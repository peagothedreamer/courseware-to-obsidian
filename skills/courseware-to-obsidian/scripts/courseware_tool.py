#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deterministic helpers for the courseware-to-obsidian skill.

Claude reads and writes the notes; this script only does the mechanical work:

  extract  <file> [--out PATH] [--titles] [--pages 1-9]
  render   <file> <page> <out.png>
  finalize <course_dir> [--vault-root ROOT]
  check    <course_dir>

Supported inputs:
  .pptx/.ppt  python-pptx; legacy .ppt is converted to .pptx first
  .pdf        PyMuPDF
  .docx/.doc  converted to PDF when LibreOffice/Word is available,
              otherwise .docx is sliced by headings (no screenshots)
  .md/.txt    sliced by top-level headings
  images      one page each; no OCR, read the file with the Read tool

Requires: python-pptx, pymupdf, python-docx.
"""
import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# Windows consoles default to GBK; paths/messages here are Chinese.
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import graph_config as gcfg
from ppt_parser import parse_pptx
from pdf_parser import parse_pdf
from docx_parser import parse_docx
from plain_parser import is_image, parse_image, parse_text

CACHE_NAME = "courseware-to-obsidian"
PPT_EXTS = (".pptx", ".ppt")
WORD_EXTS = (".docx", ".doc")
TEXT_EXTS = (".md", ".markdown", ".txt")
SUPPORTED = PPT_EXTS + WORD_EXTS + TEXT_EXTS + (".pdf",)


# ---------------------------------------------------------------- extraction

_DATE_NOISE_RE = re.compile(
    r"\d{4}\s*[年/\-]\s*\d{1,2}\s*[月/\-]\s*\d{1,2}\s*日?\s*(?:星期\w)?|"
    r"[-–]\s*年\s*[-–]\s*月\s*[-–]\s*日"
)


def clean_noise(txt: str) -> str:
    """Drop slide-footer date stamps, e.g. '2026 8 30 年 月 日星期日'."""
    return _DATE_NOISE_RE.sub("", txt or "").strip()


def _cache_dir(src: str) -> str:
    key = hashlib.md5(os.path.abspath(src).encode("utf-8")).hexdigest()[:10]
    d = os.path.join(tempfile.gettempdir(), CACHE_NAME, key)
    os.makedirs(d, exist_ok=True)
    return d


def _find_soffice():
    for cand in ("soffice", "libreoffice"):
        p = shutil.which(cand)
        if p:
            return p
    for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.isfile(cand):
            return cand
    return None


_PS_TO_PPTX = r"""
$ErrorActionPreference = 'Stop'
$src = $env:C2O_SRC
$out = $env:C2O_OUT
$app = New-Object -ComObject PowerPoint.Application
try {
  try { $pres = $app.Presentations.Open($src, $true, $false, $false) }
  catch { $pres = $app.Presentations.Open($src) }
  try { $pres.SaveAs($out, 24) } finally { $pres.Close() }
} finally { $app.Quit() }
"""


def _to_pptx(src: str) -> str:
    """Convert legacy .ppt to .pptx (cached); python-pptx cannot read .ppt."""
    cache_dir = _cache_dir(src)
    stem = os.path.splitext(os.path.basename(src))[0]
    out = os.path.join(cache_dir, stem + ".pptx")
    if os.path.isfile(out) and os.path.getmtime(out) >= os.path.getmtime(src):
        return out

    soffice = _find_soffice()
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pptx",
                        "--outdir", cache_dir, src],
                       timeout=180, check=False, capture_output=True)
        if os.path.isfile(out):
            return out

    if os.name == "nt":
        env = os.environ.copy()
        env["C2O_SRC"] = os.path.abspath(src)
        env["C2O_OUT"] = os.path.abspath(out)
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-Command", _PS_TO_PPTX],
                       env=env, timeout=240, check=False, capture_output=True)
        if os.path.isfile(out):
            return out

    raise RuntimeError("cannot open legacy .ppt; save it as .pptx first "
                       "(install LibreOffice or PowerPoint)")


_PS_PPT_TO_PDF = r"""
$ErrorActionPreference = 'Stop'
$src = $env:C2O_SRC
$pdf = $env:C2O_PDF
$app = New-Object -ComObject PowerPoint.Application
try {
  try { $pres = $app.Presentations.Open($src, $true, $false, $false) }
  catch { $pres = $app.Presentations.Open($src) }
  try { $pres.SaveAs($pdf, 32) } finally { $pres.Close() }
} finally { $app.Quit() }
"""

_PS_WORD_TO_PDF = r"""
$ErrorActionPreference = 'Stop'
$src = $env:C2O_SRC
$pdf = $env:C2O_PDF
$app = New-Object -ComObject Word.Application
$app.Visible = $false
try {
  $doc = $app.Documents.Open($src, $false, $true)
  try { $doc.ExportAsFixedFormat($pdf, 17) } finally { $doc.Close($false) }
} finally { $app.Quit() }
"""


def _to_pdf(src: str) -> str:
    """Return a PDF for src, converting Office files once and caching in TEMP."""
    if src.lower().endswith(".pdf"):
        return src

    cache_dir = _cache_dir(src)
    stem = os.path.splitext(os.path.basename(src))[0]
    pdf = os.path.join(cache_dir, stem + ".pdf")
    if os.path.isfile(pdf) and os.path.getsize(pdf) > 1000 and \
            os.path.getmtime(pdf) >= os.path.getmtime(src):
        return pdf

    soffice = _find_soffice()
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", cache_dir, src],
                       timeout=180, check=False, capture_output=True)
        if os.path.isfile(pdf) and os.path.getsize(pdf) > 1000:
            return pdf

    if os.name == "nt":
        low = src.lower()
        script = _PS_WORD_TO_PDF if low.endswith(WORD_EXTS) else _PS_PPT_TO_PDF
        env = os.environ.copy()
        env["C2O_SRC"] = os.path.abspath(src)
        env["C2O_PDF"] = os.path.abspath(pdf)
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-Command", script],
                       env=env, timeout=240, check=False, capture_output=True)
        if os.path.isfile(pdf) and os.path.getsize(pdf) > 1000:
            return pdf

    raise RuntimeError("cannot convert to PDF (install LibreOffice or Office)")


def load_doc(path: str):
    low = path.lower()
    if low.endswith(".pdf"):
        return parse_pdf(path)
    if low.endswith(".ppt"):
        original = (path, os.path.splitext(os.path.basename(path))[0])
        doc = parse_pptx(_to_pptx(path))
        doc.path, doc.name = original
        return doc
    if low.endswith(".pptx"):
        return parse_pptx(path)
    if low.endswith(WORD_EXTS):
        try:
            doc = parse_pdf(_to_pdf(path))
            doc.path = path
            doc.name = os.path.splitext(os.path.basename(path))[0]
            return doc
        except Exception:
            if low.endswith(".docx"):
                return parse_docx(path)
            raise
    if low.endswith(TEXT_EXTS):
        return parse_text(path)
    if is_image(path):
        return parse_image(path)
    raise SystemExit("unsupported file (need PPT/PDF/Word/Markdown/text/image): " + path)


_RANGE_RE = re.compile(r"^\s*(\d+)\s*[-~至到]\s*(\d+)\s*$")


def parse_pages(spec: str, total: int):
    out = set()
    for part in re.split(r"[,\s;，；、]+", spec or ""):
        if not part:
            continue
        m = _RANGE_RE.match(part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.update(range(min(a, b), max(a, b) + 1))
        elif part.isdigit():
            out.add(int(part))
    return {p for p in out if 1 <= p <= total}


def cmd_extract(args):
    doc = load_doc(args.file)
    wanted = parse_pages(args.pages, doc.total_pages) if args.pages else None
    lines = [f"SOURCE: {doc.path}", f"NAME: {doc.name}", f"PAGES: {doc.total_pages}"]
    for s in doc.slides:
        if wanted is not None and s.index not in wanted:
            continue
        text = clean_noise(s.to_text())
        n = len(text)
        title = clean_noise(s.title or "")
        if args.titles:
            body = text.split("\n", 1)[1] if text.startswith("# ") else text
            body = re.sub(r"\s+", " ", body)
            short = " [少字·可能图/公式]" if n < 18 else ""
            lines.append(f"p{s.index} | 图{s.image_count} | {n}字{short} | "
                         f"{title or body[:40] or '(空白页)'}")
        else:
            lines.append(f"\n===== p{s.index} | 标题：{title or '(无)'} | "
                         f"图{s.image_count} | {n}字 =====")
            lines.append(text or "(空白页)")
    out = "\n".join(lines) + "\n"
    if args.titles and not args.out:
        print(out, end="")
        return
    target = args.out
    if not target:
        cache = os.path.join(tempfile.gettempdir(), CACHE_NAME)
        os.makedirs(cache, exist_ok=True)
        key = hashlib.md5(os.path.abspath(doc.path).encode("utf-8")).hexdigest()[:10]
        tag = re.sub(r"[^0-9,\-]", "", args.pages or "")
        target = os.path.join(cache, f"{key}{'-p' + tag if tag else ''}-extract.txt")
    with open(target, "w", encoding="utf-8") as f:
        f.write(out)
    print(target)


# ---------------------------------------------------------------- rendering

def _pdf_page_to_png(src_pdf: str, page_no: int, out_png: str, dpi: int = 110) -> bool:
    try:
        import pymupdf
    except Exception:
        import fitz as pymupdf
    doc = pymupdf.open(src_pdf)
    try:
        if not (1 <= page_no <= doc.page_count):
            return False
        doc[page_no - 1].get_pixmap(dpi=dpi).save(out_png)
    finally:
        doc.close()
    return os.path.isfile(out_png) and os.path.getsize(out_png) > 500


def cmd_render(args):
    try:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        if is_image(args.file):
            if args.page != 1:
                raise RuntimeError("image sources have exactly one page")
            shutil.copyfile(args.file, args.out)
            ok = True
        elif args.file.lower().endswith(TEXT_EXTS):
            raise RuntimeError("Markdown/text sources have no pages to render")
        else:
            pdf = _to_pdf(args.file)
            ok = _pdf_page_to_png(pdf, args.page, args.out)
    except Exception as e:
        print(f"render failed: {e}", file=sys.stderr)
        sys.exit(1)
    if not ok:
        print(f"render failed: page {args.page} out of range", file=sys.stderr)
        sys.exit(1)
    print(args.out)


# ---------------------------------------------------------------- contact sheet

_CONTACT_MAX_PAGES = 40


def _page_count(src: str) -> int:
    if is_image(src):
        return 1
    if src.lower().endswith(TEXT_EXTS):
        return 0
    try:
        import pymupdf
    except Exception:
        import fitz as pymupdf
    pdf = _to_pdf(src)
    doc = pymupdf.open(pdf)
    try:
        return doc.page_count
    finally:
        doc.close()


def _contact_sheet(src_pdf: str, pages, out_png: str, cols: int, width: int) -> bool:
    try:
        import pymupdf
    except Exception:
        import fitz as pymupdf
    src = pymupdf.open(src_pdf)
    try:
        thumbs = [(n, src[n - 1].get_pixmap(dpi=60))
                  for n in pages if 1 <= n <= src.page_count]
    finally:
        src.close()
    if not thumbs:
        return False
    # ponytail: one cell ratio for the whole sheet; decks are uniform, revisit if mixed sizes distort
    ratio = max(p.height / p.width for _, p in thumbs)
    cell_h = int(width * ratio)
    label_h, pad = 20, 10
    rows = (len(thumbs) + cols - 1) // cols
    sheet = pymupdf.open()
    page = sheet.new_page(width=cols * width + (cols + 1) * pad,
                          height=rows * (cell_h + label_h) + (rows + 1) * pad)
    for i, (n, pix) in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (width + pad)
        y = pad + r * (cell_h + label_h + pad)
        rect = pymupdf.Rect(x, y, x + width, y + cell_h)
        page.insert_image(rect, pixmap=pix)
        page.draw_rect(rect, color=(0.7, 0.7, 0.7), width=0.5)
        page.insert_text((x, y + cell_h + 14), f"p{n}", fontsize=11)
    page.get_pixmap(dpi=80).save(out_png)
    sheet.close()
    return os.path.isfile(out_png) and os.path.getsize(out_png) > 500


def cmd_contact(args):
    total = _page_count(args.file)
    if total <= 0:
        print("contact needs a paged source (PPT/PDF/Word/image)", file=sys.stderr)
        sys.exit(1)
    pages = sorted(parse_pages(args.pages, total))
    if not pages:
        print(f"no valid pages in range for a {total}-page source", file=sys.stderr)
        sys.exit(1)
    if len(pages) > _CONTACT_MAX_PAGES:
        print(f"contact sheet supports <= {_CONTACT_MAX_PAGES} pages; split the range",
              file=sys.stderr)
        sys.exit(1)
    cols = max(1, args.cols)
    width = max(80, args.width)
    out = args.out
    if not out:
        cache = os.path.join(tempfile.gettempdir(), CACHE_NAME)
        os.makedirs(cache, exist_ok=True)
        key = hashlib.md5(os.path.abspath(args.file).encode("utf-8")).hexdigest()[:10]
        tag = re.sub(r"[^0-9,\-]", "", args.pages)
        out = os.path.join(cache, f"{key}-contact-{tag}.png")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    if is_image(args.file):
        shutil.copyfile(args.file, out)
        print(out)
        return
    if not _contact_sheet(_to_pdf(args.file), pages, out, cols, width):
        print("contact failed", file=sys.stderr)
        sys.exit(1)
    print(out)


# ---------------------------------------------------------------- vault hub

COURSE_HUB_SUFFIX = " 课程总目录"
HUB_TAGS = ["课程总目录", "层级/课程"]
INDEX_FALLBACK_SUFFIX = "-目录"
GRAPH_HINT = ("> [!tip]- 🕸️ 怎么看关系图\n"
              "> 打开 Obsidian 左侧「关系图」：同一门课的笔记会聚成一团。"
              "默认每门课一个专属色；层级标签（紫＝课程枢纽、金橙＝课件目录、蓝＝正文、绿＝概念、灰蓝＝附录）"
              "在「按层级」配色下生效。")


def course_hub_name(course_dir: str) -> str:
    name = os.path.basename(os.path.normpath(os.path.abspath(course_dir))) or "课程"
    return f"{name}{COURSE_HUB_SUFFIX}"


def _read_fm_field(md_path: str, key: str = "title") -> str:
    try:
        with open(md_path, "r", encoding="utf-8") as f:
            head = f.read(2000)
    except Exception:
        return ""
    m = re.search(rf"^{re.escape(key)}:\s*(.+?)\s*$", head, re.M)
    return m.group(1).strip().strip('"').strip("'") if m else ""


def find_deck_index(deck_dir: str):
    """Folder note: <folder>.md -> <folder>-目录.md -> 00_MOC.md."""
    name = os.path.basename(os.path.normpath(deck_dir))
    for cand in (f"{name}.md", f"{name}{INDEX_FALLBACK_SUFFIX}.md", "00_MOC.md"):
        p = os.path.join(deck_dir, cand)
        if os.path.isfile(p):
            return p
    return None


def build_course_index(course_dir: str, vault_root: str):
    hub_path = os.path.join(course_dir, course_hub_name(course_dir) + ".md")
    prefix = gcfg.rel_prefix(course_dir, vault_root)
    course_name = os.path.basename(os.path.normpath(os.path.abspath(course_dir))) or "课程"

    entries = []
    for sub in sorted(os.listdir(course_dir)):
        if sub.startswith(".") or sub.lower() == "attachments":
            continue
        pdir = os.path.join(course_dir, sub)
        if not os.path.isdir(pdir):
            continue
        idx = find_deck_index(pdir)
        if not idx:
            continue
        stem = os.path.splitext(os.path.basename(idx))[0]
        n_sec = sum(1 for f in os.listdir(pdir)
                    if f.lower().endswith(".md") and f != os.path.basename(idx))
        entries.append({
            "title": _read_fm_field(idx) or sub,
            "link": "/".join(x for x in (prefix, sub, stem) if x),
            "sec": n_sec,
            "src": _read_fm_field(idx, "source"),
        })

    if not entries:
        if os.path.isfile(hub_path):
            os.remove(hub_path)
        return None

    def esc(s):
        return str(s or "").replace("|", "\\|")

    lines = ["---", f'title: "{course_name}{COURSE_HUB_SUFFIX}"',
             "type: course-index", "tags:"]
    ctag = gcfg.course_tag(course_name)
    lines += [f"  - {t}" for t in (HUB_TAGS + [ctag])]
    lines += ["---", "", f"# 📚 {course_name}{COURSE_HUB_SUFFIX}", "",
              f"> 本页串联《{course_name}》下全部课件笔记; 有了它, 关系图里这门课会自然聚成一团。",
              "", GRAPH_HINT, "", f"## 课件（{len(entries)}）", "",
              "| 课件 | 小节 | 来源课件 |", "|------|------|----------|"]
    for e in entries:
        lines.append(f"| [[{e['link']}\\|{esc(e['title'])}]] | {e['sec']} | {esc(e['src'])} |")
    lines += ["", "---", "*由 courseware-to-obsidian skill 生成; 每转换一份课件后自动刷新。*"]
    with open(hub_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return hub_path


def cmd_finalize(args):
    course_dir = os.path.abspath(args.course_dir)
    if not os.path.isdir(course_dir):
        print(f"course directory not found: {course_dir}", file=sys.stderr)
        sys.exit(1)
    root = args.vault_root or gcfg.find_vault_root(course_dir)
    if not os.path.isdir(os.path.join(root, ".obsidian")):
        print(f"not an Obsidian vault (no .obsidian): {root}", file=sys.stderr)
        sys.exit(1)
    hub = build_course_index(course_dir, root)
    ok = gcfg.apply_graph_colors(course_dir, vault_root=root,
                                 progress=lambda m: print(" ·", m))
    print(f"course_dir: {course_dir}")
    print(f"vault_root: {root}")
    print(f"hub: {hub or '(no decks found)'}")
    print(f"graph: {'ok' if ok else 'skipped'}")
    if not hub:
        sys.exit(1)


# ---------------------------------------------------------------- checking

_WIKILINK_RE = re.compile(r"\[\[([^\]\|#]*)(?:#([^\]\|]*))?(?:\|[^\]]*)?\]\]")
_HEADING_RE = re.compile(r"^#{1,6}\s+(.*?)\s*$", re.M)
_IMG_RE = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
_PLACEHOLDER_RE = re.compile(r"【\s*(?:课件)?截图\s*[:：]?\s*(?:第\s*|p\s*)?\d{1,4}\s*页?\s*】|"
                             r"<!--\s*截图\s*:")


def cmd_check(args):
    course_dir = os.path.abspath(args.course_dir)
    if not os.path.isdir(course_dir):
        print(f"course directory not found: {course_dir}", file=sys.stderr)
        sys.exit(1)
    root = args.vault_root or gcfg.find_vault_root(course_dir)
    md_files = []
    for dirpath, dirnames, filenames in os.walk(course_dir):
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d.lower() != "attachments"]
        md_files += [os.path.join(dirpath, f) for f in filenames if f.lower().endswith(".md")]

    stems = set()
    headings = {}
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for f in filenames:
            if not f.lower().endswith(".md"):
                continue
            stem = os.path.splitext(f)[0]
            stems.add(stem)
            try:
                with open(os.path.join(dirpath, f), "r", encoding="utf-8") as fh:
                    found = _HEADING_RE.findall(fh.read())
            except OSError:
                found = []
            headings.setdefault(stem, set()).update(h.strip() for h in found)

    problems = []
    missing_index = []
    for sub in sorted(os.listdir(course_dir)):
        p = os.path.join(course_dir, sub)
        if sub.startswith(".") or sub.lower() == "attachments" or not os.path.isdir(p):
            continue
        # Only NN- folders are decks; manual folders and source-PDF folders
        # are not expected to have a folder note.
        if not re.match(r"^\d{2}-", sub):
            continue
        if not find_deck_index(p):
            missing_index.append(sub)
    for name in missing_index:
        problems.append(f"deck folder without an index page: {name}/")

    placeholders = 0
    for fp in md_files:
        with open(fp, "r", encoding="utf-8") as f:
            text = f.read()
        rel = os.path.relpath(fp, root).replace(os.sep, "/")
        if _PLACEHOLDER_RE.search(text):
            placeholders += 1
            problems.append(f"unresolved screenshot placeholder: {rel}")
        for m in _WIKILINK_RE.finditer(text):
            # table cells escape the alias separator as "\|"; drop that backslash
            target = m.group(1).strip().rstrip("\\").lstrip("/")
            anchor = (m.group(2) or "").strip().rstrip("\\")
            if target:
                cand = os.path.join(root, target + ("" if target.lower().endswith(".md") else ".md"))
                if not (os.path.isfile(cand) or ("/" not in target and target in stems)):
                    problems.append(f"broken wikilink: {rel} -> [[{m.group(1)}]]")
                    continue
                base = os.path.basename(target)
                stem = base[:-3] if base.lower().endswith(".md") else base
            else:
                # [[#标题]] points at the current note
                stem = os.path.splitext(os.path.basename(fp))[0]
            if anchor and not anchor.startswith("^") and anchor not in headings.get(stem, set()):
                problems.append(f"broken anchor: {rel} -> [[{m.group(1)}#{anchor}]]")
        for m in _IMG_RE.finditer(text):
            src = m.group(1).strip()
            if "://" in src or src.startswith("#"):
                continue
            src = src.split("|")[0]
            if not os.path.isfile(os.path.join(os.path.dirname(fp), src)):
                problems.append(f"missing image: {rel} -> {src}")

    print(f"notes: {len(md_files)} | placeholders: {placeholders} | problems: {len(problems)}")
    for p in problems:
        print(" -", p)
    if problems:
        sys.exit(1)


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(prog="courseware_tool")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("extract")
    p.add_argument("file")
    p.add_argument("--out")
    p.add_argument("--titles", action="store_true", help="page titles only (cheap outline pass)")
    p.add_argument("--pages", help="e.g. 4-9 or 3,7-9")
    p.set_defaults(func=cmd_extract)

    p = sub.add_parser("render")
    p.add_argument("file")
    p.add_argument("page", type=int)
    p.add_argument("out")
    p.set_defaults(func=cmd_render)

    p = sub.add_parser("contact")
    p.add_argument("file")
    p.add_argument("pages", help="e.g. 3,7-12 (at most 40 pages)")
    p.add_argument("--out")
    p.add_argument("--cols", type=int, default=4)
    p.add_argument("--width", type=int, default=320)
    p.set_defaults(func=cmd_contact)

    p = sub.add_parser("finalize")
    p.add_argument("course_dir")
    p.add_argument("--vault-root")
    p.set_defaults(func=cmd_finalize)

    p = sub.add_parser("check")
    p.add_argument("course_dir")
    p.add_argument("--vault-root")
    p.set_defaults(func=cmd_check)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
