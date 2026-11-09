# -*- coding: utf-8 -*-
"""DOCX fallback parser.

The normal path converts DOCX/DOC to PDF first, so page numbers match
``render``.  This parser is only used when no converter (LibreOffice /
Word COM) is available: it slices the document by top-level headings and
cannot render screenshots.
"""
import os
import re

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    _DOCX_OK = True
except Exception:
    _DOCX_OK = False

from ppt_parser import PresentationDoc, Slide, SlideBlock


def _iter_blocks(doc):
    """Yield paragraphs and tables in document order."""
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def _table_to_markdown(table) -> str:
    rows = []
    for row in table.rows:
        cells = [" ".join(c.text.split()).replace("|", "\\|") for c in row.cells]
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |",
           "| " + " | ".join(["---"] * width) + " |"]
    out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(out)


def _heading_level(para) -> int:
    try:
        name = (para.style.name or "").lower()
    except Exception:
        return 0
    m = re.match(r"heading\s*(\d)", name)
    if m:
        return int(m.group(1))
    return 1 if name in ("title", "标题") else 0


def parse_docx(path: str) -> PresentationDoc:
    """Parse .docx into one Slide per top-level heading (fallback only)."""
    if not _DOCX_OK:
        raise RuntimeError("python-docx is not installed: pip install python-docx")
    if not os.path.isfile(path):
        raise FileNotFoundError(f"file not found: {path}")
    if not path.lower().endswith(".docx"):
        raise ValueError("parse_docx only accepts .docx; convert .doc first")

    doc = Document(path)
    stem = os.path.splitext(os.path.basename(path))[0]
    sections = []
    current = {"title": stem, "blocks": []}

    for block in _iter_blocks(doc):
        if isinstance(block, Table):
            md = _table_to_markdown(block)
            if md:
                current["blocks"].append(SlideBlock(kind="table", text=md))
            continue
        text = (block.text or "").strip()
        if not text:
            continue
        level = _heading_level(block)
        if level == 1:
            if current["blocks"]:
                sections.append(current)
            current = {"title": text[:80], "blocks": []}
        elif level >= 2:
            current["blocks"].append(SlideBlock(kind="heading", text=text))
        else:
            current["blocks"].append(SlideBlock(kind="text", text=text, level=1))
    if current["blocks"] or current["title"] != stem:
        sections.append(current)
    if not sections:
        sections = [{"title": stem, "blocks": []}]

    pdoc = PresentationDoc(path=path, name=stem)
    for i, sec in enumerate(sections, start=1):
        slide = Slide(index=i, title=sec["title"])
        slide.blocks = sec["blocks"]
        pdoc.slides.append(slide)
    return pdoc
