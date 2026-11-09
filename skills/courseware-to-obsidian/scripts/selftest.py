#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Minimal self-check: python selftest.py"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from ppt_parser import parse_pptx
from pdf_parser import parse_pdf
from docx_parser import parse_docx
from plain_parser import parse_text, parse_image


def test_pptx():
    from pptx import Presentation
    fd, path = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    try:
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "1.1 热力学概论"
        frame = slide.placeholders[1].text_frame
        frame.text = "内能"
        frame.add_paragraph().text = "焓"
        prs.save(path)

        doc = parse_pptx(path)
        assert doc.total_pages == 1, doc.total_pages
        text = doc.slides[0].to_text()
        assert "1.1 热力学概论" in doc.slides[0].title, doc.slides[0].title
        assert "内能" in text and "焓" in text, text
    finally:
        os.remove(path)


def test_pdf():
    import pymupdf
    fd, path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    try:
        pdf = pymupdf.open()
        page = pdf.new_page()
        page.insert_text((72, 72), "1.2 焓", fontsize=20)
        page.insert_text((72, 120), "H = U + pV", fontsize=12)
        pdf.save(path)
        pdf.close()

        doc = parse_pdf(path)
        assert doc.total_pages == 1, doc.total_pages
        assert "H = U + pV" in doc.slides[0].to_text(), doc.slides[0].to_text()
    finally:
        os.remove(path)


def test_docx():
    from docx import Document
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        d = Document()
        d.add_heading("1.1 随机事件", level=1)
        d.add_paragraph("样本空间")
        d.add_heading("性质", level=2)
        d.add_paragraph("P(A) ≥ 0")
        d.save(path)

        doc = parse_docx(path)
        assert doc.total_pages >= 1, doc.total_pages
        joined = "\n".join(s.to_text() for s in doc.slides)
        assert "1.1 随机事件" in joined and "P(A) ≥ 0" in joined, joined
    finally:
        os.remove(path)


def test_text_and_image():
    fd, path = tempfile.mkstemp(suffix=".md")
    os.close(fd)
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("# 第一节\n正文\n## 小节\n要点\n# 第二节\n内容\n")
        doc = parse_text(path)
        assert doc.total_pages == 2, doc.total_pages
        assert doc.slides[0].title == "第一节", doc.slides[0].title
        assert "要点" in doc.slides[0].to_text(), doc.slides[0].to_text()
    finally:
        os.remove(path)

    fd, img = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        doc = parse_image(img)
        assert doc.total_pages == 1 and doc.slides[0].image_count == 1
    finally:
        os.remove(img)


def _run(*args):
    return subprocess.run([sys.executable, os.path.join(HERE, "courseware_tool.py"), *args],
                          capture_output=True, text=True, encoding="utf-8")


def test_vault_check():
    root = tempfile.mkdtemp()
    try:
        vault = os.path.join(root, "vault")
        course = os.path.join(vault, "测试课")
        deck = os.path.join(course, "01-第一讲")
        att = os.path.join(deck, "attachments")
        os.makedirs(os.path.join(vault, ".obsidian"))
        os.makedirs(att)
        with open(os.path.join(vault, ".obsidian", "app.json"), "w", encoding="utf-8") as f:
            f.write("{}")
        with open(os.path.join(deck, "01-第一讲.md"), "w", encoding="utf-8") as f:
            f.write('---\ntitle: "第一讲"\nsource: "x.pdf"\n---\n\n'
                    '# 📗 第一讲\n\n- [[测试课/01-第一讲/1.1-测试|1.1 测试]]\n')
        with open(os.path.join(deck, "1.1-测试.md"), "w", encoding="utf-8") as f:
            f.write("# 1.1 测试\n\n![图](attachments/x.png)\n")
        open(os.path.join(att, "x.png"), "wb").close()

        ok = _run("check", course, "--vault-root", vault)
        assert ok.returncode == 0, ok.stdout + ok.stderr
        assert "problems: 0" in ok.stdout, ok.stdout

        fin = _run("finalize", course, "--vault-root", vault)
        assert fin.returncode == 0, fin.stdout + fin.stderr
        assert os.path.isfile(os.path.join(course, "测试课 课程总目录.md"))

        with open(os.path.join(deck, "1.1-测试.md"), "a", encoding="utf-8") as f:
            f.write("\n[[测试课/01-第一讲/1.1-测试#1.1 测试|有效锚点]]\n")
        good_anchor = _run("check", course, "--vault-root", vault)
        assert good_anchor.returncode == 0, good_anchor.stdout + good_anchor.stderr

        with open(os.path.join(deck, "1.1-测试.md"), "a", encoding="utf-8") as f:
            f.write("\n[[测试课/01-第一讲/1.1-测试#不存在的标题|坏锚点]]\n")
        bad_anchor = _run("check", course, "--vault-root", vault)
        assert bad_anchor.returncode != 0 and "broken anchor" in bad_anchor.stdout, bad_anchor.stdout

        with open(os.path.join(deck, "1.1-测试.md"), "a", encoding="utf-8") as f:
            f.write("\n[[测试课/不存在|断链]]\n")
        bad = _run("check", course, "--vault-root", vault)
        assert bad.returncode != 0 and "broken wikilink" in bad.stdout, bad.stdout
    finally:
        shutil.rmtree(root, ignore_errors=True)


def test_contact():
    import pymupdf
    root = tempfile.mkdtemp()
    try:
        pdf_path = os.path.join(root, "deck.pdf")
        pdf = pymupdf.open()
        for i in range(3):
            page = pdf.new_page()
            page.insert_text((72, 72), f"page {i + 1}", fontsize=20)
        pdf.save(pdf_path)
        pdf.close()

        out = os.path.join(root, "sheet.png")
        run = _run("contact", pdf_path, "1-3", "--out", out)
        assert run.returncode == 0, run.stdout + run.stderr
        assert os.path.isfile(out) and os.path.getsize(out) > 1000, run.stdout
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    test_pptx()
    test_pdf()
    test_docx()
    test_text_and_image()
    test_vault_check()
    test_contact()
    print("selftest ok")
