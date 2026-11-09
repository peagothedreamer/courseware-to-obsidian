# -*- coding: utf-8 -*-
"""Parsers for Markdown / plain text / image courseware.

Markdown and text are sliced by top-level headings (``#`` or ``##``).
An image file is a single page; the caller (Claude) must view it with the
Read tool, because this parser does no OCR.
"""
import os
import re

from ppt_parser import PresentationDoc, Slide, SlideBlock

_HEADING_RE = re.compile(r"^#\s+(.+?)\s*$")
_SUBHEAD_RE = re.compile(r"^(#{2,6})\s+(.+?)\s*$")
_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff"}


def is_image(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in _IMAGE_EXTS


def parse_text(path: str) -> PresentationDoc:
    """Slice a Markdown/plain-text file into one Slide per top-level heading."""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"file not found: {path}")
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.read().splitlines()

    stem = os.path.splitext(os.path.basename(path))[0]
    sections = []
    current = {"title": stem, "blocks": []}

    for line in lines:
        m = _HEADING_RE.match(line)
        sub = _SUBHEAD_RE.match(line)
        if m:
            if current["blocks"]:
                sections.append(current)
            current = {"title": m.group(1)[:80], "blocks": []}
        elif sub:
            current["blocks"].append(SlideBlock(kind="heading", text=sub.group(2)))
        elif line.strip():
            current["blocks"].append(SlideBlock(kind="text", text=line.rstrip()))
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


def parse_image(path: str) -> PresentationDoc:
    """Treat one image file as one page. No OCR; Claude reads it visually."""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"file not found: {path}")
    stem = os.path.splitext(os.path.basename(path))[0]
    slide = Slide(index=1, title=stem, has_images=True, image_count=1)
    slide.blocks.append(SlideBlock(kind="image"))
    pdoc = PresentationDoc(path=path, name=stem)
    pdoc.slides.append(slide)
    return pdoc
