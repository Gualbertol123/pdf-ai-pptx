"""PDF side of scoring: page renders, page text and word lines.

Scoring-only helpers. The converter's own ingest (M2, ``ingest/parse.py``) is
separate; keeping the reference side independent avoids grading our parser
with itself.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pdfplumber
import pypdfium2 as pdfium

from .metrics import Word


@dataclass
class PdfPageRef:
    index: int
    width: float
    height: float
    text: str
    lines: list[list[Word]]


def render_pages(pdf: Path, dpi: int) -> list[np.ndarray]:
    """RGB renders at exactly round(size_pt * dpi / 72) pixels."""
    doc = pdfium.PdfDocument(str(pdf))
    out = []
    try:
        for i in range(len(doc)):
            page = doc[i]
            w, h = page.get_size()
            tw, th = round(w * dpi / 72), round(h * dpi / 72)
            img = page.render(scale=dpi / 72).to_pil().convert("RGB")
            if img.size != (tw, th):
                img = img.resize((tw, th))
            out.append(np.asarray(img))
    finally:
        doc.close()
    return out


def page_texts(pdf: Path) -> list[str]:
    doc = pdfium.PdfDocument(str(pdf))
    try:
        texts = []
        for i in range(len(doc)):
            tp = doc[i].get_textpage()
            texts.append(tp.get_text_range())
            tp.close()
        return texts
    finally:
        doc.close()


def group_lines(words: list[Word]) -> list[list[Word]]:
    """Words -> visual lines: vertical bands, then split at large horizontal gaps."""
    lines: list[list[Word]] = []
    bands: list[list[Word]] = []
    for w in sorted(words, key=lambda w: ((w.y0 + w.y1) / 2, w.x0)):
        cy, hh = (w.y0 + w.y1) / 2, (w.y1 - w.y0)
        if bands:
            band = bands[-1]
            bcy = sum((b.y0 + b.y1) / 2 for b in band) / len(band)
            bh = min(b.y1 - b.y0 for b in band)
            if abs(cy - bcy) <= 0.45 * min(hh, bh):
                band.append(w)
                continue
        bands.append([w])
    for band in bands:
        band.sort(key=lambda w: w.x0)
        h = float(np.median([w.y1 - w.y0 for w in band]))
        cur = [band[0]]
        for w in band[1:]:
            if w.x0 - cur[-1].x1 > 1.0 * h:
                lines.append(cur)
                cur = [w]
            else:
                cur.append(w)
        lines.append(cur)
    return lines


def page_refs(pdf: Path) -> list[PdfPageRef]:
    texts = page_texts(pdf)
    refs = []
    with pdfplumber.open(str(pdf)) as doc:
        for i, page in enumerate(doc.pages):
            words = [
                Word(w["text"], float(w["x0"]), float(w["top"]), float(w["x1"]),
                     float(w["bottom"]))
                for w in page.extract_words(keep_blank_chars=False, use_text_flow=False)
            ]
            refs.append(PdfPageRef(i, float(page.width), float(page.height), texts[i],
                                   group_lines(words)))
    return refs
