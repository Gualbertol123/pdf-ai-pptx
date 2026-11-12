"""Block roles from position, size and repetition across pages.

Roles: ``title``, ``subtitle``, ``body``, ``footer``, ``header``, ``page_number``,
``caption``. Each decision sets ``TextBlock.role_confidence``.

* **Repetition**: a block whose text (digits masked) and position repeat on at least
  half of the pages (and on two or more) is page furniture.
* **page_number**: a short number ("7", "Page 7", "7 / 20", "- 7 -") in the top or
  bottom band; confident when its value minus the page index is constant across
  pages. When the number equals the slide number its run becomes a slide-number field.
* **footer / header**: bottom / top band, repeated or smaller than the body text.
* **title**: the largest text in the top 40% of the page, clearly larger than the
  page's body text; **subtitle**: the block right under it, smaller than the title.
* **caption**: small text right above or below a picture.
"""

from __future__ import annotations

import re
import statistics
from collections import Counter, defaultdict

from ..ir import Document, Image, Page, TextBlock

PAGE_NUM_RE = re.compile(
    r"^(?:(?:page|pag\.?|pagina|p\.|slide)\s*)?[-–]?\s*(\d{1,4})\s*[-–]?"
    r"(?:\s*(?:/|of|di|de)\s*\d{1,4})?$", re.IGNORECASE)
CAPTION_RE = re.compile(r"^(fig\.?|figure|figura|table|tabella|tab\.|source|fonte|chart|"
                        r"grafico|immagine|image)\b", re.IGNORECASE)


def block_text(tb: TextBlock) -> str:
    return "\n".join("".join(r.text for r in p.runs) for p in tb.paragraphs)


def _size(tb: TextBlock) -> float:
    """Character-weighted median run size (script runs excluded)."""
    sizes = [r.size for p in tb.paragraphs for r in p.runs if not r.baseline_shift
             for _ in range(max(len(r.text.strip()), 1))]
    return statistics.median(sizes) if sizes else 12.0


def _n_chars(tb: TextBlock) -> int:
    return sum(len(r.text.strip()) for p in tb.paragraphs for r in p.runs)


def _body_size(blocks: list[TextBlock]) -> float | None:
    sizes = [s for tb in blocks for s in [_size(tb)] * max(_n_chars(tb), 1)]
    return statistics.median(sizes) if sizes else None


def _rep_key(tb: TextBlock) -> tuple[str, int, int]:
    text = re.sub(r"\d+", "#", block_text(tb).strip().lower())
    return text, round(tb.bbox[0] / 6), round(tb.bbox[1] / 6)


def assign_roles(doc: Document, blocks: dict[int, list[TextBlock]]) -> None:
    pages = {p.index: p for p in doc.pages}
    n_pages = len([i for i, b in blocks.items() if b])
    rep = Counter(k for bl in blocks.values() for k in {_rep_key(tb) for tb in bl})
    min_rep = max(2, (n_pages + 1) // 2)
    doc_body = _body_size([tb for bl in blocks.values() for tb in bl]) or 12.0
    # page numbers: value - page index must be constant to be confident
    offsets: dict[str, int] = defaultdict(int)
    num_cands: dict[str, tuple[int, int]] = {}
    for idx, bl in blocks.items():
        page = pages[idx]
        for tb in bl:
            m = PAGE_NUM_RE.match(block_text(tb).strip())
            if m and len(tb.paragraphs) == 1 and _in_band(tb, page, 0.15, 0.10):
                num_cands[tb.id] = (idx, int(m.group(1)))
                offsets[str(int(m.group(1)) - idx)] += 1
    best_off = max(offsets.items(), key=lambda kv: (kv[1], kv[0]))[0] if offsets else None
    for idx, bl in blocks.items():
        page = pages[idx]
        body = _body_size([tb for tb in bl if _n_chars(tb) >= 20]) or doc_body
        for tb in bl:
            tb.role, tb.role_confidence = "body", 0.7
        _furniture(tb_list=bl, page=page, rep=rep, min_rep=min_rep, body=body,
                   num_cands=num_cands, best_off=best_off, offsets=offsets)
        _title(bl, page, body)
        _captions(bl, page, body)


def _in_band(tb: TextBlock, page: Page, bottom: float, top: float) -> bool:
    return tb.bbox[1] >= (1 - bottom) * page.height or tb.bbox[3] <= top * page.height


def _furniture(*, tb_list: list[TextBlock], page: Page, rep: Counter[tuple[str, int, int]],
               min_rep: int, body: float, num_cands: dict[str, tuple[int, int]],
               best_off: str | None, offsets: dict[str, int]) -> None:
    for tb in tb_list:
        repeated = rep[_rep_key(tb)] >= min_rep
        if tb.id in num_cands:
            idx, value = num_cands[tb.id]
            consistent = best_off is not None and str(value - idx) == best_off and offsets[
                best_off] >= 2
            if not consistent and value != idx + 1:
                continue  # a lone number in the band (axis tick, legend year): not furniture
            tb.role = "page_number"
            tb.role_confidence = 0.95 if consistent else 0.75
            if value == idx + 1 and tb.role_confidence >= 0.75:
                for p in tb.paragraphs:
                    for r in p.runs:
                        if r.text.strip() == str(value):
                            r.field_type = "slidenum"
            continue
        bottom = tb.bbox[1] >= 0.88 * page.height
        top = tb.bbox[3] <= 0.08 * page.height
        small = _size(tb) <= 0.85 * body
        if bottom and (repeated or small):
            tb.role, tb.role_confidence = "footer", 0.9 if repeated else 0.6
        elif top and repeated:
            tb.role, tb.role_confidence = "header", 0.85


def _title(bl: list[TextBlock], page: Page, body: float) -> None:
    cands = [tb for tb in bl if tb.role == "body" and tb.bbox[1] < 0.4 * page.height
             and len(tb.paragraphs) <= 3 and _n_chars(tb) > 0]
    if not cands:
        return
    best = max(cands, key=lambda tb: (_size(tb), -tb.bbox[1]))
    ratio = _size(best) / max(body, 1.0)
    alone = len([tb for tb in bl if tb.role == "body"]) == 1
    if ratio >= 1.2 or (alone and _size(best) >= 20):
        best.role = "title"
        best.role_confidence = round(min(0.95, 0.55 + (ratio - 1.0) * 0.6), 3) if not alone \
            else 0.6
        _subtitle(bl, best, body)


def _subtitle(bl: list[TextBlock], title: TextBlock, body: float) -> None:
    ts = _size(title)
    below = [tb for tb in bl if tb.role == "body" and tb.bbox[1] >= title.bbox[1]
             and tb.bbox[1] - title.bbox[3] <= 3.0 * ts and len(tb.paragraphs) <= 3
             and min(tb.bbox[2], title.bbox[2]) - max(tb.bbox[0], title.bbox[0]) > 0]
    if not below:
        return
    sub = min(below, key=lambda tb: tb.bbox[1])
    sw, tw = sub.bbox[2] - sub.bbox[0], title.bbox[2] - title.bbox[0]
    two_lines = (sub.bbox[3] - sub.bbox[1]) <= 2.5 * 1.2 * _size(sub)
    # larger than the body text, or the page is just a title and this block
    standout = _size(sub) >= 1.1 * body or len([tb for tb in bl if tb.role == "body"]) <= 1
    if (_size(sub) < ts and sw >= 0.25 * tw and _n_chars(sub) >= 4 and two_lines and standout
            and not any(p.bullet.kind != "none" for p in sub.paragraphs)):
        sub.role, sub.role_confidence = "subtitle", 0.6


def _captions(bl: list[TextBlock], page: Page, body: float) -> None:
    images = [e for e in page.elements if isinstance(e, Image) and not e.is_background
              and (e.bbox[2] - e.bbox[0]) * (e.bbox[3] - e.bbox[1]) > 0.01 * page.width
              * page.height]
    for tb in bl:
        if tb.role != "body" or len(tb.paragraphs) > 3:
            continue
        s = _size(tb)
        text = block_text(tb).strip()
        for img in images:
            ix0, iy0, ix1, iy1 = img.bbox
            w = tb.bbox[2] - tb.bbox[0]
            overlap = min(tb.bbox[2], ix1) - max(tb.bbox[0], ix0)
            gap = min(abs(tb.bbox[1] - iy1), abs(iy0 - tb.bbox[3]))
            if overlap >= 0.5 * w and gap <= 2.0 * s and (s < 0.95 * body
                                                          or CAPTION_RE.match(text)):
                tb.role = "caption"
                tb.role_confidence = 0.85 if CAPTION_RE.match(text) else 0.65
                break
