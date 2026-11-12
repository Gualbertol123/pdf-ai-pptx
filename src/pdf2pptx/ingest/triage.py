"""TRIAGE: classify each page as born_digital | broken_text | outlined_text | scanned.

Rules only (PLAN.md §8 rule 2), computed from the ingested IR:

* **scanned**: (almost) no visible glyphs and images cover most of the page.
  An invisible OCR text layer over the image is reported in the reason.
* **outlined_text**: many small filled, curved sub-paths sized like letters and
  more of them than real glyphs (text exported as outlines).
* **broken_text**: many glyphs are private-use, unmapped or U+FFFD, or the text
  reads badly (Latin letter-bigram plausibility score).
* **born_digital**: everything else.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass

import numpy as np

from ..ir import Document, Image, Page, Shape, Triage

# Letter bigrams of ordinary English + Italian prose. A page of real Latin-script
# text scores ~0.95+ against this set; text from a broken /ToUnicode map scores
# far lower.
_REFERENCE = """
the quick brown fox jumps over the lazy dog while the results of this quarter show
that revenue grew strongly across every region and the board approved the new plan
for next year which includes hiring more people in sales marketing and engineering
we expect customers to adopt the product quickly because it solves real problems
with privacy security performance and quality management reports tables charts
figure summary introduction conclusion method analysis growth market strategy
il presente documento descrive i risultati del trimestre e la strategia per il
prossimo anno la crescita dei ricavi è stata forte in tutte le regioni e il
consiglio ha approvato il nuovo piano che prevede assunzioni nelle vendite nel
marketing e nello sviluppo della qualità dei servizi offerti ai clienti città
perché attività università società più già però così può gestione relazione
tabella figura introduzione conclusione analisi mercato obiettivi sviluppo
which know would knowledge acknowledge equip squad quiz jazz pizza zone lazy
by my sky gym rhythm yellow yes eye boy you joy enjoy jump just major project
expect excel extra next box six mix exam taxi oxygen view vivid have give over
awkward wh wr sw tw dw kn kl ck nk gh ph sh th ch sc sk sp st sn sm sl qu
"""

LETTER_MIN, LETTER_MAX = 2.0, 60.0  # pt: size of a sub-path that may be a letter


def _fold(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text.lower())
                   if not unicodedata.combining(c))


def _bigrams(text: str) -> list[str]:
    t = _fold(text)
    return [t[i:i + 2] for i in range(len(t) - 1)
            if "a" <= t[i] <= "z" and "a" <= t[i + 1] <= "z"]


_REF_BIGRAMS = frozenset(_bigrams(_REFERENCE))


def bigram_score(text: str) -> tuple[float, int]:
    """(share of letter bigrams that occur in ordinary prose, number of bigrams)."""
    bg = _bigrams(text)
    if not bg:
        return 1.0, 0
    return sum(b in _REF_BIGRAMS for b in bg) / len(bg), len(bg)


def _bad_char(ch: str, notes: list[str]) -> bool:
    o = ord(ch)
    return (0xE000 <= o <= 0xF8FF or o == 0xFFFD or o < 0x20 or "unmapped" in notes
            or o >= 0xF0000)


def image_coverage(page: Page, grid: int = 64) -> float:
    if page.width <= 0 or page.height <= 0:
        return 0.0
    mask = np.zeros((grid, grid), dtype=bool)
    for e in page.elements:
        if not isinstance(e, Image):
            continue
        x0, y0, x1, y1 = e.bbox
        if e.rotation not in (0.0, 180.0):  # rotated frame: use its rotated extent
            cx, cy, r = (x0 + x1) / 2, (y0 + y1) / 2, max(x1 - x0, y1 - y0) / 2
            x0, y0, x1, y1 = cx - r, cy - r, cx + r, cy + r
        c0 = int(np.clip(x0 / page.width * grid, 0, grid))
        c1 = int(np.clip(np.ceil(x1 / page.width * grid), 0, grid))
        r0 = int(np.clip(y0 / page.height * grid, 0, grid))
        r1 = int(np.clip(np.ceil(y1 / page.height * grid), 0, grid))
        mask[r0:r1, c0:c1] = True
    return float(mask.mean())


def letter_like_subpaths(page: Page) -> int:
    """Filled, closed, curved sub-paths whose size is in the range of a letter."""
    n = 0
    for e in page.elements:
        if not isinstance(e, Shape) or e.path is None or e.fill.kind == "none":
            continue
        pts: list[tuple[float, float]] = []
        curved = False
        for seg in [*e.path, None]:
            if seg is None or seg.op == "M":
                if len(pts) >= 3 and curved:
                    xs = [p[0] for p in pts]
                    ys = [p[1] for p in pts]
                    w, h = max(xs) - min(xs), max(ys) - min(ys)
                    if LETTER_MIN <= max(w, h) <= LETTER_MAX:
                        n += 1
                pts, curved = [], False
            if seg is None:
                break
            pts.extend(seg.pts)
            curved = curved or seg.op == "C"
    return n


@dataclass
class PageStats:
    visible_glyphs: int
    invisible_glyphs: int
    bad_glyphs: int
    image_coverage: float
    letter_paths: int
    bigram: float
    n_bigrams: int


def page_stats(page: Page) -> PageStats:
    visible = [g for g in page.glyphs if g.render_mode not in (3, 7) and not g.text.isspace()]
    invisible = sum(1 for g in page.glyphs if g.render_mode in (3, 7))
    bad = sum(1 for g in visible if _bad_char(g.text, g.notes))
    score, nb = bigram_score("".join(g.text if not g.text.isspace() else " "
                                     for g in page.glyphs if g.render_mode not in (3, 7)))
    return PageStats(len(visible), invisible, bad, image_coverage(page),
                     letter_like_subpaths(page), score, nb)


def classify(page: Page) -> Triage:
    s = page_stats(page)
    facts = (f"glyphs={s.visible_glyphs} invisible={s.invisible_glyphs} bad={s.bad_glyphs} "
             f"image_cov={s.image_coverage:.2f} letter_paths={s.letter_paths} "
             f"bigram={s.bigram:.2f}/{s.n_bigrams}")
    if s.visible_glyphs < 10 and s.image_coverage >= 0.5:
        ocr = " with an invisible OCR text layer" if s.invisible_glyphs >= 10 else ""
        return Triage(kind="scanned", confidence=min(1.0, 0.5 + s.image_coverage / 2),
                      reason=f"no visible text, images cover {s.image_coverage:.0%}{ocr}; "
                             + facts)
    if s.letter_paths >= 20 and s.letter_paths > s.visible_glyphs:
        return Triage(kind="outlined_text", confidence=0.7,
                      reason=f"{s.letter_paths} letter-sized curved filled sub-paths vs "
                             f"{s.visible_glyphs} glyphs; " + facts)
    if s.visible_glyphs >= 20 and s.bad_glyphs / s.visible_glyphs >= 0.25:
        return Triage(kind="broken_text", confidence=0.8,
                      reason=f"{s.bad_glyphs / s.visible_glyphs:.0%} of glyphs are private-use"
                             f" / unmapped; " + facts)
    if s.n_bigrams >= 60 and s.bigram < 0.6:
        return Triage(kind="broken_text", confidence=0.6,
                      reason=f"text reads badly (bigram score {s.bigram:.2f}); " + facts)
    conf = 1.0 if s.letter_paths < 10 else 0.8
    return Triage(kind="born_digital", confidence=conf, reason=facts)


def triage_document(doc: Document) -> Document:
    """Classify every page (in place) and return the document with stage=triage."""
    for page in doc.pages:
        prev = page.triage
        t = classify(page)
        if prev.fallback_image:  # parse failed: keep the failure, add the facts
            t.fallback_image = True
            t.confidence = 0.0
            t.reason = prev.reason + "; " + t.reason
        elif prev.reason:  # e.g. "skipped: shading=2" from ingest
            t.reason = t.reason + "; " + prev.reason
        page.triage = t
    doc.stage = "triage"
    return doc
