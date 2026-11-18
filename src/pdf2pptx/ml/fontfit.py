"""Font geometry for OCR text: which installed family and size explain the word widths.

For a word, the ink width in the image is ``size × (Σ advance − lsb(first) − rsb(last))``
in em of the right font. Fitting one size per line over all its words, the family
with the smallest relative residual over the page wins (Arial's wide ``a`` and
narrow ``t`` give a different width profile from Calibri's, even though their
x-height / width ratios are nearly equal).
"""

from __future__ import annotations

import functools
import math
from dataclasses import dataclass
from typing import Any

from .. import log
from ..fonts import index as fidx

LOG = log.get(__name__)

# the families office documents use most; each face costs ~50 ms to load
CANDIDATES = ("Calibri", "Arial", "Times New Roman", "Cambria", "Verdana")
DEFAULT_FAMILY = "Calibri"
FALLBACK_EM = 0.5  # advance of an unknown character


@dataclass(frozen=True)
class CharGeo:
    adv: float  # em
    x0: float  # ink bounds, em (x from the origin, y up from the baseline)
    x1: float
    y0: float
    y1: float


class Face:
    """Advance and ink bounds of an installed face, per character (lazy)."""

    def __init__(self, family: str, bold: bool, font: Any) -> None:
        self.family, self.bold, self.font = family, bold, font
        self.upm = float(font["head"].unitsPerEm)
        self.cmap = font.getBestCmap() or {}
        self.gs = font.getGlyphSet()
        self.hmtx = font["hmtx"].metrics
        os2 = font["OS/2"]
        self.ascent = float(os2.usWinAscent) / self.upm
        self.descent = float(os2.usWinDescent) / self.upm
        xh = float(getattr(os2, "sxHeight", 0) or 0) / self.upm
        self.x_height = xh if xh > 0 else (self.char("x").y1 if self.char("x") else 0.5)
        self._cache: dict[str, CharGeo | None] = {}

    def char(self, ch: str) -> CharGeo | None:
        if ch in self._cache:
            return self._cache[ch]
        g = self.cmap.get(ord(ch)) if len(ch) == 1 else None
        geo = None
        if g is not None and g in self.hmtx:
            from fontTools.pens.boundsPen import BoundsPen

            adv = self.hmtx[g][0] / self.upm
            try:
                bp = BoundsPen(self.gs)
                self.gs[g].draw(bp)
                b = bp.bounds
            except Exception:
                b = None
            if b is None:  # blank glyph (space)
                geo = CharGeo(adv, 0.0, adv, 0.0, 0.0)
            else:
                geo = CharGeo(adv, b[0] / self.upm, b[2] / self.upm, b[1] / self.upm,
                              b[3] / self.upm)
        self._cache[ch] = geo
        return geo

    def adv(self, ch: str) -> float:
        g = self.char(ch)
        return g.adv if g is not None else FALLBACK_EM

    def ink_em(self, word: str) -> float:
        """Ink width of ``word`` in em: advances minus the outer side bearings."""
        if not word:
            return 0.0
        total = sum(self.adv(c) for c in word)
        first, last = self.char(word[0]), self.char(word[-1])
        lsb = first.x0 if first is not None else 0.05
        rsb = (last.adv - last.x1) if last is not None else 0.05
        return max(total - lsb - rsb, 0.1)


@functools.cache
def face(family: str, bold: bool = False) -> Face | None:
    fonts = fidx.installed_fonts()
    names = ([f"{family} bold".lower()] if bold else []) + [family.lower()]
    for name in names:
        p = fonts.get(name)
        if p is None or not p.is_file():
            continue
        try:
            from fontTools.ttLib import TTFont

            kw = {"fontNumber": 0} if p.suffix.lower() == ".ttc" else {}
            f = TTFont(str(p), lazy=True, **kw)  # type: ignore[arg-type]
            return Face(family, bold and name.endswith(" bold"), f)
        except Exception as exc:
            LOG.debug("font %s unreadable: %s", p, exc)
    return None


def candidates() -> list[str]:
    return [f for f in CANDIDATES if face(f) is not None] or [DEFAULT_FAMILY]


@dataclass
class WordObs:
    text: str
    width: float  # ink width (any unit; sizes come out in the same unit)


def fit_size(f: Face, words: list[WordObs]) -> tuple[float, float]:
    """(size, relative residual) of the least-squares size through the origin."""
    e = [f.ink_em(w.text) for w in words]
    w = [o.width for o in words]
    se2 = sum(x * x for x in e)
    if se2 <= 0:
        return 0.0, math.inf
    s = sum(a * b for a, b in zip(w, e, strict=True)) / se2
    res = sum((a - s * b) ** 2 for a, b in zip(w, e, strict=True))
    norm = sum(a * a for a in w) or 1.0
    return s, math.sqrt(res / norm)


def choose_family(lines: list[tuple[list[WordObs], bool]]) -> tuple[str, dict[str, float]]:
    """Family whose width profile fits the page's lines best (character weighted).
    ``lines``: (words, bold) per line; lines with fewer than 2 words or 6 letters
    say little and are skipped."""
    scores: dict[str, float] = {}
    for fam in candidates():
        tot, wsum = 0.0, 0.0
        for words, bold in lines:
            n = sum(len(w.text) for w in words)
            if len(words) < 2 or n < 6:
                continue
            f = (face(fam, True) if bold else None) or face(fam)
            if f is None:
                break
            _, r = fit_size(f, words)
            if math.isfinite(r):
                tot += r * n
                wsum += n
        else:
            if wsum > 0:
                scores[fam] = round(tot / wsum, 5)
    if not scores:
        return DEFAULT_FAMILY, {}
    best = min(sorted(scores), key=lambda k: scores[k])
    return best, dict(sorted(scores.items()))
