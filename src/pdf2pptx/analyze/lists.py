"""List detection: bullets, numbering, picture bullets and indent levels.

A *label* is the leading item of a text line that PowerPoint would draw as a bullet:

* a bullet glyph (``•``, ``–``, ``▪`` ...) that is its own word, or a strong bullet
  glyph glued to the text;
* a number label: ``1.`` ``2)`` ``(3)`` ``a.`` ``b)`` ``iv.`` ``B.`` ...;
* a small picture just left of the line, vertically centred on the text
  (detected in ``text_group``, which knows the page images).

``resolve_numbering`` fixes the scheme of ambiguous labels (``i.`` is roman or alpha)
from the sequence, and ``assign_levels`` maps label x positions inside a block to
``Paragraph.level`` (one level per distinct indent, left to right).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

STRONG_BULLETS = "•◦▪▫■□●○◆◇►▸▹▶➢➤➣✓✔✗✘❖‣⁃∙⦁⚫⬛❯"
WEAK_BULLETS = "–—-·*+>»"
BULLET_CHARS = STRONG_BULLETS + WEAK_BULLETS

_NUM_PATTERNS = [
    (re.compile(r"^(\d{1,3})\.$"), "arabic", "Period"),
    (re.compile(r"^(\d{1,3})\)$"), "arabic", "ParenR"),
    (re.compile(r"^\((\d{1,3})\)$"), "arabic", "ParenBoth"),
    (re.compile(r"^([ivxlcdm]{1,6})\.$"), "romanLc", "Period"),
    (re.compile(r"^([ivxlcdm]{1,6})\)$"), "romanLc", "ParenR"),
    (re.compile(r"^\(([ivxlcdm]{1,6})\)$"), "romanLc", "ParenBoth"),
    (re.compile(r"^([IVXLCDM]{1,6})\.$"), "romanUc", "Period"),
    (re.compile(r"^([IVXLCDM]{1,6})\)$"), "romanUc", "ParenR"),
    (re.compile(r"^\(([IVXLCDM]{1,6})\)$"), "romanUc", "ParenBoth"),
    (re.compile(r"^([a-z])\.$"), "alphaLc", "Period"),
    (re.compile(r"^([a-z])\)$"), "alphaLc", "ParenR"),
    (re.compile(r"^\(([a-z])\)$"), "alphaLc", "ParenBoth"),
    (re.compile(r"^([A-Z])\.$"), "alphaUc", "Period"),
    (re.compile(r"^([A-Z])\)$"), "alphaUc", "ParenR"),
    (re.compile(r"^\(([A-Z])\)$"), "alphaUc", "ParenBoth"),
]
_ROMAN = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}


def roman_value(s: str) -> int | None:
    s = s.lower()
    if not s or any(c not in _ROMAN for c in s):
        return None
    total = 0
    for i, c in enumerate(s):
        v = _ROMAN[c]
        total += -v if i + 1 < len(s) and _ROMAN[s[i + 1]] > v else v
    return total if 0 < total < 4000 and to_roman(total) == s else None


def to_roman(n: int) -> str:
    vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
            (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for v, r in vals:
        while n >= v:
            out += r
            n -= v
    return out


@dataclass
class Label:
    kind: str  # char | autonum | picture
    text: str  # as printed
    n_glyphs: int  # leading non-space glyphs that form the label
    pos: float  # along-position of the label (its origin)
    text_start: float  # along-position where the paragraph text starts
    confidence: float
    char: str | None = None
    family: str | None = None  # roman | alpha | arabic (autonum only, before resolving)
    scheme: str | None = None  # DrawingML autonum type, e.g. arabicPeriod
    value: int | None = None
    alts: dict[str, int] = field(default_factory=dict)  # scheme -> value, ambiguous labels
    image_id: str | None = None


def parse_number(text: str) -> dict[str, int]:
    """All autonum schemes a label could be, with its value in each."""
    out: dict[str, int] = {}
    for rx, kind, suffix in _NUM_PATTERNS:
        m = rx.match(text)
        if not m:
            continue
        tok = m.group(1)
        if kind == "arabic":
            v: int | None = int(tok)
        elif kind.startswith("roman"):
            v = roman_value(tok)
        else:
            v = ord(tok.lower()) - ord("a") + 1
        if v is not None and v > 0:
            out[f"{kind}{suffix}"] = v
    return out


def detect_label(words: list[tuple[str, float, float, int]], em: float) -> Label | None:
    """Label at the start of a line.

    ``words``: (text, start, end, n_glyphs) along the line; ``em``: line font size.
    """
    if len(words) < 2:
        return None
    text, a0, a1, n = words[0]
    nxt = words[1]
    gap = nxt[1] - a1
    if len(text) == 1 and text in BULLET_CHARS:
        strong = text in STRONG_BULLETS
        if not strong and gap < 0.2 * em:
            return None
        return Label("char", text, n, a0, nxt[1], 0.95 if strong else 0.75, char=text)
    if text[0] in STRONG_BULLETS and len(text) > 1:  # "•Text": bullet glued to the word
        return Label("char", text[0], 1, a0, a0 + (a1 - a0) / len(text), 0.7, char=text[0])
    alts = parse_number(text)
    if alts and gap >= 0.15 * em:
        scheme = min(alts, key=lambda k: _scheme_rank(k, alts[k]))
        return Label("autonum", text, n, a0, nxt[1], 0.85 if gap >= 0.3 * em else 0.7,
                     scheme=scheme, value=alts[scheme], alts=alts)
    return None


def _scheme_rank(s: str, value: int = 0) -> tuple[int, str]:
    """Default preference: arabic, then roman for small values (i, v, x), then alpha."""
    if s.startswith("arabic"):
        return (0, s)
    if s.startswith("roman"):
        return (1 if value <= 10 else 3, s)
    return (2, s)


def resolve_numbering(labels: list[Label]) -> None:
    """Pick one scheme for ambiguous labels ("i." roman or alpha) from their neighbours.

    ``labels`` are the autonum labels of one level of one block, in order. A scheme
    scores one point per other label that can be read in it, and two when that
    label continues the sequence (value + 1) from the label before it.
    """
    auto = [lb for lb in labels if lb.kind == "autonum"]
    for i, lb in enumerate(auto):
        if len(lb.alts) <= 1:
            continue
        prev = auto[i - 1] if i > 0 else None
        nxt = auto[i + 1] if i + 1 < len(auto) else None

        def score(scheme: str, lb: Label = lb, prev: Label | None = prev,
                  nxt: Label | None = nxt) -> tuple[int, int]:
            pts = sum(1 for o in auto if o is not lb and scheme in o.alts)
            v = lb.alts[scheme]
            if prev is not None and prev.alts.get(scheme) == v - 1:
                pts += 2
            if nxt is not None and nxt.alts.get(scheme) == v + 1:
                pts += 2
            return pts, -_scheme_rank(scheme, v)[0]

        best = max(sorted(lb.alts), key=score)
        lb.scheme, lb.value = best, lb.alts[best]


def assign_levels(positions: list[float], tol: float) -> list[int]:
    """Cluster label x positions; level = rank of the cluster (left to right)."""
    centers: list[float] = []
    for x in sorted(positions):
        if not centers or x - centers[-1] > tol:
            centers.append(x)
    out = []
    for x in positions:
        lvl = min(range(len(centers)), key=lambda i: abs(centers[i] - x))
        out.append(min(lvl, 8))
    return out
