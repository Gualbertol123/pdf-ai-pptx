"""Scores from PLAN.md §6. Pure functions over arrays / word lists.

Word lists are sequences of lines; each line is a list of ``Word`` (text +
bbox in pt, top-left origin). The PDF side comes from ``verify/pdf_text.py``,
the PPTX side from PowerPoint's text layout (COM).

Word matching is order-independent: words are normalised to alphanumeric
tokens and, for each distinct token, PDF and PPTX occurrences are paired by a
minimum-total-distance assignment. ``word_recall``, ``pos_err_pt`` and
``linebreak_match`` all use the same pairing.
"""

from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter, defaultdict
from collections.abc import Sequence
from dataclasses import dataclass

import numpy as np
from rapidfuzz.distance import Levenshtein
from scipy.optimize import linear_sum_assignment
from skimage.metrics import structural_similarity

# Bullet glyphs are paragraph formatting in PPTX but literal glyphs in PDF text.
BULLET_CHARS = "•◦▪▫■□●○➢►▶✓❖·"
_BULLET_RE = re.compile(f"[{re.escape(BULLET_CHARS)}]")
_TOKEN_RE = re.compile(r"\w+", re.UNICODE)

COMPOSITE_WEIGHTS = {
    "ssim": 0.4,
    "text_sim": 0.2,
    "real_text_ratio": 0.2,
    "linebreak_match": 0.1,
    "native_shape_ratio": 0.1,
}


@dataclass(frozen=True)
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def center(self) -> tuple[float, float]:
        return ((self.x0 + self.x1) / 2, (self.y0 + self.y1) / 2)

    @property
    def anchor(self) -> tuple[float, float]:
        """Point used for matching and pos_err: horizontal centre, box bottom.

        Both sides put the bottom on the font's descent line (pdfplumber: font
        bbox; PowerPoint: line box minus space-after), while box tops differ
        (PowerPoint line boxes include leading and space-before).
        """
        return ((self.x0 + self.x1) / 2, self.y1)


Lines = Sequence[Sequence[Word]]


# --------------------------------------------------------------------------- text normalisation


def normalize_text(s: str) -> str:
    s = unicodedata.normalize("NFC", s)
    s = s.replace("­", "")  # soft hyphen
    s = s.replace("’", "'").replace("‘", "'")
    s = _BULLET_RE.sub(" ", s)
    return " ".join(s.split())


def tokens(s: str) -> list[str]:
    return [t.casefold() for t in _TOKEN_RE.findall(normalize_text(s))]


def _split_word(w: Word) -> list[Word]:
    """Split a word into alphanumeric tokens, sharing out its box by character offset."""
    text = normalize_text(w.text)
    n = max(len(text), 1)
    out = []
    for m in _TOKEN_RE.finditer(text):
        a, b = m.start() / n, m.end() / n
        out.append(
            Word(m.group().casefold(), w.x0 + (w.x1 - w.x0) * a, w.y0, w.x0 + (w.x1 - w.x0) * b,
                 w.y1)
        )
    return out


def _token_lines(lines: Lines) -> list[list[Word]]:
    return [[t for w in line for t in _split_word(w)] for line in lines]


# --------------------------------------------------------------------------- image metrics


def to_gray(img: np.ndarray) -> np.ndarray:
    a = np.asarray(img)
    if a.ndim == 3:
        a = a[..., :3].astype(np.float64) @ np.array([0.299, 0.587, 0.114])
    return a.astype(np.float64)


def ssim_map(a: np.ndarray, b: np.ndarray) -> tuple[float, np.ndarray]:
    """Grayscale SSIM (Gaussian window, sigma 1.5). Images must have equal size."""
    ga, gb = to_gray(a), to_gray(b)
    if ga.shape != gb.shape:
        raise ValueError(f"image sizes differ: {ga.shape} vs {gb.shape}")
    score, full = structural_similarity(
        ga, gb, data_range=255.0, gaussian_weights=True, sigma=1.5,
        use_sample_covariance=False, full=True,
    )
    return float(score), full


def region_scores(smap: np.ndarray, grid: int = 4) -> list[list[float]]:
    h, w = smap.shape
    ys = np.linspace(0, h, grid + 1).astype(int)
    xs = np.linspace(0, w, grid + 1).astype(int)
    return [
        [float(smap[ys[r]:ys[r + 1], xs[c]:xs[c + 1]].mean()) for c in range(grid)]
        for r in range(grid)
    ]


def region_ssim_min(smap: np.ndarray, grid: int = 4) -> float:
    return min(min(row) for row in region_scores(smap, grid))


# --------------------------------------------------------------------------- text metrics


def text_sim(pdf_text: str, pptx_text: str) -> float:
    """Normalised Levenshtein similarity on whitespace-normalised text."""
    a, b = normalize_text(pdf_text), normalize_text(pptx_text)
    if not a and not b:
        return 1.0
    return float(Levenshtein.normalized_similarity(a, b))


def token_recall(pdf_text: str, pptx_text: str, char_weighted: bool = False) -> float | None:
    """Multiset recall of PDF tokens in the PPTX text (order-free)."""
    pa, pb = Counter(tokens(pdf_text)), Counter(tokens(pptx_text))
    if not pa:
        return None
    if char_weighted:
        num = sum(min(c, pb[t]) * len(t) for t, c in pa.items())
        den = sum(c * len(t) for t, c in pa.items())
    else:
        num = sum(min(c, pb[t]) for t, c in pa.items())
        den = sum(pa.values())
    return num / den


def real_text_ratio(pdf_text: str, pptx_frame_text: str) -> float | None:
    """Share of visible (PDF) characters that exist as real text in PPTX text frames."""
    return token_recall(pdf_text, pptx_frame_text, char_weighted=True)


# --------------------------------------------------------------------------- word matching


@dataclass(frozen=True)
class Match:
    pdf: tuple[int, int]  # (line, index in line) in token lines
    pptx: tuple[int, int]
    dist: float


def match_words(pdf_lines: Lines, pptx_lines: Lines) -> tuple[list[Match], int]:
    """Pair equal tokens by minimum total centre distance. Returns (matches, n_pdf_tokens)."""
    pl, xl = _token_lines(pdf_lines), _token_lines(pptx_lines)
    by_tok_p: dict[str, list[tuple[int, int]]] = defaultdict(list)
    by_tok_x: dict[str, list[tuple[int, int]]] = defaultdict(list)
    for li, line in enumerate(pl):
        for wi, w in enumerate(line):
            by_tok_p[w.text].append((li, wi))
    for li, line in enumerate(xl):
        for wi, w in enumerate(line):
            by_tok_x[w.text].append((li, wi))
    n_pdf = sum(len(v) for v in by_tok_p.values())
    matches: list[Match] = []
    for tok in sorted(by_tok_p):
        ps, xs = by_tok_p[tok], by_tok_x.get(tok)
        if not xs:
            continue
        pc = np.array([pl[a][b].anchor for a, b in ps])
        xc = np.array([xl[a][b].anchor for a, b in xs])
        cost = np.linalg.norm(pc[:, None, :] - xc[None, :, :], axis=2)
        rows, cols = linear_sum_assignment(cost)
        for r, c in zip(rows, cols, strict=True):
            matches.append(Match(ps[r], xs[c], float(cost[r, c])))
    matches.sort(key=lambda m: m.pdf)
    return matches, n_pdf


def words_from_layout(lines: Sequence[object]) -> list[list[Word]]:
    """PowerPoint text layout (``ppt_com.LineBox`` list) -> word lines for matching.

    This is where ``pos_err_pt`` / ``linebreak_match`` get their PPTX side: one line
    per rendered PowerPoint line, words with COM ``Bound*`` boxes (slide pt).
    """
    return [[Word(w.text, w.x0, w.y0, w.x1, w.y1) for w in ln.words]  # type: ignore[attr-defined]
            for ln in lines]


def map_lines(lines: Lines, scale: float, dx: float, dy: float) -> list[list[Word]]:
    """Slide pt -> page pt for a page placed on the slide at ``scale`` + (dx, dy)."""
    s = scale or 1.0
    return [[Word(w.text, (w.x0 - dx) / s, (w.y0 - dy) / s, (w.x1 - dx) / s, (w.y1 - dy) / s)
             for w in line] for line in lines]


def word_recall(pdf_lines: Lines, pptx_lines: Lines) -> float | None:
    matches, n = match_words(pdf_lines, pptx_lines)
    return None if n == 0 else len(matches) / n


def pos_err_pt(pdf_lines: Lines, pptx_lines: Lines, max_dist: float | None = None) -> float | None:
    """Median distance (pt) between the anchors of matched words."""
    matches, _ = match_words(pdf_lines, pptx_lines)
    d = [m.dist for m in matches if max_dist is None or m.dist <= max_dist]
    return float(np.median(d)) if d else None


def linebreak_match(pdf_lines: Lines, pptx_lines: Lines) -> float | None:
    """Share of PDF lines whose last word also ends a line in PowerPoint."""
    pl, xl = _token_lines(pdf_lines), _token_lines(pptx_lines)
    matches, _ = match_words(pdf_lines, pptx_lines)
    to_x = {m.pdf: m.pptx for m in matches}
    total = hit = 0
    for li, line in enumerate(pl):
        if not line:
            continue
        total += 1
        x = to_x.get((li, len(line) - 1))
        if x is not None and x[1] == len(xl[x[0]]) - 1:
            hit += 1
    return None if total == 0 else hit / total


# --------------------------------------------------------------------------- structure metrics


def native_shape_ratio(n_preset: int, n_freeform: int) -> float | None:
    total = n_preset + n_freeform
    return None if total == 0 else n_preset / total


def composite(scores: dict[str, float | None]) -> float | None:
    """Weighted sum (PLAN.md §6); weights renormalised over the scores that exist."""
    num = den = 0.0
    for key, w in COMPOSITE_WEIGHTS.items():
        v = scores.get(key)
        if v is None or (isinstance(v, float) and math.isnan(v)):
            continue
        num += w * v
        den += w
    return None if den == 0 else num / den
