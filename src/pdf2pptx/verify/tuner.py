"""Fidelity tuner (M5): render -> compare -> adjust -> re-render, per text block.

For every text block the tuner compares the PDF's lines with the lines PowerPoint
actually lays out:

* **Line breaks** come from COM ``text_layout`` (one ``TextShape`` per block: the
  writer names each box ``TextBlock <id>``). Both sides are reduced to offsets in
  the block's non-space character stream, so a break either matches or not.
* **Positions** come from a PDF round trip (PowerPoint's ``SaveAs`` PDF; glyphs are read
  only on the changed slides): exported glyphs inside a COM line box give that line's
  exact baseline and start x.

Fixes, in this order and bounded by the mode (``config.ModeLimits``):

1. line breaks wrong: box width; then character spacing (``spc``); then font size in
   0.1 pt steps, chosen from a wrap model built from PowerPoint's own word widths;
   as a last resort (``forced_line_breaks``, fidelity) a soft break at each PDF
   break position;
2. breaks right: exact line spacing (through ``units.ppt_spacing``, PowerPoint's
   rounding), then a vertical (and, for left-aligned blocks, horizontal) offset that
   puts the first baseline on the PDF's.

A block is re-measured after every change, and the best version seen is kept, so a
block never ends worse than it started. After the text rounds, a whole-slide pass
compares pixels of shapes and pictures and translates those with a clear, consistent
offset (kept only if their local SSIM improves).

Edits go through the IR and a rewrite of the PPTX (python-pptx), not COM: measured
in M5, COM edits save ~0.1 s per round but would make the measured file differ from
the shipped one. Every change is recorded in the IR (``source="tuner"`` and
``tuner:*`` notes) and in ``Document.metadata["pdf2pptx:tuner"]``.
"""

from __future__ import annotations

import copy
import ctypes
import itertools
import json
import math
import statistics
import time
from collections import defaultdict
from collections.abc import Sequence
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import numpy as np

from .. import log
from ..build.units import ppt_spacing
from ..build.writer import (
    LINE_BREAK,
    SINGLE_LINE,
    TEXT_SHAPE_PREFIX,
    TextOptions,
    _drop_hyphens,
    slide_size,
    write_pptx,
)
from ..config import ModeLimits
from ..ir import Document, Glyph, Image, Page, Paragraph, Shape, TextBlock

LOG = log.get(__name__)

SPC_EFFECT = 0.955  # measured (M5): rendered extra advance per char = 0.955 x spc
SPC_STEP = 0.01  # pt (spc is written in 1/100 pt)
SIZE_STEP = 0.1  # pt
FIT_MARGIN = 0.3  # pt kept from a wrap constraint by a chosen width
POS_OK = 0.3  # pt: a block whose lines are this close needs no position fix
PITCH_TOL = 0.3  # pt
MIN_SHIFT = 0.2  # pt: smaller offsets are not worth a change


def _ns(s: str) -> str:
    return "".join(c for c in s if not c.isspace())


# --------------------------------------------------------------------------- geometry


@dataclass(frozen=True)
class Xf:
    """Page pt <-> slide pt (uniform scale + offset, as the writer places pages)."""

    s: float = 1.0
    dx: float = 0.0
    dy: float = 0.0

    def to_page_x(self, v: float) -> float:
        return (v - self.dx) / self.s

    def to_page_y(self, v: float) -> float:
        return (v - self.dy) / self.s


def page_xf(doc: Document, page: Page) -> Xf:
    sw, sh, _ = slide_size(doc)
    if page.width <= 0 or page.height <= 0:
        return Xf()
    s = min(sw / page.width, sh / page.height)
    return Xf(s, (sw - page.width * s) / 2, (sh - page.height * s) / 2)


# --------------------------------------------------------------------------- the PDF side


@dataclass
class Line:
    """One line of a block, as offsets [start, end) in its non-space char stream."""

    para: int
    start: int
    end: int
    baseline: float | None = None  # page pt
    x0: float | None = None  # origin x of the first char, page pt
    x1: float | None = None  # right edge of the last glyph, page pt (PDF side)
    para_last: bool = False  # the paragraph's last line (its break is not a wrap)


def written_paragraphs(tb: TextBlock, keep_hyphens: bool = True) -> list[str]:
    """Each paragraph's text as the writer writes it."""
    out = []
    for p in tb.paragraphs:
        runs = p.runs if keep_hyphens else _drop_hyphens(p)
        out.append("".join(r.text for r in runs))
    return out


def pdf_lines(tb: TextBlock, glyphs: dict[str, Glyph], keep_hyphens: bool = True
              ) -> list[Line] | None:
    """The block's PDF lines from its runs' glyphs (None if they cannot be mapped)."""
    lines: list[Line] = []
    offset = 0
    for pi, para in enumerate(tb.paragraphs):
        runs = para.runs if keep_hyphens else _drop_hyphens(para)
        cur: list[tuple[Glyph, int, bool]] = []  # glyph, end offset, shifted run
        para_lines: list[list[tuple[Glyph, int, bool]]] = []
        for run in runs:
            n = len(_ns(run.text))
            gl = [glyphs[g] for g in run.glyph_ids if g in glyphs and _ns(glyphs[g].text)]
            weights = [len(_ns(g.text)) for g in gl]
            total = sum(weights)
            if not gl or total == 0:
                offset += n
                continue
            acc = 0
            for g, w in zip(gl, weights, strict=True):
                acc += w
                end = offset + (acc if total == n else round(acc * n / total))
                shifted = bool(run.baseline_shift)
                if cur and not shifted:
                    ref = [c[0] for c in cur if not c[2]] or [cur[0][0]]
                    if abs(g.baseline - statistics.median(c.baseline for c in ref)) > (
                            0.5 * max(g.size, 1.0)):
                        para_lines.append(cur)
                        cur = []
                cur.append((g, end, shifted))
            offset += n
        if cur:
            para_lines.append(cur)
        if not para_lines:
            if any(_ns(r.text) for r in runs):
                return None  # text without glyphs: offsets cannot be placed
            continue
        for k, pl in enumerate(para_lines):
            start = lines[-1].end if lines else 0
            base = [c[0].baseline for c in pl if not c[2]] or [c[0].baseline for c in pl]
            lines.append(Line(pi, start, max(c[1] for c in pl),
                              baseline=statistics.median(base), x0=pl[0][0].origin_x,
                              x1=max(c[0].bbox[2] for c in pl),
                              para_last=k == len(para_lines) - 1))
        lines[-1].end = max(lines[-1].end, offset)
    if not lines:
        return None
    lines[-1].end = offset
    if any(ln.end <= ln.start for ln in lines):
        return None
    return lines


# --------------------------------------------------------------------------- the PowerPoint side


@dataclass
class RtGlyph:
    """A glyph of the PowerPoint round-trip PDF, page pt."""

    ch: str
    cx: float
    cy: float
    origin_x: float
    baseline: float


def roundtrip_glyphs(pdf: Path, index: int, xf: Xf) -> list[RtGlyph]:
    """Glyphs of slide ``index`` of a PDF exported by PowerPoint, mapped to page pt."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as R

    doc = pdfium.PdfDocument(str(pdf))
    out: list[RtGlyph] = []
    try:
        page = doc[index]
        h = page.get_size()[1]
        tp = page.get_textpage()
        try:
            for i in range(tp.count_chars()):
                u = R.FPDFText_GetUnicode(tp.raw, i)
                if u <= 32 or R.FPDFText_IsGenerated(tp.raw, i) == 1:
                    continue
                cl, r, b, t = (ctypes.c_double() for _ in range(4))
                R.FPDFText_GetCharBox(tp.raw, i, cl, r, b, t)
                ox, oy = ctypes.c_double(), ctypes.c_double()
                R.FPDFText_GetCharOrigin(tp.raw, i, ox, oy)
                out.append(RtGlyph(chr(u), xf.to_page_x((cl.value + r.value) / 2),
                                   xf.to_page_y(h - (b.value + t.value) / 2),
                                   xf.to_page_x(ox.value), xf.to_page_y(h - oy.value)))
        finally:
            tp.close()
    finally:
        doc.close()
    return out


@dataclass
class PptWord:
    text: str
    x0: float  # slide pt
    x1: float
    start: int  # non-space offsets in the block stream
    end: int


@dataclass
class PptLine(Line):
    words: list[PptWord] = field(default_factory=list)
    text: str = ""
    bound: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)  # slide pt


def ppt_lines(shape: Any, stream: str) -> tuple[list[PptLine], bool]:
    """COM lines as offsets; second value: the words map onto ``stream`` exactly."""
    out: list[PptLine] = []
    off = 0
    words_ok = True
    for ln in shape.lines:
        t = _ns(ln.text)
        pl = PptLine(-1, off, off + len(t), text=ln.text, bound=tuple(ln.bound))
        k = off
        for w in ln.words:
            wt = _ns(w.text)
            if stream[k:k + len(wt)] != wt:
                words_ok = False
                break
            pl.words.append(PptWord(wt, w.x0, w.x1, k, k + len(wt)))
            k += len(wt)
        if k != off + len(t):
            words_ok = False
        off += len(t)
        out.append(pl)
    return out, words_ok


def attach_roundtrip(lines: Sequence[PptLine], glyphs: Sequence[RtGlyph], xf: Xf) -> None:
    """Baseline and first-char x of each PowerPoint line from the round-trip glyphs."""
    for ln in lines:
        left, top, w, h = ln.bound
        x0, x1 = xf.to_page_x(left) - 0.5, xf.to_page_x(left + w) + 0.5
        y0, y1 = xf.to_page_y(top) - 0.5, xf.to_page_y(top + h) + 0.5
        inside = [g for g in glyphs if x0 <= g.cx <= x1 and y0 <= g.cy <= y1]
        if not inside:
            continue
        ln.baseline = statistics.median(g.baseline for g in inside)
        first = _ns(ln.text)[:1]
        cands = [g.origin_x for g in inside if g.ch == first]
        ln.x0 = min(cands) if cands else None


# --------------------------------------------------------------------------- evaluation


@dataclass
class Eval:
    breaks: int  # break offsets present on one side only (or |line count diff|)
    mapped: bool  # PDF and PowerPoint streams agree
    pos: float | None  # median distance of line anchors (baseline, start x), pt
    dy_first: float | None = None
    dx: float | None = None
    dx_spread: float = 0.0  # max - min of the per-line horizontal offsets
    pitch: dict[int, tuple[float, float]] = field(default_factory=dict)  # para: (pdf, ppt)

    @property
    def ok(self) -> bool:
        return self.breaks == 0 and (self.pos is None or self.pos <= POS_OK)

    def key(self) -> tuple[int, float]:
        return self.breaks, round(self.pos or 0.0, 3)

    def better_than(self, other: Eval) -> bool:
        return self.breaks < other.breaks or (
            self.breaks == other.breaks and (self.pos or 0.0) < (other.pos or 0.0) - 0.05)

    def worse_than(self, other: Eval) -> bool:
        return self.breaks > other.breaks or (
            self.breaks == other.breaks and (self.pos or 0.0) > (other.pos or 0.0) + 0.1)


def evaluate(ref: list[Line], ppt: list[PptLine], mapped: bool) -> Eval:
    if not mapped:
        return Eval(abs(len(ref) - len(ppt)), False, None)
    rb = {ln.end for ln in ref[:-1]}
    pb = {ln.end for ln in ppt[:-1]}
    ev = Eval(len(rb ^ pb), True, None)
    if ev.breaks:
        return ev
    d: list[float] = []
    dys: list[float] = []
    dxs: list[float] = []
    for r, p in zip(ref, ppt, strict=True):
        p.para, p.para_last = r.para, r.para_last
        if r.baseline is None or p.baseline is None:
            continue
        dy = r.baseline - p.baseline
        dx = (r.x0 - p.x0) if r.x0 is not None and p.x0 is not None else 0.0
        dys.append(dy)
        dxs.append(dx)
        d.append(math.hypot(dx, dy))
    if d:
        ev.pos = statistics.median(d)
        ev.dy_first = dys[0]
        ev.dx = statistics.median(dxs)
        ev.dx_spread = max(dxs) - min(dxs)
    by_para: dict[int, list[tuple[Line, PptLine]]] = defaultdict(list)
    for r, p in zip(ref, ppt, strict=True):
        if r.baseline is not None and p.baseline is not None:
            by_para[r.para].append((r, p))
    for pi, pairs in by_para.items():
        if len(pairs) >= 2:
            n = len(pairs) - 1
            ev.pitch[pi] = ((pairs[-1][0].baseline - pairs[0][0].baseline) / n,  # type: ignore[operator]
                            (pairs[-1][1].baseline - pairs[0][1].baseline) / n)  # type: ignore[operator]
    if ev.pitch and any(abs(a - b) > PITCH_TOL for a, b in ev.pitch.values()):
        ev.dy_first = statistics.median(dys)
    return ev


# --------------------------------------------------------------------------- the wrap model


@dataclass
class _Tok:
    start: int
    end: int
    para: int
    spaces: int  # whitespace chars after it within the paragraph (0 = hyphen split)
    width: float = 0.0  # slide pt
    gap: float | None = None  # measured gap to the next token (slide pt)


def _tokens(paras: list[str]) -> list[_Tok]:
    toks: list[_Tok] = []
    off = 0
    for pi, text in enumerate(paras):
        i, n = 0, len(text)
        while i < n:
            if text[i].isspace():
                i += 1
                continue
            j = i
            while j < n and not text[j].isspace():
                j += 1
                if text[j - 1] == "-" and j < n and text[j].isalpha() and j - i > 1:
                    break  # PowerPoint may wrap after a hyphen
            k = j
            while k < n and text[k].isspace() and text[k] != LINE_BREAK:
                k += 1
            ln = len(text[i:j])
            toks.append(_Tok(off, off + ln, pi, k - j))
            off += ln
            i = k if k > j else j
    return toks


@dataclass
class WrapModel:
    """Per PDF line: required width and the width at which the next word would fit,
    as linear functions of (spc change, size factor). Slide pt, box-inner."""

    req: list[tuple[float, float, float]]  # (start, text span, chars)
    nxt: list[tuple[float, float, float]]  # (start, span incl. next word, chars)
    width: float  # current inner width (slide pt)
    prefer: float | None = None  # the PDF's justified column width, if any (slide pt)

    def interval(self, delta: float, f: float) -> tuple[float, float]:
        lo = max(s + f * w + SPC_EFFECT * delta * n for s, w, n in self.req)
        hi = min((s + f * w + SPC_EFFECT * delta * n for s, w, n in self.nxt), default=math.inf)
        return lo, hi


def wrap_model(tb: TextBlock, paras: list[str], ref: list[Line], ppt: list[PptLine],
               xf: Xf) -> WrapModel | None:
    """Build the model from PowerPoint's word boxes. None when it cannot be built."""
    # x extent of every char from the COM words (chars inside a word share it evenly)
    xs: dict[int, tuple[float, float, int]] = {}
    for li, ln in enumerate(ppt):
        for w in ln.words:
            n = max(w.end - w.start, 1)
            for k in range(w.start, w.end):
                a = w.x0 + (w.x1 - w.x0) * (k - w.start) / n
                b = w.x0 + (w.x1 - w.x0) * (k - w.start + 1) / n
                xs[k] = (a, b, li)
    total = ref[-1].end
    if len(xs) != total:
        return None
    toks = _tokens(paras)
    if not toks or toks[-1].end != total:
        return None
    for t in toks:
        pieces: dict[int, list[int]] = defaultdict(list)
        for k in range(t.start, t.end):
            pieces[xs[k][2]].append(k)
        t.width = sum(xs[ks[-1]][1] - xs[ks[0]][0] for ks in pieces.values())
    ml, _, mr, _ = tb.margins
    box_left = tb.bbox[0] * xf.s + xf.dx + ml * xf.s
    inner = (tb.bbox[2] - tb.bbox[0] - ml - mr) * xf.s
    # natural gaps per paragraph, measured on unjustified lines
    per_space: dict[int, list[float]] = defaultdict(list)
    last_line: dict[int, int] = {}  # paragraph -> its last PowerPoint line
    for t in toks:
        last_line[t.para] = xs[t.end - 1][2]
    for a, b in itertools.pairwise(toks):
        if a.para != b.para or a.spaces == 0:
            continue
        la, lb = xs[a.end - 1][2], xs[b.start][2]
        stretched = tb.paragraphs[a.para].alignment in ("justify", "distributed") and (
            la != last_line[a.para])  # justified lines stretch their gaps; the last does not
        if la == lb and not stretched:
            a.gap = xs[b.start][0] - xs[a.end - 1][1]
            per_space[a.para].append(a.gap / a.spaces)
    all_gaps = [g for v in per_space.values() for g in v]
    sizes = [r.size for p in tb.paragraphs for r in p.runs if r.text.strip()]
    fallback = (statistics.median(all_gaps) if all_gaps
                else 0.25 * xf.s * (statistics.median(sizes) if sizes else 12.0))

    def gap(t: _Tok) -> float:
        if t.spaces == 0:
            return 0.0
        if t.gap is not None:
            return t.gap
        g = per_space.get(t.para)
        return t.spaces * (min(g) if g else fallback)

    # where lines start (box-inner): measured for left/justified paragraphs, else IR
    first_start: dict[int, float] = {}
    cont_start: dict[int, float] = {}
    for ln in ppt:
        if not ln.words:
            continue
        x = ln.words[0].x0 - box_left
        pi = next((t.para for t in toks if t.start <= ln.start < t.end), None)
        if pi is None:
            continue
        is_first = any(t.start == ln.start and (i == 0 or toks[i - 1].para != t.para)
                       for i, t in enumerate(toks))
        if tb.paragraphs[pi].alignment in ("left", "justify", "distributed"):
            (first_start if is_first else cont_start).setdefault(pi, x)

    def start_of(pi: int, first: bool) -> float:
        p: Paragraph = tb.paragraphs[pi]
        d = first_start if first else cont_start
        if pi in d:
            return d[pi]
        if first and p.bullet.kind == "none":
            return (p.indent_left + p.indent_first) * xf.s
        return (p.indent_left + max(0.0, p.indent_first if first and p.bullet.kind != "none"
                                    else 0.0)) * xf.s

    by_start = {t.start: i for i, t in enumerate(toks)}
    req: list[tuple[float, float, float]] = []
    nxt: list[tuple[float, float, float]] = []
    prev_para = -1
    for ln in ref:
        i0 = by_start.get(ln.start)
        if i0 is None:
            return None  # a PDF break inside a token
        i = i0
        span = 0.0
        chars = 0
        while i < len(toks) and toks[i].end <= ln.end:
            if i > i0:
                span += gap(toks[i - 1])
                chars += toks[i - 1].spaces
            span += toks[i].width
            chars += toks[i].end - toks[i].start
            i += 1
        if i == i0 or toks[i - 1].end != ln.end:
            return None
        st = start_of(ln.para, ln.para != prev_para)
        prev_para = ln.para
        req.append((st, span, chars))
        if not ln.para_last and i < len(toks) and toks[i].para == ln.para:
            t = toks[i]
            nxt.append((st, span + gap(toks[i - 1]) + t.width,
                        chars + toks[i - 1].spaces + t.end - t.start))
    # justified lines end on the column edge: that width also gives the PDF's spacing
    edges = [ln.x1 * xf.s + xf.dx - box_left for ln in ref
             if ln.x1 is not None and not ln.para_last
             and tb.paragraphs[ln.para].alignment in ("justify", "distributed")]
    return WrapModel(req, nxt, inner, statistics.median(edges) if edges else None)


def _deltas(limit: float) -> list[float]:
    n = round(limit / SPC_STEP)
    out = [0.0]
    for k in range(1, n + 1):
        out += [-k * SPC_STEP, k * SPC_STEP]
    return out


@dataclass
class WrapFix:
    width: float  # new inner width, slide pt
    delta: float = 0.0  # spc change, slide pt
    size: float = 0.0  # change of the main font size, page pt
    feasible: bool = True


def solve_wrap(model: WrapModel, w0: float, limits: ModeLimits, base_size: float,
               spc_room: tuple[float, float], s: float) -> WrapFix | None:
    """Width first, then spc, then size: the first feasible combination."""
    wmin, wmax = w0 * (1 - limits.max_width_change), w0 * (1 + limits.max_width_change)
    n_size = round(limits.max_font_size_change * base_size / SIZE_STEP + 1e-9)
    sizes = [0.0]
    for k in range(1, n_size + 1):
        sizes += [-k * SIZE_STEP, k * SIZE_STEP]
    deltas = [d for d in _deltas(limits.max_char_spacing_pt * s)
              if spc_room[0] - 1e-9 <= d <= spc_room[1] + 1e-9]
    for ds in sizes:
        f = (base_size + ds) / base_size
        for d in deltas:
            lo, hi = model.interval(d, f)
            lo_ok, hi_ok = max(lo + FIT_MARGIN, wmin), min(hi - FIT_MARGIN, wmax)
            if lo_ok > hi_ok:
                continue
            m = min(1.5, 0.3 * (hi - lo)) if math.isfinite(hi) else 1.5
            a, b = max(lo + m, wmin), min(hi - m, wmax)
            if a > b:
                a, b = lo_ok, hi_ok
            want = model.prefer if model.prefer is not None and d == 0 and ds == 0 else (
                model.width)
            return WrapFix(min(max(want, a), b), d, ds)
    return None


def best_effort_width(model: WrapModel, w0: float, limits: ModeLimits) -> float | None:
    """Width (within limits) violating the fewest wrap constraints, if fewer than now."""
    wmin, wmax = w0 * (1 - limits.max_width_change), w0 * (1 + limits.max_width_change)

    def violations(w: float) -> int:
        return (sum(1 for st, sp, _ in model.req if st + sp > w)
                + sum(1 for st, sp, _ in model.nxt if st + sp <= w))

    cands = {wmin, wmax, model.width}
    cands |= {st + sp + 0.5 for st, sp, _ in model.req}
    cands |= {st + sp - 0.5 for st, sp, _ in model.nxt}
    cands = {min(max(c, wmin), wmax) for c in cands}
    best = min(sorted(cands), key=lambda w: (violations(w), abs(w - model.width)))
    return best if violations(best) < violations(model.width) else None


# --------------------------------------------------------------------------- applying fixes


def _anchor(tb: TextBlock) -> str:
    al = {p.alignment for p in tb.paragraphs if any(r.text.strip() for r in p.runs)}
    if al == {"center"}:
        return "center"
    if al == {"right"}:
        return "right"
    return "left"


def set_width(tb: TextBlock, inner: float, xf: Xf) -> float:
    """New inner width (slide pt) -> bbox, keeping the alignment's anchor. Returns dw."""
    ml, _, mr, _ = tb.margins
    x0, y0, x1, y1 = tb.bbox
    new_w = inner / xf.s + ml + mr
    dw = new_w - (x1 - x0)
    anchor = _anchor(tb)
    if anchor == "center":
        x0 -= dw / 2
    elif anchor == "right":
        x0 -= dw
    tb.bbox = (round(x0, 3), y0, round(x0 + new_w, 3), y1)
    return dw


def main_size(tb: TextBlock) -> float:
    count: dict[float, int] = defaultdict(int)
    for p in tb.paragraphs:
        for r in p.runs:
            if not r.baseline_shift:
                count[r.size] += len(_ns(r.text))
    return max(sorted(count), key=lambda k: count[k]) if count else 12.0


def spc_room(tb: TextBlock, orig: TextBlock, limit: float, s: float) -> tuple[float, float]:
    """Allowed spc change (slide pt) so no run moves more than ``limit`` from the original."""
    lo, hi = -limit, limit
    for p, op in zip(tb.paragraphs, orig.paragraphs, strict=True):
        for r, orr in zip(p.runs, op.runs, strict=False):
            used = r.char_spacing - orr.char_spacing
            lo, hi = max(lo, -limit - used), min(hi, limit - used)
    return lo * s, hi * s


def apply_wrap(tb: TextBlock, fix: WrapFix, xf: Xf, base_size: float) -> list[str]:
    notes = []
    if fix.size:
        f = (base_size + fix.size) / base_size
        for p in tb.paragraphs:
            for r in p.runs:
                r.size = round(base_size + fix.size, 2) if r.size == base_size else round(
                    r.size * f, 2)
        notes.append(f"size:{fix.size:+.1f}")
    if fix.delta:
        for p in tb.paragraphs:
            for r in p.runs:
                r.char_spacing = round(r.char_spacing + fix.delta / xf.s, 2)
        notes.append(f"spc:{fix.delta / xf.s:+.2f}")
    dw = set_width(tb, fix.width, xf)
    if abs(dw) > 1e-3:
        notes.append(f"width:{dw:+.2f}")
    return notes


def force_breaks(tb: TextBlock, ref: list[Line]) -> int:
    """Soft breaks at every PDF wrap position. Returns how many were inserted."""
    by_para: dict[int, list[int]] = defaultdict(list)
    para_start: dict[int, int] = {}
    for ln in ref:
        para_start.setdefault(ln.para, ln.start)
        if not ln.para_last:
            by_para[ln.para].append(ln.end - para_start[ln.para])
    n = 0
    for pi, ends in by_para.items():
        para = tb.paragraphs[pi]
        chars = [(ri, ci) for ri, r in enumerate(para.runs) for ci in range(len(r.text))]
        texts = [list(r.text) for r in para.runs]
        cuts = []
        seen = 0
        targets = set(ends)
        for j, (ri, ci) in enumerate(chars):
            if not texts[ri][ci].isspace():
                seen += 1
                if seen in targets:
                    cuts.append(j)
        for j in reversed(cuts):
            k = j + 1
            if k < len(chars) and texts[chars[k][0]][chars[k][1]].isspace():
                ri, ci = chars[k]
                texts[ri][ci] = LINE_BREAK
            else:
                ri, ci = chars[j]
                texts[ri][ci] = texts[ri][ci] + LINE_BREAK
            n += 1
        for r, t in zip(para.runs, texts, strict=True):
            r.text = "".join(t)
    return n


# --------------------------------------------------------------------------- the driver


@dataclass
class _BlockState:
    page: int
    index: int  # position in page.elements
    orig: TextBlock
    ref: list[Line]
    best: TextBlock
    best_eval: Eval | None = None
    first_eval: Eval | None = None
    last_eval: Eval | None = None
    frozen: bool = False
    forced: bool = False
    moved: tuple[float, float] = (0.0, 0.0)  # the tuner's own offsets so far, pt
    bracket: tuple[float, float] = (-math.inf, math.inf)  # inner width: too narrow / wide
    too_narrow: bool = False
    steps: list[str] = field(default_factory=list)
    reverted: bool = False


@dataclass
class TuneReport:
    enabled: bool = True
    skipped: str | None = None
    rounds: int = 0
    seconds: float = 0.0
    slide_seconds: dict[int, float] = field(default_factory=dict)
    blocks: int = 0
    blocks_changed: int = 0
    reverted: int = 0
    breaks_before: int = 0
    breaks_after: int = 0
    matched_before: int = 0
    matched_after: int = 0
    forced_breaks: int = 0
    unmapped: int = 0
    shifted_shapes: int = 0
    slide_pass_rejected: int = 0
    timing: dict[str, float] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["slide_seconds"] = {str(k): round(v, 3) for k, v in sorted(self.slide_seconds.items())}
        d["timing"] = {k: round(v, 3) for k, v in sorted(self.timing.items())}
        d["seconds"] = round(self.seconds, 3)
        return d


class Tuner:
    def __init__(self, doc: Document, out: Path, work_dir: Path, pdf: Path, opts: TextOptions,
                 limits: ModeLimits, *, rounds: int = 4, engine: Any = None,
                 slide_pass: bool | None = None, render_dpi: int = 150) -> None:
        self.doc, self.out, self.work_dir, self.pdf = doc, out, work_dir, pdf
        self.opts, self.limits, self.rounds = opts, limits, max(0, rounds)
        self.engine = engine
        self.slide_pass = limits.slide_pass if slide_pass is None else slide_pass
        self.dpi = render_dpi
        self.rep = TuneReport()
        self.states: dict[str, _BlockState] = {}
        self.xfs = {p.index: page_xf(doc, p) for p in doc.pages}
        self.tmp = work_dir / "tune"
        self.stale = False  # the IR has changes the file on disk does not

    # -- bookkeeping ---------------------------------------------------------

    def _time(self, key: str, t0: float, slides: Sequence[int]) -> None:
        dt = time.perf_counter() - t0
        self.rep.timing[key] = self.rep.timing.get(key, 0.0) + dt
        for i in slides:
            self.rep.slide_seconds[i] = self.rep.slide_seconds.get(i, 0.0) + dt / max(
                len(slides), 1)

    def _write(self, slides: Sequence[int]) -> None:
        t0 = time.perf_counter()
        write_pptx(self.doc, self.out, self.work_dir, self.pdf, self.opts)
        self.stale = False
        self._time("write", t0, slides)

    def _block(self, st: _BlockState) -> TextBlock:
        e = self.doc.pages[st.page].elements[st.index]
        assert isinstance(e, TextBlock)
        return e

    def _set_block(self, st: _BlockState, tb: TextBlock) -> None:
        self.doc.pages[st.page].elements[st.index] = copy.deepcopy(tb)

    # -- setup ---------------------------------------------------------------

    def _collect(self) -> None:
        keep = self.opts.keep_hyphens
        for page in self.doc.pages:
            if page.triage.fallback_image:
                continue
            glyphs = {g.id: g for g in page.glyphs}
            for idx, e in enumerate(page.elements):
                if not isinstance(e, TextBlock) or e.rotation or e.vertical:
                    continue
                if not any(_ns(r.text) for p in e.paragraphs for r in p.runs):
                    continue
                if any(LINE_BREAK in r.text for p in e.paragraphs for r in p.runs):
                    continue
                ref = pdf_lines(e, glyphs, keep)
                if ref is None:
                    continue
                self.states[e.id] = _BlockState(page.index, idx, copy.deepcopy(e), ref,
                                                copy.deepcopy(e))
        self.rep.blocks = len(self.states)

    # -- measuring -----------------------------------------------------------

    def _measure(self, slides: list[int], rnd: int) -> dict[str, tuple[Eval, list[PptLine], bool]]:
        eng = self.engine
        t0 = time.perf_counter()
        info = eng.open(self.out)
        self._time("open", t0, slides)
        layouts = {}
        for i in slides:
            t0 = time.perf_counter()
            layouts[i] = {s.name: s for s in eng.text_layout(info, i)}
            self._time("text_layout", t0, [i])
        rt: dict[int, list[RtGlyph]] = {}
        t0 = time.perf_counter()
        try:
            # the whole deck: a ranged ExportAsFixedFormat crashed PowerPoint after some
            # decks (M5), SaveAs PDF has been stable since M1 and costs about the same
            path = eng.export_pdf(info, self.tmp / f"round{rnd}.pdf")
            self._time("export", t0, slides)
            t0 = time.perf_counter()
            for i in slides:
                rt[i] = roundtrip_glyphs(path, i, self.xfs[i])
        except Exception as exc:  # positions are optional: breaks still work
            LOG.warning("tuner: PDF round trip failed (%s); positions not tuned", exc)
        self._time("roundtrip", t0, slides)
        eng.close(info)
        out = {}
        for bid, st in sorted(self.states.items()):
            if st.page not in layouts or st.frozen:
                continue
            shape = layouts[st.page].get(f"{TEXT_SHAPE_PREFIX}{bid}")
            if shape is None:
                continue
            tb = self._block(st)
            paras = written_paragraphs(tb, self.opts.keep_hyphens)
            stream = "".join(_ns(t) for t in paras)
            pl, words_ok = ppt_lines(shape, stream)
            mapped = (sum(len(_ns(ln.text)) for ln in pl) == st.ref[-1].end == len(stream))
            if st.page in rt:
                attach_roundtrip(pl, rt[st.page], self.xfs[st.page])
            out[bid] = (evaluate(st.ref, pl, mapped), pl, words_ok and mapped)
        return out

    # -- fixing --------------------------------------------------------------

    def _fix(self, st: _BlockState, ev: Eval, pl: list[PptLine], model_ok: bool) -> bool:
        """Apply the next fix to the block. False when there is nothing left to try."""
        tb = self._block(st)
        xf = self.xfs[st.page]
        if ev.breaks:
            if not ev.mapped or st.forced:
                return False
            paras = written_paragraphs(tb, self.opts.keep_hyphens)
            model = wrap_model(tb, paras, st.ref, pl, xf) if model_ok else None
            ml, _, mr, _ = st.orig.margins
            w0 = (st.orig.bbox[2] - st.orig.bbox[0] - ml - mr) * xf.s
            self._observe(st, tb, pl, xf)
            if model is None:
                return self._nudge(st, tb, ev, pl, w0, xf)
            base = main_size(tb)
            room = spc_room(tb, st.orig, self.limits.max_char_spacing_pt, xf.s)
            fix = solve_wrap(model, w0, self.limits, base, room, xf.s)
            if fix is not None:
                if fix.delta == 0 and fix.size == 0:
                    fix.width = self._bracketed(st, fix.width, w0)
                else:
                    st.bracket = (-math.inf, math.inf)  # text widths change: start over
                notes = apply_wrap(tb, fix, xf, base)
                st.steps.extend(notes)
                # the model says the current state already fits: it is off, so step
                return bool(notes) or self._nudge(st, tb, ev, pl, w0, xf)
            if self.limits.forced_line_breaks:
                n = force_breaks(tb, st.ref)
                lo, _ = model.interval(0.0, 1.0)
                if lo + FIT_MARGIN > model.width:
                    set_width(tb, lo + 1.0, xf)
                st.forced = True
                st.steps.append(f"br:{n}")
                self.rep.forced_breaks += n
                return n > 0
            w = best_effort_width(model, w0, self.limits)
            if w is None:
                return False
            dw = set_width(tb, w, xf)
            st.steps.append(f"width*:{dw:+.2f}")
            return True
        if ev.pos is None or ev.ok:
            return False
        # 1. line spacing: exact pitch where PowerPoint's differs from the PDF's
        changed = False
        for pi, (want, got) in sorted(ev.pitch.items()):
            if abs(want - got) <= PITCH_TOL:
                continue
            para = tb.paragraphs[pi]
            single = SINGLE_LINE * xf.s * max((r.size for r in para.runs if not r.baseline_shift),
                                              default=12.0)
            real = ppt_spacing(want * xf.s, single)[2] / xf.s
            if abs(real - want) < abs(got - want) - 0.15 and not (
                    para.line_spacing_rule == "exact" and para.line_spacing == round(want, 2)):
                para.line_spacing_rule, para.line_spacing = "exact", round(want, 2)
                st.steps.append(f"lnspc:{pi}={want:.2f}")
                changed = True
        if changed:
            return True
        # 2. offset: first baseline; start x too when every line agrees on it
        dy = ev.dy_first or 0.0
        dx = (ev.dx or 0.0) if ev.dx_spread <= 2 * MIN_SHIFT else 0.0
        dy = 0.0 if abs(dy) < MIN_SHIFT else dy
        dx = 0.0 if abs(dx) < MIN_SHIFT else dx
        cx, cy = st.moved
        lim = self.limits.max_offset_pt
        dx = max(-lim - cx, min(lim - cx, dx))
        dy = max(-lim - cy, min(lim - cy, dy))
        if abs(dx) < 1e-3 and abs(dy) < 1e-3:
            return False
        x0, y0, x1, y1 = tb.bbox
        tb.bbox = (round(x0 + dx, 3), round(y0 + dy, 3), round(x1 + dx, 3), round(y1 + dy, 3))
        st.moved = (cx + dx, cy + dy)
        st.steps.append(f"move:{dx:+.2f},{dy:+.2f}")
        return True

    @staticmethod
    def _inner(tb: TextBlock, xf: Xf) -> float:
        ml, _, mr, _ = tb.margins
        return (tb.bbox[2] - tb.bbox[0] - ml - mr) * xf.s

    def _observe(self, st: _BlockState, tb: TextBlock, pl: list[PptLine], xf: Xf) -> None:
        """Narrow the width bracket from what PowerPoint did at the current width.

        At the first break where the two sides differ, the lines before it agree. If
        PowerPoint breaks there and the PDF does not, the PDF's line did not fit: the box
        is too narrow. Otherwise PowerPoint fitted more: the box is too wide."""
        rb = {ln.end for ln in st.ref[:-1]}
        pb = {ln.end for ln in pl[:-1]}
        if not rb ^ pb:
            return
        cur = self._inner(tb, xf)
        lo, hi = st.bracket
        if min(rb ^ pb) in pb:
            lo = max(lo, cur)
        else:
            hi = min(hi, cur)
        st.bracket = (lo, hi)
        st.too_narrow = min(rb ^ pb) in pb

    def _bracketed(self, st: _BlockState, target: float, w0: float) -> float:
        """``target`` if it lies inside the observed bracket, else bisect / step."""
        lo, hi = st.bracket
        lim = self.limits.max_width_change
        wmin, wmax = w0 * (1 - lim), w0 * (1 + lim)
        if lo < target < hi:
            out = target
        elif math.isfinite(lo) and math.isfinite(hi):
            out = (lo + hi) / 2
        elif math.isfinite(lo):
            out = lo * 1.015
        else:
            out = hi / 1.015
        return min(max(out, wmin), wmax)

    def _nudge(self, st: _BlockState, tb: TextBlock, ev: Eval, pl: list[PptLine], w0: float,
               xf: Xf) -> bool:
        """Model-free width step inside the bracket (wider if the box was too narrow)."""
        if not tb.wrap:
            return False
        cur = self._inner(tb, xf)
        new = self._bracketed(st, cur * (1.015 if st.too_narrow else 1 / 1.015), w0)
        if abs(new - cur) < 1e-3:
            return False
        dw = set_width(tb, new, xf)
        st.steps.append(f"width~:{dw:+.2f}")
        return True

    # -- main loop -------------------------------------------------------------

    def run(self) -> TuneReport:
        t_start = time.perf_counter()
        self._collect()
        if not self.states:
            self.rep.skipped = "no tunable text blocks"
            return self._finish(t_start)
        dirty = sorted({st.page for st in self.states.values()})
        for rnd in range(self.rounds + 1):
            evals = self._measure(dirty, rnd)
            self.rep.rounds = rnd
            changed: set[int] = set()
            for bid, (ev, pl, model_ok) in sorted(evals.items()):
                st = self.states[bid]
                if st.first_eval is None:
                    st.first_eval = ev
                st.last_eval = ev
                if st.best_eval is None or ev.better_than(st.best_eval):
                    st.best_eval, st.best = ev, copy.deepcopy(self._block(st))
                elif ev.worse_than(st.best_eval):
                    self._set_block(st, st.best)  # keep the best version seen
                    st.frozen = st.reverted = True
                    st.steps.append("revert")
                    self.stale = True
                    continue
                if ev.ok:
                    st.frozen = True
                    continue
                if rnd == self.rounds:
                    continue
                if self._fix(st, ev, pl, model_ok):
                    changed.add(st.page)
                else:
                    st.frozen = True
            if not changed:
                break
            dirty = sorted(changed)
            self._write(dirty)
        # every change was measured (the loop measures after each write); a block whose
        # measured best is not its last version goes back to the best
        for st in self.states.values():
            if (st.best_eval is not None and st.last_eval is not None and not st.reverted
                    and st.best_eval.better_than(st.last_eval)):
                self._set_block(st, st.best)
                st.reverted = True
                self.stale = True
        if self.slide_pass:
            try:
                self._slide_pass()
            except Exception as exc:
                LOG.warning("tuner: slide pass failed (%s); skipped", exc)
        self._record()
        if self.stale:
            self._write(sorted({st.page for st in self.states.values()}))
        return self._finish(t_start)

    def _finish(self, t0: float) -> TuneReport:
        self.rep.seconds = time.perf_counter() - t0
        self.doc.metadata["pdf2pptx:tuner"] = json.dumps(self.rep.as_dict(), sort_keys=True)
        return self.rep

    def _record(self) -> None:
        """Net adjustments -> IR (source="tuner", ``tuner:*`` notes) and the report."""
        rep = self.rep
        for st in self.states.values():
            tb = self._block(st)
            fe, be = st.first_eval, st.best_eval
            if fe is not None:
                rep.breaks_before += fe.breaks
                rep.matched_before += fe.ok
                rep.unmapped += not fe.mapped
            if be is not None:
                rep.breaks_after += be.breaks
                rep.matched_after += be.ok
            rep.reverted += st.reverted
            notes = _diff_notes(st.orig, tb)
            if not notes:
                if be is not None and not be.ok:
                    tb.notes = [*tb.notes, f"tuner:unresolved:breaks={be.breaks}"
                                + (f",pos={be.pos:.2f}" if be.pos is not None else "")]
                continue
            rep.blocks_changed += 1
            tb.notes = [*tb.notes, f"tuner:from:{st.orig.source}", *notes]
            if fe is not None and be is not None:
                tb.notes.append(f"tuner:eval:breaks {fe.breaks}->{be.breaks}"
                                + (f",pos {fe.pos:.2f}->{be.pos:.2f}"
                                   if fe.pos is not None and be.pos is not None else ""))
            tb.source = "tuner"

    # -- the whole-slide pixel pass --------------------------------------------

    def _slide_pass(self) -> None:
        from .diff import align
        from .metrics import ssim_map
        from .pdf_text import render_pages

        pages = [p for p in self.doc.pages if not p.triage.fallback_image and any(
            isinstance(e, (Shape, Image)) for e in p.elements)]
        if not pages:
            return
        slides = [p.index for p in pages]
        if self.stale:
            self._write(slides)
        t0 = time.perf_counter()
        pdf_imgs = render_pages(self.pdf, self.dpi)
        self._time("slide_pass", t0, slides)
        before = self._render(slides, "before")
        moves: dict[int, list[tuple[int, float, float]]] = {}
        t0 = time.perf_counter()
        for page in pages:
            img = before.get(page.index)
            if img is None or page.index >= len(pdf_imgs):
                continue
            ref = pdf_imgs[page.index]
            got = align(img, (page.width, page.height), self._slide_wh(),
                        (ref.shape[1], ref.shape[0]))
            found = []
            for idx, e in enumerate(page.elements):
                if not isinstance(e, (Shape, Image)) or e.is_background:
                    continue
                if "bullet_image" in e.notes:
                    continue
                shift = element_shift(ref, got, e.bbox, self.dpi)
                if shift is None:
                    continue
                dx, dy = shift
                if max(abs(dx), abs(dy)) > self.limits.max_offset_pt:
                    continue
                found.append((idx, dx, dy))
            if found:
                moves[page.index] = found
        self._time("slide_pass", t0, slides)
        if not moves:
            return
        backup = {pi: {idx: copy.deepcopy(self.doc.pages[pi].elements[idx]) for idx, _, _ in mv}
                  for pi, mv in moves.items()}
        for pi, mv in moves.items():
            for idx, dx, dy in mv:
                _translate(self.doc.pages[pi].elements[idx], dx, dy)
        changed = sorted(moves)
        self._write(changed)
        after = self._render(changed, "after")
        t0 = time.perf_counter()
        for pi, mv in moves.items():
            page = self.doc.pages[pi]
            ref = pdf_imgs[pi]
            a = align(before[pi], (page.width, page.height), self._slide_wh(),
                      (ref.shape[1], ref.shape[0]))
            b = align(after[pi], (page.width, page.height), self._slide_wh(),
                      (ref.shape[1], ref.shape[0]))
            _, ma = ssim_map(ref, a)
            _, mb = ssim_map(ref, b)
            k = self.dpi / 72.0
            for idx, dx, dy in mv:
                e = page.elements[idx]
                x0, y0, x1, y1 = backup[pi][idx].bbox
                xa = max(int(min(x0, x0 + dx) * k) - 2, 0)
                ya = max(int(min(y0, y0 + dy) * k) - 2, 0)
                xb, yb = int(max(x1, x1 + dx) * k) + 3, int(max(y1, y1 + dy) * k) + 3
                sa, sb = float(ma[ya:yb, xa:xb].mean()), float(mb[ya:yb, xa:xb].mean())
                if sb > sa + 1e-3:
                    e.notes = [*e.notes, f"tuner:from:{e.source}",
                               f"tuner:move:{dx:+.2f},{dy:+.2f}", f"tuner:ssim {sa:.3f}->{sb:.3f}"]
                    e.source = "tuner"
                    self.rep.shifted_shapes += 1
                else:
                    page.elements[idx] = backup[pi][idx]
                    self.rep.slide_pass_rejected += 1
                    self.stale = True
        self._time("slide_pass", t0, changed)

    def _slide_wh(self) -> tuple[float, float]:
        sw, sh, _ = slide_size(self.doc)
        return sw, sh

    def _render(self, slides: list[int], tag: str) -> dict[int, np.ndarray]:
        from PIL import Image as PILImage

        t0 = time.perf_counter()
        info = self.engine.open(self.out)
        paths = self.engine.render(info, slides, self.dpi, self.tmp / f"png_{tag}")
        self.engine.close(info)
        out = {i: np.asarray(PILImage.open(p).convert("RGB")) for i, p in
               zip(slides, paths, strict=True)}
        self._time("slide_pass", t0, slides)
        return out


def _translate(e: Any, dx: float, dy: float) -> None:
    x0, y0, x1, y1 = e.bbox
    e.bbox = (round(x0 + dx, 3), round(y0 + dy, 3), round(x1 + dx, 3), round(y1 + dy, 3))
    if isinstance(e, Shape) and e.path:
        for seg in e.path:
            seg.pts = [(round(px + dx, 3), round(py + dy, 3)) for px, py in seg.pts]


def element_shift(ref: np.ndarray, got: np.ndarray, bbox: Sequence[float], dpi: float,
                  radius: int = 6) -> tuple[float, float] | None:
    """Translation (pt) that best lines the PowerPoint render of ``bbox`` up with the PDF.

    None unless the match is clear (peak NCC >= 0.9, at least 0.03 above the unshifted
    one, and confirmed by the reverse match)."""
    import cv2

    from .metrics import to_gray

    k = dpi / 72.0
    h, w = ref.shape[:2]
    x0, y0 = math.floor(bbox[0] * k) - 2, math.floor(bbox[1] * k) - 2
    x1, y1 = math.ceil(bbox[2] * k) + 2, math.ceil(bbox[3] * k) + 2
    if x1 - x0 < 10 or y1 - y0 < 10 or (x1 - x0) * (y1 - y0) > 0.5 * w * h:
        return None
    if x0 - radius < 0 or y0 - radius < 0 or x1 + radius > w or y1 + radius > h:
        return None
    ga, gb = to_gray(ref).astype(np.float32), to_gray(got).astype(np.float32)
    tpl = ga[y0:y1, x0:x1]
    if float(tpl.std()) < 8.0:
        return None
    win = gb[y0 - radius:y1 + radius, x0 - radius:x1 + radius]
    res = cv2.matchTemplate(win, tpl, cv2.TM_CCOEFF_NORMED)
    _, peak, _, loc = cv2.minMaxLoc(res)
    sx, sy = loc[0] - radius, loc[1] - radius  # where the element is in PowerPoint
    zero = float(res[radius, radius])
    if (sx, sy) == (0, 0) or peak < 0.9 or peak - zero < 0.03:
        return None
    # reverse: the PowerPoint crop found in the PDF at the opposite shift
    tpl2 = gb[y0 + sy:y1 + sy, x0 + sx:x1 + sx]
    res2 = cv2.matchTemplate(ga[y0 + sy - radius:y1 + sy + radius,
                                x0 + sx - radius:x1 + sx + radius], tpl2, cv2.TM_CCOEFF_NORMED)
    _, _, _, loc2 = cv2.minMaxLoc(res2)
    if (loc2[0] - radius, loc2[1] - radius) != (-sx, -sy):
        return None
    return -sx / k, -sy / k  # move the element back onto the PDF's position


def _diff_notes(orig: TextBlock, tb: TextBlock) -> list[str]:
    notes = []
    ow, nw = orig.bbox[2] - orig.bbox[0], tb.bbox[2] - tb.bbox[0]
    if abs(nw - ow) > 1e-3:
        notes.append(f"tuner:width:{nw - ow:+.2f}")
    dx, dy = tb.bbox[0] - orig.bbox[0], tb.bbox[1] - orig.bbox[1]
    if abs(dy) > 1e-3 or (abs(dx) > 1e-3 and abs(nw - ow) <= 1e-3):
        notes.append(f"tuner:move:{dx:+.2f},{dy:+.2f}")
    for pi, (op, np_) in enumerate(zip(orig.paragraphs, tb.paragraphs, strict=True)):
        if (op.line_spacing_rule, op.line_spacing) != (np_.line_spacing_rule, np_.line_spacing):
            notes.append(f"tuner:lnspc:{pi}:{np_.line_spacing_rule}={np_.line_spacing}")
        for orr, nr in zip(op.runs, np_.runs, strict=True):
            if abs(nr.char_spacing - orr.char_spacing) > 1e-6:
                notes.append(f"tuner:spc:{nr.char_spacing - orr.char_spacing:+.2f}")
                break
        for orr, nr in zip(op.runs, np_.runs, strict=True):
            if abs(nr.size - orr.size) > 1e-6:
                notes.append(f"tuner:size:{nr.size - orr.size:+.2f}")
                break
        n_br = sum(r.text.count(LINE_BREAK) for r in np_.runs)
        if n_br:
            notes.append(f"tuner:br:{pi}:{n_br}")
    return sorted(set(notes), key=notes.index)


def tune_document(doc: Document, out: Path, work_dir: Path, pdf: Path, opts: TextOptions,
                  limits: ModeLimits, *, rounds: int = 4, engine: Any = None,
                  slide_pass: bool | None = None, render_dpi: int = 150) -> TuneReport:
    """Tune ``out`` (already written from ``doc``) in place. Never raises: without
    PowerPoint the deck is left as written and the report says why."""
    if opts.layout != "blocks":
        rep = TuneReport(enabled=False, skipped="text layout 'lines' is not tuned")
        return rep
    if engine is None:
        try:
            from . import ppt_com

            engine = ppt_com.shared()
        except Exception as exc:
            rep = TuneReport(enabled=False, skipped=f"PowerPoint unavailable: {exc}")
            doc.metadata["pdf2pptx:tuner"] = json.dumps(rep.as_dict(), sort_keys=True)
            return rep
    tuner = Tuner(doc, out, work_dir, pdf, opts, limits, rounds=rounds, engine=engine,
                  slide_pass=slide_pass, render_dpi=render_dpi)
    snapshot = copy.deepcopy(doc.pages)
    try:
        return tuner.run()
    except Exception as exc:  # never fail a conversion because of the tuner
        LOG.error("tuner failed (%s); keeping the untuned deck", exc,
                  exc_info=LOG.isEnabledFor(10))
        doc.pages = snapshot
        write_pptx(doc, out, work_dir, pdf, opts)
        rep = TuneReport(enabled=False, skipped=f"tuner error: {type(exc).__name__}: {exc}")
        doc.metadata["pdf2pptx:tuner"] = json.dumps(rep.as_dict(), sort_keys=True)
        return rep
