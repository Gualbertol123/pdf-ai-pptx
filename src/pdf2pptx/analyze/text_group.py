"""ANALYZE: glyphs -> words -> lines -> paragraphs -> TextBlocks.

All thresholds are relative to the local font size (em) or to statistics of the page
(median word gap, line pitch inside a block), never absolute points.

1. **Local frames.** Glyphs are grouped by rotation and projected onto (along, across)
   axes, so rotated text is grouped exactly like horizontal text.
2. **Lines.** Content order first (a glyph continues the current line if it sits on
   the same baseline and does not jump back or far ahead). Smaller glyphs within
   0.6 em of the baseline join as sub/superscripts. Same-baseline fragments that a
   PDF drew out of order are merged afterwards.
3. **Words / splits.** Space glyphs and gaps wider than 0.15 em separate words. A
   word gap wider than ``max(0.8 em, 3 x page median gap)`` splits the line (columns,
   table cells), except in lines whose gaps are uniformly wide (justified text).
4. **Labels.** Bullets, numbers and picture bullets at line starts (``lists.py``).
5. **Blocks.** A line joins the block whose last line is just above it: horizontal
   overlap, compatible size, and a baseline distance within ~2 line pitches. A line
   that would touch two blocks starts its own.
6. **Paragraphs.** A new paragraph starts at a label, at extra vertical space, at a
   size change, at an indent change, or after a line that ended early (the next
   line's first word would have fitted).
7. **Properties.** Alignment from line edges; indents from label / text positions;
   line spacing from the baselines (single when within 2% of PowerPoint's single
   pitch, else exact points); space after from the remaining gap; the frame width
   from the wrap constraints (wide enough for every line, narrow enough that every
   wrapped line still wraps).

PowerPoint geometry used for the frame (measured, see PROGRESS.md M2/M4):

* single spacing: line height 1.2 x size, first baseline at
  ``fonts.index.first_baseline_offset`` below the line top;
* other spacing (written as a percentage r of single: PowerPoint rounds ``spcPts``
  line and paragraph spacing to whole points, percentages are exact): line height
  r x single; baseline at 0.75 x the line height when r >= 1, else the single
  offset minus (1 - r) x single height; space after = percentage of single height;
* super/subscript runs render at 2/3 of their size, raised by ``baseline`` % of it.
"""

from __future__ import annotations

import dataclasses
import itertools
import math
import statistics
from dataclasses import dataclass, field

from ..build.units import ppt_spacing
from ..fonts.index import LINE_FACTOR, advance_em, first_baseline_offset
from ..fonts.match import FontMatch
from ..ingest import geom
from ..ir import (
    Annotation,
    Bullet,
    Color,
    FontRef,
    Glyph,
    Image,
    Page,
    Paragraph,
    TextBlock,
    TextRun,
)
from . import lists

SAME_BASELINE = 0.25  # em
SCRIPT_SIZE = 0.85  # a glyph this much smaller than the line may be a sub/superscript
SCRIPT_SHIFT = 0.6  # em of the line size: max baseline offset of a sub/superscript
SCRIPT_MIN_SHIFT = 0.08  # em: smaller shifts are noise, not scripts
BACKSTEP = 0.3  # em
MAX_JOIN_GAP = 3.0  # em: content order never continues a line across a wider gap
SPACE_GAP = 0.15  # em
SPLIT_MIN = 0.8  # em
WIDE_GAP = 0.3  # em beyond a space: kept as character spacing in one-line paragraphs
EXACT_BASELINE_RATIO = 0.75  # PowerPoint: baseline at 0.75 x line height (non-single)
SINGLE_TOL = 0.02  # pitch within 2% of 1.2 x size -> single spacing
LOW_CONFIDENCE = 0.6


# --------------------------------------------------------------------------- data


@dataclass
class LG:
    """A glyph in its local frame."""

    g: Glyph
    seq: int  # content order
    a: float  # origin along the text direction
    b: float  # baseline (across, downwards)
    lo: float  # ink extent along
    hi: float
    top: float  # ink extent across
    bot: float
    adv: float  # advance in pt
    space: bool
    script: bool = False


@dataclass
class Word:
    glyphs: list[LG]
    gap_before: float = 0.0  # pt between the previous word's end and this word's start

    @property
    def text(self) -> str:
        return "".join(x.g.text for x in self.glyphs)

    @property
    def a0(self) -> float:
        return self.glyphs[0].a

    @property
    def a1(self) -> float:
        return max(x.a + x.adv for x in self.glyphs)

    @property
    def natural(self) -> float:
        return sum(x.adv for x in self.glyphs)


@dataclass
class Line:
    glyphs: list[LG]
    base: float
    size: float  # largest non-script size
    words: list[Word] = field(default_factory=list)
    label: lists.Label | None = None
    space_adv: float = 0.0
    italic: bool = False  # most of the line's characters are italic

    @property
    def a0(self) -> float:
        return self.words[0].a0 if self.words else self.glyphs[0].a

    @property
    def a1(self) -> float:
        return max(w.a1 for w in self.words) if self.words else self.glyphs[-1].a

    @property
    def end(self) -> float:
        return max(x.a + max(x.adv, x.hi - x.a) for x in self.glyphs)

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)

    @property
    def text_start(self) -> float:
        return self.label.text_start if self.label else self.a0

    def text_words(self) -> list[Word]:
        if self.label is None or self.label.kind == "picture":
            return self.words
        return self.words[1:] if self.label.n_glyphs >= len(self.words[0].glyphs) else [
            Word(self.words[0].glyphs[self.label.n_glyphs:]), *self.words[1:]]

    def wide_extra(self) -> float:
        """Width beyond single spaces of the oversized gaps between text words."""
        ws = self.text_words()
        return sum(e for w in ws[1:] if (e := w.gap_before - self.space_adv)
                   > WIDE_GAP * self.size)

    def natural_text_width(self) -> float:
        ws = self.text_words()
        return sum(w.natural for w in ws) + self.space_adv * max(len(ws) - 1, 0)


@dataclass
class Para:
    lines: list[Line]
    reasons: list[str] = field(default_factory=list)
    confidence: float = 1.0
    alignment: str = "left"
    align_conf: float = 1.0
    level: int = 0


@dataclass
class Block:
    rot: float
    lines: list[Line]
    paras: list[Para] = field(default_factory=list)
    confidence: float = 1.0
    notes: list[str] = field(default_factory=list)

    @property
    def a0(self) -> float:
        return min(ln.a0 for ln in self.lines)

    @property
    def a1(self) -> float:
        return max(ln.a1 for ln in self.lines)


# --------------------------------------------------------------------------- context


@dataclass
class Ctx:
    page: Page
    fonts: dict[str, FontRef]
    matches: dict[str, FontMatch]
    rot: float = 0.0
    u: tuple[float, float] = (1.0, 0.0)
    n: tuple[float, float] = (0.0, 1.0)

    def family(self, font_id: str) -> tuple[str, bool, bool]:
        m = self.matches.get(font_id)
        if m is not None:
            return m.family, m.bold, m.italic
        fr = self.fonts.get(font_id)
        return (fr.family if fr else (font_id or "Arial"), bool(fr and fr.bold),
                bool(fr and fr.italic))

    def advance(self, g: Glyph, ch: str) -> float | None:
        fr = self.fonts.get(g.font)
        if fr is not None and ch in fr.widths and fr.widths[ch] > 0:
            return fr.widths[ch] * g.size
        fam, b, i = self.family(g.font)
        a = advance_em(fam, b, i, ch)
        return a * g.size if a is not None else None

    def local(self, x: float, y: float) -> tuple[float, float]:
        return x * self.u[0] + y * self.u[1], x * self.n[0] + y * self.n[1]

    def page_pt(self, a: float, b: float) -> tuple[float, float]:
        return a * self.u[0] + b * self.n[0], a * self.u[1] + b * self.n[1]


def _dirs(rot_ccw: float) -> tuple[tuple[float, float], tuple[float, float]]:
    r = math.radians(rot_ccw)
    return (math.cos(r), -math.sin(r)), (math.sin(r), math.cos(r))


def _to_local(ctx: Ctx, glyphs: list[tuple[int, Glyph]]) -> list[LG]:
    out = []
    for seq, g in glyphs:
        a, b = ctx.local(g.origin_x, g.baseline)
        x0, y0, x1, y1 = g.bbox
        corners = [ctx.local(x, y) for x, y in ((x0, y0), (x1, y0), (x0, y1), (x1, y1))]
        lo, hi = min(c[0] for c in corners), max(c[0] for c in corners)
        top, bot = min(c[1] for c in corners), max(c[1] for c in corners)
        space = g.text.isspace()
        adv = ctx.advance(g, g.text if not space else " ")
        if adv is None:
            adv = max(hi - a, 0.25 * g.size) if hi > lo else 0.25 * g.size
        out.append(LG(g=g, seq=seq, a=a, b=b, lo=lo, hi=hi, top=top, bot=bot, adv=adv,
                      space=space))
    return out


# --------------------------------------------------------------------------- lines


def _line_join(cur: Line, x: LG) -> bool:
    em = max(cur.size, x.g.size, 1.0)
    db = x.b - cur.base
    gap = x.a - cur.end
    if gap < -BACKSTEP * em or gap > MAX_JOIN_GAP * em:
        return False
    if abs(db) <= SAME_BASELINE * em:
        return True
    if x.space:
        return False
    if x.g.size <= SCRIPT_SIZE * cur.size and abs(db) <= SCRIPT_SHIFT * cur.size:
        return True
    # the line so far is all small glyphs (a leading superscript): rebase onto x
    return (cur.size <= SCRIPT_SIZE * x.g.size and abs(db) <= SCRIPT_SHIFT * x.g.size
            and all(not g.space for g in cur.glyphs))


def _form_lines(gs: list[LG]) -> list[Line]:
    lines: list[Line] = []
    cur: Line | None = None
    for x in gs:
        if cur is not None and _line_join(cur, x):
            cur.glyphs.append(x)
            if not x.space and abs(x.b - cur.base) <= SAME_BASELINE * max(cur.size, x.g.size):
                cur.size = max(cur.size, x.g.size)
            elif not x.space and x.g.size > cur.size:  # rebase on the bigger glyph
                cur.base, cur.size = x.b, x.g.size
            continue
        if x.space:
            continue
        cur = Line(glyphs=[x], base=x.b, size=x.g.size)
        lines.append(cur)
    for ln in lines:
        _mark_scripts(ln)
    return lines


def _mark_scripts(ln: Line) -> None:
    for x in ln.glyphs:
        x.script = (not x.space and x.g.size <= SCRIPT_SIZE * ln.size
                    and abs(x.b - ln.base) >= SCRIPT_MIN_SHIFT * ln.size)


def _words(ctx: Ctx, ln: Line) -> None:
    words: list[Word] = []
    cur: list[LG] = []
    pending_space = False
    prev_end = None
    for x in ln.glyphs:
        if x.space:
            pending_space = True
            continue
        if cur and prev_end is not None:
            gap = x.a - prev_end
            if pending_space or gap > SPACE_GAP * max(x.g.size, 1.0):
                words.append(Word(cur))
                cur = []
        if not cur and words and prev_end is not None:
            pass
        cur.append(x)
        prev_end = x.a + x.adv if prev_end is None else max(prev_end, x.a + x.adv)
        pending_space = False
    if cur:
        words.append(Word(cur))
    for i in range(1, len(words)):
        words[i].gap_before = words[i].a0 - words[i - 1].a1
    ln.words = words
    main = [x for x in ln.glyphs if not x.space and not x.script] or ln.glyphs
    ln.italic = 2 * sum(1 for x in main if ctx.family(x.g.font)[2]) > len(main)
    sp = ctx.advance(main[0].g, " ")
    ln.space_adv = sp if sp is not None else 0.25 * ln.size


def _gap_stats(lines: list[Line]) -> float:
    gaps = [w.gap_before / max(ln.size, 1.0) for ln in lines for w in ln.words[1:]]
    return statistics.median(gaps) if len(gaps) >= 5 else 0.25


def _split_line(ln: Line, thr_em: float) -> list[Line]:
    em = max(ln.size, 1.0)
    gaps = [w.gap_before for w in ln.words[1:]]
    if len(gaps) >= 4:
        med = statistics.median(gaps)
        if 0 < med <= 1.5 * em and max(gaps) <= 1.6 * med:
            return [ln]  # uniformly wide gaps: justified, not columns
    out: list[Line] = []
    start = 0
    for i in range(1, len(ln.words)):
        if ln.words[i].gap_before > thr_em * em:
            out.append(_subline(ln, ln.words[start:i]))
            start = i
    if not out:
        return [ln]
    out.append(_subline(ln, ln.words[start:]))
    return out


def _subline(ln: Line, words: list[Word]) -> Line:
    ids = {id(x) for w in words for x in w.glyphs}
    first, last = words[0].glyphs[0].seq, words[-1].glyphs[-1].seq
    glyphs = [x for x in ln.glyphs if id(x) in ids or (x.space and first < x.seq < last)]
    sub = Line(glyphs=glyphs, base=ln.base, size=max(
        (x.g.size for x in glyphs if not x.script and not x.space), default=ln.size),
        space_adv=ln.space_adv)
    sub.words = [Word(w.glyphs, w.gap_before) for w in words]
    sub.words[0].gap_before = 0.0
    return sub


def _merge_fragments(ctx: Ctx, lines: list[Line], thr_em: float) -> list[Line]:
    """Join same-baseline fragments that the PDF drew out of order."""
    lines = sorted(lines, key=lambda ln: (round(ln.base, 1), ln.a0))
    out: list[Line] = []
    for ln in lines:
        prev = out[-1] if out else None
        if prev is not None:
            em = max(prev.size, ln.size, 1.0)
            gap = ln.a0 - prev.a1
            if (abs(ln.base - prev.base) <= 0.1 * em and 0.8 < ln.size / prev.size < 1.25
                    and -BACKSTEP * em <= gap <= min(thr_em, 1.5) * em
                    and max(x.seq for x in prev.glyphs) > min(x.seq for x in ln.glyphs)):
                prev.glyphs = prev.glyphs + ln.glyphs
                prev.size = max(prev.size, ln.size)
                _mark_scripts(prev)
                _words(ctx, prev)
                continue
        out.append(ln)
    return out


# --------------------------------------------------------------------------- labels


def _detect_labels(ctx: Ctx, lines: list[Line], images: list[Image]) -> None:
    for ln in lines:
        words = [(w.text, w.a0, w.a1, len(w.glyphs)) for w in ln.words]
        ln.label = lists.detect_label(words, ln.size)
        if ln.label is None and ctx.rot == 0.0 and images:
            ln.label = _picture_bullet(ln, images)


def _picture_bullet(ln: Line, images: list[Image]) -> lists.Label | None:
    s = ln.size
    x_mid = ln.base - 0.35 * s
    for img in images:
        x0, y0, x1, y1 = img.bbox
        w, h = x1 - x0, y1 - y0
        if not (0 < h <= 1.6 * s and 0 < w <= 2.0 * s and 0.4 <= w / h <= 2.5):
            continue
        gap = ln.a0 - x1
        if -0.1 * s <= gap <= 1.5 * s and abs((y0 + y1) / 2 - x_mid) <= 0.5 * s:
            return lists.Label("picture", "", 0, x0, ln.a0, 0.7, image_id=img.id)
    return None


# --------------------------------------------------------------------------- blocks


def _overlap(a0: float, a1: float, b0: float, b1: float) -> float:
    return min(a1, b1) - max(a0, b0)


def _pitch(lines: list[Line]) -> float | None:
    ds = [b.base - a.base for a, b in itertools.pairwise(lines)
          if 0.85 < b.size / a.size < 1.18 and b.base > a.base]
    return statistics.median(ds) if ds else None


def _short_item(ln: Line) -> bool:
    """A short label-less line: at most 3 words and 8 em."""
    return (ln.label is None and len(ln.words) <= 3
            and ln.natural_text_width() <= 8.0 * max(ln.size, 1.0))


def _block_accepts(blk: Block, ln: Line) -> float | None:
    """Baseline distance if ``ln`` may continue ``blk``, else None."""
    last = blk.lines[-1]
    d = ln.base - last.base
    if d <= 0.3 * max(ln.size, last.size):
        return None
    big, small = max(ln.size, last.size), min(ln.size, last.size)
    ratio = big / max(small, 0.1)
    if ratio > 1.6:
        return None
    p = _pitch(blk.lines)
    limit = (2.0 * p if p else 2.0 * LINE_FACTOR * big) if ratio <= 1.3 else 1.5 * big
    if d > limit:
        return None
    if _short_item(last) and _short_item(ln) and d > 1.3 * LINE_FACTOR * big:
        return None  # spaced-out labels (axis ticks, table cells) are separate items
    if last.italic != ln.italic and ratio >= 1.05 and last.label is None:
        return None  # an italic caption next to upright body text (or the reverse)
    if _overlap(ln.a0, ln.a1, blk.a0, blk.a1) <= 0:
        return None
    return d


def _blocks(lines: list[Line], rot: float) -> list[Block]:
    blocks: list[Block] = []
    for ln in sorted(lines, key=lambda x: (x.base, x.a0)):
        cands = [(d, b) for b in blocks if (d := _block_accepts(b, ln)) is not None]
        if len(cands) > 1:
            # a line that touches two blocks at the same height spans columns
            cands.sort(key=lambda c: c[0])
            if cands[1][0] - cands[0][0] < 0.5 * ln.size:
                cands = []
        if cands:
            blk = min(cands, key=lambda c: c[0])[1]
            blk.lines.append(ln)
        else:
            blocks.append(Block(rot=rot, lines=[ln]))
    return blocks


# --------------------------------------------------------------------------- paragraphs


def _block_shape(lines: list[Line]) -> str:
    """left | center | right: which edge the lines of a block share."""
    if len(lines) < 2:
        return "left"
    tol = max(0.15 * max(ln.size for ln in lines), 1.5)

    def spread(v: list[float]) -> float:
        return max(v) - min(v)

    lefts = [ln.text_start for ln in lines[1:]]
    if spread(lefts) <= tol:
        return "left"
    if spread([(ln.text_start + ln.a1) / 2 for ln in lines]) <= 2 * tol:
        return "center"
    if spread([ln.a1 for ln in lines]) <= tol:
        return "right"
    return "left"


SENTENCE_END = (".", "!", "?", ":", ";", "…")


def _paragraphs(blk: Block) -> None:
    lines = blk.lines
    right = max(ln.a1 for ln in lines)
    left = min(ln.a0 for ln in lines)
    shape = _block_shape(lines)
    pitch = _pitch(lines)
    paras: list[Para] = [Para(lines=[lines[0]])]
    for prev, ln in itertools.pairwise(lines):
        reasons: list[str] = []
        conf = 1.0
        em = max(ln.size, 1.0)
        d = ln.base - prev.base
        exp = pitch or LINE_FACTOR * max(ln.size, prev.size)
        cont_left = paras[-1].lines[-1].text_start if len(paras[-1].lines) > 1 else (
            paras[-1].lines[0].text_start)
        if ln.label is not None:
            reasons.append("label")
        if d > 1.3 * exp:
            reasons.append("space")
        if not 0.87 < ln.size / max(prev.size, 0.1) < 1.15:
            reasons.append("size")
        first_w = ln.text_words()[0].natural if ln.text_words() else 0.0
        if shape == "left":
            slack = right - (prev.a1 + prev.space_adv + first_w)
        else:  # centred / right aligned: compare widths, not right edges
            slack = (right - left) - (prev.a1 - prev.text_start + prev.space_adv + first_w)
        if slack > 0.25 * em:
            reasons.append("short_line")
            if slack < 1.0 * em and not {"label", "space", "size"} & set(reasons):
                conf = 0.7  # close call
        if (shape == "left" and abs(ln.a0 - cont_left) > 0.5 * em and prev.label is None
                and not reasons):
            reasons.append("indent")
            conf = min(conf, 0.75)
        if (not reasons and len(paras[-1].lines) == 1 and prev.text.endswith(SENTENCE_END)
                and ln.text[:1].isupper() and prev.a1 >= right - 0.25 * em
                and ln.a1 - ln.text_start < 0.85 * (right - left)
                and not any(o is not prev and o.a1 >= right - 0.5 * em for o in lines)):
            # the widest line ends a sentence and a short line starting a sentence
            # follows: a wrapped paragraph and two paragraphs look the same; prefer two
            # (same layout). A nearly full next line, or another line reaching the same
            # right edge, shows the block's wrap width: then it is a wrap, no split.
            reasons.append("sentence")
            conf = 0.65
        if len(prev.text_words()) == 1 and not reasons and (
                len(ln.text_words()) == 1 or d > 1.35 * LINE_FACTOR * ln.size):
            # stacked single words (axis ticks, table columns) are not a wrapped paragraph
            reasons.append("single_word")
            conf = min(conf, 0.75)
        if ln.label is None and prev.label is not None and len(paras[-1].lines) == 1 and (
                abs(ln.a0 - prev.label.text_start) <= 0.5 * em):
            # a wrapped list item continues at the label's text start
            reasons = [r for r in reasons if r in ("label", "space", "size")]
        if prev.text.endswith(("-", "­")) and ln.words and ln.words[0].text[:1].islower():
            reasons = [r for r in reasons if r in ("label", "space", "size")]
        if reasons:
            paras.append(Para(lines=[ln], reasons=reasons, confidence=conf))
        else:
            paras[-1].lines.append(ln)
    blk.paras = paras


def _alignment(blk: Block) -> None:
    lft = blk.a0
    rgt = blk.a1
    cen = (lft + rgt) / 2
    for p in blk.paras:
        em = max(p.lines[0].size, 1.0)
        tol = max(0.15 * em, 1.5)
        lines = p.lines
        if len(lines) >= 2:
            cont = lines[1:]
            lefts = [ln.a0 for ln in cont]
            rights = [ln.a1 for ln in lines[:-1]]
            centers = [(ln.text_start + ln.a1) / 2 for ln in lines]
            stretched = any(ln.a1 - ln.text_start - ln.natural_text_width()
                            > max(0.25 * em, 1.5) for ln in lines[:-1])
            if (max(lefts) - min(lefts) <= tol and max(rights) - min(rights) <= tol
                    and stretched):
                p.alignment, p.align_conf = "justify", 0.9
            elif max(lefts) - min(lefts) <= tol and (
                    abs(lines[0].text_start - lefts[0]) <= tol or lines[0].label is not None
                    or lines[0].a0 > lefts[0]):
                p.alignment, p.align_conf = "left", 0.9
            elif max(centers) - min(centers) <= 2 * tol:
                p.alignment, p.align_conf = "center", 0.85
            elif max(ln.a1 for ln in lines) - min(ln.a1 for ln in lines) <= tol:
                p.alignment, p.align_conf = "right", 0.85
            else:
                p.alignment, p.align_conf = "left", 0.5
    # single-line paragraphs: vote on the edge most of them share
    singles = [p for p in blk.paras if len(p.lines) == 1]

    def fits(p: Para) -> dict[str, bool]:
        ln = p.lines[0]
        tol = max(0.15 * ln.size, 1.5)
        return {"left": abs(ln.a0 - lft) <= tol or ln.label is not None,
                "center": abs((ln.text_start + ln.a1) / 2 - cen) <= tol,
                "right": abs(ln.a1 - rgt) <= tol}

    votes = {k: sum(1 for p in singles if fits(p)[k]) for k in ("left", "center", "right")}
    pref = max(votes, key=lambda k: (votes[k], k == "left", k == "center"))
    for p in singles:
        f = fits(p)
        many = votes[pref] >= 2
        if f[pref]:
            p.alignment, p.align_conf = pref, 0.8 if many else 0.6
        elif any(f.values()):
            p.alignment = next(k for k in ("left", "center", "right") if f[k])
            p.align_conf = 0.7
        else:
            p.alignment, p.align_conf = "left", 0.5
    # single-line paragraphs follow a block that is clearly centred / right aligned
    multi = [p.alignment for p in blk.paras if len(p.lines) > 1 and p.align_conf >= 0.8]
    if multi and all(a == multi[0] for a in multi) and multi[0] in ("center", "right"):
        for p in blk.paras:
            if len(p.lines) == 1 and p.lines[0].label is None:
                p.alignment = multi[0]


def _single_line_block_alignment(ctx: Ctx, blk: Block, shapes: list[tuple[float, float, float,
                                                                            float]]) -> None:
    """A lone line centred in its enclosing shape or on the page is centred."""
    if len(blk.lines) != 1 or ctx.rot != 0.0 or blk.lines[0].label is not None:
        return
    ln = blk.lines[0]
    p = blk.paras[0]
    cx = (ln.a0 + ln.a1) / 2
    em = max(ln.size, 1.0)
    for x0, y0, x1, y1 in shapes:
        if x0 - 1 <= ln.a0 and ln.a1 <= x1 + 1 and y0 <= ln.base <= y1 and (
                x1 - x0) < 0.9 * ctx.page.width:
            if abs(cx - (x0 + x1) / 2) <= max(0.25 * em, 2.0):
                p.alignment, p.align_conf = "center", 0.8
            return
    w = ctx.page.width
    if abs(cx - w / 2) <= 0.015 * w and ln.a0 > 0.12 * w:
        p.alignment, p.align_conf = "center", 0.6


# --------------------------------------------------------------------------- geometry


def _text_glyphs(ln: Line) -> list[LG]:
    """The line's text glyphs (no label, no spaces, no sub/superscripts)."""
    out = [x for w in ln.text_words() for x in w.glyphs if not x.script]
    return out or [x for x in ln.glyphs if not x.space] or ln.glyphs


def _text_size(ln: Line) -> float:
    return _nice_size(max(x.g.size for x in _text_glyphs(ln)))


def _single_metrics(ctx: Ctx, ln: Line) -> tuple[float, float, float]:
    """(size, single line height, single first-baseline offset) of a line.

    The bullet glyph does not count: PowerPoint lays the line out with the text font.
    """
    big = max(_text_glyphs(ln), key=lambda x: x.g.size)
    size = _text_size(ln)
    return size, LINE_FACTOR * size, first_baseline_offset(ctx.family(big.g.font)[0], size)


def _line_metrics(ctx: Ctx, ln: Line, rule: str, value: float | None,
                  psize: float) -> tuple[float, float]:
    """(line height, baseline offset from the line top) in PowerPoint.

    ``value`` (rule "exact") is the pitch in pt; the writer stores it as a percentage
    of single spacing of the paragraph size ``psize`` (PowerPoint rounds ``spcPts`` to
    whole points). Measured: with a ratio r >= 1 the baseline sits at 0.75 x the line
    height; with r < 1 the line loses (1 - r) x single height at the top.
    """
    _size, single, b = _single_metrics(ctx, ln)
    if rule != "exact" or not value:
        return single, b
    r = value / (LINE_FACTOR * psize)
    h = r * single
    return h, (EXACT_BASELINE_RATIO * h if r >= 1.0 else b - (1.0 - r) * single)


def _ratio_for_offset(ctx: Ctx, ln: Line, offset: float) -> float | None:
    """Spacing ratio r that puts the line's baseline ``offset`` below its line top."""
    _size, single, b = _single_metrics(ctx, ln)
    if offset >= EXACT_BASELINE_RATIO * single:
        return offset / (EXACT_BASELINE_RATIO * single)
    r = 1.0 - (b - offset) / single
    return r if r >= 0.3 else None


def _para_size(p: Para) -> float:
    return max(_text_size(ln) for ln in p.lines)


def _nice_size(s: float) -> float:
    r = round(s * 2) / 2
    return r if abs(r - s) <= 0.004 * s else round(s, 2)


@dataclass
class _ParaGeom:
    rule: str = "auto"
    value: float | None = None
    space_after: float = 0.0
    marl: float = 0.0
    indent: float = 0.0


def _vertical(ctx: Ctx, blk: Block) -> tuple[list[_ParaGeom], float, float, list[str]]:
    """Per-paragraph spacing; frame top and bottom (local across)."""
    notes: list[str] = []
    geoms: list[_ParaGeom] = []
    sizes = [_para_size(p) for p in blk.paras]
    eff: list[float | None] = []  # the pitch PowerPoint will really use (after rounding)
    for p, ps in zip(blk.paras, sizes, strict=True):
        g = _ParaGeom()
        e = None
        if len(p.lines) >= 2:
            pitch = statistics.median(b.base - a.base for a, b in itertools.pairwise(p.lines))
            single = LINE_FACTOR * ps
            if abs(pitch - single) > SINGLE_TOL * single:
                g.rule, g.value = "exact", round(pitch, 2)
                e = ppt_spacing(pitch, single)[2]
        geoms.append(g)
        eff.append(e)

    def metrics(i: int, ln: Line) -> tuple[float, float]:
        g = geoms[i]
        return _line_metrics(ctx, ln, g.rule, eff[i] if eff[i] is not None else g.value,
                             sizes[i])

    # walk down the block predicting PowerPoint's baselines; each paragraph gap takes
    # the rounded space that best re-aligns the prediction with the PDF
    pred = blk.paras[0].lines[0].base
    for i, p in enumerate(blk.paras):
        if i > 0:
            ha, ba = metrics(i - 1, blk.paras[i - 1].lines[-1])
            _hb, bb = metrics(i, p.lines[0])
            want = p.lines[0].base - pred - ((ha - ba) + bb)
            ga = geoms[i - 1]
            if want >= 0.3:
                ga.space_after = round(ppt_spacing(want, LINE_FACTOR * sizes[i - 1])[2], 2)
            elif want < -0.3:
                r = _ratio_for_offset(ctx, p.lines[0], p.lines[0].base - pred - (ha - ba))
                if len(p.lines) == 1 and r is not None:
                    v = r * LINE_FACTOR * sizes[i]
                    geoms[i].rule, geoms[i].value = "exact", round(v, 2)
                    eff[i] = ppt_spacing(v, LINE_FACTOR * sizes[i])[2]
                    _hb, bb = metrics(i, p.lines[0])
                else:
                    notes.append(f"tight:{i}")
            pred = pred + (ha - ba) + bb + ga.space_after
        if len(p.lines) > 1:
            pitch = eff[i] if eff[i] is not None else LINE_FACTOR * sizes[i]
            pred += pitch * (len(p.lines) - 1)
    _h0, b0 = metrics(0, blk.paras[0].lines[0])
    top = blk.paras[0].lines[0].base - b0
    hl, bl = metrics(len(blk.paras) - 1, blk.paras[-1].lines[-1])
    bottom = pred + (hl - bl)
    if abs(pred - blk.paras[-1].lines[-1].base) > 0.25 * blk.paras[-1].lines[-1].size:
        notes.append(f"vdrift:{pred - blk.paras[-1].lines[-1].base:+.1f}")
    return geoms, top, bottom, notes


def _horizontal(blk: Block, geoms: list[_ParaGeom]) -> tuple[float, float, list[str]]:
    """Frame left and width (local along); sets indents on ``geoms``."""
    notes: list[str] = []
    left_paras = [p for p in blk.paras if p.alignment in ("left", "justify")]
    x0 = min(min(ln.a0 for ln in p.lines) for p in left_paras) if left_paras else blk.a0
    lo = 0.0
    hi = math.inf
    just_w = None
    for p, g in zip(blk.paras, geoms, strict=True):
        aligned_left = p.alignment in ("left", "justify")
        first = p.lines[0]
        if aligned_left:
            cont = p.lines[1:]
            text_left = min((ln.a0 for ln in cont), default=first.text_start)
            if first.label is not None:
                text_left = first.text_start if not cont else text_left
            g.marl = max(text_left - x0, 0.0)
            g.indent = (first.a0 if first.label is None else first.label.pos) - x0 - g.marl
            g.indent = max(g.indent, -g.marl)
        for i, ln in enumerate(p.lines):
            off = (g.marl + (g.indent if i == 0 and first.label is None else 0.0)
                   if aligned_left else 0.0)
            need = off + ln.natural_text_width() + (
                ln.wide_extra() if len(p.lines) == 1 and p.alignment != "justify" else 0.0)
            lo = max(lo, need)
            if i + 1 < len(p.lines):
                nxt = p.lines[i + 1].text_words()
                if nxt:
                    hi = min(hi, need + ln.space_adv + nxt[0].natural)
        if p.alignment == "justify" and len(p.lines) > 1:
            w = max(ln.a1 for ln in p.lines[:-1]) - x0
            just_w = w if just_w is None else max(just_w, w)
    em = max(blk.lines[0].size, 1.0)
    if just_w is not None and max(just_w, lo + 0.05) < hi:
        # justified: the column width, or just enough for a line the PDF squeezed
        # (some engines shrink spaces; PowerPoint cannot)
        width = max(just_w, lo + 0.05)
    elif hi < math.inf:
        if hi <= lo:
            width = lo + 0.5
            notes.append("wrap_conflict")
        else:
            width = lo + min((hi - lo) / 2, max(2.0, 0.03 * lo))
    else:
        width = lo + max(2.0, 0.04 * lo + 0.25 * em)
    if left_paras and just_w is None:
        # right / centred paragraphs next to left ones pin the frame's right edge
        rights = [ln.a1 for p in blk.paras if p.alignment == "right" for ln in p.lines]
        centers = [(ln.text_start + ln.a1) / 2 for p in blk.paras if p.alignment == "center"
                   for ln in p.lines]
        target = (max(rights) - x0 if rights else
                  2 * (statistics.median(centers) - x0) if centers else None)
        if target is not None and lo - 0.5 <= target < hi:
            width = max(target, lo)
    if not left_paras:
        cs = [(ln.text_start + ln.a1) / 2 for p in blk.paras for ln in p.lines]
        if all(p.alignment == "right" for p in blk.paras):
            x0 = max(ln.a1 for ln in blk.lines) - width
        else:
            x0 = statistics.median(cs) - width / 2
    return x0, width, notes


# --------------------------------------------------------------------------- runs


def _link_at(x: float, y: float, links: list[Annotation]) -> str | None:
    for a in links:
        x0, y0, x1, y1 = a.bbox
        if x0 - 1 <= x <= x1 + 1 and y0 - 1 <= y <= y1 + 1:
            if a.uri:
                return a.uri
            if a.dest_page is not None:
                return f"#slide={a.dest_page + 1}"
    return None


def _color_key(c: Color) -> tuple[str, float]:
    return (c.rgb or "000000", round(c.alpha, 3))


def _runs(ctx: Ctx, para: Para, pid: str, links: list[Annotation]) -> tuple[
        list[TextRun], list[str]]:
    """Runs of a paragraph: one per change of font, size, colour, script shift or link."""
    runs: list[TextRun] = []
    notes: list[str] = []
    last_key: list[tuple[object, ...] | None] = [None]

    keep_gaps = len(para.lines) == 1 and para.alignment != "justify"

    pending = [False]  # a space waits for the next glyph: it never joins a script run

    def space() -> None:
        if runs and not runs[-1].text.endswith(" "):
            pending[0] = True

    def wide_space(extra: float) -> None:
        # a single-line paragraph keeps an oversized gap (legend swatch, tab stop) as
        # character spacing on its space, so the next word starts where the PDF has it
        pending[0] = False
        prev = runs[-1]
        runs.append(dataclasses.replace(prev, id=f"{pid}_r{len(runs) + 1}", text=" ",
                                        char_spacing=round(extra, 2), glyph_ids=[]))
        last_key[0] = None

    def glyph(ln: Line, x: LG) -> None:
        if x.script:
            nominal = _nice_size(x.g.size * 1.5)
            shift = round((ln.base - x.b) / nominal * 100.0, 1)
        else:
            nominal, shift = _nice_size(x.g.size), 0.0
        cx, cy = (x.g.bbox[0] + x.g.bbox[2]) / 2, (x.g.bbox[1] + x.g.bbox[3]) / 2
        link = _link_at(cx, cy, links) if links else None
        fam, bold, italic = ctx.family(x.g.font)
        key = (fam, bold, italic, nominal, *_color_key(x.g.color), shift, link)
        lead = ""
        if pending[0]:
            pending[0] = False
            if runs[-1].baseline_shift and not shift:
                lead = " "  # the space goes with the following normal text
            else:
                runs[-1].text += " "
        if runs and key == last_key[0]:
            r = runs[-1]
            r.text += lead + x.g.text
            r.glyph_ids.append(x.g.id)
            b, gb = r.bbox, x.g.bbox
            r.bbox = (min(b[0], gb[0]), min(b[1], gb[1]), max(b[2], gb[2]), max(b[3], gb[3]))
            return
        runs.append(TextRun(id=f"{pid}_r{len(runs) + 1}", bbox=x.g.bbox, z=x.g.z,
                            source="rule", text=lead + x.g.text, font=fam, size=nominal,
                            color=x.g.color, bold=bold, italic=italic, baseline_shift=shift,
                            hyperlink=link, glyph_ids=[x.g.id]))
        last_key[0] = key

    for li, ln in enumerate(para.lines):
        words = ln.text_words()
        if li > 0 and words:
            prev = runs[-1].text if runs else ""
            if prev.endswith(("-", "­")) and words[0].text[:1].islower():
                notes.append(f"hyphen:{len(runs) - 1}:{len(prev) - 1}")
            else:
                space()
        for wi, w in enumerate(words):
            if wi > 0:
                extra = w.gap_before - ln.space_adv
                if keep_gaps and extra > WIDE_GAP * ln.size and runs:
                    wide_space(extra)
                else:
                    space()
            for x in w.glyphs:
                glyph(ln, x)
    while runs and not runs[-1].text.strip():
        runs.pop()
    if runs:
        runs[-1].text = runs[-1].text.rstrip(" ") or runs[-1].text
    return runs, notes


# --------------------------------------------------------------------------- build


def _bullet(ctx: Ctx, ln: Line) -> Bullet:
    lb = ln.label
    if lb is None:
        return Bullet()
    if lb.kind == "picture":
        return Bullet(kind="picture", image_id=lb.image_id, size_pct=100.0)
    lg = next((x for x in ln.glyphs if not x.space), None)
    text_lg = next((x for w in ln.text_words() for x in w.glyphs), None)
    size_pct = None
    color = None
    font = None
    if lg is not None:
        if text_lg is not None and text_lg.g.size > 0:
            pct = lg.g.size / text_lg.g.size * 100.0
            size_pct = round(pct, 1) if abs(pct - 100.0) > 1.0 else None
        color = lg.g.color
        font = ctx.family(lg.g.font)[0]
    if lb.kind == "autonum":
        return Bullet(kind="autonum", scheme=lb.scheme, start_at=lb.value or 1, color=color,
                      size_pct=size_pct, font=None)
    return Bullet(kind="char", char=lb.char, font=font, color=color, size_pct=size_pct)


def _levels_and_numbering(blk: Block) -> None:
    labelled = [p for p in blk.paras if p.lines[0].label is not None]
    if not labelled:
        return
    em = max(blk.lines[0].size, 1.0)
    pos = [p.lines[0].label.pos for p in labelled]  # type: ignore[union-attr]
    lvls = lists.assign_levels(pos, max(0.4 * em, 3.0))
    for p, lv in zip(labelled, lvls, strict=True):
        p.level = lv
    # text starts per level (for non-labelled continuation paragraphs)
    starts: dict[int, float] = {}
    for p in labelled:
        starts.setdefault(p.level, p.lines[0].text_start)
    for p in blk.paras:
        if p.lines[0].label is None:
            x = p.lines[0].a0
            fits = [lv for lv, s in starts.items() if s <= x + 0.4 * em]
            p.level = max(fits) if fits else 0
    by_level: dict[int, list[lists.Label]] = {}
    for p in labelled:
        lb = p.lines[0].label
        assert lb is not None
        if lb.kind == "autonum":
            by_level.setdefault(p.level, []).append(lb)
    for labs in by_level.values():
        lists.resolve_numbering(labs)
    # numbering runs: PowerPoint continues a run while type and startAt are equal
    expect: dict[int, tuple[str, int, int]] = {}  # level -> (scheme, start, next value)
    for p in blk.paras:
        for deeper in [k for k in expect if k > p.level]:
            del expect[deeper]
        lb = p.lines[0].label
        if lb is None or lb.kind != "autonum" or lb.scheme is None or lb.value is None:
            expect.pop(p.level, None)
            continue
        run = expect.get(p.level)
        if run and run[0] == lb.scheme and run[2] == lb.value:
            expect[p.level] = (run[0], run[1], lb.value + 1)
            lb.value = run[1]  # start_at of the run
        else:
            expect[p.level] = (lb.scheme, lb.value, lb.value + 1)


def group_page(page: Page, fonts: dict[str, FontRef], matches: dict[str, FontMatch]) -> list[
        TextBlock]:
    """TextBlocks for one page (not yet added to ``page.elements``)."""
    # invisible text (render mode 3/7, or a fully transparent fill: PowerPoint exports
    # tagged 1-pt copies of link text that way) never becomes a text box
    visible = [(i, g) for i, g in enumerate(page.glyphs)
               if g.render_mode not in (3, 7) and g.color.alpha > 0.01]
    if not visible:
        return []
    links = [a for a in page.annotations if a.subtype == "link"]
    images = [e for e in page.elements if isinstance(e, Image) and not e.is_background]
    shapes = [e.bbox for e in page.elements
              if getattr(e, "kind", "") == "shape" and getattr(e, "fill", None) is not None
              and e.fill.kind != "none"]  # type: ignore[union-attr]
    by_rot: dict[float, list[tuple[int, Glyph]]] = {}
    for i, g in visible:
        r = round(geom.norm_angle(g.rotation) * 2) / 2 % 360.0
        by_rot.setdefault(r, []).append((i, g))
    out: list[TextBlock] = []
    for rot in sorted(by_rot):
        u, n = _dirs(rot)
        ctx = Ctx(page=page, fonts=fonts, matches=matches, rot=rot, u=u, n=n)
        lgs = _to_local(ctx, by_rot[rot])
        lines = _form_lines(lgs)
        for ln in lines:
            _words(ctx, ln)
        lines = [ln for ln in lines if ln.words]
        thr = max(SPLIT_MIN, 3.0 * _gap_stats(lines))
        lines = [s for ln in lines for s in _split_line(ln, thr)]
        lines = _merge_fragments(ctx, lines, thr)
        _detect_labels(ctx, lines, images if rot == 0.0 else [])
        for blk in _blocks(lines, rot):
            _paragraphs(blk)
            _alignment(blk)
            _single_line_block_alignment(ctx, blk, shapes)
            _levels_and_numbering(blk)
            out.append(_to_block(ctx, blk, links, len(out) + 1))
    return out


def _to_block(ctx: Ctx, blk: Block, links: list[Annotation], k: int) -> TextBlock:
    pi = ctx.page.index
    bid = f"p{pi}_b{k}"
    geoms, top, bottom, vnotes = _vertical(ctx, blk)
    x0, width, hnotes = _horizontal(blk, geoms)
    notes = [*blk.notes, *vnotes, *hnotes]
    paragraphs: list[Paragraph] = []
    conf = 1.0
    for j, (p, g) in enumerate(zip(blk.paras, geoms, strict=True)):
        pid = f"{bid}_p{j + 1}"
        runs, rnotes = _runs(ctx, p, pid, links)
        if not runs:
            continue
        lb = p.lines[0].label
        pconf = min(p.confidence, p.align_conf, lb.confidence if lb else 1.0)
        conf = min(conf, pconf)
        xs = [c for ln in p.lines for x in ln.glyphs if not x.space for c in (
            x.g.bbox[0], x.g.bbox[2])]
        ys = [c for ln in p.lines for x in ln.glyphs if not x.space for c in (
            x.g.bbox[1], x.g.bbox[3])]
        pnotes = [f"split:{'+'.join(p.reasons)}"] if p.reasons else []
        pnotes += rnotes
        paragraphs.append(Paragraph(
            id=pid, bbox=(round(min(xs), 3), round(min(ys), 3), round(max(xs), 3),
                          round(max(ys), 3)),
            z=max(x.g.z for ln in p.lines for x in ln.glyphs), confidence=round(pconf, 3),
            source="rule", notes=pnotes, runs=runs, alignment=p.alignment,
            indent_left=round(g.marl, 2), indent_first=round(g.indent, 2), level=p.level,
            bullet=_bullet(ctx, p.lines[0]), line_spacing=g.value,
            line_spacing_rule=g.rule, space_after=g.space_after))
    if "wrap_conflict" in notes:
        conf = min(conf, 0.55)
    # frame: local (x0, top, width, height) -> unrotated page frame + rotation
    h = bottom - top
    ca, cb = x0 + width / 2, top + h / 2
    cx, cy = ctx.page_pt(ca, cb)
    frame = (round(cx - width / 2, 3), round(cy - h / 2, 3), round(cx + width / 2, 3),
             round(cy + h / 2, 3))
    rot_cw = round((-ctx.rot) % 360.0, 3)
    return TextBlock(
        id=bid, bbox=frame, z=_block_z(ctx, blk, frame),
        confidence=round(conf, 3), source="rule", notes=notes, paragraphs=paragraphs,
        autofit="none", wrap=True, rotation=0.0 if rot_cw >= 359.999 else rot_cw,
        margins=(0.0, 0.0, 0.0, 0.0))


def _block_z(ctx: Ctx, blk: Block, frame: tuple[float, float, float, float]) -> int:
    """z of the first glyph, unless a graphic drawn among the block's glyphs overlaps it.

    The first glyph keeps the PDF's drawing order (and so its text order); a filled
    graphic drawn between the block's first and last glyph that overlaps the frame
    would cover later lines, so then the block goes on top (z of the last glyph).
    """
    zs = [x.g.z for ln in blk.lines for x in ln.glyphs]
    lo, hi = min(zs), max(zs)
    if lo == hi:
        return lo
    fx0, fy0, fx1, fy1 = reading_aabb(frame, (-ctx.rot) % 360.0)
    for e in ctx.page.elements:
        if not lo < e.z < hi or e.kind not in ("shape", "image"):
            continue
        if e.kind == "shape" and getattr(e, "fill", None) is not None and \
                e.fill.kind == "none":  # type: ignore[union-attr]
            continue
        x0, y0, x1, y1 = e.bbox
        if x0 < fx1 and fx0 < x1 and y0 < fy1 and fy0 < y1:
            return hi
    return lo


def reading_aabb(frame: tuple[float, float, float, float], rot_cw: float) -> tuple[
        float, float, float, float]:
    from .reading_order import rotated_aabb

    return rotated_aabb(frame, rot_cw)


def picture_bullet_ids(blocks: list[TextBlock]) -> set[str]:
    return {p.bullet.image_id for b in blocks for p in b.paragraphs
            if p.bullet.kind == "picture" and p.bullet.image_id}
