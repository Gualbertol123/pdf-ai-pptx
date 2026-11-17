"""ANALYZE (tables): ruled grids and whitespace tables -> native ``Table`` elements.

Runs before text grouping, on the page's glyphs and vector paths:

* **Ruled tables.** Axis-aligned rules (stroked lines, stroked rectangles, thin filled
  rectangles) and cell fills (filled rectangles) that touch each other form a
  component. Its grid lines are the clustered rule positions and fill edges. Two
  neighbouring grid cells are separated when a rule covers at least half of their
  common edge or they lie in different fill rectangles; otherwise they are merged.
  Merged regions must be rectangles. Fills, borders (width, colour, dash) and the
  header row come from the same paths.
* **Whitespace tables.** Text lines split at wide gaps give segments; consecutive,
  evenly spaced rows with two or more segments whose x ranges leave common empty
  gutters form columns. Rules lying on row boundaries (booktabs style) become
  borders. Guards against prose columns and lists: short cells, a numeric column,
  no bullet labels.

Cell text is grouped by the M4 machinery (``text_group.group_page`` on the cell's
glyphs, with the column's alignment) and positioned through the cell margins.

Each table records a confidence; below ``MIN_CONFIDENCE`` nothing changes and the
page keeps its text boxes and shapes (no regression).
"""

from __future__ import annotations

import itertools
import re
import statistics
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

from .. import log
from ..fonts.match import FontMatch
from ..ir import (
    Color,
    Fill,
    FontRef,
    Glyph,
    Page,
    Paragraph,
    Shape,
    Stroke,
    Table,
    TableCell,
    TextBlock,
)
from . import text_group as tg
from .shapes import simplify, subpaths

LOG = log.get(__name__)

MIN_CONFIDENCE = 0.6
AXIS_TOL = 0.6  # pt: a segment this close to horizontal / vertical is axis-aligned
THIN_RULE = 2.5  # pt: a filled rectangle this thin is a rule
TOUCH = 2.0  # pt: rules / fills closer than this belong to one component
CLUSTER = 1.5  # pt: grid line positions closer than this are one line
COVER = 0.5  # share of a cell edge a rule must cover to separate the cells
MAX_ELEMENTS = 2000

_LABEL_RE = re.compile(r"^([•▪►✓\-–—·*]|\(?[0-9ivxlcIVXLC]{1,4}[.)]|\(?[a-zA-Z][.)])$")
_NUM_RE = re.compile(r"^[\s(+\-−–]*[€$£%]?\s*[0-9][0-9.,'\s]*[%€$£)]?\s*$")


@dataclass
class Seg:
    orient: str  # h | v
    pos: float
    a: float
    b: float
    width: float
    color: Color | None
    dash: str
    dash_pattern: list[float] | None
    sid: str

    def stroke(self) -> Stroke:
        return Stroke(color=self.color, width=round(self.width, 3), dash=self.dash,
                      dash_pattern=self.dash_pattern)


@dataclass
class FillRect:
    x0: float
    y0: float
    x1: float
    y1: float
    color: Color
    sid: str
    z: int


@dataclass
class _Comp:
    segs: list[Seg] = field(default_factory=list)
    fills: list[FillRect] = field(default_factory=list)


# --------------------------------------------------------------------------- rules


def _axis_rect(sh: Shape) -> tuple[float, float, float, float] | None:
    subs = subpaths(sh.path)
    if len(subs) != 1 or subs[0].curves:
        return None
    c = simplify(subs[0].corners, True)
    if len(c) != 4:
        return None
    xs = sorted({round(p[0], 1) for p in c})
    ys = sorted({round(p[1], 1) for p in c})
    if len(xs) != 2 or len(ys) != 2:
        return None
    return min(p[0] for p in c), min(p[1] for p in c), max(p[0] for p in c), max(p[1] for p in c)


def collect(page: Page) -> tuple[list[Seg], list[FillRect], dict[str, int]]:
    """Rules and cell fills on the page; ``{shape id: number of its segments}``."""
    segs: list[Seg] = []
    fills: list[FillRect] = []
    nseg: dict[str, int] = {}
    for e in page.elements:
        if not isinstance(e, Shape) or e.is_background or not e.path or e.rotation:
            continue
        if e.stroke is not None and e.fill.kind == "none":
            mine: list[Seg] = []
            ok = True
            for sub in subpaths(e.path):
                if sub.curves:
                    ok = False
                    break
                pts = list(sub.pts) + ([sub.pts[0]] if sub.closed else [])
                for (x0, y0), (x1, y1) in itertools.pairwise(pts):
                    if abs(y1 - y0) <= AXIS_TOL and abs(x1 - x0) > AXIS_TOL:
                        mine.append(Seg("h", (y0 + y1) / 2, min(x0, x1), max(x0, x1),
                                        e.stroke.width, e.stroke.color, e.stroke.dash,
                                        e.stroke.dash_pattern, e.id))
                    elif abs(x1 - x0) <= AXIS_TOL and abs(y1 - y0) > AXIS_TOL:
                        mine.append(Seg("v", (x0 + x1) / 2, min(y0, y1), max(y0, y1),
                                        e.stroke.width, e.stroke.color, e.stroke.dash,
                                        e.stroke.dash_pattern, e.id))
                    elif abs(x1 - x0) > AXIS_TOL or abs(y1 - y0) > AXIS_TOL:
                        ok = False
            if ok and mine:
                segs.extend(mine)
                nseg[e.id] = len(mine)
            continue
        if e.fill.kind == "solid" and e.fill.color is not None and e.stroke is None:
            r = _axis_rect(e)
            if r is None:
                continue
            x0, y0, x1, y1 = r
            w, h = x1 - x0, y1 - y0
            if h <= THIN_RULE and w > 3 * h:
                segs.append(Seg("h", (y0 + y1) / 2, x0, x1, h, e.fill.color, "solid", None,
                                e.id))
                nseg[e.id] = 1
            elif w <= THIN_RULE and h > 3 * w:
                segs.append(Seg("v", (x0 + x1) / 2, y0, y1, w, e.fill.color, "solid", None,
                                e.id))
                nseg[e.id] = 1
            elif w > THIN_RULE and h > THIN_RULE and e.fill.color.alpha >= 0.999 and (
                    w * h < 0.5 * page.width * page.height):
                fills.append(FillRect(x0, y0, x1, y1, e.fill.color, e.id, e.z))
    return segs, fills, nseg


def _box(x: Seg | FillRect) -> tuple[float, float, float, float]:
    if isinstance(x, FillRect):
        return x.x0, x.y0, x.x1, x.y1
    if x.orient == "h":
        return x.a, x.pos, x.b, x.pos
    return x.pos, x.a, x.pos, x.b


def components(segs: list[Seg], fills: list[FillRect]) -> list[_Comp]:
    items: list[Seg | FillRect] = [*segs, *fills]
    if len(items) > MAX_ELEMENTS:
        return []
    boxes = [_box(x) for x in items]
    parent = list(range(len(items)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    order = sorted(range(len(items)), key=lambda i: boxes[i][0])
    for k, i in enumerate(order):
        a = boxes[i]
        for j in order[k + 1:]:
            b = boxes[j]
            if b[0] > a[2] + TOUCH:
                break
            if b[1] <= a[3] + TOUCH and a[1] <= b[3] + TOUCH:
                parent[find(j)] = find(i)
    comps: dict[int, _Comp] = defaultdict(_Comp)
    for i, x in enumerate(items):
        c = comps[find(i)]
        (c.fills if isinstance(x, FillRect) else c.segs).append(x)
    return [comps[k] for k in sorted(comps)]


def _cluster(vals: list[float], tol: float = CLUSTER) -> list[float]:
    out: list[list[float]] = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [round(statistics.median(c), 3) for c in out]


def _coverage(segs: list[Seg], orient: str, pos: float, a: float, b: float) -> tuple[
        float, Seg | None]:
    """Share of [a, b] covered by rules of ``orient`` at ``pos``; the widest such rule."""
    if b <= a:
        return 0.0, None
    spans = []
    best = None
    for s in segs:
        if s.orient != orient or abs(s.pos - pos) > CLUSTER:
            continue
        lo, hi = max(a, s.a), min(b, s.b)
        if hi > lo:
            spans.append((lo, hi))
            if best is None or (hi - lo, s.width) > (min(b, best.b) - max(a, best.a), best.width):
                best = s
    spans.sort()
    tot, cur_lo, cur_hi = 0.0, None, None
    for lo, hi in spans:
        if cur_hi is None or lo > cur_hi:
            if cur_hi is not None and cur_lo is not None:
                tot += cur_hi - cur_lo
            cur_lo, cur_hi = lo, hi
        else:
            cur_hi = max(cur_hi, hi)
    if cur_hi is not None and cur_lo is not None:
        tot += cur_hi - cur_lo
    return tot / (b - a), best


def _fill_at(fills: list[FillRect], x: float, y: float) -> FillRect | None:
    hits = [f for f in fills if f.x0 - 0.1 <= x <= f.x1 + 0.1 and f.y0 - 0.1 <= y <= f.y1 + 0.1]
    return max(hits, key=lambda f: f.z) if hits else None


# --------------------------------------------------------------------------- grid model


@dataclass
class Grid:
    xs: list[float]
    ys: list[float]
    # (r, c) -> (row_span, col_span) for origins; covered cells map to their origin
    origin: dict[tuple[int, int], tuple[int, int]] = field(default_factory=dict)
    spans: dict[tuple[int, int], tuple[int, int]] = field(default_factory=dict)
    kind: str = "ruled"
    confidence: float = 1.0
    notes: list[str] = field(default_factory=list)

    @property
    def rows(self) -> int:
        return len(self.ys) - 1

    @property
    def cols(self) -> int:
        return len(self.xs) - 1

    def rect(self, r: int, c: int) -> tuple[float, float, float, float]:
        rs, cs = self.spans.get((r, c), (1, 1))
        return self.xs[c], self.ys[r], self.xs[c + cs], self.ys[r + rs]


def merge_regions(grid: Grid, sep_h: dict[tuple[int, int], bool],
                  sep_v: dict[tuple[int, int], bool]) -> bool:
    """Union cells without a separator; every region must be a rectangle.
    ``sep_v[(r, c)]``: boundary between (r, c) and (r, c+1); ``sep_h[(r, c)]``: between
    (r, c) and (r+1, c)."""
    n_r, n_c = grid.rows, grid.cols
    parent = {(r, c): (r, c) for r in range(n_r) for c in range(n_c)}

    def find(k: tuple[int, int]) -> tuple[int, int]:
        while parent[k] != k:
            parent[k] = parent[parent[k]]
            k = parent[k]
        return k

    for r in range(n_r):
        for c in range(n_c):
            if c + 1 < n_c and not sep_v[(r, c)]:
                parent[find((r, c + 1))] = find((r, c))
            if r + 1 < n_r and not sep_h[(r, c)]:
                parent[find((r + 1, c))] = find((r, c))
    regions: dict[tuple[int, int], list[tuple[int, int]]] = defaultdict(list)
    for k in parent:
        regions[find(k)].append(k)
    for cells in regions.values():
        rs = [k[0] for k in cells]
        cs = [k[1] for k in cells]
        r0, r1, c0, c1 = min(rs), max(rs), min(cs), max(cs)
        if len(cells) != (r1 - r0 + 1) * (c1 - c0 + 1):
            return False
        grid.spans[(r0, c0)] = (r1 - r0 + 1, c1 - c0 + 1)
        for k in cells:
            grid.origin[k] = (r0, c0)
    return True


# --------------------------------------------------------------------------- text helpers


def _visible(page: Page, skip: set[str]) -> list[Glyph]:
    return [g for g in page.glyphs if g.render_mode not in (3, 7) and g.color.alpha > 0.01
            and not g.text.isspace() and g.id not in skip and abs(g.rotation) < 0.5]


def _center(g: Glyph) -> tuple[float, float]:
    return (g.bbox[0] + g.bbox[2]) / 2, (g.bbox[1] + g.bbox[3]) / 2


def _glyphs_in(gs: list[Glyph], box: tuple[float, float, float, float]) -> list[Glyph]:
    x0, y0, x1, y1 = box
    return [g for g in gs if x0 <= _center(g)[0] <= x1 and y0 <= _center(g)[1] <= y1]


def _cell_text(page: Page, fonts: dict[str, FontRef], matches: dict[str, FontMatch],
               ids: set[str], align: str, prefix: str) -> TextBlock | None:
    if not ids:
        return None
    blocks = tg.group_page(page, fonts, matches, only=ids, align=align, prefix=prefix)
    if not blocks:
        return None
    blocks.sort(key=lambda b: (b.bbox[1], b.bbox[0]))
    first = blocks[0]
    if len(blocks) == 1:
        return first
    # several blocks in one cell: one frame, later blocks as paragraphs with space before
    paras: list[Paragraph] = list(first.paragraphs)
    bottom = first.bbox[3]
    for b in blocks[1:]:
        if b.paragraphs:
            b.paragraphs[0].space_before = round(max(b.bbox[1] - bottom, 0.0), 2)
        paras.extend(b.paragraphs)
        bottom = b.bbox[3]
    first.paragraphs = paras
    first.bbox = (min(b.bbox[0] for b in blocks), first.bbox[1],
                  max(b.bbox[2] for b in blocks), bottom)
    first.notes.append(f"cell_blocks:{len(blocks)}")
    first.confidence = min(first.confidence, 0.7)
    return first


def _alignments(grid: Grid, cell_glyphs: dict[tuple[int, int], list[Glyph]],
                header_rows: int) -> dict[tuple[int, int], str]:
    """Per cell: left | center | right, decided per column (body) and per header row."""
    def decide(items: list[tuple[tuple[float, float, float, float], list[Glyph]]]) -> str:
        lefts, rights, centers = [], [], []
        for (x0, _, x1, _), gs in items:
            if not gs:
                continue
            tx0, tx1 = min(g.bbox[0] for g in gs), max(g.bbox[2] for g in gs)
            lefts.append(tx0 - x0)
            rights.append(x1 - tx1)
            centers.append((tx0 + tx1) / 2 - (x0 + x1) / 2)
        if len(lefts) < 2:
            if lefts and abs(centers[0]) < 1.0 and min(lefts[0], rights[0]) > 3.0:
                return "center"
            return "left"
        sl = max(lefts) - min(lefts)
        sr = max(rights) - min(rights)
        sc = max(centers) - min(centers)
        if sl <= 1.5:
            return "left"
        if sr <= 1.5:
            return "right"
        if sc <= 2.0:
            return "center"
        return "left"

    out: dict[tuple[int, int], str] = {}
    for c in range(grid.cols):
        body = [(grid.rect(r, c), cell_glyphs.get((r, c), [])) for r in range(header_rows,
                                                                          grid.rows)
                if grid.origin.get((r, c)) == (r, c) and grid.spans.get((r, c), (1, 1))[1] == 1]
        a = decide(body)
        for r in range(header_rows, grid.rows):
            out[(r, c)] = a
    for r in range(header_rows):
        items = [(grid.rect(r, c), cell_glyphs.get((r, c), [])) for c in range(grid.cols)
                 if grid.origin.get((r, c)) == (r, c)]
        a = decide(items)
        for c in range(grid.cols):
            out[(r, c)] = a
    return out


def build_table(page: Page, fonts: dict[str, FontRef], matches: dict[str, FontMatch],
                grid: Grid, gs: list[Glyph], segs: list[Seg], fills: list[FillRect], tid: str,
                z: int) -> tuple[Table, set[str]] | None:
    cell_glyphs: dict[tuple[int, int], list[Glyph]] = {}
    used: set[str] = set()
    for (r, c) in sorted(grid.spans):
        inside = [g for g in _glyphs_in(gs, grid.rect(r, c)) if g.id not in used]
        cell_glyphs[(r, c)] = inside
        used |= {g.id for g in inside}
    n_text = sum(1 for v in cell_glyphs.values() if v)
    if n_text < 3 or n_text < 0.3 * len(grid.spans):
        return None
    # header rows: a distinct fill, bold text, or a heavier rule under the first row
    header = 0
    if grid.rows >= 3:
        f0 = _fill_at(fills, *_mid(grid.rect(0, 0)))
        f1 = _fill_at(fills, *_mid(grid.rect(grid.rows - 1, 0)))
        bold0 = _bold_share(cell_glyphs, fonts, matches, [0]) >= 0.8
        bold_b = _bold_share(cell_glyphs, fonts, matches, list(range(1, grid.rows))) >= 0.8
        if (f0 is not None and (f1 is None or f0.color.rgb != f1.color.rgb)) or (
                bold0 and not bold_b):
            header = max(rs for (r, _), (rs, _) in grid.spans.items() if r == 0)
    aligns = _alignments(grid, cell_glyphs, header)
    cells: list[TableCell] = []
    for r in range(grid.rows):
        for c in range(grid.cols):
            o = grid.origin[(r, c)]
            if o != (r, c):
                cells.append(TableCell(row=r, col=c, merged_into=o))
                continue
            rs, cs = grid.spans[(r, c)]
            x0, y0, x1, y1 = grid.rect(r, c)
            ids = {g.id for g in cell_glyphs.get((r, c), [])}
            text = _cell_text(page, fonts, matches, ids, aligns.get((r, c), "left"),
                              f"{tid}_r{r}c{c}_b")
            borders: dict[str, Stroke] = {}
            for side, (orient, pos, a, b) in {
                    "left": ("v", x0, y0, y1), "right": ("v", x1, y0, y1),
                    "top": ("h", y0, x0, x1), "bottom": ("h", y1, x0, x1)}.items():
                cov, s = _coverage(segs, orient, pos, a, b)
                if s is not None and cov >= COVER:
                    borders[side] = s.stroke()
            f = _fill_at(fills, (x0 + x1) / 2, (y0 + y1) / 2)
            fill = Fill(kind="solid", color=f.color) if f is not None else Fill()
            cells.append(TableCell(row=r, col=c, row_span=rs, col_span=cs, text=text,
                                   fill=fill, borders=borders,
                                   margins=_default_margins(text, (x0, y0, x1, y1))))
    t = Table(id=tid, bbox=(grid.xs[0], grid.ys[0], grid.xs[-1], grid.ys[-1]), z=z,
              source="rule", confidence=round(grid.confidence, 3),
              row_heights=[round(b - a, 3) for a, b in zip(grid.ys, grid.ys[1:], strict=False)],
              col_widths=[round(b - a, 3) for a, b in zip(grid.xs, grid.xs[1:], strict=False)],
              cells=cells, header_rows=header,
              notes=[f"table:{grid.kind}", f"grid:{grid.rows}x{grid.cols}", *grid.notes])
    return t, used


def _default_margins(text: TextBlock | None, rect: tuple[float, float, float, float]) -> tuple[
        float, float, float, float]:
    if text is None:
        return (min(7.2, (rect[2] - rect[0]) / 4), 1.0, min(7.2, (rect[2] - rect[0]) / 4), 0.0)
    return (round(max(text.bbox[0] - rect[0], 0.0), 3), round(max(text.bbox[1] - rect[1], 0.0), 3),
            round(max(rect[2] - text.bbox[2], 0.0), 3), 0.0)


def _mid(r: tuple[float, float, float, float]) -> tuple[float, float]:
    return (r[0] + r[2]) / 2, (r[1] + r[3]) / 2


def _bold_share(cell_glyphs: dict[tuple[int, int], list[Glyph]], fonts: dict[str, FontRef],
                matches: dict[str, FontMatch], rows: list[int]) -> float:
    gs = [g for (r, _), v in cell_glyphs.items() if r in rows for g in v]
    if not gs:
        return 0.0

    def bold(g: Glyph) -> bool:
        m = matches.get(g.font)
        if m is not None:
            return bool(m.bold)
        f = fonts.get(g.font)
        return bool(f and f.bold)

    return sum(1 for g in gs if bold(g)) / len(gs)


# --------------------------------------------------------------------------- ruled tables


def ruled_grid(comp: _Comp, gs: list[Glyph]) -> Grid | None:
    vs = [s for s in comp.segs if s.orient == "v"]
    hs = [s for s in comp.segs if s.orient == "h"]
    xs = _cluster([s.pos for s in vs] + [v for f in comp.fills for v in (f.x0, f.x1)])
    ys = _cluster([s.pos for s in hs] + [v for f in comp.fills for v in (f.y0, f.y1)])
    if len(xs) < 3 or len(ys) < 3 or len(xs) > 60 or len(ys) > 200:
        return None
    if not vs and not comp.fills:
        return None  # rows only: the whitespace detector decides the columns
    grid = Grid(xs=xs, ys=ys)
    box = (xs[0], ys[0], xs[-1], ys[-1])
    inside = _glyphs_in(gs, box)
    if len(inside) < 3:
        return None
    sep_v: dict[tuple[int, int], bool] = {}
    sep_h: dict[tuple[int, int], bool] = {}
    crossing = 0
    for r in range(grid.rows):
        for c in range(grid.cols):
            cx, cy = (xs[c] + xs[c + 1]) / 2, (ys[r] + ys[r + 1]) / 2
            if c + 1 < grid.cols:
                cov, _ = _coverage(comp.segs, "v", xs[c + 1], ys[r], ys[r + 1])
                fa = _fill_at(comp.fills, cx, cy)
                fb = _fill_at(comp.fills, (xs[c + 1] + xs[c + 2]) / 2, cy)
                sep = cov >= COVER or (fa is not fb and (fa is not None or fb is not None))
                if sep and any(g.bbox[0] < xs[c + 1] - 1.0 < xs[c + 1] + 1.0 < g.bbox[2]
                               and ys[r] <= _center(g)[1] <= ys[r + 1] for g in inside):
                    crossing += 1
                sep_v[(r, c)] = sep
            if r + 1 < grid.rows:
                cov, _ = _coverage(comp.segs, "h", ys[r + 1], xs[c], xs[c + 1])
                fa = _fill_at(comp.fills, cx, cy)
                fb = _fill_at(comp.fills, cx, (ys[r + 1] + ys[r + 2]) / 2)
                sep_h[(r, c)] = cov >= COVER or (fa is not fb and (fa is not None or
                                                                   fb is not None))
    if not merge_regions(grid, sep_h, sep_v):
        return None
    if len(grid.spans) < 4:
        return None
    if crossing:
        grid.confidence -= 0.3
        grid.notes.append(f"text_crosses_rules:{crossing}")
    return grid


# --------------------------------------------------------------------------- whitespace tables


@dataclass
class _WSeg:
    x0: float
    x1: float
    base: float
    size: float
    top: float
    bot: float
    text: str
    ids: list[str]


def text_segments(page: Page, fonts: dict[str, FontRef], matches: dict[str, FontMatch],
                  skip: set[str]) -> list[_WSeg]:
    """Upright text lines split at wide gaps (the same rule as text grouping)."""
    vis = [(i, g) for i, g in enumerate(page.glyphs)
           if g.render_mode not in (3, 7) and g.color.alpha > 0.01 and g.id not in skip
           and abs(g.rotation) < 0.5]
    if not vis:
        return []
    ctx = tg.Ctx(page=page, fonts=fonts, matches=matches)
    lines = tg._form_lines(tg._to_local(ctx, vis))
    for ln in lines:
        tg._words(ctx, ln)
    lines = [ln for ln in lines if ln.words]
    thr = max(tg.SPLIT_MIN, 3.0 * tg._gap_stats(lines))
    out = []
    for ln in lines:
        for sub in tg._split_line(ln, thr):
            gl = [x for w in sub.words for x in w.glyphs]
            out.append(_WSeg(x0=sub.a0, x1=sub.a1, base=sub.base, size=sub.size,
                             top=min(x.top for x in gl), bot=max(x.bot for x in gl),
                             text=sub.text, ids=[x.g.id for x in sub.glyphs if not x.space]))
    return out


def _rows(segs: list[_WSeg]) -> list[list[_WSeg]]:
    out: list[list[_WSeg]] = []
    for s in sorted(segs, key=lambda s: (s.base, s.x0)):
        if out and abs(out[-1][0].base - s.base) <= 0.3 * max(s.size, 1.0):
            out[-1].append(s)
        else:
            out.append([s])
    for r in out:
        r.sort(key=lambda s: s.x0)
    return out


def _gutters(rows: list[list[_WSeg]]) -> list[tuple[float, float]] | None:
    segs = [s for r in rows for s in r]
    size = statistics.median([s.size for s in segs])
    spans = sorted((s.x0, s.x1) for s in segs)
    cols: list[list[float]] = []
    for a, b in spans:
        if cols and a <= cols[-1][1] + max(0.8 * size, 4.0):
            cols[-1][1] = max(cols[-1][1], b)
        else:
            cols.append([a, b])
    return [(a, b) for a, b in cols] if len(cols) >= 2 else None


def whitespace_grids(segs: list[_WSeg], hrules: list[Seg]) -> list[tuple[Grid, list[_WSeg]]]:
    rows = _rows(segs)
    out: list[tuple[Grid, list[_WSeg]]] = []
    i = 0
    while i < len(rows):
        if len(rows[i]) < 2:
            i += 1
            continue
        run = [rows[i]]
        j = i + 1
        while j < len(rows):
            prev, cur = run[-1], rows[j]
            size = max(prev[0].size, 1.0)
            pitch = cur[0].base - prev[0].base
            if not (0.8 * size <= pitch <= 2.6 * size) or not (
                    0.75 <= cur[0].size / size <= 1.33):
                break
            if len(cur) < 2 and not (j + 1 < len(rows) and len(rows[j + 1]) >= 2):
                break
            run.append(cur)
            j += 1
        grid = _ws_grid(run, hrules)
        if grid is not None:
            out.append((grid, [s for r in run for s in r]))
            i = j
        else:
            i += 1
    return out


def _ws_grid(run: list[list[_WSeg]], hrules: list[Seg]) -> Grid | None:
    if len(run) < 3:
        return None
    cols = _gutters(run)
    if cols is None or len(cols) > 12:
        return None
    n_c = len(cols)
    cells: dict[tuple[int, int], list[_WSeg]] = defaultdict(list)
    spans: dict[tuple[int, int], int] = {}
    for r, row in enumerate(run):
        for s in row:
            hit = [k for k, (a, b) in enumerate(cols) if s.x0 < b + 0.5 and s.x1 > a - 0.5]
            if not hit:
                return None
            if len(hit) > 1:
                if r != 0:
                    return None  # only a header may span columns
                spans[(r, hit[0])] = len(hit)
            cells[(r, hit[0])].append(s)
    filled = len(cells)
    if filled < 0.6 * len(run) * n_c:
        return None
    for k in range(n_c):
        if sum(1 for r in range(len(run)) if (r, k) in cells) < 0.5 * len(run):
            return None
    texts = [s.text for v in cells.values() for s in v]
    if statistics.mean(len(t) for t in texts) > 20 or max(len(t) for t in texts) > 40:
        return None
    col0 = [s.text for r in range(len(run)) for s in cells.get((r, 0), [])]
    if col0 and all(_LABEL_RE.match(t.split(" ")[0]) for t in col0):
        return None
    numeric = any(
        sum(1 for r in range(len(run)) for s in cells.get((r, k), []) if _NUM_RE.match(s.text))
        >= 0.5 * len(run) for k in range(n_c))
    near_rules = [h for h in hrules if h.b - h.a > 0.8 * (cols[-1][1] - cols[0][0])]
    if not numeric and len(near_rules) < 2:
        return None
    size = statistics.median([s.size for row in run for s in row])
    pad = 0.4 * size
    xs = [cols[0][0] - pad]
    xs += [(cols[k][1] + cols[k + 1][0]) / 2 for k in range(n_c - 1)]
    xs.append(cols[-1][1] + pad)
    tops = [min(s.top for s in row) for row in run]
    bots = [max(s.bot for s in row) for row in run]
    ys = [tops[0] - 0.3 * size]
    ys += [(bots[r] + tops[r + 1]) / 2 for r in range(len(run) - 1)]
    ys.append(bots[-1] + 0.3 * size)
    # rules on row boundaries (above the first row, below the last) join the table
    for h in near_rules:
        if h.a <= xs[0] + 2 * size and h.b >= xs[-1] - 2 * size:
            k = min(range(len(ys)), key=lambda q: abs(ys[q] - h.pos))
            lim = 0.6 * size if 0 < k < len(ys) - 1 else 1.2 * size
            if abs(ys[k] - h.pos) <= lim:
                ys[k] = h.pos
    xs = [round(v, 3) for v in xs]
    ys = [round(v, 3) for v in ys]
    if any(b <= a for a, b in itertools.pairwise(ys)):
        return None
    grid = Grid(xs=xs, ys=ys, kind="whitespace", confidence=0.85 if numeric else 0.7)
    sep_v = {(r, c): not (r == 0 and any(c0 <= c < c0 + n - 1 for (rr, c0), n in spans.items()
                                         if rr == 0))
             for r in range(len(run)) for c in range(n_c - 1)}
    sep_h = {(r, c): True for r in range(len(run) - 1) for c in range(n_c)}
    if not merge_regions(grid, sep_h, sep_v):
        return None
    return grid


# --------------------------------------------------------------------------- page driver


def analyze_page(page: Page, fonts: dict[str, FontRef], matches: dict[str, FontMatch]) -> set[
        str]:
    """Detect tables, add them to ``page.elements`` (consumed paths removed). Returns the
    glyph ids now owned by table cells (text grouping skips them)."""
    used_glyphs: set[str] = set()
    segs, fills, nseg = collect(page)
    tables: list[Table] = []
    consumed: dict[str, int] = defaultdict(int)
    for comp in components(segs, fills):
        gs = _visible(page, used_glyphs)
        grid = ruled_grid(comp, gs)
        if grid is None or grid.confidence < MIN_CONFIDENCE:
            continue
        tid = f"p{page.index}_t{len(tables) + 1}"
        zs = [e.z for e in page.elements if getattr(e, "id", None) in
              {s.sid for s in comp.segs} | {f.sid for f in comp.fills}]
        res = build_table(page, fonts, matches, grid, gs, comp.segs, comp.fills, tid,
                          min(zs) if zs else 0)
        if res is None:
            continue
        t, ids = res
        box = t.bbox
        for s in comp.segs:
            if box[0] - 2 <= _box(s)[0] and _box(s)[2] <= box[2] + 2 and \
                    box[1] - 2 <= _box(s)[1] and _box(s)[3] <= box[3] + 2:
                consumed[s.sid] += 1
        for f in comp.fills:
            consumed[f.sid] += 1
        tables.append(t)
        used_glyphs |= ids
    # whitespace tables on the remaining text
    hrules = [s for s in segs if s.orient == "h" and consumed.get(s.sid, 0) < nseg.get(s.sid, 1)]
    wsegs = text_segments(page, fonts, matches, used_glyphs)
    for grid, members in whitespace_grids(wsegs, hrules):
        if grid.confidence < MIN_CONFIDENCE:
            continue
        gs = [g for g in _visible(page, used_glyphs) if g.id in {i for m in members
                                                                  for i in m.ids}]
        tid = f"p{page.index}_t{len(tables) + 1}"
        box = (grid.xs[0], grid.ys[0], grid.xs[-1], grid.ys[-1])
        rules = [h for h in hrules if box[1] - 1 <= h.pos <= box[3] + 1
                 and h.a >= box[0] - 20 and h.b <= box[2] + 20]
        res = build_table(page, fonts, matches, grid, gs, rules, [], tid,
                          min((g.z for g in gs), default=0))
        if res is None:
            continue
        t, ids = res
        for h in rules:
            consumed[h.sid] += 1
        tables.append(t)
        used_glyphs |= ids
    if not tables:
        return set()
    gone = {sid for sid, n in consumed.items() if n >= nseg.get(sid, 1)}
    for t in tables:
        t.fallback_ids = sorted(gone & _ids_in(page, t.bbox))
    page.elements = [e for e in page.elements if getattr(e, "id", None) not in gone]
    page.elements.extend(tables)
    LOG.info("page %d: %d table(s): %s", page.index, len(tables),
             ", ".join(f"{t.id} {len(t.row_heights)}x{len(t.col_widths)}" for t in tables))
    return used_glyphs


def _ids_in(page: Page, box: tuple[float, float, float, float]) -> set[str]:
    out = set()
    for e in page.elements:
        b: Any = getattr(e, "bbox", None)
        if b is not None and b[0] >= box[0] - 3 and b[2] <= box[2] + 3 and \
                b[1] >= box[1] - 3 and b[3] <= box[3] + 3:
            out.add(e.id)
    return out
