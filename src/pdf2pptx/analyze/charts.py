"""ANALYZE (charts, M9): vector chart drawings -> native ``Chart`` elements.

Runs after ``shapes`` (it reads the fitted presets) on born-digital pages. Rules only;
the Max tier's vision model cross-checks and handles raster charts
(``llm/charts_vlm.py``).

* **Value axis.** Three or more numeric text blocks in a column (column / line charts)
  or a row (bar charts) whose values are a linear function of their position. Each
  label is snapped to the gridline or tick mark next to it; the least-squares fit
  maps positions to values.
* **Columns / bars.** Filled rectangles inside the plot standing on the base line
  (the value-axis minimum, or zero). Fill colour = series; the category is the slot
  of the plot the bar's centre falls in (PowerPoint spreads categories evenly).
  Different colours over the same x range = stacked.
* **Lines.** Stroked open paths inside the plot; one value per vertex, markers
  (small fitted shapes on the vertices) give the marker style.
* **Pies.** Two or more ``pie`` presets with one centre: the slice angles give the
  shares; data labels inside the slices give the exact values when they agree.
* **Categories / legend / data labels.** Category labels sit under the base line (left
  of it for bars). Legend swatches are small shapes in a series colour with words to
  their right (words are taken glyph by glyph, so one text block may hold the whole
  legend). Numeric labels next to bar ends or inside slices are data labels.

Every element a chart consumes is removed from the page and kept in
``Chart.fallback``; a chart that fails its render check puts them back
(``restore``). A text block that is only partly chart text keeps the chart from
being built (nothing changes).
"""

from __future__ import annotations

import itertools
import math
import re
import statistics
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

from .. import log
from ..ir import Chart, ChartSeries, Color, Glyph, Group, Image, Page, Shape, TextBlock

LOG = log.get(__name__)

VERSION = "9.0"
MIN_LABELS = 3
FIT_TOL = 0.02  # max residual of the axis fit, as a share of the axis range
SNAP = 4.0  # pt: a label snaps to a gridline / tick this close
SWATCH_MAX = 16.0  # pt: legend swatch size
MIN_CONFIDENCE = 0.6
NUM_RE = re.compile(r"^[\s(]*([-+−–]?)\s*[€$£¥]?\s*([0-9][0-9.,\s']*)\s*([%‰kKmM]?)\s*[€$£)]?$")

BBox = tuple[float, float, float, float]


# --------------------------------------------------------------------------- helpers


def block_text(b: TextBlock) -> str:
    return "\n".join("".join(r.text for r in p.runs) for p in b.paragraphs).strip()


def parse_number(text: str, thousands: str | None = None) -> float | None:
    """A number as written on a chart: ``1,250`` ``1.250`` ``0,5`` ``42%`` ``-3`` ``€ 12``.
    ``thousands``: the separator to drop ("," or "."); None guesses per label."""
    m = NUM_RE.match(text.strip())
    if not m:
        return None
    sign, body, unit = m.groups()
    body = body.strip().rstrip(".,")
    # spaces / apostrophes only as thousands groups ("1 250", "1'250"), not "2024 2025"
    if (" " in body or "'" in body) and not re.fullmatch(r"\d{1,3}([ ']\d{3})+([.,]\d+)?",
                                                         body):
        return None
    body = body.replace(" ", "").replace("'", "")
    if not body or not body[0].isdigit():
        return None
    if thousands is None:
        if "," in body and "." in body:
            thousands = "," if body.rfind(".") > body.rfind(",") else "."
        elif re.fullmatch(r"\d{1,3}([.,]\d{3})+", body):
            thousands = body[len(body.split(",")[0].split(".")[0])]
        else:
            thousands = ""
    if thousands:
        body = body.replace(thousands, "")
    body = body.replace(",", ".")
    if body.count(".") > 1:
        return None
    try:
        v = float(body)
    except ValueError:
        return None
    if unit in ("k", "K"):
        v *= 1e3
    elif unit in ("m", "M"):
        v *= 1e6
    return -v if sign in ("-", "−", "–") else v


def _c(b: BBox) -> tuple[float, float]:
    return (b[0] + b[2]) / 2, (b[1] + b[3]) / 2


def _inside(b: BBox, box: BBox, tol: float = 1.5) -> bool:
    return b[0] >= box[0] - tol and b[1] >= box[1] - tol and b[2] <= box[2] + tol and \
        b[3] <= box[3] + tol


def _union(boxes: list[BBox]) -> BBox:
    return (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes),
            max(b[3] for b in boxes))


def _rgb(c: Color | None) -> str | None:
    return c.rgb.upper() if c is not None and c.rgb and c.alpha > 0.5 else None


def _fill_rgb(s: Shape) -> str | None:
    return _rgb(s.fill.color) if s.fill.kind == "solid" else None


def _stroke_rgb(s: Shape) -> str | None:
    return _rgb(s.stroke.color) if s.stroke is not None else None


def _is_hline(s: Shape) -> bool:
    return s.preset == "line" and abs(s.bbox[3] - s.bbox[1]) < 0.6


def _is_vline(s: Shape) -> bool:
    return s.preset == "line" and abs(s.bbox[2] - s.bbox[0]) < 0.6


# --------------------------------------------------------------------------- axis


@dataclass
class Axis:
    vertical: bool  # labels in a column: values along y (column / line charts)
    blocks: list[TextBlock]
    values: list[float]
    positions: list[float]  # snapped tick positions (y for vertical axes, x otherwise)
    a: float  # value = a + b * pos
    b: float
    residual: float
    percent: bool = False
    number_format: str = "General"

    def value(self, pos: float) -> float:
        return self.a + self.b * pos

    def pos(self, value: float) -> float:
        return (value - self.a) / self.b

    @property
    def vmin(self) -> float:
        return min(self.values)

    @property
    def vmax(self) -> float:
        return max(self.values)

    @property
    def major(self) -> float:
        v = sorted(self.values)
        return statistics.median(b - a for a, b in itertools.pairwise(v))

    @property
    def label_box(self) -> BBox:
        return _union([b.bbox for b in self.blocks])


def _linfit(xs: list[float], ys: list[float]) -> tuple[float, float, float]:
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    if sxx <= 1e-9:
        return my, 0.0, float("inf")
    b = sum((x - mx) * (y - my) for x, y in zip(xs, ys, strict=True)) / sxx
    a = my - b * mx
    res = max(abs(a + b * x - y) for x, y in zip(xs, ys, strict=True))
    return a, b, res


def _snap(pos: float, lines: list[float]) -> float:
    best = min(lines, key=lambda v: abs(v - pos), default=None)
    return best if best is not None and abs(best - pos) <= SNAP else pos


def _number_format(texts: list[str], percent: bool) -> str:
    dec = 0
    for t in texts:
        m = re.search(r"[.,](\d+)\s*%?\s*$", t.strip())
        if m and not re.fullmatch(r"[^\d]*\d{1,3}([.,]\d{3})+[^\d]*", t.strip()):
            dec = max(dec, len(m.group(1)))
    base = "0" + ("." + "0" * dec if dec else "")
    if any(re.search(r"\d[.,]\d{3}", t) for t in texts) and not dec:
        base = "#,##0"
    return base + ("%" if percent else "")


def find_axes(blocks: list[TextBlock], hlines: list[float], vlines: list[float]) -> list[Axis]:
    """Columns / rows of numeric labels whose values are linear in their position."""
    nums: list[tuple[TextBlock, float, str]] = []
    for b in blocks:
        t = block_text(b)
        if "\n" in t:
            continue
        v = parse_number(t)
        if v is not None:
            nums.append((b, v, t))
    axes: list[Axis] = []
    for vertical in (True, False):
        # group: same right edge / centre (columns) or same centre y (rows)
        groups: list[list[tuple[TextBlock, float, str]]] = []
        for item in sorted(nums, key=lambda it: (it[0].bbox[2] if vertical else _c(
                it[0].bbox)[1])):
            b = item[0]
            key_r, key_c = (b.bbox[2], _c(b.bbox)[0]) if vertical else (_c(b.bbox)[1],
                                                                         _c(b.bbox)[1])
            for g in groups:
                g0 = g[0][0].bbox
                gr, gc = (g0[2], _c(g0)[0]) if vertical else (_c(g0)[1], _c(g0)[1])
                if abs(key_r - gr) < 2.5 or abs(key_c - gc) < 2.5:
                    g.append(item)
                    break
            else:
                groups.append([item])
        for g in groups:
            if len(g) < MIN_LABELS:
                continue
            g.sort(key=lambda it: _c(it[0].bbox)[1 if vertical else 0])
            pos = [_snap(_c(it[0].bbox)[1 if vertical else 0], hlines if vertical else vlines)
                   for it in g]
            # evenly spaced, distinct positions
            gaps = [q - p for p, q in itertools.pairwise(pos)]
            if min(gaps) < 4 or max(gaps) > 1.25 * min(gaps) + 2:
                continue
            texts = [it[2] for it in g]
            # thousands separators: try both conventions, keep the better fit
            best: tuple[float, float, float, list[float]] | None = None
            for th in (None, ",", "."):
                vals = [parse_number(t, th) for t in texts]
                if any(v is None for v in vals):
                    continue
                fv = [float(v) for v in vals if v is not None]
                if len(set(fv)) < len(fv):
                    continue
                a, bb, res = _linfit(pos, fv)
                if best is None or res < best[2]:
                    best = (a, bb, res, fv)
            if best is None:
                continue
            a, bb, res, vals = best
            rng = max(vals) - min(vals)
            if rng <= 0 or res > FIT_TOL * rng or bb == 0:
                continue
            percent = all(t.strip().endswith("%") for t in texts)
            axes.append(Axis(vertical, [it[0] for it in g], vals, pos, a, bb, res / rng,
                             percent, _number_format(texts, percent)))
    return axes


# --------------------------------------------------------------------------- words


def words(page: Page, blocks: list[TextBlock]) -> list[tuple[str, BBox, str]]:
    """The words of ``blocks`` as (text, bbox, block id), from the glyphs (a legend often
    is one text block)."""
    gmap = {g.id: g for g in page.glyphs}
    out: list[tuple[str, BBox, str]] = []
    for b in blocks:
        for p in b.paragraphs:
            gl: list[Glyph] = [gmap[i] for r in p.runs for i in r.glyph_ids if i in gmap]
            if not gl:
                if "".join(r.text for r in p.runs).strip():
                    out.append(("".join(r.text for r in p.runs).strip(), p.bbox, b.id))
                continue
            gl.sort(key=lambda g: (round(g.baseline), g.bbox[0]))
            cur: list[Glyph] = []
            for g in [*gl, None]:  # type: ignore[list-item]
                split = g is None or g.text.isspace() or (cur and (
                    abs(g.baseline - cur[-1].baseline) > 2
                    or g.bbox[0] - cur[-1].bbox[2] > 0.25 * g.size))
                if split and cur:
                    out.append(("".join(c.text for c in cur),
                                _union([c.bbox for c in cur]), b.id))
                    cur = []
                if g is not None and not g.text.isspace():
                    cur.append(g)
    return out


# --------------------------------------------------------------------------- legend


@dataclass
class LegendEntry:
    color: str
    swatch: Shape
    extra: list[Shape]  # a line chart's legend line + marker
    text: str
    bbox: BBox
    word_blocks: set[str]


def find_legend(shapes: list[Shape], wds: list[tuple[str, BBox, str]],
                colors: set[str], exclude: set[str]) -> list[LegendEntry]:
    sw = []
    for s in shapes:
        if s.id in exclude or s.is_background:
            continue
        w, h = s.bbox[2] - s.bbox[0], s.bbox[3] - s.bbox[1]
        col = _fill_rgb(s) or (_stroke_rgb(s) if _is_hline(s) else None)
        if col in colors and w <= 3 * SWATCH_MAX and h <= SWATCH_MAX and (
                (s.preset != "line" and w <= SWATCH_MAX) or (_is_hline(s) and w <= 3 * SWATCH_MAX)):
            sw.append((s, col))
    # merge a legend line with the marker drawn on it
    entries: list[tuple[list[Shape], str, BBox]] = []
    for s, col in sorted(sw, key=lambda t: (t[0].bbox[1], t[0].bbox[0])):
        for ent in entries:
            if ent[1] == col and _inside(_c(s.bbox) * 2, ent[2], 1.0):  # type: ignore[arg-type]
                ent[0].append(s)
                break
            if ent[1] == col and _inside(_c(ent[2]) * 2, s.bbox, 1.0):  # type: ignore[arg-type]
                ent[0].append(s)
                break
        else:
            entries.append(([s], col, s.bbox))
    out: list[LegendEntry] = []
    for parts, col, _ in entries:
        box = _union([p.bbox for p in parts])
        cy = _c(box)[1]
        # next swatch on the same row bounds the text
        nxt = min((_union([p.bbox for p in o[0]])[0] for o in entries
                   if o[0] is not parts and abs(_c(_union([p.bbox for p in o[0]]))[1] - cy) < 4
                   and _union([p.bbox for p in o[0]])[0] > box[2]), default=1e9)
        ws = sorted((w for w in wds if abs(_c(w[1])[1] - cy) < 5 and w[1][0] >= box[2] - 1
                     and w[1][2] <= nxt + 1), key=lambda w: w[1][0])
        # contiguous words right of the swatch
        picked: list[tuple[str, BBox, str]] = []
        last = box[2]
        for w in ws:
            if w[1][0] - last > 30:
                break
            picked.append(w)
            last = w[1][2]
        if not picked:
            continue
        main = parts[0] if len(parts) == 1 else max(parts, key=lambda p: (p.preset == "line"))
        out.append(LegendEntry(col, main, [p for p in parts if p is not main],
                               " ".join(w[0] for w in picked),
                               _union([box] + [w[1] for w in picked]), {w[2] for w in picked}))
    return out


# --------------------------------------------------------------------------- results


@dataclass
class Found:
    chart: Chart
    consumed: set[str] = field(default_factory=set)  # element ids
    words_used: list[tuple[str, BBox, str]] = field(default_factory=list)


def _marker(s: Shape) -> str:
    return {"diamond": "diamond", "rect": "square", "triangle": "triangle", "ellipse": "circle",
            "plus": "plus", "mathMultiply": "x"}.get(s.preset or "", "auto")


def _font_style(blocks: list[TextBlock]) -> dict[str, float | str | None]:
    for b in blocks:
        for p in b.paragraphs:
            for r in p.runs:
                if r.text.strip():
                    return {"font_size": round(r.size, 1), "font_family": r.font,
                            "font_color": _rgb(r.color)}
    return {}


def _category_labels(blocks: list[TextBlock], box: BBox, vertical: bool,
                     used: set[str]) -> list[TextBlock]:
    """Text right under the plot (column / line) or left of it (bar)."""
    out = []
    for b in blocks:
        if b.id in used:
            continue
        cx, cy = _c(b.bbox)
        below = (vertical and box[0] - 20 <= cx <= box[2] + 20
                 and box[3] - 2 <= b.bbox[1] <= box[3] + 30)
        left = (not vertical and box[1] - 10 <= cy <= box[3] + 10
                and box[0] - 200 <= b.bbox[0] and b.bbox[2] <= box[0] + 2)
        if below or left:
            out.append(b)
    return out


def _bars(shapes: list[Shape], axis: Axis, plot: BBox, used: set[str]) -> list[Shape]:
    base = axis.pos(0.0) if axis.vmin <= 0 <= axis.vmax else axis.pos(axis.vmin)
    out = []
    for s in shapes:
        if s.id in used or s.preset != "rect" or _fill_rgb(s) is None or s.is_background:
            continue
        x0, y0, x1, y1 = s.bbox
        if not _inside(s.bbox, plot, 1.5) or (x1 - x0) < 1.5 or (y1 - y0) < 0.5:
            continue
        ends = (y0, y1) if axis.vertical else (x0, x1)
        if min(abs(ends[0] - base), abs(ends[1] - base)) < 1.2:
            out.append(s)
    return out


def _vertices(s: Shape) -> list[tuple[float, float]]:
    pts: list[tuple[float, float]] = []
    for seg in s.path or []:
        if (seg.op in ("M", "L") and seg.pts) or (seg.op == "C" and seg.pts):
            pts.append(seg.pts[-1])
    return pts


def _plot_box(axis: Axis, shapes: list[Shape], page: Page) -> tuple[BBox, list[Shape]]:
    """The plot rectangle from the gridlines / axis lines at the tick positions."""
    lo, hi = min(axis.positions), max(axis.positions)
    lab = axis.label_box
    if axis.vertical:
        cand = [s for s in shapes if _is_hline(s) and lab[2] - 2 <= s.bbox[0] <= lab[2] + 40
                and lo - 1.5 <= s.bbox[1] <= hi + 1.5 and s.bbox[2] - s.bbox[0] > 20]
        if cand:
            x0 = statistics.median(s.bbox[0] for s in cand)
            x1 = max(s.bbox[2] for s in cand)
        else:
            x0, x1 = lab[2] + 4, page.width
        return (x0, lo, x1, hi), cand
    cand = [s for s in shapes if _is_vline(s) and lab[1] - 40 <= s.bbox[3] <= lab[1] + 2
            and lo - 1.5 <= s.bbox[0] <= hi + 1.5 and s.bbox[3] - s.bbox[1] > 20]
    if cand:
        y0 = min(s.bbox[1] for s in cand)
        y1 = statistics.median(s.bbox[3] for s in cand)
    else:
        y0, y1 = 0.0, lab[1] - 4
    return (lo, y0, hi, y1), cand


def _axis_chart(page: Page, axis: Axis, shapes: list[Shape], blocks: list[TextBlock],
                wds: list[tuple[str, BBox, str]], cid: str) -> Found | None:
    plot, grid = _plot_box(axis, shapes, page)
    used: set[str] = {b.id for b in axis.blocks}
    bars = _bars(shapes, axis, plot, used)
    lines = [s for s in shapes if s.preset is None and s.path and s.stroke is not None
             and s.fill.kind == "none" and _inside(s.bbox, plot, 2.0)
             and len(_vertices(s)) >= 2] if axis.vertical else []
    if not bars and not lines:
        return None
    cats = _category_labels(blocks, plot, axis.vertical, used)
    if cats and any("\n" in block_text(c) for c in cats):
        cats = [c for c in cats if "\n" not in block_text(c)]
    # category slots
    if axis.vertical:
        cats.sort(key=lambda b: _c(b.bbox)[0])
        span = (plot[0], plot[2])
    else:
        cats.sort(key=lambda b: -_c(b.bbox)[1])  # first category at the bottom
        span = (plot[3], plot[1])
    n = len(cats)
    marks: list[float] = []
    if bars:
        marks = [_c(s.bbox)[0 if axis.vertical else 1] for s in bars]
    for ln in lines:
        marks += [p[0] for p in _vertices(ln)]
    if n == 0:  # no labels: count slots from the bar / vertex clusters
        cl = sorted(marks)
        groups = [[cl[0]]]
        for v in cl[1:]:
            if v - groups[-1][-1] > 0.25 * abs(span[1] - span[0]) / max(len(cl), 1) + 4:
                groups.append([v])
            else:
                groups[-1].append(v)
        n = len(groups)
    if n == 0:
        return None
    width = (span[1] - span[0]) / n

    def slot(v: float) -> int:
        return max(0, min(n - 1, int((v - span[0]) / width)))

    # labels must sit on their slots (else this is not our category axis)
    for i, c in enumerate(cats):
        cc = _c(c.bbox)[0 if axis.vertical else 1]
        if slot(cc) != i:
            LOG.info("page %d: chart candidate dropped (category %r off its slot)",
                     page.index, block_text(c))
            return None
    categories = [block_text(c) for c in cats] or [str(i + 1) for i in range(n)]
    series: list[ChartSeries] = []
    consumed = {b.id for b in axis.blocks} | {b.id for b in cats}
    chart_type = ""
    style: dict[str, float | str | None] = {}
    if bars:
        by_col: dict[str, list[Shape]] = defaultdict(list)
        for s in bars:
            by_col[_fill_rgb(s) or ""].append(s)
        # stacked: two colours over the same slot range
        stacked = False
        for a in bars:
            for b in bars:
                if _fill_rgb(a) != _fill_rgb(b):
                    ia = (a.bbox[0], a.bbox[2]) if axis.vertical else (a.bbox[1], a.bbox[3])
                    ib = (b.bbox[0], b.bbox[2]) if axis.vertical else (b.bbox[1], b.bbox[3])
                    ov = min(ia[1], ib[1]) - max(ia[0], ib[0])
                    if ov > 0.5 * (ia[1] - ia[0]):
                        stacked = True
        cols = sorted(by_col, key=lambda k: statistics.median(
            (_c(s.bbox)[0] - (span[0] + (slot(_c(s.bbox)[0]) + 0.5) * width)) if axis.vertical
            else -(_c(s.bbox)[1] - (span[0] + (slot(_c(s.bbox)[1]) + 0.5) * width))
            for s in by_col[k]))
        if stacked:  # order stacked series from the base outwards
            cols = sorted(by_col, key=lambda k: statistics.median(
                abs((s.bbox[3] if axis.vertical else s.bbox[0]) - axis.pos(
                    max(axis.vmin, 0.0))) for s in by_col[k]))
        for k in cols:
            vals: list[float | None] = [None] * n
            for s in by_col[k]:
                i = slot(_c(s.bbox)[0 if axis.vertical else 1])
                if axis.vertical:
                    top, bot = axis.value(s.bbox[1]), axis.value(s.bbox[3])
                else:
                    top, bot = axis.value(s.bbox[2]), axis.value(s.bbox[0])
                v = (top - bot) if stacked else (top if abs(bot - max(axis.vmin, 0.0)) < abs(
                    top - max(axis.vmin, 0.0)) else bot)
                vals[i] = (vals[i] or 0.0) + v if stacked else v
                consumed.add(s.id)
            series.append(ChartSeries(name="", values=vals, color=Color(rgb=k)))
        # one bar per category, each in its own colour: one series, varied colours
        if not stacked and len(series) > 1 and all(
                sum(v is not None for v in x.values) == 1 for x in series) and all(
                sum(x.values[i] is not None for x in series) <= 1 for i in range(n)):
            vals = [next((x.values[i] for x in series if x.values[i] is not None), None)
                    for i in range(n)]
            pcs = [next((x.color for x in series if x.values[i] is not None), None)
                   for i in range(n)]
            series = [ChartSeries(name="", values=vals, point_colors=pcs)]
            style["vary_colors"] = 1.0
        kind = "column" if axis.vertical else "bar"
        chart_type = f"{kind}_{'stacked' if stacked else 'clustered'}"
        bw = statistics.median((s.bbox[2] - s.bbox[0]) if axis.vertical else (s.bbox[3] - s.bbox[1])
                               for s in bars)
        per = 1 if stacked else len(series)
        style["gap_width"] = round(max(0.0, min(500.0, (abs(width) - per * bw) / bw * 100)))
        style["overlap"] = 100.0 if stacked else 0.0
    else:
        markers_used: set[str] = set()
        smalls = [s for s in shapes if s.id not in used and s.preset not in (None, "line")
                  and _fill_rgb(s) and (s.bbox[2] - s.bbox[0]) <= SWATCH_MAX
                  and _inside(s.bbox, plot, 8.0)]
        for ln in sorted(lines, key=lambda s: _c(s.bbox)[1]):
            col = _stroke_rgb(ln) or ""
            vals = [None] * n
            for x, y in _vertices(ln):
                vals[slot(x)] = axis.value(y)
            mk = [s for s in smalls if any(
                abs(_c(s.bbox)[0] - x) < 1.5 and abs(_c(s.bbox)[1] - y) < 1.5
                for x, y in _vertices(ln))]
            for s in mk:
                markers_used.add(s.id)
            consumed.add(ln.id)
            series.append(ChartSeries(
                name="", values=vals, color=Color(rgb=col) if col else None,
                marker=_marker(mk[0]) if mk else "none",
                line_width=ln.stroke.width if ln.stroke else None))
        consumed |= markers_used
        chart_type = "line_markers" if markers_used else "line"
    # data labels: numbers near bar ends that are not axis labels
    nums_left = [b for b in blocks if b.id not in consumed and parse_number(block_text(b))
                 is not None and _inside(b.bbox, (plot[0] - 5, plot[1] - 25, plot[2] + 40,
                                                  plot[3] + 5), 0)]
    labels = []
    if bars:
        for b in nums_left:
            cx, cy = _c(b.bbox)
            hit = any(abs(cx - _c(s.bbox)[0]) < (s.bbox[2] - s.bbox[0]) / 2 + 2
                      and (abs(b.bbox[3] - s.bbox[1]) < 6 or s.bbox[1] <= cy <= s.bbox[3])
                      if axis.vertical else
                      abs(cy - _c(s.bbox)[1]) < (s.bbox[3] - s.bbox[1]) / 2 + 2
                      and (abs(b.bbox[0] - s.bbox[2]) < 8 or s.bbox[0] <= cx <= s.bbox[2])
                      for s in bars)
            if hit:
                labels.append(b)
    # grid / axis lines / ticks
    tick_lines = [s for s in shapes if s.id not in consumed and s.preset == "line" and (
        _is_hline(s) or _is_vline(s)) and _inside(s.bbox, (plot[0] - 9, plot[1] - 9,
                                                             plot[2] + 9, plot[3] + 9), 0.5)]
    consumed |= {s.id for s in tick_lines}
    style["gridlines"] = "major" if grid else "none"
    style |= _font_style(axis.blocks)
    cats_style = _font_style(cats)
    if cats_style.get("font_size"):
        style["cat_font_size"] = cats_style["font_size"]
    colors = {s.color.rgb for s in series if s.color and s.color.rgb} | {
        c.rgb for s in series for c in (s.point_colors or []) if c and c.rgb}
    legend = find_legend([s for s in shapes if s.id not in consumed], wds, colors, consumed)
    leg_words: list[tuple[str, BBox, str]] = []
    for ent in legend:
        for s in series:
            if s.color and s.color.rgb == ent.color and not s.name:
                s.name = ent.text
        consumed.add(ent.swatch.id)
        consumed |= {x.id for x in ent.extra} | ent.word_blocks
        leg_words += [w for w in wds if w[2] in ent.word_blocks and _inside(w[1], ent.bbox, 0.5)]
    step = axis.major / 100.0
    leg_style = _font_style([b for b in blocks if any(b.id in e.word_blocks for e in legend)])
    if leg_style.get("font_size"):
        style["legend_font_size"] = leg_style["font_size"]
    for i, s in enumerate(series):
        s.name = s.name or f"Series {i + 1}"
        s.values = [None if v is None else _snap_value(v, step) for v in s.values]
    data_labels = bool(labels) and len(labels) >= 0.6 * sum(
        1 for s in series for v in s.values if v is not None)
    if data_labels:
        consumed |= {b.id for b in labels}
        # exact values where the labels agree with the geometry
        _apply_bar_labels(series, labels, bars, axis, slot)
    axes = {"value": {"min": axis.vmin, "max": axis.vmax, "major": axis.major,
                      "format": axis.number_format}}
    chart = Chart(id=cid, bbox=plot, chart_type=chart_type, categories=categories,
                  series=series, has_legend=bool(legend),
                  legend_pos=_legend_pos(legend, plot), axes=axes, data_labels=data_labels,
                  plot_bbox=plot, legend_bbox=_union([e.bbox for e in legend]) if legend else None,
                  style=style, source="rule", confidence=_confidence(axis, series, cats, n))
    return Found(chart, consumed, leg_words)


def _snap_value(v: float, step: float) -> float:
    """Geometry carries ~0.01 pt of noise: snap to 1/100 of the major unit."""
    if step <= 0:
        return v
    q = round(v / step) * step
    return round(q, 10) if abs(q - v) <= step / 2 else v


def _apply_bar_labels(series: list[ChartSeries], labels: list[TextBlock], bars: list[Shape],
                      axis: Axis, slot: Any) -> None:
    for b in labels:
        v = parse_number(block_text(b))
        if v is None:
            continue
        cx, cy = _c(b.bbox)
        bar = min(bars, key=lambda s: math.dist(_c(s.bbox), (cx, cy)))
        i = slot(_c(bar.bbox)[0 if axis.vertical else 1])
        for s in series:
            pc = s.point_colors[i] if s.point_colors else None
            col = (pc or s.color).rgb if (pc or s.color) else None  # type: ignore[union-attr]
            if col == _fill_rgb(bar) and s.values[i] is not None:
                cur = float(s.values[i] or 0.0)
                if abs(cur - v) <= 0.03 * max(abs(axis.vmax - axis.vmin), 1e-9):
                    s.values[i] = v


def _legend_pos(legend: list[LegendEntry], plot: BBox) -> str:
    if not legend:
        return "right"
    lb = _union([e.bbox for e in legend])
    cx, _cy = _c(lb)
    if lb[1] >= plot[3] - 2:
        return "bottom"
    if lb[3] <= plot[1] + 2:
        return "top"
    return "left" if cx < plot[0] else "right"


def _confidence(axis: Axis, series: list[ChartSeries], cats: list[TextBlock], n: int) -> float:
    c = 0.95 - 4 * axis.residual
    if not cats:
        c -= 0.15
    missing = sum(1 for s in series for v in s.values if v is None)
    c -= 0.1 * missing / max(1, n * len(series))
    return round(max(0.0, min(1.0, c)), 3)


# --------------------------------------------------------------------------- pie


def _pie_chart(page: Page, pies: list[Shape], blocks: list[TextBlock], shapes: list[Shape],
               wds: list[tuple[str, BBox, str]], cid: str) -> Found | None:
    box = pies[0].bbox
    cx, cy = _c(box)

    def ang(a: float) -> float:  # ooxml 60000ths of a degree, clockwise from 3 o'clock
        return (a / 60000.0) % 360.0

    def rel(d: float) -> float:  # clockwise from 12 o'clock
        return (d - 270.0) % 360.0

    wedges = []
    for s in pies:
        st, en = ang(s.adjust.get("adj1", 0.0)), ang(s.adjust.get("adj2", 0.0))
        sweep = (en - st) % 360.0 or 360.0
        wedges.append((rel(st), sweep, s))
    wedges.sort(key=lambda w: w[0])
    first = wedges[0][0]
    total = sum(w[1] for w in wedges)
    if abs(total - 360.0) > 4.0:
        return None
    shares = [w[1] / total * 100.0 for w in wedges]
    consumed = {w[2].id for w in wedges}
    # data labels inside / next to slices
    rad = (box[2] - box[0]) / 2
    labels: dict[int, tuple[float, TextBlock]] = {}
    for b in blocks:
        v = parse_number(block_text(b))
        if v is None:
            continue
        bx, by = _c(b.bbox)
        if math.dist((bx, by), (cx, cy)) > rad * 1.35:
            continue
        a = rel(math.degrees(math.atan2(by - cy, bx - cx)) % 360.0)
        for i, (st, sw, _) in enumerate(wedges):
            if (a - st) % 360.0 <= sw:
                if i not in labels:
                    labels[i] = (v, b)
                break
    values: list[float | None] = [round(v, 6) for v in shares]
    data_labels = False
    if len(labels) == len(wedges):
        lv = [labels[i][0] for i in range(len(wedges))]
        tot = sum(lv)
        if tot > 0 and all(abs(lv[i] / tot * 100 - shares[i]) < 1.5 for i in range(len(lv))):
            values = list(lv)
            data_labels = True
            consumed |= {labels[i][1].id for i in labels}
    colors = [_fill_rgb(w[2]) for w in wedges]
    legend = find_legend([s for s in shapes if s.id not in consumed], wds,
                         {c for c in colors if c}, consumed)
    names = {e.color: e.text for e in legend}
    cats = [names.get(c or "", f"Slice {i + 1}") for i, c in enumerate(colors)]
    leg_words: list[tuple[str, BBox, str]] = []
    for ent in legend:
        consumed.add(ent.swatch.id)
        consumed |= {x.id for x in ent.extra} | ent.word_blocks
        leg_words += [w for w in wds if w[2] in ent.word_blocks and _inside(w[1], ent.bbox, 0.5)]
    series = [ChartSeries(name="Series 1", values=values,
                          point_colors=[Color(rgb=c) if c else None for c in colors])]
    conf = 0.9 if (legend or data_labels) else 0.75
    style: dict[str, float | str | None] = {"first_slice_angle": round(first, 2)}
    style |= _font_style([b for b in blocks if any(b.id in e.word_blocks for e in legend)]
                         or ([labels[0][1]] if labels else []))
    if style.get("font_size"):
        style["legend_font_size"] = style["font_size"]
    if data_labels:
        style |= {f"label_{k}": v for k, v in _font_style([labels[0][1]]).items()}
    chart = Chart(id=cid, bbox=box, chart_type="pie", categories=cats, series=series,
                  has_legend=bool(legend), legend_pos=_legend_pos(legend, box),
                  data_labels=data_labels, plot_bbox=box,
                  legend_bbox=_union([e.bbox for e in legend]) if legend else None,
                  style=style, source="rule", confidence=conf)
    return Found(chart, consumed, leg_words)


# --------------------------------------------------------------------------- page


def _frame(page: Page, box: BBox, consumed: set[str]) -> tuple[BBox, set[str]]:
    """An invisible rectangle around the chart (PowerPoint's graphic frame) gives its
    bounds; otherwise the union of the chart's parts plus a margin."""
    best: Shape | None = None
    for e in page.elements:
        if isinstance(e, Shape) and e.preset == "rect" and e.fill.kind == "none" and (
                e.stroke is None) and _inside(box, e.bbox, 0.5) and not e.is_background and (
                best is None or (e.bbox[2] - e.bbox[0]) < (best.bbox[2] - best.bbox[0])):
            best = e
    if best is not None:
        return best.bbox, {best.id}
    pad = 6.0
    return (max(0.0, box[0] - pad), max(0.0, box[1] - pad), min(page.width, box[2] + pad),
            min(page.height, box[3] + pad)), set()


def analyze_page(page: Page) -> dict[str, int]:
    """Find vector charts on ``page``; replace their parts by ``Chart`` elements."""
    if page.triage.fallback_image or page.triage.kind == "scanned":
        return {}
    shapes = [e for e in page.elements if isinstance(e, Shape)]
    blocks = [e for e in page.elements if isinstance(e, TextBlock) and not e.rotation]
    if len(shapes) < 3:
        return {}
    hl = [s.bbox[1] for s in shapes if _is_hline(s)]
    vl = [s.bbox[0] for s in shapes if _is_vline(s)]
    found: list[Found] = []
    taken: set[str] = set()
    # pies: slices sharing a centre
    pies: dict[tuple[int, int, int], list[Shape]] = defaultdict(list)
    for s in shapes:
        if s.preset == "pie" and _fill_rgb(s):
            w = s.bbox[2] - s.bbox[0]
            pies[(round(_c(s.bbox)[0]), round(_c(s.bbox)[1]), round(w))].append(s)
    for key in sorted(pies):
        grp = pies[key]
        if len(grp) < 2:
            continue
        f = _pie_chart(page, grp, blocks, shapes, words(page, blocks),
                       f"p{page.index}_c{len(found) + 1}")
        if f is not None:
            found.append(f)
            taken |= f.consumed
    for axis in sorted(find_axes(blocks, hl, vl), key=lambda a: (not a.vertical,
                                                                  a.label_box)):
        if any(b.id in taken for b in axis.blocks):
            continue
        avail_s = [s for s in shapes if s.id not in taken]
        avail_b = [b for b in blocks if b.id not in taken]
        f = _axis_chart(page, axis, avail_s, avail_b, words(page, avail_b),
                        f"p{page.index}_c{len(found) + 1}")
        if f is not None and f.chart.confidence >= MIN_CONFIDENCE:
            found.append(f)
            taken |= f.consumed
    if not found:
        return {}
    stats: dict[str, int] = defaultdict(int)
    by_id = {e.id: e for e in page.elements}
    for f in found:
        # every consumed text block must be chart text only
        partial = []
        for bid in sorted(f.consumed):
            b = by_id.get(bid)
            if isinstance(b, TextBlock) and bid in {w[2] for w in f.words_used}:
                all_w = [w for w in words(page, [b])]
                used = [w for w in f.words_used if w[2] == bid]
                if len(used) < len(all_w):
                    partial.append(bid)
        if partial:
            LOG.info("page %d: chart %s skipped: text blocks %s hold other text", page.index,
                     f.chart.id, partial)
            stats["charts_skipped_partial_text"] += 1
            continue
        frame, extra = _frame(page, _union([by_id[i].bbox for i in f.consumed if i in by_id]
                                           + [f.chart.plot_bbox or f.chart.bbox]), f.consumed)
        f.consumed |= extra
        # groups whose children all go
        for g in page.elements:
            if isinstance(g, Group) and g.child_ids and all(c in f.consumed for c in g.child_ids):
                f.consumed.add(g.id)
        ch = f.chart
        ch.bbox = frame
        ch.fallback = [e for e in page.elements if e.id in f.consumed]
        ch.fallback_ids = sorted(f.consumed)
        ch.z = min((e.z for e in ch.fallback if not isinstance(e, Group)), default=0)
        ch.notes.append(f"charts:{VERSION}")
        # (a surviving group keeps its child ids: the writer skips absent children)
        page.elements = [e for e in page.elements if e.id not in f.consumed]
        page.elements.append(ch)
        stats["charts"] += 1
        stats[f"charts_{ch.chart_type}"] += 1
        LOG.info("page %d: chart %s %s, %d categories x %d series (confidence %.2f)",
                 page.index, ch.id, ch.chart_type, len(ch.categories), len(ch.series),
                 ch.confidence)
    return dict(stats)


def restore(page: Page, chart_id: str, reason: str) -> bool:
    """Undo a chart: its fallback elements return to the page (render check failed)."""
    ch = next((e for e in page.elements if isinstance(e, Chart) and e.id == chart_id), None)
    if ch is None:
        return False
    page.elements = [e for e in page.elements if e is not ch] + list(ch.fallback)
    LOG.info("page %d: chart %s reverted to its shapes (%s)", page.index, chart_id, reason)
    return True


def chart_regions(page: Page) -> list[Image]:
    """Pictures the layout model labelled ``chart`` (raster charts, Max tier)."""
    out = []
    for e in page.elements:
        if isinstance(e, Image) and not e.is_background and any(
                n == "layout:chart" or n.startswith("layout:chart") for n in e.notes):
            out.append(e)
    return out
