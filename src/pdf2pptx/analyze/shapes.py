"""ANALYZE (graphics): vector paths -> built-in PowerPoint shapes.

Per page, in order:

1. **Merge** a fill-only path and the stroke-only path drawn right after it with the
   same geometry into one shape.
2. **Background**: the bottom-most full-bleed rectangle or picture becomes the slide
   background (``is_background``).
3. **Gradients**: a pattern fill that the parser rasterised (``raster:<image>``) is
   fitted with a linear or radial gradient (angle, stops, alpha) on the raster's
   pixels; accepted only if the fitted gradient reproduces the raster.
4. **Connectors**: stroked lines and elbows, and the outlined lines PowerPoint's PDF
   export writes (a thin body polygon plus one polygon per arrowhead), become
   connectors with arrowheads (type and size), dash, cap and join. Endpoints on a
   shape's connection site are attached (``stCxn`` / ``endCxn``).
5. **Presets**: closed paths are fitted to rect, roundRect, ellipse, triangle,
   rtTriangle, diamond, parallelogram, trapezoid, chevron, homePlate and the four
   block arrows (adjust values fitted), with rotation and flips. A preset is accepted
   only when its outline lies within tolerance of the path *and* a rasterised diff of
   the region (fill and stroke) is within tolerance; otherwise the freeform stays.
   Compound paths of disjoint parts are split (and grouped).
6. **Groups**: elements the producer drew as one unit (one marked-content id or one
   form XObject; also the parts of a split compound path) become a ``Group`` when
   they are contiguous in z and nothing else is drawn between them over their area.
7. **SVG**: dense clusters of freeforms that fit no preset (icons, illustrations,
   outlined text) become one SVG picture with a PNG fallback, positioned exactly.

Every decision records ``source="rule"``, a confidence and a ``fit:``/``svg:`` note.
"""

from __future__ import annotations

import hashlib
import io
import itertools
import math
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import numpy as np

from .. import log
from ..ir import (
    Color,
    Connector,
    Fill,
    Gradient,
    GradientStop,
    Group,
    Image,
    Page,
    PathSegment,
    Shape,
    Stroke,
    TextBlock,
)
from . import presets as P

LOG = log.get(__name__)

Pt = tuple[float, float]

BEZIER_STEPS = 16
FIT_TOL_PT = 0.75  # boundary distance accepted for a preset (or FIT_TOL_REL x diagonal)
FIT_TOL_REL = 0.01
FIT_TOL_MAX = 3.0
RASTER_TOL = 0.03  # XOR / union of the rasterised region
RASTER_PX = 600  # longest side of the raster check, px
ADJ_SNAP = 0.1  # pt: a default adjust value within this much of the best fit is kept
LINE_TOL = 0.35  # pt: straightness of a "line"
BODY_MAX_T = 12.0  # pt: thickest outlined line body
ATTACH_TOL = 1.5  # pt (or 1.5 x line width): endpoint on a connection site
SVG_MIN_PATHS = 6  # freeforms in a cluster that becomes one SVG picture
SVG_MIN_SEGMENTS = 120  # or this many segments
SVG_GAP = 3.0  # pt between paths of one cluster
SVG_PX_PER_PT = 2.0  # PNG fallback resolution

# preset candidates by polygon vertex count (curved paths: CURVED)
BY_VERTICES: dict[int, tuple[str, ...]] = {
    3: ("triangle", "rtTriangle"),
    4: ("rect", "diamond", "parallelogram", "trapezoid"),
    5: ("homePlate",),
    6: ("chevron",),
    7: ("rightArrow", "leftArrow", "upArrow", "downArrow"),
}
CURVED = ("ellipse", "roundRect")
PRIORITY = {p: i for i, p in enumerate((
    "rect", "ellipse", "roundRect", "rtTriangle", "triangle", "diamond", "parallelogram",
    "trapezoid", "homePlate", "chevron", "rightArrow", "leftArrow", "upArrow", "downArrow",
    "pie"))}
FLIPS: dict[str, tuple[tuple[bool, bool], ...]] = {
    "rtTriangle": ((False, False), (True, False), (False, True), (True, True)),
    "triangle": ((False, False), (False, True)),
    "parallelogram": ((False, False), (True, False)),
    "trapezoid": ((False, False), (False, True)),
    "chevron": ((False, False), (True, False)),
    "homePlate": ((False, False), (True, False)),
}
# PowerPoint dash presets in multiples of the line width (dash, space, ...)
PRESET_DASHES: dict[str, tuple[float, ...]] = {
    "dash": (4, 3), "dot": (1, 3), "lgDash": (8, 3), "dashDot": (4, 3, 1, 3),
    "lgDashDot": (8, 3, 1, 3), "lgDashDotDot": (8, 3, 1, 3, 1, 3), "sysDash": (3, 1),
    "sysDot": (1, 1), "sysDashDot": (3, 1, 1, 1), "sysDashDotDot": (3, 1, 1, 1, 1, 1),
}


# --------------------------------------------------------------------------- path geometry


@dataclass
class Sub:
    """One subpath: flattened points, its corner points, closed, curved."""

    pts: list[Pt]
    corners: list[Pt]
    closed: bool
    curves: int
    segs: list[PathSegment]


def _bez(p0: Pt, p1: Pt, p2: Pt, p3: Pt, n: int = BEZIER_STEPS) -> list[Pt]:
    out = []
    for k in range(1, n + 1):
        t = k / n
        u = 1 - t
        out.append((u * u * u * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0]
                    + t * t * t * p3[0],
                    u * u * u * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1]
                    + t * t * t * p3[1]))
    return out


def subpaths(path: list[PathSegment] | None) -> list[Sub]:
    out: list[Sub] = []
    cur: Sub | None = None
    for seg in path or []:
        if seg.op == "M" or cur is None:
            if cur is not None and cur.pts:
                out.append(cur)
            start = seg.pts[0] if seg.pts else (0.0, 0.0)
            cur = Sub(pts=[start], corners=[start], closed=False, curves=0, segs=[
                PathSegment(op="M", pts=[start])])
            if seg.op == "M":
                continue
        if seg.op == "L":
            cur.pts.append(seg.pts[0])
            cur.corners.append(seg.pts[0])
        elif seg.op == "C" and len(seg.pts) == 3:
            cur.pts.extend(_bez(cur.pts[-1], seg.pts[0], seg.pts[1], seg.pts[2]))
            cur.corners.append(seg.pts[2])
            cur.curves += 1
        elif seg.op == "Z":
            cur.closed = True
        cur.segs.append(seg)
    if cur is not None and cur.pts:
        out.append(cur)
    return out


def _dist(a: Pt, b: Pt) -> float:
    return math.hypot(a[0] - b[0], a[1] - b[1])


def simplify(pts: list[Pt], closed: bool, tol: float = 0.06) -> list[Pt]:
    """Drop repeated and collinear points."""
    out: list[Pt] = []
    for p in pts:
        if not out or _dist(out[-1], p) > tol:
            out.append(p)
    if closed and len(out) > 1 and _dist(out[0], out[-1]) <= tol:
        out.pop()
    changed = True
    while changed and len(out) > 2:
        changed = False
        n = len(out)
        rng = range(n) if closed else range(1, n - 1)
        for i in rng:
            a, b, c = out[i - 1], out[i], out[(i + 1) % n]
            ab = _dist(a, c)
            if ab < 1e-9:
                continue
            cross = abs((c[0] - a[0]) * (b[1] - a[1]) - (c[1] - a[1]) * (b[0] - a[0])) / ab
            t = ((b[0] - a[0]) * (c[0] - a[0]) + (b[1] - a[1]) * (c[1] - a[1])) / (ab * ab)
            if cross <= tol and -0.01 <= t <= 1.01:
                out.pop(i)
                changed = True
                break
    return out


def resample(pts: list[Pt], closed: bool, step: float) -> np.ndarray:
    """Points every ``step`` pt along the polyline (closed: back to the start)."""
    poly = list(pts) + ([pts[0]] if closed and pts else [])
    out = [poly[0]] if poly else []
    for a, b in itertools.pairwise(poly):
        d = _dist(a, b)
        n = max(1, math.ceil(d / step))
        out.extend((a[0] + (b[0] - a[0]) * k / n, a[1] + (b[1] - a[1]) * k / n)
                   for k in range(1, n + 1))
    return np.asarray(out, dtype=float).reshape(-1, 2)


def _bbox(pts: list[Pt] | np.ndarray) -> tuple[float, float, float, float]:
    a = np.asarray(pts, dtype=float).reshape(-1, 2)
    return float(a[:, 0].min()), float(a[:, 1].min()), float(a[:, 0].max()), float(a[:, 1].max())


def _rot(pts: np.ndarray, deg: float, c: Pt) -> np.ndarray:
    """Rotate points clockwise (y down) by ``deg`` about ``c``."""
    r = math.radians(deg)
    cs, sn = math.cos(r), math.sin(r)
    d = pts - np.asarray(c)
    return np.stack([c[0] + d[:, 0] * cs - d[:, 1] * sn, c[1] + d[:, 0] * sn + d[:, 1] * cs],
                    axis=1)


class _Target:
    """A closed outline to compare candidates against (KD-tree of dense samples)."""

    def __init__(self, pts: list[Pt], step: float) -> None:
        from scipy.spatial import cKDTree

        self.step = step
        self.samples = resample(pts, True, step)
        self.tree = cKDTree(self.samples)

    def error(self, cand: list[Pt]) -> tuple[float, float]:
        """(max, mean) symmetric distance between the outlines."""
        from scipy.spatial import cKDTree

        cs = resample(cand, True, self.step)
        d1, _ = self.tree.query(cs)
        d2, _ = cKDTree(cs).query(self.samples)
        return float(max(d1.max(), d2.max())), float((d1.mean() + d2.mean()) / 2)


# --------------------------------------------------------------------------- preset fitting


@dataclass
class Fit:
    preset: str
    frame: tuple[float, float, float, float]
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    adjust: dict[str, float] = field(default_factory=dict)
    err: float = math.inf
    mean: float = math.inf

    def outline(self) -> list[Pt]:
        x0, y0, x1, y1 = self.frame
        return P.place(P.outline(self.preset, x1 - x0, y1 - y0, self.adjust), self.frame,
                       self.rotation, self.flip_h, self.flip_v)


def _adj_bounds(preset: str, w: float, h: float) -> dict[str, tuple[float, float]]:
    ss = max(min(w, h), 1e-6)
    return {
        "roundRect": {"adj": (0.0, 50000.0)},
        "triangle": {"adj": (0.0, 100000.0)},
        "parallelogram": {"adj": (0.0, 100000.0 * w / ss)},
        "trapezoid": {"adj": (0.0, 50000.0 * w / ss)},
        "chevron": {"adj": (0.0, 100000.0 * w / ss)},
        "homePlate": {"adj": (0.0, 100000.0 * w / ss)},
        "rightArrow": {"adj1": (0.0, 100000.0), "adj2": (0.0, 100000.0 * w / ss)},
        "leftArrow": {"adj1": (0.0, 100000.0), "adj2": (0.0, 100000.0 * w / ss)},
        "upArrow": {"adj1": (0.0, 100000.0), "adj2": (0.0, 100000.0 * h / ss)},
        "downArrow": {"adj1": (0.0, 100000.0), "adj2": (0.0, 100000.0 * h / ss)},
    }.get(preset, {})


def _fit_preset(target: _Target, preset: str, frame: tuple[float, float, float, float],
                rotation: float, flip_h: bool, flip_v: bool) -> Fit:
    from scipy.optimize import minimize, minimize_scalar

    x0, y0, x1, y1 = frame
    w, h = x1 - x0, y1 - y0
    bounds = _adj_bounds(preset, w, h)
    names = sorted(bounds)

    def make(vals: list[float]) -> Fit:
        adj = {n: float(min(max(v, bounds[n][0]), bounds[n][1])) for n, v in zip(
            names, vals, strict=True)}
        return Fit(preset, frame, rotation, flip_h, flip_v, adj)

    def cost(vals: list[float]) -> float:
        return target.error(make(vals).outline())[1]

    if not names:
        best = make([])
    elif len(names) == 1:
        lo, hi = bounds[names[0]]
        grid = np.linspace(lo, hi, 11)
        k = int(np.argmin([cost([g]) for g in grid]))
        a, b = grid[max(k - 1, 0)], grid[min(k + 1, len(grid) - 1)]
        res = minimize_scalar(lambda v: cost([v]), bounds=(a, b), method="bounded",
                              options={"xatol": 20.0})
        best = make([float(res.x)])
    else:
        grids = [np.linspace(*bounds[n], 5) for n in names]
        starts = sorted((cost([a, b]), a, b) for a in grids[0] for b in grids[1])[:1]
        _, a, b = starts[0]
        res = minimize(lambda v: cost(list(v)), np.array([a, b]), method="Nelder-Mead",
                       options={"xatol": 20.0, "fatol": 1e-3, "maxiter": 120})
        best = make([float(v) for v in res.x])
    best.err, best.mean = target.error(best.outline())
    # prefer the preset's default adjust values when they fit as well (exact for
    # shapes drawn with defaults, and nicer to edit)
    if names:
        dflt = Fit(preset, frame, rotation, flip_h, flip_v, {n: float(P.DEFAULTS[preset][n])
                                                            for n in names})
        dflt.err, dflt.mean = target.error(dflt.outline())
        if dflt.err <= best.err + ADJ_SNAP:
            best = dflt
    best.adjust = {k: float(round(v)) for k, v in best.adjust.items()}
    return best


def _frames(pts: np.ndarray, corners: list[Pt]) -> list[tuple[float, tuple[float, float,
                                                                           float, float]]]:
    """(rotation, unrotated frame) candidates: axis-aligned, then the directions of the
    min-area rectangle and of the polygon's longest edges (each + k x 90 degrees)."""
    import cv2

    out = [(0.0, _bbox(pts))]
    rect = cv2.minAreaRect(pts.astype(np.float32))
    box = cv2.boxPoints(rect)
    bases = [math.degrees(math.atan2(float(box[1][1] - box[0][1]), float(box[1][0] - box[0][0])))]
    edges = sorted(((_dist(a, b), a, b) for a, b in zip(corners, corners[1:] + corners[:1],
                                                         strict=True)), reverse=True)
    for _, a, b in edges[:3]:
        bases.append(math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])))
    c = _bbox(pts)
    pivot = ((c[0] + c[2]) / 2, (c[1] + c[3]) / 2)
    seen: set[float] = set()
    for base in bases:
        for k in range(4):
            ang = round((base + 90 * k) % 360.0, 1) % 360.0
            if min(ang % 90, 90 - ang % 90) < 0.3 or ang in seen:
                continue  # axis aligned: covered by rotation 0
            seen.add(ang)
            local = _rot(pts, -ang, pivot)
            lx0, ly0, lx1, ly1 = _bbox(local)
            lc = _rot(np.array([[(lx0 + lx1) / 2, (ly0 + ly1) / 2]]), ang, pivot)[0]
            w, h = lx1 - lx0, ly1 - ly0
            cx, cy = float(lc[0]), float(lc[1])
            out.append((ang, (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)))
    return out


def fit_closed(sub: Sub, stroke_w: float = 0.0) -> Fit | None:
    """Best preset for a closed subpath, or None (keep the freeform)."""
    pts = np.asarray(sub.pts, dtype=float)
    if len(pts) < 3:
        return None
    x0, y0, x1, y1 = _bbox(pts)
    diag = math.hypot(x1 - x0, y1 - y0)
    if diag < 0.5:
        return None
    tol = min(max(FIT_TOL_PT, FIT_TOL_REL * diag), FIT_TOL_MAX)
    corners = simplify(sub.corners, True)
    if sub.curves or len(corners) > 8:  # flattened curves are polygons with many vertices
        cands: tuple[str, ...] = CURVED
    else:
        cands = BY_VERTICES.get(len(corners), ())
    if not cands:
        return None
    step = max(0.2, min(1.0, diag / 300))
    target = _Target(sub.pts, step)
    # fast path: an axis-aligned rectangle
    if not sub.curves and len(corners) == 4 and all(
            min(abs(px - x0), abs(px - x1)) < 0.05 and min(abs(py - y0), abs(py - y1)) < 0.05
            for px, py in corners):
        return Fit("rect", (x0, y0, x1, y1), err=0.0, mean=0.0)
    best: Fit | None = None

    def consider(f: Fit) -> None:
        nonlocal best
        if best is None or (f.err < best.err - 0.05) or (
                abs(f.err - best.err) <= 0.05 and (abs(_signed(f.rotation)), PRIORITY[
                    f.preset]) < (abs(_signed(best.rotation)), PRIORITY[best.preset])):
            best = f

    frames = _frames(pts, corners)
    rot0, frame0 = frames[0]
    for preset in cands:
        for fh_, fv_ in FLIPS.get(preset, ((False, False),)):
            consider(_fit_preset(target, preset, frame0, rot0, fh_, fv_))
    if (best is None or best.err > tol) and len(frames) > 1:
        # rotated: rank every (frame, preset, flips) at default adjust values and
        # optimise only the most promising few
        ranked = []
        for rot, frame in frames[1:]:
            if frame[2] - frame[0] < 0.2 or frame[3] - frame[1] < 0.2:
                continue
            for preset in cands:
                for fh_, fv_ in FLIPS.get(preset, ((False, False),)):
                    f = Fit(preset, frame, rot, fh_, fv_, dict(P.DEFAULTS[preset]))
                    ranked.append((target.error(f.outline())[1], rot, preset, fh_, fv_, frame))
        for _, rot, preset, fh_, fv_, frame in sorted(ranked, key=lambda r: r[:5])[:6]:
            consider(_fit_preset(target, preset, frame, rot, fh_, fv_))
    if (sub.curves or len(corners) > 8) and (best is None or best.err > tol):
        pie = fit_pie(sub, target)
        if pie is not None and (best is None or pie.err < best.err):
            best = pie
    if best is None or best.err > tol:
        return None
    if raster_diff(sub.pts, best.outline(), stroke_w) > RASTER_TOL:
        return None
    return best


def fit_pie(sub: Sub, target: _Target) -> Fit | None:
    """A pie slice: two radii meeting at the centre and one circular arc."""
    corners = simplify(sub.corners, True, 0.05)
    ring = list(sub.pts)
    if len(ring) > 1 and _dist(ring[0], ring[-1]) < 0.05:
        ring.pop()
    # start the ring right after the centre-candidate run so the arc is contiguous
    n = len(ring)
    if n < 6:
        return None
    # the centre: a corner whose neighbouring flattened points are both straight edges
    best = None
    for c in corners:
        k = min(range(n), key=lambda i: _dist(ring[i], c))
        seq = ring[k + 1:] + ring[:k]  # the ring without the centre, in path order
        ds = np.array([_dist(p, c) for p in seq])
        far = ds[ds > 0.5]
        if len(far) < 4:
            continue
        r = float(np.median(far))
        arc = [p for p, d in zip(seq, ds, strict=True) if abs(d - r) <= max(0.6, 0.01 * r)]
        if len(arc) < max(4, 0.6 * len(far)):
            continue
        angs = [math.degrees(math.atan2(p[1] - c[1], p[0] - c[0])) % 360.0 for p in arc]
        # order along the path; PowerPoint sweeps clockwise from adj1 to adj2 (y down)
        a0, a1 = angs[0], angs[-1]
        mid = angs[len(angs) // 2]
        cw = (mid - a0) % 360.0 <= (a1 - a0) % 360.0
        st, en = (a0, a1) if cw else (a1, a0)
        frame = (c[0] - r, c[1] - r, c[0] + r, c[1] + r)
        f = Fit("pie", frame, adjust={"adj1": float(round(st * 60000)),
                                      "adj2": float(round(en * 60000))})
        f.err, f.mean = target.error(f.outline())
        if best is None or f.err < best.err:
            best = f
    return best


def _signed(a: float) -> float:
    return a - 360.0 if a > 180.0 else a


def raster_diff(a: list[Pt], b: list[Pt], stroke_w: float = 0.0) -> float:
    """XOR / union of two closed outlines rendered over their common region."""
    import cv2

    pts = np.asarray(list(a) + list(b), dtype=float)
    x0, y0, x1, y1 = _bbox(pts)
    pad = stroke_w / 2 + 1.0
    x0, y0, x1, y1 = x0 - pad, y0 - pad, x1 + pad, y1 + pad
    s = min(4.0, RASTER_PX / max(x1 - x0, y1 - y0, 1e-6))
    w, h = math.ceil((x1 - x0) * s) + 1, math.ceil((y1 - y0) * s) + 1

    def mask(poly: list[Pt]) -> np.ndarray:
        m = np.zeros((h, w), np.uint8)
        q = np.round((np.asarray(poly) - [x0, y0]) * s * 16).astype(np.int32)
        cv2.fillPoly(m, [q], 1, lineType=cv2.LINE_8, shift=4)
        if stroke_w > 0:
            cv2.polylines(m, [q], True, 1, thickness=max(1, round(stroke_w * s)),
                          shift=4)
        return m

    ma, mb = mask(list(a)), mask(list(b))
    union = int(np.count_nonzero(ma | mb))
    return float(np.count_nonzero(ma ^ mb)) / union if union else 0.0


# --------------------------------------------------------------------------- lines


def _is_line(sub: Sub) -> tuple[Pt, Pt] | None:
    """Endpoints if the open subpath is one straight segment."""
    if sub.closed or sub.curves or len(sub.pts) < 2:
        return None
    a, b = sub.pts[0], sub.pts[-1]
    d = _dist(a, b)
    if d < 0.5:
        return None
    for p in sub.pts[1:-1]:
        cross = abs((b[0] - a[0]) * (p[1] - a[1]) - (b[1] - a[1]) * (p[0] - a[0])) / d
        if cross > LINE_TOL:
            return None
    return a, b


def _elbow(pts: list[Pt]) -> list[Pt] | None:
    """Corner points if the polyline is axis-aligned with 1 or 2 bends."""
    c = simplify(pts, False, 0.2)
    if len(c) not in (3, 4):
        return None
    for a, b in itertools.pairwise(c):
        if abs(a[0] - b[0]) > 0.3 and abs(a[1] - b[1]) > 0.3:
            return None
    dirs = ["h" if abs(a[1] - b[1]) <= 0.3 else "v" for a, b in itertools.pairwise(c)]
    if any(d1 == d2 for d1, d2 in itertools.pairwise(dirs)):
        return None
    return c


_CXN_LOCAL = {
    "bentConnector2": lambda w, h, a: [(0.0, 0.0), (w, 0.0), (w, h)],
    "bentConnector3": lambda w, h, a: [(0.0, 0.0), (w * a / 100000, 0.0), (w * a / 100000, h),
                                       (w, h)],
}


def connector_for(pts: list[Pt]) -> tuple[str, tuple[float, float, float, float], float, bool,
                                          bool, dict[str, float]] | None:
    """(preset, frame, rotation, flip_h, flip_v, adjust) of a straight or elbow connector
    whose path runs from ``pts[0]`` to ``pts[-1]``."""
    a, b = pts[0], pts[-1]
    if len(pts) == 2:
        x0, x1 = sorted((a[0], b[0]))
        y0, y1 = sorted((a[1], b[1]))
        return "line", (x0, y0, x1, y1), 0.0, b[0] < a[0], b[1] < a[1], {}
    preset = "bentConnector2" if len(pts) == 3 else "bentConnector3"
    bx0, by0, bx1, by1 = _bbox(pts)
    cx, cy = (bx0 + bx1) / 2, (by0 + by1) / 2
    best = None
    for rot in (0.0, 90.0, 180.0, 270.0):
        w, h = (bx1 - bx0, by1 - by0) if rot in (0.0, 180.0) else (by1 - by0, bx1 - bx0)
        frame = (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
        adj = 50000.0
        if preset == "bentConnector3":
            # the middle segment's position along the local x axis
            local = _rot(np.asarray(pts, dtype=float), -rot, (cx, cy))
            if w > 1e-6:
                adj = float(round((local[1][0] - frame[0]) / w * 100000))
        for fh in (False, True):
            for fv in (False, True):
                cand = P.place(_CXN_LOCAL[preset](w, h, adj), frame, rot, fh, fv)
                err = max(_dist(p, q) for p, q in zip(cand, pts, strict=True))
                key = (round(err, 2), rot != 0.0, fh, fv)
                if best is None or key < best[0]:
                    best = (key, (preset, frame, rot, fh, fv, {"adj1": adj} if preset ==
                                  "bentConnector3" else {}))
    if best is None or best[0][0] > 0.5:
        return None
    return best[1]


def dash_preset(stroke: Stroke) -> Stroke:
    """Custom dash pattern -> PowerPoint preset dash when it matches one."""
    if stroke.dash != "custom" or not stroke.dash_pattern or stroke.width <= 0:
        return stroke
    pat = [v / stroke.width for v in stroke.dash_pattern]
    if len(pat) % 2:
        pat = pat * 2
    for name, ref in PRESET_DASHES.items():
        if len(ref) != len(pat):
            continue
        # round caps add a width to each dash in the PDF; PowerPoint's preset does not
        adj = [p - (1.0 if stroke.cap in ("round", "square") and i % 2 == 0 else 0.0)
               if stroke.cap in ("round", "square") else p for i, p in enumerate(pat)]
        if all(abs(p - r) <= 0.12 * r + 0.05 for p, r in zip(adj, ref, strict=True)):
            stroke.dash, stroke.dash_pattern = name, None
            return stroke
    return stroke


# --------------------------------------------------------------------------- outlined lines


def _thickness(poly: list[Pt]) -> float:
    """Width of a thin strip from its area and perimeter (L x t: P = 2(L + t), A = L t)."""
    a = np.asarray(poly, dtype=float)
    area = 0.5 * abs(float(np.dot(a[:, 0], np.roll(a[:, 1], -1))
                           - np.dot(a[:, 1], np.roll(a[:, 0], -1))))
    per = sum(_dist(p, q) for p, q in zip(poly, poly[1:] + poly[:1], strict=True))
    half = per / 2
    disc = half * half - 4 * area
    return (half - math.sqrt(disc)) / 2 if disc >= 0 else (2 * area / per if per else 0.0)


def _clusters(vals: list[float], gap: float) -> list[tuple[float, float]]:
    out: list[list[float]] = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= gap:
            out[-1].append(v)
        else:
            out.append([v])
    return [(c[0], c[-1]) for c in out]


def _snap(v: float, cl: list[tuple[float, float]]) -> float:
    for lo, hi in cl:
        if lo - 1e-6 <= v <= hi + 1e-6:
            return (lo + hi) / 2
    return v


def centerline(sub: Sub, t: float) -> list[Pt] | None:
    """The centre line of an outlined (thin, filled) line body, start -> end."""
    corners = simplify(sub.corners, True, 0.05)
    if len(corners) == 4:  # a straight body at any angle: midpoints of the short edges
        edges = [(corners[i], corners[(i + 1) % 4]) for i in range(4)]
        lens = [_dist(a, b) for a, b in edges]
        k = int(np.argmin(lens))
        if lens[k] <= 1.6 * t and abs(lens[k] - lens[(k + 2) % 4]) <= 0.3 * t + 0.1:
            e1, e2 = edges[k], edges[(k + 2) % 4]
            m1 = ((e1[0][0] + e1[1][0]) / 2, (e1[0][1] + e1[1][1]) / 2)
            m2 = ((e2[0][0] + e2[1][0]) / 2, (e2[0][1] + e2[1][1]) / 2)
            return [m1, m2]
    # an axis-aligned elbow: snap every outline point to the centre of its coordinate
    # cluster; the snapped walk goes out along the line and back
    xs = _clusters([p[0] for p in sub.pts], 0.6 * t + 0.05)
    ys = _clusters([p[1] for p in sub.pts], 0.6 * t + 0.05)
    if any(hi - lo > 1.3 * t + 0.1 for lo, hi in xs + ys):
        return None
    walk: list[Pt] = []
    for p in sub.pts:
        q = (_snap(p[0], xs), _snap(p[1], ys))
        if not walk or _dist(walk[-1], q) > 1e-6:
            walk.append(q)
    if len(walk) > 1 and _dist(walk[0], walk[-1]) <= 1e-6:
        walk.pop()
    n = len(walk)
    if n < 2:
        return None
    turns = [i for i in range(n) if n == 2 or _dist(walk[i - 1], walk[(i + 1) % n]) <= 1e-6]
    if len(turns) != 2:
        return None
    i, j = turns
    line = [walk[k % n] for k in range(i, j + 1)] if j > i else None
    if line is None or len(line) > 4:
        return None
    # a walk that is not a clean out-and-back is not a line
    back = [walk[k % n] for k in range(j, i + n + 1)]
    if [*reversed(back)] != line:
        return None
    return line


def head_kind(sub: Sub, tip_dir: Pt) -> str:
    if sub.curves:
        # an oval head lies on its ellipse; an open "arrow" head is outlined with round
        # joins (curves too) but is far from elliptical
        x0, y0, x1, y1 = _bbox(sub.pts)
        a, b = max((x1 - x0) / 2, 1e-6), max((y1 - y0) / 2, 1e-6)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        r = [math.hypot((x - cx) / a, (y - cy) / b) for x, y in sub.pts]
        if max(r) - min(r) < 0.15:
            return "oval"
        return "arrow"
    c = simplify(sub.corners, True, 0.05)
    if len(c) == 3:
        return "triangle"
    if len(c) == 4:
        # stealth: one vertex is concave (the notch); diamond: convex and symmetric
        a = np.asarray(c)
        crosses = []
        for i in range(4):
            p0, p1, p2 = a[i - 1], a[i], a[(i + 1) % 4]
            d1, d2 = p1 - p0, p2 - p1
            crosses.append(float(d1[0] * d2[1] - d1[1] * d2[0]))
        signs = {math.copysign(1, v) for v in crosses if abs(v) > 1e-6}
        return "diamond" if len(signs) == 1 else "stealth"
    return "arrow"


def _size_word(r: float) -> str:
    return "sm" if r < 2.5 else ("med" if r < 4.0 else "lg")


def decode_outlined(shape: Shape) -> Shape | None:
    """A filled compound path that is an outlined line + arrowheads -> a connector."""
    if shape.fill.kind != "solid" or shape.stroke is not None or shape.fill.color is None:
        return None
    subs = subpaths(shape.path)
    if not subs or not all(s.closed for s in subs) or len(subs) > 3:
        return None
    lengths = [max(_bbox(s.pts)[2] - _bbox(s.pts)[0], _bbox(s.pts)[3] - _bbox(s.pts)[1])
               for s in subs]
    bi = int(np.argmax(lengths))
    body = subs[bi]
    if body.curves > 2:
        return None
    t = _thickness(body.pts)
    if not 0.05 < t <= BODY_MAX_T or lengths[bi] < 4 * t:
        return None
    line = centerline(body, t)
    if line is None:
        return None
    heads = [s for k, s in enumerate(subs) if k != bi]
    ends: dict[int, tuple[str, Pt, float, float]] = {}
    for h in heads:
        hb = _bbox(h.pts)
        hc = ((hb[0] + hb[2]) / 2, (hb[1] + hb[3]) / 2)
        if max(hb[2] - hb[0], hb[3] - hb[1]) > 12 * max(t, 1.0):
            return None
        k = 0 if _dist(hc, line[0]) <= _dist(hc, line[-1]) else -1
        p, q = line[k], line[1] if k == 0 else line[-2]
        u = (p[0] - q[0], p[1] - q[1])
        un = math.hypot(*u) or 1.0
        u = (u[0] / un, u[1] / un)
        kind = head_kind(h, u)
        proj = [(x - p[0]) * u[0] + (y - p[1]) * u[1] for x, y in h.pts]
        perp = [-(x - p[0]) * u[1] + (y - p[1]) * u[0] for x, y in h.pts]
        length, width = max(proj) - min(proj), max(perp) - min(perp)
        if kind in ("oval", "diamond"):
            tip = hc
        else:
            far = max(proj)
            tip = (p[0] + u[0] * far, p[1] + u[1] * far)
        ends[0 if k == 0 else 1] = (kind, tip, width, length)
    backwards = line[-1][0] < line[0][0] - 0.01 or (
        abs(line[-1][0] - line[0][0]) <= 0.01 and line[-1][1] < line[0][1])
    # one head: it is the tail (PowerPoint's default); none or two: left -> right
    if (0 in ends and 1 not in ends) or (len(ends) != 1 and backwards):
        line = list(reversed(line))
        ends = {1 - k: v for k, v in ends.items()}
    pts = list(line)
    for k, (_, tip, _, _) in ends.items():
        pts[0 if k == 0 else -1] = tip
    if len(pts) > 2:  # keep the elbow axis-aligned after moving the ends to the tips
        pts[1] = (pts[1][0], pts[0][1]) if abs(line[0][1] - line[1][1]) <= 0.3 else (
            pts[0][0], pts[1][1])
        pts[-2] = (pts[-2][0], pts[-1][1]) if abs(line[-1][1] - line[-2][1]) <= 0.3 else (
            pts[-1][0], pts[-2][1])
    cx = connector_for(pts)
    if cx is None:
        return None
    preset, frame, rot, fh, fv, adj = cx
    eff = max(t, 2.0)
    stroke = Stroke(color=shape.fill.color, width=round(t, 3), cap="flat", join="round")
    for k, (kind, _, width, length) in sorted(ends.items()):
        size = f"{_size_word(width / eff)}-{_size_word(length / eff)}"
        if kind in ("oval", "diamond"):
            size = f"{_size_word(width / eff)}-{_size_word(width / eff)}"
        if k == 0:
            stroke.head_end, stroke.head_size = kind, size
        else:
            stroke.tail_end, stroke.tail_size = kind, size
    return _connector_shape(shape, preset, frame, rot, fh, fv, adj, stroke, pts,
                            note="fit:outlined_line")


def _connector_shape(src: Shape, preset: str, frame: tuple[float, float, float, float],
                     rot: float, fh: bool, fv: bool, adj: dict[str, float], stroke: Stroke,
                     pts: list[Pt], note: str) -> Shape:
    kind = "straight" if preset == "line" else "elbow"
    out = Shape(id=src.id, bbox=_r4(frame), z=src.z, source="rule", confidence=0.9,
                notes=[*[n for n in src.notes if n.startswith(("mcid:", "form:"))], note,
                       f"cxn_ends:{pts[0][0]:.2f},{pts[0][1]:.2f},{pts[-1][0]:.2f},"
                       f"{pts[-1][1]:.2f}"],
                preset=preset, adjust=adj, fill=Fill(), stroke=stroke, rotation=rot,
                flip_h=fh, flip_v=fv, connector=Connector(kind=kind))
    return out


def _r4(b: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
    return (round(b[0], 3), round(b[1], 3), round(b[2], 3), round(b[3], 3))


# --------------------------------------------------------------------------- per shape


def _same_path(a: Shape, b: Shape, tol: float = 0.05) -> bool:
    pa, pb = a.path or [], b.path or []
    if len(pa) != len(pb):
        return False
    for sa, sb in zip(pa, pb, strict=True):
        if sa.op != sb.op or len(sa.pts) != len(sb.pts):
            return False
        for p, q in zip(sa.pts, sb.pts, strict=True):
            if abs(p[0] - q[0]) > tol or abs(p[1] - q[1]) > tol:
                return False
    return True


def merge_fill_stroke(elements: list[Any]) -> int:
    """Fill-only path + stroke-only path with the same geometry, drawn one after the
    other -> one shape. Returns the number of merges."""
    order = sorted(range(len(elements)), key=lambda i: elements[i].z)
    drop: set[int] = set()
    for i, j in itertools.pairwise(order):
        a, b = elements[i], elements[j]
        if i in drop or not isinstance(a, Shape) or not isinstance(b, Shape):
            continue
        if a.fill.kind in ("solid", "pattern") and a.stroke is None and \
                b.fill.kind == "none" and b.stroke is not None and _same_path(a, b):
            a.stroke = b.stroke
            a.notes.append(f"merged:{b.id}")
            drop.add(j)
    for j in sorted(drop, reverse=True):
        elements.pop(j)
    return len(drop)


def _part(shape: Shape, sub: Sub, k: int) -> Shape:
    segs = list(sub.segs) + ([PathSegment(op="Z")] if sub.closed and not any(
        s.op == "Z" for s in sub.segs) else [])
    return Shape(id=f"{shape.id}_{k}", bbox=_r4(_bbox(sub.pts)), z=shape.z, path=segs,
                 fill=shape.fill, stroke=shape.stroke, opacity=shape.opacity,
                 notes=[*shape.notes, f"part_of:{shape.id}"], source=shape.source)


def fit_shape(shape: Shape) -> Shape | list[Shape] | None:
    """A preset / connector version of ``shape``; a list for a split compound path;
    None to keep the freeform."""
    if shape.preset or not shape.path or shape.is_background:
        return None
    if shape.fill.kind == "pattern":
        return None
    subs = subpaths(shape.path)
    if not subs:
        return None
    if len(subs) > 1:
        outlined = decode_outlined(shape)
        if outlined is not None:
            return outlined
        if _disjoint(subs) and len(subs) <= 64:
            parts = [_part(shape, s, k + 1) for k, s in enumerate(subs)]
            fitted = [fit_shape(p) for p in parts]
            if all(isinstance(f, Shape) for f in fitted):
                return [f for f in fitted if isinstance(f, Shape)]
        return None
    sub = subs[0]
    sw = shape.stroke.width if shape.stroke is not None else 0.0
    if shape.stroke is not None:
        shape.stroke = dash_preset(shape.stroke)
    if shape.fill.kind == "none" and shape.stroke is not None and not sub.closed:
        ends = _is_line(sub)
        poly = [ends[0], ends[1]] if ends else _elbow(sub.pts)
        if poly is None:
            return None
        cx = connector_for(poly)
        if cx is None:
            return None
        preset, frame, rot, fh, fv, adj = cx
        return _connector_shape(shape, preset, frame, rot, fh, fv, adj, shape.stroke, poly,
                                note="fit:line" if preset == "line" else "fit:elbow")
    if not sub.closed and shape.fill.kind != "none":
        sub = Sub(sub.pts, sub.corners, True, sub.curves, sub.segs)  # fills close implicitly
    if not sub.closed:
        return None
    if len(subpaths(shape.path)) == 1 and shape.fill.kind == "solid" and shape.stroke is None:
        outlined = decode_outlined(shape)  # a plain outlined line without heads
        if outlined is not None:
            return outlined
    fit = fit_closed(sub, sw)
    if fit is None:
        return None
    out = Shape(id=shape.id, bbox=_r4(fit.frame), z=shape.z, source="rule",
                confidence=round(max(0.5, 1.0 - fit.err / 4.0), 3),
                notes=[*shape.notes, f"fit:{fit.preset}:err={fit.err:.2f}"],
                preset=fit.preset, adjust=fit.adjust, fill=shape.fill, stroke=shape.stroke,
                opacity=shape.opacity, rotation=round(fit.rotation, 2), flip_h=fit.flip_h,
                flip_v=fit.flip_v, hyperlink=shape.hyperlink)
    return out


def _disjoint(subs: list[Sub]) -> bool:
    boxes = [_bbox(s.pts) for s in subs]
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a, b = boxes[i], boxes[j]
            if a[0] < b[2] - 0.05 and b[0] < a[2] - 0.05 and a[1] < b[3] - 0.05 and \
                    b[1] < a[3] - 0.05:
                return False
    return True


# --------------------------------------------------------------------------- arrowheads of
# stroked lines (other producers draw the head as a separate filled triangle)


def attach_separate_heads(elements: list[Any]) -> int:
    lines = [e for e in elements if isinstance(e, Shape) and e.preset == "line"
             and e.stroke is not None and "fit:line" in e.notes]
    heads = [e for e in elements if isinstance(e, Shape) and e.preset in ("triangle",)
             and e.fill.kind == "solid"]
    used: set[str] = set()
    for ln in lines:
        ends = _ends(ln)
        if ends is None:
            continue
        for h in heads:
            if h.id in used or abs(h.z - ln.z) > 2 or ln.stroke is None:
                continue
            hb = h.bbox
            if max(hb[2] - hb[0], hb[3] - hb[1]) > 12 * max(ln.stroke.width, 1.0):
                continue
            for k, p in enumerate(ends):
                inside = hb[0] - 0.5 <= p[0] <= hb[2] + 0.5 and hb[1] - 0.5 <= p[1] <= hb[3] + 0.5
                if not inside:
                    continue
                w = max(ln.stroke.width, 2.0)
                size = f"{_size_word(max(hb[2] - hb[0], hb[3] - hb[1]) / w)}-med"
                if k == 0:
                    ln.stroke.head_end, ln.stroke.head_size = "triangle", size
                else:
                    ln.stroke.tail_end, ln.stroke.tail_size = "triangle", size
                used.add(h.id)
                ln.notes.append(f"head_from:{h.id}")
                break
    if used:
        elements[:] = [e for e in elements if getattr(e, "id", None) not in used]
    return len(used)


def _ends(e: Shape) -> tuple[Pt, Pt] | None:
    for n in e.notes:
        if n.startswith("cxn_ends:"):
            v = [float(x) for x in n.split(":", 1)[1].split(",")]
            return (v[0], v[1]), (v[2], v[3])
    return None


# --------------------------------------------------------------------------- connections


def attach_connectors(elements: list[Any]) -> int:
    targets = [e for e in elements if isinstance(e, Shape) and e.preset in (
        "rect", "roundRect", "ellipse", "diamond") and not e.is_background and e.connector is None]
    n = 0
    for c in elements:
        if not isinstance(c, Shape) or c.connector is None:
            continue
        ends = _ends(c)
        if ends is None:
            continue
        tol = max(ATTACH_TOL, 1.5 * (c.stroke.width if c.stroke else 0.0))
        for k, p in enumerate(ends):
            best: tuple[float, str, int] | None = None
            for t in targets:
                x0, y0, x1, y1 = t.bbox
                sites = P.place(P.connection_sites(t.preset or "", x1 - x0, y1 - y0), t.bbox,
                                t.rotation, t.flip_h, t.flip_v)
                for si, s in enumerate(sites):
                    d = _dist(p, s)
                    if d <= tol and (best is None or (d, t.id, si) < best):
                        best = (d, t.id, si)
            if best is not None:
                if k == 0:
                    c.connector.start_id, c.connector.start_site = best[1], best[2]
                else:
                    c.connector.end_id, c.connector.end_site = best[1], best[2]
                n += 1
    return n


# --------------------------------------------------------------------------- background


def detect_background(page: Page) -> str | None:
    """Mark the bottom-most full-bleed rectangle / picture as the slide background."""
    els = sorted(page.elements, key=lambda e: e.z)
    for e in els:
        if isinstance(e, TextBlock):
            return None
        x0, y0, x1, y1 = e.bbox
        full = (x0 <= 1.0 and y0 <= 1.0 and x1 >= page.width - 1.0 and y1 >= page.height - 1.0)
        if isinstance(e, Shape):
            subs = subpaths(e.path)
            rect = (len(subs) == 1 and not subs[0].curves
                    and len(simplify(subs[0].corners, True)) == 4) or e.preset == "rect"
            if full and rect and e.fill.kind == "solid" and e.fill.color is not None and \
                    e.fill.color.alpha >= 0.999 and e.stroke is None and not e.rotation:
                e.is_background = True
                e.preset, e.path = "rect", None
                e.bbox = (0.0, 0.0, page.width, page.height)
                e.source = "rule"
                e.notes.append("background")
                return e.id
            return None
        if isinstance(e, Image):
            if full and not e.rotation and e.opacity >= 0.999 and e.mask_file is None:
                e.is_background = True
                e.source = "rule"
                e.notes.append("background")
                return e.id
            return None
        return None
    return None


# --------------------------------------------------------------------------- gradients


GRAD_BINS = 64
GRAD_TOL = 4.0  # mean abs error (0..255) of the reconstructed gradient
GRAD_STOP_TOL = 3.0  # piecewise-linear tolerance (0..255) when choosing stops


def _stops(profile: np.ndarray, valid: np.ndarray) -> list[tuple[float, np.ndarray]]:
    """Douglas-Peucker on a (bins x 4) RGBA profile -> [(pos 0..1, rgba)]."""
    idx = np.flatnonzero(valid)
    if len(idx) < 2:
        return []
    xs = (idx + 0.5) / len(profile)
    ys = profile[idx]

    def rec(i: int, j: int, keep: set[int]) -> None:
        if j <= i + 1:
            return
        t = (xs[i + 1:j] - xs[i]) / (xs[j] - xs[i])
        interp = ys[i] + np.outer(t, ys[j] - ys[i])
        err = np.abs(ys[i + 1:j] - interp).max(axis=1)
        k = int(np.argmax(err))
        if err[k] > GRAD_STOP_TOL:
            keep.add(i + 1 + k)
            rec(i, i + 1 + k, keep)
            rec(i + 1 + k, j, keep)

    keep = {0, len(idx) - 1}
    rec(0, len(idx) - 1, keep)
    out = [(float(xs[k]), ys[k]) for k in sorted(keep)]
    # extend the end stops to the ends of the axis
    out[0] = (0.0, out[0][1])
    out[-1] = (1.0, out[-1][1])
    return out


def _lin_eval(stops: list[tuple[float, np.ndarray]], t: np.ndarray) -> np.ndarray:
    pos = np.array([s[0] for s in stops])
    cols = np.stack([s[1] for s in stops])
    return np.stack([np.interp(t, pos, cols[:, c]) for c in range(cols.shape[1])], axis=-1)


def fit_gradient(rgba: np.ndarray, mask: np.ndarray, shape_box: tuple[float, float, float, float],
                 img_box: tuple[float, float, float, float]) -> tuple[Gradient, float] | None:
    """Linear (any angle) or radial gradient reproducing the masked raster pixels.

    Positions follow DrawingML: a linear gradient (``a:lin scaled=0``) runs along its
    angle across the shape's bounding box; a radial (``path="circle"``) one from the
    box centre to its corners. Returns (gradient, mean abs error) or None."""
    h, w = mask.shape
    if mask.sum() < 50:
        return None
    ys, xs = np.nonzero(mask)
    sx = (img_box[2] - img_box[0]) / w
    sy = (img_box[3] - img_box[1]) / h
    px = img_box[0] + (xs + 0.5) * sx
    py = img_box[1] + (ys + 0.5) * sy
    col = rgba[ys, xs].astype(float)
    bx0, by0, bx1, by1 = shape_box
    bw, bh = max(bx1 - bx0, 1e-6), max(by1 - by0, 1e-6)

    def profile(t: np.ndarray) -> tuple[np.ndarray, np.ndarray, float]:
        k = np.clip((t * GRAD_BINS).astype(int), 0, GRAD_BINS - 1)
        cnt = np.bincount(k, minlength=GRAD_BINS)
        prof = np.stack([np.bincount(k, col[:, c], minlength=GRAD_BINS) for c in range(4)],
                        axis=1) / np.maximum(cnt, 1)[:, None]
        resid = float(np.abs(col - prof[k]).mean())
        return prof, cnt > 0, resid

    best: tuple[float, str, float, np.ndarray] | None = None
    for ang in range(0, 360, 3):
        r = math.radians(ang)
        u = (math.cos(r), math.sin(r))
        corners = [(bx0, by0), (bx1, by0), (bx0, by1), (bx1, by1)]
        proj = [cx * u[0] + cy * u[1] for cx, cy in corners]
        lo, hi = min(proj), max(proj)
        t = np.clip((px * u[0] + py * u[1] - lo) / max(hi - lo, 1e-6), 0, 1)
        _, _, resid = profile(t)
        if best is None or resid < best[0] - 1e-9:
            best = (resid, "linear", float(ang), t)
    # refine the angle
    if best is not None:
        a0 = best[2]
        for ang10 in range(int(a0 * 10) - 30, int(a0 * 10) + 31, 5):
            ang = (ang10 / 10.0) % 360
            r = math.radians(ang)
            u = (math.cos(r), math.sin(r))
            corners = [(bx0, by0), (bx1, by0), (bx0, by1), (bx1, by1)]
            proj = [cx * u[0] + cy * u[1] for cx, cy in corners]
            lo, hi = min(proj), max(proj)
            t = np.clip((px * u[0] + py * u[1] - lo) / max(hi - lo, 1e-6), 0, 1)
            _, _, resid = profile(t)
            if resid < best[0] - 1e-9:
                best = (resid, "linear", ang, t)
    # radial about the box centre
    cxp, cyp = (bx0 + bx1) / 2, (by0 + by1) / 2
    rmax = math.hypot(bw / 2, bh / 2)
    tr = np.clip(np.hypot(px - cxp, py - cyp) / rmax, 0, 1)
    _, _, rres = profile(tr)
    if best is None or rres < best[0] - 0.5:
        best = (rres, "radial", 0.0, tr)
    assert best is not None
    resid, kind, ang, t = best
    prof, valid, _ = profile(t)
    stops = _stops(prof, valid)
    if len(stops) < 2:
        return None
    if kind == "linear" and ang >= 180.0:
        # same gradient, angle in 0..180 with the stops reversed (PowerPoint's UI range)
        ang -= 180.0
        stops = [(1.0 - p, c) for p, c in reversed(stops)]
        t = 1.0 - t
    err = float(np.abs(col - _lin_eval(stops, t)).mean())
    if err > GRAD_TOL:
        return None
    gstops = [GradientStop(pos=round(p, 4), color=Color(
        rgb=f"{round(c[0]):02X}{round(c[1]):02X}{round(c[2]):02X}",
        alpha=round(float(c[3]) / 255.0, 3))) for p, c in stops]
    g = Gradient(kind=kind, angle=round(ang, 1), stops=gstops,
                 center=(0.5, 0.5) if kind == "radial" else None)
    return g, err


def gradients(page: Page, work_dir: Path) -> int:
    """Pattern-filled paths (rasterised by the parser) -> native gradients."""
    import cv2
    from PIL import Image as PILImage

    imgs = {e.id: e for e in page.elements if isinstance(e, Image)}
    n = 0
    drop: set[str] = set()
    for e in page.elements:
        if not isinstance(e, Shape) or e.fill.kind != "pattern" or not e.fill.image_id:
            continue
        img = imgs.get(e.fill.image_id)
        if img is None:
            continue
        try:
            rgba = np.asarray(PILImage.open(work_dir / img.file).convert("RGBA"))
        except OSError:
            continue
        h, w = rgba.shape[:2]
        ix0, iy0, ix1, iy1 = img.bbox
        m = np.zeros((h, w), np.uint8)
        polys = []
        for sub in subpaths(e.path):
            q = (np.asarray(sub.pts) - [ix0, iy0]) * [w / (ix1 - ix0), h / (iy1 - iy0)]
            polys.append(np.round(q * 16).astype(np.int32))
        cv2.fillPoly(m, polys, 1, shift=4)
        m = cv2.erode(m, np.ones((5, 5), np.uint8)) & (rgba[..., 3] > 250)
        res = fit_gradient(rgba, m.astype(bool), e.bbox, img.bbox)
        if res is None:
            e.notes.append("gradient:no_fit")
            continue
        g, err = res
        e.fill = Fill(kind="gradient", gradient=g)
        e.notes.append(f"gradient:{g.kind}:err={err:.2f}")
        e.source = "rule"
        drop.add(img.id)
        n += 1
    if drop:
        page.elements = [e for e in page.elements if getattr(e, "id", None) not in drop]
    return n


# --------------------------------------------------------------------------- groups


def _group_key(e: Any) -> str | None:
    forms = [n for n in e.notes if n.startswith("form:")]
    if forms:
        return forms[0]
    mc = [n for n in e.notes if n.startswith("mcid:")]
    return mc[0] if mc else None


def _overlap(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> bool:
    return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]


def make_groups(page: Page, extra: dict[str, list[str]]) -> int:
    """Group elements drawn as one unit. ``extra``: split compound path -> part ids."""
    els = [e for e in page.elements if isinstance(e, (Shape, Image)) and not e.is_background
           and "bullet_image" not in e.notes]
    keyed: dict[str, list[Any]] = defaultdict(list)
    for e in els:
        k = _group_key(e)
        if k is not None:
            keyed[k].append(e)
    for src, ids in extra.items():
        members = [e for e in els if e.id in set(ids)]
        if len(members) >= 2:
            keyed[f"split:{src}"] = members
    taken: set[str] = set()
    groups: list[Group] = []
    allz = sorted(page.elements, key=lambda e: e.z)
    for key in sorted(keyed, key=lambda k: min(e.z for e in keyed[k])):
        members = [e for e in keyed[key] if e.id not in taken]
        if len(members) < 2 or len(members) > 200:
            continue
        ids = {e.id for e in members}
        box = (min(e.bbox[0] for e in members), min(e.bbox[1] for e in members),
               max(e.bbox[2] for e in members), max(e.bbox[3] for e in members))
        if (box[2] - box[0]) * (box[3] - box[1]) > 0.8 * page.width * page.height:
            continue
        z0, z1 = min(e.z for e in members), max(e.z for e in members)
        between = [e for e in allz if z0 < e.z < z1 and e.id not in ids]
        if any(_overlap(e.bbox, box) for e in between):
            continue
        taken |= ids
        gid = f"p{page.index}_grp{len(groups) + 1}"
        groups.append(Group(id=gid, bbox=_r4(box), z=z0, source="rule", confidence=0.9,
                            child_ids=[e.id for e in sorted(members, key=lambda e: (e.z, e.id))],
                            name=key, notes=[f"group:{key}"]))
    page.elements.extend(groups)
    return len(groups)


# --------------------------------------------------------------------------- SVG clusters


def _svg_path_d(shape: Shape, ox: float, oy: float) -> str:
    out = []
    for seg in shape.path or []:
        pts = " ".join(f"{x - ox:.3f} {y - oy:.3f}" for x, y in seg.pts)
        out.append(seg.op + (" " + pts if pts else ""))
    return " ".join(out)


def _svg_paint(c: Color | None) -> tuple[str, str]:
    if c is None:
        return "none", ""
    op = f' {{}}-opacity="{c.alpha:.3f}"' if c.alpha < 0.999 else ""
    return f"#{(c.rgb or '000000').upper()}", op


def svg_document(shapes: list[Shape], box: tuple[float, float, float, float]) -> str:
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{w:.3f}pt" height="{h:.3f}pt" '
             f'viewBox="0 0 {w:.3f} {h:.3f}">']
    for s in shapes:
        fill, fop = _svg_paint(s.fill.color if s.fill.kind == "solid" else None)
        attrs = [f'd="{_svg_path_d(s, x0, y0)}"', f'fill="{fill}"']
        if fop:
            attrs.append(fop.strip().format("fill"))
        if "fill_rule:evenodd" in s.notes:
            attrs.append('fill-rule="evenodd"')
        if s.stroke is not None:
            sc, sop = _svg_paint(s.stroke.color or Color(rgb="000000"))
            attrs += [f'stroke="{sc}"', f'stroke-width="{s.stroke.width:.3f}"',
                      f'stroke-linecap="{ {"flat": "butt"}.get(s.stroke.cap, s.stroke.cap)}"',
                      f'stroke-linejoin="{s.stroke.join}"']
            if sop:
                attrs.append(sop.strip().format("stroke"))
            if s.stroke.dash_pattern:
                attrs.append('stroke-dasharray="' + " ".join(
                    f"{v:.3f}" for v in s.stroke.dash_pattern) + '"')
        parts.append("<path " + " ".join(attrs) + "/>")
    parts.append("</svg>")
    return "\n".join(parts) + "\n"


def svg_png(shapes: list[Shape], box: tuple[float, float, float, float]) -> bytes:
    """PNG fallback: the paths rasterised with 4x supersampling (RGBA)."""
    import cv2
    from PIL import Image as PILImage

    x0, y0, x1, y1 = box
    k = SVG_PX_PER_PT * 4
    w = max(1, math.ceil((x1 - x0) * k))
    h = max(1, math.ceil((y1 - y0) * k))
    img = np.zeros((h, w, 4), np.uint8)
    for s in shapes:
        polys = [np.round((np.asarray(sub.pts) - [x0, y0]) * k * 16).astype(np.int32)
                 for sub in subpaths(s.path)]
        if s.fill.kind == "solid" and s.fill.color is not None:
            layer = np.zeros((h, w), np.uint8)
            cv2.fillPoly(layer, polys, 255, shift=4)
            _blend(img, layer, s.fill.color)
        if s.stroke is not None:
            layer = np.zeros((h, w), np.uint8)
            subs = subpaths(s.path)
            for poly, sub in zip(polys, subs, strict=True):
                cv2.polylines(layer, [poly], sub.closed, 255,
                              thickness=max(1, round(s.stroke.width * k)), shift=4)
            _blend(img, layer, s.stroke.color or Color(rgb="000000"))
    small = PILImage.fromarray(img, "RGBA").resize(
        (max(1, w // 4), max(1, h // 4)), PILImage.Resampling.BOX)
    buf = io.BytesIO()
    small.save(buf, "PNG", optimize=False)
    return buf.getvalue()


def _blend(img: np.ndarray, layer: np.ndarray, c: Color) -> None:
    rgb = [int((c.rgb or "000000")[i:i + 2], 16) for i in (0, 2, 4)]
    a = (layer.astype(float) / 255.0) * c.alpha
    da = img[..., 3].astype(float) / 255.0
    out_a = a + da * (1 - a)
    for ch in range(3):
        num = rgb[ch] * a + img[..., ch].astype(float) * da * (1 - a)
        img[..., ch] = np.where(out_a > 0, num / np.maximum(out_a, 1e-9), 0).astype(np.uint8)
    img[..., 3] = (out_a * 255).astype(np.uint8)


def svg_clusters(page: Page, work_dir: Path) -> int:
    """Dense groups of unfitted freeforms -> one SVG picture each."""
    free = [e for e in page.elements if isinstance(e, Shape) and e.preset is None and e.path
            and not e.is_background and e.fill.kind in ("none", "solid")]
    if not free:
        return 0
    free.sort(key=lambda e: e.z)
    # union-find over nearby freeforms
    parent = list(range(len(free)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i in range(len(free)):
        a = free[i].bbox
        for j in range(i + 1, len(free)):
            b = free[j].bbox
            if a[0] - SVG_GAP < b[2] and b[0] - SVG_GAP < a[2] and a[1] - SVG_GAP < b[3] and \
                    b[1] - SVG_GAP < a[3]:
                parent[find(j)] = find(i)
    clusters: dict[int, list[Shape]] = defaultdict(list)
    for i, e in enumerate(free):
        clusters[find(i)].append(e)
    allz = sorted(page.elements, key=lambda e: e.z)
    n = 0
    for members in sorted(clusters.values(), key=lambda m: m[0].z):
        segs = sum(len(m.path or []) for m in members)
        if len(members) < SVG_MIN_PATHS and segs < SVG_MIN_SEGMENTS:
            continue
        ids = {m.id for m in members}
        box = (min(m.bbox[0] for m in members), min(m.bbox[1] for m in members),
               max(m.bbox[2] for m in members), max(m.bbox[3] for m in members))
        pad = max((m.stroke.width / 2 if m.stroke else 0.0) for m in members) + 0.5
        box = (box[0] - pad, box[1] - pad, box[2] + pad, box[3] + pad)
        z0, z1 = members[0].z, max(m.z for m in members)
        between = [e for e in allz if z0 < e.z < z1 and e.id not in ids]
        if any(_overlap(e.bbox, box) for e in between):
            continue
        svg = svg_document(members, box)
        png = svg_png(members, box)
        key = hashlib.sha256(svg.encode()).hexdigest()[:16]
        (work_dir / "images").mkdir(parents=True, exist_ok=True)
        svg_rel, png_rel = f"images/svg_{key}.svg", f"images/svg_{key}.png"
        (work_dir / svg_rel).write_text(svg, encoding="utf-8", newline="\n")
        (work_dir / png_rel).write_bytes(png)
        n += 1
        img = Image(id=f"p{page.index}_svg{n}", bbox=_r4(box), z=z0, file=png_rel,
                    svg_file=svg_rel, sha256=hashlib.sha256(png).hexdigest(), source="rule",
                    confidence=0.9, pixel_size=(0, 0),
                    notes=[f"svg:{len(members)} paths, {segs} segments",
                           *(f"replaces:{m.id}" for m in members)])
        page.elements = [e for e in page.elements if e.id not in ids]
        page.elements.append(img)
    return n


# --------------------------------------------------------------------------- page driver


def analyze_page(page: Page, work_dir: Path) -> dict[str, int]:
    stats: dict[str, int] = defaultdict(int)
    if page.triage.fallback_image:
        return dict(stats)
    stats["merged"] = merge_fill_stroke(page.elements)
    if detect_background(page):
        stats["background"] = 1
    stats["gradients"] = gradients(page, work_dir)
    out: list[Any] = []
    split: dict[str, list[str]] = {}
    for e in page.elements:
        if not isinstance(e, Shape) or e.is_background or e.fill.kind == "pattern":
            out.append(e)
            continue
        try:
            res = fit_shape(e)
        except Exception as exc:  # a fit failure keeps the freeform
            LOG.debug("page %d %s: fit failed: %s", page.index, e.id, exc)
            res = None
        if res is None:
            out.append(e)
            stats["freeform"] += 1
        elif isinstance(res, list):
            out.extend(res)
            split[e.id] = [r.id for r in res]
            stats["split"] += 1
            stats["preset"] += len(res)
        else:
            out.append(res)
            stats["connector" if res.connector else "preset"] += 1
    page.elements = out
    stats["heads"] = attach_separate_heads(page.elements)
    stats["attached"] = attach_connectors(page.elements)
    stats["svg"] = svg_clusters(page, work_dir)
    stats["groups"] = make_groups(page, split)
    return dict(stats)
