"""DrawingML preset geometry: outlines and connection sites of the built-in shapes.

The formulas follow ECMA-376 ``presetShapeDefinitions.xml`` (adjust values in
1/100000, ``ss`` = the shorter side). Outlines are point lists in the shape's own
frame (origin top-left, y down, before flips and rotation); arcs are sampled.
``tests/test_shapes.py`` checks every outline against PowerPoint's own PDF export.
"""

from __future__ import annotations

import math
from collections.abc import Callable

Pt = tuple[float, float]

ARC_STEPS = 24  # samples per quarter arc

# adjust names, defaults and which side bounds them (for pinning)
DEFAULTS: dict[str, dict[str, float]] = {
    "rect": {},
    "roundRect": {"adj": 16667},
    "ellipse": {},
    "triangle": {"adj": 50000},
    "rtTriangle": {},
    "diamond": {},
    "parallelogram": {"adj": 25000},
    "trapezoid": {"adj": 25000},
    "chevron": {"adj": 50000},
    "homePlate": {"adj": 50000},
    "rightArrow": {"adj1": 50000, "adj2": 50000},
    "leftArrow": {"adj1": 50000, "adj2": 50000},
    "upArrow": {"adj1": 50000, "adj2": 50000},
    "downArrow": {"adj1": 50000, "adj2": 50000},
    "pie": {"adj1": 0, "adj2": 16200000},
}

# number of polygon vertices (None: has curves)
VERTICES: dict[str, int | None] = {
    "rect": 4, "roundRect": None, "ellipse": None, "triangle": 3, "rtTriangle": 3,
    "diamond": 4, "parallelogram": 4, "trapezoid": 4, "chevron": 6, "homePlate": 5,
    "rightArrow": 7, "leftArrow": 7, "upArrow": 7, "downArrow": 7, "pie": None,
}


def _pin(lo: float, v: float, hi: float) -> float:
    return max(lo, min(hi, v))


def _arc(cx: float, cy: float, rx: float, ry: float, a0: float, sweep: float,
         steps: int | None = None) -> list[Pt]:
    """Points on an elliptical arc; angles in degrees, clockwise from +x (y down)."""
    n = steps or max(2, math.ceil(abs(sweep) / 90.0 * ARC_STEPS))
    out = []
    for k in range(n + 1):
        a = math.radians(a0 + sweep * k / n)
        out.append((cx + rx * math.cos(a), cy + ry * math.sin(a)))
    return out


def _round_rect(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    r = ss * _pin(0, a["adj"], 50000) / 100000
    if r <= 1e-6:
        return [(0, 0), (w, 0), (w, h), (0, h)]
    pts: list[Pt] = []
    pts += _arc(r, r, r, r, 180, 90)
    pts += _arc(w - r, r, r, r, 270, 90)
    pts += _arc(w - r, h - r, r, r, 0, 90)
    pts += _arc(r, h - r, r, r, 90, 90)
    return pts


def _ellipse(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    return _arc(w / 2, h / 2, w / 2, h / 2, 0, 360, 4 * ARC_STEPS)[:-1]


def _triangle(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    x2 = w * _pin(0, a["adj"], 100000) / 100000
    return [(0, h), (x2, 0), (w, h)]


def _rt_triangle(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    return [(0, 0), (w, h), (0, h)]


def _diamond(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    return [(0, h / 2), (w / 2, 0), (w, h / 2), (w / 2, h)]


def _parallelogram(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    x2 = ss * _pin(0, a["adj"], 100000 * w / ss) / 100000
    return [(0, h), (x2, 0), (w, 0), (w - x2, h)]


def _trapezoid(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    x2 = ss * _pin(0, a["adj"], 50000 * w / ss) / 100000
    return [(0, h), (x2, 0), (w - x2, 0), (w, h)]


def _chevron(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    x1 = ss * _pin(0, a["adj"], 100000 * w / ss) / 100000
    return [(0, 0), (w - x1, 0), (w, h / 2), (w - x1, h), (0, h), (x1, h / 2)]


def _home_plate(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    x1 = w - ss * _pin(0, a["adj"], 100000 * w / ss) / 100000
    return [(0, 0), (x1, 0), (w, h / 2), (x1, h), (0, h)]


def _right_arrow(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    a1 = _pin(0, a["adj1"], 100000)
    a2 = _pin(0, a["adj2"], 100000 * w / ss)
    x1 = w - ss * a2 / 100000
    dy1 = h * a1 / 200000
    y1, y2 = h / 2 - dy1, h / 2 + dy1
    return [(0, y1), (x1, y1), (x1, 0), (w, h / 2), (x1, h), (x1, y2), (0, y2)]


def _left_arrow(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    a1 = _pin(0, a["adj1"], 100000)
    a2 = _pin(0, a["adj2"], 100000 * w / ss)
    x2 = ss * a2 / 100000
    dy1 = h * a1 / 200000
    y1, y2 = h / 2 - dy1, h / 2 + dy1
    return [(0, h / 2), (x2, 0), (x2, y1), (w, y1), (w, y2), (x2, y2), (x2, h)]


def _up_arrow(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    a1 = _pin(0, a["adj1"], 100000)
    a2 = _pin(0, a["adj2"], 100000 * h / ss)
    y2 = ss * a2 / 100000
    dx1 = w * a1 / 200000
    x1, x2 = w / 2 - dx1, w / 2 + dx1
    return [(0, y2), (w / 2, 0), (w, y2), (x2, y2), (x2, h), (x1, h), (x1, y2)]


def _down_arrow(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    ss = min(w, h)
    a1 = _pin(0, a["adj1"], 100000)
    a2 = _pin(0, a["adj2"], 100000 * h / ss)
    y1 = h - ss * a2 / 100000
    dx1 = w * a1 / 200000
    x1, x2 = w / 2 - dx1, w / 2 + dx1
    return [(0, y1), (x1, y1), (x1, 0), (x2, 0), (x2, y1), (w, y1), (w / 2, h)]


def _pie(w: float, h: float, a: dict[str, float]) -> list[Pt]:
    """Visual angles (as PowerPoint draws them) on the frame's ellipse."""
    st = a["adj1"] / 60000.0
    en = a["adj2"] / 60000.0
    sw = (en - st) % 360.0
    if sw <= 1e-6:
        sw = 360.0
    cx, cy, rx, ry = w / 2, h / 2, w / 2, h / 2

    def param(ang: float) -> float:  # visual angle -> ellipse parameter
        r = math.radians(ang)
        return math.degrees(math.atan2(math.sin(r) * rx, math.cos(r) * ry))

    p0 = param(st)
    p1 = param(st + sw)
    swp = (p1 - p0) % 360.0 or 360.0
    return [*_arc(cx, cy, rx, ry, p0, swp), (cx, cy)]


OUTLINES: dict[str, Callable[[float, float, dict[str, float]], list[Pt]]] = {
    "rect": lambda w, h, a: [(0, 0), (w, 0), (w, h), (0, h)],
    "roundRect": _round_rect,
    "ellipse": _ellipse,
    "triangle": _triangle,
    "rtTriangle": _rt_triangle,
    "diamond": _diamond,
    "parallelogram": _parallelogram,
    "trapezoid": _trapezoid,
    "chevron": _chevron,
    "homePlate": _home_plate,
    "rightArrow": _right_arrow,
    "leftArrow": _left_arrow,
    "upArrow": _up_arrow,
    "downArrow": _down_arrow,
    "pie": _pie,
}


def outline(preset: str, w: float, h: float, adjust: dict[str, float] | None = None) -> list[Pt]:
    """The preset's closed outline in its own frame (w x h)."""
    a = {**DEFAULTS[preset], **(adjust or {})}
    return OUTLINES[preset](max(w, 1e-6), max(h, 1e-6), a)


def place(pts: list[Pt], frame: tuple[float, float, float, float], rotation: float = 0.0,
          flip_h: bool = False, flip_v: bool = False) -> list[Pt]:
    """Frame-local points -> page points (flips, then clockwise rotation about the centre)."""
    x0, y0, x1, y1 = frame
    w, h = x1 - x0, y1 - y0
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    c, s = math.cos(math.radians(rotation)), math.sin(math.radians(rotation))
    out = []
    for px, py in pts:
        if flip_h:
            px = w - px
        if flip_v:
            py = h - py
        dx, dy = px - w / 2, py - h / 2
        out.append((cx + dx * c - dy * s, cy + dx * s + dy * c))
    return out


def connection_sites(preset: str, w: float, h: float) -> list[Pt]:
    """Connection sites (frame-local), indexed as PowerPoint's ``stCxn/endCxn idx``."""
    if preset in ("rect", "roundRect", "diamond"):
        return [(w / 2, 0), (0, h / 2), (w / 2, h), (w, h / 2)]
    if preset == "ellipse":
        c = math.cos(math.radians(45))
        il, ir = w / 2 - w / 2 * c, w / 2 + w / 2 * c
        it, ib = h / 2 - h / 2 * c, h / 2 + h / 2 * c
        return [(w / 2, 0), (il, it), (0, h / 2), (il, ib), (w / 2, h), (ir, ib), (w, h / 2),
                (ir, it)]
    return []
