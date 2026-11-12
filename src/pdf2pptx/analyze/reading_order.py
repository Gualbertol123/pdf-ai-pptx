"""Reading order: recursive XY-cut with column detection.

At each level the region's boxes are projected on both axes and whitespace gaps are
found. A vertical gap is a **column split** when both sides hold a substantial share
of the region's height (two or more boxes each, or at least half the height); a
column split wins over a horizontal cut, so a two-column page is read column by
column even when some paragraph gaps happen to line up across columns. Otherwise the
widest horizontal gap cuts the region into rows (top to bottom), then the widest
vertical gap into columns (left to right). Boxes that overlap with no gap left are
ordered by top, then left.
"""

from __future__ import annotations

from collections.abc import Sequence

BBox = tuple[float, float, float, float]

MIN_GAP = 2.0  # pt


def _gaps(intervals: list[tuple[float, float]]) -> list[tuple[float, float]]:
    """Whitespace gaps between the union of ``intervals`` (sorted, merged)."""
    out: list[tuple[float, float]] = []
    end = None
    for a, b in sorted(intervals):
        if end is not None and a - end >= MIN_GAP:
            out.append((end, a))
        end = b if end is None else max(end, b)
    return out


def _column_split(idx: list[int], boxes: Sequence[BBox]) -> float | None:
    """x of the best column split, or None."""
    gaps = _gaps([(boxes[i][0], boxes[i][2]) for i in idx])
    if not gaps:
        return None
    y0 = min(boxes[i][1] for i in idx)
    y1 = max(boxes[i][3] for i in idx)
    height = max(y1 - y0, 1e-6)
    best = None
    for a, b in gaps:
        cut = (a + b) / 2
        left = [i for i in idx if boxes[i][2] <= cut]
        right = [i for i in idx if boxes[i][0] >= cut]

        def tall(side: list[int]) -> bool:
            h = max(boxes[i][3] for i in side) - min(boxes[i][1] for i in side)
            return len(side) >= 2 or h >= 0.5 * height

        if left and right and tall(left) and tall(right):
            width = b - a
            if best is None or width > best[0]:
                best = (width, cut)
    return best[1] if best else None


def _order(idx: list[int], boxes: Sequence[BBox], depth: int = 0) -> list[int]:
    if len(idx) <= 1 or depth > 64:
        return sorted(idx, key=lambda i: (boxes[i][1], boxes[i][0], i))
    col = _column_split(idx, boxes)
    if col is not None:
        left = [i for i in idx if boxes[i][2] <= col]
        right = [i for i in idx if boxes[i][0] >= col]
        return _order(left, boxes, depth + 1) + _order(right, boxes, depth + 1)
    for axis in (1, 0):  # rows first, then columns
        gaps = _gaps([(boxes[i][axis], boxes[i][axis + 2]) for i in idx])
        if not gaps:
            continue
        a, b = max(gaps, key=lambda g: (g[1] - g[0], -g[0]))
        cut = (a + b) / 2
        first = [i for i in idx if boxes[i][axis + 2] <= cut]
        second = [i for i in idx if boxes[i][axis] >= cut]
        return _order(first, boxes, depth + 1) + _order(second, boxes, depth + 1)
    return sorted(idx, key=lambda i: (boxes[i][1], boxes[i][0], i))


def xy_cut(boxes: Sequence[BBox]) -> list[int]:
    """Indices of ``boxes`` in reading order."""
    return _order(list(range(len(boxes))), boxes)


def rotated_aabb(bbox: BBox, rotation_cw: float) -> BBox:
    """Axis-aligned bounds of a frame rotated about its centre."""
    import math

    if not rotation_cw:
        return bbox
    x0, y0, x1, y1 = bbox
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    hw, hh = (x1 - x0) / 2, (y1 - y0) / 2
    r = math.radians(rotation_cw)
    c, s = abs(math.cos(r)), abs(math.sin(r))
    ex, ey = hw * c + hh * s, hw * s + hh * c
    return cx - ex, cy - ey, cx + ex, cy + ey
