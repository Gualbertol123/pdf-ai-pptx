"""Affine matrices in PDF convention: ``(a, b, c, d, e, f)`` maps
``x' = a*x + c*y + e``, ``y' = b*x + d*y + f``.
"""

from __future__ import annotations

import math

Matrix = tuple[float, float, float, float, float, float]

IDENTITY: Matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def concat(first: Matrix, then: Matrix) -> Matrix:
    """The matrix that applies ``first`` and then ``then``."""
    a1, b1, c1, d1, e1, f1 = first
    a2, b2, c2, d2, e2, f2 = then
    return (
        a1 * a2 + b1 * c2,
        a1 * b2 + b1 * d2,
        c1 * a2 + d1 * c2,
        c1 * b2 + d1 * d2,
        e1 * a2 + f1 * c2 + e2,
        e1 * b2 + f1 * d2 + f2,
    )


def apply(m: Matrix, x: float, y: float) -> tuple[float, float]:
    a, b, c, d, e, f = m
    return a * x + c * y + e, b * x + d * y + f


def scale_of(m: Matrix) -> float:
    """Mean linear scale (sqrt of |det|), used for line widths and dash lengths."""
    return math.sqrt(abs(m[0] * m[3] - m[1] * m[2]))


def page_matrix(cropbox: tuple[float, float, float, float], rotation: int) -> Matrix:
    """PDF user space -> display space (points, origin top-left, y down, /Rotate applied).

    ``rotation`` is the page's /Rotate in degrees, clockwise when displayed.
    """
    x0, y0, x1, y1 = cropbox
    r = rotation % 360
    if r == 90:
        return (0.0, 1.0, 1.0, 0.0, -y0, -x0)
    if r == 180:
        return (-1.0, 0.0, 0.0, 1.0, x1, -y0)
    if r == 270:
        return (0.0, -1.0, -1.0, 0.0, y1, x1)
    return (1.0, 0.0, 0.0, -1.0, -x0, y1)


def bbox_of(points: list[tuple[float, float]]) -> tuple[float, float, float, float]:
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return (min(xs), min(ys), max(xs), max(ys))


def transform_rect(m: Matrix, left: float, b: float, r: float, t: float) -> tuple[
        float, float, float, float]:
    return bbox_of([apply(m, left, b), apply(m, r, b), apply(m, left, t), apply(m, r, t)])


def norm_angle(deg: float) -> float:
    """Normalise to (-180, 180]."""
    a = math.fmod(deg, 360.0)
    if a <= -180.0:
        a += 360.0
    elif a > 180.0:
        a -= 360.0
    return 0.0 if abs(a) < 1e-9 else a
