"""Unit conversions. The only place where EMU appears (PLAN.md §8 rule 5)."""

from __future__ import annotations

EMU_PER_INCH = 914_400
PT_PER_INCH = 72
EMU_PER_PT = 12_700  # 914400 / 72
EMU_PER_PX_96 = 9_525  # one CSS pixel at 96 DPI


def pt_to_emu(pt: float) -> int:
    """Points to EMU, rounded half away from zero to the nearest integer EMU."""
    v = pt * EMU_PER_PT
    return int(v + 0.5) if v >= 0 else -int(-v + 0.5)


def emu_to_pt(emu: int | float) -> float:
    return float(emu) / EMU_PER_PT


def inch_to_pt(inch: float) -> float:
    return inch * PT_PER_INCH


def pt_to_inch(pt: float) -> float:
    return pt / PT_PER_INCH


def pt_to_px(pt: float, dpi: float) -> float:
    return pt * dpi / PT_PER_INCH


def px_to_pt(px: float, dpi: float) -> float:
    return px * PT_PER_INCH / dpi


def pt_to_centipoints(pt: float) -> int:
    """Font sizes and character spacing in DrawingML are in 1/100 pt."""
    return round(pt * 100)


def angle_to_ooxml(deg: float) -> int:
    """DrawingML angles are in 1/60000 degree."""
    return round(deg * 60_000)


def pct_to_ooxml(fraction: float) -> int:
    """DrawingML percentages are in 1/1000 percent (100% = 100000)."""
    return round(fraction * 100_000)


def _half_up(v: float) -> int:
    return int(v + 0.5) if v >= 0 else -int(-v + 0.5)


def ppt_spacing(points: float, single: float) -> tuple[str, int, float]:
    """Best PowerPoint encoding of a line / paragraph spacing given in points.

    Measured in M4: PowerPoint rounds ``spcPts`` to whole points and ``spcPct``
    (a percentage of the single line height ``single``) to whole percent, both half
    up. Returns ``(tag, ooxml value, effective points)`` for whichever encoding lands
    closer to ``points``.
    """
    pts = float(_half_up(points))
    pct = _half_up(points / single * 100.0) if single > 0 else 0
    pct_pts = pct / 100.0 * single
    if single > 0 and abs(pct_pts - points) < abs(pts - points):
        return "spcPct", pct * 1000, pct_pts
    return "spcPts", int(pts) * 100, pts
