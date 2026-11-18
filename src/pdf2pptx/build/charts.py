"""BUILD (charts, M9): ``Chart`` IR -> a native PowerPoint chart with an embedded
workbook (python-pptx + raw XML for what it does not expose).

Placed to match the PDF: the plot area and legend get manual layouts (fractions of the
chart frame) from ``Chart.plot_bbox`` / ``legend_bbox``; axis bounds, major unit,
gridlines, gap width, colours, markers, fonts and data labels come from the IR.
"""

from __future__ import annotations

import re
from typing import Any

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from ..ir import Chart, Color

CHART_SHAPE_PREFIX = "Chart "  # + the chart id: shape name of a chart's graphic frame

TYPES = {
    "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "column_stacked_100": XL_CHART_TYPE.COLUMN_STACKED_100,
    "bar_clustered": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_stacked": XL_CHART_TYPE.BAR_STACKED,
    "bar_stacked_100": XL_CHART_TYPE.BAR_STACKED_100,
    "line": XL_CHART_TYPE.LINE,
    "line_markers": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "area": XL_CHART_TYPE.AREA,
    "area_stacked": XL_CHART_TYPE.AREA_STACKED,
}
LEGEND = {"right": XL_LEGEND_POSITION.RIGHT, "left": XL_LEGEND_POSITION.LEFT,
          "top": XL_LEGEND_POSITION.TOP, "bottom": XL_LEGEND_POSITION.BOTTOM}
MARKERS = {"diamond": XL_MARKER_STYLE.DIAMOND, "square": XL_MARKER_STYLE.SQUARE,
           "triangle": XL_MARKER_STYLE.TRIANGLE, "circle": XL_MARKER_STYLE.CIRCLE,
           "x": XL_MARKER_STYLE.X, "plus": XL_MARKER_STYLE.PLUS,
           "none": XL_MARKER_STYLE.NONE, "auto": XL_MARKER_STYLE.AUTOMATIC}


def _rgb(c: Color | None) -> RGBColor | None:
    return RGBColor.from_string(c.rgb.upper()) if c is not None and c.rgb else None


def _safe(s: str) -> str:
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)


def _manual_layout(parent: Any, box: tuple[float, float, float, float],
                   frame: tuple[float, float, float, float], inner: bool) -> None:
    """``c:layout/c:manualLayout`` with fractions of the chart frame."""
    fw, fh = frame[2] - frame[0], frame[3] - frame[1]
    if fw <= 0 or fh <= 0:
        return
    old = parent.find(qn("c:layout"))
    if old is not None:
        parent.remove(old)
    lay = etree.SubElement(parent, qn("c:layout"))
    # c:layout must come first in c:plotArea / c:legend (after c:legendPos)
    parent.remove(lay)
    idx = 1 if parent.tag == qn("c:legend") else 0
    parent.insert(idx, lay)
    ml = etree.SubElement(lay, qn("c:manualLayout"))
    if inner:
        etree.SubElement(ml, qn("c:layoutTarget")).set("val", "inner")
    etree.SubElement(ml, qn("c:xMode")).set("val", "edge")
    etree.SubElement(ml, qn("c:yMode")).set("val", "edge")

    def f(v: float) -> str:
        return f"{max(0.0, min(1.0, v)):.5f}"

    etree.SubElement(ml, qn("c:x")).set("val", f((box[0] - frame[0]) / fw))
    etree.SubElement(ml, qn("c:y")).set("val", f((box[1] - frame[1]) / fh))
    etree.SubElement(ml, qn("c:w")).set("val", f((box[2] - box[0]) / fw))
    etree.SubElement(ml, qn("c:h")).set("val", f((box[3] - box[1]) / fh))


def chart_data(ch: Chart) -> CategoryChartData:
    cd = CategoryChartData()
    cd.categories = [_safe(c) for c in ch.categories]
    fmt = str((ch.axes.get("value") or {}).get("format") or "General")
    for s in ch.series:
        n = len(ch.categories)
        vals = (list(s.values) + [None] * n)[:n]
        cd.add_series(_safe(s.name), vals, number_format=fmt)
    return cd


def add_chart(shapes: Any, ch: Chart, x: int, y: int, cx: int, cy: int, scale: float) -> Any:
    """Add ``ch`` as a native chart; returns the graphic frame."""
    xl = TYPES.get(ch.chart_type)
    if xl is None:
        raise ValueError(f"unsupported chart type {ch.chart_type!r}")
    gf = shapes.add_chart(xl, Emu(x), Emu(y), Emu(cx), Emu(cy), chart_data(ch))
    gf.name = CHART_SHAPE_PREFIX + ch.id
    chart = gf.chart
    st = ch.style
    fs = float(st.get("font_size") or 10.0) * scale
    chart.font.size = Pt(fs)
    if st.get("font_family"):
        chart.font.name = str(st["font_family"])
    if st.get("font_color"):
        chart.font.color.rgb = RGBColor.from_string(str(st["font_color"]))
    chart.has_title = False
    cs = chart._chartSpace
    atd = cs.find(".//" + qn("c:autoTitleDeleted"))
    if atd is not None:
        atd.set("val", "1")
    chart.has_legend = ch.has_legend
    if ch.has_legend:
        chart.legend.position = LEGEND.get(ch.legend_pos, XL_LEGEND_POSITION.RIGHT)
        chart.legend.include_in_layout = False
        if st.get("legend_font_size"):
            chart.legend.font.size = Pt(float(st["legend_font_size"]) * scale)
    plot = chart.plots[0]
    is_pie = ch.chart_type in ("pie", "doughnut")
    if "gap_width" in st and hasattr(plot, "gap_width"):
        plot.gap_width = int(float(st["gap_width"] or 150))
    if "overlap" in st and hasattr(plot, "overlap") and not is_pie:
        plot.overlap = int(float(st["overlap"] or 0))
    plot.vary_by_categories = bool(st.get("vary_colors")) or is_pie
    if ch.data_labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_value = True
        lfs = st.get("label_font_size") or st.get("font_size")
        if lfs:
            dl.font.size = Pt(float(lfs) * scale)
        if is_pie:
            from pptx.enum.chart import XL_LABEL_POSITION

            dl.position = XL_LABEL_POSITION.BEST_FIT
    for k, s in enumerate(ch.series):
        ser = plot.series[k]
        line_type = ch.chart_type.startswith("line")
        if s.color is not None and _rgb(s.color) is not None:
            if line_type:
                ser.format.line.color.rgb = _rgb(s.color)
                if s.line_width:
                    ser.format.line.width = Pt(s.line_width * scale)
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = _rgb(s.color)
        if line_type:
            ser.smooth = False
            if s.marker:
                ser.marker.style = MARKERS.get(s.marker, XL_MARKER_STYLE.AUTOMATIC)
                if s.color is not None and s.marker != "none":
                    ser.marker.format.fill.solid()
                    ser.marker.format.fill.fore_color.rgb = _rgb(s.color)
                    ser.marker.format.line.color.rgb = _rgb(s.color)
        for i, pc in enumerate(s.point_colors or []):
            if pc is not None and _rgb(pc) is not None and i < len(ch.categories):
                pt = ser.points[i]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = _rgb(pc)
    if is_pie:
        ang = round(float(st.get("first_slice_angle") or 0.0)) % 360
        pie = cs.find(".//" + qn("c:pieChart"))
        if pie is not None:
            fsa = pie.find(qn("c:firstSliceAng"))
            if fsa is None:
                fsa = etree.SubElement(pie, qn("c:firstSliceAng"))
            fsa.set("val", str(ang))
    else:
        va = chart.value_axis
        ax = ch.axes.get("value") or {}
        if ax.get("min") is not None:
            va.minimum_scale = float(ax["min"])  # type: ignore[arg-type]
        if ax.get("max") is not None:
            va.maximum_scale = float(ax["max"])  # type: ignore[arg-type]
        if ax.get("major"):
            va.major_unit = float(ax["major"])  # type: ignore[arg-type]
        va.has_major_gridlines = st.get("gridlines", "major") == "major"
        if not ax.get("major") and "gridlines" not in st:
            va.has_major_gridlines = True
        va.tick_labels.font.size = Pt(fs)
        if ax.get("format"):
            va.tick_labels.number_format = str(ax["format"])
            va.tick_labels.number_format_is_linked = False
        ca = chart.category_axis
        cfs = st.get("cat_font_size")
        ca.tick_labels.font.size = Pt(float(cfs) * scale if cfs else fs)
        if st.get("hide_tick_labels"):  # the source chart shows no labels
            from pptx.enum.chart import XL_TICK_LABEL_POSITION, XL_TICK_MARK

            va.tick_label_position = XL_TICK_LABEL_POSITION.NONE
            ca.tick_label_position = XL_TICK_LABEL_POSITION.NONE
            for axis in (va, ca):
                axis.major_tick_mark = XL_TICK_MARK.NONE
                axis.format.line.color.rgb = RGBColor(0, 0, 0)
                axis.format.line.width = Pt(1.5 * scale)
    # positions: plot area and legend as in the PDF
    frame = ch.bbox
    pa = cs.find(".//" + qn("c:plotArea"))
    if pa is not None and ch.plot_bbox is not None:
        _manual_layout(pa, ch.plot_bbox, frame, inner=True)
    leg = cs.find(".//" + qn("c:legend"))
    if leg is not None and ch.legend_bbox is not None:
        lb = ch.legend_bbox
        pad = 2.0
        _manual_layout(leg, (lb[0] - pad, lb[1] - pad, lb[2] + pad, lb[3] + pad), frame,
                       inner=False)
    return gf
