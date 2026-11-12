"""Raw DrawingML for what python-pptx cannot express (custom geometry with curves,
alpha, custom dashes). Every helper takes EMU / OOXML units already converted by
``build/units.py``.
"""

from __future__ import annotations

from xml.sax.saxutils import quoteattr

from lxml import etree
from pptx.oxml import parse_xml

from ..ir import Color, Stroke
from .units import pct_to_ooxml, pt_to_emu

NSDECL = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')

_CAP = {"flat": "flat", "round": "rnd", "square": "sq"}


def color_xml(c: Color) -> str:
    alpha = "" if c.alpha >= 0.9995 else f'<a:alpha val="{pct_to_ooxml(max(c.alpha, 0.0))}"/>'
    return f'<a:solidFill><a:srgbClr val="{c.rgb or "000000"}">{alpha}</a:srgbClr></a:solidFill>'


def line_xml(stroke: Stroke | None, scale: float = 1.0) -> str:
    if stroke is None:
        return "<a:ln><a:noFill/></a:ln>"
    width_pt = max(stroke.width * scale, 0.25)
    w = pt_to_emu(width_pt)
    fill = color_xml(stroke.color or Color(rgb="000000"))
    dash = ""
    if stroke.dash == "custom" and stroke.dash_pattern:
        pat = list(stroke.dash_pattern)
        if len(pat) % 2:
            pat = pat * 2
        parts = "".join(
            f'<a:ds d="{pct_to_ooxml(max(d * scale, 0.01) / width_pt)}" '
            f'sp="{pct_to_ooxml(max(s * scale, 0.01) / width_pt)}"/>'
            for d, s in zip(pat[::2], pat[1::2], strict=True))
        dash = f"<a:custDash>{parts}</a:custDash>"
    join = {"round": "<a:round/>", "bevel": "<a:bevel/>"}.get(stroke.join,
                                                              '<a:miter lim="800000"/>')
    return f'<a:ln w="{w}" cap="{_CAP.get(stroke.cap, "flat")}">{fill}{dash}{join}</a:ln>'


def freeform(shape_id: int, name: str, off: tuple[int, int], ext: tuple[int, int],
             path_cmds: str, fill: Color | None, stroke: Stroke | None, *,
             stroke_scale: float = 1.0, path_closed_fill: bool = True) -> etree._Element:
    """A ``p:sp`` with custom geometry. ``path_cmds`` is the inner XML of ``a:path``
    in a coordinate space of ``ext`` (EMU)."""
    cx, cy = max(ext[0], 1), max(ext[1], 1)
    fill_xml = color_xml(fill) if fill is not None else "<a:noFill/>"
    path_attrs = f'w="{cx}" h="{cy}"'
    if fill is None or not path_closed_fill:
        path_attrs += ' fill="none"'
    if stroke is None:
        path_attrs += ' stroke="0"'
    xml = (
        f"<p:sp {NSDECL}>"
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name={quoteattr(name)}/><p:cNvSpPr/><p:nvPr/>'
        f"</p:nvSpPr>"
        f"<p:spPr>"
        f'<a:xfrm><a:off x="{off[0]}" y="{off[1]}"/><a:ext cx="{ext[0]}" cy="{ext[1]}"/></a:xfrm>'
        f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        f'<a:rect l="0" t="0" r="r" b="b"/>'
        f"<a:pathLst><a:path {path_attrs}>{path_cmds}</a:path></a:pathLst></a:custGeom>"
        f"{fill_xml}{line_xml(stroke, stroke_scale)}"
        f"</p:spPr>"
        f'<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
        f'<a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>'
        f"</p:sp>"
    )
    el: etree._Element = parse_xml(xml)
    return el


def pt(x: int, y: int) -> str:
    return f'<a:pt x="{x}" y="{y}"/>'
