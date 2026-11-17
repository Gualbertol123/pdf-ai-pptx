"""Raw DrawingML for what python-pptx cannot express: custom geometry with curves,
alpha, custom and preset dashes, gradients, arrowheads, preset shapes with adjust
values, connectors with connection references, SVG pictures, table cell borders and
fills, slide backgrounds and groups.

Every helper takes EMU / OOXML units already converted by ``build/units.py`` (stroke
widths and dash lengths in points are converted here through ``units`` too).
``tests/test_xml_ext.py`` opens a deck using every helper in PowerPoint (COM) and
checks that it loads without repair and renders.
"""

from __future__ import annotations

from contextvars import ContextVar
from typing import Any
from xml.sax.saxutils import quoteattr

from lxml import etree
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

from ..ir import Color, Fill, Gradient, Stroke
from .units import angle_to_ooxml, pct_to_ooxml, pt_to_emu

NSDECL = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
NS_ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
NO_STYLE_TABLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # "No Style, No Grid"

_CAP = {"flat": "flat", "round": "rnd", "square": "sq"}
_HEADS = {"triangle", "stealth", "diamond", "oval", "arrow"}
_SIZES = {"sm", "med", "lg"}


# M7: while set, colours with ``Color.theme`` are written as theme references
# (``writer.write_pptx`` sets it in editable / balanced mode).
THEME_REFS: ContextVar[bool] = ContextVar("theme_refs", default=False)
SCHEME_NAMES = {"dk1": "tx1", "lt1": "bg1", "dk2": "tx2", "lt2": "bg2"}


def srgb_xml(c: Color) -> str:
    alpha = "" if c.alpha >= 0.9995 else f'<a:alpha val="{pct_to_ooxml(max(c.alpha, 0.0))}"/>'
    if c.theme and THEME_REFS.get():
        return f'<a:schemeClr val="{SCHEME_NAMES.get(c.theme, c.theme)}">{alpha}</a:schemeClr>'
    return f'<a:srgbClr val="{(c.rgb or "000000").upper()}">{alpha}</a:srgbClr>'


def color_xml(c: Color) -> str:
    return f"<a:solidFill>{srgb_xml(c)}</a:solidFill>"


def gradient_xml(g: Gradient) -> str:
    stops = "".join(f'<a:gs pos="{pct_to_ooxml(min(max(s.pos, 0.0), 1.0))}">'
                    f"{srgb_xml(s.color)}</a:gs>" for s in g.stops)
    if g.kind == "linear":
        shade = f'<a:lin ang="{angle_to_ooxml(g.angle % 360.0)}" scaled="0"/>'
    else:
        cx, cy = g.center or (0.5, 0.5)
        l_, t_ = pct_to_ooxml(cx), pct_to_ooxml(cy)
        r_, b_ = pct_to_ooxml(1 - cx), pct_to_ooxml(1 - cy)
        path = "circle" if g.kind == "radial" else ("rect" if g.kind == "rect" else "shape")
        shade = (f'<a:path path="{path}"><a:fillToRect l="{l_}" t="{t_}" r="{r_}" b="{b_}"/>'
                 f"</a:path>")
    return (f'<a:gradFill flip="none" rotWithShape="1"><a:gsLst>{stops}</a:gsLst>{shade}'
            "</a:gradFill>")


def fill_xml(fill: Fill | None, opacity: float = 1.0) -> str:
    if fill is None or fill.kind not in ("solid", "gradient"):
        return "<a:noFill/>"
    if fill.kind == "gradient" and fill.gradient is not None and fill.gradient.stops:
        return gradient_xml(fill.gradient)
    c = fill.color or Color(rgb="000000")
    if opacity < 0.9995:
        c = Color(rgb=c.rgb, theme=c.theme, alpha=c.alpha * opacity)
    return color_xml(c)


def _end(tag: str, kind: str | None, size: str) -> str:
    if not kind or kind not in _HEADS:
        return ""
    w, _, ln = size.partition("-")
    w = w if w in _SIZES else "med"
    ln = ln if ln in _SIZES else "med"
    return f'<a:{tag} type="{kind}" w="{w}" len="{ln}"/>'


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
    elif stroke.dash not in ("solid", "custom", ""):
        dash = f'<a:prstDash val="{stroke.dash}"/>'
    join = {"round": "<a:round/>", "bevel": "<a:bevel/>"}.get(stroke.join,
                                                              '<a:miter lim="800000"/>')
    ends = (_end("headEnd", stroke.head_end, stroke.head_size)
            + _end("tailEnd", stroke.tail_end, stroke.tail_size))
    return (f'<a:ln w="{w}" cap="{_CAP.get(stroke.cap, "flat")}">{fill}{dash}{join}{ends}'
            f"</a:ln>")


def xfrm_xml(off: tuple[int, int], ext: tuple[int, int], rot: float = 0.0, flip_h: bool = False,
             flip_v: bool = False, tag: str = "a:xfrm") -> str:
    attrs = ""
    if rot % 360.0:
        attrs += f' rot="{angle_to_ooxml(rot % 360.0)}"'
    if flip_h:
        attrs += ' flipH="1"'
    if flip_v:
        attrs += ' flipV="1"'
    return (f'<{tag}{attrs}><a:off x="{off[0]}" y="{off[1]}"/>'
            f'<a:ext cx="{max(ext[0], 0)}" cy="{max(ext[1], 0)}"/></{tag}>')


def _empty_text() -> str:
    return ('<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>')


def freeform(shape_id: int, name: str, off: tuple[int, int], ext: tuple[int, int],
             path_cmds: str, fill: Color | Fill | None, stroke: Stroke | None, *,
             stroke_scale: float = 1.0, path_closed_fill: bool = True) -> etree._Element:
    """A ``p:sp`` with custom geometry. ``path_cmds`` is the inner XML of ``a:path``
    in a coordinate space of ``ext`` (EMU)."""
    cx, cy = max(ext[0], 1), max(ext[1], 1)
    if isinstance(fill, Fill):
        fill_x = fill_xml(fill)
        has_fill = fill.kind in ("solid", "gradient")
    else:
        fill_x = color_xml(fill) if fill is not None else "<a:noFill/>"
        has_fill = fill is not None
    path_attrs = f'w="{cx}" h="{cy}"'
    if not has_fill or not path_closed_fill:
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
        f"{fill_x}{line_xml(stroke, stroke_scale)}"
        f"</p:spPr>"
        f"{_empty_text()}"
        f"</p:sp>"
    )
    el: etree._Element = parse_xml(xml)
    return el


def _av_list(adjust: dict[str, float]) -> str:
    gds = "".join(f'<a:gd name="{k}" fmla="val {round(v)}"/>' for k, v in sorted(
        adjust.items()))
    return f"<a:avLst>{gds}</a:avLst>"


def preset_shape(shape_id: int, name: str, off: tuple[int, int], ext: tuple[int, int],
                 preset: str, adjust: dict[str, float], fill: Fill | None,
                 stroke: Stroke | None, *, rot: float = 0.0, flip_h: bool = False,
                 flip_v: bool = False, stroke_scale: float = 1.0, opacity: float = 1.0,
                 locked: bool = False) -> etree._Element:
    """A ``p:sp`` with a preset geometry (``a:prstGeom`` + adjust values)."""
    locks = ('<a:spLocks noGrp="1" noSelect="1" noRot="1" noMove="1" noResize="1"/>'
             if locked else "")
    xml = (
        f"<p:sp {NSDECL}>"
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name={quoteattr(name)}/>'
        f"<p:cNvSpPr>{locks}</p:cNvSpPr><p:nvPr/></p:nvSpPr>"
        f"<p:spPr>{xfrm_xml(off, ext, rot, flip_h, flip_v)}"
        f'<a:prstGeom prst="{preset}">{_av_list(adjust)}</a:prstGeom>'
        f"{fill_xml(fill, opacity)}{line_xml(stroke, stroke_scale)}</p:spPr>"
        f"{_empty_text()}"
        f"</p:sp>"
    )
    el: etree._Element = parse_xml(xml)
    return el


def connector(shape_id: int, name: str, off: tuple[int, int], ext: tuple[int, int],
              preset: str, adjust: dict[str, float], stroke: Stroke | None, *,
              rot: float = 0.0, flip_h: bool = False, flip_v: bool = False,
              start: tuple[int, int] | None = None, end: tuple[int, int] | None = None,
              stroke_scale: float = 1.0) -> etree._Element:
    """A ``p:cxnSp`` (line / bentConnector2 / bentConnector3 ...) with optional
    connection references ``(shape id, site index)``."""
    refs = ""
    if start is not None:
        refs += f'<a:stCxn id="{start[0]}" idx="{start[1]}"/>'
    if end is not None:
        refs += f'<a:endCxn id="{end[0]}" idx="{end[1]}"/>'
    xml = (
        f"<p:cxnSp {NSDECL}>"
        f'<p:nvCxnSpPr><p:cNvPr id="{shape_id}" name={quoteattr(name)}/>'
        f"<p:cNvCxnSpPr>{refs}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>"
        f"<p:spPr>{xfrm_xml(off, ext, rot, flip_h, flip_v)}"
        f'<a:prstGeom prst="{preset}">{_av_list(adjust)}</a:prstGeom>'
        f"<a:noFill/>{line_xml(stroke, stroke_scale)}</p:spPr>"
        f"</p:cxnSp>"
    )
    el: etree._Element = parse_xml(xml)
    return el


def set_connection(cxn: Any, start: tuple[int, int] | None, end: tuple[int, int] | None) -> None:
    """(Re)write the ``stCxn`` / ``endCxn`` references of a ``p:cxnSp``."""
    c_nv = cxn.find(qn("p:nvCxnSpPr")).find(qn("p:cNvCxnSpPr"))
    for child in list(c_nv):
        if child.tag in (qn("a:stCxn"), qn("a:endCxn")):
            c_nv.remove(child)
    for tag, ref in (("a:stCxn", start), ("a:endCxn", end)):
        if ref is not None:
            el = etree.SubElement(c_nv, qn(tag))
            el.set("id", str(ref[0]))
            el.set("idx", str(ref[1]))
    # schema order: stCxn before endCxn, both before extLst (none written)


def svg_blip(blip: Any, svg_rid: str) -> None:
    """Add the Office 2016 SVG extension to a picture's ``a:blip`` (the PNG stays the
    fallback for older readers)."""
    ext_lst = blip.find(qn("a:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(blip, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"))
    ext.set("uri", SVG_EXT_URI)
    svg = etree.SubElement(ext, f"{{{NS_ASVG}}}svgBlip", nsmap={"asvg": NS_ASVG})
    svg.set(qn("r:embed"), svg_rid)


def add_svg_part(slide_part: Any, svg_bytes: bytes) -> str:
    """Store an SVG image part and relate it to the slide; returns the rId."""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI

    package = slide_part.package
    cache = getattr(package, "_pdf2pptx_svgs", None)
    if cache is None:
        cache = {}
        package._pdf2pptx_svgs = cache
    part = cache.get(svg_bytes)
    if part is None:
        partname = package.next_partname("/ppt/media/image%d.svg")
        part = Part(PackURI(str(partname)), "image/svg+xml", package, svg_bytes)
        cache[svg_bytes] = part
    rid: str = slide_part.relate_to(part, RT.IMAGE)
    return rid


def cell_props(tc: Any, *, margins: tuple[int, int, int, int], anchor: str,
               borders: dict[str, Stroke | None], fill: Fill | None,
               stroke_scale: float = 1.0) -> None:
    """``a:tcPr`` of a table cell: margins (EMU), anchor, the four borders (None = no
    line) and the fill. Children are written in schema order."""
    tcpr = tc.find(qn("a:tcPr"))
    if tcpr is None:
        tcpr = etree.SubElement(tc, qn("a:tcPr"))
    for child in list(tcpr):
        tcpr.remove(child)
    ml, mt, mr, mb = margins
    tcpr.set("marL", str(max(ml, 0)))
    tcpr.set("marR", str(max(mr, 0)))
    tcpr.set("marT", str(max(mt, 0)))
    tcpr.set("marB", str(max(mb, 0)))
    tcpr.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}.get(anchor, "t"))
    for side, tag in (("left", "a:lnL"), ("right", "a:lnR"), ("top", "a:lnT"),
                      ("bottom", "a:lnB")):
        ln = frag(line_xml(borders.get(side), stroke_scale))
        ln.tag = qn(tag)
        tcpr.append(ln)
    tcpr.append(frag(fill_xml(fill)))


def frag(xml: str) -> Any:
    """Parse a DrawingML fragment written without namespace declarations."""
    wrapper = parse_xml(f"<a:wrap {NSDECL}>{xml}</a:wrap>")
    return wrapper[0]


def plain_table(tbl: Any) -> None:
    """No table style effects: the "No Style, No Grid" style, no banding flags."""
    tbl_pr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow", "lastRow", "firstCol", "lastCol", "bandCol"):
        if tbl_pr.get(attr) is not None:
            del tbl_pr.attrib[attr]
    sid = tbl_pr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tbl_pr, qn("a:tableStyleId"))
    sid.text = NO_STYLE_TABLE


def background(slide: Any, fill: Fill | None = None, image_rid: str | None = None) -> None:
    """``p:bg`` with a solid / gradient fill or a stretched picture."""
    c_sld = slide._element.find(qn("p:cSld"))
    old = c_sld.find(qn("p:bg"))
    if old is not None:
        c_sld.remove(old)
    if image_rid is not None:
        body = (f'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{image_rid}"/>'
                f"<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>")
    else:
        body = fill_xml(fill)
    bg = parse_xml(f"<p:bg {NSDECL}><p:bgPr>{body}<a:effectLst/></p:bgPr></p:bg>")
    c_sld.insert(0, bg)


def lock_picture(pic: Any) -> None:
    """Lock a picture (a background that could not be the slide background)."""
    c_nv = pic.find(qn("p:nvPicPr")).find(qn("p:cNvPicPr"))
    locks = c_nv.find(qn("a:picLocks"))
    if locks is None:
        locks = etree.SubElement(c_nv, qn("a:picLocks"))
    for k in ("noGrp", "noSelect", "noRot", "noMove", "noResize", "noChangeAspect"):
        locks.set(k, "1")


def pt(x: int, y: int) -> str:
    return f'<a:pt x="{x}" y="{y}"/>'
