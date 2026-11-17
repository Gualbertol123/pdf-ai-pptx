"""BUILD: IR -> PPTX (baseline writer, M2).

* Slide size from the page size (the most common page size; other pages are
  scaled uniformly and centred).
* Text (M4): one wrapping, zero-margin text box per TextBlock with paragraph
  properties (alignment, indents, line spacing, space after, level) and real
  PowerPoint bullets (character, autonumber, picture). Autofit is off except in
  ``editable`` mode (shrink on overflow). ``TextOptions.layout="lines"`` (or a page
  without TextBlocks) writes the M2 baseline instead: one non-wrapping box per line
  (``build/lines.py``).
* Images placed with crop, rotation and flips; every vector path as a freeform.
* z-order kept: elements are written in IR ``z`` order.
* A page that fails to build (or failed to parse) becomes its 300-DPI render as a
  full-slide picture, and is marked ``triage.fallback_image`` in the IR.
* Output is byte-deterministic: zip entries are rewritten with a fixed timestamp.
"""

from __future__ import annotations

import dataclasses
import datetime
import hashlib
import io
import zipfile
from collections import Counter
from contextvars import ContextVar
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.shapes.shapetree import SlideShapeFactory, SlideShapes
from pptx.util import Emu

from .. import log
from ..analyze.builds import absorbed_pages, element_step, page_to_slide, paragraph_steps
from ..ingest.render import ensure_render
from ..ir import (
    Bullet,
    Color,
    Document,
    Fill,
    Group,
    Image,
    Page,
    Paragraph,
    Shape,
    Table,
    TextBlock,
    TextRun,
)
from . import structure, xml_ext
from .lines import line_blocks
from .units import pct_to_ooxml, ppt_spacing, pt_to_centipoints, pt_to_emu

TEXT_LAYOUTS = ("blocks", "lines")
SINGLE_LINE = 1.2  # PowerPoint single line height / font size (measured, all fonts)
BULLET_IMAGE_NOTE = "bullet_image"  # analyze/text.py marks images drawn as bullets
TEXT_SHAPE_PREFIX = "TextBlock "  # + the block id: shape name of a block's text box
LINE_BREAK = "\v"  # in TextRun.text: a forced line break (a:br), set by the tuner
_ALIGN = {"left": "l", "center": "ctr", "right": "r", "justify": "just",
          "distributed": "dist"}

LOG = log.get(__name__)

MIN_SLIDE_PT, MAX_SLIDE_PT = 72.0, 4032.0  # PowerPoint limits: 1 in .. 56 in
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
FIXED_CORE_TIME = datetime.datetime(2000, 1, 1)
BLANK_LAYOUT = 6


@dataclass
class _Xf:
    """Page points -> slide points: uniform scale + offset."""

    s: float = 1.0
    dx: float = 0.0
    dy: float = 0.0

    def x(self, v: float) -> int:
        return pt_to_emu(v * self.s + self.dx)

    def y(self, v: float) -> int:
        return pt_to_emu(v * self.s + self.dy)

    def d(self, v: float) -> int:
        return pt_to_emu(v * self.s)


@dataclass
class TextOptions:
    """What the writer may do, from the mode (text, and since M7 the deck structure)."""

    layout: str = "blocks"  # blocks | lines (the M2 baseline, for comparison)
    autofit: str = "none"  # none | shrink
    keep_hyphens: bool = True  # False: drop end-of-line hyphenation (editable)
    # M7. structure: slide master / layouts with the repeated elements, slide-number
    # placeholders. placeholders: title / body text in placeholders. theme_refs: colours
    # and fonts as theme references. builds: build-up pages collapsed into animations.
    structure: bool = True
    placeholders: bool = False
    theme_refs: bool = False
    builds: bool = False
    slide_no: int | None = None  # set per slide by the writer (slide-number fields)

    @classmethod
    def for_mode(cls, mode: str, layout: str = "blocks", *, builds: bool | None = None,
                 structure: bool = True) -> TextOptions:
        from ..config import MODE_LIMITS

        editable = mode in ("editable", "deck")
        lim = MODE_LIMITS[mode]
        structured = structure and layout == "blocks"
        return cls(layout=layout, autofit="shrink" if editable else "none",
                   keep_hyphens=not editable, structure=structured,
                   placeholders=structured and lim.use_placeholders,
                   theme_refs=structured and lim.use_theme,
                   builds=structured and (lim.builds if builds is None else builds))

    def flat(self) -> TextOptions:
        """One slide per page, everything on the slides (the tuner's deck, M5/M6 output)."""
        return dataclasses.replace(self, structure=False, placeholders=False,
                                   theme_refs=False, builds=False)


@dataclass
class WriteReport:
    slides: int = 0
    fallback_pages: list[int] = field(default_factory=list)
    counts: Counter[str] = field(default_factory=Counter)
    errors: list[str] = field(default_factory=list)


def slide_size(doc: Document) -> tuple[float, float, float]:
    """(width pt, height pt, global scale) for the deck."""
    if not doc.pages:
        return 960.0, 540.0, 1.0
    sizes = Counter((p.width, p.height) for p in doc.pages if p.width > 0 and p.height > 0)
    w, h = min(sizes.items(), key=lambda kv: (-kv[1], kv[0]))[0] if sizes else (960.0, 540.0)
    g = 1.0
    if max(w, h) > MAX_SLIDE_PT:
        g = MAX_SLIDE_PT / max(w, h)
    if min(w, h) * g < MIN_SLIDE_PT:
        g = MIN_SLIDE_PT / min(w, h)
    return w * g, h * g, g


def _xf_for(page: Page, sw: float, sh: float) -> _Xf:
    if page.width <= 0 or page.height <= 0:
        return _Xf()
    s = min(sw / page.width, sh / page.height)
    return _Xf(s=s, dx=(sw - page.width * s) / 2, dy=(sh - page.height * s) / 2)


_BAD_XML = {c: None for c in range(0x20) if c not in (0x09, 0x0A, 0x0D)}


def _xml_safe(text: str) -> str:
    text = text.translate(_BAD_XML)
    return "".join(c for c in text if not 0xD800 <= ord(c) <= 0xDFFF and c not in "￾￿")


def _rgb(c: Color | None) -> RGBColor:
    return RGBColor.from_string((c.rgb if c and c.rgb else "000000").upper())


# --------------------------------------------------------------------------- elements


def _el(parent: Any, tag: str, **attrs: str) -> Any:
    el = etree.SubElement(parent, qn(tag))
    for k, v in attrs.items():
        el.set(k, v)
    return el


def _srgb(parent: Any, c: Color) -> None:
    if c.theme and xml_ext.THEME_REFS.get():
        _el(parent, "a:schemeClr", val=xml_ext.SCHEME_NAMES.get(c.theme, c.theme))
        return
    _el(parent, "a:srgbClr", val=(c.rgb or "000000").upper())


# M7: (major, minor) theme fonts while theme references are on (set by write_pptx)
THEME_FONTS: ContextVar[tuple[str | None, str | None]] = ContextVar("theme_fonts",
                                                                     default=(None, None))


def _typeface(font: str) -> str:
    if xml_ext.THEME_REFS.get():
        major, minor = THEME_FONTS.get()
        if font == minor:
            return "+mn-lt"
        if font == major:
            return "+mj-lt"
    return font


def _para_props(p: Any, para: Paragraph, xf: _Xf, opts: TextOptions, slide: Any,
                work_dir: Path, pictures: dict[str, Image]) -> None:
    """``a:pPr`` attributes and children, in schema order."""
    ppr = p._p.get_or_add_pPr()
    if para.alignment in _ALIGN and para.alignment != "left":
        ppr.set("algn", _ALIGN[para.alignment])
    if para.level:
        ppr.set("lvl", str(min(para.level, 8)))
    if para.indent_left or para.indent_first or para.bullet.kind != "none" or para.level:
        ppr.set("marL", str(xf.d(para.indent_left)))
        ppr.set("indent", str(xf.d(para.indent_first)))
    # Spacing in points goes out as whole points or whole percent of single spacing
    # (1.2 x the paragraph's font size), whichever PowerPoint rounds closer (M4).
    single = SINGLE_LINE * xf.s * max((r.size for r in para.runs if not r.baseline_shift),
                                      default=12.0)

    def spacing(tag: str, pt: float) -> None:
        kind, val, _ = ppt_spacing(pt * xf.s, single)
        _el(_el(ppr, tag), f"a:{kind}", val=str(val))

    if para.line_spacing_rule == "exact" and para.line_spacing:
        spacing("a:lnSpc", para.line_spacing)
    elif para.line_spacing_rule == "multiple" and para.line_spacing:
        _el(_el(ppr, "a:lnSpc"), "a:spcPct", val=str(pct_to_ooxml(para.line_spacing)))
    if para.space_before:
        spacing("a:spcBef", para.space_before)
    if para.space_after:
        spacing("a:spcAft", para.space_after)
    _bullet_props(ppr, para.bullet, slide, work_dir, pictures)


def _bullet_props(ppr: Any, b: Bullet, slide: Any, work_dir: Path,
                  pictures: dict[str, Image]) -> None:
    if b.kind == "none":
        if ppr.get("marL") is not None:  # an indented plain paragraph: no inherited bullet
            _el(ppr, "a:buNone")
        return
    if b.kind == "picture":
        img = pictures.get(b.image_id or "")
        if img is not None:
            _el(ppr, "a:buSzPct", val=str(pct_to_ooxml((b.size_pct or 100.0) / 100.0)))
            _, rid = slide.part.get_or_add_image_part(str(work_dir / img.file))
            blip = _el(_el(ppr, "a:buBlip"), "a:blip")
            blip.set(qn("r:embed"), rid)
            return
        b = Bullet(kind="char", char="•")
    if b.color is not None and b.color.rgb:
        _srgb(_el(ppr, "a:buClr"), b.color)
    if b.size_pct:
        _el(ppr, "a:buSzPct", val=str(pct_to_ooxml(b.size_pct / 100.0)))
    if b.kind == "char":
        if b.font:
            _el(ppr, "a:buFont", typeface=b.font)
        _el(ppr, "a:buChar", char=_xml_safe(b.char or "•") or "•")
    elif b.kind == "autonum":
        an = _el(ppr, "a:buAutoNum", type=b.scheme or "arabicPeriod")
        if b.start_at and b.start_at != 1:
            an.set("startAt", str(b.start_at))


def _run_props(rpr: Any, run: TextRun, xf: _Xf) -> None:
    rpr.set("lang", rpr.get("lang") or "en-US")
    rpr.set("sz", str(pt_to_centipoints(run.size * xf.s)))
    if run.bold:
        rpr.set("b", "1")
    if run.italic:
        rpr.set("i", "1")
    if run.underline:
        rpr.set("u", "sng")
    if run.strike:
        rpr.set("strike", "sngStrike")
    if run.baseline_shift:
        rpr.set("baseline", str(round(run.baseline_shift * 1000)))
    if run.char_spacing:
        rpr.set("spc", str(pt_to_centipoints(run.char_spacing * xf.s)))
    _srgb(_el(rpr, "a:solidFill"), run.color)
    if run.font:
        _el(rpr, "a:latin", typeface=_typeface(run.font))
        _el(rpr, "a:cs", typeface=run.font)


def _field(p: Any, run: TextRun, text: str, xf: _Xf, key: str) -> None:
    g = hashlib.sha256(key.encode()).hexdigest().upper()
    fld = etree.SubElement(p._p, qn("a:fld"))
    fld.set("id", f"{{{g[:8]}-{g[8:12]}-{g[12:16]}-{g[16:20]}-{g[20:32]}}}")
    fld.set("type", run.field_type or "slidenum")
    _run_props(etree.SubElement(fld, qn("a:rPr")), run, xf)
    etree.SubElement(fld, qn("a:t")).text = text
    end = p._p.find(qn("a:endParaRPr"))
    if end is not None:  # endParaRPr stays last
        p._p.remove(end)
        p._p.append(end)


def _drop_hyphens(para: Paragraph) -> list[TextRun]:
    """Runs with end-of-line hyphenation removed (``hyphen:<run>:<offset>`` notes)."""
    cuts = sorted(((int(n.split(":")[1]), int(n.split(":")[2])) for n in para.notes
                   if n.startswith("hyphen:")), reverse=True)
    runs = [dataclasses.replace(r) for r in para.runs]
    for ri, off in cuts:
        if 0 <= ri < len(runs) and runs[ri].text[off:off + 1] in ("-", "­"):
            runs[ri].text = runs[ri].text[:off] + runs[ri].text[off + 1:]
    return runs


def _placeholder_kind(tb: TextBlock, opts: TextOptions) -> str | None:
    """The placeholder ``tb`` is written into (M7), or None for a text box."""
    if not opts.structure or tb.placeholder not in structure.PH:
        return None
    if tb.placeholder == "sldNum" or opts.placeholders:
        return tb.placeholder
    return None


def _add_text(slide: Any, tb: TextBlock, xf: _Xf, internal: list[tuple[Any, int]],
              opts: TextOptions | None = None, work_dir: Path | None = None,
              pictures: dict[str, Image] | None = None, shapes: Any = None) -> Any:
    opts = opts or TextOptions()
    shape = _text_box(slide, tb, xf, internal, opts, work_dir, pictures, shapes)
    kind = _placeholder_kind(tb, opts)
    if kind is not None:
        shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        structure.to_placeholder(shape._element, kind)
    return shape


def _text_box(slide: Any, tb: TextBlock, xf: _Xf, internal: list[tuple[Any, int]],
              opts: TextOptions, work_dir: Path | None, pictures: dict[str, Image] | None,
              shapes: Any) -> Any:
    x0, y0, x1, y1 = tb.bbox
    shape = (shapes if shapes is not None else slide.shapes).add_textbox(
        Emu(xf.x(x0)), Emu(xf.y(y0)), Emu(xf.d(x1 - x0)), Emu(xf.d(y1 - y0)))
    if tb.rotation:
        shape.rotation = tb.rotation
    baseline = "baseline_line" in tb.notes
    if not baseline:  # the tuner maps PowerPoint's shapes back to blocks by name
        shape.name = f"{TEXT_SHAPE_PREFIX}{tb.id}"
    tf = shape.text_frame
    tf.word_wrap = tb.wrap
    tf.auto_size = (MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE if opts.autofit == "shrink" and not baseline
                    else MSO_AUTO_SIZE.NONE)
    ml, mt, mr, mb = tb.margins
    tf.margin_left, tf.margin_top = Emu(xf.d(ml)), Emu(xf.d(mt))
    tf.margin_right, tf.margin_bottom = Emu(xf.d(mr)), Emu(xf.d(mb))
    _fill_frame(tf, tb, xf, internal, opts, slide, work_dir, pictures, shape, baseline)
    return shape


def _fill_frame(tf: Any, tb: TextBlock, xf: _Xf, internal: list[tuple[Any, int]],
                opts: TextOptions, slide: Any, work_dir: Path | None,
                pictures: dict[str, Image] | None, shape: Any, baseline: bool = False) -> None:
    """Paragraphs and runs of ``tb`` into a text frame (text box or table cell)."""
    for pi, para in enumerate(tb.paragraphs):
        p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
        if not baseline:
            _para_props(p, para, xf, opts, slide, work_dir or Path("."), pictures or {})
        runs = para.runs if opts.keep_hyphens else _drop_hyphens(para)
        for ri, run in enumerate(runs):
            # a forced break (VT) is a control char: split on it before _xml_safe drops it
            text = LINE_BREAK.join(_xml_safe(t) for t in run.text.split(LINE_BREAK))
            if not text:
                continue
            # a slide-number field only where it shows the PDF's number (M7: build
            # collapsing renumbers the slides after a build-up)
            if run.field_type and not baseline and (
                    run.field_type != "slidenum" or opts.slide_no is None
                    or text.strip() == str(opts.slide_no)):
                _field(p, run, text, xf, f"{tb.id}/{pi}/{ri}")
                continue
            if LINE_BREAK in text and not baseline:  # forced line breaks (tuner, fidelity)
                for k, piece in enumerate(text.split(LINE_BREAK)):
                    if k:
                        _run_props(etree.SubElement(etree.SubElement(p._p, qn("a:br")),
                                                    qn("a:rPr")), run, xf)
                    if piece:
                        _add_run(p, run, piece, xf, shape, internal)
                continue
            r = p.add_run()
            r.text = text
            if baseline:  # the M2 path, byte-for-byte as before
                f = r.font
                if run.font:
                    f.name = run.font
                f.size = Emu(xf.d(run.size))
                f.bold = run.bold
                f.italic = run.italic
                f.color.rgb = _rgb(run.color)
                if run.char_spacing:
                    r._r.get_or_add_rPr().set("spc",
                                              str(pt_to_centipoints(run.char_spacing * xf.s)))
            else:
                _run_props(r._r.get_or_add_rPr(), run, xf)
            if run.hyperlink and not run.hyperlink.startswith("#slide="):
                r.hyperlink.address = run.hyperlink
            elif run.hyperlink and shape is not None:
                internal.append((shape, int(run.hyperlink.split("=", 1)[1]) - 1))


def _add_run(p: Any, run: TextRun, text: str, xf: _Xf, shape: Any,
             internal: list[tuple[Any, int]]) -> None:
    r = p.add_run()
    r.text = text
    _run_props(r._r.get_or_add_rPr(), run, xf)
    if run.hyperlink and not run.hyperlink.startswith("#slide="):
        r.hyperlink.address = run.hyperlink
    elif run.hyperlink and shape is not None:
        internal.append((shape, int(run.hyperlink.split("=", 1)[1]) - 1))


def _path_cmds(shape: Shape, xf: _Xf) -> tuple[str, bool]:
    """Inner XML of ``a:path`` relative to the frame; and whether any subpath is closed."""
    x0, y0 = shape.bbox[0], shape.bbox[1]
    out: list[str] = []
    closed = False
    started = False
    for seg in shape.path or []:
        pts = [(xf.d(px - x0), xf.d(py - y0)) for px, py in seg.pts]
        if seg.op == "M" or (not started and pts):
            out.append(f"<a:moveTo>{xml_ext.pt(*pts[0])}</a:moveTo>")
            started = True
            pts = pts[1:] if seg.op == "M" else []
            if not pts:
                continue
        if seg.op == "L":
            out.append(f"<a:lnTo>{xml_ext.pt(*pts[0])}</a:lnTo>")
        elif seg.op == "C" and len(pts) == 3:
            out.append("<a:cubicBezTo>" + "".join(xml_ext.pt(*p) for p in pts) + "</a:cubicBezTo>")
        elif seg.op == "Z" and started:
            out.append("<a:close/>")
            closed = True
    return "".join(out), closed


def _add_freeform(slide: Any, shape: Shape, xf: _Xf, shapes: Any = None) -> bool:
    shapes = shapes if shapes is not None else slide.shapes
    if not shape.path:
        return False
    cmds, _closed = _path_cmds(shape, xf)
    if not cmds:
        return False
    x0, y0, x1, y1 = shape.bbox
    fill: Any = shape.fill if shape.fill.kind in ("solid", "gradient") else None
    if fill is not None and shape.opacity < 0.9995 and fill.color is not None:
        fill = Fill(kind="solid", color=Color(rgb=fill.color.rgb, theme=fill.color.theme,
                                              alpha=fill.color.alpha * shape.opacity))
    sid = shapes._next_shape_id
    el = xml_ext.freeform(sid, f"Freeform {sid}", (xf.x(x0), xf.y(y0)),
                          (xf.d(x1 - x0), xf.d(y1 - y0)), cmds, fill, shape.stroke,
                          stroke_scale=xf.s)
    shapes._spTree.append(el)
    return True


_PRESET_NAMES = {"rect": "Rectangle", "roundRect": "Rounded Rectangle", "ellipse": "Oval",
                 "line": "Straight Connector", "bentConnector2": "Elbow Connector",
                 "bentConnector3": "Elbow Connector"}


@dataclass
class _Ids:
    """IR id -> written shape id, and connectors whose references are patched last."""

    sids: dict[str, int] = field(default_factory=dict)
    pending: list[tuple[Any, Shape]] = field(default_factory=list)

    def patch(self) -> None:
        for el, e in self.pending:
            c = e.connector
            if c is None:
                continue
            st = (self.sids[c.start_id], c.start_site or 0) if c.start_id in self.sids else None
            en = (self.sids[c.end_id], c.end_site or 0) if c.end_id in self.sids else None
            if st or en:
                xml_ext.set_connection(el, st, en)


def _add_shape(shapes: Any, e: Shape, xf: _Xf, ids: _Ids, locked: bool = False) -> str | None:
    """A preset shape, connector or freeform. Returns the count key, or None."""
    if e.fill.kind == "pattern":
        return None  # drawn by its raster image
    if not e.preset:
        return "freeforms" if _add_freeform(None, e, xf, shapes) else None
    x0, y0, x1, y1 = e.bbox
    off, ext = (xf.x(x0), xf.y(y0)), (xf.d(x1 - x0), xf.d(y1 - y0))
    sid = shapes._next_shape_id
    name = f"{_PRESET_NAMES.get(e.preset, e.preset[:1].upper() + e.preset[1:])} {sid}"
    if e.connector is not None:
        el = xml_ext.connector(sid, name, off, ext, e.preset, e.adjust, e.stroke,
                               rot=e.rotation, flip_h=e.flip_h, flip_v=e.flip_v,
                               stroke_scale=xf.s)
        ids.pending.append((el, e))
        key = "connectors"
    else:
        el = xml_ext.preset_shape(sid, name, off, ext, e.preset, e.adjust, e.fill, e.stroke,
                                  rot=e.rotation, flip_h=e.flip_h, flip_v=e.flip_v,
                                  stroke_scale=xf.s, opacity=e.opacity, locked=locked)
        key = "preset_shapes"
    shapes._spTree.append(el)
    ids.sids[e.id] = sid
    return key


def _covers_slide(page: Page, xf: _Xf, sw: float, sh: float) -> bool:
    return abs(page.width * xf.s - sw) < 0.5 and abs(page.height * xf.s - sh) < 0.5


def _add_background(slide: Any, e: Shape | Image, page: Page, xf: _Xf, sw: float, sh: float,
                    work_dir: Path, ids: _Ids) -> str:
    if _covers_slide(page, xf, sw, sh):
        if isinstance(e, Shape):
            xml_ext.background(slide, fill=e.fill)
        else:
            _, rid = slide.part.get_or_add_image_part(str(work_dir / e.file))
            xml_ext.background(slide, image_rid=rid)
        return "backgrounds"
    if isinstance(e, Shape):
        _add_shape(slide.shapes, e, xf, ids, locked=True)
    else:
        pic = _add_picture(slide, e, xf, work_dir)
        xml_ext.lock_picture(pic._element)
    return "locked_backgrounds"


def _add_picture(slide: Any, img: Image, xf: _Xf, work_dir: Path, shapes: Any = None) -> Any:
    shapes = shapes if shapes is not None else slide.shapes
    x0, y0, x1, y1 = img.bbox
    path = work_dir / img.file
    pic = shapes.add_picture(str(path), Emu(xf.x(x0)), Emu(xf.y(y0)),
                             Emu(max(xf.d(x1 - x0), 1)), Emu(max(xf.d(y1 - y0), 1)))
    if img.svg_file:
        svg = (work_dir / img.svg_file).read_bytes()
        rid = xml_ext.add_svg_part(shapes.part, svg)
        blip = pic._element.find(qn("p:blipFill")).find(qn("a:blip"))
        xml_ext.svg_blip(blip, rid)
    cl, ct, cr, cb = img.crop
    if any(img.crop):
        # The frame already shows only the visible part: crop = the hidden fractions.
        pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = cl, ct, cr, cb
    if img.rotation:
        pic.rotation = img.rotation
    if img.flip_h or img.flip_v:
        xfrm = pic._element.spPr.find(qn("a:xfrm"))
        if img.flip_h:
            xfrm.set("flipH", "1")
        if img.flip_v:
            xfrm.set("flipV", "1")
    return pic


TABLE_SHAPE_PREFIX = "Table "  # + the table id: the tuner maps cells by this name


def table_cell_rect(t: Table, row: int, col: int, row_span: int = 1,
                    col_span: int = 1) -> tuple[float, float, float, float]:
    x0 = t.bbox[0] + sum(t.col_widths[:col])
    y0 = t.bbox[1] + sum(t.row_heights[:row])
    return (x0, y0, x0 + sum(t.col_widths[col:col + col_span]),
            y0 + sum(t.row_heights[row:row + row_span]))


def cell_margins(t: Table, cell: Any) -> tuple[float, float, float, float]:
    """Cell margins (pt) that put the cell's text frame where ``cell.text.bbox`` is."""
    if cell.text is None:
        return cell.margins
    cx0, cy0, cx1, _ = table_cell_rect(t, cell.row, cell.col, cell.row_span, cell.col_span)
    tx0, ty0, tx1, _ = cell.text.bbox
    ml, mr = max(tx0 - cx0, 0.0), max(cx1 - tx1, 0.0)
    # Measured (M6): PowerPoint widens a merged cell's last column until it holds
    # marL + marR, so a spanning cell keeps both within that column's width
    last = t.col_widths[cell.col + cell.col_span - 1]
    excess = ml + mr - max(last - 1.0, 0.0)
    if cell.col_span > 1 and excess > 0:
        cut = min(mr, excess)
        mr -= cut
        ml = max(ml - (excess - cut), 0.0)
    return ml, max(ty0 - cy0, 0.0), mr, 0.0


def _add_table(slide: Any, t: Table, xf: _Xf, internal: list[tuple[Any, int]],
               opts: TextOptions, work_dir: Path, pictures: dict[str, Image]) -> Any:
    """A native table: column widths, row heights, merges, fills, borders, cell margins
    from the cell text's frame (``TableCell.text.bbox``), and the cell paragraphs."""
    n_rows, n_cols = len(t.row_heights), len(t.col_widths)
    x0, y0, x1, y1 = t.bbox
    gf = slide.shapes.add_table(n_rows, n_cols, Emu(xf.x(x0)), Emu(xf.y(y0)),
                                Emu(xf.d(x1 - x0)), Emu(xf.d(y1 - y0)))
    gf.name = f"{TABLE_SHAPE_PREFIX}{t.id}"
    tbl = gf.table
    xml_ext.plain_table(tbl)
    for c, w in enumerate(t.col_widths):
        tbl.columns[c].width = Emu(xf.d(w))
    for r, h in enumerate(t.row_heights):
        tbl.rows[r].height = Emu(xf.d(h))
    cells = {(c.row, c.col): c for c in t.cells}
    for c in t.cells:
        if c.merged_into is None and (c.row_span > 1 or c.col_span > 1):
            tbl.cell(c.row, c.col).merge(tbl.cell(c.row + c.row_span - 1,
                                                  c.col + c.col_span - 1))
    for r in range(n_rows):
        for col in range(n_cols):
            ic = cells.get((r, col))
            if ic is None:
                continue
            cell = tbl.cell(r, col)
            _, cy0, _, cy1 = table_cell_rect(t, r, col, ic.row_span, ic.col_span)
            ml, mt, mr, mb = cell_margins(t, ic)
            xml_ext.cell_props(cell._tc, margins=(xf.d(ml), xf.d(mt), xf.d(mr), xf.d(mb)),
                               anchor=ic.anchor, borders=dict(ic.borders), fill=ic.fill,
                               stroke_scale=xf.s)
            tf = cell.text_frame
            if ic.text is not None and ic.merged_into is None:
                _fill_frame(tf, ic.text, xf, internal, opts, slide, work_dir, pictures, None)
            # an empty paragraph mark must not make the row taller than the PDF's
            room = max((cy1 - cy0 - mt - mb) / SINGLE_LINE * 0.9, 1.0)
            for p in tf.paragraphs:
                end = p._p.get_or_add_endParaRPr()
                if not p.runs:
                    end.set("sz", str(pt_to_centipoints(min(room, 10.0) * xf.s)))
    return gf


def _write_element(slide: Any, shapes: Any, e: Any, xf: _Xf,
                   internal: list[tuple[Any, int]], rep: WriteReport, opts: TextOptions,
                   work_dir: Path, pictures: dict[str, Image], ids: _Ids) -> int | None:
    """Write one element; returns its shape id (None when nothing was drawn)."""
    sid: int | None = None
    proxy: Any = None
    if isinstance(e, TextBlock):
        box = _add_text(slide, e, xf, internal, opts, work_dir, pictures, shapes)
        sid = box.shape_id
        kind = _placeholder_kind(e, opts)
        rep.counts["text_boxes" if kind is None else f"placeholders_{kind}"] += 1
    elif isinstance(e, Image):
        proxy = _add_picture(slide, e, xf, work_dir, shapes)
        sid = proxy.shape_id
        rep.counts["svg_pictures" if e.svg_file else "pictures"] += 1
    elif isinstance(e, Shape):
        n = len(shapes._spTree)
        key = _add_shape(shapes, e, xf, ids)
        if key:
            rep.counts[key] += 1
            el = shapes._spTree[n] if len(shapes._spTree) > n else None
            if el is not None and el.tag != qn("p:cxnSp"):
                proxy = SlideShapeFactory(el, shapes)
                sid = proxy.shape_id
            elif el is not None:
                sid = int(el.find(".//" + qn("p:cNvPr")).get("id"))
    elif isinstance(e, Table):
        gf = _add_table(slide, e, xf, internal, opts, work_dir, pictures)
        sid = gf.shape_id
        rep.counts["tables"] += 1
    else:
        rep.counts[f"skipped_{e.kind}"] += 1
    link = getattr(e, "hyperlink", None)
    if proxy is not None and link:
        if link.startswith("#slide="):
            internal.append((proxy, int(link.split("=", 1)[1]) - 1))
        else:
            proxy.click_action.hyperlink.address = link
        rep.counts["shape_links"] += 1
    return sid


def _add_fallback(slide: Any, page: Page, pdf: Path, work_dir: Path, sw: float,
                  sh: float) -> None:
    png = ensure_render(pdf, work_dir, page.index, 300)
    page.renders["300"] = str(png.relative_to(work_dir)).replace("\\", "/")
    xf = _xf_for(page, sw, sh)
    slide.shapes.add_picture(str(png), Emu(xf.x(0)), Emu(xf.y(0)), Emu(xf.d(page.width)),
                             Emu(xf.d(page.height)))


def _clear(slide: Any) -> None:
    tree = slide.shapes._spTree
    for el in list(tree):
        if el.tag in (qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"),
                      qn("p:cxnSp")):
            tree.remove(el)


def _drawn_elsewhere(e: Any) -> bool:
    """M7: drawn by the slide master or a layout (``analyze/masters.py`` notes)."""
    return any(n.startswith(("in_master:", "in_layout:")) for n in e.notes)


def _build_page(slide: Any, page: Page, doc: Document, work_dir: Path, sw: float, sh: float,
                internal: list[tuple[Any, int]], rep: WriteReport,
                opts: TextOptions | None = None, items: list[Any] | None = None
                ) -> dict[str, int]:
    """Write ``page`` (or ``items`` placed like ``page``) onto ``slide``; returns the
    shape id of every top-level element written (IR id -> shape id)."""
    opts = opts or TextOptions()
    xf = _xf_for(page, sw, sh)
    own = items is None
    items = list(page.elements) if items is None else list(items)
    if opts.layout == "blocks" and any(isinstance(e, TextBlock) for e in items):
        items = [e for e in items if not (isinstance(e, Image) and BULLET_IMAGE_NOTE in e.notes)]
    elif own:
        items = [e for e in items if not isinstance(e, TextBlock)]
        items.extend(line_blocks(page, doc.fonts))
    if own and opts.structure:
        items = [e for e in items if not _drawn_elsewhere(e)]
    pictures = {e.id: e for e in page.elements if isinstance(e, Image)}
    ids = _Ids()
    present = {e.id for e in items}
    groups = [g for g in items if isinstance(g, Group)
              and sum(1 for c in g.child_ids if c in present) >= 2]
    child_of = {c: g.id for g in groups for c in g.child_ids if c in present}
    by_id = {e.id: e for e in items}
    top = [e for e in items if e.id not in child_of and not (
        isinstance(e, Group) and e not in groups)]
    # stable: by z; text ties broken by reading order
    order = sorted(range(len(top)), key=lambda i: (
        top[i].z, getattr(top[i], "reading_order", None) or 0, i))
    written: dict[str, int] = {}
    for i in order:
        e = top[i]
        if isinstance(e, (Shape, Image)) and e.is_background:
            rep.counts[_add_background(slide, e, page, xf, sw, sh, work_dir, ids)] += 1
        elif isinstance(e, Group):
            grp = slide.shapes.add_group_shape()
            grp.name = f"Group {grp.shape_id}"
            kids = sorted((by_id[c] for c in e.child_ids if c in by_id),
                          key=lambda k: (k.z, k.id))
            for k in kids:
                _write_element(slide, grp.shapes, k, xf, internal, rep, opts, work_dir,
                               pictures, ids)
            grp.shapes._recalculate_extents()
            written[e.id] = grp.shape_id
            rep.counts["groups"] += 1
        else:
            sid = _write_element(slide, slide.shapes, e, xf, internal, rep, opts, work_dir,
                                 pictures, ids)
            if sid is not None:
                written[e.id] = sid
    ids.patch()
    return written


class _Owner:
    """A slide master or layout seen as a slide by the element writers."""

    def __init__(self, real: Any) -> None:
        self.real = real
        self.part = real.part
        self._element = real._element
        self.shapes = SlideShapes(real.shapes._spTree, real)


def _emu_box(b: tuple[float, float, float, float], xf: _Xf) -> tuple[int, int, int, int]:
    return xf.x(b[0]), xf.y(b[1]), xf.d(b[2] - b[0]), xf.d(b[3] - b[1])


def _rep_page(doc: Document, items: list[Any], fallback: Page) -> Page:
    """The page a master / layout copy was taken from (for its placement)."""
    for e in items:
        for n in e.notes:
            if n.startswith("from:"):
                pi = int(n[5:].split(":", 1)[0])
                if 0 <= pi < len(doc.pages):
                    return doc.pages[pi]
    return fallback


LAYOUT_TYPES = {"Title Slide": "title", "Title and Content": "obj", "Two Content": "twoObj",
                "Title Only": "titleOnly", "Blank": "blank",
                "Blank (no master graphics)": "blank"}


def _write_masters(prs: Any, doc: Document, work_dir: Path, sw: float, sh: float,
                   rep: WriteReport, opts: TextOptions) -> dict[str, Any]:
    """Slide master + layouts from ``doc.masters[0]``; returns layout id -> layout."""
    m = doc.masters[0]
    pages = [p for p in doc.pages if p.layout_id and not p.triage.fallback_image]
    ref = pages[0] if pages else doc.pages[0]
    xf = _xf_for(ref, sw, sh)
    structure.neutral_text_styles(prs)
    master = prs.slide_master
    structure.clear_tree(master)
    mtree = master.shapes._spTree
    slots: dict[str, tuple[float, float, float, float]] = {}
    for lay in m.layouts:
        for k, b in lay.placeholders.items():
            slots.setdefault(k, b)
    defaults = {"title": (0.05 * ref.width, 0.04 * ref.height, 0.95 * ref.width,
                          0.2 * ref.height),
                "body": (0.05 * ref.width, 0.23 * ref.height, 0.95 * ref.width,
                         0.88 * ref.height)}
    for k in ("title", "body", "sldNum"):
        box = slots.get(k) or defaults.get(k)
        if box is not None:
            sid = structure.next_id(master)
            mtree.append(structure.placeholder_xml(sid, k, _emu_box(box, xf), master=True))
    internal: list[tuple[Any, int]] = []
    owner = _Owner(master)
    _build_page(owner, _rep_page(doc, m.elements, ref), doc, work_dir, sw, sh, internal, rep,
                opts, items=m.elements)
    structure.renumber_ids(master)
    members = {c for g in m.elements if isinstance(g, Group) for c in g.child_ids}
    rep.counts["master_elements"] += len([e for e in m.elements if e.id not in members])
    stock = list(prs.slide_layouts)
    out: dict[str, Any] = {}
    for k, lay in enumerate(m.layouts[:len(stock)]):
        pl = stock[k]
        structure.clear_tree(pl)
        bg = pl._element.find(qn("p:cSld")).find(qn("p:bg"))
        if bg is not None:
            bg.getparent().remove(bg)
        structure.set_layout_props(pl, lay.name, LAYOUT_TYPES.get(lay.name), lay.show_master)
        for kind in ("title", "subTitle", "body", "body2", "sldNum"):
            if kind in lay.placeholders:
                sid = structure.next_id(pl)
                pl.shapes._spTree.append(structure.placeholder_xml(
                    sid, kind, _emu_box(lay.placeholders[kind], xf)))
        if lay.elements:
            _build_page(_Owner(pl), _rep_page(doc, lay.elements, ref), doc, work_dir, sw, sh,
                        internal, rep, opts, items=lay.elements)
            structure.renumber_ids(pl)
            lm = {c for g in lay.elements if isinstance(g, Group) for c in g.child_ids}
            rep.counts["layout_elements"] += len([e for e in lay.elements if e.id not in lm])
        out[lay.id] = pl
    for pl in stock[len(out):]:
        prs.slide_layouts.remove(pl)
    rep.counts["layouts"] = len(out)
    return out


def _timing(page: Page, written: dict[str, int]) -> tuple[Any, int] | None:
    """The appear animations of a build-up's final page (``analyze/builds.py`` notes)."""
    steps: dict[int, list[tuple[int, tuple[int, int] | None]]] = {}
    bld: list[tuple[int, str]] = []
    for e in sorted(page.elements, key=lambda x: (x.z, x.id)):
        sid = written.get(e.id)
        if sid is None or getattr(e, "is_background", False):
            continue
        paras = paragraph_steps(e) if isinstance(e, TextBlock) else {}
        if paras:
            for pi in sorted(paras):
                steps.setdefault(paras[pi], []).append((sid, (pi, pi)))
            bld.append((sid, "p"))
            continue
        st = element_step(e)
        if st > 0:
            steps.setdefault(st, []).append((sid, None))
            if isinstance(e, (TextBlock, Shape)):
                bld.append((sid, "sp"))
            elif isinstance(e, Table):
                bld.append((sid, "graphic"))
    if not steps:
        return None
    ordered = [steps[k] for k in sorted(steps)]
    return structure.timing_xml(ordered, bld), sum(len(s) for s in ordered)


def write_pptx(doc: Document, out: Path, work_dir: Path, pdf: Path,
               opts: TextOptions | None = None) -> WriteReport:
    """Write ``doc`` to ``out``. Never fails because of a single page.

    With ``opts.structure`` (M7) the repeated elements go to the slide master and
    layouts, page numbers become slide-number placeholders, and (per ``opts``) titles
    and bodies become placeholders, colours / fonts theme references and build-up pages
    one animated slide. Notes, sections, links and document properties are always
    written."""
    opts = opts or TextOptions()
    sw, sh, g = slide_size(doc)
    for attempt in (opts, opts.flat()):
        rep = WriteReport()
        prs = _new_presentation(doc, sw, sh)
        use_theme = attempt.theme_refs and attempt.structure and doc.theme is not None
        tok_refs = xml_ext.THEME_REFS.set(use_theme)
        tok_fonts = THEME_FONTS.set((doc.theme.major_font, doc.theme.minor_font)
                                    if use_theme and doc.theme else (None, None))
        try:
            _write_slides(prs, doc, work_dir, pdf, attempt, rep, sw, sh)
            break
        except _Restart:
            continue  # the flat deck
        finally:
            xml_ext.THEME_REFS.reset(tok_refs)
            THEME_FONTS.reset(tok_fonts)
    if g != 1.0:
        LOG.warning("page size outside PowerPoint limits: deck scaled by %.3f", g)
    buf = io.BytesIO()
    prs.save(buf)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_name(out.name + ".tmp")
    tmp.write_bytes(normalize_zip(buf.getvalue()))
    tmp.replace(out)
    return rep


def _new_presentation(doc: Document, sw: float, sh: float) -> Any:
    prs = Presentation()
    prs.slide_width = Emu(pt_to_emu(sw))
    prs.slide_height = Emu(pt_to_emu(sh))
    cp = prs.core_properties
    cp.title = _xml_safe(doc.metadata.get("Title", ""))[:255]
    cp.author = _xml_safe(doc.metadata.get("Author", ""))[:255] or "pdf2pptx"
    cp.subject = _xml_safe(doc.metadata.get("Subject", ""))[:255]
    cp.keywords = _xml_safe(doc.metadata.get("Keywords", ""))[:255]
    cp.last_modified_by = "pdf2pptx"
    cp.revision = 1
    cp.created = cp.modified = FIXED_CORE_TIME
    structure.apply_theme(prs, doc.theme)
    return prs


def _plain_layout_id(doc: Document) -> str | None:
    for m in doc.masters:
        for lay in m.layouts:
            if not lay.show_master and not lay.elements and not set(lay.placeholders) - {
                    "sldNum"}:
                return lay.id
    return None


def _write_slides(prs: Any, doc: Document, work_dir: Path, pdf: Path, opts: TextOptions,
                  rep: WriteReport, sw: float, sh: float) -> None:
    layouts: dict[str, Any] = {}
    if opts.structure and doc.masters:
        try:
            layouts = _write_masters(prs, doc, work_dir, sw, sh, rep, opts)
        except Exception as exc:  # a flat deck instead of a broken master
            msg = f"master/layouts failed ({type(exc).__name__}: {exc}); flat deck"
            LOG.error(msg, exc_info=LOG.isEnabledFor(10))
            rep.errors.append(msg)
            raise _Restart(msg) from exc
    structured = bool(layouts)
    if not structured:
        opts = dataclasses.replace(opts, structure=False, placeholders=False)
    blank = next((pl for pl in prs.slide_layouts if pl.name == "Blank"), prs.slide_layouts[-1])
    plain = layouts.get(_plain_layout_id(doc) or "")
    gone = absorbed_pages(doc) if opts.builds else set()
    slide_of = page_to_slide(doc, opts.builds)
    internal: list[tuple[Any, int]] = []
    slides = []
    for page in doc.pages:
        if page.index in gone:
            continue
        layout = layouts.get(page.layout_id or "") if structured else None
        if layout is None:
            layout = plain if structured and plain is not None else (
                next(iter(layouts.values())) if structured else blank)
        slide = prs.slides.add_slide(layout)
        _clear(slide)  # placeholders cloned from the layout: ours are written explicitly
        slides.append(slide)
        page_opts = dataclasses.replace(opts, slide_no=len(slides))
        if page.notes_text:
            slide.notes_slide.notes_text_frame.text = _xml_safe(page.notes_text)
            rep.counts["notes"] += 1
        if not page.triage.fallback_image:
            try:
                n_links = len(internal)
                written = _build_page(slide, page, doc, work_dir, sw, sh, internal, rep,
                                      page_opts)
                if opts.builds and page.build_of is not None:
                    tm = _timing(page, written)
                    if tm is not None:
                        structure.set_timing(slide, tm[0])
                        rep.counts["animations"] += tm[1]
                rep.slides += 1
                continue
            except Exception as exc:
                del internal[n_links:]
                msg = f"page {page.index}: build failed ({type(exc).__name__}: {exc})"
                LOG.error("%s; using the page render", msg, exc_info=LOG.isEnabledFor(10))
                rep.errors.append(msg)
                page.triage.fallback_image = True
                page.triage.reason = (page.triage.reason + "; " if page.triage.reason
                                      else "") + msg
                _clear(slide)
                old = slide._element.find(qn("p:timing"))
                if old is not None:
                    slide._element.remove(old)
        _add_fallback(slide, page, pdf, work_dir, sw, sh)
        rep.fallback_pages.append(page.index)
        rep.counts["fallback_pictures"] += 1
        rep.slides += 1
    for shape, target in internal:
        if 0 <= target < len(slide_of) and slide_of[target] < len(slides):
            shape.click_action.target_slide = slides[slide_of[target]]
    if doc.sections:
        secs: list[tuple[str, list[int]]] = []
        for k, sec in enumerate(doc.sections):
            first = slide_of[sec.first_page] if sec.first_page < len(slide_of) else len(slides)
            nxt = doc.sections[k + 1].first_page if k + 1 < len(doc.sections) else len(
                doc.pages)
            last = slide_of[nxt] if nxt < len(slide_of) else len(slides)
            idxs = [i for i in range(first, last) if not secs or i > secs[-1][1][-1]]
            if idxs:
                secs.append((_xml_safe(sec.name), idxs))
        structure.set_sections(prs, secs)
        rep.counts["sections"] = len(secs)


class _Restart(Exception):
    """The structured write failed before any slide: write the flat deck instead."""


def normalize_zip(data: bytes) -> bytes:
    """Same entries, same order, fixed timestamps and attributes -> deterministic bytes."""
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            zi = zipfile.ZipInfo(info.filename, date_time=FIXED_ZIP_TIME)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            dst.writestr(zi, src.read(info.filename))
    return out.getvalue()
