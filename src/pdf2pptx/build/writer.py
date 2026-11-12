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
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu

from .. import log
from ..ingest.render import ensure_render
from ..ir import Bullet, Color, Document, Image, Page, Paragraph, Shape, TextBlock, TextRun
from . import xml_ext
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
    """What the writer may do with text, from the mode."""

    layout: str = "blocks"  # blocks | lines (the M2 baseline, for comparison)
    autofit: str = "none"  # none | shrink
    keep_hyphens: bool = True  # False: drop end-of-line hyphenation (editable)

    @classmethod
    def for_mode(cls, mode: str, layout: str = "blocks") -> TextOptions:
        editable = mode in ("editable", "deck")
        return cls(layout=layout, autofit="shrink" if editable else "none",
                   keep_hyphens=not editable)


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
    _el(parent, "a:srgbClr", val=(c.rgb or "000000").upper())


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
        _el(rpr, "a:latin", typeface=run.font)
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


def _add_text(slide: Any, tb: TextBlock, xf: _Xf, internal: list[tuple[Any, int]],
              opts: TextOptions | None = None, work_dir: Path | None = None,
              pictures: dict[str, Image] | None = None) -> None:
    opts = opts or TextOptions()
    x0, y0, x1, y1 = tb.bbox
    shape = slide.shapes.add_textbox(Emu(xf.x(x0)), Emu(xf.y(y0)), Emu(xf.d(x1 - x0)),
                                     Emu(xf.d(y1 - y0)))
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
            if run.field_type and not baseline:
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
            elif run.hyperlink:
                internal.append((shape, int(run.hyperlink.split("=", 1)[1]) - 1))


def _add_run(p: Any, run: TextRun, text: str, xf: _Xf, shape: Any,
             internal: list[tuple[Any, int]]) -> None:
    r = p.add_run()
    r.text = text
    _run_props(r._r.get_or_add_rPr(), run, xf)
    if run.hyperlink and not run.hyperlink.startswith("#slide="):
        r.hyperlink.address = run.hyperlink
    elif run.hyperlink:
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


def _add_freeform(slide: Any, shape: Shape, xf: _Xf) -> bool:
    if not shape.path:
        return False
    cmds, _closed = _path_cmds(shape, xf)
    if not cmds:
        return False
    x0, y0, x1, y1 = shape.bbox
    fill = shape.fill.color if shape.fill.kind == "solid" else None
    sid = slide.shapes._next_shape_id
    el = xml_ext.freeform(sid, f"Freeform {sid}", (xf.x(x0), xf.y(y0)),
                          (xf.d(x1 - x0), xf.d(y1 - y0)), cmds, fill, shape.stroke,
                          stroke_scale=xf.s)
    slide.shapes._spTree.append(el)
    return True


def _add_picture(slide: Any, img: Image, xf: _Xf, work_dir: Path) -> None:
    x0, y0, x1, y1 = img.bbox
    path = work_dir / img.file
    pic = slide.shapes.add_picture(str(path), Emu(xf.x(x0)), Emu(xf.y(y0)),
                                   Emu(max(xf.d(x1 - x0), 1)), Emu(max(xf.d(y1 - y0), 1)))
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


def _build_page(slide: Any, page: Page, doc: Document, work_dir: Path, sw: float, sh: float,
                internal: list[tuple[Any, int]], rep: WriteReport,
                opts: TextOptions | None = None) -> None:
    opts = opts or TextOptions()
    xf = _xf_for(page, sw, sh)
    items: list[Any] = list(page.elements)
    if opts.layout == "blocks" and any(isinstance(e, TextBlock) for e in items):
        items = [e for e in items if not (isinstance(e, Image) and BULLET_IMAGE_NOTE in e.notes)]
    else:
        items = [e for e in items if not isinstance(e, TextBlock)]
        items.extend(line_blocks(page, doc.fonts))
    pictures = {e.id: e for e in page.elements if isinstance(e, Image)}
    # stable: by z; text ties broken by reading order
    order = sorted(range(len(items)), key=lambda i: (
        items[i].z, getattr(items[i], "reading_order", None) or 0, i))
    for i in order:
        e = items[i]
        if isinstance(e, TextBlock):
            _add_text(slide, e, xf, internal, opts, work_dir, pictures)
            rep.counts["text_boxes"] += 1
        elif isinstance(e, Image):
            _add_picture(slide, e, xf, work_dir)
            rep.counts["pictures"] += 1
        elif isinstance(e, Shape):
            if _add_freeform(slide, e, xf):
                rep.counts["freeforms"] += 1
        else:
            rep.counts[f"skipped_{e.kind}"] += 1


def write_pptx(doc: Document, out: Path, work_dir: Path, pdf: Path,
               opts: TextOptions | None = None) -> WriteReport:
    """Write ``doc`` to ``out``. Never fails because of a single page."""
    opts = opts or TextOptions()
    rep = WriteReport()
    prs = Presentation()
    sw, sh, g = slide_size(doc)
    prs.slide_width = Emu(pt_to_emu(sw))
    prs.slide_height = Emu(pt_to_emu(sh))
    cp = prs.core_properties
    cp.title = _xml_safe(doc.metadata.get("Title", ""))[:255]
    cp.author = cp.last_modified_by = "pdf2pptx"
    cp.revision = 1
    cp.created = cp.modified = FIXED_CORE_TIME
    layout = prs.slide_layouts[BLANK_LAYOUT]
    internal: list[tuple[Any, int]] = []
    slides = []
    for page in doc.pages:
        slide = prs.slides.add_slide(layout)
        slides.append(slide)
        if not page.triage.fallback_image:
            try:
                n_links = len(internal)
                _build_page(slide, page, doc, work_dir, sw, sh, internal, rep, opts)
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
        _add_fallback(slide, page, pdf, work_dir, sw, sh)
        rep.fallback_pages.append(page.index)
        rep.counts["fallback_pictures"] += 1
        rep.slides += 1
    for shape, target in internal:
        if 0 <= target < len(slides):
            shape.click_action.target_slide = slides[target]
    if g != 1.0:
        LOG.warning("page size outside PowerPoint limits: deck scaled by %.3f", g)
    buf = io.BytesIO()
    prs.save(buf)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_name(out.name + ".tmp")
    tmp.write_bytes(normalize_zip(buf.getvalue()))
    tmp.replace(out)
    return rep


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
