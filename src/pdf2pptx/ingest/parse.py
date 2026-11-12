"""INGEST: PDF -> IR with pypdfium2 (glyphs, paths, images, annotations, metadata).

All coordinates are converted to display space: points, origin at the top-left
of the page *as displayed* (cropbox, /Rotate applied), y pointing down.

Conventions for rotated elements (shared with the writer):

* ``Glyph.rotation`` is counter-clockwise degrees in display space (as in the PDF
  text matrix); ``Glyph.origin_x`` / ``Glyph.baseline`` are the display-space
  coordinates of the glyph origin.
* ``Image.rotation`` / ``Shape.rotation`` are clockwise degrees (PowerPoint's
  convention) and their ``bbox`` is the *unrotated* frame centred on the element,
  exactly like a PowerPoint ``a:xfrm``.

One bad page never stops the document: :func:`parse_pages` catches the error,
logs it and returns a page marked ``triage.fallback_image`` so the writer places
that page's render instead.
"""

from __future__ import annotations

import ctypes
import hashlib
import io
import math
import re
from collections import Counter
from collections.abc import Sequence
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pypdfium2 as pdfium
import pypdfium2.raw as R

from .. import log
from ..ir import (
    Annotation,
    AnyElement,
    Color,
    Document,
    Fill,
    FontRef,
    Glyph,
    Image,
    Page,
    PathSegment,
    Shape,
    Stroke,
    Triage,
)
from . import geom
from .geom import Matrix

LOG = log.get(__name__)

PARSER_VERSION = "2.0"

_SUBSET_RE = re.compile(r"^[A-Z]{6}\+")
_STYLES = (r"Bold|Italic|Oblique|Regular|Semibold|SemiBold|Light|Medium|Black|Heavy|"
           r"Book|Demi|Condensed|Narrow|BoldItalic|BoldOblique|It|BoldMT|MT|PSMT|PS")
_STYLE_RE = re.compile(rf"[-,_ ]?({_STYLES})+$")  # trailing, no separator needed
_TAIL_RE = re.compile(rf"[-,]({_STYLES}|Roman)+")  # after a separator ("Times-Roman")
_CAPS = {0: "flat", 1: "round", 2: "square"}
_JOINS = {0: "miter", 1: "round", 2: "bevel"}


def _addr(handle: Any) -> int:
    return int(ctypes.cast(handle, ctypes.c_void_p).value or 0)


def _hex(r: int, g: int, b: int) -> str:
    return f"{r:02X}{g:02X}{b:02X}"


def clean_family(base_name: str) -> str:
    """'ABCDEF+Calibri-Bold' -> 'Calibri'; 'TimesNewRomanPS-BoldMT' -> 'TimesNewRoman'."""
    name = _SUBSET_RE.sub("", base_name)
    for sep in ("-", ","):
        if sep in name:
            head, tail = name.split(sep, 1)
            if head and _TAIL_RE.fullmatch(sep + tail):
                name = head
                break
    prev = None
    while prev != name:
        prev = name
        m = _STYLE_RE.search(name)
        if m and m.start() > 0:
            name = name[: m.start()]
    return name.strip() or base_name


# --------------------------------------------------------------------------- helpers


def _fill_color(obj: Any) -> Color | None:
    r, g, b, a = (ctypes.c_uint() for _ in range(4))
    if not R.FPDFPageObj_GetFillColor(obj, r, g, b, a):
        return None
    return Color(rgb=_hex(r.value, g.value, b.value), alpha=round(a.value / 255, 4))


def _stroke_color(obj: Any) -> Color | None:
    r, g, b, a = (ctypes.c_uint() for _ in range(4))
    if not R.FPDFPageObj_GetStrokeColor(obj, r, g, b, a):
        return None
    return Color(rgb=_hex(r.value, g.value, b.value), alpha=round(a.value / 255, 4))


def _matrix(obj: Any) -> Matrix:
    m = R.FS_MATRIX()
    if not R.FPDFPageObj_GetMatrix(obj, m):
        return geom.IDENTITY
    return (m.a, m.b, m.c, m.d, m.e, m.f)


def _clip_bbox(obj: Any, to_display: Matrix) -> tuple[float, float, float, float] | None:
    """Intersection of the bboxes of the object's clip paths, in display space."""
    clip = R.FPDFPageObj_GetClipPath(obj)
    if not clip:
        return None
    box: list[float] | None = None
    for p in range(R.FPDFClipPath_CountPaths(clip)):
        pts = []
        for s in range(R.FPDFClipPath_CountPathSegments(clip, p)):
            seg = R.FPDFClipPath_GetPathSegment(clip, p, s)
            x, y = ctypes.c_float(), ctypes.c_float()
            if R.FPDFPathSegment_GetPoint(seg, x, y):
                pts.append(geom.apply(to_display, x.value, y.value))
        if not pts:
            continue
        b = geom.bbox_of(pts)
        box = list(b) if box is None else [max(box[0], b[0]), max(box[1], b[1]),
                                           min(box[2], b[2]), min(box[3], b[3])]
    return None if box is None else (box[0], box[1], box[2], box[3])


def _r(v: float, nd: int = 3) -> float:
    out = round(float(v), nd)
    return 0.0 if out == 0 else out


def _rbox(b: Sequence[float]) -> tuple[float, float, float, float]:
    return (_r(b[0]), _r(b[1]), _r(b[2]), _r(b[3]))


# --------------------------------------------------------------------------- page parser


@dataclass
class _PageCtx:
    doc: Any
    page: Any
    index: int
    display: Matrix
    width: float
    height: float
    work_dir: Path
    elements: list[AnyElement] = field(default_factory=list)
    text_objects: dict[int, tuple[int, int]] = field(default_factory=dict)  # addr->(z, mode)
    skipped: Counter[str] = field(default_factory=Counter)
    # top-level path objects that claim an opaque white fill: (shape, top-level index)
    white_fills: list[tuple[Shape, int]] = field(default_factory=list)
    top_index: int = -1
    z: int = 0
    n_img: int = 0
    n_shape: int = 0


def _walk(ctx: _PageCtx, obj: Any, parent: Matrix) -> None:
    """Visit one page object (recursing into form XObjects), in content order."""
    t = R.FPDFPageObj_GetType(obj)
    to_display = geom.concat(parent, ctx.display)
    if t == R.FPDF_PAGEOBJ_FORM:
        child_parent = geom.concat(_matrix(obj), parent)
        for i in range(R.FPDFFormObj_CountObjects(obj)):
            _walk(ctx, R.FPDFFormObj_GetObject(obj, i), child_parent)
        return
    ctx.z += 1
    z = ctx.z
    if t == R.FPDF_PAGEOBJ_TEXT:
        ctx.text_objects[_addr(obj)] = (z, int(R.FPDFTextObj_GetTextRenderMode(obj)))
    elif t == R.FPDF_PAGEOBJ_PATH:
        shape = _path(ctx, obj, to_display, z)
        if shape is not None:
            ctx.elements.append(shape)
            col = shape.fill.color
            if (parent == geom.IDENTITY and col is not None and col.rgb == "FFFFFF"
                    and col.alpha >= 1.0):
                ctx.white_fills.append((shape, ctx.top_index))
    elif t == R.FPDF_PAGEOBJ_IMAGE:
        img = _image(ctx, obj, parent, to_display, z)
        if img is not None:
            ctx.elements.append(img)
    elif t == R.FPDF_PAGEOBJ_SHADING:
        ctx.skipped["shading"] += 1
    else:
        ctx.skipped[f"type{t}"] += 1


def _path(ctx: _PageCtx, obj: Any, to_display: Matrix, z: int) -> Shape | None:
    fill_mode, stroke_on = ctypes.c_int(), ctypes.c_int()
    R.FPDFPath_GetDrawMode(obj, fill_mode, stroke_on)
    if fill_mode.value == 0 and not stroke_on.value:
        return None  # clip-only / invisible path
    total = geom.concat(_matrix(obj), to_display)
    segs: list[PathSegment] = []
    pts_all: list[tuple[float, float]] = []
    n = R.FPDFPath_CountSegments(obj)
    i = 0
    while i < n:
        seg = R.FPDFPath_GetPathSegment(obj, i)
        kind = R.FPDFPathSegment_GetType(seg)
        x, y = ctypes.c_float(), ctypes.c_float()
        R.FPDFPathSegment_GetPoint(seg, x, y)
        p = geom.apply(total, x.value, y.value)
        last = seg
        if kind == R.FPDF_SEGMENT_BEZIERTO and i + 2 < n:
            pts = [p]
            for k in (1, 2):
                s2 = R.FPDFPath_GetPathSegment(obj, i + k)
                x2, y2 = ctypes.c_float(), ctypes.c_float()
                R.FPDFPathSegment_GetPoint(s2, x2, y2)
                pts.append(geom.apply(total, x2.value, y2.value))
                last = s2
            segs.append(PathSegment(op="C", pts=[(_r(a), _r(b)) for a, b in pts]))
            pts_all.extend(pts)
            i += 3
        else:
            op = "M" if kind == R.FPDF_SEGMENT_MOVETO or not segs else "L"
            segs.append(PathSegment(op=op, pts=[(_r(p[0]), _r(p[1]))]))
            pts_all.append(p)
            i += 1
        if R.FPDFPathSegment_GetClose(last):
            segs.append(PathSegment(op="Z"))
    if not pts_all:
        return None
    notes: list[str] = []
    fill = Fill()
    if fill_mode.value:
        col = _fill_color(obj)
        if col is None:
            notes.append("unsupported_fill:pattern")
            ctx.skipped["pattern_fill"] += 1
        else:
            fill = Fill(kind="solid", color=col)
        if fill_mode.value == R.FPDF_FILLMODE_ALTERNATE:
            notes.append("fill_rule:evenodd")
    stroke = None
    if stroke_on.value:
        w = ctypes.c_float()
        R.FPDFPageObj_GetStrokeWidth(obj, w)
        sc = geom.scale_of(total)
        dash_pattern = None
        nd = R.FPDFPageObj_GetDashCount(obj)
        if nd > 0:
            arr = (ctypes.c_float * nd)()
            if R.FPDFPageObj_GetDashArray(obj, arr, nd):
                dash_pattern = [_r(v * sc) for v in arr]
                if not any(dash_pattern):
                    dash_pattern = None
        stroke = Stroke(
            color=_stroke_color(obj) or Color(rgb="000000"),
            width=_r(w.value * sc),
            dash="custom" if dash_pattern else "solid",
            dash_pattern=dash_pattern,
            cap=_CAPS.get(R.FPDFPageObj_GetLineCap(obj), "flat"),
            join=_JOINS.get(R.FPDFPageObj_GetLineJoin(obj), "miter"),
        )
    if fill.kind == "none" and stroke is None:
        return None
    clip = _clip_bbox(obj, to_display)
    if clip is not None:
        notes.append("clip:" + ",".join(f"{v:.2f}" for v in clip))
    ctx.n_shape += 1
    return Shape(id=f"p{ctx.index}_s{ctx.n_shape}", bbox=_rbox(geom.bbox_of(pts_all)), z=z,
                 path=segs, fill=fill, stroke=stroke, notes=notes)


def _image(ctx: _PageCtx, obj: Any, parent: Matrix, to_display: Matrix, z: int) -> Image | None:
    total = geom.concat(_matrix(obj), to_display)
    tl = geom.apply(total, 0, 1)
    tr = geom.apply(total, 1, 1)
    bl = geom.apply(total, 0, 0)
    ex = (tr[0] - tl[0], tr[1] - tl[1])
    ey = (bl[0] - tl[0], bl[1] - tl[1])
    w, h = math.hypot(*ex), math.hypot(*ey)
    if w < 0.01 or h < 0.01:
        return None
    rot = math.degrees(math.atan2(ex[1], ex[0]))
    flip_v = ex[0] * ey[1] - ex[1] * ey[0] < 0
    rot = round(rot % 360.0, 3)
    if rot >= 359.999:
        rot = 0.0
    cx, cy = geom.apply(total, 0.5, 0.5)
    frame = (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
    pw, ph = ctypes.c_uint(), ctypes.c_uint()
    R.FPDFImageObj_GetImagePixelSize(obj, pw, ph)
    notes: list[str] = []

    data, ext, mask_bytes = _image_bytes(ctx, obj, axis_aligned=(rot == 0.0 and not flip_v),
                                         notes=notes)
    if data is None:
        ctx.skipped["image_unreadable"] += 1
        return None
    ctx.n_img += 1
    img_id = f"p{ctx.index}_i{ctx.n_img}"
    img_dir = ctx.work_dir / "images"
    img_dir.mkdir(parents=True, exist_ok=True)
    rel = f"images/{img_id}.{ext}"
    (ctx.work_dir / rel).write_bytes(data)
    mask_rel = None
    if mask_bytes is not None:
        mask_rel = f"images/{img_id}_mask.png"
        (ctx.work_dir / mask_rel).write_bytes(mask_bytes)

    crop = (0.0, 0.0, 0.0, 0.0)
    clip = _clip_bbox(obj, geom.concat(parent, ctx.display))
    if clip is not None and rot == 0.0 and not flip_v:
        x0, y0, x1, y1 = frame
        cl = max(0.0, (clip[0] - x0) / w)
        ct = max(0.0, (clip[1] - y0) / h)
        cr = max(0.0, (x1 - clip[2]) / w)
        cb = max(0.0, (y1 - clip[3]) / h)
        if cl + cr < 0.999 and ct + cb < 0.999 and max(cl, ct, cr, cb) > 1e-3:
            crop = (_r(cl, 5), _r(ct, 5), _r(cr, 5), _r(cb, 5))
            frame = (x0 + cl * w, y0 + ct * h, x1 - cr * w, y1 - cb * h)
            notes.append("crop_from_clip")
    return Image(id=img_id, bbox=_rbox(frame), z=z, file=rel,
                 sha256=hashlib.sha256(data).hexdigest(), crop=crop, mask_file=mask_rel,
                 rotation=rot, flip_v=flip_v, pixel_size=(int(pw.value), int(ph.value)),
                 notes=notes)


def _image_bytes(ctx: _PageCtx, obj: Any, *, axis_aligned: bool,
                 notes: list[str]) -> tuple[bytes | None, str, bytes | None]:
    """(file bytes, extension, mask PNG bytes or None).

    pdfium does not report soft masks (``/SMask``) through HasTransparency, so the
    alpha comes from the object's rendered bitmap (about 1 px per pt; upscaled to
    the image's pixel size). Rotated / flipped images keep no mask (noted).
    """
    from PIL import Image as PILImage

    filters = []
    for i in range(R.FPDFImageObj_GetImageFilterCount(obj)):
        n = R.FPDFImageObj_GetImageFilter(obj, i, None, 0)
        buf = ctypes.create_string_buffer(n)
        R.FPDFImageObj_GetImageFilter(obj, i, buf, n)
        filters.append(buf.value.decode("latin-1"))
    meta = R.FPDF_IMAGEOBJ_METADATA()
    R.FPDFImageObj_GetImageMetadata(obj, ctx.page, meta)
    plain_cs = meta.colorspace in (R.FPDF_COLORSPACE_DEVICERGB, R.FPDF_COLORSPACE_DEVICEGRAY)

    rendered = None
    rraw = R.FPDFImageObj_GetRenderedBitmap(ctx.doc, ctx.page, obj)
    if rraw:
        rendered = pdfium.PdfBitmap.from_raw(rraw).to_pil().convert("RGBA")
    alpha = rendered.getchannel("A") if rendered is not None else None
    transparent = alpha is not None and alpha.getextrema()[0] < 255

    if filters == ["DCTDecode"] and plain_cs and not transparent:
        n = R.FPDFImageObj_GetImageDataRaw(obj, None, 0)
        buf = ctypes.create_string_buffer(n)
        R.FPDFImageObj_GetImageDataRaw(obj, buf, n)
        return buf.raw, "jpg", None

    raw_bm = R.FPDFImageObj_GetBitmap(obj)
    pil = pdfium.PdfBitmap.from_raw(raw_bm).to_pil().copy() if raw_bm else None
    mask_png = None
    if transparent and alpha is not None:
        if pil is None:
            pil = rendered
        elif axis_aligned:
            alpha = alpha.resize(pil.size, PILImage.Resampling.BILINEAR)
            pil = pil.convert("RGB")
            pil.putalpha(alpha)
            mbuf = io.BytesIO()
            alpha.save(mbuf, "PNG", optimize=False)
            mask_png = mbuf.getvalue()
        else:
            notes.append("mask_dropped:rotated")
    if pil is None:
        return None, "", None
    if pil.mode not in ("RGB", "RGBA", "L", "LA"):
        pil = pil.convert("RGBA" if "A" in pil.mode else "RGB")
    out = io.BytesIO()
    pil.save(out, "PNG", optimize=False)
    return out.getvalue(), "png", mask_png


MAX_FILL_CHECKS = 200  # per page
CHECK_SCALE = 0.5  # px per pt for the uniformity check
RASTER_SCALE = 2.0  # px per pt (144 DPI) for a rasterised fill


def _isolated_render(ctx: _PageCtx, top_index: int, scale: float) -> Any:
    """RGBA render of one top-level object alone, transparent background (display space)."""
    import numpy as np

    page = ctx.page
    n = R.FPDFPage_CountObjects(page)
    objs = [R.FPDFPage_GetObject(page, i) for i in range(n)]
    try:
        for i, o in enumerate(objs):
            R.FPDFPageObj_SetIsActive(o, int(i == top_index))
        bm = page.render(scale=scale, fill_color=(0, 0, 0, 0), may_draw_forms=False)
        return np.asarray(bm.to_pil().convert("RGBA"))
    finally:
        for o in objs:
            R.FPDFPageObj_SetIsActive(o, 1)


def _rasterize_pattern_fills(ctx: _PageCtx) -> None:
    """pdfium reports pattern (e.g. gradient) fills as opaque white with no other hint.

    Check each such path with a low-res isolated render; a path that does not render
    as uniform white is replaced by an isolated 144-DPI render with alpha (an Image).
    """
    from PIL import Image as PILImage

    checks = ctx.white_fills[:MAX_FILL_CHECKS]
    if len(ctx.white_fills) > MAX_FILL_CHECKS:
        ctx.skipped["fill_checks_capped"] += len(ctx.white_fills) - MAX_FILL_CHECKS
    for shape, top in checks:
        x0, y0, x1, y1 = shape.bbox
        if (x1 - x0) * (y1 - y0) < 4.0:
            continue
        arr = _isolated_render(ctx, top, CHECK_SCALE)
        h, w = arr.shape[:2]
        c0, c1 = max(0, int(x0 * CHECK_SCALE)), min(w, int(x1 * CHECK_SCALE) + 1)
        r0, r1 = max(0, int(y0 * CHECK_SCALE)), min(h, int(y1 * CHECK_SCALE) + 1)
        sub = arr[r0:r1, c0:c1]
        inside = sub[..., 3] > 250
        if inside.sum() < 4:
            continue
        rgb = sub[..., :3][inside].astype(float)
        if rgb.mean() > 247 and rgb.std(axis=0).max() < 4:
            continue  # really white
        big = _isolated_render(ctx, top, RASTER_SCALE)
        h, w = big.shape[:2]
        c0, c1 = max(0, int(x0 * RASTER_SCALE)), min(w, int(x1 * RASTER_SCALE + 1))
        r0, r1 = max(0, int(y0 * RASTER_SCALE)), min(h, int(y1 * RASTER_SCALE + 1))
        if c1 <= c0 or r1 <= r0:
            continue
        crop = PILImage.fromarray(big[r0:r1, c0:c1])
        buf = io.BytesIO()
        crop.save(buf, "PNG", optimize=False)
        data = buf.getvalue()
        ctx.n_img += 1
        img_id = f"p{ctx.index}_i{ctx.n_img}"
        rel = f"images/{img_id}.png"
        (ctx.work_dir / "images").mkdir(parents=True, exist_ok=True)
        (ctx.work_dir / rel).write_bytes(data)
        frame = (c0 / RASTER_SCALE, r0 / RASTER_SCALE, c1 / RASTER_SCALE, r1 / RASTER_SCALE)
        img = Image(id=img_id, bbox=_rbox(frame), z=shape.z, file=rel,
                    sha256=hashlib.sha256(data).hexdigest(), pixel_size=crop.size,
                    source="rule", confidence=0.8,
                    notes=["raster:pattern_fill", f"replaces:{shape.id}"])
        ctx.elements[ctx.elements.index(shape)] = img
        ctx.skipped["pattern_fill_rasterized"] += 1


def _font_info(tp: Any, i: int) -> tuple[str, int]:
    flags = ctypes.c_int()
    n = R.FPDFText_GetFontInfo(tp, i, None, 0, flags)
    if n <= 0:
        return "", 0
    buf = ctypes.create_string_buffer(n)
    R.FPDFText_GetFontInfo(tp, i, buf, n, flags)
    return buf.value.decode("utf-8", "replace"), int(flags.value)


def _glyphs(ctx: _PageCtx, fonts: dict[str, FontRef]) -> list[Glyph]:
    tp = R.FPDFText_LoadPage(ctx.page)
    if not tp:
        return []
    out: list[Glyph] = []
    try:
        n = R.FPDFText_CountChars(tp)
        page_rot = R.FPDFPage_GetRotation(ctx.page) * 90
        font_cache: dict[int, tuple[str, int]] = {}
        for i in range(n):
            if R.FPDFText_IsGenerated(tp, i) == 1:
                continue
            u = R.FPDFText_GetUnicode(tp, i)
            if u in (0x0D, 0x0A, 0x02):
                continue
            ch = chr(u) if 0 < u < 0x110000 else "�"
            tobj = R.FPDFText_GetTextObject(tp, i)
            z, mode = ctx.text_objects.get(_addr(tobj), (0, 0))
            cl, r, b, t = (ctypes.c_double() for _ in range(4))
            R.FPDFText_GetCharBox(tp, i, cl, r, b, t)
            ox, oy = ctypes.c_double(), ctypes.c_double()
            R.FPDFText_GetCharOrigin(tp, i, ox, oy)
            m = R.FS_MATRIX()
            R.FPDFText_GetMatrix(tp, i, m)
            angle = geom.norm_angle(math.degrees(math.atan2(m.b, m.a)) - page_rot)
            size = float(R.FPDFText_GetFontSize(tp, i))
            fkey = _addr(tobj)
            if fkey not in font_cache:
                font_cache[fkey] = _font_info(tp, i)
            base, flags = font_cache[fkey]
            font_id = base or "unknown"
            if font_id not in fonts:
                fonts[font_id] = _font_ref(font_id, flags, int(R.FPDFText_GetFontWeight(tp, i)),
                                           tobj)
            cr, cg, cb, ca = (ctypes.c_uint() for _ in range(4))
            color = Color(rgb="000000")
            if R.FPDFText_GetFillColor(tp, i, cr, cg, cb, ca):
                color = Color(rgb=_hex(cr.value, cg.value, cb.value),
                              alpha=round(ca.value / 255, 4))
            notes = []
            if R.FPDFText_HasUnicodeMapError(tp, i) == 1:
                notes.append("unmapped")
            ox_d, oy_d = geom.apply(ctx.display, ox.value, oy.value)
            box = geom.transform_rect(ctx.display, cl.value, b.value, r.value, t.value)
            out.append(Glyph(id=f"p{ctx.index}_g{i}", bbox=_rbox(box), z=z, text=ch,
                             font=font_id, size=_r(size), color=color, baseline=_r(oy_d),
                             origin_x=_r(ox_d), rotation=_r(angle, 2), render_mode=mode,
                             notes=notes))
    finally:
        R.FPDFText_ClosePage(tp)
    return out


def _font_ref(base: str, flags: int, weight: int, tobj: Any) -> FontRef:
    low = base.lower()
    embedded = False
    if tobj:
        font = R.FPDFTextObj_GetFont(tobj)
        if font:
            embedded = bool(R.FPDFFont_GetIsEmbedded(font))
    if weight <= 0:
        weight = 700 if "bold" in low else 400
    return FontRef(
        id=base,
        pdf_name=base,
        family=clean_family(base),
        bold=weight >= 600 or "bold" in low or "black" in low or "heavy" in low,
        italic=bool(flags & 64) or "italic" in low or "oblique" in low,
        weight=weight,
        flags=flags,
        embedded=embedded,
        subset=bool(_SUBSET_RE.match(base)),
    )


def _wide(buf: ctypes.Array[ctypes.c_char]) -> str:
    return bytes(buf).decode("utf-16-le", "replace").rstrip("\x00")


def _annotations(ctx: _PageCtx) -> list[Annotation]:
    out: list[Annotation] = []
    subtypes = {R.FPDF_ANNOT_LINK: "link", 1: "text", 9: "highlight", 3: "freetext"}
    for i in range(R.FPDFPage_GetAnnotCount(ctx.page)):
        annot = R.FPDFPage_GetAnnot(ctx.page, i)
        if not annot:
            continue
        try:
            st = R.FPDFAnnot_GetSubtype(annot)
            rect = R.FS_RECTF()
            if not R.FPDFAnnot_GetRect(annot, rect):
                continue
            box = geom.transform_rect(ctx.display, rect.left, rect.bottom, rect.right, rect.top)
            uri = None
            dest_page = None
            if st == R.FPDF_ANNOT_LINK:
                link = R.FPDFAnnot_GetLink(annot)
                action = R.FPDFLink_GetAction(link) if link else None
                dest = None
                if action:
                    at = R.FPDFAction_GetType(action)
                    if at == R.PDFACTION_URI:
                        n = R.FPDFAction_GetURIPath(ctx.doc, action, None, 0)
                        buf = ctypes.create_string_buffer(n)
                        R.FPDFAction_GetURIPath(ctx.doc, action, buf, n)
                        uri = buf.value.decode("utf-8", "replace")
                    elif at == R.PDFACTION_GOTO:
                        dest = R.FPDFAction_GetDest(ctx.doc, action)
                elif link:
                    dest = R.FPDFLink_GetDest(ctx.doc, link)
                if dest:
                    idx = R.FPDFDest_GetDestPageIndex(ctx.doc, dest)
                    dest_page = int(idx) if idx >= 0 else None
            n = R.FPDFAnnot_GetStringValue(annot, b"Contents", None, 0)
            contents = ""
            if n > 2:
                cbuf = ctypes.create_string_buffer(n)
                R.FPDFAnnot_GetStringValue(annot, b"Contents",
                                           ctypes.cast(cbuf, ctypes.POINTER(ctypes.c_ushort)), n)
                contents = _wide(cbuf)
            out.append(Annotation(id=f"p{ctx.index}_a{len(out) + 1}", bbox=_rbox(box),
                                  subtype=subtypes.get(st, f"subtype{st}"), uri=uri,
                                  dest_page=dest_page, contents=contents))
        finally:
            R.FPDFPage_CloseAnnot(annot)
    return out


def page_geometry(page: Any) -> tuple[Matrix, float, float, int]:
    """(display matrix, display width, display height, /Rotate degrees)."""
    rot = int(R.FPDFPage_GetRotation(page)) * 90
    left, b, r, t = page.get_cropbox()
    x0, x1 = sorted((float(left), float(r)))
    y0, y1 = sorted((float(b), float(t)))
    w, h = x1 - x0, y1 - y0
    if rot in (90, 270):
        w, h = h, w
    return geom.page_matrix((x0, y0, x1, y1), rot), _r(w), _r(h), rot


def parse_page(doc: Any, index: int, work_dir: Path, fonts: dict[str, FontRef]) -> Page:
    page = doc[index]
    try:
        display, w, h, rot = page_geometry(page)
        ctx = _PageCtx(doc=doc, page=page, index=index, display=display, width=w, height=h,
                       work_dir=work_dir)
        for i in range(R.FPDFPage_CountObjects(page)):
            ctx.top_index = i
            _walk(ctx, R.FPDFPage_GetObject(page, i), geom.IDENTITY)
        if ctx.white_fills:
            _rasterize_pattern_fills(ctx)
        glyphs = _glyphs(ctx, fonts)
        annots = _annotations(ctx)
        out = Page(index=index, width=w, height=h, rotation=rot, elements=ctx.elements,
                   glyphs=glyphs, annotations=annots)
        if ctx.skipped:
            out.triage.reason = "skipped: " + ", ".join(
                f"{k}={v}" for k, v in sorted(ctx.skipped.items()))
            LOG.info("page %d: %s", index, out.triage.reason)
        return out
    finally:
        page.close()


def _failed_page(doc: Any, index: int, exc: BaseException) -> Page:
    w = h = 0.0
    rot = 0
    try:
        page = doc[index]
        try:
            _, w, h, rot = page_geometry(page)
        finally:
            page.close()
    except Exception:
        w, h = 720.0, 540.0
    reason = f"parse error: {type(exc).__name__}: {exc}"
    return Page(index=index, width=w, height=h, rotation=rot,
                triage=Triage(kind="born_digital", reason=reason, confidence=0.0,
                              fallback_image=True))


def parse_pages(pdf: Path, indices: Sequence[int], work_dir: Path) -> tuple[
        list[Page], dict[str, FontRef]]:
    """Parse some pages of ``pdf``. Never raises for a bad page."""
    doc = pdfium.PdfDocument(str(pdf))
    fonts: dict[str, FontRef] = {}
    pages = []
    try:
        for i in indices:
            try:
                pages.append(parse_page(doc, i, work_dir, fonts))
            except Exception as exc:
                LOG.error("%s page %d: parse failed, falling back to a picture: %s",
                          pdf.name, i, exc, exc_info=LOG.isEnabledFor(10))
                pages.append(_failed_page(doc, i, exc))
    finally:
        doc.close()
    return pages, fonts


def _chunks(n: int, k: int) -> list[list[int]]:
    k = max(1, min(k, n))
    return [list(range(n))[i::k] for i in range(k)]


def metadata(pdf: Path) -> tuple[int, dict[str, str]]:
    doc = pdfium.PdfDocument(str(pdf))
    try:
        meta = {k: str(v) for k, v in sorted(doc.get_metadata_dict(skip_empty=True).items())}
        meta["pdf_version"] = str(doc.get_version() or "")
        return len(doc), meta
    finally:
        doc.close()


def parse_document(pdf: Path, work_dir: Path, *, sha256: str = "", workers: int = 1,
                   min_pages_per_worker: int = 4) -> Document:
    """Ingest a whole PDF. Pages are parsed in a process pool when ``workers > 1``."""
    n, meta = metadata(pdf)
    work_dir.mkdir(parents=True, exist_ok=True)
    k = min(workers, max(1, n // min_pages_per_worker))
    pages: list[Page] = []
    fonts: dict[str, FontRef] = {}
    if k <= 1:
        pages, fonts = parse_pages(pdf, range(n), work_dir)
    else:
        with ProcessPoolExecutor(max_workers=k) as pool:
            for ps, fs in pool.map(parse_pages, [pdf] * k, _chunks(n, k), [work_dir] * k):
                pages.extend(ps)
                for key, f in fs.items():
                    fonts.setdefault(key, f)
    pages.sort(key=lambda p: p.index)
    sizes = Counter(f"{p.width:g}x{p.height:g}" for p in pages)
    meta["page_sizes"] = ", ".join(f"{s} ({c})" for s, c in sorted(sizes.items()))
    return Document(source_path=str(pdf), sha256=sha256, metadata=meta, pages=pages,
                    fonts=dict(sorted(fonts.items())), stage="ingest")


# --------------------------------------------------------------------------- cross-check


def crosscheck_pdfplumber(pdf: Path, page: Page) -> dict[str, Any]:
    """Compare our glyphs with pdfplumber's chars on one page; log disagreements.

    Only meaningful for unrotated pages (pdfplumber ignores /Rotate).
    """
    import difflib

    import pdfplumber

    ours = [g for g in page.glyphs if not g.text.isspace()]
    with pdfplumber.open(str(pdf)) as doc:
        pp = doc.pages[page.index]
        theirs = [c for c in pp.chars if not c["text"].isspace()]
    a = "".join(g.text for g in ours)
    b = "".join(str(c["text"]) for c in theirs)
    sm = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
    dists = []
    for blk in sm.get_matching_blocks():
        for k in range(blk.size):
            g, c = ours[blk.a + k], theirs[blk.b + k]
            if g.rotation == 0:
                dists.append(max(abs(g.origin_x - float(c["x0"])), abs(g.size - float(c["size"]))))
    dists.sort()
    res = {
        "page": page.index,
        "ours": len(a),
        "pdfplumber": len(b),
        "text_ratio": round(sm.ratio(), 4) if (a or b) else 1.0,
        "median_dx": round(dists[len(dists) // 2], 3) if dists else None,
        "max_dx": round(dists[-1], 3) if dists else None,
    }
    if page.rotation == 0 and (res["text_ratio"] < 0.98 or (res["max_dx"] or 0) > 0.5):
        LOG.warning("%s p%d: pdfium vs pdfplumber disagree: %s", pdf.name, page.index, res)
    return res
