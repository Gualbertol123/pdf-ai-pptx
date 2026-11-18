"""OCR stage (M8): hard pages become IR glyphs that the M4 text pipeline groups.

Runs only on pages triage flagged (PLAN.md §8 rule 2):

* **scanned**: render at the scan's resolution, deskew, OCR (PP-OCR on ONNX, or
  Windows OCR when the models are missing), then turn every read word into glyph
  records (``source="ocr"``, confidence = the recogniser's) placed with the
  installed font whose word widths fit best. The page picture becomes a clean
  background: the text pixels are masked and inpainted (``ml/inpaint.py``).
  Layout regions (``ml/layout.py``) are stored for the text grouping and roles;
  text inside formula / picture / chart regions stays in the picture.
* **broken_text**: the glyphs are in the right place but their Unicode is wrong.
  The page render is OCR'd and each line's PDF glyphs take the OCR text (one to
  one when the counts agree, else replaced by OCR glyphs with the PDF's font size).
* **outlined_text**: clusters of letter-shaped filled paths are rendered, turned
  upright along their main direction and OCR'd; they are replaced by glyphs only
  when the OCR confidence is high.

Everything is optional: no engine, no model or any failure leaves the page as it
was (picture / paths), with the reason in ``triage.reason``.
"""

from __future__ import annotations

import itertools
import json
import math
import statistics
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import cv2
import numpy as np

from .. import log
from ..ingest.render import ensure_render
from ..ir import Color, Document, FontRef, Glyph, Image, Page, Shape
from . import fontfit, inpaint, layout, models, ocr_engine, runtime, winocr
from .ocr_engine import OcrChar, OcrLine

LOG = log.get(__name__)

OCR_VERSION = "8.1"
KINDS = ("scanned", "broken_text", "outlined_text")
SCAN_DPI_MIN, SCAN_DPI_MAX = 150, 300
OUTLINE_DPI = 300
MIN_WORD_CONF = 0.35  # words read below this stay in the picture
MIN_LINE_CONF = 0.5
OUTLINE_MIN_CONF = 0.9  # outlined text is replaced only above this
BULLET_CONF = 0.6  # confidence given to a bullet recognised by its shape
SIZE_CLUSTER = 0.045  # line sizes closer than this (relative) are one size
SNAP_WHOLE, SNAP_HALF = 0.03, 0.015
WIDTH_FIT_OK = 0.05  # relative residual of the word-width size fit that is trusted
INK_BIAS = 1.012  # thresholded scan ink is ~1.2% narrower than the font's (measured, s01)
# stroke width / x-height (measured M8: regular Arial / Calibri 0.149-0.166, bold
# Arial 0.21-0.215); bold when above BOLD_XH or BOLD_REL x the page's median line
BOLD_XH = 0.19
BOLD_REL = 1.22
SKEW_SNAP = 0.1  # degrees: smaller skews are treated as none
KEEP_IN_PICTURE = {"display_formula", "inline_formula", "formula_number", "image", "chart",
                   "seal", "header_image", "footer_image"}
FORMULA_LABELS = {"display_formula", "inline_formula"}
META_KEY = "pdf2pptx:ocr"
NOTE = "ocr"


def engine_id() -> str:
    """Which OCR engine this machine would use (part of the stage cache key)."""
    if ocr_engine.ppocr_available() and runtime.ort_module() is not None:
        return "ppocr:" + models.fingerprint("ocr_det", "ocr_cls", "ocr_rec_latin",
                                             "ocr_rec_multi", "layout", "inpaint")
    if winocr.available():
        return "winocr"
    return "none"


def stage_version() -> str:
    return f"{OCR_VERSION}+{engine_id()}"


# --------------------------------------------------------------------------- images


def read_bgr(path: Path) -> np.ndarray:
    data = np.fromfile(str(path), dtype=np.uint8)  # cv2.imread fails on non-ASCII paths
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"cannot read {path}")
    return img


def write_image(path: Path, img: np.ndarray, quality: int = 88) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    ext = path.suffix.lower()
    params = [cv2.IMWRITE_JPEG_QUALITY, quality] if ext in (".jpg", ".jpeg") else []
    ok, buf = cv2.imencode(ext, img, params)
    if not ok:
        raise ValueError(f"cannot encode {path}")
    tmp = path.with_name(path.name + ".tmp")
    buf.tofile(str(tmp))
    tmp.replace(path)


@dataclass
class Frame:
    """Deskewed image pixels <-> page points."""

    dpi: float
    skew: float  # degrees, counter-clockwise; the deskewed image = image rotated by -skew
    w: int
    h: int
    m: np.ndarray = field(default_factory=lambda: np.eye(2, 3))  # image -> deskewed
    minv: np.ndarray = field(default_factory=lambda: np.eye(2, 3))  # deskewed -> image
    origin: tuple[float, float] = (0.0, 0.0)  # page pt of image pixel (0, 0)
    angle: float = 0.0  # extra rotation of a crop (outlined text), degrees ccw

    @classmethod
    def make(cls, dpi: float, skew: float, w: int, h: int,
             origin: tuple[float, float] = (0.0, 0.0)) -> Frame:
        m = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), -skew, 1.0)
        return cls(dpi=dpi, skew=skew, w=w, h=h, m=m, minv=cv2.invertAffineTransform(m),
                   origin=origin)

    def to_pt(self, x: float, y: float) -> tuple[float, float]:
        ix = self.minv[0, 0] * x + self.minv[0, 1] * y + self.minv[0, 2]
        iy = self.minv[1, 0] * x + self.minv[1, 1] * y + self.minv[1, 2]
        k = 72.0 / self.dpi
        return self.origin[0] + ix * k, self.origin[1] + iy * k

    def from_pt(self, x: float, y: float) -> tuple[float, float]:
        k = self.dpi / 72.0
        ix, iy = (x - self.origin[0]) * k, (y - self.origin[1]) * k
        return (self.m[0, 0] * ix + self.m[0, 1] * iy + self.m[0, 2],
                self.m[1, 0] * ix + self.m[1, 1] * iy + self.m[1, 2])

    @property
    def rotation(self) -> float:
        return round(self.skew, 2)


# --------------------------------------------------------------------------- line geometry


@dataclass
class WordGeo:
    text: str
    x0: float  # ink extent, deskewed px
    x1: float
    conf: float


@dataclass
class LineGeo:
    line: OcrLine
    baseline: float  # deskewed px
    x_height: float
    words: list[WordGeo]
    rgb: str
    stroke: float  # px
    ink_top: float
    ink_bot: float
    bullet: bool = False
    size_px: float = 0.0
    bold: bool = False
    family: str = fontfit.DEFAULT_FAMILY
    region: str | None = None  # layout label of the region holding the line
    thresh: float = 128.0  # ink threshold of the line crop
    dark: bool = True  # dark ink on a light background

    @property
    def conf(self) -> float:
        return self.line.conf


def _ink(gray: np.ndarray) -> tuple[np.ndarray, float]:
    """Ink mask (dark on light, or light on dark) of a crop and its threshold."""
    t, mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    ink = mask > 0
    if ink.mean() > 0.6:  # light text on a dark background
        ink = ~ink
    return ink, float(t)


def _color(bgr: np.ndarray, gray: np.ndarray, ink: np.ndarray) -> str:
    vals = gray[ink]
    if vals.size < 4:
        return "000000"
    dark = ink.mean() <= 0.6 and float(np.median(gray[ink])) < float(np.median(gray[~ink]) if (
        ~ink).any() else 255)
    q = np.percentile(vals, 35 if dark else 65)
    core = ink & ((gray <= q) if dark else (gray >= q))
    px = bgr[core] if core.any() else bgr[ink]
    b, g, r = (round(float(np.median(px[:, k]))) for k in range(3))
    return f"{r:02X}{g:02X}{b:02X}"


def _runs(cols: np.ndarray) -> list[tuple[int, int]]:
    out, start = [], None
    for i, v in enumerate(cols):
        if v and start is None:
            start = i
        elif not v and start is not None:
            out.append((start, i))
            start = None
    if start is not None:
        out.append((start, len(cols)))
    return out


def _split_words(chars: list[OcrChar]) -> list[list[OcrChar]]:
    words: list[list[OcrChar]] = [[]]
    for c in chars:
        if c.ch.isspace():
            if words[-1]:
                words.append([])
        else:
            words[-1].append(c)
    return [w for w in words if w]


def analyse_line(bgr: np.ndarray, gray: np.ndarray, line: OcrLine) -> LineGeo | None:
    h_img, w_img = gray.shape
    x0, y0, x1, y1 = line.box
    ix0, iy0 = max(0, math.floor(x0)), max(0, math.floor(y0))
    ix1, iy1 = min(w_img, math.ceil(x1)), min(h_img, math.ceil(y1))
    if ix1 - ix0 < 2 or iy1 - iy0 < 2:
        return None
    g = gray[iy0:iy1, ix0:ix1]
    ink, thresh = _ink(g)
    n, lab, st, _ = cv2.connectedComponentsWithStats(ink.astype(np.uint8), connectivity=8)
    hgt = iy1 - iy0
    min_area = max(3, int((0.04 * hgt) ** 2))
    comps = [st[i] for i in range(1, n) if st[i][4] >= min_area]
    if not comps:
        return None
    clean = np.isin(lab, [i for i in range(1, n) if st[i][4] >= min_area])
    tall = sorted(c[3] for c in comps)
    ref_h = tall[int(0.8 * (len(tall) - 1))]
    body = [c for c in comps if c[3] >= 0.35 * ref_h] or comps
    baseline_c = float(statistics.median(c[1] + c[3] for c in body))
    near = [c for c in body if abs(c[1] + c[3] - baseline_c) <= 0.12 * ref_h] or body
    xh = float(statistics.median(baseline_c - c[1] for c in near))
    ink_top = float(min(c[1] for c in comps))
    ink_bot = float(max(c[1] + c[3] for c in comps))
    # stroke width ~ 2 * area / perimeter over the letter bodies
    contours, _ = cv2.findContours(clean.astype(np.uint8), cv2.RETR_LIST,
                                   cv2.CHAIN_APPROX_NONE)
    perim = sum(cv2.arcLength(c, True) for c in contours)
    stroke = 2.0 * float(clean.sum()) / max(perim, 1.0)
    cols = clean.any(axis=0)
    runs = _runs(cols)
    words: list[WordGeo] = []
    groups = _split_words(line.chars)
    # every ink run belongs to the word whose stretch holds its centre; stretches meet
    # half-way between one word's last character and the next word's first (the CTC
    # positions are good enough for that, not for the exact ink ends)
    mids = [(a[-1].x1 + b[0].x0) / 2 - ix0 for a, b in itertools.pairwise(groups)]
    cuts = [-math.inf, *mids, math.inf]
    for k, wc in enumerate(groups):
        sel = [r for r in runs if cuts[k] <= (r[0] + r[1]) / 2 < cuts[k + 1]]
        if sel:
            wx0, wx1 = sel[0][0] + ix0, sel[-1][1] + ix0
        else:
            wx0, wx1 = wc[0].x0, wc[-1].x1
        words.append(WordGeo("".join(c.ch for c in wc), float(wx0), float(wx1),
                             float(np.mean([c.conf for c in wc]))))
    bullet = False
    if len(words) == 1 and len(words[0].text) <= 1 and len(body) == 1:
        c = body[0]
        fill = c[4] / max(c[2] * c[3], 1)
        bullet = fill > 0.6 and 0.6 <= c[2] / max(c[3], 1) <= 1.6
    return LineGeo(line=line, baseline=baseline_c + iy0, x_height=max(xh, 1.0), words=words,
                   rgb=_color(bgr[iy0:iy1, ix0:ix1], g, clean), stroke=stroke,
                   ink_top=ink_top + iy0, ink_bot=ink_bot + iy0, bullet=bullet,
                   thresh=thresh, dark=ink.mean() <= 0.5 or bool((g[ink] < thresh).mean() > 0.5))


def _attach_small_lines(geos: list[LineGeo]) -> None:
    """A bullet or one-letter line next to a longer line on the same row takes that
    line's baseline (a bullet dot sits at mid x-height, not on the baseline)."""
    for g in geos:
        if sum(len(w.text) for w in g.words) > 2:
            continue
        h = g.line.box[3] - g.line.box[1]
        best = None
        for o in geos:
            if o is g or sum(len(w.text) for w in o.words) <= 2:
                continue
            ov = min(g.line.box[3], o.line.box[3]) - max(g.line.box[1], o.line.box[1])
            dx = o.line.box[0] - g.line.box[2]
            if ov >= 0.4 * min(h, o.line.box[3] - o.line.box[1]) and -h <= dx <= 4 * h and (
                    best is None or dx < best[0]):
                best = (dx, o)
        if best is not None:
            o = best[1]
            g.baseline, g.x_height = o.baseline, o.x_height
            g.stroke, g.size_px, g.bold = o.stroke, o.size_px, o.bold
            if g.bullet:  # a round blob before a text line: a shape rule, not the reader
                g.words[0].text = "•"
                g.words[0].conf = max(g.words[0].conf, BULLET_CONF)


def _snap(size: float) -> float:
    """OCR sizes carry ~2% noise: prefer a whole point, then a half point."""
    whole = round(size)
    if whole and abs(whole - size) <= SNAP_WHOLE * size:
        return float(whole)
    half = round(size * 2) / 2
    if abs(half - size) <= SNAP_HALF * size:
        return half
    return round(size, 2)


def unify_sizes(geos: list[LineGeo], k: float) -> None:
    """Lines whose sizes differ by less than ``SIZE_CLUSTER`` share one (character
    weighted median) size, snapped to a whole / half point. ``k``: pt per px."""
    items = sorted((g.size_px * k, sum(len(w.text) for w in g.words), i)
                   for i, g in enumerate(geos) if g.size_px > 0 and not g.bullet)
    groups: list[list[tuple[float, int, int]]] = []
    for it in items:
        if groups and it[0] <= groups[-1][-1][0] * (1 + SIZE_CLUSTER):
            groups[-1].append(it)
        else:
            groups.append([it])
    for grp in groups:
        weighted = sorted(v for v, n, _ in grp for _ in range(max(n, 1)))
        size = _snap(weighted[len(weighted) // 2])
        for _, _, i in grp:
            geos[i].size_px = size / k


def find_bullets(gray: np.ndarray, geos: list[LineGeo]) -> int:
    """Bullet dots / squares / dashes left of a line that the detector skipped (small
    blobs), found by their shape. Returns how many were added."""
    h, _w = gray.shape
    added = 0
    for g in geos:
        if g.bullet or not g.words or g.size_px <= 0:
            continue
        first = g.words[0]
        em, xh = g.size_px, g.x_height
        if len(first.text) == 1 and first.text in "-—−–" and len(g.words) > 1:
            first.text = "–"  # a dash label read as hyphen / em dash
            continue
        if len(first.text) == 1 and first.text in "•·●▪■":
            continue
        x0, x1 = int(max(0, first.x0 - 2.2 * em)), int(max(0, first.x0 - 0.12 * em))
        y0, y1 = int(max(0, g.baseline - 1.3 * xh)), int(min(h, g.baseline + 0.2 * xh))
        if x1 - x0 < 3 or y1 - y0 < 3:
            continue
        win = gray[y0:y1, x0:x1]
        ink = (win < g.thresh) if g.dark else (win > g.thresh)
        n, _, st, _ = cv2.connectedComponentsWithStats(ink.astype(np.uint8), connectivity=8)
        best = None
        for i in range(1, n):
            cx, cy, cw, ch, area = (int(v) for v in st[i])
            if area < max(4, (0.1 * xh) ** 2) or cx + cw >= ink.shape[1] - 1 or cx == 0:
                continue
            mid = y0 + cy + ch / 2.0
            fill = area / max(cw * ch, 1)
            kind = None
            if (0.6 <= cw / max(ch, 1) <= 1.6 and 0.15 * xh <= ch <= 0.9 * xh and fill > 0.55
                    and g.baseline - 0.95 * xh <= mid <= g.baseline - 0.05 * xh):
                kind = "▪" if fill > 0.9 else "•"
            elif (ch <= 0.3 * xh and cw / max(ch, 1) >= 2.0 and 0.25 * em <= cw <= 1.3 * em
                  and fill > 0.7 and g.baseline - 0.8 * xh <= mid <= g.baseline - 0.2 * xh):
                kind = "–"
            if kind and (best is None or cx > best[0]):
                best = (cx, kind, cw)
        if best is None:
            continue
        bx0 = float(x0 + best[0])
        g.words.insert(0, WordGeo(best[1], bx0, bx0 + best[2], BULLET_CONF))
        added += 1
    return added


def fit_fonts(geos: list[LineGeo], family: str | None = None) -> dict[str, Any]:
    """Size (px) and bold per line; one family per page from the word widths."""
    ratios = [(g.stroke / max(g.x_height, 1.0), sum(len(w.text) for w in g.words))
              for g in geos]
    weighted = sorted(r for r, n in ratios for _ in range(max(n, 1)))
    med = weighted[len(weighted) // 2] if weighted else 0.0
    for g, (r, _) in zip(geos, ratios, strict=True):
        g.bold = r >= BOLD_XH or (med > 0 and r >= BOLD_REL * med)
    obs = [([fontfit.WordObs(w.text, w.x1 - w.x0) for w in g.words if w.x1 > w.x0], g.bold)
           for g in geos if not g.bullet]
    scores: dict[str, float] = {}
    if family is None:
        family, scores = fontfit.choose_family(obs)
    for g in geos:
        g.family = family
        f = fontfit.face(family, g.bold) or fontfit.face(family)
        words = [fontfit.WordObs(w.text, w.x1 - w.x0) for w in g.words if w.x1 > w.x0]
        size_x = g.x_height / f.x_height if f is not None and f.x_height > 0 else g.x_height * 2
        size = size_x
        if f is not None and words and not g.bullet:
            s, r = fontfit.fit_size(f, words)
            # the width fit is the better estimate; the x-height (fooled by ascender-
            # heavy text) only vetoes a fit that is poor and far from it
            if s > 0 and ((len(words) >= 2 and r < WIDTH_FIT_OK)
                          or abs(math.log(s / size_x)) < math.log(1.3)):
                size = s
        g.size_px = size * INK_BIAS
    _attach_small_lines(geos)
    return {"family": family, "family_scores": scores}


# --------------------------------------------------------------------------- glyphs


def font_ref(doc: Document, family: str, bold: bool) -> str:
    fid = f"ocr:{family}{'-Bold' if bold else ''}"
    if fid not in doc.fonts:
        doc.fonts[fid] = FontRef(id=fid, pdf_name=f"{family}{'-Bold' if bold else ''}",
                                 family=family, bold=bold, weight=700 if bold else 400)
    return fid


def make_glyphs(doc: Document, page: Page, geos: list[LineGeo], fr: Frame, *,
                engine: str, start: int = 0, size_override: dict[int, float] | None = None,
                color_override: str | None = None, font_override: str | None = None,
                ) -> list[Glyph]:
    """Glyph records for the analysed lines (one per character, spaces included)."""
    out: list[Glyph] = []
    k = 72.0 / fr.dpi
    rot = round(fr.rotation + fr.angle, 2) % 360.0
    n = start
    for li, g in enumerate(geos):
        f = fontfit.face(g.family, g.bold) or fontfit.face(g.family)
        size_pt = (size_override or {}).get(li, g.size_px * k)
        size_pt = round(size_pt, 2)
        s_px = size_pt / k
        fid = font_override or font_ref(doc, g.family, g.bold)
        rgb = color_override or g.rgb
        words = [w for w in g.words if w.conf >= MIN_WORD_CONF and w.text]
        prev_end: float | None = None
        for w in words:
            geo = [(c, f.char(c) if f else None) for c in w.text]
            ink_em = f.ink_em(w.text) if f else len(w.text) * fontfit.FALLBACK_EM
            first = geo[0][1]
            lsb = first.x0 if first is not None else 0.05
            scale = (w.x1 - w.x0) / max(ink_em * s_px, 1e-6) if w.x1 > w.x0 else 1.0
            scale = min(max(scale, 0.7), 1.4)
            x = w.x0 - lsb * s_px * scale
            if prev_end is not None and x > prev_end:  # the space between words
                out.append(_glyph(page, n, " ", prev_end, x - prev_end, g, None, s_px, fr,
                                  size_pt, fid, rgb, rot, w.conf, engine))
                n += 1
            for ch, cg in geo:
                adv = (cg.adv if cg is not None else fontfit.FALLBACK_EM) * s_px * scale
                out.append(_glyph(page, n, ch, x, adv, g, cg, s_px, fr, size_pt, fid, rgb,
                                  rot, w.conf, engine))
                n += 1
                x += adv
            prev_end = x
    return out


def _glyph(page: Page, n: int, ch: str, x: float, adv: float, g: LineGeo,
           cg: fontfit.CharGeo | None, s_px: float, fr: Frame, size_pt: float, fid: str,
           rgb: str, rot: float, conf: float, engine: str) -> Glyph:
    if cg is None or ch.isspace():
        ix0, ix1, iy0, iy1 = x, x + adv, g.baseline - 0.45 * s_px, g.baseline
    else:
        ix0, ix1 = x + cg.x0 * s_px, x + cg.x1 * s_px
        iy0, iy1 = g.baseline - cg.y1 * s_px, g.baseline - cg.y0 * s_px
    corners = [fr.to_pt(px, py) for px, py in ((ix0, iy0), (ix1, iy0), (ix0, iy1), (ix1, iy1))]
    xs, ys = [c[0] for c in corners], [c[1] for c in corners]
    ox, oy = fr.to_pt(x, g.baseline)
    return Glyph(id=f"p{page.index}_o{n}", bbox=(_r(min(xs)), _r(min(ys)), _r(max(xs)),
                                                  _r(max(ys))),
                 text=ch, font=fid, size=size_pt, color=Color(rgb=rgb), baseline=_r(oy),
                 origin_x=_r(ox), rotation=rot, confidence=round(conf, 3), source="ocr",
                 notes=[f"{NOTE}:{engine}"])


def _r(v: float) -> float:
    return round(float(v), 2)


# --------------------------------------------------------------------------- engines


def read_lines(img: np.ndarray) -> tuple[list[OcrLine], str]:
    """OCR one (deskewed) image with the best available engine."""
    if ocr_engine.ppocr_available() and runtime.ort_module() is not None:
        try:
            return ocr_engine.ppocr(img), "ppocr"
        except Exception as exc:
            LOG.warning("PP-OCR failed (%s: %s); trying Windows OCR", type(exc).__name__, exc)
    if winocr.available():
        return winocr.recognise(img), "winocr"
    raise RuntimeError("no OCR engine: run `pdf2pptx setup models --lite` "
                       "(or install a Windows OCR language)")


def _scan_dpi(page: Page) -> float:
    best = None
    for e in page.elements:
        if isinstance(e, Image) and e.pixel_size[0] > 0:
            w = e.bbox[2] - e.bbox[0]
            if w > 0 and (best is None or w > best[0]):
                best = (w, e.pixel_size[0] / (w / 72.0))
    dpi = best[1] if best else 200.0
    return float(min(max(round(dpi), SCAN_DPI_MIN), SCAN_DPI_MAX))


def _page_image(pdf: Path, work_dir: Path, page: Page, dpi: int) -> np.ndarray:
    return read_bgr(ensure_render(pdf, work_dir, page.index, dpi))


def _region_of(g: LineGeo, regions: list[layout.Region]) -> tuple[str | None, float]:
    """(label, reading rank) of the region holding the line; unplaced lines rank by
    the nearest region above them."""
    x0, y0, x1, y1 = g.line.box
    area = max((x1 - x0) * (y1 - y0), 1.0)
    best = None
    for r in regions:
        ov = max(0.0, min(x1, r.box[2]) - max(x0, r.box[0])) * max(
            0.0, min(y1, r.box[3]) - max(y0, r.box[1]))
        if ov >= 0.5 * area and (best is None or ov > best[0]):
            best = (ov, r)
    if best is not None:
        return best[1].label, float(best[1].order)
    above = [r for r in regions if r.box[3] <= y1 and r.box[0] < x1 and r.box[2] > x0]
    if above:
        return None, max(r.order for r in above) + 0.5
    return None, -1.0 if not regions else float(min(r.order for r in regions)) - 0.5


# --------------------------------------------------------------------------- page kinds


def ocr_scanned(doc: Document, page: Page, pdf: Path, work_dir: Path) -> dict[str, Any]:
    t0 = time.perf_counter()
    dpi = int(_scan_dpi(page))
    img = _page_image(pdf, work_dir, page, dpi)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    skew = ocr_engine.estimate_skew(gray)
    if abs(skew) < SKEW_SNAP:
        skew = 0.0
    fr = Frame.make(dpi, skew, img.shape[1], img.shape[0])
    paper = int(np.median(gray))
    de = img if skew == 0.0 else ocr_engine.rotate_image(img, -skew, (paper,) * 3)
    de_gray = cv2.cvtColor(de, cv2.COLOR_BGR2GRAY)
    t_ocr = time.perf_counter()
    lines, engine = read_lines(de)
    t_ocr = time.perf_counter() - t_ocr
    regions = layout.detect(de) if layout.available() else []
    geos = [g for g in (analyse_line(de, de_gray, ln) for ln in lines) if g is not None]
    kept, in_picture = [], []
    order: dict[int, tuple[float, ...]] = {}
    for g in geos:
        g.region, rank = _region_of(g, regions)
        order[id(g)] = (rank, g.line.box[1], g.line.box[0])
        if g.region in KEEP_IN_PICTURE or g.line.flipped:  # upside-down text: a gap
            in_picture.append(g)
        elif g.conf >= MIN_LINE_CONF or g.bullet:
            kept.append(g)
    # glyphs go out in reading order (the layout model's region order, then top-down),
    # as a born-digital PDF's content stream would have them
    kept.sort(key=lambda g: order[id(g)])
    fonts = fit_fonts(kept)
    unify_sizes(kept, 72.0 / dpi)
    bullets = find_bullets(de_gray, kept)
    glyphs = make_glyphs(doc, page, kept, fr, engine=engine)
    page.glyphs.extend(glyphs)
    # the clean background: text pixels masked (in the original, skewed image) and filled
    t_inp = time.perf_counter()
    mask = inpaint.text_mask(de_gray, kept, fr, img.shape[:2])
    bg, how = inpaint.fill(img, mask)
    t_inp = time.perf_counter() - t_inp
    rel = f"ocr/p{page.index:04d}_background.jpg"
    write_image(work_dir / rel, bg)
    _set_background(page, work_dir, rel, bg)
    regions_pt = [r.to_page(fr) for r in regions]
    words = [w for g in kept for w in g.words]
    formulas = sum(1 for r in regions if r.label in FORMULA_LABELS)
    info = {
        "kind": "scanned", "engine": engine, "dpi": dpi, "skew": skew,
        "lines": len(kept), "words": len(words),
        "low_conf_words": sum(1 for w in words if w.conf < MIN_WORD_CONF),
        "mean_conf": round(float(np.mean([w.conf for w in words])), 4) if words else None,
        "kept_in_picture": len(in_picture), "formula_regions": formulas,
        "bullets_by_shape": bullets,
        "family": fonts["family"], "family_scores": fonts["family_scores"],
        "inpaint": how, "regions": regions_pt,
        "seconds": {"ocr": round(t_ocr, 3), "inpaint": round(t_inp, 3),
                    "total": round(time.perf_counter() - t0, 3)},
    }
    if formulas:
        info["gaps"] = [f"{formulas} formula region(s) kept as picture (no math model)"]
    return info


def _set_background(page: Page, work_dir: Path, rel: str, bg: np.ndarray) -> None:
    """The cleaned render replaces the page pictures (it already contains them)."""
    import hashlib

    data = (work_dir / rel).read_bytes()
    scan = Image(id=f"p{page.index}_ocrbg", bbox=(0.0, 0.0, page.width, page.height), z=0,
                 file=rel, sha256=hashlib.sha256(data).hexdigest(),
                 pixel_size=(int(bg.shape[1]), int(bg.shape[0])), is_background=True,
                 source="model:ocr+inpaint", alt_text="scanned page (text removed)",
                 notes=["ocr_background"])
    keep = [e for e in page.elements if not isinstance(e, Image)]
    page.elements = [scan, *keep]


def ocr_broken(doc: Document, page: Page, pdf: Path, work_dir: Path) -> dict[str, Any]:
    """Re-read a page whose text layer decodes to garbage."""
    t0 = time.perf_counter()
    dpi = 300
    img = _page_image(pdf, work_dir, page, dpi)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    fr = Frame.make(dpi, 0.0, img.shape[1], img.shape[0])
    lines, engine = read_lines(img)
    geos = [g for g in (analyse_line(img, gray, ln) for ln in lines) if g is not None]
    vis = [g for g in page.glyphs if g.render_mode not in (3, 7) and not g.text.isspace()]
    replaced = remapped = 0
    new: list[Glyph] = []
    drop: set[str] = set()
    for g in geos:
        if g.conf < MIN_LINE_CONF:
            continue
        x0, y0 = fr.to_pt(g.line.box[0], g.line.box[1])
        x1, y1 = fr.to_pt(g.line.box[2], g.line.box[3])
        inside = [p for p in vis if p.id not in drop and x0 <= (p.bbox[0] + p.bbox[2]) / 2 <= x1
                  and y0 <= (p.bbox[1] + p.bbox[3]) / 2 <= y1]
        if not inside:
            continue
        chars = [c for c in g.line.text if not c.isspace()]
        inside.sort(key=lambda p: p.origin_x)
        if len(chars) == len(inside):  # same count: keep the PDF geometry, fix the text
            for p, c in zip(inside, chars, strict=True):
                if p.text != c:
                    p.text = c
                    p.source = "ocr"
                    p.confidence = round(g.conf, 3)
                    p.notes = [*[x for x in p.notes if x != "unmapped"], f"{NOTE}:{engine}"]
            remapped += 1
            continue
        size = statistics.median(p.size for p in inside)
        font = Counter_most([p.font for p in inside])
        color = inside[0].color.rgb or "000000"
        g.family = doc.fonts[font].family if font in doc.fonts else fontfit.DEFAULT_FAMILY
        g.size_px = size * dpi / 72.0
        made = make_glyphs(doc, page, [g], fr, engine=engine, start=len(page.glyphs) + len(new),
                           size_override={0: size}, color_override=color, font_override=font)
        new.extend(made)
        drop.update(p.id for p in inside)
        replaced += 1
    if drop:
        page.glyphs = [p for p in page.glyphs if p.id not in drop]
    page.glyphs.extend(new)
    return {"kind": "broken_text", "engine": engine, "lines": len(geos),
            "lines_remapped": remapped, "lines_replaced": replaced,
            "seconds": {"total": round(time.perf_counter() - t0, 3)}}


def Counter_most(items: list[str]) -> str:
    counts: dict[str, int] = {}
    for it in items:
        counts[it] = counts.get(it, 0) + 1
    return max(sorted(counts), key=lambda k: counts[k])


# --------------------------------------------------------------------------- outlined text


def _letter_shapes(page: Page) -> list[Shape]:
    """Filled, curved shapes whose sub-paths are all letter-sized (text as outlines)."""
    from ..ingest.triage import LETTER_MAX

    out = []
    for e in page.elements:
        if not isinstance(e, Shape) or e.path is None or e.fill.kind != "solid" or (
                e.stroke is not None and e.stroke.color is not None):
            continue
        w, h = e.bbox[2] - e.bbox[0], e.bbox[3] - e.bbox[1]
        if max(w, h) > 0.6 * max(page.width, page.height) or min(w, h) <= 0:
            continue
        subs, cur, curved = [], [], False
        for seg in [*e.path, None]:
            if seg is None or seg.op == "M":
                if cur:
                    subs.append((cur, curved))
                cur, curved = [], False
            if seg is None:
                break
            cur.extend(seg.pts)
            curved = curved or seg.op == "C"
        if not subs:
            continue
        sizes = [max(max(p[0] for p in s) - min(p[0] for p in s),
                     max(p[1] for p in s) - min(p[1] for p in s)) for s, _ in subs]
        if (sum(1 for s, (_, c) in zip(sizes, subs, strict=True) if c) >= 0.4 * len(subs)
                and max(sizes) <= LETTER_MAX and len(subs) >= 2):
            out.append(e)
    return out


def _clusters(shapes: list[Shape]) -> list[list[Shape]]:
    """Shapes close to each other (gap below their height) belong to one label."""
    n = len(shapes)
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i in range(n):
        a = shapes[i].bbox
        for j in range(i + 1, n):
            b = shapes[j].bbox
            tol = 0.8 * max(min(a[3] - a[1], a[2] - a[0]), min(b[3] - b[1], b[2] - b[0]), 4.0)
            if (b[0] - a[2] <= tol and a[0] - b[2] <= tol and b[1] - a[3] <= tol
                    and a[1] - b[3] <= tol):
                parent[find(i)] = find(j)
    groups: dict[int, list[Shape]] = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(shapes[i])
    return [sorted(g, key=lambda s: s.id) for _, g in sorted(groups.items())]


def _points(shapes: list[Shape]) -> np.ndarray:
    return np.array([p for s in shapes for seg in (s.path or []) for p in seg.pts], np.float32)


def ocr_outlined(doc: Document, page: Page, pdf: Path, work_dir: Path) -> dict[str, Any]:
    t0 = time.perf_counter()
    shapes = _letter_shapes(page)
    if not shapes:
        return {"kind": "outlined_text", "clusters": 0, "replaced": 0}
    dpi = OUTLINE_DPI
    full = _page_image(pdf, work_dir, page, dpi)
    # one label is too little to tell fonts apart: use the document's main family
    doc_family = dominant_family(doc)
    k = dpi / 72.0
    replaced, tried, removed = 0, 0, set()
    details = []
    engine = "none"
    for cl in _clusters(shapes):
        pts = _points(cl)
        if len(pts) < 8:
            continue
        (cx, cy), (rw, rh), ang = cv2.minAreaRect(pts)
        if rw < rh:  # the long side is the text direction
            rw, rh, ang = rh, rw, ang + 90.0
        # minAreaRect angle is clockwise on screen (y down); glyph rotation is ccw
        ccw = round(-ang, 2)
        while ccw <= -90.0:
            ccw += 180.0
        while ccw > 90.0:
            ccw -= 180.0
        if abs(ccw) < 0.5:
            ccw = 0.0
        pad = 0.6 * rh + 4
        side = math.ceil((max(rw, rh) + 2 * pad) * k)
        ox, oy = cx - side / (2 * k), cy - side / (2 * k)
        crop = _crop_pt(full, ox, oy, side)
        crop = _only_shapes(crop, cl, ox, oy, k)
        up = ocr_engine.rotate_image(crop, -ccw, (255, 255, 255))
        tried += 1
        try:
            lines, engine = read_lines(up)
        except Exception as exc:
            details.append({"shapes": [s.id for s in cl], "error": str(exc)})
            continue
        upg = cv2.cvtColor(up, cv2.COLOR_BGR2GRAY)
        geos = [g for g in (analyse_line(up, upg, ln) for ln in lines) if g is not None]
        conf = min((g.conf for g in geos), default=0.0)
        text = " / ".join(g.line.text for g in geos)
        ok = bool(geos) and conf >= OUTLINE_MIN_CONF
        details.append({"shapes": [s.id for s in cl], "text": text, "conf": round(conf, 3),
                        "angle": ccw, "replaced": ok})
        if not ok:
            continue
        fr = Frame.make(dpi, ccw, side, side, origin=(ox, oy))
        fit_fonts(geos, family=doc_family)
        fill = cl[0].fill.color.rgb if cl[0].fill.color and cl[0].fill.color.rgb else None
        glyphs = make_glyphs(doc, page, geos, fr, engine=engine, start=len(page.glyphs),
                             color_override=fill)
        for gl in glyphs:
            gl.z = min(s.z for s in cl)
        page.glyphs.extend(glyphs)
        removed.update(s.id for s in cl)
        replaced += 1
    if removed:
        page.elements = [e for e in page.elements if e.id not in removed]
    return {"kind": "outlined_text", "engine": engine, "clusters": tried, "replaced": replaced,
            "shapes_removed": len(removed), "labels": details,
            "seconds": {"total": round(time.perf_counter() - t0, 3)}}


def _only_shapes(crop: np.ndarray, shapes: list[Shape], ox: float, oy: float,
                 k: float) -> np.ndarray:
    """White out everything but the cluster's own paths (neighbouring text, lines)."""
    mask = np.zeros(crop.shape[:2], np.uint8)
    for s in shapes:
        sub: list[tuple[float, float]] = []
        for seg in [*(s.path or []), None]:
            if (seg is None or seg.op == "M") and len(sub) >= 3:
                poly = np.array([[(x - ox) * k, (y - oy) * k] for x, y in sub], np.int32)
                cv2.fillPoly(mask, [poly], 255)
            if seg is None:
                break
            sub = list(seg.pts) if seg.op == "M" else [*sub, *seg.pts]
    mask = cv2.dilate(mask, np.ones((5, 5), np.uint8))
    out = crop.copy()
    out[mask == 0] = 255
    return out


def dominant_family(doc: Document) -> str | None:
    """The installed family of most real (non-OCR) glyphs in the document."""
    from ..fonts.match import match_font

    counts: dict[str, int] = {}
    for p in doc.pages:
        for g in p.glyphs:
            fr = doc.fonts.get(g.font)
            if g.source == "ocr" or fr is None or g.text.isspace():
                continue
            counts[g.font] = counts.get(g.font, 0) + 1
    fams: dict[str, int] = {}
    for fid, n in counts.items():
        try:
            fam = match_font(doc.fonts[fid]).family
        except Exception:
            fam = doc.fonts[fid].family
        if fontfit.face(fam) is not None:
            fams[fam] = fams.get(fam, 0) + n
    return max(sorted(fams), key=lambda k: fams[k]) if fams else None


def _crop_pt(img: np.ndarray, ox: float, oy: float, side: int) -> np.ndarray:
    """A side × side crop of the page image whose top-left is page point (ox, oy),
    padded with white outside the page."""
    k = OUTLINE_DPI / 72.0
    x0, y0 = round(ox * k), round(oy * k)
    out = np.full((side, side, 3), 255, np.uint8)
    h, w = img.shape[:2]
    sx0, sy0, sx1, sy1 = max(0, x0), max(0, y0), min(w, x0 + side), min(h, y0 + side)
    if sx1 > sx0 and sy1 > sy0:
        out[sy0 - y0:sy1 - y0, sx0 - x0:sx1 - x0] = img[sy0:sy1, sx0:sx1]
    return out


# --------------------------------------------------------------------------- stage


HANDLERS = {"scanned": ocr_scanned, "broken_text": ocr_broken, "outlined_text": ocr_outlined}


def run(doc: Document, pdf: Path, work_dir: Path) -> Document:
    """The OCR stage: handles every flagged page, records what ran in
    ``metadata["pdf2pptx:ocr"]`` (per page, plus model providers and timings)."""
    targets = [p for p in doc.pages if p.triage.kind in KINDS and not p.triage.fallback_image]
    report: dict[str, Any] = {"pages": {}, "flagged": len(targets), "total_pages": len(
        doc.pages)}
    if not targets:
        doc.metadata[META_KEY] = json.dumps(report, sort_keys=True)
        return doc
    eng = engine_id()
    report["engine"] = eng
    before = runtime.stats()
    if eng == "none":
        msg = "OCR unavailable (run `pdf2pptx setup models --lite`); page kept as is"
        LOG.warning("%s: %d flagged page(s): %s", pdf.name, len(targets), msg)
        for p in targets:
            p.triage.reason += f"; {msg}"
        doc.metadata[META_KEY] = json.dumps(report, sort_keys=True)
        return doc
    for page in targets:
        try:
            info = HANDLERS[page.triage.kind](doc, page, pdf, work_dir)
        except Exception as exc:
            LOG.error("page %d: OCR (%s) failed (%s: %s); page kept as is", page.index,
                      page.triage.kind, type(exc).__name__, exc,
                      exc_info=LOG.isEnabledFor(10))
            info = {"kind": page.triage.kind, "error": f"{type(exc).__name__}: {exc}"}
        report["pages"][str(page.index)] = info
        summary = (f"ocr: {info.get('engine', '-')}, {info.get('words', info.get('replaced', 0))}"
                   f" {'words' if 'words' in info else 'replaced'}")
        page.triage.reason += f"; {summary}" if "error" not in info else f"; ocr failed: {info[
            'error']}"
    report["models"] = runtime.stats_since(before)
    doc.metadata[META_KEY] = json.dumps(report, sort_keys=True)
    return doc


def page_regions(doc: Document, index: int) -> list[dict[str, Any]]:
    """Layout regions stored for a page by the OCR stage (page points)."""
    try:
        rep = json.loads(doc.metadata.get(META_KEY, "{}"))
        return list(rep.get("pages", {}).get(str(index), {}).get("regions", []))
    except ValueError:
        return []
