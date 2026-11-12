"""Aligned per-slide comparison of a PDF page render with its PowerPoint render.

* **Alignment.** The builder places a page on the slide with a uniform scale and
  centring (``build/writer._xf_for``). The PowerPoint render is cropped to that page
  rectangle and resampled to the PDF render's pixel size, so both images cover the
  same page area pixel for pixel. When page and slide sizes agree (the usual case)
  this is the identity and scores equal a plain whole-image SSIM.
* **Scores.** Grayscale Gaussian SSIM map (``metrics.ssim_map``), its mean, and the
  4×4 grid of region means (``region_ssim_min`` = the worst cell).
* **Mismatch regions.** Low local SSIM, smoothed and closed so a mis-placed word
  becomes one blob, split into connected components. Each region carries its page
  bbox (pt), its mean SSIM and the ids of IR elements whose bboxes overlap it.
* **Heatmap.** A PNG: the PDF render faded to gray, dissimilarity in colour on top,
  mismatch regions outlined.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import cv2
import numpy as np

from . import metrics

GRID = 4


@dataclass(frozen=True)
class ElementBox:
    """What diff needs from an IR element: id, kind, source and page bbox (pt)."""

    id: str
    kind: str
    source: str
    bbox: tuple[float, float, float, float]


@dataclass
class Mismatch:
    bbox: tuple[float, float, float, float]  # page pt, top-left origin
    ssim: float  # mean local SSIM inside the bbox
    severity: float  # sum of (1 - ssim) over the region's pixels, in pt²
    element_ids: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"bbox": [round(v, 2) for v in self.bbox], "ssim": round(self.ssim, 4),
                "severity": round(self.severity, 2), "element_ids": self.element_ids}


@dataclass
class SlideDiff:
    ssim: float
    region_scores: list[list[float]]
    mismatches: list[Mismatch]
    px_per_pt: float
    smap: np.ndarray = field(repr=False)

    @property
    def region_ssim_min(self) -> float:
        return min(min(r) for r in self.region_scores)

    def to_dict(self) -> dict[str, Any]:
        return {"ssim": round(self.ssim, 4), "region_ssim_min": round(self.region_ssim_min, 4),
                "regions": [[round(v, 4) for v in row] for row in self.region_scores],
                "mismatches": [m.to_dict() for m in self.mismatches]}


# --------------------------------------------------------------------------- alignment


def page_rect_on_slide(page_w: float, page_h: float, slide_w: float,
                       slide_h: float) -> tuple[float, float, float]:
    """(scale, dx, dy): page pt -> slide pt, as the builder places pages."""
    if page_w <= 0 or page_h <= 0:
        return 1.0, 0.0, 0.0
    s = min(slide_w / page_w, slide_h / page_h)
    return s, (slide_w - page_w * s) / 2, (slide_h - page_h * s) / 2


def align(pptx_img: np.ndarray, page_size: tuple[float, float],
          slide_size: tuple[float, float] | None, out_wh: tuple[int, int]) -> np.ndarray:
    """Crop the page rectangle out of a slide render and resample it to ``out_wh``."""
    img = np.asarray(pptx_img)
    if slide_size is not None:
        (pw, ph), (sw, sh) = page_size, slide_size
        if abs(pw - sw) > 0.05 or abs(ph - sh) > 0.05:
            s, dx, dy = page_rect_on_slide(pw, ph, sw, sh)
            kx, ky = img.shape[1] / sw, img.shape[0] / sh
            x0, y0 = round(dx * kx), round(dy * ky)
            x1, y1 = round((dx + pw * s) * kx), round((dy + ph * s) * ky)
            img = img[max(y0, 0):max(y1, y0 + 1), max(x0, 0):max(x1, x0 + 1)]
    if (img.shape[1], img.shape[0]) != out_wh:
        img = cv2.resize(img, out_wh, interpolation=cv2.INTER_LANCZOS4)
    return img


# --------------------------------------------------------------------------- comparison


def _overlap(a: Sequence[float], b: Sequence[float]) -> float:
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0.0


def overlapping_ids(bbox: Sequence[float], boxes: Sequence[ElementBox],
                    limit: int = 20) -> list[str]:
    """Ids of elements overlapping ``bbox``, best match first (ties by id).

    Ranked by overlap² / (region area × element area): 1 when region and element
    coincide, small for a page-sized background that merely contains the region.
    """
    pad = 0.5
    area = max((bbox[2] - bbox[0]) * (bbox[3] - bbox[1]), 1e-6)
    hits = []
    for e in boxes:
        x0, y0, x1, y1 = e.bbox
        eb = (min(x0, x1) - pad, min(y0, y1) - pad, max(x0, x1) + pad, max(y0, y1) + pad)
        ov = _overlap(bbox, eb)
        if ov > 0:
            hits.append((-ov * ov / (area * (eb[2] - eb[0]) * (eb[3] - eb[1])), e.id))
    return [i for _, i in sorted(hits)[:limit]]


def mismatch_regions(smap: np.ndarray, px_per_pt: float, boxes: Sequence[ElementBox] = (),
                     *, threshold: float = 0.6, min_area_pt2: float = 6.0,
                     limit: int = 40) -> list[Mismatch]:
    dis = np.clip(1.0 - smap, 0.0, 1.0).astype(np.float32)
    sigma = max(1.0, 1.5 * px_per_pt)  # ~1.5 pt
    blur = cv2.GaussianBlur(dis, (0, 0), sigma)
    mask = (blur > 1.0 - threshold).astype(np.uint8)
    k = max(3, round(3 * px_per_pt) | 1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((k, k), np.uint8))
    n, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    px_area = 1.0 / (px_per_pt * px_per_pt)
    out = []
    for i in range(1, n):
        x, y, w, h, area = (int(v) for v in stats[i])
        if area * px_area < min_area_pt2:
            continue
        sel = labels[y:y + h, x:x + w] == i
        region = smap[y:y + h, x:x + w][sel]
        bbox = (x / px_per_pt, y / px_per_pt, (x + w) / px_per_pt, (y + h) / px_per_pt)
        out.append(Mismatch(bbox=bbox, ssim=float(region.mean()),
                            severity=float((1.0 - region).clip(0).sum() * px_area)))
    out.sort(key=lambda m: (-m.severity, m.bbox))
    out = out[:limit]
    for m in out:
        m.element_ids = overlapping_ids(m.bbox, boxes)
    return out


def compare(pdf_img: np.ndarray, pptx_img: np.ndarray, page_size: tuple[float, float], *,
            slide_size: tuple[float, float] | None = None,
            boxes: Sequence[ElementBox] = ()) -> SlideDiff:
    """Compare a PDF page render with the PowerPoint render of its slide."""
    h, w = pdf_img.shape[:2]
    aligned = align(pptx_img, page_size, slide_size, (w, h))
    score, smap = metrics.ssim_map(pdf_img, aligned)
    px_per_pt = w / page_size[0] if page_size[0] > 0 else 1.0
    return SlideDiff(ssim=score, region_scores=metrics.region_scores(smap, GRID),
                     mismatches=mismatch_regions(smap, px_per_pt, boxes),
                     px_per_pt=px_per_pt, smap=smap)


# --------------------------------------------------------------------------- heatmap


def heatmap(diff: SlideDiff, pdf_img: np.ndarray | None = None) -> np.ndarray:
    """RGB heatmap: faded page, dissimilarity colour-mapped on top, regions outlined."""
    dis = np.clip(1.0 - diff.smap, 0.0, 1.0)
    h, w = dis.shape
    if pdf_img is not None:
        gray = metrics.to_gray(pdf_img).astype(np.float32)
        base = np.repeat((180 + gray[..., None] * (75 / 255.0)), 3, axis=2)
    else:
        base = np.full((h, w, 3), 255.0, np.float32)
    color = cv2.applyColorMap((dis * 255).astype(np.uint8), cv2.COLORMAP_INFERNO)[..., ::-1]
    alpha = np.clip(dis * 2.5, 0.0, 0.85)[..., None]
    out = (base * (1 - alpha) + color.astype(np.float32) * alpha).astype(np.uint8)
    out = np.ascontiguousarray(out)
    k = diff.px_per_pt
    for m in diff.mismatches:
        x0, y0, x1, y1 = (round(v * k) for v in m.bbox)
        cv2.rectangle(out, (x0, y0), (x1 - 1, y1 - 1), (0, 160, 255), max(1, int(k)))
    ys = np.linspace(0, h, GRID + 1).astype(int)
    xs = np.linspace(0, w, GRID + 1).astype(int)
    for v in xs[1:-1]:
        cv2.line(out, (int(v), 0), (int(v), h - 1), (150, 150, 150), 1)
    for v in ys[1:-1]:
        cv2.line(out, (0, int(v)), (w - 1, int(v)), (150, 150, 150), 1)
    return out


def write_heatmap(diff: SlideDiff, path: Path, pdf_img: np.ndarray | None = None) -> Path:
    from PIL import Image

    path.parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(heatmap(diff, pdf_img)).save(path, optimize=True)
    return path


# --------------------------------------------------------------------------- IR -> boxes


def page_boxes(page: Any, fonts: dict[str, Any] | None = None) -> list[ElementBox]:
    """Element boxes of an IR page, plus the per-line text boxes the builder makes
    from glyphs when the page has no TextBlocks yet (M2 baseline)."""
    from ..ir import TextBlock

    out = [ElementBox(e.id, e.kind, e.source, tuple(e.bbox)) for e in page.elements]
    if fonts is not None and not any(isinstance(e, TextBlock) for e in page.elements):
        from ..build.lines import line_blocks

        out += [ElementBox(b.id, b.kind, b.source, tuple(b.bbox))
                for b in line_blocks(page, fonts)]
    return out
