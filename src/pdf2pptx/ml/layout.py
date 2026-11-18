"""Page layout regions from PP-DocLayoutV3 (25 classes) and how they are used.

* :func:`detect` on any page image: regions with label, score and the model's
  reading order (image pixels).
* Scanned pages (``ml/ocr.py``): text inside formula / picture / chart regions stays
  in the picture; the regions are stored with the page for roles and the report.
* :func:`refine_roles` on born-digital pages runs the model only when a text block's
  rule confidence is below ``TRIGGER`` (PLAN.md §8 rule 2). A disagreement between
  the rule role and the region label is logged (``metadata["pdf2pptx:layout"]`` and
  the debug report); the model's role is adopted only when the rule was unsure, the
  model is sure (score ≥ ``ADOPT``) and both roles are among title / body / caption.
"""

from __future__ import annotations

import json
import time
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

import cv2
import numpy as np

from .. import log
from ..ir import Document, Page, TextBlock
from . import models, runtime

if TYPE_CHECKING:
    from .ocr import Frame

LOG = log.get(__name__)

INPUT = 800
THRESHOLD = 0.5
TRIGGER = 0.65  # rule role confidence below this asks the model (born-digital pages)
ADOPT = 0.8  # model score needed to change a role
SCANNED_TRIGGER = 0.75  # on OCR pages the rules see noisier geometry: ask more often
META_KEY = "pdf2pptx:layout"
MODEL_SOURCE = "model:pp-doclayoutv3"
DEFAULT_LABELS = (
    "abstract", "algorithm", "aside_text", "chart", "content", "display_formula",
    "doc_title", "figure_title", "footer", "footer_image", "footnote", "formula_number",
    "header", "header_image", "image", "inline_formula", "number", "paragraph_title",
    "reference", "reference_content", "seal", "table", "text", "vertical_text",
    "vision_footnote")
# region label -> block role it suggests (None: no opinion)
LABEL_ROLE: dict[str, str | None] = {
    "doc_title": "title", "paragraph_title": None, "text": "body", "content": "body",
    "abstract": "body", "reference": "body", "reference_content": "body",
    "aside_text": "body", "algorithm": "body", "footer": "footer", "header": "header",
    "number": "page_number", "figure_title": "caption", "vision_footnote": "caption",
    "footnote": "footer",
}
ADOPTABLE = {"title", "body", "caption"}


@dataclass
class Region:
    label: str
    score: float
    box: tuple[float, float, float, float]  # image px (x0, y0, x1, y1)
    order: int = 0

    def to_page(self, fr: Frame) -> dict[str, Any]:
        corners = [fr.to_pt(x, y) for x in (self.box[0], self.box[2])
                   for y in (self.box[1], self.box[3])]
        xs, ys = [c[0] for c in corners], [c[1] for c in corners]
        return {"label": self.label, "score": round(self.score, 3), "order": self.order,
                "bbox": [round(min(xs), 1), round(min(ys), 1), round(max(xs), 1),
                         round(max(ys), 1)]}


def available() -> bool:
    return models.available("layout") and runtime.ort_module() is not None


def labels() -> tuple[str, ...]:
    p = models.path("layout_config")
    if p is not None:
        try:
            got = json.loads(p.read_text(encoding="utf-8")).get("label_list")
            if got:
                return tuple(str(x) for x in got)
        except ValueError:
            pass
    return DEFAULT_LABELS


def _warmup(s: runtime.Session) -> None:
    s.run(_feeds(np.full((INPUT, INPUT, 3), 255, np.uint8)))


def _feeds(img: np.ndarray) -> dict[str, np.ndarray]:
    h, w = img.shape[:2]
    x = cv2.resize(img, (INPUT, INPUT), interpolation=cv2.INTER_LINEAR)
    x = cv2.cvtColor(x, cv2.COLOR_BGR2RGB).astype(np.float32) / 255.0  # config: /255 only
    return {"im_shape": np.array([[INPUT, INPUT]], np.float32),
            "image": np.ascontiguousarray(x.transpose(2, 0, 1)[None]),
            "scale_factor": np.array([[INPUT / h, INPUT / w]], np.float32)}


def detect(img: np.ndarray, threshold: float = THRESHOLD) -> list[Region]:
    """Regions of a BGR page image (pixels of that image)."""
    s = runtime.session("layout", _warmup)
    if s is None:
        return []
    out = s.run(_feeds(img))[0]
    names = labels()
    regs = []
    h, w = img.shape[:2]
    for row in out:
        k, score = int(row[0]), float(row[1])
        if score < threshold or not 0 <= k < len(names):
            continue
        x0, y0, x1, y1 = (float(v) for v in row[2:6])
        x0, y0, x1, y1 = max(0.0, x0), max(0.0, y0), min(float(w), x1), min(float(h), y1)
        if x1 - x0 < 2 or y1 - y0 < 2:
            continue
        regs.append(Region(names[k], round(score, 4), (round(x0, 1), round(y0, 1),
                                                         round(x1, 1), round(y1, 1)),
                           int(row[6]) if len(row) > 6 else 0))
    return _dedupe(regs)


def _iou(a: tuple[float, ...], b: tuple[float, ...]) -> float:
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    ua = (a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - inter
    return inter / ua if ua > 0 else 0.0


def _dedupe(regs: list[Region]) -> list[Region]:
    regs = sorted(regs, key=lambda r: (-r.score, r.box))
    keep: list[Region] = []
    for r in regs:
        if all(_iou(r.box, k.box) < 0.8 for k in keep):
            keep.append(r)
    return sorted(keep, key=lambda r: (r.order, r.box))


# --------------------------------------------------------------------------- roles


def region_for(bbox: tuple[float, float, float, float],
               regions: list[dict[str, Any]]) -> dict[str, Any] | None:
    """The region (page points) holding most of ``bbox`` (≥ 50% of its area)."""
    area = max((bbox[2] - bbox[0]) * (bbox[3] - bbox[1]), 1e-6)
    best = None
    for r in regions:
        b = r["bbox"]
        ov = max(0.0, min(bbox[2], b[2]) - max(bbox[0], b[0])) * max(
            0.0, min(bbox[3], b[3]) - max(bbox[1], b[1]))
        if ov >= 0.5 * area and (best is None or ov > best[0]):
            best = (ov, r)
    return best[1] if best else None


def compare_roles(page: Page, regions: list[dict[str, Any]], trigger: float,
                  adopt: bool = True) -> list[dict[str, Any]]:
    """Check the page's low-confidence blocks against the regions; returns the
    disagreements (adopted or only logged)."""
    out = []
    for tb in page.elements:
        if not isinstance(tb, TextBlock) or tb.role_confidence >= trigger:
            continue
        r = region_for(tb.bbox, regions)
        if r is None:
            continue
        want = LABEL_ROLE.get(r["label"])
        if want is None or want == tb.role:
            continue
        rec = {"block": tb.id, "rule": tb.role, "rule_conf": round(tb.role_confidence, 3),
               "model": r["label"], "model_role": want, "model_score": r["score"],
               "adopted": False}
        if (adopt and r["score"] >= ADOPT and want in ADOPTABLE and tb.role in ADOPTABLE
                and tb.placeholder is None):
            tb.notes.append(f"role:{MODEL_SOURCE} {tb.role}->{want} ({r['score']:.2f}; "
                            f"rule conf {tb.role_confidence:.2f})")
            tb.role = want
            tb.role_confidence = round(float(r["score"]), 3)
            rec["adopted"] = True
        out.append(rec)
    return out


def refine_roles(doc: Document, pdf: Path, work_dir: Path, *, dpi: int = 150) -> dict[str, Any]:
    """Born-digital pages with a low-confidence block: run the layout model on the page
    render and compare (OCR pages use the regions stored by the OCR stage)."""
    from ..ingest.render import ensure_render
    from .ocr import page_regions, read_bgr

    rep: dict[str, Any] = {"pages_checked": 0, "model_runs": 0, "pages_total": len(doc.pages),
                           "blocks_low": 0,
                           "disagreements": [], "available": available()}
    t0 = time.perf_counter()
    before = runtime.stats()
    for page in doc.pages:
        if page.triage.fallback_image:
            continue
        ocr_page = page.triage.kind == "scanned" and any(g.source == "ocr" for g in page.glyphs)
        trigger = SCANNED_TRIGGER if ocr_page else TRIGGER
        low = [e for e in page.elements if isinstance(e, TextBlock)
               and e.role_confidence < trigger]
        rep["blocks_low"] += len(low)
        if not low:
            continue
        regions = page_regions(doc, page.index) if ocr_page else []
        if not regions and not ocr_page:
            if not rep["available"]:
                continue
            try:
                img = read_bgr(ensure_render(pdf, work_dir, page.index, dpi))
                k = 72.0 / dpi
                regions = [{"label": r.label, "score": round(r.score, 3), "order": r.order,
                            "bbox": [round(v * k, 1) for v in r.box]} for r in detect(img)]
                rep["model_runs"] += 1
            except Exception as exc:
                LOG.warning("page %d: layout model failed (%s)", page.index, exc)
                continue
        rep["pages_checked"] += 1
        for d in compare_roles(page, regions, trigger):
            d["page"] = page.index
            rep["disagreements"].append(d)
    rep["seconds"] = round(time.perf_counter() - t0, 3)
    rep["models"] = runtime.stats_since(before)
    born = sum(1 for p in doc.pages if p.triage.kind != "scanned")
    rep["trigger_rate"] = round(rep["model_runs"] / max(born, 1), 4)
    doc.metadata[META_KEY] = json.dumps({k: v for k, v in rep.items() if k != "seconds"},
                                        sort_keys=True)
    return rep
