"""Render check of native charts (M9): each chart's frame on the PowerPoint render is
compared with the same region of the PDF page; a chart below ``MIN_SSIM`` goes back
to the shapes it replaced (``analyze.charts.restore``) and the deck is rewritten.

Without PowerPoint the charts stay (recorded as ``unchecked``).
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import cv2
import numpy as np
from skimage.metrics import structural_similarity

from .. import log
from ..ir import Chart, Document

LOG = log.get(__name__)

MIN_SSIM = 0.85  # chart frame SSIM (grayscale, render dpi) below which the shapes return
NOTE = "chart_check:"


def charts_of(doc: Document) -> list[tuple[int, Chart]]:
    return [(p.index, e) for p in doc.pages for e in p.elements if isinstance(e, Chart)]


def region_ssim(a: np.ndarray, b: np.ndarray, box: tuple[float, float, float, float],
                scale: float) -> float:
    h, w = a.shape[:2]
    x0, y0 = max(0, int(box[0] * scale)), max(0, int(box[1] * scale))
    x1, y1 = min(w, round(box[2] * scale)), min(h, round(box[3] * scale))
    if x1 - x0 < 8 or y1 - y0 < 8:
        return 1.0
    ca, cb = a[y0:y1, x0:x1], b[y0:y1, x0:x1]
    return float(structural_similarity(ca, cb, data_range=255))


def check(doc: Document, pptx: Path, work_dir: Path, pdf: Path, *, dpi: int = 150,
          builds: bool = False, engine: Any = None) -> dict[str, Any]:
    """Render ``pptx``, score every chart's frame; revert the failing ones in ``doc``.
    Returns {"charts": n, "kept": [...], "reverted": [...], "ssim": {id: v}} (the caller
    rewrites the deck when something was reverted)."""
    from ..analyze.builds import page_to_slide
    from ..analyze.charts import restore
    from ..ingest.render import ensure_render

    items = charts_of(doc)
    res: dict[str, Any] = {"charts": len(items), "kept": [], "reverted": [], "ssim": {}}
    if not items:
        return res
    t0 = time.perf_counter()
    try:
        from . import ppt_com

        eng = engine or ppt_com.shared()
        slide_of = page_to_slide(doc, builds)
        pages = sorted({pi for pi, _ in items})
        out_dir = work_dir / "chart_check"
        with eng.opened(pptx.resolve()) as info:
            pngs = eng.render(info, [slide_of[pi] for pi in pages], dpi, out_dir)
            native = eng.charts(info)
    except Exception as exc:  # PowerPoint unavailable: keep the charts
        res["unchecked"] = f"{type(exc).__name__}: {exc}"
        for _, ch in items:
            res["kept"].append(ch.id)
        return res
    for pi, png in zip(pages, pngs, strict=True):
        page = doc.pages[pi]
        ref = cv2.imread(str(ensure_render(pdf, work_dir, pi, dpi)), cv2.IMREAD_GRAYSCALE)
        got = cv2.imread(str(png), cv2.IMREAD_GRAYSCALE)
        if ref is None or got is None:
            continue
        got = cv2.resize(got, (ref.shape[1], ref.shape[0]), interpolation=cv2.INTER_AREA)
        scale = ref.shape[1] / page.width
        names = {n for n, has, _ in native[slide_of[pi]] if has}
        for ch in [e for e in page.elements if isinstance(e, Chart)]:
            v = region_ssim(ref, got, ch.bbox, scale)
            res["ssim"][ch.id] = round(v, 4)
            ch.notes = [n for n in ch.notes if not n.startswith(NOTE)]
            is_native = f"Chart {ch.id}" in names
            if v < MIN_SSIM or not is_native:
                why = f"render ssim {v:.3f} < {MIN_SSIM}" if is_native else "not native"
                restore(page, ch.id, why)
                res["reverted"].append({"id": ch.id, "page": pi, "ssim": round(v, 4),
                                        "reason": why})
            else:
                ch.notes.append(f"{NOTE}ssim={v:.4f}")
                res["kept"].append(ch.id)
    res["seconds"] = round(time.perf_counter() - t0, 2)
    LOG.info("%s: chart check: %d kept, %d reverted (%.1fs)", pdf.name, len(res["kept"]),
             len(res["reverted"]), res["seconds"])
    return res
