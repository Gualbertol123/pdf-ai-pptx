"""verify/diff.py: alignment, region scores, mismatch regions -> IR ids, heatmap."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest
from PIL import Image

from pdf2pptx.verify import diff, metrics
from pdf2pptx.verify.diff import ElementBox


def _page(w: int = 400, h: int = 300) -> np.ndarray:
    img = np.full((h, w, 3), 255, np.uint8)
    img[40:60, 40:200] = 0  # a "text line"
    img[150:250, 250:350] = (30, 90, 200)  # a "shape"
    return img


def test_identical_images_have_no_mismatch() -> None:
    a = _page()
    d = diff.compare(a, a.copy(), (400, 300))  # 1 px per pt
    assert d.ssim == pytest.approx(1.0)
    assert d.region_ssim_min == pytest.approx(1.0)
    assert d.mismatches == []
    assert len(d.region_scores) == 4 and all(len(r) == 4 for r in d.region_scores)


def test_mismatch_region_names_overlapping_elements() -> None:
    a = _page()
    b = a.copy()
    b[150:250, 250:350] = 255  # the shape is missing in the PPTX render
    boxes = [ElementBox("t1", "text_block", "rule", (40, 40, 200, 60)),
             ElementBox("s1", "shape", "pdf", (250, 150, 350, 250)),
             ElementBox("bg", "shape", "pdf", (0, 0, 400, 300))]
    d = diff.compare(a, b, (400, 300), boxes=boxes)
    assert d.mismatches, "the missing shape must be reported"
    top = d.mismatches[0]
    x0, y0, x1, y1 = top.bbox
    assert 230 <= x0 <= 260 and 330 <= x1 <= 370 and 130 <= y0 <= 160 and 240 <= y1 <= 270
    assert top.element_ids[0] == "s1" and "bg" in top.element_ids and "t1" not in top.element_ids
    assert top.ssim < 0.6
    # the damage is confined to the bottom-right quadrant of the 4x4 grid
    assert d.region_scores[0][0] == pytest.approx(1.0)
    assert min(d.region_scores[2][2:] + d.region_scores[3][2:]) < 0.9


def test_alignment_crops_letterboxed_page() -> None:
    """A 400x300 page placed centred on a 400x400 slide (scale 1, dy = 50)."""
    page = _page()
    slide = np.full((400, 400, 3), 255, np.uint8)
    slide[50:350] = page
    assert diff.page_rect_on_slide(400, 300, 400, 400) == (1.0, 0.0, 50.0)
    d = diff.compare(page, slide, (400, 300), slide_size=(400, 400))
    assert d.ssim == pytest.approx(1.0) and d.mismatches == []
    # and at a different render resolution (2 px per pt on the slide render)
    big = np.asarray(Image.fromarray(slide).resize((800, 800), Image.Resampling.NEAREST))
    assert diff.compare(page, big, (400, 300), slide_size=(400, 400)).ssim > 0.99


def test_map_lines_inverts_page_placement() -> None:
    w = metrics.Word("x", 10, 70, 20, 80)
    (m,), = metrics.map_lines([[w]], 0.5, 0, 50)
    assert (m.x0, m.y0, m.x1, m.y1) == (20, 40, 40, 60)


def test_heatmap_png(tmp_path: Path) -> None:
    a = _page()
    b = a.copy()
    b[40:60, 40:200] = 255
    d = diff.compare(a, b, (400, 300))
    p = diff.write_heatmap(d, tmp_path / "h.png", a)
    img = np.asarray(Image.open(p))
    assert img.shape == (300, 400, 3)
    # the missing line is coloured, untouched white stays light
    assert img[50, 120].astype(int).sum() < img[280, 20].astype(int).sum()
