"""Scoring end to end: PDF-side extraction, plus metric sanity through PowerPoint."""

from __future__ import annotations

from pathlib import Path

import pytest
from reportlab.pdfgen import canvas

from pdf2pptx.config import Config
from pdf2pptx.verify import ppt_com, score
from pdf2pptx.verify.pdf_text import page_refs, render_pages
from pdf2pptx.verify.pptx_stats import blank_deck, deck_stats

from .conftest import CORPUS


def _two_column_pdf(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=(400, 300))
    c.setFont("Helvetica", 12)
    for i in range(3):
        y = 300 - 50 - i * 16
        c.drawString(20, y, f"left column line {i}")
        c.drawString(220, y, f"right column line {i}")
    c.save()


def test_pdf_lines_split_columns(tmp_path: Path) -> None:
    pdf = tmp_path / "cols.pdf"
    _two_column_pdf(pdf)
    ref = page_refs(pdf)[0]
    assert (ref.width, ref.height) == (400, 300)
    assert len(ref.lines) == 6
    texts = sorted(" ".join(w.text for w in line) for line in ref.lines)
    assert texts[0] == "left column line 0"
    first = next(line for line in ref.lines if line[0].text == "left")
    assert first[0].x0 == pytest.approx(20, abs=0.5)
    # reportlab y=250 from the bottom -> baseline at 50 from the top; bottom = baseline + descent
    assert first[0].y1 == pytest.approx(50 + 12 * 0.22, abs=1.5)
    img = render_pages(pdf, 150)[0]
    assert img.shape == (625, 833, 3)


def test_blank_deck_stats(tmp_path: Path) -> None:
    p = blank_deck(tmp_path / "b.pptx", [(960.0, 540.0)] * 2)
    st = deck_stats(p)
    assert (st.width_pt, st.height_pt) == (960.0, 540.0)
    assert len(st.slides) == 2 and st.slides[0].text == ""


@pytest.mark.corpus
def test_deck_stats_counts() -> None:
    st = deck_stats(CORPUS / "synthetic" / "d09_shapes_basic.pptx").slides[0]
    assert st.n_preset == 8 and st.n_freeform == 0
    st = deck_stats(CORPUS / "synthetic" / "d07_table_simple.pptx").slides[0]
    assert st.n_tables == 1 and st.n_text_boxes == 2


@pytest.mark.corpus
@pytest.mark.powerpoint
def test_truth_scores_high_blank_scores_low(tmp_path: Path) -> None:
    cfg = Config(work_dir=tmp_path)
    syn = CORPUS / "synthetic"
    with ppt_com.PowerPointEngine() as eng:
        for name in ("d01_title_bullets", "d07_table_simple", "d23_long_paragraphs"):
            _check_truth_vs_blank(cfg, syn, name, eng, tmp_path)
    # an instance this engine used is gone unless another live engine still uses it
    shared = ppt_com._SHARED.owned_pids if ppt_com._SHARED else []
    assert set(eng.owned_history) & set(ppt_com.owned_alive()) <= set(shared)


def _check_truth_vs_blank(cfg: Config, syn: Path, name: str, eng: ppt_com.PowerPointEngine,
                          tmp_path: Path) -> None:
    res = score.score_pair(syn / f"{name}.pdf", syn / f"{name}.pptx", cfg, name=name,
                           engine=eng)
    assert len(res["com_seconds_per_slide"]) == res["slides_in_pptx"]
    assert res["powerpoint"] == "ok"
    s = res["summary"]
    assert s["ssim"]["median"] >= 0.97, name
    assert s["text_sim"]["median"] >= 0.99, name
    assert s["pos_err_pt"]["median"] <= 1.0, name
    assert s["linebreak_match"]["median"] >= 0.99, name
    refs = page_refs(syn / f"{name}.pdf")
    blank = blank_deck(tmp_path / f"{name}_blank.pptx", [(r.width, r.height) for r in refs])
    b = score.score_pair(syn / f"{name}.pdf", blank, cfg, name=f"{name}_blank", engine=eng)
    assert b["summary"]["text_sim"]["median"] == 0.0
    assert b["summary"]["composite"]["median"] < 0.6
    assert b["summary"]["region_ssim_min"]["median"] < s["region_ssim_min"]["median"]


@pytest.mark.corpus
def test_score_without_powerpoint_is_null_with_reason(tmp_path: Path) -> None:
    cfg = Config(work_dir=tmp_path)
    syn = CORPUS / "synthetic"
    res = score.score_pair(syn / "d01_title_bullets.pdf", syn / "d01_title_bullets.pptx", cfg,
                           engine_error="PowerPointUnavailable: simulated")
    assert res["powerpoint"] is None and "simulated" in res["powerpoint_error"]
    s = res["summary"]
    for k in ("ssim", "region_ssim_min", "pos_err_pt", "linebreak_match"):
        assert s[k]["median"] is None, k
    assert s["text_sim"]["median"] == 1.0 and s["composite"]["median"] is not None
