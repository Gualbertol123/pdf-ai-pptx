"""Pipeline + baseline writer: end-to-end, caching, determinism, page fallback."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from reportlab.pdfgen import canvas

from pdf2pptx import pipeline
from pdf2pptx.config import Config
from pdf2pptx.ingest import parse
from pdf2pptx.ir import Document

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture()
def two_pages(tmp_path: Path) -> Path:
    pdf = tmp_path / "two.pdf"
    c = canvas.Canvas(str(pdf), pagesize=(720, 405))
    for i in range(2):
        c.setFont("Helvetica", 20)
        c.drawString(50, 350, f"Slide {i + 1} title")
        c.drawString(50, 300, "Left")
        c.drawString(400, 300, "Right column")
        c.setFillColorRGB(0, 0.5, 1)
        c.circle(200, 150, 40, stroke=1, fill=1)
        c.showPage()
    c.save()
    return pdf


def _cfg(tmp_path: Path) -> Config:
    return Config(work_dir=tmp_path / "work", out_dir=tmp_path / "out", workers=1)


def test_convert_end_to_end(two_pages: Path, tmp_path: Path) -> None:
    cfg = _cfg(tmp_path)
    out = tmp_path / "two.pptx"
    res = pipeline.convert(two_pages, out, cfg)
    assert res.pages == 2 and not res.fallback_pages and not res.errors
    prs = Presentation(str(out))
    assert len(prs.slides) == 2
    assert abs(prs.slide_width - 720 * 12700) <= 1 and abs(prs.slide_height - 405 * 12700) <= 1
    s = prs.slides[0]
    texts = sorted(sh.text_frame.text for sh in s.shapes
                   if sh.has_text_frame and sh.text_frame.text)
    # the far-apart "Left" / "Right column" stay separate blocks
    assert texts == ["Left", "Right column", "Slide 1 title"]
    box = next(sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text == "Left")
    tf = box.text_frame
    assert tf.word_wrap is True and tf.margin_left == 0 and tf.margin_top == 0
    assert any(sh.shape_type == MSO_SHAPE_TYPE.FREEFORM for sh in s.shapes)
    work = cfg.work_dir / pipeline.sha256_file(two_pages)
    for stage in ("ingest", "triage", "analyze", "build"):
        assert (work / f"ir_{stage}.json").is_file()
    assert res.crosscheck is not None and res.crosscheck["text_ratio"] == 1.0


def test_lines_layout_keeps_the_baseline_path(two_pages: Path, tmp_path: Path) -> None:
    cfg = Config(work_dir=tmp_path / "work", out_dir=tmp_path / "out", workers=1,
                 text_layout="lines")
    out = tmp_path / "lines.pptx"
    pipeline.convert(two_pages, out, cfg)
    s = Presentation(str(out)).slides[0]
    boxes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text]
    assert [b.text_frame.text for b in boxes] == ["Slide 1 title", "Left", "Right column"]
    assert all(b.text_frame.word_wrap is False for b in boxes)  # one non-wrapping box per line


def test_cache_and_determinism(two_pages: Path, tmp_path: Path) -> None:
    cfg = _cfg(tmp_path)
    a, b = tmp_path / "a.pptx", tmp_path / "b.pptx"
    pipeline.convert(two_pages, a, cfg, use_cache=False)
    res = pipeline.convert(two_pages, b, cfg)
    assert res.cached_stages == ["ingest", "triage", "analyze"]
    assert hashlib.sha256(a.read_bytes()).digest() == hashlib.sha256(b.read_bytes()).digest()


def test_bad_page_falls_back_to_picture(two_pages: Path, tmp_path: Path,
                                        monkeypatch: pytest.MonkeyPatch) -> None:
    real = parse.parse_page

    def flaky(doc, index, work_dir, fonts):  # type: ignore[no-untyped-def]
        if index == 1:
            raise RuntimeError("boom")
        return real(doc, index, work_dir, fonts)

    monkeypatch.setattr(parse, "parse_page", flaky)
    cfg = _cfg(tmp_path)
    out = tmp_path / "fb.pptx"
    res = pipeline.convert(two_pages, out, cfg, use_cache=False)
    assert res.fallback_pages == [1]
    prs = Presentation(str(out))
    assert len(prs.slides) == 2
    shapes = list(prs.slides[1].shapes)
    assert len(shapes) == 1 and shapes[0].shape_type == MSO_SHAPE_TYPE.PICTURE
    assert shapes[0].width == prs.slide_width
    work = cfg.work_dir / pipeline.sha256_file(two_pages)
    final = Document.from_json((work / "ir_build.json").read_text(encoding="utf-8"))
    assert final.pages[1].triage.fallback_image and "boom" in final.pages[1].triage.reason
    assert (work / "renders" / "p0001_300.png").is_file()


@pytest.mark.corpus
@pytest.mark.powerpoint
def test_corpus_outputs_open_in_powerpoint(tmp_path: Path) -> None:
    cfg = Config(work_dir=tmp_path / "work", out_dir=tmp_path / "out", workers=2)
    names = ["d01_title_bullets", "d12_gradients_alpha", "d28_rotated_text", "s02_scanned_report"]
    s = pipeline.convert_corpus(cfg, only=names, check=True)
    assert s["converted"] == len(names) == s["opened_in_powerpoint"], s
    assert not s["failures"]
    r = json.loads((cfg.out_dir / "corpus" / "d28_rotated_text.convert.json").read_text(
        encoding="utf-8"))
    assert r["triage"] == {"outlined_text": 1}
