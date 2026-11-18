"""M8: model files, runtime, OCR, layout, inpainting. Tests that need the ONNX models
skip when they are not installed (``pdf2pptx setup models --lite``)."""

from __future__ import annotations

import codecs
import hashlib
import io
import json
import re
from pathlib import Path
from typing import Any

import cv2
import numpy as np
import pytest

from pdf2pptx.ml import fontfit, inpaint, layout, models, ocr, ocr_engine, runtime

ROOT = Path(__file__).resolve().parents[1]
SYN = ROOT / "corpus" / "synthetic"

needs_ocr = pytest.mark.skipif(
    not (ocr_engine.ppocr_available() and runtime.ort_module() is not None),
    reason="OCR models not installed (pdf2pptx setup models --lite)")
needs_layout = pytest.mark.skipif(not layout.available(), reason="layout model not installed")
needs_lama = pytest.mark.skipif(models.path("inpaint") is None, reason="LaMa not installed")
needs_fonts = pytest.mark.skipif(fontfit.face("Arial") is None or fontfit.face("Calibri") is None,
                                 reason="Arial / Calibri not installed")


def _work(pdf: Path) -> Path:
    from pdf2pptx.pipeline import sha256_file

    wd = ROOT / "work" / sha256_file(pdf)
    wd.mkdir(parents=True, exist_ok=True)
    return wd


def _doc(pdf: Path) -> Any:
    from pdf2pptx.ingest import parse, triage

    wd = _work(pdf)
    return triage.triage_document(parse.parse_document(pdf, wd, workers=1)), wd


# --------------------------------------------------------------------------- models


def test_manifest_is_pinned() -> None:
    spec = models.spec()
    assert {"ocr_det", "ocr_cls", "ocr_rec_latin", "ocr_rec_multi", "layout", "inpaint"} <= set(
        spec)
    for m in spec.values():
        assert re.fullmatch(r"[0-9a-f]{64}", m.sha256), m.name
        assert re.fullmatch(r"[0-9a-f]{40}", m.revision), m.name  # a commit, not "main"
        assert m.license == "Apache-2.0"
        assert m.revision in m.url and m.size > 0
    assert sum(m.size for m in spec.values() if m.pack == "lite") < models.ASK_ABOVE_BYTES


class _Resp(io.BytesIO):
    def __init__(self, data: bytes, status: int) -> None:
        super().__init__(data)
        self.status = status

    def __enter__(self) -> _Resp:
        return self

    def __exit__(self, *a: object) -> None:
        self.close()


def _fake_model(data: bytes) -> models.ModelFile:
    return models.ModelFile(name="tiny", pack="test", role="", version="1", license="MIT",
                            upstream="", repo="x/y", revision="0" * 40, path="t.bin",
                            file="sub/t.bin", size=len(data),
                            sha256=hashlib.sha256(data).hexdigest())


def test_install_resumes_and_verifies(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    data = bytes(range(256)) * 40
    m = _fake_model(data)
    monkeypatch.setattr(models, "spec", lambda: {"tiny": m})
    part = tmp_path / "sub" / "t.bin.part"
    part.parent.mkdir(parents=True)
    part.write_bytes(data[:1000])  # an interrupted earlier download
    ranges: list[str | None] = []

    def urlopen(req: Any, timeout: float = 0) -> _Resp:
        rng = req.get_header("Range")
        ranges.append(rng)
        start = int(rng.split("=")[1].rstrip("-")) if rng else 0
        return _Resp(data[start:], 206 if rng else 200)

    monkeypatch.setattr(models.urllib.request, "urlopen", urlopen)
    res = models.install("test", tmp_path, progress=lambda _: None)
    assert res["installed"] == ["tiny"]
    assert ranges == ["bytes=1000-"]
    assert (tmp_path / "sub" / "t.bin").read_bytes() == data
    assert not part.exists()
    man = json.loads((tmp_path / "manifest.json").read_text(encoding="utf-8"))
    assert man["models"][0]["installed"] is True


def test_install_rejects_bad_checksum(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    m = _fake_model(b"good bytes")
    monkeypatch.setattr(models, "spec", lambda: {"tiny": m})
    monkeypatch.setattr(models.urllib.request, "urlopen",
                        lambda req, timeout=0: _Resp(b"evil bytes", 200))
    with pytest.raises(RuntimeError, match="sha256"):
        models.install("test", tmp_path, progress=lambda _: None)
    assert not (tmp_path / "sub" / "t.bin").exists()


def test_install_asks_above_one_gb(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    big = models.ModelFile(**{**_fake_model(b"x").__dict__, "size": 2 << 30})
    monkeypatch.setattr(models, "spec", lambda: {"tiny": big})
    asked: list[str] = []
    res = models.install("test", tmp_path, ask=lambda q: asked.append(q) or False,
                         progress=lambda _: None)
    assert asked and res["installed"] == [] and "1 GB" in res["reason"]


def test_missing_models_degrade(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PDF2PPTX_MODELS_DIR", str(tmp_path))
    runtime.reset()
    try:
        assert models.path("ocr_det") is None
        assert runtime.session("ocr_det") is None
        assert not ocr_engine.ppocr_available() and not layout.available()
        assert layout.detect(np.full((100, 100, 3), 255, np.uint8)) == []
    finally:
        runtime.reset()


def test_no_engine_keeps_page(monkeypatch: pytest.MonkeyPatch) -> None:
    """No OCR engine at all: the scanned page stays a picture, with the reason."""
    pdf = SYN / "s01_scanned_deck_it.pdf"
    if not pdf.is_file():
        pytest.skip("corpus not built")
    doc, wd = _doc(pdf)
    monkeypatch.setattr(ocr, "engine_id", lambda: "none")
    ocr.run(doc, pdf, wd)
    assert all(not p.glyphs for p in doc.pages)
    assert "setup models" in doc.pages[0].triage.reason
    assert any(e.kind == "image" for e in doc.pages[0].elements)


# --------------------------------------------------------------------------- rules


def _lines_image(angle: float) -> np.ndarray:
    img = np.full((900, 1400), 245, np.uint8)
    for k, y in enumerate(range(120, 800, 55)):
        x = 80 if k % 2 else 700  # two columns whose lines do not line up
        cv2.putText(img, "The quick brown fox", (x, y + (k % 2) * 20),
                    cv2.FONT_HERSHEY_SIMPLEX, 1.1, 20, 2)
    return ocr_engine.rotate_image(img, angle, 245)


@pytest.mark.parametrize("angle", [-1.4, -0.7, 0.0, 0.45, 1.0])
def test_skew_estimate(angle: float) -> None:
    assert ocr_engine.estimate_skew(_lines_image(angle)) == pytest.approx(angle, abs=0.08)


def test_size_snap_and_clusters() -> None:
    assert ocr._snap(31.7) == 32.0 and ocr._snap(10.12) == 10.0 and ocr._snap(10.6) == 10.5
    assert ocr._snap(13.3) == 13.3 or ocr._snap(13.3) == 13.0
    g = [ocr.LineGeo(line=ocr_engine.OcrLine((0, 0, 1, 1), "ab", 1.0), baseline=0,
                     x_height=1, words=[ocr.WordGeo("abc", 0, 1, 1.0)], rgb="000000", stroke=1,
                     ink_top=0, ink_bot=1, size_px=s) for s in (31.2, 31.9, 32.3, 18.1)]
    ocr.unify_sizes(g, 1.0)
    assert [x.size_px for x in g] == [32.0, 32.0, 32.0, 18.0]


@needs_fonts
def test_font_family_from_word_widths() -> None:
    words = ["the", "quick", "brown", "fox", "jumps", "over", "a", "lazy", "dog", "while",
             "tables", "and", "charts", "wait"]
    for fam in ("Arial", "Calibri"):
        f = fontfit.face(fam)
        assert f is not None
        obs = [fontfit.WordObs(w, f.ink_em(w) * 40.0) for w in words]  # 40 px per em
        lines = [(obs[i:i + 5], False) for i in range(0, len(obs), 5)]
        best, scores = fontfit.choose_family(lines)
        assert best == fam, scores
        assert fontfit.fit_size(f, obs)[0] == pytest.approx(40.0, rel=1e-3)


def test_bullet_found_by_shape() -> None:
    img = np.full((120, 600), 250, np.uint8)
    cv2.circle(img, (40, 62), 7, 10, -1)  # a bullet dot at mid x-height
    line = ocr_engine.OcrLine((70, 30, 560, 95), "Item", 0.99,
                              [ocr_engine.OcrChar("I", 80, 90, 0.99)])
    g = ocr.LineGeo(line=line, baseline=80.0, x_height=30.0,
                    words=[ocr.WordGeo("Item", 80, 200, 0.99)], rgb="000000", stroke=4,
                    ink_top=40, ink_bot=80, size_px=60.0, thresh=128.0)
    assert ocr.find_bullets(img, [g]) == 1
    assert g.words[0].text == "•" and g.words[0].conf == ocr.BULLET_CONF


def test_uniform_fill_is_rule_based() -> None:
    rng = np.random.default_rng(0)
    img = np.clip(rng.normal(240, 4, (300, 400, 3)), 0, 255).astype(np.uint8)
    cv2.putText(img, "Text", (60, 180), cv2.FONT_HERSHEY_SIMPLEX, 3, (20, 20, 20), 8)
    mask = cv2.dilate(((img[:, :, 0] < 128) * 255).astype(np.uint8), np.ones((5, 5), np.uint8))
    out, st = inpaint.fill(img, mask)
    assert st["lama"] == 0 and st["uniform"] >= 1
    assert out[mask > 0].mean() == pytest.approx(240, abs=3)
    assert (out[mask == 0] == img[mask == 0]).all()


# --------------------------------------------------------------------------- models


@needs_lama
def test_lama_on_a_gradient() -> None:
    yy, xx = np.mgrid[0:400, 0:600]  # stripes: a background no rule fill can fake
    wave = 120 + 70 * np.sin(xx / 9.0)
    clean = np.dstack([wave, wave * 0.8 + 30, 200 - yy * 0.2]).astype(np.uint8)
    img = clean.copy()
    cv2.putText(img, "Hello", (100, 230), cv2.FONT_HERSHEY_SIMPLEX, 3, (0, 0, 0), 10)
    mask = cv2.dilate(((img.sum(axis=2) < 30) * 255).astype(np.uint8), np.ones((7, 7), np.uint8))
    out, st = inpaint.fill(img, mask)
    assert st["lama"] >= 1 and st["lama_provider"] == runtime.CPU
    err = np.abs(out.astype(int) - clean.astype(int))[mask > 0].mean()
    assert err < 30, err


@needs_ocr
def test_ocr_reads_a_rendered_line() -> None:
    img = np.full((140, 900, 3), 255, np.uint8)
    cv2.putText(img, "Quarterly report 2026", (30, 90), cv2.FONT_HERSHEY_SIMPLEX, 2.0,
                (0, 0, 0), 4)
    lines = ocr_engine.ppocr(img)
    assert lines and "Quarterly report 2026" in " ".join(ln.text for ln in lines)
    ln = lines[0]
    assert ln.conf > 0.9 and ln.chars[0].x0 < ln.chars[-1].x0


@needs_ocr
@pytest.mark.corpus
def test_ocr_scanned_page() -> None:
    pdf = SYN / "s01_scanned_deck_it.pdf"
    doc, wd = _doc(pdf)
    ocr.run(doc, pdf, wd)
    page = doc.pages[0]
    text = "".join(g.text for g in page.glyphs)
    assert "Perché visitare il museo?" in text and "3.000" in text
    assert text.count("•") == 4  # found by shape, the detector skips the dots
    assert all(g.source == "ocr" and 0 < g.confidence <= 1 for g in page.glyphs)
    assert {round(g.size) for g in page.glyphs if not g.text.isspace()} == {44, 32, 28}
    assert abs(page.glyphs[0].rotation - 1.0) < 0.1  # the scan's skew, counter-clockwise
    bg = [e for e in page.elements if e.kind == "image"]
    assert len(bg) == 1 and bg[0].is_background and (wd / bg[0].file).is_file()
    rep = json.loads(doc.metadata[ocr.META_KEY])
    assert rep["pages"]["0"]["inpaint"]["lama"] == 0  # paper: the rule fill suffices
    assert rep["models"]["ocr_rec_latin"]["calls"] > 0


@needs_ocr
@pytest.mark.corpus
def test_ocr_outlined_text() -> None:
    pdf = SYN / "d28_rotated_text.pdf"
    doc, wd = _doc(pdf)
    assert doc.pages[0].triage.kind == "outlined_text"
    n_shapes = sum(1 for e in doc.pages[0].elements if e.kind == "shape")
    ocr.run(doc, pdf, wd)
    page = doc.pages[0]
    ocr_text = "".join(g.text for g in page.glyphs if g.source == "ocr")
    assert "Tilted by minus 30 degrees" in ocr_text and "Slightly tilted label" in ocr_text
    rots = {round(g.rotation) for g in page.glyphs if g.source == "ocr"}
    assert rots == {30, 352}
    assert sum(1 for e in page.elements if e.kind == "shape") == n_shapes - 2
    fonts = {doc.fonts[g.font].family for g in page.glyphs if g.source == "ocr"}
    assert fonts == {"Calibri"}  # the document's own family


@needs_ocr
@pytest.mark.corpus
def test_ocr_broken_text_layer() -> None:
    """A text layer that decodes to garbage (simulated with rot13) is re-read."""
    from pdf2pptx.ingest import triage

    pdf = SYN / "d23_long_paragraphs.pdf"
    doc, wd = _doc(pdf)
    page = doc.pages[0]
    want = "".join(g.text for g in page.glyphs if not g.text.isspace())
    for g in page.glyphs:
        g.text = codecs.encode(g.text, "rot13")
    page.triage = triage.classify(page)
    assert page.triage.kind == "broken_text"
    info = ocr.ocr_broken(doc, page, pdf, wd)
    assert info["lines_remapped"] >= 5
    got = "".join(g.text for g in page.glyphs if not g.text.isspace())
    from rapidfuzz.distance import Levenshtein

    assert Levenshtein.normalized_similarity(sorted(want), sorted(got)) > 0.98


@needs_layout
@pytest.mark.corpus
def test_layout_regions() -> None:
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(SYN / "r03_report_figure.pdf"))
    img = cv2.cvtColor(pdf[0].render(scale=100 / 72).to_numpy()[:, :, :3], cv2.COLOR_RGB2BGR)
    pdf.close()
    regs = layout.detect(img)
    labels = {r.label for r in regs}
    assert "chart" in labels and "text" in labels and "footer" in labels
    assert all(0.5 <= r.score <= 1 for r in regs)


def test_role_disagreement_is_logged_not_forced() -> None:
    from pdf2pptx.ir import Page, Paragraph, TextBlock, TextRun

    def block(i: str, role: str, conf: float, y: float) -> TextBlock:
        return TextBlock(id=i, bbox=(50, y, 300, y + 30), role=role, role_confidence=conf,
                         paragraphs=[Paragraph(id=i + "p", bbox=(50, y, 300, y + 30),
                                               runs=[TextRun(id=i + "r", bbox=(50, y, 300,
                                                                               y + 30),
                                                             text="Hello")])])
    page = Page(index=0, width=600, height=400, elements=[
        block("a", "body", 0.6, 20), block("b", "subtitle", 0.6, 100),
        block("c", "body", 0.9, 200)])
    regions = [{"label": "doc_title", "score": 0.9, "bbox": [40, 10, 320, 60]},
               {"label": "text", "score": 0.95, "bbox": [40, 90, 320, 140]},
               {"label": "doc_title", "score": 0.99, "bbox": [40, 190, 320, 240]}]
    dis = layout.compare_roles(page, regions, layout.TRIGGER)
    by = {d["block"]: d for d in dis}
    assert by["a"]["adopted"] and page.elements[0].role == "title"  # unsure rule, sure model
    assert page.elements[0].role_confidence == 0.9
    assert "role:model:pp-doclayoutv3" in page.elements[0].notes[0]
    assert not by["b"]["adopted"] and page.elements[1].role == "subtitle"  # logged only
    assert "c" not in by  # confident rule: the model is not asked


def test_cpu_forced(monkeypatch: pytest.MonkeyPatch) -> None:
    if models.path("ocr_det") is None or runtime.ort_module() is None:
        pytest.skip("models not installed")
    monkeypatch.setenv(runtime.DEVICE_ENV, "cpu")
    runtime.reset()
    try:
        s = runtime.session("ocr_det", ocr_engine._det_warmup)
        assert s is not None and s.provider == runtime.CPU
        assert "forced" in s.stats.fallback
    finally:
        runtime.reset()
