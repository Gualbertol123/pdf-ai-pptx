"""INGEST on a hand-made reportlab PDF with known glyph positions."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pypdfium2 as pdfium
import pytest
from PIL import Image as PILImage
from reportlab.lib.colors import Color as RLColor
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

from pdf2pptx.build.lines import line_blocks
from pdf2pptx.ingest import geom
from pdf2pptx.ingest.parse import clean_family, parse_document
from pdf2pptx.ir import Image, Shape

H = 792.0
TEXT = "Hello World"


@pytest.fixture(scope="module")
def known_pdf(tmp_path_factory: pytest.TempPathFactory) -> Path:
    d = tmp_path_factory.mktemp("parse")
    png = d / "img.png"
    PILImage.new("RGB", (40, 20), (10, 200, 30)).save(png)
    pdf = d / "known.pdf"
    c = canvas.Canvas(str(pdf), pagesize=(612, H))
    c.setFillColor(RLColor(0.2, 0.4, 0.6))
    c.setFont("Helvetica", 24)
    c.drawString(72, 700, TEXT)
    c.setFillColor(RLColor(1, 0, 0))
    c.rect(100, 100, 100, 50, stroke=0, fill=1)
    c.drawImage(str(png), 400, 500, 100, 50)
    c.saveState()
    c.translate(300, 400)
    c.rotate(30)
    c.setFillColor(RLColor(0, 0, 0))
    c.setFont("Helvetica", 18)
    c.drawString(0, 0, "Rotated line")
    c.restoreState()
    c.showPage()
    c.setPageRotation(90)
    c.setFont("Helvetica", 24)
    c.drawString(72, 500, "Turned")
    c.showPage()
    c.save()
    return pdf


@pytest.fixture(scope="module")
def doc(known_pdf: Path, tmp_path_factory: pytest.TempPathFactory):  # type: ignore[no-untyped-def]
    return parse_document(known_pdf, tmp_path_factory.mktemp("work"))


def test_glyph_positions_size_and_color(doc) -> None:  # type: ignore[no-untyped-def]
    page = doc.pages[0]
    glyphs = [g for g in page.glyphs if g.size > 20 and g.rotation == 0]
    assert "".join(g.text for g in glyphs) == TEXT.replace(" ", "") or "".join(
        g.text for g in glyphs) == TEXT
    nonspace = [g for g in glyphs if not g.text.isspace()]
    expected = [72 + stringWidth(TEXT[:i], "Helvetica", 24) for i, ch in enumerate(TEXT)
                if not ch.isspace()]
    for g, x in zip(nonspace, expected, strict=True):
        assert abs(g.origin_x - x) <= 0.5, (g.text, g.origin_x, x)
        assert abs(g.baseline - (H - 700)) <= 0.5
        assert abs(g.size - 24) <= 0.1
        assert g.color.rgb == "336699"
    assert doc.fonts[nonspace[0].font].family == "Helvetica"


def test_rotated_line_detected(doc) -> None:  # type: ignore[no-untyped-def]
    page = doc.pages[0]
    rot = [g for g in page.glyphs if abs(g.rotation - 30) < 0.5]
    assert "".join(g.text for g in rot if not g.text.isspace()) == "Rotatedline"
    first = rot[0]
    assert abs(first.origin_x - 300) <= 0.5 and abs(first.baseline - (H - 400)) <= 0.5
    blocks = [b for b in line_blocks(page, doc.fonts) if b.rotation]
    assert len(blocks) == 1
    assert abs(blocks[0].rotation - 330) < 0.5  # PowerPoint: clockwise
    assert "".join(r.text for r in blocks[0].paragraphs[0].runs) == "Rotated line"


def test_paths_and_images(doc) -> None:  # type: ignore[no-untyped-def]
    page = doc.pages[0]
    shapes = [e for e in page.elements if isinstance(e, Shape)]
    red = [s for s in shapes if s.fill.color and s.fill.color.rgb == "FF0000"]
    assert len(red) == 1
    assert np.allclose(red[0].bbox, (100, H - 150, 200, H - 100), atol=0.5)
    imgs = [e for e in page.elements if isinstance(e, Image)]
    assert len(imgs) == 1
    assert np.allclose(imgs[0].bbox, (400, H - 550, 500, H - 500), atol=0.5)
    assert imgs[0].pixel_size == (40, 20)
    # z-order is content order: text (1st object) < rect < image
    text_z = page.glyphs[0].z
    assert text_z < red[0].z < imgs[0].z


def test_page_rotation(doc, known_pdf: Path) -> None:  # type: ignore[no-untyped-def]
    page = doc.pages[1]
    assert page.rotation == 90
    # reportlab stores the rotated page as a 792x612 mediabox + /Rotate 90
    assert (page.width, page.height) == (612, H)
    g = page.glyphs[0]
    assert abs(g.origin_x - 500) <= 0.5 and abs(g.baseline - 72) <= 0.5
    assert abs(g.rotation - (-90)) < 0.5
    # independent check: pdfium's own (rotated) render has ink inside our glyph boxes
    pdf = pdfium.PdfDocument(str(known_pdf))
    img = np.asarray(pdf[1].render(scale=1).to_pil().convert("L"))
    pdf.close()
    for gl in page.glyphs:
        x0, y0, x1, y1 = (int(v) for v in gl.bbox)
        assert img[y0:y1 + 1, x0:x1 + 1].min() < 128, gl.text


def test_matrix_helpers() -> None:
    m = geom.page_matrix((0, 0, 100, 200), 270)
    assert geom.apply(m, 0, 0) == (200, 100)
    assert geom.apply(m, 100, 0) == (200, 0)
    assert geom.norm_angle(270) == -90
    assert clean_family("ABCDEF+Calibri-Bold") == "Calibri"
    assert clean_family("TimesNewRomanPS-BoldItalicMT") == "TimesNewRoman"
