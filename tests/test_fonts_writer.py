"""Font index / matching / embedding and the M4 text writer (paragraphs, bullets, fields)."""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from PIL import Image as PILImage
from pptx import Presentation
from pptx.oxml.ns import qn

from pdf2pptx.build.units import ppt_spacing
from pdf2pptx.build.writer import TextOptions, write_pptx
from pdf2pptx.fonts import embed
from pdf2pptx.fonts import index as fidx
from pdf2pptx.fonts.match import match_document, match_font, name_parts
from pdf2pptx.ir import (
    Bullet,
    Color,
    Document,
    FontRef,
    Image,
    Page,
    Paragraph,
    TextBlock,
    TextRun,
)

CORPUS = Path(__file__).resolve().parents[1] / "corpus"
needs_fonts = pytest.mark.skipif(not fidx.face_index(), reason="no installed fonts indexed")


# --------------------------------------------------------------------------- fonts


@pytest.mark.parametrize("name,parts", [
    ("BCDFEE+ArialMT", ("Arial", [])),
    ("TimesNewRomanPS-BoldMT", ("TimesNewRoman", ["Bold"])),
    ("Calibri-Light", ("Calibri", ["Light"])),
    ("Helvetica-BoldOblique", ("Helvetica", ["Bold", "Oblique"])),
])
def test_name_parts(name: str, parts: tuple[str, list[str]]) -> None:
    assert name_parts(name) == parts


@needs_fonts
@pytest.mark.parametrize("pdf_name,family,bold,italic,source", [
    ("ABCDEF+Calibri-Light", "Calibri Light", False, False, "name"),
    ("TimesNewRomanPS-BoldMT", "Times New Roman", True, False, "name"),
    ("SegoeUI-Semibold", "Segoe UI Semibold", False, False, "name"),
    ("Helvetica-Oblique", "Arial", False, True, "alias"),
    ("Courier", "Courier New", False, False, "alias"),
])
def test_match_by_name(pdf_name: str, family: str, bold: bool, italic: bool,
                       source: str) -> None:
    m = match_font(FontRef(id=pdf_name, pdf_name=pdf_name, family=pdf_name))
    assert (m.family, m.bold, m.italic, m.source) == (family, bold, italic, source)


@needs_fonts
def test_match_by_width_profile() -> None:
    georgia = fidx.best_face("Georgia", False, False)
    if georgia is None:
        pytest.skip("Georgia not installed")
    used = {ch: georgia.widths[ch] for ch in "The quick brown fox jumps over lazy dog"
            if ch in georgia.widths}
    fr = FontRef(id="X", pdf_name="QWERTY+Mystery-Regular", family="Mystery", flags=2,
                 widths=used)
    m = match_font(fr)
    assert m.source == "widths" and m.family == "Georgia" and m.confidence > 0.9
    assert m.distance is not None and m.distance < 0.01


@needs_fonts
def test_face_index_has_metrics_and_is_cached(tmp_path: Path,
                                              monkeypatch: pytest.MonkeyPatch) -> None:
    faces = fidx.face_index()
    arial = fidx.best_face("Arial", False, False)
    assert arial is not None and arial.widths["a"] > 0 and arial.upm > 0
    assert any(f.monospace for f in faces) and any(f.italic for f in faces)
    path = tmp_path / "idx.json"
    monkeypatch.setenv("PDF2PPTX_FONT_INDEX", str(path))
    first = fidx.build_index(path)
    assert path.is_file() and fidx.build_index(path) == first  # reused, deterministic


def test_embedding_permissions_and_candidates() -> None:
    assert embed.program_extension(b"\x00\x01\x00\x00rest") == ".ttf"
    assert embed.program_extension(b"OTTOrest") == ".otf"
    assert embed.program_extension(b"\x01\x00\x04\x02") == ".cff"
    assert embed.embedding_allowed(0) is True
    assert embed.embedding_allowed(0x0002) is False  # restricted
    assert embed.embedding_allowed(0x0004) is True  # preview & print
    assert embed.embedding_allowed(0x0200) is False  # bitmap only
    assert embed.embedding_allowed(None) is None
    doc = Document(fonts={
        "A": FontRef(id="A", pdf_name="ABCDEF+Mystery", family="Mystery", embedded=True,
                     subset=True, fs_type=0, match_source="widths", matched_family="Georgia"),
        "B": FontRef(id="B", pdf_name="Arial", family="Arial", embedded=True, fs_type=0,
                     match_source="name"),
        "C": FontRef(id="C", pdf_name="Locked", family="Locked", embedded=True, fs_type=2,
                     match_source="widths"),
    })
    cands = embed.embed_candidates(doc)
    assert [c["font"] for c in cands] == ["A"]
    assert doc.fonts["C"].embeddable is False and doc.fonts["B"].embeddable is True


@needs_fonts
def test_read_fs_type_of_an_installed_font() -> None:
    arial = fidx.best_face("Arial", False, False)
    assert arial is not None
    assert embed.read_fs_type(Path(arial.path)) == arial.fs_type


@pytest.mark.corpus
def test_ingest_records_widths_and_programs(tmp_path: Path) -> None:
    from pdf2pptx.ingest import parse

    pdf = CORPUS / "synthetic" / "d24_font_families.pdf"
    if not pdf.is_file():
        pytest.skip("corpus not built")
    doc = parse.parse_document(pdf, tmp_path)
    georgia = next(f for f in doc.fonts.values() if "Georgia" in f.pdf_name)
    assert 0.5 < georgia.widths["o"] < 0.6 and georgia.widths["i"] < georgia.widths["o"]
    assert georgia.embedded and georgia.program and (tmp_path / georgia.program).is_file()
    assert georgia.fs_type is not None
    matches = match_document(doc)
    if fidx.face_index():
        assert {m.family for m in matches.values()} >= {"Georgia", "Consolas", "Verdana"}


# --------------------------------------------------------------------------- writer


def test_ppt_spacing_picks_the_closer_rounding() -> None:
    # PowerPoint rounds spcPts to whole points and spcPct to whole percent
    assert ppt_spacing(13.5, 12.0) == ("spcPct", 113000, pytest.approx(13.56))
    assert ppt_spacing(26.0, 19.2) == ("spcPts", 2600, 26.0)
    _, _, eff = ppt_spacing(6.77, 12.0)
    assert abs(eff - 6.77) <= 0.06


def _doc(tmp_path: Path) -> tuple[Document, Path]:
    work = tmp_path / "work"
    work.mkdir()
    buf = io.BytesIO()
    PILImage.new("RGB", (8, 8), (200, 0, 0)).save(buf, "PNG")
    (work / "dot.png").write_bytes(buf.getvalue())

    def run(text: str, **kw: object) -> TextRun:
        return TextRun(id=f"r{text}", bbox=(0, 0, 1, 1), text=text, font="Arial",
                       size=10.0, **kw)  # type: ignore[arg-type]

    paras = [
        Paragraph(id="p1", bbox=(0, 0, 1, 1), runs=[run("Bullet item")], indent_left=18,
                  indent_first=-18, bullet=Bullet(kind="char", char="•", font="Arial",
                                                  color=Color(rgb="C00000"), size_pct=80.0),
                  line_spacing=13.5, line_spacing_rule="exact", space_after=6.0),
        Paragraph(id="p2", bbox=(0, 0, 1, 1), runs=[run("Third")], level=1,
                  indent_left=40, indent_first=-20,
                  bullet=Bullet(kind="autonum", scheme="alphaLcParenR", start_at=3)),
        Paragraph(id="p3", bbox=(0, 0, 1, 1), alignment="justify",
                  runs=[run("E = mc"), run("2", baseline_shift=30.0), run(" exam-"),
                        run("ple")], notes=["hyphen:2:5"]),
        Paragraph(id="p4", bbox=(0, 0, 1, 1), runs=[run("Pic")], indent_left=14,
                  indent_first=-14, bullet=Bullet(kind="picture", image_id="img1")),
        Paragraph(id="p5", bbox=(0, 0, 1, 1), runs=[run("7", field_type="slidenum")]),
    ]
    page = Page(index=0, width=400, height=300, elements=[
        Image(id="img1", bbox=(10, 10, 18, 18), file="dot.png", notes=["bullet_image"]),
        TextBlock(id="b1", bbox=(20, 20, 300, 200), paragraphs=paras),
    ])
    return Document(pages=[page]), work


def _paras(out: Path) -> list[object]:
    s = Presentation(str(out)).slides[0]
    (box,) = [sh for sh in s.shapes if sh.has_text_frame]
    assert not [sh for sh in s.shapes if sh.shape_type == 13]  # bullet image not placed
    return list(box.text_frame.paragraphs)


def test_writer_paragraph_properties_and_bullets(tmp_path: Path) -> None:
    doc, work = _doc(tmp_path)
    out = tmp_path / "t.pptx"
    write_pptx(doc, out, work, tmp_path / "none.pdf", TextOptions.for_mode("fidelity"))
    p1, p2, p3, p4, p5 = (p._p for p in _paras(out))  # type: ignore[attr-defined]
    ppr = p1.pPr
    assert ppr.get("marL") == str(18 * 12700) and ppr.get("indent") == str(-18 * 12700)
    assert ppr.find(qn("a:lnSpc")).find(qn("a:spcPct")).get("val") == "113000"
    assert ppr.find(qn("a:spcAft")).find(qn("a:spcPts")).get("val") == "600"
    assert ppr.find(qn("a:buChar")).get("char") == "•"
    assert ppr.find(qn("a:buFont")).get("typeface") == "Arial"
    assert ppr.find(qn("a:buSzPct")).get("val") == "80000"
    tags = [el.tag.split("}")[1] for el in ppr]
    assert tags == ["lnSpc", "spcAft", "buClr", "buSzPct", "buFont", "buChar"]  # schema order
    an = p2.pPr.find(qn("a:buAutoNum"))
    assert an.get("type") == "alphaLcParenR" and an.get("startAt") == "3"
    assert p2.pPr.get("lvl") == "1"
    assert p3.pPr.get("algn") == "just"
    sup = next(r for r in p3.findall(qn("a:r")) if r.find(qn("a:t")).text == "2")
    assert sup.find(qn("a:rPr")).get("baseline") == "30000"
    assert "".join(t.text for t in p3.iter(qn("a:t"))) == "E = mc2 exam-ple"  # fidelity
    blip = p4.pPr.find(qn("a:buBlip")).find(qn("a:blip"))
    assert blip.get(qn("r:embed"))
    fld = p5.find(qn("a:fld"))
    assert fld is not None and fld.get("type") == "slidenum"
    body = Presentation(str(out)).slides[0].shapes[0].text_frame._txBody.bodyPr
    assert body.find(qn("a:noAutofit")) is not None and body.get("wrap") == "square"


def test_writer_editable_mode(tmp_path: Path) -> None:
    doc, work = _doc(tmp_path)
    out = tmp_path / "e.pptx"
    write_pptx(doc, out, work, tmp_path / "none.pdf", TextOptions.for_mode("editable"))
    p3 = _paras(out)[2]._p  # type: ignore[attr-defined]
    assert "".join(t.text for t in p3.iter(qn("a:t"))) == "E = mc2 example"
    body = Presentation(str(out)).slides[0].shapes[0].text_frame._txBody.bodyPr
    assert body.find(qn("a:normAutofit")) is not None  # shrink text on overflow


@pytest.mark.corpus
def test_structure_matches_truth(tmp_path: Path) -> None:
    """Synthetic decks: paragraphs + levels on >= 95% of slides, fonts on >= 95% of runs."""
    from pdf2pptx.config import Config
    from pdf2pptx.verify import structure

    if not (CORPUS / "manifest.json").is_file():
        pytest.skip("corpus not built")
    cfg = Config(work_dir=tmp_path / "work", out_dir=tmp_path / "out", workers=1)
    res = structure.structure_corpus(cfg)
    s = res["summary"]
    assert s["slides"] >= 50
    assert s["slide_match_rate"] >= 0.95, structure.summary_text(res)
    if fidx.face_index():
        assert s["font_match_rate"] >= 0.95, structure.summary_text(res)
    assert s["bullets_ok"] == s["bullets_checked"]
