"""M7 document structure: links / notes / sections, build-ups, theme, masters and
layouts, the structured writer, build expansion for scoring, and (PowerPoint) that
structured decks open, animate and render like the flat deck."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from pdf2pptx.analyze import builds, links, masters, theme
from pdf2pptx.build.structure import read_sections
from pdf2pptx.build.writer import TextOptions, write_pptx
from pdf2pptx.config import load_config
from pdf2pptx.ingest.parse import OUTLINE_KEY
from pdf2pptx.ir import (
    Annotation,
    Color,
    Document,
    Fill,
    Image,
    Page,
    Paragraph,
    Shape,
    TextBlock,
    TextRun,
)
from pdf2pptx.verify.expand import expand_builds

from .conftest import CORPUS


def _tb(tid: str, box: tuple[float, float, float, float], texts: list[str], *, size: float = 20,
        role: str = "body", color: str = "000000") -> TextBlock:
    paras = []
    y = box[1]
    for k, t in enumerate(texts):
        pb = (box[0], y, box[2], y + size * 1.2)
        paras.append(Paragraph(id=f"{tid}_p{k}", bbox=pb, runs=[TextRun(
            id=f"{tid}_r{k}", bbox=pb, text=t, font="Arial", size=size,
            color=Color(rgb=color))]))
        y += size * 1.5
    return TextBlock(id=tid, bbox=box, paragraphs=paras, role=role, role_confidence=0.9)


def _analyzed(name: str):  # type: ignore[no-untyped-def]
    from pdf2pptx.pipeline import analyzed

    cfg = load_config()
    return analyzed(CORPUS / "synthetic" / f"{name}.pdf", cfg, workers=1)


# --------------------------------------------------------------------------- links


def test_sections_from_nested_outline_and_not_from_slide_bookmarks() -> None:
    doc = Document(pages=[Page(index=i, width=960, height=540) for i in range(6)])
    doc.metadata[OUTLINE_KEY] = json.dumps([
        [0, "Intro", 0], [1, "Slide 1: A", 0], [1, "Slide 2: B", 1],
        [0, "Results", 2], [1, "Slide 3", 2], [1, "Slide 4", 3],
        [0, "Plans", 4], [1, "Slide 5", 4], [1, "Slide 6", 5]])
    secs = links.sections_from_outline(doc)
    assert [(s.name, s.first_page) for s in secs] == [("Intro", 0), ("Results", 2), ("Plans", 4)]
    # PowerPoint's per-slide bookmarks of a deck without sections are not sections
    doc.metadata[OUTLINE_KEY] = json.dumps([[0, f"Diapositiva {i + 1}: T", i] for i in range(6)])
    assert links.sections_from_outline(doc) == []
    # report chapters: a section per entry, a default first section before the first one
    doc.metadata[OUTLINE_KEY] = json.dumps([[0, "Chapter 1", 2], [0, "Chapter 2", 4]])
    secs = links.sections_from_outline(doc)
    assert [(s.name, s.first_page) for s in secs] == [
        (links.DEFAULT_SECTION, 0), ("Chapter 1", 2), ("Chapter 2", 4)]


def test_comments_become_notes_and_link_annotations_move_onto_shapes() -> None:
    btn = Shape(id="b", bbox=(100, 100, 300, 160), z=2, preset="roundRect",
                fill=Fill(kind="solid", color=Color(rgb="4F81BD")))
    label = _tb("t", (150, 115, 250, 140), ["Go"])
    label.paragraphs[0].runs[0].hyperlink = "#slide=2"
    page = Page(index=0, width=960, height=540, elements=[btn, label], annotations=[
        Annotation(id="a1", bbox=(100, 100, 300, 160), subtype="link", dest_page=1),
        Annotation(id="a2", bbox=(900, 10, 920, 30), subtype="text", contents="Say hello"),
        Annotation(id="a3", bbox=(900, 10, 920, 30), subtype="subtype16", contents="Say hello"),
    ])
    doc = Document(pages=[page, Page(index=1, width=960, height=540)])
    stats = links.analyze_links(doc)
    assert stats["shape_links"] == 1 and btn.hyperlink == "#slide=2"
    assert label.paragraphs[0].runs[0].hyperlink is None  # the button carries it
    assert page.notes_text == "Say hello"  # the popup duplicate is not repeated


def test_producer_detection() -> None:
    doc = Document(pages=[Page(index=0, width=960, height=540)],
                   metadata={"Producer": "Microsoft® PowerPoint® per Microsoft 365"})
    assert links.detect_producer(doc)["app"] == "powerpoint"
    doc = Document(pages=[Page(index=0, width=595, height=842)],
                   metadata={"Producer": "ReportLab PDF Library"})
    info = links.detect_producer(doc)
    assert info["app"] == "other" and not info["slide_producer"]


# --------------------------------------------------------------------------- builds


def _build_doc() -> Document:
    pages = []
    items = ["One", "Two", "Three"]
    for n in range(1, 4):
        title = _tb(f"p{n}_t", (100, 40, 500, 80), ["Our plan"], size=32, role="title")
        body = _tb(f"p{n}_b", (100, 120, 600, 120 + 30 * n), items[:n])
        els = [title, body]
        if n == 3:
            els.append(Image(id="p3_i", bbox=(700, 300, 800, 400), z=5, file="x.png",
                             sha256="abc"))
        pages.append(Page(index=n - 1, width=960, height=540, elements=els))
    return Document(pages=pages)


def test_build_up_pages_collapse_into_steps() -> None:
    doc = _build_doc()
    stats = builds.analyze_builds(doc)
    assert stats == {"build_sequences": 1, "build_pages_absorbed": 2, "build_effects": 3}
    assert [p.build_of for p in doc.pages] == [None, 0, 1]
    final = doc.pages[2]
    body = next(e for e in final.elements if e.id == "p3_b")
    assert builds.paragraph_steps(body) == {1: 1, 2: 2}
    img = next(e for e in final.elements if e.id == "p3_i")
    assert builds.element_step(img) == 2
    assert builds.page_to_slide(doc, True) == [0, 0, 0]
    assert builds.page_to_slide(doc, False) == [0, 1, 2]


def test_different_pages_are_not_a_build() -> None:
    doc = _build_doc()
    doc.pages[0].elements[0] = _tb("p1_t", (100, 40, 500, 80), ["Other title"], size=32,
                                   role="title")
    assert builds.analyze_builds(doc)["build_sequences"] == 1  # only pages 2-3 remain
    assert [p.build_of for p in doc.pages] == [None, None, 1]
    same = _build_doc()
    same.pages[1] = Page(index=1, width=960, height=540, elements=[
        _tb("q_t", (100, 40, 500, 80), ["Our plan"], size=32, role="title"),
        _tb("q_b", (100, 120, 600, 150), ["One"])])
    same.pages = same.pages[:2]
    assert builds.analyze_builds(same)["build_sequences"] == 0  # identical pages


# --------------------------------------------------------------------------- theme


def test_theme_recognises_a_stock_palette_and_references_its_colours() -> None:
    els = [Shape(id=f"s{k}", bbox=(10 + 100 * k, 10, 90 + 100 * k, 90), z=k, preset="rect",
                 fill=Fill(kind="solid", color=Color(rgb=c)))
           for k, c in enumerate(["4F81BD", "C0504D", "9BBB59", "123456"])]
    els.append(_tb("t", (10, 200, 400, 240), ["Hello"], color="000000"))
    doc = Document(pages=[Page(index=0, width=960, height=540, elements=els)])
    theme.analyze_theme(doc)
    assert doc.theme is not None
    c = doc.theme.colors
    assert [c[f"accent{i}"] for i in range(1, 4)] == ["4F81BD", "C0504D", "9BBB59"]
    assert "123456" in c.values()  # a document colour takes an unused palette slot
    assert c["dk1"] == "000000" and c["lt1"] == "FFFFFF"
    assert els[0].fill.color is not None and els[0].fill.color.theme == "accent1"
    assert doc.theme.major_font == doc.theme.minor_font == "Arial"


# --------------------------------------------------------------------------- corpus decks


@pytest.mark.corpus
def test_d17_furniture_goes_to_the_master() -> None:
    doc, _, _ = _analyzed("d17_footer_logo_number")
    assert len(doc.masters) == 1
    m = doc.masters[0]
    kinds = sorted(e.kind for e in m.elements)
    assert kinds == ["image", "shape", "text_block"]  # logo, white background, footer
    for p in doc.pages:
        num = [e for e in p.elements if isinstance(e, TextBlock) and e.role == "page_number"]
        assert num and num[0].placeholder == "sldNum"
        assert sum(1 for e in p.elements if any(n.startswith(masters.MASTER_NOTE)
                                                for n in e.notes)) == 3
        assert p.layout_id == m.layouts[0].id
    assert {"title", "body", "sldNum"} <= set(m.layouts[0].placeholders)


@pytest.mark.corpus
def test_corpus_structure_facts() -> None:
    d18, _, _ = _analyzed("d18_buildup")
    assert [p.build_of for p in d18.pages] == [None, 0, 1, 2, 3]
    d21, _, _ = _analyzed("d21_sections")
    assert [(s.name, s.first_page) for s in d21.sections] == [
        ("Introduction", 0), ("Results", 2), ("Plans", 4)]
    d20, _, _ = _analyzed("d20_speaker_notes")
    truth = json.loads((CORPUS / "synthetic" / "d20_speaker_notes.truth.json").read_text(
        encoding="utf-8"))
    assert [p.notes_text for p in d20.pages] == [s["notes"] for s in truth["slides"]]
    assert d20.metadata.get("Title") == truth["core"]["title"]


@pytest.mark.corpus
def test_structured_writer_and_build_expansion(tmp_path: Path) -> None:
    doc, wd, _ = _analyzed("d17_footer_logo_number")
    pdf = CORPUS / "synthetic" / "d17_footer_logo_number.pdf"
    out = tmp_path / "d17.pptx"
    rep = write_pptx(doc, out, wd, pdf, TextOptions.for_mode("balanced"))
    assert rep.counts["master_elements"] == 3 and rep.counts["placeholders_sldNum"] == 5
    prs = Presentation(str(out))
    master_pics = [s for s in prs.slide_master.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(master_pics) == 1
    for sl in prs.slides:
        assert not [s for s in sl.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
        num = [s for s in sl.shapes if s.is_placeholder
               and s.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER]
        assert len(num) == 1 and b'type="slidenum"' in num[0]._element.xml.encode()
    # the flat deck has everything on the slides, as in M6
    flat = tmp_path / "d17_flat.pptx"
    write_pptx(doc, flat, wd, pdf, TextOptions.for_mode("balanced").flat())
    prs = Presentation(str(flat))
    assert all(any(s.shape_type == MSO_SHAPE_TYPE.PICTURE for s in sl.shapes)
               for sl in prs.slides)
    # build-up: one slide, four paragraph steps; expanded back to five states
    doc, wd, _ = _analyzed("d18_buildup")
    pdf = CORPUS / "synthetic" / "d18_buildup.pdf"
    out = tmp_path / "d18.pptx"
    rep = write_pptx(doc, out, wd, pdf, TextOptions.for_mode("balanced"))
    assert rep.slides == 1 and rep.counts["animations"] == 4
    exp = expand_builds(out, tmp_path / "d18_expanded.pptx")
    assert (exp.slides_in, exp.slides_out) == (1, 5)
    prs = Presentation(str(exp.path))
    bodies = [max((s for s in sl.shapes if s.has_text_frame),
                  key=lambda s: len(s.text_frame.paragraphs)) for sl in prs.slides]
    assert [len(b.text_frame.paragraphs) for b in bodies] == [1, 2, 3, 4, 5]
    with zipfile.ZipFile(exp.path) as z:
        assert z.testzip() is None
    # builds off: every page again
    rep = write_pptx(doc, tmp_path / "d18_all.pptx", wd, pdf,
                     TextOptions.for_mode("balanced", builds=False))
    assert rep.slides == 5 and not rep.counts["animations"]
    # sections survive the writer
    doc, wd, _ = _analyzed("d21_sections")
    out = tmp_path / "d21.pptx"
    write_pptx(doc, out, wd, CORPUS / "synthetic" / "d21_sections.pdf",
               TextOptions.for_mode("balanced"))
    assert read_sections(Presentation(str(out))) == [
        {"name": "Introduction", "slides": [0, 1]}, {"name": "Results", "slides": [2, 3]},
        {"name": "Plans", "slides": [4, 5]}]


def test_slide_number_field_only_where_it_shows_the_pdf_number(tmp_path: Path) -> None:
    doc = _build_doc()
    doc.pages[2].elements = [e for e in doc.pages[2].elements if not isinstance(e, Image)]
    num = _tb("n", (900, 500, 920, 520), ["3"], size=12, role="page_number")
    num.paragraphs[0].runs[0].field_type = "slidenum"
    doc.pages[2].elements.append(num)
    builds.analyze_builds(doc)
    out = tmp_path / "num.pptx"
    write_pptx(doc, out, tmp_path, tmp_path / "none.pdf",
               TextOptions.for_mode("balanced"))
    xml = Presentation(str(out)).slides[0]._element.xml
    assert 'type="slidenum"' not in xml and ">3<" in xml  # slide 1 shows the PDF's "3"


# --------------------------------------------------------------------------- PowerPoint


@pytest.mark.powerpoint
@pytest.mark.corpus
def test_structured_decks_open_animate_and_render_like_the_flat_deck(tmp_path: Path) -> None:
    from pdf2pptx.verify import ppt_com
    from pdf2pptx.verify.deck import render_check

    cfg = load_config()
    eng = ppt_com.shared()
    for name, steps in (("d17_footer_logo_number", None), ("d18_buildup", [4]),
                        ("d21_sections", None), ("d19_hyperlinks", None)):
        doc, wd, _ = _analyzed(name)
        pdf = CORPUS / "synthetic" / f"{name}.pdf"
        out = tmp_path / f"{name}.pptx"
        opts = TextOptions.for_mode("balanced")
        write_pptx(doc, out, wd, pdf, opts)
        with eng.opened(out) as info:
            counts = eng.animation_counts(info)
        if steps is not None:
            assert counts == steps
            continue
        flat = tmp_path / f"{name}_flat.pptx"
        write_pptx(doc, flat, wd, pdf, opts.flat())
        verdict = render_check(out, flat, cfg)
        assert verdict["ok"], (name, verdict)
