from __future__ import annotations

import json

import pytest

from pdf2pptx import ir
from pdf2pptx.ir import (
    Annotation,
    Bullet,
    Chart,
    ChartSeries,
    Color,
    Connector,
    Document,
    Equation,
    Fill,
    FontRef,
    Glyph,
    Gradient,
    GradientStop,
    Group,
    Image,
    Layout,
    Master,
    Page,
    Paragraph,
    PathSegment,
    Section,
    Shape,
    Stroke,
    Table,
    TableCell,
    TextBlock,
    TextRun,
    Theme,
    Triage,
)


def _run(i: int = 0) -> TextRun:
    return TextRun(id=f"r{i}", bbox=(1.0, 2.0, 30.5, 14.25), text="Perché è così? ",
                   font="F1", size=11.5, color=Color(rgb="1F497D", alpha=0.5), bold=True,
                   baseline_shift=30.0, hyperlink="https://example.org/", glyph_ids=["g1"])


def _block(i: int = 0) -> TextBlock:
    para = Paragraph(
        id=f"p{i}", bbox=(1.0, 2.0, 100.0, 20.0), runs=[_run(i)], alignment="justify",
        indent_left=18.0, indent_first=-18.0, level=1,
        bullet=Bullet(kind="char", char="•", color=Color(theme="accent1", lum_mod=0.75)),
        line_spacing=14.4, line_spacing_rule="exact", space_before=6.0,
    )
    return TextBlock(id=f"tb{i}", bbox=(1.0, 2.0, 100.0, 40.0), paragraphs=[para], role="title",
                     role_confidence=0.8, autofit="shrink", margins=(0.0, 0.0, 0.0, 0.0),
                     placeholder="title", reading_order=0, source="rule", notes=["merged 3 lines"])


def one_of_each() -> list[ir.AnyElement]:
    grad = Gradient(kind="linear", angle=45.0,
                    stops=[GradientStop(pos=0.0, color=Color(rgb="000000")),
                           GradientStop(pos=1.0, color=Color(rgb="FFFFFF", alpha=0.25))])
    return [
        Glyph(id="g1", bbox=(1.0, 2.0, 8.0, 14.0), text="P", font="F1", size=12.0, baseline=12.0,
              origin_x=1.0, rotation=90.0),
        _run(9),
        Paragraph(id="p9", bbox=(0.0, 0.0, 1.0, 1.0), runs=[_run(8)]),
        _block(),
        Shape(id="s1", bbox=(10.0, 10.0, 110.0, 60.0), z=2, preset="roundRect",
              adjust={"adj": 16667.0}, fill=Fill(kind="gradient", gradient=grad),
              stroke=Stroke(color=Color(rgb="FF0000"), width=1.5, dash="dash", head_end="triangle"),
              rotation=30.0, text=_block(2), hyperlink="#slide=2"),
        Shape(id="s2", bbox=(0.0, 0.0, 50.0, 50.0), path=[
            PathSegment(op="M", pts=[(0.0, 0.0)]),
            PathSegment(op="C", pts=[(10.0, 0.0), (20.0, 5.0), (25.5, 25.5)]),
            PathSegment(op="Z"),
        ], connector=Connector(kind="elbow", start_id="s1", start_site=3, end_id="s3"),
              source="model:qwen3-vl-8b", confidence=0.61),
        Image(id="i1", bbox=(0.0, 0.0, 200.0, 100.0), file="images/i1.png", sha256="ab" * 32,
              crop=(0.1, 0.0, 0.1, 0.0), mask_file="images/i1_mask.png", pixel_size=(640, 320),
              alt_text="A courtyard"),
        Table(id="t1", bbox=(0.0, 0.0, 300.0, 100.0), row_heights=[20.0, 20.0],
              col_widths=[100.0, 200.0], header_rows=1, cells=[
                  TableCell(row=0, col=0, col_span=2, text=_block(3),
                            fill=Fill(kind="solid", color=Color(rgb="DCE6F1")),
                            borders={"top": Stroke(width=0.5)}),
                  TableCell(row=0, col=1, merged_into=(0, 0)),
              ]),
        Chart(id="c1", bbox=(0.0, 0.0, 300.0, 200.0), chart_type="column_clustered",
              categories=["Q1", "Q2"], series=[ChartSeries(name="2026", values=[1.5, None],
                                                           color=Color(rgb="4F81BD"))],
              axes={"value": {"min": 0.0, "max": 60.0, "title": None}}, fallback_ids=["s2"]),
        Equation(id="e1", bbox=(0.0, 0.0, 10.0, 10.0), latex=r"E = mc^2"),
        Group(id="grp1", bbox=(0.0, 0.0, 110.0, 60.0), child_ids=["s1", "s2"], name="group"),
        Annotation(id="a1", bbox=(0.0, 0.0, 10.0, 10.0), subtype="text", contents="note",
                   author="X"),
    ]


def sample_document() -> Document:
    els = one_of_each()
    page = Page(index=0, width=960.0, height=540.0, rotation=90,
                triage=Triage(kind="scanned", reason="no glyphs, image coverage 0.98",
                              confidence=0.9),
                elements=els, glyphs=[els[0]],  # type: ignore[list-item]
                annotations=[els[-1]], renders={"150": "renders/p0_150.png"},  # type: ignore[list-item]
                notes_text="Speaker notes", build_of=None)
    return Document(
        source_path="in.pdf", sha256="cd" * 32, metadata={"producer": "PowerPoint"},
        pages=[page],
        fonts={"F1": FontRef(id="F1", pdf_name="ABCDEF+Calibri-Bold", family="Calibri",
                             bold=True, subset=True, embedded=True, matched_family="Calibri",
                             match_score=0.02)},
        theme=Theme(colors={"accent1": "4F81BD"}, major_font="Calibri Light", minor_font="Calibri"),
        masters=[Master(id="m1", elements=[_block(7)],
                        layouts=[Layout(id="l1", name="Title", placeholders={
                            "title": (10.0, 10.0, 900.0, 80.0)})])],
        sections=[Section(name="Intro", first_page=0)],
    )


def test_every_element_kind_is_covered() -> None:
    kinds = {e.kind for e in one_of_each()}
    assert kinds == set(ir.ELEMENT_TYPES)


def test_round_trip_exact() -> None:
    doc = sample_document()
    s = doc.to_json()
    back = Document.from_json(s)
    assert back == doc
    assert back.to_json() == s
    # indent does not matter for equality
    assert Document.from_json(doc.to_json(indent=2)) == doc


@pytest.mark.parametrize("element", one_of_each(), ids=lambda e: e.kind)
def test_round_trip_each_element(element: ir.AnyElement) -> None:
    page = Page(index=0, width=100.0, height=100.0, elements=[element])
    doc = Document(pages=[page])
    back = Document.from_json(doc.to_json())
    assert back == doc
    assert type(back.pages[0].elements[0]) is type(element)


def test_tuples_and_types_restored() -> None:
    back = Document.from_json(sample_document().to_json())
    el = back.pages[0].elements
    assert isinstance(el[0].bbox, tuple)
    table = el[7]
    assert isinstance(table, Table)
    assert table.cells[1].merged_into == (0, 0)
    shape = el[5]
    assert isinstance(shape, Shape) and shape.path is not None
    assert shape.path[1].pts[2] == (25.5, 25.5)
    assert isinstance(back.masters[0].layouts[0].placeholders["title"], tuple)


def test_kind_is_set_automatically() -> None:
    assert Shape(id="x", bbox=(0, 0, 1, 1)).kind == "shape"
    assert TextBlock(id="x", bbox=(0, 0, 1, 1)).kind == "text_block"


def test_rejects_wrong_version() -> None:
    data = json.loads(Document().to_json())
    data["ir_version"] = "0.1"
    with pytest.raises(ir.IRError):
        Document.from_json(json.dumps(data))


def test_rejects_unknown_kind_and_fields() -> None:
    data = json.loads(sample_document().to_json())
    data["pages"][0]["elements"][0]["kind"] = "hologram"
    with pytest.raises(ir.IRError):
        Document.from_json(json.dumps(data))
    data = json.loads(sample_document().to_json())
    data["pages"][0]["bogus"] = 1
    with pytest.raises(ir.IRError):
        Document.from_json(json.dumps(data))


def test_rejects_nan() -> None:
    doc = Document(pages=[Page(index=0, width=float("nan"), height=1.0)])
    with pytest.raises(ir.IRError):
        doc.to_json()


def test_deterministic_serialisation() -> None:
    assert sample_document().to_json() == sample_document().to_json()
