"""Table detection (M6) on hand-made pages: ruled grids with merges, fills and borders;
whitespace tables; prose columns and lists that must stay text."""

from __future__ import annotations

from pathlib import Path

import pytest

from pdf2pptx.analyze import tables
from pdf2pptx.ir import Color, Fill, PathSegment, Shape, Stroke, Table

from .test_text_group import FONT, MATCH, WORDS, PageBuilder


def _line(sid: str, x0: float, y0: float, x1: float, y1: float, w: float = 0.75) -> Shape:
    return Shape(id=sid, bbox=(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)), z=5,
                 path=[PathSegment(op="M", pts=[(x0, y0)]), PathSegment(op="L", pts=[(x1, y1)])],
                 stroke=Stroke(color=Color(rgb="000000"), width=w))


def _rect(sid: str, x0: float, y0: float, x1: float, y1: float, rgb: str) -> Shape:
    pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    segs = [PathSegment(op="M", pts=[pts[0]]), *(PathSegment(op="L", pts=[p]) for p in pts[1:]),
            PathSegment(op="Z")]
    return Shape(id=sid, bbox=(x0, y0, x1, y1), z=2, path=segs,
                 fill=Fill(kind="solid", color=Color(rgb=rgb)))


def _tables(b: PageBuilder) -> list[Table]:
    tables.analyze_page(b.page, {"F": FONT}, MATCH)
    return [e for e in b.page.elements if isinstance(e, Table)]


def _cell_text(t: Table, r: int, c: int) -> str:
    cell = next(x for x in t.cells if x.row == r and x.col == c)
    if cell.text is None:
        return ""
    return " ".join("".join(run.text for run in p.runs) for p in cell.text.paragraphs)


def test_ruled_grid_with_merges_fills_and_borders() -> None:
    """4 rows x 3 cols; the header cell "Year" spans two columns; the first column's
    header spans two rows; header fill differs from the body."""
    b = PageBuilder()
    xs, ys = [100, 220, 340, 460], [100, 130, 160, 190, 220]
    b.page.elements += [_rect("hdr", 100, 100, 460, 160, "4F81BD")]
    n = 0
    for y in ys:  # horizontal rules; the one between header rows skips column 0
        n += 1
        x_start = 220 if y == 130 else 100
        b.page.elements.append(_line(f"h{n}", x_start, y, 460, y, 2.0 if y == 160 else 0.75))
    for x in xs:  # vertical rules; the one between the "Year" columns skips row 0
        n += 1
        b.page.elements.append(_line(f"v{n}", x, 130 if x == 340 else 100, x, 220))
    b.text("Area", 106, 120)
    b.text("Year", 226, 120)
    b.text("2025", 226, 150)
    b.text("2026", 346, 150)
    for r, (a, v1, v2) in enumerate([("North", "12", "14"), ("South", "9", "11")]):
        y = 180 + 30 * r
        b.text(a, 106, y)
        b.text(v1, 226, y)
        b.text(v2, 346, y)
    (t,) = _tables(b)
    assert (len(t.row_heights), len(t.col_widths)) == (4, 3)
    merges = sorted((c.row, c.col, c.row_span, c.col_span) for c in t.cells
                    if c.merged_into is None and (c.row_span > 1 or c.col_span > 1))
    assert merges == [(0, 0, 2, 1), (0, 1, 1, 2)]
    assert _cell_text(t, 0, 1) == "Year" and _cell_text(t, 3, 2) == "11"
    head = next(c for c in t.cells if (c.row, c.col) == (0, 0))
    assert head.fill.kind == "solid" and head.fill.color is not None
    assert head.fill.color.rgb == "4F81BD"
    body = next(c for c in t.cells if (c.row, c.col) == (2, 1))
    assert body.fill.kind == "none" and body.borders["top"].width == 2.0
    assert t.header_rows >= 1 and t.confidence >= tables.MIN_CONFIDENCE
    # consumed rules are gone from the page; the text is in the cells only
    assert not any(isinstance(e, Shape) for e in b.page.elements)
    assert all(e.kind != "text_block" for e in b.page.elements)


def test_whitespace_table_with_numbers() -> None:
    b = PageBuilder()
    rows = [("Voce", "Previsto", "Effettivo"), ("Personale", "88", "91"),
            ("Materiali", "12", "10"), ("Affitti", "24", "24")]
    for r, row in enumerate(rows):
        for c, txt in enumerate(row):
            b.text(txt, 100 + 120 * c, 100 + 15 * r, size=10)
    (t,) = _tables(b)
    assert (len(t.row_heights), len(t.col_widths)) == (4, 3)
    assert "table:whitespace" in t.notes
    assert _cell_text(t, 2, 0) == "Materiali" and _cell_text(t, 3, 2) == "24"


def test_prose_columns_are_not_a_table() -> None:
    b = PageBuilder()
    b.wrapped(WORDS, 60, 100, 40)
    b.wrapped(WORDS, 400, 100, 40)
    assert _tables(b) == []


def test_numbered_list_is_not_a_table() -> None:
    b = PageBuilder()
    for k, item in enumerate(["Budget", "Partners", "Timeline", "Evaluation"]):
        b.text(f"{k + 1}.", 100, 100 + 18 * k)
        b.text(item, 130, 100 + 18 * k)
    assert _tables(b) == []


def test_chart_gridlines_without_text_are_not_a_table() -> None:
    b = PageBuilder()
    for k in range(5):
        b.page.elements.append(_line(f"g{k}", 100, 100 + 30 * k, 500, 100 + 30 * k))
    b.page.elements += [_line("ax", 100, 100, 100, 220), _rect("bar", 150, 160, 180, 220, "C00000")]
    b.text("0", 85, 224)
    assert _tables(b) == []


@pytest.mark.corpus
def test_corpus_tables_match_truth() -> None:
    """Every truth table: rows x cols and merges (synthetic decks and the reports)."""
    import json

    from pdf2pptx.config import load_config
    from pdf2pptx.pipeline import analyzed

    cfg = load_config()
    found = 0
    for name in ("d07_table_simple", "d08_table_merged", "d26_kitchen_sink",
                 "d27_italian_report_4x3", "r02_report_tables_it"):
        pdf = next(cfg.corpus_dir.rglob(f"{name}.pdf"))
        truth = json.loads(Path(str(pdf).replace(".pdf", ".truth.json")).read_text("utf-8"))
        doc, _, _ = analyzed(pdf, cfg)
        pages = truth.get("slides") or truth.get("pages")
        for tp, page in zip(pages, doc.pages, strict=True):
            got = sorted((e for e in page.elements if isinstance(e, Table)),
                         key=lambda t: (t.bbox[1], t.bbox[0]))
            want = sorted(tp["tables"], key=lambda t: -1)
            assert len(got) == len(want), name
            for t, w in zip(got, want, strict=True):
                assert (len(t.row_heights), len(t.col_widths)) == (w["rows"], w["cols"]), name
                merges = sorted([c.row, c.col, c.row_span, c.col_span] for c in t.cells
                                if c.merged_into is None and (c.row_span > 1 or c.col_span > 1))
                assert merges == sorted(w.get("merges", [])), name
                found += 1
    assert found == 6  # d07, d08, d26, d27 and the two tables of r02
