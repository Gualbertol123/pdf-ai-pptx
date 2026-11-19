"""Regressions found on real bank results decks (M10 follow-up, UCG / RBI 9M 2026):

* a financial statement whose figure columns are only a space apart (right-aligned,
  edges spread a few points) is one native table, labels included, with a side panel
  left out;
* scattered chart data labels are not a table;
* a title slide with a few real glyphs over a large photo is not a scan;
* when a scanned page is cleaned, shapes the photo used to hide do not cover it.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np

from pdf2pptx.analyze import tables
from pdf2pptx.ingest.triage import classify
from pdf2pptx.ir import Color, Fill, Glyph, Image, Page, PathSegment, Shape, Table
from pdf2pptx.ml import ocr

from .test_text_group import ADV, FONT, MATCH, PageBuilder


def _tables(b: PageBuilder) -> list[Table]:
    tables.analyze_page(b.page, {"F": FONT}, MATCH)
    return [e for e in b.page.elements if isinstance(e, Table)]


def _cell(t: Table, r: int, c: int) -> str:
    cell = next(x for x in t.cells if x.row == r and x.col == c)
    if cell.text is None:
        return ""
    return " ".join("".join(run.text for run in p.runs) for p in cell.text.paragraphs)


def _right(b: PageBuilder, s: str, right: float, y: float, size: float = 10.0) -> None:
    b.text(s, right - len(s) * ADV * size, y, size=size)


def test_tight_numeric_statement_is_one_table() -> None:
    b = PageBuilder()
    b.text("Group", 20, 80, size=28)  # a side panel's title, beside the table
    b.text("P&L", 20, 110, size=28)
    heads = ["1Q25", "2Q25", "3Q25", "4Q25", "9M25"]
    rights = [330, 370, 410, 450, 500]  # one space-width apart at 10 pt
    for k, h in enumerate(heads):
        _right(b, h, rights[k], 60)
    rows = [("Total revenues", ["6,586", "6,115", "6,173", "5,709", "18,874"]),
            ("o/w Fees", ["2,327", "2,149", "2,071", "2,073", "6,547"]),
            ("o/w Trading", ["497", "18", "284", "-285", "798"]),
            ("Operating costs", ["-2,321", "-2,315", "-2,291", "-2,514", "-6,927"]),
            ("Net profit", ["2,771", "3,344", "2,631", "1,833", "8,746"])]
    for r, (label, vals) in enumerate(rows):
        y = 75 + 13 * r
        b.text(label, 180, y, size=10)
        for k, v in enumerate(vals):
            # right edges spread by up to 3 pt, as in the real deck
            _right(b, v, rights[k] + (3 if r % 2 else 0), y)
    (t,) = _tables(b)
    assert "table:numeric" in t.notes
    assert (len(t.row_heights), len(t.col_widths)) == (6, 6)
    assert _cell(t, 0, 1) == "1Q25" and _cell(t, 0, 5) == "9M25"
    assert _cell(t, 1, 0) == "Total revenues" and _cell(t, 1, 1) == "6,586"
    assert _cell(t, 4, 0) == "Operating costs" and _cell(t, 4, 4) == "-2,514"
    assert t.bbox[0] > 150  # the side panel stays out


def test_chart_data_labels_are_not_a_table() -> None:
    b = PageBuilder()
    # bar labels at different heights, no row labels
    for k, (v, y) in enumerate([("3.7", 120), ("3.56%", 180), ("3.54%", 150), ("1,716", 90),
                                ("1,856", 200), ("8.67", 140), ("8.64%", 170), ("8.67%", 110)]):
        _right(b, v, 300 + 60 * (k % 4), y)
    assert _tables(b) == []


def _photo_page(cover_h: float, text: str) -> Page:
    page = Page(index=0, width=960, height=540)
    page.elements = [Image(id="img", bbox=(0, 0, 960, cover_h), z=5, file="x.jpg")]
    page.glyphs = [Glyph(id=f"g{i}", bbox=(60 + 20 * i, 250, 78 + 20 * i, 280), text=ch,
                         font="F", size=40) for i, ch in enumerate(text)]
    return page


def test_title_over_a_photo_is_not_a_scan() -> None:
    assert classify(_photo_page(455, "Annex22")).kind == "born_digital"  # 84 % photo
    assert classify(_photo_page(540, "")).kind == "scanned"  # a scan has no text
    assert classify(_photo_page(540, "Annex")).kind == "scanned"  # full-page picture


def test_cleaned_background_drops_shapes_the_photo_hid(tmp_path: Path) -> None:
    pts = [(0, 0), (960, 0), (960, 540), (0, 540)]
    white = Shape(id="white", bbox=(0, 0, 960, 540), z=1,
                  path=[PathSegment(op="M", pts=[pts[0]]),
                        *(PathSegment(op="L", pts=[p]) for p in pts[1:]), PathSegment(op="Z")],
                  fill=Fill(kind="solid", color=Color(rgb="FFFFFF")))
    logo = Shape(id="logo", bbox=(900, 480, 930, 510), z=2, preset="ellipse",
                 fill=Fill(kind="solid", color=Color(rgb="E2001A")))
    page = Page(index=0, width=960, height=540, elements=[
        white, logo, Image(id="img", bbox=(0, 0, 960, 455), z=5, file="x.jpg")])
    rel = "ocr/bg.jpg"
    (tmp_path / "ocr").mkdir()
    (tmp_path / rel).write_bytes(b"jpeg")
    ocr._set_background(page, tmp_path, rel, np.zeros((54, 96, 3), dtype=np.uint8))
    ids = [e.id for e in page.elements]
    assert ids[0] == "p0_ocrbg" and "white" not in ids and "logo" in ids and "img" not in ids
