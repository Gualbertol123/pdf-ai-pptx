"""Charts (M9): vector chart geometry -> native charts, the writer, the render check,
and the synthetic corpus against its truth."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from pptx import Presentation

from pdf2pptx.analyze import charts
from pdf2pptx.ir import (
    Chart,
    Color,
    Document,
    Fill,
    Group,
    Page,
    Paragraph,
    Shape,
    Stroke,
    TextBlock,
    TextRun,
)

TRUTH_CHARTS = [("d14_chart_bar", 0), ("d14_chart_bar", 1), ("d15_chart_line", 0),
                ("d16_chart_pie", 0), ("d26_kitchen_sink", 0), ("d27_italian_report_4x3", 1)]


def _tb(bid: str, text: str, x0: float, y0: float, x1: float, y1: float) -> TextBlock:
    run = TextRun(id=bid + "_r", bbox=(x0, y0, x1, y1), text=text, font="Calibri", size=12.0)
    para = Paragraph(id=bid + "_p", bbox=(x0, y0, x1, y1), runs=[run])
    return TextBlock(id=bid, bbox=(x0, y0, x1, y1), paragraphs=[para], source="rule")


def _hline(sid: str, x0: float, x1: float, y: float) -> Shape:
    return Shape(id=sid, bbox=(x0, y, x1, y), preset="line", stroke=Stroke(
        color=Color(rgb="808080")), source="rule")


def _rect(sid: str, x0: float, y0: float, x1: float, y1: float, rgb: str) -> Shape:
    return Shape(id=sid, bbox=(x0, y0, x1, y1), preset="rect",
                 fill=Fill(kind="solid", color=Color(rgb=rgb)), source="rule")


def column_page() -> Page:
    """Two series x three categories; axis 0..30 step 10 (y 400 -> 100); a legend."""
    els: list = []
    for k, v in enumerate((0, 10, 20, 30)):
        y = 400 - k * 100
        els.append(_tb(f"a{k}", str(v), 40, y - 8, 60, y + 8))
        els.append(_hline(f"g{k}", 70, 670, y))
    vals = {"4F81BD": (12, 25, 7), "C0504D": (20, 15, 30)}
    for s, (rgb, vs) in enumerate(vals.items()):
        for i, v in enumerate(vs):
            x0 = 70 + i * 200 + 50 + s * 50
            els.append(_rect(f"b{s}{i}", x0, 400 - v * 10, x0 + 50, 400, rgb))
    for i, name in enumerate(("North", "South", "East")):
        cx = 70 + i * 200 + 100
        els.append(_tb(f"c{i}", name, cx - 20, 410, cx + 20, 426))
    for s, (rgb, name) in enumerate((("4F81BD", "2025"), ("C0504D", "2026"))):
        x = 300 + s * 80
        els.append(_rect(f"sw{s}", x, 446, x + 8, 454, rgb))
        els.append(_tb(f"l{s}", name, x + 12, 442, x + 40, 458))
    els.append(_tb("t", "Unrelated text far away", 700, 20, 900, 40))
    els.append(Group(id="grp", bbox=(70, 100, 670, 400), child_ids=["g0", "g1", "g2"]))
    return Page(index=0, width=960, height=540, elements=els)


def test_parse_number_conventions() -> None:
    assert charts.parse_number("1,250") == 1250
    assert charts.parse_number("1.250") == 1250
    assert charts.parse_number("0,5") == 0.5
    assert charts.parse_number("42%") == 42
    assert charts.parse_number("-3") == -3
    assert charts.parse_number("€ 12") == 12
    assert charts.parse_number("North") is None
    assert charts.parse_number("2024 2025") is None


def test_column_chart_from_geometry() -> None:
    page = column_page()
    stats = charts.analyze_page(page)
    assert stats.get("charts") == 1
    ch = next(e for e in page.elements if isinstance(e, Chart))
    assert ch.chart_type == "column_clustered"
    assert ch.categories == ["North", "South", "East"]
    assert [s.name for s in ch.series] == ["2025", "2026"]
    assert ch.series[0].values == [12, 25, 7] and ch.series[1].values == [20, 15, 30]
    assert ch.series[0].color is not None and ch.series[0].color.rgb == "4F81BD"
    assert ch.axes["value"]["min"] == 0 and ch.axes["value"]["max"] == 30
    assert ch.has_legend and ch.legend_pos == "bottom"
    # the parts left the page and are kept as the fallback; other text stays
    ids = {e.id for e in page.elements}
    assert "t" in ids and not ids & {"a0", "b00", "c1", "sw0", "l1", "g3", "grp"}
    assert {e.id for e in ch.fallback} >= {"a0", "b00", "c1", "sw0", "l1", "grp"}
    # restore puts them back
    assert charts.restore(page, ch.id, "test")
    assert not any(isinstance(e, Chart) for e in page.elements)
    assert {"a0", "b00", "c1", "sw0", "l1", "grp"} <= {e.id for e in page.elements}


def test_partial_text_block_keeps_the_page() -> None:
    page = column_page()
    # the legend's text block also holds other words: no chart (nothing changes)
    lab = next(e for e in page.elements if e.id == "l0")
    lab.paragraphs.append(Paragraph(id="extra", bbox=(312, 470, 400, 486), runs=[
        TextRun(id="x", bbox=(312, 470, 400, 486), text="note", font="Calibri", size=12)]))
    lab.bbox = (312, 442, 400, 486)
    before = [e.id for e in page.elements]
    charts.analyze_page(page)
    after = [e.id for e in page.elements]
    assert before == after


def test_no_chart_from_a_plain_number_list() -> None:
    els = [_tb(f"n{k}", str(k), 40, 100 + 30 * k, 60, 116 + 30 * k) for k in range(5)]
    page = Page(index=0, width=960, height=540, elements=els)
    assert charts.analyze_page(page) == {}


def test_writer_native_chart_and_determinism(tmp_path: Path) -> None:
    from pdf2pptx.build.writer import TextOptions, write_pptx

    page = column_page()
    charts.analyze_page(page)
    doc = Document(pages=[page])
    a, b = tmp_path / "a.pptx", tmp_path / "b.pptx"
    write_pptx(doc, a, tmp_path, tmp_path / "none.pdf", TextOptions().flat())
    write_pptx(doc, b, tmp_path, tmp_path / "none.pdf", TextOptions().flat())
    assert a.read_bytes() == b.read_bytes()  # embedded workbook timestamps normalised
    prs = Presentation(str(a))
    frames = [s for s in prs.slides[0].shapes if getattr(s, "has_chart", False) and s.has_chart]
    assert len(frames) == 1
    chart = frames[0].chart
    assert list(chart.plots[0].categories) == ["North", "South", "East"]
    assert [list(s.values) for s in chart.plots[0].series] == [[12, 25, 7], [20, 15, 30]]
    assert chart.value_axis.maximum_scale == 30


def test_chart_text_for_scoring(tmp_path: Path) -> None:
    from pdf2pptx.build.writer import TextOptions, write_pptx
    from pdf2pptx.verify.pptx_stats import deck_stats

    page = column_page()
    charts.analyze_page(page)
    out = tmp_path / "c.pptx"
    write_pptx(Document(pages=[page]), out, tmp_path, tmp_path / "none.pdf",
               TextOptions().flat())
    st = deck_stats(out).slides[0]
    assert st.n_charts == 1 and st.chart_boxes
    # ticks ascending, categories, legend (PowerPoint's PDF reading order)
    assert "0\n10\n20\n30\nNorth\nSouth\nEast\n2025\n2026" in st.text


@pytest.mark.corpus
def test_corpus_charts_match_truth() -> None:
    """Chart type 100% right; >= 90% of the values within 2% of the axis range."""
    from pdf2pptx.config import load_config
    from pdf2pptx.pipeline import analyzed

    cfg = load_config()
    hit = total = 0
    for name, i in TRUTH_CHARTS:
        pdf = cfg.corpus_dir / "synthetic" / f"{name}.pdf"
        truth = json.loads((cfg.corpus_dir / "synthetic" / f"{name}.truth.json")
                           .read_text("utf-8"))
        doc, _, _ = analyzed(pdf, cfg)
        got = [e for e in doc.pages[i].elements if isinstance(e, Chart)]
        want = truth["slides"][i]["charts"]
        assert len(got) == len(want) == 1, name
        ch, tw = got[0], want[0]
        assert ch.chart_type == tw["type"].lower(), name
        assert ch.categories == tw["categories"], name
        rng = max(abs(v) for s in tw["series"] for v in s["values"])
        for k, s in enumerate(tw["series"]):
            vals = ch.series[0].values if len(ch.series) == 1 and k == 0 else ch.series[k].values
            for g, w in zip(vals, s["values"], strict=True):
                total += 1
                hit += g is not None and abs(g - w) <= 0.02 * rng
    assert total and hit / total >= 0.9


@pytest.mark.corpus
@pytest.mark.powerpoint
def test_charts_are_native_in_powerpoint_and_render_check(tmp_path: Path) -> None:
    from pdf2pptx.config import load_config
    from pdf2pptx.pipeline import convert
    from pdf2pptx.verify import chart_check, ppt_com

    cfg = load_config({"work_dir": str(tmp_path / "w")})
    pdf = cfg.corpus_dir / "synthetic" / "d16_chart_pie.pdf"
    out = tmp_path / "d16.pptx"
    res = convert(pdf, out, cfg, tune=False)
    assert res.charts and res.charts["kept"] == ["p0_c1"] and not res.charts["reverted"]
    with ppt_com.shared().opened(out.resolve()) as info:
        native = ppt_com.shared().charts(info)
    assert native == [[("Chart p0_c1", True, 5)]]  # xlPie
    # a failing render check puts the shapes back
    old = chart_check.MIN_SSIM
    try:
        chart_check.MIN_SSIM = 1.01
        res2 = convert(pdf, tmp_path / "d16b.pptx", cfg, tune=False)
    finally:
        chart_check.MIN_SSIM = old
    assert res2.charts and res2.charts["reverted"] and not res2.counts.get("charts")
    prs = Presentation(str(tmp_path / "d16b.pptx"))
    assert not any(getattr(s, "has_chart", False) and s.has_chart for s in prs.slides[0].shapes)
