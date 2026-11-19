"""Deck mode (M10): outline, number checker, grounding, part planning, the extractive
fallback, the build, and the audit. The model is a fake server; the PowerPoint fit
check runs only with PowerPoint."""

from __future__ import annotations

import dataclasses
import json
from pathlib import Path
from typing import Any

import pytest

from pdf2pptx.config import load_config
from pdf2pptx.ir import (
    Document,
    Image,
    Page,
    Paragraph,
    Table,
    TableCell,
    TextBlock,
    TextRun,
)
from pdf2pptx.llm import deck
from pdf2pptx.llm import scheduler as sch

from .conftest import CORPUS


def _tb(bid: str, texts: list[str], y: float, size: float = 10.0, role: str = "body",
        bold: bool = False) -> TextBlock:
    paras = [Paragraph(id="", bbox=(0, 0, 0, 0), runs=[TextRun(
        id="", bbox=(0, 0, 0, 0), text=t, size=size, bold=bold)]) for t in texts]
    return TextBlock(id=bid, bbox=(50, y, 500, y + 20 * len(texts)), paragraphs=paras,
                     role=role, reading_order=int(y))


def toy_doc() -> Document:
    cells = []
    rows = [["Site", "2025"], ["North", "1.250"], ["South", "980"]]
    for r, row in enumerate(rows):
        for c, t in enumerate(row):
            cells.append(TableCell(row=r, col=c, text=_tb(f"c{r}{c}", [t], 0)))
    t = Table(id="p1_t1", bbox=(60, 300, 260, 360), row_heights=[20, 20, 20],
              col_widths=[100, 100], cells=cells)
    p0 = Page(index=0, width=595, height=842, elements=[
        _tb("p0_b1", ["Annual report 2026"], 40, size=18, role="title"),
        _tb("p0_b2", ["Growth", "The program grew to 5 sites and 42 volunteers in 2026."],
            100, bold=True),
        _tb("p0_b3", ["Families liked the evening sessions; attendance rose by 12%."], 200),
        _tb("p0_f", ["Annual report 2026"], 800, size=8, role="footer"),
    ])
    p1 = Page(index=1, width=595, height=842, elements=[
        _tb("p1_b1", ["Costs"], 60, size=13, bold=True),
        _tb("p1_b2", ["Spending stayed within the plan at 88 thousand euro."], 100),
        _tb("p1_c", ["Table 1. Participants per site"], 280, size=9, role="caption"),
        t,
    ])
    return Document(source_path="toy.pdf", pages=[p0, p1])


def test_outline_sections_captions_and_skips() -> None:
    o = deck.build_outline(toy_doc())
    assert o.title == "Annual report 2026" and o.lang == "en"
    titles = [s.title for s in o.sections]
    assert titles == ["Growth", "Costs"]
    ids = [i.id for s in o.sections for i in s.items]
    assert "p0_f" not in ids  # footers are not content
    assert o.visuals["p1_t1"].caption.startswith("Table 1")
    assert "Participants per site" in o.texts["p1_t1"]  # the caption counts as its text
    assert o.pages_of["p1_b2"] == 1


@pytest.mark.parametrize("text,source,missing", [
    ("rose by 12%", "attendance rose by 12%", []),
    ("1,250 visits", "North 1.250", []),  # Italian thousands vs English
    ("1250 visits", "North 1.250", []),
    ("0,5 points", "a gain of 0.5", []),
    ("13 sites", "5 sites and 42 volunteers", ["13"]),
    ("in 2027", "in 2026", ["2027"]),
])
def test_number_checker(text: str, source: str, missing: list[str]) -> None:
    assert deck.missing_numbers(text, source) == missing


def test_grounder_drops_flags_and_recites() -> None:
    o = deck.build_outline(toy_doc())
    g = deck.Grounder(o)
    ok = g.bullet("The program grew to 5 sites", ["p0_b2"], "t")
    assert ok is not None and ok.sources == ["p0_b2"] and ok.pages == [1] and not ok.flags
    assert g.bullet("The program grew to 7 sites", ["p0_b2"], "t") is None  # invented
    assert g.bullet("Totally unrelated words here", ["nope"], "t") is None  # no source
    moved = g.bullet("North had 1.250 participants", ["p0_b2"], "t")  # number elsewhere
    assert moved is not None and "p1_t1" in moved.sources and moved.flags
    re = g.bullet("Families liked the evening sessions", [], "t")  # missing cite: regrounded
    assert re is not None and re.sources == ["p0_b3"] and "regrounded" in re.flags
    notes = g.clean_text("Spending stayed within the plan. Costs will double in 2030. "
                         "This shows remarkable managerial excellence overall.", "n")
    assert notes == "Spending stayed within the plan."
    reasons = {d["reason"].split(":")[0] for d in g.dropped}
    assert {"no valid source", "numbers not in the source",
            "words mostly not in the source"} <= reasons


def test_group_units_and_allocate() -> None:
    units = [deck.Unit("x" * n, {str(k)}, set(), (k + 1, k + 1))
             for k, n in enumerate([5000, 9000, 9000, 4000, 20000, 1000])]
    groups = deck.group_units(units, 15000, 10)
    assert [sum(len(u.text) for u in g) for g in groups] == [14000, 13000, 20000, 1000]
    assert sum(deck.allocate([14000, 13000, 21000], 11)) == 11
    assert deck.allocate([14000, 13000, 21000], 11) == [3, 3, 5]
    assert deck.allocate([1, 1, 1, 1], 2) == [1, 1, 1, 1]  # at least one each
    assert len(deck.group_units(units, 15000, 2)) == 2


def test_extractive_is_grounded_and_approaches_n() -> None:
    o = deck.build_outline(toy_doc())
    slides = deck.extractive(o, 3)
    assert 2 <= len(slides) <= 3
    for s in slides:
        for b in s.bullets:
            assert any(b.text.rstrip(" …") in o.texts[i] or b.text[:20] in o.texts[i]
                       for i in b.sources)
            assert not deck.missing_numbers(b.text, o.all_text())


class PlanServer:
    """Fake llama-server: answers the deck plan with good and bad content."""

    def __init__(self, name: str, **kw: Any) -> None:
        from pdf2pptx.llm.server import ServerStats

        self.name, self.plan, self.mmproj_cpu_fallback = name, None, False
        self.stats = ServerStats()

    def start(self, **kw: Any) -> PlanServer:
        return self

    def stop(self) -> None:
        pass

    def chat(self, prompt: str, **kw: Any) -> dict[str, Any]:
        out = {"title": "Annual report 2026", "subtitle": "The program in 2026",
               "slides": [
                   {"title": "Growth in 2026", "visual": "", "notes":
                    "The program grew to 5 sites. Revenue will reach 9 million.",
                    "bullets": [
                        {"text": "The program grew to 5 sites and 42 volunteers",
                         "sources": ["p0_b2"]},
                        {"text": "Attendance rose by 15%", "sources": ["p0_b3"]},
                        {"text": "Evening sessions were liked", "sources": ["p0_b3"]}]},
                   {"title": "Costs in 2031", "visual": "p1_t1", "notes": "",
                    "bullets": [{"text": "Spending stayed within the plan at 88 thousand "
                                         "euro", "sources": ["p1_b2"]}]},
                   {"title": "Empty", "visual": "p9_x", "notes": "", "bullets": []}]}
        return {"content": json.dumps(out), "json": out, "timings": {},
                "finish_reason": "stop"}


def test_model_plan_validated(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(sch, "LlamaServer", PlanServer)
    monkeypatch.setattr(sch, "model_id", lambda n: f"{n}@abc")
    monkeypatch.setattr(deck, "_text_model", lambda cfg: "Fake")
    cfg = load_config({"work_dir": str(tmp_path), "tier": "lite"})
    o = deck.build_outline(toy_doc())
    plan = deck.plan_deck(o, cfg, 3, use_model=True)
    assert plan.source == "model:Fake"
    assert [s.title for s in plan.slides] == ["Growth in 2026", "Costs"]  # 2031: invented
    texts = [b.text for s in plan.slides for b in s.bullets]
    assert "Attendance rose by 15%" not in texts  # 15 is not in the source
    assert plan.slides[1].visual == "p1_t1"
    assert plan.slides[0].notes == "The program grew to 5 sites."
    assert any(d["text"] == "p9_x" for d in plan.dropped)


def test_build_and_audit_without_powerpoint(tmp_path: Path) -> None:
    src = toy_doc()
    o = deck.build_outline(src)
    plan = deck.DeckPlan(title=o.title, subtitle="", lang="en", source="rule",
                         slides=deck.extractive(o, 3))
    plan.slides[-1].visual = "p1_t1"
    doc = deck.build_deck_document(plan, o, src)
    assert doc.masters and next(p.layout_id for p in doc.pages) == "L_title"
    vis = [e for p in doc.pages for e in p.elements if isinstance(e, Table)]
    assert vis and vis[0].id.endswith("p1_t1") and vis[0].bbox[2] <= deck.RIGHT.x1 + 0.01
    assert vis[0].col_widths[0] > 100  # scaled up with its text
    out = tmp_path / "deck.pptx"
    deck._write(doc, out, tmp_path, Path("toy.pdf"))
    srcs = deck.sources_json(plan, o, {}, Path("toy.pdf"))
    # the audit reads the PDF text layer; the toy has no PDF: patch an empty one
    import pdf2pptx.verify.pdf_text as pt

    orig = pt.page_texts
    pt.page_texts = lambda pdf: []  # type: ignore[assignment]
    try:
        a = deck.audit(Path("toy.pdf"), out, srcs, src, len(doc.pages), com=False)
    finally:
        pt.page_texts = orig  # type: ignore[assignment]
    assert a["bullets"] == sum(len(s.bullets) for s in plan.slides) > 0
    assert a["bullets_with_valid_source"] == a["bullets"]
    assert a["numbers_absent"] == [] and a["slide_count_ok"] and a["ok"]


def test_place_visual_scales_geometry() -> None:
    img = Image(id="i1", bbox=(100, 100, 200, 150), file="x.png")
    out = deck.place_visual(img, deck.Box(0, 0, 400, 400), "s1_")
    assert out.id == "s1_i1" and out.bbox == pytest.approx((50, 8, 350, 158))
    assert img.bbox == (100, 100, 200, 150)  # the source element is untouched
    t = toy_doc().pages[1].elements[-1]
    big = deck.place_visual(t, deck.Box(0, 0, 1000, 1000), "s2_")
    assert big.cells[0].text.paragraphs[0].runs[0].size == pytest.approx(10 * 2.2)


@pytest.mark.corpus
@pytest.mark.powerpoint
def test_deck_end_to_end_extractive(tmp_path: Path) -> None:
    """r02 in deck mode without a model: grounded, no overflow, audit ok."""
    cfg = load_config({"tier": "lite", "deck_slides": 4})
    cfg = dataclasses.replace(cfg, mode="deck")
    out = tmp_path / "r02_deck.pptx"
    res = deck.convert_deck(CORPUS / "synthetic" / "r02_report_tables_it.pdf", out, cfg,
                            use_model=False)
    a = res["audit"]
    assert a["ok"], a
    assert abs(res["slides"] - 4) <= 2 and a["overflow"] == []
    srcs = json.loads(Path(res["sources"]).read_text(encoding="utf-8"))
    assert all(b["sources"] for s in srcs["slides"] for b in s.get("bullets", []))


@pytest.mark.powerpoint
def test_fit_check_detects_overflow(tmp_path: Path) -> None:
    src = toy_doc()
    o = deck.build_outline(src)
    long = "word " * 120
    plan = deck.DeckPlan(title="T", subtitle="", lang="en", source="rule", slides=[
        deck.DeckSlide(title="Overflow", bullets=[
            deck.DeckBullet(long, ["p0_b2"], [1]) for _ in range(5)])])
    doc = deck.build_deck_document(plan, o, src, sizes={"s1_body": 28})
    out = tmp_path / "over.pptx"
    deck._write(doc, out, tmp_path, Path("toy.pdf"))
    over = deck.check_fit(out, doc)
    assert "s1_body" in over and over["s1_body"]["need"] > over["s1_body"]["have"]
    _, _, fit = deck.fit_loop(plan, o, src, out, tmp_path, Path("toy.pdf"), rounds=4,
                              sizes={"s1_body": 28})
    assert fit["shrunk"]["s1_body"] < 28 and fit["overflow"] == {}
