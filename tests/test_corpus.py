"""Corpus checks: manifest completeness and deterministic ground truth."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pdf2pptx.corpus import reports, synthetic, truth
from pdf2pptx.corpus.build import load_manifest

from .conftest import CORPUS

pytestmark = pytest.mark.corpus


def test_manifest_complete() -> None:
    files = load_manifest(CORPUS)
    assert len(files) >= 30
    cats = {f["category"] for f in files}
    assert cats == {"synthetic_deck", "report", "scanned"}
    assert sum(f["category"] == "synthetic_deck" for f in files) >= 25
    assert sum(f["category"] == "report" for f in files) == 3
    assert sum(f["category"] == "scanned" for f in files) == 2
    for f in files:
        assert (CORPUS / f["pdf"]).is_file(), f["pdf"]
        assert (CORPUS / f["truth"]).is_file(), f["truth"]
        if f["truth_pptx"]:
            assert (CORPUS / f["truth_pptx"]).is_file()
        assert f["pages"] >= 1


def test_required_features_present() -> None:
    feats = {x for f in load_manifest(CORPUS) for x in f["features"]}
    for needed in ("multilevel", "multi_column", "mixed_fonts", "italian", "table", "merged_cells",
                   "preset", "arrows", "chevron", "connectors", "gradient", "transparency",
                   "images", "bar", "line", "pie", "repeated_footer", "slide_number", "logo",
                   "build_sequence", "hyperlinks", "speaker_notes", "sections", "aspect_4_3",
                   "scanned", "whitespace_table", "caption"):
        assert needed in feats, needed


def test_truth_contents() -> None:
    t = json.loads((CORPUS / "synthetic" / "d08_table_merged.truth.json").read_text("utf-8"))
    tbl = t["slides"][0]["tables"][0]
    assert (tbl["rows"], tbl["cols"]) == (6, 5)
    assert sorted(map(tuple, tbl["merges"])) == [(0, 0, 2, 1), (0, 1, 1, 2), (0, 3, 1, 2)]
    t = json.loads((CORPUS / "synthetic" / "d21_sections.truth.json").read_text("utf-8"))
    assert [s["name"] for s in t["sections"]] == ["Introduction", "Results", "Plans"]
    t = json.loads((CORPUS / "synthetic" / "d17_footer_logo_number.truth.json").read_text("utf-8"))
    assert all(s["has_slide_number_field"] for s in t["slides"])
    t = json.loads((CORPUS / "synthetic" / "d16_chart_pie.truth.json").read_text("utf-8"))
    assert t["slides"][0]["charts"][0]["series"][0]["values"] == [42.0, 28.0, 20.0, 10.0]
    t = json.loads((CORPUS / "synthetic" / "d22_aspect_4x3.truth.json").read_text("utf-8"))
    assert (t["slide_width_pt"], t["slide_height_pt"]) == (720.0, 540.0)


def test_truth_is_deterministic(tmp_path: Path) -> None:
    """Rebuilding every deck yields byte-identical truth (no PowerPoint needed)."""
    assets = synthetic.make_assets(tmp_path / "_assets")
    for spec in synthetic.DECKS:
        pptx = synthetic.build_deck(spec, tmp_path / f"{spec.name}.pptx", assets)
        fresh = truth.extract(pptx)
        stored = json.loads((CORPUS / "synthetic" / f"{spec.name}.truth.json").read_text("utf-8"))
        for key in ("name", "category", "features", "aspect", "source", "truth_pptx"):
            stored.pop(key)
        assert json.loads(json.dumps(fresh, sort_keys=True)) == stored, spec.name


def test_report_truth_is_deterministic(tmp_path: Path) -> None:
    for name in reports.REPORTS:
        a = reports.build_report(name, tmp_path / "a.pdf", tmp_path)
        b = reports.build_report(name, tmp_path / "b.pdf", tmp_path)
        assert a == b
        stored = json.loads((CORPUS / "synthetic" / f"{name}.truth.json").read_text("utf-8"))
        assert stored["pages"] == json.loads(json.dumps(a["pages"]))
