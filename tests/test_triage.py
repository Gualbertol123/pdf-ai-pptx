from __future__ import annotations

import json
from pathlib import Path

import pytest

from pdf2pptx.ingest.parse import parse_document
from pdf2pptx.ingest.triage import bigram_score, classify
from pdf2pptx.ir import Glyph, Page

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


def _page(text: str) -> Page:
    gl = [Glyph(id=f"g{i}", bbox=(i * 6.0, 0, i * 6.0 + 5, 10), text=ch, font="F", size=10)
          for i, ch in enumerate(text)]
    return Page(index=0, width=600, height=800, glyphs=gl)


def test_bigram_score_separates_garbled_text() -> None:
    good, _ = bigram_score("The quarterly report shows strong growth in every region.")
    bad, _ = bigram_score("Wkh txduwhuob uhsruw vkrzv vwurqj jurzwk lq hyhub uhjlrq.")
    assert good > 0.85 > 0.6 > bad


def test_private_use_is_broken_text() -> None:
    assert classify(_page("" * 10)).kind == "broken_text"
    assert classify(_page("A perfectly normal sentence with ordinary words.")).kind == \
        "born_digital"


def test_garbled_text_is_broken_text() -> None:
    t = classify(_page("Wkh txduwhuob uhsruw vkrzv vwurqj jurzwk lq hyhub uhjlrq dqg pdunhw " * 2))
    assert t.kind == "broken_text"


@pytest.mark.corpus
def test_corpus_triage(tmp_path: Path) -> None:
    manifest = json.loads((CORPUS / "manifest.json").read_text(encoding="utf-8"))
    for e in manifest["files"]:
        doc = parse_document(CORPUS / e["pdf"], tmp_path / e["name"])
        kinds = {classify(p).kind for p in doc.pages}
        if e["category"] == "scanned":
            want = {"scanned"}
        elif "pdf_outlined_text" in e["features"]:
            want = {"outlined_text"}
        else:
            want = {"born_digital"}
        assert kinds == want, (e["name"], kinds)
