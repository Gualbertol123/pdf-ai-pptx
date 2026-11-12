"""``pdf2pptx corpus build``: generate the synthetic corpus with ground truth.

Layout (stable names)::

    corpus/synthetic/_assets/            deterministic images
    corpus/synthetic/<name>.pptx         answer key (decks only)
    corpus/synthetic/<name>.pdf          input for the converter
    corpus/synthetic/<name>.truth.json   ground truth
    corpus/manifest.json                 every file with its category
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any

import pypdfium2 as pdfium

from .. import log
from ..verify import ppt_com
from . import reports, scanned, synthetic, truth

LOG = log.get(__name__)
LF = chr(10)  # truth files use LF on every platform (stable hashes, clean git diffs)

SCANNED = {
    # name: (source corpus item, seed)
    "s01_scanned_deck_it": ("d06_italian_text", 501),
    "s02_scanned_report": ("r01_report_two_column", 502),
}


def _write_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=1, sort_keys=True) + "\n",
                    encoding="utf-8", newline=LF)


def _page_count(pdf: Path) -> int:
    doc = pdfium.PdfDocument(str(pdf))
    try:
        return len(doc)
    finally:
        doc.close()


def build(corpus_dir: Path, only: list[str] | None = None) -> dict[str, Any]:
    t0 = time.perf_counter()
    syn = corpus_dir / "synthetic"
    syn.mkdir(parents=True, exist_ok=True)
    asset_set = synthetic.make_assets(syn / "_assets")
    entries: list[dict[str, Any]] = []

    def wanted(name: str) -> bool:
        return not only or name in only

    # 1. decks -> pptx + truth
    decks = [d for d in synthetic.DECKS if wanted(d.name)]
    for spec in decks:
        pptx = syn / f"{spec.name}.pptx"
        synthetic.build_deck(spec, pptx, asset_set)
        t = truth.extract(pptx)
        t.update(name=spec.name, category="synthetic_deck", features=list(spec.features),
                 aspect=spec.aspect, source="python-pptx", truth_pptx=pptx.name)
        _write_json(syn / f"{spec.name}.truth.json", t)
        LOG.info("deck %s: %d slides", spec.name, t["slide_count"])

    # 2. decks -> PDF through PowerPoint (one session for all)
    if decks:
        ppt = ppt_com.shared()
        LOG.info("PowerPoint %s: exporting %d decks to PDF", ppt.version, len(decks))
        for spec in decks:
            with ppt.opened(syn / f"{spec.name}.pptx") as pres:
                ppt.export_pdf(pres, syn / f"{spec.name}.pdf")
    for spec in decks:
        entries.append({
            "name": spec.name,
            "category": "synthetic_deck",
            "features": list(spec.features),
            "pdf": f"synthetic/{spec.name}.pdf",
            "truth": f"synthetic/{spec.name}.truth.json",
            "truth_pptx": f"synthetic/{spec.name}.pptx",
            "pages": _page_count(syn / f"{spec.name}.pdf"),
        })

    # 3. report-style PDFs (reportlab)
    for name, features in reports.REPORTS.items():
        if not wanted(name):
            continue
        pdf = syn / f"{name}.pdf"
        t = reports.build_report(name, pdf, syn / "_assets")
        t.update(name=name, category="report", features=list(features), source="reportlab",
                 truth_pptx=None, slide_count=len(t["pages"]))
        _write_json(syn / f"{name}.truth.json", t)
        entries.append({
            "name": name,
            "category": "report",
            "features": list(features),
            "pdf": f"synthetic/{name}.pdf",
            "truth": f"synthetic/{name}.truth.json",
            "truth_pptx": None,
            "pages": _page_count(pdf),
        })

    # 4. scanned PDFs derived from the files above
    for name, (src, seed) in SCANNED.items():
        if not wanted(name):
            continue
        src_pdf = syn / f"{src}.pdf"
        if not src_pdf.is_file():
            LOG.warning("skip %s: source %s missing", name, src_pdf.name)
            continue
        pdf = syn / f"{name}.pdf"
        n = scanned.make_scanned(src_pdf, pdf, seed=seed)
        src_truth = json.loads((syn / f"{src}.truth.json").read_text(encoding="utf-8"))
        src_truth.update(name=name, category="scanned", derived_from=src,
                         features=["scanned", "rotation", "noise", "jpeg"], scan_dpi=200)
        _write_json(syn / f"{name}.truth.json", src_truth)
        entries.append({
            "name": name,
            "category": "scanned",
            "features": ["scanned", "rotation", "noise", "jpeg"],
            "pdf": f"synthetic/{name}.pdf",
            "truth": f"synthetic/{name}.truth.json",
            "truth_pptx": src_truth.get("truth_pptx") and f"synthetic/{src_truth['truth_pptx']}",
            "derived_from": src,
            "pages": n,
        })

    manifest_path = corpus_dir / "manifest.json"
    if only and manifest_path.is_file():  # partial rebuild: merge into existing manifest
        old = json.loads(manifest_path.read_text(encoding="utf-8"))["files"]
        names = {e["name"] for e in entries}
        entries += [e for e in old if e["name"] not in names]
    entries.sort(key=lambda e: e["name"])
    manifest = {"version": 1, "files": entries}
    _write_json(manifest_path, manifest)
    LOG.info("corpus: %d PDFs in %.1fs", len(entries), time.perf_counter() - t0)
    return manifest


def load_manifest(corpus_dir: Path) -> list[dict[str, Any]]:
    path = corpus_dir / "manifest.json"
    if not path.is_file():
        raise FileNotFoundError(f"{path} not found; run `pdf2pptx corpus build` first")
    files: list[dict[str, Any]] = json.loads(path.read_text(encoding="utf-8"))["files"]
    return files
