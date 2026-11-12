"""PDF -> PPTX pipeline: ingest -> triage -> analyze -> build.

Each stage's IR is cached as ``work/<sha256>/ir_<stage>.json``. A cached stage is
reused only when its cache key (the code versions of that stage and every stage
before it) matches ``cache.json``. After the build, ``ir_build.json`` holds the
final IR, including pages that fell back to a picture.
"""

from __future__ import annotations

import hashlib
import json
import time
from collections import Counter
from collections.abc import Callable
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from . import log
from .build.writer import write_pptx
from .config import Config
from .ingest import parse, triage
from .ir import Document

LOG = log.get(__name__)

STAGE_VERSIONS = {
    "ingest": parse.PARSER_VERSION,
    "triage": "2.0",
    "analyze": "identity-2.0",
}
STAGES = tuple(STAGE_VERSIONS)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def stage_key(stage: str) -> str:
    keys = []
    for s in STAGES:
        keys.append(f"{s}={STAGE_VERSIONS[s]}")
        if s == stage:
            break
    return ";".join(keys)


def _read_cache_index(work_dir: Path) -> dict[str, str]:
    p = work_dir / "cache.json"
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return {str(k): str(v) for k, v in data.items()}
    except (OSError, ValueError):
        return {}


def _save_ir(work_dir: Path, name: str, doc: Document) -> None:
    p = work_dir / f"ir_{name}.json"
    tmp = p.with_suffix(".tmp")
    tmp.write_text(doc.to_json(indent=0), encoding="utf-8", newline="\n")
    tmp.replace(p)


def run_stage(work_dir: Path, stage: str, fn: Callable[[], Document], use_cache: bool) -> tuple[
        Document, bool]:
    """Load ``ir_<stage>.json`` if its key matches, else run ``fn`` and cache it."""
    key = stage_key(stage)
    index = _read_cache_index(work_dir)
    path = work_dir / f"ir_{stage}.json"
    if use_cache and index.get(stage) == key and path.is_file():
        try:
            return Document.from_json(path.read_text(encoding="utf-8")), True
        except Exception as exc:  # stale or corrupt cache: recompute
            LOG.warning("cache %s unreadable (%s); recomputing", path, exc)
    doc = fn()
    doc.stage = stage
    _save_ir(work_dir, stage, doc)
    index = {k: v for k, v in index.items() if k in STAGES and STAGES.index(k) < STAGES.index(
        stage)}
    index[stage] = key
    (work_dir / "cache.json").write_text(json.dumps(index, indent=1, sort_keys=True),
                                         encoding="utf-8", newline="\n")
    return doc, False


def analyze(doc: Document) -> Document:
    """Identity for M2. M4+ adds text grouping, lists, shapes, tables, ..."""
    doc.stage = "analyze"
    return doc


@dataclass
class ConvertResult:
    pdf: str
    output: str
    pages: int = 0
    seconds: float = 0.0
    seconds_per_page: float = 0.0
    cached_stages: list[str] = field(default_factory=list)
    triage: dict[str, int] = field(default_factory=dict)
    fallback_pages: list[int] = field(default_factory=list)
    counts: dict[str, int] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)
    crosscheck: dict[str, Any] | None = None

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def convert(pdf: Path, out: Path, cfg: Config, *, workers: int | None = None,
            use_cache: bool = True) -> ConvertResult:
    """Convert one PDF. Raises only if the whole document cannot be opened/written."""
    t0 = time.perf_counter()
    pdf = Path(pdf)
    sha = sha256_file(pdf)
    work_dir = cfg.work_dir / sha
    work_dir.mkdir(parents=True, exist_ok=True)
    res = ConvertResult(pdf=str(pdf), output=str(out))
    n_workers = cfg.workers if workers is None else workers

    def _ingest() -> Document:
        doc = parse.parse_document(pdf, work_dir, sha256=sha, workers=n_workers)
        first = next((p for p in doc.pages if p.glyphs and p.rotation == 0
                      and not p.triage.fallback_image), None)
        if first is not None:
            try:
                xc = parse.crosscheck_pdfplumber(pdf, first)
                doc.metadata["pdf2pptx:crosscheck"] = json.dumps(xc, sort_keys=True)
            except Exception as exc:
                LOG.warning("%s: pdfplumber cross-check failed: %s", pdf.name, exc)
        return doc

    doc, hit = run_stage(work_dir, "ingest", _ingest, use_cache)
    if hit:
        res.cached_stages.append("ingest")
    doc, hit = run_stage(work_dir, "triage", lambda: triage.triage_document(doc), use_cache)
    if hit:
        res.cached_stages.append("triage")
    doc, hit = run_stage(work_dir, "analyze", lambda: analyze(doc), use_cache)
    if hit:
        res.cached_stages.append("analyze")

    rep = write_pptx(doc, out, work_dir, pdf)
    doc.stage = "build"
    _save_ir(work_dir, "build", doc)

    res.pages = len(doc.pages)
    res.seconds = round(time.perf_counter() - t0, 3)
    res.seconds_per_page = round(res.seconds / max(res.pages, 1), 3)
    res.triage = dict(sorted(Counter(p.triage.kind for p in doc.pages).items()))
    res.fallback_pages = sorted(p.index for p in doc.pages if p.triage.fallback_image)
    res.counts = dict(sorted(rep.counts.items()))
    res.errors = rep.errors
    xc = doc.metadata.get("pdf2pptx:crosscheck")
    res.crosscheck = json.loads(xc) if xc else None
    LOG.info("%s: %d pages in %.2fs -> %s (triage %s, fallback %s)", pdf.name, res.pages,
             res.seconds, out, res.triage, res.fallback_pages or "none")
    return res


# --------------------------------------------------------------------------- corpus


def _convert_entry(pdf: Path, out: Path, cfg: Config, use_cache: bool) -> dict[str, Any]:
    log.setup(False)
    try:
        return convert(pdf, out, cfg, workers=1, use_cache=use_cache).as_dict()
    except Exception as exc:  # a whole-document failure is recorded, not raised
        LOG.error("%s: conversion failed: %s", pdf.name, exc)
        return {"pdf": str(pdf), "output": str(out), "error": f"{type(exc).__name__}: {exc}"}


def convert_corpus(cfg: Config, *, only: list[str] | None = None, use_cache: bool = True,
                   check: bool = True) -> dict[str, Any]:
    """Convert every corpus PDF into ``out/corpus`` (process pool over files)."""
    from .corpus.build import load_manifest

    entries = [e for e in load_manifest(cfg.corpus_dir) if not only or e["name"] in only]
    out_dir = cfg.out_dir / "corpus"
    out_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.perf_counter()
    results: dict[str, dict[str, Any]] = {}
    with ProcessPoolExecutor(max_workers=max(1, min(cfg.workers, len(entries) or 1))) as pool:
        futs = {pool.submit(_convert_entry, cfg.corpus_dir / e["pdf"],
                            out_dir / f"{e['name']}.pptx", cfg, use_cache): e for e in entries}
        for fut in as_completed(futs):
            e = futs[fut]
            try:
                r = fut.result()
            except Exception as exc:  # worker died
                r = {"pdf": e["pdf"], "error": f"{type(exc).__name__}: {exc}"}
            r["name"] = e["name"]
            r["expected_pages"] = e["pages"]
            results[e["name"]] = r
    if check:
        _com_check(results)
    for name, r in results.items():
        (out_dir / f"{name}.convert.json").write_text(json.dumps(r, indent=1, sort_keys=True),
                                                      encoding="utf-8", newline="\n")
    ok = [r for r in results.values() if "error" not in r and r.get("pages") == r[
        "expected_pages"]]
    summary = {
        "files": len(results),
        "converted": len(ok),
        "opened_in_powerpoint": sum(1 for r in results.values() if r.get("com_open") == "ok"),
        "com_checked": check,
        "seconds": round(time.perf_counter() - t0, 2),
        "triage": dict(sorted(sum((Counter(r.get("triage", {})) for r in results.values()),
                                  Counter()).items())),
        "fallback_pages": {n: r["fallback_pages"] for n, r in sorted(results.items())
                           if r.get("fallback_pages")},
        "failures": {n: r.get("error") or r.get("com_open") for n, r in sorted(results.items())
                     if "error" in r or (check and r.get("com_open") != "ok")},
    }
    (out_dir / "_summary.json").write_text(json.dumps(summary, indent=1, sort_keys=True),
                                           encoding="utf-8", newline="\n")
    return summary


def _com_check(results: dict[str, dict[str, Any]]) -> None:
    """Open every output in PowerPoint; record whether it opened and the slide count."""
    from .verify import _com_temp

    try:
        with _com_temp.session() as ppt:
            for name in sorted(results):
                r = results[name]
                if "error" in r:
                    continue
                try:
                    n = ppt.slide_count(Path(r["output"]))
                    r["com_slides"] = n
                    r["com_open"] = "ok" if n == r["expected_pages"] else (
                        f"slide count {n} != {r['expected_pages']}")
                except Exception as exc:
                    r["com_open"] = f"open failed: {type(exc).__name__}: {exc}"
    except Exception as exc:
        LOG.warning("PowerPoint check skipped: %s", exc)
        for r in results.values():
            r.setdefault("com_open", f"unavailable: {exc}")
