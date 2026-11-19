"""PDF -> PPTX pipeline: ingest -> triage -> ocr -> analyze -> build.

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

from . import log, progress
from .build.writer import TextOptions, WriteReport, write_pptx
from .config import Config
from .ingest import parse, triage
from .ir import Document

LOG = log.get(__name__)

STAGE_VERSIONS = {
    "ingest": parse.PARSER_VERSION,
    "triage": "2.0",
    "ocr": "8.1",  # + the OCR engine / models installed (stage_versions)
    "analyze": "charts-9.0",
}
STAGES = tuple(STAGE_VERSIONS)


def stage_versions() -> dict[str, str]:
    """Stage versions plus, for ``ocr``, which engine and models this machine has:
    installing the models invalidates the OCR stage and everything after it."""
    from .ml import ocr

    v = dict(STAGE_VERSIONS)
    v["ocr"] = f"{v['ocr']}+{ocr.engine_id()}"
    return v


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def stage_key(stage: str) -> str:
    keys = []
    versions = stage_versions()
    for s in STAGES:
        keys.append(f"{s}={versions[s]}")
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


def analyze(doc: Document, work_dir: Path | None = None, pdf: Path | None = None) -> Document:
    """M4: fonts, text blocks, lists, roles, reading order. M6: tables (before text, so
    cell glyphs are grouped per cell) and graphics (presets, connectors, groups, SVG).
    M8: the layout model checks low-confidence roles (born-digital pages) and OCR
    pages' roles against their stored regions.
    M7: document structure: links / notes / sections / producer, build-up sequences,
    theme, then masters / layouts / placeholders (each optional: a failure leaves the
    flat deck)."""
    from .analyze import charts, shapes, tables
    from .analyze.text import analyze_text
    from .fonts.match import match_document

    matches = match_document(doc)
    skip: dict[int, set[str]] = {}
    stats: Counter[str] = Counter()
    for page in doc.pages:
        if page.triage.fallback_image:
            continue
        try:
            skip[page.index] = tables.analyze_page(page, doc.fonts, matches)
        except Exception as exc:  # tables are optional: the page keeps text + shapes
            LOG.error("page %d: table detection failed (%s: %s)", page.index,
                      type(exc).__name__, exc, exc_info=LOG.isEnabledFor(10))
    doc = analyze_text(doc, matches=matches, skip_glyphs=skip)
    wd = work_dir or Path(".")
    if pdf is not None:
        try:
            from .ml import layout

            layout.refine_roles(doc, pdf, wd)
        except Exception as exc:  # the model is optional: rule roles stay
            LOG.error("layout refinement failed (%s: %s)", type(exc).__name__, exc,
                      exc_info=LOG.isEnabledFor(10))
    for page in doc.pages:
        progress.report("analyze", page.index, len(doc.pages))
        try:
            stats.update(shapes.analyze_page(page, wd))
        except Exception as exc:  # graphics are optional: freeforms stay
            LOG.error("page %d: shape analysis failed (%s: %s)", page.index,
                      type(exc).__name__, exc, exc_info=LOG.isEnabledFor(10))
        try:  # M9: vector charts -> native charts (their parts kept as the fallback)
            stats.update(charts.analyze_page(page))
        except Exception as exc:  # charts are optional: the shapes stay
            LOG.error("page %d: chart analysis failed (%s: %s)", page.index,
                      type(exc).__name__, exc, exc_info=LOG.isEnabledFor(10))
    doc.metadata["pdf2pptx:graphics"] = json.dumps(dict(sorted(stats.items())), sort_keys=True)
    doc.metadata["pdf2pptx:structure"] = json.dumps(analyze_structure(doc), sort_keys=True)
    doc.stage = "analyze"
    return doc


def analyze_structure(doc: Document) -> dict[str, int]:
    """M7 passes in order; each one is optional (logged, never raised)."""
    from .analyze import builds, links, masters, theme

    stats: dict[str, int] = {}
    for name, fn in (("links", links.analyze_links), ("builds", builds.analyze_builds),
                     ("theme", theme.analyze_theme), ("masters", masters.analyze_masters)):
        try:
            stats.update(fn(doc))
        except Exception as exc:
            LOG.error("%s analysis failed (%s: %s); skipped", name, type(exc).__name__, exc,
                      exc_info=LOG.isEnabledFor(10))
            if name == "builds":
                for p in doc.pages:
                    p.build_of = None
            if name == "masters":
                doc.masters = []
                for p in doc.pages:
                    p.layout_id = None
                    for e in p.elements:
                        e.notes = [n for n in e.notes
                                   if not n.startswith((masters.MASTER_NOTE, masters.LAYOUT_NOTE))]
                        if hasattr(e, "placeholder"):
                            e.placeholder = None
    return dict(sorted(stats.items()))


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
    tune: dict[str, Any] | None = None
    slides: int = 0  # slides written (fewer than pages when build-ups are collapsed)
    structure_check: dict[str, Any] | None = None  # fidelity: structured vs flat render
    ml: dict[str, Any] | None = None  # M8: OCR / layout / inpainting: what ran, providers
    charts: dict[str, Any] | None = None  # M9: native charts and their render check
    llm: dict[str, Any] | None = None  # M9: max tier: judgment / chart model calls

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def analyzed(pdf: Path, cfg: Config, *, workers: int | None = None,
             use_cache: bool = True) -> tuple[Document, Path, list[str]]:
    """ingest -> triage -> analyze (each cached): (IR, work dir, cached stage names)."""
    pdf = Path(pdf)
    parse.open_pdf(pdf).close()  # InputError: missing / encrypted / damaged / not a PDF
    sha = sha256_file(pdf)
    work_dir = cfg.work_dir / sha
    work_dir.mkdir(parents=True, exist_ok=True)
    n_workers = cfg.workers if workers is None else workers
    cached: list[str] = []

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
        cached.append("ingest")
    doc, hit = run_stage(work_dir, "triage", lambda: triage.triage_document(doc), use_cache)
    if hit:
        cached.append("triage")
    doc, hit = run_stage(work_dir, "ocr", lambda: _ocr(doc, pdf, work_dir), use_cache)
    if hit:
        cached.append("ocr")
    doc, hit = run_stage(work_dir, "analyze", lambda: analyze(doc, work_dir, pdf), use_cache)
    if hit:
        cached.append("analyze")
    return doc, work_dir, cached


def _ocr(doc: Document, pdf: Path, work_dir: Path) -> Document:
    from .ml import ocr

    return ocr.run(doc, pdf, work_dir)


def ml_summary(doc: Document) -> dict[str, Any] | None:
    """What the models did for this document (compact; the full record is in the IR
    metadata ``pdf2pptx:ocr`` / ``pdf2pptx:layout``)."""
    def load(key: str) -> dict[str, Any]:
        try:
            return dict(json.loads(doc.metadata.get(key) or "{}"))
        except ValueError:
            return {}

    ocr, lay = load("pdf2pptx:ocr"), load("pdf2pptx:layout")
    if not ocr and not lay:
        return None
    pages = {}
    for k, v in sorted(ocr.get("pages", {}).items(), key=lambda kv: int(kv[0])):
        keep = ("kind", "engine", "skew", "words", "mean_conf", "lines_remapped",
                "lines_replaced", "replaced", "kept_in_picture", "family", "error", "gaps")
        pages[k] = {x: v[x] for x in keep if x in v}
        if "inpaint" in v:
            pages[k]["inpaint"] = {x: v["inpaint"][x] for x in ("uniform", "lama", "telea")}
        pages[k]["seconds"] = v.get("seconds", {}).get("total")
    return {
        "ocr_pages": pages,
        "ocr_flagged": ocr.get("flagged", 0),
        "models": {n: {"provider": m.get("provider"), "seconds": m.get("seconds"),
                       "calls": m.get("calls")}
                   for n, m in sorted((ocr.get("models", {}) | lay.get("models", {})).items())},
        "layout": {k: lay.get(k) for k in ("pages_checked", "model_runs", "blocks_low",
                                           "trigger_rate") if k in lay}
        | {"disagreements": len(lay.get("disagreements", [])),
           "adopted": sum(1 for d in lay.get("disagreements", []) if d.get("adopted"))},
    }


def write_options(cfg: Config) -> TextOptions:
    return TextOptions.for_mode(cfg.mode, cfg.text_layout, builds=cfg.builds_enabled,
                                structure=cfg.structure)


def _tune(doc: Document, out: Path, work_dir: Path, pdf: Path, cfg: Config) -> dict[str, Any]:
    """Tune on the flat deck (one slide per page, M5/M6): ``_finalize`` writes the
    structured deck from the tuned IR afterwards."""
    from .verify.tuner import tune_document

    flat = write_options(cfg).flat()
    write_pptx(doc, out, work_dir, pdf, flat)
    rep = tune_document(doc, out, work_dir, pdf, flat, cfg.limits, rounds=cfg.tune_rounds,
                        render_dpi=cfg.render_dpi)
    if rep.skipped:
        LOG.info("%s: tuner skipped (%s)", pdf.name, rep.skipped)
    else:
        LOG.info("%s: tuned %d/%d blocks in %d rounds, %.2fs (breaks %d -> %d)", pdf.name,
                 rep.blocks_changed, rep.blocks, rep.rounds, rep.seconds, rep.breaks_before,
                 rep.breaks_after)
    return rep.as_dict()


def _finalize(doc: Document, out: Path, work_dir: Path, pdf: Path, cfg: Config, *,
              check: bool) -> tuple[WriteReport, dict[str, Any] | None]:
    """The final write (M7 structure per the mode). In fidelity mode the master /
    layouts are kept only if every slide renders like the flat deck (``check``)."""
    from .analyze.masters import refresh_copies

    opts = write_options(cfg)
    refresh_copies(doc)
    rep = write_pptx(doc, out, work_dir, pdf, opts)
    if not (check and cfg.mode == "fidelity" and opts.structure and doc.masters):
        return rep, None
    from .verify.deck import render_check

    flat_out = work_dir / "structure_check" / "flat.pptx"
    write_pptx(doc, flat_out, work_dir, pdf, opts.flat())
    verdict = render_check(out, flat_out, cfg)
    if not verdict.get("ok"):
        LOG.info("%s: structured deck differs from the flat one (%s); keeping the flat deck",
                 pdf.name, verdict.get("reason") or verdict.get("min_ssim"))
        rep = write_pptx(doc, out, work_dir, pdf, opts.flat())
        verdict["kept"] = "flat"
    else:
        verdict["kept"] = "structured"
    return rep, verdict


def _check_charts(doc: Document, out: Path, work_dir: Path, pdf: Path, cfg: Config,
                  rep: WriteReport) -> tuple[WriteReport, dict[str, Any] | None]:
    """M9: render-check the native charts; rewrite the deck if one went back to shapes."""
    from .verify import chart_check

    if not chart_check.charts_of(doc):
        return rep, None
    res = chart_check.check(doc, out, work_dir, pdf, dpi=cfg.render_dpi,
                            builds=write_options(cfg).builds)
    if res["reverted"]:
        rep, _ = _finalize(doc, out, work_dir, pdf, cfg, check=False)
    return rep, res


def _max(doc: Document, pdf: Path, work_dir: Path, cfg: Config) -> dict[str, Any]:
    """M9 max tier on one document: judgment + chart reading + keep-best."""
    from .llm import max_tier

    try:
        return max_tier.run_many([(doc, pdf, work_dir)], cfg)[0]
    except Exception as exc:  # the model layer is optional: the Lite IR stays
        LOG.error("%s: max tier failed (%s: %s); Lite result kept", pdf.name,
                  type(exc).__name__, exc, exc_info=LOG.isEnabledFor(10))
        return {"error": f"{type(exc).__name__}: {exc}"}


def convert(pdf: Path, out: Path, cfg: Config, *, workers: int | None = None,
            use_cache: bool = True, tune: bool | None = None,
            check_structure: bool = True, max_tier: bool = True) -> ConvertResult:
    """Convert one PDF. Raises only if the whole document cannot be opened/written.

    ``tune``: run the fidelity tuner (M5) after writing; None = the mode's default.
    ``check_structure``: in fidelity mode, render-check the master / layouts (M7).
    """
    t0 = time.perf_counter()
    pdf = Path(pdf)
    res = ConvertResult(pdf=str(pdf), output=str(out))
    doc, work_dir, res.cached_stages = analyzed(pdf, cfg, workers=workers, use_cache=use_cache)
    if max_tier and cfg.resolved_tier() == "max" and cfg.judge != "off":
        res.llm = _max(doc, pdf, work_dir, cfg)

    if (cfg.tune_enabled if tune is None else tune) and cfg.text_layout == "blocks":
        res.tune = _tune(doc, out, work_dir, pdf, cfg)
    rep, res.structure_check = _finalize(doc, out, work_dir, pdf, cfg, check=check_structure)
    if check_structure:
        rep, res.charts = _check_charts(doc, out, work_dir, pdf, cfg, rep)
    doc.stage = "build"
    _save_ir(work_dir, "build", doc)

    res.pages = len(doc.pages)
    res.slides = rep.slides
    res.seconds = round(time.perf_counter() - t0, 3)
    res.seconds_per_page = round(res.seconds / max(res.pages, 1), 3)
    res.triage = dict(sorted(Counter(p.triage.kind for p in doc.pages).items()))
    res.fallback_pages = sorted(p.index for p in doc.pages if p.triage.fallback_image)
    res.counts = dict(sorted(rep.counts.items()))
    res.errors = rep.errors
    xc = doc.metadata.get("pdf2pptx:crosscheck")
    res.crosscheck = json.loads(xc) if xc else None
    res.ml = ml_summary(doc)
    LOG.info("%s: %d pages in %.2fs -> %s (triage %s, fallback %s)", pdf.name, res.pages,
             res.seconds, out, res.triage, res.fallback_pages or "none")
    return res


# --------------------------------------------------------------------------- corpus


def _convert_entry(pdf: Path, out: Path, cfg: Config, use_cache: bool) -> dict[str, Any]:
    log.setup(False)
    try:
        return convert(pdf, out, cfg, workers=1, use_cache=use_cache, tune=False,
                       check_structure=False, max_tier=False).as_dict()
    except Exception as exc:  # a whole-document failure is recorded, not raised
        LOG.error("%s: conversion failed: %s", pdf.name, exc)
        return {"pdf": str(pdf), "output": str(out), "error": f"{type(exc).__name__}: {exc}"}


def convert_corpus(cfg: Config, *, only: list[str] | None = None, use_cache: bool = True,
                   check: bool = True) -> dict[str, Any]:
    """Convert every corpus PDF into ``out/corpus`` (process pool over files)."""
    from .corpus.build import load_manifest

    entries = [e for e in load_manifest(cfg.corpus_dir) if not only or e["name"] in only]
    try:  # index installed fonts once here, not in every worker
        from .fonts.index import face_index

        face_index()
    except Exception as exc:
        LOG.warning("font index unavailable: %s", exc)
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
    if cfg.resolved_tier() == "max" and cfg.judge != "off":
        _max_corpus(results, cfg)
    if (cfg.tune_enabled and cfg.text_layout == "blocks") or cfg.mode == "fidelity":
        _tune_corpus(results, cfg, tune=cfg.tune_enabled and cfg.text_layout == "blocks")
    if check:
        _chart_check_corpus(results, cfg)
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


def _max_corpus(results: dict[str, dict[str, Any]], cfg: Config) -> None:
    """M9: judge every converted file with one scheduler (each model loads once for the
    whole corpus), then rewrite the decks."""
    from .llm import max_tier

    items, names = [], []
    for name in sorted(results):
        r = results[name]
        if "error" in r:
            continue
        pdf = Path(r["pdf"])
        work_dir = cfg.work_dir / sha256_file(pdf)
        doc = Document.from_json((work_dir / "ir_build.json").read_text(encoding="utf-8"))
        items.append((doc, pdf, work_dir))
        names.append(name)
    t0 = time.perf_counter()
    try:
        reps = max_tier.run_many(items, cfg)
    except Exception as exc:
        LOG.error("max tier failed (%s: %s); Lite results kept", type(exc).__name__, exc)
        for n in names:
            results[n]["llm"] = {"error": f"{type(exc).__name__}: {exc}"}
        return
    per = (time.perf_counter() - t0) / max(len(items), 1)
    for name, (doc, pdf, work_dir), rep in zip(names, items, reps, strict=True):
        r = results[name]
        r["llm"] = rep
        if rep.get("applied"):
            wr, _ = _finalize(doc, Path(r["output"]), work_dir, pdf, cfg, check=False)
            r["counts"] = dict(sorted(wr.counts.items()))
            r["slides"] = wr.slides
        _save_ir(work_dir, "build", doc)
        r["seconds"] = round(r["seconds"] + per, 3)
        r["seconds_per_page"] = round(r["seconds"] / max(r.get("pages") or 1, 1), 3)


def _tune_corpus(results: dict[str, dict[str, Any]], cfg: Config, tune: bool = True) -> None:
    """Tune every converted file here, one after the other, on one PowerPoint engine,
    then write the final (structured) deck: ``_finalize``."""
    for name in sorted(results):
        r = results[name]
        if "error" in r:
            continue
        try:
            pdf, out = Path(r["pdf"]), Path(r["output"])
            work_dir = cfg.work_dir / sha256_file(pdf)
            doc = Document.from_json((work_dir / "ir_build.json").read_text(encoding="utf-8"))
            t0 = time.perf_counter()
            if tune:
                r["tune"] = _tune(doc, out, work_dir, pdf, cfg)
            rep, r["structure_check"] = _finalize(doc, out, work_dir, pdf, cfg, check=True)
            r["counts"] = dict(sorted(rep.counts.items()))
            r["slides"] = rep.slides
            _save_ir(work_dir, "build", doc)
            r["seconds"] = round(r["seconds"] + time.perf_counter() - t0, 3)
            r["seconds_per_page"] = round(r["seconds"] / max(r.get("pages") or 1, 1), 3)
        except Exception as exc:  # the untuned deck stays
            LOG.error("%s: tuning failed: %s", name, exc)
            r["tune"] = {"skipped": f"{type(exc).__name__}: {exc}"}


def _chart_check_corpus(results: dict[str, dict[str, Any]], cfg: Config) -> None:
    """M9: render-check native charts of every converted file (one PowerPoint)."""
    for name in sorted(results):
        r = results[name]
        if "error" in r or not r.get("counts", {}).get("charts"):
            continue
        try:
            pdf, out = Path(r["pdf"]), Path(r["output"])
            work_dir = cfg.work_dir / sha256_file(pdf)
            doc = Document.from_json((work_dir / "ir_build.json").read_text(encoding="utf-8"))
            rep, r["charts"] = _check_charts(doc, out, work_dir, pdf, cfg, WriteReport())
            if r["charts"] and r["charts"]["reverted"]:
                r["counts"] = dict(sorted(rep.counts.items()))
            _save_ir(work_dir, "build", doc)
        except Exception as exc:  # the unchecked deck stays
            LOG.error("%s: chart check failed: %s", name, exc)
            r["charts"] = {"error": f"{type(exc).__name__}: {exc}"}


def _com_check(results: dict[str, dict[str, Any]]) -> None:
    """Open every output in PowerPoint; record whether it opened and the slide count."""
    from .verify import ppt_com

    try:
        ppt = ppt_com.shared()
        for name in sorted(results):
            r = results[name]
            if "error" in r:
                continue
            try:
                n = ppt.slide_count(Path(r["output"]))
                r["com_slides"] = n
                want = r.get("slides") or r["expected_pages"]  # builds: fewer slides
                r["com_open"] = "ok" if n == want else f"slide count {n} != {want}"
            except ppt_com.PowerPointUnavailable:
                raise
            except Exception as exc:
                r["com_open"] = f"open failed: {type(exc).__name__}: {exc}"
    except Exception as exc:
        LOG.warning("PowerPoint check skipped: %s", exc)
        for r in results.values():
            r.setdefault("com_open", f"unavailable: {exc}")
