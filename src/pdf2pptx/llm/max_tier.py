"""Max tier (M9): judgment + chart reading with the local vision model, then keep-best.

``run_many`` plans the jobs of every document first and runs them through one
:class:`Scheduler` (all jobs of a model across all documents, then the next model),
applies the validated answers, and keeps a page's changes only if the page scores at
least as well as its Lite version (composite, PowerPoint render; tolerance
``KEEP_TOL``). Without PowerPoint the changes cannot be measured and are dropped.
"""

from __future__ import annotations

import copy
import json
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .. import log
from ..config import Config
from ..ir import Document
from . import judge
from .scheduler import Scheduler
from .setup import candidate, candidates, installed, llama_server

LOG = log.get(__name__)

KEEP_TOL = 1e-4
# preferred judge / chart models when the benchmark has not chosen (best first)
DEFAULT_ORDER = ["Qwen3.6-35B-A3B", "Qwen3.5-9B", "Qwen3.5-4B"]


@dataclass
class Stack:
    judge: str | None
    chart: str | None
    text: str | None

    def as_dict(self) -> dict[str, str | None]:
        return {"judge": self.judge, "chart": self.chart, "text": self.text}


def _fits(name: str) -> bool:
    from .server import plan_memory

    try:
        c = candidate(name)
        return plan_memory(name, int(c.get("ctx", 8192))).fits
    except Exception:
        return False


def resolve_stack(cfg: Config, check_memory: bool = False) -> Stack:
    """Configured models (``llm_judge`` / ``llm_chart`` / ``llm_text``, written by
    ``pdf2pptx bench``) if installed, else the best installed default. With
    ``check_memory`` a model that does not fit the RAM free right now (available - 8 GB,
    measured peaks) is replaced by the next one that does."""
    def pick(want: str | None, kinds: tuple[str, ...]) -> str | None:
        order: list[str] = []
        if want:
            try:
                order.append(str(candidate(want)["name"]))
            except KeyError:
                LOG.warning("configured model %r is not a candidate", want)
        order += DEFAULT_ORDER + [c["name"] for c in candidates()]
        for n in order:
            c = candidate(n)
            if c["kind"] in kinds and installed(c) and (not check_memory or _fits(n)):
                if want and n != order[0]:
                    LOG.warning("%s does not fit in RAM now; using %s", order[0], n)
                return str(c["name"])
        return None

    return Stack(pick(cfg.llm_judge, ("vlm",)), pick(cfg.llm_chart or cfg.llm_judge, ("vlm",)),
                 pick(cfg.llm_text, ("text", "vlm")))


def available(cfg: Config) -> str | None:
    """None if Max can run, else the reason."""
    if llama_server() is None:
        return "llama.cpp missing (pdf2pptx setup llm)"
    if resolve_stack(cfg).judge is None:
        return "no vision model installed (pdf2pptx setup llm)"
    return None


def _score(pdf: Path, pptx: Path, cfg: Config, name: str) -> list[float | None]:
    from ..verify import score

    res = score.score_pair(pdf, pptx, cfg, name=name)
    if res.get("powerpoint_error"):
        raise RuntimeError(res["powerpoint_error"])
    return [s.get("composite") for s in res["slides"]]


def keep_best(doc: Document, lite: Document, pages: list[int], pdf: Path, work_dir: Path,
              cfg: Config) -> dict[str, Any]:
    """Score the affected pages with and without the judgment (flat, untuned decks);
    pages that got worse go back to their Lite elements."""
    from ..build.writer import TextOptions, write_pptx

    out: dict[str, Any] = {"pages": pages, "before": {}, "after": {}, "reverted": []}
    if not pages:
        return out
    opts = TextOptions.for_mode(cfg.mode, cfg.text_layout).flat()
    d = work_dir / "max"
    a, b = d / "lite.pptx", d / "judged.pptx"
    try:
        write_pptx(lite, a, work_dir, pdf, opts)
        write_pptx(doc, b, work_dir, pdf, opts)
        before = _score(pdf, a, cfg, f"{pdf.stem}__lite")
        after = _score(pdf, b, cfg, f"{pdf.stem}__judged")
    except Exception as exc:
        out["error"] = f"{type(exc).__name__}: {exc}"
        judge.drop_page_changes(doc, lite, pages)
        out["reverted"] = list(pages)
        LOG.warning("%s: keep-best unavailable (%s): judgment dropped", pdf.name, exc)
        return out
    for i in pages:
        x, y = before[i] if i < len(before) else None, after[i] if i < len(after) else None
        out["before"][i], out["after"][i] = x, y
        if x is not None and (y is None or y < x - KEEP_TOL):
            out["reverted"].append(i)
    if out["reverted"]:
        judge.drop_page_changes(doc, lite, out["reverted"])
    return out


def run_many(items: list[tuple[Document, Path, Path]], cfg: Config,
             scope: str | None = None) -> list[dict[str, Any]]:
    """Judge every (doc, pdf, work_dir) in place; one report per document."""
    t0 = time.perf_counter()
    reason = available(cfg)
    if reason:
        return [{"skipped": reason} for _ in items]
    stack = resolve_stack(cfg, check_memory=True)
    if stack.judge is None:
        why = "no installed vision model fits in the free RAM (available - 8 GB)"
        return [{"skipped": why} for _ in items]
    sched = Scheduler(cfg.work_dir / "llm_cache", log_dir=cfg.work_dir / "llm_logs")
    plans = []
    lites = []
    for doc, pdf, work_dir in items:
        lites.append(copy.deepcopy(doc))
        plans.append(judge.plan(doc, pdf, work_dir, sched, stack.judge, scope or cfg.judge,
                                stack.chart))
    t_plan = time.perf_counter() - t0
    runs = sched.run()
    sched.close()
    reports = []
    for (doc, pdf, work_dir), lite, rep in zip(items, lites, plans, strict=True):
        judge.apply(doc, rep)
        kb = keep_best(doc, lite, judge.affected_pages(rep), pdf, work_dir, cfg)
        r = rep.as_dict() | {"keep_best": kb, "stack": stack.as_dict()}
        doc.metadata["pdf2pptx:llm"] = json.dumps(
            {k: v for k, v in r.items() if k != "changes"}, sort_keys=True, default=str)
        reports.append(r)
    total = round(time.perf_counter() - t0, 2)
    run_info = [{"model": m.model, "jobs": m.jobs, "cached": m.cached, "failed": m.failed,
                 "seconds": m.seconds, "server": m.server, "memory": m.memory,
                 "error": m.error} for m in runs]
    for r in reports:
        r["models"] = run_info
        r["seconds_total_batch"] = total
        r["seconds_plan"] = round(t_plan, 2)
    return reports
