"""Model-call cache and scheduler.

* **Cache**: every call is keyed by (model id = candidate + weights sha256, prompt hash,
  image hashes, schema, limits) and stored as JSON under ``<work>/llm_cache``. A
  cached call never starts a server.
* **Scheduler**: jobs are queued per model; all jobs of model A run across the whole
  document (or corpus), then the server is swapped for model B. One model at a time,
  unless the machine has more than 48 GB of RAM (PLAN.md §8 rule 7): then a server
  stays resident until :meth:`Scheduler.close`.
"""

from __future__ import annotations

import hashlib
import json
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import psutil

from .. import log, progress
from ..ml import models
from .server import LlamaServer, LLMUnavailable
from .setup import candidate

LOG = log.get(__name__)

RESIDENT_ABOVE = 48 << 30


def model_id(name: str) -> str:
    c = candidate(name)
    m = models.spec()[c["model"]]
    return f"{c['name']}@{m.sha256[:12]}"


def call_key(name: str, prompt: str, images: list[bytes], schema: dict[str, Any] | None,
             system: str | None, max_tokens: int) -> str:
    h = hashlib.sha256()
    h.update(json.dumps({
        "model": model_id(name),
        "prompt": hashlib.sha256(prompt.encode("utf-8")).hexdigest(),
        "system": hashlib.sha256((system or "").encode("utf-8")).hexdigest(),
        "images": [hashlib.sha256(b).hexdigest() for b in images],
        "schema": schema, "max_tokens": max_tokens,
    }, sort_keys=True).encode("utf-8"))
    return h.hexdigest()


class Cache:
    def __init__(self, root: Path) -> None:
        self.root = root

    def _path(self, key: str) -> Path:
        return self.root / key[:2] / f"{key}.json"

    def get(self, key: str) -> dict[str, Any] | None:
        try:
            data: dict[str, Any] = json.loads(self._path(key).read_text(encoding="utf-8"))
            return data
        except (OSError, ValueError):
            return None

    def put(self, key: str, value: dict[str, Any]) -> None:
        p = self._path(key)
        p.parent.mkdir(parents=True, exist_ok=True)
        tmp = p.with_suffix(".tmp")
        tmp.write_text(json.dumps(value, sort_keys=True, ensure_ascii=False),
                       encoding="utf-8", newline="\n")
        tmp.replace(p)


@dataclass
class Job:
    model: str
    prompt: str
    images: list[bytes] = field(default_factory=list)
    schema: dict[str, Any] | None = None
    system: str | None = None
    max_tokens: int = 1024
    tag: str = ""  # caller's label (e.g. "judge:p3:roles")
    on_done: Callable[[dict[str, Any]], None] | None = None
    key: str = ""
    result: dict[str, Any] | None = None


@dataclass
class ModelRun:
    model: str
    jobs: int = 0
    cached: int = 0
    failed: int = 0
    seconds: float = 0.0
    server: dict[str, Any] = field(default_factory=dict)
    memory: dict[str, Any] = field(default_factory=dict)
    error: str | None = None


class Scheduler:
    def __init__(self, cache_dir: Path, *, log_dir: Path | None = None,
                 server_opts: dict[str, dict[str, Any]] | None = None,
                 resident: bool | None = None) -> None:
        self.cache = Cache(cache_dir)
        self.log_dir = log_dir
        self.server_opts = server_opts or {}
        self.resident = (psutil.virtual_memory().total > RESIDENT_ABOVE) if resident is None \
            else resident
        self.queue: list[Job] = []
        self.runs: list[ModelRun] = []
        self._servers: dict[str, LlamaServer] = {}

    def submit(self, job: Job) -> Job:
        job.key = call_key(job.model, job.prompt, job.images, job.schema, job.system,
                           job.max_tokens)
        self.queue.append(job)
        return job

    def _server(self, name: str) -> LlamaServer:
        srv = self._servers.get(name)
        if srv is None:
            srv = LlamaServer(name, **self.server_opts.get(name, {}))
            srv.start(log_dir=self.log_dir)
            if self.resident:
                self._servers[name] = srv
        return srv

    def run(self) -> list[ModelRun]:
        """Run every queued job, grouped by model in first-submission order."""
        order: list[str] = []
        for j in self.queue:
            if j.model not in order:
                order.append(j.model)
        done: list[ModelRun] = []
        for name in order:
            jobs = [j for j in self.queue if j.model == name]
            mr = ModelRun(model=name, jobs=len(jobs))
            todo = []
            for j in jobs:
                hit = self.cache.get(j.key)
                if hit is not None:
                    j.result = hit | {"cached": True}
                    mr.cached += 1
                else:
                    todo.append(j)
            t0 = time.perf_counter()
            if todo:
                srv: LlamaServer | None = None
                try:
                    srv = self._server(name)
                    mr.memory = srv.plan.as_dict() if srv.plan else {}
                    for k, j in enumerate(todo):
                        progress.report("model", k, len(todo), f"{name}: {j.tag}")
                        try:
                            r = srv.chat(j.prompt, images=j.images, schema=j.schema,
                                         system=j.system, max_tokens=j.max_tokens)
                            j.result = {"content": r["content"], "json": r["json"],
                                        "timings": r["timings"],
                                        "finish_reason": r["finish_reason"],
                                        "model": model_id(name)}
                            self.cache.put(j.key, j.result)
                        except Exception as exc:  # one failed call: the rule result stays
                            mr.failed += 1
                            j.result = {"error": f"{type(exc).__name__}: {exc}"}
                            LOG.warning("%s: %s failed: %s", name, j.tag, exc)
                except LLMUnavailable as exc:
                    mr.error = str(exc)
                    mr.failed += len(todo)
                    for j in todo:
                        j.result = {"error": str(exc)}
                    LOG.warning("%s unavailable: %s", name, exc)
                finally:
                    if srv is not None:
                        mr.server = srv.stats.as_dict() | {
                            "n_cpu_moe": srv.plan.n_cpu_moe if srv.plan else 0,
                            "mmproj_cpu_fallback": srv.mmproj_cpu_fallback}
                        if not self.resident:
                            srv.stop()
            mr.seconds = round(time.perf_counter() - t0, 2)
            for j in jobs:
                if j.on_done is not None and j.result is not None:
                    j.on_done(j.result)
            done.append(mr)
        self.queue.clear()
        self.runs.extend(done)
        return done

    def close(self) -> None:
        for srv in self._servers.values():
            srv.stop()
        self._servers.clear()
