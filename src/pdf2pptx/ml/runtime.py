"""One ONNX Runtime session factory for every model (PLAN.md §5, rule 3).

* Provider: DirectML (the iGPU) first, CPU as the fallback. ``PDF2PPTX_ML_DEVICE``
  = ``cpu`` forces the CPU, ``dml`` forbids the fallback (for tests), ``auto`` (the
  default) tries DirectML and falls back when the session cannot be created, its
  warm-up fails, or a later run raises.
* Per model the session is created once per process and warmed up on a small input.
* Every run is timed: :func:`stats` gives provider, load / warm-up time, calls, items
  and seconds per model (reported by ``convert`` and ``doctor``).
* Missing onnxruntime or model file: :func:`session` returns None and the caller
  falls back to rules (never a crash).
"""

from __future__ import annotations

import os
import threading
import time
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import numpy as np

from .. import log
from . import models

LOG = log.get(__name__)

DEVICE_ENV = "PDF2PPTX_ML_DEVICE"
DML = "DmlExecutionProvider"
CPU = "CPUExecutionProvider"
# Models that stay on the CPU in ``auto`` (measured on the Radeon 780M, M8):
# * inpaint (LaMa): a MatMul of its Fourier unit fails on DirectML ("parameter not
#   correct") after a ~20 s graph compile; CPU: 2.2 s per 512 tile.
# * ocr_cls: 0.02 s per page on the CPU vs 0.2 s on DirectML (tiny model, launch cost).
PREFER_CPU = {"inpaint": "LaMa fails on DirectML", "ocr_cls": "faster on the CPU"}


def device_pref() -> str:
    v = os.environ.get(DEVICE_ENV, "auto").strip().lower()
    return v if v in ("auto", "cpu", "dml") else "auto"


def ort_module() -> Any | None:
    try:
        import onnxruntime as ort  # type: ignore[import-not-found]
    except Exception:  # not installed, or a broken DLL
        return None
    return ort


@dataclass
class ModelStats:
    name: str
    provider: str = ""
    load_s: float = 0.0
    warmup_s: float = 0.0
    calls: int = 0
    items: int = 0
    seconds: float = 0.0
    fallback: str = ""  # why DirectML was not used

    def as_dict(self) -> dict[str, Any]:
        d = asdict(self)
        for k in ("load_s", "warmup_s", "seconds"):
            d[k] = round(d[k], 3)
        return d


@dataclass
class Session:
    """An ONNX session plus its timing; ``run`` retries once on the CPU if DirectML
    fails at run time."""

    name: str
    model_path: Path
    sess: Any
    stats: ModelStats
    warm: Callable[[Session], None] | None = None
    lock: threading.Lock = field(default_factory=threading.Lock)

    @property
    def provider(self) -> str:
        return self.stats.provider

    @property
    def input_names(self) -> list[str]:
        return [i.name for i in self.sess.get_inputs()]

    def run(self, feeds: dict[str, np.ndarray], items: int = 1) -> list[np.ndarray]:
        with self.lock:
            t0 = time.perf_counter()
            try:
                out = self.sess.run(None, feeds)
            except Exception as exc:
                if self.stats.provider != DML or device_pref() == "dml":
                    raise
                LOG.warning("%s: DirectML run failed (%s); switching to CPU", self.name, exc)
                self.sess = _create(self.model_path, [CPU])
                self.stats.provider = CPU
                self.stats.fallback = f"run failed: {type(exc).__name__}"
                out = self.sess.run(None, feeds)
            self.stats.calls += 1
            self.stats.items += items
            self.stats.seconds += time.perf_counter() - t0
            return list(out)

    def metadata(self) -> dict[str, str]:
        return dict(self.sess.get_modelmeta().custom_metadata_map)


_SESSIONS: dict[str, Session | None] = {}
_LOCK = threading.Lock()


def _create(model_path: Path, providers: list[str]) -> Any:
    ort = ort_module()
    assert ort is not None
    so = ort.SessionOptions()
    so.log_severity_level = 3
    if DML in providers:  # required by the DirectML provider
        so.enable_mem_pattern = False
        so.execution_mode = ort.ExecutionMode.ORT_SEQUENTIAL
    else:
        so.intra_op_num_threads = max(1, min(8, (os.cpu_count() or 4) // 2))
    return ort.InferenceSession(str(model_path), sess_options=so, providers=providers)


def session(name: str, warmup: Callable[[Session], None] | None = None) -> Session | None:
    """The (cached) session of model ``name`` from the manifest, or None if it cannot
    be used. ``warmup`` runs once on creation; a DirectML warm-up failure moves the
    model to the CPU."""
    with _LOCK:
        if name in _SESSIONS:
            return _SESSIONS[name]
        _SESSIONS[name] = s = _open(name, warmup)
        return s


def _open(name: str, warmup: Callable[[Session], None] | None) -> Session | None:
    ort = ort_module()
    mp = models.path(name)
    if ort is None or mp is None:
        LOG.info("model %s unavailable (%s)", name,
                 "onnxruntime missing" if ort is None else "file not installed")
        return None
    pref = device_pref()
    tried: list[list[str]] = []
    has_dml = DML in ort.get_available_providers()
    if (pref == "dml" or (pref == "auto" and has_dml and name not in PREFER_CPU)) and has_dml:
        tried.append([DML, CPU])
    if pref != "dml" or not tried:
        tried.append([CPU])
    if tried[0][0] == DML:
        reason = ""
    elif pref == "cpu":
        reason = "forced by " + DEVICE_ENV
    elif has_dml and name in PREFER_CPU:
        reason = PREFER_CPU[name]
    else:
        reason = "DirectML provider not available"
    for providers in tried:
        st = ModelStats(name=name, fallback=reason)
        t0 = time.perf_counter()
        try:
            sess = _create(mp, providers)
        except Exception as exc:
            reason = f"session failed: {type(exc).__name__}"
            LOG.warning("%s: %s session failed (%s)", name, providers[0], exc)
            continue
        st.load_s = time.perf_counter() - t0
        st.provider = sess.get_providers()[0]
        s = Session(name=name, model_path=mp, sess=sess, stats=st, warm=warmup)
        if warmup is not None:
            t1 = time.perf_counter()
            try:
                warmup(s)
            except Exception as exc:
                reason = f"warm-up failed: {type(exc).__name__}: {exc}"
                LOG.warning("%s: warm-up on %s failed (%s)", name, st.provider, exc)
                if providers[0] == DML and pref != "dml":
                    continue
                return None
            st.warmup_s = time.perf_counter() - t1
            st.calls = st.items = 0
            st.seconds = 0.0
        if st.provider != DML and not st.fallback:
            st.fallback = reason
        LOG.debug("%s on %s (load %.2fs, warm-up %.2fs)", name, st.provider, st.load_s,
                  st.warmup_s)
        return s
    return None


def stats() -> dict[str, dict[str, Any]]:
    """Per loaded model: provider and timings, sorted by name."""
    with _LOCK:
        return {n: s.stats.as_dict() for n, s in sorted(_SESSIONS.items()) if s is not None}


def stats_since(before: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Per model: what ran after the ``before`` snapshot of :func:`stats` (sessions
    live for the whole process, a corpus worker converts several documents)."""
    out = {}
    for n, st in stats().items():
        b = before.get(n, {})
        calls = st["calls"] - b.get("calls", 0)
        if calls <= 0:
            continue
        out[n] = {**st, "calls": calls, "items": st["items"] - b.get("items", 0),
                  "seconds": round(st["seconds"] - b.get("seconds", 0.0), 3),
                  "load_s": st["load_s"] if n not in before else 0.0,
                  "warmup_s": st["warmup_s"] if n not in before else 0.0}
    return out


def reset() -> None:
    """Drop every cached session (tests; switching ``PDF2PPTX_ML_DEVICE``)."""
    with _LOCK:
        _SESSIONS.clear()
