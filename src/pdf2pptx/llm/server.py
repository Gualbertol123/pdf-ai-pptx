"""llama-server (llama.cpp, Vulkan) management: one model at a time.

* Start / health / stop of ``tools/llama.cpp/llama-server.exe`` on 127.0.0.1 with
  ``-ngl``, ``--n-cpu-moe``, flash attention, a q8_0 KV cache and a context size.
* **Memory guard** (PLAN.md §8 rule 7). Before a start, the committed memory is
  estimated from the GGUF tensor table: the weights that go to the GPU (on this iGPU
  that is system RAM), the projector, the KV cache and a compute buffer. Expert
  weights kept on the CPU with ``--n-cpu-moe`` stay memory-mapped from the file (the
  OS can page them out), so they are not counted. If the estimate exceeds the
  available RAM minus 8 GB the start is refused with a suggested ``--n-cpu-moe``
  value, or, with ``auto_moe``, that value is used.
* While a server runs a sampler records the lowest available RAM: ``peak_used`` =
  available before the start - lowest available (includes the GPU's shared memory,
  which a per-process RSS would miss).
* Requests: ``/v1/chat/completions`` with images as data URLs and output forced by a
  JSON schema (``response_format``). Thinking is switched off. The response's
  ``timings`` give prompt and generation speed.
"""

from __future__ import annotations

import base64
import functools
import json
import os
import re
import socket
import subprocess
import threading
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import psutil

from .. import log
from ..ml import models
from . import gguf
from .setup import candidate, llama_server

LOG = log.get(__name__)

RESERVE_BYTES = 8 << 30  # leave this much RAM free (the prompt's rule)
COMPUTE_BUFFER = 768 << 20  # llama.cpp compute buffers, measured order of magnitude
HOST = "127.0.0.1"
MAX_IMAGE_PX = 768  # longest image side sent to a vision model (Vulkan encoder limit)


class LLMUnavailable(RuntimeError):
    """llama.cpp or the model is missing, or it does not fit in memory."""


@dataclass
class MemoryPlan:
    model: str
    available: int
    budget: int  # available - reserve
    need: int  # RAM the running model takes (all weights + KV + buffers)
    n_cpu_moe: int
    expert_cpu_bytes: int
    gpu_need: int  # what goes to the Vulkan device (system RAM on this iGPU)
    gpu_budget: int
    fits: bool
    suggestion: str = ""

    def as_dict(self) -> dict[str, Any]:
        gb = 2**30
        return {"model": self.model, "available_gb": round(self.available / gb, 2),
                "budget_gb": round(self.budget / gb, 2), "need_gb": round(self.need / gb, 2),
                "gpu_need_gb": round(self.gpu_need / gb, 2),
                "gpu_budget_gb": round(self.gpu_budget / gb, 2),
                "n_cpu_moe": self.n_cpu_moe,
                "experts_on_cpu_gb": round(self.expert_cpu_bytes / gb, 2), "fits": self.fits,
                "suggestion": self.suggestion}


def reserve_bytes() -> int:
    v = os.environ.get("PDF2PPTX_LLM_RESERVE_GB")
    try:
        return int(float(v) * 2**30) if v else RESERVE_BYTES
    except ValueError:
        return RESERVE_BYTES


@functools.cache
def vulkan_free_bytes() -> int:
    """Free memory of the first Vulkan device (``llama-server --list-devices``)."""
    exe = llama_server()
    if exe is not None:
        try:
            out = subprocess.run([str(exe), "--list-devices"], capture_output=True, text=True,
                                 timeout=60, creationflags=getattr(subprocess,
                                                                   "CREATE_NO_WINDOW", 0))
            m = re.search(r"Vulkan0:.*?\((\d+) MiB, (\d+) MiB free\)", out.stdout + out.stderr)
            if m:
                return int(m.group(2)) << 20
        except (OSError, subprocess.SubprocessError):
            pass
    return 14 << 30


def _measured_path() -> Path:
    return models.models_dir() / "llm" / "measured_peak.json"


def measured_peak(name: str) -> int:
    """The highest RAM use measured for this model (bytes, 0 if never measured)."""
    try:
        data = json.loads(_measured_path().read_text(encoding="utf-8"))
        return int(data.get(name, 0))
    except (OSError, ValueError):
        return 0


def record_peak(name: str, peak: int) -> None:
    """Keep the highest measured peak per model (the memory guard uses it)."""
    p = _measured_path()
    try:
        data = json.loads(p.read_text(encoding="utf-8")) if p.is_file() else {}
    except (OSError, ValueError):
        data = {}
    if peak > int(data.get(name, 0)):
        data[name] = int(peak)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(data, indent=1, sort_keys=True) + "\n", encoding="utf-8",
                     newline="\n")


def plan_memory(name: str, ctx: int, n_cpu_moe: int | None = None,
                available: int | None = None, gpu_free: int | None = None) -> MemoryPlan:
    """Memory plan for candidate ``name``.

    Measured (M9, Qwen3.6-35B-A3B with 18 layers of experts on the CPU): the CPU-side
    experts are memory-mapped but become resident as they are used, so the RAM need is
    the whole model + projector + KV cache + compute buffers whatever ``--n-cpu-moe``
    is. ``--n-cpu-moe`` only moves weights off the Vulkan device: it is chosen as the
    smallest value whose GPU part fits the device's free memory (90 %)."""
    c = candidate(name)
    mp = models.path(c["model"])
    if mp is None:
        raise LLMUnavailable(f"{c['name']}: model file missing (pdf2pptx setup llm)")
    info = gguf.read(mp)
    proj = models.path(c["mmproj"]) if c.get("mmproj") else None
    proj_b = proj.stat().st_size if proj else 0
    avail = psutil.virtual_memory().available if available is None else available
    budget = avail - reserve_bytes()
    gpu_budget = int(0.9 * (vulkan_free_bytes() if gpu_free is None else gpu_free))
    per_layer = info.expert_bytes_per_layer()
    extra = info.kv_bytes(ctx) + COMPUTE_BUFFER
    # measured (M9): on the 780M the Vulkan buffers come out of system RAM and the peak
    # is well above the file sizes (Qwen3.5-9B 9.4 GB for 7.2 GB of files + KV), so a
    # measured peak, when there is one, wins
    need = max(info.file_size + proj_b + extra, measured_peak(str(c["name"])))

    def gpu_for(k: int) -> tuple[int, int]:
        cpu = sum(v for layer, v in per_layer.items() if layer < k)
        return info.file_size - cpu + extra, cpu

    if n_cpu_moe is None:
        k = 0
        gpu, cpu = gpu_for(0)
        while gpu > gpu_budget and per_layer and k < info.n_layers:
            k += 1
            gpu, cpu = gpu_for(k)
    else:
        k = n_cpu_moe
        gpu, cpu = gpu_for(k)
    ram_ok, gpu_ok = need <= budget, gpu <= gpu_budget
    sug = ""
    if not ram_ok:
        short = (need - budget) / 2**30
        sug = (f"{short:.1f} GB short of RAM (needs {need / 2**30:.1f} GB + "
               f"{reserve_bytes() / 2**30:.0f} GB reserve): close programs, or a smaller "
               "model / quant")
    elif not gpu_ok:
        sug = "the GPU part does not fit the Vulkan device even with every expert on the CPU"
    elif k and n_cpu_moe is None:
        sug = f"--n-cpu-moe {k}"
    return MemoryPlan(c["name"], avail, budget, need, k, cpu, gpu, gpu_budget,
                      ram_ok and gpu_ok, sug)


def _free_port() -> int:
    with socket.socket() as s:
        s.bind((HOST, 0))
        return int(s.getsockname()[1])


@dataclass
class ServerStats:
    load_seconds: float = 0.0
    available_before: int = 0
    lowest_available: int = 0
    requests: int = 0
    prompt_tokens: int = 0
    prompt_ms: float = 0.0
    gen_tokens: int = 0
    gen_ms: float = 0.0
    peak_private: int = 0  # llama-server's highest committed memory (incl. GPU buffers)

    @property
    def peak_used(self) -> int:
        """The server's own peak: its committed memory (on this iGPU the Vulkan buffers
        are charged to it). The drop in available RAM is logged separately: it also
        counts every other program."""
        return self.peak_private

    @property
    def system_drop(self) -> int:
        return max(0, self.available_before - self.lowest_available)

    def as_dict(self) -> dict[str, Any]:
        return {"load_s": round(self.load_seconds, 2),
                "peak_ram_used_gb": round(self.peak_used / 2**30, 2),
                "system_available_drop_gb": round(self.system_drop / 2**30, 2),
                "lowest_available_gb": round(self.lowest_available / 2**30, 2),
                "requests": self.requests,
                "prompt_tok_s": round(self.prompt_tokens / (self.prompt_ms / 1000), 1)
                if self.prompt_ms else None,
                "gen_tok_s": round(self.gen_tokens / (self.gen_ms / 1000), 1)
                if self.gen_ms else None}


@dataclass
class LlamaServer:
    name: str  # candidate name (llm/manifest.json)
    ctx: int | None = None
    ngl: int = 99
    n_cpu_moe: int | None = None  # None = auto (smallest that fits)
    flash_attn: bool = True
    kv_type: str = "q8_0"
    threads: int = 8
    port: int | None = None
    start_timeout: float = 600.0
    request_timeout: float = 600.0
    extra_args: list[str] = field(default_factory=list)
    stats: ServerStats = field(default_factory=ServerStats)
    plan: MemoryPlan | None = None
    _proc: subprocess.Popen[bytes] | None = None
    _sampler: threading.Thread | None = None
    _stop_sampling: threading.Event = field(default_factory=threading.Event)
    _log_path: Path | None = None
    mmproj_cpu_fallback: bool = False
    # Measured (b12503, Radeon 780M): the Vulkan vision encoder crashes llama-server on
    # page-sized images (Qwen3.5-9B at 594 x 768 px, Qwen3.5-4B above ~800 px wide);
    # on the CPU it costs ~0.5 s more per image. So the encoder runs on the CPU.
    mmproj_gpu: bool = False

    @property
    def cand(self) -> dict[str, Any]:
        return candidate(self.name)

    @property
    def url(self) -> str:
        return f"http://{HOST}:{self.port}"

    @property
    def model_id(self) -> str:
        """Identifies the weights in cache keys: candidate + model sha256."""
        m = models.spec()[self.cand["model"]]
        return f"{self.cand['name']}@{m.sha256[:12]}"

    def command(self) -> list[str]:
        exe = llama_server()
        if exe is None:
            raise LLMUnavailable("llama.cpp missing (pdf2pptx setup llm)")
        c = self.cand
        mp = models.path(c["model"])
        if mp is None:
            raise LLMUnavailable(f"{c['name']}: model missing (pdf2pptx setup llm)")
        cmd = [str(exe), "-m", str(mp), "--host", HOST, "--port", str(self.port),
               "-c", str(self.ctx or c.get("ctx", 8192)), "-ngl", str(self.ngl),
               "-np", "1", "-t", str(self.threads), "--jinja", "--no-webui",
               "-fa", "on" if self.flash_attn else "off",
               "-ctk", c.get("kv_type", self.kv_type), "-ctv", c.get("kv_type", self.kv_type),
               "--seed", "1234", "--temp", "0",
               # no host-RAM prompt cache (default 8 GB; prompts are never reused here)
               "--cache-ram", "0"]
        if c.get("mmproj"):
            pp = models.path(c["mmproj"])
            if pp is None:
                raise LLMUnavailable(f"{c['name']}: projector missing (pdf2pptx setup llm)")
            cmd += ["--mmproj", str(pp)]
        if (c.get("mmproj") and not self.mmproj_gpu
                and "--no-mmproj-offload" not in self.extra_args):
            cmd += ["--no-mmproj-offload"]
        if self.plan is not None and self.plan.n_cpu_moe:
            cmd += ["--n-cpu-moe", str(self.plan.n_cpu_moe)]
        return cmd + list(c.get("server_args", [])) + list(self.extra_args)

    def _sample(self) -> None:
        proc: psutil.Process | None = None
        while not self._stop_sampling.wait(0.25):
            a = psutil.virtual_memory().available
            self.stats.lowest_available = min(self.stats.lowest_available, a)
            try:
                if proc is None and self._proc is not None:
                    proc = psutil.Process(self._proc.pid)
                if proc is not None:
                    mi = proc.memory_info()
                    # committed memory (GPU buffers) or resident pages (memory-mapped
                    # weights are not committed), whichever is larger
                    used = max(int(getattr(mi, "private", 0) or 0), int(mi.rss))
                    self.stats.peak_private = max(self.stats.peak_private, used)
            except (psutil.Error, OSError):
                proc = None

    def start(self, auto_moe: bool = True, log_dir: Path | None = None) -> LlamaServer:
        if self._proc is not None:
            return self
        ctx = self.ctx or int(self.cand.get("ctx", 8192))
        self.plan = plan_memory(self.name, ctx, self.n_cpu_moe)
        if not self.plan.fits or (self.plan.n_cpu_moe and not auto_moe
                                  and self.n_cpu_moe is None):
            raise LLMUnavailable(f"{self.name}: needs {self.plan.need / 2**30:.1f} GB, "
                                 f"budget {self.plan.budget / 2**30:.1f} GB "
                                 f"(available - 8 GB); {self.plan.suggestion}")
        self.port = self.port or _free_port()
        cmd = self.command()
        self._log_path = (log_dir or Path(os.environ.get("TEMP", "."))) / \
            f"llama_{self.cand['model']}_{self.port}.log"
        self._log_path.parent.mkdir(parents=True, exist_ok=True)
        LOG.info("starting %s (n_cpu_moe=%s, need %.1f GB of %.1f GB budget)", self.name,
                 self.plan.n_cpu_moe, self.plan.need / 2**30, self.plan.budget / 2**30)
        self.stats = ServerStats(available_before=psutil.virtual_memory().available)
        self.stats.lowest_available = self.stats.available_before
        self._stop_sampling.clear()
        self._sampler = threading.Thread(target=self._sample, daemon=True)
        self._sampler.start()
        t0 = time.perf_counter()
        logf = self._log_path.open("wb")
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        self._proc = subprocess.Popen(cmd, stdout=logf, stderr=subprocess.STDOUT,
                                      creationflags=flags)
        try:
            while True:
                if self._proc.poll() is not None:
                    tail = self._log_path.read_text(encoding="utf-8", errors="replace")[-1500:]
                    raise LLMUnavailable(f"{self.name}: llama-server exited "
                                         f"({self._proc.returncode}): {tail}")
                if self.healthy():
                    break
                if time.perf_counter() - t0 > self.start_timeout:
                    raise LLMUnavailable(f"{self.name}: not ready after {self.start_timeout} s")
                time.sleep(0.25)
        except BaseException:
            self.stop()
            raise
        self.stats.load_seconds = time.perf_counter() - t0
        LOG.info("%s ready in %.1fs", self.name, self.stats.load_seconds)
        return self

    def healthy(self) -> bool:
        try:
            with urllib.request.urlopen(self.url + "/health", timeout=2) as r:
                return bool(r.status == 200 and json.load(r).get("status") == "ok")
        except (urllib.error.URLError, OSError, ValueError):
            return False

    def stop(self) -> None:
        self._stop_sampling.set()
        if self._proc is not None and self.stats.requests:
            record_peak(str(self.cand["name"]), self.stats.peak_used)
        if self._proc is not None:
            if self._proc.poll() is None:
                self._proc.terminate()
                try:
                    self._proc.wait(15)
                except subprocess.TimeoutExpired:
                    self._proc.kill()
                    self._proc.wait(5)
            self._proc = None
        if self._sampler is not None:
            self._sampler.join(2)
            self._sampler = None

    def __enter__(self) -> LlamaServer:
        return self.start()

    def __exit__(self, *exc: object) -> None:
        self.stop()

    # ----------------------------------------------------------------- requests

    def chat(self, prompt: str, *, images: list[bytes] | None = None,
             schema: dict[str, Any] | None = None, system: str | None = None,
             max_tokens: int = 1024) -> dict[str, Any]:
        """One chat completion. Returns {"content", "json" (parsed or None), "timings"}.

        Measured (b12503, Radeon 780M): the vision encoder crashes llama-server on
        Vulkan for images above ~800 px; callers cap images at ``MAX_IMAGE_PX``. If the
        server dies anyway, it restarts once with the encoder on the CPU."""
        try:
            return self._chat(prompt, images=images, schema=schema, system=system,
                              max_tokens=max_tokens)
        except (urllib.error.URLError, ConnectionError, TimeoutError) as exc:
            died = False
            if self._proc is not None:
                try:
                    self._proc.wait(5)  # a crashed server takes a moment to exit
                    died = True
                except subprocess.TimeoutExpired:
                    died = False
            if not (images and died and self.mmproj_gpu):
                raise
            LOG.warning("%s died on an image request (%s); restarting with the vision "
                        "encoder on the CPU", self.name, exc)
            self.stop()
            self.mmproj_gpu = False
            self.mmproj_cpu_fallback = True
            self.start()
            return self._chat(prompt, images=images, schema=schema, system=system,
                              max_tokens=max_tokens)

    def _chat(self, prompt: str, *, images: list[bytes] | None = None,
              schema: dict[str, Any] | None = None, system: str | None = None,
              max_tokens: int = 1024) -> dict[str, Any]:
        if self._proc is None:
            raise LLMUnavailable(f"{self.name}: server not started")
        content: list[dict[str, Any]] = []
        for img in images or []:
            b64 = base64.b64encode(img).decode("ascii")
            content.append({"type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{b64}"}})
        content.append({"type": "text", "text": prompt})
        msgs: list[dict[str, Any]] = []
        if system:
            msgs.append({"role": "system", "content": system})
        msgs.append({"role": "user", "content": content if images else prompt})
        body: dict[str, Any] = {"messages": msgs, "temperature": 0.0, "seed": 1234,
                                "max_tokens": max_tokens, "cache_prompt": False,
                                "chat_template_kwargs": {"enable_thinking": False}}
        if self.cand.get("reasoning"):
            body["reasoning_effort"] = self.cand["reasoning"]
        if schema is not None:
            body["response_format"] = {"type": "json_schema",
                                       "json_schema": {"name": "out", "schema": schema}}
        req = urllib.request.Request(self.url + "/v1/chat/completions",
                                     data=json.dumps(body).encode("utf-8"),
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=self.request_timeout) as r:
            resp = json.load(r)
        msg = resp["choices"][0]["message"]
        text = msg.get("content") or ""
        t = resp.get("timings") or {}
        self.stats.requests += 1
        self.stats.prompt_tokens += int(t.get("prompt_n") or 0)
        self.stats.prompt_ms += float(t.get("prompt_ms") or 0.0)
        self.stats.gen_tokens += int(t.get("predicted_n") or 0)
        self.stats.gen_ms += float(t.get("predicted_ms") or 0.0)
        parsed = None
        if schema is not None:
            try:
                parsed = json.loads(text)
            except ValueError:
                parsed = None
        return {"content": text, "json": parsed, "timings": t,
                "finish_reason": resp["choices"][0].get("finish_reason")}
