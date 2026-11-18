"""Model files: the pinned manifest, where they live, and ``pdf2pptx setup models``.

Downloads happen only here (PLAN.md §8 rule 9): resumable (HTTP Range on a ``.part``
file), checked against the manifest's sha256, installed atomically. Conversion only
looks files up (:func:`path`) and never touches the network.
"""

from __future__ import annotations

import hashlib
import json
import os
import time
import urllib.error
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from functools import cache
from pathlib import Path
from typing import Any

from .. import log
from ..config import default_models_dir

LOG = log.get(__name__)

SPEC_PATH = Path(__file__).with_name("manifest.json")
# pack "llm" (M9): GGUF files for llama.cpp, same pinning and install code
LLM_SPEC_PATH = Path(__file__).resolve().parents[1] / "llm" / "manifest.json"
ASK_ABOVE_BYTES = 1 << 30  # ask before downloading more than 1 GB in total
HF_URL = "https://huggingface.co/{repo}/resolve/{revision}/{path}"


@dataclass(frozen=True)
class ModelFile:
    name: str
    pack: str
    role: str
    version: str
    license: str
    upstream: str
    repo: str
    revision: str
    path: str
    file: str
    size: int
    sha256: str

    @property
    def url(self) -> str:
        return HF_URL.format(repo=self.repo, revision=self.revision, path=self.path)


@cache
def spec() -> dict[str, ModelFile]:
    out: dict[str, ModelFile] = {}
    for sp in (SPEC_PATH, LLM_SPEC_PATH):
        if sp.is_file():
            data = json.loads(sp.read_text(encoding="utf-8"))
            out.update({m["name"]: ModelFile(**m) for m in data["models"]})
    return out


def models_dir() -> Path:
    env = os.environ.get("PDF2PPTX_MODELS_DIR")
    return Path(env) if env else default_models_dir()


def path(name: str, root: Path | None = None) -> Path | None:
    """The installed file of model ``name``, or None when it is missing (size check
    only: the sha256 is verified at install time and by ``verify``)."""
    m = spec().get(name)
    if m is None:
        return None
    p = (root or models_dir()) / m.file
    try:
        return p if p.is_file() and p.stat().st_size == m.size else None
    except OSError:
        return None


def available(*names: str, root: Path | None = None) -> bool:
    return all(path(n, root) is not None for n in names)


def fingerprint(*names: str, root: Path | None = None) -> str:
    """Short id of which of ``names`` are installed (part of the OCR stage cache key)."""
    have = [f"{n}:{spec()[n].sha256[:8]}" for n in names if path(n, root) is not None]
    return hashlib.sha256(";".join(have).encode()).hexdigest()[:10] if have else "none"


def sha256_of(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def status(root: Path | None = None, verify: bool = False,
           pack: str | None = None) -> list[dict[str, Any]]:
    out = []
    for m in spec().values():
        if pack is not None and m.pack != pack:
            continue
        p = (root or models_dir()) / m.file
        st = "missing"
        if p.is_file():
            st = "ok" if p.stat().st_size == m.size else "size mismatch"
            if st == "ok" and verify and sha256_of(p) != m.sha256:
                st = "sha256 mismatch"
        out.append({"name": m.name, "file": str(p), "size": m.size, "status": st,
                    "license": m.license, "version": m.version})
    return out


def _download(m: ModelFile, dest: Path, progress: Callable[[str], None],
              retries: int = 4) -> None:
    part = dest.with_name(dest.name + ".part")
    dest.parent.mkdir(parents=True, exist_ok=True)
    for attempt in range(1, retries + 1):
        have = part.stat().st_size if part.is_file() else 0
        if have > m.size:
            part.unlink()
            have = 0
        if have < m.size:
            req = urllib.request.Request(m.url, headers={"User-Agent": "pdf2pptx-setup"})
            if have:
                req.add_header("Range", f"bytes={have}-")
            try:
                with urllib.request.urlopen(req, timeout=60) as resp:
                    if have and resp.status != 206:  # server ignored the range: restart
                        have = 0
                    mode = "ab" if have else "wb"
                    t0, last = time.monotonic(), 0.0
                    with part.open(mode) as fh:
                        while chunk := resp.read(1 << 20):
                            fh.write(chunk)
                            have += len(chunk)
                            now = time.monotonic()
                            if now - last > 2.0:
                                last = now
                                rate = have / max(now - t0, 1e-3) / 2**20
                                progress(f"  {m.name}: {have / 2**20:.1f} / "
                                         f"{m.size / 2**20:.1f} MB ({rate:.1f} MB/s)")
            except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
                progress(f"  {m.name}: attempt {attempt} interrupted ({exc}); resuming")
                time.sleep(min(2.0 * attempt, 8.0))
                continue
        if part.stat().st_size != m.size:
            continue
        got = sha256_of(part)
        if got != m.sha256:
            part.unlink()
            raise RuntimeError(f"{m.name}: sha256 mismatch ({got} != {m.sha256}); "
                               "the partial file was removed, run setup again")
        part.replace(dest)
        return
    raise RuntimeError(f"{m.name}: download failed after {retries} attempts")


def install(pack: str = "lite", root: Path | None = None, *, assume_yes: bool = False,
            ask: Callable[[str], bool] | None = None,
            progress: Callable[[str], None] = print,
            names: list[str] | None = None) -> dict[str, Any]:
    """Download the missing files of ``pack`` (verified), then write
    ``<root>/manifest.json``. Asks before more than 1 GB in total."""
    root = root or models_dir()
    todo = [m for m in spec().values() if m.pack == pack and path(m.name, root) is None
            and (names is None or m.name in names)]
    total = sum(m.size for m in todo)
    if total > ASK_ABOVE_BYTES and not assume_yes:
        question = f"Download {total / 2**30:.2f} GB of models into {root}?"
        if ask is None or not ask(question):
            return {"installed": [], "skipped": [m.name for m in todo],
                    "reason": "more than 1 GB: confirm with --yes"}
    progress(f"{len(todo)} file(s) to download, {total / 2**20:.1f} MB, into {root}")
    done = []
    for m in todo:
        progress(f"- {m.name} ({m.size / 2**20:.1f} MB, {m.license}) from {m.repo}")
        _download(m, root / m.file, progress)
        done.append(m.name)
    write_manifest(root)
    return {"installed": done, "skipped": [], "total_bytes": total}


def write_manifest(root: Path | None = None) -> Path:
    """The installed copy of the manifest: every pinned file plus its state."""
    from dataclasses import asdict

    root = root or models_dir()
    root.mkdir(parents=True, exist_ok=True)
    about = json.loads(SPEC_PATH.read_text(encoding="utf-8")).get("about", "")
    st = {s["name"]: s["status"] for s in status(root)}
    entries = [asdict(m) | {"url": m.url, "installed": st.get(m.name) == "ok"}
               for m in spec().values()]
    out = root / "manifest.json"
    out.write_text(json.dumps({"about": about, "models": entries}, indent=1,
                              ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    return out
