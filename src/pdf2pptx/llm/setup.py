"""``pdf2pptx setup llm``: llama.cpp (Windows Vulkan build) + the GGUF candidates.

Downloads only happen here (PLAN.md §8 rule 9). llama.cpp comes from the newest GitHub
release that has a ``win-vulkan-x64`` asset (``--llama-tag`` pins one); its sha256 is
the digest GitHub publishes for the asset. It is unpacked into ``tools/llama.cpp`` and
the version is recorded in ``tools/llama.cpp/VERSION.json`` and in the installed
``<models>/manifest.json``. The GGUF files reuse ``ml/models.py`` (pack ``llm``).
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import shutil
import time
import urllib.request
import zipfile
from collections.abc import Callable
from functools import cache
from pathlib import Path
from typing import Any

from .. import log
from ..config import PROJECT_ROOT
from ..ml import models

LOG = log.get(__name__)

SPEC_PATH = Path(__file__).with_name("manifest.json")
GITHUB_API = "https://api.github.com/repos/{repo}/releases?per_page=30"


@cache
def spec() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(SPEC_PATH.read_text(encoding="utf-8"))
    return data


def llama_dir() -> Path:
    return PROJECT_ROOT / spec()["llama_cpp"]["dir"]


def llama_server() -> Path | None:
    exe = llama_dir() / "llama-server.exe"
    return exe if exe.is_file() else None


def llama_version() -> dict[str, Any] | None:
    p = llama_dir() / "VERSION.json"
    try:
        v: dict[str, Any] = json.loads(p.read_text(encoding="utf-8"))
        return v
    except (OSError, ValueError):
        return None


def candidates() -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = spec()["candidates"]
    return out


def candidate(name: str) -> dict[str, Any]:
    for c in candidates():
        if c["name"] == name or c["model"] == name:
            return c
    raise KeyError(name)


def candidate_files(c: dict[str, Any]) -> list[str]:
    return [n for n in (c["model"], c.get("mmproj")) if n]


def installed(c: dict[str, Any], root: Path | None = None) -> bool:
    return models.available(*candidate_files(c), root=root)


def _fetch_json(url: str) -> Any:
    req = urllib.request.Request(url, headers={"User-Agent": "pdf2pptx-setup",
                                               "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.load(resp)


def resolve_release(tag: str | None = None) -> dict[str, Any]:
    """The release to install: ``tag`` or the newest one with a Windows Vulkan zip."""
    ll = spec()["llama_cpp"]
    for rel in _fetch_json(GITHUB_API.format(repo=ll["repo"])):
        if tag and rel["tag_name"] != tag:
            continue
        want = ll["asset"].format(tag=rel["tag_name"])
        for a in rel.get("assets", []):
            if a["name"] == want:
                digest = (a.get("digest") or "").removeprefix("sha256:")
                return {"tag": rel["tag_name"], "published": rel.get("published_at"),
                        "asset": a["name"], "url": a["browser_download_url"],
                        "size": a["size"], "sha256": digest}
    raise RuntimeError(f"no llama.cpp release with {ll['asset']!r}"
                       + (f" for tag {tag}" if tag else ""))


def install_llama(tag: str | None = None, progress: Callable[[str], None] = print,
                  force: bool = False) -> dict[str, Any]:
    rel = resolve_release(tag)
    have = llama_version()
    if have and have.get("tag") == rel["tag"] and llama_server() and not force:
        progress(f"llama.cpp {rel['tag']} already installed in {llama_dir()}")
        return have
    dest = llama_dir()
    dest.parent.mkdir(parents=True, exist_ok=True)
    zpath = dest.parent / rel["asset"]
    progress(f"- llama.cpp {rel['tag']} ({rel['size'] / 2**20:.1f} MB, MIT) {rel['url']}")
    req = urllib.request.Request(rel["url"], headers={"User-Agent": "pdf2pptx-setup"})
    h = hashlib.sha256()
    t0 = time.monotonic()
    with urllib.request.urlopen(req, timeout=120) as resp, zpath.open("wb") as fh:
        while chunk := resp.read(1 << 20):
            fh.write(chunk)
            h.update(chunk)
    if rel["sha256"] and h.hexdigest() != rel["sha256"]:
        zpath.unlink()
        raise RuntimeError(f"llama.cpp zip sha256 mismatch ({h.hexdigest()})")
    tmp = dest.with_name(dest.name + ".new")
    shutil.rmtree(tmp, ignore_errors=True)
    with zipfile.ZipFile(zpath) as z:
        z.extractall(tmp)
    # some releases nest the binaries in a folder
    inner = next((p.parent for p in tmp.rglob("llama-server.exe")), None)
    if inner is None:
        raise RuntimeError("llama-server.exe not found in the release zip")
    if inner != tmp:
        for p in inner.iterdir():
            p.rename(tmp / p.name)
    shutil.rmtree(dest, ignore_errors=True)
    tmp.rename(dest)
    zpath.unlink()
    info = rel | {"sha256": h.hexdigest(), "installed": time.strftime("%Y-%m-%d"),
                  "seconds": round(time.monotonic() - t0, 1)}
    (dest / "VERSION.json").write_text(json.dumps(info, indent=1) + "\n", encoding="utf-8",
                                       newline="\n")
    return info


def plan(names: list[str] | None = None) -> list[models.ModelFile]:
    """The GGUF files still to download (``names``: candidate names; None = all)."""
    want: set[str] | None = None
    if names:
        want = set()
        for n in names:
            want.update(candidate_files(candidate(n)))
    return [m for m in models.spec().values() if m.pack == "llm"
            and models.path(m.name) is None and (want is None or m.name in want)]


def install(names: list[str] | None = None, *, assume_yes: bool = False,
            ask: Callable[[str], bool] | None = None, tag: str | None = None,
            progress: Callable[[str], None] = print) -> dict[str, Any]:
    """llama.cpp + the GGUF candidates. Always asks (with the total size) unless
    ``assume_yes``: the candidates are tens of GB."""
    todo = plan(names)
    total = sum(m.size for m in todo)
    if todo and not assume_yes:
        lines = "\n".join(f"  {m.name:24s} {m.size / 2**30:6.2f} GB  {m.license}  {m.repo}"
                          for m in todo)
        question = (f"{lines}\nDownload {total / 2**30:.2f} GB of GGUF models into "
                    f"{models.models_dir() / 'llm'} (plus llama.cpp, ~40 MB)?")
        if ask is None or not ask(question):
            return {"installed": [], "skipped": [m.name for m in todo],
                    "reason": "not confirmed (use --yes)"}
    llama = install_llama(tag, progress)
    res = models.install("llm", assume_yes=True, progress=progress,
                         names=[m.name for m in todo])
    write_manifest()
    return {"llama_cpp": llama.get("tag"), **res}


def write_manifest() -> Path:
    """``<models>/manifest.json`` plus the llama.cpp version and the candidates."""
    out = models.write_manifest()
    data = json.loads(out.read_text(encoding="utf-8"))
    data["llama_cpp"] = llama_version()
    data["llm_candidates"] = [c | {"installed": installed(c)} for c in candidates()]
    out.write_text(json.dumps(data, indent=1, ensure_ascii=False) + "\n", encoding="utf-8",
                   newline="\n")
    return out


def remove(keep: list[str], progress: Callable[[str], None] = print) -> list[str]:
    """Delete the GGUF files of every candidate not in ``keep`` (after the benchmark)."""
    keep_files = {f for n in keep for f in candidate_files(candidate(n))}
    gone = []
    for m in models.spec().values():
        if m.pack != "llm" or m.name in keep_files:
            continue
        p = models.models_dir() / m.file
        if p.is_file():
            p.unlink()
            gone.append(m.name)
            progress(f"removed {m.name} ({m.size / 2**30:.2f} GB)")
            with contextlib.suppress(OSError):
                p.parent.rmdir()
    write_manifest()
    return gone
