"""``pdf2pptx doctor``: report what this machine can do. Never raises."""

from __future__ import annotations

import os
import platform
import sys
from pathlib import Path
from typing import Any

from .config import Config


def check_powerpoint(timeout: float = 90.0) -> dict[str, Any]:
    """Probe COM through the engine (its worker process has its own timeouts, so a
    hung PowerPoint cannot hang doctor)."""
    if sys.platform != "win32":
        return {"available": False, "detail": "not Windows"}
    try:
        from .verify import ppt_com

        version = ppt_com.probe(timeout=timeout)
    except Exception as exc:
        return {"available": False, "detail": f"{type(exc).__name__}: {exc}"}
    return {"available": True, "detail": None, "version": version}


def check_onnxruntime() -> dict[str, Any]:
    try:
        import onnxruntime as ort  # type: ignore[import-not-found]
    except ImportError:
        return {"installed": False}
    return {"installed": True, "version": ort.__version__,
            "providers": list(ort.get_available_providers())}


def check_vulkan() -> dict[str, Any]:
    sysroot = Path(os.environ.get("SYSTEMROOT", r"C:\Windows"))
    dll = sysroot / "System32" / "vulkan-1.dll"
    return {"vulkan-1.dll": dll.is_file(), "path": str(dll)}


def check_ram() -> dict[str, Any]:
    import psutil

    vm = psutil.virtual_memory()
    return {"total_gb": round(vm.total / 2**30, 1), "available_gb": round(vm.available / 2**30, 1)}


def check_gpu() -> dict[str, Any]:
    if sys.platform != "win32":
        return {"names": []}
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        wmi = win32com.client.GetObject("winmgmts:")
        gpus = [
            {"name": str(g.Name), "driver": str(g.DriverVersion),
             "dedicated_mb": round(int(g.AdapterRAM or 0) / 2**20)}
            for g in wmi.InstancesOf("Win32_VideoController")
        ]
        return {"gpus": gpus}
    except Exception as exc:
        return {"gpus": [], "error": str(exc)}


def check_models(load: bool = True) -> dict[str, Any]:
    """M8 models: installed files and, with ``load``, the provider each one runs on."""
    try:
        from .ml import models, probe

        if not load:
            return {"models": {s["name"]: {"installed": s["status"] == "ok"}
                               for s in models.status()}}
        return probe.run()
    except Exception as exc:
        return {"error": f"{type(exc).__name__}: {exc}"}


def run(cfg: Config, load_models: bool = True) -> dict[str, Any]:
    return {
        "python": {"version": platform.python_version(), "executable": sys.executable},
        "os": platform.platform(),
        "cpu": {"name": platform.processor(), "logical_cores": os.cpu_count()},
        "ram": check_ram(),
        "gpu": check_gpu(),
        "powerpoint_com": check_powerpoint(cfg.com_timeout_s + 30),
        "onnxruntime": check_onnxruntime(),
        "vulkan": check_vulkan(),
        "models_dir": {"path": str(cfg.models_dir), "exists": cfg.models_dir.is_dir()},
        "ml": check_models(load_models),
        "tier": cfg.resolved_tier(),
        "mode": cfg.mode,
    }


def format_report(r: dict[str, Any]) -> str:
    ppt = r["powerpoint_com"]
    ort = r["onnxruntime"]
    gpus = r["gpu"].get("gpus", [])
    lines = [
        "pdf2pptx doctor",
        f"  Python          {r['python']['version']} ({r['python']['executable']})",
        f"  OS              {r['os']}",
        f"  CPU             {r['cpu']['name']} - {r['cpu']['logical_cores']} logical cores",
        f"  RAM             {r['ram']['total_gb']} GB total, {r['ram']['available_gb']} GB free",
    ]
    for g in gpus or [{"name": "unknown", "driver": "?", "dedicated_mb": 0}]:
        lines.append(f"  GPU             {g['name']} (driver {g['driver']}, "
                     f"{g['dedicated_mb']} MB dedicated)")
    lines.append(
        "  PowerPoint COM  "
        + (f"available (version {ppt['version']})" if ppt["available"]
           else f"NOT available - {ppt['detail']}")
    )
    lines.append(
        "  onnxruntime     "
        + (f"{ort['version']}, providers: {', '.join(ort['providers'])}" if ort["installed"]
           else "not installed (needed from Milestone 8)")
    )
    vulkan = "present" if r["vulkan"]["vulkan-1.dll"] else "MISSING"
    lines.append(f"  Vulkan          vulkan-1.dll {vulkan}")
    md = r["models_dir"]
    state = "exists" if md["exists"] else "not created yet"
    lines.append(f"  Models dir      {md['path']} ({state})")
    ml = r.get("ml", {})
    if "error" in ml:
        lines.append(f"  Models          unavailable: {ml['error']}")
    for name, m in sorted(ml.get("models", {}).items()):
        if not m.get("installed"):
            lines.append(f"  - {name:15s} not installed (pdf2pptx setup models --lite)")
        elif "provider" not in m:
            lines.append(f"  - {name:15s} installed")
        else:
            prov = (m.get("provider") or "FAILED").replace("ExecutionProvider", "")
            why = f" ({m['fallback']})" if m.get("fallback") else ""
            lines.append(f"  - {name:15s} {prov}{why}, load {m.get('load_s')} s, "
                         f"warm-up {m.get('warmup_s')} s")
    if "windows_ocr" in ml:
        lines.append(f"  Windows OCR     {'available' if ml['windows_ocr'] else 'not available'}"
                     " (fallback OCR engine)")
    if ml.get("device") and ml["device"] != "auto":
        lines.append(f"  ML device       {ml['device']} (PDF2PPTX_ML_DEVICE)")
    lines.append(f"  Tier / mode     {r['tier']} / {r['mode']}")
    return "\n".join(lines)
