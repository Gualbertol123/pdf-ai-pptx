"""``doctor``: load every installed model once and report its provider and timings."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from . import inpaint, layout, models, ocr_engine, runtime, winocr

WARMUPS: dict[str, Callable[[runtime.Session], None]] = {
    "ocr_det": ocr_engine._det_warmup,
    "ocr_cls": ocr_engine._cls_warmup,
    "ocr_rec_latin": ocr_engine._rec_warmup,
    "ocr_rec_multi": ocr_engine._rec_warmup,
    "layout": layout._warmup,
    "inpaint": inpaint._lama_warmup,
}


def run() -> dict[str, Any]:
    """Per model: installed, provider, load / warm-up seconds, why not DirectML."""
    out: dict[str, Any] = {"device": runtime.device_pref(), "models_dir": str(
        models.models_dir()), "models": {}}
    for name, warm in WARMUPS.items():
        if models.path(name) is None:
            out["models"][name] = {"installed": False}
            continue
        s = runtime.session(name, warm)
        if s is None:
            out["models"][name] = {"installed": True, "provider": None,
                                   "error": "could not be loaded"}
            continue
        st = s.stats
        out["models"][name] = {"installed": True, "provider": st.provider,
                               "load_s": round(st.load_s, 2), "warmup_s": round(st.warmup_s, 2),
                               "fallback": st.fallback}
    out["windows_ocr"] = winocr.available()
    return out
