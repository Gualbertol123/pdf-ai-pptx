"""Configuration: defaults <- pdf2pptx.toml <- CLI overrides.

Modes (PLAN.md §3): fidelity | balanced | editable | deck.
Tiers: lite | max (``auto`` picks from hardware and installed models).
"""

from __future__ import annotations

import os
import tomllib
from dataclasses import asdict, dataclass, field, fields, replace
from pathlib import Path
from typing import Any

MODES = ("fidelity", "balanced", "editable", "deck")
TIERS = ("auto", "lite", "max")
TEXT_LAYOUTS = ("blocks", "lines")  # lines = the M2 baseline, one box per line

PROJECT_ROOT = Path(__file__).resolve().parents[2]
CONFIG_FILENAME = "pdf2pptx.toml"


def default_models_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "pdf2pptx" / "models"


@dataclass(frozen=True, kw_only=True)
class ModeLimits:
    """What the tuner/builder may do in a mode (M5 tuner: verify/tuner.py)."""

    forced_line_breaks: bool  # tuner may insert a:br at the PDF's break positions
    max_font_size_change: float  # fraction, e.g. 0.03 = ±3%
    max_char_spacing_pt: float  # max |change| of character spacing (spc), pt
    max_width_change: float  # fraction of the analysed text box width
    max_offset_pt: float  # max |shift| of a text box / shape by the tuner
    slide_pass: bool  # pixel-diff translation pass over shapes and images
    use_placeholders: bool
    use_theme: bool
    autofit: str  # none | shrink
    tune: bool


MODE_LIMITS: dict[str, ModeLimits] = {
    "fidelity": ModeLimits(
        forced_line_breaks=True,
        max_font_size_change=0.10,
        max_char_spacing_pt=1.5,
        max_width_change=0.25,
        max_offset_pt=6.0,
        slide_pass=True,
        use_placeholders=False,
        use_theme=False,
        autofit="none",
        tune=True,
    ),
    "balanced": ModeLimits(
        forced_line_breaks=False,
        max_font_size_change=0.05,
        max_char_spacing_pt=0.8,
        max_width_change=0.15,
        max_offset_pt=4.0,
        slide_pass=True,
        use_placeholders=True,
        use_theme=True,
        autofit="none",
        tune=True,
    ),
    "editable": ModeLimits(
        forced_line_breaks=False,
        max_font_size_change=0.03,
        max_char_spacing_pt=0.4,
        max_width_change=0.10,
        max_offset_pt=3.0,
        slide_pass=False,
        use_placeholders=True,
        use_theme=True,
        autofit="shrink",
        tune=False,
    ),
    "deck": ModeLimits(
        forced_line_breaks=False,
        max_font_size_change=0.03,
        max_char_spacing_pt=0.0,
        max_width_change=0.10,
        max_offset_pt=3.0,
        slide_pass=False,
        use_placeholders=True,
        use_theme=True,
        autofit="shrink",
        tune=True,
    ),
}


@dataclass(frozen=True, kw_only=True)
class Config:
    mode: str = "balanced"
    tier: str = "auto"
    work_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "work")
    models_dir: Path = field(default_factory=default_models_dir)
    corpus_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "corpus")
    out_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "out")
    workers: int = 6
    render_dpi: int = 150
    hires_dpi: int = 300
    tune: bool | None = None  # None = mode default
    tune_rounds: int = 4
    com_timeout_s: float = 60.0
    seed: int = 1234
    text_layout: str = "blocks"

    def __post_init__(self) -> None:
        if self.mode not in MODES:
            raise ValueError(f"mode must be one of {MODES}, got {self.mode!r}")
        if self.tier not in TIERS:
            raise ValueError(f"tier must be one of {TIERS}, got {self.tier!r}")
        if self.workers < 1:
            raise ValueError("workers must be >= 1")
        if self.text_layout not in TEXT_LAYOUTS:
            raise ValueError(f"text_layout must be one of {TEXT_LAYOUTS}, "
                             f"got {self.text_layout!r}")

    @property
    def limits(self) -> ModeLimits:
        return MODE_LIMITS[self.mode]

    @property
    def tune_enabled(self) -> bool:
        return self.limits.tune if self.tune is None else self.tune

    @property
    def scores_dir(self) -> Path:
        return self.work_dir / "scores"

    def resolved_tier(self) -> str:
        return detect_tier(self) if self.tier == "auto" else self.tier

    def as_dict(self) -> dict[str, Any]:
        d = asdict(self)
        return {k: str(v) if isinstance(v, Path) else v for k, v in d.items()}


_PATH_FIELDS = {"work_dir", "models_dir", "corpus_dir", "out_dir"}


def _coerce(values: dict[str, Any], base: Path) -> dict[str, Any]:
    known = {f.name for f in fields(Config)}
    out: dict[str, Any] = {}
    for k, v in values.items():
        k = k.replace("-", "_")
        if k not in known:
            raise ValueError(f"unknown config key {k!r}")
        if v is None:
            continue
        if k in _PATH_FIELDS:
            p = Path(os.path.expandvars(str(v)))
            v = p if p.is_absolute() else (base / p)
        out[k] = v
    return out


def load_config(
    overrides: dict[str, Any] | None = None, config_path: Path | None = None
) -> Config:
    """Defaults, then ``pdf2pptx.toml`` (project root or given path), then overrides."""
    cfg = Config()
    path = config_path or (PROJECT_ROOT / CONFIG_FILENAME)
    if path.is_file():
        with path.open("rb") as fh:
            data = tomllib.load(fh)
        cfg = replace(cfg, **_coerce(data.get("pdf2pptx", data), path.parent))
    if overrides:
        cfg = replace(cfg, **_coerce(overrides, Path.cwd()))
    return cfg


def detect_tier(cfg: Config) -> str:
    """Stub until M9: ``max`` needs llama.cpp + at least one GGUF model + >= 24 GB RAM."""
    try:
        import psutil

        ram_gb = psutil.virtual_memory().total / 2**30
    except Exception:  # pragma: no cover - psutil is a dependency
        ram_gb = 0.0
    llama = PROJECT_ROOT / "tools" / "llama.cpp" / "llama-server.exe"
    has_gguf = cfg.models_dir.is_dir() and any(cfg.models_dir.rglob("*.gguf"))
    return "max" if (llama.is_file() and has_gguf and ram_gb >= 24) else "lite"
