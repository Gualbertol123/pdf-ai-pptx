from __future__ import annotations

from pathlib import Path

import pytest

from pdf2pptx.build import units
from pdf2pptx.config import MODE_LIMITS, Config, load_config


def test_pt_emu() -> None:
    assert units.pt_to_emu(1) == 12_700
    assert units.pt_to_emu(72) == units.EMU_PER_INCH
    assert units.emu_to_pt(12_192_000) == pytest.approx(960.0)
    assert units.pt_to_emu(-1.5) == -19_050
    for pt in (0.0, 0.1, 13.37, 540.0):
        assert units.emu_to_pt(units.pt_to_emu(pt)) == pytest.approx(pt, abs=1 / 12_700)


def test_px_and_ooxml_units() -> None:
    assert units.pt_to_px(72, 150) == 150
    assert units.px_to_pt(150, 150) == 72
    assert units.angle_to_ooxml(90) == 5_400_000
    assert units.pct_to_ooxml(0.5) == 50_000
    assert units.pt_to_centipoints(18) == 1800


def test_defaults_and_modes() -> None:
    cfg = Config()
    assert cfg.mode == "balanced"
    assert cfg.tune_enabled
    assert not Config(mode="editable").tune_enabled
    assert MODE_LIMITS["fidelity"].forced_line_breaks
    assert not MODE_LIMITS["editable"].forced_line_breaks
    assert str(cfg.models_dir).endswith(str(Path("pdf2pptx") / "models"))
    with pytest.raises(ValueError):
        Config(mode="pretty")


def test_toml_then_overrides(tmp_path: Path) -> None:
    toml = tmp_path / "pdf2pptx.toml"
    toml.write_text('[pdf2pptx]\nmode = "fidelity"\nworkers = 3\nwork_dir = "w"\n',
                    encoding="utf-8")
    cfg = load_config(config_path=toml)
    assert cfg.mode == "fidelity"
    assert cfg.workers == 3
    assert cfg.work_dir == tmp_path / "w"
    cfg = load_config({"workers": 8}, config_path=toml)
    assert cfg.workers == 8 and cfg.mode == "fidelity"
    with pytest.raises(ValueError):
        load_config({"colour": "red"}, config_path=toml)


def test_tier_detection_stub() -> None:
    assert Config(tier="lite").resolved_tier() == "lite"
    assert Config().resolved_tier() in ("lite", "max")
