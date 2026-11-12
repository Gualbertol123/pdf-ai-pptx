from __future__ import annotations

import pytest

from pdf2pptx import cli, doctor
from pdf2pptx.config import Config


def test_parser_has_all_subcommands() -> None:
    p = cli.build_parser()
    for cmd in ("convert", "score", "corpus", "report", "doctor", "setup"):
        args = p.parse_args([cmd] if cmd not in ("corpus",) else [cmd, "list"])
        assert args.command == cmd


def test_stubs_exit_cleanly(capsys: pytest.CaptureFixture[str]) -> None:
    assert cli.main(["convert", "x.pdf"]) == 3
    assert "Milestone 2" in capsys.readouterr().err


def test_doctor_parts_without_powerpoint() -> None:
    assert doctor.check_ram()["total_gb"] > 0
    assert "vulkan-1.dll" in doctor.check_vulkan()
    assert "installed" in doctor.check_onnxruntime()
    gpus = doctor.check_gpu().get("gpus", [])
    assert isinstance(gpus, list)


@pytest.mark.powerpoint
def test_doctor_full_report() -> None:
    r = doctor.run(Config())
    assert r["powerpoint_com"]["available"]
    text = doctor.format_report(r)
    for key in ("Python", "PowerPoint COM", "onnxruntime", "Vulkan", "RAM", "GPU"):
        assert key in text
