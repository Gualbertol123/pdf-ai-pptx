from __future__ import annotations

from pathlib import Path

import pytest

from pdf2pptx import cli, doctor
from pdf2pptx.config import Config


def test_parser_has_all_subcommands() -> None:
    p = cli.build_parser()
    for cmd in ("convert", "score", "corpus", "report", "doctor", "setup"):
        args = p.parse_args([cmd] if cmd not in ("corpus",) else [cmd, "list"])
        assert args.command == cmd


def test_stubs_exit_cleanly(capsys: pytest.CaptureFixture[str]) -> None:
    assert cli.main(["setup", "models"]) == 3
    assert "Milestone 8" in capsys.readouterr().err


def test_report_without_powerpoint(tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
                                   capsys: pytest.CaptureFixture[str]) -> None:
    """PowerPoint missing: the report still renders, scores needing it are null, and
    the page says why."""
    from reportlab.pdfgen import canvas

    from pdf2pptx.verify import ppt_com

    def unavailable(**_: object) -> ppt_com.PowerPointEngine:
        raise ppt_com.PowerPointUnavailable("simulated: PowerPoint not installed")

    monkeypatch.setattr(ppt_com, "shared", unavailable)
    pdf = tmp_path / "doc.pdf"
    c = canvas.Canvas(str(pdf), pagesize=(400, 300))
    c.setFont("Helvetica", 14)
    c.drawString(40, 240, "Hello report")
    c.rect(40, 60, 120, 80, fill=1)
    c.save()
    out = tmp_path / "r.html"
    assert cli.main(["report", str(pdf), "-o", str(out), "--work-dir", str(tmp_path / "w")]) == 0
    html = out.read_text(encoding="utf-8")
    assert "PowerPoint unavailable" in html and "simulated: PowerPoint not installed" in html
    assert "data:image/jpeg;base64," in html  # the PDF render is inlined
    assert 'data-kind="text_block"' in html and 'data-kind="shape"' in html
    assert "http://" not in html and "https://" not in html  # self-contained
    assert "src=\"http" not in html
    assert "PowerPoint unavailable" in capsys.readouterr().out


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
