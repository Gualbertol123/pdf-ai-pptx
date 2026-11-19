"""Local UI (M10): binds 127.0.0.1, refuses foreign hosts / origins, queues uploads,
reports clear errors, lists the cache. Input errors (M10 hardening) have clear messages."""

from __future__ import annotations

import json
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

import pytest

from pdf2pptx import progress
from pdf2pptx.config import load_config
from pdf2pptx.ingest.parse import InputError, open_pdf
from pdf2pptx.ui import server


@pytest.fixture()
def ui(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Any:
    cfg = load_config({"work_dir": str(tmp_path / "work"), "tier": "lite"})
    (tmp_path / "work").mkdir()
    ran: list[str] = []

    def fake_run(self: server.App, job: server.Job) -> None:  # no conversion in this test
        from pdf2pptx.ingest.parse import InputError, open_pdf

        ran.append(job.id)
        job.status = "running"
        try:
            open_pdf(Path(job.pdf)).close()
            job.status, job.summary = "done", {"pages": 1}
        except InputError as exc:
            job.status, job.error = "error", str(exc)

    monkeypatch.setattr(server.App, "_run", fake_run)
    httpd_box: dict[str, Any] = {}
    orig = server.ThreadingHTTPServer

    class Capture(orig):  # type: ignore[misc, valid-type]
        def __init__(self, *a: Any, **k: Any) -> None:
            super().__init__(*a, **k)
            httpd_box["s"] = self

    monkeypatch.setattr(server, "ThreadingHTTPServer", Capture)
    ready = threading.Event()
    t = threading.Thread(target=server.serve, args=(cfg,), kwargs={
        "port": 0, "open_browser": False, "ready": ready}, daemon=True)
    t.start()
    assert ready.wait(10)
    s = httpd_box["s"]
    base = f"http://127.0.0.1:{s.server_address[1]}"
    yield base, ran, cfg
    s.shutdown()


def _req(base: str, path: str, data: bytes | None = None, headers: dict[str, str] | None = None,
         method: str | None = None) -> tuple[int, bytes]:
    r = urllib.request.Request(base + path, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(r, timeout=10) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


POST = {"X-Requested-With": "pdf2pptx", "Content-Type": "application/pdf"}


def test_ui_serves_offline_page_and_guards(ui: Any) -> None:
    base, _, _ = ui
    code, body = _req(base, "/")
    assert code == 200 and b"pdf2pptx" in body
    assert b"https://" not in body and b"http://" not in body.replace(
        b"http://127.0.0.1", b"")  # no CDN / external assets
    assert _req(base, "/api/jobs", headers={"Host": "evil.example"})[0] == 403
    assert _req(base, "/api/jobs?mode=balanced", data=b"%PDF-1.4", method="POST")[0] == 403
    assert _req(base, "/api/jobs?mode=balanced", data=b"%PDF-1.4", method="POST",
                headers=POST | {"Origin": "http://evil.example"})[0] == 403
    assert _req(base, "/api/jobs?mode=nope", data=b"%PDF-1.4", method="POST",
                headers=POST)[0] == 400
    assert _req(base, "/api/jobs/zzz")[0] == 404


def test_ui_upload_job_and_history(ui: Any, tmp_path: Path) -> None:
    base, ran, cfg = ui
    pdf = Path(__file__).resolve().parents[1] / "corpus" / "synthetic" / "d01_title_bullets.pdf"
    if not pdf.is_file():
        pytest.skip("corpus not built")
    code, body = _req(base, "/api/jobs?name=My%20deck.pdf&mode=balanced&tier=lite",
                      data=pdf.read_bytes(), method="POST", headers=POST)
    assert code == 201
    job = json.loads(body)
    code, body = _req(base, "/api/jobs?name=bad.pdf&mode=balanced&tier=lite",
                      data=b"not a pdf at all", method="POST", headers=POST)
    bad = json.loads(body)
    deadline = time.time() + 10
    while time.time() < deadline:
        jobs = {j["id"]: j for j in json.loads(_req(base, "/api/jobs")[1])["jobs"]}
        if jobs[job["id"]]["status"] == "done" and jobs[bad["id"]]["status"] == "error":
            break
        time.sleep(0.1)
    assert jobs[job["id"]]["name"] == "My deck.pdf" and jobs[job["id"]]["status"] == "done"
    assert "not a PDF" in jobs[bad["id"]]["error"]
    assert ran == [job["id"], bad["id"]]  # one at a time, in order
    # persisted: a new App sees the history
    app = server.App(cfg)
    assert set(app.jobs) >= {job["id"], bad["id"]}


def test_input_errors(tmp_path: Path) -> None:
    from reportlab.lib.pdfencrypt import StandardEncryption
    from reportlab.pdfgen import canvas

    enc = tmp_path / "enc.pdf"
    c = canvas.Canvas(str(enc), encrypt=StandardEncryption("pw", canPrint=1))
    c.drawString(100, 700, "secret")
    c.save()
    with pytest.raises(InputError, match="encrypted"):
        open_pdf(enc)
    bad = tmp_path / "bad.pdf"
    bad.write_bytes(b"%PDF-1.7\n" + bytes(range(256)) * 20)
    with pytest.raises(InputError, match=r"damaged|cannot open"):
        open_pdf(bad)
    txt = tmp_path / "t.pdf"
    txt.write_text("hello", encoding="utf-8")
    with pytest.raises(InputError, match="not a PDF"):
        open_pdf(txt)
    with pytest.raises(InputError, match="not found"):
        open_pdf(tmp_path / "missing.pdf")
    empty = tmp_path / "e.pdf"
    empty.write_bytes(b"")
    with pytest.raises(InputError, match="empty"):
        open_pdf(empty)


def test_cli_reports_input_error(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    from pdf2pptx.cli import main

    txt = tmp_path / "t.pdf"
    txt.write_text("hello", encoding="utf-8")
    rc = main(["convert", str(txt), "-o", str(tmp_path / "o.pptx"), "--tier", "lite",
               "--work-dir", str(tmp_path / "w")])
    assert rc == 3 and "not a PDF" in capsys.readouterr().err


def test_progress_callback_is_optional_and_safe() -> None:
    seen: list[tuple[str, int, int]] = []
    progress.report("read", 0, 1)  # no callback: nothing happens
    progress.set_callback(lambda s, d, t, m: seen.append((s, d, t)))
    try:
        progress.report("analyze", 2, 5)
        progress.set_callback(lambda s, d, t, m: 1 / 0)
        progress.report("analyze", 3, 5)  # a failing callback never raises
    finally:
        progress.set_callback(None)
    assert seen == [("analyze", 2, 5)]
