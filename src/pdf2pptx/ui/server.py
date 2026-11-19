"""``pdf2pptx ui``: a local web app (M10). Binds 127.0.0.1 only; nothing leaves the PC.

* ``GET /``: the single page (``index.html``, no external assets).
* ``POST /api/jobs?name=..&mode=..&tier=..[&slides=..&lang=..&audience=..]``: the PDF is
  the raw request body; the job is queued. Jobs run one at a time (one PowerPoint, one
  model) in a worker thread; progress comes from :mod:`pdf2pptx.progress`.
* ``GET /api/jobs``: jobs (persisted in ``<work>/ui/jobs/<id>/job.json``) plus the
  documents in the conversion cache (``<work>/<sha256>/ir_build.json``).
* ``POST /api/cached/<sha256>?mode=..``: convert a cached document again.
* ``GET /api/jobs/<id>``, ``/download``, ``/report``, ``/sources``.

The preview is the M3 debug report (PDF page, PowerPoint render and diff heatmap side by
side) for page conversions, and the rendered slides with their sources for deck mode.

Requests are refused unless the Host header is this server (DNS rebinding) and every
POST carries ``X-Requested-With: pdf2pptx`` and, if the browser sends one, a local
Origin (other sites cannot post files here).
"""

from __future__ import annotations

import dataclasses
import hashlib
import html
import json
import queue
import re
import threading
import time
import traceback
import webbrowser
from dataclasses import dataclass, field
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, urlparse

from .. import __version__, log, progress
from ..config import MODES, TIERS, Config

LOG = log.get(__name__)

MAX_UPLOAD = 2 << 30  # 2 GB
INDEX = Path(__file__).with_name("index.html")
JOB_ID = re.compile(r"^[0-9a-f]{16}$")
SHA = re.compile(r"^[0-9a-f]{64}$")
STAGES = ("read", "ocr", "analyze", "plan", "tune", "write", "check", "preview")


@dataclass
class Job:
    id: str
    name: str
    sha256: str
    pdf: str
    mode: str = "balanced"
    tier: str = "auto"
    options: dict[str, Any] = field(default_factory=dict)
    status: str = "queued"  # queued | running | done | error | interrupted
    stage: str = ""
    done: int = 0
    total: int = 0
    message: str = ""
    created: float = 0.0
    started: float | None = None
    finished: float | None = None
    error: str | None = None
    output: str | None = None
    preview: str | None = None
    summary: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        d = dataclasses.asdict(self)
        d.pop("pdf")
        d["output_name"] = Path(self.output).name if self.output else None
        return d


class App:
    def __init__(self, cfg: Config, root: Path | None = None) -> None:
        self.cfg = cfg
        self.root = root or cfg.work_dir / "ui"
        (self.root / "jobs").mkdir(parents=True, exist_ok=True)
        (self.root / "uploads").mkdir(parents=True, exist_ok=True)
        self.jobs: dict[str, Job] = {}
        self.lock = threading.Lock()
        self.q: queue.Queue[str] = queue.Queue()
        self._load()
        self.worker = threading.Thread(target=self._work, name="pdf2pptx-ui-worker",
                                       daemon=True)
        self.worker.start()

    # -- persistence ---------------------------------------------------------------------

    def _job_dir(self, jid: str) -> Path:
        return self.root / "jobs" / jid

    def _save(self, job: Job) -> None:
        d = self._job_dir(job.id)
        d.mkdir(parents=True, exist_ok=True)
        tmp = d / "job.json.tmp"
        tmp.write_text(json.dumps(dataclasses.asdict(job), indent=1, ensure_ascii=False),
                       encoding="utf-8", newline="\n")
        tmp.replace(d / "job.json")

    def _load(self) -> None:
        for p in sorted((self.root / "jobs").glob("*/job.json")):
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
                job = Job(**{k: v for k, v in data.items()
                             if k in {f.name for f in dataclasses.fields(Job)}})
            except Exception as exc:
                LOG.warning("ui: unreadable job %s: %s", p, exc)
                continue
            if job.status in ("queued", "running"):
                job.status, job.error = "interrupted", "the UI was closed during this job"
            self.jobs[job.id] = job

    # -- jobs ----------------------------------------------------------------------------

    def submit(self, name: str, pdf: Path, sha: str, mode: str, tier: str,
               options: dict[str, Any]) -> Job:
        if mode not in MODES:
            raise ValueError(f"mode must be one of {MODES}")
        if tier not in TIERS:
            raise ValueError(f"tier must be one of {TIERS}")
        now = time.time()
        jid = hashlib.sha256(f"{sha}:{mode}:{tier}:{options}:{now}".encode()).hexdigest()[:16]
        job = Job(id=jid, name=name, sha256=sha, pdf=str(pdf), mode=mode, tier=tier,
                  options=options, created=now)
        with self.lock:
            self.jobs[jid] = job
        self._save(job)
        self.q.put(jid)
        return job

    def _work(self) -> None:
        while True:
            jid = self.q.get()
            job = self.jobs.get(jid)
            if job is None:
                continue
            try:
                self._run(job)
            except Exception as exc:  # the server never dies with a job
                job.status, job.error = "error", f"{type(exc).__name__}: {exc}"
                LOG.error("ui job %s failed: %s", jid, traceback.format_exc())
            finally:
                progress.set_callback(None)
                job.finished = time.time()
                self._save(job)

    def _progress(self, job: Job) -> progress.Callback:
        def cb(stage: str, done: int, total: int, message: str) -> None:
            job.stage, job.done, job.total, job.message = stage, done, total, message
        return cb

    def _run(self, job: Job) -> None:
        from ..ingest.parse import InputError

        job.status, job.started = "running", time.time()
        self._save(job)
        progress.set_callback(self._progress(job))
        o = job.options
        cfg = dataclasses.replace(
            self.cfg, mode=job.mode, tier=job.tier,
            deck_slides=int(o["slides"]) if o.get("slides") else None,
            deck_lang=o.get("lang") or "auto", deck_audience=o.get("audience") or "")
        d = self._job_dir(job.id)
        stem = re.sub(r"[^\w\-. ]+", "_", Path(job.name).stem)[:80] or "converted"
        out = d / f"{stem}.pptx"
        pdf = Path(job.pdf)
        try:
            if job.mode == "deck":
                from ..llm import deck

                res = deck.convert_deck(pdf, out, cfg)
                job.summary = {k: res.get(k) for k in (
                    "pages", "slides", "target_slides", "plan_source", "lang", "dropped",
                    "flagged", "seconds")} | {"audit_ok": res["audit"]["ok"]}
                job.stage, job.message = "preview", "rendering the slides"
                job.preview = str(self._deck_preview(job, out, Path(res["sources"])))
            else:
                from .. import pipeline
                from ..report import debug_html

                r = pipeline.convert(pdf, out, cfg)
                job.summary = {"pages": r.pages, "slides": r.slides, "seconds": r.seconds,
                               "seconds_per_page": r.seconds_per_page, "triage": r.triage,
                               "fallback_pages": r.fallback_pages,
                               "max_tier": bool(r.llm and r.llm.get("applied"))}
                job.output = str(out)
                job.stage, job.message = "check", "rendering in PowerPoint for the preview"
                path, rep = debug_html.write_report(pdf, cfg, d / "report.html", pptx=out)
                s = rep["score"].get("summary", {})
                job.summary["scores"] = {k: (s.get(k) or {}).get("median") for k in (
                    "ssim", "text_sim", "composite")}
                job.summary["seconds_with_preview"] = round(time.time() - (job.started or 0),
                                                            1)
                job.preview = str(path)
            job.output = str(out)
            job.status = "done"
        except InputError as exc:
            job.status, job.error = "error", str(exc)

    def _deck_preview(self, job: Job, pptx: Path, sources: Path) -> Path:
        """Rendered slides (PowerPoint, when available) next to their bullets' sources."""
        import base64

        srcs = json.loads(sources.read_text(encoding="utf-8"))
        imgs: list[str] = []
        why = ""
        try:
            from ..verify import ppt_com

            ppt = ppt_com.shared()
            with ppt.opened(pptx) as pres:
                paths = ppt.render(pres, None, 72, self._job_dir(job.id) / "slides")
            imgs = ["data:image/png;base64," + base64.b64encode(p.read_bytes()).decode()
                    for p in paths]
        except Exception as exc:
            why = f"PowerPoint unavailable, no slide pictures ({type(exc).__name__}: {exc})"
        audit = srcs.get("audit", {})
        e = html.escape
        rows = []
        for k, s in enumerate(srcs["slides"]):
            img = f'<img src="{imgs[k]}" alt="slide {k + 1}">' if k < len(imgs) else ""
            items = "".join(
                f"<li>{e(b['text'])} <span class=src>p. {', '.join(map(str, b['pages']))} · "
                f"{e(', '.join(b['sources']))}</span>"
                + (f" <span class=flag>{e('; '.join(b['flags']))}</span>" if b.get("flags")
                   else "") + "</li>" for b in s.get("bullets", []))
            vis = s.get("visual")
            vis_txt = (f"<div class=src>visual: {e(vis['kind'])} {e(vis['id'])} "
                       f"(p. {vis['page']})</div>" if vis else "")
            rows.append(f"<section><div class=pic>{img}</div><div><h3>{k + 1}. "
                        f"{e(s.get('title', ''))}</h3><ul>{items}</ul>{vis_txt}</div></section>")
        dropped = "".join(f"<li>{e(x['where'])}: {e(x['text'])} <span class=flag>"
                          f"{e(x['reason'])}</span></li>" for x in srcs.get("dropped", []))
        checks = (f"bullets {audit.get('bullets')} · with a valid source "
                  f"{audit.get('bullets_with_valid_source')} · numbers absent from the source "
                  f"{len(audit.get('numbers_absent', []))} · overflow "
                  f"{e(str(audit.get('overflow')))} · slides {audit.get('slides')} "
                  f"(target {audit.get('target')})")
        page = f"""<!doctype html><html lang=en><head><meta charset=utf-8>
<title>{e(job.name)}: deck</title><style>
body{{font:14px/1.45 system-ui,Segoe UI,sans-serif;margin:0;padding:16px;background:#f6f7f9;
color:#1d2433}}
section{{display:grid;grid-template-columns:minmax(0,480px) 1fr;gap:16px;background:#fff;
border:1px solid #dde1e8;border-radius:8px;padding:12px;margin:0 0 12px}}
.pic img{{width:100%;border:1px solid #dde1e8}} h3{{margin:0 0 6px;font-size:15px}}
ul{{margin:0;padding-left:18px}} .src{{color:#5b6475;font-size:12px}}
.flag{{color:#a15c00;font-size:12px}} .bar{{margin:0 0 12px;color:#5b6475}}
@media (max-width:700px){{section{{grid-template-columns:1fr}}}}
@media (prefers-color-scheme:dark){{body{{background:#14171c;color:#e3e7ee}}
section{{background:#1c2027;border-color:#2c323c}} .src{{color:#9aa4b5}}}}
</style></head><body><div class=bar>{e(srcs.get('plan_source', ''))} · {checks}
{(' · ' + e(why)) if why else ''}</div>{''.join(rows)}
{('<h3>Left out (not grounded in the source)</h3><ul>' + dropped + '</ul>') if dropped else ''}
</body></html>"""
        out = self._job_dir(job.id) / "preview.html"
        out.write_text(page, encoding="utf-8", newline="\n")
        return out

    # -- history -------------------------------------------------------------------------

    def cached(self) -> list[dict[str, Any]]:
        """Documents in the conversion cache (newest first): what ``ir_build.json`` says
        about its source, read from the head of the file only."""
        out = []
        for d in self.cfg.work_dir.iterdir() if self.cfg.work_dir.is_dir() else []:
            ir = d / "ir_build.json"
            if not SHA.match(d.name) or not ir.is_file():
                continue
            try:
                with ir.open("r", encoding="utf-8") as fh:
                    head = fh.read(4096)
                m = re.search(r'"source_path":\s*"((?:[^"\\]|\\.)*)"', head)
                src = json.loads(f'"{m.group(1)}"') if m else ""
            except (OSError, ValueError):
                continue
            out.append({"sha256": d.name, "name": Path(src).name or d.name[:12],
                        "path": src, "available": bool(src) and Path(src).is_file(),
                        "modified": ir.stat().st_mtime})
        out.sort(key=lambda r: -r["modified"])
        return out[:200]

    def listing(self) -> dict[str, Any]:
        with self.lock:
            jobs = sorted(self.jobs.values(), key=lambda j: -j.created)
        return {"version": __version__, "jobs": [j.as_dict() for j in jobs],
                "cached": self.cached()}


# --------------------------------------------------------------------------- HTTP


class Handler(BaseHTTPRequestHandler):
    server_version = f"pdf2pptx/{__version__}"
    app: App
    port: int

    def log_message(self, format: str, *args: Any) -> None:
        LOG.debug("ui %s - %s", self.address_string(), format % args)

    # -- helpers -------------------------------------------------------------------------

    def _host_ok(self) -> bool:
        host = (self.headers.get("Host") or "").lower()
        return host in (f"127.0.0.1:{self.port}", f"localhost:{self.port}")

    def _post_ok(self) -> bool:
        if self.headers.get("X-Requested-With") != "pdf2pptx":
            return False
        origin = self.headers.get("Origin")
        return origin in (None, f"http://127.0.0.1:{self.port}", f"http://localhost:{self.port}")

    def _send(self, code: int, body: bytes, ctype: str, extra: dict[str, str] | None = None
              ) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        if ctype.startswith("text/html"):
            # inline scripts/styles and data: images only; nothing from the network
            self.send_header("Content-Security-Policy",
                             "default-src 'self'; img-src 'self' data:; style-src 'self' "
                             "'unsafe-inline'; script-src 'self' 'unsafe-inline'; "
                             "connect-src 'self'; frame-src 'self'; frame-ancestors 'self'")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, obj: Any, code: int = 200) -> None:
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _error(self, code: int, msg: str) -> None:
        self._json({"error": msg}, code)

    def _file(self, path: Path, ctype: str, download: str | None = None) -> None:
        if not path.is_file():
            self._error(404, "not found")
            return
        extra = {}
        if download:
            extra["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(download)}"
        self._send(200, path.read_bytes(), ctype, extra)

    def _job(self, jid: str) -> Job | None:
        if not JOB_ID.match(jid):
            return None
        return self.app.jobs.get(jid)

    # -- routes --------------------------------------------------------------------------

    def do_GET(self) -> None:
        if not self._host_ok():
            self._error(403, "forbidden host")
            return
        url = urlparse(self.path)
        parts = [p for p in url.path.split("/") if p]
        if not parts:
            self._send(200, INDEX.read_bytes(), "text/html; charset=utf-8")
            return
        if parts == ["api", "jobs"]:
            self._json(self.app.listing())
            return
        if len(parts) >= 3 and parts[:2] == ["api", "jobs"]:
            job = self._job(parts[2])
            if job is None:
                self._error(404, "no such job")
                return
            if len(parts) == 3:
                self._json(job.as_dict())
                return
            what = parts[3]
            if what == "download" and job.output:
                self._file(Path(job.output),
                           "application/vnd.openxmlformats-officedocument.presentationml."
                           "presentation", download=Path(job.output).name)
                return
            if what == "report" and job.preview:
                self._file(Path(job.preview), "text/html; charset=utf-8")
                return
            if what == "sources" and job.output:
                self._file(Path(job.output).with_name(Path(job.output).stem
                                                      + ".deck_sources.json"),
                           "application/json; charset=utf-8")
                return
        self._error(404, "not found")

    def do_HEAD(self) -> None:
        self.do_GET()

    def do_POST(self) -> None:
        if not self._host_ok() or not self._post_ok():
            self._error(403, "forbidden")
            return
        url = urlparse(self.path)
        parts = [p for p in url.path.split("/") if p]
        qs = {k: v[0] for k, v in parse_qs(url.query).items()}
        opts = {k: qs[k] for k in ("slides", "lang", "audience") if qs.get(k)}
        mode, tier = qs.get("mode", "balanced"), qs.get("tier", "auto")
        try:
            if parts == ["api", "jobs"]:
                self._upload(qs.get("name") or "document.pdf", mode, tier, opts)
                return
            if len(parts) == 3 and parts[:2] == ["api", "cached"] and SHA.match(parts[2]):
                hit = next((c for c in self.app.cached() if c["sha256"] == parts[2]), None)
                up = self.app.root / "uploads" / f"{parts[2]}.pdf"
                src = up if up.is_file() else Path(hit["path"]) if hit and hit["available"] \
                    else None
                if src is None or hit is None:
                    self._error(404, "the source PDF of this cache entry is gone")
                    return
                job = self.app.submit(hit["name"], src, parts[2], mode, tier, opts)
                self._json(job.as_dict(), 201)
                return
        except ValueError as exc:
            self._error(400, str(exc))
            return
        self._error(404, "not found")

    def _upload(self, name: str, mode: str, tier: str, opts: dict[str, Any]) -> None:
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            self._error(411, "send the PDF as the request body")
            return
        if n > MAX_UPLOAD:
            self._error(413, "file too large (2 GB at most)")
            return
        up = self.app.root / "uploads"
        tmp = up / f".upload_{threading.get_ident()}.part"
        h = hashlib.sha256()
        left = n
        with tmp.open("wb") as fh:
            while left:
                chunk = self.rfile.read(min(left, 1 << 20))
                if not chunk:
                    break
                fh.write(chunk)
                h.update(chunk)
                left -= len(chunk)
        if left:
            tmp.unlink(missing_ok=True)
            self._error(400, "upload interrupted")
            return
        sha = h.hexdigest()
        dest = up / f"{sha}.pdf"
        tmp.replace(dest)
        job = self.app.submit(Path(name).name, dest, sha, mode, tier, opts)
        self._json(job.as_dict(), 201)


def serve(cfg: Config, port: int = 8765, open_browser: bool = True,
          ready: threading.Event | None = None) -> None:
    """Run the UI until interrupted (Ctrl+C)."""
    app = App(cfg)
    handler = type("BoundHandler", (Handler,), {"app": app})
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", port), handler)
    except OSError as exc:  # port taken (or reserved): any free port instead
        LOG.warning("port %d unavailable (%s); using a free port", port, exc)
        httpd = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    handler.port = httpd.server_address[1]
    url = f"http://127.0.0.1:{handler.port}/"
    LOG.info("pdf2pptx UI on %s (local only; Ctrl+C to stop)", url)
    print(f"pdf2pptx UI: {url}  (Ctrl+C to stop)", flush=True)
    if open_browser:
        webbrowser.open(url)
    if ready is not None:
        ready.set()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
