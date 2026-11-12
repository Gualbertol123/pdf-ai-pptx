"""PowerPoint COM engine: the ONLY module that talks to PowerPoint (PLAN.md §8 rule 6).

PowerPoint runs behind a worker *process* (``multiprocessing`` spawn context) whose
main thread is a single COM STA thread. The client (:class:`PowerPointEngine`) sends
requests over a queue and waits with a per-call timeout, watching a heartbeat the
worker updates between COM operations. When the worker dies, hangs, or loses its
PowerPoint (RPC disconnect), the client kills the worker and only the POWERPNT.EXE
processes that worker started, starts a new worker, re-opens the presentations that
were open, and retries the call.

Operations: ``open`` / ``render`` / ``text_layout`` / ``export_pdf`` / ``close`` /
``ping``. Slides are addressed by 0-based index (= IR page index).

Safety rules:

* PowerPoint is started with ``DispatchEx``, presentations open read-only with
  ``WithWindow=False``, ``DisplayAlerts`` is off.
* A POWERPNT.EXE counts as ours only if it appeared while our worker created the
  application object. Ours are recorded in a registry (``%TEMP%/pdf2pptx/ppt_owned``)
  with their creation time, so a later run can reap them if this process dies.
* PowerPoint is never quit while it holds presentations that are not ours, and a
  PowerPoint the user started is never quit or killed.
"""

from __future__ import annotations

import atexit
import json
import multiprocessing as mp
import multiprocessing.util  # registers its atexit hook before ours (see the end)
import os
import queue
import tempfile
import threading
import time
from collections.abc import Iterator, Sequence
from contextlib import contextmanager, suppress
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .. import log

LOG = log.get(__name__)

PP_SAVE_AS_PDF = 32
PP_ALERTS_NONE = 1
MSO_TRUE, MSO_FALSE = -1, 0
MSO_GROUP = 6
PARAGRAPH_ENDS = (chr(13), chr(10), chr(11))  # CR, LF, vertical tab (line break)

REGISTRY_DIR = Path(tempfile.gettempdir()) / "pdf2pptx" / "ppt_owned"


class PowerPointUnavailable(RuntimeError):
    """PowerPoint cannot be started or keeps failing; callers degrade (scores = null)."""


class PowerPointError(RuntimeError):
    """A single call failed (bad file, bad slide index) but PowerPoint itself is fine."""


# --------------------------------------------------------------------------- result types


@dataclass
class WordBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass
class LineBox:
    """One rendered line. ``bound`` is COM's raw (BoundLeft, BoundTop, BoundWidth,
    BoundHeight), slide points. ``x0..y1`` is the same box with the bottom moved up
    by the paragraph's space-after (so it sits on the descent line like a PDF font
    bbox bottom); words are trimmed of trailing whitespace and share the line's y."""

    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    words: list[WordBox] = field(default_factory=list)
    shape: str = ""
    bound: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)


@dataclass
class TextShape:
    """Text layout of one text-bearing shape (a table cell counts as one shape)."""

    name: str
    shape_id: int
    lines: list[LineBox] = field(default_factory=list)


@dataclass
class PresInfo:
    handle: str
    path: str
    slides: int
    width_pt: float
    height_pt: float


def layout_lines(shapes: Sequence[TextShape]) -> list[LineBox]:
    return [ln for s in shapes for ln in s.lines]


# --------------------------------------------------------------------------- process helpers


def _ppt_procs() -> dict[int, float]:
    """pid -> create_time of every running POWERPNT.EXE."""
    import psutil

    out: dict[int, float] = {}
    for p in psutil.process_iter(["name", "create_time"]):
        try:
            if (p.info["name"] or "").lower() == "powerpnt.exe":
                out[p.pid] = float(p.info["create_time"] or 0.0)
        except Exception:
            continue
    return out


def _alive(pid: int, ctime: float) -> bool:
    import psutil

    try:
        p = psutil.Process(pid)
        return p.name().lower() == "powerpnt.exe" and abs(p.create_time() - ctime) < 1.0
    except Exception:
        return False


def _kill(pid: int, ctime: float) -> None:
    """Kill pid only if it is still the same POWERPNT.EXE (guards against pid reuse)."""
    import psutil

    if not _alive(pid, ctime):
        return
    with suppress(Exception):
        psutil.Process(pid).kill()
        LOG.info("killed POWERPNT.EXE pid %d (started by pdf2pptx)", pid)


def _reg_file(pid: int, ctime: float, owner: int) -> Path:
    return REGISTRY_DIR / f"ppt_{pid}_{int(ctime * 1000)}_{owner}.json"


def _register(pid: int, ctime: float, owner: int) -> None:
    with suppress(OSError):
        REGISTRY_DIR.mkdir(parents=True, exist_ok=True)
        _reg_file(pid, ctime, owner).write_text(
            json.dumps({"pid": pid, "ctime": ctime, "owner": owner}), encoding="utf-8")


def _unregister(pid: int, ctime: float, owner: int) -> None:
    with suppress(OSError):
        _reg_file(pid, ctime, owner).unlink()


def registered() -> list[dict[str, Any]]:
    """Registry entries: one per (POWERPNT.EXE started by pdf2pptx, worker using it).

    PowerPoint is single-instance, so a second engine attaches to the first one's
    process; it then registers as a co-owner. An instance is quit or killed only
    when no live co-owner is left.
    """
    out = []
    for f in sorted(REGISTRY_DIR.glob("ppt_*.json")) if REGISTRY_DIR.is_dir() else []:
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            out.append({"pid": int(d["pid"]), "ctime": float(d["ctime"]),
                        "owner": int(d["owner"])})
        except (OSError, ValueError, KeyError):
            with suppress(OSError):
                f.unlink()
    return out


def owned_alive() -> list[int]:
    """Pids of POWERPNT.EXE processes started by pdf2pptx that are still running."""
    return sorted({r["pid"] for r in registered() if _alive(r["pid"], r["ctime"])})


def _pid_running(pid: int) -> bool:
    import psutil

    return psutil.pid_exists(pid)


def _live_co_owners(pid: int, ctime: float, exclude: int) -> list[int]:
    return [r["owner"] for r in registered()
            if r["pid"] == pid and abs(r["ctime"] - ctime) < 1e-3 and r["owner"] != exclude
            and _pid_running(r["owner"])]


def _release(pid: int, ctime: float, owner: int, wait_s: float) -> None:
    """``owner`` stops using an instance; kill it if nobody else uses it and it does
    not exit by itself within ``wait_s``."""
    _unregister(pid, ctime, owner)
    if _live_co_owners(pid, ctime, owner):
        return
    deadline = time.time() + wait_s
    while time.time() < deadline and _alive(pid, ctime):
        time.sleep(0.2)
    _kill(pid, ctime)
    for r in registered():  # entries of dead co-owners
        if r["pid"] == pid and abs(r["ctime"] - ctime) < 1e-3:
            _unregister(r["pid"], r["ctime"], r["owner"])


def reap_orphans(wait_s: float = 8.0) -> list[int]:
    """Kill registered PowerPoints none of whose workers is alive; forget dead entries.

    An orphan first gets ``wait_s`` to finish quitting on its own (a worker that
    just exited normally may still be shutting PowerPoint down).
    """
    groups: dict[tuple[int, float], list[dict[str, Any]]] = {}
    for r in registered():
        groups.setdefault((r["pid"], r["ctime"]), []).append(r)
    pending = []
    for (pid, ct), rs in groups.items():
        if not _alive(pid, ct):
            for r in rs:
                _unregister(pid, ct, r["owner"])
        elif not any(_pid_running(r["owner"]) for r in rs):
            pending.append((pid, ct, rs))
    deadline = time.time() + wait_s
    while pending and time.time() < deadline and any(_alive(p, c) for p, c, _ in pending):
        time.sleep(0.25)
    killed = []
    for pid, ct, rs in pending:
        if _alive(pid, ct):
            _kill(pid, ct)
            killed.append(pid)
        for r in rs:
            _unregister(pid, ct, r["owner"])
    return killed


# --------------------------------------------------------------------------- worker side


class _Worker:
    """Runs in the worker process, on its main (STA) thread."""

    def __init__(self, heartbeat: Any) -> None:
        self.heartbeat = heartbeat
        self.app: Any = None
        self.owned: dict[int, float] = {}
        self.pres: dict[str, Any] = {}  # handle -> Presentation
        self.paths: dict[str, Path] = {}
        self.plans: dict[str, list[list[_PlanItem]]] = {}
        self.counts: dict[str, int] = {}

    def beat(self) -> None:
        self.heartbeat.value = time.time()

    # ---------------------------------------------------------------- lifecycle

    def start(self, adopt_paths: list[str]) -> str:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()  # this thread becomes the process's single STA
        before = _ppt_procs()
        ours_before = {(r["pid"], r["ctime"]) for r in registered()}
        app = win32com.client.DispatchEx("PowerPoint.Application")
        self.owned = {pid: ct for pid, ct in _ppt_procs().items() if pid not in before}
        if not self.owned and before and all(
                any(p == q and abs(c - d) < 1e-3 for q, d in ours_before)
                for p, c in before.items()):
            # attached to a PowerPoint another pdf2pptx worker started: co-own it
            self.owned = dict(before)
        for pid, ct in self.owned.items():
            _register(pid, ct, os.getpid())
        self.app = app
        self.app.DisplayAlerts = PP_ALERTS_NONE
        version = str(self.app.Version)
        # A previous worker of ours died inside a PowerPoint it did not own: close
        # the presentations it left open there (never anything else).
        if adopt_paths:
            wanted = {os.path.normcase(p) for p in adopt_paths}
            for i in range(int(self.app.Presentations.Count), 0, -1):
                with suppress(Exception):
                    p = self.app.Presentations.Item(i)
                    if os.path.normcase(str(p.FullName)) in wanted:
                        p.Close()
        return version

    def healthy(self) -> bool:
        try:
            _ = self.app.Version
            return True
        except Exception:
            return False

    def shutdown(self) -> None:
        for h in list(self.pres):
            with suppress(Exception):
                self.pres.pop(h).Close()
        me = os.getpid()
        shared = any(_live_co_owners(p, c, me) for p, c in self.owned.items())
        if self.app is not None and self.owned and not shared:
            foreign = 0
            with suppress(Exception):
                foreign = int(self.app.Presentations.Count)
            if foreign == 0:
                with suppress(Exception):
                    self.app.Quit()
            else:  # the user opened something in our instance: it is theirs now
                LOG.warning("PowerPoint has %d presentations not opened by pdf2pptx; "
                            "leaving it running", foreign)
                for pid, ct in self.owned.items():
                    _unregister(pid, ct, me)
                self.owned = {}
        self.app = None
        with suppress(Exception):
            import pythoncom

            pythoncom.CoUninitialize()
        for pid, ct in self.owned.items():
            _release(pid, ct, me, wait_s=8.0)

    # ---------------------------------------------------------------- operations

    def op_ping(self) -> str:
        return str(self.app.Version)

    def op_open(self, handle: str, path: str) -> PresInfo:
        full = os.path.normpath(str(Path(path).resolve()))
        if handle in self.pres:
            with suppress(Exception):
                self.pres.pop(handle).Close()
        # ReadOnly=msoTrue, Untitled=msoFalse, WithWindow=msoFalse
        pres = self.app.Presentations.Open(full, MSO_TRUE, MSO_FALSE, MSO_FALSE)
        self.pres[handle] = pres
        self.paths[handle] = Path(full)
        self.plans.pop(handle, None)
        self.counts[handle] = int(pres.Slides.Count)
        return PresInfo(handle, full, self.counts[handle],
                        float(pres.PageSetup.SlideWidth), float(pres.PageSetup.SlideHeight))

    def op_close(self, handle: str) -> None:
        pres = self.pres.pop(handle, None)
        self.paths.pop(handle, None)
        self.plans.pop(handle, None)
        self.counts.pop(handle, None)
        if pres is not None:
            pres.Close()

    def _get(self, handle: str) -> Any:
        try:
            return self.pres[handle]
        except KeyError:
            raise PowerPointError(f"presentation {handle!r} is not open") from None

    def _slide(self, handle: str, index: int) -> Any:
        pres = self._get(handle)
        n = self.counts[handle]
        if not 0 <= index < n:
            raise PowerPointError(f"slide index {index} out of range (0..{n - 1})")
        return pres.Slides.Item(index + 1)

    def op_render(self, handle: str, slides: list[int] | None, dpi: float,
                  out_dir: str) -> list[str]:
        pres = self._get(handle)
        w_pt, h_pt = float(pres.PageSetup.SlideWidth), float(pres.PageSetup.SlideHeight)
        w_px, h_px = round(w_pt * dpi / 72), round(h_pt * dpi / 72)
        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)
        idx = list(range(self.counts[handle])) if slides is None else list(slides)
        paths = []
        for i in idx:
            p = out / f"slide_{i + 1:03d}.png"
            self._slide(handle, i).Export(os.path.normpath(str(p.resolve())), "PNG", w_px, h_px)
            paths.append(str(p))
            self.beat()
        return paths

    def op_text_layout(self, handle: str, slide: int) -> list[TextShape]:
        self._slide(handle, slide)  # validates handle and index
        if handle not in self.plans:
            self.plans[handle] = _text_plan(self.paths[handle])
        plan = self.plans[handle]
        items = plan[slide] if slide < len(plan) else []
        out: list[TextShape] = []
        if not items:
            return out
        shapes = _RAW(self._slide(handle, slide)._oleobj_, "Slide", "Shapes")
        for item in items:
            shape = _RAW(shapes, "Shapes", "Item", item.index)
            if item.kind == "text":
                _collect_text_frame(_RAW(shape, "Shape", "TextFrame2"), item.name,
                                    item.shape_id, out)
            else:  # table / group
                _collect_shape(shape, out, item.spanned)
            self.beat()
        return out

    def op_export_pdf(self, handle: str, out: str) -> str:
        pres = self._get(handle)
        p = Path(out).resolve()
        p.parent.mkdir(parents=True, exist_ok=True)
        pres.SaveAs(os.path.normpath(str(p)), PP_SAVE_AS_PDF)
        return str(p)

    def op_set_geometry(self, handle: str, slide: int,
                        edits: list[tuple[str, float, float, float]]) -> int:
        """Set (left, top, width) of named shapes in the open (in-memory) presentation.

        Only used to measure the COM-edit alternative to rewriting the file (M5)."""
        shapes = _RAW(self._slide(handle, slide)._oleobj_, "Slide", "Shapes")
        n = 0
        for name, left, top, width in edits:
            shp = _RAW(shapes, "Shapes", "Item", name)
            shp.Invoke(_dispid(shp, "Left"), 0, 4, False, float(left))  # DISPATCH_PROPERTYPUT
            shp.Invoke(_dispid(shp, "Top"), 0, 4, False, float(top))
            shp.Invoke(_dispid(shp, "Width"), 0, 4, False, float(width))
            n += 1
        return n

    def op_crash(self) -> None:  # test hook: simulate a worker that dies mid-call
        os._exit(17)

    def op_hang(self, seconds: float) -> None:  # test hook: simulate a hung COM call
        time.sleep(seconds)


class _RawDispatch:
    """Late-bound COM calls with DISPIDs cached per (interface, member).

    Profiling (M3) showed the cost of text layout is not PowerPoint's work but
    pywin32's wrapping of every new range object: a type-info lookup (dynamic
    dispatch) or a cross-process QueryInterface (makepy classes), 2-10 ms each.
    Raw ``IDispatch.Invoke`` is one round trip per member (~0.25 ms). Callers name
    the interface of the object, since DISPIDs are per interface.
    """

    def __init__(self) -> None:
        self.ids: dict[tuple[str, str], int] = {}

    def __call__(self, obj: Any, iface: str, name: str, *args: Any) -> Any:
        import pythoncom

        key = (iface, name)
        dispid = self.ids.get(key)
        if dispid is None:
            dispid = self.ids[key] = int(obj.GetIDsOfNames(name))
        flags = pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD
        return obj.Invoke(dispid, 0, flags, True, *args)


_RAW = _RawDispatch()


def _dispid(obj: Any, name: str) -> int:
    key = ("Shape", name)
    if key not in _RAW.ids:
        _RAW.ids[key] = int(obj.GetIDsOfNames(name))
    return _RAW.ids[key]


def _collect_text_frame(tf2: Any, name: str, sid: int, out: list[TextShape],
                        offset: tuple[float, float] = (0.0, 0.0)) -> None:
    """Lines and words of one TextFrame2 (raw IDispatch)."""
    ox, oy = offset
    raw = _RAW
    ts = TextShape(name=name, shape_id=sid)
    lines = raw(raw(tf2, "TextFrame2", "TextRange"), "TextRange2", "Lines")
    n_lines = int(raw(lines, "TextRange2", "Count"))
    for li in range(1, n_lines + 1):
        line = raw(lines, "TextRange2", "Item", li)
        text = str(raw(line, "TextRange2", "Text"))
        if not text.strip():
            continue
        left = float(raw(line, "TextRange2", "BoundLeft"))
        top = float(raw(line, "TextRange2", "BoundTop"))
        width = float(raw(line, "TextRange2", "BoundWidth"))
        height = float(raw(line, "TextRange2", "BoundHeight"))
        y0 = top + oy
        y1 = y0 + height
        # space-after belongs to the paragraph's last line only; a line that ends in a
        # soft break (VT) is not a paragraph end
        if text.endswith(PARAGRAPH_ENDS[:2]) or li == n_lines:
            pf = raw(line, "TextRange2", "ParagraphFormat")
            after = max(0.0, float(raw(pf, "ParagraphFormat2", "SpaceAfter")))
            if after and int(raw(pf, "ParagraphFormat2", "LineRuleAfter")):
                # SpaceAfter in lines: one line = 1.2 x the font size (measured, M4)
                size = float(raw(raw(line, "TextRange2", "Font"), "Font2", "Size"))
                after = after * 1.2 * size if size > 0 else 0.0
            y1 -= after
        lb = LineBox(text=text, x0=left + ox, y0=y0, x1=left + width + ox, y1=y1, shape=name,
                     bound=(left + ox, top + oy, width, height))
        words = raw(line, "TextRange2", "Words")
        for wi in range(1, int(raw(words, "TextRange2", "Count")) + 1):
            w = raw(words, "TextRange2", "Item", wi)
            wraw = str(raw(w, "TextRange2", "Text"))
            wt = wraw.strip()
            if not wt:
                continue
            if wraw != wraw.rstrip():  # a new range object costs a round trip: only if needed
                w = raw(w, "TextRange2", "TrimText")
            wl = float(raw(w, "TextRange2", "BoundLeft"))
            ww = float(raw(w, "TextRange2", "BoundWidth"))
            lb.words.append(WordBox(wt, wl + ox, y0, wl + ww + ox, y1))
        ts.lines.append(lb)
    out.append(ts)


def _collect_shape(shape: Any, out: list[TextShape],
                   spanned: dict[int, set[tuple[int, int]]]) -> None:
    """Generic walk of one shape (raw IDispatch): groups recurse, tables per cell."""
    raw = _RAW
    name = str(raw(shape, "Shape", "Name"))
    if int(raw(shape, "Shape", "Type")) == MSO_GROUP:
        items = raw(shape, "Shape", "GroupItems")
        for k in range(1, int(raw(items, "GroupShapes", "Count")) + 1):
            _collect_shape(raw(items, "GroupShapes", "Item", k), out, spanned)
        return
    sid = int(raw(shape, "Shape", "Id"))
    if raw(shape, "Shape", "HasTable"):
        # Measured (M1): table cell text bounds are relative to the cell; x already
        # includes the left margin, y does not include the top margin.
        tbl = raw(shape, "Shape", "Table")
        rows, cols = raw(tbl, "Table", "Rows"), raw(tbl, "Table", "Columns")
        n_rows, n_cols = int(raw(rows, "Rows", "Count")), int(raw(cols, "Columns", "Count"))
        widths = [float(raw(raw(cols, "Columns", "Item", c), "Column", "Width"))
                  for c in range(1, n_cols + 1)]
        heights = [float(raw(raw(rows, "Rows", "Item", r), "Row", "Height"))
                   for r in range(1, n_rows + 1)]
        skip = spanned.get(sid, set())
        left, top = float(raw(shape, "Shape", "Left")), float(raw(shape, "Shape", "Top"))
        for r in range(1, n_rows + 1):
            for c in range(1, n_cols + 1):
                if (r - 1, c - 1) in skip:
                    continue
                cell_shape = raw(raw(tbl, "Table", "Cell", r, c), "Cell", "Shape")
                if not raw(cell_shape, "Shape", "HasTextFrame"):
                    continue
                tf2 = raw(cell_shape, "Shape", "TextFrame2")
                if raw(tf2, "TextFrame2", "HasText"):
                    margin_top = float(raw(tf2, "TextFrame2", "MarginTop"))
                    off = (left + sum(widths[: c - 1]), top + sum(heights[: r - 1]) + margin_top)
                    _collect_text_frame(tf2, f"{name}[{r},{c}]", sid, out, off)
        return
    if raw(shape, "Shape", "HasTextFrame"):
        tf2 = raw(shape, "Shape", "TextFrame2")
        if raw(tf2, "TextFrame2", "HasText"):
            _collect_text_frame(tf2, name, sid, out)


@dataclass
class _PlanItem:
    index: int  # 1-based position in Slide.Shapes (z-order, as python-pptx iterates)
    kind: str  # text | table | group
    shape_id: int
    name: str
    spanned: dict[int, set[tuple[int, int]]] = field(default_factory=dict)


def _has_text(tf: Any) -> bool:
    return any(r.text.strip() for p in tf.paragraphs for r in p.runs) or bool(
        tf.text.strip())


def _group_has_text(grp: Any) -> bool:
    for shp in grp.shapes:
        if shp.shape_type == 6 and _group_has_text(shp):  # MSO_SHAPE_TYPE.GROUP
            return True
        if getattr(shp, "has_text_frame", False) and shp.has_text_frame and _has_text(
                shp.text_frame):
            return True
        if getattr(shp, "has_table", False) and shp.has_table:
            return True
    return False


def _text_plan(pptx: Path) -> list[list[_PlanItem]]:
    """Per slide, the shapes that can carry text (read with python-pptx).

    Only these are visited through COM, so pictures and freeforms cost nothing.
    Tables also record which cells are covered by a merge.
    """
    from pptx import Presentation

    out = []
    for slide in Presentation(str(pptx)).slides:
        items: list[_PlanItem] = []
        for idx, shp in enumerate(slide.shapes, start=1):
            sid, name = int(shp.shape_id), str(shp.name)
            if shp.shape_type == 6:
                if _group_has_text(shp):
                    items.append(_PlanItem(idx, "group", sid, name))
            elif getattr(shp, "has_table", False) and shp.has_table:
                t = shp.table
                spanned = {(r, c) for r in range(len(t.rows)) for c in range(len(t.columns))
                           if t.cell(r, c).is_spanned}
                items.append(_PlanItem(idx, "table", sid, name, {sid: spanned}))
            elif getattr(shp, "has_text_frame", False) and shp.has_text_frame and _has_text(
                    shp.text_frame):
                items.append(_PlanItem(idx, "text", sid, name))
        out.append(items)
    return out


def _worker_main(req_q: Any, resp_q: Any, heartbeat: Any, parent_pid: int,
                 adopt_paths: list[str]) -> None:
    """Worker process entry point. Protocol: requests ``(call_id, op, args)``,
    responses ``(call_id, status, payload)`` with status ok | error | fatal."""
    w = _Worker(heartbeat)
    w.beat()
    try:
        version = w.start(adopt_paths)
    except Exception as exc:
        resp_q.put((0, "unavailable", f"{type(exc).__name__}: {exc}"))
        with suppress(Exception):
            w.shutdown()
        return
    resp_q.put((0, "hello", {"version": version, "pid": os.getpid(),
                             "owned": sorted(w.owned.items())}))
    try:
        while True:
            w.beat()
            if not _pid_running(parent_pid):
                break  # client is gone: clean up and leave
            try:
                call_id, op, args = req_q.get(timeout=0.5)
            except queue.Empty:
                continue
            if op == "shutdown":
                break
            try:
                result = getattr(w, f"op_{op}")(*args)
                resp_q.put((call_id, "ok", result))
            except PowerPointError as exc:
                resp_q.put((call_id, "error", str(exc)))
            except Exception as exc:
                msg = f"{type(exc).__name__}: {exc}"
                if w.healthy():
                    resp_q.put((call_id, "error", msg))
                else:  # PowerPoint died or disconnected: let the client restart us
                    resp_q.put((call_id, "fatal", msg))
                    return
    finally:
        w.shutdown()
        with suppress(Exception):
            resp_q.put((-1, "bye", None))


# --------------------------------------------------------------------------- client side


@dataclass
class _Handle:
    path: str
    stat: tuple[float, int]
    info: PresInfo


class PowerPointEngine:
    """Client for the PowerPoint worker process. Thread-safe (calls are serialised)."""

    def __init__(self, *, call_timeout: float = 60.0, start_timeout: float = 90.0,
                 hang_timeout: float = 30.0, max_restarts: int = 3) -> None:
        self.call_timeout = call_timeout
        self.start_timeout = start_timeout
        self.hang_timeout = hang_timeout
        self.max_restarts = max_restarts
        self.version: str | None = None
        self.restarts = 0
        self.calls = 0
        self._ctx = mp.get_context("spawn")
        self._proc: Any = None
        self._req: Any = None
        self._resp: Any = None
        self._hb: Any = None
        self._owned: list[tuple[int, float]] = []
        self._worker_pid_seen = 0
        self.owned_history: set[int] = set()
        self._handles: dict[str, _Handle] = {}
        self._next_call = 1
        self._next_handle = 1
        self._lock = threading.RLock()

    # ---------------------------------------------------------------- lifecycle

    @property
    def worker_pid(self) -> int | None:
        return self._proc.pid if self._proc is not None and self._proc.is_alive() else None

    @property
    def owned_pids(self) -> list[int]:
        return [p for p, _ in self._owned]

    def start(self) -> PowerPointEngine:
        with self._lock:
            if self._proc is None or not self._proc.is_alive():
                self._spawn(adopt=[])
            return self

    def _spawn(self, adopt: list[str]) -> None:
        reap_orphans()
        self._req, self._resp = self._ctx.Queue(), self._ctx.Queue()
        self._hb = self._ctx.Value("d", time.time())
        self._proc = self._ctx.Process(
            target=_worker_main, name="pdf2pptx-powerpoint",
            args=(self._req, self._resp, self._hb, os.getpid(), adopt), daemon=True)
        self._proc.start()
        deadline = time.time() + self.start_timeout
        while True:
            try:
                _cid, status, payload = self._resp.get(timeout=0.2)
                break
            except queue.Empty:
                if not self._proc.is_alive():
                    self._discard_worker()
                    raise PowerPointUnavailable(
                        f"PowerPoint worker exited during start (code {self._proc.exitcode})"
                    ) from None
                if time.time() > deadline:
                    self._discard_worker(kill=True)
                    raise PowerPointUnavailable(
                        f"PowerPoint did not start within {self.start_timeout:.0f}s") from None
        if status != "hello":
            self._discard_worker(kill=True)
            raise PowerPointUnavailable(f"cannot start PowerPoint via COM: {payload}")
        self.version = payload["version"]
        self._worker_pid_seen = int(payload["pid"])
        self._owned = [(int(p), float(c)) for p, c in payload["owned"]]
        self.owned_history.update(p for p, _ in self._owned)
        LOG.debug("PowerPoint %s worker pid %s (owned POWERPNT.EXE: %s)", self.version,
                  payload["pid"], self.owned_pids or "none, attached to a running one")

    def _discard_worker(self, kill: bool = False) -> None:
        """Forget the current worker; with ``kill``, kill it and its PowerPoint first."""
        proc = self._proc
        if proc is not None:
            if kill and proc.is_alive():
                with suppress(Exception):
                    proc.kill()
            with suppress(Exception):
                proc.join(5)
        wpid = self._worker_pid_seen
        for pid, ct in self._owned:
            if proc is not None and proc.is_alive():
                continue  # still running (join timed out, not killed): leave it to reap
            _release(pid, ct, wpid, wait_s=0.0 if kill else 5.0)
        for q in (self._req, self._resp):
            if q is not None:
                with suppress(Exception):
                    q.cancel_join_thread()
                    q.close()
        self._proc = self._req = self._resp = self._hb = None
        self._owned = []

    def _restart(self, reason: str) -> None:
        self.restarts += 1
        LOG.warning("PowerPoint worker restart #%d: %s", self.restarts, reason)
        self._discard_worker(kill=True)
        self._spawn(adopt=sorted({h.path for h in self._handles.values()}))
        for handle, h in sorted(self._handles.items()):  # re-open what was open
            h.info = self._roundtrip("open", (handle, h.path), self.call_timeout)

    def stop(self) -> None:
        """Close presentations, quit our PowerPoint, end the worker. Idempotent."""
        with self._lock:
            proc = self._proc
            self._handles.clear()
            if proc is None:
                return
            if proc.is_alive():
                with suppress(Exception):
                    self._req.put((0, "shutdown", ()))
                proc.join(20)
            self._discard_worker(kill=proc.is_alive())

    def kill_worker(self) -> None:
        """Test hook: kill the worker process abruptly (as a crash would)."""
        if self._proc is not None and self._proc.is_alive():
            self._proc.kill()
            self._proc.join(5)

    def __enter__(self) -> PowerPointEngine:
        return self.start()

    def __exit__(self, *exc: object) -> None:
        self.stop()

    # ---------------------------------------------------------------- calls

    def _roundtrip(self, op: str, args: tuple[Any, ...], timeout: float) -> Any:
        """One request on the current worker. Raises _WorkerLost on death/hang/fatal."""
        if self._proc is None or not self._proc.is_alive():
            raise _WorkerLost("worker is not running")
        cid = self._next_call
        self._next_call += 1
        self.calls += 1
        self._req.put((cid, op, args))
        t0 = time.time()
        while True:
            try:
                rid, status, payload = self._resp.get(timeout=0.1)
            except queue.Empty:
                now = time.time()
                if not self._proc.is_alive():
                    raise _WorkerLost(f"worker died (exit code {self._proc.exitcode})") from None
                if now - t0 > timeout:
                    raise _WorkerLost(f"{op} timed out after {timeout:.0f}s") from None
                if now - self._hb.value > self.hang_timeout and now - t0 > self.hang_timeout:
                    raise _WorkerLost(f"{op}: no heartbeat for {self.hang_timeout:.0f}s") from None
                continue
            if rid != cid:
                continue  # a late answer to a call we already gave up on
            if status == "ok":
                return payload
            if status == "error":
                raise PowerPointError(f"{op}: {payload}")
            raise _WorkerLost(f"{op}: PowerPoint failed ({payload})")

    def _call(self, op: str, *args: Any, timeout: float | None = None) -> Any:
        with self._lock:
            last = ""
            for attempt in range(self.max_restarts + 1):
                try:
                    if self._proc is None or not self._proc.is_alive():
                        if attempt or self._proc is not None:
                            self._restart(last or "worker not running")
                        else:
                            self._spawn(adopt=[])
                    return self._roundtrip(op, args, timeout or self.call_timeout)
                except _WorkerLost as exc:
                    last = str(exc)
                    LOG.warning("PowerPoint call %s failed: %s", op, last)
                    self._discard_worker(kill=True)
                except PowerPointUnavailable as exc:
                    last = str(exc)
            raise PowerPointUnavailable(
                f"{op}: gave up after {self.max_restarts} restarts ({last})")

    def ping(self) -> float:
        """Round-trip time in seconds (also proves PowerPoint answers)."""
        t0 = time.perf_counter()
        self._call("ping", timeout=15)
        return time.perf_counter() - t0

    def open(self, pptx: Path | str) -> PresInfo:
        """Open read-only (or reuse an open, unchanged copy) and describe it."""
        p = os.path.normpath(str(Path(pptx).resolve()))
        st = os.stat(p)
        stat = (st.st_mtime, st.st_size)
        with self._lock:
            for handle, h in self._handles.items():
                if os.path.normcase(h.path) == os.path.normcase(p):
                    if h.stat == stat:
                        return h.info
                    self.close(handle)  # the file changed on disk: re-open it
                    break
            handle = f"h{self._next_handle}"
            self._next_handle += 1
            info: PresInfo = self._call("open", handle, p)
            self._handles[handle] = _Handle(p, stat, info)
            return info

    def close(self, pres: PresInfo | str) -> None:
        handle = pres.handle if isinstance(pres, PresInfo) else pres
        with self._lock:
            if self._handles.pop(handle, None) is None:
                return
            if self._proc is not None and self._proc.is_alive():
                with suppress(PowerPointError, _WorkerLost):
                    self._roundtrip("close", (handle,), 30)

    def close_all(self) -> None:
        for handle in list(self._handles):
            self.close(handle)

    def _h(self, pres: PresInfo | str) -> str:
        handle = pres.handle if isinstance(pres, PresInfo) else pres
        if handle not in self._handles:
            raise PowerPointError(f"presentation {handle!r} is not open")
        return handle

    def render(self, pres: PresInfo | str, slides: Sequence[int] | None, dpi: float,
               out_dir: Path) -> list[Path]:
        """Export the given slides (0-based; None = all) as ``slide_NNN.png``."""
        h = self._h(pres)
        n = len(slides) if slides is not None else self._handles[h].info.slides
        paths = self._call("render", h, None if slides is None else [int(s) for s in slides],
                           float(dpi), str(out_dir), timeout=self.call_timeout + 5 * n)
        return [Path(p) for p in paths]

    def text_layout(self, pres: PresInfo | str, slide: int) -> list[TextShape]:
        """Lines and words of every text shape on ``slide`` (0-based), slide points."""
        result: list[TextShape] = self._call("text_layout", self._h(pres), int(slide))
        return result

    def export_pdf(self, pres: PresInfo | str, out: Path) -> Path:
        return Path(self._call("export_pdf", self._h(pres), str(out)))

    def set_geometry(self, pres: PresInfo | str, slide: int,
                     edits: Sequence[tuple[str, float, float, float]]) -> int:
        """COM edit of named shapes' left/top/width in memory (benchmark only, M5)."""
        n: int = self._call("set_geometry", self._h(pres), int(slide),
                            [(str(a), float(b), float(c), float(d)) for a, b, c, d in edits])
        return n

    # convenience ---------------------------------------------------------

    @contextmanager
    def opened(self, pptx: Path | str) -> Iterator[PresInfo]:
        info = self.open(pptx)
        try:
            yield info
        finally:
            self.close(info)

    def slide_count(self, pptx: Path | str) -> int:
        with self.opened(pptx) as info:
            return info.slides


class _WorkerLost(Exception):
    pass


# --------------------------------------------------------------------------- shared instance

_SHARED: PowerPointEngine | None = None
_SHARED_ERROR: str | None = None
_SHARED_LOCK = threading.Lock()


def shared(**kwargs: Any) -> PowerPointEngine:
    """The process-wide long-lived engine (started on first use, stopped at exit).

    Raises :class:`PowerPointUnavailable` (and remembers it, so a missing PowerPoint
    costs one start attempt per process, not one per file).
    """
    global _SHARED, _SHARED_ERROR
    with _SHARED_LOCK:
        if _SHARED_ERROR is not None:
            raise PowerPointUnavailable(_SHARED_ERROR)
        if _SHARED is None:
            eng = PowerPointEngine(**kwargs)
            try:
                eng.start()
            except PowerPointUnavailable as exc:
                _SHARED_ERROR = str(exc)
                raise
            except Exception as exc:  # pywin32 missing, not Windows, ...
                _SHARED_ERROR = f"{type(exc).__name__}: {exc}"
                raise PowerPointUnavailable(_SHARED_ERROR) from exc
            _SHARED = eng
        return _SHARED


def shutdown_shared() -> None:
    global _SHARED, _SHARED_ERROR
    with _SHARED_LOCK:
        if _SHARED is not None:
            _SHARED.stop()
        _SHARED = None
        _SHARED_ERROR = None


# atexit runs hooks in reverse order: ours (registered after multiprocessing.util's)
# stops the worker cleanly before multiprocessing terminates daemon children.
atexit.register(shutdown_shared)


def probe(timeout: float = 90.0) -> str:
    """Start a private engine, return the PowerPoint version, stop it."""
    eng = PowerPointEngine(start_timeout=timeout)
    try:
        eng.start()
        eng.ping()
        return str(eng.version)
    finally:
        eng.stop()
