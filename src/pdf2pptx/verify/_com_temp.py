"""TEMPORARY PowerPoint COM helper (Milestone 1 only).

Prompt 3 replaces this module with ``verify/ppt_com.py`` (separate worker process,
heartbeat, timeouts, restart). Every COM call in M1 goes through here, so there
is exactly one place to replace.

Safety rules already honoured here:
* If PowerPoint was already running (the user's instance), we never call Quit and
  never kill it; we only close the presentations we opened.
* Otherwise we Quit on exit and kill any POWERPNT.EXE that appeared during our
  session and is still alive afterwards.
"""

from __future__ import annotations

import os
import time
from collections.abc import Iterator
from contextlib import contextmanager, suppress
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

PP_SAVE_AS_PDF = 32
MSO_GROUP = 6
PARAGRAPH_ENDS = (chr(13), chr(10), chr(11))  # CR, LF, vertical tab (line break)


class PowerPointUnavailable(RuntimeError):
    pass


# POWERPNT.EXE processes started by this Python process. A previous session's
# instance may still be shutting down when the next session starts; it must not
# be mistaken for the user's own PowerPoint.
_STARTED_BY_US: set[int] = set()


def _reap_ours(timeout: float = 10.0) -> None:
    import psutil

    deadline = time.time() + timeout
    while _STARTED_BY_US & _ppt_pids() and time.time() < deadline:
        time.sleep(0.25)
    for pid in list(_STARTED_BY_US & _ppt_pids()):
        with suppress(Exception):
            psutil.Process(pid).kill()
    _STARTED_BY_US.intersection_update(_ppt_pids())


def _ppt_pids() -> set[int]:
    import psutil

    pids: set[int] = set()
    for p in psutil.process_iter(["name"]):
        try:
            if (p.info["name"] or "").lower() == "powerpnt.exe":
                pids.add(p.pid)
        except Exception:
            continue
    return pids


@dataclass
class WordBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass
class LineBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    words: list[WordBox] = field(default_factory=list)
    shape: str = ""


class PowerPointSession:
    def __init__(self) -> None:
        self.app: Any = None
        self._pre_existing: set[int] = set()
        self._opened: list[Any] = []

    def start(self) -> PowerPointSession:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:  # pragma: no cover - non-Windows
            raise PowerPointUnavailable(f"pywin32 not available: {exc}") from exc
        pythoncom.CoInitialize()
        _reap_ours()
        self._pre_existing = _ppt_pids()
        try:
            self.app = win32com.client.DispatchEx("PowerPoint.Application")
            self.app.DisplayAlerts = 1  # ppAlertsNone
        except Exception as exc:
            raise PowerPointUnavailable(f"cannot start PowerPoint via COM: {exc}") from exc
        if not self._pre_existing:
            _STARTED_BY_US.update(_ppt_pids())
        return self

    @property
    def version(self) -> str:
        return str(self.app.Version)

    def open(self, path: Path) -> Any:
        full = os.path.normpath(str(Path(path).resolve()))
        # ReadOnly=True, Untitled=False, WithWindow=False
        pres = self.app.Presentations.Open(full, True, False, False)
        self._opened.append(pres)
        return pres

    def close(self, pres: Any) -> None:
        try:
            pres.Close()
        finally:
            if pres in self._opened:
                self._opened.remove(pres)

    def stop(self) -> None:
        for pres in list(self._opened):
            with suppress(Exception):
                pres.Close()
        self._opened.clear()
        if self.app is not None and not self._pre_existing:
            with suppress(Exception):
                self.app.Quit()
        self.app = None
        # kill only processes that we caused to start
        deadline = time.time() + 5
        while True:
            ours = _ppt_pids() - self._pre_existing
            if not ours or time.time() > deadline:
                break
            time.sleep(0.25)
        if ours:
            import psutil

            for pid in ours:
                with suppress(Exception):
                    psutil.Process(pid).kill()
        _STARTED_BY_US.intersection_update(_ppt_pids())
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:
            pass

    # ----------------------------------------------------------------- operations

    def export_pdf(self, pptx: Path, pdf: Path) -> None:
        pres = self.open(pptx)
        try:
            pdf.parent.mkdir(parents=True, exist_ok=True)
            pres.SaveAs(os.path.normpath(str(pdf.resolve())), PP_SAVE_AS_PDF)
        finally:
            self.close(pres)

    def slide_count(self, pptx: Path) -> int:
        """Open read-only (fails on a file that needs repair) and count slides."""
        pres = self.open(pptx)
        try:
            return int(pres.Slides.Count)
        finally:
            self.close(pres)

    def render(self, pptx: Path, out_dir: Path, dpi: int) -> list[Path]:
        """Export every slide as PNG at ``dpi`` (relative to the slide size in pt)."""
        out_dir.mkdir(parents=True, exist_ok=True)
        pres = self.open(pptx)
        try:
            w_pt = float(pres.PageSetup.SlideWidth)
            h_pt = float(pres.PageSetup.SlideHeight)
            w_px, h_px = round(w_pt * dpi / 72), round(h_pt * dpi / 72)
            paths = []
            for i in range(1, pres.Slides.Count + 1):
                p = out_dir / f"slide_{i:03d}.png"
                pres.Slides(i).Export(os.path.normpath(str(p.resolve())), "PNG", w_px, h_px)
                paths.append(p)
            return paths
        finally:
            self.close(pres)

    def text_layout(self, pptx: Path) -> list[list[LineBox]]:
        """Per slide: every rendered text line with its word boxes (pt, top-left)."""
        pres = self.open(pptx)
        try:
            spanned = _spanned_cells(pptx)
            slides = []
            for i in range(1, pres.Slides.Count + 1):
                lines: list[LineBox] = []
                for j in range(1, pres.Slides(i).Shapes.Count + 1):
                    _collect_shape(pres.Slides(i).Shapes(j), lines, spanned[i - 1])
                slides.append(lines)
            return slides
        finally:
            self.close(pres)


def _collect_text_frame(
    tf2: Any, name: str, out: list[LineBox], offset: tuple[float, float] = (0.0, 0.0)
) -> None:
    """Lines and words of one text frame.

    Word boxes are trimmed of trailing whitespace. ``y1`` of a line/word is the
    line box bottom minus the paragraph's space-after (on a paragraph's last
    line), which sits on the descent line like a PDF font-bbox bottom.
    """
    ox, oy = offset
    tr = tf2.TextRange
    n_lines = tr.Lines.Count
    for li in range(1, n_lines + 1):
        line = tr.Lines.Item(li)
        text = str(line.Text)
        if not text.strip():
            continue
        y0 = float(line.BoundTop) + oy
        y1 = y0 + float(line.BoundHeight)
        if text.endswith(PARAGRAPH_ENDS) or li == n_lines:
            pf = line.ParagraphFormat
            if not int(pf.LineRuleAfter):  # msoFalse: SpaceAfter is in points
                y1 -= max(0.0, float(pf.SpaceAfter))
        lb = LineBox(
            text=text,
            x0=float(line.BoundLeft) + ox,
            y0=y0,
            x1=float(line.BoundLeft) + float(line.BoundWidth) + ox,
            y1=y1,
            shape=name,
        )
        words = line.Words
        for wi in range(1, words.Count + 1):
            w = words.Item(wi).TrimText()
            wt = str(w.Text).strip()
            if not wt:
                continue
            lb.words.append(
                WordBox(
                    text=wt,
                    x0=float(w.BoundLeft) + ox,
                    y0=y0,
                    x1=float(w.BoundLeft) + float(w.BoundWidth) + ox,
                    y1=y1,
                )
            )
        out.append(lb)


def _collect_shape(
    shape: Any, out: list[LineBox], spanned: dict[int, set[tuple[int, int]]]
) -> None:
    name = str(shape.Name)
    if int(shape.Type) == MSO_GROUP:
        for k in range(1, shape.GroupItems.Count + 1):
            _collect_shape(shape.GroupItems(k), out, spanned)
        return
    if shape.HasTable:
        # Table cell text bounds come back relative to the cell origin.
        tbl = shape.Table
        skip = spanned.get(int(shape.Id), set())
        widths = [float(tbl.Columns(c).Width) for c in range(1, tbl.Columns.Count + 1)]
        heights = [float(tbl.Rows(r).Height) for r in range(1, tbl.Rows.Count + 1)]
        left, top = float(shape.Left), float(shape.Top)
        for r in range(1, tbl.Rows.Count + 1):
            for c in range(1, tbl.Columns.Count + 1):
                if (r - 1, c - 1) in skip:
                    continue
                cell_shape = tbl.Cell(r, c).Shape
                if cell_shape.HasTextFrame and cell_shape.TextFrame2.HasText:
                    # measured: relative x already includes the left margin, y does not
                    # include the top margin
                    margin_top = float(cell_shape.TextFrame2.MarginTop)
                    off = (left + sum(widths[: c - 1]),
                           top + sum(heights[: r - 1]) + margin_top)
                    _collect_text_frame(cell_shape.TextFrame2, f"{name}[{r},{c}]", out, off)
        return
    if shape.HasTextFrame and shape.TextFrame2.HasText:
        _collect_text_frame(shape.TextFrame2, name, out)


def _spanned_cells(pptx: Path) -> list[dict[int, set[tuple[int, int]]]]:
    """Per slide: table shape id -> cells covered by a merge (read with python-pptx)."""
    from pptx import Presentation

    out = []
    for slide in Presentation(str(pptx)).slides:
        d: dict[int, set[tuple[int, int]]] = {}
        for shp in slide.shapes:
            if getattr(shp, "has_table", False) and shp.has_table:
                t = shp.table
                d[int(shp.shape_id)] = {
                    (r, c)
                    for r in range(len(t.rows))
                    for c in range(len(t.columns))
                    if t.cell(r, c).is_spanned
                }
        out.append(d)
    return out


@contextmanager
def session() -> Iterator[PowerPointSession]:
    s = PowerPointSession().start()
    try:
        yield s
    finally:
        s.stop()


def probe() -> str:
    """Return the PowerPoint version, or raise PowerPointUnavailable."""
    with session() as s:
        return s.version
