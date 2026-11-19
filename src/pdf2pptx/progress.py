"""Progress reporting for long conversions (the UI shows it per page).

Stages call :func:`report` with ``(stage, done, total)``; nothing happens unless a
callback is installed with :func:`set_callback` (the UI does, for its one running job).
A failing callback never stops a conversion. Work in pool processes (page parsing with
several workers) is reported when each chunk comes back.
"""

from __future__ import annotations

from collections.abc import Callable
from contextlib import suppress

Callback = Callable[[str, int, int, str], None]
_CB: Callback | None = None


def set_callback(cb: Callback | None) -> None:
    global _CB
    _CB = cb


def report(stage: str, done: int, total: int, message: str = "") -> None:
    cb = _CB
    if cb is None:
        return
    with suppress(Exception):  # progress is cosmetic
        cb(stage, done, total, message)
