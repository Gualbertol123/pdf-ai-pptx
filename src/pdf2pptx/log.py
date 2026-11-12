"""Logging setup shared by the CLI and worker processes."""

from __future__ import annotations

import logging
import sys

_CONFIGURED = False


class _StderrHandler(logging.StreamHandler):  # type: ignore[type-arg]
    """Writes to the *current* ``sys.stderr`` (pytest swaps and closes capture streams;
    a handler bound once would log into a closed file, e.g. from atexit hooks)."""

    @property  # type: ignore[override]
    def stream(self) -> object:
        return sys.stderr

    @stream.setter
    def stream(self, _value: object) -> None:
        pass


def setup(verbose: bool = False) -> None:
    global _CONFIGURED
    level = logging.DEBUG if verbose else logging.INFO
    if _CONFIGURED:
        logging.getLogger().setLevel(level)
        return
    handler = _StderrHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(name)s: %(message)s",
                                           "%H:%M:%S"))
    root = logging.getLogger()
    root.addHandler(handler)
    root.setLevel(level)
    for noisy in ("pdfminer", "PIL", "fontTools"):
        logging.getLogger(noisy).setLevel(logging.WARNING)
    _CONFIGURED = True


def get(name: str) -> logging.Logger:
    return logging.getLogger(name)
