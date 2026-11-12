from __future__ import annotations

import functools
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
CORPUS = ROOT / "corpus"


@functools.cache
def powerpoint_available() -> bool:
    from pdf2pptx.doctor import check_powerpoint

    return bool(check_powerpoint(timeout=90)["available"])


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    has_corpus = (CORPUS / "manifest.json").is_file()
    for item in items:
        if "corpus" in item.keywords and not has_corpus:
            item.add_marker(pytest.mark.skip(reason="corpus not built (pdf2pptx corpus build)"))
        if "powerpoint" in item.keywords and not powerpoint_available():
            item.add_marker(pytest.mark.skip(reason="PowerPoint COM not available"))
