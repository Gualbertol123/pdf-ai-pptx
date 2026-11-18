"""JSON schemas for constrained model output (llama.cpp turns them into a grammar),
the prompts that go with them, and validation of what comes back.

Every validator returns None for output that is invalid for the given input; the
caller then keeps the rule result.
"""

from __future__ import annotations

import math
from typing import Any

REGION_LABELS = ["timeline", "process", "matrix", "org_chart", "table", "chart", "diagram",
                 "text", "photo", "logo", "decorative"]
ROLES = ["title", "subtitle", "body", "caption", "header", "footer", "page_number", "other"]
CHART_TYPES = ["column_clustered", "column_stacked", "bar_clustered", "bar_stacked", "line",
               "line_markers", "pie", "doughnut", "area", "scatter", "other"]

_CONF = {"type": "number", "minimum": 0, "maximum": 1}

REGION = {
    "type": "object",
    "properties": {"label": {"type": "string", "enum": REGION_LABELS}, "confidence": _CONF},
    "required": ["label", "confidence"],
}

ROLE_FIX = {
    "type": "object",
    "properties": {"blocks": {"type": "array", "items": {
        "type": "object",
        "properties": {"id": {"type": "string"},
                       "role": {"type": "string", "enum": ROLES},
                       "confidence": _CONF},
        "required": ["id", "role", "confidence"]}}},
    "required": ["blocks"],
}

READING_ORDER = {
    "type": "object",
    "properties": {"order": {"type": "array", "items": {"type": "integer"}},
                   "confidence": _CONF},
    "required": ["order", "confidence"],
}

TABLE_FIX = {
    "type": "object",
    "properties": {"rows": {"type": "integer", "minimum": 1, "maximum": 60},
                   "cols": {"type": "integer", "minimum": 1, "maximum": 30},
                   "header_rows": {"type": "integer", "minimum": 0, "maximum": 5},
                   "cells": {"type": "array", "items": {"type": "array",
                                                         "items": {"type": "string"}}},
                   "confidence": _CONF},
    "required": ["rows", "cols", "header_rows", "cells", "confidence"],
}

CHART = {
    "type": "object",
    "properties": {
        "chart_type": {"type": "string", "enum": CHART_TYPES},
        "title": {"type": "string"},
        "categories": {"type": "array", "items": {"type": "string"}},
        "series": {"type": "array", "items": {
            "type": "object",
            "properties": {"name": {"type": "string"},
                           "values": {"type": "array", "items": {"type": ["number", "null"]}}},
            "required": ["name", "values"]}},
        "axis_labels_visible": {"type": "boolean"},
        "value_axis_min": {"type": ["number", "null"]},
        "value_axis_max": {"type": ["number", "null"]},
        "confidence": _CONF,
    },
    "required": ["chart_type", "categories", "series", "confidence"],
}

SYSTEM = ("You check the output of a PDF-to-PowerPoint converter. Answer only with JSON that "
          "matches the schema. Use only what is visible in the image and the given data.")


def region_prompt(kind: str, n_children: int) -> str:
    return ("The image is one region of a document page (" + kind + f", {n_children} parts). "
            "Classify it: timeline, process (steps / flow), matrix (grid of boxes), "
            "org_chart, table, chart (data plot), diagram (other), text, photo, logo or "
            "decorative. Give your confidence (0-1).")


def role_prompt(blocks: list[tuple[str, str, str]]) -> str:
    lines = "\n".join(f'- id {bid}: "{text[:160]}" (current role: {role})'
                      for bid, text, role in blocks)
    return ("The image is a document page. These text blocks have uncertain roles:\n"
            f"{lines}\n"
            "For each id give its role: title (the page or slide title), subtitle, body, "
            "caption (describes a figure or table), header (running header), footer, "
            "page_number or other, and your confidence (0-1).")


def order_prompt(blocks: list[tuple[int, str]]) -> str:
    lines = "\n".join(f"{i}: \"{t[:100]}\"" for i, t in blocks)
    return ("The image is a document page with numbered boxes around its text blocks:\n"
            f"{lines}\n"
            "Give the numbers in the order a person reads the page (all numbers, each once), "
            "and your confidence (0-1).")


def table_prompt(rows: list[list[str]]) -> str:
    grid = "\n".join(" | ".join(c[:40] for c in r) for r in rows)
    return ("The image is a table from a document. A converter read it as this grid "
            f"({len(rows)} rows):\n{grid}\n"
            "Give the table's true number of rows and columns, how many header rows it has, "
            "the text of every cell row by row (merged cells: text in the first cell, empty "
            "strings in the rest) and your confidence (0-1).")


def chart_prompt(hint: str = "") -> str:
    return ("The image is a chart. Extract its data: chart type (column = vertical bars, "
            "bar = horizontal bars; clustered = side by side, stacked = on top of each "
            "other; line_markers = lines with point markers), category labels in axis "
            "order (for bar charts bottom to top as plotted: the first category is the one "
            "nearest the axis origin), every series with its name (from the legend; empty if "
            "none) and one value per category read from the bars / points / slices and their "
            "labels, whether the axes show tick labels (numbers or category names), the value "
            "axis minimum and maximum if shown, and your confidence (0-1). "
            + hint)


# --------------------------------------------------------------------------- validation


def _conf(d: dict[str, Any]) -> float | None:
    c = d.get("confidence")
    if not isinstance(c, (int, float)) or math.isnan(c) or not 0 <= c <= 1:
        return None
    return float(c)


def valid_region(d: Any) -> dict[str, Any] | None:
    if not isinstance(d, dict) or d.get("label") not in REGION_LABELS or _conf(d) is None:
        return None
    return d


def valid_roles(d: Any, ids: set[str]) -> dict[str, tuple[str, float]] | None:
    if not isinstance(d, dict) or not isinstance(d.get("blocks"), list):
        return None
    out: dict[str, tuple[str, float]] = {}
    for b in d["blocks"]:
        if not isinstance(b, dict) or b.get("id") not in ids or b.get("role") not in ROLES:
            continue
        c = _conf(b)
        if c is not None:
            out[str(b["id"])] = (str(b["role"]), c)
    return out or None


def valid_order(d: Any, n: int) -> tuple[list[int], float] | None:
    if not isinstance(d, dict):
        return None
    order, c = d.get("order"), _conf(d)
    if c is None or not isinstance(order, list) or sorted(order) != list(range(n)):
        return None
    return [int(i) for i in order], c


def valid_table(d: Any) -> dict[str, Any] | None:
    if not isinstance(d, dict) or _conf(d) is None:
        return None
    rows, cols, cells = d.get("rows"), d.get("cols"), d.get("cells")
    if not isinstance(rows, int) or not isinstance(cols, int) or not isinstance(cells, list):
        return None
    if not cells or any(not isinstance(r, list) for r in cells):
        return None
    # the cells are the evidence: stated counts that disagree with them are corrected
    rows, cols = len(cells), max(len(r) for r in cells)
    d = d | {"rows": rows, "cols": cols}
    # ragged rows (merged cells written once): pad / cut to the stated width
    d = d | {"cells": [[str(c) for c in r][:cols] + [""] * (cols - len(r)) for r in cells]}
    return d


def valid_chart(d: Any) -> dict[str, Any] | None:
    if not isinstance(d, dict) or d.get("chart_type") not in CHART_TYPES or _conf(d) is None:
        return None
    cats, series = d.get("categories"), d.get("series")
    if not isinstance(cats, list) or not cats or not isinstance(series, list) or not series:
        return None
    n = len(cats)
    # a pie written as one single-value series per slice: one series
    if (d["chart_type"] in ("pie", "doughnut") and len(series) == n and n > 1
            and all(isinstance(s, dict) and isinstance(s.get("values"), list)
                    and len(s["values"]) == 1 for s in series)):
        series = [{"name": "", "values": [s["values"][0] for s in series]}]
        d = d | {"series": series}
    for s in series:
        vals = s.get("values") if isinstance(s, dict) else None
        if not isinstance(vals, list) or len(vals) != n:
            return None
        if any(v is not None and (not isinstance(v, (int, float)) or math.isnan(v))
               for v in vals):
            return None
    return d
