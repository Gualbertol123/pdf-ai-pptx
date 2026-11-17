"""Content fingerprints of IR elements, independent of ids and page position (M7).

Two elements with the same fingerprint draw the same thing at their own bbox: same
text and run styling, same picture bytes and crop, same geometry relative to the frame,
same fill and line. Used to find repeated page furniture (``masters.py``) and build-up
sequences (``builds.py``).
"""

from __future__ import annotations

import hashlib
import json
import re
from typing import Any

from ..ir import Color, Fill, Group, Image, Paragraph, Shape, Stroke, Table, TextBlock

_DIGITS = re.compile(r"\d+")


def _r(v: float, q: float = 0.1) -> float:
    return round(round(v / q) * q, 3)


def _color(c: Color | None) -> Any:
    return None if c is None else [c.rgb, c.theme, _r(c.alpha, 0.01)]


def _fill(f: Fill) -> Any:
    g = f.gradient
    return [f.kind, _color(f.color), f.image_id and "img",
            None if g is None else [g.kind, _r(g.angle, 0.5),
                                    [[_r(s.pos, 0.01), _color(s.color)] for s in g.stops]]]


def _stroke(s: Stroke | None) -> Any:
    if s is None:
        return None
    return [_color(s.color), _r(s.width, 0.05), s.dash, s.head_end, s.tail_end, s.head_size,
            s.tail_size]


def paragraph_key(p: Paragraph, mask_digits: bool = False) -> list[Any]:
    runs = []
    for r in p.runs:
        text = _DIGITS.sub("#", r.text) if mask_digits else r.text
        runs.append([text, r.font, _r(r.size), _color(r.color), r.bold, r.italic, r.underline,
                     r.baseline_shift, r.hyperlink, r.field_type])
    b = p.bullet
    return [runs, p.alignment, p.level, b.kind, b.char, b.scheme, _r(p.indent_left),
            _r(p.indent_first)]


def paragraph_text(p: Paragraph) -> str:
    return "".join(r.text for r in p.runs)


def _content(e: Any, mask_digits: bool = False) -> Any:
    x0, y0 = e.bbox[0], e.bbox[1]
    size = [_r(e.bbox[2] - e.bbox[0], 0.5), _r(e.bbox[3] - e.bbox[1], 0.5)]
    if isinstance(e, TextBlock):
        # the frame of a text block follows its text: compare the text, not the frame size
        return ["text", [paragraph_key(p, mask_digits) for p in e.paragraphs], _r(e.rotation),
                e.vertical, e.placeholder]
    if isinstance(e, Image):
        return ["image", e.sha256 or e.file, [_r(c, 0.001) for c in e.crop], _r(e.rotation),
                e.flip_h, e.flip_v, _r(e.opacity, 0.01), size, e.hyperlink]
    if isinstance(e, Shape):
        path = None
        if e.path:
            path = [[s.op, [[_r(px - x0, 0.5), _r(py - y0, 0.5)] for px, py in s.pts]]
                    for s in e.path]
        adjust = {k: _r(v, 0.01) for k, v in sorted(e.adjust.items())}
        return ["shape", e.preset, adjust, path, _fill(e.fill), _stroke(e.stroke),
                _r(e.opacity, 0.01), _r(e.rotation), e.flip_h, e.flip_v, e.is_background, size,
                e.hyperlink, e.connector is not None]
    if isinstance(e, Table):
        cells = [[c.row, c.col, c.row_span, c.col_span,
                  [paragraph_key(p) for p in c.text.paragraphs] if c.text else None,
                  _fill(c.fill)] for c in e.cells]
        return ["table", [_r(v, 0.5) for v in e.row_heights], [_r(v, 0.5) for v in e.col_widths],
                cells]
    return [e.kind, size]


def fingerprint(e: Any, mask_digits: bool = False, children: list[Any] | None = None) -> str:
    """A short hash of what ``e`` draws (``children``: a group's members)."""
    data: Any = _content(e, mask_digits)
    if isinstance(e, Group) and children is not None:
        data = ["group", sorted(
            [_r(c.bbox[0] - e.bbox[0], 0.5), _r(c.bbox[1] - e.bbox[1], 0.5),
             fingerprint(c, mask_digits)] for c in children)]
    blob = json.dumps(data, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha1(blob.encode("utf-8")).hexdigest()[:16]


def same_box(a: tuple[float, float, float, float], b: tuple[float, float, float, float],
             tol: float) -> bool:
    return all(abs(u - v) <= tol for u, v in zip(a, b, strict=True))
