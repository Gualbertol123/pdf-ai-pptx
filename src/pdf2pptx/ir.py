"""Intermediate representation (IR) shared by every pipeline stage.

Conventions (PLAN.md §4.1):

* Units are PDF points, origin at the top-left of the page, y pointing down.
  Conversion to EMU happens only in ``build/units.py``.
* Every element carries ``id``, ``bbox`` (x0, y0, x1, y1), ``z``, ``confidence``
  (0-1), ``source`` (``pdf`` | ``rule`` | ``ocr`` | ``tuner`` | ``model:<name>``)
  and ``notes``.
* Elements are tagged by a ``kind`` field so heterogeneous lists round-trip.
* ``Document.to_json()`` / ``Document.from_json()`` round-trip exactly:
  ``from_json(to_json(d)) == d`` and ``to_json(from_json(s)) == s`` for any
  ``s`` produced by ``to_json``.
"""

from __future__ import annotations

import json
import types
import typing
from dataclasses import MISSING, dataclass, field, fields, is_dataclass
from typing import Any, ClassVar, Self, TypeVar, Union, get_args, get_origin, get_type_hints

IR_VERSION = "1.0"

BBox = tuple[float, float, float, float]
Point = tuple[float, float]

T = TypeVar("T")


# --------------------------------------------------------------------------- styling


@dataclass(kw_only=True)
class Color:
    """An sRGB color (``rgb`` = ``RRGGBB``) or a theme color reference."""

    rgb: str | None = None
    theme: str | None = None  # dk1, lt1, accent1..6, ...
    alpha: float = 1.0  # 0 = transparent, 1 = opaque
    lum_mod: float | None = None
    lum_off: float | None = None


@dataclass(kw_only=True)
class GradientStop:
    pos: float  # 0..1
    color: Color


@dataclass(kw_only=True)
class Gradient:
    kind: str = "linear"  # linear | radial | rect | path
    angle: float = 0.0  # degrees, clockwise from +x (linear only)
    stops: list[GradientStop] = field(default_factory=list)
    center: Point | None = None  # radial: relative center (0..1)


@dataclass(kw_only=True)
class Fill:
    kind: str = "none"  # none | solid | gradient | image | pattern
    color: Color | None = None
    gradient: Gradient | None = None
    image_id: str | None = None


@dataclass(kw_only=True)
class Stroke:
    color: Color | None = None
    width: float = 0.75  # pt
    dash: str = "solid"  # solid | dash | dot | dashDot | lgDash | sysDash | custom
    dash_pattern: list[float] | None = None  # custom dash lengths in pt
    cap: str = "flat"  # flat | round | square
    join: str = "miter"  # miter | round | bevel
    head_end: str | None = None  # none | triangle | arrow | stealth | oval | diamond
    tail_end: str | None = None


@dataclass(kw_only=True)
class FontRef:
    """A font as found in the PDF plus the installed font it maps to."""

    id: str
    pdf_name: str  # e.g. "ABCDEF+Calibri-Bold"
    family: str  # cleaned family name, e.g. "Calibri"
    bold: bool = False
    italic: bool = False
    weight: int = 400
    flags: int = 0  # PDF font descriptor flags
    embedded: bool = False
    subset: bool = False
    matched_family: str | None = None  # installed family chosen by fonts/match.py
    match_score: float | None = None
    embeddable: bool | None = None  # OS/2 fsType allows embedding


# --------------------------------------------------------------------------- elements


@dataclass(kw_only=True)
class Element:
    """Fields shared by every IR element."""

    KIND: ClassVar[str] = "element"

    id: str
    bbox: BBox
    z: int = 0
    confidence: float = 1.0
    source: str = "pdf"
    notes: list[str] = field(default_factory=list)
    kind: str = field(default="element", init=False)

    def __post_init__(self) -> None:
        self.kind = type(self).KIND


@dataclass(kw_only=True)
class Glyph(Element):
    KIND: ClassVar[str] = "glyph"

    text: str
    font: str  # FontRef.id
    size: float  # pt
    color: Color = field(default_factory=lambda: Color(rgb="000000"))
    baseline: float = 0.0  # y of the baseline (pt, page space)
    origin_x: float = 0.0
    rotation: float = 0.0  # degrees, counter-clockwise as in PDF text matrix
    render_mode: int = 0  # PDF text render mode (3 = invisible)


@dataclass(kw_only=True)
class TextRun(Element):
    KIND: ClassVar[str] = "run"

    text: str
    font: str = ""  # FontRef.id or a family name when no FontRef exists
    size: float = 12.0
    color: Color = field(default_factory=lambda: Color(rgb="000000"))
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    baseline_shift: float = 0.0  # % of size; >0 superscript, <0 subscript
    char_spacing: float = 0.0  # pt
    hyperlink: str | None = None  # URL, or "#slide=<n>" for internal jumps
    field_type: str | None = None  # e.g. "slidenum"
    glyph_ids: list[str] = field(default_factory=list)


@dataclass(kw_only=True)
class Bullet:
    kind: str = "none"  # none | char | autonum | picture
    char: str | None = None
    scheme: str | None = None  # arabicPeriod, alphaLcParenR, romanLcPeriod, ...
    start_at: int = 1
    font: str | None = None
    color: Color | None = None
    size_pct: float | None = None
    image_id: str | None = None


@dataclass(kw_only=True)
class Paragraph(Element):
    KIND: ClassVar[str] = "paragraph"

    runs: list[TextRun] = field(default_factory=list)
    alignment: str = "left"  # left | center | right | justify | distributed
    indent_left: float = 0.0  # pt
    indent_first: float = 0.0  # pt, relative to indent_left (negative = hanging)
    level: int = 0
    bullet: Bullet = field(default_factory=Bullet)
    line_spacing: float | None = None  # value interpreted by line_spacing_rule
    line_spacing_rule: str = "auto"  # auto | multiple | exact
    space_before: float = 0.0  # pt
    space_after: float = 0.0  # pt


@dataclass(kw_only=True)
class TextBlock(Element):
    KIND: ClassVar[str] = "text_block"

    paragraphs: list[Paragraph] = field(default_factory=list)
    role: str = "body"  # title | subtitle | body | footer | page_number | caption | header
    role_confidence: float = 1.0
    autofit: str = "none"  # none | shrink | resize
    wrap: bool = True
    rotation: float = 0.0
    anchor: str = "top"  # top | middle | bottom
    margins: BBox = (0.0, 0.0, 0.0, 0.0)  # left, top, right, bottom (pt)
    placeholder: str | None = None  # title | body | ftr | sldNum | dt ...
    reading_order: int | None = None
    vertical: bool = False


@dataclass(kw_only=True)
class PathSegment:
    op: str  # M | L | C | Z
    pts: list[Point] = field(default_factory=list)


@dataclass(kw_only=True)
class Connector:
    kind: str = "straight"  # straight | elbow | curved
    start_id: str | None = None
    start_site: int | None = None
    end_id: str | None = None
    end_site: int | None = None


@dataclass(kw_only=True)
class Shape(Element):
    KIND: ClassVar[str] = "shape"

    preset: str | None = None  # rect, roundRect, ellipse, ... (None = custom path)
    adjust: dict[str, float] = field(default_factory=dict)
    path: list[PathSegment] | None = None
    fill: Fill = field(default_factory=Fill)
    stroke: Stroke | None = None
    opacity: float = 1.0
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    connector: Connector | None = None
    text: TextBlock | None = None
    hyperlink: str | None = None
    clip_id: str | None = None
    is_background: bool = False


@dataclass(kw_only=True)
class Image(Element):
    KIND: ClassVar[str] = "image"

    file: str  # path relative to the work dir
    sha256: str = ""
    crop: BBox = (0.0, 0.0, 0.0, 0.0)  # left, top, right, bottom fractions
    mask_file: str | None = None
    opacity: float = 1.0
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    pixel_size: tuple[int, int] = (0, 0)
    hyperlink: str | None = None
    alt_text: str = ""
    is_background: bool = False


@dataclass(kw_only=True)
class TableCell:
    row: int
    col: int
    row_span: int = 1
    col_span: int = 1
    text: TextBlock | None = None
    fill: Fill = field(default_factory=Fill)
    borders: dict[str, Stroke] = field(default_factory=dict)  # left/top/right/bottom
    margins: BBox = (7.2, 3.6, 7.2, 3.6)
    anchor: str = "top"
    merged_into: tuple[int, int] | None = None  # covered cells point to their origin


@dataclass(kw_only=True)
class Table(Element):
    KIND: ClassVar[str] = "table"

    row_heights: list[float] = field(default_factory=list)
    col_widths: list[float] = field(default_factory=list)
    cells: list[TableCell] = field(default_factory=list)
    header_rows: int = 0
    banded_rows: bool = False
    fallback_ids: list[str] = field(default_factory=list)


@dataclass(kw_only=True)
class ChartSeries:
    name: str
    values: list[float | None] = field(default_factory=list)
    color: Color | None = None
    x_values: list[float] | None = None


@dataclass(kw_only=True)
class Chart(Element):
    KIND: ClassVar[str] = "chart"

    chart_type: str  # bar_clustered, column_clustered, line, pie, scatter, area, ...
    categories: list[str] = field(default_factory=list)
    series: list[ChartSeries] = field(default_factory=list)
    title: str | None = None
    has_legend: bool = True
    legend_pos: str = "right"
    axes: dict[str, dict[str, float | str | None]] = field(default_factory=dict)
    fallback_ids: list[str] = field(default_factory=list)


@dataclass(kw_only=True)
class Equation(Element):
    KIND: ClassVar[str] = "equation"

    latex: str = ""
    omml: str | None = None
    fallback_ids: list[str] = field(default_factory=list)


@dataclass(kw_only=True)
class Group(Element):
    KIND: ClassVar[str] = "group"

    child_ids: list[str] = field(default_factory=list)
    name: str = ""


@dataclass(kw_only=True)
class Annotation(Element):
    """A PDF annotation: link, text note, comment."""

    KIND: ClassVar[str] = "annotation"

    subtype: str = "link"  # link | text | highlight | ...
    uri: str | None = None
    dest_page: int | None = None
    contents: str = ""
    author: str = ""


AnyElement = Union[  # noqa: UP007 - runtime dispatch needs a typing.Union object
    Glyph, TextRun, Paragraph, TextBlock, Shape, Image, Table, Chart, Equation, Group, Annotation
]

ELEMENT_TYPES: dict[str, type[Element]] = {
    cls.KIND: cls
    for cls in (
        Glyph,
        TextRun,
        Paragraph,
        TextBlock,
        Shape,
        Image,
        Table,
        Chart,
        Equation,
        Group,
        Annotation,
    )
}


# --------------------------------------------------------------------------- page / document


@dataclass(kw_only=True)
class Triage:
    kind: str = "born_digital"  # born_digital | broken_text | outlined_text | scanned
    reason: str = ""
    confidence: float = 1.0
    fallback_image: bool = False  # page was placed as a full-slide picture


@dataclass(kw_only=True)
class Page:
    index: int
    width: float  # pt (after applying rotation)
    height: float  # pt
    rotation: int = 0  # PDF /Rotate
    triage: Triage = field(default_factory=Triage)
    elements: list[AnyElement] = field(default_factory=list)
    glyphs: list[Glyph] = field(default_factory=list)
    annotations: list[Annotation] = field(default_factory=list)
    renders: dict[str, str] = field(default_factory=dict)  # dpi -> path
    layout_id: str | None = None
    notes_text: str = ""
    build_of: int | None = None  # page index this page builds on (animation steps)


@dataclass(kw_only=True)
class Theme:
    name: str = "pdf2pptx"
    colors: dict[str, str] = field(default_factory=dict)  # dk1, lt1, dk2, lt2, accent1..6
    major_font: str | None = None
    minor_font: str | None = None


@dataclass(kw_only=True)
class Layout:
    id: str
    name: str
    elements: list[AnyElement] = field(default_factory=list)
    placeholders: dict[str, BBox] = field(default_factory=dict)


@dataclass(kw_only=True)
class Master:
    id: str
    name: str = "Master"
    elements: list[AnyElement] = field(default_factory=list)
    layouts: list[Layout] = field(default_factory=list)


@dataclass(kw_only=True)
class Section:
    name: str
    first_page: int


@dataclass(kw_only=True)
class Document:
    ir_version: str = IR_VERSION
    source_path: str = ""
    sha256: str = ""
    metadata: dict[str, str] = field(default_factory=dict)
    pages: list[Page] = field(default_factory=list)
    fonts: dict[str, FontRef] = field(default_factory=dict)
    theme: Theme | None = None
    masters: list[Master] = field(default_factory=list)
    sections: list[Section] = field(default_factory=list)
    stage: str = "ingest"

    def to_json(self, indent: int | None = None) -> str:
        return to_json(self, indent=indent)

    @classmethod
    def from_json(cls, text: str) -> Self:
        return from_json(cls, text)


# --------------------------------------------------------------------------- codec


class IRError(ValueError):
    """Raised when JSON cannot be decoded into the IR."""


def to_dict(obj: Any) -> Any:
    """Encode an IR object to plain JSON-compatible data (field order preserved)."""
    if is_dataclass(obj) and not isinstance(obj, type):
        return {f.name: to_dict(getattr(obj, f.name)) for f in fields(obj)}
    if isinstance(obj, (list, tuple)):
        return [to_dict(v) for v in obj]
    if isinstance(obj, dict):
        return {str(k): to_dict(v) for k, v in obj.items()}
    if isinstance(obj, float) and obj != obj:  # NaN is not valid JSON
        raise IRError("NaN is not allowed in the IR")
    return obj


def to_json(obj: Any, indent: int | None = None) -> str:
    return json.dumps(to_dict(obj), ensure_ascii=False, indent=indent, allow_nan=False)


_HINTS: dict[type, dict[str, Any]] = {}


def _hints(cls: type) -> dict[str, Any]:
    if cls not in _HINTS:
        _HINTS[cls] = get_type_hints(cls)
    return _HINTS[cls]


def _decode(tp: Any, value: Any) -> Any:
    if tp is Any:
        return value
    origin = get_origin(tp)
    if origin is Union or origin is types.UnionType:
        args = get_args(tp)
        if value is None:
            if type(None) in args:
                return None
            raise IRError(f"null not allowed for {tp}")
        non_none = [a for a in args if a is not type(None)]
        element_args = [a for a in non_none if isinstance(a, type) and issubclass(a, Element)]
        if element_args and isinstance(value, dict):
            kind = value.get("kind")
            cls = ELEMENT_TYPES.get(str(kind))
            if cls is None or cls not in element_args:
                raise IRError(f"unknown element kind {kind!r}")
            return _decode(cls, value)
        if len(non_none) == 1:
            return _decode(non_none[0], value)
        for a in non_none:  # primitive unions like float | str
            if a in (bool, int, float, str) and isinstance(value, a):
                return value
        if float in non_none and isinstance(value, int):
            return value
        raise IRError(f"cannot decode {value!r} as {tp}")
    if origin is list:
        (item_tp,) = get_args(tp)
        if not isinstance(value, list):
            raise IRError(f"expected list for {tp}, got {type(value).__name__}")
        return [_decode(item_tp, v) for v in value]
    if origin is tuple:
        args = get_args(tp)
        if not isinstance(value, list):
            raise IRError(f"expected list for {tp}")
        if len(args) == 2 and args[1] is Ellipsis:
            return tuple(_decode(args[0], v) for v in value)
        if len(args) != len(value):
            raise IRError(f"expected {len(args)} items for {tp}, got {len(value)}")
        return tuple(_decode(a, v) for a, v in zip(args, value, strict=True))
    if origin is dict:
        _, val_tp = get_args(tp)
        if not isinstance(value, dict):
            raise IRError(f"expected object for {tp}")
        return {k: _decode(val_tp, v) for k, v in value.items()}
    if isinstance(tp, type) and is_dataclass(tp):
        if not isinstance(value, dict):
            raise IRError(f"expected object for {tp.__name__}")
        hints = _hints(tp)
        kwargs: dict[str, Any] = {}
        for f in fields(tp):
            if not f.init:
                continue
            if f.name in value:
                kwargs[f.name] = _decode(hints[f.name], value[f.name])
            elif f.default is MISSING and f.default_factory is MISSING:
                raise IRError(f"{tp.__name__}: missing field {f.name!r}")
        unknown = set(value) - {f.name for f in fields(tp)}
        if unknown:
            raise IRError(f"{tp.__name__}: unknown fields {sorted(unknown)}")
        return tp(**kwargs)
    if tp is float:
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise IRError(f"expected number, got {value!r}")
        return value
    if tp in (int, str, bool):
        if not isinstance(value, tp) or (tp is int and isinstance(value, bool)):
            raise IRError(f"expected {tp.__name__}, got {value!r}")
        return value
    raise IRError(f"unsupported type {tp!r}")


def from_dict[T](cls: type[T], data: Any) -> T:
    return typing.cast(T, _decode(cls, data))


def from_json[T](cls: type[T], text: str) -> T:
    data = json.loads(text)
    if cls is Document:
        version = data.get("ir_version") if isinstance(data, dict) else None
        if version != IR_VERSION:
            raise IRError(f"unsupported ir_version {version!r} (expected {IR_VERSION})")
    return from_dict(cls, data)


def iter_elements(page: Page) -> typing.Iterator[AnyElement]:
    """All elements on a page in a stable order (z, then list order)."""
    yield from sorted(page.elements, key=lambda e: e.z)
