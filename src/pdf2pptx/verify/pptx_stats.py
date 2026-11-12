"""PPTX side of scoring that needs no PowerPoint: text, text frames, shape counts."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from ..build.units import emu_to_pt


@dataclass
class SlideStats:
    index: int
    text: str = ""  # all visible text, z-order
    frame_text: str = ""  # text in text frames, tables and charts (real, editable text)
    n_text_boxes: int = 0  # shapes carrying text (a table counts once)
    n_preset: int = 0  # built-in autoshapes + connectors (excluding text boxes)
    n_freeform: int = 0
    n_pictures: int = 0
    n_tables: int = 0
    n_charts: int = 0
    parts: list[str] = field(default_factory=list)


@dataclass
class DeckStats:
    width_pt: float
    height_pt: float
    slides: list[SlideStats]


def _para_text(p: Any) -> str:
    out = []
    for child in p._p:
        if child.tag in (qn("a:r"), qn("a:fld")):
            t = child.find(qn("a:t"))
            if t is not None and t.text:
                out.append(t.text)
        elif child.tag == qn("a:br"):
            out.append("\n")
    return "".join(out)


def _roman(n: int) -> str:
    vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
            (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for v, r in vals:
        while n >= v:
            out += r
            n -= v
    return out


def autonum_label(scheme: str, n: int) -> str:
    """Render a DrawingML autonumber (e.g. arabicPeriod, alphaLcParenR) for item n."""
    m = re.match(r"(arabic|alphaLc|alphaUc|romanLc|romanUc)(\w*)", scheme)
    kind, suffix = (m.group(1), m.group(2)) if m else ("arabic", "Period")
    if kind == "alphaLc":
        num = chr(ord("a") + (n - 1) % 26)
    elif kind == "alphaUc":
        num = chr(ord("A") + (n - 1) % 26)
    elif kind == "romanLc":
        num = _roman(n)
    elif kind == "romanUc":
        num = _roman(n).upper()
    else:
        num = str(n)
    return {"Period": f"{num}.", "ParenR": f"{num})", "ParenBoth": f"({num})",
            "Plain": num}.get(suffix, f"{num}.")


def _ppr_bullet(ppr: Any) -> tuple[str, str | None] | None:
    if ppr is None:
        return None
    if ppr.find(qn("a:buNone")) is not None:
        return ("none", None)
    ch = ppr.find(qn("a:buChar"))
    if ch is not None:
        return ("char", ch.get("char"))
    an = ppr.find(qn("a:buAutoNum"))
    if an is not None:
        return ("autonum", an.get("type"))
    return None


def _lvl_bullet(container: Any, level: int) -> tuple[str, str | None] | None:
    """Bullet set on lvlNpPr of a lstStyle / txStyles element, or None."""
    if container is None:
        return None
    return _ppr_bullet(container.find(qn(f"a:lvl{level + 1}pPr")))


def _inherited_bullet(shape: Any, level: int) -> tuple[str, str | None]:
    """Placeholder inheritance: own lstStyle -> layout -> master placeholder -> txStyles."""
    if shape is None or not shape.is_placeholder:
        return ("none", None)
    el = shape
    for _ in range(3):  # slide -> layout -> master
        body = el._element.find(qn("p:txBody"))
        if body is not None:
            b = _lvl_bullet(body.find(qn("a:lstStyle")), level)
            if b is not None:
                return b
        try:
            el = el._base_placeholder
        except Exception:
            el = None
        if el is None:
            break
    ph_type = str(shape.placeholder_format.type)
    styles = shape.part.slide_layout.slide_master._element.find(qn("p:txStyles"))
    if styles is None:
        return ("none", None)
    if "TITLE" in ph_type:
        key = "p:titleStyle"
    elif "BODY" in ph_type or "OBJECT" in ph_type:
        key = "p:bodyStyle"
    else:
        key = "p:otherStyle"
    return _lvl_bullet(styles.find(qn(key)), level) or ("none", None)


def _frame_text(tf: Any, shape: Any = None) -> str:
    """Visible text of a frame, with list labels (bullet chars, autonumbers) rendered."""
    lines = []
    counters: dict[int, tuple[str, int]] = {}
    for p in tf.paragraphs:
        t = _para_text(p)
        if not t.strip():
            continue
        level = int(p.level)
        bullet = _ppr_bullet(p._p.pPr) or _inherited_bullet(shape, level)
        for deeper in [k for k in counters if k > level]:
            del counters[deeper]
        label = ""
        if bullet[0] == "char" and bullet[1]:
            label = bullet[1]
            counters.pop(level, None)
        elif bullet[0] == "autonum" and bullet[1]:
            prev = counters.get(level)
            n = prev[1] + 1 if prev and prev[0] == bullet[1] else 1
            counters[level] = (bullet[1], n)
            label = autonum_label(bullet[1], n)
        else:
            counters.pop(level, None)
        lines.append(f"{label} {t}" if label else t)
    return "\n".join(lines)


def _chart_text(chart: Any) -> str:
    parts = []
    if chart.has_title and chart.chart_title.has_text_frame:
        parts.append(_frame_text(chart.chart_title.text_frame))
    plot = chart.plots[0]
    parts.extend(str(c) for c in plot.categories)
    if chart.has_legend:
        parts.extend(s.name for s in plot.series)
    return "\n".join(p for p in parts if p)


def _walk(shapes: Any, st: SlideStats) -> None:
    for shape in shapes:
        stype = shape.shape_type
        if stype == MSO_SHAPE_TYPE.GROUP:
            _walk(shape.shapes, st)
            continue
        if getattr(shape, "has_table", False) and shape.has_table:
            st.n_tables += 1
            texts = [shape.table.cell(r, c).text
                     for r in range(len(shape.table.rows))
                     for c in range(len(shape.table.columns))
                     if not shape.table.cell(r, c).is_spanned]
            txt = "\n".join(t for t in texts if t.strip())
            if txt:
                st.n_text_boxes += 1
                st.parts.append(txt)
            continue
        if getattr(shape, "has_chart", False) and shape.has_chart:
            st.n_charts += 1
            txt = _chart_text(shape.chart)
            if txt:
                st.parts.append(txt)
            continue
        if stype == MSO_SHAPE_TYPE.PICTURE or (
            shape.is_placeholder and shape._element.tag == qn("p:pic")
        ):
            st.n_pictures += 1
            continue
        if stype == MSO_SHAPE_TYPE.FREEFORM:
            st.n_freeform += 1
        elif shape._element.tag == qn("p:cxnSp") or (
            stype == MSO_SHAPE_TYPE.AUTO_SHAPE
        ):
            st.n_preset += 1
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            txt = _frame_text(shape.text_frame, shape)
            if txt.strip():
                st.n_text_boxes += 1
                st.parts.append(txt)


def deck_stats(pptx: Path) -> DeckStats:
    prs = Presentation(str(pptx))
    slides = []
    for i, slide in enumerate(prs.slides):
        st = SlideStats(index=i)
        _walk(slide.shapes, st)
        st.text = "\n".join(st.parts)
        st.frame_text = st.text
        slides.append(st)
    return DeckStats(emu_to_pt(prs.slide_width), emu_to_pt(prs.slide_height), slides)


def blank_deck(out: Path, sizes_pt: list[tuple[float, float]]) -> Path:
    """A PPTX with one empty slide per page (sanity baseline: should score low)."""
    from ..build.units import pt_to_emu

    prs = Presentation()
    w, h = sizes_pt[0] if sizes_pt else (960.0, 540.0)
    prs.slide_width = pt_to_emu(w)
    prs.slide_height = pt_to_emu(h)
    for _ in sizes_pt:
        prs.slides.add_slide(prs.slide_layouts[6])
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return out
