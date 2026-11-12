"""Ground truth extracted from a PPTX (the answer key for a synthetic deck).

Everything here must be deterministic: no paths, times or random ids.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

from ..build.units import emu_to_pt

NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}


def _r(v: float) -> float:
    return round(v, 2)


def theme_info(prs: Any) -> dict[str, Any]:
    for part in prs.part.package.iter_parts():
        if re.search(r"/theme\d*\.xml$", str(part.partname)):
            xml = part.blob.decode("utf-8")
            colors: dict[str, str] = {}
            for key in ("dk1", "lt1", "dk2", "lt2", *(f"accent{i}" for i in range(1, 7))):
                m = re.search(rf"<a:{key}>\s*<a:(srgbClr val|sysClr val=\"\w+\" lastClr)=\"(\w+)\"",
                              xml)
                if m:
                    colors[key] = m.group(2).upper()
            major = re.search(r"<a:majorFont>\s*<a:latin typeface=\"([^\"]*)\"", xml)
            minor = re.search(r"<a:minorFont>\s*<a:latin typeface=\"([^\"]*)\"", xml)
            return {
                "colors": colors,
                "major_font": major.group(1) if major else None,
                "minor_font": minor.group(1) if minor else None,
            }
    return {"colors": {}, "major_font": None, "minor_font": None}


def _shape_kind(shape: Any) -> tuple[str, str | None]:
    st = shape.shape_type
    if shape.is_placeholder:
        pf = shape.placeholder_format
        if pf.type == PP_PLACEHOLDER.PICTURE or st == MSO_SHAPE_TYPE.PICTURE:
            return "picture", None
        return "placeholder", str(pf.type).split(" ")[0].lower()
    if st == MSO_SHAPE_TYPE.PICTURE:
        return "picture", None
    if st == MSO_SHAPE_TYPE.TEXT_BOX:
        return "text_box", None
    if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
        prst = shape._element.spPr.find(qn("a:prstGeom"))
        return "auto_shape", prst.get("prst") if prst is not None else None
    if st == MSO_SHAPE_TYPE.FREEFORM:
        return "freeform", None
    if st == MSO_SHAPE_TYPE.LINE or shape._element.tag == qn("p:cxnSp"):
        prst = shape._element.spPr.find(qn("a:prstGeom"))
        return "connector", prst.get("prst") if prst is not None else None
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group", None
    if getattr(shape, "has_table", False) and shape.has_table:
        return "table", None
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return "chart", None
    return str(st).split(" ")[0].lower() if st is not None else "unknown", None


def _resolve_font(run: Any, is_title: bool, theme: dict[str, Any]) -> str | None:
    name = run.font.name
    if name in (None, "+mn-lt"):
        return theme["minor_font"] if not (is_title and name is None) else theme["major_font"]
    if name == "+mj-lt":
        return theme["major_font"]
    return str(name)


def _bullet_of(p: Any, is_body_placeholder: bool) -> str:
    pPr = p._p.pPr
    if pPr is not None:
        if pPr.find(qn("a:buNone")) is not None:
            return "none"
        ch = pPr.find(qn("a:buChar"))
        if ch is not None:
            return f"char:{ch.get('char')}"
        an = pPr.find(qn("a:buAutoNum"))
        if an is not None:
            return f"autonum:{an.get('type')}"
    return "inherited" if is_body_placeholder else "none"


def _paragraphs(tf: Any, is_title: bool, is_body: bool, theme: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for p in tf.paragraphs:
        text = "".join(r.text for r in p.runs)
        # fields (slide number) are not runs in python-pptx
        for fld in p._p.findall(qn("a:fld")):
            t = fld.find(qn("a:t"))
            text += t.text if t is not None and t.text else ""
        if not text.strip():
            continue
        out.append(
            {
                "text": text,
                "level": int(p.level),
                "bullet": _bullet_of(p, is_body),
                "fonts": sorted({f for r in p.runs if (f := _resolve_font(r, is_title, theme))}),
                "has_field": bool(p._p.findall(qn("a:fld"))),
            }
        )
    return out


def _hyperlinks(shape: Any) -> list[str]:
    links = []
    try:
        ca = shape.click_action
        if ca.hyperlink.address:
            links.append(ca.hyperlink.address)
        elif ca.target_slide is not None:
            links.append("#slide")
    except Exception:
        pass
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.hyperlink.address:
                    links.append(r.hyperlink.address)
    return links


def _table_info(tbl: Any) -> dict[str, Any]:
    rows, cols = len(tbl.rows), len(tbl.columns)
    merges = []
    cells = []
    for r in range(rows):
        row_txt = []
        for c in range(cols):
            cell = tbl.cell(r, c)
            if cell.is_merge_origin:
                merges.append([r, c, int(cell.span_height), int(cell.span_width)])
            row_txt.append("" if cell.is_spanned else cell.text)
        cells.append(row_txt)
    return {"rows": rows, "cols": cols, "merges": merges, "cells": cells}


def _chart_info(chart: Any) -> dict[str, Any]:
    plot = chart.plots[0]
    return {
        "type": str(chart.chart_type).split(" ")[0],
        "categories": [str(c) for c in plot.categories],
        "series": [{"name": s.name, "values": [float(v) for v in s.values]} for s in plot.series],
        "has_legend": bool(chart.has_legend),
        "data_labels": bool(plot.has_data_labels),
    }


def _walk(shapes: Any, theme: dict[str, Any], out: dict[str, Any], depth: int = 0) -> None:
    for shape in shapes:
        kind, preset = _shape_kind(shape)
        is_title = shape.is_placeholder and shape.placeholder_format.type in _TITLE_TYPES
        is_body = shape.is_placeholder and not is_title
        entry: dict[str, Any] = {
            "kind": kind,
            "preset": preset,
            "bbox_pt": [
                _r(emu_to_pt(shape.left or 0)),
                _r(emu_to_pt(shape.top or 0)),
                _r(emu_to_pt((shape.left or 0) + (shape.width or 0))),
                _r(emu_to_pt((shape.top or 0) + (shape.height or 0))),
            ],
            "rotation": _r(float(shape.rotation)) if hasattr(shape, "rotation") else 0.0,
        }
        if kind == "placeholder":
            entry["placeholder"] = preset
            entry["preset"] = None
        out["shape_list"].append(entry)
        out["shapes"][kind] = out["shapes"].get(kind, 0) + 1
        out["hyperlinks"].extend(_hyperlinks(shape))
        if kind == "group":
            _walk(shape.shapes, theme, out, depth + 1)
            continue
        if getattr(shape, "has_table", False) and shape.has_table:
            info = _table_info(shape.table)
            out["tables"].append(info)
            out["texts"].extend(t for row in info["cells"] for t in row if t)
            continue
        if getattr(shape, "has_chart", False) and shape.has_chart:
            out["charts"].append(_chart_info(shape.chart))
            continue
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            paras = _paragraphs(shape.text_frame, is_title, is_body, theme)
            if paras:
                entry["role"] = "title" if is_title else ("body" if is_body else "text")
                out["paragraphs"].extend(
                    {**p, "role": entry["role"]} for p in paras
                )
                out["texts"].append("\n".join(p["text"] for p in paras))


def _sections(prs: Any) -> list[dict[str, Any]]:
    pres = prs.part._element
    order = {s.get("id"): i for i, s in enumerate(pres.find(qn("p:sldIdLst")))}
    out = []
    for sec in pres.iter(f"{{{NS_P14}}}section"):
        ids = [order[e.get("id")] for e in sec.iter(f"{{{NS_P14}}}sldId") if e.get("id") in order]
        out.append({"name": sec.get("name"), "slides": ids})
    return out


def extract(pptx: Path) -> dict[str, Any]:
    prs = Presentation(str(pptx))
    theme = theme_info(prs)
    slides = []
    for i, slide in enumerate(prs.slides):
        info: dict[str, Any] = {
            "shapes": {},
            "shape_list": [],
            "paragraphs": [],
            "texts": [],
            "tables": [],
            "charts": [],
            "hyperlinks": [],
        }
        _walk(slide.shapes, theme, info)
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text
        bg = slide._element.find(qn("p:cSld")).find(qn("p:bg"))
        slides.append(
            {
                "index": i,
                "layout": slide.slide_layout.name,
                "text": "\n".join(info.pop("texts")),
                **info,
                "shapes": dict(sorted(info["shapes"].items())),
                "notes": notes,
                "has_background": bg is not None,
                "has_slide_number_field": any(p["has_field"] for p in info["paragraphs"]),
            }
        )
    return {
        "slide_width_pt": _r(emu_to_pt(prs.slide_width)),
        "slide_height_pt": _r(emu_to_pt(prs.slide_height)),
        "slide_count": len(slides),
        "theme": theme,
        "sections": _sections(prs),
        "slides": slides,
    }
