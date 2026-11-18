"""Chart data from a vision model (Max tier, M9).

* **Vector charts** (built by ``analyze/charts.py`` from the PDF geometry): the model
  reads the same region; :func:`compare` checks type, categories and values. The
  geometry stays the source of the values; the model's verdict is recorded
  (``vlm_check:``) and raises or lowers the chart's confidence.
* **Raster charts** (pictures): the model gives type, categories, series names and
  approximate values; for bar / column charts the bar lengths are then measured in the
  pixels (:func:`measure_bars`) and replace the model's values, scaled to its axis.
  Series colours are sampled from the picture. The chart keeps the picture as its
  fallback; the render check decides.
"""

from __future__ import annotations

import io
import itertools
import re
from typing import Any

import cv2
import numpy as np

from ..ir import Chart, ChartSeries, Color, Image
from . import schemas
from .server import MAX_IMAGE_PX

MIN_IMAGE_PT = 80.0  # pictures smaller than this (either side) are not chart candidates
MIN_ADOPT = 0.6  # model confidence needed to build a chart from a picture


def crop_png(img: np.ndarray, bbox: tuple[float, float, float, float], scale: float,
             pad: float = 4.0, max_px: int = MAX_IMAGE_PX) -> bytes:
    """``bbox`` (pt) of a page image (BGR, ``scale`` px/pt) as PNG, longest side capped."""
    h, w = img.shape[:2]
    x0, y0 = max(0, int((bbox[0] - pad) * scale)), max(0, int((bbox[1] - pad) * scale))
    x1, y1 = min(w, int((bbox[2] + pad) * scale)), min(h, int((bbox[3] + pad) * scale))
    crop = img[y0:y1, x0:x1]
    f = min(1.0, max_px / max(crop.shape[:2]))
    if f < 1.0:
        crop = cv2.resize(crop, (max(1, int(crop.shape[1] * f)), max(1, int(crop.shape[0] * f))),
                          interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".png", crop)
    return bytes(buf) if ok else b""


def family(t: str) -> str:
    t = t.lower()
    for k in ("column", "bar", "line", "pie", "doughnut", "area", "scatter"):
        if t.startswith(k):
            return k
    return "other"


def _norm(s: str) -> str:
    return re.sub(r"[^0-9a-zà-ÿ]", "", s.lower())


def compare(ch: Chart, vlm: dict[str, Any]) -> dict[str, Any]:
    """How well the model's reading agrees with a geometry-built chart."""
    out: dict[str, Any] = {"type_ok": family(vlm["chart_type"]) == family(ch.chart_type)}
    vc = [_norm(c) for c in vlm["categories"]]
    rc = [_norm(c) for c in ch.categories]
    if ch.chart_type.startswith("bar") and vc == list(reversed(rc)):
        vlm = vlm | {"categories": list(reversed(vlm["categories"])),
                     "series": [s | {"values": list(reversed(s["values"]))}
                                for s in vlm["series"]]}
        vc = rc
    out["categories_ok"] = vc == rc
    ax = ch.axes.get("value") or {}
    rng = float(ax.get("max") or 0) - float(ax.get("min") or 0)
    if not rng:
        rng = max((abs(v) for s in ch.series for v in s.values if v is not None), default=1.0)
    agree = total = 0
    for k, s in enumerate(ch.series):
        vs = vlm["series"][k]["values"] if k < len(vlm["series"]) else []
        for i, v in enumerate(s.values):
            if v is None:
                continue
            total += 1
            m = vs[i] if i < len(vs) else None
            if m is not None and abs(float(m) - v) <= 0.02 * rng:
                agree += 1
    out["values_within_2pct"] = round(agree / total, 3) if total else None
    out["agree"] = bool(out["type_ok"] and out["categories_ok"]
                        and (out["values_within_2pct"] or 0) >= 0.9)
    return out


def apply_check(ch: Chart, vlm: dict[str, Any] | None, model: str) -> dict[str, Any] | None:
    """Record the model's verdict on a geometry chart (values are not replaced)."""
    v = schemas.valid_chart(vlm)
    if v is None:
        ch.notes.append(f"vlm_check:{model}:invalid")
        return None
    res = compare(ch, v)
    ch.notes.append(f"vlm_check:{model}:{'agree' if res['agree'] else 'disagree'}")
    if res["agree"]:
        ch.confidence = round(min(1.0, ch.confidence + 0.04), 3)
    else:
        ch.confidence = round(max(0.0, ch.confidence - 0.1), 3)
    # unnamed series: take the model's names
    for k, s in enumerate(ch.series):
        if s.name.startswith("Series ") and k < len(v["series"]) and v["series"][k]["name"]:
            s.name = str(v["series"][k]["name"])[:80]
    return res


# --------------------------------------------------------------------------- raster


def _saturated(img: np.ndarray) -> np.ndarray:
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    return (hsv[:, :, 1] > 60) & (hsv[:, :, 2] > 40)


def dominant_colors(img: np.ndarray, k: int) -> list[str]:
    """The ``k`` most frequent saturated colours (RRGGBB), most frequent first."""
    m = _saturated(img)
    px = img[m].reshape(-1, 3)
    if len(px) == 0:
        return []
    q = (px // 8).astype(np.int32)
    keys = q[:, 0] * 1024 + q[:, 1] * 32 + q[:, 2]
    vals, counts = np.unique(keys, return_counts=True)
    out: list[str] = []
    for key in vals[np.argsort(-counts)]:
        b, g, r = (int(key) // 1024) * 8 + 4, (int(key) // 32 % 32) * 8 + 4, (int(key) % 32) * 8 + 4
        rgb = f"{r:02X}{g:02X}{b:02X}"
        if all(sum(abs(int(rgb[i:i + 2], 16) - int(o[i:i + 2], 16)) for i in (0, 2, 4)) > 60
               for o in out):
            out.append(rgb)
        if len(out) >= k:
            break
    return out


def bar_orientation(img: np.ndarray) -> str | None:
    """``column`` if the solid bars share a bottom edge, ``bar`` if they share a left
    edge, None if neither (or fewer than two bars)."""
    m = _saturated(img).astype(np.uint8)
    n, _lab, stats, _ = cv2.connectedComponentsWithStats(m, 8)
    rects = [stats[i] for i in range(1, n) if stats[i][4] >= 30
             and stats[i][4] >= 0.85 * stats[i][2] * stats[i][3]]
    if len(rects) < 2:
        return None
    bottoms = [r[1] + r[3] for r in rects]
    lefts = [r[0] for r in rects]
    same_bottom = sum(1 for b in bottoms if abs(b - max(bottoms)) <= 3)
    same_left = sum(1 for x in lefts if abs(x - min(lefts)) <= 3)
    if same_bottom == len(rects) and same_left < len(rects):
        return "column"
    if same_left == len(rects) and same_bottom < len(rects):
        return "bar"
    return None


def bar_geometry(img: np.ndarray, vertical: bool) -> dict[str, Any] | None:
    """Plot rectangle (px: base line to the tallest bar, across the category slots) and
    gap width of a single-series bar picture."""
    m = _saturated(img).astype(np.uint8)
    n, _lab, stats, _ = cv2.connectedComponentsWithStats(m, 8)
    rects = [stats[i][:4] for i in range(1, n) if stats[i][4] >= 30
             and stats[i][4] >= 0.85 * stats[i][2] * stats[i][3]]
    if len(rects) < 2:
        return None
    if not vertical:  # transpose: same code for bars
        img = cv2.transpose(img)
        g = bar_geometry(img, True)
        if g is None:
            return None
        x0, y0, x1, y1 = g["plot"]
        return {"plot": (y0, x0, y1, x1), "gap_width": g["gap_width"]}
    rects.sort(key=lambda r: r[0])
    bw = float(np.median([r[2] for r in rects]))
    pitch = float(np.median([b[0] - a[0] for a, b in itertools.pairwise(rects)]))
    if pitch <= bw:
        return None
    base = max(r[1] + r[3] for r in rects)
    top = min(r[1] for r in rects)
    # slots are centred on the bars
    c0 = rects[0][0] + rects[0][2] / 2
    x0 = c0 - pitch / 2
    x1 = x0 + pitch * len(rects)
    return {"plot": (x0, float(top), x1, float(base)),
            "gap_width": round((pitch - bw) / bw * 100)}


def measure_bars(img: np.ndarray, vertical: bool = True) -> list[tuple[float, float, str]]:
    """Solid saturated bars in a chart picture: (position along the category axis,
    length in px, colour), sorted by position. Bars share a base line."""
    m = _saturated(img).astype(np.uint8)
    n, _lab, stats, _ = cv2.connectedComponentsWithStats(m, 8)
    bars = []
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if area < 30 or area < 0.85 * w * h or min(w, h) < 3:
            continue  # not a solid rectangle
        col = img[y + h // 2, x + w // 2]
        rgb = f"{int(col[2]):02X}{int(col[1]):02X}{int(col[0]):02X}"
        bars.append((x, y, w, h, rgb))
    if not bars:
        return []
    if vertical:
        base = max(b[1] + b[3] for b in bars)
        bars = [b for b in bars if abs(b[1] + b[3] - base) <= 3]
        return sorted((b[0] + b[2] / 2, float(b[3]), b[4]) for b in bars)
    base = min(b[0] for b in bars)
    bars = [b for b in bars if abs(b[0] - base) <= 3]
    return sorted((-(b[1] + b[3] / 2), float(b[2]), b[4]) for b in bars)


def from_picture(img_el: Image, crop: np.ndarray, vlm: dict[str, Any] | None, model: str,
                 cid: str) -> Chart | None:
    """A native chart for a chart picture, or None (not a chart / not readable)."""
    v = schemas.valid_chart(vlm)
    if v is None or v["chart_type"] == "other" or float(v["confidence"]) < MIN_ADOPT:
        return None
    ctype = v["chart_type"] if v["chart_type"] in (
        "column_clustered", "column_stacked", "bar_clustered", "bar_stacked", "line",
        "line_markers", "pie", "doughnut", "area") else "column_clustered"
    cats = [str(c)[:60] for c in v["categories"]]
    series = [ChartSeries(name=str(s.get("name") or f"Series {k + 1}")[:80],
                          values=[None if x is None else float(x) for x in s["values"]])
              for k, s in enumerate(v["series"])]
    notes = [f"vlm_chart:{model}"]
    fam = family(ctype)
    if fam in ("column", "bar") and len(series) == 1:
        orient = bar_orientation(crop)
        if orient is not None and orient != fam:  # the pixels decide
            notes.append(f"orientation:{fam}->{orient}")
            ctype = ctype.replace(fam, orient, 1)
            fam = orient
        bars = measure_bars(crop, vertical=fam == "column")
        geom = bar_geometry(crop, fam == "column")
        if len(bars) == len(cats):
            vmax = v.get("value_axis_max")
            known = [x for x in series[0].values if x is not None]
            ref = max(known) if known else None
            top = max(b[1] for b in bars)
            scale = (ref / top) if ref else ((float(vmax) / top) if vmax else 100.0 / top)
            series[0].values = [round(b[1] * scale, 4) for b in bars]
            series[0].color = Color(rgb=bars[0][2])
            notes.append("values:pixels")
    if series[0].color is None:
        cols = dominant_colors(crop, len(cats) if fam in ("pie", "doughnut") else len(series))
        if fam in ("pie", "doughnut"):
            series[0].point_colors = [Color(rgb=c) for c in cols] or None
        else:
            for s, c in zip(series, cols, strict=False):
                s.color = Color(rgb=c)
    ch = Chart(id=cid, bbox=img_el.bbox, chart_type=ctype, categories=cats, series=series,
               title=v.get("title") or None, has_legend=len(series) > 1 or fam == "pie",
               legend_pos="bottom", source=f"model:{model}",
               confidence=round(float(v["confidence"]), 3), fallback=[img_el],
               fallback_ids=[img_el.id], z=img_el.z, notes=notes)
    placeholder = all(re.fullmatch(r"(category|cat|item|series|group)?\s*\d+", c.strip(),
                                   re.I) for c in cats)
    if (fam in ("column", "bar") and len(series) == 1 and geom is not None
            and "values:pixels" in notes):
        # plot rectangle, axis maximum and gap width from the pixels
        sx = (img_el.bbox[2] - img_el.bbox[0]) / crop.shape[1]
        sy = (img_el.bbox[3] - img_el.bbox[1]) / crop.shape[0]
        x0, y0, x1, y1 = geom["plot"]
        ch.plot_bbox = (img_el.bbox[0] + x0 * sx, img_el.bbox[1] + y0 * sy,
                        img_el.bbox[0] + x1 * sx, img_el.bbox[1] + y1 * sy)
        vmax = max(x for x in series[0].values if x is not None)
        ch.axes = {"value": {"min": 0.0, "max": float(vmax), "major": None,
                             "format": "General"}}
        ch.style["gap_width"] = geom["gap_width"]
        ch.style["overlap"] = 0.0
        ch.notes.append("layout:pixels")
    if v.get("axis_labels_visible") is False or placeholder:
        ch.style |= {"hide_tick_labels": 1.0, "gridlines": "none"}
        ch.has_legend = ch.has_legend and fam in ("pie", "doughnut")
    lo, hi = v.get("value_axis_min"), v.get("value_axis_max")
    if (fam not in ("pie", "doughnut") and lo is not None and hi is not None and hi > lo
            and not ch.axes):
        ch.axes = {"value": {"min": float(lo), "max": float(hi), "major": None,
                             "format": "General"}}
    return ch


def png_bytes(img: np.ndarray) -> bytes:
    ok, buf = cv2.imencode(".png", img)
    return bytes(buf) if ok else b""


def decode_png(data: bytes) -> np.ndarray:
    arr = np.frombuffer(io.BytesIO(data).getvalue(), dtype=np.uint8)
    out = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    return out if out is not None else np.zeros((1, 1, 3), np.uint8)
