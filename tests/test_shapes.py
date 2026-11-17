"""Graphics analysis (M6): preset fitting, connectors, styling, background, groups, SVG.

The PowerPoint test draws every supported preset (with non-default adjust values,
rotation and flips), connectors with every arrowhead kind and a gradient, exports the
deck to PDF through PowerPoint, parses it and checks that analyze/shapes.py recovers
them. That validates both the preset formulas (analyze/presets.py) and the fitter.
"""

from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pytest
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Pt

from pdf2pptx.analyze import presets, shapes
from pdf2pptx.ingest import parse
from pdf2pptx.ir import Color, Fill, Group, Image, Page, PathSegment, Shape, Stroke

# --------------------------------------------------------------------------- helpers


def _path(pts: list[tuple[float, float]], closed: bool = True) -> list[PathSegment]:
    segs = [PathSegment(op="M", pts=[pts[0]])]
    segs += [PathSegment(op="L", pts=[p]) for p in pts[1:]]
    if closed:
        segs.append(PathSegment(op="Z"))
    return segs


def _shape(sid: str, pts: list[tuple[float, float]], *, fill: str | None = "4F81BD",
           stroke: Stroke | None = None, closed: bool = True, z: int = 1,
           notes: list[str] | None = None) -> Shape:
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return Shape(id=sid, bbox=(min(xs), min(ys), max(xs), max(ys)), z=z, path=_path(pts, closed),
                 fill=Fill(kind="solid", color=Color(rgb=fill)) if fill else Fill(),
                 stroke=stroke, notes=list(notes or []))


def _preset_path(preset: str, frame: tuple[float, float, float, float], adj: dict | None = None,
                 rot: float = 0.0, fh: bool = False, fv: bool = False) -> list[tuple[float, float]]:
    w, h = frame[2] - frame[0], frame[3] - frame[1]
    return presets.place(presets.outline(preset, w, h, adj), frame, rot, fh, fv)


# --------------------------------------------------------------------------- unit


@pytest.mark.parametrize("preset,adj,rot,fh,fv", [
    ("rect", None, 0.0, False, False),
    ("roundRect", {"adj": 30000}, 0.0, False, False),
    ("ellipse", None, 0.0, False, False),
    ("triangle", {"adj": 25000}, 0.0, False, False),
    ("rtTriangle", None, 0.0, True, False),
    ("diamond", None, 0.0, False, False),
    ("parallelogram", {"adj": 40000}, 0.0, False, False),
    ("trapezoid", {"adj": 15000}, 0.0, False, True),
    ("chevron", {"adj": 30000}, 0.0, True, False),
    ("homePlate", {"adj": 70000}, 0.0, False, False),
    ("rightArrow", {"adj1": 30000, "adj2": 80000}, 0.0, False, False),
    ("leftArrow", None, 0.0, False, False),
    ("upArrow", None, 0.0, False, False),
    ("downArrow", {"adj1": 60000, "adj2": 40000}, 0.0, False, False),
    ("rightArrow", None, 30.0, False, False),
    ("rect", None, 15.0, False, False),
])
def test_fit_recovers_preset(preset: str, adj: dict | None, rot: float, fh: bool,
                             fv: bool) -> None:
    frame = (100.0, 100.0, 300.0, 180.0)
    pts = _preset_path(preset, frame, adj, rot, fh, fv)
    res = shapes.fit_shape(_shape("s", pts))
    assert isinstance(res, Shape) and res.preset == preset, res
    assert abs(((res.rotation - rot + 180) % 360) - 180) < 0.5 or preset == "rect"
    for k, v in (adj or {}).items():
        assert abs(res.adjust[k] - v) < 1500, (k, res.adjust)
    assert res.source == "rule" and any(n.startswith(f"fit:{preset}") for n in res.notes)


def test_a_wavy_path_stays_a_freeform() -> None:
    pts = [(100 + 200 * t, 140 + 30 * math.sin(t * 12)) for t in np.linspace(0, 1, 60)]
    pts += [(300, 200), (100, 200)]
    assert shapes.fit_shape(_shape("w", pts)) is None


def test_raster_check_rejects_a_near_miss() -> None:
    # a 5-gon that is almost a homePlate but with its tip moved off-centre by 15 pt
    pts = [(100, 100), (250, 100), (300, 155), (250, 180), (100, 180)]
    assert shapes.fit_shape(_shape("h", pts)) is None


def test_pie_slice() -> None:
    frame = (100.0, 100.0, 300.0, 300.0)
    pts = _preset_path("pie", frame, {"adj1": 0, "adj2": 5400000})
    res = shapes.fit_shape(_shape("p", pts))
    assert isinstance(res, Shape) and res.preset == "pie"
    assert abs(res.adjust["adj2"] / 60000 - 90) < 1.0


def test_stroked_line_and_elbow_become_connectors() -> None:
    st = Stroke(color=Color(rgb="000000"), width=1.5, dash="custom", dash_pattern=[6.0, 4.5])
    line = shapes.fit_shape(_shape("l", [(300, 100), (100, 100)], fill=None, stroke=st,
                                   closed=False))
    assert isinstance(line, Shape) and line.preset == "line" and line.connector is not None
    assert line.stroke is not None and line.stroke.dash == "dash"  # 4w-3w -> preset
    elbow = shapes.fit_shape(_shape("e", [(100, 100), (200, 100), (200, 200), (300, 200)],
                                    fill=None, stroke=Stroke(width=1.0), closed=False))
    assert isinstance(elbow, Shape) and elbow.preset == "bentConnector3"
    assert elbow.adjust["adj1"] == 50000


def test_outlined_line_with_heads() -> None:
    """PowerPoint's export: a 2 pt body polygon + a triangle head at the end."""
    body = [(100, 101), (194, 101), (194, 99), (100, 99)]
    head = [(194, 103), (200, 100), (194, 97)]
    segs = _path(body) + _path(head)
    sh = Shape(id="c", bbox=(100, 97, 200, 103), path=segs,
               fill=Fill(kind="solid", color=Color(rgb="404040")))
    res = shapes.fit_shape(sh)
    assert isinstance(res, Shape) and res.preset == "line"
    assert res.stroke is not None and res.stroke.tail_end == "triangle"
    assert res.stroke.head_end is None and abs(res.stroke.width - 2.0) < 0.05
    assert res.bbox[2] == pytest.approx(200, abs=0.05)  # the line ends at the tip


def test_connector_attaches_to_sites() -> None:
    a = shapes.fit_shape(_shape("a", _preset_path("rect", (0, 0, 100, 50))))
    b = shapes.fit_shape(_shape("b", _preset_path("ellipse", (200, 0, 300, 50))))
    ln = shapes.fit_shape(_shape("l", [(100, 25), (200, 25)], fill=None,
                                 stroke=Stroke(width=1.0), closed=False))
    els = [a, b, ln]
    assert shapes.attach_connectors(els) == 2
    assert isinstance(ln, Shape) and ln.connector is not None
    assert (ln.connector.start_id, ln.connector.start_site) == ("a", 3)
    assert (ln.connector.end_id, ln.connector.end_site) == ("b", 2)


def test_fill_and_stroke_paths_merge() -> None:
    pts = _preset_path("rect", (10, 10, 60, 40))
    f = _shape("f", pts, z=1)
    s = _shape("s", pts, fill=None, stroke=Stroke(color=Color(rgb="FF0000"), width=2), z=2)
    els = [f, s]
    assert shapes.merge_fill_stroke(els) == 1
    assert els == [f] and f.stroke is not None and f.stroke.width == 2


def test_background_and_groups() -> None:
    page = Page(index=0, width=400, height=300)
    bg = _shape("bg", [(0, 0), (400, 0), (400, 300), (0, 300)], fill="1B2A41", z=1)
    a = _shape("a", _preset_path("rect", (10, 10, 60, 40)), z=2, notes=["mcid:P/1"])
    img = Image(id="i", bbox=(8, 12, 62, 44), z=3, file="x.png", notes=["mcid:P/1"])
    c = _shape("c", _preset_path("rect", (200, 10, 260, 40)), z=4)
    page.elements = [bg, a, img, c]
    assert shapes.detect_background(page) == "bg" and bg.is_background
    assert shapes.make_groups(page, {}) == 1
    g = next(e for e in page.elements if isinstance(e, Group))
    assert g.child_ids == ["a", "i"]


def test_group_skipped_when_something_is_drawn_between() -> None:
    page = Page(index=0, width=400, height=300)
    a = _shape("a", _preset_path("rect", (10, 10, 60, 40)), z=1, notes=["mcid:P/1"])
    mid = _shape("m", _preset_path("rect", (20, 20, 50, 30)), z=2)
    b = _shape("b", _preset_path("rect", (40, 10, 90, 40)), z=3, notes=["mcid:P/1"])
    page.elements = [a, mid, b]
    assert shapes.make_groups(page, {}) == 0


def test_dense_cluster_becomes_svg(tmp_path: Path) -> None:
    page = Page(index=0, width=400, height=300)
    for k in range(8):  # eight wavy strokes touching each other
        pts = [(50 + 4 * k + 30 * t, 50 + 10 * math.sin(t * 9 + k)) for t in np.linspace(0, 1, 20)]
        page.elements.append(_shape(f"w{k}", pts, fill=None, stroke=Stroke(width=1.0),
                                    closed=False, z=k + 1))
    assert shapes.svg_clusters(page, tmp_path) == 1
    img = page.elements[0]
    assert isinstance(img, Image) and img.svg_file and len(page.elements) == 1
    svg = (tmp_path / img.svg_file).read_text(encoding="utf-8")
    assert svg.count("<path ") == 8 and "viewBox" in svg
    assert (tmp_path / img.file).read_bytes()[:4] == b"\x89PNG"


def test_linear_gradient_fit() -> None:
    h, w = 60, 120
    xs = np.linspace(0, 1, w)[None, :].repeat(h, 0)
    rgba = np.zeros((h, w, 4), np.uint8)
    rgba[..., 0] = (255 * (1 - xs)).astype(np.uint8)  # red -> blue, left to right
    rgba[..., 2] = (255 * xs).astype(np.uint8)
    rgba[..., 3] = 255
    res = shapes.fit_gradient(rgba, np.ones((h, w), bool), (0, 0, 120, 60), (0, 0, 120, 60))
    assert res is not None
    g, err = res
    assert g.kind == "linear" and abs(g.angle) < 2.0 and err < 2.0
    assert g.stops[0].color.rgb[:2] in ("FF", "FE", "FD") and g.stops[-1].color.rgb[4:] >= "FA"


# --------------------------------------------------------------------------- PowerPoint


def _rgb(s: str) -> RGBColor:
    return RGBColor.from_string(s)


@pytest.mark.powerpoint
def test_presets_round_trip_through_powerpoint(tmp_path: Path) -> None:
    from pdf2pptx.verify import ppt_com

    cases = [  # (MSO_SHAPE, expected preset, adjustments, rotation, flipH)
        (MSO_SHAPE.RECTANGLE, "rect", [], 0, False),
        (MSO_SHAPE.ROUNDED_RECTANGLE, "roundRect", [0.3], 0, False),
        (MSO_SHAPE.OVAL, "ellipse", [], 0, False),
        (MSO_SHAPE.ISOSCELES_TRIANGLE, "triangle", [0.3], 0, False),
        (MSO_SHAPE.RIGHT_TRIANGLE, "rtTriangle", [], 0, True),
        (MSO_SHAPE.DIAMOND, "diamond", [], 0, False),
        (MSO_SHAPE.PARALLELOGRAM, "parallelogram", [0.4], 0, False),
        (MSO_SHAPE.TRAPEZOID, "trapezoid", [0.2], 0, False),
        (MSO_SHAPE.CHEVRON, "chevron", [0.3], 0, False),
        (MSO_SHAPE.PENTAGON, "homePlate", [0.7], 0, False),
        (MSO_SHAPE.RIGHT_ARROW, "rightArrow", [0.3, 0.8], 0, False),
        (MSO_SHAPE.LEFT_ARROW, "leftArrow", [], 0, False),
        (MSO_SHAPE.UP_ARROW, "upArrow", [], 0, False),
        (MSO_SHAPE.DOWN_ARROW, "downArrow", [0.6, 0.4], 0, False),
        (MSO_SHAPE.RIGHT_ARROW, "rightArrow", [], 30, False),
        (MSO_SHAPE.RECTANGLE, "rect", [], 20, False),
    ]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(960), Pt(540)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    for k, (kind, _, adjs, rot, fh) in enumerate(cases):
        x, y = 40 + (k % 4) * 230, 30 + (k // 4) * 125
        shp = s.shapes.add_shape(kind, Pt(x), Pt(y), Pt(170), Pt(80))
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb("4F81BD")
        shp.line.fill.background()
        shp.shadow.inherit = False
        for i, v in enumerate(adjs):
            shp.adjustments[i] = v
        if rot:
            shp.rotation = rot
        if fh:
            shp._element.spPr.find(qn("a:xfrm")).set("flipH", "1")
    pptx, pdf = tmp_path / "p.pptx", tmp_path / "p.pdf"
    prs.save(str(pptx))
    eng = ppt_com.shared()
    info = eng.open(pptx)
    eng.export_pdf(info, pdf)
    eng.close(info)
    page = parse.parse_pages(pdf, [0], tmp_path / "w")[0][0]
    shapes.analyze_page(page, tmp_path / "w")
    got = sorted((e for e in page.elements if isinstance(e, Shape) and not e.is_background),
                 key=lambda e: (round(((e.bbox[1] + e.bbox[3]) / 2 - 30) // 125),
                                (e.bbox[0] + e.bbox[2]) / 2))
    assert [e.preset for e in got] == [c[1] for c in cases]
    for e, (_, preset, adjs, rot, _) in zip(got, cases, strict=True):
        assert abs(((e.rotation - rot + 180) % 360) - 180) < 0.5 or preset == "rect"
        names = sorted(presets.DEFAULTS[preset])
        for name, v in zip(names, adjs, strict=False):
            assert abs(e.adjust[name] - v * 100000) < 2000, (preset, e.adjust)


@pytest.mark.powerpoint
def test_connectors_and_gradient_round_trip(tmp_path: Path) -> None:
    from pdf2pptx.verify import ppt_com

    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(960), Pt(540)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    ends = [("triangle", None), ("stealth", "oval"), ("diamond", None), ("arrow", None)]
    for k, (tail, head) in enumerate(ends):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(60), Pt(60 + 60 * k), Pt(400),
                                   Pt(60 + 60 * k))
        c.line.width = Pt(2)
        c.line.color.rgb = _rgb("404040")
        ln = c.line._get_or_add_ln()
        if head:
            etree.SubElement(ln, qn("a:headEnd")).set("type", head)
        etree.SubElement(ln, qn("a:tailEnd")).set("type", tail)
    g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(500), Pt(60), Pt(300), Pt(150))
    g.fill.gradient()
    g.fill.gradient_angle = 0
    g.fill.gradient_stops[0].color.rgb = _rgb("FF0000")
    g.fill.gradient_stops[1].color.rgb = _rgb("0000FF")
    g.line.fill.background()
    g.shadow.inherit = False
    pptx, pdf = tmp_path / "c.pptx", tmp_path / "c.pdf"
    prs.save(str(pptx))
    eng = ppt_com.shared()
    info = eng.open(pptx)
    eng.export_pdf(info, pdf)
    eng.close(info)
    wd = tmp_path / "w"
    page = parse.parse_pages(pdf, [0], wd)[0][0]
    shapes.analyze_page(page, wd)
    lines = sorted((e for e in page.elements if isinstance(e, Shape) and e.preset == "line"),
                   key=lambda e: e.bbox[1])
    assert len(lines) == 4
    for e, (tail, head) in zip(lines, ends, strict=True):
        assert e.stroke is not None
        assert e.stroke.tail_end == tail and e.stroke.head_end == head, e.stroke
        assert abs(e.stroke.width - 2.0) < 0.1
    grad = [e for e in page.elements if isinstance(e, Shape) and e.fill.kind == "gradient"]
    assert len(grad) == 1 and grad[0].preset == "rect"
    gr = grad[0].fill.gradient
    assert gr is not None and gr.kind == "linear" and abs(gr.angle) < 3
    assert not any(isinstance(e, Image) and "raster:pattern_fill" in e.notes
                   for e in page.elements)  # the raster was replaced
