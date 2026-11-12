"""Synthetic decks built with python-pptx: the answer keys of the corpus.

Each builder gets a fresh Presentation and a seeded RNG and adds slides. Builders
are registered with ``@deck(name, features, aspect)``; names are stable and
become file names. Nothing here depends on wall-clock time.
"""

from __future__ import annotations

import random
import re
import uuid
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import assets

W169 = 12_192_000
W43 = 9_144_000
H = 6_858_000
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
UUID_NS = uuid.UUID("6f1c1d2e-7d38-4b8e-9a51-2f2d6c0d9a10")

Builder = Callable[[Any, random.Random, "AssetSet"], None]


@dataclass(frozen=True)
class DeckSpec:
    name: str
    features: tuple[str, ...]
    aspect: str
    builder: Builder
    seed: int


@dataclass(frozen=True)
class AssetSet:
    photo: Path
    photo2: Path
    logo: Path
    logo_red: Path


DECKS: list[DeckSpec] = []


def deck(name: str, features: tuple[str, ...], aspect: str = "16:9") -> Callable[[Builder], Builder]:
    def wrap(fn: Builder) -> Builder:
        DECKS.append(DeckSpec(name, features, aspect, fn, seed=1000 + len(DECKS)))
        return fn

    return wrap


def make_assets(dir_: Path) -> AssetSet:
    return AssetSet(
        photo=assets.photo(dir_ / "photo_a.jpg", seed=7),
        photo2=assets.photo(dir_ / "photo_b.jpg", w=480, h=480, seed=11),
        logo=assets.logo(dir_ / "logo_blue.png"),
        logo_red=assets.logo(dir_ / "logo_red.png", color=(192, 80, 77)),
    )


# --------------------------------------------------------------------------- helpers


def new_presentation(aspect: str) -> Any:
    prs = Presentation()
    prs.core_properties.title = ""
    if aspect == "16:9":
        f = W169 / W43
        prs.slide_width = Emu(W169)
        prs.slide_height = Emu(H)
        master = prs.slide_master
        # Read every geometry (possibly inherited from the master) before writing any:
        # setting one value creates an xfrm whose other values would fall back to 0,
        # and scaling the master first would double-scale inheriting layouts.
        todo = []
        for container in [master, *master.slide_layouts]:
            for shp in container.shapes:
                geo = (shp.left, shp.top, shp.width, shp.height)
                if None not in geo:
                    todo.append((shp, geo))
        for shp, (left, top, width, height) in todo:
            shp.left = Emu(round(left * f))
            shp.top = Emu(top)
            shp.width = Emu(round(width * f))
            shp.height = Emu(height)
    else:
        prs.slide_width = Emu(W43)
        prs.slide_height = Emu(H)
    return prs


def stable_guid(*parts: object) -> str:
    return "{" + str(uuid.uuid5(UUID_NS, "/".join(map(str, parts)))).upper() + "}"


def rgb(hex_: str) -> RGBColor:
    return RGBColor.from_string(hex_)


def blank(prs: Any) -> Any:
    return prs.slides.add_slide(prs.slide_layouts[6])


def set_title(slide: Any, text: str) -> None:
    slide.shapes.title.text = text


def textbox(
    slide: Any,
    x: float,
    y: float,
    w: float,
    h: float,
    paras: list[Any],
    *,
    font: str | None = None,
    size: float = 18,
    color: str | None = None,
    align: Any = None,
    wrap: bool = True,
    anchor: Any = None,
) -> Any:
    """Add a text box (inches). ``paras`` items: str, or dict(text|runs, level, bullet, ...)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    if anchor is not None:
        tf.vertical_anchor = anchor
    for i, spec in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(spec, str):
            spec = {"text": spec}
        runs = spec.get("runs") or [{"text": spec.get("text", "")}]
        for rs in runs:
            r = p.add_run()
            r.text = rs["text"]
            f = r.font
            f.size = Pt(rs.get("size", spec.get("size", size)))
            name = rs.get("font", spec.get("font", font))
            if name:
                f.name = name
            c = rs.get("color", spec.get("color", color))
            if c:
                f.color.rgb = rgb(c)
            if rs.get("bold", spec.get("bold")):
                f.bold = True
            if rs.get("italic", spec.get("italic")):
                f.italic = True
            if rs.get("underline"):
                f.underline = True
            if rs.get("link"):
                r.hyperlink.address = rs["link"]
            if rs.get("baseline"):
                r._r.get_or_add_rPr().set("baseline", str(rs["baseline"]))
        a = spec.get("align", align)
        if a is not None:
            p.alignment = a
        if spec.get("level"):
            p.level = spec["level"]
        if spec.get("bullet"):
            set_bullet(p, spec["bullet"], level=spec.get("level", 0), size=spec.get("size", size))
        if "line_spacing" in spec:
            p.line_spacing = spec["line_spacing"]
        if "space_before" in spec:
            p.space_before = Pt(spec["space_before"])
        if "space_after" in spec:
            p.space_after = Pt(spec["space_after"])
    return tb


def set_bullet(p: Any, char: str, level: int = 0, size: float = 18) -> None:
    pPr = p._p.get_or_add_pPr()
    indent = Pt(size * 1.1)
    pPr.set("marL", str(int(indent * (level + 1))))
    pPr.set("indent", str(-int(indent)))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if char.startswith("auto:"):
        bu = etree.SubElement(pPr, qn("a:buAutoNum"))
        bu.set("type", char[5:])
    else:
        bu = etree.SubElement(pPr, qn("a:buChar"))
        bu.set("char", char)


def bullets_placeholder(slide: Any, items: list[tuple[int, str]], idx: int = 1) -> None:
    body = slide.placeholders[idx].text_frame
    for i, (level, text) in enumerate(items):
        p = body.paragraphs[0] if i == 0 else body.add_paragraph()
        p.text = text
        p.level = level


def set_alpha(shape: Any, alpha: float) -> None:
    for clr in shape._element.spPr.iter(qn("a:srgbClr")):
        for old in clr.findall(qn("a:alpha")):
            clr.remove(old)
        a = etree.SubElement(clr, qn("a:alpha"))
        a.set("val", str(round(alpha * 100_000)))


def autoshape(
    slide: Any,
    kind: Any,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: str | None = None,
    line: str | None = None,
    line_w: float | None = None,
    text: str | None = None,
    rotation: float = 0.0,
    text_color: str = "FFFFFF",
    size: float = 16,
) -> Any:
    shp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        if line_w:
            shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    if rotation:
        shp.rotation = rotation
    if text:
        tf = shp.text_frame
        tf.text = text
        tf.word_wrap = True
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.size = Pt(size)
                r.font.color.rgb = rgb(text_color)
    return shp


def arrow_ends(connector: Any, head: str | None = None, tail: str | None = "triangle") -> None:
    ln = connector.line._get_or_add_ln()
    if head:
        etree.SubElement(ln, qn("a:headEnd")).set("type", head)
    if tail:
        etree.SubElement(ln, qn("a:tailEnd")).set("type", tail)


def slide_number_box(slide: Any, number: int, x: float, y: float, size: float = 12) -> Any:
    tb = textbox(slide, x, y, 0.8, 0.35, [""], size=size, color="595959", align=PP_ALIGN.RIGHT)
    p = tb.text_frame.paragraphs[0]._p
    for r in p.findall(qn("a:r")):
        p.remove(r)
    fld = etree.SubElement(p, qn("a:fld"))
    fld.set("id", stable_guid("sldnum", number))
    fld.set("type", "slidenum")
    rpr = etree.SubElement(fld, qn("a:rPr"))
    rpr.set("lang", "it-IT")
    rpr.set("sz", str(int(size * 100)))
    solid = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(solid, qn("a:srgbClr")).set("val", "595959")
    etree.SubElement(fld, qn("a:t")).text = str(number)
    # endParaRPr must stay last
    end = p.find(qn("a:endParaRPr"))
    if end is not None:
        p.remove(end)
        p.append(end)
    return tb


def add_sections(prs: Any, sections: list[tuple[str, list[int]]]) -> None:
    """Write p14:sectionLst. ``sections`` = [(name, [slide indices 0-based])]."""
    pres = prs.part._element
    sld_ids = [s.get("id") for s in pres.find(qn("p:sldIdLst"))]
    ext_lst = pres.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(pres, qn("p:extLst"))
    ext = etree.SubElement(ext_lst, qn("p:ext"))
    ext.set("uri", SECTION_EXT_URI)
    lst = etree.SubElement(ext, f"{{{NS_P14}}}sectionLst", nsmap={"p14": NS_P14})
    for name, idxs in sections:
        sec = etree.SubElement(lst, f"{{{NS_P14}}}section")
        sec.set("name", name)
        sec.set("id", stable_guid("section", name))
        ids = etree.SubElement(sec, f"{{{NS_P14}}}sldIdLst")
        for i in idxs:
            etree.SubElement(ids, f"{{{NS_P14}}}sldId").set("id", sld_ids[i])


def set_theme_colors(prs: Any, colors: dict[str, str]) -> None:
    for part in prs.part.package.iter_parts():
        if re.search(r"/theme\d*\.xml$", str(part.partname)):
            xml = part.blob.decode("utf-8")
            for key, val in colors.items():
                xml = re.sub(
                    rf"(<a:{key}>)\s*<a:(?:srgbClr|sysClr)[^>]*/>",
                    rf'\1<a:srgbClr val="{val}"/>',
                    xml,
                )
            part._blob = xml.encode("utf-8")


def notes(slide: Any, text: str) -> None:
    slide.notes_slide.notes_text_frame.text = text


def add_chart(
    slide: Any,
    kind: Any,
    cats: list[str],
    series: list[tuple[str, list[float]]],
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    legend: bool = True,
    labels: bool = False,
    title: str | None = None,
) -> Any:
    data = CategoryChartData()
    data.categories = cats
    for name, vals in series:
        data.add_series(name, vals)
    gf = slide.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), data)
    chart = gf.chart
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    if labels:
        chart.plots[0].has_data_labels = True
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    else:
        chart.has_title = False
    return gf


def table(
    slide: Any,
    rows: list[list[str]],
    x: float,
    y: float,
    w: float,
    h: float,
    size: float = 14,
) -> Any:
    gf = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = gf.table
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)
    return gf


# --------------------------------------------------------------------------- content

LOREM_EN = (
    "Our programs reached more communities this year than in any year before. "
    "Volunteers logged thousands of hours, partners opened new sites, and the team "
    "rebuilt the intake process so that families wait less and get help sooner."
)
LOREM_IT = (
    "Quest'anno le attività del museo hanno coinvolto più scuole e famiglie che mai. "
    "Le visite guidate, i laboratori didattici e le aperture serali hanno portato "
    "nuovi visitatori, e la città ha risposto con entusiasmo alle iniziative più originali."
)


# --------------------------------------------------------------------------- decks


@deck("d01_title_bullets", ("title", "bullets", "placeholders"))
def d01(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "Annual Review 2026"
    s.placeholders[1].text = "Programs, people and the year ahead"
    s = prs.slides.add_slide(prs.slide_layouts[1])
    set_title(s, "What we achieved")
    bullets_placeholder(
        s,
        [
            (0, "Served 4,200 families across three districts"),
            (0, "Opened two new community sites"),
            (0, "Trained 180 new volunteers"),
            (0, "Cut average waiting time from 12 to 5 days"),
        ],
    )
    s = prs.slides.add_slide(prs.slide_layouts[1])
    set_title(s, "Priorities for next year")
    bullets_placeholder(
        s,
        [
            (0, "Grow the after-school program"),
            (0, "Strengthen partnerships with local schools"),
            (0, "Launch a monthly giving program"),
        ],
    )


@deck("d02_multilevel_bullets", ("bullets", "multilevel", "numbering"))
def d02(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[1])
    set_title(s, "Project structure")
    bullets_placeholder(
        s,
        [
            (0, "Research"),
            (1, "Interviews with 40 partners"),
            (2, "Transcribed and coded"),
            (1, "Survey of 1,100 visitors"),
            (0, "Design"),
            (1, "Three prototypes"),
            (2, "Paper prototype"),
            (3, "Tested with staff"),
            (0, "Delivery"),
        ],
    )
    s = blank(prs)
    textbox(s, 0.8, 0.4, 11, 1.0, ["Steps to apply"], size=36, font="Calibri Light")
    textbox(
        s,
        1.0,
        1.6,
        10,
        4.5,
        [
            {"text": "Read the guidelines", "bullet": "auto:arabicPeriod"},
            {"text": "Check eligibility", "bullet": "auto:arabicPeriod"},
            {"text": "Budget and timeline", "bullet": "auto:alphaLcParenR", "level": 1},
            {"text": "Partner letters", "bullet": "auto:alphaLcParenR", "level": 1},
            {"text": "Submit before the deadline", "bullet": "auto:arabicPeriod"},
            {"text": "Wait for the decision", "bullet": "–"},
        ],
        size=22,
    )


@deck("d03_two_columns", ("two_column", "placeholders"))
def d03(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[3])
    set_title(s, "Before and after")
    bullets_placeholder(
        s, [(0, "Paper forms"), (0, "Phone queue"), (0, "Weekly reports by hand")], idx=1
    )
    bullets_placeholder(
        s, [(0, "Online form"), (0, "Text reminders"), (0, "Live dashboard")], idx=2
    )
    s = blank(prs)
    textbox(s, 0.7, 0.4, 12, 0.9, ["Two columns of running text"], size=32)
    textbox(s, 0.7, 1.5, 5.7, 5.2, [LOREM_EN, LOREM_EN], size=16, align=PP_ALIGN.LEFT)
    textbox(s, 6.9, 1.5, 5.7, 5.2, [LOREM_EN, LOREM_EN], size=16, align=PP_ALIGN.LEFT)


@deck("d04_three_columns", ("multi_column", "wrapping"))
def d04(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.4, 12, 0.9, ["Three program areas"], size=32)
    heads = ["Education", "Health", "Housing"]
    for i, head in enumerate(heads):
        x = 0.6 + i * 4.1
        textbox(
            s,
            x,
            1.5,
            3.8,
            5.2,
            [
                {"text": head, "bold": True, "size": 20, "space_after": 6},
                {"text": LOREM_EN, "size": 14},
            ],
            size=14,
        )


@deck("d05_mixed_runs", ("mixed_fonts", "mixed_sizes", "colors", "superscript"))
def d05(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.4, 12, 0.9, ["Mixed formatting in one paragraph"], size=30)
    textbox(
        s,
        0.6,
        1.5,
        12,
        2.0,
        [
            {
                "runs": [
                    {"text": "Normal text, "},
                    {"text": "bold text, ", "bold": True},
                    {"text": "italic text, ", "italic": True},
                    {"text": "red text, ", "color": "C00000"},
                    {"text": "larger text, ", "size": 28},
                    {"text": "Georgia text", "font": "Georgia"},
                    {"text": " and "},
                    {"text": "underlined", "underline": True},
                    {"text": "."},
                ]
            }
        ],
        size=20,
    )
    textbox(
        s,
        0.6,
        3.6,
        12,
        1.0,
        [
            {
                "runs": [
                    {"text": "E = mc"},
                    {"text": "2", "baseline": 30000},
                    {"text": " and H"},
                    {"text": "2", "baseline": -25000},
                    {"text": "O are classic examples."},
                ]
            }
        ],
        size=22,
    )
    for i, (col, sz) in enumerate([("1F497D", 12), ("4F81BD", 16), ("9BBB59", 24), ("F79646", 32)]):
        textbox(s, 0.6 + i * 3.0, 5.0, 2.9, 1.4, [f"{sz} pt"], size=sz, color=col)


@deck("d06_italian_text", ("italian", "accents", "bullets"))
def d06(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[1])
    set_title(s, "Perché visitare il museo?")
    bullets_placeholder(
        s,
        [
            (0, "Collezioni uniche: più di 3.000 opere d'arte"),
            (0, "Attività per bambini e famiglie, ogni sabato"),
            (1, "Laboratori creativi già prenotabili"),
            (0, "Biglietto ridotto: 5 € per gli studenti"),
            (0, "Città, territorio e identità: «il nostro patrimonio»"),
        ],
    )
    s = blank(prs)
    textbox(s, 0.7, 0.4, 12, 0.9, ["Il programma dell'anno"], size=32)
    textbox(s, 0.7, 1.5, 11.5, 4.5, [LOREM_IT, "Perché è così? Però è già così: àèéìòù ÀÈÉÌÒÙ."],
            size=18, align=PP_ALIGN.JUSTIFY)


@deck("d07_table_simple", ("table",))
def d07(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Visitors by quarter")
    rows = [["Site", "Q1", "Q2", "Q3", "Q4"]]
    for site in ["North", "South", "East", "West"]:
        rows.append([site, *(f"{rng.randint(800, 4000):,}" for _ in range(4))])
    table(s, rows, 1.0, 1.8, 11.3, 3.5, size=16)


@deck("d08_table_merged", ("table", "merged_cells"))
def d08(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Budget overview")
    rows = [
        ["Area", "2025", "", "2026", ""],
        ["", "Plan", "Actual", "Plan", "Actual"],
        ["Programs", "120", "118", "135", "—"],
        ["Staff", "80", "84", "88", "—"],
        ["Operations", "30", "27", "32", "—"],
        ["Total", "230", "229", "255", "—"],
    ]
    gf = table(s, rows, 1.0, 1.7, 11.3, 4.2, size=15)
    t = gf.table
    t.cell(0, 1).merge(t.cell(0, 2))
    t.cell(0, 3).merge(t.cell(0, 4))
    t.cell(0, 0).merge(t.cell(1, 0))
    t.cell(0, 1).text = "2025"
    t.cell(0, 3).text = "2026"
    t.cell(0, 0).text = "Area"


@deck("d09_shapes_basic", ("shapes", "preset", "shape_text"))
def d09(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["Basic shapes"], size=30)
    autoshape(s, MSO_SHAPE.RECTANGLE, 0.8, 1.6, 2.6, 1.6, fill="4F81BD", text="Rectangle")
    autoshape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 3.9, 1.6, 2.6, 1.6, fill="C0504D", text="Rounded")
    autoshape(s, MSO_SHAPE.OVAL, 7.0, 1.6, 2.6, 1.6, fill="9BBB59", text="Ellipse")
    autoshape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, 10.1, 1.6, 2.4, 1.6, fill="8064A2")
    autoshape(s, MSO_SHAPE.DIAMOND, 0.8, 4.2, 2.6, 2.0, fill="4BACC6")
    autoshape(s, MSO_SHAPE.RIGHT_TRIANGLE, 3.9, 4.2, 2.2, 2.0, fill="F79646")
    autoshape(s, MSO_SHAPE.PARALLELOGRAM, 6.6, 4.2, 2.8, 1.6, fill="1F497D", text="Para")
    autoshape(s, MSO_SHAPE.TRAPEZOID, 9.9, 4.2, 2.6, 1.6, line="000000", line_w=2)


@deck("d10_arrows_chevrons", ("shapes", "arrows", "chevron", "rotation"))
def d10(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["Process"], size=30)
    labels = ["Plan", "Build", "Test", "Launch"]
    for i, lab in enumerate(labels):
        kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        autoshape(s, kind, 0.8 + i * 3.0, 1.8, 3.2, 1.2, fill=["1F497D", "4F81BD", "4BACC6", "9BBB59"][i],
                  text=lab, size=18)
    autoshape(s, MSO_SHAPE.RIGHT_ARROW, 0.8, 4.0, 3.0, 1.0, fill="C0504D")
    autoshape(s, MSO_SHAPE.LEFT_ARROW, 4.3, 4.0, 3.0, 1.0, fill="F79646")
    autoshape(s, MSO_SHAPE.UP_ARROW, 8.0, 3.7, 1.0, 2.0, fill="8064A2")
    autoshape(s, MSO_SHAPE.RIGHT_ARROW, 9.6, 4.2, 2.6, 0.9, fill="9BBB59", rotation=30)
    autoshape(s, MSO_SHAPE.RECTANGLE, 1.0, 5.6, 2.0, 1.0, fill="4F81BD", rotation=15)


@deck("d11_connectors", ("connectors", "arrowheads", "shapes"))
def d11(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["How a request flows"], size=30)
    boxes = []
    for i, lab in enumerate(["Request", "Review", "Approve"]):
        boxes.append(
            autoshape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.9 + i * 4.2, 2.0, 2.8, 1.3, fill="4F81BD",
                      text=lab, size=18)
        )
    for b0, b1 in zip(boxes, boxes[1:], strict=False):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 0, 0)
        c.begin_connect(b0, 3)
        c.end_connect(b1, 1)
        c.line.color.rgb = rgb("404040")
        c.line.width = Pt(2)
        arrow_ends(c)
    done = autoshape(s, MSO_SHAPE.OVAL, 9.3, 4.8, 2.2, 1.3, fill="9BBB59", text="Done", size=18)
    c = s.shapes.add_connector(MSO_CONNECTOR.ELBOW, 0, 0, 0, 0)
    c.begin_connect(boxes[1], 2)
    c.end_connect(done, 1)
    c.line.color.rgb = rgb("C0504D")
    c.line.width = Pt(1.5)
    arrow_ends(c, head="oval", tail="stealth")
    line = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.9), Inches(6.6), Inches(8.0),
                                  Inches(6.6))
    line.line.color.rgb = rgb("7F7F7F")
    line.line.dash_style = MSO_LINE.DASH


@deck("d12_gradients_alpha", ("gradient", "transparency", "overlap"))
def d12(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["Gradients and transparency"], size=30)
    shp = autoshape(s, MSO_SHAPE.RECTANGLE, 0.8, 1.6, 5.5, 2.4)
    shp.fill.gradient()
    shp.fill.gradient_angle = 0
    st = shp.fill.gradient_stops
    st[0].color.rgb = rgb("1F497D")
    st[0].position = 0.0
    st[1].color.rgb = rgb("4BACC6")
    st[1].position = 1.0
    shp2 = autoshape(s, MSO_SHAPE.OVAL, 6.8, 1.6, 2.6, 2.6)
    shp2.fill.gradient()
    shp2.fill.gradient_angle = 90
    st = shp2.fill.gradient_stops
    st[0].color.rgb = rgb("F79646")
    st[1].color.rgb = rgb("C0504D")
    for i, col in enumerate(["C0504D", "9BBB59", "4F81BD"]):
        c = autoshape(s, MSO_SHAPE.OVAL, 1.2 + i * 1.5, 4.4, 2.4, 2.4, fill=col)
        set_alpha(c, 0.5)
    r = autoshape(s, MSO_SHAPE.RECTANGLE, 7.0, 4.6, 5.0, 1.8, fill="000000", text="50% black",
                  size=20)
    set_alpha(r, 0.5)


@deck("d13_images", ("images", "crop", "caption", "png_alpha"))
def d13(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["Pictures"], size=30)
    s.shapes.add_picture(str(a.photo), Inches(0.8), Inches(1.5), Inches(6.0))
    pic = s.shapes.add_picture(str(a.photo2), Inches(7.4), Inches(1.5), Inches(3.0), Inches(3.0))
    pic.crop_left = 0.1
    pic.crop_right = 0.1
    s.shapes.add_picture(str(a.logo), Inches(11.0), Inches(1.5), Inches(1.5))
    textbox(s, 0.8, 5.6, 6.0, 0.5, ["Figure 1: The new reading room"], size=14,
            color="595959")
    s = prs.slides.add_slide(prs.slide_layouts[8])
    set_title(s, "Picture with caption")
    s.placeholders[1].insert_picture(str(a.photo))
    s.placeholders[2].text = "The courtyard after the restoration work."


@deck("d14_chart_bar", ("chart", "bar"))
def d14(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Visitors by site")
    cats = ["North", "South", "East", "West"]
    series = [(str(y), [float(rng.randint(10, 60)) for _ in cats]) for y in (2024, 2025, 2026)]
    add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cats, series, 0.8, 1.5, 11.5, 5.5)
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Share of budget")
    add_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, ["Programs", "Staff", "Operations", "Outreach"],
              [("Budget", [52.0, 31.0, 11.0, 6.0])], 0.8, 1.5, 11.5, 5.5, legend=False,
              labels=True)


@deck("d15_chart_line", ("chart", "line"))
def d15(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Monthly members")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"]
    v1, v2 = 100.0, 80.0
    s1, s2 = [], []
    for _ in months:
        v1 += rng.randint(-5, 15)
        v2 += rng.randint(-3, 10)
        s1.append(v1)
        s2.append(v2)
    add_chart(s, XL_CHART_TYPE.LINE_MARKERS, months, [("Members", s1), ("Donors", s2)],
              0.8, 1.5, 11.5, 5.5)


@deck("d16_chart_pie", ("chart", "pie", "data_labels"))
def d16(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Where funding comes from")
    add_chart(s, XL_CHART_TYPE.PIE, ["Individuals", "Foundations", "Government", "Events"],
              [("Funding", [42.0, 28.0, 20.0, 10.0])], 2.5, 1.5, 8.0, 5.5, labels=True)


@deck("d17_footer_logo_number", ("repeated_footer", "logo", "slide_number", "master_candidate"))
def d17(prs: Any, rng: random.Random, a: AssetSet) -> None:
    topics = ["Introduction", "Context", "Approach", "Results", "Next steps"]
    for i, t in enumerate(topics, start=1):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        set_title(s, t)
        bullets_placeholder(s, [(0, f"{t}: key point one"), (0, f"{t}: key point two")])
        s.shapes.add_picture(str(a.logo), Inches(12.35), Inches(0.2), Inches(0.75))
        textbox(s, 0.6, 6.95, 8.0, 0.35, ["Community Report 2026 · Confidential"], size=11,
                color="7F7F7F")
        slide_number_box(s, i, 11.9, 6.95)


@deck("d18_buildup", ("build_sequence",))
def d18(prs: Any, rng: random.Random, a: AssetSet) -> None:
    items = [
        "Listen to the community",
        "Design with partners",
        "Pilot in one district",
        "Measure what changes",
        "Scale what works",
    ]
    for n in range(1, len(items) + 1):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        set_title(s, "Our approach")
        bullets_placeholder(s, [(0, t) for t in items[:n]])


@deck("d19_hyperlinks", ("hyperlinks", "internal_links"))
def d19(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s1 = blank(prs)
    textbox(s1, 0.6, 0.3, 12, 0.9, ["Useful links"], size=30)
    textbox(
        s1,
        0.8,
        1.6,
        11,
        3.0,
        [
            {"runs": [{"text": "Website: "}, {"text": "example.org", "link": "https://example.org/"}]},
            {"runs": [{"text": "Docs: "}, {"text": "docs.example.org/guide",
                                           "link": "https://docs.example.org/guide"}]},
            {"runs": [{"text": "Write to us: "}, {"text": "info@example.org",
                                                  "link": "mailto:info@example.org"}]},
        ],
        size=22,
    )
    btn = autoshape(s1, MSO_SHAPE.ROUNDED_RECTANGLE, 0.8, 5.0, 3.0, 1.0, fill="4F81BD",
                    text="Go to details", size=18)
    s2 = blank(prs)
    textbox(s2, 0.6, 0.3, 12, 0.9, ["Details"], size=30)
    textbox(s2, 0.8, 1.6, 11, 2.0, [LOREM_EN], size=18)
    btn.click_action.target_slide = s2
    ext = autoshape(s2, MSO_SHAPE.RECTANGLE, 0.8, 5.0, 3.0, 1.0, fill="9BBB59", text="Open site",
                    size=18)
    ext.click_action.hyperlink.address = "https://example.org/details"


@deck("d20_speaker_notes", ("speaker_notes",))
def d20(prs: Any, rng: random.Random, a: AssetSet) -> None:
    for i, t in enumerate(["Welcome", "The problem", "Our answer"], start=1):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        set_title(s, t)
        bullets_placeholder(s, [(0, f"Point {i}.1"), (0, f"Point {i}.2")])
        notes(s, f"Speaker notes for slide {i}: say thank you and introduce '{t}'.")


@deck("d21_sections", ("sections",))
def d21(prs: Any, rng: random.Random, a: AssetSet) -> None:
    titles = ["Agenda", "Where we are", "Numbers", "Stories", "Plans", "Questions"]
    for t in titles:
        s = prs.slides.add_slide(prs.slide_layouts[5] if t != "Agenda" else prs.slide_layouts[2])
        set_title(s, t)
        if t == "Agenda":
            s.placeholders[1].text = "Introduction, results and plans"
    add_sections(prs, [("Introduction", [0, 1]), ("Results", [2, 3]), ("Plans", [4, 5])])


@deck("d22_aspect_4x3", ("aspect_4_3", "bullets", "shapes"), aspect="4:3")
def d22(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "A 4:3 presentation"
    s.placeholders[1].text = "Classic slide size"
    s = prs.slides.add_slide(prs.slide_layouts[1])
    set_title(s, "Key facts")
    bullets_placeholder(s, [(0, "Founded in 1998"), (0, "35 staff, 400 volunteers"),
                            (1, "Most of them part-time")])
    s = blank(prs)
    autoshape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.8, 1.0, 3.8, 2.2, fill="4F81BD", text="Mission")
    autoshape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 5.4, 1.0, 3.8, 2.2, fill="C0504D", text="Vision")
    autoshape(s, MSO_SHAPE.OVAL, 3.1, 4.0, 3.8, 2.6, fill="9BBB59", text="Values")


@deck("d23_long_paragraphs", ("justified", "line_spacing", "space_before_after"))
def d23(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["A page of text"], size=30)
    textbox(
        s,
        0.8,
        1.4,
        11.6,
        5.6,
        [
            {"text": LOREM_EN, "align": PP_ALIGN.JUSTIFY, "space_after": 12},
            {"text": LOREM_EN, "align": PP_ALIGN.JUSTIFY, "line_spacing": 1.3, "space_after": 12},
            {"text": LOREM_EN, "align": PP_ALIGN.LEFT, "line_spacing": Pt(26), "space_before": 6},
        ],
        size=16,
    )
    s = blank(prs)
    textbox(s, 0.8, 0.8, 11.6, 1.0, ["Centered heading"], size=36, align=PP_ALIGN.CENTER)
    textbox(s, 2.0, 2.2, 9.3, 3.5, [LOREM_IT], size=20, align=PP_ALIGN.CENTER)
    textbox(s, 0.8, 6.2, 11.6, 0.6, ["Right aligned note"], size=14, align=PP_ALIGN.RIGHT)


@deck("d24_font_families", ("fonts", "serif", "mono"))
def d24(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["Font families"], size=30)
    fams = ["Arial", "Times New Roman", "Georgia", "Verdana", "Consolas", "Segoe UI", "Cambria",
            "Calibri"]
    for i, fam in enumerate(fams):
        textbox(s, 0.8 + (i % 2) * 6.2, 1.4 + (i // 2) * 1.35, 6.0, 1.2,
                [f"{fam}: The quick brown fox jumps over the lazy dog"], font=fam, size=16)


@deck("d25_dark_background", ("background", "dark", "theme_colors"))
def d25(prs: Any, rng: random.Random, a: AssetSet) -> None:
    set_theme_colors(prs, {"accent1": "00B0F0", "accent2": "FFC000", "accent3": "92D050",
                           "accent4": "FF5050", "accent5": "B4A7D6", "accent6": "F4B183",
                           "dk2": "1B2A41", "lt2": "E7E6E6"})
    for i, t in enumerate(["Night mode", "Highlights"]):
        s = blank(prs)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = rgb("1B2A41")
        textbox(s, 0.6, 0.4, 12, 1.0, [t], size=36, color="FFFFFF")
        textbox(s, 0.6, 1.6, 7.0, 3.0, ["White text on a dark background.", "Second paragraph."],
                size=20, color="E7E6E6")
        for k, col in enumerate(["00B0F0", "FFC000", "92D050"]):
            autoshape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 8.2 + (k % 2) * 2.3, 1.6 + k * 1.5, 2.0, 1.2,
                      fill=col, text=f"{(i + 1) * (k + 3)}%", text_color="1B2A41", size=20)


@deck("d26_kitchen_sink", ("title", "bullets", "table", "chart", "shapes", "images",
                           "theme_colors"))
def d26(prs: Any, rng: random.Random, a: AssetSet) -> None:
    set_theme_colors(prs, {"accent1": "2E75B6", "accent2": "ED7D31", "accent3": "A5A5A5",
                           "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47"})
    s = blank(prs)
    textbox(s, 0.5, 0.25, 12, 0.8, ["Dashboard"], size=28)
    textbox(s, 0.5, 1.1, 4.0, 2.6,
            [{"text": "Visitors up 12%", "bullet": "•"}, {"text": "Members up 8%", "bullet": "•"},
             {"text": "Costs flat", "bullet": "•"}], size=16)
    table(s, [["KPI", "Value"], ["Events", "48"], ["Schools", "31"]], 0.5, 4.0, 4.0, 1.8, size=12)
    add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, ["Q1", "Q2", "Q3", "Q4"],
              [("2026", [12.0, 18.0, 15.0, 22.0])], 4.8, 1.1, 4.6, 3.2, legend=False)
    s.shapes.add_picture(str(a.photo2), Inches(9.8), Inches(1.1), Inches(3.0), Inches(3.0))
    autoshape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 4.8, 4.7, 3.6, 1.2, fill="2E75B6", text="On track",
              size=18)
    autoshape(s, MSO_SHAPE.OVAL, 9.0, 4.6, 1.5, 1.5, fill="ED7D31", text="!", size=28)
    s.shapes.add_picture(str(a.logo_red), Inches(12.2), Inches(6.6), Inches(0.7))


@deck("d27_italian_report_4x3", ("italian", "table", "chart", "aspect_4_3"), aspect="4:3")
def d27(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Visitatori per mese")
    table(s, [["Mese", "Ingressi", "Variazione"], ["Gennaio", "1.240", "+4%"],
              ["Febbraio", "1.385", "+12%"], ["Marzo", "1.910", "+38%"]], 0.7, 1.8, 8.6, 2.4,
          size=16)
    textbox(s, 0.7, 4.6, 8.6, 1.5, ["Nota: i dati di marzo includono la Settimana dei Musei e "
                                     "l'apertura straordinaria di venerdì sera."], size=14,
            color="595959")
    s = prs.slides.add_slide(prs.slide_layouts[5])
    set_title(s, "Andamento delle attività")
    add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, ["Visite", "Laboratori", "Eventi"],
              [("2025", [320.0, 140.0, 45.0]), ("2026", [410.0, 190.0, 52.0])], 0.7, 1.6, 8.6,
              5.2)


@deck("d28_rotated_text", ("rotation", "rotated_text", "vertical_text", "pdf_outlined_text"))
def d28(prs: Any, rng: random.Random, a: AssetSet) -> None:
    s = blank(prs)
    textbox(s, 0.6, 0.3, 12, 0.9, ["Rotated text"], size=30)
    tb = textbox(s, 0.2, 3.0, 3.5, 0.6, ["Rotated 90 degrees"], size=18)
    tb.rotation = 270
    tb = textbox(s, 3.5, 2.5, 4.5, 0.6, ["Tilted by minus 30 degrees"], size=18, color="C0504D")
    tb.rotation = -30
    tb = textbox(s, 8.0, 4.5, 4.5, 0.6, ["Slightly tilted label"], size=18, color="1F497D")
    tb.rotation = 8
    textbox(s, 3.0, 6.0, 7.0, 0.6, ["Normal text for comparison"], size=18,
            anchor=MSO_ANCHOR.MIDDLE)


def build_deck(spec: DeckSpec, out_pptx: Path, asset_set: AssetSet) -> Path:
    prs = new_presentation(spec.aspect)
    spec.builder(prs, random.Random(spec.seed), asset_set)
    out_pptx.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_pptx))
    return out_pptx
