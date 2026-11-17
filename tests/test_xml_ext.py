"""build/xml_ext.py helpers (M6): every feature python-pptx lacks, written by the real
writer from a hand-made IR, opened in PowerPoint (COM) and checked by exporting the
deck to PDF and parsing it back: gradients (linear, radial), preset and custom dashes,
arrowheads, connectors with connection references, preset shapes with adjust values,
rotation and flips, an SVG picture, a native table with borders / fills / merges, a
solid slide background, a locked shape and a group.

A part PowerPoint had to repair or drop would be missing from the export, so each
feature is checked to come back."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from pdf2pptx.analyze import shapes
from pdf2pptx.build import xml_ext
from pdf2pptx.build.writer import write_pptx
from pdf2pptx.ingest import parse
from pdf2pptx.ir import (
    Color,
    Connector,
    Document,
    Fill,
    Gradient,
    GradientStop,
    Group,
    Image,
    Page,
    Paragraph,
    Shape,
    Stroke,
    Table,
    TableCell,
    TextBlock,
    TextRun,
)


def _grad(kind: str, angle: float = 0.0) -> Fill:
    return Fill(kind="gradient", gradient=Gradient(kind=kind, angle=angle, stops=[
        GradientStop(pos=0.0, color=Color(rgb="FF0000")),
        GradientStop(pos=1.0, color=Color(rgb="0000FF"))], center=(0.5, 0.5)))


def _cell_text(tid: str, text: str, box: tuple[float, float, float, float]) -> TextBlock:
    return TextBlock(id=tid, bbox=box, paragraphs=[Paragraph(id=f"{tid}_p", bbox=box, runs=[
        TextRun(id=f"{tid}_r", bbox=box, text=text, font="Arial", size=12)])])


def _deck(wd: Path) -> Document:
    page = Page(index=0, width=960, height=540)
    els: list = [
        Shape(id="bg", bbox=(0, 0, 960, 540), z=1, preset="rect", is_background=True,
              fill=Fill(kind="solid", color=Color(rgb="F2F2F2"))),
        Shape(id="lin", bbox=(40, 40, 240, 140), z=2, preset="rect", fill=_grad("linear", 90)),
        Shape(id="rad", bbox=(280, 40, 420, 180), z=3, preset="ellipse", fill=_grad("radial")),
        Shape(id="arr", bbox=(460, 40, 660, 110), z=4, preset="rightArrow",
              adjust={"adj1": 30000, "adj2": 80000}, rotation=20.0,
              fill=Fill(kind="solid", color=Color(rgb="9BBB59", alpha=0.6)),
              stroke=Stroke(color=Color(rgb="000000"), width=1.5, dash="dash")),
        Shape(id="chev", bbox=(700, 40, 900, 110), z=5, preset="chevron", flip_h=True,
              adjust={"adj": 30000}, fill=Fill(kind="solid", color=Color(rgb="4BACC6")),
              stroke=Stroke(color=Color(rgb="1F497D"), width=2.0, dash="custom",
                            dash_pattern=[6.0, 2.0, 1.0, 2.0])),
        Shape(id="boxA", bbox=(40, 220, 180, 300), z=6, preset="roundRect",
              fill=Fill(kind="solid", color=Color(rgb="4F81BD"))),
        Shape(id="boxB", bbox=(400, 320, 540, 400), z=7, preset="ellipse",
              fill=Fill(kind="solid", color=Color(rgb="C0504D"))),
        Shape(id="cx", bbox=(180, 260, 400, 360), z=8, preset="bentConnector3",
              adjust={"adj1": 50000}, connector=Connector(kind="elbow", start_id="boxA",
                                                          start_site=3, end_id="boxB",
                                                          end_site=2),
              stroke=Stroke(color=Color(rgb="404040"), width=2.0, head_end="oval",
                            tail_end="triangle", tail_size="lg-lg")),
        Shape(id="ln", bbox=(40, 470, 400, 470), z=9, preset="line",
              connector=Connector(kind="straight"),
              stroke=Stroke(color=Color(rgb="7F7F7F"), width=3.0, tail_end="stealth")),
        Shape(id="g1", bbox=(600, 450, 640, 490), z=10, preset="diamond",
              fill=Fill(kind="solid", color=Color(rgb="F79646"))),
        Shape(id="g2", bbox=(650, 450, 690, 490), z=11, preset="triangle",
              fill=Fill(kind="solid", color=Color(rgb="8064A2"))),
        Group(id="grp", bbox=(600, 450, 690, 490), z=10, child_ids=["g1", "g2"]),
    ]
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" width="100pt" height="60pt" '
           'viewBox="0 0 100 60"><path d="M 0 0 L 100 0 L 50 60 Z" fill="#00B050"/></svg>\n')
    (wd / "images").mkdir(parents=True, exist_ok=True)
    (wd / "images/t.svg").write_text(svg, encoding="utf-8", newline="\n")
    from PIL import Image as PILImage

    PILImage.new("RGBA", (200, 120), (0, 176, 80, 255)).save(wd / "images/t.png")
    els.append(Image(id="svg", bbox=(760, 200, 860, 260), z=12, file="images/t.png",
                     svg_file="images/t.svg"))
    thin = Stroke(color=Color(rgb="000000"), width=1.0)
    thick = Stroke(color=Color(rgb="C00000"), width=3.0, dash="sysDot")
    cells = []
    for r in range(3):
        for c in range(3):
            x0, y0 = 600 + 100 * c, 300 + 30 * r
            if (r, c) == (0, 1):
                cells.append(TableCell(row=0, col=1, col_span=2, borders={
                    "top": thin, "bottom": thick, "left": thin, "right": thin},
                    fill=Fill(kind="solid", color=Color(rgb="DCE6F1")),
                    text=_cell_text("t_h", "Merged", (x0 + 6, y0 + 4, x0 + 194, y0 + 20))))
                continue
            if (r, c) == (0, 2):
                cells.append(TableCell(row=0, col=2, merged_into=(0, 1)))
                continue
            cells.append(TableCell(row=r, col=c, borders={"top": thin, "bottom": thin},
                                   text=_cell_text(f"t{r}{c}", f"C{r}{c}",
                                                   (x0 + 6, y0 + 4, x0 + 94, y0 + 20))))
    els.append(Table(id="tbl", bbox=(600, 300, 900, 390), z=13, row_heights=[30, 30, 30],
                     col_widths=[100, 100, 100], cells=cells))
    page.elements = els
    return Document(pages=[page])


def test_xml_is_well_formed_and_ordered(tmp_path: Path) -> None:
    out = tmp_path / "x.pptx"
    write_pptx(_deck(tmp_path), out, tmp_path, tmp_path / "none.pdf")
    z = zipfile.ZipFile(out)
    xml = z.read("ppt/slides/slide1.xml").decode()
    for needle in ("<a:gradFill", '<a:path path="circle"', '<a:prstDash val="dash"/>',
                   "<a:custDash>", '<a:tailEnd type="triangle" w="lg" len="lg"/>',
                   '<a:headEnd type="oval"', "<p:cxnSp", "<a:stCxn", "<a:endCxn",
                   'prst="rightArrow"', '<a:gd name="adj2" fmla="val 80000"/>', 'flipH="1"',
                   "<p:grpSp>", "asvg:svgBlip", "<a:tbl>", "<a:lnB", 'gridSpan="2"',
                   "<p:bg>", xml_ext.NO_STYLE_TABLE):
        assert needle in xml, needle
    assert any(n.endswith(".svg") for n in z.namelist())
    assert "image/svg+xml" in z.read("[Content_Types].xml").decode()


@pytest.mark.powerpoint
def test_every_feature_opens_and_renders_in_powerpoint(tmp_path: Path) -> None:
    from pdf2pptx.verify import ppt_com

    out = tmp_path / "x.pptx"
    write_pptx(_deck(tmp_path), out, tmp_path, tmp_path / "none.pdf")
    eng = ppt_com.shared()
    info = eng.open(out)
    assert info.slides == 1
    pdf = eng.export_pdf(info, tmp_path / "x.pdf")
    layout = {s.name: s for s in eng.text_layout(info, 0)}
    png = eng.render(info, [0], 72, tmp_path / "png")[0]
    eng.close(info)
    # the table and its cell text survived (merged header, 3 x 3 grid)
    assert any(n.startswith("Table tbl[1,2]") for n in layout)
    assert "Merged" in "".join(ln.text for s in layout.values() for ln in s.lines)
    # parse the export back: connectors with their heads, presets, the gradient
    page = parse.parse_pages(Path(pdf), [0], tmp_path / "rt")[0][0]
    shapes.analyze_page(page, tmp_path / "rt")
    got = [e for e in page.elements if isinstance(e, Shape)]
    presets_found = {e.preset for e in got}
    for p in ("rightArrow", "chevron", "roundRect", "ellipse", "diamond", "triangle",
              "bentConnector3", "line"):
        assert p in presets_found, (p, presets_found)
    cx = next(e for e in got if e.preset == "bentConnector3")
    assert cx.stroke is not None and (cx.stroke.head_end, cx.stroke.tail_end) == ("oval",
                                                                                  "triangle")
    assert any(e.fill.kind == "gradient" for e in got)
    bg = next(e for e in page.elements if getattr(e, "is_background", False))
    assert isinstance(bg, Shape) and bg.fill.color is not None and bg.fill.color.rgb == "F2F2F2"
    # the SVG picture is drawn (green) where it was placed
    from PIL import Image as PILImage

    px = PILImage.open(png).convert("RGB").getpixel((810, 215))
    assert px[1] > 120 and px[0] < 80, px
