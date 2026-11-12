"""Report-style PDFs generated with reportlab, with per-page ground truth.

The truth is recorded while the document is laid out (``afterFlowable``), so
each paragraph, table and figure is attributed to the page it lands on.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from . import assets

PARA_EN = [
    "The program started in a single neighbourhood and now runs in five. Each site has a "
    "coordinator, a group of volunteers and a small budget for materials.",
    "Families told us that the most useful part was the evening sessions, because parents "
    "could join after work. We moved two thirds of the sessions to the evening in spring.",
    "Attendance grew steadily through the year. The largest increase came after the local "
    "library agreed to host the sessions and to promote them in its newsletter.",
    "We tracked three outcomes: reading level, attendance and parent participation. All "
    "three improved, although the change in reading level was smaller than planned.",
]
PARA_IT = [
    "Il progetto è nato in un solo quartiere e oggi è attivo in cinque. Ogni sede ha un "
    "referente, un gruppo di volontari e un piccolo budget per i materiali.",
    "Le famiglie ci hanno detto che la parte più utile erano gli incontri serali, perché "
    "i genitori potevano partecipare dopo il lavoro.",
    "La partecipazione è cresciuta per tutto l'anno, soprattutto dopo che la biblioteca "
    "comunale ha accettato di ospitare gli incontri.",
]


class TruthDoc(BaseDocTemplate):
    def __init__(self, filename: str, **kw: Any) -> None:
        super().__init__(filename, **kw)
        self.truth_pages: dict[int, dict[str, list[Any]]] = {}

    def _page(self) -> dict[str, list[Any]]:
        idx = self.page - 1
        return self.truth_pages.setdefault(idx, {"paragraphs": [], "tables": [], "figures": []})

    def afterFlowable(self, flowable: Any) -> None:  # noqa: N802 - reportlab API
        role = getattr(flowable, "_truth_role", None)
        if isinstance(flowable, Paragraph):
            self._page()["paragraphs"].append(
                {"text": flowable.getPlainText(), "role": role or "body"}
            )
        elif isinstance(flowable, Table) and role is not None:
            self._page()["tables"].append(flowable._truth)  # type: ignore[attr-defined]
        elif isinstance(flowable, Image):
            self._page()["figures"].append({"role": "figure"})


def _p(text: str, style: ParagraphStyle, role: str = "body") -> Paragraph:
    p = Paragraph(text, style)
    p._truth_role = role  # type: ignore[attr-defined]
    return p


def _styles() -> dict[str, ParagraphStyle]:
    ss = getSampleStyleSheet()
    body = ParagraphStyle("body", parent=ss["BodyText"], fontName="Helvetica", fontSize=10,
                          leading=13.5, alignment=TA_JUSTIFY, spaceAfter=6)
    return {
        "body": body,
        "h1": ParagraphStyle("h1", parent=ss["Heading1"], fontName="Helvetica-Bold",
                             fontSize=18, leading=22, spaceAfter=10),
        "h2": ParagraphStyle("h2", parent=ss["Heading2"], fontName="Helvetica-Bold",
                             fontSize=13, leading=16, spaceBefore=6, spaceAfter=4),
        "caption": ParagraphStyle("caption", parent=body, fontName="Helvetica-Oblique",
                                  fontSize=9, leading=11),
    }


def _footer(canvas: Any, doc: Any) -> None:
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.drawString(2 * cm, 1.2 * cm, "Annual report 2026")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, str(doc.page))
    canvas.restoreState()


def _table(rows: list[list[str]], *, ruled: bool, header: bool = True) -> Table:
    t = Table(rows, hAlign="LEFT")
    cmds: list[Any] = [("FONT", (0, 0), (-1, -1), "Helvetica", 9),
                       ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    if header:
        cmds.append(("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9))
    if ruled:
        cmds += [("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                 ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE6F1"))]
    t.setStyle(TableStyle(cmds))
    t._truth_role = "table"  # type: ignore[attr-defined]
    t._truth = {"rows": len(rows), "cols": len(rows[0]), "ruled": ruled,  # type: ignore[attr-defined]
                "cells": rows}
    return t


def _doc(path: Path, columns: int) -> TruthDoc:
    doc = TruthDoc(str(path), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm,
                   topMargin=2 * cm, bottomMargin=2 * cm, title="", author="", creator="pdf2pptx",
                   invariant=1)
    gap = 0.8 * cm
    fw = (doc.width - gap * (columns - 1)) / columns
    frames = [Frame(doc.leftMargin + i * (fw + gap), doc.bottomMargin, fw, doc.height, id=f"c{i}")
              for i in range(columns)]
    doc.addPageTemplates([PageTemplate(id="main", frames=frames, onPage=_footer)])
    return doc


def report_two_column(path: Path) -> dict[str, Any]:
    st = _styles()
    doc = _doc(path, columns=2)
    story: list[Any] = [_p("Community Reading Program", st["h1"], "title")]
    for k in range(3):
        story.append(_p(f"{k + 1}. Section {k + 1}", st["h2"], "heading"))
        for para in PARA_EN:
            story.append(_p(para, st["body"]))
    doc.build(story)
    return {"columns": 2, "pages": _pages(doc)}


def report_tables(path: Path) -> dict[str, Any]:
    st = _styles()
    doc = _doc(path, columns=1)
    story: list[Any] = [_p("Relazione annuale: i numeri", st["h1"], "title"), _p(PARA_IT[0], st["body"])]
    story.append(_p("Tabella 1. Partecipanti per sede", st["caption"], "caption"))
    story.append(_table([["Sede", "2024", "2025", "2026"], ["Centro", "120", "145", "171"],
                         ["Nord", "80", "96", "110"], ["Sud", "64", "70", "88"],
                         ["Totale", "264", "311", "369"]], ruled=True))
    story.append(Spacer(1, 12))
    story.append(_p(PARA_IT[1], st["body"]))
    story.append(_p("Tabella 2. Costi (migliaia di euro), senza linee", st["caption"], "caption"))
    story.append(_table([["Voce", "Previsto", "Effettivo"], ["Personale", "88", "91"],
                         ["Materiali", "12", "10"], ["Affitti", "24", "24"],
                         ["Comunicazione", "6", "7"]], ruled=False))
    story.append(Spacer(1, 12))
    story.append(_p(PARA_IT[2], st["body"]))
    doc.build(story)
    return {"columns": 1, "pages": _pages(doc)}


def report_figure(path: Path, asset_dir: Path) -> dict[str, Any]:
    st = _styles()
    doc = _doc(path, columns=1)
    img_path = assets.diagram(asset_dir / "diagram.png")
    story: list[Any] = [_p("Results at a glance", st["h1"], "title"), _p(PARA_EN[0], st["body"])]
    img = Image(str(img_path), width=12 * cm, height=7.2 * cm)
    story += [img, _p("Figure 1. Sessions held per month, January to June.", st["caption"],
                      "caption")]
    for para in PARA_EN[1:]:
        story.append(_p(para, st["body"]))
    doc.build(story)
    return {"columns": 1, "pages": _pages(doc)}


def _pages(doc: TruthDoc) -> list[dict[str, Any]]:
    out = []
    for i in sorted(doc.truth_pages):
        page = doc.truth_pages[i]
        footer = [{"text": "Annual report 2026", "role": "footer"},
                  {"text": str(i + 1), "role": "page_number"}]
        out.append({"index": i, **page, "paragraphs": page["paragraphs"] + footer})
    return out


REPORTS = {
    "r01_report_two_column": ("multi_column", "headings", "footer", "page_number"),
    "r02_report_tables_it": ("table", "whitespace_table", "italian", "caption"),
    "r03_report_figure": ("figure", "caption", "image"),
}


def build_report(name: str, path: Path, asset_dir: Path) -> dict[str, Any]:
    path.parent.mkdir(parents=True, exist_ok=True)
    if name == "r01_report_two_column":
        return report_two_column(path)
    if name == "r02_report_tables_it":
        return report_tables(path)
    if name == "r03_report_figure":
        return report_figure(path, asset_dir)
    raise KeyError(name)
