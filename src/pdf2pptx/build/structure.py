"""Deck-level PPTX structure (M7): theme part, slide master and layouts, placeholders,
build animations, sections and speaker notes.

Placeholder text must render exactly like the text box it replaces, so the master's
text styles (``p:txStyles``) are made neutral: title, body and other styles all get
the presentation's default text style (the one text boxes inherit), and the master and
layout placeholders have zero insets, top anchoring and no autofit. Slide placeholders
keep their own frame, insets and paragraph / run properties (``writer._add_text``).

``tests/test_structure.py`` opens decks using every helper in PowerPoint and checks
the master shapes, slide-number fields, animation counts and sections through COM.
"""

from __future__ import annotations

import copy
import hashlib
from collections.abc import Sequence
from typing import Any
from xml.sax.saxutils import escape

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

from ..ir import Theme
from .xml_ext import NSDECL

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
THEME_SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4",
               "accent5", "accent6", "hlink", "folHlink")
# placeholder kind (IR) -> (ph type or None for content, idx or None, sz or None, name)
PH = {
    "title": ("title", None, None, "Title"),
    "subTitle": ("subTitle", 1, None, "Subtitle"),
    "body": (None, 1, None, "Content Placeholder"),
    "body2": (None, 2, "half", "Content Placeholder"),
    "sldNum": ("sldNum", 12, "quarter", "Slide Number Placeholder"),
}
PROMPTS = {"title": "Click to edit Master title style", "subTitle": "Click to edit subtitle",
           "body": "Click to edit Master text styles", "body2": "Click to edit Master text styles"}


def guid(*parts: object) -> str:
    g = hashlib.sha256("/".join(map(str, parts)).encode()).hexdigest().upper()
    return f"{{{g[:8]}-{g[8:12]}-{g[12:16]}-{g[16:20]}-{g[20:32]}}}"


# --------------------------------------------------------------------------- theme


def theme_part(prs: Any) -> Any:
    return prs.slide_master.part.part_related_by(RT.THEME)


def apply_theme(prs: Any, theme: Theme | None) -> None:
    """Theme colours (``srgbClr`` in every slot) and major / minor latin fonts."""
    if theme is None:
        return
    part = theme_part(prs)
    root = etree.fromstring(part.blob)
    scheme = root.find(f".//{{{NS_A}}}clrScheme")
    if scheme is not None:
        scheme.set("name", theme.name or "pdf2pptx")
        for slot in THEME_SLOTS:
            val = theme.colors.get(slot)
            el = scheme.find(f"{{{NS_A}}}{slot}")
            if not val or el is None:
                continue
            for child in list(el):
                el.remove(child)
            etree.SubElement(el, f"{{{NS_A}}}srgbClr").set("val", val.upper())
    fonts = root.find(f".//{{{NS_A}}}fontScheme")
    if fonts is not None:
        fonts.set("name", theme.name or "pdf2pptx")
        for tag, face in (("majorFont", theme.major_font), ("minorFont", theme.minor_font)):
            latin = fonts.find(f"{{{NS_A}}}{tag}/{{{NS_A}}}latin")
            if face and latin is not None:
                latin.set("typeface", face)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# --------------------------------------------------------------------------- master


def neutral_text_styles(prs: Any) -> None:
    """Master title / body / other styles = the presentation's default text style."""
    default = prs.part._element.find(qn("p:defaultTextStyle"))
    styles = prs.slide_master._element.find(qn("p:txStyles"))
    if default is None or styles is None:
        return
    levels = [el for el in default if el.tag.endswith("pPr") and "lvl" in el.tag]
    for tag in ("p:titleStyle", "p:bodyStyle", "p:otherStyle"):
        st = styles.find(qn(tag))
        if st is None:
            continue
        for child in list(st):
            st.remove(child)
        for lv in levels:
            st.append(copy.deepcopy(lv))


def clear_tree(container: Any) -> None:
    tree = container.shapes._spTree
    for el in list(tree):
        if el.tag in (qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"),
                      qn("p:cxnSp"), qn("p:contentPart")):
            tree.remove(el)


def renumber_ids(container: Any) -> None:
    """Renumber the shape ids of a master / layout from 1 in document order.

    python-pptx's ``_next_shape_id`` counts every ``@id`` of the part; on a slide master
    that includes the layout ids (2147483649 and up), and PowerPoint refuses shape ids
    that large. Connector references follow the renumbering."""
    tree = container.shapes._spTree
    mapping: dict[str, str] = {}
    for k, el in enumerate(tree.iter(qn("p:cNvPr")), start=1):
        old = el.get("id", "")
        mapping[old] = str(k)
        el.set("id", str(k))
        name = el.get("name", "")
        if old and name.endswith(f" {old}"):  # python-pptx names shapes "<Kind> <id>"
            el.set("name", f"{name[:-len(old)]}{k}")
    for tag in ("a:stCxn", "a:endCxn"):
        for el in tree.iter(qn(tag)):
            if el.get("id") in mapping:
                el.set("id", mapping[el.get("id")])


def next_id(container: Any) -> int:
    """The next free shape id of a master / layout (shape ids only)."""
    ids = [int(el.get("id", "0")) for el in container.shapes._spTree.iter(qn("p:cNvPr"))]
    return max([i for i in ids if i < 2**31 - 1], default=1) + 1


def placeholder_xml(sid: int, kind: str, box: tuple[int, int, int, int],
                    prompt: bool = True, master: bool = False) -> Any:
    """A master / layout placeholder ``p:sp`` (EMU box: x, y, cx, cy). On the master the
    body placeholder must be typed ``body`` (a content placeholder fails to open)."""
    ph_type, idx, sz, name = PH[kind]
    if master and kind == "body":
        ph_type = "body"
    attrs = (f' type="{ph_type}"' if ph_type else "") + (f' sz="{sz}"' if sz else "") + (
        f' idx="{idx}"' if idx is not None else "")
    x, y, cx, cy = box
    if kind == "sldNum":
        para = (f'<a:p><a:pPr algn="r"/><a:fld id="{guid("sldnum-ph", sid)}" type="slidenum">'
                '<a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>')
    else:
        text = escape(PROMPTS.get(kind, "")) if prompt else ""
        para = f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p>'
    return parse_xml(
        f"<p:sp {NSDECL}><p:nvSpPr><p:cNvPr id=\"{sid}\" name=\"{name} {sid}\"/>"
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f"<p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr>"
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{max(cx, 1)}" cy="{max(cy, 1)}"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" '
        'rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
        f"{para}</p:txBody></p:sp>")


def set_layout_props(layout: Any, name: str, kind: str | None, show_master: bool) -> None:
    el = layout._element
    el.find(qn("p:cSld")).set("name", name)
    if kind:
        el.set("type", kind)
    elif "type" in el.attrib:
        del el.attrib["type"]
    el.set("preserve", "1")
    if show_master:
        el.attrib.pop("showMasterSp", None)
    else:
        el.set("showMasterSp", "0")


def to_placeholder(sp: Any, kind: str) -> None:
    """Turn a text box ``p:sp`` written by the writer into a slide placeholder."""
    ph_type, idx, sz, _ = PH[kind]
    nv = sp.find(qn("p:nvSpPr"))
    c_nv = nv.find(qn("p:cNvSpPr"))
    c_nv.attrib.pop("txBox", None)
    if c_nv.find(qn("a:spLocks")) is None:
        etree.SubElement(c_nv, qn("a:spLocks")).set("noGrp", "1")
    nv_pr = nv.find(qn("p:nvPr"))
    ph = etree.SubElement(nv_pr, qn("p:ph"))
    nv_pr.remove(ph)
    nv_pr.insert(0, ph)
    if ph_type:
        ph.set("type", ph_type)
    if sz:
        ph.set("sz", sz)
    if idx is not None:
        ph.set("idx", str(idx))


# --------------------------------------------------------------------------- animation


def timing_xml(steps: Sequence[Sequence[tuple[int, tuple[int, int] | None]]],
               bld: Sequence[tuple[int, str]]) -> Any:
    """``p:timing`` with one click per step; each step makes its targets appear.

    ``steps[k]`` = [(shape id, (first, last) paragraph or None)]; ``bld`` = [(shape id,
    kind)] with kind ``p`` (paragraph build), ``sp`` (whole shape with text) or
    ``graphic`` (table)."""
    ids = iter(range(3, 100000))
    clicks = []
    for targets in steps:
        effects = []
        for j, (sid, rng) in enumerate(targets):
            tgt = f'<p:spTgt spid="{sid}"/>'
            if rng is not None:
                tgt = (f'<p:spTgt spid="{sid}"><p:txEl><p:pRg st="{rng[0]}" end="{rng[1]}"/>'
                       "</p:txEl></p:spTgt>")
            node = "clickEffect" if j == 0 else "withEffect"
            eff_id, set_id = next(ids), next(ids)
            effects.append(
                f'<p:par><p:cTn id="{eff_id}" presetID="1" presetClass="entr" presetSubtype="0" '
                f'fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/>'
                f'</p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{set_id}" dur="1" '
                'fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f"<p:tgtEl>{tgt}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility"
                '</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to>'
                "</p:set></p:childTnLst></p:cTn></p:par>")
        outer, inner = next(ids), next(ids)
        clicks.append(
            f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'</p:stCondLst><p:childTnLst><p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst>'
            '<p:cond delay="0"/></p:stCondLst><p:childTnLst>' + "".join(effects)
            + "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>")
    blds = []
    for sid, kind in bld:
        if kind == "p":
            blds.append(f'<p:bldP spid="{sid}" grpId="0" build="p"/>')
        elif kind == "sp":
            blds.append(f'<p:bldP spid="{sid}" grpId="0" animBg="1"/>')
        elif kind == "graphic":
            blds.append(f'<p:bldGraphic spid="{sid}" grpId="0"><p:bldAsOne/></p:bldGraphic>')
    bld_lst = f"<p:bldLst>{''.join(blds)}</p:bldLst>" if blds else ""
    return parse_xml(
        f"<p:timing {NSDECL}><p:tnLst><p:par>"
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f"<p:childTnLst>{''.join(clicks)}</p:childTnLst></p:cTn>"
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
        '</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
        "</p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>"
        f"{bld_lst}</p:timing>")


def set_timing(slide: Any, timing: Any) -> None:
    sld = slide._element
    old = sld.find(qn("p:timing"))
    if old is not None:
        sld.remove(old)
    # schema order: cSld, clrMapOvr, transition, timing, extLst
    after = None
    for tag in ("p:transition", "p:clrMapOvr", "p:cSld"):
        after = sld.find(qn(tag))
        if after is not None:
            break
    if after is None:
        sld.append(timing)
    else:
        after.addnext(timing)


# --------------------------------------------------------------------------- sections


def set_sections(prs: Any, sections: Sequence[tuple[str, Sequence[int]]]) -> None:
    """``p14:sectionLst`` (``sections`` = [(name, [slide indices])])."""
    pres = prs.part._element
    sld_ids = [s.get("id") for s in pres.find(qn("p:sldIdLst"))]
    ext_lst = pres.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(pres, qn("p:extLst"))
    for ext in ext_lst.findall(qn("p:ext")):
        if ext.get("uri") == SECTION_EXT_URI:
            ext_lst.remove(ext)
    ext = etree.Element(qn("p:ext"))
    ext.set("uri", SECTION_EXT_URI)
    ext_lst.insert(0, ext)
    lst = etree.SubElement(ext, f"{{{NS_P14}}}sectionLst", nsmap={"p14": NS_P14})
    for k, (name, idxs) in enumerate(sections):
        sec = etree.SubElement(lst, f"{{{NS_P14}}}section")
        sec.set("name", name)
        sec.set("id", guid("section", k, name))
        ids = etree.SubElement(sec, f"{{{NS_P14}}}sldIdLst")
        for i in idxs:
            if 0 <= i < len(sld_ids):
                etree.SubElement(ids, f"{{{NS_P14}}}sldId").set("id", sld_ids[i])


def read_sections(prs: Any) -> list[dict[str, Any]]:
    pres = prs.part._element
    sld_ids = [s.get("id") for s in pres.find(qn("p:sldIdLst"))]
    out = []
    for sec in pres.iter(f"{{{NS_P14}}}section"):
        ids = [el.get("id") for el in sec.iter(f"{{{NS_P14}}}sldId")]
        out.append({"name": sec.get("name"),
                    "slides": [sld_ids.index(i) for i in ids if i in sld_ids]})
    return out
