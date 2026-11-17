"""Expand build animations into one slide per click state, for scoring (M7).

A collapsed build-up (``analyze/builds.py``) is one slide whose entrance animations
reproduce the PDF's pages click by click. PowerPoint renders and exports a slide in
its final state, so page-by-page scoring needs the intermediate states: for a slide
with K clicks, K copies are inserted before it, copy ``s`` without the shapes and
paragraphs that appear after click ``s``. The original slide stays as the final
state. This is done on the zip parts (no PowerPoint): slide XML copied, the targets of
later clicks removed (whole shapes by id; paragraphs ``pRg st..end`` from their text
body), ``p:timing`` dropped, the slide's relationships copied without its notes slide,
and the new parts registered in ``[Content_Types].xml``, ``presentation.xml`` and its
relationships. Sections are dropped from the copy (scoring does not use them).
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
NOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


@dataclass
class Target:
    spid: int
    paras: tuple[int, int] | None = None


@dataclass
class Expansion:
    path: Path
    slides_in: int
    slides_out: int
    # expanded slide index -> (original slide index, click state; -1 = final / unchanged)
    states: list[tuple[int, int]] = field(default_factory=list)


def click_targets(sld: etree._Element) -> list[list[Target]]:
    """Entrance-effect targets per click of a slide's main sequence."""
    seq = sld.find(".//p:cTn[@nodeType='mainSeq']", NS)
    if seq is None:
        return []
    lst = seq.find("p:childTnLst", NS)
    if lst is None:
        return []
    clicks: list[list[Target]] = []
    for par in lst.findall("p:par", NS):
        targets = []
        for eff in par.iter(f"{{{NS['p']}}}cTn"):
            if eff.get("presetClass") != "entr":
                continue
            for tgt in eff.iter(f"{{{NS['p']}}}spTgt"):
                rng = tgt.find("p:txEl/p:pRg", NS)
                paras = (int(rng.get("st", "0")), int(rng.get("end", "0"))) if rng is not None \
                    else None
                targets.append(Target(int(tgt.get("spid", "0")), paras))
        clicks.append(targets)
    return clicks


def _state(sld_xml: bytes, hidden: list[Target]) -> bytes:
    root = etree.fromstring(sld_xml)
    timing = root.find("p:timing", NS)
    if timing is not None:
        root.remove(timing)
    tree = root.find("p:cSld/p:spTree", NS)
    whole = {t.spid for t in hidden if t.paras is None}
    paras: dict[int, set[int]] = {}
    for t in hidden:
        if t.paras is not None:
            paras.setdefault(t.spid, set()).update(range(t.paras[0], t.paras[1] + 1))
    for c_nv in list(tree.iter(f"{{{NS['p']}}}cNvPr")):
        sid = int(c_nv.get("id", "0"))
        shape = c_nv.getparent().getparent()  # nvXxPr -> shape element
        if sid in whole and shape.getparent() is not None:
            shape.getparent().remove(shape)
        elif sid in paras:
            body = shape.find("p:txBody", NS)
            if body is None:
                continue
            ps = body.findall("a:p", NS)
            for k in sorted(paras[sid], reverse=True):
                if 0 <= k < len(ps):
                    body.remove(ps[k])
            if body.find("a:p", NS) is None:
                etree.SubElement(body, f"{{{NS['a']}}}p")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def expand_builds(pptx: Path, out: Path) -> Expansion:
    """Write the expanded deck to ``out`` (or return ``pptx`` itself when no slide has
    entrance animations)."""
    src = zipfile.ZipFile(pptx)
    names = src.namelist()
    pres = etree.fromstring(src.read("ppt/presentation.xml"))
    rels = etree.fromstring(src.read("ppt/_rels/presentation.xml.rels"))
    rel_target = {r.get("Id"): r.get("Target") for r in rels}
    sld_list = pres.find("p:sldIdLst", NS)
    entries = list(sld_list) if sld_list is not None else []
    plan: list[tuple[etree._Element, str, list[list[Target]]]] = []
    for el in entries:
        target = rel_target[el.get(f"{{{NS['r']}}}id")]
        part = "ppt/" + target.lstrip("/").removeprefix("ppt/")
        plan.append((el, part, click_targets(etree.fromstring(src.read(part)))))
    n_in = len(entries)
    if not any(clicks for _, _, clicks in plan):
        return Expansion(pptx, n_in, n_in, [(i, -1) for i in range(n_in)])
    numbers = [int(m.group(1)) for n in names if (m := re.match(r"ppt/slides/slide(\d+)\.xml$",
                                                               n))]
    next_no = max(numbers, default=0) + 1
    next_id = max(int(e.get("id")) for e in entries) + 1
    rel_ids = {r.get("Id") for r in rels}
    ct = etree.fromstring(src.read("[Content_Types].xml"))
    new_parts: dict[str, bytes] = {}
    states: list[tuple[int, int]] = []
    for i, (el, part, clicks) in enumerate(plan):
        xml = src.read(part)
        rels_name = part.replace("slides/", "slides/_rels/") + ".rels"
        for s in range(len(clicks)):
            hidden = [t for k in range(s, len(clicks)) for t in clicks[k]]
            name = f"ppt/slides/slide{next_no}.xml"
            new_parts[name] = _state(xml, hidden)
            if rels_name in names:
                srels = etree.fromstring(src.read(rels_name))
                for r in list(srels):
                    if r.get("Type") == NOTES_RT:
                        srels.remove(r)
                new_parts[f"ppt/slides/_rels/slide{next_no}.xml.rels"] = etree.tostring(
                    srels, xml_declaration=True, encoding="UTF-8", standalone=True)
            rid = f"rIdX{next_no}"
            while rid in rel_ids:
                rid += "x"
            rel_ids.add(rid)
            r = etree.SubElement(rels, f"{{{NS_REL}}}Relationship")
            r.set("Id", rid)
            r.set("Type", SLIDE_RT)
            r.set("Target", f"slides/slide{next_no}.xml")
            ov = etree.SubElement(ct, f"{{{NS_CT}}}Override")
            ov.set("PartName", f"/{name}")
            ov.set("ContentType", SLIDE_CT)
            sld_id = etree.Element(f"{{{NS['p']}}}sldId")
            sld_id.set("id", str(next_id))
            sld_id.set(f"{{{NS['r']}}}id", rid)
            el.addprevious(sld_id)
            states.append((i, s))
            next_no += 1
            next_id += 1
        states.append((i, -1))
    ext_lst = pres.find("p:extLst", NS)
    if ext_lst is not None:
        for ext in ext_lst.findall("p:ext", NS):
            if ext.get("uri") == SECTION_EXT_URI:
                ext_lst.remove(ext)
    changed = {
        "ppt/presentation.xml": etree.tostring(pres, xml_declaration=True, encoding="UTF-8",
                                               standalone=True),
        "ppt/_rels/presentation.xml.rels": etree.tostring(rels, xml_declaration=True,
                                                          encoding="UTF-8", standalone=True),
        "[Content_Types].xml": etree.tostring(ct, xml_declaration=True, encoding="UTF-8",
                                              standalone=True),
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for n in names:
            dst.writestr(n, changed.get(n) or src.read(n))
        for n in sorted(new_parts):
            dst.writestr(n, new_parts[n])
    src.close()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(buf.getvalue())
    return Expansion(out, n_in, len(states), states)
