"""Deck-level checks (M7): the structured deck against the flat one, and against truth.

* ``render_check``: PowerPoint renders the structured and the flat deck; every slide
  pair must have SSIM >= ``RENDER_SSIM`` (fidelity mode keeps masters / theme only
  then). Without PowerPoint the check fails (fidelity keeps the flat deck).
* ``check_deck`` compares a converted deck with the synthetic truth: slide count after
  build collapsing, animation steps per slide (COM ``TimeLine.MainSequence.Count``),
  repeated elements in the master and not on the slides, slide-number fields, theme
  accents, sections, speaker notes, hyperlinks and document properties.
* ``pdf2pptx deck`` runs it over the corpus (``work/scores/deck.json`` + ``.md``).
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from skimage.metrics import structural_similarity

from .. import log
from ..build.structure import read_sections
from ..config import Config
from ..corpus.truth import theme_info

LOG = log.get(__name__)
RENDER_SSIM = 0.995
RENDER_DPI = 72


def _gray(path: Path) -> np.ndarray:
    return np.asarray(Image.open(path).convert("L"), dtype=np.float64)


def render_check(structured: Path, flat: Path, cfg: Config) -> dict[str, Any]:
    from . import ppt_com

    try:
        eng = ppt_com.shared(call_timeout=cfg.com_timeout_s)
    except Exception as exc:
        return {"ok": False, "reason": f"PowerPoint unavailable: {exc}"}
    with tempfile.TemporaryDirectory() as tmp:
        try:
            with eng.opened(structured) as a, eng.opened(flat) as b:
                if a.slides != b.slides:
                    return {"ok": True, "reason": "slide counts differ (builds collapsed)"}
                pa = eng.render(a, None, RENDER_DPI, Path(tmp) / "s")
                pb = eng.render(b, None, RENDER_DPI, Path(tmp) / "f")
        except Exception as exc:
            return {"ok": False, "reason": f"{type(exc).__name__}: {exc}"}
        scores = [float(structural_similarity(_gray(Path(x)), _gray(Path(y)), data_range=255))
                  for x, y in zip(pa, pb, strict=True)]
    worst = min(scores, default=1.0)
    return {"ok": worst >= RENDER_SSIM, "min_ssim": round(worst, 5),
            "slides": [round(s, 5) for s in scores], "threshold": RENDER_SSIM}


# --------------------------------------------------------------------------- facts


def _texts(shapes: Any, out: list[str]) -> None:
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            _texts(s.shapes, out)
        elif getattr(s, "has_text_frame", False) and s.has_text_frame and s.text_frame.text:
            out.append(s.text_frame.text.strip())


def _pictures(shapes: Any) -> int:
    n = 0
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            n += _pictures(s.shapes)
        elif s.shape_type == MSO_SHAPE_TYPE.PICTURE:
            n += 1
    return n


def _links(shapes: Any, slide_index: dict[str, int]) -> list[str]:
    out: list[str] = []
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            out.extend(_links(s.shapes, slide_index))
            continue
        try:
            ca = s.click_action
            if ca.target_slide is not None:
                out.append(f"#slide{slide_index.get(ca.target_slide.part.partname, -1) + 1}")
            elif ca.hyperlink.address:
                out.append(ca.hyperlink.address)
        except Exception:
            pass
        if getattr(s, "has_text_frame", False) and s.has_text_frame:
            for p in s.text_frame.paragraphs:
                for r in p.runs:
                    if r.hyperlink.address:
                        out.append(r.hyperlink.address)
    return out


def _has_slidenum(shapes: Any) -> bool:
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP and _has_slidenum(s.shapes):
            return True
        if getattr(s, "has_text_frame", False) and s.has_text_frame:
            for f in s._element.iter(qn("a:fld")):
                if f.get("type") == "slidenum":
                    return True
    return False


def deck_facts(pptx: Path, animations: list[int] | None = None) -> dict[str, Any]:
    prs = Presentation(str(pptx))
    index = {sl.part.partname: i for i, sl in enumerate(prs.slides)}
    master = prs.slide_master
    m_texts: list[str] = []
    _texts([s for s in master.shapes if not s.is_placeholder], m_texts)
    layout_texts: list[str] = []
    for lo in prs.slide_layouts:
        _texts([s for s in lo.shapes if not s.is_placeholder], layout_texts)
    slides = []
    for i, sl in enumerate(prs.slides):
        texts: list[str] = []
        _texts(sl.shapes, texts)
        ph = sorted(str(s.placeholder_format.type).split(" ")[0] for s in sl.shapes
                    if s.is_placeholder)
        slides.append({
            "layout": sl.slide_layout.name,
            "texts": texts,
            "pictures": _pictures(sl.shapes),
            "placeholders": ph,
            "slide_number_field": _has_slidenum(sl.shapes),
            "slide_number_placeholder": any(
                s.is_placeholder and s.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER
                for s in sl.shapes),
            "notes": sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else "",
            "links": _links(sl.shapes, index),
            "animations": animations[i] if animations and i < len(animations) else None,
        })
    return {
        "slides": slides,
        "master_pictures": _pictures([s for s in master.shapes if not s.is_placeholder]),
        "master_texts": m_texts,
        "layout_texts": layout_texts,
        "layouts": [lo.name for lo in prs.slide_layouts],
        "theme": theme_info(prs),
        "sections": read_sections(prs),
        "title": prs.core_properties.title or "",
        "author": prs.core_properties.author or "",
    }


# --------------------------------------------------------------------------- truth


def _norm(s: str) -> str:
    return " ".join(s.split()).casefold()


def _truth_link(link: str) -> str:
    if link.startswith("slide") and link.endswith(".xml"):
        return f"#slide{link[5:-4]}"
    return link


def check_deck(truth: dict[str, Any], facts: dict[str, Any],
               pdf_links: set[str] | None = None,
               pdf_colors: set[str] | None = None) -> dict[str, Any]:
    """Every check that the truth file can answer for this deck. ``pdf_links`` /
    ``pdf_colors``: what the PDF itself contains (links its export kept, colours it
    draws), so that what is recoverable is reported apart from what is not."""
    ts = truth["slides"]
    out: dict[str, Any] = {}
    builds = truth.get("builds") or []
    absorbed = sum(len(b["slides"]) - 1 for b in builds)
    out["slides"] = {"truth": len(ts), "expected": len(ts) - absorbed,
                     "got": len(facts["slides"]),
                     "ok": len(facts["slides"]) == len(ts) - absorbed}
    if builds:
        anim = [s["animations"] for s in facts["slides"]]
        want = [b["steps"] for b in builds]
        got = sorted(a for a in anim if a)
        out["build_steps"] = {"expected": want, "got": anim, "ok": sorted(want) == got}
    # slide index of each truth slide after collapsing
    slide_of = []
    k = 0
    gone = {i for b in builds for i in b["slides"][:-1]}
    for i in range(len(ts)):
        slide_of.append(k)
        if i not in gone:
            k += 1
    if "master_candidate" in truth.get("features", []):
        repeated = _repeated_texts(ts)
        slides_with = [sum(1 for t in s["texts"] if _norm(t) in repeated)
                       for s in facts["slides"]]
        out["master"] = {
            "master_pictures": facts["master_pictures"],
            "master_texts": facts["master_texts"],
            "slide_pictures": [s["pictures"] for s in facts["slides"]],
            "repeated_text_on_slides": slides_with,
            "ok": facts["master_pictures"] >= 1 and bool(facts["master_texts"])
            and not any(s["pictures"] for s in facts["slides"]) and not any(slides_with),
        }
    want_num = [i for i, s in enumerate(ts) if s.get("has_slide_number_field")]
    if want_num:
        got = [i for i in want_num if facts["slides"][slide_of[i]]["slide_number_field"]]
        out["slide_number_fields"] = {"expected": len(want_num), "got": len(got),
                                      "placeholders": sum(s["slide_number_placeholder"]
                                                          for s in facts["slides"]),
                                      "ok": len(got) == len(want_num)}
    tt, ft = truth.get("theme", {}).get("colors", {}), facts["theme"]["colors"]
    accents = [f"accent{i}" for i in range(1, 7)]
    out["theme"] = {
        "accents_same_slot": sum(1 for a in accents if tt.get(a) and tt.get(a) == ft.get(a)),
        "accents_in_theme": sum(1 for a in accents if tt.get(a) in {ft.get(b) for b in accents}),
        "accents_visible": None if pdf_colors is None else sum(
            1 for a in accents if tt.get(a) in pdf_colors),
        "fonts_ok": (truth.get("theme", {}).get("major_font") == facts["theme"].get("major_font")
                     and truth.get("theme", {}).get("minor_font")
                     == facts["theme"].get("minor_font")),
    }
    if truth.get("sections"):
        want = [{"name": s["name"], "slides": sorted({slide_of[i] for i in s["slides"]})}
                for s in truth["sections"]]
        out["sections"] = {"expected": want, "got": facts["sections"],
                           "ok": want == facts["sections"]}
    notes = [(i, s["notes"]) for i, s in enumerate(ts) if s.get("notes")]
    if notes:
        got = [i for i, n in notes if _norm(facts["slides"][slide_of[i]]["notes"]) == _norm(n)]
        out["notes"] = {"expected": len(notes), "got": len(got), "ok": len(got) == len(notes)}
    links = [(i, _truth_link(h)) for i, s in enumerate(ts) for h in s.get("hyperlinks", [])]
    if links:
        found = [(i, h) for i, h in links if h in facts["slides"][slide_of[i]]["links"]]
        missing = [h for i, h in links if (i, h) not in found]
        lost = [h for h in missing if pdf_links is not None and h not in pdf_links]
        out["hyperlinks"] = {"expected": len(links), "got": len(found), "missing": missing,
                             "not_in_pdf": lost, "recoverable": len(links) - len(lost),
                             "ok": len(found) == len(links) - len(lost)}
    if truth.get("core"):
        c = truth["core"]
        out["core"] = {"expected": c, "got": {"title": facts["title"], "author": facts["author"]},
                       "ok": c.get("title", "") == facts["title"]
                       and (not c.get("author") or c["author"] == facts["author"])}
    return out


def _repeated_texts(ts: list[dict[str, Any]]) -> set[str]:
    """Text paragraphs that appear on at least 3 truth slides (digits ignored)."""
    from collections import Counter

    cnt: Counter[str] = Counter()
    for s in ts:
        cnt.update({_norm(p["text"]) for p in s.get("paragraphs", []) if p["role"] not in (
            "title", "body") and not p["text"].strip().isdigit()})
    return {t for t, n in cnt.items() if n >= 3 and t}


def pdf_facts(pdf: Path, cfg: Config) -> tuple[set[str], set[str]]:
    """(link targets, drawn colours) of a PDF, from its analyzed IR: URIs and
    ``#slide<n>`` for internal jumps; RGB hex of every fill, line and text colour."""
    from ..analyze.theme import collect
    from ..pipeline import analyzed

    doc, _, _ = analyzed(pdf, cfg, workers=1)
    links: set[str] = set()
    for page in doc.pages:
        for a in page.annotations:
            if a.subtype == "link" and a.uri:
                links.add(a.uri)
            elif a.subtype == "link" and a.dest_page is not None:
                links.add(f"#slide{a.dest_page + 1}")
    w, bg, _, _ = collect(doc)
    return links, set(w) | set(bg)


def deck_corpus(cfg: Config, only: list[str] | None = None,
                out_dir: Path | None = None) -> dict[str, Any]:
    from ..corpus.build import load_manifest
    from . import ppt_com

    out_dir = out_dir or cfg.out_dir / "corpus"
    try:
        eng: Any = ppt_com.shared(call_timeout=cfg.com_timeout_s)
    except Exception as exc:
        LOG.warning("PowerPoint unavailable (%s): animation counts not checked", exc)
        eng = None
    files = []
    for e in load_manifest(cfg.corpus_dir):
        if e["category"] != "synthetic_deck" or (only and e["name"] not in only):
            continue
        pptx = out_dir / f"{e['name']}.pptx"
        if not pptx.is_file():
            continue
        truth = json.loads((cfg.corpus_dir / e["truth"]).read_text(encoding="utf-8"))
        anim = None
        if eng is not None:
            with eng.opened(pptx) as info:
                anim = eng.animation_counts(info)
        facts = deck_facts(pptx, anim)
        links, colors = pdf_facts(cfg.corpus_dir / e["pdf"], cfg)
        files.append({"name": e["name"], "features": e.get("features", []),
                      "checks": check_deck(truth, facts, links, colors),
                      "layouts": facts["layouts"]})
    summary: dict[str, Any] = {}
    for f in files:
        for k, v in f["checks"].items():
            if isinstance(v, dict) and "ok" in v:
                s = summary.setdefault(k, {"files": 0, "ok": 0, "failing": []})
                s["files"] += 1
                s["ok"] += bool(v["ok"])
                if not v["ok"]:
                    s["failing"].append(f["name"])
    th = [f["checks"]["theme"] for f in files]
    vis = [t for t in th if (t["accents_visible"] or 0) >= 4]
    summary["theme_accents"] = {
        "decks": len(th),
        "decks_showing_4_accents": len(vis),
        "of_those_4_of_6_same_slot": sum(1 for t in vis if t["accents_same_slot"] >= 4),
        "of_those_4_of_6_in_theme": sum(1 for t in vis if t["accents_in_theme"] >= 4),
        "decks_4_of_6_same_slot": sum(1 for t in th if t["accents_same_slot"] >= 4),
        "decks_4_of_6_in_theme": sum(1 for t in th if t["accents_in_theme"] >= 4),
        "fonts_ok": sum(1 for t in th if t["fonts_ok"]),
    }
    res = {"summary": summary, "files": files}
    path = cfg.scores_dir / "deck.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(res, indent=1, ensure_ascii=False), encoding="utf-8",
                    newline="\n")
    return res


def summary_text(res: dict[str, Any]) -> str:
    lines = []
    for k, v in res["summary"].items():
        if k == "theme_accents":
            lines.append(f"theme accents: {v['decks_4_of_6_same_slot']}/{v['decks']} decks with "
                         f">= 4 of 6 in the same slot, {v['decks_4_of_6_in_theme']}/{v['decks']} "
                         f"anywhere in the theme; fonts {v['fonts_ok']}/{v['decks']}")
            lines.append(f"  decks whose PDF draws >= 4 truth accents: "
                         f"{v['decks_showing_4_accents']}; of those, >= 4 of 6 in the same slot "
                         f"{v['of_those_4_of_6_same_slot']}, in the theme "
                         f"{v['of_those_4_of_6_in_theme']}")
        else:
            lines.append(f"{k}: {v['ok']}/{v['files']}" + (
                f" (failing: {', '.join(v['failing'])})" if v["failing"] else ""))
    for f in res["files"]:
        t = f["checks"]["theme"]
        lines.append(f"  {f['name']}: accents same slot {t['accents_same_slot']}/6, "
                     f"in theme {t['accents_in_theme']}/6, drawn in the PDF "
                     f"{t['accents_visible']}/6")
        h = f["checks"].get("hyperlinks")
        if h and h["not_in_pdf"]:
            lines.append(f"    links not in the PDF (not recoverable): {h['not_in_pdf']}")
    return "\n".join(lines)
