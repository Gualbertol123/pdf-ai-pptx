"""Text structure vs ground truth (synthetic decks): paragraphs, levels, bullets, fonts.

For every truth slide the analyzed IR's paragraphs are compared with the answer key's
paragraphs (``truth.json``):

* paragraphs inside a truth table or chart frame are left out on both sides (tables
  and charts are M6 / M9; chart tick labels exist only in the PDF);
* each truth paragraph is paired with the IR paragraph of equal normalised text
  (in order, one to one);
* a slide **matches** when both sides have the same number of paragraphs, every truth
  paragraph is paired, and every pair has the same level;
* bullets are compared as present / absent (an ``inherited`` truth bullet is not
  checked: its visibility depends on the layout);
* **fonts**: every IR run of a paired paragraph counts as correct when its family is
  one of the truth paragraph's fonts.

Graphics (M6), for every truth auto-shape and connector: the IR shape whose frame
overlaps it most (IoU >= 0.5, frames grown by 1 pt so lines have an area) must have
the same preset. Tables: every truth table must be an IR ``Table`` (IoU >= 0.5) with
the same rows x cols and the same merges.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from ..config import Config
from ..ir import Document, Paragraph, Shape, Table, TextBlock
from .metrics import normalize_text

BBox = tuple[float, float, float, float]


def _norm(s: str) -> str:
    return normalize_text(s).casefold().replace(" ", "")


def _inside(bbox: BBox, frames: list[BBox]) -> bool:
    cx, cy = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
    return any(x0 - 1 <= cx <= x1 + 1 and y0 - 1 <= cy <= y1 + 1 for x0, y0, x1, y1 in frames)


def _ir_paragraphs(doc: Document, index: int, frames: list[BBox],
                   scale: tuple[float, float, float]) -> list[Paragraph]:
    s, dx, dy = scale
    page = next((p for p in doc.pages if p.index == index), None)
    if page is None:
        return []
    blocks = sorted((e for e in page.elements if isinstance(e, TextBlock)),
                    key=lambda b: (b.reading_order if b.reading_order is not None else 0, b.id))
    out = []
    for b in blocks:
        for p in b.paragraphs:
            x0, y0, x1, y1 = p.bbox
            slide_box = (x0 * s + dx, y0 * s + dy, x1 * s + dx, y1 * s + dy)
            if not _inside(slide_box, frames) and "".join(r.text for r in p.runs).strip():
                out.append(p)
    return out


def _has_bullet(truth_bullet: str) -> bool | None:
    if truth_bullet == "inherited":
        return None
    return truth_bullet != "none"


def compare_slide(doc: Document, slide: dict[str, Any], scale: tuple[float, float, float]
                  ) -> dict[str, Any]:
    frames = [tuple(e["bbox_pt"]) for e in slide.get("shape_list", [])
              if e.get("kind") in ("table", "chart")]
    # truth.py never takes paragraphs from tables or charts
    truth = [p for p in slide.get("paragraphs", []) if p["text"].strip()]
    ir = _ir_paragraphs(doc, slide["index"], frames, scale)  # type: ignore[arg-type]
    used: set[int] = set()
    pairs: list[tuple[dict[str, Any], Paragraph]] = []
    for tp in truth:
        key = _norm(tp["text"])
        j = next((k for k, p in enumerate(ir) if k not in used
                  and _norm("".join(r.text for r in p.runs)) == key), None)
        if j is not None:
            used.add(j)
            pairs.append((tp, ir[j]))
    levels_ok = all(tp["level"] == p.level for tp, p in pairs)
    bullets = [(b, p.bullet.kind != "none") for tp, p in pairs
               if (b := _has_bullet(tp["bullet"])) is not None]
    # a paragraph without runs in the truth (only a slide-number field) has no fonts
    runs = [(r.font, tp["fonts"]) for tp, p in pairs if tp["fonts"] for r in p.runs
            if r.text.strip()]
    font_ok = sum(1 for f, fonts in runs if f in fonts)
    match = len(truth) == len(ir) == len(pairs) and levels_ok
    return {
        "index": slide["index"],
        "truth_paragraphs": len(truth),
        "ir_paragraphs": len(ir),
        "paired": len(pairs),
        "levels_ok": levels_ok,
        "match": match,
        "bullets_checked": len(bullets),
        "bullets_ok": sum(1 for a, b in bullets if a == b),
        "runs": len(runs),
        "runs_font_ok": font_ok,
        "font_misses": sorted({f"{f} not in {fonts}" for f, fonts in runs if f not in fonts}),
        "unpaired_truth": [tp["text"] for tp in truth if all(tp is not q for q, _ in pairs)],
        "unpaired_ir": ["".join(r.text for r in p.runs) for k, p in enumerate(ir)
                        if k not in used],
    }


def _scale(doc: Document, truth: dict[str, Any], index: int) -> tuple[float, float, float]:
    """Page pt -> truth slide pt (the deck was exported at slide size, so usually 1:1)."""
    page = next((p for p in doc.pages if p.index == index), None)
    sw, sh = truth["slide_width_pt"], truth["slide_height_pt"]
    if page is None or page.width <= 0:
        return 1.0, 0.0, 0.0
    s = min(sw / page.width, sh / page.height)
    return s, (sw - page.width * s) / 2, (sh - page.height * s) / 2


def _iou(a: BBox, b: BBox) -> float:
    a = (a[0] - 1, a[1] - 1, a[2] + 1, a[3] + 1)
    b = (b[0] - 1, b[1] - 1, b[2] + 1, b[3] + 1)
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    union = (a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - inter
    return inter / union if union > 0 else 0.0


def compare_graphics(doc: Document, slide: dict[str, Any], scale: tuple[float, float, float]
                     ) -> dict[str, Any]:
    s, dx, dy = scale
    page = next((p for p in doc.pages if p.index == slide["index"]), None)
    ir_shapes = ([e for e in page.elements if isinstance(e, Shape) and not e.is_background]
                 if page else [])
    ir_tables = [e for e in page.elements if isinstance(e, Table)] if page else []

    def sbox(b: BBox) -> BBox:
        return (b[0] * s + dx, b[1] * s + dy, b[2] * s + dx, b[3] * s + dy)

    shapes_out = []
    used: set[str] = set()
    for t in slide.get("shape_list", []):
        if t["kind"] not in ("auto_shape", "connector") or not t.get("preset"):
            continue
        tb = tuple(t["bbox_pt"])
        best = max(((_iou(tb, sbox(e.bbox)), e) for e in ir_shapes if e.id not in used),
                   key=lambda x: x[0], default=(0.0, None))
        got = best[1].preset if best[1] is not None and best[0] >= 0.5 else None
        if best[1] is not None and best[0] >= 0.5:
            used.add(best[1].id)
        shapes_out.append({"truth": t["preset"], "ir": got or "missing",
                           "ok": got == t["preset"], "bbox": list(tb)})
    frames = [tuple(e["bbox_pt"]) for e in slide.get("shape_list", []) if e["kind"] == "table"]
    tables_out = []
    for tt, fr in zip(slide.get("tables", []), frames, strict=False):
        best = max(((_iou(fr, sbox(t.bbox)), t) for t in ir_tables), key=lambda x: x[0],
                   default=(0.0, None))
        t = best[1] if best[0] >= 0.5 else None
        merges = (sorted([c.row, c.col, c.row_span, c.col_span] for c in t.cells
                         if c.merged_into is None and (c.row_span > 1 or c.col_span > 1))
                  if t is not None else None)
        dims = (len(t.row_heights), len(t.col_widths)) if t is not None else None
        tables_out.append({"truth": f"{tt['rows']}x{tt['cols']}",
                           "ir": f"{dims[0]}x{dims[1]}" if dims else "missing",
                           "merges_ok": merges == sorted(tt.get("merges", [])),
                           "ok": dims == (tt["rows"], tt["cols"]) and merges == sorted(
                               tt.get("merges", []))})
    return {"shapes": shapes_out, "tables": tables_out}


def compare_document(doc: Document, truth: dict[str, Any]) -> list[dict[str, Any]]:
    out = []
    for sl in truth["slides"]:
        sc = _scale(doc, truth, sl["index"])
        out.append({**compare_slide(doc, sl, sc), **compare_graphics(doc, sl, sc)})
    return out


def structure_corpus(cfg: Config, only: list[str] | None = None,
                     use_cache: bool = True) -> dict[str, Any]:
    from ..corpus.build import load_manifest
    from ..pipeline import analyzed

    files = []
    for e in load_manifest(cfg.corpus_dir):
        if e["category"] != "synthetic_deck" or (only and e["name"] not in only):
            continue
        truth = json.loads((cfg.corpus_dir / e["truth"]).read_text(encoding="utf-8"))
        doc, _, _ = analyzed(cfg.corpus_dir / e["pdf"], cfg, workers=1, use_cache=use_cache)
        files.append({"name": e["name"], "features": e.get("features", []),
                      "slides": compare_document(doc, truth)})
    slides = [s for f in files for s in f["slides"]]
    n = len(slides)
    runs = sum(s["runs"] for s in slides)
    summary = {
        "slides": n,
        "slides_matching": sum(1 for s in slides if s["match"]),
        "slide_match_rate": round(sum(1 for s in slides if s["match"]) / n, 4) if n else None,
        "runs": runs,
        "runs_font_ok": sum(s["runs_font_ok"] for s in slides),
        "font_match_rate": round(sum(s["runs_font_ok"] for s in slides) / runs, 4) if runs
        else None,
        "bullets_checked": sum(s["bullets_checked"] for s in slides),
        "bullets_ok": sum(s["bullets_ok"] for s in slides),
        "mismatching": [f"{f['name']}#{s['index']}" for f in files for s in f["slides"]
                        if not s["match"]],
    }
    shp = [g for sl in slides for g in sl.get("shapes", [])]
    tbl = [g for sl in slides for g in sl.get("tables", [])]
    summary["shapes"] = len(shp)
    summary["shapes_preset_ok"] = sum(1 for g in shp if g["ok"])
    summary["preset_match_rate"] = (round(summary["shapes_preset_ok"] / len(shp), 4) if shp
                                    else None)
    summary["preset_misses"] = [f"{f['name']}#{sl['index']}: {g['truth']} -> {g['ir']}"
                                for f in files for sl in f["slides"]
                                for g in sl.get("shapes", []) if not g["ok"]]
    summary["tables"] = len(tbl)
    summary["tables_ok"] = sum(1 for g in tbl if g["ok"])
    out = {"summary": summary, "files": files}
    path = cfg.scores_dir / "structure.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8",
                    newline="\n")
    return out


def summary_text(res: dict[str, Any]) -> str:
    s = res["summary"]
    lines = [
        f"slides: {s['slides_matching']}/{s['slides']} match truth paragraphs + levels "
        f"({s['slide_match_rate']:.1%})",
        f"runs:   {s['runs_font_ok']}/{s['runs']} with the truth font family "
        f"({s['font_match_rate']:.1%})",
        f"bullets present/absent: {s['bullets_ok']}/{s['bullets_checked']}",
    ]
    if s.get("shapes"):
        lines.append(f"shapes: {s['shapes_preset_ok']}/{s['shapes']} with the truth preset "
                     f"({s['preset_match_rate']:.1%})"
                     + (f"; misses {s['preset_misses'][:6]}" if s["preset_misses"] else ""))
    if s.get("tables"):
        lines.append(f"tables: {s['tables_ok']}/{s['tables']} with truth rows x cols + merges")
    for f in res["files"]:
        for sl in f["slides"]:
            if not sl["match"] or sl["font_misses"]:
                lines.append(f"  {f['name']}#{sl['index']}: truth {sl['truth_paragraphs']} / "
                             f"ir {sl['ir_paragraphs']} / paired {sl['paired']}, levels "
                             f"{'ok' if sl['levels_ok'] else 'WRONG'}; missing "
                             f"{sl['unpaired_truth'][:3]}, extra {sl['unpaired_ir'][:3]}"
                             + (f"; fonts {sl['font_misses'][:3]}" if sl["font_misses"] else ""))
    return "\n".join(lines)


def write_markdown(res: dict[str, Any], path: Path) -> None:
    path.write_text(summary_text(res) + "\n", encoding="utf-8", newline="\n")
