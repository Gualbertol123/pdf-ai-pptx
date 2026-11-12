"""``pdf2pptx score``: compare a PDF with a PPTX, per slide, using PLAN.md §6 scores.

Scores that need PowerPoint (ssim, region_ssim_min, pos_err_pt, linebreak_match,
COM-based word_recall) are ``null`` when PowerPoint is unavailable, and the
result says why.
"""

from __future__ import annotations

import json
import statistics
import time
from collections import defaultdict
from contextlib import ExitStack
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image

from .. import log
from ..config import Config
from . import _com_temp, metrics
from .metrics import Word
from .pdf_text import page_refs, render_pages
from .pptx_stats import blank_deck, deck_stats

LOG = log.get(__name__)

METRIC_KEYS = [
    "ssim",
    "region_ssim_min",
    "text_sim",
    "word_recall",
    "pos_err_pt",
    "linebreak_match",
    "real_text_ratio",
    "native_shape_ratio",
    "boxes_per_slide",
    "seconds_per_page",
    "composite",
]
HIGHER_IS_BETTER = {k: k not in ("pos_err_pt", "boxes_per_slide", "seconds_per_page")
                    for k in METRIC_KEYS}


def _com_lines(layout: list[_com_temp.LineBox]) -> list[list[Word]]:
    return [[Word(w.text, w.x0, w.y0, w.x1, w.y1) for w in lb.words] for lb in layout]


def _load_png(path: Path, size: tuple[int, int]) -> np.ndarray:
    img = Image.open(path).convert("RGB")
    if img.size != size:
        img = img.resize(size, Image.Resampling.LANCZOS)
    return np.asarray(img)


def _r(v: float | None, nd: int = 4) -> float | None:
    return None if v is None else round(float(v), nd)


def score_pair(
    pdf: Path,
    pptx: Path,
    cfg: Config,
    *,
    ppt: _com_temp.PowerPointSession | None = None,
    name: str | None = None,
    seconds_per_page: float | None = None,
    text_ref_pdf: Path | None = None,
) -> dict[str, Any]:
    """Score ``pptx`` against ``pdf``.

    ``text_ref_pdf``: take reference text and word positions from this PDF instead
    (used for scanned PDFs, which have no text layer: the clean source is the truth).
    """
    name = name or pdf.stem
    t0 = time.perf_counter()
    refs = page_refs(text_ref_pdf or pdf)
    renders = render_pages(pdf, cfg.render_dpi)
    stats = deck_stats(pptx)

    com_error: str | None = None
    pngs: list[Path] = []
    layouts: list[list[_com_temp.LineBox]] = []
    with ExitStack() as stack:
        try:
            if ppt is None:
                ppt = stack.enter_context(_com_temp.session())
            render_dir = cfg.scores_dir / "renders" / name
            pngs = ppt.render(pptx, render_dir, cfg.render_dpi)
            layouts = ppt.text_layout(pptx)
        except Exception as exc:  # PowerPoint missing or failing: degrade, never crash
            com_error = f"{type(exc).__name__}: {exc}"
            LOG.warning("%s: PowerPoint scores unavailable (%s)", name, com_error)

    n = max(len(refs), len(stats.slides))
    slides = []
    for i in range(n):
        row: dict[str, Any] = {"index": i}
        if i >= len(refs) or i >= len(stats.slides):
            row.update({k: 0.0 for k in METRIC_KEYS if HIGHER_IS_BETTER[k]})
            row.update(pos_err_pt=None, boxes_per_slide=None, seconds_per_page=None,
                       note="page/slide missing on one side")
            slides.append(row)
            continue
        ref, st = refs[i], stats.slides[i]
        row["text_sim"] = metrics.text_sim(ref.text, st.text)
        row["real_text_ratio"] = metrics.real_text_ratio(ref.text, st.frame_text)
        row["native_shape_ratio"] = metrics.native_shape_ratio(st.n_preset, st.n_freeform)
        row["boxes_per_slide"] = st.n_text_boxes
        row["seconds_per_page"] = seconds_per_page
        if com_error is None and i < len(pngs):
            h, w = renders[i].shape[:2]
            s, smap = metrics.ssim_map(renders[i], _load_png(pngs[i], (w, h)))
            row["ssim"] = s
            row["region_ssim_min"] = metrics.region_ssim_min(smap)
            pl = _com_lines(layouts[i])
            row["word_recall"] = metrics.word_recall(ref.lines, pl)
            row["pos_err_pt"] = metrics.pos_err_pt(ref.lines, pl)
            row["linebreak_match"] = metrics.linebreak_match(ref.lines, pl)
        else:
            row["ssim"] = row["region_ssim_min"] = None
            row["pos_err_pt"] = row["linebreak_match"] = None
            row["word_recall"] = metrics.token_recall(ref.text, st.text)
        row["composite"] = metrics.composite(row)
        slides.append({k: (_r(v) if isinstance(v, float) else v) for k, v in row.items()})

    return {
        "name": name,
        "pdf": str(pdf),
        "pptx": str(pptx),
        "text_reference": str(text_ref_pdf or pdf),
        "pages": len(refs),
        "slides_in_pptx": len(stats.slides),
        "powerpoint": None if com_error else "ok",
        "powerpoint_error": com_error,
        "render_dpi": cfg.render_dpi,
        "slides": slides,
        "summary": summarize(slides),
        "score_seconds": round(time.perf_counter() - t0, 2),
    }


def _vals(rows: list[dict[str, Any]], key: str) -> list[float]:
    return [float(r[key]) for r in rows if r.get(key) is not None]


def summarize(rows: list[dict[str, Any]]) -> dict[str, Any]:
    out: dict[str, Any] = {"slides": len(rows)}
    for k in METRIC_KEYS:
        v = _vals(rows, k)
        out[k] = {
            "median": _r(statistics.median(v)) if v else None,
            "mean": _r(statistics.fmean(v)) if v else None,
            "min": _r(min(v)) if v else None,
            "n": len(v),
        }
    return out


# --------------------------------------------------------------------------- corpus


def _pptx_for(entry: dict[str, Any], against: str, cfg: Config) -> Path | None:
    if against == "truth":
        rel = entry.get("truth_pptx")
        return cfg.corpus_dir / rel if rel else None
    if against == "blank":
        from .pdf_text import page_refs as _refs

        refs = _refs(cfg.corpus_dir / entry["pdf"])
        return blank_deck(cfg.scores_dir / "blank" / f"{entry['name']}.pptx",
                          [(r.width, r.height) for r in refs])
    p = cfg.out_dir / "corpus" / f"{entry['name']}.pptx"
    return p if p.is_file() else None


def score_corpus(cfg: Config, against: str = "out", label: str | None = None,
                 only: list[str] | None = None) -> dict[str, Any]:
    from ..corpus.build import load_manifest

    label = label or against
    out_dir = cfg.scores_dir / label
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    skipped = []
    with ExitStack() as stack:
        ppt: _com_temp.PowerPointSession | None
        try:
            ppt = stack.enter_context(_com_temp.session())
        except Exception as exc:
            LOG.warning("PowerPoint unavailable: %s", exc)
            ppt = None
        for entry in load_manifest(cfg.corpus_dir):
            if only and entry["name"] not in only:
                continue
            pptx = _pptx_for(entry, against, cfg)
            if pptx is None:
                skipped.append(entry["name"])
                continue
            text_ref = None
            if entry.get("derived_from"):  # scanned: no text layer, use the clean source
                text_ref = cfg.corpus_dir / "synthetic" / f"{entry['derived_from']}.pdf"
            spp = None
            if against == "out":
                meta = cfg.out_dir / "corpus" / f"{entry['name']}.convert.json"
                if meta.is_file():
                    spp = json.loads(meta.read_text(encoding="utf-8")).get("seconds_per_page")
            res = score_pair(cfg.corpus_dir / entry["pdf"], pptx, cfg, ppt=ppt,
                             name=entry["name"], text_ref_pdf=text_ref, seconds_per_page=spp)
            if ppt is None:
                res["powerpoint_error"] = res["powerpoint_error"] or "PowerPoint unavailable"
            res["category"] = entry["category"]
            (out_dir / f"{entry['name']}.json").write_text(json.dumps(res, indent=1),
                                                           encoding="utf-8")
            LOG.info("%-28s ssim=%s text_sim=%s", entry["name"],
                     res["summary"]["ssim"]["median"], res["summary"]["text_sim"]["median"])
            results.append(res)
    summary = corpus_summary(results, label, against, skipped)
    (out_dir / "summary.json").write_text(json.dumps(summary, indent=1), encoding="utf-8")
    md = summary_markdown(summary)
    (out_dir / "summary.md").write_text(md, encoding="utf-8")
    return summary


def corpus_summary(results: list[dict[str, Any]], label: str, against: str,
                   skipped: list[str]) -> dict[str, Any]:
    by_cat: dict[str, list[dict[str, Any]]] = defaultdict(list)
    files = []
    for r in results:
        by_cat[r["category"]].extend(r["slides"])
        by_cat["ALL"].extend(r["slides"])
        files.append({"name": r["name"], "category": r["category"], **{
            k: r["summary"][k]["median"] for k in METRIC_KEYS}, "slides": len(r["slides"])})
    return {
        "label": label,
        "against": against,
        "files": files,
        "categories": {c: summarize(rows) for c, rows in sorted(by_cat.items())},
        "skipped": skipped,
    }


_SHORT = {
    "ssim": "ssim", "region_ssim_min": "reg_min", "text_sim": "text_sim",
    "word_recall": "w_recall", "pos_err_pt": "pos_err", "linebreak_match": "lb_match",
    "real_text_ratio": "real_txt", "native_shape_ratio": "native", "boxes_per_slide": "boxes",
    "seconds_per_page": "s/page", "composite": "composite",
}


def _fmt(v: Any) -> str:
    if v is None:
        return "n/a"
    return f"{v:.3f}" if isinstance(v, float) else str(v)


def summary_markdown(summary: dict[str, Any]) -> str:
    keys = METRIC_KEYS
    lines = [f"# Scores: {summary['label']} (against {summary['against']})", "",
             "## By category (median over slides; mean in parentheses for ssim/text_sim)", "",
             "| category | slides | " + " | ".join(_SHORT[k] for k in keys) + " |",
             "|---|---|" + "---|" * len(keys)]
    for cat, s in summary["categories"].items():
        cells = []
        for k in keys:
            c = _fmt(s[k]["median"])
            if k in ("ssim", "text_sim") and s[k]["mean"] is not None:
                c += f" ({s[k]['mean']:.3f})"
            cells.append(c)
        lines.append(f"| {cat} | {s['slides']} | " + " | ".join(cells) + " |")
    lines += ["", "## By file (median over slides)", "",
              "| file | cat | slides | " + " | ".join(_SHORT[k] for k in keys) + " |",
              "|---|---|---|" + "---|" * len(keys)]
    for f in summary["files"]:
        lines.append(f"| {f['name']} | {f['category']} | {f['slides']} | "
                     + " | ".join(_fmt(f[k]) for k in keys) + " |")
    if summary["skipped"]:
        lines += ["", f"Skipped (no PPTX): {', '.join(summary['skipped'])}"]
    return "\n".join(lines) + "\n"
