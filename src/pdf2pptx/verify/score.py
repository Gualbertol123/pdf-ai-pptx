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
from contextlib import suppress
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image

from .. import log
from ..config import Config
from . import diff, metrics, ppt_com
from .diff import ElementBox
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


def _com_lines(layout: list[ppt_com.TextShape]) -> list[list[Word]]:
    return metrics.words_from_layout(ppt_com.layout_lines(layout))


def _load_png(path: Path) -> np.ndarray:
    return np.asarray(Image.open(path).convert("RGB"))


def _r(v: float | None, nd: int = 4) -> float | None:
    return None if v is None else round(float(v), nd)


def engine_or_reason(cfg: Config) -> tuple[ppt_com.PowerPointEngine | None, str | None]:
    """The shared PowerPoint engine, or (None, why it is unavailable)."""
    try:
        return ppt_com.shared(call_timeout=cfg.com_timeout_s), None
    except Exception as exc:
        reason = f"{type(exc).__name__}: {exc}"
        LOG.warning("PowerPoint unavailable: %s", reason)
        return None, reason


def score_pair(
    pdf: Path,
    pptx: Path,
    cfg: Config,
    *,
    engine: ppt_com.PowerPointEngine | None = None,
    engine_error: str | None = None,
    name: str | None = None,
    seconds_per_page: float | None = None,
    text_ref_pdf: Path | None = None,
    boxes: list[list[ElementBox]] | None = None,
    heatmap_dir: Path | None = None,
    collect: dict[int, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Score ``pptx`` against ``pdf``.

    ``engine``: a running PowerPoint engine; if None the shared one is used (unless
    ``engine_error`` says it is unavailable). ``text_ref_pdf``: take reference text
    and word positions from this PDF instead (scanned PDFs have no text layer: the
    clean source is the truth). ``boxes``: IR element boxes per page, used to name
    the elements under each mismatch region. ``heatmap_dir``: write one diff heatmap
    PNG per slide there. ``collect``: filled per slide with the PowerPoint render
    path, text layout and :class:`diff.SlideDiff` (for the debug report).
    """
    name = name or pdf.stem
    t0 = time.perf_counter()
    refs = page_refs(text_ref_pdf or pdf)
    renders = render_pages(pdf, cfg.render_dpi)
    stats = deck_stats(pptx)

    com_error: str | None = engine_error
    if engine is None and com_error is None:
        engine, com_error = engine_or_reason(cfg)
    info: ppt_com.PresInfo | None = None
    if engine is not None and com_error is None:
        try:
            info = engine.open(pptx)
        except Exception as exc:  # PowerPoint missing or failing: degrade, never crash
            com_error = f"{type(exc).__name__}: {exc}"
            LOG.warning("%s: PowerPoint scores unavailable (%s)", name, com_error)

    render_dir = cfg.scores_dir / "renders" / name
    n = max(len(refs), len(stats.slides))
    slides = []
    com_times: list[float] = []
    try:
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
            got = None
            if engine is not None and info is not None and com_error is None and i < info.slides:
                try:
                    tc = time.perf_counter()
                    png = engine.render(info, [i], cfg.render_dpi, render_dir)[0]
                    layout = engine.text_layout(info, i)
                    com_times.append(time.perf_counter() - tc)
                    got = (png, layout)
                except Exception as exc:
                    com_error = f"{type(exc).__name__}: {exc}"
                    LOG.warning("%s: PowerPoint scores unavailable (%s)", name, com_error)
            if got is not None and info is not None:
                png, layout = got
                d = diff.compare(renders[i], _load_png(png), (ref.width, ref.height),
                                 slide_size=(info.width_pt, info.height_pt),
                                 boxes=boxes[i] if boxes and i < len(boxes) else ())
                row["ssim"] = d.ssim
                row["region_ssim_min"] = d.region_ssim_min
                pl = _com_lines(layout)
                if abs(info.width_pt - ref.width) > 0.05 or abs(info.height_pt - ref.height) > 0.05:
                    pl = metrics.map_lines(pl, *diff.page_rect_on_slide(
                        ref.width, ref.height, info.width_pt, info.height_pt))
                row["word_recall"] = metrics.word_recall(ref.lines, pl)
                row["pos_err_pt"] = metrics.pos_err_pt(ref.lines, pl)
                row["linebreak_match"] = metrics.linebreak_match(ref.lines, pl)
                row["regions"] = [[_r(v) for v in rr] for rr in d.region_scores]
                row["mismatches"] = [m.to_dict() for m in d.mismatches[:10]]
                if heatmap_dir is not None:
                    hm = diff.write_heatmap(d, heatmap_dir / f"heat_{i + 1:03d}.png", renders[i])
                    row["heatmap"] = str(hm)
                if collect is not None:
                    collect[i] = {"png": png, "layout": layout, "diff": d, "lines": pl}
            else:
                row["ssim"] = row["region_ssim_min"] = None
                row["pos_err_pt"] = row["linebreak_match"] = None
                row["word_recall"] = metrics.token_recall(ref.text, st.text)
            row["composite"] = metrics.composite(row)
            slides.append({k: (_r(v) if isinstance(v, float) else v) for k, v in row.items()})
    finally:
        if info is not None and engine is not None:
            with suppress(Exception):
                engine.close(info)

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
        "slide_size": [info.width_pt, info.height_pt] if info is not None else None,
        "slides": slides,
        "summary": summarize(slides),
        "com_seconds_per_slide": [round(t, 3) for t in com_times],
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
    engine, engine_error = engine_or_reason(cfg)
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
        res = score_pair(cfg.corpus_dir / entry["pdf"], pptx, cfg, engine=engine,
                         engine_error=engine_error, name=entry["name"], text_ref_pdf=text_ref,
                         seconds_per_page=spp)
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
    com = sorted(t for r in results for t in r.get("com_seconds_per_slide", []))
    timing = {"n": len(com), "median": _r(statistics.median(com), 3) if com else None,
              "mean": _r(statistics.fmean(com), 3) if com else None,
              "max": _r(com[-1], 3) if com else None}
    return {
        "label": label,
        "against": against,
        "powerpoint_render_layout_seconds_per_slide": timing,
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
    t = summary.get("powerpoint_render_layout_seconds_per_slide") or {}
    if t.get("n"):
        lines += ["", f"PowerPoint render + text layout per slide: median {t['median']:.3f} s, "
                      f"mean {t['mean']:.3f} s, max {t['max']:.3f} s (n={t['n']})"]
    if summary["skipped"]:
        lines += ["", f"Skipped (no PPTX): {', '.join(summary['skipped'])}"]
    return "\n".join(lines) + "\n"
