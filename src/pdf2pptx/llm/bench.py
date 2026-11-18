"""``pdf2pptx bench``: every installed LLM candidate on a fixed task set (M9).

Tasks (all from the corpus, scored against its truth):

* 10 judgment items: block roles on 4 pages (r01, r02, r03, d17), reading order on 2
  pages (r01 two columns, d04 three columns), table structure on 3 tables (d07,
  d08 merged, r02) and region classification on 1 slide (d10 chevrons: ``process``).
* 5 chart regions: d14 (column, 3 series), d14 (bar, data labels), d15 (line),
  d16 (pie), d27 (Italian column).
* 2 text items (text-capable models): a grounded deck outline of r01 and of d27.

Vision / text models get the JSON-schema prompts of ``llm/schemas.py`` (the same as in
conversion). Document models (granite-docling, PaddleOCR-VL) do not follow free
prompts: they get their native task prompts and their table / chart output (OTSL,
markdown or HTML) is parsed.

Per model: load time, seconds per image request, prompt and generation speed (from
llama.cpp's timings), peak RAM (available before the start - lowest available while
it runs), accuracy per task family. Writes ``bench/results.md`` + ``results.json`` and
the recommended stack to ``pdf2pptx.toml`` (``llm_judge`` / ``llm_chart`` /
``llm_doc`` / ``llm_text``).
"""

from __future__ import annotations

import json
import re
import statistics
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import cv2
import numpy as np
from rapidfuzz import fuzz

from .. import log
from ..analyze.charts import parse_number
from ..analyze.roles import block_text
from ..config import PROJECT_ROOT, Config
from ..ir import Document, Table, TextBlock
from . import charts_vlm, schemas
from .server import LlamaServer, LLMUnavailable, reserve_bytes
from .setup import candidates, installed

LOG = log.get(__name__)

BENCH_DIR = PROJECT_ROOT / "bench"
DPI = 150
ROLE_MAP = {"text": None, "heading": "heading", "body": "body"}  # truth -> ours
REGION_TRUTH = {"d10_arrows_chevrons": "process"}
DOC_PROMPTS = {
    "paddleocr_vl": {"table": "Table Recognition:", "chart": "Chart Recognition:"},
    "granite_docling": {"table": "Convert table to OTSL.", "chart": "Convert chart to table."},
}


@dataclass
class Item:
    kind: str  # roles | order | table | region | chart | text
    name: str
    image: bytes = b""
    prompt: str = ""
    schema: dict[str, Any] | None = None
    truth: Any = None
    data: dict[str, Any] = field(default_factory=dict)


# --------------------------------------------------------------------------- task set


def _doc(cfg: Config, name: str) -> tuple[Document, Path, Path, dict[str, Any]]:
    from ..pipeline import analyzed

    pdf = cfg.corpus_dir / "synthetic" / f"{name}.pdf"
    doc, work_dir, _ = analyzed(pdf, cfg, workers=1)
    truth = json.loads((cfg.corpus_dir / "synthetic" / f"{name}.truth.json").read_text(
        encoding="utf-8"))
    return doc, pdf, work_dir, truth


def _render(pdf: Path, work_dir: Path, i: int) -> np.ndarray:
    from ..ingest.render import ensure_render

    img = cv2.imread(str(ensure_render(pdf, work_dir, i, DPI)), cv2.IMREAD_COLOR)
    assert img is not None
    return img


def _crop(img: np.ndarray, bbox: tuple[float, float, float, float], pad: float = 4.0) -> bytes:
    return charts_vlm.crop_png(img, bbox, DPI / 72, pad=pad)


def _page_png(img: np.ndarray) -> bytes:
    h, w = img.shape[:2]
    return charts_vlm.crop_png(img, (0, 0, w * 72 / DPI, h * 72 / DPI), DPI / 72, pad=0)


def _truth_page(truth: dict[str, Any], i: int) -> dict[str, Any]:
    pages = truth.get("slides") or truth.get("pages") or []
    page: dict[str, Any] = pages[i]
    return page


def build_items(cfg: Config) -> list[Item]:
    items: list[Item] = []
    # roles
    for name in ("r01_report_two_column", "r02_report_tables_it", "r03_report_figure",
                 "d17_footer_logo_number"):
        doc, pdf, wd, truth = _doc(cfg, name)
        page = doc.pages[0]
        tp = _truth_page(truth, 0)
        blocks = [b for b in page.elements if isinstance(b, TextBlock)]
        want: dict[str, str] = {}
        for b in blocks:
            t = block_text(b)
            best = max(tp["paragraphs"], key=lambda p: fuzz.ratio(p["text"], t), default=None)
            if best is None or fuzz.ratio(best["text"], t) < 80:
                continue
            role = ROLE_MAP.get(best["role"], best["role"])
            if role:
                want[b.id] = role
        sel = [b for b in blocks if b.id in want][:12]
        items.append(Item("roles", f"{name}:p0", _page_png(_render(pdf, wd, 0)),
                          schemas.role_prompt([(b.id, block_text(b), "?") for b in sel]),
                          schemas.ROLE_FIX, {b.id: want[b.id] for b in sel}))
    # reading order
    for name in ("r01_report_two_column", "d04_three_columns"):
        doc, pdf, wd, truth = _doc(cfg, name)
        page = doc.pages[0]
        tp = _truth_page(truth, 0)
        blocks = [b for b in page.elements if isinstance(b, TextBlock)
                  and b.role not in ("footer", "page_number")]
        texts = [p["text"] for p in tp["paragraphs"]]

        def rank(b: TextBlock, texts: list[str] = texts) -> int:
            t = block_text(b)
            scores = [fuzz.partial_ratio(x[:80], t[:80]) for x in texts]
            return scores.index(max(scores))

        # shuffle-free but not in truth order: sort by x then y (column-blind)
        shown = sorted(blocks, key=lambda b: (round(b.bbox[1] / 40), b.bbox[0]))
        from .judge import _numbered

        img = _render(pdf, wd, 0)
        truth_order = sorted(range(len(shown)), key=lambda k: (rank(shown[k]), shown[k].bbox[1]))
        items.append(Item("order", f"{name}:p0",
                          _numbered(img, [b.bbox for b in shown]),
                          schemas.order_prompt([(k, block_text(b)) for k, b in enumerate(shown)]),
                          schemas.READING_ORDER, truth_order))
    # tables
    for name, k in (("d07_table_simple", 0), ("d08_table_merged", 0),
                    ("r02_report_tables_it", 0)):
        doc, pdf, wd, truth = _doc(cfg, name)
        tp = _truth_page(truth, 0)
        tt = tp["tables"][k]
        tabs = sorted((e for e in doc.pages[0].elements if isinstance(e, Table)),
                      key=lambda t: t.bbox[1])
        if tabs:
            bbox = tabs[min(k, len(tabs) - 1)].bbox
            from .judge import _table_rows

            rows = _table_rows(tabs[min(k, len(tabs) - 1)])
        else:
            box = next(s["bbox_pt"] for s in tp["shape_list"] if s["kind"] == "table")
            bbox, rows = tuple(box), [[""]]
        items.append(Item("table", f"{name}:t{k}", _crop(_render(pdf, wd, 0), bbox),
                          schemas.table_prompt(rows), schemas.TABLE_FIX, tt))
    # region
    for name, label in REGION_TRUTH.items():
        doc, pdf, wd, truth = _doc(cfg, name)
        img = _render(pdf, wd, 0)
        h, w = img.shape[:2]
        items.append(Item("region", f"{name}:p0", _crop(img, (0, 100, w * 72 / DPI,
                                                                h * 72 / DPI), 0),
                          schemas.region_prompt("slide body", 8), schemas.REGION, label))
    # charts
    for name, i, k in (("d14_chart_bar", 0, 0), ("d14_chart_bar", 1, 0), ("d15_chart_line", 0, 0),
                       ("d16_chart_pie", 0, 0), ("d27_italian_report_4x3", 1, 0)):
        doc, pdf, wd, truth = _doc(cfg, name)
        tp = _truth_page(truth, i)
        box = [s["bbox_pt"] for s in tp["shape_list"] if s["kind"] == "chart"][k]
        items.append(Item("chart", f"{name}:p{i}", _crop(_render(pdf, wd, i), tuple(box), 0),
                          schemas.chart_prompt(), schemas.CHART, tp["charts"][k]))
    # text: grounded outline
    for name in ("r01_report_two_column", "d27_italian_report_4x3"):
        doc, pdf, wd, truth = _doc(cfg, name)
        blocks = [(b.id, block_text(b)) for p in doc.pages for b in p.elements
                  if isinstance(b, TextBlock) and len(block_text(b)) > 2]
        src = "\n".join(f"[{bid}] {t}" for bid, t in blocks)
        prompt = ("Source document blocks (id in brackets):\n" + src[:6000] + "\n\n"
                  "Plan a 3-slide summary deck. Each slide: a title and 2-4 short bullets; "
                  "every bullet lists the ids of the blocks it comes from. Use only facts "
                  "and numbers from the source.")
        items.append(Item("text", name, b"", prompt, OUTLINE,
                          {"ids": [b for b, _ in blocks], "text": src}))
    return items


OUTLINE = {
    "type": "object",
    "properties": {"slides": {"type": "array", "items": {
        "type": "object",
        "properties": {"title": {"type": "string"},
                       "bullets": {"type": "array", "items": {
                           "type": "object",
                           "properties": {"text": {"type": "string"},
                                          "sources": {"type": "array",
                                                      "items": {"type": "string"}}},
                           "required": ["text", "sources"]}}},
        "required": ["title", "bullets"]}}},
    "required": ["slides"],
}


# --------------------------------------------------------------------------- scoring


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())


def score_roles(item: Item, out: Any) -> float:
    v = schemas.valid_roles(out, set(item.truth))
    if not v:
        return 0.0
    ok = 0
    for bid, want in item.truth.items():
        got = v.get(bid, ("", 0))[0]
        if got == want or (want == "heading" and got in ("subtitle", "heading", "title")):
            ok += 1
    return ok / len(item.truth)


def score_order(item: Item, out: Any) -> float:
    v = schemas.valid_order(out, len(item.truth))
    if v is None:
        return 0.0
    pos = {b: k for k, b in enumerate(v[0])}
    want = item.truth
    pairs = [(want[a], want[b]) for a in range(len(want)) for b in range(a + 1, len(want))]
    return sum(1 for x, y in pairs if pos[x] < pos[y]) / max(len(pairs), 1)


def _cell_sim(a: list[list[str]], b: list[list[str]]) -> float:
    fa = [_norm(c) for r in a for c in r]
    fb = [_norm(c) for r in b for c in r]
    return fuzz.ratio(" | ".join(fa), " | ".join(fb)) / 100.0


def score_table(item: Item, rows: int, cols: int, cells: list[list[str]]) -> float:
    t = item.truth
    shape = 0.5 * (rows == t["rows"]) + 0.5 * (cols == t["cols"])
    return 0.5 * shape + 0.5 * _cell_sim(cells, t["cells"])


def score_chart(item: Item, ctype: str, cats: list[str], series: list[list[float | None]]
                ) -> dict[str, float]:
    t = item.truth
    tt = t["type"].lower()
    type_ok = float(charts_vlm.family(ctype) == charts_vlm.family(tt))
    tc = [_norm(c) for c in t["categories"]]
    gc = [_norm(c) for c in cats]
    if tt.startswith("bar") and gc == list(reversed(tc)):
        gc = tc
        series = [list(reversed(s)) for s in series]
    cats_ok = float(gc == tc)
    rng = max(abs(v) for s in t["series"] for v in s["values"]) or 1.0
    hit = tot = 0
    for k, s in enumerate(t["series"]):
        got = series[k] if k < len(series) else []
        for i, v in enumerate(s["values"]):
            tot += 1
            g = got[i] if i < len(got) else None
            if g is not None and abs(g - v) <= 0.02 * rng:
                hit += 1
    vals = hit / max(tot, 1)
    return {"type": type_ok, "categories": cats_ok, "values": vals,
            "score": (type_ok + cats_ok + vals) / 3}


def score_text(item: Item, out: Any) -> float:
    if not isinstance(out, dict) or not isinstance(out.get("slides"), list) or not out["slides"]:
        return 0.0
    ids = set(item.truth["ids"])
    src_nums = set(re.findall(r"\d+(?:[.,]\d+)?", item.truth["text"]))
    n = good = 0
    for s in out["slides"]:
        for b in s.get("bullets", []):
            n += 1
            cited = [x for x in b.get("sources", []) if x in ids]
            nums = re.findall(r"\d+(?:[.,]\d+)?", b.get("text", ""))
            if cited and all(x in src_nums for x in nums):
                good += 1
    return good / n if n else 0.0


# --------------------------------------------------------------------------- doc models


def parse_grid(text: str) -> list[list[str]]:
    """A table from OTSL tokens, an HTML table or a markdown table."""
    if "<fcel>" in text or "<ecel>" in text or "<nl>" in text:
        rows: list[list[str]] = [[]]
        for tok, val in re.findall(r"<(fcel|ecel|lcel|ucel|xcel|ched|rhed|srow|nl)>([^<]*)",
                                   text):
            if tok == "nl":
                rows.append([])
            elif tok in ("fcel", "ched", "rhed", "srow"):
                rows[-1].append(val.strip())
            else:
                rows[-1].append("")
        return [r for r in rows if r]
    if "<tr" in text.lower():
        rows = []
        for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", text, re.S | re.I):
            rows.append([re.sub(r"<[^>]+>", "", c).strip()
                         for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S | re.I)])
        return [r for r in rows if r]
    rows = []
    for line in text.splitlines():
        if "|" not in line or re.fullmatch(r"[\s|:\-]+", line):
            continue
        rows.append([c.strip() for c in line.strip().strip("|").split("|")])
    return rows


def grid_to_chart(grid: list[list[str]]) -> tuple[list[str], list[list[float | None]]]:
    """Header row = series names; first column = categories (or transposed)."""
    if len(grid) < 2:
        return [], []
    body = grid[1:]
    cats = [r[0] for r in body if r]
    ncol = max(len(r) for r in grid)
    series = []
    for c in range(1, ncol):
        series.append([parse_number(r[c]) if c < len(r) else None for r in body])
    numeric_first = sum(parse_number(c) is not None for c in cats)
    if numeric_first > len(cats) / 2 and len(grid[0]) > 2:  # transposed
        cats = grid[0][1:]
        series = [[parse_number(c) for c in r[1:]] for r in body]
    return cats, series


# --------------------------------------------------------------------------- run


def _run_model(c: dict[str, Any], items: list[Item], log_dir: Path) -> dict[str, Any]:
    import psutil

    from .server import RESERVE_BYTES, plan_memory

    name = c["name"]
    res: dict[str, Any] = {"model": name, "kind": c["kind"], "tasks": {}, "items": []}
    avail = psutil.virtual_memory().available
    default = plan_memory(name, int(c.get("ctx", 8192)), available=avail + 0)
    # would it start under the default 8 GB reserve with the RAM free right now?
    res["fits_default_reserve"] = default.need <= avail - RESERVE_BYTES
    res["need_gb"] = round(default.need / 2**30, 2)
    res["available_gb_at_start"] = round(avail / 2**30, 2)
    srv = LlamaServer(name)
    t0 = time.perf_counter()
    try:
        srv.start(log_dir=log_dir)
    except LLMUnavailable as exc:
        res["error"] = str(exc)
        res["memory"] = srv.plan.as_dict() if srv.plan else None
        return res
    res["memory"] = srv.plan.as_dict() if srv.plan else None
    img_secs: list[float] = []
    try:
        for it in items:
            if c["kind"] == "text" and it.kind != "text":
                continue
            if c["kind"] == "docvlm" and it.kind not in ("table", "chart"):
                continue
            t1 = time.perf_counter()
            entry: dict[str, Any] = {"item": it.name, "kind": it.kind}
            try:
                if c["kind"] == "docvlm":
                    r = srv.chat(DOC_PROMPTS[c["model"]][it.kind], images=[it.image],
                                 max_tokens=1500)
                    grid = parse_grid(r["content"])
                    if it.kind == "table":
                        cols = max((len(x) for x in grid), default=0)
                        entry["score"] = score_table(it, len(grid), cols,
                                                     [x + [""] * (cols - len(x)) for x in grid])
                    else:
                        cats, ser = grid_to_chart(grid)
                        guess = it.truth["type"].lower() if cats else "other"
                        sc = score_chart(it, guess, cats, ser)
                        sc["type"] = 0.0  # no type from the model: not credited
                        sc["score"] = (sc["categories"] + sc["values"]) / 3
                        entry |= {"score": sc["score"], "detail": sc}
                    entry["raw"] = r["content"][:400]
                else:
                    imgs = [it.image] if it.image and c["kind"] == "vlm" else []
                    if it.kind != "text" and not imgs:
                        continue
                    r = srv.chat(it.prompt, images=imgs, schema=it.schema,
                                 system=schemas.SYSTEM, max_tokens=1500)
                    out = r["json"]
                    if it.kind == "roles":
                        entry["score"] = score_roles(it, out)
                    elif it.kind == "order":
                        entry["score"] = score_order(it, out)
                    elif it.kind == "table":
                        v = schemas.valid_table(out)
                        entry["score"] = 0.0 if v is None else score_table(
                            it, v["rows"], v["cols"], v["cells"])
                    elif it.kind == "region":
                        v2 = schemas.valid_region(out)
                        entry["score"] = float(v2 is not None and v2["label"] == it.truth)
                        entry["label"] = v2["label"] if v2 else None
                    elif it.kind == "chart":
                        v3 = schemas.valid_chart(out)
                        if v3 is None:
                            entry["score"] = 0.0
                        else:
                            sc = score_chart(it, v3["chart_type"], v3["categories"],
                                             [s["values"] for s in v3["series"]])
                            entry |= {"score": sc["score"], "detail": sc}
                    elif it.kind == "text":
                        entry["score"] = score_text(it, out)
                    entry["valid_json"] = out is not None
                    if not entry.get("score"):
                        entry["raw"] = r["content"][:600]
                t = r.get("timings") or {}
                entry["seconds"] = round(time.perf_counter() - t1, 2)
                entry["prompt_tokens"] = t.get("prompt_n")
                if it.image:
                    img_secs.append(entry["seconds"])
            except Exception as exc:
                entry["error"] = f"{type(exc).__name__}: {exc}"
                entry["score"] = 0.0
                LOG.warning("%s %s: %s", name, it.name, exc)
            res["items"].append(entry)
    finally:
        srv.stop()
    res["server"] = srv.stats.as_dict() | {"n_cpu_moe": srv.plan.n_cpu_moe if srv.plan else 0}
    res["seconds"] = round(time.perf_counter() - t0, 1)
    res["s_per_image_request"] = round(statistics.median(img_secs), 2) if img_secs else None
    fam: dict[str, list[float]] = {}
    for e in res["items"]:
        fam.setdefault(_family(e["kind"]), []).append(float(e.get("score", 0.0)))
    res["tasks"] = {k: round(sum(v) / len(v), 3) for k, v in sorted(fam.items())}
    return res


def _family(kind: str) -> str:
    return {"roles": "judgment", "order": "judgment", "region": "judgment"}.get(kind, kind)


def recommend(results: list[dict[str, Any]]) -> dict[str, str | None]:
    """Best model per role among those that start under the default 8 GB reserve with
    the RAM free during the benchmark (all measured models if none does). Accuracy
    first; within 0.02 the judge model wins (no model swap during a conversion), then
    the faster one."""
    ok = [r for r in results if not r.get("error")]
    fit = [r for r in ok if r.get("fits_default_reserve")]
    ok = fit or ok

    def judge_score(r: dict[str, Any]) -> float:
        t = r["tasks"]
        return (t.get("judgment", 0) + t.get("table", 0)) / 2

    vlm = [r for r in ok if r["kind"] == "vlm"]
    judge = None
    if vlm:
        top = max(judge_score(r) for r in vlm)
        judge = str(min((r for r in vlm if judge_score(r) >= top - 0.02),
                        key=lambda r: r["seconds"])["model"])

    def best(kind: tuple[str, ...], task: str) -> str | None:
        cand = [r for r in ok if r["kind"] in kind and task in r["tasks"]]
        if not cand:
            return None
        top = max(r["tasks"][task] for r in cand)
        near = [r for r in cand if r["tasks"][task] >= top - 0.02]
        if any(r["model"] == judge for r in near):
            return judge
        return str(min(near, key=lambda r: r["seconds"])["model"])

    return {"llm_judge": judge, "llm_chart": best(("vlm",), "chart"),
            "llm_doc": best(("docvlm",), "table"), "llm_text": best(("text", "vlm"), "text")}


def write_results(results: list[dict[str, Any]], stack: dict[str, str | None],
                  meta: dict[str, Any]) -> Path:
    BENCH_DIR.mkdir(parents=True, exist_ok=True)
    (BENCH_DIR / "results.json").write_text(json.dumps(
        {"meta": meta, "stack": stack, "results": results}, indent=1, ensure_ascii=False)
        + "\n", encoding="utf-8", newline="\n")

    def f(v: Any, nd: int = 2) -> str:
        return "-" if v is None else (f"{v:.{nd}f}" if isinstance(v, float) else str(v))

    def overall(r: dict[str, Any]) -> float:
        t = r.get("tasks") or {}
        return sum(t.values()) / len(t) if t else 0.0

    ranked = sorted(results, key=lambda r: (bool(r.get("error")), -overall(r), r.get(
        "seconds") or 1e9))
    lines = [
        "# LLM benchmark (pdf2pptx bench)", "",
        f"Machine: {meta.get('machine')}. llama.cpp {meta.get('llama_cpp')} (Vulkan). "
        f"Run {meta.get('date')}, {meta.get('seconds')} s in total.", "",
        "Tasks: 10 judgment items (4 role pages, 2 reading-order pages, 3 tables, "
        "1 region), 5 chart regions, 2 grounded text outlines. Accuracy 0-1 against the "
        "corpus truth; document models answer tables / charts only, text models text only.",
        "",
        f"RAM reserve during the run: {meta.get('reserve_gb')} GB (default 8). 'fits 8 GB' = "
        "the model would start under the default reserve with the RAM that was free.", "",
        "| rank | model | kind | load s | s / image req | prompt tok/s | gen tok/s | "
        "peak RAM GB | RAM need GB | fits 8 GB | --n-cpu-moe | judgment | table | chart | "
        "text | overall | total s |",
        "|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|",
    ]
    for k, r in enumerate(ranked, 1):
        sv, t = r.get("server") or {}, r.get("tasks") or {}
        if r.get("error"):
            lines.append(f"| - | {r['model']} | {r['kind']} | not run: "
                         f"{r['error'][:120]} |||||||||||||")
            continue
        lines.append(
            f"| {k} | {r['model']} | {r['kind']} | {f(sv.get('load_s'), 1)} | "
            f"{f(r.get('s_per_image_request'), 1)} | {f(sv.get('prompt_tok_s'), 1)} | "
            f"{f(sv.get('gen_tok_s'), 1)} | {f(sv.get('peak_ram_used_gb'))} | "
            f"{f(r.get('need_gb'))} | {'yes' if r.get('fits_default_reserve') else 'no'} | "
            f"{sv.get('n_cpu_moe', 0)} | {f(t.get('judgment'))} | {f(t.get('table'))} | "
            f"{f(t.get('chart'))} | {f(t.get('text'))} | {f(overall(r))} | "
            f"{f(r.get('seconds'), 0)} |")
    lines += ["", "## Recommended stack (written to pdf2pptx.toml)", ""]
    lines += [f"- `{k}` = {v or '(none)'}" for k, v in stack.items()]
    lines += ["", "## Per item", "",
              "| model | item | score | s | detail |", "|---|---|---|---|---|"]
    for r in ranked:
        for e in r.get("items", []):
            det = e.get("detail") or e.get("label") or e.get("error") or ""
            lines.append(f"| {r['model']} | {e['item']} | {f(e.get('score'))} | "
                         f"{f(e.get('seconds'), 1)} | {json.dumps(det) if det else ''} |")
    out = BENCH_DIR / "results.md"
    out.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    return out


def write_stack(stack: dict[str, str | None], path: Path | None = None) -> Path:
    """Set the ``llm_*`` keys of ``pdf2pptx.toml`` (other content is kept)."""
    path = path or PROJECT_ROOT / "pdf2pptx.toml"
    text = path.read_text(encoding="utf-8") if path.is_file() else ""
    lines = [ln for ln in text.splitlines() if not re.match(r"\s*llm_(judge|chart|doc|text)\s*=",
                                                            ln)]
    if not any(ln.strip() == "[pdf2pptx]" for ln in lines):
        lines = ["[pdf2pptx]", *lines]
    at = next(i for i, ln in enumerate(lines) if ln.strip() == "[pdf2pptx]") + 1
    new = ["# written by `pdf2pptx bench` (bench/results.md)"] if not any(
        "pdf2pptx bench" in ln for ln in lines) else []
    new += [f'{k} = "{v}"' for k, v in stack.items() if v]
    lines[at:at] = new
    path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    return path


def run(cfg: Config, only: list[str] | None = None, write_config: bool = True
        ) -> dict[str, Any]:
    import platform

    from .setup import llama_version

    t0 = time.perf_counter()
    items = build_items(cfg)
    prev: dict[str, dict[str, Any]] = {}
    pj = BENCH_DIR / "results.json"
    if only and pj.is_file():  # partial re-run: keep the other models' results
        prev = {r["model"]: r for r in json.loads(pj.read_text(encoding="utf-8"))["results"]}
    results = []
    for c in candidates():
        if only and c["name"] not in only:
            if c["name"] in prev:
                results.append(prev[c["name"]])
            continue
        if not installed(c):
            results.append({"model": c["name"], "kind": c["kind"],
                            "error": "not installed (pdf2pptx setup llm)"})
            continue
        LOG.info("bench: %s", c["name"])
        results.append(_run_model(c, items, cfg.work_dir / "llm_logs"))
        # save after every model: a stopped run keeps what it measured (--only resumes)
        BENCH_DIR.mkdir(parents=True, exist_ok=True)
        pj.write_text(json.dumps({"meta": {"partial": True}, "stack": {},
                                  "results": results}, indent=1, ensure_ascii=False) + "\n",
                      encoding="utf-8", newline="\n")
    stack = recommend(results)
    import psutil

    meta = {"date": time.strftime("%Y-%m-%d %H:%M"), "seconds": round(time.perf_counter() - t0),
            "llama_cpp": (llama_version() or {}).get("tag"),
            "machine": f"{platform.processor() or platform.machine()}, "
                       f"{psutil.virtual_memory().total / 2**30:.1f} GB RAM",
            "items": [i.name for i in items],
            "reserve_gb": round(reserve_bytes() / 2**30, 1)}
    md = write_results(results, stack, meta)
    cfgp = write_stack(stack) if write_config and any(stack.values()) else None
    return {"results": md, "config": str(cfgp) if cfgp else None, "stack": stack}


# --------------------------------------------------------------------------- long deck


LONG_DECK_SOURCES = ["d01_title_bullets", "d02_multilevel_bullets", "d03_two_columns",
                     "d05_mixed_runs", "d06_italian_text", "d07_table_simple",
                     "d08_table_merged", "d09_shapes_basic", "d10_arrows_chevrons",
                     "d12_gradients_alpha", "d13_images", "d14_chart_bar", "d15_chart_line",
                     "d16_chart_pie", "d17_footer_logo_number", "d19_hyperlinks",
                     "d20_speaker_notes", "d21_sections", "d23_long_paragraphs",
                     "d24_font_families", "d25_dark_background", "d26_kitchen_sink"]


def build_long_deck(cfg: Config, pages: int = 30) -> Path:
    """A ``pages``-page 16:9 deck made of corpus slides (same order every time)."""
    import pypdfium2 as pdfium

    out = cfg.work_dir / "bench" / f"deck{pages}.pdf"
    out.parent.mkdir(parents=True, exist_ok=True)
    dst = pdfium.PdfDocument.new()
    n = 0
    while n < pages:
        for name in LONG_DECK_SOURCES:
            src = pdfium.PdfDocument(str(cfg.corpus_dir / "synthetic" / f"{name}.pdf"))
            take = min(len(src), pages - n)
            dst.import_pages(src, list(range(take)))
            n += take
            src.close()
            if n >= pages:
                break
    dst.save(str(out))
    dst.close()
    return out


def long_deck(cfg: Config, pages: int = 30) -> dict[str, Any]:
    """Convert the long deck in the Max tier; total time and peak RAM."""
    import threading

    import psutil

    from ..pipeline import convert

    pdf = build_long_deck(cfg, pages)
    low = [psutil.virtual_memory().available]
    before = low[0]
    stop = threading.Event()

    def sample() -> None:
        while not stop.wait(0.25):
            low[0] = min(low[0], psutil.virtual_memory().available)

    th = threading.Thread(target=sample, daemon=True)
    th.start()
    t0 = time.perf_counter()
    try:
        res = convert(pdf, pdf.with_suffix(".pptx"), cfg, use_cache=False)
    finally:
        stop.set()
        th.join(2)
    llm = res.llm or {}
    return {"pdf": pdf.name, "pages": res.pages, "seconds": round(time.perf_counter() - t0, 1),
            "seconds_per_page": round((time.perf_counter() - t0) / max(res.pages, 1), 2),
            "tier": cfg.resolved_tier(), "llm_seconds": llm.get("seconds_total_batch"),
            "models": llm.get("models"), "applied": llm.get("applied"),
            "keep_best": llm.get("keep_best"),
            "ram_available_before_gb": round(before / 2**30, 2),
            "ram_lowest_available_gb": round(low[0] / 2**30, 2),
            "peak_ram_used_gb": round((before - low[0]) / 2**30, 2),
            "total_ram_gb": round(psutil.virtual_memory().total / 2**30, 1)}
