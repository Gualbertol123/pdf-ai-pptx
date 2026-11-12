"""``pdf2pptx report``: one self-contained HTML debug page per PDF.

Per page: PDF render | PowerPoint render | diff heatmap, with IR bounding boxes
drawn over all three (toggle by element kind and by source, colour by either,
hover shows the element's JSON), PowerPoint's own line boxes, the mismatch
regions, the page's triage result and its score row. On top: the document
status and a sortable per-page score table.

No external assets: images are inlined as data URIs, CSS and JS are inline.
If PowerPoint is unavailable the report still renders (PDF panel, overlays,
text scores) and says why the rest is missing.
"""

from __future__ import annotations

import base64
import html
import io
import json
import time
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image

from .. import log
from ..analyze.text_group import LOW_CONFIDENCE
from ..config import Config
from ..ir import Document, to_dict
from ..verify import diff, ppt_com, score
from ..verify.pdf_text import render_pages

LOG = log.get(__name__)
LOW_CONF_SHOWN = LOW_CONFIDENCE

KIND_COLORS = {
    "text_block": "#1f6fd1", "shape": "#e07b00", "image": "#1a9a4a", "table": "#8a3fd1",
    "chart": "#0f9a9a", "equation": "#b0336b", "group": "#6b6b6b", "annotation": "#d1338a",
    "run": "#4f8fe0", "paragraph": "#3f7fd0", "glyph": "#999999",
}
SOURCE_COLORS = {"pdf": "#555f6d", "rule": "#1f6fd1", "ocr": "#e07b00"}
EXTRA_COLORS = ["#c0392b", "#16a085", "#8e44ad", "#d35400", "#2c3e50", "#27ae60"]
MAX_OVERLAYS = 4000
JPEG_QUALITY = 82


def _data_uri(img: np.ndarray | Image.Image, fmt: str = "JPEG") -> str:
    im = img if isinstance(img, Image.Image) else Image.fromarray(np.asarray(img))
    buf = io.BytesIO()
    if fmt == "JPEG":
        im.convert("RGB").save(buf, "JPEG", quality=JPEG_QUALITY, optimize=True)
        mime = "image/jpeg"
    else:
        im.save(buf, "PNG", optimize=True)
        mime = "image/png"
    return f"data:{mime};base64,{base64.b64encode(buf.getvalue()).decode('ascii')}"


def _compact(d: Any, depth: int = 0) -> Any:
    """Element JSON for the hover box: long lists are cut (paths can be huge)."""
    if isinstance(d, dict):
        return {k: _compact(v, depth + 1) for k, v in d.items()}
    if isinstance(d, list):
        if len(d) > 8:
            return [_compact(v, depth + 1) for v in d[:6]] + [f"... {len(d) - 6} more"]
        return [_compact(v, depth + 1) for v in d]
    if isinstance(d, float):
        return round(d, 2)
    return d


def _source_color(src: str) -> str:
    if src in SOURCE_COLORS:
        return SOURCE_COLORS[src]
    if src.startswith("model:"):
        return "#8e44ad"
    return EXTRA_COLORS[sum(map(ord, src)) % len(EXTRA_COLORS)]


def _fmt(v: Any) -> str:
    if v is None:
        return "n/a"
    if isinstance(v, float):
        return f"{v:.3f}"
    return str(v)


# --------------------------------------------------------------------------- data collection


def _load_ir(work_dir: Path) -> Document | None:
    p = work_dir / "ir_build.json"
    if not p.is_file():
        return None
    try:
        return Document.from_json(p.read_text(encoding="utf-8"))
    except Exception as exc:
        LOG.warning("cannot read %s: %s", p, exc)
        return None


def _page_overlays(doc: Document | None, index: int) -> tuple[list[dict[str, Any]],
                                                              dict[str, Any]]:
    """Overlay rects (page pt) and the hover JSON by id for one page."""
    if doc is None or index >= len(doc.pages):
        return [], {}
    page = doc.pages[index]
    items: list[Any] = list(page.elements)
    from ..ir import TextBlock

    if not any(isinstance(e, TextBlock) for e in items) and not page.triage.fallback_image:
        from ..build.lines import line_blocks

        items += line_blocks(page, doc.fonts)  # what the builder writes for this page
    items += list(page.annotations)
    rects, data = [], {}
    for e in items[:MAX_OVERLAYS]:
        x0, y0, x1, y1 = e.bbox
        rects.append({"id": e.id, "kind": e.kind, "source": e.source,
                      "bbox": [round(min(x0, x1), 2), round(min(y0, y1), 2),
                               round(max(x0, x1), 2), round(max(y0, y1), 2)],
                      "rot": float(getattr(e, "rotation", 0.0) or 0.0)
                      if e.kind == "text_block" else 0.0})
        data[e.id] = _compact(to_dict(e))
    return rects, data


def low_confidence_blocks(doc: Document | None, index: int,
                          threshold: float | None = None) -> list[dict[str, Any]]:
    """Text blocks of a page whose grouping, role or a paragraph decision is uncertain."""
    from ..ir import TextBlock

    thr = LOW_CONFIDENCE if threshold is None else threshold
    if doc is None or index >= len(doc.pages):
        return []
    out = []
    for b in doc.pages[index].elements:
        if not isinstance(b, TextBlock):
            continue
        weak = [p for p in b.paragraphs if p.confidence < thr]
        if b.confidence >= thr and b.role_confidence >= thr and not weak:
            continue
        text = " / ".join("".join(r.text for r in p.runs) for p in b.paragraphs)
        why = [n for n in b.notes] + [f"{p.id}: {', '.join(p.notes) or 'low'} "
                                      f"({p.confidence:.2f})" for p in weak]
        if b.role_confidence < thr:
            why.append(f"role {b.role} ({b.role_confidence:.2f})")
        out.append({"id": b.id, "confidence": b.confidence, "role": b.role,
                    "role_confidence": b.role_confidence, "text": text[:160], "why": why})
    return sorted(out, key=lambda d: (d["confidence"], d["id"]))


def _ppt_line_rects(layout: list[ppt_com.TextShape], page_wh: tuple[float, float],
                    slide_wh: tuple[float, float]) -> list[dict[str, Any]]:
    s, dx, dy = diff.page_rect_on_slide(*page_wh, *slide_wh)
    out = []
    for k, ln in enumerate(ppt_com.layout_lines(layout)):
        left, top, w, h = ln.bound
        out.append({"id": f"pptline{k}", "text": ln.text.strip()[:120], "shape": ln.shape,
                    "words": len(ln.words),
                    "bbox": [round((left - dx) / s, 2), round((top - dy) / s, 2),
                             round((left + w - dx) / s, 2), round((top + h - dy) / s, 2)]})
    return out


def collect(pdf: Path, cfg: Config, *, pptx: Path | None = None,
            use_cache: bool = True) -> dict[str, Any]:
    """Convert (unless ``pptx`` is given), score, and gather everything the page shows."""
    from .. import pipeline

    t0 = time.perf_counter()
    pdf = Path(pdf)
    sha = pipeline.sha256_file(pdf)
    work_dir = cfg.work_dir / sha
    conv: dict[str, Any] | None = None
    if pptx is None:
        pptx = cfg.out_dir / "report" / f"{pdf.stem}.pptx"
        conv = pipeline.convert(pdf, pptx, cfg, use_cache=use_cache).as_dict()
    doc = _load_ir(work_dir)
    boxes = [diff.page_boxes(p, doc.fonts) for p in doc.pages] if doc else None
    engine, engine_error = score.engine_or_reason(cfg)
    got: dict[int, dict[str, Any]] = {}
    res = score.score_pair(pdf, pptx, cfg, engine=engine, engine_error=engine_error,
                           name=f"report_{pdf.stem}",
                           seconds_per_page=conv["seconds_per_page"] if conv else None,
                           boxes=boxes, collect=got)
    pdf_imgs = render_pages(pdf, cfg.render_dpi)
    pages = []
    for i, row in enumerate(res["slides"]):
        rects, data = _page_overlays(doc, i)
        page_ir = doc.pages[i] if doc and i < len(doc.pages) else None
        wh = ((page_ir.width, page_ir.height) if page_ir else
              (pdf_imgs[i].shape[1] * 72 / cfg.render_dpi, pdf_imgs[i].shape[0] * 72 /
               cfg.render_dpi) if i < len(pdf_imgs) else (0.0, 0.0))
        entry: dict[str, Any] = {"index": i, "size": wh, "row": row, "rects": rects,
                                 "data": data, "ppt_lines": [], "mismatches": []}
        entry["triage"] = to_dict(page_ir.triage) if page_ir else None
        entry["low_conf"] = low_confidence_blocks(doc, i)
        if i < len(pdf_imgs):
            entry["pdf_img"] = _data_uri(pdf_imgs[i])
        c = got.get(i)
        if c is not None:
            entry["pptx_img"] = _data_uri(Image.open(c["png"]))
            d: diff.SlideDiff = c["diff"]
            entry["heat_img"] = _data_uri(diff.heatmap(d, pdf_imgs[i]))
            entry["mismatches"] = [m.to_dict() for m in d.mismatches]
            entry["regions"] = d.region_scores
            slide_wh = tuple(res["slide_size"]) if res.get("slide_size") else wh
            entry["ppt_lines"] = _ppt_line_rects(c["layout"], wh, slide_wh)
        pages.append(entry)
    return {"pdf": str(pdf), "pptx": str(pptx), "sha256": sha, "convert": conv, "score": res,
            "pages": pages, "powerpoint_version": engine.version if engine else None,
            "seconds": round(time.perf_counter() - t0, 2)}


# --------------------------------------------------------------------------- HTML

CSS = """
:root{--bg:#f6f7f9;--fg:#1d2330;--muted:#5d6675;--card:#fff;--line:#d9dde3;--bad:#c0392b;
--ok:#1a7f4b;--warn:#b7791f}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);
font:14px/1.45 system-ui,-apple-system,Segoe UI,Roboto,sans-serif}
header{padding:16px 20px;background:var(--card);border-bottom:1px solid var(--line);
position:sticky;top:0;z-index:5}
h1{font-size:18px;margin:0 0 4px}h2{font-size:16px;margin:0}
.muted{color:var(--muted)}.mono{font-family:Consolas,ui-monospace,monospace;font-size:12px}
main{padding:16px 20px;max-width:1800px;margin:0 auto}
.banner{padding:10px 14px;border-radius:6px;margin:10px 0;background:#fff4e5;
border:1px solid #f0c27a}
.controls{display:flex;flex-wrap:wrap;gap:6px 18px;align-items:center;margin-top:8px}
.controls fieldset{border:1px solid var(--line);border-radius:6px;padding:2px 8px;margin:0}
.controls legend{font-size:11px;color:var(--muted);padding:0 4px}
.controls label{margin-right:8px;white-space:nowrap;cursor:pointer}
.sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:3px;
vertical-align:-1px}
table.scores{border-collapse:collapse;background:var(--card);width:100%;font-size:13px}
table.scores th,table.scores td{border:1px solid var(--line);padding:4px 6px;text-align:right}
table.scores th{background:#eef1f5;cursor:pointer;user-select:none;position:sticky;top:0}
table.scores th.sorted-asc::after{content:" \\25B2"}table.scores th.sorted-desc::after{
content:" \\25BC"}
table.scores td:first-child,table.scores th:first-child{text-align:left}
.page{background:var(--card);border:1px solid var(--line);border-radius:8px;margin:18px 0;
padding:12px 14px}
.page-head{display:flex;flex-wrap:wrap;gap:8px 16px;align-items:baseline;margin-bottom:8px}
.chip{display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;
background:#eef1f5}
.chip.born_digital{background:#e3f4ea}.chip.scanned{background:#fde8e6}
.chip.outlined_text,.chip.broken_text{background:#fff4e5}
.panels{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}
.panel{margin:0}.panel figcaption{font-size:12px;color:var(--muted);margin-bottom:3px}
.stage{position:relative;border:1px solid var(--line);background:#fff;line-height:0}
.stage img{width:100%;height:auto;display:block}
.stage svg{position:absolute;inset:0;width:100%;height:100%;overflow:visible}
.stage .na{line-height:1.4;padding:30px 12px;color:var(--muted);text-align:center}
svg rect{fill:transparent;stroke-width:1.2;vector-effect:non-scaling-stroke}
svg rect:hover{fill:rgba(255,220,0,.25);stroke-width:2.2}
svg rect.mm{stroke:var(--bad);stroke-dasharray:4 2;stroke-width:1.6}
svg rect.pl{stroke:#00a0c8;stroke-dasharray:2 2}
.row td.bad{color:var(--bad);font-weight:600}
#tip{position:fixed;z-index:20;max-width:560px;max-height:70vh;overflow:auto;
background:#111a;backdrop-filter:blur(2px);background:rgba(20,24,32,.95);color:#e8edf5;
padding:8px 10px;border-radius:6px;font:12px/1.35 Consolas,ui-monospace,monospace;
white-space:pre-wrap;pointer-events:none;display:none}
@media (max-width:900px){.panels{grid-template-columns:1fr}}
"""

JS = r"""
const KIND_COLORS = __KIND_COLORS__;
const SRC_COLORS = __SRC_COLORS__;
const DATA = JSON.parse(document.getElementById('ir-data').textContent);
const tip = document.getElementById('tip');
function colorFor(el){
  const by = document.querySelector('input[name=colorby]:checked').value;
  return by === 'kind' ? (KIND_COLORS[el.dataset.kind] || '#888')
                       : (SRC_COLORS[el.dataset.source] || '#888');
}
function refresh(){
  const kinds = new Set([...document.querySelectorAll('input.k:checked')].map(i => i.value));
  const srcs = new Set([...document.querySelectorAll('input.s:checked')].map(i => i.value));
  const layers = new Set([...document.querySelectorAll('input.l:checked')].map(i => i.value));
  document.querySelectorAll('svg rect.el').forEach(el => {
    const on = layers.has('ir') && kinds.has(el.dataset.kind) && srcs.has(el.dataset.source);
    el.style.display = on ? '' : 'none';
    el.style.stroke = colorFor(el);
  });
  document.querySelectorAll('svg rect.pl').forEach(el =>
    el.style.display = layers.has('ppt') ? '' : 'none');
  document.querySelectorAll('svg rect.mm').forEach(el =>
    el.style.display = layers.has('mm') ? '' : 'none');
}
document.querySelectorAll('.controls input').forEach(i => i.addEventListener('change', refresh));
document.addEventListener('mouseover', ev => {
  const r = ev.target.closest && ev.target.closest('svg rect');
  if (!r) return;
  let d;
  if (r.classList.contains('el')) d = DATA[r.dataset.page][r.dataset.id];
  else d = JSON.parse(r.dataset.info);
  tip.textContent = JSON.stringify(d, null, 1);
  tip.style.display = 'block';
});
document.addEventListener('mousemove', ev => {
  if (tip.style.display !== 'block') return;
  const w = tip.offsetWidth, h = tip.offsetHeight;
  let x = ev.clientX + 14, y = ev.clientY + 14;
  if (x + w > innerWidth - 8) x = ev.clientX - w - 14;
  if (y + h > innerHeight - 8) y = Math.max(8, innerHeight - h - 8);
  tip.style.left = x + 'px'; tip.style.top = y + 'px';
});
document.addEventListener('mouseout', ev => {
  if (ev.target.closest && ev.target.closest('svg rect')) tip.style.display = 'none';
});
refresh();
"""

SORT_JS = r"""
document.querySelectorAll('table.scores').forEach(tbl => {
  tbl.querySelectorAll('th').forEach((th, col) => th.addEventListener('click', () => {
    const asc = !th.classList.contains('sorted-asc');
    tbl.querySelectorAll('th').forEach(h => h.classList.remove('sorted-asc', 'sorted-desc'));
    th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');
    const body = tbl.tBodies[0];
    const rows = [...body.rows];
    const key = r => { const c = r.cells[col]; const v = c.dataset.v ?? c.textContent;
                       const n = parseFloat(v); return isNaN(n) ? v : n; };
    rows.sort((a, b) => { const x = key(a), y = key(b);
      const nx = typeof x === 'number', ny = typeof y === 'number';
      if (nx !== ny) return nx ? -1 : 1;   // n/a sorts last
      return (x < y ? -1 : x > y ? 1 : 0) * (asc ? 1 : -1); });
    rows.forEach(r => body.appendChild(r));
  }));
});
"""

METRICS = score.METRIC_KEYS
SHORT = {"ssim": "ssim", "region_ssim_min": "reg_min", "text_sim": "text_sim",
         "word_recall": "w_recall", "pos_err_pt": "pos_err", "linebreak_match": "lb_match",
         "real_text_ratio": "real_txt", "native_shape_ratio": "native",
         "boxes_per_slide": "boxes", "seconds_per_page": "s/page", "composite": "composite"}
BAD = {"ssim": 0.95, "region_ssim_min": 0.85, "text_sim": 0.98, "word_recall": 0.98,
       "linebreak_match": 0.98, "real_text_ratio": 0.98, "composite": 0.85}


def _esc(s: Any) -> str:
    return html.escape(str(s), quote=True)


def _score_cells(row: dict[str, Any]) -> str:
    cells = []
    for k in METRICS:
        v = row.get(k)
        bad = k in BAD and isinstance(v, (int, float)) and v < BAD[k]
        bad = bad or (k == "pos_err_pt" and isinstance(v, (int, float)) and v > 1.5)
        dv = "" if v is None else f' data-v="{v}"'
        cells.append(f'<td{dv} class="{"bad" if bad else ""}">{_fmt(v)}</td>')
    return "".join(cells)


def _svg(page: dict[str, Any], *, ir: bool = True, ppt: bool = True, mm: bool = True) -> str:
    w, h = page["size"]
    parts = [f'<svg viewBox="0 0 {w:.2f} {h:.2f}" preserveAspectRatio="none">']
    pi = page["index"]
    if ir:
        for r in page["rects"]:
            x0, y0, x1, y1 = r["bbox"]
            tr = ""
            if r["rot"]:
                tr = f' transform="rotate({r["rot"]:.2f} {(x0 + x1) / 2:.2f} {(y0 + y1) / 2:.2f})"'
            parts.append(
                f'<rect class="el" data-page="{pi}" data-id="{_esc(r["id"])}" '
                f'data-kind="{_esc(r["kind"])}" data-source="{_esc(r["source"])}" '
                f'x="{x0}" y="{y0}" width="{max(x1 - x0, 0.3):.2f}" '
                f'height="{max(y1 - y0, 0.3):.2f}"{tr}/>')
    if ppt:
        for r in page["ppt_lines"]:
            x0, y0, x1, y1 = r["bbox"]
            parts.append(f'<rect class="pl" data-info="{_esc(json.dumps(r))}" x="{x0}" '
                         f'y="{y0}" width="{max(x1 - x0, 0.3):.2f}" '
                         f'height="{max(y1 - y0, 0.3):.2f}"/>')
    if mm:
        for k, m in enumerate(page["mismatches"]):
            x0, y0, x1, y1 = m["bbox"]
            info = {"mismatch": k + 1, **m}
            parts.append(f'<rect class="mm" data-info="{_esc(json.dumps(info))}" x="{x0}" '
                         f'y="{y0}" width="{x1 - x0:.2f}" height="{y1 - y0:.2f}"/>')
    parts.append("</svg>")
    return "".join(parts)


def _panel(title: str, img: str | None, svg: str, missing: str) -> str:
    inner = (f'<img src="{img}" alt="{_esc(title)}">{svg}' if img
             else f'<div class="na">{_esc(missing)}</div>')
    return (f'<figure class="panel"><figcaption>{_esc(title)}</figcaption>'
            f'<div class="stage">{inner}</div></figure>')


def render_html(rep: dict[str, Any]) -> str:
    res = rep["score"]
    pages = rep["pages"]
    title = Path(rep["pdf"]).name
    kinds = sorted({r["kind"] for p in pages for r in p["rects"]})
    sources = sorted({r["source"] for p in pages for r in p["rects"]})
    ppt_err = res.get("powerpoint_error")
    out = ["<!doctype html><html lang=en><head><meta charset=utf-8>",
           '<meta name=viewport content="width=device-width,initial-scale=1">',
           f"<title>pdf2pptx report: {_esc(title)}</title><style>{CSS}</style></head><body>",
           f"<header><h1>{_esc(title)}</h1>",
           f'<div class="muted mono">{_esc(rep["pdf"])} → {_esc(rep["pptx"])}</div>']
    out.append('<div class="controls"><fieldset><legend>layers</legend>'
               '<label><input type=checkbox class=l value=ir checked> IR boxes</label>'
               '<label><input type=checkbox class=l value=ppt> PowerPoint lines</label>'
               '<label><input type=checkbox class=l value=mm checked> mismatch regions</label>'
               '</fieldset><fieldset><legend>colour by</legend>'
               '<label><input type=radio name=colorby value=kind checked> kind</label>'
               '<label><input type=radio name=colorby value=source> source</label></fieldset>')
    out.append('<fieldset><legend>kind</legend>' + "".join(
        f'<label><input type=checkbox class=k value="{_esc(k)}" checked>'
        f'<span class=sw style="background:{KIND_COLORS.get(k, "#888")}"></span>{_esc(k)}</label>'
        for k in kinds) + "</fieldset>")
    out.append('<fieldset><legend>source</legend>' + "".join(
        f'<label><input type=checkbox class=s value="{_esc(s)}" checked>'
        f'<span class=sw style="background:{_source_color(s)}"></span>{_esc(s)}</label>'
        for s in sources) + "</fieldset></div></header><main>")

    if ppt_err:
        out.append(f'<div class="banner"><b>PowerPoint unavailable:</b> {_esc(ppt_err)}<br>'
                   "The PowerPoint render, the diff heatmap and the PowerPoint-based scores "
                   "(ssim, region_ssim_min, pos_err_pt, linebreak_match) are not available; "
                   "word_recall falls back to order-free token recall.</div>")
    conv = rep.get("convert") or {}
    facts = [f"{len(pages)} pages",
             f"PowerPoint {rep['powerpoint_version']}" if rep["powerpoint_version"] else
             "PowerPoint: n/a",
             f"convert {conv.get('seconds', 'n/a')} s" if conv else "PPTX given",
             f"score {res.get('score_seconds')} s",
             f"triage {conv.get('triage')}" if conv else "",
             f"fallback pages {conv.get('fallback_pages') or 'none'}" if conv else ""]
    cs = res.get("com_seconds_per_slide") or []
    if cs:
        facts.append(f"render+layout median {float(np.median(cs)):.3f} s/slide")
    out.append('<p class="muted">' + " · ".join(_esc(f) for f in facts if f) + "</p>")

    # index: sortable score table
    out.append('<h2>Pages</h2><p class="muted">Click a column to sort; '
               'click a page to jump to it.</p><table class="scores"><thead><tr><th>page</th>'
               "<th>triage</th>" + "".join(f"<th>{SHORT[k]}</th>" for k in METRICS)
               + "<th>mismatches</th></tr></thead><tbody>")
    for p in pages:
        tri = (p["triage"] or {}).get("kind", "n/a")
        out.append(f'<tr class="row"><td data-v="{p["index"] + 1}"><a href="#page-'
                   f'{p["index"] + 1}">page {p["index"] + 1}</a></td><td>{_esc(tri)}</td>'
                   + _score_cells(p["row"]) + f'<td>{len(p["mismatches"])}</td></tr>')
    summ = res.get("summary", {})
    out.append("</tbody><tfoot><tr><td>median</td><td></td>" + "".join(
        f"<td>{_fmt(summ.get(k, {}).get('median'))}</td>" for k in METRICS)
        + "<td></td></tr></tfoot></table>")

    for p in pages:
        n = p["index"] + 1
        tri = p["triage"] or {}
        chip = (f'<span class="chip {_esc(tri.get("kind", ""))}">{_esc(tri.get("kind", "n/a"))}'
                f'{" · fallback image" if tri.get("fallback_image") else ""}</span>')
        reason = (f'<span class="muted">{_esc(tri.get("reason", ""))} '
                  f'(confidence {_fmt(tri.get("confidence"))})</span>' if tri else "")
        out.append(f'<section class="page" id="page-{n}"><div class="page-head">'
                   f"<h2>Page {n}</h2>{chip}{reason}"
                   f'<span class="muted">{p["size"][0]:.1f} × {p["size"][1]:.1f} pt · '
                   f'{len(p["rects"])} IR boxes · {len(p["ppt_lines"])} PowerPoint lines'
                   "</span></div>")
        out.append('<table class="scores"><thead><tr>' + "".join(
            f"<th>{SHORT[k]}</th>" for k in METRICS) + "</tr></thead><tbody><tr class=row>"
            + _score_cells(p["row"]) + "</tr></tbody></table>")
        why = ppt_err or "not rendered"
        out.append('<div class="panels" style="margin-top:10px">')
        out.append(_panel("PDF (pdfium)", p.get("pdf_img"), _svg(p, ppt=False), "no render"))
        out.append(_panel("PPTX (PowerPoint)", p.get("pptx_img"), _svg(p),
                          f"PowerPoint render unavailable: {why}"))
        out.append(_panel("diff heatmap (1 − local SSIM; grid = 4×4 regions)",
                          p.get("heat_img"), _svg(p, ppt=False),
                          f"no heatmap: {why}"))
        out.append("</div>")
        if p.get("low_conf"):
            items = "".join(
                f"<li><b>{_esc(b['id'])}</b> confidence {b['confidence']:.2f}, role "
                f"{_esc(b['role'])} ({b['role_confidence']:.2f}): “{_esc(b['text'])}” "
                f"<span class=muted>{_esc('; '.join(b['why']))}</span></li>"
                for b in p["low_conf"])
            out.append(f'<details open><summary class="muted">{len(p["low_conf"])} '
                       f'low-confidence text blocks (&lt; {LOW_CONF_SHOWN})</summary>'
                       f'<ul class="mono">{items}</ul></details>')
        if p["mismatches"]:
            items = "".join(
                f"<li>#{k + 1} ssim {m['ssim']:.3f}, bbox {m['bbox']}: "
                f"{_esc(', '.join(m['element_ids'][:8]) or 'no IR element')}</li>"
                for k, m in enumerate(p["mismatches"][:8]))
            out.append(f'<details><summary class="muted">{len(p["mismatches"])} mismatch '
                       f'regions</summary><ul class="mono">{items}</ul></details>')
        out.append("</section>")

    data = {str(p["index"]): p["data"] for p in pages}
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    out.append(f'</main><div id="tip"></div><script type="application/json" id="ir-data">'
               f"{payload}</script>")
    js = JS.replace("__KIND_COLORS__", json.dumps(KIND_COLORS)).replace(
        "__SRC_COLORS__", json.dumps({s: _source_color(s) for s in sources}))
    out.append(f"<script>{js}{SORT_JS}</script></body></html>")
    return "".join(out)


def write_report(pdf: Path, cfg: Config, out_html: Path | None = None, *,
                 pptx: Path | None = None, use_cache: bool = True) -> tuple[Path, dict[str, Any]]:
    rep = collect(pdf, cfg, pptx=pptx, use_cache=use_cache)
    out_html = out_html or cfg.out_dir / "report" / f"{Path(pdf).stem}.html"
    out_html.parent.mkdir(parents=True, exist_ok=True)
    out_html.write_text(render_html(rep), encoding="utf-8", newline="\n")
    LOG.info("report: %s (%.1f s)", out_html, rep["seconds"])
    return out_html, rep


def render_index(entries: list[tuple[Path, dict[str, Any]]], out_html: Path) -> str:
    """Index of several reports: one sortable row per document (medians)."""
    rows = []
    for path, rep in entries:
        s = rep["score"]["summary"]
        rel = Path(path).name
        cells = "".join(
            f'<td data-v="{s[k]["median"]}">{_fmt(s[k]["median"])}</td>'
            if s[k]["median"] is not None else "<td>n/a</td>" for k in METRICS)
        rows.append(f'<tr><td><a href="{_esc(rel)}">{_esc(Path(rep["pdf"]).name)}</a></td>'
                    f'<td data-v="{len(rep["pages"])}">{len(rep["pages"])}</td>{cells}</tr>')
    head = "".join(f"<th>{SHORT[k]}</th>" for k in METRICS)
    return ("<!doctype html><html lang=en><head><meta charset=utf-8>"
            '<meta name=viewport content="width=device-width,initial-scale=1">'
            f"<title>pdf2pptx reports</title><style>{CSS}</style></head><body><header>"
            "<h1>pdf2pptx debug reports</h1><div class=muted>median over pages; click a "
            "column to sort</div></header><main><table class=scores><thead><tr><th>file</th>"
            f"<th>pages</th>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table></main>"
            f"<script>{SORT_JS}</script></body></html>")
