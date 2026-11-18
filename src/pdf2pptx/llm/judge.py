"""Judgment pass (Max tier, M9): a vision model looks at the regions the rules were
unsure about and its schema-constrained answers are merged into the IR.

Targets (``scope="low"``, the default; ``"all"`` takes every candidate):

* **roles**: text blocks whose rule role confidence is below ``ROLE_TRIGGER`` (the M8
  layout trigger) or where the layout model disagreed (``pdf2pptx:layout``).
* **reading order**: pages whose body text sits in two or more columns (the XY-cut
  order is a guess there).
* **tables**: tables below ``TABLE_TRIGGER`` confidence (header rows; the grid itself is
  only logged).
* **regions**: groups of three or more shapes and large pictures are classified
  (timeline / process / matrix / org chart / table / chart / ...): recorded as notes
  for later milestones.
* **charts**: every geometry chart is checked by the model; every large picture is
  offered as a possible chart (a native chart is built when the model reads one).

Each answer is validated (``schemas.valid_*``); invalid output keeps the rule result.
Changes carry ``source="model:<name>"`` and the model's confidence and are adopted
only at ``MIN_ADOPT`` or above. Whole pages are then kept only where they score at
least as well as the Lite result (``keep_best`` in ``llm/max_tier.py``).
"""

from __future__ import annotations

import json
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import cv2
import numpy as np

from .. import log
from ..analyze.roles import block_text
from ..ir import Chart, Document, Group, Image, Page, Table, TextBlock
from . import charts_vlm, schemas
from .scheduler import Job, Scheduler

LOG = log.get(__name__)

ROLE_TRIGGER = 0.65
TABLE_TRIGGER = 0.8
MIN_ADOPT = 0.75
MAX_ROLE_BLOCKS = 12  # per page and call
RENDER_DPI = 150


@dataclass
class Proposal:
    kind: str  # roles | order | table | region | chart_check | chart_new
    page: int
    target: str  # element id (or page)
    data: dict[str, Any] | None = None
    applied: bool = False
    note: str = ""


@dataclass
class JudgeReport:
    targets: dict[str, int] = field(default_factory=dict)
    proposals: list[Proposal] = field(default_factory=list)
    invalid: int = 0
    not_chart: int = 0

    def as_dict(self) -> dict[str, Any]:
        applied: dict[str, int] = {}
        for p in self.proposals:
            if p.applied:
                applied[p.kind] = applied.get(p.kind, 0) + 1
        return {"targets": dict(sorted(self.targets.items())), "applied": applied,
                "invalid": self.invalid, "not_chart": self.not_chart,
                "changes": [{"kind": p.kind, "page": p.page, "target": p.target,
                             "note": p.note} for p in self.proposals if p.applied]}


def _page_image(pdf: Path, work_dir: Path, index: int) -> np.ndarray:
    from ..ingest.render import ensure_render

    img = cv2.imread(str(ensure_render(pdf, work_dir, index, RENDER_DPI)), cv2.IMREAD_COLOR)
    if img is None:
        raise RuntimeError(f"page {index}: render missing")
    return img


class _LazyImage:
    """The page render, read on first use."""

    def __init__(self, pdf: Path, work_dir: Path, index: int) -> None:
        self.args = (pdf, work_dir, index)
        self.img: np.ndarray | None = None

    def __call__(self) -> np.ndarray:
        if self.img is None:
            self.img = _page_image(*self.args)
        return self.img


def _whole_page(img: np.ndarray) -> bytes:
    return charts_vlm.crop_png(img, (0, 0, img.shape[1] * 72 / RENDER_DPI,
                                     img.shape[0] * 72 / RENDER_DPI), RENDER_DPI / 72, pad=0)


def _layout_disagreements(doc: Document) -> set[str]:
    try:
        lay = json.loads(doc.metadata.get("pdf2pptx:layout") or "{}")
    except ValueError:
        return set()
    return {str(d.get("block")) for d in lay.get("disagreements", []) if d.get("block")}


def _columns(blocks: list[TextBlock]) -> bool:
    body = [b for b in blocks if b.role == "body" and len(block_text(b)) > 40]
    for a in body:
        for b in body:
            if a is b:
                continue
            xo = min(a.bbox[2], b.bbox[2]) - max(a.bbox[0], b.bbox[0])
            yo = min(a.bbox[3], b.bbox[3]) - max(a.bbox[1], b.bbox[1])
            if xo < 0 and yo > 0:
                return True
    return False


def _numbered(img: np.ndarray, boxes: list[tuple[float, float, float, float]]) -> bytes:
    out = img.copy()
    s = RENDER_DPI / 72
    for i, b in enumerate(boxes):
        p0, p1 = (int(b[0] * s), int(b[1] * s)), (int(b[2] * s), int(b[3] * s))
        cv2.rectangle(out, p0, p1, (0, 0, 230), 2)
        cv2.putText(out, str(i), (p0[0] + 2, p0[1] + 22), cv2.FONT_HERSHEY_SIMPLEX, 0.8,
                    (0, 0, 230), 2)
    return _whole_page(out)


def plan(doc: Document, pdf: Path, work_dir: Path, sched: Scheduler, model: str,
         scope: str = "low", chart_model: str | None = None) -> JudgeReport:
    """Queue the judgment jobs of ``doc`` on ``sched`` (run them with ``sched.run()``)."""
    rep = JudgeReport()
    chart_model = chart_model or model
    disagree = _layout_disagreements(doc)
    for page in doc.pages:
        if page.triage.fallback_image:
            continue
        page_img = _LazyImage(pdf, work_dir, page.index)

        blocks = [e for e in page.elements if isinstance(e, TextBlock)]
        # roles
        weak = [b for b in blocks if scope == "all" or b.role_confidence < ROLE_TRIGGER
                or b.id in disagree][:MAX_ROLE_BLOCKS]
        if weak:
            rep.targets["roles"] = rep.targets.get("roles", 0) + len(weak)
            prop = Proposal("roles", page.index, ",".join(b.id for b in weak))
            rep.proposals.append(prop)
            sched.submit(Job(model, schemas.role_prompt(
                [(b.id, block_text(b), b.role) for b in weak]), [_whole_page(page_img())],
                schemas.ROLE_FIX, schemas.SYSTEM, 600, f"roles:p{page.index}",
                _setter(prop)))
        # reading order
        ordered = sorted((b for b in blocks if b.reading_order is not None),
                         key=lambda b: b.reading_order or 0)
        if len(ordered) >= 3 and (scope == "all" or _columns(ordered)) and len(ordered) <= 30:
            rep.targets["order"] = rep.targets.get("order", 0) + 1
            prop = Proposal("order", page.index, f"p{page.index}")
            prop.data = {"ids": [b.id for b in ordered]}
            rep.proposals.append(prop)
            sched.submit(Job(model, schemas.order_prompt(
                [(i, block_text(b)) for i, b in enumerate(ordered)]),
                [_numbered(page_img(), [b.bbox for b in ordered])], schemas.READING_ORDER,
                schemas.SYSTEM, 300, f"order:p{page.index}", _setter(prop, keep=True)))
        # tables
        for t in (e for e in page.elements if isinstance(e, Table)):
            if scope != "all" and t.confidence >= TABLE_TRIGGER:
                continue
            rep.targets["tables"] = rep.targets.get("tables", 0) + 1
            prop = Proposal("table", page.index, t.id)
            rep.proposals.append(prop)
            sched.submit(Job(model, schemas.table_prompt(_table_rows(t)),
                             [charts_vlm.crop_png(page_img(), t.bbox, RENDER_DPI / 72)],
                             schemas.TABLE_FIX, schemas.SYSTEM, 1500, f"table:{t.id}",
                             _setter(prop)))
        # regions: multi-part graphics
        for g in (e for e in page.elements if isinstance(e, Group)):
            if len(g.child_ids) < 3 or (scope != "all" and g.confidence >= 0.95):
                continue
            rep.targets["regions"] = rep.targets.get("regions", 0) + 1
            prop = Proposal("region", page.index, g.id)
            rep.proposals.append(prop)
            sched.submit(Job(model, schemas.region_prompt("vector graphic", len(g.child_ids)),
                             [charts_vlm.crop_png(page_img(), g.bbox, RENDER_DPI / 72)],
                             schemas.REGION, schemas.SYSTEM, 100, f"region:{g.id}",
                             _setter(prop)))
        # charts: check the geometry charts, offer large pictures
        for ch in (e for e in page.elements if isinstance(e, Chart)):
            rep.targets["chart_checks"] = rep.targets.get("chart_checks", 0) + 1
            prop = Proposal("chart_check", page.index, ch.id)
            rep.proposals.append(prop)
            sched.submit(Job(chart_model, schemas.chart_prompt(),
                             [charts_vlm.crop_png(page_img(), ch.bbox, RENDER_DPI / 72)],
                             schemas.CHART, schemas.SYSTEM, 1200, f"chart:{ch.id}",
                             _setter(prop)))
        for im in (e for e in page.elements if isinstance(e, Image)):
            w, h = im.bbox[2] - im.bbox[0], im.bbox[3] - im.bbox[1]
            if im.is_background or w < charts_vlm.MIN_IMAGE_PT or h < charts_vlm.MIN_IMAGE_PT:
                continue
            if page.triage.kind == "scanned" and not any("chart" in n for n in im.notes):
                continue
            rep.targets["pictures"] = rep.targets.get("pictures", 0) + 1
            prop = Proposal("chart_new", page.index, im.id)
            crop = charts_vlm.crop_png(page_img(), im.bbox, RENDER_DPI / 72, pad=0)
            prop.data = {"crop": crop}
            rep.proposals.append(prop)
            sched.submit(Job(chart_model, schemas.chart_prompt(
                "If the picture is not a chart, use chart_type 'other'."), [crop],
                schemas.CHART, schemas.SYSTEM, 1200, f"picture:{im.id}",
                _setter(prop, keep=True)))
    return rep


def _setter(prop: Proposal, keep: bool = False) -> Callable[[dict[str, Any]], None]:
    def done(res: dict[str, Any]) -> None:
        out = {"json": res.get("json"), "error": res.get("error"),
               "model": res.get("model")}
        prop.data = (prop.data or {}) | out if keep else out
    return done


def _table_rows(t: Table) -> list[list[str]]:
    rows = [["" for _ in t.col_widths] for _ in t.row_heights]
    for c in t.cells:
        if c.text is not None and c.row < len(rows) and c.col < len(rows[0]):
            rows[c.row][c.col] = block_text(c.text)
    return rows


def _model_name(p: Proposal) -> str:
    m = str((p.data or {}).get("model") or "model")
    return m.split("@")[0]


def apply(doc: Document, rep: JudgeReport) -> JudgeReport:
    """Merge the validated answers into ``doc`` (in place)."""
    pages = {p.index: p for p in doc.pages}
    for p in rep.proposals:
        d = p.data or {}
        if d.get("error") or d.get("json") is None:
            if d.get("json") is None and not d.get("error"):
                rep.invalid += 1
            continue
        page = pages[p.page]
        name = _model_name(p)
        src = f"model:{name}"
        by_id = {e.id: e for e in page.elements}
        if p.kind == "roles":
            ids = set(p.target.split(","))
            v = schemas.valid_roles(d["json"], ids)
            if v is None:
                rep.invalid += 1
                continue
            changed = []
            for bid, (role, conf) in sorted(v.items()):
                b = by_id.get(bid)
                if not isinstance(b, TextBlock) or role == "other" or role == b.role:
                    continue
                if conf < MIN_ADOPT or b.placeholder:
                    continue
                b.notes.append(f"role:{b.role}->{role}:{src}:{conf:.2f}")
                b.role, b.role_confidence, b.source = role, round(conf, 3), src
                changed.append(f"{bid}={role}")
            if changed:
                p.applied, p.note = True, " ".join(changed)
        elif p.kind == "order":
            ids = list(d.get("ids") or [])
            v2 = schemas.valid_order(d["json"], len(ids))
            if v2 is None:
                rep.invalid += 1
                continue
            order, conf = v2
            if conf < MIN_ADOPT or order == list(range(len(ids))):
                continue
            for rank, i in enumerate(order):
                b = by_id.get(ids[i])
                if isinstance(b, TextBlock):
                    b.reading_order = rank
                    b.notes.append(f"reading_order:{src}:{conf:.2f}")
            # the writer emits text in z order: follow the new reading order
            texts = sorted((e for e in page.elements if isinstance(e, TextBlock)
                            and e.id in ids), key=lambda e: e.reading_order or 0)
            zs = sorted(e.z for e in texts)
            for e, z in zip(texts, zs, strict=True):
                e.z = z
            p.applied, p.note = True, f"order {order}"
        elif p.kind == "table":
            t = by_id.get(p.target)
            v3 = schemas.valid_table(d["json"])
            if v3 is None or not isinstance(t, Table):
                rep.invalid += 1
                continue
            same_grid = v3["rows"] == len(t.row_heights) and v3["cols"] == len(t.col_widths)
            t.notes.append(f"vlm_table:{src}:{'same' if same_grid else 'different'}_grid:"
                           f"{v3['rows']}x{v3['cols']}")
            if same_grid and v3["header_rows"] != t.header_rows and v3["confidence"] >= MIN_ADOPT:
                t.notes.append(f"header_rows:{t.header_rows}->{v3['header_rows']}:{src}")
                t.header_rows = int(v3["header_rows"])
                t.source = src
                p.applied, p.note = True, f"header_rows={t.header_rows}"
        elif p.kind == "region":
            g = by_id.get(p.target)
            v4 = schemas.valid_region(d["json"])
            if v4 is None or g is None:
                rep.invalid += 1
                continue
            g.notes.append(f"region:{v4['label']}:{src}:{float(v4['confidence']):.2f}")
            p.applied, p.note = True, str(v4["label"])
        elif p.kind == "chart_check":
            ch = by_id.get(p.target)
            if isinstance(ch, Chart):
                res = charts_vlm.apply_check(ch, d["json"], name)
                if res is None:
                    rep.invalid += 1
                else:
                    p.applied, p.note = True, json.dumps(res, sort_keys=True)
        elif p.kind == "chart_new":
            im = by_id.get(p.target)
            if not isinstance(im, Image):
                continue
            crop = charts_vlm.decode_png(d.get("crop") or b"")
            n = sum(1 for e in page.elements if isinstance(e, Chart))
            ch2 = charts_vlm.from_picture(im, crop, d["json"], name,
                                          f"p{page.index}_vc{n + 1}")
            if ch2 is None:
                js = d["json"] if isinstance(d["json"], dict) else {}
                if js.get("chart_type") == "other":
                    rep.not_chart += 1  # the model says it is no chart: fine
                elif schemas.valid_chart(js) is None:
                    rep.invalid += 1
                continue
            page.elements = [e for e in page.elements if e is not im] + [ch2]
            p.applied, p.note = True, f"{ch2.chart_type} {len(ch2.categories)}x{len(ch2.series)}"
    return rep


def affected_pages(rep: JudgeReport) -> list[int]:
    return sorted({p.page for p in rep.proposals if p.applied})


def drop_page_changes(doc: Document, lite: Document, pages: list[int]) -> None:
    """Keep-best: put the Lite elements back on ``pages``."""
    for i in pages:
        doc.pages[i].elements = lite.pages[i].elements


def touched(page: Page) -> bool:
    return any(getattr(e, "source", "").startswith("model:") for e in page.elements)
