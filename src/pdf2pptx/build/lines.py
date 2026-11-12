"""Baseline text layout: glyphs -> one single-line, non-wrapping text block per line.

Used by the writer for pages whose IR has no TextBlocks yet (until M4's analyzer
produces real paragraphs). A "line" is a run of glyphs in content order that share
rotation and baseline and have no large gap; a gap wider than ``SPLIT_GAP`` em
starts a new box so that columns and table cells keep their exact positions.
Inside a line, every space gets character spacing so that the next word starts at
its PDF origin (advances from the installed font; ``fonts/index.py``).

Geometry: the block's ``bbox`` is the unrotated frame (PowerPoint ``a:xfrm``) and
``rotation`` is clockwise degrees. The frame top is placed so that PowerPoint's
first baseline (see ``fonts/index.py``) lands on the PDF baseline.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from dataclasses import dataclass, field

from ..fonts.index import LINE_FACTOR, advance_em, first_baseline_offset, resolve_family
from ..ingest import geom
from ..ir import Annotation, FontRef, Glyph, Page, Paragraph, TextBlock, TextRun

SPLIT_GAP = 0.8  # em: a gap wider than this starts a new text box
SPACE_GAP = 0.15  # em: a gap wider than this (without a space glyph) inserts a space
SAME_BASELINE = 0.25  # em: max baseline offset across the line direction
BACKSTEP = 0.3  # em: how far a glyph may start before the previous glyph's end


@dataclass
class _Line:
    first: Glyph
    rot: float
    u: tuple[float, float]
    n: tuple[float, float]
    end: float = 0.0
    pen: float = 0.0  # where PowerPoint's pen will be after the last glyph (along)
    # (style glyph, text, char spacing in pt for inserted spaces)
    glyphs: list[tuple[Glyph, str, float | None]] = field(default_factory=list)
    pending_space: bool = False

    def along_across(self, g: Glyph) -> tuple[float, float]:
        dx, dy = g.origin_x - self.first.origin_x, g.baseline - self.first.baseline
        return dx * self.u[0] + dy * self.u[1], dx * self.n[0] + dy * self.n[1]

    def extent(self, g: Glyph) -> float:
        """Far edge of ``g`` along the line direction (relative to the line origin)."""
        x0, y0, x1, y1 = g.bbox
        best = -math.inf
        for x, y in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
            dx, dy = x - self.first.origin_x, y - self.first.baseline
            best = max(best, dx * self.u[0] + dy * self.u[1])
        along, _ = self.along_across(g)
        return max(best, along + 0.25 * g.size) if x1 - x0 <= 0 else best


def _dirs(rot_ccw: float) -> tuple[tuple[float, float], tuple[float, float]]:
    a = math.radians(rot_ccw)
    return (math.cos(a), -math.sin(a)), (math.sin(a), math.cos(a))


Advance = Callable[[Glyph, str], float | None]  # em, or None if unknown


def _lines(glyphs: list[Glyph], adv: Advance) -> list[_Line]:
    lines: list[_Line] = []
    cur: _Line | None = None

    def pen_after(line: _Line, g: Glyph, along: float) -> float:
        a = adv(g, g.text)
        return along + a * g.size if a is not None else line.extent(g)

    for g in glyphs:
        ws = g.text.isspace()
        if cur is not None:
            same_rot = abs(geom.norm_angle(g.rotation - cur.rot)) < 0.5
            along, across = cur.along_across(g)
            em = max(g.size, cur.first.size, 1.0)
            if ws:
                if same_rot and abs(across) <= SAME_BASELINE * em:
                    cur.pending_space = True
                continue
            gap = along - cur.end
            if (same_rot and abs(across) <= SAME_BASELINE * em
                    and gap >= -BACKSTEP * em and gap <= SPLIT_GAP * em):
                if (cur.pending_space or gap > SPACE_GAP * g.size) and cur.glyphs:
                    prev = cur.glyphs[-1][0]
                    sp = adv(prev, " ")
                    sp_pt = (sp if sp is not None else 0.25) * prev.size
                    cs = along - cur.pen - sp_pt
                    cur.glyphs.append((prev, " ", round(cs, 2) if abs(cs) > 0.25 else None))
                cur.pending_space = False
                cur.glyphs.append((g, g.text, None))
                cur.end = max(cur.end, cur.extent(g))
                cur.pen = pen_after(cur, g, along)
                continue
        if ws:
            continue
        u, n = _dirs(g.rotation)
        cur = _Line(first=g, rot=g.rotation, u=u, n=n)
        cur.glyphs.append((g, g.text, None))
        cur.end = cur.extent(g)
        cur.pen = pen_after(cur, g, 0.0)
        lines.append(cur)
    return lines


def _style_key(g: Glyph) -> tuple[str, float, str]:
    return (g.font, round(g.size, 1), g.color.rgb or "")


def _link_for(box: tuple[float, float, float, float], links: list[Annotation]) -> str | None:
    cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
    for a in links:
        x0, y0, x1, y1 = a.bbox
        if x0 - 1 <= cx <= x1 + 1 and y0 - 1 <= cy <= y1 + 1:
            if a.uri:
                return a.uri
            if a.dest_page is not None:
                return f"#slide={a.dest_page + 1}"
    return None


def line_blocks(page: Page, fonts: dict[str, FontRef]) -> list[TextBlock]:
    """One TextBlock (one paragraph, one line) per visual text line of ``page``."""
    visible = [g for g in page.glyphs if g.render_mode not in (3, 7)]
    links = [a for a in page.annotations if a.subtype == "link"]

    def adv(g: Glyph, ch: str) -> float | None:
        fr = fonts.get(g.font)
        fam = resolve_family(fr.family) if fr else g.font
        return advance_em(fam, bool(fr and fr.bold), bool(fr and fr.italic), ch)

    blocks: list[TextBlock] = []
    for k, ln in enumerate(_lines(visible, adv)):
        runs: list[TextRun] = []
        last_key: tuple[str, float, str] | None = None
        for g, text, cs in ln.glyphs:
            key = _style_key(g)
            if runs and cs is None and last_key is not None and (
                    key == last_key or (text == " " and runs[-1].char_spacing == 0.0)):
                runs[-1].text += text
                continue
            fr = fonts.get(g.font)
            runs.append(TextRun(id=f"p{page.index}_t{k + 1}_r{len(runs) + 1}", bbox=g.bbox,
                                z=g.z, source="rule", text=text,
                                font=resolve_family(fr.family) if fr else g.font,
                                size=g.size, color=g.color,
                                bold=bool(fr and fr.bold), italic=bool(fr and fr.italic),
                                char_spacing=cs or 0.0))
            last_key = None if cs is not None else key
        while len(runs) > 1 and not runs[-1].text.strip():
            runs.pop()
        runs[-1].text = runs[-1].text.rstrip(" ") or runs[-1].text
        big = max(runs, key=lambda r: r.size)
        size = max(big.size, 1.0)
        off = first_baseline_offset(big.font, size)
        o = (ln.first.origin_x, ln.first.baseline)
        w = max(ln.end, 0.5 * size) + 0.05 * size
        h = LINE_FACTOR * size
        tl = (o[0] - off * ln.n[0], o[1] - off * ln.n[1])
        cx = tl[0] + w / 2 * ln.u[0] + h / 2 * ln.n[0]
        cy = tl[1] + w / 2 * ln.u[1] + h / 2 * ln.n[1]
        frame = (round(cx - w / 2, 3), round(cy - h / 2, 3), round(cx + w / 2, 3),
                 round(cy + h / 2, 3))
        link = _link_for(frame, links) if links else None
        if link:
            for r in runs:
                r.hyperlink = link
        rot_cw = round((-ln.rot) % 360.0, 3)
        blocks.append(TextBlock(
            id=f"p{page.index}_t{k + 1}", bbox=frame, z=ln.first.z, source="rule",
            paragraphs=[Paragraph(id=f"p{page.index}_t{k + 1}_p1", bbox=frame, z=ln.first.z,
                                  source="rule", runs=runs)],
            autofit="none", wrap=False, rotation=0.0 if rot_cw >= 359.999 else rot_cw,
            margins=(0.0, 0.0, 0.0, 0.0), notes=["baseline_line"]))
    return blocks
