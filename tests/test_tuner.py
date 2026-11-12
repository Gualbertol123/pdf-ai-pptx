"""Fidelity tuner (M5): wrap model, fixes, mode limits, and end-to-end runs in PowerPoint."""

from __future__ import annotations

import dataclasses
import hashlib
import zipfile
from pathlib import Path

import pytest

from pdf2pptx.build.writer import LINE_BREAK, TextOptions, write_pptx
from pdf2pptx.config import MODE_LIMITS, load_config
from pdf2pptx.ir import Document, TextBlock
from pdf2pptx.verify import tuner
from pdf2pptx.verify.tuner import Line, PptLine, WrapModel, evaluate, solve_wrap

from .test_text_group import WORDS, PageBuilder

# --------------------------------------------------------------------------- unit


def _wrapped_block(chars: int = 40) -> tuple[TextBlock, dict, int]:
    tb, _, glyphs, n = _wrapped_page(chars)
    return tb, glyphs, n


def _wrapped_page(chars: int = 40) -> tuple[TextBlock, Document, dict, int]:
    b = PageBuilder()
    n = b.wrapped(WORDS, 60, 100, chars)
    blocks = b.blocks()
    assert len(blocks) == 1
    b.page.elements = list(blocks)
    return blocks[0], Document(pages=[b.page]), {g.id: g for g in b.page.glyphs}, n


def test_pdf_lines_follow_the_glyph_baselines() -> None:
    tb, glyphs, n = _wrapped_block()
    lines = tuner.pdf_lines(tb, glyphs)
    assert lines is not None and len(lines) == n
    stream = "".join(tuner._ns(t) for t in tuner.written_paragraphs(tb))
    assert lines[-1].end == len(stream)
    # every line starts with a word of the text and ends at a word end
    words = [w for w in WORDS]
    ends, k = set(), 0
    for w in words:
        k += len(w)
        ends.add(k)
    assert all(ln.end in ends for ln in lines)
    assert [ln.para_last for ln in lines] == [False] * (n - 1) + [True]
    assert all(ln.x1 is not None and ln.x1 > (ln.x0 or 0) for ln in lines)
    assert [round(ln.baseline or 0) for ln in lines] == [round(100 + 14.4 * i) for i in range(n)]


def test_tokens_split_after_hyphens() -> None:
    toks = tuner._tokens(["well-known items, x - y", "next"])
    texts = [(t.start, t.end, t.spaces, t.para) for t in toks]
    # "well-" | "known" | "items," | "x" | "-" | "y" || "next"
    assert [e - s for s, e, _, _ in texts] == [5, 5, 6, 1, 1, 1, 4]
    assert [sp for _, _, sp, _ in texts][:2] == [0, 1]
    assert texts[-1][3] == 1


def _ppt(ends: list[int]) -> list[PptLine]:
    out, start = [], 0
    for e in ends:
        out.append(PptLine(-1, start, e))
        start = e
    return out


def test_evaluate_counts_break_differences() -> None:
    ref = [Line(0, 0, 10), Line(0, 10, 20), Line(0, 20, 25, para_last=True)]
    assert evaluate(ref, _ppt([10, 20, 25]), True).breaks == 0
    ev = evaluate(ref, _ppt([8, 18, 25]), True)
    assert ev.breaks == 4 and not ev.ok
    assert evaluate(ref, _ppt([25]), False).breaks == 2  # unmapped: line count difference


def test_solve_wrap_prefers_width_then_spc_then_size() -> None:
    lim = MODE_LIMITS["balanced"]
    # three lines need 100, 98, 95 pt; the next word would make them 110, 109
    m = WrapModel(req=[(0, 100, 50), (0, 98, 50), (0, 95, 40)],
                  nxt=[(0, 110, 55), (0, 109, 55)], width=90)
    fix = solve_wrap(m, 90, lim, 20.0, (-0.8, 0.8), 1.0)
    assert fix is not None and fix.delta == 0 and fix.size == 0
    assert 100 + tuner.FIT_MARGIN <= fix.width < 109 - tuner.FIT_MARGIN
    # no width room (limit 0): spc has to make the lines fit 90 pt
    tight = dataclasses.replace(lim, max_width_change=0.0)
    fix = solve_wrap(m, 90, tight, 20.0, (-0.8, 0.8), 1.0)
    assert fix is not None and fix.delta < 0 and fix.size == 0
    lo, hi = m.interval(fix.delta, 1.0)
    assert lo + tuner.FIT_MARGIN <= 90 < hi
    # neither width nor spc: the font size (0.1 pt steps); 98 pt needs about -2.5%
    nospc = dataclasses.replace(tight, max_char_spacing_pt=0.0)
    m = dataclasses.replace(m, width=98)
    fix = solve_wrap(m, 98, nospc, 20.0, (0.0, 0.0), 1.0)
    assert fix is not None and fix.delta == 0 and fix.size < 0
    assert abs(fix.size * 10 - round(fix.size * 10)) < 1e-9
    assert abs(fix.size) <= 0.05 * 20.0 + 1e-9
    # editable limits: nothing fits within ±3% size, ±0.4 pt spc, ±10% width
    far = WrapModel(req=[(0, 150, 50)], nxt=[(0, 152, 52)], width=100)
    assert solve_wrap(far, 100, MODE_LIMITS["editable"], 20.0, (-0.4, 0.4), 1.0) is None


def test_force_breaks_puts_soft_breaks_at_the_pdf_wraps(tmp_path: Path) -> None:
    tb, glyphs, n = _wrapped_block()
    lines = tuner.pdf_lines(tb, glyphs)
    assert lines is not None
    assert tuner.force_breaks(tb, lines) == n - 1
    text = "".join(r.text for r in tb.paragraphs[0].runs)
    pieces = text.split(LINE_BREAK)
    assert len(pieces) == n and all(p == p.strip() for p in pieces)
    doc = Document(pages=[PageBuilder().page])
    doc.pages[0].elements = [tb]
    out = tmp_path / "br.pptx"
    write_pptx(doc, out, tmp_path, tmp_path / "none.pdf", TextOptions())
    xml = zipfile.ZipFile(out).read("ppt/slides/slide1.xml").decode()
    assert xml.count("<a:br>") == n - 1
    assert f'name="TextBlock {tb.id}"' in xml


def test_without_powerpoint_the_deck_is_left_alone(tmp_path: Path) -> None:
    _, doc, _, _ = _wrapped_page()
    out = tmp_path / "x.pptx"
    write_pptx(doc, out, tmp_path, tmp_path / "none.pdf", TextOptions())
    before = out.read_bytes()

    class Dead:
        def open(self, *a: object) -> None:
            raise RuntimeError("PowerPoint is gone")

    rep = tuner.tune_document(doc, out, tmp_path, tmp_path / "none.pdf", TextOptions(),
                              MODE_LIMITS["balanced"], engine=Dead())
    assert rep.skipped and "PowerPoint is gone" in rep.skipped
    assert out.read_bytes() == before


# --------------------------------------------------------------------------- PowerPoint


def _d23() -> tuple[Document, Path, Path]:
    from pdf2pptx.pipeline import analyzed

    cfg = load_config()
    pdf = next(cfg.corpus_dir.rglob("d23_long_paragraphs.pdf"))
    doc, wd, _ = analyzed(pdf, cfg)
    return doc, wd, pdf


def _narrow(doc: Document, factor: float) -> TextBlock:
    """Shrink the long (justified) paragraphs on page 1, keeping their left edge."""
    tb = next(e for e in doc.pages[0].elements if isinstance(e, TextBlock)
              and e.wrap and all(p.alignment in ("left", "justify") for p in e.paragraphs)
              and len(tuner.written_paragraphs(e)[0]) > 100)
    x0, y0, x1, y1 = tb.bbox
    tb.bbox = (x0, y0, x0 + (x1 - x0) * factor, y1)
    return tb


def _run(doc: Document, wd: Path, pdf: Path, out: Path, mode: str,
         limits: object | None = None, tune_flag: bool = True) -> tuner.TuneReport:
    from pdf2pptx.verify import ppt_com

    opts = TextOptions.for_mode(mode)
    write_pptx(doc, out, wd, pdf, opts)
    t = tuner.Tuner(doc, out, wd, pdf, opts, limits or MODE_LIMITS[mode],  # type: ignore[arg-type]
                    engine=ppt_com.shared(), rounds=4)
    rep = t.run()
    for st in t.states.values():  # never worse than the start
        assert st.first_eval is not None and st.best_eval is not None
        assert not st.best_eval.worse_than(st.first_eval)
    rep._tuner = t  # type: ignore[attr-defined]
    return rep


@pytest.mark.powerpoint
@pytest.mark.corpus
def test_a_block_10_percent_too_narrow_is_fixed(tmp_path: Path) -> None:
    doc, wd, pdf = _d23()
    tb = _narrow(doc, 0.9)
    rep = _run(doc, wd, pdf, tmp_path / "n.pptx", "balanced")
    st = rep._tuner.states[tb.id]  # type: ignore[attr-defined]
    assert st.first_eval.breaks > 0
    assert st.best_eval.breaks == 0 and st.best_eval.ok
    assert rep.rounds <= 4
    fixed = next(e for e in doc.pages[0].elements if e.id == tb.id)
    assert fixed.source == "tuner" and any(n.startswith("tuner:width:+") for n in fixed.notes)
    # deterministic: the same input gives the same adjustments and the same bytes
    doc2, _, _ = _d23()
    _narrow(doc2, 0.9)
    _run(doc2, wd, pdf, tmp_path / "n2.pptx", "balanced")
    assert hashlib.sha256((tmp_path / "n.pptx").read_bytes()).hexdigest() == hashlib.sha256(
        (tmp_path / "n2.pptx").read_bytes()).hexdigest()


def _breaks_in(pptx: Path) -> int:
    z = zipfile.ZipFile(pptx)
    return sum(z.read(n).decode().count("<a:br") for n in z.namelist()
               if n.startswith("ppt/slides/slide") and n.endswith(".xml"))


@pytest.mark.powerpoint
@pytest.mark.corpus
def test_editable_never_forces_line_breaks_fidelity_may(tmp_path: Path) -> None:
    # 30% too narrow: more than editable's ±10% width (and spc/size) can repair
    doc, wd, pdf = _d23()
    _narrow(doc, 0.7)
    rep = _run(doc, wd, pdf, tmp_path / "e.pptx", "editable")
    assert rep.forced_breaks == 0
    assert _breaks_in(tmp_path / "e.pptx") == 0
    # fidelity with no width/spc/size room falls back to forced breaks, and they match
    doc, wd, pdf = _d23()
    tb = _narrow(doc, 0.9)
    stiff = dataclasses.replace(MODE_LIMITS["fidelity"], max_width_change=0.0,
                                max_char_spacing_pt=0.0, max_font_size_change=0.0)
    rep = _run(doc, wd, pdf, tmp_path / "f.pptx", "fidelity", stiff)
    assert rep.forced_breaks > 0 and _breaks_in(tmp_path / "f.pptx") == rep.forced_breaks
    st = rep._tuner.states[tb.id]  # type: ignore[attr-defined]
    assert st.best_eval.breaks == 0


def test_element_shift_finds_a_consistent_offset() -> None:
    import numpy as np

    rng = np.random.default_rng(1234)
    ref = np.full((300, 400, 3), 255, np.uint8)
    patch = (rng.random((40, 60)) * 255).astype(np.uint8)
    ref[100:140, 150:210] = patch[..., None]
    got = np.full_like(ref, 255)
    got[103:143, 148:208] = patch[..., None]  # PowerPoint drew it 2 px left, 3 px down
    k = 72 / 150
    bbox = (150 * k, 100 * k, 210 * k, 140 * k)
    dx, dy = tuner.element_shift(ref, got, bbox, 150)  # type: ignore[misc]
    assert (round(dx / k), round(dy / k)) == (2, -3)  # move it back right and up
    assert tuner.element_shift(ref, ref, bbox, 150) is None  # already in place
    flat = np.full_like(ref, 255)
    assert tuner.element_shift(flat, flat, bbox, 150) is None  # nothing to match
