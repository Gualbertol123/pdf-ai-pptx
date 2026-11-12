"""The PowerPoint COM engine: layout accuracy, recovery, timeouts, process hygiene."""

from __future__ import annotations

import time
from pathlib import Path

import psutil
import pytest
from pptx import Presentation
from pptx.util import Pt

from pdf2pptx.verify import ppt_com

from .conftest import CORPUS

pytestmark = pytest.mark.powerpoint


def _layout_deck(path: Path) -> list[int]:
    """Five slides whose rendered line count is known by construction."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(720), Pt(405)
    blank = prs.slide_layouts[6]
    expected = []

    def box(slide, x, y, w, h, wrap=False):  # type: ignore[no-untyped-def]
        tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        return tf

    # 1. three paragraphs, no wrapping -> 3 lines
    tf = box(prs.slides.add_slide(blank), 40, 40, 400, 200)
    tf.text = "first paragraph"
    for t in ("second paragraph", "third one here"):
        tf.add_paragraph().text = t
    expected.append(3)

    # 2. one paragraph with two soft line breaks (vertical tab) -> 3 lines
    tf = box(prs.slides.add_slide(blank), 40, 40, 400, 200)
    tf.paragraphs[0].text = "alpha beta\vgamma delta\vepsilon"
    expected.append(3)

    # 3. a 2x2 table, one cell with an explicit second paragraph -> 5 lines
    s = prs.slides.add_slide(blank)
    tbl = s.shapes.add_table(2, 2, Pt(40), Pt(40), Pt(400), Pt(120)).table
    for (r, c), t in {(0, 0): "one", (0, 1): "two", (1, 0): "three", (1, 1): "four"}.items():
        tbl.cell(r, c).text = t
    tbl.cell(1, 1).text_frame.add_paragraph().text = "five"
    expected.append(5)

    # 4. two text boxes inside a group -> 2 lines
    s = prs.slides.add_slide(blank)
    grp = s.shapes.add_group_shape()
    for k in range(2):
        tb = grp.shapes.add_textbox(Pt(40), Pt(40 + 60 * k), Pt(300), Pt(40))
        tb.text_frame.text = f"grouped line {k + 1}"
    expected.append(2)

    # 5. wrapping: 10 five-char words of 20 pt Consolas (0.55 em -> 11 pt per char,
    #    "aaaa " = 55 pt) in a 120 pt box -> 2 words per line -> 5 lines
    tf = box(prs.slides.add_slide(blank), 40, 40, 120, 300, wrap=True)
    run = tf.paragraphs[0].add_run()
    run.text = " ".join(["word"] * 10)
    run.font.name, run.font.size = "Consolas", Pt(20)
    expected.append(5)

    prs.save(str(path))
    return expected


def test_text_layout_line_counts(tmp_path: Path) -> None:
    pptx = tmp_path / "layout.pptx"
    expected = _layout_deck(pptx)
    with ppt_com.PowerPointEngine() as eng, eng.opened(pptx) as info:
        assert info.slides == 5 and (info.width_pt, info.height_pt) == (720, 405)
        got = [len(ppt_com.layout_lines(eng.text_layout(info, i))) for i in range(5)]
        assert got == expected
        # bounds are slide points: slide 1 starts at the box origin (zero insets)
        first = ppt_com.layout_lines(eng.text_layout(info, 0))[0]
        assert first.bound[0] == pytest.approx(40, abs=1.5)
        assert first.bound[1] == pytest.approx(40, abs=1.5)
        assert [w.text for w in first.words] == ["first", "paragraph"]
        wrapped = ppt_com.layout_lines(eng.text_layout(info, 4))
        assert all(len(ln.words) == 2 for ln in wrapped)
        assert wrapped[0].words[0].x1 - wrapped[0].words[0].x0 == pytest.approx(44, abs=2)
        with pytest.raises(ppt_com.PowerPointError):
            eng.text_layout(info, 9)


def test_render_only_requested_slides(tmp_path: Path) -> None:
    pptx = tmp_path / "layout.pptx"
    _layout_deck(pptx)
    with ppt_com.PowerPointEngine() as eng, eng.opened(pptx) as info:
        paths = eng.render(info, [3, 1], 72, tmp_path / "r")
        assert [p.name for p in paths] == ["slide_004.png", "slide_002.png"]
        assert sorted(p.name for p in (tmp_path / "r").iterdir()) == ["slide_002.png",
                                                                      "slide_004.png"]
        from PIL import Image

        assert Image.open(paths[0]).size == (720, 405)  # 72 dpi = 1 px per pt
        # re-opening the unchanged file reuses the open presentation
        assert eng.open(pptx).handle == info.handle


def test_export_pdf(tmp_path: Path) -> None:
    """The path ``corpus build`` uses to make the synthetic PDFs."""
    import pypdfium2 as pdfium

    pptx = tmp_path / "layout.pptx"
    _layout_deck(pptx)
    with ppt_com.PowerPointEngine() as eng, eng.opened(pptx) as info:
        out = eng.export_pdf(info, tmp_path / "sub" / "layout.pdf")
    doc = pdfium.PdfDocument(str(out))
    try:
        assert len(doc) == 5 and doc[0].get_size() == pytest.approx((720, 405))
    finally:
        doc.close()


@pytest.mark.corpus
def test_stress_recovers_from_killed_worker() -> None:
    """200 renders in a row on corpus decks; the worker is killed after call 70 and
    its PowerPoint after call 140. Every call must still complete."""
    decks = sorted((CORPUS / "synthetic").glob("d*.pptx"))[:8]
    assert decks
    ppt_com.shutdown_shared()  # so this engine starts (and owns) its own PowerPoint
    out = Path(ppt_com.REGISTRY_DIR).parent / "stress"
    eng = ppt_com.PowerPointEngine(call_timeout=60)
    done = 0
    t0 = time.perf_counter()
    try:
        eng.start()
        for k in range(200):
            deck = decks[k % len(decks)]
            info = eng.open(deck)
            eng.render(info, [k % info.slides], 50, out)
            done += 1
            if k == 70:
                eng.kill_worker()  # the client must notice and restart
            if k == 140:
                if eng.owned_pids:  # PowerPoint itself dies under the worker
                    for pid in eng.owned_pids:
                        psutil.Process(pid).kill()
                else:  # attached to a PowerPoint the user started: never touch it
                    eng.kill_worker()
        assert done == 200
        assert eng.restarts >= 2
        assert len(eng.owned_history) >= 2  # a new PowerPoint after each crash
        # the instances of the dead workers are gone; only the live one remains
        assert set(eng.owned_history) & set(ppt_com.owned_alive()) == set(eng.owned_pids)
    finally:
        eng.stop()
    print(f"stress: 200 renders, {eng.restarts} restarts, {time.perf_counter() - t0:.1f} s")
    assert not set(eng.owned_history) & set(ppt_com.owned_alive())


def test_timeout_restarts_worker() -> None:
    ppt_com.shutdown_shared()
    eng = ppt_com.PowerPointEngine(max_restarts=0, hang_timeout=60)
    try:
        eng.start()
        pid = eng.worker_pid
        with pytest.raises(ppt_com.PowerPointUnavailable):
            eng._call("hang", 30, timeout=2)
        assert eng.ping() < 5  # a fresh worker answers
        assert eng.worker_pid != pid
    finally:
        eng.stop()
    assert not set(eng.owned_history) & set(ppt_com.owned_alive())


def test_bad_file_is_a_call_error_not_a_crash(tmp_path: Path) -> None:
    bad = tmp_path / "broken.pptx"
    bad.write_bytes(b"not a zip at all")
    with ppt_com.PowerPointEngine() as eng:
        with pytest.raises(ppt_com.PowerPointError):
            eng.open(bad)
        assert eng.restarts == 0 and eng.ping() < 5
