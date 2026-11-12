"""Text grouping on hand-made pages: columns, justified text, lists, scripts, rotation."""

from __future__ import annotations

import math

import pytest

from pdf2pptx.analyze import lists, reading_order, roles, text_group
from pdf2pptx.fonts.match import FontMatch
from pdf2pptx.ir import Color, Document, FontRef, Glyph, Image, Page, TextBlock

ADV = 0.5  # every test glyph is 0.5 em wide (a monospace "font")
FONT = FontRef(id="F", pdf_name="Arial", family="Arial",
               widths={chr(c): ADV for c in range(32, 0x2100) if chr(c).isprintable()})
MATCH = {"F": FontMatch("Arial", False, False, 1.0, "name")}


class PageBuilder:
    def __init__(self, w: float = 960, h: float = 540) -> None:
        self.page = Page(index=0, width=w, height=h)
        self.n = 0

    def text(self, s: str, x: float, y: float, size: float = 12.0, rot: float = 0.0,
             gaps: dict[int, float] | None = None, color: str = "000000") -> float:
        """Draw ``s`` with its first origin at (x, y); returns the end (along)."""
        u = (math.cos(math.radians(rot)), -math.sin(math.radians(rot)))
        pos = 0.0
        for i, ch in enumerate(s):
            if gaps and i in gaps:
                pos += gaps[i]
            ox, oy = x + pos * u[0], y + pos * u[1]
            w = ADV * size
            corners = [(ox + a * u[0] + b * -u[1], oy + a * u[1] + b * u[0])
                       for a in (0, w) for b in (0.2 * size, -0.7 * size)]
            xs, ys = [c[0] for c in corners], [c[1] for c in corners]
            self.page.glyphs.append(Glyph(
                id=f"g{self.n}", bbox=(min(xs), min(ys), max(xs), max(ys)), z=1, text=ch,
                font="F", size=size, color=Color(rgb=color), baseline=round(oy, 3),
                origin_x=round(ox, 3), rotation=rot))
            self.n += 1
            pos += w
        return pos

    def wrapped(self, words: list[str], x: float, y: float, chars: int, size: float = 12.0,
                pitch: float | None = None, justify: bool = False) -> int:
        """Greedy-wrap ``words`` at ``chars`` per line (what a PDF engine would do)."""
        pitch = pitch or 1.2 * size
        lines: list[list[str]] = [[]]
        for w in words:
            if lines[-1] and len(" ".join([*lines[-1], w])) > chars:
                lines.append([])
            lines[-1].append(w)
        for k, ln in enumerate(lines):
            text = " ".join(ln)
            gaps = None
            if justify and k + 1 < len(lines) and len(ln) > 1:
                extra = (chars - len(text)) * ADV * size / (len(ln) - 1)
                gaps, pos = {}, 0
                for w in ln[:-1]:
                    pos += len(w) + 1
                    gaps[pos] = extra
            self.text(text, x, y + k * pitch, size, gaps=gaps)
        return len(lines)

    def blocks(self) -> list[TextBlock]:
        return text_group.group_page(self.page, {"F": FONT}, MATCH)


TEXT = ("the program started in a single neighbourhood and now runs in five each site has "
        "a coordinator a group of volunteers and a small budget for materials")
WORDS = TEXT.split(" ")


def ptext(p: object) -> str:
    return "".join(r.text for r in p.runs)  # type: ignore[attr-defined]


def test_two_columns_and_reading_order() -> None:
    b = PageBuilder()
    b.text("Two columns", 50, 40, size=28)
    b.wrapped(WORDS, 50, 100, chars=34)
    b.wrapped(WORDS[::-1], 520, 100, chars=34)
    blocks = b.blocks()
    assert len(blocks) == 3
    boxes = [tb.bbox for tb in blocks]
    order = [blocks[i] for i in reading_order.xy_cut(boxes)]
    assert ptext(order[0].paragraphs[0]) == "Two columns"
    assert order[1].bbox[0] < 100 < 500 < order[2].bbox[0]  # left column before right
    for tb in order[1:]:
        assert len(tb.paragraphs) == 1  # wrapped lines join one paragraph
        assert tb.paragraphs[0].alignment == "left"
    assert ptext(order[1].paragraphs[0]) == " ".join(WORDS)
    # the frame is wide enough for every line and too narrow for the next word
    w = order[1].bbox[2] - order[1].bbox[0]
    assert 34 * 6 - 1 <= w < 34 * 6 + 6 * 7


def test_justified_paragraph() -> None:
    b = PageBuilder()
    n = b.wrapped(WORDS, 60, 100, chars=40, justify=True)
    (tb,) = b.blocks()
    (p,) = tb.paragraphs
    assert n >= 3 and p.alignment == "justify"
    assert abs((tb.bbox[2] - tb.bbox[0]) - 40 * 6) < 0.6  # frame = the justified column
    assert ptext(p) == " ".join(WORDS)


def test_paragraph_breaks_line_spacing_and_space_after() -> None:
    b = PageBuilder()
    n1 = b.wrapped(WORDS[:14], 60, 100, chars=40, pitch=18.0)
    b.wrapped(WORDS[14:], 60, 100 + (n1 - 1) * 18.0 + 18.0 + 9.0, chars=40, pitch=18.0)
    (tb,) = b.blocks()
    assert len(tb.paragraphs) == 2
    p1 = tb.paragraphs[0]
    assert p1.line_spacing_rule == "exact" and abs((p1.line_spacing or 0) - 18.0) < 0.01
    assert abs(p1.space_after - 9.0) < 0.6  # rounded the way PowerPoint will round it


def test_nested_bullets_and_numbering() -> None:
    b = PageBuilder()
    items = [(0, "•", "Research"), (1, "–", "Interviews with partners"),
             (2, "•", "Transcribed and coded"), (1, "–", "Survey of visitors"),
             (0, "•", "Design")]
    for k, (lvl, ch, txt) in enumerate(items):
        x = 50 + lvl * 36
        b.text(ch, x, 100 + k * 24, size=16)
        b.text(txt, x + 18, 100 + k * 24, size=16)
    nums = [(0, "1.", "Read the guidelines"), (0, "2.", "Check eligibility"),
            (1, "a)", "Budget"), (1, "b)", "Letters"), (0, "3.", "Submit")]
    for k, (lvl, lab, txt) in enumerate(nums):
        x = 520 + lvl * 30
        b.text(lab, x, 100 + k * 24, size=16)
        b.text(txt, x + 24, 100 + k * 24, size=16)
    bl, nl = sorted(b.blocks(), key=lambda tb: tb.bbox[0])
    assert [p.level for p in bl.paragraphs] == [0, 1, 2, 1, 0]
    assert [p.bullet.char for p in bl.paragraphs] == ["•", "–", "•", "–", "•"]
    assert [ptext(p) for p in bl.paragraphs] == [t for _, _, t in items]  # bullet removed
    p0 = bl.paragraphs[0]
    assert abs(p0.indent_left - 18) < 0.01 and abs(p0.indent_first + 18) < 0.01
    assert [p.level for p in nl.paragraphs] == [0, 0, 1, 1, 0]
    assert [p.bullet.scheme for p in nl.paragraphs] == [
        "arabicPeriod", "arabicPeriod", "alphaLcParenR", "alphaLcParenR", "arabicPeriod"]
    # PowerPoint numbers a run of equal buAutoNum; "3." continues after the sub-items
    assert [p.bullet.start_at for p in nl.paragraphs] == [1, 1, 1, 1, 1]


def test_superscript_and_subscript_stay_in_the_line() -> None:
    b = PageBuilder()
    end = b.text("E = mc", 50, 100, size=20)
    b.text("2", 50 + end, 94, size=13.33)  # raised 6 pt, 2/3 size
    end2 = end + 13.33 * ADV
    end3 = end2 + b.text(" and H", 50 + end2, 100, size=20)
    b.text("2", 50 + end3, 104, size=13.33)  # lowered 4 pt
    b.text("O", 50 + end3 + 13.33 * ADV, 100, size=20)
    (tb,) = b.blocks()
    (p,) = tb.paragraphs
    assert ptext(p) == "E = mc2 and H2O"
    shifts = [(r.text, r.baseline_shift, r.size) for r in p.runs if r.baseline_shift]
    assert [s[0] for s in shifts] == ["2", "2"]
    assert shifts[0][1] > 0 > shifts[1][1]
    assert all(abs(s[2] - 20) < 0.1 for s in shifts)  # nominal size: PowerPoint shows 2/3
    assert abs(shifts[0][1] - 30.0) < 0.5  # 6 pt of a 20 pt nominal size


def test_rotated_block() -> None:
    b = PageBuilder()
    # text running upwards (90 degrees counter-clockwise), two wrapped lines
    b.text("Rotated label that", 100, 400, size=14, rot=90)
    b.text("wraps once", 100 + 1.2 * 14, 400, size=14, rot=90)
    (tb,) = b.blocks()
    assert tb.rotation == 270.0
    assert [ptext(p) for p in tb.paragraphs] == ["Rotated label that wraps once"]
    # the unrotated frame is wide along the text and ~2 lines tall
    w, h = tb.bbox[2] - tb.bbox[0], tb.bbox[3] - tb.bbox[1]
    assert w > 18 * 7 - 1 and 2 * 14 < h < 3 * 14


def test_picture_bullet_and_hyphenation() -> None:
    b = PageBuilder()
    for k in range(2):
        y = 100 + k * 20
        b.page.elements.append(Image(id=f"img{k}", bbox=(50, y - 10, 58, y - 2), file="x.png"))
        b.text(["First item", "Second item"][k], 66, y, size=14)
    b.text("A long word broken by exam-", 300, 100, size=14)
    b.text("ple here.", 300, 116.8, size=14)
    blocks = sorted(b.blocks(), key=lambda tb: tb.bbox[0])
    pics = [p.bullet for p in blocks[0].paragraphs]
    assert [x.kind for x in pics] == ["picture", "picture"]
    assert [x.image_id for x in pics] == ["img0", "img1"]
    assert text_group.picture_bullet_ids(blocks) == {"img0", "img1"}
    (hp,) = blocks[1].paragraphs
    assert ptext(hp) == "A long word broken by exam-ple here."
    assert any(n.startswith("hyphen:") for n in hp.notes)


def test_invisible_text_is_dropped() -> None:
    b = PageBuilder()
    b.text("Visible", 50, 100)
    n = b.n
    b.text("Hidden", 50, 200)
    for g in b.page.glyphs[n:]:
        g.color = Color(rgb="000000", alpha=0.0)
    assert [ptext(tb.paragraphs[0]) for tb in b.blocks()] == ["Visible"]


def test_axis_ticks_are_not_a_paragraph() -> None:
    b = PageBuilder()
    for k, v in enumerate(["400", "300", "200", "100", "0"]):
        b.text(v, 80 - len(v) * 6, 100 + k * 30)  # right-aligned, 2.5 lines apart
    blocks = b.blocks()
    assert len(blocks) == 5 and all(len(tb.paragraphs) == 1 for tb in blocks)


# --------------------------------------------------------------------------- lists


def test_number_labels_and_roman_resolution() -> None:
    assert lists.parse_number("iv.") == {"romanLcPeriod": 4}
    assert lists.parse_number("(3)") == {"arabicParenBoth": 3}
    assert lists.parse_number("2026.") == {}
    labs = [lists.detect_label([(t, 0, 5, len(t)), ("x", 10, 20, 1)], 10) for t in
            ("h.", "i.", "j.")]
    lists.resolve_numbering([lb for lb in labs if lb])
    assert [lb.scheme for lb in labs if lb] == ["alphaLcPeriod"] * 3
    labs = [lists.detect_label([(t, 0, 5, len(t)), ("x", 10, 20, 1)], 10) for t in
            ("i.", "ii.", "iii.")]
    lists.resolve_numbering([lb for lb in labs if lb])
    assert [(lb.scheme, lb.value) for lb in labs if lb] == [("romanLcPeriod", k) for k in (1, 2, 3)]
    assert lists.assign_levels([50, 86, 122, 86.5, 50.2], 3.0) == [0, 1, 2, 1, 0]
    # a dash glued to a word is not a bullet; a strong bullet is
    assert lists.detect_label([("-x", 0, 10, 2), ("y", 14, 20, 1)], 10) is None
    assert lists.detect_label([("•Text", 0, 25, 5), ("y", 29, 35, 1)], 10) is not None


# --------------------------------------------------------------------------- reading order


def test_xy_cut_prefers_columns_over_aligned_gaps() -> None:
    # two columns whose paragraph gaps line up, under a full-width title
    boxes = [(50, 20, 900, 60),
             (50, 100, 450, 200), (500, 100, 900, 200),
             (50, 220, 450, 320), (500, 220, 900, 320)]
    assert reading_order.xy_cut(boxes) == [0, 1, 3, 2, 4]
    # a plain grid of short items reads row by row
    grid = [(50, 100, 100, 110), (200, 100, 250, 110), (50, 130, 100, 140),
            (200, 130, 250, 140)]
    assert reading_order.xy_cut(grid) in ([0, 1, 2, 3], [0, 2, 1, 3])


# --------------------------------------------------------------------------- roles


def test_roles_title_footer_page_number() -> None:
    pages = []
    blocks: dict[int, list[TextBlock]] = {}
    for i in range(3):
        b = PageBuilder()
        b.page.index = i
        b.text(f"Section {i + 1} title", 50, 60, size=36)
        b.wrapped(WORDS, 50, 150, chars=60, size=16)
        b.text("Community Report 2026", 50, 515, size=10)
        b.text(str(i + 1), 900, 515, size=10)
        pages.append(b.page)
        blocks[i] = text_group.group_page(b.page, {"F": FONT}, MATCH)
    doc = Document(pages=pages, fonts={"F": FONT})
    roles.assign_roles(doc, blocks)
    for bl in blocks.values():
        by_role = {tb.role: tb for tb in bl}
        assert set(by_role) == {"title", "body", "footer", "page_number"}
        num = by_role["page_number"]
        assert num.role_confidence >= 0.9
        assert num.paragraphs[0].runs[0].field_type == "slidenum"
        assert by_role["footer"].role_confidence >= 0.9  # repeated on every page


@pytest.mark.parametrize("text,ok", [("7", True), ("Page 7", True), ("7 / 20", True),
                                     ("- 7 -", True), ("2026 report", False)])
def test_page_number_pattern(text: str, ok: bool) -> None:
    assert bool(roles.PAGE_NUM_RE.match(text)) is ok
