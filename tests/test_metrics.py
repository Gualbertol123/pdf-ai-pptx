from __future__ import annotations

import numpy as np
import pytest

from pdf2pptx.verify import metrics
from pdf2pptx.verify.metrics import Word
from pdf2pptx.verify.pptx_stats import autonum_label


def _line(y: float, words: list[str], x: float = 10.0, dx: float = 0.0) -> list[Word]:
    out, cx = [], x + dx
    for w in words:
        out.append(Word(w, cx, y, cx + 6 * len(w), y + 12))
        cx += 6 * len(w) + 4
    return out


def test_ssim_identical_and_different() -> None:
    rng = np.random.default_rng(0)
    img = (rng.random((120, 200)) * 255).astype(np.uint8)
    s, smap = metrics.ssim_map(img, img)
    assert s == pytest.approx(1.0)
    assert metrics.region_ssim_min(smap) == pytest.approx(1.0)
    other = img.copy()
    other[:60, :100] = 255  # damage one quadrant
    s2, smap2 = metrics.ssim_map(img, other)
    assert s2 < 0.9
    grid = metrics.region_scores(smap2)
    assert grid[0][0] < 0.2 and grid[3][3] == pytest.approx(1.0)
    with pytest.raises(ValueError):
        metrics.ssim_map(img, img[:-1])


def test_text_sim_normalises() -> None:
    assert metrics.text_sim("• Hello   world", "Hello\nworld") == 1.0
    assert metrics.text_sim("", "") == 1.0
    assert metrics.text_sim("abc", "") == 0.0
    assert 0.5 < metrics.text_sim("Perché è così", "Perche e cosi") < 1.0


def test_token_and_real_text_recall() -> None:
    assert metrics.token_recall("a b c d", "a b") == 0.5
    assert metrics.token_recall("", "x") is None
    # char weighted: the long word counts more
    assert metrics.real_text_ratio("aaaaaaaa b", "aaaaaaaa") == pytest.approx(8 / 9)


def test_word_matching_perfect() -> None:
    pdf = [_line(10, ["Hello", "wonderful", "world"]), _line(30, ["second", "line"])]
    m, n = metrics.match_words(pdf, pdf)
    assert n == 5 and len(m) == 5
    assert metrics.word_recall(pdf, pdf) == 1.0
    assert metrics.pos_err_pt(pdf, pdf) == 0.0
    assert metrics.linebreak_match(pdf, pdf) == 1.0


def test_pos_err_offset() -> None:
    pdf = [_line(10, ["alpha", "beta"])]
    ppt = [_line(10, ["alpha", "beta"], dx=3.0)]
    assert metrics.pos_err_pt(pdf, ppt) == pytest.approx(3.0)


def test_linebreak_detects_rewrap() -> None:
    pdf = [_line(10, ["one", "two", "three"]), _line(30, ["four", "five"])]
    ppt = [_line(10, ["one", "two"]), _line(30, ["three", "four", "five"])]
    # first PDF line ends with "three", which no longer ends a line; "five" still does
    assert metrics.linebreak_match(pdf, ppt) == 0.5
    assert metrics.word_recall(pdf, ppt) == 1.0


def test_matching_handles_repeated_words() -> None:
    pdf = [_line(10, ["go", "go"]), _line(40, ["go"])]
    ppt = [_line(40, ["go"]), _line(10, ["go", "go"])]  # order differs, positions equal
    assert metrics.pos_err_pt(pdf, ppt) == 0.0
    assert metrics.linebreak_match(pdf, ppt) == 1.0


def test_punctuation_split() -> None:
    pdf = [[Word("world,", 0, 0, 36, 12)]]
    ppt = [[Word("world", 0, 0, 30, 12)]]
    assert metrics.word_recall(pdf, ppt) == 1.0


def test_composite_renormalises() -> None:
    full = {"ssim": 1.0, "text_sim": 1.0, "real_text_ratio": 1.0, "linebreak_match": 1.0,
            "native_shape_ratio": 1.0}
    assert metrics.composite(full) == pytest.approx(1.0)
    partial = {"ssim": 0.5, "text_sim": 1.0, "real_text_ratio": None}
    assert metrics.composite(partial) == pytest.approx((0.4 * 0.5 + 0.2) / 0.6)
    assert metrics.composite({}) is None
    assert metrics.native_shape_ratio(0, 0) is None
    assert metrics.native_shape_ratio(4, 1) == 0.8


def test_autonum_labels() -> None:
    assert autonum_label("arabicPeriod", 3) == "3."
    assert autonum_label("alphaLcParenR", 2) == "b)"
    assert autonum_label("romanUcPeriod", 4) == "IV."
    assert autonum_label("arabicParenBoth", 1) == "(1)"
