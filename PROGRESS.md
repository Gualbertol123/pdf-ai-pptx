# pdf2pptx: Progress

| # | Milestone | Status | Date |
|---|---|---|---|
| 1 | Foundation, IR, corpus, scoring harness | **done** | 2026-11-11 |
| 2 | Ingest and baseline converter | **done** | 2026-11-12 |
| 3 | PowerPoint check engine and debug report | **done** | 2026-11-12 |
| 4 | Text understanding and fonts | **done** | 2026-11-12 |
| 5 | Fidelity tuner | **done** | 2026-11-12 |
| 6 | Graphics and tables | not started | |
| 7 | Document structure | not started | |
| 8 | Hard pages (OCR, layout model, background cleanup) | not started | |
| 9 | Local model service, benchmark, judgment, charts | not started | |
| 10 | Document → deck, UI, packaging | not started | |

---

## Milestone 1: foundation, IR, corpus, scoring harness (2026-11-11)

### What exists

- **Project.** git repo, `.venv` (Python 3.12.7), `pyproject.toml` (src layout, console
  script `pdf2pptx`, minimum versions pinned), ruff / pytest / mypy (strict on `ir.py`
  and `build/units.py`), `.gitattributes` (LF). Dependencies are all MIT, BSD, Apache,
  HPND or PSF (no AGPL).
- **`ir.py`.** The full IR from PLAN.md §4.1: `Document`, `Page`, `Triage`, `FontRef`,
  `Theme`, `Master` / `Layout`, `Section`, and the 11 element kinds (`Glyph`, `TextRun`,
  `Paragraph`, `TextBlock`, `Shape`, `Image`, `Table` + `TableCell`, `Chart` +
  `ChartSeries`, `Equation`, `Group`, `Annotation`). Styling types are `Color`, `Fill`,
  `Gradient`, `Stroke`, `Bullet`, `PathSegment` and `Connector`. A generic,
  type-hint-driven JSON codec round-trips exactly (`from_json(to_json(d)) == d` and
  the re-dump is byte-identical). It rejects unknown kinds, unknown fields, a wrong
  `ir_version` and NaN.
- **`build/units.py`.** pt↔EMU, px, and DrawingML angle / percent / centipoint units.
- **`config.py`.** Modes `fidelity|balanced|editable|deck` with per-mode limits (used
  by the tuner and builder later), tiers `auto|lite|max` (detection stub: `max` needs
  llama.cpp, a GGUF model and at least 24 GB RAM), and paths. Settings load from
  defaults, then `pdf2pptx.toml`, then CLI flags.
- **CLI.** `convert`, `score`, `corpus`, `report`, `doctor`, `setup`. `convert`,
  `report` and `setup` are stubs that exit with code 3 and say which milestone
  adds them.
- **`doctor`.** Reports Python, OS, CPU, RAM, GPU and driver (WMI), PowerPoint COM
  (probed in a child process with a timeout), onnxruntime providers, `vulkan-1.dll`,
  the models dir and the tier.
- **Corpus** (`pdf2pptx corpus build`, about 60 s): **33 PDFs.**
  - 28 python-pptx decks exported to PDF through PowerPoint (56 slides): titles and
    multi-level bullets, autonumbering, 2 and 3 columns, mixed runs and
    super/subscript, Italian accented text, simple and merged tables, preset shapes,
    arrows / chevrons / rotation, straight and elbow connectors with arrowheads,
    gradients and alpha, images with crop and PNG alpha, bar / column / line / pie
    charts, a repeated footer + logo + slide-number field, a 5-step build-up,
    external and internal hyperlinks, speaker notes, p14 sections, 16:9 and 4:3,
    justified text and line spacing, 8 font families, a dark background with custom
    theme colours, a kitchen sink, and rotated text.
  - 3 reportlab reports: a 2-column report, an Italian report with a ruled table
    and a whitespace (unruled) table, and a figure with a caption. Truth for these
    is recorded per page while the layout runs.
  - 2 "scanned" PDFs: 200 DPI, rotated ≤ 1°, with noise, speckle and JPEG q55,
    rewrapped as image-only PDFs.
  - `corpus/manifest.json` plus `<name>.truth.json` per file. Truth covers text,
    paragraphs with level / bullet / fonts, shape kinds and presets with bboxes,
    tables (rows, cols, merges, cells), charts (type, categories, series values),
    hyperlinks, notes, sections, theme colours and fonts, and slide-number fields.
- **Scoring** (`verify/metrics.py`, `pdf_text.py`, `pptx_stats.py`, `score.py`).
  Every PLAN.md §6 metric is implemented. `seconds_per_page` is plumbed in but stays
  null until a converter exists.
  - `ssim` / `region_ssim_min` compare a pdfium render with PowerPoint's
    `Slide.Export` PNG at 150 DPI (grayscale Gaussian SSIM over a 4×4 grid).
  - `text_sim` is normalised Levenshtein. Bullet glyphs are ignored, and list
    labels (bullet characters, inherited master bullets and autonumbers) are
    rendered on the PPTX side.
  - `word_recall`, `pos_err_pt` and `linebreak_match` use an order-free word pairing
    (a Hungarian assignment per token) between pdfplumber word lines and
    PowerPoint's COM line/word bounds.
  - `real_text_ratio` is char-weighted token recall over text frames.
  - `native_shape_ratio`, `boxes_per_slide` and `composite` (weights renormalised
    over the metrics that are present).
  - Without PowerPoint, the PowerPoint-only scores are `null` and the reason is
    recorded.
- **Temporary COM** (`verify/_com_temp.py`). This is the only module that touches
  COM; Prompt 3 replaces it. It handles PDF export, PNG render and text layout. It
  never quits or kills a PowerPoint the user already had open, and it tracks and
  reaps the instances it started itself.

### Checks

| Check | Result |
|---|---|
| `pytest -q` | **48 passed** (IR round-trip for every element kind, units, config, metrics, corpus, CLI, PowerPoint scoring) |
| `ruff check .` / `mypy` (ir.py, units.py) | clean / clean |
| `pdf2pptx doctor` | PowerPoint COM = **available (16.0)**, Radeon 780M, 30.7 GB RAM, vulkan-1.dll present, onnxruntime not installed yet |
| `corpus build` | 33 PDFs + manifest + 33 truth files |
| Build twice | All 33 `truth.json` files and `manifest.json` are sha256-identical |
| No leftover POWERPNT.EXE | Verified after the test suite and after the corpus builds |

### Metric sanity: ground-truth PPTX vs its own PDF (`score --corpus --against truth`)

| set | slides | ssim median (mean, min) | text_sim median (mean) | pos_err_pt median | linebreak median | composite median |
|---|---|---|---|---|---|---|
| synthetic decks | 56 | **0.994** (0.992, 0.955) | **1.000** (0.955) | 0.38 | 1.000 | 0.997 |
| decks without charts / outlined text | 48 | mean 0.993, min 0.955 | mean **0.994** | | | |
| scanned (vs clean source deck) | 2 | 0.755 | 1.000 | 0.44 | 1.000 | 0.891 |
| **blank PPTX (sanity)** | 62 | 0.941 (0.912) | **0.000** | n/a | 0.000 | **0.418** |

The PLAN thresholds are ssim ≥ 0.97 and text_sim ≥ 0.99. The median meets both. The
text_sim mean sits below 0.99 only because of the slides listed below, where the PDF
really does differ from what the PPTX text model can see. Every case was inspected:

- **Charts** (d14, d15, d16, d26, d27): text_sim is 0.66–0.87. PowerPoint computes the
  axis tick labels ("0, 10, 20 …"), so they exist only in the PDF. Native charts
  arrive in M9, and chart-region handling is open until then.
- **d28 rotated text**: text_sim is 0.54. PowerPoint's PDF export writes text rotated
  at −30° and 8° as **outlines**. pdfium and pdfplumber find no glyphs for it. This
  makes a good triage test case for M2/M8, and it's tagged `pdf_outlined_text` in
  the manifest.
- **d19 hyperlinks**: text_sim is 0.92. The exported PDF contains the "Go to details"
  button text twice.

Blank decks score clearly low on text (0) and composite (0.42 vs 0.997). **SSIM alone
is lenient on sparse slides**: an empty white slide still gets 0.94 against a
mostly white page. `region_ssim_min` (0.80 vs 0.97) and the composite tell the two
apart. Judge visual progress on region_ssim_min as well, and consider an
ink-weighted SSIM in M3.

Full tables: `work/scores/truth/summary.md`, `work/scores/blank/summary.md`.

### Measured facts that later milestones rely on

- COM `TextRange2.Lines/Words` bounds are **layout boxes**. They include trailing
  spaces, leading and space-before, and the bottom includes the paragraph's
  space-after. Words are therefore trimmed (`TrimText`) and matched on
  (x-centre, bottom − space-after). Doing this took pos_err on truth from about
  4 pt to 0.38 pt.
- COM bounds for **table cell text are relative to the cell**. x includes the left
  margin and y excludes the top margin. With the offset fixed, table pos_err is
  0.16–0.18 pt.
- PDF export through COM needs backslash paths. Office here runs in an Italian
  locale, so error messages come back in Italian.
- pdfium vs PowerPoint rasterisation on an identical slide gives SSIM ≈ 0.997.

### Known gaps

- `convert`, `report` and `setup` are stubs (M2, M3, M8/M9).
- `seconds_per_page` is null until M2. Chart tick labels aren't modelled on the PPTX
  side. Rotated words (non-upright chars) aren't grouped by the scoring word
  extractor.
- `text_sim` uses content-stream order (PDF) against z-order (PPTX). PowerPoint
  exports keep these consistent, but arbitrary PDFs need the reading order from M4.
- Report-style PDFs have no answer-key PPTX, so they can only be scored against
  converter output (M2+).
- pytest prints a handled `0x80010108` COM disconnect from faulthandler, which goes
  away in M3 when COM moves out of the test process.

### Next

Milestone 2: pypdfium2 ingest, triage, baseline writer, `convert --corpus`, and
baseline scores in `work/scores/baseline.json`.

---

## Milestone 2: ingest and baseline converter (2026-11-12)

### What exists

- **`ingest/parse.py`** (pypdfium2). Per page:
  - glyphs: unicode, font name / flags / weight, size, fill colour + alpha, tight bbox,
    origin, baseline, rotation from the char matrix, render mode, unmapped-unicode flag;
  - vector paths: M/L/C/Z segments, fill / stroke colour + alpha, line width, dash, cap,
    join, an even-odd note and a clip bbox note. Object and form-XObject matrices are
    resolved to page space;
  - images: JPEG passthrough for plain DCT, PNG otherwise, placement matrix → frame +
    rotation + flip, crop from clip, soft mask;
  - annotations: links with a URI or destination page, other subtypes with contents;
  - metadata: producer, creator, page sizes, PDF version.

  Everything is in display space: points, top-left origin, cropbox and `/Rotate`
  applied. A page that throws is logged and comes back marked `triage.fallback_image`.
- **pdfplumber cross-check** on the first unrotated page of every document. The result
  is stored in the IR metadata (`pdf2pptx:crosscheck`) and logged on disagreement.
  Corpus: 31/31 checked pages have text ratio 1.000 and a max origin difference
  ≤ 0.001 pt.
- **`ingest/render.py`**: 150 / 300 DPI renders, cached as
  `work/<sha>/renders/pNNNN_<dpi>.png`. They're made lazily (only fallback pages use
  them for now) and never held in RAM.
- **`ingest/triage.py`**: rules on the IR. Corpus result: 58 born_digital, 3 scanned
  (s01 ×2, s02), 1 outlined_text (d28: 60 letter-sized curved sub-paths vs 50 glyphs).
  On the bigram check, real text scores 0.77–1.00 and Caesar-shifted text scores 0.39.
- **`build/writer.py`** + `build/lines.py` + `build/xml_ext.py`:
  - slide size from the most common page size; other pages are scaled and centred, and
    PowerPoint's 1–56 in limit is enforced;
  - one zero-margin, no-wrap, no-autofit text box per line;
  - images with crop, rotation, flips and alpha;
  - every path as a `custGeom` freeform (alpha, custom dash, caps, joins);
  - z-order kept; URI links on runs, internal links as slide jumps;
  - a page that fails to build is cleared and replaced by its 300-DPI render;
  - byte-deterministic output (zip entries and core properties get fixed timestamps).
- **`fonts/index.py`**: installed fonts from the registry, PDF-name → Windows-family
  resolution, and `winAscent/winDescent` and per-glyph advances via fontTools.
- **`pipeline.py`**: ingest → triage → analyze (identity) → build. Each stage is cached
  as `work/<sha>/ir_<stage>.json`, plus `ir_build.json` (the final IR, with fallbacks
  marked). The cache key is the chain of stage versions (`cache.json`).
- **CLI**:
  - `pdf2pptx convert in.pdf -o out.pptx [--mode] [--workers N] [--no-cache]`;
  - `pdf2pptx convert --corpus [--only …] [--no-check]`. The corpus run uses a process
    pool over files (default 6 workers); a single large PDF is parsed in parallel page
    chunks;
  - `score … --as-baseline` writes `work/scores/baseline.json`;
  - `score --corpus --against out` now fills `seconds_per_page` from
    `out/corpus/<name>.convert.json`.

### Measured facts that later milestones rely on

- **PowerPoint's first baseline** in a zero-inset text box (single spacing) sits
  `1.2 × winAscent / (winAscent + winDescent) × size` below the box top. This held for
  Calibri, Arial, Times New Roman, Georgia, Consolas, Verdana, Segoe UI and Courier New
  (sizes 20 and 40) to within about 0.3% of the size. No single other font metric
  (hhea, typo, win ascent alone) fits.
- **One box per line can still keep word positions.** Each inter-word space gets a
  character spacing (`spc`) computed from installed-font advances, so the next word
  starts at its PDF origin. Without this, bullet decks had pos_err of 5–7 pt because
  the bullet → text gap collapsed into one space. With it, they're at about 0.7 pt.
- **pdfium hides pattern (gradient) fills.** `FPDFPageObj_GetFillColor` returns opaque
  white and `GetRenderedFillPattern` returns nothing. The parser now renders each
  opaque-white top-level path on its own (`FPDFPageObj_SetIsActive`, transparent
  background, 0.5 px/pt, about 5 ms). If that render isn't uniform white, the path is
  replaced by an isolated 144-DPI render with alpha (`raster:pattern_fill`). This took
  d12 from ssim 0.938 / region 0.669 to 0.997 / 0.987.
- **pdfium also misses soft masks.** `FPDFPageObj_HasTransparency` is 0 for images with
  an `/SMask`, so the alpha comes from `FPDFImageObj_GetRenderedBitmap` (about 1 px/pt,
  upscaled). This matters because PowerPoint exports shadows as black images with alpha.
- PowerPoint's PDF export puts every object under the page clip. It writes rotated text
  at −30° and 8° as one path with about 2 000 segments.
- reportlab's `setPageRotation(90)` swaps the mediabox itself (792×612 + /Rotate 90).

### Checks

| Check | Result |
|---|---|
| `pytest -q` | **61 passed**. New tests cover:<br>• parse: positions within 0.5 pt, exact colour, rotated line, page rotation checked against pdfium's own render<br>• triage: unit tests and the whole corpus<br>• pipeline: end-to-end, cache, byte determinism, bad-page fallback<br>• corpus outputs open in PowerPoint |
| `ruff check .` / `mypy` | clean / clean |
| `convert --corpus --no-cache` | **33/33 files, 62/62 pages, 0 fallbacks, 12.2 s wall** (6 workers) |
| PowerPoint COM `Presentations.Open` + slide count | **33/33 open, slide counts match** |

### BASELINE scores (`work/scores/baseline.json`; `score --corpus --against out --label baseline --as-baseline`)

Median over slides; the mean is in parentheses for ssim and text_sim.

| category | slides | ssim | reg_min | text_sim | w_recall | pos_err | lb_match | real_txt | native | boxes | s/page | composite |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| ALL | 62 | 0.989 (0.987) | 0.957 | 1.000 (0.951) | 1.000 | 0.508 | 1.000 | 1.000 | 0.000 | 5.000 | 0.163 | 0.896 |
| report | 3 | 0.958 (0.947) | 0.865 | 1.000 (0.996) | 1.000 | 0.575 | 1.000 | 1.000 | 0.000 | 46.000 | 0.484 | 0.964 |
| scanned | 3 | 0.984 (0.981) | 0.977 | 0.000 (0.000) | 0.000 | n/a | 0.000 | 0.000 | n/a | 0.000 | 0.173 | 0.437 |
| synthetic_deck | 56 | 0.992 (0.989) | 0.957 | 1.000 (1.000) | 1.000 | 0.506 | 1.000 | 1.000 | 0.000 | 5.000 | 0.150 | 0.896 |

Reference: the ground-truth decks (M1) score ssim 0.994, reg_min 0.974, pos_err 0.38,
boxes 2, native 1.000 and composite 0.997 (synthetic medians).

<details><summary>By file (median over slides)</summary>

| file | cat | slides | ssim | reg_min | text_sim | w_recall | pos_err | lb_match | real_txt | native | boxes | s/page | composite |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| d01_title_bullets | synthetic_deck | 3 | 0.975 | 0.836 | 1.000 | 1.000 | 0.721 | 1.000 | 1.000 | 0.000 | 4.000 | 0.220 | 0.890 |
| d02_multilevel_bullets | synthetic_deck | 2 | 0.983 | 0.922 | 1.000 | 1.000 | 0.514 | 1.000 | 1.000 | 0.000 | 8.500 | 0.323 | 0.893 |
| d03_two_columns | synthetic_deck | 2 | 0.981 | 0.909 | 1.000 | 1.000 | 0.452 | 1.000 | 1.000 | 0.000 | 12.000 | 0.435 | 0.892 |
| d04_three_columns | synthetic_deck | 1 | 0.974 | 0.899 | 1.000 | 1.000 | 0.241 | 1.000 | 1.000 | 0.000 | 22.000 | 0.765 | 0.889 |
| d05_mixed_runs | synthetic_deck | 1 | 0.983 | 0.924 | 0.995 | 0.917 | 0.508 | 1.000 | 0.979 | 0.000 | 9.000 | 0.614 | 0.888 |
| d06_italian_text | synthetic_deck | 2 | 0.964 | 0.868 | 1.000 | 1.000 | 0.596 | 1.000 | 1.000 | 0.000 | 5.500 | 0.372 | 0.886 |
| d07_table_simple | synthetic_deck | 1 | 0.996 | 0.991 | 1.000 | 1.000 | 0.233 | 1.000 | 1.000 | 0.000 | 26.000 | 0.244 | 0.899 |
| d08_table_merged | synthetic_deck | 1 | 0.997 | 0.987 | 1.000 | 1.000 | 0.224 | 1.000 | 1.000 | 0.000 | 28.000 | 0.246 | 0.899 |
| d09_shapes_basic | synthetic_deck | 1 | 0.997 | 0.987 | 1.000 | 1.000 | 0.234 | 1.000 | 1.000 | 0.000 | 5.000 | 0.444 | 0.899 |
| d10_arrows_chevrons | synthetic_deck | 1 | 0.998 | 0.993 | 1.000 | 1.000 | 0.258 | 1.000 | 1.000 | 0.000 | 5.000 | 0.436 | 0.899 |
| d11_connectors | synthetic_deck | 1 | 0.995 | 0.982 | 1.000 | 1.000 | 0.343 | 1.000 | 1.000 | 0.000 | 5.000 | 0.364 | 0.898 |
| d12_gradients_alpha | synthetic_deck | 1 | 0.997 | 0.987 | 1.000 | 1.000 | 0.420 | 1.000 | 1.000 | 0.000 | 2.000 | 0.349 | 0.899 |
| d13_images | synthetic_deck | 2 | 0.994 | 0.979 | 1.000 | 1.000 | 0.257 | 1.000 | 1.000 | 0.000 | 2.000 | 0.150 | 0.898 |
| d14_chart_bar | synthetic_deck | 2 | 0.991 | 0.980 | 1.000 | 1.000 | 0.252 | 1.000 | 1.000 | 0.000 | 16.000 | 0.124 | 0.896 |
| d15_chart_line | synthetic_deck | 1 | 0.984 | 0.972 | 1.000 | 1.000 | 0.252 | 1.000 | 1.000 | 0.000 | 20.000 | 0.235 | 0.894 |
| d16_chart_pie | synthetic_deck | 1 | 0.995 | 0.978 | 1.000 | 1.000 | 0.281 | 1.000 | 1.000 | 0.000 | 9.000 | 0.146 | 0.898 |
| d17_footer_logo_number | synthetic_deck | 5 | 0.989 | 0.910 | 1.000 | 1.000 | 0.715 | 1.000 | 1.000 | 0.000 | 5.000 | 0.100 | 0.895 |
| d18_buildup | synthetic_deck | 5 | 0.986 | 0.880 | 1.000 | 1.000 | 0.717 | 1.000 | 1.000 | 0.000 | 4.000 | 0.068 | 0.894 |
| d19_hyperlinks | synthetic_deck | 2 | 0.990 | 0.958 | 1.000 | 1.000 | 0.374 | 1.000 | 1.000 | 0.000 | 6.000 | 0.171 | 0.896 |
| d20_speaker_notes | synthetic_deck | 3 | 0.995 | 0.938 | 1.000 | 1.000 | 0.720 | 1.000 | 1.000 | 0.000 | 3.000 | 0.071 | 0.898 |
| d21_sections | synthetic_deck | 6 | 0.999 | 0.995 | 1.000 | 1.000 | 0.620 | 1.000 | 1.000 | 0.000 | 1.000 | 0.040 | 0.900 |
| d22_aspect_4x3 | synthetic_deck | 3 | 0.996 | 0.977 | 1.000 | 1.000 | 0.610 | 1.000 | 1.000 | 0.000 | 3.000 | 0.135 | 0.898 |
| d23_long_paragraphs | synthetic_deck | 2 | 0.979 | 0.864 | 1.000 | 1.000 | 0.417 | 1.000 | 1.000 | 0.000 | 6.000 | 0.242 | 0.891 |
| d24_font_families | synthetic_deck | 1 | 0.982 | 0.933 | 1.000 | 1.000 | 0.331 | 1.000 | 1.000 | 0.000 | 12.000 | 0.392 | 0.893 |
| d25_dark_background | synthetic_deck | 2 | 0.996 | 0.981 | 1.000 | 1.000 | 0.503 | 1.000 | 1.000 | 0.000 | 6.000 | 0.154 | 0.898 |
| d26_kitchen_sink | synthetic_deck | 1 | 0.992 | 0.974 | 1.000 | 1.000 | 0.249 | 1.000 | 1.000 | 0.000 | 22.000 | 0.331 | 0.897 |
| d27_italian_report_4x3 | synthetic_deck | 2 | 0.988 | 0.958 | 1.000 | 1.000 | 0.314 | 1.000 | 1.000 | 0.000 | 15.000 | 0.148 | 0.895 |
| d28_rotated_text | synthetic_deck | 1 | 0.997 | 0.984 | 1.000 | 0.273 | 0.262 | 0.111 | 1.000 | 0.000 | 3.000 | 0.555 | 0.810 |
| r01_report_two_column | report | 1 | 0.920 | 0.725 | 1.000 | 1.000 | 0.575 | 1.000 | 1.000 | n/a | 70.000 | 0.967 | 0.964 |
| r02_report_tables_it | report | 1 | 0.964 | 0.893 | 0.989 | 0.945 | 0.445 | 0.848 | 0.967 | 0.000 | 46.000 | 0.484 | 0.862 |
| r03_report_figure | report | 1 | 0.958 | 0.865 | 1.000 | 1.000 | 0.858 | 1.000 | 1.000 | n/a | 12.000 | 0.475 | 0.981 |
| s01_scanned_deck_it | scanned | 2 | 0.984 | 0.978 | 0.000 | 0.000 | n/a | 0.000 | 0.000 | n/a | 0.000 | 0.173 | 0.437 |
| s02_scanned_report | scanned | 1 | 0.975 | 0.964 | 0.000 | 0.000 | n/a | 0.000 | 0.000 | n/a | 0.000 | 0.222 | 0.433 |

</details>

This is what a baseline should look like:

- Every text-bearing slide has real, editable text (real_txt 1.000).
- Positions are within about 0.5 pt, and the slides look right (ssim median 0.989).

It isn't usefully editable yet:

- There are 5 boxes per slide (median; mean 8.6, max 70 on r01), against 2 in the
  truth decks.
- There are no native shapes; everything is a freeform or a picture.
- The composite median of 0.896 is held down mainly by `native_shape_ratio = 0`.

### The 5 worst slides (by composite)

| # | slide | composite | why |
|---|---|---|---|
| 1 | s02_scanned_report p0 | 0.434 | Scanned, so the page image is placed and there's no text at all (OCR is M8). It looks right (ssim 0.975). |
| 2 | s01_scanned_deck_it p1 | 0.437 | Same: scanned, image only. |
| 3 | s01_scanned_deck_it p0 | 0.437 | Same: scanned, image only. |
| 4 | d28_rotated_text p0 | 0.810 | The −30° and 8° text exists only as outlines in the PDF, so it comes out as freeforms (triage: outlined_text; OCR or re-typing is M8). word_recall 0.27 and lb_match 0.11 are partly a scoring gap: the scoring word extractor doesn't group rotated words (M1 known gap), so the vertical text box gets matched character by character. |
| 5 | r02_report_tables_it p0 | 0.862 | Whitespace-table cells closer than 0.8 em merge into one line box (e.g. "1 1 0"). Character spacing keeps their positions, but pdfplumber splits them into separate lines, which gives lb_match 0.85 and word_recall 0.95. Table detection is M6. |

The next lowest by ssim:

- **r01 (0.920, 70 boxes).** The layout is correct. The difference is font rendering:
  pdfium draws the non-embedded Helvetica with its substitute and PowerPoint uses Arial,
  which shows on dense 9 pt text.
- **d06 (0.953).** The bullet / spacing differences of single-line boxes.

### Known gaps

- Analyze is the identity, so text is one box per line. M4 builds paragraphs, lists and
  reading order. There are no native shapes, tables or charts yet (M6/M9).
- Shading objects (`sh` operator) are skipped and logged, not rasterised. Only top-level
  pattern-filled paths are checked (at most 200 per page); paths inside form XObjects
  aren't.
- Soft-mask alpha is low-resolution (about 1 px/pt), and it's dropped for rotated or
  flipped images.
- The writer skips invisible (render mode 3) text, so OCR'd scans keep only the image.
- Text alpha, blend modes and the even-odd fill rule aren't written; the IR notes them.
- On PowerPoint exports, pdfium reports font sizes about 0.1% too large (32.04 vs 32).
  They're left as is.

### Next

Milestone 3: the `verify/ppt_com.py` worker process (replacing `_com_temp.py`), the fast
check engine and the HTML debug report.

---

## Milestone 3: PowerPoint check engine and debug report (2026-11-12)

### What exists

- **`verify/ppt_com.py`**, the only module that talks to PowerPoint.
  - **Worker.** PowerPoint runs in a worker *process* (`multiprocessing` spawn), whose
    main thread is the single COM STA thread. The client, `PowerPointEngine`, is
    thread-safe and serialises calls.
  - **Operations** (over a queue): `open(pptx)` → `PresInfo` (slides, size),
    `render(pres, slides, dpi, dir)` → PNGs, `text_layout(pres, slide)`, `export_pdf`,
    `close`, `ping`. Slides are addressed by 0-based index (= IR page index).
    `text_layout` returns, per text shape (a table cell counts as one shape), lines
    with text, raw `BoundLeft/Top/Width/Height` and trimmed word boxes.
  - **Robustness.**
    - A heartbeat is updated between COM operations, and every call has a timeout.
    - The engine restarts on worker death, a hang, or an RPC disconnect: it kills the
      worker, releases its PowerPoint, starts a new worker, re-opens the open
      presentations and retries the call.
    - A bad file is a `PowerPointError`, not a restart.
    - The worker exits if its client dies.
  - **PowerPoint settings.** `DispatchEx`, read-only `WithWindow=False` presentations,
    `DisplayAlerts` off.
  - **Process hygiene.**
    - Every POWERPNT.EXE we start is registered in `%TEMP%\pdf2pptx\ppt_owned` with its
      creation time, which guards against pid reuse.
    - PowerPoint is single-instance, so a second engine *attaches* to the first one's
      process; it then registers as co-owner. An instance is quit or killed only when
      no live co-owner is left and it holds no foreign presentations.
    - Orphans (all owners dead) are reaped when the next engine starts.
    - A PowerPoint the user started is never quit or killed.
  - **Lifetime.** `ppt_com.shared()` gives one long-lived engine per process; an atexit
    hook stops it before multiprocessing terminates daemon children. If PowerPoint is
    missing, the error is remembered, so it costs one start attempt per process.
  - **Speed.** Unchanged files reuse the open presentation, and only the requested
    slides are rendered.
- **`verify/_com_temp.py` is gone.** `score`, `pipeline` (the corpus open check),
  `corpus build` (PDF export), `doctor` and the tests all use `ppt_com`.
- **`verify/diff.py`**: aligned per-slide comparison.
  - **Alignment.** The PowerPoint render is cropped to the page rectangle that the
    builder's uniform scale + centring puts on the slide, then resampled to the PDF
    render's size. When sizes are equal this is the identity.
  - **Scores.** SSIM map, 4×4 region scores, `region_ssim_min`.
  - **Mismatch regions.** Low local SSIM, smoothed and closed, then connected
    components, each with page bbox (pt), mean SSIM, severity, and the overlapping IR
    element ids. Ids are ranked by overlap² / (region area × element area), so the
    specific element comes before a page background.
  - **Heatmap PNG.** The page faded to gray, 1 − SSIM in colour, regions outlined,
    4×4 grid.
- **`metrics.py`.** `words_from_layout()` turns `text_layout()` lines into the PPTX
  side of `pos_err_pt` / `linebreak_match` / `word_recall`. `map_lines()` maps slide
  pt back to page pt when page and slide sizes differ (letterboxed pages).
- **`score.py`.**
  - It renders and lays out each slide through the engine, records
    `com_seconds_per_slide` and keeps the 4×4 region grid plus the top-10 mismatch
    regions per slide in the JSON.
  - Corpus summaries print `PowerPoint render + text layout per slide: median …`.
  - Without PowerPoint, its scores are null and `powerpoint_error` says why.
- **`report/debug_html.py`** + **`pdf2pptx report <pdf> [<pdf> …] [-o] [--pptx]
  [--corpus/--only]`.**
  - Converts (cached), scores and writes one self-contained HTML per PDF
    (`out/report/<stem>.html`, no external assets).
  - Top of the page: document status and a sortable per-page score table (click a
    header; n/a sorts last; a page link jumps to it).
  - Per page: triage chip + reason, the page's score row, and three panels: PDF
    (pdfium) | PPTX (PowerPoint) | diff heatmap.
  - SVG overlays on every panel:
    - IR boxes: the elements, the per-line text boxes the builder writes, and
      annotations. They toggle by kind and by source and are coloured by either;
      rotated text blocks are drawn rotated.
    - PowerPoint's own line boxes, mapped back to page pt.
    - Mismatch regions.
  - Hovering a box shows its JSON (long lists cut).
  - Several PDFs also get an `index.html` with a sortable per-document table.
  - Without PowerPoint, a banner says why and which scores are missing; the PDF panel
    and overlays still work.
- **`log.py`.** The handler writes to the *current* `sys.stderr`. Before, logging from
  atexit hooks under pytest hit a closed capture stream.

### Measured facts that later milestones rely on

- **COM cost is in object creation, not in PowerPoint's layout.**
  - A property read on an existing object costs ~0.25 ms.
  - Each *new* object costs 2–10 ms because of pywin32's wrapping. Dynamic dispatch
    does a type-info lookup per object; `TextRange2` / `TextFrame2` live in the Office
    type library, so even PowerPoint makepy doesn't cover them. Makepy classes for
    both libraries still do one cross-process `QueryInterface` per object.
  - Raw `IDispatch.Invoke` with DISPIDs cached per (interface, member) costs one round
    trip per member.
  - Together with a python-pptx plan (only text-bearing shapes are visited) and
    trimming only words with trailing whitespace, text layout went from 9.3 s to
    0.9 s on r01 (70 lines, 376 words) and from 0.5–1.5 s to 0.05–0.2 s on typical
    slides.
- **Render.** `Slide.Export` at 150 DPI takes 0.07–0.14 s per slide. Opening a
  presentation takes 0.06–0.35 s, and starting the worker + PowerPoint about 2.5–4 s.
- **Killing PowerPoint under a live worker** shows up as RPC error `0x800706BE`
  ("La chiamata di procedura remota non è riuscita"). The worker then reports
  `fatal`, and the client restarts it.
- **Headless Edge** (`msedge --headless --screenshot`) exits with code 21 on this PC,
  probably because of corporate policy. Report checks therefore re-draw the HTML's
  own SVG rects onto its inlined images (see below).

### Checks

| Check | Result |
|---|---|
| Stress: 200 `render` calls in a row over 8 corpus decks; worker killed after call 70, its POWERPNT.EXE killed after call 140 | **200/200 complete, 2 restarts, 21.6 s.** The dead workers' PowerPoints are gone, and nothing is left after `stop()` (`test_stress_recovers_from_killed_worker`) |
| Hang: a call that sleeps 30 s, with a 2 s timeout | Times out, the worker and its PowerPoint are killed, and a fresh worker answers `ping` (`test_timeout_restarts_worker`) |
| Speed: render + text layout per slide, corpus (`score --corpus --against out`) | **median 0.144 s**, mean 0.187 s, max 0.995 s (r01, 376 words, includes the first-call shape plan); 60/62 slides ≤ 0.5 s. Truth decks: median 0.125 s, max 0.504 s |
| `text_layout` line counts on 5 slides with known truth: 3 paragraphs; 2 soft breaks; a 2×2 table with a 2-paragraph cell; a group of 2 boxes; 10 Consolas words wrapping 2 per line | **3, 3, 5, 2, 5: all correct.** Origins within 1.5 pt, and the Consolas word width is 44 ± 2 pt (`test_text_layout_line_counts`) |
| `pdf2pptx report` on d01_title_bullets (3 pages), r02_report_tables_it, d26_kitchen_sink | HTML written in 6.8 / 2.4 / 1.8 s (1.1 MB / 0.6 MB / 0.4 MB), plus `index.html` |
| Overlay alignment (d26, r02, d01) | See below |
| No leftover PowerPoint | No POWERPNT.EXE and an empty registry after the test suite (session fixture asserts it), `convert --corpus`, `score --corpus` ×3, `report` ×3, `doctor` and single-file `score` |
| No PowerPoint (simulated: `ppt_com.shared` raises) | `report` still writes the HTML, with the banner + reason (`test_report_without_powerpoint`); `score_pair` gives null ssim / region / pos_err / lb_match with the reason, while text scores and composite stay (`test_score_without_powerpoint_is_null_with_reason`) |
| `pytest -q` | **74 passed in ~90 s** (M2: 61 in ~3 min). The `0x80010108` faulthandler message is gone, since COM is no longer in the test process |
| `ruff check .` / `mypy` | clean / clean |
| `convert --corpus --no-cache` | 33/33 files, 33/33 open in PowerPoint, 8.7 s |

**How the overlays were checked.** I parsed each generated HTML and drew its `<rect>`s
onto its own inlined images with the browser's mapping (`viewBox` in pt,
`preserveAspectRatio="none"` over the image box), then looked at the results.

- **d26.** The blue IR text boxes frame every PDF text line ("Dashboard", the three
  bullets, the axis ticks 0–25 and Q1–Q4, the table cells, "On track") with no visible
  offset. The orange shape and image boxes sit on the chart frame, the bars, the
  table rules, the rounded rectangle, the circle, the logo and the gradient picture.
- **d26, PPTX panel.** PowerPoint's own line boxes (cyan) coincide with the IR boxes.
  The one visible difference is real, not an overlay error: the "!" inside the
  circle is drawn a few pt left of its box by PowerPoint.
- **Ink share.** 100% of dark pixels fall inside text boxes on d01's three pages.
  r02 has 91.8%; the remainder is the table rules.
- **Heatmap panel.** The 4×4 grid, the colour and the boxes line up on r02. Its hot
  spots are the known Helvetica → Arial font substitution from M2.

### Scores: unchanged, and the new full baseline

Every metric on all 62 slides is **identical** to the M2 baseline (max abs difference
0.0 for ssim, region_ssim_min, text_sim, word_recall, pos_err_pt, linebreak_match,
real_text_ratio, native_shape_ratio, boxes_per_slide and composite). Rescoring the
truth decks also reproduces the M1 truth scores exactly (58 slides). The new data
per slide (4×4 regions, mismatch regions, `com_seconds_per_slide`) is extra.
`seconds_per_page` differs only because the corpus was converted again (wall time:
0.163 → 0.137 median).

`work/scores/baseline.json` was rewritten from this run (`score --corpus --against out
--label baseline --as-baseline`):

| category | slides | ssim | reg_min | text_sim | w_recall | pos_err | lb_match | real_txt | native | boxes | s/page | composite |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| ALL | 62 | 0.989 (0.987) | 0.957 | 1.000 (0.951) | 1.000 | 0.508 | 1.000 | 1.000 | 0.000 | 5.000 | 0.137 | 0.896 |
| report | 3 | 0.958 (0.947) | 0.865 | 1.000 (0.996) | 1.000 | 0.575 | 1.000 | 1.000 | 0.000 | 46.000 | 0.318 | 0.964 |
| scanned | 3 | 0.984 (0.981) | 0.977 | 0.000 (0.000) | 0.000 | n/a | 0.000 | 0.000 | n/a | 0.000 | 0.129 | 0.437 |
| synthetic_deck | 56 | 0.992 (0.989) | 0.957 | 1.000 (1.000) | 1.000 | 0.506 | 1.000 | 1.000 | 0.000 | 5.000 | 0.105 | 0.896 |

PowerPoint render + text layout per slide: median 0.144 s, mean 0.187 s, max 0.995 s
(n = 62). The by-file table equals M2's apart from `s/page`; it is in
`work/scores/baseline/summary.md`.

### Known gaps

- The report re-renders the PDF pages once more for display (scoring already rendered
  them). That costs ~0.1 s per page. It is simple to share later.
- Mismatch regions use a fixed local-SSIM threshold (0.6). On font-substituted text
  (r01/r02) almost every line becomes a region, which is correct but noisy. The tuner
  (M5) should rank by severity and look at the region's element ids.
- `text_layout` skips charts (as M1 did), so chart tick labels are still PDF-only
  text.
- Report HTML size grows ~350 KB per page (three JPEG panels). Fine for decks; a
  100-page report would want a lower DPI or lazy loading.
- The overlay check is done by re-drawing from the HTML, not by a browser screenshot,
  because headless Edge is blocked here.

### Next

Milestone 4: text understanding and fonts (paragraphs, lists, reading order). The
target is ≥ 5× fewer text boxes per slide and word_recall ≥ 0.99 on born-digital
pages, measured against this baseline.

**Starting points for Prompt 4** (checked against the code at `85441ce`):

- **Hook.** `pipeline.analyze()` is the identity (`STAGE_VERSIONS["analyze"] =
  "identity-2.0"`). Bump that version when it starts emitting TextBlocks, or stale
  `ir_analyze.json` caches are reused.
- **Writer.** `build/writer._build_page` already writes `TextBlock` elements when a
  page has any. It falls back to `build/lines.line_blocks()` (one box per line) only
  when it has none. Keep that function: it becomes the "baseline line-per-box" path
  that Prompt 4 must keep behind a flag. `_add_text` applies the block's `wrap`,
  `margins`, `rotation` and the runs' font, size, colour, bold, italic, `spc` and
  links. It writes no paragraph properties yet (alignment, indents, spacing,
  bullets, level).
- **Fonts.** `fonts/index.py` already has the installed-font registry,
  `resolve_family()` (PDF name → Windows family), `advance_em()` (per-glyph advances)
  and `first_baseline_offset()` (the measured `1.2 × winAscent/(winAscent+winDescent)`
  rule). Prompt 4 adds `match.py` (width-profile matching) and `embed.py` next to it.
- **Truth for the checks.** Each `truth.json` slide has `paragraphs`:
  `{text, level, bullet: none|inherited|char|autonum, role, fonts, has_field}`.
  That is enough for "paragraph count and bullet levels match on ≥ 95% of slides"
  and "matched family = truth family on ≥ 95% of runs".
- **Measuring.**
  1. `convert --corpus --no-cache`, then `score --corpus --against out --label m4`.
  2. Compare with `work/scores/baseline.json`.
  3. `pdf2pptx report <pdf>` shows the new blocks (overlay kind `text_block`, source
     `rule`) and PowerPoint's own line boxes, so wrap and line-break errors show up
     as cyan vs blue mismatches.
  4. The prompt asks for low-confidence blocks to be listed in the report: add that
     to `report/debug_html.py` (`_page_overlays` already has every element's
     `confidence`).
- **Speed headroom.** Fewer boxes means fewer COM shape visits, so scoring should get
  faster. A wrapped multi-line box costs one `Lines.Item` per line plus the words.

---

## Milestone 4: text understanding and fonts (2026-11-12)

### What exists

- **`analyze/text_group.py`**: glyphs → words → lines → paragraphs → TextBlocks. Every
  threshold is in em of the local font size or relative to page/block statistics.
  - **Local frames.** Glyphs are grouped by rotation and projected on (along, across),
    so rotated text is grouped exactly like horizontal text. The frame goes back to
    the page as an unrotated box + clockwise rotation.
  - **Lines.** Content order first, with a sub/superscript tolerance (≤ 0.85× size
    within 0.6 em of the baseline). A merge pass joins same-baseline fragments a PDF
    drew out of order. Invisible text (render mode 3/7 or a fully transparent fill)
    is dropped.
  - **Splits.** Words split at space glyphs or gaps > 0.15 em. A word gap wider than
    `max(0.8 em, 3 × page median gap)` splits the line (columns, cells), except in
    lines whose gaps are uniformly wide (justified).
  - **Blocks.** A line joins the block whose last line is just above it: horizontal
    overlap, size ratio ≤ 1.6, baseline distance ≤ 2 block pitches (≤ 1.5 × size when
    sizes differ by > 1.3×). A line touching two blocks at the same height starts its
    own (it spans columns). Short label-less items more than 1.3 lines apart (axis
    ticks, table cells) stay separate. An italic/upright change together with a size
    change starts a new block (captions).
  - **Paragraphs.** A new paragraph starts at a label, at extra space (> 1.3 × block
    pitch), at a size change, at an indent change, after a short line (the next line's
    first word would have fitted), at stacked single words, or, as a low-confidence
    (0.65) tie-breaker, when the widest line ends a sentence before a short line and
    nothing else in the block shows the wrap width. Centred/right blocks use
    width-based tests. End-of-line hyphenation joins without a space and is recorded
    (`hyphen:` note).
  - **Properties.** Alignment from line edges (justify needs equal edges and a stretch
    > max(0.25 em, 1.5 pt); single-line paragraphs vote on the edge they share, or
    are centred in an enclosing filled shape / on the page). Indents from label and
    text positions (marL / indent). Line spacing: single when within 2% of 1.2 × size,
    else the exact pitch in points. Space after from the remaining gap.
  - **Frame.** The width comes from the wrap constraints: at least every line's natural
    width, and less than (line + space + next word) for every wrapped line. Justified
    blocks use the column width (or just enough for a line the PDF squeezed). The
    vertical position walks down the block predicting PowerPoint's baselines (with its
    rounding, see below) and puts the rounding error into the next paragraph's space.
  - **Runs.** One per change of family, bold, italic, size, colour, script shift or
    link. Sub/superscripts get `baseline_shift` and 1.5× their glyph size (PowerPoint
    draws them at ⅔). In a single-line paragraph an oversized gap (legend swatch, tab)
    keeps its width as character spacing on the space.
  - **z.** The first glyph's z (keeps the PDF's drawing and text order), or the last
    glyph's when a filled graphic drawn between them overlaps the block.
- **`analyze/lists.py`**: bullet glyphs (strong `•▪►✓…`; weak `– — - · *` only as a
  separate word followed by a gap), number labels (`1.` `2)` `(3)` `a.` `iv.` `B)`;
  `i.` resolved as roman or alpha from its neighbours), picture bullets (a small image
  left of the line, centred on the x-height), indent levels (clustered label
  positions) and numbering runs (PowerPoint continues a run while `type` and `startAt`
  are equal; a broken sequence gets its own `startAt`). Bullet glyphs leave the text
  and become real `buChar` / `buAutoNum` / `buBlip` bullets with font, colour and size.
- **`analyze/reading_order.py`**: recursive XY-cut. A vertical gap with substantial
  content on both sides is a column split and wins over a horizontal cut. It sets
  `TextBlock.reading_order`, which the writer uses to break z ties.
- **`analyze/roles.py`**: title, subtitle, body, footer, header, page_number and caption
  from position, size and repetition (text with digits masked + position, on at least
  half of the pages). A page number equal to the slide number becomes a real
  slide-number field. Every role has a `role_confidence`.
- **`analyze/text.py`**: font matching → embedding candidates → grouping per page (a
  page that fails keeps the one-box-per-line fallback, logged) → roles → reading
  order. `STAGE_VERSIONS["analyze"] = "text-4.5"`.
- **`fonts/index.py`**: an index of every installed face in `C:\Windows\Fonts` and
  `%LOCALAPPDATA%\Microsoft\Windows\Fonts` (354 faces here, ~2 s cold): family (nameID
  1, the name PowerPoint shows), typographic family, style, weight, width, italic,
  monospace, serif (PANOSE), fsType, win metrics and a width profile (ASCII + common
  Latin-1). It is cached in `%LOCALAPPDATA%\pdf2pptx\cache\font_index.json` and
  re-read per changed file. `convert --corpus` builds it once before starting workers.
- **`fonts/match.py`**:
  - Name first: the subset prefix and PS suffixes are stripped, camel case is split
    and weight words are kept (`Calibri-Light` → Calibri Light, `SegoeUI-Semibold` →
    Segoe UI Semibold). Base-14 names go through aliases (Helvetica → Arial, …). The
    widths are still checked.
  - Otherwise the width profile of the characters actually used (`sum|Δw| / sum w`
    plus style penalties); otherwise the serif / fixed-pitch flags.
  - Bold/italic come from the name, the flags, ForceBold and the stem weight; a family
    that is itself heavy is not bolded again.
  - It records `matched_family`, `match_score` and `match_source` on each FontRef.
- **`fonts/embed.py`** + ingest: embedded font programs are extracted to
  `work/<sha>/fonts/` and their OS/2 fsType is read. Fonts that are embedded, not
  installed and licensed for embedding are listed in
  `metadata["pdf2pptx:embed_candidates"]`. **Writing them into the PPTX is a gap**
  (see Known gaps).
- **Ingest** (`PARSER_VERSION = "2.1"`): per font, the PDF advance width (em) of every
  character used (`FPDFFont_GetGlyphWidth` takes a Unicode code point), plus the
  program path and fsType. IR: `FontRef` gains `widths`, `program`, `fs_type` and
  `match_source` (additive; the IR version is unchanged).
- **Writer**:
  - One wrapping, zero-inset box per TextBlock with `algn`, `marL`/`indent`, `lvl`,
    line spacing, space after, bullets (char / autonum / picture), sub/superscript
    `baseline` and slide-number fields.
  - Autofit is off except in `editable`/`deck` (normAutofit, and end-of-line hyphens
    removed). Images used as picture bullets are not placed.
  - **`--text-layout lines`** (config `text_layout`) keeps the M2 path. Rescoring it
    reproduces the baseline with max |Δ| = 0.0 on every metric of every slide.
- **`verify/structure.py`** + **`pdf2pptx structure`**: analyzed paragraphs vs the
  synthetic truth (count, levels, bullet present, font family per run); table and
  chart frames are left out. It writes `work/scores/structure.json`.
- **Debug report**: each page lists its low-confidence text blocks (< 0.6) with the
  reasons (paragraph split rule, role, notes such as `wrap_conflict`, `vdrift`).
- **Scorer fix** (`ppt_com._collect_text_frame`): a `SpaceAfter` in *lines*
  (`LineRuleAfter`) is now converted (lines × 1.2 × font size) before it is taken off
  the line bottom. Before, it was ignored, which shifted every word of a paragraph
  with percentage spacing (pos_err 8 pt on d01). Point spacing is unchanged, so the
  baseline and truth scores do not move.

### Measured facts that later milestones rely on

- **PowerPoint rounds `spcPts` to whole points** (line spacing and space before/after:
  13.5 → 14, 10.2 → 10, 32.4 → 32, 6.77 → 7) **and `spcPct` to whole percent**
  (112.5% → 113%), both half up. This was measured on 96 boxes (Arial/Calibri, 10/16/24
  pt, ratios 0.85–1.6) exported to PDF and parsed. The writer picks the encoding that
  lands closer (`build/units.ppt_spacing`). r01's drift went from +0.54 pt per line to
  ≤ 0.37 pt over the whole column.
- **First baseline with non-single spacing** (ratio r of single): for r ≥ 1 it sits at
  0.75 × the line height, font-independent (±0.005); for r < 1 it is the single offset
  minus (1 − r) × 1.2 × size. Single spacing keeps the M2 rule. Space after in percent
  = r × 1.2 × size.
- **Super/subscript** runs are drawn at ⅔ of the run size (0.6655–0.6666) and raised by
  `baseline`% of the *nominal* size.
- **A bullet's font does not change the line metrics**: the line is laid out with the
  text font. An Arial bullet on a Calibri line moved our first baseline by 1.1 pt
  until the label glyph was excluded.
- **PowerPoint's PDF export adds invisible 1.2 pt copies of link text** (fill alpha 0)
  under the link. pdfplumber counts them, so `text_sim`/`word_recall` on d19 now equal
  the truth deck's (0.92) rather than 1.0.
- **reportlab shrinks spaces** (2.69 vs 2.78 pt) to fit a justified line that overflows
  by < 1 pt. PowerPoint cannot, so the frame takes `max(column, natural)`.
- **Layout check by round trip.** Exporting our PPTX to PDF with PowerPoint and parsing
  it gives exact positions. On d04, r01 and d23 every glyph is within 0.45 pt (dx
  median 0.05 pt) horizontally and 0.4 pt vertically of the source PDF. d04's lower
  ssim (0.954) comes from PowerPoint's raster `Slide.Export` of wrapped text; the truth
  deck scores the same 0.955.

### Checks

| Check | Target | Result |
|---|---|---|
| Unit tests on hand-made pages | 2 columns, justified, nested bullets, superscripts, rotated block | **pass** (`tests/test_text_group.py`, 17 tests; also picture bullets, hyphenation, invisible text, axis ticks, roles, XY-cut, numbering) |
| Synthetic: paragraph count + bullet levels = truth | ≥ 95% of slides | **96.4%** (54/56). Bullets present/absent 141/141 |
| Font family = truth family | ≥ 95% of runs | **100%** (215/215). PowerPoint exports name their fonts, so these are name matches; the width-profile path is unit-tested (Georgia recovered from widths alone, distance < 0.01) |
| `boxes_per_slide` ↓ ≥ 5× vs baseline | 5× | **Not met: 531 → 348 boxes (1.5×); on slides without tables/charts/scans 318 → 145 (2.2×).** The truth decks themselves have 158 boxes (3.4× below the baseline) because a table counts as 1 box and a chart as 0; those come in M6/M9. On text-only slides M4 has 145 boxes vs the truth's 137 (+6%) |
| Born-digital `word_recall` ≥ 0.99, `text_sim` ≥ 0.98 | | **met**: median 1.000 / 1.000, mean 0.992 / 0.997 (58 slides; scans and the outlined d28 excluded) |
| `ssim` not below baseline by > 2% | | **met**: synthetic median 0.995 vs 0.992, report median 0.966 vs 0.958. Worst slide d04 −2.0% (= the truth deck) |
| `pytest -q`, `ruff check .`, `mypy` | clean | **110 passed** (~95 s), clean, clean |
| `convert --corpus --no-cache` | | 33/33, all open in PowerPoint, ~9 s. `--mode editable` outputs open too |
| Determinism | byte-identical | two full corpus runs: all 33 PPTX sha256-identical |
| `--text-layout lines` = baseline | | max abs difference 0.0 on all 10 metrics, all 62 slides |
| No leftover POWERPNT.EXE | | none after the tests, corpus runs and scoring |

### Scores (`score --corpus --against out --label m4`) vs BASELINE

Median over slides (mean in parentheses for ssim and text_sim).

| set | label | ssim | reg_min | text_sim | w_recall | pos_err | lb_match | boxes (total) | composite |
|---|---|---|---|---|---|---|---|---|---|
| ALL (62) | baseline | 0.989 (0.987) | 0.957 | 1.000 (0.951) | 1.000 | 0.508 | 1.000 | 5 (531) | 0.896 |
| ALL (62) | **m4** | **0.995 (0.991)** | **0.977** | 1.000 (0.949) | 1.000 | **0.444** | 1.000 | **3 (348)** | **0.898** |
| synthetic (56) | baseline | 0.992 (0.989) | 0.957 | 1.000 (1.000) | 1.000 | 0.506 | 1.000 | 5 | 0.896 |
| synthetic (56) | **m4** | **0.995 (0.993)** | **0.977** | 1.000 (0.997) | 1.000 | **0.436** | 1.000 | **3** | **0.898** |
| synthetic (56) | truth decks | 0.994 (0.992) | 0.975 | 1.000 | 1.000 | 0.380 | 1.000 | 2 | 0.997 |
| report (3) | baseline | 0.958 (0.947) | 0.865 | 1.000 (0.996) | 1.000 | 0.575 | 1.000 | 46 | 0.964 |
| report (3) | **m4** | **0.966 (0.954)** | 0.865 | 1.000 (1.000) | 1.000 | 1.350 | 1.000 | **6** | 0.963 |
| scanned (3) | both | 0.984 | 0.977 | 0.000 | 0.000 | n/a | 0.000 | 0 | 0.437 |

Full tables are in `work/scores/m4/summary.md`; the structure check is in
`work/scores/structure.json`. PowerPoint render + text layout per slide: median
0.158 s (baseline 0.144 s).

**Per-slide drops > 1% vs baseline.** Every one is explained:

| slide | metric | baseline → m4 | truth deck | why |
|---|---|---|---|---|
| d02 p1 | word_recall | 1.000 → 0.808 | 0.808 | real autonumbers: COM `Words` has no list labels ("1.", "a)") |
| d04 | ssim / reg_min | 0.974 / 0.899 → 0.954 / 0.814 | 0.955 / 0.820 | PowerPoint's raster of wrapped text; the layout is verified by round trip (≤ 0.45 pt) |
| d05 | lb_match | 1.000 → 0.875 | 0.875 | `E = mc²` is one COM word but two pdfplumber words |
| d19 p0/p1 | text_sim, w_recall, lb, real_txt | 1.0 → 0.88 / 0.96 | identical | the invisible duplicate link text (alpha 0) is no longer written |
| d16 | reg_min 0.978 → 0.960, lb 1.0 → 0.667 | | 0.981 / 0.111 | the pie legend is one line with kept gaps, and pdfplumber splits it. Charts are M9 |

`pos_err` on the reports (0.58 → 1.35) reflects the scorer's anchor, not the
positions. COM line boxes put a percentage-spaced line's extra leading above the
baseline, so the "bottom − space after" anchor sits lower than for a single-spaced
line. The truth deck d23 (130% spacing) shows the same bias (1.46). The round trip
puts the glyphs within 0.4 pt.

### The 5 worst slides (by composite)

| # | slide | composite | why |
|---|---|---|---|
| 1–3 | s02 p0, s01 p0/p1 | 0.43–0.44 | Scanned, image only (OCR is M8). |
| 4 | d28 p0 | 0.810 | The −30° / 8° text exists only as outlines in the PDF (M8). The 270° box is grouped and rotated correctly. |
| 5 | d19 p0 | 0.836 | See above (the truth deck scores the same text). It also has the one real grouping miss: "Docs: …" and "Write to us: …" are two equal-width lines with nothing that tells a wrap from a paragraph break, so they become one paragraph. |

### Known gaps

- **Font embedding is not written.** Candidates are found and recorded, but a
  PowerPoint `.fntdata` part needs an EOT-style wrapper around a *complete* font with
  a usable cmap. PDF programs are subsets with re-mapped or missing cmaps, and bare
  CFF/Type 1 programs have no OS/2 table at all. Doing it needs a cmap rebuilt from
  ToUnicode plus `name`/`OS/2` tables. No corpus file needs it (every font is
  installed: 0 candidates on 33 PDFs), so it could not be verified either. The
  fallback is the closest installed font.
- Tables and charts stay as separate small boxes (in content order, so text_sim is
  unchanged). They are most of the remaining box count (M6/M9).
- Vertical (stacked) text, underline/strike detection (underlines stay as freeform
  lines) and RTL scripts are not handled.
- Mixed alignment inside one block uses one frame. A centred paragraph next to left
  ones is centred on that frame, which can be off by the frame slack.
- Two equal-width lines with no other evidence (d19 p0) are ambiguous. The
  sentence-end tie-breaker only covers the "sentence + short line" case.
- The `pos_err` anchor bias for non-single spacing (above) should be fixed in the
  scorer: anchor on the baseline (top + 0.75 × line height for r ≥ 1).
- Wrap safety relies on the PDF's own advance widths. A substituted font with
  different widths can still re-wrap (the frame has 2–6 pt of slack). The M5 tuner
  should check PowerPoint's line breaks and adjust the width or spacing.
- The analyze cache key is the stage version only. When iterating on `analyze/`, use
  `--no-cache` or bump the version.

### Next

Milestone 5: the fidelity tuner. **Starting points for Prompt 5** (checked against the
code at `4290aa2`):

- **Compare against M4, not the M2 baseline**, for "no score regresses vs M4":
  `work/scores/m4/` (per file + `summary.json`) and `work/scores/structure.json`.
  `work/scores/baseline.json` stays the M2 baseline. If they are missing, rerun
  `convert --corpus --no-cache` then `score --corpus --against out --label m4`.
- **Where to hook in.** `pipeline.convert()` = `analyzed()` (cached IR) →
  `write_pptx(doc, out, work_dir, pdf, TextOptions.for_mode(...))`. A tuner fits after
  the write: adjust the IR (`TextBlock.bbox`, `Paragraph.line_spacing`,
  `TextRun.char_spacing` / `size`), mark changes `source="tuner"` + a note, rewrite,
  re-render only the changed slides. `Config.tune_enabled` / `tune_rounds` and
  `ModeLimits` (`forced_line_breaks`, `max_font_size_change`, `max_char_spacing_pt`)
  already exist in `config.py`.
- **Comparing lines per block.** PowerPoint side: `ppt_com.shared().text_layout(pres,
  slide)` → `TextShape` per box (`name` = "TextBox N", in spTree order) with
  `LineBox`es and `WordBox`es. PDF side: the block's paragraphs carry `glyph_ids` per
  run, so each block's PDF lines can be rebuilt from `page.glyphs` (group by
  baseline). The writer creates text boxes in `_build_page` order (z, then
  `reading_order`); map shapes to blocks by writing the block id into the shape name
  (not done yet: cheap to add in `writer._add_text`).
- **Rounding.** PowerPoint rounds `spcPts` to whole points and `spcPct` to whole
  percent (M4). Any "line spacing exact" fix must go through
  `build/units.ppt_spacing()`, or it will be silently re-rounded. Whether `spc`
  (character spacing) is rounded too is **not measured yet**; measure it first.
- **First candidates.** `TextBlock.notes` carries `wrap_conflict`, `vdrift:<pt>` and
  `tight:<i>`; paragraphs carry their split reasons; the report lists blocks with
  confidence < 0.6. Frames have 2–6 pt of wrap slack (`text_group._horizontal`).
- **An exact signal.** The round trip (PowerPoint `export_pdf` of the PPTX, then
  `ingest.parse.parse_pages` and compare glyph origins in order) is exact and costs
  ~1 s per deck. It is a better line-break / position check than COM word boxes.
- **Known scorer bias.** `pos_err` uses "line bottom − space after" as the anchor;
  with percentage line spacing that sits ~1 pt lower than the baseline (reports,
  d23). Fix it in the scorer (anchor = top + 0.75 × line height when the ratio is ≥ 1)
  before using pos_err as the tuner's objective, and rescore truth/baseline/m4 when
  you do.
- **Test fixtures.** `tests/test_text_group.py::PageBuilder` builds hand-made pages
  (monospace 0.5 em glyphs, wrapped/justified text). Use it for the "block 10% too
  narrow" test: shrink the analyzed `TextBlock.bbox` before writing.

---

## Milestone 5: fidelity tuner (2026-11-12)

### What exists

- **`verify/tuner.py`**: render → compare → adjust → re-render, per text block.
  - **Mapping.** The writer names every block's text box `TextBlock <block id>`, so
    COM `text_layout` shapes map back to IR blocks. Both sides are reduced to offsets
    in the block's non-space character stream. PDF lines come from the runs'
    `glyph_ids` (a new line where the baseline jumps; sub/superscripts stay on their
    line). PowerPoint lines come from `text_layout`. A break matches or it doesn't.
  - **Positions.** A PDF round trip (PowerPoint `SaveAs` PDF; glyphs read only on the
    measured slides). The exported glyphs inside a COM line box give that line's exact
    baseline and first-character x.
  - **Wrap model.** Built from PowerPoint's own word boxes: token widths, natural gaps
    (measured on unstretched lines; for justified paragraphs on their last line), and
    line starts. PowerPoint may wrap after a hyphen. Per PDF line it gives the width
    the line needs, and the width at which the next word would join it, as linear
    functions of the spc change δ and the size factor f.
  - **Fixes, in order, bounded by `ModeLimits`:**
    1. Breaks wrong: the first feasible (f, δ) in the order width only → spc → size
       (0.1 pt steps). Justified blocks prefer the PDF's column width (same word
       spacing); others keep the current width when it is inside the interval.
       A per-block **width bracket** is kept from what PowerPoint actually did at each
       width: at the first differing break, PowerPoint breaking first means too
       narrow, otherwise too wide. A model answer outside the bracket is replaced by
       bisection, which ended an oscillation on r01. Last resort, only when
       `forced_line_breaks` is allowed (fidelity): a soft break (`a:br`) at every PDF
       wrap position.
    2. Breaks right: exact line spacing where PowerPoint's pitch differs from the
       PDF's by more than 0.3 pt. It goes through `units.ppt_spacing` and is applied
       only if the rounded value is closer.
    3. Then an offset that puts the first baseline on the PDF's. Start x moves too
       when all lines agree (spread ≤ 0.4 pt). The offset is capped by
       `max_offset_pt` (the tuner's own cumulative move).
  - **Never worse.** Every change is re-measured. The best version per block (break
    mismatches, then median line-anchor distance) is kept; a worse result is reverted
    and the block frozen. Blocks that match (0 breaks, pos ≤ 0.3 pt) stop early.
  - **Rounds.** ≤ `tune_rounds` (default 4) re-renders; only slides with changes are
    re-measured.
  - **Whole-slide pass** (fidelity, balanced). After the text rounds, each shape and
    picture's PDF render is matched against the PowerPoint render (±6 px, NCC ≥ 0.9,
    ≥ 0.03 above the unshifted match, confirmed by the reverse match). It is
    translated back, kept only if its local SSIM improves, and reverted otherwise.
  - **Record.** A changed element gets `source="tuner"`, `tuner:from:<old source>`
    and net-change notes (`tuner:width:+3.20`, `tuner:spc:-0.05`, `tuner:size:-0.2`,
    `tuner:lnspc:1:exact=14.4`, `tuner:move:dx,dy`, `tuner:br:<para>:<n>`,
    `tuner:eval:breaks a->b,pos a->b`). Unfixed blocks get `tuner:unresolved:…`. The
    run summary (rounds, timing by phase and by slide, counts) is in
    `Document.metadata["pdf2pptx:tuner"]` and in `<name>.convert.json["tune"]`.
  - **Degrades.** Without PowerPoint, or on any tuner error, the untuned deck stays
    and the reason is recorded. Rotated and vertical blocks, and `--text-layout
    lines`, are not tuned.
- **Mode limits** (`config.ModeLimits`; new fields: `max_width_change`, `max_offset_pt`,
  `slide_pass`):

  | mode | forced breaks | size | spc | width | offset | slide pass | tune by default |
  |---|---|---|---|---|---|---|---|
  | fidelity | yes | ±10% | ±1.5 pt | ±25% | 6 pt | yes | yes |
  | balanced | no | ±5% | ±0.8 pt | ±15% | 4 pt | yes | yes |
  | editable | no | ±3% | ±0.4 pt | ±10% | 3 pt | no | no |

- **Pipeline / CLI.**
  - `convert … --tune/--no-tune --tune-rounds N` (default: the mode's setting).
  - A new common flag `--out-dir`, so per-mode corpus runs can live side by side
    (`convert --corpus --mode fidelity --out-dir out/fidelity`, then
    `score --corpus --out-dir out/fidelity`).
  - `convert --corpus` builds in the process pool, then tunes every file
    sequentially in the main process on one PowerPoint engine. Tuning time is added
    to `seconds` / `seconds_per_page`.
- **Writer.** Shape names `TextBlock <id>`; `\v` in a run is written as `<a:br>` with
  the run's properties (`_xml_safe` used to strip it).
- **Scorer fix** (`ppt_com._collect_text_frame`). A line that ends in a soft break
  (VT) is no longer treated as a paragraph end: space-after belongs to the
  paragraph's last line only. Rescoring the truth decks gives no per-slide change
  > 1%.
- **Report.** Each page lists its tuner adjustments (id, kind, notes), and
  `source=tuner` boxes have their own colour.
- **`ppt_com.set_geometry`**: a COM in-memory edit of left/top/width, kept only to
  re-run the edit-path benchmark below.

### Measured facts that later milestones rely on

- **`spc` is not rounded.** The rendered extra advance per character is ≈ 0.955 × spc
  (0.13 pt → 0.116, 0.5 → 0.471, 1.0 → 0.955, −0.5 → −0.497; 20 pt Arial). The
  model uses `SPC_EFFECT = 0.955`.
- **Font sizes in 0.1 pt steps are honoured** (20.1/20.2/20.5/20.7 pt all render at
  their size).
- **Edit path: python-pptx rewrite + reopen vs COM in-memory edits** (per round,
  including `text_layout` of the slide; median of 3):

  | deck | rewrite + reopen + layout | COM edit + layout | export 1 slide (ranged) | export deck |
  |---|---|---|---|---|
  | d23 (2 slides) | 0.35 s | 0.25 s | 0.18 s | 0.19 s |
  | r01 (1 slide, 376 words) | 0.88 s | 0.72 s | 0.29 s | 0.26 s |
  | d26 (1 slide, 20 boxes) | 0.28 s | 0.19 s | 0.28 s | 0.27 s |
  | d17 (5 slides) | 0.16 s | 0.06 s | 0.13 s | 0.16 s |

  **Chosen: python-pptx.** COM saves ~0.1 s per round because PowerPoint's layout
  dominates, but it would measure a file that differs from the one shipped, with
  different rounding of what COM sets.
- **A ranged `ExportAsFixedFormat` (PrintOptions.Ranges) crashes PowerPoint**
  (`0x800706BA`, "Server RPC non disponibile") on the first export after some decks
  (d12 → d13, d13 → d14, reproducible in sequence, never in isolation). The engine
  recovered each time, but the tuner now uses whole-deck `SaveAs` PDF, which has
  been stable since M1 and costs about the same.
- **Intermittent `SaveAs` PDF stalls of 7–10.4 s.** Seen in 3 of ~10 full corpus runs,
  on 1–5 decks each, never on the same deck twice; timing shows the wait inside the
  export call. No error is raised. The cause is not established (a constant
  ~10.2 s suggests a retry on a locked file, e.g. a scanner). They push single
  slides past 3 s but not the median.

### Checks

| Check | Target | Result |
|---|---|---|
| Synthetic, balanced: `linebreak_match` | ≥ 0.98 | **median 1.000**, mean 0.970. The truth decks themselves score mean 0.884. Every slide below 0.98 scores the same as or better than its truth deck: d28 (outlined text, 0.111), d16 (pie legend split by pdfplumber, 0.667 vs truth 0.111), d19 p0/p1 (invisible duplicate link text, 0.833), d05 (`E = mc²`, 0.875). Without d28 the mean is 0.986 |
| Synthetic, balanced: median `ssim` | ≥ 0.95 | **0.995** |
| Synthetic, balanced: median `pos_err_pt` | ≤ 1.5 | **0.428** (M4 0.436) |
| Editable: no forced line breaks | none in output | `test_editable_never_forces_line_breaks_fidelity_may`: a block 30% too narrow in editable → 0 `a:br` and 0 forced; the same block in fidelity with no width/spc/size room → forced breaks, all breaks match. Corpus: 0 forced breaks in every mode |
| A block 10% too narrow is fixed within the round limit | ≤ 4 rounds | `test_a_block_10_percent_too_narrow_is_fixed` (d23 justified paragraphs, balanced): breaks 9 → 0 in 1 fix round, `tuner:width:+…` recorded; two runs give byte-identical PPTX |
| Never worse | | Every PowerPoint test asserts best ≥ first per block; corpus runs: 0 reverts |
| Deterministic | | two tuned corpus runs: **33/33 PPTX byte-identical** |
| Tuning time per slide (median) | ≤ 3 s | **0.37 s** (mean 0.57, p90 1.24, max 2.97; n = 59 slides, balanced; per document 0.68 s/slide median). Stress set (all blocks perturbed): **0.88 s** median, p90 2.36, max 3.97 |
| No score regresses vs M4 by > 1% | | **none** in balanced, fidelity, editable or editable --tune. The one per-slide change > 1% is better (d16 region_ssim_min 0.960 → 0.980) |
| `pytest -q`, `ruff check .`, `mypy` | clean | **119 passed** (~100 s), clean, clean |
| `convert --corpus --no-cache` | | 33/33, 33/33 open in PowerPoint, 44 s including tuning; no COM warnings in the final runs |

### Scores per mode (median over slides; mean in parentheses)

| set | label | ssim | reg_min | text_sim | w_recall | pos_err | lb_match | boxes | composite |
|---|---|---|---|---|---|---|---|---|---|
| ALL (62) | M4 | 0.995 (0.991) | 0.977 | 1.000 | 1.000 | 0.444 (0.460) | 1.000 (0.925) | 3 | 0.898 |
| ALL (62) | **fidelity** | 0.995 (0.991) | 0.977 | 1.000 | 1.000 | **0.440** (0.457) | 1.000 (0.925) | 3 | 0.898 |
| ALL (62) | **balanced** | 0.995 (0.991) | 0.977 | 1.000 | 1.000 | **0.440** (0.457) | 1.000 (0.925) | 3 | 0.898 |
| ALL (62) | **editable** (untuned default) | 0.995 (0.991) | 0.977 | 1.000 | 1.000 | 0.444 (0.460) | 1.000 (0.925) | 3 | 0.898 |
| ALL (62) | editable --tune | 0.995 (0.991) | 0.977 | 1.000 | 1.000 | 0.440 (0.457) | 1.000 (0.925) | 3 | 0.898 |
| synthetic (56) | M4 | 0.995 (0.993) | 0.977 | 1.000 | 1.000 | 0.436 (0.423) | 1.000 (0.970) | 3 | 0.898 |
| synthetic (56) | balanced / fidelity | 0.995 (0.993) | 0.977 | 1.000 | 1.000 | **0.428** (0.420) | 1.000 (0.970) | 3 | 0.898 |
| report (3) | M4 = every mode | 0.966 (0.954) | 0.865 | 1.000 | 1.000 | 1.350 (1.144) | 1.000 (1.000) | 6 | 0.963 |

Labels: `work/scores/m5_balanced`, `m5_fidelity`, `m5_editable`,
`m5_editable_tuned`; `truth_m5` is the truth rescore.

**What the tuner found on the real corpus.** Almost nothing to do: M4's analysis
already places the corpus within PowerPoint's rounding. Of the 347 tunable blocks,
346 match before tuning: 0 break mismatches, and every line's baseline and start
within 0.3 pt. The one exception is d16's pie legend, 0.37 pt off; it was moved
−0.36 pt (0.37 → 0.07). The slide pass found 8 candidate shape shifts, and all 8
failed the SSIM check and were reverted. Fidelity and balanced therefore score
the same here; they differ only in what they are *allowed* to do (see the stress set).

### Stress set: every block perturbed (balanced)

To exercise the tuner on the corpus, every wrapping, unrotated block on every
non-scanned page (347 blocks) was made 10% narrower (alignment anchor kept) and
moved 1.5 pt down. The perturbed decks were then scored untuned and tuned.

| | breaks (sum) | blocks matched | ssim | reg_min | pos_err median (mean) | lb_match mean | composite |
|---|---|---|---|---|---|---|---|
| unperturbed (balanced) | 0 | 347 | 0.995 | 0.977 | 0.440 (0.457) | 0.925 | 0.898 |
| perturbed, untuned | **291** | 0 | 0.959 | 0.815 | 2.166 (20.19) | 0.859 | 0.879 |
| perturbed, **tuned** | **0** | **347** | **0.994** | **0.975** | **0.379** (0.411) | **0.972** | **0.898** |

- 30 of 31 decks finish in 2 rounds (one fix + one position fix) and one in 3, with
  0 reverts.
- Reports alone: ssim 0.919 → 0.967 (unperturbed 0.966) and lb_match 0.625 → 1.000.
- Getting there needed three fixes found on this set: the width bracket, natural
  gaps of justified paragraphs, and the justified-column preference. A
  dataclass-field-order bug from the last of these had silently dropped every
  wrap upper bound; a unit test now guards it.

### Top remaining failure types

1. **Scorer-side line-break misses, not layout.** d28 (text outlined in the PDF), the
   d16 legend, d19 (invisible duplicate link text) and d05 (`E = mc²`) keep
   `linebreak_match` mean at 0.970. The truth decks score the same or lower.
2. **Any width inside the feasible interval is accepted.** On the stress set a few
   slides keep 1–2% lower region SSIM than the unperturbed analysis (d01, d19, d22,
   d23, r01) with every break and baseline right. The residue is sub-0.3 pt
   horizontal differences and justified stretch within the 1.5 pt margin.
3. **Width first, spc second.** When the cause is a wider substituted font, the
   block is widened until the breaks match. The breaks are then right, but words
   drift right along the line; spc is only used when width alone cannot work.
   Charging horizontal word error in the block score would let spc win there.
4. **Intermittent 7–10 s PowerPoint PDF-export stalls** (see above). They are rare
   but set the per-slide maximum.
5. **Not tuned:** rotated and vertical text, `wrap=False` boxes' widths, table and
   chart frames (still separate text boxes until M6/M9). The slide pass costs up to
   ~2 s on chart slides for candidates that, on this corpus, are all rejected.

### Next

Milestone 6: graphics and tables (native shapes, native tables). The tuner works on
any `TextBlock`. When tables arrive, their cell text shows up in `text_layout` as
`<table>[r,c]` shapes and needs its own mapping (cell → block) before it can be
tuned.
