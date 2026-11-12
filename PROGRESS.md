# pdf2pptx: Progress

| # | Milestone | Status | Date |
|---|---|---|---|
| 1 | Foundation, IR, corpus, scoring harness | **done** | 2026-11-11 |
| 2 | Ingest and baseline converter | **done** | 2026-11-12 |
| 3 | PowerPoint check engine and debug report | **done** | 2026-11-12 |
| 4 | Text understanding and fonts | not started | |
| 5 | Fidelity tuner | not started | |
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
