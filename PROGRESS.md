# pdf2pptx: Progress

| # | Milestone | Status | Date |
|---|---|---|---|
| 1 | Foundation, IR, corpus, scoring harness | **done** | 2026-11-11 |
| 2 | Ingest and baseline converter | not started | |
| 3 | PowerPoint check engine and debug report | not started | |
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
