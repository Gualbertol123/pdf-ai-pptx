# pdf2pptx: 10 build prompts for Claude Code (Opus)

## How to use

- Run them **in order**, **one prompt per fresh Claude Code session**, started from the `pdf_pptx/` folder.
- Each prompt can be pasted as-is. It points Claude at `PLAN.md` (the design) and `PROGRESS.md` (what's done so far).
- For the large prompts (4, 5, 6, 9, 10), start in **plan mode** (Shift+Tab), review the plan, then let it build.
- Between prompts, read `PROGRESS.md` and look at the debug report. Don't move on if a check failed. Re-run the same prompt and add "continue: the following checks failed: …".
- Put your own test PDFs in `corpus/real/` at any time. Every later prompt scores them too.

Each prompt has: **Context → Objective → Deliverables → Constraints → Checks → Report**.

---

## Prompt 1: Foundation, IR, synthetic corpus, scoring harness

```text
CONTEXT
You are starting the pdf2pptx project in this folder. Read PLAN.md completely
before doing anything: it is the design document and its §8 "Engineering rules"
are binding for all work. Target machine: Windows 11, Python 3.12, PowerPoint
installed, Radeon 780M iGPU, 31 GB RAM. The folder currently contains only PLAN.md
and PROMPTS.md.

OBJECTIVE
Build the foundation everything else stands on: project skeleton, the IR,
config/modes, a synthetic corpus with ground truth, and the scoring harness that
all later milestones are judged by.

DELIVERABLES
1. git init; .gitignore (work/, .venv/, tools/, corpus/real/, models, *.pptx
   outputs under out/); a venv at .venv created with `py -3.12 -m venv .venv`.
2. pyproject.toml (src layout, package `pdf2pptx`, console script `pdf2pptx`),
   ruff + pytest + mypy config. Pin dependency minimum versions.
3. src/pdf2pptx/ir.py: full IR per PLAN.md §4.1 as dataclasses, with
   to_json/from_json, ir_version, and exact JSON round-trip. units.py with pt↔EMU.
4. config.py: modes (fidelity|balanced|editable|deck), tier detection stub,
   paths (work dir, %LOCALAPPDATA%\pdf2pptx\models), loaded from defaults +
   optional pdf2pptx.toml + CLI overrides.
5. cli.py with subcommands (stubs are fine where later milestones fill them in):
   convert, score, corpus, report, doctor, setup.
   `pdf2pptx doctor` must already report: Python, PowerPoint COM availability,
   onnxruntime providers (if installed), Vulkan dll presence, RAM, GPU name.
6. corpus generator (`pdf2pptx corpus build`):
   - ≥ 25 synthetic decks made with python-pptx covering: titles+bullets (multi-level),
     multi-column text, mixed fonts/sizes/colors, Italian accented text, tables
     (incl. merged cells), preset shapes (rect, roundRect, ellipse, arrows,
     chevron, connectors), gradients and transparency, images, a native bar/line/pie
     chart, footer + slide number + logo repeated across slides, a build-up
     sequence (same slide repeated with one more bullet each time), hyperlinks,
     speaker notes, sections, 16:9 and 4:3.
   - Export each to PDF through PowerPoint COM (SaveAs ppSaveAsPDF=32), keep the
     source .pptx as ground truth alongside the PDF and a truth.json (text per
     slide, shape kinds/counts, tables, charts).
   - 3 "report-style" PDFs with reportlab (multi-column, table, figure caption).
   - 2 "scanned" PDFs: rasterize two synthetic PDFs at 200 DPI, add slight
     rotation (≤1°), noise and JPEG artifacts, rewrap as image-only PDFs.
   - A manifest corpus/manifest.json listing every file with its category.
7. verify/metrics.py implementing every metric in PLAN.md §6 that does not need
   later milestones. For metrics that need PowerPoint rendering (ssim, pos_err,
   linebreak_match) implement them against PNG renders and a word list input;
   the PowerPoint renderer itself comes in Prompt 3, so for now render the
   *ground-truth* PPTX via a minimal COM export to validate the metrics.
   `pdf2pptx score <pdf> <pptx>` and `pdf2pptx score --corpus` produce
   per-slide JSON + a summary table (markdown) in work/scores/.
8. PROGRESS.md created with a milestone table and a "Milestone 1" entry.
9. README.md: setup, commands, what needs IT (driver update, power mode).

CONSTRAINTS
- Follow PLAN.md §5 license rules (no AGPL dependencies).
- Deterministic corpus: fixed seeds, stable file names.
- COM code here is temporary; keep it isolated in one module so Prompt 3 can
  replace it.

CHECKS (all must pass before you finish; show the output)
- `.venv\Scripts\pytest -q` passes, incl. IR round-trip tests for every element type.
- `.venv\Scripts\ruff check .` clean; mypy clean on ir.py and units.py.
- `pdf2pptx doctor` runs and reports PowerPoint COM = available.
- `pdf2pptx corpus build` produces ≥ 30 PDFs with manifest + truth files.
- Sanity check of the metrics: scoring each ground-truth PPTX against its own
  PDF gives ssim ≥ 0.97 and text_sim ≥ 0.99 (this proves the metrics, not the
  converter). Scoring a blank PPTX gives clearly low scores.
- Running corpus build twice gives identical truth.json files.

REPORT
Update PROGRESS.md (what exists, metric sanity results, known gaps). Commit
with message "M1: foundation, IR, corpus, scoring". Summarize in ≤ 15 lines.
```

---

## Prompt 2: Ingest and baseline converter

```text
CONTEXT
pdf2pptx project. Read PLAN.md (design + §8 rules) and PROGRESS.md (state).
Milestone 1 is done: IR, config, CLI, synthetic corpus, scoring harness.

OBJECTIVE
Get a complete, deliberately simple end-to-end conversion working for every
corpus file and record the baseline scores that all later work must beat.

DELIVERABLES
1. ingest/parse.py with pypdfium2: per page extract glyphs (unicode, font
   name/flags, size, color, bbox, baseline, rotation via text matrix), vector
   paths (segments, fill/stroke, line width, dash, opacity, transforms resolved
   to page space, clipping recorded), images (bitmap, placement matrix, soft
   mask), annotations/links, document metadata (producer, creator, page sizes).
   Cross-check text against pdfplumber on a sample and log disagreements.
2. ingest/render.py: page renders at 150 and 300 DPI, created lazily and cached in
   work/<hash>/renders/.
3. ingest/triage.py: classify each page born_digital | broken_text |
   outlined_text | scanned, plus a reason. Heuristics: no glyphs with large
   image coverage → scanned; much private-use / unmapped unicode or
   unreadable text (character-bigram score) → broken_text; many small filled
   paths shaped like letters with no glyphs → outlined_text.
4. build/writer.py baseline: slide size from page size; one text box per text
   line (exact position, font, size, color, no autofit, no wrap); images placed
   with crop; every path as a freeform shape; z-order kept. Text boxes with
   zero margins.
5. pipeline.py: ingest → triage → (analyze = identity for now) → build, caching
   the IR per stage in work/<sha256>/ir_<stage>.json.
6. `pdf2pptx convert in.pdf -o out.pptx [--mode]` and
   `pdf2pptx convert --corpus` (parallel over pages/files with a process pool,
   default workers = 6).

CONSTRAINTS
- Never crash on one bad page: catch, log, fall back to placing that page's
  300-DPI render as a full-slide picture, and mark it in the IR.
- Rotated text and non-zero page rotation must be handled (not just ignored).
- Memory: stream pages; don't hold all renders in RAM.

CHECKS
- Unit tests: parse a hand-made PDF (reportlab) with known glyph positions;
  positions within 0.5 pt; colors exact; a rotated text line detected.
- Triage tests: the two scanned corpus PDFs → scanned; born-digital → born_digital.
- `pdf2pptx convert --corpus` converts 100% of files; every output opens in
  PowerPoint through COM without a repair prompt (check Presentations.Open
  succeeds and the slide count matches).
- `pdf2pptx score --corpus` runs; save the result as the BASELINE in
  work/scores/baseline.json and paste the summary table into PROGRESS.md.
- pytest and ruff clean.

REPORT
PROGRESS.md "Milestone 2" with baseline table and the 5 worst slides (with
reasons). Commit "M2: ingest + baseline converter".
```

---

## Prompt 3: PowerPoint check engine and debug report

```text
CONTEXT
pdf2pptx. Read PLAN.md (§6 scores, §8 rule 6 on COM) and PROGRESS.md.
Baseline conversion and scores exist. Rendering through COM is currently a
temporary helper from M1.

OBJECTIVE
Build the production PowerPoint COM engine that drives the check-and-correct
loop, plus a visual debug report that makes every conversion easy to inspect.

DELIVERABLES
1. verify/ppt_com.py: a PowerPoint worker in a separate process (multiprocessing,
   single STA thread) exposing over a queue:
   open(pptx) / render(slide_ids, dpi) → PNGs / text_layout(slide_id) → for each
   text shape: lines (text, BoundLeft/Top/Width/Height), words with bounds /
   close / ping. Heartbeat, per-call timeouts, automatic restart, cleanup of
   only the POWERPNT.EXE processes it started, WithWindow=False, no dialogs
   (DisplayAlerts off). One long-lived instance reused across calls.
2. Replace the M1 temporary COM helper with this module everywhere.
3. verify/diff.py: aligned per-slide comparison: SSIM map, 4×4 region scores,
   a heatmap PNG, and a list of mismatch regions with the IR element ids that
   overlap them.
4. Fill in pos_err_pt and linebreak_match in metrics.py using text_layout().
5. report/debug_html.py: `pdf2pptx report <pdf>` writes one self-contained HTML
   (no external assets) with, per page: PDF render | PPTX render | diff heatmap,
   toggleable IR bbox overlays coloured by element type and source, hover
   showing the element's JSON, page triage result and score row. Index page
   with a sortable score table.
6. Speed: reuse the open presentation; render only the slides that were asked for.

CONSTRAINTS
- If PowerPoint is missing or broken, everything still runs: scores that need
  it are reported as null and the report says why.
- No leftover PowerPoint windows/processes after any command (verify).

CHECKS
- Stress test: 200 render calls in a row on the corpus with a deliberately
  killed worker in the middle → the engine recovers and all calls complete.
- Median render+layout time ≤ 1.0 s per slide on the corpus (print timings).
- text_layout returns correct line counts on 5 synthetic slides where the
  ground truth is known.
- `pdf2pptx report` produces the HTML for 3 corpus files; open one and confirm
  overlays align with the renders (describe what you checked).
- After the test suite: no POWERPNT.EXE started by us is still running.
- pytest and ruff clean. Corpus scores re-run; baseline numbers unchanged except
  for newly available metrics.

REPORT
PROGRESS.md "Milestone 3" with timings and the new full baseline. Commit
"M3: PowerPoint verify engine + debug report".
```

---

## Prompt 4: Text understanding and fonts

```text
CONTEXT
pdf2pptx. Read PLAN.md, PROGRESS.md and the latest scores in work/scores/.
Baseline writes one text box per line. This milestone makes text real:
paragraphs, blocks, lists, reading order, font matching.

OBJECTIVE
Turn glyphs into well-structured, editable text that PowerPoint lays out like
the PDF, with far fewer boxes.

DELIVERABLES
1. analyze/text_group.py: glyphs → words → lines → paragraphs → TextBlocks, using
   page-adaptive statistics (median char width, line pitch histograms), not
   fixed thresholds. Handle justified text, hyphenation at line ends, superscript
   and subscript, mixed runs (font/size/color/bold/italic changes inside a line),
   rotated blocks.
2. analyze/reading_order.py: recursive XY-cut with column detection; a
   reading-order index on blocks.
3. analyze/lists.py: detect bullets (glyph types: •, –, ▪, images used as bullets)
   and numbering (1. / a) / i.) and indent levels → Paragraph.level + bullet
   definition; remove the bullet glyph from the text.
4. Paragraph properties: alignment, left/first-line indent, line spacing (exact
   points from baselines), space before/after.
5. Block roles (title, subtitle, body, footer, page_number, caption) from position,
   size and repetition heuristics; confidence recorded.
6. fonts/index.py: index fonts in C:\Windows\Fonts and
   %LOCALAPPDATA%\Microsoft\Windows\Fonts with fontTools (family, style, weight,
   width, advance-width table, embedding flags), cached.
   fonts/match.py: map a PDF font (name + flags + widths of the glyphs actually
   used) to the closest installed font by width-profile distance; strip subset
   prefixes (ABCDEF+); bold/italic from name/flags/stem weight.
   fonts/embed.py: extract embedded PDF fonts; if a font is installable
   (OS/2 fsType allows it) and not installed, record it for PPTX embedding
   (implement embedding if feasible; otherwise document it as a gap).
7. writer.py: one text box per TextBlock with wrapping on, margins 0, width from
   the block bbox (+ small tolerance), paragraph/run properties, bullets as
   real PowerPoint bullets; autofit off in fidelity mode, shrink-on-overflow in
   editable mode.

CONSTRAINTS
- The writer must keep the baseline line-per-box path behind a flag for comparison.
- Every grouping decision that is uncertain gets a confidence; low confidence
  blocks are listed in the debug report.

CHECKS
- Unit tests for grouping on hand-made pages: 2-column text, justified
  paragraph, nested bullets, superscripts, rotated block.
- On synthetic decks: paragraph count and bullet levels match truth.json for
  ≥ 95% of slides.
- Corpus: boxes_per_slide decreases ≥ 5× vs baseline; word_recall ≥ 0.99 and
  text_sim ≥ 0.98 on born-digital pages; ssim not lower than baseline by more
  than 2% (the tuner in M5 will recover the rest).
- Font matching: for synthetic decks the matched family equals the truth family
  for ≥ 95% of runs.
- pytest and ruff clean.

REPORT
PROGRESS.md "Milestone 4" with before/after score table and the remaining
worst cases. Commit "M4: text understanding + fonts".
```

---

## Prompt 5: Fidelity tuner (the check-and-correct loop)

```text
CONTEXT
pdf2pptx. Read PLAN.md (§3 modes, §6 scores), PROGRESS.md, latest scores.
Text is now real paragraphs, but PowerPoint wraps/positions text slightly
differently from the PDF. The COM engine can report exact line breaks and
word bounds.

OBJECTIVE
Close the gap automatically: render → compare → adjust → re-render, until line
breaks and positions match, while respecting the selected mode.

DELIVERABLES
1. verify/tuner.py, per text block:
   - Compare PowerPoint lines (text_layout) with the PDF lines for the block.
   - Fixes, applied in this order and bounded by mode: box width ±; character
     spacing (spc) ±; font size in 0.1 pt steps; line spacing exact; vertical
     offset to align the first baseline; as a last resort in fidelity mode,
     a forced line break at the PDF break position.
   - Per-mode limits from config (e.g. editable: no forced breaks, max ±3%
     size change; fidelity: everything allowed).
2. A whole-slide pass after the text pass: a pixel diff to find residual offsets
   (shapes/images) and correct them by translation if consistent.
3. Iteration control: ≤ 4 rounds; re-render only slides with changes; stop
   early when a block matches; all adjustments recorded in the IR (source="tuner")
   so they're visible in the report.
4. Batch API: tune all slides of a deck in one open presentation session
   (edit through python-pptx between rounds, or directly via COM if faster:
   measure both and choose).
5. CLI: `--tune/--no-tune`, `--tune-rounds N`; on by default for fidelity and
   balanced.

CONSTRAINTS
- Tuning must never make a block's score worse: keep the best version seen.
- Deterministic: same input → same adjustments.
- Total tuning time ≤ 3 s per slide median on this machine.

CHECKS
- Synthetic corpus: linebreak_match ≥ 0.98, median ssim ≥ 0.95, pos_err_pt
  median ≤ 1.5 pt, in balanced mode.
- Editable mode: no forced line breaks in output (assert in test).
- A test where a block is deliberately 10% too narrow → the tuner fixes it
  within the round limit.
- Print time per slide; meet the time limit.
- No score regresses vs M4 by more than 1% (except explained trade-offs).
- pytest and ruff clean.

REPORT
PROGRESS.md "Milestone 5": scores per mode (fidelity/balanced/editable), timing,
and the top remaining failure types. Commit "M5: fidelity tuner".
```

---

## Prompt 6: Graphics and tables

```text
CONTEXT
pdf2pptx. Read PLAN.md, PROGRESS.md, latest scores. Paths are currently written
as freeform shapes; tables come out as loose text and lines.

OBJECTIVE
Rebuild vector graphics as built-in PowerPoint shapes with correct styling, and
detect tables as native, editable tables.

DELIVERABLES
1. analyze/shapes.py:
   - Fit paths to preset shapes with a tolerance: rect, roundRect (fit the
     corner-radius adjust value), ellipse, triangle, rtTriangle, diamond,
     parallelogram, trapezoid, chevron, homePlate, rightArrow/leftArrow (fit
     adjust values), line, plus rotation.
   - Lines/arrows whose endpoints touch shapes → connectors (straight/elbow)
     with arrowheads.
   - Merge a fill path and its stroke path into one shape; group shapes that
     were grouped in the PDF content (form XObjects / marked content).
   - Styling: solid fill, linear/radial gradient (stops, angle), opacity,
     stroke width/color/dash/cap/join.
   - Background detection: full-bleed rect or image behind everything → slide
     background (or bottom-most locked shape).
   - Very dense vector clusters that don't fit (icons, illustrations) → one SVG
     picture with a PNG fallback (PowerPoint supports SVG), positioned exactly.
2. analyze/tables.py:
   - Detect tables from (a) ruling lines forming a grid and (b) whitespace-aligned
     text columns/rows with no lines; build the cell grid, find merged cells,
     header row, cell padding, fills, borders.
   - Output native tables (graphicFrame) with matching borders, fills,
     alignment, fonts; cell text uses the M4 paragraph machinery.
   - Confidence; low confidence → keep as text + shapes (no regression).
3. build/xml_ext.py helpers for everything python-pptx lacks (gradients, dash,
   arrowheads, connectors, SVG pictures, table borders), each covered by a test
   that opens the result in PowerPoint through COM.

CONSTRAINTS
- Shape fitting must not change the look: accept a preset only if the render
  diff of that region is within tolerance; otherwise keep the freeform.
- Run the M5 tuner after these changes; table cells participate in tuning.

CHECKS
- Synthetic: native_shape_ratio ≥ 0.80; preset kinds match truth.json for
  ≥ 90% of shapes; every truth table is detected with correct rows×cols and
  merges for ≥ 90% of tables.
- Report-style PDFs: tables detected (whitespace tables included).
- ssim does not regress vs M5 by more than 1%.
- Every xml_ext feature opens in PowerPoint without repair.
- pytest and ruff clean.

REPORT
PROGRESS.md "Milestone 6" with scores and examples of remaining misfits.
Commit "M6: graphics + tables".
```

---

## Prompt 7: Document-level structure

```text
CONTEXT
pdf2pptx. Read PLAN.md, PROGRESS.md, latest scores. Slides are now good one by
one; the deck as a whole has no masters, theme, sections, notes or animations.

OBJECTIVE
Recover the document's structure so the PPTX behaves like a deck someone built,
not a set of separate pages.

DELIVERABLES
1. analyze/masters.py: find elements that repeat across pages (same content
   hash + position within tolerance: logos, footers, page numbers, background
   art); move them into the slide master/layouts; replace page numbers with the
   slide-number field. Cluster pages by layout signature (title position, body
   regions) into ≤ N layouts and map title/body blocks to placeholders when
   the geometry fits (editable/balanced modes).
2. analyze/theme.py: extract the color palette (frequency weighted, from fills
   and text) → theme colors (dk1, lt1, accent1–6); dominant heading/body fonts
   → theme fonts; elements using these colors/fonts reference the theme instead
   of hard-coded values (editable/balanced).
3. analyze/builds.py: detect build-up sequences (page n+1 = page n + new
   elements, everything else identical within tolerance) → collapse into one
   slide with "appear" animations in order (raw XML timing tree). Config flag
   to disable.
4. analyze/links.py: external hyperlinks on text/shapes; internal links → jumps
   to slides; PDF outline/bookmarks → PowerPoint sections; text annotations and
   comments → speaker notes; document title/author → core properties.
5. Producer detection: if the PDF was made by PowerPoint/Keynote/Google Slides
   (metadata + page size heuristics), enable stronger placeholder mapping.

CONSTRAINTS
- Fidelity mode: masters/theme may be used only if the render stays within ssim
  tolerance; otherwise keep hard-coded values.
- Build collapsing must be reversible by config and noted in the report.

CHECKS
- Synthetic: the repeated logo/footer/slide number appear once in a master,
  not per slide; slide number fields present; the build-up deck collapses to
  the right number of slides with the right number of animation steps (verify
  via COM: TimeLine.MainSequence.Count).
- Theme accent colors match the truth deck's theme for ≥ 4 of 6 accents.
- Sections and speaker notes match truth.json.
- No score regresses by more than 1% vs M6.
- pytest and ruff clean.

REPORT
PROGRESS.md "Milestone 7". Commit "M7: masters, theme, builds, links".
```

---

## Prompt 8: Hard pages: OCR, layout model, background cleanup

```text
CONTEXT
pdf2pptx. Read PLAN.md (§2 hardware, §5 components, §8 rules 2/3/9),
PROGRESS.md. Triage already flags scanned / broken_text / outlined_text pages,
which currently fall back to images. GPU: Radeon 780M via ONNX Runtime DirectML.

OBJECTIVE
Make hard pages editable with local models, running on the iGPU with CPU
fallback, and use a layout model to resolve ambiguous regions on any page.

DELIVERABLES
1. `pdf2pptx setup models --lite`: downloads (with checksums, resumable) the
   ONNX models into %LOCALAPPDATA%\pdf2pptx\models: RapidOCR det/rec/cls
   (Latin + multilingual), PP-DocLayout (Apache-2.0), LaMa inpainting.
   Record exact versions, licenses and sha256 in models/manifest.json. Ask
   before downloading more than 1 GB total.
2. ml/runtime.py: one ONNX session factory (DirectML → CPU fallback), warmup,
   batching, timing; `pdf2pptx doctor` reports which provider each model uses.
3. ml/ocr.py: OCR on scanned and broken_text pages → lines/words with boxes and
   confidence → fed into the M4 text pipeline (glyph-like records with
   source="ocr"). Deskew the page first. Windows built-in OCR
   (Windows.Media.Ocr via winrt) as a fallback engine.
   Outlined text: render the region, OCR it, and replace the paths with text
   only if the OCR confidence is high.
4. ml/layout.py: PP-DocLayout regions (title, text, table, figure, formula,
   chart, header/footer) used (a) on scanned pages for block structure and
   (b) on born-digital pages only for blocks/regions whose rule confidence is
   below a threshold. Disagreements are logged in the report.
5. ml/inpaint.py: for scanned pages, build a text mask from OCR boxes (dilated),
   inpaint it with LaMa to a clean background, set it as the slide background,
   and place editable text on top, color sampled from the original pixels.
6. Formulas: if a math model is not feasible within licensing/time, keep
   formula regions as pictures with alt text and log it as a gap.

CONSTRAINTS
- No network during convert; missing models → clear message + rule-only fallback.
- Models only on flagged pages/regions (rule 2). Measure how often they run.

CHECKS
- Scanned corpus PDFs: real_text_ratio ≥ 0.95, text_sim ≥ 0.95, ssim ≥ 0.90,
  ≤ 3 s/page median (print provider and timings).
- Born-digital corpus: scores unchanged within 1% (models rarely triggered;
  report the trigger rate).
- Forcing CPU (env var) still works, slower.
- pytest and ruff clean; model tests skip gracefully when models are absent.

REPORT
PROGRESS.md "Milestone 8" with timings per provider and trigger rates.
Commit "M8: OCR, layout model, inpainting".
```

---

## Prompt 9: Local model service, benchmark, judgment pass, charts

```text
CONTEXT
pdf2pptx. Read PLAN.md (§2 hardware: about 15–16 GB GPU-shareable memory,
memory bandwidth is the bottleneck; §3 tiers; §5 candidate models), PROGRESS.md.
Lite is complete. This milestone adds the Max-local tier.

OBJECTIVE
Run larger vision/text models locally through llama.cpp (Vulkan) and use
them where rules and small models are uncertain: judgment on ambiguous regions
and native charts from chart regions, then pick the best model stack for this
PC by benchmark.

DELIVERABLES
1. `pdf2pptx setup llm`: download the latest llama.cpp Windows Vulkan release
   into tools/llama.cpp (record version), and the candidate GGUF models (ask
   for confirmation with total size first): a document vision model
   (granite-docling-258M and/or PaddleOCR-VL-0.9B), judgment vision models
   (Qwen3-VL-8B and Qwen3-VL-30B-A3B), text models (Qwen3-30B-A3B,
   gpt-oss-20b). Check licenses; prefer newer equivalents if they exist and are
   supported by the release you downloaded; record choices in models/manifest.json.
2. llm/server.py: manage llama-server (start/stop/health, port, -ngl, --n-cpu-moe,
   flash attention, q8_0 KV cache, context size), one model at a time.
   llm/scheduler.py: queue jobs by model; run all jobs for model A across
   the document, then swap; keep a model resident when RAM > 48 GB.
3. bench/: `pdf2pptx bench` runs each candidate on a fixed set of 10 corpus
   regions/pages + 5 chart regions; measures load time, image-prompt speed,
   generation tok/s, peak RAM, accuracy against truth (structure/table/chart
   values). Writes bench/results.md and a recommended stack to config.
4. llm/schemas.py + llm/judge.py: JSON-schema constrained prompts for:
   region classification (e.g. "timeline / process / matrix / org chart /
   table / chart / decorative"), reading-order fixes, table structure
   correction, block role fixes. Inputs: cropped region image (capped
   resolution) + its IR JSON. Outputs merged into the IR with
   source="model:<name>" and confidence; on invalid output → keep the rule result.
5. llm/charts_vlm.py + analyze/charts.py: for chart regions (vector or raster),
   extract chart type, categories, series values, axis ranges and labels;
   cross-check against vector geometry when available; build a native chart
   (embedded workbook) with matching colors; fall back to the original shapes
   if the rebuilt chart's render diff is too large.
6. Tier wiring: `--tier lite|max`, auto-detect; Max runs judgment only on
   low-confidence regions by default (`--judge all` to force).

CONSTRAINTS
- All model calls are cached by (model, prompt hash, image hash).
- Watch memory: refuse to start a model that would exceed available RAM minus
  8 GB; suggest --n-cpu-moe values.
- Measured numbers only in reports; do not copy estimates from PLAN.md.

CHECKS
- `pdf2pptx bench` completes and writes results.md with a ranked table.
- Synthetic chart slides: chart type correct for 100%, series values within 2%
  of truth for ≥ 90% of points, output charts open as native charts in
  PowerPoint (COM: shape.HasChart).
- On low-confidence regions in the corpus, judgment improves the composite score
  of affected slides (show before/after); no slide gets worse than its Lite
  result (keep-best rule).
- A 30-page corpus deck in Max tier: total time reported; no memory
  over-commit (peak RAM logged).
- pytest and ruff clean; LLM tests skip when models are absent.

REPORT
PROGRESS.md "Milestone 9" with the benchmark table, chosen stack, and timings.
Commit "M9: local LLM service, benchmark, judgment, charts".
```

---

## Prompt 10: Document → deck, UI, packaging

```text
CONTEXT
pdf2pptx. Read PLAN.md (§3 `deck` mode), PROGRESS.md, bench/results.md.
Page-to-slide conversion is complete in Lite and Max tiers. This final
milestone adds rewriting a document into a deck and turns the project into a
product.

OBJECTIVE
(1) A `deck` mode that turns a long document (e.g. a 40-page report) into a real
presentation using the local text model; (2) a simple local UI; (3) a one-file
Windows build.

DELIVERABLES
1. llm/deck.py, `--mode deck [--slides N] [--audience ...] [--lang it|en|auto]`:
   - Build a document outline from the IR (sections, headings, key figures,
     tables, charts) without the model.
   - Text model (schema-constrained) plans the deck: slide list with title,
     3–5 bullet points, which source figures/tables/charts to reuse, speaker
     notes holding the detail, and source page references for every claim.
   - Validate: every bullet traceable to a source block (citations by block id);
     drop or flag anything not grounded. No invented numbers: numbers in bullets
     must appear in the source.
   - Build the slides with the existing builder + theme/masters from M7
     (or a clean default theme), then run the M5 tuner for overflow.
   - Output also deck_sources.json (slide → source pages/blocks).
2. Local UI: `pdf2pptx ui` starts a localhost-only web app (bind 127.0.0.1):
   drag-and-drop PDF, mode/tier selection, progress per page, side-by-side
   PDF/PPTX preview with the diff heatmap (reuse debug report components),
   download button, job history from the cache. No external CDN assets.
3. Packaging: PyInstaller one-folder (and one-file if it works reliably) build of
   the CLI + UI; models and llama.cpp stay out of the bundle and are fetched with
   `pdf2pptx setup`; version info; a smoke-test script that runs the built exe
   on 3 corpus files.
4. Hardening: a 300-page PDF, a PDF with 50 MB of images, an encrypted PDF
   (clear error), a corrupted PDF (clear error), and a PowerPoint-not-installed
   simulation all behave well.
5. Final documentation: README (install, setup, modes/tiers, what needs IT,
   troubleshooting), and a short docs/ARCHITECTURE.md that matches the code.

CONSTRAINTS
- Deck mode must never present information that is not in the source document.
- The UI must work fully offline and never send data anywhere.

CHECKS
- Deck mode on the 3 report-style PDFs: every bullet has a valid source
  reference; a number checker finds 0 numbers absent from the source; no text
  overflows its placeholder (COM check); slide count within ±2 of --slides.
- `pdf2pptx ui` works end to end on one file (describe what you checked).
- The packaged exe converts 3 corpus files on this machine with the same scores
  as the dev build (±0.5%).
- Hardening cases: no crash, clear messages, peak RAM logged for the 300-page PDF.
- Full corpus scoring: final table vs baseline in PROGRESS.md.
- pytest and ruff clean.

REPORT
PROGRESS.md "Milestone 10" + a final summary: scores vs baseline per category,
speed per tier, known limitations, recommended next steps (e.g. RAM upgrade to
64 GB, NPU use, math OCR). Commit "M10: deck mode, UI, packaging".
```
