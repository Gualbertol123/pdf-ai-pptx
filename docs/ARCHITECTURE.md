# pdf2pptx architecture

How the code is organised, as of Milestone 10. The design goals are in
[PLAN.md](../PLAN.md); measured results are in [PROGRESS.md](../PROGRESS.md).

## Pipeline

```
PDF ─► ingest ─► triage ─► ocr ─► analyze ─► [max: judge] ─► build ─► [tune] ─► finalize ─► PPTX
         │          │        │        │            │             │        │          │
      parse.py  triage.py  ml/ocr  analyze/*   llm/max_tier  build/*  verify/   structured deck,
                                                                       tuner     chart check
--mode deck:  ingest … analyze ─► llm/deck.py (outline ─► plan ─► validate ─► build ─► fit ─► audit)
```

`pipeline.py` runs the stages. Each stage reads and writes the **IR** (`ir.py`) and
caches it as `work/<sha256 of the PDF>/ir_<stage>.json`. A stage is reused only if its
cache key matches: the code version of that stage and of every stage before it
(`pipeline.STAGE_VERSIONS`, plus the installed OCR engine and models for `ocr`). The
final IR is saved as `ir_build.json`.

| Stage | Module | What it does |
|---|---|---|
| ingest | `ingest/parse.py` | pypdfium2: glyphs (font, size, colour, matrix), vector paths, images, annotations, outline, metadata. Pages are parsed in a process pool for longer documents. `open_pdf` turns missing, empty, encrypted, damaged and non-PDF inputs into an `InputError` with a clear message. A page that fails becomes a full-page picture. |
| triage | `ingest/triage.py` | Each page is classed as `born_digital`, `broken_text`, `outlined_text` or `scanned`. |
| ocr | `ml/ocr.py`, `ml/ocr_engine.py`, `ml/layout.py`, `ml/inpaint.py`, `ml/winocr.py` | Hard pages only: PP-OCR (ONNX, DirectML or CPU) or Windows OCR produce ordinary `Glyph`s (`source="ocr"`). Text is cleaned out of the background with LaMa or Telea inpainting. The PP-DocLayout regions are stored for later stages. |
| analyze | `analyze/*.py`, `fonts/*.py` | Fonts are matched to installed families. Then tables (ruled and whitespace), then text (glyphs → words → lines → paragraphs → blocks; lists, roles, XY-cut reading order), then shapes (presets, connectors, groups, SVG), then charts (vector drawings become native `Chart`s, keeping their parts as a fallback). Then the document structure: links, notes, sections, build-ups, theme, masters and layouts. Every pass is optional: one that fails is logged and skipped. |
| max (optional) | `llm/max_tier.py`, `llm/judge.py`, `llm/charts_vlm.py` | A local vision model checks low-confidence regions (roles, reading order, tables, regions, charts). A page keeps the changes only if it scores at least as well as Lite in PowerPoint (`keep_best`). |
| build | `build/writer.py`, `build/structure.py`, `build/xml_ext.py`, `build/charts.py`, `build/units.py` | IR → PPTX: python-pptx plus raw DrawingML. Covers text boxes or placeholders, presets, custom geometry, gradients, pictures, tables, native charts with embedded workbooks, the slide master and layouts, theme, timing (build animations), sections and notes. EMU appear only in `units.py`. The output is deterministic, including a zip with fixed timestamps. |
| tune | `verify/tuner.py` | Renders the deck in PowerPoint, compares each text block's line breaks and baselines with the PDF, adjusts width, spacing, size and position, and renders again. Then a pixel pass shifts shapes and pictures. It always works on the flat deck (one slide per page). |
| finalize | `pipeline._finalize`, `verify/deck.py`, `verify/chart_check.py` | Writes the structured deck. In fidelity mode the structured deck is kept only if it renders like the flat one. Native charts are render-checked, and any chart below 0.85 frame SSIM goes back to its original shapes. |

## The IR (`ir.py`)

Units are PDF points with the origin at the top left. Every element has `id`, `bbox`,
`z`, `confidence`, `source` (`pdf` | `rule` | `ocr` | `model:<name>`) and `notes`. The
element types are `Glyph`, `TextBlock` (made of `Paragraph`s and `TextRun`s), `Shape`,
`Image`, `Table`, `Chart`, `Equation` and `Group`. A `Document` holds pages, fonts,
theme, masters (with layouts) and sections. The JSON codec round-trips exactly
(`ir_version`).

## Verification and scoring (`verify/`)

- `ppt_com.py` is the only code that talks to PowerPoint. It runs a worker process with a
  heartbeat, timeouts and restart. It never quits a PowerPoint the user started, and it
  registers the instances it started under `%TEMP%\pdf2pptx\ppt_owned` so they can be
  reaped after a crash. `PDF2PPTX_NO_POWERPOINT=1` simulates a PC without PowerPoint.
- `score.py` and `metrics.py` compute the PLAN.md §6 scores per slide (SSIM, text_sim,
  word_recall, pos_err, linebreak_match, real_text_ratio, native_shape_ratio, composite).
  Without PowerPoint, the scores that need it are `null`.
- `diff.py` aligns the renders, builds heatmaps and lists mismatches. `report/debug_html.py`
  writes the self-contained HTML report (PDF | PPTX | heatmap, IR overlays). The UI uses
  the same report as its preview.

## Model layer (optional)

| Part | Module | Notes |
|---|---|---|
| ONNX models (Lite) | `ml/models.py`, `ml/runtime.py` | Pinned manifest with sha256. Installed by `pdf2pptx setup models` into `%LOCALAPPDATA%\pdf2pptx\models`. DirectML with CPU fallback (`PDF2PPTX_ML_DEVICE`). |
| llama.cpp (Max) | `llm/setup.py`, `llm/server.py`, `llm/gguf.py` | `pdf2pptx setup llm` installs llama.cpp (Vulkan) and GGUF candidates. One `llama-server` runs at a time on 127.0.0.1. Output is forced by a JSON schema. A memory guard estimates the need from the GGUF tensor table or measured peaks, and keeps an 8 GB reserve. |
| Scheduler + cache | `llm/scheduler.py` | All jobs of one model run before the next model loads. Every call is cached in `work/llm_cache`, keyed by model weights, prompt, images and schema. |
| Benchmark | `llm/bench.py` | Picks the stack (`pdf2pptx.toml`: `llm_judge`, `llm_chart`, `llm_text`, `llm_doc`). |

No network is used at conversion time. Downloads happen only in `pdf2pptx setup …`.

## Deck mode (`llm/deck.py`)

1. **Outline** (rules): sections from heading paragraphs (size, weight or role). Text
   blocks keep their IR ids and pages. Tables, charts and large pictures get their
   captions, and a caption counts as part of its visual's text. Footers, page numbers
   and master furniture are skipped.
2. **Plan** (text model, schema-constrained). The source is split into contiguous parts
   that each fit one call (about 22,000 characters). Slides are allocated to parts in
   proportion to their length, so the whole document is covered. A document with more
   parts than slides is first condensed chunk by chunk into grounded key points (`k<n>`
   ids that map to block ids). Each slide gets a title, 3–5 bullets citing block ids, an
   optional visual id, and speaker notes.
3. **Validate** (`Grounder`):
   - A bullet without a valid source is regrounded by text similarity or dropped.
   - A bullet with a number missing from its cited blocks either cites the block that
     has it or is dropped.
   - Notes and subtitle sentences with numbers missing from the document, or with
     words mostly not in it, are dropped.
   - Titles with invented numbers fall back to the section heading.
   - Anything dropped is listed in `deck_sources.json`.
4. **Build**: a clean master with Title, Content, Two Content and Title Only layouts, the
   source theme, and placeholders. Reused visuals are deep copies of the IR elements,
   scaled and moved with all their geometry (`_xform`).
5. **Fit**: the deck is written, PowerPoint lays the text out (COM), and any text box
   whose text runs past its bottom is set smaller, for up to 4 rounds.
6. **Audit**: the PPTX is read back. Every bullet must be listed with valid source ids,
   every number on the slides and in the notes must appear in the PDF text layer or the
   IR, there must be no overflow, and the slide count must be within 2 of the target.

Without a text model the plan is extractive: the first sentences of each section's
paragraphs (`source="rule"`).

## UI (`ui/server.py`, `ui/index.html`)

The UI runs on stdlib `http.server`, bound to 127.0.0.1. It refuses foreign `Host`
headers, requires `X-Requested-With: pdf2pptx` and a local `Origin` on every POST, and
sends a CSP that blocks network assets. Jobs run one at a time in a worker thread. Each
job is persisted as `work/ui/jobs/<id>/job.json`, and uploads are stored as
`work/ui/uploads/<sha256>.pdf`. Progress comes from `progress.report(stage, done,
total)`, which parse, ocr, analyze, write, tune, check and the model scheduler all
call. The preview is the debug report, or for deck mode the rendered slides with each
bullet's sources. The history lists the UI jobs and the documents in the conversion
cache.

## Packaging (`packaging/`)

- `pdf2pptx.spec` is a PyInstaller spec with one-folder and one-file variants.
- `entry.py` calls `multiprocessing.freeze_support()` before the CLI, because the COM
  worker and the parse pool re-launch the exe.
- In the packaged exe, `config.PROJECT_ROOT` is `%LOCALAPPDATA%\pdf2pptx` (work cache,
  `tools\llama.cpp`, `pdf2pptx.toml`). `PDF2PPTX_HOME` overrides it.
- Models and llama.cpp are not bundled.
- `smoke_test.py` converts 3 corpus files with the exe and the dev build and compares
  their scores. `scripts/hardening.py` runs the hardening cases.

## Invariants

- Rules first, models second. Every model decision records `source` and `confidence`.
- Everything optional degrades: without PowerPoint there is no tuner, no checks and no
  fit loop. Without models there is no OCR, no judgment and the deck is extractive.
  Without a GPU everything runs on the CPU.
- Deterministic: stable ids, sorted iteration, fixed seeds, LF line endings, fixed zip
  timestamps.
- Privacy: documents never leave the PC, and `corpus/real/` is not in git.
