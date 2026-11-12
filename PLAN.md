# pdf2pptx: Project Plan

A local, offline PDF → PPTX converter that rebuilds **native, editable PowerPoint objects** (real text with paragraphs and bullets, built-in shapes, tables, charts, slide masters, theme, animations) and checks its own output by rendering it in PowerPoint and correcting mismatches automatically.

It runs entirely on this PC. No document ever leaves the machine.

---

## 1. Goals

| # | Goal | How we measure it |
|---|---|---|
| G1 | The slide looks like the PDF page | Median visual similarity (SSIM) per slide ≥ 0.95 on born-digital PDFs |
| G2 | The content can be edited | ≥ 98% of visible text is real text. ≥ 80% of simple shapes are built-in shapes. Few text boxes per slide. |
| G3 | Text flows the way it should | ≥ 98% of PowerPoint line breaks match the PDF's |
| G4 | It understands the document's structure | Tables, charts, masters, theme, sections and build animations are rebuilt natively |
| G5 | It handles hard inputs | Scanned or broken PDFs come out as editable text on a clean background |
| G6 | It's as close to "Max" as this PC allows | Local vision and language models for judgment calls, chart data and document → deck rewriting |

**Non-goals (for now):** cloud services, fine-tuning models, recreating SmartArt, real-time Max mode, macOS/Linux support.

---

## 2. Target machine (measured 2026-11-11)

| Component | Value | Design consequence |
|---|---|---|
| Machine | Lenovo ThinkPad 21MD (T14 Gen 5 AMD class) | Corporate PC: drivers and BIOS may need IT |
| CPU | Ryzen 7 PRO 8840U, 8 cores / 16 threads | Process about 6–8 pages in parallel |
| RAM | 2 × 16 GB DDR5-5600 (2 SODIMM slots) | About 18–20 GB free for one large model. **Upgrading to 64 GB is the best hardware step.** |
| GPU | Radeon 780M (integrated), 1 GB dedicated video memory | The GPU can borrow about 15–16 GB of RAM. Memory bandwidth (about 60–70 GB/s) is the bottleneck. |
| NPU | AMD IPU (first-generation XDNA) | Not used in v1 |
| GPU driver | 31.0.22056.6001 (March 2024) | Update it for better Vulkan performance |
| Office | PowerPoint installed | **Reference renderer through COM automation** |
| Python | 3.12.7 | Project runtime |
| Not installed | LibreOffice, Tesseract, Ollama | Not needed |

Set the Windows power mode to **Best performance**, and stay plugged in during Max runs.

---

## 3. Tiers and modes

**Tiers** are picked automatically from hardware and settings:

| Tier | Adds | Rough speed here |
|---|---|---|
| Lite | Rules, layout model, OCR, background cleanup, PowerPoint check loop | about 1–3 s per page |
| Max-local | A small document-specialist vision model, plus a larger vision model for judgment calls and chart data, plus a text model for document → deck | Minutes per deck (queued, can run overnight) |

**Modes**, a fidelity ↔ editability slider:

- `fidelity`: matches the PDF as closely as possible, even if more pieces end up in fixed positions
- `balanced` (default)
- `editable`: clean placeholders, theme-based styling and reflowable text, with small visual differences allowed
- `deck` (Max only): rewrites the document into a real presentation

---

## 4. Architecture

```
PDF ──► 1 INGEST ──► 2 TRIAGE ──► 3 ANALYZE ──► 4 BUILD ──► 5 VERIFY/TUNE ──► PPTX
        pdfium        per page:     rules first;   python-pptx   PowerPoint COM:
        glyphs,       born-digital  models only    + raw XML     render, compare,
        paths,        / broken      on flagged                   read line breaks,
        images,       text layer /  regions                      adjust, repeat
        renders       scanned
                           │
                    ┌──────▼───────────────────────────────┐
                    │ Model layer (optional per tier)      │
                    │ ONNX Runtime with DirectML: layout,  │
                    │   OCR, background cleanup            │
                    │ llama.cpp (Vulkan) server: document  │
                    │   vision model, judgment vision      │
                    │   model, text model                  │
                    │ Scheduler: ONE large model loaded    │
                    │   at a time, run on all pages        │
                    └──────────────────────────────────────┘
```

Every stage reads and writes the **intermediate representation (IR)** and saves it to `work/<hash>/` as JSON. Stages can then be cached, tested and inspected on their own.

### 4.1 Intermediate representation (IR, `src/pdf2pptx/ir.py`)

- Units: **PDF points, origin at the top left**, y pointing down. Convert to EMU (1 pt = 12 700 EMU) only in BUILD.
- Common fields on every element: `id`, `bbox (x0,y0,x1,y1)`, `z`, `confidence (0–1)`, `source` (`pdf` | `rule` | `ocr` | `model:<name>`), `notes`.
- Types:
  - `Document`: pages, metadata (producer, creator), fonts, theme, masters, sections
  - `Page`: index, size, rotation, triage result, elements, render paths
  - `Glyph` and `TextRun`: text, font reference, size, color, bbox, baseline, rotation
  - `Paragraph`: runs, alignment, indent, level, bullet or numbering, line spacing, space before and after
  - `TextBlock`: paragraphs, role (title, body, footer, caption, page number…), autofit policy
  - `Shape`: preset (with adjust values) or custom path, fill (solid, gradient), stroke, opacity, rotation, connector endpoints
  - `Image`: file, crop, mask, opacity
  - `Table`: rows, columns, cells (row and column span, text blocks, fill, borders)
  - `Chart`: type, categories, series, styling, fallback element ids
  - `Equation`: LaTeX, output as PowerPoint math (OMML)
  - `Group`: child ids
- The IR is versioned (`ir_version`), and every saved JSON round-trips exactly.

### 4.2 Repository layout

```
pdf_pptx/
├─ PLAN.md  PROMPTS.md  PROGRESS.md  README.md  pyproject.toml
├─ src/pdf2pptx/
│  ├─ cli.py  config.py  ir.py  pipeline.py  cache.py  log.py
│  ├─ ingest/   parse.py  render.py  triage.py
│  ├─ analyze/  text_group.py reading_order.py lists.py shapes.py tables.py
│  │            charts.py masters.py theme.py builds.py links.py
│  ├─ fonts/    index.py  match.py  embed.py
│  ├─ ml/       runtime.py  ocr.py  layout.py  inpaint.py
│  ├─ llm/      server.py  scheduler.py  schemas.py  judge.py  charts_vlm.py  deck.py
│  ├─ build/    writer.py  xml_ext.py  units.py
│  ├─ verify/   ppt_com.py  diff.py  tuner.py  metrics.py
│  └─ report/   debug_html.py
├─ corpus/
│  ├─ synthetic/   (generated decks with ground truth)
│  └─ real/        (user PDFs, git-ignored if sensitive)
├─ bench/   tests/   tools/ (llama.cpp binaries, git-ignored)
└─ work/  (cache, git-ignored)
Models live in %LOCALAPPDATA%\pdf2pptx\models (git-ignored, downloaded on first run)
```

---

## 5. Components and dependencies

Only permissive licenses (MIT, BSD, Apache-2.0). **No AGPL**: no PyMuPDF, no Ultralytics/DocLayout-YOLO, no MinerU.

| Area | Choice |
|---|---|
| PDF parsing and rendering | `pypdfium2`, `pdfplumber` (for cross-checking) |
| PPTX writing | `python-pptx` plus `lxml` for raw XML (gradients, animations, sections, charts) |
| PowerPoint automation | `pywin32` (COM) in its own worker process with a watchdog |
| Fonts | `fontTools` |
| Image comparison | `opencv-python-headless`, `scikit-image`, `numpy` |
| Synthetic PDFs | `python-pptx` → PowerPoint PDF export; `reportlab` for report-style pages |
| Small models | `onnxruntime-directml` with CPU fallback: PP-DocLayout, RapidOCR, LaMa |
| Large models | `llama.cpp` Vulkan build (`llama-server`), with output forced to match a JSON schema |
| Candidate large models | Document vision: granite-docling-258M, PaddleOCR-VL-0.9B. Judgment vision: Qwen3-VL-8B or Qwen3-VL-30B-A3B. Text: Qwen3-30B-A3B or gpt-oss-20b. **Final choice made by benchmark.** |
| Tooling | `pytest`, `ruff`, `mypy` (for `ir.py` and `build/`) |

---

## 6. Scores (`verify/metrics.py`)

Computed per slide and per document, and stored as JSON plus a table:

| Score | Definition |
|---|---|
| `ssim` | Grayscale SSIM between the PDF page and the PowerPoint render at 150 DPI, same size |
| `region_ssim_min` | Worst score over a 4×4 grid of page regions (catches problems in one spot) |
| `text_sim` | Normalized Levenshtein similarity between PDF text and PPTX text, in reading order |
| `word_recall` | Share of PDF words found in the PPTX |
| `pos_err_pt` | Median distance between the centers of matched words. PPTX word positions come from COM `TextRange.Words(i).Bound*`. |
| `linebreak_match` | Share of PDF lines whose break PowerPoint reproduces |
| `real_text_ratio` | Characters in text frames ÷ all visible characters |
| `native_shape_ratio` | Built-in shapes ÷ all vector shapes |
| `boxes_per_slide` | Number of text boxes per slide (lower is better, for the same fidelity) |
| `seconds_per_page` | Wall-clock time |
| `composite` | 0.4·ssim + 0.2·text_sim + 0.2·real_text_ratio + 0.1·linebreak_match + 0.1·native_shape_ratio |

The **synthetic corpus** has exact ground truth: decks generated with python-pptx, exported to PDF by PowerPoint, then converted back. The original deck is the answer key.

---

## 7. Milestones (each one maps to a prompt in PROMPTS.md)

| # | Milestone | Done when |
|---|---|---|
| 1 | Foundation, IR, corpus, scoring harness | One command scores the whole corpus |
| 2 | Ingest and baseline converter | Every corpus file converts. Baseline scores recorded. |
| 3 | PowerPoint check engine and debug report | Any PPTX is scored in under about 1 s per slide. HTML overlay report works. |
| 4 | Text understanding and fonts | Text boxes per slide ↓ ≥ 5× vs baseline. word_recall ≥ 0.99 on born-digital. |
| 5 | Fidelity tuner | Median ssim ≥ 0.95 and linebreak_match ≥ 0.98 on synthetic |
| 6 | Graphics and tables | native_shape_ratio ≥ 0.8. Tables come out as native tables. |
| 7 | Document structure | Masters, theme, sections, notes, links, build animations recovered |
| 8 | Hard pages (OCR, layout model, background cleanup) | Scanned pages become editable at under about 3 s per page |
| 9 | Local model service, benchmark, judgment pass, charts | Benchmark picks the stack. Low-confidence regions improve. Charts are native. |
| 10 | Document → deck, UI, packaging | `deck` mode works. Local UI runs. A single .exe converts a 30-page deck. |

---

## 8. Engineering rules (every prompt follows these)

1. **Measure first.** Run the corpus scoring before and after every milestone, and record both in `PROGRESS.md`. Never accept a regression above 1% on any score without saying why.
2. **Rules first, models second.** Models only run on regions that are flagged or low-confidence. Every model decision records its `source` and `confidence` in the IR.
3. **Everything is optional and degrades gracefully.** Missing PowerPoint, models or GPU must lead to a working fallback (fewer features), never a crash.
4. **Deterministic output.** Same input and settings give byte-identical output: stable ids, sorted iteration, fixed seeds.
5. **Units:** points internally, EMU only in `build/units.py`.
6. **PowerPoint COM** only through `verify/ppt_com.py`: separate process, single thread, timeouts, automatic restart, cleanup of stuck `POWERPNT.EXE` processes started by us only.
7. **Memory:** render pages as they're processed. At most one large model loaded at a time (unless more than 48 GB RAM is detected).
8. **Tests:** every analyzer gets unit tests on small hand-made pages, plus at least one corpus-level check. `pytest -q` and `ruff check` must pass.
9. **No network at conversion time.** Downloads only happen in explicit setup commands (`pdf2pptx setup …`).
10. **Privacy:** never upload documents. Keep `corpus/real/` out of git.
11. Update `PROGRESS.md` at the end of every milestone: what was built, scores before and after, known gaps, next steps.

---

## 9. Risks

| Risk | Mitigation |
|---|---|
| PowerPoint COM hangs or leaves stuck processes | Worker process with heartbeat, a timeout per call and a kill-and-restart path |
| python-pptx is missing features | Raw XML helpers in `build/xml_ext.py`, each checked by opening the file in PowerPoint |
| Unusual PDFs (Type 3 fonts, blend modes, clipping) | Triage step, falling back to an image for that region, logged in the report |
| Font licenses | Check each font's embedding permission. Otherwise substitute the closest metric match. |
| DirectML or Vulkan driver problems | Automatic CPU fallback. A `doctor` command reports what works. |
| Corporate restrictions (driver, BIOS, installs) | Everything installs per user. The README documents what needs IT. |
| Model output drifts | Output forced by JSON schema, plus validation, plus a fallback to the rule-based result |
