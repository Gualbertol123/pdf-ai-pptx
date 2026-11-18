# pdf2pptx: notes for Claude Code sessions

Local, offline PDF → PPTX converter. Design: `PLAN.md`. Build prompts: `PROMPTS.md`
(one prompt per fresh session, in order). What's done and measured: `PROGRESS.md`. Read
the latest PROGRESS.md entry before starting a prompt.

## Commands (Windows, PowerShell, from the repo root)

```powershell
.venv\Scripts\pytest -q                      # ~1.5 min with PowerPoint; COM tests skip without it
.venv\Scripts\ruff check .
.venv\Scripts\mypy                           # strict on ir.py and build/units.py
.venv\Scripts\pdf2pptx doctor
.venv\Scripts\pdf2pptx corpus build          # ~60 s, deterministic, regenerates corpus/synthetic
.venv\Scripts\pdf2pptx convert in.pdf -o out.pptx      # IR cached in work/<sha256>/ir_<stage>.json
.venv\Scripts\pdf2pptx convert --corpus --no-cache     # ~12 s -> out/corpus, + PowerPoint open check
.venv\Scripts\pdf2pptx score --corpus --against truth|blank|out --label NAME [--as-baseline]
.venv\Scripts\pdf2pptx report a.pdf [b.pdf ...]          # -> out/report/<stem>.html (+ index.html)
.venv\Scripts\pdf2pptx structure [--no-cache]          # paragraphs/levels/fonts + presets/tables vs truth
.venv\Scripts\pdf2pptx convert ... --text-layout lines  # the M2 one-box-per-line baseline
.venv\Scripts\pdf2pptx convert ... --tune/--no-tune --tune-rounds N   # M5 tuner (on: fidelity, balanced)
.venv\Scripts\pdf2pptx convert --corpus --mode fidelity --out-dir out/fidelity  # then score --out-dir out/fidelity
.venv\Scripts\pdf2pptx convert ... --builds/--no-builds --structure/--no-structure  # M7 deck structure
.venv\Scripts\pdf2pptx deck [--out-dir D]     # masters / slide numbers / builds / theme / sections / notes vs truth
.venv\Scripts\pdf2pptx setup models --lite   # M8 ONNX models (~346 MB, pinned + sha256) -> %LOCALAPPDATA%\pdf2pptx\models
.venv\Scripts\pdf2pptx setup models --status --verify
.venv\Scripts\pdf2pptx doctor [--no-models]  # provider per model (loading LaMa takes ~15 s)
$env:PDF2PPTX_ML_DEVICE = "cpu"   # force the CPU (auto | cpu | dml); PDF2PPTX_MODELS_DIR overrides the folder
.venv\Scripts\pdf2pptx setup llm [--only NAME ...] [--status]   # M9 llama.cpp (Vulkan) -> tools/llama.cpp + GGUF candidates
.venv\Scripts\pdf2pptx setup llm --remove-except NAME ...      # delete the candidates the bench did not pick
.venv\Scripts\pdf2pptx bench [--only NAME ...]   # LLM benchmark -> bench/results.md + stack in pdf2pptx.toml
.venv\Scripts\pdf2pptx bench --long-deck 30 --tier max          # 30-page deck in Max: total time + peak RAM
.venv\Scripts\pdf2pptx convert ... --tier lite|max --judge low|all|off   # tier auto = max when a VLM is installed
```

Scoring runs must pass `--tier lite` (or `max`) explicitly: `auto` picks Max on this machine
once a vision model is installed.

Scores go to `work/scores/<label>/` (git-ignored). `work/scores/truth` and `blank` hold
the M1 metric sanity runs; `work/scores/baseline.json` is the M2 BASELINE every later
milestone must beat. Re-run them if missing.

## Binding rules (PLAN.md §8)

- Measure before and after every milestone; record both in PROGRESS.md. No silent
  regression above 1%.
- Rules first, models second. Every model decision records `source` and `confidence`.
- Everything optional degrades gracefully (no PowerPoint / models / GPU → fallback, never a crash).
- Deterministic output: stable ids, sorted iteration, fixed seeds, LF line endings.
- Points internally; EMU only in `src/pdf2pptx/build/units.py`.
- PowerPoint COM only through `verify/ppt_com.py` (worker process; use
  `ppt_com.shared()` for the long-lived engine). Never quit or kill a PowerPoint the
  user started.
- No network at conversion time; downloads only in `pdf2pptx setup …`.
- Privacy: never upload documents; `corpus/real/` stays out of git.
- No AGPL dependencies (no PyMuPDF, Ultralytics, MinerU).

## Gotchas learned so far

- Write text files with `newline="\n"` (Windows Python defaults to CRLF; `.gitattributes`
  enforces LF). Avoid shell heredocs for Python source containing `\r`/`\n` escapes.
- COM paths must be normalised backslash paths (`os.path.normpath`).
- Office runs in an Italian locale: COM error messages are in Italian.
- COM speed (M3): property reads cost ~0.25 ms, but every *new* COM object costs
  2-10 ms when pywin32 wraps it (type-info lookup or a cross-process QueryInterface).
  Hot loops in `ppt_com` use raw `IDispatch.Invoke` with cached DISPIDs (`_RAW`) and
  skip non-text shapes using a python-pptx plan. Keep it that way.
- PowerPoint is single-instance: a second engine attaches to the first one's
  POWERPNT.EXE and registers as co-owner (`%TEMP%/pdf2pptx/ppt_owned`). Tests that
  need their own instance call `ppt_com.shutdown_shared()` first.
- Bump the stage version in `pipeline.STAGE_VERSIONS` (or `parse.PARSER_VERSION`) when a
  stage's output changes, or stale `work/<sha>/ir_*.json` caches are reused.
- pdfium quirks (M2): pattern fills report as opaque white, soft masks don't show in
  `HasTransparency`; see PROGRESS.md M2 for the workarounds.
- More measured COM facts (text-layout boxes, table-cell offsets): PROGRESS.md, M1.
- PowerPoint rounds `spcPts` (line and paragraph spacing) to whole points and `spcPct`
  to whole percent. The writer picks whichever lands closer (`build/units.ppt_spacing`)
  and the analyzer predicts positions with the same rounding. Details: PROGRESS.md M4.
- The installed-font index is cached in `%LOCALAPPDATA%\pdf2pptx\cache\font_index.json`
  (`PDF2PPTX_FONT_INDEX` overrides); a changed font file is re-read automatically.
- Tuner (M5): text boxes are named `TextBlock <block id>`; that is how COM layouts map
  back to IR blocks. `spc` is not rounded (effect ≈ 0.955 × spc); 0.1 pt font sizes
  are honoured. Do not use a ranged `ExportAsFixedFormat`: it crashed PowerPoint
  after some decks. Whole-deck `export_pdf` is stable but stalls 7–10 s now and then.
- Adding a field to a dataclass that is built positionally shifts the later
  arguments silently (it broke the tuner's `Line`); build such objects with keywords.
- While iterating on `analyze/`, run `convert`/`structure` with `--no-cache` (or bump
  `STAGE_VERSIONS["analyze"]`): the analyze cache key is the stage version only.
- Graphics/tables (M6): `analyze` runs fonts → tables → text (table glyphs skipped) →
  shapes. Table cells are `Table <id>[r,c]` in COM layouts; cell margins are derived
  from `TableCell.text.bbox` (`writer.cell_margins`). COM cell `BoundTop` is
  `(marT − marB)/2` relative to the cell; PowerPoint widens a merged cell's last column
  to hold marL + marR. Details: PROGRESS.md M6.
- PowerPoint's PDF export draws lines with arrowheads as filled outlines (body + head
  polygons); `shapes.decode_outlined` turns them back into connectors. Pattern fills
  are kept as `Fill(kind="pattern")` shapes (not drawn) next to their raster.
- Scripts that start the COM worker need an `if __name__ == "__main__":` guard
  (multiprocessing spawn). Shell heredocs choke on some quoting: write longer files
  with the editor tools instead.
- Structure (M7): pages keep every element; master / layout copies live in
  `Document.masters` and page elements carry `in_master:` / `in_layout:` notes. The
  tuner always works on the flat deck (`TextOptions.flat()`, slide = page);
  `pipeline._finalize` writes the structured deck afterwards (and, in fidelity, keeps it
  only if it renders like the flat one).
- python-pptx `_next_shape_id` counts every `@id` of the part: on a slide master that
  includes the layout ids (2147483649+), and PowerPoint refuses such shape ids. Use
  `structure.next_id` / `renumber_ids`. A master's body placeholder must be
  `type="body"` (a content placeholder there makes the file unopenable).
- COM line bounds include a paragraph's space-after, except for the frame's last
  paragraph (measured M7); the scorer subtracts it only where it is included.
- Collapsed build-ups are scored through `verify/expand.py` (one slide per click
  state). Master / layout shapes count in text stats and COM text layouts.
- Hard pages (M8): stages are ingest → triage → **ocr** → analyze. The OCR stage key
  includes which engine / models are installed (`pipeline.stage_versions`). OCR glyphs
  are ordinary `Glyph`s with `source="ocr"` (ids `p<i>_o<n>`, fonts `ocr:<Family>[-Bold]`)
  so the M4 grouping, M5 tuner and M7 structure handle them.
- Model runtime facts (Radeon 780M, ORT-DirectML 1.24): LaMa fails on DirectML (a MatMul
  of its Fourier unit) after a ~20 s compile and takes ~10-15 s just to load on the CPU:
  it is pinned to the CPU (`runtime.PREFER_CPU`) and only runs on non-uniform
  backgrounds. Warm a model up on realistic input: an all-zero mask passed LaMa's DML
  warm-up. The cls model is 10x faster on the CPU; det / rec / layout run on DML.
- The PP-OCR orientation (cls) model calls some upright lines upside down with > 0.9
  confidence: flagged lines are read both ways and the better reading wins.
- Skew: projection profiles need soft binning + smoothing (integer bins favour exactly
  0°) and per-strip scoring (unaligned columns pulled s02 0.3° off). 0.1° of skew error
  costs ~0.01 SSIM on a dense page. `text_group` writes a bucket's median exact angle.
- The scorer's reference text is the PDF text layer: recovering outlined text (d28)
  *lowers* text_sim against it although the slide now matches the truth deck.
- Charts (M9): `analyze/charts.py` runs after shapes; a `Chart` keeps the elements it
  replaced in `Chart.fallback` (`charts.restore` puts them back; the render check in
  `verify/chart_check.py` does that below 0.85 frame SSIM). Chart text is scored in
  PowerPoint's PDF reading order (data labels, ticks ascending, categories, legend);
  chart words are left out of the COM word metrics. Embedded workbooks are
  re-zipped with fixed timestamps (`writer.normalize_zip`).
- llama.cpp b12503 on the 780M: the Vulkan vision encoder crashes llama-server on
  page-sized images, so `--no-mmproj-offload` is the default (~0.5 s / image). q8_0 KV
  crashes granite-docling at start (manifest `kv_type: f16`). Images are capped at
  768 px. The memory guard reads the GGUF tensor table: CPU-side experts
  (`--n-cpu-moe`) stay memory-mapped and do not count.
- LLM calls are cached in `work/llm_cache` by (model + weights sha, prompt, images,
  schema); delete it to re-measure. Max tier keeps a page's judgment only if its
  composite is not lower than Lite's (`llm/max_tier.keep_best`).
