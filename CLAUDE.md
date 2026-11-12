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
.venv\Scripts\pdf2pptx structure [--no-cache]          # paragraphs/levels/fonts vs synthetic truth
.venv\Scripts\pdf2pptx convert ... --text-layout lines  # the M2 one-box-per-line baseline
.venv\Scripts\pdf2pptx convert ... --tune/--no-tune --tune-rounds N   # M5 tuner (on: fidelity, balanced)
.venv\Scripts\pdf2pptx convert --corpus --mode fidelity --out-dir out/fidelity  # then score --out-dir out/fidelity
```

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
