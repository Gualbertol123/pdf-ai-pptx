# pdf2pptx: notes for Claude Code sessions

Local, offline PDF → PPTX converter. Design: `PLAN.md`. Build prompts: `PROMPTS.md`
(one prompt per fresh session, in order). What's done and measured: `PROGRESS.md`. Read
the latest PROGRESS.md entry before starting a prompt.

## Commands (Windows, PowerShell, from the repo root)

```powershell
.venv\Scripts\pytest -q                      # ~3 min with PowerPoint; COM tests skip without it
.venv\Scripts\ruff check .
.venv\Scripts\mypy                           # strict on ir.py and build/units.py
.venv\Scripts\pdf2pptx doctor
.venv\Scripts\pdf2pptx corpus build          # ~60 s, deterministic, regenerates corpus/synthetic
.venv\Scripts\pdf2pptx score --corpus --against truth|blank|out --label NAME
```

Scores go to `work/scores/<label>/` (git-ignored). `work/scores/truth` and `blank` hold
the M1 metric sanity runs; re-run them if missing.

## Binding rules (PLAN.md §8)

- Measure before and after every milestone; record both in PROGRESS.md. No silent
  regression above 1%.
- Rules first, models second. Every model decision records `source` and `confidence`.
- Everything optional degrades gracefully (no PowerPoint / models / GPU → fallback, never a crash).
- Deterministic output: stable ids, sorted iteration, fixed seeds, LF line endings.
- Points internally; EMU only in `src/pdf2pptx/build/units.py`.
- PowerPoint COM only through `verify/ppt_com.py` (until Prompt 3 replaces it, the
  temporary `verify/_com_temp.py`). Never quit or kill a PowerPoint the user started.
- No network at conversion time; downloads only in `pdf2pptx setup …`.
- Privacy: never upload documents; `corpus/real/` stays out of git.
- No AGPL dependencies (no PyMuPDF, Ultralytics, MinerU).

## Gotchas learned so far

- Write text files with `newline="\n"` (Windows Python defaults to CRLF; `.gitattributes`
  enforces LF). Avoid shell heredocs for Python source containing `\r`/`\n` escapes.
- COM paths must be normalised backslash paths (`os.path.normpath`).
- Office runs in an Italian locale: COM error messages are in Italian.
- pytest may print `Windows fatal exception: code 0x80010108` after COM tests; it is a
  handled disconnect, not a failure.
- More measured COM facts (text-layout boxes, table-cell offsets): PROGRESS.md, M1.
