# Helper-AA-Clean — Technical Documentation

Stable technical reference for the Helper-AA-Clean tools.

- `README.md` — entry-level overview.
- This file — full technical documentation (stable).
- `docs/project-context.md` — living working memory; it changes every task and
  is not a stable reference.

## Overview

Two standalone Python scripts (no build step; run directly from `src/`):

| Script | Purpose | Runtime |
| --- | --- | --- |
| `src/clean_aa_backups.py` | Prune old timestamped backup sheets from AA workbooks | openpyxl (cross-platform) |
| `src/run_aa_workbook.py` | Run the workbook's own VBA workflow steps | Excel COM via pywin32 (Windows only) |

## AA workbook format

The target workbook is a macro-enabled `.xlsm`. It contains live ("main") sheets
that are never touched, and timestamped backup sheets:

- Live sheets: `Main`, `AnalysisLabels`, `ShortNames`, `AnalysisValues`,
  `Exclusions`, `SL_CategoryMAP`, `SL_MAPPING`, `AnalysisValidation`.
- Backup sheets: `<Name>_MMDDYY_HHMMSS`, e.g. `AnalysisLabels_041124_131721`. A
  backup is a triplet of sheets (`AnalysisLabels`, `ShortNames`,
  `AnalysisValues`) whose timestamps differ by a few seconds because each sheet
  is stamped as it is written.

## Tool: clean_aa_backups.py

Removes old backup sheets while leaving live sheets untouched. Retention is
applied per sheet name, so whole backup triplets are kept or removed together:

1. Keep every backup sheet from the last **6 months** (configurable).
2. Among older sheets, keep the single most recent one per sheet name per
   calendar quarter (Jan–Mar, Apr–Jun, Jul–Sep, Oct–Dec) per year.
3. Delete everything else.

The 6-month rule runs first, so a sheet kept by it is never removed even if it
is not the most recent for its name in its quarter.

```bash
python src/clean_aa_backups.py [options] <paths...>
```

| Option | Description |
| --- | --- |
| `paths` | Excel file(s) or folder(s) to clean (`.xlsx` / `.xlsm`) |
| `-n`, `--dry-run` | print what would be deleted; do not modify files |
| `--keep-months N` | keep snapshots newer than N months (default 6) |
| `-r`, `--recursive` | scan subfolders of a given folder |
| `--backup` | copy each file to `<file>.bak` before modifying it |

Behavior notes:

- `.xlsm` files load with `keep_vba=True`, so macros survive re-save.
- openpyxl re-saves the whole workbook; unusual elements (some charts/images/form
  controls) may not round-trip — `--dry-run` / `--backup` are the safety net.

## Tool: run_aa_workbook.py

Runs the AA workbook's own VBA buttons — apply definitions, load definitions,
and validation — through `Excel.Application`, instead of clicking and waiting.

Workflow:

1. Create Excel via `win32com.client.DispatchEx("Excel.Application")` (hidden
   unless `--visible`).
2. Open the workbook (`DisplayAlerts` off, link updates off).
3. Run each step's macro via `Application.Run("'file.xlsm'!Macro")`, in order.
4. Wait for Excel to go idle (`Application.Ready` and not `Application.Busy`)
   after each macro, then an optional settle delay.
5. Save (unless `--read-only` / `--no-save`) and close, or leave open with
   `--leave-open`.

`Application.Run` is synchronous; the idle poll is a safety net for macros that
hand work off asynchronously.

Requirements (Windows only):

- Excel installed.
- [pywin32](https://pypi.org/project/pywin32/) — `pip install pywin32`.

The `win32com` import is lazy, so `--help`, `--dry-run`, and the test suite run
without pywin32 (e.g. on macOS).

```bash
python src/run_aa_workbook.py [options] <path>
```

| Option | Description |
| --- | --- |
| `path` | Excel file (`.xlsx` / `.xlsm`) to open and run |
| `--visible` | show the Excel window (default hidden) |
| `--read-only` | open read-only and never save |
| `--no-save` | do not save after the run |
| `--leave-open` | keep Excel/workbook open (debugging) |
| `--dry-run` | print the plan without launching Excel |
| `--skip STEP` | skip a step; repeatable |
| `--macro-<step> NAME` | override a step's VBA macro name |
| `--idle-timeout S` | seconds to wait for Excel to go idle (default 300) |
| `--settle S` | extra seconds after idle (default 1.0) |

Step keys: `apply_definitions`, `load_definitions`, `run_validation`.

### Macro configuration

The three macros are placeholders in one `STEPS` table at the top of the script:

| Step key | Default macro |
| --- | --- |
| `apply_definitions` | `ApplyDefinitions` |
| `load_definitions` | `LoadDefinitions` |
| `run_validation` | `RunValidation` |

The real VBA procedure names are not part of this repo; edit the table or pass
`--macro-<step>` to match the actual names.

### Limitations

- The macros must be `Public` subs taking no arguments.
- They must not show modal UI (`MsgBox`, `InputBox`, user forms); there is no
  one to click them, so the script would hang.

## Development

```bash
python -m pytest
```

Tests cover pure logic only (retention planning, step planning, idle-wait); no
COM/Excel needed, so the suite runs on any OS.
