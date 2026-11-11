# Helper-AA-Clean — Technical Documentation

Stable technical reference for the Helper-AA-Clean tools.

- `README.md` — entry-level overview.
- This file — full technical documentation (stable).
- `docs/project-context.md` — living working memory; it changes every task and
  is not a stable reference.

## Overview

Two standalone Python packages (no build step; run directly from `src/`):

| Package | Purpose | Runtime |
| --- | --- | --- |
| `src/clean_aa_backups/` | Prune old timestamped backup sheets from AA workbooks | openpyxl (cross-platform) |
| `src/run_aa_workbook/` | Run the workbook's own VBA actions (reload / commit-save / validate / auto) | Excel COM via pywin32 (Windows only) |

## AA workbook format

The target workbook is a macro-enabled `.xlsm`. It contains live ("main") sheets
that are never touched, and timestamped backup sheets:

- Live sheets: `Main`, `AnalysisLabels`, `ShortNames`, `AnalysisValues`,
  `Exclusions`, `SL_CategoryMAP`, `SL_MAPPING`, `AnalysisValidation`.
- Backup sheets: `<Name>_MMDDYY_HHMMSS`, e.g. `AnalysisLabels_041124_131721`. A
  backup is a triplet of sheets (`AnalysisLabels`, `ShortNames`,
  `AnalysisValues`) whose timestamps differ by a few seconds because each sheet
  is stamped as it is written.

## Tool: clean_aa_backups

Removes old backup sheets while leaving live sheets untouched. Retention is
applied per sheet name, so whole backup triplets are kept or removed together:

1. Keep every backup sheet from the last **6 months** (configurable).
2. Among older sheets, keep the single most recent one per sheet name per
   calendar quarter (Jan–Mar, Apr–Jun, Jul–Sep, Oct–Dec) per year.
3. Delete everything else.

The 6-month rule runs first, so a sheet kept by it is never removed even if it
is not the most recent for its name in its quarter.

```bash
python src/clean_aa_backups/clean_aa_backups.py [options] <paths...>
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

## Tool: run_aa_workbook

Runs the AA workbook's own VBA buttons — reload, commit/save, and validation —
through `Excel.Application`, instead of clicking and waiting.  A synthetic
`auto` action runs a full cycle of commit/save -> reload -> commit/save.

Workflow:

1. Copy the input to a working copy in the same directory
   (`<name>-working-automated-YYYY-MM-DD.xlsm`), overwriting any existing
   copy.  Every later step operates on that copy, never the original.
2. Create Excel via `win32com.client.DispatchEx("Excel.Application")` (hidden
   unless `--visible`).
3. Open the workbook (`DisplayAlerts` off, link updates off).
4. Pre-step #0: read the external data-file path from cell `C13` on the
   `Main` sheet into the module global `DATA_FILE_PATH` and verify the file
   physically exists (fail fast if missing).
5. Run each step's macro via `Application.Run("'file.xlsm'!Macro")`, in order,
   while a background thread watches for the step's message boxes and clicks
   the right button (OK / Cancel), recording validation-error text.
6. Wait for Excel to go idle (`Application.Ready` and not `Application.Busy`)
   after each macro, then an optional settle delay.
7. After each successful step, save the workbook explicitly so its on-disk
   state is clean, then close — or leave open with `--leave-open`.

`Application.Run` is synchronous; the idle poll is a safety net for macros that
hand work off asynchronously.

### Message boxes and pass/fail detection

The macros show modal `MsgBox` prompts.  A `DialogWatcher` thread matches them
by window title/text and clicks the OK/Cancel button to dismiss them, retrying
while the dialog is still open and recording the text of the dialogs that
matter.  Standard Win32 message boxes (window class `#32770`) are clicked via
`BM_CLICK` / `WM_COMMAND`; VBA UserForms (whose buttons do not carry
`IDOK`/`IDCANCEL` and ignore `BM_CLICK`) are handled by locating the button
child window and sending it a real mouse click.

| Action | Success dialog | Failure dialog |
| --- | --- | --- |
| reload | (none) | (none) |
| commit/save | `Save Successful` / "Your Analysis Specifications have been saved" | `Validation Error` / "An error occured while processing your specifications." |
| validate | "Validation Complete! No Errors found!" | `Validation Error` / "An error occured while processing your specifications." |

Success/failure is judged from which dialog was seen:

- commit/save succeeds only when the success dialog was shown **and** the data
  file's modification time advanced during the save (i.e. it was recently
  edited).  Otherwise the run fails.
- validate succeeds only when its success dialog was shown.
- Any validation failure raises `ValidationError` carrying the dialog text and
  the full `AnalysisValidation` sheet, printed in red with a non-zero exit
  code.

The `auto` action runs commit/save -> reload -> commit/save with these rules:

- first commit/save — normal rules (validation failure or any other error is
  a failure);
- reload — any error is a failure;
- second commit/save — its expected `Validation Error` is the success signal;
  any other error is a failure.

Any other modal prompt (not listed above) still blocks the run, because there
is nobody to click it.

### Logging

Every action that actually runs appends one row to a CSV log
(`aa_run_log.csv`) in the same directory as the AA file (the original input,
not the working copy):

| Column | Meaning |
| --- | --- |
| `datetime` | when the action was invoked |
| `program` | `commit/save`, `(re)load`, `run validation`, or `auto` |
| `aa_file` | path to the AA Excel file |
| `data_file` | path to the data file (from `Main`!C13) |
| `result` | `success` or the exception message |
| `validation_details` | message box text plus `AnalysisValidation` sheet when a validation failed |

The file is created on first use and appended to afterwards.  For `auto`, the
`validation_details` column holds the first validation failure in the cycle:
the first commit/save's failure stops the run, otherwise the second
commit/save's expected failure, or empty if none failed.

### Working copy

Before launching Excel, the script copies the input next to itself as
`<name>-working-automated-YYYY-MM-DD.xlsm` (same extension) and runs only
against that copy.  The original is never opened or saved, so the macros'
repeated saves (and any crash mid-run) leave the original untouched.  A
previous working copy from the same day is overwritten.

Requirements (Windows only):

- Excel installed.
- [pywin32](https://pypi.org/project/pywin32/) —
  `pip install -r requirements-run-aa.txt`.

The `win32com` import is lazy, so `--help` and the test suite run without
pywin32 (e.g. on macOS).

```bash
python src/run_aa_workbook/run_aa_workbook.py --program {reload,commit_save,validate,auto} [options] <path>
```

| Option | Description |
| --- | --- |
| `path` | Excel file (`.xlsx` / `.xlsm`) to open and run |
| `--program KEY` | which program to run; exactly one per run (required) |
| `--visible` | show the Excel window (default hidden) |
| `--leave-open` | keep Excel/workbook open (debugging) |
| `--idle-timeout S` | seconds to wait for Excel to go idle (default 300) |
| `--settle S` | extra seconds after idle (default 1.0) |

Program keys: `reload`, `commit_save`, `validate`, `auto` (`auto` expands to
`commit_save` -> `reload` -> `commit_save`).  Exactly one is selected per
invocation via `--program`.

### Macro configuration

The three macros are the real VBA entry points behind the AA buttons, in one
`STEPS` table at the top of the script:

| Step key | Default macro | VBA module |
| --- | --- | --- |
| `reload` | `ListAll` | `Simplified` |
| `commit_save` | `SaveAll` | `Simplified` |
| `validate` | `Button1_Click` | `Addon` |

`auto` is not a macro; it is a synthetic action expanded into
`commit_save` -> `reload` -> `commit_save`.

The full VBA source of these three macros is captured in
`docs/aa-vba-macros.md`.

### Limitations

- The macros must be `Public` subs.  `ListAll` is called with no arguments, so
  its `Optional RJob` is `IsMissing` and the overwrite prompt is shown.
- Only the message boxes listed above are handled; any other modal UI
  (`InputBox`, user forms, or an unknown `MsgBox`) will hang the script
  because there is nobody to click it.

## Development

```bash
python -m pytest
```

Tests cover pure logic only (retention planning, step planning, idle-wait); no
COM/Excel needed, so the suite runs on any OS.
