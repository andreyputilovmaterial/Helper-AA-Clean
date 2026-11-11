# Helper-AA-Clean

Removes old timestamped backup sheets from AA (Excel) workbooks, pruning years of
accumulated backups down to a small, sensible set — without touching the live
"main" sheets.

Backup sheets follow the naming convention `<Name>_MMDDYY_HHMMSS`, e.g.
`AnalysisLabels_041124_131721`. Sheets without that suffix (`Main`,
`AnalysisLabels`, `ShortNames`, `AnalysisValues`, `Exclusions`, ...) are left
untouched.

## Retention rules

Applied per sheet name, so whole backup sets (AnalysisLabels + ShortNames +
AnalysisValues) are kept or removed together:

1. Keep every backup sheet from the last **6 months**.
2. Among the older sheets, keep only the single most recent one per sheet name
   per calendar quarter (Jan–Mar, Apr–Jun, Jul–Sep, Oct–Dec) per year.
3. Delete everything else.

## Requirements

- Python 3
- [openpyxl](https://pypi.org/project/openpyxl/) — `pip install openpyxl`

No live Excel automation is used: files are read and written directly with
openpyxl, so nothing blocks on a running Excel instance.

## Usage

```bash
# Preview only (recommended first)
python src/clean_aa_backups.py --dry-run path/to/file.xlsm

# Clean one file
python src/clean_aa_backups.py path/to/file.xlsm

# Clean every .xlsx/.xlsm in a folder (add -r to recurse)
python src/clean_aa_backups.py path/to/folder

# Copy each file to <file>.bak before modifying it
python src/clean_aa_backups.py --backup path/to/file.xlsm

# Change the "recent" window (default 6 months)
python src/clean_aa_backups.py --keep-months 12 path/to/file.xlsm
```

Macro-enabled `.xlsm` files are loaded with `keep_vba=True`, so VBA macros are
preserved on re-save. As with any tool that re-saves a workbook, use
`--dry-run` or `--backup` first.

## Running the AA workflow (Windows)

`src/run_aa_workbook.py` drives the AA workbook's own VBA buttons — apply
definitions, load definitions, and validation — through `Excel.Application`
(pywin32 / `win32com.client`), instead of clicking and waiting by hand.

Requirements (Windows only):

- Excel installed
- [pywin32](https://pypi.org/project/pywin32/) — `pip install pywin32`

```bash
# Preview the plan without launching Excel
python src/run_aa_workbook.py --dry-run path/to/file.xlsm

# Run all three steps, hidden, then save
python src/run_aa_workbook.py path/to/file.xlsm

# Show the Excel window and leave it open afterwards (for debugging)
python src/run_aa_workbook.py --visible --leave-open path/to/file.xlsm

# Skip a step, and/or point a step at a differently named macro
python src/run_aa_workbook.py --skip load_definitions \
    --macro-run-validation ValidateNow path/to/file.xlsm
```

The macro names are configurable placeholders (`ApplyDefinitions`,
`LoadDefinitions`, `RunValidation`) because the real VBA procedure names are
not part of this repo. Edit the `STEPS` table at the top of the script, or
override with `--macro-...` flags, to match the actual names in your workbook.
The macros must be `Public` subs taking no arguments and must not show modal
dialogs (message boxes, input boxes, or user forms), since there is no one to
click them.

## Layout

- `src/clean_aa_backups.py` — the script.
- `src/run_aa_workbook.py` — Windows-only VBA automation for the AA run steps.
- `tests/` — pytest suite for the retention logic and in-place clean/dry-run,
  plus the run-workbook step planning / wait logic.

## Development

```bash
python -m pytest
```
