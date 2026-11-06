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

## Layout

- `src/clean_aa_backups.py` — the script.
- `tests/` — pytest suite for the retention logic and in-place clean/dry-run.

## Development

```bash
python -m pytest
```
