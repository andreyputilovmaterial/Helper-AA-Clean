# Helper-AA-Clean

Small command-line helpers for working with AA (Excel) workbooks.

- `clean_aa_backups` — removes old timestamped backup sheets (openpyxl).
- `run_aa_workbook` — runs the workbook's own VBA actions automatically
  (reload, commit/save, validate, auto_apply_aa, auto_reload_aa) on
  Windows + Excel (pywin32).

## Quick start

```bash
# Preview a cleanup before touching anything (recommended first)
python src/clean_aa_backups/clean_aa_backups.py --dry-run path/to/file.xlsm

# Run one AA action (Windows only): reload, commit_save, validate,
# auto_apply_aa, or auto_reload_aa
python src/run_aa_workbook/run_aa_workbook.py --program validate path/to/file.xlsm
```

## Requirements

- Python 3
- `clean_aa_backups`: [openpyxl](https://pypi.org/project/openpyxl/) —
  `pip install -r requirements-clean-aa.txt`
- `run_aa_workbook` (Windows only): Excel and
  [pywin32](https://pypi.org/project/pywin32/) —
  `pip install -r requirements-run-aa.txt`

## Documentation

See [docs/architecture.md](docs/architecture.md) for the full technical
documentation.
