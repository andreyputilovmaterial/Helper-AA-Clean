# Helper-AA-Clean

Small command-line helpers for working with AA (Excel) workbooks.

- `clean_aa_backups.py` — removes old timestamped backup sheets.
- `run_aa_workbook.py` — runs the workbook's own VBA steps automatically
  (apply definitions, load definitions, validation) on Windows + Excel.

## Quick start

```bash
# Preview a cleanup before touching anything (recommended first)
python src/clean_aa_backups.py --dry-run path/to/file.xlsm

# Run the AA workflow (Windows only)
python src/run_aa_workbook.py path/to/file.xlsm
```

## Requirements

- Python 3
- `clean_aa_backups.py`: [openpyxl](https://pypi.org/project/openpyxl/) —
  `pip install openpyxl`
- `run_aa_workbook.py` (Windows only): Excel and
  [pywin32](https://pypi.org/project/pywin32/) — `pip install pywin32`

## Documentation

See [docs/architecture.md](docs/architecture.md) for the full technical
documentation.
