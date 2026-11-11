# Helper-AA-Clean

Command-line helpers for working with AA (Excel) workbooks. There are two
programs:

| Program | Purpose | Platform |
| --- | --- | --- |
| `run_aa_workbook` | Automate running the AA workbook's own VBA macros | Windows + Excel (pywin32) |
| `clean_aa_backups` | Purge/clean old timestamped backup sheets from AA workbooks | Any OS (openpyxl) |

---

## 1. Automate running — `run_aa_workbook`

**Purpose.** Run the workbook's own VBA macros automatically (through Excel)
instead of clicking the buttons and waiting for each message box. It works on a
temporary working copy, so the original file is never touched.

Each run executes exactly one program, chosen with `--program`. The individual
programs just execute the macro behind the corresponding AA button:

| `--program` key | Purpose | VBA macro |
| --- | --- | --- |
| `commit_save` | commit/save — validate, then write specifications back to the data file | `SaveAll` |
| `reload` | (re)load — re-load specifications from the data file | `ListAll` |
| `validate` | run validation — run the project-team validation | `Button1_Click` |

And two combined programs, `auto_apply_aa` and `auto_reload_aa`, which just run
the three steps **commit/save → (re)load → commit/save** in order:

| `--program` key | Purpose |
| --- | --- |
| `auto_apply_aa` | full cycle commit/save → (re)load → commit/save; a validation failure is a real failure (non-zero exit) |
| `auto_reload_aa` | same cycle; the second commit/save's expected validation error is treated as the success signal |

```bash
.venv/Scripts/python.exe src/run_aa_workbook/run_aa_workbook.py --program auto_apply_aa path/to/analysis_author.xlsm
```

**Expected output.** Results are recorded in two log files in the same
directory as the AA file:

- `aa_run_log.csv` — full log of all runs, one row per program with a
  timestamp (`datetime`, `program`, `aa_file`, `data_file`, `result`,
  `validation_details`). Validation results are captured and stored here too —
  they may not be immediately visible because Excel shows only the first line
  of a cell until you click it, and validation details can be stored as
  multi-line text.
- `aa_event_log.csv` — a more detailed event log for debugging (`datetime`,
  `event`, `phase`).

## 2. Purge / clean backups — `clean_aa_backups`

**Purpose.** Remove old, timestamped backup sheets from an AA workbook while
leaving the live "main" sheets untouched. Retention keeps everything from the
last 6 months, then the most recent backup per sheet name per calendar quarter.

**Expected output.** A modified Excel file: the same workbook re-saved with the
old backup sheets deleted. The original is not changed unless you ask it to be —
use `--dry-run` to preview what would be deleted and `--backup` to keep a
`<file>.bak` copy first.

```bash
.venv/Scripts/python.exe src/clean_aa_backups/clean_aa_backups.py path/to/analysis_author.xlsm
```

---

## Setup

Python 3 is required. On Windows the venv's Python lives in `.venv\Scripts`;
on macOS/Linux it lives in `.venv/bin`.

### To automate running (Windows + Excel only)

```bash
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements-run-aa.txt
.venv/Scripts/python.exe src/run_aa_workbook/run_aa_workbook.py --program auto_apply_aa path/to/analysis_author.xlsm
```

`--program` accepts `commit_save`, `reload`, `validate`, `auto_apply_aa`, or
`auto_reload_aa`.

### To purge/clean backups

```bash
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements-clean-aa.txt
.venv/Scripts/python.exe src/clean_aa_backups/clean_aa_backups.py path/to/analysis_author.xlsm
```

## Documentation

See the [docs/](docs/) folder for more details — in particular
[docs/architecture.md](docs/architecture.md) for the full technical
documentation.
