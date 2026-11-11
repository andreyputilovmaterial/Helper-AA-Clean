# Project Context — Helper-AA-Clean

> Last updated: 2026-11-11
> (2026-11-11) Reorganized `src/` into one package directory per tool:
> `src/clean_aa_backups/` (openpyxl) and `src/run_aa_workbook/` (pywin32), each
> with its own module and `__init__.py`. Tests now import via
> `from clean_aa_backups import clean_aa_backups` /
> `from run_aa_workbook import run_aa_workbook`.
> (2026-11-11) Split `requirements.txt` into `requirements-clean-aa.txt`
> (openpyxl) and `requirements-run-aa.txt` (pywin32) so each tool's dependency
> can be installed independently.
> (2026-11-11) Slimmed `run_aa_workbook.py`'s CLI: removed `--read-only`,
> `--no-save`, `--dry-run`, and the `--macro-<step>` overrides. Macro names are
> hardcoded in `STEPS`. The workbook is now saved explicitly after every
> successful step (commit/save, reload, validate, and each step inside `auto`)
> instead of a single end-of-run save, so the working copy is always left in a
> clean, closable state.
> (2026-11-11) `run_aa_workbook.py` now runs exactly one program per
> invocation: a required `--program {reload,commit_save,validate,auto}` option
> replaced the old run-all-by-default + `--skip` model. `plan_actions()` was
> removed and `plan_steps(actions, macro_overrides=None)` no longer takes a
> `skip` argument or a default-all actions value.
> (2026-11-06) Initial version. Created `src/clean_aa_backups.py` and its pytest
> suite. Retention: keep everything from the last 6 months, then the single most
> recent backup per sheet name per calendar quarter. Grouping is per sheet name
> (not per exact timestamp) because the three sheets of one backup are stamped a
> few seconds apart. Uses openpyxl (no live Excel automation), `keep_vba=True`
> for `.xlsm`.
> (2026-11-11) Added `src/run_aa_workbook.py` (Windows-only, pywin32/COM) to
> automate the AA workbook's VBA buttons (apply definitions, load definitions,
> run validation) instead of clicking and waiting. Macro names are placeholders
> until the real VBA procedure names are known.
> (2026-11-11) Documentation split: README is now entry-level only; full
> technical docs moved to `docs/architecture.md` (stable). Replaced the
> `08-commit-messages` clinerule with `20-git-commit-messages-wording.md`,
> `22-when-making-commits.md` and `29-github-issues.md` (skipped
> `21-git-submodules-commits-order.md` — no submodules here).
> (2026-11-11) `run_aa_workbook.py` now makes a working copy of the input
> (`<name>-working-automated-YYYY-MM-DD.xlsm`, same dir, overwritten daily)
> before running and operates only on that copy, so the VBA macros' repeated
> saves never touch the original. `--dry-run` does not create the copy.
> (2026-11-11) `run_aa_workbook.py` reworked to the real three actions —
> reload (`ListAll`), commit/save (`SaveAll`), validate (`Button1_Click`) —
> with a `DialogWatcher` thread that clicks the message boxes, and a pre-step
> that reads the data-file path from `Main`!C13 into `DATA_FILE_PATH`.
> Success is confirmed by the dialog seen (plus an mtime check on the data
> file for commit/save); validation failures capture the dialog text and the
> `AnalysisValidation` sheet.
> (2026-11-11) Added a synthetic `auto` action (commit/save -> reload ->
> commit/save) and renamed the data-file path symbols to clear, non-cryptic
> naming.
> (2026-11-11) Clarified the `auto` pass/fail rules: the second commit/save's
> expected validation failure is the success signal; all other errors fail.
> (2026-11-11) Added CSV logging: each action appends a row to
> `aa_run_log.csv` next to the AA file (datetime, program, aa_file, data_file,
> result, validation_details). `auto` logs one row with the first validation
> failure of its cycle.

## 1. Purpose

Clean old backup sheets out of AA (Excel) workbooks. Backup sheets are named
`<Name>_MMDDYY_HHMMSS`; live "main" sheets have no suffix and are never touched.

## 2. Layout

- `README.md` — entry-level overview only (links to `docs/architecture.md`).
- `docs/architecture.md` — stable technical documentation (full docs).
- `docs/project-context.md` — living working memory (this file; not stable).
- `src/clean_aa_backups/` — the clean_aa_backups tool (openpyxl, no build step).
- `src/run_aa_workbook/` — Windows-only COM automation of the AA actions
  (reload / commit-save / validate) with programmatic message-box handling.
- `tests/test_clean_aa_backups.py` — pytest suite.
- `tests/test_run_aa_workbook.py` — pytest suite for step planning, dialog
  matching, data-file-path reading, and pass/fail outcome logic (no COM
  needed, so it runs on any OS).
- `.clinerules/` — repo rules for non-interactive commands, project context,
  code style/errors, commit-message wording/process, and GitHub issue
  references.

## 3. Key decisions & reasoning

- **openpyxl, not a live Excel instance.** Re-saving via Excel COM is
  interface-blocking and depends on Excel not hanging; openpyxl reads/writes the
  file directly.
- **Retention applied per sheet name, not per exact timestamp.** Each backup is
  a set of sheets (`AnalysisLabels`, `ShortNames`, `AnalysisValues`) whose
  timestamps differ by seconds (e.g. `..._101031`, `..._101048`, `..._101058`).
  Grouping by the exact suffix would split a set apart, so the most recent sheet
  per name per quarter is what gets kept.
- **6-month window uses real calendar months.** `subtract_months` clamps the day
  to the target month's length (Mar 31 − 6 months = Sep 30) instead of
  approximating with `timedelta(days=...)`.
- **No build/dist step.** The scripts run straight from `src/`; there are no
  `.bat`/`.sh` wrappers.
- **Run steps go through the workbook's own VBA.** `run_aa_workbook.py` opens the
  file with `Excel.Application` (`win32com.client.DispatchEx`) and calls
  `Application.Run("'file.xlsm'!Macro")`, so it exercises exactly what the
  buttons do rather than reimplementing the logic.
- **`win32com` import is lazy.** It is imported inside a function, so `--help`
  and the pytest suite run on non-Windows machines without pywin32.
- **The three actions map to the real VBA procedures.** `reload` → `ListAll`
  (Simplified), `commit_save` → `SaveAll` (Simplified), `validate` →
  `Button1_Click` (Addon), fixed in one `STEPS` table (no CLI override).
- **`auto` is a synthetic action.** It expands to commit/save -> reload ->
  commit/save (a full cycle of 2x commit/save with a reload in between) in
  `plan_steps()`, so the existing per-step dialog/outcome handling applies.
  Its second commit/save is expected to fail validation, and that `Validation
  Error` is treated as the success signal (any other error fails).
- **Explicit save after each successful step.** After a step's macro returns
  and its outcome is evaluated as passing, the workbook is saved
  (`workbook.Save()`) so its on-disk state is clean and it can be closed
  without Excel prompting. There is no `--read-only`, `--no-save`, or
  `--dry-run` mode: the macros save internally anyway, and dry-running is not
  meaningful for COM/VBA side effects.
- **One program per run.** The CLI requires `--program
  {reload,commit_save,validate,auto}`; there is no default "run everything"
  mode and no `--skip` flag. `main()` passes `[args.program]` straight into
  `plan_steps()`.
- **Message boxes are dismissed from a background thread.** `Application.Run`
  blocks on modal `MsgBox`s, so a `DialogWatcher` thread matches them by
  title/text and sends `WM_COMMAND` (IDOK/IDCANCEL); validation-error text and
  the `AnalysisValidation` sheet are captured for the failure report.
- **Pre-step #0 reads the data file path.** Cell `C13` on `Main` is read
  into the module global `DATA_FILE_PATH` (a `Path`), its existence is
  verified, and after commit/save the file's mtime must advance (the
  "recently edited" check).
- **CSV log next to the AA file.** Each action appends one row to
  `aa_run_log.csv` (created on first use) with datetime, program, aa_file,
  data_file, result, and validation_details. `auto` is logged as one row with
  the first validation failure of its cycle.

## 4. Findings / gotchas

- The three sheets of one snapshot are written a few seconds apart, so the exact
  `HHMMSS` suffix is unique per sheet, not per snapshot.
- openpyxl re-saves the whole workbook; unusual elements (some charts/images/form
  controls) may not round-trip. `--dry-run` / `--backup` are the safety net.
- `Application.Run` is synchronous, but the script still polls `Application.Ready`
  / `Application.Busy` as a safety net for macros that hand work off
  asynchronously.
- The run macros must be `Public`.  They do show modal `MsgBox` prompts; the
  known ones are dismissed by the watcher, but any other modal UI (`InputBox`,
  user forms, unknown `MsgBox`) still hangs the script.

## 5. Roadmap / follow-ups

- `clean_aa_backups.py`: a `--now` override for reproducible CLI testing;
  directory recursion is already available via `-r`.
- `run_aa_workbook.py`: consider `--update-links` if the workbook prompts about
  external links; consider a watchdog timeout for unexpected modal dialogs.

## 6. Open questions

- Whether the quarter rule should also count snapshots inside the 6-month window
  (currently it does not: the 6-month keep is applied first, so a quarter can end
  up with more than one kept set if several saves fall in the recent window).
- (resolved) The VBA procedure names: `ListAll`, `SaveAll`, `Button1_Click`.
