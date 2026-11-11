# Project Context — Helper-AA-Clean

> Last updated: 2026-11-11
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

## 1. Purpose

Clean old backup sheets out of AA (Excel) workbooks. Backup sheets are named
`<Name>_MMDDYY_HHMMSS`; live "main" sheets have no suffix and are never touched.

## 2. Layout

- `README.md` — entry-level overview only (links to `docs/architecture.md`).
- `docs/architecture.md` — stable technical documentation (full docs).
- `docs/project-context.md` — living working memory (this file; not stable).
- `src/clean_aa_backups.py` — the whole tool (single file, no build step).
- `src/run_aa_workbook.py` — Windows-only COM automation of the AA run steps.
- `tests/test_clean_aa_backups.py` — pytest suite.
- `tests/test_run_aa_workbook.py` — pytest suite for step planning / wait logic
  (no COM needed, so it runs on any OS).
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
- **`win32com` import is lazy.** It is imported inside a function, so `--help`,
  `--dry-run`, and the pytest suite run on non-Windows machines without pywin32.
- **Macro names are placeholders.** The real VBA procedure names are not in this
  repo, so they live in one `STEPS` table (plus `--macro-...` CLI overrides)
  instead of being hard-coded everywhere.

## 4. Findings / gotchas

- The three sheets of one snapshot are written a few seconds apart, so the exact
  `HHMMSS` suffix is unique per sheet, not per snapshot.
- openpyxl re-saves the whole workbook; unusual elements (some charts/images/form
  controls) may not round-trip. `--dry-run` / `--backup` are the safety net.
- `Application.Run` is synchronous, but the script still polls `Application.Ready`
  / `Application.Busy` as a safety net for macros that hand work off
  asynchronously.
- The run macros must be `Public`, take no arguments, and must not show modal UI
  (`MsgBox`, `InputBox`, user forms), otherwise the script hangs waiting for a
  click that never comes.

## 5. Roadmap / follow-ups

- `clean_aa_backups.py`: a `--now` override for reproducible CLI testing;
  directory recursion is already available via `-r`.
- `run_aa_workbook.py`: fill in the real VBA macro names once known; consider
  `--update-links` if the workbook prompts about external links.

## 6. Open questions

- Whether the quarter rule should also count snapshots inside the 6-month window
  (currently it does not: the 6-month keep is applied first, so a quarter can end
  up with more than one kept set if several saves fall in the recent window).
- The exact VBA procedure names behind the apply/load/validate buttons.
