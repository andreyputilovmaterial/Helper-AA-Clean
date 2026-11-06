# Project Context — Helper-AA-Clean

> Last updated: 2026-11-06
> (2026-11-06) Initial version. Created `src/clean_aa_backups.py` and its pytest
> suite. Retention: keep everything from the last 6 months, then the single most
> recent backup per sheet name per calendar quarter. Grouping is per sheet name
> (not per exact timestamp) because the three sheets of one backup are stamped a
> few seconds apart. Uses openpyxl (no live Excel automation), `keep_vba=True`
> for `.xlsm`.

## 1. Purpose

Clean old backup sheets out of AA (Excel) workbooks. Backup sheets are named
`<Name>_MMDDYY_HHMMSS`; live "main" sheets have no suffix and are never touched.

## 2. Layout

- `src/clean_aa_backups.py` — the whole tool (single file, no build step).
- `tests/test_clean_aa_backups.py` — pytest suite.
- `.clinerules/` — repo rules for non-interactive commands, project context,
  code style/errors, and commit messages.

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
- **No build/dist step.** The script runs straight from `src/`; there are no
  `.bat`/`.sh` wrappers.

## 4. Findings / gotchas

- The three sheets of one snapshot are written a few seconds apart, so the exact
  `HHMMSS` suffix is unique per sheet, not per snapshot.
- openpyxl re-saves the whole workbook; unusual elements (some charts/images/form
  controls) may not round-trip. `--dry-run` / `--backup` are the safety net.

## 5. Roadmap / follow-ups

- None currently. Possible: a `--now` override for reproducible CLI testing;
  directory recursion is already available via `-r`.

## 6. Open questions

- Whether the quarter rule should also count snapshots inside the 6-month window
  (currently it does not: the 6-month keep is applied first, so a quarter can end
  up with more than one kept set if several saves fall in the recent window).
