"""Remove old timestamped backup sheets from AA (Excel) workbooks.

Backup sheets follow the naming convention ``<Name>_MMDDYY_HHMMSS``.  Sheets
without that suffix (``Main``, ``AnalysisLabels``, ...) are left untouched.

A backup is saved as a group of sheets (e.g. ``AnalysisLabels``, ``ShortNames``
and ``AnalysisValues``) whose timestamps differ by only a few seconds, because
each sheet is stamped as it is written.  Retention is therefore applied per
sheet name, so that keeping the most recent sheet of each name in a quarter
preserves whole backup sets together.

Retention rules:

1. Keep every backup sheet created within the last ``KEEP_MONTHS`` (default 6).
2. Among the remaining (older) sheets, keep only the single most recent one per
   sheet name per calendar quarter (Jan-Mar, Apr-Jun, Jul-Sep, Oct-Dec) per year.
3. Delete everything else.

The 6-month rule is applied first, so a sheet kept by that rule is never removed
even if it is not the most recent one for its name in its quarter.
"""

import argparse
import calendar
import re
import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path

import openpyxl


STDOUT_COLOR_RED = "\033[31m"
STDOUT_COLOR_RESET = "\033[0m"

KEEP_MONTHS_DEFAULT = 6
SUPPORTED_SUFFIXES = (".xlsx", ".xlsm")

# ``<Name>_MMDDYY_HHMMSS`` — six digits for the date, six for the time.
BACKUP_SHEET_RE = re.compile(r"^(?P<base>.+)_(?P<date>\d{6})_(?P<time>\d{6})$")


def parse_sheet_timestamp(name):
    """Return ``(base, datetime)`` for a backup sheet name, else ``None``.

    ``base`` is the sheet name without the ``MMDDYY_HHMMSS`` suffix.  A name
    that looks like a backup but has an impossible date (e.g. month 13) is
    treated as a non-backup sheet so it is never touched.
    """
    match = BACKUP_SHEET_RE.match(name)
    if match is None:
        return None

    date_str = match.group("date")  # MMDDYY
    time_str = match.group("time")  # HHMMSS
    try:
        timestamp = datetime(
            2000 + int(date_str[4:6]),  # YY -> 20YY
            int(date_str[0:2]),         # MM
            int(date_str[2:4]),         # DD
            int(time_str[0:2]),         # HH
            int(time_str[2:4]),         # MM
            int(time_str[4:6]),         # SS
        )
    except ValueError:
        return None

    return match.group("base"), timestamp


def quarter_key(timestamp):
    """Return the ``(year, quarter)`` a timestamp belongs to (quarter 0-3)."""
    return timestamp.year, (timestamp.month - 1) // 3


def subtract_months(timestamp, months):
    """Return ``timestamp`` shifted back by ``months`` calendar months.

    Clamps the day to the end of the target month (e.g. Mar 31 minus one month
    becomes Feb 28/29), so a 6-month window stays a real 6-month window.
    """
    total_months = timestamp.year * 12 + (timestamp.month - 1) - months
    year, zero_based_month = divmod(total_months, 12)
    month = zero_based_month + 1
    day = min(timestamp.day, calendar.monthrange(year, month)[1])
    return timestamp.replace(year=year, month=month, day=day)


def plan_deletions(sheet_names, now=None, keep_months=KEEP_MONTHS_DEFAULT):
    """Decide which sheet names to delete.

    Returns a list of sheet names.  ``now`` is the reference time used for the
    6-month window; it defaults to the current wall-clock time and can be passed
    explicitly in tests.
    """
    if now is None:
        now = datetime.now()
    cutoff = subtract_months(now, keep_months)

    backups = []  # (sheet name, base name, timestamp)
    for name in sheet_names:
        parsed = parse_sheet_timestamp(name)
        if parsed is None:
            continue  # main sheet — never a candidate for deletion
        base, timestamp = parsed
        backups.append((name, base, timestamp))

    if not backups:
        return []

    keep = set()
    older = []
    for name, _base, timestamp in backups:
        if timestamp >= cutoff:
            keep.add(name)  # rule 1: within the recent window
        else:
            older.append((name, _base, timestamp))

    # Rule 2: among the older sheets, keep the most recent per (name, quarter).
    best = {}
    for name, base, timestamp in older:
        key = (base, quarter_key(timestamp))
        if key not in best or timestamp > best[key][1]:
            best[key] = (name, timestamp)
    keep.update(name for name, _timestamp in best.values())

    return [name for name, _base, _timestamp in backups if name not in keep]


def collect_workbook_paths(paths, recursive=False):
    """Expand CLI paths into a deduplicated, ordered list of workbook files."""
    seen = set()
    result = []

    def add(path):
        resolved = path.resolve()
        if resolved not in seen:
            seen.add(resolved)
            result.append(resolved)

    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            candidates = sorted(path.rglob("*")) if recursive else sorted(path.iterdir())
            for candidate in candidates:
                if candidate.is_file() and candidate.suffix.lower() in SUPPORTED_SUFFIXES:
                    add(candidate)
        elif path.is_file():
            add(path)
        else:
            raise FileNotFoundError(f"{path} not found")

    return result


def clean_workbook(path, now=None, keep_months=KEEP_MONTHS_DEFAULT, dry_run=False, backup=False):
    """Clean one workbook in place; return the number of deleted sheets."""
    path = Path(path)

    # keep_vba=True preserves macros when re-saving an .xlsm workbook.
    if path.suffix.lower() == ".xlsm":
        workbook = openpyxl.load_workbook(path, keep_vba=True)
    else:
        workbook = openpyxl.load_workbook(path)

    try:
        sheet_names = workbook.sheetnames
        to_delete = plan_deletions(sheet_names, now=now, keep_months=keep_months)
        kept = [name for name in sheet_names if name not in set(to_delete)]

        if not to_delete:
            print(f"{path}: nothing to delete (kept {len(kept)} sheet(s))")
            return 0

        if not kept:
            raise RuntimeError("refusing to delete every sheet; the workbook must keep at least one")

        if dry_run:
            print(f"{path}: would delete {len(to_delete)} sheet(s) and keep {len(kept)}")
            for name in to_delete:
                print(f"  - {name}")
            return len(to_delete)

        if backup:
            backup_path = path.with_name(path.name + ".bak")
            shutil.copy2(path, backup_path)
            print(f"{path}: backed up to {backup_path}")

        for name in to_delete:
            workbook.remove(workbook[name])

        workbook.save(path)
        print(f"{path}: deleted {len(to_delete)} sheet(s); kept {len(kept)}")
        for name in to_delete:
            print(f"  - {name}")
        return len(to_delete)
    finally:
        workbook.close()


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="clean_aa_backups.py",
        description="Remove old MMDDYY_HHMMSS backup sheets from AA (Excel) workbooks.",
    )
    parser.add_argument(
        "paths",
        nargs="+",
        help="Excel file(s) or folder(s) to clean (.xlsx / .xlsm)",
    )
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="only print what would be deleted, do not modify files",
    )
    parser.add_argument(
        "--keep-months",
        type=int,
        default=KEEP_MONTHS_DEFAULT,
        help=f"keep snapshots newer than this many months (default {KEEP_MONTHS_DEFAULT})",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="when a folder is given, also scan its subfolders",
    )
    parser.add_argument(
        "--backup",
        action="store_true",
        help="copy each file to <file>.bak before modifying it",
    )
    args = parser.parse_args(argv)

    now = datetime.now()
    print(f"{parser.prog}: started at {now}")

    files = collect_workbook_paths(args.paths, recursive=args.recursive)
    if not files:
        print("no .xlsx/.xlsm files found", file=sys.stderr)
        return 1

    try:
        total_deleted = 0
        for path in files:
            total_deleted += clean_workbook(
                path,
                now=now,
                keep_months=args.keep_months,
                dry_run=args.dry_run,
                backup=args.backup,
            )

        verb = "would delete" if args.dry_run else "deleted"
        print(
            f"{parser.prog}: finished; {len(files)} file(s) checked, "
            f"{total_deleted} sheet(s) {verb}"
        )
        return 0
    except Exception as exc:
        print("", file=sys.stderr)
        print("Stack trace:", file=sys.stderr)
        print("", file=sys.stderr)
        traceback.print_exc()
        print("", file=sys.stderr)
        print("Error:", file=sys.stderr)
        print(f"{STDOUT_COLOR_RED}{exc}{STDOUT_COLOR_RESET}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

