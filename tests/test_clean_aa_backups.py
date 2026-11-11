import sys
from datetime import datetime
from pathlib import Path

import openpyxl

SRC_DIR = Path(__file__).resolve().parents[1] / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from clean_aa_backups import clean_aa_backups as cab


# --- parse_sheet_timestamp ---------------------------------------------------

def test_parse_backup_sheet_name():
    base, dt = cab.parse_sheet_timestamp("AnalysisLabels_041124_131721")
    assert base == "AnalysisLabels"
    assert dt == datetime(2024, 4, 11, 13, 17, 21)


def test_parse_backup_sheet_with_underscored_base():
    base, dt = cab.parse_sheet_timestamp("SL_CategoryMAP_011224_120000")
    assert base == "SL_CategoryMAP"
    assert dt == datetime(2024, 1, 12, 12, 0, 0)


def test_parse_main_sheet_returns_none():
    for name in ("Main", "AnalysisLabels", "ShortNames", "AnalysisValues", "Exclusions"):
        assert cab.parse_sheet_timestamp(name) is None


def test_parse_invalid_date_returns_none():
    assert cab.parse_sheet_timestamp("AnalysisLabels_132524_120000") is None


# --- quarter_key -------------------------------------------------------------

def test_quarter_key():
    assert cab.quarter_key(datetime(2024, 1, 1)) == (2024, 0)
    assert cab.quarter_key(datetime(2024, 3, 31)) == (2024, 0)
    assert cab.quarter_key(datetime(2024, 4, 1)) == (2024, 1)
    assert cab.quarter_key(datetime(2024, 7, 1)) == (2024, 2)
    assert cab.quarter_key(datetime(2024, 12, 31)) == (2024, 3)


# --- subtract_months ---------------------------------------------------------

def test_subtract_months_basic():
    assert cab.subtract_months(datetime(2024, 12, 1, 12, 0, 0), 6) == datetime(2024, 6, 1, 12, 0, 0)


def test_subtract_months_clamps_day():
    # Mar 31 minus one month -> Feb 29 in a leap year.
    assert cab.subtract_months(datetime(2024, 3, 31), 1) == datetime(2024, 2, 29)
    assert cab.subtract_months(datetime(2024, 3, 31), 6) == datetime(2023, 9, 30)


# --- plan_deletions ----------------------------------------------------------

BASES = ("AnalysisLabels", "ShortNames", "AnalysisValues")


def _suffix(month, day, year, hour, minute, second):
    return f"{month:02d}{day:02d}{year % 100:02d}_{hour:02d}{minute:02d}{second:02d}"


def _triplet(month, day, hour, minute, second, year=2024):
    # Each sheet of one backup is stamped a few seconds apart as it is written.
    offsets = (0, 5, 8)
    return [
        f"{base}_{_suffix(month, day, year, hour, minute, second + off)}"
        for base, off in zip(BASES, offsets)
    ]


def _sheets(*triplets):
    names = ["Main", "AnalysisLabels", "ShortNames", "AnalysisValues"]
    for triplet in triplets:
        names.extend(triplet)
    return names


def test_plan_deletions_keeps_recent_and_one_quarter_per_name():
    now = datetime(2024, 12, 1, 12, 0, 0)  # 6-month cutoff = 2024-06-01 12:00
    recent_a = _triplet(11, 1, 10, 0, 0)             # within 6m -> keep
    recent_b = _triplet(6, 1, 13, 0, 0)              # just inside window -> keep
    q2_newer = _triplet(5, 1, 10, 0, 0)              # Q2 most recent -> keep
    q2_older = _triplet(4, 1, 10, 0, 0)              # Q2 -> delete
    q1 = _triplet(2, 1, 10, 0, 0)                    # Q1 -> keep
    q4_prev_newer = _triplet(12, 1, 10, 0, 0, year=2023)  # Q4 2023 most recent -> keep
    q4_prev_older = _triplet(10, 1, 10, 0, 0, year=2023)  # Q4 2023 -> delete
    q3_prev = _triplet(7, 1, 10, 0, 0, year=2023)    # Q3 2023 -> keep

    names = _sheets(
        recent_a, recent_b, q2_newer, q2_older, q1,
        q4_prev_newer, q4_prev_older, q3_prev,
    )
    deleted = cab.plan_deletions(names, now=now, keep_months=6)

    assert set(deleted) == set(q2_older + q4_prev_older)


def test_plan_deletions_keeps_full_triplets_together():
    now = datetime(2024, 12, 1, 12, 0, 0)
    latest = _triplet(5, 1, 10, 0, 0)   # older than 6m, Q2 most recent
    older = _triplet(4, 1, 10, 0, 0)    # older than 6m, Q2 -> delete
    deleted = cab.plan_deletions(_sheets(latest, older), now=now, keep_months=6)

    assert set(deleted) == set(older)
    assert not set(latest) & set(deleted)


def test_plan_deletions_all_recent_keeps_everything():
    now = datetime(2024, 12, 1, 12, 0, 0)
    names = _sheets(
        _triplet(11, 1, 10, 0, 0),
        _triplet(10, 1, 10, 0, 0),
        _triplet(7, 1, 10, 0, 0),
    )
    assert cab.plan_deletions(names, now=now, keep_months=6) == []


def test_plan_deletions_no_backups():
    names = ["Main", "AnalysisLabels", "ShortNames", "AnalysisValues"]
    assert cab.plan_deletions(names, now=datetime(2024, 12, 1, 12, 0, 0)) == []


# --- clean_workbook (integration) --------------------------------------------

def test_clean_workbook_removes_backups_and_keeps_mains(tmp_path):
    path = tmp_path / "backup_test.xlsx"
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for name in ("Main", "AnalysisLabels", "ShortNames", "AnalysisValues"):
        workbook.create_sheet(name)
    for triplet in (_triplet(5, 1, 10, 0, 0), _triplet(4, 1, 10, 0, 0)):
        for name in triplet:
            workbook.create_sheet(name)
    workbook.save(path)

    now = datetime(2024, 12, 1, 12, 0, 0)
    deleted = cab.clean_workbook(path, now=now, keep_months=6)
    assert deleted == 3

    reloaded = openpyxl.load_workbook(path)
    expected = {"Main", "AnalysisLabels", "ShortNames", "AnalysisValues"}
    expected.update(_triplet(5, 1, 10, 0, 0))
    assert set(reloaded.sheetnames) == expected


def test_clean_workbook_dry_run_does_not_modify(tmp_path):
    path = tmp_path / "backup_test.xlsx"
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    workbook.create_sheet("Main")
    for triplet in (_triplet(5, 1, 10, 0, 0), _triplet(4, 1, 10, 0, 0)):
        for name in triplet:
            workbook.create_sheet(name)
    workbook.save(path)

    deleted = cab.clean_workbook(
        path, now=datetime(2024, 12, 1, 12, 0, 0), keep_months=6, dry_run=True
    )
    assert deleted == 3

    reloaded = openpyxl.load_workbook(path)
    for name in _triplet(4, 1, 10, 0, 0):
        assert name in reloaded.sheetnames

