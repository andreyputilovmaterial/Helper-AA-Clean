import sys
from datetime import date
from pathlib import Path

import pytest

SRC_DIR = Path(__file__).resolve().parents[1] / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import run_aa_workbook as ruw


class FakeExcel:
    """Minimal stand-in for an Excel.Application COM object."""

    def __init__(self, ready=True, busy=False):
        self.ready = ready
        self.busy = busy
        self.calls = []

    @property
    def Ready(self):
        return self.ready

    @property
    def Busy(self):
        return self.busy

    def Run(self, qualified):
        self.calls.append(qualified)


# --- plan_steps -------------------------------------------------------------

def test_plan_steps_runs_all_by_default():
    steps = ruw.plan_steps(ruw.STEPS)
    assert [s.key for s in steps] == [
        "apply_definitions",
        "load_definitions",
        "run_validation",
    ]


def test_plan_steps_skips_a_step():
    steps = ruw.plan_steps(ruw.STEPS, skip=["load_definitions"])
    assert [s.key for s in steps] == ["apply_definitions", "run_validation"]


def test_plan_steps_overrides_macro_name():
    steps = ruw.plan_steps(ruw.STEPS, macro_overrides={"run_validation": "ValidateNow"})
    by_key = {s.key: s for s in steps}
    assert by_key["run_validation"].macro == "ValidateNow"
    assert by_key["apply_definitions"].macro == "ApplyDefinitions"


def test_plan_steps_all_skipped_is_empty():
    keys = [s.key for s in ruw.STEPS]
    assert ruw.plan_steps(ruw.STEPS, skip=keys) == []


# --- wait_for_excel_idle ----------------------------------------------------

def test_wait_for_excel_idle_returns_when_ready():
    xl = FakeExcel(ready=True, busy=False)
    ruw.wait_for_excel_idle(xl, timeout=1, poll_interval=0.01)


def test_wait_for_excel_idle_times_out():
    xl = FakeExcel(ready=False, busy=True)
    with pytest.raises(RuntimeError):
        ruw.wait_for_excel_idle(xl, timeout=0.05, poll_interval=0.01)


# --- run_macro --------------------------------------------------------------

def test_run_macro_runs_qualified_name():
    xl = FakeExcel(ready=True, busy=False)
    ruw.run_macro(xl, "aaa.xlsm", "ApplyDefinitions", timeout=1, settle_seconds=0)
    assert xl.calls == ["'aaa.xlsm'!ApplyDefinitions"]


# --- working copy ------------------------------------------------------------

def test_working_copy_path_appends_date_to_stem(tmp_path):
    src = tmp_path / "MyAA.xlsm"
    dest = ruw.working_copy_path(src, today=date(2026, 11, 11))
    assert dest == tmp_path / "MyAA-working-automated-2026-11-11.xlsm"


def test_working_copy_path_keeps_extension_and_zero_padded_date(tmp_path):
    src = tmp_path / "AA.xlsx"
    dest = ruw.working_copy_path(src, today=date(2026, 1, 5))
    assert dest == tmp_path / "AA-working-automated-2026-01-05.xlsx"


def test_working_copy_path_is_in_same_directory_as_source(tmp_path):
    sub = tmp_path / "sub"
    sub.mkdir()
    src = sub / "AA.xlsm"
    dest = ruw.working_copy_path(src, today=date(2026, 11, 11))
    assert dest.parent == sub
    assert dest == sub / "AA-working-automated-2026-11-11.xlsm"


def test_make_working_copy_creates_copy_and_leaves_original(tmp_path):
    src = tmp_path / "AA.xlsm"
    src.write_bytes(b"v1")
    dest = ruw.make_working_copy(src, today=date(2026, 11, 11))
    assert dest.exists()
    assert dest.read_bytes() == b"v1"
    assert src.read_bytes() == b"v1"


def test_make_working_copy_overwrites_existing_copy(tmp_path):
    src = tmp_path / "AA.xlsm"
    src.write_bytes(b"v1")
    dest = ruw.make_working_copy(src, today=date(2026, 11, 11))

    src.write_bytes(b"v2")
    dest2 = ruw.make_working_copy(src, today=date(2026, 11, 11))
    assert dest2 == dest
    assert dest.read_bytes() == b"v2"


# --- argument parsing -------------------------------------------------------

def test_parser_skip_and_macro_override():
    args = ruw.build_parser().parse_args(
        [
            "some.xlsm",
            "--skip",
            "load_definitions",
            "--macro-run-validation",
            "ValidateNow",
        ]
    )
    assert args.skip == ["load_definitions"]
    assert args.macro_run_validation == "ValidateNow"
    assert args.macro_apply_definitions is None
