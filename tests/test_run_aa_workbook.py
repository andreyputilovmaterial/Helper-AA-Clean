import csv
import os
import sys
from datetime import date, datetime
from pathlib import Path

import pytest

SRC_DIR = Path(__file__).resolve().parents[1] / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from run_aa_workbook import run_aa_workbook as ruw


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

def test_plan_steps_requires_explicit_actions():
    with pytest.raises(TypeError):
        ruw.plan_steps()


def test_plan_steps_auto_expands_to_commit_reload_commit():
    steps = ruw.plan_steps(actions=[ruw.AUTO])
    assert [s.key for s in steps] == ["commit_save", "reload", "commit_save"]


def test_plan_steps_uses_fixed_macros():
    steps = ruw.plan_steps([ruw.RELOAD, ruw.COMMIT_SAVE, ruw.VALIDATE])
    assert [s.macro for s in steps] == ["ListAll", "SaveAll", "Button1_Click"]


def test_plan_steps_auto_tags_second_commit_as_expected():
    steps = ruw.plan_steps(actions=[ruw.AUTO])
    assert [s.context for s in steps] == ["", "", ruw.AUTO_COMMIT_SECOND]


def test_plan_steps_auto_labels_are_prefixed():
    steps = ruw.plan_steps(actions=[ruw.AUTO])
    assert [s.label for s in steps] == [
        "auto: commit/save specifications",
        "auto: reload specifications",
        "auto: commit/save specifications",
    ]


def test_program_names_mapping():
    assert ruw.PROGRAM_NAMES[ruw.COMMIT_SAVE] == "commit/save"
    assert ruw.PROGRAM_NAMES[ruw.RELOAD] == "(re)load"
    assert ruw.PROGRAM_NAMES[ruw.VALIDATE] == "run validation"
    assert ruw.PROGRAM_NAMES[ruw.AUTO] == "auto"


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
    ruw.run_macro(xl, "aaa.xlsm", "ListAll", timeout=1, settle_seconds=0)
    assert xl.calls == ["'aaa.xlsm'!ListAll"]


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


# --- find_dialog ------------------------------------------------------------

def dialog(title="", text="", action="ok"):
    return ruw.Dialog(key="d", title=title, text=text, action=action)


def test_find_dialog_matches_title_and_text():
    found = ruw.find_dialog(
        "Validation Error",
        "An error occured while processing your specs.",
        [dialog(title="Validation Error", text="error occured")],
    )
    assert found is not None


def test_find_dialog_is_case_insensitive():
    found = ruw.find_dialog(
        "validation error",
        "AN ERROR OCCURED",
        [dialog(title="Validation Error", text="error occured")],
    )
    assert found is not None


def test_find_dialog_empty_title_matches_any_caption():
    found = ruw.find_dialog(
        "Some Caption",
        "Validation Complete! No Errors found!",
        [dialog(title="", text="Validation Complete!")],
    )
    assert found is not None


def test_find_dialog_returns_none_when_no_match():
    assert ruw.find_dialog("Other", "Other text", [dialog(title="Overwrite?")]) is None


# --- find_button_child ---------------------------------------------------------

def child(hwnd, text="", class_name="", ctrl_id=0):
    return {"hwnd": hwnd, "text": text, "class": class_name, "id": ctrl_id}


def test_find_button_child_prefers_standard_control_id():
    children = [
        child(1, "OK", "ThunderCommandButton", ruw.WIN32_IDOK),
        child(2, "Cancel", "ThunderCommandButton", ruw.WIN32_IDCANCEL),
    ]
    assert ruw.find_button_child(children, "ok")["hwnd"] == 1
    assert ruw.find_button_child(children, "cancel")["hwnd"] == 2


def test_find_button_child_recognises_yes_no_control_ids():
    children = [
        child(7, "Yes", "Button", ruw.WIN32_IDYES),
        child(8, "No", "Button", ruw.WIN32_IDNO),
    ]
    assert ruw.find_button_child(children, "ok")["hwnd"] == 7
    assert ruw.find_button_child(children, "cancel")["hwnd"] == 8


def test_find_button_child_falls_back_to_caption():
    children = [
        child(3, "OK", "ThunderCommandButton"),
        child(4, "Cancel", "ThunderCommandButton"),
    ]
    assert ruw.find_button_child(children, "ok")["hwnd"] == 3
    assert ruw.find_button_child(children, "cancel")["hwnd"] == 4


def test_find_button_child_falls_back_to_window_class():
    children = [child(5, "Click here", "ThunderCommandButton")]
    assert ruw.find_button_child(children, "ok")["hwnd"] == 5


def test_find_button_child_returns_none_when_no_button():
    assert ruw.find_button_child([], "ok") is None
    assert ruw.find_button_child([child(6, "note", "Static")], "ok") is None


# --- read_data_file_path ------------------------------------------------------

class FakeCell:
    def __init__(self, value):
        self.Value = value


class FakeMainSheet:
    def __init__(self, value):
        self.value = value

    def Range(self, _cell):
        return FakeCell(self.value)


class FakeMainWorkbook:
    def __init__(self, value):
        self.value = value

    def Sheets(self, _name):
        return FakeMainSheet(self.value)


def test_read_data_file_path_returns_existing_file(tmp_path):
    target = tmp_path / "data.xlsm"
    target.write_bytes(b"x")
    result = ruw.read_data_file_path(FakeMainWorkbook(str(target)))
    assert isinstance(result, Path)
    assert result == target


def test_read_data_file_path_strips_quotes_and_whitespace(tmp_path):
    target = tmp_path / "data.xlsm"
    target.write_bytes(b"x")
    result = ruw.read_data_file_path(FakeMainWorkbook(f'  "{target}"  '))
    assert result == target


def test_read_data_file_path_missing_file_raises():
    with pytest.raises(FileNotFoundError):
        ruw.read_data_file_path(FakeMainWorkbook("Z:/no/such/file.xlsm"))


def test_read_data_file_path_empty_cell_raises():
    with pytest.raises(RuntimeError):
        ruw.read_data_file_path(FakeMainWorkbook("   "))


# --- read_sheet_text ---------------------------------------------------------

class FakeUsedRange:
    def __init__(self, value):
        self.Value = value


class FakeSheet:
    def __init__(self, value):
        self.UsedRange = FakeUsedRange(value)


class FakeSheetWorkbook:
    def __init__(self, value):
        self.value = value

    def Sheets(self, _name):
        return FakeSheet(self.value)


def test_read_sheet_text_empty():
    assert ruw.read_sheet_text(FakeSheetWorkbook(None), "AnalysisValidation") == ""


def test_read_sheet_text_single_value():
    assert ruw.read_sheet_text(FakeSheetWorkbook("oops"), "x") == "oops"


def test_read_sheet_text_rows_and_none_cells():
    value = (("a", None, "c"), ("d", "e"))
    assert ruw.read_sheet_text(FakeSheetWorkbook(value), "x") == "a\t\tc\nd\te"


# --- verify_data_file_recently_edited ----------------------------------------

def test_verify_data_file_recently_edited_passes_after_edit(tmp_path):
    target = tmp_path / "data.xlsm"
    target.write_bytes(b"x")
    before = ruw.data_file_mtime(target)
    os.utime(target, (before, before + 10))
    ruw.verify_data_file_recently_edited(target, before)


def test_verify_data_file_recently_edited_raises_when_unchanged(tmp_path):
    target = tmp_path / "data.xlsm"
    target.write_bytes(b"x")
    before = ruw.data_file_mtime(target)
    with pytest.raises(RuntimeError):
        ruw.verify_data_file_recently_edited(target, before)


# --- logging helpers ---------------------------------------------------------

def test_format_validation_details_without_sheet():
    assert ruw.format_validation_details("msg") == "msg"


def test_format_validation_details_with_sheet():
    assert ruw.format_validation_details("msg", "s1\ns2") == (
        "msg\n--- AnalysisValidation sheet ---\ns1\ns2"
    )


def test_write_log_row_creates_header_and_row(tmp_path):
    log = tmp_path / "aa_run_log.csv"
    now = datetime(2026, 11, 11, 12, 34, 56)
    ruw.write_log_row(
        log, "auto", "/x/AA.xlsm", "/x/data.xlsm", "success", "details", now=now
    )
    assert log.read_text(encoding="utf-8").splitlines() == [
        "datetime,program,aa_file,data_file,result,validation_details",
        "2026-11-11 12:34:56,auto,/x/AA.xlsm,/x/data.xlsm,success,details",
    ]


def test_write_log_row_appends_without_second_header(tmp_path):
    log = tmp_path / "aa_run_log.csv"
    now = datetime(2026, 11, 11, 12, 0, 0)
    ruw.write_log_row(log, "commit/save", "a", "b", "success", "", now=now)
    ruw.write_log_row(log, "auto", "a", "b", "boom", "", now=now)
    lines = log.read_text(encoding="utf-8").splitlines()
    assert lines[0].startswith("datetime,program")
    assert len(lines) == 3
    assert "boom" in lines[2]


def test_write_log_row_escapes_csv_specials(tmp_path):
    log = tmp_path / "aa_run_log.csv"
    now = datetime(2026, 11, 11, 12, 0, 0)
    details = 'a,b\n"c"'
    ruw.write_log_row(log, "auto", "a", "b", "success", details, now=now)
    with open(log, newline="", encoding="utf-8") as handle:
        rows = list(csv.reader(handle))
    assert rows[0] == ruw.LOG_HEADER
    assert rows[1][5] == details


# --- evaluate_step_outcome ---------------------------------------------------

def step(key):
    return ruw.Step(key, key, key)


def record(key, text=""):
    dlg = ruw.Dialog(key=key, title="", text="", action="ok")
    return ruw.DialogRecord(dialog=dlg, captured_text=text, matched_caption="")


def test_evaluate_reload_does_nothing():
    ruw.evaluate_step_outcome(step(ruw.RELOAD), [], None, FakeSheetWorkbook(None))


def test_evaluate_commit_save_success_with_edited_data_file(tmp_path):
    target = tmp_path / "data.xlsm"
    target.write_bytes(b"x")
    before = ruw.data_file_mtime(target)
    os.utime(target, (before, before + 10))
    ruw.DATA_FILE_PATH = target
    try:
        ruw.evaluate_step_outcome(
            step(ruw.COMMIT_SAVE),
            [record("save_success", "Your Analysis Specifications have been saved")],
            before,
            FakeSheetWorkbook(None),
        )
    finally:
        ruw.DATA_FILE_PATH = None


def test_evaluate_commit_save_success_but_data_file_not_edited_raises(tmp_path):
    target = tmp_path / "data.xlsm"
    target.write_bytes(b"x")
    before = ruw.data_file_mtime(target)
    ruw.DATA_FILE_PATH = target
    try:
        with pytest.raises(RuntimeError):
            ruw.evaluate_step_outcome(
                step(ruw.COMMIT_SAVE),
                [record("save_success")],
                before,
                FakeSheetWorkbook(None),
            )
    finally:
        ruw.DATA_FILE_PATH = None


def test_evaluate_commit_save_without_success_raises():
    with pytest.raises(RuntimeError):
        ruw.evaluate_step_outcome(
            step(ruw.COMMIT_SAVE), [], None, FakeSheetWorkbook(None)
        )


def test_evaluate_commit_save_validation_error_captures_dialog_and_sheet():
    err = ruw.DialogRecord(
        dialog=ruw.VALIDATION_ERROR_DIALOG,
        captured_text=(
            "An error occured while processing your specifications.\n\n"
            "Message: bad"
        ),
        matched_caption="Validation Error",
    )
    with pytest.raises(ruw.ValidationError) as excinfo:
        ruw.evaluate_step_outcome(
            step(ruw.COMMIT_SAVE),
            [err],
            None,
            FakeSheetWorkbook(("row1col1", "row1col2")),
        )
    message = str(excinfo.value)
    assert "Message: bad" in message
    assert "row1col1" in message


def test_evaluate_validate_success_passes():
    ruw.evaluate_step_outcome(
        step(ruw.VALIDATE),
        [record("validation_success", "Validation Complete! No Errors found!")],
        None,
        FakeSheetWorkbook(None),
    )


def test_evaluate_validate_without_success_raises():
    with pytest.raises(RuntimeError):
        ruw.evaluate_step_outcome(
            step(ruw.VALIDATE), [], None, FakeSheetWorkbook(None)
        )


def test_evaluate_validate_error_raises():
    err = ruw.DialogRecord(
        dialog=ruw.VALIDATION_ERROR_DIALOG,
        captured_text="An error occured while processing your specifications.",
        matched_caption="Validation Error",
    )
    with pytest.raises(ruw.ValidationError):
        ruw.evaluate_step_outcome(
            step(ruw.VALIDATE), [err], None, FakeSheetWorkbook(("a", "b"))
        )


def auto_second_commit_step():
    return ruw.Step(
        ruw.COMMIT_SAVE,
        "auto: commit/save specifications",
        "SaveAll",
        ruw.AUTO_COMMIT_SECOND,
    )


def test_evaluate_auto_second_commit_validation_failure_is_success():
    err = ruw.DialogRecord(
        dialog=ruw.VALIDATION_ERROR_DIALOG,
        captured_text="An error occured while processing your specifications.",
        matched_caption="Validation Error",
    )
    ruw.evaluate_step_outcome(
        auto_second_commit_step(), [err], None, FakeSheetWorkbook(None)
    )


def test_evaluate_auto_second_commit_validation_failure_returns_details():
    err = ruw.DialogRecord(
        dialog=ruw.VALIDATION_ERROR_DIALOG,
        captured_text="An error occured while processing your specifications.",
        matched_caption="Validation Error",
    )
    details = ruw.evaluate_step_outcome(
        auto_second_commit_step(), [err], None, FakeSheetWorkbook(("a", "b"))
    )
    assert "An error occured" in details
    assert "AnalysisValidation sheet" in details


def test_evaluate_auto_second_commit_save_success_is_success():
    ruw.evaluate_step_outcome(
        auto_second_commit_step(),
        [record("save_success", "Your Analysis Specifications have been saved")],
        None,
        FakeSheetWorkbook(None),
    )


def test_evaluate_auto_second_commit_without_expected_message_raises():
    with pytest.raises(RuntimeError):
        ruw.evaluate_step_outcome(
            auto_second_commit_step(), [], None, FakeSheetWorkbook(None)
        )


# --- argument parsing -------------------------------------------------------

def test_parser_program_and_visible():
    args = ruw.build_parser().parse_args(
        ["some.xlsm", "--program", "validate", "--visible"]
    )
    assert args.program == "validate"
    assert args.visible is True


def test_parser_accepts_auto_program():
    args = ruw.build_parser().parse_args(["some.xlsm", "--program", "auto"])
    assert args.program == "auto"


def test_parser_requires_program():
    with pytest.raises(SystemExit):
        ruw.build_parser().parse_args(["some.xlsm"])


def test_parser_rejects_unknown_program():
    with pytest.raises(SystemExit):
        ruw.build_parser().parse_args(["some.xlsm", "--program", "nope"])
