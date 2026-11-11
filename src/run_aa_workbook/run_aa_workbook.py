"""Run the AA (Excel) workbook actions automatically via Excel.Application.

The AA workbook ships three buttons that drive its VBA workflow:

* reload (``ListAll``) — re-load specifications from the data file
* commit/save (``SaveAll``) — validate, then write specifications back
* validate (``Button1_Click``) — run the project-team validation

There are also two synthetic ``auto`` actions that run the same full cycle
of commit/save -> reload -> commit/save (2x commit/save with a reload in
between), expanded from the three leaf steps above: ``auto_apply_aa`` and
``auto_reload_aa``.  Each run executes exactly one action, chosen with
``--program``.

Clicking them by hand and waiting for each message box is slow and
error-prone.  This script opens the workbook with ``Excel.Application``
(pywin32 / ``win32com.client``), reads the data-file path from cell C13
of the ``Main`` sheet, then calls the same VBA procedures those buttons call,
dismissing the message boxes programmatically as they appear.

Before running anything it makes a working copy of the input file next to
the original (``<name>-working-automated-YYYY-MM-DD.xlsm``) and runs only
against that copy, so the original is never touched even though the macros
save repeatedly.

Message boxes
-------------
The macros show modal ``MsgBox`` prompts.  A background thread watches for
them by title/text and clicks the right button (OK / Cancel), recording the
text of the ones that matter (validation errors) so the script can report it:

* reload: ``Overwrite?`` / "Are you sure you want to load specifications?"
* commit/save success: ``Save Successful`` / "Your Analysis Specifications
  have been saved"
* commit/save failure: ``Validation Error`` / "An error occured while
  processing your specifications."
* validate success: "Validation Complete! No Errors found!"
* validate failure: ``Validation Error`` / "An error occured while processing
  your specifications."

Any other modal prompt will still block the run, because there is nobody to
click it.

Auto action pass/fail
---------------------
The synthetic ``auto_apply_aa`` and ``auto_reload_aa`` actions both run
commit/save -> reload -> commit/save:

* first commit/save — normal rules (validation failure or any other error
  is a failure);
* reload — any error is a failure;
* second commit/save — ``auto_reload_aa`` treats its expected validation
  failure as the success signal (any other error is a failure), while
  ``auto_apply_aa`` uses the normal commit/save rules, so a validation
  failure is a real failure and the run exits non-zero.

Data file (pre-step #0)
-----------------------
The path to the external data file is read from cell ``C13`` on the
``Main`` sheet before any action runs and exposed as the module global
``DATA_FILE_PATH``.  Its physical existence is verified (fail fast if it is
missing), and after a successful commit/save the script verifies the file was
modified during the save.

Logging
-------
Every action that actually runs appends one row to a CSV log
(``aa_run_log.csv``) in the same directory as the AA file.  Columns are
``datetime``, ``program`` (commit/save, (re)load, run validation,
auto_apply_aa, auto_reload_aa),
``aa_file``, ``data_file``, ``result`` (``success`` or the exception message),
and ``validation_details`` (the message box text plus the
``AnalysisValidation`` sheet when a validation failed).

A second, more frequently updated event log (``aa_event_log.csv``) in the
same directory records one row per lifecycle event with three columns:
``datetime``, ``event``, and ``phase`` (``started`` / ``finished``).  The
events are the program start/finish, each macro's start/finish, a
failed-validation signal, exceptions, explicit saves, and captured dialog
boxes.

Windows only
------------
This uses COM automation, so it runs on Windows with Excel installed and
``pywin32`` available (``pip install pywin32``).  On other platforms the
import is delayed so ``--help`` and the test suite still work.
"""

import argparse
import csv
import shutil
import sys
import threading
import time
import traceback
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path


STDOUT_COLOR_RED = "\033[31m"
STDOUT_COLOR_RESET = "\033[0m"

SUPPORTED_SUFFIXES = (".xlsx", ".xlsm")

# Seconds to wait, after calling a macro, for Excel to report itself idle
# again before continuing.  ``Application.Run`` is synchronous, so the macro
# has already returned when this wait starts; the poll is a best-effort safety
# net for any asynchronous tidying Excel does afterwards, and a timeout only
# logs a warning (see ``wait_for_excel_idle``).  A short default keeps a stale
# ``Busy``/``Ready`` flag (which can happen after a modal MsgBox is dismissed
# from another thread) from stalling every step for minutes.
IDLE_TIMEOUT_DEFAULT = 10
IDLE_POLL_INTERVAL = 0.25
SETTLE_SECONDS_DEFAULT = 1.0

# 0 = do not update links on open (and never prompt); use 3 to update all.
UPDATE_LINKS = 0

# Cell on the "Main" sheet that holds the path to the external data file,
# and the sheet that receives the validation-failure detail.
MAIN_SHEET_NAME = "Main"
DATA_FILE_CELL = "C13"
ANALYSIS_VALIDATION_SHEET = "AnalysisValidation"

# How often the background thread scans for modal message boxes.
DIALOG_POLL_INTERVAL = 0.1

# Step keys (used for --program choices and per-step dialog wiring).
RELOAD = "reload"
COMMIT_SAVE = "commit_save"
VALIDATE = "validate"
AUTO_APPLY_AA = "auto_apply_aa"
AUTO_RELOAD_AA = "auto_reload_aa"

#: Path to the external data file the macros read/write.  Read from
#: ``Main``!C13 before any action runs (pre-step #0) and shared globally so
#: every part of the run can use it.
DATA_FILE_PATH = None

# CSV log written next to the AA file: one row per action run.
LOG_FILENAME = "aa_run_log.csv"
LOG_HEADER = [
    "datetime",
    "program",
    "aa_file",
    "data_file",
    "result",
    "validation_details",
]

# Detailed event log written next to the AA file: one row per event, updated
# more frequently than the per-action CSV log above.  Columns are a timestamp,
# the event, and its phase ("started" / "finished").
EVENT_LOG_FILENAME = "aa_event_log.csv"
EVENT_LOG_HEADER = ["datetime", "event", "phase"]

#: Phase column values for the event log.
EVENT_PHASE_STARTED = "started"
EVENT_PHASE_FINISHED = "finished"

#: Path of the event log for the current run (set by ``main``); ``None`` means
#: no event logging (tests, or library use without a configured log).
EVENT_LOG_PATH = None

# Serialises appends to the event log, which is written from both the main
# thread (macro start/finish, save, validation) and the dialog-watcher thread.
_EVENT_LOG_LOCK = threading.Lock()

# Event names written to the event log (column 2).  Start/finish of macros and
# dialogs append the macro name / dialog caption at the call site.
EVENT_PROGRAM_STARTED = "started the program (overall py script)"
EVENT_PROGRAM_FINISHED = "finished the program (overall py script)"
EVENT_MACRO_STARTED = "started the macro"
EVENT_MACRO_FINISHED = "finished the macro"
EVENT_VALIDATION_FAILURE = "received a signal about failed validation"
EVENT_EXCEPTION = "exception"
EVENT_SAVE = "explicitly called save on Excel file"
EVENT_DIALOG_CAPTURED = "captured a dialog box"

# Column 2 values (the programs/actions the script can run).
PROGRAM_NAMES = {
    RELOAD: "(re)load",
    COMMIT_SAVE: "commit/save",
    VALIDATE: "run validation",
    AUTO_APPLY_AA: "auto_apply_aa",
    AUTO_RELOAD_AA: "auto_reload_aa",
}


@dataclass(frozen=True)
class Step:
    """One VBA step of the AA run.

    ``context`` marks a step as part of a composite action (used by
    ``evaluate_step_outcome`` for special pass/fail rules); it is ``""`` for
    ordinary steps.
    """

    key: str
    label: str
    macro: str
    context: str = ""


# The macro names are the VBA entry points behind the AA buttons:
# * reload      -> "Simplified" module: ListAll
# * commit/save -> "Simplified" module: SaveAll
# * validate    -> "Addon" module: Button1_Click
STEPS = (
    Step(RELOAD, "reload specifications", "ListAll"),
    Step(COMMIT_SAVE, "commit/save specifications", "SaveAll"),
    Step(VALIDATE, "run validation", "Button1_Click"),
)

# The two synthetic "auto" actions each run the same full cycle of commit/save
# -> reload -> commit/save (2x commit/save with a reload in between).
# plan_steps() expands them into the leaf steps above, tagging each leaf with
# its role in the cycle.  They differ only in how the second commit/save is
# judged:
#
# * auto_reload_aa — the second commit/save is expected to fail validation,
#   so that ``Validation Error`` is the success signal (the behaviour the
#   previous ``auto`` action had); see evaluate_step_outcome.
# * auto_apply_aa — the second commit/save uses the normal commit/save rules,
#   so a validation failure is a real failure and the run exits non-zero.
AUTO_COMMIT_SECOND = "auto_commit_second"

AUTO_EXPANSIONS = {
    AUTO_RELOAD_AA: (
        (COMMIT_SAVE, ""),
        (RELOAD, ""),
        (COMMIT_SAVE, AUTO_COMMIT_SECOND),
    ),
    AUTO_APPLY_AA: (
        (COMMIT_SAVE, ""),
        (RELOAD, ""),
        (COMMIT_SAVE, ""),
    ),
}

# All selectable actions, in run order: the leaf steps, then the auto cycles.
ACTIONS = tuple(step.key for step in STEPS) + (AUTO_APPLY_AA, AUTO_RELOAD_AA)


@dataclass(frozen=True)
class Dialog:
    """A modal message box the watcher should recognise and dismiss.

    ``title`` / ``text`` are case-insensitive substrings matched against the
    window caption / window text (an empty string matches anything).
    ``action`` is ``"ok"`` (IDOK) or ``"cancel"`` (IDCANCEL).  ``capture``
    records the full dialog text so the run can report it later.
    """

    key: str
    title: str
    text: str
    action: str
    capture: bool = True


@dataclass
class DialogRecord:
    """A matched-and-dismissed message box, as recorded by the watcher."""

    dialog: Dialog
    captured_text: str
    matched_caption: str


def find_dialog(caption, text, dialogs):
    """Return the first :class:`Dialog` matching ``caption``/``text`` or ``None``.

    Matching is case-insensitive substring matching; an empty ``title`` or
    ``text`` on a dialog matches anything.
    """
    caption = (caption or "").strip()
    text = (text or "").strip()
    for dialog in dialogs:
        if dialog.title and dialog.title.lower() not in caption.lower():
            continue
        if dialog.text and dialog.text.lower() not in text.lower():
            continue
        return dialog
    return None


OVERWRITE_DIALOG = Dialog(
    key="overwrite_prompt",
    title="Overwrite?",
    text="Are you sure you want to load specifications?",
    action="ok",
    capture=True,
)

SAVE_SUCCESS_DIALOG = Dialog(
    key="save_success",
    title="Save Successful",
    text="Your Analysis Specifications have been saved",
    action="ok",
    capture=True,
)

VALIDATION_ERROR_DIALOG = Dialog(
    key="validation_error",
    title="Validation Error",
    text="An error occured while processing your specifications.",
    action="ok",
    capture=True,
)

VALIDATION_SUCCESS_DIALOG = Dialog(
    key="validation_success",
    title="",
    text="Validation Complete! No Errors found!",
    action="ok",
    capture=True,
)

# Which message boxes each action may show.  The watcher clicks them as they
# appear; ``evaluate_step_outcome`` later turns the records into pass/fail.
STEP_DIALOGS = {
    RELOAD: (OVERWRITE_DIALOG,),
    COMMIT_SAVE: (SAVE_SUCCESS_DIALOG, VALIDATION_ERROR_DIALOG),
    VALIDATE: (VALIDATION_SUCCESS_DIALOG, VALIDATION_ERROR_DIALOG),
}

# Control ids used by standard Win32 dialogs (kept as plain ints so the
# button-locating logic stays testable without pywin32).
WIN32_IDOK = 1
WIN32_IDCANCEL = 2
WIN32_IDYES = 6
WIN32_IDNO = 7

# Button captions used to locate the OK/Cancel button on non-standard dialogs
# (VBA UserForms / MSForms), where the button does not carry IDOK/IDCANCEL.
OK_LABELS = ("ok", "yes", "continue", "finish", "done")
CANCEL_LABELS = ("cancel", "no")
BUTTON_CLASS_MARKERS = ("button",)


def find_button_child(children, action):
    """Return the child most likely to be the OK/Cancel button, or ``None``.

    ``children`` is a list of dicts with ``text``, ``class`` and ``id`` keys
    (the shape produced by :func:`_child_windows`).  ``action`` is ``"ok"``
    or ``"cancel"``.  Preference order: a child carrying a standard Win32
    control id (IDOK/IDYES for OK, IDCANCEL/IDNO for Cancel), then a child
    whose caption is a known OK/Yes or Cancel/No label, then any child whose
    window class suggests a button.
    """
    wanted_ids = (
        (WIN32_IDOK, WIN32_IDYES)
        if action == "ok"
        else (WIN32_IDCANCEL, WIN32_IDNO)
    )
    labels = OK_LABELS if action == "ok" else CANCEL_LABELS

    for child in children:
        if child.get("id") in wanted_ids:
            return child
    for child in children:
        label = (child.get("text") or "").strip().lower().rstrip(".")
        if label in labels:
            return child
    for child in children:
        class_name = (child.get("class") or "").lower()
        if any(marker in class_name for marker in BUTTON_CLASS_MARKERS):
            return child
    return None


def plan_steps(actions):
    """Return the ordered leaf steps for the given action keys.

    ``actions`` is a sequence of action keys (the CLI passes a single key per
    run).  The synthetic ``auto_apply_aa`` and ``auto_reload_aa`` keys are
    expanded into their full cycle of leaf steps (commit/save -> reload ->
    commit/save).  Macro names are fixed in ``STEPS``.
    """
    by_key = {step.key: step for step in STEPS}
    planned = []
    for key in actions:
        if key in AUTO_EXPANSIONS:
            for sub_key, context in AUTO_EXPANSIONS[key]:
                sub = by_key[sub_key]
                planned.append(
                    Step(
                        sub.key,
                        f"{key}: {sub.label}",
                        sub.macro,
                        context,
                    )
                )
            continue
        step = by_key[key]
        planned.append(Step(step.key, step.label, step.macro))
    return planned


def _import_win32com():
    """Import ``win32com.client`` lazily.

    Kept inside a function so ``--help`` works on machines without pywin32
    (e.g. this repo's macOS dev box, or the test suite).
    """
    try:
        import win32com.client as win32com_client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 (win32com.client) is required but not installed. "
            "Install it with: pip install pywin32"
        ) from exc
    return win32com_client


def working_copy_path(path, today=None):
    """Path of the automated working copy for ``path``.

    The copy lives next to the original and is named
    ``<stem>-working-automated-YYYY-MM-DD<suffix>``, e.g.
    ``MyAA-working-automated-2026-11-11.xlsm``.  The date is the run date,
    or ``today`` (a ``datetime.date``) when provided for testing.
    """
    path = Path(path)
    today = today or date.today()
    return path.with_name(
        f"{path.stem}-working-automated-{today:%Y-%m-%d}{path.suffix}"
    )


def make_working_copy(path, today=None):
    """Copy ``path`` to its working copy, overwriting any existing copy.

    The original ``path`` is left untouched; the run works only on the
    returned copy, so the VBA macros' repeated saves never touch the
    original.  Returns the working copy path.
    """
    dest = working_copy_path(path, today=today)
    shutil.copy2(path, dest)
    return dest


def wait_for_excel_idle(xl, timeout, poll_interval=IDLE_POLL_INTERVAL):
    """Wait up to ``timeout`` seconds for Excel to report itself idle.

    ``Application.Run`` is synchronous, so by the time this runs the macro has
    normally already returned and Excel is idle; the poll is a best-effort
    safety net for macros that hand work off to Excel asynchronously.
    ``Application.Busy`` / ``Application.Ready`` are not always trustworthy,
    however: after a modal message box has been dismissed from the watcher
    thread (or while Excel finishes tidying up) they can keep reporting
    ``Busy`` even though the macro has finished.  A timeout is therefore a
    warning, not an error — the run continues, and any real problem surfaces
    on the next Excel call.
    """
    deadline = time.monotonic() + timeout
    last_state = ""
    while True:
        try:
            busy = bool(xl.Busy)
            ready = bool(xl.Ready)
            last_state = f"Busy={busy}, Ready={ready}"
        except Exception as exc:
            busy, ready = True, False
            last_state = f"could not read Busy/Ready ({exc})"
        if ready and not busy:
            return
        if time.monotonic() >= deadline:
            message = (
                f"Excel did not report idle within {timeout:g}s ({last_state}); "
                f"continuing because Application.Run has already returned"
            )
            print(f"warning: {message}")
            log_event(f"warning: {message}", EVENT_PHASE_FINISHED)
            return
        time.sleep(poll_interval)


def run_macro(xl, workbook_name, macro, timeout, settle_seconds):
    """Run ``macro`` from ``workbook_name`` and wait for Excel to settle."""
    qualified = f"'{workbook_name}'!{macro}"
    print(f"  -> {qualified}")
    log_event(f"{EVENT_MACRO_STARTED} {macro}", EVENT_PHASE_STARTED)
    xl.Run(qualified)
    wait_for_excel_idle(xl, timeout)
    if settle_seconds:
        time.sleep(settle_seconds)
    log_event(f"{EVENT_MACRO_FINISHED} {macro}", EVENT_PHASE_FINISHED)



class ValidationError(RuntimeError):
    """Raised when the workbook reports a validation failure."""

    def __init__(self, dialog_text, analysis_validation_text=""):
        self.dialog_text = dialog_text
        self.analysis_validation_text = analysis_validation_text
        message = f"validation failed: {dialog_text}"
        if analysis_validation_text:
            message += (
                "\n\n--- AnalysisValidation sheet ---\n"
                + analysis_validation_text
                + "\n----------------------------------"
            )
        super().__init__(message)


def read_data_file_path(workbook):
    """Read the data-file path from ``Main``!C13 and return it as ``Path``.

    Fails with an exception when the cell is empty, the path does not exist,
    or it is not a file.  This is pre-step #0; the returned path is the
    "global config" the rest of the run relies on.
    """
    try:
        worksheet = workbook.Sheets(MAIN_SHEET_NAME)
        raw = worksheet.Range(DATA_FILE_CELL).Value
    except Exception as exc:
        raise RuntimeError(
            f"could not read the data file path from "
            f"'{MAIN_SHEET_NAME}'!{DATA_FILE_CELL}: {exc}"
        ) from exc

    text = str(raw).strip() if raw is not None else ""
    text = text.strip('"')
    if not text:
        raise RuntimeError(
            f"data file path is empty in '{MAIN_SHEET_NAME}'!{DATA_FILE_CELL}"
        )
    path = Path(text)
    if not path.exists():
        raise FileNotFoundError(
            f"data file not found: {path} "
            f"(from '{MAIN_SHEET_NAME}'!{DATA_FILE_CELL})"
        )
    if not path.is_file():
        raise RuntimeError(f"data file path is not a file: {path}")
    return path


def data_file_mtime(path):
    """Modification time (``st_mtime``) of the data file."""
    return Path(path).stat().st_mtime


def verify_data_file_recently_edited(path, before_mtime):
    """Raise unless ``path`` was modified after ``before_mtime``.

    This is the "the file captured in step #0 was recently edited" check: a
    successful commit/save must have written the data file, which advances
    its modification time past the value captured just before the save ran.
    """
    after_mtime = data_file_mtime(path)
    if after_mtime <= before_mtime:
        raise RuntimeError(
            f"data file was not edited by the save: {path} "
            f"(mtime {after_mtime!r} not newer than {before_mtime!r})"
        )


def read_sheet_text(workbook, sheet_name):
    """Return the used range of ``sheet_name`` as tab/newline-separated text."""
    worksheet = workbook.Sheets(sheet_name)
    used = worksheet.UsedRange
    if used is None:
        return ""
    value = used.Value
    if value is None:
        return ""
    if not isinstance(value, tuple):
        value = ((value,),)
    lines = []
    for row in value:
        if not isinstance(row, tuple):
            row = (row,)
        cells = ["" if cell is None else str(cell) for cell in row]
        lines.append("\t".join(cells))
    return "\n".join(lines)


def format_validation_details(dialog_text, sheet_text=""):
    """Combine a validation message box with the AnalysisValidation sheet text."""
    if sheet_text:
        return f"{dialog_text}\n--- AnalysisValidation sheet ---\n{sheet_text}"
    return dialog_text


def write_log_row(log_path, program, aa_file, data_file, result, validation_details, now=None):
    """Append one CSV log row, writing the header first if the file is new/empty."""
    now = now or datetime.now()
    row = [
        now.strftime("%Y-%m-%d %H:%M:%S"),
        program,
        str(aa_file),
        str(data_file) if data_file is not None else "",
        result,
        validation_details,
    ]
    log_path = Path(log_path)
    write_header = not log_path.exists() or log_path.stat().st_size == 0
    with open(log_path, "a", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        if write_header:
            writer.writerow(LOG_HEADER)
        writer.writerow(row)


def write_event_log_row(event_log_path, event, phase, now=None):
    """Append one event-log CSV row, writing the header first if new/empty.

    Columns are ``datetime`` (with microseconds so rapid events stay in
    order), ``event``, and ``phase`` (``started`` / ``finished``).  Appends to
    an existing file; a new or empty file gets the header first.
    """
    now = now or datetime.now()
    row = [
        now.strftime("%Y-%m-%d %H:%M:%S.%f"),
        event,
        phase,
    ]
    event_log_path = Path(event_log_path)
    write_header = (
        not event_log_path.exists() or event_log_path.stat().st_size == 0
    )
    with _EVENT_LOG_LOCK:
        with open(event_log_path, "a", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            if write_header:
                writer.writerow(EVENT_LOG_HEADER)
            writer.writerow(row)


def log_event(event, phase):
    """Record ``event`` to the configured event log, if one is configured.

    No-op when :data:`EVENT_LOG_PATH` is ``None`` (the default outside a CLI
    run).  Used from both the main thread and the dialog-watcher thread; the
    append is serialised by ``_EVENT_LOG_LOCK``.
    """
    if EVENT_LOG_PATH is None:
        return
    write_event_log_row(EVENT_LOG_PATH, event, phase)


def _first_record(records, key):
    """Return the first :class:`DialogRecord` for ``key``, or ``None``."""
    for record in records:
        if record.dialog.key == key:
            return record
    return None


def evaluate_step_outcome(step, records, before_mtime, workbook):
    """Check a finished step's message-box records; return details or raise.

    Returns the step's validation-details text (``""`` when there was no
    validation failure) on success, and raises on failure.

    Success is judged by which message box the macro showed:

    * commit/save — ``Save Successful``, plus the data file was modified
    * validate   — "Validation Complete! No Errors found!"
    * auto_reload_aa's second commit/save — its expected ``Validation Error``
      (see below)

    A ``Validation Error`` box normally means validation failed; its text and
    the ``AnalysisValidation`` sheet are captured into the raised error.  If
    neither the success nor the failure box was seen, the run fails because
    success could not be confirmed.

    The one exception is the second commit/save of the ``auto_reload_aa``
    cycle (``step.context == AUTO_COMMIT_SECOND``): that commit/save is
    expected to fail validation, so its ``Validation Error`` box is treated
    as success rather than failure (its details are returned for the log).
    """
    if step.key == RELOAD:
        return ""

    if step.context == AUTO_COMMIT_SECOND:
        error_record = _first_record(records, "validation_error")
        if error_record is not None:
            log_event(EVENT_VALIDATION_FAILURE, EVENT_PHASE_FINISHED)
            print(
                "auto_reload_aa: second commit/save failed validation as "
                "expected (success)"
            )
            return format_validation_details(
                error_record.captured_text,
                read_sheet_text(workbook, ANALYSIS_VALIDATION_SHEET),
            )
        if _first_record(records, "save_success") is None:
            raise RuntimeError(
                "auto_reload_aa: second commit/save completed without an "
                "expected message ('Save Successful' or the expected "
                "'Validation Error')"
            )
        return ""

    error_record = _first_record(records, "validation_error")
    if error_record is not None:
        log_event(EVENT_VALIDATION_FAILURE, EVENT_PHASE_FINISHED)
        sheet_text = read_sheet_text(workbook, ANALYSIS_VALIDATION_SHEET)
        raise ValidationError(error_record.captured_text, sheet_text)

    if step.key == COMMIT_SAVE:
        if _first_record(records, "save_success") is None:
            raise RuntimeError(
                "commit/save completed without the success message "
                "'Your Analysis Specifications have been saved'"
            )
        verify_data_file_recently_edited(DATA_FILE_PATH, before_mtime)
        return ""

    if step.key == VALIDATE:
        if _first_record(records, "validation_success") is None:
            raise RuntimeError(
                "validation completed without the success message "
                "'Validation Complete! No Errors found!'"
            )
        return ""


def _import_win32gui():
    """Import the Win32 GUI helpers lazily (Windows only)."""
    try:
        import win32con
        import win32gui
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 (win32gui) is required but not installed. "
            "Install it with: pip install pywin32"
        ) from exc
    return win32gui, win32con


def _window_text(win32gui, hwnd):
    """Return a window's caption plus the text of its child controls."""
    parts = [win32gui.GetWindowText(hwnd)]
    try:
        def collect(child, _extra):
            child_text = win32gui.GetWindowText(child)
            if child_text:
                parts.append(child_text)
            return True

        win32gui.EnumChildWindows(hwnd, collect, None)
    except Exception:
        pass
    return "\n".join(parts)


def _enumerate_visible_windows(win32gui):
    """Return ``(hwnd, caption, full_text)`` for visible top-level windows."""
    windows = []

    def collect(hwnd, _extra):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        caption = win32gui.GetWindowText(hwnd)
        text = _window_text(win32gui, hwnd)
        windows.append((hwnd, caption, text))
        return True

    win32gui.EnumWindows(collect, None)
    return windows


def _window_class(win32gui, hwnd):
    """Return the window class name for ``hwnd`` (``""`` on failure)."""
    try:
        return win32gui.GetClassName(hwnd) or ""
    except Exception:
        return ""


def _child_windows(win32gui, hwnd):
    """Return the children of ``hwnd`` as ``{hwnd, text, class, id}`` dicts.

    ``id`` is the control id (``GetDlgCtrlID``); it is 0 for controls that do
    not use the standard Win32 dialog control-id scheme (e.g. VBA UserForm
    buttons).
    """
    children = []

    def collect(child, _extra):
        try:
            children.append(
                {
                    "hwnd": child,
                    "text": win32gui.GetWindowText(child) or "",
                    "class": win32gui.GetClassName(child) or "",
                    "id": win32gui.GetDlgCtrlID(child),
                }
            )
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, collect, None)
    except Exception:
        pass
    return children


def _send_message(win32gui, hwnd, message, wparam, lparam):
    """Send one window message, ignoring failures (the watcher retries)."""
    try:
        return win32gui.SendMessage(hwnd, message, wparam, lparam)
    except Exception:
        return 0


def _click_dialog(win32gui, win32con, hwnd, action):
    """Dismiss ``hwnd`` by clicking its OK/Cancel button.

    Standard Win32 message boxes (window class ``#32770``) expose the button
    through ``GetDlgItem(IDOK/IDCANCEL)``; those are clicked with ``BM_CLICK``
    plus the matching ``WM_COMMAND``.  Yes/No (or Retry/Cancel) message boxes
    use ``IDYES``/``IDNO`` instead, so the button is located among the child
    windows.  Non-standard dialogs (VBA UserForms / MSForms) do not carry the
    standard ids and ignore ``BM_CLICK``, so their button is located by control
    id / caption / class and clicked with a real mouse click.  ``SendMessage``
    is used so each click is processed before the next poll; the watcher
    retries while the dialog is still visible.
    """
    command = win32con.IDOK if action == "ok" else win32con.IDCANCEL
    class_name = _window_class(win32gui, hwnd)
    children = _child_windows(win32gui, hwnd)

    if class_name == "#32770":
        target = None
        try:
            button = win32gui.GetDlgItem(hwnd, command)
        except Exception:
            button = 0
        if button:
            target = {"hwnd": button, "id": command}
        else:
            # Yes/No (or Retry/Cancel) message boxes use IDYES/IDNO rather
            # than IDOK/IDCANCEL; find the button among the children instead.
            target = find_button_child(children, action)

        if target is not None:
            target_hwnd = target["hwnd"]
            _send_message(win32gui, target_hwnd, win32con.BM_CLICK, 0, 0)
            _send_message(
                win32gui,
                hwnd,
                win32con.WM_COMMAND,
                target["id"],
                target_hwnd,
            )
            return

        _send_message(win32gui, hwnd, win32con.WM_COMMAND, command, 0)
        return

    target = find_button_child(children, action)
    if target is not None:
        target_hwnd = target["hwnd"]
        _send_message(
            win32gui,
            target_hwnd,
            win32con.WM_LBUTTONDOWN,
            win32con.MK_LBUTTON,
            0,
        )
        _send_message(win32gui, target_hwnd, win32con.WM_LBUTTONUP, 0, 0)
        _send_message(win32gui, target_hwnd, win32con.BM_CLICK, 0, 0)
        _send_message(
            win32gui,
            hwnd,
            win32con.WM_COMMAND,
            target["id"],
            target_hwnd,
        )
        return

    # Last resort: press Enter to trigger the dialog's default button, then
    # send the classic IDOK/IDCANCEL command.
    _send_message(win32gui, hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    _send_message(win32gui, hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    _send_message(win32gui, hwnd, win32con.WM_COMMAND, command, 0)


class DialogWatcher:
    """Watch for known message boxes and click their buttons in a thread.

    ``Application.Run`` blocks while a modal ``MsgBox`` is on screen, so the
    dialog must be dismissed from another thread.  Start the watcher before
    calling the macro; it polls for the configured dialogs and clicks their
    OK/Cancel button (standard ``#32770`` message boxes via ``BM_CLICK`` /
    ``WM_COMMAND``, VBA UserForms via a real mouse click on the button).
    Dismissed dialogs are collected in :attr:`records`.
    """

    def __init__(self, dialogs, poll_interval=DIALOG_POLL_INTERVAL):
        self.dialogs = tuple(dialogs)
        self.poll_interval = poll_interval
        self.records = []
        self._stop = threading.Event()
        self._thread = None

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        self.stop()
        return False

    def start(self):
        if not self.dialogs:
            return
        self._stop.clear()
        self._thread = threading.Thread(
            target=self._run, name="aa-dialog-watcher", daemon=True
        )
        self._thread.start()

    def stop(self):
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=self.poll_interval * 2 + 0.5)
            self._thread = None

    def _run(self):
        win32gui, win32con = _import_win32gui()
        recorded = set()
        while not self._stop.is_set():
            for hwnd, caption, text in _enumerate_visible_windows(win32gui):
                dialog = find_dialog(caption, text, self.dialogs)
                if dialog is None:
                    continue
                if hwnd not in recorded:
                    recorded.add(hwnd)
                    class_name = _window_class(win32gui, hwnd)
                    children = _child_windows(win32gui, hwnd)
                    child_summary = [
                        (child["text"], child["class"], child["id"])
                        for child in children
                    ]
                    log_event(
                        f"{EVENT_DIALOG_CAPTURED} {caption!r}",
                        EVENT_PHASE_FINISHED,
                    )
                    print(
                        f"dialog watcher: dismissing {dialog.key!r} ({caption!r}) "
                        f"[class={class_name!r}, children={child_summary}]"
                    )
                    if dialog.capture:
                        self.records.append(DialogRecord(dialog, text, caption))
                _click_dialog(win32gui, win32con, hwnd, dialog.action)
            time.sleep(self.poll_interval)


def run_step(xl, workbook, workbook_name, step, idle_timeout, settle_seconds):
    """Run one step's macro, dismissing and recording its message boxes.

    On success the workbook is explicitly saved, so its on-disk state is clean
    and it can be closed without Excel prompting.  Returns the step's
    validation-details text (``""`` if none); raises on failure.
    """
    dialogs = STEP_DIALOGS.get(step.key, ())

    before_mtime = None
    if step.key == COMMIT_SAVE and DATA_FILE_PATH is not None:
        before_mtime = data_file_mtime(DATA_FILE_PATH)

    with DialogWatcher(dialogs) as watcher:
        run_macro(xl, workbook_name, step.macro, idle_timeout, settle_seconds)

    details = evaluate_step_outcome(step, watcher.records, before_mtime, workbook)
    log_event(f"{EVENT_SAVE} {workbook_name}", EVENT_PHASE_FINISHED)
    workbook.Save()
    return details


def run_workbook(
    path,
    actions,
    log_path=None,
    aa_file=None,
    visible=False,
    leave_open=False,
    idle_timeout=IDLE_TIMEOUT_DEFAULT,
    settle_seconds=SETTLE_SECONDS_DEFAULT,
):
    """Open ``path`` and run each action in order, logging one CSV row each.

    ``actions`` is a sequence of top-level action keys (``auto_apply_aa`` and
    ``auto_reload_aa`` run their full cycle as one logged action).  Before
    anything runs, the data-file path is read from ``Main``!C13 into the
    module global ``DATA_FILE_PATH`` (pre-step
    #0).  After each successful step the workbook is saved explicitly, so it
    is left in a clean, closable state.  Each action's result and any
    validation details are appended to ``log_path``; a failure is logged and
    then re-raised so the run still stops with an exception.  Returns the
    number of leaf steps run.  When ``leave_open`` is true the workbook and
    Excel instance are left open (for debugging); otherwise they are closed in
    a ``finally`` block even if an action raises.
    """
    global DATA_FILE_PATH

    win32com_client = _import_win32com()
    xl = None
    workbook = None
    try:
        xl = win32com_client.DispatchEx("Excel.Application")
        xl.Visible = visible
        xl.DisplayAlerts = False

        abs_path = str(Path(path).resolve())
        workbook = xl.Workbooks.Open(
            Filename=abs_path,
            UpdateLinks=UPDATE_LINKS,
        )
        workbook.Activate()
        workbook_name = Path(path).name

        # Pre-step #0: the external data file the macros read/write.
        DATA_FILE_PATH = read_data_file_path(workbook)
        print(f"data file: {DATA_FILE_PATH}")

        aa_file = aa_file if aa_file is not None else Path(path)

        ran = 0
        for action_key in actions:
            program = PROGRAM_NAMES.get(action_key, action_key)
            started = datetime.now()
            try:
                validation_details = ""
                for step in plan_steps([action_key]):
                    print(f"Step {step.label}:")
                    step_details = run_step(
                        xl, workbook, workbook_name, step, idle_timeout, settle_seconds
                    )
                    if step_details:
                        validation_details = step_details
                    ran += 1
            except ValidationError as exc:
                if log_path is not None:
                    write_log_row(
                        log_path,
                        program,
                        aa_file,
                        DATA_FILE_PATH,
                        str(exc),
                        format_validation_details(
                            exc.dialog_text, exc.analysis_validation_text
                        ),
                        now=started,
                    )
                raise
            except Exception as exc:
                if log_path is not None:
                    write_log_row(
                        log_path,
                        program,
                        aa_file,
                        DATA_FILE_PATH,
                        str(exc),
                        "",
                        now=started,
                    )
                raise
            if log_path is not None:
                write_log_row(
                    log_path,
                    program,
                    aa_file,
                    DATA_FILE_PATH,
                    "success",
                    validation_details,
                    now=started,
                )

        return ran
    finally:
        if workbook is not None and not leave_open:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if xl is not None and not leave_open:
            try:
                xl.Quit()
            except Exception:
                pass


def build_parser():
    parser = argparse.ArgumentParser(
        prog="run_aa_workbook.py",
        description=(
            "Run the AA workbook's VBA actions (reload, commit/save, "
            "validate, auto_apply_aa, auto_reload_aa) via Excel.Application."
        ),
    )
    parser.add_argument("path", help="Excel file (.xlsx / .xlsm) to open and run")
    parser.add_argument(
        "--program",
        required=True,
        choices=list(ACTIONS),
        help="which program to run (exactly one per run); choices: %(choices)s",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="show the Excel window (default: run hidden)",
    )
    parser.add_argument(
        "--leave-open",
        action="store_true",
        help="keep Excel and the workbook open after the run (for debugging)",
    )
    parser.add_argument(
        "--idle-timeout",
        type=float,
        default=IDLE_TIMEOUT_DEFAULT,
        help=(
            "max seconds to wait for Excel to report idle after a macro; "
            "a timeout only warns and continues because Application.Run is "
            f"synchronous (default {IDLE_TIMEOUT_DEFAULT})"
        ),
    )
    parser.add_argument(
        "--settle",
        type=float,
        default=SETTLE_SECONDS_DEFAULT,
        help=(
            "extra seconds to wait after Excel reports idle "
            f"(default {SETTLE_SECONDS_DEFAULT})"
        ),
    )
    return parser


def main(argv=None):
    global EVENT_LOG_PATH

    parser = build_parser()
    args = parser.parse_args(argv)

    path = Path(args.path)
    if path.is_dir():
        print(f"expected a file, got a directory: {path}", file=sys.stderr)
        return 1
    if not path.exists():
        print(f"file not found: {path}", file=sys.stderr)
        return 1
    if path.suffix.lower() not in SUPPORTED_SUFFIXES:
        print(
            f"unsupported file type: {path} (expected .xlsx / .xlsm)",
            file=sys.stderr,
        )
        return 1

    actions = [args.program]

    steps = plan_steps(actions)
    print(f"{parser.prog}: {path}")
    for step in steps:
        print(f"  - {step.label}")

    EVENT_LOG_PATH = path.parent / EVENT_LOG_FILENAME
    print(f"event log: {EVENT_LOG_PATH}")

    try:
        log_event(EVENT_PROGRAM_STARTED, EVENT_PHASE_STARTED)

        working_path = make_working_copy(path)
        print(f"working copy: {working_path}")

        log_path = path.parent / LOG_FILENAME
        print(f"log: {log_path}")

        ran = run_workbook(
            working_path,
            actions,
            log_path=log_path,
            aa_file=path,
            visible=args.visible,
            leave_open=args.leave_open,
            idle_timeout=args.idle_timeout,
            settle_seconds=args.settle,
        )
        print(f"{parser.prog}: finished; {ran} step(s) run")
        return 0
    except Exception as exc:
        log_event(f"{EVENT_EXCEPTION}: {exc}", EVENT_PHASE_FINISHED)
        print("", file=sys.stderr)
        print("Stack trace:", file=sys.stderr)
        print("", file=sys.stderr)
        traceback.print_exc()
        print("", file=sys.stderr)
        print("Error:", file=sys.stderr)
        print(f"{STDOUT_COLOR_RED}{exc}{STDOUT_COLOR_RESET}", file=sys.stderr)
        return 1
    finally:
        log_event(EVENT_PROGRAM_FINISHED, EVENT_PHASE_FINISHED)
        EVENT_LOG_PATH = None


if __name__ == "__main__":
    sys.exit(main())

