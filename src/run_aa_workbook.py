"""Run the AA (Excel) workflow steps automatically via Excel.Application.

The AA workbook ships a set of buttons that drive its VBA workflow:

* apply definitions
* load definitions
* run validation

Clicking them by hand and waiting for each one to finish is slow and
error-prone.  This script opens the workbook with ``Excel.Application``
(pywin32 / ``win32com.client``) and calls the same VBA procedures those
buttons call, in order, waiting for Excel to go idle after each one.

Macro names
-----------
The entry points below (``ApplyDefinitions``, ``LoadDefinitions``,
``RunValidation``) are the likely procedure names, but they are plain VBA
procedure names and must match the real ones in the workbook.  Adjust the
``STEPS`` table below, or pass ``--macro-...`` flags, if your workbook names
them differently.  The macros must be ``Public`` subs taking no arguments and
must not show message boxes or other modal UI, because there is nobody there
to click them.

Windows only
------------
This uses COM automation, so it runs on Windows with Excel installed and
``pywin32`` available (``pip install pywin32``).  On other platforms the
import is delayed so ``--help`` and ``--dry-run`` still work.
"""

import argparse
import sys
import time
import traceback
from dataclasses import dataclass
from pathlib import Path


STDOUT_COLOR_RED = "\033[31m"
STDOUT_COLOR_RESET = "\033[0m"

SUPPORTED_SUFFIXES = (".xlsx", ".xlsm")

# Seconds to wait, after calling a macro, for Excel to report itself idle
# again before considering the step finished (and before raising an error).
IDLE_TIMEOUT_DEFAULT = 300
IDLE_POLL_INTERVAL = 0.25
SETTLE_SECONDS_DEFAULT = 1.0

# 0 = do not update links on open (and never prompt); use 3 to update all.
UPDATE_LINKS = 0


@dataclass(frozen=True)
class Step:
    """One VBA step of the AA run."""

    key: str
    label: str
    macro: str


# The macro names are the VBA entry points behind the AA buttons.  Edit them
# here, or override at runtime with the matching ``--macro-<step>`` flag.
STEPS = (
    Step("apply_definitions", "apply definitions", "ApplyDefinitions"),
    Step("load_definitions", "load definitions", "LoadDefinitions"),
    Step("run_validation", "run validation", "RunValidation"),
)


def plan_steps(steps=STEPS, skip=(), macro_overrides=None):
    """Return the ordered steps to run after applying ``skip`` and overrides.

    ``skip`` is a collection of step keys to drop.  ``macro_overrides`` maps a
    step key to a replacement macro name (used for the ``--macro-...`` flags).
    """
    macro_overrides = macro_overrides or {}
    skip = set(skip or ())
    planned = []
    for step in steps:
        if step.key in skip:
            continue
        planned.append(
            Step(step.key, step.label, macro_overrides.get(step.key, step.macro))
        )
    return planned


def _import_win32com():
    """Import ``win32com.client`` lazily.

    Kept inside a function so ``--help`` and ``--dry-run`` work on machines
    without pywin32 (e.g. this repo's macOS dev box, or the test suite).
    """
    try:
        import win32com.client as win32com_client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 (win32com.client) is required but not installed. "
            "Install it with: pip install pywin32"
        ) from exc
    return win32com_client


def wait_for_excel_idle(xl, timeout, poll_interval=IDLE_POLL_INTERVAL):
    """Block until Excel reports ``Ready`` and not ``Busy``, or raise.

    ``Application.Run`` is synchronous, so normally the macro has already
    returned by the time we get here; this loop is a safety net for macros
    that hand work off to Excel asynchronously.
    """
    deadline = time.monotonic() + timeout
    while True:
        try:
            busy = bool(xl.Busy)
            ready = bool(xl.Ready)
        except Exception:
            busy, ready = True, False
        if ready and not busy:
            return
        if time.monotonic() >= deadline:
            raise RuntimeError(
                f"timed out after {timeout}s waiting for Excel to go idle"
            )
        time.sleep(poll_interval)


def run_macro(xl, workbook_name, macro, timeout, settle_seconds):
    """Run ``macro`` from ``workbook_name`` and wait for Excel to settle."""
    qualified = f"'{workbook_name}'!{macro}"
    print(f"  -> {qualified}")
    xl.Run(qualified)
    wait_for_excel_idle(xl, timeout)
    if settle_seconds:
        time.sleep(settle_seconds)



def run_workbook(
    path,
    steps,
    visible=False,
    read_only=False,
    save=True,
    leave_open=False,
    idle_timeout=IDLE_TIMEOUT_DEFAULT,
    settle_seconds=SETTLE_SECONDS_DEFAULT,
):
    """Open ``path``, run each step's macro in order, then save and close.

    Returns the number of steps that ran.  When ``leave_open`` is true the
    workbook and Excel instance are left open (for debugging); otherwise they
    are closed in a ``finally`` block even if a step raises.
    """
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
            ReadOnly=read_only,
            UpdateLinks=UPDATE_LINKS,
        )
        workbook.Activate()
        workbook_name = Path(path).name

        ran = 0
        for step in steps:
            print(f"Step {step.label}:")
            run_macro(xl, workbook_name, step.macro, idle_timeout, settle_seconds)
            ran += 1

        if save and not read_only:
            print("saving workbook")
            workbook.Save()

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
            "Run the AA workbook's VBA steps (apply definitions, load "
            "definitions, validation) via Excel.Application."
        ),
    )
    parser.add_argument("path", help="Excel file (.xlsx / .xlsm) to open and run")
    parser.add_argument(
        "--visible",
        action="store_true",
        help="show the Excel window (default: run hidden)",
    )
    parser.add_argument(
        "--read-only",
        action="store_true",
        help="open the workbook read-only and never save",
    )
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="do not save the workbook after the run",
    )
    parser.add_argument(
        "--leave-open",
        action="store_true",
        help="keep Excel and the workbook open after the run (for debugging)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="print the planned steps without launching Excel",
    )
    parser.add_argument(
        "--skip",
        action="append",
        choices=[step.key for step in STEPS],
        help="skip a step (repeatable); choices: %(choices)s",
    )
    for step in STEPS:
        parser.add_argument(
            f"--macro-{step.key.replace('_', '-')}",
            default=None,
            metavar="NAME",
            help=(
                f"override the VBA macro name for '{step.label}' "
                f"(default: {step.macro!r})"
            ),
        )
    parser.add_argument(
        "--idle-timeout",
        type=float,
        default=IDLE_TIMEOUT_DEFAULT,
        help=(
            "seconds to wait for Excel to go idle after a macro "
            f"(default {IDLE_TIMEOUT_DEFAULT})"
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

    macro_overrides = {}
    for step in STEPS:
        value = getattr(args, f"macro_{step.key}")
        if value:
            macro_overrides[step.key] = value

    steps = plan_steps(STEPS, skip=args.skip, macro_overrides=macro_overrides)
    if not steps:
        print("no steps to run (all skipped)")
        return 0

    print(f"{parser.prog}: {path}")
    for step in steps:
        print(f"  - {step.label} ({step.macro})")

    if args.dry_run:
        print("dry run: nothing executed")
        return 0

    try:
        ran = run_workbook(
            path,
            steps,
            visible=args.visible,
            read_only=args.read_only,
            save=not args.no_save,
            leave_open=args.leave_open,
            idle_timeout=args.idle_timeout,
            settle_seconds=args.settle,
        )
        print(f"{parser.prog}: finished; {ran} step(s) run")
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

