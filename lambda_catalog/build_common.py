"""Build scaffolding shared by the production build script and other
Excel-required build drivers (build_test_models.py, build_demo_workbook.py).

``build_production.py`` builds the unified ``Lambda_Library.xlsx`` workbook
which carries the Regression template, the Univariate template, the LAMBDA
function catalog, reference sheets, and the data sheets. It follows a
three-phase structure (write sheets + sync names; recalculate and save;
optional deep verify), and needs Excel-open retry handling and the same
"open headless Excel, run CalculateFullRebuild, save under a chosen
calculation mode" recalc step. Those helpers live here so the driver scripts
stay thin and do not duplicate the retry/recalc logic.
"""
from __future__ import annotations

import argparse
import contextlib
import io
import sys
import time
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import TextIO

import xlwings as xw

from lambda_catalog.workbook_builder import (
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
    NameSyncResult,
)
from lambda_catalog.workbook_helpers import (
    LOCK_HINT,
    OPEN_WORKBOOK_ERRORS,
    raise_excel_access_error,
)

# Phrases that indicate Excel's COM server disappeared mid-call rather than a
# genuine build error — worth one retry in a fresh Excel instance.
_RPC_FAILURE_PHRASES = (
    "remote procedure call failed",
    "rpc server is unavailable",
)


def _is_rpc_failure(exc: BaseException) -> bool:
    """Return True when Excel's COM server disappeared mid-call."""
    message = str(exc).lower()
    return any(phrase in message for phrase in _RPC_FAILURE_PHRASES)


def _close_workbook_quietly(workbook: xw.Book | None) -> None:
    """Best-effort close; Excel may already be gone after an RPC failure."""
    if workbook is None:
        return
    try:
        workbook.close()
    except OPEN_WORKBOOK_ERRORS:
        pass


def _quit_app_quietly(app: xw.App | None) -> None:
    """Best-effort quit; cleanup must not mask the original Excel error."""
    if app is None:
        return
    try:
        app.quit()
    except OPEN_WORKBOOK_ERRORS:
        pass


# Where a driver's terminal transcript is archived. Every Excel-required
# driver writes here — build_production.py, build_test_models.py and
# build_demo_workbook.py — because none of them can run on the GitHub-hosted
# Linux CI, so this directory is the cross-tool substitute for the CI log a
# failed run would otherwise leave only in somebody's terminal scrollback.
# It is tracked, not gitignored: the transcript is committed on the branch
# that did the work (see excel-only-runs/README.md for when to retire one).
# The directory is named after what produced it (Excel-only run artifact)
# rather than where (a developer's box), so a contributor who has never run
# a Windows-only verifier build still knows what belongs here.
RUN_LOG_DIR_NAME = "excel-only-runs"
RUN_LOG_FILE_SUFFIX = ".log"


# Floor shared by every grid-size surface: the --beta-grid-size flag here, and
# the in-sheet Grid Points Validation and conditional format in
# write_sheet_univariate (_MIN_GRID_POINTS).  A stage's Step cell is
# (Max-Min)/(N-1), so N=1 divides by zero, and a one-point grid searches
# nothing.  Full_Factorial itself tolerates N=1 via its MAX(1,N-1) divisor;
# the sheet built around it does not.
MIN_GRID_POINTS = 2


def positive_grid_size(value: str) -> int:
    """argparse type for a grid-size flag: a whole number >= MIN_GRID_POINTS.

    Without this an accidental ``--beta-grid-size 0`` builds a workbook whose
    Step cells are #DIV/0! and whose Full_Factorial grids are degenerate — a
    broken artifact produced without a single error message.
    """
    try:
        parsed = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"expected a whole number, got {value!r}"
        ) from None
    if parsed < MIN_GRID_POINTS:
        raise argparse.ArgumentTypeError(
            f"must be at least {MIN_GRID_POINTS} (the Step cell divides by "
            f"grid size - 1), got {parsed}"
        )
    return parsed


class _Tee(io.TextIOBase):
    """Write to two streams at once, flushing both."""

    def __init__(self, primary: TextIO, secondary: TextIO) -> None:
        self._primary = primary
        self._secondary = secondary

    def write(self, s: str) -> int:
        self._secondary.write(s)
        self._secondary.flush()
        return self._primary.write(s)

    def flush(self) -> None:
        self._primary.flush()
        self._secondary.flush()


def run_log_path(root_dir: Path, script_name: str, argv: list[str]) -> Path:
    """Return the archive path for one run's transcript.

    Named after the script and the flags it ran with — ``build_test_models
    verify include heavy.log`` — so a directory listing reads as a list of
    what was actually run, and two different invocations do not overwrite
    each other's evidence. That is the convention the first hand-uploaded
    logs already used; this just stops it being a manual step.
    """
    flags = [
        token.lstrip("-").replace("-", " ")
        for token in argv
        if token.startswith("--")
    ]
    stem = " ".join([Path(script_name).stem, *flags]).strip()
    return root_dir / RUN_LOG_DIR_NAME / f"{stem}{RUN_LOG_FILE_SUFFIX}"


@contextlib.contextmanager
def tee_run_log(log_path: Path, command: str) -> Iterator[Path]:
    """Mirror stdout AND stderr into ``log_path`` for the duration.

    stderr matters as much as stdout here: the failure this is most likely
    to be capturing is a `pywintypes.com_error` traceback, which never
    touches stdout. The first hand-captured log of a failed run only
    contained its traceback because a human was copying the terminal.

    The file is flushed on every write, so a run that hangs or is killed
    still leaves everything it had reached — the opposite of buffering,
    and the whole point when the thing being diagnosed is a build that
    stops partway.
    """
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with open(log_path, "w", encoding="utf-8") as handle:
        real_stdout, real_stderr = sys.stdout, sys.stderr
        sys.stdout = _Tee(real_stdout, handle)
        sys.stderr = _Tee(real_stderr, handle)
        try:
            print(command)
            yield log_path
        except BaseException:
            # Land the traceback in the file before the streams are
            # restored, so an aborted run archives its own cause.
            import traceback

            traceback.print_exc()
            raise
        finally:
            sys.stdout, sys.stderr = real_stdout, real_stderr


def print_name_sync_summary(result: NameSyncResult) -> None:
    """Print the workbook-scope sync counts every build script reports.

    ``Removed names`` and ``Skipped names`` are the interesting two: the first
    counts workbook-scoped residue the sync cleared out (a name the catalog
    retired), the second names catalog entries this workbook cannot carry
    because they reference a worksheet it does not have. Both are normal on
    the first build after a change and should settle to 0 / none on a rebuild.
    """
    print(f"Created names: {result.created}")
    print(f"Updated names: {result.updated}")
    print(f"Removed names: {result.removed}")
    if result.skipped:
        print(
            "Skipped names: "
            + ", ".join(result.skipped)
            + " (reference a worksheet this workbook does not have)"
        )


def set_calculate_before_save(app: xw.App, value: bool) -> bool | None:
    """Set Excel's "recalculate workbook before saving", returning the old value.

    Under Manual calculation this setting is what decides whether
    ``workbook.save()`` triggers a full calculation. It has to be off for a
    genuinely calculation-free build; leaving it off would change the user's
    Excel for every later session, so callers restore the returned value.

    ``Application.CalculateBeforeSave`` is application-level and
    environment-dependent, so a failure to read or write it is not a build
    error — it just means there was nothing to suppress and nothing to
    restore.

    Parameters
    ----------
    app : xw.App
        The Excel application whose setting to change.
    value : bool
        The new setting.

    Returns
    -------
    bool or None
        The previous setting, or None when Excel would not report or accept
        it — in which case the caller has nothing to restore.
    """
    try:
        previous = bool(app.api.CalculateBeforeSave)
        app.api.CalculateBeforeSave = value
        return previous
    except Exception:  # pylint: disable=broad-except
        return None


def _recalculate_and_save(
    workbook_path: Path,
    *,
    calc_mode: int = XL_CALCULATION_AUTOMATIC,
) -> None:
    """Fully calculate the workbook after name sync, then save under ``calc_mode``.

    Opens a headless Excel, sets Manual mode (so the open does not trigger a
    full recalc), runs ``CalculateFullRebuild`` (which rebuilds the dependency
    tree and evaluates every formula, including the Univariate grid spills
    when present), sets the workbook's persisted calculation mode to
    ``calc_mode`` (full Automatic for both artifacts post-split), and saves.

    Parameters
    ----------
    workbook_path : Path
        Workbook to open, recalculate, and save in place.
    calc_mode : int
        Excel ``XlCalculation`` constant to persist as the workbook's
        calculation mode. Both production artifacts use full Automatic
        (``XL_CALCULATION_AUTOMATIC``); the Univariate grid spills recalculate
        on edit under that mode.
    """
    app: xw.App | None = None
    workbook: xw.Book | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.api.DisplayAlerts = False
        app.api.AskToUpdateLinks = False

        workbook = app.books.open(str(workbook_path))
        try:
            app.api.Calculation = XL_CALCULATION_MANUAL
            app.api.CalculateFullRebuild()
            app.api.Calculation = calc_mode
            workbook.save(str(workbook_path))
        finally:
            _close_workbook_quietly(workbook)
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "recalculate and save", exc)
    finally:
        _quit_app_quietly(app)


def _retry_on_open(
    action_label: str,
    fn: Callable[[], None],
    *,
    retry_rpc: bool = False,
    max_rpc_retries: int = 1,
) -> None:
    """Call fn(); retry open-workbook locks and optionally dropped COM sessions.

    A workbook locked open in Excel by the user (the most common cause of a
    mid-build failure once the write itself is correct) prompts an interactive
    retry rather than aborting the whole multi-minute build. When ``retry_rpc``
    is set, a dropped COM session also gets one automatic retry in a fresh
    Excel instance before propagating.
    """
    rpc_retries = 0
    while True:
        try:
            fn()
            return
        except RuntimeError as exc:
            if retry_rpc and _is_rpc_failure(exc) and rpc_retries < max_rpc_retries:
                rpc_retries += 1
                print(
                    "\nExcel's COM session dropped; "
                    "retrying in a fresh Excel instance...",
                    file=sys.stderr,
                    flush=True,
                )
                if sys.stdin.isatty():
                    time.sleep(2)
                continue
            if LOCK_HINT not in str(exc):
                raise
            if not sys.stdin.isatty():
                raise
            input(
                f"\n{action_label} — close it in Excel "
                "and press Enter to retry (or Ctrl+C to cancel): "
            )


def _lock_file_owner(workbook_path: Path) -> str | None:
    """Return the owner name from Excel's ``~$`` sidecar, or None.

    Excel drops a hidden ``~$<name>.xlsx`` next to an open workbook whose first
    byte is the owner-name length and whose remainder is that name in UTF-16LE.
    The format is undocumented and varies across Excel versions, so every
    failure mode here — missing file, short read, bad decode, a build box where
    the sidecar is unreadable — degrades to None and the caller falls back to
    the unattributed message. Attribution is a nicety; never let it raise.
    """
    sidecar = workbook_path.with_name(f"~${workbook_path.name}")
    try:
        raw = sidecar.read_bytes()
    except OSError:
        return None
    try:
        length = raw[0]
        name = raw[1 : 1 + length * 2].decode("utf-16-le").strip()
    except (IndexError, UnicodeDecodeError):
        return None
    return name or None


def workbook_lock_holder(workbook_path: Path) -> str | None:
    """Return a description of what holds the workbook, or None if it is free.

    The probe opens the file for writing (``r+b``) and closes it immediately —
    it reads nothing and writes nothing, so it cannot disturb the workbook.
    Excel holds an open .xlsx with write sharing denied, which is exactly what
    makes this the cheap pre-flight signal: a locked workbook fails here in
    milliseconds, where the build otherwise discovers it only when
    ``workbook.save()`` fails at the end of a multi-minute write.

    A missing file is not a lock — the build creates it.
    """
    if not workbook_path.exists():
        return None
    try:
        with workbook_path.open("r+b"):
            pass
    except OSError:
        owner = _lock_file_owner(workbook_path)
        if owner is not None:
            return f"{LOCK_HINT} (owner: {owner})"
        if workbook_path.with_name(f"~${workbook_path.name}").exists():
            return LOCK_HINT
        return "locked by another process"
    return None


def warn_if_workbook_open(
    workbook_path: Path,
    *,
    action_label: str,
    probe: Callable[[Path], str | None] = workbook_lock_holder,
) -> None:
    """Warn before a build touches a workbook something else already holds.

    ``_retry_on_open`` catches the same condition, but only reactively: Excel
    opens a locked workbook read-only without raising, so the first call that
    actually fails is the ``save()`` at the *end* of the write, and the user
    pays the full build before being told to close Excel. This runs the cheap
    probe up front and, on a TTY, offers the same close-and-press-Enter prompt
    ``_retry_on_open`` uses, so the two surfaces read identically.

    Off a TTY this warns and returns rather than raising. The probe is a
    heuristic — a network share, an antivirus scanner or a transient handle can
    all deny the write open on a workbook no human has open — and a false
    positive must not abort an agentic ``poe build-verify``. The reactive path
    still catches a genuine lock at save time exactly as it does today, so
    non-interactive exit codes are unchanged.

    ``probe`` is injectable so the unit suite can exercise the prompt loop
    without a real locked file.
    """
    while True:
        holder = probe(workbook_path)
        if holder is None:
            return
        print(
            f"\nWarning: {workbook_path.name} is {holder}. "
            "The build writes every sheet before Excel refuses the save, so "
            "closing it now saves the whole write.",
            file=sys.stderr,
            flush=True,
        )
        if not sys.stdin.isatty():
            return
        instruction = "close it in Excel" if LOCK_HINT in holder else "release the lock"
        input(
            f"\n{action_label} — {instruction} and press Enter to retry (or Ctrl+C to cancel): "
        )
