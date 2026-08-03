"""Build scaffolding shared by the two production build scripts.

From v3.0 the project ships two artifacts — `build_production.py` (the
Regression workbook) and `build_univariate.py` (the standalone Univariate
workbook). Both follow the same three-phase structure (write sheets + sync
names; recalculate and save; optional deep verify), and both need the same
Excel-open retry handling and the same "open headless Excel, run
CalculateFullRebuild, save under a chosen calculation mode" recalc step. Those
helpers live here so the two driver scripts stay thin and do not duplicate the
retry/recalc logic.
"""
from __future__ import annotations

import sys
import time
from collections.abc import Callable
from pathlib import Path

import xlwings as xw

from lambda_catalog.workbook_builder import (
    NameSyncResult,
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
)
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error

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


def print_name_sync_summary(result: NameSyncResult) -> None:
    """Print the workbook-scope sync counts every build script reports.

    ``Removed names`` and ``Skipped names`` are the interesting two: the first
    counts workbook-scoped residue the sync cleared out (a name the catalog
    retired, or one belonging to the other artifact), the second names the
    catalog entries this artifact cannot carry because they reference a
    worksheet it does not have. Both are normal on the first build after a
    change and should settle to 0 / none on a rebuild — except the Univariate
    artifact's permanent ``Base_Period_Delta`` skip.
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
    tree and evaluates every formula — including the Univariate Data Tables
    when present), sets the workbook's persisted calculation mode to
    ``calc_mode`` (full Automatic for both artifacts post-split), and saves.

    Parameters
    ----------
    workbook_path : Path
        Workbook to open, recalculate, and save in place.
    calc_mode : int
        Excel ``XlCalculation`` constant to persist as the workbook's
        calculation mode. Both production artifacts use full Automatic
        (``XL_CALCULATION_AUTOMATIC``); the Univariate Data Tables recalculate
        on edit under that mode, which is the point of the split.
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
            if "likely open in Excel" not in str(exc):
                raise
            if not sys.stdin.isatty():
                raise
            input(
                f"\n{action_label} — close it in Excel "
                "and press Enter to retry (or Ctrl+C to cancel): "
            )