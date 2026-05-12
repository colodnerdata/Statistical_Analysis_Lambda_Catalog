"""Shared xlwings helpers and error types for workbook sheet writers."""
from __future__ import annotations

from pathlib import Path

import pywintypes  # type: ignore[import-untyped]
import xlwings as xw


XL_SRC_RANGE = 1
XL_YES = 1
OPEN_WORKBOOK_ERRORS: tuple[type[BaseException], ...] = tuple(
    dict.fromkeys((getattr(pywintypes, "com_error", OSError), OSError))
)


def excel_error_message(exc: BaseException) -> str:
    """Extract a human-readable message from an Excel COM or OS exception.

    Parameters
    ----------
    exc : BaseException
        The exception raised by xlwings or pywintypes.

    Returns
    -------
    str
        The stripped string representation, or the class name when the
        string representation is empty.
    """
    return str(exc).strip() or exc.__class__.__name__


def raise_excel_access_error(
    workbook_path: Path, action: str, exc: BaseException
) -> None:
    """Raise a RuntimeError describing why Excel could not access a workbook.

    Parameters
    ----------
    workbook_path : Path
        Path to the workbook that could not be accessed.
    action : str
        Short description of the attempted operation (e.g. ``'open or save'``).
    exc : BaseException
        The underlying exception from xlwings or the OS.

    Raises
    ------
    RuntimeError
        Always raised. The message includes a hint to close Excel when the
        error appears to be a file-lock or permission issue.
    """
    message = excel_error_message(exc)
    normalized = message.lower()
    likely_locked = any(
        phrase in normalized
        for phrase in (
            "cannot access read-only document",
            "read-only",
            "currently in use",
            "sharing violation",
            "permission denied",
        )
    )

    if likely_locked:
        raise RuntimeError(
            f"Excel could not {action} {workbook_path.name!r}. "
            "The workbook is likely open in Excel or locked by another process. "
            f"Close Excel and retry. Original error: {message}"
        ) from exc

    raise RuntimeError(
        f"Excel could not {action} {workbook_path.name!r}: {message}"
    ) from exc


def open_or_create_workbook(
    app: xw.App, workbook_path: Path
) -> tuple[xw.Book, bool]:
    """Open an existing workbook or create a new empty one.

    Parameters
    ----------
    app : xw.App
        The running xlwings Application instance.
    workbook_path : Path
        Path to the workbook file.

    Returns
    -------
    tuple[xw.Book, bool]
        A 2-tuple of (workbook, workbook_existed) where workbook_existed
        is True when the file was found on disk and opened.
    """
    if workbook_path.exists():
        return app.books.open(str(workbook_path)), True
    return app.books.add(), False


def get_or_create_sheet(workbook: xw.Book, sheet_name: str) -> xw.Sheet:
    """Return the named sheet, creating it at the end of the workbook if absent.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook.
    sheet_name : str
        Name of the sheet to retrieve or create.

    Returns
    -------
    xw.Sheet
        The existing or newly created sheet.
    """
    for sheet in workbook.sheets:
        if sheet.name == sheet_name:
            return sheet
    return workbook.sheets.add(name=sheet_name, after=workbook.sheets[-1])


def reset_generated_sheet(sheet: xw.Sheet) -> None:
    """Clear all ListObjects and cell content from a generated sheet.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet to reset.
    """
    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()


def reset_column_groups(sheet: xw.Sheet) -> None:
    """Remove existing column outlines and ensure all columns are visible.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet whose column grouping state should be reset.
    """
    sheet.api.Cells.ClearOutline()
    sheet.api.Cells.EntireColumn.Hidden = False


def group_and_hide_columns(sheet: xw.Sheet, start_col: int, end_col: int) -> None:
    """Group a contiguous column range and collapse it by hiding those columns.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet where grouping should be applied.
    start_col : int
        1-based start column index of the group.
    end_col : int
        1-based end column index of the group.
    """
    if start_col > end_col:
        return
    grouped_columns = sheet.range((1, start_col), (1, end_col)).api.EntireColumn
    grouped_columns.Group()
    grouped_columns.Hidden = True
