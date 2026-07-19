"""Shared xlwings helpers: error handling, cell formatting, sheet operations, and name management."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable, NamedTuple, NoReturn

try:
    import pywintypes  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - non-Windows environments
    class _ComError(OSError):
        """Fallback Excel COM error type for non-Windows environments."""

    class _PyWinTypesFallback:
        """Fallback namespace for non-Windows environments."""

        com_error = _ComError

        def __getattr__(self, name: str) -> NoReturn:
            raise AttributeError(
                f"pywintypes.{name} is unavailable (pywintypes not installed)"
            )

    pywintypes = _PyWinTypesFallback()
import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER, INPUT_COLOR as _INPUT


XL_SRC_RANGE = 1
XL_YES = 1
MAX_EXCEL_ROW = 1_048_576
_XL_EXPRESSION = 2
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
) -> NoReturn:
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
    start_letter = col_letter(start_col)
    end_letter = col_letter(end_col)
    grouped_columns = sheet.range(f"{start_letter}:{end_letter}").api.EntireColumn
    grouped_columns.Group()
    grouped_columns.Hidden = True


def col_letter(col_idx: int) -> str:
    """Convert a 1-based column index to an Excel column letter string."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


# ── Cell address helpers ───────────────────────────────────────────────────────

def rc(row: int, col: int) -> tuple[int, int]:
    return (row, col)


def a1(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


# ── Column layout helpers ──────────────────────────────────────────────────────

class ColumnSpec(NamedTuple):
    """One column's layout facts: index, width, header text, and a note.

    A file-local declarative alternative to separate constant/comment/width
    sections. Feed `(c.index, c.width)` pairs from a tuple of these to
    `set_column_widths` (e.g. ``((c.index, c.width) for c in _COLUMNS)``),
    and derive `_C_*` constants from `.index` via tuple-unpacking so
    existing call sites keep working unchanged.
    """

    index: int
    width: float | None = None
    header: str = ""
    note: str = ""


def set_column_widths(sheet: xw.Sheet, columns: Iterable[tuple[int, float | None]]) -> None:
    """Set column widths from an iterable of (1-based index, width) pairs.

    ``width=None`` entries are skipped, so a `ColumnSpec` list can be passed
    directly (e.g. ``((c.index, c.width) for c in _COLUMNS)``) without
    filtering out columns that intentionally have no width opinion.
    """
    for col, width in columns:
        if width is None:
            continue
        sheet.range(rc(1, col), rc(1, col)).column_width = width


# ── Cell value / formula helpers ───────────────────────────────────────────────

def val(sheet: xw.Sheet, row: int, col: int, value: object) -> None:
    sheet.range(rc(row, col)).value = value


def f(sheet: xw.Sheet, row: int, col: int, formula: str) -> None:
    sheet.range(rc(row, col)).api.Formula2 = formula


def f_structured(sheet: xw.Sheet, row: int, col: int, formula: str) -> None:
    """Write a formula using the ``Formula`` property (not ``Formula2``).

    ``Formula2`` is the locale-aware variant that rejects structured table
    references like ``[@Column]``; use this helper whenever a formula
    refers to a structured table column or row.
    """
    sheet.range(rc(row, col)).api.Formula = formula


# ── Cell formatting helpers ────────────────────────────────────────────────────

def bold(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(rc(row, col)).api.Font.Bold = True


def bold_row(sheet: xw.Sheet, row: int, col1: int, col2: int) -> None:
    sheet.range(rc(row, col1), rc(row, col2)).api.Font.Bold = True


def section_heading(sheet: xw.Sheet, row: int, col: int, label: str) -> None:
    val(sheet, row, col, label)
    sheet.range(rc(row, col)).api.Font.Bold = True
    sheet.range(rc(row, col)).color = _HEADER


def format_input(sheet: xw.Sheet, row: int, col: int) -> None:
    sheet.range(rc(row, col)).color = _INPUT


def border_box(sheet: xw.Sheet, r1: int, c1: int, r2: int, c2: int) -> None:
    rng = sheet.range(rc(r1, c1), rc(r2, c2)).api
    for edge in [7, 8, 9, 10]:   # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
        rng.Borders(edge).LineStyle = 1   # xlContinuous
        rng.Borders(edge).Weight = 2      # xlThin


def excel_color(rgb: tuple[int, int, int]) -> int:
    """Convert an RGB tuple to the OLE color integer expected by Excel COM."""
    red, green, blue = rgb
    return red + green * 256 + blue * 65536


def add_expression_format(
    sheet: xw.Sheet,
    address: str,
    formula: str,
    *,
    fill: tuple[int, int, int] | None = None,
    font_color: tuple[int, int, int] | None = None,
    bold_font: bool | None = None,
    strikethrough: bool | None = None,
    stop_if_true: bool = False,
):
    """Add a formula-based conditional-formatting rule to a range."""
    condition = sheet.range(address).api.FormatConditions.Add(
        Type=_XL_EXPRESSION,
        Formula1=formula,
    )

    if fill is not None:
        condition.Interior.Color = excel_color(fill)

    if font_color is not None:
        condition.Font.Color = excel_color(font_color)

    if bold_font is not None:
        condition.Font.Bold = bold_font
    if strikethrough is not None:
        condition.Font.Strikethrough = strikethrough

    condition.StopIfTrue = stop_if_true
    return condition


# ── Name management ───────────────────────────────────────────────────────────

def drop_local_name(sheet: xw.Sheet, name: str) -> None:
    for idx in range(sheet.api.Names.Count, 0, -1):
        local = sheet.api.Names(idx).Name.split("!", 1)[-1]
        if local.lower() == name.lower():
            sheet.api.Names(idx).Delete()
