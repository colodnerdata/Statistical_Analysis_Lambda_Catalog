"""Shared xlwings helpers: error handling, cell formatting, sheet operations, and name management."""
from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import NamedTuple, NoReturn

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
import math

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER
from .sheet_styles import INPUT_COLOR as _INPUT

XL_SRC_RANGE = 1
XL_YES = 1
MAX_EXCEL_ROW = 1_048_576
_XL_EXPRESSION = 2
OPEN_WORKBOOK_ERRORS: tuple[type[BaseException], ...] = tuple(
    dict.fromkeys((getattr(pywintypes, "com_error", OSError), OSError))
)

# The sentinel raise_excel_access_error plants in the message of a RuntimeError
# that a locked workbook caused. Every retry loop in the repo recognises a lock
# by looking for this substring, so it is a shared constant rather than a phrase
# repeated at each site: rewording the message without updating a hand-copied
# literal would silently stop every retry loop from retrying. Imported by
# build_common (_retry_on_open, workbook_lock_holder).
LOCK_HINT = "likely open in Excel"


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
            f"The workbook is {LOCK_HINT} or locked by another process. "
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


def safe_activate(sheet: xw.Sheet) -> None:
    """Activate a worksheet, tolerating environments where Excel has no foreground window.

    ``Sheet.activate()`` raises when Excel cannot become the active
    application (no interactive desktop session, focus stolen by another
    process, etc.), even though the workbook write itself succeeds. Which
    sheet is on top when the file opens is cosmetic, so that failure must
    not abort the build.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet to bring to the front.
    """
    try:
        sheet.activate()
    except Exception:
        pass


def safe_freeze_top_row(sheet: xw.Sheet) -> None:
    """Freeze the header row of a worksheet, tolerating a missing active window.

    Freezing panes reads and writes ``Application.ActiveWindow``, which is
    only populated when Excel has an active window for the sheet. In the
    same headless/no-focus sessions ``safe_activate`` guards against, that
    property access can itself raise, so this is best-effort too.

    Parameters
    ----------
    sheet : xw.Sheet
        The worksheet whose top row should be frozen.
    """
    try:
        window = sheet.api.Application.ActiveWindow
        window.SplitRow = 1
        window.SplitColumn = 0
        window.FreezePanes = True
    except Exception:
        pass


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


def copy_static_sheet(workbook: xw.Book, template_path: Path, sheet_name: str) -> xw.Sheet:
    """Replace ``sheet_name`` in ``workbook`` with a copy from a static template.

    For reference/documentation sheets whose content never depends on the
    target dataset (e.g. Regression Instructions, Diagnostic Guide), copying
    an already-styled sheet via Excel's own ``Copy`` is far cheaper than
    reconstructing every cell with COM calls on each build, and keeps the
    content itself in one editable place — the template workbook — rather
    than a Python row list. ``workbook`` must be open in a running Excel
    instance; the template is opened in that same instance (or reused if
    already open there) for the cross-workbook copy to work.

    Parameters
    ----------
    workbook : xw.Book
        The open target workbook.
    template_path : Path
        Path to the ``.xlsx`` template holding the source sheet.
    sheet_name : str
        Name of the sheet to copy from the template. Any existing sheet of
        the same name in ``workbook`` is deleted first.

    Returns
    -------
    xw.Sheet
        The freshly copied sheet, now part of ``workbook``.
    """
    for sheet in list(workbook.sheets):
        if sheet.name == sheet_name:
            sheet.delete()
            break

    app = workbook.app
    resolved_template_path = template_path.resolve()

    # Reuse the template if it's already open in this Excel instance (e.g. a
    # developer editing it by hand) so this helper only closes what it opened
    # — closing someone else's open workbook out from under them would be a
    # surprising side effect of what looks like a read-only copy operation.
    template_book = None
    for book in app.books:
        try:
            if Path(book.fullname).resolve() == resolved_template_path:
                template_book = book
                break
        except OPEN_WORKBOOK_ERRORS:
            continue

    opened_here = template_book is None
    if opened_here:
        try:
            # read_only because this helper only ever copies a sheet OUT of the
            # template — and because it is the one file two concurrent builds
            # could contend over. `poe verify` runs its three Excel drivers in
            # parallel; a read-write open would give the second one Excel's
            # file-in-use path (a modal prompt with no window to show it in, or
            # a silent read-only downgrade) instead of a clean shared read.
            template_book = app.books.open(str(resolved_template_path), read_only=True)
        except OPEN_WORKBOOK_ERRORS as exc:
            # Attribute the failure to the template, not `workbook` — a
            # missing/locked/corrupt template would otherwise surface as a
            # misleading error about the (perfectly fine) target workbook.
            raise_excel_access_error(resolved_template_path, "open", exc)

    try:
        template_book.sheets[sheet_name].api.Copy(After=workbook.sheets[-1].api)
    finally:
        if opened_here:
            template_book.close()

    return workbook.sheets[sheet_name]


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


def quoted_sheet_name(name: str) -> str:
    """Quote a sheet name for use in a formula or a ``RefersTo``.

    Two things have to happen and only the first was ever done consistently.

    **Wrap in single quotes.** A sheet name containing a space — which every
    generated test-model sheet has ("M01 Baseline Categoricals") — makes an
    unquoted reference an invalid formula, and Excel rejects the whole
    ``Names.Add`` with "There's a problem with this formula". Quoting a name
    that does not need it is always legal, which is why this is unconditional.

    **Double any embedded apostrophe.** ``'`` is the quote character, so a
    literal one inside the name must be written ``''`` — ``'M17 Cook''s D'!$A$1``.
    ``test_model_sheets.validate_sheet_name`` rejects a LEADING or TRAILING
    apostrophe (Excel's own rule) but permits an internal one, so a case named
    "M17 Cook's D Threshold" is legal, plausible in this repo, and would
    otherwise build a reference Excel cannot parse. No sheet name carries one
    today; this closes the gap before one does rather than after.

    Parameters
    ----------
    name : str
        Bare sheet name.

    Returns
    -------
    str
        The name wrapped in single quotes, apostrophes escaped — ready to be
        followed by ``!`` and a cell reference.
    """
    return "'" + name.replace("'", "''") + "'"


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

    index: int  # type: ignore[assignment]  # Intentionally shadows tuple.index.
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
    cell = sheet.range(rc(row, col))
    api = cell.api
    if api is None:
        cell.formula = formula
        return
    api.Formula2 = formula


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

# ── Comment/Note shape sizing ─────────────────────────────────────────────────
# Excel's default comment box is small (~150x60 pt) and clips multi-sentence
# notes. Both writers use anchored, sized comment shapes instead. The width
# scales with text length; height grows to fit the wrapped line count. A
# per-`label` SIZE_OVERRIDES entry (regression-local, in write_sheet_regression)
# replaces either axis for individual notes the heuristic guesses wrong for.

_NOTE_MIN_WIDTH = 150.0     # points
_NOTE_MAX_WIDTH = 320.0     # points
_NOTE_BASE_WIDTH = 200.0    # width used for a ~80-char note before scaling
_NOTE_CHARS_PER_LINE_PER_POINT = 1.0 / 5.2  # ~5.2pt per character at 8pt Tahoma
_NOTE_LINE_HEIGHT = 12.0    # points per wrapped line
_NOTE_MIN_HEIGHT = 32.0     # points
_NOTE_VERTICAL_PADDING = 10.0  # points added above/below the wrapped text


def note_dimensions(
    label: str,
    text: str,
    size_overrides: dict[str, tuple[float | None, float | None]] | None = None,
) -> tuple[float, float]:
    """Guess a (width, height) in points that fits `text` without clipping.

    Width grows with text length (clamped to a readable range); height is
    then derived from how many lines that width wraps the text into. A
    per-`label` entry in `size_overrides` replaces either axis (or both)
    for hand-tuning notes the heuristic guesses wrong for. Pass `None` to
    use purely the heuristic (no overrides).
    """
    length = len(text)
    width = min(
        _NOTE_MAX_WIDTH,
        max(_NOTE_MIN_WIDTH, _NOTE_BASE_WIDTH + (length - 80) * 0.35),
    )
    chars_per_line = max(1, int(width * _NOTE_CHARS_PER_LINE_PER_POINT))
    lines = sum(
        max(1, math.ceil(len(paragraph) / chars_per_line))
        for paragraph in text.split("\n")
    )
    height = max(_NOTE_MIN_HEIGHT, lines * _NOTE_LINE_HEIGHT + _NOTE_VERTICAL_PADDING)

    override_width, override_height = (None, None)
    if size_overrides is not None:
        override_width, override_height = size_overrides.get(label, (None, None))
    return (
        override_width if override_width is not None else width,
        override_height if override_height is not None else height,
    )


def anchor_comment_right_of_cell(
    sheet: xw.Sheet, row: int, col: int, width: float, height: float,
) -> None:
    """Position the cell's comment box immediately to the cell's right.

    Best-effort: wrapped in a try/except so the build succeeds even when
    the COM API for Shape.Left/Top/Width/Height is unavailable (CI/headless).
    """
    cell = sheet.range(rc(row, col))
    try:
        comment_shape = cell.api.Comment.Shape
        comment_shape.Width = width
        comment_shape.Height = height
        comment_shape.Left = cell.left + cell.width
        comment_shape.Top = cell.top
    except Exception:  # pylint: disable=broad-exception-caught
        pass


def drop_local_name(sheet: xw.Sheet, name: str) -> None:
    """Delete every sheet-scoped name whose local part matches ``name``.

    Iterates the collection in REVERSE so a delete cannot shift the indices of
    entries not yet visited.

    Each entry is handled defensively because a worksheet ``Names`` collection
    is not guaranteed to hand back a live object at every index: Excel can
    return ``None`` for a broken or orphaned entry (one whose ``RefersTo``
    target no longer exists), and reading ``.Name`` on it raised
    ``AttributeError: 'NoneType' object has no attribute 'Name'``, which
    aborted the whole production build partway through the Regression sheet.
    A name we cannot even read is, by definition, not the one being dropped,
    so skipping it is both safe and the only thing that can be done with it.
    """
    for idx in range(sheet.api.Names.Count, 0, -1):
        try:
            entry = sheet.api.Names(idx)
            if entry is None:
                continue
            local = entry.Name.split("!", 1)[-1]
        except Exception:  # pylint: disable=broad-exception-caught
            continue
        if local.lower() == name.lower():
            entry.Delete()
