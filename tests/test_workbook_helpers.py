"""Tests for the best-effort Excel-window helpers in workbook_helpers."""
# pylint: disable=missing-function-docstring
from pathlib import Path
from types import SimpleNamespace

from lambda_catalog.workbook_helpers import (
    copy_static_sheet,
    safe_activate,
    safe_freeze_top_row,
)


class _RaisingSheet:
    def activate(self) -> None:
        raise RuntimeError("Could not activate App! Try to instantiate the App with visible=True.")


def test_safe_activate_swallows_activation_failure() -> None:
    safe_activate(_RaisingSheet())  # must not raise


def test_safe_activate_calls_through_on_success() -> None:
    calls: list[str] = []
    sheet = SimpleNamespace(activate=lambda: calls.append("activated"))
    safe_activate(sheet)
    assert calls == ["activated"]


def test_safe_freeze_top_row_swallows_missing_active_window() -> None:
    class _NoActiveWindowApplication:
        @property
        def ActiveWindow(self) -> None:
            raise RuntimeError("ActiveWindow is not available in this session")

    sheet = SimpleNamespace(api=SimpleNamespace(Application=_NoActiveWindowApplication()))

    safe_freeze_top_row(sheet)  # must not raise


def test_safe_freeze_top_row_sets_split_and_freeze_on_success() -> None:
    window = SimpleNamespace(SplitRow=None, SplitColumn=None, FreezePanes=None)
    sheet = SimpleNamespace(
        api=SimpleNamespace(Application=SimpleNamespace(ActiveWindow=window))
    )

    safe_freeze_top_row(sheet)

    assert window.SplitRow == 1
    assert window.SplitColumn == 0
    assert window.FreezePanes is True


# --- copy_static_sheet ------------------------------------------------------
#
# The template open is the one file access two concurrent builds could collide
# over, and `poe verify` now runs its three Excel drivers at once. Excel is not
# available here, so these pin the call shape against COM doubles.


class _Sheets:
    """The slice of ``Book.sheets`` this helper touches: iterate, index by
    position, index by name."""

    def __init__(self, names: list[str]) -> None:
        self._sheets = []
        for name in names:
            sheet = SimpleNamespace(
                name=name, api=SimpleNamespace(Copy=lambda **_: None)
            )
            sheet.delete = lambda sheet=sheet: self._sheets.remove(sheet)
            self._sheets.append(sheet)
        # The helper deletes the same-named sheet before copying, then returns
        # workbook.sheets[name] — which in Excel is the freshly copied one.
        self._restore_on_lookup = names

    def __iter__(self):
        return iter(self._sheets)

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._sheets[key]
        for sheet in self._sheets:
            if sheet.name == key:
                return sheet
        # Deleted, then re-created by the Copy the helper just performed.
        if key in self._restore_on_lookup:
            return SimpleNamespace(name=key, api=SimpleNamespace())
        raise KeyError(key)


class _Books:
    def __init__(self, already_open: list[object]) -> None:
        self._open = already_open
        self.opens: list[dict] = []

    def __iter__(self):
        return iter(self._open)

    def open(self, path: str, **kwargs):
        self.opens.append({"path": path, **kwargs})
        book = SimpleNamespace(
            fullname=path, sheets=_Sheets(["Diagnostic Guide"]), closed=False
        )
        book.close = lambda: setattr(book, "closed", True)
        return book


def _target_workbook(books: _Books) -> SimpleNamespace:
    return SimpleNamespace(
        app=SimpleNamespace(books=books),
        sheets=_Sheets(["Regression", "Diagnostic Guide"]),
    )


def test_copy_static_sheet_opens_the_template_read_only(tmp_path: Path) -> None:
    """A read-write open is what would give a concurrent build Excel's
    file-in-use path — a modal prompt with no window to show it in, or a silent
    read-only downgrade — instead of a clean shared read."""
    template = tmp_path / "static_sheets.xlsx"
    template.write_bytes(b"")
    books = _Books(already_open=[])

    copy_static_sheet(_target_workbook(books), template, "Diagnostic Guide")

    assert len(books.opens) == 1
    assert books.opens[0]["read_only"] is True


def test_copy_static_sheet_reuses_and_leaves_open_a_template_it_did_not_open(
    tmp_path: Path,
) -> None:
    """Closing a workbook somebody else opened would be a surprising side
    effect of what looks like a read-only copy."""
    template = tmp_path / "static_sheets.xlsx"
    template.write_bytes(b"")
    already = SimpleNamespace(
        fullname=str(template), sheets=_Sheets(["Diagnostic Guide"]), closed=False
    )
    already.close = lambda: setattr(already, "closed", True)
    books = _Books(already_open=[already])

    copy_static_sheet(_target_workbook(books), template, "Diagnostic Guide")

    assert books.opens == []
    assert already.closed is False
