"""Tests for the best-effort Excel-window helpers in workbook_helpers."""
# pylint: disable=missing-function-docstring
from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast

from lambda_catalog.workbook_helpers import (
    copy_static_sheet,
    f,
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


def _make_sheet(api_value):
    """Return a minimal sheet stub whose range() returns a cell with given api."""

    class _Cell:
        def __init__(self):
            self.api = api_value
            self.formula = None

    class _Sheet:
        def range(self, _addr):
            return _Cell()

    return _Sheet()


def test_f_uses_formula2_when_api_available() -> None:
    recorded = {}

    class _Api:
        def __setattr__(self, name, value):
            if name != "__class__":
                recorded[name] = value
            super().__setattr__(name, value)

    sheet = _make_sheet(_Api())
    f(sheet, 1, 1, "=SUM(A1:A10)")
    assert recorded.get("Formula2") == "=SUM(A1:A10)"


def test_f_falls_back_to_formula_property_when_api_is_none() -> None:
    class _Cell:
        def __init__(self):
            self.api = None
            self.formula = None

    cell = _Cell()

    class _Sheet:
        def range(self, _addr):
            return cell

    f(_Sheet(), 1, 1, "=SUM(A1:A10)")
    assert cell.formula == "=SUM(A1:A10)"



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


def test_quoted_sheet_name_escapes_embedded_apostrophes() -> None:
    """Two rules, and only the first was ever applied consistently.

    Wrapping in single quotes is what makes a sheet name with a SPACE usable —
    every generated test-model sheet has one, and an unquoted reference makes
    Excel reject the whole Names.Add. That much the writers already did inline.

    Doubling an embedded apostrophe is the half that was missing. `'` is the
    quote character, so a literal one has to be written `''`.
    `test_model_sheets.validate_sheet_name` rejects a LEADING or TRAILING
    apostrophe — Excel's own rule — but permits an internal one, so a case
    named "M17 Cook's D Threshold" is legal and, in a repo that already ships
    Cook's Distance diagnostics, entirely plausible. Without escaping it builds
    a reference Excel cannot parse, and the failure lands at build time on a
    machine with Excel, which is the slowest place to find it.
    """
    from lambda_catalog.workbook_helpers import quoted_sheet_name

    # The ordinary cases: quoted unconditionally, because quoting a name that
    # does not need it is always legal and removes the per-call-site judgement.
    assert quoted_sheet_name("Regression") == "'Regression'"
    assert quoted_sheet_name("M01 Baseline Categoricals") == "'M01 Baseline Categoricals'"

    # The case this function exists for.
    assert quoted_sheet_name("M17 Cook's D Threshold") == "'M17 Cook''s D Threshold'"
    assert quoted_sheet_name("a'b'c") == "'a''b''c'"


def test_a_sheet_name_with_an_apostrophe_still_builds_valid_defined_names() -> None:
    """End to end: the escaping reaches the RefersTo strings, not just the helper."""
    import sys

    sys.path.insert(0, ".")
    from tests.recording_sheet import RecordingSheet
    from lambda_catalog.regression_materialization import _write_materialization_zone

    sheet = RecordingSheet(name="M17 Cook's D Threshold")
    _write_materialization_zone(cast(Any, sheet), closures=())

    refers_to = {
        item.Name.split("!", 1)[-1]: item.RefersTo for item in sheet.api.Names.items
    }
    for name in ("Fit_Context", "Fit_Sample_Include", "Fit_Design_Columns"):
        assert "'M17 Cook''s D Threshold'!" in refers_to[name], name
        # The unescaped form would terminate the quoted name early at "Cook'".
        assert "'M17 Cook's" not in refers_to[name], name


# --- drop_local_name --------------------------------------------------------
#
# A worksheet Names collection can hand back None at a valid index when the
# entry is broken or orphaned. Reading .Name on it raised
# "AttributeError: 'NoneType' object has no attribute 'Name'" partway through
# the Regression sheet and aborted the whole production build, leaving the
# shipped dist/ workbook stale. These pin the guard.


class _Name:
    def __init__(self, name: str) -> None:
        self.Name = name
        self.deleted = False

    def Delete(self) -> None:
        self.deleted = True


class _Names:
    """1-indexed COM-style Names collection; a None entry models a broken name."""

    def __init__(self, entries: list[object]) -> None:
        self._entries = entries

    @property
    def Count(self) -> int:
        return len(self._entries)

    def __call__(self, idx: int) -> object:
        return self._entries[idx - 1]


def _sheet_with_names(entries: list[object]) -> object:
    return SimpleNamespace(api=SimpleNamespace(Names=_Names(entries)))


def test_drop_local_name_deletes_the_matching_sheet_scoped_name() -> None:
    from lambda_catalog.workbook_helpers import drop_local_name

    target = _Name("Regression!Fit_Context")
    other = _Name("Regression!Fit_Sample_Include")
    drop_local_name(cast(Any, _sheet_with_names([other, target])), "Fit_Context")

    assert target.deleted is True
    assert other.deleted is False


def test_drop_local_name_matches_case_insensitively_and_ignores_the_sheet_prefix() -> None:
    from lambda_catalog.workbook_helpers import drop_local_name

    target = _Name("Some Sheet!FIT_CONTEXT")
    drop_local_name(cast(Any, _sheet_with_names([target])), "fit_context")

    assert target.deleted is True


def test_drop_local_name_skips_a_none_entry_and_still_deletes_the_match() -> None:
    """The regression: a None entry must not abort the sweep."""
    from lambda_catalog.workbook_helpers import drop_local_name

    target = _Name("Regression!Fit_Context")
    # None sits BELOW the target so the reverse walk reaches the target only
    # if the None is skipped rather than raised on.
    drop_local_name(
        cast(Any, _sheet_with_names([None, target, None])), "Fit_Context"
    )

    assert target.deleted is True


def test_drop_local_name_skips_an_entry_whose_name_raises() -> None:
    from lambda_catalog.workbook_helpers import drop_local_name

    class _Broken:
        @property
        def Name(self) -> str:
            raise OSError("name refers to a deleted range")

    target = _Name("Regression!Fit_Context")
    drop_local_name(
        cast(Any, _sheet_with_names([_Broken(), target])), "Fit_Context"
    )

    assert target.deleted is True
