"""Tests for write_sheet_version_history's per-artifact version lineage.

The v3.0 split gives the two workbooks independent version lines (shared
function-library version + per-workbook version). The writer selects the
lineage via an ``artifact`` kwarg; these tests pin that selection and the
bad-artifact guard without needing Excel.
"""
# pylint: disable=missing-function-docstring,protected-access
from __future__ import annotations

from types import SimpleNamespace

import pytest

from lambda_catalog import write_sheet_version_history as vh


class _VHRange:
    def __init__(self) -> None:
        self.value = None
        self.color = None
        self.api = SimpleNamespace(
            Font=SimpleNamespace(Bold=None),
            EntireRow=SimpleNamespace(AutoFit=lambda: None),
        )


class _VHSheet:
    def __init__(self, name: str = "Version History") -> None:
        self.name = name
        self._ranges: dict[tuple, _VHRange] = {}

    def range(self, *addresses) -> _VHRange:
        key = tuple(addresses)
        if key not in self._ranges:
            self._ranges[key] = _VHRange()
        return self._ranges[key]


@pytest.fixture
def stub_helpers(monkeypatch):
    """Stub the Excel-touching helpers so the writer runs headless against a
    single in-memory sheet; return the sheet so a test can read written values."""
    sheet = _VHSheet()
    monkeypatch.setattr(vh, "get_or_create_sheet", lambda workbook, name: sheet)
    monkeypatch.setattr(vh, "reset_generated_sheet", lambda _sheet: None)
    monkeypatch.setattr(vh, "set_column_widths", lambda _sheet, _columns: None)
    return sheet


def _version_rows(sheet: _VHSheet) -> list[str]:
    """Collect the version strings written in column 1, rows 3 onward."""
    rows: list[str] = []
    row = 3
    # sheet.range((row, 1)) is keyed as a 1-tuple wrapping the (row, 1) pair.
    while ((row, 1),) in sheet._ranges:
        rows.append(sheet._ranges[((row, 1),)].value)
        row += 1
    return rows


def test_univariate_artifact_writes_univariate_lineage(stub_helpers) -> None:
    """artifact='univariate' writes the Univariate lineage, which begins at the
    v3.0 split — NOT the Regression lineage."""
    vh.write_version_history_sheet(object(), artifact="univariate")

    assert _version_rows(stub_helpers) == ["1.0.0", "2.0.0"]


def test_regression_artifact_writes_regression_lineage(stub_helpers) -> None:
    """artifact='regression' (the default) writes the Regression lineage, in
    ascending order from 1.0.0 with no gaps."""
    vh.write_version_history_sheet(object())  # default artifact

    rows = _version_rows(stub_helpers)
    assert rows[0] == "1.0.0"
    assert rows[-1] == "3.3.0"
    # The Regression lineage carries the pre-split Univariate history
    # (1.1.0's "Univariate Analysis release"); the standalone Univariate
    # artifact's lineage does not — its history begins at the split.
    assert "1.1.0" in rows
    # Every shipped minor has an entry. This sheet IS the changelog for
    # non-git users, so a release reaching them with no row is the defect —
    # 2.1.0 (Fixed Effects) and 2.2.0 (Transforms) both did exactly that
    # until they were backfilled.
    assert rows == [
        "1.0.0", "1.1.0", "1.2.0", "2.0.0", "2.1.0", "2.2.0", "3.0.0", "3.1.0",
        "3.3.0",
    ]
    # No 3.2.0: that number was reserved for the resequenced "transforms
    # remainder" milestone, which shipped under 3.3.0 (the unit-space
    # dispatcher, Duan smearing, and the model-formula label land together).
    assert "3.2.0" not in rows


def test_no_shipped_version_is_missing_from_the_regression_changelog(
    stub_helpers,
) -> None:
    # Ascending, strictly increasing, no duplicates — checked structurally so
    # a future entry inserted in the wrong place fails here rather than
    # shipping a changelog that reads out of order.
    vh.write_version_history_sheet(object())

    parsed = [tuple(int(part) for part in row.split(".")) for row in _version_rows(stub_helpers)]
    assert parsed == sorted(parsed)
    assert len(set(parsed)) == len(parsed)


def test_unknown_artifact_raises_value_error(stub_helpers) -> None:
    with pytest.raises(ValueError, match="artifact must be 'regression' or 'univariate'"):
        vh.write_version_history_sheet(object(), artifact="mileage")