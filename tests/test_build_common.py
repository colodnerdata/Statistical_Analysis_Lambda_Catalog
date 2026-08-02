"""Tests for the shared build scaffolding (lambda_catalog.build_common).

These helpers were extracted from build_production.py when the project split
into two build scripts (build_production.py for the Regression workbook,
build_univariate.py for the standalone Univariate workbook). They live in
lambda_catalog.build_common so both driver scripts stay thin and share the
Excel-open retry handling and the "open headless Excel, run
CalculateFullRebuild, save under a chosen calculation mode" recalc step
without duplicating it.
"""
# pylint: disable=invalid-name,missing-function-docstring,protected-access
from __future__ import annotations

from pathlib import Path

import lambda_catalog.build_common as build_common
from lambda_catalog.workbook_builder import (
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
)


class _FakeApi:
    def __init__(self, events: list[tuple[str, object]]) -> None:
        self.events = events
        self.DisplayAlerts = True
        self.AskToUpdateLinks = True
        self.calculation_values: list[int] = []
        self.full_rebuilds = 0

    @property
    def Calculation(self) -> int | None:
        return self.calculation_values[-1] if self.calculation_values else None

    @Calculation.setter
    def Calculation(self, value: int) -> None:
        self.events.append(("calculation", value))
        self.calculation_values.append(value)

    def CalculateFullRebuild(self) -> None:
        self.events.append(("full_rebuild", None))
        self.full_rebuilds += 1


class _FakeBook:
    def __init__(self) -> None:
        self.saved_paths: list[str] = []
        self.closed = False

    def save(self, path: str) -> None:
        self.saved_paths.append(path)

    def close(self) -> None:
        self.closed = True


class _FakeBooks:
    def __init__(self, book: _FakeBook, events: list[tuple[str, object]]) -> None:
        self.book = book
        self.events = events
        self.opened_paths: list[str] = []

    def open(self, path: str) -> _FakeBook:
        self.events.append(("open", path))
        self.opened_paths.append(path)
        return self.book


class _FakeApp:
    def __init__(self) -> None:
        self.events: list[tuple[str, object]] = []
        self.api = _FakeApi(self.events)
        self.book = _FakeBook()
        self.books = _FakeBooks(self.book, self.events)
        self.quit_called = False

    def quit(self) -> None:
        self.quit_called = True


def test_recalculate_uses_full_rebuild_then_persists_automatic_mode(monkeypatch) -> None:
    """_recalculate_and_save opens headless Excel in Manual (so the open does
    not trigger a recalc), runs CalculateFullRebuild, then persists the chosen
    calc mode — full Automatic by default for both post-split artifacts."""
    app = _FakeApp()
    monkeypatch.setattr(build_common.xw, "App", lambda **_: app)
    workbook_path = Path("Example.xlsx")

    build_common._recalculate_and_save(workbook_path)

    assert app.books.opened_paths == [str(workbook_path)]
    assert app.api.full_rebuilds == 1
    assert app.api.calculation_values == [
        XL_CALCULATION_MANUAL,
        XL_CALCULATION_AUTOMATIC,
    ]
    assert app.events[:3] == [
        ("open", str(workbook_path)),
        ("calculation", XL_CALCULATION_MANUAL),
        ("full_rebuild", None),
    ]
    assert app.book.saved_paths == [str(workbook_path)]
    assert app.book.closed is True
    assert app.quit_called is True


def test_recalculate_persists_explicit_calc_mode(monkeypatch) -> None:
    """The calc_mode parameter controls the mode persisted to the saved
    workbook (both production artifacts pass Automatic, but the helper must
    honour whatever it is given)."""
    app = _FakeApp()
    monkeypatch.setattr(build_common.xw, "App", lambda **_: app)

    build_common._recalculate_and_save(Path("Example.xlsx"), calc_mode=XL_CALCULATION_MANUAL)

    assert app.api.calculation_values == [
        XL_CALCULATION_MANUAL,  # set before the rebuild to suppress open-recalc
        XL_CALCULATION_MANUAL,  # persisted after the rebuild
    ]


def test_retry_on_open_retries_dropped_rpc_session(capsys) -> None:
    """A dropped COM session gets one automatic retry in a fresh Excel
    instance before the error propagates."""
    calls = 0

    def flaky() -> None:
        nonlocal calls
        calls += 1
        if calls == 1:
            raise RuntimeError("Excel could not recalculate: The RPC server is unavailable.")

    build_common._retry_on_open("Example.xlsx is open", flaky, retry_rpc=True)

    assert calls == 2
    assert "retrying in a fresh Excel instance" in capsys.readouterr().err