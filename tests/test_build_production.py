"""Tests for the production build driver that do not require Excel."""
# pylint: disable=invalid-name,missing-function-docstring,protected-access,too-few-public-methods
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import cast

import build_production
from lambda_catalog.workbook_builder import (
    NameSyncResult,
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
    XL_CALCULATION_SEMIAUTOMATIC,
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


def test_recalculate_uses_full_rebuild_without_automatic_mode(monkeypatch) -> None:
    app = _FakeApp()
    monkeypatch.setattr(build_production.xw, "App", lambda **_: app)
    workbook_path = Path("Example.xlsx")

    build_production._recalculate_and_save(workbook_path)

    assert app.books.opened_paths == [str(workbook_path)]
    assert app.api.full_rebuilds == 1
    assert app.api.calculation_values == [
        XL_CALCULATION_MANUAL,
        XL_CALCULATION_SEMIAUTOMATIC,
    ]
    assert app.events[:3] == [
        ("open", str(workbook_path)),
        ("calculation", XL_CALCULATION_MANUAL),
        ("full_rebuild", None),
    ]
    assert XL_CALCULATION_AUTOMATIC not in app.api.calculation_values
    assert app.book.saved_paths == [str(workbook_path)]
    assert app.book.closed is True
    assert app.quit_called is True


def test_retry_on_open_retries_dropped_rpc_session(capsys) -> None:
    calls = 0

    def flaky() -> None:
        nonlocal calls
        calls += 1
        if calls == 1:
            raise RuntimeError("Excel could not recalculate: The RPC server is unavailable.")

    build_production._retry_on_open("Example.xlsx is open", flaky, retry_rpc=True)

    assert calls == 2
    assert "retrying in a fresh Excel instance" in capsys.readouterr().err


def test_main_skips_data_table_recalculation_when_requested(
    monkeypatch,
    capsys,
) -> None:
    workbook_path = Path("Example.xlsx")
    calls: list[tuple[str, object]] = []

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=workbook_path,
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=True,
            skip_data_table_calculations=True,
        ),
    )

    def fake_build_production_workbook(**kwargs) -> NameSyncResult:
        calls.append(("build", kwargs))
        return NameSyncResult(created=1, updated=2)

    def fail_recalculate(workbook_path: Path) -> None:
        raise AssertionError("_recalculate_and_save should be skipped")

    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        fake_build_production_workbook,
    )
    monkeypatch.setattr(build_production, "_recalculate_and_save", fail_recalculate)
    monkeypatch.setattr(
        build_production.subprocess,
        "Popen",
        lambda args: calls.append(("popen", args)),
    )

    build_production.main()

    build_call = cast(dict[str, object], next(value for name, value in calls if name == "build"))
    assert build_call["recalculate"] is False
    assert "  Recalculate:    skipped" in capsys.readouterr().out
    assert not any(name == "recalculate" for name, _ in calls)
