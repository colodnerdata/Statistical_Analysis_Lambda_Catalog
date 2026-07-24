"""Tests for the production build driver that do not require Excel."""
# pylint: disable=invalid-name,missing-function-docstring,protected-access,too-few-public-methods
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import cast

import pytest

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


class _FakeSheet:
    def __init__(self, name: str) -> None:
        self.name = name
        self.deleted = False

    def delete(self) -> None:
        self.deleted = True


class _FakeSheetCollection:
    def __init__(self) -> None:
        self.items = [_FakeSheet("Sheet1")]

    def __iter__(self):
        return iter(self.items)

    def __getitem__(self, key):
        if isinstance(key, str):
            return next(sheet for sheet in self.items if sheet.name == key)
        return self.items[key]


class _CleanupFailingBook:
    def __init__(self) -> None:
        self.sheets = _FakeSheetCollection()

    def save(self, path: str) -> None:
        raise AssertionError("save should not be reached")

    def close(self) -> None:
        raise OSError("cleanup close masked the original error")


class _BuildBooks:
    def __init__(self, book: _CleanupFailingBook) -> None:
        self.book = book

    def add(self) -> _CleanupFailingBook:
        return self.book

    def open(self, path: str) -> _CleanupFailingBook:
        raise AssertionError("new workbook path should not be opened")


class _CleanupFailingApp:
    def __init__(self) -> None:
        self.events: list[tuple[str, object]] = []
        self.api = _FakeApi(self.events)
        self.book = _CleanupFailingBook()
        self.books = _BuildBooks(self.book)

    def quit(self) -> None:
        raise OSError("cleanup quit masked the original error")


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


def test_build_preserves_original_write_error_when_cleanup_fails(
    monkeypatch,
    tmp_path,
) -> None:
    app = _CleanupFailingApp()
    monkeypatch.setattr(build_production.xw, "App", lambda **_: app)
    monkeypatch.setattr(
        build_production,
        "load_catalog_document",
        lambda _: SimpleNamespace(
            functions=(),
            regression_sheet_notes={},
            functions_for_sheet=lambda _sheet: (),
        ),
    )
    monkeypatch.setattr(
        build_production,
        "load_life_expectancy_rows",
        lambda _: ([], []),
    )
    monkeypatch.setattr(
        build_production,
        "load_mileage_rows",
        lambda _: ([], []),
    )
    for writer_name in [
        "write_catalog_sheet",
        "write_life_expectancy_sheet",
        "write_mileage_sheet",
        "write_univariate_sheet",
        "write_regression_instructions_sheet",
        "write_diagnostic_guide_sheet",
        "write_version_history_sheet",
    ]:
        monkeypatch.setattr(build_production, writer_name, lambda *_, **__: None)

    def fail_regression_write(*_, **__) -> None:
        raise OSError("remote procedure call failed while writing Formula2")

    monkeypatch.setattr(
        build_production,
        "write_regression_output_sheet",
        fail_regression_write,
    )
    monkeypatch.setattr(
        build_production,
        "sync_workbook_names",
        lambda *_, **__: pytest.fail("sync should not be reached"),
    )

    with pytest.raises(RuntimeError) as exc_info:
        build_production.build_production_workbook(
            workbook_path=tmp_path / "Example.xlsx",
            definitions_path=tmp_path / "lambda_functions.json",
            csv_path=tmp_path / "life_expectancy.csv",
            recalculate=False,
        )

    message = str(exc_info.value)
    assert "remote procedure call failed while writing Formula2" in message
    assert "cleanup close masked" not in message
    assert "cleanup quit masked" not in message


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


def test_main_retries_dropped_rpc_session_during_sheet_write(monkeypatch, capsys) -> None:
    calls: list[bool] = []
    popen_calls: list[tuple[str, ...]] = []

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_univariate=False,
            skip_data_table_calculations=True,
            verify=False,
            no_launch=False,
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=1, updated=2),
    )

    def record_retry(_label, fn, *, retry_rpc=False, **_kwargs) -> None:
        calls.append(cast(bool, retry_rpc))
        fn()

    monkeypatch.setattr(build_production, "_retry_on_open", record_retry)
    # main() ends by shelling out to `start "" <workbook>` to open the just-built
    # file in Excel. That shell call must not fire during tests — the test
    # workbook path is a stub (Example.xlsx) and even if the file existed, no
    # Excel window should pop up from a unit test. Capture the call instead.
    monkeypatch.setattr(
        build_production.subprocess,
        "Popen",
        lambda args: popen_calls.append(tuple(args)),
    )

    build_production.main()

    assert calls == [True]
    # build_production_workbook resolves the workbook path to an absolute path
    # before opening; the same absolute path is what the final `cmd /c start`
    # gets. Compare via Path so the assertion is robust to CWD differences.
    assert len(popen_calls) == 1
    popen_args = popen_calls[0]
    assert popen_args[:4] == ("cmd", "/c", "start", "")
    assert Path(popen_args[4]) == Path("Example.xlsx").resolve()
    output = capsys.readouterr().out
    assert "Created names: 1" in output
    assert "Updated names: 2" in output


class _RecordingBook:
    def __init__(self) -> None:
        self.sheets = _FakeSheetCollection()
        self.saved_paths: list[str] = []
        self.closed = False

    def save(self, path: str) -> None:
        self.saved_paths.append(path)

    def close(self) -> None:
        self.closed = True


class _RecordingBooks:
    def __init__(self, book: _RecordingBook) -> None:
        self.book = book

    def add(self) -> _RecordingBook:
        return self.book

    def open(self, path: str) -> _RecordingBook:
        raise AssertionError("new workbook path should not be opened")


class _RecordingApp:
    def __init__(self) -> None:
        self.events: list[tuple[str, object]] = []
        self.api = _FakeApi(self.events)
        self.book = _RecordingBook()
        self.books = _RecordingBooks(self.book)
        self.quit_called = False

    def quit(self) -> None:
        self.quit_called = True


def test_build_skips_univariate_sheet_when_requested(monkeypatch, tmp_path) -> None:
    app = _RecordingApp()
    monkeypatch.setattr(build_production.xw, "App", lambda **_: app)
    monkeypatch.setattr(
        build_production,
        "load_catalog_document",
        lambda _: SimpleNamespace(
            functions=(),
            workbook_functions=(),
            regression_sheet_notes={},
            functions_for_sheet=lambda _sheet: (),
        ),
    )
    monkeypatch.setattr(
        build_production,
        "load_life_expectancy_rows",
        lambda _: ([], []),
    )
    monkeypatch.setattr(
        build_production,
        "load_mileage_rows",
        lambda _: ([], []),
    )

    writer_calls: list[str] = []
    for writer_name in [
        "write_catalog_sheet",
        "write_life_expectancy_sheet",
        "write_mileage_sheet",
        "write_univariate_sheet",
        "write_regression_instructions_sheet",
        "write_diagnostic_guide_sheet",
        "write_version_history_sheet",
        "write_regression_output_sheet",
    ]:
        monkeypatch.setattr(
            build_production,
            writer_name,
            lambda *_, _name=writer_name, **__: writer_calls.append(_name),
        )
    monkeypatch.setattr(
        build_production,
        "sync_workbook_names",
        lambda *_, **__: NameSyncResult(created=0, updated=0),
    )

    build_production.build_production_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
        skip_univariate=True,
    )

    assert "write_univariate_sheet" not in writer_calls
    assert "write_regression_output_sheet" in writer_calls
    assert app.book.saved_paths == [str(tmp_path / "Example.xlsx")]


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
            skip_univariate=False,
            skip_data_table_calculations=True,
            verify=False,
            no_launch=False,
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


def test_main_no_launch_suppresses_post_build_excel_handoff(
    monkeypatch,
    capsys,
) -> None:
    """When --no-launch is set, main() must not shell out to `cmd /c start`.

    The agentic verifier loop runs the build with --no-launch so the freshly
    built workbook is never opened in Excel — the verifier has already
    proven the workbook is correct. A real `cmd /c start` from the test
    would also pop an unwanted Excel window on the developer's machine.
    """
    popen_calls: list[tuple[str, ...]] = []

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_univariate=False,
            skip_data_table_calculations=True,
            verify=False,
            no_launch=True,
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_production.subprocess,
        "Popen",
        lambda args: popen_calls.append(tuple(args)),
    )

    build_production.main()

    assert popen_calls == []
    assert "Updated names: 0" in capsys.readouterr().out


def test_main_runs_deep_verify_and_exits_zero_on_pass(
    monkeypatch,
    capsys,
) -> None:
    """When --verify is set and the verifier passes, main() must NOT
    shell out to `cmd /c start` until the report is rendered, then
    must shell out as usual on the success path."""
    popen_calls: list[tuple[str, ...]] = []
    verify_calls: list[tuple[Path, Path]] = []

    def fake_run_deep_verify(workbook_path, csv_path, *, verbose=False):
        verify_calls.append((workbook_path, csv_path))
        from lambda_catalog.verify_report import VerifyReport
        return VerifyReport(
            passed=True,
            categories={},
            failures=(),
            elapsed_seconds=0.0,
            mode="spec",
            workbook=str(workbook_path),
        )

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_univariate=False,
            skip_data_table_calculations=True,
            verify=True,
            no_launch=False,
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(build_production, "_run_deep_verify", fake_run_deep_verify)
    monkeypatch.setattr(
        build_production.subprocess,
        "Popen",
        lambda args: popen_calls.append(tuple(args)),
    )

    build_production.main()

    assert len(verify_calls) == 1
    assert verify_calls[0][0] == Path("Example.xlsx").resolve()
    assert verify_calls[0][1] == Path("life_expectancy.csv")
    # Excel handoff fires AFTER verify passes.
    assert len(popen_calls) == 1
    assert popen_calls[0][:4] == ("cmd", "/c", "start", "")
    output = capsys.readouterr().out
    assert "Verify: passed" in output
    assert "spec mode" in output


def test_main_verify_failure_skips_excel_handoff_and_exits_nonzero(
    monkeypatch,
    capsys,
) -> None:
    """When --verify is set and the verifier reports drift, main() must
    NOT shell out to `cmd /c start` (so a stale workbook cannot be
    launched in place of a fresh one) and must sys.exit(1)."""
    popen_calls: list[tuple[str, ...]] = []

    def fake_run_deep_verify(workbook_path, csv_path, *, verbose=False):
        from lambda_catalog.verify_report import VerifyReport
        return VerifyReport(
            passed=False,
            categories={"Regression/scalars": 2},
            failures=(
                "[Regression/scalars] k=1 expected=1.0 excel_calc=1.5",
                "[Regression/scalars] k=2 expected=2.0 excel_calc=2.5",
            ),
            elapsed_seconds=1.23,
            mode="spec",
            workbook=str(workbook_path),
        )

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_univariate=False,
            skip_data_table_calculations=True,
            verify=True,
            no_launch=False,
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(build_production, "_run_deep_verify", fake_run_deep_verify)
    monkeypatch.setattr(
        build_production.subprocess,
        "Popen",
        lambda args: popen_calls.append(tuple(args)),
    )

    with pytest.raises(SystemExit) as exc_info:
        build_production.main()

    assert exc_info.value.code == 1
    # The Excel handoff must NOT have fired — drift means a stale workbook
    # would be opened in place of a fresh one.
    assert popen_calls == []
    output = capsys.readouterr().out
    assert "ERROR Verify mismatch totals" in output
    assert "Regression/scalars=2" in output
