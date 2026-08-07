"""Tests for the Regression production build driver that do not require Excel."""

# pylint: disable=invalid-name,missing-function-docstring,protected-access,too-few-public-methods
from __future__ import annotations

import contextlib
from pathlib import Path
from types import SimpleNamespace
from typing import cast

import pytest

from lambda_catalog.verify_report import VerifyReport
from lambda_catalog.workbook_builder import NameSyncResult
from tests.script_loader import load_script_module

build_production = load_script_module("build_production")


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
        "load_csv_rows",
        lambda _csv_path, _config: ([], []),
    )
    monkeypatch.setattr(
        build_production, "write_csv_dataset_sheet", lambda *_, **__: None
    )
    for writer_name in [
        "write_catalog_sheet",
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
            raise RuntimeError(
                "Excel could not recalculate: The RPC server is unavailable."
            )

    build_production._retry_on_open("Example.xlsx is open", flaky, retry_rpc=True)

    assert calls == 2
    assert "retrying in a fresh Excel instance" in capsys.readouterr().err


def test_main_retries_dropped_rpc_session_during_sheet_write(
    monkeypatch, capsys, tmp_path
) -> None:
    calls: list[bool] = []
    popen_calls: list[tuple[str, ...]] = []

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=False,
            no_launch=False,
            log=tmp_path / "build.log",
            regression_dataset="auto_mpg",
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
    # The Regression build always runs the recalculate phase (the verifier's
    # per-sheet Calculate doesn't rebuild the dependency tree after a name
    # sync), so main() calls _recalculate_and_save even when
    # --skip-data-table-calculations is set. Stub it so no real Excel opens.
    monkeypatch.setattr(
        build_production, "_recalculate_and_save", lambda *_args, **_kwargs: None
    )
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

    # _retry_on_open wraps both the build phase and the recalc phase.
    assert calls == [True, True]
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


def test_workbook_open_check_runs_before_the_first_sheet_write(
    monkeypatch, tmp_path
) -> None:
    """The pre-flight check must precede the write, or it buys nothing.

    Excel opens a locked workbook read-only without raising, so the phase 1
    _retry_on_open only discovers the lock when the closing save fails — after
    every sheet has been written. Ordering IS the feature here: a future
    refactor that moves this call below the build would restore the exact
    behaviour it was added to fix, with no other test noticing.
    """
    order: list[str] = []

    monkeypatch.setattr(
        build_production,
        "warn_if_workbook_open",
        lambda _path, **_kwargs: order.append("check"),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: (order.append("build"), NameSyncResult(created=0, updated=0))[1],
    )
    monkeypatch.setattr(
        build_production, "_recalculate_and_save", lambda *_a, **_k: None
    )

    build_production._build_and_verify(
        SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=False,
            no_launch=True,
            log=tmp_path / "build.log",
            regression_dataset="auto_mpg",
        ),
        Path("Example.xlsx").resolve(),
    )

    assert order == ["check", "build"]


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


def _stub_regression_build_writers(monkeypatch, writer_calls: list[str]) -> None:
    """Stub every writer build_production_workbook calls, recording names."""
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
        "load_csv_rows",
        lambda _csv_path, _config: ([], []),
    )
    monkeypatch.setattr(
        build_production,
        "write_csv_dataset_sheet",
        lambda *_, **__: writer_calls.append("write_csv_dataset_sheet"),
    )
    for writer_name in [
        "write_catalog_sheet",
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


def test_build_never_writes_univariate_sheet(monkeypatch, tmp_path) -> None:
    """The Regression workbook ships no Univariate sheet (it moved to its own
    artifact in v3.0), so build_production_workbook must never call a Univariate
    writer — and indeed does not even import one."""
    app = _RecordingApp()
    monkeypatch.setattr(build_production.xw, "App", lambda **_: app)

    writer_calls: list[str] = []
    _stub_regression_build_writers(monkeypatch, writer_calls)

    build_production.build_production_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
    )

    assert "write_univariate_sheet" not in writer_calls
    assert "write_regression_output_sheet" in writer_calls
    assert "write_csv_dataset_sheet" in writer_calls
    assert app.book.saved_paths == [str(tmp_path / "Example.xlsx")]
    # build_production must not import the Univariate writer at all.
    assert not hasattr(build_production, "write_univariate_sheet")


def test_build_sets_full_automatic_calc_mode(monkeypatch, tmp_path) -> None:
    """The Regression workbook returns to full Automatic (the Univariate Data
    Tables that forced SEMIAUTOMATIC moved to their own artifact)."""
    app = _RecordingApp()
    monkeypatch.setattr(build_production.xw, "App", lambda **_: app)

    writer_calls: list[str] = []
    _stub_regression_build_writers(monkeypatch, writer_calls)

    build_production.build_production_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
    )

    from lambda_catalog.workbook_builder import XL_CALCULATION_AUTOMATIC

    assert app.api.calculation_values[-1] == XL_CALCULATION_AUTOMATIC


def test_main_always_rebuilds_regression(
    monkeypatch,
    tmp_path,
) -> None:
    """main() always runs the recalc rebuild for the Regression workbook (the
    verifier's per-sheet Calculate doesn't rebuild the dependency tree after
    a name sync, so the Regression sheet needs the rebuild to avoid every QC
    value reading nan)."""
    workbook_path = Path("Example.xlsx")
    recalc_calls: list[Path] = []

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=workbook_path,
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=True,
            verify=False,
            no_launch=True,
            log=tmp_path / "build.log",
            regression_dataset="auto_mpg",
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=1, updated=2),
    )
    monkeypatch.setattr(
        build_production,
        "_recalculate_and_save",
        lambda path: recalc_calls.append(path),
    )
    monkeypatch.setattr(
        build_production.subprocess,
        "Popen",
        lambda args: None,
    )

    build_production.main()

    assert len(recalc_calls) == 1


def test_main_no_launch_suppresses_post_build_excel_handoff(
    monkeypatch,
    capsys,
    tmp_path,
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
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=False,
            no_launch=True,
            log=tmp_path / "build.log",
            regression_dataset="auto_mpg",
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_production, "_recalculate_and_save", lambda *_args, **_kwargs: None
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
    tmp_path,
) -> None:
    """When --verify is set and the verifier passes, main() must NOT
    shell out to `cmd /c start` until the report is rendered, then
    must shell out as usual on the success path."""
    popen_calls: list[tuple[str, ...]] = []
    verify_calls: list[tuple[Path, Path]] = []

    def fake_run_deep_verify(
        workbook_path,
        csv_path,
        *,
        mileage_path=None,
        production_lots_path=None,
        verbose=False,
    ):
        del mileage_path, production_lots_path, verbose
        verify_calls.append((workbook_path, csv_path))
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
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=True,
            no_launch=False,
            log=tmp_path / "build.log",
            regression_dataset="auto_mpg",
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_production, "_recalculate_and_save", lambda *_args, **_kwargs: None
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
    tmp_path,
) -> None:
    """When --verify is set and the verifier reports drift, main() must
    NOT shell out to `cmd /c start` (so a stale workbook cannot be
    launched in place of a fresh one) and must sys.exit(1)."""
    popen_calls: list[tuple[str, ...]] = []

    def fake_run_deep_verify(
        workbook_path,
        csv_path,
        *,
        mileage_path=None,
        production_lots_path=None,
        verbose=False,
    ):
        del csv_path, mileage_path, production_lots_path, verbose
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
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=True,
            no_launch=False,
            log=tmp_path / "build.log",
            regression_dataset="auto_mpg",
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_production, "_recalculate_and_save", lambda *_args, **_kwargs: None
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


def test_main_archives_the_verify_report_to_the_run_log(
    monkeypatch,
    tmp_path,
) -> None:
    """A failing --verify run must leave its VerifyReport in the transcript.

    `poe verify-deep` cannot run on the GitHub-hosted Linux CI (no Excel),
    so `excel-only-runs/` is the only place a reviewer or a future agent can
    read what a failed deep verify actually said. The report has to be IN the
    file, not just on the terminal that ran it.
    """
    log_path = tmp_path / "build_production verify.log"

    def fake_run_deep_verify(workbook_path, csv_path, **_kwargs):
        del csv_path
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
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=True,
            no_launch=True,
            log=log_path,
            regression_dataset="auto_mpg",
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_production, "_recalculate_and_save", lambda *_args, **_kwargs: None
    )
    monkeypatch.setattr(build_production, "_run_deep_verify", fake_run_deep_verify)

    with pytest.raises(SystemExit) as exc_info:
        build_production.main()

    assert exc_info.value.code == 1
    archived = log_path.read_text(encoding="utf-8")
    assert "python build_production.py" in archived
    assert "ERROR Verify mismatch totals" in archived
    assert "Regression/scalars=2" in archived
    assert "[Regression/scalars] k=1 expected=1.0 excel_calc=1.5" in archived
    # The report is the payload, not a SystemExit stack trace: main() exits
    # after the tee closes so the archived file stays readable.
    assert "Traceback" not in archived
    assert f"Log: {log_path}" in archived


def test_main_archives_the_traceback_when_the_build_aborts(
    monkeypatch,
    tmp_path,
) -> None:
    """A build that dies mid-run archives its own cause.

    The failure most worth capturing here is a pywintypes.com_error, which
    never touches stdout — so the transcript has to carry stderr and the
    traceback, and the exception still has to propagate to the shell.
    """
    log_path = tmp_path / "build_production verify.log"

    def explode(**_kwargs):
        raise RuntimeError("Excel dispatch blew up")

    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=True,
            no_launch=True,
            log=log_path,
            regression_dataset="auto_mpg",
        ),
    )
    monkeypatch.setattr(build_production, "build_production_workbook", explode)

    with pytest.raises(RuntimeError, match="Excel dispatch blew up"):
        build_production.main()

    archived = log_path.read_text(encoding="utf-8")
    assert "Traceback" in archived
    assert "Excel dispatch blew up" in archived


def test_main_defaults_the_run_log_to_excel_only_runs(monkeypatch) -> None:
    """Without --log, the transcript lands in excel-only-runs/ named after the run.

    Recorded rather than written: an actual file here would drop a unit-test
    artifact into a tracked directory whose contents are meant to be real
    Excel-required runs.
    """
    recorded: list[Path] = []

    @contextlib.contextmanager
    def fake_tee(log_path, command):
        del command
        recorded.append(log_path)
        yield log_path

    monkeypatch.setattr(build_production, "tee_run_log", fake_tee)
    monkeypatch.setattr(
        build_production.sys,
        "argv",
        ["scripts/build_production.py", "--verify", "--no-launch"],
    )
    monkeypatch.setattr(
        build_production,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            mileage_csv=Path("mileage.csv"),
            production_lots_csv=Path("production_lots.csv"),
            validate_reopen=False,
            verbose=False,
            verify=False,
            no_launch=True,
            log=None,
            regression_dataset="auto_mpg",
        ),
    )
    monkeypatch.setattr(
        build_production,
        "build_production_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_production, "_recalculate_and_save", lambda *_args, **_kwargs: None
    )

    build_production.main()

    assert recorded == [
        build_production.ROOT_DIR
        / "excel-only-runs"
        / "build_production verify no launch.log"
    ]


def test_build_uses_life_expectancy_source_table_when_requested(
    monkeypatch, tmp_path
) -> None:
    """regression_dataset='life_expectancy' must pass LifeExpectancyData[#All]
    as source_table_ref to write_regression_output_sheet."""
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
        "load_csv_rows",
        lambda _csv_path, _config: ([], []),
    )

    called: dict[str, object] = {"source_table_ref": None}
    monkeypatch.setattr(build_production, "write_catalog_sheet", lambda *_, **__: None)
    monkeypatch.setattr(
        build_production, "write_csv_dataset_sheet", lambda *_, **__: None
    )
    monkeypatch.setattr(
        build_production, "write_regression_instructions_sheet", lambda *_, **__: None
    )
    monkeypatch.setattr(
        build_production, "write_diagnostic_guide_sheet", lambda *_, **__: None
    )
    monkeypatch.setattr(
        build_production, "write_version_history_sheet", lambda *_, **__: None
    )
    monkeypatch.setattr(
        build_production,
        "write_regression_output_sheet",
        lambda *args, **kwargs: called.__setitem__(
            "source_table_ref", kwargs.get("source_table_ref")
        ),
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
        regression_dataset="life_expectancy",
    )

    assert called["source_table_ref"] == "=LifeExpectancyData[#All]"


def test_build_defaults_to_life_expectancy_source_table(monkeypatch, tmp_path) -> None:
    """Default regression_dataset must resolve to LifeExpectancyData[#All]."""
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
        build_production, "load_csv_rows", lambda _csv_path, _config: ([], [])
    )

    called: dict[str, object] = {"source_table_ref": None}
    monkeypatch.setattr(
        build_production, "write_csv_dataset_sheet", lambda *_, **__: None
    )
    for name in [
        "write_catalog_sheet",
        "write_regression_instructions_sheet",
        "write_diagnostic_guide_sheet",
        "write_version_history_sheet",
    ]:
        monkeypatch.setattr(build_production, name, lambda *_, **__: None)
    monkeypatch.setattr(
        build_production,
        "write_regression_output_sheet",
        lambda *args, **kwargs: called.__setitem__(
            "source_table_ref", kwargs.get("source_table_ref")
        ),
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
    )

    assert called["source_table_ref"] == "=LifeExpectancyData[#All]"


class _TabOrderSheet:
    def __init__(self, name: str) -> None:
        self.name = name


class _TabOrderSheets:
    def __init__(self, names: list[str]) -> None:
        self._items = [_TabOrderSheet(name) for name in names]

    def __iter__(self):
        return iter(self._items)

    def __getitem__(self, key):
        if isinstance(key, str):
            return next(sheet for sheet in self._items if sheet.name == key)
        return self._items[key]

    def names(self) -> list[str]:
        return [sheet.name for sheet in self._items]

    def move_before(self, moving_name: str, anchor_name: str) -> None:
        moving_index = self.names().index(moving_name)
        anchor_index = self.names().index(anchor_name)
        sheet = self._items.pop(moving_index)
        if moving_index < anchor_index:
            anchor_index -= 1
        self._items.insert(anchor_index, sheet)


class _TabOrderBook:
    def __init__(self, names: list[str]) -> None:
        self.sheets = _TabOrderSheets(names)


def test_reorder_and_style_sheet_tabs_orders_front_matter_and_sets_colors(
    monkeypatch,
) -> None:
    """The Regression workbook's tab order puts the three data sheets first, then
    Version History, then the Regression workbench sheets, with LAMBDA_functions
    last. There is no Univariate tab (it ships in its own workbook)."""
    book = _TabOrderBook(
        [
            "LAMBDA_functions",
            "Regression",
            "Diagnostic Guide",
            "Mileage Data",
            "Life Expectancy Data",
            "Production Lots",
            "Version History",
            "Regression Instructions",
        ]
    )
    tab_colors: dict[str, tuple[int, int, int]] = {}

    def fake_move_sheet_before(sheet, anchor) -> None:
        book.sheets.move_before(sheet.name, anchor.name)

    def fake_set_tab_color(sheet, color) -> None:
        tab_colors[sheet.name] = color

    monkeypatch.setattr(build_production, "_move_sheet_before", fake_move_sheet_before)
    monkeypatch.setattr(build_production, "_set_tab_color", fake_set_tab_color)

    build_production._reorder_and_style_sheet_tabs(book)

    assert book.sheets.names()[:8] == [
        "Mileage Data",
        "Life Expectancy Data",
        "Production Lots",
        "Version History",
        "Regression Instructions",
        "Regression",
        "Diagnostic Guide",
        "LAMBDA_functions",
    ]
    assert tab_colors == {
        "Mileage Data": (217, 217, 217),
        "Life Expectancy Data": (217, 217, 217),
        "Production Lots": (217, 217, 217),
        "Version History": (128, 128, 128),
        "Regression Instructions": build_production.SUBHDR_COLOR,
        "Regression": build_production.SUBHDR_COLOR,
        "Diagnostic Guide": build_production.SUBHDR_COLOR,
    }


class _FakeVerifyRecorder:
    """Records kwargs passed to deep_verify.verify_test_sheets."""

    def __init__(self) -> None:
        self.verify_kwargs: dict = {}

    def verify_test_sheets(
        self, workbook, regression_sheet_configs, csv_path, **kwargs
    ) -> None:
        self.verify_kwargs = {
            "workbook": workbook,
            "regression_sheet_configs": regression_sheet_configs,
            "csv_path": csv_path,
            **kwargs,
        }


def test_run_deep_verify_forwards_skip_regression_false_and_skip_univariate_true(
    monkeypatch,
) -> None:
    """The Regression artifact's verify path must run the full Regression check
    (skip_regression defaults to False) while still skipping the Univariate
    check (the Regression workbook ships no Univariate sheet)."""
    app = _FakeApp()
    monkeypatch.setattr(build_production.xw, "App", lambda **_: app)

    fake_verify = _FakeVerifyRecorder()
    monkeypatch.setattr(
        build_production, "verify_test_sheets", fake_verify.verify_test_sheets
    )
    monkeypatch.setattr(
        build_production,
        "build_regression_spec_qc_configs",
        lambda mileage_path: ["_fake_regression_configs"],
    )

    report = build_production._run_deep_verify(
        Path("Example.xlsx"),
        Path("life_expectancy.csv"),
        mileage_path=Path("mileage.csv"),
        production_lots_path=Path("production_lots.csv"),
    )

    assert report.passed is True
    # _run_deep_verify does not pass skip_regression explicitly, so it defaults
    # to False — the full Regression check runs (this artifact HAS a Regression
    # sheet). skip_univariate=True is passed explicitly (no Univariate sheet).
    assert fake_verify.verify_kwargs["skip_univariate"] is True
    assert fake_verify.verify_kwargs.get("skip_regression", False) is False
    # The Regression path passes real regression_sheet_configs (not None).
    assert fake_verify.verify_kwargs["regression_sheet_configs"] is not None
