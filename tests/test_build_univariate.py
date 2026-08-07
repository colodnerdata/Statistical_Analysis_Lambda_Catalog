"""Tests for the standalone Univariate build driver (build_univariate.py).

These tests do not require Excel. They mirror tests/test_build_production.py's
stub-the-writers / capture-the-shell-out pattern: the writers and the Excel
open are stubbed via monkeypatch, and the tests assert the Univariate
artifact's sheet set, default output path, calculation mode, the
--skip-data-table-calculations gate on the rebuild, and that --verify
forwards skip_regression=True to the spec-driven verifier.
"""

# pylint: disable=invalid-name,missing-function-docstring,protected-access,too-few-public-methods
from __future__ import annotations

import contextlib
from pathlib import Path
from types import SimpleNamespace

import pytest

from lambda_catalog.verify_report import VerifyReport
from lambda_catalog.workbook_builder import (
    XL_CALCULATION_AUTOMATIC,
    XL_CALCULATION_MANUAL,
    NameSyncResult,
)
from tests.script_loader import load_script_module

build_univariate = load_script_module("build_univariate")


class _FakeApi:
    def __init__(self, events: list[tuple[str, object]]) -> None:
        self.events = events
        self.DisplayAlerts = True
        self.AskToUpdateLinks = True
        self.calculation_values: list[int] = []
        # Excel's default: under Manual, saving still recalculates first.
        self._calculate_before_save = True
        self.calculate_before_save_values: list[bool] = []

    @property
    def CalculateBeforeSave(self) -> bool:
        return self._calculate_before_save

    @CalculateBeforeSave.setter
    def CalculateBeforeSave(self, value: bool) -> None:
        self.events.append(("calculate_before_save", value))
        self.calculate_before_save_values.append(value)
        self._calculate_before_save = value

    @property
    def Calculation(self) -> int | None:
        return self.calculation_values[-1] if self.calculation_values else None

    @Calculation.setter
    def Calculation(self, value: int) -> None:
        self.events.append(("calculation", value))
        self.calculation_values.append(value)


class _FakeSheet:
    def __init__(self, name: str) -> None:
        self.name = name
        self.deleted = False

    def delete(self) -> None:
        self.deleted = True


class _FakeSheetCollection:
    """Mutable sheet collection: starts with only the target Univariate sheets
    so the stray-sheet-deletion path can be exercised by adding extras."""

    def __init__(self, names: list[str] | None = None) -> None:
        self.items = [_FakeSheet(n) for n in (names or ["Sheet1"])]

    def __iter__(self):
        return iter(self.items)

    def __getitem__(self, key):
        if isinstance(key, str):
            return next(sheet for sheet in self.items if sheet.name == key)
        return self.items[key]


class _FakeBook:
    def __init__(self, sheet_names: list[str] | None = None) -> None:
        self.sheets = _FakeSheetCollection(sheet_names)
        self.saved_paths: list[str] = []
        self.closed = False

    def save(self, path: str) -> None:
        self.saved_paths.append(path)

    def close(self) -> None:
        self.closed = True


class _FakeBooks:
    def __init__(self, book: _FakeBook) -> None:
        self.book = book

    def add(self) -> _FakeBook:
        return self.book

    def open(self, path: str) -> _FakeBook:
        return self.book


class _FakeApp:
    def __init__(self, sheet_names: list[str] | None = None) -> None:
        self.events: list[tuple[str, object]] = []
        self.api = _FakeApi(self.events)
        self.book = _FakeBook(sheet_names)
        self.books = _FakeBooks(self.book)
        self.quit_called = False

    def quit(self) -> None:
        self.quit_called = True


def test_default_workbook_path_is_univariate_xlsx() -> None:
    """The default output is the capital-U Lambda_Library_Univariate.xlsx per
    the README/CONTRIBUTING two-artifact naming convention."""
    assert (
        build_univariate.DEFAULT_UNIVARIATE_WORKBOOK_PATH.name
        == "Lambda_Library_Univariate.xlsx"
    )


def _stub_writers(monkeypatch, writer_calls: list[str]) -> None:
    """Stub every writer build_univariate_workbook calls, recording names."""
    monkeypatch.setattr(
        build_univariate,
        "load_catalog_document",
        lambda _: SimpleNamespace(
            functions=(),
            workbook_functions=(),
            univariate_sheet_notes={},
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "load_csv_rows",
        lambda _csv_path, _config: ([], []),
    )
    monkeypatch.setattr(
        build_univariate,
        "write_csv_dataset_sheet",
        lambda *_, **__: writer_calls.append("write_csv_dataset_sheet"),
    )
    for writer_name in [
        "write_catalog_sheet",
        "write_univariate_sheet",
        "write_version_history_sheet",
    ]:
        monkeypatch.setattr(
            build_univariate,
            writer_name,
            lambda *_, _name=writer_name, **__: writer_calls.append(_name),
        )
    monkeypatch.setattr(
        build_univariate,
        "sync_workbook_names",
        lambda *_, **__: NameSyncResult(created=0, updated=0),
    )


def test_build_writes_univariate_sheet_set(monkeypatch, tmp_path) -> None:
    """The standalone Univariate workbook ships exactly four sheets:
    LAMBDA_functions, Life Expectancy Data, Univariate, Version History — and
    NOT the Regression / Mileage / Production Lots / instructions / diagnostic
    sheets (those belong to the Regression artifact)."""
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)

    writer_calls: list[str] = []
    _stub_writers(monkeypatch, writer_calls)

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
    )

    assert "write_univariate_sheet" in writer_calls
    assert "write_version_history_sheet" in writer_calls
    assert "write_catalog_sheet" in writer_calls
    assert writer_calls.count("write_csv_dataset_sheet") == 1  # Life Expectancy only
    # Regression-side writers are not even imported by this script.
    for absent in (
        "write_regression_output_sheet",
        "write_regression_instructions_sheet",
        "write_diagnostic_guide_sheet",
    ):
        assert absent not in writer_calls
        assert not hasattr(build_univariate, absent)
    assert app.book.saved_paths == [str(tmp_path / "Example.xlsx")]


def test_build_passes_univariate_artifact_to_version_history(
    monkeypatch, tmp_path
) -> None:
    """The Version History sheet must be written with artifact='univariate' so
    it carries the Univariate lineage (1.0.0), not the Regression lineage."""
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)
    monkeypatch.setattr(
        build_univariate,
        "load_catalog_document",
        lambda _: SimpleNamespace(
            functions=(), workbook_functions=(), univariate_sheet_notes={}
        ),
    )
    monkeypatch.setattr(build_univariate, "load_csv_rows", lambda *_a, **_k: ([], []))
    monkeypatch.setattr(build_univariate, "write_catalog_sheet", lambda *_, **__: None)
    monkeypatch.setattr(
        build_univariate, "write_csv_dataset_sheet", lambda *_, **__: None
    )
    monkeypatch.setattr(
        build_univariate, "write_univariate_sheet", lambda *_, **__: None
    )

    version_calls: list[dict] = []

    def capture_version_history(_workbook, *, artifact="regression"):
        version_calls.append({"artifact": artifact})

    monkeypatch.setattr(
        build_univariate, "write_version_history_sheet", capture_version_history
    )
    monkeypatch.setattr(
        build_univariate,
        "sync_workbook_names",
        lambda *_, **__: NameSyncResult(created=0, updated=0),
    )

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
    )

    assert version_calls == [{"artifact": "univariate"}]


def test_build_sets_full_automatic_calc_mode(monkeypatch, tmp_path) -> None:
    """The Univariate workbook ships in full Automatic so the Beta
    Full_Factorial grid and the Weibull/Gamma profile-NLL columns recalculate
    on edit (the whole point of the split)."""
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)

    writer_calls: list[str] = []
    _stub_writers(monkeypatch, writer_calls)

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
    )

    assert app.api.calculation_values[-1] == XL_CALCULATION_AUTOMATIC


def test_no_calculation_never_leaves_manual_mode(monkeypatch, tmp_path) -> None:
    """calculate=False must never set Automatic.

    Setting Automatic on the open workbook is itself a full calculation — the
    Beta Full_Factorial grid included — so it is the step that has to not
    happen, not just the final rebuild. --skip-data-table-calculations does not
    cover it.
    """
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)

    writer_calls: list[str] = []
    _stub_writers(monkeypatch, writer_calls)

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
        calculate=False,
    )

    assert XL_CALCULATION_AUTOMATIC not in app.api.calculation_values
    assert app.api.calculation_values == [XL_CALCULATION_MANUAL]
    # The sheets are still written — the point is inspecting them uncalculated.
    assert "write_univariate_sheet" in writer_calls
    assert app.book.saved_paths == [str(tmp_path / "Example.xlsx")]


def test_no_calculation_suppresses_calculate_before_save_then_restores_it(
    monkeypatch, tmp_path
) -> None:
    """Manual mode alone still calculates on save; the setting must go off.

    And it must come back on: Application.CalculateBeforeSave is Excel-wide, so
    leaving it off would silently change every later session on that machine.
    """
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)
    _stub_writers(monkeypatch, [])

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
        calculate=False,
    )

    assert app.api.calculate_before_save_values == [False, True]
    assert app.api.CalculateBeforeSave is True

    # The save happens while it is still suppressed, not after the restore.
    save_index = app.events.index(("calculate_before_save", True))
    assert ("calculate_before_save", False) in app.events[:save_index]


def test_default_build_leaves_calculate_before_save_alone(
    monkeypatch, tmp_path
) -> None:
    """An ordinary build must not touch the user's Excel-wide setting at all."""
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)
    _stub_writers(monkeypatch, [])

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
    )

    assert app.api.calculate_before_save_values == []


def test_no_calculation_skips_the_rebuild_even_when_recalculate_is_true(
    monkeypatch, tmp_path
) -> None:
    """calculate=False wins over recalculate=True — nothing calculates."""
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)
    _stub_writers(monkeypatch, [])

    recalc_calls: list[Path] = []
    monkeypatch.setattr(
        build_univariate,
        "_recalculate_and_save",
        lambda path: recalc_calls.append(path),
    )

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=True,
        calculate=False,
    )

    assert recalc_calls == []


def test_main_no_calculation_skips_rebuild_and_warns(
    monkeypatch, capsys, tmp_path
) -> None:
    """--no-calculation skips phase 2 and says the artifact is not shippable."""
    recalc_calls: list[Path] = []
    build_kwargs: dict = {}

    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_data_table_calculations=False,
            no_calculation=True,
            verify=False,
            no_launch=True,
            beta_grid_size=10,
            log=tmp_path / "build.log",
        ),
    )

    def capture_build(**kwargs):
        build_kwargs.update(kwargs)
        return NameSyncResult(created=0, updated=0)

    monkeypatch.setattr(build_univariate, "build_univariate_workbook", capture_build)
    monkeypatch.setattr(
        build_univariate,
        "_recalculate_and_save",
        lambda path: recalc_calls.append(path),
    )
    monkeypatch.setattr(build_univariate.subprocess, "Popen", lambda args: None)

    build_univariate.main()

    assert build_kwargs["calculate"] is False
    # Skipped despite --skip-data-table-calculations being off.
    assert recalc_calls == []
    out = capsys.readouterr().out
    assert "--no-calculation" in out
    assert "do not ship it" in out


def test_main_warns_when_verify_is_combined_with_no_calculation(
    monkeypatch, capsys, tmp_path
) -> None:
    """Verifying a workbook that was never calculated compares nothing useful."""
    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_data_table_calculations=True,
            no_calculation=True,
            verify=True,
            no_launch=True,
            beta_grid_size=10,
            log=tmp_path / "build.log",
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "build_univariate_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_univariate, "_recalculate_and_save", lambda *_a, **_k: None
    )
    monkeypatch.setattr(
        build_univariate,
        "_run_deep_verify",
        lambda *_a, **_k: VerifyReport(
            passed=True,
            categories={},
            failures=(),
            elapsed_seconds=0.0,
            mode="spec",
            workbook="Example.xlsx",
        ),
    )
    monkeypatch.setattr(build_univariate.subprocess, "Popen", lambda args: None)

    build_univariate.main()

    assert "spurious mismatches" in capsys.readouterr().err


def test_build_drops_stray_regression_sheets_from_reused_file(
    monkeypatch, tmp_path
) -> None:
    """If the output path is a reused workbook (e.g. a copied Lambda_Library.xlsx),
    build_univariate_workbook must delete every sheet that is not part of the
    Univariate artifact before writing, so no stray Regression / Mileage /
    Production Lots sheets leak into the shipped Univariate workbook."""
    stray_names = ["Regression", "Mileage Data", "Production Lots", "Diagnostic Guide"]
    target_names = [
        "LAMBDA_functions",
        "Life Expectancy Data",
        "Univariate",
        "Version History",
    ]
    app = _FakeApp(sheet_names=stray_names + target_names)
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)

    writer_calls: list[str] = []
    _stub_writers(monkeypatch, writer_calls)
    # Tab reorder/coloring touches sheet.api.Move / sheet.api.Tab.Color; the
    # fake sheets have no .api, so stub the two tab helpers so the post-write
    # _reorder_and_style_sheet_tabs survives. (This test is about stray-sheet
    # deletion, not tab order.)
    monkeypatch.setattr(build_univariate, "_move_sheet_before", lambda *_a, **_k: None)
    monkeypatch.setattr(build_univariate, "_set_tab_color", lambda *_a, **_k: None)

    build_univariate.build_univariate_workbook(
        workbook_path=tmp_path / "Example.xlsx",
        definitions_path=tmp_path / "lambda_functions.json",
        csv_path=tmp_path / "life_expectancy.csv",
        recalculate=False,
    )

    deleted = {sheet.name for sheet in app.book.sheets.items if sheet.deleted}
    assert deleted == set(stray_names)
    # Target sheets (and the placeholder Sheet1) are not deleted.
    for name in target_names:
        assert not any(
            sheet.name == name and sheet.deleted for sheet in app.book.sheets.items
        )


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


def test_reorder_puts_univariate_front_and_lambda_last(monkeypatch) -> None:
    """The Univariate workbench sits front-most (the reason the file exists),
    followed by its data source and version history, with the LAMBDA_functions
    catalog last."""
    book = _TabOrderBook(
        ["LAMBDA_functions", "Version History", "Life Expectancy Data", "Univariate"]
    )
    tab_colors: dict[str, tuple[int, int, int]] = {}

    def fake_move_sheet_before(sheet, anchor) -> None:
        book.sheets.move_before(sheet.name, anchor.name)

    def fake_set_tab_color(sheet, color) -> None:
        tab_colors[sheet.name] = color

    monkeypatch.setattr(build_univariate, "_move_sheet_before", fake_move_sheet_before)
    monkeypatch.setattr(build_univariate, "_set_tab_color", fake_set_tab_color)

    build_univariate._reorder_and_style_sheet_tabs(book)

    assert book.sheets.names() == [
        "Univariate",
        "Life Expectancy Data",
        "Version History",
        "LAMBDA_functions",
    ]
    assert tab_colors == {
        "Univariate": build_univariate.SUBHDR_COLOR,
        "Life Expectancy Data": (217, 217, 217),
        "Version History": (128, 128, 128),
    }


def test_main_recalculate_skipped_when_skip_data_table_calculations(
    monkeypatch, capsys, tmp_path
) -> None:
    """--skip-data-table-calculations skips the (slow, Full_Factorial grid +
    profile-NLL) recalc rebuild for the Univariate artifact — this is the
    flag's now-primary purpose."""
    recalc_calls: list[Path] = []

    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=True,
            skip_data_table_calculations=True,
            no_calculation=False,
            verify=False,
            no_launch=True,
            beta_grid_size=10,
            log=tmp_path / "build.log",
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "build_univariate_workbook",
        lambda **_: NameSyncResult(created=1, updated=2),
    )
    monkeypatch.setattr(
        build_univariate,
        "_recalculate_and_save",
        lambda path: recalc_calls.append(path),
    )
    monkeypatch.setattr(build_univariate.subprocess, "Popen", lambda args: None)

    build_univariate.main()

    assert recalc_calls == []
    assert (
        "Recalculate:    skipped (--skip-data-table-calculations)"
        in capsys.readouterr().out
    )


def test_main_recalculate_runs_by_default(monkeypatch, tmp_path) -> None:
    """Without --skip-data-table-calculations, main() runs the rebuild so the
    shipped artifact's Beta Full_Factorial grid and profile-NLL columns are
    computed (not stale pending Ctrl+Alt+F9)."""
    recalc_calls: list[Path] = []

    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_data_table_calculations=False,
            no_calculation=False,
            verify=False,
            no_launch=True,
            beta_grid_size=10,
            log=tmp_path / "build.log",
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "build_univariate_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_univariate,
        "_recalculate_and_save",
        lambda path: recalc_calls.append(path),
    )
    monkeypatch.setattr(build_univariate.subprocess, "Popen", lambda args: None)

    build_univariate.main()

    assert len(recalc_calls) == 1


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


def test_run_deep_verify_forwards_skip_regression_true(monkeypatch) -> None:
    """The Univariate artifact's verify path must skip every Regression /
    Mileage / Production Lots check (skip_regression=True) and pass None for
    regression_sheet_configs — this artifact carries none of those sheets."""
    app = _FakeApp()
    monkeypatch.setattr(build_univariate.xw, "App", lambda **_: app)

    fake_verify = _FakeVerifyRecorder()
    monkeypatch.setattr(
        build_univariate, "verify_test_sheets", fake_verify.verify_test_sheets
    )

    report = build_univariate._run_deep_verify(
        Path("Example.xlsx"),
        Path("life_expectancy.csv"),
    )

    assert report.passed is True
    assert fake_verify.verify_kwargs["skip_regression"] is True
    assert fake_verify.verify_kwargs["regression_sheet_configs"] is None
    # The Univariate check is NOT skipped (the sheet is present in this artifact).
    assert fake_verify.verify_kwargs.get("skip_univariate") is not True


def test_main_verify_forwards_skip_regression_true(monkeypatch, tmp_path) -> None:
    """--verify on the Univariate build must run _run_deep_verify, which
    forwards skip_regression=True to the verifier."""
    verify_calls: list[tuple] = []

    def fake_run_deep_verify(workbook_path, csv_path, *, verbose=False):
        verify_calls.append((workbook_path, csv_path, verbose))
        return VerifyReport(
            passed=True,
            categories={},
            failures=(),
            elapsed_seconds=0.0,
            mode="spec",
            workbook=str(workbook_path),
        )

    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_data_table_calculations=True,
            no_calculation=False,
            verify=True,
            no_launch=True,
            beta_grid_size=10,
            log=tmp_path / "build.log",
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "build_univariate_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_univariate, "_recalculate_and_save", lambda *_a, **_k: None
    )
    monkeypatch.setattr(build_univariate, "_run_deep_verify", fake_run_deep_verify)
    monkeypatch.setattr(build_univariate.subprocess, "Popen", lambda args: None)

    build_univariate.main()

    assert len(verify_calls) == 1
    assert verify_calls[0][0] == Path("Example.xlsx").resolve()


def test_main_archives_the_verify_report_to_the_run_log(monkeypatch, tmp_path) -> None:
    """A failing --verify run must leave its VerifyReport in the transcript.

    Same reasoning as the Regression driver's copy of this test: this build
    needs Excel, so `excel-only-runs/` is the only place the output of a
    failed `poe verify-deep-univariate` survives the terminal that ran it.
    """
    log_path = tmp_path / "build_univariate verify.log"

    def fake_run_deep_verify(workbook_path, csv_path, *, verbose=False):
        del csv_path, verbose
        return VerifyReport(
            passed=False,
            categories={"Univariate/fits": 1},
            failures=("[Univariate/fits] Weibull shape expected=2.0 excel_calc=2.5",),
            elapsed_seconds=4.5,
            mode="spec",
            workbook=str(workbook_path),
        )

    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_data_table_calculations=False,
            no_calculation=False,
            verify=True,
            no_launch=True,
            beta_grid_size=10,
            log=log_path,
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "build_univariate_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_univariate, "_recalculate_and_save", lambda *_a, **_k: None
    )
    monkeypatch.setattr(build_univariate, "_run_deep_verify", fake_run_deep_verify)

    with pytest.raises(SystemExit) as exc_info:
        build_univariate.main()

    assert exc_info.value.code == 1
    archived = log_path.read_text(encoding="utf-8")
    assert "python build_univariate.py" in archived
    assert "ERROR Verify mismatch totals" in archived
    assert "Univariate/fits=1" in archived
    assert "[Univariate/fits] Weibull shape expected=2.0 excel_calc=2.5" in archived
    assert "Traceback" not in archived


def test_main_archives_the_no_calculation_verify_warning(monkeypatch, tmp_path) -> None:
    """The --no-calculation + --verify warning goes to stderr; it must be archived.

    stderr is exactly what a terminal-only transcript loses first, and this
    warning is what explains a wall of spurious fit mismatches further down
    the same file.
    """
    log_path = tmp_path / "build_univariate verify no calculation.log"

    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_data_table_calculations=False,
            no_calculation=True,
            verify=True,
            no_launch=True,
            beta_grid_size=10,
            log=log_path,
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "build_univariate_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_univariate,
        "_run_deep_verify",
        lambda workbook_path, csv_path, *, verbose=False: VerifyReport(
            passed=True,
            categories={},
            failures=(),
            elapsed_seconds=0.0,
            mode="spec",
            workbook=str(workbook_path),
        ),
    )

    build_univariate.main()

    archived = log_path.read_text(encoding="utf-8")
    assert "--verify with --no-calculation" in archived


def test_main_defaults_the_run_log_to_excel_only_runs(monkeypatch) -> None:
    """Without --log, the transcript lands in excel-only-runs/ named after the run.

    Recorded rather than written, so a unit test never drops an artifact
    into a tracked directory that is meant to hold real Excel-required runs.
    """
    recorded: list[Path] = []

    @contextlib.contextmanager
    def fake_tee(log_path, command):
        del command
        recorded.append(log_path)
        yield log_path

    monkeypatch.setattr(build_univariate, "tee_run_log", fake_tee)
    monkeypatch.setattr(
        build_univariate.sys,
        "argv",
        ["scripts/build_univariate.py", "--verify", "--no-launch"],
    )
    monkeypatch.setattr(
        build_univariate,
        "parse_args",
        lambda: SimpleNamespace(
            workbook=Path("Example.xlsx"),
            definitions=Path("lambda_functions.json"),
            csv=Path("life_expectancy.csv"),
            validate_reopen=False,
            verbose=False,
            skip_data_table_calculations=True,
            no_calculation=False,
            verify=False,
            no_launch=True,
            beta_grid_size=10,
            log=None,
        ),
    )
    monkeypatch.setattr(
        build_univariate,
        "build_univariate_workbook",
        lambda **_: NameSyncResult(created=0, updated=0),
    )
    monkeypatch.setattr(
        build_univariate, "_recalculate_and_save", lambda *_a, **_k: None
    )

    build_univariate.main()

    assert recorded == [
        build_univariate.ROOT_DIR
        / "excel-only-runs"
        / "build_univariate verify no launch.log"
    ]
