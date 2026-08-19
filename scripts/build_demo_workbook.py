"""Build presentations/Lambda_Library_Regression_Demo.xlsx — one sheet per talk beat.

A build target alongside ``build_production.py`` (the shipped
Regression artifact) and ``build_test_models.py`` (the QC fixture). This one
is a **presentation
handout**: the workbook a speaker opens on stage so the audience can follow
the talk's regression beats live. Both decks headline the same four-driver
Life Expectancy model (slide 19's coefficient table), so the demo opens on
exactly that model — the shipped default — and walks the declare → screen →
trim → diagnose → predict → fixed-effects arc one tab at a time.

Why a separate workbook and not the shipped artifact. The shipped
``dist/Lambda_Library.xlsx`` is one Regression sheet the speaker would have to
re-spec on stage for every beat. This workbook has one sheet per beat, each
pre-fitted and cached, so a low-compute presentation machine opens to fully
computed values and only recomputes one sheet at a time when the speaker
toggles a cell. Charts are on the three plot-teaching beats (Diagnose, Predict,
Fixed Effects) and off the three spec-building beats (Declare, Screen, Trim),
where the visual is not the point — keeping the build light on a machine that
may have little to spare.

What it reuses. Nothing reimplements the Regression sheet. Every beat sheet is
``write_test_model_sheet`` (the same writer ``build_test_models.py`` uses for
the QC suite) called with ``include_charts`` on the plot beats. Beats 1 and 6
reuse the registered ``life_talk_demo`` (L11) and ``production_lots_fixed_effects``
(P01) cases verbatim, retargeted to a presentation tab name; beats 2–5 are
unregistered ``RegressionSpecCase`` snapshots with explicit sheet names
(``"2 Screen"`` etc.) that bypass the QC registry's plan-id naming contract —
they are talk snapshots, not QC cases, so the numbered tab names are legal
here. Each beat's oracle is computed via ``calculate_regression_spec_case``
(the same NumPy/statsmodels oracle the QC harness uses) to prefill the
prediction inputs at training means and stamp the model-formula provenance.

Usage::

    python scripts/build_demo_workbook.py                     # build + launch
    python scripts/build_demo_workbook.py --verify --no-launch  # build, check, exit
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from dataclasses import replace
from pathlib import Path

import xlwings as xw

from lambda_catalog.analyze_regression_spec import (
    RegressionSpecCase,
    _life_spec,
    _spec_var,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
    build_regression_spec_cases,
    calculate_regression_spec_case,
)
from lambda_catalog.build_common import (
    RUN_LOG_DIR_NAME,
    RUN_LOG_FILE_SUFFIX,
    _recalculate_and_save,
    _retry_on_open,
    print_name_sync_summary,
    run_log_path,
    tee_run_log,
    warn_if_workbook_open,
)
from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.write_sheet_regression import (
    _C_MODEL_FORMULA,
    _ROW_MODEL_FORMULA,
)
from lambda_catalog.write_sheet_test_model import write_test_model_sheet
from lambda_catalog.workbook_builder import (
    XL_CALCULATION_MANUAL,
    XL_CALCULATION_SEMIAUTOMATIC,
    NameSyncResult,
    sync_workbook_names,
)
from lambda_catalog.workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    raise_excel_access_error,
)
from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    PRODUCTION_LOTS,
    load_csv_rows,
    write_csv_dataset_sheet,
)
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Presentations" / "Lambda_Library_Regression_Demo.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"


# ── Beat specs ────────────────────────────────────────────────────────────
# Every Life Expectancy beat is built with _life_spec so the 23 spec rows stay
# in the dataset's own column order: Country is the Identifier, Year the
# Sequence axis (Omit), and every column not named below is Omit. Only the
# kwarg's SpecVariable.name matters; the kwarg name itself is ignored (see
# _life_spec's docstring), so the readable keyword is just a label.


def _declare_spec() -> list:
    """Beat 1 — the curated four-driver model both decks headline.

    Identical to the registered ``life_talk_demo`` (L11) case, so the workbook
    opens on the shipped default and the slide-19 coefficient table is live.
    """
    return _life_spec(
        response=_spec_var("Life expectancy", _ROLE_RESPONSE),
        status=_spec_var("Status", _ROLE_PREDICTOR, True, "Categorical"),
        adult_mortality=_spec_var("Adult Mortality", _ROLE_PREDICTOR, True, "Continuous"),
        alcohol=_spec_var("Alcohol", _ROLE_PREDICTOR, True, "Continuous"),
        percentage_expenditure=_spec_var(
            "percentage expenditure", _ROLE_PREDICTOR, True, "Continuous"
        ),
    )


def _screen_spec() -> list:
    """Beat 2 — add Schooling and BMI for a richer screen panel."""
    return _life_spec(
        response=_spec_var("Life expectancy", _ROLE_RESPONSE),
        status=_spec_var("Status", _ROLE_PREDICTOR, True, "Categorical"),
        adult_mortality=_spec_var("Adult Mortality", _ROLE_PREDICTOR, True, "Continuous"),
        alcohol=_spec_var("Alcohol", _ROLE_PREDICTOR, True, "Continuous"),
        percentage_expenditure=_spec_var(
            "percentage expenditure", _ROLE_PREDICTOR, True, "Continuous"
        ),
        schooling=_spec_var("Schooling", _ROLE_PREDICTOR, True, "Continuous"),
        bmi=_spec_var("BMI", _ROLE_PREDICTOR, True, "Continuous"),
    )


def _trim_spec() -> list:
    """Beat 3 — add under-five deaths, GDP, thinness 5-9 (collinear; high VIF).

    The presenter toggles ``under-five deaths`` off live to show collinearity
    resolving, arriving at the diagnosed model on the next tab.
    """
    return _life_spec(
        response=_spec_var("Life expectancy", _ROLE_RESPONSE),
        status=_spec_var("Status", _ROLE_PREDICTOR, True, "Categorical"),
        adult_mortality=_spec_var("Adult Mortality", _ROLE_PREDICTOR, True, "Continuous"),
        alcohol=_spec_var("Alcohol", _ROLE_PREDICTOR, True, "Continuous"),
        percentage_expenditure=_spec_var(
            "percentage expenditure", _ROLE_PREDICTOR, True, "Continuous"
        ),
        schooling=_spec_var("Schooling", _ROLE_PREDICTOR, True, "Continuous"),
        bmi=_spec_var("BMI", _ROLE_PREDICTOR, True, "Continuous"),
        under_five_deaths=_spec_var("under-five deaths", _ROLE_PREDICTOR, True, "Continuous"),
        gdp=_spec_var("GDP", _ROLE_PREDICTOR, True, "Continuous"),
        thinness_5_9=_spec_var("thinness 5-9 years", _ROLE_PREDICTOR, True, "Continuous"),
    )


def _diagnose_spec() -> list:
    """Beat 4/5 — the trimmed model: drop under-five deaths, Schooling, BMI.

    The non-collinear set the diagnose and predict beats fit: Adult Mortality,
    Alcohol, percentage expenditure, GDP, thinness 5-9 years, C(Status).
    """
    return _life_spec(
        response=_spec_var("Life expectancy", _ROLE_RESPONSE),
        status=_spec_var("Status", _ROLE_PREDICTOR, True, "Categorical"),
        adult_mortality=_spec_var("Adult Mortality", _ROLE_PREDICTOR, True, "Continuous"),
        alcohol=_spec_var("Alcohol", _ROLE_PREDICTOR, True, "Continuous"),
        percentage_expenditure=_spec_var(
            "percentage expenditure", _ROLE_PREDICTOR, True, "Continuous"
        ),
        gdp=_spec_var("GDP", _ROLE_PREDICTOR, True, "Continuous"),
        thinness_5_9=_spec_var("thinness 5-9 years", _ROLE_PREDICTOR, True, "Continuous"),
    )


# ── Beat registry ──────────────────────────────────────────────────────────
# Each beat is a (sheet_name, RegressionSpecCase, include_charts) triple. The
# case's sheet_name is overwritten with the presentation tab name; beats 1 and
#6 reuse the registered QC cases (so their oracle is the suite's), beats 2–5
# are unregistered snapshots. None of these sheet names go through the QC
# registry's validate_sheet_name (which requires an [MLPG] plan-id prefix), so
# the numbered talk tab names are legal here — they are presentation labels,
# not QC case identities.


def _build_beats() -> list[tuple[str, RegressionSpecCase, bool]]:
    """Return the six (sheet_name, case, include_charts) beats, in tab order."""
    registered = {case.name: case for case in build_regression_spec_cases()}

    # Beat 1 — reuse the registered life_talk_demo (L11) case verbatim.
    beat1 = replace(registered["life_talk_demo"], sheet_name="1 Declare")

    # Beat 6 — reuse the registered production_lots_fixed_effects (P01) case.
    beat6 = replace(registered["production_lots_fixed_effects"], sheet_name="6 Fixed Effects")

    # Beats 2–5 — unregistered Life Expectancy snapshots. plan_id is a label
    # for the provenance stamp only; it carries no QC meaning.
    beat2 = RegressionSpecCase(
        name="demo_screen",
        spec=tuple(_screen_spec()),
        allow_intercept=True,
        plan_id="Beat 2",
        sheet_name="2 Screen",
        source_csv_path=registered["life_talk_demo"].source_csv_path,
        row_loader=registered["life_talk_demo"].row_loader,
        source_table_ref="=LifeExpectancyData[#All]",
        back_transform="Duan",
    )
    beat3 = RegressionSpecCase(
        name="demo_trim",
        spec=tuple(_trim_spec()),
        allow_intercept=True,
        plan_id="Beat 3",
        sheet_name="3 Trim",
        source_csv_path=registered["life_talk_demo"].source_csv_path,
        row_loader=registered["life_talk_demo"].row_loader,
        source_table_ref="=LifeExpectancyData[#All]",
        back_transform="Duan",
    )
    beat4 = RegressionSpecCase(
        name="demo_diagnose",
        spec=tuple(_diagnose_spec()),
        allow_intercept=True,
        plan_id="Beat 4",
        sheet_name="4 Diagnose",
        source_csv_path=registered["life_talk_demo"].source_csv_path,
        row_loader=registered["life_talk_demo"].row_loader,
        source_table_ref="=LifeExpectancyData[#All]",
        back_transform="Duan",
    )
    beat5 = RegressionSpecCase(
        name="demo_predict",
        spec=tuple(_diagnose_spec()),
        allow_intercept=True,
        plan_id="Beat 5",
        sheet_name="5 Predict",
        source_csv_path=registered["life_talk_demo"].source_csv_path,
        row_loader=registered["life_talk_demo"].row_loader,
        source_table_ref="=LifeExpectancyData[#All]",
        back_transform="Duan",
    )

    # Charts on the plot-teaching beats only (Diagnose, Predict, Fixed Effects).
    return [
        ("1 Declare", beat1, False),
        ("2 Screen", beat2, False),
        ("3 Trim", beat3, False),
        ("4 Diagnose", beat4, True),
        ("5 Predict", beat5, True),
        ("6 Fixed Effects", beat6, True),
    ]


def build_demo_workbook(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    *,
    verbose: bool = False,
) -> tuple[NameSyncResult, list[str]]:
    """Write the six beat sheets, sync the catalog names, and recalculate.

    Returns the name-sync result and the ordered list of sheet names built.
    """
    document = load_catalog_document(definitions_path)
    closures = document.functions_for_sheet("Regression")
    beats = _build_beats()

    # Compute every oracle BEFORE opening Excel. A spec error should fail in
    # seconds, not after a multi-minute build has already written sheets.
    if verbose:
        print("Oracles: computing", flush=True)
    expectations = [calculate_regression_spec_case(case) for _, case, _ in beats]
    if verbose:
        print("Oracles: done", flush=True)

    headers = {
        config.name: load_csv_rows(config.default_csv_path, config)
        for config in (LIFE_EXPECTANCY, PRODUCTION_LOTS)
    }

    workbook_path = workbook_path.resolve()
    built: list[str] = []
    start = time.monotonic()

    def _run_write() -> None:
        # The first session writes every sheet under XL_CALCULATION_MANUAL so
        # the spec-block writes don't trigger a recalc mid-build, then saves.
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            # Always from scratch: a stale beat sheet from a previous build
            # would carry an old spec and present the wrong model.
            workbook = app.books.add()
            try:
                app.api.Calculation = XL_CALCULATION_MANUAL
                for sheet in list(workbook.sheets)[1:]:
                    sheet.delete()

                # Data sheets FIRST. A beat's spec block binds Source_Table to
                # a structured reference (``=LifeExpectancyData[#All]``), and
                # Excel validates that RefersTo at Names.Add time — so the
                # ListObject must already exist, or the name is rejected as a
                # broken formula. The default sheet becomes the Life Expectancy
                # data sheet (reused by write_csv_dataset_sheet); Production Lots
                # is added after it. Auto MPG is omitted: unused by any beat,
                # and keeping the workbook lean matters on a low-compute host.
                if verbose:
                    print("Data sheets", flush=True)
                workbook.sheets[0].name = LIFE_EXPECTANCY.sheet_name
                for config in (LIFE_EXPECTANCY, PRODUCTION_LOTS):
                    config_headers, config_rows = headers[config.name]
                    write_csv_dataset_sheet(workbook, config_headers, config_rows, config)

                if verbose:
                    print("Sheets: writing 6 beats", flush=True)
                for index, (sheet_name, _case, _charts) in enumerate(beats, start=1):
                    if verbose:
                        print(f"  {index}/6 {sheet_name}", end="", flush=True)
                    t0 = time.monotonic()
                    write_test_model_sheet(
                        workbook,
                        expectations[index - 1],
                        document.regression_sheet_notes,
                        closures,
                        include_charts=beats[index - 1][2],
                    )
                    built.append(sheet_name)
                    if verbose:
                        print(f" {time.monotonic() - t0:5.1f}s", flush=True)

                if verbose:
                    print("Catalog sheet", flush=True)
                # LAMBDA_functions appended after the beats. Tab order is now
                # [Life Expectancy Data, Production Lots, 6 beats, LAMBDA_functions];
                # the data sheets are moved to the end below so the presenter
                # opens on "1 Declare".
                write_catalog_sheet(workbook, document.functions)
                for data_name in (LIFE_EXPECTANCY.sheet_name, PRODUCTION_LOTS.sheet_name):
                    workbook.sheets[data_name].api.Move(
                        After=workbook.sheets[-1].api
                    )

                if verbose:
                    print("Save", flush=True)
                app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
                workbook.save(str(workbook_path))
            finally:
                workbook.close()

    # Checked before the write: this phase only touches workbook_path at the
    # closing save, so a lock surfaces after every beat sheet has been written.
    warn_if_workbook_open(
        workbook_path, action_label=f"{workbook_path.name} is open in Excel"
    )

    try:
        _retry_on_open(
            f"{workbook_path.name} is open in Excel",
            _run_write,
            retry_rpc=True,
        )
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)

    write_elapsed = time.monotonic() - start

    if verbose:
        print("Sync names", flush=True)
    sync_start = time.monotonic()
    result = sync_workbook_names(workbook_path, document.workbook_functions)
    sync_elapsed = time.monotonic() - sync_start

    # The rebuild is not optional: the name sync rewrites the catalog's
    # workbook-scoped LAMBDAs, and a per-sheet Calculate() does not rebuild
    # the dependency tree behind them, so every engine value would read nan
    # without a full CalculateFullRebuild. The shared helper does exactly that
    # and saves in full Automatic, so the workbook opens to computed values.
    if verbose:
        print("CalculateFullRebuild", flush=True)
    rebuild_start = time.monotonic()
    _recalculate_and_save(workbook_path)
    if verbose:
        print(
            f"Timing: write sheets  {write_elapsed:.1f}s\n"
            f"Timing: sync names    {sync_elapsed:.1f}s\n"
            f"Timing: rebuild       {time.monotonic() - rebuild_start:.1f}s",
            flush=True,
        )

    return result, built


def _run_verify(workbook_path: Path, beats: list[tuple[str, object, bool]]) -> int:
    """Structural verification: sheets present, charts on the right beats.

    The demo's beats are presentation snapshots, not QC cases, so the QC
    verifier (which looks sheets up by registered case name) cannot read them.
    This check confirms what the talk depends on instead: every beat sheet
    exists, the catalog closures resolved (the Model Formula readout is a real
    formula, not ``#NAME?``), and charts landed on the three plot-teaching
    beats and nowhere else. The oracle prefill already guarantees the right
    model is written; this confirms it survived the COM write and recalc.
    """
    beat_names = [name for name, _, _ in beats]
    charted = {name for name, _, charts in beats if charts}
    plain = {name for name, _, charts in beats if not charts}
    expected_sheets = set(beat_names) | {"LAMBDA_functions", "Life Expectancy Data", "Production Lots"}

    failures: list[str] = []
    try:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            workbook = app.books.open(str(workbook_path))
            try:
                present = {sheet.name for sheet in workbook.sheets}
                missing = expected_sheets - present
                if missing:
                    failures.append(f"Missing sheets: {sorted(missing)}")

                for name in beat_names:
                    if name not in present:
                        continue  # already reported above
                    sheet = workbook.sheets[name]
                    chart_count = sheet.api.ChartObjects().Count
                    if name in charted and chart_count == 0:
                        failures.append(f"{name}: expected charts (plot beat), found 0")
                    elif name in plain and chart_count > 0:
                        failures.append(
                            f"{name}: expected no charts (spec beat), found {chart_count}"
                        )
                    # The Model Formula readout must resolve to a real string,
                    # not an error — a #NAME? here means the catalog closures
                    # did not install sheet-scoped on this beat.
                    formula_cell = sheet.range(
                        (_ROW_MODEL_FORMULA, _C_MODEL_FORMULA)
                    ).value
                    if not formula_cell or str(formula_cell).startswith("#"):
                        failures.append(
                            f"{name}: Model Formula readout did not resolve "
                            f"(got {formula_cell!r})"
                        )
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "verify", exc)

    if not failures:
        print("Verify: passed")
        return 0
    for message in failures:
        print(f"ERROR {message}")
    print(f"Verify: FAILED with {len(failures)} problem(s).")
    return 1


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the demo workbook build."""
    parser = argparse.ArgumentParser(
        description="Build the presentation Regression demo workbook (one sheet per beat)."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK_PATH,
        help="Path to the demo workbook to create.",
    )
    parser.add_argument(
        "--definitions",
        type=Path,
        default=DEFAULT_DEFINITIONS_PATH,
        help="Path to the JSON file containing lambda definitions.",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="After building, check sheet structure and chart placement.",
    )
    parser.add_argument(
        "--no-launch",
        action="store_true",
        help="Do not open the workbook in Excel when the build finishes.",
    )
    parser.add_argument(
        "--log",
        type=Path,
        default=None,
        help="Where to archive this run's transcript. Defaults to "
        f"'{RUN_LOG_DIR_NAME}/<script> <flags>{RUN_LOG_FILE_SUFFIX}'.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Name every sheet as it is written, with timings.",
    )
    return parser.parse_args()


def main() -> None:
    """Build the demo workbook, optionally verifying it.

    The whole run is teed into ``excel-only-runs/`` so a failure is a file
    somebody can hand over, not a terminal scrollback. The script needs Excel,
    so the transcript only exists on the Windows machine that ran it.
    """
    args = parse_args()
    log_path = (
        args.log
        if args.log is not None
        else run_log_path(ROOT_DIR, Path(__file__).name, sys.argv[1:])
    )
    command = f"python {Path(__file__).name} " + " ".join(sys.argv[1:])

    with tee_run_log(log_path, command):
        total_start = time.monotonic()

        result, built = build_demo_workbook(
            workbook_path=args.workbook,
            definitions_path=args.definitions,
            verbose=args.verbose,
        )

        print(f"Workbook: {args.workbook.resolve()}")
        print(f"Sheets built: {len(built)} ({', '.join(built)})")
        print_name_sync_summary(result)

        exit_code = 0
        if args.verify:
            beats = _build_beats()
            exit_code = _run_verify(args.workbook.resolve(), beats)

        print(f"Timing: total         {time.monotonic() - total_start:.1f}s")
        print(f"Log: {log_path}")

    if not args.no_launch and exit_code == 0:
        subprocess.Popen(["cmd", "/c", "start", "", str(args.workbook.resolve())])

    sys.exit(exit_code)


if __name__ == "__main__":
    main()