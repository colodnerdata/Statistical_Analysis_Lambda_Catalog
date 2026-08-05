"""Tests for the one-sheet-per-test-model framework.

Three things this covers, none of which needs Excel:

1. The sheet-naming contract — Excel's limits, uniqueness, and the
   ``<PlanID> <Concept>`` shape that ties a tab back to
   docs/MODEL_TESTING_ASSETS.md.
2. Coverage in both directions — every case has a sheet, every sheet has a
   case, and no plan ID is claimed twice.
3. The writer's per-sheet parameterization — a generated sheet must point its
   chart-label formulas at itself and skip charts, and its spec block must
   size itself from the source table rather than from a ListObject built at
   the wrong width.
"""
# pylint: disable=missing-function-docstring
from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
import xlwings as xw

from lambda_catalog.analyze_regression_guard_states import build_guard_state_cases
from lambda_catalog.analyze_regression_spec import build_regression_spec_cases
from tests.script_loader import load_script_module
from lambda_catalog.test_model_sheets import (
    ILLEGAL_SHEET_NAME_CHARS,
    MAX_SHEET_NAME_LENGTH,
    SheetNameError,
    assert_sheet_names_unique,
    plan_id_of,
    validate_sheet_name,
)
from lambda_catalog.write_sheet_regression import (
    _C_CHART_LABEL_NAME,
    _C_CHART_TITLE,
    _ROW_CHART_LABELS,
    _diagnostic_chart_specs,
    _write_chart_label_cells,
)
from lambda_catalog.write_sheet_test_model import profile_key_for
from tests.recording_sheet import RecordingSheet

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.csv"


def _as_xw_sheet(sheet: RecordingSheet) -> xw.Sheet:
    return sheet  # type: ignore[return-value]


def _all_cases() -> list:
    return [*build_regression_spec_cases(), *build_guard_state_cases()]


# ── The naming contract ──────────────────────────────────────────────────


def test_validate_sheet_name_accepts_a_well_formed_name() -> None:
    validate_sheet_name("M05 Log-Log NA Masking")
    validate_sheet_name("G01b Empty Model")


@pytest.mark.parametrize(
    "name, reason",
    [
        ("", "empty"),
        ("M05 " + "x" * 40, "characters"),
        ("M05 Bad/Name", "illegal"),
        ("M05 Bad:Name", "illegal"),
        ("M05 Bad[Name]", "illegal"),
        ("'M05 Quoted", "apostrophe"),
        ("M05 Trailing ", "whitespace"),
        ("History", "reserved"),
        ("Just A Concept", "plan ID"),
        ("m05 lowercase tier", "plan ID"),
    ],
)
def test_validate_sheet_name_rejects_illegal_names(name: str, reason: str) -> None:
    with pytest.raises(SheetNameError) as excinfo:
        validate_sheet_name(name)
    assert reason in str(excinfo.value)


def test_excel_limits_are_the_real_ones() -> None:
    """A wrong constant here is invisible until Sheets.Add raises mid-build."""
    assert MAX_SHEET_NAME_LENGTH == 31
    assert ILLEGAL_SHEET_NAME_CHARS == frozenset("[]:*?/\\")


def test_assert_sheet_names_unique_folds_case() -> None:
    """Excel worksheet-name uniqueness ignores case, so the check must too —
    two cases differing only in case would collide on the second Sheets.Add."""
    assert_sheet_names_unique(["M01 One", "M02 Two"])
    with pytest.raises(SheetNameError, match="Duplicate"):
        assert_sheet_names_unique(["M01 One", "m01 one"])


def test_plan_id_is_recoverable_from_a_sheet_name() -> None:
    assert plan_id_of("M05 Log-Log NA Masking") == "M05"
    assert plan_id_of("M03b All Continuous NoInt") == "M03b"
    assert plan_id_of("G01b Empty Model") == "G01b"


# ── Coverage, both directions ────────────────────────────────────────────


def test_every_case_has_a_legal_sheet_name() -> None:
    for case in _all_cases():
        validate_sheet_name(case.sheet_name)
        assert case.plan_id, case.name
        assert case.sheet_name.startswith(case.plan_id + " "), case.name


def test_sheet_names_are_unique_across_model_and_guard_cases() -> None:
    """The two registries are built independently, so nothing but this test
    stops them claiming the same tab — where the second write would land on
    the first's sheet and one case would silently never be verified."""
    assert_sheet_names_unique([case.sheet_name for case in _all_cases()])


def test_plan_ids_are_unique_across_model_and_guard_cases() -> None:
    plan_ids = [case.plan_id for case in _all_cases()]
    assert len(plan_ids) == len(set(plan_ids)), sorted(
        pid for pid in plan_ids if plan_ids.count(pid) > 1
    )


def test_case_names_are_unique_across_model_and_guard_cases() -> None:
    names = [case.name for case in _all_cases()]
    assert len(names) == len(set(names))


def test_every_registered_dataset_has_a_spec_profile() -> None:
    """A case whose Source_Table has no profile would get its spec DEFAULTS
    from the wrong dataset — which fails as wrong NUMBERS, not as an error."""
    for case in _all_cases():
        assert profile_key_for(case.source_table_ref)


def test_profile_key_for_refuses_an_unregistered_source_table() -> None:
    with pytest.raises(ValueError, match="SPEC_DATASET_PROFILES"):
        profile_key_for("=NoSuchTable[#All]")


def test_heavy_is_exactly_the_two_gated_fittable_cases() -> None:
    """Two fittable cases are gated behind ``--include-heavy``:

    * L08 (173 Fixed Effects groups over 2909 rows) — the sheet build is
      what hurts; the per-row residual / leverage / Cook's path is 173×
      wider than the next case.
    * L05 (Kitchen Sink Profile, k = 19, n = 2117, ~5% missingness on
      every predictor) — the statsmodels OLS reference and Excel's OLS
      implementation diverge in the 5th–6th decimal place on most
      coefficients and residuals. Not a defect; both sides go through a
      QR-with-column-pivoting path on an ill-conditioned Gram matrix and
      produce near-tied numerics. Keeping the case on the regular roster
      with a loosened tolerance would paper over the floor; gating it
      preserves the case as a deliberate showcase for the floor that ships.

    L07 was the third candidate until the live run showed the workbook
    cannot fit a 205-column design at all — it is a guard state now, and
    guard sheets are cheap.
    """
    heavy = {case.name for case in build_regression_spec_cases() if case.heavy}
    assert heavy == {"life_country_fixed_effects", "life_full_profile"}


# ── The writer's per-sheet parameterization ──────────────────────────────


def test_chart_specs_qualify_references_with_the_given_sheet_name() -> None:
    """Chart SERIES formulas live above the sheet layer, so they carry a sheet
    prefix even for worksheet-scoped names. A spec built with the wrong name
    still PARSES — it just points at another sheet's data — so the prefix has
    to track the sheet actually being written."""
    specs = _diagnostic_chart_specs("M05 Log-Log NA Masking")
    references = [spec[2] for spec in specs if spec[2]] + [spec[3] for spec in specs]
    assert references
    for reference in references:
        assert reference.startswith("='M05 Log-Log NA Masking'!"), reference
    assert not any("'Regression'!" in reference for reference in references)


def test_chart_label_cells_follow_the_live_sheet_name() -> None:
    sheet = RecordingSheet(name="M09 Cat x Cat Full Product")
    _write_chart_label_cells(_as_xw_sheet(sheet))

    keys = [spec[0] for spec in _diagnostic_chart_specs(sheet.name)]
    for index, key in enumerate(keys):
        row = _ROW_CHART_LABELS + index
        assert sheet.ranges[((row, _C_CHART_LABEL_NAME),)].state.value == key
    titles = [
        sheet.ranges[((_ROW_CHART_LABELS + i, _C_CHART_TITLE),)].state.formula2
        for i in range(len(keys))
    ]
    assert titles and all(title for title in titles)


def test_chart_specs_still_default_to_the_production_sheet() -> None:
    """Every default is chosen so build_production.py's output is unchanged
    by the parameterization — this is the guard on that."""
    for spec in _diagnostic_chart_specs():
        for reference in (spec[2], spec[3]):
            if reference:
                assert reference.startswith("='Regression'!")


# ── Spec-table threading ─────────────────────────────────────────────────


def test_a_cases_typed_sequence_period_lands_in_spec_column_i() -> None:
    """P01/P02 declare Δ = 1, and the sheet has to receive it.

    Base_Period_Delta() reads the TYPED Sequence Period from column I and
    returns #N/A when the cell is blank — it is the override accessor,
    never a silent 1 — and the BFN panel Durbin-Watson cell passes it as
    its delta. If the builder skips this write, AE12 sits at #N/A while the
    oracle holds a real number, and the only symptom is a QC mismatch in a
    multi-minute Excel run that reads like a broken statistic rather than
    an unwritten input cell. Asserted headlessly for that reason.

    The value must land on the SEQUENCE-FLAGGED row specifically, because
    that is the row Base_Period_Delta's own XMATCH resolves.
    """
    from lambda_catalog.analyze_regression_spec import (
        build_regression_spec_cases,
        calculate_regression_spec_case,
    )
    from lambda_catalog.regression_spec_sheet_io import apply_sequence_period_overrides
    from lambda_catalog.write_sheet_model_construction import (
        _C_SEQUENCE_PERIOD,
        _FIRST_DATA_ROW,
    )
    from lambda_catalog.write_sheet_test_model import _padded

    case = next(
        c for c in build_regression_spec_cases()
        if c.name == "production_lots_fixed_effects"
    )
    assert case.sequence_period == 1.0, "P01 is the case that makes BFN live"

    sheet = RecordingSheet(name=case.sheet_name)
    padded = _padded(calculate_regression_spec_case(case))
    apply_sequence_period_overrides(
        _as_xw_sheet(sheet),
        padded.case.spec,
        {item.name: case.sequence_period for item in case.spec if item.sequence},
    )

    written = {
        offset: sheet.cell(_FIRST_DATA_ROW + offset, _C_SEQUENCE_PERIOD).value
        for offset in range(len(padded.case.spec))
        if sheet.cell(_FIRST_DATA_ROW + offset, _C_SEQUENCE_PERIOD).value is not None
    }
    flagged = [
        offset for offset, item in enumerate(padded.case.spec) if item.sequence
    ]
    assert len(flagged) == 1, "exactly one Sequence axis, or H2 errors"
    assert written == {flagged[0]: 1.0}
    assert padded.case.spec[flagged[0]].name == "Fiscal_Year"


def test_only_cases_that_need_a_period_declare_one() -> None:
    """A typed Sequence Period is a declaration, not boilerplate.

    It exists to make the BFN cell computable, which needs Fixed Effects as
    well as a Sequence axis — so a case declaring a period without FE would
    be typing into a cell no diagnostic reads. Pinned so the field does not
    spread by copy-paste the way the Sequence flag itself once did.
    """
    from lambda_catalog.analyze_regression_spec import build_regression_spec_cases
    from lambda_catalog.write_sheet_model_construction import _ROLE_FIXED_EFFECTS

    for case in build_regression_spec_cases():
        if case.sequence_period is None:
            continue
        assert any(item.sequence for item in case.spec), (
            f"{case.plan_id} types a period with no Sequence axis to attach it to"
        )
        assert any(item.role == _ROLE_FIXED_EFFECTS for item in case.spec), (
            f"{case.plan_id} types a period but has no Fixed Effects, so no "
            "diagnostic reads it"
        )


def test_the_spec_block_creates_no_list_object() -> None:
    """The block used to build a SpecTable ListObject sized to the profile,
    which pinned it to the build-time dataset. It must not come back: a
    ListObject cannot be resized by a formula, and a spill cannot live
    inside one, so its return would silently re-break the retarget."""
    from lambda_catalog.write_sheet_model_construction import _write_spec_block

    sheet = RecordingSheet(name="M05 Log-Log NA Masking")
    _write_spec_block(_as_xw_sheet(sheet))

    assert list(sheet.api.ListObjects.items) == []


def test_spec_bands_are_sized_to_the_live_source_table() -> None:
    """Each band TAKE-trims its column to COLUMNS(Source_Data), which is what
    makes a Source_Table retarget resize the spec block. A band pinned to a
    fixed row count is the bug this replaced."""
    from lambda_catalog.write_sheet_model_construction import _set_sheet_scoped_names

    sheet = RecordingSheet(name="M05 Log-Log NA Masking")
    _set_sheet_scoped_names(_as_xw_sheet(sheet), ())

    spec_names = [
        name
        for name in sheet.api.Names.items
        if name.Name.split("!", 1)[-1].startswith("Spec_")
    ]
    assert spec_names
    for name in spec_names:
        assert "COLUMNS(Source_Data)" in name.RefersTo, name.Name
        assert "'M05 Log-Log NA Masking'!" in name.RefersTo, name.Name
        # No structured reference may survive: that is what bound the band
        # to a fixed-height table.
        assert "[[#Data]," not in name.RefersTo, name.Name


def test_spec_bands_carry_no_retired_names() -> None:
    """Spec_Base_Period_Delta was renamed to Spec_Sequence_Period and its
    $I$4:$I$15989 band is still in the shipped artifact — sync_workbook_names
    only sweeps WORKBOOK-scoped residue, so this one has to be dropped by
    the writer that stopped creating it."""
    from lambda_catalog.write_sheet_model_construction import _set_sheet_scoped_names

    sheet = RecordingSheet(name="Regression")
    _set_sheet_scoped_names(_as_xw_sheet(sheet), ())

    registered = {name.Name.split("!", 1)[-1] for name in sheet.api.Names.items}
    assert "Spec_Base_Period_Delta" not in registered
    assert "Spec_Sequence_Period" in registered


def test_spec_role_band_is_the_dataset_sized_role_column() -> None:
    from lambda_catalog.write_sheet_model_construction import _set_sheet_scoped_names

    sheet = RecordingSheet(name="Regression")
    _set_sheet_scoped_names(_as_xw_sheet(sheet), ())

    role = next(
        name for name in sheet.api.Names.items if name.Name.endswith("Spec_Role")
    )
    assert role.RefersTo == (
        "=TAKE('Regression'!$B$4:$B$16000,MAX(1,COLUMNS(Source_Data)))"
    )


# ── Sheet names with spaces must be quoted in every RefersTo ─────────────


def test_every_refers_to_quotes_a_sheet_name_containing_spaces() -> None:
    """A sheet name with a space is invalid in a formula unless single-quoted.

    This is the bug that stopped `build_test_models.py` on its very first
    sheet: `_setup_local_names` built `=M01 Baseline Categoricals!$AB$12`,
    and Excel rejected the whole `Names.Add` with "There's a problem with
    this formula". Four of that function's seven references were unquoted —
    invisible for the entire life of the project because the only sheet it
    ever wrote was named `Regression`, a single word.

    Every generated test-model sheet has a space in its name by design (the
    `<PlanID> <Concept>` contract), so this asserts the property directly
    rather than trusting the convention.
    """
    from lambda_catalog.write_sheet_model_construction import _set_sheet_scoped_names
    from lambda_catalog.write_sheet_regression import (
        _setup_local_names,
        _write_materialization_zone,
    )

    spaced = "M01 Baseline Categoricals"
    quoted = f"'{spaced}'"

    def _unquoted(sheet) -> list[tuple[str, str]]:
        return [
            (name.Name, name.RefersTo)
            for name in sheet.api.Names.items
            if spaced in name.RefersTo and quoted not in name.RefersTo
        ]

    for writer in (
        lambda s: _setup_local_names(s, ()),
        lambda s: _set_sheet_scoped_names(s, ()),
        lambda s: _write_materialization_zone(s, ()),
    ):
        sheet = RecordingSheet(name=spaced)
        writer(_as_xw_sheet(sheet))
        assert _unquoted(sheet) == []

    # Chart series references live above the sheet layer and carry the same
    # requirement.
    for spec in _diagnostic_chart_specs(spaced):
        for reference in (spec[2], spec[3]):
            if reference and spaced in reference:
                assert quoted in reference, reference


# ── Row constants vs. the writers' actual layout ─────────────────────────


def test_row_constants_match_the_writers_own_layout() -> None:
    """`regression_spec_sheet_io`'s row constants must name the rows the zone
    writers actually use.

    Nothing links the two automatically: the writers state these positions as
    literals inside their own formula loops, so the reader keeps a parallel
    set of constants. That is exactly the shape of coupling this repo's
    layout-constant rule exists to catch — if a zone moved, the reader would
    go on reading the old rows and report a wrong NUMBER rather than an
    error, which is far worse than a crash.

    Each zone writer is run against a RecordingSheet and its labelled cells
    are matched against the constants, so a moved row fails here, headlessly,
    rather than during an Excel-only verify run.
    """
    from lambda_catalog import regression_spec_sheet_io as io
    from lambda_catalog.write_sheet_regression import (
        _C_AA,
        _C_AE,
        _C_AG,
        _C_AJ,
        _write_anova,
        _write_coefficients,
        _write_diagnostics,
        _write_prediction_interval,
        _write_prediction_inputs,
        _write_regression_statistics,
        _write_unit_space_block,
    )

    def _labels(writer, col: int) -> dict[int, str]:
        sheet = RecordingSheet(name="Regression")
        writer(_as_xw_sheet(sheet))
        return {
            key[0][0]: cell.state.value
            for key, cell in sheet.ranges.items()
            if len(key) == 1
            and isinstance(key[0], tuple)
            and key[0][1] == col
            and isinstance(cell.state.value, str)
            and cell.state.value
        }

    statistics = _labels(_write_regression_statistics, _C_AA)
    assert statistics[io.ROW_MULTIPLE_R] == "Multiple R"
    assert statistics[io.ROW_R_SQUARED] == "R Square"
    assert statistics[io.ROW_ADJ_R2] == "Adjusted R Square"
    assert statistics[io.ROW_SE_REG] == "Standard Error"
    assert statistics[io.ROW_OBS] == "Observations"

    # The diagnostics block labels sit one column left of their values, which
    # are the cells the reader pulls from _C_AE.
    diagnostics = _labels(_write_diagnostics, _C_AE - 1)
    assert diagnostics[io.ROW_PRESS] == "PRESS"
    assert diagnostics[io.ROW_PRESS_R2] == "PRESS R²"
    assert diagnostics[io.ROW_MEAN_LEV] == "Mean Leverage"
    assert diagnostics[io.ROW_AIC] == "AIC"
    assert diagnostics[io.ROW_BIC] == "BIC"
    assert diagnostics[io.ROW_AICC] == "AICc"
    assert diagnostics[io.ROW_QQ_CORR] == "QQ Correlation"
    assert diagnostics[io.ROW_DURBIN_WATSON] == "Durbin-Watson"

    anova = _labels(_write_anova, _C_AA)
    assert anova[io.ROW_ANOVA_REG] == "Regression"
    assert anova[io.ROW_ANOVA_RES] == "Residual"
    assert anova[io.ROW_ANOVA_TOT] == "Total"

    unit_space = _labels(_write_unit_space_block, _C_AG)
    assert unit_space[io.ROW_BACK_TRANSFORM] == "Back-Transform"
    assert unit_space[5] == "Smearing Factor"
    assert unit_space[6] == "R Square (Unit)"
    assert unit_space[7] == "Adj R Square (Unit)"
    assert unit_space[8] == "RMSE (Unit)"

    interval = _labels(_write_prediction_interval, _C_AJ)
    assert interval[io.ROW_PI_POINT] == "Point Estimate"
    assert interval[io.ROW_FE_GROUP] == "FE Group"
    assert interval[io.ROW_GROUP_MEAN] == "Group Mean (y)"
    assert interval[io.ROW_GROUP_COUNT] == "Group Count"

    # The coefficient LABELS spill from ROW_COEFF_DATA, two rows under the
    # zone heading and one under its column sub-headers.
    coefficients = _labels(_write_coefficients, _C_AA)
    assert coefficients[io.ROW_COEFF_DATA - 2] == "COEFFICIENTS"

    # Prediction inputs: the reader writes from ROW_PRED_INPUT_FIRST, which
    # must be the row after the band's own sub-header row.
    inputs = _labels(_write_prediction_inputs, _C_AJ)
    assert inputs[io.ROW_PRED_INPUT_FIRST - 2] == "Predictor"


# ── Build-driver selection ───────────────────────────────────────────────


def _load_build_test_models():
    """Load scripts/build_test_models.py without relying on the repo root being on sys.path.

    The script moved to ``scripts/`` alongside its siblings in the Chunk 1
    reorganization; the tests below exercise its driver directly.
    """
    return load_script_module("build_test_models")


def test_default_build_excludes_heavy_cases_and_include_heavy_adds_them() -> None:
    build_test_models = _load_build_test_models()

    default_models, default_guards = build_test_models._selected_cases(None, False)
    heavy_models, heavy_guards = build_test_models._selected_cases(None, True)

    assert default_guards == heavy_guards
    assert len(heavy_models) == len(default_models) + 2
    assert not any(case.heavy for case in default_models)


def test_case_filter_matches_plan_id_or_case_name_and_overrides_heavy() -> None:
    build_test_models = _load_build_test_models()

    models, guards = build_test_models._selected_cases({"M09", "G10"}, False)
    assert [case.plan_id for case in models] == ["M09"]
    assert [case.plan_id for case in guards] == ["G10"]

    # By case name, and a heavy case named explicitly is built anyway.
    models, _ = build_test_models._selected_cases(
        {"life_country_fixed_effects"}, False
    )
    assert [case.plan_id for case in models] == ["L08"]


def test_unmatched_case_filter_is_an_error_not_an_empty_build() -> None:
    build_test_models = _load_build_test_models()

    with pytest.raises(ValueError, match="No test-model case matches"):
        build_test_models._selected_cases({"M09", "nonexistent"}, False)


# ── Run-log archiving and verbose progress ───────────────────────────────


def test_run_log_path_names_the_file_after_the_script_and_its_flags() -> None:
    """The convention the first hand-uploaded logs used, now automatic: a
    directory listing reads as a list of what was actually run, and two
    different invocations cannot overwrite each other's evidence."""
    from lambda_catalog.build_common import RUN_LOG_DIR_NAME, run_log_path

    root = Path("/tmp/project")
    assert run_log_path(root, "build_test_models.py", ["--verify", "--no-launch"]) == (
        root / RUN_LOG_DIR_NAME / "build_test_models verify no launch.log"
    )
    assert run_log_path(root, "build_test_models.py", []) == (
        root / RUN_LOG_DIR_NAME / "build_test_models.log"
    )
    # Values attached to a flag are not part of the name — otherwise a path
    # would end up embedded in a filename.
    assert run_log_path(root, "build_test_models.py", ["--cases", "M09,G10"]) == (
        root / RUN_LOG_DIR_NAME / "build_test_models cases.log"
    )


def test_tee_run_log_captures_stdout_stderr_and_the_traceback(tmp_path) -> None:
    """The failure most likely to need archiving is a com_error traceback,
    which never touches stdout. The first hand-captured log of a failed run
    only contained its traceback because a human was copying the terminal."""
    from lambda_catalog.build_common import tee_run_log

    log_path = tmp_path / "logs" / "run.txt"
    with pytest.raises(RuntimeError, match="boom"):
        with tee_run_log(log_path, "python scripts/build_test_models.py --verify"):
            print("progress line")
            print("a warning", file=sys.stderr)
            raise RuntimeError("boom")

    written = log_path.read_text(encoding="utf-8")
    assert "python scripts/build_test_models.py --verify" in written
    assert "progress line" in written
    assert "a warning" in written
    assert "RuntimeError: boom" in written
    assert "Traceback" in written
    # Streams must be restored even though the body raised.
    assert sys.stdout is not None and not hasattr(sys.stdout, "_secondary")


def test_progress_names_the_sheet_before_doing_the_work(capsys) -> None:
    """A summary printed after each sheet says nothing about the one that
    hung. The name has to be on screen, flushed, before the work starts."""
    build_test_models = _load_build_test_models()

    progress = build_test_models._Progress(enabled=True, run_start=0.0)
    case = SimpleNamespace(plan_id="M09", sheet_name="M09 Cat x Cat Full Product")

    with progress.sheet(3, 46, case):
        # Mid-work: the line is already out, without its duration.
        partial = capsys.readouterr().out
        assert "3/46" in partial
        assert "M09 Cat x Cat Full Product" in partial
        assert "s\n" not in partial
    assert capsys.readouterr().out.rstrip().endswith("s")


def test_progress_marks_the_failing_sheet(capsys) -> None:
    build_test_models = _load_build_test_models()

    progress = build_test_models._Progress(enabled=True, run_start=0.0)
    case = SimpleNamespace(plan_id="L08", sheet_name="L08 High Cardinality FE")

    with pytest.raises(ValueError):
        with progress.sheet(1, 1, case):
            raise ValueError("com error")
    assert "FAILED" in capsys.readouterr().out


def test_progress_phases_print_even_without_verbose(capsys) -> None:
    """Phase checkpoints are the minimum that makes a hang attributable, so
    they are not gated on --verbose; only the per-sheet detail is."""
    build_test_models = _load_build_test_models()

    progress = build_test_models._Progress(enabled=False, run_start=0.0)
    progress.phase("Sync names")
    progress.detail("suppressed")
    output = capsys.readouterr().out
    assert "Sync names" in output
    assert "suppressed" not in output


# ── One spec row per source-table column ─────────────────────────────────


def test_every_case_spec_covers_its_whole_source_table() -> None:
    """The invariant that 74,065 mismatches came from violating.

    Every constructor on the sheet opens `n_c = COLUMNS(Source_Data)` and then
    indexes the `Spec_*` bands at `1..n_c`. A spec block one row short of the
    data therefore makes `INDEX(rl, n_c)` run off the end: the row mask
    errors, and every engine cell downstream reads as an error rather than a
    number. Nothing about that is loud — the build succeeds, the sheet looks
    right, and the failure surfaces only as a verify run where an entire case
    reads `None`.

    It bit because `Is_USA` was added to `MileageData` for M15's benefit,
    widening the table to 13 columns while every other Auto MPG case still
    wrote a 12-row spec. M15 was the one Auto MPG sheet that worked.
    """
    from lambda_catalog.write_sheet_test_model import (
        effective_variables,
        pad_spec_to_source_table,
        profile_key_for,
    )

    for case in _all_cases():
        key = profile_key_for(case.source_table_ref)
        padded = pad_spec_to_source_table(case.spec, key)
        assert len(padded) == len(effective_variables(key)), case.name
        # Positional: spec row i describes source column i, so padding must
        # append in the table's own order, never interleave.
        assert [item.name for item in padded] == list(effective_variables(key)), (
            case.name
        )


def test_padding_is_omit_so_it_cannot_change_the_model() -> None:
    """Omit contributes no design column and imposes no mask condition, which
    is why the Python oracle can stay ignorant of the fixture columns."""
    from lambda_catalog.write_sheet_test_model import pad_spec_to_source_table

    case = next(
        c for c in build_regression_spec_cases() if c.name == "default_t0_intercept"
    )
    padded = pad_spec_to_source_table(case.spec, "auto_mpg")
    added = padded[len(case.spec):]
    assert [item.name for item in added] == ["Is_USA"]
    assert all(item.role == "Omit" and not item.include for item in added)
    # The declared rows are untouched.
    assert padded[: len(case.spec)] == tuple(case.spec)


def test_a_spec_naming_a_column_the_table_lacks_is_rejected() -> None:
    """Padding can only ADD rows, so a spec naming something not in the table
    stays too long — caught here rather than as wrong numbers on a sheet."""
    from lambda_catalog.analyze_model_construction import SpecVariable
    from lambda_catalog.write_sheet_test_model import pad_spec_to_source_table

    case = next(
        c for c in build_regression_spec_cases() if c.name == "default_t0_intercept"
    )
    bogus = (*case.spec, SpecVariable("Not_A_Column", "Omit", False, "Continuous"))
    with pytest.raises(ValueError, match="Not_A_Column"):
        pad_spec_to_source_table(bogus, "auto_mpg")


def test_fixture_columns_are_declared_once_for_both_writers() -> None:
    """The data-sheet writer and the spec block must read ONE list. Two copies
    is how the table grew a column the spec block did not know about."""
    build_test_models = _load_build_test_models()
    from lambda_catalog.write_sheet_test_model import (
        FIXTURE_COLUMNS,
        effective_variables,
    )

    assert set(FIXTURE_COLUMNS) <= set(build_test_models._CONFIG_BY_PROFILE_KEY)
    for profile_key, columns in FIXTURE_COLUMNS.items():
        for name in columns:
            assert name in effective_variables(profile_key)


# ── Provenance must not touch the Model Context block ────────────────────


def test_provenance_leaves_the_fit_context_block_intact() -> None:
    """The bug behind the #VALUE! wave.

    Provenance was written at rows 1-2 of the Model Context columns. Row 1 is
    that block's heading and row 2 is the FIRST of the four cells
    `Fit_Context()` reads as a fixed range — so `Allow_Intercept` became the
    string "M04 — continuous_subset_intercept", and every one of the ~30
    engine call sites taking `Fit_Context()` returned #VALUE!: Multiple R,
    R Square, Adjusted R Square, the ANOVA block, SS Total, Beta Weights,
    PRESS R², the fit-space prediction outputs.

    What made it hard to see is what still worked. The coefficients computed
    fine, so the sheet looked like it had a subtle numerical problem rather
    than two clobbered cells.

    This builds the real materialization zone and then the real provenance
    over the top, and asserts every context cell still holds its formula.
    """
    from lambda_catalog.write_sheet_regression import (
        _C_MODEL_CONTEXT,
        _C_MODEL_CONTEXT_LABEL,
        _MATERIALIZATION_FIRST_ROW,
        _MODEL_CONTEXT_ELEMENTS,
        _C_MODEL_FORMULA,
        _C_MODEL_FORMULA_LABEL,
        _MODEL_CONTEXT_LAST_ROW,
        _ROW_MODEL_CONTEXT_CHECK,
        _ROW_MODEL_FORMULA,
        _write_materialization_zone,
    )
    from lambda_catalog.write_sheet_test_model import (
        _ROW_PROVENANCE_COVERS,
        _ROW_PROVENANCE_ID,
        _write_provenance,
    )

    sheet = RecordingSheet(name="M04 Excluded Candidates")
    _write_materialization_zone(_as_xw_sheet(sheet), ())
    _write_provenance(
        _as_xw_sheet(sheet), "M04", "continuous_subset_intercept", "a note"
    )

    for offset, element in enumerate(_MODEL_CONTEXT_ELEMENTS):
        row = _MATERIALIZATION_FIRST_ROW + offset
        assert sheet.ranges[((row, _C_MODEL_CONTEXT),)].state.formula2 == (
            f"={element.formula}"
        ), element.name
        assert sheet.ranges[((row, _C_MODEL_CONTEXT_LABEL),)].state.value == (
            element.label
        ), element.name

    # And the provenance rows are clear of the block AND its health check.
    for row in (_ROW_PROVENANCE_ID, _ROW_PROVENANCE_COVERS):
        assert row > _ROW_MODEL_CONTEXT_CHECK
        assert not (_MATERIALIZATION_FIRST_ROW <= row <= _MODEL_CONTEXT_LAST_ROW)

    # ...and clear of the Model Formula readout, the other thing the
    # Regression writer leaves in this band. Provenance runs AFTER that
    # writer, so an overlap would not conflict — it would silently replace
    # the readout with the case name, and every case's model-formula
    # comparison would then fail against a caption. The readout is on row 1
    # of the design-matrix zone today, columns away from this block, but the
    # two sets of constants are independent, so assert the cells are disjoint
    # rather than assuming they stay that way.
    provenance_cells = {
        (row, col)
        for row in (_ROW_PROVENANCE_ID, _ROW_PROVENANCE_COVERS)
        for col in (_C_MODEL_CONTEXT_LABEL, _C_MODEL_CONTEXT)
    }
    readout_cells = {
        (_ROW_MODEL_FORMULA, col)
        for col in (_C_MODEL_FORMULA_LABEL, _C_MODEL_FORMULA)
    }
    assert not provenance_cells & readout_cells
    assert sheet.ranges[((_ROW_MODEL_FORMULA, _C_MODEL_FORMULA),)].state.formula2 == (
        "=Model_Formula()"
    )
    assert sheet.ranges[
        ((_ROW_MODEL_FORMULA, _C_MODEL_FORMULA_LABEL),)
    ].state.value == "Model Formula"


def test_provenance_still_lands_on_the_sheet() -> None:
    """Moving it must not silently drop it — the tab still explains itself."""
    from lambda_catalog.write_sheet_regression import (
        _C_MODEL_CONTEXT,
        _C_MODEL_CONTEXT_LABEL,
    )
    from lambda_catalog.write_sheet_test_model import (
        _ROW_PROVENANCE_COVERS,
        _ROW_PROVENANCE_ID,
        _write_provenance,
    )

    sheet = RecordingSheet(name="M04 Excluded Candidates")
    _write_provenance(_as_xw_sheet(sheet), "M04", "continuous_subset_intercept", "why")

    assert sheet.ranges[((_ROW_PROVENANCE_ID, _C_MODEL_CONTEXT_LABEL),)].state.value == (
        "Test Model"
    )
    assert sheet.ranges[((_ROW_PROVENANCE_ID, _C_MODEL_CONTEXT),)].state.value == (
        "M04 — continuous_subset_intercept"
    )
    assert sheet.ranges[((_ROW_PROVENANCE_COVERS, _C_MODEL_CONTEXT),)].state.value == "why"
