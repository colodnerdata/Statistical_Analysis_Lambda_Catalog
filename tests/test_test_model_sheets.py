"""Tests for the one-sheet-per-test-model framework.

Three things this covers, none of which needs Excel:

1. The sheet-naming contract — Excel's limits, uniqueness, and the
   ``<PlanID> <Concept>`` shape that ties a tab back to
   docs/MODEL_TESTING_ASSETS.md.
2. Coverage in both directions — every case has a sheet, every sheet has a
   case, and no plan ID is claimed twice.
3. The writer's per-sheet parameterization — a generated sheet must name its
   own ListObject, point its chart-label formulas at itself, and skip charts.
"""
# pylint: disable=missing-function-docstring
from __future__ import annotations

from pathlib import Path

import pytest
import xlwings as xw

from lambda_catalog.analyze_regression_guard_states import build_guard_state_cases
from lambda_catalog.analyze_regression_spec import build_regression_spec_cases
from lambda_catalog.test_model_sheets import (
    ILLEGAL_SHEET_NAME_CHARS,
    MAX_SHEET_NAME_LENGTH,
    SheetNameError,
    assert_sheet_names_unique,
    plan_id_of,
    spec_table_name,
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


def test_spec_table_names_are_unique_per_sheet() -> None:
    """Excel ListObject names are WORKBOOK-scoped: two sheets both naming
    their table SpecTable is an error from ListObjects.Add, not a rename."""
    table_names = [spec_table_name(case.plan_id) for case in _all_cases()]
    assert len(table_names) == len(set(table_names))
    assert spec_table_name("M05") == "SpecTable_M05"


def test_every_registered_dataset_has_a_spec_profile() -> None:
    """A case whose Source_Table has no profile would get a spec table sized
    for the wrong dataset — which fails as wrong NUMBERS, not as an error."""
    for case in _all_cases():
        assert profile_key_for(case.source_table_ref)


def test_profile_key_for_refuses_an_unregistered_source_table() -> None:
    with pytest.raises(ValueError, match="SPEC_DATASET_PROFILES"):
        profile_key_for("=NoSuchTable[#All]")


def test_heavy_cases_are_exactly_the_two_oversized_life_expectancy_models() -> None:
    heavy = {case.name for case in build_regression_spec_cases() if case.heavy}
    assert heavy == {"life_country_width_guard", "life_country_fixed_effects"}


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


def test_spec_block_names_its_table_and_binds_the_bands_to_it() -> None:
    """The Spec_* band names must reference the SAME table name the block
    created. They are sheet-scoped, so each generated sheet binds its own —
    but only if both sides read one parameter."""
    from lambda_catalog.write_sheet_model_construction import _create_spec_table

    sheet = RecordingSheet(name="M05 Log-Log NA Masking")
    _create_spec_table(_as_xw_sheet(sheet), None, "SpecTable_M05")

    assert [table.Name for table in sheet.api.ListObjects.items] == ["SpecTable_M05"]


def test_spec_scoped_names_reference_the_per_sheet_table() -> None:
    from lambda_catalog.write_sheet_model_construction import _set_sheet_scoped_names

    sheet = RecordingSheet(name="M05 Log-Log NA Masking")
    _set_sheet_scoped_names(
        _as_xw_sheet(sheet), (), spec_table_name="SpecTable_M05"
    )

    spec_names = [
        name
        for name in sheet.api.Names.items
        if name.Name.split("!", 1)[-1].startswith("Spec_")
    ]
    assert spec_names
    for name in spec_names:
        assert "SpecTable_M05[[#Data]," in name.RefersTo, name.Name
        assert "'M05 Log-Log NA Masking'!" in name.RefersTo, name.Name


def test_spec_scoped_names_default_to_the_production_table_name() -> None:
    from lambda_catalog.write_sheet_model_construction import _set_sheet_scoped_names

    sheet = RecordingSheet(name="Regression")
    _set_sheet_scoped_names(_as_xw_sheet(sheet), ())

    role = next(
        name for name in sheet.api.Names.items if name.Name.endswith("Spec_Role")
    )
    assert role.RefersTo == "='Regression'!SpecTable[[#Data],[Role]]"


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
        lambda s: _set_sheet_scoped_names(s, (), spec_table_name="SpecTable_M01"),
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


def test_default_build_excludes_heavy_cases_and_include_heavy_adds_them() -> None:
    import build_test_models

    default_models, default_guards = build_test_models._selected_cases(None, False)
    heavy_models, heavy_guards = build_test_models._selected_cases(None, True)

    assert default_guards == heavy_guards
    assert len(heavy_models) == len(default_models) + 2
    assert not any(case.heavy for case in default_models)


def test_case_filter_matches_plan_id_or_case_name_and_overrides_heavy() -> None:
    import build_test_models

    models, guards = build_test_models._selected_cases({"M09", "G10"}, False)
    assert [case.plan_id for case in models] == ["M09"]
    assert [case.plan_id for case in guards] == ["G10"]

    # By case name, and a heavy case named explicitly is built anyway.
    models, _ = build_test_models._selected_cases(
        {"life_country_width_guard"}, False
    )
    assert [case.plan_id for case in models] == ["L07"]


def test_unmatched_case_filter_is_an_error_not_an_empty_build() -> None:
    import build_test_models

    with pytest.raises(ValueError, match="No test-model case matches"):
        build_test_models._selected_cases({"M09", "nonexistent"}, False)
