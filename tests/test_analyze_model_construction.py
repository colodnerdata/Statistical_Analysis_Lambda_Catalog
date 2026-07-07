"""Tests for the Model Construction QC analyzer.

Excel-side reads run only in the QC build; these tests pin the pure-Python
side — the default-spec expectations derived from the sample CSV (the human
test plan's T0 numbers), the stratified-Filter degeneracy case the QC build
drives (the T8 mechanism from the T0 base state), the mask/label/level
semantics that mirror the sheet's formulas, and the observed-vs-expected
comparison layer's failure messages.
"""
# pylint: disable=missing-function-docstring,protected-access
from pathlib import Path

import pytest

from lambda_catalog.analyze_model_construction import (
    _EXTRA_FILTER_FORMULA,
    ModelConstructionObserved,
    SpecVariable,
    _is_truthy_filter,
    build_default_spec,
    calculate_model_construction_expectations,
    compare_observed_to_expected,
    load_source_rows,
)
from lambda_catalog.write_sheet_model_construction import (
    _DEFAULT_SEQUENCE_VARIABLES,
    _DEFAULT_SPEC,
    _ROLE_FILTER,
    _ROLE_OMIT,
    _VARIABLES,
)

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"

_EXPECTED_T0_NAMES = (
    *(f"Year: {year}" for year in range(2001, 2016)),
    "Status: Developing",
    "Adult Mortality",
    "GDP",
    "Schooling",
)


@pytest.fixture(scope="module", name="rows")
def _rows() -> list[dict[str, object]]:
    return load_source_rows(CSV_PATH)


@pytest.fixture(scope="module", name="t0_expected")
def _t0_expected(rows):
    return calculate_model_construction_expectations(build_default_spec(), rows)


# ---------------------------------------------------------------------------
# Spec model
# ---------------------------------------------------------------------------

def test_default_spec_mirrors_the_writer_prefill() -> None:
    spec = build_default_spec()
    assert [v.name for v in spec] == _VARIABLES
    for variable in spec:
        if variable.name in _DEFAULT_SPEC:
            role, include, var_type = _DEFAULT_SPEC[variable.name]
            assert (variable.role, variable.include, variable.var_type) == (
                role,
                include,
                var_type,
            )
        else:
            assert (variable.role, variable.include) == ("Predictor (x)", False)
        assert variable.reference == ""
        assert variable.sequence == (variable.name in _DEFAULT_SEQUENCE_VARIABLES)


# ---------------------------------------------------------------------------
# T0 (default spec) expectations — the human test plan's pinned numbers
# ---------------------------------------------------------------------------

def test_t0_expectations_pin_the_csv_derived_values(rows, t0_expected) -> None:
    assert t0_expected.total_rows == 2938
    # Full_Data ships as Omit (not Filter), so the mask is completeness-only on
    # the response and the model's three continuous predictors — 2482 rows, vs
    # 1649 under the old all-features Full_Data filter (which over-filtered).
    assert t0_expected.included_rows == 2482
    assert t0_expected.k == 19
    # Year is the shipped Sequence axis; Population's Omit role leaves k
    # unchanged (Omit contributes no column, same as a not-included row).
    assert t0_expected.sequence_flags == 1
    assert t0_expected.response_name == "Life expectancy"
    assert t0_expected.responses_count == 1
    assert t0_expected.first_filtered_label == "Afghanistan"
    assert t0_expected.first_filtered_response == 65
    assert t0_expected.level_counts == {"Year": 16, "Status": 2}
    # References in effect, surfaced even when defaulted: first sorted
    # masked level (numbers before text, so Year -> 2000).
    assert t0_expected.references_in_use == {"Year": 2000, "Status": "Developed"}
    assert t0_expected.degenerate_categoricals == ()
    assert len(rows) == t0_expected.total_rows


def test_t0_constructed_names_are_level_qualified_in_spec_order(t0_expected) -> None:
    # 15 Year dummies (reference 2000 dropped), the Status dummy (reference
    # Developed dropped), then the three included continuous predictors in
    # table order — 19 names, matching audit k.
    assert t0_expected.constructed_column_names == _EXPECTED_T0_NAMES
    assert len(t0_expected.constructed_column_names) == t0_expected.k


# ---------------------------------------------------------------------------
# The QC build's degenerate-Categorical case (T8 mechanism from T0 state)
# ---------------------------------------------------------------------------

def test_stratifying_filter_degenerates_status(rows) -> None:
    spec = build_default_spec() + [
        SpecVariable("Is_Developing", _ROLE_FILTER, False, "Continuous")
    ]
    mutated = [
        {**row, "Is_Developing": 1 if row["Status"] == "Developing" else 0}
        for row in rows
    ]
    expected = calculate_model_construction_expectations(spec, mutated)

    # Full_Data ships as Omit, so the only Filter is Is_Developing: the mask is
    # completeness-on-the-model's-predictors AND Status = Developing → 2034
    # (was 1407 when Full_Data's all-features filter was also ANDed in).
    assert expected.included_rows == 2034
    assert expected.degenerate_categoricals == ("Status",)
    assert expected.level_counts == {"Year": 16, "Status": 1}
    # Inside the stratified mask the only Status level left IS the default
    # reference — the display still shows it while H=1 flags degeneracy.
    assert expected.references_in_use == {"Year": 2000, "Status": "Developing"}
    # Status is skipped, not errored: zero contributed columns, everything
    # else still constructed — 15 Year dummies + 3 continuous.
    assert expected.k == 18
    assert not any(
        name.startswith("Status:") for name in expected.constructed_column_names
    )
    assert expected.first_filtered_label == "Afghanistan"


def test_invalid_reference_degenerates_the_predictor(rows) -> None:
    spec = [
        variable
        if variable.name != "Status"
        else SpecVariable(
            variable.name,
            variable.role,
            variable.include,
            variable.var_type,
            reference="Nonexistent Level",
        )
        for variable in build_default_spec()
    ]
    expected = calculate_model_construction_expectations(spec, rows)
    assert "Status" in expected.degenerate_categoricals
    assert expected.k == 18
    # The display echoes the explicit (invalid) E verbatim; E's red CF is
    # what carries the error signal on the sheet.
    assert expected.references_in_use["Status"] == "Nonexistent Level"


# ---------------------------------------------------------------------------
# Mask / label semantics mirroring the sheet formulas
# ---------------------------------------------------------------------------

def test_the_injected_filter_formula_is_balanced() -> None:
    assert _EXTRA_FILTER_FORMULA.count("(") == _EXTRA_FILTER_FORMULA.count(")")
    assert _EXTRA_FILTER_FORMULA.count('"') % 2 == 0


def test_filter_truthiness_matches_the_sheet_mask() -> None:
    # N(IFERROR(N(col)=1,FALSE)): TRUE and 1 pass; FALSE/0/blank/text fail.
    assert _is_truthy_filter(True)
    assert _is_truthy_filter(1)
    assert _is_truthy_filter(1.0)
    assert not _is_truthy_filter(False)
    assert not _is_truthy_filter(0)
    assert not _is_truthy_filter("1")
    assert not _is_truthy_filter(None)


def test_row_labels_fall_back_to_positional_without_identifiers(rows) -> None:
    spec = [
        variable
        if variable.name != "Country"
        else SpecVariable(variable.name, _ROLE_OMIT, False, variable.var_type)
        for variable in build_default_spec()
    ]
    expected = calculate_model_construction_expectations(spec, rows)
    # Row 1 (Afghanistan 2015) passes the mask, so the positional label is
    # the first observation number.
    assert expected.first_filtered_label == "Obs. 1"


# ---------------------------------------------------------------------------
# Comparison layer
# ---------------------------------------------------------------------------

def _observed_matching(expected) -> ModelConstructionObserved:
    """An Observed exactly as a correct sheet would read back (floats)."""
    return ModelConstructionObserved(
        audit_k=float(expected.k),
        audit_rows=float(expected.total_rows),
        audit_response=expected.response_name,
        audit_responses=float(expected.responses_count),
        audit_included=float(expected.included_rows),
        # The shipped spec flags Year as the Sequence axis, so the zero-or-one
        # audit count reads 1; a correct sheet reads back this expected count.
        audit_sequence_flags=float(expected.sequence_flags),
        header_strip=expected.constructed_column_names,
        level_cells={
            name: float(count) for name, count in expected.level_counts.items()
        },
        reference_cells={
            name: float(ref) if isinstance(ref, int) else ref
            for name, ref in expected.references_in_use.items()
        },
        row_labels_height=expected.total_rows,
        mask_height=expected.total_rows,
        mask_true_count=expected.included_rows,
        filtered_labels_height=expected.included_rows,
        filtered_response_height=expected.included_rows,
        matrix_labels_height=expected.included_rows,
        matrix_height=expected.included_rows,
        matrix_width=expected.k,
        first_filtered_label=expected.first_filtered_label,
        y_header=f"y: {expected.response_name}",
        first_filtered_response=float(expected.first_filtered_response),
    )


def test_comparison_passes_on_a_matching_sheet(t0_expected) -> None:
    observed = _observed_matching(t0_expected)
    assert compare_observed_to_expected(observed, t0_expected, "default spec") == []


def test_comparison_reports_standard_format_failures(t0_expected) -> None:
    observed = _observed_matching(t0_expected)
    broken = ModelConstructionObserved(
        **{
            **observed.__dict__,
            "audit_k": 7.0,
            "first_filtered_label": "Albania",
        }
    )
    failures = compare_observed_to_expected(broken, t0_expected, "default spec")

    assert len(failures) == 3  # audit k, the twin tripwire, first label
    assert all(f.startswith("[Model Construction] [default spec]") for f in failures)
    k_failure = next(f for f in failures if "check='audit k'" in f)
    assert "expected=19" in k_failure and "excel_calc=7.0" in k_failure
    assert any("twin tripwire" in f for f in failures)
    label_failure = next(f for f in failures if "first filtered label" in f)
    assert "'Afghanistan'" in label_failure and "'Albania'" in label_failure
