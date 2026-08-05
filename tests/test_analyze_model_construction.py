"""Tests for the Model Construction QC analyzer.

Excel-side reads run only in the Excel verifier; these tests pin the pure-Python
side — the default-spec expectations derived from the sample Mileage/Auto
MPG CSV (the human test plan's T0 numbers), the stratified-Filter
degeneracy case the Excel verifier drives (the T8 mechanism from the T0 base
state), the mask/label/level semantics that mirror the sheet's formulas,
and the observed-vs-expected comparison layer's failure messages.
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
CSV_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.csv"

_EXPECTED_T0_NAMES = (
    "Horsepower",
    "Weight",
    *(f"Model Year: {year}" for year in range(71, 83)),
    "Origin: Europe",
    "Origin: US",
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
    assert t0_expected.total_rows == 406
    # Full_Data ships as Omit (not Filter), so the mask is completeness-only on
    # the response and the model's two continuous predictors — 392 rows.
    assert t0_expected.included_rows == 392
    assert t0_expected.k == 16
    # Auto MPG ships NO Sequence axis. It is cross-sectional — one row per
    # car model, no unit repeated across periods — so the shipped T0 spec
    # declares nothing on column H and the serial-correlation / base-period
    # layer self-reports "n/a — requires Sequence". The Life Expectancy and
    # Production Lots profiles are the ones that ship a flag, because those
    # two datasets are real panels.
    assert t0_expected.sequence_flags == 0
    assert t0_expected.response_name == "MPG"
    assert t0_expected.responses_count == 1
    assert t0_expected.first_filtered_label == "chevrolet chevelle malibu"
    assert t0_expected.first_filtered_response == 18
    assert t0_expected.level_counts == {"Model Year": 13, "Origin": 3}
    # References in effect, surfaced even when defaulted: first sorted
    # masked level (Model Year is numeric so lowest wins; Origin is now
    # text-valued so the first alphabetical level — "Asia" — wins).
    assert t0_expected.references_in_use == {"Model Year": 70, "Origin": "Asia"}
    assert t0_expected.degenerate_categoricals == ()
    assert len(rows) == t0_expected.total_rows


def test_t0_constructed_names_are_level_qualified_in_spec_order(t0_expected) -> None:
    # 2 included continuous predictors, then 12 Model Year dummies (reference
    # 70 dropped), then 2 Origin dummies (reference "Asia" dropped) — in
    # spec/table order — 16 names, matching audit k.
    assert t0_expected.constructed_column_names == _EXPECTED_T0_NAMES
    assert len(t0_expected.constructed_column_names) == t0_expected.k


# ---------------------------------------------------------------------------
# The Excel verifier's degenerate-Categorical case (T8 mechanism from T0 state)
# ---------------------------------------------------------------------------

def test_stratifying_filter_degenerates_origin(rows) -> None:
    spec = build_default_spec() + [
        SpecVariable("Is_USA", _ROLE_FILTER, False, "Continuous")
    ]
    mutated = [
        {**row, "Is_USA": 1 if row["Origin"] == "US" else 0}
        for row in rows
    ]
    expected = calculate_model_construction_expectations(spec, mutated)

    # Full_Data ships as Omit, so the only Filter is Is_USA: the mask is
    # completeness-on-the-model's-predictors AND Origin = "US" → 245.
    assert expected.included_rows == 245
    assert expected.degenerate_categoricals == ("Origin",)
    assert expected.level_counts == {"Model Year": 13, "Origin": 1}
    # Inside the stratified mask the only Origin level left IS the default
    # reference — the display still shows it while H=1 flags degeneracy.
    assert expected.references_in_use == {"Model Year": 70, "Origin": "US"}
    # Origin is skipped, not errored: zero contributed columns, everything
    # else still constructed — 2 continuous + 12 Model Year dummies.
    assert expected.k == 14
    assert not any(
        name.startswith("Origin:") for name in expected.constructed_column_names
    )
    assert expected.first_filtered_label == "chevrolet chevelle malibu"


def test_invalid_reference_degenerates_the_predictor(rows) -> None:
    spec = [
        variable
        if variable.name != "Origin"
        else SpecVariable(
            variable.name,
            variable.role,
            variable.include,
            variable.var_type,
            reference=99,
        )
        for variable in build_default_spec()
    ]
    expected = calculate_model_construction_expectations(spec, rows)
    assert "Origin" in expected.degenerate_categoricals
    assert expected.k == 14
    # The display echoes the explicit (invalid) E verbatim; E's red CF is
    # what carries the error signal on the sheet.
    assert expected.references_in_use["Origin"] == 99


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
        if variable.name != "Car Name"
        else SpecVariable(variable.name, _ROLE_OMIT, False, variable.var_type)
        for variable in build_default_spec()
    ]
    expected = calculate_model_construction_expectations(spec, rows)
    # Row 1 (chevrolet chevelle malibu) passes the mask, so the positional
    # label is the first observation number.
    assert expected.first_filtered_label == "Obs. 1"


def test_blank_role_behaves_identically_to_explicit_omit(rows) -> None:
    # No constructor closure and no check here ever compares Role against
    # _ROLE_OMIT by name (see the comment next to _ROLE_OMIT in
    # write_sheet_model_construction.py) — Omit is purely the implicit
    # "none of the active roles" bucket, so a blank Role cell must be
    # indistinguishable from an explicit "Omit" in every respect. This is
    # what makes a freshly-added, not-yet-classified spec row safe: it's
    # inert the instant it exists, whether or not its Role has been set.
    #
    # Exercised across every variable in the default spec, one at a time,
    # so the guarantee holds regardless of which row is left unclassified.
    for target in build_default_spec():
        omit_spec = [
            variable
            if variable.name != target.name
            else SpecVariable(variable.name, _ROLE_OMIT, False, variable.var_type)
            for variable in build_default_spec()
        ]
        blank_spec = [
            variable
            if variable.name != target.name
            else SpecVariable(variable.name, "", False, variable.var_type)
            for variable in build_default_spec()
        ]
        omit_expected = calculate_model_construction_expectations(omit_spec, rows)
        blank_expected = calculate_model_construction_expectations(blank_spec, rows)
        assert blank_expected == omit_expected, target.name


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
            "first_filtered_label": "ford torino",
        }
    )
    failures = compare_observed_to_expected(broken, t0_expected, "default spec")

    assert len(failures) == 3  # audit k, the twin tripwire, first label
    assert all(f.startswith("[Model Construction] [default spec]") for f in failures)
    k_failure = next(f for f in failures if "check='audit k'" in f)
    assert "expected=16" in k_failure and "excel_calc=7.0" in k_failure
    assert any("twin tripwire" in f for f in failures)
    label_failure = next(f for f in failures if "first filtered label" in f)
    assert "'chevrolet chevelle malibu'" in label_failure and "'ford torino'" in label_failure
