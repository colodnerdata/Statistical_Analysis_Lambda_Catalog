"""Tests for the Regression spec-block QC verifier.

Excel-side reads run only in the QC build; these tests pin the pure-Python
side — the comparison layer's pass/fail behavior and message format, the
sheet coordinates the reads anchor to, and the delegation to the shared
expectation calculator (whose T0/T8 numbers are already pinned in
tests/test_analyze_model_construction.py — not re-pinned here).
"""
# pylint: disable=missing-function-docstring,protected-access
from pathlib import Path

import pytest

from lambda_catalog import analyze_model_construction, analyze_regression_spec_block
from lambda_catalog.analyze_regression_spec_block import (
    RegressionSpecObserved,
    _ROW_COEFF_FIRST,
    _ROW_NAMES_SPILL,
    _ROW_OBSERVATIONS,
    _ROW_RESID_FIRST,
    _ROW_RESPONSE_READOUT,
    compare_spec_observed_to_expected,
)
from lambda_catalog.analyze_model_construction import (
    build_default_spec,
    calculate_model_construction_expectations,
    load_source_rows,
)
from lambda_catalog.workbook_helpers import col_letter
from lambda_catalog.write_sheet_model_construction import (
    _C_SEQUENCE,
    _INTERCEPT_ROW,
)
from lambda_catalog.write_sheet_regression import _C_AH, _C_M, _C_U, _C_V, _C_Z

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "Life Expectancy Data.csv"


@pytest.fixture(scope="module", name="t0_expected")
def _t0_expected():
    rows = load_source_rows(CSV_PATH)
    return calculate_model_construction_expectations(build_default_spec(), rows)


# ---------------------------------------------------------------------------
# Coordinate pins — a future layout shift must break loudly here
# ---------------------------------------------------------------------------

def test_reads_anchor_to_the_documented_regression_cells() -> None:
    # M3 = TRANSPOSE(Constructed_Column_Names()) spill
    assert (col_letter(_C_M), _ROW_NAMES_SPILL) == ("M", 3)
    # U21 = coefficient label spill (k+1 rows with the default intercept ON)
    assert (col_letter(_C_U), _ROW_COEFF_FIRST) == ("U", 21)
    # V8 = Observations cell
    assert (col_letter(_C_V), _ROW_OBSERVATIONS) == ("V", 8)
    # Z2 = Predicted Variable readout
    assert (col_letter(_C_Z), _ROW_RESPONSE_READOUT) == ("Z", 2)
    # AH3 = FILTER(Row_Labels(), Sample_Include()) spill
    assert (col_letter(_C_AH), _ROW_RESID_FIRST) == ("AH", 3)
    # H2 = Sequence status line (shared spec-block coordinates)
    assert (col_letter(_C_SEQUENCE), _INTERCEPT_ROW) == ("H", 2)


def test_expectations_delegate_to_the_shared_calculator() -> None:
    # One expectation side for both verifiers: the T0/T8 numbers pinned in
    # test_analyze_model_construction.py cover this module too.
    assert (
        analyze_regression_spec_block.calculate_model_construction_expectations
        is analyze_model_construction.calculate_model_construction_expectations
    )
    assert (
        analyze_regression_spec_block.build_default_spec
        is analyze_model_construction.build_default_spec
    )


# ---------------------------------------------------------------------------
# Comparison layer
# ---------------------------------------------------------------------------

def _observed_matching(expected) -> RegressionSpecObserved:
    """An Observed exactly as a correct Regression sheet would read back."""
    return RegressionSpecObserved(
        constructed_names=expected.constructed_column_names,
        coeff_label_height=expected.k + 1,
        response_readout=expected.response_name,
        observations_cell=float(expected.included_rows),
        level_cells={
            name: float(count) for name, count in expected.level_counts.items()
        },
        reference_cells={
            name: float(ref) if isinstance(ref, int) else ref
            for name, ref in expected.references_in_use.items()
        },
        resid_labels_height=expected.included_rows,
        first_resid_label=expected.first_filtered_label,
        # A legal spec (zero-or-one flags) renders the status line blank;
        # COM reads an Excel "" back as None, which the comparison
        # blank-normalizes.
        sequence_status=None,
    )


def test_comparison_passes_on_a_matching_sheet(t0_expected) -> None:
    observed = _observed_matching(t0_expected)
    assert (
        compare_spec_observed_to_expected(observed, t0_expected, "default spec")
        == []
    )


def test_comparison_reports_standard_format_failures(t0_expected) -> None:
    observed = _observed_matching(t0_expected)
    broken = RegressionSpecObserved(
        **{
            **observed.__dict__,
            "coeff_label_height": 7,
            "reference_cells": {
                **observed.reference_cells,
                "Status": "Developing",
            },
        }
    )
    failures = compare_spec_observed_to_expected(broken, t0_expected, "default spec")

    assert len(failures) == 2  # twin tripwire, Status reference
    assert all(f.startswith("[Regression Spec] [default spec]") for f in failures)
    tripwire = next(f for f in failures if "twin tripwire" in f)
    assert "expected=20" in tripwire and "excel_calc=7" in tripwire
    reference = next(f for f in failures if "Reference In Use" in f)
    assert "'Developed'" in reference and "'Developing'" in reference
