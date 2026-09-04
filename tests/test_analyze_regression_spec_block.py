"""Tests for the Regression spec-block QC verifier.

Excel-side reads run only in the Excel verifier; these tests pin the pure-Python
side — the comparison layer's pass/fail behavior and message format, the
sheet coordinates the reads anchor to, and the delegation to the shared
expectation calculator (whose T0/T8 numbers are already pinned in
tests/test_analyze_model_construction.py — not re-pinned here).
"""
# pylint: disable=missing-function-docstring,protected-access
from pathlib import Path
from types import SimpleNamespace

import pytest

from lambda_catalog import analyze_model_construction, analyze_regression_spec_block
from lambda_catalog.analyze_model_construction import (
    build_default_spec,
    calculate_model_construction_expectations,
    load_source_rows,
)
from lambda_catalog.analyze_regression_spec_block import (
    _ROW_COEFF_FIRST,
    _ROW_NAMES_SPILL,
    _ROW_OBSERVATIONS,
    _ROW_RESID_FIRST,
    _ROW_RESPONSE_READOUT,
    RegressionSpecObserved,
    compare_spec_observed_to_expected,
)
from lambda_catalog.workbook_helpers import col_letter
from lambda_catalog.write_spec_block import (
    _C_SEQUENCE,
    _INTERCEPT_ROW,
    SPEC_DATASET_PROFILES,
)
from lambda_catalog.write_sheet_regression import _C_AA, _C_AB, _C_AF, _C_AN, _C_S

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.csv"


@pytest.fixture(scope="module", name="t0_expected")
def _t0_expected():
    rows = load_source_rows(CSV_PATH)
    return calculate_model_construction_expectations(build_default_spec(), rows)


# ---------------------------------------------------------------------------
# Coordinate pins — a future layout shift must break loudly here
# ---------------------------------------------------------------------------

def test_reads_anchor_to_the_documented_regression_cells() -> None:
    # These are the cells read_observed_spec_values actually reads. Pinning
    # the letters here is what makes a layout shift break loudly instead of
    # letting the verifier read a neighbouring column and compare the wrong
    # numbers — so import the SAME constants the module under test uses,
    # never a parallel set that happens to spell the old letters.
    # S4 = TRANSPOSE(Constructed_Column_Names()) spill
    assert (col_letter(_C_S), _ROW_NAMES_SPILL) == ("S", 4)
    # AA22 = coefficient label spill (k+1 rows with the default intercept ON)
    assert (col_letter(_C_AA), _ROW_COEFF_FIRST) == ("AA", 22)
    # AB9 = Observations cell
    assert (col_letter(_C_AB), _ROW_OBSERVATIONS) == ("AB", 9)
    # AF3 = Predicted Variable readout
    assert (col_letter(_C_AF), _ROW_RESPONSE_READOUT) == ("AF", 3)
    # AN4 = FILTER(Row_Labels(), Sample_Include()) spill
    assert (col_letter(_C_AN), _ROW_RESID_FIRST) == ("AN", 4)
    # H2 = Sequence status line (shared spec-block coordinates)
    assert (col_letter(_C_SEQUENCE), _INTERCEPT_ROW) == ("H", 2)


def test_expectations_delegate_to_the_shared_calculator() -> None:
    # One expectation side for both verifiers: the numbers pinned in
    # test_analyze_model_construction.py cover this module too.
    assert (
        analyze_regression_spec_block.calculate_model_construction_expectations
        is analyze_model_construction.calculate_model_construction_expectations
    )
    # build_PROFILE_spec, not build_default_spec: this verifier describes
    # whichever dataset the workbook in front of it targets, and
    # build_default_spec is Auto MPG specifically.
    assert (
        analyze_regression_spec_block.build_profile_spec
        is analyze_model_construction.build_profile_spec
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
                "Origin": 2.0,
            },
        }
    )
    failures = compare_spec_observed_to_expected(broken, t0_expected, "default spec")

    assert len(failures) == 2  # twin tripwire, Origin reference
    assert all(f.startswith("[Regression Spec] [default spec]") for f in failures)
    tripwire = next(f for f in failures if "twin tripwire" in f)
    assert "expected=17" in tripwire and "excel_calc=7" in tripwire
    reference = next(f for f in failures if "Reference In Use" in f)
    assert "expected='Asia'" in reference and "excel_calc=2.0" in reference


# ---------------------------------------------------------------------------
# Dataset resolution — the layer that stops this verifier comparing the
# wrong dataset when the shipped default changes
# ---------------------------------------------------------------------------


class _FakeNames:
    def __init__(self, refers_to: object) -> None:
        self._refers_to = refers_to

    def Item(self, name: str):  # noqa: N802 — mirrors the COM API
        assert name == "Source_Table"
        if self._refers_to is None:
            raise KeyError(name)
        return SimpleNamespace(RefersTo=self._refers_to)


class _FakeSheet:
    """Just enough of an xlwings Sheet to answer 'what is Source_Table?'."""

    def __init__(self, refers_to: object) -> None:
        self.api = SimpleNamespace(Names=_FakeNames(refers_to))


@pytest.mark.parametrize(
    "refers_to, expected_key",
    [
        ("=MileageData[#All]", "auto_mpg"),
        ("=LifeExpectancyData[#All]", "life_expectancy"),
        ("=ProductionLotsData[#All]", "production_lots"),
        # Excel hands the ref back requoted or sheet-qualified depending on
        # how it was written; matching on the table name absorbs that.
        ("='Mileage Data'!MileageData[#All]", "auto_mpg"),
        ("=lifeexpectancydata[#all]", "life_expectancy"),
    ],
)
def test_source_table_resolves_to_the_shipped_dataset(refers_to, expected_key) -> None:
    resolved = analyze_regression_spec_block._resolve_shipped_profile(
        _FakeSheet(refers_to)
    )

    assert resolved is not None
    key, profile = resolved
    assert key == expected_key
    assert profile is SPEC_DATASET_PROFILES[expected_key]


@pytest.mark.parametrize("refers_to", [None, "=SomeOtherTable[#All]", "=$A$1:$C$9"])
def test_unresolvable_source_table_is_a_failure_not_a_default(refers_to) -> None:
    """The whole point of resolving rather than assuming.

    Falling back to a default dataset here is what produced the original bug:
    the verifier compared Auto MPG expectations against a workbook shipping
    Life Expectancy and reported ten mismatches that all meant one thing. An
    unknown target has to say so.
    """
    assert (
        analyze_regression_spec_block._resolve_shipped_profile(_FakeSheet(refers_to))
        is None
    )


def test_every_registered_profile_has_a_loader_and_they_agree() -> None:
    """A profile the spec block can target but this module cannot load would
    resolve to None and report 'no registered profile' — a confusing way to
    say 'somebody added a dataset and not its loader'."""
    assert set(analyze_regression_spec_block._DATASET_LOADERS) == set(
        SPEC_DATASET_PROFILES
    )
    for key, profile in SPEC_DATASET_PROFILES.items():
        config, loader = analyze_regression_spec_block._DATASET_LOADERS[key]
        assert config.table_name in profile.source_table_ref, key
        assert callable(loader), key


@pytest.mark.parametrize("profile_key", sorted(SPEC_DATASET_PROFILES))
def test_every_profile_produces_computable_expectations(profile_key) -> None:
    """Each shipped default must be describable by the expectation side.

    This is the assertion that would have caught the breakage at unit-test
    time instead of on a Windows machine: it fits every profile the
    --regression-dataset flag can select, not just the one that happens to
    ship today.
    """
    config, loader = analyze_regression_spec_block._DATASET_LOADERS[profile_key]
    if not config.default_csv_path.exists():
        pytest.skip(f"{config.default_csv_path.name} not present")

    expected = calculate_model_construction_expectations(
        analyze_model_construction.build_profile_spec(
            SPEC_DATASET_PROFILES[profile_key]
        ),
        loader(config.default_csv_path),
    )

    assert expected.response_name
    assert expected.k >= 1
    assert 0 < expected.included_rows <= expected.total_rows
    assert len(expected.constructed_column_names) == expected.k
