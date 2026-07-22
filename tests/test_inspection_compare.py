"""Tests for workbook inspection value comparisons."""

from lambda_catalog.inspection_compare import compare_values


def test_compare_values_treats_one_missing_value_as_deviation():
    assert compare_values(1.0, None) == (None, 0)
    assert compare_values(None, 1.0) == (None, 0)


def test_compare_values_accepts_two_missing_values():
    assert compare_values(None, None) == (None, None)


def test_compare_values_treats_nan_or_inf_expected_as_missing():
    # A leverage-1 row makes the residual denominator 0: the Python oracle
    # computes NaN/Inf there, and Excel's IFERROR(...,NA()) guard reads back
    # as None. Both sides being "undefined" is not a deviation.
    assert compare_values(float("nan"), None) == (None, None)
    assert compare_values(float("inf"), None) == (None, None)
    assert compare_values(float("-inf"), None) == (None, None)


def test_compare_values_nan_expected_against_a_real_number_is_still_a_deviation():
    # If Excel computed a real number while the oracle is undefined, that is
    # a genuine mismatch worth flagging, not a silently-dropped comparison.
    assert compare_values(float("nan"), 1.0) == (None, 0)
