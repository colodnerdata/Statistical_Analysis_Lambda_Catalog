"""Tests for workbook inspection value comparisons."""

from lambda_catalog.inspection_compare import compare_values


def test_compare_values_treats_one_missing_value_as_deviation():
    assert compare_values(1.0, None) == (None, 0)
    assert compare_values(None, 1.0) == (None, 0)


def test_compare_values_accepts_two_missing_values():
    assert compare_values(None, None) == (None, None)
