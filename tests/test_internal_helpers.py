"""Tests for private helpers in analyze_life_expectancy."""
# pylint: disable=missing-class-docstring,missing-function-docstring
from __future__ import annotations

import unittest

import numpy as np

from lambda_catalog.analyze_life_expectancy import (
    FEATURE_COLUMNS,
    TARGET_COLUMN,
    _build_training_arrays,
    _normalize_header,
    _parse_float,
    _predict_single_row,
    _validate_required_headers,
)


class ParseFloatTests(unittest.TestCase):
    def test_valid_integer_string(self) -> None:
        self.assertEqual(_parse_float("42"), 42.0)

    def test_valid_float_string(self) -> None:
        result = _parse_float("3.14")
        assert result is not None
        self.assertAlmostEqual(result, 3.14)

    def test_none_input_returns_none(self) -> None:
        self.assertIsNone(_parse_float(None))

    def test_empty_string_returns_none(self) -> None:
        self.assertIsNone(_parse_float(""))

    def test_whitespace_only_returns_none(self) -> None:
        self.assertIsNone(_parse_float("   "))

    def test_non_numeric_string_returns_none(self) -> None:
        self.assertIsNone(_parse_float("not_a_number"))

    def test_whitespace_padded_float(self) -> None:
        result = _parse_float("  2.5  ")
        assert result is not None
        self.assertAlmostEqual(result, 2.5)

    def test_negative_value(self) -> None:
        result = _parse_float("-7.0")
        assert result is not None
        self.assertAlmostEqual(result, -7.0)

    def test_scientific_notation(self) -> None:
        result = _parse_float("1.5e3")
        assert result is not None
        self.assertAlmostEqual(result, 1500.0)


class NormalizeHeaderTests(unittest.TestCase):
    def test_strips_leading_trailing_whitespace(self) -> None:
        self.assertEqual(_normalize_header("  Life expectancy  "), "Life expectancy")

    def test_collapses_internal_whitespace(self) -> None:
        self.assertEqual(_normalize_header("Adult  Mortality"), "Adult Mortality")

    def test_tab_treated_as_whitespace(self) -> None:
        self.assertEqual(_normalize_header("Adult\tMortality"), "Adult Mortality")

    def test_newline_treated_as_whitespace(self) -> None:
        self.assertEqual(_normalize_header("Adult\nMortality"), "Adult Mortality")

    def test_no_change_for_clean_single_word(self) -> None:
        self.assertEqual(_normalize_header("BMI"), "BMI")

    def test_multiple_internal_spaces_collapsed_to_one(self) -> None:
        self.assertEqual(_normalize_header("a   b   c"), "a b c")


class ValidateRequiredHeadersTests(unittest.TestCase):
    def test_raises_when_target_column_missing(self) -> None:
        with self.assertRaises(ValueError) as ctx:
            _validate_required_headers(FEATURE_COLUMNS[:3])
        self.assertIn("Life expectancy", str(ctx.exception))

    def test_raises_when_a_feature_column_is_missing(self) -> None:
        # Headers contain target + only first feature; validate against first two features.
        headers = [TARGET_COLUMN, FEATURE_COLUMNS[0]]
        with self.assertRaises(ValueError):
            _validate_required_headers(headers, feature_columns=FEATURE_COLUMNS[:2])

    def test_passes_when_all_required_headers_present(self) -> None:
        headers = [TARGET_COLUMN, *FEATURE_COLUMNS[:3]]
        _validate_required_headers(headers, feature_columns=FEATURE_COLUMNS[:3])

    def test_normalizes_header_whitespace_before_comparison(self) -> None:
        # Header with extra spaces should still match after normalization.
        headers = ["Life  expectancy", FEATURE_COLUMNS[0]]
        _validate_required_headers(headers, feature_columns=FEATURE_COLUMNS[:1])


class PredictSingleRowTests(unittest.TestCase):
    def test_returns_none_for_none_features(self) -> None:
        coefficients = np.array([1.0, 2.0, 3.0])
        self.assertIsNone(_predict_single_row(coefficients, None, True))

    def test_simple_dot_product_no_intercept(self) -> None:
        coefficients = np.array([2.0, 3.0])
        result = _predict_single_row(coefficients, [1.0, 1.0], False)
        assert result is not None
        self.assertAlmostEqual(result, 5.0)

    def test_dot_product_with_intercept(self) -> None:
        # coef = [intercept=1.0, slope=2.0], features = [3.0]
        # prediction = 1.0*1 + 2.0*3 = 7.0
        coefficients = np.array([1.0, 2.0])
        result = _predict_single_row(coefficients, [3.0], True)
        assert result is not None
        self.assertAlmostEqual(result, 7.0)

    def test_zero_features_give_intercept_only(self) -> None:
        coefficients = np.array([5.0, 2.0, 3.0])
        result = _predict_single_row(coefficients, [0.0, 0.0], True)
        assert result is not None
        self.assertAlmostEqual(result, 5.0)


class BuildTrainingArraysTests(unittest.TestCase):
    def _rows(
        self,
        targets: list,
        features: list,
    ) -> list[dict[str, str]]:
        rows = []
        for y, x in zip(targets, features):
            row: dict[str, str] = {
                TARGET_COLUMN: "" if y is None else str(y),
                FEATURE_COLUMNS[0]: "" if x is None else str(x),
            }
            for col in FEATURE_COLUMNS[1:]:
                row[col] = "1.0"
            rows.append(row)
        return rows

    def test_row_count_matches_valid_targets(self) -> None:
        normalized_rows = self._rows([1.0, None, 3.0], [1.0, 1.0, 1.0])
        _, y_train, _, _ = _build_training_arrays(
            normalized_rows,
            include_intercept=True,
            feature_columns=FEATURE_COLUMNS[:1],
        )
        self.assertEqual(len(y_train), 2)

    def test_intercept_column_is_all_ones_when_requested(self) -> None:
        normalized_rows = self._rows([5.0, 7.0, 9.0], [2.0, 3.0, 4.0])
        x_train, _, _, _ = _build_training_arrays(
            normalized_rows,
            include_intercept=True,
            feature_columns=FEATURE_COLUMNS[:1],
        )
        self.assertTrue(all(x_train[:, 0] == 1.0))

    def test_no_intercept_column_when_flag_is_false(self) -> None:
        normalized_rows = self._rows([5.0, 7.0, 9.0], [2.0, 3.0, 4.0])
        x_train, _, _, _ = _build_training_arrays(
            normalized_rows,
            include_intercept=False,
            feature_columns=FEATURE_COLUMNS[:1],
        )
        self.assertEqual(x_train.shape[1], 1)

    def test_rows_with_missing_feature_excluded(self) -> None:
        normalized_rows = self._rows([1.0, 2.0, 3.0], [1.0, None, 3.0])
        _, y_train, _, _ = _build_training_arrays(
            normalized_rows,
            include_intercept=True,
            feature_columns=FEATURE_COLUMNS[:1],
        )
        self.assertEqual(len(y_train), 2)

    def test_raises_when_no_usable_rows(self) -> None:
        normalized_rows = self._rows([None, None], [None, None])
        with self.assertRaises(ValueError):
            _build_training_arrays(
                normalized_rows,
                include_intercept=True,
                feature_columns=FEATURE_COLUMNS[:1],
            )

    def test_parsed_features_none_for_excluded_rows(self) -> None:
        normalized_rows = self._rows([1.0, 2.0], [1.0, None])
        _, _, parsed_features, _ = _build_training_arrays(
            normalized_rows,
            include_intercept=True,
            feature_columns=FEATURE_COLUMNS[:1],
        )
        self.assertIsNone(parsed_features[1])
        self.assertIsNotNone(parsed_features[0])


if __name__ == "__main__":
    unittest.main()
