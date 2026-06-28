"""Validate the Univariate worksheet against independent Python calculations."""
from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import numpy as np
import xlwings as xw

from lambda_catalog.analyze_univariate import (
    bin_counts,
    bin_edges,
    descriptive_stats,
    fd_bins,
    scott_bins,
    sturges_bins,
)
from lambda_catalog.inspection_compare import compare_values, to_float_or_none
from lambda_catalog.write_sheet_life_expectancy_data import load_life_expectancy_rows
from lambda_catalog.write_sheet_univariate import (
    UNIVARIATE_SHEET_NAME,
    _C_D,
    _C_E,
    _C_FD,
    _C_SCOTT,
    _C_STUR,
    _HB_COUNT,
    _HB_EDGE,
    _ROW_HIST_START,
    _ROW_SECTION_HDR,
    _ROW_STATS_START,
)


TOLERANCE_DECIMALS = 6
_DATA_HEADER = "Life expectancy"
_STAT_SPECS = (
    ("Mean", "mean"),
    ("Median", "median"),
    ("Mode", "mode"),
    ("Std Dev", "sd"),
    ("Variance", "variance"),
    ("Min", "min"),
    ("Max", "max"),
    ("Range", "range"),
    ("Skewness", "skewness"),
    ("Kurtosis", "kurtosis"),
    ("Count", "count"),
    ("Missing", "missing_count"),
)
_HISTOGRAM_SPECS = (
    ("Sturges", _C_STUR + _HB_EDGE, _C_STUR + _HB_COUNT, sturges_bins),
    ("Scott", _C_SCOTT + _HB_EDGE, _C_SCOTT + _HB_COUNT, scott_bins),
    ("FD", _C_FD + _HB_EDGE, _C_FD + _HB_COUNT, fd_bins),
)


def _column_values(sheet: xw.Sheet, start_row: int, col: int, count: int) -> list[Any]:
    if count <= 0:
        return []
    cell_range = sheet.range((start_row, col), (start_row + count - 1, col))
    values = cell_range.value
    return values if count > 1 else [values]


def _numeric_mismatch(expected: float, actual: Any, tolerance: int) -> bool:
    actual_number = to_float_or_none(actual)
    if actual_number is None:
        return True
    if tolerance <= 0:
        return expected != actual_number
    _, first_deviation = compare_values(float(expected), actual_number)
    return first_deviation is not None and first_deviation <= tolerance


def _load_source_data(csv_path: Path) -> list[object]:
    headers, rows = load_life_expectancy_rows(csv_path)
    try:
        data_col = headers.index(_DATA_HEADER)
    except ValueError as exc:
        raise ValueError(f"CSV is missing required column {_DATA_HEADER!r}.") from exc
    return [row[data_col] if data_col < len(row) else None for row in rows]


def read_univariate_failures(
    workbook: xw.Book,
    csv_path: Path,
    tolerance: int = TOLERANCE_DECIMALS,
) -> list[str]:
    """Return all descriptive-statistic and histogram mismatches on Univariate."""
    if UNIVARIATE_SHEET_NAME not in {sheet.name for sheet in workbook.sheets}:
        return [f"[{UNIVARIATE_SHEET_NAME}] sheet is missing"]

    sheet = workbook.sheets[UNIVARIATE_SHEET_NAME]
    source_data = _load_source_data(csv_path)
    stats_expected = descriptive_stats(source_data)
    failures: list[str] = []

    for offset, (expected_label, stat_key) in enumerate(_STAT_SPECS):
        row = _ROW_STATS_START + offset
        actual_label = sheet.range((row, _C_D)).value
        actual_value = sheet.range((row, _C_E)).value
        if actual_label != expected_label:
            failures.append(
                f"[Univariate/stats] row={row}: expected label={expected_label!r}, "
                f"excel_label={actual_label!r}"
            )

        expected_value = stats_expected[stat_key]
        if isinstance(expected_value, float) and math.isnan(expected_value):
            if str(actual_value).upper() not in {"N/A", "#N/A", "NAN"}:
                failures.append(
                    f"[Univariate/stats] stat={expected_label!r}: expected no unique mode, "
                    f"excel_calc={actual_value!r}"
                )
        elif _numeric_mismatch(float(expected_value), actual_value, tolerance):
            failures.append(
                f"[Univariate/stats] stat={expected_label!r}: "
                f"expected={expected_value!r}, excel_calc={actual_value!r}"
            )

    for method, edge_col, count_col, bin_rule in _HISTOGRAM_SPECS:
        expected_bin_count = bin_rule(source_data)
        actual_bin_count = sheet.range((_ROW_SECTION_HDR, count_col)).value
        if _numeric_mismatch(float(expected_bin_count), actual_bin_count, 0):
            failures.append(
                f"[Univariate/{method}] bin count: expected={expected_bin_count}, "
                f"excel_calc={actual_bin_count!r}"
            )
            continue

        expected_edges = bin_edges(source_data, method)
        expected_counts = bin_counts(source_data, expected_edges)
        actual_edges = _column_values(sheet, _ROW_HIST_START, edge_col, expected_bin_count)
        actual_counts = _column_values(sheet, _ROW_HIST_START, count_col, expected_bin_count)

        for index, (expected, actual) in enumerate(zip(expected_edges, actual_edges), start=1):
            if _numeric_mismatch(float(expected), actual, tolerance):
                failures.append(
                    f"[Univariate/{method}] edge={index}: "
                    f"expected={float(expected)!r}, excel_calc={actual!r}"
                )

        for index, (expected, actual) in enumerate(zip(expected_counts, actual_counts), start=1):
            if _numeric_mismatch(float(expected), actual, 0):
                failures.append(
                    f"[Univariate/{method}] count={index}: "
                    f"expected={int(expected)}, excel_calc={actual!r}"
                )

        actual_total = sum(
            value for value in (to_float_or_none(item) for item in actual_counts)
            if value is not None
        )
        expected_total = int(np.sum(expected_counts))
        if actual_total != expected_total:
            failures.append(
                f"[Univariate/{method}] total count: "
                f"expected={expected_total}, excel_calc={actual_total!r}"
            )

    return failures
