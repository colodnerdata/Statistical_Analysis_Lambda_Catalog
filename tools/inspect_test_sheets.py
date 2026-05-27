"""Inspect Calc columns in MLR test sheets and compare against Python-computed expected values.

Expected values come from the analysis cache / analyze_life_expectancy.py, not from
the (Exp.) columns in the workbook, so the tool survives removal of those columns.

Usage:
    python tools/inspect_test_sheets.py Lambda_Library.xlsx [--csv path/to/data.csv]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

import pandas as pd  # type: ignore[import-untyped]  # pyright: ignore[reportMissingTypeStubs]
import xlwings as xw

from lambda_catalog.analyze_life_expectancy import DEFAULT_INPUT_CSV, RegressionVectors
from lambda_catalog.analysis_cache import get_analysis_results
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error


# Matches _D in write_sheet_mlr_vector_outputs_test.py
_D = 3
TOLERANCE_DECIMALS = _D * 2  # 6 decimal places

# Layout constants (must match write_sheet_mlr_vector_outputs_test.py and write_sheet_mlr_observation_test.py)
_STATS = [
    "Coefficients",
    "SE_Coefficients",
    "T_Stats",
    "P_Values",
    "CI_Lower",
    "CI_Upper",
]
# RegressionVectors field names in the same order as _STATS
_STAT_FIELDS = [
    "coefficients",
    "std_errors",
    "t_stats",
    "p_values",
    "ci_lower",
    "ci_upper",
]
_TERM_COL = 1
_CALC_START_COL = 2

_DF_SCALAR_COLS = [
    "k", "allow_intercept", "stat_name",
    "expected", "excel_calc", "abs_diff", "first_digit_deviation",
]
_DF_VECTOR_COLS = [
    "k", "allow_intercept", "term_name", "stat_name",
    "expected", "excel_calc", "abs_diff", "first_digit_deviation",
]
_DF_OBS_COLS = ["k", "allow_intercept", "row_idx", "stat_name", "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
_OBS_STATS = [
    "Observation_Num",
    "Percentile",
    "Y_Ranked",
    "Normal_Scores",
    "Predictions",
    "Residuals",
    "Scaled_Residuals",
    "Scaled_Residuals_Ranked",
]
_OBS_STAT_FIELDS = [
    "observation_num",
    "percentile",
    "y_ranked",
    "normal_scores",
    "predictions",
    "residuals",
    "scaled_residuals",
    "scaled_residuals_ranked",
]


def _first_digit_deviation(expected: float, actual: float) -> int | None:
    """Return the decimal place where expected and actual first deviate.

    Scans d from 15 down to 0. Returns d+1 where d is the finest precision
    at which round(expected, d) == round(actual, d). Returns None if the
    values are identical, 0 if they differ even at d=0.
    """
    if expected == actual:
        return None
    for d in range(15, -1, -1):
        if round(expected, d) == round(actual, d):
            return d + 1
    return 0


def read_scalar_df(
    workbook: xw.Book,
    scalar_row_configs: list,
) -> pd.DataFrame:
    """Read MLR_Scalar_Test Calc columns and compare against Python-computed expected values.

    Expected values are taken from ``scalar_row_configs`` (Python-computed OLS metrics),
    not from the (Exp.) columns in the sheet. Only the Calc columns and the row-identity
    columns (ind_vars, Allow_Intercept) are read from the workbook.

    Each row in the returned DataFrame corresponds to one (k, allow_intercept, stat_name)
    triple. Columns: k, allow_intercept, stat_name, expected, excel_calc, abs_diff,
    first_digit_deviation.
    """
    sheet = workbook.sheets["MLR_Scalar_Test"]
    data: list[list[Any]] = sheet.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame(columns=_DF_SCALAR_COLS)

    headers = [str(h).strip() if h is not None else "" for h in data[0]]

    calc_cols: dict[str, int] = {}
    ind_vars_col: int | None = None
    intercept_col: int | None = None

    _SCALAR_FIXED = {"", "X_Variables", "ind_vars", "Allow_Intercept"}
    for i, h in enumerate(headers):
        h_norm = h.replace("\n", " ").strip()
        if h_norm == "ind_vars":
            ind_vars_col = i
        elif h_norm == "Allow_Intercept":
            intercept_col = i
        elif h_norm not in _SCALAR_FIXED:
            calc_cols[h_norm] = i

    # Build lookup: (k, allow_intercept) -> expected_dict
    config_lookup: dict[tuple[int, bool], dict[str, float]] = {}
    for item in scalar_row_configs:
        row_vals = item[0]
        expected_values = item[1]
        k_val = row_vals.get("ind_vars")
        ai_val = bool(row_vals.get("Allow_Intercept"))
        if k_val is not None:
            config_lookup[(int(k_val), ai_val)] = expected_values

    rows = []
    for row in data[1:]:
        if not any(v is not None for v in row):
            continue
        k = (
            int(row[ind_vars_col])
            if ind_vars_col is not None and row[ind_vars_col] is not None
            else None
        )
        allow_intercept = (
            bool(row[intercept_col])
            if intercept_col is not None and row[intercept_col] is not None
            else None
        )

        expected_dict: dict[str, float] = {}
        if k is not None and allow_intercept is not None:
            expected_dict = config_lookup.get((k, allow_intercept), {})

        for stat_name, calc_i in calc_cols.items():
            calc_val = row[calc_i] if calc_i < len(row) else None
            calc_f = float(calc_val) if calc_val is not None else None

            raw_exp = expected_dict.get(stat_name)
            exp_f = float(raw_exp) if raw_exp is not None else None

            abs_diff = (
                abs(calc_f - exp_f)
                if calc_f is not None and exp_f is not None
                else None
            )
            fdd = (
                _first_digit_deviation(exp_f, calc_f)
                if exp_f is not None and calc_f is not None
                else None
            )

            rows.append({
                "k": k,
                "allow_intercept": allow_intercept,
                "stat_name": stat_name,
                "expected": exp_f,
                "excel_calc": calc_f,
                "abs_diff": abs_diff,
                "first_digit_deviation": fdd,
            })

    return pd.DataFrame(rows, columns=_DF_SCALAR_COLS)


def read_vector_df(
    workbook: xw.Book,
    vector_row_configs: list,
) -> pd.DataFrame:
    """Read MLR_Vector_Outputs_Test Calc columns and compare against Python-computed expected values.

    Expected values are taken from ``vector_row_configs`` (Python-computed OLS vectors),
    not from the (Exp.) columns in the sheet. Only the Calc columns and the Term column
    are read from the workbook.

    Each row in the returned DataFrame corresponds to one
    (k, allow_intercept, term_name, stat_name) quad. Columns: k, allow_intercept,
    term_name, stat_name, expected, excel_calc, abs_diff, first_digit_deviation.
    """
    sheet = workbook.sheets["MLR_Vector_Outputs_Test"]
    data: list[list[Any]] = sheet.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame(columns=_DF_VECTOR_COLS)

    term_col_0 = _TERM_COL - 1           # 0-indexed: 7
    calc_start_0 = _CALC_START_COL - 1   # 0-indexed: 8

    # Build lookup: (k, allow_intercept) -> RegressionVectors
    config_lookup: dict[tuple[int, bool], RegressionVectors] = {
        (k, allow_intercept): vectors
        for k, allow_intercept, vectors in vector_row_configs
    }

    rows = []
    section_k: int | None = None
    section_intercept: bool | None = None
    term_offset = 0  # position within the current section's term list

    for row in data[1:]:
        if not row:
            continue
        col_a = row[0]

        # Section label row: "k=1 | Intercept=TRUE"
        if isinstance(col_a, str) and col_a.startswith("k="):
            parts = col_a.split("|")
            k_str = parts[0].strip()
            int_str = parts[1].strip() if len(parts) > 1 else ""
            try:
                section_k = int(k_str.split("=")[1])
            except (IndexError, ValueError):
                section_k = None
            section_intercept = "TRUE" in int_str.upper()
            term_offset = 0
            continue

        if section_k is None:
            continue

        term_name = row[term_col_0] if len(row) > term_col_0 else None
        if not term_name:
            term_offset = 0
            continue

        vectors = None
        if section_intercept is not None:
            vectors = config_lookup.get((section_k, section_intercept))

        for stat_idx, stat_name in enumerate(_STATS):
            calc_0 = calc_start_0 + stat_idx
            calc_val = row[calc_0] if calc_0 < len(row) else None
            calc_f = float(calc_val) if calc_val is not None else None

            exp_f: float | None = None
            if vectors is not None:
                vector_data: tuple = getattr(vectors, _STAT_FIELDS[stat_idx])
                if term_offset < len(vector_data):
                    exp_f = float(vector_data[term_offset])

            abs_diff = (
                abs(calc_f - exp_f)
                if calc_f is not None and exp_f is not None
                else None
            )
            fdd = (
                _first_digit_deviation(exp_f, calc_f)
                if exp_f is not None and calc_f is not None
                else None
            )

            rows.append({
                "k": section_k,
                "allow_intercept": section_intercept,
                "term_name": str(term_name),
                "stat_name": stat_name,
                "expected": exp_f,
                "excel_calc": calc_f,
                "abs_diff": abs_diff,
                "first_digit_deviation": fdd,
            })

        term_offset += 1

    return pd.DataFrame(rows, columns=_DF_VECTOR_COLS)


def read_observation_df(workbook: xw.Book, observation_row_configs: list) -> pd.DataFrame:
    sheet = workbook.sheets["MLR_Observation_Test"]
    data: list[list[Any]] = sheet.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame(columns=_DF_OBS_COLS)
    lookup = {(k, ai): v for k, ai, v in observation_row_configs}
    rows = []
    section_k = None
    section_intercept = None
    row_offset = 0
    section_n: int | None = None
    calc_start = _CALC_START_COL - 1
    for row in data[1:]:
        if isinstance(row[0] if row else None, str) and str(row[0]).startswith("k="):
            section_k = int(str(row[0]).split(",")[0].split("=")[1])
            section_intercept = "TRUE" in str(row[0]).upper()
            vectors = None
            if section_intercept is not None:
                vectors = lookup.get((section_k, section_intercept))
            section_n = len(vectors.observation_num) if vectors is not None else None
            row_offset = 0
            continue
        if section_k is None or not row or (section_n is not None and row_offset >= section_n):
            continue
        vectors = None
        if section_intercept is not None:
            vectors = lookup.get((section_k, section_intercept))
        for i, stat in enumerate(_OBS_STATS):
            calc_val = row[calc_start + i] if len(row) > calc_start + i else None
            calc_f = float(calc_val) if calc_val is not None else None
            exp_f = None
            if vectors is not None and row_offset < len(getattr(vectors, _OBS_STAT_FIELDS[i])):
                exp_f = float(getattr(vectors, _OBS_STAT_FIELDS[i])[row_offset])
            rows.append({"k": section_k, "allow_intercept": section_intercept, "row_idx": row_offset + 1, "stat_name": stat, "expected": exp_f, "excel_calc": calc_f, "abs_diff": abs(calc_f-exp_f) if calc_f is not None and exp_f is not None else None, "first_digit_deviation": _first_digit_deviation(exp_f, calc_f) if calc_f is not None and exp_f is not None else None})
        row_offset += 1
    return pd.DataFrame(rows, columns=_DF_OBS_COLS)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inspect Calc columns in MLR test sheets against Python-computed expected values."
    )
    parser.add_argument("workbook", type=Path, help="Path to the Excel workbook.")
    parser.add_argument(
        "--csv",
        type=Path,
        default=DEFAULT_INPUT_CSV,
        help="Path to the Life Expectancy CSV used to compute expected values.",
    )
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Error: workbook not found: {workbook_path}", file=sys.stderr)
        sys.exit(1)

    scalar_row_configs, vector_row_configs, observation_row_configs = get_analysis_results(args.csv)

    try:
        with xw.App(visible=False, add_book=False) as app:
            try:
                workbook = app.books.open(str(workbook_path))
            except OPEN_WORKBOOK_ERRORS as exc:
                raise_excel_access_error(workbook_path, "open", exc)
            try:
                app.calculate()
                scalar_df = read_scalar_df(workbook, scalar_row_configs)
                vector_df = read_vector_df(workbook, vector_row_configs)
                observation_df = read_observation_df(workbook, observation_row_configs)
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "inspect", exc)

    print("=== MLR_Scalar_Test ===")
    print(scalar_df.to_string(index=False))
    print()
    print("=== MLR_Vector_Outputs_Test ===")
    print(vector_df.to_string(index=False))
    print()
    print("=== MLR_Observation_Test ===")
    print(observation_df.to_string(index=False))


if __name__ == "__main__":
    main()
