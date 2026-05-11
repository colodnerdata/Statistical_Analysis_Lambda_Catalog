"""Inspect Calc vs Exp columns in MLR test sheets and report discrepancies.

Usage:
    python tools/inspect_test_sheets.py Lambda_Library.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

import pandas as pd
import xlwings as xw

from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error


# Matches _D in write_sheet_mlr_vector_outputs_test.py
_D = 3
TOLERANCE_DECIMALS = _D * 2  # 6 decimal places

# Vector sheet layout constants (must match write_sheet_mlr_vector_outputs_test.py)
_STATS = [
    "Coefficients",
    "SE_Coefficients",
    "T_Stats",
    "P_Values",
    "CI_Lower",
    "CI_Upper",
]
_TESTS_COLS = 7
_FORMULA_COLS = 7
_TERM_COL = _TESTS_COLS + 1           # col 8 (1-based)
_CALC_START_COL = _TESTS_COLS + 2     # col 9 (1-based)
_EXPECTED_START_COL = _TESTS_COLS + _FORMULA_COLS + 1  # col 15 (1-based)


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


def read_scalar_df(workbook: xw.Book) -> pd.DataFrame:
    """Read MLR_Scalar_Test and return a DataFrame comparing Calc vs Exp values.

    Each row corresponds to one (k, allow_intercept, stat_name) triple.
    Columns: k, allow_intercept, stat_name, expected, excel_calc, abs_diff,
    first_digit_deviation.
    """
    sheet = workbook.sheets["MLR_Scalar_Test"]
    data: list[list[Any]] = sheet.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame(
            columns=["k", "allow_intercept", "stat_name", "expected",
                     "excel_calc", "abs_diff", "first_digit_deviation"]
        )

    headers = [str(h).strip() if h is not None else "" for h in data[0]]

    calc_cols: dict[str, int] = {}
    exp_cols: dict[str, int] = {}
    ind_vars_col: int | None = None
    intercept_col: int | None = None

    for i, h in enumerate(headers):
        h_norm = h.replace("\n", " ").strip()
        if h_norm.endswith(" (Calc.)"):
            stat = h_norm[: -len(" (Calc.)")].strip()
            calc_cols[stat] = i
        elif h_norm.endswith(" (Exp.)"):
            stat = h_norm[: -len(" (Exp.)")].strip()
            exp_cols[stat] = i
        elif h_norm == "ind_vars":
            ind_vars_col = i
        elif h_norm == "Allow_Intercept":
            intercept_col = i

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

        for stat_name, calc_i in calc_cols.items():
            if stat_name not in exp_cols:
                continue
            exp_i = exp_cols[stat_name]
            calc_val = row[calc_i] if calc_i < len(row) else None
            exp_val = row[exp_i] if exp_i < len(row) else None

            calc_f = float(calc_val) if calc_val is not None else None
            exp_f = float(exp_val) if exp_val is not None else None

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

    return pd.DataFrame(
        rows,
        columns=["k", "allow_intercept", "stat_name", "expected",
                 "excel_calc", "abs_diff", "first_digit_deviation"],
    )


def read_vector_df(workbook: xw.Book) -> pd.DataFrame:
    """Read MLR_Vector_Outputs_Test and return a DataFrame comparing Calc vs Exp values.

    Each row corresponds to one (k, allow_intercept, term_name, stat_name) quad.
    Columns: k, allow_intercept, term_name, stat_name, expected, excel_calc,
    abs_diff, first_digit_deviation.
    """
    sheet = workbook.sheets["MLR_Vector_Outputs_Test"]
    data: list[list[Any]] = sheet.used_range.value
    if not data or len(data) < 2:
        return pd.DataFrame(
            columns=["k", "allow_intercept", "term_name", "stat_name",
                     "expected", "excel_calc", "abs_diff", "first_digit_deviation"]
        )

    term_col_0 = _TERM_COL - 1           # 0-indexed: 7
    calc_start_0 = _CALC_START_COL - 1   # 0-indexed: 8
    exp_start_0 = _EXPECTED_START_COL - 1  # 0-indexed: 14

    rows = []
    section_k: int | None = None
    section_intercept: bool | None = None

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
            continue

        if section_k is None:
            continue

        term_name = row[term_col_0] if len(row) > term_col_0 else None
        if not term_name:
            continue

        for stat_idx, stat_name in enumerate(_STATS):
            calc_0 = calc_start_0 + stat_idx
            exp_0 = exp_start_0 + stat_idx

            calc_val = row[calc_0] if calc_0 < len(row) else None
            exp_val = row[exp_0] if exp_0 < len(row) else None

            calc_f = float(calc_val) if calc_val is not None else None
            exp_f = float(exp_val) if exp_val is not None else None

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

    return pd.DataFrame(
        rows,
        columns=["k", "allow_intercept", "term_name", "stat_name",
                 "expected", "excel_calc", "abs_diff", "first_digit_deviation"],
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inspect Calc vs Exp columns in MLR test sheets."
    )
    parser.add_argument("workbook", type=Path, help="Path to the Excel workbook.")
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Error: workbook not found: {workbook_path}", file=sys.stderr)
        sys.exit(1)

    try:
        with xw.App(visible=False, add_book=False) as app:
            try:
                workbook = app.books.open(str(workbook_path))
            except OPEN_WORKBOOK_ERRORS as exc:
                raise_excel_access_error(workbook_path, "open", exc)
            try:
                app.calculate()
                scalar_df = read_scalar_df(workbook)
                vector_df = read_vector_df(workbook)
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open", exc)

    print("=== MLR_Scalar_Test ===")
    print(scalar_df.to_string(index=False))
    print()
    print("=== MLR_Vector_Outputs_Test ===")
    print(vector_df.to_string(index=False))


if __name__ == "__main__":
    main()
