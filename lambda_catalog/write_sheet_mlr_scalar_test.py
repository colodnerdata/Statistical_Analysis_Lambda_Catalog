"""Build and write the MLR_Scalar_Test worksheet for regression formula smoke tests."""
from __future__ import annotations

from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import TYPE_CHECKING

import xlwings as xw

from .catalog_schema import CatalogFunction
from .make_test_sheet import (
    _ColumnSpec,
    _RowConfig,
    write_test_table,
)
from .regression_shared import FEATURE_COLUMNS
from .workbook_helpers import (
    drop_local_name,
    reset_column_groups,
    safe_activate,
    safe_freeze_top_row,
)

if TYPE_CHECKING:
    from .regression_shared import RegressionSummary


_MLR_K_VALUES: list[int] = [1, 5, 10, 18]
_MLR_TABLE_HEADER_ROW = 1

# Replaces x_s in every formula; resolves dynamically from each row's ind_vars value.
_MLR_X_S_OFFSET = "OFFSET(y,0,1,ROWS(y),[@[ind_vars]])"


def _set_sheet_scoped_names(sheet: xw.Sheet) -> None:
    local_names = {
        "y": "=LifeExpectancyData[Life expectancy]",
        "Regression_Sample_Include": "=LifeExpectancyData[Full_Data]",
    }

    for legacy in ("fil", "x_s", "x_s_k1", "x_s_k5", "x_s_k10", "x_s_k18"):
        drop_local_name(sheet, legacy)

    for name, refers_to in local_names.items():
        drop_local_name(sheet, name)
        sheet.api.Names.Add(Name=name, RefersTo=refers_to)

    drop_local_name(sheet, "Allow_Intercept")


def _actual_formula(
    function_name: str,
    argument_names: Iterable[str],
    x_s_name: str = "x_s",
) -> str:
    args = list(argument_names)
    uses_x_s = any(a.strip().lower() == "x_s" for a in args)
    use_let = uses_x_s and x_s_name != "x_s"

    reference_map = {
        "y": "y",
        "x_s": "x_s" if use_let else x_s_name,
        "filter": "Regression_Sample_Include",
        "include": "Regression_Sample_Include",
        "allow_intercept": "[@[Allow_Intercept]]",
    }

    resolved_arguments = [
        reference_map.get(argument_name.strip().lower(), argument_name)
        for argument_name in args
    ]
    call = f"{function_name}({', '.join(resolved_arguments)})"
    return f"=LET(x_s, {x_s_name}, {call})" if use_let else f"={call}"


def _expected_values_map(summary: RegressionSummary) -> dict[str, float | int]:
    return {
        "Observations": summary.observations,
        "Regression_Degrees_Of_Freedom": summary.df_regression,
        "Total_Degrees_Of_Freedom": summary.df_total,
        "R_Squared": summary.r_squared,
        "Residual_Degrees_Of_Freedom": summary.df_residual,
        "Multiple_R": summary.multiple_r,
        "Adjusted_R_Squared": summary.adjusted_r2,
        "SS_Total": summary.ss_total,
        "SS_Residual": summary.ss_residual,
        "SS_Regression": summary.ss_regression,
        "SE_Regression": summary.se_regression,
        "PRESS": summary.press,
        "Durbin_Watson": summary.durbin_watson,
        "F_Statistic": summary.f_stat,
        "F_Statistic_P_Value": summary.p_value_f,
        "AIC": summary.aic,
        "BIC": summary.bic,
        "AICc": summary.aicc,
        "QQ_Correlation": summary.qq_correlation,
    }


def build_test_columns(
    test_table: str,
    definitions: Sequence[CatalogFunction],
) -> list[_ColumnSpec]:
    """Build the ordered column specs for the ``MLR_Scalar_Test`` table.

    The leading fixed columns (``X_Variables``, ``ind_vars``, ``Allow_Intercept``)
    are followed by one Calc column per catalog function whose ``test_table``
    equals ``test_table``; each Calc column's formula calls the function with
    the dynamic ``x_s`` OFFSET substitution (``_MLR_X_S_OFFSET``).
    """
    filtered = [d for d in definitions if d.test_table == test_table]
    fixed_columns: list[_ColumnSpec] = [
        (
            "X_Variables",
            "=TEXTJOIN(\", \",TRUE,OFFSET(y,-1,1,1,[@[ind_vars]]))",
            "General",
        ),
        ("ind_vars", None, "0"),
        ("Allow_Intercept", None, "General"),
    ]
    calc_columns: list[_ColumnSpec] = [
        (
            d.name,
            _actual_formula(d.name, d.argument_names, _MLR_X_S_OFFSET),
            d.number_format,
        )
        for d in filtered
    ]
    return fixed_columns + calc_columns


def build_mlr_row_configs(csv_path: Path) -> list[_RowConfig]:
    """Build the per-row input + expected-value configs for ``MLR_Scalar_Test``.

    For each ``k`` in ``_MLR_K_VALUES``, computes an OLS summary against the
    Life Expectancy CSV with intercept on and off (``FEATURE_COLUMNS[:k]``) and
    packs each into a row config whose inputs are ``ind_vars=k`` and the
    intercept flag, and whose expected values come from
    :func:`_expected_values_map`.
    """
    from .analyze_life_expectancy import calculate_regression_summary

    row_configs: list[_RowConfig] = []
    for k in _MLR_K_VALUES:
        feature_cols = FEATURE_COLUMNS[:k]
        summary_true = calculate_regression_summary(
            input_csv_path=csv_path,
            include_intercept=True,
            feature_columns=feature_cols,
        )
        summary_false = calculate_regression_summary(
            input_csv_path=csv_path,
            include_intercept=False,
            feature_columns=feature_cols,
        )
        row_configs.append(
            ({"ind_vars": k, "Allow_Intercept": True}, _expected_values_map(summary_true))
        )
        row_configs.append(
            ({"ind_vars": k, "Allow_Intercept": False}, _expected_values_map(summary_false))
        )
    return row_configs


def write_mlr_scalar_test_sheet(
    workbook: xw.Book,
    definitions: Sequence[CatalogFunction],
    row_configs: list[_RowConfig],
) -> None:
    """Write the ``MLR_Scalar_Test`` sheet into ``workbook``.

    Creates or reuses the sheet, installs its sheet-scoped names, and writes a
    structured table whose Calc columns call the scalar regression LAMBDA
    functions; ``row_configs`` (from :func:`build_mlr_row_configs`) supplies
    each row's inputs and expected values for the QC verification pass.
    """
    test_table = "MLR_Scalar_Test"

    sheet_names = {sheet.name for sheet in workbook.sheets}
    if test_table in sheet_names:
        sheet = workbook.sheets[test_table]
    elif len(workbook.sheets) == 1 and not workbook.sheets[0]["A1"].value:
        sheet = workbook.sheets[0]
        sheet.name = test_table
    else:
        sheet = workbook.sheets.add(name=test_table, after=workbook.sheets[-1])

    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()
    reset_column_groups(sheet)

    _set_sheet_scoped_names(sheet)

    columns = build_test_columns(test_table, definitions)
    header_row = _MLR_TABLE_HEADER_ROW
    last_data_row = header_row + len(row_configs)
    col_count = len(columns)

    rows_data = [
        (header_row + 1 + i, row_vals)
        for i, (row_vals, _) in enumerate(row_configs)
    ]

    write_test_table(sheet, test_table, columns, header_row, rows_data)

    sheet.range("A1").column_width = 100
    sheet.range((header_row, 1), (last_data_row, 1)).api.WrapText = True
    sheet.range((header_row, 1), (last_data_row, col_count)).api.EntireRow.AutoFit()

    sheet.range((header_row, 1), (header_row, col_count)).api.WrapText = True
    sheet.range((header_row, 1), (header_row, col_count)).api.EntireRow.AutoFit()

    sheet.range((header_row, 2), (last_data_row, col_count)).columns.autofit()

    safe_activate(sheet)
    safe_freeze_top_row(sheet)
