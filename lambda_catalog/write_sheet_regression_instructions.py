"""Write the Regression Instructions sheet into the target workbook."""
from __future__ import annotations

import argparse
from pathlib import Path

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR
from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    get_or_create_sheet,
    open_or_create_workbook,
    raise_excel_access_error,
    reset_generated_sheet,
)


SHEET_NAME = "Regression Instructions"
_COL_A_WIDTH = 110

# (row, text, style): style is "heading", "body", or None (empty spacer)
_ROWS: list[tuple[int, str, str | None]] = [
    (
        1,
        (
            "How to use the Regression sheet for Multilinear Regression (MLR)"
        ),
        "heading",
    ),
    (
        2,
        (
            "To analyze your own dataset, copy the Regression sheet into your workbook "
            "(right-click the tab → Move or Copy). This carries over all workbook-scoped "
            "LAMBDA functions as well as the sheet-scoped named ranges used by the Regression sheet."
        ),
        "body",
    ),
    (
        3,
        (
            "Ensure your data is organized as a structured Excel table. Select the range "
            "including column headers and press Ctrl+T, or go to Home → Format as Table."
        ),
        "body",
    ),
    (
        4,
        (
            "Next, update the named ranges that point to your data. Open the Name Manager "
            "from Formulas → Name Manager. You will see the named ranges and the "
            "workbook-scoped LAMBDA functions listed there."
        ),
        "body",
    ),
    (5, "", None),
    (6, "Required Name Range Updates:", "heading"),
    (
        7,
        (
            'In the Name Manager, update All_Xs (the “Refers To” field) to span '
            "all potential independent variable columns in your dataset."
        ),
        "body",
    ),
    (
        8,
        (
            "Update y to refer to your dependent variable column. "
            "It must span exactly the same rows as All_Xs."
        ),
        "body",
    ),
    (9, "Optional — sample include mask:", "heading"),
    (
        10,
        (
            'Update Regression_Sample_Include (the “Refers To” field) to a column in the same table '
            “containing TRUE or FALSE for each row. A recommended approach is to add a completeness “
            “column using the Data_Completeness function:\n\n”
            “=Data_Completeness(YourTable[@[First_Predictor]:[Last_Predictor]])\n\n”
            “Replace First_Predictor and Last_Predictor with your first and last predictor column names. “
            “This returns TRUE only when every predictor value in the row is numeric, “
            “automatically excluding rows with blanks or non-numeric values from the regression.”
        ),
        “body”,
    ),
    (11, "", None),
    (12, "Optional — point prediction:", "heading"),
    (
        13,
        (
            "To generate a point prediction and prediction interval, enter a value for each "
            "independent variable in the orange cells in column V under the PREDICTION INPUTS heading. "
            "Results appear automatically in the PREDICTION OUTPUTS box above. "
            "By default, these inputs are set to the mean of each independent variable in the data, "
            "so the SE prediction equals the SE of the regression.\n\n"
            "If you are making a prediction using the regression, overwrite these defaults "
            "with the values of the independent variables for your scenario."
        ),
        "body",
    ),
]


def write_regression_instructions_sheet(workbook: xw.Book) -> None:
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    sheet.range("A:A").column_width = _COL_A_WIDTH

    for row, text, style in _ROWS:
        cell = sheet.range((row, 1))
        if style is None:
            continue
        cell.value = text
        if style == "heading":
            cell.api.Font.Bold = True
            cell.color = _HEADER_COLOR
        else:
            cell.api.WrapText = True

    sheet.autofit("rows")


def _main() -> None:
    parser = argparse.ArgumentParser(
        description=f"Write the {SHEET_NAME!r} sheet into a workbook."
    )
    parser.add_argument("workbook", type=Path, help="Path to the target workbook.")
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    try:
        with xw.App(visible=True, add_book=False) as app:
            workbook, _ = open_or_create_workbook(app, workbook_path)
            try:
                write_regression_instructions_sheet(workbook)
                workbook.save(str(workbook_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)

    print(f"Sheet updated: {SHEET_NAME}")
    print(f"Workbook: {workbook_path}")


if __name__ == "__main__":
    _main()
