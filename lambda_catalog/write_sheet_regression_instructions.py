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
            "containing TRUE or FALSE for each row. A recommended approach is to add a completeness "
            "column using the Data_Completeness function:\n\n"
            "=Data_Completeness(YourTable[@[First_Predictor]:[Last_Predictor]])\n\n"
            "Replace First_Predictor and Last_Predictor with your first and last predictor column names. "
            "This returns TRUE only when every predictor value in the row is numeric, "
            "automatically excluding rows with blanks or non-numeric values from the regression."
        ),
        "body",
    ),
    (11, "", None),
    (12, "Optional — row identifiers:", "heading"),
    (
        13,
        (
            'Update data_identifiers (the "Refers To" field) to a single column in the same table '
            "containing a label for each row (e.g., a name, ID, or date). It must span exactly the "
            "same rows as All_Xs. These labels appear in the Residual Output section so you can trace "
            "flagged observations back to the records they came from. If left unset or the FILTER "
            'fails, rows are labeled generically as "Observation 1", "Observation 2", etc.'
        ),
        "body",
    ),
    (14, "", None),
    (15, "Optional — point prediction:", "heading"),
    (
        16,
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
    (17, "", None),
    (18, "Optional — categorical predictors (factor / dummy coding):", "heading"),
    (
        19,
        (
            "The regression functions require numeric predictors, so a categorical variable "
            "(e.g., Region or Treatment_Group) cannot be included in All_Xs directly. To analyze "
            "a categorical variable as a factor in the regression, transform it into dummy "
            "variables — one 0/1 indicator column per category level — using the built-in "
            "Dummy_Levels and Dummy_Code functions. Dummy_Levels returns the labels of the levels "
            "that become columns; Dummy_Code returns the matching matrix of 1s and 0s, "
            "with one row per row of your data."
        ),
        "body",
    ),
    (
        20,
        (
            "Place the dummy columns in the first empty columns to the right of your data table, "
            "leaving no gap, so that All_Xs can span your numeric predictor columns and the dummy "
            "columns as one contiguous block. In the header row, enter Dummy_Levels to spill the "
            "retained level labels; in the first data row directly beneath it, enter Dummy_Code "
            "to spill the indicator matrix:\n\n"
            "=Dummy_Levels(YourTable[Category_Column])\n"
            "=Dummy_Code(YourTable[Category_Column])\n\n"
            "If Excel automatically expands the table to absorb the new column, press Ctrl+Z once "
            "to undo the expansion — spilled arrays cannot be placed inside a structured table."
        ),
        "body",
    ),
    (
        21,
        (
            "Both functions sort the category levels and drop one reference level. By default the "
            "reference is the first level in sort order; to choose it yourself, pass it as the "
            "second argument, e.g. =Dummy_Code(YourTable[Region], \"West\"). One level must always "
            "be dropped when the model includes an intercept: with every level present, the dummy "
            "columns would sum to 1 in every row and duplicate the intercept column, making the "
            "fit impossible (perfect multicollinearity). Each dummy coefficient is then read as "
            "the average difference in the dependent variable between that level and the reference "
            "level, holding the other predictors constant. If the reference you name does not "
            "appear in the data, both functions return an error message rather than silently "
            "keeping the redundant column."
        ),
        "body",
    ),
    (
        22,
        (
            "Finally, open the Name Manager and extend All_Xs to include the new dummy columns. "
            "The predictor names shown in the MODEL SELECTION and PREDICTOR SUMMARY zones are read "
            "from the header row directly above All_Xs, which is why the Dummy_Levels labels belong "
            "in that row. Rows whose category cell is blank spill \"\" across every dummy column, so "
            "a Data_Completeness-based Regression_Sample_Include mask (see above) excludes those "
            "rows automatically. You can also pass the same include mask as the third argument to "
            "both functions to blank out excluded rows explicitly:\n\n"
            '=Dummy_Code(YourTable[Region], "West", YourTable[Include])'
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
