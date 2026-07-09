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
            "LAMBDA functions as well as the sheet-scoped names the Regression sheet uses."
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
    (4, "", None),
    (5, "Point the sheet at your data (one edit):", "heading"),
    (
        6,
        (
            "Open the Name Manager (Formulas → Name Manager) and update Source_Table "
            '(the "Refers To" field) to your table, including its header row — for '
            "example =MyTable[#All]. Everything else derives from this one name: the "
            "header row, the data body, and the variable list in the MODEL SPECIFICATION "
            "block all update automatically."
        ),
        "body",
    ),
    (7, "", None),
    (
        8,
        "Define your model in the MODEL SPECIFICATION block (columns A–K):",
        "heading",
    ),
    (
        9,
        (
            "The Variable column fills itself from your table's headers — one "
            "specification row per column. For each row, choose a Role:\n"
            "Response (y) — the dependent variable; declare exactly one.\n"
            "Predictor (x) — a candidate model input; the Include toggle turns it on "
            "or off for the current model run.\n"
            "Identifier (Row Label) — labeling columns such as names, IDs, or dates; "
            "they appear as row labels in the RESIDUAL OUTPUT section (multiple "
            'Identifier columns are joined with "|").\n'
            "Filter — a TRUE/FALSE or 1/0 column; only rows where every Filter column "
            "is truthy enter the regression. Declare several Filter columns to "
            "stratify — they are ANDed together.\n"
            "Omit — never used for anything; helper columns and notes."
        ),
        "body",
    ),
    (
        10,
        (
            "For each included Predictor, set its Type. Continuous predictors enter "
            "the design matrix as-is. Categorical predictors are dummy-coded "
            "automatically: each level except the reference level becomes a 0/1 "
            'column with a level-qualified name (e.g. "Status: Developing"). The '
            "reference level defaults to the first level in sort order; type a level "
            "into Reference Level to override it (the cell turns red if that level "
            "does not exist in the analysis sample). The Levels and Reference In Use "
            "columns display what the model will do: the distinct level count in the "
            "analysis sample, and the reference actually in effect. A categorical "
            "with one level (or an invalid reference) is flagged red and contributes "
            "no columns — the rest of the model still computes."
        ),
        "body",
    ),
    (
        11,
        (
            "If your table has more columns than the shipped dataset, fill in Role, "
            "Include, and Type for the extra specification rows — the dropdowns are "
            "available all the way down. A row with a blank Role is ignored. The "
            "Order and Transform columns are placeholders for a future version and "
            "are not read by any formula; they are hidden on the sheet. The Sequence "
            "column marks at most one variable as the ordering axis for future "
            "lag/difference/serial-correlation features — it is independent of Role "
            "and Type (a Predictor can also be the sequence axis). Leave it blank "
            "for non-panel data; marking two or more rows shows a red error in the "
            "status cell above the column. Sequence Period and Period In Use are "
            "Sequence's companion pair, the same input/display split as Reference "
            "Level and Reference In Use: Sequence Period is blank unless you type a "
            "number over it on the flagged row; Period In Use shows the number "
            "actually in effect — your typed override when present, else the "
            "computed candidate spacing (the most common gap between consecutive "
            "periods within a group). Lag_By and Difference_By read Period In Use "
            "when their [delta] argument is omitted — never a silent 1. The "
            "Sequence Spacing block (below the spec's 16000-row input band, with a "
            "jump link near the spec) shows the spacing spectrum and warns about "
            "gaps, off-grid spacings, and calendar-date axes (prefer an integer "
            "period index like YEAR over raw dates)."
        ),
        "body",
    ),
    (12, "", None),
    (13, "Intercept:", "heading"),
    (
        14,
        (
            "The Intercept toggle sits at the top of the Include column (C2). Leave "
            "it TRUE for models with Categorical predictors: reference-level dummy "
            "coding relies on the intercept to carry the baseline, and the toggle "
            "flags red if it is FALSE while a Categorical predictor is included."
        ),
        "body",
    ),
    (15, "", None),
    (16, "Which rows are used:", "heading"),
    (
        17,
        (
            "The analysis sample is derived automatically — a row is included when "
            "the Response is numeric, every included Continuous predictor is "
            "numeric, and every Filter column is truthy. There is nothing to "
            "configure beyond the Roles. Note that Categorical predictors impose no "
            "completeness requirement: rows with a blank category still enter the "
            "sample (all of that variable's dummies are blank there) unless a "
            "Filter excludes them."
        ),
        "body",
    ),
    (18, "", None),
    (19, "Optional — point prediction:", "heading"),
    (
        20,
        (
            "Enter a value for each design-matrix column in the orange cells under "
            "PREDICTION INPUTS (column AC) — one row per constructed column, "
            "including one per dummy (use 1 for the scenario's level, 0 for its "
            "siblings). The Training Mean column beside the inputs shows each "
            "column's mean over the analysis sample, which is also the prefilled "
            "default — so the untouched prediction is at the data's center and the "
            "SE prediction equals the SE of the regression. Results appear in the "
            "PREDICTION OUTPUTS box above; the confidence level is controlled by "
            "the Alpha cell (T12) in REGRESSION OUTPUTS."
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
