"""Write the Regression Instructions sheet into the target workbook.

This sheet is a fixed, dataset-independent guide, so artifact builds
just copy the already-styled sheet out of ``TEMPLATE_PATH`` (see
``copy_static_sheet``) instead of re-running the row-by-row COM writes on
every build. ``_ROWS`` and ``_write_template_sheet`` remain the authored
source of the content; run this module's CLI after editing ``_ROWS`` to
regenerate the template, then commit the updated ``.xlsx``.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR
from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    copy_static_sheet,
    get_or_create_sheet,
    open_or_create_workbook,
    raise_excel_access_error,
    reset_generated_sheet,
)

SHEET_NAME = "Regression Instructions"
_COL_A_WIDTH = 110

_ROOT_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = _ROOT_DIR / "templates" / "static_sheets.xlsx"

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
            "block all update automatically. Source_Table points at LifeExpectancyData "
            "by default (a curated four-driver Life Expectancy model — Adult Mortality, "
            "Alcohol, percentage expenditure, and Status); a second sample table, "
            "MileageData (on the Mileage Data sheet), ships with the workbook for a "
            "multi-level categorical-encoding demo (MPG ~ Horsepower + Weight + "
            "C(Model Year) + C(Origin) — reach it with --regression-dataset auto_mpg). "
            "Practice the retarget on either before pointing at your own data."
        ),
        "body",
    ),
    (7, "", None),
    (
        8,
        "Define your model in the MODEL SPECIFICATION block (columns A–O):",
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
            "Omit — never used for anything; helper columns and notes.\n"
            "Fixed Effects — declare exactly one panel-grouping variable; the model "
            "absorbs a separate intercept per group (one-way within transformation) "
            "instead of pooled OLS, crediting the absorbed degrees of freedom back "
            "into inference. Type and Reference Level don't apply to this row — "
            "leave them as-is."
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
            "Order column is a placeholder for a future version and is not read by "
            "any formula; it is hidden on the sheet. The Transform column offers "
            "Log on the Response row and on Continuous Predictor rows: the model "
            "fits in natural-log space for that variable, and affected outputs are "
            "labeled \"(Log)\" — but the Unit-Space Fit block at AG4:AH10 reports "
            "the back-transformed R² / Adj R² / RMSE in original units (Duan "
            "smearing by default, Naive on the AH5 toggle), and the Prediction "
            "Outputs block's Original Units column (AL) carries the back-"
            "transformed point estimate and the four CI/PI bounds. The two "
            "new Residual Output columns (AZ, BA) carry Predicted Y and "
            "Residual in original units. Under Log, the Duan point estimate "
            "does not sit at the centre of the four Naive CI/PI bounds (the "
            "conditional mean is offset from the conditional median). Log "
            "is not valid on a Categorical Predictor; setting it there flags the "
            "cell red and the fit still runs with that row's Transform ignored "
            "(dummy-coded as usual). The Sequence "
            "column marks at most one variable as the ordering axis for future "
            "lag/difference/serial-correlation features — it is independent of Role "
            "and Type (a Predictor can also be the sequence axis). Leave it blank "
            "for non-panel data; marking two or more rows shows a red error in the "
            "status cell above the column. The Sequence Period column (I) is the "
            "typed override input: type a number on the flagged row to declare a Δ "
            "that differs from the candidate. The Period In Use column (J) is the "
            "live companion — it shows the typed override if column I is non-blank, "
            "otherwise the computed candidate (the most common gap between "
            "consecutive periods within a group). Lag_By and Difference_By fall "
            "back to Base_Period_Delta() when their [delta] argument is omitted "
            "— never a silent 1. The Interaction Term (M) and Interaction "
            "Operation (N) columns declare an interaction between this row "
            "and another Predictor — pick the other variable and one of "
            "Product, Difference, or Ratio. They are placeholders in this "
            "release: the dropdowns, the marginality warning (amber when "
            "the named Predictor is excluded), and the duplicate-column "
            "error (red when two rows declare the same symmetric "
            "interaction of each other) are all live, but no constructor "
            "builds the columns yet. The Design Columns column (O) shows how "
            "many design-matrix columns each row contributes — 1 for a "
            "Continuous predictor, one less than the Levels count for a "
            "Categorical one. The total above it, with the intercept added, "
            "is the full width of the constructed design matrix; it turns "
            "amber on a model wide enough to be slow and red on one too wide "
            "for the sheet."
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
            "PREDICTION INPUTS (column AK) — one row per constructed column, "
            "including one per dummy (use 1 for the scenario's level, 0 for its "
            "siblings; no Intercept row — the model's own baseline is handled "
            "automatically). The Training Mean column beside the inputs shows each "
            "column's mean over the analysis sample, which is also the prefilled "
            "default — so the untouched prediction is at the data's center. Results "
            "appear in the PREDICTION OUTPUTS box above as both a mean-response "
            "confidence interval (CI) and a wider new-observation prediction "
            "interval (PI); the confidence level is controlled by the Alpha cell "
            "(AB12) in REGRESSION OUTPUTS. If a Fixed Effects variable is declared, "
            "the FE Group cell above the inputs selects which group's own average "
            "the prediction is anchored to."
        ),
        "body",
    ),
    (21, "", None),
    (
        22,
        "Optional – faster charts on a fixed-size dataset:",
        "heading",
    ),
    (
        23,
        (
            "The seven diagnostic charts read each series from a worksheet-scoped "
            "named range (the RegChart* names) built with OFFSET and sized to the "
            "live observation count at $AB$9. OFFSET is volatile, so those ranges "
            "re-evaluate on every recalculation pass. If your dataset size is "
            "stable, you can remove that overhead by pointing each chart series "
            "directly at the absolute range the name resolves to — from row 4 "
            "(one below the column header) down to row 3+N, where N is the count "
            "shown at $AB$9 (for 200 observations, for example "
            "='Regression'!$AP$4:$AP$203). Select the chart, click a series, and "
            "replace the ='Regression'!RegChart... reference in the formula bar "
            "with the absolute range. The trade-off is that the chart no longer "
            "resizes if you add or drop rows — you must re-point it when the "
            "row count changes. The named ranges can stay defined; there is no "
            "need to delete them."
        ),
        "body",
    ),
]


def _write_template_sheet(workbook: xw.Book) -> None:
    """(Re)build this sheet's content directly from ``_ROWS``.

    Used only to author ``TEMPLATE_PATH`` — artifact builds never call
    this; they call ``write_regression_instructions_sheet``, which copies the
    sheet this function last wrote into the template.
    """
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


def write_regression_instructions_sheet(workbook: xw.Book) -> None:
    """Create or refresh the ``Regression Instructions`` sheet from the static template.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to receive the sheet.
    """
    copy_static_sheet(workbook, TEMPLATE_PATH, SHEET_NAME)


def _main() -> None:
    """Regenerate this sheet inside ``templates/static_sheets.xlsx``.

    Run after editing ``_ROWS``, then commit the updated template file.
    """
    parser = argparse.ArgumentParser(
        description=(
            f"Rebuild the {SHEET_NAME!r} sheet inside the static template workbook."
        )
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=TEMPLATE_PATH,
        help="Path to the static template workbook (default: templates/static_sheets.xlsx).",
    )
    args = parser.parse_args()

    template_path = args.template.resolve()
    try:
        with xw.App(visible=True, add_book=False) as app:
            workbook, _ = open_or_create_workbook(app, template_path)
            try:
                _write_template_sheet(workbook)
                workbook.save(str(template_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(template_path, "open or save", exc)

    print(f"Template sheet updated: {SHEET_NAME}")
    print(f"Template workbook: {template_path}")


if __name__ == "__main__":
    _main()
