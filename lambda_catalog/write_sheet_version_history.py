"""Write the Version History sheet into the target workbook."""
from __future__ import annotations

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR, SUBHDR_COLOR as _SUBHDR_COLOR
from .workbook_helpers import get_or_create_sheet, reset_generated_sheet


SHEET_NAME = "Version History"
_TABLE_HEADER_COLOR = _SUBHDR_COLOR

# Versions follow the interface-based semantic-versioning convention in
# ROADMAP.md: MAJOR only when a workbook built against the prior version would
# stop working or silently compute something different. Under that definition
# the ladder was renumbered on 2026-07-05 — Univariate (previously released as
# 2.0.0) is 1.1.0, the hardening release (previously 2.1.0) is 1.2.0, and
# Specification-Driven Regression (previously 3.0.0) is 2.0.0, the one MAJOR.
_VERSIONS = [
    {
        "version": "1.0.0",
        "date": "2026-06-16",
        "breaking": "No",
        "summary": (
            "Initial release — complete OLS/MLR engine with full model-fit and ANOVA "
            "statistics, coefficient inference (SE, t, p, CIs, partial R²/correlation), "
            "multicollinearity diagnostics (VIF, Tolerance, correlation matrix), residual "
            "and influence diagnostics (residuals, studentized residuals, hat diagonal, "
            "Cook's distance, Q-Q machinery, Durbin-Watson), cross-validation (PRESS, "
            "LOOCV), information criteria (AIC, AICc, BIC), distributional exploration "
            "(skewness, kurtosis, Pearson/Spearman), and prediction with confidence "
            "intervals. Ships with Regression sheet, Diagnostic Guide, Regression "
            "Instructions, pre-built diagnostic charts, and this Version History."
        ),
    },
    {
        "version": "1.1.0",
        "date": "2026-06-29",
        "breaking": "No",
        "summary": (
            "Univariate Analysis release. Adds a live Univariate sheet with descriptive "
            "statistics, missing-count handling, Sturges/Scott/Freedman-Diaconis "
            "histogram binning, dynamic histogram charts, lower/upper edge and midpoint "
            "helpers, eight CDF and NLL distribution families (Normal, Lognormal, "
            "Exponential, Weibull, Gamma, Triangular, Beta, BetaPERT), goodness-of-fit "
            "ranking with AIC, BIC, Anderson-Darling, and Kolmogorov-Smirnov, and "
            "two-stage native Data Table grid-search fitting for Weibull, Gamma, and "
            "Beta. Also establishes the shared sheet-style palette and the xlwings COM "
            "chart-writing pattern used by generated worksheets."
        ),
    },
    {
        "version": "1.2.0",
        "date": "2026-07-03",
        "breaking": "No",
        "summary": (
            "Workbook hardening and regression usability update. Adds Name "
            "Manager notes to catalog functions, a predicted-variable readout and row "
            "identifiers on the Regression sheet, LOOCV_Residual as an observation-level "
            "diagnostic, stronger intercept-only and undersized-sample guards, explicit "
            "error handling instead of silent first-predictor fallbacks, a chart data "
            "series for the identity line, and safer production-build retry/RPC handling. "
            "Includes categorical helper groundwork and documentation while leaving the "
            "full specification-driven regression sheet for the v2.0 milestone."
        ),
    },
    {
        "version": "2.0.0",
        "date": "2026-07-05",
        "breaking": "Yes",
        "summary": (
            "Specification-Driven Regression release (MAJOR — existing workbook "
            "inputs change meaning). The Regression sheet's control block is "
            "replaced by a declarative variable-specification block (Role, Include, "
            "Type, Reference Level, with reserved Order and Transform columns) "
            "spanning every column of the source table, and X_s is promoted from a "
            "column filter to a model-matrix constructor. Sheet-scoped constructor "
            "names (Source_Data as the dataset-retarget point, Sample_Include with "
            "role-aware completeness, Response_Column, Row_Labels, X_s, and its "
            "structural twin Constructed_Column_Names) dissolve the v1 hard-wired "
            "ranges, with filtered display zones and a row-1 audit strip (k, rows, "
            "response, responses, included rows). Categorical predictors are "
            "reference-dropped via the rebuilt Dummy_Levels/Dummy_Code, which now "
            "signal failure with a real #N/A error instead of text; degenerate or "
            "invalid-reference categoricals contribute zero columns and are flagged "
            "red rather than erroring the sheet. Includes the canonical rename pass "
            "(e.g. P_Value_F to F_Statistic_P_Value, F_Stat to F_Statistic, "
            "Grid_Argmin to Grid_Argument_Minimum). The QC build gains a Model "
            "Construction analyzer asserting the default-spec audit values, the "
            "X_s/Constructed_Column_Names twin widths, the full-height row-mask "
            "contract, and a stratified-Filter degeneracy case."
        ),
    },
]


def write_version_history_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Version History sheet."""
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    sheet.range("A:A").column_width = 12
    sheet.range("B:B").column_width = 16
    sheet.range("C:C").column_width = 14
    sheet.range("D:D").column_width = 90

    cell = sheet.range("A1")
    cell.value = "VERSION HISTORY"
    cell.api.Font.Bold = True
    sheet.range("A1:D1").color = _HEADER_COLOR

    headers = ["Version", "Release Date", "Breaking?", "Summary of Changes"]
    for col, header in enumerate(headers, start=1):
        cell = sheet.range((2, col))
        cell.value = header
        cell.api.Font.Bold = True
        cell.color = _TABLE_HEADER_COLOR

    for row_offset, entry in enumerate(_VERSIONS):
        row = 3 + row_offset
        sheet.range((row, 1)).value = entry["version"]
        sheet.range((row, 2)).value = entry["date"]
        sheet.range((row, 3)).value = entry["breaking"]
        cell = sheet.range((row, 4))
        cell.value = entry["summary"]
        cell.api.WrapText = True
        sheet.range((row, 1), (row, 4)).api.EntireRow.AutoFit()
