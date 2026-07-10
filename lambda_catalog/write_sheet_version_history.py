"""Write the Version History sheet into the target workbook."""
from __future__ import annotations

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR, SUBHDR_COLOR as _SUBHDR_COLOR
from .workbook_helpers import ColumnSpec, get_or_create_sheet, reset_generated_sheet, set_column_widths


SHEET_NAME = "Version History"
_TABLE_HEADER_COLOR = _SUBHDR_COLOR

_COLUMNS: tuple[ColumnSpec, ...] = (
    ColumnSpec(1, 12, "Version"),
    ColumnSpec(2, 16, "Release Date"),
    ColumnSpec(3, 14, "Breaking?"),
    ColumnSpec(4, 90, "Summary of Changes"),
)

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
            "replaced by a declarative variable-specification block (Role, "
            "Include, Type, Reference Level, plus reserved Order/Transform "
            "columns); X_s is promoted from a column filter to a model-matrix "
            "constructor with sheet-scoped twins (Source_Data, Sample_Include, "
            "Response_Column, Row_Labels, X_s, Constructed_Column_Names) that "
            "dissolve the v1 hard-wired ranges. Categorical predictors are "
            "reference-dropped via the rebuilt Dummy_Levels/Dummy_Code, now "
            "NA()-error-signaling with degenerate or invalid-reference cases "
            "flagged red rather than erroring. Ships the canonical rename pass "
            "and a Model Construction QC analyzer that asserts the default-spec "
            "audit values, the X_s/Constructed_Column_Names twin widths, and "
            "the full-height row-mask contract."
        ),
    },
]


def write_version_history_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Version History sheet."""
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    set_column_widths(sheet, ((c.index, c.width) for c in _COLUMNS))

    cell = sheet.range("A1")
    cell.value = "VERSION HISTORY"
    cell.api.Font.Bold = True
    sheet.range("A1:D1").color = _HEADER_COLOR

    for column in _COLUMNS:
        cell = sheet.range((2, column.index))
        cell.value = column.header
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
