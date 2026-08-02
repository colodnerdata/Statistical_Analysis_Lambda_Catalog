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
_LAST_COL = _COLUMNS[-1].index

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
    {
        "version": "3.0.0",
        "date": "2026-08-02",
        "breaking": "Yes",
        "summary": (
            "v3.0: bounded model context, constructor pipeline, the "
            "two-workbook split, and the layout break. The Regression fit "
            "chain is rebuilt around a bounded [Context] argument "
            "(Has_Intercept, DF_Absorbed, response and predictor transforms) "
            "threaded through a single Model_Context constructor with "
            "Context_* field accessors, and Design_Columns / Design_Response "
            "construct the fit-time design matrix (the within estimator's "
            "demean-by-group stage lives here, not in the engines). "
            "Univariate Analysis moves to its own workbook "
            "(Lambda_Library_Univariate.xlsx) so each artifact can set its "
            "own calculation mode; the Regression workbook returns to full "
            "Automatic. The model specification block gains three columns: "
            "Interaction Term (M) and Interaction Operation (N) declare an "
            "interaction between a spec row and another Predictor, and "
            "Design Columns (O) shows how many design-matrix columns each "
            "row contributes, totalled above it as the constructed matrix's "
            "full width. A pre-flight width guard reads that total and flags "
            "a model wide enough to be slow (amber) or too wide for the "
            "sheet (red) — computed from the specification, before any "
            "matrix is built. The sheet's far right gains the Constructed "
            "Design Matrix zone, which terminates the materialization band; "
            "nothing may ever be placed to its right. M and N are reserved: "
            "validated and flagged now, read by no constructor until the "
            "interaction wiring release. WHAT BREAKS: cell ADDRESSES, not "
            "meanings. Columns A-L keep both their letters and their "
            "meanings, so a saved specification survives the upgrade "
            "unchanged and no fitted number moves — but the three new "
            "columns push every zone right of the spec block three columns "
            "over (Alpha moves Y12 to AB12, Prediction Inputs move from "
            "column AH to AK, Residual Output from AK to AN), so any "
            "formula of your own that points at a cell on this sheet needs "
            "re-pointing."
        ),
    },
]


# The Univariate artifact's own version lineage. It starts at 1.0.0 with the
# v3.0 split (Univariate Analysis becoming its own workbook) — see DECISIONS.md
# § v3.0 "Univariate becomes its own workbook" and the two-version scheme in
# README.md ("Univariate Workbook 1.0.0 · Function Library 3.0.0"). The
# pre-split Univariate history (1.1.0's "Univariate Analysis release" onward)
# belongs to the Regression workbook's lineage, which carried the sheet until
# 3.0.0; this artifact's history begins at the split.
_UNIVARIATE_VERSIONS = [
    {
        "version": "1.0.0",
        "date": "2026-08-02",
        "breaking": "No",
        "summary": (
            "Initial release of the standalone Univariate workbook. Ships "
            "Univariate Analysis in its own file so its six two-input Data "
            "Tables (Weibull, Gamma, Beta across two stages; 2,400 NLL "
            "evaluations per recalculation) can run in full Automatic "
            "calculation — distribution-fit results are never stale pending a "
            "manual Ctrl+Alt+F9. Carries the complete 126-function LAMBDA "
            "library, the Life Expectancy Data sheet the Univariate data zone "
            "reads, and this Version History. The sheet content is unchanged "
            "from the Univariate Analysis that shipped inside Lambda_Library.xlsx "
            "through 3.0.0; only the packaging (own workbook, own calc mode) is new."
        ),
    },
]


def write_version_history_sheet(workbook: xw.Book, *, artifact: str = "regression") -> None:
    """Create or refresh the Version History sheet.

    Parameters
    ----------
    workbook : xw.Book
        The target workbook.
    artifact : str, optional
        Which artifact's version lineage to write: ``"regression"`` (default)
        for the Regression workbook (Lambda_Library.xlsx) or ``"univariate"``
        for the standalone Univariate workbook (Lambda_Library_Univariate.xlsx).
        The two artifacts have independent version lines under the two-version
        scheme (shared function-library version + per-workbook version); see
        README.md "Versions".
    """
    if artifact == "univariate":
        versions = _UNIVARIATE_VERSIONS
    elif artifact == "regression":
        versions = _VERSIONS
    else:
        raise ValueError(
            f"artifact must be 'regression' or 'univariate', got {artifact!r}"
        )

    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    set_column_widths(sheet, ((c.index, c.width) for c in _COLUMNS))

    cell = sheet.range("A1")
    cell.value = "VERSION HISTORY"
    cell.api.Font.Bold = True
    sheet.range((1, 1), (1, _LAST_COL)).color = _HEADER_COLOR

    for column in _COLUMNS:
        cell = sheet.range((2, column.index))
        cell.value = column.header
        cell.api.Font.Bold = True
        cell.color = _TABLE_HEADER_COLOR

    for row_offset, entry in enumerate(versions):
        row = 3 + row_offset
        sheet.range((row, 1)).value = entry["version"]
        sheet.range((row, 2)).value = entry["date"]
        sheet.range((row, 3)).value = entry["breaking"]
        cell = sheet.range((row, _LAST_COL))
        cell.value = entry["summary"]
        cell.api.WrapText = True
        sheet.range((row, 1), (row, _LAST_COL)).api.EntireRow.AutoFit()
