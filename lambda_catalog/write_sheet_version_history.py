"""Write the Version History sheet into the target workbook."""
from __future__ import annotations

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR
from .workbook_helpers import get_or_create_sheet, reset_generated_sheet


SHEET_NAME = "Version History"
_TABLE_HEADER_COLOR = (217, 225, 242)

_VERSIONS = [
    {
        "version": "1.0.0",
        "date": "2026-06-16",
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
]


def write_version_history_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Version History sheet."""
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    sheet.range("A:A").column_width = 12
    sheet.range("B:B").column_width = 16
    sheet.range("C:C").column_width = 90

    cell = sheet.range("A1")
    cell.value = "VERSION HISTORY"
    cell.api.Font.Bold = True
    sheet.range("A1:C1").color = _HEADER_COLOR

    for col, header in enumerate(["Version", "Release Date", "Summary of Changes"], start=1):
        cell = sheet.range((2, col))
        cell.value = header
        cell.api.Font.Bold = True
        cell.color = _TABLE_HEADER_COLOR

    for row_offset, entry in enumerate(_VERSIONS):
        row = 3 + row_offset
        sheet.range((row, 1)).value = entry["version"]
        sheet.range((row, 2)).value = entry["date"]
        cell = sheet.range((row, 3))
        cell.value = entry["summary"]
        cell.api.WrapText = True
        sheet.range((row, 1), (row, 3)).api.EntireRow.AutoFit()
