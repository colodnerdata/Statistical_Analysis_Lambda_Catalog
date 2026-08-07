"""Excel integration checks for the grid-search helper LAMBDAs."""
from __future__ import annotations

import os
from pathlib import Path

import pytest
import xlwings as xw

from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import sync_workbook_names

ROOT_DIR = Path(__file__).resolve().parents[1]
DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="set RUN_EXCEL_INTEGRATION=1 to run Excel COM integration tests",
)


def _start_excel_or_skip() -> xw.App:
    try:
        app = xw.App(visible=False, add_book=False)
        app.api.DisplayAlerts = False
        app.api.AskToUpdateLinks = False
        return app
    except Exception as exc:  # pragma: no cover - depends on host Excel installation
        pytest.skip(f"Excel COM is unavailable: {exc}")


def test_grid_helpers_evaluate_known_grid_and_row_major_tie(tmp_path: Path) -> None:
    workbook_path = tmp_path / "grid_helpers.xlsx"
    definitions = load_catalog_document(DEFINITIONS_PATH).functions
    helpers = [
        fn for fn in definitions
        if fn.name in {"Grid_Argument_Minimum", "Min_NLL_Params"}
    ]
    assert {fn.name for fn in helpers} == {"Grid_Argument_Minimum", "Min_NLL_Params"}

    app = _start_excel_or_skip()
    try:
        book = app.books.add()
        book.save(str(workbook_path))
        book.close()
        sync_workbook_names(workbook_path, helpers)
        book = app.books.open(str(workbook_path))
        sheet = book.sheets[0]
        sheet.range("A2:B3").value = [[100, 5], [200, 6]]
        sheet.range("E2:F4").value = [[10], [20], [30]]
        sheet.range("G2:G4").value = [[5], [1], [3]]
        sheet.range("B2:D3").value = [[5, 1, 3], [4, 2, 6]]
        sheet.range("F1").api.Formula2 = "=Grid_Argument_Minimum(B2:D3)"
        sheet.range("J1").api.Formula2 = "=Min_NLL_Params(E2:E4,G2:G4)"
        sheet.range("J3").api.Formula2 = "=Min_NLL_Params(A2:B3,HSTACK(C2,C3))"
        app.api.CalculateFullRebuild()

        assert sheet.range("F1:H1").options(ndim=2).value == [[1.0, 1.0, 2.0]]
        assert sheet.range("J1").value == 20.0
        assert sheet.range("J3:K3").options(ndim=2).value == [[100.0, 5.0]]

        # The second minimum is earlier in column-major order but later in row-major order.
        sheet.range("B2:D3").value = [[5, 1, 3], [1, 2, 6]]
        app.api.CalculateFullRebuild()
        assert sheet.range("F1:H1").options(ndim=2).value == [[1.0, 1.0, 2.0]]
        assert sheet.range("J1").value == 20.0

        book.save()
        book.close()
        reopened = app.books.open(str(workbook_path))
        assert reopened.sheets[0].range("J3:K3").options(ndim=2).value == [[100.0, 5.0]]
        reopened.close()
    finally:
        app.quit()
