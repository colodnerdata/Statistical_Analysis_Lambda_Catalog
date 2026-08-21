"""Excel COM integration checks for the v3.9 Categorical & model-construction
trio: ``Dummy_Column``, ``Interact``, ``Model_Matrix``.

The companion ``tests/test_categorical_model_construction_verification.py``
verifies pure-Python mirrors and catalog formula *text*, but never evaluates
the LAMBDAs in Excel, so it cannot catch parser or array-evaluation
differences. This module is the workbook-backed oracle: it builds a throwaway
workbook, writes the three catalog definitions into it via
``sync_workbook_names``, calls each function with fixed inputs, forces a full
recalculation, and reads the spilled results back.

Gated behind ``RUN_EXCEL_INTEGRATION=1`` (the same opt-in flag as
``test_weibull_grid_excel.py``): the GitHub-hosted ``windows-latest`` runner has
no Office, so these checks skip in CI and run only on a developer machine with
desktop Excel. A green run here is independent confirmation that Excel parses
the catalog formulas and evaluates them to the contract values — not just that
the strings match.
"""
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


def _trio_definitions():
    definitions = load_catalog_document(DEFINITIONS_PATH).functions
    helpers = [
        fn for fn in definitions
        if fn.name in {"Dummy_Column", "Interact", "Model_Matrix"}
    ]
    assert {fn.name for fn in helpers} == {
        "Dummy_Column", "Interact", "Model_Matrix"
    }
    return helpers


def test_trio_evaluate_in_excel(tmp_path: Path) -> None:
    """Exercise each function against fixed inputs and read the spills back.

    Coverage matches the contract the Python mirror pins:
    * ``Dummy_Column`` — a blank category row (default and explicit include),
      an explicit include mask on non-blank rows, and a missing level.
    * ``Interact`` — an n x k by n x 1 broadcast, plus blank and #N/A
      propagation through the ``""`` excluded-row sentinel.
    * ``Model_Matrix`` — the default intercept-on block and an explicit FALSE.
    """
    workbook_path = tmp_path / "trio.xlsx"
    helpers = _trio_definitions()

    app = _start_excel_or_skip()
    try:
        book = app.books.add()
        book.save(str(workbook_path))
        book.close()
        sync_workbook_names(workbook_path, helpers)
        book = app.books.open(str(workbook_path))
        sheet = book.sheets[0]

        # ── Inputs ──────────────────────────────────────────────────────
        # Dummy_Column category with a blank cell at row 2 (A3 left empty).
        sheet.range("A2").value = "A"
        sheet.range("A4").value = "A"
        sheet.range("A5").value = "B"
        sheet.range("A6").value = "A"
        # Explicit include mask, all TRUE — used to show a blank row stays
        # excluded even when "included" (the structural (x<>"")*inc factor).
        sheet.range("H2:H6").value = [[1], [1], [1], [1], [1]]
        # Clean category (no blanks) for the missing-level case.
        sheet.range("J2:J5").value = [["A"], ["B"], ["A"], ["C"]]

        # Interact: n x k dummy matrix (L2:M4) by n x 1 continuous (N2:N4).
        sheet.range("L2:M4").value = [[1, 0], [0, 1], [1, 0]]
        sheet.range("N2:N4").value = [[10], [20], [30]]
        # Interact blank passthrough: x1 has an empty cell at S3.
        sheet.range("S2").value = 2
        sheet.range("S4").value = 5
        sheet.range("T2:T4").value = [[10], [20], [30]]
        # Interact #N/A propagation: x1 has =NA() at X3.
        sheet.range("X2").value = 2
        sheet.range("X4").value = 5
        sheet.range("X3").api.Formula2 = "=NA()"
        sheet.range("Y2:Y4").value = [[10], [20], [30]]

        # Model_Matrix predictor block.
        sheet.range("AC2:AD3").value = [[2, 3], [4, 5]]

        # ── Formulas ────────────────────────────────────────────────────
        sheet.range("D2").api.Formula2 = '=Dummy_Column(A2:A6,"A")'
        sheet.range("F2").api.Formula2 = '=Dummy_Column(A2:A6,"A",H2:H6)'
        sheet.range("K2").api.Formula2 = '=Dummy_Column(J2:J5,"Z")'

        sheet.range("P2").api.Formula2 = "=Interact(L2:M4,N2:N4)"
        sheet.range("V2").api.Formula2 = "=Interact(S2:S4,T2:T4)"
        sheet.range("AA2").api.Formula2 = "=Interact(X2:X4,Y2:Y4)"

        sheet.range("AF2").api.Formula2 = "=Model_Matrix(AC2:AD3)"
        sheet.range("AJ2").api.Formula2 = "=Model_Matrix(AC2:AD3,FALSE)"

        # Probe spills so blank / #N/A rows read back as unambiguous strings
        # (xlwings readback of "" and #N/A varies by version; "BLANK"/"NA"
        # text does not).
        sheet.range("E2").api.Formula2 = '=IF(D2:D6="","BLANK","NUM")'
        sheet.range("G2").api.Formula2 = '=IF(F2:F6="","BLANK","NUM")'
        sheet.range("W2").api.Formula2 = '=IF(V2:V4="","BLANK","NUM")'
        sheet.range("AB2").api.Formula2 = '=IF(ISNA(AA2:AA4),"NA","NUM")'

        app.api.CalculateFullRebuild()

        # ── Dummy_Column: default include, blank row excluded ───────────
        # D spills D2:D6 = [1, "", 1, 0, 1]. The blank row is row 2 (D3).
        assert sheet.range("E2:E6").options(ndim=2).value == [
            ["NUM"], ["BLANK"], ["NUM"], ["NUM"], ["NUM"]
        ]
        assert sheet.range("D2").value == 1.0
        assert sheet.range("D4").value == 1.0
        assert sheet.range("D5").value == 0.0
        assert sheet.range("D6").value == 1.0

        # ── Dummy_Column: explicit include all-TRUE, blank STAYS excluded ─
        # F spills F2:F6 = [1, "", 1, 0, 1]. The blank row (F3) is still ""
        # because active = (x<>"")*inc ANDs the non-blank factor — an include
        # mask cannot revive a blank category row. This is the Excel-side
        # confirmation of the divergence the mirror pins (a naive mirror
        # returned 0 here; the formula returns "").
        assert sheet.range("G2:G6").options(ndim=2).value == [
            ["NUM"], ["BLANK"], ["NUM"], ["NUM"], ["NUM"]
        ]
        assert sheet.range("F2").value == 1.0
        assert sheet.range("F4").value == 1.0
        assert sheet.range("F5").value == 0.0
        assert sheet.range("F6").value == 1.0

        # ── Dummy_Column: missing level -> all-0 column, no error ─────────
        assert sheet.range("K2:K5").options(ndim=2).value == [
            [0.0], [0.0], [0.0], [0.0]
        ]

        # ── Interact: n x k by n x 1 broadcast -> n x k ───────────────────
        assert sheet.range("P2:Q4").options(ndim=2).value == [
            [10.0, 0.0], [0.0, 20.0], [30.0, 0.0]
        ]

        # ── Interact: blank passthrough -> "" at the blank row ────────────
        # V spills V2:V4 = [20, "", 150].
        assert sheet.range("W2:W4").options(ndim=2).value == [
            ["NUM"], ["BLANK"], ["NUM"]
        ]
        assert sheet.range("V2").value == 20.0
        assert sheet.range("V4").value == 150.0

        # ── Interact: #N/A propagation -> #N/A at the error row ───────────
        # AA spills AA2:AA4 = [20, #N/A, 150].
        assert sheet.range("AB2:AB4").options(ndim=2).value == [
            ["NUM"], ["NA"], ["NUM"]
        ]
        assert sheet.range("AA2").value == 20.0
        assert sheet.range("AA4").value == 150.0

        # ── Model_Matrix: default intercept-on -> ones column prepended ──
        assert sheet.range("AF2:AH3").options(ndim=2).value == [
            [1.0, 2.0, 3.0], [1.0, 4.0, 5.0]
        ]

        # ── Model_Matrix: explicit FALSE -> X unchanged ───────────────────
        assert sheet.range("AJ2:AK3").options(ndim=2).value == [
            [2.0, 3.0], [4.0, 5.0]
        ]

        book.save()
        book.close()
    finally:
        app.quit()