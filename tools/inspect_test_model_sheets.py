"""Verify Lambda_Library_TestModels.xlsx — one pre-built sheet per test model.

The complement to ``tools/inspect_regression_sheet.py``. That one writes each
case onto the shared ``Regression`` sheet and reads it back; this one reads
sheets that were already built and populated by ``build_test_models.py`` and
**writes nothing**.

That difference is the point, not an optimization:

* No state leaks between cases, because no case touches another's sheet. The
  defensive re-setting the legacy path does on every iteration — Source_Table,
  the intercept toggle, ``$AK$12``, the whole spec block, the prediction
  inputs — exists only because one sheet is reused.
* No recalculation per case. The workbook was rebuilt once at the end of the
  build, so every sheet already holds final values.
* A failure leaves something to open. ``M09 Cat x Cat Full Product`` is a
  tab, and the numbers that disagreed are on it.

Two kinds of sheet, two kinds of check. A fittable case compares the five
numeric zones through the shared reader. A guard-state case compares status
cells, the per-row Design Columns audit and the Model Formula — no fit
statistics, because most guard configurations have no fittable model.

Usage:
    python tools/inspect_test_model_sheets.py Lambda_Library_TestModels.xlsx
"""
from __future__ import annotations

import argparse
import sys
from collections.abc import Callable
from pathlib import Path

import xlwings as xw

from lambda_catalog.analyze_regression_guard_states import (
    GuardStateExpected,
    build_guard_state_cases,
    calculate_guard_state_case,
)
from lambda_catalog.analyze_regression_spec import (
    RegressionSpecExpected,
    build_regression_spec_cases,
    calculate_regression_spec_case,
)
from lambda_catalog.regression_spec_sheet_io import (
    read_case_comparison_rows,
    read_model_formula,
    read_response_readout,
)
from lambda_catalog.workbook_helpers import OPEN_WORKBOOK_ERRORS, raise_excel_access_error
from lambda_catalog.write_sheet_model_construction import (
    _C_DESIGN_COLUMNS,
    _C_ROLE,
    _C_SEQUENCE,
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
    _INTERCEPT_ROW,
)

# Shared with inspect_regression_sheet.py: a comparison fails when the first
# differing digit sits at or before this decimal place.
_D = 3
TOLERANCE_DECIMALS = _D * 2  # 6

# Where the guard-state facts live, all imported rather than spelled out.
# _C_ROLE is column B (the Fixed Effects cardinality error at B1);
# _C_SEQUENCE is column H (the multi-flag status at H2). The width guard's
# M2 and the Σ total at O1 come from the interaction pair's own columns.
_ROW_STATUS = 1
_ROW_SEQUENCE_STATUS = _INTERCEPT_ROW  # H2, on the intercept-control row


def _as_text(value: object) -> str:
    """Render a cell value as the text a status comparison expects."""
    if value is None:
        return ""
    return str(value).strip()


def verify_model_sheet(
    sheet: xw.Sheet, expected: RegressionSpecExpected
) -> list[str]:
    """Compare one fittable case's sheet against its oracle."""
    failures: list[str] = []
    sections = read_case_comparison_rows(sheet, expected)
    for section, rows in sections.items():
        for row in rows:
            deviation = row["first_digit_deviation"]
            if deviation is not None and deviation <= TOLERANCE_DECIMALS:
                identity = " ".join(
                    f"{key}={row[key]!r}"
                    for key in (
                        "config_name", "stat_name", "predictor_name",
                        "term_name", "row_idx",
                    )
                    if key in row
                )
                failures.append(
                    f"[TestModel/{expected.case.sheet_name}/{section}] {identity}: "
                    f"expected={row['expected']!r}, "
                    f"excel_calc={row['excel_calc']!r}, "
                    f"abs_diff={row['abs_diff']!r}, fdd={deviation}"
                )

    # The Model Formula cell is a string, so it is compared exactly rather
    # than numerically — and it is the one assertion that catches a spec
    # applied to the wrong sheet, since every other zone would still be
    # internally consistent.
    actual_formula = _as_text(read_model_formula(sheet))
    wanted_formula = expected.results.unit_space.model_formula
    if actual_formula != wanted_formula:
        failures.append(
            f"[TestModel/{expected.case.sheet_name}/formula] "
            f"expected={wanted_formula!r}, excel_calc={actual_formula!r}"
        )
    return failures


def verify_guard_sheet(sheet: xw.Sheet, expected: GuardStateExpected) -> list[str]:
    """Compare one guard-state sheet's status cells and audit displays."""
    failures: list[str] = []
    name = expected.case.sheet_name

    def _check(label: str, wanted: object, actual: object) -> None:
        if wanted != actual:
            failures.append(
                f"[TestModel/{name}/{label}] expected={wanted!r}, "
                f"excel_calc={actual!r}"
            )

    _check(
        "model_formula",
        expected.model_formula,
        _as_text(read_model_formula(sheet)),
    )
    _check(
        "response_readout",
        expected.response_name,
        _as_text(read_response_readout(sheet)),
    )
    _check(
        "sequence_status",
        expected.sequence_status,
        _as_text(sheet.range(_ROW_SEQUENCE_STATUS, _C_SEQUENCE).value),
    )
    _check(
        "fixed_effects_status",
        expected.fixed_effects_status,
        _as_text(sheet.range(_ROW_STATUS, _C_ROLE).value),
    )

    # Column O, one cell per spec row. A non-Predictor row shows "" and a
    # Predictor row a count, so the comparison is on the rendered value.
    for offset, wanted in enumerate(expected.design_columns):
        actual = sheet.range(_SPEC_FIRST_DATA_ROW + offset, _C_DESIGN_COLUMNS).value
        rendered: object = "" if actual is None else actual
        if isinstance(rendered, float) and rendered.is_integer():
            rendered = int(rendered)
        if rendered != wanted:
            failures.append(
                f"[TestModel/{name}/design_columns row={offset + 1}] "
                f"expected={wanted!r}, excel_calc={rendered!r}"
            )
    return failures


def verify_test_model_workbook(
    workbook: xw.Book,
    sheet_names: list[str] | None = None,
    on_sheet: Callable[[int, int, object, int], None] | None = None,
) -> list[str]:
    """Verify every test-model sheet present in ``workbook``.

    ``sheet_names`` bounds the check to what a particular build produced
    (``build_test_models.py --cases`` builds a subset). When omitted, every
    registered case whose sheet is present is checked and absent ones are
    skipped silently — a partial workbook is a legitimate developer state,
    not a failure.

    ``on_sheet(index, total, case, failure_count)`` is called after each
    sheet so a caller can report progress. Reading ~46 sheets takes minutes,
    and a per-sheet failure count is what turns a 20,000-line failure list
    into "case L07 is the problem" without post-processing.
    """
    present = {sheet.name for sheet in workbook.sheets}
    if sheet_names is not None:
        present &= set(sheet_names)

    model_cases = [c for c in build_regression_spec_cases() if c.sheet_name in present]
    guard_cases = [c for c in build_guard_state_cases() if c.sheet_name in present]
    total = len(model_cases) + len(guard_cases)

    failures: list[str] = []
    for index, case in enumerate(model_cases, start=1):
        expected = calculate_regression_spec_case(case)
        found = verify_model_sheet(workbook.sheets[case.sheet_name], expected)
        failures.extend(found)
        if on_sheet is not None:
            on_sheet(index, total, case, len(found))
    for offset, case in enumerate(guard_cases, start=1):
        found = verify_guard_sheet(
            workbook.sheets[case.sheet_name], calculate_guard_state_case(case)
        )
        failures.extend(found)
        if on_sheet is not None:
            on_sheet(len(model_cases) + offset, total, case, len(found))
    return failures


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Verify the test-model workbook against the Python oracles."
    )
    parser.add_argument("workbook", type=Path, help="Path to the Excel workbook.")
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Error: workbook not found: {workbook_path}", file=sys.stderr)
        sys.exit(1)

    try:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            try:
                workbook = app.books.open(str(workbook_path))
            except OPEN_WORKBOOK_ERRORS as exc:
                raise_excel_access_error(workbook_path, "open", exc)
            try:
                failures = verify_test_model_workbook(workbook)
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "inspect", exc)

    for message in failures:
        print(f"ERROR {message}")
    if failures:
        print(f"\n{len(failures)} mismatch(es).")
        sys.exit(1)
    print("All test-model sheets match the oracle.")


if __name__ == "__main__":
    main()
