"""Independent verification of Numeric_Complete_Cases (v3.9 sample-construction
primitive).

Numeric_Complete_Cases is the listwise-deletion sample mask the rest of the
standalone Data Transformation library composes onto: every other helper in the
bundle takes an ``[include]`` mask, and this is the mask a caller builds before
passing it in, so it takes no ``[include]`` itself. Verified the same way the
other transform-family functions are: a pure-Python mirror cross-checked against
hand-constructed fixtures (Excel ``ISNUMBER`` is stricter than
``pd.to_numeric``, which coerces numeric strings, so pandas is NOT a faithful
oracle here), plus implementation-shape assertions on the actual catalog formula
text so a future edit cannot silently change the contract without a test
noticing.

Runnable standalone (``python tests/test_numeric_complete_cases_verification.py``)
or under pytest.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import numpy as np

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

NA = float("nan")
BLANK = ""


# ── Pure-Python mirror of the workbook formula ───────────────────────────────

def _is_excel_number(v) -> bool:
    """Mirror of Excel's ISNUMBER(v).

    Excel's ISNUMBER(TRUE)/ISNUMBER(FALSE) is FALSE — booleans are their
    own type, not numbers, in Excel's type system — but Python's bool is
    a subclass of int, so a bare isinstance(v, (int, float)) check would
    wrongly accept True/False. Also accepts numpy scalar numerics
    (np.integer / np.floating), which a plain (int, float) check misses,
    and rejects NaN (ISNUMBER on an error/NaN cell is FALSE in Excel).
    """
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float, np.integer, np.floating)):
        return not (isinstance(v, float) and math.isnan(v))
    return False


def numeric_complete_cases_mirror(data):
    """Mirror of Numeric_Complete_Cases(data).

    ``data`` is a 2-D iterable of rows (list of lists). Returns a list
    (n x 1) of 1/0: 1 when every value in the row is Excel-numeric
    (ISNUMBER), 0 otherwise — a blank, text, boolean, or NaN cell makes
    its row 0. This is a *detector*, not a transform output, so a
    non-numeric row returns 0 (not "" and not NA): the library's
    NA()-exception convention applies to transform outputs like
    Ln_Positive, where an included-but-incomputable row must surface
    loudly; a mask has no such state, and 0 is the signal downstream
    callers filter on.
    """
    out = []
    for row in data:
        row = list(row)
        complete = all(_is_excel_number(v) for v in row)
        out.append(1 if complete else 0)
    return out


# ── Numeric acceptance (hand-constructed fixtures are the faithful oracle) ────

def test_all_numeric_matrix_is_all_ones() -> None:
    data = [
        [1.0, 2.0, 3.0],
        [4, 5, 6],
        [7.0, 8.0, 9.0],
    ]
    assert numeric_complete_cases_mirror(data) == [1, 1, 1]


def test_a_text_value_in_a_row_makes_it_zero() -> None:
    data = [
        [1.0, 2.0, 3.0],
        [4.0, "missing", 6.0],
        [7.0, 8.0, 9.0],
    ]
    assert numeric_complete_cases_mirror(data) == [1, 0, 1]


def test_a_blank_cell_makes_its_row_zero() -> None:
    # An empty string is how a formula renders a blank cell; ISNUMBER("")
    # is FALSE in Excel, so the row is incomplete.
    data = [
        [1.0, 2.0, 3.0],
        [4.0, BLANK, 6.0],
    ]
    assert numeric_complete_cases_mirror(data) == [1, 0]


def test_a_none_cell_makes_its_row_zero() -> None:
    # A truly empty cell (None here) is not a number.
    data = [
        [1.0, 2.0],
        [None, 4.0],
    ]
    assert numeric_complete_cases_mirror(data) == [1, 0]


def test_a_nan_cell_makes_its_row_zero() -> None:
    # ISNUMBER on a NaN/error cell is FALSE in Excel.
    data = [
        [1.0, 2.0, 3.0],
        [4.0, NA, 6.0],
    ]
    assert numeric_complete_cases_mirror(data) == [1, 0]


def test_a_boolean_cell_makes_its_row_zero() -> None:
    # ISNUMBER(TRUE)/ISNUMBER(FALSE) is FALSE — booleans are their own
    # type in Excel, not numbers — and Python's bool subclasses int, so
    # the mirror's _is_excel_number must reject bool explicitly.
    data = [
        [1.0, 2.0, 3.0],
        [4.0, True, 6.0],
        [7.0, 8.0, False],
    ]
    assert numeric_complete_cases_mirror(data) == [1, 0, 0]


def test_numpy_scalars_are_accepted_as_numeric() -> None:
    data = [
        [np.int64(1), np.float64(2.5)],
        [np.float64(3.0), NA],
    ]
    assert numeric_complete_cases_mirror(data) == [1, 0]


def test_single_column_input_works() -> None:
    # Each row is 1x1, so the all-numeric test reduces to ISNUMBER on the
    # one cell — the degenerate case the formula's COLUMNS(r) = 1 branch
    # must handle.
    data = [[1.0], ["x"], [3.0], [NA]]
    assert numeric_complete_cases_mirror(data) == [1, 0, 1, 0]


def test_a_row_is_0_1_mask_not_truthy_falsy_strings() -> None:
    # The mirror returns integer 0/1, not ""/TRUE, so the column is a
    # clean numeric array a SUMPRODUCT or further BYROW can reduce.
    result = numeric_complete_cases_mirror([[1.0], ["x"]])
    assert result == [1, 0]
    assert all(isinstance(v, int) for v in result)


# ── Implementation-shape assertions (by construction, not convention) ───────

def _catalog_functions() -> dict:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    return {fn["name"]: fn for fn in payload["functions"]}


def _formula(name: str) -> str:
    from lambda_catalog.lambda_formula_parser import (  # pylint: disable=import-outside-toplevel
        _normalize_user_formula,
        _strip_non_string_whitespace,
    )

    functions = _catalog_functions()
    return _strip_non_string_whitespace(_normalize_user_formula(functions[name]["formula_display"]))


def test_numeric_complete_cases_is_a_workbook_scoped_data_transformation_function() -> None:
    fn = _catalog_functions()["Numeric_Complete_Cases"]
    assert fn.get("category") == "Data Transformation"
    assert fn.get("subcategory") == "Sample Construction & Diagnostics"
    assert "scope" not in fn  # workbook-scoped, like Ln_Positive/Is_Balanced_Panel


def test_numeric_complete_cases_takes_only_data_no_include_argument() -> None:
    fn = _catalog_functions()["Numeric_Complete_Cases"]
    arg_names = [arg["name"] for arg in fn["arguments"]]
    assert arg_names == ["data"]
    assert fn["arguments"][0]["optional"] is False


def test_numeric_complete_cases_uses_byrow_per_row_not_a_whole_input_and_aggregate() -> None:
    formula = _formula("Numeric_Complete_Cases")
    # AND() aggregates an array to one scalar — the exact class of bug the
    # Ln_Positive degeneracy-guard comment warns about, in reverse. This
    # function must judge each row independently (the definition of
    # listwise deletion), so the all-numeric test is written as a per-row
    # count-equals-width check inside BYROW, not an AND over the input.
    assert "AND(" not in formula
    assert "BYROW(data,LAMBDA(r,N(SUM(--ISNUMBER(r))=COLUMNS(r))))" in formula


def test_numeric_complete_cases_returns_zero_one_not_blank_or_na() -> None:
    # A mask is a detector, not a transform output, so it returns 0/1.
    # The NA()-exception convention ("" = excluded, #N/A = included-but-
    # incomputable) applies to transform OUTPUTS like Ln_Positive; a
    # non-numeric row here is simply not a complete case (0), not a
    # surfaced error.
    formula = _formula("Numeric_Complete_Cases")
    assert "NA()" not in formula
    assert '""' not in formula


def test_numeric_complete_cases_appears_before_is_balanced_panel_in_document_order() -> None:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    names = [fn["name"] for fn in payload["functions"]]
    assert names.index("Numeric_Complete_Cases") < names.index("Is_Balanced_Panel")


def main() -> None:  # pragma: no cover - standalone runner
    """Run every verification directly (no pytest needed)."""
    checks = [
        (name, fn)
        for name, fn in sorted(globals().items())
        if name.startswith("test_") and callable(fn)
    ]
    for name, fn in checks:
        fn()
        print(f"PASS {name}")
    print(f"{len(checks)} verifications passed.")


if __name__ == "__main__":
    main()