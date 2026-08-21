"""Independent verification of the v3.9 Categorical & model-construction trio:
``Dummy_Column``, ``Interact``, ``Model_Matrix``.

These are the free-form standalone counterparts to the v3.0 spec-driven constructor
path (``Predictor_Columns`` encodes inline; the spec-block interaction columns are
built by the constructor, not by ``Interact``). The spec-driven path does not call any
of the three — the same relationship ``Predictor_Columns`` already has with
``Dummy_Code``. Specified in ``docs/ARCHITECTURE.md`` §5 ("Categorical & Model
Construction") and recorded in ``docs/DECISIONS.md`` as "specified, not yet built."

Verified the same way the other transform-family functions are: pure-Python mirrors
cross-checked against hand-constructed fixtures (Excel ``ISNUMBER``/arithmetic
semantics, not pandas coercion — pandas would coerce numeric strings and lose the
``""``-excluded-row distinction), plus implementation-shape assertions on the actual
catalog formula text so a future edit cannot silently change the contract without a
test noticing.

Runnable standalone (``python tests/test_categorical_model_construction_verification.py``)
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


# ── Shared helpers ────────────────────────────────────────────────────────────

def _is_blank(v) -> bool:
    """The library's excluded-row sentinel: a None (truly empty cell) or "" (the
    string a formula renders a blank as)."""
    return v is None or v == BLANK


def _is_na(v) -> bool:
    return isinstance(v, float) and math.isnan(v)


def _to_num(v):
    """Excel arithmetic coercion: bool -> 1/0 (TRUE*5 = 5), number -> itself, numpy
    scalars -> float. Returns None for non-numeric text (which Excel would surface
    as #VALUE!, mirrored here as NA)."""
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    return None


# ── Pure-Python mirrors of the workbook formulas ─────────────────────────────

def dummy_column_mirror(category, level, include=None):
    """Mirror of Dummy_Column(category, level, [include]).

    One indicator column for an explicit level: 1 where category == level, 0
    otherwise, for included rows; "" on excluded rows and blank category cells.
    Blank cells are normalized to "" first (mirroring
    ``TOCOL(IF(category="", "", category), 0)``), then the default include mask
    excludes them (``active = (x <> "") * inc``) — the same blank-exclusion
    convention Dummy_Levels/Dummy_Code use. A level not present in the sample
    yields an all-0 column, NOT an error (the free-form primitive imposes no
    treatment-coding semantics and no reference validation).
    """
    cat = [BLANK if v is None else v for v in category]
    if include is None:
        inc = [(v != BLANK) for v in cat]
    else:
        inc = [bool(v) for v in include]
    return [
        BLANK if not inc[i] else (1 if cat[i] == level else 0)
        for i in range(len(cat))
    ]


def _cell_product(a, b):
    """Mirror of ``IF((a="")+(b=""), "", a*b)``: blank propagates to ""; #N/A and
    other errors propagate as NA; numbers (and bools, coerced) multiply;
    non-numeric text -> #VALUE! mirrored as NA."""
    if _is_blank(a) or _is_blank(b):
        return BLANK
    if _is_na(a) or _is_na(b):
        return NA
    na, nb = _to_num(a), _to_num(b)
    if na is None or nb is None:
        return NA  # text x number -> #VALUE!
    return na * nb


def interact_mirror(x1, x2):
    """Mirror of Interact(x1, x2): elementwise product with Excel broadcasting
    (n x k x n x 1 -> n x k), "" passthrough, and error propagation. ``x1`` and
    ``x2`` are lists of rows (list of lists). Non-conformable shapes (k != m and
    neither is 1) -> NA, mirroring Excel's #VALUE!."""
    x1 = [[BLANK if c is None else c for c in row] for row in x1]
    x2 = [[BLANK if c is None else c for c in row] for row in x2]
    n = len(x1)
    if n == 0:
        return []
    k1 = len(x1[0])
    k2 = len(x2[0])
    if k1 == 1 and k2 > 1:  # column broadcasts across x2's matrix
        return [[_cell_product(x1[i][0], x2[i][j]) for j in range(k2)] for i in range(n)]
    if k2 == 1 and k1 > 1:  # column broadcasts across x1's matrix
        return [[_cell_product(x1[i][j], x2[i][0]) for j in range(k1)] for i in range(n)]
    if k1 == k2:  # elementwise
        return [[_cell_product(x1[i][j], x2[i][j]) for j in range(k1)] for i in range(n)]
    return NA  # non-conformable -> #VALUE!


def model_matrix_mirror(X, add_intercept=True):
    """Mirror of Model_Matrix(X, [add_intercept]): prepend a column of 1s when
    add_intercept is TRUE (the default); return X unchanged when FALSE."""
    X = [list(row) for row in X]
    if add_intercept:
        return [[1] + row for row in X]
    return X


# ── Numeric acceptance: Dummy_Column ─────────────────────────────────────────

def test_dummy_column_indicators_a_matching_level() -> None:
    category = ["A", "A", "A"]
    assert dummy_column_mirror(category, "A") == [1, 1, 1]


def test_dummy_column_is_zero_where_category_differs() -> None:
    category = ["A", "B", "A", "C"]
    assert dummy_column_mirror(category, "A") == [1, 0, 1, 0]


def test_dummy_column_excludes_blank_category_rows_by_default() -> None:
    # Blank cells (None or "") are normalized to "" and excluded by the default
    # mask, the same convention Dummy_Levels/Dummy_Code use — a blank never
    # coerces to a spurious match.
    category = ["A", None, "A", ""]
    assert dummy_column_mirror(category, "A") == [1, "", 1, ""]


def test_dummy_column_respects_an_explicit_include_mask() -> None:
    category = ["A", "A", "A"]
    include = [1, 0, 1]
    assert dummy_column_mirror(category, "A", include) == [1, "", 1]


def test_dummy_column_a_missing_level_is_all_zero_not_an_error() -> None:
    # The free-form primitive imposes no reference validation: a level not in
    # the sample is a valid (if useless) all-0 indicator, NOT #N/A. This
    # deliberately diverges from Dummy_Levels/Dummy_Code's invalid-reference
    # NA() contract — those are treatment-coding functions where a bad
    # reference reintroduces collinearity; an unobserved indicator is harmless.
    category = ["A", "B", "A"]
    result = dummy_column_mirror(category, "Z")
    assert result == [0, 0, 0]
    assert not any(_is_na(v) for v in result)


def test_dummy_column_a_genuine_zero_category_value_passes_through() -> None:
    # IF(category = "", "", category) leaves a real 0 untouched (0 <> "" is
    # TRUE), so a 0 level is a legitimate category, not a blank.
    category = [0, 1, 0]
    assert dummy_column_mirror(category, 0) == [1, 0, 1]


def test_dummy_column_single_level_sample_matches_that_level() -> None:
    category = ["X", "X"]
    assert dummy_column_mirror(category, "X") == [1, 1]
    assert dummy_column_mirror(category, "Y") == [0, 0]


# ── Numeric acceptance: Interact ─────────────────────────────────────────────

def test_interact_elementwise_product_of_two_columns() -> None:
    x1 = [[2.0], [3.0], [5.0]]
    x2 = [[10.0], [20.0], [30.0]]
    assert interact_mirror(x1, x2) == [[20.0], [60.0], [150.0]]


def test_interact_broadcasts_a_column_across_a_dummy_matrix() -> None:
    # n x k dummy matrix times n x 1 continuous -> n x k, one interaction column
    # per retained level (the documented continuous x categorical use).
    x1 = [[1, 0], [0, 1], [1, 0]]  # 3 x 2 dummy
    x2 = [[10.0], [20.0], [30.0]]  # 3 x 1 continuous
    assert interact_mirror(x1, x2) == [[10.0, 0.0], [0.0, 20.0], [30.0, 0.0]]


def test_interact_broadcasts_the_other_direction() -> None:
    # column x matrix -> matrix, same broadcast.
    x1 = [[10.0], [20.0], [30.0]]  # 3 x 1
    x2 = [[1, 0], [0, 1], [1, 0]]  # 3 x 2
    assert interact_mirror(x1, x2) == [[10.0, 0.0], [0.0, 20.0], [30.0, 0.0]]


def test_interact_self_interaction_yields_the_square() -> None:
    # x . x = x^2 — the documented way to get a quadratic term (DECISIONS.md:
    # self-interaction is allowed and documented).
    x = [[2.0], [3.0], [4.0]]
    assert interact_mirror(x, x) == [[4.0], [9.0], [16.0]]


def test_interact_passes_blank_through_from_either_operand() -> None:
    # The "" excluded-row sentinel in either operand -> "" in the output, so
    # Interact composes with transform outputs (Dummy_Code, Ln_Positive) and
    # stays row-aligned. A bare x1*x2 would turn "" into #VALUE!.
    x1 = [[2.0], [BLANK], [5.0]]
    x2 = [[10.0], [20.0], [30.0]]
    assert interact_mirror(x1, x2) == [[20.0], [BLANK], [150.0]]

    x1b = [[2.0], [3.0], [5.0]]
    x2b = [[10.0], [BLANK], [30.0]]
    assert interact_mirror(x1b, x2b) == [[20.0], [BLANK], [150.0]]


def test_interact_propagates_na_from_an_operand() -> None:
    # An included-but-incomputable value (a #N/A from a Log of a non-positive
    # number, say) surfaces as #N/A in the interaction too — the NA()-exception
    # convention is respected; errors are never silently coerced to 0.
    x1 = [[2.0], [NA], [5.0]]
    x2 = [[10.0], [20.0], [30.0]]
    result = interact_mirror(x1, x2)
    assert math.isclose(result[0][0], 20.0)
    assert _is_na(result[1][0])
    assert math.isclose(result[2][0], 150.0)


def test_interact_text_times_number_is_an_error() -> None:
    # Non-numeric text x number -> Excel #VALUE!, mirrored as NA.
    x1 = [["abc"], [2.0]]
    x2 = [[10.0], [20.0]]
    result = interact_mirror(x1, x2)
    assert _is_na(result[0][0])
    assert math.isclose(result[1][0], 40.0)


def test_interact_bool_operand_coerces_to_one_or_zero() -> None:
    # Excel coerces TRUE/FALSE to 1/0 in arithmetic (TRUE*5 = 5).
    x1 = [[True], [False], [True]]
    x2 = [[5.0], [5.0], [5.0]]
    assert interact_mirror(x1, x2) == [[5.0], [0.0], [5.0]]


def test_interact_non_conformable_shapes_error() -> None:
    # n x k x n x m with k != m and neither 1 -> #VALUE!.
    x1 = [[1, 0, 0], [0, 1, 0]]  # 2 x 3
    x2 = [[1, 0], [0, 1]]  # 2 x 2
    assert _is_na(interact_mirror(x1, x2))


# ── Numeric acceptance: Model_Matrix ────────────────────────────────────────

def test_model_matrix_prepends_an_intercept_by_default() -> None:
    X = [[2.0, 3.0], [4.0, 5.0], [6.0, 7.0]]
    assert model_matrix_mirror(X) == [[1, 2.0, 3.0], [1, 4.0, 5.0], [1, 6.0, 7.0]]


def test_model_matrix_explicit_true_prepends_ones() -> None:
    X = [[2.0], [3.0]]
    assert model_matrix_mirror(X, True) == [[1, 2.0], [1, 3.0]]


def test_model_matrix_false_returns_x_unchanged() -> None:
    X = [[2.0, 3.0], [4.0, 5.0]]
    assert model_matrix_mirror(X, False) == [[2.0, 3.0], [4.0, 5.0]]


def test_model_matrix_intercept_is_one_for_every_row() -> None:
    X = [[10.0], [20.0], [30.0]]
    result = model_matrix_mirror(X)
    assert [row[0] for row in result] == [1, 1, 1]


def test_model_matrix_handles_a_single_predictor() -> None:
    X = [[5.0], [6.0]]
    assert model_matrix_mirror(X) == [[1, 5.0], [1, 6.0]]


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


def test_the_trio_are_workbook_scoped_categorical_model_construction_functions() -> None:
    functions = _catalog_functions()
    for name in ("Dummy_Column", "Interact", "Model_Matrix"):
        fn = functions[name]
        assert fn.get("category") == "Data Transformation"
        assert fn.get("subcategory") == "Categorical & Model Construction"
        assert fn.get("scope", "workbook") == "workbook"  # workbook-scoped, like Dummy_Levels/Dummy_Code


def test_dummy_column_shape() -> None:
    fn = _catalog_functions()["Dummy_Column"]
    assert [(a["name"], a["optional"]) for a in fn["arguments"]] == [
        ("category", False), ("level", False), ("include", True)
    ]
    formula = _formula("Dummy_Column")
    assert "--(x=level)" in formula
    assert 'TOCOL(IF(category="","",category),0)' in formula  # blank normalization
    # No error path: a missing level is an all-0 column, not #N/A. The free-form
    # primitive deliberately diverges from Dummy_Levels/Dummy_Code's NA() contract.
    assert "NA()" not in formula
    # Standalone primitive — does not delegate level determination to Dummy_Levels.
    assert "Dummy_Levels(" not in formula


def test_interact_shape() -> None:
    fn = _catalog_functions()["Interact"]
    # No [include]: the operands carry their own row alignment.
    assert [(a["name"], a["optional"]) for a in fn["arguments"]] == [
        ("x1", False), ("x2", False)
    ]
    formula = _formula("Interact")
    assert "x1*x2" in formula  # the elementwise product
    assert '(x1="")+(x2="")' in formula  # the "" passthrough guard
    # Not an error-returning function; errors propagate through x1*x2, they are
    # not generated here.
    assert "NA()" not in formula
    assert "[include]" not in formula and "include" not in formula


def test_model_matrix_shape() -> None:
    fn = _catalog_functions()["Model_Matrix"]
    assert [(a["name"], a["optional"]) for a in fn["arguments"]] == [
        ("X", False), ("add_intercept", True)
    ]
    # Not variadic: exactly two arguments (X + add_intercept). Predictors are
    # assembled explicitly with HSTACK before being passed in.
    assert len(fn["arguments"]) == 2
    formula = _formula("Model_Matrix")
    assert "HSTACK(SEQUENCE(ROWS(X),1,1,0),X)" in formula  # the ones column prepend
    assert "ISOMITTED(add_intercept)" in formula
    # Intercept defaults ON (the common regression case).
    assert "ISOMITTED(add_intercept),TRUE,add_intercept)" in formula
    assert "NA()" not in formula


def test_the_trio_appear_in_architecture_section_order() -> None:
    # ARCHITECTURE.md §5 lists the subcategory as Dummy_Levels, Dummy_Code,
    # Dummy_Column, Interact, Model_Matrix. The catalog document order matches.
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    names = [fn["name"] for fn in payload["functions"]]
    assert names.index("Dummy_Levels") < names.index("Dummy_Code")
    assert names.index("Dummy_Code") < names.index("Dummy_Column")
    assert names.index("Dummy_Column") < names.index("Interact")
    assert names.index("Interact") < names.index("Model_Matrix")


def test_none_of_the_trio_calls_linest() -> None:
    # Guard the whole-catalog invariant in test_intercept_relocation.py
    # (test_every_linest_call_disables_its_own_intercept pins the LINEST-caller
    # set to {Coefficients, SE_Coefficients, SS_Residual}). None of these three
    # is a fitter, so none calls LINEST.
    for name in ("Dummy_Column", "Interact", "Model_Matrix"):
        assert "LINEST(" not in _formula(name)


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