"""Independent verification of Ln_Positive (v2.2 Transform=Log MVP).

Ln_Positive is the primitive behind the Regression sheet's spec column G
(Transform): Response_Column(), X_s(), and Constructed_Column_Names() all
call it when a row declares Log. Verified here the same way the other
transform-family functions are: a pure-Python mirror cross-checked against
math.log, plus implementation-shape assertions on the actual catalog
formula text so a future edit cannot silently change the contract without
a test noticing.

Runnable standalone (``python tests/test_ln_positive_verification.py``) or
under pytest.
"""
from __future__ import annotations

from pathlib import Path

import json
import math
import sys

import numpy as np

if __package__ in (None, ""):  # standalone run: make `tests.` importable
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

NA = float("nan")
BLANK = ""


# ── Pure-Python mirror of the workbook formula ───────────────────────────────

def ln_positive_mirror(x, include=None):
    """Mirror of Ln_Positive(x, [include]).

    Excluded row -> "" (BLANK). Included row: LN(x) when x is numeric and
    strictly positive; NA() when x is zero, negative, or non-numeric —
    the library's NA()-exception convention (BLANK = excluded, NA =
    included-but-incomputable), the same one Lag_By/Difference_By use.
    Default mask (include omitted) is ISNUMBER(x).
    """
    x = list(x)
    n = len(x)
    if include is None:
        inc = [isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)) for v in x]
    else:
        inc = [bool(v) for v in include]

    out = []
    for i in range(n):
        if not inc[i]:
            out.append(BLANK)
            continue
        v = x[i]
        is_number = isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v))
        if is_number and v > 0:
            out.append(math.log(v))
        else:
            out.append(NA)
    return out


def _is_na(value) -> bool:
    return isinstance(value, float) and math.isnan(value)


# ── Numeric acceptance ────────────────────────────────────────────────────────

def test_ln_positive_matches_log_for_positive_included_values() -> None:
    x = [1.0, math.e, 10.0, 100.0]
    mirrored = ln_positive_mirror(x)
    assert np.allclose(mirrored, np.log(x))


def test_ln_positive_is_na_for_zero_or_negative_included_values() -> None:
    x = [0.0, -5.0, 3.0]
    mirrored = ln_positive_mirror(x)
    assert _is_na(mirrored[0])
    assert _is_na(mirrored[1])
    assert not _is_na(mirrored[2])
    assert math.isclose(mirrored[2], math.log(3.0))


def test_ln_positive_is_blank_for_excluded_rows() -> None:
    x = [10.0, 20.0]
    include = [True, False]
    mirrored = ln_positive_mirror(x, include)
    assert math.isclose(mirrored[0], math.log(10.0))
    assert mirrored[1] == BLANK


def test_ln_positive_default_mask_excludes_non_numeric_rather_than_erroring() -> None:
    # No explicit include: the default mask is ISNUMBER(x), so a blank/
    # non-numeric cell is EXCLUDED ("") rather than flagged NA() — NA() is
    # reserved for a row explicitly included despite being incomputable.
    x = [5.0, NA, 8.0]
    mirrored = ln_positive_mirror(x)
    assert math.isclose(mirrored[0], math.log(5.0))
    assert mirrored[1] == BLANK
    assert math.isclose(mirrored[2], math.log(8.0))


def test_ln_positive_is_na_for_an_explicitly_included_non_numeric_row() -> None:
    # Forced inclusion of a non-numeric row is the "included but
    # incomputable" case -> NA(), not "".
    x = [NA]
    include = [True]
    mirrored = ln_positive_mirror(x, include)
    assert _is_na(mirrored[0])


def test_ln_positive_geometric_mean_round_trip() -> None:
    # The Prediction Inputs band's double-log fix (write_sheet_regression.py
    # AI19) relies on this exact identity: EXP(mean(ln x)) fed back through
    # Ln_Positive recovers mean(ln x), so the default prediction still
    # lands on X_s()'s own centroid.
    x = [50.0, 75.0, 120.0, 30.0]
    logged = np.log(x)
    geometric_mean_input = math.exp(float(np.mean(logged)))
    recovered = ln_positive_mirror([geometric_mean_input])[0]
    assert math.isclose(recovered, float(np.mean(logged)), rel_tol=1e-12)


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


def test_ln_positive_is_a_workbook_scoped_data_transformation_function() -> None:
    fn = _catalog_functions()["Ln_Positive"]
    assert fn.get("category") == "Data Transformation"
    assert fn.get("subcategory") == "Location & Scale"
    assert "scope" not in fn  # workbook-scoped, like Dummy_Levels/Demean_By
    assert fn.get("test_table") is None


def test_ln_positive_is_elementwise_not_an_aggregate_and() -> None:
    formula = _formula("Ln_Positive")
    # AND() aggregates an array to one scalar — the exact class of bug the
    # X_s() degeneracy-guard comment warns about, in reverse: this function
    # must stay elementwise so it broadcasts over a full column.
    assert "AND(" not in formula
    assert "IF(inc=0,\"\",IF(pos,LN(x_v),NA()))" in formula


def test_ln_positive_default_mask_is_isnumber() -> None:
    formula = _formula("Ln_Positive")
    assert "IF(ISOMITTED(include),ISNUMBER(x_v),include*1)" in formula


def test_ln_positive_appears_before_dummy_levels_in_document_order() -> None:
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    names = [fn["name"] for fn in payload["functions"]]
    assert names.index("Ln_Positive") < names.index("Dummy_Levels")


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
