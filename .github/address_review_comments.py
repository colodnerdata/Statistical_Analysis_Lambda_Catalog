from pathlib import Path


def replace_exact(path: str, old: str, new: str, expected: int = 1) -> None:
    p = Path(path)
    text = p.read_text(encoding="utf-8")
    count = text.count(old)
    if count != expected:
        raise RuntimeError(
            f"{path}: expected {expected} occurrences, found {count}: {old!r}"
        )
    p.write_text(text.replace(old, new), encoding="utf-8")


# Guard-state oracle: match the sheet's singular/plural amber status exactly.
replace_exact(
    "lambda_catalog/analyze_regression_guard_states.py",
    '    return f"{dropped} rows excluded: Log of ≤ 0" if dropped else ""\n',
    '    if dropped == 0:\n'
    '        return ""\n'
    '    suffix = "row" if dropped == 1 else "rows"\n'
    '    return f"{dropped} {suffix} excluded: Log of ≤ 0"\n',
)

# Catalog-schema dependency-order comment: describe the new two-mask contract,
# not the removed optional FALSE-argument reader behavior.
replace_exact(
    "tests/test_catalog_schema.py",
    '''                # The v3.2 name-promotion: the _Calc computational leaf precedes\n                # the public reader that delegates to it. Sample_Include_Calc is\n                # the REDUCE leaf the spill cell (=Log_Drop_Log_Drop_Sample_Include_Calc()) calls;\n                # Sample_Include is the reader over that spill (and delegates\n                # FALSE to Sample_Include_Calc). Document order IS dependency\n                # order (see _set_sheet_scoped_names).\n''',
    '''                # Two-mask contract: Sample_Include_Calc computes ordinary\n                # eligibility and public Sample_Include delegates to it. The\n                # later Log_Drop_Sample_Include_Calc adds only the explicit\n                # Log (drop ≤ 0) positivity layer and is the final-mask spill\n                # source read by Fit_Sample_Include(). Document order IS\n                # dependency order (see _set_sheet_scoped_names).\n''',
)

# N()-coercion guard docstring: neither public mask reader has a FALSE argument
# anymore; the test rejects N() over either array/range-returning thunk.
replace_exact(
    "tests/test_sheet_writers.py",
    '''    N() of a range/array-returning thunk (Fit_Sample_Include(), or the no-arg\n    Sample_Include() which on main returns that same reader) collapses to the\n    top-left cell, so SUMPRODUCT(N(<reader>())) returns 1 for any non-empty\n    sample. That is the Log_Domain_Status amber bug (PR #237) and the\n    Intercept_Only_N bug this test sits next to: both made a per-sheet count\n    read 1 instead of the included-row count.\n\n    No import can reach a JSON string literal or a Python RefersTo string, so a\n    source scan is the only thing that catches a future call site retaining the\n    pattern. The guard sweeps every catalog body AND every RefersTo the\n    Regression sheet-writer registers (the constructor closures AND the\n    local-only names like Intercept_Only_N), so neither half can regress alone.\n    N(Sample_Include()) / N(base) are NOT defects — the FALSE arg returns\n    an array leaf, which N() sums correctly — so only the no-arg reader forms\n    are rejected.\n''',
    '''    N() of an array/range-returning thunk such as Sample_Include() or\n    Fit_Sample_Include() can collapse to the top-left cell, so\n    SUMPRODUCT(N(<reader>())) may return 1 for any non-empty sample. That is\n    the Log_Domain_Status amber bug (PR #237) and the Intercept_Only_N bug this\n    test sits next to: both made a per-sheet count read 1 instead of the row\n    count represented by the mask.\n\n    No import can reach a JSON string literal or a Python RefersTo string, so a\n    source scan is the only thing that catches a future call site retaining the\n    pattern. The guard sweeps every catalog body AND every RefersTo the\n    Regression sheet-writer registers (the constructor closures AND the\n    local-only names like Intercept_Only_N), so neither half can regress alone.\n    Both no-argument mask thunks are rejected here; callers that need a count\n    must coerce the returned array explicitly with -- rather than N().\n''',
)
