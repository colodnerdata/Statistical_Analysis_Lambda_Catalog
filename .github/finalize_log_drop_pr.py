from pathlib import Path
import json


def replace_exact(path: str, old: str, new: str, expected: int = 1) -> None:
    p = Path(path)
    text = p.read_text(encoding="utf-8")
    count = text.count(old)
    if count != expected:
        raise RuntimeError(f"{path}: expected {expected} occurrences, found {count}: {old!r}")
    p.write_text(text.replace(old, new), encoding="utf-8")


# ---------------------------------------------------------------------------
# Spec-block calls: every display/check that claims to mirror the actual
# constructed model must use the final fitted sample, not ordinary eligibility.
# ---------------------------------------------------------------------------
path = Path("lambda_catalog/write_spec_block.py")
text = path.read_text(encoding="utf-8")

# K Levels, L Reference In Use, O Design Columns all mirror constructor state.
if text.count('"si,Sample_Include(),"') != 3:
    raise RuntimeError("expected exactly three fit-dependent spec-display masks")
text = text.replace('"si,Sample_Include(),"', '"si,Fit_Sample_Include(),"')

# Strict Log red CF must agree with G2 RED: only a bad value that actually
# reaches the fitted sample poisons the fit.
if text.count('"SUMPRODUCT(--Sample_Include(),"') != 1:
    raise RuntimeError("strict-Log CF mask occurrence changed")
text = text.replace(
    '"SUMPRODUCT(--Sample_Include(),"',
    '"SUMPRODUCT(--Fit_Sample_Include(),"',
)

# Reference validity must be checked against the same categorical level set
# the constructor uses after intentional Log-drop row removal.
old_ref = 'f"$E{_FIRST_DATA_ROW},Sample_Include())))"'
new_ref = 'f"$E{_FIRST_DATA_ROW},Fit_Sample_Include())))"'
if text.count(old_ref) != 1:
    raise RuntimeError("invalid-reference CF mask occurrence changed")
text = text.replace(old_ref, new_ref)

# Current architecture documentation in the source file.
text = text.replace(
    "    S             T        U     V           W            X     Y           Z →\n"
    "    Δ    Count          Row Labels    Included (brk) Filt.Labels Filt.y       (brk) Filt.Labels Filtered Predictor_Columns\n"
    "    (=Row_Labels() spill at S4; =Sample_Include() spill at T4 — both\n"
    "     full-height, never internally filtered. V/W/Y/Z are the FILTERED\n"
    "     display zones: the only place on the sheet where Sample_Include()\n"
    "     row-filters anything. Y repeats the filtered labels so the matrix\n",
    "    S             T        U     V           W            X     Y           Z →\n"
    "    Δ    Count          Row Labels    Eligible (brk) Filt.Labels Filt.y       (brk) Filt.Labels Filtered Predictor_Columns\n"
    "    (=Row_Labels() spill at S4; =Sample_Include() spill at T4 — both\n"
    "     full-height, never internally filtered. T is ordinary pre-Log-drop\n"
    "     eligibility. V/W/Y/Z are the FILTERED display zones and use\n"
    "     Fit_Sample_Include(), the actual fitted mask. Y repeats the filtered\n"
    "     labels so the matrix\n",
)
text = text.replace(
    "    included rows = SUMPRODUCT(N(Sample_Include())) ·\n",
    "    included rows = SUMPRODUCT(--Fit_Sample_Include()) ·\n",
)
text = text.replace(
    "SUMPRODUCT(N(Sample_Include())) = 392 (completeness-only on the response\n",
    "SUMPRODUCT(--Fit_Sample_Include()) = 392 in the shipped no-Log-drop default (completeness-only on the response\n",
)
text = text.replace(
    "# count L over the mask-included rows, with Dummy_Levels' blank\n",
    "# count L over the fitted rows, with Dummy_Levels' blank\n",
)
text = text.replace(
    "# shows Dummy_Levels' own default, the first sorted level over the\n# mask-included sample.",
    "# shows Dummy_Levels' own default, the first sorted level over the\n# fitted sample.",
)
text = text.replace(
    "# Everything invariant across rows is hoisted into the outer LET —\n# Sample_Include(), TOROW(Header_Names) and the kk helper itself.",
    "# Everything invariant across rows is hoisted into the outer LET —\n# Fit_Sample_Include(), TOROW(Header_Names) and the kk helper itself.",
)
text = text.replace(
    "# The row test mirrors Sample_Include's own eligibility branch exactly:\n",
    "# The row test mirrors the transform eligibility branch exactly:\n",
)
text = text.replace(
    "# Sample_Include() is the mask BEFORE the positivity layer, which is\n# what makes this count the rows the fit would otherwise have used.\n",
    "# Fit_Sample_Include() is the final mask AFTER any explicitly declared\n# Log-drop exclusions, so this flags only strict-Log values that can actually\n# reach Ln_Positive in the fitted model.\n",
)
text = text.replace(
    "# Dummy_Levels fail — the constructor's exact skip condition, tested\n",
    "# Dummy_Levels fail on the fitted sample — the constructor's exact skip condition, tested\n",
)
path.write_text(text, encoding="utf-8")


# ---------------------------------------------------------------------------
# Materialization documentation: public Sample_Include is no longer a reader.
# ---------------------------------------------------------------------------
path = Path("lambda_catalog/regression_materialization.py")
text = path.read_text(encoding="utf-8")
old = '''    **The v3.2 name-promotion landed here, via a ``_Calc`` split.** The
    spill-source cells call the computational leaves ``=Sample_Include_Calc()``
    and ``=Design_Columns_Calc()`` (catalog LAMBDAs holding the REDUCE bodies),
    NOT the public ``Sample_Include()`` / ``Design_Columns()`` names. The public
    names are READERS over these spills — ``Sample_Include`` dispatches to
    ``Fit_Sample_Include()`` for the default and to ``Sample_Include_Calc(FALSE)``
    for the pre-positivity mask; ``Design_Columns`` delegates to
    ``Fit_Design_Columns()``. Producing the spill from the ``_Calc`` leaf is what
    breaks the self-reference that kept the promotion deferred: the producing
    cell no longer calls the name that reads its own spill. ``Sample_Include``
    keeps its ordinary pre-drop eligibility semantics because
    ``Sample_Include()`` — the mask BEFORE the positivity layer that
    ``Log_Domain_Status`` differences against the default — still delegates to
    ``Sample_Include_Calc(FALSE)``; only the default mask is materialized here.
'''
new = '''    The sample-mask surface now makes the Log-drop distinction explicit.
    ``Sample_Include_Calc()`` computes ordinary eligibility only (Filters plus
    numeric completeness), and public ``Sample_Include()`` delegates directly
    to that ordinary mask. ``Log_Drop_Sample_Include_Calc()`` starts from it and
    adds positivity only where the spec declares exactly ``Log (drop ≤ 0)``.
    THAT final mask is what the spill-source cell materializes here, and
    ``Fit_Sample_Include()`` is the sheet-scoped reader over the materialized
    spill. Strict ``Log`` therefore never filters a row; a surviving non-positive
    value reaches ``Ln_Positive`` and fails visibly. ``Design_Columns_Calc()``
    remains the design-matrix spill source and public ``Design_Columns()`` reads
    it through ``Fit_Design_Columns()``.
'''
if old not in text:
    raise RuntimeError("stale materialization architecture paragraph not found")
text = text.replace(old, new)
old = '''    # rows beside it. The spill-source cell calls the _Calc computational leaf
    # (=Log_Drop_Sample_Include_Calc()), while public Sample_Include() remains the ordinary pre-drop eligibility mask. Before the
    # v3.2 promotion the spill cell WAS =Sample_Include(); pointing the public
    # name — now a reader over THIS spill via Fit_Sample_Include() — at a cell
    # that called itself would have been self-referential. The _Calc split is
    # what breaks that cycle: the producing cell calls the leaf, so the public
    # reader can read the spill without self-reference (see the docstring).
'''
new = '''    # rows beside it. The spill-source cell calls the explicitly Log-drop-aware
    # computational leaf (=Log_Drop_Sample_Include_Calc()). Public
    # Sample_Include() is deliberately different: it returns ordinary eligibility
    # before transform-driven dropping. Fit_Sample_Include() is the reader over
    # THIS final fitted-mask spill (see the docstring).
'''
if old not in text:
    raise RuntimeError("stale materialization inline comment not found")
text = text.replace(old, new)
text = text.replace(
    '            "block (Include, and any Role=Filter column), not here. Full "',
    '            "block (Include, Role=Filter, or Transform=Log (drop ≤ 0)), not here. Full "',
)
path.write_text(text, encoding="utf-8")


# ---------------------------------------------------------------------------
# Python mirrors: behavior was already correct; make the documentation name
# ordinary versus final masks accurately.
# ---------------------------------------------------------------------------
path = Path("lambda_catalog/analyze_model_construction.py")
text = path.read_text(encoding="utf-8")
text = text.replace(
    '    """Mirror ``Sample_Include()``: AND over role-derived row conditions.\n',
    '    """Compute the model row mask; by default include the explicit Log-drop layer.\n',
)
text = text.replace(
    '    ``apply_log_drop=False`` mirrors ``Sample_Include() before Log drop``: the mask\n'
    '    before layer 3, which is what the sheet\'s G2 status cell differences\n'
    '    against the default to report how many rows the transform excluded.\n',
    '    ``apply_log_drop=False`` mirrors public ``Sample_Include()``: ordinary\n'
    '    eligibility before layer 3. The default mirrors the final fitted mask\n'
    '    produced by ``Log_Drop_Sample_Include_Calc()`` / read by\n'
    '    ``Fit_Sample_Include()``. Their population difference is the number of\n'
    '    distinct rows intentionally removed by Log (drop ≤ 0).\n',
)
path.write_text(text, encoding="utf-8")

path = Path("lambda_catalog/analyze_regression_guard_states.py")
text = path.read_text(encoding="utf-8")
old = '''    """Every strict-``Log`` column, with its count of non-positive fit rows.

    "Fit rows" means the mask BEFORE the positivity layer — the rows the model
    would otherwise have used — which is what ``Sample_Include() before Log drop`` gives
    the sheet. Only columns that actually reach ``Ln_Positive`` are considered:
    the Response and included Continuous Predictors, exactly the eligibility
    branch ``_compute_mask`` uses. Columns with a zero count are included so
    callers can tell "checked, clean" from "not checked".
    """
'''
new = '''    """Every strict-``Log`` column, with its count of non-positive fitted rows.

    "Fitted rows" means the FINAL mask after any explicitly declared
    ``Log (drop ≤ 0)`` exclusions. A strict-Log non-positive on a row already
    removed by another Log-drop variable cannot poison the fit and therefore is
    not counted. Only the Response and included Continuous Predictors can reach
    ``Ln_Positive``. Columns with a zero count are included so callers can tell
    "checked, clean" from "not checked".
    """
'''
if old not in text:
    raise RuntimeError("guard-state strict-log doc block not found")
text = text.replace(old, new)
path.write_text(text, encoding="utf-8")

path = Path("lambda_catalog/analyze_regression_spec.py")
text = path.read_text(encoding="utf-8")
text = text.replace(
    "    # non-positive, for the same reason the cell differences two calls to\n"
    "    # Sample_Include: one predicate, two evaluations, nothing to drift.\n",
    "    # non-positive. Compare ordinary eligibility against the final Log-drop\n"
    "    # mask: one eligibility predicate plus one explicit transform layer,\n"
    "    # matching Sample_Include() versus Fit_Sample_Include() on the sheet.\n",
)
path.write_text(text, encoding="utf-8")


# ---------------------------------------------------------------------------
# Catalog prose cleanup where the formulas were already rewired.
# ---------------------------------------------------------------------------
catalog = Path("lambda_functions.json")
document = json.loads(catalog.read_text(encoding="utf-8"))
by_name = {item["name"]: item for item in document["functions"]}
predictor = by_name["Predictor_Columns"]
predictor["description"] = predictor["description"].replace(
    "Reads Sample_Include only to fix categorical level sets and to gate the Log transform's row mask;",
    "Reads Log_Drop_Sample_Include_Calc only to fix categorical level sets and to gate the Log transform's row mask;",
)
catalog.write_text(json.dumps(document, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


# ---------------------------------------------------------------------------
# Contract tests for the fit-dependent spec-block displays / CFs.
# ---------------------------------------------------------------------------
path = Path("tests/test_spec_block_writer.py")
text = path.read_text(encoding="utf-8")
text = text.replace(
    '        f"$E{r},Sample_Include())))"\n',
    '        f"$E{r},Fit_Sample_Include())))"\n',
)
# Existing strict-Log CF assertion if present.
text = text.replace(
    '"SUMPRODUCT(--Sample_Include(),"',
    '"SUMPRODUCT(--Fit_Sample_Include(),"',
)
path.write_text(text, encoding="utf-8")

# Static source-level tripwire for the three computed display masks.
path = Path("tests/test_sheet_writers.py")
text = path.read_text(encoding="utf-8")
anchor = "def test_"
addition = '''\n\ndef test_spec_fit_dependent_displays_use_materialized_fit_mask() -> None:\n    from lambda_catalog.write_spec_block import (\n        _DESIGN_COLUMNS_SPILL_FORMULA,\n        _LEVELS_SPILL_FORMULA,\n        _REF_IN_USE_SPILL_FORMULA,\n    )\n\n    for formula in (\n        _LEVELS_SPILL_FORMULA,\n        _REF_IN_USE_SPILL_FORMULA,\n        _DESIGN_COLUMNS_SPILL_FORMULA,\n    ):\n        assert "Fit_Sample_Include()" in formula\n        assert "Sample_Include()" not in formula.replace("Fit_Sample_Include()", "")\n'''
if "test_spec_fit_dependent_displays_use_materialized_fit_mask" not in text:
    text += addition
path.write_text(text, encoding="utf-8")
