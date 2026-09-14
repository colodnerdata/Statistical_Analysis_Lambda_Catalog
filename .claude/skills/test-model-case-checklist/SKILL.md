---
name: test-model-case-checklist
description: Step-by-step checklist for adding or editing a RegressionSpecCase or GuardStateCase in the QC test-model suite. Use before writing a new spec-case builder, before registering a case in build_regression_spec_cases(), or whenever a change touches _EXPECTED_CASE_NAMES / _EXPECTED_GUARD_NAMES / _CASE_SHEET_IDENTITY.
---

# Adding a test-model case

`docs/MODEL_TESTING_ASSETS.md` is the **plan of record** for which configurations the QC harness covers. Read it before adding or changing a case; add to it before adding a case it doesn't list. This repo runs a *covering array* (~25–30 fittable models + ~10 guard states), not a full factorial — don't add a case that doesn't buy a genuinely new corner.

## The 8 steps

1. **Read `docs/MODEL_TESTING_ASSETS.md`** — confirm the corner you're covering is listed (§1.5 coverage matrix or §2 roadmap ladder). If not, add a row there first, in the same PR.
2. **Write the spec builder** as a `RegressionSpecCase` in `lambda_catalog/analyze_regression_spec.py` (or a `GuardStateCase` in `analyze_regression_guard_states.py` if the configuration is meant to raise/fail by design). Give it a docstring naming the specific corner it covers — not a restatement of its variables.
3. **Register the case** in `build_regression_spec_cases()` (or the guard-state registry equivalent).
4. **Add sheet identity** to `_CASE_SHEET_IDENTITY`: name format is `<PlanID> <Concept>`, ≤31 characters, legal Excel sheet-name charset, and unique across both model and guard cases. Name the **concept under test** (`M05 Log-Log NA Masking`), never the variables (`MPG ~ Ln(Weight) + Ln(HP)`) — `lambda_catalog/test_model_sheets.py` validates this at registry-build time.
5. **Pin the name** — add the exact sheet name to `_EXPECTED_CASE_NAMES` in `tests/test_regression_spec_qc.py` (or `_EXPECTED_GUARD_NAMES` in `tests/test_regression_guard_states.py`) **in the same commit**. This is the mechanism that turns a silent add/rename/reorder/drop into a failing test.
6. **Non-default dataset trap** — if this case doesn't use the default source table, set `source_csv_path`, `row_loader`, and `source_table_ref` **all together**. `Source_Table` is the one name that retargets the data sheet; setting only one or two of the three lands spec rows on the wrong columns with no error.
7. **Run the suite in order**: pure-Python unit tests first (`uv run pytest`), then the Excel-machine Layer 2 verify (needs desktop Excel — see `deep-verify-before-done`). Don't skip straight to claiming success from the Python suite alone; it doesn't exercise the actual workbook.
8. **Update `docs/MODEL_TESTING_ASSETS.md` §1.5** — flip the coverage-matrix row's status from new → existing.

## Guard states are different from fittable cases

A `GuardStateCase` is expected to raise in the spec oracle *by design* — it asserts status text, the Design Columns audit, and that the right conditional-formatting rule fires, never fit statistics. `GuardFlag` is a predicate you recompute (the rule condition), never a pixel — don't assert against `DisplayFormat.Interior.Color`.
