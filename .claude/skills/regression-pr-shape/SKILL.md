---
name: regression-pr-shape
description: Checklist for the four-rule PR shape every Regression feature PR must satisfy in this repo (feature, oracle, test-model case, excel-only-runs transcript). Use before finishing or opening a PR that touches lambda_catalog/analyze_regression*.py, write_sheet_regression*.py, the Regression spec block, or lambda_functions.json entries used by the Regression sheet.
---

# Regression PR shape

This repo's `CONTRIBUTING.md` ("The PR-shape rules — what every Regression PR must contain") requires **all four** of the following in the same PR, not just the code change. A PR is not done until all four are checked off, or item 4 fails against a *tracked, already-open* issue (never a new failure you just introduced).

## 1. The feature itself
The actual behavior change — catalog entry, spec-block column, sheet-writer logic, etc.

## 2. An oracle in the spec-driven verifier chain
A new or extended branch in the independent NumPy/statsmodels comparison path (`calculate_regression_spec_case` in `lambda_catalog/analyze_regression_spec.py`, or the guard-state equivalent in `analyze_regression_guard_states.py`). **Reading the cell back and asserting it equals itself is not an oracle** — the comparison must be against an independently computed expected value.

## 3. A registered, pinned test-model case
- The case is a `RegressionSpecCase` (or `GuardStateCase` for a deliberately-erroring configuration) — not a sheet fixture.
- Check `docs/MODEL_TESTING_ASSETS.md` first: if the corner isn't listed, add it to that document before writing the case.
- Register it in `build_regression_spec_cases()`.
- Give it a sheet identity in `_CASE_SHEET_IDENTITY` (≤31 chars, legal Excel sheet-name charset, `<PlanID> <Concept>` — name the *concept under test*, e.g. `M05 Log-Log NA Masking`, never the variables).
- Add the exact name to `_EXPECTED_CASE_NAMES` (or `_EXPECTED_GUARD_NAMES`) **in the same commit** — this is what stops a case being silently added, renamed, reordered, or dropped.
- If the case uses a non-default dataset, set `source_csv_path`, `row_loader`, and `source_table_ref` **together** — `Source_Table` is the one name that retargets the data sheet; omitting it lands spec rows on the wrong columns silently.
- Flip the row in `docs/MODEL_TESTING_ASSETS.md` §1.5's coverage matrix from new → existing.

## 4. A committed excel-only-runs transcript
Because Layer 2 (Excel-based spec-driven) verification cannot run in this container or in CI, the only evidence a human or future agent has that the build+verify actually succeeded on a machine with Excel is a transcript in `excel-only-runs/`, **committed via `git add`** (not pasted into the PR description). Run the build and verify on a machine with desktop Excel, then commit the transcript for the new/changed sheet.

## Before claiming the PR is ready
Walk through all four items explicitly. If you cannot run item 4 in the current environment (no Excel available), say so plainly instead of claiming verification passed — see the `deep-verify-before-done` skill.
