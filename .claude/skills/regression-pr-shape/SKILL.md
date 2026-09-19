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
Building one is a procedure in its own right — follow the **`test-model-case-checklist`** skill for the eight steps (plan-of-record row first, sheet identity, registration, name pinning, the non-default-dataset trap, the coverage-matrix flip). Two facts belong in the PR-shape context, because they are about *PR completeness* rather than how-to:

- The case must be a `RegressionSpecCase` (or `GuardStateCase` for a deliberately-erroring configuration) — not a sheet fixture.
- Its exact sheet name goes into `_EXPECTED_CASE_NAMES` (or `_EXPECTED_GUARD_NAMES`) **in the same commit** — that pin is what stops a case being silently added, renamed, reordered, or dropped.

## 4. A committed excel-only-runs transcript
Because Layer 2 (Excel-based spec-driven) verification cannot run in this container or in CI, the only evidence a human or future agent has that the build+verify actually succeeded on a machine with Excel is a transcript in `excel-only-runs/`, **committed via `git add`** (not pasted into the PR description). Run the build and verify on a machine with desktop Excel, then commit the transcript for the new/changed sheet.

## Before claiming the PR is ready
Walk through all four items explicitly. If you cannot run item 4 in the current environment (no Excel available), say so plainly instead of claiming verification passed — see the `deep-verify-before-done` skill.
