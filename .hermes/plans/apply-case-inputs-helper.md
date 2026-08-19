# Part 0 fix now; apply-case-inputs helper still planned

**Branch (plan):** `investigate/verify-baseline-failures` (this document)
**Branch (implementation):** `refactor/apply-case-inputs-helper` (off `main`, not yet created)
**Date:** 2026-08-18
**Status:** Part 0 has already landed on this branch: the surgical BFN fix in
`tools/inspect_regression_sheet.py` plus the source-grep guard test in
`tests/test_inspect_regression_sheet_wires_sequence_period.py`. This document
is the still-planned permanent-prevention follow-up: replace the duplicated
three-step sequence with one helper so the two fittable call sites cannot drift
again.

| Part | State | Lands as |
|---|---|---|
| 0. Surgical BFN fix + source-grep guard (the investigation's output) | **On this branch** | `tools/inspect_regression_sheet.py` + `tests/test_inspect_regression_sheet_wires_sequence_period.py` |
| 1. Add `apply_case_inputs` to `regression_spec_sheet_io.py` | **Planned** | one function + docstring |
| 2. Migrate the test-model builder to it | **Planned** | `write_sheet_test_model.write_test_model_sheet` |
| 3. Migrate the verify inspector to it | **Planned** | `tools/inspect_regression_sheet.read_regression_df` |
| 4. Replace the source-grep guard with a behavioral test of the helper | **Planned** | `tests/test_apply_case_inputs.py` (new); retire `test_inspect_regression_sheet_wires_sequence_period.py` |
| 5. Leave the guard-state path as-is | **Closed — structurally separate** | see "Why the guard path is not migrated" |

## Goal

One function owns the "write a fittable case's visible inputs onto a sheet"
sequence — `apply_spec_case` → typed Sequence Period → `set_prediction_inputs` —
so a second or third call site cannot forget the middle step again. This is
the "bigger change" deferred during the verify-baseline investigation, where
the root cause was identified but not fixed permanently because the
`_padded` asymmetry made a naive extraction non-trivial.

## The bug this prevents

`tools/inspect_regression_sheet.read_regression_df` (the loop behind
`build_production.py --verify`) applied each `RegressionSpecCase` with
`apply_spec_case` + `set_prediction_inputs` but **never typed column I** — the
Sequence *Period*. `apply_spec_case` writes the Sequence *flag* (column H) but
not the typed period (column I). `Base_Period_Delta()` reads the TYPED value
and returns `#N/A` when column I is blank, and the BFN panel Durbin-Watson cell
(AE12) passes that as its delta. With `step` at `#N/A` every `Difference_By`
lookup misses, the `n_terms` guard fires `#N/A`, and AE12 reads `nan` — while
the oracle holds `case.sequence_period` (1.0 on P01/P02) and computes a real
BFN (0.985). The mismatch reads as a broken statistic and is really a blank
input cell the harness never typed.

The test-model **builder** (`write_sheet_test_model.write_test_model_sheet`)
had the wiring correct, with a documenting comment. The **inspector** duplicated
the sequence and dropped the middle step. That duplication is the cause: two
places doing the same three things, one of which drifted. The BFN cell sat at
`nan` for every verify run on record — the gap was never green, which is why
it read as "pre-existing baseline noise" rather than a bug.

## The three call sites (the shape of the problem)

`grep` over `apply_spec_case` / `set_prediction_inputs` /
`apply_sequence_period_overrides` finds exactly three writers of a spec block:

1. **Builder (fittable case)** — `write_sheet_test_model.write_test_model_sheet`
   (lines 287–313). Pads the spec to the source table's width, applies it,
   types column I from the scalar `case.sequence_period` (building the
   `{name: period}` dict from the Sequence-flagged rows), then prediction
   inputs. Correct today.
2. **Inspector (fittable case)** — `tools/inspect_regression_sheet.read_regression_df`
   (lines 134–160). Does **not** pad, applies the spec, [after the surgical
   fix] types column I from `expected.case.sequence_period`, then prediction
   inputs. Correct only after the fix on this branch.
3. **Builder (guard state)** — `write_sheet_test_model._apply_guard_spec`
   (lines 400–403). Applies the spec through a synthetic `_ExpectedView`
   shim, types column I from the pre-built **dict** `case.sequence_period_override`
   (a different field), and writes **no** prediction inputs (a guard case is
   not a fittable model). Correct and structurally different.

The two fittable-case paths (1, 2) are the ones the helper unifies. The guard
path (3) is not — see "Why the guard path is not migrated."

## The `_padded` asymmetry — and why it does not block the helper

The complication that made a naive extraction risky: the **builder pads** the
spec to the source table's width (`_padded(expected)` → `pad_spec_to_source_table`,
which **appends** an `Omit` row for every unnamed source column); the
**inspector does not**. So a helper that owned padding would either change the
inspector's behavior or need a padding flag.

The resolution: **padding is the caller's concern, not the helper's.** The
helper takes one `expected` and uses `expected.case.spec` / `.results` /
`.design` uniformly. The builder passes the **padded** expected; the inspector
passes `expected` directly. This preserves both behaviors exactly, for three
reasons that all reduce to one fact — **padding only appends `Omit` rows at the
end** (`padded = (*spec, *padding)` in `pad_spec_to_source_table`):

* **The Sequence-Period dict is identical.** The dict is built from
  `Sequence`-flagged rows, which are declared variables. Padding adds `Omit`
  rows (never `Sequence`), so the dict is the same whether built from the
  padded or unpadded spec. The helper building it from `expected.case.spec`
  (padded, in the builder's case) yields the same dict as the builder's current
  `{item.name: case.sequence_period for item in case.spec if item.sequence}`.
* **Row offsets are stable.** `apply_sequence_period_overrides` derives row
  positions from `enumerate(spec)`. Appended padding does not shift any
  declared row's offset, so the typed period lands on the same row whether
  the helper reads padded or unpadded positions.
* **`.results` and `.design` are unchanged by padding.** `_padded` does
  `replace(expected, case=replace(case, spec=padded))` — only `.case.spec` is
  replaced. `expected.results` and `expected.design` (which the helper reads
  for prediction inputs) are the same object on the padded and unpadded
  expected.

So the helper is a pure de-duplication: same three writes, same values, one
place. The `Omit`-appends-only invariant is what makes that true and is
asserted by `pad_spec_to_source_table`'s own width check already; this plan
adds no new invariant, it relies on the existing one.

## Part 1 — Add `apply_case_inputs`

In `lambda_catalog/regression_spec_sheet_io.py`, alongside
`apply_spec_case` / `apply_sequence_period_overrides` / `set_prediction_inputs`
(the module exists so "the one function that writes a spec onto a sheet stays
identical for the builder and both verifiers" — same reason the helper lives
here):

```python
def apply_case_inputs(sheet: xw.Sheet, expected: RegressionSpecExpected) -> None:
    """Apply a fittable case's visible inputs: spec, typed Sequence Period, prediction inputs.

    The three writes every spec-driven case needs on top of a sheet whose spec
    block already exists: ``apply_spec_case`` (Source_Table retarget, intercept,
    FE group, back-transform, the spec rows), the typed Sequence Period into
    column I, and ``set_prediction_inputs``. Centralizing the sequence is what
    stops a second call site forgetting the middle step — which is how the BFN
    panel Durbin-Watson cell (AE12) sat at ``nan`` for every verify run on
    record while the oracle held a real number: the inspector duplicated this
    sequence and dropped the override, and the only symptom was a multi-minute
    Excel run that read like a broken diagnostic.

    Padding the spec to the source table's width is the CALLER's concern, not
    this helper's: the builder pads (it writes a fresh sheet sized to the
    dataset and wants every column shown), the inspector does not (it writes
    the fixed production sheet). Pass the padded ``expected`` if you want
    padding — the Sequence-Period dict, row offsets, and prediction inputs are
    all unchanged by appended ``Omit`` padding, so the helper reads the same
    values either way.

    Only cases that declare a period are touched by the override, so non-panel
    configs are unaffected.
    """
    apply_spec_case(sheet, expected)
    case = expected.case
    if case.sequence_period is not None:
        apply_sequence_period_overrides(
            sheet,
            case.spec,
            {item.name: case.sequence_period for item in case.spec if item.sequence},
        )
    set_prediction_inputs(
        sheet,
        expected.results.prediction_interval.pred_input_values,
        expected.design.constructed_column_transforms,
    )
```

`RegressionSpecExpected` is already imported in this module. `xw` is already
imported.

## Part 2 — Migrate the builder

`lambda_catalog/write_sheet_test_model.py`, `write_test_model_sheet`
(lines 287–313) becomes:

```python
    padded = _padded(expected)
    apply_case_inputs(sheet, padded)
```

The block's documenting comment (lines 289–307) is subsumed by the helper's
docstring; delete it. `_write_provenance` (line 314+) is unchanged and stays
after the call.

Import: add `apply_case_inputs` to the
`from lambda_catalog.regression_spec_sheet_io import (...)` block, and drop
the now-unused `apply_spec_case`, `apply_sequence_period_overrides`,
`set_prediction_inputs` from that import **only if no other site in this file
uses them**. `_apply_guard_spec` (line 400) still uses `apply_spec_case` and
`apply_sequence_period_overrides` directly — so those two stay imported;
`set_prediction_inputs` becomes unused here and is dropped.

**Equivalence check (the thing to verify by reading, not by running):** the
builder currently builds the override dict from `case.spec` (unpadded) but
passes `padded.case.spec` as the spec arg. The helper builds the dict from
`expected.case.spec` where `expected` is `padded`, i.e. from `padded.case.spec`.
Because padding only appends `Omit` (never `Sequence`), the dict is identical.
The spec arg is `padded.case.spec` in both. Prediction inputs read
`expected.results` / `expected.design` — `padded` carries the same `.results`
/ `.design` as the unpadded `expected`. Byte-equivalent.

## Part 3 — Migrate the inspector

`tools/inspect_regression_sheet.py`, `read_regression_df` (lines 134–160) —
replace the `apply_spec_case` + conditional `apply_sequence_period_overrides` +
`set_prediction_inputs` block with:

```python
            apply_case_inputs(sheet, expected)
```

The long comment block (lines 135–145) is subsumed by the helper's docstring;
delete it. `_apply_extra_columns` (line 133) stays before the call — it writes
fixture columns to the *data* sheet, not the spec, and is not part of this
sequence. `sheet.api.Calculate()` (line 164) stays after.

Import: replace the `apply_sequence_period_overrides`, `apply_spec_case`,
`set_prediction_inputs` names with `apply_case_inputs` in the
`from lambda_catalog.regression_spec_sheet_io import (...)` block. (`read_case_comparison_rows`
stays.)

This is the whole point: the inspector no longer re-implements the sequence,
so it cannot drift off it again.

## Part 4 — Tests: behavioral, not source-grep

The investigation added `tests/test_inspect_regression_sheet_wires_sequence_period.py`,
a source-grep guard that asserts the inspector *contains* the override call in
the right neighbourhood. After the migration the inspector contains
`apply_case_inputs` — one name — and the override is inside the helper, so the
grep guard no longer applies and is retired.

Replace it with a **behavioral** test of the helper, which is strictly
stronger: it exercises the full composition headlessly via `RecordingSheet`
(the same mock the spec-block tests use — `apply_spec_case`,
`apply_sequence_period_overrides`, and `set_prediction_inputs` are all plain
cell writes with no recalculation). New file
`tests/test_apply_case_inputs.py`:

* **`test_apply_case_inputs_types_column_i_and_writes_prediction_inputs`** —
  build the P01 expected (`production_lots_fixed_effects`, the only case with
  `sequence_period == 1.0`), call `apply_case_inputs(sheet, _padded(expected))`
  on a `RecordingSheet`, and assert: (a) the Sequence-flagged row's column I
  cell is `1.0` and no other row's is; (b) the spec rows are written (Role
  column non-blank for the declared width); (c) the prediction-input cells
  received `expected.results.prediction_interval.pred_input_values`. This is
  the test the BFN gap should have had: it would have gone red the moment the
  inspector dropped the override, and it covers the builder's path too (the
  helper is the builder's path now).
* **`test_apply_case_inputs_skips_the_override_when_no_period_declared`** —
  a case with `sequence_period is None` leaves column I blank on every row
  while still writing spec + prediction inputs. Pins the guard.
* **`test_apply_case_inputs_is_the_only_spec_writer_the_inspector_calls`** —
  source pin that `read_regression_df` calls `apply_case_inputs` and does
  *not* call `apply_spec_case` / `set_prediction_inputs` /
  `apply_sequence_period_overrides` directly (the inverse of the retired grep
  guard: the inspector must route through the helper, not re-inline the
  sequence). A light source-grep, kept because the failure mode (re-inlining)
  is silent in Excel.

The existing `tests/test_test_model_sheets.py::test_a_cases_typed_sequence_period_lands_in_spec_column_i`
tests the primitive `apply_sequence_period_overrides` directly. **Keep it** —
the primitive stays public (the guard path calls it), and the test pins the
row-position semantics the helper relies on. Add a one-line comment noting the
helper also exercises it via the composition test.

Delete `tests/test_inspect_regression_sheet_wires_sequence_period.py`.

## Part 5 — Why the guard path is not migrated

`_apply_guard_spec` (call site 3) is structurally different and stays as-is:

* It applies the spec through a **synthetic `_ExpectedView` / `_SpecView`**
  shim, not a real `RegressionSpecExpected` — the shim carries
  `resolved_prediction_group = "(all)"` and a padded `spec` but has **no
  `.results` and no `.design`**, so the helper (which reads
  `expected.results.prediction_interval...` and `expected.design...`) would
  `AttributeError` on it. Faking those would mean inventing a prediction
  interval and constructed-column-transforms for a case that declares neither
  — a guard case is not a fittable model.
* It types column I from `case.sequence_period_override` — a **pre-built
  dict**, a different field from the scalar `case.sequence_period` the helper
  builds its dict from.
* It writes **no prediction inputs** — there is no oracle prediction to prefill.

Forcing the guard path through the helper would either complicate the helper
with two override shapes and optional prediction inputs (re-introducing the
"two shapes that can drift" problem this plan exists to remove), or require a
synthetic expected that lies about the case. Neither is worth it: the guard
path is already correct, has its own test coverage
(`tests/test_regression_guard_states.py`), and is a different shape by design.
The helper is "apply a **fittable** case's inputs"; a guard state is not one.

## Execution order

```
Part 1   Add apply_case_inputs to regression_spec_sheet_io.py        (no Excel)
         uv run pytest tests/test_regression_spec_qc.py             (no behavior change yet)
---  checkpoint: commit  ---
Part 2   Migrate the builder (write_sheet_test_model)                (no Excel)
         uv run pytest tests/test_test_model_sheets.py
---  checkpoint: commit  ---
Part 3   Migrate the inspector (tools/inspect_regression_sheet)      (no Excel)
Part 4   Add tests/test_apply_case_inputs.py; retire the grep guard  (no Excel)
         uv run pytest            (full suite, must be green)
         uv run poe lint
         uv run poe verify-headless
---  checkpoint: commit  ---
Final    uv run poe verify-deep   (NEEDS Excel — owner-run)
         Confirm: BFN row for P01/P02 now PASSES (was nan on every prior run),
         and the life_full_profile precision-boundary count is unchanged
         (the helper does not touch k=19 residual stats).
```

## Risk assessment

**Low.** The helper is a pure de-duplication of three existing calls; it
introduces no new logic and no new invariants. The one subtlety — that padding
only appends `Omit` — is an existing invariant of `pad_spec_to_source_table`,
not a new one, and it is what makes the padded and unpadded paths
byte-equivalent through the helper.

**The thing to get right is the equivalence check in Part 2** (builder: dict
from `case.spec` → dict from `padded.case.spec`). Read it, confirm the
`Omit`-appends-only property makes the dicts identical, and the
`Sequence`-flagged row offsets stable. The behavioral test in Part 4 exercises
exactly this path with the padded P01 expected, so a regression here fails
headlessly.

**Excel gate.** The headless suite cannot exercise `read_regression_df` (it
opens Excel). The final `poe verify-deep` is the only gate that does, and it
is owner-run (not in CI — `windows-latest` lacks Office, per CLAUDE.md). The
expected outcome is the BFN row for P01/P02 flips from `nan`-mismatch to PASS,
and nothing else moves. If the life_full_profile precision-boundary count
changes, the helper touched something it should not have — revert and
re-examine.

**The surgical fix on this branch is the floor.** If the helper is deferred or
rejected, the surgical fix (inspector types column I inline) still repairs the
BFN gap. The helper is the permanent prevention; the fix is the immediate
remedy. They are not coupled — either can land without the other.

## Files touched (summary)

New files:
- `tests/test_apply_case_inputs.py` (behavioral test of the helper)

Deleted files:
- `tests/test_inspect_regression_sheet_wires_sequence_period.py` (source-grep guard, retired)

Modified files:
- `lambda_catalog/regression_spec_sheet_io.py` (+ `apply_case_inputs`)
- `lambda_catalog/write_sheet_test_model.py` (builder → `apply_case_inputs`; import trim)
- `tools/inspect_regression_sheet.py` (inspector → `apply_case_inputs`; import trim)

Untouched (deliberately):
- `lambda_catalog/write_sheet_test_model.py::_apply_guard_spec` (guard path — see Part 5)