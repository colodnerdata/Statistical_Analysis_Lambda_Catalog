# Open-PR error diagnosis — #225 and #226

**Date:** 2026-08-19
**Base:** `main` @ `46194fe`
**Scope:** the two open PRs, both off the same investigation

| # | Branch | What it is | CI | Mergeable |
|---|---|---|---|---|
| 225 | `investigate/verify-baseline-failures` | Plan doc + surgical BFN fix + source-grep guard | **never ran** | `unstable` |
| 226 | `refactor/apply-case-inputs-helper` | The `apply_case_inputs` helper (the plan, implemented) | 10/10 green | `clean` |

Five findings. One is a real correctness bug and is fixed in this branch; the
other four are reported here and need owner action on the PRs themselves.

---

## F1 — #226 leaks a typed Sequence Period between cases (confirmed)

**Copilot flagged this; it is correct, and it is the finding that matters.**
Thread [`r3809599073`](https://github.com/colodnerdata/Statistical_Analysis_Lambda_Catalog/pull/226#discussion_r3809599073), unresolved.

`apply_spec_case` clears exactly eight spec columns before rewriting a case:

```
B Role · C Include · D Type · E Reference · G Transform · H Sequence
M Interaction Term · N Interaction Operation
```

**Column I (Sequence Period) is not among them** — and I is an input column
(the spec block's input band is B–I; `write_spec_block` line 1129 states the
band as `(_C_ROLE, _C_SEQUENCE_PERIOD)`). `apply_sequence_period_overrides`
only writes rows named in its `overrides` dict, so a case that declares no
period never resets I.

### Why it was invisible before, and reachable now

The two callers that type column I today each write a case onto its **own
fresh sheet**:

* `write_sheet_test_model.write_test_model_sheet` — one sheet per case
* `write_sheet_test_model._apply_guard_spec` — one sheet per guard state

Nothing is reused, so nothing can leak. The verify **inspector** is the one
caller that applies *every* case in turn to a **single reused `Regression`
sheet**:

```python
for expected in regression_sheet_configs:      # tools/inspect_regression_sheet.py
    _apply_extra_columns(data_sheet, expected, all_extra_names)
    apply_case_inputs(sheet, expected)         # same `sheet` every iteration
```

Before #225/#226 the inspector never wrote column I at all, so the missing
clear was inert. **Both PRs make the inspector type column I, which is what
turns the latent gap into a live leak.** The bug the PRs fix and the bug they
introduce are the same missing invariant, one column apart.

### Why the stale value is not harmless

`Period In Use` (column J) prefers a typed period over the computed candidate:

```
=LET(nc,COLUMNS(Source_Data), sq,TAKE(Spec_Sequence,nc), sp,TAKE(Spec_Sequence_Period,nc),
     cand,IFERROR(Base_Period_Delta_Candidate(),""),
     MAP(SEQUENCE(nc),LAMBDA(i,
       IF(INDEX(sq,i)<>TRUE,"",
       IF(N(INDEX(sp,i))<>0,INDEX(sp,i),cand)))))
```

A row that is **not** Sequence-flagged returns `""`, so a stale I there is
inert. But a row that **is** flagged and has a non-zero stale I takes that
branch and **silently replaces the candidate**.

That is reachable. Only two cases declare a period — `production_lots_fixed_effects`
(P01) and `production_lots_log_transform` (P02), both Δ=1, registered first at
`analyze_regression_spec.py:1609` and `:1633`. Seven later cases flag a
Sequence axis (`Fiscal_Year`, `sequence=True` at `:1194 :1224 :1270 :1300 :1326
:1352 :1512`) and declare **no** period, expecting the computed candidate. P01/P02
run first and type 1.0; every later `production_lots` case whose Sequence row
shares that row offset inherits it.

The candidate for an annual `Fiscal_Year` axis is *also* 1.0, so today the leak
most likely produces **no visible verify failure**. That is the worst version
of this bug, not the harmless one: the inspector stops testing
`Base_Period_Delta_Candidate()` on those seven cases and nobody finds out until
a dataset with a non-unit spacing arrives.

### Fix (applied in this branch)

Root-cause, in `apply_spec_case` rather than in `apply_case_inputs`: the clear
list is hoisted to a named `_SPEC_INPUT_COLUMNS` constant with column I added.
This is deliberately **not** a fix on #226's branch:

* the defect is in `apply_spec_case`, which is on `main` and which **neither PR
  touches** — so the fix lands independently and closes the leak for whichever
  of #225/#226 merges, or both;
* `apply_spec_case` already enforces exactly this invariant for `Source_Table`
  and `$AK$12` ("a case must never be evaluated against whatever dataset the
  previous write left behind"). Column I was a missing entry in that list, not
  a new rule;
* both existing callers are already clear-then-write ordered
  (`apply_spec_case` → `apply_sequence_period_overrides`), so clearing I is
  safe at every call site, guard path included;
* nothing prefills column I on the spec rows, so clearing it cannot erase a
  default. Row 2's Spacing Verdict lives above `_FIRST_DATA_ROW` and is outside
  the clear range.

Pinned by `tests/test_spec_input_columns_are_cleared.py`, which asserts the
**partition** (block − label − computed displays − reserved F) rather than the
tuple's literal contents, so a newly appended input column that nobody adds to
the clear list fails headlessly instead of leaking on the next multi-minute
Excel run. Verified to fail 3/5 with column I removed.

---

## F2 — #226 test docstring contradicts its own code (minor)

Copilot thread [`r3809599106`](https://github.com/colodnerdata/Statistical_Analysis_Lambda_Catalog/pull/226#discussion_r3809599106), unresolved.
`tests/test_apply_case_inputs.py` `_p01_expected` says *"(not padded here — the
builder pads before calling the helper, so the test mirrors the builder and
pads too)"* and then returns `_padded(calculate_regression_spec_case(case))`.
The parenthetical contradicts itself and the code. Docs only. Same wording at
line 70. Owner fix on that branch.

---

## F3 — #225's CI never ran; it is gated, not broken

`mergeable_state: unstable` with **zero** check runs reads as a CI failure. It
is not. The head commit `18ae08b` was authored by `copilot-swe-agent[bot]`, and
both workflow runs for it completed with conclusion **`action_required`**:

```
18ae08be  pull_request  completed  action_required  02:31:49
18ae08be  push          completed  action_required  02:31:45
3573566e  pull_request  completed  success          02:17:13   <- previous, human-authored
3573566e  push          completed  success          02:15:20
```

`.github/workflows/ci.yml` triggers on `push: ["**"]` and `pull_request: [main]`,
so the trigger is fine — GitHub is holding the run behind the
approve-workflows-for-this-contributor gate. **Nothing to fix in the repo.**
The owner clicks *Approve and run* on the PR's Actions tab (or pushes one
human-authored commit) and the checks report.

---

## F4 — #225 silently weakens a guard assertion

`tests/test_sheet_writers.py`, the only non-plan, non-tool change in the PR
(`1 -`, unexplained in the commit message):

```diff
     for row in (4, 5, 6, 7, 8, 9, 10):
         formula = _formula(diag, row, _C_AE)
         assert "Fit_Design_Columns()" not in formula, row
         assert "Fit_Sample_Include()" not in formula, row
         assert "Design_Columns()" in formula, row
-        assert "Sample_Include()" in formula, row
```

This is the half of `test_regression_statistics_zone_reads_the_materialized_spills`
that pins *"...and the next zone has NOT moved yet"* — the assertion that keeps
the materialized-reads migration one zone per PR.

**The assertion passes on `main`** (verified: `1 passed`), and #225 touches
nothing that could affect the diagnostics zone — its only source change is to
`tools/inspect_regression_sheet.py`. The deletion is gratuitous and should be
reverted.

---

## F5 — #225 and #226 conflict, and #226 supersedes #225

`git merge-tree` confirms a real conflict in **`tools/inspect_regression_sheet.py`**
(both edit the same import block and the same loop body). Whichever merges
second needs manual resolution.

They are not independent changes. #225 is the plan; #226 is that plan carried
out, and #226 explicitly retires what #225 adds:

| #225 adds | #226 does |
|---|---|
| inline `apply_sequence_period_overrides` block in the inspector | replaces it with `apply_case_inputs` |
| `tests/test_inspect_regression_sheet_wires_sequence_period.py` (source-grep guard) | **retires it** — #225's own plan, Part 4, says to delete it |

Merging #225 after #226 would resurrect a guard test that #226's own plan
document declares retired, and re-inline the sequence the helper exists to own.

**Recommendation:** merge **#226** (green, `clean`, the actual implementation).
Reduce **#225** to the plan document only — drop its `tools/inspect_regression_sheet.py`
change, drop the retired grep-guard test, and restore the `Sample_Include()`
assertion from F4 — or close it, since #226's PR body already carries the
rationale and the plan doc is committed on its branch.

---

## Summary

| Finding | PR | Severity | Owner |
|---|---|---|---|
| F1 Sequence Period leaks between cases | 226 (root cause on `main`) | **correctness** | **fixed in this branch** |
| F2 test docstring contradicts code | 226 | docs | owner, on that branch |
| F3 CI gated at `action_required` | 225 | none — process | owner clicks *Approve and run* |
| F4 guard assertion deleted without cause | 225 | test coverage | revert on that branch |
| F5 conflicts with #226 and is superseded | 225 | process | merge 226; reduce/close 225 |
