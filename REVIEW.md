# Architecture Review — 2026-08-01

A standing review of cross-cutting structural pressure in the library, written
against commit `f2123b2` ("Size and position Regression sheet notes to fit
their content", #145).

**What this file is.** A triage list of problems that span versions and
therefore have no natural home in the existing docs: [ROADMAP.md](ROADMAP.md)
holds the version plan, [DECISIONS.md](DECISIONS.md) holds *resolved* decisions
with rationale, [ARCHITECTURE.md](ARCHITECTURE.md) holds patterns a new feature
must honor, and [TODOs.md](TODOs.md) holds scoped work items. None of those is
the right place to record "this pattern is individually correct and
collectively expensive."

**What this file is not.** Not a decision record. When a finding is resolved it
moves to DECISIONS.md with its rationale and is struck from this file.

**Status as of the v3.0 documentation pass.** Seven of the eight findings — F1,
F2, F3, F4, F5, F6, F8 — plus the Minor item are **resolved**, recorded in
[DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline),
and struck below. They are kept in place rather than deleted so the reasoning
that produced the v3.0 design stays legible.

**F5 is a different kind of closure from the rest.** Six findings were resolved
by decisions this pass made. F5 was **already fixed in the code when the review
was written** — the spec block is imported by `write_sheet_regression.py`, not
duplicated in it — so the finding was never true as stated. It is struck as
corrected rather than as decided.

**Still open: F7 alone.** Documentation drift is a standing condition rather
than something a single pass closes — and F5 is now itself an instance of it.

**Method.** Read: `lambda_functions.json` (126 functions at the time of writing;
131 as of the v3.0 polish pass), `ARCHITECTURE.md`,
`DECISIONS.md`, `TODOs.md`, `ROADMAP.md`,
`build_production.py`, `build_qc.py`, `write_sheet_regression.py`,
`write_sheet_univariate.py`, `write_sheet_model_construction.py`,
`test_analyze_model_construction.py`, `test_difference_by_verification.py`
(the automated test modules that supersede the retired v2.0 human test plan).
Counts and quotes below are from those sources, not from the README (which is
stale — see F7).

---

## Summary

Every finding below traces to a decision that was correct in isolation, argued
well, and recorded properly. The cost is in the sum, not in any single call.
That is what makes this hard to see from inside a PR: the review that would
catch it is not a review of any one change.

Three clusters:

- **F1, F2** ✅ — no canonical place to put "properties of this fit," so fit
  properties accumulate in argument lists and constructor names.
- **F3** ✅, **F6** ✅, **F5** ✅*(never true)* — the sheet layout has no eviction
  mechanism, ~~is implemented twice,~~ and has no representation for the one
  feature that cannot be expressed as a column.
- **F4, F8** ✅ — the single-workbook delivery model has begun charging users of
  one sheet for the cost of another.

**F7** ⬜ (doc drift) is separate and is the reason the other seven are hard to
see: no single document currently describes the shipped state.

✅ resolved at v3.0 · ⬜ still open.

**What the resolutions have in common.** Each cluster was fixed by giving the
thing that was accumulating a *bounded home*, not by trimming instances: fit
properties get a four-element context block, the sheet gets a terminal zone with
an increasing-width ordering rule, and delivery gets a second artifact with its
own version. In each case the old rule constrained the count of changes without
constraining what could accumulate between them, which is why "additive"
authorized the growth indefinitely.

---

## F1 — Optional-argument accretion is now doctrine  — ~~OPEN~~ **RESOLVED at v3.0**

> **Resolved.** Properties of a fit now travel in the bounded `Model_Context` block; engine signatures collapse from five arguments to four. ARCHITECTURE § 7's reserved-slot pattern no longer applies to argument lists. Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).


**Observation.** `[DF_Absorbed]` is carried by **24 functions** in
`lambda_functions.json`:

```
Residual_Degrees_Of_Freedom, Adjusted_R_Squared, MS_Residual, SE_Regression,
AIC, BIC, AICc, QQ_Correlation, SE_Coefficients, T_Statistics, P_Values,
Confidence_Interval_Lower, Confidence_Interval_Upper, Partial_R_Squared,
Partial_Correlation, Prediction_Interval, Group_Prediction_Interval,
Scaled_Residuals, Studentized_Residuals, Studentized_Residuals_Ranked,
Cooks_Distance, Scaled_Residuals_Ranked, F_Statistic, F_Statistic_P_Value
```

The v2.1 df-plumbing decision (DECISIONS.md § v2.1) was sound on its own terms:
default 0, no-FE models compute identically, MINOR instead of MAJOR. The
problem is that it has since been generalized into a rule. DECISIONS.md § v2.6
specifies `[Weights]` as "a single optional `[Weights]` argument (default
uniform)" and states that the `[DF_Absorbed]` precedent "is the exact pattern
to follow." ARCHITECTURE.md § 7 promotes the same idea to a named
reserved-slot pattern applying to "a LAMBDA's argument list or its internal
`SWITCH`."

**Trajectory.** `(X_s, Y, [Allow_Intercept], [Include], [Alpha],
[DF_Absorbed], [Weights], …)`, with a `Cluster` key wanted at v2.7+. Excel has
no keyword arguments, so every addition converts to positional comma-counting
for any user writing a call by hand.

**Present worst case.** The Regression sheet's prediction cell is a
**695-character** two-branch construction ending in a nine-argument call with a
hard-coded cell address mid-list:

```
Group_Prediction_Interval(X_s(),Response_Column(),pred_input,
  Prediction_Group_Column(),$AH$12,Allow_Intercept,
  Sample_Include(),alpha,Absorbed_Degrees_Of_Freedom())
```

Ordinary statistics remain legible
(`=R_Squared(X_s_Within(),y_s(),Allow_Intercept,Sample_Include())`), but every
diagnostic now carries a fixed repeated tail of
`,Allow_Intercept,Sample_Include(),Absorbed_Degrees_Of_Freedom()`.

**Why it does not self-correct.** Each addition is individually non-breaking,
which is the criterion the versioning policy uses to authorize it. There is no
step at which the rule says stop, and the rule is written down as a virtue.

**Triage:** high. Blocks nothing today; makes v2.6 measurably worse; the cost
of unwinding grows with each carrier added.

---

## F2 — One concept, three mechanisms  — ~~OPEN~~ **RESOLVED at v3.0**

> **Resolved.** One constructor pipeline (`Design_Columns()` / `Design_Response()`) plus one named escape hatch (`Predictor_Columns()`) replaces the constructor name fork, applying declared stages in a fixed order. Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).


**Observation.** Fixed effects is implemented as a **constructor variant**
(`X_s_Within()`, `y_s()`) *and* as a **signature argument** (`[DF_Absorbed]`).
WLS is planned as signature-only (DECISIONS.md § v2.6). Two-way FE will require
a third constructor (`Absorb_Two_Way_Fixed_Effects`). All three are the same
kind of operation — a transformation of the design matrix and/or response —
arriving by different routes.

**Visible cost today.** The sheet must know which constructor each call site
wants. Fit statistics take `X_s_Within()`; `GVIF`, `Generalized_Tolerance`,
`Pearson_R`, `Spearman_R`, `Skewness`, and `Kurtosis` take `X_s()`. That
distinction is statistically correct — collinearity and marginal diagnostics
want pre-demeaning columns — and entirely invisible in the names. Nothing
enforces it.

**Trajectory.** Weighting and two-way absorption produce a cross product of
constructor variants, each with its own correct-call-site rule and none of them
checkable by the build.

**Triage:** high. Shares a root with F1; a resolution that addresses one should
be evaluated against both.

---

## F3 — The Regression sheet has no eviction mechanism  — ~~OPEN~~ **RESOLVED at v3.0**

> **Resolved.** ARCHITECTURE § 4b sets the boundary this finding said was missing: materialized zones run in increasing width and terminate in the unbounded Constructed Design Matrix, and nothing may be placed to its right. Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).


**Observation.** `write_sheet_regression.py` defines column anchors from `_C_A`
(1) through `_C_AW` (49) — ~~seven~~ **five** content zones separated by gap
columns. The committed plan adds more: a unit-space block (v2.2), comparison
anchors (v2.3), weight display (v2.6).

*Correction (v3.0 doc pass).* This finding originally said seven zones. `_ZONES`
defines five — `(A,N) (P,V) (X,AE) (AG,AI) (AK,AV)` — separated by four
ungrouped gap columns (O, W, AF, AJ), with AW a non-content gutter. The finding
stands: the count was wrong, the growth trajectory it describes was not.

Every one of those is correctly classified as *additive*, which is precisely
why nothing ever leaves. "Additive" is the property that makes a change a
MINOR; it is not evidence that the sheet can absorb it.

**Triage:** medium. Not a correctness problem. Becomes a usability problem at
some width, and no one has set that width.

---

## F4 — Univariate taxes every other user, and violates the stated philosophy  — ~~OPEN~~ **RESOLVED at v3.0**

> **Resolved.** The build emits two workbooks, each setting its own calculation mode. Both carry the complete function library. The Regression workbook returns to full Automatic and Univariate's fits are live. Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).


**Observation.** `build_production.py` leaves the shipped workbook in
`XL_CALCULATION_SEMIAUTOMATIC` — Automatic except Data Tables. This is forced
by the Univariate sheet's six two-input Data Tables (Weibull, Gamma, Beta ×
two stages, 20×20 each), **2,400 NLL evaluations per full recalculation**.

Two consequences, both shipped:

1. Every Regression user receives a non-default calculation mode as a
   side effect of a sheet they may never open.
2. **Univariate fit results are stale until the user presses Ctrl+Alt+F9.**

Consequence (2) is the serious one. The library's stated design philosophy is
live recalculation, formula transparency, and visible failure. The flagship
distribution-fitting sheet currently displays a previous answer with no
indication that it has done so — silent wrongness, in the one place the
philosophy exists to prevent it.

**Mitigating note (pre-resolution).** `--skip-univariate` and
`--skip-data-table-calculations` already existed in the build, so bundling was
closer than it looked — it was simply not a first-class concept. It is now:
`build_production.py` and `build_univariate.py` are the two build targets
(`--skip-univariate` retired with the split).

**Compounds with.** v2.4 Resampling adds a pre-drawn `Bootstrap_Random_Draws`
table and n-resample machinery to the same workbook.

**Triage:** high, and the only finding with a live correctness dimension.

---

## F5 — The spec block is implemented twice  — ~~OPEN~~ **RESOLVED — already fixed in the code**

> **Resolved, and it was resolved before this review was written.** The spec
> block is implemented **once**. `write_sheet_regression.py` imports the
> spec-block writers — `_write_spec_block`, `_write_spec_feedback`,
> `_write_intercept_control`, `_set_sheet_scoped_names`,
> `_set_spec_block_column_widths` — plus every `_C_*` column constant from
> `write_sheet_model_construction.py`, and calls them. Its module docstring
> states the intent: *"the spec-block writers are imported from
> write_sheet_model_construction so the two sheets can never drift."*
> Separately, the Model Construction **sheet** is deleted by both builds
> (`build_production.py`, `build_qc.py`), so only one spec block ships at all.
> Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#the-spec-block-is-implemented-once-not-twice).

**Observation.** ~~`write_sheet_regression.py` (1,862 lines) and
`write_sheet_model_construction.py` (1,512 lines) each implement a spec block —
against the WHO and Mileage datasets respectively. A layout change touches
both writers.~~

*Correction (v3.0 doc pass).* The line counts are right and the inference from
them is wrong. Both files are large and both concern the spec block, but the
relationship is **import**, not duplication: one module owns the writers, the
other calls them. A layout change touches one writer. What the review measured
was file size; what it concluded was coupling.

This matters for v2.3, whose design assumes a single canonical spec/status
block shape that a second sheet can read through a fixed anchor
(`Comparison_Anchor`, DECISIONS.md § v2.3). ~~There are already two
implementations of the shape that contract would be written against.~~ There is
one, which is what that contract needs.

**Triage:** ~~medium~~ none — no work outstanding.

**What is left is a naming problem, not a structural one.**
`write_sheet_model_construction.py` no longer writes a shipped sheet; it is the
spec-block component library the Regression sheet is built from, and its name
still refers to a sheet that both builds delete. Renaming it (and dropping the
dead standalone-CLI path) is tracked in
[TODOs.md](TODOs.md). That is cosmetic — it changes no behavior and closes no
finding.

---

## F6 — Interactions have no representation and no reserved slot  — ~~OPEN~~ **RESOLVED at v3.0**

> **Resolved.** Interactions are declared with two spec columns (M Interaction Term, N Interaction Operation) and a closed operation vocabulary carrying a symmetry attribute. A second spec section below the per-column block was considered and rejected. Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).


**Observation.** ~~`Interact(x1, x2)` exists as a standalone catalog function,
but~~ `Interact(x1, x2)` is **specified but not built**, and in either case
the spec block is **one row per source column** and an interaction term is
not a column. It fits neither declared axis: Predictor Type is documented as
permanently closed (ARCHITECTURE.md § 3), and Role describes what a column
*is*.

*Correction (v3.0 doc pass).* This finding originally asserted that `Interact`
ships. It does not — `Interact`, `Model_Matrix`, and `Dummy_Column` are all
specified in ARCHITECTURE.md § 5 and listed as v2.2 work items in TODOs.md, but
none is in `lambda_functions.json`. The correction strengthens the finding
rather than weakening it: there is no standalone escape hatch either, so
interactions are currently unreachable by any route.

The reserved-column policy exists exactly to prevent a second layout break —
column F (Order) and column G (Transform) were reserved for that reason, and
G's activation at v2.2 is documented as the pattern succeeding. **No slot is
reserved for interactions.**

**Why this one is different.** Every other pending feature can be absorbed
additively. An interaction mechanism cannot be retrofitted the way column G
was, because it is not a per-column property — it needs either a new spec
column with non-column semantics or a second spec section. Both are layout
changes.

**Relevance.** Two-factor ANOVA is on the v2.7+ candidate list; factor ×
continuous is the natural request immediately after one-way FE ships.

**Triage:** medium urgency, high irreversibility. The decision is worth making
before the next layout touch even though the feature itself is far down the
roadmap.

---

## F7 — Documentation drift is measurable

**Observation.**

| Claim | Source | Actual |
|---|---|---|
| v2.1 is "Planned" | ROADMAP.md ladder | Built; `y_s`, `X_s_Within`, `Absorbed_Degrees_Of_Freedom`, `Group_Prediction_Interval` all in the JSON; human test plan written |
| Spec block is A–I | ROADMAP.md § v2.0 | A–L (ARCHITECTURE.md § 4) |
| Spec-driven regression is v2.0 | ROADMAP.md (renumbered from 3.0.0) | `write_sheet_model_construction.py` docstring and the human test plan both say v3.0 |
| `F_Stat`, `P_Value_F`, `R_squared`, `Hat_diagonal`, `Grid_Argmin` | README.md function reference | Renamed by the v2.0 pass; JSON has `F_Statistic`, `F_Statistic_P_Value`, `R_Squared`, `Hat_Diagonal`, `Grid_Argument_Minimum` |
| `GVIF` / `Generalized_Tolerance` | absent from ROADMAP | shipped |

**Why it matters beyond tidiness.** `lambda_functions.json` is the source of
truth for *functions*. Nothing is the source of truth for the *shipped state of
the plan*. Four cross-referencing planning documents plus CLAUDE.md/AGENTS.md
plus two test plans are hand-maintained, and a reader starting from the README
or ROADMAP forms a materially wrong picture of what exists.

**Triage:** medium, but it is the precondition for the other findings being
visible at all.

### Status after the v3.0 documentation pass — **STILL OPEN**

The listed instances are reconciled, but the finding stands: nothing prevents
the next one. Row by row:

| Row | Now |
|---|---|
| v2.1 "Planned" | Fixed — the ladder now reads shipped inside the 3.0.0 artifact, with the automated gate that was actually met named in place of the retired hand-run plan |
| Spec block A–I | Was already A–L in ROADMAP before this pass; **A–O** since v3.0 stage 3 — M and N (the interaction pair) *and* O (the Design Columns audit) |
| Spec-driven regression is v2.0 vs v3.0 | Partly fixed — `write_sheet_model_construction.py`'s docstring is corrected and the test plan that carried the old label in its filename is retired. The label survives in comments in three test modules, `build_production.py`, and `analyze_regression_spec_block.py`; tracked in TODOs.md |
| `F_Stat`, `P_Value_F`, … | Already corrected upstream — the current README carries no function reference table at all |
| `GVIF` / `Generalized_Tolerance` | Fixed — named in the v2.1 ladder row |

Four further drift instances this pass found and fixed, none of which were in
the original table: the Role dropdown values in ARCHITECTURE.md § 3 omitted the
parenthetical suffixes that formulas actually string-compare against; `Interact`,
`Model_Matrix`, and `Dummy_Column` were documented as though they ship;
this file's own zone count was wrong (see F3); and **F5 itself described a
duplication that the code had already replaced with an import** (see F5).

That last one is the sharpest illustration of why this finding matters. F5 was
not a doc lagging behind a *plan* — it was a review of the code, written from
the code, that reached a wrong conclusion because it compared two file sizes
instead of reading one import statement. Drift is not only docs falling behind
code; it is also the absence of any check that a claim *about* the code still
holds.

One instance was found and **not** fixed at the time, because the file is
outside the documentation set: CLAUDE.md / AGENTS.md described the Regression
sheet's zones with stale column letters. **That one has since closed** — the
v3.0 stage 3 layout break rewrote both files, and they now read A–Q / S–Y /
AA–AH / AJ–AL / AN–AY with gap columns R, Z, AI, AM, matching `_ZONES` in
`write_sheet_regression.py`.

### Status after the 2026-08-03 review — **STILL OPEN**

A second reconciliation pass, run against the post-v3.0 tree, found six more
instances. All are now fixed; none were caught by anything other than reading
the source:

| Instance | Was | Is |
|---|---|---|
| Catalog size | "126 LAMBDA definitions", in README ×2, ROADMAP ×2, CONTRIBUTING ×2, DECISIONS, and this file | **131** — v3.0 stage 2 added `Model_Context` plus the four `Context_*` accessors |
| `X_s` | Named as a live sheet-scoped closure in README and CLAUDE/AGENTS | Renamed `Predictor_Columns` at v3.0 stage 1; `X_s` is not in the catalog |
| §4b gap columns | ARCHITECTURE cited `_C_O` / `_C_W` / `_C_AF` / `_C_AJ` | `_C_R` / `_C_Z` / `_C_AI` / `_C_AM` — the pre-stage-3 letters, in prose |
| Univariate Data Tables | CONTRIBUTING's Univariate-build section said "Weibull/Gamma/Beta Data-Table grid-search" | Weibull and Gamma were demoted to formula grids at v3.0; Beta alone uses Data Tables |
| Headless verifier | CLAUDE/AGENTS/CONTRIBUTING said Layer 1 is discovered "once it lands" | It landed; `tests/test_workbook_invariants.py` runs in CI on every push |
| TODOs stage 2 | Heading read "VERIFICATION GATE OUTSTANDING" while a bullet in the same section recorded the gate as passed | Section contradicted itself; heading corrected to match stage 1 |

**This file was itself an instance.** The row above reading "now A–N with the
v3.0 interaction columns" was wrong (it is A–O), and the "not fixed" paragraph
above claimed a live CLAUDE/AGENTS defect that had already been repaired. F7 is
the finding about documentation drift, and the record *of* F7 had drifted —
which is the same lesson F5 taught, arriving a second time.

**Why the finding stays open.** Every fix above was made by hand, by reading the
source. That is exactly the mechanism the finding says does not scale. Two
mechanical checks are proposed in
[CONTRIBUTING.md § Documentation drift](CONTRIBUTING.md#documentation-drift-proposed-check--not-yet-implemented)
— function names resolving against the JSON, and cross-document anchors
resolving against real headings. Neither is built. Until one is, this finding is
a standing condition.

The 2026-08-03 pass ran the anchor check by hand over every `.md` in the repo
and found **zero** unresolved targets — and then deleted three files, which is
precisely the change that breaks anchors silently. The check found nothing
because it was run; it is not running on its own, which is the point.

---

## F8 — The versioning definition does not survive a multi-artifact split  — ~~OPEN~~ **RESOLVED at v3.0**

> **Resolved.** One library version covers the shared catalog; a per-workbook version covers each artifact's input surface. The `Breaking?` flag attaches to the workbook version. Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).


**Observation.** ROADMAP.md defines the public interface as "the user's inputs
to the workbook" — singular. Any split into multiple emitted workbooks requires
per-artifact versioning, or a Univariate change bumps the number a Regression
user reads as the answer to "do my existing inputs still work?"

The in-workbook Version History sheet and its `Breaking?` column inherit the
same problem.

**Triage:** low until a split is actually scheduled; blocking at that point.

---

## Minor — ~~OPEN~~ **RESOLVED at v3.0**

> **Resolved.** Correct, and the reason is mechanical rather than a judgment
> call. `PRESS` is `SUMSQ(LOOCV_Residual(…))`, and neither `eᵢ/(1−hᵢ)` nor the
> leverage depends on a df count, so there is no term for absorbed df to enter.
> `QQ_Correlation` calls `Scaled_Residuals_Ranked`, which divides by a σ estimate
> computed on residual df. The generalized rule — **a statistic needs
> `[DF_Absorbed]` exactly when it divides by a residual-df-based variance
> estimate** — goes in each function's JSON `notes` field, so the asymmetry is
> legible from the catalog sheet without reading both formulas. Full rationale in
> [DECISIONS.md § v3.0](DECISIONS.md#press-correctly-omits-df_absorbed).

- `PRESS` does not carry `[DF_Absorbed]`; `QQ_Correlation` does. PRESS as a sum
  of squared leave-one-out residuals genuinely does not need residual df, so
  this is probably correct — but the two appear in adjacent zones of the same
  sheet and the asymmetry is not legible from the signatures. Worth a comment
  in the JSON `notes` field either way.

---

## Compounding map

| Cluster | Findings | Shared root | Status |
|---|---|---|---|
| Fit properties have no home | F1, F2 | No canonical carrier for "properties of this fit"; they accumulate in argument lists and constructor names | ✅ `Model_Context` + the constructor pipeline |
| Sheet layout | F3, F5, F6 | One growing surface, ~~implemented twice~~, with no slot for the one non-column feature | ✅ F3 and F6 resolved at v3.0; **F5 was never true** — the spec block is imported, not duplicated |
| Delivery model | F4, F8 | Single workbook, single calc mode, single version number | ✅ two artifacts, two calc modes, two version numbers |
| Visibility | F7 | No document describes the shipped state | ⬜ instances reconciled by hand; no mechanism yet |

**Correction to an earlier draft of this pass.** A previous revision claimed v3.0
made F5 *worse*, on the reasoning that the two new spec columns, the audit
column, and the materialization zone would each have to be built in both
writers. That was wrong, and wrong in the same way F5 itself was: it inferred
duplication from two large files without checking the import. Each of those
lands in **one** writer and the other inherits it. F5 is not a v3.0 cost.

**What the cluster's shared root actually was.** Two of the three findings held —
one growing surface with no eviction rule, and no slot for interactions. The
"implemented twice" leg did not. Worth noting because it is the second time in
this file that a correct-sounding structural inference was drawn from
file-level evidence without reading the coupling; see also F6's `Interact`
premise.

---

## Sequencing implications

Recorded as observations, not as a plan. Written before the v3.0 decisions; the
outcome of each is noted.

1. **F1 gets more expensive per release, not per month.** The unwind cost is
   proportional to carrier count. v2.6 adds `[Weights]` to roughly the same 24
   functions. Any decision to change the mechanism is cheaper before v2.6 than
   after.

2. **F3, F5, and F6 all want the same breaking change.** Resolving them
   separately spends three layout breaks where one would do — the exact outcome
   the reserved-column policy was written to avoid. *(F5 turned out not to be a
   layout problem at all; the observation holds for F3 and F6.)*

3. **F6 is the only finding that is irreversible if deferred.** Everything else
   can be fixed later at higher cost; an interaction mechanism cannot be added
   additively to a one-row-per-column spec. The representation decision is
   worth making before the next layout touch, independent of when interactions
   are actually implemented.

4. **F4 is the only finding with a live correctness dimension.** Stale
   Univariate results are shipping now.

5. **F7 should be cheap and should probably go first.** It is the reason the
   rest are hard to see, and reconciling the docs against the JSON and the
   sheet writers is mechanical work that could be partially build-enforced.

### How each played out

1. **Still live, and now a scheduling constraint.** The v3.0 decision changes the
   WLS mechanism from a threaded `[Weights]` argument to √w scaling in the
   constructor. If v2.6 ships first it needs the argument anyway, and v3.0 then
   unwinds it across the same ~24 functions. Recorded in the ROADMAP v2.6 entry.
2. **Accepted for F3 and F6, and it shaped the recommended v3.0 scope.** The
   recommendation ships the interaction columns and the audit column as *layout*
   inside v3.0 — reserved and unwired — so the insertions cost one break rather
   than three. F5 dropped out of the cluster once the import was checked.
3. **Accepted.** F6's representation decision is resolved at v3.0 even though the
   feature itself ships later, which is exactly what this observation argued for.
4. **Confirmed and acted on first in the write-up.** F4 is the finding with a
   live correctness dimension, and the split is packaging-only and non-breaking,
   so it carries the least risk of any v3.0 change.
5. **Half-right.** Reconciling the docs was cheap and did go first — this pass.
   But "should probably go first" understated it: the reconciliation is what
   surfaced three drift instances the original table missed, including two
   factual errors in this file. The mechanism it calls for still does not exist.
