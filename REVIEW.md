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

**What this file is not.** Not a decision record. Nothing here is resolved.
When a finding is resolved it moves to DECISIONS.md with its rationale and is
struck from this file.

**Method.** Read: `lambda_functions.json` (126 functions), `ARCHITECTURE.md`,
`DECISIONS.md`, `TODOs.md`, `ROADMAP.md`, `HUMAN_TEST_PLAN_v3_model_construction.md`,
`build_production.py`, `build_qc.py`, `write_sheet_regression.py`,
`write_sheet_univariate.py`, `write_sheet_model_construction.py`. Counts and
quotes below are from those sources, not from the README (which is stale — see
F7).

---

## Summary

Every finding below traces to a decision that was correct in isolation, argued
well, and recorded properly. The cost is in the sum, not in any single call.
That is what makes this hard to see from inside a PR: the review that would
catch it is not a review of any one change.

Three clusters:

- **F1, F2** — no canonical place to put "properties of this fit," so fit
  properties accumulate in argument lists and constructor names.
- **F3, F5, F6** — the sheet layout has no eviction mechanism, is implemented
  twice, and has no representation for the one feature that cannot be
  expressed as a column.
- **F4, F8** — the single-workbook delivery model has begun charging users of
  one sheet for the cost of another.

**F7** (doc drift) is separate and is the reason the other seven are hard to
see: no single document currently describes the shipped state.

---

## F1 — Optional-argument accretion is now doctrine

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

## F2 — One concept, three mechanisms

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

## F3 — The Regression sheet has no eviction mechanism

**Observation.** `write_sheet_regression.py` defines column anchors from `_C_A`
(1) through `_C_AW` (49) — seven content zones separated by gap columns. The
committed plan adds more: a unit-space block (v2.2), comparison anchors (v2.3),
weight display (v2.6).

Every one of those is correctly classified as *additive*, which is precisely
why nothing ever leaves. "Additive" is the property that makes a change a
MINOR; it is not evidence that the sheet can absorb it.

**Triage:** medium. Not a correctness problem. Becomes a usability problem at
some width, and no one has set that width.

---

## F4 — Univariate taxes every other user, and violates the stated philosophy

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

**Mitigating note.** `--skip-univariate` and `--skip-data-table-calculations`
already exist in the build. Bundling is closer than it looks; it is simply not
a first-class concept.

**Compounds with.** v2.4 Resampling adds a pre-drawn `Bootstrap_Random_Draws`
table and n-resample machinery to the same workbook.

**Triage:** high, and the only finding with a live correctness dimension.

---

## F5 — The spec block is implemented twice

**Observation.** `write_sheet_regression.py` (1,862 lines) and
`write_sheet_model_construction.py` (1,512 lines) each implement a spec block —
against the WHO and Mileage datasets respectively. A layout change touches
both writers.

This matters for v2.3, whose design assumes a single canonical spec/status
block shape that a second sheet can read through a fixed anchor
(`Comparison_Anchor`, DECISIONS.md § v2.3). There are already two
implementations of the shape that contract would be written against.

**Triage:** medium. Cheap now, expensive after v2.3 formalizes the anchor
contract as a public-interface commitment.

---

## F6 — Interactions have no representation and no reserved slot

**Observation.** `Interact(x1, x2)` exists as a standalone catalog function,
but the spec block is **one row per source column** and an interaction term is
not a column. It fits neither declared axis: Predictor Type is documented as
permanently closed (ARCHITECTURE.md § 3), and Role describes what a column
*is*.

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

---

## F8 — The versioning definition does not survive a multi-artifact split

**Observation.** ROADMAP.md defines the public interface as "the user's inputs
to the workbook" — singular. Any split into multiple emitted workbooks requires
per-artifact versioning, or a Univariate change bumps the number a Regression
user reads as the answer to "do my existing inputs still work?"

The in-workbook Version History sheet and its `Breaking?` column inherit the
same problem.

**Triage:** low until a split is actually scheduled; blocking at that point.

---

## Minor

- `PRESS` does not carry `[DF_Absorbed]`; `QQ_Correlation` does. PRESS as a sum
  of squared leave-one-out residuals genuinely does not need residual df, so
  this is probably correct — but the two appear in adjacent zones of the same
  sheet and the asymmetry is not legible from the signatures. Worth a comment
  in the JSON `notes` field either way.

---

## Compounding map

| Cluster | Findings | Shared root |
|---|---|---|
| Fit properties have no home | F1, F2 | No canonical carrier for "properties of this fit"; they accumulate in argument lists and constructor names |
| Sheet layout | F3, F5, F6 | One growing surface, implemented twice, with no slot for the one non-column feature |
| Delivery model | F4, F8 | Single workbook, single calc mode, single version number |
| Visibility | F7 | No document describes the shipped state |

---

## Sequencing implications

Recorded as observations, not as a plan.

1. **F1 gets more expensive per release, not per month.** The unwind cost is
   proportional to carrier count. v2.6 adds `[Weights]` to roughly the same 24
   functions. Any decision to change the mechanism is cheaper before v2.6 than
   after.

2. **F3, F5, and F6 all want the same breaking change.** Resolving them
   separately spends three layout breaks where one would do — the exact outcome
   the reserved-column policy was written to avoid.

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
