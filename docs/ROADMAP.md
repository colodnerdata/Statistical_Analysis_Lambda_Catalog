# Lambda Library — Roadmap

A native-Excel statistical and regression library (LAMBDA-based, no VBA) intended to
replace and expand upon the Excel Analysis ToolPak. Every release ships **functions plus
a pre-built worksheet** that drives them — the worksheet is a first-class deliverable, not
an afterthought.

Design philosophy: live recalculation, formula transparency, auditability, and no Solver /
no VBA wherever it can be avoided. The goal is that any result can be interrogated by
clicking the cell.

**Documentation map.** This file is the version plan only — the ladder, what's shipped,
and what's next. The "why" behind any design choice lives in
[DECISIONS.md](DECISIONS.md). The between-versions rules (naming convention, function
categories, the Role / Type / Sequence taxonomy, the spec block, the data-transformation
taxonomy) live in [ARCHITECTURE.md](ARCHITECTURE.md). Active work lives in
[TODOs.md](TODOs.md). The test-model suite each milestone has to grow — and the
ordering principle this ladder follows from v3.4 on — is
[docs/MODEL_TESTING_ASSETS.md](MODEL_TESTING_ASSETS.md). A reader who needs the
rationale behind a milestone's design should follow the cross-link from the
milestone's bullet to the corresponding section in DECISIONS.md.

---

## Versioning & Release Conventions

Semantic versioning, `MAJOR.MINOR.PATCH`:

- **MAJOR** — a breaking change to the library's **public interface**, defined below.
- **MINOR** — new functions, new sheets, or new sheet capabilities that do **not**
  break the public interface. A large additive feature is still a MINOR.
- **PATCH** — bug fixes, formula corrections, documentation edits.

**What the "public interface" (API) is for this library.** Unlike a code library
whose API is a set of function signatures, this library's API is **the user's inputs
to the workbook**: the layout and semantics of the input cells and control blocks a
user fills in, the sheet-scoped and workbook-scoped named ranges a user may reference
in their own formulas, and the meaning of an existing specification already saved in a
user's file. A release is **breaking (MAJOR)** when a workbook a user built against the
prior version would, on adopting the new version, either stop working or silently
compute something different — e.g. a control block moving or changing meaning, a named
range disappearing or being repurposed, or `x_s()` changing what it returns for the
same inputs. A release is **non-breaking (MINOR)** when every specification valid under
the prior version produces the same result, and the new capability is reached only by
new inputs the user opts into. Canonical function *names* are part of the API only to
the extent a user types them into cells; internal implementation is not.

This is why, for example, Univariate is a MINOR (a wholly new sheet; no existing input
changes meaning) while Specification-Driven Regression is a MAJOR (it changes what
`x_s()` returns and restructures the Regression sheet's control block — existing
formulas and saved specs change behavior).

Each release maintains a changelog. For distribution to non-git users (e.g. cost
estimators), a "Version History" sheet inside the workbook mirrors the changelog so the
history travels with the file. That sheet carries an explicit **`Breaking?` (yes/no)**
column so a workbook user gets the one signal the version number is *for* — "do my
existing inputs still work?" — without the number also having to convey "how big is
this release." Deliverable size is described in the changelog prose; breakage is the
flag.

### One workbook, one number

The build emits one workbook (`Lambda_Library.xlsx`), so there is one workbook
version alongside the one library version. Both move under the same conditions
stated above; the `Breaking?` flag attaches to the workbook version, since it
answers a question about a user's saved inputs.

**History.** From v3.0 through the Beta `Full_Factorial` rework, the build
emitted two artifacts — a Regression workbook and a Univariate workbook — and
the versioning scheme carried a separate per-workbook version. That split was
driven by a Data Tables calculation-mode conflict: the Univariate sheet's Beta
fit used two two-input Data Tables, forcing `XL_CALCULATION_SEMIAUTOMATIC`, which
stale-ified the Regression sheet's live results. Once Beta moved to a
`Full_Factorial` dynamic-array spill the Data Tables were gone, the
calculation-mode conflict vanished, and the two workbooks were reunified into
one. The split decision is marked SUPERSEDED in
[DECISIONS.md § Univariate becomes its own workbook](DECISIONS.md#univariate-becomes-its-own-workbook--superseded);
the reunification rationale is in
[DECISIONS.md § the Data-Table driver for the workbook split retires](DECISIONS.md#v3x--betas-grid-search-becomes-a-full_factorial-spill-the-data-table-driver-for-the-workbook-split-retires).

**Version ladder (current plan):**

| Version | Milestone | Breaking? | Status |
|---|---|---|---|
| v1.0 | Multivariate OLS / MLR | — (baseline) | Shipped |
| v1.1 | Univariate (descriptives, histograms, distribution fitting) | No | **Shipped 2026-06-29** (workbook 1.1.0; renumbered from 2.0.0). MoM-vs-MLE resolved: MLE throughout. New sheet, no existing input changes meaning. PDF functions dropped as unnecessary — the histogram tables already compute per-bin probabilities as CDF deltas between bin boundaries. The two post-release leftovers (per-distribution Q-Q plots and combo-chart overlay lines built on those CDF-delta columns) shipped with the next workbook build |
| v1.2 | Workbook hardening & regression usability (Name Manager notes, identity-line data series, intercept-only and undersized-sample guards, LOOCV_Residual, build retry/RPC handling) | No | **Shipped 2026-07-03** (workbook 1.2.0; renumbered from 2.1.0) |
| v2.0 | Specification-Driven Regression (roles: Continuous / Categorical) | **Yes** | **Shipped 2026-07-05** (workbook 2.0.0; renumbered from 3.0.0) — MAJOR. Changed `x_s()` return semantics and restructured the Regression control block; includes the canonical rename pass. Shipped with `Transform` as a reserved placeholder column as planned; users transform their own variables via extra input-table columns in the interim |
| v2.1 | Sequence axis + gap-aware longitudinal + serial-correlation diagnostics + Fixed Effects (Role axis, one-way only) + Generalized VIF | No | **Shipped inside the 3.0.0 artifact** — every TODOs #1–#10 item is DONE and verified against a live build (0 mismatches across all 12 spec-driven QC cases), with the FE engine independently pinned against `statsmodels` LSDV by `test_within_estimator`, `test_df_absorbed_threading`, and `test_group_prediction_interval`. `Design_Response` and `Design_Columns` (shipped at v2.1 as `y_s` / `X_s_Within`; renamed by the v3.0 constructor pipeline), `Absorbed_Degrees_Of_Freedom`, `Group_Prediction_Interval`, `GVIF`, and `Generalized_Tolerance` are all in `lambda_functions.json`. Never got its own release build — the features reached users inside 3.0.0, and the 2.1.0 Version History entry was backfilled later, at v3.1, rather than written at release. DEFERRED follow-on polish remains |
| v2.2 | Transforms (Response / Predictor Log, unit-space comparability) + the standalone Data Transformation function library | No | Partially delivered — MINOR, and likewise shipped inside the 3.0.0 artifact; its 2.2.0 Version History entry was backfilled later, at v3.1, rather than written at release. Column-G `Log` wiring shipped (`Response_Column()`/`X_s()` — renamed `Predictor_Columns()` at v3.0 — plus `Constructed_Column_Names()`/`Constructed_Column_Transforms()`, the Prediction Inputs auto-log step, `Ln_Positive`); the unit-space dispatcher, Duan back-transformation, and the rest of the standalone transform library (Center, Zscore, Winsorize, …) remain open and ship as **v3.3**, after v3.0 |
| **v3.0** | **The engine-interface release** — bounded `Model_Context`, intercept relocation, the constructor pipeline, the Univariate split (later reunified — see below), and the layout break | **Yes** | **Shipped 2026-08-02** (workbook 3.0.0; Univariate artifact 1.0.0, later reunified). Three stages plus the split, landed as separate reviewable pull requests: stage 1 (constructor pipeline + intercept relocation), stage 2 (the `Model_Context` / `[Context]` collapse), the Univariate split, and stage 3 (the layout break). Stages 1-2 and the split were non-breaking — they restructure the engine and the packaging, not the user-typed spec block, so a Regression spec saved under 2.0.0 produces identical output (stage one QC: zero mismatches across all twelve cases; stage two gate green). Stage 3 is where the `Breaking?` flag turns **Yes**, and it breaks ADDRESSES, not meanings: three columns are APPENDED to the spec block (M/N interaction pair, O Design Columns audit), so A–L keep their letters and their meanings and no fitted number moves, but every zone right of the spec block shifts three columns. See the milestone entry below |
| v3.1 | Interaction wiring — the constructor actually builds the interaction columns v3.0 stage 3 inserted | No | **Shipped 2026-08-03** (workbook 3.1.0) — MINOR, and exactly the follow-on the reserved columns were for: three LAMBDA definitions and one audit formula changed, and no column moved. `Predictor_Columns()` and its two twins read M/N and emit the pairwise combination (1 column for Continuous × Continuous, L−1 for Continuous × Categorical, (L₁−1)(L₂−1) for Categorical × Categorical); the Design Columns audit gained its `k(row)×k(operand)` term in the same edit, off the same width helper. A spec with M and N blank computes identically to 3.0.0 |
| v3.2 | Full materialization of the design matrix | No | Delivered — MINOR. Stage 3 established the terminal zone and its width guard, and the spills that fill it — `Design_Columns()` into the design-matrix zone, `Sample_Include()` into its own — landed in the code, replacing both `"reserved"` placeholders. The ~30 engine call sites are now repointed at those spills via the `Fit_Design_Columns()` / `Fit_Sample_Include()` readers (sheet zones and the eight catalog LAMBDA bodies that called the constructors), the performance win is banked (recalculate 21.8s → 10.5s, cell-by-cell spec verifier green), and the artifact rebuild carries it to users. Still open: the cosmetic `Sample_Include` *name*-over-a-spill promotion (self-referential at the producing cell, now optional since every call site already reads the reader) |
| v3.3 | Transforms remainder — unit-space dispatcher, Duan back-transformation, the model formula label | No | **SHIPPED** — MINOR. *Planned as the second half of v2.2*, moved after v3.0 with the rest of the feature train; the column-G `Log` wiring already shipped at v2.2. The **standalone transform library** was planned inside this milestone and now ships as **v3.9** — it is the ladder's most expensive item to test, and nothing else waits on it |
| v3.4 | Model Comparison Sheet | No | Planned — MINOR, a *nice-to-have*. *Planned as v2.3.* Read-only across finished Regression sheets; ships after the Transforms remainder (v3.3) so its comparisons are unit-space-honest from day one. **Test scale: additive (~1×)** — it reads models the suite already has |
| v3.5 | `Cluster` Role (clustered-robust SEs) | No | Planned — MINOR. *Planned as v2.7+; promoted out of the unordered bucket by the [ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth).* Forward-wired from `Serial_Correlation_Group()`'s dormant branch. **Test scale: near-additive** — a variance-estimator variant over a few existing models |
| v3.6 | `Time` Role + lag/difference semantics | No | Planned — MINOR. *Planned as v2.7+; promoted out of the unordered bucket.* Partially forward-wired via the v2.1 Sequence axis. **The sheet half moved to v3.12** — this milestone is the engine work (the Role, cross-sheet `Lag_By` / `Difference_By` time semantics, the calendar-dated dataset), which is Regression-track and belongs here; the worksheet that consumes it is a new analysis surface and ships with the trailing block. **Test scale: near-additive — and it closes a coverage gap that exists today**: its calendar-dated dataset is what finally makes the Sequence calendar-signature verdict testable |
| v3.7 | `Weight` Role (WLS) | No | Planned — MINOR. *Planned as v2.6; claimed as v3.7 all along, though it reaches the slot by a different route.* User-supplied weights as the first stage; variance-driver-derived weights and FGLS as later follow-ons. The `Weight` Role, its cardinality rule, and the three-stage scope stand; the **implementation mechanism changed at v3.0** — √w scaling in the constructor, not a threaded `[Weights]` argument. Shipping after v3.0 is what makes that the first implementation rather than a rewrite. **Test scale: ~2×** over a representative subset |
| v3.8 | Two-way Fixed Effects | No | Planned — MINOR. *Planned as v2.7+; promoted out of the unordered bucket.* Forward wiring from the v2.1 FE engine. **Test scale: ~2×** over the FE family |
| v3.9 | Standalone Data Transformation library (`Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Zscore_By`, `Decompose_By`, `Numeric_Complete_Cases`, `Dummy_Column`, `Interact`, `Model_Matrix`) | No | Planned — MINOR. *Planned as the second half of v2.2, then carried as the v3.3 remainder.* **The last regression milestone**, because it is **the ~10× axis-widener** — every added Transform value multiplies the response × predictor dispatch table, so it lands against the most mature harness the Regression track ever has |
| v3.10 | Bivariate / two-sample (one-sample t, two-sample t [equal-var / Welch / paired], F-test, Covariance) | No | Planned — MINOR. *Planned as v2.5, claimed as v3.6, briefly held at v3.5.* **The first milestone that is not Regression work** — a new sheet and a new analysis surface, held until the Regression template is feature-complete. F-test feeds a recommendation cell that selects the t-test variant; Covariance complements the existing `Correlation_Matrix`. **Test scale: additive** — a fixed set of cases on two small new datasets |
| v3.11 | Resampling & Simulation (bootstrap, Monte Carlo) | No | Planned — MINOR. *Planned as v2.4; claimed as v3.5, briefly held at v3.6.* The second non-Regression milestone. Pre-drawn random table (`Bootstrap_Random_Draws` named range) indexed at use time; non-volatile by design (every recalc reproduces the same draw). The artifact build seeds the table from a SHA-256 digest of the source CSV, so the draw is reproducible from the data alone. **Test scale: additive** — no new data at all |
| v3.12+ | Multi-group means (ANOVA), Fourier, Decision analysis | mixed | Unordered (deliberate — see Future section). *Planned as v2.7+.* Design-not-started, and nothing about their test cost sequences them |
| v3.13 | QR decomposition (`Coefficients_QR`) | No | Planned — MINOR. Replace the normal-equations path (form X'X, invert via MINVERSE) with QR decomposition (X = QR, solve Rb = Q'y). This avoids squaring the condition number, which is what causes Excel to return all-nan results above ~200 constructed columns on the current path. Lifts the O2 width guard's hard error threshold substantially. The engine already supports a drop-in replacement — `Coefficients` is the only function that forms X'X; a `Coefficients_QR` sibling would leave every downstream statistic untouched. The challenge is implementing Householder reflections or Gram-Schmidt in LAMBDA without exceeding Excel's formula-size limits on the intermediate arrays. **Test scale: near-additive** — re-runs existing models against a new solver and compares |


**Ladder rationale.** Under the interface definition above, exactly two milestones
break user inputs. Specification-Driven Regression took v2.0; everything after it
was additive and opt-in, forming a v2.x train directly analogous to Python's 3.x
line — one breaking 3.0 followed by years of large but non-breaking minors
(async/await, pattern matching) that never forced a new major.

**That train stops at v2.2 and resumes at v3.3.** Semver does not permit shipping
2.3.0 after 3.0.0, so once v3.0 was planned every unfinished v2.x milestone had to
move behind it and renumber. **That** renumbering was nothing more than that — the
claimed sequence was unchanged, each entry carries the number it was planned under,
and the DECISIONS entries keep their original headings as the record of when each
decision was actually made.

A **second** reordering came later and did change the sequence: everything from
v3.4 on is ordered by two keys — Regression work first, then the test-suite growth
each milestone forces. Three milestones changed number, three candidates left the
unordered bucket, and WLS keeps v3.7 by coincidence rather than by inheritance. It
is described in
[Ladder order from v3.4 on](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth)
below. The two renumberings are independent: the v3.0 one answers "what number can
this ship under," this one answers "what should ship next."

One milestone gets materially cheaper rather than merely renumbered. `Weight`
(WLS, now v3.7) carried a standing warning that shipping before v3.0 would force
the `[Weights]` argument it was designed around, with v3.0 then unwinding it across
the same ~24 functions. Behind v3.0 the constructor already owns the intercept, so
√w scaling is the first implementation instead of a rewrite. Two others —
Model Comparison (v3.4) and the Transforms remainder (v3.3) — keep their relative
order for the reason they always had: comparison is only honest once the numbers
being compared are unit-space comparable.

### Ladder order from v3.4 on: Regression work first, then test-suite growth

Everything at v3.4 and beyond is sequenced by two keys, in this order:

1. **Finish the Regression template first.** Every milestone that extends the
   Regression sheet, its spec block, or its engine ships before either milestone
   that opens a *new* analysis surface. Two-sample (v3.10) and Resampling (v3.11)
   are the only two of the latter, and they go last as a block.
2. **Within the Regression track, order by how much the test-model suite has to
   grow** — additive features first, per-model multipliers next, axis-wideners
   last, and within a tier the most commonly used feature first.

The full test-scale analysis — per-feature scale effect, the datasets each one
needs, and the covering-array philosophy the suite is built on — is
[docs/MODEL_TESTING_ASSETS.md § 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).
That table is the source; this ladder follows it.

**The Regression track — key 2 orders these:**

| Tier | Effect on the suite | Milestones |
|---|---|---|
| Additive | a fixed number of new cases | v3.4 Model Comparison |
| Near-additive | a variant over a few existing models | v3.5 `Cluster` · v3.6 `Time` Role |
| ~2× multiplier | re-runs a whole model family | v3.7 WLS · v3.8 Two-way FE |
| ~10× axis-widener | widens an axis every model is crossed against | v3.9 standalone transform library |

**Then the new surfaces:** v3.10 Two-sample, v3.11 Resampling, v3.12 Time Series.
All three are flat-cost to test — cheaper than four of the milestones ahead of
them — and they are held anyway, because key 1 outranks key 2.

**Why key 1 outranks key 2.** Test cost is the right tiebreaker *within* one
artifact; it is the wrong primary key across two. Everything in the Regression
track extends surfaces that already exist and is verified by the harness that
already exists — a milestone there is a spec column, an engine change, and more
cases in the same oracle. Two-sample and Resampling each need a new sheet writer,
a new layout, and a verification path that shares nothing with
`calculate_regression_spec_case`. Interleaving them means carrying two half-built
analysis surfaces at once, and it means the Regression template — the thing users
actually have — sits feature-incomplete for longer while effort goes somewhere
else. Deferring them costs nothing in rework: neither depends on any Regression
milestone, and neither is depended on by one.

Consequences worth stating plainly, because each moved a number:

- **`Cluster` (v3.5) and `Time` (v3.6) leave the unordered bucket and land ahead
  of WLS (v3.7).** A variance-estimator variant on a handful of models is cheaper
  to cover than a weighted re-run of one model per dispatch-pair family. `Time`
  earns its slot twice over: its calendar-dated dataset is the only asset that
  closes a Section-1 coverage gap *existing today* — the Sequence
  calendar-signature verdict has no test because no wired dataset carries real
  dates.
- **The time-series *sheet* splits off v3.6 and becomes v3.12.** v3.6 was
  carrying two unlike things under one number: engine wiring (the `Time` Role
  and the lag/difference semantics — Regression-track work) and a new worksheet.
  Key 1 says a new analysis surface ships in the trailing block, so the sheet
  goes there and v3.6 keeps the engine. Splitting also makes v3.6's
  near-additive rating honest: its test cost was always the dataset, never the
  sheet.
- **The standalone transform library leaves v3.3 for v3.9**, the last slot in the
  Regression track. It is the one item that widens the predictor-transform axis
  {None, Log}, and every widening multiplies the response × predictor dispatch
  table that every other model is scored against. v3.3 keeps its number for what
  actually shipped.
- **Two-sample (v3.10) still precedes Resampling (v3.11), and both precede Time
  Series (v3.12).** All three are flat-cost, so the ties break on value and on
  dependency: two-sample tests are the ToolPak-parity gap a user hits first;
  Time Series goes last of the three because it is the only one waiting on
  another milestone's asset.

This is a rework-minimizing default, not a commitment. The tool is single-user and
pre-release; a user pressing for one of these reorders it, and reordering means
editing the MODEL_TESTING_ASSETS table first and the ladder second.

**v3.0 is the second break, and it is not a failure of that plan.** The v2.0 record
says "one breaking restructure, never a second," and the reasoning behind it still
holds: no single v2.x feature justified another. What accumulated instead was the
*sum* of correctly-classified additive changes — 24 functions carrying
`[DF_Absorbed]`, 48 carrying `[Allow_Intercept]`, two constructor names for one
pipeline, and no representation for interactions at all. "Additive" is the property
that makes a change a MINOR; it was never evidence that the interface could absorb
it indefinitely. v3.0 unwinds that accumulation in the *engine* (the constructor
pipeline and the bounded `Model_Context`), which is non-breaking at the public
interface — a spec saved under 2.0.0 produces identical output under 3.0.0 — and
in the *layout* (the interaction spec columns, the audit column, the
materialization zone), which is the one genuine break the unwinding needs. Both
land inside v3.0, as separate stages behind separate pull requests: the engine
stages carry a zero-mismatch verification gate, and the layout stage carries the
`Breaking?` flag. The discipline v3.0 replaces the old rule with is stated in the
milestone entry below.

The next MAJOR is reserved for the next genuine interface break, whenever that
is. The layout stage bought room for two of them not to happen: M/N ship
reserved-and-unwired and the terminal materialization zone ships reserved, so
v3.1 and v3.2 are formula changes against columns that already exist.

Univariate shipped **before** Specification-Driven Regression despite the lower version
gap, as planned: its engine was already implemented and its sheet writer was wired into
`build_production.py` (now the single build script; during the split era it had a
separate `build_univariate.py`). Specification-Driven Regression was greenfield by
comparison, so the near-finished milestone shipped first.

**Fixed Effects breakage flag (v2.1) — RESOLVED as non-breaking.** The absorbed-df
correction is threaded as an **optional `[DF_Absorbed]` argument defaulting to 0**,
leaving the no-FE df path untouched, so a model with no Fixed-Effects Role behaves
identically to v2.0. FE therefore stays a MINOR at v2.1. The full argument-threading
rationale is in [DECISIONS.md § v2.1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

---

## v1.0 — Multivariate (OLS / MLR) — SHIPPED

The complete OLS package and the first stable release: full model-fit and ANOVA
statistics, coefficient inference, multicollinearity, residual and influence
diagnostics, cross-validation, information criteria, distributional exploration, and
prediction. `Include` argument supports stratified OLS natively. Five-zone Regression
sheet (model inputs · predictor summary · regression outputs · prediction · residual
output), Diagnostic Guide sheet, Regression Instructions sheet, Version History sheet,
seven pre-built diagnostic charts, conditional formatting on out-of-bounds
diagnostics, QC via `analyze_regression_sheet.py` across six configurations
(sparse/medium/full predictor sets × intercept/no-intercept).

The full v1.0 gate history is in the git history of this file rather than restated
here. The release predates the open-decisions convention, so DECISIONS.md has no v1.0
section.

---

## v1.1 — Univariate — SHIPPED 2026-06-29

*(Renumbered from v2.0 to v1.1 — a new sheet that changes no existing input is a MINOR
under the interface definition in Versioning & Release Conventions. Released in the
workbook Version History as 1.1.0.)*

The foundation layer. A single input column drives three coordinated sections:
**describe** the data, **visualize** its shape, then **formalize** that shape with a
fitted distribution. The skewness value from the descriptive section motivates the
distribution choice in the fitting section.

- **Descriptive statistics** — mean, median, mode, sd, variance, min/max/range,
  skewness, kurtosis, count, and a prominent missing count.
- **Histogram binning** — three methods side-by-side: Sturges, Scott,
  Freedman-Diaconis. Each method gets its own column chart (gap width 0) built
  from the computed bin table.
- **Distribution fitting** — eight candidates (Normal, Lognormal, Exponential,
  Weibull, Gamma, Triangular, Beta, BetaPERT) ranked in a single comparison
  table. Closed-form MLE where possible; search-based MLE for the two-parameter
  shape family (Weibull and Gamma via 1-D profile searches, Beta via a 2-D
  `Full_Factorial` spill — the Cartesian product of candidate parameters).
  Per-distribution Q-Q plots
  and histogram distribution overlays (post-release v1.1 leftovers, shipped
  with the next workbook build).
- **The MLE-via-grid reframing** — the wall was never "MLE without Solver"; it
  was "MLE in closed form." Grid search clears the no-Solver bar for the
  two-parameter likelihood class.
- **PDF functions dropped as unnecessary** — the histogram tables already
  compute per-bin probabilities as `CDF(upper edge) − CDF(lower edge)`,
  the PDF integrated exactly over the bin. The histogram overlay is
  delivered via combo charts with the CDF-delta columns as line series,
  not via `PDF_*` LAMBDAs at bin midpoints.

Design rationale and resolved decisions: [DECISIONS.md § v1.1](DECISIONS.md#v11--univariate).

---

## v1.2 — Workbook hardening & regression usability — SHIPPED 2026-07-03

*(Released in the workbook Version History as 1.2.0; renumbered from 2.1.0. A
non-breaking MINOR consolidating the v1.0 / v1.1 surface into a workbook that holds up
under hand-use by cost estimators and other non-git users. The technical core is
unchanged; every existing v1.0 / v1.1 input computes identically.)*

- **Name Manager notes on every worksheet-scoped name** — every worksheet-scoped
  named range (constructor closures, `RegChart*` chart feeds, helper formulas)
  carries a `Comment` describing what the name is for and which sheet zone
  owns it. The Name Manager `Comments` column is the in-workbook index for
  non-git users.
- **Identity-line data series on the diagnostic charts** — the QQ,
  Actual-vs-Predicted, and Studentized-vs-Leverage charts need a `y = x`
  reference line. Real data series pointing both `XValues` and `Values` at
  the same named range, with `ChartType = _XL_XY_SCATTER_LINES_NO_MARKERS`.
  The `Shapes.AddLine` approach was tried in v1.0 and abandoned — the
  pixel-coordinate positioning silently went wrong on resize.
- **Intercept-only and undersized-sample guards** — visible `n/a` tokens in
  the affected cells rather than silently-wrong numbers. Consistent with
  the library's "visible failure" philosophy.
- **`LOOCV_Residual` (new function, pair to `LOOCV_Prediction`)** — the
  leave-one-out residual vector `eᵢ / (1 − hᵢ)` in a single call. Makes
  `PRESS = SUMSQ(LOOCV_Residual(...))` an obvious idiom and unifies the
  LOOCV machinery under the Sherman-Morrison-Woodbury update.
- **Build-phase retry / RPC handling** — Excel's COM automation occasionally
  returns `RPC server unavailable` mid-`Calculate`. `_retry_on_open` in
  `build_production.py`; `main()` is split into two phases (sheet writes,
  no retry; recalculate + save, retry) so a transient failure in the
  short phase doesn't restart the multi-minute sheet-writing phase.

Design rationale and resolved decisions: [DECISIONS.md § v1.2](DECISIONS.md#v12--workbook-hardening).

---

## v2.0 — Specification-Driven Regression — SHIPPED 2026-07-05

*(Released in the workbook Version History as 2.0.0; renumbered from 3.0.0. The
canonical rename pass shipped inside this MAJOR as planned. The reserved
`Order`/`Transform` columns shipped unread by any formula, confirmed by construction.)*

The central idea: factor (categorical) and panel (fixed-effects) regression are not
new estimators — they are OLS on a transformed design matrix. Rather than telling
the user to manipulate their dataset into MLR form by hand, the sheet's control
block becomes a **declarative model specification spanning the entire source table**,
and `x_s()` is promoted from a column filter into a **model-matrix constructor**
that reads the spec and emits the numeric design matrix. Because every engine
function already consumes `x_s()`, the entire engine inherits the new capability
without a signature change.

The specification dissolves all three of the v1 sheet's hard-wired range names into
declarations: `y` (hard-wired) → the row whose Role is Response; `All_Xs`
(hard-wired span) → the table reference; `Regression_Sample_Include`
(hard-wired to `[Full_Data]`) → the column(s) whose Role is Filter. After v2.0
there is nothing left to hard-wire — the spec block plus one source-table
reference *is* the model.

- **Two-axis taxonomy (Role / Type) plus a structural Sequence axis** —
  Variable Role (Response / Predictor / Identifier / Filter / Omit),
  Predictor Type (Continuous / Categorical — closed, never grows), and
  Sequence (a structural flag, never grows). The full taxonomy and the
  cardinality rules are in
  [ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence).
- **Spec block A–L on the Regression sheet** *(A–O from v3.0, which appends the
  two interaction columns and the Design Columns audit)* — every column of the
  source table, one row per column. Cascading-relevance CF grays out cells
  irrelevant to the column's Role. The full A–O layout, the
  reserved-column policy, and the "Display derives, never feeds" rule
  are in [ARCHITECTURE.md § 4](ARCHITECTURE.md#4-the-model-spec-block-ao).
- **Spec-order assembly for `x_s()`** with the **level-vector split** for
  Categorical Predictors — training and prediction both call the same
  encoder with the same training level vector. Reference-level validation
  via `NA()`, not silent fallback.
- **Canonical rename pass** — every LAMBDA function renamed to the
  Title_Case_With_Underscores convention. The retained initialisms
  (`AIC`, `AICc`, `BIC`, `VIF`, `PRESS`, `CDF`, `NLL`, `LOOCV`, `PERT`,
  `R`, `QQ`, `GoF`, `MS`, `SS`, `SE`) and the per-pass sign-offs are in
  [ARCHITECTURE.md § 1](ARCHITECTURE.md#1-naming-convention).
- **One breaking restructure, never a second** — the spec-block
  changeover at v2.0 is the only MAJOR change to the Regression sheet.
  All v2.1+ additions are either additive sections on the
  already-restructured sheet or backward-compatible engine signature
  additions.
- **Supersessions** — separate Factor / Panel Regression sheets →
  one spec-driven sheet; WLS parallel function set → `Weight` Role
  axis value; single-axis "Predictor Type" → two-axis taxonomy.
  Full supersession record in
  [DECISIONS.md § Supersession log](DECISIONS.md#supersession-log).

Design rationale and resolved decisions: [DECISIONS.md § v2.0](DECISIONS.md#v20--specification-driven-regression).

---

## v2.1 — Sequence, fixed effects, and the forward-wiring chain — SHIPPED WITHIN 3.0.0

The 2.1 milestone bundles three coherent pieces that share the Sequence axis
and the FE Role: the Sequence/Base Period/longitudinal/serial-correlation chain
that the v2.0 work record says is "the v2.1 work that doesn't need the FE
engine proper" (now substantially complete in the repo), the FE engine
proper (`y_s`, `[DF_Absorbed]`, `Demean_By`, `Group_Mean`,
`Absorbed_Degrees_Of_Freedom`, `Is_Balanced_Panel`), and the sheet work that
activates the engine (FE Role dropdown, status-block validation, CI+PI
prediction layout, FE group dropdown, BFN cell flips active when FE is set).
Two-way FE remains a post-2.1 milestone (see
[v3.8](#v38--two-way-fixed-effects--planned)).

The 2.1.0 release ships as **a single release**, not as a sequence of preview
builds — the FE Role dropdown, the CI+PI prediction layout, and the FE
activation on the sheet are all gated to ship with the engine, so users never
see a workbook that says "FE is in the dropdown but the engine is forthcoming."

**Shipped since v2.0.0:**

- Sequence structural axis (spec column H) and reserved Base Period Δ (column I). **(PR #101)**
- Gap-aware `Difference_By` / `Lag_By` and the Base Period Δ layer, with the Sequence Spacing block. **(PR #102)**
- Sequence-aware Durbin-Watson. **(PR #103)**
- BFN panel Durbin-Watson. **(PR #105)**
- Grouping-key resolver (`Serial_Correlation_Group()` with the dormant Cluster branch). **(PR #106)**
- v1.1 leftovers — histogram distribution overlays and per-distribution Q-Q plots. **(PRs #96, #97, #99, #100)**
- `--skip-univariate` CLI option. **(PR #98; retired by the v3.0 split — Regression no longer ships a Univariate sheet)**
- Spec-driven QC refactor (`analyze_regression_spec.py` and `test_regression_spec_qc.py`). **(PR #103)**
- Durbin-Watson under FE — second cell + mutual gating (BFN + resolver releases). **(PRs #105, #106)**
- Second sample dataset — Auto MPG as the **Mileage Data** sheet, and the default `Source_Table` retarget to it, demonstrating the one-name-edit dataset changeover. **(PRs #123, #125, #126, #127)**
- Generalized VIF for the multi-level-categorical case. **(PR #124)**
- Sequence Period (column I) / Period In Use (column J) split — TODOs #1. **(PRs #111, #112, #114, #129)**
- FE engine proper: `Demean_By`/`Group_Mean` primitives, the fit-time demeaned `y_s()`/`X_s_Within()` pair, `[DF_Absorbed]` threaded through 23 engine functions, the FE Role dropdown + status-block validation + intercept×FE red flag, and the group-mean-recovery Prediction Outputs rebuild (`Group_Prediction_Interval`) — TODOs #2–#9. **(PR #128, with the array-shaped `Group_Mean`/BYCOL fix in PR #130)**
- FE-aware residual-output headers (conditional on the FE Role being active, not just on a label). **(PRs #131, #132)**
- Third sample dataset — **Production Lots**, a small unbalanced learning-curve panel with a natural Fixed Effects grouping column (Facility) and Sequence column (Fiscal_Year); the `--regression-dataset {auto_mpg,life_expectancy,production_lots}` build flag; extended the spec-driven QC oracle to compute and verify Fixed Effects (group/within-transform support); a guard against degenerate `df_residual` when FE absorbs too many degrees of freedom. **(PR #133)**
- QC-oracle rebuild for the shipped 9-value CI+PI/group-mean-recovery Prediction Interval shape (`regression_shared.RegressionPredictionInterval`, `analyze_regression_sheet.py`, `analysis_cache.py` schema v16, `tools/inspect_regression_sheet.py`), closing the 58 automated mismatches the pre-v2.1 6-value oracle was reporting; a `selected_group` validation guard so an unrecognized group name errors loudly instead of dividing by zero — TODOs #10. **(PRs #134, #136)**
- Build-time sheet tab order and colors; Auto MPG `Origin` values decoded to region labels (US/Europe/Asia) in the source dataset. **(pre-#128 commits)**
- `safe_activate()` / `safe_freeze_top_row()` guards so a headless/no-focus Excel session (no interactive desktop, focus denied by the OS) cannot abort the build when a sheet writer activates its sheet or freezes its header row. **(PR #135)**

**The verification gate — met by the automated evidence.**

Every numbered TODOs #1–#10 item is DONE and verified against a live build
(0 mismatches across all 12 spec-driven QC cases), and the FE engine is pinned
independently against `statsmodels` LSDV fits by `test_within_estimator`,
`test_df_absorbed_threading`, and `test_group_prediction_interval`. A written
human test plan for the FE engine existed and was **retired unrun**: by the
time it would have been executed, v2.1's features had already shipped to users
inside the 3.0.0 artifact, behind that artifact's own verifier pass. A hand-run
gate for code that is already released gates nothing. The v2.0 spec-block plan
was executed and signed off PASS on 2026-07-05 before its release; that record
lives in the git history of this file and in the T0–T19 cases now carried by
`tests/test_difference_by_verification.py` and
`tests/test_analyze_model_construction.py`.

**Still pending:**

- **Follow-on polish**: BFN critical values (**DEFERRED** — N,T-dependent
  bounds), Categorical × FE prediction encoding (**DEFERRED** — encode
  `x_new`/`x̄ᵢ` through `Dummy_Code` before the FE formula), and a
  residual-output relabel + Diagnostic Guide paragraph on residuals under FE
  (documentation-only). Full list in
  [TODOs.md § v2.1 follow-on polish](TODOs.md#v21-leftovers--follow-on-polish).
- **The Version History entry** — CLOSED. The 2.1.0 and 2.2.0 rows were never
  written at release, so the workbook's shipped changelog jumped 2.0.0 → 3.0.0
  and Fixed Effects, the Sequence axis, GVIF, and the Log transform reached users
  with no entry describing them. Both rows were backfilled into `_VERSIONS` at
  v3.1 and are present in the committed artifact.

Design rationale and resolved decisions: [DECISIONS.md § v2.1](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

---

## v2.2 — Transforms & Unit-Space Comparability — PARTIALLY DELIVERED

Wires the reserved spec-block column G (`Transform`) and ships the standalone Data
Transformation function library. The `Transform` dropdown gains one real value:
`Log`, applicable to a Continuous Response and/or Continuous Predictors. With
`Transform = None` (the default) the model computes identically to v2.1 — this
is what keeps the release non-breaking.

**Why this is its own milestone, not a Model Comparison sub-feature:** the Model
Comparison sheet is only an honest comparison tool if the numbers it lines up mean the
same thing. An R² computed on `Ln(Life expectancy)` and an R² computed on raw
`Life expectancy` are not the same quantity, and putting them in adjacent cells of a
comparison table without correction is exactly the kind of silent misfiring the
library's design philosophy exists to prevent. Transforms is the release that makes
cross-model comparison trustworthy, not just possible — and it is what completes the
Regression sheet as a fully functional deliverable, which is why it precedes the
Model Comparison convenience layer.

**Shipped:**

- **Column-G Log wiring** — `Response_Column()` and `X_s()` (renamed
  `Predictor_Columns()` by the v3.0 constructor pipeline) apply
  `Ln_Positive` in place when a Response or Continuous Predictor row
  declares `Log`; `Constructed_Column_Names()` relabels the column
  `Ln(name)`; the new structural twin `Constructed_Column_Transforms()`
  carries the per-constructed-column Log/None flag the Prediction Inputs
  band needs (a Categorical Predictor's dummy columns always read
  `None` — Log is disallowed there and flagged red, never silently
  applied). The v2.1 FE-demeaning wrappers (`y_s()`/`X_s_Within()`) needed
  no changes — they compose transform-then-demean automatically. The
  Prediction Inputs band takes a raw value and auto-logs it internally;
  its Training Mean spill emits the geometric mean for a logged column to
  avoid double-logging the default prediction. New catalog function
  `Ln_Positive(x, [include])`. Verified against a real learning-curve
  model (Production Lots, raw columns with `transform="Log"`) matching
  the pre-existing precomputed-log-column case to floating-point
  precision. Full design rationale in
  [DECISIONS.md § v2.2 Transform column wiring](DECISIONS.md#v22--transforms--unit-space-comparability).

**Not delivered — moved to v3.3.** Back-transformation and cross-model
comparability were deliberately excluded from this pass, which shipped only a
correctly-fitted log-space model. They are no longer v2.2 work: when the feature
train was resequenced behind v3.0, the unfinished half became its own milestone.
See [v3.3](#v33--transforms-remainder--shipped-dispatcher--duan--model-formula-label)
for the half that shipped and
[v3.9](#v39--standalone-data-transformation-library--planned) for the standalone
transform library.

Design rationale and resolved decisions: [DECISIONS.md § v2.2](DECISIONS.md#v22--transforms--unit-space-comparability).

---

## v3.0 — The engine-interface release — SHIPPED 2026-08-02

The engine-interface release. It responds to the 2026 architecture
review, whose findings share one shape: each decision was correct
in isolation and the cost is in the sum. Every design question below is
**resolved** in
[DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).
The release shipped as three stages plus the split (workbook 3.0.0,
Univariate artifact 1.0.0, later reunified), each a separate reviewable pull request with its own
verification gate (see the scope section below).

- **Bounded `Model_Context`** — engine signatures collapse from
  `(X_s, Y, [Allow_Intercept], [Include], [DF_Absorbed])` to
  `(X, Y, [Include], [Context])`. Exactly four elements (`Has_Intercept`,
  `DF_Absorbed`, `Response_Transform`, `Predictor_Transform`), materialized once
  into a spill range. `[Include]` is a permanent floor, not a transitional
  state — an n×1 row mask would break boundedness. Resolves F1.
- **Intercept relocation** — the intercept column moves into the constructor and
  `Design_Matrix` stops synthesizing it. `SS_Total` is redefined as the
  intercept-only residual sum of squares, which collapses the ones / absent / √w
  cases into one projection formula.
- **The constructor pipeline** — `Design_Columns()` / `Design_Response()` /
  `Predictor_Columns()` replace the `X_s()` / `X_s_Within()` name fork, applying
  declared stages in the fixed order `encode → transform → demean → intercept →
  weight`. Resolves F2.
- **Interaction spec columns** — M (Interaction Term) and N (Interaction
  Operation), with a closed operation vocabulary carrying a symmetry attribute.
  Resolves F6.
- **The materialization zone** — `Model_Context`, `Sample_Include`, and the
  Constructed Design Matrix at the far right, in increasing width, terminating in
  the unbounded zone. Resolves F3.
- **One workbook, one version** — the build emits one workbook
  (`Lambda_Library.xlsx`). The split into two artifacts that originally shipped
  with v3.0 was later reunified (see
  [Versioning & Release Conventions](#one-workbook-one-number) above). Resolves
  F4 and F8.

**What replaces "one breaking restructure, never a second."** That rule failed
because it constrained the *count* of breaks without constraining what could
accumulate between them. The v3.0 replacement constrains the accumulation
directly, in two rules that live in ARCHITECTURE because they bind every future
feature, not just this release:

1. **Properties of a fit travel in the context block, never as new optional
   arguments.** The reserved-slot pattern explicitly no longer applies to
   argument lists ([ARCHITECTURE.md § 7](ARCHITECTURE.md#7-reserved-spec-column-pattern-general)).
2. **Nothing may be placed right of the Constructed Design Matrix**, and
   materialized zones run in increasing width
   ([ARCHITECTURE.md § 4b](ARCHITECTURE.md#4b-the-materialization-zone)).

### Scope — **RESOLVED**

Which of the five pieces ship together was the one open question. They are
interdependent: the bounded context requires intercept relocation, which requires
the pipeline order; the materialization zone depends on the constructor pipeline
being settled first, or two variants of a soon-to-change architecture get
materialized; and the interaction columns and the audit column both touch the
spec-block layout. That argued for one release, against the general principle of
small increments.

**Resolved: shipped as three stages plus the split.** v3.0 shipped as
four reviewable pull requests — workbook 3.0.0, Univariate artifact 1.0.0 (later
reunified),
2026-08-02. Splitting the release into stages did not split the release: all four
land under the one version number, because they answer the one question together.

| Stage | Contents | Status |
|---|---|---|
| **1** | Constructor pipeline · intercept relocation | **Shipped** (merged as #148, QC gate cleared) |
| **2** | `Model_Context` · `Sample_Include` materialized · `[Has_Intercept]`/`[DF_Absorbed]` collapse into `[Context]` · two-name split (`Model_Context` constructor / `Fit_Context` reader) + 4 context accessors · rows 3-4 populated | **Shipped** (merged as #150, QC gate green) |
| **+ split** | Univariate Analysis → its own workbook (later reunified); Regression workbook → full Automatic | **Shipped** with v3.0 (merged as #151) |
| **3** | Interaction spec columns M/N (reserved) · Design Columns audit column + pre-flight width guard · Constructed Design Matrix zone | **Shipped** — the layout break; spec-driven verifier passed, no fitted number moved |

v3.3 onward is the feature train resequenced behind v3.0 — a different thing, and
not part of it.

The order was forced by the same dependencies listed above. Stage one carried a
verification property the others did not — **no number moves**, so the spec-driven
QC pass had to report zero mismatches across all twelve cases — which is why it
went first despite touching the most functions. Stage three went last because the
materialization zone could not be positioned until the constructor pipeline was
settled; placing it earlier would have materialized two variants of an
architecture that was still changing.

**Where the `Breaking?` flag comes from.** Stages 1-2 and the split are
non-breaking: they restructure the engine and the packaging, not the user-typed
spec block, so a spec saved under 2.0.0 produces identical output. Stage three is
the break, and it is worth being precise about what kind. Its three columns are
**appended** to the spec block, not inserted into it, so A–L keep both their
addresses and their meanings — a saved specification survives untouched and no
fitted number moves. What moves is every zone to the *right* of the spec block,
three columns over. A user who only fills in the spec block notices nothing; a
user whose own formulas point at cells on this sheet has to re-point them. That is
a real break, so the flag is **Yes**, but it is an address break rather than a
meaning break — the far more recoverable of the two. See
[DECISIONS.md § v3.0 shipped in stages](DECISIONS.md#v30-shipped-in-stages-the-layout-break-lands-last).

**Release contents — §3 through §7, all shipped:**

| Stage | Contents | Break |
|---|---|---|
| 1-2 + split | `Model_Context` · intercept relocation · constructor pipeline · Univariate split (later reunified; Regression → full Automatic) | **No** |
| 3 | Interaction spec columns M/N reserved-and-unwired · Design Columns audit column · Constructed Design Matrix zone + width guard | **Yes** (addresses right of the spec block shift three columns) |
| v3.1 | Interaction wiring — the constructor actually builds the columns stage 3 declared | MINOR (follows v3.0) |
| v3.2 | Full materialization of the design matrix | MINOR (follows v3.0) |

**Justification.** REVIEW.md's own sequencing note observes that F3 and F6 "all
want the same breaking change — resolving them separately spends three layout
breaks where one would do." That is why the two interaction columns, the audit
column, and the materialization zone all landed in the single stage-three change
rather than one per release. The wiring of each is then a formula change against a
column that already exists — precisely the reserved-column pattern, and exactly
how column G went live at v2.2 — so v3.1 and v3.2 stay genuinely additive MINOR
work with no second layout break behind them.

**What this scope does *not* cost.** F5 read as though the layout work would have
to be done twice — "the spec block is implemented twice; a layout change touches
both writers." It is not, and it does not. `write_sheet_regression.py` imports
the spec-block writers from `write_spec_block.py` and calls them,
so the two interaction columns, the audit column, and the materialization zone
each land in **one** writer. That single-implementation structure is part of what
makes this scope affordable, and it is why F5 does not appear in the release
above. See
[DECISIONS.md § v3.0 spec block](DECISIONS.md#the-spec-block-is-implemented-once-not-twice).

The counter-argument weighed and answered: v3.0 is a large release that is hard to
verify in one pass, and the human test plan for it is substantial. The answer was
not to shrink the release but to stage it — four pull requests, each a reviewable
diff with its own verification gate, under one version number. The engine stages
cleared a zero-mismatch numeric gate before the layout stage was allowed to move a
single column, which is what made the layout stage's own gate meaningful: with the
numbers already pinned, any mismatch it produced could only have come from the
layout.

**What stage one actually cost, against the estimates in this entry.** Two numbers
moved and both are recorded in DECISIONS: `Has_Intercept` survives in **thirteen**
functions rather than the estimated seven (the R²/sums-of-squares chain needs it as
an identifier, which the estimate did not anticipate), and `R_Squared` turned out
to be a **third** LINEST `const` site — the one that would have failed silently,
since LINEST reports the uncentered R² under `const = FALSE`. Neither changes the
design; both are why stage one is the one with the zero-mismatch gate.

---

## v3.1 — Interaction wiring — SHIPPED 2026-08-03

The first of the two follow-ons v3.0 stage 3 bought room for, and the payoff on
the reserved-column policy. Spec columns M (Interaction Term) and N (Interaction
Operation) shipped at 3.0.0 validated, flagged, and **read by nothing**; this
release makes the constructor read them. No column moved, no address changed, and
a specification with M and N blank — every specification saved under 3.0.0 —
computes identically. Workbook 3.1.0, library 3.1.0.

**What it builds.** `Predictor_Columns()` resolves the operand named in M to its
spec row, encodes it with the **same** `blk()` that encodes the declaring row, and
combines the two blocks pairwise under the operation in N:

| Operands | Columns |
|---|---|
| Continuous × Continuous | 1 |
| Continuous × Categorical | L−1 |
| Categorical × Categorical | (L₁−1)(L₂−1) |

The columns land immediately after the declaring row's own block, so the matrix
stays in spec order and the per-row audit describes adjacent columns. Headers are
the two operands' own names joined by the **operation's own symbol** —
`Weight × Origin: US`, `Weight ÷ Displacement` — so a header says what was done
as well as what it was done to. (v3.0 specified R's colon; three operations make
one shared separator ambiguous, and it collided with the `": "` already inside a
level-qualified name. Superseded — see DECISIONS.)
`Constructed_Column_Names()` and `Constructed_Column_Transforms()` gate
interactions identically — the twin property is what keeps the header strip and
the transform strip exactly as wide as the matrix.

**What it enforces.** The operation axis stays closed (`Product` · `Difference` ·
`Ratio`), with `Ratio` returning `NA()` on a zero denominator rather than a bare
`#DIV/0!`. An operand that is an *excluded* Predictor still builds — the
flagged-amber marginality case. An operand that is not a Predictor, or matches no
column, contributes nothing and stays flagged red: the constructor degrades to
the main effect rather than taking the whole sheet down for one mistyped cell,
the same way an invalid Reference Level already degrades. A row pointing at
itself under `Product` is the documented quadratic term. Two-way only.

**The audit column earns its keep here.** Column O gained `k(row) × k(operand)`
off the *same* per-row width helper it already used, so it cannot disagree with
the constructor about how wide a categorical operand is — and the pre-flight
width guard reads that total, so Status × Country on the WHO data announces its
155 columns before anything is built.

**One thing is deliberately not automatic.** An interaction row in the Prediction
Inputs band is an independent input, not recomputed from its operand rows. The
default state is self-consistent (the whole band sits on the design matrix's own
centroid); a partial override is not, and the band's header note says so rather
than silently rewriting one user input because another changed. Rationale, and
the deferred derive-on-change design, in
[DECISIONS.md § v3.1](DECISIONS.md#v31--interaction-wiring).

**Verification — the Excel gate ran and cleared.** Three QC cases
(`interaction_continuous_product`, `interaction_quadratic_self_product`,
`interaction_categorical_broadcast`) join the spec-driven oracle, covering all
three width regimes, and `tests/test_interaction_wiring.py` pins the semantics
headlessly against the Python mirror. `build_production.py --verify --no-launch`
was then run on a machine with Excel and reported **no spec-driven QC
mismatch** — neither on the three new cases nor on the twelve pre-existing ones,
which is the behaviour-preserving property this release had to hold: a spec with
M and N blank must compute exactly what it computed under 3.0.0.

The run's one reported failure, `[Univariate] sheet is missing`, was the verifier
checking a sheet the Regression-only artifact stopped carrying at v3.0 (during the
split era). It was a false positive against the post-split layout, not a result —
`skip_univariate` reaches the force-calc list but does not yet guard the check
itself — moot after reunification: the unified build produces both sheets, so the check cannot fire.

Design rationale: [DECISIONS.md § v3.1](DECISIONS.md#v31--interaction-wiring),
building on the representation decisions in
[§ v3.0](DECISIONS.md#interactions-are-declared-with-two-spec-columns).

---

## v3.3 — Transforms remainder — SHIPPED (dispatcher + Duan + model formula label)

*Planned as the second half of v2.2. Moved after v3.0 when the feature train was
resequenced — see the [ladder rationale](#versioning--release-conventions).*

The column-G `Log` wiring shipped at [v2.2](#v22--transforms--unit-space-comparability--partially-delivered);
this milestone finishes the release. Unit-space dispatch (eight new catalog
functions under the `Back-Transformation` subcategory), Duan back-transformation
with the `[Method]` Duan/Naive toggle, the original-units Prediction Outputs
column (AL), the original-units residual columns (AZ/BA), the Model Formula
readout, and the `Comparison_*` sheet-scoped named ranges are now in
production.

**The standalone transform library is no longer part of this milestone.** It was
planned here, never started, and now ships as
[v3.9](#v39--standalone-data-transformation-library--planned), the last milestone
in the Regression track — it is that track's single most expensive item to test
(the ~10× axis-widener), and nothing between here and there waits on it. See
[Ladder order from v3.4 on](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth).

- **Unit-space dispatcher, RESOLVED & SHIPPED** — eight catalog functions
  (`Smearing_Factor`, `Back_Transform_Response`, `Unit_Space_Predictions`,
  `Unit_Space_Observed`, `Unit_Space_Residuals`, `Unit_Space_R_Squared`,
  `Unit_Space_Adjusted_R_Squared`, `Unit_Space_RMSE`) under the
  `Back-Transformation` subcategory. The
  transform pair is read off `Fit_Context()` rather than passed as
  positional arguments. `SWITCH` on the six recognised `(response, predictor)`
  pairs and `NA()` outside. The dispatcher is the first deliberate departure
  from "one canonical name, one LAMBDA" — justified by the combinatorial
  blow-up the exception avoids.
- **Prediction back-transformation, RESOLVED & SHIPPED** — Duan's smearing
  estimator as the default, with a per-cell `Back_Transform_Method` toggle
  (`Duan` default | `Naive`). Naive is biased (Jensen's inequality); Duan is
  unbiased under iid residuals. Caveat visible on the sheet as a note on the Back-Transform label at AG4.
- **Model Formula label, RESOLVED & SHIPPED** — the sheet-scoped
  `Model_Formula()` catalog closure, rendered by a labelled readout on
  row 1 of the terminal Constructed Design Matrix zone, right of that zone's heading (header two columns over, readout three columns past the header). It shipped as an
  inline expression at `AA2:AB2` and moved off that cell afterwards: row 2
  of the Regression Outputs zone wraps and AutoFits, so the sheet's longest
  string set the height of the whole header row. Row 1 of the design-matrix
  zone is the one row the matrix itself can never reach, so with wrap off the
  caption overflows across as much empty width as it needs. Built from the response-name
  lookup (which already emits `Ln(name)` when Log), `Allow_Intercept`,
  `Constructed_Column_Names()`, and the FE-name suffix gated by the Fixed
  Effects count. The mixed Log/None predictor case
  renders correctly with no extra work because `Constructed_Column_Names()`
  already emits `Ln(name)` per logged predictor, level-qualified dummy
  names, and `left × right` interaction names.
- **`Comparison_*` named ranges, RESOLVED & SHIPPED** — sheet-scoped
  `Comparison_Anchor` (`$AF$2`), `Comparison_Headline_GoF` (`$AH$6:$AH$8`),
  `Comparison_Model_Formula` (the Model Formula readout in the §4b band).
  v3.4 Model Comparison reads from these surfaces — by NAME, which is why
  moving the readout off `$AB$2` cost its consumer nothing.
- **Statistics with a unit-space counterpart:** R², Adjusted R², RMSE
  (SHIPPED at v3.3); LOOCV RMSE / MAE and the LOOCV residual column
  (SHIPPED at v3.4 — v3.3 left cross-validation in fit space; the v3.4
  unit-space LOOCV block closes that gap, in original units). AIC / AICc /
  BIC deferred (likelihood depends on the Jacobian of the transformation;
  the "right" comparison is on the original response's likelihood, not the
  transformed one's).
- **Standalone transform library, remainder — MOVED to
  [v3.9](#v39--standalone-data-transformation-library--planned).**
  `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Zscore_By`,
  `Decompose_By`, `Numeric_Complete_Cases`, `Dummy_Column`, `Interact`,
  `Model_Matrix` (`Ln_Positive` shipped early with the column-G wiring
  above). The full taxonomy and the `""`-vs-`NA()` row-alignment
  convention are in
  [ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy),
  and remain version-independent.

Design rationale and resolved decisions: [DECISIONS.md § v3.3](DECISIONS.md#v33--transforms-remainder-unit-space-dispatch--duan-back-transformation--model-formula-label),
recorded there under v3.3 (the original v2.2 entries describe the standalone
library, now [v3.9](#v39--standalone-data-transformation-library--planned)).

---

## v3.4 — Model Comparison Sheet — PLANNED

*Planned as v2.3. Moved after v3.0 when the feature train was resequenced — see the [ladder rationale](#versioning--release-conventions).*

Every v2.0 Regression sheet already exposes a fixed-height, fixed-position **Model
Spec status block** (response in effect, constructed column count, level-qualified
names, included row count, error state). That block is an interface, not just a
display — the Model Comparison sheet is what happens when a second sheet is allowed
to *read* it. No new modeling capability is added; this is purely a cross-sheet
aggregation and navigation layer, which is why it is a MINOR.

- **Model registry** — one row per registered Regression sheet, hyperlink
  + display text from `Model_Formula_String(anchor_cell)`, link
  target a fixed anchor cell inside the spec block.
- **GoF table** — R², Adjusted R², AIC, AICc, BIC, PRESS, LOOCV,
  F-statistic, F p-value, n, k. References the **unit-space
  headline** statistic from v3.3 — a logged model and a level
  model line up as comparable quantities by construction.
- **Shared prediction inputs** — the Comparison sheet is the source;
  individual Regression sheets pull from it via `XLOOKUP` keyed on
  spec name, so one shared "what-if" scenario drives every
  registered model simultaneously.
- **Interface contract, RESOLVED** — three sheet-scoped named ranges
  per Regression sheet (`Comparison_Anchor`,
  `Comparison_Headline_GoF`, `Comparison_Prediction_Output`) become
  part of the library's public interface the moment they ship.
  The changelog entry for v3.4.0 must name them explicitly so
  the commitment is discoverable.

**Test assets — additive (~1×).** No new data and no new models: the sheet reads
models the suite already has (M1, L2, P2 supply ≥3 registered models with shared
prediction inputs). One **mismatched-predictor-set pair** (M1 vs M14) is added to
exercise the `XLOOKUP [if_not_found]` open question. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 1](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

Design rationale and resolved decisions: [DECISIONS.md § v2.3](DECISIONS.md#v23--model-comparison-sheet),
recorded there under the original milestone number.

---

## v3.5 — `Cluster` Role (clustered-robust SEs) — PLANNED

*Planned as v2.7+ and carried in the unordered bucket until the
[ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth)
gave it a number: it is Regression work, and a variance-estimator variant over a
few existing models is cheaper to cover than anything below it in that track.*

A `Cluster` value on the Role axis (at most one, per the same cardinality rule as
Response / Time / Weight), producing a clustered-robust V_β. It has **partial
forward wiring already in the codebase**: `Serial_Correlation_Group()` carries a
dormant `Cluster` branch (the reserved-spec-column pattern, pinned by
`tests/test_serial_correlation_group_resolver.py`), so the resolver side lights up
by retargeting one name. The engine side — the cluster-robust variance estimator
itself — is not started. Shipping it also lifts the v2.1 `n/a — engine forthcoming`
token on the BFN cell when Cluster is active.

**Test assets — near-additive.** Within-group correlated data, which the wired
datasets already supply: Production Lots' three facilities are enough to start (and
deliberately few, so the small-cluster warning path is exercised); `Grunfeld`
arrives with v3.8 and provides 10–11 proper clusters. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 2](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

---

## v3.6 — `Time` Role + lag/difference semantics — PLANNED

*Planned as v2.7+ and carried in the unordered bucket until the
[ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth).
It is the one milestone that closes a coverage gap existing **today**.*

**Engine work only — the sheet this used to carry is now [v3.12](#v312--time-series-analysis-sheet--planned).**
This entry was written as "`Time` Role + time series" and bundled two unlike
things: time-index semantics on the Regression spec block, and a whole new
worksheet. Key 1 of the ladder ordering says a new analysis surface ships in the
trailing block, so `Moving_Average`, `Exponential_Smoothing`, the forecast
output and its error metrics move to v3.12. What stays here is what extends the
Regression artifact.

The `Time` Role gives a column time-index semantics, which is what cross-sheet
`Lag_By` / `Difference_By` calls resolve against — today those functions read
the `Sequence` axis, which answers a narrower question (what orders the rows)
than "what is the calendar". Partially forward-wired via the v2.1 Sequence axis:
`Sequence_Deltas`, `Sequence_Delta_Spectrum` and `Base_Period_Delta_Candidate`
already compute the spacing this Role needs to interpret. The open design
question stands: **can a column be both `Sequence` and `Time`**, or are they
mutually exclusive?

The milestone also wires the calendar-dated dataset, and that is the half with a
consumer waiting on it in two directions — the Section-1 coverage gap below, and
every case on the v3.12 sheet.

**Test assets — near-additive, plus one gap closed.** This milestone brings the
first **calendar-dated** dataset into the workbook (~144 rows, AirPassengers-shaped,
with a real date column). No wired dataset carries dates today, which is why the
Sequence **calendar-signature verdict** (~28–31 / ~90–92 / ~365–366-day spacing
clusters) is the single uncovered axis in the Section-1 coverage matrix. That test
becomes writable as soon as the dataset is wired — *before* the `Time` Role itself
ships. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 3](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order)
and [§ 1.5](MODEL_TESTING_ASSETS.md#15-coverage-matrix).

---

## v3.7 — `Weight` Role (WLS) — PLANNED

*Planned as v2.6 and claimed as v3.7; it keeps that number, but now as the first
~2× item in the Regression track rather than by inheritance — the
[ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth)
put `Cluster` and `Time` ahead of it and Two-sample and Resampling behind it.*

A `Weight` value on the Role axis (see
[ARCHITECTURE.md § 3](ARCHITECTURE.md#3-variable-role--predictor-type--sequence)
for the cardinality rule). Three-stage scope: user-supplied weights →
variance-driver-derived weights → FGLS. This milestone ships the first stage only.

**The mechanism changed at v3.0; the feature did not.** This milestone was
planned as "a single optional `[Weights]` argument (default uniform), following
the `[DF_Absorbed]` precedent." That threading is superseded — with the intercept
owned by the constructor, √w scaling of the design matrix and response yields the
exact WLS estimator, standard errors, leverage, and Cook's distance, because the
intercept column correctly becomes √w rather than remaining ones. **WLS becomes a
constructor concern rather than an engine argument.** Everything else carries
forward unchanged: the `Weight` Role, its cardinality rule, the status-block
validation, and the three-stage scope. Weights are still declared in the spec
block.

One trap this avoids, recorded because it is the kind that ships silently:
`DEVSQ(√w ⊙ y)` is *not* the weighted total sum of squares — it centres on
mean(√w·y) rather than ȳ_w — so a naive "scale everything by √w" implementation
would leave SS_Total, and therefore R², wrong under WLS with no error anywhere.
The v3.0 projection form of `SS_Total` is correct by construction. See
[DECISIONS.md § v3.0 SS_Total](DECISIONS.md#ss_total-redefined-as-the-intercept-only-residual-sum-of-squares).

**Sequencing note — resolved.** This entry used to warn that shipping before v3.0
would force the `[Weights]` argument after all, with v3.0 then unwinding it across
the same ~24 functions. Resequencing the feature train behind v3.0 is what removes
that cost: the constructor owns the intercept before WLS is built, so √w scaling
is the *first* implementation rather than a rewrite of a threaded argument. This is
the clearest single case for the resequencing, and it was already visible in the
review's first sequencing implication — that changing the mechanism is cheaper
before this milestone than after.

**Test assets — ~2× over a representative subset.** Grouped/heteroskedastic data
with a natural weight column (R/MASS `Insurance`, 64 rows, claims with exposure
`Holders`, or a grouped-mean aggregation of an existing dataset). The plan is
**weighted variants of ~6 representative Section-1 models — one per dispatch-pair
family — not the whole suite**; that bound is what keeps a ~2× item from becoming a
full re-run. The trap above is an oracle assertion, not just prose:
`DEVSQ(√w ⊙ y)` ≠ weighted SST. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 4](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

---

## v3.8 — Two-way Fixed Effects — PLANNED

*Planned as v2.7+ and carried in the unordered bucket until the
[ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth).*

The trio `Absorb_Two_Way_Fixed_Effects` (alternating-projection demeaning for
unbalanced panels), `Demean_Two_Way_Balanced`, and
`Fixed_Effects_Convergence_Check`, plus the two-way `Is_Balanced_Panel` check,
lifting the v2.1 one-FE-variable status-block error, and the two-way prediction
question (group intercepts are not recoverable as simple group means). Forward
wiring from the v2.1 FE engine; the one-way-scope rationale is in
[DECISIONS.md § v2.1 scope](DECISIONS.md#v21--sequence-gap-aware-longitudinal-serial-correlation-diagnostics-fixed-effects).

**Test assets — ~2× over the FE family.** A balanced two-factor panel (R
`Grunfeld`, 200 rows, 10 firms × 20 years) plus an **unbalanced variant** (rows
deleted) to exercise `Is_Balanced_Panel` and the convergence check, with the FE
family (P1/P2/L8 analogues) re-run two-way. `Grunfeld` also back-fills v3.5's
cluster count. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 5](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

---

## v3.9 — Standalone Data Transformation library — PLANNED

*Planned as the second half of v2.2, then carried as the v3.3 remainder. Moved to
the end of the **Regression track** by the
[ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth) —
it is the only item that widens an axis every model is crossed against, so it lands
against the most mature harness that track ever has.*

`Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Zscore_By`, `Decompose_By`,
`Numeric_Complete_Cases`, `Dummy_Column`, `Interact`, `Model_Matrix`.
(`Ln_Positive`, `Demean_By`, `Group_Mean`, `Lag_By`, and `Difference_By` shipped
early, with the v2.2 column-G wiring and the v2.1 FE work.) The full taxonomy and
the `""`-vs-`NA()` row-alignment convention are in
[ARCHITECTURE.md § 5](ARCHITECTURE.md#5-data-transformation-taxonomy).

**Test assets — the ~10× axis-widener, and no new data.** Every new `Transform`
value widens the predictor-transform axis that today holds {None, Log}, and each
widening multiplies the response × predictor dispatch table (six recognized pairs
now). The existing datasets cover all of them. **Sequencing within the milestone**
matters as much as its position on the ladder:

1. **The additive helpers first** — `Numeric_Complete_Cases`, `Dummy_Column`,
   `Interact`, `Model_Matrix`. Standalone LAMBDAs with a fixed test count; they
   widen nothing.
2. **Predictor-side location/scale transforms next** — each adds dispatch pairs
   but no back-transformation semantics.
3. **Any response-side extension last.** A response transform also multiplies the
   back-transformation / unit-space semantics — what *is* the smearing analogue
   for Zscore⁻¹? — which is the single most expensive kind of growth this project
   has.

See
[docs/MODEL_TESTING_ASSETS.md § 2 item 6](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

---

## v3.10 — Bivariate / Two-sample — PLANNED

*Planned as v2.5, then claimed as v3.6, briefly held at v3.5. It lands here because
the [ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth)
ships all remaining Regression work first: this is the first milestone that opens a
**new analysis surface** rather than extending the Regression sheet. It still
precedes Resampling — both are flat-cost to test, and two-sample is the
ToolPak-parity gap a user hits first.*

**Nothing about this milestone got harder by waiting.** It depends on no Regression
milestone and none depends on it, so the deferral is pure sequencing: a new sheet
writer, a new layout, and a verification path that shares nothing with
`calculate_regression_spec_case` all cost the same whenever they are built.

`T_Test_OneSample`, `T_Test_TwoSample` (equal-variance / Welch / paired
variants — the 3-way flag or separate `paired` boolean is the open design
question, see [DECISIONS.md § v2.5](DECISIONS.md#v25--claimed), recorded there under the original number),
`F_Test_Variance` (feeds a recommendation cell that selects the t-test
variant), `Covariance_Matrix` (complement to the existing
`Correlation_Matrix`). Dedicated sheet layout with test selector and
F-test assumption check.

**Test assets — additive.** A fixed set of cases on two small new datasets: a
two-group dataset (R `ToothGrowth`, 60 rows, or the in-repo `Status` split of Life
Expectancy) and a **paired** dataset (R `sleep`, 20 rows). Cases: equal-variance t,
Welch t, paired t, and the F-test of variances feeding the selector cell. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 7](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

---

## v3.11 — Resampling & Simulation — PLANNED

*Planned as v2.4, then claimed as v3.5, briefly held at v3.6. The second
non-Regression milestone, behind Two-sample — see the
[ladder reordering](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth).*

Bootstrap confidence intervals and Monte Carlo simulation. Validated as worthwhile
differentiators by their presence in Pyrcz's Excel demos and squarely in cost-estimation
territory (three-point estimates, MCS, risk analysis). These depend on nothing else on
the ladder. Bootstrap and Monte Carlo pair naturally and may share a single sheet.

- **`Bootstrap_Random_Draws` table** — sheet-scoped named range
  holding a pre-drawn uniformly-distributed table, indexed at use
  time. Seeded from a SHA-256 digest of the source CSV, so the same
  data reproduces the same draw. Non-volatile by design.
- **`RANDARRAY()` rejected** — silently re-drawing per recalc is
  the opposite of the library's auditability philosophy. A cost
  estimator who sees a 90% CI of (4.2, 5.7) one moment and
  (4.0, 5.9) the next, with no record of which sample produced
  which, has not been given a tool.
- **Functions** — `Bootstrap_CI(data, stat_lambda, n_resamples,
  alpha, [include])`, `MC_Percentile(dist_params, n_samples,
  percentile)`, `PERT_Sample(min, mode, max, n_samples)`.

**Test assets — additive, and no new data at all.** The seeded pre-drawn
`Bootstrap_Random_Draws` table *is* the asset; Production Lots (n = 51) is the
natural small-n bootstrap target (slope CI on P3), and PERT/MC cases need only
parameter cells. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 8](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

Design rationale and resolved decisions: [DECISIONS.md § v2.4](DECISIONS.md#v24--resampling--simulation),
recorded there under the original milestone number.

---

## v3.12 — Time Series Analysis sheet — PLANNED

*New milestone, and the last of the three new analysis surfaces. It takes this
slot for two reasons: [key 1](#ladder-order-from-v34-on-regression-work-first-then-test-suite-growth)
puts new surfaces behind all Regression work, and within that block it goes last
because it is the only one waiting on another milestone's asset — the
calendar-dated series [v3.6](#v36--time-role--lagdifference-semantics--planned)
wires. It **absorbs the sheet half of v3.6**, which was carrying a worksheet
inside the Regression track by historical accident.*

A dedicated `Time Series` sheet organized around the **identification
workflow**, in the order an analyst actually works: declare the series and how
it is differenced → look at the ACF and PACF → test whether what is left is
white noise and whether the level is stationary → decompose, then smooth and
forecast. Each step is a zone, and every zone reads the *constructed* series, so
changing the differencing order re-drives the whole sheet — the same one-edit
promise `Source_Table` makes on the Regression sheet.

This is also the roadmap's clearest parity-exceeding claim. The Analysis ToolPak
has no ACF, no PACF, and no stationarity test of any kind; its entire
time-series offering is Moving Average and Exponential Smoothing, both of which
land here as one bullet out of six.

- **Series spec block** — the Regression spec-block idiom narrowed to one
  series: Time column, Value column, optional Group, a `Log` toggle reusing
  `Ln_Positive` and the established Transform vocabulary, regular differencing
  order `d`, seasonal order `D`, and period `m`. **Differencing is declared, not
  hand-built.** The reasoning is the Transform column's, exactly: transforming
  or narrowing the series changes the model, so it belongs in a specification
  the sheet can read and report, not in a helper column whose existence no
  status cell knows about. The spacing / regularity verdict reuses the v2.1
  Sequence machinery — `Sequence_Deltas`, `Sequence_Delta_Spectrum`,
  `Base_Period_Delta_Candidate` — rather than implementing the same question
  twice.
- **ACF / PACF tables and charts** — `Autocorrelation(data, max_lag,
  [include])` spilling lag / r / SE, and `Partial_Autocorrelation` by the
  **Durbin–Levinson recursion** rather than a Yule–Walker solve: no
  `Gram_Inverse` dependency, no singularity path, and a recursion is
  inspectable lag by lag in a way a matrix inversion is not. Two column charts,
  with the significance bands as **real data series, never shapes** (the
  `_add_identity_line` rule — a shape is positioned in plot-area pixels fixed at
  creation time and silently goes wrong on resize). Bartlett's widening band on
  the ACF, ±1.96/√n on the PACF, with `Bartlett_Bands` exposed as its own
  function so the band is a value you can click.
- **White-noise test** — `Ljung_Box_Q(acf, n, h, [df_fitted])` and
  `Box_Pierce_Q`. The optional `df_fitted` is what lets one cell serve both
  audiences: a raw series tests on `h` degrees of freedom, fitted residuals on
  `h − p − q`. It is also what would let this statistic appear on the Regression
  sheet's diagnostic band beside `Durbin_Watson` — see the open questions.
- **Unit-root and stationarity, as a pair** — `ADF_Statistic(data, lags, spec)`
  built as a regression of Δy on y(t−1) plus lagged differences and evaluated
  **through the existing OLS engine**, not a second estimator; and
  `KPSS_Statistic(data, lags, spec)` for the complementary null. They ship
  together deliberately: their nulls are opposites, the four
  reject / fail-to-reject combinations are the actual reading, and shipping ADF
  alone invites the standard error of treating "failed to reject" as "is
  stationary". The sheet names which of the four cells the series lands in.
- **Critical values are a table, not a formula** — both statistics have
  non-standard null distributions, so each ships a companion
  `*_Critical_Value(n, spec, alpha)` lookup. The sheet shows the statistic, the
  critical value it was compared against, and the verdict side by side, so the
  comparison stays clickable. A `p`-value alone would be the one number in this
  library a user cannot interrogate.
- **Classical decomposition** — additive and multiplicative: trend by centered
  `Moving_Average` at period `m`, seasonal indices as the de-trended period
  means (normalized to sum 0 or average 1 respectively), remainder by
  subtraction or division. `Classical_Decomposition` returns the three columns
  as one spill, charted under the time plot.
- **Smoothing and forecast** — `Moving_Average(data, window, [include])` and
  `Exponential_Smoothing(data, alpha_smooth, [include])`, forecast output scored
  by MAE / RMSE / MAPE, and an actual-vs-smoothed chart. **Moved here from
  v3.6.** The `alpha_smooth` argument name is deliberate — `alpha` is already
  the significance level throughout the catalog, and a smoothing parameter
  sharing that name in a library whose whole premise is clickable formulas would
  be a standing trap.
- **Interface contract** — the `TSChart`-prefixed sheet-scoped named ranges the
  charts read (following the `RegChart` precedent in `_setup_local_names`) and
  the series-constructor closures become part of the public interface the moment
  they ship. The v3.12.0 changelog entry must name them, exactly as v3.4's
  entry must name its three.

**Open design questions:**

- **Does the sheet declare its own series, or consume a Regression sheet's
  `Time` Role?** Leaning independent — this is not a Regression sheet, and
  cross-sheet reading is the v3.4 Comparison sheet's job. Consuming would couple
  a new surface to a Regression sheet's spec for no capability gain.
- **Is the differenced series a materialized column, or a constructor closure
  each consumer evaluates?** The ARCHITECTURE §4b lesson — one materialized
  spill everything reads, rather than N re-evaluations of the same
  transformation — argues for the column, as does the fact that ACF, PACF, both
  tests and the decomposition all want the identical vector.
- **Should `Ljung_Box_Q` also appear on the Regression sheet**, next to
  `Durbin_Watson`? Durbin–Watson tests lag 1 only; a portmanteau test over the
  first h lags is strictly more informative about the residuals, and the
  `df_fitted` argument already exists to make it correct there.

**Test assets — additive, and no new data.** The calendar-dated monthly series
v3.6 wires (~144 rows, AirPassengers-shaped) is the only dataset this milestone
needs, because **one series supplies both sides of every verdict**: in levels it
is the non-stationary seasonal case (ADF fails to reject, KPSS rejects, ACF
decays slowly, decomposition is multiplicative), and log-differenced at
`d=1, D=1, m=12` it is the stationary counterpart with the opposite verdict on
both tests. Oracles come from `statsmodels.tsa` (`acf`, `pacf`,
`acorr_ljungbox`, `adfuller`, `kpss`, `seasonal_decompose`) — the same library
the Regression oracle already depends on — through a `TimeSeriesSpecCase`
registry mirroring `RegressionSpecCase`. One guard state, a series shorter than
`max_lag`, pins the PACF recursion's degenerate path. See
[docs/MODEL_TESTING_ASSETS.md § 2 item 9](MODEL_TESTING_ASSETS.md#section-2--assets-for-roadmap-features-in-ladder-order).

---

## v3.13+ — Unordered (no claim; planned as v2.7+)

What is left after the reordering gave the other candidates numbers, and after
the Time Series sheet took v3.12: multi-group means (ANOVA, with Tukey HSD or
Bonferroni post-hoc comparisons); Fourier analysis and Decision analysis
(long-tail, out of planning horizon).

These stay unordered because **nothing about their test cost sequences them** —
ANOVA-as-regression needs only `warpbreaks` plus the existing categorical
machinery, and the other two are design-not-started. A user pressing for one would
reorder it; absent that signal, a single maintainer should not pre-order work
nobody is asking for. That is precisely how the Time Series sheet got a number:
it was asked for.

---

## Univariate releases — the split (SHIPPED, later reunified), then the grid shrink (Weibull/Gamma SHIPPED; Beta SUPERSEDED)

The Univariate sheet shipped as its own workbook from v3.0 until the Beta
`Full_Factorial` rework retired the Data Tables calculation-mode conflict that
motivated the split. The split shipped with v3.0 as the Univariate artifact's
1.0.0 initial release; on the Regression side it was bundled into the
non-breaking 3.0.0. The grid shrink was MAJOR for the Univariate workbook alone.
The two workbooks have since been reunified. The Version History sheet now
carries one lineage — the Regression one, which always held the pre-split
Univariate history — and the two standalone Univariate releases described below
are recorded here, in this section, rather than in a second changelog no build
could publish.

**The split** moved Univariate Analysis into its own workbook. Both artifacts
carried the complete function library — there was no bundling, no dependency
closure, and no per-artifact function subsetting; they differed only in which
sheets they contained. It was **non-breaking for both**.

Splitting let each artifact set its own calculation mode, so the Regression
workbook ran in full Automatic. **Shipped:** `build_production.py` emitted the
Regression artifact and `build_univariate.py` emitted the Univariate artifact
(shared scaffolding in `lambda_catalog/build_common.py`); the verifier carried a
`skip_regression` mode for the Univariate-only workbook. Both scripts were
merged back into the single `build_production.py` when the workbooks were
reunified.

**The grid shrink** follows as a separate release of the Univariate artifact, and
is **MAJOR for that workbook's version only**. Weibull and Gamma collapse to
one-dimensional searches by profiling out the scale/rate parameter in closed form,
and use a profile-NLL line chart — an upgrade in legibility: the basin, the
interior minimum, and any boundary hit are more visible in a line chart than
across a 2-D grid. Profiling is still genuine MLE — the profile maximizer is the
joint maximizer — so this extends the v1.1 MLE-via-grid reframing rather than
replacing it. The planned Beta half — a method-of-moments start and a smaller
fixed grid to take the total to ~370 — was never implemented and is superseded:
the PR #203 rework moved Beta onto a live `Full_Factorial` spill with an editable
`Grid Points` cell, retiring the Data-Table-recalc-cost driver that motivated a
fixed shrink. See [DECISIONS.md § the grid shrink](DECISIONS.md#the-grid-shrink-ships-as-a-later-release-of-the-univariate-artifact).

It is MAJOR because the Scale Min/Max/Step input cells change or disappear, so a
user's saved bounds stop meaning anything. That is a workbook-interface break, and
it does not move the Regression workbook version.

**Status — the Weibull and Gamma half SHIPPED as Univariate 2.0.0.** Those four
stage blocks are 1-D profile searches with closed-form starting values and
profile-NLL line charts. **The Beta half is superseded (overcome by events)** as
noted above — the method-of-moments start and fixed smaller grid were a
Data-Table-era recalc-cost optimization, retired by the PR #203 rework onto a
live `Full_Factorial` spill with an editable `Grid Points` cell.

Design rationale: [DECISIONS.md § v3.0](DECISIONS.md#v30--two-artifacts-a-bounded-model-context-and-the-constructor-pipeline).

---

## A note on the "v3.0" label in the codebase

The Specification-Driven Regression changeover was planned as v3.0 and renumbered
to **v2.0** before release, under the interface definition above. The old label
survives in comments and docstrings across `write_spec_block.py`,
`analyze_regression_spec_block.py`, `build_production.py`, and three test modules,
where "v3.0" means the spec-block changeover.

**v3.0 now means the engine-interface release.** The two are unrelated, and the
collision is live. `write_spec_block.py`'s docstring is corrected,
and the human test plan that carried the old label in its filename has been
retired; the remaining comment references are tracked as a cleanup item in
[TODOs.md](TODOs.md). They are comments only — no executable logic reads the
label.

---

## Analysis ToolPak Parity Reference

The ToolPak ships 19 tools. Tracking which are covered, planned, or intentionally skipped.

**Covered or exceeded (v1):** Regression (with diagnostics, influence measures,
cross-validation, information criteria, and prediction), Correlation, partial descriptive
stats; Descriptive Statistics + Histogram (v1.1). **v2.0 exceeded further:** categorical
predictors via the declarative spec, which the ToolPak has never offered; fixed-effects
panel regression follows at v2.1.

**Planned:** Rank/Percentile; t-tests, F-test,
Covariance (v3.10 two-sample); one-way ANOVA (unordered, v3.13+); Moving Average +
Exponential Smoothing (v3.12 Time Series sheet).

**Planned to exceed — the Time Series sheet (v3.12).** The ToolPak's entire
time-series offering is Moving Average and Exponential Smoothing. It has no ACF,
no PACF, and no stationarity or white-noise test of any kind, so ACF/PACF with
significance bands, Ljung–Box, ADF and KPSS are not parity items — there is
nothing to reach. They are listed here because a reader scanning this table for
"what does the library do that the ToolPak cannot" should find them.

**Intentionally skipped:**

- **z-Test (two-sample for means)** — assumes known population variance; rarely applicable.
- **Fourier Analysis** — engineering/signal domain; out of scope (may add later).
- **Two-factor ANOVA** — complexity vs. demand. *(Note: post-v2.0, two-way fixed effects
  via the spec covers adjacent territory; revisit whether this skip still holds.)*
- **Random Number Generation / Sampling** — largely redundant with native Excel functions.

**Why the library exists (ToolPak flaws it fixes):** ToolPak output is static (pasted
values that never update when inputs change), opaque (no formula trace), one-sheet-at-a-time
with manual reruns, locked behind a modal dialog, and diagnostically dated (no VIF, Cook's
distance, leverage, studentized residuals, PRESS, AIC/BIC, or cross-validation). The Lambda
Library is live, transparent, auditable, reusable, and diagnostically modern.
