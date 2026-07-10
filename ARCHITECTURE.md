# Architecture

Foundational patterns that don't change version-to-version. The version plan
lives in [ROADMAP.md](ROADMAP.md); resolved design decisions with their
rationale live in [DECISIONS.md](DECISIONS.md); this file is the
between-versions reference — the rules and patterns a new feature must honor
even if the version it ships in changes.

**Reading order for someone new to the codebase.** §1 (Naming) and §2
(Categories) first — they apply to every catalog function. §3 (Role / Type /
Sequence) and §4 (Spec block) before touching the Regression sheet. §5 (Data
Transformation) before writing transforms. §6 (Chart patterns) and §7
(Reserved-spec-column pattern) when relevant.

---

## 1. Naming Convention

Canonical function names use **Title_Case_With_Underscores**, fully spelled
out — no abbreviations (e.g. `Absorb_Two_Way_Fixed_Effects`, not
`ABSORB_2FE`). This is a deliberate departure from Excel's all-caps native
functions: a mixed-case name in a formula bar is immediately recognizable as
library code, not a built-in. A reviewer should be able to scan a nested
formula and tell at a glance which calls are library functions versus native
Excel (`SUM`, `FILTER`, `XLOOKUP`) without cross-referencing a function list.

Rules:

- Full English words only. Spell out what an abbreviation would have stood for
  (`Two_Way`, not `2WAY`; `Fixed_Effects`, not `FE`).
- Numerals appear only when the numeral is itself the statistical quantity
  (e.g. a literal lag of `2`, not a stand-in for "two-way").
- Underscores separate words; no camelCase.
- One canonical name, one LAMBDA, one place it can be wrong.

### Retained initialisms and abbreviations

Words in their own right, never expanded. These are the explicit sign-offs
from the v2.0 rename pass (2026-07-03), recorded as exceptions to the
"no abbreviations" rule:

- `AIC`, `AICc`, `BIC` — information-criterion names
- `VIF` — variance inflation factor
- `PRESS` — prediction sum of squares
- `CDF`, `NLL` — cumulative distribution function, negative log-likelihood
- `LOOCV` — leave-one-out cross-validation
- `PERT` — Beta / PERT
- `R` — Pearson / Spearman correlation
- `QQ` — as in `QQ_Correlation`
- `GoF` — the `GoF_*` goodness-of-fit family
- `MS`, `SS`, `SE` — the classical ANOVA-table shorthands (`MS_Residual`,
  `SS_Total`, `SE_Regression`, `SE_Coefficients`, etc.)

Other sign-offs from the same pass:

- `CDF_BetaPERT` / `NLL_BetaPERT` — kept (no underscore inserted)
- `P_Value_F` → `F_Statistic_P_Value` (renamed to pair with `F_Stat` →
  `F_Statistic`)
- `Grid_Argmin` → `Grid_Argument_Minimum` (spelled out per the convention)

### Alias layer (future, optional)

A separate, optional layer of short, ALL-CAPS aliases may be added in a
later pass for power-user typing speed (e.g. `ABSORB2FE` as an alias for
`Absorb_Two_Way_Fixed_Effects`). Aliases are thin wrappers — each alias
LAMBDA's entire body is a call to the canonical function, with no
independent logic:

```excel
ABSORB2FE = LAMBDA(x, group1, group2, [include], [passes],
    Absorb_Two_Way_Fixed_Effects(x, group1, group2, include, passes)
)
```

This keeps a single source of truth: if the canonical implementation
changes, every alias inherits the fix automatically. Aliases are never the
documented or taught form — they exist purely as optional shortcuts and
should be introduced only after the canonical library is stable, to avoid
maintaining two names for a function that's still under active revision.

See [DECISIONS.md § Aliases](DECISIONS.md#aliases) for the full alias table
and the deferral record.

### Naming-style departures

Recorded exceptions to the "one canonical name, one LAMBDA" rule. Pattern:
when a family is closed-form (one shape per name), use the per-shape style;
when a family is combinatorial in its inputs, use a dispatcher. Future
combinatorially-named families follow the same exception.

- v2.2 — the `Unit_Space_*` dispatcher family (`Unit_Space_R_Squared`,
  `Unit_Space_Adjusted_R_Squared`, `Unit_Space_RMSE`). See
  [DECISIONS.md § v2.2 unit-space dispatcher](DECISIONS.md#v22--transforms--unit-space-comparability).

---

## 2. Function Categories and Subcategories

Every catalog entry carries a `category` and a `subcategory`, used to drive
filtering on the `LAMBDA_functions` sheet. This taxonomy is **purely
functional — it does not encode version.** Version is a property of the
library's release history (tracked in the changelog and Version History
sheet); category is a property of what the function *does*, and a function's
category should not change just because it shipped in a later release.

Subcategories are scoped *within* a category — each category defines its own
subcategory list rather than sharing one flat list across categories, so a
category can grow its own subdivisions independently as it fills up.

| Category | Subcategories |
|---|---|
| **Model Construction** | MLR Core · Coefficient Inference · Prediction · Specification & Design Matrix · Transforms & Unit-Space |
| **Diagnostics** | Residual · Influence & Leverage · Multicollinearity · Cross-Validation · Information Criteria |
| **Data Transformation** | Sample Construction & Diagnostics · Location & Scale · Group & Panel · Categorical & Model Construction · Longitudinal & Panel-Time |
| **Distribution Fitting** | Descriptive · Histogram Binning · Parameter Estimation · Goodness-of-Fit |
| **Resampling & Simulation** | Bootstrap · Monte Carlo |
| **Model Comparison** | Spec String & Registry *(provisional — promote from a subcategory only once it holds 2+ functions; see v2.3)* |

Notes:

- **Specification & Design Matrix** is under Model Construction to house the
  v2.0 constructor functions (`x_s`, `Constructed_Column_Names`, etc.).
- **Transforms & Unit-Space** is under Model Construction for the v2.2
  unit-space fit statistics.
- Data Transformation subcategories are unchanged across versions; those
  functions serve double duty as constructor internals and standalone
  user-callable transforms.
- **Model Comparison** is listed provisionally as a top-level category for
  the v2.3 sheet's single function. Per the "categories describe what a
  function does" rule it earns top-level status once it holds 2+ functions;
  until then it may equally live as a Model Construction subcategory.

This table is the source of truth for the controlled vocabulary; `category`
and `subcategory` values in `lambda_functions.json` should be drawn only
from this list, not invented ad hoc per function.

---

## 3. Variable Role / Predictor Type / Sequence

The model specification is organized along three orthogonal axes. The axes
are deliberately separate — a column's value on one axis says nothing about
its value on another.

| Axis | Values | Future values | Meaning |
|---|---|---|---|
| **Variable Role** | Response · Predictor · Identifier · Filter · Omit | Fixed Effects, Weight, Time (v2.1+) | What the column *is* in the model |
| **Predictor Type** | Continuous · Categorical | *(closed — never grows)* | How a Predictor *enters* the design matrix — meaningful only when Role = Predictor |
| **Sequence** *(structural, post-v2.0)* | TRUE · blank | *(flag — never grows)* | Which column *orders* the data, for lag/difference/serial-correlation features |

Literature precedent: this is essentially the tidymodels `recipes` role
system (outcome / predictor / ID) plus Stata's factor-variable notation on
the Type axis.

**The Sequence structural axis is deliberately NOT a Role value and NOT a
Predictor Type.** It annotates *structure*, orthogonal to what a column is
in the model: a column can be Role = Predictor, Type = Continuous AND
Sequence = TRUE simultaneously (the escalation-index case), and an
Identifier like Year is a typical sequence axis. No constructor formula
reads the flag for design-matrix assembly — it exists so the
lag/difference/serial-correlation features land on a declared axis.

### Cardinality rules

- **Exactly one** Response. Status-block error at zero or two-plus.
- **At most one** Sequence (zero-or-one). Zero flags is a valid spec
  (non-panel data); two-plus is a visible status-line error (same pattern as
  exactly-one-Response, with a >1 threshold).
- **At most one** of each future Role value that is single-instance by
  nature: Fixed Effects, Weight, Time. The status block validates each the
  same way it validates exactly-one-Response.
- **Zero-or-more** Filter columns. AND-composed: no Filter columns → all rows
  included; multiple → logical AND (declarative stratification, e.g.
  `Full_Data` AND a hand-built "Developing only" flag).
- **No cardinality constraint** on Identifier or Omit.

### Role semantics

- **Response** — the dependent variable. Derived, not hard-wired: `y`
  becomes "the column whose Role is Response" (`XMATCH` over the Role
  range). Swapping the dependent variable is a dropdown change — cost vs.
  log-cost vs. cost-per-unit against the same table is a workflow feature,
  not plumbing.
- **Predictor** — candidate model input; the only Role for which Include,
  Type, Reference, and Levels are meaningful (cascading relevance, see §4).
- **Identifier** — labeling columns (Country, Year): available to the
  residual-output zone for row labeling; never enter the design matrix.
- **Filter** — supplies the row mask (zero-or-more, ANDed). Filter columns
  hold boolean / 0-1 values; blanks are FALSE.
- **Omit** — dataset semantics, set once: "never a candidate for anything"
  (helper columns, notes). Distinct from Include, which is *model-iteration
  state* flipped constantly ("candidate, currently out of this run").
  Collapsing them would destroy the record of what a column *is* every time
  a toggle experiment runs. Cost of keeping both: one dropdown value.
- **Fixed Effects** *(v2.1)* — the panel role: enters no column; the entire
  design matrix and the Response are demeaned by its groups (one FE
  variable → `Demean_By`; two-plus → visible error — two-way absorption via
  `Absorb_Two_Way_Fixed_Effects` is its own post-v2.1 milestone, see v2.7+).
- **Weight** *(v2.6+)* — WLS. See DECISIONS § v2.6.
- **Time** *(v2.7+)* — time-index designation. Partially forward-wired via
  the Sequence axis but the full semantics still need design work.
- **Cluster** *(v2.7+)* — clustered-robust variance estimator. Forward
  wiring is partial (the dormant branch in `Serial_Correlation_Group()`'s
  SWITCH).

### Predictor Type semantics

- **Continuous** — enters the design matrix as a single column.
- **Categorical** — encoded via the level-vector split: training and
  prediction both call the same encoder with the same training level vector
  (see [DECISIONS.md § v2.0 level-vector split](DECISIONS.md#v20--specification-driven-regression)).
Closed: the Predictor Type axis never grows. Fixed Effects is not a Type
value; it contributes no columns and no coefficients. It moves to the Role
axis (v2.1), and the Type axis becomes permanently Continuous/Categorical.

---

## 4. The Model Spec block (A–L)

The spec spans **every column of the source table**, one row per column:

| Col | Contents | UX |
|---|---|---|
| A | Variable name | Spills from the table's header row |
| B | **Role** | Dropdown: `Response` · `Predictor` · `Identifier` · `Filter` · `Omit`. Pre-filled by the build — no blank cells. Role precedes Include because it is the larger declaration. |
| C | Include toggle | Orange input; meaningful only when Role = Predictor |
| D | **Predictor Type** | Dropdown: `Continuous` · `Categorical`; meaningful only when Role = Predictor; pre-filled `Continuous` |
| E | **Reference Level** | Orange input, meaningful only for Categorical Predictors. Blank = **first level in sort order** (confirmed default, matching R). CF: red when the entered level does not exist in the analysis sample. |
| F | **Order** *(reserved, not implemented v2.0)* | Input, integer. Will control user-specified ordering of Identifier columns in the row-label text-join; v2.0 always joins in table order. Present now so the layout absorbs the feature without a future column insertion. Cell comment marks it reserved; no validation yet (no fixed domain). |
| G | **Transform** *(reserved, not implemented v2.0)* | Input, dropdown. Will apply a transform (e.g. `Log`) to a Continuous Response or Predictor. v2.0 dropdown list is `None` only; build pre-fills `None`. Cell comment marks it reserved. |
| H | **Sequence** *(structural axis, post-v2.0)* | Orange input flag, dropdown `TRUE`/blank. The shipped default pre-flags **Year** `TRUE` (the WHO panel's ordering axis; every other row blank) so the Sequence machinery is live at T0; on a non-panel dataset leave it blank. Marks **at most one** variable as the ordering axis. Status line at H2: red error at two-plus flags (zero is valid); per-cell red CF points at the offending rows. Read by the validation layer, by the sequence-spacing layer (`Sequence_Deltas`, `Base_Period_Delta`) since the base-period release, and — since the DW-gate release — by the serial-correlation accessor `Sequence_Column` (which feeds the gated `Durbin_Watson_By` diagnostic cell). No design-matrix constructor consumes it: Sequence orders the data, it never enters the model matrix. |
| I | **Sequence Period** *(typed override input, post-v2.1 Sequence fix)* | Orange input — the user types a number on the Sequence-flagged row to declare a Δ that differs from the computed candidate. Blank by default; the spec falls back to the candidate. Read only by the in-use display at column J, not by any constructor. The cell is the load-bearing override point of the reference-level pattern. |
| J | **Period In Use** *(live — base-period release; Sequence companion)* | **Computed-with-override display**, the reference-level pattern: shows the typed value at I if non-blank, otherwise the candidate closure's value (`Base_Period_Delta_Candidate()` — MODE of within-group consecutive spacings, MIN fallback when no spacing repeats). Read only by the base-period layer — the `Base_Period_Delta()` accessor (the omitted-`[delta]` default of `Lag_By`/`Difference_By`) and the Sequence Spacing block's Δ-in-use display. Override flagging lives on the Sequence Spacing block (verdict lines on rows 31–34), not on the spec block itself: the J cell stays plain so the spec reads top-to-bottom as a clean declaration. The block (rows 28–34 under the spec) also shows the delta spectrum and verdicts: Regularity (any spacing ≠ Δ), Off-grid (spacing not a whole multiple of Δ), the no-natural-base-period override prompt, and calendar-signature guidance (~28–31/90–92/365–366 clusters → recommend an integer period index upstream; day counts are never quantized to a scalar Δ). |
| K | **Levels** | **Computed display**: distinct level count over the mask-included rows, shown only for Categorical Predictors. Live against stratification. CF: **red when L ≤ 1 while included** (contributes L−1 = 0 columns). Large L needs no flag — the visible count is the warning. |
| L | **Reference In Use** | **Computed display**: the reference level the constructor will actually drop, surfaced even when defaulted. The Σ(design columns) = COLUMNS(x_s()) audit lives in the status strip's `k` cell, and the gap column right of the spec block still visually reserves a future Design Columns slot. |

### Reserved-column policy (F, G)

Neither reserved column is read by any formula — confirmed by construction,
not by convention: `X_s()`, `Constructed_Column_Names()`, `Row_Labels()`,
and `Sample_Include()` must not reference `Spec_Order` or `Spec_Transform`
(and may not reference `Spec_Sequence` or `Spec_Sequence_Period` either —
those names are consumed only by the zero-or-one validation and the
base-period layer, never by a constructor). The columns exist purely so
the *sheet layout* absorbs the future feature now; wiring them in a later
release is additive (a formula change), not a second column-insertion
breaking the sheet a second time — exactly how column I went live in the
base-period release.

### Cascading relevance

C–G and K–L gray out (conditional formatting) whenever Role ≠ Predictor —
the same pattern as Reference-only-for-Categorical, applied one level up.
H–J key on the **Sequence flag itself**, not on Role: they gray on every
row that is not the sequence axis, because Sequence is structural and
Role-independent.

### Display derives, never feeds

Columns J (the Period In Use display), K, and L must not be inputs to the
constructor. The J cell calls `Base_Period_Delta_Candidate()` and reads
column I; the K and L cells call the same mask-aware primitive
(`Dummy_Levels`); the constructor calls `Base_Period_Delta()` (which reads
J) and the same primitive. Display and constructor read the same closure,
so they are provably consistent. Letting the engine read a display column
would make it load-bearing. One source of truth is the *function*.

This is also what settled the `Dummy_Code` design: the level-vector split
is required — it makes display and constructor provably consistent, and
gives prediction the training level set.

### The effective row mask

Composes two layers:

1. *Automatic role-aware completeness* — numeric required for Continuous
   Predictors and the Response; non-blank for Categorical Predictors;
   ignored for Identifier/Filter/Omit.
2. *Declared Filter columns* — zero-or-more, ANDed.

This settles the architecture of the former role-aware-completeness open
item by construction; only the auto-completeness LAMBDA remains to build,
and the hard-coded `Data_Completeness(...[Life expectancy]:[Schooling])`
span dies with it. The `Sample_Include()` shipped-with role-aware
completeness layer is the resolution — see
[DECISIONS.md § v2.0 auto-completeness](DECISIONS.md#v20).

### The `x_s()` row-mask contract

The constructor reads the effective mask *only* to fix Categorical level
sets; it always emits **full-height** columns and leaves row filtering to
the engine functions, exactly as v1 does. Filtering rows inside `x_s()`
would double-filter against the engine's own mask application.

### Spec-validation semantics

Failure is signaled by `NA()` in the affected output cell, not by a silent
fallback or a descriptive string. This is the `NA()`-based error-signaling
convention used by `Dummy_Levels` / `Dummy_Code`; it lets every downstream
`IFERROR`/`ISNA` guard work without special-casing. The status block
aggregates these signals into a single visible Error State cell.

### Reserved-spec-column pattern (general)

The F/G "reserved, not implemented" pattern is a general technique:
introduce a column in the sheet layout, ship it as a placeholder that no
formula reads, and absorb the future feature additively (formula change)
instead of with a second column-insertion that breaks the sheet a second
time.

The function-side equivalent: a SWITCH or IF with a dormant branch that
returns a `RESERVED — vN+` token. The v2.6+ `Cluster` branch in
`Serial_Correlation_Group()`'s SWITCH is the worked example — supplying
the grouping key from a Cluster role for pooled-panel diagnostics without
absorption is a resolver-only edit, no engine change.

---

## 5. Data Transformation taxonomy

Cross-cutting infrastructure as a *taxonomy*: these functions serve **double
duty** — internals of the spec-driven constructor, and standalone
user-callable transforms for free-form work on the data sheet. Tracked as
its own catalog category, separate from the version ladder.

**Delivery, however, is pinned to the ladder** (none of these standalone
functions is built yet as of v2.0): the user-callable transform library
ships at **v2.2** alongside the column-G wiring, with two exceptions —
`Demean_By` and `Group_Mean` ship at **v2.1** as Fixed-Effects internals,
and the two-way functions (`Absorb_Two_Way_Fixed_Effects`,
`Demean_Two_Way_Balanced`, `Fixed_Effects_Convergence_Check`) follow the
**two-way FE milestone (post-v2.1)**.

### Subcategories

**Sample Construction & Diagnostics**

- `Numeric_Complete_Cases(data)` — listwise-deletion sample mask; `1` when
  every value in a row is numeric.
- `Is_Balanced_Panel(group, time, [include])` — `TRUE` only when every
  included group has exactly one observation for every included time
  period.
- `Fixed_Effects_Convergence_Check(x, group1, group2, [include])` — largest
  absolute remaining group mean across two fixed-effect dimensions; near
  zero indicates convergence after `Absorb_Two_Way_Fixed_Effects`.
  Surfaced in the status block whenever two Fixed Effects variables are
  active *(two-way FE milestone, post-v2.1)*.

**Location & Scale**

- `Center(x, [include])` — grand-mean centering, \(x_i - \bar{x}\).
- `Zscore(x, [include])` — standardization via `STDEV.S`,
  \((x_i - \bar{x}) / s_x\).
- `Minmax_Scale(x, [include])` — scales to \([0, 1]\).
- `Winsorize(x, [lower_p], [upper_p], [include])` — caps values outside
  selected percentiles (default 1st/99th). Remains an explicit modeling
  decision, never an automatic preprocessing step.
- `Ln_Positive(x, [include])` — natural log, restricted to strictly
  positive numeric values; returns `""` rather than a worksheet error for
  zero, negative, or non-numeric input.

**Group & Panel**

- `Group_Mean(x, group, [include])` — matching group mean on every
  included row.
- `Demean_By(x, group, [include])` — one-way within transformation,
  \(x_{ig} - \bar{x}_g\). **v2.1 constructor internal** for a single
  Fixed Effects variable.
- `Zscore_By(x, group, [include])` — within-group standardization.
- `Decompose_By(x, group, [include])` — returns the between-group mean
  and within-group deviation as two columns, exposing
  \(x_{ig} = \bar{x}_g + (x_{ig} - \bar{x}_g)\).
- `Demean_Two_Way_Balanced(x, group1, group2, [include])` — direct
  two-way demeaning, exact only for a balanced panel; check with
  `Is_Balanced_Panel` first. *(Two-way FE milestone, post-v2.1.)*
- `Absorb_Two_Way_Fixed_Effects(x, group1, group2, [include], [passes])` —
  iterative alternating-projection demeaning for unbalanced panels.
  Convergence is not verified internally; always pair with
  `Fixed_Effects_Convergence_Check`. **Constructor internal for two Fixed
  Effects variables at the two-way FE milestone (post-v2.1).**

**Categorical & Model Construction**

- `Dummy_Levels(category, [reference], [include])` — **rebuilt for v2.0**
  (an earlier version existed as a catalog function with string-based
  error returns; dropped and replaced rather than amended — see
  [DECISIONS.md § v2.0 Dummy rebuild](DECISIONS.md#v20--specification-driven-regression)). Signals
  every downstream `IFERROR`/`ISNA` guard works without special-casing.
  Retained categorical levels as a horizontal header row; backs the v2.0
  prediction-input validation lists and is the hard dependency of
  `X_s()`'s level-vector split.
- `Dummy_Code(category, [reference], [include])` — **rebuilt for v2.0**
  alongside `Dummy_Levels`, calling it internally for level determination
  (one source of truth, same NA()-based error contract). Dummy-coded
  matrix. Use a reference level (treatment coding) when the design
  includes an intercept; full one-hot coding plus an intercept causes
  perfect multicollinearity. **Reference-level validation (confirming the
  requested reference actually exists in the included sample) is required
  at implementation, not deferred** — an invalid reference must error
  (`NA()`), never silently fail to drop a column and reintroduce the
  exact collinearity the function exists to prevent. Standalone; `X_s()`
  does not call it directly (it encodes inline via broadcast) but is held
  to the same standard. **v2.0 constructor internal** for Categorical
  roles, via `Dummy_Levels`.
- `Dummy_Column(category, level, [include])` — single indicator column
  per explicit call.
- `Interact(x1, x2)` — elementwise product \(x_1 x_2\); broadcasts across
  dummy-coded matrices to produce one interaction column per retained
  level.
- `Model_Matrix(X, [add_intercept])` — optionally prepends an intercept
  column. Intentionally not variadic — predictors are assembled
  explicitly with `HSTACK` so the specification stays visible and
  auditable.

**Longitudinal & Panel-Time**

Shipped early, in the base-period release (ahead of the v2.2 bundle), to
the gap-aware semantics recorded in
[DECISIONS.md § v2.1 base-period layer](DECISIONS.md#v21):

- `Lag_By(x, group, seq, [delta], [include])` — prior-period value
  within the same group, keyed on `group`/`seq` **by exact time value,
  not physical row order** (exact-match lookup of `(group, seq − Δ)` pairs,
  never OFFSET/row arithmetic). A gap — `seq − Δ` absent within the
  group — returns `NA()`, never the previous available row.
- `Difference_By(x, group, seq, [delta], [include])` — within-group time
  difference \(\Delta x_{it} = x_{it} - x_{i,t-\Delta}\), delegating the
  pair lookup to `Lag_By`. Each group's first period returns `NA()` —
  never a fabricated 0 that would enter a design matrix silently.
- `Base_Period_Delta()` — the Δ in effect: spec column I on the
  Sequence-flagged row (computed candidate or typed override). The
  omitted-`[delta]` default of both functions above — Δ is never a silent
  1; with no declared axis they return `NA()` everywhere (visible
  failure). Companion sheet-scoped closures `Sequence_Deltas` /
  `Base_Period_Delta_Candidate` / `Sequence_Delta_Spectrum` drive the
  spec's Sequence Spacing block.

### The row-aligned `""` convention

Row-aligned transforms accept an optional `include` mask (`1`/`TRUE` keeps
the row, `0`/`FALSE` excludes it). When `include` is omitted, those
functions construct a default mask from the required inputs. Excluded rows
return `""` so spilled arrays stay aligned with the source rows — `""` was
chosen deliberately over `NA()` because it round-trips cleanly through
`ISNUMBER`-based keep logic elsewhere in the library (`ISNUMBER("")` is
`FALSE`, so a downstream mask built on `ISNUMBER` correctly re-excludes
it) without erroring inside `SUM`/`AVERAGE` the way a propagated error
value would.

### The `NA()` exception (recorded)

An *included* row whose difference is incomputable (first period, gap,
non-numeric value) returns `NA()`, not `""` — it is a visible
incomputable observation, not an excluded row. `""` remains the
excluded-row return. Verification: `tests/test_difference_by_verification.py`
(WHO exact counts plus the punched-out-year and calendar-date synthetic
cases); human test plan T17–T19.

---

## 6. Chart patterns and pitfalls

The full chart-creation rules live in [CLAUDE.md](CLAUDE.md) / [AGENTS.md](AGENTS.md)
at the project-instructions tier (xlwings COM, never openpyxl; `.Text` for
static titles, `.Formula` for cell-linked; histogram `GapWidth = 0`;
identity lines as real data series; chart title cells outside the
try/except guard; etc.). ARCHITECTURE does not duplicate those — the
file would be three copies instead of two. The single line of
cross-cutting context that's worth recording here:

- **Why xlwings COM, never openpyxl.** openpyxl's `load_workbook`/`save`
  rewrites the entire .xlsx package and silently drops chart parts, VML
  drawings, and chartUserShapes it didn't create. Loading a workbook that
  already has Excel-created charts and saving it back will destroy those
  charts. The zipfile-patching workaround (generate chart XML in
  throwaway openpyxl workbooks, then splice into the real zip) is fragile,
  hard to test, and was ultimately abandoned. All charts in this project
  use `sheet.api.ChartObjects().Add(...)` via xlwings COM, which writes
  directly into the live Excel instance without round-tripping the file.

For the implementation details (chart-creation pattern, histogram
formatting, chart positioning, identity-line construction, build-phase
retry separation), see [CLAUDE.md § Charts](CLAUDE.md).

---

## 7. Reserved-spec-column pattern (general)

See §4 "Reserved-column policy (F, G)" for the sheet-side form. The
function-side form is the same idea applied to a LAMBDA's argument list
or its internal `SWITCH`:

- A `SWITCH` argument that is a Role-axis value can carry a dormant branch
  for a not-yet-implemented Role, returning a `RESERVED — vN+` token. The
  resolver (`Serial_Correlation_Group()`) does this for the `Cluster`
  role — supplying the grouping key from a Cluster role for pooled-panel
  diagnostics without absorption is a resolver-only edit, no engine
  change. The v2.6+ `Cluster` work lights up the dormant branch by adding
  the engine-side estimator.

The general principle: when a feature lands across multiple versions, the
sheet layout and the function signature can each carry reserved slots
that absorb the future feature additively, instead of forcing a second
restructure. The cost is one dormant cell or branch; the benefit is a
single breaking change per version axis rather than per feature.
