# TODOs

## Design note — chart series data ranges

Chart `SERIES` formulas require explicit range references; the `#` spill operator is not reliably supported in chart series formulas. However, referencing all 1,048,576 rows can significantly degrade performance or crash Excel when the populated dataset is much smaller.

Instead, define dynamically sized named ranges using the row count in `$T$8`. For example:

```excel
RegChartQQX = OFFSET($AN$2,1,0,MAX(IFERROR($T$8,1),1),1)
RegChartQQY = OFFSET($AO$2,1,0,MAX(IFERROR($T$8,1),1),1)
```

These formulas define ranges beginning at `AN3` and `AO3`, respectively, and extending for exactly the number of rows specified in `$T$8`. All chart series names use the `RegChart` prefix and are documented in CONTRIBUTING.md.

---

## Alias layer

Design/planning item — see ROADMAP.md for the architectural rationale. Aliases are thin ALL-CAPS wrappers whose entire body is a single call to the canonical function; the canonical function remains the single source of truth. Implement only after the canonical library is stable.

Suggested alias names:

**Regression — scalar outputs**

| Alias | Canonical |
|---|---|
| `R2` | `R_Squared` |
| `ADJ_R2` | `Adjusted_R_Squared` |
| `MULT_R` | `Multiple_R` |
| `SE_REG` | `SE_Regression` |
| `DW` | `Durbin_Watson` |
| `QQ_CORR` | `QQ_Correlation` |
| `DFR` | `Regression_Degrees_Of_Freedom` |
| `DFE` | `Residual_Degrees_Of_Freedom` |
| `DFT` | `Total_Degrees_Of_Freedom` |
| `SSR` | `SS_Regression` |
| `SSE` | `SS_Residual` |
| `SST` | `SS_Total` |
| `MSR` | `MS_Regression` |
| `MSE` | `MS_Residual` |
| `P_VAL_F` | `F_Statistic_P_Value` |

**Regression — coefficient vectors**

| Alias | Canonical |
|---|---|
| `COEF` | `Coefficients` |
| `SE_COEF` | `SE_Coefficients` |
| `TSTAT` | `T_Statistics` |
| `PVALS` | `P_Values` |
| `CI_LOW` | `Confidence_Interval_Lower` |
| `CI_UP` | `Confidence_Interval_Upper` |
| `P_R2` | `Partial_R_Squared` |
| `P_COR` | `Partial_Correlation` |
| `BETA_W` | `Beta_Weights` |

**Regression — observation vectors**

| Alias | Canonical |
|---|---|
| `PRED` | `Predictions` |
| `RESID` | `Residuals` |
| `STDR` | `Studentized_Residuals` |
| `LEV` | `Hat_Diagonal` |
| `COOK_D` | `Cooks_Distance` |
| `LOOCV` | `LOOCV_Prediction` |
| `PI` | `Prediction_Interval` |

**Regression — utilities**

| Alias | Canonical |
|---|---|
| `COMPLETE` | `Complete_Cases_Filter` |
| `CORMAT` | `Correlation_Matrix` |
| `DESIGN` | `Design_Matrix` |

**Univariate — descriptive**

| Alias | Canonical |
|---|---|
| `DSTAT` | `Descriptive_Statistics` |
| `NMISS` | `Missing_Count` |

**Univariate — histogram binning**

| Alias | Canonical |
|---|---|
| `NBINS` | `Number_Of_Histogram_Bins` |
| `EDGES` | `Bin_Edges` |
| `UEDGES` | `Upper_Bin_Edges` |
| `LEDGES` | `Bin_Lower_Edges` |
| `BIN_MIDS` | `Bin_Midpoints` |
| `BIN_FREQS` | `Bin_Counts` |

**Univariate — goodness-of-fit**

| Alias | Canonical |
|---|---|
| `GOF_AD` | `GoF_Anderson_Darling` |
| `GOF_KS` | `GoF_Kolmogorov_Smirnov` |

**Grid-search helpers**

| Alias | Canonical |
|---|---|
| `GS_MIN` | `Grid_Argument_Minimum` |
| `GS_OPT` | `Grid_Search_Optimum` |

---

## v1.x — Regression sheet

- TODO: Add reference lines to the Cook's Distance, PRESS Residuals, and Leverage vs. Studentized charts. Format thematically similar to the conditional formatting in the table (yellow = mild, red = strong). Use minimalist helper columns (2 anchor points using max/min of the relevant threshold; place them beneath the chart area).

---

## v1.1 — Univariate (shipped; leftovers)

### LAMBDA functions

**PDF functions — DROPPED (unnecessary, will not be implemented)**
- ~~Implement `PDF_Normal` … `PDF_BetaPERT` evaluated at bin midpoints~~ — superseded.
  The histogram tables already compute per-bin probabilities as **CDF deltas between the
  bin boundaries** (`CDF(upper edge) − CDF(lower edge)`, the 8 `CDF_*` probability
  columns in each histogram block). That delta is the bin's probability mass (the PDF
  integrated exactly over the bin) — more faithful to the histogram than a midpoint PDF
  evaluation — so no PDF LAMBDAs are needed. See ROADMAP.md § v1.1 Distribution
  fitting for the full rationale.

### Sheet writer (`write_sheet_univariate.py`)

**Q-Q plots and histogram overlays — DONE (ships with the next workbook build)**
- ~~Per-distribution Q-Q plots (8 charts) using OFFSET-based named ranges~~ — done.
  Zone 6 (cols CW–DF) holds Hazen plotting positions `(i−0.5)/n` (the same
  convention as `QQ_Correlation`/`Normal_Scores`), the sorted sample, and eight
  theoretical-quantile columns (native `NORM.INV`/`LOGNORM.INV`/`GAMMA.INV`/
  `BETA.INV` plus closed-form inverses for Exponential, Weibull, and Triangular,
  validated against scipy in `tests/test_univariate.py`). Eight XY scatter charts
  stack under the histogram charts at G74–G233, each with an identity-line data
  series, fed by OFFSET-based `UV_QQ_*` named ranges.
- ~~Histogram overlays as combo charts~~ — done. Each histogram chart keeps its
  gapless count bars and adds one smoothed, markerless line series per
  distribution. The axis question was settled as **expected counts on the shared
  count axis** (not a secondary axis): `UV_<method>_<Dist>_Expected` named
  formulas multiply the CDF-delta column by the Count stat cell ($E$14).
- TODO: Investigate suppressing worst-fit / N/A-error distributions from the combo
  charts. Best outcome would be dynamically hiding those columns — hidden columns
  drop out of charts automatically (`PlotVisibleOnly` default) — but it is unclear
  whether column hiding can be driven from cell values without VBA (manual hiding
  works; data-driven hiding may be VBA-only, which the library forbids). No-VBA
  fallback to evaluate: emit `NA()` across a suppressed distribution's column, since
  line charts skip `#N/A` points — same chart effect without hiding.

**Additional distributions (long-term)**
- TODO: Add support for more distribution families: Bernoulli, Binomial, Geometric, Negative Binomial, Hypergeometric, Poisson, Uniform, Chi-Square, Student-t.

---

## v2.0 — Specification-Driven Regression (shipped; leftovers)

Human test plan fully executed and signed off PASS 2026-07-05 (T0–T16). One open
decision remains from it:

- TODO: Resolve the blank-categorical caveat — `Sample_Include()`'s role-aware
  completeness layer requires numeric Response and numeric included Continuous
  Predictors, but Categorical Predictors impose no non-blank condition; a blank
  category value encodes as all-zero dummies (indistinguishable from the reference
  level). Run the caveat verification step in `HUMAN_TEST_PLAN_v3_model_construction.md`
  and record the decision: accept as documented behavior, or extend `Sample_Include()`
  with a non-blank condition for included Categorical Predictors. Interim workaround:
  a completeness column declared as a Filter.

---

## v2.1 — Sequence, gap-aware longitudinal, serial-correlation diagnostics, fixed effects (in progress)

Two-way FE is deliberately deferred until this framework is finished — see the
v2.5+ section. Items below are in the locked ship order: the Sequence fix is
the prerequisite for the 2.1.0 entry; the FE Role dropdown, the CI+PI
prediction layout, and the engine are gated to ship as a single release so
users never see "FE is in the dropdown but the engine is forthcoming."

### Shipped (since v2.0.0)

- ~~TODO: Sequence structural axis (spec column H) and reserved Base Period Δ (column I).~~ **DONE (PR #101):** spec block grew A–I to A–K; zero-or-one validation at H2; pre-filled blank by the build; read only by the validation layer and the base-period layer.
- ~~TODO: Gap-aware `Difference_By` / `Lag_By` and the Base Period Δ layer.~~ **DONE (PR #102):** `Difference_By` / `Lag_By` keyed on exact `(group, seq−Δ)` pairs (no OFFSET/row arithmetic); NA() at first periods and panel gaps; the Sequence Spacing block (rows 28–34) and the parser's zero-parameter LAMBDA support. Verified by `tests/test_difference_by_verification.py` (T17–T19).
- ~~TODO: Sequence-aware Durbin-Watson.~~ **DONE (PR #103):** `Durbin_Watson_By(X_s, Y, seq, [Allow_Intercept], [Include])` — sequence-axis DW, row-order invariant; gated cell at Regression X11/Y11 with `n/a — requires Sequence` / `n/a — multiple Sequence flags` tokens.
- ~~TODO: BFN panel Durbin-Watson.~~ **DONE (PR #105):** `BFN_Panel_Durbin_Watson(X_s, Y, group, seq, [delta], …)` (Bhargava–Franzini–Narendranathan 1982); within-group differencing via `Difference_By`; mutual-gate trigger matrix with the DW cell; X12/Y12 cell.
- ~~TODO: Grouping-key resolver.~~ **DONE (PR #106):** `Serial_Correlation_Group()` SWITCH with the dormant `Cluster` branch (the reserved-spec-column pattern for the v2.6+ Cluster role).
- ~~TODO: v1.1 leftovers — histogram distribution overlays and per-distribution Q-Q plots.~~ **DONE (PRs #96, #97, #99, #100):** eight per-distribution Q-Q charts with closed-form quantile inverses; histogram combo charts with overlay lines on the shared count axis; BetaPERT singularity fix; 4×2 Q-Q layout.
- ~~TODO: `--skip-univariate` CLI option.~~ **DONE (PR #98):** build-ergonomics for the slow grid-search step.
- ~~TODO: Spec-driven QC refactor.~~ **DONE (PR #103):** `analyze_regression_spec.py` and `test_regression_spec_qc.py` replace the legacy MLR smoke-test-sheet path; `Full_Data` ships as Omit in the default spec; Year flagged as Sequence, Population as Omit demonstrator.
- ~~TODO: Durbin-Watson under FE — relabel, caveat, or suppress.~~ **DONE (BFN + resolver releases):** resolved as "second cell + mutual gating" (PR #105), with the resolver (PR #106) carrying the dormant Cluster forward-wiring.

### Pending (in ship order; #1 prerequisite, #2/#3/#9 gated to #5)

- TODO: **Sequence axis auto-detection and override** (renames column I to **`Sequence Period`** (the typed override input), adds column J **`Period In Use`** following the Reference Level / Reference In Use pattern (displays the typed override if non-blank, otherwise the candidate)). The current override mechanic has a spill-collision risk for source tables wider than the shipped WHO sample: the spec block reads its own H/I cells, and a longer table could let the override spill overrun an input band. Fix: relocate the override spill and bound every read of the H/I/J band by `COLUMNS(Source_Data)` (the spill-placement principle from `CLAUDE.md`). **Override flagging lives ONLY on the Sequence Spacing block's verdict lines (rows 31–34) — the J spec-block cell stays plain, so the spec reads top-to-bottom as a clean declaration.** Update the Sequence Spacing block (rows 28–34), the spec layout constants, the named-range rename (`Spec_Base_Period_Delta` → `Spec_Sequence_Period`), and the QC analyzers. **Significant testing. Resolve before writing the 2.1.0 Version History entry.**

- TODO: **FE Role dropdown + status-block validation** (gated to ship with the engine). `Fixed Effects` in the Role axis; status-block cells for "active FE variable," "group count" (`n/a — engine forthcoming` until the engine lands), "absorbed df" (`n/a — engine forthcoming` until the engine lands); visible error at 2+ FE variables; intercept × FE red flag; CF bands update. Sheet-only; no engine change.

- TODO: **Surface BOTH intervals in adjacent cells of the prediction outputs section** (gated to ship with the engine). Three lines: point · CI low/high · PI low/high. The PI half-width is `√(σ²·(1+1/T) + q)` and the CI half-width is `√(σ²/T + q)`; same center, same x_new, same t-critical. Sheet layout only — the math lives in the engine. The FE engine drops in group-keyed inputs at the activation step without restructuring this layout.

- TODO: **`Demean_By(x, group, [include])` and `Group_Mean(x, group, [include])`** (constructor internals, also user-callable transforms).

- TODO: **`Is_Balanced_Panel(group, time, [include])`** — one-way/panel diagnostic; ships with `Demean_By` (shares the "valid group set" primitive).

- TODO: **`Absorbed_Degrees_Of_Freedom(spec)`** — Σ(Gᵢ − 1) from the spec.

- TODO: **`y_s()`** — demeaned-Response constructor (new function, not a replacement wired into existing no-FE call sites).

- TODO: **`[DF_Absorbed]` argument (default 0) threaded through df / MS-residual / t-critical.** **Significant testing** — assert bit-equality of every existing engine test with and without the argument, plus FE-active cases for the full inferential chain (SE, t, p, CIs, AIC/BIC).

- TODO: **FE group selection dropdown + ȳᵢ / x̄ᵢ / Tᵢ cells** (gated to ship with the engine). AVERAGEIFS/COUNTIFS respecting the Include/Filter mask; the prediction outputs section from the CI+PI layout activates the group-mean form (ȳᵢ + (x_new − x̄ᵢ)′β̂) with group-keyed inputs; BFN cell (X12/Y12) flips from `n/a — no fixed effects` to active when FE is set. **#2, #3, and this item ship as one release with the engine.**

- TODO: **BFN critical values** (follow-on, ships with 2.1.0 if there's room). N,T-dependent bounds per Bhargava et al. 1982 tables; do NOT present standard DW bounds next to the BFN cell.

- TODO: **Categorical × FE prediction encoding** (follow-on, ships with 2.1.0 if there's room). x_new and x̄ᵢ formed in constructed design-matrix space; UI wire to encode through `Dummy_Code` before reaching the FE formula. Largely subsumed by v2.0 categorical prediction; recorded so the encoding step is not forgotten.

- TODO: **Relabel within-model residual outputs + Diagnostic Guide paragraph on residuals under FE** (follow-on, ships with 2.1.0 if there's room). Documentation-only.

---

## v2.2 — Transforms & the standalone transform library

### Transform wiring (spec column G)

- TODO: `Transform` dropdown gains `Log`; wire `X_s()` / `Constructed_Column_Names()` /
  prediction to read column G.
- TODO: **Unit-space dispatcher function, RESOLVED:** `Unit_Space_R_Squared(model, response_transform, predictor_transform)`,
  `Unit_Space_Adjusted_R_Squared(...)`, `Unit_Space_RMSE(...)`. Single dispatcher per
  statistic, internal `SWITCH` on the `(response_transform, predictor_transform)` pair;
  `NA()` on unrecognised values. Argument order: `model` first, then `response_transform`
  then `predictor_transform` (matches spec-block reading order on column G).
- TODO: Unit-space section on the Regression sheet — SWITCH on column G, one headline
  comparable statistic (the cell v2.3 Model Comparison will reference).
- TODO: **Prediction back-transformation, RESOLVED:** Duan's smearing estimator as the
  default, with a per-cell `Back_Transform_Method` toggle (`Duan` default | `Naive`).
  Caveat row visible on the sheet:
  *Duan = Duan (1983) smearing; Naive = textbook EXP(ŷ), biased.*

### Standalone Data Transformation functions (specs in ROADMAP.md)

- TODO: Location & Scale — `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Ln_Positive`.
- TODO: Group & Panel — `Zscore_By`, `Decompose_By` (`Demean_By`/`Group_Mean` arrive at
  v2.1; two-way functions follow the two-way FE milestone).
- ~~Longitudinal — `Lag_By`, `Difference_By`~~ — **DONE (shipped early, base-period
  release)** with the gap-aware t−Δ semantics: exact-match lookup of
  (group, seq−Δ) pairs, `NA()` at first periods and gaps, `[delta]` defaulting
  to the spec's Period In Use cell via `Base_Period_Delta()` (never a silent 1).
  The same release wired spec column I (typed override → Sequence Period) and
  J (candidate-with-override display → Period In Use) plus the Sequence
  Spacing block (delta spectrum, Regularity/Off-grid flags, calendar-signature
  guidance). Verification: `tests/test_difference_by_verification.py`; human
  test plan T17–T19.
- TODO: Sample construction — `Numeric_Complete_Cases`.
- TODO: Categorical & model construction — `Dummy_Column`, `Interact`, `Model_Matrix`.

---

## v2.3 — Model Comparison Sheet

- ~~TODO: Resolve the spec-string function name (~~`Regression_Model_Spec_String` ~~vs.~~ ~~`Regression_Spec_Label`~~ ~~vs.~~ ~~`Model_Formula_String`~~ ~~) and the argument type (lean: anchor-cell reference, not sheet-name text — avoids volatile `INDIRECT`).~~ **RESOLVED.** Function name: **`Model_Formula_String(anchor_cell)`**. Argument type: **anchor cell** (consistent with the project's `INDIRECT`-avoidance stance).
- TODO: Implement the `Model_Formula_String` LAMBDA with header-signature validation (`NA()` on
  non-Regression targets).
- TODO: Sheet layout — model registry (hyperlinks), GoF table referencing the v2.2
  unit-space headline cells, shared prediction inputs (Comparison sheet is the source;
  Regression sheets pull via XLOOKUP), prediction results table.
- ~~TODO: Formalize `Comparison_Anchor` sheet-scoped named ranges (interface contract —
  becomes part of the public interface, a versioning commitment).~~ **RESOLVED:** three
  named ranges ship at v2.3 as a public-interface commitment:
  - `Comparison_Anchor` — single anchor cell in the status block (the
    `Model_Formula_String` first-argument; the model-registry hyperlink target)
  - `Comparison_Headline_GoF` — the v2.2 unit-space headline cells (R²,
    Adjusted R², RMSE) — the GoF table's source
  - `Comparison_Prediction_Output` — the v2.1 prediction outputs center cell —
    the prediction-results table's source

  All three are sheet-scoped and the changelog entry for v2.3.0 must name them
  explicitly so the public-interface commitment is discoverable.
- TODO: Decide the mismatched-predictor-set fallback (XLOOKUP `[if_not_found]`).

---

## v2.4 — Resampling & Simulation

- TODO: **No-volatile constraint, RESOLVED: pre-drawn random table.** A single
  sheet-scoped named range `Bootstrap_Random_Draws` holds a uniformly-distributed
  random table pre-drawn once at build time, seeded from the same SHA-derived seed
  the QC build already uses (`analysis_cache.py`). `Bootstrap_CI` indexes via
  `INDEX(Bootstrap_Random_Draws, MOD(SEQUENCE(n_resamples), ROWS(Bootstrap_Random_Draws))+1)`.
  Same inputs → same output, every recalc. `RANDARRAY()` rejected: silently re-drawing
  per recalc is the opposite of the library's auditability philosophy. To get a new
  draw, regenerate the workbook via `build_production.py` (deliberate, not a limitation).
- TODO: Implement `Bootstrap_CI(data, stat_lambda, n_resamples, alpha, [include])` — bootstrap
  confidence interval for an arbitrary statistic passed as a LAMBDA. Uses the pre-drawn
  table above.
- TODO: Implement `MC_Percentile(dist_params, n_samples, percentile)` — Monte Carlo draw
  from a fitted distribution; complements v2.0 fitting. Uses the same pre-drawn table.
- TODO: Implement `PERT_Sample(min, mode, max, n_samples)` — BetaPERT sampling for
  cost/schedule risk analysis. Uses the same pre-drawn table.
- TODO: Design sheet layout (bootstrap section + Monte Carlo section; may share one sheet).
  Implement `write_sheet_simulation.py`.

---

## v2.5+ — Future (sequence TBD; first two claimed)

The v2.5+ bucket previously had seven candidates with no order. Two-sample
tests are now v2.5 (next MINOR after v2.4) and the `Weight` Role is v2.6
(after v2.5). The rest are deliberately unordered pending actual user
demand — a single maintainer should not pre-order work that may not be
the next thing actually needed.

### v2.5 — Bivariate / Two-sample *(claimed, next MINOR after v2.4)*

- TODO: Implement `T_Test_OneSample(data, mu0, alpha, [include])` → test statistic, p-value, CI.
- TODO: Implement `T_Test_TwoSample(data1, data2, alpha, equal_var, [include1], [include2])` —
  equal-variance, Welch, and paired variants. Open design question: paired is a separate
  code path the `equal_var` flag does not cover — 3-way flag or separate `paired` boolean?
- TODO: Implement `F_Test_Variance(data1, data2, alpha, [include1], [include2])` — output
  feeds a recommendation cell that selects the appropriate t-test variant.
- TODO: Implement `Covariance_Matrix(data, [include])` — sample covariance (consistent
  with the existing catalog's sample-statistic convention); complement to the existing
  `Correlation_Matrix`.
- TODO: Design two-sample sheet layout: inputs, test selector, F-test assumption check,
  output (test statistic, df, p-value, CI, effect size). Implement `write_sheet_two_sample.py`.

### v2.6 — `Weight` Role (WLS) *(claimed, after v2.5)*

The standalone WLS milestone and its `[weights]`-argument-vs-parallel-function-set
debate are superseded by a **`Weight` value on the Role axis** (see ROADMAP *Future
roles*). Three-stage scope carried forward: user-supplied weights →
variance-driver-derived weights → FGLS. v2.6 ships the first stage only.

- TODO: Implement the `Weight` Role (at most one, per the cardinality rule that
  Response, Time, and Weight share; status-block validation identical to
  exactly-one-Response).
- TODO: Thread weights through the engine per the Role-axis design: a single optional
  `[Weights]` argument (default uniform) on the inferential chain. Default-uniform
  means every existing OLS call computes identically — the v2.1 `[DF_Absorbed]`
  precedent (default 0 → identical no-FE model) is the exact pattern to follow.
- TODO: Update the Diagnostic Guide to describe which diagnostics change interpretation
  under WLS. (WLS closes the loop opened by v1's Scale-Location diagnostic.)

### v2.7+ — Unordered candidates (no claim)

The following are real candidate work but deliberately unordered. Two-way FE
and `Cluster` have partial forward wiring (from v2.1 FE and the
`Serial_Correlation_Group()` resolver); the rest are design-not-started.
A user-pressing-for-them signal would reorder these; absent that, they
stay in this unordered bucket.

#### Two-way Fixed Effects

- TODO: Implement `Absorb_Two_Way_Fixed_Effects(x, group1, group2, [include], [passes])`
  (alternating-projection demeaning for unbalanced panels).
- TODO: Implement `Demean_Two_Way_Balanced(x, group1, group2, [include])` and the
  two-way `Is_Balanced_Panel` check.
- TODO: Implement `Fixed_Effects_Convergence_Check(x, group1, group2, [include])`;
  surface in the status block whenever two FE variables are active.
- TODO: Lift the v2.1 one-FE-variable status-block error; resolve the two-way
  prediction question (group intercepts are not recoverable as simple group means).

#### Multi-group means (ANOVA)

- TODO: Implement one-way ANOVA as regression on group dummies, reusing the existing SS/MS/F machinery. Frame explicitly as "ANOVA is regression" — group means, SS decomposition, and F-test should match the MLR output exactly.
- TODO: Add post-hoc comparisons (Tukey HSD or Bonferroni) as an optional output section.

#### `Cluster` Role (clustered SEs)

- TODO: Implement the `Cluster` Role (at most one) — clustered-robust variance estimator.
  Has partial forward wiring from `Serial_Correlation_Group()`'s dormant Cluster branch
  (PR #106), but the engine-side estimator (cluster-robust V_β) is not implemented.
- TODO: Lift the v2.1 `n/a — engine forthcoming` token on the BFN cell when Cluster is
  active (the BFN formula already uses `Serial_Correlation_Group()` as its resolver, so
  the wiring is partial).

#### `Time` Role (time-index designation)

- TODO: Design and implement the `Time` Role. Partially forward-wired via the v2.1
  Sequence axis, but the full `Time` Role adds time-index semantics (for the future
  time-series sheet, for cross-sheet `Lag_By`/`Difference_By` calls). Open design
  question: can a column be both `Sequence` and `Time`, or are they mutually exclusive?

#### Time series

- TODO: Implement `Moving_Average(data, window, [include])`.
- TODO: Implement `Exponential_Smoothing(data, alpha_smooth, [include])` — note: use
  `alpha_smooth` to distinguish from the significance-level `alpha`.
- TODO: Implement `write_sheet_time_series.py` with forecast output, error metrics
  (MAE, RMSE, MAPE), and an actual vs. smoothed series chart.

#### Long-tail (out of planning horizon)

- **Fourier analysis** — long-tail; the *ToolPak Parity Reference* notes it is
  "intentionally skipped" and a later addition-by-demand decision, not a planned milestone.
- **Decision analysis** — long-tail (loss functions, cost/risk oriented). Not on the
  planning horizon.
