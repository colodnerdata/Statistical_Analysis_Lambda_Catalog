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

**CDF/PDF functions (histogram overlay curves)**
- TODO: Implement `PDF_Normal`, `PDF_Lognormal`, `PDF_Exponential`, `PDF_Weibull`, `PDF_Gamma`, `PDF_Triangular`, `PDF_Beta`, `PDF_BetaPERT` — evaluated at bin midpoints.

### Sheet writer (`write_sheet_univariate.py`)

**Q-Q plots and histogram overlays**
- TODO: Per-distribution Q-Q plots (8 charts) using OFFSET-based named ranges.
- TODO: Histogram overlay: fitted PDF curve on each histogram using bin midpoints as X values (depends on PDF LAMBDAs).

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

## v2.1 — Fixed Effects (one-way only)

Two-way FE is deliberately deferred until this framework is finished — see the
v2.5+ section.

### Engine

- TODO: Implement `y_s()` — the demeaned-Response constructor (new function, not a
  replacement wired into existing no-FE call sites).
- TODO: Thread the optional `[DF_Absorbed]` argument (default 0) through the df /
  MS-residual / t-critical inference chain; no-FE models must compute identically.
- TODO: Implement `Demean_By(x, group, [include])` and `Group_Mean(x, group, [include])`
  (constructor internals, also user-callable).
- TODO: Implement `Absorbed_Degrees_Of_Freedom(...)` — Σ(Gᵢ − 1) from the spec.
- TODO: Implement `Is_Balanced_Panel(group, time, [include])` (one-way/panel diagnostic).

### Sheet

- TODO: Restructure the prediction zone into the general group-mean form
  ŷ = ȳᵢ + (x_new − x̄ᵢ)′β̂ with the whole sample as the degenerate G = 1 group
  (v2.0 shipped the standard `Prediction_Interval` form, so this is a rebuild, not an
  activation — see the ROADMAP post-ship correction).
- TODO: Surface BOTH intervals — mean-response CI and new-observation PI (three lines:
  point · CI low/high · PI low/high).
- TODO: FE group selection dropdown sourced from the observed level list; ȳᵢ / x̄ᵢ / Tᵢ
  via AVERAGEIFS/COUNTIFS respecting the Include/Filter mask.
- TODO: Status block — active FE variable, group count, absorbed df; visible error when
  more than one FE variable is declared; intercept × FE red flag (flag, don't force).
- TODO: Relabel within-model residual outputs; Diagnostic Guide paragraph on residuals
  under FE.

### Open decisions

- TODO: Durbin-Watson under FE — relabel, caveat, or suppress.
- TODO: Categorical × FE prediction encoding — x_new and x̄ᵢ formed in constructed
  design-matrix space (largely subsumed by v2.0 categorical prediction; recorded so the
  encoding step is not forgotten).

---

## v2.2 — Transforms & the standalone transform library

### Transform wiring (spec column G)

- TODO: `Transform` dropdown gains `Log`; wire `X_s()` / `Constructed_Column_Names()` /
  prediction to read column G.
- TODO: Unit-space fit statistics (R², Adjusted R², RMSE at minimum) — resolve the
  one-LAMBDA-per-combination vs. `Unit_Space_*` dispatcher decision BEFORE implementing
  (sets the pattern for every future transform).
- TODO: Unit-space section on the Regression sheet — SWITCH on column G, one headline
  comparable statistic (the cell v2.3 Model Comparison will reference).
- TODO: Prediction back-transformation — decide naive `EXP()` with documented caveat
  vs. Duan smearing estimator (a statistical decision, not an implementation detail).

### Standalone Data Transformation functions (specs in ROADMAP.md)

- TODO: Location & Scale — `Center`, `Zscore`, `Minmax_Scale`, `Winsorize`, `Ln_Positive`.
- TODO: Group & Panel — `Zscore_By`, `Decompose_By` (`Demean_By`/`Group_Mean` arrive at
  v2.1; two-way functions follow the two-way FE milestone).
- TODO: Longitudinal — `Lag_By`, `Difference_By`.
- TODO: Sample construction — `Numeric_Complete_Cases`.
- TODO: Categorical & model construction — `Dummy_Column`, `Interact`, `Model_Matrix`.

---

## v2.3 — Model Comparison Sheet

- TODO: Resolve the spec-string function name (`Regression_Model_Spec_String` vs.
  `Regression_Spec_Label` vs. `Model_Formula_String`) and the argument type (lean:
  anchor-cell reference, not sheet-name text — avoids volatile `INDIRECT`).
- TODO: Implement the spec-string LAMBDA with header-signature validation (`NA()` on
  non-Regression targets).
- TODO: Sheet layout — model registry (hyperlinks), GoF table referencing the v2.2
  unit-space headline cells, shared prediction inputs (Comparison sheet is the source;
  Regression sheets pull via XLOOKUP), prediction results table.
- TODO: Formalize `Comparison_Anchor` sheet-scoped named ranges (interface contract —
  becomes part of the public interface, a versioning commitment).
- TODO: Decide the mismatched-predictor-set fallback (XLOOKUP `[if_not_found]`).

---

## v2.4 — Resampling & Simulation

- TODO: Implement `Bootstrap_CI(data, stat_lambda, n_resamples, alpha, [include])` — bootstrap confidence interval for an arbitrary statistic passed as a LAMBDA. Evaluate whether `RANDARRAY`-based resampling is viable or whether a pre-drawn random table is needed.
- TODO: Implement `MC_Percentile(dist_params, n_samples, percentile)` — Monte Carlo draw from a fitted distribution; complements v2.0 fitting.
- TODO: Implement `PERT_Sample(min, mode, max, n_samples)` — BetaPERT sampling for cost/schedule risk analysis.
- TODO: Design sheet layout (bootstrap section + Monte Carlo section; may share one sheet). Implement `write_sheet_simulation.py`.

---

## v2.5+ — Future (sequence TBD)

### Two-way Fixed Effects (first candidate after v2.1)

- TODO: Implement `Absorb_Two_Way_Fixed_Effects(x, group1, group2, [include], [passes])`
  (alternating-projection demeaning for unbalanced panels).
- TODO: Implement `Demean_Two_Way_Balanced(x, group1, group2, [include])` and the
  two-way `Is_Balanced_Panel` check.
- TODO: Implement `Fixed_Effects_Convergence_Check(x, group1, group2, [include])`;
  surface in the status block whenever two FE variables are active.
- TODO: Lift the v2.1 one-FE-variable status-block error; resolve the two-way
  prediction question (group intercepts are not recoverable as simple group means).

### Weighted regression — superseded by the `Weight` Role

The standalone WLS milestone and its `[weights]`-argument-vs-parallel-function-set
debate are superseded by a **`Weight` value on the Role axis** (see ROADMAP *Future
roles*). Three-stage scope carried forward: user-supplied weights →
variance-driver-derived weights → FGLS.

- TODO: Implement the `Weight` Role (at most one; status-block validation) and thread
  weights through the engine per the Role-axis design.
- TODO: Update the Diagnostic Guide to describe which diagnostics change interpretation
  under WLS. (WLS closes the loop opened by v1's Scale-Location diagnostic.)

### Bivariate / Two-sample

- TODO: Implement `T_Test_OneSample(data, mu0, alpha, [include])` → test statistic, p-value, CI.
- TODO: Implement `T_Test_TwoSample(data1, data2, alpha, equal_var, [include1], [include2])` — equal-variance, Welch, and paired variants via `equal_var` flag.
- TODO: Implement `F_Test_Variance(data1, data2, alpha, [include1], [include2])` — test for equality of variances; output feeds a recommendation cell that selects the appropriate t-test variant.
- TODO: Implement `Covariance_Matrix(data, [include])` — complement to the existing `Correlation_Matrix`.
- TODO: Design two-sample sheet layout: inputs, test selector, F-test assumption check, output (test statistic, df, p-value, CI, effect size). Implement `write_sheet_two_sample.py`.

### Multi-group means (ANOVA)

- TODO: Implement one-way ANOVA as regression on group dummies, reusing the existing SS/MS/F machinery. Frame explicitly as "ANOVA is regression" — group means, SS decomposition, and F-test should match the MLR output exactly.
- TODO: Add post-hoc comparisons (Tukey HSD or Bonferroni) as an optional output section.

### Time series

- TODO: Implement `Moving_Average(data, window, [include])`.
- TODO: Implement `Exponential_Smoothing(data, alpha_smooth, [include])` — note: use `alpha_smooth` to distinguish from the significance-level `alpha`.
- TODO: Implement `write_sheet_time_series.py` with forecast output, error metrics (MAE, RMSE, MAPE), and an actual vs. smoothed series chart.
