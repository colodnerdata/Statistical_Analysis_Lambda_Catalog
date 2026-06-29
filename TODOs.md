# TODOs

## Design note — chart series data ranges

Chart `SERIES` formulas require explicit range references; the `#` spill operator is not reliably supported in chart series formulas. However, referencing all 1,048,576 rows can significantly degrade performance or crash Excel when the populated dataset is much smaller.

Instead, define dynamically sized named ranges using the row count in `$M$8`. For example:

```excel
RegChartQQX = OFFSET($AE$2,1,0,$M$8,1)
RegChartQQY = OFFSET($AF$2,1,0,$M$8,1)
```

These formulas define ranges beginning at `AE3` and `AF3`, respectively, and extending for exactly the number of rows specified in `M8`. All chart series names use the `RegChart` prefix and are documented in CONTRIBUTING.md.

---

## Alias layer

Design/planning item — see ROADMAP.md for the architectural rationale. Aliases are thin ALL-CAPS wrappers whose entire body is a single call to the canonical function; the canonical function remains the single source of truth. Implement only after the canonical library is stable.

Suggested alias names:

**Regression — scalar outputs**

| Alias | Canonical |
|---|---|
| `R2` | `R_squared` |
| `ADJ_R2` | `Adjusted_R2` |
| `MULT_R` | `Multiple_R` |
| `SE_REG` | `SE_Regression` |
| `DW` | `Durbin_Watson` |
| `QQ_CORR` | `QQ_Correlation` |
| `DFR` | `DF_Regression` |
| `DFE` | `DF_Residual` |
| `DFT` | `DF_Total` |
| `SSR` | `SS_Regression` |
| `SSE` | `SS_Residual` |
| `SST` | `SS_Total` |
| `MSR` | `MS_Regression` |
| `MSE` | `MS_Residual` |
| `P_VAL_F` | `P_Value_F` |

**Regression — coefficient vectors**

| Alias | Canonical |
|---|---|
| `COEF` | `Coefficients` |
| `SE_COEF` | `SE_Coefficients` |
| `TSTAT` | `T_Stats` |
| `PVALS` | `P_Values` |
| `CI_LOW` | `CI_Lower` |
| `CI_UP` | `CI_Upper` |
| `P_R2` | `Partial_R2` |
| `P_COR` | `Partial_Correlation` |
| `BETA_W` | `Beta_Weights` |

**Regression — observation vectors**

| Alias | Canonical |
|---|---|
| `PRED` | `Predictions` |
| `RESID` | `Residuals` |
| `STDR` | `Studentized_Residuals` |
| `LEV` | `Hat_diagonal` |
| `COOK_D` | `Cooks_Distance` |
| `LOOCV` | `LOOCV_prediction` |
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
| `DSTAT` | `Descriptive_Stats` |
| `NMISS` | `Missing_Count` |

**Univariate — histogram binning**

| Alias | Canonical |
|---|---|
| `NBINS` | `num_histogram_bins` |
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
| `GS_MIN` | `Grid_Argmin` |
| `GS_OPT` | `Grid_Search_Optimum` |

---

## v1.x — Regression sheet

- TODO: Add reference lines to the Cook's Distance, PRESS Residuals, and Leverage vs. Studentized charts. Format thematically similar to the conditional formatting in the table (yellow = mild, red = strong). Use minimalist helper columns (2 anchor points using max/min of the relevant threshold; place them beneath the chart area).

---

## v2.0 — Univariate

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

## v3.0 — Resampling & Simulation

- TODO: Implement `Bootstrap_CI(data, stat_lambda, n_resamples, alpha, [include])` — bootstrap confidence interval for an arbitrary statistic passed as a LAMBDA. Evaluate whether `RANDARRAY`-based resampling is viable or whether a pre-drawn random table is needed.
- TODO: Implement `MC_Percentile(dist_params, n_samples, percentile)` — Monte Carlo draw from a fitted distribution; complements v2.0 fitting.
- TODO: Implement `PERT_Sample(min, mode, max, n_samples)` — BetaPERT sampling for cost/schedule risk analysis.
- TODO: Design sheet layout (bootstrap section + Monte Carlo section; may share one sheet). Implement `write_sheet_simulation.py`.

---

## v4.0 — Future (sequence TBD)

### Weighted regression (WLS)

- TODO: Thread a `[weights]` argument through the core regression LAMBDAs (`Coefficients`, `Predictions`, `Residuals`, `Hat_diagonal`, `Cooks_Distance`, etc.). WLS closes the loop opened by v1's Scale-Location diagnostic.
- TODO: Update the Regression sheet to expose a weights column selector.
- TODO: Update the Diagnostic Guide to describe which diagnostics change interpretation under WLS.

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
