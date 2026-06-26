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

## v1.x — Regression sheet

- TODO: Add reference lines to the Cook's Distance, PRESS Residuals, and Leverage vs. Studentized charts. Format thematically similar to the conditional formatting in the table (yellow = mild, red = strong). Use minimalist helper columns (2 anchor points using max/min of the relevant threshold; place them beneath the chart area).

---

## v2.0 — Univariate

### LAMBDA functions

**Descriptive statistics**
- TODO: Implement `Descriptive_Stats(data, [include])` — returns a column vector: mean, median, mode, SD, variance, min, max, range, skewness, kurtosis, count, missing count. Reuse existing `Skewness` and `Kurtosis` LAMBDAs.
- TODO: Implement `Missing_Count(data, [include])` — count of blank/non-numeric cells; surface prominently in the sheet.
- TODO: Resolve `Missing_Count` handling for formula blanks (`""`). Excel formulas like `IF(...,"",...)` produce visually blank cells; the workbook `Missing_Count` LAMBDA currently uses `d<>""` to find the last active row, while `lambda_catalog.analyze_univariate.missing_count()` does not treat `""` as blank and can count trailing formula blanks as missing. Decide whether the LAMBDA should change or whether the Python oracle/tests should be updated to match the intended behavior.

**Histogram binning**
- TODO: Implement `Sturges_Bins(data, [include])` → bin count via `CEILING(LOG(n,2)+1, 1)`.
- TODO: Implement `Scott_Bins(data, [include])` → bin width `3.49 × SD × n^(-1/3)`; bin count from range/width.
- TODO: Implement `FD_Bins(data, [include])` → Freedman-Diaconis bin width `2 × IQR × n^(-1/3)`.
- TODO: Implement `Bin_Edges(data, method, [include])` → vector of upper bin edges (half-open `(lower, upper]` intervals, matching Excel `FREQUENCY` convention).
- TODO: Implement `Bin_Counts(data, edges, [include])` → frequency vector for the given edges.

**Distribution fitting — negative log-likelihood (NLL)**
- TODO: Implement `NLL_Normal(data, mean, sd, [include])`.
- TODO: Implement `NLL_Lognormal(data, meanlog, sdlog, [include])`.
- TODO: Implement `NLL_Exponential(data, rate, [include])`.
- TODO: Implement `NLL_Weibull(data, shape, scale, [include])` — two-parameter; grid-search MLE.
- TODO: Implement `NLL_Gamma(data, shape, rate, [include])` — two-parameter; grid-search MLE.
- TODO: Implement `NLL_Triangular(data, min, mode, max, [include])` — fit by direct min/mode/max estimation (likelihood non-differentiable at mode; not grid-search).
- TODO: Implement `NLL_Beta(data, alpha_param, beta_param, [include])` — requires data rescaled to `[0,1]`; wrap in `IFERROR` sentinel for values exactly at 0 or 1.
- TODO: Implement `NLL_BetaPERT(data, min, mode, max, [include])` — closed-form via PERT reparameterization of Beta.

**Goodness-of-fit statistics**
- TODO: Implement `GoF_AIC(nll, k)` — `2k + 2 × NLL`; comparable across distributions on the same data.
- TODO: Implement `GoF_BIC(nll, k, n)` — `k × ln(n) + 2 × NLL`.
- TODO: Implement `GoF_AndersonDarling(data, dist_cdf, [include])` — handle bounded-support distributions (Beta, Triangular, BetaPERT) at support edges explicitly.
- TODO: Implement `GoF_KS(data, dist_cdf, [include])` — Kolmogorov-Smirnov statistic.

**CDF/PDF functions (histogram overlay curves)**
- TODO: Implement `PDF_Normal`, `PDF_Lognormal`, `PDF_Exponential`, `PDF_Weibull`, `PDF_Gamma`, `PDF_Triangular`, `PDF_Beta`, `PDF_BetaPERT` — evaluated at bin midpoints.

### Sheet writer (`write_sheet_univariate.py`)

**Design decisions to resolve before building**
- TODO: Confirm fitting approach for closed-form-friendly distributions: method of moments (fully live, formula-transparent) vs. MLE (statistically principled). Document in module docstring.
- TODO: Design three-zone sheet layout: (1) descriptive statistics + histogram, (2) distribution fitting grid-search tables, (3) fit comparison table + per-distribution Q-Q plots. Document before implementing.
- TODO: Add support for more distribution families: Bernoulli, Binomial, Geometric, Negative Binomial, Hypergeometric, Poisson, Uniform, Chi-Square, Student-t.


**Grid-search MLE (two-input Data Table)**
- TODO: Two-input Data Table grid for Weibull, Gamma, and Beta: row headers = parameter 1 candidates, column headers = parameter 2 candidates, corner cell = NLL. Grid fills all combinations via native Excel recalculation.
- TODO: 2D argmin lookup: flatten the grid with `TOCOL`, locate `MIN`, recover row/column via integer division and modulo against grid width to extract fitted parameters.
- TODO: Multi-stage refinement (2 stages minimum, 3 available): each stage centers on the prior stage's minimum. Next-stage bounds must extend one current-grid step on either side of the located cell — do not use a fixed shrink factor.
- TODO: Boundary guard: conditionally format the located parameter cell red when it lands on the first or last row/column of the grid.
- TODO: Apply `IFERROR` sentinel (large finite value) to NLL cells to handle overflow/undefined regions.
- TODO: Color-scale conditional formatting on the NLL grid as a heatmap — confirms the minimum is interior and reveals overflow zones.

**Histogram section**
- TODO: Three side-by-side bin tables (Sturges, Scott, Freedman-Diaconis) with column charts (gap width 0). Use OFFSET-based named ranges for chart series per the convention in CONTRIBUTING.md.

**Fit comparison table and charts**
- TODO: One row per distribution; columns: AIC, BIC, Anderson-Darling, K-S, QQ Correlation. Highlight best-fit row.
- TODO: Per-distribution Q-Q plots (8 charts) using OFFSET-based named ranges.
- TODO: Histogram overlay: fitted PDF curve on each histogram using bin midpoints as X values.

---

## v3.0 — Resampling & Simulation

- TODO: Implement `Bootstrap_CI(data, stat_lambda, n_resamples, alpha, [include])` — bootstrap confidence interval for an arbitrary statistic passed as a LAMBDA. Evaluate whether `RANDARRAY`-based resampling is viable or whether a pre-drawn random table is needed.
- TODO: Implement `MC_Percentile(dist_params, n_samples, percentile)` — Monte Carlo draw from a fitted distribution; complements v2.0 fitting.
- TODO: Implement `PERT_Sample(min, mode, max, n_samples)` — BetaPERT sampling for cost/schedule risk analysis.
- TODO: Design sheet layout (bootstrap section + Monte Carlo section; may share one sheet). Implement `write_sheet_simulation.py`.

---

## v4.0 — Bivariate / Two-sample

- TODO: Implement `T_Test_OneSample(data, mu0, alpha, [include])` → test statistic, p-value, CI.
- TODO: Implement `T_Test_TwoSample(data1, data2, alpha, equal_var, [include1], [include2])` — equal-variance, Welch, and paired variants via `equal_var` flag.
- TODO: Implement `F_Test_Variance(data1, data2, alpha, [include1], [include2])` — test for equality of variances; output feeds a recommendation cell that selects the appropriate t-test variant.
- TODO: Implement `Covariance_Matrix(data, [include])` — complement to the existing `Correlation_Matrix`.
- TODO: Design two-sample sheet layout: inputs, test selector, F-test assumption check, output (test statistic, df, p-value, CI, effect size). Implement `write_sheet_two_sample.py`.

---

## v4.x — Multi-group means (ANOVA)

- TODO: Implement one-way ANOVA as regression on group dummies, reusing the existing SS/MS/F machinery. Frame explicitly as "ANOVA is regression" — group means, SS decomposition, and F-test should match the MLR output exactly.
- TODO: Add post-hoc comparisons (Tukey HSD or Bonferroni) as an optional output section.

---

## v4.x — Weighted regression (WLS)

- TODO: Thread a `[weights]` argument through the core regression LAMBDAs (`Coefficients`, `Predictions`, `Residuals`, `Hat_diagonal`, `Cooks_Distance`, etc.). WLS closes the loop opened by v1's Scale-Location diagnostic.
- TODO: Update the Regression sheet to expose a weights column selector.
- TODO: Update the Diagnostic Guide to describe which diagnostics change interpretation under WLS.

---

## v4.x — Time series

- TODO: Implement `Moving_Average(data, window, [include])`.
- TODO: Implement `Exponential_Smoothing(data, alpha_smooth, [include])` — note: use `alpha_smooth` to distinguish from the significance-level `alpha`.
- TODO: Implement `write_sheet_time_series.py` with forecast output, error metrics (MAE, RMSE, MAPE), and an actual vs. smoothed series chart.
