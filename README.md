# Statistical Analysis Lambda Catalog

Excel 365 LAMBDA functions that replicate and extend Analysis ToolPak regression statistics — no VBA, no add-ins, no installation. Download `Lambda_Library.xlsx`, open it in Excel, and all functions are immediately available in formulas.

## Getting started

1. Download `Lambda_Library.xlsx` from this repository.
2. Open it in Excel 365 (Windows or Mac).
3. Enter your data in columns on any sheet, then call any function by name.

All functions are defined as workbook-scoped names, so they work in any cell formula within the workbook. The **Regression** sheet provides a ready-to-use analysis interface: select your predictor and outcome columns, and it computes the full regression output automatically. The **Univariate Analysis** sheet demonstrates descriptive statistics, histogram binning, and distribution fitting via grid-search MLE. The **LAMBDA_functions** sheet is a browsable catalog of every function with its full description and argument documentation.

To use these functions in a different workbook, open both files in Excel at the same time. You can reference functions as `='[Lambda_Library.xlsx]'!FunctionName(args)`, or use Name Manager (Formulas → Name Manager → New) to copy individual definitions into your own workbook.

## Common function signature

Most regression functions share this signature:

```
FunctionName(X_s, Y, [Allow_Intercept], [Include])
```

| Argument | Description |
|---|---|
| `X_s` | Predictor column range (one or more columns) |
| `Y` | Outcome column range (single column) |
| `[Allow_Intercept]` | TRUE to fit an intercept (default), FALSE to force through the origin |
| `[Include]` | Boolean column — TRUE includes the row, FALSE excludes it |

Square brackets indicate optional arguments. A few functions have different signatures, as noted in their entries below.

## Function reference

### Model fit — scalar outputs

These functions return a single number summarizing the regression.

| Function | Returns |
|---|---|
| `Observations(Y, [Include])` | Number of observations (n) |
| `Regression_Degrees_Of_Freedom(X_s)` | Degrees of freedom for the model (k predictors) |
| `Residual_Degrees_Of_Freedom(X_s, Y, [Allow_Intercept], [Include])` | Degrees of freedom for residuals (n − k − 1) |
| `Total_Degrees_Of_Freedom(Y, [Allow_Intercept], [Include])` | Total degrees of freedom (n − 1) |
| `R_Squared(X_s, Y, [Allow_Intercept], [Include])` | R² — proportion of variance explained, 0 to 1 |
| `Multiple_R(X_s, Y, [Allow_Intercept], [Include])` | √R² — multiple correlation coefficient |
| `Adjusted_R_Squared(X_s, Y, [Allow_Intercept], [Include])` | R² penalized for number of predictors |
| `SE_Regression(X_s, Y, [Allow_Intercept], [Include])` | Standard error of the regression |
| `SS_Regression(X_s, Y, [Allow_Intercept], [Include])` | Model sum of squares |
| `SS_Residual(X_s, Y, [Allow_Intercept], [Include])` | Residual (error) sum of squares |
| `SS_Total(Y, [Allow_Intercept], [Include])` | Total sum of squares |
| `MS_Regression(X_s, Y, [Allow_Intercept], [Include])` | Mean square for the model (SS_Regression / df_Regression) |
| `MS_Residual(X_s, Y, [Allow_Intercept], [Include])` | Mean square for residual error (SS_Residual / df_Residual) |
| `PRESS(X_s, Y, [Allow_Intercept], [Include])` | Leave-one-out cross-validation error sum |
| `Durbin_Watson(X_s, Y, [Allow_Intercept], [Include])` | Serial autocorrelation test for residuals (2 = no autocorrelation) |
| `AIC(X_s, Y, [Allow_Intercept], [Include])` | Akaike Information Criterion |
| `AICc(X_s, Y, [Allow_Intercept], [Include])` | Corrected AIC (for small samples) |
| `BIC(X_s, Y, [Allow_Intercept], [Include])` | Bayesian Information Criterion |
| `QQ_Correlation(X_s, Y, [Allow_Intercept], [Include])` | Filliben Q-Q normality statistic for residuals |

### Coefficient-level outputs — (k+1)×1 vector

These functions return one value per coefficient (intercept first when included, then predictors in input order). Enter them as array formulas and let them spill.

| Function | Returns |
|---|---|
| `Coefficients(X_s, Y, [Allow_Intercept], [Include])` | OLS coefficient estimates |
| `SE_Coefficients(X_s, Y, [Allow_Intercept], [Include])` | Standard errors of the coefficients |
| `T_Statistics(X_s, Y, [Allow_Intercept], [Include])` | t-statistics (coefficient / standard error) |
| `P_Values(X_s, Y, [Allow_Intercept], [Include])` | Two-tailed p-values |
| `Confidence_Interval_Lower(X_s, Y, [Allow_Intercept], [Include], [Alpha])` | Lower confidence interval bounds (default 95%) |
| `Confidence_Interval_Upper(X_s, Y, [Allow_Intercept], [Include], [Alpha])` | Upper confidence interval bounds (default 95%) |
| `Partial_R_Squared(X_s, Y, [Allow_Intercept], [Include])` | Partial R² per coefficient |
| `Partial_Correlation(X_s, Y, [Allow_Intercept], [Include])` | Partial correlation per coefficient |
| `Beta_Weights(X_s, Y, [Allow_Intercept], [Include])` | k×1 vector of standardized coefficients (Beta weights) — intercept excluded |

### Multicollinearity — k×1 vector

These functions return one value per predictor (no intercept row). `VIF > 10` or `Tolerance < 0.1` indicates severe collinearity.

| Function | Returns |
|---|---|
| `VIF(X_s, [Allow_Intercept], [Include])` | Variance inflation factor per predictor |
| `Tolerance(X_s, [Allow_Intercept], [Include])` | Tolerance (= 1 / VIF) per predictor |

### Observation-level outputs — n×1 vector

These functions return one value per observation in the filtered dataset, spilled in original data order.

| Function | Returns |
|---|---|
| `Predictions(X_s, Y, [Allow_Intercept], [Include])` | Fitted (predicted) values |
| `Residuals(X_s, Y, [Allow_Intercept], [Include])` | Raw residuals (Y − Ŷ) |
| `Scaled_Residuals(X_s, Y, [Allow_Intercept], [Include])` | Residuals divided by SE_Regression |
| `Scaled_Residuals_Ranked(X_s, Y, [Allow_Intercept], [Include])` | Scaled residuals sorted ascending |
| `Studentized_Residuals(X_s, Y, [Allow_Intercept], [Include])` | Internally studentized residuals |
| `Studentized_Residuals_Ranked(X_s, Y, [Allow_Intercept], [Include])` | Studentized residuals sorted ascending |
| `Hat_Diagonal(X_s, [Allow_Intercept], [Include])` | Leverage values (diagonal of the hat matrix) |
| `Cooks_Distance(X_s, Y, [Allow_Intercept], [Include])` | Cook's distance per observation |
| `LOOCV_Prediction(X_s, Y, [Allow_Intercept], [Include])` | Leave-one-out predicted value per observation |
| `Observation_Number(Y, [Include])` | Sequential row indices 1 through n |
| `Rank_Fraction(Y, [Include])` | Empirical CDF fractions in original data order |
| `Y_Ranked(Y, [Include])` | Sorted filtered outcome values |
| `Normal_Scores(Y, [Include])` | Theoretical standard-normal quantiles |

### Prediction interval

`Prediction_Interval(X_s, Y, X_new, [Allow_Intercept], [Include], [alpha])` returns a 6-element vertical array: point estimate, SE of prediction, critical t value, lower bound, upper bound, and confidence level. `X_new` is a single-row range of predictor values for the new observation.

### Model significance — scalar outputs

These functions return the overall F-test results for the regression model.

| Function | Returns |
|---|---|
| `F_Statistic(X_s, Y, [Allow_Intercept], [Include])` | F-statistic for overall model significance |
| `F_Statistic_P_Value(X_s, Y, [Allow_Intercept], [Include])` | p-value for the overall F-test |

### Descriptive and correlation functions

These functions support exploratory analysis before model fitting. `Pearson_R`, `Spearman_R`, `Skewness`, and `Kurtosis` accept a multi-column `X_s` and return a k×1 array — one value per predictor. Comparing `Pearson_R` and `Spearman_R` for the same predictor diagnoses nonlinearity: a large gap between the two suggests curvature that a linear model will not capture.

| Function | Signature | Returns |
|---|---|---|
| `Pearson_R` | `(X_s, Y, [Include])` | k×1 vector of Pearson correlations between each predictor and Y |
| `Spearman_R` | `(X_s, Y, [Include])` | k×1 vector of Spearman rank correlations between each predictor and Y |
| `Skewness` | `(X_s, [Include])` | k×1 vector of skewness for each predictor column |
| `Kurtosis` | `(X_s, [Include])` | k×1 vector of excess kurtosis (Fisher convention) for each predictor column |
| `Correlation_Matrix` | `(X_s, [Include])` | k×k symmetric matrix of pairwise Pearson correlations among predictors |

### Utility functions

| Function | Returns |
|---|---|
| `Design_Matrix(X_s, [Allow_Intercept], [Include])` | Filtered numeric design matrix as a spilled array |
| `Gram_Inverse(X)` | Inverse of the Gram matrix (X'X)⁻¹ — used internally by `Hat_Diagonal`, `PRESS`, `LOOCV_Prediction`, and `Prediction_Interval` |
| `Complete_Cases_Filter(X_s, [Y])` | Boolean column — TRUE for rows with no missing values |
| `Column_Select(table, col_nums)` | Selected columns from an array in the specified order |
| `Leave_One_Out_Prediction(X_s, Y, n, [Allow_Intercept], [Include])` | Leave-one-out prediction for a single observation n |
| `This_Row(array)` | 1-to-n relative row indices |
| `Exclude_Row_N(array, n)` | Array with row n removed |
| `Dependent_Variable(Y, [Include])` | Filtered dependent variable vector |
| `Data_Completeness(predictor_row)` | TRUE if every cell in a row is numeric, FALSE if any value is blank, text, or error |

### Grid-search helpers and Univariate Weibull fitting

`Grid_Argument_Minimum(grid)` returns a horizontal three-value array:

```text
minimum value | 1-based row location | 1-based column location
```

Tied minima resolve to the first occurrence in row-major order. `Grid_Search_Optimum(grid)` uses those locations to return a vertical two-value array:

```text
best column-parameter value
best row-parameter value
```

For `Grid_Search_Optimum`, `grid` must be the complete rectangular Data Table body. Column-parameter values must sit immediately above it, and row-parameter values immediately to its left.

The **Univariate** sheet applies these helpers to a two-stage Weibull search. Shape is the column parameter and Scale is the row parameter. Each stage evaluates 20 values per axis—a 20×20 grid and 400 likelihood evaluations. The visible **Rows/Columns** value documents this generated physical size; manually changing it does not resize the Data Table.

Each axis includes both endpoints. Its spacing is:

```text
Step Size = (Max - Min) / (Rows/Columns - 1)
```

Stage 2 centers its narrower range on the Stage 1 optimum and extends one Stage 1 step in each direction. A red Best cell warns when the optimum lands on the corresponding grid boundary.

## Sample data and Regression sheet

`Lambda_Library.xlsx` includes the WHO Life Expectancy dataset (2,938 rows across 193 countries, 2000–2015) as a structured table on the **Life Expectancy Data** sheet. The workbook ships with eight sheets:

- **LAMBDA_functions** — browsable catalog of all function definitions with signatures, descriptions, and plain-language summaries, filterable by category and subcategory.
- **Life Expectancy Data** — the WHO dataset as a structured table.
- **Univariate Analysis** — descriptive statistics, three side-by-side histogram binning methods (Sturges, Scott, Freedman-Diaconis), and two-stage Weibull grid-search distribution fitting via native Data Tables.
- **Regression Instructions** — step-by-step guide for adapting the Regression sheet to a new dataset, including Name Manager updates and table setup.
- **Diagnostic Guide** — interpretation guide for regression diagnostics with Tier 1/Tier 2 plot specifications, threshold reference table, and "Common Patterns & Next Steps" guidance.
- **Version History** — changelog that travels with the workbook for non-git users.
- **Regression** — full ToolPak-style multiple regression analysis interface (see below).
- **Model Construction** — declarative model specification: assign each table column a Role (Response / Predictor / Identifier / Filter / Omit) and each predictor a Type (Continuous / Categorical, with reference-level control), and the sheet derives the row mask, row labels, and the constructed design matrix `X_s()` with level-qualified column names.

### Regression sheet

The **Regression** sheet uses the Life Expectancy dataset to demonstrate a full multiple regression analysis. It is organized in four zones:

- **Cols A–J — Main analysis.** Prediction inputs, regression statistics, ANOVA (including `F_Statistic` and `F_Statistic_P_Value`), prediction interval for a new observation, and coefficient table.
- **Predictor summary (below the coefficient table).** A per-predictor panel showing `Pearson_R`, `Spearman_R`, `Skewness`, `Kurtosis`, `VIF`, and `Tolerance` for every predictor — all filtered and recomputed whenever `Full_Data` changes.
- **Cols L–X — Residual output.** One row per filtered observation: predicted values, residuals, LOOCV predictions, leverage, studentized residuals, Cook's distance, and ranked/scored variants.
- **Diagnostic charts (col AC+).** Seven charts — Residuals vs Fitted, Normal Q-Q, Actual vs Predicted, Scale-Location (scatter), Cook's Distance, PRESS Residuals (bar), and Studentized Residuals vs Leverage (scatter) — that update automatically via OFFSET-based named ranges.

`Correlation_Matrix(X_s, [Include])` returns a k×k pairwise correlation matrix and can be entered in any blank cell on the sheet.

## Univariate Analysis functions

These functions power the **Univariate Analysis** sheet and can be used independently on any column of data.

### Descriptive statistics

| Function | Returns |
|---|---|
| `Missing_Count(data, [filter])` | Count of non-numeric (blank or text) cells in the active rows |
| `Descriptive_Statistics(data, [filter])` | 12×1 column vector: mean, median, mode, SD, variance, min, max, range, skewness, kurtosis, count, missing count |

### Histogram binning

| Function | Returns |
|---|---|
| `Sturges_Bins(data, [filter])` | Integer bin count via Sturges' rule |
| `Scott_Bins(data, [filter])` | Integer bin count via Scott's normal reference rule |
| `Freedman_Diaconis_Bins(data, [filter])` | Integer bin count via the Freedman-Diaconis rule |
| `Number_Of_Histogram_Bins(data, [method], [filter])` | Integer bin count for the chosen method (`"Sturges"`, `"Scott"`, or `"FD"`; defaults to `"FD"` when omitted) |
| `Bin_Edges(data, [method], [filter])` | k+1 × 1 column vector of full bin boundaries — the data minimum followed by k evenly-spaced upper edges |
| `Upper_Bin_Edges(data, [method], [filter])` | k × 1 upper bin edges (`DROP(Bin_Edges, 1)`) |
| `Bin_Lower_Edges(data, [method], [filter])` | k × 1 lower bin edges (`DROP(Bin_Edges, -1)`) |
| `Bin_Midpoints(data, [method], [filter])` | k × 1 bin midpoints — average of consecutive boundaries from `Bin_Edges` |
| `Bin_Counts(data, [method], [filter])` | k × 1 bin frequencies |

### CDF functions (for distribution fitting and GoF)

Each CDF function returns the probability of the interval `(minimum, maximum]`. When `minimum` is omitted it defaults to the distribution's lower bound.

| Function | Signature | Distribution |
|---|---|---|
| `CDF_Normal` | `(maximum, mean, sd, [minimum])` | Normal |
| `CDF_Lognormal` | `(maximum, meanlog, sdlog, [minimum])` | Lognormal |
| `CDF_Exponential` | `(maximum, rate, [minimum])` | Exponential |
| `CDF_Weibull` | `(maximum, shape, scale, [minimum])` | Weibull |
| `CDF_Gamma` | `(maximum, shape, rate, [minimum])` | Gamma |
| `CDF_Triangular` | `(maximum, min_val, mode, max_val, [minimum])` | Triangular |
| `CDF_Beta` | `(maximum, alpha, beta, [minimum])` | Beta (on [0, 1]) |
| `CDF_BetaPERT` | `(maximum, min_val, mode, max_val, [minimum])` | BetaPERT |

### NLL functions (for grid-search MLE)

Each NLL function returns the negative log-likelihood for the given distribution and parameters. Invalid parameter combinations (e.g., non-positive scale) return `1E+15` so the grid-search `MIN` stays finite.

| Function | Signature |
|---|---|
| `NLL_Normal` | `(data, mean, sd, [filter])` |
| `NLL_Lognormal` | `(data, meanlog, sdlog, [filter])` |
| `NLL_Exponential` | `(data, rate, [filter])` |
| `NLL_Weibull` | `(data, shape, scale, [filter])` |
| `NLL_Gamma` | `(data, shape, rate, [filter])` |
| `NLL_Triangular` | `(data, min_val, mode, max_val, [filter])` |
| `NLL_Beta` | `(data, alpha, beta, [filter])` |
| `NLL_BetaPERT` | `(data, min_val, mode, max_val, [filter])` |

### Goodness-of-fit functions

| Function | Signature | Returns |
|---|---|---|
| `GoF_AIC` | `(nll, k)` | AIC = 2k + 2·NLL |
| `GoF_BIC` | `(nll, k, n)` | BIC = k·ln(n) + 2·NLL |
| `GoF_Anderson_Darling` | `(data, dist_cdf, [include])` | A² = −n − (1/n)·Σᵢ(2i−1)[ln F(xᵢ) + ln(1−F(x_{n+1−i}))] (lower = better fit) |
| `GoF_Kolmogorov_Smirnov` | `(data, dist_cdf, [include])` | D = max\|F_n(xᵢ) − F(xᵢ)\| (lower = better fit) |

`GoF_Anderson_Darling` and `GoF_Kolmogorov_Smirnov` accept any of the `CDF_*` functions (partially applied via LAMBDA) as the `dist_cdf` argument.
