# Worked example: life expectancy

This example uses the supplied **Life Expectancy Data** table to explain
model setup, coefficients, diagnostics, and predictions. It is an
observational country/year dataset; use the example to learn the controls,
not to infer the effect of changing a country's policy or behavior.

## 1. Set up the model

Open a working copy of the workbook and select **Regression**. In
**Formulas → Name Manager**, select `Source_Table` with scope **Regression**
and check that **Refers To** is `=LifeExpectancyData[#All]`.

A response is the variable being modeled; predictors are the variables
used to explain its variation. In **MODEL SPECIFICATION**, use:

| Variable | Role | Include | Type |
|---|---|---|---|
| Life expectancy | Response (y) | Leave as supplied | Not used for this role |
| Status | Predictor (x) | TRUE | Categorical |
| Adult Mortality | Predictor (x) | TRUE | Continuous |
| Alcohol | Predictor (x) | TRUE | Continuous |
| percentage expenditure | Predictor (x) | TRUE | Continuous |
| Country | Identifier (Row Label) | Leave as supplied | Not used for this role |
| Year | Omit | Leave as supplied | Not used for this role |
| All remaining variables | Omit | Leave as supplied | Not used for this role |

Leave **Intercept** TRUE, **Alpha** at 0.05, transforms at **None**, and
interaction fields blank. Leave Status's **Reference Level** blank;
**Reference In Use** should show **Developed**. The supplied Year Sequence
flag can remain TRUE; it does not add a predictor or a country fixed effect.

A categorical predictor is represented by indicator columns. Here
**Status: Developing** is 1 for Developing and 0 for the reference level,
Developed. The model has four predictor columns plus an intercept, so
**Σ Design Columns** should be **5**.

## 2. Check the results

These rounded reference values were calculated from the repository's
supplied data using its independent regression oracle (`life_talk_demo`).
They describe this specification and sample, not every model of these data.

| Output | Expected value |
|---|---:|
| Observations | 2,735 |
| R² | 0.5885 |
| Adjusted R² | 0.5879 |
| Standard Error of Regression | 6.1313 years |
| Intercept coefficient | 77.2036 |
| Status: Developing coefficient | −3.6081 |
| Adult Mortality coefficient | −0.04434 |
| Alcohol coefficient | 0.37786 |
| percentage expenditure coefficient | 0.0006101 |

R² describes the proportion of response variation explained within this
sample. It is not a measure of performance on new data.

For **Adult Mortality**, a difference of 10 units corresponds to a fitted
life-expectancy difference of about **−0.4434 years**, holding the other
predictors fixed. The Status coefficient compares Developing with
Developed at the same values of the continuous predictors. These are
model associations, not estimates of causal effects.

If the observation count differs, check the source table, roles, Include
settings, filters, and transforms before comparing coefficients.

## 3. Change one setting and inspect the sample

Set **Include = FALSE** on the Alcohol row. Its coefficient disappears,
and Σ Design Columns becomes **4**. You should see:

| Output | Alcohol included | Alcohol excluded |
|---|---:|---:|
| Observations | 2,735 | 2,928 |
| R² | 0.5885 | 0.5749 |
| Adjusted R² | 0.5879 | 0.5745 |
| Adult Mortality coefficient | −0.04434 | −0.04520 |

The observation count rises because rows missing Alcohol no longer fail
that predictor's numeric-completeness check. Both the model and its sample
changed, so the R² difference does not isolate Alcohol's contribution.
For a comparison on the same rows, use a Filter column to retain a common
sample. Restore **Include = TRUE** before continuing.

## 4. Review the diagnostics

A residual is the observed response minus the model's prediction. Inspect
**Residuals vs. Fitted** for curves or changes in spread, then inspect
**Normal Q-Q**, which compares ordered residuals with a normal distribution.
For this model, **QQ Correlation** is about **0.9327**, below the workbook's
0.95 warning threshold. Inspect the plot rather than treating this number
as a formal normality test.

Also review **Cook's Distance** and **Studentized Residuals vs. Leverage**
for influential observations. These country/year rows include repeated
observations within countries; ordinary regression standard errors do not
automatically account for that dependence. A chart alone cannot establish
independence. See {doc}`generated/diagnostic-guide` for the checks and their
limits.

## 5. Make a prediction

The orange **PREDICTION INPUTS** cells correspond to constructed columns.
With the supplied training-mean defaults, the point estimate is about
**69.1689 years**. At Alpha = 0.05, the mean-response confidence interval
is approximately **68.9390–69.3988**, and the new-observation prediction
interval is **57.1443–81.1935**. These intervals use the fitted model's
assumptions.

The default Status indicator is a sample proportion, so this prediction
represents the center of the design data, not a particular country's category.
For a specific scenario, enter **1** for Status: Developing or **0** for
Developed and supply the three continuous values in their source units.

For example, enter **1**, **150**, **5**, and **500** for Status: Developing,
Adult Mortality, Alcohol, and percentage expenditure, respectively. The
point prediction should be about **69.1387 years**. Raising Adult Mortality
to **160**, with the other inputs unchanged, lowers it to about **68.6953**.
These are illustrative inputs, not a forecast for a named country.

A confidence interval describes uncertainty about the model's mean
response at those inputs. A prediction interval also includes variation
of an individual observation and is therefore wider.

## Use your own data

Continue with {doc}`walkthrough` to connect either template sheet to your
chosen data source. Changing a source name updates the variable list;
review and reset the model specifications for the new columns.
