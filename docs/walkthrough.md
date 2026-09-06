# Connect and use the template sheets

This page describes the steps in a regression analysis: point
the sheet at data, declare a model, read the outputs, check the
diagnostics, make a prediction. The tab-by-tab inventory is on
{doc}`generated/workbook-tour`; the three built-in manual sheets have
their own generated pages — {doc}`generated/regression-instructions`
(the *how*), {doc}`generated/modeling-concepts` (the *why*), and
{doc}`generated/diagnostic-guide` (the *what to look for*).

## Step 1 — Connect the Regression sheet to your data

The Regression sheet reads its source table through the name
`Source_Table`. The header row, data body, and variable list in the
specification block derive from this name. To connect your data:

1. Put your data in an **Excel Table** (select the range with headers,
   press **Ctrl+T**, or Home → Format as Table). It must be a true
   Excel Table, not a plain named range — the `[#All]` structured
   reference in the example below is table syntax, and Tables carry
   their own headers and resize when rows are added. If you haven't
   made one before, Microsoft's guide covers both routes:
   <https://support.microsoft.com/en-us/excel/get-started/create-and-format-tables>
2. Open the **Name Manager** (Formulas → Name Manager), find
   `Source_Table` with scope **Regression**, and edit *Refers To* to your table including its
   header row — e.g. `=MyTable[#All]`.
3. Review the specification block, which displays one row per column of
   your table. Set the roles and other inputs for the new variables.

The shipped default is `LifeExpectancyData` (a curated four-driver
model of life expectancy); a second sample table, `MileageData`, ships
for a multi-level categorical example. You can use either sample table
or connect your own.

## Step 2 — Declare the model in MODEL SPECIFICATION (columns A–O)

The block at the left of the sheet has **one row per column of your
table**. Enter specifications in the orange cells. Rows 1–3 stay visible
as you scroll down. The column-by-column reference is on
{doc}`generated/spec-block`. Set each variable's Role and, for predictors,
its Include setting and Type. Other fields configure transformations,
interactions, and panel ordering:

- **Role (B)** — what this column *is*: `Response (y)` for the one
  variable being modeled (exactly one, or the status cell above the
  column says so in plain English), `Predictor (x)` for model inputs,
  `Identifier` for labels, `Filter` for TRUE/FALSE sample masks, `Omit`
  for unused helpers, `Fixed Effects` for a panel grouping.
- **Include (C)** — a per-row on/off switch, so you can try a predictor
  without deleting its specification.
- **Type (D)** — `Continuous` (enters the fit as-is) or `Categorical`
  (dummy-coded automatically, one 0/1 column per level except the
  reference; type a different reference in **E** if you want).
- **Transform (G)** — `Log` fits in log space (for percentage effects
  and skewed data); `Log (drop ≤ 0)` also excludes zeros and negatives
  and tells you how many.
- **Interaction Term / Operation (M/N)** — pick another predictor and
  one of Product / Difference / Ratio to add an interaction column;
  naming the row's *own* variable under Product is the documented way
  to add a quadratic (x²). The interaction is added to the design matrix.
- **Sequence (H) + Sequence Period (I)** — for panel data, mark the
  ordering axis and (optionally) type a period step Δ; lag and
  difference features follow it within each group.

Two displays help you review the specification. First, **the status cells**:
the cells in row 2 are blank when their checks pass, and
when a check fails, the relevant cell displays a message in plain
English — *no Response row*, *multiple Fixed Effects*, *this Log column
contains N values ≤ 0*, *the declared reference level is not in the
sample*. Check row 2 after editing the model. Second, **the Design Columns audit
(column O)**: each row shows how many design-matrix columns it
contributes, and the total above the block (with the intercept) is the
full width of the model — it turns amber when a model gets slow and red
when it gets too wide to fit.

## Step 3 — Read the outputs

Right of the spec block, each zone holds one part of the answer:

- **REGRESSION OUTPUTS** — R² and Adjusted R², the
  ANOVA F-test, the coefficient table (estimate, standard error,
  t-statistic, P-value, confidence bounds) with non-significant terms
  flagged red.
- **PREDICTOR SUMMARY** — each term's partial R², tolerance and GVIF
  (a collinearity check).
- **PREDICTION OUTPUTS** — prediction inputs are prefilled with each
  constructed column's **Training Mean**; type
  scenario values over them to ask "what if?" The **Original Units**
  column back-transforms log-space predictions to real units.
- **RESIDUAL OUTPUT** — per-observation fitted values, residuals,
  studentized residuals, leverage, Cook's Distance and PRESS residuals,
  one row per observation, feeding the diagnostic charts.

### Back-transformation: Duan vs. Naive

When the response is Log-transformed, the fit runs in log space. The
**Unit-Space Fit** block (AG4:AH10) reports R², Adjusted R², and RMSE in
original units using the method selected in **AH5**:

- **Duan** (default) multiplies EXP(predicted log response) by a
  *smearing factor*: the average of EXP(residuals).
- **Naive** uses EXP(predicted log response) without that adjustment.

The toggle also changes the original-unit point prediction and residual
columns. The interval bounds always use the Naive back-transformation.

## Step 4 — Judge the model with the charts

Seven charts sit right of the residual columns, and each one's **title
carries the live statistic it judges against** — the Q-Q correlation, the
Adjusted R², the mean leverage, the Cook's Distance cutoff, the PRESS
total — so a chart can be read without scrolling back to its source
cell. The {doc}`generated/diagnostic-guide` page (the workbook's own
Diagnostic Guide sheet) provides further guidance. The main checks are:

1. **Residuals vs. Fitted**: look for scatter around zero; a curve
   suggests nonlinearity, and a funnel suggests non-constant variance.
2. **Normal Q-Q**: look for departures from the diagonal; the title warns
   "— check normality" when the correlation drops below 0.95.
3. **Actual vs. Predicted**: compare the points with the 45° line; the title
   carries the live Adjusted R².
4. **Cook's Distance / leverage / PRESS**: the influence checks for
   Tier 2 — individual rows that unduly drive the fit, with the cutoff
   printed in the title and flagged bars labeled.

Use these plots to look for problems with the model. An apparently clean
plot does not establish that all model assumptions hold. Influence flags
identify observations to inspect; an extreme value alone is not a reason
to delete an observation.

## Step 5 — Analyze a column on the Univariate sheet

To connect Univariate to your own data, put the column in an Excel Table
in the same workbook. For a table named `MyTable` with a column named
`Measurement`, replace the formula in **A4** with:

```excel
=IF(MyTable[Measurement]="","",MyTable[Measurement])
```

Replace the companion numeric-inclusion formula in **B4** with:

```excel
=ISNUMBER(MyTable[Measurement])
```

Both formulas must read
the same column; changing Regression's `Source_Table` does not change
Univariate's source. Keep the cells below the two formula anchors clear
so the results can spill.

The **Univariate** sheet analyzes a single data column. It reports descriptive
statistics, a histogram (with three bin-width rules — Sturges, Scott,
Freedman-Diaconis — for comparing bin choices), and a
**distribution-fitting comparison**: Weibull, Gamma and Beta fits are
found by a two-stage grid search (a wide first pass, then a refined
second pass around the best point — the Weibull and Gamma passes are drawn on
profile-likelihood charts), with the best-fitting distribution's
parameters highlighted, and Q-Q plots for every candidate against your
data. You can use this sheet to inspect individual variables before fitting a
regression.

Each fit's **Grid Points** cell is a live input — raise it for a finer
search and the grid, its NLL column and the charts all grow with it.

## Where to go next

- The **spec-block reference** — {doc}`generated/spec-block` — one row
  per column of the model specification.
- The **formula review** — {doc}`generated/formula-review` — the actual
  formulas behind all of this, each annotated in plain English.
- The **LAMBDA reference** — {doc}`generated/lambda-reference` — every
  catalog function, its arguments, and what it returns.