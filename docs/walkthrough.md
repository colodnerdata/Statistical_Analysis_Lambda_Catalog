# Part 1 — A walk through the workbook

This page walks one full analysis, the way you would actually do it: point
the sheet at data, declare a model, read the outputs, check the
diagnostics, make a prediction. The tab-by-tab inventory is on
{doc}`generated/workbook-tour`; the three built-in manual sheets have
their own generated pages — {doc}`generated/regression-instructions`
(the *how*), {doc}`generated/modeling-concepts` (the *why*), and
{doc}`generated/diagnostic-guide` (the *what to look for*).

## Step 1 — Point the sheet at your data: one edit

The Regression sheet does not know or care what data it ships with. It
reads **one name**: `Source_Table`. Everything else — the header row, the
data body, the variable list in the specification block — derives from
it. So retargeting the sheet to your own data is a single edit:

1. Put your data in an **Excel Table** (select the range with headers,
   press **Ctrl+T**, or Home → Format as Table). Tables carry their own
   headers and resize when rows are added.
2. Open the **Name Manager** (Formulas → Name Manager), find
   `Source_Table`, and edit *Refers To* to your data including its
   header row — e.g. `=MyTable[#All]`.
3. Done. The specification block below gains one row per column of your
   table, and every dropdown, band and output resizes with it.

A structured Table is the convenient form, not a requirement: a **Named
Range** works just as well. `Source_Table` can point at any rectangular
range whose first row is headers — `=Data!$A$1:$F$100` — because the
sheet reads it positionally: `Header_Names` takes the first row, and
`Source_Data` is everything below it. The one trade-off to know: a Table
grows by itself when you append rows, a fixed range does not, so point
the name at the new extent after adding data.

The shipped default is `LifeExpectancyData` (a curated four-driver
model of life expectancy); a second sample table, `MileageData`, ships
for a multi-level categorical demo. Practice the retarget on either
before pointing at your own data.

## Step 2 — Declare the model in MODEL SPECIFICATION (columns A–O)

The block at the left of the sheet (frozen, so it stays visible while you
scroll) is the model's control panel: **one row per column of your
table**, and you fill in the orange cells. The full column-by-column
reference is on {doc}`generated/spec-block`; the short version. The three
fields that actually have to be set — for every row you want in the
model — are the **main mandatory fields: Role (B), Include (C) and Type
(D)**; everything after them is an optional refinement with a safe
default:

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
  to add a quadratic (x²). These columns are fully wired: the
  interaction is built into the design matrix automatically.
- **Sequence (H) + Sequence Period (I)** — for panel data, mark the
  ordering axis and (optionally) type a period step Δ; lag and
  difference features follow it within each group.

Two ideas carry the whole block. First, **the status cells**: above the
block, the cells in row 2 are quiet when the specification is legal, and
when it is not, the relevant one states its own problem in plain
English — *no Response row*, *multiple Fixed Effects*, *this Log column
contains N values ≤ 0*, *the declared reference level is not in the
sample*. Look at row 2 first, always. Second, **the Design Columns audit
(column O)**: each row shows how many design-matrix columns it
contributes, and the total above the block (with the intercept) is the
full width of the model — it turns amber when a model gets slow and red
when it gets too wide to fit.

## Step 3 — Read the outputs

Right of the spec block, each zone holds one part of the answer:

- **REGRESSION OUTPUTS** — the classic table: R² and Adjusted R², the
  ANOVA F-test, the coefficient table (estimate, standard error,
  t-statistic, P-value, confidence bounds) with non-significant terms
  flagged red, and the **Predictor Summary** (each term's partial R²,
  tolerance and GVIF — a collinearity check).
- **PREDICTION OUTPUTS** — prefilled with each predictor's **Training
  Mean** so it always shows a valid prediction out of the box; type
  scenario values over them to ask "what if?" The **Original Units**
  column back-transforms log-space predictions to real units.
- **RESIDUAL OUTPUT** — per-observation fitted values, residuals,
  studentized residuals, leverage, Cook's Distance and PRESS residuals,
  one row per observation, feeding the diagnostic charts.

### The one subtlety worth understanding: Duan vs. Naive

When the response is Log-transformed, the fit runs in log space, and
`EXP(predicted)` is the **median** of the response — not the mean.
That is not a technicality: for skewed data the mean and median differ
materially. The **Unit-Space Fit** block (AG4:AH10) reports the
back-transformed R² / Adj R² / RMSE both ways, on the **AH5 toggle**:

- **Duan** (default) multiplies by a *smearing factor* — the average of
  EXP(residuals) — recovering the conditional **mean**;
- **Naive** is plain EXP — the conditional **median**.

If you need "average value given these inputs" (forecasting a total),
use Duan; if you need the typical case, Naive. The toggle re-points the
prediction bounds and the original-units residual columns at the same
time, so nothing can be read against the wrong basis.

## Step 4 — Judge the model with the charts

Seven charts sit right of the residual columns, and each one's **title
carries the live statistic it judges against** — the Q-Q correlation, the
Adjusted R², the mean leverage, the Cook's Distance cutoff, the PRESS
total — so a chart can be read without scrolling back to its source
cell. The {doc}`generated/diagnostic-guide` page (the workbook's own
Diagnostic Guide sheet) is the full curriculum; the 30-second version:

1. **Residuals vs. Fitted**: random scatter is good; a curve means
   nonlinearity, a funnel means non-constant variance.
2. **Normal Q-Q**: points on the diagonal are good; the title warns
   "— check normality" when the correlation drops below 0.95.
3. **Actual vs. Predicted**: hugging the 45° line is good; the title
   carries the live Adjusted R².
4. **Cook's Distance / leverage / PRESS**: the influence checks for
   Tier 2 — individual rows that unduly drive the fit, with the cutoff
   printed in the title and flagged bars labeled.

A clean Tier 1 usually means you can trust the coefficients and their
P-values; a flagged Tier 2 plot means one or a few rows deserve
inspection — the guide is explicit that an extreme value is a reason to
*investigate*, never a license to delete data.

## Step 5 — The Univariate sheet: one column, understood

The **Univariate** sheet runs the same philosophy on a single data
column: point it at a column of values and it reports descriptive
statistics, a histogram (with three bin-width rules — Sturges, Scott,
Freedman-Diaconis — so the binning can't drive the conclusion), and a
**distribution-fitting comparison**: Weibull, Gamma and Beta fits are
found by a two-stage grid search (a wide first pass, then a refined
second pass around the best point — the two passes are both drawn on the
profile-likelihood charts), with the best-fitting distribution's
parameters highlighted, and Q-Q plots for every candidate against your
data. Use it before the regression: know what each variable looks like
on its own.

Each fit's **Grid Points** cell is a live input — raise it for a finer
search and the grid, its NLL column and the charts all grow with it.

## Where to go next

- The **spec-block reference** — {doc}`generated/spec-block` — one row
  per column of the model specification.
- The **formula review** — {doc}`generated/formula-review` — the actual
  formulas behind all of this, each annotated in plain English.
- The **LAMBDA reference** — {doc}`generated/lambda-reference` — every
  catalog function, its arguments, and what it returns.