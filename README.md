# Statistical Analysis Lambda Catalog

Excel 365 LAMBDA functions that replicate and extend Analysis ToolPak regression statistics — no VBA, no add-ins, no installation. Download a workbook, open it in Excel, and all functions are immediately available in formulas.

## Which workbook do I want?

From v3.0 the build emits **two workbooks**. Both carry the **complete function library** — all 141 LAMBDA definitions, identical in each — so whichever you pick, every function in the catalog is available in your formulas. They differ only in which pre-built analysis sheets they contain.

| Workbook | Contains | Pick this if |
|---|---|---|
| **`Lambda_Library.xlsx`** | The Regression workbench, the three sample datasets, and the reference sheets | You are fitting models — regression, fixed effects, prediction, diagnostics. **This is the default.** |
| **`Lambda_Library_Univariate.xlsx`** | The Univariate Analysis sheet — descriptive statistics, histogram binning, distribution fitting | You are characterizing a single variable's distribution, or fitting a distribution for cost/risk work |

Nothing is lost by choosing one: the function library is the same in both, and you can open both at once if you want both sets of sheets.

## Getting started

1. Download the workbook you want (see above; `Lambda_Library.xlsx` if unsure).
2. Open it in Excel 365 (Windows or Mac).
3. Enter your data in columns on any sheet, then call any function by name.

Most functions are defined as workbook-scoped names, so they work in any cell formula within either workbook. The **Regression** sheet (Regression workbook) also installs a small set of sheet-scoped constructor names for the spec block and provides a ready-to-use analysis interface: declare each table column's Role and Type on the spec block, and the sheet derives the row mask, the constructed design matrix, and the full regression output automatically. The **Univariate Analysis** sheet (Univariate workbook) demonstrates descriptive statistics, histogram binning, and distribution fitting via search-based MLE. The **LAMBDA_functions** sheet ships in both and is the canonical, always-in-sync function catalog — see [Documentation map](#documentation-map) below.

Around 30 catalog functions are called by no pre-built sheet. That is deliberate: they are the **standalone user-callable layer** — `Correlation_Matrix`, `Lag_By`, `Descriptive_Statistics`, `Design_Matrix` and others you call in your own cells on your own data. The pre-built sheets demonstrate the library; they are not the whole of it.

To use these functions in a different workbook, you have two options. **The easy one:** if you are using one of the pre-built sheets (the Regression workbench or the Univariate Analysis sheet), copy the sheet into your own workbook. The sheet's named-range dependencies come with it — the **workbook-scoped LAMBDA definitions** (the 123 portable functions) and the **sheet-scoped definitions** (the Regression sheet's 18 constructor closures like `Predictor_Columns`, `Sample_Include`, `Design_Columns`; the Univariate sheet's `UV_Data`, `UV_Include`, `GoF_AIC`) travel inside the sheet-copy and are renamed automatically. Open your workbook, the sheet calculates, and every function is ready in formulas. **The other one:** if you only want a single function, or you would rather not pull in the pre-built sheet at all, open both files in Excel at the same time and reference functions as `='[Lambda_Library.xlsx]'!FunctionName(args)`, or use Name Manager (Formulas → Name Manager → New) to copy individual definitions into your own workbook.

## Versions

Two numbers, because there are two workbooks:

- **Function library version** — the shared catalog of 141 LAMBDA definitions, identical in both workbooks. Moves when a function is added, renamed, or changes what it returns.
- **Workbook version** — one per artifact, covering that workbook's sheets, input cells, and control blocks. Moves when its input surface changes.

Each workbook's **Version History** sheet shows both, with its own workbook version as the headline:

```
Regression Workbook 3.3.0   ·   Function Library 3.3.0
Univariate Workbook 2.0.0   ·   Function Library 3.3.0
```

The **`Breaking?` flag belongs to the workbook version**, since it answers a question about your saved inputs. A library-version bump that adds a function breaks nothing. A change to the Univariate workbook's inputs does not move the number a Regression user reads.

## What ships in the workbooks

### `Lambda_Library.xlsx` — the Regression workbook

Includes the WHO Life Expectancy dataset (2,938 rows across 193 countries, 2000–2015) as a structured table on the **Life Expectancy Data** sheet. Eight sheets:

- **LAMBDA_functions** — the **canonical function catalog**: every function's name, scope, full LAMBDA definition, arguments, yields, plain-language summary, and long description. This is the source of truth for "what functions exist and what they do" — generated from `lambda_functions.json` by `lambda_catalog/write_sheet_lambda_functions.py`, so the sheet and the source are never out of sync.
- **Life Expectancy Data** — the WHO dataset as a structured table. It ships as a ready-made target for practicing the Regression sheet's `Source_Table` retarget: point `Source_Table` at `LifeExpectancyData[#All]` in Name Manager and the spec block re-populates from its columns, no data of your own required.
- **Mileage Data** — the Auto MPG dataset (406 vehicles) as a structured table. This is the dataset the Regression sheet's `Source_Table` targets by default (`=MileageData[#All]`).
- **Production Lots** — a small unbalanced learning-curve panel (3 facilities, 51 lots) as a structured table. The only shipped dataset with a natural Fixed Effects grouping column (Facility) and Sequence column (Fiscal_Year) — retarget `Source_Table` at `ProductionLotsData[#All]` (or pass `--regression-dataset production_lots` to `build_production.py`) for a ready-made Fixed Effects example.
- **Regression** — the spec-driven regression workbench, and the reason this workbook exists. Each table column gets a Role (`Response (y)` / `Predictor (x)` / `Identifier (Row Label)` / `Filter` / `Omit` / `Fixed Effects`). Each predictor also gets a Type (Continuous / Categorical, with reference-level control), an optional Transform (natural log, on the Response row and/or Continuous Predictor rows), and an optional interaction with another predictor (Product, Difference, or Ratio — a row pointing at itself under Product is how you write a quadratic term). Any column can carry the structural Sequence flag for lag/difference/serial-correlation features. The spec block derives the row mask, the constructed design matrix, and level-qualified column names; the sheet then produces the full regression output, diagnostics, and — for a Fixed Effects model — both a mean-response CI and a new-observation PI in the prediction outputs.

  The sheet is laid out as **five content zones** — Model Specification, Predictor Summary, Regression Outputs, Prediction Outputs, and Residual Output — each in its own collapsible outline group, separated by narrow gap columns, with seven pre-built diagnostic charts to the right. Past the charts sits a band of **materialized** blocks — the Model Context, the row mask, and the Constructed Design Matrix — which exist so the engines read a computed value instead of recomputing it at every call site. Only the Model Context block collapses; the row mask and the design matrix stay expanded, because hiding the columns a spilled array occupies is what leaves it stale when the model recalculates.
- **Regression Instructions** — step-by-step guide for adapting the Regression sheet to a new dataset, including Name Manager updates and table setup.
- **Diagnostic Guide** — interpretation guide for regression diagnostics with Tier 1/Tier 2 plot specifications, threshold reference table, and "Common Patterns & Next Steps" guidance.
- **Version History** — changelog that travels with the workbook for non-git users. Every release's "Breaking? (yes/no)" flag is here so a workbook user gets the one signal the version number is *for* — "do my existing inputs still work?" — without the number also having to convey "how big is this release."

### `Lambda_Library_Univariate.xlsx` — the Univariate workbook

The same **LAMBDA_functions** catalog and **Version History** sheets, plus:

- **Univariate Analysis** — descriptive statistics, three side-by-side histogram binning methods (Sturges, Scott, Freedman-Diaconis), and two-stage distribution fitting across eight candidate distributions. Weibull and Gamma profile their scale/rate parameter out in closed form and search a 20-point profile-NLL curve per stage; Beta searches both parameters on a `Full_Factorial` dynamic-array grid (N² negative-log-likelihood evaluations per stage, N editable live in the sheet); Normal, Log-Normal, Exponential, Triangular, and BetaPERT are closed-form. The fitted-distribution Q-Q plots, the Weibull and Gamma profile-NLL line charts, and the histogram distribution overlays live alongside.

This workbook ships in full **Automatic** calculation mode — no Data Tables anywhere — so fitted parameters update live as you change the input column or the Beta grid-points cell.

## Documentation map

This README is intentionally short — it tells you what the library is and how to open it. The rest of the documentation is split by purpose:

| File | What's in it | When to read it |
|---|---|---|
| [CONTRIBUTING.md](CONTRIBUTING.md) | Dev guide: setup, tests, build, file structure, adding a new LAMBDA function, Regression sheet chart conventions | You're contributing to the codebase |
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | Foundational patterns: naming convention, function categories, the Role / Type / Sequence taxonomy, the Model Spec block (A–O), the data-transformation taxonomy, the reserved-spec-column pattern | You're adding a feature that has to honor the library's design |
| [ROADMAP.md](docs/ROADMAP.md) | Version plan: the public-interface definition, the version ladder, what's shipped, what's next, ToolPak parity reference | You want to know what's planned or whether a change breaks the public interface |
| [DECISIONS.md](docs/DECISIONS.md) | Resolved design decisions with their rationale, indexed by version, plus the supersession log | You want to know *why* a design choice was made |
| [MODEL_TESTING_ASSETS.md](docs/MODEL_TESTING_ASSETS.md) | The regression test-model suite: which model configurations the QC harness covers, which corner each one is there for, the datasets future milestones need, and the ordering the version ladder follows from v3.4 on (Regression work first, then test-scale growth) | You're adding or changing a QC model case, wiring a new dataset, or wondering why the roadmap is sequenced the way it is |
| [TODOs.md](docs/TODOs.md) | Open work only — every item tagged status · size · whether it needs Excel, with a "Pick something to work on" index at the top and cross-links to DECISIONS for the "why" | You want to pick up the next piece of work |

For the canonical function reference (every function's signature and description), open the **LAMBDA_functions** sheet in `Lambda_Library.xlsx`. It is generated from the same source as the codebase's catalog and is the only place the reference is always in sync with the workbook.
