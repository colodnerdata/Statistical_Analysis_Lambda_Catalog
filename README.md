# Statistical Analysis Lambda Catalog

Excel 365 LAMBDA functions that replicate and extend Analysis ToolPak regression statistics — no VBA, no add-ins, no installation. Download `Lambda_Library.xlsx`, open it in Excel, and all functions are immediately available in formulas.

## Getting started

1. Download `Lambda_Library.xlsx` from this repository.
2. Open it in Excel 365 (Windows or Mac).
3. Enter your data in columns on any sheet, then call any function by name.

All functions are defined as workbook-scoped names, so they work in any cell formula within the workbook. The **Regression** and **Model Construction** sheets provide a ready-to-use analysis interface: declare each table column's Role and Type on the spec block, and the sheet derives the row mask, the constructed design matrix, and the full regression output automatically. The **Univariate Analysis** sheet demonstrates descriptive statistics, histogram binning, and distribution fitting via grid-search MLE. The **LAMBDA_functions** sheet is the canonical, always-in-sync function catalog — see [Documentation map](#documentation-map) below.

To use these functions in a different workbook, open both files in Excel at the same time. You can reference functions as `='[Lambda_Library.xlsx]'!FunctionName(args)`, or use Name Manager (Formulas → Name Manager → New) to copy individual definitions into your own workbook.

## What ships in the workbook

`Lambda_Library.xlsx` includes the WHO Life Expectancy dataset (2,938 rows across 193 countries, 2000–2015) as a structured table on the **Life Expectancy Data** sheet. The workbook ships with eight sheets:

- **LAMBDA_functions** — the **canonical function catalog**: every function's name, scope, full LAMBDA definition, arguments, yields, plain-language summary, and long description. This is the source of truth for "what functions exist and what they do" — generated from `lambda_functions.json` by `lambda_catalog/write_sheet_lambda_functions.py`, so the sheet and the source are never out of sync.
- **Life Expectancy Data** — the WHO dataset as a structured table.
- **Univariate Analysis** — descriptive statistics, three side-by-side histogram binning methods (Sturges, Scott, Freedman-Diaconis), and two-stage Weibull grid-search distribution fitting via native Data Tables. The fitted-distribution Q-Q plots and the histogram distribution overlays live alongside.
- **Model Construction** — declarative model specification: each table column gets a Role (Response / Predictor / Identifier / Filter / Omit), and the shipped workbook forward-wires Fixed Effects without exposing it in the Role dropdown yet. Each predictor also gets a Type (Continuous / Categorical, with reference-level control). The sheet derives the row mask, the constructed design matrix `X_s()`, and level-qualified column names.
- **Regression Instructions** — step-by-step guide for adapting the Regression sheet to a new dataset, including Name Manager updates and table setup.
- **Diagnostic Guide** — interpretation guide for regression diagnostics with Tier 1/Tier 2 plot specifications, threshold reference table, and "Common Patterns & Next Steps" guidance.
- **Version History** — changelog that travels with the workbook for non-git users. Every release's "Breaking? (yes/no)" flag is here so a workbook user gets the one signal the version number is *for* — "do my existing inputs still work?" — without the number also having to convey "how big is this release."

## Documentation map

This README is intentionally short — it tells you what the library is and how to open it. The rest of the documentation is split by purpose:

| File | What's in it | When to read it |
|---|---|---|
| [CONTRIBUTING.md](CONTRIBUTING.md) | Dev guide: setup, tests, build, file structure, adding a new LAMBDA function, Regression sheet chart conventions | You're contributing to the codebase |
| [ARCHITECTURE.md](ARCHITECTURE.md) | Foundational patterns: naming convention, function categories, the Role / Type / Sequence taxonomy, the Model Spec block (A–L), the data-transformation taxonomy, the reserved-spec-column pattern | You're adding a feature that has to honor the library's design |
| [ROADMAP.md](ROADMAP.md) | Version plan: the public-interface definition, the version ladder, what's shipped, what's next, ToolPak parity reference | You want to know what's planned or whether a change breaks the public interface |
| [DECISIONS.md](DECISIONS.md) | Resolved design decisions with their rationale, indexed by version, plus the supersession log and the alias-layer table | You want to know *why* a design choice was made |
| [TODOs.md](TODOs.md) | Active work only, with cross-links to DECISIONS for the "why" | You want to know what to work on next |
| [HUMAN_TEST_PLAN_v3_model_construction.md](HUMAN_TEST_PLAN_v3_model_construction.md) | Executed test plan for the v2.0 / v2.1 spec-driven Regression sheet (T0–T19) | You want to see the manual verification history for the spec block |

For the canonical function reference (every function's signature and description), open the **LAMBDA_functions** sheet in `Lambda_Library.xlsx`. It is generated from the same source as the codebase's catalog and is the only place the reference is always in sync with the workbook.
