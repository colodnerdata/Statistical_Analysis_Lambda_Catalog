# How the workbook works

## What you get is a spreadsheet

The analysis workbook is `dist/Lambda_Library.xlsx`. Its calculations use
Excel formulas, including LAMBDA and dynamic-array functions. Using the
workbook requires an Excel version that supports those functions; it does
not require Python, macros, or add-ins.

The repository's Python code builds and verifies the workbook.

## How the workbook uses Excel formulas

The following concepts explain how the formulas are organized.

### 1. LAMBDA and the Name Manager — Excel's own function language

Modern Excel lets you define a formula once, name it, and then call it by
name like a built-in function. Where `=SUM(A1:A10)` is Excel's function,
`Multiple_R(...)` is **this workbook's** function — a formula stored in the
Name Manager under that name, using Excel's `LAMBDA` feature. Its everyday
partner is `LET`, which names intermediate results *inside* one formula —
where a cell formula would need a helper column, `LET` keeps the whole
chain in one place. You will see `LAMBDA` and `LET` in nearly every
formula this workbook writes.

The workbook's LAMBDA
functions — the regression engine, the distribution fitters, the
diagnostic statistics — form a **catalog** of named functions, one row
each on the `LAMBDA_functions` tab, every one documented. The complete
list is on the {doc}`generated/lambda-reference` page of this site. Most
are *workbook-scoped* (they work on any sheet); a couple dozen are
*sheet-scoped*, living on the Regression sheet itself because they read
that sheet's layout.

### 2. Dynamic arrays — formulas that spill

In old Excel, one formula filled one cell. In modern Excel, a formula can
**spill**: entering `=A1:A10` in C1 fills C1:C10. The reference `C1#`
refers to that spilled result, using the address of the formula cell.

Spills allow computed ranges to resize when the source table or model
specification changes. The formulas use these functions:

| Function | What it does |
|---|---|
| `SEQUENCE(n)` | The numbers 1..n as a column |
| `TAKE(range, n)` | The first n rows of a range when n is positive |
| `MAP(range, LAMBDA(...))` | Evaluate a formula once per element |
| `BYROW(range, LAMBDA(...))` | Evaluate once per **row** of a 2-D range |
| `FILTER(range, condition)` | Keep the rows where the condition is TRUE |
| `VSTACK` / `HSTACK` | Glue ranges together vertically / horizontally |
| `Full_Factorial(...)` | A grid of every combination — the search engine's axes |
| `LET(name, value, ...)` | Name intermediate results inside one formula |
| `INDEX(range, r, c)` | Read one element out of a range |

### 3. Python as the builder, not the engine

The build scripts use `xlwings` to open Excel, write formulas and formatting,
copy template sheets, and save the workbook. Excel evaluates the formulas
when you use the resulting file.

Building with Python makes the workbook's formulas and layout reproducible
and allows the project's tests to check the result.

## The color grammar

The sheets use the following color conventions:

| Color | Meaning |
|---|---|
| **Light orange** | An input cell — this one is yours to type in |
| **Light blue** | A heading — structure, not data |
| **Red fill, dark red text** | A validation or diagnostic threshold has been crossed; inspect the associated value or message |
| **Yellow/amber fill** | A borderline warning — legal, but worth a look |
| No fill | Computed — the workbook owns this cell |

Red and yellow are **conditional formats**: the fill appears when a rule
fires. Some specification checks also display a message in row 2; for
example, the cell above the Role column reports multiple Response rows.

## One caution before you start

The Regression sheet computes its model **when Excel recalculates**. If
you set calculation to Manual (Formulas → Calculation Options), the
workbook will not update as you edit the spec — the shipped file saves in
Automatic, so leave it that way.

With that, you are ready for the tour: {doc}`generated/workbook-tour`.
