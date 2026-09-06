# Part 0 — What this is, in plain terms

## What you get is a spreadsheet

Everything this project produces is **one Excel file**:
`dist/Lambda_Library.xlsx`. Open it, and you are looking at the whole
thing. There is no app, no installer, and — importantly — **no macros**,
so it opens without any security warnings and works in Excel for Windows
and Mac alike (the workbook itself needs no Python; the formulas use
functions built into modern Excel).

The rest of this repository — all the Python code — is the **factory**
that builds that one file. You never run Python to *use* the workbook, only
to rebuild it.

## Three technologies, from zero

If you have used Excel's `SUM`, you are three ideas away
from understanding every formula in this workbook:

### 1. LAMBDA and the Name Manager — Excel's own function language

Modern Excel lets you define a formula once, name it, and then call it by
name like a built-in function. Where `=SUM(A1:A10)` is Excel's function,
`Multiple_R(...)` is **this workbook's** function — a formula stored in the
Name Manager under that name, using Excel's `LAMBDA` feature. Its everyday
partner is `LET`, which names intermediate results *inside* one formula —
where a cell formula would need a helper column, `LET` keeps the whole
chain in one place. You will see `LAMBDA` and `LET` in nearly every
formula this workbook writes.

This project takes that idea further than most: the workbook's LAMBDA
functions — the regression engine, the distribution fitters, the
diagnostic statistics — form a **catalog** of named functions, one row
each on the `LAMBDA_functions` tab, every one documented. The complete
list is on the {doc}`generated/lambda-reference` page of this site. Most
are *workbook-scoped* (they work on any sheet); a couple dozen are
*sheet-scoped*, living on the Regression sheet itself because they read
that sheet's layout.

### 2. Dynamic arrays — formulas that spill

In old Excel, one formula filled one cell. In modern Excel, a formula can
**spill**: `=A1:A10` written in one cell fills ten cells downward with a
blue outline around the result. That spilled range is referred to as
`A1#` — "whatever that spill currently is."

This workbook leans on spills everywhere, for one reason: **the model
resizes**. Point the sheet at a table with 12 columns and every
specification row, design-matrix column and output resizes itself. The
helper vocabulary you will see over and over:

| Function | What it does |
|---|---|
| `SEQUENCE(n)` | The numbers 1..n as a column |
| `TAKE(range, n)` | The first n rows of a range — non-volatile, so cheap |
| `MAP(range, LAMBDA(...))` | Evaluate a formula once per element |
| `BYROW(range, LAMBDA(...))` | Evaluate once per **row** of a 2-D range |
| `FILTER(range, condition)` | Keep the rows where the condition is TRUE |
| `VSTACK` / `HSTACK` | Glue ranges together vertically / horizontally |
| `Full_Factorial(...)` | A grid of every combination — the search engine's axes |
| `LET(name, value, ...)` | Name intermediate results inside one formula |
| `INDEX(range, r, c)` | Read one element out of a range |

### 3. Python as the builder, not the engine

The Python in this repository is a **factory robot**: it opens Excel
through a library called `xlwings`, writes every sheet, formula, chart,
named range and color into a new workbook, saves it, and exits. Its
output is a normal file you can email to a colleague; from then on
Excel alone does all the math. (The workbook needs no Python; it needs
only the Excel you already have.)

Why build it with Python at all? Because writing 150+ interlocking
formulas by hand is error-prone — the factory writes them the same way
every time, and the project's tests can verify the result.

## The color grammar

Every sheet follows one convention — learn it once, read any sheet:

| Color | Meaning |
|---|---|
| **Light orange** | An input cell — this one is yours to type in |
| **Light blue** | A heading — structure, not data |
| **Red fill, dark red text** | Something is wrong; read the message and fix it |
| **Yellow/amber fill** | A borderline warning — legal, but worth a look |
| No fill | Computed — the workbook owns this cell |

Red and yellow are **conditional formats**: the fill appears only when a
rule fires, and the matching status cell near the top of the sheet states
the problem in plain English (for example, the cell above the Role column
says *"ERROR: multiple Response (y) rows — mark exactly one."*). Nothing
fails silently.

## One caution before you start

The Regression sheet computes its model **when Excel recalculates**. If
you set calculation to Manual (Formulas → Calculation Options), the
workbook will not update as you edit the spec — the shipped file saves in
Automatic, so leave it that way.

With that, you are ready for the tour: {doc}`generated/workbook-tour`.