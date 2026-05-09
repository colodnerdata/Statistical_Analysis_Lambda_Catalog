# Statistical Analysis Lambda Catalog

Reproduce the functionality of the Analysis Toolpak using named-object LAMBDA functions
registered directly in an Excel workbook's Name Manager.

## Requirements

- Python 3.10+
- Desktop Excel (Windows or macOS) — xlwings drives Excel through COM/AppleScript
- Dependencies declared in `pyproject.toml`; install with:

  ```
  pip install xlwings
  ```

  or, if using [uv](https://github.com/astral-sh/uv):

  ```
  uv sync
  ```

## Building the workbook

```
python build_workbook.py
```

This script:

1. Reads all function definitions from `lambda_functions.json`.
2. Opens `Lambda_Library.xlsx` if it already exists, or creates it from scratch.
3. Ensures the workbook has a starter sheet named **MLR** (required by `format_workbook.py`).
4. Syncs every function into the workbook's Name Manager as a workbook-level name using the
   compact LAMBDA formula string.  Existing names are overwritten so the script is safe to
   re-run.
5. Prints a short summary: workbook path, created count, updated count, invalid count.

**Note:** `build_workbook.py` depends on desktop Excel being installed.  It opens Excel
invisibly in the background, writes the names, saves, and closes — no manual interaction
is needed.

## Files

| File | Purpose |
|---|---|
| `build_workbook.py` | Primary entry point — builds/updates `Lambda_Library.xlsx` |
| `lambda_functions.json` | Source of truth for function names and LAMBDA formulas |
| `format_workbook.py` | Visual formatting pass (run after `build_workbook.py` when ready) |
| `Lambda_Library.xlsx` | Output workbook (created by `build_workbook.py`) |

## Functions registered

| Name | Returns | Notes |
|---|---|---|
| `Observations` | n — row count after filtering | |
| `DF_Regression` | k — number of predictors | |
| `DF_Total` | n−1 (with intercept) or n (without) | |
| `DF_Residual` | n−k−1 (with intercept) or n−k (without) | |
| `R2` | Coefficient of determination | See note below |
| `Multiple_R` | Square root of R² | |
| `Adjusted_R2` | R² penalised for predictor count | |

After running `build_workbook.py`, open `Lambda_Library.xlsx` in Excel and press
**Ctrl+F3** to verify the names appear in Name Manager as workbook-level definitions.

### Known limitation: `R2`

Excel's Name Manager rejects `R2` via the COM API because it conflicts with R1C1 cell
reference notation (`R2` = row 2).  `build_workbook.py` reports this as a `Failed: 1`
entry and continues — the other 6 names are registered successfully.

**Workaround:** Add `R2` manually through the Name Manager UI (Formulas → Name Manager →
New → Name: `R2`, Refers to: paste the compact formula from `lambda_functions.json`).
Excel's interactive Name Manager accepts `R2` even though the COM API does not.  Once
`R2` is registered manually, re-running `build_workbook.py` will maintain the other 6
names without disturbing the manually-added `R2`.
