# CLAUDE.md — Development Guidelines

## Excel / xlwings sheet layout

### Dynamic array spill safety

**Never place a static value or formula in a cell that falls within the potential spill range of a dynamic array formula.**

A dynamic array formula (FILTER, SEQUENCE, spill-returning LAMBDAs, etc.) spills downward from its anchor cell into as many rows as it returns. Any non-empty cell below the anchor — even a seemingly harmless header or label — will cause a `#SPILL!` error.

**Rule:** Always derive section start rows programmatically from the number of rows the preceding spill occupies, not from hardcoded magic numbers. Leave at least one blank row as a buffer.

```python
# Bad — hardcoded row that is fragile if k changes
_v(sheet, 41, col, "Pearson R")   # could collide with coefficients spill

# Good — computed from the spill size
COEFF_ANCHOR = 20
# Coefficients returns k+1 values (intercept + k predictors); spill ends at COEFF_ANCHOR + k
header_row = COEFF_ANCHOR + k + 2   # one blank gap row after spill end
_v(sheet, header_row, col, "Pearson R")
```

This applies to:
- Section title rows placed below a spill
- Column header rows placed above a new spill
- Any static label in a column that is also used by a spill formula

### Spill column isolation

Keep static labels and spill formulas in separate columns where possible. A label column (e.g., col C with predictor names) can share rows with a spill column (e.g., col D with Pearson R values) without conflict because Excel evaluates spill blocking per-column.
