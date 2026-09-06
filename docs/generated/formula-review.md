<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# Code review — the formulas, annotated

Every formula below is pulled from the repo at generation time
(imported constants, the JSON catalog, or a single-match source
extraction), so the docs cannot drift from the workbook. Each is
annotated in plain English.

## The self-sizing spec band

*Source: lambda_catalog/write_spec_block.py — `_spec_band` (the `Spec_Role` name)*

```excel
=TAKE(Regression!$B$4:$B$16000,MAX(1,COLUMNS(Source_Data)))
```

Take the first `COLUMNS(Source_Data)` rows of the fixed 16000-row band under column B. Because it is `TAKE` (not the volatile `OFFSET`), the name costs nothing until the table is retargeted — and retargeting `Source_Table` resizes every spec column at once. This one formula is the whole "one edit" promise.

## A status cell is a call, not text

*Source: lambda_catalog/write_spec_block.py — the B2 cell; body from lambda_functions.json*

```excel
=Role_Status()

=LAMBDA(
    IF(SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response (y)"))=0,"ERROR: no Response (y) row — mark the variable being modeled.",IF(SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response (y)"))>1,"ERROR: multiple Response (y) rows — mark exactly one.",IF(SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Fixed Effects"))>1,"ERROR: multiple Fixed Effects rows — mark at most one.","")))
)
```

Click B2 and the formula bar shows `=Role_Status()` — one name. The check itself lives in the Name Manager: exactly one Response row, at most one Fixed Effects row, and a plain-English error message in the cell when a rule is broken. Red fill comes from a conditional format keyed on the cell being non-blank.

## One formula, many rows — the computed spec columns

*Source: lambda_catalog/write_spec_block.py — `_PERIOD_IN_USE_SPILL_FORMULA`*

```excel
=LET(nc,COLUMNS(Source_Data),sq,TAKE(Spec_Sequence,nc),sp,TAKE(Spec_Sequence_Period,nc),cand,IFERROR(Base_Period_Delta_Candidate(),""),MAP(SEQUENCE(nc),LAMBDA(i,IF(INDEX(sq,i)<>TRUE,"",IF(N(INDEX(sp,i))<>0,INDEX(sp,i),cand)))))
```

Column J (Period In Use) is a single `MAP(SEQUENCE(nc), ...)` spill: for each spec row, show the typed override (column I) if one was entered, otherwise the computed candidate — the most common gap between consecutive periods. One formula covers every row and resizes with the table; `nc` is just the table's column count.

## Every statistic is a named function over the fitted model

*Source: lambda_catalog/write_sheet_regression.py — the Regression Statistics block*

```excel
=Multiple_R(Fit_Design_Columns(),Design_Response(),Fit_Sample_Include(),Fit_Context())
```

Multiple R is not a bespoke formula — it is the catalog's `Multiple_R` LAMBDA called with the materialized model: the design matrix, the response, the row mask, and the fit context. Every cell in the Regression Statistics block reads the same way, which is why a spec edit updates all of them at once.

## The Duan/Naive toggle drives the back-transformation

*Source: lambda_catalog/write_sheet_regression.py — the Unit-Space Fit block*

```excel
=Unit_Space_R_Squared(Fit_Design_Columns(),Design_Response(),Response_Column(),Fit_Sample_Include(),Fit_Context(),$AH$5)
```

When the response is Log-transformed, the fit runs in log space and `EXP(ŷ)` is the *median* prediction, not the mean. The last argument here is the AH5 toggle: Duan multiplies by the smearing factor (the mean of EXP(residuals)) to recover the conditional mean; Naive is plain EXP. One toggle re-points R², RMSE, the prediction bounds and the residual columns together.

## A distribution fit is one formula — Weibull's profile NLL

*Source: lambda_catalog/write_sheet_univariate.py — `_stage_nll` (Weibull)*

```excel
=LET(x,FILTER(UV_Data,UV_Include),BYROW($BP$33#,LAMBDA(r,LET(p,INDEX(r,1,1),IFERROR(NLL_Weibull(x,p,((AVERAGE(x^p))^(1/p))),1E+15)))))
```

The Univariate sheet's Weibull fit evaluates N points, not N², because the scale parameter is profiled out in closed form (λ̂ = (mean of xᵏ)^(1/k)). One `BYROW` walks the grid axis; `IFERROR` sits INSIDE the `LAMBDA` so one non-evaluable trial costs only its own row; `INDEX(r,1,1)` scalarizes the 1×1 row `BYROW` hands the callback. The anchor `$BP$33#` is the Stage-1 grid spill — the `#` operator reads the whole spilled axis whatever its height.

## The Beta fit's 2-D grid — N² evaluations in one spill pair

*Source: lambda_catalog/write_sheet_univariate.py — `_stage_nll` (Beta)*

```excel
=LET(d,FILTER(UV_Data,UV_Include),range_,MAX(d)-MIN(d),pad,range_*0.001,scale_,range_+2*pad,z,(d-MIN(d)+pad)/scale_,BYROW($BY$33#,LAMBDA(ab,IFERROR(NLL_Beta(z,INDEX(ab,1,1),INDEX(ab,1,2))+COUNT(d)*LN(scale_),1E+15))))
```

Beta has no closed-form partner, so each stage is a Cartesian product: an N²×2 `Full_Factorial` grid spill (Alpha | Beta) and this `BYROW` NLL column reading it via `#`. The sample is rescaled once into `z` on [pad, 1] — `NLL_Beta` needs a bounded support — and `COUNT(d)*LN(scale_)` is the Jacobian that puts the rescaled NLL back on the original scale.

## Finding the optimum: `Grid_Argument_Minimum`

*Source: lambda_functions.json — the Grid_Argument_Minimum entry*

```excel
=LAMBDA(grid,
  LET(
    value_count, COUNT(grid),
    IF(
      value_count = 0,
      HSTACK(NA(), NA(), NA()),
      LET(
        min_value, MIN(grid),
        flat_location, XMATCH(min_value, TOCOL(grid)),
        column_count, COLUMNS(grid),
        row_location, QUOTIENT(flat_location - 1, column_count) + 1,
        column_location, MOD(flat_location - 1, column_count) + 1,
        HSTACK(min_value, row_location, column_location)
      )
    )
  ))
```

The whole search-recovery trick in one LAMBDA: find the minimum of a grid, then recover WHERE it sits. `TOCOL` flattens the grid row-major, `XMATCH` finds the first minimum's flat position, and `QUOTIENT`/`MOD` convert it back to (row, column). The boundary guard reads the column: if the best shape is the first or last grid point, the optimum is on the edge and the Min/Max bounds should be widened — the sheet turns that cell red.
