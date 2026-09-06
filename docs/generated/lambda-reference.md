<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# LAMBDA function reference

Every catalog entry in `lambda_functions.json` — 152
functions. Workbook-scoped names work on any sheet; sheet-scoped names
(marked) are defined per-sheet.

## `Absorbed_Degrees_Of_Freedom` *(sheet-scoped: Regression)*

**How many degrees of freedom the Fixed Effects variable soaks up (G−1 groups) — 0 when there's no Fixed Effects row.**

Absorbed_Degrees_Of_Freedom is the one-way FE df correction for the Model Construction sheet: G − 1, where G is the count of distinct Fixed Effects levels among the included rows. Every model with no declared Fixed Effects row (the entire v2.0 shipped surface) returns 0 — the same non-breaking default the [DF_Absorbed] argument threads through the inference chain, so a no-FE model's df, MS-Residual, t, p, CI, and AIC/BIC are bit-identical to before this function existed.

Reuses Dummy_Levels for the level count rather than a fresh UNIQUE/COUNT — the exact same mask-scoped, reference-dropped level set the design matrix would use if the FE column were instead one-hot encoded (which is the algebraic content of 'absorbing' a fixed effect: G−1 degrees of freedom are spent whether they are spent explicitly, as G−1 dummy coefficients, or implicitly, by demeaning). COLUMNS(Dummy_Levels(...)) is exactly G−1 by construction, so this function can never disagree with what an equivalent LSDV fit would report. A degenerate FE variable (every included row shares one level) makes Dummy_Levels return #N/A; this function absorbs that into 0 rather than propagating the error, because a single-level 'group' fixes no group-specific intercept and absorbs no degrees of freedom.

One-way only: a spec with two or more Fixed Effects rows is a visible status-block error (validated elsewhere), so this function reads only the first Fixed Effects row via Fixed_Effects_Column() — two-way absorption (Σ over both dimensions, minus the overlap correction) is deferred to the two-way FE milestone.

Returns: The degrees of freedom absorbed by the Fixed Effects group: G − 1 where G is the FE column's level count over the included sample; 0 when no Fixed Effects row is declared or the FE variable is degenerate (G ≤ 1).

One-way FE df correction: G−1 via COLUMNS(Dummy_Levels(...)) on the FE column, 0 when no FE row exists or the FE variable is degenerate. Default 0 keeps every no-FE model bit-identical. Two-way absorption deferred to the two-way FE milestone.

```excel
=LAMBDA(LET(
    n_c,       COLUMNS(Source_Data),
    fe_active, SUMPRODUCT(N(TAKE(Spec_Role, n_c) = "Fixed Effects")) > 0,
    IF(NOT(fe_active),
       0,
       LET(
           lv, Dummy_Levels(Fixed_Effects_Column(), "", Log_Drop_Sample_Include_Calc()),
           IF(ISNA(INDEX(lv, 1, 1)), 0, COLUMNS(lv))
       )
    )
))
```

## `Adjusted_R_Squared`

**R-squared adjusted so extra predictors do not get free credit.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Adjusted R² penalises R² for the number of predictors relative to sample size, correcting for the inflation that occurs when irrelevant predictors are added. It equals 1 - (1 - R²) × df_total / df_residual, which reduces to the familiar 1 - (1 - R²)(n-1)/(n-k-1) when an intercept is fit. Unlike R² it can decrease when a predictor adds less explanatory power than expected by chance, and can be negative when the model fits worse than a horizontal line.

Computed as 1 - (1 - R_Squared(...)) * Total_Degrees_Of_Freedom(...) / Residual_Degrees_Of_Freedom(...), inheriting the intercept correction from Total_Degrees_Of_Freedom (the residual df needs none — COLUMNS(X) already counts the intercept column). An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to Residual_Degrees_Of_Freedom() so the penalty term correctly reflects absorbed group degrees of freedom.

Returns: Adjusted R² as a scalar, typically in [0, 1] but can be negative

R² penalised for model complexity: 1 − (1 − R²)·df_total/df_residual. Can be negative when the model fits worse than the mean. Optional DF_Absorbed (default 0) corrects df_residual under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg,     IF(ISOMITTED(Include),       TRUE, Include),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    1 - (1 - R_Squared(X, Y, filt_arg, context_arg))
      * Total_Degrees_Of_Freedom(Y, filt_arg, context_arg)
      / Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg)
  ))
```

## `AIC`

**A score for comparing models that rewards fit but penalizes extra complexity.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

AIC = n·ln(SSR/n) + 2p, where p = COLUMNS(X), the number of regression parameters. Since the v3.0 intercept relocation the design matrix carries its own intercept column, so COLUMNS(X) counts it directly and no separate intercept term is added. Smaller values indicate better-fitting models; compare across models on the same dataset.

Under v2.1 Fixed Effects, an optional trailing DF_Absorbed adds directly into p: the G−1 group intercepts a Fixed Effects variable absorbs are genuinely estimated parameters even though the within transformation never materializes them as explicit coefficients, so AIC must count them or understate model complexity.

Returns: Akaike Information Criterion (AIC) scalar

Akaike Information Criterion = n·ln(SSR/n) + 2p. Lower = better. Optional DF_Absorbed (default 0) adds absorbed FE group intercepts into p.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    n,   Observations(Y, filt_arg),
    ssr, SS_Residual(X, Y, filt_arg),
    p,   COLUMNS(X) + absorbed_arg,
    n * LN(ssr / n) + 2 * p
  ))
```

## `AICc`

**AIC with an extra small-sample penalty.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

AICc = AIC + 2p(p+1)/(n−p−1). The small-sample correction reduces upward bias in AIC when n/p < 40. Converges to AIC as n → ∞.

Under v2.1 Fixed Effects, an optional trailing DF_Absorbed adds directly into p, propagating into both the AIC term and the small-sample correction term.

Returns: Corrected Akaike Information Criterion (AICc) scalar

AIC with a small-sample correction (+2p(p+1)/(n−p−1)). Preferred when n/p < 40; converges to AIC as n → ∞. Optional DF_Absorbed (default 0) adds absorbed FE group intercepts into p.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    n,   Observations(Y, filt_arg),
    ssr, SS_Residual(X, Y, filt_arg),
    p,   COLUMNS(X) + absorbed_arg,
    aic, n * LN(ssr / n) + 2 * p,
    aic + 2 * p * (p + 1) / (n - p - 1)
  ))
```

## `Back_Transform_Response`

**Convert a fit-space column to response units — EXP(x) under Log (with or without Duan smearing); pass-through under None.**

Arguments:

- **Values** — n × 1 column of fit-space values to back-transform (LN(y) for a Log response; y unchanged for None)
- **Context** — the materialized Model_Context() 4x1 array; defaults to VSTACK(TRUE,0,"None","None") when omitted
- **Method** — Duan (multiply by smearing factor — predicts the conditional mean) or Naive (EXP only — predicts the conditional median, biased for the mean); defaults to "Duan" so unit-space R² always corresponds to the displayed predictions
- **Smearing** — the scalar smearing factor (Smearing_Factor(X, Y, Include, Context)); defaults to 1, so an omitted smearing under Duan collapses to the Naive result

Back_Transform_Response returns the response-scale counterpart of a fit-space column, applying Duan smearing or a naive EXP back-transform when the response is log-transformed. It is the single dispatch site for the response-side transform pair — every other catalog function in the back-transformation family (Unit_Space_Predictions, Unit_Space_Residuals) is built on top of it, so the (response, method) SWITCH lives in exactly one place.

Three recognised branches: Context_Response_Transform = "None" returns Values unchanged (the spec's response transform carries through to the display); "Log" with Method = "Duan" returns EXP(Values) * Smearing; "Log" with Method = "Naive" returns EXP(Values). Any other (transform, method) pair returns #N/A — a deliberately honest signal, not a silent default to Duan. The Duan case is the default because unit-space R² must always correspond to the predictions the sheet shows, and the sheet's default back-transform-method toggle is Duan.

CI/PI bounds are quantiles, not means, so a back-transformation that smears them (multiplies both bounds by the same factor) preserves coverage but would skew the interval away from its nominal width — see the AMENDMENT in docs/v3.3-unit-space-and-back-transformation.md about EXP-only bound back-transformation. The sheet forces the bounds' Back_Transform_Response call to "Naive" for that reason; the caller controls the method, not the function.

Returns: n × 1 column: Values unchanged when Context_Response_Transform = "None"; EXP(Values) * Smearing under Duan, EXP(Values) under Naive, when = "Log"; #N/A outside the recognised response-transform × method pairs.

Back-transforms a fit-space column to response units. "None" → pass-through. "Log" + Duan → EXP(x)*Smearing. "Log" + Naive → EXP(x). Anything else → #N/A. The (response, method) SWITCH lives here only; the Unit_Space_* names are built on top of it.

```excel
=LAMBDA(Values, [Context], [Method], [Smearing],
  LET(
    context_arg,  IF(ISOMITTED(Context), Model_Context(), Context),
    method_arg,   IF(ISOMITTED(Method), "Duan", Method),
    smear_arg,    IF(ISOMITTED(Smearing), 1, Smearing),
    rt,           Context_Response_Transform(context_arg),
    IF(rt="None", Values,
      IF(rt="Log",
        SWITCH(method_arg,
          "Duan",  IFERROR(EXP(Values) * smear_arg, NA()),
          "Naive", IFERROR(EXP(Values), NA()),
          NA()),
        NA()))
  )
)
```

## `Base_Period_Delta` *(sheet-scoped: Regression)*

**The gap between consecutive time periods (Δ) currently set in the model spec — what Lag_By and Difference_By use when you don't pass one.**

Base_Period_Delta returns the base period Δ currently in effect for THIS sheet's sequence axis: the value of the Period In Use spec cell (column J) on the row whose Sequence flag (column H) is TRUE. That cell is computed-with-override — the build pre-fills it with Base_Period_Delta_Candidate() (the MODE of within-group consecutive spacings), and a user-typed number into column I (Sequence Period) replaces the formula to override. Either way, this accessor reads whatever the J cell shows, so the Δ that drives Lag_By and Difference_By defaults is always the Δ visible in the spec.

Δ is never silently assumed to be 1. When no variable is flagged as the Sequence axis, or the flagged row's J cell holds no number, the function returns #N/A — and an omitted-delta Lag_By/Difference_By call therefore returns #N/A everywhere rather than fabricating single-step differences. This is the same visible-default-with-override pattern as the categorical reference level (blank E cell → surfaced default in L).

SHEET-SCOPED (scope: Regression), like the constructor closures it reads. It was workbook-scoped and sheet-qualified ('Regression'!Spec_Sequence), which had two costs. A workbook with more than one Regression-shaped sheet — the test-model artifact has 47 — would have every one of them read the single sheet named 'Regression'; and a workbook with NO such sheet had the function skipped at build time to avoid an external link, leaving #NAME? wherever it was called. Unqualified names resolve against the sheet the calling formula lives on, so one sheet-scoped definition per Regression sheet gives each its own Δ. The narrow cost: an omitted-delta Lag_By/Difference_By evaluated on a sheet that has no spec block (a data sheet, say) now returns #NAME? instead of borrowing the Regression sheet's Δ — which is the honest answer, since that sheet declares no sequence axis.

Returns: The base period Δ in effect — the numeric value of spec cell J (Period In Use) on the Sequence-flagged row of the Regression sheet; #N/A when no Sequence axis is flagged or the cell holds no number.

The base period Δ in effect: spec column J (Period In Use) on the Sequence-flagged row of the Regression sheet. #N/A when no axis is flagged or the cell is not numeric. Default for Lag_By / Difference_By.

```excel
=LAMBDA(
    LET(
        n_c,      COLUMNS(Source_Data),
        position, XMATCH(TRUE, TAKE(Spec_Sequence, n_c), 0),
        value,    INDEX(TAKE(Spec_Sequence_Period, n_c), position),
        IF(ISNUMBER(value), value, NA())
    )
)
```

## `Base_Period_Delta_Candidate` *(sheet-scoped: Regression)*

**The suggested time step for the data — the most common gap between consecutive periods within a group (1 for yearly data).**

Base_Period_Delta_Candidate computes the natural base period of the declared sequence axis: the MODE.SNGL of Sequence_Deltas() — the most common within-group consecutive spacing. For a complete yearly panel that is 1; for a decadal survey it is 10. The build writes this candidate into spec cell J (Period In Use) of the Sequence-flagged row; cell I (Sequence Period) is the user-typed override input, and a non-blank I replaces the formula via the reference-level pattern.

When MODE.SNGL returns #N/A because every spacing is distinct, there is no natural base period; the candidate falls back to the MIN of the positive spacings (the finest grid the data could support) and the Sequence Spacing block surfaces an explicit override prompt ("no repeated spacing — set Δ"). When there are no within-group spacings at all — no sequence axis flagged, or every group has a single observation — the candidate is #N/A and cell J displays blank, awaiting a typed Δ in I.

The candidate never quantizes calendar spacings: monthly serial dates produce a spectrum clustered at 28–31 and a MODE of 31, which is a wrong Δ for eleven months of the year. That situation is surfaced (calendar-signature guidance recommending an integer period index upstream), not papered over with a scalar.

Returns: The computed Base Period Δ candidate: MODE of the within-group spacings, falling back to their MIN when no spacing repeats; #N/A when there are no spacings at all.

Computed Base Period Δ candidate: MODE.SNGL of Sequence_Deltas(), MIN fallback when no spacing repeats. Pre-fills spec cell J (Period In Use); cell I (Sequence Period) is the override input.

```excel
=LAMBDA(
    LET(
        d, Sequence_Deltas(),
        IF(COUNT(d) = 0, NA(), IFERROR(MODE.SNGL(d), MIN(d)))
    )
)
```

## `Beta_Weights`

**Coefficients rescaled so predictors on different units can be compared fairly.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Standardized regression coefficients — each slope rescaled by the ratio of its predictor's standard deviation to the response's, making magnitudes comparable across predictors measured in different units.

The intercept is excluded from both halves: its coefficient is dropped, and the per-column standard deviations are taken over the predictor columns of X only. A standard deviation of the intercept column would be zero and the whole vector would come back as zeros or errors.

Returns: k×1 column vector of standardized coefficients (Beta weights), one per predictor — intercept row excluded

Standardised coefficients in SD units for comparing predictors measured on different scales. k rows; no intercept row.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg,   IF(ISOMITTED(Include),       TRUE, Include),
    coefs,      Coefficients(X, Y, filt_arg),
    preds,      IF(has_arg, DROP(X, , 1), X),
    pred_coefs, IF(has_arg, DROP(coefs, 1), coefs),
    sd_y,       STDEV(FILTER(Y, filt_arg)),
    k,          COLUMNS(preds),
    sd_x,       MAKEARRAY(k, 1, LAMBDA(j, _, STDEV(FILTER(CHOOSECOLS(preds, j), filt_arg)))),
    pred_coefs * sd_x / sd_y
  ))
```

## `BFN_Panel_Durbin_Watson`

**The panel version of the Durbin-Watson autocorrelation check — residual changes are measured within each group (country) only, so group boundaries never fake a correlation.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **group** — single-column panel-unit identifiers (the Fixed Effects variable, e.g. Country) — residual differences never cross a group boundary
- **seq** — single-column ordering axis (the declared Sequence column, e.g. Year) — residuals are sorted by this before differencing, so the result does not depend on physical row order
- **delta** — base period Δ: differences pair û(i,t) with û(i,t−delta) by exact time value; when omitted, the spec's Base Period Δ (Base_Period_Delta()) is used — never a silent 1
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

BFN_Panel_Durbin_Watson is the Bhargava–Franzini–Narendranathan (1982) panel form of the Durbin-Watson statistic: BFN = Σᵢ Σₜ₌₂ (û(i,t) − û(i,t−Δ))² / Σᵢ Σₜ û(i,t)², where the numerator differences residuals only within a panel unit.

Why this exists: under fixed effects, ordinary Durbin-Watson is invalid twice over — the residuals are within-demeaned, and a panel has no single ordering. Naive row-adjacent differencing manufactures correlation at every group seam (unit A's last period sits next to unit B's first), and even the sequence-sorted Durbin_Watson_By still differences across units that share a time value. BFN restricts the numerator's differencing to within-group pairs, so seams contribute nothing by construction.

The within-group differencing is delegated to Difference_By (one source of truth): the prior residual is the exact-match (group, seq − Δ) lookup, never row arithmetic, so the statistic is invariant to physical row order — permuting whole group blocks leaves it unchanged. Difference_By returns #N/A at each group's first period and across panel gaps (a punched-out year contributes no fabricated difference); the IFERROR(…, 0) mask in the numerator zeroes exactly those terms, locally and visibly, per the consumer-owns-masking rule. The denominator sums every residual's square — first periods and gap rows still count there, as the BFN definition requires. A filtered-in row whose group is blank or whose seq is not numeric is a data defect and surfaces as a visible error rather than being silently dropped.

The mask never turns an error STATE into a number: when not a single difference is computable — every group a singleton, or Δ unresolved (no Base Period Δ declared and delta omitted, so step is #N/A and every lookup misses) — the all-#N/A difference column would otherwise mask to an all-zero numerator and display BFN = 0, a fake strong-negative-autocorrelation reading. The n_terms guard returns a genuine #N/A instead: zero computable differences is an error state, not a statistic.

Interpretation caveat: read it like Durbin-Watson — near 2 means no first-order autocorrelation in the within residuals — but do NOT apply the standard DW significance bounds. BFN has its own critical values that depend on both N (units) and T (periods); Bhargava, Franzini & Narendranathan (1982) tabulate them. Surfacing those bounds on the sheet is a recorded open item.

Output is always a scalar; residual computation, filtering, and differencing stay inside the LAMBDA so no helper column or spill is exposed.

Returns: Bhargava–Franzini–Narendranathan panel Durbin-Watson statistic as a scalar: Σᵢ Σₜ (û(i,t) − û(i,t−Δ))² / Σᵢ Σₜ û(i,t)², with differencing restricted to within-group (group, seq−Δ) pairs

Bhargava–Franzini–Narendranathan panel DW: differences residuals within (group, seq−Δ) pairs via Difference_By, NA→0 masked locally; denominator sums all e². Near 2 = no autocorrelation; N,T-dependent critical values — standard DW bounds do not apply.

```excel
=LAMBDA(X, Y, group, seq, [delta], [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    step,     IF(ISOMITTED(delta), Base_Period_Delta(), delta),
    e,        Residuals(X, Y, filt_arg),
    g_f,      FILTER(group, filt_arg),
    t_f,      FILTER(seq, filt_arg),
    d,        Difference_By(e, g_f, t_f, step),
    n_terms,  SUMPRODUCT(N(ISNUMBER(d))),
    IF(n_terms = 0, NA(),
      SUMPRODUCT(IFERROR(d, 0)^2) / SUMPRODUCT(e^2))
  ))
```

## `BIC`

**A model-comparison score that penalizes extra complexity more strongly than AIC.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

BIC = n·ln(SSR/n) + p·ln(n), where p = number of regression parameters. Applies a heavier penalty than AIC for additional parameters, favouring sparser models in large samples.

Under v2.1 Fixed Effects, an optional trailing DF_Absorbed adds directly into p, the same absorbed-group-intercepts-are-parameters correction AIC applies.

Returns: Bayesian Information Criterion (BIC) scalar

Bayesian Information Criterion = n·ln(SSR/n) + p·ln(n). Penalises model complexity more strongly than AIC. Optional DF_Absorbed (default 0) adds absorbed FE group intercepts into p.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    n,   Observations(Y, filt_arg),
    ssr, SS_Residual(X, Y, filt_arg),
    p,   COLUMNS(X) + absorbed_arg,
    n * LN(ssr / n) + p * LN(n)
  ))
```

## `Bin_Counts`

**How many data values fall into each histogram bin.**

Arguments:

- **data** — single-column numeric data range
- **method** — bin-count rule: "Sturges", "Scott", or "FD" (default when omitted)
- **filter** — optional boolean array

Bin_Counts returns the number of observations in each bin, using Excel's FREQUENCY function applied to the upper edges from Upper_Bin_Edges.

Returns: k × 1 column vector of bin frequencies

Counts of observations in each of k bins using Excel FREQUENCY on Upper_Bin_Edges. Returns a k×1 column vector.

```excel
=LAMBDA(data, [method], [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    upper, Upper_Bin_Edges(data, method, filter),
    TAKE(FREQUENCY(x, upper), ROWS(upper))
  ))
```

## `Bin_Edges`

**The complete list of k+1 bin boundary values from min to max.**

Arguments:

- **data** — single-column numeric data range
- **method** — bin-count rule: "Sturges", "Scott", or "FD" (default when omitted)
- **filter** — optional boolean array

Bin_Edges returns the complete set of k+1 bin boundaries as a column vector, starting at the data minimum and ending at the exact data maximum. Intermediate boundaries are spaced by width = (max - min) / k. Upper_Bin_Edges, Bin_Lower_Edges, Bin_Midpoints, and Bin_Counts all derive from this vector.

Returns: k+1 × 1 column vector of bin boundaries — the minimum of the data followed by k evenly-spaced upper edges

k+1 evenly-spaced bin boundaries from data minimum to maximum. The master boundary vector; all other Bin_* functions derive from it.

```excel
=LAMBDA(data, [method], [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    min_x, MIN(x),
    max_x, MAX(x),
    k, Number_Of_Histogram_Bins(data, IF(ISOMITTED(method), "FD", method), filt_arg),
    width, (max_x - min_x) / k,
    VSTACK(SEQUENCE(k, 1, min_x, width), max_x)
  ))
```

## `Bin_Lower_Edges`

**The lower boundary of each histogram bin — data minimum for bin 1, then each previous upper edge.**

Arguments:

- **data** — single-column numeric data range
- **method** — bin-count rule: "Sturges", "Scott", or "FD" (default when omitted)
- **filter** — optional boolean array

Bin_Lower_Edges returns the k lower bin edges derived from Bin_Edges by dropping the last boundary (the data maximum). The first lower edge is the data minimum.

Returns: k × 1 column vector of lower bin edges (the k+1 boundary vector with the last element dropped)

k lower bin boundaries (Bin_Edges without the last element). First element = data minimum.

```excel
=LAMBDA(data, [method], [filter],
  DROP(Bin_Edges(data, method, filter), -1)
)
```

## `Bin_Midpoints`

**The centre of each histogram bin — use as the X-axis value for distribution overlay curves.**

Arguments:

- **data** — single-column numeric data range
- **method** — bin-count rule: "Sturges", "Scott", or "FD" (default when omitted)
- **filter** — optional boolean array

Bin_Midpoints returns the midpoint of each bin, computed as the average of consecutive boundaries from Bin_Edges: (upper_edge + lower_edge) / 2. Bin_Edges is called once via LET to avoid recomputation.

Returns: k × 1 column vector of bin midpoints

k bin centre values = (lower + upper) / 2. Use as x-axis values when overlaying fitted distribution PDF curves.

```excel
=LAMBDA(data, [method], [filter],
  LET(
    e, Bin_Edges(data, method, filter),
    (DROP(e, 1) + DROP(e, -1)) / 2
  ))
```

## `CDF_Beta`

**The probability that a rescaled Beta variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array, on the original data scale)
- **alpha** — Beta shape parameter α (must be > 0)
- **beta** — Beta shape parameter β (must be > 0)
- **data_min** — minimum of the fitted data (used for rescaling to [0,1])
- **data_range** — max − min of the fitted data (used for rescaling to [0,1])
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from the padded support minimum (data_min − pad)

CDF_Beta returns the probability that a Beta(alpha, beta) random variable (rescaled from [0,1] to the original data range) falls in the interval (minimum, maximum]. Rescaling uses the same 0.1 % padding as NLL_Beta to keep boundary data points interior to the support.

Takes data_min and data_range rather than the full data array to avoid redundant computation when called repeatedly (e.g. once per histogram bin).

Returns: CDF(maximum) − CDF(minimum) for rescaled Beta(alpha, beta); probability of the interval

P(minimum < X ≤ maximum) for Beta(alpha, beta) rescaled to the original data range using data_min and data_range.

```excel
=LAMBDA(maximum, alpha, beta, data_min, data_range, [minimum],
  LET(
    pad, MAX(data_range * 0.001, 1E-30),
    scale_, data_range + 2 * pad,
    z_max, (maximum - data_min + pad) / scale_,
    cdf_max, IFERROR(BETA.DIST(z_max, alpha, beta, TRUE), 0.5),
    cdf_min, IF(ISOMITTED(minimum), 0,
      LET(z_min, (minimum - data_min + pad) / scale_,
        IFERROR(BETA.DIST(z_min, alpha, beta, TRUE), 0.5))),
    cdf_max - cdf_min
  ))
```

## `CDF_BetaPERT`

**The probability that a BetaPERT variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array)
- **min_val** — distribution minimum
- **mode_val** — distribution mode (most likely value)
- **max_val** — distribution maximum
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from min_val

CDF_BetaPERT returns the probability that a BetaPERT(min, mode, max) random variable falls in the interval (minimum, maximum]. The λ=4 PERT mapping α = 1 + 4·(mode − min)/range and β = 1 + 4·(max − mode)/range (the same parameterisation as NLL_BetaPERT) is evaluated via BETA.DIST on the [0,1]-rescaled value. This form is algebraically identical to the μ-based reparameterisation but has no 0/0 singularity at a symmetric mode (mode = (min + max)/2), where the μ form degenerates to α = β = 0.

Epsilon guards (1E-30) prevent division by zero for degenerate parameters.

Returns: CDF(maximum) − CDF(minimum) for BetaPERT(min_val, mode_val, max_val); probability of the interval

P(minimum < X ≤ maximum) for BetaPERT(min, mode, max). PERT parameters are auto-derived from three-point estimates.

```excel
=LAMBDA(maximum, min_val, mode_val, max_val, [minimum],
  LET(
    range_, max_val - min_val + 1E-30,
    alpha_param, 1 + 4 * (mode_val - min_val) / range_,
    beta_param, 1 + 4 * (max_val - mode_val) / range_,
    z_max, (maximum - min_val) / range_,
    cdf_max, BETA.DIST(z_max, alpha_param, beta_param, TRUE),
    cdf_min, IF(ISOMITTED(minimum), 0,
      LET(z_min, (minimum - min_val) / range_,
        BETA.DIST(z_min, alpha_param, beta_param, TRUE))),
    cdf_max - cdf_min
  ))
```

## `CDF_Exponential`

**The probability that an Exponential variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array)
- **rate** — rate parameter λ = 1/mean
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from 0

CDF_Exponential returns the probability that an Exponential(rate) random variable falls in the interval (minimum, maximum]. When minimum is omitted the result is the cumulative probability from 0 to maximum.

Accepts arrays for maximum and minimum, spilling element-wise.

Returns: CDF(maximum) − CDF(minimum) for Exponential(rate); probability of the interval

P(minimum < X ≤ maximum) for Exponential(rate = 1/mean). Data must be ≥ 0. Accepts arrays for element-wise calculation.

```excel
=LAMBDA(maximum, rate, [minimum],
  LET(
    cdf_max, EXPON.DIST(maximum, rate, TRUE),
    cdf_min, IF(ISOMITTED(minimum), 0, EXPON.DIST(minimum, rate, TRUE)),
    cdf_max - cdf_min
  ))
```

## `CDF_Gamma`

**The probability that a Gamma variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array)
- **shape** — shape parameter α (must be > 0)
- **rate** — rate parameter β = 1/scale (must be > 0)
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from 0

CDF_Gamma returns the probability that a Gamma(shape, rate) random variable falls in the interval (minimum, maximum]. GAMMA.DIST takes scale = 1/rate, matching the rate parameterisation used by NLL_Gamma.

Accepts arrays for maximum and minimum, spilling element-wise.

Returns: CDF(maximum) − CDF(minimum) for Gamma(shape, rate); probability of the interval

P(minimum < X ≤ maximum) for Gamma(shape, rate = 1/scale). Data must be > 0. Accepts arrays for element-wise calculation.

```excel
=LAMBDA(maximum, shape, rate, [minimum],
  LET(
    cdf_max, GAMMA.DIST(maximum, shape, 1 / rate, TRUE),
    cdf_min, IF(ISOMITTED(minimum), 0, GAMMA.DIST(minimum, shape, 1 / rate, TRUE)),
    cdf_max - cdf_min
  ))
```

## `CDF_Lognormal`

**The probability that a Lognormal variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array)
- **meanlog** — mean of ln(x)
- **sdlog** — standard deviation of ln(x) (must be > 0)
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from 0

CDF_Lognormal returns the probability that a Lognormal(meanlog, sdlog) random variable falls in the interval (minimum, maximum]. When minimum is omitted the result is the cumulative probability from 0 to maximum.

Accepts arrays for maximum and minimum, spilling element-wise.

Returns: CDF(maximum) − CDF(minimum) for Lognormal(meanlog, sdlog); probability of the interval

P(minimum < X ≤ maximum) for Lognormal(meanlog, sdlog). Data must be positive. Accepts arrays for element-wise calculation.

```excel
=LAMBDA(maximum, meanlog, sdlog, [minimum],
  LET(
    cdf_max, LOGNORM.DIST(maximum, meanlog, sdlog, TRUE),
    cdf_min, IF(ISOMITTED(minimum), 0, LOGNORM.DIST(minimum, meanlog, sdlog, TRUE)),
    cdf_max - cdf_min
  ))
```

## `CDF_Normal`

**The probability that a Normal variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array)
- **mean** — distribution mean μ
- **sd** — distribution standard deviation σ (must be > 0)
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from −∞

CDF_Normal returns the probability that a Normal(mean, sd) random variable falls in the interval (minimum, maximum]. When minimum is omitted the result is the cumulative probability from −∞ to maximum.

Accepts arrays for maximum and minimum, spilling element-wise. Use with Bin_Edges as maximum and Bin_Lower_Edges as minimum to compute the expected proportion of each histogram bin under the fitted Normal distribution.

Returns: CDF(maximum) − CDF(minimum) for Normal(mean, sd); probability of the interval

P(minimum < X ≤ maximum) for Normal(mean, sd). Omit minimum for P(X ≤ maximum). Accepts arrays; use with Bin_Edges for histogram overlay.

```excel
=LAMBDA(maximum, mean, sd, [minimum],
  LET(
    cdf_max, NORM.DIST(maximum, mean, sd, TRUE),
    cdf_min, IF(ISOMITTED(minimum), 0, NORM.DIST(minimum, mean, sd, TRUE)),
    cdf_max - cdf_min
  ))
```

## `CDF_Triangular`

**The probability that a Triangular variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array)
- **min_val** — distribution minimum a
- **mode_val** — distribution mode c (must satisfy a < c < b)
- **max_val** — distribution maximum b
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from min_val

CDF_Triangular returns the probability that a Triangular(a, c, b) random variable falls in the interval (minimum, maximum]. Uses the piecewise closed-form CDF since Excel has no built-in Triangular CDF. MAP applies the nested CDF LAMBDA element-wise so the function works with array inputs.

The epsilon guard (1E-30) prevents division by zero for degenerate parameters. IFERROR maps any remaining errors to 0.5.

Returns: CDF(maximum) − CDF(minimum) for Triangular(min_val, mode_val, max_val); probability of the interval

P(minimum < X ≤ maximum) for Triangular(min_val, mode_val, max_val). No Excel built-in; uses closed-form piecewise CDF.

```excel
=LAMBDA(maximum, min_val, mode_val, max_val, [minimum],
  LET(
    cdf_at, LAMBDA(x,
      IFERROR(IF(x < min_val, 0,
        IF(x < mode_val,
          (x - min_val)^2 / ((max_val - min_val) * (mode_val - min_val) + 1E-30),
          IF(x <= max_val,
            1 - (max_val - x)^2 / ((max_val - min_val) * (max_val - mode_val) + 1E-30),
            1))), 0.5)),
    cdf_max, MAP(maximum, cdf_at),
    cdf_min, IF(ISOMITTED(minimum), 0, MAP(minimum, cdf_at)),
    cdf_max - cdf_min
  ))
```

## `CDF_Weibull`

**The probability that a Weibull variable falls between minimum and maximum.**

Arguments:

- **maximum** — upper bound of the interval (scalar or array)
- **shape** — shape parameter k (must be > 0)
- **scale** — scale parameter λ (must be > 0)
- **minimum** — lower bound of the interval; when omitted, returns CDF(maximum) from 0

CDF_Weibull returns the probability that a Weibull(shape, scale) random variable falls in the interval (minimum, maximum]. When minimum is omitted the result is the cumulative probability from 0 to maximum.

Accepts arrays for maximum and minimum, spilling element-wise.

Returns: CDF(maximum) − CDF(minimum) for Weibull(shape, scale); probability of the interval

P(minimum < X ≤ maximum) for Weibull(shape, scale). Data must be > 0. Accepts arrays for element-wise calculation.

```excel
=LAMBDA(maximum, shape, scale, [minimum],
  LET(
    cdf_max, WEIBULL.DIST(maximum, shape, scale, TRUE),
    cdf_min, IF(ISOMITTED(minimum), 0, WEIBULL.DIST(minimum, shape, scale, TRUE)),
    cdf_max - cdf_min
  ))
```

## `Coefficients`

**The fitted intercept and predictor weights used to make predictions.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Returns the fitted OLS coefficient vector, in design-matrix column order.

LINEST is called with const = FALSE. This is not a stylistic choice: the design matrix arrives with its intercept column already in it, and letting Excel fit a second intercept on top would produce two exactly collinear terms, a singular Gram matrix, and a result that is wrong rather than absent. With const = FALSE, LINEST returns exactly COLUMNS(X) coefficients in reverse column order, so unwinding it is a single reversal — the old intercept-splicing branch is gone.

Returns: vertical array of fitted OLS coefficients — intercept first (when included), then predictors in input order

Fitted OLS coefficients in design-matrix column order. LINEST runs with const = FALSE because the design matrix already carries its intercept column.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    ls,       LINEST(FILTER(Y, filt_arg), FILTER(X, filt_arg), FALSE, FALSE),
    k,        COLUMNS(X),
    TRANSPOSE(CHOOSECOLS(ls, SEQUENCE(1, k, k, -1)))
  ))
```

## `Column_Select`

**A helper that picks specific columns from a larger table.**

Arguments:

- **table** — source range or table body (multi-column)
- **col_nums** — horizontal array of 1-based column numbers, e.g. {1,3,5}; negative indices count from the right

Wraps CHOOSECOLS to give column selection a named, self-documenting form. The output is a calculated array with the same row count as the source table, valid as X in any library function.

Primary use: define a named range — e.g. X_Predictors = Column_Select(LifeData, {3,7,9}) — then pass that name as X. Non-contiguous column sets are fully supported. Combine with Complete_Cases_Filter to drop rows that have missing or non-numeric values in the chosen columns.

Returns: array of the selected columns in the specified order

Named wrapper for CHOOSECOLS: picks non-contiguous predictor columns by index. Returns an array suitable for X.

```excel
=LAMBDA(table, col_nums,
  CHOOSECOLS(table, col_nums)
)
```

## `Complete_Cases_Filter`

**A TRUE/FALSE mask that keeps only rows with usable numeric data.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Y** — single-column outcome range

Produces a ready-made Include argument for any library function. TRUE rows have no blanks, text, or errors in any of the supplied columns.

When Y is omitted, only Predictors columns are checked. When Y is provided, it is HSTACKed onto Predictors so that a single MMULT row-sum determines completeness across all regression inputs at once. ISNUMBER returns FALSE for errors and blanks, so error cells are safely treated as incomplete without propagating into the MMULT result.

Typical usage: R_Squared(Predictors, Y, , Complete_Cases_Filter(Predictors, Y))

Returns: boolean column vector — TRUE for rows where every included column is numeric

Boolean vector TRUE where every included column is numeric. Pass as the Include argument to exclude incomplete rows.

```excel
=LAMBDA(Predictors, [Y],
  LET(
    combined, IF(ISOMITTED(Y), Predictors, HSTACK(Predictors, Y)),
    MMULT((ISNUMBER(combined)) * 1, SEQUENCE(COLUMNS(combined), 1, 1, 0)) = COLUMNS(combined)
  )
)
```

## `Confidence_Interval_Lower`

**The low end of the likely range for each coefficient.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Alpha** — significance level (0–1); default 0.05 yields 95% CIs.
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Returns a vertical spill array of the lower bounds of the confidence intervals for each OLS coefficient. Uses a two-tailed t critical value at significance level Alpha (default 0.05, yielding 95% CIs).

Computed as Coefficients(...) - T.INV.2T(Alpha, Residual_Degrees_Of_Freedom(...)) * SE_Coefficients(...). An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to both the t-critical value's df and the coefficient SE.

Returns: vertical array of lower confidence interval bounds for each coefficient

Lower CI bounds per coefficient. Default Alpha = 0.05 gives 95% CIs. Optional DF_Absorbed (default 0) corrects both df and SE under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Alpha], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    alpha_arg,    IF(ISOMITTED(Alpha),       0.05, Alpha),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    Coefficients(X, Y, filt_arg)
      - T.INV.2T(alpha_arg, Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg))
        * SE_Coefficients(X, Y, filt_arg, context_arg)
  ))
```

## `Confidence_Interval_Upper`

**The high end of the likely range for each coefficient.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Alpha** — significance level (0–1); default 0.05 yields 95% CIs.
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Returns a vertical spill array of the upper bounds of the confidence intervals for each OLS coefficient. Uses a two-tailed t critical value at significance level Alpha (default 0.05, yielding 95% CIs).

Computed as Coefficients(...) + T.INV.2T(Alpha, Residual_Degrees_Of_Freedom(...)) * SE_Coefficients(...). An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to both the t-critical value's df and the coefficient SE.

Returns: vertical array of upper confidence interval bounds for each coefficient

Upper CI bounds per coefficient. Default Alpha = 0.05 gives 95% CIs. Optional DF_Absorbed (default 0) corrects both df and SE under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Alpha], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    alpha_arg,    IF(ISOMITTED(Alpha),       0.05, Alpha),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    Coefficients(X, Y, filt_arg)
      + T.INV.2T(alpha_arg, Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg))
        * SE_Coefficients(X, Y, filt_arg, context_arg)
  ))
```

## `Constructed_Column_Names` *(sheet-scoped: Regression)*

**The column headings for the design matrix — like 'Status: Developing', 'Year: 2001', or 'Ln(Weight)' for a logged predictor — one per X_s column.**

Sheet-scoped header strip for the Model Construction sheet: the structural twin of X. Same iteration, same Predictor-AND-Include predicate, same Dummy_Levels skip conditions, emitting a name per contributed column — the bare header for a Continuous predictor (or 'Ln(header)' when that row's Transform is Log), 'Header: level' for each retained Categorical level (never relabelled — Log is disallowed on Categorical). Twinning guarantees the header width always equals COLUMNS(X()); QC asserts the two never diverge. The degenerate skip tests ISNA(INDEX(lv,1,1)), a scalar: lv is a 1x(L-1) row, and an array ISNA condition in front of a wider HSTACK branch broadcasts to #N/A, so the guard must be scalarized.

Returns: One header per constructed column, in the same order and width as Predictor_Columns(): the raw name, Ln(name) under Transform=Log, "name: level" per retained categorical level, and for an interaction column the two operands' own names joined by the operation's symbol — " × " Product, " − " Difference, " ÷ " Ratio, " ? " otherwise.

Structural twin of Predictor_Columns: identical iteration, skip, and interaction gating, emitting headers instead of columns. An interaction header joins the two operands' own names with its operation's symbol (GDP × Schooling).

```excel
=LAMBDA(LET(n_c,COLUMNS(Source_Data),rl,TAKE(Spec_Role,n_c),inc,TAKE(Spec_Include,n_c),typ,TAKE(Spec_Type,n_c),refs,TAKE(Spec_Reference,n_c),trn,TAKE(Spec_Transform,n_c),itm,TAKE(Spec_Interaction_Term,n_c),iop,TAKE(Spec_Interaction_Operation,n_c),hdrs,TOROW(Header_Names),mate,LAMBDA(j,LET(t,INDEX(itm,j),o,INDEX(iop,j),q,IFERROR(XMATCH(t,hdrs),0),IF(OR(LEN(t&"")=0,LEN(o&"")=0,q=0),0,IF(INDEX(rl,q)<>"Predictor (x)",0,q)))),keep,LAMBDA(x,arr,IF(INDEX(typ,x)<>"Categorical",TRUE,NOT(ISNA(INDEX(arr,1,1))))),seed,"",si,Log_Drop_Sample_Include_Calc(),blk,LAMBDA(x,IF(INDEX(typ,x)<>"Categorical",IF(OR(INDEX(trn,x)="Log",INDEX(trn,x)="Log (drop ≤ 0)"),"Ln("&INDEX(hdrs,1,x)&")",INDEX(hdrs,1,x)),LET(col,INDEX(Source_Data,0,x),d,INDEX(refs,x),r,IF(LEN(d&"")=0,"",d),lv,Dummy_Levels(col,r,si),IF(ISNA(INDEX(lv,1,1)),NA(),INDEX(hdrs,1,x)&": "&lv)))),built,REDUCE(seed,SEQUENCE(n_c),LAMBDA(acc,j,IF(OR(INDEX(rl,j)<>"Predictor (x)",INDEX(inc,j)<>TRUE),acc,LET(a,blk(j),IF(NOT(keep(j,a)),acc,LET(m,HSTACK(acc,a),q,mate(j),IF(q=0,m,LET(b,blk(q),IF(NOT(keep(q,b)),m,LET(o,INDEX(iop,j),REDUCE(m,SEQUENCE(COLUMNS(a)),LAMBDA(p,ai,REDUCE(p,SEQUENCE(COLUMNS(b)),LAMBDA(pp,bi,HSTACK(pp,INDEX(a,0,ai)&SWITCH(o,"Product"," × ","Difference"," − ","Ratio"," ÷ "," ? ")&INDEX(b,0,bi)))))))))))))))),DROP(built,,1)))
```

## `Constructed_Column_Transforms` *(sheet-scoped: Regression)*

**Which constructed columns are natural-logged ('Log') versus untouched ('None') — one flag per X_s() column, with every dummy column from a Categorical Predictor always 'None'.**

Sheet-scoped Transform-flag strip for the Model Construction sheet: the third structural twin of X and Constructed_Column_Names (same iteration, same Predictor-AND-Include predicate, same Dummy_Levels skip conditions). Where Constructed_Column_Names answers 'what is this constructed column called', this answers 'was this constructed column logged' — one flag per column, in constructed-column space rather than spec-row space.

That distinction is the whole reason this function exists rather than a caller reading Spec_Transform directly: a Categorical Predictor contributes a variable number of dummy columns (Dummy_Levels' retained-level count, not one column per spec row), so a spec-row-indexed Transform vector cannot align with X()'s output. This closure re-derives the flag in the constructed column's own space — a Continuous column reads 'Log' or 'None' from its row's Transform cell; every dummy column from a Categorical Predictor reads 'None' unconditionally (EXPAND'd across the retained-level width), because Log is disallowed on Categorical Predictors and must never leak into the design matrix even if the cell happens to hold 'Log' (flagged red on the sheet, but still computationally inert here).

Consumed by the Prediction Inputs band (write_sheet_regression.py) to decide, per predictor, whether the user-typed raw prediction value needs Ln_Positive applied before it reaches Group_Prediction_Interval, and to compute the Training Mean spill in the same (possibly logged) space X() itself uses. The retained-levels EXPAND is inlined, never LET-bound, for the same eager-empty-array reason X()'s own dummy branch inlines its broadcast.

Returns: "Log" or "None" per constructed column, same order and width as Predictor_Columns(). Every dummy column from a Categorical Predictor reads "None", and so does every interaction column — the transform lives on each operand's own column and is applied before the two are combined.

Per-constructed-column Transform flags (Log/None), aligned to X() (twin of X/Constructed_Column_Names). Categorical dummy columns always read None regardless of their spec row Transform cell. Feeds Prediction Inputs auto-log and Training Mean.

```excel
=LAMBDA(LET(n_c,COLUMNS(Source_Data),rl,TAKE(Spec_Role,n_c),inc,TAKE(Spec_Include,n_c),typ,TAKE(Spec_Type,n_c),refs,TAKE(Spec_Reference,n_c),trn,TAKE(Spec_Transform,n_c),itm,TAKE(Spec_Interaction_Term,n_c),iop,TAKE(Spec_Interaction_Operation,n_c),hdrs,TOROW(Header_Names),mate,LAMBDA(j,LET(t,INDEX(itm,j),o,INDEX(iop,j),q,IFERROR(XMATCH(t,hdrs),0),IF(OR(LEN(t&"")=0,LEN(o&"")=0,q=0),0,IF(INDEX(rl,q)<>"Predictor (x)",0,q)))),keep,LAMBDA(x,arr,IF(INDEX(typ,x)<>"Categorical",TRUE,NOT(ISNA(INDEX(arr,1,1))))),seed,"",si,Log_Drop_Sample_Include_Calc(),blk,LAMBDA(x,IF(INDEX(typ,x)<>"Categorical",IF(OR(INDEX(trn,x)="Log",INDEX(trn,x)="Log (drop ≤ 0)"),"Log","None"),LET(col,INDEX(Source_Data,0,x),d,INDEX(refs,x),r,IF(LEN(d&"")=0,"",d),lv,Dummy_Levels(col,r,si),IF(ISNA(INDEX(lv,1,1)),NA(),EXPAND("None",1,COLUMNS(lv),"None"))))),built,REDUCE(seed,SEQUENCE(n_c),LAMBDA(acc,j,IF(OR(INDEX(rl,j)<>"Predictor (x)",INDEX(inc,j)<>TRUE),acc,LET(a,blk(j),IF(NOT(keep(j,a)),acc,LET(m,HSTACK(acc,a),q,mate(j),IF(q=0,m,LET(b,blk(q),IF(NOT(keep(q,b)),m,LET(o,INDEX(iop,j),HSTACK(m,EXPAND("None",1,COLUMNS(a)*COLUMNS(b),"None")))))))))))),DROP(built,,1)))
```

## `Context_DF_Absorbed`

**Read element 2 of the model context array.**

Arguments:

- **Context** — the fixed-effects absorbed degrees of freedom - element 2 of Model_Context()

Accessor for element 2 of the v3.0 [Context] array. Routing every context read through this one-line function means the Model_Context row order is a contract enforced in one place rather than 32 hard-coded positional indices.

Returns: Element 2 of the Model_Context() 4x1 array.

Part of the v3.0 context row-order contract (append only, never insert).

```excel
=LAMBDA(Context,
  INDEX(Context, 2)
)
```

## `Context_Has_Intercept`

**Read element 1 of the model context array.**

Arguments:

- **Context** — the intercept flag (column-1 identity, not a synthesize switch) - element 1 of Model_Context()

Accessor for element 1 of the v3.0 [Context] array. Routing every context read through this one-line function means the Model_Context row order is a contract enforced in one place rather than 32 hard-coded positional indices.

Returns: Element 1 of the Model_Context() 4x1 array.

Part of the v3.0 context row-order contract (append only, never insert).

```excel
=LAMBDA(Context,
  INDEX(Context, 1)
)
```

## `Context_Predictor_Transform`

**Read element 4 of the model context array.**

Arguments:

- **Context** — the predictor transform summary (None/Log/Mixed) - element 4 of Model_Context()

Accessor for element 4 of the v3.0 [Context] array. Routing every context read through this one-line function means the Model_Context row order is a contract enforced in one place rather than 32 hard-coded positional indices.

Returns: Element 4 of the Model_Context() 4x1 array.

Part of the v3.0 context row-order contract (append only, never insert).

```excel
=LAMBDA(Context,
  INDEX(Context, 4)
)
```

## `Context_Response_Transform`

**Read element 3 of the model context array.**

Arguments:

- **Context** — the response transform summary (None/Log) - element 3 of Model_Context()

Accessor for element 3 of the v3.0 [Context] array. Routing every context read through this one-line function means the Model_Context row order is a contract enforced in one place rather than 32 hard-coded positional indices.

Returns: Element 3 of the Model_Context() 4x1 array.

Part of the v3.0 context row-order contract (append only, never insert).

```excel
=LAMBDA(Context,
  INDEX(Context, 3)
)
```

## `Cooks_Distance`

**How much one row can pull the fitted model around.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

D_i = r_i² · h_i / ((1−h_i)·p) where r_i is the internally studentized residual, h_i the hat-matrix diagonal, and p the number of regression parameters. Combines outlier signal (r_i) with leverage (h_i) into a single influence measure. Observations with D_i > 4/n or D_i > 1 are conventionally flagged as influential.

Returns: n-element vector of Cook’s distances

Influence measure Dᵢ = rᵢ²·hᵢ / ((1−hᵢ)·p). Values > 4/n or > 1 flag influential rows. Optional DF_Absorbed (default 0) under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    r, Studentized_Residuals(X, Y, filt_arg, context_arg),
    h, Hat_Diagonal(X, filt_arg),
    p, COLUMNS(Design_Matrix(X, filt_arg)),
    r ^ 2 * h / ((1 - h) * p)
  ))
```

## `Correlation_Matrix`

**The pairwise correlations among all predictor columns.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Correlation_Matrix returns the k×k Pearson correlation matrix for all columns of Predictors after filtering. Each cell [r, c] equals CORREL(column r, column c). The diagonal is always 1. The matrix is symmetric. Enter as a single formula and let it spill.

Use this to screen for multicollinearity before fitting a model: pairs with |r| > 0.7 are candidates for VIF investigation. High off-diagonal correlations indicate redundant predictors whose coefficients will be unstable and hard to interpret. For formal post-fit diagnostics, use VIF and Tolerance instead.

Returns: k×k symmetric matrix of pairwise Pearson correlations (diagonal = 1)

k×k symmetric Pearson correlation matrix among all Predictors columns. Use to screen for multicollinearity before fitting.

```excel
=LAMBDA(Predictors, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    xf, FILTER(Predictors, filt_arg),
    k,  COLUMNS(Predictors),
    MAKEARRAY(k, k, LAMBDA(r, c, CORREL(CHOOSECOLS(xf, r), CHOOSECOLS(xf, c))))
  ))
```

## `Data_Completeness`

**Returns TRUE when every predictor value in the row is a usable number.**

Arguments:

- **predictor_row** — single-row range covering the predictor columns for one data row — typically a structured table row reference such as Table[@[FirstCol]:[LastCol]]

Data_Completeness tests whether a single data row has usable numeric values in every predictor column. It applies ISNUMBER to each cell in predictor_row and reduces the results with AND, returning TRUE only when no cell is blank, text, or an error.

Designed to be entered as a calculated table column formula — e.g. =Data_Completeness(Table[@[First]:[Last]]) — so that every row carries its own TRUE/FALSE completeness flag. That column can then be used directly as the Regression_Sample_Include argument to any regression library function, restricting the model to rows with complete predictor data.

Contrast with Complete_Cases_Filter, which operates on full column ranges and returns a boolean vector for all rows at once.

Returns: TRUE if every cell in predictor_row is numeric, FALSE if any value is blank, text, or an error

Returns TRUE when every predictor cell in one table row is numeric. Use as a calculated table column and pass as Include.

```excel
=LAMBDA(predictor_row,
  AND(ISNUMBER(predictor_row))
)
```

## `Demean_By`

**A column with each group's average subtracted out — the within-group deviation that Fixed Effects fits on instead of the raw values.**

Arguments:

- **x** — single-column numeric data range to demean
- **group** — single-column group identifiers — the one-way Fixed Effects panel unit
- **include** — boolean mask — TRUE keeps the row when computing each group's mean; when omitted, numeric x rows are used

Demean_By is the one-way within transformation: x_ig − x̄_g, where x̄_g is Group_Mean(x, group, include). It is the v2.1 Fixed Effects constructor internal — the entire design matrix and the Response are demeaned by the Fixed Effects group before fitting, which is algebraically equivalent to (and numerically more stable than) fitting a full dummy variable per group (LSDV) — the group intercepts are absorbed rather than estimated.

Delegates group-mean lookup to Group_Mean — the same one-source-of-truth pattern Dummy_Code uses for Dummy_Levels — so a caller can never see Demean_By and Group_Mean disagree about a group's mean. #N/A propagates from either a non-numeric x or an undefined group mean (blank group, or a group with no included members), the same NA()-based error contract used throughout the catalog: a genuine Excel error every downstream IFERROR/ISNA guard can catch, never a silently-wrong zero.

With a single group spanning the whole included sample (G = 1 — no Fixed Effects row declared, or a degenerate one-level FE variable), Demean_By collapses to ordinary grand-mean centering — the same degenerate-collapse property the v2.1 prediction and df-plumbing design relies on to build the FE machinery once and have it specialize correctly to the no-FE case.

Returns: n × 1 column, row-aligned to x: x_ig − Group_Mean(x, group, include); #N/A when x is non-numeric or the row's group mean is undefined.

One-way within transformation x_ig − Group_Mean(x,group,include). v2.1 Fixed Effects constructor internal: demeaning by the FE group is algebraically equivalent to LSDV without materializing group dummy columns. G=1 collapses to grand-mean centering.

```excel
=LAMBDA(x, group, [include],
    LET(
        x_v, TOCOL(IF(x = "", "", x), 0),
        gm,  Group_Mean(x, group, include),
        IF(ISNUMBER(x_v) * ISNUMBER(gm), x_v - gm, NA())
    )
)
```

## `Dependent_Variable`

**The filtered outcome column in the order the model uses it.**

Arguments:

- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Returns the dependent-variable vector used by residual and diagnostic outputs.

Returns: filtered dependent variable vector (all rows when Include is omitted)

Filtered Y column in the order the model uses it. Companion to Predictions and Residuals for scatter plots.

```excel
=LAMBDA(Y, [Include],
  FILTER(Y, IF(ISOMITTED(Include), TRUE, Include))
)
```

## `Descriptive_Statistics`

**All 12 standard summary statistics for a data column, returned as a column vector.**

Arguments:

- **data** — single-column data range
- **filter** — optional boolean array — TRUE includes the row, FALSE excludes it; defaults to ISNUMBER(data)

Descriptive_Statistics returns a 12-element column vector of summary statistics for a single data column. When no filter is supplied, non-numeric cells are automatically excluded via ISNUMBER. Mode returns NA() when no unique mode exists. Skewness and kurtosis reuse the existing Skewness and Kurtosis LAMBDAs (Excel SKEW and KURT bias-correction formulas).

Use INDEX to extract individual values: INDEX(Descriptive_Statistics(data), 1) = mean, INDEX(..., 4) = SD, etc. Or enter as a single spill formula to populate all 12 rows at once.

Returns: 12×1 column vector: mean, median, mode, SD, variance, min, max, range, skewness, kurtosis, count, missing count

12×1 column vector: mean, median, mode, SD, variance, min, max, range, skewness, kurtosis, count, missing count.

```excel
=LAMBDA(data, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    n, COUNT(x),
    skew_val, INDEX(Skewness(data, filt_arg), 1),
    kurt_val, INDEX(Kurtosis(data, filt_arg), 1),
    VSTACK(
      AVERAGE(x),
      MEDIAN(x),
      IFERROR(MODE.SNGL(x), NA()),
      STDEV.S(x),
      VAR.S(x),
      MIN(x),
      MAX(x),
      MAX(x) - MIN(x),
      skew_val,
      kurt_val,
      n,
      Missing_Count(data, IF(ISOMITTED(filter), TRUE, filter))
    )
  ))
```

## `Design_Columns` *(sheet-scoped: Regression)*

**The design matrix the model actually fits: unchanged with no Fixed Effects, or with each predictor column's group average subtracted out when a Fixed Effects variable is declared.**

Sheet-scoped reader over the materialized fit-time design matrix for the Regression sheet. Reads the ONE §4b spill via Fit_Design_Columns(), so any call site — the standalone user-callable layer, and any cell that writes =Design_Columns() — reads the single already-computed matrix instead of re-running the column-by-column REDUCE+HSTACK build. The computational body (the no-FE pass-through, the FE demeaning, the direct Allow_Intercept read) lives in Design_Columns_Calc, which is what the spill-source cell calls; pointing this name at the spill is therefore NOT self-referential — the producing cell no longer calls Design_Columns().

This is the v3.2 name-promotion, the predictor-side twin of Sample_Include's. The engine call sites already read Fit_Design_Columns() directly, so this reader is what makes the public/standalone-facing name cheap to call; it delegates to the same Fit_Design_Columns thunk rather than holding the spill's cell address itself, because the catalog cannot spell an A1 address (only the writer knows the anchor from its _C_* layout constants).

Returns: The fit-time design matrix: X() unchanged when no Fixed Effects row is declared, or each of its columns one-way within-transformed (group-demeaned) when one is; full height, same width as X().

Reader over the materialized design-matrix spill (Fit_Design_Columns()). Computation lives in the Design_Columns_Calc leaf; the spill cell calls _Calc, so this name is not self-referential. Engines read Fit_Design_Columns() directly.

```excel
=LAMBDA(Fit_Design_Columns())
```

## `Design_Columns_Calc` *(sheet-scoped: Regression)*

**The design matrix the model actually fits: unchanged with no Fixed Effects, or with each predictor column's group average subtracted out when a Fixed Effects variable is declared.**

The computational core of the fit-time design matrix — what the materialization spill cell (=Design_Columns_Calc()) produces. Design_Columns() does not call this leaf; it reads the produced spill via Fit_Design_Columns() (see Design_Columns's own description). The predictor-side counterpart to Design_Response. With no declared Fixed Effects row it is X() unchanged; with one declared, a REDUCE+HSTACK builds a new matrix column-by-column, demeaning each constructed column of X() by the Fixed Effects group via Demean_By, reproducing the LSDV fit's coefficients on the non-FE predictors without ever materializing the G−1 group dummy columns an explicit one-hot encoding would require. (Not BYCOL: Demean_By returns a whole column, and BYCOL's callback contract requires a single scalar per call — the same REDUCE+HSTACK pattern X() itself uses for column-by-column assembly.)

Shipped alongside Design_Response, for the identical reason: X() has raw-predictor consumers that must not see demeaned values — the Predictor Summary zone (Pearson_R, Spearman_R, Skewness, Kurtosis, GVIF, Generalized_Tolerance) characterizes each predictor's own distribution, and Constructed_Column_Names() labels columns positionally against X(), not Design_Columns(). Only the fit/inference/prediction/residual chain (Coefficients, Predictions, Residuals, Hat_Diagonal, Cooks_Distance, and everything built from them) reads Design_Columns() (via Fit_Design_Columns()) — repointing which columns those formulas READ, not a change to what X() itself returns.

The no-FE branch returns X() itself, unchanged, so every existing no-FE model is unaffected by construction — including its error behavior: a zero-predictor spec that makes X() itself error propagates through Design_Columns_Calc() identically, never masked. The constructor reads the Allow_Intercept toggle directly (not via Fit_Context): a large spill depending on another spill via a thunk is not resolved reliably by per-sheet Worksheet.Calculate — see project_v3_stage2_gate_fails.

Returns: The fit-time design matrix: X() unchanged when no Fixed Effects row is declared, or each of its columns one-way within-transformed (group-demeaned) when one is; full height, same width as X().

Fit-time design-matrix _Calc leaf: X() unchanged (no FE), else each column demeaned by Fixed_Effects_Column() via Demean_By (REDUCE+HSTACK). Reads Allow_Intercept directly, not Fit_Context. Predictor Summary stays on raw X().

```excel
=LAMBDA(LET(
    n_c,       COLUMNS(Source_Data),
    fe_active, SUMPRODUCT(N(TAKE(Spec_Role, n_c) = "Fixed Effects")) > 0,
    has_int,   N(Allow_Intercept) = 1,
    ones,      SEQUENCE(ROWS(Source_Data), 1, 1, 0),
    k_p,       IFERROR(COLUMNS(Predictor_Columns()), 0),
    IF(k_p = 0,
       IF(has_int, ones, Predictor_Columns()),
       LET(
           demeaned, IF(NOT(fe_active),
                        Predictor_Columns(),
                        LET(
                            fe,    Fixed_Effects_Column(),
                            inc,   Log_Drop_Sample_Include_Calc(),
                            xp,    Predictor_Columns(),
                            seed,  SEQUENCE(ROWS(xp), 1, 0, 0),
                            built, REDUCE(seed, SEQUENCE(k_p), LAMBDA(acc, j,
                                HSTACK(acc, Demean_By(INDEX(xp, 0, j), fe, inc))
                            )),
                            DROP(built, , 1)
                        )),
           IF(has_int, HSTACK(ones, demeaned), demeaned)
       )
    )
))
```

## `Design_Matrix`

**The cleaned predictor matrix the regression actually fits.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Returns the design matrix restricted to the included rows — FILTER(X, filt_arg).

It no longer synthesizes an intercept column. Since v3.0 the constructor (Design_Columns on the Regression sheet) owns every construction stage, including the intercept, so by the time a design matrix reaches an engine function its intercept column is already there. What remains here is the row mask, applied in the engine so that constructors can stay full-height and row-aligned with the source table.

Returns: filtered numeric design matrix as spilled array

The design matrix restricted to the included rows. Does not add an intercept column — the constructor owns that stage from v3.0.

```excel
=LAMBDA(X, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    FILTER(X, filt_arg)
  ))
```

## `Design_Response` *(sheet-scoped: Regression)*

**The response the model actually fits: unchanged with no Fixed Effects, or with each group's average subtracted out when a Fixed Effects variable is declared.**

Design_Response is the fit-time Response for the Regression sheet — the v2.1 Fixed Effects counterpart to X on the predictor side. With no declared Fixed Effects row it is Response_Column() unchanged (the entire v2.0 shipped surface passes through untouched); with one declared, it is Demean_By(Response_Column(), Fixed_Effects_Column(), Log_Drop_Sample_Include_Calc()) — the one-way within transformation that lets ordinary OLS on the demeaned data reproduce the LSDV (dummy-per-group) fit without materializing G−1 group dummy columns.

Shipped as a NEW function, deliberately not a replacement of Response_Column() at its existing call sites: Response_Column() keeps its raw-response consumers that are unrelated to this fit (Intercept_Only_Point/S/SE/N, the zero-predictor closed-form branch). The Residual Output table's Y column is repointed to Design_Response() too, not left raw — under FE the whole row-level table (Y, Predicted Y, Residuals, Hat Diagonal, …) must read as one internally consistent block in the same fit space, or Residuals would visibly fail to match Y − Predicted Y. Only the fit/inference/prediction/residual chain (Coefficients, Predictions, Residuals, and everything built from them) is repointed to Design_Response() — a caller-side choice, not a change to what Response_Column() means.

The no-FE branch is checked first and returns the exact same object Response_Column() would — not merely a numerically-equal recomputation — so every existing no-FE model is unaffected by construction, the same non-breaking-default property Absorbed_Degrees_Of_Freedom relies on for the df side.

Returns: The fit-time Response: Response_Column() unchanged when no Fixed Effects row is declared, or its one-way within-transformed (group-demeaned) form when one is; full height.

Fit-time Response: Response_Column() unchanged (no FE), else Demean_By(Response_Column(),Fixed_Effects_Column(),Log_Drop_Sample_Include_Calc()). New function — the zero-predictor Intercept_Only_* branch stays on raw Response_Column().

```excel
=LAMBDA(LET(
    n_c,       COLUMNS(Source_Data),
    fe_active, SUMPRODUCT(N(TAKE(Spec_Role, n_c) = "Fixed Effects")) > 0,
    IF(NOT(fe_active),
       Response_Column(),
       Demean_By(Response_Column(), Fixed_Effects_Column(), Log_Drop_Sample_Include_Calc())
    )
))
```

## `Design_Width_Status` *(sheet-scoped: Regression)*

**Warns before the design matrix gets too wide to compute — read off the spec, before any matrix is built.**

Design_Width_Status is the O2 readout: the pre-flight design-matrix width guard. Both thresholds are computed FROM THE SPEC, never from COLUMNS(Design_Columns()) — a matrix too wide to fit cannot be built in order to be measured, which is the failure this guard exists to prevent.

The hard limit is empirical: the point where Excel cannot reliably invert the Gram matrix (Gram_Inverse is O(k^3) in MMULT, and MINVERSE on X'X squares the condition number). The soft limit fires earlier on either k or n×k, warning that recalculation is getting slow — every materialized cell recalculates on any input change.

It recomputes k rather than reading the O1 audit total, because a catalog body cannot import the sheet's column constants and hard-coding an A1 address is what turns a column insertion into a silent-wrong-answer bug. The expression is the same one O1 holds, so the number is identical; the two cells are simply independent, and an error in O1 no longer propagates into this verdict.

The thresholds are duplicated from regression_layout.py, which also uses the soft column count to size the design-matrix band. Python remains the source of truth and a unit test pins this body's numbers to it.

SHEET-SCOPED (scope: Regression).

Returns: An ERROR string above the hard design-matrix column limit, a WARNING at the soft column or materialized-cell threshold, otherwise "".

O2 pre-flight width guard. ERROR above the hard Gram-inversion column limit, WARNING at the soft column or n*k cell threshold, else "". Computed from the spec, never from the built matrix. Sheet-scoped.

```excel
=LAMBDA(
    LET(k,SUM(TAKE(Spec_Design_Columns,COLUMNS(Source_Data)))+N(Allow_Intercept),n,ROWS(Source_Data),IF(k>200,"ERROR: the design matrix has "&k&" constructed columns; the Gram matrix cannot be reliably inverted above 200 columns in Excel (MINVERSE on X'X squares the condition number). Reduce a Categorical predictor's levels, exclude predictors, or group them.",IF(OR(k>100,n*k>200000),"WARNING: "&k&" constructed columns x "&n&" rows = "&(n*k)&" materialized cells. Recalculation is O(k^3) in MINVERSE, and every materialized cell recalculates on any input change. Approaching the Gram inversion limit.","")))
)
```

## `Difference_By`

**The change in a column since the group's previous time period — #N/A at each group's first period and across gaps, so no fake zeros or multi-year jumps sneak into a model.**

Arguments:

- **x** — single-column numeric data range to difference
- **group** — single-column group identifiers (e.g. Country) — differences never cross a group boundary
- **seq** — single-column numeric time index (e.g. Year) — the lookup key, not the row order
- **delta** — base period Δ: Δx(t) = x(t) − x(t − delta) by exact time value; when omitted, the spec's Base Period Δ (Base_Period_Delta()) is used — never a silent 1
- **include** — boolean mask — TRUE keeps the row; when omitted, rows with a non-blank group and numeric seq are used

Difference_By returns the within-group time difference Δx(i,t) = x(i,t) − x(i, t−Δ), where Δ is the base period. Three semantic commitments, in order of how expensively they fail when violated:

(A) Each group's first period returns #N/A — never a fabricated 0. A zero would enter a design matrix silently as a real observation; #N/A is caught by every downstream ISNUMBER/IFERROR mask.

(B) "Prior period" means the literal time value t − Δ. If t − Δ is absent within the group (a gap), the result is #N/A. The function never falls back to the previous available row: that would inject a spurious multi-period jump and present it as a one-Δ change. Lookup is by time value, not row position, by construction — the prior value comes from Lag_By's exact-match XLOOKUP of (group, seq − Δ) composite keys, not from OFFSET/row arithmetic — so the result is invariant to how the rows happen to be sorted.

(C) Δ is computed-and-displayed-with-override, never silently assumed to be 1. An omitted delta resolves through Base_Period_Delta() — the visible Base Period Δ spec cell — and returns #N/A everywhere when no sequence axis or Δ is declared. Same pattern as the categorical reference level.

The prior value is delegated to Lag_By (one source of truth for the pair lookup, as Dummy_Code delegates to Dummy_Levels). A row whose own x, or whose t−Δ partner's x, is blank or non-numeric yields #N/A: an incomputable difference on an included row is a visible #N/A observation, not a silently dropped one — which is why the default include mask requires only a non-blank group and numeric seq, deliberately NOT numeric x. Rows failing the include mask return "" so the spilled column stays row-aligned with the source.

Returns: n × 1 column, row-aligned to x: Δx(i,t) = x(i,t) − x(i,t−Δ); #N/A when the difference is not computable (first period, gap, or a non-numeric value); "" on excluded rows.

Within-group difference x(t) − x(t−Δ) by exact time value. First period or gap → #N/A (never 0, never the previous available row). Omitted delta reads the spec's Base Period Δ (never a silent 1). Excluded rows return "".

```excel
=LAMBDA(x, group, seq, [delta], [include],
    LET(
        x_v,  TOCOL(IF(x = "", "", x), 0),
        g_v,  TOCOL(IF(group = "", "", group), 0),
        t_v,  TOCOL(IF(seq = "", "", seq), 0),
        step, IF(ISOMITTED(delta), Base_Period_Delta(), delta),
        inc,  IF(ISOMITTED(include), (g_v <> "") * ISNUMBER(t_v), include * 1),
        prior, Lag_By(x, group, seq, step, inc),
        IF(inc, IF(ISNUMBER(x_v) * ISNUMBER(prior), x_v - prior, NA()), "")
    )
)
```

## `Dummy_Code`

**Treatment-coded dummy matrix for a categorical predictor, with excluded rows as blank and the reference level dropped.**

Arguments:

- **category** — single-column categorical data range
- **reference** — the level to drop (treatment coding); when omitted or "", the first sorted level is dropped
- **include** — boolean mask — TRUE keeps the row; when omitted all non-blank rows are used

Dummy_Code returns a treatment-coded dummy matrix for a categorical variable. For each included row, the value is compared to each retained level and encoded as 1 (match) or 0 (no match). Excluded rows (include = FALSE or blank category value) return "" across all k columns so the spilled array stays row-aligned with the source data.

Level determination is delegated to Dummy_Levels — one source of truth for which levels are retained — so the two functions can never disagree about the level set. The reference level is dropped to avoid perfect multicollinearity when the design includes an intercept. Reference-level validation is enforced: if the supplied reference does not appear in the included sample, the function returns #N/A rather than silently failing to drop a column — an invalid reference reintroduces the exact collinearity the function exists to prevent.

Pair with Dummy_Levels to get the matching column headers. Pass the output directly to HSTACK to assemble a multi-predictor design for Coefficients, Prediction_Interval, and related functions.

Failure is signaled by type, not by text: degenerate inputs return a genuine Excel error (#N/A via NA()) that every downstream IFERROR/ISNA guard can catch, never a descriptive string that would silently participate in array math. The three failure conditions — an included sample with no non-blank category values, a reference level not present in the included sample, and a sample containing only one level (nothing left to code after the reference is dropped) — all return #N/A.

Returns: n × k matrix of 0/1 dummy indicators aligned to every source row; excluded rows return "" in all columns; #N/A on degenerate input

Treatment-coded 0/1 dummy matrix, row-aligned to the source data. Levels come from Dummy_Levels (one source of truth); the reference level is dropped to avoid collinearity with the intercept; excluded rows return "". Returns #N/A on degenerate input.

```excel
=LAMBDA(category, [reference], [include],
    LET(
        inc,    IF(ISOMITTED(include), TRUE, include),
        x,      TOCOL(IF(category = "", "", category), 0),
        active, (x<>"") * inc,
        lv,     Dummy_Levels(category, reference, include),
        IF(ISNA(lv),
           NA(),
           MAKEARRAY(ROWS(x), COLUMNS(lv), LAMBDA(r, c,
               IF(INDEX(active, r), --(INDEX(x, r) = INDEX(lv, 1, c)), "")
           ))
        )
    )
)
```

## `Dummy_Column`

**One indicator column for a single named level — 1 where category equals it, 0 elsewhere; no treatment coding, no reference validation.**

Arguments:

- **category** — single-column categorical data range
- **level** — the single level to indicator — 1 where category equals this level, 0 otherwise
- **include** — boolean mask — TRUE keeps the row; when omitted all non-blank rows are used

Dummy_Column returns a single indicator (dummy) column for one explicitly named level: 1 where the category equals that level, 0 elsewhere. It is the free-form counterpart to the spec-driven categorical path — the constructor encodes inline via broadcast and delegates level determination to Dummy_Levels/Dummy_Code, but a standalone user building their own design matrix needs the raw primitive: one column, one level, no treatment-coding machinery attached.

Unlike Dummy_Levels and Dummy_Code, Dummy_Column imposes no treatment-coding semantics and no reference validation. A level not present in the included sample yields a column of 0s — a valid but useless indicator, NOT #N/A. The treatment-coding functions return #N/A on an invalid reference because a bad reference reintroduces the exact collinearity they exist to prevent; an indicator column for a level that simply isn't observed is harmless (all zeros), so there is nothing to signal. The caller names the level and takes responsibility for it.

Blank category cells are normalized to "" and excluded by the default include mask (active = (x <> "") * inc), the same blank-exclusion convention Dummy_Levels/Dummy_Code use, so a blank never coerces to a spurious match against a level named "". Excluded rows return "" so the spilled column stays row-aligned with the source — the standard row-aligned-transform contract every other transform in the catalog honours.

Returns: n × 1 column, row-aligned to category: 1 where the row's category equals level, 0 where it does not (for included rows); "" on excluded rows and on blank category cells.

Single indicator column for one explicit level: 1 where category = level, else 0. No treatment coding, no reference validation — a missing level is an all-0 column, not #N/A. Free-form counterpart to the spec-driven path. Excluded rows return "".

```excel
=LAMBDA(category, level, [include],
    LET(
        inc,    IF(ISOMITTED(include), TRUE, include),
        x,      TOCOL(IF(category = "", "", category), 0),
        active, (x <> "") * inc,
        IF(active, --(x = level), "")
    )
)
```

## `Dummy_Levels`

**The k category labels that become dummy columns (sorted, reference level removed).**

Arguments:

- **category** — single-column categorical data range
- **reference** — the level to drop (treatment coding); when omitted or "", the first sorted level is dropped
- **include** — boolean mask — TRUE keeps the row; when omitted all non-blank rows are used

Dummy_Levels returns the set of category labels that will become dummy columns after treatment coding. It sorts all levels present in the included sample, drops the reference level, and returns the remainder as a horizontal row vector.

When reference is omitted or given as an empty string "", the first sorted level is used as the reference (alphabetically first) — so a blank reference cell passed straight through means "use the default". When any other reference is supplied, it is validated against the included sample; if not found, the function returns #N/A rather than silently failing to drop a column.

Use Dummy_Levels as a header row above Dummy_Code to label dummy columns, or to inspect which levels the model will contain before fitting.

Truly empty cells are normalized to "" before the row mask is applied, so a blank category cell is excluded rather than coerced to a spurious numeric-zero level; genuine 0 values pass through untouched.

Failure is signaled by type, not by text: degenerate inputs return a genuine Excel error (#N/A via NA()) that every downstream IFERROR/ISNA guard can catch, never a descriptive string that would silently participate in array math. The three failure conditions — an included sample with no non-blank category values, a reference level not present in the included sample, and a sample containing only one level (nothing left to code after the reference is dropped) — all return #N/A. The human-readable explanation is carried by the consuming sheet's conditional formatting, not by the return value.

Returns: 1 × k horizontal vector of retained category labels (sorted, reference level removed); #N/A on degenerate input

Retained category labels after treatment coding: sorted levels from the included sample with the reference level dropped. Use as a header row above the Dummy_Code matrix. Returns #N/A (a real Excel error, catchable by ISNA/IFERROR) on degenerate input.

```excel
=LAMBDA(category, [reference], [include],
    LET(
        inc,        IF(ISOMITTED(include), TRUE, include),
        x,          TOCOL(IF(category = "", "", category), 0),
        active,     (x<>"") * inc,
        x_active,   FILTER(x, active, NA()),
        All_Levels, IFERROR(TOROW(SORT(UNIQUE(x_active))), NA()),
        ref_raw,    IF(ISOMITTED(reference), "", reference),
        ref,        IF(ref_raw = "", IFERROR(INDEX(All_Levels, 1, 1), NA()), ref_raw),
        ref_ok,     IFERROR(ISNUMBER(MATCH(ref, All_Levels, 0)), FALSE),
        IF(OR(ISNA(All_Levels), NOT(ref_ok)),
           NA(),
           IFERROR(FILTER(All_Levels, All_Levels <> ref), NA())
        )
    )
)
```

## `Durbin_Watson`

**Whether nearby residuals tend to move together instead of staying independent.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

DW = Σ(eₜ − eₜ₋₁)² / Σeₜ², testing for first-order serial correlation in residuals. Values near 2 indicate no autocorrelation; values below ~1.5 suggest positive autocorrelation, values above ~2.5 suggest negative autocorrelation.

Computed as SUM((DROP(e,1) − DROP(e,−1))²) / SUM(e²), where DROP(e,1) removes the first residual (giving e₂..eₙ) and DROP(e,−1) removes the last (giving e₁..eₙ₋₁), so consecutive differences are formed without OFFSET or index arithmetic.

Returns: Durbin-Watson statistic as a scalar

Test statistic for first-order serial autocorrelation in residuals. Values near 2 = no autocorrelation; < 1.5 or > 2.5 flag problems.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    e,        Residuals(X, Y, filt_arg),
    SUM((DROP(e, 1) - DROP(e, -1))^2) / SUM(e^2)
  ))
```

## `Durbin_Watson_By`

**The Durbin-Watson autocorrelation check, but measured along a declared time/order column instead of the accidental row order.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **seq** — single-column ordering axis (the declared Sequence column, e.g. Year) — residuals are sorted by this before differencing, so the result does not depend on physical row order
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Durbin_Watson_By is the sequence-aware Durbin-Watson statistic: DW = Σ(eₜ − eₜ₋₁)² / Σeₜ², but with the residuals ordered by a declared sequence axis instead of trusting physical row order.

Why this exists: the ordinary Durbin_Watson differences residuals in the order the rows happen to sit in the sheet. That is only meaningful if row adjacency IS the hypothesized serial axis. On arbitrarily sorted data — e.g. rows sorted by GDP — it silently assumes row order = time order and, under a meaningful non-time sort with mild misspecification, reports strong spurious autocorrelation. This function removes that hazard: it computes residuals, sorts them ascending by seq (SORTBY, all internal — the sort never spills to the sheet), and only then forms consecutive differences.

The result is invariant to how the source rows are physically ordered: shuffling the data and recomputing gives the identical statistic, because adjacency is defined by seq, not by row position. seq should uniquely order the observations of a single series (a proper time index); on a pooled panel where seq repeats across groups (many countries sharing a Year), adjacency along seq alone is not the within-group serial axis — that group-aware, seam-safe case is BFN_Panel_Durbin_Watson, not this function.

Output is always a scalar; the DROP/SORTBY array work stays inside the LAMBDA so no helper column or spill is exposed.

Returns: Durbin-Watson statistic as a scalar, computed along seq order rather than row order

Sequence-aware Durbin-Watson: sorts residuals by seq before differencing, so the statistic is invariant to physical row order. Use when a Sequence axis is declared; ordinary Durbin_Watson trusts row order. Scalar output, no spill.

```excel
=LAMBDA(X, Y, seq, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    e,        Residuals(X, Y, filt_arg),
    ord,      FILTER(seq, filt_arg),
    es,       SORTBY(e, ord, 1),
    SUM((DROP(es, 1) - DROP(es, -1))^2) / SUM(es^2)
  ))
```

## `Exclude_Row_N`

**The same array with one chosen row removed.**

Arguments:

- **array** — any contiguous array to drop a row from
- **n** — 1-based row to exclude; negative values count from the end (−1 = last row)

Returns the input array with exactly one row removed. Positive n selects from the top (1 = first row); negative n selects from the bottom (−1 = last row, −2 = second-to-last, …). Returns NA() when the resolved index is less than 1 or greater than ROWS(array).

Implemented as CHOOSEROWS(array, FILTER(SEQUENCE(total_rows), SEQUENCE(total_rows) <> n_abs)): SEQUENCE produces the full index 1..n, FILTER removes the target index, and CHOOSEROWS reassembles the array without that row.

Returns: array with row n removed; NA() if n is out of bounds

Array with one row removed. Positive n counts from top, negative from bottom. Returns NA() when n is out of bounds.

```excel
=LAMBDA(array, n,
  LET(
    total_rows, ROWS(array),
    n_abs,      IF(n < 0, total_rows + n + 1, n),
    IF(
      OR(n_abs < 1, n_abs > total_rows),
      NA(),
      LET(
        rows, SEQUENCE(total_rows),
        CHOOSEROWS(array, FILTER(rows, rows <> n_abs))
      )
    )
  )
)
```

## `F_Statistic`

**Whether the model, taken as a whole, explains more than random noise would.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

F_Statistic returns the F-statistic for the omnibus test of the null hypothesis that all regression coefficients are simultaneously zero. It equals (MS_Regression / MS_Residual) = (SS_Regression / df_regression) / (SS_Residual / df_residual).

A large F-statistic (and correspondingly small F_Statistic_P_Value) means the model explains significantly more variance than would be expected by chance. Note that a significant F does not guarantee that every individual predictor is significant — individual t-statistics (T_Statistics) address that question. F_Statistic appears in the ANOVA table on the Regression sheet; this function exposes it as a named reference for use in other formulas.

An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads only to the residual (denominator) df — the regression (numerator) df is the count of estimated slope coefficients, which absorbed FE groups are not.

Returns: F-statistic for the overall regression model as a scalar

Overall F = MS_Regression / MS_Residual. Optional DF_Absorbed (default 0) corrects only the denominator df under Fixed Effects — the numerator df is unaffected.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg,     IF(ISOMITTED(Include),       TRUE, Include),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    MS_Regression(X, Y, filt_arg, context_arg) / MS_Residual(X, Y, filt_arg, context_arg)
  ))
```

## `F_Statistic_P_Value`

**How surprising the overall F-statistic would be if the model had no real signal.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

F_Statistic_P_Value returns the p-value for the overall model F-test: the probability of observing an F-statistic at least as large as F_Statistic under the null hypothesis that all coefficients are zero. Computed as F.DIST.RT(F_Statistic, df_regression, df_residual).

Conventionally, F_Statistic_P_Value < 0.05 indicates that the model as a whole fits the data significantly better than a null model. This p-value appears in the Significance F column of the Regression sheet ANOVA table; this function exposes it for programmatic use. An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to both the F-statistic and its denominator df.

Returns: p-value for the overall F-test (right-tail probability of the F distribution)

Right-tail p-value for the overall F-test. Optional DF_Absorbed (default 0) corrects both F_Statistic and its denominator df under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg,     IF(ISOMITTED(Include),       TRUE, Include),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    F.DIST.RT(
      F_Statistic(X, Y, filt_arg, context_arg),
      Regression_Degrees_Of_Freedom(X, context_arg),
      Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg)
    )
  ))
```

## `Fixed_Effects_Column` *(sheet-scoped: Regression)*

**The column marked as the Fixed Effects (panel unit) variable in the spec — e.g. Country in a Country/Year panel.**

Sheet-scoped accessor for the panel-unit (grouping) variable on the Regression sheet. Returns the data column of the first spec row whose Role is "Fixed Effects", via exact-match XMATCH over the TAKE-trimmed Spec_Role vector — the Role-axis sibling of Response_Column (Role=Response) and the structural counterpart of Sequence_Column (the H flag). Full-height by the same row-mask contract as the other sheet closures.

It is the group input for the seam-safe panel serial-correlation diagnostic (BFN_Panel_Durbin_Watson), keeping the panel unit a single canonical name rather than an inline XMATCH repeated at every call site. When no row carries the Fixed Effects role, XMATCH does not match and the accessor is #N/A — the caller's gate (the FE-variable count) shows the not-applicable token before this is ever evaluated. When two or more rows carry the role (two-way absorption, out of scope until its own milestone) it resolves the first, and callers gate on the count instead of computing.

Forward wiring for the v2.1 Fixed Effects release: the "Fixed Effects" Role token is not yet offered by the Role dropdown because the design-matrix engine does not absorb it yet — the accessor and the diagnostics gated on it activate when the role ships. Spec_Role is TAKE-trimmed to COLUMNS(Source_Data) so cells below the live spec rows are never scanned.

Returns: The single Role=Fixed Effects data column (the panel-unit identifier), full height; #N/A when no row carries the role.

Derived panel-unit column for the Regression sheet: the data column of the first Role="Fixed Effects" spec row. #N/A when none is declared. Feeds BFN_Panel_Durbin_Watson; forward wiring for the v2.1 Fixed Effects role.

```excel
=LAMBDA(LET(n_c,COLUMNS(Source_Data),INDEX(Source_Data,0,XMATCH("Fixed Effects",TAKE(Spec_Role,n_c),0))))
```

## `Freedman_Diaconis_Bins`

**How many histogram bins the Freedman-Diaconis rule recommends.**

Arguments:

- **data** — single-column numeric data range
- **filter** — optional boolean array

Freedman_Diaconis_Bins computes the bin width h = 2 × IQR × n^(−1/3) where IQR is the interquartile range (QUARTILE.EXC), then returns k = ⌈range / h⌉.

Freedman-Diaconis is robust to outliers because it uses the IQR rather than the standard deviation. It is the preferred rule for skewed or heavy-tailed distributions and for large samples where the shape deviates from normality.

Returns: integer bin count via the Freedman-Diaconis rule

Bin count from Freedman-Diaconis rule: width = 2·IQR·n^(−1/3). Robust to outliers; preferred for skewed data.

```excel
=LAMBDA(data, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    n, COUNT(x),
    iqr, QUARTILE.EXC(x, 3) - QUARTILE.EXC(x, 1),
    width, 2 * iqr * n ^ (-1/3),
    range_, MAX(x) - MIN(x),
    IF(iqr = 0, Sturges_Bins(data, filt_arg), CEILING(range_ / width, 1))
  ))
```

## `Full_Factorial`

**Every N-level combination of d variables, each spanning its own [min, max] range — the full-factorial grid.**

Arguments:

- **N** — number of levels per axis — the grid has N evenly-spaced values along every dimension; N must be ≥ 1 (the MAX(1,N-1) divisor guards the N=1 case, which yields the single point at the minimum)
- **minimums** — lower bound for each dimension — a 1×d row vector or d×1 column vector; TOROW normalizes either to a single row, so orientation does not matter
- **maximums** — upper bound for each dimension — same shape contract as minimums; the c-th entry pairs with the c-th minimum

Full_Factorial generates the complete N-level Cartesian-product grid over a d-dimensional box bounded by minimums and maximums. TOROW normalizes both bound vectors to single rows, so a row or column vector may be passed interchangeably; d is COLUMNS of the normalized minimums, and the two bound vectors must agree in length. MAKEARRAY builds N^d rows by d columns, where the level index for dimension c at row r is i = MOD(QUOTIENT(r-1, N^(d-c)), N) — mixed-radix counting with column 1 the slowest-varying axis and column d the fastest — and the value is the linear interpolation mins[c] + i·(maxs[c]-mins[c])/MAX(1,N-1). With i ranging 0 to N-1, each dimension spans its full [min, max] range in N evenly-spaced points: the N=2 case is the 2^d corners of the box, and larger N densifies every dimension uniformly. The MAX(1,N-1) divisor guards the N=1 case (where N-1 would be zero): N=1 yields a single row at the minimum of every dimension, so the single-dimension case d=1 reduces to SEQUENCE(N,1,min,(max-min)/MAX(1,N-1)) — an N-point linear spacing from minimums to maximums, and a single point when N=1. N must be ≥ 1. This is the general grid the Beta (and any two-parameter) grid search is a specialization of — a 2-D N×N grid is Full_Factorial(N, VSTACK(alpha_min, beta_min), VSTACK(alpha_max, beta_max)).

Returns: N^d × d matrix — one row per level combination, one column per dimension, each dimension evenly spaced from its minimum to its maximum; column 1 is the slowest-varying axis, column d the fastest

N^d × d full-factorial grid: N evenly-spaced levels per dimension between minimums and maximums. Column 1 slowest-varying, column d fastest. N must be ≥ 1; MAX(1,N-1) guards the N=1 single-point case.

```excel
=LAMBDA(N,minimums,maximums,
    LET(
        mins,TOROW(minimums),
        maxs,TOROW(maximums),
        d,COLUMNS(mins),
        MAKEARRAY(
            N^d,
            d,
            LAMBDA(r,c,
                LET(
                    i,MOD(QUOTIENT(r-1,N^(d-c)),N),
                    INDEX(mins,1,c)
                        +i*(INDEX(maxs,1,c)-INDEX(mins,1,c))/MAX(1,N-1)
                )
            )
        )
    )
)
```

## `Generalized_Tolerance`

**How much unique information a variable still has after overlap with the others, generalized to whole categorical predictors.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Column_Names** — 1×k row vector of column labels aligned with X_s, following the "Header: level" convention (e.g. Constructed_Column_Names()); columns sharing the text before the first ": " are treated as one variable's group
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Generalized Tolerance = 1 / GVIF, the reciprocal of the Generalized Variance Inflation Factor. Mirrors Tolerance's own one-line relationship to VIF, forwarding all arguments to GVIF.

Returns: k×1 column vector of Generalized Tolerance values (= 1/GVIF), one per column of Predictors

1/GVIF per source variable. Fraction of unique variance not shared with other predictors, generalized across a categorical predictor's dummy columns.

```excel
=LAMBDA(Predictors, Column_Names, [Include],
  1 / GVIF(Predictors, Column_Names, Include)
)
```

## `GoF_AIC`

**A model-comparison score that balances fit quality against the number of parameters.**

Arguments:

- **nll** — negative log-likelihood from any NLL_* function
- **k** — number of free parameters in the distribution (1 for Exponential, 2 for Normal/Lognormal/Weibull/Gamma/Beta, 3 for Triangular/BetaPERT)

GoF_AIC computes the Akaike Information Criterion: AIC = 2k + 2·NLL. Lower AIC indicates a better balance of fit quality and model complexity. AIC is comparable across all distributions fitted to the same dataset — the distribution with the lowest AIC is the best-fitting model under this criterion.

AIC does not penalise extra parameters as strongly as BIC, so it may favour more complex distributions (larger k) when sample sizes are large. For small samples, use AICc instead (not implemented here — add a correction term 2k(k+1)/(n−k−1) if n/k < 40).

Returns: AIC = 2k + 2·NLL as a scalar

AIC = 2k + 2·NLL for distribution comparison. k = number of free parameters. Lower = better. Compare across distributions on same data.

```excel
=LAMBDA(nll, k, 2 * k + 2 * nll)
```

## `GoF_Anderson_Darling`

**Anderson-Darling A² = −n − (1/n)·Σ(2i−1)[ln F(xᵢ) + ln(1−F(x_{n+1−i}))] (lower = better fit).**

Arguments:

- **data** — single-column data range
- **dist_cdf** — CDF values evaluated at each data point under the fitted distribution (same size as data; e.g. NORM.DIST(data, mean, sd, TRUE))
- **include** — optional boolean array — TRUE keeps the row, FALSE excludes it; defaults to ISNUMBER(data)

GoF_Anderson_Darling computes the Anderson-Darling goodness-of-fit statistic: A² = −n − (1/n) Σᵢ (2i−1)·[ln F(X₍ᵢ₎) + ln(1 − F(X₍ₙ₊₁₋ᵢ₎))]. CDF values are sorted by data internally. An epsilon clamp (1E-10) prevents ln(0) for bounded-support distributions (Beta, Triangular, BetaPERT) whose CDF can reach exactly 0 or 1 at support edges.

A-D is more sensitive to tail deviations than K-S, making it better at detecting departures in the extremes. Lower A² indicates a better fit. Compare alongside AIC, BIC, and K-S in the fit comparison table.

Returns: Anderson-Darling A² statistic as a scalar (lower = better fit)

Anderson-Darling A² goodness-of-fit. More sensitive to tail deviations than K-S. Lower = better. Compare with AIC, BIC, and K-S.

```excel
=LAMBDA(data, dist_cdf, [include],
  LET(
    inc, IF(ISOMITTED(include), ISNUMBER(data), include),
    d, FILTER(data, inc),
    f, FILTER(dist_cdf, inc),
    n, ROWS(d),
    IF(n = 0, NA(),
    LET(
      fs, SORTBY(f, d),
      eps, 1E-10,
      fc, IF(fs < eps, eps, IF(fs > 1 - eps, 1 - eps, fs)),
      i, SEQUENCE(n),
      fc_rev, SORTBY(fc, i, -1),
      -n - SUMPRODUCT((2 * i - 1) * (LN(fc) + LN(1 - fc_rev))) / n
    ))
  ))
```

## `GoF_BIC`

**A model-comparison score that penalises extra parameters more strongly as sample size grows.**

Arguments:

- **nll** — negative log-likelihood from any NLL_* function
- **k** — number of free parameters in the distribution
- **n** — number of observations used in the fit (from UV_n or COUNT(data))

GoF_BIC computes the Bayesian Information Criterion: BIC = k·ln(n) + 2·NLL. Lower BIC indicates a better model. BIC penalises extra parameters more strongly than AIC (penalty grows with log(n) rather than 2), favouring simpler distributions in larger samples.

Compare AIC and BIC together: if they agree, the best-fit distribution is clear; if they disagree, AIC is selecting a more complex distribution and BIC is selecting a simpler one. The difference is informative about how much the extra parameters contribute to fit.

Returns: BIC = k·ln(n) + 2·NLL as a scalar

BIC = k·ln(n) + 2·NLL. Penalises extra parameters more than AIC (penalty scales with log n). Lower = better.

```excel
=LAMBDA(nll, k, n, k * LN(n) + 2 * nll)
```

## `GoF_Kolmogorov_Smirnov`

**Kolmogorov-Smirnov D = max|F_n(xᵢ) − F(xᵢ)| (lower = better fit).**

Arguments:

- **data** — single-column data range
- **dist_cdf** — CDF values evaluated at each data point under the fitted distribution (same size as data; e.g. NORM.DIST(data, mean, sd, TRUE))
- **include** — optional boolean array — TRUE keeps the row, FALSE excludes it; defaults to ISNUMBER(data)

GoF_Kolmogorov_Smirnov computes the Kolmogorov-Smirnov goodness-of-fit statistic: D = max(D⁺, D⁻) where D⁺ = maxᵢ(i/n − F(X₍ᵢ₎)) and D⁻ = maxᵢ(F(X₍ᵢ₎) − (i−1)/n). This is the maximum vertical distance between the empirical CDF and the fitted CDF.

K-S is sensitive to the single largest deviation between the empirical and fitted distributions, making it complementary to A-D (which weights tails more heavily). Lower D indicates a better fit. D is bounded on [0, 1]; values near 0 mean the fitted CDF closely tracks the empirical CDF.

Returns: Kolmogorov-Smirnov D statistic as a scalar in [0, 1] (lower = better fit)

Kolmogorov-Smirnov D = max|F_n(xᵢ) − F(xᵢ)|. Maximum deviation between empirical and fitted CDFs. Lower = better.

```excel
=LAMBDA(data, dist_cdf, [include],
  LET(
    inc, IF(ISOMITTED(include), ISNUMBER(data), include),
    d, FILTER(data, inc),
    f, FILTER(dist_cdf, inc),
    n, ROWS(d),
    IF(n = 0, NA(),
    LET(
      fs, SORTBY(f, d),
      i, SEQUENCE(n),
      d_plus, MAX(i / n - fs),
      d_minus, MAX(fs - (i - 1) / n),
      MAX(d_plus, d_minus)
    ))
  ))
```

## `Gram_Inverse`

**The matrix inverse the regression uses for leverage, intervals, and influence math.**

Arguments:

- **X** — design matrix (already filtered and intercept-augmented as needed)

Returns (X'X)⁻¹ for a design matrix X. Single point of inversion used by Hat_Diagonal, PRESS, LOOCV_Prediction, and Prediction_Interval — any future stability upgrade (e.g., Tikhonov regularisation) only needs to be made here.

Pass the output of Design_Matrix — already filtered and intercept-augmented — directly as X.

Returns: p×p matrix — the inverse of the Gram matrix (X'X)⁻¹

Inverse of the Gram matrix (X’X)⁻¹. Central to leverage, prediction intervals, and LOOCV calculations. Pass Design_Matrix output.

```excel
=LAMBDA(X,
  MINVERSE(MMULT(TRANSPOSE(X), X))
)
```

## `Grid_Argument_Minimum`

**The smallest value in a grid and the row and column where its first occurrence appears.**

Arguments:

- **grid** — rectangular numeric grid to search; pass the complete body range only

Grid_Argument_Minimum finds the minimum numeric value in a rectangular grid and returns that value together with its 1-based row and column locations. TOCOL scans the grid by row, XMATCH selects the first occurrence of the minimum, and QUOTIENT/MOD convert the flattened position back to grid coordinates. Returning one matched position avoids combining the row of one tied minimum with the column of another.

For a rectangular grid such as a 2-D Full_Factorial spill, pass only the grid body, excluding any header and axis-value columns. The grid is expected to contain numeric results; an all-nonnumeric grid returns NA() in all three output positions.

Returns: 1×3 row vector: minimum value, 1-based row location, and 1-based column location of the first minimum

Minimum value in a grid plus its 1-based row and column. Pass the body of the grid to locate MLE parameter estimates.

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

## `Group_Count_At`

**How many included rows belong to one specific group — the sample size behind a group-mean prediction's uncertainty.**

Arguments:

- **group** — single-column group identifiers
- **selected_group** — the one group value to count included rows for
- **include** — boolean mask — TRUE keeps the row; when omitted, rows with a non-blank group are used

Group_Count_At is Tᵢ — the number of included observations in one specific group — the third group-keyed summary (alongside Group_Mean_At for ȳᵢ and x̄ᵢ) the v2.1 Fixed Effects group-mean prediction recovery needs: the mean-response and new-observation prediction variances both scale by 1/Tᵢ.

Unlike Group_Mean_At, a group with zero included members returns 0 (a genuine, meaningful count), not #N/A — counting is always well-defined even when the resulting mean would not be.

Returns: The count of included rows where group = selected_group, as a scalar integer; 0 when selected_group has no included members.

Count of included rows matching selected_group. 0 (not #N/A) when the group has no included members — counting is always well-defined. Backs the v2.1 FE group-mean prediction recovery.

```excel
=LAMBDA(group, selected_group, [include],
    LET(
        g_v, TOCOL(IF(group = "", "", group), 0),
        inc, IF(ISOMITTED(include), (g_v <> ""), include * 1),
        SUMPRODUCT(inc, --(g_v = selected_group))
    )
)
```

## `Group_Mean`

**The average of a column within each row's own group (e.g. each Facility's average cost) — one source of truth for Demean_By.**

Arguments:

- **x** — single-column data range to average
- **group** — single-column group identifiers — the mean is taken within each row's group
- **include** — boolean mask — TRUE keeps the row when computing each group's mean; when omitted, numeric x rows are used

Group_Mean returns, for every row, the mean of x among the included rows sharing that row's group. Groups are resolved once via UNIQUE over the included sample, and each row looks up its group's mean via XLOOKUP — this avoids an O(n²) pairwise comparison when there are many rows and few groups (a Country/Year panel, or a Facility-by-lot dataset).

A row's own inclusion does not gate whether it receives a value: the function reports what that row's group mean IS over the included sample, even for a row currently excluded from the model (e.g. failing a Filter) — full height, matching the same row-mask contract as X() and Sample_Include(). Only two conditions return #N/A: a blank group value, or a group with zero included members (every member of that group is itself excluded, or the group value never appears among included rows).

Group_Mean is a standalone user-callable transform and also the primitive behind Demean_By — one source of truth, so the two functions can never disagree about a group's mean.

Returns: n × 1 column, row-aligned to x: the mean of x within each row's group, over the included rows of that group; #N/A when the row's group is blank or has no included members.

Group mean per row's own group, over included members only; full-height (a row's own inclusion doesn't gate its output). #N/A for a blank group or a group with no included members. Backs Demean_By.

```excel
=LAMBDA(x, group, [include],
    LET(
        x_v,      TOCOL(IF(x = "", "", x), 0),
        g_v,      TOCOL(IF(group = "", "", group), 0),
        inc,      IF(ISOMITTED(include), ISNUMBER(x_v), include * 1),
        active,   inc * (g_v <> ""),
        IFERROR(
            LET(
                groups, UNIQUE(FILTER(g_v, active)),
                means, BYROW(groups, LAMBDA(gk,
                    LET(mask, active * (g_v = gk),
                        SUMPRODUCT(mask, IF(ISNUMBER(x_v), x_v, 0)) / SUM(mask))
                )),
                XLOOKUP(g_v, groups, means, NA(), 0)
            ),
            NA()
        )
    )
)
```

## `Group_Mean_At`

**The average of a column for one specific group only — e.g. Facility='Site B''s average cost, for a group-mean prediction.**

Arguments:

- **x** — single-column data range to average
- **group** — single-column group identifiers
- **selected_group** — the one group value to return the mean for (a scalar, e.g. a dropdown-selected panel unit)
- **include** — boolean mask — TRUE keeps the row when computing the group's mean; when omitted, numeric x rows are used

Group_Mean_At is the scalar counterpart to Group_Mean: the mean for ONE specific group, not a per-row broadcast. Delegates entirely to Group_Mean — one source of truth for the group-averaging arithmetic — and looks up the requested group's value via XLOOKUP, exact match. This is the ȳᵢ / x̄ᵢ primitive the v2.1 Fixed Effects group-mean prediction recovery uses: the group-mean-recovery point estimate ŷ = ȳᵢ + (x_new − x̄ᵢ)′β̂ needs the SELECTED group's own mean, not every row's.

#N/A when selected_group is not among the included group values — never a silent 0, which would misrepresent an unobserved or fully-excluded group as having a real (zero) mean.

Returns: The mean of x among included rows where group = selected_group, as a scalar; #N/A when selected_group does not appear in group.

Scalar group mean for one selected group: XLOOKUP into Group_Mean(x,group,include). #N/A when selected_group isn't among the included groups. Backs the v2.1 FE group-mean prediction recovery.

```excel
=LAMBDA(x, group, selected_group, [include],
    XLOOKUP(selected_group, group, Group_Mean(x, group, include), NA(), 0)
)
```

## `Group_Prediction_Interval`

**Predicts for a specific group (e.g. one Facility) using that group's own average as the baseline, adjusted for how the new inputs differ from that group's typical values — with both a mean-response interval and a wider individual-observation interval.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Y** — single-column outcome range
- **X_new** — column vector of predictor values for the new observation — intercept term first (=1 or 0) when Allow_Intercept is TRUE, matching the layout of pred_input on the Regression sheet
- **group** — single-column panel-unit identifiers (the Fixed Effects variable, e.g. Country) — residual differences never cross a group boundary
- **selected_group** — the one group value to predict for (a dropdown-selected panel unit, or the constant "(all)" sentinel when no Fixed Effects row is declared)
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **alpha** — significance level (0–1); default 0.05 yields 95% prediction intervals
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Group_Prediction_Interval is the v2.1 group-mean-recovery prediction: ŷ = ȳᵢ + (x_new − x̄ᵢ)′β̂, the algebraic identity that recovers the group-specific intercept an LSDV fit would have materialized (α̂ᵢ = ȳᵢ − x̄ᵢ′β̂) without ever fitting G−1 group dummies. Surfaces BOTH intervals DECISIONS.md calls for: a mean-response CI (Var = σ²/Tᵢ + (x_new−x̄ᵢ)′V_β(x_new−x̄ᵢ)) and a new-observation PI (the same quadratic term plus σ²), same center, same t-critical, differing by exactly one variance term — reusing the existing leverage quadratic-form machinery (Gram_Inverse, Design_Matrix) fed the *deviation* (x_new−x̄ᵢ) instead of raw x_new.

Re-demeans Predictors/Y internally via Demean_By(·, group, include) rather than accepting pre-demeaned Design_Columns()/Design_Response(): those return the RAW Predictors()/Response_Column() unchanged when no Fixed Effects row is declared (preserving the v2.0 model's own intercept exactly), which is correct for the Coefficients/ANOVA/Diagnostics reporting chain but wrong for this function's quadratic-form term, which needs a properly (grand-mean-)centered matrix regardless of FE status. Passing Prediction_Group_Column() as group makes this re-demeaning reduce to the real Fixed_Effects_Column() when FE is active, or to plain grand-mean centering (the standard textbook regression-prediction-variance decomposition) when it is not — an intercept fit on top of either is always mathematically inert (the coefficient is exactly 0 by construction, confirmed independently), so beta only needs the slope rows, never the intercept row, regardless of Has_Intercept.

Degenerate collapse (docs/DECISIONS.md's 'build it once'): with group = Prediction_Group_Column()'s constant "(all)" fallback (G=1, whole sample is the group), ȳᵢ/x̄ᵢ/Tᵢ become the grand mean/grand mean/full count, and this function's point estimate and PI reduce EXACTLY to Prediction_Interval()'s own numbers — verified to floating-point precision, not merely close. Predicting at the selected group's own centroid (x_new = x̄ᵢ) kills the quadratic term and the mean-CI collapses to t·√(σ²/Tᵢ), the standard error of ȳᵢ — the DECISIONS.md sanity check.

Scope, per DECISIONS.md: one-way FE only, iid errors, existing groups only (selected_group must be an observed level — Group_Mean_At/Group_Count_At return #N/A/0 rather than fabricating statistics for an unobserved group). Categorical-predictor × FE input encoding (forming X_new/x̄ᵢ in constructed design-matrix space when a non-FE Categorical Predictor coexists with Fixed Effects) is explicitly DEFERRED — this function's X_new is continuous-predictor space only.

Returns: 9-element vertical array: [Point Estimate, SE (mean), SE (new obs), t Critical, CI Lower, CI Upper, PI Lower, PI Upper, Confidence Level]

v2.1 group-mean-recovery prediction: ŷ=ȳᵢ+(x_new−x̄ᵢ)′β̂. Surfaces both mean-CI and new-obs-PI. G=1 (Prediction_Group_Column()'s "(all)" fallback) collapses exactly to Prediction_Interval(). One-way FE, iid errors, continuous predictors only.

```excel
=LAMBDA(Predictors, Y, X_new, group, selected_group, [Include], [alpha], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg,      IF(ISOMITTED(Include),       TRUE, Include),
    alpha_arg,     IF(ISOMITTED(alpha),         0.05, alpha),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    x_within,      LET(
                       n_x,   COLUMNS(Predictors),
                       seed,  SEQUENCE(ROWS(Predictors), 1, 0, 0),
                       built, REDUCE(seed, SEQUENCE(n_x), LAMBDA(acc, j,
                           HSTACK(acc, Demean_By(INDEX(Predictors, 0, j), group, filt_arg))
                       )),
                       DROP(built, , 1)
                   ),
    y_within,      Demean_By(Y, group, filt_arg),
    design_within, IF(has_arg, HSTACK(SEQUENCE(ROWS(x_within), 1, 1, 0), x_within), x_within),
    beta_full,     Coefficients(design_within, y_within, filt_arg),
    beta,          IF(has_arg, DROP(beta_full, 1), beta_full),
    s,             SE_Regression(design_within, y_within, filt_arg, context_arg),
    df,            Residual_Degrees_Of_Freedom(design_within, y_within, filt_arg, context_arg),
    t_crit,        T.INV.2T(alpha_arg, df),
    X_design,      Design_Matrix(x_within, filt_arg),
    XtX_inv,       Gram_Inverse(X_design),
    xbar_i,        TRANSPOSE(BYCOL(Predictors, LAMBDA(col, Group_Mean_At(col, group, selected_group, filt_arg)))),
    ybar_i,        Group_Mean_At(Y, group, selected_group, filt_arg),
    T_i,           Group_Count_At(group, selected_group, filt_arg),
    deviation,     X_new - xbar_i,
    quad,          MMULT(MMULT(TRANSPOSE(deviation), XtX_inv), deviation),
    point_est,     ybar_i + SUMPRODUCT(deviation, beta),
    se_mean,       s * SQRT(1 / T_i + quad),
    se_new,        s * SQRT(1 + 1 / T_i + quad),
    margin_mean,   t_crit * se_mean,
    margin_new,    t_crit * se_new,
    VSTACK(
      point_est,
      se_mean,
      se_new,
      t_crit,
      point_est - margin_mean,
      point_est + margin_mean,
      point_est - margin_new,
      point_est + margin_new,
      1 - alpha_arg
    )
  ))
```

## `GVIF`

**How much overlap with the other predictors is inflating a variable's instability — one shared number per variable, even when it spans several dummy columns.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Column_Names** — 1×k row vector of column labels aligned with X_s, following the "Header: level" convention (e.g. Constructed_Column_Names()); columns sharing the text before the first ": " are treated as one variable's group
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Generalized Variance Inflation Factor (Fox & Monette, 1992): GVIF_g = det(R11) * det(R22) / det(R), where R is the k×k correlation matrix of Predictors (via Correlation_Matrix), R11 is the submatrix restricted to source-variable g's own columns, and R22 is the submatrix restricted to every other column. Collapses a categorical predictor's whole dummy block into one coding-invariant number instead of L−1 separate, coding-dependent per-dummy VIF values.

Groups are recovered from Column_Names by splitting on the first ": " — the same convention Constructed_Column_Names() already uses to label dummy columns. When a group has exactly one column (an ordinary continuous predictor, or a two-level categorical), GVIF is numerically identical to VIF's own 1/(1−R²ⱼ) — a matrix identity, not an approximation — so the existing 5/10 review thresholds stay valid for those rows. When every column resolves to the same group (only one source variable in the whole matrix) GVIF returns 1 for all columns, the direct generalization of VIF's own k=1 special case.

det(R) appears in every column's denominator, so exact multicollinearity anywhere in Predictors (det(R) = 0) produces #DIV/0! for every column, not only the collinear ones — the same all-or-nothing failure mode VIF's own 1/(1−R²ⱼ) already has at R²ⱼ = 1, surfaced once instead of per column.

Returns: k×1 column vector of GVIF values, one per column of Predictors — every dummy column sharing a source variable's group carries the same value

Generalized VIF, one value per source variable (shared across a categorical predictor's dummy columns): det(R11)*det(R22)/det(R). Equals ordinary VIF when a variable is a single column.

```excel
=LAMBDA(Predictors, Column_Names, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    k,        COLUMNS(Predictors),
    grp,      IFERROR(LEFT(Column_Names, FIND(": ", Column_Names) - 1), Column_Names),
    IF(COUNTA(UNIQUE(grp, TRUE)) <= 1,
      MAKEARRAY(k, 1, LAMBDA(j, _, 1)),
      LET(
        R,    Correlation_Matrix(Predictors, filt_arg),
        detR, MDETERM(R),
        MAKEARRAY(k, 1, LAMBDA(j, _,
          LET(
            g,   INDEX(grp, 1, j),
            own, grp = g,
            R11, FILTER(FILTER(R, TRANSPOSE(own)), own),
            R22, FILTER(FILTER(R, TRANSPOSE(NOT(own))), NOT(own)),
            MDETERM(R11) * MDETERM(R22) / detR
          )
        ))
      )
    )
  ))
```

## `Hat_Diagonal`

**Each row's leverage, meaning how unusual its predictor values are.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Returns h_i = x_iᵀ(XᵀX)⁻¹x_i for each observation. Values range from p/n (minimum leverage, uniform design) to 1. Observations with h_i > 2p/n are conventionally flagged as high-leverage. Used internally by Studentized_Residuals, Cooks_Distance, PRESS, and LOOCV_Prediction; exposed here for direct inspection. Inversion is delegated to Gram_Inverse.

Returns: n-element hat-matrix diagonal (leverage values) as spilled vector

Leverage hᵢ = xᵢ’(X’X)⁻¹xᵢ per row. Rows with hᵢ > 2p/n have high leverage; used by Cook’s Distance and PRESS.

```excel
=LAMBDA(X, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    X_design, Design_Matrix(X, filt_arg),
    p,        COLUMNS(X_design),
    XtX_inv,  Gram_Inverse(X_design),
    Z,        MMULT(X_design, XtX_inv),
    MMULT(Z * X_design, SEQUENCE(p, 1, 1, 0))
  ))
```

## `Interact`

**Elementwise product of two operands, broadcasting a column across a dummy matrix — the standalone way to build an interaction column (including x·x for a quadratic).**

Arguments:

- **x1** — first operand — a column or a dummy-coded matrix (n × k)
- **x2** — second operand — a column (broadcasts across x1's columns when x1 is a matrix)

Interact returns the elementwise product of two operands — the standalone, free-form counterpart to the v3.0 spec-block interaction columns (§4 M/N). The spec-driven path does not call it: the constructor encodes interactions inline, the same relationship Predictor_Columns() already has with Dummy_Code. A user assembling their own design matrix calls Interact directly to form a continuous × categorical interaction (a dummy matrix times a continuous predictor broadcasts to one interaction column per retained level) or a self-interaction x·x for a quadratic term.

It takes no [include] mask: the operands carry their own row alignment. An excluded row in either operand is the library's "" sentinel, and Interact passes it through — IF((x1="")+(x2=""), "", x1*x2) returns "" where either side is blank, so the result stays row-aligned when composing transform outputs (Dummy_Code, Ln_Positive) that mark excluded rows with "". This is the same excluded-row contract every row-aligned transform honours; a bare x1*x2 would turn a "" into #VALUE! and break that alignment.

Genuine errors and #N/A propagate through x1*x2 rather than being swallowed — the NA()-exception convention is respected: an included-but-incomputable value (a Log of a non-positive number, say) surfaces as #N/A in the interaction too, never silently coerced to 0. Excel's native array arithmetic does the broadcasting (n × k × n × 1 → n × k); non-conformable shapes (n × k × n × m with k ≠ m and neither 1) produce #VALUE!, the honest answer for inputs that cannot be multiplied.

Returns: elementwise product x1·x2, broadcasting n × k × n × 1 → n × k (one interaction column per retained level when one operand is a dummy matrix and the other a column); "" on rows where either operand is the excluded-row sentinel; #N/A and other errors propagate through the product.

Elementwise product x1·x2 with Excel broadcasting (n×k × n×1 → n×k). "" in either operand → "" (stays row-aligned, composes with transform outputs). #N/A/errors propagate. No [include]. Free-form counterpart to the spec-block interaction columns.

```excel
=LAMBDA(x1, x2,
    IF((x1 = "") + (x2 = ""), "", x1 * x2)
)
```

## `Is_Balanced_Panel`

**TRUE when every group has exactly one observation at every time period — a complete panel grid with no gaps or duplicates.**

Arguments:

- **group** — single-column group identifiers (the panel unit)
- **time** — single-column time index
- **include** — boolean mask — TRUE keeps the row; when omitted, rows with a non-blank group and time are used

Is_Balanced_Panel checks whether the included sample forms a complete rectangular (group × time) grid: every included group observed at every included time period, exactly once. Two conditions, both required: the included row count equals |groups| × |times| (no missing cells), and every (group, time) pair is distinct (no duplicate cells) — together they rule out both gaps and doubled-up observations, either of which a simple row-count check alone would miss.

This is a diagnostic, not a constructor internal: it does not feed Demean_By (the one-way within transformation is exact and requires no balance), but it flags when a stronger method (e.g. a two-way within transformation via direct demeaning, Demean_Two_Way_Balanced, valid only for a balanced panel) may be safely used instead of the general-but-iterative Absorb_Two_Way_Fixed_Effects — a later-milestone concern, ships now as a standalone diagnostic ahead of that milestone.

Returns: TRUE when every included group has exactly one observation for every included time period (a complete, non-duplicated panel grid); FALSE otherwise; #N/A when no rows are included.

TRUE iff the included sample is a complete (group × time) grid: row count = groups×times AND every (group,time) pair distinct. Diagnostic only — flags when a two-way demeaning shortcut is valid; Demean_By itself needs no balance.

```excel
=LAMBDA(group, time, [include],
    LET(
        g_v,     TOCOL(IF(group = "", "", group), 0),
        t_v,     TOCOL(IF(time = "", "", time), 0),
        inc,     IF(ISOMITTED(include), (g_v <> "") * (t_v <> ""), include * 1),
        active,  inc * (g_v <> "") * (t_v <> ""),
        n,       SUM(active),
        IF(n = 0,
           NA(),
           LET(
               sep,     CHAR(31),
               groups,  UNIQUE(FILTER(g_v, active)),
               times,   UNIQUE(FILTER(t_v, active)),
               keys,    FILTER(g_v & sep & t_v, active),
               AND(n = ROWS(groups) * ROWS(times), n = ROWS(UNIQUE(keys)))
           )
        )
    )
)
```

## `Kurtosis`

**Whether a variable has unusually heavy tails or outliers.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Kurtosis returns the excess kurtosis for each filtered column of Predictors. Uses Excel's KURT function (Fisher's excess kurtosis: kurtosis − 3, so a normal distribution returns 0).

Positive excess kurtosis (leptokurtic) indicates heavier tails and more extreme outliers than a normal distribution. Negative (platykurtic) indicates lighter tails. High kurtosis in a predictor or in regression residuals signals that OLS standard errors may be unreliable, and that robust standard errors or a different distributional assumption may be warranted.

Returns: k×1 column vector of excess kurtosis values (one per column of Predictors)

Excess kurtosis per Predictors column. 0 = normal-distributed tails; positive = heavier tails with more frequent extremes.

```excel
=LAMBDA(Predictors, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    xf, FILTER(Predictors, filt_arg),
    MAKEARRAY(COLUMNS(Predictors), 1, LAMBDA(r, c, KURT(CHOOSECOLS(xf, r))))
  ))
```

## `Lag_By`

**The previous period's value of a column within the same group — matched by actual time value, so gaps give #N/A instead of grabbing the wrong row.**

Arguments:

- **x** — single-column data range to lag
- **group** — single-column group identifiers (e.g. Country) — lags never cross a group boundary
- **seq** — single-column numeric time index (e.g. Year) — the lookup key, not the row order
- **delta** — base period Δ: the lag looks up seq − delta exactly; when omitted, the spec's Base Period Δ (Base_Period_Delta()) is used — never a silent 1
- **include** — boolean mask — TRUE keeps the row; when omitted, rows with a non-blank group and numeric seq are used

Lag_By returns the prior-period value of x within the same group: for each included row, the value of x on the row whose group matches and whose seq equals exactly seq − Δ.

The lookup is by time value, not row position — by construction, not convention: the implementation is an exact-match XLOOKUP of composite (group, seq − Δ) keys, never OFFSET or row arithmetic. If the literal period seq − Δ is absent within the group (a gap in the panel), the result is #N/A. It never falls back to "the previous available row", because that would silently splice a multi-period jump into a series that claims to be a one-Δ lag. A group's first observed period has no seq − Δ row and is #N/A for the same reason.

Δ is computed-and-displayed-with-override, never silently assumed: an omitted delta resolves through Base_Period_Delta() — the visible Base Period Δ spec cell (column I of the Sequence-flagged row) — and when no sequence axis or Δ is declared, every included row returns #N/A rather than guessing Δ = 1.

Rows failing the include mask return "" so the spilled column stays row-aligned with the source (the standard row-aligned-transform contract). Excluded rows are also removed from the lookup source, so a lag never reads through an excluded observation. Blank x values are normalized to "" and never served as lag values (the lookup misses → #N/A). The composite key joins group and seq with CHAR(31), so ordinary text groups cannot collide; duplicate (group, seq) pairs are a data defect — the first occurrence wins. seq should be an integer period index (Year, or YEAR*12+MONTH for monthly data); raw calendar serial dates make seq − Δ land between real dates — see the Sequence Spacing block's calendar-signature guidance.

Returns: n × 1 column, row-aligned to x: the value of x at (group, seq − Δ); #N/A when that exact prior period does not exist; "" on excluded rows.

Prior-period value within the group: exact-match lookup of (group, seq − Δ). Gap or first period → #N/A, never the previous available row. Omitted delta reads the spec's Base Period Δ (never a silent 1). Excluded rows return "".

```excel
=LAMBDA(x, group, seq, [delta], [include],
    LET(
        x_v,  TOCOL(IF(x = "", "", x), 0),
        g_v,  TOCOL(IF(group = "", "", group), 0),
        t_v,  TOCOL(IF(seq = "", "", seq), 0),
        step, IF(ISOMITTED(delta), Base_Period_Delta(), delta),
        inc,  IF(ISOMITTED(include), (g_v <> "") * ISNUMBER(t_v), include * 1),
        sep,  CHAR(31),
        source_keys, IF(inc * ISNUMBER(t_v) * (x_v <> ""), g_v & sep & t_v, ""),
        prior, XLOOKUP(g_v & sep & (t_v - step), source_keys, x_v, NA(), 0),
        IF(inc, IF(ISNUMBER(t_v), prior, NA()), "")
    )
)
```

## `Leave_One_Out_Prediction`

**One leave-one-out prediction for the requested row.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **n** — 1-based row to exclude; negative values count from the end (−1 = last row)
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Applies the filter once in the outer scope to produce X_filt and Y_filt, then delegates to inner LAMBDAs without re-passing the filter. Exclude_Row_N removes observation n from both X and Y; Coefficients fits OLS on the remaining n−1 rows; Design_Matrix builds the single-row design vector for row n (with or without intercept column); MMULT computes the dot product.

Note: for n LOO predictions at once use LOOCV_Prediction, which avoids n separate OLS fits via the hat-matrix shortcut.

Returns: scalar LOO prediction for observation n — the fitted value from a model trained on all rows except n, evaluated at row n of X

Single leave-one-out prediction for row n by refitting on n−1 rows. For all rows at once use LOOCV_Prediction.

```excel
=LAMBDA(X, Y, n, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    X_filt,   FILTER(X, filt_arg),
    Y_filt,   FILTER(Y, filt_arg),
    X_loo,    Exclude_Row_N(X_filt, n),
    Y_loo,    Exclude_Row_N(Y_filt, n),
    x_n,      Design_Matrix(CHOOSEROWS(X_filt, n)),
    beta_loo, Coefficients(X_loo, Y_loo),
    MMULT(x_n, beta_loo)
  )
)
```

## `Ln_Positive`

**Natural log of x, restricted to strictly positive values — zero/negative/non-numeric input on an included row is a visible #N/A, not a silent blank.**

Arguments:

- **x** — single-column numeric data range to transform
- **include** — boolean mask — TRUE keeps the row; when omitted, all numeric rows are used

Ln_Positive returns the natural log of x, restricted to strictly positive numeric input. It is the primitive behind the Regression sheet's Transform = Log spec option (column G): the model spec block's Response_Column(), X(), and Constructed_Column_Names() all call this function when a row declares Log, so every log-transformed column in the workbook shares one implementation.

The row-mask contract matches every other transform in the catalog (Demean_By, Dummy_Code): an excluded row returns "" so a spilled array stays row-aligned with its source, exactly the same as a row this function was never asked about. An included row that cannot be logged — zero, negative, or non-numeric — returns #N/A, not a blank. This follows the library's recorded NA()-exception convention (the same one Lag_By/Difference_By use for a first period or a panel gap): "" means "this row isn't part of the sample"; #N/A means "this row is part of the sample and the requested computation is genuinely undefined for it." A non-positive value under a declared Log transform is exactly that — a modeling error that must surface loudly (a visible #N/A the fit can't quietly absorb into a MMULT), never a silent blank that degrades the whole design matrix without explanation.

The Regression sheet reaches this function through TWO spec tokens, and the difference between them is entirely in the mask it is handed, not in anything here. Under Transform = “Log”, Sample_Include() still contains the non-positive rows, so they arrive included and this function returns #N/A for them — the loud failure described above. Under Transform = “Log (drop ≤ 0)”, Sample_Include() has already excluded them, so they arrive with include = 0 and return "" instead. One implementation, two declared behaviours, and the #N/A branch stays exactly as specified for every direct library caller.

When include is omitted, the default mask is ISNUMBER(x) — the same default every other row-masked transform in the catalog uses when the caller doesn't supply one explicitly.

Returns: n × 1 column, row-aligned to x: LN(x) for an included, strictly-positive numeric value; #N/A for an included value that is zero, negative, or non-numeric; "" on excluded rows.

LN(x) for strictly positive numeric input. Excluded row -> "". Included but non-positive/non-numeric -> #N/A (the NA()-exception convention: "" means excluded, #N/A means included-but-incomputable). Backs the Regression sheet's Transform = Log option.

```excel
=LAMBDA(x, [include],
    LET(
        x_v, TOCOL(IF(x = "", "", x), 0),
        inc, IF(ISOMITTED(include), ISNUMBER(x_v), include * 1),
        pos, IFERROR(ISNUMBER(x_v) * (x_v > 0), FALSE),
        IF(inc = 0, "", IF(pos, LN(x_v), NA()))
    )
)
```

## `Log_Domain_Status` *(sheet-scoped: Regression)*

**Reports whether a strict Log makes the fitted model fail, or how many otherwise usable rows "Log (drop ≤ 0)" deliberately removed.**

Log_Domain_Status is the G2 readout: what the declared Log transforms are doing to the model and its sample. Two states, most severe first.

RED — an eligible variable declaring strict "Log" contains one or more zero or negative values among the rows that actually reach the fit. Plain Log does not filter those rows, so Ln_Positive returns #N/A and the error propagates through the fit. The message names the first variable with the largest offending-row count and points to "Log (drop ≤ 0)" as the explicit row-dropping alternative.

AMBER — "Log (drop ≤ 0)" has removed otherwise eligible model rows. The count is the difference between Sample_Include(), which represents ordinary eligibility before transform-driven dropping, and Fit_Sample_Include(), the materialized final mask produced from Log_Drop_Sample_Include_Calc(). Because that is the only additional exclusion between those masks, the difference is exactly the number of distinct records removed by the Log-drop rule. A row violating two drop-Log variables counts once.

Blank means neither applies.

SHEET-SCOPED (scope: Regression).

Returns: An ERROR string naming the variable and non-positive fitted-row count when a strict Log makes the fit invalid; otherwise the number of ordinary eligible rows removed by "Log (drop ≤ 0)"; otherwise "".

G2: RED for fitted strict-Log nonpositives; otherwise AMBER for ordinary eligible rows removed by Log (drop ≤ 0), else blank.

```excel
=LAMBDA(
    LET(
      Number_Of_Columns,COLUMNS(Source_Data),
      Spec_Roles,TAKE(Spec_Role,Number_Of_Columns),
      Spec_Includes,TAKE(Spec_Include,Number_Of_Columns),
      Spec_Types,TAKE(Spec_Type,Number_Of_Columns),
      Spec_Transforms,TAKE(Spec_Transform,Number_Of_Columns),
      Headers,TOROW(Header_Names),
      Base_Sample_Include,Sample_Include(),
      Fitted_Sample_Include,Fit_Sample_Include(),
      Eligible_Columns,((Spec_Roles="Response (y)")+((Spec_Roles="Predictor (x)")*(Spec_Includes=TRUE)*(Spec_Types="Continuous")))>0,
      Strict_Log_Nonpositive_Counts,
        MAP(
          SEQUENCE(Number_Of_Columns),
          LAMBDA(Column_Number,
            IF(
              AND(
                INDEX(Eligible_Columns,Column_Number),
                INDEX(Spec_Transforms,Column_Number)="Log"
              ),
              SUMPRODUCT(
                --Fitted_Sample_Include,
                --IFERROR((INDEX(Source_Data,0,Column_Number)+0)<=0,FALSE)
              ),
              0
            )
          )
        ),
      Maximum_Strict_Log_Nonpositive_Count,MAX(Strict_Log_Nonpositive_Counts),
      Log_Drop_Excluded_Row_Count,SUMPRODUCT(--Base_Sample_Include)-SUMPRODUCT(--Fitted_Sample_Include),
      IF(
        Maximum_Strict_Log_Nonpositive_Count>0,
        LET(
          Worst_Column_Number,XMATCH(Maximum_Strict_Log_Nonpositive_Count,Strict_Log_Nonpositive_Counts),
          "ERROR: "&INDEX(Headers,1,Worst_Column_Number)&" has "&Maximum_Strict_Log_Nonpositive_Count&" values ≤ 0 under Log — the fit is #N/A. Use Log (drop ≤ 0)."
        ),
        IF(
          Log_Drop_Excluded_Row_Count=0,
          "",
          Log_Drop_Excluded_Row_Count&" row"&IF(Log_Drop_Excluded_Row_Count=1,"","s")&" excluded: Log of ≤ 0"
        )
      )
    )
)
```

## `Log_Drop_Sample_Include_Calc` *(sheet-scoped: Regression)*

**Takes the ordinary model sample and removes rows that are non-positive only where "Log (drop ≤ 0)" was explicitly selected.**

Log_Drop_Sample_Include_Calc adds the one transform-specific row-removal rule supported by the regression spec to the ordinary Sample_Include_Calc mask.

It begins with Sample_Include_Calc(), so every Filter and required-numeric-data condition has already been applied. It then walks only transform-eligible model columns: the Response (y), and included Continuous Predictor (x) rows. When one of those columns declares exactly "Log (drop ≤ 0)", rows whose value is zero or negative are removed from the mask.

Plain "Log" is intentionally ignored. A strict Log does not modify the sample: a non-positive value remains present and Ln_Positive returns #N/A, causing the fit to fail visibly. This gives the two transform choices distinct semantics — "Log" means transform or fail, while "Log (drop ≤ 0)" means explicitly remove invalid-domain rows and fit the remaining sample.

This function produces the materialized sample-inclusion spill read by Fit_Sample_Include(). Separating it from Sample_Include_Calc makes the transform-driven sample modification visible in both the function name and dependency graph.

SHEET-SCOPED (scope: Regression).

Returns: The final model-sample mask after ordinary eligibility plus the positivity rule explicitly requested by every eligible "Log (drop ≤ 0)" transform.

Final-sample leaf: ordinary eligibility plus positivity only for eligible columns declaring Log (drop ≤ 0).

```excel
=LAMBDA(
    LET(
      Number_Of_Columns,COLUMNS(Source_Data),
      Spec_Roles,TAKE(Spec_Role,Number_Of_Columns),
      Spec_Includes,TAKE(Spec_Include,Number_Of_Columns),
      Spec_Types,TAKE(Spec_Type,Number_Of_Columns),
      Spec_Transforms,TAKE(Spec_Transform,Number_Of_Columns),
      Base_Sample_Mask,--Sample_Include_Calc(),
      Log_Drop_Sample_Mask,
        REDUCE(
          Base_Sample_Mask,
          SEQUENCE(Number_Of_Columns),
          LAMBDA(Accumulated_Mask,Column_Number,
            LET(
              Source_Column,INDEX(Source_Data,0,Column_Number),
              Eligible_Column,
                OR(
                  INDEX(Spec_Roles,Column_Number)="Response (y)",
                  AND(
                    INDEX(Spec_Roles,Column_Number)="Predictor (x)",
                    INDEX(Spec_Includes,Column_Number)=TRUE,
                    INDEX(Spec_Types,Column_Number)="Continuous"
                  )
                ),
              IF(
                AND(
                  Eligible_Column,
                  INDEX(Spec_Transforms,Column_Number)="Log (drop ≤ 0)"
                ),
                Accumulated_Mask*--IFERROR((Source_Column+0)>0,FALSE),
                Accumulated_Mask
              )
            )
          )
        ),
      Log_Drop_Sample_Mask=1
    )
)
```

## `LOOCV_Prediction`

**Each row's prediction from a model refit without that row.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Returns the leave-one-out predicted value for every observation simultaneously, without fitting n separate models. Uses the Sherman-Morrison-Woodbury update: ŷ_{−i}(xᵢ) = ŷᵢ − hᵢ eᵢ / (1 − hᵢ), where ŷᵢ = Predictions, eᵢ = Y − ŷᵢ, and hᵢ is the hat-matrix diagonal from Hat_Diagonal.

The LOOCV residual for each row equals eᵢ / (1 − hᵢ), which is exposed directly by LOOCV_Residual; SUMSQ(LOOCV_Residual(...)) = PRESS.

Returns: n-element spill vector of LOO predictions for every filtered observation

All n leave-one-out predictions at once using the hat-matrix shortcut ŷ₋ᵢ = ŷᵢ − hᵢeᵢ/(1−hᵢ). No n separate refits needed.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    yhat,     Predictions(X, Y, filt_arg),
    e,        FILTER(Y, filt_arg) - yhat,
    h,        Hat_Diagonal(X, filt_arg),
    yhat - h * e / (1 - h)
  )
)
```

## `LOOCV_Residual`

**Each row's prediction error when that row is left out of the fit.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Returns the leave-one-out cross-validation residual for every observation simultaneously. The shortcut formula is e_i / (1 - h_i), where e_i is the ordinary OLS residual and h_i is the hat-matrix diagonal from Hat_Diagonal. Squaring and summing this vector gives PRESS. A row with h_i = 1 (its fit is fully determined by that single observation, e.g. a singleton dummy level) makes the denominator 0; that row returns #N/A instead of #DIV/0!.

Returns: n-element spill vector of leave-one-out residuals for every filtered observation

Leave-one-out residuals e_i/(1-h_i). SUMSQ of this spilled vector equals PRESS. #N/A (not #DIV/0!) where h_i = 1.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    e, Residuals(X, Y, filt_arg),
    h, Hat_Diagonal(X, filt_arg),
    IFERROR(e / (1 - h), NA())
  )
)
```

## `Min_NLL_Params`

**The parameter value or values attached to the smallest NLL result.**

Arguments:

- **Param_Array** — candidate parameter rows, one row per trial and one column per parameter
- **NLL_Array** — matching NLL values, one per trial row

Min_NLL_Params returns the parameter row aligned to the first minimum in an NLL vector. XMATCH locates MIN(NLL_Array) within the NLL results, and INDEX returns the full matching row from Param_Array across however many parameter columns it contains.

This works for a single searched parameter (for example a 1-D profile-NLL axis) and for multi-parameter searches such as a Full_Factorial Alpha/Beta grid materialized as an N^2×2 parameter array.

Returns: 1×d row vector containing the parameter row associated with the minimum NLL

Use with a parameter array and its matching NLL array to recover the optimal parameter row at the first minimum.

```excel
=LAMBDA(Param_Array,NLL_Array,INDEX(Param_Array,XMATCH(MIN(NLL_Array),NLL_Array),SEQUENCE(,COLUMNS(Param_Array))))
```

## `Missing_Count`

**How many cells in the data column are blank or non-numeric.**

Arguments:

- **data** — single-column data range (may contain blanks or non-numeric cells)
- **filter** — optional boolean array — TRUE includes the row, FALSE excludes it

Missing_Count counts cells that are present in the active data range but not numeric. The active range is determined from the first row down to the last non-empty row (using MATCH to find the last non-empty entry), so trailing blanks beyond the user's data do not inflate the count.

Returns 0 when no active rows exist. A non-zero result signals that some cells the user likely intended as data values were blank or contained text and were silently excluded from statistical calculations.

Returns: count of non-numeric (blank or text) cells within the active rows of the data range

Count of non-numeric (blank or text) cells in the active data range up to the last non-empty row. Returns 0 when data is clean.

```excel
=LAMBDA(data, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), TRUE, filter),
    d, FILTER(data, filt_arg),
    last_row, IFERROR(LOOKUP(2, 1/(d<>""), SEQUENCE(ROWS(d))), 0),
    IF(last_row = 0, 0,
      last_row - COUNT(TAKE(d, last_row)))
  ))
```

## `Model_Context`

**Build the 4-row context array (intercept flag, absorbed df, response transform, predictor transform) that every v3.0 engine function receives.**

Arguments:

- **Has_Intercept** — TRUE when column 1 of the design matrix is the intercept column
- **DF_Absorbed** — degrees of freedom absorbed by declared Fixed Effects groups; defaults to 0
- **Response_Transform** — transform applied to the response (None/Log/...); defaults to None
- **Predictor_Transform** — per-predictor transform summary (None/Log/Mixed); defaults to None

Model_Context is the workbook-scoped constructor for the bounded 4x1 context array threaded through the v3.0 engine as the [Context] argument. It bundles the spec-derived scalars an engine function cannot derive from its own design-matrix argument: the intercept flag (column-1 identity, not a synthesize switch), the fixed-effects absorbed df, and the two transform summaries. Free-form callers outside the Regression sheet build a context with it; on the Regression sheet the context is materialized once into a spill range and read back by Fit_Context(), a sheet-scoped zero-arg reader thunk over that fixed range. The reader carries its own name so this constructor is not shadowed on the sheet (the pre-stage-two design reused the name Model_Context for the thunk, which is what the unshadowing split removed).

Returns: A 4x1 vertical array: [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform] - the bounded, materializable context every engine function reads via its [Context] argument.

Row order is a versioned public contract - append only, never insert. ROWS(Model_Context()) is the build-time constant asserted in the build. Defaults reproduce the pre-v3.0 no-intercept-extra, no-FE surface: VSTACK(TRUE,0,'None','None').

```excel
=LAMBDA([Has_Intercept], [DF_Absorbed], [Response_Transform], [Predictor_Transform],
  LET(
    has_arg, IF(ISOMITTED(Has_Intercept), TRUE, Has_Intercept),
    df_arg,  IF(ISOMITTED(DF_Absorbed), 0, DF_Absorbed),
    rt_arg,  IF(ISOMITTED(Response_Transform), "None", Response_Transform),
    pt_arg,  IF(ISOMITTED(Predictor_Transform), "None", Predictor_Transform),
    VSTACK(has_arg, df_arg, rt_arg, pt_arg)
  ))
```

## `Model_Formula` *(sheet-scoped: Regression)*

**The model written out in one line — "mpg ~ 1 + weight + horsepower" — built from the spec block, so it always says exactly what is being fitted.**

Model_Formula renders the saved spec as the one-line formula the model actually is — the caption a reader (and the v3.4 Model Comparison sheet) uses to tell two fits apart. It assembles four pieces out of names that already exist, so it can never describe a model other than the one being fitted:

Response — the Role=Response (y) row's header, wrapped as "Ln(name)" when that row's Transform is Log. It reads the same XMATCH-on-Role position Response_Column() reads, so the label and the fitted column cannot disagree; while no Response row is declared it degrades to "(none)" rather than an error.

Intercept — "1 + " or "0 + ", from the Allow_Intercept toggle.

Predictors — TEXTJOIN over Constructed_Column_Names(), which already emits "Ln(name)" per logged predictor, level-qualified dummy names such as "Status[Developing]", and interaction names, so a mixed Log/None spec renders correctly with no extra work here. IFERROR degrades to an empty right-hand side while the spec names nothing.

Fixed effects — a " | <variable>" suffix when a Role=Fixed Effects row is declared, naming the absorbed variable rather than a bare "FE" token; omitted entirely when the count is zero.

It is a DISPLAY: every piece derives from the spec and no constructor reads it back (ARCHITECTURE §4 — display derives, never feeds). It lives in the catalog rather than inline in one cell so the assembly rules are documented on the LAMBDA_functions sheet and every Regression-shaped sheet renders its own spec identically.

Sheet-scoped, not workbook-scoped: the body reads the sheet-scoped spec names (Spec_Role, Spec_Transform, Header_Names, Allow_Intercept) and the Constructed_Column_Names() closure beside them, so each Regression-shaped sheet in a workbook gets its own formula string — the same reason Base_Period_Delta is sheet-scoped rather than hardcoding 'Regression'!.

Returns: The model formula as one string: "<response> ~ 1 + <constructed predictors>", with "0 + " when the intercept is off and a " | <FE variable>" suffix when a Fixed Effects row is declared.

The saved spec rendered as "<response> ~ 1 + <predictors> [| <FE>]". Response and predictor names carry their Log wrapping and dummy level qualification. Display only — no constructor reads it. Feeds Comparison_Model_Formula.

```excel
=LAMBDA(
    LET(
        n_c,        COLUMNS(Source_Data),
        p,          XMATCH("Response (y)", TAKE(Spec_Role, n_c)),
        h,          INDEX(TOROW(Header_Names), p),
        response,   IFERROR(IF(OR(INDEX(TAKE(Spec_Transform, n_c), p) = "Log", INDEX(TAKE(Spec_Transform, n_c), p) = "Log (drop ≤ 0)"), "Ln(" & h & ")", h), "(none)"),
        intercept,  IF(Allow_Intercept, "1 + ", "0 + "),
        predictors, IFERROR(TEXTJOIN(" + ", TRUE, Constructed_Column_Names()), ""),
        fe_count,   SUMPRODUCT(N(TAKE(Spec_Role, n_c) = "Fixed Effects")),
        fe_name,    IFERROR(INDEX(TOROW(Header_Names), XMATCH("Fixed Effects", TAKE(Spec_Role, n_c))), "FE"),
        response & " ~ " & intercept & predictors & IF(fe_count > 0, " | " & fe_name, "")
    )
)
```

## `Model_Matrix`

**A design matrix from an HSTACKed predictor block, with an intercept column of 1s prepended by default — pass FALSE only when X already contains a constant (e.g. a full one-hot block that kept every level, not a reference-dropped Dummy_Code block).**

Arguments:

- **X** — predictor matrix assembled explicitly with HSTACK (the specification stays visible and auditable)
- **add_intercept** — TRUE (the default) prepends a column of 1s; FALSE returns X unchanged — pass FALSE only when X already contains a constant, i.e. a full one-hot block that kept every level (Dummy_Levels output, which does NOT drop the reference). Dummy_Code output drops the reference level, so its columns are linearly independent of an intercept — keep the default TRUE there; passing FALSE fits a no-intercept model, which changes the model rather than avoiding collinearity

Model_Matrix assembles a design matrix from an already-HSTACKed predictor block, optionally prepending an intercept column of 1s. It is intentionally NOT variadic: predictors are assembled explicitly with HSTACK before being passed in, so the model specification stays visible and auditable rather than buried in a variadic argument list. The standalone, free-form counterpart to the v3.0 spec-driven constructor pipeline, which owns its own intercept (§4a) and builds the design matrix from the spec block — Model_Matrix is for a user constructing a matrix by hand from catalog transforms (Dummy_Column, Interact, Center, …) and feeding it to Coefficients or a LINEST call.

The intercept defaults to ON because the common regression design includes one. The one case to opt out with add_intercept = FALSE is when X already contains a constant — most importantly a full one-hot block that kept every level (Dummy_Levels output), where prepending an intercept would re-create the perfect multicollinearity that an all-levels dummy block has with a constant. That opt-out is the same decision the spec-driven constructor makes when Allow_Intercept is FALSE: no ones column, and the design-matrix header asymmetry (Design_Columns one wider than Constructed_Column_Names) disappears. Do NOT pass FALSE merely because X came from Dummy_Code: Dummy_Code drops the reference level, so its columns are linearly independent of an intercept and keeping the default is correct — passing FALSE there fits a no-intercept model (no constant term), which changes the model, not a collinearity fix. Model_Matrix does not validate that choice — it prepends or it doesn't, as asked — because the spec-driven pipeline, not this function, owns intercept correctness from v3.0 on.

The ones column is SEQUENCE(ROWS(X), 1, 1, 0) — a constant column of 1s the same height as X — HSTACKed in front of X. Row alignment is the caller's concern: X may contain the library's "" excluded-row sentinel from its constituent transforms, and the intercept is 1 for every row including those, which is correct (the sample mask is applied downstream at fit time, not by the design-matrix assembler).

Returns: X with a leading column of 1s prepended when add_intercept is TRUE (the default); X unchanged when add_intercept is FALSE.

Design matrix from an HSTACKed X: prepends a ones column when add_intercept is TRUE (default), else X unchanged. Keep the intercept with Dummy_Code (reference dropped). Pass FALSE only for a full one-hot block (all levels) that already has a constant.

```excel
=LAMBDA(X, [add_intercept],
    LET(
        ai, IF(ISOMITTED(add_intercept), TRUE, add_intercept),
        IF(ai, HSTACK(SEQUENCE(ROWS(X), 1, 1, 0), X), X)
    )
)
```

## `MS_Regression`

**Average explained variation per model degree of freedom.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

MS_Regression is the mean square for the regression model: SS_Regression divided by df_Regression. It is the numerator of the overall F-statistic and represents the average explained variation per predictor degree of freedom.

Returns: mean square for regression (MS_regression = SS_regression / df_regression) as a scalar

Mean square for regression = SS_Regression / df_Regression. Numerator of the overall F-statistic.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg, IF(ISOMITTED(Include),       TRUE, Include),
    SS_Regression(X, Y, filt_arg, context_arg) / Regression_Degrees_Of_Freedom(X, context_arg)
  ))
```

## `MS_Residual`

**Average unexplained variation per residual degree of freedom.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

MS_Residual is the mean square for residual error: SS_Residual divided by df_Residual. It equals SE_Regression² and is the denominator of the overall F-statistic. It represents the average unexplained variation per residual degree of freedom.

An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to Residual_Degrees_Of_Freedom() so the divisor correctly reflects absorbed group degrees of freedom.

Returns: mean square for residual error (MS_residual = SS_residual / df_residual) as a scalar

Mean square residual = SS_Residual / df_Residual = SE_Regression². Denominator of the F-statistic. Optional DF_Absorbed (default 0) corrects df_Residual under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    SS_Residual(X, Y, filt_arg) / Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg)
  ))
```

## `Multiple_R`

**How closely the model's predictions track the real outcome values.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Multiple R is the Pearson correlation between the observed values of Y and the values predicted by the regression model. It ranges from 0 (no linear relationship) to 1 (perfect linear fit) and is the square root of R².

Computed as SQRT(R_Squared(...)).

Returns: Multiple R (correlation coefficient) as a scalar in [0, 1]

Correlation between observed and fitted Y values = √R². Ranges [0, 1]. Square to recover R_Squared.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg, IF(ISOMITTED(Include),       TRUE, Include),
    SQRT(R_Squared(X, Y, filt_arg, context_arg))
  ))
```

## `NLL_Beta`

**How unlikely the Beta distribution is to have generated this (0,1)-bounded data at the given shape parameters.**

Arguments:

- **data** — single-column data range; values must be strictly on (0, 1)
- **alpha_param** — first shape parameter α (must be > 0)
- **beta_param** — second shape parameter β (must be > 0)
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_Beta evaluates −Σ ln f(xᵢ | α, β) via BETA.DIST. Values exactly at 0 or 1 cause BETA.DIST to return errors; the IFERROR sentinel catches these and maps the entire NLL to 1E+15, making the cell visible as an overflow region in the spill search.

Beta fits proportions and rates bounded on [0, 1] (e.g. efficiencies, fractions). Data must be rescaled to (0, 1) before fitting. No closed-form MLE; fit by searching a 2-D Full_Factorial spill (the Cartesian product of candidate α, β). Pass to GoF_AIC and GoF_BIC with k = 2.

Returns: negative log-likelihood for Beta(α, β) given data on (0,1); returns 1E+15 for boundary or invalid values

Negative log-likelihood for Beta(alpha, beta). Data must be strictly on (0, 1). No closed-form MLE; search a 2-D Full_Factorial spill (the Cartesian product of candidate parameters).

```excel
=LAMBDA(data, alpha_param, beta_param, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    IFERROR(-SUMPRODUCT(LN(BETA.DIST(x, alpha_param, beta_param, FALSE))), 1E+15)
  ))
```

## `NLL_BetaPERT`

**How unlikely the BetaPERT distribution is to have generated this data at the given three-point bounds.**

Arguments:

- **data** — single-column data range
- **min_val** — lower bound (minimum possible value)
- **mode_val** — most likely value (peak of the PERT curve)
- **max_val** — upper bound (maximum possible value)
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_BetaPERT reparameterises the Beta distribution using the PERT convention: α = 1 + 4·(mode − min)/(max − min), β = 1 + 4·(max − mode)/(max − min). Data are rescaled to z = (x − min)/(max − min); the Jacobian correction log(1/(max − min)) converts the Beta log-likelihood back to the original data scale.

BetaPERT is the standard distribution for three-point cost and schedule estimates. Parameters are estimated directly from min, mode, max; no optimisation is needed. Pass to GoF_AIC and GoF_BIC with k = 3.

Returns: negative log-likelihood for BetaPERT(min, mode, max) given data; returns 1E+15 for invalid parameters

Negative log-likelihood for BetaPERT(min, mode, max). Parameters come from three-point estimates; no optimisation needed.

```excel
=LAMBDA(data, min_val, mode_val, max_val, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    range_, max_val - min_val,
    z, (x - min_val) / range_,
    alpha_p, 1 + 4 * (mode_val - min_val) / range_,
    beta_p,  1 + 4 * (max_val  - mode_val) / range_,
    pdf_z, BETA.DIST(z, alpha_p, beta_p, FALSE),
    pdf_x, pdf_z / range_,
    IFERROR(-SUMPRODUCT(LN(pdf_x)), 1E+15)
  ))
```

## `NLL_Exponential`

**How unlikely the Exponential distribution is to have generated this data at the given rate.**

Arguments:

- **data** — single-column data range (all values must be ≥ 0)
- **rate** — rate parameter λ = 1/mean (must be > 0)
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_Exponential evaluates −Σ ln f(xᵢ | λ) = −n·ln(λ) + λ·Σxᵢ via EXPON.DIST. The IFERROR sentinel maps invalid regions (rate ≤ 0, or data violating the ≥ 0 support) to 1E+15, keeping MIN and grid-search argmin robust. The MLE closed form is: rate = 1/AVERAGE(data).

Exponential fits well for memoryless waiting times, inter-arrival times, and time-to-failure data. Pass to GoF_AIC and GoF_BIC with k = 1.

Returns: negative log-likelihood for Exponential(rate) given data; returns 1E+15 for invalid parameter combinations

Negative log-likelihood for Exponential(rate = 1/mean). Data must be ≥ 0. Closed-form MLE: rate = 1/AVERAGE(data).

```excel
=LAMBDA(data, rate, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    IFERROR(-SUMPRODUCT(LN(EXPON.DIST(x, rate, FALSE))), 1E+15)
  ))
```

## `NLL_Gamma`

**How unlikely the Gamma distribution is to have generated this data at the given shape and rate.**

Arguments:

- **data** — single-column data range (all values must be > 0)
- **shape** — shape parameter α (must be > 0)
- **rate** — rate parameter β = 1/scale (must be > 0)
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_Gamma evaluates −Σ ln f(xᵢ | α, β) via GAMMA.DIST, converting rate to scale internally (scale = 1/rate). IFERROR maps overflow near shape → 0 to a large sentinel.

Gamma fits well for right-skewed positive data (insurance claims, rainfall, queuing times). No closed-form MLE; fit by searching a Full_Factorial spill. Pass to GoF_AIC and GoF_BIC with k = 2.

Returns: negative log-likelihood for Gamma(shape, rate) given data; returns 1E+15 for invalid parameter combinations

Negative log-likelihood for Gamma(shape, rate = 1/scale). Data must be > 0. No closed-form MLE; fit by searching a Full_Factorial spill.

```excel
=LAMBDA(data, shape, rate, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    scale, 1 / rate,
    IFERROR(-SUMPRODUCT(LN(GAMMA.DIST(x, shape, scale, FALSE))), 1E+15)
  ))
```

## `NLL_Lognormal`

**How unlikely the Lognormal distribution is to have generated this data at the given parameters.**

Arguments:

- **data** — single-column data range (all values must be > 0)
- **meanlog** — mean of ln(data), μ_ln
- **sdlog** — standard deviation of ln(data), σ_ln (must be > 0)
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_Lognormal evaluates −Σ ln f(xᵢ | μ_ln, σ_ln) using Excel's LOGNORM.DIST PDF. The IFERROR sentinel maps overflow or invalid regions (non-positive data or sdlog ≤ 0) to 1E+15, keeping MIN and grid-search argmin robust. The MLE closed form is: μ_ln = AVERAGE(LN(data)), σ_ln = STDEV.S(LN(data)).

Lognormal fits well when data is right-skewed and multiplicative (incomes, sizes, concentrations). Pass to GoF_AIC and GoF_BIC with k = 2.

Returns: negative log-likelihood for Lognormal(meanlog, sdlog) given data; returns 1E+15 for invalid parameter combinations

Negative log-likelihood for Lognormal(meanlog, sdlog). Data must be positive. MLE: meanlog = AVERAGE(LN(data)), sdlog = STDEV.S(LN(data)).

```excel
=LAMBDA(data, meanlog, sdlog, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    IFERROR(-SUMPRODUCT(LN(LOGNORM.DIST(x, meanlog, sdlog, FALSE))), 1E+15)
  ))
```

## `NLL_Normal`

**How unlikely the Normal distribution is to have generated this data at the given parameters.**

Arguments:

- **data** — single-column data range
- **mean** — distribution mean parameter μ
- **sd** — distribution standard deviation σ (must be > 0)
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_Normal computes −Σ ln f(xᵢ | μ, σ) where f is the Normal PDF evaluated via NORM.DIST. The IFERROR sentinel maps overflow or invalid regions (sd ≤ 0, or non-positive PDF values) to 1E+15, keeping MIN and grid-search argmin robust. Minimising NLL over (μ, σ) gives the MLE; for the Normal distribution the MLE is the sample mean and sample standard deviation.

The filter defaults to ISNUMBER(data) so blank cells in the input range are automatically excluded. Pass this value to GoF_AIC and GoF_BIC with k = 2.

Returns: negative log-likelihood for Normal(mean, sd) given data; returns 1E+15 for invalid parameter combinations

Negative log-likelihood for Normal(mean, sd). Minimise to fit by MLE. Closed-form MLE: mean = AVERAGE, sd = STDEV.S.

```excel
=LAMBDA(data, mean, sd, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    IFERROR(-SUMPRODUCT(LN(NORM.DIST(x, mean, sd, FALSE))), 1E+15)
  ))
```

## `NLL_Triangular`

**How unlikely the Triangular distribution is to have generated this data at the given bounds.**

Arguments:

- **data** — single-column data range
- **min_val** — lower bound of the triangular distribution
- **mode_val** — peak (most likely value)
- **max_val** — upper bound
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_Triangular implements the triangular PDF directly: 2(x − min) / (range × (mode − min)) on [min, mode] and 2(max − x) / (range × (max − mode)) on (mode, max]. The likelihood is non-differentiable at the mode, so MLE is non-standard; the conventional approach is to set min = data min, mode = data mode, max = data max.

Triangular is common in project risk and cost estimation (three-point estimates). Pass to GoF_AIC and GoF_BIC with k = 3.

Returns: negative log-likelihood for Triangular(min, mode, max) given data; returns 1E+15 for invalid parameters

Negative log-likelihood for Triangular(min, mode, max). Conventional MLE: set min, mode, max to the data min, mode, and max.

```excel
=LAMBDA(data, min_val, mode_val, max_val, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    range_, max_val - min_val,
    pdf, IF(
      x <= mode_val,
      2 * (x - min_val)  / (range_ * (mode_val - min_val)),
      2 * (max_val - x)  / (range_ * (max_val - mode_val))
    ),
    IFERROR(-SUMPRODUCT(LN(pdf)), 1E+15)
  ))
```

## `NLL_Weibull`

**How unlikely the Weibull distribution is to have generated this data at the given shape and scale.**

Arguments:

- **data** — single-column data range (all values must be > 0)
- **shape** — shape parameter k (must be > 0)
- **scale** — scale parameter λ (must be > 0)
- **filter** — optional boolean array; defaults to ISNUMBER(data)

NLL_Weibull evaluates −Σ ln f(xᵢ | k, λ) using WEIBULL.DIST. The IFERROR sentinel maps overflow or invalid regions (shape → 0, or data violating support) to 1E+15, keeping MIN and grid-search argmin robust.

Weibull is the standard model for failure times and wind speeds. No closed-form MLE exists; fit by searching a Full_Factorial spill. Pass to GoF_AIC and GoF_BIC with k = 2.

Returns: negative log-likelihood for Weibull(shape, scale) given data; returns 1E+15 for invalid parameter combinations

Negative log-likelihood for Weibull(shape, scale). Data must be > 0. No closed-form MLE; fit by searching a Full_Factorial spill.

```excel
=LAMBDA(data, shape, scale, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    IFERROR(-SUMPRODUCT(LN(WEIBULL.DIST(x, shape, scale, FALSE))), 1E+15)
  ))
```

## `Normal_Scores`

**The bell-curve quantiles used to build a Q-Q plot.**

Arguments:

- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Intended for normal probability plot x-values.

Returns: theoretical standard-normal scores as a spilled vector

Rankit scores Φ⁻¹((rank−0.5)/n); rank = strict-less count + 1. The count rounds both operands to 9 decimals so tied values collapse to one rank identically across Excel and the Python oracle, not split by sub-ULP float differences. Q-Q plot x-axis.

```excel
=LAMBDA(Y, [Include],
  LET(
    filtered, FILTER(Y, IF(ISOMITTED(Include), TRUE, Include)),
    n,        ROWS(filtered),
    rounded,  ROUND(filtered, 9),
    NORM.S.INV(
      (BYROW(rounded, LAMBDA(v, SUMPRODUCT((rounded < v) * 1))) + 0.5) / n
    )
  ))
```

## `Number_Of_Histogram_Bins`

**How many histogram bins a chosen rule (default: Freedman-Diaconis) recommends for this data.**

Arguments:

- **data** — single-column numeric data range
- **method** — bin-count rule: "Sturges", "Scott", or "FD" (default when omitted)
- **filter** — optional boolean array

Number_Of_Histogram_Bins returns the recommended number of histogram bins k using the specified rule, defaulting to Freedman-Diaconis when method is omitted.

All three methods are computed inline without calling other named LAMBDAs:
  Sturges: k = ⌈log₂(n) + 1⌉
  Scott:   width = 3.49 × SD × n^(−1/3);  k = ⌈range / width⌉  (k=1 when SD=0)
  FD:      width = 2 × IQR × n^(−1/3);   k = ⌈range / width⌉  (falls back to Sturges when IQR=0)

Used internally by Bin_Edges. Can also be called directly to get the bin count for a specific rule.

Returns: integer bin count k for the chosen method

Bin count by chosen rule ("Sturges", "Scott", or default "FD"). Used internally by Bin_Edges; call directly to inspect the count.

```excel
=LAMBDA(data, [method], [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    n, COUNT(x),
    m, IF(ISOMITTED(method), "FD", method),
    SWITCH(m,
      "Sturges", CEILING(LOG(n, 2) + 1, 1),
      "Scott",   LET(
                   sd, STDEV.S(x),
                   range_, MAX(x) - MIN(x),
                   IF(sd = 0, 1, CEILING(range_ / (3.49 * sd * n ^ (-1/3)), 1))
                 ),
      LET(
        iqr, QUARTILE.EXC(x, 3) - QUARTILE.EXC(x, 1),
        range_, MAX(x) - MIN(x),
        IF(iqr = 0,
          CEILING(LOG(n, 2) + 1, 1),
          CEILING(range_ / (2 * iqr * n ^ (-1/3)), 1))
      )
    )
  ))
```

## `Numeric_Complete_Cases`

**Listwise-deletion mask — 1 when every value in a row is numeric, 0 otherwise (the sample-construction primitive the rest of the bundle composes onto).**

Arguments:

- **data** — the data range to screen — any shape; each row is judged independently

Numeric_Complete_Cases returns a listwise-deletion sample mask: 1 for a row in which every value is numeric, 0 for any row containing a blank, text, boolean, or error cell. It is the primitive the rest of the standalone Data Transformation library composes onto — every other sample-construction helper in the bundle accepts an [include] mask, and this is the mask a caller builds before passing it in, so the primitive itself takes no [include] (its job is to BE the mask, not to consume one).

This is a detector, not a transform output, so a non-numeric row returns 0 — not #N/A and not a blank. The library's NA()-exception convention ("" means excluded, #N/A means included-but-incomputable) applies to transform outputs like Ln_Positive, where a row that is part of the sample yet cannot be computed must surface loudly. A mask has no such state: a row either is a complete numeric case (1) or it is not (0), and 0 is the signal downstream callers filter on. Returning 0 (plotted, falsy, summable) rather than "" also keeps the column a clean numeric array a SUMPRODUCT or further BYROW can reduce without type coercion.

The check is row-wise by design — "every value in a row" is the definition of listwise deletion — so BYROW judges each row independently and the result stays row-aligned to data. The all-numeric test is written as SUM(--ISNUMBER(r)) = COLUMNS(r) rather than AND(ISNUMBER(r)): AND aggregates a whole array to one scalar, which would collapse the per-row judgement, while the count-equals-width form stays inside the per-row LAMBDA and is explicit about what "complete" means. A single-column input works too (each row is 1×1, so the test reduces to ISNUMBER on the one cell).

Returns: n × 1 column, row-aligned to data: 1 when every value in the row is numeric (Excel ISNUMBER), 0 otherwise (a blank, text, #N/A, or boolean cell makes its row 0).

Row-wise listwise-deletion mask: 1 when every cell in the row is numeric (ISNUMBER), else 0. The primitive the [include]-taking transforms compose onto; takes no [include] itself. Detector, not a transform output — rows return 0, not #N/A.

```excel
=LAMBDA(data,
    BYROW(data, LAMBDA(r, N(SUM(--ISNUMBER(r)) = COLUMNS(r))))
)
```

## `Observation_Number`

**A simple 1-to-n row counter for the filtered data.**

Arguments:

- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Spills 1..n, where n is the number of filtered observations.

Returns: 1..n observation index as a spilled vector

Sequence 1..n for the filtered dataset. Use as an observation index axis on diagnostic scatter plots.

```excel
=LAMBDA(Y, [Include],
  SEQUENCE(ROWS(FILTER(Y, IF(ISOMITTED(Include), TRUE, Include))))
)
```

## `Observations`

**How many rows of data the model is using.**

Arguments:

- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

n (number of observations) is the count of data points available to the model after any filtering is applied. It is the fundamental sample size that drives all degrees-of-freedom calculations.

Returns ROWS(FILTER(Y, filt_arg)). X is not required because row count is fully determined by Y.

Returns: n — count of rows used in the regression as a scalar integer

Count of rows used in the regression after optional filtering. The fundamental sample size n driving all degrees-of-freedom and model-comparison calculations.

```excel
=LAMBDA(Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    ROWS(FILTER(Y, filt_arg))
  ))
```

## `P_Values`

**How surprising each coefficient would be if its true effect were zero.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Returns a vertical spill array of two-tailed p-values for each OLS coefficient, testing the null hypothesis that the true coefficient equals zero.

Computed as T.DIST.2T(ABS(T_Statistics(...)), Residual_Degrees_Of_Freedom(...)). An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to both — the t-statistic's SE and the distribution's df must use the same corrected residual df.

Returns: vertical array of two-tailed p-values for each coefficient

Two-tailed p-values per coefficient testing H₀: true effect = 0. Optional DF_Absorbed (default 0) corrects both the t-stat's SE and the test's df under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    T.DIST.2T(ABS(T_Statistics(X, Y, filt_arg, context_arg)), Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg))
  ))
```

## `Partial_Correlation`

**The predictor-outcome relationship left over after accounting for the other predictors.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Partial correlation for coefficient j = tⱼ / √(t²ⱼ + df_residual), the sign-preserving linear correlation between Y and Xⱼ after partialling out all other predictors from both Y and Xⱼ. Equals √(Partial_R_Squared) × sign(t).

Computed as T_Statistics / SQRT(T_Statistics² + Residual_Degrees_Of_Freedom), element-wise. Ranges from −1 to +1. An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to both.

Returns: vertical array of partial correlation coefficients — same order as Coefficients (intercept first when included)

Sign-preserving partial correlation between Y and each predictor after removing all other predictor effects. Range [−1, 1]. Optional DF_Absorbed (default 0) under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    t,  T_Statistics(X, Y, filt_arg, context_arg),
    df, Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg),
    t / SQRT(t^2 + df)
  ))
```

## `Partial_R_Squared`

**How much unique explanatory value each predictor adds after controlling for the rest.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Partial R² for coefficient j = t²ⱼ / (t²ⱼ + df_residual), the proportion of residual variance explained by predictor j after controlling for all other predictors. Equivalently, it is the squared partial correlation between Y and Xⱼ given all other Xₛ.

Computed as T_Statistics² / (T_Statistics² + Residual_Degrees_Of_Freedom), element-wise. The intercept row is included (row 1) for consistency with T_Statistics and P_Values. An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to both.

Returns: vertical array of partial R² values — same order as Coefficients (intercept first when included)

Unique proportion of residual variance explained by each predictor after controlling for all others. Range [0, 1]. Optional DF_Absorbed (default 0) under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    t,  T_Statistics(X, Y, filt_arg, context_arg),
    df, Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg),
    t^2 / (t^2 + df)
  ))
```

## `Pearson_R`

**The straight-line relationship between each predictor and the outcome.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Pearson_R computes the linear (Pearson product-moment) correlation between each column of Predictors and Y after applying the optional Include argument. When Predictors has a single column the result is a scalar; when Predictors has k columns it spills as a k×1 column vector.

A Pearson_R close to ±1 indicates a strong linear association. Comparing Pearson_R with Spearman_R on the same data reveals nonlinearity: a large gap (Spearman > Pearson in magnitude) suggests a monotonic but curved relationship that may warrant a log or polynomial transformation.

Returns: k×1 column vector of Pearson correlations between each predictor and Y

Linear (Pearson) correlations between each Predictors column and Y. Returns a k×1 vector. Compare with Spearman_R to detect nonlinearity.

```excel
=LAMBDA(Predictors, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    xf, FILTER(Predictors, filt_arg),
    yf, FILTER(Y,          filt_arg),
    MAKEARRAY(COLUMNS(Predictors), 1, LAMBDA(r, c, CORREL(CHOOSECOLS(xf, r), yf)))
  ))
```

## `Prediction_Group_Column` *(sheet-scoped: Regression)*

**Which column groups observations for prediction purposes — the Fixed Effects variable when one is declared, or 'everyone is one group' when none is.**

Prediction_Group_Column is the grouping key behind the v2.1 group-mean-recovery prediction (Group_Prediction_Interval): with a Fixed Effects row declared, it is Fixed_Effects_Column(); with none, it is a constant "(all)" column treating the entire sample as one group.

The constant fallback is what makes the group-mean recovery formula ŷ = ȳᵢ + (x_new − x̄ᵢ)′β̂ collapse exactly to the ordinary no-FE prediction: with one universal group, ȳᵢ is the grand mean of y, x̄ᵢ is the grand mean of each predictor, and Tᵢ is the full included count — precisely the quantities an intercept-based OLS prediction already reduces to. Uses the same "(all)" single-group idiom as Sequence_Deltas' no-Identifier fallback (EXPAND("(all)", n, 1, "(all)")) — one degenerate-group sentinel across the catalog.

Deliberately NOT the same object as Design_Response()/Design_Columns()'s own fe_active branch: those return the RAW (uncentered) Response_Column()/X() when inactive, preserving the v2.0 model's own separately-fit intercept exactly. Group_Prediction_Interval needs a properly group-demeaned (here, grand-mean-centered) matrix for its variance term regardless of FE status, so it always re-demeans by this column rather than reusing Design_Columns() directly.

Returns: The Fixed Effects group column when one is declared; a constant "(all)" column (every row, one universal group) when none is — full height either way.

Grouping key for group-mean-recovery prediction: Fixed_Effects_Column() when FE is declared, else a constant "(all)" column (one universal group). Makes the no-FE case a degenerate collapse of the same formula.

```excel
=LAMBDA(LET(
    n_c,       COLUMNS(Source_Data),
    fe_active, SUMPRODUCT(N(TAKE(Spec_Role, n_c) = "Fixed Effects")) > 0,
    IF(fe_active,
       Fixed_Effects_Column(),
       EXPAND("(all)", ROWS(Source_Data), 1, "(all)")
    )
))
```

## `Prediction_Interval`

**A predicted outcome plus the range a new real-world observation is likely to fall in.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **X_new** — column vector of predictor values for the new observation — intercept term first (=1 or 0) when Allow_Intercept is TRUE, matching the layout of pred_input on the Regression sheet
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **alpha** — significance level (0–1); default 0.05 yields 95% prediction intervals
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Returns a 6-element VSTACK: (1) point estimate = x_new · beta, (2) SE of prediction = s √(1 + x_new'(X'X)⁻¹x_new), (3) t critical value at alpha, (4) lower prediction bound, (5) upper prediction bound, (6) confidence level = 1 − alpha.

When Has_Intercept is FALSE the first element of X_new (the intercept placeholder) is dropped before the dot product and leverage calculation to align dimensions with the k-column design matrix. Leverage h_new is computed as x_new'(X'X)⁻¹x_new using Gram_Inverse(X_design) and MMULT.

An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to both s and df, so the t-critical value and the prediction SE reflect the correct absorbed-group residual df. This function alone still assumes x_new lives in the same raw covariate space as X — it is not a well-defined FE prediction on its own under Fixed Effects; Group_Prediction_Interval is the group-mean-relative sibling that adds the ȳᵢ/x̄ᵢ recovery machinery.

Returns: 6-element vertical array: [Point Estimate, SE Prediction, t Critical, Lower bound, Upper bound, Confidence Level]

6-element vector: point estimate, SE_prediction, t-critical, lower/upper PI bound, confidence level. Optional DF_Absorbed (default 0) corrects df/SE under Fixed Effects.

```excel
=LAMBDA(X, Y, X_new, [Include], [alpha], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg,      IF(ISOMITTED(Include),       TRUE, Include),
    alpha_arg,     IF(ISOMITTED(alpha),         0.05, alpha),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    X_design,      Design_Matrix(X, filt_arg),
    beta,          Coefficients(X, Y, filt_arg),
    s,             SE_Regression(X, Y, filt_arg, context_arg),
    df,            Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg),
    t_crit,        T.INV.2T(alpha_arg, df),
    x_new_aligned, IF(has_arg, X_new, INDEX(X_new, SEQUENCE(ROWS(X_new) - 1, 1, 2, 1))),
    XtX_inv,       Gram_Inverse(X_design),
    h_new,         MMULT(MMULT(TRANSPOSE(x_new_aligned), XtX_inv), x_new_aligned),
    point_est,     SUMPRODUCT(x_new_aligned, beta),
    se_pred,       s * SQRT(1 + h_new),
    margin,        t_crit * se_pred,
    VSTACK(
      point_est,
      se_pred,
      t_crit,
      point_est - margin,
      point_est + margin,
      1 - alpha_arg
    )
  ))
```

## `Predictions`

**The model's fitted Y values for each row.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Computes X_design * beta using MMULT.

Returns: fitted Y values as a spilled vector

Model-fitted ŷ values (X·β) for each filtered row. Subtract from Y to get Residuals.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    MMULT(
      Design_Matrix(X, filt_arg),
      Coefficients(X, Y, filt_arg)
    )
  ))
```

## `Predictor_Columns` *(sheet-scoped: Regression)*

**The design matrix built from the spec — continuous columns as-is (or natural-logged), categorical columns turned into reference-dropped dummies (never logged).**

Sheet-scoped design-matrix constructor for the Regression sheet. A REDUCE assembles columns in spec order for rows where Role=Predictor and Include=TRUE. Each row's own block is built by the LET-bound blk(): a Continuous predictor is its raw column — or Ln_Positive(column) when that row's Transform (spec column G) is Log — and a Categorical predictor is --(col = retained_levels), where Dummy_Levels supplies the mask-scoped, reference-dropped level set (bound once; its #N/A on a degenerate or invalid-reference column is the single skip signal, so the variable contributes nothing rather than erroring the sheet). The Categorical branch never reads Transform: Log on a Categorical Predictor is disallowed and flagged red on the sheet rather than silently applied. From v3.1 the row's block is followed by its INTERACTION columns: mate() resolves spec column M (Interaction Term) to the operand's spec row — 0 when M or N is blank, when the name matches no column, or when the operand's Role is not Predictor, in which case the row contributes its main effect only — and the operand block is built by the SAME blk(), so an operand that is itself Categorical, Log-transformed, or excluded (the flagged-amber marginality case) encodes exactly as it would on its own. The two blocks are combined pairwise by a nested REDUCE, giving COLUMNS(a)*COLUMNS(b) columns: 1 for Continuous x Continuous, L-1 for Continuous x Categorical, (L1-1)(L2-1) for Categorical x Categorical. The operation (spec column N) is the closed Product | Difference | Ratio vocabulary; Ratio returns NA() on a zero denominator rather than a bare #DIV/0!. A row pointing at itself under Product is the documented quadratic term. Two-way only — one operand per spec row, so (AxB)xC is not expressible. Reads Log_Drop_Sample_Include_Calc only to fix categorical level sets and to gate the Log transform's row mask; the output is always full-height (the row-mask contract). The retained-levels broadcast is inlined, never LET-bound, to avoid an eager empty-array error. The degenerate skip tests ISNA(INDEX(arr,1,1)), a scalar, and only on the Categorical branch: lv is a 1x(L-1) row, and an array ISNA condition in front of a wider HSTACK branch broadcasts to #N/A.

Returns: The constructed design matrix in spec order: continuous predictors passed through (Ln_Positive-transformed when that row's Transform is Log), categoricals reference-dropped one-hot (Transform ignored — see the spec block's red flag for a Categorical row with Log set), each followed immediately by any interaction columns that row declares; full height.

Design matrix in spec order: continuous raw or Ln_Positive per Transform=Log; categoricals reference-dropped one-hot via Dummy_Levels; each row's interaction columns (spec M/N) follow its own block. Full-height; degenerate categoricals skipped.

```excel
=LAMBDA(LET(n_c,COLUMNS(Source_Data),rl,TAKE(Spec_Role,n_c),inc,TAKE(Spec_Include,n_c),typ,TAKE(Spec_Type,n_c),refs,TAKE(Spec_Reference,n_c),trn,TAKE(Spec_Transform,n_c),itm,TAKE(Spec_Interaction_Term,n_c),iop,TAKE(Spec_Interaction_Operation,n_c),hdrs,TOROW(Header_Names),mate,LAMBDA(j,LET(t,INDEX(itm,j),o,INDEX(iop,j),q,IFERROR(XMATCH(t,hdrs),0),IF(OR(LEN(t&"")=0,LEN(o&"")=0,q=0),0,IF(INDEX(rl,q)<>"Predictor (x)",0,q)))),keep,LAMBDA(x,arr,IF(INDEX(typ,x)<>"Categorical",TRUE,NOT(ISNA(INDEX(arr,1,1))))),seed,SEQUENCE(ROWS(Source_Data),1,0,0),si,Log_Drop_Sample_Include_Calc(),blk,LAMBDA(x,IF(INDEX(typ,x)<>"Categorical",LET(col,INDEX(Source_Data,0,x),IF(OR(INDEX(trn,x)="Log",INDEX(trn,x)="Log (drop ≤ 0)"),Ln_Positive(col,si),col)),LET(col,INDEX(Source_Data,0,x),d,INDEX(refs,x),r,IF(LEN(d&"")=0,"",d),lv,Dummy_Levels(col,r,si),IF(ISNA(INDEX(lv,1,1)),NA(),--(col=lv))))),built,REDUCE(seed,SEQUENCE(n_c),LAMBDA(acc,j,IF(OR(INDEX(rl,j)<>"Predictor (x)",INDEX(inc,j)<>TRUE),acc,LET(a,blk(j),IF(NOT(keep(j,a)),acc,LET(m,HSTACK(acc,a),q,mate(j),IF(q=0,m,LET(b,blk(q),IF(NOT(keep(q,b)),m,LET(o,INDEX(iop,j),REDUCE(m,SEQUENCE(COLUMNS(a)),LAMBDA(p,ai,REDUCE(p,SEQUENCE(COLUMNS(b)),LAMBDA(pp,bi,HSTACK(pp,SWITCH(o,"Product",INDEX(a,0,ai)*INDEX(b,0,bi),"Difference",INDEX(a,0,ai)-INDEX(b,0,bi),"Ratio",IFERROR(INDEX(a,0,ai)/INDEX(b,0,bi),NA()),NA())))))))))))))))),DROP(built,,1)))
```

## `PRESS`

**How well the model predicts rows it did not get to train on.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

PRESS is the leave-one-out cross-validation (LOOCV) sum of squared prediction errors. For each observation i the model is refitted on all other observations and evaluated at x_i; the shortcut formula avoids n separate refits by computing Σ(eᵢ / (1 − hᵢ))² where eᵢ are the OLS residuals and hᵢ are the hat-matrix diagonal (leverage) values.

Computed as SUMSQ(LOOCV_Residual(...)), where LOOCV_Residual returns eᵢ / (1 − hᵢ) using Residuals() and Hat_Diagonal().

Returns: PRESS (Prediction Residual Error Sum of Squares) as a scalar

Leave-one-out cross-validation sum of squares = Σ(eᵢ/(1−hᵢ))². Lower values mean better out-of-sample predictions.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    SUMSQ(LOOCV_Residual(X, Y, filt_arg))
  ))
```

## `QQ_Correlation`

**How close the residual distribution is to a normal bell shape.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

CORREL(Scaled_Residuals_Ranked, Normal_Scores) — the Filliben approximation to the Shapiro-Wilk statistic. Values near 1 indicate approximately normal residuals; values below ~0.99 suggest non-normality. The Q-Q plot is the corresponding visual: plot Normal_Scores (x) vs Scaled_Residuals_Ranked (y) and look for alignment with the identity line.

Returns: Filliben Q-Q correlation coefficient (normality statistic) scalar

Filliben Q-Q correlation between ranked scaled residuals and normal quantiles. Optional DF_Absorbed (default 0) under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    sr, Scaled_Residuals_Ranked(X, Y, filt_arg, context_arg),
    n,  ROWS(sr),
    ns, NORM.S.INV((SEQUENCE(n) - 0.5) / n),
    CORREL(sr, ns)
  ))
```

## `R_Squared`

**How much of the outcome's movement the model explains.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

R² is the proportion of variation in the response explained by the model — 1 - SS_Residual / SS_Total.

Computed from the sums of squares rather than read out of LINEST. That is deliberate and load-bearing: since the v3.0 intercept relocation every LINEST call passes const = FALSE (the design matrix supplies its own intercept column), and under const = FALSE LINEST's own R² is the UNCENTERED R², computed against Σy² instead of DEVSQ(y). Reading it would silently change R², Adjusted R², Multiple R and the F-statistic for every model with an intercept.

Because SS_Total is itself the intercept-only residual sum of squares, the centered and uncentered cases both fall out of the one expression: with an intercept the denominator is DEVSQ(y), without one it is SUMSQ(y).

Returns: R² (coefficient of determination) as a scalar in [0, 1]

Coefficient of determination, 1 - SS_Residual/SS_Total. Derived from the sums of squares, NOT from LINEST — under const = FALSE LINEST reports the uncentered R².

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg, IF(ISOMITTED(Include),       TRUE, Include),
    1 - SS_Residual(X, Y, filt_arg) / SS_Total(X, Y, filt_arg, context_arg)
  ))
```

## `Rank_Fraction`

**Each row's percentile-style position within the filtered outcome values.**

Arguments:

- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Empirical CDF: fraction of filtered observations with value <= each observation. Equivalent to SUMPRODUCT((filtered<=v)*1)/n for each v. Output is in original data order.

Returns: empirical-CDF fractions as a spilled vector, in original data order

Empirical CDF fractions for filtered Y in original data order. Each value is the row’s percentile within the filtered outcome.

```excel
=LAMBDA(Y, [Include],
  LET(
    filtered, FILTER(Y, IF(ISOMITTED(Include), TRUE, Include)),
    n,        ROWS(filtered),
    BYROW(filtered, LAMBDA(v, SUMPRODUCT((filtered <= v) * 1))) / n
  ))
```

## `Regression_Degrees_Of_Freedom`

**How many predictor slots the model is trying to fit.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

df_regression is the number of degrees of freedom attributable to the regression model — equal to k, the number of predictor columns. The intercept term is not counted here; it is accounted for in df_total.

Returns COLUMNS(X) - N(has_arg). Since the v3.0 intercept relocation the design matrix carries its own intercept column, so the count has to be corrected back to the ANOVA convention; Has_Intercept identifies the column rather than switching a behaviour.

Returns: df_regression (k, number of predictors) as a scalar integer

Degrees of freedom for regression = number of predictors k (= COLUMNS(X) less the intercept column, if the design matrix has one). Used as the numerator df in the F-test.

```excel
=LAMBDA(X, [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    COLUMNS(X) - N(has_arg)
  ))
```

## `Residual_Degrees_Of_Freedom`

**How much data is left to estimate random error after fitting the model, minus any degrees of freedom a Fixed Effects variable has already absorbed.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

df_residual (also called df_error) is the degrees of freedom remaining after the regression model has been fit. With an intercept it equals n - k - 1; without an intercept it equals n - k. It is the denominator degrees of freedom for the F-test and the penalty term in Adjusted R².

Computed as Observations(Y, filt_arg) - COLUMNS(X) - absorbed_arg. Because the design matrix carries its own intercept column, COLUMNS(X) already counts every fitted parameter and no separate intercept correction is needed — which is why this function does not take Has_Intercept.

Returns: df_residual (n-k-1 with intercept, n-k without, minus DF_Absorbed) as a scalar integer

Residual df = Total_Degrees_Of_Freedom − Regression_Degrees_Of_Freedom − DF_Absorbed (default 0). Denominator df for the F-test, Adjusted R², and — new in v2.1 — the single source of truth for FE df correction downstream.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    Observations(Y, filt_arg) - COLUMNS(X) - absorbed_arg
  ))
```

## `Residuals`

**How far each actual Y value is above or below its fitted value.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Residuals are Y - fitted Y.

Returns: raw residuals as a spilled vector

Raw residuals = Y − ŷ for each filtered row. Starting point for all diagnostic plots.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    FILTER(Y, filt_arg) - Predictions(X, Y, filt_arg)
  ))
```

## `Response_Column` *(sheet-scoped: Regression)*

**The y column — whichever variable is marked Response in the spec, natural-logged first if that row's Transform is Log.**

Sheet-scoped derived response for the Model Construction sheet. Returns the data column of the first spec row whose Role is Response, via XMATCH over the trimmed role vector, then applies Ln_Positive to it when that row's Transform (spec column G) is Log. Full-height; consumers wrap it in IFERROR to absorb the no-Response and multiple-Response states, which the audit strip's responses count flags separately.

The transform is applied here — inside the accessor itself — rather than via a separate wrapper, so every existing consumer (Intercept_Only_Point/S, Pearson_R/Spearman_R, Durbin_Watson_By, Group_Prediction_Interval, Group_Mean_At, and Design_Response()) automatically fits in log space the instant Log is declared, with no call-site changes. This differs from the v2.1 FE precedent (Design_Response() wraps Response_Column() rather than modifying it) because FE demeaning has genuine raw-response consumers (the zero-predictor Intercept_Only_* branch needs the undemeaned response to fit its own intercept); a declared Transform has none — every consumer of the response wants it in the space the model spec declares. No Type gate is applied on the Response row: Type (column D) is itself hidden-in-place on Response rows by the spec block's own cascading conditional formatting, so gating on an invisible cell would be a trap; Ln_Positive's own #N/A-on-non-numeric behavior is the only guard needed.

Returns: The single Role=Response data column, full height; #N/A when no row is declared Response. When that row's Transform is Log, every value is Ln_Positive(value) instead of the raw value.

Derived y for the Model Construction sheet: the data column of the first Role=Response spec row, Ln_Positive-transformed in place when that row's Transform is Log. #N/A when none is declared; the audit strip flags zero or multiple responses.

```excel
=LAMBDA(LET(n_c,COLUMNS(Source_Data),rl,TAKE(Spec_Role,n_c),trn,TAKE(Spec_Transform,n_c),p,XMATCH("Response (y)",rl),col,INDEX(Source_Data,0,p),IF(OR(INDEX(trn,p)="Log",INDEX(trn,p)="Log (drop ≤ 0)"),Ln_Positive(col,Log_Drop_Sample_Include_Calc()),col)))
```

## `Role_Status` *(sheet-scoped: Regression)*

**Checks the Role column: exactly one Response, at most one Fixed Effects row. Blank when the spec is legal.**

Role_Status is the B2 readout: role cardinality, in severity order. A model needs exactly one Response (y) row — with none there is nothing to fit, and with two the constructor silently takes the first, so both states are errors rather than warnings. Sequence and Fixed Effects each allow zero, so only a second row of either is flagged; two Fixed Effects rows would be two-way absorption, which this workbook does not implement (the engine absorbs the first and ignores the second).

The severity order lives inside the function, which is why B2 needs only one conditional-format rule — the cell is either blank or carrying the most severe message that applies.

SHEET-SCOPED (scope: Regression), like every name it reads. Spec_Role and Source_Data are unqualified, so they resolve against the sheet the calling formula lives on and each Regression sheet gets its own verdict — the same reason Base_Period_Delta is sheet-scoped.

Returns: The Role-column verdict for this sheet's spec: an ERROR string when the Response count is not exactly one or a second Fixed Effects row is declared, otherwise "".

B2 role-cardinality verdict. ERROR when the Response count is not exactly one, or a second Fixed Effects row is declared; "" otherwise. Severity order is inside the function. Sheet-scoped.

```excel
=LAMBDA(
    IF(SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response (y)"))=0,"ERROR: no Response (y) row — mark the variable being modeled.",IF(SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Response (y)"))>1,"ERROR: multiple Response (y) rows — mark exactly one.",IF(SUMPRODUCT(N(TAKE(Spec_Role,COLUMNS(Source_Data))="Fixed Effects"))>1,"ERROR: multiple Fixed Effects rows — mark at most one.","")))
)
```

## `Row_Labels` *(sheet-scoped: Regression)*

**A readable label per row — the Identifier columns joined, or Obs. 1, Obs. 2, … when there are none.**

Sheet-scoped observation labels for the Model Construction sheet. Dispatches structurally on whether any spec row has Role=Identifier: with none, positional labels 'Obs. 1', 'Obs. 2', …; with some, a per-row TEXTJOIN of all Identifier columns in table order, '|'-separated with ignore_empty=FALSE so field positions stay aligned when an identifier cell is blank. Full-height by the same row-mask contract as X and Sample_Include.

Returns: Full-height text column of observation labels: joined Identifier columns, or 'Obs. n' when none are declared.

Observation labels for the Model Construction sheet: Identifier columns joined by '|', or positional 'Obs. n' when no Identifier is declared. Full-height.

```excel
=LAMBDA(LET(n_c,COLUMNS(Source_Data),rl,TAKE(Spec_Role,n_c),ids,IFERROR(TRANSPOSE(FILTER(TRANSPOSE(Source_Data),rl="Identifier (Row Label)")),NA()),IF(SUM(--(rl="Identifier (Row Label)"))=0,"Obs. "&SEQUENCE(ROWS(Source_Data)),BYROW(ids,LAMBDA(r,TEXTJOIN("|",FALSE,r))))))
```

## `Sample_Include` *(sheet-scoped: Regression)*

**The model's ordinary eligible rows, before the special "Log (drop ≤ 0)" rule removes anything.**

Sample_Include is the public ordinary-eligibility mask. It returns Sample_Include_Calc() directly: Filters have been applied and required model variables have been checked for numeric values, but transform-specific row dropping has not occurred.

A plain "Log" transform never changes Sample_Include: zero and negative values remain eligible and subsequently cause Ln_Positive to return #N/A. Likewise, a row that will later be removed under "Log (drop ≤ 0)" is still TRUE here if it otherwise satisfies the model's Filters and numeric-data requirements.

The actual sample entering the regression is exposed separately through Fit_Sample_Include(), whose materialized source is Log_Drop_Sample_Include_Calc(). Keeping those concepts separate makes a call to Sample_Include() mean ordinary sample eligibility rather than silently including transform-specific filtering.

SHEET-SCOPED (scope: Regression); delegates to this sheet's Sample_Include_Calc closure.

Returns: The ordinary model-sample inclusion mask, before any rows are removed by "Log (drop ≤ 0)".

Public ordinary-sample mask. No arguments. Transform-driven row dropping is handled separately.

```excel
=LAMBDA(
    Sample_Include_Calc()
)
```

## `Sample_Include_Calc` *(sheet-scoped: Regression)*

**Which rows are otherwise usable by the model: Filters pass and all required numeric model variables contain numbers.**

Sample_Include_Calc is the computational leaf for ordinary model-sample eligibility. It applies only the rules that determine whether a source row is otherwise usable by the declared model. Filter rows must evaluate to 1 or TRUE. The Response (y) and every included Continuous Predictor (x) must contain numeric values.

Transforms do not participate here. In particular, neither "Log" nor "Log (drop ≤ 0)" changes this mask. A strict Log with a zero or negative therefore remains eligible here and is allowed to fail later in Ln_Positive. The explicit dropping behaviour of "Log (drop ≤ 0)" belongs to Log_Drop_Sample_Include_Calc instead.

Full-height by contract: it never row-filters; it returns one Boolean per source row. Filter truthiness coerces with (Source_Column+0), not N(Source_Column), because N() on a bare range reference may implicit-intersect to a scalar.

SHEET-SCOPED (scope: Regression): Source_Data and the Spec_* arrays resolve against the calling sheet.

Returns: A full-height Boolean row mask identifying rows ordinarily eligible for the model before any transform-specific row dropping is applied.

Ordinary eligibility mask only: Filters + numeric response/included continuous predictors. No transform-driven row dropping.

```excel
=LAMBDA(
    LET(
      Number_Of_Columns,COLUMNS(Source_Data),
      Spec_Roles,TAKE(Spec_Role,Number_Of_Columns),
      Spec_Includes,TAKE(Spec_Include,Number_Of_Columns),
      Spec_Types,TAKE(Spec_Type,Number_Of_Columns),
      Initial_Mask,SEQUENCE(ROWS(Source_Data),1,1,0),
      Sample_Mask,
        REDUCE(
          Initial_Mask,
          SEQUENCE(Number_Of_Columns),
          LAMBDA(Accumulated_Mask,Column_Number,
            LET(
              Source_Column,INDEX(Source_Data,0,Column_Number),
              IF(
                INDEX(Spec_Roles,Column_Number)="Filter",
                Accumulated_Mask*--IFERROR((Source_Column+0)=1,FALSE),
                IF(
                  OR(
                    INDEX(Spec_Roles,Column_Number)="Response (y)",
                    AND(
                      INDEX(Spec_Roles,Column_Number)="Predictor (x)",
                      INDEX(Spec_Includes,Column_Number)=TRUE,
                      INDEX(Spec_Types,Column_Number)="Continuous"
                    )
                  ),
                  Accumulated_Mask*N(ISNUMBER(Source_Column)),
                  Accumulated_Mask
                )
              )
            )
          )
        ),
      Sample_Mask=1
    )
)
```

## `Scaled_Residuals`

**Residuals converted to a common error scale.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Residual divided by SE_Regression. This is not Excel Analysis ToolPak 'Standard Residuals' and is not leverage-adjusted/studentized.

Returns: residual divided by SE_Regression as a spilled vector

Residuals divided by SE_Regression. Not leverage-adjusted. Optional DF_Absorbed (default 0) corrects SE_Regression under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    Residuals(X, Y, filt_arg)
      / SE_Regression(X, Y, filt_arg, context_arg)
  ))
```

## `Scaled_Residuals_Ranked`

**Scaled residuals sorted from smallest to largest.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Sorted Scaled_Residuals, intended for plotting against Normal_Scores.

Returns: sorted scaled residuals as spilled vector

Scaled residuals sorted ascending. Paired with Normal_Scores for a standard Q-Q residual plot. Optional DF_Absorbed (default 0) under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    SORT(Scaled_Residuals(X, Y, filt_arg, context_arg))
  ))
```

## `Scott_Bins`

**How many histogram bins Scott's normal-reference rule recommends.**

Arguments:

- **data** — single-column numeric data range
- **filter** — optional boolean array

Scott_Bins computes the bin width h = 3.49 × σ × n^(−1/3) that minimises mean integrated squared error assuming the data are normally distributed, then returns k = ⌈range / h⌉.

Scott's rule is more adaptive than Sturges for larger samples and performs well when the data are approximately normal. For strongly skewed or multimodal data, Freedman-Diaconis is more appropriate.

Returns: integer bin count via Scott's normal reference rule

Bin count from Scott’s normal-reference rule: width = 3.49·σ·n^(−1/3). Optimal when data are approximately normal.

```excel
=LAMBDA(data, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    n, COUNT(x),
    sd, STDEV.S(x),
    width, 3.49 * sd * n ^ (-1/3),
    range_, MAX(x) - MIN(x),
    IF(sd = 0, 1, CEILING(range_ / width, 1))
  ))
```

## `SE_Coefficients`

**How uncertain each fitted coefficient is — rescaled under Fixed Effects since Excel's LINEST can't be told about the absorbed groups directly.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Returns the standard errors of the fitted coefficients, in design-matrix column order.

LINEST is called with const = FALSE for the same reason as Coefficients: the intercept is a column of X, not something Excel should add. Under const = FALSE LINEST computes its residual df as n - COLUMNS(X), which is the correct value once the intercept is counted among the columns.

Under fixed effects the absorbed groups are not visible to LINEST, so the standard errors are rescaled by SQRT(naive_df / true_df) rather than reimplementing the fit.

Returns: vertical array of standard errors of the OLS coefficients — same order as Coefficients

Standard errors of the OLS coefficients. LINEST can't accept an external df, so DF_Absorbed (default 0) rescales its output by SQRT(naive_df/true_df) — exact, since SSR and (X'X)⁻¹ don't change.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    ls,       LINEST(FILTER(Y, filt_arg), FILTER(X, filt_arg), FALSE, TRUE),
    ses,      INDEX(ls, 2),
    k,        COLUMNS(X),
    naive_df, Residual_Degrees_Of_Freedom(X, Y, filt_arg),
    true_df,  naive_df - absorbed_arg,
    scale,    SQRT(naive_df / true_df),
    TRANSPOSE(CHOOSECOLS(ses, SEQUENCE(1, k, k, -1))) * scale
  ))
```

## `SE_Regression`

**The typical size of the model's prediction mistakes, in Y units.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

SE_regression (also called the root mean square error or S) measures the typical distance between observed Y values and the regression line, in the same units as Y. It equals √(SS_residual / df_residual) = √MSE. Smaller values indicate a tighter fit.

Computed as SQRT(SS_Residual(...) / Residual_Degrees_Of_Freedom(...)). An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads through so a Fixed Effects model's typical residual size is computed on the correct (absorbed) df, not an inflated one.

Returns: SE_regression (standard error of the regression) as a scalar

Standard error of the regression (root MSE = √(SS_Residual / df_residual)). Typical prediction error in Y units. Optional DF_Absorbed (default 0) corrects df_residual under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    SQRT(SS_Residual(X, Y, filt_arg) / Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg))
  ))
```

## `Sequence_Column` *(sheet-scoped: Regression)*

**The column marked as the Sequence (ordering) axis in the spec — whichever variable has its Sequence flag set.**

Sheet-scoped accessor for the declared ordering axis on the Regression sheet. Returns the data column of the variable whose Sequence flag (spec column H) is TRUE, via XMATCH over the TAKE-trimmed Spec_Sequence vector — the structural sibling of Response_Column (which does the same for Role=Response). Full-height by the same row-mask contract as the other sheet closures.

It is the seq input for sequence-aware serial-correlation diagnostics (Durbin_Watson_By), keeping the axis a single canonical name rather than an inline XMATCH repeated at every call site. When no variable is flagged as the Sequence axis, XMATCH does not match and the accessor is #N/A — the caller's gate (the zero-or-one Sequence-flag count) shows the not-applicable token before this is ever evaluated. When two or more are flagged (a spec error surfaced at the H2 status line) it resolves the first flagged column. Spec_Sequence is TAKE-trimmed to COLUMNS(Source_Data) so cells below the live spec rows, including the Sequence Spacing block, are never scanned.

Returns: The single Sequence-flagged data column, full height; #N/A when no variable is flagged as the Sequence axis.

Derived ordering axis for the Regression sheet: the data column of the Sequence-flagged spec row (column H). #N/A when none is flagged. Feeds Durbin_Watson_By and future serial-correlation diagnostics.

```excel
=LAMBDA(LET(n_c,COLUMNS(Source_Data),INDEX(Source_Data,0,XMATCH(TRUE,TAKE(Spec_Sequence,n_c),0))))
```

## `Sequence_Delta_Spectrum` *(sheet-scoped: Regression)*

**A little table of every gap size found between consecutive time periods and how often each occurs — all 1s means an evenly spaced panel.**

Sequence_Delta_Spectrum tabulates the delta spectrum: every distinct within-group consecutive spacing of the sequence axis alongside how many times it occurs, sorted ascending. It is the display behind the Sequence Spacing block's Delta/Count table and the evidence base for its verdict flags.

Reading it: a complete yearly panel shows the single row {1, n−G}. A punched-out year adds a {2, 1} row — the Regularity flag fires, but 2 is still on the Δ = 1 grid so the Off-grid flag stays quiet. First-of-month calendar serials show the signature cluster {28, 29, 30, 31} — Regularity and Off-grid both fire and the calendar-signature guidance recommends an integer period index (YEAR, or YEAR*12 + MONTH) upstream instead of quantizing day counts to a scalar Δ.

Returns #N/A when there are no within-group spacings (no Sequence axis flagged, or every group is a single observation); the consuming cells wrap IFERROR to degrade to a quiet blank.

Returns: Two-column array — each distinct within-group spacing with its count, ascending (the delta spectrum); #N/A when there are no spacings.

The delta spectrum: distinct within-group spacings of the sequence axis with counts, ascending. Drives the Sequence Spacing block's Delta/Count display and verdict flags. #N/A when no spacings exist.

```excel
=LAMBDA(
    LET(
        d, Sequence_Deltas(),
        u, SORT(UNIQUE(d)),
        IF(COUNT(d) = 0, NA(), HSTACK(u, BYROW(u, LAMBDA(v, SUM(--(d = v))))))
    )
)
```

## `Sequence_Deltas` *(sheet-scoped: Regression)*

**All the observed gaps between consecutive time periods within each group — e.g. all 1s for a complete yearly panel, with a 2 appearing where a year is missing.**

Sequence_Deltas returns every within-group consecutive spacing of the sequence axis: for each group, the differences between successive observed values of the Sequence-flagged column, pooled across groups into one column. It is the raw material for the Base Period Δ candidate (its MODE) and the delta spectrum display.

Seam-safety is the point of the implementation: rows are sorted by (group, seq) and consecutive diffs are kept only where the sorted neighbors share a group, so group seams can never contaminate the spacings. A pooled column-wise diff over the raw row order is forbidden — on Country/Year panel data it would manufacture large negative seam deltas at every country boundary. Zero spacings (duplicate (group, seq) pairs) are also dropped: the spectrum describes the time grid, and duplicates are a data defect surfaced elsewhere.

Groups are keyed by the Identifier-role columns, excluding the sequence column itself (a Year flagged as both Identifier and Sequence must not group by itself), joined with CHAR(31) when there are several. With no Identifier columns the whole sample is one group. A row whose Identifier fields are ALL blank produces a delimiter-only join key, which the emptiness check strips (SUBSTITUTE off the CHAR(31) separators) so unidentified rows never form a spurious group. The spectrum deliberately ignores Sample_Include(): the time grid is dataset structure, while the mask is model-iteration state — a filtered model must not make the data look gappier than it is.

All spec-band reads are TAKE-trimmed to COLUMNS(Source_Data), so the Sequence Spacing block below the spec rows is never scanned.

Returns: Column of all within-group consecutive positive spacings of the Sequence-flagged column (the raw material of the delta spectrum); #N/A when no Sequence axis is flagged or no group has two observations.

Within-group consecutive spacings of the Sequence-flagged column, seam-safe (sorted per group; seams and duplicates dropped). Feeds the Base Period Δ candidate (MODE) and the delta spectrum. #N/A when no axis is flagged.

```excel
=LAMBDA(
    LET(
        n_c,     COLUMNS(Source_Data),
        s_pos,   XMATCH(TRUE, TAKE(Spec_Sequence, n_c), 0),
        t_raw,   INDEX(Source_Data, 0, s_pos),
        t_v,     IF(t_raw = "", "", t_raw),
        rl,      TAKE(Spec_Role, n_c),
        id_mask, (rl = "Identifier (Row Label)") * (SEQUENCE(n_c) <> s_pos),
        g_v,     IF(SUM(id_mask) = 0,
            EXPAND("(all)", ROWS(t_v), 1, "(all)"),
            BYROW(CHOOSECOLS(Source_Data, TOROW(FILTER(SEQUENCE(n_c), id_mask))),
                LAMBDA(r, TEXTJOIN(CHAR(31), FALSE, IF(r = "", "", r))))),
        ok,      ISNUMBER(t_v) * (SUBSTITUTE(g_v, CHAR(31), "") <> ""),
        g_f,     FILTER(g_v, ok, NA()),
        t_f,     FILTER(t_v, ok, NA()),
        g_s,     SORTBY(g_f, g_f, 1, t_f, 1),
        t_s,     SORTBY(t_f, g_f, 1, t_f, 1),
        d,       DROP(t_s, 1) - DROP(t_s, -1),
        same,    DROP(g_s, 1) = DROP(g_s, -1),
        IFERROR(FILTER(d, same * (d > 0), NA()), NA())
    )
)
```

## `Sequence_Status` *(sheet-scoped: Regression)*

**Checks that at most one variable is flagged as the Sequence (time) axis. Blank when the spec is legal.**

Sequence_Status is the H2 readout. Zero or one Sequence flag is the legal range: zero is an ordinary non-panel spec, one designates the ordering axis for the lag, difference and serial-correlation features, and two or more is a spec error because there is no defined answer to which axis a lag is taken along.

The flagged cells in the Sequence column turn red at the same time, so this line says what is wrong and the column says where.

SHEET-SCOPED (scope: Regression): Spec_Sequence and Source_Data are unqualified and resolve against the calling sheet.

Returns: An ERROR string when more than one Sequence row is flagged, otherwise "".

H2 Sequence-cardinality verdict. ERROR when two or more Sequence rows are flagged; "" for zero or one. Sheet-scoped.

```excel
=LAMBDA(
    IF(SUMPRODUCT(N(TAKE(Spec_Sequence,COLUMNS(Source_Data))=TRUE))>1,"ERROR: multiple Sequence rows — mark at most one.","")
)
```

## `Serial_Correlation_Group` *(sheet-scoped: Regression)*

**Which column groups the residuals for the serial-correlation checks — the Fixed Effects variable today, with a reserved slot for the future Cluster role.**

Serial_Correlation_Group is the grouping-key resolver for the serial-correlation diagnostics: it answers "which dimension partitions the residuals for within-group differencing?" as a single dispatch point, so consumers (the BFN panel Durbin-Watson wiring) never reference a role column directly. Today the answer is the Fixed Effects column; when the Cluster role supplies pooled-panel grouping without absorption, ONLY this resolver changes — every consumer retargets for free.

The SWITCH dispatches on the declared grouping role, highest-priority first: a Role="Fixed Effects" row wins and returns Fixed_Effects_Column() (one source of truth for the FE lookup — the resolver dispatches, the accessor looks up); a Role="Cluster" row with no FE returns the not-applicable token "n/a — Cluster grouping RESERVED" — the branch is present in the SWITCH but deliberately inert, the switch that currently does nothing, mirroring the reserved-spec-column pattern (Order/Transform columns named now, read by nothing); with neither role the default returns the none-sentinel "none", signalling the ordinary single-series DW path. The Cluster branch is dormant AND unreachable from the shipped cells: "Cluster" is not in the Role dropdown, and the diagnostic cells gate on the FE-variable count before this resolver is ever evaluated — a hand-typed Cluster role reaches the token path, never an error.

Both non-column results are explicit text tokens, not NA() (reserved for genuine errors) — friendlier to a future Model-Comparison XLOOKUP and self-explaining in the audit trail. Spec_Role is TAKE-trimmed to COLUMNS(Source_Data) so cells below the live spec rows are never scanned.

Returns: The grouping-key column for serial-correlation diagnostics: the Fixed Effects data column when a Role="Fixed Effects" row is declared; the token "n/a — Cluster grouping RESERVED" when only a Cluster role is declared; the sentinel "none" when no grouping role exists (ordinary DW path).

Grouping-key resolver for serial-correlation diagnostics: the FE group column when Role="Fixed Effects" is declared; "n/a — Cluster grouping RESERVED" on a declared Cluster role (dormant branch); "none" otherwise (ordinary DW path).

```excel
=LAMBDA(LET(
    roles, TAKE(Spec_Role, COLUMNS(Source_Data)),
    key, IFS(
        SUMPRODUCT(N(roles = "Fixed Effects")) > 0, "Fixed Effects",
        SUMPRODUCT(N(roles = "Cluster")) > 0, "Cluster",
        TRUE, "none"
    ),
    SWITCH(key,
        "Fixed Effects", Fixed_Effects_Column(),
        "Cluster", "n/a — Cluster grouping RESERVED",
        "none"
    )
))
```

## `Skewness`

**Whether a variable has a longer tail on one side than the other.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Skewness returns the sample skewness (Fisher-Pearson standardized third moment) for each column of Predictors after filtering. Uses Excel's SKEW function, which divides by (n-1)(n-2).

Interpretation: 0 = symmetric; positive = right tail (mean > median, common in counts and incomes); negative = left tail. |Skewness| > 1 typically signals that a log or square-root transformation would linearize relationships and stabilize regression residual variance. Compare with Kurtosis for a fuller picture of distributional shape.

Returns: k×1 column vector of sample skewness values (one per column of Predictors)

Sample skewness per Predictors column after filtering. 0 = symmetric; > 1 = long right tail; < −1 = long left tail.

```excel
=LAMBDA(Predictors, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    xf, FILTER(Predictors, filt_arg),
    MAKEARRAY(COLUMNS(Predictors), 1, LAMBDA(r, c, SKEW(CHOOSECOLS(xf, r))))
  ))
```

## `Smearing_Factor`

**Duan's smearing factor — 1 when response is untransformed, mean(EXP(residuals)) under Log.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range (the model's fit-space response: transformed under Log, within-demeaned under FE)
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df and a None response transform

Smearing_Factor returns the Duan (1983) smearing factor: the mean of EXP(residuals) over the included sample, used to correct the bias of EXP(ŷ) as a predictor of the conditional mean of y when the response is log-transformed. The simple textbook back-transform EXP(ŷ) is an unbiased predictor of the conditional median; smearing lifts it to an unbiased predictor of the conditional mean.

With Context_Response_Transform = "None" the response is fit in its original units and no smearing is required, so the function returns 1.0 — never NA() — so the sheet's cells stay uniform across the None / Log boundary (a model with smearing=1 reproduces EXP-only exactly, the same way the v2.2 Log wiring reduces to the un-transformed behavior). With "Log" it returns AVERAGE(EXP(Residuals(X, Y, Include))). Any other response-transform value returns #N/A.

The factor is computed on the fit-time pair (X/Y) and is therefore a function of the same residuals the rest of the workbook reads. Under Fixed Effects, Residuals() already returns within residuals (the design is within-demeaned), so the smearing factor uses the within residuals directly; the FE level shift is reintroduced at the back-transformation step, not here.

Returns: Scalar smearing factor. 1.0 when Context_Response_Transform is "None"; AVERAGE(EXP(residuals)) when "Log"; #N/A for any other response-transform value.

Duan (1983) smearing factor: 1.0 when Context_Response_Transform = "None"; AVERAGE(EXP(Residuals(X,Y,Include))) when "Log"; #N/A otherwise. Under FE reads within residuals; the level shift is added back at the back-transformation step.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    filt_arg,    IF(ISOMITTED(Include), TRUE, Include),
    rt,          Context_Response_Transform(context_arg),
    IF(rt="None", 1,
      IF(rt="Log",
        IFERROR(AVERAGE(EXP(Residuals(X, Y, filt_arg))), NA()),
        NA()))
  )
)
```

## `Spearman_R`

**The rank-order relationship between each predictor and the outcome.**

Arguments:

- **xf** — predictor column range (single or multi-column); returns one correlation per column
- **yf** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Spearman_R computes the rank (Spearman) correlation between each column of xf and yf. For each predictor column, rows are kept where both that column and yf are numeric (and pass the optional Include). Ties receive average ranks via a MAP-based helper (avoids RANK.AVG on dynamic arrays). Returns a k×1 vector when xf has k columns.

Unlike Pearson_R, Spearman_R measures any monotonic relationship, not just linear ones. The diagnostic value of the pair is in their difference: if |Spearman_R| is noticeably larger than |Pearson_R| for the same predictor, the relationship is curved and a transformation (e.g. LOG) or a nonlinear term may improve model fit.

Returns: k×1 column vector of Spearman rank correlations between each predictor and Y

Rank (Spearman) correlations between each X column and Y. Detects monotonic but curved relationships. Returns a k×1 vector.

```excel
=LAMBDA(xf, yf, [Include],
  LET(
    baseFilter, IF(ISOMITTED(Include), TRUE, Include),
    RankAvgAsc, LAMBDA(a,
      MAP(a, LAMBDA(v, 1+SUM(--(a<v))+(SUM(--(a=v))-1)/2))
    ),
    TRANSPOSE(
      BYCOL(xf,
        LAMBDA(xcol,
          LET(
            valid, baseFilter*ISNUMBER(xcol)*ISNUMBER(yf),
            x, FILTER(xcol, valid),
            y, FILTER(yf, valid),
            rx, RankAvgAsc(x),
            ry, RankAvgAsc(y),
            CORREL(rx, ry)
          )
        )
      )
    )
  ))
```

## `SS_Regression`

**The part of Y's variation the model successfully explains.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

SS_Regression (also SSR, the model sum of squares) is the variation explained by the fitted model.

Computed as SS_Total - SS_Residual, which makes the ANOVA decomposition SS_Total = SS_Regression + SS_Residual true by construction.

Returns: SS_regression (explained/model sum of squares) as a scalar

Explained variation in Y = SS_Total × R². Grows as model fit improves.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg, IF(ISOMITTED(Include),       TRUE, Include),
    SS_Total(X, Y, filt_arg, context_arg) - SS_Residual(X, Y, filt_arg)
  ))
```

## `SS_Residual`

**The part of Y's variation the model still misses.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

SS_Residual (also SSE, the error sum of squares) is Σ(yᵢ - ŷᵢ)² — the variation the model fails to explain.

Read directly from LINEST's fifth output row, which holds the residual sum of squares regardless of the const setting. Taking it from LINEST rather than deriving it as SS_Total × (1 - R²) is what keeps the sums-of-squares chain acyclic: R_Squared is now derived FROM this value, not the other way round.

Returns: SS_residual (residual/error sum of squares) as a scalar

Residual (error) sum of squares Σ(y - ŷ)², read from LINEST's fifth output row. The base quantity from which R², SS_Regression and the mean squares are derived.

```excel
=LAMBDA(X, Y, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    INDEX(LINEST(FILTER(Y, filt_arg), FILTER(X, filt_arg), FALSE, TRUE), 5, 2)
  ))
```

## `SS_Total`

**All of the variation in Y before the model explains any of it.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

SS_Total is the total sum of squares — the residual sum of squares from the intercept-only model, that is, the projection of y off whatever the intercept column actually is:

    SS_Total = ||y||² - (c'y)² / (c'c)

One formula covers all three cases. With c a column of ones it reduces to DEVSQ(y) = Σy² - (Σy)²/n; with no intercept column it is SUMSQ(y); and under weighted least squares, where c is √w, it gives Σw(y - ȳ_w)².

The projection form is used rather than DEVSQ specifically because DEVSQ(√w ⊙ y) is NOT the weighted total sum of squares — it centers on mean(√w·y) instead of ȳ_w, which would leave SS_Total and therefore R² silently wrong under WLS. The decomposition SS_Total = SS_Regression + SS_Residual holds in every case.

Returns: SS_total (total sum of squares) as a scalar

Total sum of squares: the residual sum of squares from the intercept-only model, ||y||² - (c'y)²/(c'c) where c is the intercept column. Reduces to DEVSQ(y) for an ordinary intercept and SUMSQ(y) with none.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg, IF(ISOMITTED(Include),       TRUE, Include),
    y_filt,   FILTER(Y, filt_arg),
    IF(NOT(has_arg),
      SUMSQ(y_filt),
      LET(
        c, FILTER(CHOOSECOLS(X, 1), filt_arg),
        SUMSQ(y_filt) - SUMPRODUCT(c, y_filt) ^ 2 / SUMSQ(c)
      )
    )
  ))
```

## `Studentized_Residuals`

**Residuals adjusted for both model error and leverage.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

r_i = e_i / (s√(1−h_i)), the leverage-adjusted scaled residual. Unlike Scaled_Residuals (which divides by s only), studentized residuals have equal variance under the normal model and are more sensitive to outliers in high-leverage regions. Values beyond ±2 or ±3 indicate potential outliers. The denominator is 0 when the fit is exact for a row — a perfect-fit model (s = 0) or a leverage-1 row (h_i = 1, e_i = 0 for that row) — and that row returns #N/A rather than #DIV/0!. Cooks_Distance and the Scale-Location column both consume this output, so the #N/A propagates to them automatically.

An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads to SE_Regression() so s reflects the correct absorbed-group df.

Returns: n-element vector of internally studentized residuals

Internally studentized residuals rᵢ = eᵢ / (s√(1−hᵢ)). Values beyond ±2 flag outliers; ±3 are probable outliers. #N/A when s = 0 or hᵢ = 1. Optional DF_Absorbed (default 0) corrects s under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    e,  Residuals(X, Y, filt_arg),
    se, SE_Regression(X, Y, filt_arg, context_arg),
    h,  Hat_Diagonal(X, filt_arg),
    IFERROR(e / (se * SQRT(1 - h)), NA())
  ))
```

## `Studentized_Residuals_Ranked`

**Studentized residuals sorted from smallest to largest.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Ascending sort of Studentized_Residuals. Paired with Normal_Scores gives a leverage-corrected Q-Q plot; paired with Percentile gives a cumulative distribution plot.

Returns: sorted studentized residuals as spilled vector

Studentized residuals sorted ascending. Paired with Normal_Scores for a leverage-corrected Q-Q residual plot. Optional DF_Absorbed (default 0) under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    SORT(Studentized_Residuals(X, Y, filt_arg, context_arg))
  ))
```

## `Sturges_Bins`

**How many histogram bins Sturges' rule recommends for this data.**

Arguments:

- **data** — single-column numeric data range
- **filter** — optional boolean array

Sturges_Bins returns the bin count k = ⌈log₂(n) + 1⌉, where n is the count of numeric values after filtering. Sturges' rule works well for roughly symmetric, unimodal data with moderate sample sizes (50–200 observations). It tends to undersmooth (too few bins) for skewed or heavy-tailed data.

Used by Bin_Edges to generate upper bin edges for the Sturges histogram column chart.

Returns: integer bin count via Sturges' rule

Bin count k = ⌈log₂(n)+1⌉ (Sturges’ rule). Reliable for roughly symmetric data with 50–200 observations.

```excel
=LAMBDA(data, [filter],
  LET(
    filt_arg, IF(ISOMITTED(filter), ISNUMBER(data), filter),
    x, FILTER(data, filt_arg),
    n, COUNT(x),
    CEILING(LOG(n, 2) + 1, 1)
  ))
```

## `T_Statistics`

**Each coefficient measured in standard-error units.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

Returns a vertical spill array of t-statistics for each OLS coefficient. Each element equals the corresponding Coefficients value divided by its SE_Coefficients value.

Computed as Coefficients(...) / SE_Coefficients(...). An optional trailing DF_Absorbed (v2.1 Fixed Effects) threads only to SE_Coefficients() — the coefficient point estimates never depend on df.

Returns: vertical array of t-statistics for each coefficient (coefficient / standard error)

t-statistics = Coefficients / SE_Coefficients. Same order as Coefficients. Optional DF_Absorbed (default 0) threads to SE_Coefficients under Fixed Effects.

```excel
=LAMBDA(X, Y, [Include], [Context],
  LET(
    filt_arg,     IF(ISOMITTED(Include),     TRUE, Include),
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    absorbed_arg, Context_DF_Absorbed(context_arg),
    Coefficients(X, Y, filt_arg) / SE_Coefficients(X, Y, filt_arg, context_arg)
  ))
```

## `This_Row`

**Relative row numbers for the current spilled array.**

Arguments:

- **array** — any contiguous array (including one that has already been filtered)

Returns a 1-based integer sequence 1, 2, …, n where n = ROWS(array). The row numbers are relative to the top of the supplied array, so they are stable regardless of where the array sits on the sheet or whether it has been pre-filtered by the caller.

Designed to be passed to CHOOSEROWS or used as a comparison mask inside FILTER; see Exclude_Row_N for the canonical usage.

Returns: 1..n column vector of relative row numbers

Relative row numbers 1..n for the current spilled array. Pass to CHOOSEROWS or use as a mask inside FILTER.

```excel
=LAMBDA(array,
  SEQUENCE(ROWS(array))
)
```

## `Tolerance`

**How much unique information a predictor still has after overlap with the others.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Tolerance is the reciprocal of the Variance Inflation Factor, 1 / VIF — the share of a predictor's variance not explained by the other predictors. Values below roughly 0.1-0.2 indicate problematic multicollinearity.

Takes predictor columns, never a design matrix.

Returns: k×1 column vector of Tolerance values (= 1/VIF), one per predictor

1/VIF per predictor. Fraction of unique variance not shared with other predictors. < 0.2: review; < 0.1: strong collinearity.

```excel
=LAMBDA(Predictors, [Include],
  1 / VIF(Predictors, Include)
)
```

## `Total_Degrees_Of_Freedom`

**The total amount of independent information in the outcome column.**

Arguments:

- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array - [Has_Intercept, DF_Absorbed, Response_Transform, Predictor_Transform]; defaults to VSTACK(TRUE,0,"None","None") when omitted, so every ad-hoc caller sees an intercept model with no absorbed df

df_total is the total degrees of freedom for the sum of squares. When an intercept is fit (default), one df is spent estimating the grand mean, giving n - 1. When the model is forced through the origin (Has_Intercept = FALSE), no df is spent on the mean, giving n. df_total partitions into df_regression + df_residual.

Calls Observations() for n, then returns n - N(has_arg).

Returns: df_total (n-1 with intercept, n without) as a scalar integer

Total degrees of freedom = n − 1 with intercept or n without. Partitions into Regression_Degrees_Of_Freedom + Residual_Degrees_Of_Freedom.

```excel
=LAMBDA(Y, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    has_arg,      Context_Has_Intercept(context_arg),
    filt_arg, IF(ISOMITTED(Include),       TRUE, Include),
    n,        Observations(Y, filt_arg),
    n - N(has_arg)
  ))
```

## `Unit_Space_Adjusted_R_Squared`

**Adjusted R² in response units — same penalty form as the in-space version, df threads through Context_DF_Absorbed.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome in fit space: transformed under Log, within-demeaned under FE
- **Y_Full** — Response_Column() — the FIT-space response with the within-demean removed (ln(y) under a Log Response row), NOT the observed response in original units; the function back-transforms it internally.
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array; defaults to VSTACK(TRUE,0,"None","None") when omitted
- **Method** — Duan (default) or Naive; passed through to Unit_Space_R_Squared

Unit_Space_Adjusted_R_Squared penalises the unit-space R² for the model's parameter count, using the standard 1 − (1 − R²_unit) × df_total / df_residual form. df_total and df_residual are read through the same Total_Degrees_Of_Freedom / Residual_Degrees_Of_Freedom functions the centered R² uses, so the FE df correction (Context_DF_Absorbed) threads through automatically — under Fixed Effects the penalty term shrinks by the absorbed group count, exactly as it does for the in-space Adjusted R².

The reduction invariant: with Context_Response_Transform = "None" and no FE, Unit_Space_Adjusted_R_Squared = Adjusted_R_Squared to floating-point precision. The same six recognised (response, predictor) pairs govern when the function returns a number; outside them Unit_Space_R_Squared returns #N/A and this function's arithmetic propagates it. Can be negative when the unit-space model fits worse than the mean — same honest reading as the in-space Adjusted R², never clamped.

Returns: Unit-space Adjusted R² as a scalar. 1 − (1 − R²_unit) × df_total / df_residual, the standard penalty form. Reuses Total_Degrees_Of_Freedom and Residual_Degrees_Of_Freedom so Context_DF_Absorbed is honoured under FE.

1 − (1 − R²_unit) × df_total / df_residual, reusing the existing df functions so the FE correction is honoured. Reduces to Adjusted_R_Squared when no transforms and no FE.

```excel
=LAMBDA(X, Y, Y_Full, [Include], [Context], [Method],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    filt_arg,    IF(ISOMITTED(Include), TRUE, Include),
    1 - (1 - Unit_Space_R_Squared(X, Y, Y_Full, filt_arg, context_arg, IF(ISOMITTED(Method),"Duan",Method)))
      * Total_Degrees_Of_Freedom(Y, filt_arg, context_arg)
      / Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg)
  )
)
```

## `Unit_Space_Observed`

**The observed response, read in the same units as the back-transformed predictions.**

Arguments:

- **Y** — Design_Response() — the fit-space response the model was actually fitted on (transformed and, under Fixed Effects, within-demeaned).
- **Y_Full** — Response_Column() — the same response with the within-demean removed. Its difference from Y is the level the within transformation took out.
- **Include** — boolean row mask; when omitted every row is used
- **Context** — the Model_Context() array; defaults to the constructor's own defaults when omitted

Unit_Space_Observed is the observed half of every unit-space comparison. It exists because the observed response and the back-transformed prediction have to be read in the SAME space, and which space that is depends on the response transform — so the choice belongs in one function rather than being restated at each of the three goodness-of-fit call sites.

Both branches fall out of one expression. The level shift is IF(rt = "Log", FILTER(Y_Full, Include) - FILTER(Y, Include), 0) — the same shift Unit_Space_Predictions adds to its fitted values — and y_level is Dependent_Variable(Y, Include) + shift. Under Log that is Y_Full (the un-demeaned log response) and the Naive back-transform returns the raw response; under None the shift is zero, y_level is Y itself, and Back_Transform_Response is a pass-through, so the observed side stays the within-demeaned column the ordinary statistics are computed on.

The back-transform is forced to Naive. The smearing factor lifts a prediction from the conditional median to the conditional mean; an observation is neither, and smearing one would corrupt SSE and SST alike.

Returns: n x 1 column of the OBSERVED response, in the same space as Unit_Space_Predictions' output: the raw response under a Log Response row, the within-demeaned fit-space response under None.

Back_Transform_Response(Dependent_Variable(Y, Include) + shift, Context, "Naive", 1) with shift = IF(rt="Log", FILTER(Y_Full,Include)-FILTER(Y,Include), 0). Log -> raw y; None -> the within-demeaned Y, so the reduction invariant holds under FE.

```excel
=LAMBDA(Y, Y_Full, [Include], [Context],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    filt_arg,    IF(ISOMITTED(Include), TRUE, Include),
    rt,          Context_Response_Transform(context_arg),
    shift,       IF(rt = "Log", FILTER(Y_Full, filt_arg) - FILTER(Y, filt_arg), 0),
    y_level,     Dependent_Variable(Y, filt_arg) + shift,
    Back_Transform_Response(y_level, context_arg, "Naive", 1)
  )
)
```

## `Unit_Space_Predictions`

**Predictions in response units — adds back the level shift Y_Full − Y, then back-transforms. The (response, predictor) pair is a six-way SWITCH; outside the six, #N/A.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome in fit space: transformed under Log, within-demeaned under FE
- **Y_Full** — single-column outcome in response space with the within-demean removed: Response_Column() — transformed but NOT within-demeaned. Their difference on the filtered sample is exactly the level the within transformation removed, so the back-transformed fitted value carries the group effect without an FE-detection branch or new closure.
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array; defaults to VSTACK(TRUE,0,"None","None") when omitted
- **Method** — Duan (default — multiply by smearing) or Naive; passed through to Back_Transform_Response

Unit_Space_Predictions returns Predictions(X, Y, Include) plus a level shift, back-transformed to response units. The level shift is IF(rt = "Log", FILTER(Y_Full, Include) − FILTER(Y, Include), 0): when no Fixed Effects row is declared Y_Full and Y are the same column and the shift is exactly zero, and it is gated on a Log response because nothing is exponentiated under None — applying it there would silently convert the within-flavoured statistics into total ones and break the reduction invariant; under FE the difference is the group mean the within transformation removed, so the back-transformed fitted value carries the group effect (a unit-space statistic cannot read exp(within deviation) and call itself a prediction of the un-demeaned y).

The pair-SWITCH on (Context_Response_Transform, Context_Predictor_Transform) is the single dispatch point. The six recognised pairs are {None, Log} × {None, Log, Mixed} — the predictor half is INERT for the response-unit arithmetic (a response-unit statistic cannot depend on predictor units) but is carried as validation, so a spec with an unrecognised predictor transform (e.g. the v2.2-era "Center" placeholder) returns #N/A instead of silently producing a number. The dispatchers' actual arithmetic is a single Back_Transform_Response call per row of fitted values — keeps the Unit_Space_* names the names the v2.2 design committed to, and confines the (response, method) combinatorial switch to one place.

The function computes its own smearing factor — caller does not have to pre-compute Smearing_Factor and thread it. Under "None" the SWITCH reduces to a pass-through (Back_Transform_Response returns Values unchanged) and the smearing factor is 1, so the predictions equal Predictions(X, Y, Include) exactly. Under "Log" the smearing factor is the mean of EXP(within residuals) and the back-transform is EXP(fitted) * smearing (Duan) or EXP(fitted) (Naive).

Returns: n × 1 column of back-transformed predictions, in response units. #N/A when the (Context_Response_Transform, Context_Predictor_Transform) pair is outside the six recognised pairs {None, Log} × {None, Log, Mixed}.

Predictions(X,Y,Include) + shift, then Back_Transform_Response. shift = IF(rt="Log", FILTER(Y_Full,Include)-FILTER(Y,Include), 0) — the FE group mean, gated on Log because nothing is exponentiated under None. Six recognised pairs; else #N/A.

```excel
=LAMBDA(X, Y, Y_Full, [Include], [Context], [Method],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    method_arg,  IF(ISOMITTED(Method), "Duan", Method),
    filt_arg,    IF(ISOMITTED(Include), TRUE, Include),
    rt,          Context_Response_Transform(context_arg),
    pt,          Context_Predictor_Transform(context_arg),
    sm,          Smearing_Factor(X, Y, filt_arg, context_arg),
    shift,       IF(rt = "Log", FILTER(Y_Full, filt_arg) - FILTER(Y, filt_arg), 0),
    fitted,      Predictions(X, Y, filt_arg) + shift,
    SWITCH(rt & "|" & pt,
      "None|None",  Back_Transform_Response(fitted, context_arg, method_arg, sm),
      "None|Log",   Back_Transform_Response(fitted, context_arg, method_arg, sm),
      "None|Mixed", Back_Transform_Response(fitted, context_arg, method_arg, sm),
      "Log|None",   Back_Transform_Response(fitted, context_arg, method_arg, sm),
      "Log|Log",    Back_Transform_Response(fitted, context_arg, method_arg, sm),
      "Log|Mixed",  Back_Transform_Response(fitted, context_arg, method_arg, sm),
      NA())
  )
)
```

## `Unit_Space_R_Squared`

**R² in response units — 1 − SUMSQ(y − ŷ_unit) / SST_unit. May be negative; not clamped.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome in fit space: transformed under Log, within-demeaned under FE
- **Y_Full** — Response_Column() — the FIT-space response with the within-demean removed (ln(y) under a Log Response row), NOT the observed response in original units; the function back-transforms it internally.
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array; defaults to VSTACK(TRUE,0,"None","None") when omitted
- **Method** — Duan (default) or Naive; the unit-space R² always corresponds to the displayed predictions, so this is the same method the Prediction Outputs block uses

Unit_Space_R_Squared is the coefficient of determination in the response's original units: 1 − SSE_unit / SST_unit, where SSE_unit is the squared back-transformed residual and SST_unit is the response's total sum of squares. SST_unit is taken about the mean when the model has an intercept (the standard centered total) and about zero otherwise — the same convention SS_Total uses for the within-transform pair. The R² may be negative when the model fits worse than a horizontal line in the response's original units; that is an honest reading of a bad fit and is NOT clamped to zero.

Under FE + Log the unit-space statistics are TOTAL (about the grand mean of raw y), not within. exp(within deviation) predicts nothing, so the back-transformed fitted value necessarily carries the group effect — and so does SST_unit. Every other statistic on the sheet reports the within flavor; this one cannot, so the cell note for the unit-space block states the convention. With Context_Response_Transform = "None" and no FE row, Unit_Space_R_Squared reduces to R_Squared exactly — the reduction invariant the regression tests assert.

The Method argument is propagated to Unit_Space_Predictions so unit-space R² always corresponds to the predictions the sheet shows; the sheet's Back_Transform_Method toggle drives both via a single parameter.

The observed side is Back_Transform_Response(Dependent_Variable(Y_Full, Include), Context, "Naive", 1), not Y_Full itself: Y_Full is Response_Column(), which is ln(y) under a Log Response row, and measuring a back-transformed prediction against ln(y) would put two unit systems in one subtraction. The Naive branch is forced there because the smearing factor lifts a prediction from the conditional median to the conditional mean and has no business touching an observed value.

Returns: Unit-space R² as a scalar. 1 − SSE_unit / SST_unit (about the mean when the model has an intercept, about zero without — mirroring SS_Total). May be negative when the model fits worse than the mean; not clamped.

1 - SSE_unit/SST_unit, both formed from y_unit = the Naive back-transform of Dependent_Variable(Y_Full,Include) — original units, never fit-space Y_Full. SST_unit follows SS_Total. TOTAL under FE+Log; within (= R_Squared) under None.

```excel
=LAMBDA(X, Y, Y_Full, [Include], [Context], [Method],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    filt_arg,    IF(ISOMITTED(Include), TRUE, Include),
    method_arg,  IF(ISOMITTED(Method), "Duan", Method),
    y_unit,      Unit_Space_Observed(Y, Y_Full, filt_arg, context_arg),
    pred_unit,   Unit_Space_Predictions(X, Y, Y_Full, filt_arg, context_arg, method_arg),
    sse_unit,    SUMSQ(y_unit - pred_unit),
    sst_unit,    LET(has, Context_Has_Intercept(context_arg),
                     IF(has,
                       SUMSQ(y_unit - AVERAGE(y_unit)),
                       SUMSQ(y_unit))),
    1 - sse_unit / sst_unit
  )
)
```

## `Unit_Space_Residuals`

**Residuals in response units — observed y minus back-transformed fitted value.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome in fit space: transformed under Log, within-demeaned under FE
- **Y_Full** — Response_Column() — the FIT-space response with the within-demean removed (ln(y) under a Log Response row), NOT the observed response in original units; the function back-transforms it internally.
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array; defaults to VSTACK(TRUE,0,"None","None") when omitted
- **Method** — Duan (default) or Naive; passed through to Unit_Space_Predictions

Unit_Space_Residuals returns the observed response in original units minus the back-transformed fitted value: Back_Transform_Response(Dependent_Variable(Y_Full, Include), Context, "Naive", 1) − Unit_Space_Predictions(X, Y, Y_Full, Include, Context, Method). Y_Full is the FIT-space un-demeaned response, so the observed side goes back through the same transform as the predictions before the subtraction — always on the Naive branch, because an observation is not a prediction and must never carry the smearing factor. Subtracting Y_Full directly would put ln(y) and a response-unit prediction in one expression. The same (response, predictor) SWITCH governs the result; outside the six recognised pairs both this and Unit_Space_Predictions return #N/A, never a silently-wrong number.

When Context_Response_Transform = "None" and no Fixed Effects row is declared, Unit_Space_Residuals reduces to Residuals(X, Y, Include) exactly — Y_Full and Y are the same column, the level shift is zero, and Back_Transform_Response is a pass-through. The reduction invariant this preserves is the acceptance criterion for the whole back-transformation family: with no transforms and no FE, the new columns equal the originals to floating-point precision, so a model that was already shipped continues to compute the same numbers under v3.3.

Returns: n × 1 column of unit-space residuals (observed response in original units minus back-transformed fitted value). #N/A outside the six recognised (response, predictor) pairs.

y_unit - Unit_Space_Predictions(...), y_unit = Back_Transform_Response(Dependent_Variable(Y_Full,Include), Context, "Naive", 1). The observed side is never smeared. Reduces to Residuals(X,Y,Include) with no transforms and no FE.

```excel
=LAMBDA(X, Y, Y_Full, [Include], [Context], [Method],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    filt_arg,    IF(ISOMITTED(Include), TRUE, Include),
    method_arg,  IF(ISOMITTED(Method), "Duan", Method),
    y_unit,      Unit_Space_Observed(Y, Y_Full, filt_arg, context_arg),
    y_unit - Unit_Space_Predictions(X, Y, Y_Full, filt_arg, context_arg, method_arg)
  )
)
```

## `Unit_Space_RMSE`

**RMSE in response units — SUMSQ(unit residuals) / df_residual, df corrected for FE.**

Arguments:

- **X** — design matrix — the constructed model columns, including the intercept column when the model has one
- **Y** — single-column outcome in fit space: transformed under Log, within-demeaned under FE
- **Y_Full** — Response_Column() — the FIT-space response with the within-demean removed (ln(y) under a Log Response row), NOT the observed response in original units; the function back-transforms it internally.
- **Include** — boolean array — TRUE includes the row, FALSE excludes it
- **Context** — the materialized Model_Context() 4x1 array; defaults to VSTACK(TRUE,0,"None","None") when omitted
- **Method** — Duan (default) or Naive; passed through to Unit_Space_Predictions

Unit_Space_RMSE is the standard error of the regression in response units: SQRT(SSE_unit / Residual_Degrees_Of_Freedom(X, Y, Include, Context)). The divisor is read through Residual_Degrees_Of_Freedom, so FE df correction (Context_DF_Absorbed) threads through — under Fixed Effects, df_residual is the within residual df and the RMSE divisor matches what SE_Regression would have produced on the same model. The Method argument is propagated to Unit_Space_Predictions so the RMSE corresponds to the predictions the sheet shows.

The reduction invariant: with Context_Response_Transform = "None" and no FE, Unit_Space_RMSE = SE_Regression exactly. Outside the six recognised (response, predictor) pairs Unit_Space_Predictions returns #N/A, which propagates through the SUMSQ and the SQRT to a final #N/A — never a silently-wrong number.

The observed side is Back_Transform_Response(Dependent_Variable(Y_Full, Include), Context, "Naive", 1), not Y_Full itself: Y_Full is Response_Column(), which is ln(y) under a Log Response row, and measuring a back-transformed prediction against ln(y) would put two unit systems in one subtraction. The Naive branch is forced there because the smearing factor lifts a prediction from the conditional median to the conditional mean and has no business touching an observed value.

Returns: Unit-space RMSE as a scalar. SQRT(SSE_unit / df_residual) — the same divisor SE_Regression uses, so the None / no-FE case reduces to SE_Regression exactly.

SQRT(SUMSQ(y_unit - pred_unit) / Residual_Degrees_Of_Freedom(...)), y_unit = the Naive back-transform of Dependent_Variable(Y_Full,Include). Same divisor as SE_Regression, so the None case reduces to it exactly.

```excel
=LAMBDA(X, Y, Y_Full, [Include], [Context], [Method],
  LET(
    context_arg, IF(ISOMITTED(Context), Model_Context(), Context),
    filt_arg,    IF(ISOMITTED(Include), TRUE, Include),
    method_arg,  IF(ISOMITTED(Method), "Duan", Method),
    y_unit,      Unit_Space_Observed(Y, Y_Full, filt_arg, context_arg),
    pred_unit,   Unit_Space_Predictions(X, Y, Y_Full, filt_arg, context_arg, method_arg),
    SQRT(SUMSQ(y_unit - pred_unit) / Residual_Degrees_Of_Freedom(X, Y, filt_arg, context_arg))
  )
)
```

## `Upper_Bin_Edges`

**The k upper boundary values of the histogram bins.**

Arguments:

- **data** — single-column numeric data range
- **method** — bin-count rule: "Sturges", "Scott", or "FD" (default when omitted)
- **filter** — optional boolean array

Upper_Bin_Edges returns the k upper bin edges derived from Bin_Edges by dropping the first boundary (the data minimum). Each edge is the right endpoint of the corresponding bin. Bins are half-open intervals (lower, upper], matching Excel's FREQUENCY convention.

Returns: k × 1 column vector of upper bin edges (the k+1 boundary vector with the first element dropped)

k upper bin boundaries (Bin_Edges without the first element). Compatible with Excel’s FREQUENCY function convention.

```excel
=LAMBDA(data, [method], [filter],
  DROP(Bin_Edges(data, method, filter), 1)
)
```

## `VIF`

**How much overlap with other predictors is inflating each coefficient's instability.**

Arguments:

- **Predictors** — predictor columns before the model-fitting stages — never carries an intercept column
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Variance Inflation Factor for each predictor: 1 / (1 - R²ⱼ), where R²ⱼ comes from regressing predictor j on all the others. Values above roughly 5-10 indicate problematic multicollinearity.

Takes PREDICTOR columns, never a design matrix — a constant intercept column among the predictors would make the auxiliary regressions singular. The auxiliary fits themselves do need an intercept, so VIF builds one onto each sub-matrix itself.

Returns: k×1 column vector of VIF values, one per predictor (does not include an intercept row)

Variance Inflation Factor per predictor = 1/(1−R²ⱼ). VIF = 1: no collinearity. VIF > 5: review. VIF > 10: problem.

```excel
=LAMBDA(Predictors, [Include],
  LET(
    filt_arg, IF(ISOMITTED(Include), TRUE, Include),
    k,        COLUMNS(Predictors),
    IF(k = 1,
      1,
      LET(
        col_idx, SEQUENCE(1, k),
        ones,    SEQUENCE(ROWS(Predictors), 1, 1, 0),
        R2_vec,  MAKEARRAY(k, 1, LAMBDA(j, _,
          R_Squared(
            HSTACK(ones, CHOOSECOLS(Predictors, FILTER(col_idx, col_idx <> j))),
            CHOOSECOLS(Predictors, j),
            TRUE,
            filt_arg
          )
        )),
        1 / (1 - R2_vec)
      )
    )
  ))
```

## `Y_Ranked`

**The filtered outcome values sorted from low to high.**

Arguments:

- **Y** — single-column outcome range
- **Include** — boolean array — TRUE includes the row, FALSE excludes it

Useful for empirical distribution/probability plots of the response variable.

Returns: sorted filtered Y values as a spilled vector

Filtered Y values sorted from low to high. Paired with Normal_Scores to build a Y-variable Q-Q plot.

```excel
=LAMBDA(Y, [Include],
  SORT(FILTER(Y, IF(ISOMITTED(Include), TRUE, Include)))
)
```
