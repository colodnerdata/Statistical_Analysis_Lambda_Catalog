# Regression Sheet — Fixed Effects Human Test Plan (WHO Life Expectancy data)

Execute in order — each test is a small delta on the previous spec state.
Every test gives an **Inputs** block (exact cell edits on the `Regression`
sheet's spec block, everything else unchanged from the previous test) and an
**Expected** block (exact cells and the values they must show). Reference
numbers were computed independently in Python — an explicit LSDV
(dummy-per-group) `statsmodels` fit for every FE-active number, a plain
`statsmodels` OLS fit for the pooled baseline — the same cross-checks
`tests/test_within_estimator.py`, `tests/test_df_absorbed_threading.py`, and
`tests/test_group_prediction_interval.py` encode as pytest assertions.
Mismatches against this plan indicate implementation bugs, not stale
expectations.

**Dataset.** This plan targets the WHO Life Expectancy panel, not the
shipped default (Auto MPG has no natural panel-unit variable). Build or
retarget with:

```
python build_production.py --regression-dataset life_expectancy --verify --no-launch --skip-data-table-calculations --skip-univariate
```

**Baseline facts:** 2,938 data rows; 2,482 rows complete on Life expectancy +
GDP + Schooling. 157 countries. Every reference number below uses **only**
GDP and Schooling as predictors and Country as the panel unit — Year is
Sequence-flagged but not itself a predictor, and every other numeric column
is Omit, so the model is directly comparable to an LSDV fit with exactly
those two slopes plus country dummies.

## Cell map (Regression sheet, spec block + relevant output cells)

| Cell(s) | Contents |
|---|---|
| B4 (Country row, the spec block's first data row) | Role dropdown — this plan flips it to `Fixed Effects` in T1 |
| B1 | Fixed Effects cardinality error (blank when legal, red at 2+ rows) |
| J1 / K1 / L1 (headers), J2 / K2 / L2 (values) | FE Variable / FE Groups / FE df absorbed status block |
| C2 | Allow_Intercept toggle — red when TRUE while FE is active |
| Y4:Y8 | Multiple R, R Square, Adjusted R Square, Standard Error, Observations |
| AA7:AA12 / AB7:AB12 | AIC, BIC, AICc, QQ Correlation, Durbin-Watson, BFN Panel Durbin-Watson |
| X21:AE21+ | Coefficients block (Intercept, GDP, Schooling rows) |
| AG3:AH14 | Prediction Interval box (point/CI/PI rows 3-11), FE Group selector (12), Group Mean (y) / Group Count (13-14) |
| AG19:AH20 | Prediction Inputs — GDP then Schooling (spec table order) |

---

## T0 — Pooled baseline (no Fixed Effects declared)

**Inputs:**

| Variable | Role | Include | Type | Sequence |
|---|---|---|---|---|
| Country | Identifier | FALSE | Continuous | |
| Year | Omit | FALSE | Continuous | TRUE |
| Status | Omit | FALSE | Categorical | |
| Life expectancy | Response | FALSE | Continuous | |
| Adult Mortality | Omit | FALSE | Continuous | |
| GDP | Predictor | TRUE | Continuous | |
| Schooling | Predictor | TRUE | Continuous | |
| (all other numerics, incl. Full_Data) | Omit | FALSE | Continuous | |

Allow_Intercept (C2) = TRUE.

**Expected** (plain OLS on GDP + Schooling, N=2,482, df_resid=2,479):

| Cell | Expect |
|---|---|
| B1 | blank (zero Fixed Effects rows is legal) |
| J2 / K2 / L2 | `n/a` / `n/a` / `n/a` (no Fixed Effects row) |
| Y21 (Intercept) | 44.976161 |
| Coefficient GDP / Schooling | 0.000103 / 1.955283 |
| Y5 (R Square) | 0.588954 |
| Y7 (Standard Error) | 6.181600 |
| AA/AB7 (AIC) | 9045.31 |
| AA/AB8 (BIC) | 9062.76 |
| AA/AB11 (Durbin-Watson) | a number (Year is Sequence-flagged) |
| AA/AB12 (BFN Panel DW) | `n/a — no fixed effects` |
| AH3 (Prediction Point Estimate, inputs left at Training Mean) | ≈ the response's overall mean (pooled model, no group adjustment) |

---

## T1 — Declare Country as Fixed Effects

**Inputs:** change B4 (Country) from `Identifier` to `Fixed Effects`. Nothing
else changes from T0.

**Expected** (one-way within estimator, G=157 groups, absorbed df=156,
true residual df = 2,479 − 156 = 2,323 — matches an explicit LSDV fit's
`df_resid` exactly):

| Cell | Expect |
|---|---|
| B1 | still blank (exactly one Fixed Effects row is legal) |
| J2 (FE Variable) | `Country` |
| K2 (FE Groups) | 157 |
| L2 (FE df absorbed) | 156 |
| Y21 (Intercept) | ≈ 0 (order 1e-15 — a genuine floating-point zero, not an approximation; see DECISIONS.md's within-estimator identity) |
| Coefficient GDP | 0.0000140 (1.398×10⁻⁵) |
| Coefficient Schooling | 0.852279 |
| SE (GDP) / SE (Schooling) | 5.245×10⁻⁶ / 0.039808 |
| t (GDP) / t (Schooling) | 2.6658 / 21.4097 |
| p (GDP) / p (Schooling) | 0.00773 / ≈5.86×10⁻⁹³ |
| 95% CI, GDP | [3.697×10⁻⁶, 2.427×10⁻⁵] |
| 95% CI, Schooling | [0.774216, 0.930342] |
| Y5 (R Square) | 0.170426 — the **within** R², not the 0.94 an LSDV's own R² would report (that number credits the 156 country dummies too; this cell deliberately does not) |
| Y6 (Adjusted R Square) | 0.114002 |
| Y7 (Standard Error) | 2.424211 |
| AA/AB7 (AIC) | 4549.33 |
| AA/AB8 (BIC) | 5474.21 |
| AA/AB11 (Durbin-Watson) | `n/a — FE active` |
| AA/AB12 (BFN Panel DW) | a number (Sequence + exactly one FE row) — the pinned value for this exact panel is `0.6362023311147436` per `tests/test_bfn_panel_durbin_watson_verification.py` |
| C2 (Allow_Intercept, still TRUE) | now shaded/flagged red (Phase 4's intercept × FE coupling CF) |

---

## T2 — FE Group prediction (Albania, GDP=15000, Schooling=12)

**Inputs:** AH12 (FE Group) = `Albania`. AH19 (GDP prediction input) = 15000.
AH20 (Schooling prediction input) = 12.

**Expected:**

| Cell | Expect |
|---|---|
| AH13 (Group Mean (y)) | 75.15625 |
| AH14 (Group Count) | 16 |
| AH3 (Point Estimate) | 75.219153 |
| AH4 (SE (Mean)) | 0.609878 |
| AH5 (SE (New Obs)) | 2.499750 |
| AH6 (t Critical) | 1.960986 |
| AH7 / AH8 (CI Lower/Upper) | 74.023191 / 76.415115 |
| AH9 / AH10 (PI Lower/Upper) | 70.317179 / 80.121127 |
| AH11 (Confidence Level) | 0.95 |

**Sanity check (DECISIONS.md):** set AH19/AH20 to Albania's own GDP/Schooling
means (2119.726679 / 12.1375) instead. AH3 (Point Estimate) must equal AH13
(Group Mean (y)) exactly — 75.15625 — and AH4 (SE (Mean)) must equal
`Y7 / SQRT(AH14)` = 2.424211 / √16 = 0.606053, since predicting at the
group's own centroid kills the quadratic term.

---

## T3 — Fixed Effects cardinality error

**Inputs:** additionally set Status's Role to `Fixed Effects` (a second FE
row, alongside Country).

**Expected:**

| Cell | Expect |
|---|---|
| B1 | `ERROR: multiple Fixed Effects rows (mark at most one variable)`, red fill |
| AA/AB12 (BFN Panel DW) | `n/a — multiple FE variables` |
| AA/AB11 (Durbin-Watson) | still `n/a — FE active` (fe_vars>0 gates this cell regardless of the count) |

Revert Status to `Omit` before continuing.

---

## T4 — Degenerate FE variable (single level)

**Inputs:** with Country back as the only Fixed Effects row, additionally
Filter the sample down to a single country (e.g. add a Filter column that is
TRUE only for Albania's 16 rows).

**Expected:**

| Cell | Expect |
|---|---|
| L2 (FE df absorbed) | 0 (one level in the included sample absorbs no degrees of freedom — `Dummy_Levels` degenerates to `#N/A`, `Absorbed_Degrees_Of_Freedom()` maps that to 0, not an error) |
| K2 (FE Groups) | 1 |
| Regression Statistics / ANOVA / Coefficients | compute as an ordinary 16-row pooled OLS on GDP+Schooling for Albania alone (df_resid = 16 − 2 − 1 = 13) — the same non-breaking-default collapse T0 exercises, now triggered by data rather than by an absent Role |

Remove the Filter before continuing (or leave the workbook as a scratch
copy — this test is destructive to the working sample).

---

## Appendix — where these numbers come from

Every reference number in T0–T2 is reproduced by an independent Python
computation, runnable standalone:

```
python tests/test_within_estimator.py            # coefficients (T0, T1)
python tests/test_df_absorbed_threading.py        # SE/t/p/CI/AIC/BIC (T1)
python tests/test_group_prediction_interval.py    # point/CI/PI (T2)
python tests/test_bfn_panel_durbin_watson_verification.py  # BFN pinned value (T1)
```

None of these scripts touch Excel — they load
`sample_data/Life Expectancy Data.csv` directly and cross-check against
`statsmodels`, so a mismatch against this plan should first be chased in the
Excel formulas, not in these reference numbers.
