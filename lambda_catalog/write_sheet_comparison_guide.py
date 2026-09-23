"""Write the Model Comparison Guide sheet into the target workbook.

This sheet is a fixed, dataset-independent reference, so artifact builds
just copy the already-styled sheet out of ``TEMPLATE_PATH`` (see
``copy_static_sheet``) instead of re-running the row-by-row COM writes on
every build. ``_write_template_sheet`` remains the authored source of the
content; run ``scripts/rebuild_static_sheets.py`` after editing it to
regenerate the template, then commit the updated ``.xlsx``.

Where the sibling sheets sit: ``Regression Instructions`` is the
operational *how* (point the sheet at data, which column does what),
``Modeling Concepts`` is the conceptual *why* (one row per feature), and
``Diagnostic Guide`` is the *what to look for* in one fitted model's
residual plots. The question all three leave open — "which of these
specifications should I report, on what evidence, and is this comparison
valid at all?" — is this sheet's layer.

Every address below is literal text, not an f-string, so
``tests/test_static_template_freshness.py`` covers the whole sheet with its
strong exact-literal tier; ``tests/test_static_sheet_citations.py`` then
pins each citation back to the ``regression_layout`` / ``spec_layout``
constant it names, so a column insertion cannot leave this sheet quietly
pointing at the wrong cell.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import xlwings as xw

from .sheet_styles import HEADER_COLOR as _HEADER_COLOR
from .sheet_styles import SUBHDR_COLOR as _SUBHEADER_COLOR
from .workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    ColumnSpec,
    copy_static_sheet,
    get_or_create_sheet,
    open_or_create_workbook,
    raise_excel_access_error,
    reset_generated_sheet,
    set_column_widths,
)

SHEET_NAME = "Model Comparison Guide"

_ROOT_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = _ROOT_DIR / "templates" / "static_sheets.xlsx"

_COLUMNS: tuple[ColumnSpec, ...] = (
    ColumnSpec(1, 32, "Condition / Statistic / Step"),
    ColumnSpec(2, 42, "Where it is"),
    ColumnSpec(3, 48, "What it tells you / What it moves"),
    ColumnSpec(4, 48, "What to do about it / What it leaves alone"),
)
_LAST_COL = _COLUMNS[-1].index


# ── 1. Pre-flight ─────────────────────────────────────────────────────────────
# The conditions under which two fitted specifications may be lined up at all.
# The fifth is the non-obvious one: the unit-space formulas read their own
# method cell (AH5), so the Duan / Naive toggle moves the one surface that is
# otherwise comparable across a transform change.
_PREFLIGHT: list[list[str]] = [
    [
        "1. The same response variable",
        "The Response row's Role in Col B, Model Specification, and the response "
        "readout at Cell AF3, Regression Outputs.",
        "Both specifications declare the same variable as the Response.",
        "A different response is a different question, not a worse model. Fit and "
        "report each response separately; there is no comparison to make and no "
        "statistic that fixes one.",
    ],
    [
        "2. A matched response transform",
        "Col G, Model Specification — the Transform token on the Response row.",
        "Both read None, or both read a Log token. The two Log tokens both build "
        "the identical Ln(x) column, so Log against Log (drop ≤ 0) is a matched "
        "transform — the difference between them is which rows survive, which is "
        "condition 3, not this one.",
        "None against a Log token puts the two fits on different response scales: "
        "the coefficients, the residuals and every fit-space statistic are in the "
        "units of different variables. The only surface that survives this is the "
        "unit-space triplet at Cell AH7:AH9, Regression Outputs.",
    ],
    [
        "3. The same sample",
        "Cell AB9, Regression Statistics — Observations.",
        "The two specifications report the same count, which means the same Filter "
        "rows are active (Col B, Role = Filter), the same rows are complete, and "
        "the same non-positive rows survive the Log token.",
        "A sample change voids every comparison, because n sits inside every "
        "information criterion's penalty, every sum of squares and every degrees-"
        "of-freedom count. Refit both on the smaller sample and compare there.",
    ],
    [
        "4. The same fixed-effects structure",
        "The Fixed Effects row's Role in Col B, Model Specification, and the "
        "readouts beside the Intercept label at Cell J2 — the first of the three "
        "Fixed Effects readouts in the band at J1:L2.",
        "Both absorb no fixed effects, or both absorb the same variable. One "
        "Fixed Effects row is the legal maximum — a second is a red spec error.",
        "A fixed-effects change rewrites what the residual columns mean: they "
        "become within-group deviations, and Durbin-Watson reads n/a while the BFN "
        "Panel Durbin-Watson at Cell AE13 takes over. Coefficients answer a "
        "different question, so a fit-space comparison across the change is void.",
    ],
    [
        "5. The same back-transform method",
        "Cell AH5, Regression Outputs — the Duan / Naive toggle in the Unit-Space "
        "Fit block.",
        "Both read Duan, or both read Naive.",
        "The unit-space formulas each read that cell, so the toggle changes the "
        "unit-space triplet itself: two models whose Cell AH7:AH9 values were "
        "produced under different settings are not comparable even though the "
        "cells are the comparable-by-construction surface. Set one policy and hold "
        "it for the whole comparison.",
    ],
]


# ── 2. Change ledger ──────────────────────────────────────────────────────────
_CHANGE_LEDGER: list[list[str]] = [
    [
        "Switch a predictor's Include toggle (Col C, Model Specification)",
        "R Squared at Cell AB6, Adjusted R Squared at Cell AB7 and Standard Error "
        "at Cell AB8, Regression Statistics; PRESS at Cell AE5, PRESS R Squared at "
        "Cell AE6 and AIC, BIC and AICc at Cells AE8-AE10, Diagnostics; the ANOVA "
        "sums of squares and Significance F; the Design Columns total at Cell O1.",
        "The sample (Cell AB9) and the response transform. The predictors that "
        "stayed in are fitted on the same rows in the same units.",
        "This is the comparison the sheet exists for: same sample, same transform, "
        "different design. Read Adjusted R² rather than R², and check the added "
        "term's own P-value in Col AE, Coefficients.",
    ],
    [
        "Change the Response row's Transform token (Col G)",
        "Every coefficient and its standard error, the residuals, Standard Error, "
        "PRESS, and AIC / BIC / AICc — all computed on the transformed response. "
        "Significance F moves too: it is a ratio of transformed-scale sums of "
        "squares, so its p-value is not transform-invariant.",
        "n, the design matrix's width, and the residual degrees of freedom — all "
        "properties of the sample and the design, not of the response's scale.",
        "Fit-space statistics from before and after the change cannot be lined up. "
        "Read the unit-space triplet at Cell AH7:AH9, which is computed in original "
        "response units precisely so this comparison is legitimate.",
    ],
    [
        "Switch a Filter row on or off (Col B, Role = Filter)",
        "n at Cell AB9, and through it every statistic on the sheet: the sums of "
        "squares, the residual degrees of freedom, the standard errors, and the "
        "penalty term inside AIC, BIC and AICc.",
        "Nothing. A filter changes the population, and every reported number is a "
        "statement about that population.",
        "The most expensive accidental edit in the workbook, because the fit for "
        "the surviving rows is unchanged and the sheet looks the way it did before. "
        "Check Cell AB9 before believing any two numbers are comparable.",
    ],
    [
        "Switch the Response token from Log to Log (drop ≤ 0)",
        "n at Cell AB9 — the zero and negative rows leave the sample — and every "
        "statistic that reads n. Log Domain Status above the spec reports the "
        "excluded-row count.",
        "The Ln(x) column itself: both tokens build the identical transformation. "
        "This is a sample change wearing a transform change's clothes.",
        "Treat it as condition 3, not condition 2. If one specification drops the "
        "non-positive rows and the other leaves them in, the second one's fit is "
        "dead (#N/A through every statistic) rather than merely incomparable.",
    ],
    [
        "Turn the intercept off (Cell C2, Model Specification)",
        "The design matrix's width at Cell O1, the residual degrees of freedom at "
        "Cell AB17, and the meaning of every coefficient and its standard error.",
        "The response, the sample, and the predictors that are still in.",
        "A regression through the origin is a different model class, not a "
        "reparametrization: R² and Adjusted R² are no longer on the same footing as "
        "an intercept model's. Compare it on the unit-space triplet and on its "
        "residual plots, and say in the write-up that the intercept was removed.",
    ],
    [
        "Change a categorical Reference Level (Col E, Model Specification)",
        "The coefficient values, their standard errors, their t statistics and "
        "P-values in Col AD and Col AE, Coefficients, and the contrasts each "
        "coefficient represents.",
        "R², Adjusted R², Standard Error, PRESS, AIC, BIC and AICc. The fitted "
        "response is identical; only the parametrization moved.",
        "Nothing about the fit improved or worsened, so a coefficient table is not "
        "evidence about a specification change. Reference-level invariance holds "
        "only because the intercept is kept: Dummy_Code drops one column per "
        "categorical, which spans the same space as the full block given an "
        "intercept, and the design-width audit at Cell O1 stays put.",
    ],
    [
        "Change the Sequence Period override (Col I, Model Specification)",
        "The Spacing Verdict at Cell I2 above the spec, which reports how the "
        "declared period compares with the spacing the data actually has; the BFN "
        "Panel Durbin-Watson at Cell AE13, whose within-group differencing uses the "
        "period as its lag; and any hand-written column formula calling Lag_By or "
        "Difference_By with the delta left out, which defaults to this period.",
        "Every fit statistic. The spec block constructs no column from the period, "
        "so the design matrix, the coefficients, the standard errors, the sample "
        "and every information criterion are identical under any period setting.",
        "A Sequence edit is a declaration about the panel's timing and a change to "
        "one diagnostic — never a model edit. Two period settings are not two "
        "specifications. Check the Spacing Verdict before treating a period edit as "
        "a design edit.",
    ],
    [
        "Change the Duan / Naive toggle (Cell AH5)",
        "The unit-space fit block: the Smearing Factor at Cell AH6 and the triplet "
        "at Cell AH7:AH9.",
        "Every fit-space statistic: R², Adjusted R², Standard Error, PRESS and the "
        "information criteria do not read the toggle.",
        "Because it moves the one across-transform surface, it is a pre-flight "
        "condition rather than a modelling choice. Two unit-space readings taken "
        "under different settings are as incomparable as a None response against a "
        "logged one.",
    ],
    [
        "Change Grid Points on the Univariate sheet",
        "Nothing on the Regression sheet. The two sheets fit different models for "
        "different questions, and no name on one reads a cell on the other.",
        "Every specification on the Regression sheet.",
        "Grid Points is a search-resolution control — it resizes a fitting grid, "
        "not a regression. Do not read a finer Univariate grid as a better "
        "regression.",
    ],
]


# ── 3. The analyst's questions ────────────────────────────────────────────────
_QUESTIONS: list[list[str]] = [
    [
        "Q1. Does the model beat the intercept-only baseline?",
        "Significance F at Cell AF16, ANOVA Table, against the Alpha input at Cell "
        "AB13, Regression Statistics.",
        "A p-value at or below alpha says the fitted model explains more than the "
        "response's own mean does. The F statistic itself is at Cell AE16.",
        "Not whether the model is good. With a large sample a trivial model clears "
        "the bar, and with a small one a useful model can miss it. Read it as a "
        "floor, not a verdict.",
    ],
    [
        "Q2. Is the fit tight?",
        "Multiple R at Cell AB5, R Squared at Cell AB6, Adjusted R Squared at Cell "
        "AB7 and Standard Error at Cell AB8, Regression Statistics.",
        "Adjusted R² is the one to read when the two specifications spend a "
        "different number of columns, because it charges for the columns spent. "
        "Standard Error is in the response's units, so it answers how far a "
        "prediction typically lands.",
        "A negative Adjusted R² means the residual variation does not earn back the "
        "parameters spent — not that the model fits worse than the response's own "
        "mean, which an intercept model cannot do. R² alone never falls when a "
        "predictor is added.",
    ],
    [
        "Q3. Do the individual terms carry anything?",
        "Each coefficient's t statistic in Col AD and P-value in Col AE, "
        "Coefficients, and the Beta Weights in Col AH for relative size on "
        "standardized predictors.",
        "A term whose P-value is above alpha is not distinguishable from zero at "
        "this sample size, holding the other terms in the model.",
        "Significance F does not answer this: it tests the whole model against the "
        "intercept-only model, never one term against another. A term can be "
        "individually insignificant in a model that is globally significant.",
    ],
    [
        "Q4. Can the fit be trusted?",
        "The Diagnostic Guide's Tier 1 and Tier 2 plots, the influence columns in "
        "Residual Output, and the threshold table on that sheet.",
        "See the sections below on stability and on which evidence survives a "
        "broken assumption — this sheet does not restate the Diagnostic Guide's "
        "per-plot reading rules.",
        "A specification chosen on fit statistics alone has not been screened. The "
        "winner of a comparison still has to pass the same charts as the losers.",
    ],
    [
        "Q5. Which of two specifications should be reported?",
        "The comparison ladder below, read in order, after the five pre-flight "
        "conditions hold.",
        "Prefer the unit-space evidence where the transform differs; prefer "
        "Adjusted R² and the information criteria where it does not; settle a tie "
        "on the cleaner diagnostics.",
        "There is no ranking, no weighting and no automatic selection in this "
        "workbook. The ladder says what evidence to look at, not what answer to "
        "reach.",
    ],
    [
        "Q6. Is one observation driving the result?",
        "Cook's Distance in Col AT, the PRESS Residual column in Col AX, Hat "
        "Diagonal in Col AR and Studentized Residuals in Col AS, Residual Output; "
        "the Cook's Distance chart labels the flagged rows.",
        "Cook's Distance is screened against F.INV(0.5, p, n-p), where p is the "
        "design matrix's width from Cell O1 (intercept included) and n-p the "
        "residual degrees of freedom at Cell AB17 — so the bar scales with the "
        "model's own size.",
        "Remove a row only for a non-statistical reason: a data-entry error, or a "
        "population the study definition excludes. Never filter a row out because "
        "it is an extreme statistical value; report the refit as a sensitivity "
        "analysis instead.",
    ],
]


# ── 4. The comparison ladder ──────────────────────────────────────────────────
# Rungs are ordered by how much survives a change. Rung 1 is the only surface
# that crosses a response-transform change; rung 3 is where most of the
# workbook's statistics live and is same-transform only; rung 5 is not about
# fit quality at all.
_LADDER: list[list[str]] = [
    [
        "Rung 1 — comparable by construction, even across a "
        "transform change",
        "The unit-space fit block: R Squared (Unit) at Cell AH7, Adj R Squared "
        "(Unit) at Cell AH8 and RMSE (Unit) at Cell AH9, Regression Outputs. The "
        "Smearing Factor is at Cell AH6.",
        "Goodness of fit measured in original response units. The block is built "
        "for exactly this comparison: two models with different response "
        "transforms can be lined up on one scale, and the Response Space statement "
        "at Cell AH10 says which scale is on screen.",
        "The only across-transform evidence on the sheet. Both readings must share "
        "the Cell AH5 setting — see pre-flight condition 5. Under Fixed Effects "
        "with a logged response the unit-space values are total, not within-group.",
    ],
    [
        "Rung 2 — properties of the design and the sample, not of the response "
        "scale",
        "Observations at Cell AB9, the residual degrees of freedom at Cell AB17, "
        "the Mean Leverage at Cell AE7, and the design matrix's width at Cell O1, "
        "Model Specification.",
        "Identical under a response-transform change on the same sample, because "
        "none of them is computed from a residual. They are the numbers every "
        "penalty term and every degrees-of-freedom count is built from.",
        "Bound to the sample, not to the transform. Two specifications that differ "
        "here differ in population or in design size, which is a pre-flight failure "
        "rather than a comparison result.",
    ],
    [
        "Rung 3 — fit-space, so same-transform only",
        "Multiple R at Cell AB5, R Squared at Cell AB6, Adjusted R Squared at Cell "
        "AB7 and Standard Error at Cell AB8, Regression Statistics; PRESS at Cell "
        "AE5, PRESS R Squared at Cell AE6, AIC at Cell AE8, BIC at Cell AE9 and "
        "AICc at Cell AE10, Diagnostics; the F statistic at Cell AE16 and "
        "Significance F at Cell AF16, ANOVA Table.",
        "The workbook's main comparison surface. Every one is computed from "
        "residuals on the fitted response, so they answer how well this "
        "specification fits this transformed response.",
        "Meaningless across a response-transform change, and meaningless across a "
        "sample change. AIC, BIC and AICc are on a log-likelihood scale, not in "
        "response units — they are comparable to each other on a matched transform "
        "and nothing more.",
    ],
    [
        "Rung 4 — dimensionless, but still built from residuals",
        "Durbin-Watson at Cell AE12 and the BFN Panel Durbin-Watson at Cell AE13, "
        "Diagnostics; QQ Correlation at Cell AE11; Scale-Location in Col AW, "
        "Residual Output.",
        "These carry no units, so they will not show a scale difference the way a "
        "coefficient will — but each is a function of the residuals, so a transform "
        "change still moves them.",
        "Read them as trustworthiness evidence within one transform, not as "
        "cross-transform evidence. Both cells read n/a unless exactly one Sequence "
        "flag is declared, and the BFN cell additionally needs exactly one "
        "fixed-effects variable.",
    ],
    [
        "Rung 5 — about whether the fit can be trusted, not how good it is",
        "GVIF in Col X and Tolerance in Col Y, Predictor Summary; Hat Diagonal in "
        "Col AR, Studentized Residuals in Col AS, Cook's Distance in Col AT and "
        "PRESS Residual in Col AX, Residual Output.",
        "Collinearity among the predictors, and the influence of individual rows on "
        "the coefficients. A specification can win on Adjusted R² and still be "
        "unreportable because one row is driving it.",
        "Break a tie here. When two specifications are close on the fit statistics, "
        "prefer the one whose collinearity is lower and whose influence flags are "
        "quieter — then re-screen the winner against the Diagnostic Guide.",
    ],
]


# ── 5. Stability before reporting a coefficient ───────────────────────────────
_STABILITY: list[list[str]] = [
    [
        "Does the sign survive a specification change?",
        "Refit with the marginal terms toggled off in Col C, Model Specification, "
        "and compare the coefficient in Col AB, Coefficients.",
        "The sign and rough magnitude hold across the plausible specifications, so "
        "the effect is not an artifact of which controls were in the model.",
        "A coefficient that flips sign when a control is added is reporting the "
        "control, not the predictor. Report the range across specifications rather "
        "than a single column of numbers.",
    ],
    [
        "Is the term's own P-value stable?",
        "Col AD and Col AE, Coefficients, read at the same time as the Adjusted R² "
        "at Cell AB7, Regression Statistics.",
        "The term stays significant while Adjusted R² moves in the same direction — "
        "adding it bought something.",
        "A term that is significant but lowers Adjusted R² is spending a column for "
        "nothing. A term that is insignificant but raises Adjusted R² is worth "
        "keeping and reporting as inconclusive at this sample size.",
    ],
    [
        "Is the predictor carrying its own information?",
        "GVIF in Col X and Tolerance in Col Y, Predictor Summary.",
        "GVIF is modest and Tolerance is well above the Diagnostic Guide's "
        "thresholds, so the predictor's standard error is not inflated by the other "
        "predictors.",
        "A large GVIF means the coefficient's standard error is shared with the "
        "other predictors, so its P-value can be large for reasons that have "
        "nothing to do with the effect being absent. Use the Correlation_Matrix "
        "function to find the pairs.",
    ],
    [
        "Does the estimate depend on a few rows?",
        "Cook's Distance in Col AT, Hat Diagonal in Col AR and the PRESS Residual "
        "column in Col AX, Residual Output.",
        "No row stands out against F.INV(0.5, p, n-p), so the coefficients are not "
        "the property of a handful of observations.",
        "Refit without the flagged rows and report the shift, or report the "
        "coefficient with the caveat. Removing them outright changes the model, so "
        "it belongs in the spec as a Filter row if it belongs anywhere.",
    ],
    [
        "Are the residuals independent?",
        "Durbin-Watson at Cell AE12 and the BFN Panel Durbin-Watson at Cell AE13, "
        "Diagnostics, with the Spacing Verdict above the spec.",
        "Neither statistic shows first-order autocorrelation. Both read n/a rather "
        "than a clean bill of health wherever the specification declares no panel "
        "structure. Read the BFN statistic directionally — near 2 means no "
        "autocorrelation — but never against Durbin-Watson's published bounds: its "
        "critical values depend on the panel's N and T and are not surfaced here.",
        "Serial correlation makes the standard errors too small, so the P-values "
        "read as more certain than the evidence supports. Fix the specification "
        "before reporting a coefficient from it.",
    ],
]


# ── 6. What survives a broken assumption ──────────────────────────────────────
_ASSUMPTION_EVIDENCE: list[list[str]] = [
    [
        "Non-constant error variance (heteroscedasticity)",
        "The reported precision. Coefficients stay unbiased and consistent under "
        "OLS, but they are no longer minimum-variance, and the standard errors, t "
        "statistics, P-values and interval bounds are all computed on one common "
        "variance — so they are the parts that misstate. The coefficients "
        "themselves remain the right point estimates.",
        "Screen it on the Residuals vs. Fitted and Scale-Location charts before "
        "trusting any P-value, then address the shape in the specification: a Log "
        "transform on the response, or an interaction or quadratic term when the "
        "variance tracks a predictor's mean. Weighted Least Squares is planned, not "
        "shipped.",
    ],
    [
        "Non-normal errors",
        "The interval-based inference at small samples. QQ Correlation at Cell "
        "AE11 reads below 0.98 as a mild concern and below 0.95 as a clear one; as "
        "the sample grows, the central limit theorem carries much of the "
        "coefficient inference.",
        "Report the sample size alongside the inference, and read the Q-Q chart "
        "rather than the correlation alone. Robust standard errors address "
        "variance misspecification, not non-normality, and a bootstrap has to "
        "respect dependence in the data — neither ships here.",
    ],
    [
        "Serial correlation or dependence",
        "The independence-based standard errors. Durbin-Watson at Cell AE12 "
        "requires exactly one Sequence flag and no fixed-effects variable; the BFN "
        "Panel Durbin-Watson at Cell AE13 requires exactly one of each. Without a "
        "Sequence flag both read n/a — requires Sequence, which is a statement "
        "about the specification, not a clean bill of health.",
        "Declare the Sequence flag so the ordering axis is known, then read the "
        "statistic that matches the structure. The coefficients are unaffected; it "
        "is the precision claim that has to be qualified in the write-up.",
    ],
]


# ── 7. Moves that look like comparisons and are not ───────────────────────────
_INVALID_MOVES: list[list[str]] = [
    [
        "R² went up, so this specification is better.",
        "R² never falls when a predictor is added, so it rewards adding noise — and "
        "the version of the model with more columns always wins the comparison by "
        "construction, regardless of whether the extra columns carry information.",
        "Read Adjusted R Squared at Cell AB7 instead, which charges for the columns "
        "spent, and check the added term's own P-value in Col AE, Coefficients. "
        "Report both models' column counts from Cell O1 if the change is marginal.",
    ],
    [
        "AIC is lower, so this specification wins.",
        "Lower is better within a comparison, so the number reads like a verdict. "
        "But AIC, BIC and AICc at Cell AE8:AE10 are fit-space: they are only "
        "comparable when the response, the transform, the sample and n all match, "
        "because n enters the penalty term of every one of them.",
        "Hold the five pre-flight conditions first, then compare by hand. This "
        "workbook ships no delta-AIC, no Akaike weights and no ranking — it "
        "displays one fitted model's criteria at a time, and the arithmetic is "
        "yours.",
    ],
    [
        "The coefficients changed, so the model changed.",
        "Retargeting the Reference Level in Col E, Model Specification, rewrites "
        "the intercept, every coefficient of that categorical and any interaction "
        "term built on it, so the coefficient table looks substantially different. "
        "Nothing about the fit changed: the design spans the same space, so the "
        "fitted values, R Squared and every residual diagnostic are identical.",
        "Check whether R², Adjusted R Squared, Standard Error and the information "
        "criteria moved — if they did not, nothing about the fit changed. Compare "
        "coefficient tables only between runs with the same reference level, or "
        "compare the fitted values instead.",
    ],
    [
        "The two models came out close on Adjusted R², so I picked the one I "
        "preferred anyway.",
        "A near-tie on any single statistic is not a finding, and the workbook's "
        "statistics are computed to full precision from a sample that has its own "
        "sampling error — a gap of a few thousandths is not evidence of a "
        "difference.",
        "Settle a close call on rung 5 of the ladder: prefer the specification "
        "whose collinearity is lower (Col X, Predictor Summary), whose influence "
        "flags are quieter, and whose residual plots are cleaner. Say in the "
        "write-up that the fit statistics did not separate them.",
    ],
]


# ── 8. The decision procedure ─────────────────────────────────────────────────
# One string per step; the "Step N" label in column 1 is written from the loop
# index, so a step's text never carries its own number.
_SPEC_STEPS: list[list[str]] = [
    [
        "Fix the response and the population first. Choose the Response variable "
        "and set the Filter rows in Col B, Model Specification — these two "
        "decisions define the question, and changing either while comparing voids "
        "the comparison rather than enriching it.",
    ],
    [
        "Establish the five pre-flight conditions across every candidate: the same "
        "response, a matched transform, the same sample at Cell AB9, the same "
        "fixed-effects structure and the same Cell AH5 setting. A candidate that "
        "fails one of them belongs in its own comparison.",
    ],
    [
        "Declare one response-transform policy and hold it. If the comparison "
        "genuinely has to cross a transform change, the only legitimate surface is "
        "the unit-space triplet at Cell AH7:AH9 — every other statistic that "
        "measures fit is fit-space.",
    ],
    [
        "Fit the intercept-only baseline. Turn every Include toggle in Col C off "
        "and leave the intercept on at Cell C2, then record its R Squared at Cell "
        "AB6 and its Standard Error at Cell AB8 as the bar to beat.",
    ],
    [
        "Add predictors one decision at a time. Each addition is a single "
        "question — does this term carry information the others do not? Read the "
        "new term's P-value in Col AE, Coefficients, and the Adjusted R Squared at "
        "Cell AB7, and watch the design matrix's width grow at Cell O1.",
    ],
    [
        "Stop on evidence, not on habit. There is no stepwise, forward, backward or "
        "best-subset search in this workbook — the Include toggle is manual by "
        "design. Stop when the last term added does not earn the column it spends: "
        "an insignificant P-value with a lower Adjusted R Squared.",
    ],
    [
        "Compare the survivors down the ladder in order: unit-space evidence across "
        "a transform change, then Adjusted R Squared, PRESS R Squared and the "
        "information criteria within one transform, then the trustworthiness rung "
        "to settle a near-tie.",
    ],
    [
        "Re-screen the winner's diagnostics. It has not been screened until it has "
        "been through the same Diagnostic Guide charts as the losers; a "
        "specification that wins on Adjusted R Squared and fails on a funnel plot "
        "is not the one to report.",
    ],
    [
        "State what the preference does not license. Write down which statistics "
        "the choice was made on and which pre-flight conditions held, and do not "
        "describe a fit-space difference as though it survived a transform change.",
    ],
]


# ── 9. Deliberate absences ────────────────────────────────────────────────────
_ABSENCES: list[list[str]] = [
    [
        "No model-selection algorithm of any kind",
        "There is no stepwise, forward, backward or best-subset search, and no "
        "automatic variable screening. The Include toggle in Col C, Model "
        "Specification, is a manual TRUE/FALSE switch, and the specification is "
        "whatever the analyst declares.",
        "Choose the specification yourself and record why. The decision procedure "
        "above is the substitute for a search algorithm, and it leaves an audit "
        "trail a search would not.",
    ],
    [
        "No information-criterion comparison machinery",
        "AIC, BIC and AICc at Cell AE8:AE10 display one fitted model at a time. "
        "There is no delta-AIC, no Akaike weight, no evidence ratio and no ranking "
        "across specifications.",
        "Read the three cells for each candidate and do the differencing yourself, "
        "after confirming the pre-flight conditions. The workbook deliberately "
        "stops short of turning the comparison into an answer.",
    ],
    [
        "No unit-space or Jacobian-corrected information criteria",
        "Likelihood is defined with respect to a particular response variable, so "
        "comparing likelihoods across a transform change requires the Jacobian of "
        "that transformation — the correct comparison is the original response's "
        "likelihood evaluated at the back-transformed prediction. That is "
        "deliberately deferred, not overlooked.",
        "Use the unit-space triplet at Cell AH7:AH9, which is the comparable-by-"
        "construction surface that does ship. Do not treat an AIC difference "
        "between a logged and an untransformed response as evidence.",
    ],
    [
        "No nested or partial F test",
        "Significance F at Cell AF16, ANOVA Table, compares the fitted model to the "
        "intercept-only model. It never compares two fitted specifications to each "
        "other, however nested they are.",
        "Compare nested specifications on Adjusted R² at Cell AB7 and on the "
        "coefficient P-values in Col AE, Coefficients. A hand-computed partial F "
        "needs sums of squares this sheet does not lay out as a pair.",
    ],
    [
        "No DFFITS, DFBETAS, condition number or Mallows' Cp",
        "The influence surface is Cook's Distance in Col AT, the PRESS Residual "
        "column in Col AX, Hat Diagonal in Col AR and Studentized Residuals in Col "
        "AS — plus the PRESS total at Cell AE5 and PRESS R Squared at Cell AE6.",
        "That set answers whether individual rows drive the fit. Per-coefficient "
        "influence (which row moves which coefficient) is not computed; if a row is "
        "flagged, refit without it and compare coefficient tables directly.",
    ],
    [
        "No cross-validation surface beyond PRESS",
        "Leave-one-out cross-validation appears only as the PRESS total at Cell "
        "AE5, PRESS R Squared at Cell AE6 and the per-row PRESS Residual column in "
        "Col AX. There is no k-fold split, no holdout partition and no repeated "
        "resampling.",
        "PRESS is leave-one-out prediction error, so it is the honest out-of-sample "
        "comparison this workbook offers. Prefer PRESS R Squared over R² when two "
        "specifications spend different numbers of columns.",
    ],
    [
        "No conditional formatting on Tolerance",
        "GVIF in Col X, Predictor Summary, is coloured by threshold; Tolerance in "
        "Col Y beside it carries no conditional formatting at all, so a failing "
        "Tolerance does not turn red.",
        "Screen Tolerance by eye against the Diagnostic Guide's threshold table — "
        "below 0.2 is a possible collinearity signal and below 0.1 a strong one. "
        "The two columns are reciprocals, so they cannot disagree, but only one of "
        "them will catch your attention for you.",
    ],
]


# ── 10. Using the Model Comparison sheet ──────────────────────────────────────
# The hands-on half: how a new model gets onto the sheet at all. The first row is
# the one that goes wrong silently — a model sheet built by pasting cells reads
# the ORIGINAL sheet's sheet-scoped names, so it shows the old model's numbers.
_USING_THE_SHEET: list[list[str]] = [
    [
        "Create another model — copy the whole sheet",
        "Right-click the sheet tab of the Regression template sheet, or of any "
        "model sheet already fitted, then choose Move or Copy… and tick Create a "
        "copy.",
        "Rename the copy, then edit its Model Specification. A sheet copied this "
        "way arrives with its own set of sheet-scoped names — the spec wiring, the "
        "constructor closures, the chart ranges and the Comparison_Anchor the "
        "registry reads — every one scoped to the new sheet, so the copy fits its "
        "own specification and nothing else.",
        "Never select the cells and copy, or cut, and paste them onto a blank "
        "sheet. The formulas come across but the sheet-scoped names do not, so "
        "every formula keeps reading the ORIGINAL sheet: the new sheet shows the "
        "old model's numbers. Nothing errors — it is a silent wrong answer.",
    ],
    [
        "Point a spare row at the new model",
        "Formulas > Name Manager, with the Model Comparison sheet active. The "
        "Anchor column on that sheet shows the address each row reads.",
        "Every row owns a sheet-scoped name, Comp_Anchor_1, Comp_Anchor_2 and so "
        "on, numbered from the first data row (row 4 is Comp_Anchor_1). The build "
        "ships three spare rows whose anchors point at the sheet's own $A$1. Edit "
        "the first spare one so Refers To reads ='<new sheet>'!$AF$3 — the model "
        "sheet's response-label cell — and every cell in the row fills in.",
        "A spare row reads blank until it is pointed at a model sheet, and a "
        "mis-pointed row stays blank rather than reading an unrelated cell: the "
        "readers check for the MODEL SPECIFICATION heading at the target's A1 "
        "before reading anything.",
    ],
    [
        "Add a row when the spare rows run out",
        "Name Manager first, then the first empty row directly below the table.",
        "1. Name Manager > New: Name Comp_Anchor_<n+1>, the next unused number; "
        "Scope Model Comparison, not Workbook; Refers To ='<model sheet>'!$AF$3. "
        "2. Select the table's last row across every column and copy it down one "
        "row. 3. With only the new row selected, Find & Replace Comp_Anchor_<n> "
        "with Comp_Anchor_<n+1>, looking in Formulas.",
        "A copied row still reads the name the row above it reads, so skipping "
        "step 3 duplicates the previous model instead of adding one. Replace "
        "within the new row only. The gate flags, the row-2 verdicts and the zone "
        "shading already reach row 103, so a new row down to there is counted with "
        "no other edit.",
    ],
    [
        "Choose the reference model",
        "The first data row, row 4 — the name Comp_Anchor_1.",
        "Every gate column (Same Set?, Same Space?, Same Method?, Same Inputs?) "
        "compares its row against this one. To compare the others against a "
        "different model, point Comp_Anchor_1 at that model's sheet.",
        "The row-2 verdicts count how many rows agree with the reference, so a "
        "reference that is itself the odd one out shades every other row. Pick "
        "the specification the rest are variations on.",
    ],
]


# ── 11. Where to go next ──────────────────────────────────────────────────────
_SIBLINGS: list[list[str]] = [
    [
        "Regression Instructions",
        "The operational how: pointing the sheet at a dataset, what each spec "
        "column declares, and how a Role or a Transform token changes the fitted "
        "model.",
    ],
    [
        "Modeling Concepts",
        "The conceptual why: one row per modeling feature, the statistical method "
        "it enables, and a concrete situation for using it — including the features "
        "declared but not yet built.",
    ],
    [
        "Diagnostic Guide",
        "What to look for in one fitted model's residual plots and influence "
        "columns, with the numeric threshold table this sheet cites by column "
        "letter.",
    ],
]


def _heading(sheet: xw.Sheet, row: int, text: str) -> None:
    cell = sheet.range((row, 1))
    cell.value = text
    cell.api.Font.Bold = True
    sheet.range((row, 1), (row, _LAST_COL)).color = _HEADER_COLOR


def _subheading(sheet: xw.Sheet, row: int, text: str, cols: int = _LAST_COL) -> None:
    cell = sheet.range((row, 1))
    cell.value = text
    cell.api.Font.Bold = True
    sheet.range((row, 1), (row, cols)).color = _SUBHEADER_COLOR


def _table_header_row(sheet: xw.Sheet, row: int, headers: list[str]) -> None:
    for col, text in enumerate(headers, start=1):
        cell = sheet.range((row, col))
        cell.value = text
        cell.api.Font.Bold = True
        cell.color = _SUBHEADER_COLOR


def _row(sheet: xw.Sheet, row: int, values: list[str]) -> None:
    for col, text in enumerate(values, start=1):
        cell = sheet.range((row, col))
        cell.value = text
        cell.api.WrapText = True


def _write_template_sheet(workbook: xw.Book) -> None:
    """(Re)build this sheet's content directly.

    Used only to author ``TEMPLATE_PATH`` — artifact builds never call this;
    they call ``write_model_comparison_guide_sheet``, which copies the sheet
    this function last wrote into the template.
    """
    sheet = get_or_create_sheet(workbook, SHEET_NAME)
    reset_generated_sheet(sheet)

    set_column_widths(sheet, ((c.index, c.width) for c in _COLUMNS))

    r = 1

    # ── Title ──────────────────────────────────────────────────────────────────
    _heading(sheet, r, "MODEL COMPARISON GUIDE — WHICH MODEL, AND ON WHAT EVIDENCE")
    r += 1
    sheet.range((r, 1)).value = (
        "Fitting a model is the easy half. The hard half is deciding which of several "
        "specifications to report, on what evidence, and whether lining them up is "
        "legitimate at all. This sheet answers that in four layers: the conditions that "
        "make two models comparable, what each statistic is for and what change it "
        "survives, a step-by-step procedure for choosing a specification, and the moves "
        "that look like comparisons and are not. It is a reading guide only — it adds no "
        "formulas and no new named ranges. The Model Comparison sheet beside it is the "
        "registry this guide explains: one row per fitted model, with the comparability "
        "tests the first layer below describes already applied to each row. For the "
        "mechanics of setting a specification up, see the "
        "Regression Instructions sheet; for what each modeling feature does, see Modeling "
        "Concepts; for what to look for in the residual plots, see the Diagnostic Guide."
    )
    intro = sheet.range((r, 1), (r, _LAST_COL))
    intro.api.WrapText = True
    try:
        intro.api.HorizontalAlignment = 7  # xlHAlignCenterAcrossSelection
    except Exception:
        pass
    r += 2

    # ── 1. Pre-flight ──────────────────────────────────────────────────────────
    _subheading(sheet, r, "PRE-FLIGHT — FIVE THINGS THAT MUST MATCH, OR THE COMPARISON IS VOID")
    r += 1
    _table_header_row(
        sheet, r, ["Condition", "Where to check it", "Legal when", "If it does not match"]
    )
    r += 1
    for vals in _PREFLIGHT:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 2. Change ledger ───────────────────────────────────────────────────────
    _subheading(sheet, r, "WHAT EACH EDIT DOES TO A COMPARISON — FIND YOUR ROW")
    r += 1
    _table_header_row(
        sheet, r,
        ["The edit", "What it moves", "What it leaves alone", "What that means for a comparison"],
    )
    r += 1
    for vals in _CHANGE_LEDGER:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 3. The analyst's questions ─────────────────────────────────────────────
    _subheading(sheet, r, "THE QUESTIONS, IN THE ORDER AN ANALYST ASKS THEM")
    r += 1
    _table_header_row(
        sheet, r, ["Question", "Where the answer lives", "What it tells you", "What it does not"]
    )
    r += 1
    for vals in _QUESTIONS:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 4. The comparison ladder ───────────────────────────────────────────────
    _subheading(sheet, r, "THE COMPARISON LADDER — WHAT TO REACH FOR, IN ORDER")
    r += 1
    _table_header_row(
        sheet, r, ["Rung", "Where it is", "What it tells you", "What it does not survive"]
    )
    r += 1
    for vals in _LADDER:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 5. Stability before reporting a coefficient ────────────────────────────
    _subheading(sheet, r, "STABILITY BEFORE YOU REPORT A COEFFICIENT")
    r += 1
    _table_header_row(
        sheet, r, ["Check", "Where to look", "Stable", "Not stable — what to do"]
    )
    r += 1
    for vals in _STABILITY:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 6. What survives a broken assumption ───────────────────────────────────
    _subheading(sheet, r, "WHICH EVIDENCE SURVIVES A BROKEN ASSUMPTION", cols=3)
    r += 1
    _table_header_row(sheet, r, ["Broken assumption", "What is compromised", "What to do instead"])
    r += 1
    for vals in _ASSUMPTION_EVIDENCE:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 7. Invalid moves ───────────────────────────────────────────────────────
    _subheading(sheet, r, "FOUR MOVES THAT LOOK LIKE A COMPARISON AND ARE NOT", cols=3)
    r += 1
    _table_header_row(sheet, r, ["The move", "Why it misleads", "What to do instead"])
    r += 1
    for vals in _INVALID_MOVES:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 8. The decision procedure ──────────────────────────────────────────────
    _subheading(sheet, r, "THE SPECIFICATION DECISION PROCEDURE", cols=2)
    r += 1
    sheet.range((r, 1)).value = (
        "Nine steps, in order. Steps 1 to 3 are the conditions; steps 4 to 6 build the "
        "candidates; steps 7 to 9 choose between them and state the limits of the choice."
    )
    sheet.range((r, 1)).api.WrapText = True
    r += 2
    _table_header_row(sheet, r, ["Step", "What to do"])
    r += 1
    for index, vals in enumerate(_SPEC_STEPS, start=1):
        _row(sheet, r, [f"Step {index}", vals[0]])
        r += 1
    r += 1

    # ── 9. Deliberate absences ─────────────────────────────────────────────────
    _subheading(sheet, r, "WHAT THIS WORKBOOK DELIBERATELY DOES NOT DO", cols=3)
    r += 1
    _table_header_row(sheet, r, ["Not provided", "What that means", "What to do instead"])
    r += 1
    for vals in _ABSENCES:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 10. Using the Model Comparison sheet ────────────────────────────────────
    _subheading(sheet, r, "USING THE MODEL COMPARISON SHEET — ADDING MODELS AND ROWS")
    r += 1
    _table_header_row(sheet, r, ["Task", "Where", "What to do", "Why it matters"])
    r += 1
    for vals in _USING_THE_SHEET:
        _row(sheet, r, vals)
        r += 1
    r += 1

    # ── 11. Where to go next ───────────────────────────────────────────────────
    _subheading(sheet, r, "WHERE TO GO NEXT ON THE OTHER REFERENCE SHEETS", cols=2)
    r += 1
    _table_header_row(sheet, r, ["Sheet", "What it answers"])
    r += 1
    for vals in _SIBLINGS:
        _row(sheet, r, vals)
        r += 1

    sheet.autofit("rows")

    # Keep the title, centered introduction, and spacer visible while scrolling.
    try:
        sheet.activate()
        window = sheet.api.Application.ActiveWindow
        window.FreezePanes = False
        window.Split = False
        sheet.range("A5").select()
        window.FreezePanes = True
    except Exception:
        pass


def write_model_comparison_guide_sheet(workbook: xw.Book) -> None:
    """Create or refresh the Model Comparison Guide reference sheet from the static template.

    Parameters
    ----------
    workbook : xw.Book
        The open xlwings workbook to receive the sheet.
    """
    copy_static_sheet(workbook, TEMPLATE_PATH, SHEET_NAME)


def _main() -> None:
    """Regenerate this sheet inside ``templates/static_sheets.xlsx``.

    Prefer ``python scripts/rebuild_static_sheets.py`` — it regenerates every
    static sheet so two edited modules cannot half-regenerate. Run after
    editing the content above, then commit the updated template file.
    """
    parser = argparse.ArgumentParser(
        description=(
            f"Rebuild the {SHEET_NAME!r} sheet inside the static template workbook."
        )
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=TEMPLATE_PATH,
        help="Path to the static template workbook (default: templates/static_sheets.xlsx).",
    )
    args = parser.parse_args()

    template_path = args.template.resolve()
    try:
        with xw.App(visible=True, add_book=False) as app:
            workbook, _ = open_or_create_workbook(app, template_path)
            try:
                _write_template_sheet(workbook)
                workbook.save(str(template_path))
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(template_path, "open or save", exc)

    print(f"Template sheet updated: {SHEET_NAME}")
    print(f"Template workbook: {template_path}")


if __name__ == "__main__":
    _main()
