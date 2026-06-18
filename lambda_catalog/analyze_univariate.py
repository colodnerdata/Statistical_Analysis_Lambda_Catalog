"""Reference implementations for validating the Excel LAMBDA univariate outputs.

These functions use scipy and numpy as *independent* oracles.  The purpose is
to catch divergence between the Excel formulas and statistically correct
results — not to reproduce the same formula twice.

Intentional differences from the LAMBDA formulas
─────────────────────────────────────────────────
Skewness / kurtosis
    scipy.stats.skew(bias=False) and scipy.stats.kurtosis(bias=False) are the
    reference.  Excel's SKEW and KURT implement the same Fisher-Pearson
    estimator, so agreement validates the formula.  Disagreement flags an
    Excel formula error.

Normal / Lognormal / Exponential MLE
    scipy.stats.<dist>.fit() returns true MLE parameters (Normal uses ddof=0
    for σ; Excel's closed-form uses STDEV.S, i.e. ddof=1).  For n ≥ 30 the
    difference is negligible; for small n the discrepancy is expected and QC
    tests should use a tolerance of ~1/n.

Weibull / Gamma MLE
    scipy numerically optimises the likelihood — this is the reference that
    the v2.1 grid-search results will be validated against.

Triangular / BetaPERT
    No scipy.fit for these; parameters come from the data min/mode/max.
    The NLL is still evaluated via scipy, providing an independent check of
    the LAMBDA's hand-coded PDF implementation.

Histogram bin counts
    bin_counts() implements the FREQUENCY convention directly (prev < x ≤ edge)
    using numpy comparisons, which is independent of FREQUENCY itself.
"""
from __future__ import annotations

import math
from typing import Sequence

import numpy as np
from scipy import stats as scipy_stats


# ── Internal helpers ──────────────────────────────────────────────────────────

def _clean(data: Sequence[float | None]) -> np.ndarray:
    """Return a 1-D float array with None/NaN entries removed."""
    return np.array(
        [x for x in data if x is not None and not (isinstance(x, float) and math.isnan(x))],
        dtype=float,
    )


# ── Descriptive statistics ────────────────────────────────────────────────────

def missing_count(data: Sequence[float | None]) -> int:
    """Count of None/NaN values (blank or non-numeric cells).

    Counts from the start of the sequence to the last non-None entry, matching
    the LAMBDA's MATCH-based active-range heuristic.
    """
    items = list(data)
    # Find last non-None index
    last = -1
    for i, v in enumerate(items):
        if v is not None and not (isinstance(v, float) and math.isnan(v)):
            last = i
        elif v is not None:
            last = i   # non-None text/bad value still counts as "active row"
    if last < 0:
        return 0
    active = items[: last + 1]
    return sum(
        1 for v in active
        if v is None or (isinstance(v, float) and math.isnan(v))
    )


def descriptive_stats(data: Sequence[float | None]) -> dict[str, float | int]:
    """Compute the 12 summary statistics shown on the Univariate sheet.

    Uses scipy.stats for skewness and kurtosis so the reference is
    independent of Excel's SKEW/KURT formula.  Agreement with the LAMBDA
    validates the Excel formulas; disagreement flags an error.

    Returns
    -------
    dict with keys: mean, median, mode, sd, variance, min, max, range,
                    skewness, kurtosis, count, missing_count
    """
    x = _clean(data)
    n = int(len(x))
    mc = missing_count(data)

    try:
        mode_result = scipy_stats.mode(x, keepdims=False)
        mode_val = float("nan") if mode_result.count <= 1 else float(mode_result.mode)
    except Exception:
        mode_val = float("nan")

    return {
        "mean":          float(np.mean(x)),
        "median":        float(np.median(x)),
        "mode":          mode_val,
        "sd":            float(np.std(x, ddof=1)),
        "variance":      float(np.var(x, ddof=1)),
        "min":           float(np.min(x)),
        "max":           float(np.max(x)),
        "range":         float(np.max(x) - np.min(x)),
        # scipy bias=False matches the bias-corrected Fisher-Pearson estimator
        # used by Excel SKEW and KURT.
        "skewness":      float(scipy_stats.skew(x, bias=False)),
        "kurtosis":      float(scipy_stats.kurtosis(x, bias=False, fisher=True)),
        "count":         n,
        "missing_count": mc,
    }


# ── Histogram binning ─────────────────────────────────────────────────────────

def sturges_bins(data: Sequence[float | None]) -> int:
    """Sturges bin count: ceil(log2(n) + 1)."""
    x = _clean(data)
    n = int(np.sum(np.isfinite(x)))
    return math.ceil(math.log2(n) + 1)


def scott_bins(data: Sequence[float | None]) -> int:
    """Scott bin count from width 3.49 * SD * n^(-1/3)."""
    x = _clean(data)
    n = int(np.sum(np.isfinite(x)))
    sd = float(np.std(x, ddof=1))
    if sd == 0.0:
        return 1
    width = 3.49 * sd * n ** (-1.0 / 3.0)
    data_range = float(np.max(x) - np.min(x))
    return math.ceil(data_range / width)


def fd_bins(data: Sequence[float | None]) -> int:
    """Freedman-Diaconis bin count from width 2 * IQR * n^(-1/3).

    IQR uses scipy.stats.iqr (linear interpolation), which matches
    QUARTILE.INC (type 7) — the same method used by the LAMBDA.
    """
    x = _clean(data)
    n = int(np.sum(np.isfinite(x)))
    iqr = float(scipy_stats.iqr(x))
    if iqr == 0:
        return sturges_bins(data)
    width = 2.0 * iqr * n ** (-1.0 / 3.0)
    data_range = float(np.max(x) - np.min(x))
    return math.ceil(data_range / width)


def bin_edges(data: Sequence[float | None], method: str) -> np.ndarray:
    """Return upper bin edges for the given method (Sturges, Scott, or FD).

    Evenly spaced from min to max; the bin count comes from the chosen rule.
    """
    x = _clean(data)
    min_x, max_x = float(np.min(x)), float(np.max(x))

    if method == "Sturges":
        k = sturges_bins(data)
    elif method == "Scott":
        k = scott_bins(data)
    elif method == "FD":
        k = fd_bins(data)
    else:
        raise ValueError(f"Unknown method {method!r}. Use 'Sturges', 'Scott', or 'FD'.")

    width = (max_x - min_x) / k
    return np.array([min_x + (i + 1) * width for i in range(k)])


def bin_counts(data: Sequence[float | None], edges: np.ndarray) -> np.ndarray:
    """Frequency counts matching Excel FREQUENCY's (prev < x ≤ edge) convention.

    Implemented with numpy comparisons — independent of Excel FREQUENCY itself.
    """
    x = _clean(data)
    bins = np.asarray(edges, dtype=float)
    counts = []
    prev = -np.inf
    for edge in bins:
        counts.append(int(np.sum((x > prev) & (x <= edge))))
        prev = float(edge)
    return np.array(counts, dtype=int)


# ── Negative log-likelihood functions ────────────────────────────────────────
# NLL = -Σ log f(xᵢ | params).  All use scipy's logpdf, which is the
# independent reference the LAMBDA formulas are validated against.

def nll_normal(data: Sequence[float | None], mean: float, sd: float) -> float:
    """NLL for Normal(mean, sd) via scipy.stats.norm."""
    x = _clean(data)
    return float(-np.sum(scipy_stats.norm.logpdf(x, loc=mean, scale=sd)))


def nll_lognormal(
    data: Sequence[float | None], meanlog: float, sdlog: float
) -> float:
    """NLL for Lognormal(meanlog, sdlog) via scipy.stats.lognorm.

    meanlog = μ of ln(x), sdlog = σ of ln(x).
    """
    x = _clean(data)
    return float(-np.sum(scipy_stats.lognorm.logpdf(x, s=sdlog, scale=math.exp(meanlog))))


def nll_exponential(data: Sequence[float | None], rate: float) -> float:
    """NLL for Exponential(rate) via scipy.stats.expon.  rate = 1/mean."""
    x = _clean(data)
    return float(-np.sum(scipy_stats.expon.logpdf(x, scale=1.0 / rate)))


def nll_weibull(data: Sequence[float | None], shape: float, scale: float) -> float:
    """NLL for Weibull(shape, scale) via scipy.stats.weibull_min."""
    x = _clean(data)
    log_pdf = scipy_stats.weibull_min.logpdf(x, c=shape, scale=scale)
    if not np.all(np.isfinite(log_pdf)):
        return 1e15
    return float(-np.sum(log_pdf))


def nll_gamma(data: Sequence[float | None], shape: float, rate: float) -> float:
    """NLL for Gamma(shape, rate) via scipy.stats.gamma.  rate = 1/scale."""
    if rate <= 0.0 or shape <= 0.0:
        return 1e15
    x = _clean(data)
    log_pdf = scipy_stats.gamma.logpdf(x, a=shape, scale=1.0 / rate)
    if not np.all(np.isfinite(log_pdf)):
        return 1e15
    return float(-np.sum(log_pdf))


def nll_triangular(
    data: Sequence[float | None], min_val: float, mode_val: float, max_val: float
) -> float:
    """NLL for Triangular(min, mode, max) via scipy.stats.triang.

    Returns 1e15 for degenerate parameters or out-of-support data points.
    This matches the IFERROR(…, 1E+15) sentinel in the LAMBDA.
    """
    if mode_val <= min_val or mode_val >= max_val:
        return 1e15
    x = _clean(data)
    if np.any(x < min_val) or np.any(x > max_val):
        return 1e15
    data_range = max_val - min_val
    c = (mode_val - min_val) / data_range
    log_pdf = scipy_stats.triang.logpdf(x, c=c, loc=min_val, scale=data_range)
    if not np.all(np.isfinite(log_pdf)):
        return 1e15
    return float(-np.sum(log_pdf))


def nll_beta(
    data: Sequence[float | None], alpha_param: float, beta_param: float
) -> float:
    """NLL for Beta(alpha, beta) via scipy.stats.beta.  Data must be on (0,1).

    Returns 1e15 when any value is exactly at 0 or 1, matching the LAMBDA sentinel.
    """
    x = _clean(data)
    log_pdf = scipy_stats.beta.logpdf(x, a=alpha_param, b=beta_param)
    if not np.all(np.isfinite(log_pdf)):
        return 1e15
    return float(-np.sum(log_pdf))


def nll_betapert(
    data: Sequence[float | None], min_val: float, mode_val: float, max_val: float
) -> float:
    """NLL for BetaPERT(min, mode, max) via scipy.stats.beta with Jacobian.

    PERT reparameterisation: α = 1 + 4(mode−min)/(max−min),
    β = 1 + 4(max−mode)/(max−min).  Data are rescaled to z ∈ [0,1];
    the log-Jacobian −log(max−min) corrects back to the original scale.
    """
    x = _clean(data)
    data_range = max_val - min_val
    alpha_p = 1.0 + 4.0 * (mode_val - min_val) / data_range
    beta_p   = 1.0 + 4.0 * (max_val  - mode_val) / data_range
    z = (x - min_val) / data_range
    log_pdf = scipy_stats.beta.logpdf(z, a=alpha_p, b=beta_p) - math.log(data_range)
    if not np.all(np.isfinite(log_pdf)):
        return 1e15
    return float(-np.sum(log_pdf))


# ── MLE parameter estimates (independent of the LAMBDA closed-form) ───────────
# These use scipy's optimised fitting, not the same algebra the LAMBDAs use.
# Comparing these parameters with the LAMBDA's fitted values validates the
# closed-form expressions.

def mle_normal(data: Sequence[float | None]) -> tuple[float, float]:
    """True MLE for Normal via scipy: (mean, sd).

    scipy uses ddof=0 (true MLE); the LAMBDA uses STDEV.S (ddof=1).
    The difference = σ * sqrt(1 - 1/(n-1)).  For n ≥ 30 it is negligible.
    """
    x = _clean(data)
    loc, scale = scipy_stats.norm.fit(x)
    return float(loc), float(scale)


def mle_lognormal(data: Sequence[float | None]) -> tuple[float, float]:
    """True MLE for Lognormal via scipy: (meanlog, sdlog).

    floc=0 fixes the location to 0 (two-parameter fit matching the LAMBDA).
    """
    x = _clean(data)
    s, loc, scale = scipy_stats.lognorm.fit(x, floc=0)
    meanlog = math.log(scale)
    sdlog = s
    return float(meanlog), float(sdlog)


def mle_exponential(data: Sequence[float | None]) -> float:
    """True MLE for Exponential via scipy: rate = 1/mean.

    floc=0 fixes the location to 0; MLE rate = 1/sample_mean.
    """
    x = _clean(data)
    _loc, scale = scipy_stats.expon.fit(x, floc=0)
    return float(1.0 / scale)


def mle_weibull(data: Sequence[float | None]) -> tuple[float, float]:
    """True MLE for Weibull via scipy: (shape, scale).

    Used to validate v2.1 grid-search results.  floc=0 fixes location to 0.
    """
    x = _clean(data)
    c, _loc, scale = scipy_stats.weibull_min.fit(x, floc=0)
    return float(c), float(scale)


def mle_gamma(data: Sequence[float | None]) -> tuple[float, float]:
    """True MLE for Gamma via scipy: (shape, rate).

    Returns rate = 1/scale.  floc=0 fixes location to 0.
    Used to validate v2.1 grid-search results.
    """
    x = _clean(data)
    a, _loc, scale = scipy_stats.gamma.fit(x, floc=0)
    return float(a), float(1.0 / scale)


def mle_triangular(data: Sequence[float | None]) -> tuple[float, float, float]:
    """Estimate Triangular parameters: (min, mode, max).

    True MLE is non-differentiable at the mode; direct min/mode/max estimation
    is used instead, matching the LAMBDA's approach.  When no unique mode exists
    (all values distinct), falls back to the median — matching the LAMBDA's
    IFERROR(MODE.SNGL, MEDIAN()) fallback.
    """
    x = _clean(data)
    min_val = float(np.min(x))
    max_val = float(np.max(x))

    mode_result = scipy_stats.mode(x, keepdims=False)
    # Treat as no unique mode when every value appears exactly once
    if mode_result.count <= 1 and np.sum(x == mode_result.mode) <= 1:
        mode_val = float(np.median(x))
    else:
        mode_val = float(mode_result.mode)

    # Ensure mode is strictly interior to (min, max)
    if mode_val <= min_val or mode_val >= max_val:
        mode_val = float(np.median(x))
    if mode_val <= min_val or mode_val >= max_val:
        mode_val = (min_val + max_val) / 2.0

    return min_val, mode_val, max_val


def mle_betapert(data: Sequence[float | None]) -> tuple[float, float, float]:
    """Estimate BetaPERT parameters: (min, mode, max).  Same as Triangular."""
    return mle_triangular(data)


# ── Goodness-of-fit statistics ────────────────────────────────────────────────

def gof_aic(nll: float, k: int) -> float:
    """AIC = 2k + 2·NLL."""
    return 2.0 * k + 2.0 * nll


def gof_bic(nll: float, k: int, n: int) -> float:
    """BIC = k·ln(n) + 2·NLL."""
    return k * math.log(n) + 2.0 * nll
