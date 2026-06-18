"""Python reference implementations for univariate analysis.

These functions mirror the Excel LAMBDA formulas in lambda_functions.json and
serve as QC reference values.  Each function accepts a sequence of values that
may contain None/NaN (representing blank Excel cells) and cleans them before
computing.

Skewness and kurtosis use the exact adjustment factors from Excel's SKEW and
KURT functions so that QC comparisons pass to float-comparison precision.

IQR for Freedman-Diaconis bins is computed with numpy's default linear
interpolation (equivalent to Excel QUARTILE.INC).  The LAMBDA uses
QUARTILE.EXC; for small samples the resulting bin count may differ by ±1.
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
    """Count of None/NaN values (blank or non-numeric cells)."""
    clean = _clean(data)
    last_row = len(clean)
    numeric = int(np.sum(np.isfinite(clean)))
    return last_row - numeric


def skewness_excel(x: np.ndarray) -> float:
    """Sample skewness using Excel's SKEW formula."""
    n = len(x)
    if n < 3:
        return float("nan")
    mean = float(np.mean(x))
    s = float(np.std(x, ddof=1))
    if s == 0:
        return float("nan")
    return float(n / ((n - 1) * (n - 2)) * np.sum(((x - mean) / s) ** 3))


def kurtosis_excel(x: np.ndarray) -> float:
    """Excess kurtosis using Excel's KURT formula."""
    n = len(x)
    if n < 4:
        return float("nan")
    mean = float(np.mean(x))
    s = float(np.std(x, ddof=1))
    if s == 0:
        return float("nan")
    fourth = float(np.sum(((x - mean) / s) ** 4))
    return (
        n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * fourth
        - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3))
    )


def descriptive_stats(data: Sequence[float | None]) -> dict[str, float | int]:
    """Compute the 12 descriptive statistics displayed on the Univariate sheet.

    Returns a dict with keys:
        mean, median, mode, sd, variance, min, max, range,
        skewness, kurtosis, count, missing_count
    """
    x = _clean(data)
    n = int(len(x))
    mc = int(len(data)) - n  # cells passed in vs numeric cells

    try:
        mode_val = float(scipy_stats.mode(x, keepdims=False).mode)
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
        "skewness":      skewness_excel(x),
        "kurtosis":      kurtosis_excel(x),
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
    width = 3.49 * sd * n ** (-1.0 / 3.0)
    data_range = float(np.max(x) - np.min(x))
    return math.ceil(data_range / width)


def fd_bins(data: Sequence[float | None]) -> int:
    """Freedman-Diaconis bin count from width 2 * IQR * n^(-1/3)."""
    x = _clean(data)
    n = int(np.sum(np.isfinite(x)))
    q1, q3 = float(np.percentile(x, 25)), float(np.percentile(x, 75))
    iqr = q3 - q1
    if iqr == 0:
        return sturges_bins(data)
    width = 2.0 * iqr * n ** (-1.0 / 3.0)
    data_range = float(np.max(x) - np.min(x))
    return math.ceil(data_range / width)


def bin_edges(data: Sequence[float | None], method: str) -> np.ndarray:
    """Return upper bin edges for the given method (Sturges, Scott, or FD).

    Edges are evenly spaced from min to max, with the count determined by the
    chosen binning rule.  Matches Excel's SEQUENCE-based Bin_Edges LAMBDA.
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
    """Frequency counts matching Excel's FREQUENCY convention.

    For each bin i, counts x where edges[i-1] < x <= edges[i].
    The first bin counts x <= edges[0].
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
# All return positive NLL = -sum(log(pdf(x))).  These match their LAMBDA
# counterparts; the LAMBDAs use the same Excel built-in distribution functions.

def nll_normal(
    data: Sequence[float | None], mean: float, sd: float
) -> float:
    """NLL for Normal(mean, sd)."""
    x = _clean(data)
    return float(-np.sum(scipy_stats.norm.logpdf(x, loc=mean, scale=sd)))


def nll_lognormal(
    data: Sequence[float | None], meanlog: float, sdlog: float
) -> float:
    """NLL for Lognormal with log-space parameters (meanlog, sdlog)."""
    x = _clean(data)
    return float(-np.sum(scipy_stats.lognorm.logpdf(x, s=sdlog, scale=math.exp(meanlog))))


def nll_exponential(data: Sequence[float | None], rate: float) -> float:
    """NLL for Exponential(rate).  rate = 1/mean."""
    x = _clean(data)
    return float(-np.sum(scipy_stats.expon.logpdf(x, scale=1.0 / rate)))


def nll_weibull(data: Sequence[float | None], shape: float, scale: float) -> float:
    """NLL for Weibull(shape, scale)."""
    x = _clean(data)
    return float(-np.sum(scipy_stats.weibull_min.logpdf(x, c=shape, scale=scale)))


def nll_gamma(data: Sequence[float | None], shape: float, rate: float) -> float:
    """NLL for Gamma(shape, rate).  rate = 1/scale."""
    x = _clean(data)
    return float(-np.sum(scipy_stats.gamma.logpdf(x, a=shape, scale=1.0 / rate)))


def nll_triangular(
    data: Sequence[float | None], min_val: float, mode_val: float, max_val: float
) -> float:
    """NLL for Triangular(min, mode, max).

    Returns 1e15 for degenerate parameter sets (mode at either bound) or
    any value outside the support [min, max].
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
    """NLL for Beta(alpha, beta).  Data must be on (0, 1)."""
    x = _clean(data)
    log_pdf = scipy_stats.beta.logpdf(x, a=alpha_param, b=beta_param)
    if not np.all(np.isfinite(log_pdf)):
        return 1e15
    return float(-np.sum(log_pdf))


def nll_betapert(
    data: Sequence[float | None], min_val: float, mode_val: float, max_val: float
) -> float:
    """NLL for BetaPERT(min, mode, max).

    PERT reparameterization: alpha = 1 + 4*(mode - min)/(max - min),
    beta = 1 + 4*(max - mode)/(max - min).  The Jacobian log(1/(max-min))
    shifts the log-likelihood from the scaled [0,1] domain back to the
    original domain.
    """
    x = _clean(data)
    data_range = max_val - min_val
    alpha_p = 1.0 + 4.0 * (mode_val - min_val) / data_range
    beta_p = 1.0 + 4.0 * (max_val - mode_val) / data_range
    z = (x - min_val) / data_range
    log_pdf_z = scipy_stats.beta.logpdf(z, a=alpha_p, b=beta_p)
    log_pdf_x = log_pdf_z - math.log(data_range)
    if not np.all(np.isfinite(log_pdf_x)):
        return 1e15
    return float(-np.sum(log_pdf_x))


# ── Closed-form MLE parameter estimates ──────────────────────────────────────

def mle_normal(data: Sequence[float | None]) -> tuple[float, float]:
    """MLE for Normal: (mean, sd).  Uses sample SD (ddof=1), matching STDEV.S."""
    x = _clean(data)
    return float(np.mean(x)), float(np.std(x, ddof=1))


def mle_lognormal(data: Sequence[float | None]) -> tuple[float, float]:
    """MLE for Lognormal: (meanlog, sdlog) = (mean of ln(x), SD of ln(x))."""
    x = _clean(data)
    log_x = np.log(x)
    return float(np.mean(log_x)), float(np.std(log_x, ddof=1))


def mle_exponential(data: Sequence[float | None]) -> float:
    """MLE for Exponential: rate = 1 / mean."""
    x = _clean(data)
    return float(1.0 / np.mean(x))


def mle_triangular(
    data: Sequence[float | None],
) -> tuple[float, float, float]:
    """Triangular estimation: (min, mode, max).

    True MLE is non-differentiable at the mode; direct min/mode/max estimation
    is used instead.  When no unique mode exists (all values appear once),
    falls back to the median, matching the IFERROR(MODE.SNGL, MEDIAN())
    formula in the Univariate sheet.
    """
    x = _clean(data)
    min_val = float(np.min(x))
    max_val = float(np.max(x))
    try:
        result = scipy_stats.mode(x, keepdims=False)
        # scipy mode returns the minimum value when counts are all 1;
        # treat that as "no unique mode" and fall back to median.
        if result.count == 1 and np.sum(x == result.mode) == 1:
            mode_val = float(np.median(x))
        else:
            mode_val = float(result.mode)
    except Exception:
        mode_val = float(np.median(x))
    # Ensure mode is strictly interior to [min, max]
    if mode_val <= min_val or mode_val >= max_val:
        mode_val = float(np.median(x))
    if mode_val <= min_val or mode_val >= max_val:
        mode_val = (min_val + max_val) / 2.0
    return min_val, mode_val, max_val


def mle_betapert(
    data: Sequence[float | None],
) -> tuple[float, float, float]:
    """BetaPERT estimation: (min, mode, max), same as Triangular."""
    return mle_triangular(data)


# ── Goodness-of-fit statistics ────────────────────────────────────────────────

def gof_aic(nll: float, k: int) -> float:
    """AIC = 2k + 2*NLL."""
    return 2.0 * k + 2.0 * nll


def gof_bic(nll: float, k: int, n: int) -> float:
    """BIC = k*ln(n) + 2*NLL."""
    return k * math.log(n) + 2.0 * nll
