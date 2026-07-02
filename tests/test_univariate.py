# pylint: disable=import-outside-toplevel,missing-class-docstring,missing-function-docstring
# pylint: disable=missing-module-docstring,reimported
from __future__ import annotations

import math
import unittest

import numpy as np

from scipy import stats as scipy_stats

from lambda_catalog.analyze_univariate import (
    bin_counts,
    bin_edges,
    bin_lower_edges,
    bin_midpoints,
    cdf_beta,
    cdf_betapert,
    cdf_exponential,
    cdf_gamma,
    cdf_lognormal,
    cdf_normal,
    cdf_triangular,
    cdf_weibull,
    descriptive_stats,
    fd_bins,
    gof_aic,
    gof_anderson_darling,
    gof_bic,
    gof_ks,
    missing_count,
    mle_exponential,
    mle_gamma,
    mle_lognormal,
    mle_normal,
    mle_triangular,
    mle_weibull,
    nll_beta,
    nll_betapert,
    nll_exponential,
    nll_gamma,
    nll_lognormal,
    nll_normal,
    nll_triangular,
    nll_weibull,
    scott_bins,
    sturges_bins,
)


# ── helpers ───────────────────────────────────────────────────────────────────

def _uniform(n: int = 50, seed: int = 42) -> list[float]:
    rng = np.random.default_rng(seed)
    return rng.uniform(1.0, 10.0, n).tolist()


def _normal_data(n: int = 100, mu: float = 5.0, sigma: float = 2.0, seed: int = 0) -> list[float]:
    rng = np.random.default_rng(seed)
    return rng.normal(mu, sigma, n).tolist()


# ── missing_count ─────────────────────────────────────────────────────────────

class MissingCountTests(unittest.TestCase):
    def test_no_missing(self) -> None:
        self.assertEqual(missing_count([1.0, 2.0, 3.0]), 0)

    def test_trailing_nones_ignored(self) -> None:
        # None after the last numeric value should not be counted
        self.assertEqual(missing_count([1.0, 2.0, None, None]), 0)

    def test_interior_nones_counted(self) -> None:
        self.assertEqual(missing_count([1.0, None, 3.0]), 1)

    def test_all_none(self) -> None:
        self.assertEqual(missing_count([None, None]), 0)

    def test_empty(self) -> None:
        self.assertEqual(missing_count([]), 0)

    def test_multiple_interior_gaps(self) -> None:
        data = [1.0, None, 3.0, None, 5.0]
        self.assertEqual(missing_count(data), 2)

    def test_text_value_counts_as_missing(self) -> None:
        data = [1.0, "not numeric", 3.0]
        self.assertEqual(missing_count(data), 1)

    def test_trailing_empty_strings_ignored(self) -> None:
        self.assertEqual(missing_count([1.0, 2.0, "", ""]), 0)

    def test_interior_empty_string_counted(self) -> None:
        self.assertEqual(missing_count([1.0, "", 3.0]), 1)

    def test_all_empty_strings(self) -> None:
        self.assertEqual(missing_count(["", ""]), 0)


# ── descriptive_stats ─────────────────────────────────────────────────────────

class DescriptiveStatsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.data = [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0]

    def test_keys_present(self) -> None:
        result = descriptive_stats(self.data)
        expected_keys = {
            "mean", "median", "mode", "sd", "variance",
            "min", "max", "range", "skewness", "kurtosis",
            "count", "missing_count",
        }
        self.assertEqual(set(result.keys()), expected_keys)

    def test_mean(self) -> None:
        result = descriptive_stats(self.data)
        self.assertAlmostEqual(float(result["mean"]), float(np.mean(self.data)), places=10)

    def test_sd_uses_ddof1(self) -> None:
        result = descriptive_stats(self.data)
        self.assertAlmostEqual(float(result["sd"]), float(np.std(self.data, ddof=1)), places=10)

    def test_mode_with_unique_mode(self) -> None:
        result = descriptive_stats(self.data)
        self.assertAlmostEqual(result["mode"], 4.0, places=10)

    def test_mode_nan_when_all_unique(self) -> None:
        data = [1.0, 2.0, 3.0, 4.0, 5.0]
        result = descriptive_stats(data)
        self.assertTrue(math.isnan(result["mode"]))

    def test_count_excludes_nones(self) -> None:
        data = [1.0, None, 2.0, 3.0]
        result = descriptive_stats(data)
        self.assertEqual(result["count"], 3)

    def test_missing_count_in_output(self) -> None:
        data = [1.0, None, 3.0, None, 5.0]
        result = descriptive_stats(data)
        self.assertEqual(result["missing_count"], 2)


# ── binning rules ─────────────────────────────────────────────────────────────

class BinningTests(unittest.TestCase):
    def setUp(self) -> None:
        self.data = _uniform()  # 50 values

    def test_sturges_formula(self) -> None:
        n = len(self.data)
        expected = math.ceil(math.log2(n) + 1)
        self.assertEqual(sturges_bins(self.data), expected)

    def test_scott_bins_positive(self) -> None:
        self.assertGreater(scott_bins(self.data), 0)

    def test_fd_bins_positive(self) -> None:
        self.assertGreater(fd_bins(self.data), 0)

    def test_scott_constant_data_returns_1(self) -> None:
        self.assertEqual(scott_bins([3.0] * 20), 1)

    def test_fd_constant_data_falls_back_to_sturges(self) -> None:
        data = [5.0] * 20
        self.assertEqual(fd_bins(data), sturges_bins(data))

    def test_bin_edges_length_equals_bin_count(self) -> None:
        for method in ("Sturges", "Scott", "FD"):
            edges = bin_edges(self.data, method)
            if method == "Sturges":
                expected_k = sturges_bins(self.data)
            elif method == "Scott":
                expected_k = scott_bins(self.data)
            else:
                expected_k = fd_bins(self.data)
            self.assertEqual(len(edges), expected_k, f"method={method}")

    def test_bin_edges_span_data_range(self) -> None:
        edges = bin_edges(self.data, "Sturges")
        self.assertAlmostEqual(edges[-1], max(self.data), places=10)

    def test_bin_counts_sum_to_n(self) -> None:
        edges = bin_edges(self.data, "Sturges")
        counts = bin_counts(self.data, edges)
        self.assertEqual(int(counts.sum()), len(self.data))

    def test_bin_edges_unknown_method_raises(self) -> None:
        with self.assertRaises(ValueError):
            bin_edges(self.data, "Unknown")


# ── NLL functions ─────────────────────────────────────────────────────────────

class NLLTests(unittest.TestCase):
    def setUp(self) -> None:
        rng = np.random.default_rng(7)
        self.norm_data = rng.normal(5.0, 2.0, 80).tolist()
        self.pos_data = rng.exponential(2.0, 80).tolist()

    def test_nll_normal_finite_at_mle(self) -> None:
        mu, sigma = mle_normal(self.norm_data)
        nll = nll_normal(self.norm_data, mu, sigma)
        self.assertTrue(math.isfinite(nll))

    def test_nll_normal_minimised_at_mle(self) -> None:
        mu, sigma = mle_normal(self.norm_data)
        nll_mle = nll_normal(self.norm_data, mu, sigma)
        nll_perturbed = nll_normal(self.norm_data, mu + 0.5, sigma)
        self.assertLess(nll_mle, nll_perturbed)

    def test_nll_lognormal_finite_at_mle(self) -> None:
        mu, sigma = mle_lognormal(self.pos_data)
        nll = nll_lognormal(self.pos_data, mu, sigma)
        self.assertTrue(math.isfinite(nll))

    def test_nll_exponential_finite_at_mle(self) -> None:
        rate = mle_exponential(self.pos_data)
        nll = nll_exponential(self.pos_data, rate)
        self.assertTrue(math.isfinite(nll))

    def test_nll_weibull_finite_at_mle(self) -> None:
        shape, scale = mle_weibull(self.pos_data)
        nll = nll_weibull(self.pos_data, shape, scale)
        self.assertTrue(math.isfinite(nll))

    def test_nll_gamma_finite_at_mle(self) -> None:
        shape, rate = mle_gamma(self.pos_data)
        nll = nll_gamma(self.pos_data, shape, rate)
        self.assertTrue(math.isfinite(nll))

    def test_nll_weibull_sentinel_for_zero_scale(self) -> None:
        nll = nll_weibull(self.pos_data, 1.0, 0.0)
        self.assertGreaterEqual(nll, 1e14)

    def test_nll_gamma_sentinel_for_zero_rate(self) -> None:
        nll = nll_gamma(self.pos_data, 1.0, 0.0)
        self.assertGreaterEqual(nll, 1e14)

    def test_nll_triangular_degenerate_mode_at_min(self) -> None:
        data = [1.0, 2.0, 3.0]
        nll = nll_triangular(data, 1.0, 1.0, 3.0)
        self.assertAlmostEqual(nll, 1e15)

    def test_nll_triangular_out_of_support(self) -> None:
        data = [0.5, 1.5, 2.5]
        nll = nll_triangular(data, 1.0, 1.5, 3.0)
        self.assertAlmostEqual(nll, 1e15)

    def test_nll_triangular_finite_for_valid_params(self) -> None:
        data = [1.5, 2.0, 2.5, 3.0, 3.5]
        nll = nll_triangular(data, 1.0, 2.5, 4.0)
        self.assertTrue(math.isfinite(nll))
        self.assertLess(nll, 1e14)

    def test_nll_betapert_consistent_with_beta(self) -> None:
        # BetaPERT and Beta NLL should both be finite for valid params
        data = [1.5, 2.0, 2.5, 3.0, 3.5]
        nll_bp = nll_betapert(data, 1.0, 2.5, 4.0)
        self.assertTrue(math.isfinite(nll_bp))
        self.assertLess(nll_bp, 1e14)

    def test_nll_beta_sentinel_for_boundary_values(self) -> None:
        # 0.0 is on the boundary of (0,1); logpdf = -inf
        data = [0.0, 0.5, 1.0]
        nll = nll_beta(data, 2.0, 3.0)
        self.assertAlmostEqual(nll, 1e15)


# ── MLE functions ─────────────────────────────────────────────────────────────

class MLETests(unittest.TestCase):
    def setUp(self) -> None:
        rng = np.random.default_rng(13)
        self.norm_data = rng.normal(10.0, 3.0, 200).tolist()
        self.pos_data = rng.exponential(4.0, 200).tolist()

    def test_mle_normal_mean_close(self) -> None:
        mu, _ = mle_normal(self.norm_data)
        self.assertAlmostEqual(mu, np.mean(self.norm_data), places=10)

    def test_mle_normal_sd_close(self) -> None:
        # scipy uses ddof=0 (true MLE); tolerance vs ddof=1 is ~1/(2n)
        _, sigma = mle_normal(self.norm_data)
        self.assertAlmostEqual(sigma, np.std(self.norm_data, ddof=0), places=10)

    def test_mle_exponential_rate_close(self) -> None:
        rate = mle_exponential(self.pos_data)
        self.assertAlmostEqual(rate, 1.0 / np.mean(self.pos_data), places=10)

    def test_mle_lognormal_returns_finite(self) -> None:
        mu, sigma = mle_lognormal(self.pos_data)
        self.assertTrue(math.isfinite(mu))
        self.assertTrue(math.isfinite(sigma))
        self.assertGreater(sigma, 0.0)

    def test_mle_weibull_returns_positive_params(self) -> None:
        shape, scale = mle_weibull(self.pos_data)
        self.assertGreater(shape, 0.0)
        self.assertGreater(scale, 0.0)

    def test_mle_gamma_returns_positive_params(self) -> None:
        shape, rate = mle_gamma(self.pos_data)
        self.assertGreater(shape, 0.0)
        self.assertGreater(rate, 0.0)

    def test_mle_triangular_mode_strictly_interior(self) -> None:
        data = [1.0, 2.0, 3.0, 3.0, 3.0, 4.0, 5.0]
        min_val, mode_val, max_val = mle_triangular(data)
        self.assertGreater(mode_val, min_val)
        self.assertLess(mode_val, max_val)

    def test_mle_triangular_all_unique_falls_back_to_median(self) -> None:
        data = [1.0, 2.0, 3.0, 4.0, 5.0]
        min_val, mode_val, max_val = mle_triangular(data)
        self.assertAlmostEqual(min_val, 1.0)
        self.assertAlmostEqual(max_val, 5.0)
        self.assertGreater(mode_val, min_val)
        self.assertLess(mode_val, max_val)


# ── GoF metrics ───────────────────────────────────────────────────────────────

class GoFTests(unittest.TestCase):
    def test_aic_formula(self) -> None:
        self.assertAlmostEqual(gof_aic(nll=50.0, k=2), 2 * 2 + 2 * 50.0)

    def test_bic_formula(self) -> None:
        import math as _math
        self.assertAlmostEqual(gof_bic(nll=50.0, k=2, n=100), 2 * _math.log(100) + 2 * 50.0)

    def test_aic_lower_for_better_fit(self) -> None:
        # Lower NLL → lower AIC
        self.assertLess(gof_aic(30.0, 2), gof_aic(50.0, 2))

    def test_bic_penalises_more_params(self) -> None:
        # Same NLL, more parameters → higher BIC
        self.assertLess(gof_bic(50.0, 2, 100), gof_bic(50.0, 4, 100))

    # ── Anderson-Darling ──

    def test_ad_finite_for_good_fit(self) -> None:
        data = _normal_data(200)
        mu, sd = float(np.mean(data)), float(np.std(data, ddof=1))
        cdf = [float(scipy_stats.norm.cdf(x, mu, sd)) for x in data]
        ad = gof_anderson_darling(data, cdf)
        self.assertTrue(math.isfinite(ad))

    def test_ad_worse_for_bad_fit(self) -> None:
        data = _normal_data(200)
        mu, sd = float(np.mean(data)), float(np.std(data, ddof=1))
        good_cdf = [float(scipy_stats.norm.cdf(x, mu, sd)) for x in data]
        bad_cdf = [float(scipy_stats.norm.cdf(x, 0.0, 1.0)) for x in data]
        self.assertLess(
            gof_anderson_darling(data, good_cdf),
            gof_anderson_darling(data, bad_cdf),
        )

    def test_ad_matches_scipy(self) -> None:
        # scipy.stats.anderson standardizes with ddof=1
        data = _normal_data(200)
        mu = float(np.mean(data))
        sd = float(np.std(data, ddof=1))
        cdf = [float(scipy_stats.norm.cdf(x, mu, sd)) for x in data]
        ad = gof_anderson_darling(data, cdf)
        scipy_ad = scipy_stats.anderson(data, dist="norm").statistic
        self.assertAlmostEqual(ad, scipy_ad, places=6)

    def test_ad_bounded_support_no_crash(self) -> None:
        data = [0.0, 0.25, 0.5, 0.75, 1.0]
        cdf = [0.0, 0.25, 0.5, 0.75, 1.0]
        ad = gof_anderson_darling(data, cdf)
        self.assertTrue(math.isfinite(ad))

    # ── Kolmogorov-Smirnov ──

    def test_ks_perfect_empirical_cdf(self) -> None:
        # When F(x_i) = i/n, D⁺ = 0 but D⁻ = 1/n (empirical CDF step gap)
        n = 100
        data = list(range(1, n + 1))
        cdf = [i / n for i in range(1, n + 1)]
        ks = gof_ks(data, cdf)
        self.assertAlmostEqual(ks, 1.0 / n, places=10)

    def test_ks_bounded_01(self) -> None:
        data = _normal_data(100)
        mu, sd = float(np.mean(data)), float(np.std(data, ddof=1))
        cdf = [float(scipy_stats.norm.cdf(x, mu, sd)) for x in data]
        ks = gof_ks(data, cdf)
        self.assertGreaterEqual(ks, 0.0)
        self.assertLessEqual(ks, 1.0)

    def test_ks_worse_for_bad_fit(self) -> None:
        data = _normal_data(200)
        mu, sd = float(np.mean(data)), float(np.std(data, ddof=1))
        good_cdf = [float(scipy_stats.norm.cdf(x, mu, sd)) for x in data]
        bad_cdf = [float(scipy_stats.norm.cdf(x, 0.0, 1.0)) for x in data]
        self.assertLess(gof_ks(data, good_cdf), gof_ks(data, bad_cdf))

    def test_ks_matches_scipy(self) -> None:
        data = _normal_data(200)
        mu, sd = float(np.mean(data)), float(np.std(data, ddof=1))
        scipy_ks = scipy_stats.kstest(data, "norm", args=(mu, sd)).statistic
        cdf = [float(scipy_stats.norm.cdf(x, mu, sd)) for x in data]
        ks = gof_ks(data, cdf)
        self.assertAlmostEqual(ks, scipy_ks, places=10)


# ── bin_midpoints ────────────────────────────────────────────────────────────

class BinMidpointsTests(unittest.TestCase):
    def test_midpoints_lie_between_edges(self) -> None:
        data = _uniform()
        edges = bin_edges(data, "Sturges")
        mids = bin_midpoints(data, edges)
        self.assertEqual(len(mids), len(edges))
        width = (max(data) - min(data)) / len(edges)
        for mid, edge in zip(mids, edges):
            self.assertAlmostEqual(mid, edge - width / 2.0, places=10)

    def test_single_bin(self) -> None:
        data = [1.0, 2.0, 3.0, 4.0, 5.0]
        edges = np.array([5.0])
        mids = bin_midpoints(data, edges)
        self.assertEqual(len(mids), 1)
        self.assertAlmostEqual(float(mids[0]), 5.0 - (5.0 - 1.0) / 2.0, places=10)


# ── bin_lower_edges ──────────────────────────────────────────────────────────

class BinLowerEdgesTests(unittest.TestCase):
    def test_first_lower_edge_is_data_min(self) -> None:
        data = _uniform()
        edges = bin_edges(data, "Sturges")
        lowers = bin_lower_edges(data, edges)
        self.assertAlmostEqual(float(lowers[0]), min(data), places=10)

    def test_subsequent_lower_edges_are_previous_upper(self) -> None:
        data = _uniform()
        edges = bin_edges(data, "Sturges")
        lowers = bin_lower_edges(data, edges)
        for i in range(1, len(edges)):
            self.assertAlmostEqual(float(lowers[i]), float(edges[i - 1]), places=10)

    def test_length_matches_edges(self) -> None:
        data = _uniform()
        edges = bin_edges(data, "Scott")
        lowers = bin_lower_edges(data, edges)
        self.assertEqual(len(lowers), len(edges))

    def test_single_bin(self) -> None:
        data = [3.0, 3.0, 3.0, 5.0]
        edges = np.array([5.0])
        lowers = bin_lower_edges(data, edges)
        self.assertEqual(len(lowers), 1)
        self.assertAlmostEqual(float(lowers[0]), 3.0, places=10)


# ── CDF functions ────────────────────────────────────────────────────────────

class CDFTests(unittest.TestCase):
    """Validate CDF bin-probability functions against scipy.stats."""

    def test_cdf_normal_matches_scipy(self) -> None:
        mu, sd = 5.0, 2.0
        result = cdf_normal(10.0, mu, sd)
        expected = float(scipy_stats.norm.cdf(10.0, mu, sd))
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_normal_interval(self) -> None:
        mu, sd = 5.0, 2.0
        result = cdf_normal(7.0, mu, sd, minimum=3.0)
        expected = float(scipy_stats.norm.cdf(7.0, mu, sd) - scipy_stats.norm.cdf(3.0, mu, sd))
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_lognormal_matches_scipy(self) -> None:
        meanlog, sdlog = 1.5, 0.5
        result = cdf_lognormal(10.0, meanlog, sdlog)
        expected = float(scipy_stats.lognorm.cdf(10.0, s=sdlog, scale=np.exp(meanlog)))
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_lognormal_interval(self) -> None:
        meanlog, sdlog = 1.5, 0.5
        result = cdf_lognormal(8.0, meanlog, sdlog, minimum=2.0)
        expected = float(
            scipy_stats.lognorm.cdf(8.0, s=sdlog, scale=np.exp(meanlog))
            - scipy_stats.lognorm.cdf(2.0, s=sdlog, scale=np.exp(meanlog))
        )
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_exponential_matches_scipy(self) -> None:
        rate = 0.5
        result = cdf_exponential(4.0, rate)
        expected = float(scipy_stats.expon.cdf(4.0, scale=1.0 / rate))
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_exponential_interval(self) -> None:
        rate = 0.5
        result = cdf_exponential(4.0, rate, minimum=1.0)
        expected = float(
            scipy_stats.expon.cdf(4.0, scale=1.0 / rate)
            - scipy_stats.expon.cdf(1.0, scale=1.0 / rate)
        )
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_weibull_matches_scipy(self) -> None:
        shape, scale = 2.0, 5.0
        result = cdf_weibull(6.0, shape, scale)
        expected = float(scipy_stats.weibull_min.cdf(6.0, shape, scale=scale))
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_weibull_interval(self) -> None:
        shape, scale = 2.0, 5.0
        result = cdf_weibull(8.0, shape, scale, minimum=2.0)
        expected = float(
            scipy_stats.weibull_min.cdf(8.0, shape, scale=scale)
            - scipy_stats.weibull_min.cdf(2.0, shape, scale=scale)
        )
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_gamma_matches_scipy(self) -> None:
        shape, rate = 3.0, 0.5
        result = cdf_gamma(6.0, shape, rate)
        expected = float(scipy_stats.gamma.cdf(6.0, shape, scale=1.0 / rate))
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_gamma_interval(self) -> None:
        shape, rate = 3.0, 0.5
        result = cdf_gamma(10.0, shape, rate, minimum=2.0)
        expected = float(
            scipy_stats.gamma.cdf(10.0, shape, scale=1.0 / rate)
            - scipy_stats.gamma.cdf(2.0, shape, scale=1.0 / rate)
        )
        self.assertAlmostEqual(result, expected, places=10)

    def test_cdf_triangular_matches_scipy(self) -> None:
        a, c, b = 1.0, 4.0, 9.0
        loc, scale = a, b - a
        shape = (c - a) / (b - a)
        result = cdf_triangular(6.0, a, c, b)
        expected = float(scipy_stats.triang.cdf(6.0, shape, loc=loc, scale=scale))
        self.assertAlmostEqual(result, expected, places=8)

    def test_cdf_triangular_interval(self) -> None:
        a, c, b = 1.0, 4.0, 9.0
        loc, scale = a, b - a
        shape = (c - a) / (b - a)
        result = cdf_triangular(7.0, a, c, b, minimum=3.0)
        expected = float(
            scipy_stats.triang.cdf(7.0, shape, loc=loc, scale=scale)
            - scipy_stats.triang.cdf(3.0, shape, loc=loc, scale=scale)
        )
        self.assertAlmostEqual(result, expected, places=8)

    def test_cdf_beta_matches_scipy(self) -> None:
        data = _normal_data(100, mu=5.0, sigma=1.5)
        alpha, beta_p = 2.0, 5.0
        data_min = min(data)
        data_range = max(data) - min(data)
        pad = data_range * 0.001
        scale_ = data_range + 2 * pad
        x = 6.0
        z = (x - data_min + pad) / scale_
        result = cdf_beta(x, alpha, beta_p, data_min, data_range)
        expected = float(scipy_stats.beta.cdf(z, alpha, beta_p))
        self.assertAlmostEqual(result, expected, places=8)

    def test_cdf_beta_interval(self) -> None:
        data = _normal_data(100, mu=5.0, sigma=1.5)
        alpha, beta_p = 2.0, 5.0
        data_min = min(data)
        data_range = max(data) - min(data)
        pad = data_range * 0.001
        scale_ = data_range + 2 * pad
        x_lo, x_hi = 4.0, 7.0
        z_lo = (x_lo - data_min + pad) / scale_
        z_hi = (x_hi - data_min + pad) / scale_
        result = cdf_beta(x_hi, alpha, beta_p, data_min, data_range, minimum=x_lo)
        expected = float(
            scipy_stats.beta.cdf(z_hi, alpha, beta_p)
            - scipy_stats.beta.cdf(z_lo, alpha, beta_p)
        )
        self.assertAlmostEqual(result, expected, places=8)

    def test_cdf_betapert_matches_manual(self) -> None:
        mn, md, mx = 1.0, 4.0, 9.0
        mu = (mn + 4 * md + mx) / 6.0
        a1 = (mu - mn) * (2 * md - mn - mx) / ((md - mu) * (mx - mn))
        a2 = a1 * (mx - mu) / (mu - mn)
        x = 5.0
        z = (x - mn) / (mx - mn)
        result = cdf_betapert(x, mn, md, mx)
        expected = float(scipy_stats.beta.cdf(z, a1, a2))
        self.assertAlmostEqual(result, expected, places=8)

    def test_cdf_betapert_interval(self) -> None:
        mn, md, mx = 1.0, 4.0, 9.0
        mu = (mn + 4 * md + mx) / 6.0
        a1 = (mu - mn) * (2 * md - mn - mx) / ((md - mu) * (mx - mn))
        a2 = a1 * (mx - mu) / (mu - mn)
        x_lo, x_hi = 3.0, 7.0
        z_lo = (x_lo - mn) / (mx - mn)
        z_hi = (x_hi - mn) / (mx - mn)
        result = cdf_betapert(x_hi, mn, md, mx, minimum=x_lo)
        expected = float(
            scipy_stats.beta.cdf(z_hi, a1, a2)
            - scipy_stats.beta.cdf(z_lo, a1, a2)
        )
        self.assertAlmostEqual(result, expected, places=8)

    def test_cdf_probabilities_sum_approximately_one(self) -> None:
        """Bin probabilities from a CDF should sum close to 1 over all bins."""
        data = _normal_data(200)
        mu, sd = float(np.mean(data)), float(np.std(data, ddof=1))
        edges = bin_edges(data, "Sturges")
        lowers = bin_lower_edges(data, edges)
        total = sum(
            cdf_normal(float(hi), mu, sd, minimum=float(lo))
            for hi, lo in zip(edges, lowers)
        )
        self.assertAlmostEqual(total, 1.0, places=1)


if __name__ == "__main__":
    unittest.main()
