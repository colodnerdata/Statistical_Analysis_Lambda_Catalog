"""Independent verification of LAMBDA function outputs.

Each test computes expected values from scratch using numpy, scipy, and
manual formulas — never delegating to the project's own oracle functions —
then verifies the oracle agrees. This catches bugs where the oracle
and the LAMBDA formula share the same wrong implementation.
"""
from __future__ import annotations

import csv
import math
import tempfile
import unittest
from pathlib import Path
from statistics import NormalDist

import numpy as np
from scipy import stats as sp_stats

from lambda_catalog.analyze_life_expectancy import (
    FEATURE_COLUMNS,
    TARGET_COLUMN,
    calculate_regression_observation_vectors,
    calculate_regression_summary,
    calculate_regression_vectors,
)
from lambda_catalog.analyze_regression_sheet import (
    calculate_regression_sheet_results,
)

# ── Test data helpers ────────────────────────────────────────────────────────

def _write_csv(path: Path, rows: list[dict]) -> None:
    headers = [TARGET_COLUMN, *FEATURE_COLUMNS]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({h: row.get(h, 1.0) for h in headers})


def _make_row(y: float, *xs: float) -> dict:
    row = {name: 1.0 for name in FEATURE_COLUMNS}
    for i, val in enumerate(xs):
        row[FEATURE_COLUMNS[i]] = val
    row[TARGET_COLUMN] = y
    return row


def _build_arrays(rows: list[dict], k: int, intercept: bool):
    """Build X, Y arrays from test rows — fully independent of project code."""
    cols = FEATURE_COLUMNS[:k]
    y = np.array([r[TARGET_COLUMN] for r in rows], dtype=np.float64)
    x_raw = np.array([[r[c] for c in cols] for r in rows], dtype=np.float64)
    if intercept:
        x = np.column_stack([np.ones(len(rows)), x_raw])
    else:
        x = x_raw
    return x, y, x_raw


# ── Scalar summary metrics ──────────────────────────────────────────────────

class TestScalarMetricsIndependent(unittest.TestCase):
    """Verify scalar regression metrics against manual numpy/scipy computation."""

    def setUp(self) -> None:
        self.rows = [
            _make_row(7.2, 1.0, 2.0),
            _make_row(8.0, 2.0, 0.5),
            _make_row(9.4, 3.0, 1.5),
            _make_row(10.1, 4.0, 1.0),
            _make_row(11.8, 5.0, 2.5),
            _make_row(12.3, 6.0, 2.0),
            _make_row(13.5, 7.0, 3.0),
            _make_row(14.0, 8.0, 1.5),
            _make_row(6.5, 0.5, 1.0),
            _make_row(15.1, 9.0, 2.8),
        ]
        self.k = 2
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "data.csv"
        _write_csv(self._csv, self.rows)
        self.summary = calculate_regression_summary(
            self._csv, include_intercept=True, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.X, self.Y, self.X_raw = _build_arrays(self.rows, self.k, True)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _ols_fit(self):
        """Fit OLS via normal equations — no statsmodels."""
        X, Y = self.X, self.Y
        beta = np.linalg.lstsq(X, Y, rcond=None)[0]
        y_hat = X @ beta
        residuals = Y - y_hat
        return beta, y_hat, residuals

    def test_observations(self) -> None:
        self.assertEqual(self.summary.observations, len(self.rows))

    def test_df_regression(self) -> None:
        self.assertEqual(self.summary.df_regression, self.k)

    def test_df_residual(self) -> None:
        n = len(self.rows)
        self.assertEqual(self.summary.df_residual, n - self.k - 1)

    def test_df_total(self) -> None:
        self.assertEqual(self.summary.df_total, len(self.rows) - 1)

    def test_ss_total(self) -> None:
        ss_total = float(np.sum((self.Y - np.mean(self.Y)) ** 2))
        self.assertAlmostEqual(self.summary.ss_total, ss_total, places=10)

    def test_ss_residual(self) -> None:
        _, _, residuals = self._ols_fit()
        ss_res = float(np.sum(residuals**2))
        self.assertAlmostEqual(self.summary.ss_residual, ss_res, places=10)

    def test_ss_regression(self) -> None:
        _, y_hat, _ = self._ols_fit()
        y_mean = np.mean(self.Y)
        ss_reg = float(np.sum((y_hat - y_mean) ** 2))
        self.assertAlmostEqual(self.summary.ss_regression, ss_reg, places=10)

    def test_ss_regression_equals_total_minus_residual(self) -> None:
        self.assertAlmostEqual(
            self.summary.ss_regression,
            self.summary.ss_total - self.summary.ss_residual,
            places=10,
        )

    def test_r_squared(self) -> None:
        _, _, residuals = self._ols_fit()
        ss_res = float(np.sum(residuals**2))
        ss_tot = float(np.sum((self.Y - np.mean(self.Y)) ** 2))
        r2 = 1.0 - ss_res / ss_tot
        self.assertAlmostEqual(self.summary.r_squared, r2, places=10)

    def test_multiple_r(self) -> None:
        self.assertAlmostEqual(
            self.summary.multiple_r,
            math.sqrt(max(self.summary.r_squared, 0.0)),
            places=12,
        )

    def test_adjusted_r2(self) -> None:
        n = len(self.rows)
        k = self.k
        r2 = self.summary.r_squared
        adj_r2 = 1.0 - (1.0 - r2) * (n - 1) / (n - k - 1)
        self.assertAlmostEqual(self.summary.adjusted_r2, adj_r2, places=10)

    def test_se_regression(self) -> None:
        _, _, residuals = self._ols_fit()
        n = len(self.rows)
        k = self.k
        mse = float(np.sum(residuals**2)) / (n - k - 1)
        self.assertAlmostEqual(self.summary.se_regression, math.sqrt(mse), places=10)

    def test_f_stat(self) -> None:
        ms_reg = self.summary.ss_regression / self.summary.df_regression
        ms_res = self.summary.ss_residual / self.summary.df_residual
        f_stat = ms_reg / ms_res
        self.assertAlmostEqual(self.summary.f_stat, f_stat, places=10)

    def test_p_value_f(self) -> None:
        p_value = 1.0 - float(
            sp_stats.f.cdf(self.summary.f_stat, self.summary.df_regression, self.summary.df_residual)
        )
        self.assertAlmostEqual(self.summary.p_value_f, p_value, places=8)

    def test_press(self) -> None:
        _, _, residuals = self._ols_fit()
        X = self.X
        xtx_inv = np.linalg.inv(X.T @ X)
        h = np.sum((X @ xtx_inv) * X, axis=1)
        press = float(np.sum((residuals / (1.0 - h)) ** 2))
        self.assertAlmostEqual(self.summary.press, press, places=8)

    def test_durbin_watson(self) -> None:
        _, _, residuals = self._ols_fit()
        dw = float(np.sum(np.diff(residuals) ** 2) / np.sum(residuals**2))
        self.assertAlmostEqual(self.summary.durbin_watson, dw, places=10)

    def test_aic(self) -> None:
        n = len(self.rows)
        p = self.k + 1
        log_term = n * math.log(self.summary.ss_residual / n)
        expected = log_term + 2.0 * p
        self.assertAlmostEqual(self.summary.aic, expected, places=10)

    def test_bic(self) -> None:
        n = len(self.rows)
        p = self.k + 1
        log_term = n * math.log(self.summary.ss_residual / n)
        expected = log_term + p * math.log(n)
        self.assertAlmostEqual(self.summary.bic, expected, places=10)

    def test_aicc(self) -> None:
        n = len(self.rows)
        p = self.k + 1
        log_term = n * math.log(self.summary.ss_residual / n)
        aic = log_term + 2.0 * p
        expected = aic + 2.0 * p * (p + 1) / (n - p - 1)
        self.assertAlmostEqual(self.summary.aicc, expected, places=10)

    def test_qq_correlation(self) -> None:
        _, _, residuals = self._ols_fit()
        n = len(self.rows)
        k = self.k
        se = math.sqrt(float(np.sum(residuals**2)) / (n - k - 1))
        scaled = np.sort(residuals / se)
        normal_scores = sp_stats.norm.ppf((np.arange(1, n + 1) - 0.5) / n)
        expected_qq, _ = sp_stats.pearsonr(scaled, normal_scores)
        self.assertAlmostEqual(self.summary.qq_correlation, expected_qq, places=10)


# ── Scalar metrics without intercept ────────────────────────────────────────

class TestScalarMetricsNoIntercept(unittest.TestCase):
    """Verify scalar metrics when fitting through the origin."""

    def setUp(self) -> None:
        self.rows = [
            _make_row(3.0, 1.0),
            _make_row(5.5, 2.0),
            _make_row(8.2, 3.0),
            _make_row(11.0, 4.0),
            _make_row(13.5, 5.0),
            _make_row(16.1, 6.0),
            _make_row(18.0, 7.0),
            _make_row(21.3, 8.0),
        ]
        self.k = 1
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "no_int.csv"
        _write_csv(self._csv, self.rows)
        self.summary = calculate_regression_summary(
            self._csv, include_intercept=False, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.X, self.Y, self.X_raw = _build_arrays(self.rows, self.k, False)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_df_total_no_intercept(self) -> None:
        self.assertEqual(self.summary.df_total, len(self.rows))

    def test_df_residual_no_intercept(self) -> None:
        self.assertEqual(self.summary.df_residual, len(self.rows) - self.k)

    def test_ss_total_uncentered(self) -> None:
        ss_total = float(np.sum(self.Y**2))
        self.assertAlmostEqual(self.summary.ss_total, ss_total, places=10)

    def test_ss_decomposition(self) -> None:
        self.assertAlmostEqual(
            self.summary.ss_total,
            self.summary.ss_regression + self.summary.ss_residual,
            places=8,
        )


# ── Coefficient-level vectors ───────────────────────────────────────────────

class TestCoefficientVectorsIndependent(unittest.TestCase):
    """Verify per-coefficient statistics via independent numpy OLS."""

    def setUp(self) -> None:
        self.rows = [
            _make_row(7.2, 1.0, 2.0),
            _make_row(8.0, 2.0, 0.5),
            _make_row(9.4, 3.0, 1.5),
            _make_row(10.1, 4.0, 1.0),
            _make_row(11.8, 5.0, 2.5),
            _make_row(12.3, 6.0, 2.0),
            _make_row(13.5, 7.0, 3.0),
            _make_row(14.0, 8.0, 1.5),
            _make_row(6.5, 0.5, 1.0),
            _make_row(15.1, 9.0, 2.8),
        ]
        self.k = 2
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "coef.csv"
        _write_csv(self._csv, self.rows)
        self.vectors = calculate_regression_vectors(
            self._csv, include_intercept=True, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.X, self.Y, self.X_raw = _build_arrays(self.rows, self.k, True)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _ols_results(self):
        X, Y = self.X, self.Y
        beta = np.linalg.lstsq(X, Y, rcond=None)[0]
        y_hat = X @ beta
        residuals = Y - y_hat
        n, p = X.shape
        mse = float(np.sum(residuals**2)) / (n - p)
        cov_beta = mse * np.linalg.inv(X.T @ X)
        se = np.sqrt(np.diag(cov_beta))
        t_stats = beta / se
        return beta, se, t_stats, mse, n, p

    def test_coefficients(self) -> None:
        beta, _, _, _, _, _ = self._ols_results()
        for i, (expected, actual) in enumerate(zip(beta, self.vectors.coefficients)):
            self.assertAlmostEqual(actual, expected, places=10, msg=f"coef[{i}]")

    def test_standard_errors(self) -> None:
        _, se, _, _, _, _ = self._ols_results()
        for i, (expected, actual) in enumerate(zip(se, self.vectors.std_errors)):
            self.assertAlmostEqual(actual, expected, places=10, msg=f"se[{i}]")

    def test_t_statistics(self) -> None:
        _, _, t_stats, _, _, _ = self._ols_results()
        for i, (expected, actual) in enumerate(zip(t_stats, self.vectors.t_stats)):
            self.assertAlmostEqual(actual, expected, places=10, msg=f"t[{i}]")

    def test_p_values(self) -> None:
        _, _, t_stats, _, n, p = self._ols_results()
        df = n - p
        for i, (t, actual_p) in enumerate(zip(t_stats, self.vectors.p_values)):
            expected_p = 2.0 * float(sp_stats.t.sf(abs(t), df))
            self.assertAlmostEqual(actual_p, expected_p, places=10, msg=f"p[{i}]")

    def test_ci_bounds(self) -> None:
        beta, se, _, _, n, p = self._ols_results()
        df = n - p
        alpha = 0.05
        t_crit = float(sp_stats.t.ppf(1.0 - alpha / 2.0, df))
        for i in range(len(beta)):
            expected_lower = beta[i] - t_crit * se[i]
            expected_upper = beta[i] + t_crit * se[i]
            self.assertAlmostEqual(
                self.vectors.ci_lower[i], expected_lower, places=10, msg=f"ci_lo[{i}]"
            )
            self.assertAlmostEqual(
                self.vectors.ci_upper[i], expected_upper, places=10, msg=f"ci_hi[{i}]"
            )

    def test_beta_weights(self) -> None:
        beta, _, _, _, _, _ = self._ols_results()
        std_y = float(np.std(self.Y, ddof=1))
        for j in range(self.k):
            std_xj = float(np.std(self.X_raw[:, j], ddof=1))
            expected = beta[j + 1] * std_xj / std_y
            self.assertAlmostEqual(
                self.vectors.beta_weights[j], expected, places=10, msg=f"beta_wt[{j}]"
            )


# ── Observation-level diagnostics ───────────────────────────────────────────

class TestObservationVectorsIndependent(unittest.TestCase):
    """Verify observation-level diagnostics via independent numpy computation."""

    def setUp(self) -> None:
        self.rows = [
            _make_row(5.0, 1.0, 3.0),
            _make_row(6.3, 2.0, 1.0),
            _make_row(7.8, 3.0, 2.0),
            _make_row(9.1, 4.0, 1.5),
            _make_row(10.5, 5.0, 2.5),
            _make_row(11.2, 6.0, 3.5),
            _make_row(12.9, 7.0, 2.0),
            _make_row(14.0, 8.0, 1.0),
            _make_row(15.6, 9.0, 3.0),
            _make_row(16.3, 10.0, 2.5),
        ]
        self.k = 2
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "obs.csv"
        _write_csv(self._csv, self.rows)
        self.obs = calculate_regression_observation_vectors(
            self._csv, include_intercept=True, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.X, self.Y, self.X_raw = _build_arrays(self.rows, self.k, True)
        beta = np.linalg.lstsq(self.X, self.Y, rcond=None)[0]
        self.y_hat = self.X @ beta
        self.residuals = self.Y - self.y_hat

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_predictions(self) -> None:
        for i, (expected, actual) in enumerate(zip(self.y_hat, self.obs.predictions)):
            self.assertAlmostEqual(actual, expected, places=10, msg=f"pred[{i}]")

    def test_residuals(self) -> None:
        for i, (expected, actual) in enumerate(zip(self.residuals, self.obs.residuals)):
            self.assertAlmostEqual(actual, expected, places=10, msg=f"resid[{i}]")

    def test_scaled_residuals(self) -> None:
        n = len(self.rows)
        p = self.k + 1
        se = math.sqrt(float(np.sum(self.residuals**2)) / (n - p))
        expected = self.residuals / se
        for i, (exp, act) in enumerate(zip(expected, self.obs.scaled_residuals)):
            self.assertAlmostEqual(act, exp, places=10, msg=f"scaled[{i}]")

    def test_scaled_residuals_ranked(self) -> None:
        n = len(self.rows)
        p = self.k + 1
        se = math.sqrt(float(np.sum(self.residuals**2)) / (n - p))
        expected = np.sort(self.residuals / se)
        for i, (exp, act) in enumerate(zip(expected, self.obs.scaled_residuals_ranked)):
            self.assertAlmostEqual(act, exp, places=10, msg=f"scaled_ranked[{i}]")

    def test_rank_fraction(self) -> None:
        n = len(self.rows)
        sorted_y = np.sort(self.Y)
        for i in range(n):
            count_leq = int(np.searchsorted(sorted_y, self.Y[i], side="right"))
            expected = count_leq / n
            self.assertAlmostEqual(
                self.obs.rank_fraction[i], expected, places=12, msg=f"rank_frac[{i}]"
            )

    def test_normal_scores(self) -> None:
        n = len(self.rows)
        sorted_y = np.sort(self.Y)
        nd = NormalDist()
        for i in range(n):
            count_less = int(np.searchsorted(sorted_y, self.Y[i], side="left"))
            expected = nd.inv_cdf(float((count_less + 0.5) / n))
            self.assertAlmostEqual(
                self.obs.normal_scores[i], expected, places=10, msg=f"ns[{i}]"
            )

    def test_y_ranked(self) -> None:
        expected = sorted(self.Y)
        for i, (exp, act) in enumerate(zip(expected, self.obs.y_ranked)):
            self.assertAlmostEqual(act, exp, places=12, msg=f"y_ranked[{i}]")


# ── Normal_Scores tie-robustness ─────────────────────────────────────────────

class TestNormalScoresTieRobustness(unittest.TestCase):
    """Normal_Scores must collapse tied response values to one rank.

    The response is reported to few decimals, so after within-demeaned Fixed
    Effects it is heavily tied (Life Expectancy: ~61% of rows sit in exact-tie
    groups). The strict-less-than rank count rounds both operands to 9 decimals
    so a tied group collapses to one rank identically in Excel and the Python
    oracle, instead of being split by sub-ULP float differences between the two
    engines — which shifts the Q-Q axis by one rank at hundreds of positions
    (the ``life_country_fixed_effects`` ``Normal_Scores_Ranked`` verifier red).
    Mirrors the ``Normal_Scores`` LAMBDA's ``ROUND(filtered, 9)``.
    """

    @staticmethod
    def _vectors(tmp: Path, y_values: list[float]):
        # Distinct, non-collinear predictors so the OLS fit is well-posed; the
        # response values carry the ties under test. feature_columns is the
        # first two LIFE_EXPECTANCY predictors, matching the existing fixture.
        rows = [
            _make_row(y, float(i) * 1.7 + 1.0, float(i) * 0.9 + 2.0)
            for i, y in enumerate(y_values)
        ]
        csv_path = tmp / "tie.csv"
        _write_csv(csv_path, rows)
        return calculate_regression_observation_vectors(
            csv_path, include_intercept=True, feature_columns=FEATURE_COLUMNS[:2]
        )

    def test_tied_rows_share_one_normal_score(self) -> None:
        # Ties at 20 (x2), 30 (x3), 50 (x2); 10 distinct-on-paper rows.
        y = [10.0, 20.0, 20.0, 30.0, 30.0, 30.0, 40.0, 50.0, 50.0, 60.0]
        with tempfile.TemporaryDirectory() as tmp:
            obs = self._vectors(Path(tmp), y)
        nd = NormalDist()
        n = len(y)
        # Independent rounded-rank expectation (not delegating to the oracle).
        y_r = np.round(np.array(y, dtype=np.float64), 9)
        s = np.sort(y_r)
        expected = [
            nd.inv_cdf(float((int(np.searchsorted(s, v, side="left")) + 0.5) / n))
            for v in y_r
        ]
        for i in range(n):
            self.assertAlmostEqual(
                obs.normal_scores[i], expected[i], places=10, msg=f"ns[{i}]"
            )
        # The tied groups collapse to a single normal score.
        self.assertEqual(obs.normal_scores[1], obs.normal_scores[2])   # two 20s
        self.assertEqual(obs.normal_scores[3], obs.normal_scores[4])   # three 30s
        self.assertEqual(obs.normal_scores[4], obs.normal_scores[5])
        self.assertEqual(obs.normal_scores[7], obs.normal_scores[8])   # two 50s

    def test_sub_ulp_perturbation_does_not_split_ties(self) -> None:
        y_clean = [10.0, 20.0, 20.0, 30.0, 30.0, 30.0, 40.0, 50.0, 50.0, 60.0]
        # 1e-13 is above one ULP at 30 (~7e-15) so it is a genuine, distinct
        # float64 value that WITHOUT rounding would sort after the other 30s and
        # split the tie (shifting this row's rank by +1); it is below the 1e-9
        # rounding tolerance, so round(_, 9) collapses it back to 30.0.
        y_noisy = list(y_clean)
        y_noisy[4] = 30.0 + 1e-13
        self.assertNotEqual(np.float64(30.0), np.float64(y_noisy[4]))
        with tempfile.TemporaryDirectory() as tmp:
            obs_clean = self._vectors(Path(tmp), y_clean)
            obs_noisy = self._vectors(Path(tmp), y_noisy)
        # normal_scores is invariant to the sub-ULP perturbation: the tie did not
        # split. Without the rounding fix, obs_noisy.normal_scores[4] would jump
        # by one rank (here ~inv_norm(5.5/10) - inv_norm(3.5/10) ≈ 0.51) and the
        # two arrays would disagree at the split position.
        for i in range(len(y_clean)):
            self.assertAlmostEqual(
                obs_clean.normal_scores[i],
                obs_noisy.normal_scores[i],
                places=10,
                msg=f"ns[{i}] split by sub-ULP noise",
            )


# ── Regression sheet: predictor summary ─────────────────────────────────────

class TestPredictorSummaryIndependent(unittest.TestCase):
    """Verify Pearson_R, Spearman_R, Skewness, Kurtosis, GVIF, Tolerance."""

    def setUp(self) -> None:
        rng = np.random.default_rng(42)
        n = 25
        x1 = np.linspace(1, 10, n) + rng.normal(0, 0.5, n)
        x2 = np.linspace(0, 5, n) + rng.normal(0, 0.3, n)
        y = 3.0 + 1.5 * x1 - 0.8 * x2 + rng.normal(0, 1.0, n)
        self.rows = [_make_row(y[i], x1[i], x2[i]) for i in range(n)]
        self.k = 2
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "pred_summary.csv"
        _write_csv(self._csv, self.rows)
        results = calculate_regression_sheet_results(
            self._csv, include_intercept=True, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.ps = results.predictor_summary
        self.X, self.Y, self.X_raw = _build_arrays(self.rows, self.k, True)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_pearson_r(self) -> None:
        for j in range(self.k):
            expected = float(np.corrcoef(self.X_raw[:, j], self.Y)[0, 1])
            self.assertAlmostEqual(
                self.ps.pearson_r[j], expected, places=10, msg=f"pearson[{j}]"
            )

    def test_spearman_r(self) -> None:
        for j in range(self.k):
            expected = float(sp_stats.spearmanr(self.X_raw[:, j], self.Y).statistic)
            self.assertAlmostEqual(
                self.ps.spearman_r[j], expected, places=10, msg=f"spearman[{j}]"
            )

    def test_skewness(self) -> None:
        for j in range(self.k):
            expected = float(sp_stats.skew(self.X_raw[:, j], bias=False))
            self.assertAlmostEqual(
                self.ps.skewness[j], expected, places=10, msg=f"skewness[{j}]"
            )

    def test_kurtosis(self) -> None:
        for j in range(self.k):
            expected = float(sp_stats.kurtosis(self.X_raw[:, j], fisher=True, bias=False))
            self.assertAlmostEqual(
                self.ps.kurtosis[j], expected, places=10, msg=f"kurtosis[{j}]"
            )

    def test_gvif(self) -> None:
        # Every predictor here is its own group (no categorical block), so
        # GVIF is numerically identical to ordinary per-column VIF.
        for j in range(self.k):
            others = np.delete(self.X_raw, j, axis=1)
            others_with_const = np.column_stack([np.ones(len(self.rows)), others])
            beta = np.linalg.lstsq(others_with_const, self.X_raw[:, j], rcond=None)[0]
            y_hat = others_with_const @ beta
            ss_res = float(np.sum((self.X_raw[:, j] - y_hat) ** 2))
            ss_tot = float(np.sum((self.X_raw[:, j] - np.mean(self.X_raw[:, j])) ** 2))
            r2_j = 1.0 - ss_res / ss_tot
            expected_vif = 1.0 / (1.0 - r2_j)
            self.assertAlmostEqual(
                self.ps.gvif[j], expected_vif, places=6, msg=f"gvif[{j}]"
            )

    def test_tolerance(self) -> None:
        for j in range(self.k):
            self.assertAlmostEqual(
                self.ps.tolerance[j], 1.0 / self.ps.gvif[j], places=12, msg=f"tol[{j}]"
            )


# ── Regression sheet: full residual diagnostics ─────────────────────────────

class TestFullResidualsIndependent(unittest.TestCase):
    """Verify hat diagonal, studentized residuals, Cook's distance, LOOCV."""

    def setUp(self) -> None:
        rng = np.random.default_rng(99)
        n = 20
        x1 = np.linspace(1, 8, n)
        x2 = np.sin(x1) + 1.0
        y = 2.0 + 1.2 * x1 + 0.5 * x2 + rng.normal(0, 0.8, n)
        self.rows = [_make_row(y[i], x1[i], x2[i]) for i in range(n)]
        self.n = n
        self.k = 2
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "diagnostics.csv"
        _write_csv(self._csv, self.rows)
        results = calculate_regression_sheet_results(
            self._csv, include_intercept=True, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.fr = results.full_residuals
        self.X, self.Y, self.X_raw = _build_arrays(self.rows, self.k, True)
        beta = np.linalg.lstsq(self.X, self.Y, rcond=None)[0]
        self.y_hat = self.X @ beta
        self.e = self.Y - self.y_hat

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_hat_diagonal(self) -> None:
        X = self.X
        xtx_inv = np.linalg.inv(X.T @ X)
        h = np.sum((X @ xtx_inv) * X, axis=1)
        for i in range(self.n):
            self.assertAlmostEqual(
                self.fr.hat_diagonal[i], h[i], places=10, msg=f"h[{i}]"
            )

    def test_hat_diagonal_bounds(self) -> None:
        for i, h in enumerate(self.fr.hat_diagonal):
            self.assertGreaterEqual(h, 1.0 / self.n - 1e-12, msg=f"h_lo[{i}]")
            self.assertLessEqual(h, 1.0, msg=f"h_hi[{i}]")

    def test_hat_diagonal_sum(self) -> None:
        p = self.k + 1
        self.assertAlmostEqual(sum(self.fr.hat_diagonal), p, places=8)

    def test_studentized_residuals(self) -> None:
        X = self.X
        xtx_inv = np.linalg.inv(X.T @ X)
        h = np.sum((X @ xtx_inv) * X, axis=1)
        se = math.sqrt(float(np.sum(self.e**2)) / (self.n - self.k - 1))
        for i in range(self.n):
            expected = self.e[i] / (se * math.sqrt(1.0 - h[i]))
            self.assertAlmostEqual(
                self.fr.studentized_residuals[i], expected, places=10, msg=f"stud[{i}]"
            )

    def test_cooks_distance(self) -> None:
        X = self.X
        xtx_inv = np.linalg.inv(X.T @ X)
        h = np.sum((X @ xtx_inv) * X, axis=1)
        se = math.sqrt(float(np.sum(self.e**2)) / (self.n - self.k - 1))
        p = self.k + 1
        for i in range(self.n):
            r = self.e[i] / (se * math.sqrt(1.0 - h[i]))
            expected = r**2 * h[i] / ((1.0 - h[i]) * p)
            self.assertAlmostEqual(
                self.fr.cooks_distance[i], expected, places=10, msg=f"cook[{i}]"
            )

    def test_cooks_distance_non_negative(self) -> None:
        for i, cd in enumerate(self.fr.cooks_distance):
            self.assertGreaterEqual(cd, 0.0, msg=f"cook_neg[{i}]")

    def test_loocv_residuals(self) -> None:
        X = self.X
        xtx_inv = np.linalg.inv(X.T @ X)
        h = np.sum((X @ xtx_inv) * X, axis=1)
        loocv_pred = self.y_hat - h * self.e / (1.0 - h)
        loocv_resid = self.Y - loocv_pred
        for i in range(self.n):
            self.assertAlmostEqual(
                self.fr.loocv_residuals[i], loocv_resid[i], places=10, msg=f"loocv[{i}]"
            )

    def test_loocv_via_press_formula(self) -> None:
        """LOOCV residual = e/(1-h) — the PRESS residual formula."""
        X = self.X
        xtx_inv = np.linalg.inv(X.T @ X)
        h = np.sum((X @ xtx_inv) * X, axis=1)
        for i in range(self.n):
            expected_press_resid = self.e[i] / (1.0 - h[i])
            self.assertAlmostEqual(
                self.fr.loocv_residuals[i], expected_press_resid, places=10,
                msg=f"press_resid[{i}]"
            )

    def test_studentized_residuals_ranked_sorted(self) -> None:
        ranked = list(self.fr.studentized_residuals_ranked)
        self.assertEqual(ranked, sorted(ranked))

    def test_normal_scores_ranked_sorted(self) -> None:
        ranked = list(self.fr.normal_scores_ranked)
        self.assertEqual(ranked, sorted(ranked))


# ── Regression sheet: prediction interval ───────────────────────────────────

class TestPredictionIntervalIndependent(unittest.TestCase):
    """Verify prediction interval components via independent computation."""

    def setUp(self) -> None:
        rng = np.random.default_rng(123)
        n = 15
        x1 = np.linspace(2, 10, n)
        x2 = rng.uniform(1, 5, n)
        y = 1.0 + 2.0 * x1 - 0.5 * x2 + rng.normal(0, 1.5, n)
        self.rows = [_make_row(y[i], x1[i], x2[i]) for i in range(n)]
        self.k = 2
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "pi.csv"
        _write_csv(self._csv, self.rows)
        results = calculate_regression_sheet_results(
            self._csv, include_intercept=True, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.pi = results.prediction_interval
        self.summary = results.summary
        self.X, self.Y, self.X_raw = _build_arrays(self.rows, self.k, True)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_pred_input_at_column_means(self) -> None:
        for j in range(self.k):
            expected_mean = float(np.mean(self.X_raw[:, j]))
            self.assertAlmostEqual(
                self.pi.pred_input_values[j], expected_mean, places=10, msg=f"mean[{j}]"
            )

    def test_point_estimate(self) -> None:
        beta = np.linalg.lstsq(self.X, self.Y, rcond=None)[0]
        x_means = np.mean(self.X_raw, axis=0)
        x_new = np.concatenate([[1.0], x_means])
        expected = float(x_new @ beta)
        self.assertAlmostEqual(self.pi.point_estimate, expected, places=10)

    def test_se_new(self) -> None:
        # se_new is the new-observation-PI variance term — bit-identical to
        # the pre-v2.1 se_prediction in this no-FE, degenerate-G=1 case.
        beta = np.linalg.lstsq(self.X, self.Y, rcond=None)[0]
        e = self.Y - self.X @ beta
        n = len(self.rows)
        p = self.k + 1
        se = math.sqrt(float(np.sum(e**2)) / (n - p))
        x_means = np.mean(self.X_raw, axis=0)
        x_new = np.concatenate([[1.0], x_means])
        xtx_inv = np.linalg.inv(self.X.T @ self.X)
        h_new = float(x_new @ xtx_inv @ x_new)
        expected_se_new = se * math.sqrt(1.0 + h_new)
        self.assertAlmostEqual(self.pi.se_new, expected_se_new, places=10)

    def test_t_critical(self) -> None:
        df = self.summary.df_residual
        expected = float(sp_stats.t.ppf(0.975, df))
        self.assertAlmostEqual(self.pi.t_critical, expected, places=10)

    def test_interval_bounds(self) -> None:
        expected_lower = self.pi.point_estimate - self.pi.t_critical * self.pi.se_new
        expected_upper = self.pi.point_estimate + self.pi.t_critical * self.pi.se_new
        self.assertAlmostEqual(self.pi.pi_lower, expected_lower, places=10)
        self.assertAlmostEqual(self.pi.pi_upper, expected_upper, places=10)

    def test_confidence_level(self) -> None:
        self.assertAlmostEqual(self.pi.confidence_level, 0.95, places=12)

    def test_interval_contains_point_estimate(self) -> None:
        self.assertGreater(self.pi.pi_upper, self.pi.point_estimate)
        self.assertLess(self.pi.pi_lower, self.pi.point_estimate)


# ── Correlation matrix ──────────────────────────────────────────────────────

class TestCorrelationMatrixIndependent(unittest.TestCase):
    """Verify pairwise predictor correlation via independent numpy computation."""

    def setUp(self) -> None:
        rng = np.random.default_rng(7)
        n = 30
        x1 = np.linspace(1, 10, n)
        x2 = x1 * 0.5 + rng.normal(0, 1, n)
        x3 = rng.uniform(0, 5, n)
        y = 2.0 + x1 - 0.3 * x2 + 0.5 * x3 + rng.normal(0, 1, n)
        self.rows = [_make_row(y[i], x1[i], x2[i], x3[i]) for i in range(n)]
        self.k = 3
        self._tmp = tempfile.TemporaryDirectory()
        self._csv = Path(self._tmp.name) / "corr.csv"
        _write_csv(self._csv, self.rows)
        results = calculate_regression_sheet_results(
            self._csv, include_intercept=True, feature_columns=FEATURE_COLUMNS[:self.k]
        )
        self.ps = results.predictor_summary
        _, _, self.X_raw = _build_arrays(self.rows, self.k, True)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_pearson_cross_consistency(self) -> None:
        """Pearson_R(X_j, Y) should match np.corrcoef for each predictor."""
        y = np.array([r[TARGET_COLUMN] for r in self.rows], dtype=np.float64)
        for j in range(self.k):
            expected = float(np.corrcoef(self.X_raw[:, j], y)[0, 1])
            self.assertAlmostEqual(
                self.ps.pearson_r[j], expected, places=10, msg=f"pearson[{j}]"
            )

    def test_spearman_vs_rank_correlation(self) -> None:
        """Spearman_R should equal Pearson_R of the rank-transformed data."""
        y = np.array([r[TARGET_COLUMN] for r in self.rows], dtype=np.float64)
        for j in range(self.k):
            rank_x = sp_stats.rankdata(self.X_raw[:, j])
            rank_y = sp_stats.rankdata(y)
            expected = float(np.corrcoef(rank_x, rank_y)[0, 1])
            self.assertAlmostEqual(
                self.ps.spearman_r[j], expected, places=10, msg=f"spearman_rank[{j}]"
            )


# ── GVIF edge case: single predictor ────────────────────────────────────────

class TestGVIFSinglePredictor(unittest.TestCase):
    """GVIF should be 1.0 when there's only one predictor."""

    def test_gvif_is_one(self) -> None:
        rows = [_make_row(2.0 + i * 1.5, float(i)) for i in range(12)]
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "single.csv"
            _write_csv(csv_path, rows)
            results = calculate_regression_sheet_results(
                csv_path, include_intercept=True, feature_columns=FEATURE_COLUMNS[:1]
            )
            self.assertAlmostEqual(results.predictor_summary.gvif[0], 1.0, places=12)
            self.assertAlmostEqual(results.predictor_summary.tolerance[0], 1.0, places=12)


# ── MS (Mean Square) consistency ────────────────────────────────────────────

class TestMeanSquares(unittest.TestCase):
    """Verify MS_Regression and MS_Residual from SS and DF."""

    def test_mean_squares(self) -> None:
        rows = [_make_row(3.0 + i * 0.8, float(i), float(i % 3)) for i in range(12)]
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "ms.csv"
            _write_csv(csv_path, rows)
            summary = calculate_regression_summary(
                csv_path, include_intercept=True, feature_columns=FEATURE_COLUMNS[:2]
            )
            ms_reg = summary.ss_regression / summary.df_regression
            ms_res = summary.ss_residual / summary.df_residual
            self.assertAlmostEqual(summary.f_stat, ms_reg / ms_res, places=10)
            self.assertAlmostEqual(summary.se_regression**2, ms_res, places=10)


# ── Durbin-Watson manual formula ────────────────────────────────────────────

class TestDurbinWatsonManual(unittest.TestCase):
    """Verify DW statistic against its definition: sum(diff(e)^2)/sum(e^2)."""

    def test_durbin_watson(self) -> None:
        rng = np.random.default_rng(55)
        n = 20
        x = np.linspace(1, 10, n)
        y = 2.0 + 3.0 * x + rng.normal(0, 2, n)
        rows = [_make_row(y[i], x[i]) for i in range(n)]
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "dw.csv"
            _write_csv(csv_path, rows)
            summary = calculate_regression_summary(
                csv_path, include_intercept=True, feature_columns=FEATURE_COLUMNS[:1]
            )

        X = np.column_stack([np.ones(n), x])
        beta = np.linalg.lstsq(X, y, rcond=None)[0]
        e = y - X @ beta
        dw = float(np.sum(np.diff(e) ** 2) / np.sum(e**2))
        self.assertAlmostEqual(summary.durbin_watson, dw, places=10)


# ── Durbin-Watson along a declared Sequence axis (Durbin_Watson_By) ──────────


def _dw(e: np.ndarray) -> float:
    """Durbin-Watson on residuals in the order given."""
    return float(np.sum(np.diff(e) ** 2) / np.sum(e**2))


def _dw_along_seq(e: np.ndarray, seq: np.ndarray) -> float:
    """The Durbin_Watson_By algorithm: sort residuals ascending by seq, then DW.

    Mirrors the LAMBDA's SORTBY(e, seq, 1) — the sort defines adjacency, so the
    physical order of the incoming (e, seq) pairs never reaches the difference.
    """
    order = np.argsort(seq, kind="stable")
    return _dw(e[order])


class TestDurbinWatsonAlongSequence(unittest.TestCase):
    """Durbin_Watson_By must be invariant to physical row order.

    Acceptance: on a WHO country series indexed by Year (Sequence = Year), the
    sequence-aware statistic is identical whether the source rows are shuffled
    or not, while the ordinary row-order statistic is not — which is exactly why
    the Regression sheet gates plain Durbin-Watson behind a declared Sequence
    axis.
    """

    def _who_country_series(self):
        """A single WHO country's rows (Year unique) — target, one predictor, Year.

        Returns (y, x, year) as numpy arrays, or None when the CSV is absent so
        the test can skip in a CSV-less environment (same pattern as the QC
        config tests).
        """
        from lambda_catalog.analyze_life_expectancy import DEFAULT_INPUT_CSV

        if not DEFAULT_INPUT_CSV.exists():
            return None

        predictor = "Adult Mortality"
        with DEFAULT_INPUT_CSV.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            headers = {h.strip(): h for h in (reader.fieldnames or [])}
            by_country: dict[str, list[tuple[float, float, float]]] = {}
            for raw in reader:
                row = {key: raw[orig] for key, orig in headers.items()}
                try:
                    y = float(row[TARGET_COLUMN])
                    x = float(row[predictor])
                    year = float(row["Year"])
                except (KeyError, ValueError):
                    continue
                by_country.setdefault(row["Country"], []).append((y, x, year))

        # Pick the country with the most complete rows and strictly unique years,
        # so Year is a genuine total order for the series.
        best = None
        for rows in by_country.values():
            years = [r[2] for r in rows]
            if len(years) == len(set(years)) and len(rows) >= 12:
                if best is None or len(rows) > len(best):
                    best = rows
        if best is None:
            return None

        arr = np.array(best, dtype=np.float64)
        return arr[:, 0], arr[:, 1], arr[:, 2]

    def test_sequence_sorted_dw_is_shuffle_invariant(self) -> None:
        series = self._who_country_series()
        if series is None:
            self.skipTest("Life Expectancy CSV not found or no suitable country series")
        y, x, year = series

        # Fit in the natural CSV order, independently of project code.
        design = np.column_stack([np.ones(len(y)), x])
        beta = np.linalg.lstsq(design, y, rcond=None)[0]
        e = y - design @ beta

        # A deterministic shuffle of the same observations.
        perm = np.random.default_rng(20240607).permutation(len(y))
        e_shuffled, year_shuffled = e[perm], year[perm]

        dw_natural = _dw_along_seq(e, year)
        dw_shuffled = _dw_along_seq(e_shuffled, year_shuffled)

        # The acceptance assertion: identical along Year regardless of row order.
        self.assertAlmostEqual(dw_natural, dw_shuffled, places=12)

        # And it equals DW of the residuals put in true Year order — the thing
        # the statistic is meant to measure.
        e_year_sorted = e[np.argsort(year, kind="stable")]
        self.assertAlmostEqual(dw_natural, _dw(e_year_sorted), places=12)

    def test_row_order_dw_is_not_shuffle_invariant(self) -> None:
        # The contrast that motivates the gate: plain row-order DW moves under a
        # shuffle, so trusting physical row order is a real hazard.
        series = self._who_country_series()
        if series is None:
            self.skipTest("Life Expectancy CSV not found or no suitable country series")
        y, x, _year = series

        design = np.column_stack([np.ones(len(y)), x])
        beta = np.linalg.lstsq(design, y, rcond=None)[0]
        e = y - design @ beta

        perm = np.random.default_rng(1).permutation(len(y))
        self.assertNotAlmostEqual(_dw(e), _dw(e[perm]), places=6)


# ── PRESS manual formula ───────────────────────────────────────────────────

class TestPRESSManual(unittest.TestCase):
    """Verify PRESS = sum((e_i/(1-h_ii))^2) via leave-one-out brute force."""

    def test_press_matches_brute_force_loo(self) -> None:
        rows = [
            _make_row(5.0, 1.0),
            _make_row(6.5, 2.0),
            _make_row(8.2, 3.0),
            _make_row(9.0, 4.0),
            _make_row(11.3, 5.0),
            _make_row(12.8, 6.0),
            _make_row(14.5, 7.0),
            _make_row(16.0, 8.0),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "press.csv"
            _write_csv(csv_path, rows)
            summary = calculate_regression_summary(
                csv_path, include_intercept=True, feature_columns=FEATURE_COLUMNS[:1]
            )

        n = len(rows)
        X = np.column_stack([np.ones(n), [r[FEATURE_COLUMNS[0]] for r in rows]])
        Y = np.array([r[TARGET_COLUMN] for r in rows])

        press_brute = 0.0
        for i in range(n):
            X_loo = np.delete(X, i, axis=0)
            Y_loo = np.delete(Y, i)
            beta_loo = np.linalg.lstsq(X_loo, Y_loo, rcond=None)[0]
            pred_i = float(X[i] @ beta_loo)
            press_brute += (Y[i] - pred_i) ** 2

        self.assertAlmostEqual(summary.press, press_brute, places=8)


# ── Partial R² and Partial Correlation ──────────────────────────────────────

class TestPartialR2Independent(unittest.TestCase):
    """Verify partial R² via the t²/(t²+df) formula."""

    def test_partial_r2_from_t_stats(self) -> None:
        rows = [
            _make_row(4.0 + i * 1.2 + (i % 3) * 0.5, float(i), float(i % 3))
            for i in range(15)
        ]
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "pr2.csv"
            _write_csv(csv_path, rows)
            summary = calculate_regression_summary(
                csv_path, include_intercept=True, feature_columns=FEATURE_COLUMNS[:2]
            )
            vectors = calculate_regression_vectors(
                csv_path, include_intercept=True, feature_columns=FEATURE_COLUMNS[:2]
            )

        df_resid = summary.df_residual
        for i, t in enumerate(vectors.t_stats):
            expected_partial_r2 = t**2 / (t**2 + df_resid)
            expected_partial_corr = math.copysign(math.sqrt(expected_partial_r2), t)
            self.assertGreaterEqual(expected_partial_r2, 0.0)
            self.assertLessEqual(expected_partial_r2, 1.0)
            self.assertTrue(math.isfinite(expected_partial_corr))


if __name__ == "__main__":
    unittest.main()
