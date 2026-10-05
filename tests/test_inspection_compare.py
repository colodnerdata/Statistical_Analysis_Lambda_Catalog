"""Tests for workbook inspection value comparisons."""

from lambda_catalog.inspection_compare import compare_values


def test_compare_values_treats_one_missing_value_as_deviation():
    assert compare_values(1.0, None) == (None, 0)
    assert compare_values(None, 1.0) == (None, 0)


def test_compare_values_accepts_two_missing_values():
    assert compare_values(None, None) == (None, None)


def test_compare_values_treats_nan_or_inf_expected_as_missing():
    # A leverage-1 row makes the residual denominator 0: the Python oracle
    # computes NaN/Inf there, and Excel's IFERROR(...,NA()) guard reads back
    # as None. Both sides being "undefined" is not a deviation.
    assert compare_values(float("nan"), None) == (None, None)
    assert compare_values(float("inf"), None) == (None, None)
    assert compare_values(float("-inf"), None) == (None, None)


def test_compare_values_nan_expected_against_a_real_number_is_still_a_deviation():
    # If Excel computed a real number while the oracle is undefined, that is
    # a genuine mismatch worth flagging, not a silently-dropped comparison.
    assert compare_values(float("nan"), 1.0) == (None, 0)


# ── Response-scaled comparison (the L05 precision band) ──────────────────


def test_explicit_scale_rescues_a_residual_that_self_scaling_cannot() -> None:
    """A residual inherits the FITTED VALUE's error, not its own magnitude.

    The pair below is a residual of magnitude ~0.33 carrying the absolute
    error of the two ~74-magnitude numbers it is the difference of. Self
    scaling divides it by 1.0 and changes nothing, so an absolute
    decimal-place score rejects a value that agrees to nine significant
    figures. The response scale is the one that matches the error.
    """
    expected, actual = -0.3282595114048945, -0.32825926497521607

    _, absolute = compare_values(expected, actual)
    _, self_scaled = compare_values(expected, actual, scale_free=True)
    _, response_scaled = compare_values(expected, actual, scale=69.8580)

    assert absolute == 6, "absolute decimal places flag a 9-significant-figure match"
    assert self_scaled == 6, "max(|expected|,1.0) divides by 1.0 here — no help"
    assert response_scaled is not None and response_scaled > 6, (
        "dividing by the response scale compares like with like"
    )


def test_explicit_scale_never_tightens_the_comparison() -> None:
    """A response below 1.0 must not make the test STRICTER than absolute.

    ``compare_values`` floors the divisor at 1.0 for the same reason the
    ``scale_free`` branch does: the normalisation exists to stop a large
    response demanding impossible absolute precision, never to demand more of
    a small one.
    """
    expected, actual = 0.5000001, 0.5000002
    _, absolute = compare_values(expected, actual)
    _, tiny_scale = compare_values(expected, actual, scale=1e-6)
    assert tiny_scale == absolute


def test_explicit_scale_still_catches_a_genuinely_wrong_number() -> None:
    """Rescaling loosens the floor; it must not hide a real disagreement."""
    _, fdd = compare_values(70.0, 71.0, scale=69.8580)
    assert fdd is not None and fdd <= 6


def test_residual_band_families_are_disjoint_and_exclude_t_statistics() -> None:
    """The two families are different UNITS, and T_Statistics is in neither.

    A t-statistic is dimensionless and O(1), and its error comes from the
    coefficient — relative error on the order of eps*cond(X) — not from the
    response. A response-derived divisor is not a scale it has; conditioning
    is where its agreement has to be improved.
    """
    from lambda_catalog.regression_spec_sheet_io import (
        _RESPONSE_UNIT_STATS,
        _STANDARDIZED_RESIDUAL_STATS,
    )

    assert not (_RESPONSE_UNIT_STATS & _STANDARDIZED_RESIDUAL_STATS)
    both = _RESPONSE_UNIT_STATS | _STANDARDIZED_RESIDUAL_STATS
    assert "T_Statistics" not in both
    assert "Predictions" in _RESPONSE_UNIT_STATS
    assert "Residuals" in _RESPONSE_UNIT_STATS
    assert "Studentized_Residuals" in _STANDARDIZED_RESIDUAL_STATS


def test_unit_space_family_is_disjoint_from_the_fit_space_residual_band() -> None:
    """The original-units columns must not share the fit-space divisor.

    ``_RESPONSE_UNIT_STATS`` is scaled by the RMS of ``Design_Response()``,
    which under a Log response is in LOG units. The AZ/BA/BB columns are in
    ORIGINAL units. Putting them in the same set hands a divisor of the wrong
    order of magnitude to ``compare_values`` — on P08, ~11.7 against errors of
    order 13,000, so the precision floor is never applied; on a sub-unit Log
    response it runs the other way and makes the check too permissive.
    """
    from lambda_catalog.regression_spec_sheet_io import (
        _RESPONSE_UNIT_SCALARS,
        _RESPONSE_UNIT_STATS,
        _STANDARDIZED_RESIDUAL_STATS,
        _UNIT_SPACE_RESPONSE_STATS,
    )

    assert not (_UNIT_SPACE_RESPONSE_STATS & _RESPONSE_UNIT_STATS)
    assert not (_UNIT_SPACE_RESPONSE_STATS & _STANDARDIZED_RESIDUAL_STATS)
    assert not (_UNIT_SPACE_RESPONSE_STATS & _RESPONSE_UNIT_SCALARS)
    assert _UNIT_SPACE_RESPONSE_STATS == {
        "Unit_Space_Predictions",
        "Unit_Space_Residuals",
        "Unit_Space_LOOCV_Residual",
    }
    assert _RESPONSE_UNIT_SCALARS == {
        "Unit_Space_RMSE",
        "Unit_Space_LOOCV_RMSE",
        "Unit_Space_LOOCV_MAE",
    }
