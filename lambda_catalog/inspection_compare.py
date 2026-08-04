"""Shared helpers for QC inspector comparison: numeric coercion and deviation detection."""
from __future__ import annotations

import math
from typing import Any


def to_float_or_none(value: Any) -> float | None:
    """Best-effort numeric conversion for Excel-like cell values."""
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def first_digit_deviation(expected: float, actual: float) -> int | None:
    """Return the decimal place where expected and actual first deviate."""
    if expected == actual:
        return None
    for d in range(15, -1, -1):
        if round(expected, d) == round(actual, d):
            return d + 1
    return 0


def _finite_or_none(value: float | None) -> float | None:
    """Treat NaN/Inf as equivalent to missing — both mean "undefined here"."""
    if value is not None and not math.isfinite(value):
        return None
    return value


def compare_values(
    expected: float | None,
    actual: float | None,
    *,
    scale_free: bool = False,
) -> tuple[float | None, int | None]:
    """Return absolute difference and first-digit deviation for two values.

    A row where the underlying statistic is genuinely undefined (e.g. a
    leverage-1 observation making a residual denominator 0) can surface as
    NaN/Inf on the Python side and as Excel's #N/A on the sheet side. Both
    are normalized to None first so that case reads as "both missing"
    (no deviation) instead of a spurious first-digit-deviation-0 mismatch.

    ``scale_free`` divides BOTH values by ``max(|expected|, 1.0)`` before the
    first-digit deviation is computed, turning an absolute tolerance into a
    relative one. It is for statistics whose magnitude tracks the response
    column — sums of squares on raw production-cost data run into 1e10, where
    the IEEE-754 precision floor already sits above the third decimal and
    Excel and Python cannot agree there by construction.

    Dividing both sides by the same factor leaves their RELATIVE difference
    untouched, so this loosens the precision floor without hiding a genuinely
    wrong number. The divisor deliberately has no fixed floor in it: an
    earlier ``max(1e9, |expected|, 1.0)`` form mapped a 1e10 value to the
    intended O(1) but a 1e4 value to 1e-5, quietly loosening the tolerance on
    every small-magnitude case by five orders of magnitude — the opposite of
    what the O(1) normalisation is for.
    """
    expected = _finite_or_none(expected)
    actual = _finite_or_none(actual)
    if expected is None and actual is None:
        return None, None
    if expected is None or actual is None:
        return None, 0
    diff = abs(actual - expected)
    if scale_free:
        scale = max(abs(expected), 1.0)
        return diff, first_digit_deviation(expected / scale, actual / scale)
    return diff, first_digit_deviation(expected, actual)
