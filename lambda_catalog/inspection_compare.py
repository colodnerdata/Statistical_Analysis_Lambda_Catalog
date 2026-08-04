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
    relative_to: float | None = None,
) -> tuple[float | None, int | None]:
    """Return absolute difference and first-digit deviation for two values.

    A row where the underlying statistic is genuinely undefined (e.g. a
    leverage-1 observation making a residual denominator 0) can surface as
    NaN/Inf on the Python side and as Excel's #N/A on the sheet side. Both
    are normalized to None first so that case reads as "both missing"
    (no deviation) instead of a spurious first-digit-deviation-0 mismatch.

    ``relative_to`` rescales the expected value before the first-digit
    deviation is computed — used for statistics whose magnitude runs into
    1e10 (raw-cost sums of squares on the production-lots data) where the
    precision-floor is the seventh decimal, not the third. The normalised
    expected value is ``expected / max(relative_to, |expected|)`` so the
    comparison always sees an O(1) magnitude, and rounding-floor behaviour
    matches the original 3-decimal tolerance the rest of the suite is
    calibrated against.
    """
    expected = _finite_or_none(expected)
    actual = _finite_or_none(actual)
    if expected is None and actual is None:
        return None, None
    if expected is None or actual is None:
        return None, 0
    diff = abs(actual - expected)
    if relative_to is not None:
        scale = max(abs(relative_to), abs(expected), 1.0)
        return diff, first_digit_deviation(expected / scale, actual / scale)
    return diff, first_digit_deviation(expected, actual)
