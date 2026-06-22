"""Shared helpers for inspector comparison and numeric coercion."""
from __future__ import annotations

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


def compare_values(
    expected: float | None,
    actual: float | None,
) -> tuple[float | None, int | None]:
    """Return absolute difference and first-digit deviation for two values."""
    if expected is None and actual is None:
        return None, None
    if expected is None or actual is None:
        return None, 0
    return abs(actual - expected), first_digit_deviation(expected, actual)
