from __future__ import annotations

from lambda_catalog.analyze_mileage import (
    DEFAULT_INPUT_XLSX,
    calculate_mileage_completeness_flags,
)
from lambda_catalog.write_sheet_mileage_data import load_mileage_rows


def test_mileage_completeness_matches_expected_na_rows() -> None:
    flags = calculate_mileage_completeness_flags(DEFAULT_INPUT_XLSX)
    headers, rows = load_mileage_rows(DEFAULT_INPUT_XLSX)
    mpg_idx = headers.index("MPG")
    horsepower_idx = headers.index("Horsepower")

    legacy_flags = tuple(
        row[mpg_idx] is not None and row[horsepower_idx] is not None
        for row in rows
    )

    assert flags == legacy_flags
    assert sum(1 for flag in flags if not flag) == 14
    assert any(flags)
    assert any(not flag for flag in flags)


def test_mileage_completeness_treats_na_marker_as_incomplete() -> None:
    headers, rows = load_mileage_rows(DEFAULT_INPUT_XLSX)
    mpg_idx = headers.index("MPG")

    missing_mpg_row_indices = [
        index for index, row in enumerate(rows) if row[mpg_idx] is None
    ]

    flags = calculate_mileage_completeness_flags(DEFAULT_INPUT_XLSX)

    assert missing_mpg_row_indices
    assert all(not flags[index] for index in missing_mpg_row_indices)
