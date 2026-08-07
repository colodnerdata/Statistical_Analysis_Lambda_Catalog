"""Tests for the shared CSV dataset loader (lambda_catalog/write_sheet_csv_dataset.py)."""
from __future__ import annotations

from pathlib import Path

import pytest

from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    MILEAGE,
    PRODUCTION_LOTS,
    _excel_safe_cell_value,
    load_csv_rows,
)

_EXPECTED_MILEAGE_HEADERS = (
    "MPG",
    "Cylinders",
    "Displacement",
    "Horsepower",
    "Weight",
    "Acceleration",
    "Model Year",
    "Origin",
    "Car Name",
    "Make",
    "Model?",
)


def test_dataset_configs_carry_the_expected_sheet_and_table_names() -> None:
    assert MILEAGE.sheet_name == "Mileage Data"
    assert MILEAGE.table_name == "MileageData"
    assert PRODUCTION_LOTS.sheet_name == "Production Lots"
    assert PRODUCTION_LOTS.table_name == "ProductionLotsData"
    assert LIFE_EXPECTANCY.sheet_name == "Life Expectancy Data"
    assert LIFE_EXPECTANCY.table_name == "LifeExpectancyData"


def test_load_csv_rows_returns_expected_shape_for_mileage() -> None:
    headers, rows = load_csv_rows(MILEAGE.default_csv_path, MILEAGE)

    assert tuple(headers) == _EXPECTED_MILEAGE_HEADERS
    assert len(rows) == 406
    assert all(len(row) == len(headers) for row in rows)


def test_load_csv_rows_parses_known_complete_row_for_mileage() -> None:
    headers, rows = load_csv_rows(MILEAGE.default_csv_path, MILEAGE)
    first_row = dict(zip(headers, rows[0]))

    assert first_row["MPG"] == 18
    assert first_row["Cylinders"] == 8
    assert first_row["Horsepower"] == 130
    assert first_row["Car Name"] == "chevrolet chevelle malibu"
    assert first_row["Make"] == "chevrolet"
    assert first_row["Model?"] == "chevelle malibu"


def test_load_csv_rows_treats_na_marker_as_none_for_mileage() -> None:
    headers, rows = load_csv_rows(MILEAGE.default_csv_path, MILEAGE)
    mpg_idx = headers.index("MPG")
    horsepower_idx = headers.index("Horsepower")

    mpg_missing_rows = [row for row in rows if row[mpg_idx] is None]
    horsepower_missing_rows = [row for row in rows if row[horsepower_idx] is None]

    assert len(mpg_missing_rows) == 8
    assert len(horsepower_missing_rows) == 6
    # The two missing-value sets are disjoint: no row is missing both.
    assert not any(row[mpg_idx] is None for row in horsepower_missing_rows)


def test_load_csv_rows_returns_expected_shape_for_production_lots() -> None:
    headers, rows = load_csv_rows(PRODUCTION_LOTS.default_csv_path, PRODUCTION_LOTS)

    assert len(rows) == 51
    assert all(len(row) == len(headers) for row in rows)
    assert "Facility" in headers
    assert "Fiscal_Year" in headers


def test_load_csv_rows_normalizes_headers_for_life_expectancy(tmp_path: Path) -> None:
    csv_path = tmp_path / "life_expectancy_fixture.csv"
    csv_path.write_text(
        "Country, Life expectancy \nUSA,78.5\n",
        encoding="utf-8",
    )

    headers, rows = load_csv_rows(csv_path, LIFE_EXPECTANCY)

    assert headers == ["Country", "Life expectancy"]
    assert rows == [["USA", 78.5]]


def test_load_csv_rows_does_not_normalize_headers_for_mileage(tmp_path: Path) -> None:
    csv_path = tmp_path / "mileage_fixture.csv"
    csv_path.write_text("  MPG  ,Horsepower\n18,130\n", encoding="utf-8")

    headers, rows = load_csv_rows(csv_path, MILEAGE)

    assert headers == ["  MPG  ", "Horsepower"]
    assert rows == [[18, 130]]


def test_load_csv_rows_on_synthetic_fixture_exercises_blank_na_and_numeric(
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "mini_mpg.csv"
    csv_path.write_text("MPG,Horsepower\n18,NA\n,130\n", encoding="utf-8")

    headers, rows = load_csv_rows(csv_path, MILEAGE)

    assert headers == ["MPG", "Horsepower"]
    assert rows == [[18, None], [None, 130]]


def test_load_csv_rows_treats_only_blank_as_missing_for_production_lots(
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "mini_lots.csv"
    csv_path.write_text("Lot_ID,Facility\nLOT-1,NA\n", encoding="utf-8")

    headers, rows = load_csv_rows(csv_path, PRODUCTION_LOTS)

    # PRODUCTION_LOTS has no "NA" missing-value marker, unlike MILEAGE: the
    # literal text "NA" is a real (string) value here, not a missing marker.
    assert rows == [["LOT-1", "NA"]]


def test_load_csv_rows_skips_fully_blank_lines(tmp_path: Path) -> None:
    csv_path = tmp_path / "blank_line.csv"
    csv_path.write_text("MPG,Horsepower\n18,130\n\n25,90\n", encoding="utf-8")

    headers, rows = load_csv_rows(csv_path, MILEAGE)

    assert headers == ["MPG", "Horsepower"]
    assert rows == [[18, 130], [25, 90]]


def test_load_csv_rows_pads_short_rows_with_missing_values(tmp_path: Path) -> None:
    csv_path = tmp_path / "short_row.csv"
    csv_path.write_text("MPG,Horsepower,Weight\n18\n", encoding="utf-8")

    headers, rows = load_csv_rows(csv_path, MILEAGE)

    assert headers == ["MPG", "Horsepower", "Weight"]
    assert rows == [[18, None, None]]


def test_load_csv_rows_raises_on_row_with_too_many_columns(tmp_path: Path) -> None:
    csv_path = tmp_path / "long_row.csv"
    csv_path.write_text("MPG,Horsepower\n18,130,extra\n", encoding="utf-8")

    with pytest.raises(ValueError, match="line 2"):
        load_csv_rows(csv_path, MILEAGE)


def test_load_csv_rows_raises_on_empty_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "empty.csv"
    csv_path.write_text("", encoding="utf-8")

    with pytest.raises(ValueError, match="CSV file is empty"):
        load_csv_rows(csv_path, MILEAGE)


def test_load_csv_rows_raises_on_headers_but_no_data_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "headers_only.csv"
    csv_path.write_text("MPG,Horsepower\n", encoding="utf-8")

    with pytest.raises(ValueError, match="headers but no data rows"):
        load_csv_rows(csv_path, MILEAGE)


def test_excel_safe_cell_value_escapes_excel_error_literals() -> None:
    assert _excel_safe_cell_value("#VALUE!") == "'#VALUE!"
    assert _excel_safe_cell_value("#N/A") == "'#N/A"


def test_excel_safe_cell_value_leaves_normal_values_unchanged() -> None:
    assert _excel_safe_cell_value("subaru") == "subaru"
    assert _excel_safe_cell_value(18) == 18
    assert _excel_safe_cell_value(15.5) == 15.5
    assert _excel_safe_cell_value(None) is None
