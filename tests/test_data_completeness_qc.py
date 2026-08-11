"""QC for the Life Expectancy sheet's derived ``Developed Country after 2013``
column.

The sheet appends ``=AND([@Status]="Developed",[@Year]>2013)`` — TRUE for
developed-country rows with Year >= 2014. ``calculate_developed_country_flags``
is the Python-side mirror the source-row loader and the deep verifier rely on,
so these tests pin its semantics against the committed sample CSV and a few
hand-built edge rows.
"""
from __future__ import annotations

import csv
import tempfile
from pathlib import Path

from lambda_catalog.analyze_life_expectancy import (
    DEFAULT_INPUT_CSV,
    calculate_developed_country_flags,
)
from lambda_catalog.write_sheet_csv_dataset import LIFE_EXPECTANCY, load_csv_rows


def _write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    headers = ["Country", "Year", "Status", "Life expectancy "]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({header: row.get(header, "") for header in headers})


def test_developed_country_flags_match_the_and_rule_on_sample_data() -> None:
    flags = calculate_developed_country_flags(DEFAULT_INPUT_CSV)
    headers, rows = load_csv_rows(DEFAULT_INPUT_CSV, LIFE_EXPECTANCY)
    status_idx = headers.index("Status")
    year_idx = headers.index("Year")

    expected = tuple(
        bool(row[status_idx] == "Developed" and _year_over_2013(row[year_idx]))
        for row in rows
    )

    assert flags == expected
    assert any(flags)
    assert any(not flag for flag in flags)


def test_developed_country_flags_require_developed_status_and_year_over_2013() -> None:
    rows = [
        # Developed, 2014 -> True (the boundary: Year > 2013).
        {"Country": "A", "Year": "2014", "Status": "Developed", "Life expectancy ": "70"},
        # Developed, 2013 -> False (Year is not > 2013).
        {"Country": "B", "Year": "2013", "Status": "Developed", "Life expectancy ": "70"},
        # Developing, 2015 -> False (Status is not Developed).
        {"Country": "C", "Year": "2015", "Status": "Developing", "Life expectancy ": "70"},
        # Developed, blank Year -> False (blank is not > 2013).
        {"Country": "D", "Year": "", "Status": "Developed", "Life expectancy ": "70"},
    ]

    with tempfile.TemporaryDirectory() as tmp:
        csv_path = Path(tmp) / "developed_country.csv"
        _write_csv(csv_path, rows)

        assert calculate_developed_country_flags(csv_path) == (True, False, False, False)


def _year_over_2013(value: object) -> bool:
    """Mirror of ``analyze_life_expectancy._is_int_over_2013`` for the oracle."""
    return isinstance(value, (int, float)) and not isinstance(value, bool) and value > 2013