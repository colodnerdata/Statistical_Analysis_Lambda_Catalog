from __future__ import annotations

import csv
import tempfile
from pathlib import Path

from lambda_catalog.analyze_life_expectancy import (
    DEFAULT_INPUT_CSV,
    FEATURE_COLUMNS,
    TARGET_COLUMN,
    _load_normalized_rows,
    _parse_float,
    calculate_data_completeness_flags,
)


def _write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    headers = ["Country", "Year", "Status", TARGET_COLUMN, *FEATURE_COLUMNS]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({header: row.get(header, "") for header in headers})


def _base_row() -> dict[str, str]:
    row = {
        "Country": "Example",
        "Year": "2015",
        "Status": "Developing",
        TARGET_COLUMN: "65.0",
    }
    row.update({name: "1.0" for name in FEATURE_COLUMNS})
    return row


def test_data_completeness_matches_legacy_full_data_criterion_on_sample_data() -> None:
    flags = calculate_data_completeness_flags(DEFAULT_INPUT_CSV)
    _, normalized_rows = _load_normalized_rows(DEFAULT_INPUT_CSV)

    legacy_flags = tuple(
        all(
            _parse_float(row.get(column)) is not None
            for column in [TARGET_COLUMN, *FEATURE_COLUMNS]
        )
        for row in normalized_rows
    )

    assert flags == legacy_flags
    assert any(flags)
    assert any(not flag for flag in flags)


def test_data_completeness_treats_blank_and_text_as_incomplete() -> None:
    rows = []

    complete_row = _base_row()
    rows.append(complete_row)

    blank_row = _base_row()
    blank_row[FEATURE_COLUMNS[0]] = ""
    rows.append(blank_row)

    text_row = _base_row()
    text_row[FEATURE_COLUMNS[1]] = "not numeric"
    rows.append(text_row)

    with tempfile.TemporaryDirectory() as tmp:
        csv_path = Path(tmp) / "data_completeness.csv"
        _write_csv(csv_path, rows)

        assert calculate_data_completeness_flags(csv_path) == (True, False, False)
