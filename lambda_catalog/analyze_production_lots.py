"""Production Lots dataset source-row loader for the spec-driven Regression oracle."""
from __future__ import annotations

from pathlib import Path

from .write_sheet_csv_dataset import PRODUCTION_LOTS, load_csv_rows

DEFAULT_INPUT_CSV = PRODUCTION_LOTS.default_csv_path


def load_production_lots_source_rows(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[dict[str, object]]:
    """Load the source CSV as row dicts matching the ProductionLotsData table.

    Cell values are typed exactly as the data sheet writer types them
    (int/float/str/None). The Production Lots sheet ships no appended derived
    column, so the row dicts span only the CSV columns. Mirrors
    ``analyze_model_construction.load_source_rows`` (Mileage) and
    ``analyze_life_expectancy.load_life_expectancy_source_rows`` (which does
    append its "Developed Country after 2013" column).
    """
    headers, rows = load_csv_rows(csv_path, PRODUCTION_LOTS)
    return [dict(zip(headers, row)) for row in rows]