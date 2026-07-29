"""Python QC oracle for the Production Lots sheet's Full_Data completeness column."""
from __future__ import annotations

from pathlib import Path

from .write_sheet_csv_dataset import PRODUCTION_LOTS, load_csv_rows

DEFAULT_INPUT_CSV = PRODUCTION_LOTS.default_csv_path

# Every column except Lot_ID (text identifier) and Facility (categorical
# grouping column) is numeric -- mirrors PRODUCTION_LOTS.full_data_formula in
# write_sheet_csv_dataset.py (ProductionLotsData[@[Fiscal_Year]:[log Unit Cost]]).
_COMPLETENESS_COLUMNS = (
    "Fiscal_Year",
    "Lot_Quantity",
    "Cumulative_Units",
    "Experience_Stock",
    "Unit_Cost_BY",
    "log Cum Units",
    "log experience",
    "log Unit Cost",
)


def calculate_production_lots_completeness_flags(
    input_csv_path: Path = DEFAULT_INPUT_CSV,
) -> tuple[bool, ...]:
    """Return expected Full_Data flags matching the Production Lots completeness formula.

    Reuses ``write_sheet_csv_dataset.load_csv_rows`` rather than re-parsing
    the source CSV independently, same reasoning as
    ``analyze_mileage.calculate_mileage_completeness_flags``.
    """
    headers, rows = load_csv_rows(input_csv_path, PRODUCTION_LOTS)
    column_indices = [headers.index(column) for column in _COMPLETENESS_COLUMNS]

    return tuple(
        all(isinstance(row[index], (int, float)) for index in column_indices)
        for row in rows
    )


def load_production_lots_source_rows(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[dict[str, object]]:
    """Load the source CSV as row dicts matching the ProductionLotsData table.

    Cell values are typed exactly as the data sheet writer types them
    (int/float/str/None), and the computed ``Full_Data`` column is appended
    with the same completeness rule the sheet's ``Data_Completeness`` formula
    applies. Mirrors ``analyze_model_construction.load_source_rows``, which
    is hardcoded to the Mileage table.
    """
    headers, rows = load_csv_rows(csv_path, PRODUCTION_LOTS)
    flags = calculate_production_lots_completeness_flags(csv_path)
    table_headers = [*headers, "Full_Data"]
    return [
        dict(zip(table_headers, [*row, flag]))
        for row, flag in zip(rows, flags)
    ]
