"""Python QC oracle for the Mileage Data sheet's Full_Data completeness column."""
from __future__ import annotations

from pathlib import Path

from .write_sheet_mileage_data import DEFAULT_XLSX_PATH, load_mileage_rows

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_XLSX = DEFAULT_XLSX_PATH

# The 6 continuous-measurement columns the Mileage Data sheet's Full_Data
# formula checks (MileageData[@[MPG]:[Acceleration]]), in table order.
_COMPLETENESS_COLUMNS = (
    "MPG",
    "Cylinders",
    "Displacement",
    "Horsepower",
    "Weight",
    "Acceleration",
)


def calculate_mileage_completeness_flags(
    input_xlsx_path: Path = DEFAULT_INPUT_XLSX,
) -> tuple[bool, ...]:
    """Return expected Full_Data flags matching the Mileage Data completeness formula.

    The Mileage Data sheet computes ``Full_Data`` with
    ``Data_Completeness(MileageData[@[MPG]:[Acceleration]])``, so the
    Python-side QC expectation is that all 6 continuous-measurement columns
    parse as numeric on a given row.

    This reuses ``write_sheet_mileage_data.load_mileage_rows`` rather than
    re-parsing the source xlsx independently: unlike the Life Expectancy CSV
    (a simple format with two independently-written readers), the xlsx table
    extraction here is nontrivial enough that a second parser would be a
    maintenance and correctness risk rather than a meaningful cross-check.
    """
    headers, rows = load_mileage_rows(input_xlsx_path)
    column_indices = [headers.index(column) for column in _COMPLETENESS_COLUMNS]

    return tuple(
        all(isinstance(row[index], (int, float)) for index in column_indices)
        for row in rows
    )
