"""Python QC oracle for the Mileage Data sheet's Full_Data completeness column."""
from __future__ import annotations

from pathlib import Path

from .write_sheet_csv_dataset import MILEAGE, load_csv_rows

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_CSV = MILEAGE.default_csv_path

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
    input_csv_path: Path = DEFAULT_INPUT_CSV,
) -> tuple[bool, ...]:
    """Return expected Full_Data flags matching the Mileage Data completeness formula.

    The Mileage Data sheet computes ``Full_Data`` with
    ``Data_Completeness(MileageData[@[MPG]:[Acceleration]])``, so the
    Python-side QC expectation is that all 6 continuous-measurement columns
    parse as numeric on a given row.

    This reuses ``write_sheet_csv_dataset.load_csv_rows`` rather than
    re-parsing the source CSV independently, matching how the Life
    Expectancy oracle reuses its own loader.
    """
    headers, rows = load_csv_rows(input_csv_path, MILEAGE)
    column_indices = [headers.index(column) for column in _COMPLETENESS_COLUMNS]

    return tuple(
        all(isinstance(row[index], (int, float)) for index in column_indices)
        for row in rows
    )
