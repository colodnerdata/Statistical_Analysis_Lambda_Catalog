from pathlib import Path

from lambda_catalog.write_sheet_lambda_functions import load_catalog_entries


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"


def test_catalog_entries_include_plain_language_summaries() -> None:
    entries = load_catalog_entries(DEFINITIONS_PATH)

    assert entries
    assert all(entry.plain_language_summary.strip() for entry in entries)

    by_name = {entry.name: entry for entry in entries}
    assert (
        by_name["R_squared"].plain_language_summary
        == "How much of the outcome's movement the model explains."
    )
    assert (
        by_name["Hat_diagonal"].plain_language_summary
        == "Each row's leverage, meaning how unusual its predictor values are."
    )
