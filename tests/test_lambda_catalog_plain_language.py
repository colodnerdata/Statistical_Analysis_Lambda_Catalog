"""Verify that every LAMBDA function definition has a plain_language_summary."""
from __future__ import annotations

from pathlib import Path

from lambda_catalog.catalog_schema import load_catalog_document


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"


def test_catalog_entries_include_plain_language_summaries() -> None:
    document = load_catalog_document(DEFINITIONS_PATH)

    assert document.functions

    by_name = {f.name: f for f in document.functions}
    assert (
        by_name["R_squared"].plain_language_summary
        == "How much of the outcome's movement the model explains."
    )
    assert (
        by_name["Hat_diagonal"].plain_language_summary
        == "Each row's leverage, meaning how unusual its predictor values are."
    )
