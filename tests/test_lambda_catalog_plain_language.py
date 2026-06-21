"""Verify that every LAMBDA function definition has a plain_language_summary."""
from __future__ import annotations

import json
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"


def _load_functions() -> list[dict]:
    with DEFINITIONS_PATH.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    functions = payload.get("functions")
    assert isinstance(functions, list), "lambda_functions.json must have a 'functions' array"
    return functions


def test_catalog_entries_include_plain_language_summaries() -> None:
    functions = _load_functions()

    assert functions

    for func in functions:
        name = func.get("name", "<unknown>")
        summary = str(func.get("plain_language_summary", "")).strip()
        assert summary, f"Function {name!r} is missing a plain_language_summary"

    by_name = {func["name"]: func for func in functions}
    assert (
        by_name["R_squared"]["plain_language_summary"]
        == "How much of the outcome's movement the model explains."
    )
    assert (
        by_name["Hat_diagonal"]["plain_language_summary"]
        == "Each row's leverage, meaning how unusual its predictor values are."
    )
