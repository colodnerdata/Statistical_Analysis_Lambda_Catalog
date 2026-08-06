"""Counts of catalog functions written in prose match ``lambda_functions.json``.

This is the count half of check 3 in CONTRIBUTING.md's documentation-drift
proposal. The catalog is the source of truth for *which* functions exist; the
number of them is restated in prose in five places across four documents, and
nothing kept those in step. When this test was written the catalog held 140
entries and the docs claimed 139 (twice), 139, and 131 — plus 17 sheet-scoped
closures for a sheet that has 18. Every one of those was correct when written.

**Phrasings are pinned, not sniffed.** The test looks for three exact phrasings
rather than "a number near the word LAMBDA", because the loose version has a
trap sitting in the first line of README.md: *"Excel 365 LAMBDA functions that
replicate…"*. That 365 is the Office release, and a checker that read it as a
function count would fail the build on a sentence that is perfectly correct.
Adding a new phrasing to the docs means adding it here — the same bargain
``_EXPECTED_TASK_NAMES`` and ``_EXPECTED_CASE_NAMES`` make.

Deliberately NOT pinned: "``N`` catalog functions". It carries two different
meanings in the docs already — README's "around 30 catalog functions are called
by no pre-built sheet" is a subset, while DECISIONS.md's "131 catalog functions"
is a dated statement inside a decision log, where a superseded number is the
record doing its job rather than drift.

Pure ``json`` + ``re``: no Excel, so it runs in the Linux CI job.
"""
from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path

import pytest

ROOT_DIR = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT_DIR / "lambda_functions.json"

# Scope of a claim: the whole catalog, the portable subset, or the Regression
# sheet's closures. Each maps to a phrasing the documentation actually uses.
_TOTAL, _WORKBOOK, _SHEET_SCOPED = "total", "workbook-scoped", "Regression-scoped"

_COUNT_CLAIMS = (
    (re.compile(r"(\d+) LAMBDA definitions"), _TOTAL),
    (re.compile(r"(\d+) portable functions"), _WORKBOOK),
    (re.compile(r"(\d+) constructor closures"), _SHEET_SCOPED),
)

# The documents that describe the catalog as it is now. DECISIONS.md is left
# out on purpose: it is a dated log, and its counts are history.
_CURRENT_STATE_DOCS = ("README.md", "CONTRIBUTING.md", "docs/ROADMAP.md")


def _catalog_counts() -> dict[str, int]:
    functions = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))["functions"]
    scopes = Counter(entry.get("scope", "workbook") for entry in functions)
    return {
        _TOTAL: len(functions),
        _WORKBOOK: scopes["workbook"],
        _SHEET_SCOPED: scopes["Regression"],
    }


def _claims(text: str) -> list[tuple[str, int]]:
    return [
        (scope, int(found))
        for pattern, scope in _COUNT_CLAIMS
        for found in pattern.findall(text)
    ]


def test_every_scope_is_actually_populated() -> None:
    """Guard against a schema change silently zeroing a scope and passing."""
    counts = _catalog_counts()
    assert all(count > 0 for count in counts.values()), counts
    assert counts[_WORKBOOK] + counts[_SHEET_SCOPED] == counts[_TOTAL]


def test_the_excel_365_in_the_readme_headline_is_not_read_as_a_count() -> None:
    """The trap this file's pinned phrasings exist to avoid."""
    headline = "Excel 365 LAMBDA functions that replicate the Analysis ToolPak"
    assert _claims(headline) == []


def test_a_stale_count_is_reported_with_both_numbers() -> None:
    """The failure has to name what the docs say AND what the catalog holds,
    or whoever hits it still has to go count the JSON by hand."""
    assert _claims("all 139 LAMBDA definitions") == [(_TOTAL, 139)]


@pytest.mark.parametrize("relative_path", _CURRENT_STATE_DOCS)
def test_documented_catalog_counts_match_the_catalog(relative_path: str) -> None:
    counts = _catalog_counts()
    text = (ROOT_DIR / relative_path).read_text(encoding="utf-8")

    wrong = [
        f"says {claimed} {scope}, catalog has {counts[scope]}"
        for scope, claimed in _claims(text)
        if claimed != counts[scope]
    ]

    assert not wrong, f"{relative_path}: " + "; ".join(wrong)


def test_at_least_one_document_states_each_count() -> None:
    """A phrasing that stops appearing anywhere is drift too — the check would
    keep passing while covering nothing."""
    stated = {
        scope
        for relative_path in _CURRENT_STATE_DOCS
        for scope, _ in _claims((ROOT_DIR / relative_path).read_text(encoding="utf-8"))
    }
    assert stated == {_TOTAL, _WORKBOOK, _SHEET_SCOPED}
