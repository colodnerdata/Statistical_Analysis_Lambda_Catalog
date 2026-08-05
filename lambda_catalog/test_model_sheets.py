"""The sheet-naming contract for the one-sheet-per-test-model artifact.

Every case in the regression test-model suite — the fittable
``RegressionSpecCase`` models and the ``GuardStateCase`` guard-rail
configurations alike — is materialized as its own worksheet in
``Lambda_Library_TestModels.xlsx``. That only works if the names are legal,
unique, and mean something to a human who opens the workbook, so the rules
live here rather than being re-derived at each declaration site.

Three rules, in order of how badly breaking them hurts:

1. **Excel's own constraints.** A worksheet name is 1–31 characters and may
   not contain ``[ ] : * ? / \\``. Excel does not truncate — ``Sheets.Add``
   raises — so an over-long name is a build failure at sheet ~30 of ~40,
   minutes into a run. ``validate_sheet_name`` is called from the case
   registries themselves so the failure lands in a millisecond-long unit
   test instead.
2. **Uniqueness.** Two cases claiming one name would silently mean one case
   never gets verified (the second write lands on the first's sheet). Pinned
   by ``assert_sheet_names_unique``.
3. **The name states the CONCEPT, not the variables.** ``M05 Log-Log NA
   Masking``, never ``MPG ~ Ln(Weight) + Ln(HP)``. Thirty-one characters
   cannot hold a formula, and the formula is the least interesting thing
   about a test case anyway — the sheet exists to exercise one corner, and
   that corner is what the tab should say. The variables are one click away
   in the spec block.

The ``<PlanID> <Concept>`` shape is what ties a sheet back to a row in
``docs/MODEL_TESTING_ASSETS.md``: ``M05`` is that document's M5, ``G03`` its
G3. Renaming a case never orphans the mapping as long as the ID stays.
"""
from __future__ import annotations

import re

# Excel's hard worksheet-name limit. Not a style preference: Sheets.Add
# raises on 32+, it does not truncate.
MAX_SHEET_NAME_LENGTH = 31

# Characters Excel rejects in a worksheet name. The backslash and forward
# slash are path separators, the brackets are the structured-reference /
# external-workbook syntax, and : * ? are wildcards or reserved.
ILLEGAL_SHEET_NAME_CHARS = frozenset(r"[]:*?/\\")

# Excel reserves this one name outright (it backs the change-tracking sheet),
# case-insensitively.
_RESERVED_SHEET_NAMES = frozenset({"history"})

# `<PlanID> <Concept>`: a letter tier (M/L/P/G), a two-digit number, an
# optional lowercase variant letter for the ±intercept twins, then a space
# and the concept text.
PLAN_ID_PATTERN = re.compile(r"^(?P<plan_id>[MLPG]\d{2}[a-z]?) (?P<concept>\S.*)$")


class SheetNameError(ValueError):
    """A test-model sheet name violates the naming contract."""


def validate_sheet_name(name: str) -> None:
    """Raise ``SheetNameError`` unless ``name`` is a legal test-model sheet name.

    Checks Excel's own constraints (length, character set, reserved names,
    the no-leading/trailing-apostrophe rule) and this suite's ``<PlanID>
    <Concept>`` shape. Called from the case registries at import time, so a
    bad name fails in the unit suite rather than partway through a
    multi-minute Excel build.

    Parameters
    ----------
    name : str
        The candidate worksheet name.

    Raises
    ------
    SheetNameError
        With a message naming the specific rule broken.
    """
    if not name:
        raise SheetNameError("Sheet name is empty.")
    if len(name) > MAX_SHEET_NAME_LENGTH:
        raise SheetNameError(
            f"Sheet name {name!r} is {len(name)} characters; Excel's limit is "
            f"{MAX_SHEET_NAME_LENGTH}."
        )
    illegal = sorted(set(name) & ILLEGAL_SHEET_NAME_CHARS)
    if illegal:
        raise SheetNameError(
            f"Sheet name {name!r} contains Excel-illegal character(s): "
            + " ".join(repr(char) for char in illegal)
        )
    if name.startswith("'") or name.endswith("'"):
        raise SheetNameError(
            f"Sheet name {name!r} may not begin or end with an apostrophe."
        )
    if name != name.strip():
        raise SheetNameError(
            f"Sheet name {name!r} has leading or trailing whitespace."
        )
    if name.casefold() in _RESERVED_SHEET_NAMES:
        raise SheetNameError(f"Sheet name {name!r} is reserved by Excel.")
    if PLAN_ID_PATTERN.match(name) is None:
        raise SheetNameError(
            f"Sheet name {name!r} does not match '<PlanID> <Concept>' "
            "(e.g. 'M05 Log-Log NA Masking') — the plan ID is what ties the "
            "sheet back to docs/MODEL_TESTING_ASSETS.md."
        )


def plan_id_of(name: str) -> str:
    """Return the plan ID embedded in a validated sheet name.

    Parameters
    ----------
    name : str
        A sheet name that has passed ``validate_sheet_name``.

    Returns
    -------
    str
        The leading plan ID, e.g. ``"M05"`` or ``"M03b"``.
    """
    match = PLAN_ID_PATTERN.match(name)
    if match is None:
        raise SheetNameError(f"Sheet name {name!r} carries no plan ID.")
    return match.group("plan_id")


def assert_sheet_names_unique(names: list[str] | tuple[str, ...]) -> None:
    """Raise ``SheetNameError`` if any sheet name repeats.

    Excel worksheet names are case-insensitive for uniqueness purposes, so
    the check folds case — two cases named ``M05 Foo`` and ``m05 foo`` would
    collide on the second ``Sheets.Add``.

    Parameters
    ----------
    names : list[str] | tuple[str, ...]
        Every sheet name the test-model artifact will claim.
    """
    seen: dict[str, str] = {}
    duplicates: list[str] = []
    for name in names:
        key = name.casefold()
        if key in seen:
            duplicates.append(f"{seen[key]!r} / {name!r}")
        else:
            seen[key] = name
    if duplicates:
        raise SheetNameError(
            "Duplicate test-model sheet name(s): " + ", ".join(duplicates)
        )
