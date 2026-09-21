"""Every sheet address a static reference sheet cites names the cell it claims.

**The failure this closes.** The static reference sheets name sheet addresses as
literal text — the Diagnostic Guide's threshold table says "Col AT, Residual
Output" and the Model Comparison Guide says "Read Adjusted R Squared at Cell
AB7". Nothing computes those letters: they are typed into a content constant and
baked into ``templates/static_sheets.xlsx``, where no build re-derives them. So
a column insertion moves a statistic and leaves every citation behind, pointing
at whatever now occupies the old letter — a silent wrong answer in the one place
an analyst goes to find out where to look.

That has already happened twice in the Diagnostic Guide. Its own comment records
the first drift (AB/AC/AD/AG/AH, a whole zone out, after the Residual Output zone
moved to AN:BA) and the tolerance/GVIF citations had drifted two columns the same
way. This is the guard the comment asked for.

**What it asserts, in both directions.**

* Every address a writer cites is in the pin table below, and the citation equals
  the address the table derives from ``regression_layout`` / ``spec_layout``.
  An unrecognised citation fails rather than passing unnoticed.
* Every pin is cited by some writer. A pin nobody cites is a stale entry keeping
  the table looking healthy while it guards nothing — the same dead-entry rule
  ``test_static_template_freshness`` applies to its constant list.

**How the citations are found.** The content constants are the same ones the
freshness check pins (imported from it, so the two cannot disagree about where a
module's cell text lives). Each constant's string literals are scanned for the
two citation forms the sheets use — ``Cell AB7`` and ``Col AT``. A writer's
*interpolated* citations are not scanned: an f-string address is derived from a
layout constant at runtime and cannot drift, which is why the Diagnostic Guide's
three ``Cell`` citations are legitimately absent from the cell half of the table.

No Excel, no xlwings, no template read — pure AST plus the imported layout
constants, so it runs in the Linux CI job beside the other committed-artifact
checks.
"""
from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

from lambda_catalog import spec_layout as _sl
from lambda_catalog.regression_layout import (
    _A_ADJUSTED_R_SQUARED,
    _A_AIC,
    _A_AICC,
    _A_ALPHA,
    _A_BACK_TRANSFORM_METHOD,
    _A_BFN_DURBIN_WATSON,
    _A_BIC,
    _A_DESIGN_COLUMNS_TOTAL,
    _A_DURBIN_WATSON,
    _A_F_STATISTIC,
    _A_MEAN_LEVERAGE,
    _A_MULTIPLE_R,
    _A_OBSERVATIONS,
    _A_PRESS,
    _A_PRESS_R_SQUARED,
    _A_QQ_CORRELATION,
    _A_RESPONSE_READOUT,
    _A_RESPONSE_SPACE,
    _A_RESIDUAL_DF,
    _A_R_SQUARED,
    _A_SIGNIFICANCE_F,
    _A_SMEARING_FACTOR,
    _A_STANDARD_ERROR,
    _A_UNIT_ADJUSTED_R_SQUARED,
    _A_UNIT_RMSE,
    _A_UNIT_R_SQUARED,
    _C_AB,
    _C_AD,
    _C_AE,
    _C_AH,
    _C_AR,
    _C_AS,
    _C_AT,
    _C_AW,
    _C_AX,
    _C_X,
    _C_Y,
)
from tests.test_static_template_freshness import _CONTENT_CONSTANTS
from lambda_catalog.workbook_helpers import col_letter

ROOT_DIR = Path(__file__).resolve().parents[1]
_WRITER_DIR = ROOT_DIR / "lambda_catalog"

# ``Cell AB7, Regression Statistics`` / ``Col AT, Residual Output``. The column
# form takes a bare 1-3 letter token so ``in Col AD and Col AE`` is caught too;
# an accidental prose match fails the "unrecognised citation" assertion below
# rather than passing quietly.
_CELL_CITATION_RE = re.compile(r"\bCell ([A-Z]{1,3}[0-9]+)")
_COL_CITATION_RE = re.compile(r"\bCol ([A-Z]{1,3})\b")

# Address → the layout constant it must equal. The ``$``-stripped form of the
# constant is what a citation spells, since prose carries no absolute markers.
# Pin the address the reader is told to open, not a convenient neighbour: an
# entry here is a claim that this is where the statistic lives.
_CELL_PINS: dict[str, str] = {
    "AB5": _A_MULTIPLE_R,
    "AB6": _A_R_SQUARED,
    "AB7": _A_ADJUSTED_R_SQUARED,
    "AB8": _A_STANDARD_ERROR,
    "AB9": _A_OBSERVATIONS,
    "AB13": _A_ALPHA,
    "AB17": _A_RESIDUAL_DF,
    "AE5": _A_PRESS,
    "AE6": _A_PRESS_R_SQUARED,
    "AE7": _A_MEAN_LEVERAGE,
    "AE8": _A_AIC,
    "AE9": _A_BIC,
    "AE10": _A_AICC,
    "AE11": _A_QQ_CORRELATION,
    "AE12": _A_DURBIN_WATSON,
    "AE13": _A_BFN_DURBIN_WATSON,
    "AE16": _A_F_STATISTIC,
    "AF3": _A_RESPONSE_READOUT,
    "AF16": _A_SIGNIFICANCE_F,
    "AH5": _A_BACK_TRANSFORM_METHOD,
    "AH6": _A_SMEARING_FACTOR,
    "AH7": _A_UNIT_R_SQUARED,
    "AH8": _A_UNIT_ADJUSTED_R_SQUARED,
    "AH9": _A_UNIT_RMSE,
    "AH10": _A_RESPONSE_SPACE,
    # The spec block's own anchors, from spec_layout rather than
    # regression_layout: C2 is the intercept toggle, J2 the first of the three
    # Fixed Effects readouts, O1 the design-column total. I2 is the Spacing
    # Verdict — a status cell on the spec block's status row, in the column of
    # the spec field it is about, so it takes _C_SEQUENCE_PERIOD and the shared
    # status row rather than a row constant of its own.
    "C2": f"${col_letter(_sl._C_INCLUDE)}${_sl._INTERCEPT_ROW}",
    "I2": f"${col_letter(_sl._C_SEQUENCE_PERIOD)}${_sl._INTERCEPT_ROW}",
    "J2": f"${col_letter(_sl._C_PERIOD_IN_USE)}${_sl._INTERCEPT_ROW}",
    "O1": _A_DESIGN_COLUMNS_TOTAL,
}

# Column letter → the regression_layout constant naming that column.
_COLUMN_PINS: dict[str, int] = {
    "B": _sl._C_ROLE,
    "C": _sl._C_INCLUDE,
    "E": _sl._C_REFERENCE,
    "G": _sl._C_TRANSFORM,
    "I": _sl._C_SEQUENCE_PERIOD,
    "X": _C_X,
    "Y": _C_Y,
    "AB": _C_AB,
    "AD": _C_AD,
    "AE": _C_AE,
    "AH": _C_AH,
    "AR": _C_AR,
    "AS": _C_AS,
    "AT": _C_AT,
    "AW": _C_AW,
    "AX": _C_AX,
}

# The writers that cite sheet addresses in their content, declared rather than
# discovered. The Regression Instructions and Modeling Concepts sheets name
# their columns in prose ("column H", "the Include toggle") rather than by
# sheet address, so they carry nothing to check — but a writer that *starts*
# citing addresses has to be added here, and one that stops is caught too.
_ADDRESS_CITING_MODULES = frozenset(
    {
        "write_sheet_comparison_guide.py",
        "write_sheet_diagnostic_guide.py",
    }
)

# House floors: a detector that stopped matching would otherwise pass on an
# empty set. The Model Comparison Guide alone carries 29 cell and 16 column
# citations today; these sit below that and far above zero.
_MIN_CELL_CITATIONS = 20
_MIN_COLUMN_CITATIONS = 15


def _bare(address: str) -> str:
    """``$AB$7`` → ``AB7`` — the form a citation in prose spells."""
    return address.replace("$", "")


def _content_literals(module_name: str) -> str:
    """Every string literal in one writer's pinned content constants, joined."""
    source = (_WRITER_DIR / module_name).read_text(encoding="utf-8")
    tree = ast.parse(source)
    wanted = set(_CONTENT_CONSTANTS[module_name])
    texts: list[str] = []
    for node in tree.body:
        if not isinstance(node, (ast.Assign, ast.AnnAssign)):
            continue
        targets = node.targets if isinstance(node, ast.Assign) else [node.target]
        if len(targets) != 1 or not isinstance(targets[0], ast.Name):
            continue
        if targets[0].id not in wanted:
            continue
        for sub in ast.walk(node.value):
            if isinstance(sub, ast.Constant) and isinstance(sub.value, str):
                texts.append(sub.value)
    return "\n".join(texts)


def _citations(pattern: re.Pattern[str]) -> dict[str, set[str]]:
    """{citation: {module names citing it}} across every pinned writer."""
    found: dict[str, set[str]] = {}
    for module_name in _CONTENT_CONSTANTS:
        for citation in pattern.findall(_content_literals(module_name)):
            found.setdefault(citation, set()).add(module_name)
    return found


def _spelling_mismatch(citation: str, derived: str) -> str:
    return (
        f"cites {citation!r}, but the layout says that statistic is at "
        f"{derived!r} — a column or row moved and the citation did not follow"
    )


def test_every_cited_cell_address_matches_the_layout() -> None:
    """``Cell AB7`` in a reference sheet names AB7 and nothing else."""
    cited = _citations(_CELL_CITATION_RE)
    for citation, modules in sorted(cited.items()):
        assert citation in _CELL_PINS, (
            f"{sorted(modules)} cite Cell {citation}, which no pin covers — add "
            "the address and the regression_layout constant it derives from"
        )
        derived = _bare(_CELL_PINS[citation])
        assert citation == derived, f"{sorted(modules)} {_spelling_mismatch(citation, derived)}"


def test_every_cited_column_letter_matches_the_layout() -> None:
    """``Col AT, Residual Output`` names column AT and nothing else.

    The check that would have caught both Diagnostic Guide drifts: GVIF and
    Tolerance were cited as columns U and V while they live in X and Y.
    """
    cited = _citations(_COL_CITATION_RE)
    for citation, modules in sorted(cited.items()):
        assert citation in _COLUMN_PINS, (
            f"{sorted(modules)} cite Col {citation}, which no pin covers — add "
            "the letter and the regression_layout / spec_layout constant it "
            "derives from"
        )
        derived = col_letter(_COLUMN_PINS[citation])
        assert citation == derived, f"{sorted(modules)} {_spelling_mismatch(citation, derived)}"


def test_every_pin_is_cited_somewhere() -> None:
    """A pin nobody cites guards nothing — the dead-entry rule."""
    cited_cells = set(_citations(_CELL_CITATION_RE))
    cited_columns = set(_citations(_COL_CITATION_RE))
    unused_cells = sorted(set(_CELL_PINS) - cited_cells)
    unused_columns = sorted(set(_COLUMN_PINS) - cited_columns)
    assert not unused_cells, f"cell pins no reference sheet cites: {unused_cells}"
    assert not unused_columns, f"column pins no reference sheet cites: {unused_columns}"


def test_the_check_still_sees_work() -> None:
    """Floors, so a parser that stopped matching cannot pass quietly."""
    cells = len(_citations(_CELL_CITATION_RE))
    columns = len(_citations(_COL_CITATION_RE))
    assert cells >= _MIN_CELL_CITATIONS, f"only {cells} cell citations found"
    assert columns >= _MIN_COLUMN_CITATIONS, f"only {columns} column citations found"


def test_the_set_of_address_citing_writers_is_the_declared_one() -> None:
    """A writer that starts or stops citing addresses is a conscious change.

    Without this, a new static sheet could cite addresses that no pin covers
    and the two checks above would still pass — they only ever look at what
    they find.
    """
    observed = {
        module_name
        for module_name in _CONTENT_CONSTANTS
        if _CELL_CITATION_RE.findall(_content_literals(module_name))
        or _COL_CITATION_RE.findall(_content_literals(module_name))
    }
    assert observed == set(_ADDRESS_CITING_MODULES), (
        "address-citing writers changed: "
        f"newly citing {sorted(observed - _ADDRESS_CITING_MODULES)}, "
        f"no longer citing {sorted(set(_ADDRESS_CITING_MODULES) - observed)}"
    )


@pytest.mark.parametrize("module_name", sorted(_ADDRESS_CITING_MODULES), ids=str)
def test_every_address_citing_writer_contributes_citations(module_name: str) -> None:
    """Each declared writer is actually covered by the pin checks."""
    blob = _content_literals(module_name)
    assert _CELL_CITATION_RE.findall(blob) or _COL_CITATION_RE.findall(blob), (
        f"{module_name} cites no sheet address, so nothing above checks it"
    )
