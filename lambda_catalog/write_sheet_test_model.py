"""Materialize one test-model case as its own Regression-shaped worksheet.

The regression test-model suite has always been a list of specs pushed
through a single sheet one at a time. That works, but it means a case exists
only as a log line: a failure says "expected 0.79596, got 0.79601" with
nothing to open. It also serializes the whole suite behind one shared sheet,
and forces every case to defensively re-set every input in case the previous
case left something behind.

This writer builds each case ONCE onto its own sheet, so the artifact is a
workbook of ~48 fully-computed models a human can page through, and the
verifier's job collapses to reading — no writing, no per-case
recalculation, no state to leak.

Nothing here reimplements the Regression sheet. It calls
``write_regression_output_sheet`` with a per-case sheet name, charts off and
a per-sheet spec-table name, then pushes the case's spec through the same
``regression_spec_sheet_io`` helpers the legacy verifier uses. If the two
ever disagreed about what a case is, the test-model sheets would be
verifying something other than what the QC harness fits.
"""
from __future__ import annotations

import xlwings as xw

from .analyze_regression_guard_states import GuardStateExpected
from .analyze_regression_spec import RegressionSpecExpected
from .catalog_schema import CatalogFunction
from .regression_spec_sheet_io import (
    apply_sequence_period_overrides,
    apply_spec_case,
    set_prediction_inputs,
)
from .test_model_sheets import spec_table_name
from .workbook_helpers import bold, val
from .write_sheet_model_construction import SPEC_DATASET_PROFILES
from .write_sheet_regression import (
    _C_MODEL_CONTEXT,
    _C_MODEL_CONTEXT_LABEL,
    write_regression_output_sheet,
)

# Where a generated sheet states what it is. The Model Context block's own
# label/value column pair, four rows above its first element — inside the
# §4b materialization band, well right of every zone a reader scrolls
# through, so the provenance never displaces sheet content.
_ROW_PROVENANCE_ID = 1
_ROW_PROVENANCE_COVERS = 2

# Source_Table ref -> SPEC_DATASET_PROFILES key. The profile decides how many
# rows the spec table gets, and a table sized for Auto MPG's 12 columns would
# silently drop the last 11 rows of a 23-column Life Expectancy spec — the
# structured references simply would not reach them. Deriving the profile
# from the case's own Source_Table retarget is what keeps the two in sync;
# they are independent parameters of the writer by design (see its docstring)
# and something has to pair them.
_PROFILE_BY_SOURCE_TABLE = {
    profile.source_table_ref: key
    for key, profile in SPEC_DATASET_PROFILES.items()
}


def profile_key_for(source_table_ref: str) -> str:
    """Return the SPEC_DATASET_PROFILES key a case's Source_Table implies.

    Raises rather than defaulting to Auto MPG: a case pointed at a dataset
    with no registered profile would get a spec table sized for the wrong
    column count, which fails as wrong NUMBERS rather than as an error.
    """
    try:
        return _PROFILE_BY_SOURCE_TABLE[source_table_ref]
    except KeyError:
        raise ValueError(
            f"No SPEC_DATASET_PROFILES entry targets {source_table_ref!r}. "
            "A test-model case needs one so its spec table is sized to the "
            "dataset's column count; add the profile in "
            "write_sheet_model_construction.py."
        ) from None


def _write_provenance(sheet: xw.Sheet, plan_id: str, name: str, covers: str) -> None:
    """Label the sheet with its plan ID, case name, and the corner it covers.

    A generated sheet is worthless as a debugging surface if a reader has to
    cross-reference a Python registry to find out why it exists. These three
    cells make an opened tab self-describing, and they are plain cell writes
    so they are exercised without Excel in the unit suite.
    """
    val(sheet, _ROW_PROVENANCE_ID, _C_MODEL_CONTEXT_LABEL, "Test Model")
    bold(sheet, _ROW_PROVENANCE_ID, _C_MODEL_CONTEXT_LABEL)
    val(sheet, _ROW_PROVENANCE_ID, _C_MODEL_CONTEXT, f"{plan_id} — {name}")
    val(sheet, _ROW_PROVENANCE_COVERS, _C_MODEL_CONTEXT_LABEL, "Covers")
    bold(sheet, _ROW_PROVENANCE_COVERS, _C_MODEL_CONTEXT_LABEL)
    val(sheet, _ROW_PROVENANCE_COVERS, _C_MODEL_CONTEXT, covers)


def _write_shell(
    workbook: xw.Book,
    *,
    sheet_name: str,
    plan_id: str,
    source_table_ref: str,
    sheet_notes: dict[str, str] | None,
    closures: tuple[CatalogFunction, ...] | None,
) -> xw.Sheet:
    """Write the Regression sheet layout under a per-case identity."""
    write_regression_output_sheet(
        workbook,
        sheet_notes,
        closures,
        source_table_ref=source_table_ref,
        spec_profile=SPEC_DATASET_PROFILES[profile_key_for(source_table_ref)],
        sheet_name=sheet_name,
        # Charts are the single biggest cost in this build — roughly a dozen
        # COM chart objects per sheet across ~48 sheets — and no oracle reads
        # one. Chart wiring is verified once, on the production Regression
        # sheet, by build_production.py.
        include_charts=False,
        spec_table_name=spec_table_name(plan_id),
    )
    return workbook.sheets[sheet_name]


def write_test_model_sheet(
    workbook: xw.Book,
    expected: RegressionSpecExpected,
    sheet_notes: dict[str, str] | None = None,
    closures: tuple[CatalogFunction, ...] | None = None,
) -> xw.Sheet:
    """Build one fittable case's sheet, spec applied and inputs prefilled.

    The prediction inputs are prefilled to each design column's training
    mean, matching what the legacy verifier writes before reading the
    Prediction Interval box — so the box on a generated sheet holds a
    comparable number rather than whatever the shipped default produces.
    """
    case = expected.case
    sheet = _write_shell(
        workbook,
        sheet_name=case.sheet_name,
        plan_id=case.plan_id,
        source_table_ref=case.source_table_ref,
        sheet_notes=sheet_notes,
        closures=closures,
    )
    apply_spec_case(sheet, expected)
    set_prediction_inputs(
        sheet,
        expected.results.prediction_interval.pred_input_values,
        expected.design.constructed_column_transforms,
    )
    _write_provenance(
        sheet,
        case.plan_id,
        case.name,
        expected.results.unit_space.model_formula,
    )
    return sheet


def write_guard_state_sheet(
    workbook: xw.Book,
    expected: GuardStateExpected,
    sheet_notes: dict[str, str] | None = None,
    closures: tuple[CatalogFunction, ...] | None = None,
) -> xw.Sheet:
    """Build one guard-rail configuration's sheet.

    Differs from a fittable case's sheet in three ways, all consequences of
    a guard state not being a model: there is no oracle-computed prediction
    input to prefill (several of these specs have no fittable model at all),
    the spec is written directly rather than through a
    ``RegressionSpecExpected``, and typed Sequence Period overrides are
    applied — spec column I is an input, and it is the mechanism M16 and P07
    exist to exercise.
    """
    case = expected.case
    sheet = _write_shell(
        workbook,
        sheet_name=case.sheet_name,
        plan_id=case.plan_id,
        source_table_ref=case.source_table_ref,
        sheet_notes=sheet_notes,
        closures=closures,
    )
    _apply_guard_spec(sheet, expected)
    _write_provenance(sheet, case.plan_id, case.name, case.covers)
    return sheet


def _apply_guard_spec(sheet: xw.Sheet, expected: GuardStateExpected) -> None:
    """Write a guard case's spec block, intercept toggle and Δ overrides.

    A thin adapter over ``apply_spec_case``: that function takes a
    ``RegressionSpecExpected`` because it also needs the resolved prediction
    group and back-transform method, neither of which a guard case has. The
    adapter presents the three fields it does read rather than duplicating
    ~40 lines of cell writes, so the guard sheets and the model sheets stay
    written by one piece of code.
    """
    case = expected.case

    class _SpecView:  # pylint: disable=too-few-public-methods
        """The subset of RegressionSpecCase that apply_spec_case reads."""

        spec = case.spec
        allow_intercept = case.allow_intercept
        source_table_ref = case.source_table_ref
        back_transform = "Duan"

    class _ExpectedView:  # pylint: disable=too-few-public-methods
        case = _SpecView()
        # A guard case declares no prediction group. "(all)" is the constant
        # group a no-Fixed-Effects spec resolves to, and is what the sheet's
        # own $AK$12 default computes for one — writing it explicitly keeps
        # the cell from carrying a formula whose value depends on build order.
        resolved_prediction_group = "(all)"

    apply_spec_case(sheet, _ExpectedView())  # type: ignore[arg-type]
    apply_sequence_period_overrides(
        sheet, case.spec, case.sequence_period_override
    )
