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

from dataclasses import replace

import xlwings as xw

from .analyze_regression_guard_states import GuardStateExpected
from .analyze_model_construction import SpecVariable
from .analyze_regression_spec import RegressionSpecExpected
from .catalog_schema import CatalogFunction
from .regression_spec_sheet_io import (
    apply_sequence_period_overrides,
    apply_spec_case,
    set_prediction_inputs,
)
from .workbook_helpers import bold, val
from .write_sheet_model_construction import SPEC_DATASET_PROFILES, _ROLE_OMIT
from .write_sheet_regression import (
    _C_MODEL_CONTEXT,
    _C_MODEL_CONTEXT_LABEL,
    _C_MODEL_FORMULA,
    _C_MODEL_FORMULA_LABEL,
    _ROW_MODEL_CONTEXT_CHECK,
    _ROW_MODEL_FORMULA,
    write_regression_output_sheet,
)

# Where a generated sheet states what it is: the Model Context block's own
# two columns, BELOW the block and its health-check row with a blank row
# between.
#
# It was rows 1-2, which is inside the block — row 1 is its "MODEL CONTEXT"
# heading and row 2 is the first of the four cells `Fit_Context()` reads as a
# fixed range. Writing provenance text there replaced `Allow_Intercept` with a
# string, so every one of the ~30 engine call sites that takes `Fit_Context()`
# returned #VALUE!: Multiple R, R Square, Adjusted R Square, the whole ANOVA
# block, SS Total, Beta Weights, PRESS R², and the fit-space prediction
# outputs. The coefficients still computed, which is what made it look like a
# subtle numerical problem rather than two clobbered cells.
#
# Derived from the block's own constants so the two can never drift back into
# each other, and asserted below. The provenance runs AFTER
# write_regression_output_sheet, so any overlap with what that writer put on
# the sheet is a silent clobber, not a conflict anyone would see at build time.
_ROW_PROVENANCE_ID = _ROW_MODEL_CONTEXT_CHECK + 2
_ROW_PROVENANCE_COVERS = _ROW_MODEL_CONTEXT_CHECK + 3

# The provenance must sit strictly below the Model Context block, including
# its health-check row. A build that violates this produces a workbook whose
# engine is silently poisoned, so it fails at import rather than at Excel.
assert _ROW_PROVENANCE_ID > _ROW_MODEL_CONTEXT_CHECK, (
    "Provenance would overwrite the Model Context block that Fit_Context reads"
)

# ...and it must not land on the Model Formula readout, the other thing the
# Regression writer puts in this band. Today it cannot — the readout is on row
# 1 of the design-matrix zone, columns away from this block — but the two sets
# of constants are independent, so the disjointness is asserted rather than
# assumed. Cell-level, because the two no longer share a column to compare.
_PROVENANCE_CELLS = frozenset(
    (row, col)
    for row in (_ROW_PROVENANCE_ID, _ROW_PROVENANCE_COVERS)
    for col in (_C_MODEL_CONTEXT_LABEL, _C_MODEL_CONTEXT)
)
_MODEL_FORMULA_CELLS = frozenset(
    (_ROW_MODEL_FORMULA, col)
    for col in (_C_MODEL_FORMULA_LABEL, _C_MODEL_FORMULA)
)
assert not _PROVENANCE_CELLS & _MODEL_FORMULA_CELLS, (
    "Provenance would overwrite the Model Formula readout"
)

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


# Fixture columns this artifact adds to a shipped dataset's table, keyed by
# SPEC_DATASET_PROFILES key. `Is_USA` exists for M15, which declares it as a
# Filter to collapse Origin to one level.
#
# The legacy verifier adds such a column and deletes it again around the one
# case that needs it, because it runs against a shipped artifact that must not
# keep it. This workbook is itself a fixture, so the column is written once
# and left in place — which makes it part of EVERY case's source table, and
# therefore part of every case's spec block. See `effective_variables`.
FIXTURE_COLUMNS: dict[str, dict[str, str]] = {
    "auto_mpg": {"Is_USA": '=--([@Origin]="US")'},
}


def effective_variables(profile_key: str) -> tuple[str, ...]:
    """The source table's real column list: shipped columns plus fixtures.

    **A spec block must have exactly one row per Source_Table column.** Every
    constructor starts `n_c = COLUMNS(Source_Data)` and then indexes the
    `Spec_*` bands at `1..n_c`, so a table one row short of the data makes
    `INDEX(rl, n_c)` run off the end — the mask errors, and every engine cell
    downstream reads as an error rather than a number.

    That is not hypothetical: the first successful build of this workbook
    produced 74,065 mismatches for exactly this reason. `Is_USA` was added to
    `MileageData` for M15's benefit, which widened the table to 13 columns
    while every OTHER Auto MPG case still wrote a 12-row spec. M15 was the
    one Auto MPG sheet that worked, because its spec happens to declare the
    thirteenth column.
    """
    return (
        *SPEC_DATASET_PROFILES[profile_key].variables,
        *FIXTURE_COLUMNS.get(profile_key, {}),
    )


def pad_spec_to_source_table(
    spec: tuple[SpecVariable, ...], profile_key: str
) -> tuple[SpecVariable, ...]:
    """Append an ``Omit`` row for every source column the spec does not name.

    ``Omit`` is the right filler and not merely a convenient one: it
    contributes no design column and imposes no mask condition, so a padded
    spec fits exactly the same model as the unpadded one. The Python oracle
    never sees these rows — its loader returns only the shipped columns — and
    it does not need to, precisely because they change nothing.
    """
    declared = {variable.name for variable in spec}
    padding = tuple(
        SpecVariable(name, _ROLE_OMIT, False, "Continuous")
        for name in effective_variables(profile_key)
        if name not in declared
    )
    padded = (*spec, *padding)
    expected_width = len(effective_variables(profile_key))
    if len(padded) != expected_width:
        raise ValueError(
            f"Spec has {len(padded)} rows for a {expected_width}-column source "
            f"table. Every constructor indexes the Spec_* bands at "
            f"1..COLUMNS(Source_Data), so a mismatch makes the row mask error "
            f"and every engine cell read as an error. Declared but not in the "
            f"table: {sorted(declared - set(effective_variables(profile_key)))}"
        )
    return padded


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
    source_table_ref: str,
    sheet_notes: dict[str, str] | None,
    closures: tuple[CatalogFunction, ...] | None,
    shell_profile_key: str | None = None,
) -> xw.Sheet:
    """Write the Regression sheet layout under a per-case identity.

    ``shell_profile_key`` overrides which dataset's profile pre-fills the
    block, leaving ``source_table_ref`` to decide what the sheet actually
    reads. They agree for every case but one, which is the point: with them
    equal the retarget path is never exercised, because the block is always
    built for exactly the dataset it ends up pointed at. Setting this to a
    NARROWER dataset than ``source_table_ref`` reproduces what a user does
    by hand — edit one name and land on a table the block was not built for.
    See the ``spec_block_retarget_widens`` guard case.
    """
    profile_key = shell_profile_key or profile_key_for(source_table_ref)
    # The profile is built over the table's REAL column list — shipped
    # columns plus any fixture column this artifact added — so the shipped
    # DEFAULTS cover every Source_Table column. The block's height follows
    # COLUMNS(Source_Data) on its own; what the profile still governs is
    # which rows arrive pre-filled. See effective_variables.
    profile = replace(
        SPEC_DATASET_PROFILES[profile_key],
        variables=effective_variables(profile_key),
    )
    write_regression_output_sheet(
        workbook,
        sheet_notes,
        closures,
        source_table_ref=source_table_ref,
        spec_profile=profile,
        sheet_name=sheet_name,
        # Charts are the single biggest cost in this build — roughly a dozen
        # COM chart objects per sheet across ~48 sheets — and no oracle reads
        # one. Chart wiring is verified once, on the production Regression
        # sheet, by build_production.py.
        include_charts=False,
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
        source_table_ref=case.source_table_ref,
        sheet_notes=sheet_notes,
        closures=closures,
    )
    padded = _padded(expected)
    apply_spec_case(sheet, padded)
    if case.sequence_period is not None:
        # Type the case's Sequence Period into spec column I. Not optional
        # wiring: Base_Period_Delta() reads the TYPED value and returns #N/A
        # when the cell is blank, and the BFN panel Durbin-Watson cell passes
        # it as its delta. Skipping this would leave AE12 at #N/A while the
        # oracle held a real number — a guaranteed QC mismatch that looks
        # like a broken statistic and is really a blank input cell.
        #
        # Keyed on the Sequence-flagged row, the same row XMATCH resolves
        # inside Base_Period_Delta, so the sheet and the oracle read the
        # same declaration.
        apply_sequence_period_overrides(
            sheet,
            padded.case.spec,
            {
                item.name: case.sequence_period
                for item in case.spec
                if item.sequence
            },
        )
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
        source_table_ref=case.source_table_ref,
        sheet_notes=sheet_notes,
        closures=closures,
        shell_profile_key=case.shell_profile_key,
    )
    _apply_guard_spec(sheet, expected)
    _write_provenance(sheet, case.plan_id, case.name, case.covers)
    return sheet


def _padded(expected: RegressionSpecExpected) -> RegressionSpecExpected:
    """The same case with its spec padded to the source table's width.

    Returned as a replaced dataclass rather than by giving `apply_spec_case`
    an override parameter, so the one function that writes a spec onto a
    sheet stays identical for the builder and both verifiers — the whole
    reason it lives in regression_spec_sheet_io.
    """
    case = expected.case
    padded = pad_spec_to_source_table(
        case.spec, profile_key_for(case.source_table_ref)
    )
    if padded == case.spec:
        return expected
    return replace(expected, case=replace(case, spec=padded))


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

        spec = pad_spec_to_source_table(
            case.spec, profile_key_for(case.source_table_ref)
        )
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
