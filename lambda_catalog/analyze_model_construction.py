"""Python expected values and QC reads for the Model Construction sheet.

Follows the repo's QC split: expected values are computed in Python from the
source xlsx (the ``analyze_regression_sheet`` pattern), sheet reads return
failure strings in the standard ``[Sheet] check=...`` format (the
``read_dummy_check_failures`` pattern). The calculator mirrors the sheet's
own semantics — ``Sample_Include()``'s role-aware mask, ``Dummy_Levels``'
mask-scoped level sets with the reference dropped, and the ``Predictor_Columns()`` /
``Constructed_Column_Names()`` iteration predicate — over the same typed
rows the Mileage Data sheet is built from, so every audit cell, spill
height, and level-qualified column name has an independently derived
expectation.

Two verification passes run against the open QC workbook:

1. **Default spec (T0)** — the build's shipped spec, untouched: audit strip
   (k, rows, response, responses, included rows, sequence flags), the header-strip /
   ``Predictor_Columns()`` twin tripwire, the Levels and Reference In Use display cells,
   the full-height contract on the K/L spills, and the filtered zones'
   heights and first values.
2. **Degenerate-Categorical via a Filter column** (the human test plan's T8
   mechanism, driven from the T0 base state): an ``Is_USA`` column is
   added to the data table and declared as a Filter, which collapses Origin
   to a single level inside the mask. Origin must then contribute zero
   columns — flagged by the Levels cell, never an error on the sheet —
   while every other zone recomputes against the narrower mask. The
   mutation is reverted afterwards, and the QC verify phase closes the
   workbook without saving as a backstop.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TypeGuard

import xlwings as xw

from .analyze_mileage import calculate_mileage_completeness_flags
from .workbook_builder import XL_CALCULATION_MANUAL, XL_CALCULATION_SEMIAUTOMATIC
from .write_sheet_csv_dataset import MILEAGE, load_csv_rows
from .write_sheet_model_construction import (
    _AUDIT_PAIRS,
    _AUDIT_ROW,
    _C_FILTERED_LABELS,
    _C_FILTERED_Y,
    _C_INCLUDED,
    _C_LEVELS,
    _C_MATRIX_LABELS,
    _C_MATRIX_START,
    _C_REF_IN_USE,
    _C_ROLE,
    _C_ROW_LABELS,
    _DEFAULT_SEQUENCE_VARIABLES,
    _DEFAULT_SPEC,
    _FALLBACK_SPEC,
    _FIRST_DATA_ROW,
    _HEADER_ROW,
    _ROLE_FILTER,
    _ROLE_IDENTIFIER,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
    _VARIABLES,
    SHEET_NAME,
)

DEFAULT_INPUT_CSV = MILEAGE.default_csv_path
FULL_DATA_HEADER = MILEAGE.full_data_header
DATA_SHEET_NAME = MILEAGE.sheet_name

_QC_PREFIX = "[Model Construction]"
_TABLE_NAME = "MileageData"

# The T8-mechanism Filter column: 1 for USA-built rows, 0 otherwise. Within
# the narrowed mask every remaining Origin value is 1 (USA), so Origin
# (Predictor/Categorical/TRUE in the default spec) degenerates to one level.
_EXTRA_FILTER_HEADER = "Is_USA"
_EXTRA_FILTER_FORMULA = '=--([@Origin]="US")'

# Read a few cells beyond every expected spill so an over-long spill is
# detected rather than silently truncated at the expected size.
_READ_MARGIN = 3


# ---------------------------------------------------------------------------
# Spec model and data loading
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class SpecVariable:
    """One spec row: the Role/Include/Type/Reference/Sequence declaration."""

    name: str
    role: str
    include: bool
    var_type: str
    # "" means "use the default (first sorted level)". A non-blank override
    # matches the underlying level's type — a numeric categorical column
    # (e.g. Origin, Model Year) takes an int/float reference, mirroring how
    # Excel auto-converts a typed numeric override to a number, not text.
    reference: object = ""
    sequence: bool = False
    # "None" (default) or "Log" — mirrors spec column G (Transform, v2.2).
    # Meaningful on a Response row or a Continuous Predictor row; ignored
    # on a Categorical Predictor (the sheet flags that combination red
    # rather than applying it).
    transform: str = "None"
    # Spec columns M / N (Interaction Term, Interaction Operation), wired
    # at v3.1. ``interaction_term`` names the OTHER operand by its column
    # name; ``interaction_operation`` is the closed Product | Difference |
    # Ratio vocabulary. Both blank (the default) means no interaction —
    # and either one blank is the same thing, matching the constructor's
    # mate(), which requires both.
    interaction_term: str = ""
    interaction_operation: str = ""


def build_default_spec() -> list[SpecVariable]:
    """Return the build's shipped T0 spec, one entry per table column.

    Derived from the sheet writer's own ``_DEFAULT_SPEC``/``_FALLBACK_SPEC``
    constants so the expectation side can never drift from what the build
    pre-fills.
    """
    spec: list[SpecVariable] = []
    for variable in _VARIABLES:
        role, include, var_type = _DEFAULT_SPEC.get(variable, _FALLBACK_SPEC)
        spec.append(
            SpecVariable(
                variable,
                role,
                include,
                var_type,
                sequence=variable in _DEFAULT_SEQUENCE_VARIABLES,
            )
        )
    return spec


def load_source_rows(csv_path: Path = DEFAULT_INPUT_CSV) -> list[dict[str, object]]:
    """Load the source CSV as row dicts matching the MileageData table.

    Cell values are typed exactly as the data sheet writer types them
    (int/float/str/None), and the computed ``Full_Data`` column is appended
    with the same completeness rule the sheet's ``Data_Completeness``
    formula applies.
    """
    headers, rows = load_csv_rows(csv_path, MILEAGE)
    flags = calculate_mileage_completeness_flags(csv_path)
    table_headers = [*headers, FULL_DATA_HEADER]
    return [
        dict(zip(table_headers, [*row, flag]))
        for row, flag in zip(rows, flags)
    ]


# ---------------------------------------------------------------------------
# Expectation calculation (pure Python — no Excel)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ModelConstructionExpectations:
    """Every sheet value the QC pass asserts, derived from CSV + spec."""

    total_rows: int
    included_rows: int
    k: int
    # Count of Sequence-flagged spec rows — the zero-or-one audit value the
    # H2 status line guards. Derived from the spec so the expectation tracks
    # whatever the shipped spec flags (Year in the T0 default).
    sequence_flags: int
    constructed_column_names: tuple[str, ...]
    response_name: str
    responses_count: int
    first_filtered_label: str
    first_filtered_response: object
    # Masked distinct level count per Predictor+Categorical spec row — the
    # sheet's Levels display column (H), including 1 for a degenerate column.
    level_counts: dict[str, int]
    # Reference level in effect per Predictor+Categorical spec row — the
    # sheet's Reference In Use display column (I): the explicit E value when
    # given, else the first sorted masked level, else "" (empty sample).
    references_in_use: dict[str, object]
    # Included Categorical Predictors contributing zero columns (degenerate
    # level set or invalid reference): the ISNA-skip variables.
    degenerate_categoricals: tuple[str, ...]


def _is_number(value: object) -> TypeGuard[int | float]:
    """ISNUMBER semantics: Excel booleans are not numbers."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _is_truthy_filter(value: object) -> bool:
    """The mask's Filter test ``N(IFERROR(N(col)=1,FALSE))``: TRUE or 1 pass."""
    return value is True or (_is_number(value) and float(value) == 1.0)


def _is_blank(value: object) -> bool:
    return value is None or value == ""


def _format_value(value: object) -> str:
    """Render a cell value the way Excel's ``&`` concatenation would."""
    if value is None:
        return ""
    if value is True:
        return "TRUE"
    if value is False:
        return "FALSE"
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def _compute_mask(
    spec: list[SpecVariable], rows: list[dict[str, object]]
) -> list[bool]:
    """Mirror ``Sample_Include()``: AND over role-derived row conditions."""
    filter_columns = [v.name for v in spec if v.role == _ROLE_FILTER]
    numeric_columns = [
        v.name
        for v in spec
        if v.role == _ROLE_RESPONSE
        or (v.role == _ROLE_PREDICTOR and v.include and v.var_type == "Continuous")
    ]
    return [
        all(_is_truthy_filter(row[name]) for name in filter_columns)
        and all(_is_number(row[name]) for name in numeric_columns)
        for row in rows
    ]


def _retained_levels(
    variable: SpecVariable,
    rows: list[dict[str, object]],
    mask: list[bool],
) -> list[object] | None:
    """Mirror ``Dummy_Levels``: sorted masked levels minus the reference.

    Returns None on any of the three NA() conditions — no masked non-blank
    values, an explicit reference not among the levels, or nothing left
    after dropping the reference — which is exactly the constructor's
    ISNA-skip signal.
    """
    values = {
        row[variable.name]
        for row, included in zip(rows, mask)
        if included and not _is_blank(row[variable.name])
    }
    if not values:
        return None
    levels = sorted(values, key=lambda v: (isinstance(v, str), v))
    reference = variable.reference if variable.reference != "" else levels[0]
    if reference not in levels:
        return None
    retained = [level for level in levels if level != reference]
    return retained or None


def resolve_interaction_operand(
    variable: SpecVariable, spec: Sequence[SpecVariable]
) -> int | None:
    """Mirror the constructor's ``mate()``: the operand's spec index, or None.

    None — meaning "this row contributes its main effect only" — covers every
    case the sheet flags rather than builds: a blank Interaction Term, a blank
    Interaction Operation, a name matching no source column, and an operand
    whose Role is not Predictor. An operand that IS a Predictor but has
    Include = FALSE resolves normally; that is the flagged-amber marginality
    case, which builds columns.
    """
    if not variable.interaction_term or not variable.interaction_operation:
        return None
    for index, candidate in enumerate(spec):
        if candidate.name == variable.interaction_term:
            return index if candidate.role == _ROLE_PREDICTOR else None
    return None


def block_column_names(
    variable: SpecVariable,
    rows: list[dict[str, object]],
    mask: list[bool],
) -> list[str] | None:
    """Mirror the constructor's ``blk()``: one spec row's own column headers.

    A Continuous row is a single column, relabelled ``Ln(name)`` under
    Transform = Log; a Categorical row is its reference-dropped level set.
    ``None`` mirrors the ``#N/A`` degenerate skip — the row contributes
    nothing at all, main effect and interaction alike.
    """
    if variable.var_type != "Categorical":
        name = variable.name
        return [f"Ln({name})" if variable.transform == "Log" else name]
    retained = _retained_levels(variable, rows, mask)
    if retained is None:
        return None
    return [f"{variable.name}: {_format_value(level)}" for level in retained]


def interaction_column_names(left: Sequence[str], right: Sequence[str]) -> list[str]:
    """Pairwise interaction headers, in the constructor's emission order.

    R's colon form over this library's own constructed names, so an
    interaction header always decomposes back into the two operand columns
    that produced it: ``GDP:Schooling``, or ``GDP:Status: Developing`` when
    one side is level-qualified.
    """
    return [f"{a}:{b}" for a in left for b in right]


def calculate_model_construction_expectations(
    spec: list[SpecVariable], rows: list[dict[str, object]]
) -> ModelConstructionExpectations:
    """Compute every asserted sheet value from the spec and the table rows."""
    mask = _compute_mask(spec, rows)
    included_rows = sum(mask)

    constructed: list[str] = []
    level_counts: dict[str, int] = {}
    references_in_use: dict[str, object] = {}
    degenerate: list[str] = []
    for variable in spec:
        is_categorical_predictor = (
            variable.role == _ROLE_PREDICTOR and variable.var_type == "Categorical"
        )
        if is_categorical_predictor:
            masked_values = {
                row[variable.name]
                for row, included in zip(rows, mask)
                if included and not _is_blank(row[variable.name])
            }
            level_counts[variable.name] = len(masked_values)
            # Mirror the Reference In Use display (I): echo an explicit E
            # verbatim; otherwise Dummy_Levels' default — the first sorted
            # masked level (numbers before text, matching Excel SORT).
            if variable.reference != "":
                references_in_use[variable.name] = variable.reference
            else:
                levels = sorted(
                    masked_values, key=lambda v: (isinstance(v, str), v)
                )
                references_in_use[variable.name] = levels[0] if levels else ""
        if variable.role != _ROLE_PREDICTOR or not variable.include:
            continue
        own = block_column_names(variable, rows, mask)
        if own is None:
            degenerate.append(variable.name)
            continue
        constructed.extend(own)
        # The row's interaction columns follow its own block, exactly as
        # Predictor_Columns() emits them. A degenerate operand contributes
        # nothing, which is the same skip the declaring row just passed.
        operand_index = resolve_interaction_operand(variable, spec)
        if operand_index is None:
            continue
        other = block_column_names(spec[operand_index], rows, mask)
        if other is None:
            continue
        constructed.extend(interaction_column_names(own, other))

    response_names = [v.name for v in spec if v.role == _ROLE_RESPONSE]
    response_name = response_names[0] if response_names else "(none)"

    identifiers = [v.name for v in spec if v.role == _ROLE_IDENTIFIER]
    first_included = next(
        (index for index, included in enumerate(mask) if included), None
    )
    if first_included is None:
        first_label = ""
        first_response: object = None
    else:
        first_row = rows[first_included]
        if identifiers:
            first_label = "|".join(
                _format_value(first_row[name]) for name in identifiers
            )
        else:
            first_label = f"Obs. {first_included + 1}"
        first_response = (
            first_row[response_name] if response_names else None
        )

    return ModelConstructionExpectations(
        total_rows=len(rows),
        included_rows=included_rows,
        k=len(constructed),
        sequence_flags=sum(1 for v in spec if v.sequence),
        constructed_column_names=tuple(constructed),
        response_name=response_name,
        responses_count=len(response_names),
        first_filtered_label=first_label,
        first_filtered_response=first_response,
        level_counts=level_counts,
        references_in_use=references_in_use,
        degenerate_categoricals=tuple(degenerate),
    )


# ---------------------------------------------------------------------------
# Sheet reads and comparison
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ModelConstructionObserved:
    """The sheet-side values the QC pass reads, one field per assertion."""

    audit_k: object
    audit_rows: object
    audit_response: object
    audit_responses: object
    audit_included: object
    audit_sequence_flags: object
    header_strip: tuple[object, ...]
    level_cells: dict[str, object]
    reference_cells: dict[str, object]
    row_labels_height: int
    mask_height: int
    mask_true_count: int
    filtered_labels_height: int
    filtered_response_height: int
    matrix_labels_height: int
    matrix_height: int
    matrix_width: int
    first_filtered_label: object
    y_header: object
    first_filtered_response: object


def _contiguous_height(values: list[object]) -> int:
    """Length of the leading non-blank run — a spill's observed extent."""
    height = 0
    for value in values:
        if _is_blank(value):
            break
        height += 1
    return height


def _read_column(
    sheet: xw.Sheet, column: int, first_row: int, count: int
) -> list[object]:
    values = sheet.range(
        (first_row, column), (first_row + count - 1, column)
    ).value
    return values if isinstance(values, list) else [values]


def _read_row(
    sheet: xw.Sheet, row: int, first_column: int, count: int
) -> list[object]:
    values = sheet.range(
        (row, first_column), (row, first_column + count - 1)
    ).value
    return values if isinstance(values, list) else [values]


def read_observed_values(
    sheet: xw.Sheet, total_rows: int, k_bound: int
) -> ModelConstructionObserved:
    """Read every asserted zone from the sheet.

    ``total_rows``/``k_bound`` size the reads only — every read extends
    ``_READ_MARGIN`` cells past the expected spill so an over-long spill
    shows up as a height/width mismatch instead of being clipped.
    """
    audit_value_columns = [value_col for _, value_col in _AUDIT_PAIRS]
    audit = [
        sheet.range((_AUDIT_ROW, column)).value for column in audit_value_columns
    ]

    strip = _read_row(
        sheet, _HEADER_ROW, _C_MATRIX_START, k_bound + _READ_MARGIN
    )
    header_strip = tuple(strip[: _contiguous_height(strip)])

    level_cells = {
        variable: sheet.range((_FIRST_DATA_ROW + offset, _C_LEVELS)).value
        for offset, variable in enumerate(_VARIABLES)
    }
    reference_cells = {
        variable: sheet.range((_FIRST_DATA_ROW + offset, _C_REF_IN_USE)).value
        for offset, variable in enumerate(_VARIABLES)
    }

    column_extent = total_rows + _READ_MARGIN
    row_labels = _read_column(sheet, _C_ROW_LABELS, _FIRST_DATA_ROW, column_extent)
    mask = _read_column(sheet, _C_INCLUDED, _FIRST_DATA_ROW, column_extent)
    filtered_labels = _read_column(
        sheet, _C_FILTERED_LABELS, _FIRST_DATA_ROW, column_extent
    )
    filtered_response = _read_column(
        sheet, _C_FILTERED_Y, _FIRST_DATA_ROW, column_extent
    )
    matrix_labels = _read_column(
        sheet, _C_MATRIX_LABELS, _FIRST_DATA_ROW, column_extent
    )
    matrix_first_column = _read_column(
        sheet, _C_MATRIX_START, _FIRST_DATA_ROW, column_extent
    )
    matrix_first_row = _read_row(
        sheet, _FIRST_DATA_ROW, _C_MATRIX_START, k_bound + _READ_MARGIN
    )

    return ModelConstructionObserved(
        audit_k=audit[0],
        audit_rows=audit[1],
        audit_response=audit[2],
        audit_responses=audit[3],
        audit_included=audit[4],
        audit_sequence_flags=audit[5],
        header_strip=header_strip,
        level_cells=level_cells,
        reference_cells=reference_cells,
        row_labels_height=_contiguous_height(row_labels),
        mask_height=len([v for v in mask if not _is_blank(v)]),
        mask_true_count=len([v for v in mask if v is True]),
        filtered_labels_height=_contiguous_height(filtered_labels),
        filtered_response_height=_contiguous_height(filtered_response),
        matrix_labels_height=_contiguous_height(matrix_labels),
        matrix_height=_contiguous_height(matrix_first_column),
        matrix_width=_contiguous_height(matrix_first_row),
        first_filtered_label=filtered_labels[0] if filtered_labels else None,
        y_header=sheet.range((_HEADER_ROW, _C_FILTERED_Y)).value,
        first_filtered_response=(
            filtered_response[0] if filtered_response else None
        ),
    )


def _values_match(expected: object, actual: object) -> bool:
    """Numeric-tolerant equality: Excel hands every number back as float."""
    if _is_number(expected) and _is_number(actual):
        return float(expected) == float(actual)
    return expected == actual


def compare_observed_to_expected(
    observed: ModelConstructionObserved,
    expected: ModelConstructionExpectations,
    context: str,
) -> list[str]:
    """Return one standard-format QC failure string per mismatch."""
    failures: list[str] = []

    def check(label: str, expected_value: object, actual_value: object) -> None:
        if not _values_match(expected_value, actual_value):
            failures.append(
                f"{_QC_PREFIX} [{context}] check={label!r}: "
                f"expected={expected_value!r}, excel_calc={actual_value!r}"
            )

    check("audit k", expected.k, observed.audit_k)
    check("audit rows", expected.total_rows, observed.audit_rows)
    check("audit response", expected.response_name, observed.audit_response)
    check("audit responses", expected.responses_count, observed.audit_responses)
    check("audit included rows", expected.included_rows, observed.audit_included)
    # Zero-or-one audit count, derived from the spec: the shipped T0 spec
    # flags Year, so this reads 1 (a QC config that clears H would read 0).
    check("audit sequence flags", expected.sequence_flags, observed.audit_sequence_flags)

    # Twin tripwire: the header strip is Constructed_Column_Names() and the
    # audit k is COLUMNS(Predictor_Columns()) — their widths must always agree.
    check(
        "header strip width == audit k (twin tripwire)",
        observed.audit_k,
        len(observed.header_strip),
    )
    check(
        "header strip names",
        expected.constructed_column_names,
        observed.header_strip,
    )

    for variable, count in expected.level_counts.items():
        check(f"Levels display for {variable}", count, observed.level_cells[variable])

    for variable, reference in expected.references_in_use.items():
        # Blank-normalize both sides: an Excel formula returning "" may read
        # back as None through COM, and the expected empty-sample value is "".
        observed_reference = observed.reference_cells[variable]
        check(
            f"Reference In Use display for {variable}",
            "" if _is_blank(reference) else reference,
            "" if _is_blank(observed_reference) else observed_reference,
        )

    check("Row_Labels() full height", expected.total_rows, observed.row_labels_height)
    check("Sample_Include() full height", expected.total_rows, observed.mask_height)
    check("Sample_Include() TRUE count", expected.included_rows, observed.mask_true_count)

    check("filtered labels rows", expected.included_rows, observed.filtered_labels_height)
    check("filtered response rows", expected.included_rows, observed.filtered_response_height)
    check("matrix labels rows", expected.included_rows, observed.matrix_labels_height)
    check("filtered matrix rows", expected.included_rows, observed.matrix_height)
    check("filtered matrix width", expected.k, observed.matrix_width)

    check("first filtered label", expected.first_filtered_label, observed.first_filtered_label)
    check("y header", f"y: {expected.response_name}", observed.y_header)
    check(
        "first filtered response",
        expected.first_filtered_response,
        observed.first_filtered_response,
    )

    return failures


# ---------------------------------------------------------------------------
# QC entry point
# ---------------------------------------------------------------------------

def _recalculate_sheets(workbook: xw.Book, *sheets: xw.Sheet) -> None:
    """Recalculate the given sheets in order under manual calculation."""
    workbook.app.api.Calculation = XL_CALCULATION_MANUAL
    try:
        for sheet in sheets:
            sheet.api.Calculate()
    finally:
        workbook.app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC


def _verify_degenerate_filter_case(
    workbook: xw.Book, rows: list[dict[str, object]]
) -> list[str]:
    """Drive the T8 mechanism: add an Is_USA Filter, assert the skip.

    The added ListColumn and the typed Filter role are reverted before
    returning; the verify phase additionally closes the workbook without
    saving, so the shipped workbook stays in its T0 state either way.
    """
    data_sheet = workbook.sheets[DATA_SHEET_NAME]
    sheet = workbook.sheets[SHEET_NAME]

    spec = build_default_spec() + [
        SpecVariable(_EXTRA_FILTER_HEADER, _ROLE_FILTER, False, "Continuous")
    ]
    mutated_rows = [
        {**row, _EXTRA_FILTER_HEADER: 1 if row["Origin"] == "US" else 0}
        for row in rows
    ]
    expected = calculate_model_construction_expectations(spec, mutated_rows)

    # The new table column's spec row appears right below the shipped ones.
    filter_role_row = _FIRST_DATA_ROW + len(_VARIABLES)
    table = data_sheet.api.ListObjects(_TABLE_NAME)
    column = table.ListColumns.Add()
    try:
        column.Name = _EXTRA_FILTER_HEADER
        column.DataBodyRange.Formula = _EXTRA_FILTER_FORMULA
        sheet.range((filter_role_row, _C_ROLE)).value = _ROLE_FILTER
        _recalculate_sheets(workbook, data_sheet, sheet)

        observed = read_observed_values(
            sheet, expected.total_rows, max(expected.k, 1)
        )
        failures = compare_observed_to_expected(
            observed, expected, "Is_USA filter"
        )
        # The point of the case: Origin degenerates inside the stratified
        # mask and must be skipped (zero contributed columns), not errored.
        if "Origin" not in expected.degenerate_categoricals:
            failures.append(
                f"{_QC_PREFIX} [Is_USA filter] check='Origin degenerates': "
                f"expected Origin in degenerate_categoricals, "
                f"got {expected.degenerate_categoricals!r}"
            )
    finally:
        sheet.range((filter_role_row, _C_ROLE)).clear_contents()
        column.Delete()
        _recalculate_sheets(workbook, data_sheet, sheet)

    return failures


def read_model_construction_failures(
    workbook: xw.Book, csv_path: Path = DEFAULT_INPUT_CSV
) -> list[str]:
    """Verify the Model Construction sheet; return QC failure messages.

    Pass 1 asserts the shipped default spec (the human test plan's T0
    state); pass 2 drives the degenerate-Categorical Filter case and
    reverts it. Every mismatch produces one standard-format failure string.
    """
    rows = load_source_rows(csv_path)
    spec = build_default_spec()
    expected = calculate_model_construction_expectations(spec, rows)

    sheet = workbook.sheets[SHEET_NAME]
    observed = read_observed_values(sheet, expected.total_rows, max(expected.k, 1))
    failures = compare_observed_to_expected(observed, expected, "default spec")

    failures.extend(_verify_degenerate_filter_case(workbook, rows))
    return failures
