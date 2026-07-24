"""QC verification of the Regression sheet's spec block and constructor path.

The v3.0 changeover moved the declarative spec block onto the Regression
sheet, and the Model Construction sheet (with its dedicated verifier) left
the build. The six regression QC configurations only exercise all-continuous
designs — every config switches the default spec's Categorical predictors
(Model Year, Origin) OFF — so without this module the categorical
constructor path (dummy encoding, level-qualified names, the degenerate
skip, the Levels and Reference In Use displays) would have no live-Excel
verification at all.

This is the Model Construction verifier ported to the Regression sheet. The
expectation side is reused verbatim (``analyze_model_construction``'s
calculator is pure Python and sheet-agnostic), and the spec block occupies
identical coordinates on both sheets (columns A–K; intercept row 2, headers
row 3, spec rows 4–26 — the writers are shared), so the Levels / Reference
In Use reads port unchanged. Only the assertions against the Model
Construction sheet's display zones are remapped onto the Regression sheet's
own display of the same facts:

    MC audit k / twin tripwire   → N3 constructed-names spill width, and the
                                   V21 coefficient-label spill (k+1 rows with
                                   the default intercept ON)
    MC header strip names        → N3 spill contents (vertical)
    MC audit response            → AA2 Predicted Variable readout
    MC audit included rows       → W8 Observations cell
    MC first filtered label      → AI3 (first residual-output row label)
    MC filtered zone heights     → AI3 spill height
    MC audit sequence flags      → H2 Sequence status line (blank while the
                                   spec carries zero-or-one flags)

The full-height contract, filtered-y column, y-header, and responses-count
assertions are dropped: the Regression sheet displays none of them, and the
closures behind them are pinned by the unit suite.

Two passes, mirroring the retired verifier:

1. **Default spec (T0)** — the build's shipped state, Model Year/Origin ON.
2. **Degenerate Categorical via a Filter column** (the human test plan's T8
   mechanism): an ``Is_USA`` column is added to the data table and
   declared as a Filter in the spec, collapsing Origin to one level inside
   the mask; Origin must contribute zero columns, not an error. The mutation
   is reverted, and the QC verify phase closes the workbook without saving
   as a backstop.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import xlwings as xw

from .analyze_model_construction import (
    _EXTRA_FILTER_FORMULA,
    _EXTRA_FILTER_HEADER,
    _READ_MARGIN,
    _TABLE_NAME,
    ModelConstructionExpectations,
    SpecVariable,
    _contiguous_height,
    _is_blank,
    _read_column,
    _recalculate_sheets,
    _values_match,
    build_default_spec,
    calculate_model_construction_expectations,
    load_source_rows,
)
from .write_sheet_mileage_data import (
    DEFAULT_XLSX_PATH,
    SHEET_NAME as DATA_SHEET_NAME,
)
from .write_sheet_model_construction import (
    _C_LEVELS,
    _C_REF_IN_USE,
    _C_ROLE,
    _C_SEQUENCE,
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
    _INTERCEPT_ROW as _SEQUENCE_STATUS_ROW,
    _ROLE_FILTER,
    _VARIABLES,
)
from .write_sheet_regression import (
    REGRESSION_SHEET_NAME,
    _C_AC,
    _C_AK,
    _C_P,
    _C_X,
    _C_Y,
)

_QC_PREFIX = "[Regression Spec]"

# Row anchors on the Regression sheet (1-based; must match
# write_sheet_regression.py's section writers).
_ROW_RESPONSE_READOUT = 2   # AC2 = derived response name
_ROW_NAMES_SPILL = 3        # P3 = TRANSPOSE(Constructed_Column_Names())
_ROW_RESID_FIRST = 3        # AK3 = FILTER(Row_Labels(), Sample_Include())
_ROW_OBSERVATIONS = 8       # Y8 = Observations(...)
_ROW_COEFF_FIRST = 21       # X21 = coefficient label spill (k+1 with intercept)


@dataclass(frozen=True)
class RegressionSpecObserved:
    """The Regression-sheet values the spec-block QC pass reads."""

    constructed_names: tuple[object, ...]
    coeff_label_height: int
    response_readout: object
    observations_cell: object
    level_cells: dict[str, object]
    reference_cells: dict[str, object]
    resid_labels_height: int
    first_resid_label: object
    sequence_status: object


def read_observed_spec_values(
    sheet: xw.Sheet, total_rows: int, k_bound: int
) -> RegressionSpecObserved:
    """Read every asserted zone from the Regression sheet.

    ``total_rows``/``k_bound`` size the reads only — every read extends
    ``_READ_MARGIN`` cells past the expected spill so an over-long spill
    shows up as a height/width mismatch instead of being clipped.
    """
    names = _read_column(
        sheet, _C_P, _ROW_NAMES_SPILL, k_bound + _READ_MARGIN
    )
    constructed_names = tuple(names[: _contiguous_height(names)])

    coeff_labels = _read_column(
        sheet, _C_X, _ROW_COEFF_FIRST, k_bound + 1 + _READ_MARGIN
    )

    level_cells = {
        variable: sheet.range((_SPEC_FIRST_DATA_ROW + offset, _C_LEVELS)).value
        for offset, variable in enumerate(_VARIABLES)
    }
    reference_cells = {
        variable: sheet.range(
            (_SPEC_FIRST_DATA_ROW + offset, _C_REF_IN_USE)
        ).value
        for offset, variable in enumerate(_VARIABLES)
    }

    resid_labels = _read_column(
        sheet, _C_AK, _ROW_RESID_FIRST, total_rows + _READ_MARGIN
    )

    return RegressionSpecObserved(
        constructed_names=constructed_names,
        coeff_label_height=_contiguous_height(coeff_labels),
        response_readout=sheet.range((_ROW_RESPONSE_READOUT, _C_AC)).value,
        observations_cell=sheet.range((_ROW_OBSERVATIONS, _C_Y)).value,
        level_cells=level_cells,
        reference_cells=reference_cells,
        resid_labels_height=_contiguous_height(resid_labels),
        first_resid_label=resid_labels[0] if resid_labels else None,
        sequence_status=sheet.range(
            (_SEQUENCE_STATUS_ROW, _C_SEQUENCE)
        ).value,
    )


def compare_spec_observed_to_expected(
    observed: RegressionSpecObserved,
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

    # The constructor twin, displayed: the M3 spill IS
    # Constructed_Column_Names(), so content equality pins both the width
    # (k) and the level-qualified names in one assertion.
    check(
        "constructed names (M spill)",
        expected.constructed_column_names,
        observed.constructed_names,
    )
    # Twin tripwire against the engine side: the V21 coefficient-label spill
    # is VSTACK("Intercept", names...) under the shipped intercept-ON
    # default, so its height must be exactly k+1.
    check(
        "coefficient label rows == k+1 (twin tripwire)",
        expected.k + 1,
        observed.coeff_label_height,
    )

    check("response readout", expected.response_name, observed.response_readout)
    check("Observations cell", expected.included_rows, observed.observations_cell)

    for variable, count in expected.level_counts.items():
        check(
            f"Levels display for {variable}",
            count,
            observed.level_cells[variable],
        )

    for variable, reference in expected.references_in_use.items():
        # Blank-normalize both sides: an Excel formula returning "" may read
        # back as None through COM, and the expected empty-sample value is "".
        observed_reference = observed.reference_cells[variable]
        check(
            f"Reference In Use display for {variable}",
            "" if _is_blank(reference) else reference,
            "" if _is_blank(observed_reference) else observed_reference,
        )

    check(
        "residual label rows",
        expected.included_rows,
        observed.resid_labels_height,
    )
    check(
        "first residual label",
        expected.first_filtered_label,
        observed.first_resid_label,
    )

    # The Sequence status line (H2) errors only at two-plus flags. The shipped
    # spec carries exactly one (Year), which is valid, so the zero-or-one
    # validation must still render blank, not an error.
    check(
        "sequence status line blank",
        "",
        "" if _is_blank(observed.sequence_status) else observed.sequence_status,
    )

    return failures


def _verify_degenerate_filter_case(
    workbook: xw.Book, rows: list[dict[str, object]]
) -> list[str]:
    """Drive the T8 mechanism: add an Is_USA Filter, assert the skip.

    The added ListColumn and the typed Filter role are reverted before
    returning; the verify phase additionally closes the workbook without
    saving, so the shipped workbook stays in its T0 state either way.
    """
    data_sheet = workbook.sheets[DATA_SHEET_NAME]
    sheet = workbook.sheets[REGRESSION_SHEET_NAME]

    spec = build_default_spec() + [
        SpecVariable(_EXTRA_FILTER_HEADER, _ROLE_FILTER, False, "Continuous")
    ]
    mutated_rows = [
        {**row, _EXTRA_FILTER_HEADER: 1 if row["Origin"] == 1 else 0}
        for row in rows
    ]
    expected = calculate_model_construction_expectations(spec, mutated_rows)

    # The new table column's spec row appears right below the shipped ones.
    filter_role_row = _SPEC_FIRST_DATA_ROW + len(_VARIABLES)
    table = data_sheet.api.ListObjects(_TABLE_NAME)
    column = table.ListColumns.Add()
    try:
        column.Name = _EXTRA_FILTER_HEADER
        column.DataBodyRange.Formula = _EXTRA_FILTER_FORMULA
        sheet.range((filter_role_row, _C_ROLE)).value = _ROLE_FILTER
        _recalculate_sheets(workbook, data_sheet, sheet)

        observed = read_observed_spec_values(
            sheet, expected.total_rows, max(expected.k, 1)
        )
        failures = compare_spec_observed_to_expected(
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


def read_regression_spec_block_failures(
    workbook: xw.Book, xlsx_path: Path = DEFAULT_XLSX_PATH
) -> list[str]:
    """Verify the Regression sheet's spec block; return QC failure messages.

    Pass 1 asserts the shipped default spec (the human test plan's T0 state,
    Model Year/Origin ON — the only live-Excel coverage of the categorical
    constructor path); pass 2 drives the degenerate-Categorical Filter case
    and reverts it. Must run while the spec block is in its shipped state —
    i.e. BEFORE the six-configuration regression pass, which mutates the
    Include cells.
    """
    rows = load_source_rows(xlsx_path)
    spec = build_default_spec()
    expected = calculate_model_construction_expectations(spec, rows)

    sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    observed = read_observed_spec_values(
        sheet, expected.total_rows, max(expected.k, 1)
    )
    failures = compare_spec_observed_to_expected(
        observed, expected, "default spec"
    )

    failures.extend(_verify_degenerate_filter_case(workbook, rows))
    return failures
