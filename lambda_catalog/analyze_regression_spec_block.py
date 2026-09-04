"""QC verification of the Regression sheet's spec block and constructor path.

The declarative spec block lives on the Regression sheet, and
``analyze_model_construction`` (pure Python, sheet-agnostic) supplies the
expectation side. The six regression QC configurations that run beside this
one only exercise all-continuous designs — every config switches the
shipped spec's Categorical predictors OFF — so without this module nothing
would check the shipped workbook's own spec against an independent
expectation at all.

The expectation side is reused verbatim (``analyze_model_construction``'s
calculator is pure Python and sheet-agnostic), and the spec block occupies
fixed coordinates on the Regression sheet (columns A–O; intercept row 2,
headers row 3, the spec data rows below that — the writers are shared), so
the Levels / Reference In Use reads port unchanged. Only the assertions
against the Regression sheet's display zones are retained:

    MC audit k / twin tripwire   → S4 constructed-names spill width, and the
                                   AA22 coefficient-label spill (k+1 rows with
                                   the default intercept ON)
    MC header strip names        → S4 spill contents (vertical)
    MC audit response            → AF3 Predicted Variable readout
    MC audit included rows       → AB9 Observations cell
    MC first filtered label      → AN4 (first residual-output row label)
    MC filtered zone heights     → AN4 spill height
    MC audit sequence flags      → H2 Sequence status line (blank while the
                                   spec carries zero-or-one flags)

Every coordinate above is imported from ``write_sheet_regression`` rather
than spelled out, and pinned in tests/test_analyze_regression_spec_block.py.

The full-height contract, filtered-y column, y-header, and responses-count
assertions are dropped: the Regression sheet displays none of them, and the
closures behind them are pinned by the unit suite.

**This verifier checks whatever the workbook actually ships.** It resolves the
dataset by reading the live ``Source_Table`` name off the Regression sheet and
matching it to a registered profile, then builds its expectations from that
profile. The shipped default is a build-time choice
(``build_production.py --regression-dataset``), so a verifier that hardcoded
one dataset would compare the wrong one whenever the default flipped;
deriving the dataset from the artifact makes that unrepresentable.

**The categorical corners live in the test-model suite.** This verifier's
job is the SHIPPED artifact, so it carries no synthetic filter column and
no extra categorical pass: those describe configurations the shipped
workbook is not in, and reproducing them here would mean retargeting
``Source_Table`` on the production sheet mid-verify. They are covered on
their own sheets by ``build_test_models.py --verify``:

    multi-level dummies, level-qualified names   M01, M09, M14, M14b
    degenerate Categorical via a Filter column   M15
    Levels / Reference In Use displays           every categorical case

What remains here is the one thing only this module can check: that the
workbook a user opens computes its own shipped spec correctly.
"""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

import xlwings as xw

from .analyze_life_expectancy import load_life_expectancy_source_rows
from .analyze_model_construction import (
    _READ_MARGIN,
    ModelConstructionExpectations,
    _contiguous_height,
    _is_blank,
    _read_column,
    _values_match,
    build_profile_spec,
    calculate_model_construction_expectations,
    load_source_rows,
)
from .analyze_production_lots import load_production_lots_source_rows
from .write_sheet_csv_dataset import LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS
from .write_spec_block import (
    _C_LEVELS,
    _C_REF_IN_USE,
    _C_SEQUENCE,
    SPEC_DATASET_PROFILES,
    SpecDatasetProfile,
)
from .write_spec_block import (
    _FIRST_DATA_ROW as _SPEC_FIRST_DATA_ROW,
)
from .write_spec_block import (
    _INTERCEPT_ROW as _SEQUENCE_STATUS_ROW,
)
from .write_sheet_regression import (
    _C_AA,
    _C_AB,
    _C_AF,
    _C_AN,
    _C_S,
    REGRESSION_SHEET_NAME,
)

DEFAULT_INPUT_CSV = MILEAGE.default_csv_path
DATA_SHEET_NAME = MILEAGE.sheet_name

_QC_PREFIX = "[Regression Spec]"

# Every dataset the spec block can be pointed at, keyed by the profile key
# SPEC_DATASET_PROFILES uses. The row loader is per-dataset because each CSV
# needs its own normalization (header casing, an optional appended derived
# column — Life Expectancy appends "Developed Country after 2013"; the
# others append none) — and all three share one signature:
# (csv_path) -> list[dict].
_DATASET_LOADERS = {
    "auto_mpg": (MILEAGE, load_source_rows),
    "life_expectancy": (LIFE_EXPECTANCY, load_life_expectancy_source_rows),
    "production_lots": (PRODUCTION_LOTS, load_production_lots_source_rows),
}


def _resolve_shipped_profile(
    sheet: xw.Sheet,
) -> tuple[str, SpecDatasetProfile] | None:
    """Which dataset is this Regression sheet actually pointed at?

    ``Source_Table`` is the ONE name that decides which table the spec block
    reads, so the sheet's own copy of it is the authoritative answer — not the
    build flag that produced the file, which a verifier handed an existing
    workbook does not have.

    Returns ``(profile_key, profile)``, or ``None`` when the name is missing
    or targets a table with no registered profile. The caller turns that into
    a QC failure rather than guessing: comparing against the wrong dataset is
    exactly the failure mode this resolution exists to prevent, and silently
    falling back to a default would reintroduce it.
    """
    try:
        refers_to = str(sheet.api.Names.Item("Source_Table").RefersTo)
    except Exception:  # pylint: disable=broad-exception-caught
        return None

    # Excel may hand back the ref sheet-qualified or requoted; matching on the
    # table name rather than the whole string keeps this robust to that.
    for key, profile in SPEC_DATASET_PROFILES.items():
        config, _loader = _DATASET_LOADERS[key]
        if config.table_name.casefold() in refers_to.casefold():
            return key, profile
    return None

# Row anchors on the Regression sheet (1-based; must match
# write_sheet_regression.py's section writers).
_ROW_RESPONSE_READOUT = 3   # AF3 = derived response name
_ROW_NAMES_SPILL = 4        # S4 = TRANSPOSE(Constructed_Column_Names())
_ROW_RESID_FIRST = 4        # AN4 = FILTER(Row_Labels(), Sample_Include())
_ROW_OBSERVATIONS = 9       # AB9 = Observations(...)
_ROW_COEFF_FIRST = 22       # AA22 = coefficient label spill (k+1 with intercept)


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
    sheet: xw.Sheet,
    total_rows: int,
    k_bound: int,
    variables: Sequence[str],
) -> RegressionSpecObserved:
    """Read every asserted zone from the Regression sheet.

    ``total_rows``/``k_bound`` size the reads only — every read extends
    ``_READ_MARGIN`` cells past the expected spill so an over-long spill
    shows up as a height/width mismatch instead of being clipped.

    ``variables`` is the resolved dataset's column list, in spec-block row
    order, and it is a PARAMETER for the same reason the expectations are:
    the Levels and Reference In Use cells are read POSITIONALLY, so the
    column list must match the resolved dataset — iterating a list for the
    wrong dataset would read the wrong rows and report every categorical
    display as ``None``.
    """
    names = _read_column(
        sheet, _C_S, _ROW_NAMES_SPILL, k_bound + _READ_MARGIN
    )
    constructed_names = tuple(names[: _contiguous_height(names)])

    coeff_labels = _read_column(
        sheet, _C_AA, _ROW_COEFF_FIRST, k_bound + 1 + _READ_MARGIN
    )

    level_cells = {
        variable: sheet.range((_SPEC_FIRST_DATA_ROW + offset, _C_LEVELS)).value
        for offset, variable in enumerate(variables)
    }
    reference_cells = {
        variable: sheet.range(
            (_SPEC_FIRST_DATA_ROW + offset, _C_REF_IN_USE)
        ).value
        for offset, variable in enumerate(variables)
    }

    resid_labels = _read_column(
        sheet, _C_AN, _ROW_RESID_FIRST, total_rows + _READ_MARGIN
    )

    return RegressionSpecObserved(
        constructed_names=constructed_names,
        coeff_label_height=_contiguous_height(coeff_labels),
        response_readout=sheet.range((_ROW_RESPONSE_READOUT, _C_AF)).value,
        observations_cell=sheet.range((_ROW_OBSERVATIONS, _C_AB)).value,
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

    # The constructor twin, displayed: the S3 spill IS
    # Constructed_Column_Names(), so content equality pins both the width
    # (k) and the level-qualified names in one assertion.
    check(
        "constructed names (S spill)",
        expected.constructed_column_names,
        observed.constructed_names,
    )
    # Twin tripwire against the engine side: the AA21 coefficient-label spill
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


def read_regression_spec_block_failures(
    workbook: xw.Book, csv_path: Path | None = None
) -> list[str]:
    """Verify the Regression sheet's spec block; return QC failure messages.

    Asserts that the workbook computes ITS OWN shipped spec correctly. The
    dataset is resolved from the sheet's live ``Source_Table`` rather than
    assumed, so this tracks ``--regression-dataset`` instead of breaking on it
    — see the module docstring for what that replaced and where the
    categorical corners went.

    ``csv_path`` overrides the resolved profile's default CSV, for a caller
    whose sample data lives elsewhere. Leave it None to use the dataset's own.

    Must run while the spec block is in its shipped state — i.e. BEFORE the
    six-configuration regression pass, which mutates the Include cells.
    """
    sheet = workbook.sheets[REGRESSION_SHEET_NAME]

    resolved = _resolve_shipped_profile(sheet)
    if resolved is None:
        return [
            f"{_QC_PREFIX} [shipped spec] check='Source_Table resolves to a "
            "known dataset': the Regression sheet's Source_Table names no "
            "registered profile, so there is nothing to compare against. "
            f"Known tables: "
            + ", ".join(
                config.table_name for config, _ in _DATASET_LOADERS.values()
            )
        ]

    profile_key, profile = resolved
    config, loader = _DATASET_LOADERS[profile_key]

    rows = loader(csv_path or config.default_csv_path)
    expected = calculate_model_construction_expectations(
        build_profile_spec(profile), rows
    )

    observed = read_observed_spec_values(
        sheet, expected.total_rows, max(expected.k, 1), profile.variables
    )
    return compare_spec_observed_to_expected(
        observed, expected, f"shipped spec ({profile_key})"
    )
