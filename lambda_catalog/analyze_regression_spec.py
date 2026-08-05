"""Spec-driven expected values for the current Regression worksheet QC harness."""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace
from pathlib import Path
import math
import numpy as np

from .analyze_life_expectancy import (
    DEFAULT_INPUT_CSV as LIFE_EXPECTANCY_CSV_PATH,
    load_life_expectancy_source_rows,
)
from .analyze_mileage import DEFAULT_INPUT_CSV
from .analyze_production_lots import (
    DEFAULT_INPUT_CSV as PRODUCTION_LOTS_CSV_PATH,
    load_production_lots_source_rows,
)
from .analyze_model_construction import (
    SpecVariable,
    _compute_mask,
    _format_value,
    _is_blank,
    _is_number,
    _retained_levels,
    build_default_spec,
    calculate_model_construction_expectations,
    interaction_header_operator,
    load_source_rows,
    resolve_interaction_operand,
)
from .analyze_regression_sheet import calculate_regression_results_from_matrix
from .regression_shared import RegressionSheetResults
from .test_model_sheets import assert_sheet_names_unique, validate_sheet_name
from .write_sheet_model_construction import (
    SPEC_DATASET_PROFILES,
    _ROLE_FILTER,
    _ROLE_FIXED_EFFECTS,
    _ROLE_IDENTIFIER,
    _ROLE_OMIT,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
)

# The Mileage/Auto MPG continuous-measurement columns available as full
# continuous predictors in the "v1_full_continuous" QC case (all 5, vs. the
# curated subset the shipped T0 default and "continuous_subset" turn on).
# Kept local to this module rather than added to regression_shared's
# FEATURE_COLUMNS, which backs the unrelated MLR_*_Test QC sheets (those
# still target Life Expectancy data independently of the Regression sheet's
# spec-driven QC oracle here).
_MILEAGE_FEATURE_COLUMNS = (
    "Cylinders",
    "Displacement",
    "Horsepower",
    "Weight",
    "Acceleration",
)


@dataclass(frozen=True)
class ExtraSpecColumn:
    """A temporary source-table column used by a QC spec fixture."""

    name: str
    excel_formula: str
    value_fn: Callable[[dict[str, object]], object]


@dataclass(frozen=True)
class RegressionSpecCase:
    """One spec-driven Regression sheet QC case."""

    name: str
    spec: tuple[SpecVariable, ...]
    allow_intercept: bool
    # The row this case implements in docs/MODEL_TESTING_ASSETS.md § 1
    # ("M05", "L07", "P06"; a trailing lowercase letter marks the second
    # half of an ±intercept twin, e.g. "M03b"). Carried on the case so the
    # workbook sheet, the plan document, and the failure message all name
    # the same thing.
    plan_id: str = ""
    # The worksheet this case is materialized on in the test-model artifact.
    # Governed by lambda_catalog/test_model_sheets.py — 31 chars, legal
    # charset, unique, and naming the CONCEPT under test rather than the
    # variables. Validated at registry-build time, not at Excel-write time.
    sheet_name: str = ""
    # Cases whose sheets are expensive enough to be opt-in: L07 (k ~ 201
    # dummy columns) and L08 (193 Fixed Effects groups), both over 2938
    # rows. The Python oracle always runs — it is cheap; only the sheet
    # build is gated, behind build_test_models.py --include-heavy.
    heavy: bool = False
    alpha: float = 0.05
    extra_columns: tuple[ExtraSpecColumn, ...] = ()
    # Override the source CSV / row loader for cases that don't target the
    # default Mileage dataset (e.g. the Production Lots Fixed Effects case,
    # which has no natural analogue in the Auto MPG table). None means "use
    # the caller's default", preserving every existing Mileage-based case.
    source_csv_path: Path | None = None
    row_loader: Callable[[Path], list[dict[str, object]]] | None = None
    # The live workbook's Source_Table RefersTo to apply before writing this
    # case's spec block — Source_Table is the ONE name that retargets which
    # data sheet the Regression sheet's spec block/design matrix read from
    # (see write_sheet_model_construction._set_sheet_scoped_names), so a case
    # from a different dataset than the QC workbook's default (Mileage) must
    # switch it or its spec rows land on the wrong table's columns entirely.
    # Every case sets this explicitly (not Optional) so the QC harness resets
    # it on each case regardless of run order — no state leaks between cases.
    source_table_ref: str = "=MileageData[#All]"
    # Which group the Prediction Interval box (AK3:AK14) is anchored to —
    # written into the sheet's own $AK$12 cell before reading the box back.
    # None mirrors the sheet's own default (AK12's formula, the
    # alphabetically-first observed group) rather than hardcoding it here.
    prediction_group: str | None = None
    # The Back-Transform Method input ($AH$4): "Duan" (the sheet's shipped
    # default) or "Naive". Written into that cell before the unit-space
    # block is read back. Only a Log-response case can distinguish the two;
    # under Transform=None they coincide, so every other case leaves it at
    # the default rather than asserting a difference that does not exist.
    back_transform: str = "Duan"
    # The typed Sequence Period (spec column I) for the Sequence-flagged row.
    # None leaves the cell blank, which is what most cases want.
    #
    # It is not cosmetic: Base_Period_Delta() reads the TYPED value and
    # returns #N/A when none is present — it is the override accessor, never
    # a silent 1 — and the BFN panel Durbin-Watson cell passes it as its
    # delta. So a Fixed Effects case with no typed period leaves AE12 at
    # #N/A by design, and the panel diagnostic is unverifiable there. A case
    # that wants BFN live has to declare the period its panel actually has.
    sequence_period: float | None = None


@dataclass(frozen=True)
class RegressionSpecDesign:
    """The spec-derived arrays and display facts behind one expected result."""

    row_mask: tuple[bool, ...]
    row_labels: tuple[str, ...]
    constructed_column_names: tuple[str, ...]
    # v2.2 Log wiring: "Log"/"None" per constructed column, mirroring the
    # sheet's Constructed_Column_Transforms() — the Python side of the
    # raw/log split the Prediction Inputs band's auto-log step needs.
    # Every dummy column from a Categorical Predictor reads "None"
    # regardless of its spec row's own Transform value.
    constructed_column_transforms: tuple[str, ...]
    # v3.3 unit-space: the response transform ("Log"/"None") the spec row
    # declared for the Response, and the predictor-side transform summary
    # ("None"/"Log"/"Mixed") the v2.2 _PREDICTOR_TRANSFORM_FORMULA computes
    # on the live spec rows. Both feed the v3.3 unit-space arithmetic.
    response_transform: str
    predictor_transform: str
    x_features: np.ndarray
    # ``y_train`` is the FIT-space response — logged when the Response row
    # declares Log — and is NOT yet within-demeaned; the demeaning happens
    # inside calculate_regression_results_from_matrix, which therefore holds
    # both the demeaned and un-demeaned columns and can derive the v3.3 level
    # shift itself. This is exactly Response_Column() on the sheet, so no
    # separate "y_full" array is carried: an array that has to be fit-space
    # for the level shift and original-units for the unit-space residual is
    # wrong for one of its two callers whatever it holds.
    y_train: np.ndarray
    sequence_values: np.ndarray | None
    group_labels: np.ndarray | None
    included_rows: int
    level_counts: dict[str, int]
    references_in_use: dict[str, object]
    degenerate_categoricals: tuple[str, ...]


@dataclass(frozen=True)
class RegressionSpecExpected:
    """One spec case plus the expected current Regression sheet outputs."""

    case: RegressionSpecCase
    design: RegressionSpecDesign
    results: RegressionSheetResults
    # Always a concrete group name, even when case.prediction_group is None
    # (resolved to the alphabetically-first observed group, matching the
    # sheet's own $AK$12 default formula) — what the QC harness writes into
    # that cell before reading the Prediction Interval box back.
    resolved_prediction_group: str


def _copy_spec(spec: tuple[SpecVariable, ...] | list[SpecVariable]) -> list[SpecVariable]:
    return [
        SpecVariable(
            item.name,
            item.role,
            item.include,
            item.var_type,
            item.reference,
            item.sequence,
            item.transform,
            item.interaction_term,
            item.interaction_operation,
        )
        for item in spec
    ]


def _replace_spec_vars(
    spec: list[SpecVariable],
    **updates: SpecVariable,
) -> list[SpecVariable]:
    by_name = {item.name: item for item in updates.values()}
    return [by_name.get(item.name, item) for item in spec]


def _spec_var(
    name: str,
    role: str,
    include: bool = False,
    var_type: str = "Continuous",
    reference: object = "",
    sequence: bool = False,
    transform: str = "None",
    interaction_term: str = "",
    interaction_operation: str = "",
) -> SpecVariable:
    return SpecVariable(
        name,
        role,
        include,
        var_type,
        reference,
        sequence,
        transform,
        interaction_term,
        interaction_operation,
    )


def _row_label(row: dict[str, object], identifiers: list[str], fallback_index: int) -> str:
    if identifiers:
        return "|".join(_format_value(row[name]) for name in identifiers)
    return f"Obs. {fallback_index + 1}"


def _numeric_cell(row: dict[str, object], name: str) -> float:
    """Return a validated numeric source value."""
    value = row[name]
    if not _is_number(value):
        raise ValueError(f"Expected numeric value for {name!r}, got {value!r}")
    return float(value)


def _with_extra_columns(
    rows: list[dict[str, object]],
    extra_columns: tuple[ExtraSpecColumn, ...],
) -> list[dict[str, object]]:
    if not extra_columns:
        return rows
    return [
        {
            **row,
            **{extra.name: extra.value_fn(row) for extra in extra_columns},
        }
        for row in rows
    ]


def _numeric_response_name(spec: tuple[SpecVariable, ...]) -> str:
    responses = [item.name for item in spec if item.role == _ROLE_RESPONSE]
    if len(responses) != 1:
        raise ValueError(f"Expected exactly one response variable, got {responses!r}")
    return responses[0]


def build_spec_design(
    spec: tuple[SpecVariable, ...] | list[SpecVariable],
    rows: list[dict[str, object]],
) -> RegressionSpecDesign:
    """Build the same row mask, names, and numeric design matrix as the sheet."""
    spec_tuple = tuple(spec)
    mask = tuple(_compute_mask(list(spec_tuple), rows))
    response_name = _numeric_response_name(spec_tuple)
    response_transform = next(
        item.transform for item in spec_tuple if item.role == _ROLE_RESPONSE
    )
    identifiers = [item.name for item in spec_tuple if item.role == _ROLE_IDENTIFIER]
    included_indices = [idx for idx, included in enumerate(mask) if included]
    if not included_indices:
        raise ValueError("Spec produced no included rows")

    y_values = []
    labels = []
    for idx in included_indices:
        value = rows[idx][response_name]
        if not _is_number(value):
            raise ValueError(f"Included row has nonnumeric response: row={idx + 1}")
        numeric_value = float(value)
        if response_transform == "Log":
            # Parity contract with Ln_Positive: the sheet returns #N/A for
            # a non-positive included value under Log; a QC case must
            # describe a legal, fully-computable model, so this raises
            # instead of silently emitting NaN.
            if numeric_value <= 0:
                raise ValueError(
                    "Included row has non-positive value for Log-transformed "
                    f"response: row={idx + 1}"
                )
            numeric_value = math.log(numeric_value)
        y_values.append(numeric_value)
        labels.append(_row_label(rows[idx], identifiers, idx))

    def _block(
        variable: SpecVariable,
    ) -> tuple[list[str], list[str], list[list[float]]] | None:
        """One spec row's own block: the Python mirror of the sheet's ``blk()``.

        Returns (names, transform flags, columns), or None for the degenerate
        Categorical case the constructor skips with its ``#N/A`` guard. Called
        for the declaring row AND for an interaction operand, so an operand
        encodes exactly as it would on its own — including its own reference
        level and its own Log transform.
        """
        if variable.var_type != "Categorical":
            is_log = variable.transform == "Log"
            values = []
            for idx in included_indices:
                raw = _numeric_cell(rows[idx], variable.name)
                if is_log:
                    if raw <= 0:
                        raise ValueError(
                            "Included row has non-positive value for "
                            f"Log-transformed predictor {variable.name!r}: "
                            f"row={idx + 1}"
                        )
                    raw = math.log(raw)
                values.append(raw)
            name = f"Ln({variable.name})" if is_log else variable.name
            return ([name], ["Log" if is_log else "None"], [values])

        retained = _retained_levels(variable, rows, list(mask))
        if retained is None:
            return None
        names = []
        columns = []
        for level in retained:
            names.append(f"{variable.name}: {_format_value(level)}")
            columns.append([
                0.0 if _is_blank(rows[idx][variable.name])
                else (1.0 if rows[idx][variable.name] == level else 0.0)
                for idx in included_indices
            ])
        # Log is disallowed on Categorical Predictors (flagged red on the
        # sheet); every dummy column reads "None" unconditionally, mirroring
        # Constructed_Column_Transforms()'s EXPAND branch.
        return (names, ["None"] * len(names), columns)

    def _combine(
        left: list[float], right: list[float], operation: str, column: str
    ) -> list[float]:
        """Pairwise operand combination — the closed Product/Difference/Ratio axis.

        Case-insensitive, because the sheet dispatches on ``SWITCH``, whose
        text comparison ignores case. An unrecognized operation raises here and
        evaluates to ``NA()`` on the sheet — both refuse, neither guesses.
        """
        operation = operation.casefold()
        if operation == "product":
            return [a * b for a, b in zip(left, right)]
        if operation == "difference":
            return [a - b for a, b in zip(left, right)]
        if operation == "ratio":
            if any(b == 0 for b in right):
                # Parity with the sheet, which returns NA() here rather than
                # a bare #DIV/0!. A QC case must describe a legal, fully
                # computable model, so this raises instead of emitting NaN.
                raise ValueError(
                    f"Ratio interaction {column!r} divides by zero in the "
                    "included sample"
                )
            return [a / b for a, b in zip(left, right)]
        raise ValueError(f"Unknown interaction operation: {operation!r}")

    matrix_columns: list[list[float]] = []
    constructed_names: list[str] = []
    constructed_transforms: list[str] = []
    for variable in spec_tuple:
        if variable.role != _ROLE_PREDICTOR or not variable.include:
            continue
        own = _block(variable)
        if own is None:
            continue
        own_names, own_transforms, own_columns = own
        constructed_names.extend(own_names)
        constructed_transforms.extend(own_transforms)
        matrix_columns.extend(own_columns)

        # v3.1: the row's interaction columns follow its own block, in the
        # same nested order the sheet's paired REDUCE emits (left operand
        # outer, right operand inner).
        operand_index = resolve_interaction_operand(variable, spec_tuple)
        if operand_index is None:
            continue
        other = _block(spec_tuple[operand_index])
        if other is None:
            continue
        other_names, _, other_columns = other
        operator = interaction_header_operator(variable.interaction_operation)
        for left_name, left_column in zip(own_names, own_columns):
            for right_name, right_column in zip(other_names, other_columns):
                name = f"{left_name}{operator}{right_name}"
                constructed_names.append(name)
                matrix_columns.append(
                    _combine(
                        left_column,
                        right_column,
                        variable.interaction_operation,
                        name,
                    )
                )
                # An interaction column is never itself Log-flagged: the
                # transform lives on each operand's own column and has
                # already been applied above.
                constructed_transforms.append("None")

    if not matrix_columns:
        raise ValueError("Spec produced zero constructed columns")
    x_features = np.asarray(matrix_columns, dtype=np.float64).T

    # v3.3: predictor-transform summary, mirroring the sheet's
    # _PREDICTOR_TRANSFORM_FORMULA ("None"/"Log"/"Mixed" over the
    # included Continuous predictors — Categorical dummies are excluded so
    # their Transform value can never spuriously flip a "None" to "Mixed").
    inc_log = sum(
        1 for item in spec_tuple
        if item.role == _ROLE_PREDICTOR and item.include
        and item.var_type == "Continuous" and item.transform == "Log"
    )
    inc_none = sum(
        1 for item in spec_tuple
        if item.role == _ROLE_PREDICTOR and item.include
        and item.var_type == "Continuous" and item.transform == "None"
    )
    if inc_log == 0:
        predictor_transform = "None"
    elif inc_none == 0:
        predictor_transform = "Log"
    else:
        predictor_transform = "Mixed"

    seq_variables = [item.name for item in spec_tuple if item.sequence]
    sequence_values = None
    if len(seq_variables) == 1:
        sequence_name = seq_variables[0]
        sequence_values = np.asarray(
            [_numeric_cell(rows[idx], sequence_name) for idx in included_indices],
            dtype=np.float64,
        )

    # Fixed_Effects_Column(): the exact-match Role accessor's Python mirror.
    # Zero rows is the ordinary no-FE case; the spec-error states (2+ rows)
    # are a visible sheet warning the human test plan covers, not something
    # this Python design builder needs to reproduce — callers only ever hand
    # it a legal (0-or-1) spec.
    fe_variables = [item.name for item in spec_tuple if item.role == _ROLE_FIXED_EFFECTS]
    if len(fe_variables) > 1:
        raise ValueError(f"Expected at most one Fixed Effects variable, got {fe_variables!r}")
    group_labels = None
    if fe_variables:
        fe_name = fe_variables[0]
        group_labels = np.asarray([rows[idx][fe_name] for idx in included_indices])

    expectations = calculate_model_construction_expectations(list(spec_tuple), rows)
    return RegressionSpecDesign(
        row_mask=mask,
        row_labels=tuple(labels),
        constructed_column_names=tuple(constructed_names),
        constructed_column_transforms=tuple(constructed_transforms),
        response_transform=response_transform,
        predictor_transform=predictor_transform,
        x_features=x_features,
        y_train=np.asarray(y_values, dtype=np.float64),
        sequence_values=sequence_values,
        group_labels=group_labels,
        included_rows=len(included_indices),
        level_counts=expectations.level_counts,
        references_in_use=expectations.references_in_use,
        degenerate_categoricals=expectations.degenerate_categoricals,
    )


def calculate_regression_spec_case(
    case: RegressionSpecCase,
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> RegressionSpecExpected:
    """Compute expected current Regression sheet outputs for one spec case."""
    effective_csv_path = (
        case.source_csv_path if case.source_csv_path is not None else csv_path
    )
    loader = case.row_loader if case.row_loader is not None else load_source_rows
    rows = _with_extra_columns(loader(effective_csv_path), case.extra_columns)
    design = build_spec_design(case.spec, rows)

    # Resolve to a concrete group name up front (mirrors $AK$12's own default
    # formula: the alphabetically-first observed group) so the QC harness has
    # an explicit value to write into that cell — never ambiguous about
    # "leave it at whatever the sheet defaults to."
    if case.prediction_group is not None:
        resolved_prediction_group = case.prediction_group
    elif design.group_labels is not None:
        resolved_prediction_group = str(sorted(np.unique(design.group_labels))[0])
    else:
        resolved_prediction_group = "(all)"

    response_name = next(
        item.name for item in case.spec if item.role == _ROLE_RESPONSE
    )
    fixed_effects_name = next(
        (item.name for item in case.spec if item.role == _ROLE_FIXED_EFFECTS),
        None,
    )
    results = calculate_regression_results_from_matrix(
        x_features=design.x_features,
        y_train=design.y_train,
        predictor_names=design.constructed_column_names,
        include_intercept=case.allow_intercept,
        alpha=case.alpha,
        sequence_values=design.sequence_values,
        group_labels=design.group_labels,
        selected_group=resolved_prediction_group,
        response_transform=design.response_transform,
        predictor_transform=design.predictor_transform,
        response_name=response_name,
        fixed_effects_name=fixed_effects_name,
        back_transform=case.back_transform,
        base_period_delta=case.sequence_period,
    )
    return RegressionSpecExpected(
        case=case,
        design=design,
        results=results,
        resolved_prediction_group=resolved_prediction_group,
    )


def _v1_full_continuous_spec() -> list[SpecVariable]:
    numeric_predictors = set(_MILEAGE_FEATURE_COLUMNS)
    spec = []
    for variable in build_default_spec():
        if variable.name == "Car Name":
            spec.append(_spec_var(variable.name, _ROLE_IDENTIFIER))
        elif variable.name == "Model Year":
            spec.append(_spec_var(variable.name, _ROLE_IDENTIFIER))
        elif variable.name == "Origin":
            spec.append(_spec_var(variable.name, _ROLE_OMIT))
        elif variable.name == "MPG":
            spec.append(_spec_var(variable.name, _ROLE_RESPONSE))
        elif variable.name == "Full_Data":
            spec.append(_spec_var(variable.name, _ROLE_FILTER))
        elif variable.name in numeric_predictors:
            spec.append(_spec_var(variable.name, _ROLE_PREDICTOR, True, "Continuous"))
        else:
            spec.append(_spec_var(variable.name, _ROLE_OMIT))
    return spec


def _continuous_subset_spec() -> list[SpecVariable]:
    selected = {"Displacement", "Horsepower", "Weight"}
    spec = []
    for variable in _v1_full_continuous_spec():
        if variable.role == _ROLE_PREDICTOR and variable.name not in selected:
            spec.append(_spec_var(variable.name, _ROLE_PREDICTOR, False, variable.var_type))
        else:
            spec.append(variable)
    return spec


def _with_origin(spec: list[SpecVariable], reference: object = "") -> list[SpecVariable]:
    return _replace_spec_vars(
        spec,
        origin=_spec_var("Origin", _ROLE_PREDICTOR, True, "Categorical", reference),
    )


def _model_year_origin_categorical_spec() -> list[SpecVariable]:
    return _replace_spec_vars(
        _with_origin(_continuous_subset_spec()),
        model_year=_spec_var("Model Year", _ROLE_PREDICTOR, True, "Categorical"),
    )


def _interaction_spec(
    term: str, operation: str = "Product", *, categorical_operand: bool = False
) -> list[SpecVariable]:
    """Declare an interaction on the Weight row of the continuous subset.

    Weight is an included Continuous Predictor that sits BEFORE Origin in the
    Auto MPG column order, so the emitted interaction columns land between
    Weight's own column and Origin's dummies — which is what pins the
    constructor's emission order (each row's interaction follows its own
    block, not appended at the end of the matrix).
    """
    base = _with_origin(_continuous_subset_spec()) if categorical_operand else _continuous_subset_spec()
    return _replace_spec_vars(
        base,
        weight=_spec_var(
            "Weight",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term=term,
            interaction_operation=operation,
        ),
    )


def _mileage_log_log_na_masking_spec() -> list[SpecVariable]:
    """M5 — Ln(MPG) ~ Ln(Weight) + Ln(Horsepower) on Auto MPG.

    Covers ``(Log, Log)`` on a dataset OTHER than Production Lots, which is
    the only place the pair is currently exercised. The distinction is not
    cosmetic: Production Lots has no missing values, so its (Log, Log)
    cases never combine the transform with NA propagation. Auto MPG does —
    8 rows are missing MPG and 6 are missing Horsepower — so this case is
    the one that proves the mask is applied BEFORE the logs are taken
    rather than after (taking Ln of a blank would poison the column, not
    drop the row).

    Model Year and Car Name both stay Identifiers, so the row labels are the
    shipped ones. Neither is Sequence-flagged: Auto MPG is cross-sectional
    (one row per car model, no unit repeated across periods), so there is no
    ordering axis to declare — see ``_DEFAULT_SEQUENCE_VARIABLES``.
    """
    return [
        _spec_var("MPG", _ROLE_RESPONSE, transform="Log"),
        _spec_var("Cylinders", _ROLE_OMIT),
        _spec_var("Displacement", _ROLE_OMIT),
        _spec_var("Horsepower", _ROLE_PREDICTOR, True, "Continuous", transform="Log"),
        _spec_var("Weight", _ROLE_PREDICTOR, True, "Continuous", transform="Log"),
        _spec_var("Acceleration", _ROLE_OMIT),
        _spec_var("Model Year", _ROLE_IDENTIFIER),
        _spec_var("Origin", _ROLE_OMIT),
        _spec_var("Car Name", _ROLE_IDENTIFIER),
        _spec_var("Make", _ROLE_OMIT),
        _spec_var("Model?", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_OMIT),
    ]


def _categorical_only_design_spec() -> list[SpecVariable]:
    """M14b — MPG ~ C(Model Year) + C(Origin), with NO continuous predictors.

    The plan's M14 ("two categoricals, no continuous") turned out not to be
    what the shipped ``model_year_origin_categorical`` case builds — that
    one keeps Displacement/Horsepower/Weight on, so the categorical-only
    design has never actually been fitted. This is that design.

    Two things make it worth its own case rather than a trim of M14. The
    mask is the interesting part: with no included Continuous Predictor,
    ``Sample_Include`` reduces to "MPG is numeric", so the sample GROWS to
    398 rows (the 6 Horsepower-missing rows rejoin) — the clearest possible
    demonstration that the mask is per-model, not per-dataset. And it is
    M9's base: M9 is this design plus the interaction block, so the two
    cross-check each other's main-effect columns.
    """
    spec = []
    for variable in build_default_spec():
        if variable.name in ("Horsepower", "Weight"):
            spec.append(_spec_var(variable.name, _ROLE_PREDICTOR, False, "Continuous"))
        else:
            spec.append(variable)
    return spec


def _interaction_categorical_cross_spec() -> list[SpecVariable]:
    """M9 — MPG ~ C(Model Year) + C(Origin) + C(Model Year) x C(Origin).

    The Cat x Cat full-product width regime, the one the v3.1 interaction
    wiring emits ``(L1-1) * (L2-1)`` columns for and which no existing case
    covers: ``interaction_categorical_broadcast`` is Continuous x
    Categorical (``1 * (L-1)``), and nothing crosses two dummy blocks.
    Here that is 12 * 2 = 24 interaction columns on top of 12 + 2 main
    effects, so the constructor's nested REDUCE has to get both the count
    and the left-outer/right-inner ordering right.

    **Why Model Year x Origin and not the plan's Cylinders x Origin.** The
    Cylinders x Origin cross-tabulation is sparse — Cylinders=3 appears
    only in Asia, 5 only in Europe, 8 only in the US — so two of its eight
    product columns are identically zero and the Gram matrix is singular.
    That is a rank-deficiency test, not a width test, and Excel's MINVERSE
    would return #NUM! where NumPy's lstsq quietly returns a minimum-norm
    solution, so the two sides could not be compared at all. Model Year x
    Origin populates all 39 cells (minimum cell count 2, condition number
    ~116), which is a genuine saturated two-factor design. Model Year is
    numeric-valued, so the numeric-categorical corner the plan wanted from
    Cylinders is still covered.
    """
    return _replace_spec_vars(
        _categorical_only_design_spec(),
        model_year=_spec_var(
            "Model Year",
            _ROLE_PREDICTOR,
            True,
            "Categorical",
            interaction_term="Origin",
            interaction_operation="Product",
        ),
    )


def _interaction_difference_spec() -> list[SpecVariable]:
    """M10 — the first ``Difference`` interaction case.

    Covers the antisymmetric arm of the closed Product/Difference/Ratio
    vocabulary and the U+2212 MINUS SIGN in the constructed column header
    (``"Displacement - Horsepower"`` with the typographic minus, not a
    hyphen) — a header the QC comparison matches on, so a silent change to
    ``interaction_header_operator`` fails here.

    **Why the operand is neither Horsepower nor a main effect.** The plan
    writes this as ``MPG ~ Displacement + Horsepower + Displacement -
    Horsepower``, which is exactly singular: the difference column is a
    linear combination of the two main effects, so the design has rank k-1
    and neither Excel nor the oracle can fit it. Leaving the operand out of
    the main effects fixes the rank, which makes it an operand with
    ``Include = FALSE`` — the flagged-amber marginality state G11 documents
    as allowed, so this case exercises that path too.

    The operand is then ``Acceleration`` rather than ``Horsepower``,
    because an EXCLUDED operand imposes no mask condition. ``Sample_Include``
    tests completeness on the Response and the *included* Continuous
    Predictors only, so a row missing an excluded operand stays in the
    sample and its interaction column evaluates to ``#N/A`` — Horsepower's
    6 missing rows would poison the design. Acceleration is complete on
    every row, so the case tests the Difference operator rather than that
    interaction. (The excluded-operand-plus-missingness combination is real
    and worth flagging; it is a property of the mask, not of this case.)
    """
    return [
        _spec_var("MPG", _ROLE_RESPONSE),
        _spec_var("Cylinders", _ROLE_OMIT),
        _spec_var(
            "Displacement",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term="Acceleration",
            interaction_operation="Difference",
        ),
        _spec_var("Horsepower", _ROLE_OMIT),
        _spec_var("Weight", _ROLE_PREDICTOR, True, "Continuous"),
        _spec_var("Acceleration", _ROLE_PREDICTOR, False, "Continuous"),
        _spec_var("Model Year", _ROLE_IDENTIFIER),
        _spec_var("Origin", _ROLE_OMIT),
        _spec_var("Car Name", _ROLE_IDENTIFIER),
        _spec_var("Make", _ROLE_OMIT),
        _spec_var("Model?", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_OMIT),
    ]


def _interaction_ratio_reciprocal_spec() -> list[SpecVariable]:
    """M11 — MPG ~ Weight + Weight / Horsepower + Horsepower + Horsepower / Weight.

    Two things at once, and they are related. It is the first ``Ratio``
    case, exercising the division arm and its zero-denominator ``NA()``
    guard (no included row has a zero Weight or Horsepower, so the guard
    stays quiet here and the fit is clean). And it is the LEGAL reciprocal
    declaration: both A/B and B/A are declared, which for a symmetric
    operation would produce two identical columns and a singular Gram
    matrix — the state G10 flags red for ``Product``. Ratio is asymmetric,
    so A/B and B/A are genuinely different columns and the model fits.
    G10 and this case are the two halves of one rule.
    """
    return [
        _spec_var("MPG", _ROLE_RESPONSE),
        _spec_var("Cylinders", _ROLE_OMIT),
        _spec_var("Displacement", _ROLE_OMIT),
        _spec_var(
            "Horsepower",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term="Weight",
            interaction_operation="Ratio",
        ),
        _spec_var(
            "Weight",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term="Horsepower",
            interaction_operation="Ratio",
        ),
        _spec_var("Acceleration", _ROLE_OMIT),
        _spec_var("Model Year", _ROLE_IDENTIFIER),
        _spec_var("Origin", _ROLE_OMIT),
        _spec_var("Car Name", _ROLE_IDENTIFIER),
        _spec_var("Make", _ROLE_OMIT),
        _spec_var("Model?", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_OMIT),
    ]


_IS_USA = ExtraSpecColumn(
    name="Is_USA",
    excel_formula='=--([@Origin]="US")',
    value_fn=lambda row: 1 if row["Origin"] == "US" else 0,
)


# ── Life Expectancy specs (docs/MODEL_TESTING_ASSETS.md § 1.2) ──────────────
#
# Until now not one regression QC case targeted this dataset, so three
# things had no oracle anywhere: the (None, Mixed) and (Log, None) dispatch
# pairs, transform behaviour at 2938-row scale, and — the reason the dataset
# is in the plan at all — masking against genuinely heavy missingness
# (Population 652 blanks, GDP 448, Alcohol 194, Schooling 163). Auto MPG's
# 8-and-6 missing cells do not stress that; a model whose mask is the
# intersection of four sparse columns does.
#
# Every spec below is built by _life_spec so the 23 rows stay in the
# dataset's own column order (spec rows are positional, one per
# Source_Table column) without 23 literal lines per case.


def _life_spec(**overrides: SpecVariable) -> list[SpecVariable]:
    """Build a Life Expectancy spec: named overrides over an Omit baseline.

    Everything not named is ``Omit``, except the two structural roles the
    dataset ships with and every model wants: ``Country`` is the Identifier
    (row labels) and ``Year`` is the Sequence axis (Role=Omit, so it never
    enters the design matrix itself — it only drives the base-period /
    serial-correlation layer). Override either by name to change it.

    Keyword names are ignored; only the ``SpecVariable.name`` matters, which
    is what lets a caller write ``_life_spec(response=..., gdp=...)`` and
    have the rows land in dataset order regardless of argument order.
    """
    by_name = {variable.name: variable for variable in overrides.values()}
    spec: list[SpecVariable] = []
    for name in SPEC_DATASET_PROFILES["life_expectancy"].variables:
        if name in by_name:
            spec.append(by_name[name])
        elif name == "Country":
            spec.append(_spec_var(name, _ROLE_IDENTIFIER))
        elif name == "Year":
            spec.append(_spec_var(name, _ROLE_OMIT, sequence=True))
        else:
            spec.append(_spec_var(name, _ROLE_OMIT))
    return spec


def _life_partial_linear_log_spec(reference: object = "") -> list[SpecVariable]:
    """L1 / L9 — Life expectancy ~ Ln(Population) + Ln(GDP) + Alcohol + C(Status).

    A **partial linear-log** model, and the suite's only ``(None, Mixed)``
    dispatch pair: two logged Continuous predictors and one unlogged,
    against an untransformed response. That combination is
    what proves ``_PREDICTOR_TRANSFORM_FORMULA`` reports "Mixed" rather
    than latching to whichever transform it saw first, and that the
    unit-space block reduces cleanly (no response transform ⇒ smearing 1)
    even when the predictor side is mixed.

    It is also the heaviest masking case in the suite: the sample is the
    intersection of four columns with 652 / 448 / 194 / 0 blanks, which
    drops it well below half the 2938 rows.

    **The name is "linear-log", not "log-linear".** The two are opposite
    specifications and this case is unambiguously the first: the logs sit
    on the PREDICTORS and the response is untransformed, so a coefficient
    reads as a semi-elasticity — years of life expectancy per 100% change
    in GDP. "Log-linear" (log-lin) is the mirror image, ``ln(y) ~ x``,
    which in this suite is L2, and L2 already carries that model's other
    standard name — the exponential model. Naming this one "log-linear"
    would have given the same label to both halves of the dispatch table
    the two cases exist to tell apart. "Partial" is the ``Mixed`` half:
    Alcohol stays raw while Population and GDP are logged.

    ``Status`` is the binary categorical. With ``reference`` blank the
    first sorted level ("Developed") is dropped and "Developing" is
    retained; L9 passes "Developing" to flip that — the explicit-reference
    path on a two-level column, where getting it backwards is invisible in
    the column COUNT (one dummy either way) and only shows in the
    coefficient's sign.
    """
    return _life_spec(
        status=_spec_var("Status", _ROLE_PREDICTOR, True, "Categorical", reference),
        response=_spec_var("Life expectancy", _ROLE_RESPONSE),
        alcohol=_spec_var("Alcohol", _ROLE_PREDICTOR, True, "Continuous"),
        gdp=_spec_var("GDP", _ROLE_PREDICTOR, True, "Continuous", transform="Log"),
        population=_spec_var(
            "Population", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
    )


def _life_log_response_spec() -> list[SpecVariable]:
    """L2 / L3 — Ln(Life expectancy) ~ Adult Mortality + Schooling + C(Status).

    The **exponential model** (log-level: ``ln(y) = a + bx`` ⟺
    ``y = exp(a + bx)``, so a coefficient is a proportional change in y per
    unit of x) and the suite's only ``(Log, None)``
    dispatch pair: a logged response against entirely unlogged predictors.
    That is the pair where the v3.3 unit-space machinery does the most work
    — the smearing factor is not 1, the R2/Adj R2/RMSE in original units
    genuinely differ from the fit-space ones, and the AL prediction column
    and AZ/BA residual columns all have something to back-transform.

    L2 and L3 share this spec exactly and differ ONLY in the Back-Transform
    Method cell ($AH$4): Duan vs Naive. Keeping one spec builder for both
    is the point — any difference between the two cases' expected values is
    attributable to the toggle and nothing else.

    Schooling is deliberately used raw here, not logged. Logging it is L6's
    job, and L6 is a guard state rather than a model because the column
    contains 28 true zeros (see analyze_regression_guard_states.py).
    """
    return _life_spec(
        status=_spec_var("Status", _ROLE_PREDICTOR, True, "Categorical"),
        response=_spec_var("Life expectancy", _ROLE_RESPONSE, transform="Log"),
        adult_mortality=_spec_var(
            "Adult Mortality", _ROLE_PREDICTOR, True, "Continuous"
        ),
        schooling=_spec_var("Schooling", _ROLE_PREDICTOR, True, "Continuous"),
    )


def _life_elasticity_log_log_spec() -> list[SpecVariable]:
    """L4 — Ln(Life expectancy) ~ Ln(GDP) + Ln(Population).

    The elasticity form: ``(Log, Log)`` at 2938-row scale against sparse
    predictors. The existing (Log, Log) cases are all Production Lots, a
    complete 51-row panel, so this is the first place the pair meets both
    large-sample masking and a response whose log is taken on ~1600
    surviving rows rather than 51. Coefficients here read as elasticities,
    which is why the model is worth having beyond the dispatch coverage.
    """
    return _life_spec(
        response=_spec_var("Life expectancy", _ROLE_RESPONSE, transform="Log"),
        gdp=_spec_var("GDP", _ROLE_PREDICTOR, True, "Continuous", transform="Log"),
        population=_spec_var(
            "Population", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
    )


def _life_full_profile_spec() -> list[SpecVariable]:
    """L5 — the shipped ``life_expectancy`` spec profile, finally with an oracle.

    ``SPEC_DATASET_PROFILES["life_expectancy"]`` is what
    ``build_production.py --regression-dataset life_expectancy`` pre-fills
    into the spec block, and nothing has ever verified that the model it
    ships actually fits: all 18 continuous predictors plus C(Status),
    Country as Identifier, Year as the Sequence axis. Derived from the
    profile itself rather than restated, so the case tracks the shipped
    default automatically if that default ever changes.

    It is the suite's k-stress case at k = 19 — every predictor-summary
    statistic (GVIF in particular, which inverts a 19x19 correlation
    matrix) is computed on a wide design here.
    """
    profile = SPEC_DATASET_PROFILES["life_expectancy"]
    return [
        _spec_var(
            name,
            *profile.default_spec[name],
            sequence=name in profile.sequence_variables,
        )
        for name in profile.variables
    ]


# The eight continuous predictors L7 crosses with C(Country) + C(Year).
# Chosen for COMPLETENESS rather than modelling interest: each one keeps all
# 183 countries that survive the response's own 10 blanks, so the country
# dummy block stays at its maximum 182 columns. Swapping any of these for a
# sparser column (BMI and the two thinness columns lose 2 countries; GDP
# loses 25, Population 40) silently narrows the design and drops the case
# back under the width-guard threshold it exists to cross.
_LIFE_WIDTH_GUARD_PREDICTORS = (
    "Adult Mortality",
    "infant deaths",
    "percentage expenditure",
    "Measles",
    "HIV/AIDS",
    "under-five deaths",
    "Polio",
    "Diphtheria",
)


def _life_country_width_guard_spec() -> list[SpecVariable]:
    """L7 — Life expectancy ~ C(Country) + C(Year) + 8 continuous. k = 205.

    The soft width-guard case, and the only one: the M2 status line warns
    once the design reaches 200 columns, and nothing else in the suite
    comes within an order of magnitude of it.

    **A GUARD STATE, not a fittable model** — used by
    ``analyze_regression_guard_states.build_guard_state_cases``, not by
    ``build_regression_spec_cases``. The first live Excel run of this case
    returned ``nan`` for every engine output: at k = 205 the workbook cannot
    invert the Gram matrix, and 22,886 of that run's 22,898 mismatches were
    this one case comparing real numbers against nothing.

    That is not a defect to route around — it is the exact condition the
    width guard exists to warn about, arriving one predictor block earlier
    than the guard's own threshold suggests. So the case keeps its spec and
    its reason for existing, and asserts what the sheet actually does: the
    M2 status reads WARNING, the design-column total is 205, and the engine
    degrades visibly instead of returning a plausible wrong number. A
    numeric oracle for a model the sheet cannot compute would be comparing
    against nothing, which is the same reasoning that put M16, P07 and L06
    in the guard registry.

    **Why C(Year) is here and the plan does not mention it.** The plan
    assumed 193 countries → 192 dummies, so eight continuous predictors
    would reach k = 200 exactly. The real data does not allow that: the
    response is blank on 10 rows covering whole countries, so at most 183
    countries ever survive the mask and the dummy block caps at 182. Even
    all 18 continuous columns cannot make up the difference, because the
    sparse ones drop further countries roughly as fast as they add columns.
    Declaring Year (16 levels) as a second Categorical Predictor adds 15
    columns that cost no rows at all, putting k at 182 + 15 + 8 = 205 —
    over the threshold with margin. It also makes the design a genuine
    two-way categorical one, which is a fair description of what a
    200-column spec block looks like in practice.

    Year stays Sequence-flagged while being a Categorical Predictor, which
    is legal — Sequence is structural, not a Role, so a column can be both.
    It is also true here rather than decorative: Life Expectancy is a real
    country x year panel, which is why this dataset's shipped profile flags
    Year and Auto MPG's flags nothing (see ``_DEFAULT_SEQUENCE_VARIABLES``).
    ``Country`` moves from Identifier to Categorical Predictor, so the row
    labels fall back to positional ("Obs. 1", ...).
    """
    return _life_spec(
        country=_spec_var("Country", _ROLE_PREDICTOR, True, "Categorical"),
        year=_spec_var("Year", _ROLE_PREDICTOR, True, "Categorical", sequence=True),
        response=_spec_var("Life expectancy", _ROLE_RESPONSE),
        **{
            f"p{index}": _spec_var(name, _ROLE_PREDICTOR, True, "Continuous")
            for index, name in enumerate(_LIFE_WIDTH_GUARD_PREDICTORS)
        },
    )


def _life_country_fixed_effects_spec() -> list[SpecVariable]:
    """L8 — Life expectancy ~ Schooling + Adult Mortality | Country. HEAVY.

    High-cardinality Fixed Effects: 193 groups against Production Lots'
    three. The within-estimator and its absorbed-degrees-of-freedom
    threading are currently proven only on a 3-group, 51-row panel, where a
    df error of a few units is easy to miss; at 192 absorbed df every
    df-dependent statistic (SE, t, p, CI, AIC/BIC/AICc, F) moves visibly if
    the absorption is wrong.

    It is also the panel-spacing verdicts at scale — Year is the Sequence
    axis across 193 groups of ~16 observations each.

    Marked ``heavy`` for the same reason as L7: the oracle is cheap, the
    sheet is not.
    """
    return _life_spec(
        country=_spec_var("Country", _ROLE_FIXED_EFFECTS),
        response=_spec_var("Life expectancy", _ROLE_RESPONSE),
        adult_mortality=_spec_var(
            "Adult Mortality", _ROLE_PREDICTOR, True, "Continuous"
        ),
        schooling=_spec_var("Schooling", _ROLE_PREDICTOR, True, "Continuous"),
    )


def _production_lots_fixed_effects_spec() -> list[SpecVariable]:
    """Facility=Fixed Effects, Fiscal_Year=Sequence: log Cum Units -> log Unit Cost.

    The only shipped case that declares Role=Fixed Effects — a small
    unbalanced panel (3 facilities, 51 lots) that exercises the within-demeaned
    fit-time pair (calculate_regression_results_from_matrix's group_labels
    branch) against a real spec-driven build, unlike Auto MPG (no natural
    panel-unit variable). Column order matches the Production Lots CSV's
    header order plus the appended Full_Data column — spec rows are
    positional, one per Source_Table column.
    """
    return [
        _spec_var("Lot_ID", _ROLE_IDENTIFIER),
        _spec_var("Facility", _ROLE_FIXED_EFFECTS),
        _spec_var("Fiscal_Year", _ROLE_OMIT, sequence=True),
        _spec_var("Lot_Quantity", _ROLE_OMIT),
        _spec_var("Cumulative_Units", _ROLE_OMIT),
        _spec_var("Experience_Stock", _ROLE_OMIT),
        _spec_var("Unit_Cost_BY", _ROLE_OMIT),
        _spec_var("log Cum Units", _ROLE_PREDICTOR, True, "Continuous"),
        _spec_var("log experience", _ROLE_OMIT),
        _spec_var("log Unit Cost", _ROLE_RESPONSE),
        _spec_var("Full_Data", _ROLE_FILTER),
    ]


def _production_lots_log_transform_spec() -> list[SpecVariable]:
    """Facility=Fixed Effects, Fiscal_Year=Sequence: Cumulative_Units -Log-> Unit_Cost_BY.

    Sibling of _production_lots_fixed_effects_spec(), pointed at the RAW
    columns with transform="Log" instead of the precomputed "log Cum
    Units" / "log Unit Cost" columns it uses — the acceptance test for the
    v2.2 Transform=Log wiring on both a Predictor and the Response
    simultaneously, composed with Fixed Effects. This is the textbook
    Crawford/Wright learning-curve model (ln(unit cost) = a + b*ln(cum
    units)). tests/test_transform_threading.py cross-checks this case
    against the sibling above: the shipped "log Cum Units"/"log Unit
    Cost" columns are exact logs of the raw columns, so the two designs
    and every downstream statistic are expected to agree to floating-point
    precision — independent proof the Log wiring reproduces what the
    precomputed-column workaround already delivered.
    """
    return [
        _spec_var("Lot_ID", _ROLE_IDENTIFIER),
        _spec_var("Facility", _ROLE_FIXED_EFFECTS),
        _spec_var("Fiscal_Year", _ROLE_OMIT, sequence=True),
        _spec_var("Lot_Quantity", _ROLE_OMIT),
        _spec_var(
            "Cumulative_Units", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
        _spec_var("Experience_Stock", _ROLE_OMIT),
        _spec_var("Unit_Cost_BY", _ROLE_RESPONSE, transform="Log"),
        _spec_var("log Cum Units", _ROLE_OMIT),
        _spec_var("log experience", _ROLE_OMIT),
        _spec_var("log Unit Cost", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_FILTER),
    ]


def _production_lots_derived_log_no_fe_spec() -> list[SpecVariable]:
    """P3 — log Unit Cost ~ log Cum Units, NO Fixed Effects. Pre-derived columns.

    The pre-derived half of the no-FE pair, and P03b's twin. It fits the
    identical model by the other route: where P03b declares
    ``Transform = Log`` on the raw ``Cumulative_Units`` / ``Unit_Cost_BY``
    columns, this one points the spec straight at the shipped ``log Cum
    Units`` / ``log Unit Cost`` columns and declares no transform at all.

    **Why the suite wants both routes twice.** P01/P02 already pair
    pre-derived against transform-axis, but only *with* Fixed Effects, so
    the cross-check has never run on a design the FE machinery does not
    touch. The two mechanisms reach the design matrix by different code
    paths — one reads a column, the other computes one — and composing
    either with FE demeaning is a third thing again. Pairing them without
    FE isolates the transform axis from the FE axis, so a regression in
    one cannot hide behind the other.

    It is also the cheapest strong oracle available: the shipped log
    columns are exact logs of the raw ones, so this case and P03b must
    agree BIT-for-bit on the design matrix and response vector and to
    floating point on every downstream statistic, with neither side
    reading the workbook. ``tests/test_transform_threading.py`` asserts
    that, mirroring the P01/P02 assertion it already makes.

    The two legitimately differ on ``constructed_column_names`` — "log Cum
    Units" against "Ln(Cumulative_Units)" — and on the response display
    name. That is the mechanism showing through the label, not a
    disagreement about the fit, so the cross-check compares numerics only.
    """
    return [
        _spec_var("Lot_ID", _ROLE_IDENTIFIER),
        _spec_var("Facility", _ROLE_OMIT),
        _spec_var("Fiscal_Year", _ROLE_OMIT, sequence=True),
        _spec_var("Lot_Quantity", _ROLE_OMIT),
        _spec_var("Cumulative_Units", _ROLE_OMIT),
        _spec_var("Experience_Stock", _ROLE_OMIT),
        _spec_var("Unit_Cost_BY", _ROLE_OMIT),
        _spec_var("log Cum Units", _ROLE_PREDICTOR, True, "Continuous"),
        _spec_var("log experience", _ROLE_OMIT),
        _spec_var("log Unit Cost", _ROLE_RESPONSE),
        _spec_var("Full_Data", _ROLE_FILTER),
    ]


def _production_lots_log_no_fe_spec() -> list[SpecVariable]:
    """P3b — Log+Log with NO Fixed Effects: the (Log, Log) SWITCH branch.

    Sibling of _production_lots_log_transform_spec() with ``Facility``
    omitted instead of declared as Fixed Effects. Exercises the v3.3
    unit-space dispatcher's (Log, Log) branch where the level shift
    Y_Full − Y is exactly zero, so the unit-space arithmetic reduces
    to back-transforming the in-sample fit and the new (smeared) R²
    is computed cleanly. Reduction invariant: with no FE, the
    smearing factor uses raw residuals, not within residuals.

    It is also the transform-axis half of the no-FE pair — the model
    ``_production_lots_derived_log_no_fe_spec`` fits from the shipped
    pre-derived log columns instead. The two sit adjacent in the registry
    and on adjacent worksheets, exactly as P01/P02 do with Fixed Effects.
    """
    return [
        _spec_var("Lot_ID", _ROLE_IDENTIFIER),
        _spec_var("Facility", _ROLE_OMIT),
        _spec_var("Fiscal_Year", _ROLE_OMIT, sequence=True),
        _spec_var("Lot_Quantity", _ROLE_OMIT),
        _spec_var(
            "Cumulative_Units", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
        _spec_var("Experience_Stock", _ROLE_OMIT),
        _spec_var("Unit_Cost_BY", _ROLE_RESPONSE, transform="Log"),
        _spec_var("log Cum Units", _ROLE_OMIT),
        _spec_var("log experience", _ROLE_OMIT),
        _spec_var("log Unit Cost", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_FILTER),
    ]


def _production_lots_log_mixed_predictors_spec() -> list[SpecVariable]:
    """v3.3 — Mixed Log/None predictors with a Log response: the (Log, Mixed) branch.

    Sibling of _production_lots_log_no_fe_spec() with an additional
    untransformed Continuous predictor (``Experience_Stock``) — exercises
    the ``(Log, Mixed)`` SWITCH branch the v2.2 transform-threading
    rewrite unlocked. The mixed case is the cell the user explicitly
    asked about: a spec with one logged and one unlogged predictor
    must NOT return #N/A in the unit-space block.
    """
    return [
        _spec_var("Lot_ID", _ROLE_IDENTIFIER),
        _spec_var("Facility", _ROLE_OMIT),
        _spec_var("Fiscal_Year", _ROLE_OMIT, sequence=True),
        _spec_var("Lot_Quantity", _ROLE_OMIT),
        _spec_var(
            "Cumulative_Units", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
        _spec_var("Experience_Stock", _ROLE_PREDICTOR, True, "Continuous"),
        _spec_var("Unit_Cost_BY", _ROLE_RESPONSE, transform="Log"),
        _spec_var("log Cum Units", _ROLE_OMIT),
        _spec_var("log experience", _ROLE_OMIT),
        _spec_var("log Unit Cost", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_FILTER),
    ]


def _production_lots_log_predictor_only_spec() -> list[SpecVariable]:
    """v3.3 — Log predictor, None response: the (None, Log) branch.

    Single Log-transformed predictor against an untransformed response —
    exercises the (None, Log) SWITCH branch. The reduction invariant
    demands the unit-space block read identically to the ordinary
    statistics: with no Response transform, smearing=1, the back-
    transformation is a pass-through, and Unit_Space_R_Squared ==
    R_Squared, Unit_Space_RMSE == SE_Regression.
    """
    return [
        _spec_var("Lot_ID", _ROLE_IDENTIFIER),
        _spec_var("Facility", _ROLE_OMIT),
        _spec_var("Fiscal_Year", _ROLE_OMIT, sequence=True),
        _spec_var("Lot_Quantity", _ROLE_OMIT),
        _spec_var(
            "Cumulative_Units", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
        _spec_var("Experience_Stock", _ROLE_OMIT),
        _spec_var("Unit_Cost_BY", _ROLE_RESPONSE),
        _spec_var("log Cum Units", _ROLE_OMIT),
        _spec_var("log experience", _ROLE_OMIT),
        _spec_var("log Unit Cost", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_FILTER),
    ]


# Case name -> (plan ID, worksheet name) for every fittable case, keyed by
# the case's own name so a rename shows up here as a KeyError rather than as
# a silently unnamed sheet. The plan ID is the row in
# docs/MODEL_TESTING_ASSETS.md § 1 this case implements; the sheet name is
# what the test-model artifact's tab reads. Both are validated against
# test_model_sheets.py's contract by build_regression_spec_cases().
#
# Sheet names state the CONCEPT under test, never the variables — 31
# characters cannot hold a model formula, and the corner a case exists for
# is the useful thing to read off a tab.
_CASE_SHEET_IDENTITY: dict[str, tuple[str, str]] = {
    # § 1.1 Auto MPG — baseline, categoricals, interactions.
    "default_t0_intercept": ("M01", "M01 Baseline Categoricals"),
    "default_t0_no_intercept": ("M02", "M02 Intercept Off Categorical"),
    "v1_full_continuous_intercept": ("M03", "M03 All Continuous"),
    "v1_full_continuous_no_intercept": ("M03b", "M03b All Continuous NoInt"),
    "continuous_subset_intercept": ("M04", "M04 Excluded Candidates"),
    "continuous_subset_no_intercept": ("M04b", "M04b Excluded Cands NoInt"),
    "mileage_log_log_na_masking": ("M05", "M05 Log-Log NA Masking"),
    "interaction_quadratic_self_product": ("M06", "M06 Quadratic Self Product"),
    "interaction_continuous_product": ("M07", "M07 Continuous Product"),
    "interaction_categorical_broadcast": ("M08", "M08 Cont x Cat Broadcast"),
    "interaction_categorical_cross": ("M09", "M09 Cat x Cat Full Product"),
    "interaction_difference": ("M10", "M10 Difference Interaction"),
    "interaction_ratio_reciprocal": ("M11", "M11 Ratio Reciprocal Pair"),
    "origin_explicit_reference": ("M12", "M12 Explicit Reference"),
    "origin_default_reference": ("M13", "M13 Default Reference"),
    # The shipped case named model_year_origin_categorical keeps three
    # Continuous predictors alongside its two Categoricals, so it is NOT the
    # plan's "categorical-only design" — that corner is M14b, added
    # alongside it. See docs/MODEL_TESTING_ASSETS.md § 1.1.
    "model_year_origin_categorical": ("M14", "M14 Mixed Cat And Continuous"),
    "categorical_only_design": ("M14b", "M14b Categorical Only Design"),
    "usa_filter_degenerate_origin": ("M15", "M15 Filter Degenerate Cat"),
    # M16 (typed Sequence Period override) and P07 (irregular panel
    # spacing) are NOT here. Both fit exactly the model a neighbouring case
    # already fits — M16 is M01's fit, P07 is P02's — and everything they
    # actually test lives in the spec block's status cells (the Period In
    # Use display, the Sequence Verdict, the Δ spectrum). Registering them
    # as fittable cases would add two identical fits and violate the
    # covering-array rule, so they are guard-state cases instead. See
    # lambda_catalog/analyze_regression_guard_states.py.
    # § 1.2 Life Expectancy — transform dispatch, scale, missingness.
    "life_partial_linear_log": ("L01", "L01 Partial Linear-Log"),
    "life_log_response_duan": ("L02", "L02 Log Response Duan"),
    "life_log_response_naive": ("L03", "L03 Log Response Naive"),
    "life_elasticity_log_log": ("L04", "L04 Elasticity Log-Log"),
    "life_full_profile": ("L05", "L05 Kitchen Sink Profile"),
    # L07 is NOT here. At k = 205 the workbook cannot invert the Gram
    # matrix and every engine cell reads nan, which is precisely the state
    # the width guard warns about — so a numeric oracle for it would be
    # comparing against nothing. It ships as a guard-state case asserting
    # the M2 WARNING instead. See analyze_regression_guard_states.py.
    "life_country_fixed_effects": ("L08", "L08 High Cardinality FE"),
    "life_status_explicit_reference": ("L09", "L09 Binary Cat Reference"),
    # § 1.3 Production Lots — learning curves, fixed effects, sequence.
    "production_lots_fixed_effects": ("P01", "P01 Learning Curve FE"),
    "production_lots_log_transform": ("P02", "P02 FE Log Transform Axis"),
    # P03/P03b are a pair: the same no-FE power law reached two ways. The
    # sheet names say WHICH ROUTE, because that is the only thing that
    # differs between the two tabs and the whole reason both exist.
    "production_lots_derived_log_no_fe": ("P03", "P03 Power Law Derived Cols"),
    "production_lots_log_no_fe": ("P03b", "P03b Power Law Transform Axis"),
    "production_lots_log_mixed_predictors": ("P04", "P04 Log Mixed Predictors"),
    "production_lots_log_predictor_only": ("P05", "P05 Log Predictor Only"),
    "production_lots_lsdv_equivalence": ("P06", "P06 LSDV vs Within Estimator"),
    # § 1.4 — G8 is the one guard-rail row that IS a fittable model (the
    # invalid reference degrades to zero columns rather than erroring), so
    # it lives with the fittable cases and carries a G-tier sheet name.
    "origin_invalid_reference": ("G08", "G08 Invalid Reference Level"),
}

# Cases whose sheets are too expensive to build by default — see
# RegressionSpecCase.heavy. Kept as a set next to the identity table so the
# two facts about "which cases are special" read together.
#
# Two reasons land a case here. The first is the obvious one — L08's 173
# Fixed Effects groups make every per-row residual / leverage / Cook's
# calculation 173× wider than the next case, and the sheet build is what
# hurts, not the Python oracle. The second is the one L05 occupies: at
# k = 19 with n = 2117 and ~5% missingness on every predictor, the
# statsmodels OLS reference and Excel's OLS implementation diverge in
# the 5th–6th decimal place on most coefficients and residuals — not
# because either side is wrong, but because the QR-with-column-pivoting
# paths they each take through an ill-conditioned Gram matrix produce
# near-tied numerics. L05's 73 mismatches on the regular run are that
# floor, not a defect. A LOOSER tolerance would also work, but the
# right thing to do with "this case demonstrates a precision floor the
# production shipped formula can't go below" is keep it as a deliberate
# showcase behind --include-heavy rather than paper over it on every
# run.
_HEAVY_CASE_NAMES = frozenset({
    "life_country_fixed_effects",
    "life_full_profile",
})


def _identify(case: RegressionSpecCase) -> RegressionSpecCase:
    """Attach the plan ID / sheet name / heavy flag to a freshly built case.

    Applied by ``build_regression_spec_cases`` to every case so the identity
    table above is the single place those three facts are declared, rather
    than three more keyword arguments at each of ~37 construction sites.
    """
    plan_id, sheet_name = _CASE_SHEET_IDENTITY[case.name]
    validate_sheet_name(sheet_name)
    return replace(
        case,
        plan_id=plan_id,
        sheet_name=sheet_name,
        heavy=case.name in _HEAVY_CASE_NAMES,
    )


def _production_lots_lsdv_equivalence_spec() -> list[SpecVariable]:
    """P6 — production_lots_log_transform with Facility as a Categorical Predictor.

    The strongest cheap oracle in the suite. This is P2's spec with exactly
    one edit — ``Facility`` declared as a Categorical Predictor rather than
    Role=Fixed Effects — which makes it the least-squares dummy-variable
    form of the same model. LSDV and the within estimator are algebraically
    identical on the slope coefficients and the residual vector, so
    ``tests/test_regression_spec_qc.py`` can assert P6 == P2 to floating
    point WITHOUT either side reading the workbook.

    That matters because the FE path is the one piece of the engine with no
    independent implementation to check against: everything else is OLS,
    which statsmodels also does. Fixed Effects demeaning, the absorbed-df
    subtraction, and the level-shift recovery are bespoke, and until now
    their only oracle was a second copy of the same arithmetic. This case
    fits the same model by a completely different route.

    The two do NOT agree on everything, and the disagreements are the
    point: LSDV spends its degrees of freedom visibly (k = 3 columns:
    Ln(Cumulative_Units) plus two Facility dummies) where FE absorbs them,
    so R2, the intercept, and the coefficient count differ by construction.
    Only the slope and residuals are claimed equal.
    """
    return [
        _spec_var("Lot_ID", _ROLE_IDENTIFIER),
        _spec_var("Facility", _ROLE_PREDICTOR, True, "Categorical"),
        _spec_var("Fiscal_Year", _ROLE_OMIT, sequence=True),
        _spec_var("Lot_Quantity", _ROLE_OMIT),
        _spec_var(
            "Cumulative_Units", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
        _spec_var("Experience_Stock", _ROLE_OMIT),
        _spec_var("Unit_Cost_BY", _ROLE_RESPONSE, transform="Log"),
        _spec_var("log Cum Units", _ROLE_OMIT),
        _spec_var("log experience", _ROLE_OMIT),
        _spec_var("log Unit Cost", _ROLE_OMIT),
        _spec_var("Full_Data", _ROLE_FILTER),
    ]


def build_regression_spec_cases() -> list[RegressionSpecCase]:
    """Return the standard human-plan-core spec cases for QC.

    Every case carries its plan ID and worksheet name (from
    ``_CASE_SHEET_IDENTITY``), validated here against the naming contract in
    ``lambda_catalog/test_model_sheets.py`` — so an illegal or duplicated
    sheet name fails in the unit suite, not partway through a multi-minute
    Excel build.
    """
    cases: list[RegressionSpecCase] = []

    for allow in (True, False):
        suffix = "intercept" if allow else "no_intercept"
        cases.append(
            RegressionSpecCase(
                name=f"default_t0_{suffix}",
                spec=tuple(build_default_spec()),
                allow_intercept=allow,
            )
        )
        cases.append(
            RegressionSpecCase(
                name=f"v1_full_continuous_{suffix}",
                spec=tuple(_v1_full_continuous_spec()),
                allow_intercept=allow,
            )
        )
        cases.append(
            RegressionSpecCase(
                name=f"continuous_subset_{suffix}",
                spec=tuple(_continuous_subset_spec()),
                allow_intercept=allow,
            )
        )

    categorical_specs = [
        ("origin_default_reference", _with_origin(_continuous_subset_spec())),
        ("origin_explicit_reference", _with_origin(_continuous_subset_spec(), "Europe")),
        ("origin_invalid_reference", _with_origin(_continuous_subset_spec(), 99)),
        ("model_year_origin_categorical", _model_year_origin_categorical_spec()),
        (
            "usa_filter_degenerate_origin",
            [
                *build_default_spec(),
                _spec_var("Is_USA", _ROLE_FILTER, False, "Continuous"),
            ],
        ),
    ]
    # v3.1 interaction wiring — the three width regimes the constructor has
    # to get right: one column, the documented quadratic (a row pointing at
    # itself), and the Continuous x Categorical broadcast to L-1 columns.
    categorical_specs.extend([
        ("interaction_continuous_product", _interaction_spec("Displacement")),
        ("interaction_quadratic_self_product", _interaction_spec("Weight")),
        (
            "interaction_categorical_broadcast",
            _interaction_spec("Origin", categorical_operand=True),
        ),
    ])
    # The remaining Auto MPG rows of docs/MODEL_TESTING_ASSETS.md § 1.1:
    # the two interaction operations the closed vocabulary still had no
    # case for, the Cat x Cat width regime and its interaction-free base,
    # and (Log, Log) combined with real missingness.
    categorical_specs.extend([
        ("mileage_log_log_na_masking", _mileage_log_log_na_masking_spec()),
        ("categorical_only_design", _categorical_only_design_spec()),
        ("interaction_categorical_cross", _interaction_categorical_cross_spec()),
        ("interaction_difference", _interaction_difference_spec()),
        ("interaction_ratio_reciprocal", _interaction_ratio_reciprocal_spec()),
    ])

    for name, spec in categorical_specs:
        extra = (_IS_USA,) if name == "usa_filter_degenerate_origin" else ()
        cases.append(
            RegressionSpecCase(
                name=name,
                spec=tuple(spec),
                allow_intercept=True,
                extra_columns=extra,
            )
        )

    cases.append(
        RegressionSpecCase(
            name="production_lots_fixed_effects",
            spec=tuple(_production_lots_fixed_effects_spec()),
            allow_intercept=True,
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            # Explicit (not the alphabetically-first default) — exercises the
            # harness actually writing a non-default group into $AK$12, not
            # just accepting whatever the sheet defaults to.
            prediction_group="Site B",
            # Declares what the panel actually is: annual lots, Δ = 1. This
            # is the pair that makes the BFN panel Durbin-Watson cell live —
            # Base_Period_Delta() is the TYPED-override accessor and returns
            # #N/A when nothing is typed, so without this the AE12 cell on
            # every Fixed Effects sheet sits at #N/A and the panel
            # diagnostic is compared against nothing. P01/P02 are the
            # natural home: Fiscal_Year is a real, evenly spaced annual
            # axis, so 1 is a true statement about the data rather than
            # wiring for its own sake.
            sequence_period=1.0,
        )
    )
    cases.append(
        RegressionSpecCase(
            name="production_lots_log_transform",
            spec=tuple(_production_lots_log_transform_spec()),
            allow_intercept=True,
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            prediction_group="Site B",
            # Matches P01 — the pair must differ in exactly one thing (the
            # route to the log columns), so the period is part of what is
            # held fixed. It also extends the P01 == P02 cross-check to BFN.
            sequence_period=1.0,
        )
    )

    # The no-FE pre-derived/transform-axis PAIR, registered adjacent so the
    # two land on adjacent worksheets. Same model, two routes to it: P03
    # reads the shipped log columns, P03b computes them from the raw ones
    # via Transform=Log. P01/P02 are the same pairing with Fixed Effects;
    # having it both with and without FE is what separates a transform-axis
    # regression from an FE-demeaning one.
    cases.append(
        RegressionSpecCase(
            name="production_lots_derived_log_no_fe",
            spec=tuple(_production_lots_derived_log_no_fe_spec()),
            allow_intercept=True,
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            # Matches P03b: no Fixed Effects, so group recovery resolves to
            # "(all)". The pair must agree on this too — a differing
            # prediction group would move the prediction block and make the
            # cross-check fail for a reason that has nothing to do with
            # transforms.
            prediction_group=None,
        )
    )

    # v3.3 — three new spec cases covering the v3.3 unit-space dispatcher's
    # (Log, Log), (Log, Mixed), and (None, Log) branches. Each is a sibling of
    # production_lots_log_transform with a small spec edit. See
    # tests/test_unit_space_dispatch.py for the cross-checks against the
    # workbook arithmetic.
    cases.append(
        RegressionSpecCase(
            name="production_lots_log_no_fe",
            spec=tuple(_production_lots_log_no_fe_spec()),
            allow_intercept=True,
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            # No Fixed Effects in this spec, so group_labels is None and
            # group recovery resolves to "(all)" — leave prediction_group
            # unset and let the harness default to that.
            prediction_group=None,
        )
    )
    cases.append(
        RegressionSpecCase(
            name="production_lots_log_mixed_predictors",
            spec=tuple(_production_lots_log_mixed_predictors_spec()),
            allow_intercept=True,
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            prediction_group=None,
        )
    )
    cases.append(
        RegressionSpecCase(
            name="production_lots_log_predictor_only",
            spec=tuple(_production_lots_log_predictor_only_spec()),
            allow_intercept=True,
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            prediction_group=None,
        )
    )
    # P6 — the LSDV form of P2. Same dataset, same Filter, same prediction
    # group, so the only difference from production_lots_log_transform is
    # Facility's Role; see the spec builder for why the pair is the
    # suite's strongest cross-oracle.
    cases.append(
        RegressionSpecCase(
            name="production_lots_lsdv_equivalence",
            spec=tuple(_production_lots_lsdv_equivalence_spec()),
            allow_intercept=True,
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            # No Fixed Effects row now (Facility is an ordinary Categorical
            # Predictor), so group recovery resolves to "(all)" exactly as
            # it does for the other no-FE Production Lots cases.
            prediction_group=None,
        )
    )

    # ── Life Expectancy (§ 1.2) ─────────────────────────────────────────
    # All eight share the dataset, loader and Source_Table retarget; they
    # differ only in spec (and, for L03, in the Back-Transform toggle).
    # None declares Fixed Effects except L08, so prediction_group is left
    # to resolve to "(all)" for the rest.
    for name, spec, back_transform in (
        ("life_partial_linear_log", _life_partial_linear_log_spec(), "Duan"),
        ("life_log_response_duan", _life_log_response_spec(), "Duan"),
        # L03 is L02's spec verbatim with the sheet's $AH$4 flipped to
        # Naive: EXP(y_hat) with no smearing factor. Every unit-space
        # number moves; the CI/PI bounds do not (they are EXP-only under
        # both methods). Sharing the spec builder is what makes the
        # difference attributable to the toggle alone.
        ("life_log_response_naive", _life_log_response_spec(), "Naive"),
        ("life_elasticity_log_log", _life_elasticity_log_log_spec(), "Duan"),
        ("life_full_profile", _life_full_profile_spec(), "Duan"),
        ("life_country_fixed_effects", _life_country_fixed_effects_spec(), "Duan"),
        (
            "life_status_explicit_reference",
            _life_partial_linear_log_spec("Developing"),
            "Duan",
        ),
    ):
        cases.append(
            RegressionSpecCase(
                name=name,
                spec=tuple(spec),
                allow_intercept=True,
                source_csv_path=LIFE_EXPECTANCY_CSV_PATH,
                row_loader=load_life_expectancy_source_rows,
                source_table_ref="=LifeExpectancyData[#All]",
                prediction_group=None,
                back_transform=back_transform,
            )
        )

    identified = [_identify(case) for case in cases]
    assert_sheet_names_unique([case.sheet_name for case in identified])
    return identified


def build_regression_spec_qc_configs(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[RegressionSpecExpected]:
    """Compute expected Regression sheet outputs for all spec-driven QC cases."""
    return [
        calculate_regression_spec_case(case, csv_path)
        for case in build_regression_spec_cases()
    ]
