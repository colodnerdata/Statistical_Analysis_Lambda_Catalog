"""Spec-driven expected values for the current Regression worksheet QC harness."""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
import math
import numpy as np

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
from .write_sheet_model_construction import (
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
            spec.append(_spec_var(variable.name, _ROLE_IDENTIFIER, sequence=True))
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
        model_year=_spec_var(
            "Model Year", _ROLE_PREDICTOR, True, "Categorical", sequence=True
        ),
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


_IS_USA = ExtraSpecColumn(
    name="Is_USA",
    excel_formula='=--([@Origin]="US")',
    value_fn=lambda row: 1 if row["Origin"] == "US" else 0,
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


def _production_lots_log_no_fe_spec() -> list[SpecVariable]:
    """v3.3 — Log+Log with NO Fixed Effects: the (Log, Log) SWITCH branch.

    Sibling of _production_lots_log_transform_spec() with ``Facility``
    omitted instead of declared as Fixed Effects. Exercises the v3.3
    unit-space dispatcher's (Log, Log) branch where the level shift
    Y_Full − Y is exactly zero, so the unit-space arithmetic reduces
    to back-transforming the in-sample fit and the new (smeared) R²
    is computed cleanly. Reduction invariant: with no FE, the
    smearing factor uses raw residuals, not within residuals.
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


def build_regression_spec_cases() -> list[RegressionSpecCase]:
    """Return the standard human-plan-core spec cases for QC."""
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

    return cases


def build_regression_spec_qc_configs(
    csv_path: Path = DEFAULT_INPUT_CSV,
) -> list[RegressionSpecExpected]:
    """Compute expected Regression sheet outputs for all spec-driven QC cases."""
    return [
        calculate_regression_spec_case(case, csv_path)
        for case in build_regression_spec_cases()
    ]
