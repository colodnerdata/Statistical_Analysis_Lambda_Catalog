"""Spec-driven expected values for the current Regression worksheet QC harness."""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
import numpy as np

from .analyze_life_expectancy import DEFAULT_INPUT_CSV
from .analyze_model_construction import (
    SpecVariable,
    _compute_mask,
    _format_value,
    _is_blank,
    _is_number,
    _retained_levels,
    build_default_spec,
    build_default_spec_for_headers,
    calculate_model_construction_expectations,
    load_auto_mpg_rows,
    load_source_rows,
)
from .analyze_regression_sheet import calculate_regression_results_from_matrix
from .regression_shared import FEATURE_COLUMNS, RegressionSheetResults
from .write_sheet_auto_mpg_data import DEFAULT_AUTO_MPG_XLSX_PATH
from .write_sheet_model_construction import (
    _ROLE_FILTER,
    _ROLE_IDENTIFIER,
    _ROLE_OMIT,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
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


@dataclass(frozen=True)
class RegressionSpecDesign:
    """The spec-derived arrays and display facts behind one expected result."""

    row_mask: tuple[bool, ...]
    row_labels: tuple[str, ...]
    constructed_column_names: tuple[str, ...]
    x_features: np.ndarray
    y_train: np.ndarray
    sequence_values: np.ndarray | None
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


def _copy_spec(spec: tuple[SpecVariable, ...] | list[SpecVariable]) -> list[SpecVariable]:
    return [
        SpecVariable(
            item.name,
            item.role,
            item.include,
            item.var_type,
            item.reference,
            item.sequence,
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
    reference: str = "",
    sequence: bool = False,
) -> SpecVariable:
    return SpecVariable(name, role, include, var_type, reference, sequence)


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
        y_values.append(float(value))
        labels.append(_row_label(rows[idx], identifiers, idx))

    matrix_columns: list[list[float]] = []
    constructed_names: list[str] = []
    for variable in spec_tuple:
        if variable.role != _ROLE_PREDICTOR or not variable.include:
            continue
        if variable.var_type != "Categorical":
            constructed_names.append(variable.name)
            matrix_columns.append(
                [_numeric_cell(rows[idx], variable.name) for idx in included_indices]
            )
            continue

        retained = _retained_levels(variable, rows, list(mask))
        if retained is None:
            continue
        for level in retained:
            constructed_names.append(f"{variable.name}: {_format_value(level)}")
            matrix_columns.append([
                0.0 if _is_blank(rows[idx][variable.name])
                else (1.0 if rows[idx][variable.name] == level else 0.0)
                for idx in included_indices
            ])

    if not matrix_columns:
        raise ValueError("Spec produced zero constructed columns")
    x_features = np.asarray(matrix_columns, dtype=np.float64).T

    seq_variables = [item.name for item in spec_tuple if item.sequence]
    sequence_values = None
    if len(seq_variables) == 1:
        sequence_name = seq_variables[0]
        seq_cells = [rows[idx][sequence_name] for idx in included_indices]
        if all(_is_number(value) for value in seq_cells):
            sequence_values = np.asarray(
                [float(value) for value in seq_cells],
                dtype=np.float64,
            )

    expectations = calculate_model_construction_expectations(list(spec_tuple), rows)
    return RegressionSpecDesign(
        row_mask=mask,
        row_labels=tuple(labels),
        constructed_column_names=tuple(constructed_names),
        x_features=x_features,
        y_train=np.asarray(y_values, dtype=np.float64),
        sequence_values=sequence_values,
        included_rows=len(included_indices),
        level_counts=expectations.level_counts,
        references_in_use=expectations.references_in_use,
        degenerate_categoricals=expectations.degenerate_categoricals,
    )


def calculate_regression_spec_case(
    case: RegressionSpecCase,
    csv_path: Path = DEFAULT_INPUT_CSV,
    *,
    source_rows: list[dict[str, object]] | None = None,
) -> RegressionSpecExpected:
    """Compute expected current Regression sheet outputs for one spec case."""
    base_rows = source_rows if source_rows is not None else load_source_rows(csv_path)
    rows = _with_extra_columns(base_rows, case.extra_columns)
    design = build_spec_design(case.spec, rows)
    results = calculate_regression_results_from_matrix(
        x_features=design.x_features,
        y_train=design.y_train,
        predictor_names=design.constructed_column_names,
        include_intercept=case.allow_intercept,
        alpha=case.alpha,
        sequence_values=design.sequence_values,
        allow_singular=source_rows is not None,
    )
    return RegressionSpecExpected(case=case, design=design, results=results)


def _v1_full_continuous_spec() -> list[SpecVariable]:
    numeric_predictors = set(FEATURE_COLUMNS)
    spec = []
    for variable in build_default_spec():
        if variable.name == "Country":
            spec.append(_spec_var(variable.name, _ROLE_IDENTIFIER))
        elif variable.name == "Year":
            spec.append(_spec_var(variable.name, _ROLE_IDENTIFIER, sequence=True))
        elif variable.name == "Status":
            spec.append(_spec_var(variable.name, _ROLE_OMIT))
        elif variable.name == "Life expectancy":
            spec.append(_spec_var(variable.name, _ROLE_RESPONSE))
        elif variable.name == "Full_Data":
            spec.append(_spec_var(variable.name, _ROLE_FILTER))
        elif variable.name in numeric_predictors:
            spec.append(_spec_var(variable.name, _ROLE_PREDICTOR, True, "Continuous"))
        else:
            spec.append(_spec_var(variable.name, _ROLE_OMIT))
    return spec


def _continuous_subset_spec() -> list[SpecVariable]:
    selected = {"Adult Mortality", "GDP", "Schooling"}
    spec = []
    for variable in _v1_full_continuous_spec():
        if variable.role == _ROLE_PREDICTOR and variable.name not in selected:
            spec.append(_spec_var(variable.name, _ROLE_PREDICTOR, False, variable.var_type))
        else:
            spec.append(variable)
    return spec


def _with_status(spec: list[SpecVariable], reference: str = "") -> list[SpecVariable]:
    return _replace_spec_vars(
        spec,
        status=_spec_var("Status", _ROLE_PREDICTOR, True, "Categorical", reference),
    )


def _year_status_categorical_spec() -> list[SpecVariable]:
    return _replace_spec_vars(
        _with_status(_continuous_subset_spec()),
        year=_spec_var("Year", _ROLE_PREDICTOR, True, "Categorical", sequence=True),
    )


_IS_DEVELOPING = ExtraSpecColumn(
    name="Is_Developing",
    excel_formula='=--([@Status]="Developing")',
    value_fn=lambda row: 1 if row["Status"] == "Developing" else 0,
)


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
        ("status_default_reference", _with_status(_continuous_subset_spec())),
        ("status_explicit_reference", _with_status(_continuous_subset_spec(), "Developing")),
        ("status_invalid_reference", _with_status(_continuous_subset_spec(), "Developped")),
        ("year_status_categorical", _year_status_categorical_spec()),
        (
            "developing_filter_degenerate_status",
            [
                *build_default_spec(),
                _spec_var("Is_Developing", _ROLE_FILTER, False, "Continuous"),
            ],
        ),
    ]
    for name, spec in categorical_specs:
        extra = (_IS_DEVELOPING,) if name == "developing_filter_degenerate_status" else ()
        cases.append(
            RegressionSpecCase(
                name=name,
                spec=tuple(spec),
                allow_intercept=True,
                extra_columns=extra,
            )
        )

    return cases


def _build_auto_mpg_spec_cases(
    source_rows: list[dict[str, object]],
) -> list[RegressionSpecCase]:
    """Return default-spec QC cases for the shipped Auto MPG regression target."""
    headers = list(source_rows[0].keys())
    default_spec = build_default_spec_for_headers(headers)
    by_name = {item.name: item for item in default_spec}

    def _set(name: str, role: str, include: bool, var_type: str) -> None:
        if name not in by_name:
            return
        by_name[name] = _spec_var(name, role, include, var_type)

    # Use a numerically stable single-predictor case on the same dataset so
    # Python/scipy and Excel stay directly comparable in QC.
    for header in headers:
        _set(header, _ROLE_OMIT, False, "Continuous")
    _set("Car Name", _ROLE_IDENTIFIER, False, "Continuous")
    _set("MPG", _ROLE_RESPONSE, False, "Continuous")
    _set("Cylinders", _ROLE_PREDICTOR, True, "Continuous")
    stable_spec = tuple(by_name[header] for header in headers)

    return [
        RegressionSpecCase(
            name="auto_mpg_cylinders_intercept",
            spec=stable_spec,
            allow_intercept=True,
        )
    ]


def build_regression_spec_qc_configs(
    csv_path: Path = DEFAULT_INPUT_CSV,
    *,
    regression_dataset: str = "life_expectancy",
    auto_mpg_workbook_path: Path = DEFAULT_AUTO_MPG_XLSX_PATH,
) -> list[RegressionSpecExpected]:
    """Compute expected Regression sheet outputs for all spec-driven QC cases."""
    if regression_dataset == "auto_mpg":
        source_rows = load_auto_mpg_rows(auto_mpg_workbook_path)
        return [
            calculate_regression_spec_case(case, csv_path, source_rows=source_rows)
            for case in _build_auto_mpg_spec_cases(source_rows)
        ]
    return [
        calculate_regression_spec_case(case, csv_path)
        for case in build_regression_spec_cases()
    ]
