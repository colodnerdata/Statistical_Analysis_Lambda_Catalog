"""v3.1 interaction wiring — the constructor's Python mirror.

The spec block's M/N pair (Interaction Term / Interaction Operation) shipped
reserved-and-unwired at v3.0 stage 3; v3.1 teaches ``Predictor_Columns()`` to
build the columns they declare. These tests pin the semantics DECISIONS.md
settled — the closed operation vocabulary and its symmetry attribute, the
four operand Role/Include cases, the two-way limit, the documented quadratic,
and the width regimes the Design Columns audit has to report — against the
Python mirror in ``analyze_regression_spec.build_spec_design``, which is the
same oracle the Excel gate compares the sheet to.
"""
# pylint: disable=missing-function-docstring
from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest

from lambda_catalog.analyze_model_construction import (
    SpecVariable,
    load_source_rows,
    resolve_interaction_operand,
)
from lambda_catalog.analyze_regression_spec import (
    _continuous_subset_spec,
    _replace_spec_vars,
    _spec_var,
    _with_origin,
    build_spec_design,
)
from lambda_catalog.write_sheet_model_construction import (
    _ROLE_IDENTIFIER,
    _ROLE_OMIT,
    _ROLE_PREDICTOR,
)

ROOT_DIR = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT_DIR / "sample_data" / "auto_mpg_data.csv"

pytestmark = pytest.mark.skipif(
    not CSV_PATH.exists(), reason="Auto MPG CSV not found"
)


@pytest.fixture(name="rows", scope="module")
def _rows() -> list[dict[str, object]]:
    return load_source_rows()


def _declare(
    term: str,
    operation: str = "Product",
    *,
    on: str = "Weight",
    categorical_operand: bool = False,
    include: bool = True,
) -> list[SpecVariable]:
    """Declare an interaction on one row of the continuous-subset spec."""
    base = (
        _with_origin(_continuous_subset_spec())
        if categorical_operand
        else _continuous_subset_spec()
    )
    return _replace_spec_vars(
        base,
        target=_spec_var(
            on,
            _ROLE_PREDICTOR,
            include,
            "Continuous",
            interaction_term=term,
            interaction_operation=operation,
        ),
    )


def _column(design, name: str) -> np.ndarray:
    return design.x_features[:, design.constructed_column_names.index(name)]


# ── Width regimes ────────────────────────────────────────────────────────────

def test_continuous_by_continuous_is_one_column(rows) -> None:
    design = build_spec_design(_declare("Displacement"), rows)

    assert design.constructed_column_names == (
        "Displacement",
        "Horsepower",
        "Weight",
        "Weight:Displacement",
    )
    assert np.allclose(
        _column(design, "Weight:Displacement"),
        _column(design, "Weight") * _column(design, "Displacement"),
    )


def test_continuous_by_categorical_broadcasts_to_l_minus_one(rows) -> None:
    # Origin has three levels; one is dropped as the reference, so the
    # interaction contributes L-1 = 2 columns from a single spec row. This
    # is the width behaviour the Design Columns audit exists to surface.
    design = build_spec_design(
        _declare("Origin", categorical_operand=True), rows
    )

    interaction = [
        name for name in design.constructed_column_names if name.startswith("Weight:")
    ]
    assert interaction == ["Weight:Origin: Europe", "Weight:Origin: US"]
    for name in interaction:
        level = name.split(":", 1)[1]
        assert np.allclose(
            _column(design, name),
            _column(design, "Weight") * _column(design, level),
        )


def test_interaction_columns_follow_their_own_row_not_the_end_of_the_matrix(
    rows,
) -> None:
    # Weight precedes Origin in the source table, so Weight's interaction
    # columns land between Weight and Origin's own dummies. Appending them
    # at the end instead would still be k-correct but would break the
    # spec-order reading the sheet's residual/summary zones rely on.
    design = build_spec_design(
        _declare("Origin", categorical_operand=True), rows
    )

    assert design.constructed_column_names == (
        "Displacement",
        "Horsepower",
        "Weight",
        "Weight:Origin: Europe",
        "Weight:Origin: US",
        "Origin: Europe",
        "Origin: US",
    )


# ── The closed operation vocabulary ──────────────────────────────────────────

def test_product_difference_and_ratio_each_compute_their_own_arithmetic(
    rows,
) -> None:
    expectations = {
        "Product": lambda a, b: a * b,
        "Difference": lambda a, b: a - b,
        "Ratio": lambda a, b: a / b,
    }
    for operation, apply in expectations.items():
        design = build_spec_design(_declare("Displacement", operation), rows)
        assert np.allclose(
            _column(design, "Weight:Displacement"),
            apply(_column(design, "Weight"), _column(design, "Displacement")),
        ), operation


def test_an_unknown_operation_raises_rather_than_silently_skipping(rows) -> None:
    # The N dropdown is closed, so this is unreachable from the sheet — but
    # the mirror must not quietly emit a wrong column if it ever is.
    with pytest.raises(ValueError, match="Unknown interaction operation"):
        build_spec_design(_declare("Displacement", "Sum"), rows)


def test_ratio_by_a_zero_denominator_is_refused_not_emitted_as_inf(rows) -> None:
    # The sheet returns NA() here rather than a bare #DIV/0!. A QC case has
    # to describe a fully computable model, so the mirror raises instead of
    # putting inf/NaN into a design matrix.
    # No shipped Auto MPG column is all-zero, so the denominator is a
    # synthetic excluded Predictor — which also covers the amber
    # excluded-operand path reaching the arithmetic.
    spec = _replace_spec_vars(
        _continuous_subset_spec(),
        target=_spec_var(
            "Weight",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term="Zero_Col",
            interaction_operation="Ratio",
        ),
    )
    spec.append(_spec_var("Zero_Col", _ROLE_PREDICTOR, False, "Continuous"))
    zeroed = [{**row, "Zero_Col": 0} for row in rows]

    with pytest.raises(ValueError, match="divides by zero"):
        build_spec_design(spec, zeroed)


# ── Operand Role / Include semantics — the four cases ────────────────────────

def test_an_excluded_predictor_operand_still_builds_the_interaction(rows) -> None:
    # The marginality violation: an interaction without its main effect.
    # Flagged amber on the sheet, never blocked — blocking it would be the
    # library deciding a modelling question on the user's behalf.
    spec = _replace_spec_vars(
        _declare("Cylinders"),
        excluded=_spec_var("Cylinders", _ROLE_PREDICTOR, False, "Continuous"),
    )
    design = build_spec_design(spec, rows)

    assert "Cylinders" not in design.constructed_column_names
    assert "Weight:Cylinders" in design.constructed_column_names


def test_a_non_predictor_operand_contributes_nothing(rows) -> None:
    # Only a Predictor may be an operand. The sheet flags this red; the
    # constructor degrades to the main effect rather than erroring the whole
    # matrix, the same way an invalid Reference Level degrades.
    for role in (_ROLE_OMIT, _ROLE_IDENTIFIER):
        spec = _replace_spec_vars(
            _declare("Cylinders"),
            other=_spec_var("Cylinders", role),
        )
        design = build_spec_design(spec, rows)
        assert not [
            name for name in design.constructed_column_names if ":" in name
        ], role


def test_an_unmatched_operand_name_contributes_nothing(rows) -> None:
    design = build_spec_design(_declare("Not A Column"), rows)

    assert design.constructed_column_names == (
        "Displacement",
        "Horsepower",
        "Weight",
    )


def test_an_excluded_declaring_row_contributes_nothing_at_all(rows) -> None:
    # Cascading relevance: Include = FALSE on the declaring row drops its
    # main effect AND its interaction, not just the main effect.
    design = build_spec_design(_declare("Displacement", include=False), rows)

    assert design.constructed_column_names == ("Displacement", "Horsepower")


@pytest.mark.parametrize("missing", ["term", "operation"])
def test_both_m_and_n_are_required(rows, missing: str) -> None:
    spec = _declare("Displacement")
    spec = _replace_spec_vars(
        spec,
        target=_spec_var(
            "Weight",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term="" if missing == "term" else "Displacement",
            interaction_operation="" if missing == "operation" else "Product",
        ),
    )
    design = build_spec_design(spec, rows)

    assert not [name for name in design.constructed_column_names if ":" in name]


# ── The documented quadratic, and the two-way limit ──────────────────────────

def test_a_row_pointing_at_itself_under_product_is_the_quadratic_term(
    rows,
) -> None:
    design = build_spec_design(_declare("Weight"), rows)

    assert "Weight:Weight" in design.constructed_column_names
    assert np.allclose(
        _column(design, "Weight:Weight"), _column(design, "Weight") ** 2
    )


def test_an_operand_that_declares_its_own_interaction_contributes_only_its_block(
    rows,
) -> None:
    # Two-way only: (AxB)xC is not expressible. Declaring an interaction on
    # the OPERAND must not nest — the operand contributes its main-effect
    # block to the pairing, and its own interaction to its own position.
    spec = _replace_spec_vars(
        _declare("Displacement"),
        chained=_spec_var(
            "Displacement",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term="Horsepower",
            interaction_operation="Product",
        ),
    )
    design = build_spec_design(spec, rows)

    assert design.constructed_column_names == (
        "Displacement",
        "Displacement:Horsepower",
        "Horsepower",
        "Weight",
        "Weight:Displacement",
    )
    # The pairing used Displacement's own column, NOT its interaction column.
    assert np.allclose(
        _column(design, "Weight:Displacement"),
        _column(design, "Weight") * _column(design, "Displacement"),
    )


# ── Transform interaction ────────────────────────────────────────────────────

def test_a_logged_operand_is_combined_after_its_own_log(rows) -> None:
    # The transform is a property of the operand's column and is applied
    # before the two are combined — so this is Weight * ln(Displacement),
    # never ln(Weight * Displacement).
    spec = _replace_spec_vars(
        _declare("Ln(Displacement)"),
        logged=_spec_var(
            "Displacement", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
    )
    # M names the SOURCE column, not the relabelled constructed column.
    spec = _replace_spec_vars(
        spec,
        target=_spec_var(
            "Weight",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term="Displacement",
            interaction_operation="Product",
        ),
    )
    design = build_spec_design(spec, rows)

    assert "Weight:Ln(Displacement)" in design.constructed_column_names
    assert np.allclose(
        _column(design, "Weight:Ln(Displacement)"),
        _column(design, "Weight") * _column(design, "Ln(Displacement)"),
    )


def test_every_interaction_column_reads_transform_none(rows) -> None:
    # Constructed_Column_Transforms must stay exactly as wide as the design
    # matrix, and an interaction column is never itself Log-flagged: the
    # transform already ran on each operand.
    spec = _replace_spec_vars(
        _declare("Displacement"),
        logged=_spec_var(
            "Displacement", _ROLE_PREDICTOR, True, "Continuous", transform="Log"
        ),
    )
    design = build_spec_design(spec, rows)

    assert len(design.constructed_column_transforms) == len(
        design.constructed_column_names
    )
    flags = dict(
        zip(design.constructed_column_names, design.constructed_column_transforms)
    )
    assert flags["Ln(Displacement)"] == "Log"
    assert flags["Weight:Ln(Displacement)"] == "None"


# ── mate(), the operand resolver ─────────────────────────────────────────────

def test_resolve_interaction_operand_mirrors_the_sheets_gating() -> None:
    spec = [
        _spec_var("A", _ROLE_PREDICTOR, True, "Continuous"),
        _spec_var("B", _ROLE_OMIT),
    ]

    def declare(term: str, operation: str) -> SpecVariable:
        return _spec_var(
            "A",
            _ROLE_PREDICTOR,
            True,
            "Continuous",
            interaction_term=term,
            interaction_operation=operation,
        )

    assert resolve_interaction_operand(declare("A", "Product"), spec) == 0
    assert resolve_interaction_operand(declare("", "Product"), spec) is None
    assert resolve_interaction_operand(declare("A", ""), spec) is None
    assert resolve_interaction_operand(declare("nope", "Product"), spec) is None
    # B exists but is Role=Omit.
    assert resolve_interaction_operand(declare("B", "Product"), spec) is None
