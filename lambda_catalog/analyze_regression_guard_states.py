"""Python oracles for the spec block's guard-rail states.

``docs/MODEL_TESTING_ASSETS.md`` § 1.4 lists a dozen configurations that are
deliberately NOT fittable models — a spec with no Response, one with two
Sequence flags, an interaction naming a column that does not exist. They
verify status lines, conditional formatting and graceful degradation, and
they have never had an oracle of any kind.

They cannot reuse ``RegressionSpecCase``. ``calculate_regression_spec_case``
raises on most of them by design (zero or two Response rows, a zero-column
design, a ``#N/A`` transform), because a fittable case must describe a legal,
fully-computable model. What a guard case asserts is a different set of
facts entirely: the text in a status cell, the per-row Design Columns audit,
and which cells the conditional formatting flags and how severely.

**Flags are predicates, not pixels.** ``GuardFlag`` records that a rule
*fires*, recomputed here from the same condition the CF expression encodes —
not read back out of Excel as ``DisplayFormat.Interior.Color``. Reading the
color would only report what Excel already decided from the rule; recomputing
the predicate is what makes a silent change to a CF expression fail. The
severity vocabulary is the two the sheet uses: ``"red"`` (an error the user
must fix) and ``"amber"`` (advisory — flagged and allowed).

Three cases here are not from § 1.4 at all, and are guard states for a
reason worth recording:

* **L06** (``Ln`` of a column with true zeros) is listed in § 1.2 as a
  fittable model whose zero rows "drop out of the mask". They do not — see
  ``_LN_ZERO_GUARD_NOTE`` below.
* **M16** (typed Sequence Period override) and **P07** (irregular panel
  spacing) fit exactly the models M01 and P02 already fit. Everything they
  actually test is in the spec block's status cells, so registering them as
  fittable cases would add two duplicate fits and buy nothing.
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from .analyze_life_expectancy import (
    DEFAULT_INPUT_CSV as LIFE_EXPECTANCY_CSV_PATH,
    load_life_expectancy_source_rows,
)
from .analyze_model_construction import (
    SpecVariable,
    _compute_mask,
    _is_blank,
    _is_number,
    _retained_levels,
    build_default_spec,
    calculate_model_construction_expectations,
    load_source_rows,
)
from .analyze_regression_sheet import _build_model_formula
from .analyze_production_lots import (
    DEFAULT_INPUT_CSV as PRODUCTION_LOTS_CSV_PATH,
    load_production_lots_source_rows,
)
from .analyze_regression_spec import (
    ExtraSpecColumn,
    _life_country_width_guard_spec,
    _life_spec,
    _spec_var,
)
from .test_model_sheets import assert_sheet_names_unique, validate_sheet_name
from .write_sheet_model_construction import (
    _FIRST_DATA_ROW as SPEC_FIRST_DATA_ROW,
    _MSG_CALENDAR,
    _MSG_NO_NATURAL,
    _MSG_OFF_GRID,
    _MSG_REGULARITY,
    _ROLE_FIXED_EFFECTS,
    _ROLE_IDENTIFIER,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
)
from .write_sheet_regression import (
    _DESIGN_MATRIX_MAX_COLUMNS,
    _DESIGN_MATRIX_SOFT_CELLS,
    _DESIGN_MATRIX_SOFT_COLUMNS,
)

# The empty-model string the audit strip's IFERROR degrades to, matching
# _EMPTY_MODEL_FALLBACK in write_sheet_model_construction. Kept as a plain
# string here (that constant carries Excel's own quoting).
EMPTY_MODEL = "(empty model)"

# The two severities the spec block's conditional formatting uses. Red is an
# error the user has to fix; amber is advisory — the model still builds.
SEVERITY_RED = "red"
SEVERITY_AMBER = "amber"

_LN_ZERO_GUARD_NOTE = """L06's plan entry does not describe what the sheet does.

docs/MODEL_TESTING_ASSETS.md § 1.2 says a Log transform on a column with
true zeros makes the row "drop out of the mask, not a silent 0". It does
not. ``Sample_Include`` tests ``ISNUMBER(col)`` on the Response and the
included Continuous Predictors, and ``ISNUMBER(0)`` is TRUE, so a zero-valued
row stays in the sample; ``Ln_Positive`` then returns ``#N/A`` for it, and
the #N/A propagates through Predictor_Columns into every downstream
statistic. The mask has no Log-positivity term anywhere.

So the honest expectation for Schooling's 28 zeros is total #N/A
propagation, not a narrower sample — which is what this case asserts.

Adding a positivity term to ``Sample_Include`` would make the plan's
description true, and is arguably the better behaviour (it is what the
Response's own non-positive guard effectively achieves by raising). That is a
production change to a constructor every model depends on, so it is recorded
here as an open question rather than made as a side effect of writing an
oracle."""


@dataclass(frozen=True)
class GuardFlag:
    """One conditional-formatting rule firing on one cell.

    ``row`` is the sheet row for a per-spec-row flag, or ``None`` for the
    single-cell controls above the table (the intercept toggle at C2, the
    status cells on row 1/2). ``column`` is the spec-block column letter, so
    a flag reads as the cell a human would look at.
    """

    column: str
    severity: str
    rule: str
    row: int | None = None


@dataclass(frozen=True)
class GuardStateCase:
    """One guard-rail spec configuration and how to put it on a sheet."""

    name: str
    plan_id: str
    sheet_name: str
    spec: tuple[SpecVariable, ...]
    allow_intercept: bool = True
    source_csv_path: Path | None = None
    row_loader: Callable[[Path], list[dict[str, object]]] | None = None
    source_table_ref: str = "=MileageData[#All]"
    extra_columns: tuple[ExtraSpecColumn, ...] = ()
    # Typed Sequence Period overrides (spec column I), variable name -> Δ.
    # Column I is an INPUT: typing a number there replaces the Period In Use
    # cell's Base_Period_Delta_Candidate() formula, which is the whole point
    # of M16 and P07. Empty for every other case.
    sequence_period_override: dict[str, float] = field(default_factory=dict)
    # Why this configuration exists, in one line — written onto the sheet as
    # provenance so an opened tab explains itself.
    covers: str = ""


@dataclass(frozen=True)
class GuardStateExpected:
    """Every spec-block fact a guard case asserts against the live sheet."""

    case: GuardStateCase
    # Audit-strip counts (row 1, columns S onward).
    responses_count: int
    sequence_flag_count: int
    fixed_effects_count: int
    included_rows: int
    response_name: str
    # The AA2 Model Formula cell: "<response> ~ 1 + <p1> + ... | <FE>".
    # Degrades to "(empty model)" when the spec contributes no predictor
    # columns — G01's headline assertion, and the state that used to leak a
    # raw #CALC! before the IFERROR wrappers went in.
    model_formula: str
    # Audit "k": an int, or "(empty model)" when Predictor_Columns errors.
    audit_k: int | str
    # Per-spec-row Design Columns audit (column O): "" on a non-Predictor
    # row, else the column count this row contributes.
    design_columns: tuple[object, ...]
    design_columns_total: int
    # Status lines. Empty string is the legal, quiet state for all four.
    sequence_status: str        # E1 / H2
    fixed_effects_status: str   # B1
    width_guard_status: str     # M2
    verdict: str                # I2
    # Sequence layer: the Δ in effect on the flagged row (typed override or
    # the MODE candidate), and the (Δ, count) spectrum at P2.
    period_in_use: float | None
    delta_spectrum: tuple[tuple[float, int], ...]
    flags: tuple[GuardFlag, ...]


# ── Status-line message templates ────────────────────────────────────────
# Mirrored from _write_sequence_status / _write_spec_feedback /
# _write_design_matrix_width_guard. Imported where the source is a module
# constant; restated only where the writer inlines the literal.
_SEQUENCE_MULTI_FLAG_ERROR = (
    "ERROR: multiple Sequence flags (mark at most one variable)"
)
_FIXED_EFFECTS_MULTI_ROW_ERROR = (
    "ERROR: multiple Fixed Effects rows (mark at most one variable)"
)


def _sequence_deltas(
    spec: tuple[SpecVariable, ...], rows: list[dict[str, object]]
) -> list[float]:
    """Mirror ``Sequence_Deltas()``: within-group consecutive spacings.

    Three details of the catalog formula matter and are easy to get wrong:

    * Groups come from the **Identifier** columns, not from a Fixed Effects
      column — and the Sequence column itself is excluded even if it is also
      flagged Identifier. With no Identifier rows every observation is one
      ``"(all)"`` group.
    * The row filter is ``ISNUMBER(sequence) AND group label non-empty``.
      It does **not** consult ``Sample_Include``, so the spectrum describes
      the DATA's time grid rather than the current model's sample.
    * Only strictly positive within-group differences survive, after sorting
      by (group, sequence value).

    Returns an empty list when no Sequence axis is flagged, which is the
    ``COUNT(d)=0`` branch every consumer guards on. With two-plus flags —
    the G03 error state — the FIRST flagged row is used, because that is
    what ``XMATCH(TRUE, Spec_Sequence, 0)`` resolves to. The error is
    reported by the E1/H2 status line and the red H cells, not by the
    spacing layer going quiet, and an oracle that went quiet here would
    disagree with the sheet and fail a case that is behaving correctly.
    """
    sequence_names = [item.name for item in spec if item.sequence]
    if not sequence_names:
        return []
    sequence_name = sequence_names[0]
    identifiers = [
        item.name
        for item in spec
        if item.role == _ROLE_IDENTIFIER and item.name != sequence_name
    ]

    pairs: list[tuple[str, float]] = []
    for row in rows:
        value = row[sequence_name]
        if not _is_number(value):
            continue
        if identifiers:
            parts = [
                "" if _is_blank(row[name]) else str(row[name])
                for name in identifiers
            ]
            group = "\x1f".join(parts)
            if not group.replace("\x1f", ""):
                continue
        else:
            group = "(all)"
        pairs.append((group, float(value)))

    pairs.sort(key=lambda pair: (pair[0], pair[1]))
    return [
        after[1] - before[1]
        for before, after in zip(pairs, pairs[1:])
        if before[0] == after[0] and after[1] - before[1] > 0
    ]


def _mode_single(values: list[float]) -> float | None:
    """Mirror Excel's ``MODE.SNGL``: the most frequent value, or None on #N/A.

    ``statistics.mode`` is NOT this function and must not be used in its
    place. Since Python 3.8 it returns the first-encountered most common
    value for *any* non-empty input, including one where nothing repeats —
    ``statistics.mode([4.0])`` is ``4.0``. ``MODE.SNGL`` errors there, which
    is precisely the condition the "no natural base period" verdict and the
    candidate's fall-back-to-MIN both key on. Mirroring it with
    ``statistics.mode`` makes both unreachable.

    Ties go to the first-occurring value, which is also what MODE.SNGL does.
    """
    counts: dict[float, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    if not counts or max(counts.values()) < 2:
        return None
    best = max(counts.values())
    return next(value for value in values if counts[value] == best)


def base_period_delta_candidate(deltas: list[float]) -> float | None:
    """Mirror ``Base_Period_Delta_Candidate()``: MODE, else MIN, else None.

    ``None`` stands for the ``#N/A`` the catalog function returns when there
    are no spacings at all. When there ARE spacings but none repeats,
    ``MODE.SNGL`` errors and the catalog's ``IFERROR`` falls back to ``MIN``.
    """
    if not deltas:
        return None
    mode = _mode_single(deltas)
    return float(mode) if mode is not None else float(min(deltas))


def delta_spectrum(deltas: list[float]) -> tuple[tuple[float, int], ...]:
    """Mirror ``Sequence_Delta_Spectrum()``: sorted unique Δ with counts."""
    return tuple(
        (value, deltas.count(value)) for value in sorted(set(deltas))
    )


def sequence_verdict(deltas: list[float], period_in_use: float | None) -> str:
    """Mirror the I2 verdict switch, including its priority order.

    Blank when there is no axis or the Period In Use cell holds no number.
    Otherwise the FIRST matching branch wins, most severe first: off-grid,
    then regularity, then no-natural-period, then the calendar signature.
    """
    if not deltas or period_in_use is None:
        return ""
    if any(delta % period_in_use != 0 for delta in deltas):
        return _MSG_OFF_GRID
    if any(delta != period_in_use for delta in deltas):
        return _MSG_REGULARITY
    if _mode_single(deltas) is None:
        return _MSG_NO_NATURAL
    calendar_like = sum(
        1
        for delta in deltas
        if 28 <= delta <= 31 or 90 <= delta <= 92 or 365 <= delta <= 366
    )
    if calendar_like * 2 > len(deltas):
        return _MSG_CALENDAR
    return ""


def _row_design_columns(
    variable: SpecVariable,
    spec: tuple[SpecVariable, ...],
    rows: list[dict[str, object]],
    mask: list[bool],
) -> object:
    """Mirror ``_DESIGN_COLUMNS_ROW_FORMULA`` for one spec row (column O)."""

    def width(item: SpecVariable) -> int:
        if item.var_type != "Categorical":
            return 1
        retained = _retained_levels(item, rows, mask)
        return 0 if retained is None else len(retained)

    if variable.role != _ROLE_PREDICTOR:
        return ""
    if not variable.include:
        return 0
    own = width(variable)

    term = str(variable.interaction_term or "")
    operation = str(variable.interaction_operation or "")
    if not term or not operation:
        return own
    operand = next(
        (item for item in spec if item.name.casefold() == term.casefold()),
        None,
    )
    # A name matching no column, or an operand whose Role is not Predictor,
    # contributes nothing — the constructor's mate() returning 0. An
    # operand that IS a Predictor but excluded still counts: that is the
    # flagged-amber marginality case, which builds columns.
    if operand is None or operand.role != _ROLE_PREDICTOR:
        return own
    return own + own * width(operand)


def _width_guard_status(total_columns: int, source_rows: int) -> str:
    """Mirror the M2 width-guard status cell (its leading token only).

    The full message embeds counts and prose; what a QC comparison needs is
    which of the three states the cell is in, so this returns ``"ERROR"``,
    ``"WARNING"`` or ``""`` and the sheet-side reader matches on the same
    leading token the CF rules key on.
    """
    if total_columns > _DESIGN_MATRIX_MAX_COLUMNS:
        return "ERROR"
    if (
        total_columns > _DESIGN_MATRIX_SOFT_COLUMNS
        or source_rows * total_columns > _DESIGN_MATRIX_SOFT_CELLS
    ):
        return "WARNING"
    return ""


def _response_name(spec: tuple[SpecVariable, ...]) -> str:
    """Mirror ``_RESPONSE_NAME_FORMULA``: the first Response row's header.

    ``"(none)"`` when no row carries the role (the XMATCH #N/A branch), and
    wrapped in ``Ln(...)`` when that row declares a Log transform — the
    audit strip never claims the model fits the raw column when it fits its
    log.
    """
    response = next((item for item in spec if item.role == _ROLE_RESPONSE), None)
    if response is None:
        return "(none)"
    return f"Ln({response.name})" if response.transform == "Log" else response.name


def _collect_flags(
    spec: tuple[SpecVariable, ...],
    rows: list[dict[str, object]],
    mask: list[bool],
    allow_intercept: bool,
) -> tuple[GuardFlag, ...]:
    """Recompute every spec-block CF predicate, in sheet-rule order.

    One entry per rule that fires, so a case's expected ``flags`` tuple is
    an exact statement of what the sheet should highlight — nothing more
    fires, and nothing less.
    """
    flags: list[GuardFlag] = []
    sequence_count = sum(1 for item in spec if item.sequence)
    fixed_effects_count = sum(
        1 for item in spec if item.role == _ROLE_FIXED_EFFECTS
    )
    categorical_included = any(
        item.role == _ROLE_PREDICTOR
        and item.include
        and item.var_type == "Categorical"
        for item in spec
    )
    headers = {item.name.casefold() for item in spec}

    # C2, the intercept toggle. Three rules, red first (StopIfTrue).
    if not allow_intercept and categorical_included:
        flags.append(
            GuardFlag("C", SEVERITY_RED, "intercept_off_with_categorical")
        )
    if allow_intercept and fixed_effects_count > 0:
        flags.append(GuardFlag("C", SEVERITY_RED, "intercept_with_fixed_effects"))

    for offset, item in enumerate(spec):
        row = SPEC_FIRST_DATA_ROW + offset

        # H: every flagged Sequence cell goes red once two-plus are marked.
        if item.sequence and sequence_count > 1:
            flags.append(
                GuardFlag("H", SEVERITY_RED, "multiple_sequence_flags", row)
            )

        # K: an included Categorical Predictor whose masked level set is
        # degenerate (L <= 1) contributes zero columns.
        if (
            item.role == _ROLE_PREDICTOR
            and item.include
            and item.var_type == "Categorical"
        ):
            levels = {
                rows[index][item.name]
                for index, included in enumerate(mask)
                if included and not _is_blank(rows[index][item.name])
            }
            if len(levels) <= 1:
                flags.append(
                    GuardFlag("K", SEVERITY_RED, "degenerate_categorical", row)
                )
            # E: a nonblank reference that makes Dummy_Levels fail.
            if item.reference != "" and _retained_levels(item, rows, mask) is None:
                flags.append(
                    GuardFlag("E", SEVERITY_RED, "invalid_reference", row)
                )

        # G: Log declared on a Categorical Predictor. Note the rule does NOT
        # test Include — a Categorical Predictor carrying Log is flagged
        # whether or not it is currently in the model.
        if (
            item.role == _ROLE_PREDICTOR
            and item.var_type == "Categorical"
            and item.transform == "Log"
        ):
            flags.append(
                GuardFlag("G", SEVERITY_RED, "log_on_categorical", row)
            )

        # M/N: the interaction pair. Red-M outranks amber-M (StopIfTrue).
        term = str(item.interaction_term or "")
        if term:
            operand = next(
                (o for o in spec if o.name.casefold() == term.casefold()), None
            )
            if term.casefold() not in headers or (
                operand is not None and operand.role != _ROLE_PREDICTOR
            ):
                flags.append(
                    GuardFlag("M", SEVERITY_RED, "operand_not_a_predictor", row)
                )
            elif operand is not None and not operand.include:
                flags.append(
                    GuardFlag("M", SEVERITY_AMBER, "operand_excluded", row)
                )

        # N: a reciprocal declaration under a SYMMETRIC operation (Product)
        # or an antisymmetric one (Difference) — a duplicate or exact-negative
        # column, so a singular Gram matrix. Ratio is excluded (asymmetric),
        # and a row naming itself is excluded (that is the documented
        # quadratic form).
        operation = str(item.interaction_operation or "")
        if operation.casefold() in ("product", "difference") and term:
            operand = next(
                (o for o in spec if o.name.casefold() == term.casefold()), None
            )
            if (
                operand is not None
                and operand.name.casefold() != item.name.casefold()
                and str(operand.interaction_term or "").casefold()
                == item.name.casefold()
                and str(operand.interaction_operation or "").casefold()
                == operation.casefold()
            ):
                flags.append(
                    GuardFlag("N", SEVERITY_RED, "reciprocal_declaration", row)
                )

    return tuple(flags)


def calculate_guard_state_case(case: GuardStateCase) -> GuardStateExpected:
    """Compute every expected spec-block value for one guard configuration."""
    loader = case.row_loader or load_source_rows
    csv_path = case.source_csv_path or load_source_rows.__defaults__[0]  # type: ignore[index]
    rows = loader(csv_path)
    if case.extra_columns:
        rows = [
            {**row, **{extra.name: extra.value_fn(row) for extra in case.extra_columns}}
            for row in rows
        ]

    spec = case.spec
    mask = _compute_mask(list(spec), rows)
    # The constructed names come from the shared model-construction
    # calculator, which is pure Python, sheet-agnostic, and — unlike
    # calculate_regression_spec_case — raises on nothing a guard spec can
    # throw at it (no Response, a degenerate block, a zero-column design
    # are all ordinary inputs to it). Reusing it is what keeps the guard
    # oracle's idea of a constructed column identical to the fittable
    # oracle's.
    expectations = calculate_model_construction_expectations(list(spec), rows)
    design_columns = tuple(
        _row_design_columns(item, spec, rows, mask) for item in spec
    )
    numeric_design_columns = [
        value for value in design_columns if isinstance(value, int)
    ]
    total = sum(numeric_design_columns) + (1 if case.allow_intercept else 0)

    sequence_count = sum(1 for item in spec if item.sequence)
    fixed_effects_count = sum(
        1 for item in spec if item.role == _ROLE_FIXED_EFFECTS
    )
    responses_count = sum(1 for item in spec if item.role == _ROLE_RESPONSE)

    deltas = _sequence_deltas(spec, rows)
    override = next(
        (
            case.sequence_period_override[item.name]
            for item in spec
            if item.sequence and item.name in case.sequence_period_override
        ),
        None,
    )
    period_in_use = (
        override if override is not None else base_period_delta_candidate(deltas)
    )

    # Audit k: COLUMNS(Predictor_Columns()) with an IFERROR fallback. The
    # constructor errors — and the strip degrades to the documented string —
    # exactly when the spec contributes no predictor columns at all.
    predictor_columns = len(expectations.constructed_column_names)
    audit_k: int | str = predictor_columns if predictor_columns else EMPTY_MODEL

    fixed_effects_name = next(
        (item.name for item in spec if item.role == _ROLE_FIXED_EFFECTS), None
    )

    return GuardStateExpected(
        case=case,
        responses_count=responses_count,
        sequence_flag_count=sequence_count,
        fixed_effects_count=fixed_effects_count,
        included_rows=sum(mask),
        response_name=_response_name(spec),
        model_formula=_build_model_formula(
            response_display=_response_name(spec),
            predictor_names=expectations.constructed_column_names,
            include_intercept=case.allow_intercept,
            fixed_effects_name=fixed_effects_name,
        ),
        audit_k=audit_k,
        design_columns=design_columns,
        design_columns_total=total,
        sequence_status=(
            _SEQUENCE_MULTI_FLAG_ERROR if sequence_count > 1 else ""
        ),
        fixed_effects_status=(
            _FIXED_EFFECTS_MULTI_ROW_ERROR if fixed_effects_count > 1 else ""
        ),
        width_guard_status=_width_guard_status(total, len(rows)),
        verdict=sequence_verdict(deltas, period_in_use),
        period_in_use=period_in_use,
        delta_spectrum=delta_spectrum(deltas),
        flags=_collect_flags(spec, rows, mask, case.allow_intercept),
    )


# ── The guard configurations ─────────────────────────────────────────────


def _mileage_spec(**overrides: SpecVariable) -> list[SpecVariable]:
    """The shipped Auto MPG T0 spec with named rows replaced."""
    by_name = {variable.name: variable for variable in overrides.values()}
    return [by_name.get(item.name, item) for item in build_default_spec()]


def _irregular_panel_spacing_spec() -> list[SpecVariable]:
    """P07's spec: the learning-curve model with Facility as the Identifier.

    P02's shipped spec cannot exercise the spacing layer at all, and the
    reason is worth stating because it is easy to get backwards.
    ``Sequence_Deltas`` groups by the **Identifier** columns, not by the
    Fixed Effects column — the Identifier is what the spacing layer treats
    as the panel unit. In P01/P02 the Identifier is ``Lot_ID``, which is
    unique per row, so every group has exactly one observation, there are
    no within-group consecutive pairs, and the verdict cell is
    unconditionally blank no matter how gapped the fiscal years are.

    Declaring ``Facility`` as the Identifier makes the three facilities the
    groups, so the deltas are within-facility fiscal-year gaps — which are
    genuinely irregular (the sites cover 1998-2023, 2001-2020 and 2005-2024
    with holes), and under a typed Δ of 1 that is exactly the Regularity
    verdict. ``Lot_ID`` drops to Omit; the fitted model is otherwise P01's.
    """
    return [
        _spec_var("Lot_ID", "Omit"),
        _spec_var("Facility", _ROLE_IDENTIFIER),
        _spec_var("Fiscal_Year", "Omit", sequence=True),
        _spec_var("Lot_Quantity", "Omit"),
        _spec_var("Cumulative_Units", "Omit"),
        _spec_var("Experience_Stock", "Omit"),
        _spec_var("Unit_Cost_BY", "Omit"),
        _spec_var("log Cum Units", _ROLE_PREDICTOR, True, "Continuous"),
        _spec_var("log experience", "Omit"),
        _spec_var("log Unit Cost", _ROLE_RESPONSE),
        _spec_var("Full_Data", "Filter"),
    ]


def build_guard_state_cases() -> list[GuardStateCase]:
    """Return every guard-rail configuration, in plan order.

    G13 (the hard width error, k > 16384 minus the design-matrix origin) is
    deliberately absent: § 1.4 documents it as conceptual only. Reaching it
    needs a categorical with ~16000 levels, which is not buildable at any
    reasonable size, and the soft warning it shares a formula with is
    covered for real by L07.
    """
    cases = [
        GuardStateCase(
            name="guard_no_response",
            plan_id="G01",
            sheet_name="G01 No Response Row",
            spec=tuple(_mileage_spec(mpg=_spec_var("MPG", "Omit"))),
            covers="Zero Response rows: the response readout degrades to "
            "'(none)' and the mask loses its response-completeness term, "
            "so the sample GROWS. Nothing leaks #N/A.",
        ),
        # The plan folds "(empty model)" into G01, but the two are separate
        # states with separate causes: "(none)" comes from a missing
        # RESPONSE, "(empty model)" from missing PREDICTORS. The audit
        # strip's IFERROR fires only for the second — Predictor_Columns
        # errors on DROP(built,,1) when nothing was built — so a case that
        # only drops the Response never reaches it. Splitting them keeps
        # each failure attributable to one cause.
        GuardStateCase(
            name="guard_empty_model",
            plan_id="G01b",
            sheet_name="G01b Empty Model",
            spec=tuple(
                _spec_var(item.name, item.role, False, item.var_type,
                          item.reference, item.sequence, item.transform)
                for item in build_default_spec()
            ),
            covers="No included Predictor rows at all: Predictor_Columns "
            "errors, and every consumer's IFERROR degrades to the "
            "documented '(empty model)' string rather than leaking #CALC!.",
        ),
        GuardStateCase(
            name="guard_two_responses",
            plan_id="G02",
            sheet_name="G02 Two Response Rows",
            spec=tuple(
                _mileage_spec(
                    weight=_spec_var("Weight", _ROLE_RESPONSE),
                )
            ),
            covers="Two Response rows: the audit strip's 'responses' count "
            "goes red; the constructor still resolves the FIRST one.",
        ),
        GuardStateCase(
            name="guard_two_sequence_flags",
            plan_id="G03",
            sheet_name="G03 Two Sequence Flags",
            # Both flags are declared explicitly. Auto MPG ships with NO
            # Sequence variable (it is cross-sectional — see
            # _DEFAULT_SEQUENCE_VARIABLES), so this case cannot inherit one
            # from the T0 spec and add a second; it has to state both. That
            # is the honest shape anyway: what is under test is the H2
            # cardinality rule, which is dataset-independent — it counts
            # flags, and does not care whether either column is a real
            # ordering axis.
            spec=tuple(
                _mileage_spec(
                    model_year=_spec_var(
                        "Model Year", _ROLE_PREDICTOR, True, "Categorical",
                        sequence=True,
                    ),
                    acceleration=_spec_var(
                        "Acceleration", _ROLE_PREDICTOR, True, "Continuous",
                        sequence=True,
                    ),
                )
            ),
            covers="Two Sequence flags: the E1/H2 status line errors and "
            "every flagged H cell goes red.",
        ),
        GuardStateCase(
            name="guard_two_fixed_effects",
            plan_id="G04",
            sheet_name="G04 Two Fixed Effects Rows",
            spec=tuple(
                _mileage_spec(
                    origin=_spec_var("Origin", _ROLE_FIXED_EFFECTS),
                    make=_spec_var("Make", _ROLE_FIXED_EFFECTS),
                )
            ),
            covers="Two Fixed Effects rows: the B1 cardinality error fires; "
            "Fixed_Effects_Column still resolves the first.",
        ),
        GuardStateCase(
            name="guard_fixed_effects_with_intercept",
            plan_id="G05",
            sheet_name="G05 FE Plus Intercept",
            spec=tuple(
                _mileage_spec(
                    origin=_spec_var("Origin", _ROLE_FIXED_EFFECTS),
                )
            ),
            allow_intercept=True,
            covers="Fixed Effects with the intercept left ON: red on the "
            "toggle, because the estimated intercept is a panel artifact.",
        ),
        GuardStateCase(
            name="guard_intercept_off_with_categorical",
            plan_id="G06",
            sheet_name="G06 Intercept Off With Cat",
            spec=tuple(build_default_spec()),
            allow_intercept=False,
            covers="Intercept OFF with an included Categorical: red on the "
            "toggle. Advisory only — M02 fits this same spec.",
        ),
        GuardStateCase(
            name="guard_log_on_categorical",
            plan_id="G07",
            sheet_name="G07 Log On Categorical",
            spec=tuple(
                _mileage_spec(
                    origin=_spec_var(
                        "Origin", _ROLE_PREDICTOR, True, "Categorical",
                        transform="Log",
                    ),
                )
            ),
            covers="Transform=Log on a Categorical Predictor: red G cell. "
            "The dummies still encode normally — flagged, never applied.",
        ),
        GuardStateCase(
            name="guard_interaction_bad_operand",
            plan_id="G09",
            sheet_name="G09 Interaction Bad Operand",
            spec=tuple(
                _mileage_spec(
                    weight=_spec_var(
                        "Weight", _ROLE_PREDICTOR, True, "Continuous",
                        interaction_term="Car Name",
                        interaction_operation="Product",
                    ),
                )
            ),
            covers="Interaction Term naming a non-Predictor row (Car Name is "
            "the Identifier): red M, and the row contributes its main "
            "effect only.",
        ),
        GuardStateCase(
            name="guard_reciprocal_product",
            plan_id="G10",
            sheet_name="G10 Symmetric Product Pair",
            spec=tuple(
                _mileage_spec(
                    horsepower=_spec_var(
                        "Horsepower", _ROLE_PREDICTOR, True, "Continuous",
                        interaction_term="Weight",
                        interaction_operation="Product",
                    ),
                    weight=_spec_var(
                        "Weight", _ROLE_PREDICTOR, True, "Continuous",
                        interaction_term="Horsepower",
                        interaction_operation="Product",
                    ),
                )
            ),
            covers="Reciprocal Product declaration (A x B and B x A): red N "
            "on both rows — a duplicate column and a singular Gram matrix. "
            "The legal counterpart is M11, where Ratio makes the pair "
            "genuinely distinct.",
        ),
        GuardStateCase(
            name="guard_excluded_operand",
            plan_id="G11",
            sheet_name="G11 Excluded Operand",
            spec=tuple(
                _mileage_spec(
                    weight=_spec_var(
                        "Weight", _ROLE_PREDICTOR, True, "Continuous",
                        interaction_term="Acceleration",
                        interaction_operation="Product",
                    ),
                    acceleration=_spec_var(
                        "Acceleration", _ROLE_PREDICTOR, False, "Continuous"
                    ),
                )
            ),
            covers="Interaction operand with Include=FALSE: AMBER, not red. "
            "A marginality violation is a modelling choice, so it is "
            "flagged and allowed — the columns still build.",
        ),
        GuardStateCase(
            name="guard_unknown_interaction_operation",
            plan_id="G12",
            sheet_name="G12 Unknown Interaction Op",
            spec=tuple(
                _mileage_spec(
                    weight=_spec_var(
                        "Weight", _ROLE_PREDICTOR, True, "Continuous",
                        interaction_term="Displacement",
                        interaction_operation="Quotient",
                    ),
                    displacement=_spec_var(
                        "Displacement", _ROLE_PREDICTOR, True, "Continuous"
                    ),
                )
            ),
            covers="An Interaction Operation outside the dropdown vocabulary "
            "(pasted past the validation): the header renders ' ? ' and the "
            "design column is NA() — refused, never guessed at.",
        ),
        GuardStateCase(
            name="guard_ln_zero_propagation",
            plan_id="L06",
            sheet_name="L06 Ln Zero Guard",
            spec=tuple(
                _life_spec(
                    response=_spec_var("Life expectancy", _ROLE_RESPONSE),
                    schooling=_spec_var(
                        "Schooling", _ROLE_PREDICTOR, True, "Continuous",
                        transform="Log",
                    ),
                    adult_mortality=_spec_var(
                        "Adult Mortality", _ROLE_PREDICTOR, True, "Continuous"
                    ),
                )
            ),
            source_csv_path=LIFE_EXPECTANCY_CSV_PATH,
            row_loader=load_life_expectancy_source_rows,
            source_table_ref="=LifeExpectancyData[#All]",
            covers="Ln of a column containing true zeros (Schooling has 28). "
            "The zero rows stay in the mask and Ln_Positive returns #N/A, "
            "so the #N/A propagates — the sample does NOT narrow. See "
            "_LN_ZERO_GUARD_NOTE.",
        ),
        GuardStateCase(
            name="guard_width_guard_warning",
            plan_id="L07",
            sheet_name="L07 Width Guard Warning",
            spec=tuple(_life_country_width_guard_spec()),
            source_csv_path=LIFE_EXPECTANCY_CSV_PATH,
            row_loader=load_life_expectancy_source_rows,
            source_table_ref="=LifeExpectancyData[#All]",
            covers="k = 205 design columns: the M2 width guard reads "
            "WARNING, and the engine returns #N/A rather than a plausible "
            "wrong number — the workbook cannot invert a Gram matrix this "
            "wide, which is what the guard exists to say.",
        ),
        GuardStateCase(
            name="guard_sequence_period_override",
            plan_id="M16",
            sheet_name="M16 Sequence Period Override",
            # Model Year is Sequence-flagged HERE and nowhere else on Auto
            # MPG. The dataset is cross-sectional, so the flag is not a
            # claim about the data — it is the minimum wiring needed to
            # reach the typed-override path, which is what this case tests.
            # The override machinery reads Spec_Sequence / Spec_Sequence_
            # Period positionally and is indifferent to whether the flagged
            # column means anything, so Auto MPG's evenly spaced integer
            # years are a clean substrate for it: candidate Δ = 1, typed
            # Δ = 2, and the verdict has to follow the typed one.
            spec=tuple(
                _mileage_spec(
                    model_year=_spec_var(
                        "Model Year", _ROLE_PREDICTOR, True, "Categorical",
                        sequence=True,
                    ),
                )
            ),
            sequence_period_override={"Model Year": 2.0},
            covers="A typed Sequence Period of 2 against a candidate Δ of 1: "
            "the Period In Use cell shows the override, and the verdict is "
            "re-evaluated against it rather than against the candidate.",
        ),
        GuardStateCase(
            name="guard_irregular_panel_spacing",
            plan_id="P07",
            sheet_name="P07 Irregular Panel Spacing",
            spec=tuple(_irregular_panel_spacing_spec()),
            source_csv_path=PRODUCTION_LOTS_CSV_PATH,
            row_loader=load_production_lots_source_rows,
            source_table_ref="=ProductionLotsData[#All]",
            sequence_period_override={"Fiscal_Year": 1.0},
            covers="A genuinely gapped panel (three facilities with "
            "different, non-contiguous fiscal-year ranges) under a typed "
            "Δ of 1: the Regularity verdict fires yellow.",
        ),
    ]
    for case in cases:
        validate_sheet_name(case.sheet_name)
    assert_sheet_names_unique([case.sheet_name for case in cases])
    return cases


def build_guard_state_expectations() -> list[GuardStateExpected]:
    """Compute expected spec-block values for every guard configuration."""
    return [calculate_guard_state_case(case) for case in build_guard_state_cases()]
