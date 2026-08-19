"""The spec-driven verify path must type the Sequence Period, not just the flag.

``tools/inspect_regression_sheet.read_regression_df`` is the loop behind
``build_production.py --verify``: for each ``RegressionSpecCase`` it writes the
spec onto the Regression sheet, recalculates, and reads cells back against the
NumPy oracle. It is Excel-only, so it has no headless coverage — which is how a
gap hid for every verify run on record.

The gap: ``apply_spec_case`` writes the Sequence *flag* (column H) but not the
typed Sequence *Period* (column I). ``Base_Period_Delta()`` reads the TYPED
value and returns ``#N/A`` when column I is blank, and the BFN panel
Durbin-Watson cell (AE12) passes that as its delta. With ``step`` at ``#N/A``
every ``Difference_By`` lookup misses, the ``n_terms`` guard fires ``#N/A``, and
the cell reads ``nan`` — while the oracle holds ``case.sequence_period`` (1.0 on
P01/P02) and computes a real BFN (0.985). That mismatch reads as a broken
statistic and is really a blank input cell the harness never typed.

``write_sheet_test_model.apply_case_to_sheet`` types column I via
``apply_sequence_period_overrides`` (pinned headlessly by
``test_a_cases_typed_sequence_period_lands_in_spec_column_i``); the inspector
must do the same. This test pins that the inspector does, because the only
symptom of a regression here is a multi-minute Excel run that reports a
plausible-looking broken diagnostic.

P01/P02 are the only ``RegressionSpecCase``s with ``sequence_period`` set
(every other case leaves it ``None``), so the guard
``if expected.case.sequence_period is not None`` is the right shape: it touches
exactly the cases that make AE12 live and nothing else.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
TOOL_PATH = ROOT_DIR / "tools" / "inspect_regression_sheet.py"


def _load_inspect_module():
    """Load ``tools/inspect_regression_sheet.py`` without executing its ``main``.

    Importing the module binds its names (``apply_sequence_period_overrides``
    among them) without opening Excel — that only happens inside
    ``read_regression_df``/``main``, which this never calls.
    """
    spec = importlib.util.spec_from_file_location("inspect_regression_sheet_under_test", TOOL_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_inspector_imports_the_sequence_period_override() -> None:
    """The override must be importable from the inspector module.

    Removing the import is the first step of reintroducing the bug; binding the
    name at module scope is what the call site relies on.
    """
    module = _load_inspect_module()
    from lambda_catalog.regression_spec_sheet_io import apply_sequence_period_overrides

    assert module.apply_sequence_period_overrides is apply_sequence_period_overrides


def test_inspector_types_column_i_for_sequence_period_cases() -> None:
    """The per-config loop must call the override, guarded by ``sequence_period``.

    A name resolving to the wrong range does not error — it returns numbers from
    the wrong rows — so the BFN failure mode here is silence in Excel. The guard
    (``if expected.case.sequence_period is not None``) is load-bearing: without
    it, a non-panel config would get an empty-override call (harmless) but the
    real risk is the call going missing entirely, which is what this pins.
    """
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "apply_sequence_period_overrides(" in source, (
        "the verify loop must type the Sequence Period (column I) or the BFN "
        "panel Durbin-Watson cell reads nan while the oracle holds a real number"
    )
    assert "sequence_period is not None" in source, (
        "the override must be gated on the case declaring a period — only P01/P02 "
        "do, and typing column I on a non-panel case is a no-op at best"
    )
    # The call sits inside the per-config loop, right after apply_spec_case and
    # before set_prediction_inputs — the same order the builder uses. Pin the
    # neighbourhood so a future reordering cannot drop it back out.
    apply_index = source.index("apply_sequence_period_overrides(")
    assert source.index("apply_spec_case(") < apply_index
    assert source.index("set_prediction_inputs(") > apply_index