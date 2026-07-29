"""Tests for the best-effort Excel-window helpers in workbook_helpers."""
# pylint: disable=missing-function-docstring
from types import SimpleNamespace

from lambda_catalog.workbook_helpers import safe_activate, safe_freeze_top_row


class _RaisingSheet:
    def activate(self) -> None:
        raise RuntimeError("Could not activate App! Try to instantiate the App with visible=True.")


def test_safe_activate_swallows_activation_failure() -> None:
    safe_activate(_RaisingSheet())  # must not raise


def test_safe_activate_calls_through_on_success() -> None:
    calls: list[str] = []
    sheet = SimpleNamespace(activate=lambda: calls.append("activated"))
    safe_activate(sheet)
    assert calls == ["activated"]


def test_safe_freeze_top_row_swallows_missing_active_window() -> None:
    class _NoActiveWindowApplication:
        @property
        def ActiveWindow(self) -> None:
            raise RuntimeError("ActiveWindow is not available in this session")

    sheet = SimpleNamespace(api=SimpleNamespace(Application=_NoActiveWindowApplication()))

    safe_freeze_top_row(sheet)  # must not raise


def test_safe_freeze_top_row_sets_split_and_freeze_on_success() -> None:
    window = SimpleNamespace(SplitRow=None, SplitColumn=None, FreezePanes=None)
    sheet = SimpleNamespace(
        api=SimpleNamespace(Application=SimpleNamespace(ActiveWindow=window))
    )

    safe_freeze_top_row(sheet)

    assert window.SplitRow == 1
    assert window.SplitColumn == 0
    assert window.FreezePanes is True
