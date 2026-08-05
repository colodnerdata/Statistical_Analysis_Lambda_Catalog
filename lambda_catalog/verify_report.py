"""Structured pass/fail report for the workbook verifier.

Both verifier layers (headless zipfile check, deep spec-driven check) emit
the same ``VerifyReport`` so a downstream agentic loop can read the result
in a stable shape regardless of which layer ran.

The human-readable renderer mirrors the existing ``lambda_catalog.deep_verify._report_qc_failure``
style (the ``ERROR Verify mismatch totals: ...`` line)
so the post-build handoff looks the same whether the verifier was run by
``build_production --verify`` or by the legacy ``build_qc`` workflow.
"""

from __future__ import annotations

import json
from collections import Counter
from dataclasses import asdict, dataclass, field
from typing import Literal

VerifyMode = Literal["headless", "spec"]


@dataclass(frozen=True)
class VerifyReport:
    """Outcome of a single verifier run.

    Attributes
    ----------
    passed : bool
        True when no failures were detected.
    categories : dict[str, int]
        Per-category failure counts (e.g. ``{"Regression/scalars": 5}``).
        Empty when ``passed`` is True.
    failures : tuple[str, ...]
        Per-failure message strings. Tuple (not list) so the dataclass is
        hashable and downstream code can safely share the report across
        processes.
    elapsed_seconds : float
        Wall-clock time for the verifier phase only (excludes the build
        itself). Useful for budget tracking in CI.
    mode : str
        ``"headless"`` for the zipfile/lxml invariant check (Layer 1) or
        ``"spec"`` for the spec-driven xlwings check (Layer 2).
    workbook : str
        Path of the workbook that was verified. String (not Path) for clean
        JSON serialization.
    """

    passed: bool
    categories: dict[str, int] = field(default_factory=dict)
    failures: tuple[str, ...] = ()
    elapsed_seconds: float = 0.0
    mode: VerifyMode = "headless"
    workbook: str = ""


def report_from_failures(
    failures: list[str],
    *,
    elapsed_seconds: float,
    mode: VerifyMode,
    workbook: str,
) -> VerifyReport:
    """Build a VerifyReport from a list of failure messages.

    Empty list -> passed=True with empty categories. Otherwise categories
    are derived by splitting each message on ``]`` (the existing
    ``lambda_catalog.deep_verify._report_qc_failure`` prefix convention) and tallying the
    leading bracketed category name.
    """
    if not failures:
        return VerifyReport(
            passed=True,
            categories={},
            failures=(),
            elapsed_seconds=elapsed_seconds,
            mode=mode,
            workbook=workbook,
        )

    category_counts = Counter(
        message.split("]", 1)[0].removeprefix("[") for message in failures
    )
    return VerifyReport(
        passed=False,
        categories=dict(sorted(category_counts.items())),
        failures=tuple(failures),
        elapsed_seconds=elapsed_seconds,
        mode=mode,
        workbook=workbook,
    )


def render_human(report: VerifyReport, *, max_failures: int = 20) -> str:
    """Format a VerifyReport for terminal output.

    Mirrors the existing ``ERROR Verify mismatch totals: ...`` format so the
    post-build handoff looks the same regardless of which verifier ran.
    """
    if report.passed:
        return (
            f"Verify: passed ({report.mode} mode, {report.elapsed_seconds:.1f}s, "
            f"workbook={report.workbook!r})"
        )
    category_summary = ", ".join(
        f"{category}={count}" for category, count in sorted(report.categories.items())
    )
    shown = report.failures[:max_failures]
    extra = len(report.failures) - len(shown)
    lines = [
        f"ERROR Verify mismatch totals: {category_summary}",
        f"  mode: {report.mode}",
        f"  workbook: {report.workbook}",
        f"  elapsed: {report.elapsed_seconds:.1f}s",
        f"  failures ({len(report.failures)} total):",
    ]
    for failure in shown:
        lines.append(f"    - {failure}")
    if extra > 0:
        lines.append(f"    ... and {extra} more (use --json for the full list)")
    return "\n".join(lines)


def render_json(report: VerifyReport) -> str:
    """Format a VerifyReport as JSON for agentic consumption.

    The output schema is stable: ``passed``, ``categories``, ``failures``,
    ``elapsed_seconds``, ``mode``, ``workbook``. Downstream code can parse
    the result without depending on the human-readable format.
    """
    return json.dumps(asdict(report), indent=2, sort_keys=True)
