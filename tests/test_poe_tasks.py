"""Pins the poethepoet task table in pyproject.toml.

The tasks under ``[tool.poe.tasks]`` are the repo's command surface: CI invokes
``test-cov`` and ``lint`` by name, and CONTRIBUTING.md documents the rest. A
silent rename or deletion would break the workflow file or the docs without any
test noticing, so the task names are pinned here the same way
``_EXPECTED_CASE_NAMES`` pins the regression spec cases — adding, renaming, or
dropping a task means editing this list in the same commit.

Pure TOML reads; no Excel, no poethepoet import, no subprocess.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest

# tomllib is stdlib from 3.11; this project supports 3.10, where the tomli
# backport (dev group, same API) stands in for it.
if sys.version_info >= (3, 11):
    import tomllib
else:
    import tomli as tomllib

_PYPROJECT = Path(__file__).resolve().parent.parent / "pyproject.toml"

# Ordered exactly as they appear in pyproject.toml.
_EXPECTED_TASK_NAMES = [
    "test",
    "test-cov",
    "test-excel",
    "lint",
    "check",
    "verify-headless",
    "verify-deep",
    "verify-test-models",
    "verify-guards",
    "verify-spec-errors",
    "verify-models",
    # The resume slice: verify-models minus the fittable cases verify-spec-errors
    # already built. See test_the_two_resumable_slices_partition_the_registry.
    "verify-models-rest",
    "verify",
    "build",
    # Added with the presentation demo workbook; the pin was not updated in
    # the same commit, so this list was one merge away from failing on main.
    "build-presentation-demo",
    "build-presentation",
    "static-sheets",
    "resync-names",
    # Docs site (PR: sphinx-tutorial-site): generate the derived pages from
    # the authored sheet lists, then run sphinx-build -W.
    "docs-generate",
    "docs",
]

# The tasks .github/workflows/ci.yml runs by name. Renaming either without
# editing the workflow turns every CI run red at the "poe: unrecognised task"
# stage, which is a slower way to learn it than this assertion.
_CI_TASK_NAMES = ["test-cov", "lint"]


@pytest.fixture(scope="module")
def poe_config() -> dict[str, Any]:
    with _PYPROJECT.open("rb") as handle:
        return tomllib.load(handle)["tool"]["poe"]


@pytest.fixture(scope="module")
def tasks(poe_config: dict[str, Any]) -> dict[str, Any]:
    return poe_config["tasks"]


def test_task_names_match_the_pinned_list(tasks: dict[str, Any]) -> None:
    assert list(tasks) == _EXPECTED_TASK_NAMES


def test_ci_invoked_tasks_exist(tasks: dict[str, Any]) -> None:
    for name in _CI_TASK_NAMES:
        assert name in tasks


def test_every_task_has_help_text(tasks: dict[str, Any]) -> None:
    """Bare ``poe`` lists tasks by their help string; a task without one is
    invisible in exactly the way the old Makefile's comments were."""
    missing = [name for name, task in tasks.items() if not task.get("help")]
    assert missing == []


def test_no_task_re_enters_uv(tasks: dict[str, Any]) -> None:
    """poe's uv executor already runs each task inside the project environment.
    A task that spelled ``uv run`` itself would nest one uv invocation inside
    another and silently reintroduce the prefix the migration removed."""
    for name, task in tasks.items():
        command = task.get("cmd")
        if command is not None:
            assert not command.startswith("uv "), name


def test_the_excel_task_targets_the_suite_that_reads_the_env_var(
    tasks: dict[str, Any],
) -> None:
    """``test-excel`` must run the tests ``RUN_EXCEL_INTEGRATION`` actually gates.

    The task sets the env var, so it is only worth running against a suite
    that reads it. The committed-artifact invariants are always-on zipfile
    reads (no Excel), so pointing the task at them would set a variable
    nothing downstream consulted — a slower spelling of ``verify-headless``.
    Pinned to the pair so a future edit to either side fails here instead of
    quietly reintroducing that.
    """
    task = tasks["test-excel"]
    assert task["env"]["RUN_EXCEL_INTEGRATION"] == "1"

    target = Path(__file__).resolve().parent.parent / task["cmd"].split()[1]
    assert target.exists(), target
    assert "RUN_EXCEL_INTEGRATION" in target.read_text(encoding="utf-8"), (
        f"{target.name} does not read RUN_EXCEL_INTEGRATION, so test-excel "
        "setting it buys nothing"
    )


def test_lockfile_strictness_is_configured(poe_config: dict[str, Any]) -> None:
    """The Makefile passed ``--frozen`` on every recipe. UV_FROZEN carries that
    intent for every task at once; dropping it would let a stale lockfile
    resolve silently instead of failing."""
    assert poe_config["env"]["UV_FROZEN"] == "1"


def test_verify_builds_the_artifacts_before_screening_them(
    tasks: dict[str, Any],
) -> None:
    """The screen has to run LAST.

    ``verify-headless`` reads whatever is sitting in ``dist/``; the three deep
    tasks rewrite those files, so running the screen first would validate the
    previously committed artifacts and never look at the ones the run just
    built — a rebuild that broke a defined name or orphaned a chart
    relationship would pass ``verify`` clean.
    """
    steps = tasks["verify"]["sequence"]

    assert steps[-1] == "verify-headless"
    assert steps[0]["parallel"] == [
        "verify-deep",
        "verify-test-models",
    ]


def test_parallel_excel_output_is_buffered_not_interleaved(
    tasks: dict[str, Any],
) -> None:
    """Streaming three concurrent drivers line-by-line would shred exactly the
    transcripts ``excel-only-runs/`` exists to preserve."""
    assert tasks["verify"]["sequence"][0]["output_mode"] == "buffer"


def test_parallel_tasks_need_a_poe_that_has_them(poe_config: dict[str, Any]) -> None:
    """``parallel`` is not in every poethepoet. An older one reads the key as an
    unknown task type and refuses to run ``verify`` at all, so the floor in the
    dev dependency group is load-bearing, not cosmetic."""
    del poe_config  # the pin lives in [dependency-groups], read separately
    with _PYPROJECT.open("rb") as handle:
        dev_group = tomllib.load(handle)["dependency-groups"]["dev"]

    poe_pin = next(spec for spec in dev_group if spec.startswith("poethepoet"))
    floor = poe_pin.split(">=")[1].split(",")[0]
    assert tuple(int(part) for part in floor.split(".")) >= (0, 48)


# ── The Layer 2 slice tasks ──────────────────────────────────────────────────


def _case_selection_tokens(command: str, flag: str) -> list[str]:
    """Every comma-separated token a task passes to ``--cases`` / ``--exclude``.

    The commands are written across several lines in pyproject.toml, so the
    value is found by scanning the whitespace-split argv rather than by
    slicing the raw string.
    """
    parts = command.split()
    if flag not in parts:
        return []
    return [
        token.strip()
        for token in parts[parts.index(flag) + 1].split(",")
        if token.strip()
    ]


def test_task_case_selections_all_resolve_to_registered_cases(
    tasks: dict[str, Any],
) -> None:
    """A plan ID named by a task must exist.

    ``verify-spec-errors`` names its cases explicitly, because "which cases are
    an error surface" is a judgement rather than something derivable from the
    registry. That makes it exactly the kind of hand-maintained list that goes
    stale silently — the case gets renamed or dropped and the task quietly
    stops covering it, or fails at build time on a machine that has Excel,
    which is the slowest possible place to learn it. This resolves every token
    against the live registry instead.

    It cannot catch a NEW error surface that nobody added to the list; adding
    one is a deliberate act that the PR-shape rules already route through
    docs/MODEL_TESTING_ASSETS.md.
    """
    from lambda_catalog.analyze_regression_guard_states import build_guard_state_cases
    from lambda_catalog.analyze_regression_spec import build_regression_spec_cases

    known = {
        token.casefold()
        for case in (*build_regression_spec_cases(), *build_guard_state_cases())
        for token in (case.plan_id, case.name)
    }

    for name, task in tasks.items():
        command = task.get("cmd")
        if command is None:
            continue
        for flag in ("--cases", "--exclude"):
            for token in _case_selection_tokens(command, flag):
                assert token.casefold() in known, f"{name} {flag} {token}"


def test_the_two_resumable_slices_partition_the_registry(
    tasks: dict[str, Any],
) -> None:
    """``verify-spec-errors`` + ``verify-models-rest`` = every case, exactly once.

    The pair exists so a developer who has run one slice can finish the job
    without rebuilding sheets. That only holds if the second slice excludes
    precisely the fittable cases the first already built — G08, M15 and L12,
    which are registered as ``RegressionSpecCase`` because they produce a real
    fit and an error surface at the same time.

    Derived here rather than restated, because a hand-copied list is what went
    wrong last time: ``verify-guards``' comment claimed L07 was its only
    overlap with ``verify-spec-errors`` when in truth the whole task was a
    subset, and the cost was a redundant multi-minute Excel run. A comment
    cannot be executed; this can.
    """
    from lambda_catalog.analyze_regression_guard_states import build_guard_state_cases
    from lambda_catalog.analyze_regression_spec import build_regression_spec_cases

    spec_errors = {
        token.casefold()
        for token in _case_selection_tokens(tasks["verify-spec-errors"]["cmd"], "--cases")
    }
    excluded = {
        token.casefold()
        for token in _case_selection_tokens(tasks["verify-models-rest"]["cmd"], "--exclude")
    }

    models = {case.plan_id.casefold() for case in build_regression_spec_cases()}
    guards = {case.plan_id.casefold() for case in build_guard_state_cases()}

    # The exclusion is exactly the overlap, so the two slices are disjoint...
    assert excluded == spec_errors & models

    # ...and between them they reach every registered case. Heavy cases are the
    # documented exception: --kind models drops them unless --include-heavy.
    heavy = {
        case.plan_id.casefold()
        for case in build_regression_spec_cases()
        if getattr(case, "heavy", False)
    }
    assert (spec_errors | (models - excluded)) >= (models | guards) - heavy

    # And the claim that made this necessary: verify-guards is a strict subset
    # of verify-spec-errors, so running both after the other is pure waste.
    assert guards <= spec_errors


def test_spec_error_task_covers_every_row2_status_surface(
    tasks: dict[str, Any],
) -> None:
    """The five status cells the spec block's row-2 grammar defines.

    One case per surface, so a task that silently lost one would still look
    plausible. The width guard (O2) is the reason L07 is in this list despite
    being the case verify-guards excludes for size: it is the only case that
    reaches that cell at all.
    """
    selected = {
        token.casefold()
        for token in _case_selection_tokens(
            tasks["verify-spec-errors"]["cmd"], "--cases"
        )
    }

    for surface, plan_id in (
        ("B2 role cardinality", "G01"),
        ("G2 log domain", "L06"),
        ("H2 sequence cardinality", "G03"),
        ("I2 spacing verdict", "P07"),
        ("O2 width guard", "L07"),
    ):
        assert plan_id.casefold() in selected, surface


def test_structural_slices_do_not_hardcode_plan_ids(tasks: dict[str, Any]) -> None:
    """``verify-guards`` and ``verify-models`` select with ``--kind``.

    Enumerating either half would put a list of ~17 or ~33 plan IDs in
    pyproject.toml that nothing keeps current — the registry already knows the
    split, so the task should ask it. The single ``--exclude`` on verify-guards
    is the documented hole, not an enumeration.
    """
    guards = tasks["verify-guards"]["cmd"].split()
    models = tasks["verify-models"]["cmd"].split()

    assert "--kind" in guards and guards[guards.index("--kind") + 1] == "guards"
    assert "--kind" in models and models[models.index("--kind") + 1] == "models"
    assert "--cases" not in guards
    assert "--cases" not in models
    # The happy-path slice must not pull in the two heavy Gram matrices.
    assert "--include-heavy" not in models
    assert _case_selection_tokens(tasks["verify-guards"]["cmd"], "--exclude") == ["L07"]
