"""Build Lambda_Library_TestModels.xlsx — one worksheet per test-model case.

A third build target, alongside ``build_production.py`` (the Regression
artifact) and ``build_univariate.py`` (the Univariate artifact). Unlike
those two this one is **not shipped**: it is a QC fixture, gitignored like
``Lambda_Library_QC.xlsx``, and its only audience is whoever is changing the
spec block or the engine.

What it produces is the regression test-model suite made physical. Every
``RegressionSpecCase`` in ``docs/MODEL_TESTING_ASSETS.md`` § 1.1–1.3 and
every ``GuardStateCase`` in § 1.4 gets its own Regression-shaped sheet, named
for the corner it covers (``M09 Cat x Cat Full Product``, ``G10 Symmetric
Product Pair``), pre-populated and fully calculated. Verification then reads
those sheets and writes nothing at all, which is the whole point: the legacy
harness has to mutate one shared sheet ~33 times, so a failure leaves nothing
behind to look at and no case can be inspected independently of the ones
before it.

Usage::

    python scripts/build_test_models.py                       # every non-heavy case
    python scripts/build_test_models.py --include-heavy        # + L07 / L08
    python scripts/build_test_models.py --cases M09,G10        # just those two
    python scripts/build_test_models.py --verify --no-launch   # build, check, exit 1 on drift
"""
from __future__ import annotations

import argparse
import contextlib
import subprocess
import sys
import time
from collections.abc import Iterator
from pathlib import Path

import xlwings as xw

from lambda_catalog.analyze_regression_guard_states import (
    build_guard_state_cases,
    calculate_guard_state_case,
)
from lambda_catalog.analyze_regression_spec import (
    build_regression_spec_cases,
    calculate_regression_spec_case,
)
from lambda_catalog.build_common import (
    RUN_LOG_DIR_NAME,
    RUN_LOG_FILE_SUFFIX,
    _retry_on_open,
    print_name_sync_summary,
    run_log_path,
    tee_run_log,
    warn_if_workbook_open,
)
from lambda_catalog.catalog_schema import load_catalog_document
from lambda_catalog.workbook_builder import (
    XL_CALCULATION_MANUAL,
    XL_CALCULATION_SEMIAUTOMATIC,
    NameSyncResult,
    sync_workbook_names,
)
from lambda_catalog.workbook_helpers import (
    OPEN_WORKBOOK_ERRORS,
    raise_excel_access_error,
)
from lambda_catalog.write_sheet_csv_dataset import (
    LIFE_EXPECTANCY,
    MILEAGE,
    PRODUCTION_LOTS,
    load_csv_rows,
    write_csv_dataset_sheet,
)
from lambda_catalog.write_sheet_lambda_functions import write_catalog_sheet
from lambda_catalog.write_sheet_test_model import (
    FIXTURE_COLUMNS,
    write_guard_state_sheet,
    write_test_model_sheet,
)

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_WORKBOOK_PATH = ROOT_DIR / "Lambda_Library_TestModels.xlsx"
DEFAULT_DEFINITIONS_PATH = ROOT_DIR / "lambda_functions.json"

# Which shipped dataset each SPEC_DATASET_PROFILES key writes to. The fixture
# columns themselves are declared once in write_sheet_test_model.FIXTURE_COLUMNS,
# because the SPEC BLOCK has to know about them too: a column added here widens
# Source_Data, and a spec block one row short of the data makes every
# constructor's INDEX(band, COLUMNS(Source_Data)) run off the end. Two copies
# of that list is exactly how this workbook produced 74,065 mismatches on its
# first successful build.
_CONFIG_BY_PROFILE_KEY = {
    "auto_mpg": MILEAGE,
    "life_expectancy": LIFE_EXPECTANCY,
    "production_lots": PRODUCTION_LOTS,
}


class _Progress:
    """Verbose per-sheet progress for a build that runs for many minutes.

    A full run writes ~46 Regression-shaped sheets through COM, and the two
    questions a watcher actually has are "is it still moving?" and "which
    sheet is it stuck on?". Both need the CURRENT sheet named BEFORE the work
    starts and the line flushed immediately — a summary printed after each
    sheet tells you nothing about the one that hung.

    Every line carries elapsed-since-start as well as the per-sheet time, so
    an archived log reads as a timeline rather than a list.
    """

    def __init__(self, *, enabled: bool, run_start: float) -> None:
        self._enabled = enabled
        self._run_start = run_start

    def _stamp(self) -> str:
        return f"[{time.monotonic() - self._run_start:7.1f}s]"

    def phase(self, label: str) -> None:
        """Announce a build phase. Always printed, verbose or not — these are
        the checkpoints that make a hang attributable without a rerun."""
        print(f"{self._stamp()} {label}", flush=True)

    def detail(self, line: str) -> None:
        """A verbose-only line, indented under the phase it belongs to."""
        if self._enabled:
            print(f"{' ' * 10} {line}", flush=True)

    @contextlib.contextmanager
    def sheet(self, index: int, total: int, case) -> Iterator[None]:
        """Bracket one sheet's write, naming it before the work begins."""
        if not self._enabled:
            yield
            return
        label = f"{index:>3}/{total} {case.plan_id:<5} {case.sheet_name}"
        # No newline yet: the per-sheet duration is appended when it lands,
        # so an interrupted run leaves the hung sheet's name on screen with
        # no time after it — which is exactly the diagnostic.
        print(f"{self._stamp()} {label:<44}", end="", flush=True)
        started = time.monotonic()
        try:
            yield
        except BaseException:
            print(" FAILED", flush=True)
            raise
        else:
            print(f" {time.monotonic() - started:6.1f}s", flush=True)


def _selected_cases(
    case_filter: set[str] | None, include_heavy: bool
) -> tuple[list, list]:
    """Return (fittable cases, guard cases) after --cases / --include-heavy.

    ``case_filter`` matches a plan ID (``"M09"``) or a case name
    (``"interaction_categorical_cross"``), case-insensitively, so either
    identifier a developer has to hand works. An explicit ``--cases``
    selection overrides ``heavy``: naming L07 means you want L07.
    """
    model_cases = build_regression_spec_cases()
    guard_cases = build_guard_state_cases()

    if case_filter is not None:
        wanted = {token.casefold() for token in case_filter}

        def _matches(case) -> bool:
            return (
                case.plan_id.casefold() in wanted
                or case.name.casefold() in wanted
            )

        model_cases = [case for case in model_cases if _matches(case)]
        guard_cases = [case for case in guard_cases if _matches(case)]
        unmatched = wanted - {
            token
            for case in (*model_cases, *guard_cases)
            for token in (case.plan_id.casefold(), case.name.casefold())
        }
        if unmatched:
            raise ValueError(
                "No test-model case matches: " + ", ".join(sorted(unmatched))
            )
    elif not include_heavy:
        model_cases = [case for case in model_cases if not case.heavy]

    return model_cases, guard_cases


def _write_fixture_columns(workbook: xw.Book) -> None:
    """Add the fixture columns some specs declare, once, permanently."""
    for profile_key, columns in FIXTURE_COLUMNS.items():
        config = _CONFIG_BY_PROFILE_KEY[profile_key]
        sheet = workbook.sheets[config.sheet_name]
        table = sheet.api.ListObjects(config.table_name)
        existing = {
            table.ListColumns(index).Name
            for index in range(1, table.ListColumns.Count + 1)
        }
        for name, formula in columns.items():
            if name in existing:
                continue
            column = table.ListColumns.Add()
            column.Name = name
            column.DataBodyRange.Formula = formula


def build_test_models_workbook(
    workbook_path: Path = DEFAULT_WORKBOOK_PATH,
    definitions_path: Path = DEFAULT_DEFINITIONS_PATH,
    *,
    case_filter: set[str] | None = None,
    include_heavy: bool = False,
    verbose: bool = False,
    progress: _Progress | None = None,
) -> tuple[NameSyncResult, list[str]]:
    """Write every selected case's sheet and sync the catalog names.

    Returns the name-sync result and the ordered list of sheet names built,
    so a caller can verify exactly what was produced rather than re-deriving
    the selection.
    """
    progress = progress or _Progress(enabled=verbose, run_start=time.monotonic())

    document = load_catalog_document(definitions_path)
    closures = document.functions_for_sheet("Regression")
    model_cases, guard_cases = _selected_cases(case_filter, include_heavy)
    total_sheets = len(model_cases) + len(guard_cases)

    progress.phase(
        f"Plan: {len(model_cases)} model + {len(guard_cases)} guard "
        f"= {total_sheets} sheet(s)"
    )
    for case in (*model_cases, *guard_cases):
        progress.detail(
            f"{case.plan_id:<5} {case.sheet_name}"
            + ("  [heavy]" if getattr(case, "heavy", False) else "")
        )

    # Compute every oracle BEFORE opening Excel. A spec error should fail in
    # seconds, not after a multi-minute build has already written 30 sheets.
    progress.phase("Oracles: computing")
    model_expectations = [
        calculate_regression_spec_case(case) for case in model_cases
    ]
    guard_expectations = [calculate_guard_state_case(case) for case in guard_cases]
    progress.phase("Oracles: done")

    headers = {
        config.name: load_csv_rows(config.default_csv_path, config)
        for config in (LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS)
    }

    workbook_path = workbook_path.resolve()
    built: list[str] = []
    start = time.monotonic()

    def _run_write() -> None:
        # The first session writes every case sheet under XL_CALCULATION_MANUAL
        # so the spec-block writes don't trigger a recalc mid-build, then saves
        # the workbook. Wrapped in _retry_on_open for the same reason as
        # build_production.py / build_univariate.py: a workbook left open in
        # Excel during an 80-minute run would otherwise abort the whole build
        # rather than prompting the user to close it and retrying.
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            # Always from scratch. Sheet names encode the case list, so an
            # incremental build over a previous run would silently keep
            # sheets for cases that have since been renamed or dropped —
            # and a stale sheet that still calculates is the worst possible
            # QC artifact.
            workbook = app.books.add()
            try:
                app.api.Calculation = XL_CALCULATION_MANUAL
                for sheet in list(workbook.sheets)[1:]:
                    sheet.delete()
                workbook.sheets[0].name = "LAMBDA_functions"

                progress.phase("Catalog + data sheets")
                write_catalog_sheet(workbook, document.functions)
                for config in (LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS):
                    config_headers, config_rows = headers[config.name]
                    write_csv_dataset_sheet(
                        workbook, config_headers, config_rows, config
                    )
                _write_fixture_columns(workbook)
                progress.phase(f"Sheets: writing {total_sheets}")

                for expected in model_expectations:
                    with progress.sheet(
                        len(built) + 1, total_sheets, expected.case
                    ):
                        write_test_model_sheet(
                            workbook, expected, document.regression_sheet_notes,
                            closures,
                        )
                    built.append(expected.case.sheet_name)
                for guard in guard_expectations:
                    with progress.sheet(len(built) + 1, total_sheets, guard.case):
                        write_guard_state_sheet(
                            workbook, guard, document.regression_sheet_notes,
                            closures,
                        )
                    built.append(guard.case.sheet_name)

                progress.phase("Save")
                app.api.Calculation = XL_CALCULATION_SEMIAUTOMATIC
                workbook.save(str(workbook_path))
            finally:
                workbook.close()

    # Checked before the write, not left to _retry_on_open below. This phase
    # builds a fresh workbook and only touches workbook_path at the closing
    # save, so a lock surfaces after every one of the ~47 case sheets has been
    # written — the worst place in the project to discover it.
    warn_if_workbook_open(
        workbook_path, action_label=f"{workbook_path.name} is open in Excel"
    )

    try:
        _retry_on_open(
            f"{workbook_path.name} is open in Excel",
            _run_write,
            retry_rpc=True,
        )
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "open or save", exc)

    write_elapsed = time.monotonic() - start

    progress.phase("Sync names")
    sync_start = time.monotonic()
    result = sync_workbook_names(workbook_path, document.workbook_functions)
    sync_elapsed = time.monotonic() - sync_start

    # The rebuild is not optional here, for the same reason CLAUDE.md gives
    # for the Regression artifact: the name sync rewrites the catalog's
    # workbook-scoped LAMBDAs, and a per-sheet Calculate() does not rebuild
    # the dependency tree behind them, so every engine value would read nan.
    # This artifact has no Data Tables, so the rebuild is cheap.
    progress.phase("CalculateFullRebuild")
    rebuild_start = time.monotonic()
    # Probe: time just the CalculateFullRebuild call, separately from the
    # surrounding xw.App() lifecycle (cold-start, open, save). The two
    # numbers together make the rebuild cost visible against the rest of the
    # session — useful when the next person asks "is folding the second
    # session into the first one worth it?" since the answer is then in
    # the log rather than a hand extrapolation.
    rebuild_only: list[float] = [0.0]

    def _run_rebuild() -> None:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            workbook = app.books.open(str(workbook_path))
            try:
                inner = time.monotonic()
                app.api.CalculateFullRebuild()
                rebuild_only[0] = time.monotonic() - inner
                workbook.save(str(workbook_path))
            finally:
                workbook.close()

    try:
        _retry_on_open(
            f"{workbook_path.name} is open in Excel",
            _run_rebuild,
            retry_rpc=True,
        )
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "recalculate", exc)

    print(f"Timing: write sheets  {write_elapsed:.1f}s")
    print(f"Timing: sync names    {sync_elapsed:.1f}s")
    print(f"Timing: rebuild       {time.monotonic() - rebuild_start:.1f}s")
    print(f"Timing:   CalculateFullRebuild only   {rebuild_only[0]:.1f}s")

    return result, built


def _run_verify(
    workbook_path: Path, built: list[str], progress: _Progress
) -> int:
    """Verify the built workbook; return the process exit code."""
    from tools.inspect_test_model_sheets import verify_test_model_workbook

    progress.phase(f"Verify: reading {len(built)} sheet(s)")
    started = time.monotonic()
    per_case: list[tuple[str, int]] = []

    def _report(index: int, total: int, case, failure_count: int) -> None:
        per_case.append((case.plan_id, failure_count))
        progress.detail(
            f"{index:>3}/{total} {case.plan_id:<5} {case.sheet_name:<32}"
            + (f" {failure_count} mismatch(es)" if failure_count else " ok")
        )

    try:
        with xw.App(visible=False, add_book=False) as app:
            app.api.DisplayAlerts = False
            app.api.AskToUpdateLinks = False
            workbook = app.books.open(str(workbook_path))
            try:
                failures = verify_test_model_workbook(workbook, built, _report)
            finally:
                workbook.close()
    except OPEN_WORKBOOK_ERRORS as exc:
        raise_excel_access_error(workbook_path, "verify", exc)

    print(f"Timing: verify        {time.monotonic() - started:.1f}s")
    if not failures:
        print("Verify: passed")
        return 0

    # Per-case totals BEFORE the failure list. A wide case can contribute
    # tens of thousands of lines, and the first live run of the sibling
    # verifier buried 12 real mismatches under 22,886 from a single case —
    # this is the line that would have said so on sight.
    offenders = [(plan_id, n) for plan_id, n in per_case if n]
    print("Verify: mismatch totals by case:")
    for plan_id, count in sorted(offenders, key=lambda pair: -pair[1]):
        print(f"    {plan_id:<5} {count}")

    for message in failures[:200]:
        print(f"ERROR {message}")
    if len(failures) > 200:
        print(f"ERROR ... and {len(failures) - 200} more mismatch(es).")
    print(f"Verify: FAILED with {len(failures)} mismatch(es).")
    return 1


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the test-model workbook build."""
    parser = argparse.ArgumentParser(
        description="Build Lambda_Library_TestModels.xlsx — one sheet per test model."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK_PATH,
        help="Path to the test-model workbook to create.",
    )
    parser.add_argument(
        "--definitions",
        type=Path,
        default=DEFAULT_DEFINITIONS_PATH,
        help="Path to the JSON file containing lambda definitions.",
    )
    parser.add_argument(
        "--cases",
        type=str,
        default=None,
        help="Comma-separated plan IDs or case names to build (e.g. 'M09,G10'). "
        "Overrides --include-heavy for the cases named.",
    )
    parser.add_argument(
        "--include-heavy",
        action="store_true",
        help="Also build the cases marked heavy — currently L08 (173 Fixed "
        "Effects groups over 2909 rows) and L05 (Kitchen Sink Profile, "
        "k=19, n=2117). L08 is gated on sheet-build cost; L05 is gated on "
        "the statsmodels-vs-Excel floating-point floor at fdd=5/6 that "
        "both implementations agree on. Their Python oracles run in the "
        "unit suite either way; only the sheet is gated.",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="After building, read every sheet back and compare against the "
        "oracle. Exits 1 on drift.",
    )
    parser.add_argument(
        "--no-launch",
        action="store_true",
        help="Do not open the workbook in Excel when the build finishes.",
    )
    parser.add_argument(
        "--log",
        type=Path,
        default=None,
        help="Where to archive this run's transcript. Defaults to "
        f"'{RUN_LOG_DIR_NAME}/<script> <flags>{RUN_LOG_FILE_SUFFIX}', which is the naming "
        "the hand-uploaded logs already used.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Name every sheet as it is written and every sheet as it is "
        "verified, with per-sheet and cumulative timings. Worth it on a full "
        "run: ~46 sheets through COM takes minutes, and this is what makes a "
        "hang attributable to a case without rerunning.",
    )
    return parser.parse_args()


def main() -> None:
    """Build the test-model workbook, optionally verifying it.

    The whole run is teed into ``excel-only-runs/`` so a failure is a file
    somebody can hand over, not a terminal scrollback. The script needs
    Excel, so the transcript only exists on the Windows machine that ran
    it; ``excel-only-runs/`` is gitignored by default so a casual re-run
    does not promote the latest transcript to a commit.
    """
    args = parse_args()
    case_filter = (
        {token.strip() for token in args.cases.split(",") if token.strip()}
        if args.cases
        else None
    )
    log_path = (
        args.log
        if args.log is not None
        else run_log_path(ROOT_DIR, Path(__file__).name, sys.argv[1:])
    )
    command = f"python {Path(__file__).name} " + " ".join(sys.argv[1:])

    with tee_run_log(log_path, command):
        total_start = time.monotonic()
        progress = _Progress(enabled=args.verbose, run_start=total_start)

        result, built = build_test_models_workbook(
            workbook_path=args.workbook,
            definitions_path=args.definitions,
            case_filter=case_filter,
            include_heavy=args.include_heavy,
            verbose=args.verbose,
            progress=progress,
        )

        print(f"Workbook: {args.workbook.resolve()}")
        print(f"Sheets built: {len(built)}")
        print_name_sync_summary(result)

        exit_code = 0
        if args.verify:
            exit_code = _run_verify(args.workbook.resolve(), built, progress)

        print(f"Timing: total         {time.monotonic() - total_start:.1f}s")
        print(f"Log: {log_path}")

    if not args.no_launch and exit_code == 0:
        subprocess.Popen(["cmd", "/c", "start", "", str(args.workbook.resolve())])

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
