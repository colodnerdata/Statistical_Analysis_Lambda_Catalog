"""Automated testing script for Beta grid search performance vs accuracy tradeoffs.

This script tests the Beta distribution grid search performance by:
1. Building minimal Beta-only workbooks with different grid sizes [6, 8, 10, 12, 20]
2. Measuring build time and recalculation time for each
3. Extracting Beta parameters from the fitted model (I11 = alpha, K11 = beta)
4. Comparing with oracle parameters from scipy MLE
5. Reporting accuracy vs performance tradeoffs

The minimal workbook includes only the catalog LAMBDA sheet, the Life
Expectancy data sheet, and the Univariate sheet (no charts, no other
distributions, no version history). This isolates Beta's two-stage
``Full_Factorial`` grid search for timing.
"""

from __future__ import annotations

import csv
import subprocess
import sys
import time
from pathlib import Path

import xlwings as xw
from scipy import optimize

# Add project root to path so we can import lambda_catalog
sys.path.insert(0, str(Path(__file__).parent.parent))

from lambda_catalog.analyze_univariate import nll_beta
from lambda_catalog.workbook_helpers import col_letter
from lambda_catalog.write_sheet_csv_dataset import LIFE_EXPECTANCY, load_csv_rows

ROOT_DIR = Path(__file__).resolve().parent.parent
DIST_DIR = ROOT_DIR / "dist"
RESULTS_FILE = ROOT_DIR / "beta_grid_search_results.csv"

# The fitting table layout is stable: distribution rows start at row 5 and Beta
# is the 7th distribution, landing on row 11. Alpha lives in column I
# (_C_T1_VAL = 9), beta in column K (_C_T2_VAL = 11).
_BETA_ROW = 11
_BETA_ALPHA_COL = 9
_BETA_BETA_COL = 11


def _life_expectancy_column() -> list[float | None]:
    """Return the dataset column the Univariate sheet fits."""
    headers, rows = load_csv_rows(LIFE_EXPECTANCY.default_csv_path, LIFE_EXPECTANCY)
    column = headers.index("Life expectancy")
    return [row[column] if column < len(row) else None for row in rows]


def _beta_oracle_parameters() -> tuple[float, float]:
    """Compute the Beta MLE used as the oracle for the performance comparison."""
    data = _life_expectancy_column()
    numeric = [float(value) for value in data if isinstance(value, (int, float))]
    data_min = min(numeric)
    data_max = max(numeric)
    data_range = data_max - data_min
    pad = max(data_range * 0.001, 1e-30)
    scale_ = data_range + 2.0 * pad
    z = [(value - data_min + pad) / scale_ for value in numeric]

    result = optimize.minimize(
        lambda params: nll_beta(z, params[0], params[1]),
        x0=(2.0, 2.0),
        bounds=((0.2, 50.0), (0.2, 50.0)),
        method="L-BFGS-B",
    )
    if not result.success:
        raise RuntimeError(f"Failed to compute Beta oracle parameters: {result.message}")
    return float(result.x[0]), float(result.x[1])


def _extract_beta_parameters(workbook_path: Path) -> tuple[float, float]:
    """Read the Beta alpha and beta values from the fitting table at I11 / K11."""
    alpha_cell = f"{col_letter(_BETA_ALPHA_COL)}{_BETA_ROW}"
    beta_cell = f"{col_letter(_BETA_BETA_COL)}{_BETA_ROW}"

    app = xw.App(visible=False, add_book=False)
    try:
        workbook = app.books.open(str(workbook_path))
        try:
            sheet = workbook.sheets["Univariate"]
            return float(sheet.range(alpha_cell).value), float(sheet.range(beta_cell).value)
        finally:
            workbook.close()
    finally:
        app.quit()


def _build_workbook(workbook_path: Path, grid_size: int) -> None:
    """Run ``build_beta_only.py`` for ``grid_size`` and write to ``workbook_path``.

    Passes ``--skip-recalculation --no-calculation`` so the build leaves the
    workbook in Manual mode and the caller can measure recalc on its own.
    """
    args = [
        "--workbook", str(workbook_path),
        "--beta-grid-size", str(grid_size),
        "--skip-recalculation",
        "--no-calculation",
    ]
    result = subprocess.run(
        [sys.executable, str(Path(__file__).parent / "build_beta_only.py")] + args,
        capture_output=True,
        text=True,
        timeout=300,
    )
    if result.returncode != 0:
        print(result.stdout)
        print(result.stderr)
        raise RuntimeError(
            f"build_beta_only.py failed for grid_size={grid_size} "
            f"with return code {result.returncode}"
        )


def _recalculate_workbook(workbook_path: Path) -> float:
    """Open the workbook in Automatic mode, run CalculateFullRebuild, save, close.

    Returns the wall-clock time for the open + calculate + save + close cycle.
    """
    start = time.monotonic()
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(workbook_path))
        try:
            app.api.Calculation = xw.constants.Calculation.xlCalculationAutomatic
            wb.api.CalculateFullRebuild()
            wb.save()
        finally:
            wb.close()
    finally:
        app.quit()
    return time.monotonic() - start


def main() -> None:
    """Run the Beta grid search performance test."""
    print("Beta Grid Search Performance vs Accuracy Test")
    print("=" * 50)

    grid_sizes = [6, 8, 10, 12, 20]

    oracle_alpha, oracle_beta = _beta_oracle_parameters()
    print("Oracle parameters (scipy MLE):")
    print(f"  Alpha: {oracle_alpha:.6f}")
    print(f"  Beta:  {oracle_beta:.6f}")
    print()

    DIST_DIR.mkdir(exist_ok=True)
    results = []

    for grid_size in grid_sizes:
        print(f"Testing grid size {grid_size}...")
        workbook_path = DIST_DIR / f"Lambda_Library_BetaOnly_GridSize{grid_size}.xlsx"

        try:
            # 1) Sheet writing only — workbook stays in Manual mode, no recalc.
            build_start = time.monotonic()
            _build_workbook(workbook_path, grid_size)
            build_time = time.monotonic() - build_start

            # 2) Recalc — open in Automatic, CalculateFullRebuild, save, close.
            recalc_time = _recalculate_workbook(workbook_path)

            # 3) Extract fitted Beta parameters from the recalculated workbook.
            alpha, beta = _extract_beta_parameters(workbook_path)

            alpha_error = abs(alpha - oracle_alpha) if alpha != 0.0 else float('inf')
            beta_error = abs(beta - oracle_beta) if beta != 0.0 else float('inf')
            alpha_error_pct = (alpha_error / oracle_alpha * 100) if oracle_alpha != 0.0 else float('inf')
            beta_error_pct = (beta_error / oracle_beta * 100) if oracle_beta != 0.0 else float('inf')

            results.append({
                'grid_size': grid_size,
                'build_time': build_time,
                'recalc_time': recalc_time,
                'total_time': build_time + recalc_time,
                'alpha': alpha,
                'beta': beta,
                'alpha_error': alpha_error,
                'beta_error': beta_error,
                'alpha_error_pct': alpha_error_pct,
                'beta_error_pct': beta_error_pct,
            })

            print(f"  Build time:  {build_time:.1f}s")
            print(f"  Recalc time: {recalc_time:.1f}s")
            print(f"  Alpha: {alpha:.6f} (error: {alpha_error:.6f}, {alpha_error_pct:.2f}%)")
            print(f"  Beta:  {beta:.6f} (error: {beta_error:.6f}, {beta_error_pct:.2f}%)")
            print()

        except Exception as e:
            print(f"  FAILED: {e}")
            print()

    print("Summary Results:")
    print("-" * 90)
    print(f"{'Grid':<6} {'Build(s)':<9} {'Recalc(s)':<10} {'Total(s)':<9} {'Alpha Err%':<11} {'Beta Err%':<10}")
    print("-" * 90)
    for r in results:
        print(f"{r['grid_size']:<6} {r['build_time']:<9.1f} {r['recalc_time']:<10.1f} "
              f"{r['total_time']:<9.1f} {r['alpha_error_pct']:<11.2f} {r['beta_error_pct']:<10.2f}")

    with open(RESULTS_FILE, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'grid_size', 'build_time', 'recalc_time', 'total_time',
            'alpha', 'beta', 'alpha_error', 'beta_error',
            'alpha_error_pct', 'beta_error_pct',
        ])
        writer.writeheader()
        writer.writerows(results)

    print(f"\nDetailed results saved to: {RESULTS_FILE}")


if __name__ == "__main__":
    main()
