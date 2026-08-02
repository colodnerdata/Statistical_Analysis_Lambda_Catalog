# Statistical_Analysis_Lambda_Catalog — make targets
#
# Layer 1 (verify-headless): pure zipfile + lxml structural check, runs on
# Linux in <1 s. Catches packaging regressions: dangling defined names,
# broken Content_Types / workbook.xml.rels, orphan chart relationships,
# #REF!/#NAME? cached values, sheet drift.
#
# Layer 2 (verify-deep / verify): spec-driven xlwings check that requires
# Excel. Reuses build_qc.verify_test_sheets against the production sheets.
# On Windows; on Linux this will fail with a pywintypes import error.
#
# verify runs both layers; verify-headless is the fast screen, verify-deep
# is the source of truth.

.PHONY: verify-headless verify-deep verify-deep-univariate verify

# Fast screen — pure zipfile/lxml invariant tests. <1 s on Linux.
verify-headless:
	uv run --frozen pytest tests/test_workbook_invariants.py -v

# Spec-driven verifier for the Regression workbook. Reuses
# build_qc.verify_test_sheets(..., skip_dummy=True) against the production
# sheets. Requires Excel; not run in CI on GitHub-hosted runners (no
# Microsoft Office). Always recalculates (the recalc is the source of truth:
# the verifier's per-sheet Calculate doesn't rebuild the dependency tree
# after a name sync, so the Regression engines need CalculateFullRebuild).
verify-deep:
	uv run --frozen python build_production.py --verify --no-launch

# Spec-driven verifier for the standalone Univariate workbook. Runs
# build_qc.verify_test_sheets(..., skip_dummy=True, skip_regression=True) —
# this artifact has no Regression / Mileage / Production Lots sheets, so
# those checks are skipped; the Life Expectancy and Univariate checks run.
# Requires Excel; not run in CI.
verify-deep-univariate:
	uv run --frozen python build_univariate.py --verify --no-launch

# Both layers for both artifacts. The headless check is auto-discovered on
# Linux; run the deep checks on a machine with Microsoft Excel. Each deep
# check shell-exits 1 on drift.
verify: verify-headless verify-deep verify-deep-univariate
