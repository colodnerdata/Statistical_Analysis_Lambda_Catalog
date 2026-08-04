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

.PHONY: verify-headless verify-deep verify-deep-univariate verify-test-models verify

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

# The regression test-model suite, one worksheet per case. Builds
# Lambda_Library_TestModels.xlsx (gitignored — a fixture, not an artifact),
# then reads every sheet back against its Python oracle without writing to
# any of them. Requires Excel; not run in CI. Excludes the heavy L08 case by
# default — add --include-heavy for it; its Python oracle runs in the unit
# suite regardless.
#
# --verbose because this one runs for minutes over ~46 sheets: it names each
# sheet BEFORE writing it, so an interrupted run leaves the offending case on
# screen. The whole transcript is archived to "Local Run Logs/" either way,
# stderr included, so a com_error traceback is a file somebody can hand over
# rather than a terminal scrollback.
verify-test-models:
	uv run --frozen python build_test_models.py --verify --no-launch --verbose

# Both layers for both artifacts, plus the test-model suite. The headless
# check is auto-discovered on Linux; run the deep checks on a machine with
# Microsoft Excel. Each deep check shell-exits 1 on drift.
verify: verify-headless verify-deep verify-deep-univariate verify-test-models
