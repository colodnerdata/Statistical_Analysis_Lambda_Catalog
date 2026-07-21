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

.PHONY: verify-headless verify-deep verify

# Fast screen — pure zipfile/lxml invariant tests. <1 s on Linux.
verify-headless:
	uv run --frozen pytest tests/test_workbook_invariants.py -v

# Spec-driven verifier. Reuses build_qc.verify_test_sheets(..., skip_dummy=True)
# against the production sheets. Requires Excel; CI runs this on windows-latest.
verify-deep:
	uv run --frozen python build_production.py --verify --no-launch --skip-data-table-calculations

# Both layers. The headless check is auto-discovered on Linux; the deep
# check runs on Windows in CI. The deep check shell-exits 1 on drift.
verify: verify-headless verify-deep
