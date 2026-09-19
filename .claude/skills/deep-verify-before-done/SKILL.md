---
name: deep-verify-before-done
description: Clarifies what green CI does and does not prove for this workbook, and states the correct build-then-verify order. Use before reporting a Regression/Univariate/sheet-writer change complete, or before opening a PR touching lambda_catalog/ or any write_sheet_*.py file.
---

# Green CI is not proof the workbook is correct

This repo's verification has two layers, and only one runs in CI.

## Layer 1 — headless, CI-enforced, structural only
`poe verify-headless` / `tests/test_workbook_invariants.py` reads the committed `.xlsx` as a zipfile with `lxml` — no Excel required, runs on every push. It catches dangling names, `#REF!`/`#NAME?` cached literals, broken package parts, orphan chart-relationship targets, and sheet drift.

**A green Layer 1 run does not mean the workbook calculates correctly.** It cannot check that a formula produces the right *number* — only that the file's structure is sound.

## Layer 2 — the actual correctness check, cannot run here
`lambda_catalog.deep_verify.verify_test_sheets` (invoked via `python scripts/build_production.py --verify --no-launch`, or narrower slices like `verify-guards` / `verify-spec-errors` / `verify-models`) is the source of truth for cell-level correctness — it drives real Excel via xlwings/COM and compares cells against independent oracles. **This requires desktop Excel and cannot run in this container or in GitHub-hosted CI** (`windows-latest` has no Office; there is no `windows-verify` job, removed on purpose).

**If Excel isn't available in the current environment, say so explicitly rather than claiming verification passed.** Passing the pure-Python unit suite or Layer 1 is not the same claim as "Layer 2 verified this on a machine with Excel" — don't let those get conflated in a status update or PR description.

## Build-then-verify order matters
`poe verify` builds first, then screens (Layer 1) — deliberately. Screening a stale `dist/` artifact before rebuilding has previously passed clean even when the rebuild itself broke something. Never run a verification pass against a `dist/` you haven't just rebuilt.
