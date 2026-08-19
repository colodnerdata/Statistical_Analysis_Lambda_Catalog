# excel-only-runs

This directory holds diagnostic output from scripts that **require Microsoft Excel** to run — `python scripts/build_production.py --verify`, `python scripts/build_test_models.py --verify --no-launch`, and similar CLI invocations from a developer PowerShell session.

## Why this is tracked

The spec-driven verifier cannot run on the GitHub-hosted Linux CI without Excel (`xlwings` cannot dispatch `Excel.Application`). The headless screen (`poe verify-headless`) IS in CI because it needs only `zipfile` + `lxml`. When an Excel-required build fails, the diagnostics never reach a CI artifact.

**This directory is the cross-tool substitute for that missing CI.** Transcripts here are committed so a contributor — or a coding agent (Claude Code, Copilot, etc.) — opening the repo cold can read what is and isn't currently working on a Windows-only verifier build, the same way they'd read a CI log. A pull request that adds an Excel-required build with no transcript on its branch has no paper trail.

## What goes here

- Full stdout+stderr from `python scripts/build_production.py --verify`, `python scripts/build_test_models.py --verify --no-launch`, or `python scripts/build_demo_workbook.py --verify --no-launch`. **All three drivers write their own transcript here** — stderr and the traceback of an aborted run included — so producing one is not a manual step; committing it is. Pass `--log PATH` to send a run somewhere else (e.g. out of the way while chasing something that reruns a lot).
- Verbose-level output (`--verbose`) when comparing runs across debug sessions.
- ERROR excerpts from any single-case run (`python scripts/build_test_models.py --cases M09`) where tracking down a flaky sheet.

The naming convention `<script> <flags>.log` (set by `lambda_catalog/build_common.run_log_path`) means a `git log` on this directory reads as a list of what was actually run and when.

## This is the contributor's responsibility

The directory is **tracked, not gitignored**. Anyone who runs an Excel-required verifier build produces a transcript here, and the same run that produced the transcript also has the responsibility to commit it on the branch that did the work.

Conversely: if a transcript in this directory is *stale* — i.e. the build it captured has since changed and the failure modes it lists are not the current ones — it is also the contributor's responsibility to delete or replace it. Stale transcripts are worse than none; they mislead an agent or a reviewer into debugging a problem that no longer exists. Treat the directory as a working document, not a notebook.

A useful test before pushing a branch: `git status excel-only-runs/` should show only transcripts you generated this session. If it shows someone else's transcript from a session you didn't share, leave it alone — they own it — but add your own.

## When to retire a transcript

- **When the failure it captured is fixed** — delete the `.log` and let the next regression write a fresh one.
- **When the transcript is reusable** — i.e. it captures a known-good baseline that future runs should be compared against — leave it and add a one-line note to this README naming it, so it is not mistaken for an open problem.
- **When the underlying script changes in a way that the transcript would be misleading after** (e.g. flag semantics change) — delete and let a new run produce a replacement.

## What does not belong here

Production artifacts (`Lambda_Library.xlsx`) are tracked separately with their own `.gitignore` allowlist entry. The `Lambda_Library_TestModels.xlsx` fixture is also gitignored. Anything that should ship lives elsewhere; this directory is for the dev-side diagnostics that show how an Excel-required build actually ran.

## `.gitignore` note for the curious

This directory was gitignored at one point (in the `Local Run Logs/` and the first iteration of `excel-only-runs/`). It is no longer. Pulling this repo today should give you a populated `excel-only-runs/` from `git clone`, not an empty one you have to opt into populating.
