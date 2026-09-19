---
name: check-decisions-log
description: Grep docs/DECISIONS.md before proposing a nontrivial design change to the Regression/Univariate engine, sheet layout, or named-range scoping — several clean-looking ideas here are recorded as tried-and-reverted. Use before implementing (not after) any change that smells architectural, and add a new dated entry when you make a genuinely new nontrivial reversible-looking choice.
---

# Check the decisions log before re-trying a reverted idea

`docs/DECISIONS.md` is a long, append-only log of this project's design decisions — including several that were tried, shipped, and later **SUPERSEDED** or explicitly reverted. An agent that hasn't read the whole file (it's thousands of lines) can easily re-propose one of these as if it were a fresh idea.

## When to check
Before implementing, not after, when a change smells like any of these:
- Splitting a workbook/artifact into multiple pieces, or merging pieces back together.
- Changing a materialized zone (e.g. Model Context) from individual cells to a single spill (`VSTACK`), or vice versa.
- Using `book.names.add` instead of sheet-scoped names, or changing named-range scope in general.
- Switching between `OFFSET` and `TAKE`/dynamic-array patterns for a named range.
- Per-point COM loops for chart formatting/labels instead of a masked overlay series.
- Silently switching or ignoring a computation instead of flagging it (this repo's repeated philosophy: visible failure over a silent wrong number).
- Anything touching how `Log`/`Log (drop ≤ 0)` transforms are compared, or the QC comparison-scale conventions.

## How to check
```
grep -n "<keyword>" docs/DECISIONS.md
```
Search for the feature/mechanism name, not just the symptom. Known reverted/superseded threads worth searching for by name: "two-artifact" / "Univariate becomes its own workbook" (reversed at v3.0/v3.1), "model context is individual cells" (spill rejected), heatmap → profile-NLL line chart switch for Weibull/Gamma.

## What to do with what you find
- **Matches an existing decision** → follow it, or if you believe it should change, say so explicitly in the PR description with the reasoning — don't silently diverge.
- **Genuinely new nontrivial, reversible-looking choice** (not covered by any existing entry) → add a new dated entry to `docs/DECISIONS.md` in the same PR. The log is a rule to maintain, not just a reference to consult.
