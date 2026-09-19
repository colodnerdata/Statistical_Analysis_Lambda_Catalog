---
name: check-decisions-log
description: Consult docs/DECISIONS.md before implementing a design change, not after, and append a dated entry when the choice is genuinely new. The classes worth checking — splitting or merging a workbook/artifact, changing a materialized zone between individual cells and one spill, named-range scope or book.names.add, OFFSET vs TAKE/dynamic arrays, per-point COM loops vs a masked overlay series, and anything touching Log / Log-drop transforms or the QC comparison scale.
---

# Check the decisions log before re-trying a reverted idea

`docs/DECISIONS.md` is a long, append-only log of this project's design decisions — including several that shipped and were later **SUPERSEDED** or explicitly reverted. An agent that hasn't read the whole file (thousands of lines) can easily re-propose one of these as if it were a fresh idea.

The classes of change that trigger a check are stated in `AGENTS.md` → *PR workflow*. This skill is the **how**, not a second copy of the when.

## How to search

```
grep -n "<mechanism>" docs/DECISIONS.md
```

Search for the **mechanism name**, not the symptom. Threads known to be recorded as tried-then-reverted or superseded:

- **"two-artifact" / "Univariate becomes its own workbook"** — reversed at v3.0/v3.1.
- **"model context is individual cells"** — the single-`VSTACK`-spill version was rejected.
- **heatmap → profile-NLL line chart** for the Weibull/Gamma fits.

## What to do with what you find

- **Matches an existing decision** → follow it. If you believe it should change, say so explicitly in the PR description with the reasoning — don't silently diverge.
- **Genuinely new nontrivial, reversible-looking choice** (covered by no existing entry) → add a dated entry to `docs/DECISIONS.md` in the same PR. The log is a rule to maintain, not just a reference to consult.
