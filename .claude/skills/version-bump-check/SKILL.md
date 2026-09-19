---
name: version-bump-check
description: Decide which version number to bump — library (LAMBDA catalog), workbook, or both — before finalizing any change to lambda_catalog/ or a write_sheet_*.py module. Use near the end of a change, once the diff is settled.
---

# Which version number moves

**The authority is `CONTRIBUTING.md` → *Which version number moves*** — a three-row table: catalog content or function count → **library** version only; sheet layout with no catalog change → **workbook** version only; both → **both**. That section also carries the release procedure, whose step 5 is where the library version actually moves. This file deliberately does not restate the table — a second copy of a three-row rule is a divergence surface with no upside.

Two things this reminder adds, both about *timing* rather than content:

- **Check once the diff is settled, not as you go.** The table keys on what the diff touches, so a diff still in motion can answer it wrongly.
- **Neither number excuses the other.** A catalog edit does not mean the workbook version stands still if the same change also moved sheet layout — read the table against the whole diff.
