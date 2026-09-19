---
name: version-bump-check
description: Decide which version number to bump — LAMBDA catalog (library) version, workbook version, or both — before finalizing any change to lambda_catalog/ or a sheet writer. Use near the end of a change, once the diff is settled.
---

# Which version number moves

This repo ships two independent version numbers, and it's easy to bump the wrong one or forget one entirely. Per `CONTRIBUTING.md`'s "which version number moves" table:

| Change touches... | Bump |
|---|---|
| `lambda_functions.json` content or function count (new/changed/removed catalog LAMBDA) | **Library** version only |
| Sheet layout only (rows/columns/zones/charts on an existing sheet, no catalog change) | **Workbook** version only |
| Both a catalog change and a sheet-layout change | **Both** |

Before finalizing any change to `lambda_catalog/` or a `write_sheet_*.py` module, check the diff against this table and confirm the correct version(s) were bumped — don't assume a catalog change alone justifies skipping the workbook version, or vice versa, if the diff actually touches both.
