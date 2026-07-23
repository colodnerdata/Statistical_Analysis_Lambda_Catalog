## Summary

The four top-level docs had drifted into overlapping roles and grown past their natural scope. ROADMAP was 1,538 lines (version plan + design record + architectural reference fused); TODOs mixed active work with pre-resolved design decisions; README duplicated the LAMBDA_functions sheet with a hand-curated function reference; CONTRIBUTING duplicated the cell-styling color table from CLAUDE.md / AGENTS.md.

Split into six files, each with one role:

- **README.md** (41 lines, was 261) — first-impression overview + how to use the workbook. Function reference dropped; the `LAMBDA_functions` sheet in the workbook is the canonical, always-in-sync catalog. New "Documentation map" links to the other five files.
- **CONTRIBUTING.md** (248 lines, was 271) — dev/maintainer guide. Cell-styling color table dropped; CLAUDE.md / AGENTS.md remain the source.
- **ROADMAP.md** (459 lines, was 1,538) — version plan only. One paragraph per milestone + headline capabilities + cross-links to DECISIONS and ARCHITECTURE for depth. ToolPak parity reference stays.
- **ARCHITECTURE.md** (507 lines, new) — foundational patterns that do not change version-to-version: naming convention, function categories, the Role / Type / Sequence taxonomy, the Model Spec block (A–L), the data-transformation taxonomy, the reserved-spec-column pattern.
- **DECISIONS.md** (1,010 lines, new) — resolved design decisions with rationale, indexed by version. Each entry self-contained: question, resolution, rationale in one paragraph, date/PR. Includes the supersession log and the alias-layer table.
- **TODOs.md** (416 lines, was 400) — active work only. Resolved decisions migrated to DECISIONS; the duplicated "Design note — chart series data ranges" section (already in CONTRIBUTING) dropped. Cross-links to DECISIONS for context on the "why" of remaining items.

## Net change

- Files: 4 → 6 (two new)
- Total line count: 2,470 → 2,681 (+211)
- ROADMAP.md: 1,538 → 459 (–1,079, the largest single-file reduction)
- The +211 overall is the cost of making DECISIONS entries self-contained; previously the rationale lived compressed in-place inside ROADMAP

## Why this restructure

The structural win outweighs the line-count growth:

- ROADMAP is now readable in one pass — the version ladder and milestone summaries fit on screen
- The design record is findable by someone who needs the "why" without having to dig through 1,500 lines
- README no longer carries a stale function reference (it already showed the pre-v2.0 L–X residual zone)
- Cell-styling color table is in exactly one place (CLAUDE.md / AGENTS.md)
- TODOs is the active-work list, not the active-work list plus the design record

## Verification

- Every previously-RESOLVED TODO has a corresponding DECISIONS entry
- Every ROADMAP version section has a cross-link to DECISIONS.md or ARCHITECTURE.md (except v1.0, which correctly notes that it predates the open-decisions convention)
- README has zero references to function reference / common function signature / function name
- Cell-styling color table appears in only one place (CLAUDE.md / AGENTS.md)
- Spot-checks confirm v1.1 MLE-via-grid reframing, v2.0 spec block A–L layout, and v2.4 RANDARRAY() rejection all survived the move

## Out of scope

- **CLAUDE.md / AGENTS.md near-duplication** — these are 180 vs 182 lines with the same headings. Same problem at smaller scale, but the user scoped this PR to README/CONTRIBUTING/ROADMAP/TODOs only. Flagged for a future PR.
- **`Lambda_Library.xlsx`** — left unstaged; it's a leftover dirty state from the v2.1 sequence fix branch and is not part of the doc reorganization.

## Test plan

Docs-only change; no code or tests affected. CI coverage scope is unchanged.

🤖 Generated with [Claude Code](https://claude.com/claude-code)
