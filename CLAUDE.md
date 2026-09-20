# CLAUDE.md — Claude Code entry point

@AGENTS.md

## Claude Code

The project rules live in **[AGENTS.md](AGENTS.md)**, imported above. That file is the single source of truth for every agent tool — Claude Code reads it through this import, and Copilot CLI, Codex, and Cursor read it directly. **This file deliberately keeps no copy of those rules: add them to `AGENTS.md`, never here**, so that every tool sees the same text and none of them can drift.

Claude Code specifics:

- **Project skills** live in `.claude/skills/` — `regression-pr-shape`, `test-model-case-checklist`, `sheet-writer-conventions`, `deep-verify-before-done`, `static-sheet-regen`, `check-decisions-log`, `version-bump-check`. They are the edit-time triggers for the rules in `AGENTS.md`.
- `.agents/skills/` carries `brag-sheet`, a vendored third-party skill (MIT, from `github/awesome-copilot`) that is not a project rule trigger and is mirrored in both `.claude/skills/` and `.agents/skills/`. The project skills above are **not** mirrored there: a tool that reads `.agents/` rather than `.claude/` sees `brag-sheet` and nothing else, and a mirrored copy added later would be a second copy to keep in step.

See `AGENTS.md` for everything else — build and verify workflows, the testing regime, sheet layout and the spec block, chart conventions, and the QC comparison scale.
