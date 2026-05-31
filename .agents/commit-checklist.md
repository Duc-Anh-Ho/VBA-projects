---
file: ./.agents/commit-checklist.md
description: Commit checklist injected by .agents/hooks/commit/inject-checklist.mjs before any git commit. Edit this file to adapt the checklist without touching the hook script.
scope: project
updated-at: 2026-05-30
---

# Commit checklist

Text below is injected as additionalContext when a `git commit` command is detected. Lines starting with `-` become bullets in the prompt.

- VBA change: did you EXPORT the edited code from the VBE back to the `VBA-files-*/` text folders before committing? The binary workbook is not the source of truth.
- Did you stage the matching scope only? Keep agent-config edits (.agents/, .claude/, CLAUDE.md) in a separate commit from VBA/code edits.
- Did you avoid committing machine-local or secret files (settings.local.json, tools.local.json are gitignored)?
- Subject starts with a type word (Add / Update / Fix / Refactor / Remove / Docs / Chore) and explains WHY, not just what.
- If a planning task is now done, did you update docs/plan/plan.md or the relevant memory file?
