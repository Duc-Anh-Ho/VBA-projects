# AGENTS.md

This file bootstraps non-Claude agents for this repository.

## Required Startup Reads

Read these files before doing project work:
- `.agents/memory/MEMORY.md`
- `.agents/shared-rules.md`
- `CLAUDE.md`

Then read only the topic files referenced by `.agents/memory/MEMORY.md` that are
relevant to the current task.

## Project-Local State

All durable agent state for this repo must stay inside the project:
- memories: `.agents/memory/`
- shared rules: `.agents/shared-rules.md` and `.agents/rules/`
- hooks and scripts: `.agents/hooks/`, `.agents/scripts/`
- Codex config: `.codex/`
- Claude config: `.claude/`

Do not read or write user-level Codex or Claude memory/config paths unless the user
explicitly asks for those exact paths in the current turn.

## Coding Context

This is an Excel VBA add-in repository. The runnable code lives inside binary
workbooks, while diffable VBA exports live under `VBA-files-*` directories. For
behavior changes, edit exported `.bas`, `.cls`, and `.frm` files, not binary
workbooks.

Follow `CLAUDE.md` for repository architecture, source-of-truth rules, and VBA
conventions.
