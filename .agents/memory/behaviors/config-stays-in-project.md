---
file: ./.agents/memory/behaviors/config-stays-in-project.md
name: config stays in project
description: User prefers all agent config and memories live inside the project dir, git-tracked, never in the user home folder. Documents the one known Claude Code exception.
type: feedback
scope: project
updated-at: 2026-05-30
---

Rule: keep ALL Claude Code (and other agent) config and memory files inside the
project directory (.claude/ or .agents/), never in the user folder (~/.claude,
%USERPROFILE%\.claude).

Why: the user wants the project to be the single source of truth - portable across
machines, version-controlled, no hidden state in user-level paths. Same reasoning as
memory-storage.md (memories go in .agents/memory/), generalized to all config.

How to apply:
- New config files (settings, hooks, scripts, rules, tools manifest) go under
  .claude/ or .agents/ in the project, never under user-level paths.
- Path fields inside config (for example autoMemoryDirectory) point to project paths.
- When a Claude Code feature appears to require a user-level file, search the docs for
  a project-scoped alternative BEFORE accepting user scope. Verify, do not guess.

Known exceptions (Claude Code does not support project scope):
- ~/.claude/keybindings.json - strictly user-scoped per code.claude.com docs. There is
  no project-local keybindings file.

When you discover a new exception: verify against the official docs, add the path and
doc URL here, and tell the user.

Related: see features/memory-storage.md and features/nested-claude-md.md.
