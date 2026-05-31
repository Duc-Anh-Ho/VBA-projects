---
name: commit-clean
description: Create/update the "deploy" branch with a clean project snapshot - strips all agent config (.agents/, .claude/, CLAUDE.md), draft-and-reference/, and malware-investigation/. Only project VBA files, docs, and README remain.
---

# /commit-clean

Run the deploy-branch builder script and report results.

## Why this exists

The `deploy` branch is the clean, agent-free snapshot of the project suitable for:
- Public portfolio / open-source sharing (no internal Claude tooling visible).
- Distributing the add-in to users who only need the VBA code and workbooks.

This is why agent-config commits (.agents/, .claude/) and project commits (VBA files,
workbooks) must be kept separate - the script strips agent scope from deploy automatically.

## Steps

### 1. Run the script

```bash
node .agents/scripts/commit-clean.mjs
```

### 2. Report results

Show the script output. Key info:
- Whether the deploy branch was created or updated.
- Commit hash + file count.
- Reminder: not pushed (user pushes manually).

### 3. If the script fails

Report the error. Common causes:
- Worktree already exists from a failed run: run `git worktree prune` then retry.
- Deploy branch checked out elsewhere: remove the other worktree first.
- Git lock file: wait or remove `.git/worktrees/*/locked`.

## What gets stripped

Directories: `.agents`, `.claude`, `.codex`, `.gemini`, `.opencode`,
`draft-and-reference`, `malware-investigation`.

Root files: `CLAUDE.md`, `AGENTS.md`, `AGENTS.override.md`.

All nested `CLAUDE.md` files anywhere in the tree.

## What stays

Everything else: `final-installation/`, `xlwings/`, `part-tools/`, `Images/`,
`README.md`, `.gitignore`, `vbe-theme-regedit/`, `app/`.
