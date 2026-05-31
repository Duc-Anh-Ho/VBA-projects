---
file: ./.claude/skills/settings-sync/SKILL.md
name: settings-sync
description: Deterministic bi-directional sync between settings.local.json and settings.json. Local-only keys stay in local; all other keys are merged into both files.
scope: agent
updated-at: 2026-05-31
allowed-tools: Read Edit Write
---

# /settings-sync

Synchronize `.claude/settings.local.json` and `.claude/settings.json` so both files share
all common keys. Local-only keys remain in `settings.local.json` only and are never written
to `settings.json`.

## Local-only keys

These keys are rejected by Claude Code when present in the git-tracked `settings.json`.
They must live only in `settings.local.json`.

- autoMemoryDirectory
- autoMode
- useAutoModeDuringPlan

When merging: strip these keys from the `settings.json` output. Keep them in `settings.local.json`.

## Rules

- `settings.local.json` always wins on conflicts for shared keys.
- Missing shared keys are filled from the other file.
- Local-only keys are never copied to `settings.json`.
- Result must be deterministic and idempotent.

## Steps

### 1. Read files

Read `.claude/settings.json` and `.claude/settings.local.json`. If one does not exist, use the
existing file as source and create the missing file (omitting local-only keys if creating settings.json).

### 2. Show diff

Classify keys into:
- only_local (not local-only key) -> will be added to shared
- only_local (local-only key) -> stays in local only, NOT copied to shared
- only_shared -> will be added to local
- conflict -> values differ (local will overwrite)
- same -> no change

Highlight conflicts and local-only keys clearly. Wait for user confirmation before writing.

### 3. Merge

For `settings.json`: start with all keys, add missing shared keys from local (excluding local-only),
local overwrites on conflict, never include local-only keys.

For `settings.local.json`: start with all keys, add missing keys from shared, local overwrites.

Array merge rules (for `permissions.allow / deny / ask`): union + deduplicate, then:
- Wildcard consolidation: a specific entry is redundant if a wildcard in the same array covers it.
  `Bash(mv *)` covers `Bash(mv "D:/some/path")`. Remove redundant specifics.
- Proactive wildcard creation: if 2+ specific entries share an obvious common prefix and no wildcard
  exists, propose `Tool(prefix *)` and drop the specifics. Confirm before writing (wildcards widen scope).
- Dangerous commands -> put in `ask`, never `deny`, unless the user explicitly says "deny".

Hooks: merge by event name -> then by matcher; local overwrites on conflict.

### 4. Write

Write both files. They will NOT be identical: `settings.local.json` has local-only keys that
`settings.json` does not. All shared keys must be identical after write.

### 5. Report

Report: keys added (local->shared), keys added (shared->local), local-only keys kept,
conflicts resolved, redundant entries removed, wildcards introduced, unchanged keys.
