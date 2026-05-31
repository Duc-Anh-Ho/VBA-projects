---
file: ./.agents/memory/features/nested-claude-md.md
name: nested-claude-md
description: Project uses lazy-loaded nested CLAUDE.md per folder to save root-context tokens. When adding/moving/removing a tracked folder, update the folder CLAUDE.md, the parent folder map, and the root pointer list.
type: feature
scope: project
updated-at: 2026-05-30
---

This project uses nested CLAUDE.md files to keep the root CLAUDE.md slim. Claude Code
lazy-loads a subfolder CLAUDE.md only when a tool reads or edits a file inside that
subfolder. Root CLAUDE.md holds shared behavioral rules plus a one-paragraph
architecture and a pointer list. Folder-specific detail lives in the matching nested
CLAUDE.md.

Current nested files:
- .agents/CLAUDE.md - cross-agent shared config map (memory, rules, hooks, tools).
- .agents/memory/CLAUDE.md - memory store map.
- .agents/hooks/CLAUDE.md - hook inventory and event wiring.
- .claude/CLAUDE.md - Claude-native config map (settings, skills).

Why: every line in root CLAUDE.md is loaded into EVERY session. Folder-specific content
there is wasted tokens for sessions that do not touch that folder. Moving it to nested
CLAUDE.md keeps root small and only pays the token cost when the relevant folder is
edited.

Threshold rule of thumb: a folder earns its own CLAUDE.md at 3+ files or 1 file over
300 lines. Below that, inline a note in the parent.

How to apply when adding / moving / removing a tracked folder:
1. New folder with files: create <folder>/CLAUDE.md (one-line purpose, file list with
   one-line purpose each, folder-local invariants, a "When editing" section). Keep it
   30-100 lines. Do not duplicate root content.
2. Update the parent folder map table.
3. Add a pointer line to the root CLAUDE.md pointer list (pointer only, not detail).
4. On move/rename: move the CLAUDE.md with the folder and fix cross-references.
5. On remove: delete the nested CLAUDE.md and its pointer.

Anti-pattern: writing code in a new folder and leaving root CLAUDE.md untouched, or
pushing folder-specific detail back into root because it is easier to read in one place
(defeats the lazy-load token saving).
