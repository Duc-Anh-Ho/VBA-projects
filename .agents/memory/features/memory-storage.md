---
file: ./.agents/memory/features/memory-storage.md
name: Memory storage policy + creation protocol
description: memory-storage - storage path rule plus creation protocol with overlap-check announcement before writing a new memory file
type: feedback
scope: agent
updated-at: 2026-05-30
---

Storage path:

Always store memories in the project directory:
  D:\repos\VBA-projects\.agents\memory\

Never write memories to any other level:
- User level: %UserProfile%\.claude\CLAUDE.md or %UserProfile%\.claude\memory\
- App cache level: %UserProfile%\.claude\projects\<id>\memory\
- Any path outside the project repo

Why: only project-level memory is git-tracked and syncs across machines. Other levels
are silently out of sync and invisible during code review. The hook
check-write-rules.mjs blocks writes to the user-local memory dir, and
autoMemoryDirectory in settings.local.json redirects new memories here.

Creation protocol (announce the check result ALWAYS, even when no overlap):

Before creating any new memory file:
1. Read MEMORY.md and scan files in features/, behaviors/, project/, user/ for topic
   overlap.
2. Announce the check result in chat BEFORE writing:
   - Overlap found: state which file overlaps and merge into it instead of creating.
   - No overlap: state "Checked memory - no overlap with <topic>. Plan to create
     <path> with: <one-line summary>. Confirm to create?" and wait for confirmation.
3. For merges into an existing file: announce the merge target, then proceed.

Why announce: the user cannot see silent reads. The announcement is proof the
overlap-check happened. The hook also keyword-scores overlap and will deny a new file
when it overlaps an existing one, forcing a merge.

Scope extension - announce before modifying shared config:

The same announce-first protocol applies to non-memory shared config files:
.claude/settings.json, .claude/settings.local.json, .agents/hooks/*, .agents/rules/*.
List the file path and a one-line summary in chat before editing.
