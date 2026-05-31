---
name: memory-merge
description: Find Claude memories written to the user-local app cache, compare with project memory, and merge them into .agents/memory/ (git-tracked). Run after sessions where memories may have been saved outside the project.
allowed-tools: Read PowerShell(Get-ChildItem *) PowerShell(Remove-Item *) Write Edit
---

# /memory-merge

Compare user-local Claude memories with project memory and merge into the git-tracked
location.

## Why

By default Claude Code may write memories to a user-local path:
  %USERPROFILE%\.claude\projects\<encoded-project-path>\memory\

Project memory lives in one git-tracked location: `.agents/memory/`. The
autoMemoryDirectory in settings.local.json points new memories there. But a session run
without settings.local.json (for example on another machine) may leave memories
user-local. This skill merges them back.

## Safety rule

Before any destructive action (delete, overwrite), show the user exactly what will be
affected and wait for explicit confirmation. Never delete silently.

## Steps

### 0. Discover the user-local memory path

```powershell
Get-ChildItem "$env:USERPROFILE\.claude\projects\" -Directory -ErrorAction SilentlyContinue |
  Where-Object { Test-Path "$($_.FullName)\memory\" } |
  Select-Object -ExpandProperty FullName
```

Pick the directory matching this project (the path segment encodes the repo location:
`D--repos-VBA-projects`). If none is found, user-local memory is empty - skip to step 5.

### 1. Read both sides

List and read all .md files under the discovered user-local memory dir. Read
`.agents/memory/MEMORY.md` and the files under `.agents/memory/`.

### 2. Classify each user-local file

- already synced: exists in project with same or older content.
- needs merge: exists in project but user-local is newer.
- new: only user-local.

Destination: everything goes under `.agents/memory/<scope>/` (user, project, features,
behaviors). Show the diff and the proposed plan, then wait for confirmation.

### 3. Merge confirmed items

Write each approved item to the right scope folder. Update `.agents/memory/MEMORY.md`
with a one-line pointer for any new file. Follow features/memory-storage.md
(overlap-check first).

### 4. Delete the user-local copies (after confirmation)

Once merged and confirmed, remove the user-local .md files so the project stays the
single source of truth. Report how many were deleted.

### 5. Report

List what was merged, where it went, and how many user-local files were deleted.
