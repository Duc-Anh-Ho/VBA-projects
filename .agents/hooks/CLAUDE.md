# .agents/hooks/

Hook scripts (.mjs, Node). Each reads the hook JSON event on stdin and writes a JSON
decision on stdout. All are defensive: any internal error fails open (no block). Wiring
lives in .claude/settings.json; this folder holds the scripts.

## Inventory

| Script | Event | What it does |
|---|---|---|
| `memory/inject-index.mjs` | UserPromptSubmit | Injects .agents/memory/MEMORY.md into every prompt. |
| `memory/check-write-rules.mjs` | PreToolUse Edit/Write + UserPromptSubmit | Blocks writes to user-local memory; overlap-checks before a new memory file; warns if autoMemoryDirectory is wrong. |
| `tools/block-unlisted-tool.mjs` | PreToolUse Bash/PowerShell | Denies external CLI called by bare name or not in tools.local.json; tells you the full path to use. |
| `commit/inject-checklist.mjs` | PreToolUse Bash/PowerShell | On git commit, injects commit-checklist.md as a reminder. |
| `commit/block-message-format.mjs` | PreToolUse Bash/PowerShell | On git commit, blocks bad subject type-word / CRLF / missing blank line. |
| `commit/block-no-verify.mjs` | PreToolUse Bash/PowerShell | Denies git ... --no-verify. |
| `commit/block-without-trigger.mjs` | PreToolUse Bash/PowerShell | On git commit, surfaces Allow/Deny prompt - prevents silent carry-over from prior turn. |
| `commit/block-git-add-bypass.mjs` | PreToolUse Bash/PowerShell | Blocks `git add .` / `-A` catch-all; also blocks explicit add of machine-local files (settings.local.json, tools.local.json, .claude/cache/). |
| `edit/check-doc-style.mjs` | PreToolUse Edit/Write | Denies bold/em-dash/Unicode-arrow in .md under .agents/ and .claude/. |
| `edit/inject-vba-reminder.mjs` | PreToolUse Edit/Write | On .bas/.cls/.frm edits, injects the VBA conventions reminder. |
| `edit/inject-edit-track.mjs` | PostToolUse Edit/Write/MultiEdit/NotebookEdit | Emits an [EDIT-TRACK] +/- line delta per edit. |
| `notify/notify.mjs` | Notification + Stop | Beep + fullscreen flash when input is needed or a task finishes. |

## When editing

- Keep the .mjs extension. Read CLAUDE_PROJECT_DIR with a process.cwd() fallback.
- Read external binary paths (git, pwsh) from .agents/tools.local.json for portability.
- After adding/removing a hook, update both this table and the hooks block in
  .claude/settings.json.
