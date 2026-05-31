# Codex Project Instructions

Use project-local context only for this repository.

On session start, bootstrap from:
- `AGENTS.md`
- `CLAUDE.md`
- `.agents/memory/MEMORY.md`
- `.agents/shared-rules.md`

Do not read or write `C:\Users\Danh1\.codex` unless the user explicitly asks for
that exact user-level path in the current turn. This includes:
- `C:\Users\Danh1\.codex\memories`
- `C:\Users\Danh1\.codex\memories_*.sqlite`
- `C:\Users\Danh1\.codex\rules`
- `C:\Users\Danh1\.codex\config.toml`
- `C:\Users\Danh1\.codex\cap_sid`
- session logs, auth files, caches, and runtime state

All durable memories, rules, and agent config for this project belong in the repo:
- `.agents/memory/`
- `.agents/rules/`
- `.agents/hooks/`
- `.codex/`
- `.claude/`

If user-level Codex memory or rules are discovered because the user explicitly asks,
merge project-relevant memory/rule content into `.agents/`, verify the project copy,
then clear only the migrated user-level memory/rule content. Never touch auth,
secrets, credentials, or session logs without an explicit request.
