# .agents/memory/

Cross-agent memory store. MEMORY.md is the index (auto-injected each prompt by
hooks/memory/inject-index.mjs, and loaded via autoMemoryDirectory). Read MEMORY.md
first, then only the topic files you need.

## Folder map

| Subfolder | Holds |
|---|---|
| `user/` | Who the user is, preferences |
| `project/` | Ongoing work, goals, decisions not derivable from code/git |
| `features/` | How a project mechanism works (memory storage, nested CLAUDE.md) |
| `behaviors/` | Rules for how the agent should work (config location, commit policy) |

## When editing

- New memory file -> add a one-line pointer in MEMORY.md, and put the file in the right
  scope subfolder. Follow the overlap-check-and-announce protocol in
  features/memory-storage.md before creating.
- Each file: YAML frontmatter (name, description, type, scope, updated-at) then plain
  body. ASCII only, no bold/em-dash/arrows (doc-style hook enforces it).
- Writes to user-local memory paths are blocked by hooks/memory/check-write-rules.mjs.
