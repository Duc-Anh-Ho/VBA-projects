# .claude/

Claude-native config. Lazy-loaded when Claude touches files here. Cross-agent shared
config lives in .agents/.

## Top-level files

- `settings.json` - project settings, git-tracked. Team-shared: hook wiring,
  permissions baseline. No secrets.
- `settings.local.json` - machine-local, gitignored. Holds autoMemoryDirectory
  (points memory writes to .agents/memory) and any per-machine permission asks.

## Subfolder map

| Subfolder | Purpose |
|---|---|
| `skills/` | Claude-native skills (SKILL.md per skill). Run via /<name>. |

## settings.json vs settings.local.json

- autoMemoryDirectory MUST be in settings.local.json (Claude rejects this key in the
  git-tracked settings.json to stop shared repos redirecting memory writes).
- Hook command paths in settings.json are machine-portable (they use
  $CLAUDE_PROJECT_DIR). Per-machine tool full paths, if ever needed, go in
  settings.local.json.

## When editing

- New hook -> add it to settings.json hooks AND create the script under
  .agents/hooks/<event>/<name>.mjs. Document it in .agents/hooks/CLAUDE.md.
- New skill -> add skills/<name>/SKILL.md. Run via /<name>.
- Changing settings -> decide explicitly team-shared (settings.json) vs machine-local
  (settings.local.json), and announce the change per features/memory-storage.md.
