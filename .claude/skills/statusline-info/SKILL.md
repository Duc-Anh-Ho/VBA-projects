---
file: ./.claude/skills/statusline-info/SKILL.md
name: statusline-info
description: Explain every field shown in the project statusline (3-line layout) with concrete examples. Use when the user asks what a statusline field means, what ctx/5h/7d/dur/+N-N/style/think/effort/agent shows, or how the statusline is calculated.
disable-model-invocation: false
scope: project
updated-at: 2026-05-31
---

Explain the project statusline. The script lives at `.claude/scripts/statusline/statusline-command.mjs`
and produces up to 3 lines. Walk through every field with concrete examples.

Layout overview:

```
Claude Sonnet 4.6 | VBA-projects (dev) | style:default | v2.1.123
ctx:10% (100k/1M) | think:off          | effort:high   | +45/-12
5h:30% ↻22:40     | 7d:5% ↻Sat 17:00  | dur:12m30s    | ~$0.48
```

Line 1 - identity:

- Model name. From `model.display_name`. Examples: `Claude Opus 4.8`, `Claude Sonnet 4.6`.
- Directory plus git branch in parentheses. Examples: `VBA-projects (dev)`, `VBA-projects (main)`.
  If git fails, branch is hidden.
- Optional `session:NAME` when user set a custom session name with `--name` or `/rename`.
- Optional `agent:NAME` when running a subagent.
- `style:NAME` - current output style. Examples: `style:default`, `style:explanatory`.
- `vX.Y.Z` - Claude Code CLI version.

Line 2 - work state:

- `ctx:N% (Xk/Yk)` - context window usage. Computed from input + cache_creation + cache_read tokens.
  Warning zone above 80%, use /compact when needed. Example: `ctx:45% (450k/1000k)`.
- `200k+` - appears only when `exceeds_200k_tokens === true`.
- `think:on` or `think:off` - extended thinking state.
- `effort:LEVEL` - reasoning effort. Values: `low`, `medium`, `high`, `xhigh`, `max`.
- `+A/-R` - cumulative lines added and removed in the session from every Edit and Write.

Line 3 - limits and cost:

- `5h:N% ↻HH:MM` - 5-hour rolling rate limit. Hits 100% means throttled until reset.
- `7d:N% ↻Day HH:MM` - 7-day rolling rate limit. Pro/Max only.
- `dur:Xh Ym` or `dur:Xm Ys` - wall-clock session duration.
- `~$N.NN` - estimated session cost. The ~ marks it as estimate; actual billing is in Anthropic Console.

Field absence: every optional field auto-hides when its source is null or missing.

References:
- Source: `.claude/scripts/statusline/statusline-command.mjs`
- Schema: https://code.claude.com/docs/en/statusline
