---
file: ./.agents/tools.md
description: Index and usage policy for external CLI tools available to agents on this machine
scope: agent
updated-at: 2026-05-30
---

# Tools Index

## Usage policy

Only use tools that appear in `.agents/tools.local.json`. Do not use tools from training knowledge that are not in that manifest - they may not be installed on this machine. If a tool you need is missing, tell the user.

Full-path invocation is REQUIRED for every external CLI (anything in the `tools` map). Bare tool names are denied by `.agents/hooks/tools/block-unlisted-tool.mjs`; the deny message includes the full path to use plus alternatives. The `builtin` array lists Claude Code SDK tools that do not run through the shell and are always available - they bypass this check.

Note for this repo: prefer the dedicated Claude Code tools (Read, Glob, Grep, Edit, Write) over shelling out to cat/find/grep/sed. Use Bash/PowerShell only for things those tools cannot do (git, node, dotnet, etc.), and then invoke external binaries by their full path from the manifest.

## SKIP categories (only what cannot be tracked is skipped)

The hook SKIP set contains only tokens with no trackable external binary:
1. Shell language keywords - if, then, else, elif, fi, for, while, until, do, done, case, esac, in, function, return, exit, break, continue.
2. Pure shell intrinsics - cd, export, source, ., set, unset, shift, read, eval, exec, trap, wait, :, hash, alias, unalias, local, declare, typeset, readonly, let, ulimit, umask.
3. Bash-only operators with no external version on this machine - [[ and time.

Everything else with a binary (rm, ls, cat, echo, git, node, dotnet, powershell, ...) is in the manifest and MUST be invoked by full path.

## When you need a tool

1. Check `.agents/tools.local.json` `tools` map for the path on this machine.
2. If the value is `null` or the tool is absent: do not proceed - report it as missing.
3. Use the full path from the manifest when building a shell command.

## Machine-specific tool paths

`.agents/tools.local.json` holds full paths for THIS machine (gitignored). Regenerate it with the `/tools-init` skill after installing new tools or on a new machine. A `null` value means the tool is not installed.

## Built-in Claude Code tools

Server-side runtime tools (not shell binaries, no path check needed): Read, Write, Edit, Glob, Grep, Bash, PowerShell, WebFetch, WebSearch, Task* , Agent, ToolSearch, Skill, AskUserQuestion, Cron*, ScheduleWakeup, Monitor, Enter/ExitWorktree. MCP tools depend on `.claude/settings.json` server config.
