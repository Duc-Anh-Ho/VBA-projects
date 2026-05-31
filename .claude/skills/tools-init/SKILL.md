---
name: tools-init
description: Discover all available CLI tools on this machine and write their full paths to .agents/tools.local.json. Run once per machine or after installing new tools.
disable-model-invocation: true
allowed-tools: Bash(where *) Bash(which *) PowerShell(Get-Command *)
---

# /tools-init

Scan this machine for the known tool list, write a JSON manifest to
`.agents/tools.local.json`, then print a report. The manifest gives every hook and
every shell command a verified full path (the block-unlisted-tool hook depends on it).

## Steps

### 1. Detect OS

- Windows: `where <tool>` (or PowerShell `(Get-Command <tool>).Source`).
- Linux/WSL: `which <tool>`.

### 2. Discover tools

If `.agents/tools.local.json` already exists: read it, then for every tool re-scan a
`null` path (it may have been installed since) and verify a non-null path still exists
on disk (set to `null` if it no longer resolves).

Tool list to scan (group in parentheses):

unix: cp, diff, du, grep, sed, awk, wget, head, tail, wc, cut, xargs, uniq, unzip, tar,
vim, nano, ssh, scp, sftp, ls, cat, mkdir, rmdir, rm, mv, touch, chmod, chown, ln,
basename, dirname, env, printenv, which, tr, sort, paste, tee, find, date, sleep, ps,
tty, file, zip, gzip, gunzip, more, less, expr, bc, echo, printf, pwd, kill, true,
false, type, [, test, bash, sh, dash, zsh

dev: git, node, npm, npx, nvm, python, python3, py, rg, curl, jq, gh, fd, bat, yq,
pnpm, mise, java, javac, ffmpeg, code

devops: docker, ollama, kubectl

dotnet: dotnet

agent-cli: claude, opencode, codex, gemini, cn, kilo

windows: clip, ipconfig, pwsh, reg, robocopy, choco, scoop, systeminfo, taskkill,
tasklist, where, winget, wmic, wsl, powershell, cmd

db: sqlcmd, bcp, psql, sqlplus, sqlite3, mysql, mongosh, redis-cli

### 3. Write `.agents/tools.local.json`

```json
{
  "generated": "YYYY-MM-DD",
  "os": "win32|linux|wsl",
  "tools": { "<tool>": "<full path or null>" }
}
```

Keep the existing `builtin` array if present (the Claude Code SDK tools that bypass the
path check). A `null` value means the tool is not installed - do not assume it exists.

### 4. Ensure it is gitignored

`.agents/tools.local.json` is machine-specific and must stay in `.gitignore`. If the
entry is missing, add it.

### 5. Print report

- Found: list tools with non-null paths and their paths.
- Missing: list tools with a null path (tell the user the install source only if you
  are certain; do not invent install commands).
