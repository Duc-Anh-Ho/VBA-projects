#!/usr/bin/env node
/**
 * @file ./.agents/hooks/tools/block-unlisted-tool.mjs
 * @description PreToolUse hook for Bash/PowerShell. Enforces .agents/tools.local.json: external CLI tools must be invoked by their full path (the value stored in the `tools` map), bare names are denied with an actionable reason listing the expected full path and alternative options. Path-like primaries are allowed without lookup (project scripts, known binaries). Shell built-ins and untracked POSIX utilities fall back to the SKIP set. Defensive: when the manifest is missing or unreadable, the hook exits silently so the command proceeds.
 * @scope project
 * @updated-at 2026-05-30
 */
import { readFileSync } from "fs";

const MANIFEST_PATH = ".agents/tools.local.json";

// SKIP only contains tokens that CANNOT be tracked as external binaries:
// (1) Shell language keywords (if/then/for/while/case/function/return/...) - syntax tokens, no binary exists on any OS.
// (2) Pure shell intrinsics (cd/export/source/set/unset/eval/let/declare/...) - shell-state operations, intrinsic to shell process.
// (3) Bash-only operators with no external version on this machine ([[ and time, per /tools-init scan result).
// All other tools (rm, ls, cat, echo, printf, pwd, bash, sh, [, test, powershell, cmd, ...) MUST be in
// .agents/tools.local.json and invoked by full path. Run /tools-init to refresh paths.
// See .agents/tools.md (policy) and .agents/docs/hooks/pretooluse/tools-check.md (hook contract).
const SKIP = new Set([
  // (1) Shell language keywords
    "if", "then", "else", "elif", "fi"
  , "for", "while", "until", "do", "done"
  , "case", "esac", "in"
  , "function", "return", "exit", "break", "continue"
  // (2) Pure shell intrinsics
  , "cd", "export", "source", "."
  , "set", "unset", "shift", "read"
  , "eval", "exec", "trap", "wait", ":"
  , "hash", "alias", "unalias"
  , "local", "declare", "typeset", "readonly", "let"
  , "ulimit", "umask"
  // (3) Bash-only operators with no external version on this machine
  , "[[", "time"
]);

const input = JSON.parse(readFileSync(0, "utf8"));
if (input.tool_name !== "Bash" && input.tool_name !== "PowerShell") process.exit(0);

const command     = (input.tool_input?.command ?? "").trim();
const stripped    = command.replace(/^(\w+=\S*\s+)+/, "").trim();
const primaryTool = stripped.split(/[\s|&;(<]/)[0];

if (input.tool_name === "PowerShell" && /^[A-Z][a-zA-Z]+-[A-Z]/.test(primaryTool)) process.exit(0);
if (input.tool_name === "PowerShell" && primaryTool.startsWith("[")) process.exit(0);
if (input.tool_name === "PowerShell" && primaryTool.startsWith("$")) process.exit(0);
if (input.tool_name === "PowerShell" && primaryTool === "&") process.exit(0);

if (!primaryTool) process.exit(0);

/**
 * Emit the PreToolUse deny decision with `reason` and exit. Reason text is
 * shown back to the agent as the deny explanation.
 *
 * @param {string} reason Multi-line text including options for the agent to consider.
 */
const deny = (reason) => {
  process.stdout.write(JSON.stringify({
    hookSpecificOutput: {
        hookEventName            : "PreToolUse"
      , permissionDecision       : "deny"
      , permissionDecisionReason : reason
    }
  }));
  process.exit(0);
};

let manifest;
try {
  manifest = JSON.parse(readFileSync(MANIFEST_PATH, "utf8"));
} catch {
  process.exit(0);
}

const tools = manifest.tools ?? {};

if (primaryTool.includes("/") || primaryTool.includes("\\")) {
  process.exit(0);
}

if (primaryTool in tools) {
  const path = tools[primaryTool];
  if (path === null) {
    const installed = Object.entries(tools).filter(([, p]) => p !== null).slice(0, 8).map(([n]) => n).join(", ");
    deny(
        `Tool '${primaryTool}' is in .agents/tools.local.json but not installed (path is null).\n\n`
      + `Options:\n`
      + `  1. Install '${primaryTool}' on this machine, then run /tools-init to refresh the manifest path.\n`
      + `  2. Use an already-installed alternative from the manifest (examples: ${installed}).\n`
      + `STOP and report to the user. Do NOT substitute with another tool silently.`
    );
  }
  deny(
      `Bare tool name '${primaryTool}' is not allowed. Per .agents/tools.md usage policy, external CLI must be invoked by full path.\n\n`
    + `Use full path:\n`
    + `  ${path}\n\n`
    + `Or choose a different already-installed tool from .agents/tools.local.json if intentional.`
  );
}

if (SKIP.has(primaryTool)) process.exit(0);

const installed = Object.entries(tools).filter(([, p]) => p !== null).slice(0, 8).map(([n]) => n).join(", ");
deny(
    `Tool '${primaryTool}' is not in .agents/tools.local.json and is not a known shell built-in.\n\n`
  + `Options:\n`
  + `  1. Use an already-tracked alternative from the manifest (examples: ${installed}).\n`
  + `  2. Ask the user to install '${primaryTool}' on this machine, then add to the tool list in .claude/skills/tools-init/SKILL.md and run /tools-init.\n`
  + `STOP and report to the user. Do NOT substitute silently with a similar tool from training knowledge.`
);
