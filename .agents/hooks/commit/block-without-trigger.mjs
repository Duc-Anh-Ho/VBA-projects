#!/usr/bin/env node
/**
 * @file ./.agents/hooks/commit/block-without-trigger.mjs
 * @description ask-on-commit - PreToolUse Bash/PowerShell hook on `git commit`. Surfaces a permission prompt to the user before any commit runs, requiring an explicit Allow click in the current turn. Defensive: any internal error fails-open (allow) so a broken hook never blocks legitimate work.
 * @scope agent
 * @updated-at 2026-05-31
 */
import { readFileSync } from "fs";

const HOOK_EVENT = "PreToolUse";

const GIT_BIN       = String.raw`(?:&\s*)?(?:[A-Z_][A-Z0-9_]*=\S+\s+)*(?:"[^"]*[\\/](?:git|git\.exe)"|(?:[^\s"';&|]*[\\/])?(?:git|git\.exe))`;
const GIT_COMMIT_RE = new RegExp(String.raw`(?:^|[\n;&|])\s*${GIT_BIN}\s+commit\b`, "i");

const allow = () => {
  process.exit(0);
};

const ask = (reason) => {
  const out = { hookSpecificOutput: { hookEventName: HOOK_EVENT, permissionDecision: "ask", permissionDecisionReason: reason } };
  process.stdout.write(JSON.stringify(out) + "\n");
};

/**
 * Split a shell command into pipeline / list segments.
 *
 * @param {string} cmd
 * @returns {string[]}
 */
const segmentize = (cmd) => cmd.split(/\|\||&&|[|;&\n]/);

/**
 * Detect whether the shell command actually invokes `git commit` in any segment.
 *
 * @param {string} cmd
 * @returns {boolean}
 */
const isCommitCommand = (cmd) => {
  if (!cmd) return false;
  for (const seg of segmentize(cmd)) {
    if (GIT_COMMIT_RE.test(seg)) return true;
  }
  return false;
};

try {
  const input   = JSON.parse(readFileSync(0, "utf8"));
  const command = input.tool_input?.command ?? "";
  if (!isCommitCommand(command)) { allow(); process.exit(0); }

  ask(
      `git commit detected - per-turn confirmation required (.agents/memory/behaviors/git-commit-policy.md).\n`
    + `Click Allow to run this commit, Deny to abort.\n`
    + `Per-commit confirmation prevents prior-turn permission from carrying over silently.`
  );
} catch {
  allow();
}
