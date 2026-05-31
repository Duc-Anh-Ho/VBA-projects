#!/usr/bin/env node
/**
 * @file ./.agents/hooks/commit/block-no-verify.mjs
 * @description block-no-verify - PreToolUse Bash hook. Splits the command into pipeline segments and DENIES only when a segment that actually invokes `git <verb>` also contains the --no-verify flag. Prevents bypassing other commit-time guardrails without false-positive on commands that merely mention the literal string (echo, grep, doc edit). Defensive: any internal error fails-open (allow) so a broken hook never blocks legitimate work.
 * @scope agent
 * @updated-at 2026-05-30
 */
import { readFileSync } from "fs";

const allow = () => {
  process.exit(0);
};

const deny = (reason) => {
  process.stdout.write(JSON.stringify({ hookSpecificOutput: { hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: reason } }) + "\n");
};

/**
 * Split a bash command into pipeline / list segments. Splits on |, ;, &&, ||, newline.
 * Does NOT respect quoting (literal --no-verify inside a quoted string of a non-git
 * segment is left alone because the segment will not look like `git <verb>`).
 *
 * @param {string} cmd
 * @returns {string[]}
 */
const segmentize = (cmd) => cmd.split(/\|\||&&|[|;&\n]/);

const NO_VERIFY = /\s--no-verify\b/;

const tokenize = (seg) => {
  const matches = seg.match(/(?:[^\s"']+|"[^"]*"|'[^']*')+/g) || [];
  return matches.map((t) => t.replace(/^["']|["']$/g, ""));
};

const isGitToken = (token) => {
  const normalized = token.replace(/\\/g, "/").toLowerCase();
  const base       = normalized.split("/").pop();
  return base === "git" || base === "git.exe";
};

const OPTION_VALUE_NEXT = new Set(["-C", "-c", "--git-dir", "--work-tree", "--namespace", "--super-prefix", "--config-env", "--exec-path"]);

const findGitVerb = (seg) => {
  const tokens = tokenize(seg).filter((t) => t !== "&");
  const gitIdx = tokens.findIndex(isGitToken);
  if (gitIdx === -1) return null;

  for (let i = gitIdx + 1; i < tokens.length; i++) {
    const token = tokens[i];
    if (OPTION_VALUE_NEXT.has(token)) {
      i++;
      continue;
    }
    if (token.startsWith("--") && token.includes("=")) continue;
    if (token.startsWith("-")) continue;
    return token.toLowerCase();
  }
  return null;
};

try {
  const input = JSON.parse(readFileSync(0, "utf8"));
  const cmd   = (input.tool_input?.command ?? "").trim();
  if (!cmd) { allow(); process.exit(0); }

  for (const seg of segmentize(cmd)) {
    if (findGitVerb(seg) && NO_VERIFY.test(seg)) {
      deny(
          `BLOCKED: --no-verify bypasses git hooks.\n`
        + `Fix the underlying error instead. If a hook is wrong, fix the hook in .agents/hooks/commit/, never bypass.\n`
        + `Allowed alternatives: address the hook complaint, or unstage the offending file, then retry without --no-verify.\n`
        + `Source rule: root-fix-only + git-commit-policy.md.\n`
        + `Offending segment: ${seg.trim().slice(0, 200)}`
      );
      process.exit(0);
    }
  }

  allow();
} catch {
  allow();
}
