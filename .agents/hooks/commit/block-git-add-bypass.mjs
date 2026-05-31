#!/usr/bin/env node
/**
 * @file ./.agents/hooks/commit/block-git-add-bypass.mjs
 * @description block-git-add-bypass - PreToolUse Bash/PowerShell hook. Enforces git-commit-policy.md at git add time. Denies (a) catch-all args (., -A, -u, --all, --update) that would silently include skip-list files, and (b) explicit paths inside the skip list (.claude/settings.local.json, .agents/tools.local.json, .claude/cache/). NOTE: binary workbook files (.xlsb, .xlam, .frx) are intentionally NOT in the skip list - they are the product and must be committed explicitly. Defensive: any internal error fails-open (allow).
 * @scope agent
 * @updated-at 2026-05-31
 */
import { readFileSync } from "fs";

const allow = () => {
  process.exit(0);
};

const deny = (reason) => {
  process.stdout.write(JSON.stringify({ hookSpecificOutput: { hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: reason } }) + "\n");
};

// Files that must NEVER be committed - machine-local or sensitive only.
// Binary workbooks (.xlsb, .xlam, .frx) are NOT here - they ARE the product.
const SKIP_PREFIXES = [
    ".claude/settings.local.json"
  , ".agents/tools.local.json"
  , ".claude/cache/"
];

const CATCHALL_ARGS = new Set([".", "-A", "--all", "-u", "--update", "*"]);

const segmentize = (cmd) => cmd.split(/\|\||&&|[|;&\n]/);

/**
 * Tokenize a shell segment, respecting double and single quotes.
 *
 * @param {string} seg
 * @returns {string[]}
 */
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

const findGitVerbIndex = (tokens) => {
  const gitIdx = tokens.findIndex(isGitToken);
  if (gitIdx === -1) return -1;
  for (let i = gitIdx + 1; i < tokens.length; i++) {
    const token = tokens[i];
    if (OPTION_VALUE_NEXT.has(token)) { i++; continue; }
    if (token.startsWith("--") && token.includes("=")) continue;
    if (token.startsWith("-")) continue;
    return i;
  }
  return -1;
};

const normalize = (p) => p.replace(/\\/g, "/").replace(/^\.\//, "");

try {
  const input = JSON.parse(readFileSync(0, "utf8"));
  const cmd   = (input.tool_input?.command ?? "").trim();
  if (!cmd) { allow(); process.exit(0); }

  for (const seg of segmentize(cmd)) {
    const trimmed = seg.trim();
    const tokens  = tokenize(trimmed).filter((t) => t !== "&");
    const addIdx  = findGitVerbIndex(tokens);
    if (addIdx === -1 || tokens[addIdx].toLowerCase() !== "add") continue;
    const args = tokens.slice(addIdx + 1);

    for (const arg of args) {
      if (CATCHALL_ARGS.has(arg)) {
        deny(
            `BLOCKED: 'git add ${arg}' is a catch-all that can silently stage machine-local files.\n`
          + `Skip list: .claude/settings.local.json, .agents/tools.local.json, .claude/cache/.\n`
          + `Use explicit file names instead. Examples:\n`
          + `  git add final-installation/VBA-files-Danh-Tools-Installation/Modules/CustomUi.bas\n`
          + `  git add final-installation/Danh-Tools-Installation.xlsb\n`
          + `  git add .agents/memory/MEMORY.md\n`
          + `Source rule: .agents/memory/behaviors/git-commit-policy.md`
        );
        process.exit(0);
      }

      const norm = normalize(arg);
      for (const skip of SKIP_PREFIXES) {
        if (norm === skip || norm.startsWith(skip)) {
          deny(
              `BLOCKED: '${arg}' is a machine-local file and must NEVER be committed.\n`
            + `Reason: machine-local config (gitignored, contains machine-specific paths or autoMemoryDirectory).\n`
            + `Edit the file locally only - never stage it.\n`
            + `Offending segment: ${trimmed.slice(0, 200)}`
          );
          process.exit(0);
        }
      }
    }
  }

  allow();
} catch {
  allow();
}
