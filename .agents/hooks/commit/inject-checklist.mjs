#!/usr/bin/env node
/**
 * @file ./.agents/hooks/commit/inject-checklist.mjs
 * @description PreToolUse Bash/PowerShell hook. When the command is `git commit ...`, allows the call and injects a checklist reminder as additionalContext. The checklist text is loaded from .agents/commit-checklist.md so each project adapts without touching this script. Fallback constant is used when the markdown file is missing or unreadable.
 * @scope agent
 * @updated-at 2026-05-30
 */
import {
    readFileSync
  , existsSync
} from "fs";
import { join } from "path";

const FALLBACK_CHECKLIST = `Commit checklist - confirm before proceeding:
- Tested the change manually before committing?
- Updated tracking / docs / changelog if a task is now done?
- Commit message explains WHY, not just what?`;

const projectRoot   = (process.env.CLAUDE_PROJECT_DIR || process.cwd()).replace(/\\/g, "/");
const checklistPath = join(projectRoot, ".agents", "commit-checklist.md");
const GIT_BIN       = String.raw`(?:&\s*)?(?:[A-Z_][A-Z0-9_]*=\S+\s+)*(?:"[^"]*[\\/](?:git|git\.exe)"|(?:[^\s"';&|]*[\\/])?(?:git|git\.exe))`;
const GIT_COMMIT_RE = new RegExp(String.raw`^\s*${GIT_BIN}\s+commit\b`, "i");

const segmentize = (cmd) => cmd.split(/\|\||&&|[|;&\n]/);
const isGitCommitCommand = (cmd) => segmentize(cmd).some((seg) => GIT_COMMIT_RE.test(seg.trim()));

/**
 * Read .agents/commit-checklist.md, strip the YAML frontmatter, return the
 * remaining body. Falls back to FALLBACK_CHECKLIST when the file is missing,
 * unreadable, or contains only frontmatter.
 *
 * @returns {string} The text to inject as additionalContext for git commit.
 */
const loadChecklist = () => {
  if (!existsSync(checklistPath)) return FALLBACK_CHECKLIST;
  try {
    const raw      = readFileSync(checklistPath, "utf8");
    const stripped = raw.replace(/^---[\s\S]*?---\s*/, "").trim();
    return stripped || FALLBACK_CHECKLIST;
  } catch {
    return FALLBACK_CHECKLIST;
  }
};

/**
 * Emit no JSON for a pass-through allow. When `ctx` is provided, attach it as
 * additionalContext without permissionDecision to avoid clients that reject
 * explicit permissionDecision: "allow".
 *
 * @param {string} [ctx] Optional reminder text injected with the allow.
 */
const allow = (ctx) => {
  if (!ctx) return;
  const out = { hookSpecificOutput: { hookEventName: "PreToolUse", additionalContext: ctx } };
  process.stdout.write(JSON.stringify(out) + "\n");
};

try {
  const input    = JSON.parse(readFileSync(0, "utf8"));
  const cmd      = (input.tool_input?.command ?? "").trim();
  const isCommit = isGitCommitCommand(cmd);
  if (!isCommit) { allow(); process.exit(0); }
  allow(loadChecklist());
} catch {
  allow();
}
