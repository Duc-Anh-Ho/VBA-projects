#!/usr/bin/env node
/**
 * @file ./.agents/hooks/edit/check-doc-style.mjs
 * @description PreToolUse Edit/Write hook. DENIES writes that introduce markdown bold, Unicode arrows, or em-dash in .md files under the agent-config dirs (.agents/, .claude/). Human-facing docs under docs/ are intentionally NOT watched so plan/report files may use rich markdown. Defensive: any internal error fails-open (allow).
 * @scope project
 * @updated-at 2026-05-30
 */
import { readFileSync } from "fs";

const EVT_PRE = "PreToolUse";

// Narrowed for VBA-projects: only agent-config docs are plain-text-enforced.
// docs/ (human-facing plans/reports) is deliberately excluded so they may use bold/em-dash.
const DOC_STYLE_DIRS = [".agents/", ".claude/"];

/**
 * Emit the PreToolUse allow decision. Optional `ctx` attaches an advisory
 * additionalContext (never blocking).
 *
 * @param {string|null} [ctx]
 */
const allow = (ctx) => {
  if (!ctx) return;
  const out = { hookSpecificOutput: { hookEventName: EVT_PRE, additionalContext: ctx } };
  process.stdout.write(JSON.stringify(out) + "\n");
};

/**
 * Emit the PreToolUse deny decision with the supplied reason.
 *
 * @param {string} reason
 */
const deny = (reason) => {
  const out = { hookSpecificOutput: { hookEventName: EVT_PRE, permissionDecision: "deny", permissionDecisionReason: reason } };
  process.stdout.write(JSON.stringify(out) + "\n");
};

try {
  const input    = JSON.parse(readFileSync(0, "utf8"));
  const filePath = (input.tool_input?.file_path ?? "").replace(/\\/g, "/");
  const isDocFile = DOC_STYLE_DIRS.some(d => filePath.includes(d)) && filePath.endsWith(".md");
  if (!isDocFile) { allow(); process.exit(0); }

  const raw      = input.tool_name === "Write"
    ? (input.tool_input?.content ?? "")
    : (input.tool_input?.new_string ?? "");
  const stripped = raw
    .replace(/```[\s\S]*?```/g, "")
    .replace(/`[^`\n]+`/g, "");
  const hasBold         = /\*\*[^*]+\*\*/.test(stripped);
  const hasUnicodeArrow = /[←→⇐⇒]/.test(stripped);
  const hasEmDash       = /—/.test(stripped);
  if (hasBold || hasUnicodeArrow || hasEmDash) {
    const violations = [];
    if (hasBold)         violations.push("bold (double-asterisk) - use plain text instead");
    if (hasUnicodeArrow) violations.push("Unicode arrow (U+2192 or similar) - use -> instead");
    if (hasEmDash)       violations.push("em-dash (U+2014) - use \" - \" (space-hyphen-space) instead");
    deny(
        `DOC STYLE VIOLATION: ${violations.join("; ")}. `
      + `Rule: .agents/rules/doc/doc-style.md. Plain text only in .md files under .agents/, .claude/. Fix before retrying.`
    );
    process.exit(0);
  }

  allow();
} catch {
  allow();
}
