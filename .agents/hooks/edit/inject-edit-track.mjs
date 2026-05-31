#!/usr/bin/env node
/**
 * @file ./.agents/hooks/edit/inject-edit-track.mjs
 * @description PostToolUse hook for Edit/Write/MultiEdit/NotebookEdit. Computes per-turn line delta directly from tool_input (Edit/MultiEdit: old_string vs new_string; Write: new content vs HEAD content; NotebookEdit: old_source vs new_source) and emits an [EDIT-TRACK] additionalContext reminder so the agent sees each edit in real time.
 * @scope agent
 * @updated-at 2026-05-30
 */
import { readFileSync } from "fs";
import { execFileSync } from "child_process";

/**
 * Emit the PostToolUse allow decision. Optional `ctx` becomes additionalContext.
 *
 * @param {string} [ctx] Reminder text injected with the allow.
 */
const allow = (ctx) => {
  const out = { hookSpecificOutput: { hookEventName: "PostToolUse" } };
  if (ctx) out.hookSpecificOutput.additionalContext = ctx;
  process.stdout.write(JSON.stringify(out) + "\n");
};

const projectRoot = (process.env.CLAUDE_PROJECT_DIR || process.cwd()).replace(/\\/g, "/");

const PROJECT_PREFIXES = ["final-installation/", "part-tools/", "xlwings/", "draft-and-reference/"];
const AGENT_PREFIXES   = [".agents/", ".claude/", ".codex/", ".opencode/", ".gemini/", "docs/"];
const AGENT_ROOT_FILES = new Set(["AGENTS.md", "AGENTS.override.md", "CLAUDE.md", "GEMINI.md", "OPENCODE.md", "opencode.jsonc"]);
const TRACKED          = new Set(["Edit", "Write", "MultiEdit", "NotebookEdit"]);

/**
 * Classify a repo-relative file path into "project", "agent", or "other"
 * for the (scope) tag in the [EDIT-TRACK] line.
 *
 * @param {string} rel
 * @returns {"project"|"agent"|"other"}
 */
const scopeOf = (rel) => {
  if (AGENT_ROOT_FILES.has(rel)) return "agent";
  for (const p of AGENT_PREFIXES)   if (rel.startsWith(p)) return "agent";
  for (const p of PROJECT_PREFIXES) if (rel.startsWith(p)) return "project";
  return "other";
};

let gitPath = "git";
try {
  const m = JSON.parse(readFileSync(`${projectRoot}/.agents/tools.local.json`, "utf8"));
  if (m?.tools?.git) gitPath = m.tools.git;
} catch {}

/**
 * Run a git command rooted at projectRoot. Returns "" on failure.
 *
 * @param {string[]} args Args after `git -C projectRoot`.
 * @returns {string} stdout, or "" on failure.
 */
const git = (args) => {
  try { return execFileSync(gitPath, ["-C", projectRoot, ...args], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] }); }
  catch { return ""; }
};

const subdir = git(["rev-parse", "--show-prefix"]).replace(/\r?\n/g, "");

/**
 * Count the lines in a string. Treats a missing or empty string as 0 lines.
 * A trailing newline is not double-counted.
 *
 * @param {string} s
 * @returns {number}
 */
const numLines = (s) => {
  if (!s) return 0;
  const n = (s.match(/\n/g) || []).length;
  return s.endsWith("\n") ? n : n + 1;
};

/**
 * Compute the per-tool-call line delta from tool_input. For Write, the prior
 * line count is fetched from HEAD via `git show HEAD:<subdir><rel>` so untracked
 * files correctly report minus=0.
 *
 * @param {string} tool One of Edit | MultiEdit | Write | NotebookEdit.
 * @param {object} input The tool_input payload.
 * @param {string} rel Repo-relative path of the file.
 * @returns {{ plus:number, minus:number }}
 */
const lineDelta = (tool, input, rel) => {
  if (tool === "Edit") {
    return { plus: numLines(input?.new_string || ""), minus: numLines(input?.old_string || "") };
  }
  if (tool === "MultiEdit") {
    let plus = 0, minus = 0;
    for (const e of (input?.edits || [])) {
      plus  += numLines(e?.new_string || "");
      minus += numLines(e?.old_string || "");
    }
    return { plus, minus };
  }
  if (tool === "Write") {
    const plus  = numLines(input?.content || "");
    const prior = git(["show", `HEAD:${subdir}${rel}`]);
    return { plus, minus: numLines(prior) };
  }
  if (tool === "NotebookEdit") {
    return { plus: numLines(input?.new_source || ""), minus: numLines(input?.old_source || "") };
  }
  return { plus: 0, minus: 0 };
};

try {
  const input = JSON.parse(readFileSync(0, "utf8"));
  const tool  = input.tool_name;
  if (!TRACKED.has(tool)) { allow(); process.exit(0); }

  const filePath = input.tool_input?.file_path;
  if (!filePath) { allow(); process.exit(0); }

  const normalized = filePath.replace(/\\/g, "/");
  const rel        = normalized.startsWith(projectRoot + "/") ? normalized.slice(projectRoot.length + 1) : normalized;
  const base       = rel.split("/").pop();
  const scope      = scopeOf(rel);
  const { plus, minus } = lineDelta(tool, input.tool_input, rel);

  allow(`[EDIT-TRACK] (${scope}) ${base} (+${plus}/-${minus})`);
} catch {
  allow();
}
