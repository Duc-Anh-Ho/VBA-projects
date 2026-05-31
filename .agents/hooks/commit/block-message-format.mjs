#!/usr/bin/env node
/**
 * @file ./.agents/hooks/commit/block-message-format.mjs
 * @description PreToolUse Bash/PowerShell hook on `git commit`. BLOCKS when: (1) subject
 *   does not start with a recognized type word, (2) line endings contain CR/CRLF, (3) body
 *   is present but not separated by one blank line, (4) staged files mix agent-config scope
 *   (.agents/, .claude/, CLAUDE.md) with project scope (VBA files, docs, etc.) - required so
 *   /commit-clean can produce a clean deploy branch. Defensive: any internal error fails-open.
 * @scope agent
 * @updated-at 2026-05-31
 */
import { readFileSync } from "fs";
import { execFileSync } from "child_process";

const HOOK_EVENT = "PreToolUse";

// Project commit: "Add: description"   Agent commit: "Add (agent): description"
const SUBJECT_FORMAT       = /^(Add|Update|Fix|Refactor|Remove|Docs|Chore|Style|Test|Perf|Build|CI)(\s+\(agent\))?\s*:.+/i;
const SUBJECT_HAS_AGENT    = /^(Add|Update|Fix|Refactor|Remove|Docs|Chore|Style|Test|Perf|Build|CI)\s+\(agent\)\s*:/i;
const SUBJECT_NO_AGENT     = /^(Add|Update|Fix|Refactor|Remove|Docs|Chore|Style|Test|Perf|Build|CI)\s*:/i;

// Agent-config scope: files that belong ONLY in agent commits.
const AGENT_PREFIXES   = [".agents/", ".claude/"];
const AGENT_ROOT_FILES = new Set(["CLAUDE.md", "AGENTS.md", "AGENTS.override.md"]);

// Project scope: VBA source, workbooks, docs, images, tools.
const PROJECT_PREFIXES = [
    "final-installation/"
  , "xlwings/"
  , "part-tools/"
  , "draft-and-reference/"
  , "malware-investigation/"
  , "Images/"
  , "docs/"
  , "README.md"
  , "Danh-Tools-Installation.xlsb"
];

const allow = () => process.exit(0);

const deny = (reason) => {
  process.stdout.write(JSON.stringify({ hookSpecificOutput: { hookEventName: HOOK_EVENT, permissionDecision: "deny", permissionDecisionReason: reason } }) + "\n");
  process.exit(0);
};

const GIT_BIN       = String.raw`(?:&\s*)?(?:[A-Z_][A-Z0-9_]*=\S+\s+)*(?:"[^"]*[\\/](?:git|git\.exe)"|(?:[^\s"';&|]*[\\/])?(?:git|git\.exe))`;
const GIT_COMMIT_RE = new RegExp(String.raw`^\s*${GIT_BIN}\s+commit\b`, "i");
const segmentize    = (cmd) => cmd.split(/\|\||&&|[|;&\n]/);

const findGitCommitCommand = (cmd) => {
  for (const seg of segmentize(cmd)) {
    if (GIT_COMMIT_RE.test(seg.trim())) return seg.trim();
  }
  return null;
};

const parseCommitMessage = (cmd) => {
  const heredoc = cmd.match(/-m\s+["']\$\(\s*cat\s+<<\s*['"]?EOF['"]?\s*\r?\n([\s\S]*?)\r?\nEOF/);
  if (heredoc) return heredoc[1];
  const simple = cmd.match(/-m\s+(["'])((?:\\.|(?!\1).)+?)\1/);
  if (simple) return simple[2];
  return null;
};

/**
 * Classify a staged file path as "agent", "project", or "other".
 *
 * @param {string} filePath
 * @returns {"agent"|"project"|"other"}
 */
const scopeOf = (filePath) => {
  const norm = filePath.replace(/\\/g, "/");
  if (AGENT_ROOT_FILES.has(norm)) return "agent";
  for (const p of AGENT_PREFIXES)   if (norm.startsWith(p)) return "agent";
  for (const p of PROJECT_PREFIXES) if (norm.startsWith(p)) return "project";
  return "other";
};

const projectRoot = (process.env.CLAUDE_PROJECT_DIR || process.cwd()).replace(/\\/g, "/");

let gitPath = "git";
try {
  const m = JSON.parse(readFileSync(`${projectRoot}/.agents/tools.local.json`, "utf8"));
  if (m?.tools?.git) gitPath = m.tools.git;
} catch {}

/**
 * Run git and return stdout. Returns "" on failure.
 *
 * @param {string[]} args
 * @returns {string}
 */
const git = (args) => {
  try {
    return execFileSync(gitPath, ["-C", projectRoot, ...args], {
      encoding: "utf8", stdio: ["ignore", "pipe", "pipe"]
    });
  } catch { return ""; }
};

const listStagedFiles = () => {
  const raw = git(["diff", "--cached", "--name-only", "--diff-filter=ACMR"]).trim();
  return raw ? raw.split("\n").map(s => s.trim()).filter(Boolean) : [];
};

try {
  const input     = JSON.parse(readFileSync(0, "utf8"));
  const cmd       = (input.tool_input?.command ?? "").trim();
  const commitCmd = findGitCommitCommand(cmd);
  if (!commitCmd) allow();

  const violations = [];
  const staged     = listStagedFiles();

  // -- Scope split + (agent) marker check --
  const agentFiles   = staged.filter(f => scopeOf(f) === "agent");
  const projectFiles = staged.filter(f => scopeOf(f) === "project");

  if (agentFiles.length > 0 && projectFiles.length > 0) {
    violations.push(
        `staged files mix agent-config scope and project scope - required by /commit-clean.\n`
      + `  Agent files (${agentFiles.length}): ${agentFiles.slice(0, 4).join(", ")}${agentFiles.length > 4 ? "..." : ""}\n`
      + `  Project files (${projectFiles.length}): ${projectFiles.slice(0, 4).join(", ")}${projectFiles.length > 4 ? "..." : ""}\n`
      + `  Fix: unstage one scope and commit separately.\n`
      + `  Agent commit:   git add .agents/ .claude/ CLAUDE.md  ->  "Add (agent): description"\n`
      + `  Project commit: git add final-installation/ xlwings/  ->  "Add: description"`
    );
  }

  // -- (agent) marker + message format checks --
  const message = parseCommitMessage(commitCmd);
  if (message !== null) {
    if (/\r/.test(message)) {
      violations.push(`line endings must be LF only (no CR); use HEREDOC (cat <<'EOF') on Windows`);
    }
    const trimmed = message.replace(/\r/g, "").trim();
    const lines   = trimmed.split("\n");
    const subject = lines[0] || "";

    if (!SUBJECT_FORMAT.test(subject)) {
      violations.push(
          `subject "${subject.slice(0, 80)}" must start with a type word: `
        + `Add, Update, Fix, Refactor, Remove, Docs, Chore, Style, Test, Perf, Build, CI`
      );
    }
    if (lines.length >= 2 && lines[1].trim() !== "") {
      violations.push(
          `subject and body must be separated by one blank line; line 2 is non-empty: "${lines[1].slice(0, 60)}"`
      );
    }

    // (agent) marker enforcement (only meaningful when staged files are known).
    if (staged.length > 0) {
      const hasAgentMarker = SUBJECT_HAS_AGENT.test(subject);
      if (agentFiles.length > 0 && projectFiles.length === 0 && !hasAgentMarker) {
        violations.push(
            `agent-only commit must use "(agent)" in subject.\n`
          + `  Current:  "${subject.slice(0, 80)}"\n`
          + `  Expected: "Add (agent): description"  /  "Update (agent): description"  etc.`
        );
      }
      if (projectFiles.length > 0 && agentFiles.length === 0 && hasAgentMarker) {
        violations.push(
            `project-only commit must NOT have "(agent)" in subject.\n`
          + `  Current:  "${subject.slice(0, 80)}"\n`
          + `  Expected: "Add: description"  /  "Update: description"  etc.`
        );
      }
    }
  }

  if (violations.length === 0) allow();

  deny(
      `COMMIT BLOCKED: ${violations.length} violation(s).\n`
    + violations.map((v, i) => `  ${i + 1}. ${v}`).join("\n\n")
    + `\n\nRule: .agents/memory/behaviors/git-commit-policy.md`
  );
} catch {
  allow();
}
