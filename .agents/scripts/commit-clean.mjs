#!/usr/bin/env node
/**
 * @file ./.agents/scripts/commit-clean.mjs
 * @description Create/update "deploy" branch with clean project files only.
 *   Strips all agent config (.agents/, .claude/, CLAUDE.md, etc.) and non-product
 *   folders (draft-and-reference/, malware-investigation/) from the project.
 *   Uses git worktree so the current branch is unaffected.
 *   Run via /commit-clean skill. Adapted from QA-agent template (Tier B).
 * @scope agent
 * @updated-at 2026-05-31
 */
import { execFileSync } from "child_process";
import {
    readFileSync, mkdtempSync, rmSync
  , readdirSync, existsSync
} from "fs";
import { tmpdir } from "os";
import { join } from "path";

const BRANCH = "deploy";

// Directories stripped from the deploy branch (agent config + non-product).
const EXCLUDE_DIRS = [
    ".agents"
  , ".claude"
  , ".codex"
  , ".gemini"
  , ".opencode"
  , "draft-and-reference"
  , "malware-investigation"
];

// Root files stripped from the deploy branch.
const EXCLUDE_ROOT_FILES = [
    "CLAUDE.md"
  , "AGENTS.md"
  , "AGENTS.override.md"
];

const projectDir = (process.env.CLAUDE_PROJECT_DIR || process.cwd()).replace(/\\/g, "/");

let gitPath = "git";
try {
  const m = JSON.parse(readFileSync(join(projectDir, ".agents", "tools.local.json"), "utf8"));
  if (m?.tools?.git) gitPath = m.tools.git;
} catch {}

const git = (args, cwd) => {
  const opts = { encoding: "utf8", stdio: ["pipe", "pipe", "pipe"], maxBuffer: 50 * 1024 * 1024 };
  if (cwd) opts.cwd = cwd;
  return execFileSync(gitPath, args, opts).trim();
};

const prefix = git(["rev-parse", "--show-prefix"]).replace(/\\/g, "/").replace(/\/$/, "");

/**
 * Recursively delete every file named `name` under `dir`.
 *
 * @param {string} dir
 * @param {string} name
 */
function deleteByName(dir, name) {
  if (!existsSync(dir)) return;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) deleteByName(full, name);
    else if (entry.name === name) rmSync(full);
  }
}

try {
  let branchExists = false;
  try { git(["rev-parse", "--verify", `refs/heads/${BRANCH}`]); branchExists = true; } catch {}

  const mainRef  = git(["symbolic-ref", "--short", "HEAD"]);
  const mainHash = git(["rev-parse", "--short", "HEAD"]);

  if (!branchExists) {
    git(["branch", BRANCH, mainRef]);
    console.log(`Created branch '${BRANCH}' from ${mainRef}`);
  }

  const wtPath = mkdtempSync(join(tmpdir(), "vba-deploy-")).replace(/\\/g, "/");
  console.log(`Worktree: ${wtPath}`);

  try {
    git(["worktree", "add", wtPath, BRANCH]);
    const projInWt = prefix ? join(wtPath, prefix).replace(/\\/g, "/") : wtPath;

    // Wipe the worktree (keep .git).
    for (const entry of readdirSync(wtPath)) {
      if (entry === ".git") continue;
      rmSync(join(wtPath, entry), { recursive: true, force: true });
    }

    // Restore files from the current branch.
    if (prefix) {
      git(["checkout", mainRef, "--", prefix + "/"], wtPath);
    } else {
      git(["checkout", mainRef, "--", "."], wtPath);
    }

    // Strip agent config and non-product dirs.
    for (const dir of EXCLUDE_DIRS) {
      const t = join(projInWt, dir);
      if (existsSync(t)) rmSync(t, { recursive: true, force: true });
    }

    // Strip root agent files.
    for (const f of EXCLUDE_ROOT_FILES) {
      const t = join(projInWt, f);
      if (existsSync(t)) rmSync(t, { force: true });
    }

    // Strip any nested CLAUDE.md (agent doc files in sub-directories).
    deleteByName(projInWt, "CLAUDE.md");

    git(["add", "-A"], wtPath);

    const status = git(["status", "--porcelain"], wtPath);
    if (!status) {
      console.log("No changes - deploy branch already matches current clean snapshot.");
    } else {
      const msg      = `Deploy: clean snapshot from ${mainRef}@${mainHash}`;
      git(["commit", "-m", msg], wtPath);
      const newHash  = git(["rev-parse", "--short", "HEAD"], wtPath);
      const fileCount = git(["ls-tree", "-r", "--name-only", "HEAD"], wtPath)
        .split("\n").filter(Boolean).length;
      console.log(`Committed: ${newHash} - ${msg}`);
      console.log(`Files in deploy: ${fileCount}`);
    }
  } finally {
    try { git(["worktree", "remove", "--force", wtPath]); }
    catch { rmSync(wtPath, { recursive: true, force: true }); }
  }

  console.log(`Done. Branch '${BRANCH}' updated locally. Not pushed.`);
  console.log(`To inspect: git log ${BRANCH} --oneline -5`);
  console.log(`To push:    git push origin ${BRANCH}`);
  console.log(`\nStripped: ${EXCLUDE_DIRS.join(", ")}, ${EXCLUDE_ROOT_FILES.join(", ")}, nested CLAUDE.md files`);

} catch (err) {
  console.error("FAILED:", err.message);
  if (err.stderr) console.error(err.stderr.toString().slice(0, 500));
  process.exit(1);
}
