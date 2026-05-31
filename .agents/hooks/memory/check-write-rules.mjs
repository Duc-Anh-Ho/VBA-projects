#!/usr/bin/env node
/**
 * @file ./.agents/hooks/memory/check-write-rules.mjs
 * @description Memory write-rules merged hook. Dispatches on event shape: PreToolUse Edit/Write (blocks any write whose target is the user-local Claude/Codex memory directory, and asks for confirmation with keyword-overlap scoring before creating a brand-new memory file under .agents/memory) and UserPromptSubmit (warns when settings.local.json autoMemoryDirectory is missing or points outside the project memory directory). Defensive: any internal error fails-open (no-op).
 * @scope project
 * @updated-at 2026-05-30
 */
import {
    readFileSync
  , existsSync
  , readdirSync
} from "fs";
import {
    isAbsolute
  , dirname
  , basename
} from "path";

const TYPE_STRING       = "string";
const EVT_PRE           = "PreToolUse";
const SETTINGS_PATH     = ".claude/settings.local.json";
const EXPECTED_FRAGMENT = ".agents/memory";
const PROJECT_ROOT      = (process.env.CLAUDE_PROJECT_DIR || process.cwd()).replace(/\\/g, "/");
const PROJECT_MEMORY    = `${PROJECT_ROOT}/.agents/memory`;
const USER_LOCAL_RE     = /[\\/]\.(?:claude[\\/]projects[\\/][^\\/]+[\\/]memory|claude[\\/]memory|codex[\\/]memories)(?:[\\/]|$)|[\\/]\.codex[\\/]memories_[^\\/]+\.sqlite(?:-[a-z]+)?$/i;
const PROJECT_MEMORY_RE = /\.agents\/memory\/[^/]+\//i;

const STOPWORDS = new Set(
  ("a an the and or of for to in on at by with when if then that this is are be do does done " +
   "not no never always must should may can will would could where what who how which why " +
   "claude agent agents file files memory rule rules user write writes wrote read project")
  .split(/\s+/)
);

/**
 * Lowercase a string, drop non-alphanumeric chars, split on whitespace, and
 * keep tokens that are at least 3 characters long and not in STOPWORDS.
 *
 * @param {string} s
 * @returns {string[]} Keyword tokens used for memory-overlap scoring.
 */
const tokenize = (s) =>
  (s || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter(w => w.length >= 3 && !STOPWORDS.has(w));

/**
 * Pull the `description:` field from a memory file's YAML frontmatter.
 *
 * @param {string} content Raw file contents.
 * @returns {string} The description text, or "" when none was found.
 */
const extractDescription = (content) => {
  const m = content.match(/^description:\s*(.+)$/m);
  return m ? m[1].trim() : "";
};

/**
 * For every existing memory .md file in `dir` (excluding MEMORY.md), tokenize
 * its description, intersect with the tokenized newDesc, and score by overlap
 * count. Returns the list sorted by score descending.
 *
 * @param {string} dir Directory containing existing memory files.
 * @param {string} newDesc Description of the file being created.
 * @returns {Array<{ name:string, desc:string, overlap:string[], score:number }>}
 */
const scoreExistingMemory = (dir, newDesc) => {
  const newKw = new Set(tokenize(newDesc));
  try {
    return readdirSync(dir)
      .filter(f => f.endsWith(".md") && f !== "MEMORY.md")
      .map(f => {
        let desc = "";
        try {
          const c = readFileSync(`${dir}/${f}`, "utf8");
          desc    = extractDescription(c);
        } catch {}
        const oldKw   = new Set(tokenize(desc));
        const overlap = [...newKw].filter(w => oldKw.has(w));
        return { name: f, desc, overlap, score: overlap.length };
      })
      .sort((a, b) => b.score - a.score);
  } catch { return []; }
};

let input;
try { input = JSON.parse(readFileSync(0, "utf8")); } catch { process.exit(0); }

const isPreToolUse       = typeof input.tool_name === TYPE_STRING;
const isUserPromptSubmit = typeof input.prompt === TYPE_STRING;

if (isPreToolUse) {
  /**
   * Emit the PreToolUse allow decision and exit.
   */
  const allow = () => {
    process.exit(0);
  };
  /**
   * Emit the PreToolUse deny decision when a write targets user-local memory.
   *
   * @param {string} filePath The target file path.
   */
  const deny = (filePath) => {
    process.stdout.write(JSON.stringify({
      hookSpecificOutput: {
          hookEventName            : EVT_PRE
        , permissionDecision       : "deny"
        , permissionDecisionReason :
            `Memory write blocked: '${filePath}' is in a user-local Claude/Codex memory location. ` +
            `Project policy: all memory files MUST go to '${PROJECT_MEMORY}/' (see features/memory-storage.md). ` +
            `Re-issue the Write/Edit with file_path under .agents/memory/<user|project|features|behaviors>/ and update .agents/memory/MEMORY.md.`
      }
    }));
    process.exit(0);
  };

  if (input.tool_name !== "Edit" && input.tool_name !== "Write") allow();
  const filePath = (input.tool_input?.file_path ?? "").replace(/\\/g, "/");
  if (USER_LOCAL_RE.test(filePath)) deny(filePath);

  if (input.tool_name === "Write" && PROJECT_MEMORY_RE.test(filePath) && !existsSync(filePath)) {
    const dir        = dirname(filePath);
    const newContent = input.tool_input?.content ?? "";
    const newDesc    = extractDescription(newContent);
    const scored     = scoreExistingMemory(dir, newDesc);
    if (scored.length > 0) {
      const topOverlap   = scored.filter(s => s.score > 0).slice(0, 3);
      const allFiles     = scored.map(s => `  - ${s.name}: ${s.desc}`).join("\n");
      const overlapBlock = topOverlap.length > 0
        ? topOverlap.map(s =>
            `  - ${s.name}  [overlap=${s.score}: ${s.overlap.join(", ")}]\n    desc: ${s.desc}`
          ).join("\n")
        : "  (none - all existing files share no significant keywords)";
      process.stdout.write(JSON.stringify({
        hookSpecificOutput: {
            hookEventName            : EVT_PRE
          , permissionDecision       : topOverlap.length > 0 ? "deny" : "ask"
          , permissionDecisionReason :
              `MEMORY OVERLAP CHECK before creating '${basename(filePath)}' in '${dir}'.\n\n` +
              `New description: ${newDesc || "(missing - add description: line in frontmatter)"}\n\n` +
              `LIKELY OVERLAP (keyword-scored):\n${overlapBlock}\n\n` +
              (topOverlap.length > 0
                ? `BLOCKED: keyword overlap detected. You MUST merge into the existing file using Edit, not create a new file.\n` +
                  `Read the overlapping file(s) above, then Edit to merge your new content into the best match.`
                : `No keyword overlap found. Confirm this is a genuinely new topic before proceeding.\n` +
                  `All existing files in this dir:\n${allFiles}`)
        }
      }));
      process.exit(0);
    }
  }
  allow();
}

if (isUserPromptSubmit) {
  let dir = null;
  if (existsSync(SETTINGS_PATH)) {
    try {
      const settings = JSON.parse(readFileSync(SETTINGS_PATH, "utf8"));
      dir            = settings.autoMemoryDirectory ?? null;
    } catch {}
  }
  const isValid = dir && isAbsolute(dir) && dir.replace(/\\/g, "/").includes(EXPECTED_FRAGMENT);
  if (!isValid) {
    const current = dir ? `"${dir}"` : "not set";
    process.stdout.write(JSON.stringify({
      hookSpecificOutput: {
          hookEventName     : "UserPromptSubmit"
        , additionalContext :
            `MEMORY WARNING: autoMemoryDirectory in settings.local.json is ${current}. ` +
            `Must be an absolute path containing .agents/memory (e.g. "${PROJECT_MEMORY}"). ` +
            `Session is loading the default user-level MEMORY.md instead of the project one.`
      }
    }));
  }
  process.exit(0);
}

process.exit(0);
