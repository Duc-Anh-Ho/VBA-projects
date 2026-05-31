/**
 * file: ./.claude/scripts/statusline/statusline-command.mjs
 * description: Claude Code 3-line status - identity, work state, limits and cost
 * scope: project
 * updated-at: 2026-05-31
 */
import { execSync } from "child_process";
import { writeFileSync, mkdirSync } from "fs";
import { basename, dirname, join } from "path";

let input = "";
process.stdin.on("data", chunk => (input += chunk));
process.stdin.on("end", () => {
  const data = JSON.parse(input);

  const model = data.model?.display_name ?? "Claude";
  const cwd = data.workspace?.current_dir ?? data.cwd ?? "";
  const dir = basename(cwd);
  const version = data.version;
  const agent = data.agent?.name;
  const style = data.output_style?.name;
  const sessionName = data.session_name;

  const ctxSize = data.context_window?.context_window_size;
  const usage = data.context_window?.current_usage;
  const ctxUsed = usage != null
    ? (usage.input_tokens ?? 0)
      + (usage.cache_creation_input_tokens ?? 0)
      + (usage.cache_read_input_tokens ?? 0)
    : null;
  const ctxPct = ctxUsed != null && ctxSize
    ? ctxUsed / ctxSize * 100
    : data.context_window?.used_percentage;
  const exceeds200k = data.exceeds_200k_tokens === true;

  const thinking = data.thinking?.enabled;
  const effort = data.effort?.level;
  const linesAdded = data.cost?.total_lines_added ?? 0;
  const linesRemoved = data.cost?.total_lines_removed ?? 0;

  const fiveHourPct = data.rate_limits?.five_hour?.used_percentage;
  const fiveHourResetsAt = data.rate_limits?.five_hour?.resets_at;
  const sevenDayPct = data.rate_limits?.seven_day?.used_percentage;
  const sevenDayResetsAt = data.rate_limits?.seven_day?.resets_at;

  try {
    const cachePath = join(cwd, ".claude", "cache", "rate-limits.json");
    mkdirSync(dirname(cachePath), { recursive: true });
    writeFileSync(cachePath, JSON.stringify({
      updated_at  : new Date().toISOString(),
      five_hour   : data.rate_limits?.five_hour ?? null,
      seven_day   : data.rate_limits?.seven_day ?? null,
    }, null, 2));
  } catch {}

  const durationMs = data.cost?.total_duration_ms;
  const cost = data.cost?.total_cost_usd;

  let branch = "";
  try {
    branch = execSync(
      `git -C "${cwd}" --no-optional-locks symbolic-ref --short HEAD`,
      { encoding: "utf8", stdio: ["pipe", "pipe", "ignore"], timeout: 1000 }
    ).trim();
  } catch {}

  const pad2 = n => n.toString().padStart(2, "0");
  const formatHM = epochSec => {
    const t = new Date(epochSec * 1000);
    return `${pad2(t.getHours())}:${pad2(t.getMinutes())}`;
  };
  const formatDayHM = epochSec => {
    const t = new Date(epochSec * 1000);
    const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    return `${days[t.getDay()]} ${pad2(t.getHours())}:${pad2(t.getMinutes())}`;
  };
  const formatDuration = ms => {
    const sec = Math.floor(ms / 1000);
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = sec % 60;
    if (h > 0) return `${h}h${m}m`;
    if (m > 0) return `${m}m${s}s`;
    return `${s}s`;
  };

  const line1 = [model, branch ? `${dir} (${branch})` : dir];
  if (sessionName) line1.push(`session:${sessionName}`);
  if (agent) line1.push(`agent:${agent}`);
  if (style) line1.push(`style:${style}`);
  if (version) line1.push(`v${version}`);

  const line2 = [];
  if (ctxPct != null) {
    let s = `ctx:${Math.round(ctxPct)}%`;
    if (ctxSize != null) {
      const maxK = Math.round(ctxSize / 1000);
      const usedK = ctxUsed != null
        ? Math.round(ctxUsed / 1000)
        : Math.round(ctxPct / 100 * maxK);
      s += ` (${usedK}k/${maxK}k)`;
    }
    line2.push(s);
  }
  if (exceeds200k) line2.push("200k+");
  if (thinking != null) line2.push(`think:${thinking ? "on" : "off"}`);
  if (effort) line2.push(`effort:${effort}`);
  if (linesAdded > 0 || linesRemoved > 0) {
    line2.push(`+${linesAdded}/-${linesRemoved}`);
  }

  const line3 = [];
  if (fiveHourPct != null) {
    let s = `5h:${Math.round(fiveHourPct)}%`;
    if (fiveHourResetsAt != null) s += ` ↻${formatHM(fiveHourResetsAt)}`;
    line3.push(s);
  }
  if (sevenDayPct != null) {
    let s = `7d:${Math.round(sevenDayPct)}%`;
    if (sevenDayResetsAt != null) s += ` ↻${formatDayHM(sevenDayResetsAt)}`;
    line3.push(s);
  }
  if (durationMs != null && durationMs > 0) {
    line3.push(`dur:${formatDuration(durationMs)}`);
  }
  if (cost != null) line3.push(`~$${cost.toFixed(2)}`);

  const lines = [line1, line2, line3].filter(l => l.length > 0);
  const gridRows = [line2, line3].filter(l => l.length > 0);
  const widthSource = gridRows.length > 0 ? gridRows : lines;
  const maxCols = Math.max(...widthSource.map(l => l.length));
  const colWidths = [];
  for (let i = 0; i < maxCols; i++) {
    colWidths[i] = Math.max(...widthSource.map(l => (l[i] ?? "").length));
  }
  const padded = lines.map(l =>
    l.map((seg, i) => i === l.length - 1 ? seg : seg.padEnd(colWidths[i] ?? 0)).join(" | ")
  );
  process.stdout.write(padded.join("\n"));
});
