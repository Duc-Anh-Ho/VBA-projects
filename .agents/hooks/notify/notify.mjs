#!/usr/bin/env node
/**
 * @file ./.agents/hooks/notify/notify.mjs
 * @description Notification + Stop hook. Plays a PowerShell beep sequence and shows a fullscreen flash overlay to attract user attention. Profile is chosen per event: Notification (waiting for user input) plays a magenta double-beep; Stop (task complete) plays an ascending pentatonic with RGB color cycle. Uses execFileSync with array args (no shell parsing) and reads the pwsh full path from .agents/tools.local.json for cross-machine portability. Defensive: any failure inside the PowerShell call is swallowed so the hook never blocks Claude Code.
 * @scope project
 * @updated-at 2026-05-30
 */
import { readFileSync } from "fs";
import { execFileSync } from "child_process";

const projectRoot = (process.env.CLAUDE_PROJECT_DIR || process.cwd()).replace(/\\/g, "/");

let powerShellPath = "powershell";
try {
  const m = JSON.parse(readFileSync(`${projectRoot}/.agents/tools.local.json`, "utf8"));
  if (m?.tools?.pwsh) powerShellPath = m.tools.pwsh;
} catch {}

let input;
try { input = JSON.parse(readFileSync(0, "utf8")); } catch { process.exit(0); }

const eventName = input.hook_event_name || (input.matcher === "permission_prompt" ? "Notification" : null);

const PROFILES = {
  Notification: {
      beepFrequencies : [800, 800]
    , beepDurationMs  : 250
    , flashColors     : ["Magenta", "Magenta"]
    , flashOpacity    : 0.45
    , flashDurationMs : 140
    , flashGapMs      : 60
  },
  Stop: {
      beepFrequencies : [523, 659, 784, 1047, 1319]
    , beepDurationMs  : 180
    , flashColors     : ["Red", "Green", "Blue"]
    , flashOpacity    : 0.4
    , flashDurationMs : 180
    , flashGapMs      : 80
  }
};

const profile = PROFILES[eventName] || PROFILES.Stop;

const psScript = `
$freqs        = ${profile.beepFrequencies.join(", ")}
$beepDuration = ${profile.beepDurationMs}
$flashColors  = ${profile.flashColors.map(c => `'${c}'`).join(", ")}
$flashOpacity = ${profile.flashOpacity}
$flashDur     = ${profile.flashDurationMs}
$flashGap     = ${profile.flashGapMs}

foreach ($f in $freqs) { [console]::Beep($f, $beepDuration) }

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

function showFlash($colorName) {
  $form = New-Object System.Windows.Forms.Form
  $form.FormBorderStyle = 'None'
  $form.TopMost          = $true
  $form.ShowInTaskbar    = $false
  $form.BackColor        = [System.Drawing.Color]::FromName($colorName)
  $form.Opacity          = $flashOpacity
  $form.WindowState      = 'Maximized'

  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = $flashDur
  $timer.Add_Tick({ $form.Close(); $timer.Stop() })

  $form.Add_Shown({ $timer.Start() })
  [void]$form.ShowDialog()
  $form.Dispose()
  $timer.Dispose()
}

$count = [Math]::Max($freqs.Length, $flashColors.Length)
for ($i = 0; $i -lt $count; $i++) {
  showFlash $flashColors[$i % $flashColors.Length]
  if ($i -lt $count - 1) { Start-Sleep -Milliseconds $flashGap }
}
`;

const encoded = Buffer.from(psScript, "utf16le").toString("base64");
try {
  execFileSync(powerShellPath, ["-NoProfile", "-STA", "-EncodedCommand", encoded], { stdio: "ignore", timeout: 5000 });
} catch {}
