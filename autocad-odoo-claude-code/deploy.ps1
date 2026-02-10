# deploy.ps1 - Deploy autocad-odoo-claude-code framework
# Usage: .\deploy.ps1 [-Target <path>] [-Backup] [-DryRun] [-CleanHome]

param(
    [string]$Target = "",
    [switch]$Backup,
    [switch]$DryRun,
    [switch]$CleanHome
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

# Default target = project .claude/
if (-not $Target) {
    $Target = Join-Path $ProjectRoot ".claude"
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  AutoCAD-Odoo Claude Code Framework Deploy" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Source:  $ScriptDir" -ForegroundColor Yellow
Write-Host "Target:  $Target" -ForegroundColor Yellow
Write-Host "Project: $ProjectRoot" -ForegroundColor Yellow
Write-Host ""

# --- CleanHome: remove from HOME ~/.claude/ ---
if ($CleanHome) {
    $homeClaude = Join-Path $env:USERPROFILE ".claude"
    $dirsToClean = @("agents", "skills", "rules", "templates", "scripts")
    Write-Host "[CleanHome] Removing framework dirs from $homeClaude" -ForegroundColor Magenta
    foreach ($d in $dirsToClean) {
        $p = Join-Path $homeClaude $d
        if (Test-Path $p) {
            Write-Host "  Removing $p" -ForegroundColor DarkMagenta
            if (-not $DryRun) { Remove-Item -Path $p -Recurse -Force }
        }
    }
    # commands: only remove our files, keep others (BMad etc)
    $ourCmds = @("tdd.md", "code-review.md", "git-push.md", "plan.md")
    $cmdDir = Join-Path $homeClaude "commands"
    if (Test-Path $cmdDir) {
        foreach ($f in $ourCmds) {
            $fp = Join-Path $cmdDir $f
            if (Test-Path $fp) {
                Write-Host "  Removing $fp" -ForegroundColor DarkMagenta
                if (-not $DryRun) { Remove-Item -Path $fp -Force }
            }
        }
    }
    Write-Host "[CleanHome] Done" -ForegroundColor Magenta
    Write-Host ""
}

# --- Backup ---
if ($Backup -and (Test-Path $Target)) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupDir = "${Target}_backup_${timestamp}"
    Write-Host "[Backup] Creating backup at $backupDir" -ForegroundColor Green
    if (-not $DryRun) {
        Copy-Item -Path $Target -Destination $backupDir -Recurse -Force
    }
}

# --- Ensure target exists ---
if (-not (Test-Path $Target)) {
    Write-Host "[Init] Creating $Target" -ForegroundColor Green
    if (-not $DryRun) {
        New-Item -ItemType Directory -Path $Target -Force | Out-Null
    }
}

# --- Deploy function ---
function Deploy-Directory {
    param(
        [string]$Source,
        [string]$Dest,
        [string]$Label
    )
    if (Test-Path $Source) {
        Write-Host "[Deploy] $Label -> $Dest" -ForegroundColor Green
        if (-not $DryRun) {
            if (-not (Test-Path $Dest)) {
                New-Item -ItemType Directory -Path $Dest -Force | Out-Null
            }
            Copy-Item -Path "$Source\*" -Destination $Dest -Recurse -Force
        }
    } else {
        Write-Host "[Skip] $Label not found at $Source" -ForegroundColor Yellow
    }
}

# --- Deploy directories (merge into existing) ---
$deployItems = @(
    @{ Source = "skills";    Dest = "skills";    Label = "Skills (11 skills)" },
    @{ Source = "agents";    Dest = "agents";    Label = "Agents (6 agents)" },
    @{ Source = "commands";  Dest = "commands";  Label = "Commands (4 commands)" },
    @{ Source = "rules";     Dest = "rules";     Label = "Rules (7 rules)" },
    @{ Source = "templates"; Dest = "templates"; Label = "Templates (8 templates)" },
    @{ Source = "scripts";   Dest = "scripts";   Label = "Scripts (hooks helpers)" }
)

foreach ($item in $deployItems) {
    Deploy-Directory `
        -Source (Join-Path $ScriptDir $item.Source) `
        -Dest (Join-Path $Target $item.Dest) `
        -Label $item.Label
}

# --- Merge hooks using Node.js (reliable UTF-8 handling) ---
$hooksSource = Join-Path $ScriptDir "hooks\hooks.json"
$settingsTarget = Join-Path $Target "settings.json"

if (Test-Path $hooksSource) {
    Write-Host "[Deploy] Hooks -> $settingsTarget" -ForegroundColor Green

    if (-not $DryRun) {
        $nodeScript = @"
const fs = require('fs');
const hooksPath = process.argv[1];
const settingsPath = process.argv[2];

let raw = fs.readFileSync(hooksPath, 'utf8');
if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1);
const hooks = JSON.parse(raw);

let settings = {};
if (fs.existsSync(settingsPath)) {
    let sraw = fs.readFileSync(settingsPath, 'utf8');
    if (sraw.charCodeAt(0) === 0xFEFF) sraw = sraw.slice(1);
    settings = JSON.parse(sraw);
}

settings.hooks = hooks.hooks;
fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2), 'utf8');
console.log('  -> Merged: PreToolUse=' + settings.hooks.PreToolUse.length + ' PostToolUse=' + settings.hooks.PostToolUse.length);
"@
        $nodeScript | node - "$hooksSource" "$settingsTarget"
    }
} else {
    Write-Host "[Skip] hooks.json not found" -ForegroundColor Yellow
}

# --- CLAUDE.md info ---
Write-Host ""
Write-Host "[Info] CLAUDE.md:" -ForegroundColor Magenta
Write-Host "  Root:  claude-md/CLAUDE.md" -ForegroundColor DarkMagenta
Write-Host "  CSharp: claude-md/csharp/CLAUDE.md" -ForegroundColor DarkMagenta
Write-Host "  Note: Not auto-deployed. Merge manually if needed." -ForegroundColor Yellow

# --- Summary ---
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "  DRY RUN complete (no changes made)" -ForegroundColor Yellow
} else {
    Write-Host "  Deploy complete!" -ForegroundColor Green
}
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Deployed to: $Target" -ForegroundColor White
Write-Host "  Skills:    11  Agents:   6  Commands: 4" -ForegroundColor Gray
Write-Host "  Rules:     7   Templates: 8  Hooks:    8" -ForegroundColor Gray
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Verify $settingsTarget" -ForegroundColor Gray
Write-Host "  2. Restart Claude Code to load new config" -ForegroundColor Gray
