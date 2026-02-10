#!/usr/bin/env bash
# deploy.sh - Deploy autocad-odoo-claude-code framework
# Usage: ./deploy.sh [--target <path>] [--backup] [--dry-run] [--clean-home]

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
TARGET="${PROJECT_ROOT}/.claude"
BACKUP=false
DRY_RUN=false
CLEAN_HOME=false

while [[ $# -gt 0 ]]; do
    case $1 in
        --target) TARGET="$2"; shift 2 ;;
        --backup) BACKUP=true; shift ;;
        --dry-run) DRY_RUN=true; shift ;;
        --clean-home) CLEAN_HOME=true; shift ;;
        -h|--help)
            echo "Usage: $0 [--target <path>] [--backup] [--dry-run] [--clean-home]"
            echo ""
            echo "Options:"
            echo "  --target <path>  Deploy target (default: project .claude/)"
            echo "  --backup         Backup existing target before deploy"
            echo "  --dry-run        Preview only, no changes"
            echo "  --clean-home     Remove framework dirs from HOME ~/.claude/"
            exit 0
            ;;
        *) echo "Unknown arg: $1"; exit 1 ;;
    esac
done

echo "============================================"
echo "  AutoCAD-Odoo Claude Code Framework Deploy"
echo "============================================"
echo ""
echo "Source:  ${SCRIPT_DIR}"
echo "Target:  ${TARGET}"
echo "Project: ${PROJECT_ROOT}"
echo ""

# --- CleanHome ---
if [ "$CLEAN_HOME" = true ]; then
    HOME_CLAUDE="${HOME}/.claude"
    echo "[CleanHome] Removing framework dirs from ${HOME_CLAUDE}"
    for d in agents skills rules templates scripts; do
        if [ -d "${HOME_CLAUDE}/${d}" ]; then
            echo "  Removing ${HOME_CLAUDE}/${d}"
            [ "$DRY_RUN" = false ] && rm -r "${HOME_CLAUDE}/${d}"
        fi
    done
    for f in tdd.md code-review.md git-push.md plan.md; do
        if [ -f "${HOME_CLAUDE}/commands/${f}" ]; then
            echo "  Removing ${HOME_CLAUDE}/commands/${f}"
            [ "$DRY_RUN" = false ] && rm "${HOME_CLAUDE}/commands/${f}"
        fi
    done
    echo "[CleanHome] Done"
    echo ""
fi

# --- Backup ---
if [ "$BACKUP" = true ] && [ -d "$TARGET" ]; then
    TIMESTAMP=$(date +%Y%m%d_%H%M%S)
    BACKUP_DIR="${TARGET}_backup_${TIMESTAMP}"
    echo "[Backup] Creating backup at ${BACKUP_DIR}"
    [ "$DRY_RUN" = false ] && cp -r "$TARGET" "$BACKUP_DIR"
fi

# --- Ensure target ---
[ "$DRY_RUN" = false ] && mkdir -p "$TARGET"

# --- Deploy function ---
deploy_dir() {
    local src="$1" dest="$2" label="$3"
    if [ -d "${SCRIPT_DIR}/${src}" ]; then
        echo "[Deploy] ${label} -> ${TARGET}/${dest}"
        if [ "$DRY_RUN" = false ]; then
            mkdir -p "${TARGET}/${dest}"
            cp -r "${SCRIPT_DIR}/${src}/"* "${TARGET}/${dest}/"
        fi
    else
        echo "[Skip] ${label} not found"
    fi
}

deploy_dir "skills"    "skills"    "Skills (11 skills)"
deploy_dir "agents"    "agents"    "Agents (6 agents)"
deploy_dir "commands"  "commands"  "Commands (4 commands)"
deploy_dir "rules"     "rules"     "Rules (7 rules)"
deploy_dir "templates" "templates" "Templates (8 templates)"
deploy_dir "scripts"   "scripts"   "Scripts (hooks helpers)"

# --- Merge hooks via Node.js ---
HOOKS_SOURCE="${SCRIPT_DIR}/hooks/hooks.json"
SETTINGS_TARGET="${TARGET}/settings.json"

if [ -f "$HOOKS_SOURCE" ]; then
    echo "[Deploy] Hooks -> ${SETTINGS_TARGET}"
    if [ "$DRY_RUN" = false ]; then
        node -e "
const fs = require('fs');
let raw = fs.readFileSync(process.argv[1], 'utf8');
if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1);
const hooks = JSON.parse(raw);
let settings = {};
if (fs.existsSync(process.argv[2])) {
    let sraw = fs.readFileSync(process.argv[2], 'utf8');
    if (sraw.charCodeAt(0) === 0xFEFF) sraw = sraw.slice(1);
    settings = JSON.parse(sraw);
}
settings.hooks = hooks.hooks;
fs.writeFileSync(process.argv[2], JSON.stringify(settings, null, 2), 'utf8');
console.log('  -> Merged: PreToolUse=' + settings.hooks.PreToolUse.length + ' PostToolUse=' + settings.hooks.PostToolUse.length);
" "$HOOKS_SOURCE" "$SETTINGS_TARGET"
    fi
fi

# --- Summary ---
echo ""
echo "============================================"
if [ "$DRY_RUN" = true ]; then
    echo "  DRY RUN complete (no changes made)"
else
    echo "  Deploy complete!"
fi
echo "============================================"
echo ""
echo "Deployed to: ${TARGET}"
echo "  Skills: 11  Agents: 6  Commands: 4"
echo "  Rules:  7   Templates: 8  Hooks: 8"
echo ""
echo "Next steps:"
echo "  1. Verify ${SETTINGS_TARGET}"
echo "  2. Restart Claude Code to load new config"
