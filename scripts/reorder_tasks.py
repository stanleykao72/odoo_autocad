#!/usr/bin/env python3
"""Reorder master-task-sequence.md sprints according to actual implementation priority."""
import re
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
ORIGINAL = os.path.join(PROJECT_ROOT, "docs", "requirements", "master-task-sequence.md")
OUTPUT = os.path.join(PROJECT_ROOT, "docs", "requirements", "master-task-sequence-update.md")

# Read original file
with open(ORIGINAL, "r", encoding="utf-8") as f:
    lines = f.readlines()

# Parse task rows: extract lines matching "| ### |" pattern with task data
task_rows = {}
for line in lines:
    m = re.match(r'\| (\d+) \| (TASK-\S+) \| (.+)$', line)
    if m:
        num = int(m.group(1))
        task_rows[num] = line.strip()

def renumber_row(old_line, new_num):
    return re.sub(r'^\| \d+ \|', f'| {new_num} |', old_line)

def count_effort(task_nums):
    s = m = l = 0
    for n in task_nums:
        row = task_rows[n]
        parts = row.split('|')
        est = parts[5].strip() if len(parts) > 5 else ''
        if est == 'S':
            s += 1
        elif est == 'M':
            m += 1
        elif est == 'L':
            l += 1
    return s, m, l

# Define sprint task mappings (old task numbers)
sprint_1_tasks = list(range(1, 33))      # 32 tasks
# Sprint 2: US-007-01(old 204-211) + US-007-04/05/06(old 351-367) + US-007-13(old 410-416)
sprint_2_tasks = list(range(204, 212)) + list(range(351, 368)) + list(range(410, 417))  # 8+17+7=32
sprint_3_tasks = list(range(64, 104))     # 40 tasks (old Sprint 3)
sprint_4_tasks = list(range(33, 64))      # 31 tasks (old Sprint 2)
sprint_5_tasks = list(range(104, 146))    # 42 tasks (old Sprint 4)
sprint_6_tasks = list(range(146, 179))    # 33 tasks (old Sprint 5)
sprint_7_tasks = list(range(179, 204))    # 25 tasks (old Sprint 6 PR-only)
sprint_8_tasks = list(range(225, 282))    # 57 tasks (old Sprint 7)
sprint_9_tasks = list(range(282, 351))    # 69 tasks (old Sprint 8)
# Sprint 10: US-007-02/03(old 212-224) + US-007-07-12(old 368-409) + US-008-07(old 417-421)
sprint_10_tasks = list(range(212, 225)) + list(range(368, 410)) + list(range(417, 422))  # 13+42+5=60
sprint_11_tasks = list(range(422, 463))   # 41 tasks (old Sprint 10)
sprint_12_tasks = list(range(463, 501))   # 38 tasks (old Sprint 11)

# Verify
all_tasks = (sprint_1_tasks + sprint_2_tasks + sprint_3_tasks + sprint_4_tasks +
             sprint_5_tasks + sprint_6_tasks + sprint_7_tasks + sprint_8_tasks +
             sprint_9_tasks + sprint_10_tasks + sprint_11_tasks + sprint_12_tasks)
assert len(all_tasks) == 500, f"Expected 500 tasks, got {len(all_tasks)}"
assert len(set(all_tasks)) == 500, "Duplicate task numbers found!"
assert set(all_tasks) == set(range(1, 501)), "Missing or extra task numbers!"

sprints = [
    (1, "App Shell & Navigation", True,
     "Application skeleton, sidebar navigation, window management, fonts, clean shutdown",
     "US-008-01, US-008-02, US-008-05, US-008-08, US-008-09, US-008-10", sprint_1_tasks),
    (2, "Settings Page & Config Persistence", True,
     "Settings page with Odoo connection, AutoCAD parameters, MCP server config, and application info display",
     "US-007-01, US-007-04, US-007-05, US-007-06, US-007-13", sprint_2_tasks),
    (3, "AutoCAD & Odoo Core Connections", True,
     "AutoCAD connect/layouts/parameter extraction, Odoo credentials/test/appsettings config",
     "US-002-01, US-002-02, US-002-03, US-003-01, US-003-02, US-003-10", sprint_3_tasks),
    (4, "Status Bar, Log Panel & Dashboard", False,
     "Connection status bar, system log panel, auto-start services, dashboard views",
     "US-008-03, US-008-04, US-008-06, US-001-01, US-001-02", sprint_4_tasks),
    (5, "Odoo Advanced & BOQ Foundation", False,
     "Odoo connection status/disconnect/sync product catalog/persist credentials, BOQ extract/review",
     "US-003-03, US-003-04, US-003-05, US-003-12, US-004-01, US-004-02", sprint_5_tasks),
    (6, "BOQ Processing Pipeline", False,
     "BOQ validate against products, push to Odoo, ID writeback, skipped rows, summary",
     "US-004-03, US-004-04, US-004-05, US-004-06, US-004-07", sprint_6_tasks),
    (7, "Purchase Requisition Core", False,
     "PR conversion from BOQ, PR list display, PR submission workflow",
     "US-005-01, US-005-02, US-005-04", sprint_7_tasks),
    (8, "AutoCAD P2 & Odoo Search/Filter", False,
     "AutoCAD drawing info/PR project/clear IDs/COM monitor, Odoo search/filter/sync time/server info",
     "US-002-04, US-002-05, US-002-06, US-002-07, US-003-06, US-003-07, US-003-08, US-003-09, US-003-11",
     sprint_8_tasks),
    (9, "BOQ P2 & PR Feature Completion", False,
     "BOQ product mapping/validation errors/clear IDs/manage mappings/progress, PR details/status/filter/feedback/totals",
     "US-004-08, US-004-09, US-004-10, US-004-11, US-004-12, US-005-03, US-005-05, US-005-06, US-005-07, US-005-08",
     sprint_9_tasks),
    (10, "Settings Advanced & Keyboard Shortcuts", False,
     "Environment switching, test connection, theme/language/log level, cache/export/import config, keyboard shortcuts",
     "US-007-02, US-007-03, US-007-07, US-007-08, US-007-09, US-007-10, US-007-11, US-007-12, US-008-07",
     sprint_10_tasks),
    (11, "MCP Server Foundation", False,
     "MCP server start/stop, server status, test connection, tool registry, configuration, recent activity",
     "US-006-01, US-006-02, US-006-03, US-006-04, US-006-05, US-006-06, US-001-03", sprint_11_tasks),
    (12, "MCP Monitoring & Configuration", False,
     "MCP connection monitoring, activity logs, tool prerequisites, restart, tool execution, port config, MCP dashboard status",
     "US-006-07, US-006-08, US-006-09, US-006-10, US-006-11, US-006-12, US-001-04", sprint_12_tasks),
]

phases = [
    ("Phase 1: Core Infrastructure", [0, 1, 2]),
    ("Phase 2: UI Foundation & Business Logic", [3, 4, 5]),
    ("Phase 3: Feature Completion", [6, 7, 8]),
    ("Phase 4: Polish & AI Integration", [9, 10, 11]),
]

out = []

# Header
out.append("# Master Task Sequence \u2014 Odoo-AutoCAD C# WPF")
out.append("")
out.append("> **Total Tasks**: 500 | **Phases**: 4 | **Sprints**: 12")
out.append("> **Estimates**: 313S + 160M + 27L")
out.append("> **User Stories**: 78 | **Generated**: 2026-02-06")
out.append("> **Last Reordered**: 2026-02-10")
out.append("")
out.append("## Overview")
out.append("")
out.append("This document sequences all implementation tasks for the Odoo-AutoCAD C# WPF desktop application")
out.append("across 4 phases and 12 sprints. Sprints are ordered by actual implementation priority, with")
out.append("foundational infrastructure first and feature-specific work building on top. Each task references")
out.append("its parent User Story for full context and acceptance criteria.")
out.append("")

# Progress Summary
out.append("## Progress Summary")
out.append("")
out.append("| Phase | Sprints | User Stories | Tasks | S | M | L | Status |")
out.append("|-------|---------|-------------|-------|---|---|---|--------|")

for phase_name, sprint_indices in phases:
    sprint_range = f"{sprint_indices[0]+1}\u2013{sprint_indices[-1]+1}"
    us_count = sum(len(sprints[i][4].split(", ")) for i in sprint_indices)
    task_count = sum(len(sprints[i][5]) for i in sprint_indices)
    s_total = sum(count_effort(sprints[i][5])[0] for i in sprint_indices)
    m_total = sum(count_effort(sprints[i][5])[1] for i in sprint_indices)
    l_total = sum(count_effort(sprints[i][5])[2] for i in sprint_indices)
    any_done = any(sprints[i][2] for i in sprint_indices)
    all_done = all(sprints[i][2] for i in sprint_indices)
    if all_done:
        status = "Complete"
    elif any_done:
        status = "In Progress"
    else:
        status = "Not Started"
    out.append(f"| {phase_name} | {sprint_range} | {us_count} | {task_count} | {s_total} | {m_total} | {l_total} | {status} |")

out.append("")
out.append("---")
out.append("")

# Write sprints
new_num = 1
for si, (sprint_num, name, completed, focus, user_stories, tasks) in enumerate(sprints):
    # Phase header
    for pi, (pname, pindices) in enumerate(phases):
        if si == pindices[0]:
            out.append(f"## {pname}")
            out.append("")
            break

    s, m, l = count_effort(tasks)
    total = s + m + l
    done_marker = " \u2713" if completed else ""
    out.append(f"### Sprint {sprint_num}: {name}{done_marker}")
    out.append(f"> **Focus**: {focus}")
    out.append(f"> **User Stories**: {user_stories}")
    out.append(f"> **Tasks**: {total} ({s}S + {m}M + {l}L)")
    out.append("")
    out.append("| # | Task ID | Title | Target | Est | Depends On | Status |")
    out.append("|---|---------|-------|--------|-----|------------|--------|")

    for old_num in tasks:
        row = task_rows[old_num]
        new_row = renumber_row(row, new_num)
        out.append(new_row)
        new_num += 1

    out.append("")

    # Phase separator
    for pi, (pname, pindices) in enumerate(phases):
        if si == pindices[-1]:
            out.append("---")
            out.append("")
            break

# Cross-Sprint Dependencies
out.append("## Cross-Sprint Dependencies")
out.append("")
out.append("| Dependency | Description |")
out.append("|-----------|-------------|")
deps = [
    ("Sprint 1 \u2192 Sprint 3", "App shell, sidebar, navigation needed for AutoCAD/Odoo connection pages"),
    ("Sprint 1 \u2192 Sprint 4", "Window framework needed for status bar, log panel, dashboard"),
    ("Sprint 2 \u2192 Sprint 10", "Settings core infrastructure needed for advanced settings features"),
    ("Sprint 3 \u2192 Sprint 4", "AutoCAD/Odoo service interfaces needed for dashboard status cards"),
    ("Sprint 3 \u2192 Sprint 5", "AutoCAD/Odoo connections needed for sync and BOQ operations"),
    ("Sprint 5 \u2192 Sprint 6", "BOQ extract/review foundation needed for validate/push pipeline"),
    ("Sprint 5 \u2192 Sprint 7", "Odoo sync needed for PR conversion"),
    ("Sprint 6 \u2192 Sprint 7", "BOQ push to Odoo needed before PR conversion"),
    ("Sprint 3 \u2192 Sprint 8", "Core AutoCAD/Odoo interfaces needed for advanced features"),
    ("Sprint 6 \u2192 Sprint 9", "BOQ pipeline needed for advanced BOQ features"),
    ("Sprint 7 \u2192 Sprint 9", "PR core needed for PR feature completion"),
    ("Sprint 3 \u2192 Sprint 11", "AutoCAD/Odoo services needed for MCP tool prerequisites"),
    ("Sprint 11 \u2192 Sprint 12", "MCP server foundation needed for monitoring features"),
]
for dep, desc in deps:
    out.append(f"| {dep} | {desc} |")
out.append("")

# Effort Summary
out.append("## Effort Summary by Sprint")
out.append("")
out.append("| Sprint | Name | S | M | L | Total |")
out.append("|--------|------|---|---|---|-------|")
total_s = total_m = total_l = 0
for sprint_num, name, _, _, _, tasks in sprints:
    s, m, l = count_effort(tasks)
    total_s += s; total_m += m; total_l += l
    out.append(f"| {sprint_num} | {name} | {s} | {m} | {l} | {s+m+l} |")
out.append(f"| **Total** | | **{total_s}** | **{total_m}** | **{total_l}** | **{total_s+total_m+total_l}** |")
out.append("")

# User Story Index
us_sprint_map = {}
for sprint_num, _, _, _, user_stories, _ in sprints:
    for us in user_stories.split(", "):
        us_sprint_map[us.strip()] = sprint_num

us_index_rows = []
in_us_index = False
for line in lines:
    if line.strip().startswith("## User Story Index"):
        in_us_index = True
        continue
    if in_us_index:
        if line.strip().startswith("## "):
            break
        m = re.match(r'\| (US-\d+-\d+) \| (.+?) \| (P\d) \| (\d+) \| (\d+) \| (.+?) \|', line.strip())
        if m:
            us_id = m.group(1)
            title = m.group(2)
            priority = m.group(3)
            task_count = m.group(4)
            task_file = m.group(6)
            new_sprint = us_sprint_map.get(us_id, m.group(5))
            us_index_rows.append((us_id, title, priority, task_count, str(new_sprint), task_file))

out.append("## User Story Index")
out.append("")
out.append("| US ID | Title | Priority | Tasks | Sprint | TASK File |")
out.append("|-------|-------|----------|-------|--------|-----------|")
for us_id, title, priority, task_count, sprint, task_file in us_index_rows:
    out.append(f"| {us_id} | {title} | {priority} | {task_count} | {sprint} | {task_file} |")
out.append("")

# Legend
out.append("## Legend")
out.append("")
out.append("- **S** (Small): < 2 hours, single file, straightforward implementation")
out.append("- **M** (Medium): 2\u20134 hours, multiple files or complex logic")
out.append("- **L** (Large): 4\u20138 hours, architectural decisions, cross-cutting concerns")
out.append("- **Est** = Effort Estimate")
out.append("- **Status**: [ ] Not Started, [~] In Progress, [x] Completed")
out.append("- **\u2014** in Depends On = No dependencies (can start immediately)")
out.append("")

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write("\n".join(out))

print(f"Written {len(out)} lines to {OUTPUT}")
print(f"Total S={total_s}, M={total_m}, L={total_l}, Grand={total_s+total_m+total_l}")
print(f"New task numbering ends at: {new_num - 1}")
