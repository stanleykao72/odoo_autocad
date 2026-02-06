# TASKS: US-006-09 — View Tool Prerequisites

> **Parent US**: [US-006-09](US-006-09-view-tool-prerequisites.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 5 | **Effort**: 3S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-006-05 (View Tool Registry - prerequisites are displayed within the tool list)
- [ ] `MCPToolRegistry` exists with tool definitions including RequiresAutoCAD, RequiresOdoo metadata
- [ ] `IAutoCADService.IsConnected` and `IOdooService.IsConnected` available via DI

## Acceptance Criteria
- [ ] AC-01: Each tool in the registered tools list displays prerequisite badges: [A] for AutoCAD-required, [O] for Odoo-required, or both [A][O]
- [ ] AC-02: A prerequisite legend is displayed in the Configuration panel explaining the badge meanings
- [ ] AC-03: Tools that require AutoCAD but AutoCAD is not connected show a visual warning indicator
- [ ] AC-04: Tools that require Odoo but Odoo is not connected show a visual warning indicator
- [ ] AC-05: When a tool with unmet prerequisites is invoked, the error message includes guidance to connect to the required service (e.g., "Tool 'draw_line' requires AutoCAD connection. Please connect to AutoCAD first.")

---

## TASK-006-09-01: Add RequiresAutoCAD and RequiresOdoo boolean properties to MCPToolViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-006-09-02 |

### What to do
- Ensure `MCPToolViewModel` has `bool RequiresAutoCAD` and `bool RequiresOdoo` properties (may already be added in US-006-05)
- Add an `ObservableProperty` `bool _isAutoCADConnected` and `bool _isOdooConnected` to `MCPViewModel` to track current connection state
- Add computed properties to `MCPToolViewModel`: `bool HasUnmetAutoCADPrerequisite => RequiresAutoCAD && !IsAutoCADConnected` and `bool HasUnmetOdooPrerequisite => RequiresOdoo && !IsOdooConnected` (these need a reference to the parent ViewModel's connection state or be updated via a refresh method)
- Alternative approach: add `bool IsAutoCADAvailable` and `bool IsOdooAvailable` properties to `MCPToolViewModel` that are set when the connection state changes
- Create a method `RefreshPrerequisiteStatus(bool isAutoCADConnected, bool isOdooConnected)` on `MCPToolViewModel` that updates availability status
- Map tool prerequisites based on the documented tool list:
  - AutoCAD-required [A]: check_autocad_status, create_new_drawing, draw_line, draw_circle, create_text, add_dimension, set_layer, list_layers, scan_elements, extract_autocad_parameters, export_to_database, get_current_layout, switch_to_layout, draw_in_layout, extract_layout_parameters, export_layout_image, process_natural_language_command
  - Odoo-required [O]: check_odoo_status, sync_to_odoo
  - Both [A][O]: generate_boq, sync_drawing_to_odoo, generate_boq_from_drawing
  - Neither: test_connection, get_server_info

### How to verify
- [ ] RequiresAutoCAD and RequiresOdoo are correctly set for all 24 tools (AC-01)
- [ ] Unmet prerequisite status is computable from connection state (AC-03, AC-04)

---

## TASK-006-09-02: Display [A] and [O] badges with warning indicators for unmet prerequisites

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | TASK-006-09-01 |
| Blocks | None |

### What to do
- In the tool entry `DataTemplate` (created in US-006-05), ensure [A] and [O] badges are displayed
- Add conditional styling for unmet prerequisites: when `HasUnmetAutoCADPrerequisite` is true, add a warning icon (exclamation triangle) or change the [A] badge background to red/orange with strikethrough or opacity change
- When `HasUnmetOdooPrerequisite` is true, apply similar warning styling to the [O] badge
- Add `ToolTip` on warning indicators: "AutoCAD is not connected. This tool will fail." or "Odoo is not connected. This tool will fail."
- Use `DataTrigger` bindings: `<DataTrigger Binding="{Binding HasUnmetAutoCADPrerequisite}" Value="True">` to change badge appearance
- Normal state: [A] badge with blue/cyan background; warning state: [A] badge with red/orange background and warning icon

### How to verify
- [ ] [A] badge shows warning when AutoCAD is not connected (AC-03)
- [ ] [O] badge shows warning when Odoo is not connected (AC-04)
- [ ] Tooltips explain the warning (AC-03, AC-04)

---

## TASK-006-09-03: Add prerequisite legend section to the Configuration panel XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Configuration panel (created in US-006-06), add or verify the "Prerequisites" subsection
- Display three legend entries with matching badge styling:
  - A small Border with "[A]" text (blue/cyan) + "= AutoCAD required" label
  - A small Border with "[O]" text (purple/orange) + "= Odoo required" label
  - A small Border with "[A][O]" text (combined) + "= Both required" label
- Use the same colors and styling as the badges in the tool list for visual consistency
- Keep the legend compact (single column, small margins)
- This task may overlap with TASK-006-06-04; if that task is already done, verify and skip

### How to verify
- [ ] Prerequisite legend is displayed in Configuration panel (AC-02)
- [ ] Legend matches badge styling used in tool list (AC-02)

---

## TASK-006-09-04: Implement prerequisite checking in MCPToolRegistry.ExecuteToolAsync() with descriptive error responses

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- In `ExecuteToolAsync()`, before calling the tool executor, check the tool's prerequisite requirements
- If `tool.RequiresAutoCAD` is true: check `_autoCADService?.IsConnected ?? false`; if not connected, return error result with message "Tool '{toolName}' requires AutoCAD connection. Please connect to AutoCAD first." and set `IsError = true`
- If `tool.RequiresOdoo` is true: check `_odooService?.IsConnected ?? false`; if not connected, return error result with message "Tool '{toolName}' requires Odoo connection. Please connect to Odoo first."
- Use the custom JSON-RPC error code `-32004` (NotConnectedError) for these errors
- Add the prerequisite metadata to the `_toolDefinitions` dictionary entries during registration: set `RequiresAutoCAD` and `RequiresOdoo` on each `MCPTool` based on the tool category and name
- Log the prerequisite failure at WARNING level in the logger
- Return a `MCPToolCallResult` with `IsError = true` containing the descriptive error message

### How to verify
- [ ] AutoCAD-required tools fail with descriptive error when AutoCAD is not connected (AC-05)
- [ ] Odoo-required tools fail with descriptive error when Odoo is not connected (AC-05)
- [ ] Error code -32004 is used for prerequisite failures (AC-05)

---

## TASK-006-09-05: Query IAutoCADService.IsConnected and IOdooService.IsConnected for live prerequisite status

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Inject `IAutoCADService?` and `IOdooService?` into `MCPViewModel` constructor
- Add `[ObservableProperty]` fields: `bool _isAutoCADConnected`, `bool _isOdooConnected` (if not already present)
- Subscribe to connection state change events from both services, or poll via the existing status timer
- When connection state changes: update `IsAutoCADConnected` and `IsOdooConnected` properties
- Call `RefreshPrerequisiteStatus(IsAutoCADConnected, IsOdooConnected)` on each `MCPToolViewModel` in the `RegisteredTools` collection to update warning indicators
- Use `Dispatcher.InvokeAsync` for thread-safe property updates
- Optionally provide navigation links: when a prerequisite is unmet, offer a hyperlink to navigate to the AutoCAD Connection page (FR-002) or Odoo Settings page (FR-003) via `INavigationService`

### How to verify
- [ ] AutoCAD connection state is tracked in real-time (AC-03)
- [ ] Odoo connection state is tracked in real-time (AC-04)
- [ ] Tool warning indicators update when connection state changes (AC-03, AC-04)

---

## Dependency Graph
```
TASK-006-09-01 (ViewModel Prerequisite Properties)
       │
       └──▶ TASK-006-09-02 (XAML Warning Badges)

TASK-006-09-03 (Prerequisite Legend XAML)
       (independent)

TASK-006-09-04 (MCPToolRegistry Prerequisite Checking)
       (independent)

TASK-006-09-05 (Live Connection State Tracking)
       (independent)

Tasks 01, 03, 04, and 05 are independent and can be developed in parallel.
Task 02 depends on 01 (prerequisite properties must exist for XAML bindings).
```
