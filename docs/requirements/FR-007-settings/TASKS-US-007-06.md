# TASKS: US-007-06 — MCP Auto-Start Toggle

> **Parent US**: [US-007-06](US-007-06-mcp-auto-start-toggle.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 5 | **Effort**: 4S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [ ] `AppDbContext` with `ServerConfigs` and `SyncLogs` DbSets available

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a toggle/checkbox for MCP server auto-start in the MCP tab
- [ ] AC-02: The toggle defaults to off (disabled) for new installations
- [ ] AC-03: When enabled, the MCP SSE server starts automatically on application launch
- [ ] AC-04: When disabled, the MCP SSE server does not start on application launch (can still be started manually)
- [ ] AC-05: Saving persists the auto-start setting to `ServerConfigs` database table with key `mcp.auto_start`
- [ ] AC-06: The auto-start value is also saved to the `appsettings.json` MCP section
- [ ] AC-07: Auto-start setting is loaded from database on page initialization with fallback to false
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-06-01: Add MCP auto-start toggle/checkbox to MCP tab in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the MCP tab (below the port field from US-007-05), add a `CheckBox` with `Content="Enable auto-start on launch"` bound to `{Binding MCPAutoStart}`
- Add a descriptive `TextBlock` below the checkbox: "When enabled, the MCP SSE server will start automatically when the application launches."
- Style the checkbox consistently with the MCP tab layout

### How to verify
- [ ] Checkbox renders in the MCP tab with label "Enable auto-start on launch" (AC-01)
- [ ] Checkbox is unchecked by default for new installations (AC-02)

---

## TASK-007-06-02: Add MCPAutoStart boolean property to ViewModel with default value false

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-06-03, TASK-007-06-04 |

### What to do
- Add `[ObservableProperty] bool _mcpAutoStart` with default `false`
- Wire `OnMCPAutoStartChanged` partial to update `HasUnsavedChanges`
- No validation required for boolean fields

### How to verify
- [ ] `MCPAutoStart` property generates change notifications and is bindable (AC-01)
- [ ] Default value is false (AC-02)

---

## TASK-007-06-03: Persist MCP auto-start to ServerConfigs table and appsettings.json on Save

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-06-02 |
| Blocks | TASK-007-06-04 |

### What to do
- In the `SaveSettingsCommand` handler, add persistence for MCP auto-start:
  - Upsert `ServerConfigs` row: `ConfigKey = "mcp.auto_start"`, `ConfigValue = MCPAutoStart.ToString().ToLower()` (store as "true"/"false")
  - Update `appsettings.json` via `ConfigurationLoader.SaveToJson()`: set `AppSettings.MCP.AutoStart = MCPAutoStart`
- Include the auto-start key in the `SyncLog` details JSON
- Update the original-values snapshot for `MCPAutoStart` after save

### How to verify
- [ ] After Save, `ServerConfigs` contains row with `mcp.auto_start` and correct value (AC-05)
- [ ] After Save, `appsettings.json` MCP section reflects the auto-start boolean (AC-06)
- [ ] `SyncLog` entry includes the auto-start change (AC-08)

---

## TASK-007-06-04: Load MCP auto-start from ServerConfigs table on page initialization with fallback to false

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-06-02, TASK-007-06-03 |
| Blocks | None |

### What to do
- In the ViewModel constructor or `InitializeAsync()`, query `ServerConfigs` for key `mcp.auto_start`
- Parse the value with `bool.TryParse()`; if parsing succeeds, set `MCPAutoStart` to the parsed value
- If no DB entry exists, fall back to `ConfigurationLoader.GetSettings().MCP.AutoStart` (default false)
- Include the loaded value in the original-values snapshot

### How to verify
- [ ] On page load, checkbox reflects the persisted value from DB (AC-07)
- [ ] On first load (no DB entry), checkbox is unchecked (default false) (AC-07)

---

## TASK-007-06-05: Integrate auto-start check in application startup logic

| Field | Value |
|-------|-------|
| Target | `App.xaml.cs` or startup bootstrap |
| Estimate | M |
| Depends On | TASK-007-06-03 |
| Blocks | None |

### What to do
- In the application startup sequence (e.g., `App.OnStartup()` or `MainViewModel` constructor), add logic to check the `mcp.auto_start` setting:
  1. Query `ServerConfigs` for `mcp.auto_start`; fall back to `appsettings.json` `MCP.AutoStart` value
  2. If `true`, resolve the MCP SSE server service and call its start method (e.g., `IMCPServerService.StartAsync()`)
  3. If `false`, do not start the MCP server (it can still be started manually from the UI)
- Read the MCP port from `ServerConfigs` (key `mcp.port`) to configure the server on the correct port
- Log the auto-start decision via `ILogger` (e.g., "MCP auto-start enabled, starting on port {port}" or "MCP auto-start disabled, skipping")
- Handle startup failures gracefully: if the MCP server fails to start, log a warning but do not block the application from launching

### How to verify
- [ ] When `mcp.auto_start` is true, MCP server starts automatically on app launch (AC-03)
- [ ] When `mcp.auto_start` is false, MCP server does not start on app launch (AC-04)
- [ ] MCP server startup failure does not crash the application (AC-03, AC-04)

---

## Dependency Graph
```
TASK-007-06-01 (XAML Checkbox)       ── independent

TASK-007-06-02 (ViewModel Property)
       │
       └──▶ TASK-007-06-03 (Persist on Save)
                  │
                  ├──▶ TASK-007-06-04 (Load on Init)
                  └──▶ TASK-007-06-05 (Startup Integration)

Tasks 01 and 02 can start in parallel.
```
