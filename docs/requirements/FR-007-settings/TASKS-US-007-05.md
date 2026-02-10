# TASKS: US-007-05 — MCP Port Configuration

> **Parent US**: [US-007-05](US-007-05-mcp-port-configuration.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 5S + 1M + 0L
> **Status**: In Progress

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [ ] `AppDbContext` with `ServerConfigs` and `SyncLogs` DbSets available

## Acceptance Criteria
- [x] AC-01: Settings Page displays an input field for MCP SSE Server port number in the MCP tab with default value 8084
- [ ] AC-02: Port number validates as integer between 1024 and 65535 inclusive, with inline error for out-of-range values
- [ ] AC-03: A warning is displayed (non-blocking) if the entered port appears to be in use, with message "Port {port} appears to be in use. MCP server may not start. Choose a different port."
- [x] AC-04: Saving persists the port number to `ServerConfigs` database table with key `mcp.port`
- [x] AC-05: The port value is also saved to the `appsettings.json` MCP section
- [x] AC-06: Port setting is loaded from database on page initialization with fallback to default 8084
- [ ] AC-07: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-05-01: Add MCP tab with SSE Server Port input field in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Populate the MCP tab (stub created in US-007-01) with an "MCP Server Settings" group
- Add a `TextBox` with label "SSE Server Port:" bound to `{Binding MCPPort, UpdateSourceTrigger=PropertyChanged, ValidatesOnNotifyDataErrors=True}`
- Add `InputScope="Number"` and integer-only input restriction (reuse the numeric filter from US-007-04)
- Add a `TextBlock` below the port field for the port-in-use warning, bound to a warning message property, with `Visibility` collapsed when empty, styled with orange/yellow foreground
- Include inline validation error display via `Validation.ErrorTemplate`

### How to verify
- [x] MCP tab renders with port input field showing default 8084 (AC-01)
- [x] Warning text area is present but hidden when no warning (AC-03)

---

## TASK-007-05-02: Add MCPPort property to ViewModel with default value 8084

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-05-03, TASK-007-05-04, TASK-007-05-05 |

### What to do
- Add `[ObservableProperty] int _mcpPort` with default `8084`
- Wire `OnMCPPortChanged` partial to update `HasUnsavedChanges` and trigger port validation
- Add `[ObservableProperty] string _mcpPortWarning` (default empty) for the port-in-use warning message

### How to verify
- [x] `MCPPort` property generates change notifications and is bindable (AC-01)
- [x] Default value is 8084 (AC-01)

---

## TASK-007-05-03: Implement validation for port range (1024-65535) with inline error messages

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-05-02 |
| Blocks | None |

### What to do
- In `OnMCPPortChanged`, validate: if value < 1024 or value > 65535, call `SetErrors("MCPPort", new[] { "Port must be between 1024 and 65535." })`, else `ClearErrors("MCPPort")`
- Validation rule corresponds to VR-007-006
- Ensure `CanSave` recalculates when this error state changes
- Optionally add a check for well-known reserved ports (e.g., 3306, 5432, 8080) and set a warning (non-blocking) rather than an error

### How to verify
- [ ] Entering port outside 1024-65535 shows inline error message (AC-02)
- [ ] Save button is disabled when port validation fails (AC-02)

---

## TASK-007-05-04: Implement optional port-in-use detection and display non-blocking warning

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-05-02 |
| Blocks | None |

### What to do
- In `OnMCPPortChanged`, after a short debounce (300ms using `Task.Delay` with cancellation), attempt to detect if the port is in use
- Use `System.Net.Sockets.TcpClient` to try connecting to `localhost:{port}` with a 1-second timeout; if connection succeeds, the port is in use
- Alternatively, use `System.Net.NetworkInformation.IPGlobalProperties.GetActiveTcpListeners()` to check if any listener is bound to the port
- If in use: set `MCPPortWarning = $"Port {MCPPort} appears to be in use. MCP server may not start. Choose a different port."`
- If not in use or check fails: set `MCPPortWarning = string.Empty`
- This is a **non-blocking warning** -- the Save button remains enabled even when the warning is displayed
- Run the check on a background thread to avoid UI freezing

### How to verify
- [ ] Warning message appears when the entered port is detected as in use (AC-03)
- [ ] Warning message clears when the port is changed to an available one (AC-03)
- [ ] Save button remains enabled despite the warning (AC-03)

---

## TASK-007-05-05: Persist MCP port to ServerConfigs table and appsettings.json on Save

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-05-02 |
| Blocks | TASK-007-05-06 |

### What to do
- In the `SaveSettingsCommand` handler, add persistence for MCP port:
  - Upsert `ServerConfigs` row: `ConfigKey = "mcp.port"`, `ConfigValue = MCPPort.ToString()`
  - Update `appsettings.json` via `ConfigurationLoader.SaveToJson()`: set `AppSettings.MCP.Port = MCPPort`
- Include the MCP port key in the `SyncLog` details JSON
- Update the original-values snapshot for `MCPPort` after save

### How to verify
- [x] After Save, `ServerConfigs` contains row with `mcp.port` and correct value (AC-04)
- [x] After Save, `appsettings.json` MCP section reflects the new port number (AC-05)
- [x] `SyncLog` entry includes the port change (AC-07)

---

## TASK-007-05-06: Load MCP port from ServerConfigs table on page initialization with fallback to default

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-05-05 |
| Blocks | None |

### What to do
- In the ViewModel constructor or `InitializeAsync()`, query `ServerConfigs` for key `mcp.port`
- Parse the value with `int.TryParse()`; if parsing succeeds, set `MCPPort` to the parsed value
- If no DB entry exists, fall back to `ConfigurationLoader.GetSettings().MCP.Port` (default 8084)
- Include the loaded value in the original-values snapshot

### How to verify
- [x] On page load, port field shows the value from DB if present (AC-06)
- [x] On first load (no DB entry), port field shows default 8084 (AC-06)

---

## Dependency Graph
```
TASK-007-05-01 (XAML Port Field)     ── independent

TASK-007-05-02 (ViewModel Property)
       │
       ├──▶ TASK-007-05-03 (Port Validation)
       ├──▶ TASK-007-05-04 (Port-in-Use Detection)
       └──▶ TASK-007-05-05 (Persist on Save)
                  │
                  └──▶ TASK-007-05-06 (Load on Init)

Tasks 01 and 02 can start in parallel.
Tasks 03, 04, and 05 can proceed in parallel once 02 is complete.
```
