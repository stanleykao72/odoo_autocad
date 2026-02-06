# TASKS: US-008-03 — Persistent Connection Status

> **Parent US**: [US-008-03](US-008-03-persistent-connection-status.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 5 | **Effort**: 2S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-008-01 completed (sidebar layout must exist to host the connection status section)

## Acceptance Criteria
- [ ] AC-01: The sidebar displays a "Connection Status" section at the bottom with status indicators for AutoCAD, Odoo, and MCP Server.
- [ ] AC-02: Each connection indicator displays a colored `Ellipse` (green = connected/running, gray = disconnected/stopped) and descriptive status text.
- [ ] AC-03: AutoCAD status shows "Connected" (green) or "Disconnected" (gray) based on `IAutoCADService` connection state.
- [ ] AC-04: Odoo status shows "Connected" (green) or "Disconnected" (gray) based on `IOdooService` connection state.
- [ ] AC-05: MCP status shows "Port XXXX" (green) when the server is running, or "Stopped" (gray) when it is not.
- [ ] AC-06: Connection indicators update in real-time via data binding to `MainViewModel` observable properties (`AutoCADStatusColor`, `OdooStatusColor`, `MCPStatusColor`).
- [ ] AC-07: Connection status colors only transition between defined states: Gray (disconnected/stopped), Green (connected/running), and Red (error).
- [ ] AC-08: The connection status section is visible on all pages without requiring user interaction.

---

## TASK-008-03-01: Add connection status observable properties to MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-008-03-03, TASK-008-03-05 |

### What to do
- Verify `[ObservableProperty]` fields exist for: `_autoCADStatusText` (default "Disconnected"), `_autoCADStatusColor` (default `Brushes.Gray`), `_odooStatusText` (default "Disconnected"), `_odooStatusColor` (default `Brushes.Gray`), `_mcpStatusText` (default "Stopped"), `_mcpStatusColor` (default `Brushes.Gray`)
- Ensure all six properties use `System.Windows.Media.Brush` type for color bindings so they can bind directly to `Ellipse.Fill`
- Verify properties are already present in the skeleton `MainViewModel.cs`; if not, add them
- Restrict color values to `Brushes.Gray`, `Brushes.Green`, and `Brushes.Red` per VR-008-006

### How to verify
- [ ] All six observable properties exist and raise `PropertyChanged` (AC-06)
- [ ] Default values are "Disconnected"/Gray for AutoCAD and Odoo, "Stopped"/Gray for MCP (AC-02, AC-03, AC-04, AC-05)

---

## TASK-008-03-02: Define connection status XAML section in sidebar with Ellipse indicators and bound TextBlocks

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | M |
| Depends On | TASK-008-03-01 |
| Blocks | None |

### What to do
- In the sidebar Grid Row 2, define a `Border` with top border separator and 15px padding
- Add "Connection Status" `TextBlock` header with `FontWeight="SemiBold"` and bottom margin
- For each service (AutoCAD, Odoo, MCP Server), add a `Grid` with three columns: `Ellipse` (10x10, `Fill="{Binding AutoCADStatusColor}"`, margin-right 8), service name `TextBlock`, and status text `TextBlock` (`Text="{Binding AutoCADStatusText}"`, `TextSecondaryBrush`, 11pt)
- Repeat the three-column `Grid` pattern for Odoo (`OdooStatusColor`/`OdooStatusText`) and MCP (`MCPStatusColor`/`MCPStatusText`)
- Use consistent 8px bottom margin between each service row
- Ensure the section is inside the sidebar and therefore visible on all pages regardless of Frame content

### How to verify
- [ ] "Connection Status" section is at the bottom of the sidebar (AC-01)
- [ ] Each indicator has a colored Ellipse and descriptive text (AC-02)
- [ ] Section is visible on all pages (AC-08)

---

## TASK-008-03-03: Implement status update methods in MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | M |
| Depends On | TASK-008-03-01 |
| Blocks | TASK-008-03-05 |

### What to do
- Implement `UpdateAutoCADStatus()` to check `IAutoCADService` connection state (if available via optional DI): set `AutoCADStatusText` to "Connected"/"Disconnected" and `AutoCADStatusColor` to `Brushes.Green`/`Brushes.Gray`
- Implement `UpdateOdooStatus()` to check `IOdooService` connection state (if available via optional DI): set `OdooStatusText` to "Connected"/"Disconnected" and `OdooStatusColor` to `Brushes.Green`/`Brushes.Gray`
- Implement `UpdateMCPStatus()` to check `MCPSSEServer.IsRunning` and `MCPSSEServer.Port`: if running, set text to `$"Port {port}"` and color to `Brushes.Green`; if stopped, set text to "Stopped" and color to `Brushes.Gray`
- Add `UpdateAllStatus()` method that calls all three update methods
- Call `UpdateAllStatus()` from the constructor and from the `_statusTimer` tick handler (every 1 second)
- Restrict colors to Gray, Green, Red only per VR-008-006

### How to verify
- [ ] AutoCAD status reflects IAutoCADService state (AC-03)
- [ ] Odoo status reflects IOdooService state (AC-04)
- [ ] MCP status shows port when running, "Stopped" when not (AC-05)
- [ ] Colors are restricted to Gray/Green/Red (AC-07)

---

## TASK-008-03-04: Define connected/disconnected/error color resources

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify that `SuccessBrush` (`#388E3C`) is already defined in `App.xaml` for connected states
- Verify that `ErrorBrush` (`#D32F2F`) is already defined for error states
- Optionally add explicit `ConnectedGreenColor`/`ConnectedGreenBrush` aliases if different shade is needed
- Connection status indicators use `Brushes.Gray`, `Brushes.Green`, and `Brushes.Red` directly from code, so no additional XAML resources are strictly required
- Ensure the color palette aligns with VR-008-006 (only Gray, Green, Red transitions)

### How to verify
- [ ] Color resources for success/error states exist in ResourceDictionary (AC-07)
- [ ] No unauthorized colors used for connection status indicators (AC-07)

---

## TASK-008-03-05: Wire MCP server status to display port number when running

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | TASK-008-03-01, TASK-008-03-03 |
| Blocks | None |

### What to do
- In `UpdateMCPStatus()`, read `_mcpServer.IsRunning` and `_mcpServer.Port` properties
- When running: set `McpStatusText = $"Port {_mcpServer.Port}"` and `McpStatusColor = Brushes.Green`
- When stopped: set `McpStatusText = "Stopped"` and `McpStatusColor = Brushes.Gray`
- Ensure the `OnStatusTimerTick` calls `UpdateMCPStatus()` for real-time updates
- Verify the skeleton already implements this pattern; refine if needed

### How to verify
- [ ] MCP status displays port number when running (AC-05)
- [ ] MCP status displays "Stopped" when not running (AC-05)
- [ ] Updates are real-time via timer-driven polling (AC-06)

---

## Dependency Graph

```
TASK-008-03-01 (Observable properties)
    ├── TASK-008-03-02 (XAML indicators)
    ├── TASK-008-03-03 (Status update methods)
    │   └── TASK-008-03-05 (MCP port display)
    └── TASK-008-03-05 (MCP port display)

TASK-008-03-04 (Color resources) [independent]
```
