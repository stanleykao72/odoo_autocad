# TASKS: US-001-04 — MCP Server Status

> **Parent US**: [US-001-04](US-001-04-mcp-server-status.md)
> **Parent FR**: [FR-001](FR-001-dashboard.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 3S + 3M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-001-01 (shares the same card layout pattern and BoolToColorConverter for status indicators)
- [ ] MCP SSE Manager service interface being available for server lifecycle control (MCPSSEServer already exists in `OdooAutoCAD.MCP.Server`)

## Acceptance Criteria
- [ ] AC-01: Dashboard displays an MCP Server status card with running/stopped state
- [ ] AC-02: MCP Server status card displays the port number when the server is running (e.g., "Port: 8084")
- [ ] AC-03: Running state displays with a green indicator; stopped state displays with a red/gray indicator
- [ ] AC-04: Dashboard provides a toggle button to start/stop the MCP server
- [ ] AC-05: MCP server status updates in real-time when the server state changes
- [ ] AC-06: When MCP server fails to start, the card displays "MCP Server failed to start on port {port}. Port may be in use."
- [ ] AC-07: Toggle button label reflects current state (e.g., "Start" when stopped, "Stop" when running)

---

## TASK-001-04-01: Create MCP Server status card in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-001-04-06 |

### What to do
- In the connection status cards row (created in US-001-01 TASK-001-01-01), add a third `Border` card for MCP Server alongside the AutoCAD and Odoo cards
- The card should follow the same visual pattern: `CornerRadius="8"`, `Background="{StaticResource SurfaceBrush}"`, `BorderBrush="{StaticResource BorderBrush}"`, `BorderThickness="1"`, `Padding="16"`
- Inside the card StackPanel, add:
  - An icon (robot emoji TextBlock or a Path/icon resource)
  - A title TextBlock: "MCP Server"
  - An `Ellipse` (Width=12, Height=12) for status indicator with `Fill` bound to `{Binding MCPStatusColor}` (reusing the Brush property pattern from US-001-01)
  - A status label TextBlock bound to `{Binding MCPStatusText}` showing "Running" or "Stopped"
  - A port info TextBlock bound to `{Binding MCPPortDisplay}` (e.g., "Port: 8084"), visible only when running
  - A toggle `Button` with `Content` bound to `{Binding ToggleMCPButtonText}` and `Command` bound to `{Binding ToggleMCPCommand}`
  - An error message TextBlock bound to `{Binding MCPErrorMessage}`, collapsed when empty
- Update the cards container to accommodate 3 cards (change `UniformGrid Columns="3"` or adjust `StackPanel` widths)

### How to verify
- [ ] MCP Server card renders alongside AutoCAD and Odoo cards with consistent styling (AC-01)
- [ ] Port number is displayed when the server is running (AC-02)
- [ ] Toggle button is visible within the card (AC-04)

---

## TASK-001-04-02: Add ViewModel properties for MCP status

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-001-04-03, TASK-001-04-04, TASK-001-04-05 |

### What to do
- Add `[ObservableProperty]` fields to `DashboardViewModel`: `bool _isMCPRunning` (default false), `string _mcpStatus` (default "Stopped"), `int _mcpPort` (default 8084)
- Add `[ObservableProperty]` field: `string _mcpErrorMessage` (default empty)
- Add `[ObservableProperty]` field: `Brush _mcpStatusColor` (default `Brushes.Gray`) using `System.Windows.Media`
- Add a computed read-only property `string MCPPortDisplay` that returns `$"Port: {MCPPort}"` when `IsMCPRunning` is true, or empty string when stopped
- Add a computed read-only property `string ToggleMCPButtonText` that returns `"Stop"` when `IsMCPRunning` is true, `"Start"` when false
- In `partial void OnIsMCPRunningChanged(bool value)`, call `OnPropertyChanged(nameof(MCPPortDisplay))` and `OnPropertyChanged(nameof(ToggleMCPButtonText))`
- Inject `MCPSSEServer` into `DashboardViewModel` constructor (already available in DI as singleton) and store as `private readonly MCPSSEServer _mcpServer`

### How to verify
- [ ] All MCP properties are bindable and generate change notifications (AC-01)
- [ ] Port display and toggle button text react to IsMCPRunning changes (AC-02, AC-07)

---

## TASK-001-04-03: Implement ToggleMCPCommand with async start/stop

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | TASK-001-04-02 |
| Blocks | TASK-001-04-05 |

### What to do
- Add `[RelayCommand]` attribute on a new `private async Task ToggleMCPAsync()` method
- Use `AsyncRelayCommand` from CommunityToolkit.Mvvm for non-blocking UI execution
- Implementation logic:
  ```
  if (IsMCPRunning)
      call StopMCPAsync()
  else
      call StartMCPAsync()
  ```
- In `StartMCPAsync()`:
  - Clear `MCPErrorMessage`
  - Set `MCPStatus = "Starting..."`
  - Call `await _mcpServer.StartAsync()`
  - On success: set `IsMCPRunning = true`, `MCPStatus = "Running"`, `MCPPort = _mcpServer.Port`, `MCPStatusColor = Brushes.Green`
  - On failure: set error message (handled in TASK-001-04-05)
- In `StopMCPAsync()`:
  - Set `MCPStatus = "Stopping..."`
  - Call `await _mcpServer.StopAsync()`
  - Set `IsMCPRunning = false`, `MCPStatus = "Stopped"`, `MCPStatusColor = Brushes.Gray`
  - Clear `MCPErrorMessage`
- Follow the existing pattern from `MainViewModel.cs` methods `StartMCPServerAsync()` and `StopMCPServerAsync()`

### How to verify
- [ ] Clicking the toggle button when stopped starts the MCP server (AC-04)
- [ ] Clicking the toggle button when running stops the MCP server (AC-04)
- [ ] Toggle button label changes between "Start" and "Stop" (AC-07)

---

## TASK-001-04-04: Implement async monitoring of MCP SSE server state

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | TASK-001-04-02 |
| Blocks | None |

### What to do
- In the `DashboardViewModel` status polling timer (shared with US-001-01 TASK-001-01-04), add MCP status polling
- On each timer tick (every 1 second, matching `MainViewModel.cs` pattern), check `_mcpServer.IsRunning` and `_mcpServer.Port`
- If the state changed from the previous tick:
  - Update `IsMCPRunning`, `MCPStatus`, `MCPPort`, and `MCPStatusColor`
  - Log the state change via `_logger`
- This handles cases where the MCP server starts/stops externally (e.g., via command-line flag `--enable-mcp` in `App.xaml.cs`, or crashes)
- Alternatively, consider using a `DispatcherTimer` with 100ms interval matching the Python `on_mcp_sse_status_update` callback pattern for faster reaction, but 1-second polling is sufficient for status display
- Track previous MCP state in a `private bool _previousMCPRunning` field to detect transitions

### How to verify
- [ ] MCP status card updates automatically when the server starts or stops externally (AC-05)
- [ ] Status indicator transitions between green and gray in real-time (AC-03, AC-05)

---

## TASK-001-04-05: Add error handling for MCP server start failure

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-04-03 |
| Blocks | None |

### What to do
- In the `StartMCPAsync()` method (from TASK-001-04-03), wrap `_mcpServer.StartAsync()` in a try-catch block
- On `Exception ex`:
  - Set `MCPErrorMessage = $"MCP Server failed to start on port {_mcpServer.Port}. Port may be in use."`
  - Set `MCPStatus = "Error"`, `IsMCPRunning = false`, `MCPStatusColor = Brushes.Gray`
  - Log the error: `_logger?.LogError(ex, "Failed to start MCP Server on port {Port}", _mcpServer.Port)`
- In the `StopMCPAsync()` method, also wrap in try-catch and log errors but do not show error in UI (stopping failures are less critical)
- Clear `MCPErrorMessage` on any successful start or when the user retries (at the beginning of `StartMCPAsync`)
- The error message in the XAML card (from TASK-001-04-01) should be styled with `Foreground="Red"` or `Foreground="{StaticResource ErrorBrush}"` and `TextWrapping="Wrap"`

### How to verify
- [ ] When MCP server fails to start, the card displays "MCP Server failed to start on port {port}. Port may be in use." (AC-06)
- [ ] Error message clears when a subsequent start attempt succeeds (AC-06)

---

## TASK-001-04-06: Bind toggle button content and status indicator in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | S |
| Depends On | TASK-001-04-01 |
| Blocks | None |

### What to do
- In the MCP Server card (from TASK-001-04-01), ensure the toggle button `Content` is bound to `{Binding ToggleMCPButtonText}` so it dynamically shows "Start" or "Stop"
- Bind the status indicator `Ellipse.Fill` to `{Binding MCPStatusColor}` (direct Brush binding, consistent with AutoCAD/Odoo cards)
- Bind the port info `TextBlock.Text` to `{Binding MCPPortDisplay}`
- Add a `Visibility` binding on the port info TextBlock: `Visibility="{Binding IsMCPRunning, Converter={StaticResource BooleanToVisibilityConverter}}"` to hide port info when stopped
- Bind the error message `TextBlock.Text` to `{Binding MCPErrorMessage}` with `Visibility` collapsed when empty (use a `StringToVisibilityConverter` or `DataTrigger`)
- Ensure the `BoolToColorConverter` (from US-001-01 TASK-001-01-03) or direct Brush binding is consistently applied
- Style the toggle button to match the `PrimaryButton` style, with a visual distinction (e.g., red-tinted when showing "Stop")

### How to verify
- [ ] Toggle button shows "Start" when stopped and "Stop" when running (AC-07)
- [ ] Status indicator is green when running and gray when stopped (AC-03)

---

## TASK-001-04-07: Integrate with MCP SSE Manager service for server lifecycle management

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | L |
| Depends On | TASK-001-04-02, TASK-001-04-03, TASK-001-04-04 |
| Blocks | None |

### What to do
- Ensure `MCPSSEServer` is properly injected into `DashboardViewModel` via the DI container (it is already registered as singleton in `App.xaml.cs`)
- Update `DashboardViewModel` constructor to accept `MCPSSEServer mcpServer` parameter and store it
- In `App.xaml.cs` `ConfigureServices()`, update the `DashboardViewModel` registration to use a factory method that resolves all dependencies:
  ```csharp
  services.AddSingleton<DashboardViewModel>(sp => new DashboardViewModel(
      sp.GetService<IAutoCADService>(),
      sp.GetService<IOdooService>(),
      sp.GetRequiredService<MCPSSEServer>(),
      sp.GetRequiredService<INavigationService>(),
      sp.GetService<ILogger<DashboardViewModel>>()));
  ```
- Synchronize `DashboardViewModel` MCP status with `MainViewModel` MCP status to avoid duplicate or conflicting state:
  - Option A: `DashboardViewModel` reads directly from `MCPSSEServer.IsRunning` and `MCPSSEServer.Port` (preferred, single source of truth)
  - Option B: Use `WeakReferenceMessenger` to broadcast MCP state changes between ViewModels
- Verify the toggle command properly interacts with the `MCPSSEServer` lifecycle (`StartAsync`, `StopAsync`) and that the status timer correctly polls `IsRunning`
- Add integration test: mock `MCPSSEServer`, verify `ToggleMCPCommand` calls `StartAsync`/`StopAsync`, verify status properties update correctly
- Consider edge cases: what happens if the server is started via `MainViewModel.StartMCPServerCommand` while viewing the Dashboard? The polling timer should detect this and update the card

### How to verify
- [ ] Toggle button successfully starts and stops the MCP server via the MCPSSEServer service (AC-04)
- [ ] Status updates reflect the actual server state from MCPSSEServer.IsRunning (AC-05)
- [ ] Port number displayed matches MCPSSEServer.Port (AC-02)

---

## Dependency Graph
```
TASK-001-04-01 (MCP Card XAML) ──▶ TASK-001-04-06 (XAML Bindings)

TASK-001-04-02 (ViewModel Properties)
       │
       ├──▶ TASK-001-04-03 (Toggle Command) ──▶ TASK-001-04-05 (Error Handling) ──┐
       ├──▶ TASK-001-04-04 (Async Monitoring)                                      │
       │                                                                           ▼
       └──▶ TASK-001-04-07 (MCP Integration) ◀────────────────────────────────────┘
                (depends on 02, 03, and 04)
```
