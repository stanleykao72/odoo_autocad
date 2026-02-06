# TASKS: US-002-07 — Monitor COM Status

> **Parent US**: [US-002-07](US-002-07-monitor-com-status.md)
> **Parent FR**: [FR-002](FR-002-autocad-connection.md)
> **Priority**: P2
> **Tasks**: 8 | **Effort**: 3S + 5M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] (None -- this US can be developed independently)

## Acceptance Criteria
- [ ] AC-01: The page displays a persistent connection status indicator showing "Connected" or "Disconnected" with a visual cue (e.g., green/red icon)
- [ ] AC-02: When connected, AutoCAD version and application info are displayed
- [ ] AC-03: All COM operations are routed through IGUIProxy to ensure STA thread safety
- [ ] AC-04: If a COM operation times out, the error message "AutoCAD operation timed out. The application may be busy." is displayed with a retry option
- [ ] AC-05: If a COM thread conflict occurs, the system transparently retries via the GUI proxy
- [ ] AC-06: Connection status is continuously monitored; if AutoCAD disconnects unexpectedly, the status updates to "Disconnected" and dependent controls are disabled
- [ ] AC-07: A "Refresh Status" button allows manual re-check of the COM connection health

---

## TASK-002-07-01: Add connection status indicator and AutoCAD version info to XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Connection Status panel of `AutoCADPage.xaml`, add a persistent status indicator:
  - An `Ellipse` (10x10 px) with `Fill` bound to `StatusIndicatorColor` (a `Brush` property):
    - Green (`#4CAF50`) when connected
    - Red (`#F44336`) when disconnected but previously connected (error state)
    - Gray (`#9E9E9E`) when never connected
  - A `TextBlock` bound to `ConnectionStatusText` (e.g., "Connected", "Disconnected")
  - Layout in a horizontal `StackPanel` with the ellipse to the left of the text
- Add AutoCAD version and application info display:
  - A `TextBlock` bound to `AutoCADVersionInfo`:
    - Format: "AutoCAD {Version}" (e.g., "AutoCAD 2024")
    - Visibility bound to `IsConnected` (only shown when connected)
- Place these at the very top of the Connection Status panel for immediate visibility

### How to verify
- [ ] Status indicator is always visible with green/red/gray visual cue (AC-01)
- [ ] AutoCAD version displayed when connected (AC-02)

---

## TASK-002-07-02: Add "Refresh Status" button to the Connection Status panel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a "Refresh Status" `Button` in the Connection Status panel:
  - Position it next to the Connect/Disconnect button
  - Bind `Command` to `RefreshStatusCommand`
  - Use an icon (refresh/sync icon) or text "Refresh Status"
  - The button should always be enabled (useful to re-check status in any state)
- Style consistently with other buttons in the panel
- Add a `ToolTip` explaining: "Re-check AutoCAD COM connection health"

### How to verify
- [ ] Refresh Status button is present and clickable (AC-07)
- [ ] Button is always enabled regardless of connection state

---

## TASK-002-07-03: Implement GetStatusAsync() in AutoCADService for connection health check

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-07-05 |

### What to do
- Verify and enhance the existing `GetStatusAsync()` method:
  - The method already exists and returns `AutoCADStatus` record
  - Ensure it correctly checks connection health by attempting to access `_acadApp.Name` and `_acadApp.Version`
  - If `_acadApp` is not null but accessing properties throws `COMException`, this means AutoCAD disconnected unexpectedly:
    - Set `_isConnected = false`
    - Return status with `IsConnected = false` and appropriate `ErrorMessage`
  - Add timeout handling: wrap COM property access in a `Task.Run` with `CancellationTokenSource` and timeout
  - If the COM access times out, return status with `ErrorMessage = "AutoCAD operation timed out. The application may be busy."`
- Register IGUIProxy handler (if not already done):
  - `"get_autocad_status"` action: calls `GetStatusAsync()`, returns `AutoCADStatus`
- Ensure the existing `AutoCADStatus` record has all needed fields:
  - `IsConnected`, `ApplicationName`, `Version`, `CurrentDocument`, `OpenDocuments`, `ErrorMessage` (all confirmed present)

### How to verify
- [ ] GetStatusAsync detects when AutoCAD has disconnected unexpectedly (AC-06)
- [ ] Timeout is handled with appropriate error message (AC-04)
- [ ] Version and app info are returned when connected (AC-02)

---

## TASK-002-07-04: Define GetStatusAsync() returning AutoCADStatus in IAutoCADService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify existing interface declaration:
  - `Task<AutoCADStatus> GetStatusAsync()` -- already present (line 116)
- Verify `AutoCADStatus` record has all necessary fields:
  - `bool IsConnected` -- connection state
  - `string? ApplicationName` -- e.g., "AutoCAD"
  - `string? Version` -- e.g., "24.1" or "2024"
  - `string? CurrentDocument` -- active document name
  - `List<string>? OpenDocuments` -- all open documents
  - `string? ErrorMessage` -- error description if any
  - All fields are confirmed present (lines 80-86)
- Consider adding a `DateTime LastChecked` field to `AutoCADStatus` for tracking when the status was last verified:
  - This requires extending the record. Use `with` expression support from C# records
- Add XML doc comments referencing FR-002-003, FR-002-004, FR-002-007

### How to verify
- [ ] AutoCADStatus record contains all required fields for status display (AC-01, AC-02)

---

## TASK-002-07-05: Implement RefreshStatusCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | TASK-002-07-03 |
| Blocks | None |

### What to do
- Implement `RefreshStatusCommand` if not already present (may overlap with TASK-002-04-04):
  ```csharp
  [RelayCommand]
  private async Task RefreshStatusAsync()
  ```
- Implementation:
  - Call `_guiProxy.ExecuteInGuiAsync("get_autocad_status")` to get fresh `AutoCADStatus`
  - Update `Status`, `IsConnected`, `ConnectionStatusText`, `StatusIndicatorColor`
  - If connected: update `DocumentName`, `DocumentPath`, `AutoCADVersionInfo`
  - If disconnected unexpectedly: update all dependent properties, disable controls, show notification
  - Set `AutoCADVersionInfo` to `$"{Status.ApplicationName} {Status.Version}"` when connected
- Add an `[ObservableProperty] private string _autoCADVersionInfo = string.Empty;` property
- Add an `[ObservableProperty] private Brush _statusIndicatorColor = Brushes.Gray;` property
- Update `StatusIndicatorColor` in `partial void OnIsConnectedChanged`:
  - Connected: `Brushes.Green` (or `new SolidColorBrush(Color.FromRgb(76, 175, 80))`)
  - Disconnected after was connected: `Brushes.Red`
  - Never connected: `Brushes.Gray`

### How to verify
- [ ] RefreshStatusCommand updates connection status and version info (AC-07)
- [ ] Status indicator color changes based on connection state (AC-01)

---

## TASK-002-07-06: Implement periodic connection health check using DispatcherTimer

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-07-08 |

### What to do
- Add a `DispatcherTimer` for periodic connection health monitoring:
  ```csharp
  private readonly DispatcherTimer _healthCheckTimer;
  ```
- Initialize in the `AutoCADViewModel` constructor:
  - Interval: 5 seconds (not 100ms -- that's for `IGUIProxy.ProcessRequests()`, which is separate)
  - The 100ms timer for `ProcessRequests()` should be in the main window code-behind or App startup
  - This health check timer is for periodic status polling only
- On each tick:
  - If `IsConnected`:
    - Attempt a lightweight status check via `_guiProxy.ExecuteInGuiAsync("get_autocad_status", timeout: 3000)`
    - If the response indicates disconnection, trigger `OnUnexpectedDisconnection()`
  - If not connected: skip check (no point polling when we know it's disconnected)
- Start the timer when `IsConnected` becomes true, stop when false
- Implement `partial void OnIsConnectedChanged(bool value)` to start/stop the timer
- Keep the health check lightweight -- only access `_acadApp.Name` to verify COM is alive

### How to verify
- [ ] Health check runs periodically when connected (AC-06)
- [ ] Unexpected disconnection is detected automatically (AC-06)

---

## TASK-002-07-07: Implement COM timeout detection and retry logic via IGUIProxy

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Threading/IGUIProxy.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- The existing `IGUIProxy.ExecuteInGuiAsync()` already supports a `timeout` parameter (default 10000ms)
- Enhance the timeout handling in the concrete `GUIProxy` implementation (not the interface):
  - When a request times out (`ProxyRequestStatus.TimedOut`):
    - Set `GUIProxyResponse.ErrorMessage = "AutoCAD operation timed out. The application may be busy."`
    - Set `GUIProxyResponse.ErrorType = "COMTimeoutError"`
- Add automatic retry logic for COM thread conflicts:
  - In the ViewModel layer, wrap IGUIProxy calls with retry:
    ```csharp
    private async Task<GUIProxyResponse> ExecuteWithRetryAsync(string action, Dictionary<string, object?>? parameters = null, int maxRetries = 3, int timeout = 10000)
    {
        for (int attempt = 0; attempt < maxRetries; attempt++)
        {
            var response = await _guiProxy.ExecuteInGuiAsync(action, parameters, timeout);
            if (response.Success || response.Status != ProxyRequestStatus.TimedOut)
                return response;
            _logger?.LogWarning("COM operation '{Action}' timed out, retry {Attempt}/{Max}", action, attempt + 1, maxRetries);
            await Task.Delay(500 * (attempt + 1)); // Exponential backoff
        }
        return GUIProxyResponse.CreateTimeout("final_timeout");
    }
    ```
- Add this retry helper method to `AutoCADViewModel` (or a shared base ViewModel class)
- COM thread conflicts should be transparently retried (AC-05) without user intervention

### How to verify
- [ ] Timeout produces "AutoCAD operation timed out" error message (AC-04)
- [ ] COM thread conflicts are retried transparently (AC-05)
- [ ] All COM operations route through IGUIProxy (AC-03)

---

## TASK-002-07-08: Handle unexpected disconnection with state update and notification

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | TASK-002-07-06 |
| Blocks | None |

### What to do
- Implement an `OnUnexpectedDisconnection()` method in `AutoCADViewModel`:
  ```csharp
  private void OnUnexpectedDisconnection()
  ```
- Implementation:
  - Set `IsConnected = false`
  - Set `ConnectionStatusText = "Disconnected (unexpected)"`
  - Set `StatusIndicatorColor = Brushes.Red` (not Gray, to indicate error vs. never-connected)
  - Clear runtime data: `Layouts`, `TableRows`, `LayoutAttributes`
  - Clear document info: `DocumentName = ""`, `DocumentPath = ""`
  - Clear PR info: `PRNumber = ""`, `ProjectName = ""`, `JobWorkingPlanName = ""`
  - Set `ErrorMessage = "AutoCAD connection lost unexpectedly. Please reconnect."`
  - Disable all COM-dependent commands by triggering `CanExecute` re-evaluation
  - Log the disconnection at Warning level
- Ensure all dependent `CanExecute` methods return `false` when `IsConnected == false` (VR-002-002)
- The user should see:
  - Red status indicator
  - Error message notification
  - All action buttons disabled
  - Connect button re-enabled (back to "Connect" label)

### How to verify
- [ ] Unexpected disconnection updates status to "Disconnected" with red indicator (AC-06)
- [ ] All dependent controls are disabled when disconnected unexpectedly (AC-06)
- [ ] User is notified of the disconnection (AC-06)

---

## Dependency Graph
```
TASK-002-07-01 (XAML status indicator)

TASK-002-07-02 (XAML Refresh Status button)

TASK-002-07-04 (IAutoCADService interface verification)

TASK-002-07-03 (GetStatusAsync health check)
    |
    +---> TASK-002-07-05 (RefreshStatusCommand)

TASK-002-07-06 (Periodic health check timer)
    |
    +---> TASK-002-07-08 (Unexpected disconnection handler)

TASK-002-07-07 (COM timeout/retry logic)
```
