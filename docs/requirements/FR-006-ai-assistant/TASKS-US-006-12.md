# TASKS: US-006-12 — Configure Server Port

> **Parent US**: [US-006-12](US-006-12-configure-server-port.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 6 | **Effort**: 4S + 1M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure must be in place)
- [ ] `MCPViewModel` created with `ServerPort` property (from US-006-01)
- [ ] Configuration persistence mechanism available (YAML config or application settings)

## Acceptance Criteria
- [ ] AC-01: The page displays the currently configured server port (default: 8084)
- [ ] AC-02: The port field is editable when the server is stopped
- [ ] AC-03: The port field is disabled (read-only) when the server is running
- [ ] AC-04: Port input is validated to be an integer between 1024 and 65535
- [ ] AC-05: Invalid port values display a validation error message and prevent the server from starting
- [ ] AC-06: Changing the port automatically updates all displayed endpoint URLs (SSE, Health, Messages)
- [ ] AC-07: The configured port is persisted so it survives application restarts

---

## TASK-006-12-01: Add editable port TextBox to Configuration panel XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-006-12-06 |

### What to do
- In the Configuration panel "Server" subsection, replace the static port display TextBlock with an editable `TextBox`
- Bind `Text="{Binding ServerPort, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"` for real-time binding
- Bind `IsEnabled="{Binding IsServerRunning, Converter={StaticResource InverseBoolConverter}}"` to disable editing when server is running (VR-006-005)
- Alternative: use `IsReadOnly` binding instead of `IsEnabled` for a different visual treatment
- Set `MaxLength="5"` to limit input length (max port 65535 is 5 digits)
- Apply `InputScope` or `PreviewTextInput` handler to filter non-numeric input
- Style the TextBox consistently with other input fields in the application
- Add a label "Port:" to the left of the TextBox

### How to verify
- [ ] Port field displays current port value (AC-01)
- [ ] Field is editable when server is stopped (AC-02)
- [ ] Field is disabled when server is running (AC-03)

---

## TASK-006-12-02: Implement port validation in MCPViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Implement `INotifyDataErrorInfo` on `MCPViewModel` (CommunityToolkit.Mvvm supports this via `ObservableValidator`)
- Change `MCPViewModel` base class to `ObservableValidator` instead of `ObservableObject`
- Add validation attributes to the `ServerPort` property: `[Range(1024, 65535, ErrorMessage = "Port must be between 1024 and 65535")]`
- Override or use `partial void OnServerPortChanging(int value)` to validate: if value < 1024 or > 65535, set a validation error
- Add `[ObservableProperty] string _portValidationError` (empty when valid, error message when invalid)
- In the `partial void OnServerPortChanged(int value)` method: call `ValidateAllProperties()` or manually set `PortValidationError`
- Prevent `StartServerCommand` from executing when port is invalid: update `CanStartServer()` to include `&& string.IsNullOrEmpty(PortValidationError)`
- Handle non-integer input: if the TextBox binding fails to parse, the WPF binding engine shows a red border by default; ensure `ValidatesOnExceptions=True` is set on the binding

### How to verify
- [ ] Port validation rejects values outside 1024-65535 (AC-04)
- [ ] Validation error message is set for invalid ports (AC-05)
- [ ] Server cannot start with invalid port (AC-05)

---

## TASK-006-12-03: Update computed endpoint URL properties when ServerPort changes

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify that `partial void OnServerPortChanged(int value)` is implemented (from US-006-06)
- In the method, call `OnPropertyChanged(nameof(SSEEndpointUrl))`, `OnPropertyChanged(nameof(HealthEndpointUrl))`, `OnPropertyChanged(nameof(MessagesEndpointUrl))`, `OnPropertyChanged(nameof(RootEndpointUrl))`
- This ensures all computed URL properties recalculate and notify XAML bindings when the port changes
- The computed properties use `$"http://localhost:{ServerPort}/sse"` etc.
- Also update the client guidance text if it contains the port number
- Test: changing port from 8084 to 9090 should immediately update all displayed URLs

### How to verify
- [ ] SSE endpoint URL updates when port changes (AC-06)
- [ ] Health endpoint URL updates when port changes (AC-06)
- [ ] Messages endpoint URL updates when port changes (AC-06)
- [ ] All URLs update in real-time (AC-06)

---

## TASK-006-12-04: Persist port configuration to application settings

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | None |

### What to do
- Choose a persistence mechanism compatible with the application's existing configuration system:
  - Option A: Use `Properties.Settings.Default` (built-in .NET settings) -- add `MCPServerPort` setting of type `int` with default 8084
  - Option B: Use a YAML configuration file (consistent with the Python project's `config/*.yaml` pattern) -- read/write `mcp_server_port` key
  - Option C: Use the SQLite database via the existing configuration caching system
- In `MCPViewModel` constructor: load the persisted port value and set `ServerPort`
- In `partial void OnServerPortChanged(int value)`: if the new value is valid (passes validation), persist it immediately
- For Option A: `Properties.Settings.Default.MCPServerPort = value; Properties.Settings.Default.Save();`
- For Option B: read/write YAML using the existing config utility, key path `mcp.server.port`
- Handle first-run scenario: if no persisted value exists, use default 8084
- Handle corrupted/invalid persisted value: fall back to default 8084 with a warning log

### How to verify
- [ ] Port value persists across application restarts (AC-07)
- [ ] Default port is 8084 on first run (AC-07)
- [ ] Invalid persisted values fall back to default (AC-07)

---

## TASK-006-12-05: Add port availability check on server start and display error if port is in use

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add a private method `bool IsPortAvailable(int port)` to `MCPSSEServer`:
  - Create a `System.Net.Sockets.TcpListener` on `IPAddress.Loopback` with the given port
  - Call `listener.Start()` -- if it succeeds, the port is available; call `listener.Stop()` to release
  - Catch `SocketException` -- port is in use, return false
- In `StartAsync()`, before building the `WebApplication`, call `IsPortAvailable(Port)`
- If not available: throw `InvalidOperationException($"Port {Port} is already in use. Please choose a different port or stop the conflicting application.")`
- The ViewModel catches this exception in `StartServerCommand` and displays the error message to the user
- Log the port availability check result via `_logger`
- This implements VR-006-002 (port must not be in use before starting)

### How to verify
- [ ] Port availability is checked before server start (AC-04, AC-05)
- [ ] Port-in-use error message is descriptive (AC-05)
- [ ] Server remains stopped when port is unavailable (AC-05)

---

## TASK-006-12-06: Display validation error styling for invalid port input

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | TASK-006-12-01 |
| Blocks | None |

### What to do
- Add a `Validation.ErrorTemplate` to the port TextBox for visual error indication:
  - Red border around the TextBox when validation fails
  - Error icon (exclamation mark) next to the field
- Add a `TextBlock` below the port TextBox bound to `{Binding PortValidationError}` with red text, visible only when the error message is non-empty
- Use `Style.Triggers` or `DataTrigger` to change the TextBox `BorderBrush` to red when `PortValidationError` is not empty
- Add `PreviewTextInput` handler in code-behind to filter non-numeric characters: `e.Handled = !int.TryParse(e.Text, out _)` (basic numeric filtering)
- Also handle paste: add `DataObject.Pasting` handler to filter pasted non-numeric content
- When the user types an out-of-range value, the validation error should appear immediately (due to `UpdateSourceTrigger=PropertyChanged`)

### How to verify
- [ ] Red border appears for invalid port values (AC-05)
- [ ] Error message text appears below the field (AC-05)
- [ ] Non-numeric characters are filtered (AC-04)
- [ ] Validation feedback is immediate on typing (AC-04)

---

## Dependency Graph
```
TASK-006-12-01 (XAML Port TextBox)
       │
       └──▶ TASK-006-12-06 (Validation Error Styling)

TASK-006-12-02 (Port Validation Logic)
       (independent)

TASK-006-12-03 (Endpoint URL Update on Port Change)
       (independent)

TASK-006-12-04 (Port Persistence)
       (independent)

TASK-006-12-05 (Port Availability Check)
       (independent)

Tasks 01-05 are independent and can be developed in parallel.
Task 06 depends on 01 (TextBox must exist for validation styling).
```
