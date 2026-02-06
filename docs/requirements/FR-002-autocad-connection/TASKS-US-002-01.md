# TASKS: US-002-01 — Connect to AutoCAD

> **Parent US**: [US-002-01](US-002-01-connect-to-autocad.md)
> **Parent FR**: [FR-002](FR-002-autocad-connection.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 2S + 2M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-008-01 (navigation framework) must be completed

## Acceptance Criteria
- [ ] AC-01: Clicking the Connect button initiates a COM connection to AutoCAD
- [ ] AC-02: Connection first attempts GetActiveObject to attach to a running instance; if that fails, falls back to Dispatch to launch a new instance
- [ ] AC-03: Connection retries up to 5 times (1-second intervals) to obtain ActiveDocument
- [ ] AC-04: AutoCAD application is set to Visible upon successful connection
- [ ] AC-05: Connection status indicator updates to "Connected" with a visual indicator (e.g., green icon) on success
- [ ] AC-06: Connect button changes to "Disconnect" after a successful connection
- [ ] AC-07: All COM operations execute on the GUI/STA thread via IGUIProxy
- [ ] AC-08: If AutoCAD is not running and cannot be launched, an error message is displayed: "AutoCAD is not running. Please start AutoCAD and try again."
- [ ] AC-09: If connection fails after retries, an error message is displayed: "Could not access the active document after 5 attempts."

---

## TASK-002-01-01: Create AutoCAD page XAML with connection panel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-01-02, TASK-002-01-06 |

### What to do
- Create `AutoCADPage.xaml` as a WPF `Page` in namespace `OdooAutoCAD.App.Views.Pages`
- Add a Connection Status panel at the top with:
  - A status indicator (Ellipse or icon) bound to `IsConnected` via a `BoolToColorConverter` (green=connected, gray=disconnected)
  - A `TextBlock` bound to `Status.ApplicationName` and `Status.Version` for displaying AutoCAD info
  - A `TextBlock` bound to `DocumentName` and `DocumentPath`
- Add a Connect/Disconnect `Button` with `Content` toggled via `IsConnected` binding (show "Disconnect" when connected, "Connect" when not)
  - Bind `Command` to `ConnectCommand` when disconnected, `DisconnectCommand` when connected (use DataTrigger or converter)
- Add a "Refresh Status" `Button` bound to `RefreshStatusCommand`
- Set `DataContext` to `AutoCADViewModel` via DI or ViewModelLocator
- Include `xmlns:vm` namespace reference for design-time `d:DataContext`

### How to verify
- [ ] Connect/Disconnect button is visible and toggles label based on connection state (AC-05, AC-06)
- [ ] Status indicator displays green when connected, gray when disconnected (AC-05)

---

## TASK-002-01-02: Implement ConnectCommand and DisconnectCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | TASK-002-01-04, TASK-002-01-05 |
| Blocks | None |

### What to do
- Create `AutoCADViewModel` class in namespace `OdooAutoCAD.App.ViewModels`, inheriting `ObservableObject` from CommunityToolkit.Mvvm
- Inject `IAutoCADService`, `IGUIProxy`, and `ILogger<AutoCADViewModel>` via constructor
- Implement `ConnectCommand` as `IAsyncRelayCommand` using `[RelayCommand]` attribute on an `async Task ConnectAsync()` method:
  - Call `_guiProxy.ExecuteInGuiAsync("connect_autocad", timeout: 30000)`
  - On success: set `IsConnected = true`, populate `DocumentName`, `DocumentPath`, `Status`
  - On failure: display error message via a status property or IDialogService
  - Handle "AutoCAD is not running" and "Could not access after 5 attempts" error messages
- Implement `DisconnectCommand` as `IAsyncRelayCommand` using `[RelayCommand]` attribute on `async Task DisconnectAsync()`:
  - Call `_guiProxy.ExecuteInGuiAsync("disconnect_autocad")`
  - Set `IsConnected = false`, clear `DocumentName`, `DocumentPath`, `Status`
- Use `CanExecute` pattern: `ConnectCommand` enabled when `!IsConnected`, `DisconnectCommand` enabled when `IsConnected`

### How to verify
- [ ] ConnectCommand calls IGUIProxy.ExecuteInGuiAsync for thread-safe COM operation (AC-07)
- [ ] Error messages are set when connection fails (AC-08, AC-09)
- [ ] Connect/Disconnect commands toggle correctly based on IsConnected state (AC-01, AC-06)

---

## TASK-002-01-03: Implement ConnectAsync in AutoCADService with retry logic

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | L |
| Depends On | TASK-002-01-04 |
| Blocks | TASK-002-01-05 |

### What to do
- Refactor the existing `Connect()` private method to add retry logic for `ActiveDocument`:
  - Step 1: Try `GetActiveObject(DefaultProgId)` to get a running AutoCAD instance
  - Step 2 (fallback): Use `Activator.CreateInstance(Type.GetTypeFromProgID(DefaultProgId))` to launch a new instance
  - Step 3: Set `_acadApp.Visible = true`
  - Step 4: Retry loop up to 5 times with `Task.Delay(1000)` intervals to obtain `_acadApp.ActiveDocument`
  - Step 5: If `ActiveDocument` is null after 5 retries, throw with message "Could not access the active document after 5 attempts."
  - If no AutoCAD ProgID found, throw with message "AutoCAD is not running. Please start AutoCAD and try again."
- Extract `DocumentName` from `_acadDoc.Name` and `DocumentPath` from `_acadDoc.FullName` after successful connection
- Ensure `ConnectAsync()` wraps `Connect()` properly with `await Task.Run()`
- Add structured logging at each step for troubleshooting
- Handle `COMException` specifically for COM-related failures vs general exceptions

### How to verify
- [ ] Connection first tries GetActiveObject, then falls back to Dispatch (AC-02)
- [ ] Retries up to 5 times with 1-second intervals for ActiveDocument (AC-03)
- [ ] AutoCAD Visible is set to true on success (AC-04)

---

## TASK-002-01-04: Define ConnectAsync and DisconnectAsync in IAutoCADService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-01-02, TASK-002-01-03 |

### What to do
- Verify that the existing `IAutoCADService` interface already declares:
  - `Task<bool> ConnectAsync()` - connects to running or launches new AutoCAD instance
  - `Task DisconnectAsync()` - disconnects from AutoCAD
  - `Task<AutoCADStatus> GetStatusAsync()` - returns connection status and version info
  - `bool IsConnected { get; }` - current connection state
- These methods already exist in the interface (confirmed in skeleton). No new additions needed for US-002-01, but verify signatures match ViewModel expectations
- Ensure `AutoCADStatus` record includes `ApplicationName`, `Version`, `CurrentDocument` fields (already present)
- Add XML doc comments if missing, referencing FR-002-001 through FR-002-007

### How to verify
- [ ] Interface declares ConnectAsync, DisconnectAsync, GetStatusAsync, IsConnected (AC-01, AC-02)

---

## TASK-002-01-05: Wire COM operations through IGUIProxy for STA thread safety

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Threading/IGUIProxy.cs` |
| Estimate | M |
| Depends On | TASK-002-01-03 |
| Blocks | TASK-002-01-02 |

### What to do
- Register AutoCAD connection handlers with `IGUIProxy.RegisterHandler()`:
  - `"connect_autocad"` handler: calls `AutoCADService.ConnectAsync()` internally, marshalled to STA thread
  - `"disconnect_autocad"` handler: calls `AutoCADService.DisconnectAsync()` internally
  - `"get_autocad_status"` handler: calls `AutoCADService.GetStatusAsync()` internally
- Create a registration helper class `AutoCADProxyRegistration` in namespace `OdooAutoCAD.Core.AutoCAD`:
  - Static method `RegisterHandlers(IGUIProxy proxy, IAutoCADService service)` that registers all AutoCAD-related actions
  - Each handler is a `GUIProxyHandler` delegate taking `Dictionary<string, object?>` parameters
  - The `"connect_autocad"` handler should return an `AutoCADStatus` as the `Result` in `GUIProxyResponse`
- Ensure `ProcessRequests()` is called by the WPF `DispatcherTimer` at 100ms intervals (wiring done in App startup)
- All COM operations from ViewModel go through `_guiProxy.ExecuteInGuiAsync(actionName, params, timeout)`

### How to verify
- [ ] All AutoCAD COM operations are routed through IGUIProxy.ExecuteInGuiAsync (AC-07)
- [ ] connect_autocad, disconnect_autocad, get_autocad_status handlers are registered

---

## TASK-002-01-06: Implement connection status properties in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | TASK-002-01-01 |
| Blocks | None |

### What to do
- Add observable properties using `[ObservableProperty]` attribute from CommunityToolkit.Mvvm:
  - `private bool _isConnected;` -- overall connection state
  - `private AutoCADStatus? _status;` -- full status record from service
  - `private string _documentName = string.Empty;` -- active document filename
  - `private string _documentPath = string.Empty;` -- active document full path
  - `private string _errorMessage = string.Empty;` -- current error message for display
  - `private string _connectionStatusText = "Disconnected";` -- human-readable status label
- Implement `partial void OnIsConnectedChanged(bool value)` to:
  - Update `ConnectionStatusText` to "Connected" or "Disconnected"
  - Notify `ConnectCommand` and `DisconnectCommand` of `CanExecute` change
  - Enable/disable dependent commands (layout, extract, clear)
- Validation rule VR-002-001: Connect button only enabled when `!IsConnected`
- Validation rule VR-002-002: All COM-dependent controls disabled when `IsConnected == false`

### How to verify
- [ ] IsConnected, DocumentName, DocumentPath, Status are bindable observable properties (AC-05)
- [ ] ConnectionStatusText updates to "Connected"/"Disconnected" (AC-05)

---

## Dependency Graph
```
TASK-002-01-04 (IAutoCADService interface)
    |
    +---> TASK-002-01-03 (ConnectAsync retry logic)
    |         |
    |         +---> TASK-002-01-05 (IGUIProxy wiring)
    |                   |
    +---> TASK-002-01-02 (ViewModel commands)
              |
TASK-002-01-01 (XAML page)
    |
    +---> TASK-002-01-06 (Status properties)
```
