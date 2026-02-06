# TASKS: US-003-03 — View Connection Status

> **Parent US**: [US-003-03](US-003-03-view-connection-status.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 4S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (credential input form) must be completed so the page structure exists

## Acceptance Criteria
- [ ] AC-01: A connection status panel is always visible on the Odoo Integration Page
- [ ] AC-02: The status displays a colored indicator: green Ellipse for connected, red/gray Ellipse for disconnected
- [ ] AC-03: Descriptive text shows current state: "Connected", "Disconnected", or "Connecting..."
- [ ] AC-04: The `IsConnected` property returns true only when a valid session ID exists and authentication has succeeded
- [ ] AC-05: When connected, the panel displays the authenticated username and database name
- [ ] AC-06: When disconnected, the panel shows an appropriate disconnected state message
- [ ] AC-07: Status updates reactively when connection state changes (connect, disconnect, session expiry)
- [ ] AC-08: Other pages (Dashboard, BOQ, PR) can subscribe to connection state changes to update their enabled state

---

## TASK-003-03-01: Create Connection Status panel in XAML with colored indicators

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-003-03-04 |

### What to do
- Add a `GroupBox` or `Border` panel with `Header="Connection Status"` below the Connection Settings section
- Layout using `Grid` with two columns (labels and values):
  - Row 0: "Status:" + `Ellipse` (Width=12, Height=12) with `Fill` bound to `{Binding IsConnected, Converter={StaticResource BoolToColorConverter}}` + `TextBlock` bound to `{Binding ConnectionStatusText}`
  - Row 1: "Database:" + `TextBlock` bound to `{Binding ConnectedDatabase}`
  - Row 2: "Username:" + `TextBlock` bound to `{Binding ConnectedUsername}`
  - Row 3: "Last Sync:" + `TextBlock` bound to `{Binding LastSyncTime, Converter={StaticResource DateTimeToStringConverter}}`
  - Row 4: "Session:" + `TextBlock` bound to `{Binding SessionStatusText}` (e.g., "Active" / "Expired" / "None")
- Use `Border` with rounded corners (`CornerRadius="4"`) and consistent theme from `App.xaml` resources
- Panel visibility is always `Visible` (not collapsed based on state)

### How to verify
- [ ] Connection status panel is always visible on the page (AC-01)
- [ ] Colored Ellipse indicator is present (AC-02)
- [ ] Database and Username labels are present for connected state display (AC-05)

---

## TASK-003-03-02: Implement IsConnected property with change notification

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-03-03 |

### What to do
- Add observable property using `[ObservableProperty]`:
  - `private bool _isConnected;`
- Implement `partial void OnIsConnectedChanged(bool value)` to:
  - Update `ConnectionStatusText` to "Connected" when `true`, "Disconnected" when `false`
  - Call `NotifyCanExecuteChanged()` on all connection-dependent commands: `ConnectCommand`, `DisconnectCommand`, `SyncProductsCommand`, `SearchProductsCommand`, `SearchProjectsCommand`, `SyncToOdooCommand`, `FilterByCategoryCommand`
  - Clear `ConnectedDatabase`, `ConnectedUsername`, `ServerVersion` when `false`
  - Raise a `PropertyChanged` event for `IsFormValid` and other computed properties
- The property value should be synchronized with `IOdooService.IsConnected` after each connect/disconnect operation

### How to verify
- [ ] IsConnected property returns true only after successful authentication (AC-04)
- [ ] Status updates reactively when connection state changes (AC-07)

---

## TASK-003-03-03: Implement ConnectionStatusText and display properties

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-03-02 |
| Blocks | None |

### What to do
- Add observable properties using `[ObservableProperty]`:
  - `private string _connectionStatusText = "Disconnected";`
  - `private string _connectedDatabase = string.Empty;`
  - `private string _connectedUsername = string.Empty;`
  - `private string _sessionStatusText = "None";`
- Update these properties in the `ConnectAsync()` command handler:
  - Before connection attempt: `ConnectionStatusText = "Connecting..."`
  - On success: `ConnectionStatusText = "Connected"`, populate `ConnectedDatabase` and `ConnectedUsername` from the service response
  - On failure: `ConnectionStatusText = "Disconnected"`
- In `DisconnectAsync()`: reset to "Disconnected" and clear database/username fields
- `SessionStatusText` should reflect: "Active" when connected, "Expired" if session becomes invalid, "None" when disconnected

### How to verify
- [ ] Descriptive text shows "Connected", "Disconnected", or "Connecting..." (AC-03)
- [ ] Connected state displays username and database name (AC-05)
- [ ] Disconnected state shows appropriate message (AC-06)

---

## TASK-003-03-04: Create BoolToColorConverter for status indicator binding

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Converters/BoolToColorConverter.cs` |
| Estimate | S |
| Depends On | TASK-003-03-01 |
| Blocks | None |

### What to do
- Create `BoolToColorConverter` class implementing `IValueConverter` in namespace `OdooAutoCAD.App.Converters`
- `Convert` method: return `Brushes.Green` when value is `true`, `Brushes.Gray` when `false`
- `ConvertBack` method: throw `NotSupportedException` (one-way binding only)
- Register as a `StaticResource` in `App.xaml` or `OdooConnectionPage.xaml` resources:
  ```xml
  <converters:BoolToColorConverter x:Key="BoolToColorConverter" />
  ```
- Optionally make it configurable with `TrueColor` and `FalseColor` dependency properties for reuse across pages (Dashboard uses green/gray, error states might use red)

### How to verify
- [ ] Green Ellipse displayed when connected, red/gray when disconnected (AC-02)

---

## TASK-003-03-05: Implement IsConnected property check in OdooService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify the existing `IsConnected` property implementation in `OdooService`:
  ```csharp
  public bool IsConnected => _isConnected && !string.IsNullOrEmpty(_sessionId);
  ```
- Currently `_sessionId` is never set in `ConnectAsync()` -- fix this: extract session ID from the authentication response cookies or response body
- After `ConnectAsync` succeeds, store the session ID from the `Set-Cookie` response header (Odoo returns `session_id` cookie)
- Configure `HttpClientHandler` with `CookieContainer` to automatically manage session cookies:
  ```csharp
  var handler = new HttpClientHandler { CookieContainer = new CookieContainer() };
  _httpClient = new HttpClient(handler);
  ```
- After `DisconnectAsync`, clear the cookie container and reset `_sessionId`

### How to verify
- [ ] IsConnected returns true only when valid session ID exists and auth succeeded (AC-04)

---

## TASK-003-03-06: Add connection state change event for cross-page reactivity

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/IOdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add an event to `IOdooService` interface:
  ```csharp
  event EventHandler<ConnectionStateChangedEventArgs>? ConnectionStateChanged;
  ```
- Define the event args class:
  ```csharp
  public class ConnectionStateChangedEventArgs : EventArgs
  {
      public bool IsConnected { get; }
      public string? ServerUrl { get; }
      public string? Database { get; }
      public string? Username { get; }
      public ConnectionStateChangedEventArgs(bool isConnected, string? serverUrl = null, string? database = null, string? username = null) { ... }
  }
  ```
- In `OdooService`, raise `ConnectionStateChanged` at the end of `ConnectAsync()` (on success) and `DisconnectAsync()`
- In `MainViewModel` and other ViewModels (BOQViewModel, SettingsViewModel), subscribe to `ConnectionStateChanged` to update their own `IsOdooConnected` property
- This enables Dashboard, BOQ, and PR pages to reactively disable Odoo-dependent operations when connection drops

### How to verify
- [ ] Status updates reactively when connection state changes (AC-07)
- [ ] Other pages can subscribe to connection state changes (AC-08)

---

## Dependency Graph
```
TASK-003-03-02 (IsConnected property)
    |
    +---> TASK-003-03-03 (ConnectionStatusText and display properties)

TASK-003-03-01 (XAML status panel)
    |
    +---> TASK-003-03-04 (BoolToColorConverter)

TASK-003-03-05 (OdooService IsConnected fix)

TASK-003-03-06 (Connection state change event)
```
