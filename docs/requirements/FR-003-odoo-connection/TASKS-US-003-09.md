# TASKS: US-003-09 — View Server Info

> **Parent US**: [US-003-09](US-003-09-view-server-info.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 4S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (credential input form) must be completed
- [ ] US-003-03 (connection status panel) must be completed so the display area exists

## Acceptance Criteria
- [ ] AC-01: On successful connection, a status panel displays the Odoo server version (e.g., "17.0")
- [ ] AC-02: The status panel displays the connected database name
- [ ] AC-03: The status panel displays the authenticated username
- [ ] AC-04: Server version is retrieved via `POST /web/webclient/version_info` after authentication
- [ ] AC-05: The server info panel is only populated when `IsConnected` is true
- [ ] AC-06: Server info fields are cleared when the user disconnects
- [ ] AC-07: The session status ("Active" / "Expired") is displayed alongside the connection info

---

## TASK-003-09-01: Add server info labels to Connection Status panel in XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Connection Status panel (created in US-003-03), add or verify the following rows:
  - "Server Version:" + `TextBlock` bound to `{Binding ServerVersion}`
  - "Database:" + `TextBlock` bound to `{Binding ConnectedDatabase}`
  - "Username:" + `TextBlock` bound to `{Binding ConnectedUsername}`
  - "Session:" + `TextBlock` bound to `{Binding SessionStatusText}`
- Each `TextBlock` should have a fallback value (empty string) when the binding value is null
- Group these fields within the existing Connection Status `Border` or `GroupBox`
- Layout should match the wireframe: "Status: (o) Connected, Server Version: 17.0, Database: production, Username: admin, Session: Active"

### How to verify
- [ ] Server version label is visible in the status panel (AC-01)
- [ ] Database name label is visible (AC-02)
- [ ] Username label is visible (AC-03)
- [ ] Session status label is visible (AC-07)

---

## TASK-003-09-02: Add ServerVersion, ConnectedDatabase, ConnectedUsername properties

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-09-05 |

### What to do
- Add observable properties using `[ObservableProperty]` (may already exist from US-003-03; verify and add if missing):
  - `private string _serverVersion = string.Empty;`
  - `private string _connectedDatabase = string.Empty;`
  - `private string _connectedUsername = string.Empty;`
- These properties should:
  - Be populated after successful `ConnectAsync()` with data from `OdooStatus` record
  - Be cleared after `DisconnectAsync()` (handled in US-003-04)
  - Default to empty string when not connected
- The `OdooStatus` record contains: `IsConnected`, `ServerUrl`, `Database`, `Username`, `Version`, `ErrorMessage`

### How to verify
- [ ] ServerVersion, ConnectedDatabase, ConnectedUsername are bindable properties (AC-01, AC-02, AC-03)
- [ ] Properties are only populated when connected (AC-05)

---

## TASK-003-09-03: Implement GetStatusAsync in OdooService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-003-09-04 |

### What to do
- Verify the existing `GetStatusAsync()` implementation:
  ```csharp
  public async Task<OdooStatus> GetStatusAsync()
  {
      if (!IsConnected) return new OdooStatus(false, _serverUrl, _database, _username, null, "Not connected");
      // ... calls /web/webclient/version_info
  }
  ```
- The existing implementation correctly calls `/web/webclient/version_info` and extracts `server_version`
- Add session validity check: if the version_info call fails with a session error, update `IsConnected` to false and return disconnected status
- Add a `string? SessionId` property to `OdooStatus` or track session validity internally
- Consider caching the version info to avoid repeated API calls (version does not change during a session)
- Handle `JsonException` if the response format is unexpected

### How to verify
- [ ] GetStatusAsync returns OdooStatus with server version, database, username (AC-01, AC-02, AC-03)
- [ ] Returns disconnected status when not connected (AC-05)

---

## TASK-003-09-04: Call version_info after successful authentication

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | TASK-003-09-03 |
| Blocks | None |

### What to do
- In `ConnectAsync()`, after successful authentication (uid is valid), automatically call `GetStatusAsync()` to retrieve server version:
  ```csharp
  if (_isConnected)
  {
      var status = await GetStatusAsync();
      _serverVersion = status.Version;
  }
  ```
- Store the version internally in `_serverVersion` private field
- Add a public `string? ServerVersion` property to `OdooService` that returns the cached version
- Alternatively, return the `OdooStatus` from `ConnectAsync()` so the ViewModel has immediate access:
  - Change return type from `Task<bool>` to `Task<OdooStatus>` or add an out parameter
  - Or have the ViewModel call `GetStatusAsync()` separately after connect succeeds
- The version_info endpoint is `POST /web/webclient/version_info` with `{ "jsonrpc": "2.0", "method": "call", "params": {}, "id": 1 }`
- Extract `result.server_version` from the response (e.g., "17.0")

### How to verify
- [ ] Server version is retrieved via POST /web/webclient/version_info after authentication (AC-04)
- [ ] Version is available for display immediately after connection (AC-01)

---

## TASK-003-09-05: Populate server info after ConnectCommand execution

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-09-02 |
| Blocks | TASK-003-09-06 |

### What to do
- In the `ConnectAsync()` command handler, after successful connection:
  ```csharp
  var status = await _odooService.GetStatusAsync();
  ServerVersion = status.Version ?? "Unknown";
  ConnectedDatabase = status.Database ?? Database;
  ConnectedUsername = status.Username ?? Username;
  SessionStatusText = "Active";
  IsConnected = true;
  ConnectionStatusText = "Connected";
  ```
- If connection fails, ensure these fields remain empty and `IsConnected` stays false
- Add a `ConnectCommand` using `[RelayCommand(CanExecute = nameof(CanConnect))]` if not already present:
  - `async Task ConnectAsync()` method
  - `bool CanConnect() => IsFormValid && !IsConnected`
  - Calls `_odooService.ConnectAsync(ServerUrl.TrimEnd('/'), Database, Username, Password)`

### How to verify
- [ ] Server version displayed after successful connection (AC-01)
- [ ] Database name displayed (AC-02)
- [ ] Username displayed (AC-03)
- [ ] Fields only populated when connected (AC-05)

---

## TASK-003-09-06: Clear server info on DisconnectCommand execution

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-09-05 |
| Blocks | None |

### What to do
- In the `DisconnectAsync()` command handler (or `OnIsConnectedChanged(false)`), clear server info:
  ```csharp
  ServerVersion = string.Empty;
  ConnectedDatabase = string.Empty;
  ConnectedUsername = string.Empty;
  SessionStatusText = "None";
  ```
- This may already be handled in US-003-04 TASK-003-04-04 -- verify and ensure no duplication
- If handled in `OnIsConnectedChanged(false)`, this task is confirming the behavior exists
- Ensure the XAML bindings show empty/default text when these properties are empty strings

### How to verify
- [ ] Server info fields are cleared on disconnect (AC-06)
- [ ] Session status shows "None" after disconnect (AC-07)

---

## Dependency Graph
```
TASK-003-09-03 (OdooService GetStatusAsync)
    |
    +---> TASK-003-09-04 (version_info after auth)

TASK-003-09-02 (ViewModel server info properties)
    |
    +---> TASK-003-09-05 (Populate after connect)
              |
              +---> TASK-003-09-06 (Clear on disconnect)

TASK-003-09-01 (XAML server info labels)
```
