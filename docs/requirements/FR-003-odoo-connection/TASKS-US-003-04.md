# TASKS: US-003-04 — Disconnect from Odoo

> **Parent US**: [US-003-04](US-003-04-disconnect-from-odoo.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 4S + 1M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (credential input form) must be completed
- [ ] US-003-03 (connection status display and state change event) must be completed

## Acceptance Criteria
- [ ] AC-01: A "Disconnect" button is visible on the connection settings panel
- [ ] AC-02: Clicking "Disconnect" calls `DisconnectAsync()` and terminates the current Odoo session
- [ ] AC-03: After disconnection, session state is cleared (session cookies, uid, session ID)
- [ ] AC-04: The `IsConnected` property returns false after disconnection
- [ ] AC-05: The connection status indicator changes to red/gray with text "Disconnected"
- [ ] AC-06: Server info fields (version, database, username in status panel) are cleared
- [ ] AC-07: Product sync, project search, and sync-to-Odoo buttons are disabled after disconnection
- [ ] AC-08: The "Disconnect" button is disabled when already disconnected
- [ ] AC-09: Credential input fields remain populated so the user can reconnect or modify and reconnect

---

## TASK-003-04-01: Add "Disconnect" button to XAML button bar

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-04-05 |

### What to do
- Add a `Button` labeled "Disconnect" in the connection settings button bar `StackPanel`, positioned after the "Connect" button
- Bind `Command` to `{Binding DisconnectCommand}`
- The button should use a distinct style (e.g., warning/secondary style) to visually differentiate from "Connect"
- Style should be consistent with the application theme defined in `App.xaml`

### How to verify
- [ ] "Disconnect" button is visible on the connection settings panel (AC-01)

---

## TASK-003-04-02: Implement DisconnectCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-04-03 |
| Blocks | TASK-003-04-04 |

### What to do
- Add `DisconnectCommand` as `IAsyncRelayCommand` using `[RelayCommand(CanExecute = nameof(CanDisconnect))]` attribute on `async Task DisconnectAsync()` method
- In `DisconnectAsync()`:
  - Call `_odooService.DisconnectAsync()`
  - Set `IsConnected = false` (triggers `OnIsConnectedChanged` which updates status text and disables dependent commands)
  - Do NOT clear credential input fields (`ServerUrl`, `Database`, `Username`, `Password`) -- keep them populated for reconnection
- Add `bool CanDisconnect() => IsConnected` to disable the button when already disconnected
- Call `DisconnectCommand.NotifyCanExecuteChanged()` whenever `IsConnected` changes

### How to verify
- [ ] Clicking "Disconnect" calls DisconnectAsync (AC-02)
- [ ] IsConnected returns false after disconnection (AC-04)
- [ ] Credential input fields remain populated after disconnect (AC-09)

---

## TASK-003-04-03: Implement DisconnectAsync in OdooService with full state cleanup

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-003-04-02 |

### What to do
- Refactor the existing `DisconnectAsync()` method in `OdooService` to perform complete state cleanup:
  - Clear `_sessionId = null`
  - Clear `_userId = null`
  - Set `_isConnected = false`
  - Clear the `CookieContainer` on the `HttpClientHandler` to remove session cookies:
    ```csharp
    if (_httpClient.DefaultRequestHeaders.Contains("Cookie"))
        _httpClient.DefaultRequestHeaders.Remove("Cookie");
    // Or if using CookieContainer: create a new handler
    ```
  - Optionally attempt to call Odoo's session destroy endpoint (`/web/session/destroy`) in a try-catch (graceful; ignore failures if server is unreachable)
- Raise `ConnectionStateChanged` event with `IsConnected = false`
- Ensure the method is idempotent: calling `DisconnectAsync()` when already disconnected should not throw exceptions
- Log disconnection event via `_logger.LogInformation("Disconnected from Odoo")`
- Do NOT clear `_serverUrl`, `_database`, `_username` internal fields -- these are used for display/reconnection reference

### How to verify
- [ ] Session state is cleared: session cookies, uid, session ID (AC-03)
- [ ] DisconnectAsync is graceful; does not throw if server is unreachable (AC-02)
- [ ] ConnectionStateChanged event is raised (AC-07)

---

## TASK-003-04-04: Reset ViewModel status properties on disconnect

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-04-02 |
| Blocks | None |

### What to do
- In the `DisconnectAsync()` command handler (after service call), reset the following ViewModel properties:
  - `ServerVersion = string.Empty;`
  - `ConnectedDatabase = string.Empty;`
  - `ConnectedUsername = string.Empty;`
  - `SessionStatusText = "None";`
  - `ConnectionStatusText = "Disconnected";`
  - `ErrorMessage = string.Empty;`
- Do NOT reset:
  - `ServerUrl`, `Database`, `Username`, `Password` (input fields remain populated per AC-09)
  - `LastSyncTime` (last sync info should persist for user reference)
  - `Products`, `ProjectSearchResults` (cached data remains for viewing)
- The `OnIsConnectedChanged(false)` partial method should handle disabling dependent commands

### How to verify
- [ ] Connection status indicator changes to red/gray with "Disconnected" text (AC-05)
- [ ] Server info fields are cleared (AC-06)
- [ ] Credential inputs remain populated (AC-09)

---

## TASK-003-04-05: Bind Disconnect button IsEnabled to IsConnected

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | TASK-003-04-01 |
| Blocks | None |

### What to do
- The `DisconnectCommand` binding automatically handles `IsEnabled` via the `CanDisconnect()` delegate
- Verify that the `Command="{Binding DisconnectCommand}"` binding correctly disables the button when `CanDisconnect()` returns `false` (i.e., when `IsConnected == false`)
- No additional `IsEnabled` binding is needed since `IAsyncRelayCommand` handles this via `CanExecute`
- Ensure `NotifyCanExecuteChanged()` is called whenever `IsConnected` changes to refresh the button state

### How to verify
- [ ] "Disconnect" button is disabled when already disconnected (AC-08)

---

## TASK-003-04-06: Fire connection state change event on disconnect

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `OdooService.DisconnectAsync()`, after clearing state, raise the `ConnectionStateChanged` event:
  ```csharp
  ConnectionStateChanged?.Invoke(this, new ConnectionStateChangedEventArgs(
      isConnected: false,
      serverUrl: _serverUrl,
      database: _database,
      username: _username));
  ```
- Ensure the event is raised regardless of whether the optional server-side session destroy succeeds or fails
- Other ViewModels subscribed to this event (MainViewModel, BOQViewModel) should update their own enabled states
- This ensures product sync, project search, and sync-to-Odoo buttons on other pages are disabled after disconnection

### How to verify
- [ ] Product sync, project search, and sync-to-Odoo buttons are disabled after disconnection (AC-07)
- [ ] Connection state change event fires to notify subscribers (AC-07)

---

## Dependency Graph
```
TASK-003-04-03 (OdooService DisconnectAsync cleanup)
    |
    +---> TASK-003-04-02 (ViewModel DisconnectCommand)
              |
              +---> TASK-003-04-04 (Reset status properties)

TASK-003-04-01 (XAML Disconnect button)
    |
    +---> TASK-003-04-05 (Button IsEnabled binding)

TASK-003-04-06 (Connection state change event)
```
