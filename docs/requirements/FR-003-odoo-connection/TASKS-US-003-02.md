# TASKS: US-003-02 — Test Connection

> **Parent US**: [US-003-02](US-003-02-test-connection.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 3S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (credential input form and validation) must be completed

## Acceptance Criteria
- [ ] AC-01: A "Test Connection" button is visible on the connection settings panel
- [ ] AC-02: Clicking "Test Connection" calls `TestConnectionAsync()` without persisting the session
- [ ] AC-03: On success, a green status message displays "Connection test successful" with server version info
- [ ] AC-04: On authentication failure (uid = false), the message displays "Authentication failed. Please verify your database name, username, and password."
- [ ] AC-05: On network failure (HttpRequestException), the message displays "Unable to reach Odoo server at {ServerUrl}. Please check your network connection and server URL."
- [ ] AC-06: On timeout (TaskCanceledException), the message displays "Connection timed out after {TimeoutSeconds} seconds."
- [ ] AC-07: On SSL/TLS error, the message displays "SSL certificate validation failed for {ServerUrl}."
- [ ] AC-08: The button is disabled while the test is in progress (shows a progress indicator)
- [ ] AC-09: The button is disabled when required credential fields are empty or invalid
- [ ] AC-10: The test operation does not modify the IsConnected state or establish a persistent session

---

## TASK-003-02-01: Add "Test Connection" button to XAML button bar

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-02-06 |

### What to do
- Add a `Button` labeled "Test Connection" in the connection settings button bar `StackPanel`
- Bind `Command` to `{Binding TestConnectionCommand}`
- Add a `ProgressBar IsIndeterminate="True"` or `ProgressRing` adjacent to the button, with `Visibility` bound to `{Binding TestConnectionCommand.IsRunning}` via `BoolToVisibilityConverter`
- Add a `TextBlock` for test result message bound to `{Binding TestResultMessage}` with `Foreground` bound to `{Binding TestResultColor}` (green for success, red for failure)
- Position the button before the "Connect" button in the button bar layout

### How to verify
- [ ] "Test Connection" button is visible on the connection settings panel (AC-01)
- [ ] Progress indicator appears while test is running (AC-08)

---

## TASK-003-02-02: Implement TestConnectionCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-02-03 |
| Blocks | TASK-003-02-05 |

### What to do
- Add `TestConnectionCommand` as `IAsyncRelayCommand` using `[RelayCommand(CanExecute = nameof(CanTestConnection))]` attribute on `async Task TestConnectionAsync()` method
- In `TestConnectionAsync()`:
  - Set `TestResultMessage = "Testing connection..."` and `TestResultColor = Brushes.Gray`
  - Create a temporary `OdooService` instance (or use a dedicated test method on the injected `IOdooService`) to avoid polluting the main service state
  - Call `TestConnectionAsync(ServerUrl.TrimEnd('/'), Database, Username, Password)` with a `CancellationTokenSource` governed by `TimeoutSeconds`
  - On success: set `TestResultMessage = "Connection test successful - Server version: {version}"` and `TestResultColor = Brushes.Green`
  - On failure: delegate to error handling (TASK-003-02-05)
- Add `bool CanTestConnection() => IsFormValid && !TestConnectionCommand.IsRunning`
- Add observable properties:
  - `private string _testResultMessage = string.Empty;`
  - `private Brush _testResultColor = Brushes.Gray;`
- Ensure the test does NOT set `IsConnected = true` or retain any session state after completion

### How to verify
- [ ] Clicking "Test Connection" calls TestConnectionAsync without persisting session (AC-02)
- [ ] Test operation does not modify IsConnected state (AC-10)
- [ ] On success, green message with server version info is displayed (AC-03)

---

## TASK-003-02-03: Implement TestConnectionAsync in OdooService (non-persistent)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | TASK-003-02-04 |
| Blocks | TASK-003-02-02 |

### What to do
- Refactor or add an overload `TestConnectionAsync(string serverUrl, string database, string username, string password)` in `OdooService`
- Implementation:
  - Create a separate `HttpClient` instance (or use the existing one without storing session state)
  - Call `POST /web/session/authenticate` with `{ db, login, password }` using `serverUrl.TrimEnd('/')`
  - Check response: if `result.uid` is `false` or missing, return a result indicating auth failure
  - On success, call `POST /web/webclient/version_info` to retrieve server version string
  - Return a result object containing: `Success`, `ServerVersion`, `ErrorMessage`, `ErrorType` (auth/network/timeout/ssl)
  - Do NOT store `_sessionId`, `_userId`, or set `_isConnected` -- discard all session data after test
- Use `CancellationToken` parameter for timeout support
- Dispose the temporary `HttpClient` after the test (or use `HttpClientFactory` pattern)

### How to verify
- [ ] TestConnectionAsync validates credentials without retaining session (AC-02, AC-10)
- [ ] Returns server version on success (AC-03)
- [ ] Returns appropriate error type on failure (AC-04, AC-05, AC-06, AC-07)

---

## TASK-003-02-04: Define TestConnectionAsync overload in IOdooService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/IOdooService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-02-03 |

### What to do
- Add a new record for test connection results:
  ```csharp
  public record TestConnectionResult(
      bool Success,
      string? ServerVersion,
      string? ErrorMessage,
      ConnectionErrorType ErrorType);

  public enum ConnectionErrorType
  {
      None,
      AuthenticationFailed,
      NetworkError,
      Timeout,
      SslError,
      Unknown
  }
  ```
- Add overloaded method to `IOdooService`:
  ```csharp
  Task<TestConnectionResult> TestConnectionAsync(
      string serverUrl, string database, string username, string password,
      CancellationToken cancellationToken = default);
  ```
- Keep the existing parameterless `Task<bool> TestConnectionAsync()` for backward compatibility
- Add XML doc comments explaining that the overload does not persist session state

### How to verify
- [ ] Interface declares TestConnectionAsync overload with credential parameters (AC-02)
- [ ] TestConnectionResult record includes Success, ServerVersion, ErrorMessage, ErrorType fields (AC-03, AC-04, AC-05, AC-06, AC-07)

---

## TASK-003-02-05: Add error handling with user-friendly messages for each failure scenario

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-02-02 |
| Blocks | None |

### What to do
- In the `TestConnectionAsync()` method, wrap the service call in a try-catch block with specific exception handling:
  - `HttpRequestException` with inner `AuthenticationException`: set `TestResultMessage = "SSL certificate validation failed for {ServerUrl}."` (AC-07)
  - `HttpRequestException` (general): set `TestResultMessage = "Unable to reach Odoo server at {ServerUrl}. Please check your network connection and server URL."` (AC-05)
  - `TaskCanceledException` / `OperationCanceledException`: set `TestResultMessage = "Connection timed out after {TimeoutSeconds} seconds."` (AC-06)
- For auth failure (from `TestConnectionResult.ErrorType == AuthenticationFailed`): set `TestResultMessage = "Authentication failed. Please verify your database name, username, and password."` (AC-04)
- Set `TestResultColor = Brushes.Red` for all failure scenarios
- Log the full exception via `_logger.LogError(ex, ...)` with structured logging including ServerUrl, Database, Username (but NOT password)
- Match error messages exactly to the Error Handling table in FR-003 Section 9

### How to verify
- [ ] Auth failure displays correct message (AC-04)
- [ ] Network failure displays correct message with ServerUrl (AC-05)
- [ ] Timeout displays correct message with TimeoutSeconds (AC-06)
- [ ] SSL error displays correct message (AC-07)

---

## TASK-003-02-06: Bind button IsEnabled to validation state and IsRunning

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | TASK-003-02-01 |
| Blocks | None |

### What to do
- Ensure the "Test Connection" button `Command` binding to `TestConnectionCommand` automatically disables the button when `CanTestConnection()` returns false or when the command is running (`IsRunning == true`)
- `IAsyncRelayCommand` from CommunityToolkit.Mvvm automatically handles `IsRunning` state -- verify that the button disables during execution
- In `OdooConnectionViewModel`, call `TestConnectionCommand.NotifyCanExecuteChanged()` in each field's `OnChanged` partial method to re-evaluate enabled state when fields change
- Ensure the button is disabled when required credential fields are empty or invalid (delegates to `IsFormValid` via `CanTestConnection`)

### How to verify
- [ ] Button is disabled while test is in progress (AC-08)
- [ ] Button is disabled when required credential fields are empty or invalid (AC-09)

---

## Dependency Graph
```
TASK-003-02-04 (IOdooService interface overload)
    |
    +---> TASK-003-02-03 (OdooService TestConnectionAsync implementation)
              |
              +---> TASK-003-02-02 (ViewModel TestConnectionCommand)
                        |
                        +---> TASK-003-02-05 (Error handling messages)

TASK-003-02-01 (XAML button)
    |
    +---> TASK-003-02-06 (Button IsEnabled binding)
```
