# TASKS: US-006-04 — Test Connection

> **Parent US**: [US-006-04](US-006-04-test-connection.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 5 | **Effort**: 3S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-006-01 (Start MCP Server - server must be running to test connection)
- [ ] `MCPSSEServer` running with `/health` and `/messages` endpoints exposed
- [ ] `MCPViewModel` created with `IsServerRunning` property

## Acceptance Criteria
- [ ] AC-01: A "Test Connection" button is available in the server control panel
- [ ] AC-02: Clicking the button sends a test_connection tool call to the running MCP server
- [ ] AC-03: The test result displays success or failure status with a timestamp
- [ ] AC-04: The test verifies both the HTTP endpoint (/health) and the JSON-RPC protocol layer
- [ ] AC-05: The Test Connection button is disabled when the server is not running
- [ ] AC-06: A timeout message "Connection test timed out. Server may be unresponsive." is shown if the test does not complete within the timeout period
- [ ] AC-07: A failure message "Connection test failed: {error}. Server may not be responding to JSON-RPC requests." is shown on error

---

## TASK-006-04-01: Implement TestConnectionCommand in MCPViewModel that calls health and JSON-RPC endpoints

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-04-05 |

### What to do
- Define `IAsyncRelayCommand TestConnectionCommand` using `[RelayCommand(CanExecute = nameof(CanTestConnection))]`
- Implement `CanTestConnection()` returning `IsServerRunning` (VR-006-006)
- In `TestConnection()`: create an `HttpClient` instance (or inject `IHttpClientFactory`)
- Phase 1 -- HTTP health check: `GET http://localhost:{ServerPort}/health`, validate response is 200 with valid JSON containing `status: "healthy"`
- Phase 2 -- JSON-RPC protocol test: `POST http://localhost:{ServerPort}/messages` with body `{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"test_connection","arguments":{}}}`, validate response contains `result` with `status: "connected"`
- On success: set `IsTestConnectionSuccess = true`, `TestConnectionResult = "Connection successful"`, `LastTestTime = DateTime.Now`
- On failure: set `IsTestConnectionSuccess = false`, `TestConnectionResult = "Connection test failed: {error}. Server may not be responding to JSON-RPC requests."`
- Use `HttpClient.Timeout` (e.g., 5 seconds) for timeout detection
- Add activity log entry for the test result

### How to verify
- [ ] Clicking button sends test_connection tool call (AC-02)
- [ ] Test verifies both /health and JSON-RPC layer (AC-04)
- [ ] Success result displays with timestamp (AC-03)

---

## TASK-006-04-02: Add Test Connection button to XAML with CanExecute bound to IsServerRunning

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `Button` in the Server Control panel with `Content="Test Connection"` and `Command="{Binding TestConnectionCommand}"`
- The button's `IsEnabled` is automatically managed by `IAsyncRelayCommand.CanExecute` (returns `false` when server is not running)
- Position the button after the Start/Stop and Restart buttons in the horizontal button panel
- Apply consistent button styling matching other buttons in the panel
- Add a visual busy indicator (small `ProgressRing` or `TextBlock` showing "Testing...") bound to `TestConnectionCommand.IsRunning`

### How to verify
- [ ] Test Connection button is present in server control panel (AC-01)
- [ ] Button is disabled when server is not running (AC-05)
- [ ] Button shows busy indicator while test is in progress (AC-01)

---

## TASK-006-04-03: Implement two-phase test: HTTP GET /health followed by JSON-RPC test_connection tool call

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Create a private method `TestHealthEndpointAsync()` that performs `GET http://localhost:{ServerPort}/health` and returns `(bool success, string details)`
- Parse the JSON response and verify: `status == "healthy"`, `tools_count > 0`, extract `active_connections` and `initialized` values
- Create a private method `TestJsonRpcEndpointAsync()` that performs `POST http://localhost:{ServerPort}/messages` with the `test_connection` JSON-RPC request
- Parse the JSON-RPC response and verify: `result` is present, contains `status: "connected"`, no `error` field
- Combine both results in `TestConnection()`: if Phase 1 fails, report HTTP-level failure; if Phase 2 fails, report protocol-level failure; if both succeed, report full success
- Include response details in the `TestConnectionResult` string (e.g., "Health: OK, Protocol: OK, Tools: 24, Connections: 1")
- Handle `TaskCanceledException` for timeout scenarios

### How to verify
- [ ] Health endpoint is checked first (AC-04)
- [ ] JSON-RPC protocol is verified second (AC-04)
- [ ] Combined result provides meaningful diagnostic info (AC-03)

---

## TASK-006-04-04: Display test results in the status panel

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a test result section below the action buttons in the Server Control panel
- Include: a success/failure icon (checkmark or X) bound to `IsTestConnectionSuccess` using a `BoolToIconConverter`
- Include: a `TextBlock` bound to `TestConnectionResult` for the detail message
- Include: a `TextBlock` bound to `LastTestTime` formatted as "Last tested: HH:mm:ss" using `StringFormat`
- Apply color coding: green text for success, red text for failure, using `DataTrigger` on `IsTestConnectionSuccess`
- Set `Visibility` to `Collapsed` when `LastTestTime` is null (no test has been performed yet)

### How to verify
- [ ] Success/failure status is displayed with icon (AC-03)
- [ ] Timestamp of last test is displayed (AC-03)
- [ ] Color coding distinguishes success from failure (AC-03)

---

## TASK-006-04-05: Handle timeout and failure scenarios with appropriate user-facing error messages

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-04-01 |
| Blocks | None |

### What to do
- In `TestConnection()`, wrap the HTTP calls in `try-catch` blocks for specific exception types
- Catch `TaskCanceledException` (timeout): set `TestConnectionResult = "Connection test timed out. Server may be unresponsive."`, `IsTestConnectionSuccess = false`
- Catch `HttpRequestException` (connection refused/network error): set `TestConnectionResult = "Connection test failed: {ex.Message}. Server may not be responding to JSON-RPC requests."`, `IsTestConnectionSuccess = false`
- Catch generic `Exception`: set `TestConnectionResult = "Connection test failed: {ex.Message}. Server may not be responding to JSON-RPC requests."`
- Log all test results (success and failure) to the activity log with Source="Server"
- Set `LastTestTime = DateTime.Now` on every test completion (success or failure)
- Suggest restart in timeout scenarios by appending " Consider restarting the server." to the message

### How to verify
- [ ] Timeout shows correct message (AC-06)
- [ ] Failure shows correct message with error details (AC-07)
- [ ] All test outcomes are logged in activity log (AC-03)

---

## Dependency Graph
```
TASK-006-04-01 (TestConnectionCommand)
       │
       └──▶ TASK-006-04-05 (Error Handling)

TASK-006-04-02 (XAML Button)
       (independent)

TASK-006-04-03 (Two-Phase Test Logic)
       (independent)

TASK-006-04-04 (XAML Test Result Display)
       (independent)

Tasks 01-04 are independent and can be developed in parallel.
Task 05 depends on 01 (command must exist to add error handling logic).
```
