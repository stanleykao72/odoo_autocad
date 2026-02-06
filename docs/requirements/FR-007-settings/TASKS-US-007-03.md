# TASKS: US-007-03 — Test Odoo Connection

> **Parent US**: [US-007-03](US-007-03-test-odoo-connection.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 4S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (Connection fields must be configured in SettingsViewModel before testing)
- [ ] `IOdooService` interface defined in `OdooAutoCAD.Core` with `TestConnectionAsync()` method signature

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a "Test Connection" button in the Connection tab
- [ ] AC-02: Clicking "Test Connection" calls `IOdooService.TestConnectionAsync()` with the current URL, database, username, and token values
- [ ] AC-03: While testing, a loading indicator is displayed and the button is disabled (IsTestingConnection = true)
- [ ] AC-04: On success, a green success message "Connected successfully" is displayed inline next to the button
- [ ] AC-05: On timeout failure, the message "Connection test timed out after {timeout} seconds. Check the server URL and network." is displayed
- [ ] AC-06: On authentication failure, the message "Authentication failed. Please verify your API token and username." is displayed and the token field is highlighted
- [ ] AC-07: On network error, the message "Cannot reach server at {url}. Check network connectivity and firewall settings." is displayed
- [ ] AC-08: The test result (ConnectionTestResult and ConnectionTestMessage) is displayed until the next test or until connection fields are modified
- [ ] AC-09: Test Connection button is disabled when required connection fields are empty or invalid

---

## TASK-007-03-01: Add "Test Connection" button with inline result display area to Connection tab XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-03-06 |

### What to do
- Add a `Button` labeled "Test Connection" in the Connection tab below the Odoo fields, bound to `{Binding TestConnectionCommand}`
- Add `IsEnabled="{Binding CanTestConnection}"` to disable when fields are invalid or test is in progress
- Add a `StackPanel Orientation="Horizontal"` next to the button containing: a `ProgressRing` or `ProgressBar` (for loading state), and a `TextBlock` bound to `ConnectionTestMessage`
- The result area should be visible only when `ConnectionTestMessage` is non-empty
- Style the button consistently with other action buttons in the Settings Page

### How to verify
- [ ] "Test Connection" button renders in the Connection tab (AC-01)
- [ ] Result display area is present next to the button (AC-04, AC-05, AC-06, AC-07)

---

## TASK-007-03-02: Implement TestConnectionCommand as IAsyncRelayCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-03-05, TASK-007-03-07 |

### What to do
- Add `IAsyncRelayCommand TestConnectionCommand` initialized as `new AsyncRelayCommand(TestConnectionAsync, CanTestConnection)`
- Implement `private async Task TestConnectionAsync()` that:
  1. Sets `IsTestingConnection = true` and `ConnectionTestResult = null`
  2. Calls `await _odooService.TestConnectionAsync(OdooServerUrl, OdooDatabaseName, OdooUsername, OdooApiToken)` with a `CancellationTokenSource` configured to the current `OdooTimeoutSeconds` value
  3. On success: sets `ConnectionTestResult = true`, `ConnectionTestMessage = "Connected successfully"`
  4. On exception: delegates to error classification logic (TASK-007-03-05)
  5. Sets `IsTestingConnection = false` in a `finally` block
- Implement `private bool CanTestConnection()` that returns false when `IsTestingConnection` is true or when required connection fields have validation errors

### How to verify
- [ ] Clicking the button calls `IOdooService.TestConnectionAsync()` with current field values (AC-02)
- [ ] Button is disabled during test execution (AC-03)
- [ ] Button is disabled when required fields are empty or invalid (AC-09)

---

## TASK-007-03-03: Add IsTestingConnection, ConnectionTestResult, and ConnectionTestMessage properties

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-03-02, TASK-007-03-06 |

### What to do
- Add `[ObservableProperty] bool _isTestingConnection` (default false) -- when true, the loading indicator is shown and button is disabled
- Add `[ObservableProperty] bool? _connectionTestResult` (default null) -- null=not tested, true=success, false=failure; used for XAML visual state triggers
- Add `[ObservableProperty] string _connectionTestMessage` (default empty) -- the text displayed in the inline result area
- When `IsTestingConnection` changes, call `TestConnectionCommand.NotifyCanExecuteChanged()` to update button enabled state

### How to verify
- [ ] `IsTestingConnection = true` shows loading indicator (AC-03)
- [ ] `ConnectionTestResult` supports three states: null, true, false (AC-04, AC-05, AC-06, AC-07)
- [ ] `ConnectionTestMessage` displays the appropriate text (AC-04, AC-05, AC-06, AC-07)

---

## TASK-007-03-04: Implement IOdooService.TestConnectionAsync method

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Core/Odoo/IOdooService.cs` and implementation |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-03-02 |

### What to do
- Add `Task<bool> TestConnectionAsync(string url, string database, string username, string token, CancellationToken cancellationToken = default)` to the `IOdooService` interface
- Implement in the concrete `OdooService` class: construct an `HttpClient` request to the Odoo server's version endpoint or authentication endpoint using the provided credentials
- Use `HttpClient` with the provided `CancellationToken` for timeout support
- Return `true` on successful response (HTTP 200 with valid JSON body)
- Throw specific exceptions for different failure types: `TaskCanceledException` for timeout, `HttpRequestException` with status code for auth/network errors
- Ensure the method does not modify any service state (read-only test)

### How to verify
- [ ] Method returns `true` when Odoo server responds with valid credentials (AC-02)
- [ ] Method throws `TaskCanceledException` when timeout elapses (AC-05)
- [ ] Method throws `HttpRequestException` with 401/403 for bad credentials (AC-06)

---

## TASK-007-03-05: Add error classification logic for timeout, auth failure, and network errors

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-03-02 |
| Blocks | None |

### What to do
- In the `catch` block of `TestConnectionAsync()`, classify the exception:
  - `TaskCanceledException` or `OperationCanceledException`: set `ConnectionTestMessage = $"Connection test timed out after {OdooTimeoutSeconds} seconds. Check the server URL and network."`
  - `HttpRequestException` where `StatusCode` is 401 or 403: set `ConnectionTestMessage = "Authentication failed. Please verify your API token and username."` and highlight the token field (set a flag like `IsTokenHighlighted = true`)
  - `HttpRequestException`, `SocketException`, or other network errors: set `ConnectionTestMessage = $"Cannot reach server at {OdooServerUrl}. Check network connectivity and firewall settings."`
- In all failure cases, set `ConnectionTestResult = false`
- Log the error details via `_logger.LogWarning()` for diagnostics

### How to verify
- [ ] Timeout exception shows timeout message with seconds (AC-05)
- [ ] Auth failure shows auth message and highlights token field (AC-06)
- [ ] Network error shows network message with server URL (AC-07)

---

## TASK-007-03-06: Add XAML data triggers for success (green), failure (red), and testing (spinner) visual states

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | TASK-007-03-01, TASK-007-03-03 |
| Blocks | None |

### What to do
- Add a `DataTrigger` on the result `TextBlock` binding to `ConnectionTestResult`:
  - When `true`: set `Foreground` to green (`#22C55E`), optionally show a checkmark icon
  - When `false`: set `Foreground` to red (`#EF4444`), optionally show an error icon
  - When `null`: collapse the result area entirely
- Add a `DataTrigger` on the `ProgressRing`/`ProgressBar` binding to `IsTestingConnection`:
  - When `true`: show the spinner, set `IsIndeterminate = True`
  - When `false`: collapse the spinner
- Ensure visual transitions are smooth (consider using `Storyboard` for fade-in/out)

### How to verify
- [ ] Success state shows green text with "Connected successfully" (AC-04)
- [ ] Failure state shows red text with error message (AC-05, AC-06, AC-07)
- [ ] Testing state shows a spinner/progress indicator (AC-03)

---

## TASK-007-03-07: Clear test result when connection fields are modified

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-03-02 |
| Blocks | None |

### What to do
- In the partial `On<Property>Changed` methods for `OdooServerUrl`, `OdooDatabaseName`, `OdooUsername`, `OdooApiToken`, and `OdooTimeoutSeconds`, reset the test result:
  - Set `ConnectionTestResult = null`
  - Set `ConnectionTestMessage = string.Empty`
- This ensures stale test results are cleared when the user modifies any connection field
- Also remove any token highlighting flag (`IsTokenHighlighted = false`) when fields change

### How to verify
- [ ] Modifying any connection field clears the previous test result display (AC-08)
- [ ] Result area is hidden after field modification until a new test is run (AC-08)

---

## Dependency Graph
```
TASK-007-03-01 (XAML Button & Result Area)
       │
       └──▶ TASK-007-03-06 (Visual State Triggers) ◀── TASK-007-03-03

TASK-007-03-03 (Test State Properties)
       │
       └──▶ TASK-007-03-02 (TestConnectionCommand) ◀── TASK-007-03-04
                  │
                  ├──▶ TASK-007-03-05 (Error Classification)
                  └──▶ TASK-007-03-07 (Clear on Field Change)

TASK-007-03-04 (IOdooService.TestConnectionAsync)
       │
       └──▶ TASK-007-03-02 (TestConnectionCommand)

Tasks 01, 03, and 04 can start in parallel.
```
