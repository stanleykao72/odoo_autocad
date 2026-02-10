# US-008-09: Clean Shutdown

## User Story
**As a** System Admin,
**I want to** have all resources cleaned up properly when I close the application,
**So that** no orphan processes or port locks are left behind.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [x] AC-01: On exit, `App.OnExit` stops the GUI proxy DispatcherTimer.
- [x] AC-02: On exit, `App.OnExit` stops the MCP SSE server if it is running.
- [x] AC-03: On exit, `App.OnExit` stops the GUI proxy (`IGUIProxy.Stop()`).
- [x] AC-04: On exit, `App.OnExit` stops the DI host (`IHost.StopAsync()`).
- [x] AC-05: On exit, `App.OnExit` flushes and closes the Serilog logger (`Log.CloseAndFlush()`).
- [x] AC-06: The application registers a global exception handler for `DispatcherUnhandledException` that logs the error and marks it handled so the application can continue.
- [x] AC-07: The application registers a global exception handler for `AppDomain.UnhandledException` that logs the error as fatal.
- [x] AC-08: The application registers a global exception handler for `TaskScheduler.UnobservedTaskException` that logs the error and calls `SetObserved()`.
- [x] AC-09: If the MCP server fails to stop during shutdown, the error is logged and shutdown continues without blocking.
- [x] AC-10: No orphan processes or locked ports remain after the application exits.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-040 | OnExit: stop GUI proxy timer, stop MCP server, stop GUI proxy, stop DI host, flush Serilog | Must |
| FR-008-041 | Register global exception handlers for Dispatcher, AppDomain, and TaskScheduler exceptions | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-09-01 | Implement App.OnExit with ordered cleanup: stop timer, stop MCP, stop GUI proxy, stop host, flush logs | `App.xaml.cs` | M |
| TASK-008-09-02 | Register DispatcherUnhandledException handler that logs and marks handled | `App.xaml.cs` | S |
| TASK-008-09-03 | Register AppDomain.UnhandledException handler that logs fatal error | `App.xaml.cs` | S |
| TASK-008-09-04 | Register TaskScheduler.UnobservedTaskException handler that logs and calls SetObserved | `App.xaml.cs` | S |
| TASK-008-09-05 | Add try-catch around MCP server stop to prevent shutdown blocking on failure | `App.xaml.cs` | S |
| TASK-008-09-06 | Verify no port locks remain after shutdown via integration test | `Tests/` | M |

## Dependencies
- Depends on: None (shutdown logic can be implemented independently)
- Blocks: None directly, but US-008-06 (auto-start) references this as its counterpart

## Notes
- The Python implementation uses `on_closing()` which stops MCP/SSE servers and then calls `self.destroy()`. The C# port uses `App.OnExit` for a more structured cleanup sequence.
- The cleanup order is important: timer must stop first (to prevent new proxy calls), then MCP server (to release ports), then GUI proxy (to stop processing), then DI host (to dispose all services), and finally Serilog (to flush remaining log entries).
- Error handling during shutdown must be defensive: each cleanup step should be wrapped in try-catch to ensure subsequent steps execute even if earlier ones fail.
- The Python `WM_DELETE_WINDOW` protocol maps to the WPF `Window.Closing` event or `App.OnExit` override.
- Global exception handlers should be registered early in `App.OnStartup` before any other initialization to catch startup errors.
