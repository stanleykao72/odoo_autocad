# TASKS: US-008-09 — Clean Shutdown

> **Parent US**: [US-008-09](US-008-09-clean-shutdown.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 4S + 2M + 0L
> **Status**: Done

## Prerequisites
- [x] None (shutdown logic can be implemented independently, though US-008-06 auto-start is the counterpart)

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

---

## TASK-008-09-01: Implement App.OnExit with ordered cleanup sequence

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-008-09-05 |

### What to do
- Override `App.OnExit(ExitEventArgs e)` with the following ordered cleanup:
  1. Log "Application shutting down"
  2. Stop `_guiProxyTimer?.Stop()` to prevent new proxy calls
  3. Stop MCP server: resolve `MCPSSEServer` from DI, check `IsRunning`, call `StopAsync()` if running (wrapped in try-catch per TASK-008-09-05)
  4. Stop GUI proxy: resolve `IGUIProxy` from DI, call `Stop()` to cancel pending requests
  5. Stop DI host: `await _host.StopAsync()` then `_host.Dispose()`
  6. Flush Serilog: `await Log.CloseAndFlushAsync()`
  7. Call `base.OnExit(e)`
- Each step must be wrapped in its own try-catch to ensure subsequent steps execute even if earlier ones fail
- The cleanup order is critical: timer first, then MCP (release ports), then proxy (stop processing), then host (dispose services), then logs (flush remaining)

### How to verify
- [x] Timer is stopped first (AC-01)
- [x] MCP server is stopped if running (AC-02)
- [x] GUI proxy is stopped (AC-03)
- [x] DI host is stopped and disposed (AC-04)
- [x] Serilog is flushed and closed (AC-05)

---

## TASK-008-09-02: Register DispatcherUnhandledException handler

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `App.OnStartup`, register `DispatcherUnhandledException += OnDispatcherUnhandledException;` early, before any other initialization
- Implement `OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)`:
  - Log the exception: `Log.Error(e.Exception, "Unhandled dispatcher exception")`
  - Show a `MessageBox` with the error message for user awareness
  - Set `e.Handled = true` to prevent application crash and allow continuation
- Verify the skeleton already implements this; confirm `e.Handled = true` is set

### How to verify
- [x] DispatcherUnhandledException handler is registered (AC-06)
- [x] Error is logged (AC-06)
- [x] Exception is marked handled so the application continues (AC-06)

---

## TASK-008-09-03: Register AppDomain.UnhandledException handler

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `App.OnStartup`, register `AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;` early
- Implement `OnUnhandledException(object sender, UnhandledExceptionEventArgs e)`:
  - Cast `e.ExceptionObject` to `Exception` if possible
  - Log as fatal: `Log.Fatal(ex, "Unhandled domain exception")`
  - No UI display (this handler runs on non-UI threads and cannot safely show UI)
- Verify the skeleton already implements this; confirm logging level is `Fatal`

### How to verify
- [x] AppDomain.UnhandledException handler is registered (AC-07)
- [x] Error is logged as fatal (AC-07)

---

## TASK-008-09-04: Register TaskScheduler.UnobservedTaskException handler

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `App.OnStartup`, register `TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;`
- Implement `OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)`:
  - Log the exception: `Log.Error(e.Exception, "Unobserved task exception")`
  - Call `e.SetObserved()` to prevent the exception from escalating and terminating the process
- Verify the skeleton already implements this; confirm `SetObserved()` is called

### How to verify
- [x] TaskScheduler.UnobservedTaskException handler is registered (AC-08)
- [x] Error is logged (AC-08)
- [x] `SetObserved()` is called to prevent escalation (AC-08)

---

## TASK-008-09-05: Add try-catch around MCP server stop to prevent shutdown blocking

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | TASK-008-09-01 |
| Blocks | None |

### What to do
- In `App.OnExit`, wrap the MCP server stop call in a try-catch block:
  ```csharp
  try
  {
      var mcpServer = Services.GetService<MCPSSEServer>();
      if (mcpServer?.IsRunning == true)
      {
          await mcpServer.StopAsync();
      }
  }
  catch (Exception ex)
  {
      Log.Error(ex, "Failed to stop MCP server during shutdown");
  }
  ```
- The catch block must not rethrow or call `Shutdown` -- it logs and continues
- Ensure subsequent cleanup steps (GUI proxy stop, host stop, log flush) execute regardless of MCP stop failure
- Verify the skeleton handles this; the current implementation lacks try-catch around the MCP stop

### How to verify
- [x] MCP stop failure is caught and logged (AC-09)
- [x] Shutdown continues without blocking when MCP stop fails (AC-09)
- [x] No orphan processes or ports remain (AC-10)

---

## TASK-008-09-06: Verify no port locks remain after shutdown via integration test

| Field | Value |
|-------|-------|
| Target | `tests/` (integration test) |
| Estimate | M |
| Depends On | TASK-008-09-01, TASK-008-09-05 |
| Blocks | None |

### What to do
- Create an integration test that starts the application with `--enable-mcp`, verifies the MCP port is open, then triggers shutdown
- After shutdown completes, verify the MCP port (8084) is no longer in use
- Use `TcpClient` or `TcpListener` to probe the port availability after shutdown
- Test that the GUI proxy has zero pending requests after shutdown
- Test that `IHost` is disposed and no longer accepting service resolutions
- This test may need to run as a separate process or use `AppDomain` isolation

### How to verify
- [x] No ports remain locked after application exit (AC-10)
- [x] No orphan processes remain after application exit (AC-10)

---

## Dependency Graph

```
TASK-008-09-01 (OnExit cleanup sequence)
    └── TASK-008-09-05 (MCP try-catch)
        └── TASK-008-09-06 (Integration test)

TASK-008-09-02 (Dispatcher handler) [independent]

TASK-008-09-03 (AppDomain handler) [independent]

TASK-008-09-04 (TaskScheduler handler) [independent]
```
