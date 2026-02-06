# TASKS: US-008-06 — Auto-Start Services

> **Parent US**: [US-008-06](US-008-06-auto-start-services.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 2S + 3M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-008-05 completed (MainWindow must exist to be shown)
- [ ] US-008-09 in progress or designed (clean shutdown is the counterpart to startup)

## Acceptance Criteria
- [ ] AC-01: On startup, `App.OnStartup` configures Serilog, builds the DI host, initializes the GUI proxy timer, and shows MainWindow.
- [ ] AC-02: When the `--enable-mcp` command-line flag is present, the MCP SSE server starts automatically during application startup.
- [ ] AC-03: The `--enable-mcp` flag is checked case-insensitively against the startup arguments.
- [ ] AC-04: The GUI proxy `DispatcherTimer` (100ms interval) is initialized and started after the DI host has fully started and `IGUIProxy` is resolved.
- [ ] AC-05: `IGUIProxy.Start()` is called before the DispatcherTimer begins ticking.
- [ ] AC-06: The timer tick handler calls `IGUIProxy.ProcessRequests()` on the WPF dispatcher (UI) thread, ensuring all COM operations are thread-safe.
- [ ] AC-07: If the MCP server fails to start, a status message is shown in the UI and the application continues running without MCP functionality.
- [ ] AC-08: DI services are registered in the correct order: Core threading (IGUIProxy), MCP services (MCPToolRegistry, MCPSSEServer), ViewModels, Application services (INavigationService).

---

## TASK-008-06-01: Implement App.OnStartup with Serilog configuration, DI host build, and MainWindow creation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-008-06-02, TASK-008-06-03, TASK-008-06-04 |

### What to do
- In `App.OnStartup`, call `base.OnStartup(e)` first
- Register global exception handlers early (before any other init) for `DispatcherUnhandledException`, `AppDomain.CurrentDomain.UnhandledException`, `TaskScheduler.UnobservedTaskException`
- Configure Serilog with `LoggerConfiguration().MinimumLevel.Debug().WriteTo.Console().WriteTo.File(...)` and create logger
- Build `IHost` with `Host.CreateDefaultBuilder().UseSerilog().ConfigureServices(ConfigureServices).Build()`
- Set `Services = _host.Services` static property for service location
- Call `await _host.StartAsync()` to start the DI host
- Call `InitializeGUIProxyTimer()` after host starts (per VR-008-005)
- Create and show `MainWindow`
- Wrap entire block in try-catch: on failure, `Log.Fatal()`, show `MessageBox`, call `Shutdown(1)`
- Verify the skeleton already implements this; refine error handling and ordering

### How to verify
- [ ] App.OnStartup configures Serilog, builds DI host, inits timer, shows window (AC-01)
- [ ] DI host is fully started before GUI proxy timer begins (AC-04)
- [ ] Startup failure shows MessageBox and exits with code 1 (AC-01)

---

## TASK-008-06-02: Configure DI services registration in correct order

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | M |
| Depends On | TASK-008-06-01 |
| Blocks | None |

### What to do
- In `ConfigureServices(IServiceCollection services)`, register in documented order:
  1. Core threading: `services.AddSingleton<IGUIProxy, GUIProxy>()`
  2. MCP services: `services.AddSingleton<MCPToolRegistry>(sp => new MCPToolRegistry(...))` with factory pattern for controlled construction, `services.AddSingleton<MCPSSEServer>(sp => new MCPSSEServer(...))`
  3. ViewModels: `services.AddTransient<MainViewModel>()`, `services.AddTransient<ConnectionStatusViewModel>()`, `services.AddTransient<BOQViewModel>()`, `services.AddTransient<SettingsViewModel>()`
  4. Application services: `services.AddSingleton<INavigationService, NavigationService>()`
- Use optional service resolution (`sp.GetService<T>()`) for `IAutoCADService`, `IOdooService`, `IBOQProcessor` that are not yet implemented
- Verify the skeleton already has this registration; confirm ordering matches AC-08

### How to verify
- [ ] DI services registered in correct order: threading, MCP, ViewModels, services (AC-08)
- [ ] Optional services use `GetService<T>()` (nullable) not `GetRequiredService<T>()` (AC-08)

---

## TASK-008-06-03: Implement --enable-mcp flag parsing with case-insensitive check

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | TASK-008-06-01 |
| Blocks | TASK-008-06-05 |

### What to do
- In `App.OnStartup`, after GUI proxy timer initialization, check for `--enable-mcp` flag in `e.Args`
- Use case-insensitive comparison per VR-008-007: `e.Args.Any(a => a.Equals("--enable-mcp", StringComparison.OrdinalIgnoreCase))`
- If flag is present, call `await StartMCPServerAsync()`
- The current skeleton uses `e.Args.Contains("--enable-mcp")` which is case-sensitive; fix to use case-insensitive comparison
- Log a message when MCP auto-start is triggered

### How to verify
- [ ] `--enable-mcp` flag triggers MCP server auto-start (AC-02)
- [ ] Flag comparison is case-insensitive (AC-03)
- [ ] `--ENABLE-MCP` and `--Enable-MCP` variants are also accepted (AC-03)

---

## TASK-008-06-04: Initialize DispatcherTimer for GUI proxy polling after DI host start

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | M |
| Depends On | TASK-008-06-01 |
| Blocks | None |

### What to do
- Implement `InitializeGUIProxyTimer()` method that resolves `IGUIProxy` from `Services.GetRequiredService<IGUIProxy>()`
- Call `guiProxy.Start()` before creating the timer (per AC-05, VR-008-005)
- Create `_guiProxyTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) }`
- Set tick handler: `_guiProxyTimer.Tick += (s, e) => { guiProxy.ProcessRequests(); };`
- Call `_guiProxyTimer.Start()` to begin polling
- The timer runs on the WPF dispatcher (STA/UI) thread, ensuring all COM operations via the proxy are thread-safe
- Wrap the tick handler body in try-catch to prevent timer from stopping on individual request failures
- Log "GUI Proxy timer initialized with 100ms interval"

### How to verify
- [ ] Timer is initialized after DI host starts and IGUIProxy is resolved (AC-04)
- [ ] `IGUIProxy.Start()` is called before timer starts (AC-05)
- [ ] Tick handler calls `ProcessRequests()` on the UI thread (AC-06)
- [ ] Timer interval is 100ms (AC-04)

---

## TASK-008-06-05: Implement error handling for MCP server startup failure

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | TASK-008-06-03 |
| Blocks | None |

### What to do
- In `StartMCPServerAsync()`, wrap `mcpServer.StartAsync()` in try-catch
- On `Exception`, log with `Log.Error(ex, "Failed to start MCP SSE Server")`
- Resolve `MainViewModel` from DI and set `StatusMessage` to an error message (e.g., `$"MCP Server failed to start: {ex.Message}"`)
- Ensure the application continues running without MCP functionality (do not call `Shutdown` or rethrow)
- Verify the skeleton already implements basic error handling; enhance with StatusMessage feedback

### How to verify
- [ ] MCP startup failure is logged (AC-07)
- [ ] StatusMessage shows error feedback in the UI (AC-07)
- [ ] Application continues running without MCP (AC-07)

---

## TASK-008-06-06: Implement IGUIProxy interface with Start, Stop, ProcessRequests methods

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Threading/IGUIProxy.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-008-06-04 |

### What to do
- Define `IGUIProxy` interface in `OdooAutoCAD.Core.Threading` namespace
- Add `bool IsRunning { get; }` property
- Add `int PendingRequestCount { get; }` property
- Add `void Start()` method to begin accepting requests
- Add `void Stop()` method to cancel pending requests and stop processing
- Add `int ProcessRequests()` method returning the number of requests processed (called by DispatcherTimer)
- Add `Task<GUIProxyResponse> ExecuteInGuiAsync(string action, Dictionary<string, object?>? parameters, int timeout)` for async execution from background threads
- Add `void RegisterHandler(string action, GUIProxyHandler handler)` for action registration
- Verify the skeleton already has this interface; confirm all required members per `GUIProxy.cs` implementation

### How to verify
- [ ] IGUIProxy interface defines Start, Stop, ProcessRequests methods (AC-04, AC-05, AC-06)
- [ ] Interface supports async execution pattern for MCP thread-to-UI-thread bridging (AC-06)

---

## Dependency Graph

```
TASK-008-06-01 (App.OnStartup)
    ├── TASK-008-06-02 (DI registration)
    ├── TASK-008-06-03 (--enable-mcp flag)
    │   └── TASK-008-06-05 (MCP error handling)
    └── TASK-008-06-04 (DispatcherTimer init)

TASK-008-06-06 (IGUIProxy interface) ──> TASK-008-06-04 (DispatcherTimer init)
```
