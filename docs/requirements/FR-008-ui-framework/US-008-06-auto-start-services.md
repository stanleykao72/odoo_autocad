# US-008-06: Auto-Start Services

## User Story
**As a** System Admin,
**I want to** have the MCP server and GUI proxy start automatically when the application launches,
**So that** AI assistant integration is available without manual intervention.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: On startup, `App.OnStartup` configures Serilog, builds the DI host, initializes the GUI proxy timer, and shows MainWindow.
- [ ] AC-02: When the `--enable-mcp` command-line flag is present, the MCP SSE server starts automatically during application startup.
- [ ] AC-03: The `--enable-mcp` flag is checked case-insensitively against the startup arguments.
- [ ] AC-04: The GUI proxy `DispatcherTimer` (100ms interval) is initialized and started after the DI host has fully started and `IGUIProxy` is resolved.
- [ ] AC-05: `IGUIProxy.Start()` is called before the DispatcherTimer begins ticking.
- [ ] AC-06: The timer tick handler calls `IGUIProxy.ProcessRequests()` on the WPF dispatcher (UI) thread, ensuring all COM operations are thread-safe.
- [ ] AC-07: If the MCP server fails to start, a status message is shown in the UI and the application continues running without MCP functionality.
- [ ] AC-08: DI services are registered in the correct order: Core threading (IGUIProxy), MCP services (MCPToolRegistry, MCPSSEServer), ViewModels, Application services (INavigationService).

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-039 | OnStartup: configure Serilog, build DI host, init GUI proxy timer, optionally start MCP, show MainWindow | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-06-01 | Implement App.OnStartup with Serilog configuration, DI host build, and MainWindow creation | `App.xaml.cs` | L |
| TASK-008-06-02 | Configure DI services registration in correct order (IGUIProxy, MCP, ViewModels, NavigationService) | `App.xaml.cs` | M |
| TASK-008-06-03 | Implement --enable-mcp flag parsing (case-insensitive) to conditionally start MCP server | `App.xaml.cs` | S |
| TASK-008-06-04 | Initialize DispatcherTimer (100ms) for GUI proxy polling after DI host start | `App.xaml.cs` | M |
| TASK-008-06-05 | Implement error handling for MCP server startup failure with StatusMessage feedback | `App.xaml.cs` | S |
| TASK-008-06-06 | Implement IGUIProxy interface with Start, Stop, ProcessRequests methods | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | M |

## Dependencies
- Depends on: US-008-05 (MainWindow must exist to be shown), US-008-09 (clean shutdown is the counterpart to startup)
- Blocks: None

## Notes
- The Python startup sequence is: set theme, set title/icon, setup geometry, setup style, create UI, setup log, init utilities, start GUI proxy processing (100ms timer), auto-start SSE server, update connection status. The C# equivalent consolidates this into `App.OnStartup` (Serilog, DI, timer) and `MainWindow` constructor (XAML layout, ViewModel binding, initial navigation).
- Validation rule VR-008-005 requires the GUI proxy timer must not start until the DI host has fully started and `IGUIProxy` is resolved.
- Validation rule VR-008-007 requires `--enable-mcp` flag checked case-insensitively.
- The `DispatcherTimer` runs on the WPF dispatcher (STA) thread, which is critical for COM interop operations that require single-threaded apartment access.
- The DI container uses `Microsoft.Extensions.Hosting` with `IHost` for structured service registration and lifecycle management.
