# FR-006: AI Assistant (MCP) Page

> **Document Version**: 1.0
> **Last Updated**: 2026-02-06
> **Status**: Not Started
> **Priority**: P2

## 1. Overview

The AI Assistant (MCP) Page provides a graphical interface for managing the Model Context Protocol (MCP) Server-Sent Events (SSE) server that enables AI assistants (such as Gemini CLI or Claude) to interact with AutoCAD and Odoo systems through a standardized protocol. The page exposes server lifecycle controls (start, stop, restart), real-time status monitoring, connection testing, registered tool inspection, configuration display, and an activity log viewer.

The MCP architecture follows a layered design:

1. **Transport Layer**: SSE (Server-Sent Events) over HTTP provides the real-time communication channel. The server listens on a configurable port (default 8084) and exposes four HTTP endpoints: `GET /` (server info), `GET /health` (health check), `GET /sse` (SSE connection stream), and `POST /messages` (JSON-RPC 2.0 message handling).

2. **Protocol Layer**: JSON-RPC 2.0 carries MCP protocol messages. The server handles `initialize`, `initialized`, `tools/list`, `tools/call`, and `ping` methods. The protocol version is `2024-11-05`.

3. **Tool Layer**: A registry of 18+ tools provides AI assistants with structured access to AutoCAD COM operations (parameter extraction, layout management, drawing primitives, layer control, element scanning) and Odoo ERP operations (status checking, data synchronization, BOQ generation, purchase requisition workflows).

4. **Thread Safety Layer**: A GUI Proxy system bridges the MCP server thread (MTA) and the GUI/COM thread (STA) using a message queue architecture. All AutoCAD COM operations are queued from the MCP server thread and executed on the GUI main thread via a DispatcherTimer polling at 100ms intervals, ensuring thread-safe COM interop.

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-006-01 | CAD Engineer | Start the MCP SSE server from the UI | AI assistants can connect and help me with AutoCAD operations |
| US-006-02 | CAD Engineer | Stop the MCP SSE server | I can free the port and resources when AI assistance is not needed |
| US-006-03 | CAD Engineer | See the current server status (Running/Stopped) | I know whether AI assistants can connect |
| US-006-04 | System Admin | Test the MCP connection | I can verify the server is responding correctly to JSON-RPC requests |
| US-006-05 | System Admin | View the list of registered MCP tools | I can confirm which operations are available to AI assistants |
| US-006-06 | System Admin | See the server port and endpoint configuration | I can configure AI assistant clients to connect to the correct address |
| US-006-07 | CAD Engineer | Monitor active SSE connections | I can see how many AI assistant clients are currently connected |
| US-006-08 | System Admin | View server activity logs | I can troubleshoot issues with AI assistant operations |
| US-006-09 | CAD Engineer | See which tools require AutoCAD or Odoo connections | I know which prerequisites must be met before using AI assistant features |
| US-006-10 | System Admin | Restart the MCP server | I can recover from server issues without restarting the entire application |
| US-006-11 | CAD Engineer | See real-time tool execution results | I can monitor what the AI assistant is doing in AutoCAD |
| US-006-12 | System Admin | Configure the server port | I can avoid port conflicts with other applications |

## 3. Python Reference

### Source Files
- `mcp_server_fastmcp.py` (~2,312 lines) - FastMCP SSE server with 18+ MCP tool implementations
- `utility/util_mcp_sse_manager.py` (~925 lines) - MCPSSEManager class for GUI integration, tool registration, status management
- `utility/util_gui_proxy.py` (~605 lines) - GUI Proxy system for thread-safe COM operations via message queue
- `utility/util_mcp_sse_server.py` - StandardMCPSSEServer low-level implementation
- `forms/form_main_modern.py` - SSE control panel UI integration in main form

### Key Classes

**MCPSSEManager** (`util_mcp_sse_manager.py`)
- `__init__(port=8084, autocad_util=None, odoo_util=None)` - Initialize with port and shared utility instances
- `start_server() -> bool` - Start the SSE server in a background thread
- `stop_server() -> bool` - Stop the SSE server and clean up connections
- `restart_server() -> bool` - Stop then start (with 1-second delay)
- `get_server_status() -> Dict` - Returns `is_running`, `port`, `tools_count`, etc.
- `test_mcp_connection() -> Dict` - Tests JSON-RPC connectivity
- `set_status_callback(callback)` - Register GUI callback for status updates; signature `callback(message: str, is_running: bool)`
- `update_autocad_status_cache()` - Refreshes cached AutoCAD state (document name, version, layouts)
- `cleanup()` - Stop server and release resources

**GUIProxy** (`util_gui_proxy.py`)
- `register_handler(action_name, handler_func)` - Register a handler for an action
- `execute_in_gui(action, **kwargs) -> Dict` - Queue a request and wait for response (10s timeout)
- `process_requests() -> int` - Process pending requests from queue (called from GUI thread)
- Global `get_gui_proxy()` - Singleton accessor
- Global `setup_gui_proxy_handlers(autocad_util, logger)` - Registers all 8 default handlers

### MCP Tools (Complete Registry)

**Connection & Status Tools** (from `mcp_server_fastmcp.py`):
1. `test_connection()` - Test MCP connection, returns status string
2. `get_server_info()` - Returns name, version, tools_count, capabilities, timestamp
3. `check_autocad_status()` - Check AutoCAD COM connection; returns connected/version/document info
4. `check_odoo_status()` - Check Odoo REST API connection; returns connected/server_url/database info

**AutoCAD Drawing Tools** (from `mcp_server_fastmcp.py`):
5. `create_new_drawing(drawing_name, template_path, units, save_path)` - Create new DWG file
6. `draw_line(start_point, end_point, layer)` - Draw a line between two 3D points
7. `draw_circle(center_point, radius, layer)` - Draw a circle with center and radius
8. `create_text(position, text_content, height, rotation, layer, style, alignment)` - Create text annotation
9. `add_dimension(dimension_type, definition_points, text_position, text_override, dim_style, layer, angle)` - Add dimension (linear/angular/radial/diameter)
10. `set_layer(layer_name, color, create_if_not_exist)` - Set/create AutoCAD layer
11. `list_layers(filter_type, sort_by, include_details)` - List layers with filtering
12. `scan_elements(element_type, include_geometry, include_properties, layer_filter, bounds)` - Scan drawing elements

**Data Integration Tools** (from `mcp_server_fastmcp.py`):
13. `extract_autocad_parameters(drawing_path, use_current_drawing)` - Extract layout parameters
14. `sync_to_odoo(data, sync_type)` - Sync data to Odoo (parameters/boq/project)
15. `generate_boq(project_id, include_autocad_data)` - Generate Bill of Quantities
16. `export_to_database(drawing_name, include_geometry, incremental_update, sync_to_odoo, element_types, layer_filter)` - Export elements to SQLite
17. `sync_drawing_to_odoo(drawing_name, project_name, project_id, sync_mode, create_project, sync_elements, sync_boq, sync_parameters)` - Full drawing sync to Odoo project
18. `generate_boq_from_drawing(drawing_name, project_id, include_autocad_data, calculation_rules, element_types, layer_filter, output_format, currency)` - Generate BOQ from drawing elements

**NLP Tool** (from `mcp_server_fastmcp.py`):
19. `process_natural_language_command(command, user_context)` - Parse Chinese natural language commands and execute corresponding CAD operations

**Layout Tools** (registered via `util_mcp_sse_manager.py`, executed via GUI Proxy):
20. `get_current_layout()` - Get active layout name and tab order via GUI proxy
21. `switch_to_layout(layout_name)` - Switch to named layout via CTAB system variable
22. `draw_in_layout(layout_name, shape_type, parameters)` - Draw shapes in a specific layout
23. `extract_layout_parameters(layout_name)` - Extract header block attributes and table data from a layout
24. `export_layout_image(layout_name, export_path, image_format)` - Export layout as WMF/BMP/EPS image

### GUI Proxy Handlers (registered in `setup_gui_proxy_handlers`)
- `switch_layout` - Switches layout via `SetVariable("CTAB", layout_name)` with verification
- `get_current_layout` - Gets active layout via `autocad_util.get_active_layout()`
- `extract_parameters` - Extracts all layout values via `autocad_util.get_layouts_values()`
- `extract_layout_parameters` - Extracts single layout header blocks and table data
- `get_autocad_status` - Gets document name, path, saved state, version info
- `draw_line` - Draws line via `autocad_util.draw_line()`
- `draw_circle` - Draws circle via `autocad_util.draw_circle()`
- `export_layout_image` - Exports layout to image file with bounding box calculation

### Python UI Elements (SSE Control Panel)
- Title: "SSE Server Control Panel"
- Status display: "Running" (green) / "Stopped" (red) indicator
- Toggle button: "Start SSE Server" / "Stop SSE Server"
- Test Connection button
- View Status button
- Config info section: port number, transport mode, Gemini CLI config path
- Help section: 4-step usage instructions
- Top bar indicators: SSE status dot (green/red), toggle button, port info label

## 4. Functional Requirements

### Server Lifecycle Management

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-001 | Page SHALL provide a Start button to launch the MCP SSE server on the configured port | Must |
| FR-006-002 | Page SHALL provide a Stop button to shut down the MCP SSE server and close all SSE connections | Must |
| FR-006-003 | Page SHALL provide a Restart button that stops the server, waits briefly, and starts it again | Should |
| FR-006-004 | Start/Stop button SHALL toggle appearance based on server state (show "Start" when stopped, "Stop" when running) | Must |
| FR-006-005 | Server start SHALL verify port availability before attempting to bind | Must |
| FR-006-006 | Server start SHALL register all MCP tools with the tool registry before accepting connections | Must |
| FR-006-007 | Server stop SHALL gracefully close all active SSE connections before shutting down | Must |
| FR-006-008 | Page SHALL disable the Start button while the server is already running and disable Stop while already stopped | Must |

### Tool Registry

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-009 | Page SHALL display the complete list of registered MCP tools with name and description | Must |
| FR-006-010 | Each tool entry SHALL indicate its category (Connection, AutoCAD Drawing, Data Integration, Layout, NLP) | Should |
| FR-006-011 | Each tool entry SHALL show its input parameters schema (parameter names, types, required/optional) | Should |
| FR-006-012 | Tool list SHALL indicate which tools require AutoCAD connection and which require Odoo connection | Should |
| FR-006-013 | Tool registry SHALL support dynamic registration via `MCPToolRegistry.RegisterTool()` | Must |
| FR-006-014 | Tool count SHALL be displayed in the server status panel | Must |

### SSE Transport & Protocol

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-015 | Server SHALL expose `GET /sse` endpoint for SSE client connections with proper headers (Content-Type: text/event-stream, Cache-Control: no-cache) | Must |
| FR-006-016 | Server SHALL expose `POST /messages` endpoint for JSON-RPC 2.0 message handling | Must |
| FR-006-017 | Server SHALL expose `GET /health` endpoint returning server health, active connections, tools count, and initialized state | Must |
| FR-006-018 | Server SHALL expose `GET /` root endpoint returning server info with name, version, and protocol version | Must |
| FR-006-019 | Server SHALL send SSE heartbeat events every 30 seconds to keep connections alive | Should |
| FR-006-020 | Server SHALL send a "connected" SSE event with connection ID upon new client connection | Must |
| FR-006-021 | Server SHALL handle JSON-RPC 2.0 methods: `initialize`, `initialized`, `tools/list`, `tools/call`, `ping` | Must |
| FR-006-022 | Server SHALL return MCP protocol version `2024-11-05` in initialize response | Must |
| FR-006-023 | Server SHALL declare `tools` capability in server capabilities response | Must |

### Status Monitoring

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-024 | Page SHALL display server running state with a visual indicator (green for running, red for stopped) | Must |
| FR-006-025 | Page SHALL display the current server port number | Must |
| FR-006-026 | Page SHALL display the number of active SSE connections | Should |
| FR-006-027 | Page SHALL display the number of registered tools | Should |
| FR-006-028 | Page SHALL display whether the MCP session is initialized (client has completed handshake) | Should |
| FR-006-029 | Status SHALL update in real-time via a callback mechanism when server state changes | Must |
| FR-006-030 | Page SHALL display server uptime when running | Could |

### Connection Testing

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-031 | Page SHALL provide a "Test Connection" button that sends a `test_connection` tool call to the server | Must |
| FR-006-032 | Test result SHALL display success/failure status with timestamp | Must |
| FR-006-033 | Test SHALL verify both the HTTP endpoint (`/health`) and the JSON-RPC protocol layer | Should |
| FR-006-034 | Test button SHALL be disabled when the server is not running | Must |

### GUI Proxy Integration

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-035 | All AutoCAD COM tool operations invoked through MCP SHALL execute via the IGUIProxy message queue | Must |
| FR-006-036 | GUI main thread SHALL poll the proxy request queue every 100ms via DispatcherTimer | Must |
| FR-006-037 | GUI Proxy SHALL support a configurable timeout per request (default: 10 seconds) | Must |
| FR-006-038 | Page SHALL display GUI Proxy status (running/stopped) and pending request count | Should |
| FR-006-039 | Page SHALL display GUI Proxy statistics (total requests, successful, failed, timed out, average execution time) | Could |

### Tool Execution & Logging

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-040 | Page SHALL display a scrollable log viewer showing server activity (connections, tool calls, errors) | Should |
| FR-006-041 | Log entries SHALL include timestamp, log level (INFO/WARNING/ERROR), and message | Should |
| FR-006-042 | Log viewer SHALL auto-scroll to latest entry and support manual scroll-lock | Should |
| FR-006-043 | Page SHOULD provide a "Clear Log" button to reset the log viewer | Could |
| FR-006-044 | Tool execution results SHALL be logged with tool name, execution duration, and success/failure status | Should |

### Configuration

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-006-045 | Page SHALL display the configured server port (default: 8084) | Must |
| FR-006-046 | Page SHALL allow the user to change the server port when the server is stopped | Should |
| FR-006-047 | Page SHALL display the SSE endpoint URL (e.g., `http://localhost:8084/sse`) | Should |
| FR-006-048 | Page SHALL display configuration guidance for AI assistant clients (Gemini CLI, Claude Code) | Could |

## 5. UI Wireframe Description

```
+--------------------------------------------------------------+
|                    AI Assistant (MCP)                         |
+--------------------------------------------------------------+
|                                                               |
|  Server Control                                               |
|  +----------------------------------------------------------+|
|  | Status: [*] Running        Port: 8084                    ||
|  | Uptime: 1h 23m 45s         Connections: 2                ||
|  | Protocol: MCP 2024-11-05   Tools: 24 registered          ||
|  |                                                           ||
|  | [Stop Server]  [Restart]  [Test Connection]               ||
|  +----------------------------------------------------------+|
|                                                               |
|  Registered Tools                    |  Configuration          |
|  +--------------------------------+ |  +--------------------+ |
|  | Category: Connection (4)       | |  | Server             | |
|  |   * test_connection            | |  | Port: [8084]       | |
|  |   * get_server_info            | |  | Host: localhost     | |
|  |   * check_autocad_status  [A]  | |  | Transport: SSE     | |
|  |   * check_odoo_status     [O]  | |  |                    | |
|  |                                | |  | Endpoints          | |
|  | Category: AutoCAD Drawing (8)  | |  | SSE: /sse          | |
|  |   * create_new_drawing    [A]  | |  | RPC: /messages     | |
|  |   * draw_line             [A]  | |  | Health: /health    | |
|  |   * draw_circle           [A]  | |  |                    | |
|  |   * create_text           [A]  | |  | GUI Proxy          | |
|  |   * add_dimension         [A]  | |  | Status: Running    | |
|  |   * set_layer             [A]  | |  | Pending: 0         | |
|  |   * list_layers           [A]  | |  | Avg Time: 45ms     | |
|  |   * scan_elements         [A]  | |  |                    | |
|  |                                | |  | Prerequisites      | |
|  | Category: Data Integration (6) | |  | [A] = AutoCAD req. | |
|  |   * extract_autocad_params[A]  | |  | [O] = Odoo req.    | |
|  |   * sync_to_odoo         [O]   | |  | [A][O] = Both req. | |
|  |   * generate_boq       [A][O]  | |  +--------------------+ |
|  |   * export_to_database   [A]   | |                         |
|  |   * sync_drawing_to_odoo[A][O] | |                         |
|  |   * generate_boq_from.. [A][O] | |                         |
|  |                                | |                         |
|  | Category: Layout (5)           | |                         |
|  |   * get_current_layout    [A]  | |                         |
|  |   * switch_to_layout      [A]  | |                         |
|  |   * draw_in_layout        [A]  | |                         |
|  |   * extract_layout_params [A]  | |                         |
|  |   * export_layout_image   [A]  | |                         |
|  |                                | |                         |
|  | Category: NLP (1)              | |                         |
|  |   * process_natural_lang..[A]  | |                         |
|  +--------------------------------+ |                         |
|                                                               |
|  Activity Log                                                 |
|  +----------------------------------------------------------+|
|  | 14:30:01 [INFO]  MCP SSE Server started on port 8084     ||
|  | 14:30:02 [INFO]  Registered 24 MCP tools                 ||
|  | 14:30:15 [INFO]  SSE connection established: abc123       ||
|  | 14:30:16 [INFO]  Client initializing: gemini-cli v1.2    ||
|  | 14:30:16 [INFO]  MCP session initialized                 ||
|  | 14:30:20 [INFO]  Executing tool: check_autocad_status     ||
|  | 14:30:20 [INFO]  Tool completed: check_autocad_status     ||
|  | 14:30:45 [INFO]  Executing tool: extract_layout_params    ||
|  | 14:30:46 [INFO]  Tool completed (1.2s): extract_layout... ||
|  +----------------------------------------------------------+|
|  [Clear Log]                                     [Auto-Scroll]|
|                                                               |
+--------------------------------------------------------------+
```

### Layout Details
- **Server Control Panel**: Top section with status indicator (Ellipse with green/red fill), port, uptime, connection count, tool count, and action buttons
- **Registered Tools Panel**: Left-center scrollable list grouped by category, each tool shows name, dependency badges [A] for AutoCAD-required and [O] for Odoo-required
- **Configuration Panel**: Right-center panel showing server settings, endpoint URLs, GUI proxy status, and prerequisite legend
- **Activity Log Panel**: Bottom section with scrollable log viewer, timestamp and level prefix, clear and auto-scroll toggle buttons
- **Buttons**: Start/Stop toggles, Restart, Test Connection all in the server control panel

## 6. Data Model

### ViewModel Properties

```csharp
public class MCPViewModel : ObservableObject
{
    // Server State
    public bool IsServerRunning { get; set; }
    public string ServerStatus { get; set; }          // "Running" / "Stopped" / "Starting..." / "Stopping..."
    public int ServerPort { get; set; } = 8084;
    public string ProtocolVersion { get; set; } = "2024-11-05";
    public int ActiveConnectionCount { get; set; }
    public int RegisteredToolCount { get; set; }
    public bool IsSessionInitialized { get; set; }
    public DateTime? ServerStartTime { get; set; }
    public string ServerUptime { get; set; }           // Formatted uptime string
    public string SSEEndpointUrl { get; set; }         // "http://localhost:{port}/sse"
    public string HealthEndpointUrl { get; set; }      // "http://localhost:{port}/health"

    // GUI Proxy State
    public bool IsGUIProxyRunning { get; set; }
    public int GUIProxyPendingRequests { get; set; }
    public GUIProxyStats? ProxyStatistics { get; set; }

    // Tool Registry
    public ObservableCollection<MCPToolViewModel> RegisteredTools { get; set; }

    // Connection Test
    public string TestConnectionResult { get; set; }
    public bool IsTestConnectionSuccess { get; set; }
    public DateTime? LastTestTime { get; set; }

    // Activity Log
    public ObservableCollection<LogEntry> ActivityLog { get; set; }
    public bool IsAutoScrollEnabled { get; set; } = true;

    // Commands
    public IAsyncRelayCommand StartServerCommand { get; }
    public IAsyncRelayCommand StopServerCommand { get; }
    public IAsyncRelayCommand RestartServerCommand { get; }
    public IAsyncRelayCommand TestConnectionCommand { get; }
    public IRelayCommand ClearLogCommand { get; }
    public IRelayCommand ToggleAutoScrollCommand { get; }
}

public class MCPToolViewModel
{
    public string Name { get; set; }
    public string Description { get; set; }
    public string Category { get; set; }               // "Connection", "AutoCAD Drawing", "Data Integration", "Layout", "NLP"
    public bool RequiresAutoCAD { get; set; }
    public bool RequiresOdoo { get; set; }
    public MCPInputSchema InputSchema { get; set; }
}

public class LogEntry
{
    public DateTime Timestamp { get; set; }
    public string Level { get; set; }                  // "INFO", "WARNING", "ERROR"
    public string Message { get; set; }
    public string Source { get; set; }                  // "Server", "Tool", "SSE", "GUIProxy"
}
```

## 7. API/Service Dependencies

| Service | Interface | Methods Used |
|---------|-----------|-------------|
| MCP SSE Server | `MCPSSEServer` | `StartAsync()`, `StopAsync()`, `IsRunning`, `Port`, `ActiveConnections`, `BroadcastEventAsync()` |
| MCP Tool Registry | `MCPToolRegistry` | `GetTools()`, `ExecuteToolAsync(name, arguments)`, `RegisterTool()` |
| GUI Proxy | `IGUIProxy` | `Start()`, `Stop()`, `IsRunning`, `PendingRequestCount`, `RegisterHandler()`, `ExecuteInGuiAsync()`, `ProcessRequests()`, `GetStatistics()` |
| AutoCAD Service | `IAutoCADService` | `IsConnected` (for tool prerequisite display) |
| Odoo Service | `IOdooService` | `IsConnected` (for tool prerequisite display) |
| Navigation Service | `INavigationService` | `NavigateTo()` for linking to AutoCAD/Odoo pages when prerequisites are unmet |

## 8. Validation Rules

| Rule | Description |
|------|-------------|
| VR-006-001 | Server port must be an integer between 1024 and 65535 |
| VR-006-002 | Server port must not be in use by another process before starting |
| VR-006-003 | Server cannot be started if it is already running (idempotent start returns true) |
| VR-006-004 | Server cannot be stopped if it is already stopped (idempotent stop returns true) |
| VR-006-005 | Port configuration field must be disabled while the server is running |
| VR-006-006 | Test Connection button must be disabled when the server is not running |
| VR-006-007 | Restart must complete stop before starting (1-second minimum delay between stop and start) |
| VR-006-008 | All registered tool names must be unique within the MCPToolRegistry |
| VR-006-009 | JSON-RPC requests missing required fields (jsonrpc, method) must return InvalidRequest error (-32600) |
| VR-006-010 | `tools/call` requests with unknown tool names must return MethodNotFound error (-32601) |
| VR-006-011 | `tools/call` requests with missing required parameters must return InvalidParams error (-32602) |
| VR-006-012 | GUI Proxy requests must time out after the configured timeout (default 10 seconds) and return TimeoutError (-32003) |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| Port already in use | "Port {port} is already in use. Please choose a different port or stop the conflicting application." | Highlight port field, suggest alternative port, show "Change Port" option |
| Server start failure | "MCP Server failed to start. Error: {details}" | Display error in status panel, keep server in stopped state, log full error |
| Server crash during operation | "MCP Server stopped unexpectedly. Check logs for details." | Update status to stopped, offer restart button, log crash details |
| SSE connection dropped | "SSE connection lost for client {connectionId}." | Remove from active connections, log disconnection, no user dialog needed |
| Test connection timeout | "Connection test timed out. Server may be unresponsive." | Display timeout in test result area, suggest restart |
| Test connection failure | "Connection test failed: {error}. Server may not be responding to JSON-RPC requests." | Display error, suggest checking server logs |
| Tool execution failure | "Tool '{toolName}' failed: {error}" | Return JSON-RPC error response with code -32603, log error in activity log |
| AutoCAD not connected (tool requires it) | "Tool '{toolName}' requires AutoCAD connection. Please connect to AutoCAD first." | Return error in tool result, display in log, show link to AutoCAD page |
| Odoo not connected (tool requires it) | "Tool '{toolName}' requires Odoo connection. Please connect to Odoo first." | Return error in tool result, display in log, show link to Odoo settings |
| GUI Proxy timeout | "AutoCAD operation timed out ({action}). AutoCAD may be busy." | Return timeout error to MCP client, log in activity log, increment timeout counter in proxy stats |
| GUI Proxy handler not found | "Unknown GUI proxy action: {action}." | Return error to caller, log warning |
| JSON parse error | "Invalid JSON in request body." | Return JSON-RPC ParseError (-32700), log malformed request |
| Invalid JSON-RPC method | "Method not found: {method}." | Return JSON-RPC MethodNotFound (-32601) |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/MCPAssistantPage.xaml` - WPF page with data bindings for server control, tool list, config, and log
- `ViewModels/MCPViewModel.cs` - MVVM ViewModel with commands and state management
- `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` (exists, 376 lines) - SSE server with WebApplication endpoints
- `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` (exists, 413 lines) - Tool registry with 7 core tools
- `OdooAutoCAD.MCP/Protocol/MCPModels.cs` (exists, 337 lines) - JSON-RPC 2.0 and MCP protocol models
- `OdooAutoCAD.Core/Threading/IGUIProxy.cs` (exists, 195 lines) - GUI Proxy interface

### JSON-RPC 2.0 Protocol Implementation
The server processes JSON-RPC 2.0 requests at `POST /messages`. The existing `MCPModels.cs` defines `JsonRpcRequest` and `JsonRpcResponse` with proper `[JsonPropertyName]` attributes. Standard error codes are defined:
- `-32700` ParseError
- `-32600` InvalidRequest
- `-32601` MethodNotFound
- `-32602` InvalidParams
- `-32603` InternalError
- `-32001` AutoCADError (custom)
- `-32002` OdooError (custom)
- `-32003` TimeoutError (custom)
- `-32004` NotConnectedError (custom)

The `ProcessJsonRpcRequestAsync` method routes by method name: `initialize` (handshake), `initialized` (session ready), `tools/list` (return all tool definitions), `tools/call` (execute a tool), `ping` (keep-alive).

### SSE Transport
The SSE connection at `GET /sse` sets required headers (`text/event-stream`, `no-cache`, `keep-alive`, `X-Accel-Buffering: no`) and sends events using the `SSEEvent` model (event name, data, optional id, optional retry). The `BroadcastEventAsync` method sends events to all connected clients. Heartbeats are sent every 30 seconds. Each connection is tracked in a `ConcurrentDictionary<string, SSEConnection>`.

### MCP Protocol Version
The server reports protocol version `2024-11-05` in the `MCPServerInfo.ProtocolVersion` field, returned during the `initialize` handshake. This must match what AI assistant clients expect.

### DispatcherTimer for GUI Proxy Polling
The WPF page must set up a `DispatcherTimer` with a 100ms interval to call `IGUIProxy.ProcessRequests()`, matching the Python pattern where the Tkinter main loop calls `self.after(100, process_gui_proxy_requests)`. This ensures all COM operations queued by MCP tool handlers are executed on the STA/GUI thread:

```csharp
// In MCPAssistantPage.xaml.cs or MCPViewModel
private DispatcherTimer _proxyTimer;

private void StartProxyPolling()
{
    _proxyTimer = new DispatcherTimer
    {
        Interval = TimeSpan.FromMilliseconds(100)
    };
    _proxyTimer.Tick += (s, e) =>
    {
        var processed = _guiProxy.ProcessRequests();
        if (processed > 0)
        {
            // Update pending request count in ViewModel
            GUIProxyPendingRequests = _guiProxy.PendingRequestCount;
        }
    };
    _proxyTimer.Start();
}
```

### Thread Safety Architecture
The MCP SSE server runs on a background thread (MTA). AutoCAD COM requires STA thread. The `IGUIProxy` bridges this:
1. MCP tool handler calls `_guiProxy.ExecuteInGuiAsync("action", parameters)` from the MCP thread
2. Request is placed in a `ConcurrentQueue` with a `TaskCompletionSource` for the response
3. DispatcherTimer on the GUI thread calls `ProcessRequests()` every 100ms
4. Handler executes the COM operation on the GUI/STA thread
5. Response is set on the `TaskCompletionSource`, completing the awaited task on the MCP thread

### Tool Registration Pattern
Tools are registered in `MCPToolRegistry.RegisterAllTools()` using the private `RegisterTool(name, description, inputSchema, executor)` method. Each tool gets an `MCPTool` definition (name, description, inputSchema) and an `MCPToolExecutor` delegate. Tools that require AutoCAD COM go through `_guiProxy.ExecuteInGuiAsync()`. Tools that only need Odoo REST API can call `_odooService` directly since Odoo operations are HTTP-based and thread-safe.

### MVVM Bindings
- Server status indicator: `{Binding IsServerRunning, Converter={StaticResource BoolToColorConverter}}`
- Start/Stop button content: `{Binding IsServerRunning, Converter={StaticResource BoolToStartStopConverter}}`
- Start command: `{Binding StartServerCommand}` with `CanExecute` bound to `!IsServerRunning`
- Stop command: `{Binding StopServerCommand}` with `CanExecute` bound to `IsServerRunning`
- Test Connection: `{Binding TestConnectionCommand}` with `CanExecute` bound to `IsServerRunning`
- Tool list: `ItemsSource="{Binding RegisteredTools}"` with `GroupStyle` for category grouping
- Activity log: `ItemsSource="{Binding ActivityLog}"` in a `ListBox` or `DataGrid` with `ScrollIntoView` for auto-scroll

### Key Differences from Python
- Python uses emoji indicators in labels ("Running" / "Stopped") -- C# uses Ellipse with Fill binding and text labels
- Python `MCPSSEManager` wraps `StandardMCPSSEServer` -- C# `MCPSSEServer` directly manages WebApplication
- Python registers tools across two files (`mcp_server_fastmcp.py` decorators + `util_mcp_sse_manager.py` manual registration) -- C# centralizes in `MCPToolRegistry`
- Python GUI proxy uses a global singleton (`get_gui_proxy()`) -- C# uses DI-injected `IGUIProxy`
- Python uses `threading.Thread` for server -- C# uses `Task.Run` with `WebApplication.RunAsync`
- Python Tkinter `self.after(100, callback)` -- C# `DispatcherTimer` with 100ms interval
- Python `queue.Queue` for proxy requests -- C# `ConcurrentQueue<T>` with `TaskCompletionSource`
- Python SSE control panel is embedded in main form -- C# has a dedicated navigation page
