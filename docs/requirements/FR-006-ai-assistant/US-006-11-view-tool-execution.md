# US-006-11: View Tool Execution

## User Story
**As a** CAD Engineer,
**I want to** see real-time tool execution results,
**So that** I can monitor what the AI assistant is doing in AutoCAD.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: When an MCP tool is invoked by an AI assistant, the tool name appears in the activity log with an "Executing" status
- [ ] AC-02: When the tool completes, the log entry updates to show the execution duration and success/failure status
- [ ] AC-03: Tool execution results include the tool name, execution time in seconds, and outcome (success or error description)
- [ ] AC-04: Failed tool executions are logged at ERROR level with the failure reason
- [ ] AC-05: GUI Proxy operations (AutoCAD COM calls) show the queued action name and completion status
- [ ] AC-06: The activity log updates in real-time as tools execute, without requiring manual refresh

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-044 | Tool execution results SHALL be logged with tool name, execution duration, and success/failure status | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-11-01 | Add execution logging hooks in MCPToolRegistry.ExecuteToolAsync() to log start, completion, and failure events | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` | M |
| TASK-006-11-02 | Implement tool execution event handler in MCPViewModel to append LogEntry items to ActivityLog | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-11-03 | Add Stopwatch-based duration measurement in ExecuteToolAsync for each tool invocation | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` | S |
| TASK-006-11-04 | Log GUI Proxy action queuing and completion events with action name and duration | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | S |
| TASK-006-11-05 | Ensure log entries are dispatched to the UI thread for ObservableCollection update | `ViewModels/MCPViewModel.cs` | S |

## Dependencies
- Depends on: US-006-01 (start MCP server - tools can only execute when the server is running)
- Depends on: US-006-08 (view activity logs - execution results are displayed in the activity log)
- Blocks: None

## Notes
- Tool execution log format follows the wireframe pattern:
  - Start: "14:30:20 [INFO] Executing tool: check_autocad_status"
  - Success: "14:30:20 [INFO] Tool completed: check_autocad_status" or "14:30:46 [INFO] Tool completed (1.2s): extract_layout_params"
  - Failure: "14:30:20 [ERROR] Tool failed: draw_line - AutoCAD not connected"
- Duration is measured using System.Diagnostics.Stopwatch for precision.
- For tools that execute via the GUI Proxy (all AutoCAD COM operations), the total duration includes queue wait time + execution time. Both components can be logged separately for debugging.
- Log entries are created on the MCP server thread and must be marshalled to the UI thread via Dispatcher.InvokeAsync before being added to the ObservableCollection.
- The Source field of LogEntry should be "Tool" for tool execution events and "GUIProxy" for proxy-specific events.
