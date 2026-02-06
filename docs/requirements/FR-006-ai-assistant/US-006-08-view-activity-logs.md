# US-006-08: View Activity Logs

## User Story
**As a** System Admin,
**I want to** view server activity logs,
**So that** I can troubleshoot issues with AI assistant operations.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: The page displays a scrollable log viewer showing server activity (connections, tool calls, errors)
- [ ] AC-02: Each log entry includes a timestamp, log level (INFO/WARNING/ERROR), and message
- [ ] AC-03: The log viewer auto-scrolls to the latest entry by default
- [ ] AC-04: The user can toggle auto-scroll on/off to manually browse the log
- [ ] AC-05: A "Clear Log" button is available to reset the log viewer
- [ ] AC-06: Log entries include a source field indicating the origin (Server, Tool, SSE, GUIProxy)
- [ ] AC-07: Server start, stop, connection events, tool executions, and errors are all logged

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-040 | Page SHALL display a scrollable log viewer showing server activity | Should |
| FR-006-041 | Log entries SHALL include timestamp, log level, and message | Should |
| FR-006-042 | Log viewer SHALL auto-scroll to latest entry and support manual scroll-lock | Should |
| FR-006-043 | Page SHOULD provide a "Clear Log" button to reset the log viewer | Could |
| FR-006-044 | Tool execution results SHALL be logged with tool name, execution duration, and success/failure status | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-08-01 | Create Activity Log panel XAML with scrollable ListBox/DataGrid, Clear Log button, and Auto-Scroll toggle | `Views/Pages/MCPAssistantPage.xaml` | M |
| TASK-006-08-02 | Implement LogEntry model with Timestamp, Level, Message, and Source properties | `OdooAutoCAD.MCP/Protocol/MCPModels.cs` | S |
| TASK-006-08-03 | Add ActivityLog ObservableCollection and log management methods (AddLog, ClearLog) to ViewModel | `ViewModels/MCPViewModel.cs` | M |
| TASK-006-08-04 | Implement auto-scroll behavior with IsAutoScrollEnabled toggle and ScrollIntoView logic | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-08-05 | Implement ClearLogCommand and ToggleAutoScrollCommand relay commands | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-08-06 | Wire MCPSSEServer events (connection, tool execution, errors) to the ActivityLog collection | `ViewModels/MCPViewModel.cs` | M |

## Dependencies
- Depends on: None (log viewer is always visible; it shows entries even before the server starts)
- Blocks: None

## Notes
- The Activity Log panel is positioned at the bottom of the MCP Assistant page.
- Log entries are color-coded by level: INFO (default text), WARNING (orange/amber), ERROR (red).
- The auto-scroll behavior uses ScrollIntoView on the last item in the ListBox. When the user scrolls manually (scroll-lock), auto-scroll is temporarily disabled and can be re-enabled with the Auto-Scroll toggle button.
- Log entry format: "HH:mm:ss [LEVEL] Message" matching the wireframe format (e.g., "14:30:01 [INFO] MCP SSE Server started on port 8084").
- Tool execution logs include duration (e.g., "Tool completed (1.2s): extract_layout_params") per FR-006-044.
- The ObservableCollection should have a maximum capacity (e.g., 1000 entries) to prevent unbounded memory growth; oldest entries are removed when the limit is exceeded.
