# US-008-03: Persistent Connection Status

## User Story
**As a** CAD Engineer,
**I want to** see connection status for AutoCAD, Odoo, and MCP at all times regardless of which page I am on,
**So that** I am always aware of system connectivity.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: The sidebar displays a "Connection Status" section at the bottom with status indicators for AutoCAD, Odoo, and MCP Server.
- [ ] AC-02: Each connection indicator displays a colored `Ellipse` (green = connected/running, gray = disconnected/stopped) and descriptive status text.
- [ ] AC-03: AutoCAD status shows "Connected" (green) or "Disconnected" (gray) based on `IAutoCADService` connection state.
- [ ] AC-04: Odoo status shows "Connected" (green) or "Disconnected" (gray) based on `IOdooService` connection state.
- [ ] AC-05: MCP status shows "Port XXXX" (green) when the server is running, or "Stopped" (gray) when it is not.
- [ ] AC-06: Connection indicators update in real-time via data binding to `MainViewModel` observable properties (`AutoCADStatusColor`, `OdooStatusColor`, `MCPStatusColor`).
- [ ] AC-07: Connection status colors only transition between defined states: Gray (disconnected/stopped), Green (connected/running), and Red (error).
- [ ] AC-08: The connection status section is visible on all pages without requiring user interaction.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-022 | AutoCAD status with colored Ellipse indicator and status text | Must |
| FR-008-023 | Odoo status with colored Ellipse indicator and status text | Must |
| FR-008-024 | MCP Server status with colored Ellipse indicator and status text | Must |
| FR-008-025 | Real-time status updates via data binding to MainViewModel observable properties | Must |
| FR-008-026 | MCP status displays port number when server is running | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-03-01 | Add AutoCADStatusText, AutoCADStatusColor, OdooStatusText, OdooStatusColor, MCPStatusText, MCPStatusColor observable properties to MainViewModel | `ViewModels/MainViewModel.cs` | M |
| TASK-008-03-02 | Define connection status XAML section in sidebar with Ellipse indicators and bound TextBlocks | `Views/MainWindow.xaml` | M |
| TASK-008-03-03 | Implement status update methods in MainViewModel that poll or subscribe to service state changes | `ViewModels/MainViewModel.cs` | M |
| TASK-008-03-04 | Define connected/disconnected/error color resources (SuccessBrush, default Gray, ErrorBrush) | `Themes/Colors.xaml` | S |
| TASK-008-03-05 | Wire MCP server status to display port number when running | `ViewModels/MainViewModel.cs` | S |

## Dependencies
- Depends on: US-008-01 (sidebar layout must exist to host the connection status section)
- Blocks: None

## Notes
- The Python implementation displays connection status in the top bar area with emoji indicators (green/red circles). The C# port moves these to the sidebar bottom section using colored `Ellipse` elements for a cleaner integration.
- Validation rule VR-008-006 restricts color transitions to Gray, Green, and Red only. No other colors should be used for connection status indicators.
- Connection status properties use `Brush` type (e.g., `Brushes.Gray`, `Brushes.Green`) to bind directly to `Ellipse.Fill`.
- The `ConnectionStatusViewModel` may be used as an embedded ViewModel within the sidebar, or the properties can live directly on `MainViewModel` for simplicity.
