# US-006-09: View Tool Prerequisites

## User Story
**As a** CAD Engineer,
**I want to** see which tools require AutoCAD or Odoo connections,
**So that** I know which prerequisites must be met before using AI assistant features.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: Each tool in the registered tools list displays prerequisite badges: [A] for AutoCAD-required, [O] for Odoo-required, or both [A][O]
- [ ] AC-02: A prerequisite legend is displayed in the Configuration panel explaining the badge meanings
- [ ] AC-03: Tools that require AutoCAD but AutoCAD is not connected show a visual warning indicator
- [ ] AC-04: Tools that require Odoo but Odoo is not connected show a visual warning indicator
- [ ] AC-05: When a tool with unmet prerequisites is invoked, the error message includes guidance to connect to the required service (e.g., "Tool 'draw_line' requires AutoCAD connection. Please connect to AutoCAD first.")

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-012 | Tool list SHALL indicate which tools require AutoCAD connection and which require Odoo connection | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-09-01 | Add RequiresAutoCAD and RequiresOdoo boolean properties to MCPToolViewModel | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-09-02 | Display [A] and [O] badges in the tool list DataTemplate based on prerequisite properties | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-09-03 | Add prerequisite legend section to the Configuration panel XAML | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-09-04 | Implement prerequisite checking in MCPToolRegistry.ExecuteToolAsync() with descriptive error responses | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` | M |
| TASK-006-09-05 | Query IAutoCADService.IsConnected and IOdooService.IsConnected to determine unmet prerequisites at display time | `ViewModels/MCPViewModel.cs` | S |

## Dependencies
- Depends on: US-006-05 (view tool registry - prerequisites are displayed within the tool list)
- Blocks: None

## Notes
- Tool prerequisite mapping based on the MCP tool registry:
  - **AutoCAD-required [A]**: check_autocad_status, create_new_drawing, draw_line, draw_circle, create_text, add_dimension, set_layer, list_layers, scan_elements, extract_autocad_parameters, export_to_database, get_current_layout, switch_to_layout, draw_in_layout, extract_layout_parameters, export_layout_image, process_natural_language_command
  - **Odoo-required [O]**: check_odoo_status, sync_to_odoo
  - **Both [A][O]**: generate_boq, sync_drawing_to_odoo, generate_boq_from_drawing
  - **Neither**: test_connection, get_server_info
- When a tool is invoked via MCP and the required service is not connected, the server returns a JSON-RPC error with code -32004 (NotConnectedError) and a message identifying the missing service.
- The Navigation Service can be used to provide links from error states to the AutoCAD or Odoo settings pages.
