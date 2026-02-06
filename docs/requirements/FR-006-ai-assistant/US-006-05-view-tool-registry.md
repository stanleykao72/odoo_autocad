# US-006-05: View Tool Registry

## User Story
**As a** System Admin,
**I want to** view the list of registered MCP tools,
**So that** I can confirm which operations are available to AI assistants.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: The page displays a complete list of all registered MCP tools with name and description
- [ ] AC-02: Tools are grouped by category (Connection, AutoCAD Drawing, Data Integration, Layout, NLP)
- [ ] AC-03: Each tool entry shows its input parameters schema including parameter names, types, and required/optional status
- [ ] AC-04: Each tool entry indicates prerequisite connections with badges: [A] for AutoCAD-required, [O] for Odoo-required
- [ ] AC-05: The tool registry supports dynamic registration via MCPToolRegistry.RegisterTool()
- [ ] AC-06: The total tool count is displayed in the server status panel

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-009 | Page SHALL display the complete list of registered MCP tools with name and description | Must |
| FR-006-010 | Each tool entry SHALL indicate its category | Should |
| FR-006-011 | Each tool entry SHALL show its input parameters schema | Should |
| FR-006-012 | Tool list SHALL indicate which tools require AutoCAD connection and which require Odoo connection | Should |
| FR-006-013 | Tool registry SHALL support dynamic registration via MCPToolRegistry.RegisterTool() | Must |
| FR-006-014 | Tool count SHALL be displayed in the server status panel | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-05-01 | Create Registered Tools panel XAML with grouped ListView/ItemsControl and category headers | `Views/Pages/MCPAssistantPage.xaml` | M |
| TASK-006-05-02 | Implement MCPToolViewModel with Name, Description, Category, RequiresAutoCAD, RequiresOdoo, InputSchema properties | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-05-03 | Populate RegisteredTools ObservableCollection from MCPToolRegistry.GetTools() | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-05-04 | Add GroupStyle to ItemsControl for category-based grouping with CollectionViewSource | `Views/Pages/MCPAssistantPage.xaml` | M |
| TASK-006-05-05 | Display [A] and [O] prerequisite badges next to tool names using DataTemplate | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-05-06 | Ensure MCPToolRegistry.RegisterTool() triggers collection update via event notification | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` | S |

## Dependencies
- Depends on: None (tool registry is populated at application startup; list is viewable even when server is stopped)
- Blocks: None

## Notes
- The tool list panel is positioned in the left-center area of the MCP Assistant page.
- Categories and their expected tool counts: Connection (4), AutoCAD Drawing (8), Data Integration (6), Layout (5), NLP (1) -- totaling 24 tools.
- The CollectionViewSource groups by Category property and displays a category header with tool count (e.g., "Category: Connection (4)").
- All registered tool names must be unique within the MCPToolRegistry (VR-006-008).
- The InputSchema display can be collapsed by default and expanded on click to avoid visual clutter.
