# US-002-02: View Layouts

## User Story
**As a** CAD Engineer,
**I want to** see all layouts in my drawing,
**So that** I can select which layout to work with.

## Parent Feature
- **FR**: [FR-002-autocad-connection](FR-002-autocad-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: When connected to AutoCAD, the page displays a list of all layouts in the current drawing, excluding "Model"
- [ ] AC-02: Each layout entry shows the layout name and tab order
- [ ] AC-03: The currently active layout is visually highlighted in the list
- [ ] AC-04: The user can select a layout from the list to view its details in the right panel
- [ ] AC-05: The user can switch the active layout in AutoCAD by selecting a different layout
- [ ] AC-06: The layout list is refreshed when the connection is first established and can be manually refreshed
- [ ] AC-07: All layout-related COM operations execute on the GUI/STA thread via IGUIProxy

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-002-008 | Page SHALL display a list of all layouts in the current drawing (excluding "Model") | Must |
| FR-002-009 | Each layout entry SHALL show layout name and tab order | Must |
| FR-002-010 | User SHALL be able to select a layout to view its details | Must |
| FR-002-011 | Page SHALL show the currently active layout highlighted | Should |
| FR-002-012 | User SHALL be able to switch to a different layout | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-002-02-01 | Add Layouts ListView panel to the AutoCAD page XAML (left panel) | `Views/Pages/AutoCADPage.xaml` | M |
| TASK-002-02-02 | Implement Layouts ObservableCollection and SelectedLayout property in ViewModel | `ViewModels/AutoCADViewModel.cs` | S |
| TASK-002-02-03 | Implement GetLayouts() in AutoCADService to retrieve layout list excluding "Model" | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-02-04 | Implement GetActiveLayout() and SetActiveLayout(name) in AutoCADService | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-02-05 | Define GetLayouts(), GetActiveLayout() in IAutoCADService interface | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | S |
| TASK-002-02-06 | Implement SelectLayoutCommand to load layout details on selection change | `ViewModels/AutoCADViewModel.cs` | S |
| TASK-002-02-07 | Add visual highlighting for the active layout in the ListView (DataTemplate/Style) | `Views/Pages/AutoCADPage.xaml` | S |

## Dependencies
- Depends on: US-002-01 (AutoCAD connection must be established first)
- Blocks: US-002-03 (parameter extraction requires a selected layout)

## Notes
- The Python reference function `get_doc_layouts()` returns all layouts excluding "Model". The C# implementation should mirror this filtering.
- The Python function `get_active_layout()` returns the currently active layout, which is used to highlight the active item in the list.
- Layout names follow patterns like "S405-201", "S405-202", etc., as shown in the wireframe.
- Validation rule VR-002-002 applies: all layout controls must be disabled when AutoCAD is disconnected.
