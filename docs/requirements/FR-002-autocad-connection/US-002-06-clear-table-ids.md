# US-002-06: Clear Table IDs

## User Story
**As a** CAD Engineer,
**I want to** clear table IDs from layouts,
**So that** I can re-extract data when needed.

## Parent Feature
- **FR**: [FR-002-autocad-connection](FR-002-autocad-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: The user can update block attributes in AutoCAD from the UI
- [ ] AC-02: A "Clear IDs (This Layout)" button clears the table Detail ID column for the currently selected layout
- [ ] AC-03: A "Clear All" button clears table Detail ID columns for ALL layouts in the drawing
- [ ] AC-04: Both clear operations display a confirmation dialog before proceeding (VR-002-006)
- [ ] AC-05: Upon completion, a success or failure message is displayed to the user
- [ ] AC-06: After clearing, the table data display is refreshed to reflect the cleared IDs
- [ ] AC-07: If an attribute write fails, an error message is displayed: "Failed to update attribute '{tag}' in AutoCAD."
- [ ] AC-08: Clear buttons are only enabled when AutoCAD is connected
- [ ] AC-09: All write-back COM operations execute on the GUI/STA thread via IGUIProxy

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-002-021 | User SHALL be able to update block attributes in AutoCAD from the UI | Must |
| FR-002-022 | User SHALL be able to clear table IDs for the current layout | Must |
| FR-002-023 | User SHALL be able to clear table IDs for ALL layouts | Must |
| FR-002-024 | Write-back operations SHALL confirm success/failure to the user | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-002-06-01 | Add "Clear IDs (This Layout)" and "Clear All" buttons to the Actions Bar in XAML | `Views/Pages/AutoCADPage.xaml` | S |
| TASK-002-06-02 | Implement ClearTableIdsCommand (single layout) in ViewModel with confirmation dialog | `ViewModels/AutoCADViewModel.cs` | M |
| TASK-002-06-03 | Implement ClearAllTableIdsCommand (all layouts) in ViewModel with confirmation dialog | `ViewModels/AutoCADViewModel.cs` | M |
| TASK-002-06-04 | Implement clear_table_id equivalent in AutoCADService for single layout | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-06-05 | Implement clear_all_tables_id equivalent in AutoCADService for all layouts | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-06-06 | Implement set_attribute_value equivalent for writing block attribute values back to AutoCAD | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-06-07 | Define SetTableValue() and ClearSelection() in IAutoCADService interface | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | S |
| TASK-002-06-08 | Add success/failure notification display after write-back operations | `ViewModels/AutoCADViewModel.cs` | S |

## Dependencies
- Depends on: US-002-01 (AutoCAD connection must be established for write-back operations)
- Blocks: None

## Notes
- The Python functions `clear_table_id(layout)` and `clear_all_tables_id()` clear the Detail ID column (index 8) in layout tables. The C# implementation should provide equivalent functionality.
- The Python function `set_attribute_value(block, tag, val)` is used to write values back to AutoCAD block attributes. The C# equivalent must also go through IGUIProxy for thread safety.
- Validation rule VR-002-006: Clear operations require a confirmation dialog before execution to prevent accidental data loss.
- Clearing table IDs is necessary when a user wants to re-extract and re-sync data to Odoo, as the Detail ID is used to track which rows have already been synced.
- The Actions Bar at the bottom of the page contains: [Extract Parameters] [Clear IDs (This Layout)] [Clear All], as shown in the wireframe.
