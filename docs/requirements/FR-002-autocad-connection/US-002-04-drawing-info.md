# US-002-04: Drawing Info

## User Story
**As a** CAD Engineer,
**I want to** see the current drawing file name and path,
**So that** I know which drawing is active.

## Parent Feature
- **FR**: [FR-002-autocad-connection](FR-002-autocad-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Upon successful connection to AutoCAD, the page displays the current document filename (e.g., "Frame-001.dwg")
- [ ] AC-02: The full file path of the current document is displayed (e.g., "C:\Projects\Steel\Frame-001.dwg")
- [ ] AC-03: Document info is displayed in the Connection Status panel at the top of the page
- [ ] AC-04: Document info updates when the active document changes (on manual refresh)
- [ ] AC-05: When no document is open, a message is displayed: "No drawing is open in AutoCAD. Please open a drawing."
- [ ] AC-06: When no document is open, layout and table controls are disabled

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-002-005 | Page SHALL display current document filename and full path | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-002-04-01 | Add DocumentName and DocumentPath display fields to the Connection Status panel in XAML | `Views/Pages/AutoCADPage.xaml` | S |
| TASK-002-04-02 | Implement DocumentName and DocumentPath properties in ViewModel with change notification | `ViewModels/AutoCADViewModel.cs` | S |
| TASK-002-04-03 | Populate document info from ActiveDocument during connection in AutoCADService | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | S |
| TASK-002-04-04 | Implement RefreshStatusCommand to update document info on demand | `ViewModels/AutoCADViewModel.cs` | S |
| TASK-002-04-05 | Add "no document open" state handling with appropriate message and control disabling | `ViewModels/AutoCADViewModel.cs` | S |

## Dependencies
- Depends on: US-002-01 (AutoCAD connection must be established to read document info)
- Blocks: None

## Notes
- Document filename and path are extracted from AutoCAD's `ActiveDocument.Name` and `ActiveDocument.FullName` properties during the connection flow.
- This information is displayed in the Connection Status panel alongside the connection indicator, as shown in the wireframe.
- The Python connection flow extracts filename and path at step 6 after successfully obtaining ActiveDocument. The C# implementation should do the same within the ConnectAsync flow.
- If `ActiveDocument` is null (no drawing open), the application should gracefully degrade by showing a user-friendly message and disabling dependent controls.
