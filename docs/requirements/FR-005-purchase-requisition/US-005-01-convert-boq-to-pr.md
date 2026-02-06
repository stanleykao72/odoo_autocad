# US-005-01: Convert BOQ to PR

## User Story
**As a** CAD Engineer,
**I want to** convert BOQ entries to a Purchase Requisition,
**So that** I can initiate the procurement process for the materials in my drawing.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: A "Convert BOQ to PR" button is visible on the Purchase Requisition page
- [ ] AC-02: Clicking the Convert button calls `IOdooService.ConvertBOQToPRAsync(projectId, boqEntryIds)` with the current project context
- [ ] AC-03: When no specific BOQ entry IDs are provided, all BOQ entries for the project are included in the conversion
- [ ] AC-04: User can select specific BOQ entries to convert (partial conversion)
- [ ] AC-05: A progress indicator is displayed while the Odoo backend processes the conversion request
- [ ] AC-06: On successful conversion, the PR list refreshes and the newly created PR is highlighted
- [ ] AC-07: On failure, the error message returned by the Odoo API is displayed to the user
- [ ] AC-08: The Convert button is disabled while a conversion is already in progress (IsConverting == true)
- [ ] AC-09: The Convert button is disabled when no BOQ data exists for the project, with a warning message shown

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-001 | Page SHALL provide a "Convert BOQ to PR" button that triggers the BOQ-to-PR conversion | Must |
| FR-005-002 | Conversion SHALL call `IOdooService.ConvertBOQToPRAsync(projectId, boqEntryIds)` with the current project context | Must |
| FR-005-003 | If no specific BOQ entry IDs are provided, ALL BOQ entries for the project SHALL be included in the conversion | Must |
| FR-005-004 | User SHALL be able to select specific BOQ entries to convert (partial conversion) | Should |
| FR-005-005 | Conversion SHALL display a progress indicator while the Odoo backend processes the request | Must |
| FR-005-006 | On successful conversion, the page SHALL refresh the PR list and show the newly created PR highlighted | Must |
| FR-005-007 | On failure, the page SHALL display the error message returned by the Odoo API | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-01-01 | Add "Convert BOQ to PR" button to the page layout with proper styling | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-01-02 | Implement `ConvertBOQToPRCommand` async relay command in ViewModel | `ViewModels/PurchaseRequisitionViewModel.cs` | M |
| TASK-005-01-03 | Implement `ConvertBOQToPRAsync(projectId, boqEntryIds)` method in Odoo service | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | M |
| TASK-005-01-04 | Add optional header_id extraction path via `IAutoCADService.GetLayoutsHeaderIds()` | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | M |
| TASK-005-01-05 | Wrap AutoCAD COM calls through `IGUIProxy.ExecuteInGuiAsync()` for thread safety | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | S |
| TASK-005-01-06 | Add progress indicator binding and conversion state management (IsConverting, StatusMessage) | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-01-07 | Implement post-conversion PR list refresh and newly created PR highlighting | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-01-08 | Implement error handling for conversion failures with user-friendly messages | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-01-09 | Add BOQ entry selection UI for partial conversion support | `Views/Pages/PurchaseRequisitionPage.xaml` | M |

## Dependencies
- Depends on: US-004-04 (BOQ pushed to Odoo -- BOQ entries must exist before conversion)
- Blocks: US-005-02 (PR list needs at least one PR to display meaningfully)

## Notes
- The core conversion logic resides on the Odoo backend (`boq2pr_v2` endpoint). The C# client acts as a thin orchestrator.
- Two conversion paths exist: project-based (simpler, preferred) and header_id-based (matches Python pipeline exactly). The project-based path is recommended for the initial implementation.
- When using the header_id-based path, AutoCAD must be connected and `IGUIProxy` must be used for thread-safe COM calls.
- Duplicate conversion prevention is handled by disabling the button when `IsConverting` is true (VR-005-007).
