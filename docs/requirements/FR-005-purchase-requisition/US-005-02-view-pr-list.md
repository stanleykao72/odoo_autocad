# US-005-02: View PR List

## User Story
**As a** CAD Engineer,
**I want to** see a list of all PRs for the current project,
**So that** I can track which requisitions have been created.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: The page displays a list of all Purchase Requisitions for the current project
- [ ] AC-02: Each PR list entry shows: Reference, State (as a color-coded badge), Line Count, and Created Date
- [ ] AC-03: The PR list loads automatically when the page is opened and a valid project context exists
- [ ] AC-04: A manual "Refresh" button is available to reload the PR list
- [ ] AC-05: When no PRs exist for the project, an empty state message is displayed: "No Purchase Requisitions found for this project" with guidance on how to create one
- [ ] AC-06: A loading indicator is shown while the PR list is being fetched from Odoo
- [ ] AC-07: If the PR list fetch fails, an error message is displayed with a retry option

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-008 | Page SHALL display a list of all Purchase Requisitions for the current project | Must |
| FR-005-009 | Each PR list entry SHALL show: Reference, State, Line Count, and Created Date | Must |
| FR-005-010 | PR list SHALL load automatically when the page is opened and a project context exists | Must |
| FR-005-011 | PR list SHALL support pull-to-refresh or a manual refresh button | Must |
| FR-005-012 | Empty state SHALL display "No Purchase Requisitions found for this project" with guidance | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-02-01 | Create PR list `ListView` layout with columns for Reference, State badge, Line Count, and Date | `Views/Pages/PurchaseRequisitionPage.xaml` | M |
| TASK-005-02-02 | Implement `ObservableCollection<PRDisplayItem> PurchaseRequisitions` property and data binding | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-02-03 | Implement `GetPurchaseRequisitionsAsync(projectId)` in Odoo service interface and implementation | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | M |
| TASK-005-02-04 | Implement auto-load on page open via `OnNavigatedTo` or `Loaded` event triggering `RefreshListCommand` | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-02-05 | Add "Refresh" button to the conversion controls area bound to `RefreshListCommand` | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-02-06 | Implement empty state UI with guidance message and conditional visibility | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-02-07 | Add `IsLoading` state management with loading indicator binding | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-02-08 | Implement error handling for PR list fetch failures with retry button | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-02-09 | Map `PREntry` model to `PRDisplayItem` with computed display properties (StateDisplay, CreatedAtDisplay) | `ViewModels/PurchaseRequisitionViewModel.cs` | S |

## Dependencies
- Depends on: US-003-01 (Odoo connected -- Odoo connection must be active to fetch PR data)
- Blocks: US-005-03 (PR detail view requires selecting a PR from the list)

## Notes
- The PR list is the left panel in the master-detail layout. Selecting a PR from this list populates the detail panel (US-005-03).
- State badges use color-coded backgrounds: Draft=#E0E0E0, Submitted=#BBDEFB, Approved=#C8E6C9, Rejected=#FFCDD2.
- PR Reference must be non-empty; show "N/A" if missing (VR-005-008).
- The list should bind to `FilteredPRs` (the filtered/sorted computed collection) rather than the raw `PurchaseRequisitions` collection.
- Validation: Odoo connection must be active (VR-005-001) and a valid ProjectId must exist (VR-005-002) before loading.
