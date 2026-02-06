# US-005-08: View PR Totals

## User Story
**As a** Project Manager,
**I want to** view PR totals and line counts at a glance,
**So that** I can assess the scope of each requisition without opening it.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: A summary bar at the bottom of the page displays the count of PRs by state (e.g., "1 Draft | 2 Submitted | 3 Approved | 0 Rejected")
- [ ] AC-02: The summary bar updates automatically when PRs are added, removed, or change state
- [ ] AC-03: Each PR in the list displays its line count alongside the reference and state
- [ ] AC-04: The PR detail view shows a computed total amount (sum of Quantity x UnitPrice for all lines)
- [ ] AC-05: The total PR count is displayed in the project context panel
- [ ] AC-06: Summary counts are recalculated when the PR list is refreshed

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-025 | Page SHALL display a summary bar showing count of PRs by state (e.g., "3 Draft, 2 Submitted, 1 Approved") | Could |
| FR-005-016 | PR detail SHALL show a computed total amount (sum of Quantity x UnitPrice for all lines) | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-08-01 | Create summary bar layout at the bottom of the page with state count labels | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-08-02 | Implement `DraftCount`, `SubmittedCount`, `ApprovedCount`, `RejectedCount`, `TotalCount` properties | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-08-03 | Implement `UpdateSummaryCounts()` method that recalculates counts from the `PurchaseRequisitions` collection | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-08-04 | Call `UpdateSummaryCounts()` after PR list refresh, conversion, and submission operations | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-08-05 | Display `TotalCount` in the project context panel (e.g., "Total PRs: 6") | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-08-06 | Ensure `LineCount` is displayed in each PR list item template | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-08-07 | Bind `SelectedPRTotal` in the detail view to display the computed total amount with currency formatting | `Views/Pages/PurchaseRequisitionPage.xaml` | S |

## Dependencies
- Depends on: US-005-02 (PR list must be loaded for summary counts to be meaningful)
- Depends on: US-005-03 (PR detail total amount is part of the detail view)
- Blocks: None

## Notes
- The summary bar is positioned at the bottom of the page, below the PR list and detail panels, as shown in the wireframe: "Summary: 1 Draft | 2 Submitted | 3 Approved | 0 Rejected".
- Summary counts are computed from the full `PurchaseRequisitions` collection (not the filtered list), so the counts always reflect the total project state.
- The `TotalCount` property in the project context panel provides a quick overview without scrolling to the summary bar.
- `SelectedPRTotal` (from US-005-03) is the sum of `Quantity * UnitPrice` for all lines in the selected PR. It should be formatted as currency (e.g., "$12,450.00").
- `LineCount` on each `PRDisplayItem` is set during the mapping from `PREntry` (it equals `PREntry.Lines.Count`).
- The `UpdateSummaryCounts()` method should be called in `RefreshListAsync()`, `ConvertBOQToPRAsync()`, and `SubmitPRAsync()` to keep the summary bar current.
