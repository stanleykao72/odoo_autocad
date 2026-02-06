# US-005-04: Submit PR

## User Story
**As a** Project Manager,
**I want to** submit a PR for approval,
**So that** the purchasing team can begin procurement.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: A "Submit for Approval" button is visible in the PR detail panel for PRs in "draft" state
- [ ] AC-02: The Submit button is disabled for PRs that are not in "draft" state
- [ ] AC-03: Clicking Submit triggers a confirmation dialog: "Submit PR {reference} for approval?"
- [ ] AC-04: On confirmation, `IOdooService.SubmitPRAsync(prId)` is called
- [ ] AC-05: On successful submission, the PR state updates to "submitted" in the UI immediately
- [ ] AC-06: On failure, an error message is displayed and the PR remains in "draft" state
- [ ] AC-07: The Submit button is disabled for PRs that have no line items, with a warning: "PR has no line items and cannot be submitted"
- [ ] AC-08: A loading indicator is shown on the Submit button while the submission is in progress

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-017 | Page SHALL provide a "Submit for Approval" button for PRs in "draft" state | Must |
| FR-005-018 | Submit SHALL call `IOdooService.SubmitPRAsync(prId)` and update the PR state on success | Must |
| FR-005-019 | Submit button SHALL be disabled for PRs that are not in "draft" state | Must |
| FR-005-020 | A confirmation dialog SHALL appear before submission: "Submit PR {reference} for approval?" | Must |
| FR-005-021 | On successful submission, the PR state SHALL update to "submitted" in the UI | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-04-01 | Add "Submit for Approval" button to PR detail panel with conditional visibility/enabled binding | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-04-02 | Implement `SubmitPRCommand` async relay command with CanExecute logic (draft state + has lines) | `ViewModels/PurchaseRequisitionViewModel.cs` | M |
| TASK-005-04-03 | Implement `SubmitPRAsync(prId)` method in Odoo service interface and implementation | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | M |
| TASK-005-04-04 | Implement confirmation dialog before submission with PR reference in the message | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-04-05 | Update PR state in the list and detail view after successful submission (draft -> submitted) | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-04-06 | Implement error handling for submission failures with user-friendly messages | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-04-07 | Add loading indicator on Submit button while submission is in progress | `Views/Pages/PurchaseRequisitionPage.xaml` | S |

## Dependencies
- Depends on: US-005-02 (PR list must be loaded to select a PR)
- Depends on: US-005-03 (PR detail view must be visible to access Submit button)
- Blocks: None

## Notes
- The Submit button should bind its enabled state to `SelectedPR.CanSubmit`, which is `true` only when `State == "draft"` and `LineCount > 0`.
- After successful submission, the state badge color for the PR should update from gray (Draft) to blue (Submitted) in both the list and detail views.
- VR-005-006: Submit button is only enabled when State == "draft".
- Error scenarios to handle: network failure, backend rejection (with server-provided reason), PR with no lines.
- The confirmation dialog should use the application's standard dialog service/pattern for consistency.
