# US-005-07: Conversion Feedback

## User Story
**As a** CAD Engineer,
**I want to** see a confirmation or error when conversion completes,
**So that** I know whether the operation succeeded and what PRs were created.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: On successful conversion, the status area displays a success message including the newly created PR reference (e.g., "Successfully created PR: PR-2025-006")
- [ ] AC-02: On successful conversion, the PR list refreshes and the newly created PR is visually highlighted/selected
- [ ] AC-03: On failure, the status area displays the error message returned by the Odoo API (e.g., "Conversion failed: {error_message}")
- [ ] AC-04: When conversion completes but no PR was created, an informational message is shown: "Conversion completed but no Purchase Requisitions were created. The BOQ may already be fully converted."
- [ ] AC-05: On network timeout, a specific message is shown: "The conversion request timed out. The Odoo server may be processing a large request. Please try again."
- [ ] AC-06: The `LastConversionTime` is updated on successful conversion and displayed in the project context panel
- [ ] AC-07: Error states provide a retry option so the user can attempt conversion again without additional steps

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-006 | On successful conversion, the page SHALL refresh the PR list and show the newly created PR highlighted | Must |
| FR-005-007 | On failure, the page SHALL display the error message returned by the Odoo API | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-07-01 | Add status message area in the conversion controls section with `StatusMessage` binding | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-07-02 | Implement success feedback: set `StatusMessage`, update `LastConversionTime`, refresh list, select new PR | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-07-03 | Implement failure feedback: set `StatusMessage` with Odoo error details, enable retry | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-07-04 | Implement empty result feedback for when conversion produces no PRs | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-07-05 | Implement network timeout detection and specific timeout error message | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-07-06 | Add `LastConversionTime` display in the project context panel with formatted timestamp | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-07-07 | Style status message area with visual differentiation for success (green), error (red), and info (blue) states | `Views/Pages/PurchaseRequisitionPage.xaml` | S |

## Dependencies
- Depends on: US-005-01 (Conversion must be triggered before feedback can be displayed)
- Blocks: None

## Notes
- This user story is closely related to US-005-01 (Convert BOQ to PR) and covers the feedback/result-display aspects of the conversion flow. The implementation tasks overlap with the conversion flow in the ViewModel.
- The status message area in the conversion controls section serves as the primary feedback channel. It should use visual styling to differentiate success, error, and informational states:
  - Success: green text or green accent bar
  - Error: red text or red accent bar
  - Info: blue text or blue accent bar
- The `StatusMessage` property is bound from the ViewModel and updated throughout the conversion lifecycle:
  1. "Converting BOQ to Purchase Requisition..." (during conversion)
  2. "Successfully created PR: {reference}" (on success)
  3. "Conversion failed: {error_message}" (on failure)
  4. "Conversion completed but no Purchase Requisitions were created." (on empty result)
- The `LastConversionTime` is displayed in the project context panel as a formatted timestamp (e.g., "Last Conversion: 14:30").
- Error scenarios are defined in the FR document section 9 (Error Handling) and should map to specific user-facing messages.
