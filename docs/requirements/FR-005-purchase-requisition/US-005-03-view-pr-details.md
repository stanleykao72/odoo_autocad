# US-005-03: View PR Details

## User Story
**As a** CAD Engineer,
**I want to** view the line items within a specific PR,
**So that** I can verify the correct products and quantities before submission.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Selecting a PR from the list displays its detail view with all line items
- [ ] AC-02: The PR detail header shows: Reference, State (with badge), Created Date, and total line count
- [ ] AC-03: PR lines are displayed in a DataGrid with columns: Product Name, Quantity, Unit of Measure, Unit Price
- [ ] AC-04: A computed total amount (sum of Quantity x UnitPrice for all lines) is displayed in the detail view
- [ ] AC-05: Lines with zero or negative quantity are visually flagged as warnings
- [ ] AC-06: When no PR is selected, the detail panel shows a placeholder message

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-013 | Selecting a PR from the list SHALL display its detail view with all line items | Must |
| FR-005-014 | PR detail SHALL display: Reference, State, Created Date, and total line count | Must |
| FR-005-015 | PR lines SHALL display in a DataGrid with columns: Product Name, Quantity, Unit of Measure, Unit Price | Must |
| FR-005-016 | PR detail SHALL show a computed total amount (sum of Quantity x UnitPrice for all lines) | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-03-01 | Create PR detail panel layout with header info section (Reference, State, Date, Line Count) | `Views/Pages/PurchaseRequisitionPage.xaml` | M |
| TASK-005-03-02 | Create PR lines DataGrid with columns: Product Name, Quantity, UoM, Unit Price, Line Total | `Views/Pages/PurchaseRequisitionPage.xaml` | M |
| TASK-005-03-03 | Implement `SelectPRCommand` that loads line items into `SelectedPRLines` when a PR is selected | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-03-04 | Map `PRLine` model to `PRLineDisplayItem` with computed `LineTotal` (Quantity x UnitPrice) | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-03-05 | Implement `SelectedPRTotal` computed property (sum of all line totals) | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-03-06 | Add quantity validation highlighting for zero/negative values (VR-005-009) | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-03-07 | Add empty/placeholder state for detail panel when no PR is selected | `Views/Pages/PurchaseRequisitionPage.xaml` | S |

## Dependencies
- Depends on: US-005-02 (PR list must be loaded for the user to select a PR)
- Blocks: US-005-04 (Submit button appears in the detail panel)

## Notes
- The detail panel is the right side of the master-detail layout. It updates whenever the user selects a different PR in the list.
- `PRLineDisplayItem.LineTotal` is a computed property: `Quantity * UnitPrice`. If UnitPrice is null, LineTotal should also display as null or "--".
- The total amount (`SelectedPRTotal`) is the sum of all `LineTotal` values for the selected PR's lines.
- Quantity validation (VR-005-009): lines with `Quantity <= 0` should be highlighted with a warning style but still displayed.
- The detail view binds to `SelectedPRLines` via `{Binding SelectedPRLines}` on the DataGrid.
