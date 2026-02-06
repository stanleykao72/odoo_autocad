# US-005-06: Filter and Sort PRs

## User Story
**As a** CAD Engineer,
**I want to** filter and sort PRs by status, date, or reference,
**So that** I can quickly find specific requisitions in large projects.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: A filter dropdown is available with options: All, Draft, Submitted, Approved, Rejected
- [ ] AC-02: Selecting a filter value updates the PR list to show only PRs matching that state
- [ ] AC-03: "All" is the default filter value and shows all PRs regardless of state
- [ ] AC-04: Sorting options are available: by Reference (A-Z), by Reference (Z-A), by Date (newest first), by Date (oldest first)
- [ ] AC-05: The default sort order is "Date (newest first)"
- [ ] AC-06: Filter and sort changes are applied immediately without requiring a separate action
- [ ] AC-07: The filtered/sorted results bind to a computed `FilteredPRs` collection that the ListView uses

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-023 | Page SHALL provide a filter dropdown to filter PRs by state (All, Draft, Submitted, Approved, Rejected) | Should |
| FR-005-024 | Page SHALL provide sorting options: by Reference (A-Z), by Date (newest first), by State | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-06-01 | Add filter dropdown ComboBox to the filter/sort bar area of the page | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-06-02 | Add sort order dropdown ComboBox to the filter/sort bar area of the page | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-06-03 | Implement `SelectedStateFilter` property with change notification that triggers re-filtering | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-06-04 | Implement `SelectedSortOrder` property with change notification that triggers re-sorting | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-06-05 | Implement `FilteredPRs` computed collection applying filter and sort to `PurchaseRequisitions` | `ViewModels/PurchaseRequisitionViewModel.cs` | M |
| TASK-005-06-06 | Implement `FilterByStateCommand` and `SortByCommand` relay commands | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-06-07 | Validate that filter state values are restricted to allowed values (VR-005-010) | `ViewModels/PurchaseRequisitionViewModel.cs` | S |

## Dependencies
- Depends on: US-005-02 (PR list must be loaded before filtering/sorting is meaningful)
- Blocks: None

## Notes
- The `FilteredPRs` property is a computed/derived collection that applies both the state filter and sort order to the raw `PurchaseRequisitions` collection. The ListView should bind to `FilteredPRs`, not directly to `PurchaseRequisitions`.
- VR-005-010: Filter state values must be one of: "All", "draft", "submitted", "approved", "rejected". Invalid values should default to "All".
- Sort order options map to internal values:
  - "Newest First" -> sort by `CreatedAt` descending
  - "Oldest First" -> sort by `CreatedAt` ascending
  - "Reference (A-Z)" -> sort by `Reference` ascending
  - "Reference (Z-A)" -> sort by `Reference` descending
- When the filter or sort changes, the `FilteredPRs` collection should be recomputed using LINQ and the `ObservableCollection` updated. Consider using `CollectionViewSource` for WPF-native filtering and sorting.
- The filter/sort bar is positioned between the conversion controls area and the PR list, as shown in the wireframe.
