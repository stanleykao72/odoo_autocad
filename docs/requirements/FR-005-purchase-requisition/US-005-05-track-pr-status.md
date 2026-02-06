# US-005-05: Track PR Status

## User Story
**As a** Project Manager,
**I want to** track the status of each PR (draft, submitted, approved, rejected),
**So that** I know where each requisition stands in the approval pipeline.

## Parent Feature
- **FR**: [FR-005-purchase-requisition](../FR-005-purchase-requisition/FR-005-purchase-requisition.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Each PR in the list displays its state as a color-coded badge
- [ ] AC-02: Draft state uses gray badge (background: #E0E0E0, text: #424242, border: #9E9E9E)
- [ ] AC-03: Submitted state uses blue badge (background: #BBDEFB, text: #1565C0, border: #42A5F5)
- [ ] AC-04: Approved state uses green badge (background: #C8E6C9, text: #2E7D32, border: #66BB6A)
- [ ] AC-05: Rejected state uses red badge (background: #FFCDD2, text: #C62828, border: #EF5350)
- [ ] AC-06: The state badge is consistently displayed in both the PR list and the PR detail header
- [ ] AC-07: State changes (e.g., after submission) are reflected in real-time without requiring a manual page refresh

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-005-022 | PR states SHALL be visually distinguished with color-coded badges (draft=gray, submitted=blue, approved=green, rejected=red) | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-005-05-01 | Create reusable state badge XAML template/style with color bindings | `Views/Pages/PurchaseRequisitionPage.xaml` | M |
| TASK-005-05-02 | Implement `StateMap` dictionary mapping state strings to display names and color values | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-05-03 | Compute `StateBadgeBackground`, `StateBadgeForeground`, and `StateDisplay` on `PRDisplayItem` | `ViewModels/PurchaseRequisitionViewModel.cs` | S |
| TASK-005-05-04 | Apply state badge to PR list `ListView` item template | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-05-05 | Apply state badge to PR detail header section | `Views/Pages/PurchaseRequisitionPage.xaml` | S |
| TASK-005-05-06 | Ensure state changes propagate to UI via `ObservableObject` property change notifications | `ViewModels/PurchaseRequisitionViewModel.cs` | S |

## Dependencies
- Depends on: US-005-02 (PR list must be displayed to show status badges)
- Blocks: None

## Notes
- The state mapping is defined as a static dictionary in the ViewModel:
  ```csharp
  private static readonly Dictionary<string, (string Display, string Background, string Foreground)> StateMap = new()
  {
      ["draft"]     = ("Draft",     "#E0E0E0", "#424242"),
      ["submitted"] = ("Submitted", "#BBDEFB", "#1565C0"),
      ["approved"]  = ("Approved",  "#C8E6C9", "#2E7D32"),
      ["rejected"]  = ("Rejected",  "#FFCDD2", "#C62828"),
  };
  ```
- State badges should use `SolidColorBrush` values pre-computed during `PRDisplayItem` creation for binding efficiency.
- When a PR's state changes (e.g., after submit), the `PRDisplayItem` properties must be updated and property change notifications fired to refresh the badge in both the list and detail views.
- Consider creating a custom `StateBadge` UserControl or a DataTemplate for consistent reuse across list and detail views.
