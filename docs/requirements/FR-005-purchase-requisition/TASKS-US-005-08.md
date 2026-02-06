# TASKS: US-005-08 — View PR Totals

> **Parent US**: [US-005-08](US-005-08-view-pr-totals.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 7S + 0M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-005-02 (PR list must be loaded for summary counts to be meaningful)
- [ ] US-005-03 (PR detail total amount is part of the detail view, `SelectedPRTotal` computed)
- [ ] `ObservableCollection<PRDisplayItem> PurchaseRequisitions` populated

## Acceptance Criteria
- [ ] AC-01: A summary bar at the bottom of the page displays the count of PRs by state (e.g., "1 Draft | 2 Submitted | 3 Approved | 0 Rejected")
- [ ] AC-02: The summary bar updates automatically when PRs are added, removed, or change state
- [ ] AC-03: Each PR in the list displays its line count alongside the reference and state
- [ ] AC-04: The PR detail view shows a computed total amount (sum of Quantity x UnitPrice for all lines)
- [ ] AC-05: The total PR count is displayed in the project context panel
- [ ] AC-06: Summary counts are recalculated when the PR list is refreshed

---

## TASK-005-08-01: Create summary bar layout at page bottom

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `Border` or `StackPanel` at the bottom of the page layout, below the PR list and detail panels
- Display state counts in the format: "1 Draft | 2 Submitted | 3 Approved | 0 Rejected"
- Bind labels to: `{Binding DraftCount}`, `{Binding SubmittedCount}`, `{Binding ApprovedCount}`, `{Binding RejectedCount}`
- Use `TextBlock` elements with `Run` elements for colored count values (each count colored with its state color)
- Style with a subtle background color, horizontal orientation, and center alignment
- Use pipe "|" separators between state groups
- Apply the state badge colors to each count number for visual consistency

### How to verify
- [ ] Summary bar is visible at the bottom of the page (AC-01)
- [ ] Counts are displayed in the correct format with state labels (AC-01)

---

## TASK-005-08-02: Implement state count properties

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-08-03 |

### What to do
- Declare `[ObservableProperty] private int _draftCount;`
- Declare `[ObservableProperty] private int _submittedCount;`
- Declare `[ObservableProperty] private int _approvedCount;`
- Declare `[ObservableProperty] private int _rejectedCount;`
- Declare `[ObservableProperty] private int _totalCount;`
- These properties fire property change notifications automatically via `[ObservableProperty]`
- All counts should reflect the full `PurchaseRequisitions` collection (not the filtered list)

### How to verify
- [ ] All five count properties are declared and bindable (AC-01)
- [ ] Counts reflect the total project state, not filtered view (AC-01)

---

## TASK-005-08-03: Implement UpdateSummaryCounts method

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-08-02 |
| Blocks | TASK-005-08-04 |

### What to do
- Implement `private void UpdateSummaryCounts()` method that:
  - `DraftCount = PurchaseRequisitions.Count(p => p.State == "draft");`
  - `SubmittedCount = PurchaseRequisitions.Count(p => p.State == "submitted");`
  - `ApprovedCount = PurchaseRequisitions.Count(p => p.State == "approved");`
  - `RejectedCount = PurchaseRequisitions.Count(p => p.State == "rejected");`
  - `TotalCount = PurchaseRequisitions.Count;`
- Use LINQ `Count()` with predicate on the full `PurchaseRequisitions` collection
- Handle the case where `PurchaseRequisitions` is null or empty (all counts = 0)

### How to verify
- [ ] Counts are correctly computed from the PR collection (AC-01)
- [ ] All states are counted independently (AC-01)

---

## TASK-005-08-04: Call UpdateSummaryCounts after state-changing operations

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-08-03 |
| Blocks | None |

### What to do
- Call `UpdateSummaryCounts()` at the end of `RefreshListAsync()` after repopulating `PurchaseRequisitions`
- Call `UpdateSummaryCounts()` at the end of `ConvertBOQToPRAsync()` after successful conversion and list refresh
- Call `UpdateSummaryCounts()` at the end of `SubmitPRAsync()` after successful state change (draft -> submitted)
- This ensures the summary bar is always current after any operation that changes PR state or count
- Consider debouncing if multiple rapid updates occur

### How to verify
- [ ] Summary counts update after PR list refresh (AC-06)
- [ ] Summary counts update after conversion adds a new PR (AC-02)
- [ ] Summary counts update after submission changes a PR state (AC-02)

---

## TASK-005-08-05: Display TotalCount in project context panel

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Project Context panel at the top of the page, add a label: "Total PRs:"
- Bind to `{Binding TotalCount}` to show the total number of PRs (e.g., "Total PRs: 6")
- Position next to the project name and PR number fields as shown in the wireframe
- Update automatically when `TotalCount` property changes
- When `TotalCount == 0`, display "Total PRs: 0" (not hidden)

### How to verify
- [ ] Total PR count is visible in the project context panel (AC-05)
- [ ] Count updates when PRs are added or removed (AC-05)

---

## TASK-005-08-06: Ensure LineCount displayed in PR list item template

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the ListView `ItemTemplate` (created in TASK-005-02-01), ensure the `LineCount` column is visible
- Bind to `{Binding LineCount}` with a format like "{0} lines" or just the number
- Position the line count between the State badge and the Date columns
- Style with a muted text color to differentiate from the primary Reference text
- The `LineCount` is set during `MapToPRDisplayItem` from `PREntry.Lines.Count`

### How to verify
- [ ] Line count is visible in each PR list item (AC-03)
- [ ] Line count reflects the actual number of lines in the PR (AC-03)

---

## TASK-005-08-07: Bind SelectedPRTotal in detail view with currency formatting

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the PR detail panel (below the DataGrid), bind the total amount display to `{Binding SelectedPRTotal, StringFormat='Total: ${0:N2}'}`
- Position the total below the DataGrid, right-aligned for financial convention
- Style with bold text and slightly larger font size to emphasize the total
- When `SelectedPRTotal == 0` (all lines have null UnitPrice), display "Total: $0.00"
- Add a horizontal separator line above the total for visual clarity (mimicking a sum line)

### How to verify
- [ ] Total amount is displayed with currency formatting in the detail view (AC-04)
- [ ] Total updates when a different PR is selected (AC-04)

---

## Dependency Graph

```
TASK-005-08-02 (count properties) --> TASK-005-08-03 (UpdateSummaryCounts) --> TASK-005-08-04 (call after operations)

TASK-005-08-01 (summary bar XAML) ---------> independent
TASK-005-08-05 (TotalCount in context) ----> independent
TASK-005-08-06 (LineCount in list) --------> independent
TASK-005-08-07 (SelectedPRTotal binding) --> independent
```
