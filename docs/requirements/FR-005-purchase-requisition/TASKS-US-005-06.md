# TASKS: US-005-06 — Filter and Sort PRs

> **Parent US**: [US-005-06](US-005-06-filter-and-sort-prs.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 6S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-005-02 (PR list must be loaded before filtering/sorting is meaningful)
- [ ] `ObservableCollection<PRDisplayItem> PurchaseRequisitions` populated with PR data

## Acceptance Criteria
- [ ] AC-01: A filter dropdown is available with options: All, Draft, Submitted, Approved, Rejected
- [ ] AC-02: Selecting a filter value updates the PR list to show only PRs matching that state
- [ ] AC-03: "All" is the default filter value and shows all PRs regardless of state
- [ ] AC-04: Sorting options are available: by Reference (A-Z), by Reference (Z-A), by Date (newest first), by Date (oldest first)
- [ ] AC-05: The default sort order is "Date (newest first)"
- [ ] AC-06: Filter and sort changes are applied immediately without requiring a separate action
- [ ] AC-07: The filtered/sorted results bind to a computed `FilteredPRs` collection that the ListView uses

---

## TASK-005-06-01: Add filter dropdown ComboBox

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-06-03 |

### What to do
- Add a `ComboBox` labeled "Filter:" in the filter/sort bar area between the Conversion Controls and the PR list
- Populate with items: "All", "Draft", "Submitted", "Approved", "Rejected"
- Bind `SelectedItem` to `{Binding SelectedStateFilter, Mode=TwoWay}`
- Set default selected value to "All"
- Style to match the page theme with appropriate width and padding
- Position in the filter bar as shown in the wireframe: `Filter: [All States v]`

### How to verify
- [ ] Filter dropdown is visible with all state options (AC-01)
- [ ] "All" is the default selection (AC-03)

---

## TASK-005-06-02: Add sort order dropdown ComboBox

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-06-04 |

### What to do
- Add a `ComboBox` labeled "Sort:" next to the filter dropdown in the filter/sort bar
- Populate with items: "Newest First", "Oldest First", "Reference (A-Z)", "Reference (Z-A)"
- Bind `SelectedItem` to `{Binding SelectedSortOrder, Mode=TwoWay}`
- Set default selected value to "Newest First"
- Style consistently with the filter dropdown
- Position in the filter bar as shown in the wireframe: `Sort: [Newest First v]`

### How to verify
- [ ] Sort dropdown is visible with all sort options (AC-04)
- [ ] "Newest First" is the default selection (AC-05)

---

## TASK-005-06-03: Implement SelectedStateFilter property

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-06-01 |
| Blocks | TASK-005-06-05 |

### What to do
- Declare `[ObservableProperty] private string _selectedStateFilter = "All";`
- In the partial method `OnSelectedStateFilterChanged(string value)`, call `ApplyFilterAndSort()` to recompute `FilteredPRs`
- Validate that the filter value is one of: "All", "Draft", "Submitted", "Approved", "Rejected" (VR-005-010)
- If an invalid value is received, reset to "All" and log a warning
- Map display values to internal state strings: "Draft" -> "draft", "Submitted" -> "submitted", etc.

### How to verify
- [ ] Changing filter immediately updates the displayed PR list (AC-02, AC-06)
- [ ] Invalid filter values default to "All" (AC-03)

---

## TASK-005-06-04: Implement SelectedSortOrder property

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-06-02 |
| Blocks | TASK-005-06-05 |

### What to do
- Declare `[ObservableProperty] private string _selectedSortOrder = "Newest First";`
- In the partial method `OnSelectedSortOrderChanged(string value)`, call `ApplyFilterAndSort()` to recompute `FilteredPRs`
- Map sort display values to sort logic:
  - "Newest First" -> `OrderByDescending(p => p.CreatedAt)`
  - "Oldest First" -> `OrderBy(p => p.CreatedAt)`
  - "Reference (A-Z)" -> `OrderBy(p => p.Reference)`
  - "Reference (Z-A)" -> `OrderByDescending(p => p.Reference)`
- Default to "Newest First" (descending by CreatedAt)

### How to verify
- [ ] Changing sort order immediately re-sorts the PR list (AC-04, AC-06)
- [ ] Default sort is by date newest first (AC-05)

---

## TASK-005-06-05: Implement FilteredPRs computed collection

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-005-06-03, TASK-005-06-04 |
| Blocks | TASK-005-06-06 |

### What to do
- Declare `public ObservableCollection<PRDisplayItem> FilteredPRs { get; } = new();`
- Implement `ApplyFilterAndSort()` method that:
  1. Starts with `PurchaseRequisitions` as the source
  2. Applies state filter: if `SelectedStateFilter != "All"`, filter by `p.State == selectedState.ToLower()`
  3. Applies sort order using LINQ based on `SelectedSortOrder`
  4. Clears `FilteredPRs` and repopulates with the filtered/sorted results
- Call `ApplyFilterAndSort()` whenever: filter changes, sort changes, or `PurchaseRequisitions` is refreshed
- Consider using `CollectionViewSource` with `ListCollectionView` as a WPF-native alternative for filtering/sorting without rebuilding the collection
- The ListView in XAML binds to `FilteredPRs`, not to `PurchaseRequisitions` directly

### How to verify
- [ ] FilteredPRs reflects both filter and sort applied to PurchaseRequisitions (AC-07)
- [ ] Changes are applied immediately without a separate refresh action (AC-06)
- [ ] "All" filter shows all PRs (AC-03)

---

## TASK-005-06-06: Implement FilterByStateCommand and SortByCommand

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-06-05 |
| Blocks | None |

### What to do
- Declare `IRelayCommand<string> FilterByStateCommand` that sets `SelectedStateFilter` and calls `ApplyFilterAndSort()`
- Declare `IRelayCommand<string> SortByCommand` that sets `SelectedSortOrder` and calls `ApplyFilterAndSort()`
- These commands provide an alternative to ComboBox binding for programmatic filter/sort changes (e.g., clicking a state badge in the summary bar)
- Implement using `[RelayCommand]` attribute from CommunityToolkit.Mvvm

### How to verify
- [ ] FilterByStateCommand changes the active filter (AC-02)
- [ ] SortByCommand changes the active sort order (AC-04)

---

## TASK-005-06-07: Validate filter state values

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Define a `private static readonly HashSet<string> ValidFilterStates = new() { "All", "Draft", "Submitted", "Approved", "Rejected" };`
- In `OnSelectedStateFilterChanged`, check `ValidFilterStates.Contains(value)` (VR-005-010)
- If the value is not in the valid set, reset `SelectedStateFilter = "All"` and log a warning
- Define a mapping dictionary from display names to internal Odoo state strings: `{ "Draft": "draft", "Submitted": "submitted", "Approved": "approved", "Rejected": "rejected" }`
- Use this mapping in `ApplyFilterAndSort()` when comparing against `PRDisplayItem.State`

### How to verify
- [ ] Only valid filter states are accepted (AC-01)
- [ ] Invalid values default to "All" without crashing (AC-03)

---

## Dependency Graph

```
TASK-005-06-01 (filter ComboBox) --> TASK-005-06-03 (SelectedStateFilter)
                                              |
TASK-005-06-02 (sort ComboBox) ---> TASK-005-06-04 (SelectedSortOrder)
                                              |
                                              v
                                     TASK-005-06-05 (FilteredPRs) --> TASK-005-06-06 (commands)

TASK-005-06-07 (validation) ---------> independent
```
