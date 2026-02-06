# TASKS: US-005-02 — View PR List

> **Parent US**: [US-005-02](US-005-02-view-pr-list.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 9 | **Effort**: 7S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (Odoo connected -- Odoo connection must be active to fetch PR data)
- [ ] FR-003 Odoo Connection active (`IOdooService.IsConnected`)
- [ ] Valid ProjectId must exist in the application context (VR-005-002)

## Acceptance Criteria
- [ ] AC-01: The page displays a list of all Purchase Requisitions for the current project
- [ ] AC-02: Each PR list entry shows: Reference, State (as a color-coded badge), Line Count, and Created Date
- [ ] AC-03: The PR list loads automatically when the page is opened and a valid project context exists
- [ ] AC-04: A manual "Refresh" button is available to reload the PR list
- [ ] AC-05: When no PRs exist for the project, an empty state message is displayed: "No Purchase Requisitions found for this project" with guidance on how to create one
- [ ] AC-06: A loading indicator is shown while the PR list is being fetched from Odoo
- [ ] AC-07: If the PR list fetch fails, an error message is displayed with a retry option

---

## TASK-005-02-01: Create PR list ListView layout

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-005-02-02, TASK-005-02-05, TASK-005-02-06 |

### What to do
- Create a `ListView` in the left panel of the master-detail layout
- Define `ItemTemplate` as a `DataTemplate` with columns for: Reference (text), State (badge Border with Background/Foreground binding), Line Count (text), Created Date (formatted text)
- Bind `ItemsSource` to `{Binding FilteredPRs}` (the filtered/sorted computed collection, not raw `PurchaseRequisitions`)
- Bind `SelectedItem` to `{Binding SelectedPR, Mode=TwoWay}`
- Style the ListView with alternating row colors for readability
- PR Reference must show "N/A" if missing (VR-005-008)

### How to verify
- [ ] ListView displays all PRs for the current project (AC-01)
- [ ] Each entry shows Reference, State badge, Line Count, and Created Date (AC-02)

---

## TASK-005-02-02: Implement ObservableCollection and data binding

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-02-01 |
| Blocks | TASK-005-02-04, TASK-005-02-09 |

### What to do
- Declare `ObservableCollection<PRDisplayItem> PurchaseRequisitions` property
- Declare `[ObservableProperty] private PRDisplayItem? _selectedPR;` with property change notification
- Initialize `PurchaseRequisitions` in the constructor
- When `SelectedPR` changes, trigger loading of the selected PR's line items (for US-005-03)
- Ensure `ObservableCollection` fires `CollectionChanged` events for UI updates

### How to verify
- [ ] PR list is bound to the ListView and displays correctly (AC-01)
- [ ] Selecting a PR updates `SelectedPR` property (AC-01)

---

## TASK-005-02-03: Implement GetPurchaseRequisitionsAsync in Odoo service

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Core/Odoo/IOdooService.cs`, `OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-005-02-04 |

### What to do
- The interface method `GetPurchaseRequisitionsAsync(int projectId)` already exists in `IOdooService.cs`
- Verify/enhance `OdooService.GetPurchaseRequisitionsAsync()` to query `purchase.requisition` with `search_read` filtering by `project_id`
- Include fields: `id`, `name` (Reference), `state`, `line_ids`, `create_date`
- For each PR, also fetch line count by resolving `line_ids` array length
- Parse `create_date` from Odoo format to `DateTime`
- Map results to `List<PREntry>` with all fields populated including `Lines` count

### How to verify
- [ ] All PRs for the given project are returned (AC-01)
- [ ] Each PREntry includes Reference, State, LineCount, and CreatedAt (AC-02)

---

## TASK-005-02-04: Implement auto-load on page open

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-02-02, TASK-005-02-03 |
| Blocks | None |

### What to do
- Implement an `InitializeAsync()` or `OnNavigatedTo()` method that is called when the PurchaseRequisitionPage is loaded
- Check `IOdooService.IsConnected` (VR-005-001) and `ProjectId.HasValue` (VR-005-002)
- If both valid, call `RefreshListAsync()` to load PRs automatically
- If Odoo not connected, set `StatusMessage = "Cannot access Purchase Requisitions. Please connect to Odoo first."` and disable controls
- If no project context, set `StatusMessage = "No project selected. Open a drawing in AutoCAD or select a project to view PRs."`
- Wire the page `Loaded` event to call the ViewModel's `InitializeAsync()`

### How to verify
- [ ] PR list loads automatically when page opens with valid project context (AC-03)
- [ ] Appropriate messages shown when Odoo not connected or no project selected (AC-05)

---

## TASK-005-02-05: Add Refresh button

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-02-01 |
| Blocks | None |

### What to do
- Add a "Refresh" button in the Conversion Controls area, next to the "Convert BOQ to PR" button
- Bind `Command` to `{Binding RefreshListCommand}`
- Implement `IAsyncRelayCommand RefreshListCommand` in the ViewModel
- `RefreshListAsync()` calls `_odooService.GetPurchaseRequisitionsAsync(ProjectId.Value)`, clears `PurchaseRequisitions`, maps results to `PRDisplayItem`, and repopulates the collection
- Disable the Refresh button while `IsLoading == true`

### How to verify
- [ ] Refresh button is visible and triggers PR list reload (AC-04)
- [ ] Refresh button is disabled while loading (AC-06)

---

## TASK-005-02-06: Implement empty state UI

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-02-01 |
| Blocks | None |

### What to do
- Add a `TextBlock` or `StackPanel` that is visible when `PurchaseRequisitions.Count == 0` and `IsLoading == false`
- Display message: "No Purchase Requisitions found for this project"
- Add guidance text: "Use the 'Convert BOQ to PR' button above to create a Purchase Requisition from your BOQ entries."
- Use `Visibility` binding with a value converter: visible when collection is empty, collapsed otherwise
- Style with centered text, muted color, and appropriate icon

### How to verify
- [ ] Empty state message displays when no PRs exist (AC-05)
- [ ] Guidance text helps user understand next steps (AC-05)

---

## TASK-005-02-07: Add IsLoading state with loading indicator

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Declare `[ObservableProperty] private bool _isLoading;`
- Set `IsLoading = true` at the start of `RefreshListAsync()` and `false` in the `finally` block
- Bind a `ProgressBar` or `ProgressRing` in XAML to `{Binding IsLoading}` for visibility
- The loading indicator should overlay or replace the ListView content while loading
- Ensure `RefreshListCommand.CanExecute` returns `false` when `IsLoading == true` to prevent concurrent loads

### How to verify
- [ ] Loading indicator is visible while PR list is being fetched (AC-06)
- [ ] Loading indicator hides when fetch completes or fails (AC-06)

---

## TASK-005-02-08: Implement error handling for PR list fetch failures

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Wrap `RefreshListAsync` body in try/catch
- Catch `HttpRequestException` and set `StatusMessage = "Failed to load Purchase Requisitions: {error details}"`
- Catch general `Exception` and display user-friendly error message
- Add a retry mechanism: set a `HasError` flag that shows a "Retry" button in the UI
- Bind the Retry button to `RefreshListCommand`
- Ensure `IsLoading = false` in the `finally` block

### How to verify
- [ ] Error message displayed when PR list fetch fails (AC-07)
- [ ] Retry button allows the user to attempt loading again (AC-07)

---

## TASK-005-02-09: Map PREntry to PRDisplayItem with computed properties

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-02-02 |
| Blocks | None |

### What to do
- Create a `MapToPRDisplayItem(PREntry entry)` method in the ViewModel
- Set `PRDisplayItem.Reference` = entry.Reference, or "N/A" if empty (VR-005-008)
- Set `PRDisplayItem.State` = entry.State
- Set `PRDisplayItem.StateDisplay` from `StateMap` dictionary (e.g., "draft" -> "Draft")
- Set `PRDisplayItem.StateBadgeBackground` and `StateBadgeForeground` as `SolidColorBrush` from hex colors in `StateMap`
- Set `PRDisplayItem.LineCount` = entry.Lines.Count
- Set `PRDisplayItem.CreatedAtDisplay` = entry.CreatedAt.ToString("MM/dd") or similar format
- Set `PRDisplayItem.CanSubmit` = (entry.State == "draft" && entry.Lines.Count > 0)
- Use this mapper in `RefreshListAsync` when populating the collection

### How to verify
- [ ] Reference shows "N/A" when empty (AC-02)
- [ ] State badge colors are correctly mapped (AC-02)
- [ ] Line Count and Date are formatted for display (AC-02)

---

## Dependency Graph

```
TASK-005-02-03 (GetPurchaseRequisitionsAsync)
       |
       v
TASK-005-02-04 (auto-load on page open) <-- TASK-005-02-02 (ObservableCollection)
                                                    ^
                                                    |
TASK-005-02-01 (ListView layout) -------> TASK-005-02-02
       |                                            |
       +--------> TASK-005-02-05 (Refresh button)   +--------> TASK-005-02-09 (PREntry mapper)
       +--------> TASK-005-02-06 (empty state)

TASK-005-02-07 (IsLoading) ----------> independent
TASK-005-02-08 (error handling) -----> independent
```
