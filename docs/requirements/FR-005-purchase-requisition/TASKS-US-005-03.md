# TASKS: US-005-03 — View PR Details

> **Parent US**: [US-005-03](US-005-03-view-pr-details.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 5S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-005-02 (PR list must be loaded for the user to select a PR)
- [ ] `PRDisplayItem` and `PRLineDisplayItem` data models defined
- [ ] `IOdooService.GetPurchaseRequisitionsAsync()` returns `PREntry` with populated `Lines`

## Acceptance Criteria
- [ ] AC-01: Selecting a PR from the list displays its detail view with all line items
- [ ] AC-02: The PR detail header shows: Reference, State (with badge), Created Date, and total line count
- [ ] AC-03: PR lines are displayed in a DataGrid with columns: Product Name, Quantity, Unit of Measure, Unit Price
- [ ] AC-04: A computed total amount (sum of Quantity x UnitPrice for all lines) is displayed in the detail view
- [ ] AC-05: Lines with zero or negative quantity are visually flagged as warnings
- [ ] AC-06: When no PR is selected, the detail panel shows a placeholder message

---

## TASK-005-03-01: Create PR detail panel layout with header info

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-005-03-02, TASK-005-03-07 |

### What to do
- Create the right panel of the master-detail layout for PR detail display
- Add header section with bindings: `{Binding SelectedPR.Reference}` for PR reference, `{Binding SelectedPR.StateDisplay}` with state badge (using `StateBadgeBackground` and `StateBadgeForeground`), `{Binding SelectedPR.CreatedAtDisplay}` for created date, `{Binding SelectedPR.LineCount}` for line count
- Use a `StackPanel` or `Grid` for the header layout with labels: "Reference:", "State:", "Created:", "Lines:"
- Apply the state badge `Border` template consistent with the list view badges (reuse from US-005-05)
- Bind the entire detail panel visibility to `{Binding SelectedPR, Converter={StaticResource NullToVisibilityConverter}}`

### How to verify
- [ ] Detail header shows Reference, State badge, Created Date, and Line Count (AC-02)
- [ ] Detail panel is visible only when a PR is selected (AC-01)

---

## TASK-005-03-02: Create PR lines DataGrid

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | M |
| Depends On | TASK-005-03-01 |
| Blocks | TASK-005-03-06 |

### What to do
- Add a `DataGrid` below the detail header, bound to `{Binding SelectedPRLines}`
- Define columns: Product Name (`DataGridTextColumn` bound to `ProductName`), Quantity (`DataGridTextColumn` bound to `Quantity` with numeric formatting), Unit of Measure (`DataGridTextColumn` bound to `UnitOfMeasure`), Unit Price (`DataGridTextColumn` bound to `UnitPrice` with currency formatting `{0:N2}`), Line Total (`DataGridTextColumn` bound to `LineTotal` with currency formatting)
- Set `DataGrid.IsReadOnly = true` and `AutoGenerateColumns = false`
- Add a summary row or footer showing the total amount: `{Binding SelectedPRTotal, StringFormat='Total: ${0:N2}'}`
- Style with column headers, gridlines, and alternating row backgrounds

### How to verify
- [ ] DataGrid displays columns: Product Name, Quantity, UoM, Unit Price (AC-03)
- [ ] Total amount is displayed below the DataGrid (AC-04)

---

## TASK-005-03-03: Implement SelectPRCommand to load line items

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-03-04, TASK-005-03-05 |

### What to do
- Implement `IAsyncRelayCommand<PRDisplayItem> SelectPRCommand` or handle via `SelectedPR` property setter
- When `SelectedPR` changes (via property change partial method `OnSelectedPRChanged`), load the selected PR's line items
- Get line data from the `PREntry.Lines` collection (already fetched with the PR) or fetch from Odoo if needed
- Map `PRLine` items to `PRLineDisplayItem` instances and populate `SelectedPRLines` collection
- Clear `SelectedPRLines` when `SelectedPR` is set to null
- Calculate and set `SelectedPRTotal`

### How to verify
- [ ] Selecting a PR from the list loads its line items into the detail DataGrid (AC-01)
- [ ] Deselecting clears the detail panel (AC-06)

---

## TASK-005-03-04: Map PRLine to PRLineDisplayItem with computed LineTotal

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-03-03 |
| Blocks | None |

### What to do
- Create `MapToPRLineDisplayItem(PRLine line)` method
- Set `PRLineDisplayItem.ProductId` = line.ProductId
- Set `PRLineDisplayItem.ProductName` = line.ProductName
- Set `PRLineDisplayItem.Quantity` = line.Quantity
- Set `PRLineDisplayItem.UnitOfMeasure` = line.UnitOfMeasure
- Set `PRLineDisplayItem.UnitPrice` = line.UnitPrice
- Compute `PRLineDisplayItem.LineTotal` = line.UnitPrice.HasValue ? line.Quantity * line.UnitPrice.Value : null
- If `UnitPrice` is null, `LineTotal` should display as "--" in the DataGrid (use a value converter or StringFormat)

### How to verify
- [ ] LineTotal is correctly computed as Quantity x UnitPrice (AC-04)
- [ ] Null UnitPrice results in null/blank LineTotal display (AC-04)

---

## TASK-005-03-05: Implement SelectedPRTotal computed property

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-03-03 |
| Blocks | None |

### What to do
- Declare `[ObservableProperty] private decimal _selectedPRTotal;`
- Calculate `SelectedPRTotal` as the sum of all `PRLineDisplayItem.LineTotal` values where `LineTotal` is not null
- Update `SelectedPRTotal` whenever `SelectedPRLines` is repopulated (in `OnSelectedPRChanged`)
- If all lines have null `LineTotal`, set `SelectedPRTotal = 0`
- Bind in XAML with currency formatting: `{Binding SelectedPRTotal, StringFormat='${0:N2}'}`

### How to verify
- [ ] Total amount correctly sums all line totals (AC-04)
- [ ] Total updates when a different PR is selected (AC-04)

---

## TASK-005-03-06: Add quantity validation highlighting

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-03-02 |
| Blocks | None |

### What to do
- Add a `DataGrid.RowStyle` or `DataTrigger` that checks if `Quantity <= 0`
- When quantity is zero or negative, apply a warning background color (e.g., light yellow #FFF9C4 or light orange #FFE0B2)
- Add a warning icon or tooltip: "Quantity is zero or negative" (VR-005-009)
- Lines with invalid quantity should still be displayed, not hidden
- Use a `DataGridRow` style with `DataTrigger` binding to the Quantity property

### How to verify
- [ ] Lines with zero or negative quantity are visually flagged (AC-05)
- [ ] Flagged lines are still visible and not filtered out (AC-05)

---

## TASK-005-03-07: Add placeholder state for empty detail panel

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-03-01 |
| Blocks | None |

### What to do
- Add a `TextBlock` or `StackPanel` in the detail panel area that is visible when `SelectedPR == null`
- Display placeholder text: "Select a Purchase Requisition from the list to view its details"
- Center the text vertically and horizontally within the detail panel
- Use muted/gray text color for the placeholder
- Toggle visibility using `{Binding SelectedPR, Converter={StaticResource NullToVisibilityConverter}}` (inverse for placeholder)

### How to verify
- [ ] Placeholder message shows when no PR is selected (AC-06)
- [ ] Placeholder hides when a PR is selected (AC-06)

---

## Dependency Graph

```
TASK-005-03-01 (detail panel layout)
       |
       +--------> TASK-005-03-02 (DataGrid) --> TASK-005-03-06 (quantity validation)
       +--------> TASK-005-03-07 (placeholder state)

TASK-005-03-03 (SelectPRCommand)
       |
       +--------> TASK-005-03-04 (PRLine mapper)
       +--------> TASK-005-03-05 (SelectedPRTotal)
```
