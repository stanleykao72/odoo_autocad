# TASKS: US-005-05 — Track PR Status

> **Parent US**: [US-005-05](US-005-05-track-pr-status.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 5S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-005-02 (PR list must be displayed to show status badges)
- [ ] `PRDisplayItem` model with `State`, `StateDisplay`, `StateBadgeBackground`, `StateBadgeForeground` properties defined

## Acceptance Criteria
- [ ] AC-01: Each PR in the list displays its state as a color-coded badge
- [ ] AC-02: Draft state uses gray badge (background: #E0E0E0, text: #424242, border: #9E9E9E)
- [ ] AC-03: Submitted state uses blue badge (background: #BBDEFB, text: #1565C0, border: #42A5F5)
- [ ] AC-04: Approved state uses green badge (background: #C8E6C9, text: #2E7D32, border: #66BB6A)
- [ ] AC-05: Rejected state uses red badge (background: #FFCDD2, text: #C62828, border: #EF5350)
- [ ] AC-06: The state badge is consistently displayed in both the PR list and the PR detail header
- [ ] AC-07: State changes (e.g., after submission) are reflected in real-time without requiring a manual page refresh

---

## TASK-005-05-01: Create reusable state badge XAML template

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-005-05-04, TASK-005-05-05 |

### What to do
- Create a `DataTemplate` or `ControlTemplate` for the state badge as a reusable resource
- The badge is a `Border` with rounded corners (CornerRadius="4"), padding, and a `TextBlock` inside
- Bind `Border.Background` to `{Binding StateBadgeBackground}`
- Bind `Border.BorderBrush` to a computed border brush (or derive from the state)
- Bind `TextBlock.Foreground` to `{Binding StateBadgeForeground}`
- Bind `TextBlock.Text` to `{Binding StateDisplay}`
- Define as a `Page.Resources` entry with key `StateBadgeTemplate` for reuse in both list and detail views
- Consider creating a `StatusToColorConverter` IValueConverter as an alternative approach for direct State string binding

### How to verify
- [ ] Badge template renders with correct colors and text (AC-01)
- [ ] Template is reusable across list and detail views (AC-06)

---

## TASK-005-05-02: Implement StateMap dictionary for state-to-display mapping

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-05-03 |

### What to do
- Define a `private static readonly Dictionary<string, (string Display, string Background, string Foreground, string Border)> StateMap` with entries:
  - `["draft"]     = ("Draft",     "#E0E0E0", "#424242", "#9E9E9E")`
  - `["submitted"] = ("Submitted", "#BBDEFB", "#1565C0", "#42A5F5")`
  - `["approved"]  = ("Approved",  "#C8E6C9", "#2E7D32", "#66BB6A")`
  - `["rejected"]  = ("Rejected",  "#FFCDD2", "#C62828", "#EF5350")`
- Add a helper method `GetStateInfo(string state)` that returns the tuple, defaulting to draft values for unknown states
- This dictionary is used by `MapToPRDisplayItem()` (TASK-005-02-09) when creating display items

### How to verify
- [ ] All four states map to correct display names and color values (AC-02, AC-03, AC-04, AC-05)
- [ ] Unknown states fall back to draft appearance (AC-01)

---

## TASK-005-05-03: Compute badge properties on PRDisplayItem

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-05-02 |
| Blocks | TASK-005-05-06 |

### What to do
- In `MapToPRDisplayItem(PREntry entry)`, use `GetStateInfo(entry.State)` to populate:
  - `PRDisplayItem.StateDisplay` = stateInfo.Display
  - `PRDisplayItem.StateBadgeBackground` = new SolidColorBrush(ColorConverter.ConvertFromString(stateInfo.Background))
  - `PRDisplayItem.StateBadgeForeground` = new SolidColorBrush(ColorConverter.ConvertFromString(stateInfo.Foreground))
- Pre-compute `SolidColorBrush` values during mapping for efficient XAML binding (avoid converters at runtime)
- Ensure `PRDisplayItem` extends `ObservableObject` so that property changes trigger UI updates
- Add a `UpdateStateProperties(string newState)` method on `PRDisplayItem` that re-computes badge properties when state changes

### How to verify
- [ ] Badge background, foreground, and display text are correctly set from StateMap (AC-02, AC-03, AC-04, AC-05)
- [ ] State property changes propagate to UI via ObservableObject notifications (AC-07)

---

## TASK-005-05-04: Apply state badge to PR list ListView item template

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-05-01 |
| Blocks | None |

### What to do
- In the ListView `ItemTemplate`, include the state badge `Border` element using the `StateBadgeTemplate`
- Position the badge in the State column area of each list item, between Reference and Date
- Bind the badge's `Background`, `Foreground`, and `Text` to the `PRDisplayItem` properties
- Ensure the badge is vertically centered within the list item row
- The badge should have a fixed minimum width for visual consistency across different state text lengths

### How to verify
- [ ] Each PR in the list shows a color-coded state badge (AC-01)
- [ ] Draft=gray, Submitted=blue, Approved=green, Rejected=red (AC-02, AC-03, AC-04, AC-05)

---

## TASK-005-05-05: Apply state badge to PR detail header

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-05-01 |
| Blocks | None |

### What to do
- In the PR detail panel header section (created in TASK-005-03-01), include the same state badge template
- Bind to `{Binding SelectedPR.StateBadgeBackground}`, `{Binding SelectedPR.StateBadgeForeground}`, `{Binding SelectedPR.StateDisplay}`
- Position the badge next to the PR Reference text in the detail header
- Ensure the badge style is identical to the list view badge for visual consistency (same CornerRadius, padding, font size)

### How to verify
- [ ] State badge in detail header matches the badge in the list (AC-06)
- [ ] Badge colors are identical between list and detail views (AC-06)

---

## TASK-005-05-06: Ensure real-time state change propagation

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-05-03 |
| Blocks | None |

### What to do
- Ensure `PRDisplayItem` inherits from `ObservableObject` (CommunityToolkit.Mvvm)
- Use `[ObservableProperty]` attributes or manual `OnPropertyChanged` calls for `State`, `StateDisplay`, `StateBadgeBackground`, `StateBadgeForeground`, `CanSubmit`
- When `SubmitPRAsync` succeeds (US-005-04), call `SelectedPR.UpdateStateProperties("submitted")` which fires all property change notifications
- Verify that both the list item and the detail header update their badge colors without requiring a full list refresh
- Test that `CollectionChanged` is not needed for in-place property updates on existing items (only `PropertyChanged` on the item)

### How to verify
- [ ] State changes after submission are reflected immediately in both list and detail (AC-07)
- [ ] No manual page refresh needed to see updated state (AC-07)

---

## Dependency Graph

```
TASK-005-05-02 (StateMap dictionary)
       |
       v
TASK-005-05-03 (compute badge properties) --> TASK-005-05-06 (real-time propagation)

TASK-005-05-01 (badge XAML template)
       |
       +--------> TASK-005-05-04 (apply to list)
       +--------> TASK-005-05-05 (apply to detail header)
```
