# TASKS: US-008-08 — Page Title Header

> **Parent US**: [US-008-08](US-008-08-page-title-header.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 5 | **Effort**: 4S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-008-01 completed (navigation must update CurrentPageTitle on page change)

## Acceptance Criteria
- [ ] AC-01: A header bar is displayed at the top of the content area (above the main Frame, to the right of the sidebar).
- [ ] AC-02: The header bar displays `CurrentPageTitle` as a 20pt SemiBold TextBlock bound to the MainViewModel.
- [ ] AC-03: The header bar includes a "Refresh" button on the right side bound to `RefreshCommand`.
- [ ] AC-04: The header bar optionally includes a "Help" button on the right side.
- [ ] AC-05: The header bar uses `SurfaceBrush` background with a bottom border separator.
- [ ] AC-06: `CurrentPageTitle` updates automatically when navigation occurs, reflecting the display name of the active page.
- [ ] AC-07: Page title mapping follows the defined convention: Dashboard, AutoCAD Integration, Odoo Connection, BOQ Manager, Purchase Requisition, AI Assistant (MCP), Settings.

---

## TASK-008-08-01: Define header bar XAML with page title TextBlock and action buttons

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-008-08-05 |

### What to do
- In the content area Grid Row 0, define a `Border` with `Background="{StaticResource SurfaceBrush}"`, `Padding="20,15"`, `BorderBrush="{StaticResource BorderBrush}"`, and `BorderThickness="0,0,0,1"` (bottom border)
- Add an inner `Grid` with two columns: `*` (title fills) and Auto (action buttons)
- Add `TextBlock` in column 0 with `Text="{Binding CurrentPageTitle}"`, `FontSize="20"`, `FontWeight="SemiBold"`, `VerticalAlignment="Center"`
- Add a `StackPanel` in column 1 with `Orientation="Horizontal"` containing:
  - "Refresh" `Button` with `Style="{StaticResource PrimaryButton}"`, `Command="{Binding RefreshCommand}"`, `Margin="0,0,10,0"`
  - "Help" `Button` with `Style="{StaticResource PrimaryButton}"` (optional, no command binding required initially)
- Verify the skeleton already defines this layout; confirm correctness

### How to verify
- [ ] Header bar is at the top of the content area, right of sidebar (AC-01)
- [ ] Page title is 20pt SemiBold bound to CurrentPageTitle (AC-02)
- [ ] Refresh button is present and bound to RefreshCommand (AC-03)
- [ ] Help button is present on the right side (AC-04)

---

## TASK-008-08-02: Add CurrentPageTitle observable property to MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-008-08-04 |

### What to do
- Verify `[ObservableProperty] private string _currentPageTitle = "Dashboard";` exists in `MainViewModel`
- Default value "Dashboard" ensures the correct title on startup
- The property must raise `PropertyChanged` for the binding in the header TextBlock to update
- Already present in the skeleton; confirm it is correct

### How to verify
- [ ] CurrentPageTitle property exists with default "Dashboard" (AC-02, AC-06)
- [ ] Property raises PropertyChanged notification (AC-06)

---

## TASK-008-08-03: Add RefreshCommand to MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify `[RelayCommand] private async Task RefreshAsync()` exists in `MainViewModel`
- The method should set `StatusMessage = "Refreshing..."`, perform refresh logic (e.g., `UpdateAllStatus()`), then set `StatusMessage = "Ready"`
- The generated `RefreshCommand` is of type `IAsyncRelayCommand`, suitable for binding from XAML
- Already present in the skeleton; confirm correct implementation

### How to verify
- [ ] RefreshCommand exists and can be bound from XAML (AC-03)
- [ ] Clicking Refresh updates status message feedback (AC-03)

---

## TASK-008-08-04: Define page title display name mapping

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | TASK-008-08-02 |
| Blocks | None |

### What to do
- In the `Navigate(string pageName)` method, use a switch expression to map page names to display titles:
  - "Dashboard" -> "Dashboard"
  - "AutoCAD" -> "AutoCAD Integration"
  - "Odoo" -> "Odoo Connection"
  - "BOQ" -> "BOQ Manager"
  - "PR" -> "Purchase Requisition"
  - "MCP" -> "AI Assistant (MCP)"
  - "Settings" -> "Settings"
  - Default/fallback -> use `pageName` as-is
- Set `CurrentPageTitle` to the mapped display name
- The mapping currently exists in `MainWindow.xaml.cs` code-behind; move it to `MainViewModel` for MVVM compliance
- Alternatively, define a static `Dictionary<string, string>` or use `NavigationService` to supply display titles

### How to verify
- [ ] Page title updates to the correct display name on navigation (AC-06, AC-07)
- [ ] All seven display names match the defined convention (AC-07)

---

## TASK-008-08-05: Apply SurfaceBrush background and bottom border styling to header bar

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | TASK-008-08-01 |
| Blocks | None |

### What to do
- Verify the header `Border` uses `Background="{StaticResource SurfaceBrush}"` (white, `#FFFFFF`)
- Verify `BorderBrush="{StaticResource BorderBrush}"` (`#E0E0E0`) is set
- Verify `BorderThickness="0,0,0,1"` for bottom-only border separator
- These should use `{StaticResource}` references, not hardcoded colors, per FR-008-034
- Already in the skeleton; confirm resource references are used correctly

### How to verify
- [ ] Header uses SurfaceBrush background (AC-05)
- [ ] Header has a bottom border separator using BorderBrush (AC-05)
- [ ] All colors reference ResourceDictionary resources, not hardcoded values (AC-05)

---

## Dependency Graph

```
TASK-008-08-01 (Header XAML)
    └── TASK-008-08-05 (SurfaceBrush styling)

TASK-008-08-02 (CurrentPageTitle property)
    └── TASK-008-08-04 (Page title mapping)

TASK-008-08-03 (RefreshCommand) [independent]
```
