# TASKS: US-008-02 — Active Page Indicator

> **Parent US**: [US-008-02](US-008-02-active-page-indicator.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 4 | **Effort**: 3S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-008-01 completed (sidebar navigation buttons and NavButton style must exist)

## Acceptance Criteria
- [ ] AC-01: The currently active navigation button has a visually distinct background color (e.g., `#E3F2FD`) differentiating it from inactive buttons.
- [ ] AC-02: The currently active navigation button displays a left border accent in the primary color (`PrimaryBrush`).
- [ ] AC-03: When navigating to a different page, the previously active button returns to its default (transparent) style.
- [ ] AC-04: The active button highlighting updates atomically with the Frame navigation to prevent desynchronization.
- [ ] AC-05: On application startup, the Dashboard button is highlighted as the active button.
- [ ] AC-06: Active state is driven by `MainViewModel.ActiveNavButton` property via data binding, not code-behind UI manipulation.

---

## TASK-008-02-01: Add ActiveNavButton observable property to MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-008-02-03 |

### What to do
- Add `[ObservableProperty] private string _activeNavButton = "BtnDashboard";` to `MainViewModel`
- The default value "BtnDashboard" ensures the Dashboard button is highlighted on startup
- The property name will be code-generated as `ActiveNavButton` by CommunityToolkit.Mvvm source generators
- Ensure the property type is `string` to match the `x:Name` values of sidebar buttons

### How to verify
- [ ] `ActiveNavButton` property exists on `MainViewModel` with default "BtnDashboard" (AC-05, AC-06)
- [ ] Property raises `PropertyChanged` notification when updated (AC-06)

---

## TASK-008-02-02: Implement NavButton style DataTriggers for active state highlighting

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` (Window.Resources NavButton style) |
| Estimate | M |
| Depends On | TASK-008-02-01 |
| Blocks | TASK-008-02-04 |

### What to do
- In the `NavButton` `ControlTemplate`, add `DataTrigger` entries that compare each button's `Tag` property to `{Binding ActiveNavButton}` on the DataContext
- For each navigation button value (BtnDashboard, BtnAutoCAD, BtnOdoo, BtnBOQ, BtnPR, BtnMCP, BtnSettings), add a `DataTrigger` with `Binding="{Binding ActiveNavButton}"` and `Value` matching the button's `Tag`
- Alternatively, use an `IMultiValueConverter` or a single `DataTrigger` with `RelativeSource` binding to compare `Tag` with `ActiveNavButton`
- Set active state: `Background="#E3F2FD"` and `border.BorderBrush="{StaticResource PrimaryBrush}"` on the named `border` element
- Ensure the active trigger has higher priority than the `IsMouseOver` trigger so both co-exist correctly
- Active state must revert to transparent when another button becomes active (handled by WPF DataTrigger removal)

### How to verify
- [ ] Active button shows `#E3F2FD` background (AC-01)
- [ ] Active button shows left border accent in primary color (AC-02)
- [ ] Previously active button returns to transparent/default style (AC-03)

---

## TASK-008-02-03: Update NavigateCommand handler to set ActiveNavButton alongside CurrentPageTitle

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | TASK-008-02-01 |
| Blocks | None |

### What to do
- In the `Navigate(string pageName)` relay command method, set `ActiveNavButton = $"Btn{pageName}"` atomically with the `CurrentPageTitle` update
- Ensure both properties are set before `INavigationService.NavigateTo()` is called, or immediately after, to maintain visual synchronization per VR-008-004
- The mapping is: pageName "Dashboard" -> ActiveNavButton "BtnDashboard", "AutoCAD" -> "BtnAutoCAD", etc.
- Both `CurrentPageTitle` and `ActiveNavButton` must update in the same method invocation to prevent desynchronization

### How to verify
- [ ] `ActiveNavButton` updates atomically with `CurrentPageTitle` on navigation (AC-04)
- [ ] Dashboard button is active on startup via default property value (AC-05)

---

## TASK-008-02-04: Assign Tag property to each sidebar button matching ActiveNavButton values

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | TASK-008-02-02 |
| Blocks | None |

### What to do
- On each sidebar `Button`, add `Tag="BtnDashboard"`, `Tag="BtnAutoCAD"`, etc., matching the `x:Name` value
- The `Tag` property is used by the `DataTrigger` to determine whether the button is the active one
- Ensure all seven buttons have their `Tag` set: BtnDashboard, BtnAutoCAD, BtnOdoo, BtnBOQ, BtnPR, BtnMCP, BtnSettings
- Alternatively, the DataTrigger can use `{Binding RelativeSource={RelativeSource Self}, Path=Name}` to read the Name directly, avoiding the need for separate Tag values

### How to verify
- [ ] Each button has a `Tag` (or `Name`) that matches the `ActiveNavButton` tracking values (AC-06)
- [ ] Active state visual is driven entirely by data binding, not code-behind (AC-06)

---

## Dependency Graph

```
TASK-008-02-01 (ActiveNavButton property)
    ├── TASK-008-02-02 (NavButton DataTriggers)
    │   └── TASK-008-02-04 (Tag assignments)
    └── TASK-008-02-03 (NavigateCommand update)
```
