# TASKS: US-008-07 — Keyboard Shortcuts

> **Parent US**: [US-008-07](US-008-07-keyboard-shortcuts.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 5 | **Effort**: 4S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-008-01 completed (NavigateCommand and INavigationService must exist)

## Acceptance Criteria
- [ ] AC-01: `Ctrl+1` through `Ctrl+7` navigate to the seven sidebar pages in order: Dashboard, AutoCAD, Odoo, BOQ Manager, Purchase Requisition, AI Assistant, Settings.
- [ ] AC-02: `F5` triggers the Refresh command (equivalent to clicking the Refresh button in the header).
- [ ] AC-03: `Alt+Left` navigates back to the previous page when `CanGoBack` is true in the navigation journal.
- [ ] AC-04: Keyboard shortcuts are registered as `InputBinding` or `KeyBinding` entries in MainWindow XAML.
- [ ] AC-05: Keyboard shortcuts work regardless of which page is currently displayed.
- [ ] AC-06: When `Alt+Left` is pressed and `CanGoBack` is false, the shortcut is silently ignored (no error or feedback).

---

## TASK-008-07-01: Define InputBindings for Ctrl+1 through Ctrl+7 navigation shortcuts

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `<Window.InputBindings>` section to `MainWindow.xaml`
- Define seven `KeyBinding` elements, each with a `KeyGesture` and bound to `NavigateCommand` with `CommandParameter`:
  - `<KeyBinding Key="D1" Modifiers="Ctrl" Command="{Binding NavigateCommand}" CommandParameter="Dashboard"/>`
  - `<KeyBinding Key="D2" Modifiers="Ctrl" Command="{Binding NavigateCommand}" CommandParameter="AutoCAD"/>`
  - `<KeyBinding Key="D3" Modifiers="Ctrl" Command="{Binding NavigateCommand}" CommandParameter="Odoo"/>`
  - `<KeyBinding Key="D4" Modifiers="Ctrl" Command="{Binding NavigateCommand}" CommandParameter="BOQ"/>`
  - `<KeyBinding Key="D5" Modifiers="Ctrl" Command="{Binding NavigateCommand}" CommandParameter="PR"/>`
  - `<KeyBinding Key="D6" Modifiers="Ctrl" Command="{Binding NavigateCommand}" CommandParameter="MCP"/>`
  - `<KeyBinding Key="D7" Modifiers="Ctrl" Command="{Binding NavigateCommand}" CommandParameter="Settings"/>`
- Use `Key="D1"` through `Key="D7"` (WPF key names for digit keys 1-7)
- The `NavigateCommand` must accept a string parameter (ensured by US-008-01 TASK-008-01-05)

### How to verify
- [ ] Ctrl+1 navigates to Dashboard, Ctrl+2 to AutoCAD, etc. through Ctrl+7 for Settings (AC-01)
- [ ] Shortcuts are defined as InputBinding/KeyBinding in XAML (AC-04)
- [ ] Shortcuts work regardless of current page (AC-05)

---

## TASK-008-07-02: Define InputBinding for F5 Refresh shortcut

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `KeyBinding` for F5 in the `<Window.InputBindings>` section:
  - `<KeyBinding Key="F5" Command="{Binding RefreshCommand}"/>`
- `RefreshCommand` is already defined as `[RelayCommand]` on `MainViewModel` (`RefreshAsync` method)
- No modifier key is needed for F5
- Ensure F5 does not conflict with WPF Frame's built-in refresh behavior (NavigationUIVisibility="Hidden" prevents this)

### How to verify
- [ ] F5 triggers the Refresh command (AC-02)
- [ ] Shortcut is defined as InputBinding in XAML (AC-04)
- [ ] Works on all pages (AC-05)

---

## TASK-008-07-03: Define InputBinding for Alt+Left GoBack shortcut

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | TASK-008-07-04 |
| Blocks | None |

### What to do
- Add a `KeyBinding` for Alt+Left in the `<Window.InputBindings>` section:
  - `<KeyBinding Key="Left" Modifiers="Alt" Command="{Binding GoBackCommand}"/>`
- The `GoBackCommand` must be defined on `MainViewModel` (see TASK-008-07-04)
- When `CanGoBack` is false, the command's `CanExecute` returns false, so the shortcut is silently ignored per AC-06

### How to verify
- [ ] Alt+Left navigates back when CanGoBack is true (AC-03)
- [ ] Alt+Left is silently ignored when CanGoBack is false (AC-06)
- [ ] Shortcut is defined as InputBinding in XAML (AC-04)

---

## TASK-008-07-04: Add GoBackCommand to MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-008-07-03 |

### What to do
- Add `INavigationService` dependency to `MainViewModel` constructor (if not already added by US-008-01 tasks)
- Add `[RelayCommand(CanExecute = nameof(CanGoBack))]` method `GoBack()` that calls `_navigationService.GoBack()`
- Add `bool CanGoBack => _navigationService.CanGoBack;` property for command's CanExecute evaluation
- After `GoBack()` executes, update `CurrentPageTitle` and `ActiveNavButton` to reflect the new current page (may require querying the Frame's current content type)
- The `CanGoBack` guard per VR-008-002 prevents `InvalidOperationException` when the Frame journal is empty

### How to verify
- [ ] GoBackCommand delegates to INavigationService.GoBack() (AC-03)
- [ ] CanExecute returns false when CanGoBack is false, silently disabling the shortcut (AC-06)
- [ ] No InvalidOperationException when journal is empty (AC-06)

---

## TASK-008-07-05: Ensure NavigateCommand accepts string parameter for keyboard shortcut routing

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify that the `Navigate` relay command method accepts a `string pageName` parameter
- The `[RelayCommand]` attribute on `Navigate(string pageName)` generates `NavigateCommand` as `IRelayCommand<string>`
- Confirm that `CommandParameter` values from XAML InputBindings ("Dashboard", "AutoCAD", "Odoo", "BOQ", "PR", "MCP", "Settings") match the page name switch expression in the Navigate method
- Ensure the same page names work for both sidebar button clicks and keyboard shortcuts

### How to verify
- [ ] NavigateCommand accepts string parameter from XAML CommandParameter (AC-01)
- [ ] Same page names work for both sidebar clicks and keyboard shortcuts (AC-05)

---

## Dependency Graph

```
TASK-008-07-01 (Ctrl+1..7 InputBindings) [independent]

TASK-008-07-02 (F5 InputBinding) [independent]

TASK-008-07-04 (GoBackCommand)
    └── TASK-008-07-03 (Alt+Left InputBinding)

TASK-008-07-05 (NavigateCommand parameter) [independent]
```
