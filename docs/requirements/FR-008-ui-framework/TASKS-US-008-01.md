# TASKS: US-008-01 — Sidebar Navigation

> **Parent US**: [US-008-01](US-008-01-sidebar-navigation.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 7 | **Effort**: 2S + 4M + 1L
> **Status**: Done

## Prerequisites
- [x] None (this is the foundational US for the entire UI framework)

## Acceptance Criteria
- [x] AC-01: A fixed-width (250px) sidebar panel is displayed on the left side of the main window at all times.
- [x] AC-02: The sidebar displays a title area at the top with "Odoo AutoCAD" and "Integration System v6.0".
- [x] AC-03: The sidebar contains navigation buttons for all seven areas: Dashboard, AutoCAD, Odoo, BOQ Manager, Purchase Requisition, AI Assistant, Settings.
- [x] AC-04: Each navigation button displays an icon and a text label, left-aligned.
- [x] AC-05: Navigation buttons use the `NavButton` style with transparent background, left-aligned content, and a left-border accent on hover.
- [x] AC-06: Clicking a sidebar button loads the corresponding page into the main `Frame` via `INavigationService.NavigateTo()`.
- [x] AC-07: The application defaults to the Dashboard page on startup.
- [x] AC-08: The `Frame` element uses `NavigationUIVisibility="Hidden"` to suppress default WPF navigation chrome.
- [x] AC-09: `INavigationService.NavigateTo(pageName)` resolves pages by reflection from namespace `OdooAutoCAD.App.Views.Pages.{pageName}Page`.
- [x] AC-10: `INavigationService.GoBack()` navigates to the previous page when `CanGoBack` is true.
- [x] AC-11: If a resolved page type does not exist, navigation is silently skipped and a warning is logged.

---

## TASK-008-01-01: Define sidebar XAML layout with title area, navigation buttons, and connection status section

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-008-01-02, TASK-008-01-06 |

### What to do
- Define a `Border` in Grid column 0 with `SurfaceBrush` background and right border separator, fixed 250px width via `ColumnDefinition`
- Add a three-row inner `Grid`: title area (Auto), navigation `StackPanel` (star), connection status placeholder (Auto)
- Add title area `StackPanel` with `TextBlock` "Odoo AutoCAD" (24pt Bold, `PrimaryBrush`) and "Integration System v6.0" (12pt, `TextSecondaryBrush`), with bottom border separator
- Add seven `Button` elements (`BtnDashboard`, `BtnAutoCAD`, `BtnOdoo`, `BtnBOQ`, `BtnPR`, `BtnMCP`, `BtnSettings`) each with emoji icon and text label, using `NavButton` style and `NavButton_Click` event handler
- Add connection status `Border` at Row 2 with top border separator (placeholder for US-008-03)
- Set `x:Name` on each button matching the page mapping convention (e.g., `BtnDashboard`)
- Ensure the `Frame` element (`MainFrame`) in content area has `NavigationUIVisibility="Hidden"` and 20px margin

### How to verify
- [x] Sidebar renders at 250px width on the left side at all times (AC-01)
- [x] Title area shows "Odoo AutoCAD" and "Integration System v6.0" (AC-02)
- [x] All seven navigation buttons are present with icon and text, left-aligned (AC-03, AC-04)
- [x] Frame element has `NavigationUIVisibility="Hidden"` (AC-08)

---

## TASK-008-01-02: Implement NavButton style with transparent background, left-align, and hover left-border accent

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` (Window.Resources) or `App.xaml` |
| Estimate | M |
| Depends On | TASK-008-01-01 |
| Blocks | None |

### What to do
- Define a `Style` with `x:Key="NavButton"` targeting `Button` in the `Window.Resources` or `App.xaml ResourceDictionary`
- Set `Background` to `Transparent`, `Foreground` to `TextPrimaryBrush`, `Padding` to `20,12`, `HorizontalContentAlignment` to `Left`, `BorderThickness` to `0`, `Cursor` to `Hand`
- Create a `ControlTemplate` with a `Border` named `border` using `BorderThickness="3,0,0,0"` and `BorderBrush="Transparent"` for the left accent
- Add `ControlTemplate.Triggers`: on `IsMouseOver=True`, set `Background` to `#F0F0F0` and `border.BorderBrush` to `PrimaryBrush`
- Ensure the `ContentPresenter` uses `HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"` for left-aligned content

### How to verify
- [x] Navigation buttons have transparent background by default (AC-05)
- [x] On hover, buttons show `#F0F0F0` background and left border accent in primary color (AC-05)
- [x] Content is left-aligned within each button (AC-04)

---

## TASK-008-01-03: Create INavigationService interface with NavigateTo, GoBack, CanGoBack, and Frame property

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Services/NavigationService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-008-01-04 |

### What to do
- Define `INavigationService` interface with `Frame? Frame { get; set; }` property for the WPF `Frame` reference
- Add `void NavigateTo(string pageName)` method for page navigation by name
- Add `void GoBack()` method for journal-based back navigation
- Add `bool CanGoBack { get; }` read-only property to query Frame journal state
- Place the interface in `OdooAutoCAD.App.Services` namespace
- Validate rule VR-008-003: document that `Frame` must be set before any navigation calls

### How to verify
- [x] Interface defines `NavigateTo`, `GoBack`, `CanGoBack`, and `Frame` members (AC-06, AC-09, AC-10)
- [x] Interface is in the correct namespace for DI registration (AC-09)

---

## TASK-008-01-04: Implement NavigationService with reflection-based page resolution

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Services/NavigationService.cs` |
| Estimate | M |
| Depends On | TASK-008-01-03 |
| Blocks | TASK-008-01-06 |

### What to do
- Implement `NavigationService : INavigationService` with `Frame` property
- In `NavigateTo(string pageName)`, guard with `if (Frame == null) return;` per VR-008-003
- Resolve page type via `Type.GetType($"OdooAutoCAD.App.Views.Pages.{pageName}Page")` using reflection
- If `pageType` is null, log a warning and silently skip navigation (per VR-008-001 and AC-11)
- If `pageType` is valid, call `Frame.Navigate(Activator.CreateInstance(pageType))`
- Implement `GoBack()` with `CanGoBack` guard: `if (Frame?.CanGoBack == true) Frame.GoBack();` per VR-008-002
- Implement `CanGoBack` as `Frame?.CanGoBack ?? false`

### How to verify
- [x] Pages are resolved by reflection from `OdooAutoCAD.App.Views.Pages.{pageName}Page` namespace (AC-09)
- [x] Non-existent page types are silently skipped with a warning log (AC-11)
- [x] `GoBack()` checks `CanGoBack` before navigating (AC-10)
- [x] Null `Frame` does not throw exceptions (AC-09)

---

## TASK-008-01-05: Add NavigateCommand and page title mapping to MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | M |
| Depends On | TASK-008-01-03 |
| Blocks | TASK-008-01-06 |

### What to do
- Add `INavigationService` dependency to `MainViewModel` constructor via DI injection
- Add `[RelayCommand]` method `Navigate(string pageName)` that calls `_navigationService.NavigateTo(pageName)`
- Update `CurrentPageTitle` in the `Navigate` method using a switch expression mapping: Dashboard->"Dashboard", AutoCAD->"AutoCAD Integration", Odoo->"Odoo Connection", BOQ->"BOQ Manager", PR->"Purchase Requisition", MCP->"AI Assistant (MCP)", Settings->"Settings"
- Ensure `NavigateCommand` is of type `IRelayCommand<string>` to accept page name parameter from XAML bindings and keyboard shortcuts
- Set default `CurrentPageTitle` to "Dashboard" on construction

### How to verify
- [x] `NavigateCommand` accepts string parameter and triggers `NavigateTo` on `INavigationService` (AC-06)
- [x] `CurrentPageTitle` updates to the correct display name on navigation (AC-06)
- [x] Default page title is "Dashboard" on startup (AC-07)

---

## TASK-008-01-06: Wire MainWindow constructor to set Frame on NavigationService and navigate to Dashboard

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml.cs` |
| Estimate | S |
| Depends On | TASK-008-01-04, TASK-008-01-05 |
| Blocks | None |

### What to do
- In `MainWindow()` constructor, resolve `INavigationService` from `App.Services`
- Assign `MainFrame` (the XAML `Frame` element) to `_navigationService.Frame` per VR-008-003
- Resolve `MainViewModel` from DI and set as `DataContext`
- Call `_navigationService.NavigateTo("Dashboard")` for default startup navigation per AC-07
- Remove the inline `NavigateTo` method and delegate to `INavigationService` via ViewModel's `NavigateCommand`
- Update `NavButton_Click` to extract page name from button `Name` and invoke `_viewModel.NavigateCommand.Execute(pageName)`

### How to verify
- [x] `MainFrame` is assigned to `INavigationService.Frame` before any navigation (AC-08, AC-09)
- [x] Application defaults to Dashboard page on startup (AC-07)
- [x] Sidebar button clicks trigger `INavigationService.NavigateTo()` (AC-06)

---

## TASK-008-01-07: Register INavigationService as singleton in DI container

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | TASK-008-01-03 |
| Blocks | None |

### What to do
- In `ConfigureServices`, register `INavigationService` as singleton: `services.AddSingleton<INavigationService, NavigationService>()`
- Ensure registration is placed after ViewModels and in the "Application services" section per DI registration order
- Verify no duplicate registrations exist for `INavigationService`

### How to verify
- [x] `INavigationService` resolves correctly from `App.Services` (AC-06)
- [x] Singleton lifetime ensures the same Frame reference is shared across the application (AC-09)

---

## Dependency Graph

```
TASK-008-01-01 (XAML layout)
    └── TASK-008-01-02 (NavButton style)

TASK-008-01-03 (INavigationService interface)
    ├── TASK-008-01-04 (NavigationService impl)
    │   └── TASK-008-01-06 (MainWindow wiring)
    ├── TASK-008-01-05 (NavigateCommand in ViewModel)
    │   └── TASK-008-01-06 (MainWindow wiring)
    └── TASK-008-01-07 (DI registration)
```
