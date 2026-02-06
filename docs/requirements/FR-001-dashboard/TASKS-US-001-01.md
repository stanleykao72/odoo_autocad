# TASKS: US-001-01 — View Connection Status

> **Parent US**: [US-001-01](US-001-01-view-connection-status.md)
> **Parent FR**: [FR-001](FR-001-dashboard.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 3S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure and CommunityToolkit.Mvvm must be in place)
- [ ] IAutoCADService and IOdooService interfaces being defined (already exist in `OdooAutoCAD.Core`)

## Acceptance Criteria
- [ ] AC-01: Dashboard displays an AutoCAD connection status card with icon, label, and connected/disconnected state
- [ ] AC-02: Dashboard displays an Odoo connection status card with icon, label, and connected/disconnected state
- [ ] AC-03: Connected state displays with a green indicator; disconnected state displays with a red/gray indicator
- [ ] AC-04: Connection status cards update in real-time when the connection state changes (no manual refresh required)
- [ ] AC-05: AutoCAD status card shows the current document name when connected
- [ ] AC-06: Odoo status card shows the server URL when connected
- [ ] AC-07: When AutoCAD connection fails, the status card displays "Unable to connect to AutoCAD. Please ensure AutoCAD is running."
- [ ] AC-08: When Odoo connection fails, the status card displays "Unable to connect to Odoo server. Check connection settings."

---

## TASK-001-01-01: Create DashboardPage XAML layout with connection status cards (AutoCAD, Odoo)

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-001-01-03 |

### What to do
- Create a new WPF `Page` in the `OdooAutoCAD.App.Views.Pages` namespace with `x:Class="OdooAutoCAD.App.Views.Pages.DashboardPage"`
- Add the corresponding code-behind `DashboardPage.xaml.cs` that sets `DataContext` to `DashboardViewModel` resolved from `App.Services`
- Define two `Border` elements with `CornerRadius="8"` arranged horizontally in a `UniformGrid` or `StackPanel Orientation="Horizontal"` for the AutoCAD and Odoo status cards (leave a third slot for MCP, to be filled in US-001-04)
- Each card should contain: an icon (TextBlock with emoji or Path), a title TextBlock ("AutoCAD" / "Odoo"), an `Ellipse` (Width=12, Height=12) for the status indicator with `Fill` bound to a Brush property, a status label TextBlock bound to the status text, and a detail TextBlock bound to the document name or server URL
- Add an error message `TextBlock` inside each card, bound to an error message property, with `Visibility` collapsed when empty
- Reference `StaticResource` styles from `MainWindow.xaml` resources (BackgroundBrush, SurfaceBrush, BorderBrush, TextPrimaryBrush, TextSecondaryBrush)
- Follow the wireframe layout from FR-001: cards in a horizontal row at the top of the dashboard page

### How to verify
- [ ] AutoCAD card renders with icon, "AutoCAD" label, status indicator ellipse, and status text (AC-01)
- [ ] Odoo card renders with icon, "Odoo" label, status indicator ellipse, and status text (AC-02)
- [ ] Detail TextBlocks for document name and server URL are present and bound (AC-05, AC-06)

---

## TASK-001-01-02: Add DashboardViewModel properties for connection status

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-001-01-04, TASK-001-01-05, TASK-001-01-06 |

### What to do
- Create a new file `ViewModels/DashboardViewModel.cs` in namespace `OdooAutoCAD.App.ViewModels`
- Define `public partial class DashboardViewModel : ObservableObject` using `CommunityToolkit.Mvvm.ComponentModel`
- Add `[ObservableProperty]` fields: `bool _isAutoCADConnected`, `string _autoCADStatus` (default "Disconnected"), `string _autoCADDocumentName` (default empty), `string _autoCADErrorMessage` (default empty)
- Add `[ObservableProperty]` fields: `bool _isOdooConnected`, `string _odooStatus` (default "Disconnected"), `string _odooServerUrl` (default empty), `string _odooErrorMessage` (default empty)
- Add `[ObservableProperty]` fields: `Brush _autoCADStatusColor` (default `Brushes.Gray`), `Brush _odooStatusColor` (default `Brushes.Gray`) using `System.Windows.Media`
- Inject `IAutoCADService` and `IOdooService` via constructor (accept as nullable since they may not yet be registered in DI)
- Store injected services as `private readonly` fields for use by status update and command methods

### How to verify
- [ ] ViewModel compiles with all six connection status properties generating proper change notifications (AC-01, AC-02)
- [ ] Properties for document name and server URL are accessible via binding (AC-05, AC-06)

---

## TASK-001-01-03: Implement BoolToColorConverter for green/red-gray status indicator binding

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | S |
| Depends On | TASK-001-01-01 |
| Blocks | None |

### What to do
- Create a new `IValueConverter` class `BoolToColorConverter` in namespace `OdooAutoCAD.App.Converters` (file: `Converters/BoolToColorConverter.cs`)
- Implement `Convert`: return `Brushes.Green` (or `new SolidColorBrush(Color.FromRgb(34, 197, 94))`) when `true`, return `Brushes.Gray` (or `new SolidColorBrush(Color.FromRgb(156, 163, 175))`) when `false`
- Implement `ConvertBack`: throw `NotSupportedException` (one-way binding only)
- Register the converter in `DashboardPage.xaml` resources: `<local:BoolToColorConverter x:Key="BoolToColorConverter"/>`
- Update the Ellipse `Fill` bindings in each status card to use `{Binding IsAutoCADConnected, Converter={StaticResource BoolToColorConverter}}` and `{Binding IsOdooConnected, Converter={StaticResource BoolToColorConverter}}`
- Alternatively, the ViewModel already exposes Brush properties (`AutoCADStatusColor`, `OdooStatusColor`), so you may bind directly to those and skip the converter if preferred -- choose one consistent approach

### How to verify
- [ ] Connected state shows green indicator; disconnected state shows red/gray indicator (AC-03)

---

## TASK-001-01-04: Subscribe DashboardViewModel to IAutoCADService and IOdooService status change events for real-time updates

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | TASK-001-01-02 |
| Blocks | None |

### What to do
- In the `DashboardViewModel` constructor, if `IAutoCADService` is not null, subscribe to its status change events (or set up a `DispatcherTimer` with 1-second interval to poll `IsConnected` and `GetStatusAsync()`)
- When AutoCAD status changes: update `IsAutoCADConnected`, `AutoCADStatus` ("Connected"/"Disconnected"), `AutoCADDocumentName` (from `AutoCADStatus.CurrentDocument`), and `AutoCADStatusColor` (green/gray)
- Similarly for `IOdooService`: poll or subscribe, update `IsOdooConnected`, `OdooStatus`, `OdooServerUrl` (from `OdooStatus.ServerUrl`), and `OdooStatusColor`
- Use `Application.Current.Dispatcher.Invoke()` or `Dispatcher.BeginInvoke()` to marshal property updates to the UI thread if polling on a background thread
- Implement `IDisposable` on `DashboardViewModel` to unsubscribe from events and stop timers, preventing memory leaks
- Follow the existing pattern in `MainViewModel.cs` which uses a `DispatcherTimer` with `OnStatusTimerTick`

### How to verify
- [ ] When AutoCAD connects/disconnects, the status card updates automatically without user interaction (AC-04)
- [ ] When Odoo connects/disconnects, the status card updates automatically without user interaction (AC-04)
- [ ] Document name appears in AutoCAD card when connected (AC-05)
- [ ] Server URL appears in Odoo card when connected (AC-06)

---

## TASK-001-01-05: Register DashboardViewModel as singleton in DI container to maintain state across navigation

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-01-02 |
| Blocks | None |

### What to do
- In `App.xaml.cs` method `ConfigureServices()`, register `DashboardViewModel` as a **singleton**: `services.AddSingleton<DashboardViewModel>()`
- Inject `IAutoCADService?` and `IOdooService?` as optional parameters (using `sp.GetService<>()`) since these may not yet be registered
- Also inject `ILogger<DashboardViewModel>?` for logging support
- Update `DashboardPage.xaml.cs` code-behind to resolve the ViewModel from `App.Services.GetRequiredService<DashboardViewModel>()` instead of creating a new instance
- This ensures the dashboard state (connection status, project info) persists when the user navigates away and back

### How to verify
- [ ] Connection status is preserved when navigating away from Dashboard and back (AC-04 -- state maintained)
- [ ] ViewModel is the same instance across navigation cycles

---

## TASK-001-01-06: Implement error message display in status cards when connection attempts fail

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-01-02 |
| Blocks | None |

### What to do
- Add a private method `SetAutoCADError(string message)` that sets `AutoCADErrorMessage` and updates `AutoCADStatus` to "Error"
- Add a private method `SetOdooError(string message)` that sets `OdooErrorMessage` and updates `OdooStatus` to "Error"
- When the status polling detects a connection failure via `AutoCADStatus.ErrorMessage`, call `SetAutoCADError("Unable to connect to AutoCAD. Please ensure AutoCAD is running.")`
- When the status polling detects a connection failure via `OdooStatus.ErrorMessage`, call `SetOdooError("Unable to connect to Odoo server. Check connection settings.")`
- In the XAML, bind the error TextBlock `Visibility` to `AutoCADErrorMessage` using a `StringToVisibilityConverter` (visible when non-empty) or use a `DataTrigger` on the ViewModel error property
- Clear the error message when a subsequent connection attempt succeeds

### How to verify
- [ ] When AutoCAD connection fails, the card displays "Unable to connect to AutoCAD. Please ensure AutoCAD is running." (AC-07)
- [ ] When Odoo connection fails, the card displays "Unable to connect to Odoo server. Check connection settings." (AC-08)

---

## Dependency Graph
```
TASK-001-01-01 (XAML Layout)
       │
       └──▶ TASK-001-01-03 (BoolToColorConverter)

TASK-001-01-02 (ViewModel Properties)
       │
       ├──▶ TASK-001-01-04 (Event Subscriptions)
       ├──▶ TASK-001-01-05 (DI Registration)
       └──▶ TASK-001-01-06 (Error Messages)

No cross-dependencies between 01 and 02; they can be developed in parallel.
Tasks 03-06 all depend on their respective parent (01 or 02) but are independent of each other.
```
