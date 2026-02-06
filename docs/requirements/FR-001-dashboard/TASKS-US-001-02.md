# TASKS: US-001-02 — Quick Connect Buttons

> **Parent US**: [US-001-02](US-001-02-quick-connect-buttons.md)
> **Parent FR**: [FR-001](FR-001-dashboard.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 3S + 5M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-001-01 (connection status cards must exist before adding buttons to them)
- [ ] IAutoCADService.ConnectAsync() and IOdooService.ConnectAsync() being implemented
- [ ] INavigationService for workflow shortcut navigation (already exists in `OdooAutoCAD.App.Services`)

## Acceptance Criteria
- [ ] AC-01: Dashboard provides a "Connect to AutoCAD" quick-action button within the AutoCAD status card
- [ ] AC-02: Dashboard provides a "Connect to Odoo" quick-action button within the Odoo status card
- [ ] AC-03: Quick-connect buttons are disabled and show "Connected" state when the respective service is already connected
- [ ] AC-04: Clicking "Connect to AutoCAD" initiates an async connection attempt to AutoCAD
- [ ] AC-05: Clicking "Connect to Odoo" initiates an async connection attempt to Odoo
- [ ] AC-06: Quick-connect buttons check current connection state before attempting connection (VR-001-001)
- [ ] AC-07: Dashboard provides shortcut buttons for common workflows: "Get Parameters", "Push to BOQ", and "Transfer BOQ to PR"
- [ ] AC-08: "Get Parameters" shortcut is only enabled when both AutoCAD AND Odoo are connected (VR-001-003)
- [ ] AC-09: "Push to BOQ" shortcut is only enabled when Odoo is connected and project context is available (VR-001-004)
- [ ] AC-10: Shortcut buttons navigate to the corresponding application page when clicked

---

## TASK-001-02-01: Add connect buttons inside status cards in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-001-02-04 |

### What to do
- In the AutoCAD status card `Border` (created in TASK-001-01-01), add a `Button` at the bottom of the card's `StackPanel`
- Set `Content` to `{Binding ConnectAutoCADButtonText}` (which will show "Connect" or "Connected")
- Bind `Command` to `{Binding ConnectAutoCADCommand}`
- Apply a consistent button style matching the application theme (reference `PrimaryButton` style from `MainWindow.xaml`)
- Repeat for the Odoo status card with `Content="{Binding ConnectOdooButtonText}"` and `Command="{Binding ConnectOdooCommand}"`
- Set `IsEnabled` on each button via the command's built-in `CanExecute` (no separate binding needed when using `RelayCommand`)

### How to verify
- [ ] "Connect to AutoCAD" button is visible inside the AutoCAD status card (AC-01)
- [ ] "Connect to Odoo" button is visible inside the Odoo status card (AC-02)

---

## TASK-001-02-02: Implement ConnectAutoCADCommand with async execution

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-001-02-04, TASK-001-02-08 |

### What to do
- Add an `[RelayCommand(CanExecute = nameof(CanConnectAutoCAD))]` attribute on a new `private async Task ConnectAutoCADAsync()` method
- In `ConnectAutoCADAsync()`: set `AutoCADStatus = "Connecting..."`, clear `AutoCADErrorMessage`, then call `await _autoCADService.ConnectAsync()`
- On success: update `IsAutoCADConnected = true`, `AutoCADStatus = "Connected"`, `AutoCADStatusColor = Brushes.Green`, and fetch `AutoCADDocumentName` from `GetStatusAsync()`
- On failure (exception caught): set `AutoCADErrorMessage = "Unable to connect to AutoCAD. Please ensure AutoCAD is running."`, `AutoCADStatus = "Disconnected"`, `AutoCADStatusColor = Brushes.Gray`
- Use `AsyncRelayCommand` from CommunityToolkit.Mvvm to avoid blocking the UI thread
- Add a property `string ConnectAutoCADButtonText` that returns "Connected" when `IsAutoCADConnected` is true, "Connect" otherwise; call `OnPropertyChanged(nameof(ConnectAutoCADButtonText))` when `IsAutoCADConnected` changes
- If `_autoCADService` is null, show an appropriate message ("AutoCAD service not available")

### How to verify
- [ ] Clicking the button initiates an async connection attempt (AC-04)
- [ ] Connection state check occurs before attempting connection (AC-06)

---

## TASK-001-02-03: Implement ConnectOdooCommand with async execution

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-001-02-04, TASK-001-02-08 |

### What to do
- Add an `[RelayCommand(CanExecute = nameof(CanConnectOdoo))]` attribute on a new `private async Task ConnectOdooAsync()` method
- In `ConnectOdooAsync()`: set `OdooStatus = "Connecting..."`, clear `OdooErrorMessage`, then call `await _odooService.ConnectAsync(serverUrl, database, username, password)` using stored/configured credentials
- On success: update `IsOdooConnected = true`, `OdooStatus = "Connected"`, `OdooStatusColor = Brushes.Green`, and fetch `OdooServerUrl` from `GetStatusAsync()`
- On failure: set `OdooErrorMessage = "Unable to connect to Odoo server. Check connection settings."`, `OdooStatus = "Disconnected"`, `OdooStatusColor = Brushes.Gray`
- Use `AsyncRelayCommand` from CommunityToolkit.Mvvm
- Add a property `string ConnectOdooButtonText` that returns "Connected" when `IsOdooConnected` is true, "Connect" otherwise
- Odoo credentials should come from the application configuration (injected `IConfiguration` or a settings service); for now, accept them as constructor parameters or use placeholder logic
- If `_odooService` is null, show an appropriate message ("Odoo service not available")

### How to verify
- [ ] Clicking the button initiates an async connection attempt to Odoo (AC-05)
- [ ] Connection state check occurs before attempting connection (AC-06)

---

## TASK-001-02-04: Add CanExecute logic to connect commands

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-02-02, TASK-001-02-03 |
| Blocks | TASK-001-02-08 |

### What to do
- Implement `private bool CanConnectAutoCAD()` that returns `true` only when `!IsAutoCADConnected` and `_autoCADService != null`
- Implement `private bool CanConnectOdoo()` that returns `true` only when `!IsOdooConnected` and `_odooService != null`
- When `IsAutoCADConnected` changes, update the button text to "Connected" and the command becomes non-executable (button grayed out)
- When `IsOdooConnected` changes, same pattern
- Use `partial void OnIsAutoCADConnectedChanged(bool value)` (CommunityToolkit.Mvvm source-generated partial method) to call `ConnectAutoCADCommand.NotifyCanExecuteChanged()` and update `ConnectAutoCADButtonText`
- Use `partial void OnIsOdooConnectedChanged(bool value)` similarly for Odoo

### How to verify
- [ ] Buttons are disabled and show "Connected" when the respective service is already connected (AC-03)
- [ ] Buttons check current connection state before allowing execution (AC-06)

---

## TASK-001-02-05: Create Quick Actions section in XAML with workflow shortcut buttons

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | M |
| Depends On | TASK-001-02-01 |
| Blocks | None |

### What to do
- Below the connection status cards section, add a "Quick Actions" header `TextBlock` with `FontSize="18"` and `FontWeight="SemiBold"`
- Create a horizontal `WrapPanel` or `UniformGrid Columns="4"` containing four `Button` elements
- Button 1: Content "Get Parameters" (icon: clipboard emoji or Path icon), Command bound to `{Binding NavigateToParametersCommand}`
- Button 2: Content "Push to BOQ" (icon: chart emoji), Command bound to `{Binding NavigateToBOQCommand}`
- Button 3: Content "Transfer BOQ to PR" (icon: refresh/transfer emoji), Command bound to `{Binding NavigateToPRCommand}`
- Button 4 (optional stretch): Content "AI Assistant", Command bound to `{Binding NavigateToMCPCommand}`
- Style each button as a card-like element: `Border` with `CornerRadius="8"`, `Padding="16"`, minimum size 120x80, with icon above text in a `StackPanel`
- Each button's `IsEnabled` is automatically managed by the command's `CanExecute`

### How to verify
- [ ] Dashboard displays "Get Parameters", "Push to BOQ", and "Transfer BOQ to PR" shortcut buttons (AC-07)
- [ ] Buttons are visually styled as card-like shortcuts matching the wireframe layout

---

## TASK-001-02-06: Implement navigation commands using INavigationService

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-001-02-07 |

### What to do
- Inject `INavigationService` into `DashboardViewModel` constructor and store as `private readonly INavigationService _navigationService`
- Add `[RelayCommand(CanExecute = nameof(CanNavigateToParameters))]` on `private void NavigateToParameters()` which calls `_navigationService.NavigateTo("AutoCADParam")` (or the appropriate page name matching the `NavigationService.NavigateTo()` convention)
- Add `[RelayCommand(CanExecute = nameof(CanNavigateToBOQ))]` on `private void NavigateToBOQ()` which calls `_navigationService.NavigateTo("BOQ")`
- Add `[RelayCommand(CanExecute = nameof(CanNavigateToPR))]` on `private void NavigateToPR()` which calls `_navigationService.NavigateTo("PurchaseRequisition")`
- Optionally add `[RelayCommand]` on `private void NavigateToMCP()` which calls `_navigationService.NavigateTo("AIAssistant")`
- The page name strings should match the naming convention in `NavigationService.cs` which resolves `OdooAutoCAD.App.Views.Pages.{pageName}Page`

### How to verify
- [ ] Clicking "Get Parameters" navigates to the parameters page (AC-10)
- [ ] Clicking "Push to BOQ" navigates to the BOQ page (AC-10)
- [ ] Clicking "Transfer BOQ to PR" navigates to the PR page (AC-10)

---

## TASK-001-02-07: Add CanExecute validation for workflow shortcuts

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-02-06 |
| Blocks | TASK-001-02-08 |

### What to do
- Implement `private bool CanNavigateToParameters()` that returns `IsAutoCADConnected && IsOdooConnected` (VR-001-003: requires both AutoCAD AND Odoo connected)
- Implement `private bool CanNavigateToBOQ()` that returns `IsOdooConnected && HasProjectContext` where `HasProjectContext` is a computed property checking `!string.IsNullOrEmpty(CurrentPRNo)` (VR-001-004: requires Odoo connected and project context available)
- Implement `private bool CanNavigateToPR()` that returns `IsOdooConnected` (minimum requirement for PR operations)
- Add a helper property `bool HasProjectContext => !string.IsNullOrEmpty(CurrentPRNo)` (using `CurrentPRNo` from US-001-03, or define a placeholder `_currentPRNo` field for now)
- These CanExecute methods enforce validation rules VR-001-001 through VR-001-004 from the FR document

### How to verify
- [ ] "Get Parameters" is only enabled when both AutoCAD AND Odoo are connected (AC-08)
- [ ] "Push to BOQ" is only enabled when Odoo is connected and project context exists (AC-09)

---

## TASK-001-02-08: Re-evaluate command CanExecute states when connection status changes

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-02-04, TASK-001-02-07 |
| Blocks | None |

### What to do
- In `partial void OnIsAutoCADConnectedChanged(bool value)` (or create if not yet present), add calls to:
  - `ConnectAutoCADCommand.NotifyCanExecuteChanged()`
  - `NavigateToParametersCommand.NotifyCanExecuteChanged()`
  - Update `ConnectAutoCADButtonText` property
- In `partial void OnIsOdooConnectedChanged(bool value)`, add calls to:
  - `ConnectOdooCommand.NotifyCanExecuteChanged()`
  - `NavigateToParametersCommand.NotifyCanExecuteChanged()`
  - `NavigateToBOQCommand.NotifyCanExecuteChanged()`
  - `NavigateToPRCommand.NotifyCanExecuteChanged()`
  - Update `ConnectOdooButtonText` property
- If `CurrentPRNo` changes (project context availability), also call `NavigateToBOQCommand.NotifyCanExecuteChanged()`
- This ensures all button enabled/disabled states react immediately to connection changes without requiring manual refresh

### How to verify
- [ ] When AutoCAD connects, the "Connect to AutoCAD" button becomes disabled and shows "Connected"; "Get Parameters" becomes enabled (if Odoo also connected) (AC-03, AC-08)
- [ ] When Odoo disconnects, "Get Parameters" and "Push to BOQ" shortcuts become disabled immediately (AC-08, AC-09)

---

## Dependency Graph
```
TASK-001-02-01 (Connect Buttons XAML)
       │
       └──▶ TASK-001-02-05 (Quick Actions XAML)

TASK-001-02-02 (ConnectAutoCAD Command) ──┐
                                          ├──▶ TASK-001-02-04 (CanExecute Logic) ──┐
TASK-001-02-03 (ConnectOdoo Command) ─────┘                                        │
                                                                                   ├──▶ TASK-001-02-08 (NotifyCanExecuteChanged)
TASK-001-02-06 (Navigation Commands) ──▶ TASK-001-02-07 (Workflow Validation) ─────┘
```
