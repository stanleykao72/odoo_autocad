# TASKS: US-001-03 — Recent Activity

> **Parent US**: [US-001-03](US-001-03-recent-activity.md)
> **Parent FR**: [FR-001](FR-001-dashboard.md)
> **Priority**: P2
> **Tasks**: 8 | **Effort**: 3S + 5M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-001-01 (connection status must be available to determine when to show project info)
- [ ] IOdooService.GetLastSyncTime() being implemented (interface already exists in `OdooAutoCAD.Core.Odoo`)
- [ ] IAutoCADService providing current document and project context information

## Acceptance Criteria
- [ ] AC-01: Dashboard displays the last sync timestamp for Odoo data (e.g., "Last Sync: 2025-08-19 14:30")
- [ ] AC-02: Dashboard displays counts of recent operations including parameters extracted and BOQs pushed
- [ ] AC-03: Dashboard displays current project info (PR No, Project Name, Job Plan) when AutoCAD is connected and project context is available
- [ ] AC-04: When no project context is available, the project info panel displays "No project context available. Open a drawing with PR information."
- [ ] AC-05: Last sync timestamp updates automatically after each successful Odoo synchronization
- [ ] AC-06: Project info panel updates when a different AutoCAD drawing is opened or connection changes

---

## TASK-001-03-01: Create Current Project info panel in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Below the connection status cards section in `DashboardPage.xaml`, add a "Current Project" section with a header `TextBlock` ("Current Project", `FontSize="18"`, `FontWeight="SemiBold"`)
- Create a `Border` with `CornerRadius="8"`, `Background="{StaticResource SurfaceBrush}"`, `BorderBrush="{StaticResource BorderBrush}"`, `BorderThickness="1"`, `Padding="16"`
- Inside the Border, use a `Grid` with 2 rows and 2 columns (matching the FR wireframe layout):
  - Row 0, Col 0: Label "PR No:" + TextBlock bound to `{Binding CurrentPRNo}`
  - Row 0, Col 1: Label "Project:" + TextBlock bound to `{Binding CurrentProjectName}`
  - Row 1, Col 0: Label "Job Plan:" + TextBlock bound to `{Binding CurrentJobPlanName}`
  - Row 1, Col 1: Label "Last Sync:" + TextBlock bound to `{Binding LastSyncTimeFormatted}`
- Add a placeholder `TextBlock` with `Text="No project context available. Open a drawing with PR information."` that is visible only when `HasProjectContext` is false, using a `BooleanToVisibilityConverter` with an inverter or a `DataTrigger`
- The project info grid should be collapsed when `HasProjectContext` is false and visible when true

### How to verify
- [ ] Project info panel displays PR No, Project Name, Job Plan, and Last Sync fields (AC-03)
- [ ] Placeholder message is shown when no project context is available (AC-04)

---

## TASK-001-03-02: Add ViewModel properties for project info and sync time

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-001-03-03, TASK-001-03-04, TASK-001-03-07, TASK-001-03-08 |

### What to do
- Add `[ObservableProperty]` fields to `DashboardViewModel`: `string _currentPRNo` (default empty), `string _currentProjectName` (default empty), `string _currentJobPlanName` (default empty)
- Add `[ObservableProperty]` field: `DateTime? _lastSyncTime` (default null)
- Add a computed read-only property `string LastSyncTimeFormatted` that returns `LastSyncTime?.ToString("yyyy-MM-dd HH:mm") ?? "Never synced"`
- In `partial void OnLastSyncTimeChanged(DateTime? value)`, call `OnPropertyChanged(nameof(LastSyncTimeFormatted))` to update the formatted display
- Add a computed read-only property `bool HasProjectContext => !string.IsNullOrEmpty(CurrentPRNo)`
- In `partial void OnCurrentPRNoChanged(string value)`, call `OnPropertyChanged(nameof(HasProjectContext))` so the UI reacts to project context availability changes

### How to verify
- [ ] All project info properties are bindable and trigger UI updates (AC-03)
- [ ] LastSyncTimeFormatted returns "Never synced" when null and formatted date when set (AC-01)

---

## TASK-001-03-03: Implement logic to fetch current project context from IAutoCADService

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | TASK-001-03-02 |
| Blocks | None |

### What to do
- Add a `private async Task RefreshProjectContextAsync()` method to `DashboardViewModel`
- If `_autoCADService` is null or `!IsAutoCADConnected`, clear project properties and return
- Call `await _autoCADService.GetStatusAsync()` to get the current document info
- If a document is open, extract project context (PR No, Project Name, Job Plan) from the drawing parameters by calling `_autoCADService.GetLayoutsValues()` and parsing the `Parameters` dictionary
- Map extracted values to `CurrentPRNo`, `CurrentProjectName`, `CurrentJobPlanName`
- Call this method from the status polling timer (e.g., when AutoCAD connection state changes from disconnected to connected)
- If extraction fails or no project data is found, leave properties empty (which triggers the "no project context" placeholder via `HasProjectContext`)

### How to verify
- [ ] Project info (PR No, Project Name, Job Plan) is populated from AutoCAD when connected (AC-03)
- [ ] Project info is cleared when AutoCAD disconnects (AC-06)

---

## TASK-001-03-04: Implement logic to retrieve LastSyncTime from IOdooService

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-03-02 |
| Blocks | None |

### What to do
- Add a `private void RefreshLastSyncTime()` method to `DashboardViewModel`
- If `_odooService` is null, set `LastSyncTime = null` and return
- Call `_odooService.GetLastSyncTime()` which returns `DateTime?`
- Set `LastSyncTime` to the returned value, triggering the `OnLastSyncTimeChanged` partial method and updating `LastSyncTimeFormatted`
- Call this method during initial ViewModel construction and periodically from the status timer (e.g., every 5 seconds or after each sync operation)
- Ensure the formatted timestamp displays in user-friendly format (already handled by `LastSyncTimeFormatted` property from TASK-001-03-02)

### How to verify
- [ ] Last sync timestamp is displayed in the project info panel (AC-01)
- [ ] Timestamp updates after a successful Odoo sync (AC-05)

---

## TASK-001-03-05: Create Recent Operations summary section in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/DashboardPage.xaml` |
| Estimate | M |
| Depends On | TASK-001-03-01 |
| Blocks | None |

### What to do
- Below the Current Project panel, add a "Recent Operations" section with a header `TextBlock` ("Recent Operations", `FontSize="18"`, `FontWeight="SemiBold"`)
- Create a horizontal `StackPanel` or `UniformGrid Columns="2"` containing two summary cards:
  - Card 1: "Parameters Extracted" with a large count number bound to `{Binding ParametersExtractedCount}` and a label
  - Card 2: "BOQs Pushed" with a large count number bound to `{Binding BOQsPushedCount}` and a label
- Each card should use a `Border` with `CornerRadius="8"`, `Background="{StaticResource SurfaceBrush}"`, `Padding="16"`
- The count number should use `FontSize="32"` and `FontWeight="Bold"` with `Foreground="{StaticResource PrimaryBrush}"`
- The label should use `FontSize="12"` and `Foreground="{StaticResource TextSecondaryBrush}"`
- This section is priority "Could" (FR-001-011), so mark it with a comment indicating it can be deferred

### How to verify
- [ ] Dashboard displays counts for parameters extracted and BOQs pushed (AC-02)
- [ ] Counts are visually prominent with large numbers and descriptive labels

---

## TASK-001-03-06: Add ViewModel properties and logic for recent operation counts

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | TASK-001-03-02 |
| Blocks | None |

### What to do
- Add `[ObservableProperty]` fields: `int _parametersExtractedCount` (default 0), `int _boqsPushedCount` (default 0)
- Add public methods `void IncrementParametersExtracted(int count = 1)` and `void IncrementBOQsPushed(int count = 1)` for other ViewModels/services to call when operations complete
- Since `DashboardViewModel` is registered as singleton (TASK-001-01-05), these counts persist for the application session
- Consider adding a `ResetCounts()` method if session-based tracking is desired
- Optionally, persist counts to the local SQLite database via `AppDbContext` for cross-session tracking (stretch goal)
- Subscribe to relevant application events or use a messaging pattern (CommunityToolkit.Mvvm `WeakReferenceMessenger`) to receive operation completion notifications from other ViewModels:
  - `ParametersExtractedMessage` sent by the AutoCAD parameters ViewModel
  - `BOQPushedMessage` sent by the BOQ ViewModel
- Register message handlers in the constructor: `WeakReferenceMessenger.Default.Register<ParametersExtractedMessage>(this, (r, m) => ParametersExtractedCount += m.Count)`

### How to verify
- [ ] Counts increment when operations complete (AC-02)
- [ ] Counts persist within the application session (AC-02)

---

## TASK-001-03-07: Handle "no project context" state

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-03-02 |
| Blocks | None |

### What to do
- Add a `private void ClearProjectContext()` method that sets `CurrentPRNo = ""`, `CurrentProjectName = ""`, `CurrentJobPlanName = ""`, which in turn sets `HasProjectContext` to `false`
- Call `ClearProjectContext()` when AutoCAD disconnects (in the status change handler from US-001-01 TASK-001-01-04)
- Call `ClearProjectContext()` when the connected AutoCAD drawing does not contain project information (detected during `RefreshProjectContextAsync`)
- Add a property `string ProjectContextPlaceholder` that returns the fixed message "No project context available. Open a drawing with PR information." for XAML binding (or hardcode it in XAML)
- The XAML from TASK-001-03-01 should already show/hide the placeholder vs. project info grid based on `HasProjectContext`

### How to verify
- [ ] When no project context is available, the placeholder message is displayed (AC-04)
- [ ] When AutoCAD disconnects, the project info panel shows the placeholder (AC-04, AC-06)

---

## TASK-001-03-08: Subscribe to AutoCAD document change events to refresh project info

| Field | Value |
|-------|-------|
| Target | `ViewModels/DashboardViewModel.cs` |
| Estimate | S |
| Depends On | TASK-001-03-02, TASK-001-03-03 |
| Blocks | None |

### What to do
- In the status polling timer callback (from US-001-01 TASK-001-01-04), track the previously known document name in a `private string? _previousDocumentName` field
- On each timer tick, compare `AutoCADDocumentName` with `_previousDocumentName`
- If the document name changed (different drawing opened), call `await RefreshProjectContextAsync()` to re-extract project info from the new drawing
- Update `_previousDocumentName = AutoCADDocumentName` after processing
- Also call `RefreshProjectContextAsync()` when AutoCAD initially connects (transition from disconnected to connected)
- This polling approach is a pragmatic alternative to COM event subscriptions, following the existing `DispatcherTimer` pattern in `MainViewModel.cs`

### How to verify
- [ ] Project info updates when a different AutoCAD drawing is opened (AC-06)
- [ ] Project info updates when AutoCAD connection changes (AC-06)

---

## Dependency Graph
```
TASK-001-03-01 (Project Info XAML)
       │
       └──▶ TASK-001-03-05 (Recent Operations XAML)

TASK-001-03-02 (ViewModel Properties)
       │
       ├──▶ TASK-001-03-03 (Fetch Project Context) ──┐
       ├──▶ TASK-001-03-04 (Retrieve LastSyncTime)   │
       ├──▶ TASK-001-03-06 (Operation Counts)         │
       ├──▶ TASK-001-03-07 (No Context State)         │
       └──▶ TASK-001-03-08 (Document Change) ─────────┘
                   (depends on both 02 and 03)
```
