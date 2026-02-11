# TASKS: US-008-04 — System Log Panel

> **Parent US**: [US-008-04](US-008-04-system-log-panel.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 5S + 3M + 0L
> **Status**: Complete

## Prerequisites
- [x] US-008-01 completed (main window layout must include the content area where the status bar and log panel reside)

## Acceptance Criteria
- [x] AC-01: A status bar at the bottom of the content area displays a `StatusMessage` text, the current date/time, and the application version.
- [x] AC-02: A dedicated system log panel is available (as a collapsible bottom panel and a dedicated Logs page) to display operational log messages.
- [x] AC-03: The system log panel auto-scrolls to the latest entry when new messages are added (with checkbox toggle).
- [x] AC-04: Log entries include timestamps (`HH:mm:ss.fff`) and source prefixes ("Odoo", "AutoCAD", "Settings", "App").
- [x] AC-05: The status bar displays three columns: status message (left, fills remaining space), date/time (center-right), and version string (far right).
- [x] AC-06: The status bar uses `SurfaceBrush` background with a top border separator.

---

## TASK-008-04-01: Define status bar XAML with three-column layout ✅

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Status | **Complete** (Sprint 1) |

### Actual Implementation
Status bar in MainWindow.xaml Grid Row 3 (was Row 2 before log panel insertion): `Border` with `SurfaceBrush` background, three-column Grid with StatusMessage, CurrentDateTime, and "v6.0.0".

---

## TASK-008-04-02: Add StatusMessage, CurrentDateTime, and version observable properties to MainViewModel ✅

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Status | **Complete** (Sprint 1) |

### Actual Implementation
`[ObservableProperty]` fields for `_statusMessage`, `_currentDateTime` in MainViewModel with DispatcherTimer update.

---

## TASK-008-04-03: Create IAppLogService + AppLogService + collapsible LogPanelControl ✅

| Field | Value |
|-------|-------|
| Target | `Services/IAppLogService.cs`, `Services/AppLogService.cs`, `Views/Controls/LogPanelControl.xaml/.cs`, `Views/MainWindow.xaml` |
| Estimate | M |
| Status | **Complete** |

### Actual Implementation (deviated from spec — improved architecture)
Instead of a Serilog sink writing to `MainViewModel.LogEntries`, a centralized **`IAppLogService`** was created:

- **`IAppLogService`** interface with `Log(message, source, level)`, `Clear()`, `ObservableCollection<AppLogEntry> LogEntries`
- **`AppLogService`** singleton: thread-safe dispatcher marshaling, Serilog forwarding, 2000-entry cap
- **`AppLogEntry`** record: `(DateTime Timestamp, string Source, AppLogLevel Level, string Message)`
- **`AppLogLevel`** enum: `Debug`, `Info`, `Warning`, `Error`
- **`LogPanelControl`** reusable UserControl with:
  - ListView with GridView columns: Time (`HH:mm:ss.fff`), Source, Message
  - Color-coded rows by level: Gray=Debug, Black=Info, Orange=Warning, Red=Error
  - Source filter ComboBox (dynamically populated), Level filter ComboBox
  - Clear button + Auto-scroll checkbox
  - `CollectionViewSource` for filtering
  - Monospace font (Consolas, 12pt)
- **MainWindow** updated: content area now has 4 rows — Header, Frame, Expander (LogPanelControl, Height=180, collapsed by default), StatusBar
- **MainWindow.xaml.cs** wires `LogPanel.DataContext = IAppLogService` from DI

---

## TASK-008-04-04: Create LogsPage full-page log viewer + "Logs" nav button ✅

| Field | Value |
|-------|-------|
| Target | `Views/Pages/LogsPage.xaml/.cs`, `Views/MainWindow.xaml` |
| Estimate | M |
| Status | **Complete** |

### Actual Implementation
- **`LogsPage`**: Full-page log viewer with header ("Application Logs") and full-height `LogPanelControl`
- Code-behind resolves `IAppLogService` from DI, sets as `FullLogPanel.DataContext`
- **"Logs" nav button** (`BtnLogs`) added to sidebar between "AI Assistant" and "Settings"
- Navigation works via existing `NavigationService` reflection pattern (`LogsPage`)

---

## TASK-008-04-05: Implement auto-scroll + source/level filtering ✅

| Field | Value |
|-------|-------|
| Target | `Views/Controls/LogPanelControl.xaml.cs` |
| Estimate | S |
| Status | **Complete** |

### Actual Implementation
- Auto-scroll via `CollectionChanged` event → `LogListView.ScrollIntoView(last item)`, controlled by "Auto-scroll" checkbox
- Source filter: dynamic `ComboBox` populated from `_knownSources` HashSet as new sources appear
- Level filter: `ComboBox` with All/Debug/Info/Warning/Error options
- Both filters use `CollectionViewSource.Filter` event for real-time filtering

---

## TASK-008-04-06: Add DispatcherTimer to update CurrentDateTime every second ✅

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Status | **Complete** (Sprint 1) |

### Actual Implementation
`DispatcherTimer` with 1-second interval in MainViewModel updates `CurrentDateTime`.

---

## TASK-008-04-07: Migrate OdooConnectionViewModel to IAppLogService ✅

| Field | Value |
|-------|-------|
| Target | `ViewModels/OdooConnectionViewModel.cs`, `Views/Pages/OdooPage.xaml` |
| Estimate | M |
| Status | **Complete** |

### Actual Implementation
- Removed `ObservableCollection<string> LogEntries`, `AddLog()`, `ClearLogCommand` from OdooConnectionViewModel
- Added `IAppLogService _logService` constructor parameter
- All `AddLog("...")` calls replaced with `_logService.Log("...", "Odoo")` (with appropriate `AppLogLevel`)
- Removed "Connection Log" Expander/ListBox/Clear button block from OdooPage.xaml
- Odoo logs now appear in the shared bottom panel and Logs page

---

## TASK-008-04-08: Add IAppLogService log calls to AutoCADViewModel and SettingsViewModel ✅

| Field | Value |
|-------|-------|
| Target | `ViewModels/AutoCADViewModel.cs`, `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Status | **Complete** |

### Actual Implementation
- **AutoCADViewModel**: Added `IAppLogService _logService` parameter. Log calls at: connect success/fail, disconnect, layout refresh, extraction start/complete
- **SettingsViewModel**: Added `IAppLogService _logService` parameter. Log calls at: settings loaded, settings saved, save error
- **SettingsViewModelTests**: Updated with `Mock<IAppLogService>` constructor parameter (all 22 tests pass)

---

## Dependency Graph (Actual)

```
TASK-008-04-01 (Status bar XAML) ✅ Sprint 1
TASK-008-04-02 (ViewModel properties) ✅ Sprint 1
TASK-008-04-06 (DateTime timer) ✅ Sprint 1

TASK-008-04-03 (IAppLogService + LogPanelControl + MainWindow panel) ✅
    ├── TASK-008-04-04 (LogsPage + nav button) ✅
    ├── TASK-008-04-05 (Auto-scroll + filtering) ✅
    ├── TASK-008-04-07 (OdooConnectionViewModel migration) ✅
    └── TASK-008-04-08 (AutoCADViewModel + SettingsViewModel) ✅
```
