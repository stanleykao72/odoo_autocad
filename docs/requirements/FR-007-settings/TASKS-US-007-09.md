# TASKS: US-007-09 — Change Log Level

> **Parent US**: [US-007-09](US-007-09-change-log-level.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 6S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [ ] .NET `ILogger<T>` infrastructure configured in DI container
- [ ] `AppDbContext` with `UserPreferences` and `SyncLogs` DbSets available

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a log level selector dropdown in the Advanced tab with options: Verbose, Debug, Information, Warning, Error, Fatal
- [ ] AC-02: The currently active log level is pre-selected on page load with default "Information"
- [ ] AC-03: Log level selection must be one of the valid options (validated)
- [ ] AC-04: Settings Page displays the current log file path as a read-only text field
- [ ] AC-05: Settings Page provides an "Open Log Folder" button that opens the log directory in Windows Explorer
- [ ] AC-06: Saving persists the log level to `UserPreferences` table with key `logging.level`
- [ ] AC-07: Log level change takes effect immediately upon save without requiring application restart
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-09-01: Add Logging section to Advanced tab with log level dropdown, log path display, and Open Log Folder button

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Populate the Advanced tab (stub created in US-007-01) with a "Logging" section group
- Add a `ComboBox` with label "Log Level:" bound to `{Binding SelectedLogLevel}` with items: Verbose, Debug, Information, Warning, Error, Fatal
- Add a read-only `TextBox` with label "Log Path:" bound to `{Binding LogFilePath}` with `IsReadOnly="True"` and gray background styling
- Add a `Button` with `Content="Open Log Folder"` bound to `{Binding OpenLogFolderCommand}`
- Style consistently with other Advanced tab controls

### How to verify
- [ ] Log level dropdown renders with six options (AC-01)
- [ ] Log path field renders as read-only text (AC-04)
- [ ] "Open Log Folder" button renders next to the log path (AC-05)

---

## TASK-007-09-02: Add SelectedLogLevel and LogFilePath properties to ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-09-04, TASK-007-09-05, TASK-007-09-06 |

### What to do
- Add `[ObservableProperty] string _selectedLogLevel` with default `"Information"`
- Add `[ObservableProperty] string _logFilePath` -- populated on init from the logging configuration (e.g., `Path.Combine(AppContext.BaseDirectory, "logs")` or from Serilog/NLog file sink path)
- Wire `OnSelectedLogLevelChanged` to update `HasUnsavedChanges`
- `LogFilePath` is read-only and not user-editable

### How to verify
- [ ] `SelectedLogLevel` defaults to "Information" and is bindable (AC-02)
- [ ] `LogFilePath` displays the actual log directory path (AC-04)

---

## TASK-007-09-03: Implement OpenLogFolderCommand that opens log directory in Windows Explorer

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add `IRelayCommand OpenLogFolderCommand` initialized as `new RelayCommand(OpenLogFolder)`
- Implement `private void OpenLogFolder()`:
  - Extract the directory from `LogFilePath` using `Path.GetDirectoryName()` if it's a file path, or use directly if it's a directory
  - Call `Process.Start(new ProcessStartInfo { FileName = logFolder, UseShellExecute = true })` to open Windows Explorer
  - Handle `InvalidOperationException` and `Win32Exception` gracefully (e.g., if the path doesn't exist, show a warning)
- If the log directory doesn't exist, create it before opening (using `Directory.CreateDirectory()`)

### How to verify
- [ ] Clicking "Open Log Folder" opens Windows Explorer at the log directory (AC-05)
- [ ] Non-existent log directory is created and then opened (AC-05)

---

## TASK-007-09-04: Implement validation for log level (must be one of the six valid options)

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-09-02 |
| Blocks | None |

### What to do
- Define a private static `HashSet<string> ValidLogLevels = new() { "Verbose", "Debug", "Information", "Warning", "Error", "Fatal" }`
- In `OnSelectedLogLevelChanged`, validate: if `!ValidLogLevels.Contains(SelectedLogLevel)`, call `SetErrors("SelectedLogLevel", new[] { "Invalid log level selected." })`, else `ClearErrors("SelectedLogLevel")`
- Validation rule corresponds to VR-007-012
- Since the ComboBox restricts selection to valid items, this is a defensive validation

### How to verify
- [ ] Only valid log levels pass validation (AC-03)
- [ ] Invalid log level values trigger validation error (AC-03)

---

## TASK-007-09-05: Persist log level to UserPreferences table on Save

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-09-02 |
| Blocks | TASK-007-09-07 |

### What to do
- In the `SaveSettingsCommand` handler, add persistence for log level:
  - Upsert `UserPreferences` row: `PreferenceKey = "logging.level"`, `PreferenceValue = SelectedLogLevel`, `DataType = "string"`
- Include the log level key in the `SyncLog` details JSON with `SyncType = "settings_change"`
- Update the original-values snapshot for `SelectedLogLevel` after save

### How to verify
- [ ] After Save, `UserPreferences` table contains `logging.level` with selected value (AC-06)
- [ ] `SyncLog` entry includes the log level change (AC-08)

---

## TASK-007-09-06: Apply log level change at runtime by reconfiguring the logging provider minimum level

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-09-02 |
| Blocks | None |

### What to do
- After successful save, dynamically reconfigure the logging minimum level:
  - Map the display name to .NET `LogLevel` enum: "Verbose" -> `LogLevel.Trace`, "Debug" -> `LogLevel.Debug`, "Information" -> `LogLevel.Information`, "Warning" -> `LogLevel.Warning`, "Error" -> `LogLevel.Error`, "Fatal" -> `LogLevel.Critical`
  - If using Serilog: call `Log.Logger = new LoggerConfiguration().MinimumLevel.Is(mappedLevel)...CreateLogger()`
  - If using built-in .NET logging with a custom filter: update the `ILoggingBuilder` filter via a mutable `LogLevel` holder injected at startup (e.g., `LogLevelOptions` class with a `MinimumLevel` property that the filter reads)
  - Alternatively, use `IConfigurationRoot.Reload()` after modifying the in-memory logging configuration
- Log a message at the new level: `"Log level changed to {SelectedLogLevel}"`
- The change must take effect immediately without restart

### How to verify
- [ ] After saving a log level change, new log entries use the updated level filter (AC-07)
- [ ] Changing from "Information" to "Debug" results in debug messages appearing in log output (AC-07)

---

## TASK-007-09-07: Load log level from UserPreferences table on page initialization with fallback to "Information"

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-09-05 |
| Blocks | None |

### What to do
- In the ViewModel constructor or `InitializeAsync()`, query `UserPreferences` for key `logging.level`
- If a value exists and is valid (in `ValidLogLevels`), set `SelectedLogLevel` to the persisted value
- If no entry exists or value is invalid, fall back to `"Information"`
- Include the loaded value in the original-values snapshot

### How to verify
- [ ] On page load, log level dropdown shows the persisted value from DB (AC-02)
- [ ] On first load (no DB entry), dropdown shows "Information" (AC-02)

---

## Dependency Graph
```
TASK-007-09-01 (XAML Logging Section)  ── independent
TASK-007-09-03 (Open Log Folder)       ── independent

TASK-007-09-02 (ViewModel Properties)
       │
       ├──▶ TASK-007-09-04 (Validation)
       ├──▶ TASK-007-09-05 (Persist on Save)
       │         │
       │         └──▶ TASK-007-09-07 (Load on Init)
       └──▶ TASK-007-09-06 (Runtime Log Level)

Tasks 01, 02, and 03 can start in parallel.
```
