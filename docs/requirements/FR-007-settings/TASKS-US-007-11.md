# TASKS: US-007-11 — Export Configuration

> **Parent US**: [US-007-11](US-007-11-export-configuration.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 5S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure and SaveToJson in ConfigurationLoader)
- [ ] `AppDbContext` with `ServerConfigs`, `UserPreferences`, and `SyncLogs` DbSets available
- [ ] `IFileDialogService` interface defined (or to be created as part of this US)

## Acceptance Criteria
- [ ] AC-01: Settings Page displays an "Export Configuration" button in the Advanced tab under Data Management section
- [ ] AC-02: Clicking "Export Configuration" opens a file save dialog with default filter for JSON files (*.json)
- [ ] AC-03: The exported JSON file contains all current settings organized under top-level keys: Application, Odoo, AutoCAD, MCP
- [ ] AC-04: The exported JSON excludes sensitive fields (API token is not included in the export)
- [ ] AC-05: On successful export, a success message is displayed with the export file path
- [ ] AC-06: If export fails due to file permissions or disk error, an error dialog is shown: "Failed to export configuration to {path}. Check file permissions."
- [ ] AC-07: The export operation does not modify any current settings
- [ ] AC-08: Export operation is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-11-01: Add "Export Configuration" button to Data Management section of Advanced tab in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Data Management section of the Advanced tab (alongside the Clear Cache button from US-007-10), add a `Button` with `Content="Export Config"` bound to `{Binding ExportConfigCommand}`
- Position the button in a horizontal `StackPanel` or `WrapPanel` alongside the Clear Cache and Import Config buttons
- Add a `TextBlock` for export feedback (success/error) bound to an export feedback property, with `Visibility` collapsed when empty

### How to verify
- [ ] "Export Config" button renders in the Data Management section (AC-01)
- [ ] Feedback area is present but hidden when no message (AC-05, AC-06)

---

## TASK-007-11-02: Implement ExportConfigCommand as IAsyncRelayCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-11-05, TASK-007-11-06, TASK-007-11-07 |

### What to do
- Add `IAsyncRelayCommand ExportConfigCommand` initialized as `new AsyncRelayCommand(ExportConfigAsync)`
- Implement `private async Task ExportConfigAsync()` that:
  1. Opens file save dialog (TASK-007-11-03)
  2. If user selects a path, serializes settings to JSON (TASK-007-11-04)
  3. Writes the JSON to the file (TASK-007-11-05)
  4. Shows success or error feedback (TASK-007-11-06)
  5. Logs to SyncLog (TASK-007-11-07)
- Add `[ObservableProperty] string _exportMessage` (default empty) for feedback display

### How to verify
- [ ] `ExportConfigCommand` is callable from XAML binding (AC-01)
- [ ] Command does not modify any current ViewModel properties (AC-07)

---

## TASK-007-11-03: Implement IFileDialogService.ShowSaveFileDialog() for JSON file selection

| Field | Value |
|-------|-------|
| Target | `Services/IFileDialogService.cs` and `Services/FileDialogService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-11-02 |

### What to do
- Create `IFileDialogService` interface with method `string? ShowSaveFileDialog(string defaultFileName, string filter)`
- Implement `FileDialogService`: wrap the Win32 `SaveFileDialog` with `Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"` and `DefaultExt = ".json"`
- Set default file name to `"odoo-autocad-config.json"` or similar
- Return the selected file path, or `null` if the user cancels
- Register `FileDialogService` as transient in DI
- The abstraction enables unit testing with a mock file dialog

### How to verify
- [ ] File save dialog opens with JSON filter when Export is clicked (AC-02)
- [ ] Cancelling the dialog returns null and the export is aborted (AC-02)

---

## TASK-007-11-04: Serialize current settings to JSON excluding sensitive fields (token)

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-11-05 |

### What to do
- Add a method `public string SerializeForExport(AppSettings settings)` to `ConfigurationLoader`
- Create a deep copy of the `AppSettings` object (or use a DTO) to avoid mutating the original
- Clear sensitive fields: ensure no API token appears in the output (e.g., set `Odoo.ApiToken = null` or remove the property entirely via `JsonIgnore` conditional serialization)
- Add an `_export_metadata` section with: `ExportTimestamp = DateTime.UtcNow`, `AppVersion = settings.Application.Version`, `MachineName = Environment.MachineName`
- Serialize to indented JSON using `System.Text.Json.JsonSerializer.Serialize()` with `JsonSerializerOptions { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull }`
- The output should have top-level keys: `Application`, `Odoo`, `AutoCAD`, `MCP` (matching the `appsettings.json` structure for import compatibility)

### How to verify
- [ ] Output JSON contains Application, Odoo, AutoCAD, MCP sections (AC-03)
- [ ] API token is not present in the output JSON (AC-04)
- [ ] Output includes export metadata (timestamp, version) (AC-03)

---

## TASK-007-11-05: Write serialized JSON to the user-selected file path with error handling

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-11-02, TASK-007-11-04 |
| Blocks | TASK-007-11-06 |

### What to do
- In `ExportConfigAsync()`, after getting the file path and serialized JSON:
  - Use `await File.WriteAllTextAsync(filePath, jsonContent)` to write the file
  - Wrap in try/catch for `IOException`, `UnauthorizedAccessException`, and `DirectoryNotFoundException`
  - On success, proceed to success feedback
  - On failure, proceed to error feedback with the exception message

### How to verify
- [ ] JSON file is created at the user-selected path (AC-03)
- [ ] File write failure is caught and reported (AC-06)

---

## TASK-007-11-06: Display success or error feedback after export operation

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-11-05 |
| Blocks | None |

### What to do
- On success: set `ExportMessage = $"Configuration exported successfully to {filePath}"`
- On failure: set `ExportMessage = $"Failed to export configuration to {filePath}. Check file permissions."` and optionally show a `MessageBox` error dialog with full error details
- Clear the feedback message after 10 seconds using `Task.Delay` with a cancellation token
- Style success in green and error in red in the XAML

### How to verify
- [ ] Success message displays with the export file path (AC-05)
- [ ] Error message displays with the path and error hint on failure (AC-06)

---

## TASK-007-11-07: Log export operation to SyncLog table

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-11-02 |
| Blocks | None |

### What to do
- After export (success or failure), add a `SyncLog` entry:
  - `SyncType = "settings_change"`
  - `Direction = "download"` (exporting from app to file)
  - `Success = true/false`
  - `Details` = JSON string: `{ "action": "export_config", "file_path": "{path}" }`
  - If failed: `ErrorMessage = ex.Message`
  - `SyncTime = DateTime.UtcNow`
- Call `await _dbContext.SaveChangesAsync()` to persist the log entry

### How to verify
- [ ] After export, `SyncLog` contains entry with `SyncType = "settings_change"` and export details (AC-08)
- [ ] Failed exports are also logged with error details (AC-08)

---

## Dependency Graph
```
TASK-007-11-01 (XAML Button)         ── independent
TASK-007-11-03 (IFileDialogService)  ── independent
TASK-007-11-04 (Serialize for Export) ── independent

TASK-007-11-02 (ExportConfigCommand) ◀── TASK-007-11-03
       │
       ├──▶ TASK-007-11-05 (Write File) ◀── TASK-007-11-04
       │         │
       │         └──▶ TASK-007-11-06 (Feedback)
       │
       └──▶ TASK-007-11-07 (SyncLog Audit)

Tasks 01, 03, and 04 can start in parallel.
Task 02 needs 03 before it can call the file dialog.
Task 05 needs 02 (orchestration) and 04 (serialization).
```
