# TASKS: US-007-12 — Import Configuration

> **Parent US**: [US-007-12](US-007-12-import-configuration.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 4S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist; Save command must be functional for persisting imported values)
- [ ] US-007-11 (IFileDialogService must exist for ShowOpenFileDialog; exported JSON format defines the import schema)
- [ ] `ConfigurationLoader` with JSON parsing capability available

## Acceptance Criteria
- [ ] AC-01: Settings Page displays an "Import Configuration" button in the Advanced tab under Data Management section
- [ ] AC-02: Clicking "Import Configuration" opens a file open dialog with filter for JSON files (*.json)
- [ ] AC-03: The imported JSON file is validated as valid JSON format before applying
- [ ] AC-04: The imported JSON is validated for expected top-level keys: Application, Odoo, AutoCAD, MCP
- [ ] AC-05: If the JSON is not valid, an error dialog is shown: "The selected file is not a valid configuration file. Expected JSON format."
- [ ] AC-06: If the JSON is missing required sections, a warning is shown: "Configuration file is missing required sections: {missing_keys}. Import partially applied." and the available sections are still applied
- [ ] AC-07: After successful import, all settings fields in the ViewModel are updated with the imported values
- [ ] AC-08: Imported values are not automatically persisted; the user must click Save to persist them (import acts as "preview/load into form")
- [ ] AC-09: HasUnsavedChanges is set to true after a successful import
- [ ] AC-10: Import does not overwrite the API token unless explicitly present in the import file
- [ ] AC-11: Import operation is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-12-01: Add "Import Configuration" button to Data Management section of Advanced tab in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Data Management section of the Advanced tab (alongside Clear Cache and Export Config buttons), add a `Button` with `Content="Import Config"` bound to `{Binding ImportConfigCommand}`
- Add a `TextBlock` for import feedback (success/warning/error) bound to an import feedback property, with `Visibility` collapsed when empty
- Position consistently with the other Data Management buttons

### How to verify
- [ ] "Import Config" button renders in the Data Management section (AC-01)
- [ ] Feedback area is present but hidden when no message (AC-05, AC-06, AC-07)

---

## TASK-007-12-02: Implement ImportConfigCommand as IAsyncRelayCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-12-05, TASK-007-12-06, TASK-007-12-07 |

### What to do
- Add `IAsyncRelayCommand ImportConfigCommand` initialized as `new AsyncRelayCommand(ImportConfigAsync)`
- Implement `private async Task ImportConfigAsync()` that:
  1. Opens file open dialog (TASK-007-12-03)
  2. If user selects a file, reads and validates JSON (TASK-007-12-04)
  3. On valid: maps values to ViewModel properties (TASK-007-12-05)
  4. On invalid: shows error dialog (TASK-007-12-06)
  5. Logs to SyncLog (TASK-007-12-07)
- Add `[ObservableProperty] string _importMessage` (default empty) for feedback display
- Ensure the import does **not** call SaveSettingsCommand -- values are only loaded into the form

### How to verify
- [ ] `ImportConfigCommand` is callable from XAML binding (AC-01)
- [ ] Import loads values into form but does not persist to DB (AC-08)

---

## TASK-007-12-03: Implement IFileDialogService.ShowOpenFileDialog() for JSON file selection

| Field | Value |
|-------|-------|
| Target | `Services/IFileDialogService.cs` and `Services/FileDialogService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-12-02 |

### What to do
- Add `string? ShowOpenFileDialog(string filter)` method to `IFileDialogService` interface
- Implement in `FileDialogService`: wrap the Win32 `OpenFileDialog` with `Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"`
- Return the selected file path, or `null` if the user cancels
- Reuse the `IFileDialogService` created in US-007-11 (extend the existing interface)

### How to verify
- [ ] File open dialog appears with JSON filter when Import is clicked (AC-02)
- [ ] Cancelling the dialog returns null and the import is aborted (AC-02)

---

## TASK-007-12-04: Implement JSON parsing and validation for structure and expected top-level keys

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-12-05, TASK-007-12-06 |

### What to do
- Add a method `public (AppSettings? settings, List<string> missingKeys, string? error) ParseImportFile(string filePath)` to `ConfigurationLoader`
- Read the file content with `File.ReadAllText()`
- Attempt JSON deserialization with `JsonSerializer.Deserialize<JsonDocument>()` first to validate basic JSON structure
- If deserialization fails (malformed JSON), return `error = "The selected file is not a valid configuration file. Expected JSON format."`
- If valid JSON, check for expected top-level keys: "Application", "Odoo", "AutoCAD", "MCP"
- Collect missing keys into the `missingKeys` list
- Deserialize into `AppSettings` object for the sections that are present
- Return the parsed settings, missing keys list, and null error on success
- Validation rule corresponds to VR-007-013

### How to verify
- [ ] Invalid JSON returns an error string (AC-03, AC-05)
- [ ] Valid JSON with missing sections returns the missing keys list (AC-04, AC-06)
- [ ] Complete JSON returns all settings with empty missing keys list (AC-03, AC-04)

---

## TASK-007-12-05: Map imported JSON values to ViewModel properties and set HasUnsavedChanges

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-12-02, TASK-007-12-04 |
| Blocks | None |

### What to do
- After successful parsing, map the imported `AppSettings` values to ViewModel properties:
  - If `Odoo` section present: update `OdooServerUrl`, `OdooDatabaseName`, `OdooUsername`, `OdooTimeoutSeconds`
  - If `AutoCAD` section present: update `AutoCADConnectionTimeout`, `AutoCADRetryAttempts`
  - If `MCP` section present: update `MCPPort`, `MCPAutoStart`
  - If `Application` section present: no action (read-only info)
- For API token: only update `OdooApiToken` if the imported file explicitly contains a non-null/non-empty token value (AC-10)
- Set `HasUnsavedChanges = true` after mapping (AC-09)
- Clear any previous connection test results (`ConnectionTestResult = null`)
- Do NOT call `SaveSettingsCommand` -- the user must explicitly save

### How to verify
- [ ] ViewModel fields update to imported values (AC-07)
- [ ] `HasUnsavedChanges` is true after import (AC-09)
- [ ] API token is only updated if explicitly present in import file (AC-10)
- [ ] Values are not persisted to DB until user clicks Save (AC-08)

---

## TASK-007-12-06: Display success, partial-import warning, or error feedback based on import result

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-12-04 |
| Blocks | None |

### What to do
- Based on the return value from `ParseImportFile()`:
  - If `error` is not null (invalid JSON): show `MessageBox` with the error message and set `ImportMessage` accordingly (AC-05)
  - If `missingKeys` is non-empty: show a warning `MessageBox` with `$"Configuration file is missing required sections: {string.Join(", ", missingKeys)}. Import partially applied."` and set `ImportMessage` (AC-06)
  - If successful with no missing keys: set `ImportMessage = "Configuration imported successfully. Click Save to apply."` (AC-07)
- Clear the feedback message after 10 seconds
- Style: error in red, warning in orange, success in green

### How to verify
- [ ] Invalid JSON shows error dialog with correct message (AC-05)
- [ ] Missing sections shows warning dialog with the missing key names (AC-06)
- [ ] Full success shows success message (AC-07)

---

## TASK-007-12-07: Log import operation to SyncLog table

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-12-02 |
| Blocks | None |

### What to do
- After import (success, partial, or failure), add a `SyncLog` entry:
  - `SyncType = "settings_change"`
  - `Direction = "upload"` (importing from file into app)
  - `Success = true/false`
  - `Details` = JSON string: `{ "action": "import_config", "file_path": "{path}", "missing_sections": [...] }`
  - If failed: `ErrorMessage = error message`
  - `SyncTime = DateTime.UtcNow`
- Call `await _dbContext.SaveChangesAsync()` to persist the log entry

### How to verify
- [ ] After import, `SyncLog` contains entry with `SyncType = "settings_change"` and import details (AC-11)
- [ ] Failed imports are also logged with error details (AC-11)

---

## Dependency Graph
```
TASK-007-12-01 (XAML Button)          ── independent
TASK-007-12-03 (ShowOpenFileDialog)   ── independent
TASK-007-12-04 (JSON Parsing)         ── independent

TASK-007-12-02 (ImportConfigCommand) ◀── TASK-007-12-03
       │
       ├──▶ TASK-007-12-05 (Map to ViewModel) ◀── TASK-007-12-04
       │
       ├──▶ TASK-007-12-06 (Feedback) ◀── TASK-007-12-04
       │
       └──▶ TASK-007-12-07 (SyncLog Audit)

Tasks 01, 03, and 04 can start in parallel.
Task 02 needs 03 for file dialog.
Task 05 and 06 both need 04 for parsed results.
```
