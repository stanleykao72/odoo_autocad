# TASKS: US-007-01 — Configure Odoo Connection

> **Parent US**: [US-007-01](US-007-01-configure-odoo-connection.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 8 | **Effort**: 2S + 4M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure and CommunityToolkit.Mvvm must be in place)
- [ ] `AppDbContext` with `ServerConfigs`, `UserPreferences`, and `SyncLogs` DbSets available (already exists in `OdooAutoCAD.Data`)
- [ ] `ConfigurationLoader` with `LoadFromJson()` available (already exists in `OdooAutoCAD.Configuration`)

## Acceptance Criteria
- [ ] AC-01: Settings Page displays input fields for Odoo Server URL, Database Name, and Username
- [ ] AC-02: Settings Page displays a secure input field for Odoo API Token with masked display by default
- [ ] AC-03: API Token field has a show/hide toggle button that switches between PasswordBox and TextBox
- [ ] AC-04: Settings Page displays a connection timeout field with default value of 30 seconds
- [ ] AC-05: Server URL field validates format (must start with `http://` or `https://`) and shows inline error for invalid input
- [ ] AC-06: Database Name field validates non-empty, max 100 characters, alphanumeric/hyphens/underscores only
- [ ] AC-07: API Token field validates UUID format or non-empty string
- [ ] AC-08: Timeout field validates integer between 5 and 300 seconds inclusive
- [ ] AC-09: Clicking Save persists all Odoo connection settings to both `appsettings.json` and the `ServerConfigs` database table
- [ ] AC-10: API Token is persisted only to the `ServerConfigs` database table, never written to `appsettings.json` in plain text
- [ ] AC-11: Save button is disabled when any required field fails validation
- [ ] AC-12: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-01-01: Create SettingsPage XAML with TabControl and Connection tab containing Odoo fields

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-01-02, TASK-007-01-07 |

### What to do
- Create a new WPF `Page` in namespace `OdooAutoCAD.App.Views.Pages` with `x:Class="OdooAutoCAD.App.Views.Pages.SettingsPage"`
- Add corresponding code-behind `SettingsPage.xaml.cs` that sets `DataContext` to `SettingsViewModel` resolved from `App.Services`
- Define a `TabControl` with 5 tabs: Connection, AutoCAD, MCP, Appearance, Advanced (stubs for non-Connection tabs)
- In the Connection tab, add `TextBox` inputs for Server URL, Database Name, and Username, each with `Label` and bound to ViewModel properties via `{Binding OdooServerUrl, UpdateSourceTrigger=PropertyChanged, ValidatesOnNotifyDataErrors=True}`
- Add a `PasswordBox` and overlapping `TextBox` for API Token with visibility toggling bound to `IsTokenVisible`
- Add a numeric `TextBox` for Connection Timeout bound to `OdooTimeoutSeconds`
- Add a persistent footer bar with Save and Cancel buttons, where Save binds to `{Binding SaveSettingsCommand}` and is disabled when `CanSave` is false
- Reference `StaticResource` styles from the application theme (BackgroundBrush, SurfaceBrush, TextPrimaryBrush)

### How to verify
- [ ] Server URL, Database Name, Username text fields render in the Connection tab (AC-01)
- [ ] API Token field renders with masked display (PasswordBox visible by default) (AC-02)
- [ ] Connection Timeout field renders with default value 30 (AC-04)
- [ ] Save button is present in the footer bar (AC-09)

---

## TASK-007-01-02: Implement PasswordBox/TextBox toggle for API Token with show/hide button

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | TASK-007-01-01 |
| Blocks | None |

### What to do
- Add a `ToggleButton` or `Button` next to the API Token field bound to `ToggleTokenVisibilityCommand`
- Use a `BooleanToVisibilityConverter` to swap between `PasswordBox` (visible when `IsTokenVisible` is false) and `TextBox` (visible when `IsTokenVisible` is true)
- Both the `PasswordBox.Password` and `TextBox.Text` bind to `OdooApiToken`; since WPF `PasswordBox.Password` is not a DependencyProperty, use an attached behavior or code-behind event handler to sync values
- Add an eye icon (or text "Show"/"Hide") to the toggle button that changes based on `IsTokenVisible` state

### How to verify
- [ ] Clicking the toggle button switches between masked and plain-text display of the token (AC-03)
- [ ] Token value is preserved when toggling visibility (AC-03)

---

## TASK-007-01-03: Add SettingsViewModel with Odoo connection properties

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-01-04, TASK-007-01-05, TASK-007-01-07, TASK-007-01-08 |

### What to do
- Create `ViewModels/SettingsViewModel.cs` in namespace `OdooAutoCAD.App.ViewModels`
- Define `public partial class SettingsViewModel : ObservableValidator` using `CommunityToolkit.Mvvm.ComponentModel` (extends `ObservableObject` with `INotifyDataErrorInfo` support)
- Add `[ObservableProperty]` fields: `string _odooServerUrl`, `string _odooDatabaseName`, `string _odooUsername`, `string _odooApiToken`, `int _odooTimeoutSeconds` (default 30), `bool _isTokenVisible` (default false)
- Inject `AppDbContext`, `ConfigurationLoader`, and `ILogger<SettingsViewModel>` via constructor
- Add `IRelayCommand ToggleTokenVisibilityCommand` that toggles `IsTokenVisible`
- Load current values from `ServerConfigs` table on construction (upsert pattern: DB values override `appsettings.json` defaults)
- Register the ViewModel as **transient** in DI so it loads fresh values each time the Settings Page is navigated to
- Store a snapshot of original values for unsaved-changes tracking

### How to verify
- [ ] All six Odoo connection properties generate proper change notifications (AC-01, AC-02, AC-04)
- [ ] ToggleTokenVisibilityCommand toggles `IsTokenVisible` (AC-03)
- [ ] Properties load from `ServerConfigs` DB with fallback to `appsettings.json` defaults (AC-01)

---

## TASK-007-01-04: Implement INotifyDataErrorInfo validation for Odoo connection fields

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-01-03 |
| Blocks | TASK-007-01-05 |

### What to do
- Override the property setters (or use `[CustomValidation]` attributes) to validate each field on change:
  - `OdooServerUrl`: must not be empty (VR-007-002), must be valid URL starting with `http://` or `https://` using `Uri.TryCreate(value, UriKind.Absolute, out _)` (VR-007-001)
  - `OdooDatabaseName`: must not be empty, max 100 chars, regex `^[a-zA-Z0-9_-]+$` (VR-007-003)
  - `OdooApiToken`: must be non-empty; optionally validate UUID format via regex `^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$` (VR-007-004)
  - `OdooTimeoutSeconds`: must be integer between 5 and 300 inclusive (VR-007-005)
- Use `SetErrors()` or `ValidateProperty()` from `ObservableValidator` to surface errors
- Expose a computed `bool CanSave` property that returns `true` only when `HasErrors` is false
- Bind inline error messages in XAML via `Validation.ErrorTemplate` and `ValidatesOnNotifyDataErrors=True`

### How to verify
- [ ] Invalid URL format shows inline error (AC-05)
- [ ] Empty or invalid Database Name shows inline error (AC-06)
- [ ] Empty or invalid API Token shows inline error (AC-07)
- [ ] Out-of-range timeout shows inline error (AC-08)
- [ ] Save button is disabled when any validation fails (AC-11)

---

## TASK-007-01-05: Implement SaveSettingsCommand persisting to appsettings.json and ServerConfigs DB

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | L |
| Depends On | TASK-007-01-03, TASK-007-01-04 |
| Blocks | TASK-007-01-08 |

### What to do
- Implement `IAsyncRelayCommand SaveSettingsCommand` that executes the following on click:
  1. Re-validate all fields; abort if `HasErrors` is true
  2. Persist non-sensitive settings (Server URL, Database, Username, Timeout) to `appsettings.json` via `ConfigurationLoader.SaveToJson()` -- token is **excluded**
  3. Persist all settings (including token) to the `ServerConfigs` table using the upsert pattern: `FirstOrDefaultAsync(c => c.ConfigKey == key)`, then create or update
  4. ConfigKeys: `odoo.server_url`, `odoo.database`, `odoo.username`, `odoo.api_token`, `odoo.timeout_seconds`
  5. Call `await _dbContext.SaveChangesAsync()`
  6. Update the original-values snapshot so `HasUnsavedChanges` resets to false
  7. Set `IsSaving = true` during the operation, `false` on completion
- Handle `DbUpdateException` and `IOException` with user-friendly error dialogs matching FR error handling table
- On success, display a brief success notification

### How to verify
- [ ] After Save, non-token values appear in `appsettings.json` (AC-09)
- [ ] After Save, all values including token appear in `ServerConfigs` DB table (AC-09)
- [ ] Token is NOT present in `appsettings.json` after Save (AC-10)
- [ ] Save button is disabled when validation errors exist (AC-11)

---

## TASK-007-01-06: Extend ConfigurationLoader with SaveToJson method excluding sensitive fields

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-01-05 |

### What to do
- Add a new method `public void SaveToJson(AppSettings settings, string fileName = "appsettings.json")` to `ConfigurationLoader`
- Serialize the `AppSettings` object to JSON using `System.Text.Json.JsonSerializer.Serialize()` with `JsonSerializerOptions { WriteIndented = true }`
- Before serialization, create a sanitized copy that excludes sensitive fields: ensure no `api_token` or `token` property appears in the output JSON
- Write the JSON string to the file at `Path.Combine(AppContext.BaseDirectory, fileName)` with proper error handling for `IOException` and `UnauthorizedAccessException`
- Add a complementary method `public AppSettings LoadFromJsonFile(string fileName)` that reads and deserializes without using `IConfigurationBuilder` (for use by import/export features)

### How to verify
- [ ] `SaveToJson()` writes valid JSON to the target file (AC-09)
- [ ] Output JSON does not contain API token or sensitive credential fields (AC-10)
- [ ] `IOException` is thrown and catchable when the file is read-only or path is invalid (AC-09)

---

## TASK-007-01-07: Implement HasUnsavedChanges tracking with original-values snapshot

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | L |
| Depends On | TASK-007-01-03, TASK-007-01-01 |
| Blocks | None |

### What to do
- In the ViewModel constructor, after loading current values, capture a `Dictionary<string, string>` snapshot of all property values (Server URL, Database, Username, Token, Timeout, plus all other settings tab values)
- Override each `[ObservableProperty]` partial `On<Property>Changed` method to recompute `HasUnsavedChanges` by comparing current values against the snapshot
- Expose `[ObservableProperty] bool _hasUnsavedChanges` bound to the Save button label (e.g., "Save Settings*" when true)
- Implement `INavigationAware` or hook into `INavigationService.CanNavigateFrom()` to intercept navigation when `HasUnsavedChanges` is true
- Show a three-button dialog ("Save / Discard / Cancel") when the user tries to navigate away with unsaved changes
- Reset `HasUnsavedChanges` to false after successful Save or explicit Discard

### How to verify
- [ ] Modifying any field sets `HasUnsavedChanges` to true (AC-11 context)
- [ ] Save button label changes to indicate unsaved state (AC-11 context)
- [ ] Navigating away with unsaved changes shows a confirmation dialog (AC-11 context)

---

## TASK-007-01-08: Add SyncLog audit entry on settings save

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-01-05 |
| Blocks | None |

### What to do
- After successful save in `SaveSettingsCommand`, create a new `SyncLog` entry via `_dbContext.SyncLogs.Add()`
- Set fields: `SyncType = "settings_change"`, `Direction = "upload"`, `Success = true`, `SyncTime = DateTime.UtcNow`, `Details` = JSON string listing which ConfigKeys were changed and their old/new values (exclude token value from Details for security)
- Set `RecordsProcessed` to the count of changed settings, `RecordsUpdated` to the same count
- Call `await _dbContext.SaveChangesAsync()` to persist the log entry
- If save fails, still log the attempt with `Success = false` and `ErrorMessage` set to the exception message

### How to verify
- [ ] After saving settings, a new row exists in `SyncLogs` table with `SyncType = "settings_change"` (AC-12)
- [ ] The `Details` JSON lists changed keys but does not expose the token value (AC-12, AC-10)

---

## Dependency Graph
```
TASK-007-01-01 (SettingsPage XAML)
       │
       ├──▶ TASK-007-01-02 (Token Toggle)
       │
       └──▶ TASK-007-01-07 (Unsaved Changes Tracking) ◀── TASK-007-01-03

TASK-007-01-03 (ViewModel Properties)
       │
       ├──▶ TASK-007-01-04 (Validation)
       │         │
       │         └──▶ TASK-007-01-05 (SaveSettingsCommand) ◀── TASK-007-01-06
       │                   │
       │                   └──▶ TASK-007-01-08 (SyncLog Audit)
       │
       └──▶ TASK-007-01-07 (Unsaved Changes Tracking)

TASK-007-01-06 (ConfigurationLoader SaveToJson)
       │
       └──▶ TASK-007-01-05 (SaveSettingsCommand)

Tasks 01 and 03 can start in parallel.
Task 06 can start independently and is required before 05.
```
