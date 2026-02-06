# TASKS: US-007-02 — Switch Environments

> **Parent US**: [US-007-02](US-007-02-switch-environments.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 3S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (Odoo connection fields must exist in SettingsViewModel for environment values to populate them)
- [ ] `ConfigurationLoader` with `LoadFromJson()` available (already exists in `OdooAutoCAD.Configuration`)
- [ ] `AppDbContext` with `UserPreferences` and `SyncLogs` DbSets available

## Acceptance Criteria
- [ ] AC-01: Settings Page displays an environment selector dropdown with options: Development, Staging, Production
- [ ] AC-02: The currently active environment is pre-selected in the dropdown on page load
- [ ] AC-03: Selecting a new environment loads the corresponding settings from `appsettings.{Environment}.json`
- [ ] AC-04: A confirmation dialog appears before overwriting current values when switching environments
- [ ] AC-05: Confirmation dialog shows which settings will change (server URL, database name, etc.)
- [ ] AC-06: Cancelling the confirmation dialog reverts the dropdown to the previous environment selection
- [ ] AC-07: After confirming environment switch, all connection fields update to reflect the new environment values
- [ ] AC-08: Environment selection is persisted to `UserPreferences` table with key `app.environment`
- [ ] AC-09: Environment switch is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-02-01: Add environment selector dropdown to Connection tab in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add an "Environment" section to the Connection tab below the Odoo connection fields
- Add a `ComboBox` bound to `SelectedEnvironment` with `ItemsSource="{Binding AvailableEnvironments}"`
- Add a label "Active Environment:" next to the dropdown
- Wire the `SelectionChanged` behavior to `ChangeEnvironmentCommand` using `EventToCommand` binding or a `Command` property approach
- Style consistently with the other Connection tab fields

### How to verify
- [ ] Dropdown renders with Development, Staging, Production options (AC-01)
- [ ] The current environment is pre-selected on page load (AC-02)

---

## TASK-007-02-02: Add SelectedEnvironment, AvailableEnvironments, and ChangeEnvironmentCommand to ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-02-04, TASK-007-02-05, TASK-007-02-06 |

### What to do
- Add `[ObservableProperty] string _selectedEnvironment` with default "Production"
- Add `ObservableCollection<string> AvailableEnvironments` initialized with `{ "Development", "Staging", "Production" }`
- Add `IRelayCommand<string> ChangeEnvironmentCommand` that triggers the environment switch flow
- On construction, load the persisted environment from `UserPreferences` table (key `app.environment`) and set `SelectedEnvironment`
- Store `_previousEnvironment` field to enable revert on cancel
- Add a private field `_isEnvironmentSwitching` to prevent re-entrant switches

### How to verify
- [ ] AvailableEnvironments contains three entries (AC-01)
- [ ] SelectedEnvironment loads from UserPreferences on init (AC-02)
- [ ] ChangeEnvironmentCommand is wired and callable (AC-03)

---

## TASK-007-02-03: Extend ConfigurationLoader.GetAvailableConfigs() to discover appsettings.{env}.json files

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-02-05 |

### What to do
- Modify `GetAvailableConfigs()` to also scan `AppContext.BaseDirectory` for `appsettings.*.json` files (in addition to existing YAML scanning)
- Parse environment names from file names using pattern `appsettings.{envName}.json`
- Return discovered environment names as strings (e.g., "Development", "Staging", "Production")
- Add a new method `public AppSettings LoadEnvironmentConfig(string environmentName)` that loads `appsettings.{environmentName}.json` and returns the parsed `AppSettings`
- Handle `FileNotFoundException` gracefully when an environment file does not exist

### How to verify
- [ ] `GetAvailableConfigs()` returns environment names matching existing `appsettings.*.json` files (AC-01)
- [ ] `LoadEnvironmentConfig("Staging")` loads and returns settings from `appsettings.Staging.json` (AC-03)

---

## TASK-007-02-04: Implement confirmation dialog with change preview and cancel handling

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-02-02 |
| Blocks | TASK-007-02-05 |

### What to do
- In `ChangeEnvironmentCommand` handler, before applying changes, load the target environment config via `ConfigurationLoader.LoadEnvironmentConfig()`
- Compare current ViewModel values against the target environment values; build a list of diffs (e.g., "Server URL: current -> new", "Database: current -> new")
- Show a confirmation dialog with title "Switch Environment", message body listing the changes, and two buttons: "Confirm" and "Cancel"
- If user cancels, revert `SelectedEnvironment` to `_previousEnvironment` (suppress change event during revert)
- If user confirms, proceed to apply the new values (TASK-007-02-05)

### How to verify
- [ ] Confirmation dialog appears when selecting a new environment (AC-04)
- [ ] Dialog lists specific settings that will change (AC-05)
- [ ] Cancelling reverts the dropdown to the previous selection (AC-06)

---

## TASK-007-02-05: Load environment-specific values and populate ViewModel fields on confirm

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-02-03, TASK-007-02-04 |
| Blocks | None |

### What to do
- After user confirms the environment switch, apply the loaded `AppSettings` values to the ViewModel properties:
  - `OdooServerUrl = envSettings.Odoo.ServerUrl`
  - `OdooDatabaseName = envSettings.Odoo.Database`
  - `OdooUsername = envSettings.Odoo.Username`
  - `OdooTimeoutSeconds = envSettings.Odoo.TimeoutSeconds`
  - Apply AutoCAD and MCP settings similarly if present in the environment file
- Set `HasUnsavedChanges = true` since the loaded values are not yet persisted
- Clear any previous connection test results (`ConnectionTestResult = null`)
- Log the environment switch to `SyncLog` with details of old and new environment names

### How to verify
- [ ] After confirming, all connection fields update to new environment values (AC-07)
- [ ] HasUnsavedChanges is true after switch (values not yet saved) (AC-07)

---

## TASK-007-02-06: Persist selected environment to UserPreferences table

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-02-02 |
| Blocks | None |

### What to do
- In the `SaveSettingsCommand` handler (from US-007-01), include persisting the `SelectedEnvironment` value
- Use the upsert pattern on `UserPreferences` table: `PreferenceKey = "app.environment"`, `PreferenceValue = SelectedEnvironment`, `DataType = "string"`
- Also add a `SyncLog` entry for the environment change with `SyncType = "settings_change"` and `Details` containing old and new environment names
- On next page load, the constructor reads this preference to pre-select the correct environment

### How to verify
- [ ] After saving, `UserPreferences` table contains `app.environment` with the selected value (AC-08)
- [ ] `SyncLog` records the environment switch (AC-09)

---

## Dependency Graph
```
TASK-007-02-01 (XAML Dropdown)           ── independent, can start immediately

TASK-007-02-02 (ViewModel Properties)
       │
       ├──▶ TASK-007-02-04 (Confirmation Dialog)
       │         │
       │         └──▶ TASK-007-02-05 (Apply Environment Values) ◀── TASK-007-02-03
       │
       └──▶ TASK-007-02-06 (Persist to UserPreferences)

TASK-007-02-03 (ConfigurationLoader Extension) ── independent, can start immediately
       │
       └──▶ TASK-007-02-05 (Apply Environment Values)

Tasks 01, 02, and 03 can start in parallel.
```
