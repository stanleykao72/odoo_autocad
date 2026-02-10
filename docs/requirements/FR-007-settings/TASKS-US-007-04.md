# TASKS: US-007-04 — AutoCAD Timeout Settings

> **Parent US**: [US-007-04](US-007-04-autocad-timeout-settings.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 6S + 0M + 0L
> **Status**: In Progress

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [ ] `AppDbContext` with `ServerConfigs` and `SyncLogs` DbSets available

## Acceptance Criteria
- [x] AC-01: Settings Page displays a read-only field for AutoCAD COM ProgId showing "AutoCAD.Application" in the AutoCAD tab
- [x] AC-02: Settings Page displays a configurable connection timeout field with default value of 10 seconds
- [x] AC-03: Settings Page displays a configurable retry attempts field with default value of 3
- [ ] AC-04: Connection timeout validates as integer between 5 and 120 seconds inclusive, with inline error for out-of-range values
- [ ] AC-05: Retry attempts validates as integer between 1 and 10 inclusive, with inline error for out-of-range values
- [x] AC-06: Saving persists AutoCAD settings to `ServerConfigs` database table with keys `autocad.prog_id`, `autocad.connection_timeout`, `autocad.retry_attempts`
- [x] AC-07: AutoCAD settings are loaded from database on page initialization with fallback to defaults
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-04-01: Add AutoCAD tab with ProgId, Connection Timeout, and Retry Attempts fields in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Populate the AutoCAD tab (stub created in US-007-01) with a grouped layout
- Add a read-only `TextBox` with `IsReadOnly="True"` for COM ProgId, bound to `{Binding AutoCADProgId}`, styled with gray background to indicate non-editable
- Add a numeric `TextBox` for Connection Timeout with label "Connection Timeout (seconds):", bound to `{Binding AutoCADConnectionTimeout, UpdateSourceTrigger=PropertyChanged, ValidatesOnNotifyDataErrors=True}`
- Add a numeric `TextBox` or `NumericUpDown` for Retry Attempts with label "Retry Attempts:", bound to `{Binding AutoCADRetryAttempts, UpdateSourceTrigger=PropertyChanged, ValidatesOnNotifyDataErrors=True}`
- Add `InputScope="Number"` or `PreviewTextInput` handler to restrict input to integers only
- Include inline error display via `Validation.ErrorTemplate`

### How to verify
- [x] ProgId field renders as read-only showing "AutoCAD.Application" (AC-01)
- [x] Connection Timeout field renders with default value 10 (AC-02)
- [x] Retry Attempts field renders with default value 3 (AC-03)

---

## TASK-007-04-02: Add AutoCADProgId, AutoCADConnectionTimeout, and AutoCADRetryAttempts properties to ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-04-03, TASK-007-04-05 |

### What to do
- Add `[ObservableProperty] string _autoCADProgId` with default `"AutoCAD.Application"`
- Add `[ObservableProperty] int _autoCADConnectionTimeout` with default `10`
- Add `[ObservableProperty] int _autoCADRetryAttempts` with default `3`
- Wire the `On<Property>Changed` partials to update `HasUnsavedChanges` (from US-007-01 snapshot comparison)
- `AutoCADProgId` is read-only in the UI but still stored as a property for potential future configurability

### How to verify
- [x] Properties generate change notifications and are bindable (AC-01, AC-02, AC-03)
- [x] Modifying timeout or retry attempts sets `HasUnsavedChanges = true` (AC-06 context)

---

## TASK-007-04-03: Implement validation for timeout (5-120) and retry attempts (1-10) with inline error messages

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-04-02 |
| Blocks | None |

### What to do
- In `OnAutoCADConnectionTimeoutChanged`, validate: if value < 5 or value > 120, call `SetErrors("AutoCADConnectionTimeout", new[] { "Connection timeout must be between 5 and 120 seconds." })`, else `ClearErrors("AutoCADConnectionTimeout")`
- In `OnAutoCADRetryAttemptsChanged`, validate: if value < 1 or value > 10, call `SetErrors("AutoCADRetryAttempts", new[] { "Retry attempts must be between 1 and 10." })`, else `ClearErrors("AutoCADRetryAttempts")`
- Validation rules correspond to VR-007-008 (timeout) and VR-007-007 (retry attempts)
- Ensure `CanSave` property recalculates when these errors change

### How to verify
- [ ] Entering timeout outside 5-120 shows inline error message (AC-04)
- [ ] Entering retry attempts outside 1-10 shows inline error message (AC-05)
- [ ] Save button is disabled when validation errors exist (AC-04, AC-05)

---

## TASK-007-04-04: Add numeric input constraints to XAML fields (integer-only input)

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `PreviewTextInput` event handler (in code-behind or via attached behavior) that uses `Regex.IsMatch(e.Text, "[^0-9]")` to reject non-numeric input
- Apply this handler to both Connection Timeout and Retry Attempts TextBox controls
- Optionally use a `NumericUpDown` control from a WPF toolkit library if available, with `Minimum`, `Maximum`, and `Increment` properties set
- Handle paste events via `DataObject.Pasting` to prevent pasting non-numeric text

### How to verify
- [ ] Only integer characters can be typed into timeout and retry fields (AC-04, AC-05)
- [ ] Pasting non-numeric text is rejected (AC-04, AC-05)

---

## TASK-007-04-05: Persist AutoCAD settings to ServerConfigs table on Save

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-04-02 |
| Blocks | TASK-007-04-06 |

### What to do
- In the `SaveSettingsCommand` handler (from US-007-01), add persistence for AutoCAD settings using the upsert pattern on `ServerConfigs`:
  - `ConfigKey = "autocad.prog_id"`, `ConfigValue = AutoCADProgId`
  - `ConfigKey = "autocad.connection_timeout"`, `ConfigValue = AutoCADConnectionTimeout.ToString()`
  - `ConfigKey = "autocad.retry_attempts"`, `ConfigValue = AutoCADRetryAttempts.ToString()`
- Include these keys in the `SyncLog` details when settings are saved
- Update the original-values snapshot for AutoCAD fields after save

### How to verify
- [x] After Save, `ServerConfigs` contains rows for `autocad.prog_id`, `autocad.connection_timeout`, `autocad.retry_attempts` (AC-06)
- [x] `SyncLog` entry includes AutoCAD setting changes (AC-08)

---

## TASK-007-04-06: Load AutoCAD settings from ServerConfigs table on page initialization with default fallback

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-04-05 |
| Blocks | None |

### What to do
- In the ViewModel constructor or an `InitializeAsync()` method, query `ServerConfigs` for AutoCAD keys:
  - `autocad.prog_id` -> `AutoCADProgId` (fallback: `"AutoCAD.Application"`)
  - `autocad.connection_timeout` -> `AutoCADConnectionTimeout` (fallback: `10`, parse with `int.TryParse`)
  - `autocad.retry_attempts` -> `AutoCADRetryAttempts` (fallback: `3`, parse with `int.TryParse`)
- If no DB entries exist, fall back to values from `ConfigurationLoader.GetSettings().AutoCAD`
- Include these in the original-values snapshot for unsaved-changes tracking

### How to verify
- [x] On page load, fields show values from DB if present (AC-07)
- [x] On first load (no DB entries), fields show defaults: ProgId="AutoCAD.Application", Timeout=10, Retry=3 (AC-07)

---

## Dependency Graph
```
TASK-007-04-01 (XAML Fields)        ── independent
TASK-007-04-04 (Numeric Constraints) ── independent

TASK-007-04-02 (ViewModel Properties)
       │
       ├──▶ TASK-007-04-03 (Validation)
       └──▶ TASK-007-04-05 (Persist on Save)
                  │
                  └──▶ TASK-007-04-06 (Load on Init)

All tasks are small. Tasks 01, 02, and 04 can start in parallel.
```
