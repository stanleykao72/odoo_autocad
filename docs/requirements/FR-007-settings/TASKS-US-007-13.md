# TASKS: US-007-13 — View App Info

> **Parent US**: [US-007-13](US-007-13-view-app-info.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 7S + 0M + 0L
> **Status**: Done

## Prerequisites
- [x] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [x] `AppDbContext` available for extracting the database connection string path
- [x] `ConfigurationLoader` available for extracting the config file path

## Acceptance Criteria
- [x] AC-01: Settings Page displays the application version number (e.g., "6.0.0") as a read-only field in the Advanced tab under Application Info section
- [x] AC-02: Settings Page displays the database file path (e.g., `C:\Users\...\database.db`) as a read-only field
- [x] AC-03: Settings Page displays the config file path (e.g., `appsettings.json`) as a read-only field
- [x] AC-04: Settings Page displays the .NET Runtime version (e.g., ".NET 8.0") as a read-only field
- [x] AC-05: All Application Info fields are read-only and cannot be edited by the user
- [x] AC-06: Application info values are populated automatically on page load from runtime metadata
- [x] AC-07: Version number is sourced from the assembly version or application metadata
- [x] AC-08: Database path is sourced from the EF Core connection string or AppDbContext configuration

---

## TASK-007-13-01: Add Application Info section to Advanced tab with read-only fields

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-13-07 |

### What to do
- In the Advanced tab (below the Data Management section from US-007-10/11/12), add an "Application Info" section group
- Add four labeled `TextBox` controls, each with `IsReadOnly="True"`:
  - "Version:" bound to `{Binding AppVersion}`
  - "Database:" bound to `{Binding DatabasePath}`
  - "Config File:" bound to `{Binding ConfigFilePath}`
  - ".NET Runtime:" bound to `{Binding DotNetRuntime}`
- All fields should be non-editable with visual cues (gray background, no border, or flat styling)
- Consider adding a small "copy to clipboard" button next to each field for user convenience

### How to verify
- [x] Four read-only fields render in the Application Info section (AC-01, AC-02, AC-03, AC-04)
- [x] Fields are visually distinct as non-editable (AC-05)

---

## TASK-007-13-02: Add AppVersion, DatabasePath, ConfigFilePath, and DotNetRuntime read-only properties to ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-13-03, TASK-007-13-04, TASK-007-13-05, TASK-007-13-06 |

### What to do
- Add `[ObservableProperty] string _appVersion` (default empty)
- Add `[ObservableProperty] string _databasePath` (default empty)
- Add `[ObservableProperty] string _configFilePath` (default empty)
- Add `[ObservableProperty] string _dotNetRuntime` (default empty)
- These properties are set once during initialization and never modified by the user
- They should NOT participate in `HasUnsavedChanges` tracking (they are read-only display values)

### How to verify
- [x] All four properties are bindable and accessible (AC-01, AC-02, AC-03, AC-04)
- [x] Changing these properties does not trigger `HasUnsavedChanges` (AC-05)

---

## TASK-007-13-03: Populate AppVersion from assembly version or application metadata

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-13-02 |
| Blocks | None |

### What to do
- In the ViewModel constructor or `InitializeAsync()`, set `AppVersion`:
  - Primary: `Assembly.GetExecutingAssembly().GetName().Version?.ToString()` (returns "6.0.0.0")
  - Alternative: read from `ConfigurationLoader.GetSettings().Application.Version` (returns "6.0.0" as defined in `AppSettings.ApplicationInfo`)
  - Format to show only major.minor.patch: `version.ToString(3)` or use the `AppSettings` value directly
- Fall back to "Unknown" if both methods fail

### How to verify
- [x] `AppVersion` displays the correct version number (e.g., "6.0.0") (AC-01, AC-07)
- [x] Version is sourced from assembly metadata or AppSettings (AC-07)

---

## TASK-007-13-04: Populate DatabasePath from AppDbContext connection string

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-13-02 |
| Blocks | None |

### What to do
- In the ViewModel constructor or `InitializeAsync()`, extract the database path:
  - Get the connection string from `_dbContext.Database.GetConnectionString()` (returns e.g., `"Data Source=database.db"`)
  - Parse out the `Data Source=` value using string manipulation or `SqliteConnectionStringBuilder`
  - Resolve to an absolute path using `Path.GetFullPath()` if it's relative
- Set `DatabasePath` to the resolved absolute path (e.g., `C:\Users\...\database.db`)
- Fall back to "Unknown" if the connection string cannot be parsed

### How to verify
- [x] `DatabasePath` displays the resolved absolute path to the SQLite database file (AC-02, AC-08)
- [x] Path is sourced from the EF Core connection string (AC-08)

---

## TASK-007-13-05: Populate ConfigFilePath from ConfigurationLoader or AppDomain.CurrentDomain.BaseDirectory

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-13-02 |
| Blocks | None |

### What to do
- In the ViewModel constructor or `InitializeAsync()`, resolve the config file path:
  - `ConfigFilePath = Path.Combine(AppContext.BaseDirectory, "appsettings.json")`
  - Verify the file exists with `File.Exists()` and add a note if it doesn't (e.g., "appsettings.json (not found)")
  - Alternatively, if `ConfigurationLoader` exposes a `ConfigFilePath` property, use that
- Set `ConfigFilePath` to the resolved absolute path

### How to verify
- [x] `ConfigFilePath` displays the absolute path to `appsettings.json` (AC-03)
- [x] Path is resolved from the application base directory (AC-03)

---

## TASK-007-13-06: Populate DotNetRuntime from RuntimeInformation.FrameworkDescription

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-13-02 |
| Blocks | None |

### What to do
- In the ViewModel constructor or `InitializeAsync()`, set:
  - `DotNetRuntime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription`
  - This returns a string like ".NET 8.0.0" or ".NET 8.0.4"
- No fallback needed as this API is always available in .NET 6+

### How to verify
- [x] `DotNetRuntime` displays the .NET runtime version (e.g., ".NET 8.0") (AC-04)
- [x] Value is populated automatically without user action (AC-06)

---

## TASK-007-13-07: Style Application Info fields as read-only with distinct visual treatment

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | TASK-007-13-01 |
| Blocks | None |

### What to do
- Apply a consistent read-only style to all Application Info `TextBox` controls:
  - `Background="{StaticResource SurfaceBrush}"` or a lighter gray
  - `BorderThickness="0"` or `BorderBrush="Transparent"` for a flat appearance
  - `IsReadOnly="True"` and `IsTabStop="False"` to prevent focus
  - `Cursor="Arrow"` instead of the default text cursor
- Optionally add a `SelectAll` behavior on click so users can easily copy values
- Add small copy-to-clipboard buttons using `Clipboard.SetText()` bound to copy commands
- Ensure the section has a clear visual separation from editable settings above

### How to verify
- [x] Application Info fields have distinct read-only styling (gray, no border) (AC-05)
- [x] Fields cannot be edited by typing or pasting (AC-05)
- [x] Values can be selected and copied (AC-05)

---

## Dependency Graph
```
TASK-007-13-01 (XAML Layout)
       │
       └──▶ TASK-007-13-07 (Read-Only Styling)

TASK-007-13-02 (ViewModel Properties)
       │
       ├──▶ TASK-007-13-03 (App Version)
       ├──▶ TASK-007-13-04 (Database Path)
       ├──▶ TASK-007-13-05 (Config File Path)
       └──▶ TASK-007-13-06 (.NET Runtime)

Tasks 01 and 02 can start in parallel.
Tasks 03-06 are all independent of each other and can proceed in parallel once 02 is complete.
Task 07 depends on 01 (XAML must exist to apply styles).
```
