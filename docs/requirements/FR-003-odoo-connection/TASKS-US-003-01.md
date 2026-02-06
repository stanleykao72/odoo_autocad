# TASKS: US-003-01 — Enter Odoo Credentials

> **Parent US**: [US-003-01](US-003-01-enter-odoo-credentials.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 3S + 2M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-10 (appsettings.json configuration) must be completed so defaults can be loaded

## Acceptance Criteria
- [ ] AC-01: The page displays input fields for Server URL, Database Name, Username, and Password/API Key
- [ ] AC-02: The Password field uses masked input (PasswordBox) so characters are not visible
- [ ] AC-03: Server URL, Database, and Username fields are pre-populated from `appsettings.json` on launch
- [ ] AC-04: Server URL validates as a valid HTTPS URL pattern (`^https?://[^\s/]+`); HTTP triggers a warning but is not blocked
- [ ] AC-05: Server URL has trailing slashes stripped before use (matching `OdooService.ConnectAsync` behavior)
- [ ] AC-06: Database Name field is not empty and contains only alphanumeric characters, hyphens, and underscores
- [ ] AC-07: Username field is not empty when using session-based authentication
- [ ] AC-08: Password field is not empty with a minimum length of 1 character
- [ ] AC-09: "Connect" and "Test Connection" buttons are disabled when Server URL is empty
- [ ] AC-10: The Password/API Key is never stored in `appsettings.json` in plaintext

---

## TASK-003-01-01: Create Connection Settings GroupBox with credential input fields

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-003-01-05, TASK-003-01-06 |

### What to do
- Create `OdooConnectionPage.xaml` as a WPF `Page` in namespace `OdooAutoCAD.App.Views.Pages`
- Add a `GroupBox` with `Header="Connection Settings"` containing a `Grid` layout with 4 labeled input rows:
  - Row 0: `Label` "Server URL:" + `TextBox` bound to `{Binding ServerUrl, UpdateSourceTrigger=PropertyChanged}`
  - Row 1: `Label` "Database:" + `TextBox` bound to `{Binding Database, UpdateSourceTrigger=PropertyChanged}`
  - Row 2: `Label` "Username:" + `TextBox` bound to `{Binding Username, UpdateSourceTrigger=PropertyChanged}`
  - Row 3: `Label` "Password:" + `PasswordBox` (requires attached behavior or code-behind for binding since `PasswordBox.Password` is not a dependency property)
- Add a button bar with `StackPanel Orientation="Horizontal"` containing "Test Connection", "Connect", "Disconnect" buttons
- Set `DataContext` to `OdooConnectionViewModel` via DI or ViewModelLocator
- Use `Validation.ErrorTemplate` on each input field for inline error display
- Include `xmlns:vm` namespace reference for design-time `d:DataContext`

### How to verify
- [ ] Four labeled input fields are visible: Server URL, Database, Username, Password (AC-01)
- [ ] Password field uses `PasswordBox` with masked input (AC-02)

---

## TASK-003-01-02: Add ViewModel properties for credential fields with change notification

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-01-03, TASK-003-01-04 |

### What to do
- Create `OdooConnectionViewModel` class in namespace `OdooAutoCAD.App.ViewModels`, inheriting `ObservableObject` from CommunityToolkit.Mvvm
- Inject `IOdooService`, `IConfiguration`, `ILogger<OdooConnectionViewModel>` via constructor
- Add observable properties using `[ObservableProperty]` attribute:
  - `private string _serverUrl = string.Empty;`
  - `private string _database = string.Empty;`
  - `private string _username = string.Empty;`
  - `private string _password = string.Empty;`
  - `private int _timeoutSeconds = 30;`
- Implement `partial void OnServerUrlChanged(string value)` to strip trailing slashes via `value.TrimEnd('/')`
- Store the password separately; do not include it in any serialization or configuration write-back

### How to verify
- [ ] ServerUrl, Database, Username, Password properties exist with change notification (AC-01)
- [ ] ServerUrl has trailing slashes stripped on change (AC-05)

---

## TASK-003-01-03: Implement INotifyDataErrorInfo validation for credential fields

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-01-02 |
| Blocks | TASK-003-01-06 |

### What to do
- Implement `INotifyDataErrorInfo` interface on `OdooConnectionViewModel` (or use CommunityToolkit.Mvvm's `ObservableValidator` base class instead of `ObservableObject`)
- Add validation attributes and/or custom validation for each field:
  - `ServerUrl`: Required; must match regex `^https?://[^\s/]+`; if `http://` show warning (not error) via a separate `UrlWarningMessage` property
  - `Database`: Required; must match regex `^[a-zA-Z0-9_-]+$` (alphanumeric, hyphens, underscores only)
  - `Username`: Required; must not be empty or whitespace
  - `Password`: Required; minimum length 1 character
  - `TimeoutSeconds`: Range 5-120
- Add a `HasErrors` property and an `IsFormValid` computed property that returns `true` only when all required fields are valid
- Raise `ErrorsChanged` event when validation state changes
- Provide `GetErrors(string propertyName)` returning `IEnumerable<ValidationResult>` for per-field error display

### How to verify
- [ ] Server URL validates as HTTPS URL pattern; HTTP triggers warning but is not blocked (AC-04)
- [ ] Database Name rejects empty or invalid characters (AC-06)
- [ ] Username rejects empty value (AC-07)
- [ ] Password rejects empty value (AC-08)

---

## TASK-003-01-04: Load default values from IConfiguration on ViewModel initialization

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-01-02 |
| Blocks | None |

### What to do
- In the `OdooConnectionViewModel` constructor, read values from `IConfiguration`:
  - `ServerUrl = configuration["Odoo:ServerUrl"] ?? string.Empty;`
  - `Database = configuration["Odoo:Database"] ?? string.Empty;`
  - `Username = configuration["Odoo:Username"] ?? string.Empty;`
  - `TimeoutSeconds = int.TryParse(configuration["Odoo:TimeoutSeconds"], out var t) ? t : 30;`
- Alternatively, use `IOptions<OdooSettings>` pattern by binding `OdooSettings` POCO from `ConfigurationLoader.cs` (class already exists with `ServerUrl`, `Database`, `Username`, `TimeoutSeconds`)
- Do NOT load or set `Password` from configuration -- leave it empty for user entry
- If the `Odoo` section is missing, all fields remain at defaults (empty strings, timeout 30)

### How to verify
- [ ] Server URL, Database, Username are pre-populated from `appsettings.json` on launch (AC-03)
- [ ] Password field remains empty regardless of configuration content (AC-10)

---

## TASK-003-01-05: Create PasswordBox attached behavior for MVVM binding

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Helpers/PasswordBoxHelper.cs` |
| Estimate | M |
| Depends On | TASK-003-01-01 |
| Blocks | None |

### What to do
- Create a `PasswordBoxHelper` attached behavior class in namespace `OdooAutoCAD.App.Helpers`
- Implement an `AttachedPassword` attached `DependencyProperty` of type `string`:
  - Register with `DependencyProperty.RegisterAttached("AttachedPassword", typeof(string), typeof(PasswordBoxHelper), new PropertyMetadata(OnAttachedPasswordChanged))`
  - In the `OnAttachedPasswordChanged` callback, set `PasswordBox.Password` when the DP changes (if value differs)
  - Subscribe to `PasswordBox.PasswordChanged` event to update the DP when the user types
  - Use a reentrancy guard (`_isUpdating` flag) to prevent infinite loops between DP and PasswordChanged
- In `OdooConnectionPage.xaml`, use: `<PasswordBox helpers:PasswordBoxHelper.AttachedPassword="{Binding Password, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />`
- Ensure the Password value is only held in memory and never serialized to config or logs

### How to verify
- [ ] Password field uses PasswordBox with masked input and supports MVVM two-way binding (AC-02)
- [ ] Password value is never written to appsettings.json (AC-10)

---

## TASK-003-01-06: Bind Connect/Test Connection button IsEnabled to validation state

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | TASK-003-01-01, TASK-003-01-03 |
| Blocks | None |

### What to do
- Bind "Connect" button `IsEnabled` to `{Binding IsFormValid}` (computed property that is `true` when ServerUrl, Database, Username, Password all pass validation)
- Bind "Test Connection" button `IsEnabled` to the same `{Binding IsFormValid}` property
- Alternatively, use `IAsyncRelayCommand.CanExecute` by implementing `CanExecute` delegates on the relay commands that check `IsFormValid`
- When `ServerUrl` is empty, both buttons must be disabled regardless of other field states
- Add a `MultiBinding` with `IMultiValueConverter` if individual field bindings are preferred over a single `IsFormValid` property
- Ensure buttons re-evaluate enabled state when any credential field changes (via `NotifyCanExecuteChanged()` on each field's `OnChanged` partial method)

### How to verify
- [ ] "Connect" and "Test Connection" buttons are disabled when Server URL is empty (AC-09)
- [ ] Buttons are disabled when any required field fails validation (AC-04, AC-06, AC-07, AC-08)

---

## Dependency Graph
```
TASK-003-01-02 (ViewModel properties)
    |
    +---> TASK-003-01-03 (INotifyDataErrorInfo validation)
    |         |
    |         +---> TASK-003-01-06 (Button IsEnabled binding)
    |
    +---> TASK-003-01-04 (Load IConfiguration defaults)

TASK-003-01-01 (XAML page)
    |
    +---> TASK-003-01-05 (PasswordBox helper)
    +---> TASK-003-01-06 (Button IsEnabled binding)
```
