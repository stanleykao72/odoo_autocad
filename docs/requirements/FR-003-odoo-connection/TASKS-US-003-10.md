# TASKS: US-003-10 — Configure via appsettings.json

> **Parent US**: [US-003-10](US-003-10-configure-via-appsettings.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 6S + 0M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] None (this US is a prerequisite for US-003-01)

## Acceptance Criteria
- [ ] AC-01: The `appsettings.json` file supports an `Odoo` section with keys: `ServerUrl`, `Database`, `Username`, `TimeoutSeconds`
- [ ] AC-02: On application launch, the credential form fields are pre-populated with values from `appsettings.json`
- [ ] AC-03: The `TimeoutSeconds` value configures the `HttpClient.Timeout` (default: 30 seconds if not specified)
- [ ] AC-04: `TimeoutSeconds` is validated as a positive integer between 5 and 120 seconds
- [ ] AC-05: If `appsettings.json` is missing the `Odoo` section, the application starts with empty fields and default timeout
- [ ] AC-06: The Password field is intentionally excluded from `appsettings.json` for security
- [ ] AC-07: Users can override pre-configured values by typing in the UI fields at runtime
- [ ] AC-08: Configuration changes in the UI are not written back to `appsettings.json` (read-only configuration)

---

## TASK-003-10-01: Add Odoo section to appsettings.json

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/appsettings.json` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-10-03 |

### What to do
- Add or verify the `Odoo` section in `appsettings.json`:
  ```json
  {
      "Odoo": {
          "ServerUrl": "https://odoo.example.com",
          "Database": "production",
          "Username": "",
          "TimeoutSeconds": 30
      }
  }
  ```
- Do NOT include a `Password` key in the configuration file (security requirement per AC-06)
- Ensure the file is included in the build output: verify `.csproj` includes `<Content Include="appsettings.json"><CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory></Content>`
- Add `appsettings.Development.json` for development overrides (optional, with environment-specific URLs)

### How to verify
- [ ] appsettings.json contains Odoo section with ServerUrl, Database, Username, TimeoutSeconds (AC-01)
- [ ] Password is NOT present in appsettings.json (AC-06)

---

## TASK-003-10-02: Verify OdooSettings POCO class for strongly-typed binding

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Configuration/ConfigurationLoader.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-10-03 |

### What to do
- Verify the existing `OdooSettings` class in `ConfigurationLoader.cs`:
  ```csharp
  public class OdooSettings
  {
      public string ServerUrl { get; set; } = string.Empty;
      public string Database { get; set; } = string.Empty;
      public string Username { get; set; } = string.Empty;
      public int TimeoutSeconds { get; set; } = 30;
  }
  ```
- The class already exists and matches the required structure
- Ensure there is NO `Password` property in `OdooSettings` (confirms AC-06)
- Register for DI in `App.xaml.cs` or startup:
  ```csharp
  services.Configure<OdooSettings>(configuration.GetSection("Odoo"));
  // or
  services.AddSingleton(configuration.GetSection("Odoo").Get<OdooSettings>() ?? new OdooSettings());
  ```
- Verify `AppSettings.Odoo` property is of type `OdooSettings` (already present)

### How to verify
- [ ] OdooSettings POCO has ServerUrl, Database, Username, TimeoutSeconds properties (AC-01)
- [ ] No Password property exists in OdooSettings (AC-06)

---

## TASK-003-10-03: Inject IConfiguration and bind Odoo section in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-10-01, TASK-003-10-02 |
| Blocks | TASK-003-10-04 |

### What to do
- Inject `IOptions<OdooSettings>` (or `IConfiguration`) in the `OdooConnectionViewModel` constructor:
  ```csharp
  public OdooConnectionViewModel(
      IOdooService odooService,
      IOptions<OdooSettings> odooSettings,
      ILogger<OdooConnectionViewModel> logger)
  {
      _odooService = odooService;
      _odooSettings = odooSettings.Value;
      _logger = logger;
      LoadConfigurationDefaults();
  }
  ```
- Using `IOptions<OdooSettings>` provides strongly-typed access without manual key lookups
- If `IOptions<OdooSettings>` is not registered, fall back to `IConfiguration` with manual key access
- Store the original settings reference but do NOT write changes back (read-only per AC-08)

### How to verify
- [ ] OdooSettings is injected via DI (AC-01)
- [ ] Configuration is read-only from the ViewModel (AC-08)

---

## TASK-003-10-04: Pre-populate credential fields from configuration

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-10-03 |
| Blocks | None |

### What to do
- In a `LoadConfigurationDefaults()` method called from the constructor:
  ```csharp
  private void LoadConfigurationDefaults()
  {
      ServerUrl = _odooSettings.ServerUrl;
      Database = _odooSettings.Database;
      Username = _odooSettings.Username;
      TimeoutSeconds = _odooSettings.TimeoutSeconds;
      // Password intentionally NOT set from config
  }
  ```
- Fields are pre-populated but remain editable (the UI `TextBox` bindings allow user overrides per AC-07)
- If `OdooSettings` is `null` or has empty values, fields remain at their default empty values (AC-05)
- Do NOT attempt to write user changes back to `appsettings.json` -- UI values override config values only for the current session

### How to verify
- [ ] Credential fields are pre-populated on launch (AC-02)
- [ ] Users can override values by typing (AC-07)
- [ ] Missing Odoo section results in empty fields and default timeout (AC-05)

---

## TASK-003-10-05: Configure HttpClient.Timeout from TimeoutSeconds

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Modify the `OdooService` constructor (or `ConnectAsync`) to accept and apply a timeout value:
  ```csharp
  public OdooService(ILogger<OdooService>? logger = null, int timeoutSeconds = 30)
  {
      _logger = logger;
      _httpClient = new HttpClient();
      _httpClient.Timeout = TimeSpan.FromSeconds(Math.Clamp(timeoutSeconds, 5, 120));
      _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
  }
  ```
- Validate `timeoutSeconds` is between 5 and 120 (VR-003-007); clamp to range if out of bounds
- Default to 30 seconds if not specified or if the value is invalid
- Alternatively, accept the timeout per-request via `CancellationTokenSource` with timeout:
  ```csharp
  using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(_timeoutSeconds));
  var response = await _httpClient.PostAsync(url, content, cts.Token);
  ```
- Log the configured timeout value at startup

### How to verify
- [ ] HttpClient.Timeout is configured from TimeoutSeconds value (AC-03)
- [ ] TimeoutSeconds is validated as 5-120 range with default 30 (AC-04)

---

## TASK-003-10-06: Handle missing Odoo section gracefully

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `LoadConfigurationDefaults()`, handle null or missing configuration gracefully:
  ```csharp
  private void LoadConfigurationDefaults()
  {
      if (_odooSettings == null)
      {
          _logger?.LogWarning("Odoo configuration section not found in appsettings.json. Starting with empty defaults.");
          TimeoutSeconds = 30;
          return;
      }
      ServerUrl = _odooSettings.ServerUrl ?? string.Empty;
      Database = _odooSettings.Database ?? string.Empty;
      Username = _odooSettings.Username ?? string.Empty;
      TimeoutSeconds = _odooSettings.TimeoutSeconds > 0 ? _odooSettings.TimeoutSeconds : 30;
  }
  ```
- If `IOptions<OdooSettings>.Value` returns a default `OdooSettings` instance (all defaults), this is handled automatically by the empty string defaults in the POCO
- Test scenario: delete the `Odoo` section from `appsettings.json` and verify the application starts without errors
- Log a warning when the section is missing to help administrators troubleshoot

### How to verify
- [ ] Application starts with empty fields when Odoo section is missing (AC-05)
- [ ] Default timeout of 30 seconds is applied (AC-05)

---

## Dependency Graph
```
TASK-003-10-01 (appsettings.json Odoo section)
    |
TASK-003-10-02 (OdooSettings POCO)
    |
    +---> TASK-003-10-03 (Inject IConfiguration/IOptions)
              |
              +---> TASK-003-10-04 (Pre-populate fields)

TASK-003-10-05 (HttpClient.Timeout configuration)

TASK-003-10-06 (Graceful missing section handling)
```
