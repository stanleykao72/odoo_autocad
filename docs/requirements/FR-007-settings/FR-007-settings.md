# FR-007: Settings Page

> **Document Version**: 1.0
> **Last Updated**: 2026-02-06
> **Status**: Not Started
> **Priority**: P2

## 1. Overview

The Settings Page provides centralized configuration management for the application. It consolidates Odoo connection settings, AutoCAD integration parameters, MCP server configuration, UI appearance preferences, and advanced data management operations into a single, organized interface with tabbed navigation.

The Python application uses a dual configuration system: YAML files (`config/server.yaml`, `config/token.yaml`) for environment-specific settings and a SQLite database (`models/server.py` with SQLAlchemy ORM) for runtime caching and user preferences. The C# v6.0 migration replaces YAML with `appsettings.json` (via `Microsoft.Extensions.Configuration`) and uses Entity Framework Core with SQLite for persistent storage through the `AppDbContext` (containing `ServerConfigs`, `UserPreferences`, and related DbSets). The Settings Page is the primary UI for reading, editing, validating, and persisting these configuration values.

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-007-01 | CAD Engineer | Configure the Odoo server URL, database, and credentials | I can connect to the correct Odoo instance for my project |
| US-007-02 | CAD Engineer | Switch between different environment configurations (staging, production) | I can work against the right server without manual editing |
| US-007-03 | CAD Engineer | Test the Odoo connection from the settings page | I can verify my settings are correct before starting work |
| US-007-04 | CAD Engineer | Adjust AutoCAD connection timeout and retry attempts | I can handle slower AutoCAD startup on my machine |
| US-007-05 | System Admin | Configure the MCP SSE server port number | I can avoid port conflicts with other services on the machine |
| US-007-06 | System Admin | Enable or disable MCP server auto-start | I can control whether the AI assistant launches automatically |
| US-007-07 | CAD Engineer | Change the application theme (light, dark, system) | The UI matches my preference and reduces eye strain |
| US-007-08 | CAD Engineer | Set the application language/locale | I can use the application in my preferred language |
| US-007-09 | System Admin | Change the logging level for troubleshooting | I can increase log verbosity when diagnosing issues |
| US-007-10 | System Admin | Clear the local cache and synchronization data | I can reset the application state when data becomes stale |
| US-007-11 | System Admin | Export the current configuration to a file | I can back up or share settings across machines |
| US-007-12 | System Admin | Import a configuration from a file | I can restore settings on a new installation |
| US-007-13 | CAD Engineer | See the current application version and database path | I can report environment details when seeking support |

## 3. Python Reference

### Source Files
- `config/server.yaml` - Primary server configuration (host, db_name, url, token_file, StripMtext_file, read_csv_file, contract_product_file)
- `config/token.yaml` - User authentication token and server_file reference
- `config/server_khs.yaml`, `config/server_prod.yaml`, `config/server_prod_khs.yaml` - Environment-specific YAML configs
- `models/server.py` - SQLAlchemy `Server` model (id, host, db_name, url, token, sync_yaml, active)
- `ui/ui_theme.py` - UITheme class with COLORS, FONTS, SIZES, ICONS dictionaries; `setup_theme(appearance_mode, color_theme)`
- `ui/ui_fonts.py` - FontManager class with cross-platform font fallback (Windows: Microsoft JhengHei UI)
- `odoo.py` - Entry point, initializes database and loads YAML config on startup
- `utility/util_mcp_sse_manager.py` - MCPSSEManager with configurable port (default 8084)

### Key Functions
- `LoadYamlConfig(yaml_path)` - Loads YAML configuration files into Python dicts
- `UITheme.setup_theme(appearance_mode="system", color_theme="blue")` - Configures CustomTkinter theme
- `FontManager.get_best_font_family(font_type)` - Selects best available font with platform-specific fallback
- `Server` ORM model - Caches server connection info in SQLite (`db/database.db` or `%APPDATA%/OdooAutoCAD/database.db`)

### Python Configuration Pattern
```python
# server.yaml structure
server:
  host: 'odoo-server.example.com'
  db_name: 'odoo-database'
  url: 'https://odoo-server.example.com/api/v1/boq_import_api/swagger.json?token=...&db=...'
  token_file: 'c:\odoo\config\token.yaml'
  StripMtext_file: 'C:\\odoo\\...\\StripMtext v5-0b.lsp'
  read_csv_file: 'C:\\odoo\\...\\read_csv.lsp'
  contract_product_file: 'C:\\odoo\\...\\contract_product.lsp'

# token.yaml structure
user:
  token: '<uuid>'
  server_file: 'path/to/server_prod.yaml'

# SQLAlchemy model
class Server(Base):
    __tablename__ = 'server'
    id = Column(Integer, primary_key=True)
    host = Column(String)
    db_name = Column(String)
    url = Column(String)
    token = Column(String)
    sync_yaml = Column(Boolean, default=False)
    active = Column(Boolean, default=True)
```

### Python Theme Configuration
- Appearance modes: `"system"`, `"light"`, `"dark"`
- Color themes: `"blue"`, `"green"`, `"dark-blue"`
- Primary color: `#2B579A` (Professional Blue)
- Font: `Microsoft JhengHei UI` (with fallback chain per platform)
- Font sizes: title=16, heading=14, body=11, small=9, button=10

## 4. Functional Requirements

### Connection Settings (Odoo)

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-007-001 | Settings Page SHALL provide input fields for Odoo Server URL, Database Name, and Username | Must |
| FR-007-002 | Settings Page SHALL provide a secure input field for Odoo API Token (masked display, show/hide toggle) | Must |
| FR-007-003 | Settings Page SHALL provide a "Test Connection" button that validates the Odoo server URL and credentials by calling `IOdooService.TestConnectionAsync()` | Must |
| FR-007-004 | Test Connection SHALL display success or failure result inline with specific error details (timeout, auth failure, network error) | Must |
| FR-007-005 | Settings Page SHALL provide a connection timeout setting (in seconds, default 30) | Should |
| FR-007-006 | Settings Page SHALL persist all Odoo connection settings to both `appsettings.json` and the `ServerConfigs` database table | Must |

### Connection Settings (AutoCAD)

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-007-007 | Settings Page SHALL display the AutoCAD COM ProgId (default: `AutoCAD.Application`, read-only unless advanced mode) | Should |
| FR-007-008 | Settings Page SHALL provide a configurable connection timeout in seconds (default: 10) | Should |
| FR-007-009 | Settings Page SHALL provide a configurable retry attempt count (default: 3, range 1-10) | Should |

### MCP Server Settings

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-007-010 | Settings Page SHALL provide an input field for MCP SSE Server port number (default: 8084) | Must |
| FR-007-011 | Settings Page SHALL provide a toggle for MCP server auto-start on application launch (default: off) | Must |
| FR-007-012 | Settings Page SHALL provide a configurable heartbeat interval in seconds (default: 30) | Could |

### Appearance and Theme

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-007-013 | Settings Page SHALL provide a theme selector with options: System, Light, Dark | Must |
| FR-007-014 | Theme changes SHALL be applied immediately as a live preview without requiring application restart | Should |
| FR-007-015 | Settings Page SHALL provide a language/locale selector (English, Traditional Chinese) | Should |
| FR-007-016 | Language changes SHALL take effect after application restart, with a notification informing the user | Should |

### Logging and Diagnostics

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-007-017 | Settings Page SHALL provide a log level selector (Verbose, Debug, Information, Warning, Error, Fatal) with default "Information" | Should |
| FR-007-018 | Settings Page SHALL display the current log file path and provide an "Open Log Folder" button | Should |

### Data Management

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-007-019 | Settings Page SHALL provide a "Clear Cache" button that removes all BOQCache and ProductMapping records from the local database | Must |
| FR-007-020 | Clear Cache SHALL require confirmation dialog before executing and display affected record counts | Must |
| FR-007-021 | Settings Page SHALL provide an "Export Configuration" button that saves current settings to a user-chosen JSON file | Should |
| FR-007-022 | Settings Page SHALL provide an "Import Configuration" button that loads settings from a user-chosen JSON file and applies them | Should |

### Environment and General

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-007-023 | Settings Page SHALL provide an environment selector for switching between pre-defined configurations (Development, Staging, Production) | Should |
| FR-007-024 | Switching environments SHALL load the corresponding settings and prompt the user to confirm before overwriting current values | Should |
| FR-007-025 | Settings Page SHALL display read-only application info: Version, Database Path, Config File Path, .NET Runtime version | Must |
| FR-007-026 | Settings Page SHALL provide a "Save" button that persists all modified settings and a "Reset to Defaults" button that restores factory defaults | Must |
| FR-007-027 | Unsaved changes SHALL be tracked, and navigating away from the Settings Page with unsaved changes SHALL prompt a confirmation dialog | Should |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table with SyncType "settings_change" | Should |

## 5. UI Wireframe Description

```
+--------------------------------------------------------------+
|                       Settings Page                          |
+--------------------------------------------------------------+
|                                                               |
|  [ Connection ]  [ AutoCAD ]  [ MCP ]  [ Appearance ]  [ Advanced ] |
|  ============                                                 |
|                                                               |
|  ┌─ Connection Tab ─────────────────────────────────────────┐ |
|  │                                                           │ |
|  │  Odoo Server Connection                                   │ |
|  │  ┌─────────────────────────────────────────────────────┐  │ |
|  │  │ Server URL:  [https://odoo.example.com           ]  │  │ |
|  │  │ Database:    [production                         ]  │  │ |
|  │  │ Username:    [admin                              ]  │  │ |
|  │  │ API Token:   [********************************  👁]  │  │ |
|  │  │ Timeout(s):  [30     ]                              │  │ |
|  │  │                                                     │  │ |
|  │  │ [Test Connection]  ● Connected successfully         │  │ |
|  │  └─────────────────────────────────────────────────────┘  │ |
|  │                                                           │ |
|  │  Environment                                              │ |
|  │  ┌─────────────────────────────────────────────────────┐  │ |
|  │  │ Active Environment: [ Production        ▼]          │  │ |
|  │  │ Available: Development | Staging | Production       │  │ |
|  │  └─────────────────────────────────────────────────────┘  │ |
|  │                                                           │ |
|  └───────────────────────────────────────────────────────────┘ |
|                                                               |
|  ┌─ AutoCAD Tab ────────────────────────────────────────────┐ |
|  │  COM ProgId:         [AutoCAD.Application] (read-only)    │ |
|  │  Connection Timeout: [10    ] seconds                     │ |
|  │  Retry Attempts:     [3     ] (1-10)                      │ |
|  └───────────────────────────────────────────────────────────┘ |
|                                                               |
|  ┌─ MCP Tab ────────────────────────────────────────────────┐ |
|  │  SSE Server Port:    [8084  ]                             │ |
|  │  Auto-Start:         [x] Enable auto-start on launch      │ |
|  │  Heartbeat Interval: [30    ] seconds                     │ |
|  └───────────────────────────────────────────────────────────┘ |
|                                                               |
|  ┌─ Appearance Tab ─────────────────────────────────────────┐ |
|  │  Theme:    ( ) System  (o) Light  ( ) Dark                │ |
|  │  Language: [ Traditional Chinese  ▼]                      │ |
|  │            * Language change requires restart              │ |
|  └───────────────────────────────────────────────────────────┘ |
|                                                               |
|  ┌─ Advanced Tab ───────────────────────────────────────────┐ |
|  │  Logging                                                  │ |
|  │  Log Level:  [ Information  ▼]                            │ |
|  │  Log Path:   C:\Users\...\logs\    [Open Folder]          │ |
|  │                                                           │ |
|  │  Data Management                                          │ |
|  │  [Clear Cache]  [Export Config]  [Import Config]          │ |
|  │                                                           │ |
|  │  Application Info                                         │ |
|  │  Version:      6.0.0                                      │ |
|  │  Database:     C:\Users\...\database.db                   │ |
|  │  Config File:  appsettings.json                           │ |
|  │  .NET Runtime: .NET 8.0                                   │ |
|  │                                                           │ |
|  │  [Reset to Defaults]                                      │ |
|  └───────────────────────────────────────────────────────────┘ |
|                                                               |
|  ┌──────────────────────────────────────────────────────────┐ |
|  │     [Save Settings]                [Cancel / Revert]      │ |
|  └──────────────────────────────────────────────────────────┘ |
+--------------------------------------------------------------+
```

### Layout Details
- **Tab Control**: WPF `TabControl` with 5 tabs across the top (Connection, AutoCAD, MCP, Appearance, Advanced)
- **Connection Tab**: Primary tab, shown by default. Contains Odoo server fields and environment selector.
- **AutoCAD Tab**: AutoCAD COM configuration fields with numeric constraints.
- **MCP Tab**: MCP SSE server port, auto-start toggle, and heartbeat interval.
- **Appearance Tab**: Theme radio buttons with live preview, language dropdown.
- **Advanced Tab**: Logging settings, data management buttons, read-only application info panel.
- **Footer Bar**: Persistent Save/Cancel buttons visible regardless of active tab. Save button shows unsaved indicator (e.g., bold text or asterisk) when changes exist.

## 6. Data Model

### ViewModel

```csharp
public class SettingsViewModel : ObservableObject
{
    // Connection Settings (Odoo)
    public string OdooServerUrl { get; set; }
    public string OdooDatabaseName { get; set; }
    public string OdooUsername { get; set; }
    public string OdooApiToken { get; set; }
    public bool IsTokenVisible { get; set; }
    public int OdooTimeoutSeconds { get; set; }

    // Connection Test State
    public bool IsTestingConnection { get; set; }
    public bool? ConnectionTestResult { get; set; }     // null=not tested, true=success, false=fail
    public string ConnectionTestMessage { get; set; }

    // Environment
    public string SelectedEnvironment { get; set; }     // "Development", "Staging", "Production"
    public ObservableCollection<string> AvailableEnvironments { get; set; }

    // AutoCAD Settings
    public string AutoCADProgId { get; set; }
    public int AutoCADConnectionTimeout { get; set; }
    public int AutoCADRetryAttempts { get; set; }

    // MCP Settings
    public int MCPPort { get; set; }
    public bool MCPAutoStart { get; set; }
    public int MCPHeartbeatInterval { get; set; }

    // Appearance
    public string SelectedTheme { get; set; }           // "System", "Light", "Dark"
    public string SelectedLanguage { get; set; }        // "en-US", "zh-TW"

    // Logging
    public string SelectedLogLevel { get; set; }        // "Verbose", "Debug", "Information", etc.
    public string LogFilePath { get; set; }             // Read-only display

    // Application Info (read-only)
    public string AppVersion { get; set; }
    public string DatabasePath { get; set; }
    public string ConfigFilePath { get; set; }
    public string DotNetRuntime { get; set; }

    // State Tracking
    public bool HasUnsavedChanges { get; set; }
    public bool IsSaving { get; set; }

    // Commands
    public IAsyncRelayCommand TestConnectionCommand { get; }
    public IRelayCommand ToggleTokenVisibilityCommand { get; }
    public IAsyncRelayCommand SaveSettingsCommand { get; }
    public IRelayCommand CancelCommand { get; }
    public IRelayCommand ResetToDefaultsCommand { get; }
    public IAsyncRelayCommand ClearCacheCommand { get; }
    public IAsyncRelayCommand ExportConfigCommand { get; }
    public IAsyncRelayCommand ImportConfigCommand { get; }
    public IRelayCommand OpenLogFolderCommand { get; }
    public IRelayCommand<string> ChangeEnvironmentCommand { get; }
}
```

### Database Entities

```csharp
// ServerConfig entity (from OdooAutoCAD.Data/Entities/Entities.cs)
// Key-value store for server connection settings
public class ServerConfig
{
    public int Id { get; set; }
    public string ConfigKey { get; set; }       // Unique, max 100 chars
    public string? ConfigValue { get; set; }    // Max 1000 chars
    public string? Description { get; set; }
    public DateTime CreatedAt { get; set; }
    public DateTime UpdatedAt { get; set; }
}

// UserPreference entity (from OdooAutoCAD.Data/Entities/Entities.cs)
// Key-value store for UI preferences and app settings
public class UserPreference
{
    public int Id { get; set; }
    public string PreferenceKey { get; set; }   // Unique, max 100 chars
    public string? PreferenceValue { get; set; }
    public string? DataType { get; set; }       // "string", "int", "bool", "json"
    public DateTime UpdatedAt { get; set; }
}
```

### ConfigKey Mapping

| ConfigKey | Source (Python) | Example Value | Storage |
|-----------|----------------|---------------|---------|
| `odoo.server_url` | `server.yaml` > `host` | `https://odoo.example.com` | ServerConfig |
| `odoo.database` | `server.yaml` > `db_name` | `production` | ServerConfig |
| `odoo.username` | N/A (new in C#) | `admin` | ServerConfig |
| `odoo.api_token` | `token.yaml` > `token` | `6d4bead3-...` | ServerConfig |
| `odoo.timeout_seconds` | N/A (new in C#) | `30` | ServerConfig |
| `autocad.prog_id` | Hardcoded | `AutoCAD.Application` | ServerConfig |
| `autocad.connection_timeout` | N/A (new in C#) | `10` | ServerConfig |
| `autocad.retry_attempts` | N/A (new in C#) | `3` | ServerConfig |
| `mcp.port` | `--port 8084` | `8084` | ServerConfig |
| `mcp.auto_start` | `--enable-mcp` flag | `false` | ServerConfig |
| `mcp.heartbeat_interval` | N/A (new in C#) | `30` | ServerConfig |
| `appearance.theme` | `UITheme.setup_theme("system")` | `System` | UserPreference |
| `appearance.language` | Hardcoded zh-TW | `zh-TW` | UserPreference |
| `logging.level` | Console logging | `Information` | UserPreference |
| `app.environment` | Multiple YAML files | `Production` | UserPreference |

## 7. API/Service Dependencies

| Service | Interface | Methods Used |
|---------|-----------|-------------|
| Database Context | `AppDbContext` | `ServerConfigs` DbSet for reading/writing connection config; `UserPreferences` DbSet for reading/writing UI preferences; `BOQCache.RemoveRange()` for cache clearing; `SyncLogs.Add()` for audit logging |
| Configuration Loader | `ConfigurationLoader` | `LoadFromJson()` to read `appsettings.json`; `GetSettings()` for current `AppSettings`; `GetAvailableConfigs()` to list environment configs |
| Odoo Service | `IOdooService` | `TestConnectionAsync(url, database, username, token)` to validate connection settings from the settings page |
| Navigation Service | `INavigationService` | `CanNavigateFrom()` to intercept navigation when unsaved changes exist |
| File Dialog Service | `IFileDialogService` | `ShowSaveFileDialog()` for config export; `ShowOpenFileDialog()` for config import |
| Theme Service | `IThemeService` | `ApplyTheme(themeName)` for live theme switching |
| Logging Service | `ILogger<SettingsViewModel>` | Standard .NET logging for settings change auditing |

## 8. Validation Rules

| Rule | Field | Description |
|------|-------|-------------|
| VR-007-001 | Odoo Server URL | Must be a valid URL format starting with `http://` or `https://`. Validated with `Uri.TryCreate()` and `UriKind.Absolute`. |
| VR-007-002 | Odoo Server URL | Must not be empty or whitespace when saving connection settings. |
| VR-007-003 | Odoo Database Name | Must not be empty. Maximum 100 characters. Must contain only alphanumeric characters, hyphens, and underscores. |
| VR-007-004 | Odoo API Token | Must be a valid UUID format (regex: `^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-...`) or non-empty string. |
| VR-007-005 | Odoo Timeout | Must be an integer between 5 and 300 seconds (inclusive). |
| VR-007-006 | MCP Port | Must be an integer between 1024 and 65535 (inclusive). Must not conflict with well-known ports. |
| VR-007-007 | AutoCAD Retry Attempts | Must be an integer between 1 and 10 (inclusive). |
| VR-007-008 | AutoCAD Connection Timeout | Must be an integer between 5 and 120 seconds (inclusive). |
| VR-007-009 | MCP Heartbeat Interval | Must be an integer between 5 and 300 seconds (inclusive). |
| VR-007-010 | Theme Selection | Must be one of: `"System"`, `"Light"`, `"Dark"`. |
| VR-007-011 | Language Selection | Must be one of the supported locale codes: `"en-US"`, `"zh-TW"`. |
| VR-007-012 | Log Level Selection | Must be one of: `"Verbose"`, `"Debug"`, `"Information"`, `"Warning"`, `"Error"`, `"Fatal"`. |
| VR-007-013 | Import Config File | Imported JSON must be valid JSON and contain the expected top-level keys (`Application`, `Odoo`, `AutoCAD`, `MCP`). |
| VR-007-014 | All Required Fields | Save button SHALL be disabled when any required field fails validation. Individual fields SHALL display inline validation error messages. |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| Invalid URL format | "Please enter a valid URL starting with http:// or https://" | Inline validation error on the Server URL field; prevent save |
| Connection test timeout | "Connection test timed out after {timeout} seconds. Check the server URL and network." | Display error in test result area; enable retry |
| Connection test auth failure | "Authentication failed. Please verify your API token and username." | Display error in test result area; highlight token field |
| Connection test network error | "Cannot reach server at {url}. Check network connectivity and firewall settings." | Display error in test result area |
| Save to appsettings.json fails | "Failed to save configuration file. Check file permissions. Error: {details}" | Show error dialog; settings remain in memory but not persisted |
| Save to database fails | "Failed to save settings to database. Error: {details}" | Show error dialog with details; suggest checking database file permissions |
| Port already in use | "Port {port} appears to be in use. MCP server may not start. Choose a different port." | Show warning (not blocking); allow save with warning indicator |
| Clear cache confirmation | "This will remove {n} cached BOQ entries and {m} product mappings. This action cannot be undone. Continue?" | Confirmation dialog with record counts; proceed only on explicit confirm |
| Clear cache failure | "Failed to clear cache. Error: {details}" | Show error dialog |
| Import config invalid JSON | "The selected file is not a valid configuration file. Expected JSON format." | Show error dialog; do not apply any changes |
| Import config missing keys | "Configuration file is missing required sections: {missing_keys}. Import partially applied." | Show warning with details of what was skipped |
| Export config write failure | "Failed to export configuration to {path}. Check file permissions." | Show error dialog |
| Reset to defaults confirmation | "This will reset ALL settings to factory defaults. Your current connection settings will be lost. Continue?" | Confirmation dialog; proceed only on explicit confirm |
| Unsaved changes on navigate | "You have unsaved changes. Save before leaving?" | Three-button dialog: Save / Discard / Cancel |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/SettingsPage.xaml` - WPF page with `TabControl` and data-bound fields
- `ViewModels/SettingsViewModel.cs` - MVVM ViewModel with all settings properties and commands
- `OdooAutoCAD.Configuration/ConfigurationLoader.cs` (exists, 162 lines) - Extended for save/export operations
- `OdooAutoCAD.Data/Context/AppDbContext.cs` (exists, 92 lines) - Provides `ServerConfigs` and `UserPreferences` DbSets
- `OdooAutoCAD.Data/Entities/Entities.cs` (exists, 87 lines) - `ServerConfig` and `UserPreference` entity definitions

### Migration from Python YAML to C# appsettings.json

The Python dual-config system (YAML + SQLite) is consolidated in C# as follows:

| Python Source | C# Target | Notes |
|---------------|-----------|-------|
| `config/server.yaml` (`host`, `db_name`, `url`) | `appsettings.json` > `Odoo` section + `ServerConfigs` table | JSON file for defaults; database for runtime overrides |
| `config/token.yaml` (`token`, `server_file`) | `ServerConfigs` table (key: `odoo.api_token`) | Token stored in DB only, never in plain-text JSON on disk |
| Multiple YAML files (server_khs, server_prod, etc.) | `appsettings.{Environment}.json` pattern | .NET standard environment-based config layering |
| `models/server.py` Server model | `ServerConfig` entity via EF Core | Key-value pattern replaces fixed-column model |
| `UITheme.setup_theme(mode, theme)` | `UserPreference` entries + WPF `ResourceDictionary` switching | Theme stored as preference, applied via WPF theming system |
| `--enable-mcp` CLI flag | `appsettings.json` > `MCP.AutoStart` + `UserPreference` override | Persisted preference overrides JSON default |

### Settings Persistence Strategy

Settings are stored in two layers:
1. **`appsettings.json`** (file): Default/factory values. Read on startup. Modified only on explicit "Save" action or "Export". Located in the application directory.
2. **`ServerConfigs` / `UserPreferences` tables** (SQLite): Runtime overrides. Always take precedence over JSON file values. Read/written by EF Core through `AppDbContext`.

Load order on startup:
```
appsettings.json (base) → appsettings.{Environment}.json (override) → ServerConfigs DB (override) → UserPreferences DB (override)
```

### MVVM Bindings

```xml
<!-- Example: Odoo Server URL field with validation -->
<TextBox Text="{Binding OdooServerUrl, UpdateSourceTrigger=PropertyChanged, ValidatesOnNotifyDataErrors=True}"
         Watermark="https://odoo.example.com" />

<!-- Example: Theme radio buttons -->
<RadioButton Content="System" IsChecked="{Binding SelectedTheme, Converter={StaticResource StringMatchConverter}, ConverterParameter=System}" />
<RadioButton Content="Light"  IsChecked="{Binding SelectedTheme, Converter={StaticResource StringMatchConverter}, ConverterParameter=Light}" />
<RadioButton Content="Dark"   IsChecked="{Binding SelectedTheme, Converter={StaticResource StringMatchConverter}, ConverterParameter=Dark}" />

<!-- Example: Save button with unsaved indicator -->
<Button Content="{Binding HasUnsavedChanges, Converter={StaticResource UnsavedToLabelConverter}}"
        Command="{Binding SaveSettingsCommand}"
        IsEnabled="{Binding CanSave}" />

<!-- Example: Token field with show/hide toggle -->
<PasswordBox Password="{Binding OdooApiToken}"
             Visibility="{Binding IsTokenVisible, Converter={StaticResource BoolToVisibilityInverter}}" />
<TextBox Text="{Binding OdooApiToken}"
         Visibility="{Binding IsTokenVisible, Converter={StaticResource BoolToVisibilityConverter}}" />
<Button Content="Show/Hide" Command="{Binding ToggleTokenVisibilityCommand}" />
```

### EF Core UserPreference Usage Pattern

```csharp
// Reading a preference
public async Task<string> GetPreferenceAsync(string key, string defaultValue = "")
{
    var pref = await _dbContext.UserPreferences
        .FirstOrDefaultAsync(p => p.PreferenceKey == key);
    return pref?.PreferenceValue ?? defaultValue;
}

// Writing a preference (upsert pattern)
public async Task SetPreferenceAsync(string key, string value, string dataType = "string")
{
    var pref = await _dbContext.UserPreferences
        .FirstOrDefaultAsync(p => p.PreferenceKey == key);

    if (pref == null)
    {
        pref = new UserPreference
        {
            PreferenceKey = key,
            PreferenceValue = value,
            DataType = dataType,
            UpdatedAt = DateTime.UtcNow
        };
        _dbContext.UserPreferences.Add(pref);
    }
    else
    {
        pref.PreferenceValue = value;
        pref.UpdatedAt = DateTime.UtcNow;
    }

    await _dbContext.SaveChangesAsync();
}
```

### Special Considerations

- **Token Security**: The Odoo API token must never be written to `appsettings.json` in plain text. It is stored only in the `ServerConfigs` database table. The PasswordBox/TextBox toggle pattern provides user visibility control.
- **Live Theme Preview**: When the user selects a theme, `IThemeService.ApplyTheme()` immediately swaps the WPF `ResourceDictionary` (merging a Light or Dark theme dictionary) without requiring restart. The `SelectedTheme` property setter triggers this.
- **Unsaved Changes Tracking**: The ViewModel implements `INotifyDataErrorInfo` for validation and tracks property changes against a snapshot of the original values. `HasUnsavedChanges` is recomputed on every property change.
- **DI Registration**: `SettingsViewModel` should be registered as transient (new instance each time the page is navigated to) so it always loads fresh values from the database.
- **ConfigurationLoader Extension**: The existing `ConfigurationLoader.LoadFromJson()` reads settings; a new `SaveToJson(AppSettings)` method serializes current settings back to `appsettings.json` (excluding sensitive fields like token).
- **Clear Cache Scope**: "Clear Cache" removes all rows from `BOQCache` and `ProductMapping` tables, resets `SyncLog` entries, and invalidates any in-memory caches held by `IOdooService`.

### Differences from Python

| Aspect | Python | C# |
|--------|--------|-----|
| Config format | YAML files (`config/*.yaml`) | `appsettings.json` with environment overlays |
| Config loading | Custom `LoadYamlConfig()` | `Microsoft.Extensions.Configuration` + `ConfigurationLoader` |
| Config persistence | YAML write-back + SQLAlchemy | EF Core `ServerConfigs`/`UserPreferences` + JSON serialization |
| Environment switching | Select different YAML file path in `token.yaml` | `DOTNET_ENVIRONMENT` variable + `appsettings.{env}.json` + DB override |
| Theme system | CustomTkinter `set_appearance_mode()` | WPF `ResourceDictionary` merging with `IThemeService` |
| Token storage | Plain text in `token.yaml` | `ServerConfigs` database table (not in JSON file) |
| Settings UI | No dedicated settings page in Python (configured via YAML editing) | Full tabbed settings page with MVVM data binding |
| Validation | Manual/none on YAML values | WPF `INotifyDataErrorInfo` with inline validation messages |
