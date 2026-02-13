using System.IO;
using System.Runtime.InteropServices;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.Configuration;
using OdooAutoCAD.Core.AutoCAD;

namespace OdooAutoCAD.App.ViewModels;

/// <summary>
/// ViewModel for the Settings page.
/// Manages Odoo, AutoCAD, MCP settings and application info.
/// </summary>
public partial class SettingsViewModel : ObservableObject
{
    private readonly ISettingsService _settingsService;
    private readonly IAppLogService _logService;
    private readonly DrawingDataServiceDispatcher _dispatcher;
    private readonly ILogger<SettingsViewModel>? _logger;
    private bool _isLoading;

    // --- Connection (Odoo) ---

    [ObservableProperty]
    private string _odooSwaggerUrl = string.Empty;

    [ObservableProperty]
    private string _odooApiToken = string.Empty;

    [ObservableProperty]
    private bool _isTokenVisible;

    [ObservableProperty]
    private int _odooTimeoutSeconds = 30;

    // --- AutoCAD ---

    [ObservableProperty]
    private string _autoCADProgId = "AutoCAD.Application";

    [ObservableProperty]
    private int _autoCADConnectionTimeout = 10;

    [ObservableProperty]
    private int _autoCADRetryAttempts = 3;

    [ObservableProperty]
    private string _autoCADMode = "COM";

    /// <summary>
    /// True when COM mode is selected. Used for RadioButton binding.
    /// </summary>
    public bool IsComMode
    {
        get => AutoCADMode == "COM";
        set
        {
            if (value) AutoCADMode = "COM";
        }
    }

    /// <summary>
    /// True when File mode is selected. Used for RadioButton binding.
    /// </summary>
    public bool IsFileMode
    {
        get => AutoCADMode == "File";
        set
        {
            if (value) AutoCADMode = "File";
        }
    }

    partial void OnAutoCADModeChanged(string value)
    {
        OnPropertyChanged(nameof(IsComMode));
        OnPropertyChanged(nameof(IsFileMode));
    }

    // --- MCP ---

    [ObservableProperty]
    private int _mcpPort = 8084;

    [ObservableProperty]
    private bool _mcpAutoStart;

    [ObservableProperty]
    private int _mcpHeartbeatInterval = 30;

    // --- App Info (read-only display) ---

    [ObservableProperty]
    private string _appVersion = string.Empty;

    [ObservableProperty]
    private string _databasePath = string.Empty;

    [ObservableProperty]
    private string _configFilePath = string.Empty;

    [ObservableProperty]
    private string _dotNetRuntime = string.Empty;

    // --- State ---

    [ObservableProperty]
    private bool _hasUnsavedChanges;

    [ObservableProperty]
    private bool _isSaving;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    // Snapshot for cancel/dirty tracking
    private Dictionary<string, string?> _originalValues = new();

    public SettingsViewModel(
        ISettingsService settingsService,
        IAppLogService logService,
        DrawingDataServiceDispatcher dispatcher,
        ILogger<SettingsViewModel>? logger = null)
    {
        _settingsService = settingsService;
        _logService = logService;
        _dispatcher = dispatcher;
        _logger = logger;
    }

    [RelayCommand]
    private async Task LoadSettingsAsync()
    {
        _isLoading = true;
        try
        {
            StatusMessage = "Loading settings...";

            // Load from AppSettings (JSON/YAML)
            var appSettings = _settingsService.GetAppSettings();
            OdooSwaggerUrl = appSettings.Odoo.SwaggerUrl;
            OdooTimeoutSeconds = appSettings.Odoo.TimeoutSeconds;

            AutoCADProgId = appSettings.AutoCAD.ProgId;
            AutoCADConnectionTimeout = appSettings.AutoCAD.ConnectionTimeoutSeconds;
            AutoCADRetryAttempts = appSettings.AutoCAD.RetryAttempts;

            // Load AutoCAD mode from user preferences
            var savedMode = await _settingsService.GetPreferenceAsync("autocad_mode", "COM");
            AutoCADMode = savedMode == "File" ? "File" : "COM";

            McpPort = appSettings.MCP.Port;
            McpAutoStart = appSettings.MCP.AutoStart;
            McpHeartbeatInterval = appSettings.MCP.HeartbeatIntervalSeconds;

            AppVersion = appSettings.Application.Version;
            DatabasePath = appSettings.Database.ConnectionString;
            ConfigFilePath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            DotNetRuntime = RuntimeInformation.FrameworkDescription;

            // Load token from DB (sensitive — never in JSON)
            var configs = await _settingsService.LoadServerConfigsAsync();
            if (configs.TryGetValue("odoo_user_token", out var token))
                OdooApiToken = token ?? string.Empty;

            // Override swagger URL from DB if value exists there
            if (configs.TryGetValue("odoo_swagger_url", out var url) && !string.IsNullOrEmpty(url))
                OdooSwaggerUrl = url;

            TakeSnapshot();
            HasUnsavedChanges = false;
            StatusMessage = "Settings loaded";
            _logService.Log("Settings loaded", "Settings");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to load settings");
            _logService.Log($"Failed to load settings: {ex.Message}", "Settings", AppLogLevel.Error);
            StatusMessage = $"Failed to load settings: {ex.Message}";
        }
        finally
        {
            _isLoading = false;
        }
    }

    [RelayCommand]
    private void ToggleTokenVisibility()
    {
        IsTokenVisible = !IsTokenVisible;
    }

    [RelayCommand]
    private async Task SaveSettingsAsync()
    {
        IsSaving = true;
        StatusMessage = "Saving settings...";
        try
        {
            // Build AppSettings from current values
            var appSettings = _settingsService.GetAppSettings();
            appSettings.Odoo.SwaggerUrl = OdooSwaggerUrl;
            appSettings.Odoo.TimeoutSeconds = OdooTimeoutSeconds;

            appSettings.AutoCAD.ConnectionTimeoutSeconds = AutoCADConnectionTimeout;
            appSettings.AutoCAD.RetryAttempts = AutoCADRetryAttempts;

            appSettings.MCP.Port = McpPort;
            appSettings.MCP.AutoStart = McpAutoStart;
            appSettings.MCP.HeartbeatIntervalSeconds = McpHeartbeatInterval;

            // Save non-sensitive to JSON
            await _settingsService.SaveAppSettingsAsync(appSettings);

            // Save to DB (including token, which is sensitive)
            await _settingsService.SaveServerConfigsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = OdooSwaggerUrl,
                ["odoo_user_token"] = OdooApiToken,
                ["odoo_timeout"] = OdooTimeoutSeconds.ToString(),
                ["autocad_connection_timeout"] = AutoCADConnectionTimeout.ToString(),
                ["autocad_retry_attempts"] = AutoCADRetryAttempts.ToString(),
                ["mcp_port"] = McpPort.ToString(),
                ["mcp_auto_start"] = McpAutoStart.ToString(),
                ["mcp_heartbeat_interval"] = McpHeartbeatInterval.ToString()
            });

            // Save AutoCAD mode to user preferences
            await _settingsService.SetPreferenceAsync("autocad_mode", AutoCADMode);

            // Apply mode switch immediately at runtime
            var targetMode = AutoCADMode == "File"
                ? AutoCADOperationMode.File
                : AutoCADOperationMode.COM;
            await _dispatcher.SwitchModeAsync(targetMode);

            TakeSnapshot();
            HasUnsavedChanges = false;
            StatusMessage = "Settings saved successfully";
            _logService.Log("Settings saved successfully", "Settings");
            _logger?.LogInformation("Settings saved successfully");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to save settings");
            _logService.Log($"Failed to save settings: {ex.Message}", "Settings", AppLogLevel.Error);
            StatusMessage = $"Failed to save: {ex.Message}";
        }
        finally
        {
            IsSaving = false;
        }
    }

    [RelayCommand]
    private void CancelChanges()
    {
        RestoreSnapshot();
        HasUnsavedChanges = false;
        StatusMessage = "Changes cancelled";
    }

    /// <summary>
    /// Called by CommunityToolkit.Mvvm when any [ObservableProperty] changes.
    /// Tracks dirty state for all editable properties.
    /// </summary>
    protected override void OnPropertyChanged(System.ComponentModel.PropertyChangedEventArgs e)
    {
        base.OnPropertyChanged(e);

        // Skip dirty tracking for non-editable / state properties
        if (_isLoading) return;
        if (e.PropertyName is nameof(HasUnsavedChanges)
            or nameof(IsSaving)
            or nameof(StatusMessage)
            or nameof(AppVersion)
            or nameof(DatabasePath)
            or nameof(ConfigFilePath)
            or nameof(DotNetRuntime)
            or nameof(IsTokenVisible))
            return;

        HasUnsavedChanges = true;
    }

    private void TakeSnapshot()
    {
        _originalValues = new Dictionary<string, string?>
        {
            [nameof(OdooSwaggerUrl)] = OdooSwaggerUrl,
            [nameof(OdooApiToken)] = OdooApiToken,
            [nameof(OdooTimeoutSeconds)] = OdooTimeoutSeconds.ToString(),
            [nameof(AutoCADConnectionTimeout)] = AutoCADConnectionTimeout.ToString(),
            [nameof(AutoCADRetryAttempts)] = AutoCADRetryAttempts.ToString(),
            [nameof(AutoCADMode)] = AutoCADMode,
            [nameof(McpPort)] = McpPort.ToString(),
            [nameof(McpAutoStart)] = McpAutoStart.ToString(),
            [nameof(McpHeartbeatInterval)] = McpHeartbeatInterval.ToString()
        };
    }

    private void RestoreSnapshot()
    {
        _isLoading = true; // prevent dirty tracking during restore
        try
        {
            if (_originalValues.TryGetValue(nameof(OdooSwaggerUrl), out var url))
                OdooSwaggerUrl = url ?? string.Empty;
            if (_originalValues.TryGetValue(nameof(OdooApiToken), out var token))
                OdooApiToken = token ?? string.Empty;
            if (_originalValues.TryGetValue(nameof(OdooTimeoutSeconds), out var timeout))
                OdooTimeoutSeconds = int.TryParse(timeout, out var t) ? t : 30;
            if (_originalValues.TryGetValue(nameof(AutoCADConnectionTimeout), out var acTimeout))
                AutoCADConnectionTimeout = int.TryParse(acTimeout, out var act) ? act : 10;
            if (_originalValues.TryGetValue(nameof(AutoCADRetryAttempts), out var retry))
                AutoCADRetryAttempts = int.TryParse(retry, out var r) ? r : 3;
            if (_originalValues.TryGetValue(nameof(AutoCADMode), out var mode))
                AutoCADMode = mode == "File" ? "File" : "COM";
            if (_originalValues.TryGetValue(nameof(McpPort), out var port))
                McpPort = int.TryParse(port, out var p) ? p : 8084;
            if (_originalValues.TryGetValue(nameof(McpAutoStart), out var autoStart))
                McpAutoStart = bool.TryParse(autoStart, out var a) && a;
            if (_originalValues.TryGetValue(nameof(McpHeartbeatInterval), out var hb))
                McpHeartbeatInterval = int.TryParse(hb, out var h) ? h : 30;
        }
        finally
        {
            _isLoading = false;
        }
    }
}
