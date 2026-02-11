using System.Net.Http;
using System.Net.Http.Headers;
using System.Windows.Media;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.App.ViewModels;

/// <summary>
/// ViewModel for the Dashboard page.
/// Displays connection status for AutoCAD and Odoo, provides quick-connect
/// buttons and workflow navigation shortcuts.
/// </summary>
public partial class DashboardViewModel : ObservableObject
{
    private readonly IAutoCADService _autoCADService;
    private readonly IOdooService _odooService;
    private readonly IGUIProxy _guiProxy;
    private readonly INavigationService _navigationService;
    private readonly ISettingsService _settingsService;
    private readonly IConfiguration _configuration;
    private readonly IAppLogService _logService;
    private readonly ILogger<DashboardViewModel>? _logger;
    private readonly DispatcherTimer _statusTimer;

    // AutoCAD status
    [ObservableProperty]
    private bool _isAutoCADConnected;

    [ObservableProperty]
    private string _autoCADStatusText = "Disconnected";

    [ObservableProperty]
    private string _autoCADDocumentName = string.Empty;

    [ObservableProperty]
    private string _autoCADErrorMessage = string.Empty;

    [ObservableProperty]
    private Brush _autoCADStatusColor = Brushes.Gray;

    // Odoo status
    [ObservableProperty]
    private bool _isOdooConnected;

    [ObservableProperty]
    private string _odooStatusText = "Disconnected";

    [ObservableProperty]
    private string _odooServerUrl = string.Empty;

    [ObservableProperty]
    private string _odooErrorMessage = string.Empty;

    [ObservableProperty]
    private Brush _odooStatusColor = Brushes.Gray;

    // Button state
    [ObservableProperty]
    private string _connectAutoCADButtonText = "Connect";

    [ObservableProperty]
    private string _connectOdooButtonText = "Connect";

    [ObservableProperty]
    private bool _isConnectingAutoCAD;

    [ObservableProperty]
    private bool _isConnectingOdoo;

    public DashboardViewModel(
        IAutoCADService autoCADService,
        IOdooService odooService,
        IGUIProxy guiProxy,
        INavigationService navigationService,
        ISettingsService settingsService,
        IConfiguration configuration,
        IAppLogService logService,
        ILogger<DashboardViewModel>? logger = null)
    {
        _autoCADService = autoCADService;
        _odooService = odooService;
        _guiProxy = guiProxy;
        _navigationService = navigationService;
        _settingsService = settingsService;
        _configuration = configuration;
        _logService = logService;
        _logger = logger;

        // Status polling timer (2s interval)
        _statusTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _statusTimer.Tick += OnStatusTimerTick;
        _statusTimer.Start();

        // Initial status check
        UpdateStatusFromServices();
    }

    private void OnStatusTimerTick(object? sender, EventArgs e)
    {
        UpdateStatusFromServices();
    }

    /// <summary>
    /// Reads IsConnected from the service singletons and updates all UI properties.
    /// </summary>
    internal void UpdateStatusFromServices()
    {
        // AutoCAD
        var acConnected = _autoCADService.IsConnected;
        if (acConnected != IsAutoCADConnected)
        {
            IsAutoCADConnected = acConnected;
        }
        AutoCADStatusText = acConnected ? "Connected" : "Disconnected";
        AutoCADStatusColor = acConnected ? Brushes.Green : Brushes.Gray;

        // Odoo
        var odooConnected = _odooService.IsConnected;
        if (odooConnected != IsOdooConnected)
        {
            IsOdooConnected = odooConnected;
        }
        OdooStatusText = odooConnected ? "Connected" : "Disconnected";
        OdooStatusColor = odooConnected ? Brushes.Green : Brushes.Gray;
    }

    #region Property Changed Hooks

    partial void OnIsAutoCADConnectedChanged(bool value)
    {
        ConnectAutoCADButtonText = value ? "Connected" : "Connect";
        if (!value)
        {
            AutoCADDocumentName = string.Empty;
        }
        ConnectAutoCADCommand.NotifyCanExecuteChanged();
        NavigateToAutoCADCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsOdooConnectedChanged(bool value)
    {
        ConnectOdooButtonText = value ? "Connected" : "Connect";
        if (!value)
        {
            OdooServerUrl = string.Empty;
        }
        ConnectOdooCommand.NotifyCanExecuteChanged();
        NavigateToBOQCommand.NotifyCanExecuteChanged();
        NavigateToPRCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsConnectingAutoCADChanged(bool value)
    {
        ConnectAutoCADCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsConnectingOdooChanged(bool value)
    {
        ConnectOdooCommand.NotifyCanExecuteChanged();
    }

    #endregion

    #region Connect Commands

    [RelayCommand(CanExecute = nameof(CanConnectAutoCAD))]
    private async Task ConnectAutoCADAsync()
    {
        IsConnectingAutoCAD = true;
        ConnectAutoCADButtonText = "Connecting...";
        AutoCADErrorMessage = string.Empty;

        try
        {
            _logger?.LogInformation("Dashboard: connecting to AutoCAD");

            var response = await _guiProxy.ExecuteInGuiAsync("autocad_connect", null, timeout: 15000);

            if (response.Success && response.Result is bool connected && connected)
            {
                IsAutoCADConnected = true;
                AutoCADErrorMessage = string.Empty;
                _logService.Log("AutoCAD connected from Dashboard", "Dashboard");

                // Fetch document name
                var statusResponse = await _guiProxy.ExecuteInGuiAsync("autocad_get_status", null, timeout: 5000);
                if (statusResponse.Success && statusResponse.Result is AutoCADStatus status)
                {
                    AutoCADDocumentName = status.CurrentDocument ?? string.Empty;
                }
            }
            else
            {
                IsAutoCADConnected = false;
                AutoCADErrorMessage = "Unable to connect to AutoCAD. Please ensure AutoCAD is running.";
                _logService.Log("AutoCAD connection failed from Dashboard", "Dashboard", AppLogLevel.Warning);
            }
        }
        catch (Exception ex)
        {
            IsAutoCADConnected = false;
            AutoCADErrorMessage = "Unable to connect to AutoCAD. Please ensure AutoCAD is running.";
            _logger?.LogError(ex, "Dashboard: AutoCAD connection error");
        }
        finally
        {
            IsConnectingAutoCAD = false;
        }
    }

    private bool CanConnectAutoCAD() => !IsAutoCADConnected && !IsConnectingAutoCAD;

    [RelayCommand(CanExecute = nameof(CanConnectOdoo))]
    private async Task ConnectOdooAsync()
    {
        IsConnectingOdoo = true;
        ConnectOdooButtonText = "Connecting...";
        OdooErrorMessage = string.Empty;

        try
        {
            _logger?.LogInformation("Dashboard: connecting to Odoo");

            // Load credentials from settings
            var configs = await _settingsService.LoadServerConfigsAsync();
            configs.TryGetValue("odoo_swagger_url", out var swaggerUrl);
            configs.TryGetValue("odoo_user_token", out var userToken);

            if (string.IsNullOrEmpty(swaggerUrl) || string.IsNullOrEmpty(userToken))
            {
                // Fallback to AppSettings
                var appSettings = _settingsService.GetAppSettings();
                swaggerUrl ??= appSettings.Odoo.SwaggerUrl;
            }

            if (string.IsNullOrEmpty(swaggerUrl) || string.IsNullOrEmpty(userToken))
            {
                OdooErrorMessage = "No Odoo connection settings found. Please configure in Settings.";
                _logService.Log("Odoo connection failed - no settings", "Dashboard", AppLogLevel.Warning);
                return;
            }

            // Parse URL to extract components
            var parsed = OdooConnectionViewModel.ParseSwaggerUrl(swaggerUrl);
            if (parsed == null)
            {
                OdooErrorMessage = "Invalid Swagger URL configuration. Please check Settings.";
                return;
            }

            var (baseUrl, database, apiToken) = parsed.Value;

            // Test connection via Swagger + Basic Auth (same as OdooConnectionViewModel)
            var status = await TestSwaggerConnectionAsync(swaggerUrl, baseUrl, database, userToken);

            if (status.IsConnected)
            {
                IsOdooConnected = true;
                OdooServerUrl = baseUrl;
                OdooErrorMessage = string.Empty;
                _logService.Log($"Odoo connected from Dashboard: {baseUrl}", "Dashboard");
            }
            else
            {
                IsOdooConnected = false;
                OdooErrorMessage = status.ErrorMessage ?? "Unable to connect to Odoo server. Check connection settings.";
                _logService.Log($"Odoo connection failed: {status.ErrorMessage}", "Dashboard", AppLogLevel.Warning);
            }
        }
        catch (Exception ex)
        {
            IsOdooConnected = false;
            OdooErrorMessage = "Unable to connect to Odoo server. Check connection settings.";
            _logger?.LogError(ex, "Dashboard: Odoo connection error");
        }
        finally
        {
            IsConnectingOdoo = false;
        }
    }

    private bool CanConnectOdoo() => !IsOdooConnected && !IsConnectingOdoo;

    #endregion

    #region Navigation Commands

    [RelayCommand]
    private void NavigateToAutoCAD()
    {
        _navigationService.NavigateTo("AutoCAD");
    }

    [RelayCommand(CanExecute = nameof(CanNavigateToBOQ))]
    private void NavigateToBOQ()
    {
        _navigationService.NavigateTo("BOQ");
    }

    private bool CanNavigateToBOQ() => IsOdooConnected;

    [RelayCommand(CanExecute = nameof(CanNavigateToPR))]
    private void NavigateToPR()
    {
        _navigationService.NavigateTo("PR");
    }

    private bool CanNavigateToPR() => IsOdooConnected;

    #endregion

    #region Swagger Connection Test

    /// <summary>
    /// Tests Odoo connection via Swagger + Basic Auth (same approach as OdooConnectionViewModel).
    /// Step 1: Fetch swagger.json to validate server + API module + API access token.
    /// Step 2: POST with Basic Auth (database:userToken) to validate user credentials.
    /// </summary>
    private async Task<OdooStatus> TestSwaggerConnectionAsync(
        string swaggerUrl, string baseUrl, string database, string userToken)
    {
        try
        {
            var timeoutSeconds = _configuration.GetValue<int>("Odoo:TimeoutSeconds", 30);
            var timeout = TimeSpan.FromSeconds(Math.Clamp(timeoutSeconds, 5, 120));

            using var httpClient = new HttpClient { Timeout = timeout };

            // Step 1: Fetch swagger.json
            _logService.Log("Dashboard: Fetching swagger spec...", "Odoo");
            var swaggerResponse = await httpClient.GetAsync(swaggerUrl);

            if (!swaggerResponse.IsSuccessStatusCode)
            {
                return swaggerResponse.StatusCode switch
                {
                    System.Net.HttpStatusCode.NotFound =>
                        new OdooStatus(false, baseUrl, database, null, null,
                            "Swagger endpoint not found. Is the boq_import_api module installed?"),
                    System.Net.HttpStatusCode.Forbidden or
                    System.Net.HttpStatusCode.Unauthorized =>
                        new OdooStatus(false, baseUrl, database, null, null,
                            "API access denied. Check the token parameter in the Swagger URL"),
                    _ =>
                        new OdooStatus(false, baseUrl, database, null, null,
                            $"HTTP {(int)swaggerResponse.StatusCode}: {swaggerResponse.ReasonPhrase}")
                };
            }

            var swaggerJson = await swaggerResponse.Content.ReadAsStringAsync();

            // Parse swagger spec for basePath and version
            string? basePath = null;
            string? apiVersion = null;
            try
            {
                var swaggerDoc = System.Text.Json.JsonDocument.Parse(swaggerJson);
                if (swaggerDoc.RootElement.TryGetProperty("basePath", out var bp))
                    basePath = bp.GetString();
                if (swaggerDoc.RootElement.TryGetProperty("info", out var info) &&
                    info.TryGetProperty("version", out var ver))
                    apiVersion = ver.GetString();
            }
            catch (System.Text.Json.JsonException)
            {
                _logService.Log("Could not parse swagger spec as JSON", "Odoo", AppLogLevel.Warning);
            }

            // Step 2: Test Basic Auth with a lightweight API call
            _logService.Log("Dashboard: Testing Basic Auth...", "Odoo");
            var credentials = Convert.ToBase64String(
                System.Text.Encoding.UTF8.GetBytes($"{database}:{userToken}"));

            var apiEndpoint = basePath != null
                ? $"{baseUrl}{basePath}"
                : $"{baseUrl}/api/v1/boq_import_api";

            var testUrl = $"{apiEndpoint}/callMethodForJobWorkingPlanBoqModel";
            var testRequest = new HttpRequestMessage(HttpMethod.Post, testUrl);
            testRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            testRequest.Content = new StringContent(
                System.Text.Json.JsonSerializer.Serialize(new { method_name = "test_connection" }),
                System.Text.Encoding.UTF8,
                "application/json");

            var testResponse = await httpClient.SendAsync(testRequest);

            if (testResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                testResponse.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                return new OdooStatus(false, baseUrl, database, null, apiVersion,
                    "User token authentication failed. Check your User Token value");
            }

            // Any non-auth-error response means connection and auth work
            var versionDisplay = apiVersion != null ? $"API v{apiVersion}" : "API Available";
            return new OdooStatus(true, baseUrl, database, null, versionDisplay, null);
        }
        catch (TaskCanceledException)
        {
            return new OdooStatus(false, baseUrl, database, null, null,
                $"Connection timed out");
        }
        catch (HttpRequestException ex)
        {
            return new OdooStatus(false, baseUrl, database, null, null,
                $"Unable to reach server at {baseUrl}: {ex.Message}");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Unexpected error testing Odoo connection from Dashboard");
            return new OdooStatus(false, baseUrl, database, null, null,
                $"Unexpected error: {ex.Message}");
        }
    }

    #endregion
}
