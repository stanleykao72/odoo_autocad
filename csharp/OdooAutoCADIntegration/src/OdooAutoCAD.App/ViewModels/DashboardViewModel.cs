using System.Windows.Media;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
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
        IAppLogService logService,
        ILogger<DashboardViewModel>? logger = null)
    {
        _autoCADService = autoCADService;
        _odooService = odooService;
        _guiProxy = guiProxy;
        _navigationService = navigationService;
        _settingsService = settingsService;
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

            // Test connection (non-persistent)
            var status = await _odooService.TestConnectionAsync(baseUrl, database, apiToken, userToken);

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
                OdooErrorMessage = "Unable to connect to Odoo server. Check connection settings.";
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
}
