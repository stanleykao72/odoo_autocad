// OdooAutoCAD.App/ViewModels/MainViewModel.cs
// Main window ViewModel using CommunityToolkit.Mvvm

using System.Windows.Media;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.MCP.Server;
using OdooAutoCAD.Core.Threading;
using OdooAutoCAD.Threading;

namespace OdooAutoCAD.App.ViewModels;

/// <summary>
/// Main window ViewModel.
/// Manages navigation state and connection status.
/// </summary>
public partial class MainViewModel : ObservableObject
{
    private readonly ILogger<MainViewModel>? _logger;
    private readonly INavigationService _navigationService;
    private readonly IGUIProxy _guiProxy;
    private readonly MCPSSEServer _mcpServer;
    private readonly IAutoCADService _autoCADService;
    private readonly IOdooService _odooService;
    private readonly DispatcherTimer _statusTimer;

    [ObservableProperty]
    private string _currentPageTitle = "Dashboard";

    [ObservableProperty]
    private string _statusMessage = "Ready";

    [ObservableProperty]
    private string _currentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

    [ObservableProperty]
    private string _activeNavButton = "BtnDashboard";

    // AutoCAD Status
    [ObservableProperty]
    private string _autoCADStatusText = "Disconnected";

    [ObservableProperty]
    private Brush _autoCADStatusColor = Brushes.Gray;

    // Odoo Status
    [ObservableProperty]
    private string _odooStatusText = "Disconnected";

    [ObservableProperty]
    private Brush _odooStatusColor = Brushes.Gray;

    // MCP Status
    [ObservableProperty]
    private string _mcpStatusText = "Stopped";

    [ObservableProperty]
    private Brush _mcpStatusColor = Brushes.Gray;

    public MainViewModel(
        INavigationService navigationService,
        IGUIProxy guiProxy,
        MCPSSEServer mcpServer,
        IAutoCADService autoCADService,
        IOdooService odooService,
        ILogger<MainViewModel>? logger = null)
    {
        _navigationService = navigationService;
        _guiProxy = guiProxy;
        _mcpServer = mcpServer;
        _autoCADService = autoCADService;
        _odooService = odooService;
        _logger = logger;

        // Setup status update timer
        _statusTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _statusTimer.Tick += OnStatusTimerTick;
        _statusTimer.Start();

        UpdateAllStatus();
    }

    private void OnStatusTimerTick(object? sender, EventArgs e)
    {
        CurrentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        UpdateAllStatus();
    }

    [RelayCommand]
    private Task RefreshAsync()
    {
        StatusMessage = "Refreshing...";
        UpdateAllStatus();
        StatusMessage = "Ready";
        return Task.CompletedTask;
    }

    [RelayCommand]
    private void Navigate(string pageName)
    {
        CurrentPageTitle = pageName switch
        {
            "Dashboard" => "Dashboard",
            "AutoCAD" => "AutoCAD Integration",
            "Odoo" => "Odoo Connection",
            "BOQ" => "BOQ Manager",
            "PR" => "Purchase Requisition",
            "ParameterConfig" => "Parameter Configuration",
            "MCP" => "AI Assistant (MCP)",
            "Settings" => "Settings",
            _ => pageName
        };

        ActiveNavButton = $"Btn{pageName}";
        _navigationService.NavigateTo(pageName);
    }

    private void UpdateAllStatus()
    {
        UpdateAutoCADStatus();
        UpdateOdooStatus();
        UpdateMCPStatus();
    }

    private void UpdateAutoCADStatus()
    {
        if (_autoCADService.IsConnected)
        {
            AutoCADStatusText = "Connected";
            AutoCADStatusColor = Brushes.Green;
        }
        else
        {
            AutoCADStatusText = "Disconnected";
            AutoCADStatusColor = Brushes.Gray;
        }
    }

    private void UpdateOdooStatus()
    {
        if (_odooService.IsConnected)
        {
            OdooStatusText = "Connected";
            OdooStatusColor = Brushes.Green;
        }
        else
        {
            OdooStatusText = "Disconnected";
            OdooStatusColor = Brushes.Gray;
        }
    }

    private void UpdateMCPStatus()
    {
        if (_mcpServer.IsRunning)
        {
            McpStatusText = $"Port {_mcpServer.Port}";
            McpStatusColor = Brushes.Green;
        }
        else
        {
            McpStatusText = "Stopped";
            McpStatusColor = Brushes.Gray;
        }
    }

    [RelayCommand]
    private async Task StartMCPServerAsync()
    {
        try
        {
            StatusMessage = "Starting MCP Server...";
            await _mcpServer.StartAsync();
            UpdateMCPStatus();
            StatusMessage = "MCP Server started";
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to start MCP Server");
            StatusMessage = $"Failed to start MCP Server: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task StopMCPServerAsync()
    {
        try
        {
            StatusMessage = "Stopping MCP Server...";
            await _mcpServer.StopAsync();
            UpdateMCPStatus();
            StatusMessage = "MCP Server stopped";
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to stop MCP Server");
            StatusMessage = $"Failed to stop MCP Server: {ex.Message}";
        }
    }
}

/// <summary>
/// ViewModel for BOQ management page.
/// </summary>
public partial class BOQViewModel : ObservableObject
{
    [ObservableProperty]
    private int _selectedProjectId;

    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private string _statusMessage = string.Empty;
}

