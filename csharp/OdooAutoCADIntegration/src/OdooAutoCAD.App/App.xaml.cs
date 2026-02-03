// OdooAutoCAD.App/App.xaml.cs
// Application entry point with dependency injection setup

using System.Windows;
using System.Windows.Threading;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.MCP.Server;
using OdooAutoCAD.MCP.Tools;
using OdooAutoCAD.Core.Threading;
using OdooAutoCAD.Threading;
using Serilog;

namespace OdooAutoCAD.App;

/// <summary>
/// Application entry point.
/// Sets up dependency injection, logging, and initializes core services.
/// </summary>
public partial class App : Application
{
    private IHost? _host;
    private DispatcherTimer? _guiProxyTimer;

    public static IServiceProvider Services { get; private set; } = null!;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Configure Serilog
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.Console()
            .WriteTo.File("logs/odoo-autocad-.log",
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7)
            .CreateLogger();

        // Build host with DI
        _host = Host.CreateDefaultBuilder()
            .UseSerilog()
            .ConfigureServices((context, services) =>
            {
                ConfigureServices(services);
            })
            .Build();

        Services = _host.Services;

        // Start the host
        await _host.StartAsync();

        // Initialize GUI Proxy timer
        InitializeGUIProxyTimer();

        // Check for --enable-mcp flag
        if (e.Args.Contains("--enable-mcp"))
        {
            await StartMCPServerAsync();
        }

        Log.Information("Application started");
    }

    private void ConfigureServices(IServiceCollection services)
    {
        // Core services
        services.AddSingleton<IGUIProxy, GUIProxy>();

        // Register service implementations (these will be created later)
        // For now, use placeholder registrations
        services.AddSingleton<IAutoCADService>(sp =>
            throw new NotImplementedException("AutoCADService not yet implemented"));
        services.AddSingleton<IOdooService>(sp =>
            throw new NotImplementedException("OdooService not yet implemented"));
        services.AddSingleton<IBOQProcessor>(sp =>
            throw new NotImplementedException("BOQProcessor not yet implemented"));

        // MCP services
        services.AddSingleton<MCPToolRegistry>();
        services.AddSingleton<MCPSSEServer>();

        // ViewModels
        services.AddTransient<MainViewModel>();
        services.AddTransient<ConnectionStatusViewModel>();
        services.AddTransient<BOQViewModel>();
        services.AddTransient<SettingsViewModel>();

        // Application services
        services.AddSingleton<INavigationService, NavigationService>();
    }

    private void InitializeGUIProxyTimer()
    {
        var guiProxy = Services.GetRequiredService<IGUIProxy>();
        guiProxy.Start();

        // Create timer for processing GUI proxy requests (100ms interval)
        _guiProxyTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(100)
        };
        _guiProxyTimer.Tick += (s, e) =>
        {
            guiProxy.ProcessRequests();
        };
        _guiProxyTimer.Start();

        Log.Information("GUI Proxy timer initialized with 100ms interval");
    }

    private async Task StartMCPServerAsync()
    {
        try
        {
            var mcpServer = Services.GetRequiredService<MCPSSEServer>();
            await mcpServer.StartAsync();
            Log.Information("MCP SSE Server started automatically");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to start MCP SSE Server");
        }
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        Log.Information("Application shutting down");

        _guiProxyTimer?.Stop();

        // Stop MCP server if running
        var mcpServer = Services.GetService<MCPSSEServer>();
        if (mcpServer?.IsRunning == true)
        {
            await mcpServer.StopAsync();
        }

        // Stop GUI proxy
        var guiProxy = Services.GetService<IGUIProxy>();
        guiProxy?.Stop();

        // Stop host
        if (_host != null)
        {
            await _host.StopAsync();
            _host.Dispose();
        }

        await Log.CloseAndFlushAsync();

        base.OnExit(e);
    }
}
