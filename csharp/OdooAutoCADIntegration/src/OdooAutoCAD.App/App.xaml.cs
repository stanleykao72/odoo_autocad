// OdooAutoCAD.App/App.xaml.cs
// Application entry point with dependency injection setup

using System.Windows;
using System.Windows.Threading;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Configuration;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Data.Context;
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

        // Global exception handlers
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        // Configure Serilog
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.Console()
            .WriteTo.File("logs/odoo-autocad-.log",
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7)
            .CreateLogger();

        try
        {
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

            // Register OLE message filter — retries COM calls rejected by AutoCAD
            // during WPF layout processing (RPC_E_CALL_REJECTED)
            OleMessageFilter.Register();

            // Initialize GUI Proxy timer
            InitializeGUIProxyTimer();

            // Check for --enable-mcp flag
            if (e.Args.Contains("--enable-mcp"))
            {
                await StartMCPServerAsync();
            }

            // Manually create and show the main window
            var mainWindow = new Views.MainWindow();
            mainWindow.Show();

            Log.Information("Application started successfully with main window");
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Application startup failed");
            await Log.CloseAndFlushAsync();
            MessageBox.Show(
                $"Application failed to start:\n\n{ex.Message}\n\n{ex.StackTrace}",
                "Startup Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
        }
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        Log.Error(e.Exception, "Unhandled dispatcher exception");
        MessageBox.Show(
            $"An error occurred:\n\n{e.Exception.Message}",
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true;
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            Log.Fatal(ex, "Unhandled domain exception");
        }
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        Log.Error(e.Exception, "Unobserved task exception");
        e.SetObserved();
    }

    private void ConfigureServices(IServiceCollection services)
    {
        // Core services
        services.AddSingleton<IGUIProxy, GUIProxy>();

        // Data layer
        services.AddSingleton<AppDbContextFactory>(sp =>
            new AppDbContextFactory("Data Source=database.db"));
        services.AddTransient<AppDbContext>(sp =>
            sp.GetRequiredService<AppDbContextFactory>().CreateDbContext());

        // Configuration
        services.AddSingleton<ConfigurationLoader>(sp =>
        {
            var loader = new ConfigurationLoader();
            loader.LoadFromJson(); // Initialize settings on startup
            return loader;
        });

        // Settings service
        services.AddSingleton<ISettingsService, SettingsService>();

        // App log service (shared UI log)
        services.AddSingleton<IAppLogService, AppLogService>();

        // AutoCAD service (singleton for maintaining connection state)
        services.AddSingleton<IAutoCADService, AutoCADService>();

        // DWG file reader (ACadSharp — no COM, no running AutoCAD needed)
        services.AddSingleton<IDwgReaderService, DwgReaderService>();

        // Odoo service with configured timeout
        services.AddSingleton<IOdooService>(sp =>
        {
            var logger = sp.GetService<ILogger<OdooService>>();
            var config = sp.GetRequiredService<IConfiguration>();
            var timeoutSeconds = config.GetValue<int>("Odoo:TimeoutSeconds", 30);
            return new OdooService(logger, timeoutSeconds);
        });

        // Note: IBOQProcessor is not registered yet.
        // It will be added when real implementation is available.
        // MCPToolRegistry accepts it as optional (nullable) parameter.

        // MCP services - use factory to control construction
        services.AddSingleton<MCPToolRegistry>(sp => new MCPToolRegistry(
            sp.GetRequiredService<IGUIProxy>(),
            sp.GetService<IAutoCADService>(),
            sp.GetService<IOdooService>(),
            sp.GetService<IBOQProcessor>(),
            sp.GetService<ILogger<MCPToolRegistry>>()));
        services.AddSingleton<MCPSSEServer>(sp => new MCPSSEServer(
            sp.GetRequiredService<MCPToolRegistry>(),
            8084,
            sp.GetService<ILogger<MCPSSEServer>>()));

        // ViewModels
        services.AddTransient<MainViewModel>();
        services.AddTransient<ConnectionStatusViewModel>();
        services.AddTransient<AutoCADViewModel>();
        services.AddTransient<BOQViewModel>();
        services.AddTransient<SettingsViewModel>();
        services.AddTransient<OdooConnectionViewModel>();

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

        // Step 0: Revoke OLE message filter
        try
        {
            OleMessageFilter.Revoke();
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to revoke OLE message filter");
        }

        // Step 1: Stop GUI proxy timer
        try
        {
            _guiProxyTimer?.Stop();
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to stop GUI proxy timer");
        }

        // Step 2: Stop MCP server if running
        try
        {
            var mcpServer = Services.GetService<MCPSSEServer>();
            if (mcpServer?.IsRunning == true)
            {
                await mcpServer.StopAsync();
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to stop MCP server");
        }

        // Step 3: Stop GUI proxy
        try
        {
            var guiProxy = Services.GetService<IGUIProxy>();
            guiProxy?.Stop();
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to stop GUI proxy");
        }

        // Step 4: Stop host with timeout
        try
        {
            if (_host != null)
            {
                await _host.StopAsync(TimeSpan.FromSeconds(5));
                _host.Dispose();
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to stop host");
        }

        // Step 5: Flush logs
        try
        {
            await Log.CloseAndFlushAsync();
        }
        catch
        {
            // Silently ignore log flush failures
        }

        base.OnExit(e);
    }
}
