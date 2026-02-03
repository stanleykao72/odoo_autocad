# C# COM Architecture Design Document

> **Version**: 1.0
> **Last Updated**: February 2026
> **Target Framework**: .NET 8.0
> **Platform**: Windows Desktop (WPF)

## Table of Contents

1. [Overview and Goals](#overview-and-goals)
2. [Solution Structure](#solution-structure)
3. [Key Components](#key-components)
4. [Technology Stack](#technology-stack)
5. [Interface Definitions](#interface-definitions)
6. [Communication Patterns](#communication-patterns)
7. [Error Handling Strategy](#error-handling-strategy)

---

## Overview and Goals

### Purpose

This document describes the architecture for migrating the Odoo-AutoCAD Integration application from Python to C#. The new architecture leverages .NET 8's modern features while maintaining thread-safe COM interoperability with AutoCAD.

### Primary Goals

1. **Thread Safety**: Ensure all AutoCAD COM operations execute on the STA (Single-Threaded Apartment) thread
2. **Performance**: Improve startup time and runtime performance over Python implementation
3. **Maintainability**: Leverage strong typing and modern C# patterns for better code quality
4. **Extensibility**: Design for easy addition of new AutoCAD operations and Odoo integrations
5. **Reliability**: Implement robust error handling and recovery mechanisms

### Key Design Decisions

| Decision | Rationale |
|----------|-----------|
| WPF over Windows Forms | Modern XAML-based UI, better data binding, MVVM support |
| .NET 8 LTS | Long-term support, performance improvements, Native AOT potential |
| Channel-based messaging | High-performance, thread-safe producer/consumer pattern |
| EF Core with SQLite | Familiar ORM pattern, lightweight local storage |
| ASP.NET Core for MCP | Built-in SSE support, robust HTTP handling |

---

## Solution Structure

```
OdooAutoCAD/
├── OdooAutoCAD.sln                          # Solution file
│
├── src/
│   ├── OdooAutoCAD.Core/                    # Core domain logic
│   │   ├── Interfaces/                      # Core interfaces
│   │   │   ├── IAutoCADService.cs
│   │   │   ├── IOdooClient.cs
│   │   │   ├── IGuiProxy.cs
│   │   │   └── IBOQProcessor.cs
│   │   ├── Models/                          # Domain models
│   │   │   ├── DrawingParameter.cs
│   │   │   ├── BOQEntry.cs
│   │   │   ├── PurchaseRequisition.cs
│   │   │   └── ServerConfiguration.cs
│   │   ├── Services/                        # Core services
│   │   │   ├── BOQService.cs
│   │   │   └── ParameterValidationService.cs
│   │   └── OdooAutoCAD.Core.csproj
│   │
│   ├── OdooAutoCAD.AutoCAD/                 # AutoCAD COM integration
│   │   ├── AutoCADComService.cs             # COM wrapper service
│   │   ├── DrawingExtractor.cs              # Parameter extraction
│   │   ├── ComObjectFactory.cs              # COM object lifecycle
│   │   └── OdooAutoCAD.AutoCAD.csproj
│   │
│   ├── OdooAutoCAD.Odoo/                    # Odoo API integration
│   │   ├── OdooRestClient.cs                # REST API client
│   │   ├── OdooAuthenticator.cs             # Authentication handling
│   │   ├── Models/                          # Odoo-specific DTOs
│   │   └── OdooAutoCAD.Odoo.csproj
│   │
│   ├── OdooAutoCAD.Data/                    # Data access layer
│   │   ├── ApplicationDbContext.cs          # EF Core context
│   │   ├── Repositories/                    # Repository pattern
│   │   ├── Migrations/                      # EF Core migrations
│   │   └── OdooAutoCAD.Data.csproj
│   │
│   ├── OdooAutoCAD.GuiProxy/                # GUI Proxy system
│   │   ├── GuiProxyService.cs               # Main proxy service
│   │   ├── MessageQueue.cs                  # Request/response queues
│   │   ├── ProxyRequest.cs                  # Request message type
│   │   ├── ProxyResponse.cs                 # Response message type
│   │   └── OdooAutoCAD.GuiProxy.csproj
│   │
│   ├── OdooAutoCAD.MCP/                     # MCP SSE Server
│   │   ├── McpSseServer.cs                  # SSE endpoint handler
│   │   ├── Tools/                           # MCP tool implementations
│   │   │   ├── TestConnectionTool.cs
│   │   │   ├── CheckAutoCADStatusTool.cs
│   │   │   ├── ExtractParametersTool.cs
│   │   │   └── ...
│   │   ├── JsonRpc/                         # JSON-RPC handling
│   │   └── OdooAutoCAD.MCP.csproj
│   │
│   └── OdooAutoCAD.Desktop/                 # WPF Application
│       ├── App.xaml                         # Application entry
│       ├── MainWindow.xaml                  # Main window
│       ├── ViewModels/                      # MVVM ViewModels
│       ├── Views/                           # XAML Views
│       ├── Controls/                        # Custom controls
│       ├── Resources/                       # Styles, themes
│       └── OdooAutoCAD.Desktop.csproj
│
├── tests/
│   ├── OdooAutoCAD.Core.Tests/
│   ├── OdooAutoCAD.AutoCAD.Tests/
│   ├── OdooAutoCAD.Odoo.Tests/
│   ├── OdooAutoCAD.GuiProxy.Tests/
│   ├── OdooAutoCAD.MCP.Tests/
│   └── OdooAutoCAD.Integration.Tests/
│
├── docs/
│   └── architecture/
│       ├── CSHARP_COM_ARCHITECTURE.md
│       ├── THREADING_MODEL.md
│       ├── MCP_INTEGRATION.md
│       └── MIGRATION_GUIDE.md
│
└── tools/
    └── build/
        ├── build.ps1
        └── publish.ps1
```

---

## Key Components

### 1. AutoCAD COM Service

The `AutoCADComService` is responsible for all interactions with AutoCAD through COM interop. It must run on an STA thread to comply with COM apartment requirements.

**Responsibilities:**
- Establish and maintain AutoCAD COM connection
- Extract drawing parameters from DWG files
- Execute AutoCAD commands
- Handle COM object lifecycle and cleanup

**Key Features:**
- Lazy initialization of COM connection
- Automatic reconnection on failure
- Proper COM object disposal using `Marshal.ReleaseComObject`
- Thread affinity enforcement

### 2. GUI Proxy System

The GUI Proxy solves the fundamental problem of COM thread affinity. MCP requests arrive on background threads but must execute on the GUI's STA thread.

**Architecture:**
```
[MCP Thread] ---> [Request Queue] ---> [GUI Thread Dispatcher] ---> [AutoCAD COM]
                                              |
                                              v
[MCP Thread] <--- [Response Queue] <--- [Result/Error]
```

**Key Features:**
- `Channel<T>` based message passing
- Configurable timeout handling
- Request correlation via unique IDs
- Exception marshaling across threads

### 3. MCP Server

The MCP (Model Context Protocol) server provides AI assistant integration via Server-Sent Events (SSE).

**Responsibilities:**
- Host SSE endpoint for Gemini CLI connection
- Implement 7 MCP tools for AutoCAD/Odoo operations
- Handle JSON-RPC protocol messages
- Route tool calls through GUI Proxy

### 4. Data Layer

Entity Framework Core with SQLite provides local data persistence for configuration caching and user preferences.

**Key Entities:**
- `ServerConfiguration` - Odoo connection settings
- `ProductCache` - Cached product data from Odoo
- `UserPreference` - User settings and preferences
- `SyncHistory` - Synchronization audit log

---

## Technology Stack

### Core Technologies

| Component | Technology | Version |
|-----------|------------|---------|
| Runtime | .NET | 8.0 LTS |
| UI Framework | WPF | Built-in |
| ORM | Entity Framework Core | 8.0 |
| HTTP Client | HttpClient / Refit | Built-in / 7.0 |
| SSE Server | ASP.NET Core | 8.0 |
| DI Container | Microsoft.Extensions.DependencyInjection | 8.0 |
| Logging | Serilog | 3.1 |
| Testing | xUnit, Moq, FluentAssertions | Latest |

### NuGet Packages

```xml
<!-- Core packages -->
<PackageReference Include="Microsoft.Extensions.Hosting" Version="8.0.0" />
<PackageReference Include="Microsoft.Extensions.DependencyInjection" Version="8.0.0" />
<PackageReference Include="System.Threading.Channels" Version="8.0.0" />

<!-- Data access -->
<PackageReference Include="Microsoft.EntityFrameworkCore.Sqlite" Version="8.0.0" />
<PackageReference Include="Microsoft.EntityFrameworkCore.Design" Version="8.0.0" />

<!-- HTTP & API -->
<PackageReference Include="Refit" Version="7.0.0" />
<PackageReference Include="Refit.HttpClientFactory" Version="7.0.0" />
<PackageReference Include="Polly" Version="8.2.0" />

<!-- Logging -->
<PackageReference Include="Serilog" Version="3.1.1" />
<PackageReference Include="Serilog.Sinks.File" Version="5.0.0" />
<PackageReference Include="Serilog.Extensions.Logging" Version="8.0.0" />

<!-- Testing -->
<PackageReference Include="xunit" Version="2.6.2" />
<PackageReference Include="Moq" Version="4.20.70" />
<PackageReference Include="FluentAssertions" Version="6.12.0" />
```

---

## Interface Definitions

### IAutoCADService

```csharp
namespace OdooAutoCAD.Core.Interfaces;

/// <summary>
/// Defines the contract for AutoCAD COM operations.
/// All implementations must ensure thread-safe access to COM objects.
/// </summary>
public interface IAutoCADService
{
    /// <summary>
    /// Gets the current connection status to AutoCAD.
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// Gets information about the connected AutoCAD instance.
    /// </summary>
    AutoCADInfo? ApplicationInfo { get; }

    /// <summary>
    /// Attempts to establish a connection to a running AutoCAD instance.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>True if connection was successful, false otherwise.</returns>
    Task<bool> ConnectAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Disconnects from the current AutoCAD instance and releases COM resources.
    /// </summary>
    Task DisconnectAsync();

    /// <summary>
    /// Extracts parameters from the currently open drawing.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Collection of extracted drawing parameters.</returns>
    Task<IReadOnlyList<DrawingParameter>> ExtractParametersAsync(
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Extracts parameters from a specific drawing file.
    /// </summary>
    /// <param name="drawingPath">Full path to the DWG file.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Collection of extracted drawing parameters.</returns>
    Task<IReadOnlyList<DrawingParameter>> ExtractParametersFromFileAsync(
        string drawingPath,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets the path of the currently active drawing.
    /// </summary>
    /// <returns>Full path to the current drawing, or null if no drawing is open.</returns>
    Task<string?> GetCurrentDrawingPathAsync();
}

/// <summary>
/// Information about the connected AutoCAD application.
/// </summary>
public record AutoCADInfo(
    string Version,
    string ProductName,
    string? CurrentDrawing,
    bool IsDocumentOpen
);
```

### IOdooClient

```csharp
namespace OdooAutoCAD.Core.Interfaces;

/// <summary>
/// Defines the contract for Odoo REST API communication.
/// </summary>
public interface IOdooClient
{
    /// <summary>
    /// Gets the current connection status to Odoo.
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// Gets information about the connected Odoo server.
    /// </summary>
    OdooServerInfo? ServerInfo { get; }

    /// <summary>
    /// Authenticates with the Odoo server using stored credentials.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>True if authentication was successful.</returns>
    Task<bool> AuthenticateAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Retrieves all products from Odoo for caching.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Collection of Odoo products.</returns>
    Task<IReadOnlyList<OdooProduct>> GetProductsAsync(
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Creates a new BOQ entry in Odoo.
    /// </summary>
    /// <param name="entry">The BOQ entry to create.</param>
    /// <param name="projectId">Target project ID in Odoo.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>The created BOQ entry with Odoo ID.</returns>
    Task<BOQEntry> CreateBOQEntryAsync(
        BOQEntry entry,
        int projectId,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Submits a purchase requisition to Odoo.
    /// </summary>
    /// <param name="requisition">The purchase requisition to submit.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>The created purchase requisition with Odoo ID.</returns>
    Task<PurchaseRequisition> CreatePurchaseRequisitionAsync(
        PurchaseRequisition requisition,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Synchronizes data to Odoo based on sync type.
    /// </summary>
    /// <param name="data">Data to synchronize.</param>
    /// <param name="syncType">Type of synchronization operation.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Synchronization result.</returns>
    Task<SyncResult> SyncDataAsync(
        object data,
        SyncType syncType,
        CancellationToken cancellationToken = default);
}

public record OdooServerInfo(
    string ServerUrl,
    string DatabaseName,
    string Version,
    int UserId
);

public enum SyncType
{
    Parameters,
    BOQ,
    Project
}
```

### IGuiProxy

```csharp
namespace OdooAutoCAD.Core.Interfaces;

/// <summary>
/// Defines the contract for the GUI Proxy system that ensures
/// thread-safe execution of operations on the GUI thread.
/// </summary>
public interface IGuiProxy
{
    /// <summary>
    /// Indicates whether the proxy is currently running and accepting requests.
    /// </summary>
    bool IsRunning { get; }

    /// <summary>
    /// Executes an action on the GUI thread.
    /// </summary>
    /// <param name="action">The action to execute.</param>
    /// <param name="timeout">Maximum time to wait for completion.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>A task that completes when the action has executed.</returns>
    Task ExecuteAsync(
        Action action,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes a function on the GUI thread and returns the result.
    /// </summary>
    /// <typeparam name="T">The return type of the function.</typeparam>
    /// <param name="func">The function to execute.</param>
    /// <param name="timeout">Maximum time to wait for completion.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>The result of the function execution.</returns>
    Task<T> ExecuteAsync<T>(
        Func<T> func,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes an async function on the GUI thread and returns the result.
    /// </summary>
    /// <typeparam name="T">The return type of the function.</typeparam>
    /// <param name="asyncFunc">The async function to execute.</param>
    /// <param name="timeout">Maximum time to wait for completion.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>The result of the function execution.</returns>
    Task<T> ExecuteAsync<T>(
        Func<Task<T>> asyncFunc,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Starts the GUI proxy message processing loop.
    /// </summary>
    void Start();

    /// <summary>
    /// Stops the GUI proxy and cancels pending requests.
    /// </summary>
    Task StopAsync();
}
```

### IBOQProcessor

```csharp
namespace OdooAutoCAD.Core.Interfaces;

/// <summary>
/// Defines the contract for BOQ (Bill of Quantities) processing operations.
/// </summary>
public interface IBOQProcessor
{
    /// <summary>
    /// Validates a BOQ entry against Odoo product catalog.
    /// </summary>
    /// <param name="entry">The BOQ entry to validate.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Validation result with any errors or warnings.</returns>
    Task<ValidationResult> ValidateEntryAsync(
        BOQEntry entry,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Generates BOQ entries from drawing parameters.
    /// </summary>
    /// <param name="parameters">Drawing parameters to process.</param>
    /// <param name="projectId">Target project ID.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Generated BOQ entries.</returns>
    Task<IReadOnlyList<BOQEntry>> GenerateBOQAsync(
        IEnumerable<DrawingParameter> parameters,
        int projectId,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Pushes BOQ entries to Odoo.
    /// </summary>
    /// <param name="entries">BOQ entries to push.</param>
    /// <param name="projectId">Target project ID in Odoo.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Push result with success/failure details.</returns>
    Task<BOQPushResult> PushToOdooAsync(
        IEnumerable<BOQEntry> entries,
        int projectId,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Converts BOQ entries to purchase requisition.
    /// </summary>
    /// <param name="entries">BOQ entries to convert.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <returns>Generated purchase requisition.</returns>
    Task<PurchaseRequisition> ConvertToPurchaseRequisitionAsync(
        IEnumerable<BOQEntry> entries,
        CancellationToken cancellationToken = default);
}
```

---

## Communication Patterns

### Dependency Injection Setup

```csharp
// Program.cs or App.xaml.cs
public static IServiceProvider ConfigureServices()
{
    var services = new ServiceCollection();

    // Core services
    services.AddSingleton<IAutoCADService, AutoCADComService>();
    services.AddSingleton<IOdooClient, OdooRestClient>();
    services.AddSingleton<IGuiProxy, GuiProxyService>();
    services.AddScoped<IBOQProcessor, BOQProcessor>();

    // Data layer
    services.AddDbContext<ApplicationDbContext>(options =>
        options.UseSqlite("Data Source=database.db"));

    // HTTP clients with Polly retry policies
    services.AddHttpClient<IOdooClient, OdooRestClient>()
        .AddTransientHttpErrorPolicy(p =>
            p.WaitAndRetryAsync(3, _ => TimeSpan.FromSeconds(2)));

    // MCP Server
    services.AddSingleton<McpSseServer>();

    // Logging
    services.AddLogging(builder =>
    {
        builder.AddSerilog(new LoggerConfiguration()
            .WriteTo.File("logs/app-.log", rollingInterval: RollingInterval.Day)
            .CreateLogger());
    });

    return services.BuildServiceProvider();
}
```

### Event-Based Communication

```csharp
// Define events for cross-component communication
public class ConnectionEvents
{
    public event EventHandler<AutoCADConnectionEventArgs>? AutoCADConnectionChanged;
    public event EventHandler<OdooConnectionEventArgs>? OdooConnectionChanged;

    public void RaiseAutoCADConnectionChanged(bool isConnected, AutoCADInfo? info)
    {
        AutoCADConnectionChanged?.Invoke(this,
            new AutoCADConnectionEventArgs(isConnected, info));
    }
}

// ViewModels subscribe to events
public class MainViewModel : IDisposable
{
    private readonly ConnectionEvents _events;

    public MainViewModel(ConnectionEvents events)
    {
        _events = events;
        _events.AutoCADConnectionChanged += OnAutoCADConnectionChanged;
    }

    private void OnAutoCADConnectionChanged(object? sender, AutoCADConnectionEventArgs e)
    {
        IsAutoCADConnected = e.IsConnected;
        AutoCADVersion = e.Info?.Version ?? "Not connected";
    }
}
```

---

## Error Handling Strategy

### Exception Hierarchy

```csharp
namespace OdooAutoCAD.Core.Exceptions;

/// <summary>
/// Base exception for all OdooAutoCAD application errors.
/// </summary>
public class OdooAutoCADException : Exception
{
    public string ErrorCode { get; }

    public OdooAutoCADException(string message, string errorCode)
        : base(message)
    {
        ErrorCode = errorCode;
    }

    public OdooAutoCADException(string message, string errorCode, Exception inner)
        : base(message, inner)
    {
        ErrorCode = errorCode;
    }
}

/// <summary>
/// Exception thrown when AutoCAD COM operations fail.
/// </summary>
public class AutoCADComException : OdooAutoCADException
{
    public AutoCADComException(string message, Exception? inner = null)
        : base(message, "ACAD_COM_ERROR", inner ?? new Exception())
    {
    }
}

/// <summary>
/// Exception thrown when Odoo API operations fail.
/// </summary>
public class OdooApiException : OdooAutoCADException
{
    public int? StatusCode { get; }

    public OdooApiException(string message, int? statusCode = null)
        : base(message, "ODOO_API_ERROR")
    {
        StatusCode = statusCode;
    }
}

/// <summary>
/// Exception thrown when GUI proxy operations timeout.
/// </summary>
public class GuiProxyTimeoutException : OdooAutoCADException
{
    public TimeSpan Timeout { get; }

    public GuiProxyTimeoutException(TimeSpan timeout)
        : base($"GUI proxy operation timed out after {timeout.TotalSeconds} seconds",
               "GUI_PROXY_TIMEOUT")
    {
        Timeout = timeout;
    }
}
```

### Global Error Handler

```csharp
public class GlobalErrorHandler
{
    private readonly ILogger<GlobalErrorHandler> _logger;

    public GlobalErrorHandler(ILogger<GlobalErrorHandler> logger)
    {
        _logger = logger;
    }

    public async Task<T> ExecuteWithErrorHandlingAsync<T>(
        Func<Task<T>> operation,
        string operationName)
    {
        try
        {
            return await operation();
        }
        catch (AutoCADComException ex)
        {
            _logger.LogError(ex, "AutoCAD COM error in {Operation}", operationName);
            throw;
        }
        catch (OdooApiException ex)
        {
            _logger.LogError(ex, "Odoo API error in {Operation}: Status {Status}",
                operationName, ex.StatusCode);
            throw;
        }
        catch (GuiProxyTimeoutException ex)
        {
            _logger.LogWarning(ex, "GUI proxy timeout in {Operation}", operationName);
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error in {Operation}", operationName);
            throw new OdooAutoCADException(
                $"An unexpected error occurred in {operationName}",
                "UNEXPECTED_ERROR",
                ex);
        }
    }
}
```

---

## Next Steps

1. Review [THREADING_MODEL.md](./THREADING_MODEL.md) for detailed threading architecture
2. Review [MCP_INTEGRATION.md](./MCP_INTEGRATION.md) for MCP server implementation details
3. Follow [MIGRATION_GUIDE.md](./MIGRATION_GUIDE.md) for step-by-step migration plan

---

*Document Version: 1.0 | Created: February 2026*
