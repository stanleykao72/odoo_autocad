# MCP Integration Design Document

> **Version**: 1.0
> **Last Updated**: February 2026
> **Protocol**: Model Context Protocol (MCP) with SSE Transport

## Table of Contents

1. [Overview](#overview)
2. [ASP.NET Core SSE Server Design](#aspnet-core-sse-server-design)
3. [MCP Tools Specification](#mcp-tools-specification)
4. [JSON-RPC Protocol Handling](#json-rpc-protocol-handling)
5. [Integration with GUI Proxy](#integration-with-gui-proxy)
6. [Error Handling](#error-handling)
7. [Testing Strategy](#testing-strategy)

---

## Overview

### What is MCP?

Model Context Protocol (MCP) is a protocol that enables AI assistants (like Gemini CLI, Claude) to interact with external tools and services. This implementation provides an SSE (Server-Sent Events) transport layer for real-time bidirectional communication.

### Architecture Overview

```
+------------------+     HTTP/SSE      +------------------+
|   AI Assistant   | <===============> |   MCP Server     |
|   (Gemini CLI)   |                   |   (ASP.NET Core) |
+------------------+                   +--------+---------+
                                               |
                                               | IGuiProxy
                                               v
                                       +------------------+
                                       |   AutoCAD COM    |
                                       |   (via Proxy)    |
                                       +------------------+
                                               |
                                               v
                                       +------------------+
                                       |   Odoo REST API  |
                                       +------------------+
```

### Key Features

- **SSE Transport**: Real-time bidirectional messaging over HTTP
- **JSON-RPC 2.0**: Standard request/response protocol
- **7 MCP Tools**: Full feature parity with Python implementation
- **Thread-Safe**: All AutoCAD operations via GUI Proxy
- **Graceful Shutdown**: Proper cleanup on server stop

---

## ASP.NET Core SSE Server Design

### Server Configuration

```csharp
namespace OdooAutoCAD.MCP;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

public class McpServerHost : IHostedService
{
    private readonly IWebHost _webHost;
    private readonly ILogger<McpServerHost> _logger;
    private readonly int _port;

    public McpServerHost(
        IServiceProvider serviceProvider,
        ILogger<McpServerHost> logger,
        int port = 8084)
    {
        _logger = logger;
        _port = port;

        _webHost = new WebHostBuilder()
            .UseKestrel(options =>
            {
                options.ListenLocalhost(_port);
            })
            .ConfigureServices(services =>
            {
                // Share services from parent container
                services.AddSingleton(serviceProvider.GetRequiredService<IGuiProxy>());
                services.AddSingleton(serviceProvider.GetRequiredService<IAutoCADService>());
                services.AddSingleton(serviceProvider.GetRequiredService<IOdooClient>());
                services.AddSingleton(serviceProvider.GetRequiredService<IBOQProcessor>());

                // MCP-specific services
                services.AddSingleton<McpToolRegistry>();
                services.AddSingleton<JsonRpcHandler>();
                services.AddSingleton<SseConnectionManager>();

                // Register all tools
                services.AddSingleton<IMcpTool, TestConnectionTool>();
                services.AddSingleton<IMcpTool, GetServerInfoTool>();
                services.AddSingleton<IMcpTool, CheckAutoCADStatusTool>();
                services.AddSingleton<IMcpTool, CheckOdooStatusTool>();
                services.AddSingleton<IMcpTool, ExtractParametersTool>();
                services.AddSingleton<IMcpTool, SyncToOdooTool>();
                services.AddSingleton<IMcpTool, GenerateBOQTool>();
            })
            .Configure(app =>
            {
                app.UseRouting();
                app.UseEndpoints(endpoints =>
                {
                    endpoints.MapGet("/health", HealthEndpoint.Handle);
                    endpoints.MapGet("/sse", SseEndpoint.Handle);
                    endpoints.MapPost("/message", MessageEndpoint.Handle);
                });
            })
            .Build();
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("Starting MCP SSE server on port {Port}", _port);
        await _webHost.StartAsync(cancellationToken);
        _logger.LogInformation("MCP SSE server started successfully");
    }

    public async Task StopAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("Stopping MCP SSE server");
        await _webHost.StopAsync(cancellationToken);
    }
}
```

### SSE Endpoint Implementation

```csharp
namespace OdooAutoCAD.MCP.Endpoints;

using System.Text.Json;
using Microsoft.AspNetCore.Http;

public static class SseEndpoint
{
    public static async Task Handle(HttpContext context)
    {
        var connectionManager = context.RequestServices
            .GetRequiredService<SseConnectionManager>();
        var logger = context.RequestServices
            .GetRequiredService<ILogger<SseConnection>>();

        // Set SSE headers
        context.Response.Headers["Content-Type"] = "text/event-stream";
        context.Response.Headers["Cache-Control"] = "no-cache";
        context.Response.Headers["Connection"] = "keep-alive";
        context.Response.Headers["X-Accel-Buffering"] = "no";

        var connectionId = Guid.NewGuid().ToString();
        var connection = new SseConnection(connectionId, context.Response, logger);

        connectionManager.AddConnection(connection);

        try
        {
            // Send initial connection event
            await connection.SendEventAsync("connected", new { connectionId });

            // Keep connection alive
            var cancellationToken = context.RequestAborted;
            while (!cancellationToken.IsCancellationRequested)
            {
                // Send keepalive every 30 seconds
                await connection.SendEventAsync("ping", new { timestamp = DateTime.UtcNow });
                await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            // Client disconnected
        }
        finally
        {
            connectionManager.RemoveConnection(connectionId);
            logger.LogInformation("SSE connection {ConnectionId} closed", connectionId);
        }
    }
}

public class SseConnection
{
    private readonly string _connectionId;
    private readonly HttpResponse _response;
    private readonly ILogger _logger;
    private readonly SemaphoreSlim _writeLock = new(1, 1);

    public string ConnectionId => _connectionId;

    public SseConnection(string connectionId, HttpResponse response, ILogger logger)
    {
        _connectionId = connectionId;
        _response = response;
        _logger = logger;
    }

    public async Task SendEventAsync(string eventType, object data)
    {
        await _writeLock.WaitAsync();
        try
        {
            var json = JsonSerializer.Serialize(data);
            await _response.WriteAsync($"event: {eventType}\n");
            await _response.WriteAsync($"data: {json}\n\n");
            await _response.Body.FlushAsync();
        }
        finally
        {
            _writeLock.Release();
        }
    }

    public async Task SendJsonRpcResponseAsync(JsonRpcResponse response)
    {
        await SendEventAsync("message", response);
    }
}
```

### SSE Connection Manager

```csharp
namespace OdooAutoCAD.MCP;

using System.Collections.Concurrent;

public class SseConnectionManager
{
    private readonly ConcurrentDictionary<string, SseConnection> _connections = new();
    private readonly ILogger<SseConnectionManager> _logger;

    public SseConnectionManager(ILogger<SseConnectionManager> logger)
    {
        _logger = logger;
    }

    public int ConnectionCount => _connections.Count;

    public void AddConnection(SseConnection connection)
    {
        _connections[connection.ConnectionId] = connection;
        _logger.LogInformation(
            "SSE connection added: {ConnectionId}. Total: {Count}",
            connection.ConnectionId,
            _connections.Count);
    }

    public void RemoveConnection(string connectionId)
    {
        _connections.TryRemove(connectionId, out _);
        _logger.LogInformation(
            "SSE connection removed: {ConnectionId}. Total: {Count}",
            connectionId,
            _connections.Count);
    }

    public SseConnection? GetConnection(string connectionId)
    {
        _connections.TryGetValue(connectionId, out var connection);
        return connection;
    }

    public async Task BroadcastAsync(string eventType, object data)
    {
        var tasks = _connections.Values.Select(c => c.SendEventAsync(eventType, data));
        await Task.WhenAll(tasks);
    }
}
```

---

## MCP Tools Specification

### Tool Registry

```csharp
namespace OdooAutoCAD.MCP;

public interface IMcpTool
{
    /// <summary>
    /// Unique name of the tool.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Human-readable description.
    /// </summary>
    string Description { get; }

    /// <summary>
    /// JSON Schema for input parameters.
    /// </summary>
    JsonElement InputSchema { get; }

    /// <summary>
    /// Executes the tool with given parameters.
    /// </summary>
    Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default);
}

public record McpToolResult(
    bool Success,
    object? Data,
    string? ErrorMessage = null
);

public class McpToolRegistry
{
    private readonly Dictionary<string, IMcpTool> _tools;
    private readonly ILogger<McpToolRegistry> _logger;

    public McpToolRegistry(
        IEnumerable<IMcpTool> tools,
        ILogger<McpToolRegistry> logger)
    {
        _tools = tools.ToDictionary(t => t.Name, StringComparer.OrdinalIgnoreCase);
        _logger = logger;

        _logger.LogInformation("Registered {Count} MCP tools: {Names}",
            _tools.Count,
            string.Join(", ", _tools.Keys));
    }

    public IMcpTool? GetTool(string name)
    {
        _tools.TryGetValue(name, out var tool);
        return tool;
    }

    public IReadOnlyList<ToolDefinition> GetToolDefinitions()
    {
        return _tools.Values
            .Select(t => new ToolDefinition(t.Name, t.Description, t.InputSchema))
            .ToList();
    }
}
```

### Tool Implementations

#### 1. TestConnectionTool

```csharp
namespace OdooAutoCAD.MCP.Tools;

public class TestConnectionTool : IMcpTool
{
    public string Name => "test_connection";
    public string Description => "Test MCP server connection status";

    public JsonElement InputSchema => JsonDocument.Parse("""
        {
            "type": "object",
            "properties": {},
            "required": []
        }
        """).RootElement;

    public Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default)
    {
        return Task.FromResult(new McpToolResult(
            Success: true,
            Data: new
            {
                status = "connected",
                message = "MCP server is running",
                timestamp = DateTime.UtcNow
            }
        ));
    }
}
```

#### 2. GetServerInfoTool

```csharp
namespace OdooAutoCAD.MCP.Tools;

public class GetServerInfoTool : IMcpTool
{
    public string Name => "get_server_info";
    public string Description => "Get MCP server information and available features";

    public JsonElement InputSchema => JsonDocument.Parse("""
        {
            "type": "object",
            "properties": {},
            "required": []
        }
        """).RootElement;

    private readonly McpToolRegistry _registry;

    public GetServerInfoTool(McpToolRegistry registry)
    {
        _registry = registry;
    }

    public Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default)
    {
        var tools = _registry.GetToolDefinitions();

        return Task.FromResult(new McpToolResult(
            Success: true,
            Data: new
            {
                name = "OdooAutoCAD MCP Server",
                version = "1.0.0",
                runtime = ".NET 8.0",
                tools = tools.Select(t => new { t.Name, t.Description }),
                capabilities = new[]
                {
                    "autocad_integration",
                    "odoo_integration",
                    "boq_processing",
                    "parameter_extraction"
                }
            }
        ));
    }
}
```

#### 3. CheckAutoCADStatusTool

```csharp
namespace OdooAutoCAD.MCP.Tools;

public class CheckAutoCADStatusTool : IMcpTool
{
    public string Name => "check_autocad_status";
    public string Description => "Check AutoCAD connection status and get application info";

    public JsonElement InputSchema => JsonDocument.Parse("""
        {
            "type": "object",
            "properties": {},
            "required": []
        }
        """).RootElement;

    private readonly IAutoCADService _autocadService;
    private readonly IGuiProxy _guiProxy;

    public CheckAutoCADStatusTool(
        IAutoCADService autocadService,
        IGuiProxy guiProxy)
    {
        _autocadService = autocadService;
        _guiProxy = guiProxy;
    }

    public async Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default)
    {
        try
        {
            // Execute on GUI thread via proxy
            var info = await _guiProxy.ExecuteAsync(() =>
            {
                if (!_autocadService.IsConnected)
                {
                    return new
                    {
                        connected = false,
                        message = "AutoCAD is not connected"
                    };
                }

                var appInfo = _autocadService.ApplicationInfo;
                return new
                {
                    connected = true,
                    version = appInfo?.Version,
                    productName = appInfo?.ProductName,
                    currentDrawing = appInfo?.CurrentDrawing,
                    isDocumentOpen = appInfo?.IsDocumentOpen
                };
            }, cancellationToken: cancellationToken);

            return new McpToolResult(Success: true, Data: info);
        }
        catch (Exception ex)
        {
            return new McpToolResult(
                Success: false,
                Data: null,
                ErrorMessage: $"Failed to check AutoCAD status: {ex.Message}");
        }
    }
}
```

#### 4. CheckOdooStatusTool

```csharp
namespace OdooAutoCAD.MCP.Tools;

public class CheckOdooStatusTool : IMcpTool
{
    public string Name => "check_odoo_status";
    public string Description => "Check Odoo connection status and get server info";

    public JsonElement InputSchema => JsonDocument.Parse("""
        {
            "type": "object",
            "properties": {},
            "required": []
        }
        """).RootElement;

    private readonly IOdooClient _odooClient;

    public CheckOdooStatusTool(IOdooClient odooClient)
    {
        _odooClient = odooClient;
    }

    public async Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default)
    {
        try
        {
            if (!_odooClient.IsConnected)
            {
                return new McpToolResult(
                    Success: true,
                    Data: new { connected = false, message = "Odoo is not connected" });
            }

            var serverInfo = _odooClient.ServerInfo;

            return new McpToolResult(
                Success: true,
                Data: new
                {
                    connected = true,
                    serverUrl = serverInfo?.ServerUrl,
                    database = serverInfo?.DatabaseName,
                    version = serverInfo?.Version,
                    userId = serverInfo?.UserId
                });
        }
        catch (Exception ex)
        {
            return new McpToolResult(
                Success: false,
                Data: null,
                ErrorMessage: $"Failed to check Odoo status: {ex.Message}");
        }
    }
}
```

#### 5. ExtractParametersTool

```csharp
namespace OdooAutoCAD.MCP.Tools;

public class ExtractParametersTool : IMcpTool
{
    public string Name => "extract_autocad_parameters";
    public string Description => "Extract parameters from AutoCAD drawing";

    public JsonElement InputSchema => JsonDocument.Parse("""
        {
            "type": "object",
            "properties": {
                "drawing_path": {
                    "type": "string",
                    "description": "Path to the DWG file (optional, uses current drawing if not specified)"
                },
                "use_current_drawing": {
                    "type": "boolean",
                    "description": "Use the currently open drawing",
                    "default": true
                }
            },
            "required": []
        }
        """).RootElement;

    private readonly IAutoCADService _autocadService;
    private readonly IGuiProxy _guiProxy;

    public ExtractParametersTool(
        IAutoCADService autocadService,
        IGuiProxy guiProxy)
    {
        _autocadService = autocadService;
        _guiProxy = guiProxy;
    }

    public async Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default)
    {
        try
        {
            string? drawingPath = null;
            bool useCurrentDrawing = true;

            if (parameters.TryGetProperty("drawing_path", out var pathProp))
                drawingPath = pathProp.GetString();

            if (parameters.TryGetProperty("use_current_drawing", out var useCurrent))
                useCurrentDrawing = useCurrent.GetBoolean();

            // Execute extraction on GUI thread
            var extractedParams = await _guiProxy.ExecuteAsync(async () =>
            {
                if (!string.IsNullOrEmpty(drawingPath))
                {
                    return await _autocadService.ExtractParametersFromFileAsync(
                        drawingPath, cancellationToken);
                }
                else if (useCurrentDrawing)
                {
                    return await _autocadService.ExtractParametersAsync(cancellationToken);
                }
                else
                {
                    throw new ArgumentException(
                        "Either drawing_path or use_current_drawing must be specified");
                }
            }, timeout: TimeSpan.FromMinutes(2), cancellationToken: cancellationToken);

            return new McpToolResult(
                Success: true,
                Data: new
                {
                    parameterCount = extractedParams.Count,
                    parameters = extractedParams.Select(p => new
                    {
                        name = p.Name,
                        value = p.Value,
                        unit = p.Unit,
                        category = p.Category
                    })
                });
        }
        catch (Exception ex)
        {
            return new McpToolResult(
                Success: false,
                Data: null,
                ErrorMessage: $"Failed to extract parameters: {ex.Message}");
        }
    }
}
```

#### 6. SyncToOdooTool

```csharp
namespace OdooAutoCAD.MCP.Tools;

public class SyncToOdooTool : IMcpTool
{
    public string Name => "sync_to_odoo";
    public string Description => "Synchronize data to Odoo system";

    public JsonElement InputSchema => JsonDocument.Parse("""
        {
            "type": "object",
            "properties": {
                "data": {
                    "type": "object",
                    "description": "Data to synchronize"
                },
                "sync_type": {
                    "type": "string",
                    "enum": ["parameters", "boq", "project"],
                    "description": "Type of synchronization"
                }
            },
            "required": ["data", "sync_type"]
        }
        """).RootElement;

    private readonly IOdooClient _odooClient;

    public SyncToOdooTool(IOdooClient odooClient)
    {
        _odooClient = odooClient;
    }

    public async Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default)
    {
        try
        {
            if (!parameters.TryGetProperty("data", out var data))
                throw new ArgumentException("data parameter is required");

            if (!parameters.TryGetProperty("sync_type", out var syncTypeProp))
                throw new ArgumentException("sync_type parameter is required");

            var syncTypeStr = syncTypeProp.GetString();
            var syncType = syncTypeStr switch
            {
                "parameters" => SyncType.Parameters,
                "boq" => SyncType.BOQ,
                "project" => SyncType.Project,
                _ => throw new ArgumentException($"Invalid sync_type: {syncTypeStr}")
            };

            var result = await _odooClient.SyncDataAsync(
                data.GetRawText(),
                syncType,
                cancellationToken);

            return new McpToolResult(
                Success: result.Success,
                Data: new
                {
                    syncedCount = result.SyncedCount,
                    failedCount = result.FailedCount,
                    errors = result.Errors
                },
                ErrorMessage: result.Success ? null : "Some items failed to sync");
        }
        catch (Exception ex)
        {
            return new McpToolResult(
                Success: false,
                Data: null,
                ErrorMessage: $"Failed to sync to Odoo: {ex.Message}");
        }
    }
}
```

#### 7. GenerateBOQTool

```csharp
namespace OdooAutoCAD.MCP.Tools;

public class GenerateBOQTool : IMcpTool
{
    public string Name => "generate_boq";
    public string Description => "Generate Bill of Quantities from drawing data";

    public JsonElement InputSchema => JsonDocument.Parse("""
        {
            "type": "object",
            "properties": {
                "project_id": {
                    "type": "integer",
                    "description": "Odoo project ID"
                },
                "include_autocad_data": {
                    "type": "boolean",
                    "description": "Include data from current AutoCAD drawing",
                    "default": true
                }
            },
            "required": ["project_id"]
        }
        """).RootElement;

    private readonly IBOQProcessor _boqProcessor;
    private readonly IAutoCADService _autocadService;
    private readonly IGuiProxy _guiProxy;

    public GenerateBOQTool(
        IBOQProcessor boqProcessor,
        IAutoCADService autocadService,
        IGuiProxy guiProxy)
    {
        _boqProcessor = boqProcessor;
        _autocadService = autocadService;
        _guiProxy = guiProxy;
    }

    public async Task<McpToolResult> ExecuteAsync(
        JsonElement parameters,
        CancellationToken cancellationToken = default)
    {
        try
        {
            if (!parameters.TryGetProperty("project_id", out var projectIdProp))
                throw new ArgumentException("project_id parameter is required");

            var projectId = projectIdProp.GetInt32();
            var includeAutoCAD = true;

            if (parameters.TryGetProperty("include_autocad_data", out var includeProp))
                includeAutoCAD = includeProp.GetBoolean();

            IReadOnlyList<DrawingParameter>? drawingParams = null;

            if (includeAutoCAD && _autocadService.IsConnected)
            {
                drawingParams = await _guiProxy.ExecuteAsync(
                    () => _autocadService.ExtractParametersAsync(cancellationToken),
                    timeout: TimeSpan.FromMinutes(2),
                    cancellationToken: cancellationToken);
            }

            var boqEntries = await _boqProcessor.GenerateBOQAsync(
                drawingParams ?? Array.Empty<DrawingParameter>(),
                projectId,
                cancellationToken);

            return new McpToolResult(
                Success: true,
                Data: new
                {
                    projectId,
                    entryCount = boqEntries.Count,
                    entries = boqEntries.Select(e => new
                    {
                        name = e.Name,
                        quantity = e.Quantity,
                        unit = e.Unit,
                        unitPrice = e.UnitPrice,
                        totalPrice = e.TotalPrice
                    }),
                    includesAutoCADData = drawingParams?.Count > 0
                });
        }
        catch (Exception ex)
        {
            return new McpToolResult(
                Success: false,
                Data: null,
                ErrorMessage: $"Failed to generate BOQ: {ex.Message}");
        }
    }
}
```

---

## JSON-RPC Protocol Handling

### JSON-RPC Types

```csharp
namespace OdooAutoCAD.MCP.JsonRpc;

using System.Text.Json;
using System.Text.Json.Serialization;

public class JsonRpcRequest
{
    [JsonPropertyName("jsonrpc")]
    public string JsonRpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public JsonElement? Id { get; set; }

    [JsonPropertyName("method")]
    public string Method { get; set; } = string.Empty;

    [JsonPropertyName("params")]
    public JsonElement? Params { get; set; }
}

public class JsonRpcResponse
{
    [JsonPropertyName("jsonrpc")]
    public string JsonRpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public JsonElement? Id { get; set; }

    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Result { get; set; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public JsonRpcError? Error { get; set; }
}

public class JsonRpcError
{
    [JsonPropertyName("code")]
    public int Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("data")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Data { get; set; }

    // Standard error codes
    public const int ParseError = -32700;
    public const int InvalidRequest = -32600;
    public const int MethodNotFound = -32601;
    public const int InvalidParams = -32602;
    public const int InternalError = -32603;
    public const int ServerError = -32000;
}
```

### JSON-RPC Handler

```csharp
namespace OdooAutoCAD.MCP.JsonRpc;

public class JsonRpcHandler
{
    private readonly McpToolRegistry _toolRegistry;
    private readonly ILogger<JsonRpcHandler> _logger;

    public JsonRpcHandler(
        McpToolRegistry toolRegistry,
        ILogger<JsonRpcHandler> logger)
    {
        _toolRegistry = toolRegistry;
        _logger = logger;
    }

    public async Task<JsonRpcResponse> HandleRequestAsync(
        JsonRpcRequest request,
        CancellationToken cancellationToken = default)
    {
        _logger.LogDebug("Handling JSON-RPC request: {Method}", request.Method);

        try
        {
            return request.Method switch
            {
                "initialize" => HandleInitialize(request),
                "tools/list" => HandleToolsList(request),
                "tools/call" => await HandleToolCallAsync(request, cancellationToken),
                "ping" => HandlePing(request),
                _ => CreateErrorResponse(request.Id,
                    JsonRpcError.MethodNotFound,
                    $"Method not found: {request.Method}")
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error handling JSON-RPC request");
            return CreateErrorResponse(request.Id,
                JsonRpcError.InternalError,
                ex.Message);
        }
    }

    private JsonRpcResponse HandleInitialize(JsonRpcRequest request)
    {
        return new JsonRpcResponse
        {
            Id = request.Id,
            Result = new
            {
                protocolVersion = "2024-11-05",
                serverInfo = new
                {
                    name = "OdooAutoCAD MCP Server",
                    version = "1.0.0"
                },
                capabilities = new
                {
                    tools = new { }
                }
            }
        };
    }

    private JsonRpcResponse HandleToolsList(JsonRpcRequest request)
    {
        var tools = _toolRegistry.GetToolDefinitions();

        return new JsonRpcResponse
        {
            Id = request.Id,
            Result = new
            {
                tools = tools.Select(t => new
                {
                    name = t.Name,
                    description = t.Description,
                    inputSchema = t.InputSchema
                })
            }
        };
    }

    private async Task<JsonRpcResponse> HandleToolCallAsync(
        JsonRpcRequest request,
        CancellationToken cancellationToken)
    {
        if (!request.Params.HasValue)
        {
            return CreateErrorResponse(request.Id,
                JsonRpcError.InvalidParams,
                "params is required for tools/call");
        }

        var paramsObj = request.Params.Value;

        if (!paramsObj.TryGetProperty("name", out var nameProp))
        {
            return CreateErrorResponse(request.Id,
                JsonRpcError.InvalidParams,
                "name is required in params");
        }

        var toolName = nameProp.GetString();
        var tool = _toolRegistry.GetTool(toolName!);

        if (tool is null)
        {
            return CreateErrorResponse(request.Id,
                JsonRpcError.MethodNotFound,
                $"Tool not found: {toolName}");
        }

        JsonElement arguments = default;
        if (paramsObj.TryGetProperty("arguments", out var argsProp))
        {
            arguments = argsProp;
        }

        _logger.LogInformation("Executing tool: {ToolName}", toolName);
        var result = await tool.ExecuteAsync(arguments, cancellationToken);

        if (result.Success)
        {
            return new JsonRpcResponse
            {
                Id = request.Id,
                Result = new
                {
                    content = new[]
                    {
                        new
                        {
                            type = "text",
                            text = JsonSerializer.Serialize(result.Data)
                        }
                    }
                }
            };
        }
        else
        {
            return new JsonRpcResponse
            {
                Id = request.Id,
                Result = new
                {
                    content = new[]
                    {
                        new
                        {
                            type = "text",
                            text = result.ErrorMessage
                        }
                    },
                    isError = true
                }
            };
        }
    }

    private JsonRpcResponse HandlePing(JsonRpcRequest request)
    {
        return new JsonRpcResponse
        {
            Id = request.Id,
            Result = new { }
        };
    }

    private static JsonRpcResponse CreateErrorResponse(
        JsonElement? id,
        int code,
        string message)
    {
        return new JsonRpcResponse
        {
            Id = id,
            Error = new JsonRpcError
            {
                Code = code,
                Message = message
            }
        };
    }
}
```

---

## Integration with GUI Proxy

### Flow Diagram

```
+-------------------+
| Gemini CLI sends  |
| JSON-RPC request  |
+--------+----------+
         |
         v
+-------------------+
| SSE Connection    |
| receives message  |
+--------+----------+
         |
         v
+-------------------+
| JsonRpcHandler    |
| parses request    |
+--------+----------+
         |
         v
+-------------------+
| Tool identified   |
| (e.g., extract)   |
+--------+----------+
         |
         | Tool needs AutoCAD COM?
         |
    +----+----+
    |         |
    v NO      v YES
+--------+ +-------------------+
| Direct | | GuiProxy          |
| exec   | | .ExecuteAsync()   |
+--------+ +--------+----------+
    |              |
    |              | Enqueue request
    |              v
    |      +-------------------+
    |      | GUI Thread        |
    |      | processes queue   |
    |      +--------+----------+
    |              |
    |              | Execute COM
    |              v
    |      +-------------------+
    |      | AutoCAD COM       |
    |      +--------+----------+
    |              |
    +------+-------+
           |
           v
+-------------------+
| Result returned   |
| via TCS           |
+--------+----------+
         |
         v
+-------------------+
| JSON-RPC response |
| sent via SSE      |
+-------------------+
```

### Service Registration

```csharp
// In App.xaml.cs or Startup
services.AddSingleton<IGuiProxy>(sp =>
{
    var dispatcher = System.Windows.Application.Current.Dispatcher;
    var messageQueue = sp.GetRequiredService<MessageQueue>();
    var logger = sp.GetRequiredService<ILogger<GuiProxyService>>();
    var options = sp.GetRequiredService<IOptions<GuiProxyOptions>>();

    var proxy = new GuiProxyService(messageQueue, dispatcher, logger, options);

    // Start proxy on main thread
    dispatcher.Invoke(() => proxy.Start());

    return proxy;
});

// MCP tools receive IGuiProxy via constructor injection
services.AddSingleton<IMcpTool, CheckAutoCADStatusTool>();
services.AddSingleton<IMcpTool, ExtractParametersTool>();
// etc.
```

---

## Error Handling

### Error Categories

| Category | JSON-RPC Code | Handling |
|----------|---------------|----------|
| Parse Error | -32700 | Return immediately |
| Invalid Request | -32600 | Return immediately |
| Method Not Found | -32601 | Return immediately |
| Invalid Params | -32602 | Return immediately |
| Tool Execution Error | -32000 | Log and return error result |
| GUI Proxy Timeout | -32001 | Log warning, return timeout error |
| COM Error | -32002 | Log error, return COM-specific message |

### Error Response Examples

```json
// Method not found
{
    "jsonrpc": "2.0",
    "id": 1,
    "error": {
        "code": -32601,
        "message": "Method not found: unknown_method"
    }
}

// Tool execution error
{
    "jsonrpc": "2.0",
    "id": 2,
    "result": {
        "content": [
            {
                "type": "text",
                "text": "Failed to extract parameters: AutoCAD is not connected"
            }
        ],
        "isError": true
    }
}

// GUI Proxy timeout
{
    "jsonrpc": "2.0",
    "id": 3,
    "error": {
        "code": -32001,
        "message": "Operation timed out after 30 seconds"
    }
}
```

---

## Testing Strategy

### Unit Tests

```csharp
[Fact]
public async Task TestConnectionTool_ReturnsConnectedStatus()
{
    // Arrange
    var tool = new TestConnectionTool();

    // Act
    var result = await tool.ExecuteAsync(default);

    // Assert
    Assert.True(result.Success);
    Assert.Contains("connected", result.Data?.ToString());
}

[Fact]
public async Task CheckAutoCADStatusTool_WhenNotConnected_ReturnsDisconnectedStatus()
{
    // Arrange
    var mockAutoCAD = new Mock<IAutoCADService>();
    mockAutoCAD.Setup(a => a.IsConnected).Returns(false);

    var mockProxy = new Mock<IGuiProxy>();
    mockProxy.Setup(p => p.ExecuteAsync(It.IsAny<Func<object>>(),
            It.IsAny<TimeSpan?>(), It.IsAny<CancellationToken>()))
        .ReturnsAsync((Func<object> f, TimeSpan? _, CancellationToken __) => f());

    var tool = new CheckAutoCADStatusTool(mockAutoCAD.Object, mockProxy.Object);

    // Act
    var result = await tool.ExecuteAsync(default);

    // Assert
    Assert.True(result.Success);
    // Result should indicate not connected
}
```

### Integration Tests

```csharp
[Fact]
public async Task McpServer_HandlesToolsListRequest()
{
    // Arrange
    using var host = await CreateTestHost();
    var client = new HttpClient();

    var request = new JsonRpcRequest
    {
        Id = JsonDocument.Parse("1").RootElement,
        Method = "tools/list"
    };

    // Act
    var response = await client.PostAsJsonAsync(
        "http://localhost:8084/message", request);
    var result = await response.Content.ReadFromJsonAsync<JsonRpcResponse>();

    // Assert
    Assert.NotNull(result?.Result);
    // Should contain 7 tools
}
```

---

*Document Version: 1.0 | Created: February 2026*
