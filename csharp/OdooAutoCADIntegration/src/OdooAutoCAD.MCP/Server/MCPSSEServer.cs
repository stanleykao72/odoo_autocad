// OdooAutoCAD.MCP/Server/MCPSSEServer.cs
// MCP SSE Server Implementation - equivalent to Python mcp_server_fastmcp.py

using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.MCP.Protocol;
using OdooAutoCAD.MCP.Tools;

namespace OdooAutoCAD.MCP.Server;

/// <summary>
/// SSE Connection state.
/// </summary>
public class SSEConnection
{
    public string ConnectionId { get; } = Guid.NewGuid().ToString("N");
    public HttpResponse Response { get; init; } = null!;
    public CancellationTokenSource CancellationTokenSource { get; } = new();
    public DateTime ConnectedAt { get; } = DateTime.UtcNow;
    public string? ClientInfo { get; set; }
}

/// <summary>
/// MCP Server using Server-Sent Events (SSE) transport.
/// Implements the MCP protocol for AI assistant integration.
///
/// Endpoints:
/// - GET /         : Server info
/// - GET /health   : Health check
/// - GET /sse      : SSE connection for real-time events
/// - POST /messages: JSON-RPC requests
/// </summary>
public class MCPSSEServer
{
    private readonly ILogger<MCPSSEServer>? _logger;
    private readonly MCPToolRegistry _toolRegistry;
    private readonly ConcurrentDictionary<string, SSEConnection> _connections = new();

    private WebApplication? _app;
    private CancellationTokenSource? _serverCts;
    private Task? _serverTask;
    private bool _isInitialized;
    private MCPClientInfo? _clientInfo;

    public int Port { get; }
    public bool IsRunning => _app != null && _serverTask != null && !_serverTask.IsCompleted;
    public int ActiveConnections => _connections.Count;

    public MCPSSEServer(MCPToolRegistry toolRegistry, int port = 8084, ILogger<MCPSSEServer>? logger = null)
    {
        _toolRegistry = toolRegistry;
        Port = port;
        _logger = logger;
    }

    /// <summary>
    /// Starts the MCP SSE server.
    /// </summary>
    public Task StartAsync()
    {
        if (IsRunning)
        {
            _logger?.LogWarning("MCP SSE Server is already running");
            return Task.CompletedTask;
        }

        _serverCts = new CancellationTokenSource();

        var builder = WebApplication.CreateBuilder();

        builder.Services.AddLogging(logging =>
        {
            logging.ClearProviders();
            logging.AddConsole();
        });

        _app = builder.Build();

        // Configure endpoints
        _app.MapGet("/", HandleRootAsync);
        _app.MapGet("/health", HandleHealthAsync);
        _app.MapGet("/sse", HandleSSEAsync);
        _app.MapPost("/messages", HandleMessagesAsync);

        // Start server
        _serverTask = _app.RunAsync($"http://localhost:{Port}");

        _logger?.LogInformation("MCP SSE Server started on port {Port}", Port);
        return Task.CompletedTask;
    }

    /// <summary>
    /// Stops the MCP SSE server.
    /// </summary>
    public async Task StopAsync()
    {
        if (!IsRunning)
        {
            _logger?.LogWarning("MCP SSE Server is not running");
            return;
        }

        // Close all SSE connections
        foreach (var connection in _connections.Values)
        {
            connection.CancellationTokenSource.Cancel();
        }
        _connections.Clear();

        _serverCts?.Cancel();

        if (_app != null)
        {
            await _app.StopAsync();
            await _app.DisposeAsync();
            _app = null;
        }

        _serverTask = null;
        _isInitialized = false;

        _logger?.LogInformation("MCP SSE Server stopped");
    }

    #region Endpoint Handlers

    private async Task HandleRootAsync(HttpContext context)
    {
        var info = new MCPServerInfo
        {
            Capabilities = new MCPCapabilities
            {
                Tools = new MCPToolsCapability { ListChanged = false }
            }
        };

        context.Response.ContentType = "application/json";
        await context.Response.WriteAsJsonAsync(info);
    }

    private async Task HandleHealthAsync(HttpContext context)
    {
        var health = new Dictionary<string, object>
        {
            ["status"] = "healthy",
            ["timestamp"] = DateTime.UtcNow.ToString("O"),
            ["active_connections"] = ActiveConnections,
            ["tools_count"] = _toolRegistry.GetTools().Count,
            ["initialized"] = _isInitialized
        };

        context.Response.ContentType = "application/json";
        await context.Response.WriteAsJsonAsync(health);
    }

    private async Task HandleSSEAsync(HttpContext context)
    {
        var connection = new SSEConnection { Response = context.Response };
        _connections[connection.ConnectionId] = connection;

        context.Response.Headers["Content-Type"] = "text/event-stream";
        context.Response.Headers["Cache-Control"] = "no-cache";
        context.Response.Headers["Connection"] = "keep-alive";
        context.Response.Headers["X-Accel-Buffering"] = "no";

        _logger?.LogInformation("SSE connection established: {ConnectionId}", connection.ConnectionId);

        try
        {
            // Send initial connection event
            await SendSSEEventAsync(context.Response, new SSEEvent
            {
                Event = "connected",
                Data = JsonSerializer.Serialize(new
                {
                    connectionId = connection.ConnectionId,
                    timestamp = DateTime.UtcNow.ToString("O")
                })
            });

            // Keep connection alive
            while (!context.RequestAborted.IsCancellationRequested &&
                   !connection.CancellationTokenSource.Token.IsCancellationRequested)
            {
                // Send heartbeat every 30 seconds
                await Task.Delay(30000, connection.CancellationTokenSource.Token);

                await SendSSEEventAsync(context.Response, new SSEEvent
                {
                    Event = "heartbeat",
                    Data = DateTime.UtcNow.ToString("O")
                });
            }
        }
        catch (OperationCanceledException)
        {
            // Normal disconnection
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "SSE connection error: {ConnectionId}", connection.ConnectionId);
        }
        finally
        {
            _connections.TryRemove(connection.ConnectionId, out _);
            _logger?.LogInformation("SSE connection closed: {ConnectionId}", connection.ConnectionId);
        }
    }

    private async Task HandleMessagesAsync(HttpContext context)
    {
        try
        {
            var request = await context.Request.ReadFromJsonAsync<JsonRpcRequest>();

            if (request == null)
            {
                await WriteJsonRpcErrorAsync(context, null, JsonRpcError.InvalidRequest, "Invalid JSON-RPC request");
                return;
            }

            _logger?.LogDebug("Received JSON-RPC request: {Method}", request.Method);

            var response = await ProcessJsonRpcRequestAsync(request);

            context.Response.ContentType = "application/json";
            await context.Response.WriteAsJsonAsync(response);
        }
        catch (JsonException)
        {
            await WriteJsonRpcErrorAsync(context, null, JsonRpcError.ParseError, "JSON parse error");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error processing message");
            await WriteJsonRpcErrorAsync(context, null, JsonRpcError.InternalError, ex.Message);
        }
    }

    #endregion

    #region JSON-RPC Processing

    private async Task<JsonRpcResponse> ProcessJsonRpcRequestAsync(JsonRpcRequest request)
    {
        return request.Method switch
        {
            "initialize" => await HandleInitializeAsync(request),
            "initialized" => HandleInitialized(request),
            "tools/list" => HandleToolsList(request),
            "tools/call" => await HandleToolsCallAsync(request),
            "ping" => HandlePing(request),
            _ => JsonRpcResponse.Failure(request.Id, JsonRpcError.MethodNotFound, $"Method not found: {request.Method}")
        };
    }

    private Task<JsonRpcResponse> HandleInitializeAsync(JsonRpcRequest request)
    {
        MCPInitializeParams? initParams = null;

        if (request.Params != null)
        {
            var paramsJson = JsonSerializer.Serialize(request.Params);
            initParams = JsonSerializer.Deserialize<MCPInitializeParams>(paramsJson);
            _clientInfo = initParams?.ClientInfo;
        }

        _logger?.LogInformation("Client initializing: {ClientName} v{ClientVersion}",
            _clientInfo?.Name ?? "Unknown",
            _clientInfo?.Version ?? "Unknown");

        var result = new MCPServerInfo
        {
            Capabilities = new MCPCapabilities
            {
                Tools = new MCPToolsCapability { ListChanged = false }
            }
        };

        return Task.FromResult(JsonRpcResponse.Success(request.Id, result));
    }

    private JsonRpcResponse HandleInitialized(JsonRpcRequest request)
    {
        _isInitialized = true;
        _logger?.LogInformation("MCP session initialized");
        return JsonRpcResponse.Success(request.Id, new { });
    }

    private JsonRpcResponse HandleToolsList(JsonRpcRequest request)
    {
        var result = new MCPToolsListResult
        {
            Tools = _toolRegistry.GetTools().ToList()
        };

        return JsonRpcResponse.Success(request.Id, result);
    }

    private async Task<JsonRpcResponse> HandleToolsCallAsync(JsonRpcRequest request)
    {
        MCPToolCallParams? callParams = null;

        if (request.Params != null)
        {
            var paramsJson = JsonSerializer.Serialize(request.Params);
            callParams = JsonSerializer.Deserialize<MCPToolCallParams>(paramsJson);
        }

        if (callParams == null || string.IsNullOrEmpty(callParams.Name))
        {
            return JsonRpcResponse.Failure(request.Id, JsonRpcError.InvalidParams, "Missing tool name");
        }

        _logger?.LogInformation("Executing tool: {ToolName}", callParams.Name);

        try
        {
            var result = await _toolRegistry.ExecuteToolAsync(callParams.Name, callParams.Arguments);
            return JsonRpcResponse.Success(request.Id, result);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Tool execution failed: {ToolName}", callParams.Name);
            return JsonRpcResponse.Failure(request.Id, JsonRpcError.InternalError, ex.Message);
        }
    }

    private JsonRpcResponse HandlePing(JsonRpcRequest request)
    {
        return JsonRpcResponse.Success(request.Id, new { });
    }

    #endregion

    #region Helper Methods

    private static async Task SendSSEEventAsync(HttpResponse response, SSEEvent evt)
    {
        var eventString = evt.ToString();
        var bytes = Encoding.UTF8.GetBytes(eventString);
        await response.Body.WriteAsync(bytes);
        await response.Body.FlushAsync();
    }

    private static async Task WriteJsonRpcErrorAsync(HttpContext context, object? id, int code, string message)
    {
        var response = JsonRpcResponse.Failure(id, code, message);
        context.Response.ContentType = "application/json";
        await context.Response.WriteAsJsonAsync(response);
    }

    /// <summary>
    /// Broadcasts an SSE event to all connected clients.
    /// </summary>
    public async Task BroadcastEventAsync(string eventName, object data)
    {
        var evt = new SSEEvent
        {
            Event = eventName,
            Data = JsonSerializer.Serialize(data)
        };

        var tasks = _connections.Values.Select(conn =>
            SendSSEEventAsync(conn.Response, evt));

        await Task.WhenAll(tasks);
    }

    #endregion
}
