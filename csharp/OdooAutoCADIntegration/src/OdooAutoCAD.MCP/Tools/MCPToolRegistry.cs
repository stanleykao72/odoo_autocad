// OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs
// MCP Tool Registry - Manages the 7 MCP tools

using System.Text.Json;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.MCP.Protocol;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.MCP.Tools;

/// <summary>
/// Tool execution delegate.
/// </summary>
public delegate Task<MCPToolCallResult> MCPToolExecutor(Dictionary<string, object?>? arguments);

/// <summary>
/// Registry and executor for all MCP tools.
/// Implements the 7 tools from Python mcp_server_fastmcp.py.
/// </summary>
public class MCPToolRegistry
{
    private readonly ILogger<MCPToolRegistry>? _logger;
    private readonly IGUIProxy _guiProxy;
    private readonly IAutoCADService? _autoCADService;
    private readonly IOdooService? _odooService;
    private readonly IBOQProcessor? _boqProcessor;

    private readonly Dictionary<string, MCPTool> _toolDefinitions = new();
    private readonly Dictionary<string, MCPToolExecutor> _toolExecutors = new();

    public MCPToolRegistry(
        IGUIProxy guiProxy,
        IAutoCADService? autoCADService = null,
        IOdooService? odooService = null,
        IBOQProcessor? boqProcessor = null,
        ILogger<MCPToolRegistry>? logger = null)
    {
        _guiProxy = guiProxy;
        _autoCADService = autoCADService;
        _odooService = odooService;
        _boqProcessor = boqProcessor;
        _logger = logger;

        RegisterAllTools();
    }

    /// <summary>
    /// Gets all registered tool definitions.
    /// </summary>
    public IReadOnlyList<MCPTool> GetTools() => _toolDefinitions.Values.ToList();

    /// <summary>
    /// Executes a tool by name.
    /// </summary>
    public async Task<MCPToolCallResult> ExecuteToolAsync(string toolName, Dictionary<string, object?>? arguments)
    {
        if (!_toolExecutors.TryGetValue(toolName, out var executor))
        {
            _logger?.LogWarning("Tool not found: {ToolName}", toolName);
            return CreateErrorResult($"Tool not found: {toolName}");
        }

        try
        {
            _logger?.LogInformation("Executing tool: {ToolName}", toolName);
            var result = await executor(arguments);
            _logger?.LogInformation("Tool completed: {ToolName}", toolName);
            return result;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Tool execution failed: {ToolName}", toolName);
            return CreateErrorResult($"Tool execution failed: {ex.Message}");
        }
    }

    private void RegisterAllTools()
    {
        // Tool 1: test_connection
        RegisterTool(
            "test_connection",
            "Test the MCP server connection and verify it's working properly.",
            new MCPInputSchema { Type = "object" },
            TestConnectionAsync);

        // Tool 2: get_server_info
        RegisterTool(
            "get_server_info",
            "Get detailed information about the MCP server including version, capabilities, and available tools.",
            new MCPInputSchema { Type = "object" },
            GetServerInfoAsync);

        // Tool 3: check_autocad_status
        RegisterTool(
            "check_autocad_status",
            "Check AutoCAD connection status and get application information.",
            new MCPInputSchema { Type = "object" },
            CheckAutoCADStatusAsync);

        // Tool 4: check_odoo_status
        RegisterTool(
            "check_odoo_status",
            "Check Odoo connection status and get server information.",
            new MCPInputSchema { Type = "object" },
            CheckOdooStatusAsync);

        // Tool 5: extract_autocad_parameters
        RegisterTool(
            "extract_autocad_parameters",
            "Extract parameters and data from AutoCAD drawing layouts.",
            new MCPInputSchema
            {
                Type = "object",
                Properties = new Dictionary<string, MCPPropertySchema>
                {
                    ["drawing_path"] = new MCPPropertySchema
                    {
                        Type = "string",
                        Description = "Path to the DWG file. If empty, uses the currently open drawing."
                    },
                    ["use_current_drawing"] = new MCPPropertySchema
                    {
                        Type = "boolean",
                        Description = "If true, extracts from currently open drawing. Default: true",
                        Default = true
                    }
                }
            },
            ExtractAutoCADParametersAsync);

        // Tool 6: sync_to_odoo
        RegisterTool(
            "sync_to_odoo",
            "Synchronize data to Odoo system. Supports sync types: 'parameters', 'boq', 'project'.",
            new MCPInputSchema
            {
                Type = "object",
                Properties = new Dictionary<string, MCPPropertySchema>
                {
                    ["data"] = new MCPPropertySchema
                    {
                        Type = "object",
                        Description = "Data to sync to Odoo"
                    },
                    ["sync_type"] = new MCPPropertySchema
                    {
                        Type = "string",
                        Description = "Type of sync operation",
                        Enum = new List<string> { "parameters", "boq", "project" }
                    }
                },
                Required = new List<string> { "data", "sync_type" }
            },
            SyncToOdooAsync);

        // Tool 7: generate_boq
        RegisterTool(
            "generate_boq",
            "Generate Bill of Quantities (BOQ) for a project, optionally including AutoCAD data.",
            new MCPInputSchema
            {
                Type = "object",
                Properties = new Dictionary<string, MCPPropertySchema>
                {
                    ["project_id"] = new MCPPropertySchema
                    {
                        Type = "integer",
                        Description = "Odoo project ID to generate BOQ for"
                    },
                    ["include_autocad_data"] = new MCPPropertySchema
                    {
                        Type = "boolean",
                        Description = "If true, includes data extracted from AutoCAD. Default: true",
                        Default = true
                    }
                },
                Required = new List<string> { "project_id" }
            },
            GenerateBOQAsync);

        _logger?.LogInformation("Registered {Count} MCP tools", _toolDefinitions.Count);
    }

    private void RegisterTool(string name, string description, MCPInputSchema inputSchema, MCPToolExecutor executor)
    {
        _toolDefinitions[name] = new MCPTool
        {
            Name = name,
            Description = description,
            InputSchema = inputSchema
        };
        _toolExecutors[name] = executor;
    }

    #region Tool Implementations

    private Task<MCPToolCallResult> TestConnectionAsync(Dictionary<string, object?>? arguments)
    {
        var result = new Dictionary<string, object>
        {
            ["status"] = "connected",
            ["message"] = "MCP server is running and responding",
            ["timestamp"] = DateTime.UtcNow.ToString("O"),
            ["server"] = "OdooAutoCAD MCP Server",
            ["version"] = "6.0.0"
        };

        return Task.FromResult(CreateSuccessResult(result));
    }

    private Task<MCPToolCallResult> GetServerInfoAsync(Dictionary<string, object?>? arguments)
    {
        var result = new Dictionary<string, object>
        {
            ["name"] = "OdooAutoCAD MCP Server",
            ["version"] = "6.0.0",
            ["protocol_version"] = "2024-11-05",
            ["capabilities"] = new Dictionary<string, object>
            {
                ["tools"] = true,
                ["resources"] = false,
                ["prompts"] = false
            },
            ["tools_count"] = _toolDefinitions.Count,
            ["available_tools"] = _toolDefinitions.Keys.ToList(),
            ["gui_proxy_status"] = _guiProxy.IsRunning ? "running" : "stopped",
            ["gui_proxy_pending_requests"] = _guiProxy.PendingRequestCount
        };

        return Task.FromResult(CreateSuccessResult(result));
    }

    private async Task<MCPToolCallResult> CheckAutoCADStatusAsync(Dictionary<string, object?>? arguments)
    {
        // Execute via GUI proxy to ensure thread safety
        var response = await _guiProxy.ExecuteInGuiAsync("check_autocad_status");

        if (!response.Success)
        {
            return CreateErrorResult(response.ErrorMessage ?? "Failed to check AutoCAD status");
        }

        return CreateSuccessResult(response.Result);
    }

    private async Task<MCPToolCallResult> CheckOdooStatusAsync(Dictionary<string, object?>? arguments)
    {
        if (_odooService == null)
        {
            return CreateErrorResult("Odoo service is not configured. Please set up Odoo connection in Settings.");
        }

        try
        {
            var status = await _odooService.GetStatusAsync();

            var result = new Dictionary<string, object?>
            {
                ["connected"] = status.IsConnected,
                ["server_url"] = status.ServerUrl,
                ["database"] = status.Database,
                ["username"] = status.Username,
                ["version"] = status.Version,
                ["error"] = status.ErrorMessage
            };

            return CreateSuccessResult(result);
        }
        catch (Exception ex)
        {
            return CreateErrorResult($"Failed to check Odoo status: {ex.Message}");
        }
    }

    private async Task<MCPToolCallResult> ExtractAutoCADParametersAsync(Dictionary<string, object?>? arguments)
    {
        var drawingPath = arguments?.GetValueOrDefault("drawing_path") as string;
        var useCurrentDrawing = arguments?.GetValueOrDefault("use_current_drawing") as bool? ?? true;

        var parameters = new Dictionary<string, object?>
        {
            ["drawing_path"] = drawingPath,
            ["use_current_drawing"] = useCurrentDrawing
        };

        // Execute via GUI proxy to ensure thread safety for COM operations
        var response = await _guiProxy.ExecuteInGuiAsync("extract_autocad_parameters", parameters);

        if (!response.Success)
        {
            return CreateErrorResult(response.ErrorMessage ?? "Failed to extract parameters");
        }

        return CreateSuccessResult(response.Result);
    }

    private async Task<MCPToolCallResult> SyncToOdooAsync(Dictionary<string, object?>? arguments)
    {
        if (_odooService == null)
        {
            return CreateErrorResult("Odoo service is not configured. Please set up Odoo connection in Settings.");
        }

        var data = arguments?.GetValueOrDefault("data");
        var syncType = arguments?.GetValueOrDefault("sync_type") as string;

        if (data == null || string.IsNullOrEmpty(syncType))
        {
            return CreateErrorResult("Missing required parameters: data and sync_type");
        }

        try
        {
            var result = await _odooService.SyncToOdooAsync(data, syncType);

            var response = new Dictionary<string, object>
            {
                ["success"] = result.Success,
                ["records_processed"] = result.RecordsProcessed,
                ["records_created"] = result.RecordsCreated,
                ["records_updated"] = result.RecordsUpdated,
                ["records_failed"] = result.RecordsFailed,
                ["errors"] = result.Errors ?? new List<string>()
            };

            return CreateSuccessResult(response);
        }
        catch (Exception ex)
        {
            return CreateErrorResult($"Sync to Odoo failed: {ex.Message}");
        }
    }

    private async Task<MCPToolCallResult> GenerateBOQAsync(Dictionary<string, object?>? arguments)
    {
        if (_boqProcessor == null)
        {
            return CreateErrorResult("BOQ processor is not configured. Please set up connections first.");
        }

        var projectIdObj = arguments?.GetValueOrDefault("project_id");
        var includeAutoCADData = arguments?.GetValueOrDefault("include_autocad_data") as bool? ?? true;

        if (projectIdObj == null || !int.TryParse(projectIdObj.ToString(), out var projectId))
        {
            return CreateErrorResult("Missing or invalid required parameter: project_id");
        }

        try
        {
            var result = await _boqProcessor.GenerateBOQForProjectAsync(projectId, includeAutoCADData);

            var response = new Dictionary<string, object>
            {
                ["success"] = result.Success,
                ["total_items"] = result.TotalItems,
                ["processed_items"] = result.ProcessedItems,
                ["skipped_items"] = result.SkippedItems,
                ["entries_count"] = result.Entries.Count,
                ["entries"] = result.Entries.Select(e => new Dictionary<string, object?>
                {
                    ["product_name"] = e.ProductName,
                    ["quantity"] = e.Quantity,
                    ["unit"] = e.UnitOfMeasure,
                    ["unit_price"] = e.UnitPrice,
                    ["description"] = e.Description
                }).ToList(),
                ["warnings"] = result.Warnings,
                ["errors"] = result.Errors,
                ["generated_at"] = result.GeneratedAt.ToString("O")
            };

            return CreateSuccessResult(response);
        }
        catch (Exception ex)
        {
            return CreateErrorResult($"BOQ generation failed: {ex.Message}");
        }
    }

    #endregion

    #region Helper Methods

    private static MCPToolCallResult CreateSuccessResult(object? result)
    {
        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        return new MCPToolCallResult
        {
            Content = new List<MCPContent> { MCPContent.CreateText(json) },
            IsError = false
        };
    }

    private static MCPToolCallResult CreateErrorResult(string errorMessage)
    {
        return new MCPToolCallResult
        {
            Content = new List<MCPContent> { MCPContent.CreateText(errorMessage) },
            IsError = true
        };
    }

    #endregion
}
