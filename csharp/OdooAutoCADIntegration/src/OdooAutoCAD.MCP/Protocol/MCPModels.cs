// OdooAutoCAD.MCP/Protocol/MCPModels.cs
// MCP Protocol Models - JSON-RPC 2.0 and MCP Protocol definitions

using System.Text.Json.Serialization;

namespace OdooAutoCAD.MCP.Protocol;

#region JSON-RPC 2.0 Models

/// <summary>
/// JSON-RPC 2.0 Request
/// </summary>
public class JsonRpcRequest
{
    [JsonPropertyName("jsonrpc")]
    public string JsonRpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public object? Id { get; set; }

    [JsonPropertyName("method")]
    public string Method { get; set; } = string.Empty;

    [JsonPropertyName("params")]
    public object? Params { get; set; }
}

/// <summary>
/// JSON-RPC 2.0 Response
/// </summary>
public class JsonRpcResponse
{
    [JsonPropertyName("jsonrpc")]
    public string JsonRpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public object? Id { get; set; }

    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Result { get; set; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public JsonRpcError? Error { get; set; }

    public static JsonRpcResponse Success(object? id, object? result)
    {
        return new JsonRpcResponse { Id = id, Result = result };
    }

    public static JsonRpcResponse Failure(object? id, int code, string message, object? data = null)
    {
        return new JsonRpcResponse
        {
            Id = id,
            Error = new JsonRpcError { Code = code, Message = message, Data = data }
        };
    }
}

/// <summary>
/// JSON-RPC 2.0 Error
/// </summary>
public class JsonRpcError
{
    [JsonPropertyName("code")]
    public int Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("data")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Data { get; set; }

    // Standard JSON-RPC error codes
    public const int ParseError = -32700;
    public const int InvalidRequest = -32600;
    public const int MethodNotFound = -32601;
    public const int InvalidParams = -32602;
    public const int InternalError = -32603;

    // Custom error codes (server-defined)
    public const int AutoCADError = -32001;
    public const int OdooError = -32002;
    public const int TimeoutError = -32003;
    public const int NotConnectedError = -32004;
}

#endregion

#region MCP Protocol Models

/// <summary>
/// MCP Tool Definition
/// </summary>
public class MCPTool
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;

    [JsonPropertyName("inputSchema")]
    public MCPInputSchema InputSchema { get; set; } = new();
}

/// <summary>
/// MCP Input Schema (JSON Schema subset)
/// </summary>
public class MCPInputSchema
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "object";

    [JsonPropertyName("properties")]
    public Dictionary<string, MCPPropertySchema>? Properties { get; set; }

    [JsonPropertyName("required")]
    public List<string>? Required { get; set; }
}

/// <summary>
/// MCP Property Schema
/// </summary>
public class MCPPropertySchema
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "string";

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("default")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Default { get; set; }

    [JsonPropertyName("enum")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Enum { get; set; }
}

/// <summary>
/// MCP Server Capabilities
/// </summary>
public class MCPCapabilities
{
    [JsonPropertyName("tools")]
    public MCPToolsCapability? Tools { get; set; }

    [JsonPropertyName("resources")]
    public MCPResourcesCapability? Resources { get; set; }

    [JsonPropertyName("prompts")]
    public MCPPromptsCapability? Prompts { get; set; }
}

public class MCPToolsCapability
{
    [JsonPropertyName("listChanged")]
    public bool ListChanged { get; set; } = false;
}

public class MCPResourcesCapability
{
    [JsonPropertyName("subscribe")]
    public bool Subscribe { get; set; } = false;

    [JsonPropertyName("listChanged")]
    public bool ListChanged { get; set; } = false;
}

public class MCPPromptsCapability
{
    [JsonPropertyName("listChanged")]
    public bool ListChanged { get; set; } = false;
}

/// <summary>
/// MCP Server Info Response
/// </summary>
public class MCPServerInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "OdooAutoCAD";

    [JsonPropertyName("version")]
    public string Version { get; set; } = "6.0.0";

    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = "2024-11-05";

    [JsonPropertyName("capabilities")]
    public MCPCapabilities Capabilities { get; set; } = new();
}

/// <summary>
/// MCP Initialize Request Params
/// </summary>
public class MCPInitializeParams
{
    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = string.Empty;

    [JsonPropertyName("capabilities")]
    public MCPClientCapabilities? Capabilities { get; set; }

    [JsonPropertyName("clientInfo")]
    public MCPClientInfo? ClientInfo { get; set; }
}

public class MCPClientCapabilities
{
    [JsonPropertyName("roots")]
    public MCPRootsCapability? Roots { get; set; }

    [JsonPropertyName("sampling")]
    public object? Sampling { get; set; }
}

public class MCPRootsCapability
{
    [JsonPropertyName("listChanged")]
    public bool ListChanged { get; set; } = false;
}

public class MCPClientInfo
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("version")]
    public string Version { get; set; } = string.Empty;
}

/// <summary>
/// MCP Tools List Response
/// </summary>
public class MCPToolsListResult
{
    [JsonPropertyName("tools")]
    public List<MCPTool> Tools { get; set; } = new();
}

/// <summary>
/// MCP Tool Call Request Params
/// </summary>
public class MCPToolCallParams
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("arguments")]
    public Dictionary<string, object?>? Arguments { get; set; }
}

/// <summary>
/// MCP Tool Call Result
/// </summary>
public class MCPToolCallResult
{
    [JsonPropertyName("content")]
    public List<MCPContent> Content { get; set; } = new();

    [JsonPropertyName("isError")]
    public bool IsError { get; set; } = false;
}

/// <summary>
/// MCP Content (text or image)
/// </summary>
public class MCPContent
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "text";

    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; set; }

    [JsonPropertyName("data")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Data { get; set; }

    [JsonPropertyName("mimeType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MimeType { get; set; }

    public static MCPContent CreateText(string text) => new() { Type = "text", Text = text };

    public static MCPContent CreateImage(string base64Data, string mimeType = "image/png") =>
        new() { Type = "image", Data = base64Data, MimeType = mimeType };
}

#endregion

#region SSE Models

/// <summary>
/// SSE Event
/// </summary>
public class SSEEvent
{
    public string? Event { get; set; }
    public string Data { get; set; } = string.Empty;
    public string? Id { get; set; }
    public int? Retry { get; set; }

    public override string ToString()
    {
        var sb = new System.Text.StringBuilder();

        if (!string.IsNullOrEmpty(Event))
            sb.AppendLine($"event: {Event}");

        if (!string.IsNullOrEmpty(Id))
            sb.AppendLine($"id: {Id}");

        if (Retry.HasValue)
            sb.AppendLine($"retry: {Retry}");

        // Data can be multi-line, each line needs "data: " prefix
        foreach (var line in Data.Split('\n'))
        {
            sb.AppendLine($"data: {line}");
        }

        sb.AppendLine(); // Empty line to end the event

        return sb.ToString();
    }
}

#endregion
