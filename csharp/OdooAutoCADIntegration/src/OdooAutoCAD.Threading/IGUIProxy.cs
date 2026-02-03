// OdooAutoCAD.Threading/IGUIProxy.cs
// GUI Proxy Interface - equivalent to Python util_gui_proxy.py

using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace OdooAutoCAD.Threading;

/// <summary>
/// Status of a GUI proxy request.
/// </summary>
public enum ProxyRequestStatus
{
    Pending,
    InProgress,
    Completed,
    Failed,
    TimedOut,
    Cancelled
}

/// <summary>
/// Response from a GUI proxy execution.
/// </summary>
public class GUIProxyResponse
{
    public string RequestId { get; set; } = string.Empty;
    public bool Success { get; set; }
    public object? Result { get; set; }
    public string? ErrorMessage { get; set; }
    public string? ErrorType { get; set; }
    public ProxyRequestStatus Status { get; set; }
    public DateTime StartTime { get; set; }
    public DateTime? EndTime { get; set; }
    public TimeSpan? Duration => EndTime.HasValue ? EndTime.Value - StartTime : null;

    public static GUIProxyResponse CreateSuccess(string requestId, object? result)
    {
        return new GUIProxyResponse
        {
            RequestId = requestId,
            Success = true,
            Result = result,
            Status = ProxyRequestStatus.Completed,
            EndTime = DateTime.UtcNow
        };
    }

    public static GUIProxyResponse CreateError(string requestId, string errorMessage, string? errorType = null)
    {
        return new GUIProxyResponse
        {
            RequestId = requestId,
            Success = false,
            ErrorMessage = errorMessage,
            ErrorType = errorType,
            Status = ProxyRequestStatus.Failed,
            EndTime = DateTime.UtcNow
        };
    }

    public static GUIProxyResponse CreateTimeout(string requestId)
    {
        return new GUIProxyResponse
        {
            RequestId = requestId,
            Success = false,
            ErrorMessage = "Request timed out",
            ErrorType = "TimeoutError",
            Status = ProxyRequestStatus.TimedOut,
            EndTime = DateTime.UtcNow
        };
    }
}

/// <summary>
/// Statistics about the GUI proxy operations.
/// </summary>
public record GUIProxyStats(
    int TotalRequests,
    int SuccessfulRequests,
    int FailedRequests,
    int TimedOutRequests,
    int PendingRequests,
    TimeSpan AverageExecutionTime);

/// <summary>
/// Handler delegate for processing GUI proxy actions.
/// </summary>
/// <param name="parameters">Action parameters.</param>
/// <returns>Result of the action.</returns>
public delegate Task<object?> GUIProxyHandler(Dictionary<string, object?> parameters);

/// <summary>
/// Interface for GUI Proxy system that enables thread-safe COM operations.
///
/// The GUI Proxy solves the COM threading conflict between MCP Server (running on MTA thread)
/// and AutoCAD COM operations (requiring STA/GUI thread). All AutoCAD operations are
/// queued and executed on the GUI main thread via a message queue architecture.
///
/// Equivalent to Python's util_gui_proxy.py (605 lines).
/// </summary>
public interface IGUIProxy : IDisposable
{
    /// <summary>
    /// Gets whether the proxy is running.
    /// </summary>
    bool IsRunning { get; }

    /// <summary>
    /// Gets the number of pending requests in the queue.
    /// </summary>
    int PendingRequestCount { get; }

    /// <summary>
    /// Starts the GUI proxy service.
    /// </summary>
    void Start();

    /// <summary>
    /// Stops the GUI proxy service.
    /// </summary>
    void Stop();

    /// <summary>
    /// Registers a handler for a specific action.
    /// </summary>
    /// <param name="action">Action name.</param>
    /// <param name="handler">Handler function to execute.</param>
    void RegisterHandler(string action, GUIProxyHandler handler);

    /// <summary>
    /// Unregisters a handler for a specific action.
    /// </summary>
    /// <param name="action">Action name.</param>
    /// <returns>True if handler was unregistered.</returns>
    bool UnregisterHandler(string action);

    /// <summary>
    /// Gets all registered action names.
    /// </summary>
    IReadOnlyList<string> GetRegisteredActions();

    /// <summary>
    /// Executes an action in the GUI thread asynchronously.
    /// This is the main method for thread-safe COM operations.
    /// </summary>
    /// <param name="action">Action name.</param>
    /// <param name="parameters">Action parameters.</param>
    /// <param name="timeout">Timeout in milliseconds (default: 10000).</param>
    /// <returns>Proxy response with result or error.</returns>
    Task<GUIProxyResponse> ExecuteInGuiAsync(
        string action,
        Dictionary<string, object?>? parameters = null,
        int timeout = 10000);

    /// <summary>
    /// Executes an action in the GUI thread synchronously (blocks until complete).
    /// Use with caution - prefer async version.
    /// </summary>
    /// <param name="action">Action name.</param>
    /// <param name="parameters">Action parameters.</param>
    /// <param name="timeout">Timeout in milliseconds (default: 10000).</param>
    /// <returns>Proxy response with result or error.</returns>
    GUIProxyResponse ExecuteInGui(
        string action,
        Dictionary<string, object?>? parameters = null,
        int timeout = 10000);

    /// <summary>
    /// Processes pending requests in the queue.
    /// Must be called from the GUI thread (typically via DispatcherTimer).
    /// </summary>
    /// <returns>Number of requests processed.</returns>
    int ProcessRequests();

    /// <summary>
    /// Gets statistics about proxy operations.
    /// </summary>
    GUIProxyStats GetStatistics();

    /// <summary>
    /// Clears all pending requests.
    /// </summary>
    void ClearPendingRequests();

    /// <summary>
    /// Event raised when a request is completed.
    /// </summary>
    event EventHandler<GUIProxyResponse>? RequestCompleted;

    /// <summary>
    /// Event raised when an error occurs during request processing.
    /// </summary>
    event EventHandler<Exception>? ErrorOccurred;
}
