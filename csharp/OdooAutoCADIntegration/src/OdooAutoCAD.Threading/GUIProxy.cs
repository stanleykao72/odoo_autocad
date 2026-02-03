// OdooAutoCAD.Threading/GUIProxy.cs
// GUI Proxy Implementation - equivalent to Python util_gui_proxy.py

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.Threading;

/// <summary>
/// Internal request object for the GUI proxy queue.
/// </summary>
internal class GUIProxyRequest
{
    public string RequestId { get; } = Guid.NewGuid().ToString("N");
    public string Action { get; init; } = string.Empty;
    public Dictionary<string, object?> Parameters { get; init; } = new();
    public TaskCompletionSource<GUIProxyResponse> CompletionSource { get; } = new();
    public DateTime CreatedAt { get; } = DateTime.UtcNow;
    public int TimeoutMs { get; init; } = 10000;
    public CancellationTokenSource CancellationTokenSource { get; } = new();
}

/// <summary>
/// GUI Proxy implementation that enables thread-safe COM operations.
///
/// This class implements a message queue architecture where:
/// 1. MCP Server threads submit requests to a concurrent queue
/// 2. GUI thread (via DispatcherTimer) processes requests and executes handlers
/// 3. Results are returned via TaskCompletionSource
///
/// Key features:
/// - ConcurrentQueue for thread-safe request queuing
/// - 100ms polling interval for responsive execution
/// - 10 second default timeout to prevent deadlocks
/// - Support for both async and sync execution patterns
/// </summary>
public class GUIProxy : IGUIProxy
{
    private readonly ILogger<GUIProxy>? _logger;
    private readonly ConcurrentQueue<GUIProxyRequest> _requestQueue = new();
    private readonly ConcurrentDictionary<string, GUIProxyHandler> _handlers = new();
    private readonly ConcurrentDictionary<string, GUIProxyRequest> _pendingRequests = new();

    private bool _isRunning;
    private int _totalRequests;
    private int _successfulRequests;
    private int _failedRequests;
    private int _timedOutRequests;
    private long _totalExecutionTimeMs;

    public GUIProxy(ILogger<GUIProxy>? logger = null)
    {
        _logger = logger;
    }

    public bool IsRunning => _isRunning;

    public int PendingRequestCount => _requestQueue.Count + _pendingRequests.Count;

    public void Start()
    {
        if (_isRunning)
        {
            _logger?.LogWarning("GUIProxy is already running");
            return;
        }

        _isRunning = true;
        _logger?.LogInformation("GUIProxy started");
    }

    public void Stop()
    {
        if (!_isRunning)
        {
            _logger?.LogWarning("GUIProxy is not running");
            return;
        }

        _isRunning = false;

        // Cancel all pending requests
        foreach (var request in _pendingRequests.Values)
        {
            request.CancellationTokenSource.Cancel();
            request.CompletionSource.TrySetResult(
                GUIProxyResponse.CreateError(request.RequestId, "GUIProxy stopped", "ShutdownError"));
        }

        _pendingRequests.Clear();

        // Clear the queue
        while (_requestQueue.TryDequeue(out var request))
        {
            request.CompletionSource.TrySetResult(
                GUIProxyResponse.CreateError(request.RequestId, "GUIProxy stopped", "ShutdownError"));
        }

        _logger?.LogInformation("GUIProxy stopped");
    }

    public void RegisterHandler(string action, GUIProxyHandler handler)
    {
        if (string.IsNullOrWhiteSpace(action))
            throw new ArgumentException("Action cannot be null or empty", nameof(action));

        if (handler == null)
            throw new ArgumentNullException(nameof(handler));

        _handlers[action] = handler;
        _logger?.LogDebug("Handler registered for action: {Action}", action);
    }

    public bool UnregisterHandler(string action)
    {
        var removed = _handlers.TryRemove(action, out _);
        if (removed)
        {
            _logger?.LogDebug("Handler unregistered for action: {Action}", action);
        }
        return removed;
    }

    public IReadOnlyList<string> GetRegisteredActions()
    {
        return _handlers.Keys.ToList();
    }

    public async Task<GUIProxyResponse> ExecuteInGuiAsync(
        string action,
        Dictionary<string, object?>? parameters = null,
        int timeout = 10000)
    {
        if (!_isRunning)
        {
            return GUIProxyResponse.CreateError(
                Guid.NewGuid().ToString("N"),
                "GUIProxy is not running",
                "NotRunningError");
        }

        if (!_handlers.ContainsKey(action))
        {
            return GUIProxyResponse.CreateError(
                Guid.NewGuid().ToString("N"),
                $"No handler registered for action: {action}",
                "HandlerNotFoundError");
        }

        var request = new GUIProxyRequest
        {
            Action = action,
            Parameters = parameters ?? new Dictionary<string, object?>(),
            TimeoutMs = timeout
        };

        _requestQueue.Enqueue(request);
        _pendingRequests[request.RequestId] = request;

        Interlocked.Increment(ref _totalRequests);

        _logger?.LogDebug("Request queued: {RequestId} - {Action}", request.RequestId, action);

        try
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(
                request.CancellationTokenSource.Token);
            cts.CancelAfter(timeout);

            var result = await request.CompletionSource.Task.WaitAsync(cts.Token);
            return result;
        }
        catch (OperationCanceledException)
        {
            Interlocked.Increment(ref _timedOutRequests);
            _pendingRequests.TryRemove(request.RequestId, out _);

            var response = GUIProxyResponse.CreateTimeout(request.RequestId);
            _logger?.LogWarning("Request timed out: {RequestId} - {Action}", request.RequestId, action);

            return response;
        }
    }

    public GUIProxyResponse ExecuteInGui(
        string action,
        Dictionary<string, object?>? parameters = null,
        int timeout = 10000)
    {
        return ExecuteInGuiAsync(action, parameters, timeout).GetAwaiter().GetResult();
    }

    public int ProcessRequests()
    {
        if (!_isRunning)
            return 0;

        int processedCount = 0;
        var stopwatch = Stopwatch.StartNew();

        // Process up to 10 requests per call to avoid blocking the GUI thread too long
        const int maxRequestsPerBatch = 10;

        while (processedCount < maxRequestsPerBatch && _requestQueue.TryDequeue(out var request))
        {
            try
            {
                ProcessSingleRequest(request);
                processedCount++;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error processing request: {RequestId}", request.RequestId);

                var response = GUIProxyResponse.CreateError(
                    request.RequestId,
                    ex.Message,
                    ex.GetType().Name);

                request.CompletionSource.TrySetResult(response);
                Interlocked.Increment(ref _failedRequests);

                ErrorOccurred?.Invoke(this, ex);
            }
            finally
            {
                _pendingRequests.TryRemove(request.RequestId, out _);
            }
        }

        if (processedCount > 0)
        {
            Interlocked.Add(ref _totalExecutionTimeMs, stopwatch.ElapsedMilliseconds);
            _logger?.LogDebug("Processed {Count} requests in {ElapsedMs}ms",
                processedCount, stopwatch.ElapsedMilliseconds);
        }

        return processedCount;
    }

    private void ProcessSingleRequest(GUIProxyRequest request)
    {
        if (request.CancellationTokenSource.Token.IsCancellationRequested)
        {
            var cancelledResponse = GUIProxyResponse.CreateError(
                request.RequestId,
                "Request was cancelled",
                "CancelledError");
            request.CompletionSource.TrySetResult(cancelledResponse);
            return;
        }

        if (!_handlers.TryGetValue(request.Action, out var handler))
        {
            var notFoundResponse = GUIProxyResponse.CreateError(
                request.RequestId,
                $"No handler registered for action: {request.Action}",
                "HandlerNotFoundError");
            request.CompletionSource.TrySetResult(notFoundResponse);
            Interlocked.Increment(ref _failedRequests);
            return;
        }

        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Execute handler synchronously on GUI thread
            var resultTask = handler(request.Parameters);

            // Wait for result (this is OK because we're on GUI thread and handlers should be quick)
            var result = resultTask.GetAwaiter().GetResult();

            stopwatch.Stop();

            var response = new GUIProxyResponse
            {
                RequestId = request.RequestId,
                Success = true,
                Result = result,
                Status = ProxyRequestStatus.Completed,
                StartTime = request.CreatedAt,
                EndTime = DateTime.UtcNow
            };

            request.CompletionSource.TrySetResult(response);
            Interlocked.Increment(ref _successfulRequests);

            RequestCompleted?.Invoke(this, response);

            _logger?.LogDebug("Request completed: {RequestId} in {ElapsedMs}ms",
                request.RequestId, stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();

            var errorResponse = new GUIProxyResponse
            {
                RequestId = request.RequestId,
                Success = false,
                ErrorMessage = ex.Message,
                ErrorType = ex.GetType().Name,
                Status = ProxyRequestStatus.Failed,
                StartTime = request.CreatedAt,
                EndTime = DateTime.UtcNow
            };

            request.CompletionSource.TrySetResult(errorResponse);
            Interlocked.Increment(ref _failedRequests);

            _logger?.LogError(ex, "Request failed: {RequestId} - {Action}",
                request.RequestId, request.Action);

            ErrorOccurred?.Invoke(this, ex);
        }
    }

    public GUIProxyStats GetStatistics()
    {
        var avgExecutionTime = _successfulRequests > 0
            ? TimeSpan.FromMilliseconds(_totalExecutionTimeMs / _successfulRequests)
            : TimeSpan.Zero;

        return new GUIProxyStats(
            TotalRequests: _totalRequests,
            SuccessfulRequests: _successfulRequests,
            FailedRequests: _failedRequests,
            TimedOutRequests: _timedOutRequests,
            PendingRequests: PendingRequestCount,
            AverageExecutionTime: avgExecutionTime);
    }

    public void ClearPendingRequests()
    {
        while (_requestQueue.TryDequeue(out var request))
        {
            request.CompletionSource.TrySetResult(
                GUIProxyResponse.CreateError(request.RequestId, "Request cleared", "ClearedError"));
        }

        foreach (var request in _pendingRequests.Values)
        {
            request.CancellationTokenSource.Cancel();
            request.CompletionSource.TrySetResult(
                GUIProxyResponse.CreateError(request.RequestId, "Request cleared", "ClearedError"));
        }

        _pendingRequests.Clear();

        _logger?.LogInformation("Cleared all pending requests");
    }

    public event EventHandler<GUIProxyResponse>? RequestCompleted;
    public event EventHandler<Exception>? ErrorOccurred;

    public void Dispose()
    {
        Stop();
        _handlers.Clear();

        GC.SuppressFinalize(this);
    }
}
