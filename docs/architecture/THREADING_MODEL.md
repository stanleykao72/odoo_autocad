# Threading Model Design Document

> **Version**: 1.0
> **Last Updated**: February 2026
> **Applies To**: OdooAutoCAD C# Implementation

## Table of Contents

1. [Overview](#overview)
2. [COM Threading Requirements](#com-threading-requirements)
3. [GUI Proxy Pattern](#gui-proxy-pattern)
4. [Request/Response Queue Design](#requestresponse-queue-design)
5. [Timeout and Error Handling](#timeout-and-error-handling)
6. [Thread Architecture Diagram](#thread-architecture-diagram)
7. [Implementation Details](#implementation-details)
8. [Best Practices](#best-practices)

---

## Overview

The OdooAutoCAD application must handle a complex threading scenario:

1. **WPF GUI Thread (STA)**: Must handle all UI operations and COM calls to AutoCAD
2. **MCP Server Thread(s) (MTA)**: Handle incoming SSE connections and JSON-RPC requests
3. **Background Workers**: Handle long-running operations like Odoo API calls

The primary challenge is that AutoCAD COM objects require Single-Threaded Apartment (STA) access, while HTTP server operations typically run on ThreadPool threads (MTA). The GUI Proxy pattern solves this by marshaling all COM operations to the GUI thread.

---

## COM Threading Requirements

### Understanding COM Apartments

```
+------------------+     +------------------+
|   STA Thread     |     |   MTA Threads    |
|  (GUI Thread)    |     |  (ThreadPool)    |
+------------------+     +------------------+
|                  |     |                  |
| - WPF Dispatcher |     | - HTTP Requests  |
| - AutoCAD COM    |     | - Background I/O |
| - UI Updates     |     | - Async Tasks    |
|                  |     |                  |
+------------------+     +------------------+
         |                        |
         |    COM Marshaling      |
         |<-----------------------|
         |   (RPC/Proxy calls)    |
         |                        |
```

### STA Requirements for AutoCAD

AutoCAD's COM objects are designed for STA threads. Accessing them from MTA threads causes:

- **RPC_E_WRONG_THREAD** errors
- COM proxy creation overhead
- Potential deadlocks
- Unpredictable behavior

### Solution: Thread Affinity

All AutoCAD COM operations MUST execute on the same STA thread where the COM connection was established. The GUI Proxy pattern enforces this by:

1. Creating the AutoCAD COM connection on the GUI thread
2. Routing all COM operations through the GUI thread's Dispatcher
3. Using message queues to communicate between threads

---

## GUI Proxy Pattern

### Core Concept

The GUI Proxy acts as an intermediary between background threads (MCP, async operations) and the GUI thread where COM operations are safe.

```
+-------------------+
|   MCP Request     |
|   (MTA Thread)    |
+---------+---------+
          |
          | 1. Create ProxyRequest
          v
+-------------------+
|   Request Queue   |
|   (Channel<T>)    |
+---------+---------+
          |
          | 2. GUI Thread polls queue
          v
+-------------------+
|   GUI Dispatcher  |
|   (STA Thread)    |
+---------+---------+
          |
          | 3. Execute COM operation
          v
+-------------------+
|   AutoCAD COM     |
|   Operation       |
+---------+---------+
          |
          | 4. Store result
          v
+-------------------+
|  Response Queue   |
|  (per-request)    |
+---------+---------+
          |
          | 5. Caller receives result
          v
+-------------------+
|   MCP Response    |
|   (MTA Thread)    |
+-------------------+
```

### Key Design Decisions

| Aspect | Decision | Rationale |
|--------|----------|-----------|
| Queue Type | `Channel<T>` | High-performance, thread-safe, async-friendly |
| Request ID | `Guid` | Unique correlation for responses |
| Polling Interval | 50ms | Balance between responsiveness and CPU usage |
| Default Timeout | 30 seconds | Allow for slow AutoCAD operations |

---

## Request/Response Queue Design

### ProxyRequest Structure

```csharp
namespace OdooAutoCAD.GuiProxy;

/// <summary>
/// Represents a request to execute an operation on the GUI thread.
/// </summary>
public sealed class ProxyRequest
{
    /// <summary>
    /// Unique identifier for request/response correlation.
    /// </summary>
    public Guid RequestId { get; } = Guid.NewGuid();

    /// <summary>
    /// The operation to execute on the GUI thread.
    /// </summary>
    public Func<object?> Operation { get; init; } = null!;

    /// <summary>
    /// Maximum time to wait for the operation to complete.
    /// </summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Timestamp when the request was created.
    /// </summary>
    public DateTime CreatedAt { get; } = DateTime.UtcNow;

    /// <summary>
    /// Cancellation token to cancel the operation.
    /// </summary>
    public CancellationToken CancellationToken { get; init; }

    /// <summary>
    /// Task completion source for the response.
    /// </summary>
    internal TaskCompletionSource<ProxyResponse> ResponseTcs { get; } = new();
}
```

### ProxyResponse Structure

```csharp
namespace OdooAutoCAD.GuiProxy;

/// <summary>
/// Represents the result of a GUI proxy operation.
/// </summary>
public sealed class ProxyResponse
{
    /// <summary>
    /// The request ID this response corresponds to.
    /// </summary>
    public Guid RequestId { get; init; }

    /// <summary>
    /// Indicates whether the operation completed successfully.
    /// </summary>
    public bool IsSuccess { get; init; }

    /// <summary>
    /// The result of the operation (if successful).
    /// </summary>
    public object? Result { get; init; }

    /// <summary>
    /// The exception that occurred (if failed).
    /// </summary>
    public Exception? Exception { get; init; }

    /// <summary>
    /// Time taken to execute the operation.
    /// </summary>
    public TimeSpan ExecutionTime { get; init; }

    /// <summary>
    /// Creates a successful response.
    /// </summary>
    public static ProxyResponse Success(Guid requestId, object? result, TimeSpan executionTime)
        => new()
        {
            RequestId = requestId,
            IsSuccess = true,
            Result = result,
            ExecutionTime = executionTime
        };

    /// <summary>
    /// Creates a failed response.
    /// </summary>
    public static ProxyResponse Failure(Guid requestId, Exception exception, TimeSpan executionTime)
        => new()
        {
            RequestId = requestId,
            IsSuccess = false,
            Exception = exception,
            ExecutionTime = executionTime
        };
}
```

### Queue Implementation

```csharp
namespace OdooAutoCAD.GuiProxy;

using System.Threading.Channels;

/// <summary>
/// Manages the request queue for GUI proxy operations.
/// </summary>
public sealed class MessageQueue : IDisposable
{
    private readonly Channel<ProxyRequest> _requestChannel;
    private readonly ILogger<MessageQueue> _logger;
    private bool _disposed;

    public MessageQueue(ILogger<MessageQueue> logger)
    {
        _logger = logger;

        // Unbounded channel with single reader (GUI thread)
        _requestChannel = Channel.CreateUnbounded<ProxyRequest>(
            new UnboundedChannelOptions
            {
                SingleReader = true,
                SingleWriter = false
            });
    }

    /// <summary>
    /// Enqueues a request and returns a task that completes when the response is ready.
    /// </summary>
    public async Task<ProxyResponse> EnqueueAsync(
        ProxyRequest request,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(MessageQueue));

        _logger.LogDebug("Enqueuing request {RequestId}", request.RequestId);

        // Write request to channel
        await _requestChannel.Writer.WriteAsync(request, cancellationToken);

        // Wait for response with timeout
        using var timeoutCts = new CancellationTokenSource(request.Timeout);
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, timeoutCts.Token);

        try
        {
            return await request.ResponseTcs.Task.WaitAsync(linkedCts.Token);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
        {
            throw new GuiProxyTimeoutException(request.Timeout);
        }
    }

    /// <summary>
    /// Attempts to read the next request from the queue.
    /// </summary>
    public bool TryRead(out ProxyRequest? request)
    {
        return _requestChannel.Reader.TryRead(out request);
    }

    /// <summary>
    /// Waits for a request to become available.
    /// </summary>
    public ValueTask<bool> WaitToReadAsync(CancellationToken cancellationToken = default)
    {
        return _requestChannel.Reader.WaitToReadAsync(cancellationToken);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _requestChannel.Writer.Complete();
    }
}
```

---

## Timeout and Error Handling

### Timeout Strategy

```
Request Lifecycle:

 0s        10s       20s       30s
 |---------|---------|---------|
 ^         ^         ^         ^
 |         |         |         |
 Create    Warning   Warning   Timeout
 Request   Log       Log       Exception
```

### Timeout Configuration

```csharp
public class GuiProxyOptions
{
    /// <summary>
    /// Default timeout for proxy operations.
    /// </summary>
    public TimeSpan DefaultTimeout { get; set; } = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Extended timeout for heavy operations (e.g., opening drawings).
    /// </summary>
    public TimeSpan ExtendedTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Interval for checking the request queue on the GUI thread.
    /// </summary>
    public TimeSpan PollingInterval { get; set; } = TimeSpan.FromMilliseconds(50);

    /// <summary>
    /// Maximum number of pending requests before throttling.
    /// </summary>
    public int MaxPendingRequests { get; set; } = 100;

    /// <summary>
    /// Log warning if operation takes longer than this.
    /// </summary>
    public TimeSpan SlowOperationThreshold { get; set; } = TimeSpan.FromSeconds(5);
}
```

### Error Handling Flow

```csharp
public async Task ProcessRequestAsync(ProxyRequest request)
{
    var stopwatch = Stopwatch.StartNew();

    try
    {
        // Check for cancellation before starting
        request.CancellationToken.ThrowIfCancellationRequested();

        // Execute the operation
        var result = request.Operation();

        stopwatch.Stop();

        // Log slow operations
        if (stopwatch.Elapsed > _options.SlowOperationThreshold)
        {
            _logger.LogWarning(
                "Slow operation {RequestId} took {Elapsed:F2}s",
                request.RequestId,
                stopwatch.Elapsed.TotalSeconds);
        }

        // Complete with success
        request.ResponseTcs.SetResult(
            ProxyResponse.Success(request.RequestId, result, stopwatch.Elapsed));
    }
    catch (OperationCanceledException)
    {
        stopwatch.Stop();
        request.ResponseTcs.SetCanceled();

        _logger.LogInformation(
            "Request {RequestId} was cancelled after {Elapsed:F2}s",
            request.RequestId,
            stopwatch.Elapsed.TotalSeconds);
    }
    catch (COMException ex)
    {
        stopwatch.Stop();
        var wrappedException = new AutoCADComException(
            $"COM operation failed: {ex.Message}", ex);

        request.ResponseTcs.SetResult(
            ProxyResponse.Failure(request.RequestId, wrappedException, stopwatch.Elapsed));

        _logger.LogError(ex,
            "COM error in request {RequestId}: HRESULT={HResult}",
            request.RequestId,
            ex.HResult);
    }
    catch (Exception ex)
    {
        stopwatch.Stop();
        request.ResponseTcs.SetResult(
            ProxyResponse.Failure(request.RequestId, ex, stopwatch.Elapsed));

        _logger.LogError(ex,
            "Unexpected error in request {RequestId}",
            request.RequestId);
    }
}
```

---

## Thread Architecture Diagram

```
+===========================================================================+
|                        APPLICATION ARCHITECTURE                            |
+===========================================================================+

                    +----------------------------------+
                    |         WPF Application          |
                    |          (Process)               |
                    +----------------------------------+
                                   |
            +----------------------+----------------------+
            |                      |                      |
            v                      v                      v
    +---------------+      +---------------+      +---------------+
    |  GUI Thread   |      | MCP Server    |      | Odoo Client   |
    |    (STA)      |      | Thread Pool   |      | Thread Pool   |
    |               |      |    (MTA)      |      |    (MTA)      |
    +---------------+      +---------------+      +---------------+
    |               |      |               |      |               |
    | - WPF UI      |      | - SSE Handler |      | - HTTP Client |
    | - DispatcherTimer  | | - JSON-RPC    |      | - REST Calls  |
    | - COM Objects |      | - Tool Exec   |      | - Auth        |
    | - GUI Proxy   |      |               |      |               |
    |   Processor   |      |               |      |               |
    +-------+-------+      +-------+-------+      +---------------+
            ^                      |
            |                      |
            |    +-------------+   |
            |    |   Request   |   |
            +----+    Queue    +<--+
            |    | (Channel<T>)|
            |    +-------------+
            |
            v
    +---------------+
    |  AutoCAD      |
    |  COM Objects  |
    |               |
    | - Application |
    | - Document    |
    | - ModelSpace  |
    +---------------+


    THREAD FLOW DETAIL:
    ====================

    [MCP Request Arrives]
           |
           v
    +------------------+
    | ASP.NET Core     |
    | ThreadPool       |
    | (MTA Thread #N)  |
    +--------+---------+
             |
             | 1. Parse JSON-RPC
             | 2. Identify tool
             | 3. Need COM access?
             |
             +---> NO: Execute directly, return result
             |
             v YES
    +------------------+
    | GuiProxyService  |
    | .ExecuteAsync()  |
    +--------+---------+
             |
             | 4. Create ProxyRequest
             | 5. Write to Channel
             | 6. await ResponseTcs.Task
             |
             v
    +------------------+
    |  Request Queue   |
    |  Channel<Request>|
    +--------+---------+
             |
             | (Async boundary)
             |
             v
    +------------------+     50ms polling
    | GUI Thread       +<==================+
    | DispatcherTimer  |                   |
    +--------+---------+                   |
             |                             |
             | 7. TryRead from Channel     |
             | 8. Execute Operation        |
             |                             |
             v                             |
    +------------------+                   |
    | AutoCAD COM      |                   |
    | (Same STA)       |                   |
    +--------+---------+                   |
             |                             |
             | 9. Get result/exception     |
             | 10. SetResult on TCS        |
             |                             |
             +-----------------------------+
             |
             v
    +------------------+
    | ProxyResponse    |
    | returned to      |
    | MCP thread       |
    +------------------+
             |
             | 11. Serialize result
             | 12. Send SSE response
             |
             v
    [Response to Client]
```

---

## Implementation Details

### GuiProxyService Complete Implementation

```csharp
namespace OdooAutoCAD.GuiProxy;

using System.Windows.Threading;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

/// <summary>
/// Service that proxies operations from background threads to the GUI thread
/// for safe COM object access.
/// </summary>
public sealed class GuiProxyService : IGuiProxy, IDisposable
{
    private readonly MessageQueue _messageQueue;
    private readonly Dispatcher _dispatcher;
    private readonly ILogger<GuiProxyService> _logger;
    private readonly GuiProxyOptions _options;
    private readonly CancellationTokenSource _shutdownCts;

    private DispatcherTimer? _processingTimer;
    private bool _isRunning;
    private bool _disposed;
    private int _pendingRequestCount;

    public bool IsRunning => _isRunning;

    public GuiProxyService(
        MessageQueue messageQueue,
        Dispatcher dispatcher,
        ILogger<GuiProxyService> logger,
        IOptions<GuiProxyOptions> options)
    {
        _messageQueue = messageQueue;
        _dispatcher = dispatcher;
        _logger = logger;
        _options = options.Value;
        _shutdownCts = new CancellationTokenSource();
    }

    /// <summary>
    /// Starts the GUI proxy message processing loop.
    /// Must be called from the GUI thread.
    /// </summary>
    public void Start()
    {
        if (_isRunning)
        {
            _logger.LogWarning("GUI Proxy is already running");
            return;
        }

        _dispatcher.VerifyAccess(); // Ensure we're on GUI thread

        _processingTimer = new DispatcherTimer(
            _options.PollingInterval,
            DispatcherPriority.Normal,
            ProcessPendingRequests,
            _dispatcher);

        _processingTimer.Start();
        _isRunning = true;

        _logger.LogInformation(
            "GUI Proxy started with {Interval}ms polling interval",
            _options.PollingInterval.TotalMilliseconds);
    }

    /// <summary>
    /// Stops the GUI proxy and cancels pending requests.
    /// </summary>
    public async Task StopAsync()
    {
        if (!_isRunning) return;

        _logger.LogInformation("Stopping GUI Proxy...");

        _shutdownCts.Cancel();
        _processingTimer?.Stop();
        _isRunning = false;

        // Give pending requests a chance to complete
        await Task.Delay(100);

        _logger.LogInformation(
            "GUI Proxy stopped. Pending requests: {Count}",
            _pendingRequestCount);
    }

    /// <summary>
    /// Executes an action on the GUI thread.
    /// </summary>
    public Task ExecuteAsync(
        Action action,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        return ExecuteAsync<object?>(() =>
        {
            action();
            return null;
        }, timeout, cancellationToken);
    }

    /// <summary>
    /// Executes a function on the GUI thread and returns the result.
    /// </summary>
    public async Task<T> ExecuteAsync<T>(
        Func<T> func,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        ThrowIfNotRunning();

        var request = new ProxyRequest
        {
            Operation = () => func()!,
            Timeout = timeout ?? _options.DefaultTimeout,
            CancellationToken = cancellationToken
        };

        Interlocked.Increment(ref _pendingRequestCount);

        try
        {
            var response = await _messageQueue.EnqueueAsync(request, cancellationToken);

            if (!response.IsSuccess)
            {
                throw response.Exception ?? new InvalidOperationException("Unknown error");
            }

            return (T)response.Result!;
        }
        finally
        {
            Interlocked.Decrement(ref _pendingRequestCount);
        }
    }

    /// <summary>
    /// Executes an async function on the GUI thread.
    /// </summary>
    public async Task<T> ExecuteAsync<T>(
        Func<Task<T>> asyncFunc,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        ThrowIfNotRunning();

        // For async operations, we need to handle differently
        // The actual async work might need to return to this thread
        var tcs = new TaskCompletionSource<T>();

        var request = new ProxyRequest
        {
            Operation = () =>
            {
                // Start the async operation on the GUI thread
                // It will complete asynchronously
                asyncFunc().ContinueWith(task =>
                {
                    if (task.IsFaulted)
                        tcs.SetException(task.Exception!.InnerExceptions);
                    else if (task.IsCanceled)
                        tcs.SetCanceled();
                    else
                        tcs.SetResult(task.Result);
                }, TaskScheduler.FromCurrentSynchronizationContext());

                return null;
            },
            Timeout = timeout ?? _options.DefaultTimeout,
            CancellationToken = cancellationToken
        };

        Interlocked.Increment(ref _pendingRequestCount);

        try
        {
            // Enqueue to start the operation
            await _messageQueue.EnqueueAsync(request, cancellationToken);

            // Wait for the async operation to complete
            using var timeoutCts = new CancellationTokenSource(request.Timeout);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken, timeoutCts.Token);

            return await tcs.Task.WaitAsync(linkedCts.Token);
        }
        finally
        {
            Interlocked.Decrement(ref _pendingRequestCount);
        }
    }

    private void ProcessPendingRequests(object? sender, EventArgs e)
    {
        // Process all available requests in this tick
        int processed = 0;
        const int maxPerTick = 10; // Prevent UI freeze

        while (_messageQueue.TryRead(out var request) && processed < maxPerTick)
        {
            if (request is null) continue;

            ProcessRequest(request);
            processed++;
        }
    }

    private void ProcessRequest(ProxyRequest request)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            request.CancellationToken.ThrowIfCancellationRequested();

            var result = request.Operation();

            stopwatch.Stop();
            LogOperationTiming(request, stopwatch.Elapsed);

            request.ResponseTcs.TrySetResult(
                ProxyResponse.Success(request.RequestId, result, stopwatch.Elapsed));
        }
        catch (OperationCanceledException)
        {
            request.ResponseTcs.TrySetCanceled();
        }
        catch (COMException ex)
        {
            stopwatch.Stop();
            _logger.LogError(ex,
                "COM error processing request {RequestId}",
                request.RequestId);

            request.ResponseTcs.TrySetResult(
                ProxyResponse.Failure(
                    request.RequestId,
                    new AutoCADComException(ex.Message, ex),
                    stopwatch.Elapsed));
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.LogError(ex,
                "Error processing request {RequestId}",
                request.RequestId);

            request.ResponseTcs.TrySetResult(
                ProxyResponse.Failure(request.RequestId, ex, stopwatch.Elapsed));
        }
    }

    private void LogOperationTiming(ProxyRequest request, TimeSpan elapsed)
    {
        if (elapsed > _options.SlowOperationThreshold)
        {
            _logger.LogWarning(
                "Slow GUI proxy operation {RequestId} took {Elapsed:F2}s",
                request.RequestId,
                elapsed.TotalSeconds);
        }
        else
        {
            _logger.LogDebug(
                "Request {RequestId} completed in {Elapsed:F3}s",
                request.RequestId,
                elapsed.TotalSeconds);
        }
    }

    private void ThrowIfNotRunning()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(GuiProxyService));

        if (!_isRunning)
            throw new InvalidOperationException("GUI Proxy is not running");
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _shutdownCts.Cancel();
        _processingTimer?.Stop();
        _messageQueue.Dispose();
        _shutdownCts.Dispose();
    }
}
```

---

## Best Practices

### DO

1. **Always use GUI Proxy for COM operations**
   ```csharp
   // Correct
   var result = await _guiProxy.ExecuteAsync(() => _autocad.GetDocumentName());
   ```

2. **Use appropriate timeouts**
   ```csharp
   // For heavy operations
   var result = await _guiProxy.ExecuteAsync(
       () => _autocad.OpenDrawing(path),
       timeout: TimeSpan.FromMinutes(2));
   ```

3. **Handle cancellation properly**
   ```csharp
   try
   {
       await _guiProxy.ExecuteAsync(operation, cancellationToken: ct);
   }
   catch (OperationCanceledException)
   {
       // Handle graceful cancellation
   }
   ```

4. **Log slow operations**
   ```csharp
   var sw = Stopwatch.StartNew();
   var result = await _guiProxy.ExecuteAsync(operation);
   if (sw.Elapsed > TimeSpan.FromSeconds(5))
       _logger.LogWarning("Slow operation: {Elapsed}", sw.Elapsed);
   ```

### DON'T

1. **Never access COM objects directly from MTA threads**
   ```csharp
   // WRONG - Will cause threading issues
   Task.Run(() => _autocadApp.ActiveDocument.Name);

   // Correct
   await _guiProxy.ExecuteAsync(() => _autocadApp.ActiveDocument.Name);
   ```

2. **Never hold COM references across threads**
   ```csharp
   // WRONG - Reference may become invalid
   var doc = await _guiProxy.ExecuteAsync(() => _autocad.ActiveDocument);
   var name = doc.Name; // May fail!

   // Correct - Get everything you need in one call
   var name = await _guiProxy.ExecuteAsync(() => _autocad.ActiveDocument?.Name);
   ```

3. **Never block the GUI thread**
   ```csharp
   // WRONG - Blocks UI
   _dispatcher.Invoke(() => Thread.Sleep(5000));

   // Correct - Use async
   await Task.Delay(5000);
   ```

4. **Never ignore exceptions from COM operations**
   ```csharp
   // WRONG - Swallows important errors
   try { await _guiProxy.ExecuteAsync(op); } catch { }

   // Correct - Handle appropriately
   try { await _guiProxy.ExecuteAsync(op); }
   catch (AutoCADComException ex) { _logger.LogError(ex, "COM failed"); throw; }
   ```

---

## Performance Considerations

### Metrics to Monitor

| Metric | Target | Warning Threshold |
|--------|--------|-------------------|
| Queue Depth | < 10 | > 50 |
| Average Latency | < 100ms | > 500ms |
| Timeout Rate | < 0.1% | > 1% |
| GUI Thread Usage | < 30% | > 70% |

### Optimization Tips

1. **Batch operations when possible**
   ```csharp
   // Instead of multiple calls
   var name = await _proxy.ExecuteAsync(() => doc.Name);
   var path = await _proxy.ExecuteAsync(() => doc.FullName);

   // Batch into one call
   var (name, path) = await _proxy.ExecuteAsync(() =>
       (doc.Name, doc.FullName));
   ```

2. **Use caching for frequently accessed data**
   ```csharp
   private string? _cachedDrawingPath;

   public async Task<string?> GetCurrentDrawingPath()
   {
       _cachedDrawingPath ??= await _guiProxy.ExecuteAsync(
           () => _autocad.ActiveDocument?.FullName);
       return _cachedDrawingPath;
   }
   ```

3. **Consider request coalescing for rapid updates**
   ```csharp
   // Debounce rapid requests
   private readonly Debouncer _debouncer = new(TimeSpan.FromMilliseconds(100));

   public Task RefreshStatusAsync()
   {
       return _debouncer.ExecuteAsync(() =>
           _guiProxy.ExecuteAsync(() => _autocad.GetStatus()));
   }
   ```

---

*Document Version: 1.0 | Created: February 2026*
