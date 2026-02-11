// OdooAutoCAD.Threading.Tests/GUIProxyTests.cs
// Unit tests for GUI Proxy

using FluentAssertions;
using OdooAutoCAD.Core.Threading;
using Xunit;

namespace OdooAutoCAD.Threading.Tests;

public class GUIProxyTests
{
    [Fact]
    public void NewProxy_ShouldNotBeRunning()
    {
        // Arrange & Act
        using var proxy = new GUIProxy();

        // Assert
        proxy.IsRunning.Should().BeFalse();
    }

    [Fact]
    public void Start_ShouldSetRunningToTrue()
    {
        // Arrange
        using var proxy = new GUIProxy();

        // Act
        proxy.Start();

        // Assert
        proxy.IsRunning.Should().BeTrue();
    }

    [Fact]
    public void Stop_ShouldSetRunningToFalse()
    {
        // Arrange
        using var proxy = new GUIProxy();
        proxy.Start();

        // Act
        proxy.Stop();

        // Assert
        proxy.IsRunning.Should().BeFalse();
    }

    [Fact]
    public void RegisterHandler_ShouldAddToRegisteredActions()
    {
        // Arrange
        using var proxy = new GUIProxy();

        // Act
        proxy.RegisterHandler("test_action", (p) => Task.FromResult<object?>("result"));

        // Assert
        proxy.GetRegisteredActions().Should().Contain("test_action");
    }

    [Fact]
    public void UnregisterHandler_ShouldRemoveFromRegisteredActions()
    {
        // Arrange
        using var proxy = new GUIProxy();
        proxy.RegisterHandler("test_action", (p) => Task.FromResult<object?>("result"));

        // Act
        var result = proxy.UnregisterHandler("test_action");

        // Assert
        result.Should().BeTrue();
        proxy.GetRegisteredActions().Should().NotContain("test_action");
    }

    [Fact]
    public async Task ExecuteInGuiAsync_WhenNotRunning_ShouldReturnError()
    {
        // Arrange
        using var proxy = new GUIProxy();
        // Not calling Start()

        // Act
        var response = await proxy.ExecuteInGuiAsync("test_action");

        // Assert
        response.Success.Should().BeFalse();
        response.ErrorMessage.Should().Contain("not running");
    }

    [Fact]
    public async Task ExecuteInGuiAsync_WhenHandlerNotRegistered_ShouldReturnError()
    {
        // Arrange
        using var proxy = new GUIProxy();
        proxy.Start();

        // Act
        var response = await proxy.ExecuteInGuiAsync("unknown_action");

        // Assert
        response.Success.Should().BeFalse();
        response.ErrorMessage.Should().Contain("No handler registered");
    }

    [Fact]
    public void GetStatistics_ShouldReturnZeroForNewProxy()
    {
        // Arrange
        using var proxy = new GUIProxy();

        // Act
        var stats = proxy.GetStatistics();

        // Assert
        stats.TotalRequests.Should().Be(0);
        stats.SuccessfulRequests.Should().Be(0);
        stats.FailedRequests.Should().Be(0);
    }

    [Fact]
    public void ProcessRequests_WhenQueueEmpty_ShouldReturnZero()
    {
        // Arrange
        using var proxy = new GUIProxy();
        proxy.Start();

        // Act
        var processed = proxy.ProcessRequests();

        // Assert
        processed.Should().Be(0);
    }

    [Fact]
    public void PendingRequestCount_ShouldStartAtZero()
    {
        // Arrange & Act
        using var proxy = new GUIProxy();

        // Assert
        proxy.PendingRequestCount.Should().Be(0);
    }

    [Fact]
    public void GUIProxyResponse_CreateSuccess_ShouldSetCorrectProperties()
    {
        // Arrange & Act
        var response = GUIProxyResponse.CreateSuccess("request-123", "test result");

        // Assert
        response.Success.Should().BeTrue();
        response.RequestId.Should().Be("request-123");
        response.Result.Should().Be("test result");
        response.Status.Should().Be(ProxyRequestStatus.Completed);
    }

    [Fact]
    public void GUIProxyResponse_CreateError_ShouldSetCorrectProperties()
    {
        // Arrange & Act
        var response = GUIProxyResponse.CreateError("request-456", "Error message", "ErrorType");

        // Assert
        response.Success.Should().BeFalse();
        response.RequestId.Should().Be("request-456");
        response.ErrorMessage.Should().Be("Error message");
        response.ErrorType.Should().Be("ErrorType");
        response.Status.Should().Be(ProxyRequestStatus.Failed);
    }

    [Fact]
    public void GUIProxyResponse_CreateTimeout_ShouldSetCorrectProperties()
    {
        // Arrange & Act
        var response = GUIProxyResponse.CreateTimeout("request-789");

        // Assert
        response.Success.Should().BeFalse();
        response.RequestId.Should().Be("request-789");
        response.Status.Should().Be(ProxyRequestStatus.TimedOut);
        response.ErrorType.Should().Be("TimeoutError");
    }

    [Fact]
    public void Dispose_ShouldStopProxy()
    {
        // Arrange
        var proxy = new GUIProxy();
        proxy.Start();

        // Act
        proxy.Dispose();

        // Assert
        proxy.IsRunning.Should().BeFalse();
    }

    [Fact]
    public async Task ProcessRequests_SyncHandler_ShouldCompleteViaFastPath()
    {
        // Arrange
        using var proxy = new GUIProxy();
        proxy.Start();
        proxy.RegisterHandler("fast_action", (p) => Task.FromResult<object?>("fast_result"));

        // Enqueue a request
        var responseTask = proxy.ExecuteInGuiAsync("fast_action");

        // Act — process the queue (simulates DispatcherTimer tick)
        var processed = proxy.ProcessRequests();

        // Assert
        processed.Should().Be(1);

        var response = await responseTask;
        response.Success.Should().BeTrue();
        response.Result.Should().Be("fast_result");
        response.Status.Should().Be(ProxyRequestStatus.Completed);
    }

    [Fact]
    public async Task ProcessRequests_FailingHandler_ShouldReturnError()
    {
        // Arrange
        using var proxy = new GUIProxy();
        proxy.Start();
        proxy.RegisterHandler("fail_action", (p) =>
            throw new InvalidOperationException("handler failed"));

        // Enqueue a request
        var responseTask = proxy.ExecuteInGuiAsync("fail_action");

        // Act
        proxy.ProcessRequests();

        // Assert
        var response = await responseTask;
        response.Success.Should().BeFalse();
        response.ErrorMessage.Should().Contain("handler failed");
    }

    [Fact]
    public async Task ProcessRequests_SlowPathHandler_ShouldCompleteAsynchronously()
    {
        // Arrange
        using var proxy = new GUIProxy();
        proxy.Start();
        proxy.RegisterHandler("slow_action", async (p) =>
        {
            await Task.Delay(50);
            return (object?)"slow_result";
        });

        // Enqueue a request
        var responseTask = proxy.ExecuteInGuiAsync("slow_action", timeout: 5000);

        // Act — process to start the slow task
        proxy.ProcessRequests();

        // Assert — the response completes eventually via ContinueWith
        var response = await responseTask;
        response.Success.Should().BeTrue();
        response.Result.Should().Be("slow_result");
    }
}
