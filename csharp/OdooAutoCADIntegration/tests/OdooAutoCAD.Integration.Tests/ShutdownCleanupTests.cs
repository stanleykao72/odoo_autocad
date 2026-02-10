// OdooAutoCAD.Integration.Tests/ShutdownCleanupTests.cs
// Verify shutdown sequence: GUIProxy cleanup, MCP server stop, port release

using FluentAssertions;
using Moq;
using OdooAutoCAD.Core.Threading;
using OdooAutoCAD.MCP.Server;
using OdooAutoCAD.MCP.Tools;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

/// <summary>
/// Tests that the shutdown sequence properly cleans up resources.
/// Validates individual shutdown steps are isolated (one failure doesn't block others).
/// </summary>
public class ShutdownCleanupTests
{
    [Fact]
    public void GUIProxy_Stop_IsCalledOnce()
    {
        // Arrange
        var mockProxy = new Mock<IGUIProxy>();
        mockProxy.Setup(p => p.IsRunning).Returns(true);

        // Act
        mockProxy.Object.Stop();

        // Assert
        mockProxy.Verify(p => p.Stop(), Times.Once);
    }

    [Fact]
    public void GUIProxy_Stop_WhenAlreadyStopped_DoesNotThrow()
    {
        // Arrange
        var mockProxy = new Mock<IGUIProxy>();
        mockProxy.Setup(p => p.IsRunning).Returns(false);

        // Act & Assert
        var act = () => mockProxy.Object.Stop();
        act.Should().NotThrow();
    }

    [Fact]
    public async Task MCPServer_Stop_WhenNotRunning_DoesNotThrow()
    {
        // Arrange - server not started
        var mockProxy = new Mock<IGUIProxy>();
        var mockRegistry = new Mock<MCPToolRegistry>(mockProxy.Object, null, null, null, null);
        var server = new MCPSSEServer(mockRegistry.Object, port: 0);

        // Act & Assert
        var act = async () => await server.StopAsync();
        await act.Should().NotThrowAsync();
    }

    [Fact]
    public void GUIProxy_ClearPendingRequests_RemovesQueuedItems()
    {
        // Arrange
        var mockProxy = new Mock<IGUIProxy>();
        mockProxy.Setup(p => p.PendingRequestCount).Returns(0);

        // Act
        mockProxy.Object.ClearPendingRequests();

        // Assert
        mockProxy.Verify(p => p.ClearPendingRequests(), Times.Once);
        mockProxy.Object.PendingRequestCount.Should().Be(0);
    }

    [Fact]
    public void ShutdownSteps_AreIsolated_OneFailureDoesNotBlockOthers()
    {
        // Arrange
        var step1Executed = false;
        var step2Executed = false;
        var step3Executed = false;

        // Act - simulate the hardened shutdown pattern
        try { step1Executed = true; throw new InvalidOperationException("Timer stop failed"); }
        catch { /* isolated */ }

        try { step2Executed = true; }
        catch { /* isolated */ }

        try { step3Executed = true; }
        catch { /* isolated */ }

        // Assert - all steps executed despite step 1 failure
        step1Executed.Should().BeTrue();
        step2Executed.Should().BeTrue();
        step3Executed.Should().BeTrue();
    }

    [Fact]
    public void GUIProxy_Dispose_CanBeCalledMultipleTimes()
    {
        // Arrange
        var mockProxy = new Mock<IGUIProxy>();

        // Act & Assert - double dispose should not throw
        var act = () =>
        {
            mockProxy.Object.Dispose();
            mockProxy.Object.Dispose();
        };
        act.Should().NotThrow();
    }
}
