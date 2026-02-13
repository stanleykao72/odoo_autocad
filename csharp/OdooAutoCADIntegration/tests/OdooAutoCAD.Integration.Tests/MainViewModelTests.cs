using FluentAssertions;
using Microsoft.Extensions.Logging;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;
using OdooAutoCAD.MCP.Server;
using OdooAutoCAD.MCP.Tools;
using System.Windows.Media;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

/// <summary>
/// Tests for MainViewModel sidebar connection status indicators.
/// </summary>
public class MainViewModelTests
{
    private readonly Mock<INavigationService> _mockNav;
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<IDrawingDataService> _mockDrawingDataService;
    private readonly MCPSSEServer _mcpServer;

    public MainViewModelTests()
    {
        _mockNav = new Mock<INavigationService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockOdoo = new Mock<IOdooService>();
        _mockDrawingDataService = new Mock<IDrawingDataService>();

        // Create a real MCPSSEServer with mock registry (not started)
        var mockRegistry = new Mock<MCPToolRegistry>(_mockGuiProxy.Object, null!, null!, null!, null!);
        _mcpServer = new MCPSSEServer(mockRegistry.Object, port: 0);
    }

    private MainViewModel CreateSUT()
    {
        return new MainViewModel(
            _mockNav.Object,
            _mockGuiProxy.Object,
            _mcpServer,
            _mockAutoCAD.Object,
            _mockOdoo.Object,
            _mockDrawingDataService.Object);
    }

    [StaFact]
    public void UpdateAutoCADStatus_WhenConnected_ShowsGreen()
    {
        // Arrange
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        var sut = CreateSUT();

        // Assert (constructor calls UpdateAllStatus)
        sut.AutoCADStatusText.Should().Be("Connected");
        sut.AutoCADStatusColor.Should().Be(Brushes.Green);
    }

    [StaFact]
    public void UpdateAutoCADStatus_WhenDisconnected_ShowsGray()
    {
        // Arrange
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(false);
        var sut = CreateSUT();

        // Assert
        sut.AutoCADStatusText.Should().Be("Disconnected");
        sut.AutoCADStatusColor.Should().Be(Brushes.Gray);
    }

    [StaFact]
    public void AutoCADModeText_WhenCOM_ShowsCOM()
    {
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.COM);
        var sut = CreateSUT();

        sut.AutoCADModeText.Should().Be("(COM)");
    }

    [StaFact]
    public void AutoCADModeText_WhenFile_ShowsFile()
    {
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.File);
        var sut = CreateSUT();

        sut.AutoCADModeText.Should().Be("(File)");
    }

    [StaFact]
    public void UpdateOdooStatus_WhenConnected_ShowsGreen()
    {
        // Arrange
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();

        // Assert
        sut.OdooStatusText.Should().Be("Connected");
        sut.OdooStatusColor.Should().Be(Brushes.Green);
    }

    [StaFact]
    public void UpdateOdooStatus_WhenDisconnected_ShowsGray()
    {
        // Arrange
        _mockOdoo.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        // Assert
        sut.OdooStatusText.Should().Be("Disconnected");
        sut.OdooStatusColor.Should().Be(Brushes.Gray);
    }
}
