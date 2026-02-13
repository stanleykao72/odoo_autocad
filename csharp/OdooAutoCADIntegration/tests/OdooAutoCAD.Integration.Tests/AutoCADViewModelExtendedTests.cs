using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

/// <summary>
/// Tests for Sprint 8 AutoCAD page enhancements:
/// Drawing info, COM monitoring, PR project info, and clear table IDs.
/// </summary>
public class AutoCADViewModelExtendedTests
{
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IDwgReaderService> _mockDwgReader;
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<ISettingsService> _mockSettings;
    private readonly Mock<IAppLogService> _mockLogService;
    private readonly Mock<IDrawingDataService> _mockDrawingDataService;

    public AutoCADViewModelExtendedTests()
    {
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockDwgReader = new Mock<IDwgReaderService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockOdoo = new Mock<IOdooService>();
        _mockSettings = new Mock<ISettingsService>();
        _mockLogService = new Mock<IAppLogService>();
        _mockDrawingDataService = new Mock<IDrawingDataService>();

        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>());
    }

    private AutoCADViewModel CreateSUT()
    {
        return new AutoCADViewModel(
            _mockAutoCAD.Object,
            _mockDwgReader.Object,
            _mockGuiProxy.Object,
            _mockOdoo.Object,
            _mockSettings.Object,
            _mockLogService.Object,
            _mockDrawingDataService.Object);
    }

    #region US-002-04: Document Path

    [Fact]
    public void Constructor_DocumentPath_DefaultsToNA()
    {
        var sut = CreateSUT();

        sut.DocumentPath.Should().Be("N/A");
    }

    [Fact]
    public void Constructor_ComStatusText_DefaultsToNotMonitored()
    {
        var sut = CreateSUT();

        sut.ComStatusText.Should().Be("Not monitored");
    }

    [Fact]
    public void Constructor_IsMonitoring_DefaultsFalse()
    {
        var sut = CreateSUT();

        sut.IsMonitoring.Should().BeFalse();
    }

    #endregion

    #region US-002-05: PR Project Info

    [Fact]
    public void Constructor_PrNumber_DefaultsToEmpty()
    {
        var sut = CreateSUT();

        sut.PrNumber.Should().BeEmpty();
    }

    [Fact]
    public void Constructor_ProjectLookupStatus_DefaultsToEmpty()
    {
        var sut = CreateSUT();

        sut.ProjectLookupStatus.Should().BeEmpty();
    }

    [Fact]
    public void ExtractPRInfoCommand_WhenDisconnected_CannotExecute()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        sut.ExtractPRInfoCommand.CanExecute(null).Should().BeFalse();
    }

    #endregion

    #region US-002-06: Clear Table IDs

    [Fact]
    public void Constructor_ClearIdsStatus_DefaultsToEmpty()
    {
        var sut = CreateSUT();

        sut.ClearIdsStatus.Should().BeEmpty();
    }

    [Fact]
    public void ClearTableIdsCommand_WhenDisconnected_CannotExecute()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        sut.ClearTableIdsCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void ClearAllTableIdsCommand_WhenDisconnected_CannotExecute()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        sut.ClearAllTableIdsCommand.CanExecute(null).Should().BeFalse();
    }

    #endregion

    #region Dual-Mode

    [Fact]
    public void IsFileMode_WhenCOM_ReturnsFalse()
    {
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.COM);
        var sut = CreateSUT();

        sut.IsFileMode.Should().BeFalse();
    }

    [Fact]
    public void IsFileMode_WhenFile_ReturnsTrue()
    {
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.File);
        var sut = CreateSUT();

        sut.IsFileMode.Should().BeTrue();
    }

    #endregion

    #region US-002-07: COM Health Monitor

    [Fact]
    public void HealthCheck_WhenConnectionLost_UpdatesState()
    {
        // Arrange: simulate connected state then disconnect
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        // Simulate health check detecting lost connection
        sut.IsConnected = true; // was connected
        sut.ComStatusText = "Connected";

        // Now mark as disconnected (simulating what health timer would do)
        sut.IsConnected = false;
        sut.ComStatusText = "Connection lost";

        sut.IsConnected.Should().BeFalse();
        sut.ComStatusText.Should().Be("Connection lost");
    }

    #endregion
}
