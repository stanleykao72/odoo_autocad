using FluentAssertions;
using Microsoft.Extensions.Configuration;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Configuration;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;
using System.Windows.Media;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

/// <summary>
/// Tests for DashboardViewModel: connection status display,
/// connect commands, navigation commands, and status polling.
/// </summary>
public class DashboardViewModelTests
{
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<INavigationService> _mockNav;
    private readonly Mock<ISettingsService> _mockSettings;
    private readonly Mock<IConfiguration> _mockConfiguration;
    private readonly Mock<IAppLogService> _mockLogService;

    public DashboardViewModelTests()
    {
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockOdoo = new Mock<IOdooService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockNav = new Mock<INavigationService>();
        _mockSettings = new Mock<ISettingsService>();
        _mockConfiguration = new Mock<IConfiguration>();
        _mockLogService = new Mock<IAppLogService>();

        // Default: both disconnected
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        _mockOdoo.Setup(s => s.IsConnected).Returns(false);

        // Default empty server configs
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>());
        _mockSettings.Setup(s => s.GetAppSettings()).Returns(new AppSettings
        {
            Application = new ApplicationInfo { Name = "Test", Version = "1.0" },
            Odoo = new OdooSettings { SwaggerUrl = "" },
            AutoCAD = new AutoCADSettings(),
            MCP = new MCPSettings(),
            Database = new DatabaseSettings()
        });
    }

    private DashboardViewModel CreateSUT()
    {
        return new DashboardViewModel(
            _mockAutoCAD.Object,
            _mockOdoo.Object,
            _mockGuiProxy.Object,
            _mockNav.Object,
            _mockSettings.Object,
            _mockConfiguration.Object,
            _mockLogService.Object);
    }

    [Fact]
    public void Constructor_SetsDefaultDisconnectedState()
    {
        // Act
        var sut = CreateSUT();

        // Assert
        sut.IsAutoCADConnected.Should().BeFalse();
        sut.IsOdooConnected.Should().BeFalse();
        sut.AutoCADStatusText.Should().Be("Disconnected");
        sut.OdooStatusText.Should().Be("Disconnected");
        sut.AutoCADErrorMessage.Should().BeEmpty();
        sut.OdooErrorMessage.Should().BeEmpty();
        sut.ConnectAutoCADButtonText.Should().Be("Connect");
        sut.ConnectOdooButtonText.Should().Be("Connect");
    }

    [StaFact]
    public async Task ConnectAutoCADAsync_WhenSuccessful_SetsConnectedState()
    {
        // Arrange — ConnectAsync is now called directly (bypasses GUIProxy)
        _mockAutoCAD.Setup(s => s.ConnectAsync()).ReturnsAsync(true);

        var status = new AutoCADStatus(true, "AutoCAD", "2025", "Drawing1.dwg", null, null);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_status", null, 5000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req2", status));

        var sut = CreateSUT();

        // Act
        await sut.ConnectAutoCADCommand.ExecuteAsync(null);

        // Assert
        sut.IsAutoCADConnected.Should().BeTrue();
        sut.AutoCADDocumentName.Should().Be("Drawing1.dwg");
        sut.AutoCADErrorMessage.Should().BeEmpty();
        sut.IsConnectingAutoCAD.Should().BeFalse();
    }

    [StaFact]
    public async Task ConnectAutoCADAsync_WhenFails_SetsErrorMessage()
    {
        // Arrange — ConnectAsync is now called directly (bypasses GUIProxy)
        _mockAutoCAD.Setup(s => s.ConnectAsync()).ReturnsAsync(false);

        var sut = CreateSUT();

        // Act
        await sut.ConnectAutoCADCommand.ExecuteAsync(null);

        // Assert
        sut.IsAutoCADConnected.Should().BeFalse();
        sut.AutoCADErrorMessage.Should().Contain("Unable to connect to AutoCAD");
    }

    [StaFact]
    public async Task ConnectOdooAsync_WhenNoSettings_SetsErrorMessage()
    {
        // Arrange - default empty settings (no swagger URL or token)
        var sut = CreateSUT();

        // Act
        await sut.ConnectOdooCommand.ExecuteAsync(null);

        // Assert
        sut.IsOdooConnected.Should().BeFalse();
        sut.OdooErrorMessage.Should().Contain("No Odoo connection settings found");
    }

    [StaFact]
    public async Task ConnectOdooAsync_WhenInvalidSwaggerUrl_SetsErrorMessage()
    {
        // Arrange - URL without required token/db query params
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://odoo.example.com/invalid-url",
                ["odoo_user_token"] = "test-token"
            });

        var sut = CreateSUT();

        // Act
        await sut.ConnectOdooCommand.ExecuteAsync(null);

        // Assert
        sut.IsOdooConnected.Should().BeFalse();
        sut.OdooErrorMessage.Should().Contain("Invalid Swagger URL");
    }

    [StaFact]
    public async Task ConnectOdooAsync_WhenServerUnreachable_SetsErrorMessage()
    {
        // Arrange - valid URL format but server won't be reachable
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://unreachable.invalid/api/v1/boq_import_api/swagger.json?token=abc&db=testdb",
                ["odoo_user_token"] = "test-token"
            });

        var sut = CreateSUT();

        // Act
        await sut.ConnectOdooCommand.ExecuteAsync(null);

        // Assert
        sut.IsOdooConnected.Should().BeFalse();
        sut.OdooErrorMessage.Should().NotBeEmpty();
        sut.IsConnectingOdoo.Should().BeFalse();
    }

    [StaFact]
    public void CanConnectAutoCAD_WhenAlreadyConnected_ReturnsFalse()
    {
        // Arrange
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();
        // Force the connected state via UpdateStatusFromServices
        sut.UpdateStatusFromServices();

        // Assert
        sut.ConnectAutoCADCommand.CanExecute(null).Should().BeFalse();
    }

    [StaFact]
    public void CanConnectOdoo_WhenAlreadyConnected_ReturnsFalse()
    {
        // Arrange
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();
        sut.UpdateStatusFromServices();

        // Assert
        sut.ConnectOdooCommand.CanExecute(null).Should().BeFalse();
    }

    [StaFact]
    public void NavigateToBOQ_WhenOdooDisconnected_CanExecuteReturnsFalse()
    {
        // Arrange
        var sut = CreateSUT();

        // Assert
        sut.NavigateToBOQCommand.CanExecute(null).Should().BeFalse();
        sut.NavigateToPRCommand.CanExecute(null).Should().BeFalse();
    }

    [StaFact]
    public void NavigateToBOQ_WhenOdooConnected_CanExecuteReturnsTrue()
    {
        // Arrange
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();
        sut.UpdateStatusFromServices();

        // Assert
        sut.NavigateToBOQCommand.CanExecute(null).Should().BeTrue();
        sut.NavigateToPRCommand.CanExecute(null).Should().BeTrue();
    }

    [StaFact]
    public void StatusPolling_UpdatesFromServiceState()
    {
        // Arrange - start disconnected
        var sut = CreateSUT();
        sut.IsAutoCADConnected.Should().BeFalse();

        // Act - simulate service becoming connected
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        sut.UpdateStatusFromServices();

        // Assert
        sut.IsAutoCADConnected.Should().BeTrue();
        sut.AutoCADStatusText.Should().Be("Connected");
        sut.IsOdooConnected.Should().BeTrue();
        sut.OdooStatusText.Should().Be("Connected");
    }

    [StaFact]
    public void NavigateToAutoCAD_CallsNavigationService()
    {
        // Arrange
        var sut = CreateSUT();

        // Act
        sut.NavigateToAutoCADCommand.Execute(null);

        // Assert
        _mockNav.Verify(n => n.NavigateTo("AutoCAD"), Times.Once);
    }
}
