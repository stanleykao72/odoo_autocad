using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Configuration;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Threading;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

/// <summary>
/// Tests for SettingsViewModel: load, save, toggle token, cancel, dirty tracking.
/// </summary>
public class SettingsViewModelTests
{
    private readonly Mock<ISettingsService> _mockSettings;
    private readonly Mock<IAppLogService> _mockLogService;
    private readonly DrawingDataServiceDispatcher _dispatcher;
    private readonly SettingsViewModel _sut;

    public SettingsViewModelTests()
    {
        _mockSettings = new Mock<ISettingsService>();
        _mockLogService = new Mock<IAppLogService>();

        var mockAutoCAD = new Mock<IAutoCADService>();
        var mockProxy = new Mock<IGUIProxy>();
        var mockDwgFile = new Mock<IDwgFileService>();
        var comService = new ComDrawingDataService(mockAutoCAD.Object, mockProxy.Object);
        var fileService = new FileDrawingDataService(mockDwgFile.Object, new SidecarIdStore());
        _dispatcher = new DrawingDataServiceDispatcher(comService, fileService);

        // Default AppSettings
        _mockSettings.Setup(s => s.GetAppSettings()).Returns(new AppSettings
        {
            Application = new ApplicationInfo { Name = "Test App", Version = "6.0.0" },
            Odoo = new OdooSettings
            {
                SwaggerUrl = "https://odoo.example.com/api/v1/boq_import_api/swagger.json?token=abc&db=test_db",
                TimeoutSeconds = 30
            },
            AutoCAD = new AutoCADSettings
            {
                ProgId = "AutoCAD.Application",
                ConnectionTimeoutSeconds = 10,
                RetryAttempts = 3
            },
            MCP = new MCPSettings
            {
                Port = 8084,
                AutoStart = false,
                HeartbeatIntervalSeconds = 30
            },
            Database = new DatabaseSettings { ConnectionString = "Data Source=test.db" }
        });

        // Default empty server configs
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>());

        _sut = new SettingsViewModel(_mockSettings.Object, _mockLogService.Object, _dispatcher);
    }

    // --- LoadSettings ---

    [Fact]
    public async Task LoadSettings_PopulatesOdooFields_FromAppSettings()
    {
        // Act
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        // Assert
        _sut.OdooSwaggerUrl.Should().Be("https://odoo.example.com/api/v1/boq_import_api/swagger.json?token=abc&db=test_db");
        _sut.OdooTimeoutSeconds.Should().Be(30);
    }

    [Fact]
    public async Task LoadSettings_PopulatesAutoCADFields_FromAppSettings()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.AutoCADProgId.Should().Be("AutoCAD.Application");
        _sut.AutoCADConnectionTimeout.Should().Be(10);
        _sut.AutoCADRetryAttempts.Should().Be(3);
    }

    [Fact]
    public async Task LoadSettings_PopulatesMCPFields_FromAppSettings()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.McpPort.Should().Be(8084);
        _sut.McpAutoStart.Should().BeFalse();
        _sut.McpHeartbeatInterval.Should().Be(30);
    }

    [Fact]
    public async Task LoadSettings_PopulatesAppInfo()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.AppVersion.Should().Be("6.0.0");
        _sut.DatabasePath.Should().Be("Data Source=test.db");
        _sut.DotNetRuntime.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public async Task LoadSettings_LoadsTokenFromDatabase()
    {
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_user_token"] = "secret-token-123"
            });

        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.OdooApiToken.Should().Be("secret-token-123");
    }

    [Fact]
    public async Task LoadSettings_OverridesSwaggerUrlFromDatabase_WhenPresent()
    {
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://db-override.com/api/v1/boq_import_api/swagger.json?token=xyz&db=override_db"
            });

        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.OdooSwaggerUrl.Should().Be("https://db-override.com/api/v1/boq_import_api/swagger.json?token=xyz&db=override_db");
    }

    [Fact]
    public async Task LoadSettings_ClearsUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.HasUnsavedChanges.Should().BeFalse();
    }

    [Fact]
    public async Task LoadSettings_SetsStatusMessage()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.StatusMessage.Should().Be("Settings loaded");
    }

    // --- HasUnsavedChanges tracking ---

    [Fact]
    public async Task PropertyChange_AfterLoad_SetsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.HasUnsavedChanges.Should().BeFalse();

        _sut.OdooSwaggerUrl = "https://changed.com/swagger.json?token=a&db=b";

        _sut.HasUnsavedChanges.Should().BeTrue();
    }

    [Fact]
    public async Task PropertyChange_McpPort_SetsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.McpPort = 9999;

        _sut.HasUnsavedChanges.Should().BeTrue();
    }

    [Fact]
    public async Task PropertyChange_McpAutoStart_SetsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.McpAutoStart = true;

        _sut.HasUnsavedChanges.Should().BeTrue();
    }

    [Fact]
    public async Task PropertyChange_StatusMessage_DoesNotSetUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.StatusMessage = "Some status";

        _sut.HasUnsavedChanges.Should().BeFalse();
    }

    // --- ToggleTokenVisibility ---

    [Fact]
    public void ToggleTokenVisibility_TogglesIsTokenVisible()
    {
        _sut.IsTokenVisible.Should().BeFalse();

        _sut.ToggleTokenVisibilityCommand.Execute(null);
        _sut.IsTokenVisible.Should().BeTrue();

        _sut.ToggleTokenVisibilityCommand.Execute(null);
        _sut.IsTokenVisible.Should().BeFalse();
    }

    // --- SaveSettings ---

    [Fact]
    public async Task SaveSettings_CallsSaveAppSettings()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooSwaggerUrl = "https://new-server.com/swagger.json?token=x&db=y";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _mockSettings.Verify(s => s.SaveAppSettingsAsync(It.IsAny<AppSettings>()), Times.Once);
    }

    [Fact]
    public async Task SaveSettings_PersistsServerConfigs_IncludingToken()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooApiToken = "my-secret-token";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _mockSettings.Verify(s => s.SaveServerConfigsAsync(
            It.Is<Dictionary<string, string?>>(d =>
                d["odoo_user_token"] == "my-secret-token")),
            Times.Once);
    }

    [Fact]
    public async Task SaveSettings_ClearsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooSwaggerUrl = "https://changed.com/swagger.json?token=x&db=y";
        _sut.HasUnsavedChanges.Should().BeTrue();

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _sut.HasUnsavedChanges.Should().BeFalse();
    }

    [Fact]
    public async Task SaveSettings_SetsSuccessStatusMessage()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooSwaggerUrl = "https://changed.com/swagger.json?token=x&db=y";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _sut.StatusMessage.Should().Be("Settings saved successfully");
    }

    [Fact]
    public async Task SaveSettings_OnError_SetsErrorStatusMessage()
    {
        _mockSettings.Setup(s => s.SaveAppSettingsAsync(It.IsAny<AppSettings>()))
            .ThrowsAsync(new InvalidOperationException("Write failed"));

        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooSwaggerUrl = "changed";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _sut.StatusMessage.Should().Contain("Failed to save");
    }

    // --- CancelChanges ---

    [Fact]
    public async Task CancelChanges_RestoresOriginalValues()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        var originalUrl = _sut.OdooSwaggerUrl;

        _sut.OdooSwaggerUrl = "https://modified.com/swagger.json?token=a&db=b";
        _sut.CancelChangesCommand.Execute(null);

        _sut.OdooSwaggerUrl.Should().Be(originalUrl);
    }

    [Fact]
    public async Task CancelChanges_ClearsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooSwaggerUrl = "https://modified.com/swagger.json?token=a&db=b";
        _sut.HasUnsavedChanges.Should().BeTrue();

        _sut.CancelChangesCommand.Execute(null);

        _sut.HasUnsavedChanges.Should().BeFalse();
    }

    [Fact]
    public async Task CancelChanges_RestoresMultipleFields()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        var originalUrl = _sut.OdooSwaggerUrl;
        var originalPort = _sut.McpPort;

        _sut.OdooSwaggerUrl = "https://changed.com/swagger.json?token=x&db=y";
        _sut.McpPort = 1234;

        _sut.CancelChangesCommand.Execute(null);

        _sut.OdooSwaggerUrl.Should().Be(originalUrl);
        _sut.McpPort.Should().Be(originalPort);
    }

    [Fact]
    public async Task CancelChanges_SetsStatusMessage()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooSwaggerUrl = "changed";

        _sut.CancelChangesCommand.Execute(null);

        _sut.StatusMessage.Should().Be("Changes cancelled");
    }

    // --- AutoCAD Mode ---

    [Fact]
    public void AutoCADMode_DefaultsToCOM()
    {
        _sut.AutoCADMode.Should().Be("COM");
        _sut.IsComMode.Should().BeTrue();
        _sut.IsFileMode.Should().BeFalse();
    }

    [Fact]
    public async Task LoadSettings_LoadsModePrefFromDatabase()
    {
        _mockSettings.Setup(s => s.GetPreferenceAsync("autocad_mode", "COM"))
            .ReturnsAsync("File");

        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.AutoCADMode.Should().Be("File");
        _sut.IsFileMode.Should().BeTrue();
        _sut.IsComMode.Should().BeFalse();
    }

    [Fact]
    public async Task LoadSettings_DefaultsCOM_WhenNoPrefSaved()
    {
        _mockSettings.Setup(s => s.GetPreferenceAsync("autocad_mode", "COM"))
            .ReturnsAsync("COM");

        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.AutoCADMode.Should().Be("COM");
    }

    [Fact]
    public async Task SaveSettings_PersistsAutoCADMode()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.AutoCADMode = "File";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _mockSettings.Verify(s => s.SetPreferenceAsync("autocad_mode", "File", "string"), Times.Once);
    }

    [Fact]
    public void IsComMode_Set_ChangesAutoCADMode()
    {
        _sut.AutoCADMode = "File";
        _sut.IsFileMode.Should().BeTrue();

        _sut.IsComMode = true;
        _sut.AutoCADMode.Should().Be("COM");
    }

    [Fact]
    public void IsFileMode_Set_ChangesAutoCADMode()
    {
        _sut.AutoCADMode.Should().Be("COM");

        _sut.IsFileMode = true;
        _sut.AutoCADMode.Should().Be("File");
    }

    [Fact]
    public async Task CancelChanges_RestoresAutoCADMode()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.AutoCADMode.Should().Be("COM");

        _sut.AutoCADMode = "File";
        _sut.CancelChangesCommand.Execute(null);

        _sut.AutoCADMode.Should().Be("COM");
    }

    [Fact]
    public async Task ModeChange_SetsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.HasUnsavedChanges.Should().BeFalse();

        _sut.AutoCADMode = "File";

        _sut.HasUnsavedChanges.Should().BeTrue();
    }

    [Fact]
    public async Task SaveSettings_SwitchesDispatcherMode_ToFile()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.AutoCADMode = "File";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _dispatcher.Mode.Should().Be(AutoCADOperationMode.File);
    }

    [Fact]
    public async Task SaveSettings_SwitchesDispatcherMode_ToCOM()
    {
        // Start in File mode
        await _dispatcher.SwitchModeAsync(AutoCADOperationMode.File);
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.AutoCADMode = "COM";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _dispatcher.Mode.Should().Be(AutoCADOperationMode.COM);
    }
}
