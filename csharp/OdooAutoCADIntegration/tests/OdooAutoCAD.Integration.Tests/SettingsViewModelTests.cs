using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Configuration;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

/// <summary>
/// Tests for SettingsViewModel: load, save, toggle token, cancel, dirty tracking.
/// </summary>
public class SettingsViewModelTests
{
    private readonly Mock<ISettingsService> _mockSettings;
    private readonly SettingsViewModel _sut;

    public SettingsViewModelTests()
    {
        _mockSettings = new Mock<ISettingsService>();

        // Default AppSettings
        _mockSettings.Setup(s => s.GetAppSettings()).Returns(new AppSettings
        {
            Application = new ApplicationInfo { Name = "Test App", Version = "6.0.0" },
            Odoo = new OdooSettings
            {
                ServerUrl = "https://odoo.example.com",
                Database = "test_db",
                Username = "admin",
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

        _sut = new SettingsViewModel(_mockSettings.Object);
    }

    // --- LoadSettings ---

    [Fact]
    public async Task LoadSettings_PopulatesOdooFields_FromAppSettings()
    {
        // Act
        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        // Assert
        _sut.OdooServerUrl.Should().Be("https://odoo.example.com");
        _sut.OdooDatabase.Should().Be("test_db");
        _sut.OdooUsername.Should().Be("admin");
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
                ["odoo_api_token"] = "secret-token-123"
            });

        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.OdooApiToken.Should().Be("secret-token-123");
    }

    [Fact]
    public async Task LoadSettings_OverridesFromDatabase_WhenPresent()
    {
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_server_url"] = "https://db-override.com",
                ["odoo_database"] = "db_override",
                ["odoo_username"] = "db_user"
            });

        await _sut.LoadSettingsCommand.ExecuteAsync(null);

        _sut.OdooServerUrl.Should().Be("https://db-override.com");
        _sut.OdooDatabase.Should().Be("db_override");
        _sut.OdooUsername.Should().Be("db_user");
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

        _sut.OdooServerUrl = "https://changed.com";

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
        _sut.OdooServerUrl = "https://new-server.com";

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
                d["odoo_api_token"] == "my-secret-token")),
            Times.Once);
    }

    [Fact]
    public async Task SaveSettings_ClearsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooServerUrl = "https://changed.com";
        _sut.HasUnsavedChanges.Should().BeTrue();

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _sut.HasUnsavedChanges.Should().BeFalse();
    }

    [Fact]
    public async Task SaveSettings_SetsSuccessStatusMessage()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooServerUrl = "https://changed.com";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _sut.StatusMessage.Should().Be("Settings saved successfully");
    }

    [Fact]
    public async Task SaveSettings_OnError_SetsErrorStatusMessage()
    {
        _mockSettings.Setup(s => s.SaveAppSettingsAsync(It.IsAny<AppSettings>()))
            .ThrowsAsync(new InvalidOperationException("Write failed"));

        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooServerUrl = "changed";

        await _sut.SaveSettingsCommand.ExecuteAsync(null);

        _sut.StatusMessage.Should().Contain("Failed to save");
    }

    // --- CancelChanges ---

    [Fact]
    public async Task CancelChanges_RestoresOriginalValues()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        var originalUrl = _sut.OdooServerUrl;

        _sut.OdooServerUrl = "https://modified.com";
        _sut.CancelChangesCommand.Execute(null);

        _sut.OdooServerUrl.Should().Be(originalUrl);
    }

    [Fact]
    public async Task CancelChanges_ClearsHasUnsavedChanges()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooServerUrl = "https://modified.com";
        _sut.HasUnsavedChanges.Should().BeTrue();

        _sut.CancelChangesCommand.Execute(null);

        _sut.HasUnsavedChanges.Should().BeFalse();
    }

    [Fact]
    public async Task CancelChanges_RestoresMultipleFields()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        var originalDb = _sut.OdooDatabase;
        var originalPort = _sut.McpPort;

        _sut.OdooDatabase = "changed_db";
        _sut.McpPort = 1234;

        _sut.CancelChangesCommand.Execute(null);

        _sut.OdooDatabase.Should().Be(originalDb);
        _sut.McpPort.Should().Be(originalPort);
    }

    [Fact]
    public async Task CancelChanges_SetsStatusMessage()
    {
        await _sut.LoadSettingsCommand.ExecuteAsync(null);
        _sut.OdooServerUrl = "changed";

        _sut.CancelChangesCommand.Execute(null);

        _sut.StatusMessage.Should().Be("Changes cancelled");
    }
}
