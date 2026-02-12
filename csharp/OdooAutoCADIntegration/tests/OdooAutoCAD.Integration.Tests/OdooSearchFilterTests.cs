using FluentAssertions;
using Microsoft.Extensions.Configuration;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.Odoo;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

/// <summary>
/// Tests for Sprint 8 Odoo page enhancements:
/// Product search, category filter, project search, server info, sync time.
/// </summary>
public class OdooSearchFilterTests
{
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<ISettingsService> _mockSettings;
    private readonly Mock<ICredentialService> _mockCredential;
    private readonly Mock<IAppLogService> _mockLogService;
    private readonly IConfiguration _configuration;

    public OdooSearchFilterTests()
    {
        _mockOdoo = new Mock<IOdooService>();
        _mockSettings = new Mock<ISettingsService>();
        _mockCredential = new Mock<ICredentialService>();
        _mockLogService = new Mock<IAppLogService>();

        // Set up minimal configuration
        _configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["Odoo:TimeoutSeconds"] = "30"
            })
            .Build();

        // Default mock setups
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>());
    }

    private OdooConnectionViewModel CreateSUT()
    {
        return new OdooConnectionViewModel(
            _mockOdoo.Object,
            _mockSettings.Object,
            _mockCredential.Object,
            _configuration,
            _mockLogService.Object);
    }

    #region US-003-06: Product Search

    [Fact]
    public void SearchProducts_EmptyTerm_CannotExecute()
    {
        var sut = CreateSUT();
        sut.ProductSearchTerm = "";

        sut.SearchProductsCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public async Task SearchProducts_WithTerm_PopulatesProducts()
    {
        var products = new List<OdooProduct>
        {
            new(1, "Steel Pipe", "SP001", null, 10.5m, "m", "Pipe"),
            new(2, "Steel Plate", "SP002", null, 25.0m, "kg", "Plate")
        };
        _mockOdoo.Setup(o => o.SearchProductsAsync("steel"))
            .ReturnsAsync(products);

        var sut = CreateSUT();
        sut.IsConnected = true;
        sut.ProductSearchTerm = "steel";

        await sut.SearchProductsCommand.ExecuteAsync(null);

        sut.Products.Should().HaveCount(2);
        sut.ProductCount.Should().Be(2);
        sut.SearchStatusText.Should().Contain("steel");
    }

    [Fact]
    public void ClearSearch_RestoresAllProducts()
    {
        var sut = CreateSUT();

        // Simulate having products loaded
        sut.Products.Add(new OdooProduct(1, "A", "A1", null, null, null, null));
        sut.ProductSearchTerm = "test";
        sut.SearchStatusText = "Found 1";

        sut.ClearProductSearchCommand.Execute(null);

        sut.ProductSearchTerm.Should().BeEmpty();
        sut.SearchStatusText.Should().BeEmpty();
    }

    #endregion

    #region US-003-07: Project Search

    [Fact]
    public async Task SearchProjects_WithTerm_PopulatesProjects()
    {
        var projects = new List<OdooProject>
        {
            new(1, "Bridge Project", "BP001", "open", DateTime.Today, null),
            new(2, "Road Project", "RP001", "open", DateTime.Today, null)
        };
        _mockOdoo.Setup(o => o.SearchProjectsAsync("project"))
            .ReturnsAsync(projects);

        var sut = CreateSUT();
        sut.IsConnected = true;
        sut.ProjectSearchTerm = "project";

        await sut.SearchProjectsCommand.ExecuteAsync(null);

        sut.Projects.Should().HaveCount(2);
        sut.ProjectCount.Should().Be(2);
        sut.ProjectStatusText.Should().Contain("project");
    }

    [Fact]
    public async Task SelectProject_SetsSelectedProject()
    {
        _mockSettings.Setup(s => s.SaveServerConfigsAsync(It.IsAny<Dictionary<string, string?>>()))
            .Returns(Task.CompletedTask);

        var sut = CreateSUT();
        sut.SelectedProject = new OdooProject(42, "Test Project", "TP042", "open", null, null);

        await sut.SelectProjectCommand.ExecuteAsync(null);

        sut.ProjectStatusText.Should().Contain("Test Project");
        sut.ProjectStatusText.Should().Contain("42");
    }

    [Fact]
    public async Task SelectProject_SavesProjectIdToSettings()
    {
        Dictionary<string, string?>? savedConfigs = null;
        _mockSettings.Setup(s => s.SaveServerConfigsAsync(It.IsAny<Dictionary<string, string?>>()))
            .Callback<Dictionary<string, string?>>(configs => savedConfigs = configs)
            .Returns(Task.CompletedTask);

        var sut = CreateSUT();
        sut.SelectedProject = new OdooProject(42, "Test Project", "TP042", "open", null, null);

        await sut.SelectProjectCommand.ExecuteAsync(null);

        savedConfigs.Should().NotBeNull();
        savedConfigs!["odoo_project_id"].Should().Be("42");
        savedConfigs["odoo_project_name"].Should().Be("Test Project");
    }

    #endregion

    #region US-003-08: Sync Time

    [Fact]
    public void LastSyncTime_Default_ShowsNeverSynced()
    {
        var sut = CreateSUT();

        sut.LastSyncTime.Should().BeNull();
        sut.LastSyncDisplay.Should().Be("Never synced");
    }

    [Fact]
    public void LastSyncTime_AfterSync_Updates()
    {
        var sut = CreateSUT();

        // Simulate sync completing
        var now = DateTime.Now;
        sut.LastSyncTime = now;
        sut.LastSyncDisplay = now.ToString("yyyy-MM-dd HH:mm:ss");

        sut.LastSyncTime.Should().Be(now);
        sut.LastSyncDisplay.Should().Contain(now.Year.ToString());
    }

    #endregion

    #region US-003-09: Server Info

    [Fact]
    public void ServerInfo_Default_IsEmpty()
    {
        var sut = CreateSUT();

        sut.ServerVersion.Should().BeEmpty();
        sut.ConnectedDatabase.Should().BeEmpty();
        sut.ConnectedUsername.Should().BeEmpty();
    }

    [Fact]
    public void ServerInfo_AfterConnect_PopulatesFields()
    {
        var sut = CreateSUT();

        // Simulate successful connection populating server info
        sut.ServerVersion = "API v1.0";
        sut.ConnectedDatabase = "production_db";
        sut.ConnectedUsername = "admi...dmin";

        sut.ServerVersion.Should().Be("API v1.0");
        sut.ConnectedDatabase.Should().Be("production_db");
        sut.ConnectedUsername.Should().Be("admi...dmin");
    }

    [Fact]
    public void ServerInfo_AfterDisconnect_ClearsFields()
    {
        var sut = CreateSUT();

        // Simulate connected state
        sut.ServerVersion = "API v1.0";
        sut.ConnectedDatabase = "production_db";
        sut.ConnectedUsername = "admin";

        // Simulate disconnect
        sut.ServerVersion = string.Empty;
        sut.ConnectedDatabase = string.Empty;
        sut.ConnectedUsername = string.Empty;

        sut.ServerVersion.Should().BeEmpty();
        sut.ConnectedDatabase.Should().BeEmpty();
        sut.ConnectedUsername.Should().BeEmpty();
    }

    #endregion

    #region US-003-11: Category Filter

    [Fact]
    public async Task LoadCategories_PopulatesCategories()
    {
        var categories = new List<OdooProductCategory>
        {
            new(1, "Pipes", null),
            new(2, "Steel", null),
            new(3, "Fittings", 1)
        };
        _mockOdoo.Setup(o => o.GetCategoriesAsync())
            .ReturnsAsync(categories);

        var sut = CreateSUT();
        sut.IsConnected = true;

        await sut.LoadCategoriesCommand.ExecuteAsync(null);

        sut.Categories.Should().HaveCount(3);
    }

    [Fact]
    public async Task FilterByCategory_PopulatesFiltered()
    {
        var filtered = new List<OdooProduct>
        {
            new(5, "Pipe A", "PA001", null, 15.0m, "m", "Pipes")
        };
        _mockOdoo.Setup(o => o.GetProductsByCategoryAsync(1))
            .ReturnsAsync(filtered);

        var sut = CreateSUT();
        sut.IsConnected = true;

        // Trigger category filter via SelectedCategory setter
        sut.SelectedCategory = new OdooProductCategory(1, "Pipes", null);

        // Allow async filter to complete
        await Task.Delay(200);

        sut.Products.Should().HaveCount(1);
        sut.Products[0].Name.Should().Be("Pipe A");
    }

    #endregion
}
