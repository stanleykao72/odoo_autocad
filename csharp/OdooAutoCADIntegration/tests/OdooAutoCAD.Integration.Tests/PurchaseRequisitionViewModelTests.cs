using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Configuration;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class PurchaseRequisitionViewModelTests
{
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<IAppLogService> _mockLogService;
    private readonly Mock<ISettingsService> _mockSettingsService;

    public PurchaseRequisitionViewModelTests()
    {
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockOdoo = new Mock<IOdooService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockLogService = new Mock<IAppLogService>();
        _mockSettingsService = new Mock<ISettingsService>();
    }

    private PurchaseRequisitionViewModel CreateSUT()
    {
        return new PurchaseRequisitionViewModel(
            _mockAutoCAD.Object,
            _mockOdoo.Object,
            _mockGuiProxy.Object,
            _mockLogService.Object,
            _mockSettingsService.Object);
    }

    #region Constructor Tests

    [Fact]
    public void Constructor_DefaultsToDisconnected()
    {
        var sut = CreateSUT();

        sut.IsAutoCADConnected.Should().BeFalse();
        sut.IsOdooConnected.Should().BeFalse();
        sut.IsConverting.Should().BeFalse();
        sut.IsLoading.Should().BeFalse();
        sut.IsSubmitting.Should().BeFalse();
        sut.TotalPRs.Should().Be(0);
    }

    [Fact]
    public void Constructor_HasNoPRs()
    {
        var sut = CreateSUT();

        sut.PrItems.Should().BeEmpty();
        sut.HasData.Should().BeFalse();
        sut.HasNoData.Should().BeTrue();
    }

    #endregion

    #region Convert Command Guard Tests

    [Fact]
    public void CanConvert_WhenAutoCADDisconnected_ReturnsFalse()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();

        sut.ConvertBOQToPRCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void CanConvert_WhenOdooDisconnected_ReturnsFalse()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        sut.ConvertBOQToPRCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void CanConvert_WhenBothConnected_ReturnsTrue()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();

        sut.ConvertBOQToPRCommand.CanExecute(null).Should().BeTrue();
    }

    #endregion

    #region Convert Command Tests

    [Fact]
    public async Task Convert_NoHeaderIds_SetsErrorStatus()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string>()));

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.LastConvertResult.Should().Contain("No header IDs");
        sut.PrItems.Should().BeEmpty();
    }

    [Fact]
    public async Task Convert_MissingSwaggerConfig_SetsConfigError()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string> { "H001" }));
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>());
        _mockSettingsService
            .Setup(s => s.GetAppSettings())
            .Returns(new AppSettings());

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.LastConvertResult.Should().Contain("Missing configuration");
    }

    [Fact]
    public async Task Convert_ApiError_SetsErrorStatus()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string> { "H001" }));
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://example.com/api/v1/boq_import_api/swagger.json?token=tok123&db=testdb",
                ["odoo_user_token"] = "tok123"
            });
        _mockSettingsService
            .Setup(s => s.GetAppSettings())
            .Returns(new AppSettings());
        _mockOdoo
            .Setup(s => s.ConvertBOQToPRViaApiAsync(
                It.IsAny<List<string>>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<string>()))
            .ReturnsAsync(new Boq2PrResponse
            {
                Success = false,
                ErrorCode = "CONVERSION_ERROR",
                ErrorMessage = "No BOQ data found"
            });

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.LastConvertResult.Should().Contain("CONVERSION_ERROR");
        sut.LastConvertResult.Should().Contain("No BOQ data found");
    }

    [Fact]
    public async Task Convert_Success_PopulatesPRItems()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string> { "H001", "H002" }));
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://example.com/api/v1/boq_import_api/swagger.json?token=tok123&db=testdb",
                ["odoo_user_token"] = "tok123"
            });
        _mockSettingsService
            .Setup(s => s.GetAppSettings())
            .Returns(new AppSettings());
        _mockOdoo
            .Setup(s => s.ConvertBOQToPRViaApiAsync(
                It.IsAny<List<string>>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<string>()))
            .ReturnsAsync(new Boq2PrResponse
            {
                Success = true,
                All = new List<Boq2PrResult>
                {
                    new() { PrId = 10, Reference = "PR/2026/001", State = "draft", Lines = new List<Boq2PrLine>
                    {
                        new() { ProductId = 1, ProductName = "Widget", Quantity = 10, UnitOfMeasure = "pcs" }
                    }},
                    new() { PrId = 11, Reference = "PR/2026/002", State = "draft", Lines = new List<Boq2PrLine>
                    {
                        new() { ProductId = 2, ProductName = "Gadget", Quantity = 5, UnitOfMeasure = "m" },
                        new() { ProductId = 3, ProductName = "Bolt", Quantity = 100, UnitOfMeasure = "pcs" }
                    }}
                }
            });

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.PrItems.Should().HaveCount(2);
        sut.PrItems[0].Reference.Should().Be("PR/2026/001");
        sut.PrItems[0].LineCount.Should().Be(1);
        sut.PrItems[1].Reference.Should().Be("PR/2026/002");
        sut.PrItems[1].LineCount.Should().Be(2);
        sut.HasData.Should().BeTrue();
    }

    [Fact]
    public async Task Convert_Success_UpdatesSummary()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string> { "H001" }));
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://example.com/api/v1/boq_import_api/swagger.json?token=tok123&db=testdb",
                ["odoo_user_token"] = "tok123"
            });
        _mockSettingsService
            .Setup(s => s.GetAppSettings())
            .Returns(new AppSettings());
        _mockOdoo
            .Setup(s => s.ConvertBOQToPRViaApiAsync(
                It.IsAny<List<string>>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<string>()))
            .ReturnsAsync(new Boq2PrResponse
            {
                Success = true,
                All = new List<Boq2PrResult>
                {
                    new() { PrId = 10, Reference = "PR/001", State = "draft" }
                }
            });

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.TotalPRs.Should().Be(1);
        sut.DraftCount.Should().Be(1);
        sut.SubmittedCount.Should().Be(0);
        sut.ApprovedCount.Should().Be(0);
        sut.LastConvertResult.Should().Contain("Success");
    }

    #endregion

    #region Refresh PR List Tests

    [Fact]
    public async Task RefreshList_NoProjectId_SetsStatus()
    {
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>());

        var sut = CreateSUT();
        await sut.RefreshPRListCommand.ExecuteAsync(null);

        sut.LoadStatusText.Should().Contain("No project ID");
    }

    [Fact]
    public async Task RefreshList_Success_PopulatesItems()
    {
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?> { ["current_project_id"] = "42" });
        _mockOdoo
            .Setup(s => s.GetPurchaseRequisitionsAsync(42))
            .ReturnsAsync(new List<PREntry>
            {
                new() { Id = 1, Reference = "PR/001", State = "draft", ProjectId = 42 },
                new() { Id = 2, Reference = "PR/002", State = "submitted", ProjectId = 42 }
            });

        var sut = CreateSUT();
        await sut.RefreshPRListCommand.ExecuteAsync(null);

        sut.PrItems.Should().HaveCount(2);
        sut.TotalPRs.Should().Be(2);
    }

    #endregion

    #region Submit PR Tests

    [Fact]
    public void CanSubmit_WhenNoPRSelected_ReturnsFalse()
    {
        var sut = CreateSUT();
        sut.SelectedPR = null;

        sut.SubmitPRCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void CanSubmit_WhenDraftSelected_ReturnsTrue()
    {
        var sut = CreateSUT();
        sut.SelectedPR = new PRDisplayItem { Id = 1, State = "draft" };

        sut.SubmitPRCommand.CanExecute(null).Should().BeTrue();
    }

    [Fact]
    public void CanSubmit_WhenSubmittedSelected_ReturnsFalse()
    {
        var sut = CreateSUT();
        sut.SelectedPR = new PRDisplayItem { Id = 1, State = "submitted" };

        sut.SubmitPRCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public async Task Submit_Success_UpdatesState()
    {
        _mockOdoo.Setup(s => s.SubmitPRAsync(1)).ReturnsAsync(true);

        var sut = CreateSUT();
        var pr = new PRDisplayItem { Id = 1, Reference = "PR/001", State = "draft" };
        sut.PrItems.Add(pr);
        sut.SelectedPR = pr;
        sut.UpdateSummary();

        await sut.SubmitPRCommand.ExecuteAsync(null);

        pr.State.Should().Be("submitted");
        sut.SubmitStatusText.Should().Contain("submitted");
    }

    [Fact]
    public async Task Submit_Failure_ShowsError()
    {
        _mockOdoo.Setup(s => s.SubmitPRAsync(1)).ReturnsAsync(false);

        var sut = CreateSUT();
        var pr = new PRDisplayItem { Id = 1, Reference = "PR/001", State = "draft" };
        sut.PrItems.Add(pr);
        sut.SelectedPR = pr;

        await sut.SubmitPRCommand.ExecuteAsync(null);

        sut.SubmitStatusText.Should().Contain("Failed");
        pr.State.Should().Be("draft"); // unchanged
    }

    #endregion

    #region Summary Tests

    [Fact]
    public void UpdateSummary_CountsByState()
    {
        var sut = CreateSUT();

        sut.PrItems.Add(new PRDisplayItem { State = "draft" });
        sut.PrItems.Add(new PRDisplayItem { State = "draft" });
        sut.PrItems.Add(new PRDisplayItem { State = "submitted" });
        sut.PrItems.Add(new PRDisplayItem { State = "approved" });
        sut.PrItems.Add(new PRDisplayItem { State = "done" });

        sut.UpdateSummary();

        sut.TotalPRs.Should().Be(5);
        sut.DraftCount.Should().Be(2);
        sut.SubmittedCount.Should().Be(1);
        sut.ApprovedCount.Should().Be(2); // approved + done
    }

    #endregion

    #region Filter/Sort Tests

    [Fact]
    public void FilteredPRs_DefaultAll_ShowsAllItems()
    {
        var sut = CreateSUT();
        sut.PrItems.Add(new PRDisplayItem { State = "draft", Reference = "PR001", CreatedAt = DateTime.Now });
        sut.PrItems.Add(new PRDisplayItem { State = "submitted", Reference = "PR002", CreatedAt = DateTime.Now });
        sut.PrItems.Add(new PRDisplayItem { State = "approved", Reference = "PR003", CreatedAt = DateTime.Now });

        sut.UpdateSummary();

        sut.FilteredPRs.Should().HaveCount(3);
    }

    [Fact]
    public void FilteredPRs_FilterByDraft_ShowsOnlyDraft()
    {
        var sut = CreateSUT();
        sut.PrItems.Add(new PRDisplayItem { State = "draft", Reference = "PR001", CreatedAt = DateTime.Now });
        sut.PrItems.Add(new PRDisplayItem { State = "submitted", Reference = "PR002", CreatedAt = DateTime.Now });
        sut.PrItems.Add(new PRDisplayItem { State = "draft", Reference = "PR003", CreatedAt = DateTime.Now });

        sut.UpdateSummary();
        sut.SelectedStateFilter = "Draft";

        sut.FilteredPRs.Should().HaveCount(2);
        sut.FilteredPRs.Should().OnlyContain(p => p.State == "draft");
    }

    [Fact]
    public void FilteredPRs_SortByReferenceAZ_SortsCorrectly()
    {
        var sut = CreateSUT();
        sut.PrItems.Add(new PRDisplayItem { State = "draft", Reference = "PR003", CreatedAt = DateTime.Now });
        sut.PrItems.Add(new PRDisplayItem { State = "draft", Reference = "PR001", CreatedAt = DateTime.Now });
        sut.PrItems.Add(new PRDisplayItem { State = "draft", Reference = "PR002", CreatedAt = DateTime.Now });

        sut.UpdateSummary();
        sut.SelectedSortOrder = "Reference A-Z";

        sut.FilteredPRs[0].Reference.Should().Be("PR001");
        sut.FilteredPRs[1].Reference.Should().Be("PR002");
        sut.FilteredPRs[2].Reference.Should().Be("PR003");
    }

    #endregion

    #region Detail View Tests

    [Fact]
    public void SelectPR_PopulatesLines()
    {
        var sut = CreateSUT();
        var entry = new PREntry
        {
            Id = 1, Reference = "PR/001", State = "draft"
        };
        entry.Lines.Add(new PRLine
        {
            ProductName = "Widget", Quantity = 10, UnitOfMeasure = "pcs", UnitPrice = 5.0m
        });
        entry.Lines.Add(new PRLine
        {
            ProductName = "Bolt", Quantity = 100, UnitOfMeasure = "pcs", UnitPrice = 0.5m
        });

        var pr = new PRDisplayItem { Id = 1, Reference = "PR/001", State = "draft", OriginalEntry = entry };
        sut.PrItems.Add(pr);

        sut.SelectedPR = pr;

        sut.SelectedPRLines.Should().HaveCount(2);
        sut.SelectedPRLines[0].ProductName.Should().Be("Widget");
        sut.SelectedPRLines[1].ProductName.Should().Be("Bolt");
        sut.SelectedPRTotal.Should().Be(10 * 5.0m + 100 * 0.5m);
    }

    [Fact]
    public void SelectPR_NullSelection_ClearsLines()
    {
        var sut = CreateSUT();
        var entry = new PREntry { Id = 1, Reference = "PR/001", State = "draft" };
        entry.Lines.Add(new PRLine { ProductName = "Widget", Quantity = 10, UnitOfMeasure = "pcs", UnitPrice = 5.0m });

        var pr = new PRDisplayItem { Id = 1, OriginalEntry = entry };
        sut.PrItems.Add(pr);
        sut.SelectedPR = pr;
        sut.SelectedPRLines.Should().HaveCount(1);

        sut.SelectedPR = null;

        sut.SelectedPRLines.Should().BeEmpty();
        sut.SelectedPRTotal.Should().Be(0);
    }

    [Fact]
    public void PRLineDisplayItem_ComputesLineTotal()
    {
        var line = new PRLineDisplayItem
        {
            ProductName = "Widget", Quantity = 10, UnitPrice = 2.5m
        };

        line.LineTotal.Should().Be(25.0m);
    }

    [Fact]
    public void PRLineDisplayItem_NullUnitPrice_LineTotalIsZero()
    {
        var line = new PRLineDisplayItem
        {
            ProductName = "Widget", Quantity = 10, UnitPrice = null
        };

        line.LineTotal.Should().Be(0);
    }

    #endregion

    #region Feedback Tests

    [Fact]
    public async Task Convert_Success_SetsStatusMessageTypeSuccess()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string> { "H001" }));
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://example.com/api/v1/boq_import_api/swagger.json?token=tok123&db=testdb",
                ["odoo_user_token"] = "tok123"
            });
        _mockSettingsService
            .Setup(s => s.GetAppSettings())
            .Returns(new AppSettings());
        _mockOdoo
            .Setup(s => s.ConvertBOQToPRViaApiAsync(
                It.IsAny<List<string>>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<string>()))
            .ReturnsAsync(new Boq2PrResponse
            {
                Success = true,
                All = new List<Boq2PrResult>
                {
                    new() { PrId = 10, Reference = "PR/001", State = "draft", Lines = new List<Boq2PrLine>() }
                }
            });

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.StatusMessageType.Should().Be("success");
        sut.LastConversionTime.Should().NotBeNull();
    }

    [Fact]
    public async Task Convert_ApiError_SetsStatusMessageTypeError()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string> { "H001" }));
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://example.com/api/v1/boq_import_api/swagger.json?token=tok123&db=testdb",
                ["odoo_user_token"] = "tok123"
            });
        _mockSettingsService
            .Setup(s => s.GetAppSettings())
            .Returns(new AppSettings());
        _mockOdoo
            .Setup(s => s.ConvertBOQToPRViaApiAsync(
                It.IsAny<List<string>>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<string>()))
            .ReturnsAsync(new Boq2PrResponse
            {
                Success = false,
                ErrorCode = "ERR",
                ErrorMessage = "Something failed"
            });

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.StatusMessageType.Should().Be("error");
        sut.LastConversionTime.Should().BeNull();
    }

    [Fact]
    public async Task Convert_EmptyResult_SetsStatusMessageTypeInfo()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string> { "H001" }));
        _mockSettingsService
            .Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://example.com/api/v1/boq_import_api/swagger.json?token=tok123&db=testdb",
                ["odoo_user_token"] = "tok123"
            });
        _mockSettingsService
            .Setup(s => s.GetAppSettings())
            .Returns(new AppSettings());
        _mockOdoo
            .Setup(s => s.ConvertBOQToPRViaApiAsync(
                It.IsAny<List<string>>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<string>()))
            .ReturnsAsync(new Boq2PrResponse
            {
                Success = true,
                All = new List<Boq2PrResult>() // empty
            });

        var sut = CreateSUT();
        await sut.ConvertBOQToPRCommand.ExecuteAsync(null);

        sut.StatusMessageType.Should().Be("info");
        sut.PrItems.Should().BeEmpty();
    }

    [Fact]
    public void StatusMessageType_DefaultsToEmpty()
    {
        var sut = CreateSUT();
        sut.StatusMessageType.Should().BeEmpty();
        sut.LastConversionTime.Should().BeNull();
    }

    #endregion
}
