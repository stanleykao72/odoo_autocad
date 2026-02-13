using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;
using OdooAutoCAD.Configuration;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class BOQViewModelTests
{
    private readonly Mock<IBOQProcessor> _mockBoqProcessor;
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<IAppLogService> _mockLogService;
    private readonly Mock<ISettingsService> _mockSettingsService;
    private readonly Mock<IDrawingDataService> _mockDrawingDataService;

    public BOQViewModelTests()
    {
        _mockBoqProcessor = new Mock<IBOQProcessor>();
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockOdoo = new Mock<IOdooService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockLogService = new Mock<IAppLogService>();
        _mockSettingsService = new Mock<ISettingsService>();
        _mockDrawingDataService = new Mock<IDrawingDataService>();
    }

    private BOQViewModel CreateSUT()
    {
        return new BOQViewModel(
            _mockBoqProcessor.Object,
            _mockAutoCAD.Object,
            _mockOdoo.Object,
            _mockGuiProxy.Object,
            _mockLogService.Object,
            _mockSettingsService.Object,
            _mockDrawingDataService.Object);
    }

    #region Existing Tests

    [Fact]
    public void Constructor_DefaultsToDisconnected()
    {
        var sut = CreateSUT();

        sut.IsAutoCADConnected.Should().BeFalse();
        sut.IsOdooConnected.Should().BeFalse();
        sut.IsExtracting.Should().BeFalse();
        sut.TotalItems.Should().Be(0);
        sut.HasData.Should().BeFalse();
        sut.BoqItems.Should().BeEmpty();
    }

    [Fact]
    public void CanExtract_WhenAutoCADNotConnected_ReturnsFalse()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        sut.ExtractBOQCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void CanExtract_WhenAutoCADConnected_ReturnsTrue()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        var sut = CreateSUT();

        sut.ExtractBOQCommand.CanExecute(null).Should().BeTrue();
    }

    [Fact]
    public async Task ExtractBOQ_NoLayouts_SetsStatusMessage()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService
            .Setup(s => s.GetLayoutsAsync())
            .ReturnsAsync(new List<LayoutInfo>());

        var sut = CreateSUT();
        await sut.ExtractBOQCommand.ExecuteAsync(null);

        sut.ExtractionStatus.Should().Contain("No layouts found");
        sut.BoqItems.Should().BeEmpty();
        sut.TotalItems.Should().Be(0);
    }

    [Fact]
    public async Task ExtractBOQ_WithLayouts_PopulatesItems()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);

        var layouts = new List<LayoutInfo>
        {
            new("Layout1", 1, false, ""),
            new("Layout2", 2, false, "")
        };
        _mockDrawingDataService
            .Setup(s => s.GetLayoutsAsync())
            .ReturnsAsync(layouts);

        var layoutData1 = new LayoutData { LayoutName = "Layout1" };
        var layoutData2 = new LayoutData { LayoutName = "Layout2" };

        _mockDrawingDataService
            .Setup(s => s.ExtractParametersAsync("Layout1"))
            .ReturnsAsync(layoutData1);

        _mockDrawingDataService
            .Setup(s => s.ExtractParametersAsync("Layout2"))
            .ReturnsAsync(layoutData2);

        var result1 = new BOQGenerationResult
        {
            Success = true,
            Entries = new List<BOQEntry>
            {
                new() { ProductName = "Product A", Quantity = 10, UnitOfMeasure = "pcs" },
                new() { ProductName = "Product B", Quantity = 5, UnitOfMeasure = "m" }
            }
        };
        var result2 = new BOQGenerationResult
        {
            Success = true,
            Entries = new List<BOQEntry>
            {
                new() { ProductName = "Product C", Quantity = 3, UnitOfMeasure = "pcs" }
            }
        };

        _mockBoqProcessor
            .Setup(p => p.GenerateBOQAsync(layoutData1, It.IsAny<BOQGenerationOptions>()))
            .ReturnsAsync(result1);
        _mockBoqProcessor
            .Setup(p => p.GenerateBOQAsync(layoutData2, It.IsAny<BOQGenerationOptions>()))
            .ReturnsAsync(result2);

        var sut = CreateSUT();
        await sut.ExtractBOQCommand.ExecuteAsync(null);

        sut.BoqItems.Should().HaveCount(3);
        sut.TotalItems.Should().Be(3);
        sut.HasData.Should().BeTrue();
        sut.LayoutCount.Should().Be(2);
        sut.IsExtracting.Should().BeFalse();
    }

    [Fact]
    public void UpdateSummary_CorrectCounts()
    {
        var sut = CreateSUT();

        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1", ValidationStatus = "Valid" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1", ValidationStatus = "Valid" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L2", ValidationStatus = "Invalid" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L2", ValidationStatus = "Skipped" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L3", ValidationStatus = "Valid" });

        sut.UpdateSummary();

        sut.TotalItems.Should().Be(5);
        sut.ValidItems.Should().Be(3);
        sut.InvalidItems.Should().Be(1);
        sut.SkippedItems.Should().Be(1);
        sut.LayoutCount.Should().Be(3);
    }

    #endregion

    #region Validate Command Tests

    [Fact]
    public async Task ValidateCommand_WithValidEntries_SetsAllValid()
    {
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Widget", ProductCode = "W001", Quantity = 5
        });
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Gadget", ProductCode = "G001", Quantity = 10
        });
        sut.UpdateSummary();

        await sut.ValidateCommand.ExecuteAsync(null);

        sut.BoqItems.Should().OnlyContain(i => i.ValidationStatus == "Valid");
        sut.HasValidationErrors.Should().BeFalse();
    }

    [Fact]
    public async Task ValidateCommand_WithZeroQty_SetsWarning()
    {
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Widget", ProductCode = "W001", Quantity = 0
        });
        sut.UpdateSummary();

        await sut.ValidateCommand.ExecuteAsync(null);

        sut.BoqItems[0].ValidationStatus.Should().Be("Warning");
        sut.BoqItems[0].ValidationMessage.Should().Contain("Zero quantity");
        sut.HasValidationErrors.Should().BeFalse();
    }

    [Fact]
    public async Task ValidateCommand_WithMissingProductNo_SetsError()
    {
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "", ProductCode = "", Quantity = 5
        });
        sut.UpdateSummary();

        await sut.ValidateCommand.ExecuteAsync(null);

        sut.BoqItems[0].ValidationStatus.Should().Be("Error");
        sut.BoqItems[0].ValidationMessage.Should().Contain("Missing product number");
        sut.HasValidationErrors.Should().BeTrue();
    }

    [Fact]
    public async Task ValidateCommand_UpdatesSummaryCounts()
    {
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Good", Quantity = 10
        });
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Good", Quantity = 0
        });
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "", ProductCode = "", Quantity = 5
        });
        sut.UpdateSummary();

        await sut.ValidateCommand.ExecuteAsync(null);

        sut.ValidItems.Should().Be(1);
        sut.WarningItems.Should().Be(1);
        sut.InvalidItems.Should().Be(1);
    }

    [Fact]
    public void ValidateCommand_WhenNoData_CannotExecute()
    {
        var sut = CreateSUT();
        sut.ValidateCommand.CanExecute(null).Should().BeFalse();
    }

    #endregion

    #region Push Command Tests

    [Fact]
    public void PushCommand_WithDataAndOdooConnected_CanExecute()
    {
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1" });
        sut.UpdateSummary();

        // Push is enabled after extraction — validation is optional
        sut.PushToOdooCommand.CanExecute(null).Should().BeTrue();
    }

    [Fact]
    public async Task PushCommand_WhenValidationHasErrors_CannotExecute()
    {
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);

        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "", ProductCode = "", Quantity = 5
        });
        sut.UpdateSummary();

        // Run validation — missing product number will set errors
        await sut.ValidateCommand.ExecuteAsync(null);

        sut.HasValidationErrors.Should().BeTrue();
        sut.PushToOdooCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void PushCommand_WhenOdooDisconnected_CannotExecute()
    {
        _mockOdoo.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1" });
        sut.UpdateSummary();

        sut.PushToOdooCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void PushCommand_BuildsCorrectImportRequest()
    {
        var sut = CreateSUT();

        // Simulate extraction by populating _layoutDataMap via reflection
        var layoutData = new LayoutData
        {
            LayoutName = "TestLayout",
            Parameters = new Dictionary<string, object>
            {
                ["pr_no"] = "PR001",
                ["project_name"] = "TestProject"
            },
            Tables = new List<TableData>
            {
                new()
                {
                    Name = "MainTable",
                    ColumnCount = 9,
                    RowCount = 3,
                    Cells = new List<List<string>>
                    {
                        new() { "Position", "Product No", "Width", "Height", "Length", "Thickness", "Qty", "Description", "HEADER_ID" },  // Row 0: column headers (added by C# extraction)
                        new() { "1", "A001", "100", "200", "300", "10", "5", "Desc A", "" },  // Row 1: data
                        new() { "2", "B002", "150", "250", "350", "15", "3", "Desc B", "" }   // Row 2: data
                    }
                }
            }
        };

        // Access private field via reflection
        var field = typeof(BOQViewModel).GetField("_layoutDataMap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        var map = (Dictionary<string, LayoutData>)field!.GetValue(sut)!;
        map["TestLayout"] = layoutData;

        var request = sut.BuildImportRequest();

        request.All.Should().HaveCount(1);
        request.All[0].LayoutName.Should().Be("TestLayout");
        request.All[0].PrNo.Should().Be("PR001");
        request.All[0].ProjectName.Should().Be("TestProject");
        request.All[0].Detail.Should().HaveCount(2);
        request.All[0].Detail[0].ProductNo.Should().Be("A001");
        request.All[0].Detail[0].Qty.Should().Be("5");
        request.All[0].Detail[1].ProductNo.Should().Be("B002");
    }

    [Fact]
    public async Task PushCommand_OnApiError_SetsPushStatus()
    {
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);

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
            .Setup(s => s.ImportToBOQViaApiAsync(
                It.IsAny<BoqImportRequest>(), It.IsAny<string>(), It.IsAny<string>(),
                It.IsAny<string>(), It.IsAny<string>()))
            .ReturnsAsync(new BoqImportResponse
            {
                Success = false,
                ErrorCode = "IMPORT_ERROR",
                ErrorMessage = "Some fields are invalid"
            });

        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Widget", ProductCode = "W001", Quantity = 5
        });
        sut.UpdateSummary();

        // Validate first
        await sut.ValidateCommand.ExecuteAsync(null);

        // Populate layout data
        var field = typeof(BOQViewModel).GetField("_layoutDataMap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        var map = (Dictionary<string, LayoutData>)field!.GetValue(sut)!;
        map["L1"] = new LayoutData { LayoutName = "L1" };

        // Push
        await sut.PushToOdooCommand.ExecuteAsync(null);

        sut.LastPushResult.Should().Contain("IMPORT_ERROR");
        sut.LastPushResult.Should().Contain("Some fields are invalid");
    }

    #endregion

    #region Summary and Utility Tests

    [Fact]
    public void UpdateSummary_CountsWarningsSeparately()
    {
        var sut = CreateSUT();

        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1", ValidationStatus = "Valid" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1", ValidationStatus = "Warning" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1", ValidationStatus = "Warning" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L2", ValidationStatus = "Error" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L2", ValidationStatus = "Skipped" });

        sut.UpdateSummary();

        sut.TotalItems.Should().Be(5);
        sut.ValidItems.Should().Be(1);
        sut.WarningItems.Should().Be(2);
        sut.InvalidItems.Should().Be(1);
        sut.SkippedItems.Should().Be(1);
    }

    [Fact]
    public void SkippedBreakdown_FormatsCorrectly()
    {
        var sut = CreateSUT();
        sut.SkippedEmptyRows = 3;
        sut.SkippedIllegalTables = 1;

        sut.SkippedBreakdown.Should().Contain("Empty rows: 3");
        sut.SkippedBreakdown.Should().Contain("Illegal tables: 1");
    }

    [Fact]
    public async Task ExtractBOQ_StoresLayoutData()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);

        var layouts = new List<LayoutInfo> { new("TestLayout", 1, false, "") };
        _mockDrawingDataService
            .Setup(s => s.GetLayoutsAsync())
            .ReturnsAsync(layouts);

        var layoutData = new LayoutData { LayoutName = "TestLayout" };
        _mockDrawingDataService
            .Setup(s => s.ExtractParametersAsync("TestLayout"))
            .ReturnsAsync(layoutData);

        _mockBoqProcessor
            .Setup(p => p.GenerateBOQAsync(layoutData, It.IsAny<BOQGenerationOptions>()))
            .ReturnsAsync(new BOQGenerationResult
            {
                Success = true,
                Entries = new List<BOQEntry>
                {
                    new() { ProductName = "P1", Quantity = 1, UnitOfMeasure = "pcs" }
                }
            });

        var sut = CreateSUT();
        await sut.ExtractBOQCommand.ExecuteAsync(null);

        // Verify layout data stored via BuildImportRequest
        var request = sut.BuildImportRequest();
        request.All.Should().HaveCount(1);
        request.All[0].LayoutName.Should().Be("TestLayout");
    }

    [Fact]
    public void FindDetailId_CaseInsensitiveMatch()
    {
        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ABC001", DetailId = "42" },
            new() { ProductNo = "DEF002", DetailId = "99" }
        };

        BOQViewModel.FindDetailId(details, "abc001").Should().Be("42");
        BOQViewModel.FindDetailId(details, "ABC001").Should().Be("42");
        BOQViewModel.FindDetailId(details, "def002").Should().Be("99");
    }

    [Fact]
    public void FindDetailId_ReturnsNull_WhenNoMatch()
    {
        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ABC001", DetailId = "42" }
        };

        BOQViewModel.FindDetailId(details, "NOTFOUND").Should().BeNull();
        BOQViewModel.FindDetailId(details, "").Should().BeNull();
    }

    #endregion

    #region Validation Panel Tests

    [Fact]
    public async Task ValidateCommand_PopulatesValidationErrors()
    {
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Good", Quantity = 10
        });
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "", ProductCode = "", Quantity = 5
        });
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L2", ProductName = "Good", Quantity = 0
        });
        sut.UpdateSummary();

        await sut.ValidateCommand.ExecuteAsync(null);

        sut.ValidationErrors.Should().HaveCount(2); // 1 error (missing product) + 1 warning (zero qty)
        sut.ValidationErrorCount.Should().Be(1);
        sut.ValidationWarningCount.Should().Be(1);
        sut.IsValidationPanelVisible.Should().BeTrue();
    }

    [Fact]
    public async Task ValidateCommand_NoErrors_HidesPanel()
    {
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Widget", Quantity = 5
        });
        sut.UpdateSummary();

        await sut.ValidateCommand.ExecuteAsync(null);

        sut.ValidationErrors.Should().BeEmpty();
        sut.IsValidationPanelVisible.Should().BeFalse();
    }

    [Fact]
    public void IgnoreValidationError_RemovesWarningAndUpdatesItem()
    {
        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem
        {
            LayoutName = "L1", ProductName = "Test", ValidationStatus = "Warning",
            ValidationMessage = "Zero quantity"
        });

        var error = new BOQValidationErrorItem
        {
            Severity = "Warning", Message = "Zero quantity",
            LayoutName = "L1", RowIndex = 0, ProductName = "Test"
        };
        sut.ValidationErrors.Add(error);

        sut.IgnoreValidationErrorCommand.Execute(error);

        sut.ValidationErrors.Should().BeEmpty();
        sut.BoqItems[0].ValidationStatus.Should().Be("Valid");
    }

    #endregion

    #region Clear All IDs Tests

    [Fact]
    public async Task ClearAllIds_ClearsDetailIds()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.COM);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_clear_all_table_ids", null, 30000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", null));

        var sut = CreateSUT();
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1", DetailId = "42" });
        sut.BoqItems.Add(new BOQDisplayItem { LayoutName = "L1", DetailId = "43" });

        await sut.ClearAllIdsCommand.ExecuteAsync(null);

        sut.BoqItems.Should().OnlyContain(i => i.DetailId == null);
    }

    [Fact]
    public void ClearAllIds_WhenDisconnected_CannotExecute()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();

        sut.ClearAllIdsCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public void ClearAllIds_WhenConnected_CanExecute()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.COM);
        var sut = CreateSUT();

        sut.ClearAllIdsCommand.CanExecute(null).Should().BeTrue();
    }

    #endregion

    #region Progress Tests

    [Fact]
    public void CancelOperation_SetsCancelling()
    {
        var sut = CreateSUT();
        sut.IsCancelling.Should().BeFalse();

        sut.CancelOperationCommand.Execute(null);

        sut.IsCancelling.Should().BeTrue();
    }

    [Fact]
    public void PushProgress_DefaultsToEmpty()
    {
        var sut = CreateSUT();
        sut.PushProgressPercent.Should().Be(0);
        sut.PushProgress.Should().BeEmpty();
    }

    #endregion

    #region Dual-Mode (COM vs File) Tests

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

    [Fact]
    public void ClearAllIds_InFileMode_CannotExecute()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.File);
        var sut = CreateSUT();

        sut.ClearAllIdsCommand.CanExecute(null).Should().BeFalse();
    }

    [Fact]
    public async Task ExtractBOQ_InFileMode_UsesDrawingDataService()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.File);
        _mockOdoo.Setup(s => s.IsConnected).Returns(true);

        var layoutData = new LayoutData { LayoutName = "FileLayout1" };
        _mockDrawingDataService.Setup(s => s.GetLayoutsAsync())
            .ReturnsAsync(new List<LayoutInfo> { new("FileLayout1", 100, false, "") });
        _mockDrawingDataService.Setup(s => s.ExtractParametersAsync("FileLayout1"))
            .ReturnsAsync(layoutData);

        var result = new BOQGenerationResult
        {
            Success = true,
            Entries = new List<BOQEntry>
            {
                new() { ProductName = "Widget", Quantity = 10, UnitOfMeasure = "pcs" }
            }
        };
        _mockBoqProcessor
            .Setup(p => p.GenerateBOQAsync(layoutData, It.IsAny<BOQGenerationOptions>()))
            .ReturnsAsync(result);

        var sut = CreateSUT();
        await sut.ExtractBOQCommand.ExecuteAsync(null);

        sut.BoqItems.Should().HaveCount(1);
        sut.BoqItems[0].LayoutName.Should().Be("FileLayout1");
    }

    #endregion
}
