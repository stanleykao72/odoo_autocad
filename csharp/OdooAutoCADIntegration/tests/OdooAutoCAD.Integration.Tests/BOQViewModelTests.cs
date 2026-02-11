using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class BOQViewModelTests
{
    private readonly Mock<IBOQProcessor> _mockBoqProcessor;
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<IAppLogService> _mockLogService;

    public BOQViewModelTests()
    {
        _mockBoqProcessor = new Mock<IBOQProcessor>();
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockOdoo = new Mock<IOdooService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockLogService = new Mock<IAppLogService>();
    }

    private BOQViewModel CreateSUT()
    {
        return new BOQViewModel(
            _mockBoqProcessor.Object,
            _mockAutoCAD.Object,
            _mockOdoo.Object,
            _mockGuiProxy.Object,
            _mockLogService.Object);
    }

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
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();

        sut.ExtractBOQCommand.CanExecute(null).Should().BeTrue();
    }

    [Fact]
    public async Task ExtractBOQ_NoLayouts_SetsStatusMessage()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_layouts", null, 10000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", new List<string>()));

        var sut = CreateSUT();
        await sut.ExtractBOQCommand.ExecuteAsync(null);

        sut.ExtractionStatus.Should().Contain("No layouts found");
        sut.BoqItems.Should().BeEmpty();
        sut.TotalItems.Should().Be(0);
    }

    [Fact]
    public async Task ExtractBOQ_WithLayouts_PopulatesItems()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);

        var layouts = new List<string> { "Layout1", "Layout2" };
        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_layouts", null, 10000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req1", layouts));

        var layoutData1 = new LayoutData { LayoutName = "Layout1" };
        var layoutData2 = new LayoutData { LayoutName = "Layout2" };

        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_extract_parameters",
                It.Is<Dictionary<string, object?>>(d => d["layoutName"] as string == "Layout1"), 30000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req2", layoutData1));

        _mockGuiProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_extract_parameters",
                It.Is<Dictionary<string, object?>>(d => d["layoutName"] as string == "Layout2"), 30000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req3", layoutData2));

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
}
