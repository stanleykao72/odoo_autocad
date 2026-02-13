using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class ParameterConfigViewModelTests
{
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IOdooService> _mockOdoo;
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<ISettingsService> _mockSettingsService;
    private readonly Mock<IAppLogService> _mockLogService;
    private readonly Mock<IDrawingDataService> _mockDrawingDataService;

    public ParameterConfigViewModelTests()
    {
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockOdoo = new Mock<IOdooService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockSettingsService = new Mock<ISettingsService>();
        _mockLogService = new Mock<IAppLogService>();
        _mockDrawingDataService = new Mock<IDrawingDataService>();
    }

    private ParameterConfigViewModel CreateSUT()
    {
        return new ParameterConfigViewModel(
            _mockAutoCAD.Object,
            _mockOdoo.Object,
            _mockGuiProxy.Object,
            _mockSettingsService.Object,
            _mockLogService.Object,
            _mockDrawingDataService.Object);
    }

    #region Constructor / Default State

    [Fact]
    public void Constructor_DefaultsAllSelectionsToNull()
    {
        var sut = CreateSUT();

        sut.SelectedProduct.Should().BeNull();
        sut.SelectedSpec.Should().BeNull();
        sut.SelectedCategory.Should().BeNull();
        sut.SelectedOperationFlow.Should().BeNull();
        sut.SelectedSurfaceTreatment.Should().BeNull();
        sut.SelectedColor.Should().BeNull();
    }

    [Fact]
    public void Constructor_DefaultsAutoFillFieldsToEmpty()
    {
        var sut = CreateSUT();

        sut.Unit.Should().BeEmpty();
        sut.ColorNo.Should().BeEmpty();
    }

    [Fact]
    public void Constructor_DefaultsCollectionsToEmpty()
    {
        var sut = CreateSUT();

        sut.Products.Should().BeEmpty();
        sut.Specs.Should().BeEmpty();
        sut.Categories.Should().BeEmpty();
        sut.OperationFlows.Should().BeEmpty();
        sut.SurfaceTreatments.Should().BeEmpty();
        sut.Colors.Should().BeEmpty();
    }

    [Fact]
    public void Constructor_DefaultsStateFlags()
    {
        var sut = CreateSUT();

        sut.IsLoading.Should().BeFalse();
        sut.IsSubmitting.Should().BeFalse();
        sut.StatusMessage.Should().BeEmpty();
        sut.LastSubmitResult.Should().BeEmpty();
    }

    #endregion

    #region AllFieldsFilled

    [Fact]
    public void AllFieldsFilled_WhenNoSelections_ReturnsFalse()
    {
        var sut = CreateSUT();

        sut.AllFieldsFilled.Should().BeFalse();
    }

    [Fact]
    public void AllFieldsFilled_WhenAllSelected_ReturnsTrue()
    {
        var sut = CreateSUT();

        sut.SelectedProduct = new OdooProduct(1, "Steel", null, null, null, "kg", null);
        sut.SelectedSpec = new OdooSetupValue("3mm", "spec");
        sut.SelectedCategory = new OdooSetupValue("Plate", "product_catelog");
        sut.SelectedOperationFlow = new OdooSetupValue("Cut", "operation_flow");
        sut.SelectedSurfaceTreatment = new OdooSetupValue("Paint", "surface_treatment");
        sut.SelectedColor = new OdooColor("Red", "R001", 1);

        sut.AllFieldsFilled.Should().BeTrue();
    }

    [Fact]
    public void AllFieldsFilled_WhenOneSelectionMissing_ReturnsFalse()
    {
        var sut = CreateSUT();

        sut.SelectedProduct = new OdooProduct(1, "Steel", null, null, null, "kg", null);
        sut.SelectedSpec = new OdooSetupValue("3mm", "spec");
        sut.SelectedCategory = new OdooSetupValue("Plate", "product_catelog");
        sut.SelectedOperationFlow = new OdooSetupValue("Cut", "operation_flow");
        sut.SelectedSurfaceTreatment = new OdooSetupValue("Paint", "surface_treatment");
        // SelectedColor is null

        sut.AllFieldsFilled.Should().BeFalse();
    }

    [Fact]
    public void AllFieldsFilled_WhenColorNoEmpty_ReturnsFalse()
    {
        var sut = CreateSUT();

        sut.SelectedProduct = new OdooProduct(1, "Steel", null, null, null, "kg", null);
        sut.SelectedSpec = new OdooSetupValue("3mm", "spec");
        sut.SelectedCategory = new OdooSetupValue("Plate", "product_catelog");
        sut.SelectedOperationFlow = new OdooSetupValue("Cut", "operation_flow");
        sut.SelectedSurfaceTreatment = new OdooSetupValue("Paint", "surface_treatment");
        sut.SelectedColor = new OdooColor("Transparent", "", 1); // empty ColorNo

        sut.AllFieldsFilled.Should().BeFalse();
    }

    #endregion

    #region Auto-fill

    [Fact]
    public void SelectedProduct_SetsUnitAutomatically()
    {
        var sut = CreateSUT();

        sut.SelectedProduct = new OdooProduct(1, "Steel Sheet", null, null, null, "pcs", null);

        sut.Unit.Should().Be("pcs");
    }

    [Fact]
    public void SelectedProduct_ClearsUnitWhenNull()
    {
        var sut = CreateSUT();
        sut.SelectedProduct = new OdooProduct(1, "Steel", null, null, null, "kg", null);

        sut.SelectedProduct = null;

        sut.Unit.Should().BeEmpty();
    }

    [Fact]
    public void SelectedColor_SetsColorNoAutomatically()
    {
        var sut = CreateSUT();

        sut.SelectedColor = new OdooColor("Navy Blue", "NB-100", 5);

        sut.ColorNo.Should().Be("NB-100");
    }

    [Fact]
    public void SelectedColor_ClearsColorNoWhenNull()
    {
        var sut = CreateSUT();
        sut.SelectedColor = new OdooColor("Red", "R001", 1);

        sut.SelectedColor = null;

        sut.ColorNo.Should().BeEmpty();
    }

    #endregion

    #region CanSubmit

    [Fact]
    public void CanSubmit_WhenAllFieldsFilledAndAutoCADConnected_ReturnsTrue()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.CanSubmit.Should().BeTrue();
    }

    [Fact]
    public void CanSubmit_WhenAutoCADDisconnected_ReturnsFalse()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(false);
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.CanSubmit.Should().BeFalse();
    }

    [Fact]
    public void CanSubmit_WhenIsSubmitting_ReturnsFalse()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.IsSubmitting = true;

        sut.CanSubmit.Should().BeFalse();
    }

    [Fact]
    public void CanSubmit_WhenFieldsMissing_ReturnsFalse()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        var sut = CreateSUT();
        // Don't set any selections

        sut.CanSubmit.Should().BeFalse();
    }

    #endregion

    #region Cancel Command

    [Fact]
    public void Cancel_ClearsAllSelections()
    {
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.CancelCommand.Execute(null);

        sut.SelectedProduct.Should().BeNull();
        sut.SelectedSpec.Should().BeNull();
        sut.SelectedCategory.Should().BeNull();
        sut.SelectedOperationFlow.Should().BeNull();
        sut.SelectedSurfaceTreatment.Should().BeNull();
        sut.SelectedColor.Should().BeNull();
        sut.Unit.Should().BeEmpty();
        sut.ColorNo.Should().BeEmpty();
        sut.LastSubmitResult.Should().BeEmpty();
    }

    #endregion

    #region Record Type Tests

    [Fact]
    public void OdooSetupValue_RecordEquality()
    {
        var a = new OdooSetupValue("3mm", "spec");
        var b = new OdooSetupValue("3mm", "spec");

        a.Should().Be(b);
        (a == b).Should().BeTrue();
    }

    [Fact]
    public void OdooColor_RecordEquality()
    {
        var a = new OdooColor("Red", "R001", 1);
        var b = new OdooColor("Red", "R001", 1);

        a.Should().Be(b);
        (a == b).Should().BeTrue();
    }

    [Fact]
    public void OdooSetupValue_RecordInequality()
    {
        var a = new OdooSetupValue("3mm", "spec");
        var b = new OdooSetupValue("5mm", "spec");

        a.Should().NotBe(b);
    }

    [Fact]
    public void OdooColor_RecordInequality()
    {
        var a = new OdooColor("Red", "R001", 1);
        var b = new OdooColor("Blue", "B001", 1);

        a.Should().NotBe(b);
    }

    #endregion

    #region Dual-Mode Tests

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
    public void CanSubmit_InFileMode_WhenReady_ReturnsTrue()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.File);
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.CanSubmit.Should().BeTrue();
    }

    #endregion

    #region Layout Picker

    [Fact]
    public void Layouts_DefaultsToEmpty()
    {
        var sut = CreateSUT();
        sut.Layouts.Should().BeEmpty();
    }

    [Fact]
    public void SelectedLayout_DefaultsToNull()
    {
        var sut = CreateSUT();
        sut.SelectedLayout.Should().BeNull();
    }

    [Fact]
    public void ApplyToAllLayouts_DefaultsToFalse()
    {
        var sut = CreateSUT();
        sut.ApplyToAllLayouts.Should().BeFalse();
    }

    [Fact]
    public void SelectedLayout_UpdatesLayoutSummary()
    {
        var sut = CreateSUT();

        sut.SelectedLayout = "S405-201";

        sut.LayoutSummary.Should().Be("Will update: S405-201");
    }

    [Fact]
    public void ApplyToAllLayouts_True_UpdatesLayoutSummary()
    {
        var sut = CreateSUT();

        sut.ApplyToAllLayouts = true;

        sut.LayoutSummary.Should().Be("Will update: All layouts");
    }

    [Fact]
    public void ApplyToAllLayouts_False_WithSelection_UpdatesLayoutSummary()
    {
        var sut = CreateSUT();
        sut.SelectedLayout = "Layout1";

        sut.ApplyToAllLayouts = true;
        sut.ApplyToAllLayouts = false;

        sut.LayoutSummary.Should().Be("Will update: Layout1");
    }

    [Fact]
    public void Cancel_ResetsApplyToAllLayouts_ForCOMMode()
    {
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.COM);
        var sut = CreateSUT();
        sut.ApplyToAllLayouts = true;

        sut.CancelCommand.Execute(null);

        sut.ApplyToAllLayouts.Should().BeFalse();
    }

    [Fact]
    public void Cancel_ResetsApplyToAllLayouts_ForFileMode()
    {
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.File);
        var sut = CreateSUT();
        sut.ApplyToAllLayouts = false;

        sut.CancelCommand.Execute(null);

        sut.ApplyToAllLayouts.Should().BeTrue();
    }

    [Fact]
    public void Cancel_ResetsSelectedLayout_ToFirst()
    {
        var sut = CreateSUT();
        sut.Layouts.Add("Layout1");
        sut.Layouts.Add("Layout2");
        sut.SelectedLayout = "Layout2";

        sut.CancelCommand.Execute(null);

        sut.SelectedLayout.Should().Be("Layout1");
    }

    [Fact]
    public async Task SubmitAsync_ApplyToAll_PassesNullLayout()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService.Setup(s => s.SetAttributeValuesAsync(
                It.IsAny<Dictionary<string, string>>(), null))
            .ReturnsAsync(new List<string> { "product_name" });
        var sut = CreateSUT();
        SetAllSelections(sut);
        sut.ApplyToAllLayouts = true;

        await sut.SubmitCommand.ExecuteAsync(null);

        _mockDrawingDataService.Verify(s => s.SetAttributeValuesAsync(
            It.IsAny<Dictionary<string, string>>(), null), Times.Once);
        sut.LastSubmitResult.Should().Contain("All layouts");
    }

    [Fact]
    public async Task SubmitAsync_SingleLayout_PassesSelectedLayout()
    {
        _mockDrawingDataService.Setup(s => s.IsReady).Returns(true);
        _mockDrawingDataService.Setup(s => s.SetAttributeValuesAsync(
                It.IsAny<Dictionary<string, string>>(), "S405-201"))
            .ReturnsAsync(new List<string> { "product_name" });
        var sut = CreateSUT();
        SetAllSelections(sut);
        sut.ApplyToAllLayouts = false;
        sut.SelectedLayout = "S405-201";

        await sut.SubmitCommand.ExecuteAsync(null);

        _mockDrawingDataService.Verify(s => s.SetAttributeValuesAsync(
            It.IsAny<Dictionary<string, string>>(), "S405-201"), Times.Once);
        sut.LastSubmitResult.Should().Contain("S405-201");
    }

    [Fact]
    public async Task LoadOptionsAsync_PopulatesLayouts()
    {
        // Arrange
        _mockDrawingDataService.Setup(s => s.GetLayoutsAsync())
            .ReturnsAsync(new List<LayoutInfo>
            {
                new("Layout1", 1, false, ""),
                new("Layout2", 2, false, "")
            });
        _mockDrawingDataService.Setup(s => s.Mode).Returns(AutoCADOperationMode.File);
        _mockDrawingDataService.Setup(s => s.GetPRNumberAsync())
            .ReturnsAsync(string.Empty);
        _mockSettingsService.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = "https://odoo.test/api/v1/boq_import_api/swagger.json?token=abc&db=testdb",
                ["odoo_user_token"] = "test"
            });

        var sut = CreateSUT();

        // Act
        await sut.LoadOptionsCommand.ExecuteAsync(null);

        // Assert
        sut.Layouts.Should().HaveCount(2);
        sut.Layouts.Should().Contain("Layout1");
        sut.Layouts.Should().Contain("Layout2");
        sut.SelectedLayout.Should().Be("Layout1");
        sut.ApplyToAllLayouts.Should().BeTrue(); // File mode defaults to all
    }

    #endregion

    #region Helpers

    private static void SetAllSelections(ParameterConfigViewModel vm)
    {
        vm.SelectedProduct = new OdooProduct(1, "Steel", null, null, null, "kg", null);
        vm.SelectedSpec = new OdooSetupValue("3mm", "spec");
        vm.SelectedCategory = new OdooSetupValue("Plate", "product_catelog");
        vm.SelectedOperationFlow = new OdooSetupValue("Cut", "operation_flow");
        vm.SelectedSurfaceTreatment = new OdooSetupValue("Paint", "surface_treatment");
        vm.SelectedColor = new OdooColor("Red", "R001", 1);
    }

    #endregion
}
