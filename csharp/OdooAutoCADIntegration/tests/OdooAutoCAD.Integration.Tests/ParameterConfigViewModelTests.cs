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

    public ParameterConfigViewModelTests()
    {
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockOdoo = new Mock<IOdooService>();
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockSettingsService = new Mock<ISettingsService>();
        _mockLogService = new Mock<IAppLogService>();
    }

    private ParameterConfigViewModel CreateSUT()
    {
        return new ParameterConfigViewModel(
            _mockAutoCAD.Object,
            _mockOdoo.Object,
            _mockGuiProxy.Object,
            _mockSettingsService.Object,
            _mockLogService.Object);
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
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.CanSubmit.Should().BeTrue();
    }

    [Fact]
    public void CanSubmit_WhenAutoCADDisconnected_ReturnsFalse()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(false);
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.CanSubmit.Should().BeFalse();
    }

    [Fact]
    public void CanSubmit_WhenIsSubmitting_ReturnsFalse()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
        var sut = CreateSUT();
        SetAllSelections(sut);

        sut.IsSubmitting = true;

        sut.CanSubmit.Should().BeFalse();
    }

    [Fact]
    public void CanSubmit_WhenFieldsMissing_ReturnsFalse()
    {
        _mockAutoCAD.Setup(s => s.IsConnected).Returns(true);
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
