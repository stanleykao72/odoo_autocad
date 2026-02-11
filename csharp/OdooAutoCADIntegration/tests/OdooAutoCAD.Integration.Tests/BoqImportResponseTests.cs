using FluentAssertions;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class BoqImportResponseTests
{
    [Fact]
    public void ParseSuccessResponse_ExtractsHeaderAndDetailIds()
    {
        var response = new BoqImportResponse
        {
            Success = true,
            All = new List<BoqImportLayout>
            {
                new()
                {
                    LayoutName = "Layout1",
                    HeaderId = "H001",
                    Detail = new List<BoqImportDetail>
                    {
                        new() { ProductNo = "A001", DetailId = "D001" },
                        new() { ProductNo = "B002", DetailId = "D002" }
                    }
                }
            }
        };

        response.Success.Should().BeTrue();
        response.All.Should().HaveCount(1);
        response.All![0].HeaderId.Should().Be("H001");
        response.All[0].Detail.Should().HaveCount(2);
        response.All[0].Detail[0].DetailId.Should().Be("D001");
        response.All[0].Detail[1].DetailId.Should().Be("D002");
    }

    [Fact]
    public void ParseErrorResponse_SetsErrorCodeAndMessage()
    {
        var response = new BoqImportResponse
        {
            Success = false,
            ErrorCode = "VALIDATION_ERROR",
            ErrorMessage = "Field 'product_no' is required for row 3"
        };

        response.Success.Should().BeFalse();
        response.ErrorCode.Should().Be("VALIDATION_ERROR");
        response.ErrorMessage.Should().Contain("product_no");
        response.All.Should().BeNull();
    }

    [Fact]
    public void BuildImportRequest_MapsLayoutDataCorrectly()
    {
        var layoutData = new LayoutData
        {
            LayoutName = "TestLayout",
            Parameters = new Dictionary<string, object>
            {
                ["pr_no"] = "PR-2026-001",
                ["project_name"] = "Bridge Project",
                ["product_name"] = "Steel Beam",
                ["color_name"] = "Silver"
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
                        new() { "Position", "Product No", "Width", "Height", "Length", "Thickness", "Qty", "Description", "HEADER_ID" },
                        new() { "1", "STL-001", "100", "200", "5000", "12", "8", "Main beam", "" },
                        new() { "2", "STL-002", "80", "150", "3000", "10", "4", "Cross beam", "EXISTING-42" }
                    }
                }
            }
        };

        // Use reflection to populate the layout data map
        var vm = CreateMinimalViewModel();
        var field = typeof(BOQViewModel).GetField("_layoutDataMap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        var map = (Dictionary<string, LayoutData>)field!.GetValue(vm)!;
        map["TestLayout"] = layoutData;

        var request = vm.BuildImportRequest();

        request.All.Should().HaveCount(1);
        var layout = request.All[0];
        layout.LayoutName.Should().Be("TestLayout");
        layout.PrNo.Should().Be("PR-2026-001");
        layout.ProjectName.Should().Be("Bridge Project");
        layout.ProductName.Should().Be("Steel Beam");
        layout.ColorName.Should().Be("Silver");
        layout.Detail.Should().HaveCount(2);

        layout.Detail[0].Position.Should().Be("1");
        layout.Detail[0].ProductNo.Should().Be("STL-001");
        layout.Detail[0].Width.Should().Be("100");
        layout.Detail[0].Height.Should().Be("200");
        layout.Detail[0].Length.Should().Be("5000");
        layout.Detail[0].Thickness.Should().Be("12");
        layout.Detail[0].Qty.Should().Be("8");
        layout.Detail[0].DetailId.Should().BeNull();

        layout.Detail[1].ProductNo.Should().Be("STL-002");
        layout.Detail[1].DetailId.Should().Be("EXISTING-42");
    }

    [Fact]
    public void BuildWritebackData_MapsResponseToWriteback()
    {
        var response = new BoqImportResponse
        {
            Success = true,
            All = new List<BoqImportLayout>
            {
                new()
                {
                    LayoutName = "Layout1",
                    HeaderId = "H100",
                    Detail = new List<BoqImportDetail>
                    {
                        new() { ProductNo = "A001", DetailId = "D100" },
                        new() { ProductNo = "B002", DetailId = "D200" }
                    }
                },
                new()
                {
                    LayoutName = "Layout2",
                    HeaderId = "H200",
                    Detail = new List<BoqImportDetail>
                    {
                        new() { ProductNo = "C003", DetailId = "D300" }
                    }
                }
            }
        };

        var writebacks = BOQViewModel.BuildWritebackData(response);

        writebacks.Should().HaveCount(2);

        writebacks[0].LayoutName.Should().Be("Layout1");
        writebacks[0].HeaderId.Should().Be("H100");
        writebacks[0].Details.Should().HaveCount(2);
        writebacks[0].Details[0].ProductNo.Should().Be("A001");
        writebacks[0].Details[0].DetailId.Should().Be("D100");
        writebacks[0].Details[1].ProductNo.Should().Be("B002");
        writebacks[0].Details[1].DetailId.Should().Be("D200");

        writebacks[1].LayoutName.Should().Be("Layout2");
        writebacks[1].HeaderId.Should().Be("H200");
        writebacks[1].Details.Should().HaveCount(1);
        writebacks[1].Details[0].DetailId.Should().Be("D300");
    }

    private static BOQViewModel CreateMinimalViewModel()
    {
        var mockBoq = new Moq.Mock<IBOQProcessor>();
        var mockAcad = new Moq.Mock<IAutoCADService>();
        var mockOdoo = new Moq.Mock<OdooAutoCAD.Core.Odoo.IOdooService>();
        var mockProxy = new Moq.Mock<OdooAutoCAD.Core.Threading.IGUIProxy>();
        var mockLog = new Moq.Mock<OdooAutoCAD.App.Services.IAppLogService>();
        var mockSettings = new Moq.Mock<OdooAutoCAD.App.Services.ISettingsService>();

        return new BOQViewModel(
            mockBoq.Object, mockAcad.Object, mockOdoo.Object,
            mockProxy.Object, mockLog.Object, mockSettings.Object);
    }
}
