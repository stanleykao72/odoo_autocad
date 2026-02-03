// OdooAutoCAD.MCP.Tests/MCPToolRegistryTests.cs
// Unit tests for MCP Tool Registry

using FluentAssertions;
using Moq;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.MCP.Protocol;
using OdooAutoCAD.MCP.Tools;
using OdooAutoCAD.Threading;
using Xunit;

namespace OdooAutoCAD.MCP.Tests;

public class MCPToolRegistryTests
{
    private readonly Mock<IGUIProxy> _mockGuiProxy;
    private readonly Mock<IAutoCADService> _mockAutoCADService;
    private readonly Mock<IOdooService> _mockOdooService;
    private readonly Mock<IBOQProcessor> _mockBOQProcessor;
    private readonly MCPToolRegistry _registry;

    public MCPToolRegistryTests()
    {
        _mockGuiProxy = new Mock<IGUIProxy>();
        _mockAutoCADService = new Mock<IAutoCADService>();
        _mockOdooService = new Mock<IOdooService>();
        _mockBOQProcessor = new Mock<IBOQProcessor>();

        _registry = new MCPToolRegistry(
            _mockGuiProxy.Object,
            _mockAutoCADService.Object,
            _mockOdooService.Object,
            _mockBOQProcessor.Object);
    }

    [Fact]
    public void GetTools_ShouldReturn7Tools()
    {
        // Act
        var tools = _registry.GetTools();

        // Assert
        tools.Should().HaveCount(7);
    }

    [Theory]
    [InlineData("test_connection")]
    [InlineData("get_server_info")]
    [InlineData("check_autocad_status")]
    [InlineData("check_odoo_status")]
    [InlineData("extract_autocad_parameters")]
    [InlineData("sync_to_odoo")]
    [InlineData("generate_boq")]
    public void GetTools_ShouldContainExpectedTool(string toolName)
    {
        // Act
        var tools = _registry.GetTools();

        // Assert
        tools.Should().Contain(t => t.Name == toolName);
    }

    [Fact]
    public async Task ExecuteToolAsync_TestConnection_ShouldReturnSuccess()
    {
        // Act
        var result = await _registry.ExecuteToolAsync("test_connection", null);

        // Assert
        result.IsError.Should().BeFalse();
        result.Content.Should().NotBeEmpty();
        result.Content[0].Text.Should().Contain("connected");
    }

    [Fact]
    public async Task ExecuteToolAsync_GetServerInfo_ShouldReturnServerInfo()
    {
        // Arrange
        _mockGuiProxy.Setup(p => p.IsRunning).Returns(true);
        _mockGuiProxy.Setup(p => p.PendingRequestCount).Returns(0);

        // Act
        var result = await _registry.ExecuteToolAsync("get_server_info", null);

        // Assert
        result.IsError.Should().BeFalse();
        result.Content.Should().NotBeEmpty();
        result.Content[0].Text.Should().Contain("OdooAutoCAD");
        result.Content[0].Text.Should().Contain("6.0.0");
    }

    [Fact]
    public async Task ExecuteToolAsync_UnknownTool_ShouldReturnError()
    {
        // Act
        var result = await _registry.ExecuteToolAsync("unknown_tool", null);

        // Assert
        result.IsError.Should().BeTrue();
        result.Content[0].Text.Should().Contain("not found");
    }

    [Fact]
    public async Task ExecuteToolAsync_CheckOdooStatus_ShouldCallOdooService()
    {
        // Arrange
        _mockOdooService.Setup(s => s.GetStatusAsync())
            .ReturnsAsync(new OdooStatus(
                IsConnected: true,
                ServerUrl: "https://odoo.test.com",
                Database: "test_db",
                Username: "admin",
                Version: "16.0",
                ErrorMessage: null));

        // Act
        var result = await _registry.ExecuteToolAsync("check_odoo_status", null);

        // Assert
        result.IsError.Should().BeFalse();
        _mockOdooService.Verify(s => s.GetStatusAsync(), Times.Once);
    }

    [Fact]
    public async Task ExecuteToolAsync_CheckAutoCADStatus_ShouldUseGUIProxy()
    {
        // Arrange
        _mockGuiProxy.Setup(p => p.ExecuteInGuiAsync(
            "check_autocad_status",
            It.IsAny<Dictionary<string, object?>>(),
            It.IsAny<int>()))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("req-1", new { connected = true }));

        // Act
        var result = await _registry.ExecuteToolAsync("check_autocad_status", null);

        // Assert
        _mockGuiProxy.Verify(p => p.ExecuteInGuiAsync(
            "check_autocad_status",
            It.IsAny<Dictionary<string, object?>>(),
            It.IsAny<int>()), Times.Once);
    }

    [Fact]
    public async Task ExecuteToolAsync_SyncToOdoo_WithMissingParams_ShouldReturnError()
    {
        // Act
        var result = await _registry.ExecuteToolAsync("sync_to_odoo", null);

        // Assert
        result.IsError.Should().BeTrue();
        result.Content[0].Text.Should().Contain("Missing required parameters");
    }

    [Fact]
    public async Task ExecuteToolAsync_GenerateBOQ_WithMissingProjectId_ShouldReturnError()
    {
        // Act
        var result = await _registry.ExecuteToolAsync("generate_boq", new Dictionary<string, object?>());

        // Assert
        result.IsError.Should().BeTrue();
        result.Content[0].Text.Should().Contain("project_id");
    }

    [Fact]
    public async Task ExecuteToolAsync_GenerateBOQ_WithValidParams_ShouldCallBOQProcessor()
    {
        // Arrange
        var arguments = new Dictionary<string, object?>
        {
            ["project_id"] = 123,
            ["include_autocad_data"] = true
        };

        _mockBOQProcessor.Setup(p => p.GenerateBOQForProjectAsync(123, true))
            .ReturnsAsync(new BOQGenerationResult
            {
                Success = true,
                TotalItems = 5,
                ProcessedItems = 5
            });

        // Act
        var result = await _registry.ExecuteToolAsync("generate_boq", arguments);

        // Assert
        result.IsError.Should().BeFalse();
        _mockBOQProcessor.Verify(p => p.GenerateBOQForProjectAsync(123, true), Times.Once);
    }
}

public class MCPModelsTests
{
    [Fact]
    public void JsonRpcRequest_ShouldHaveDefaultJsonRpcVersion()
    {
        // Arrange & Act
        var request = new JsonRpcRequest();

        // Assert
        request.JsonRpc.Should().Be("2.0");
    }

    [Fact]
    public void JsonRpcResponse_Success_ShouldHaveNoError()
    {
        // Arrange & Act
        var response = JsonRpcResponse.Success(1, "result");

        // Assert
        response.JsonRpc.Should().Be("2.0");
        response.Id.Should().Be(1);
        response.Result.Should().Be("result");
        response.Error.Should().BeNull();
    }

    [Fact]
    public void JsonRpcResponse_Failure_ShouldHaveNoResult()
    {
        // Arrange & Act
        var response = JsonRpcResponse.Failure(1, -32600, "Invalid Request");

        // Assert
        response.JsonRpc.Should().Be("2.0");
        response.Id.Should().Be(1);
        response.Result.Should().BeNull();
        response.Error.Should().NotBeNull();
        response.Error!.Code.Should().Be(-32600);
        response.Error.Message.Should().Be("Invalid Request");
    }

    [Fact]
    public void MCPContent_Text_ShouldCreateTextContent()
    {
        // Arrange & Act
        var content = MCPContent.Text("Hello World");

        // Assert
        content.Type.Should().Be("text");
        content.Text.Should().Be("Hello World");
    }

    [Fact]
    public void MCPContent_Image_ShouldCreateImageContent()
    {
        // Arrange & Act
        var content = MCPContent.Image("base64data", "image/png");

        // Assert
        content.Type.Should().Be("image");
        content.Data.Should().Be("base64data");
        content.MimeType.Should().Be("image/png");
    }

    [Fact]
    public void SSEEvent_ToString_ShouldFormatCorrectly()
    {
        // Arrange
        var evt = new SSEEvent
        {
            Event = "message",
            Data = "test data",
            Id = "123"
        };

        // Act
        var result = evt.ToString();

        // Assert
        result.Should().Contain("event: message");
        result.Should().Contain("id: 123");
        result.Should().Contain("data: test data");
    }
}
