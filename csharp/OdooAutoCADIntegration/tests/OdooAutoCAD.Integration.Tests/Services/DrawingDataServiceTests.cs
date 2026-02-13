using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Moq;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Threading;
using Xunit;

namespace OdooAutoCAD.Integration.Tests.Services;

public class ComDrawingDataServiceTests
{
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IGUIProxy> _mockProxy;
    private readonly ComDrawingDataService _service;

    public ComDrawingDataServiceTests()
    {
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockProxy = new Mock<IGUIProxy>();
        _service = new ComDrawingDataService(_mockAutoCAD.Object, _mockProxy.Object);
    }

    [Fact]
    public void Mode_ReturnsCOM()
    {
        Assert.Equal(AutoCADOperationMode.COM, _service.Mode);
    }

    [Fact]
    public void SupportsWrite_ReturnsTrue()
    {
        Assert.True(_service.SupportsWrite);
    }

    [Fact]
    public void RecommendedWriteStrategy_ReturnsDirectDwg()
    {
        Assert.Equal(WriteStrategy.DirectDwg, _service.RecommendedWriteStrategy);
    }

    [Fact]
    public void IsReady_DelegatesToIsConnected()
    {
        _mockAutoCAD.Setup(a => a.IsConnected).Returns(true);
        Assert.True(_service.IsReady);

        _mockAutoCAD.Setup(a => a.IsConnected).Returns(false);
        Assert.False(_service.IsReady);
    }

    [Fact]
    public void CurrentSource_DelegatesToGetCurrentDocumentPath()
    {
        _mockAutoCAD.Setup(a => a.GetCurrentDocumentPath()).Returns("C:\\test.dwg");
        Assert.Equal("C:\\test.dwg", _service.CurrentSource);
    }

    [Fact]
    public async Task ConnectOrLoadAsync_DelegatesToGUIProxy()
    {
        _mockProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_connect", null, 15000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("r1", true));

        var result = await _service.ConnectOrLoadAsync();
        Assert.True(result);
        _mockProxy.Verify(p => p.ExecuteInGuiAsync("autocad_connect", null, 15000), Times.Once);
    }

    [Fact]
    public async Task DisconnectOrUnloadAsync_DelegatesToGUIProxy()
    {
        _mockProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_disconnect", null, 10000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("r1", null));

        await _service.DisconnectOrUnloadAsync();
        _mockProxy.Verify(p => p.ExecuteInGuiAsync("autocad_disconnect", null, 10000), Times.Once);
    }

    [Fact]
    public async Task GetLayoutsAsync_DelegatesToGUIProxy()
    {
        var layouts = new List<LayoutInfo>
        {
            new("Layout1", 1, false, ""),
            new("Layout2", 2, false, "")
        };

        _mockProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_layouts", null, 10000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("r1", layouts));

        var result = await _service.GetLayoutsAsync();
        Assert.Equal(2, result.Count);
        Assert.Equal("Layout1", result[0].Name);
    }

    [Fact]
    public async Task ExtractParametersAsync_DelegatesToGUIProxy()
    {
        var data = new LayoutData { LayoutName = "S405-201" };
        _mockProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_extract_parameters",
                It.Is<Dictionary<string, object?>>(d => d.ContainsKey("layoutName")), 10000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("r1", data));

        var result = await _service.ExtractParametersAsync("S405-201");
        Assert.Equal("S405-201", result.LayoutName);
    }

    [Fact]
    public async Task GetHeaderIdsAsync_DelegatesToGUIProxy()
    {
        var ids = new List<string> { "H-001", "H-002" };
        _mockProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_get_header_ids", null, 10000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("r1", ids));

        var result = await _service.GetHeaderIdsAsync();
        Assert.Equal(2, result.Count);
    }

    [Fact]
    public async Task WriteTableIdsAsync_DelegatesToGUIProxy()
    {
        _mockProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_write_table_ids",
                It.IsAny<Dictionary<string, object?>>(), 30000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("r1", new List<string> { "OK" }));

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };
        var result = await _service.WriteTableIdsAsync("Layout1", "H-001", details);
        Assert.True(result.Success);
        Assert.Equal(WriteStrategy.DirectDwg, result.StrategyUsed);
    }

    [Fact]
    public async Task SaveAsync_ReturnsTrue()
    {
        var result = await _service.SaveAsync();
        Assert.True(result);
    }
}

public class DrawingDataServiceDispatcherTests
{
    private readonly Mock<IAutoCADService> _mockAutoCAD;
    private readonly Mock<IGUIProxy> _mockProxy;
    private readonly Mock<IDwgFileService> _mockDwgFile;
    private readonly ComDrawingDataService _comService;
    private readonly FileDrawingDataService _fileService;

    public DrawingDataServiceDispatcherTests()
    {
        _mockAutoCAD = new Mock<IAutoCADService>();
        _mockProxy = new Mock<IGUIProxy>();
        _mockDwgFile = new Mock<IDwgFileService>();

        _comService = new ComDrawingDataService(_mockAutoCAD.Object, _mockProxy.Object);
        _fileService = new FileDrawingDataService(_mockDwgFile.Object, new SidecarIdStore());
    }

    [Fact]
    public void Constructor_DefaultMode_IsCOM()
    {
        var dispatcher = new DrawingDataServiceDispatcher(_comService, _fileService);
        Assert.Equal(AutoCADOperationMode.COM, dispatcher.Mode);
    }

    [Fact]
    public void Constructor_ExplicitFileMode()
    {
        var dispatcher = new DrawingDataServiceDispatcher(_comService, _fileService, AutoCADOperationMode.File);
        Assert.Equal(AutoCADOperationMode.File, dispatcher.Mode);
    }

    [Fact]
    public async Task SwitchMode_COMToFile_ChangesMode()
    {
        var dispatcher = new DrawingDataServiceDispatcher(_comService, _fileService);
        AutoCADOperationMode? changedTo = null;
        dispatcher.ModeChanged += (_, mode) => changedTo = mode;

        await dispatcher.SwitchModeAsync(AutoCADOperationMode.File);

        Assert.Equal(AutoCADOperationMode.File, dispatcher.Mode);
        Assert.Equal(AutoCADOperationMode.File, changedTo);
    }

    [Fact]
    public async Task SwitchMode_SameMode_DoesNothing()
    {
        var dispatcher = new DrawingDataServiceDispatcher(_comService, _fileService);
        bool modeChangeFired = false;
        dispatcher.ModeChanged += (_, _) => modeChangeFired = true;

        await dispatcher.SwitchModeAsync(AutoCADOperationMode.COM);

        Assert.False(modeChangeFired);
    }

    [Fact]
    public async Task SwitchMode_DisconnectsOldBackend()
    {
        _mockAutoCAD.Setup(a => a.IsConnected).Returns(true);
        _mockProxy
            .Setup(p => p.ExecuteInGuiAsync("autocad_disconnect", null, 10000))
            .ReturnsAsync(GUIProxyResponse.CreateSuccess("r1", null));

        var dispatcher = new DrawingDataServiceDispatcher(_comService, _fileService);
        await dispatcher.SwitchModeAsync(AutoCADOperationMode.File);

        _mockProxy.Verify(p => p.ExecuteInGuiAsync("autocad_disconnect", null, 10000), Times.Once);
    }

    [Fact]
    public void SupportsWrite_DelegatesToActiveBackend()
    {
        var dispatcher = new DrawingDataServiceDispatcher(_comService, _fileService);
        Assert.True(dispatcher.SupportsWrite); // COM SupportsWrite = true
    }

    [Fact]
    public async Task RecommendedWriteStrategy_ChangesWithMode()
    {
        var dispatcher = new DrawingDataServiceDispatcher(_comService, _fileService);

        Assert.Equal(WriteStrategy.DirectDwg, dispatcher.RecommendedWriteStrategy);

        await dispatcher.SwitchModeAsync(AutoCADOperationMode.File);

        Assert.Equal(WriteStrategy.ExportDxf, dispatcher.RecommendedWriteStrategy);
    }

    [Fact]
    public async Task Delegation_FileMode_UsesFileService()
    {
        _mockDwgFile.Setup(d => d.IsFileLoaded).Returns(true);
        _mockDwgFile.Setup(d => d.LoadedFilePath).Returns("C:\\test.dwg");
        _mockDwgFile.Setup(d => d.GetLayouts("C:\\test.dwg"))
            .Returns(new List<LayoutInfo> { new("L1", 1, false, "") });

        var dispatcher = new DrawingDataServiceDispatcher(
            _comService, _fileService, AutoCADOperationMode.File);

        var layouts = await dispatcher.GetLayoutsAsync();
        Assert.Single(layouts);
        Assert.Equal("L1", layouts[0].Name);
    }
}

public class FileDrawingDataServiceWriteTableIdsTests : IDisposable
{
    private readonly string _tempDir;
    private readonly string _tempDwgPath;
    private readonly Mock<IDwgFileService> _mockDwgFile;
    private readonly SidecarIdStore _sidecarStore;
    private readonly FileDrawingDataService _service;

    public FileDrawingDataServiceWriteTableIdsTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"fds_wt_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _tempDwgPath = Path.Combine(_tempDir, "test.dwg");
        File.WriteAllText(_tempDwgPath, "dummy");

        _mockDwgFile = new Mock<IDwgFileService>();
        _mockDwgFile.Setup(d => d.IsFileLoaded).Returns(true);
        _mockDwgFile.Setup(d => d.LoadedFilePath).Returns(_tempDwgPath);

        _sidecarStore = new SidecarIdStore();
        _service = new FileDrawingDataService(_mockDwgFile.Object, _sidecarStore);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, true); } catch { }
    }

    [Fact]
    public void RecommendedWriteStrategy_ReturnsExportDxf()
    {
        Assert.Equal(WriteStrategy.ExportDxf, _service.RecommendedWriteStrategy);
    }

    [Fact]
    public async Task WriteTableIdsAsync_CallsWriteTableIdsToDocument()
    {
        _mockDwgFile.Setup(d => d.SaveAsDxfAsync(It.IsAny<string>()))
            .ReturnsAsync(true);

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };

        var result = await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        // Verify it called WriteTableIdsToDocument on the DwgFileService
        _mockDwgFile.Verify(d => d.WriteTableIdsToDocument(
            "Layout1", "H-123",
            It.Is<IList<WritebackDetail>>(dl => dl.Count == 1 && dl[0].ProductNo == "ST-001")),
            Times.Once);
    }

    [Fact]
    public async Task WriteTableIdsAsync_ExportsDxfFile()
    {
        _mockDwgFile.Setup(d => d.SaveAsDxfAsync(It.IsAny<string>()))
            .ReturnsAsync(true);

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };

        var result = await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        Assert.True(result.Success);
        Assert.Equal(WriteStrategy.ExportDxf, result.StrategyUsed);
        Assert.Contains(".dxf", result.OutputPath);

        // Verify DXF export was attempted
        _mockDwgFile.Verify(d => d.SaveAsDxfAsync(
            It.Is<string>(p => p.EndsWith(".dxf"))), Times.Once);
    }

    [Fact]
    public async Task WriteTableIdsAsync_CreatesSidecarJsonBackup()
    {
        _mockDwgFile.Setup(d => d.SaveAsDxfAsync(It.IsAny<string>()))
            .ReturnsAsync(true);

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" },
            new() { ProductNo = "ST-002", DetailId = "D-002" }
        };

        await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        // Verify sidecar was also saved as backup
        var sidecar = await _sidecarStore.LoadAsync(_tempDwgPath);
        Assert.NotNull(sidecar);
        Assert.Equal("H-123", sidecar!.Layouts[0].HeaderId);
        Assert.Equal(2, sidecar.Layouts[0].Details.Count);
        Assert.Equal("ST-001", sidecar.Layouts[0].Details[0].ProductNo);
    }

    [Fact]
    public async Task WriteTableIdsAsync_FallsBackToSidecarOnDxfFailure()
    {
        _mockDwgFile.Setup(d => d.SaveAsDxfAsync(It.IsAny<string>()))
            .ReturnsAsync(false); // DXF export fails

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };

        var result = await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        Assert.True(result.Success); // Still succeeds — sidecar as fallback
        Assert.Equal(WriteStrategy.SidecarJson, result.StrategyUsed);
        Assert.Contains("sidecar", result.Message);
    }

    [Fact]
    public async Task WriteTableIdsAsync_NotLoaded_ReturnsFalse()
    {
        _mockDwgFile.Setup(d => d.IsFileLoaded).Returns(false);
        _mockDwgFile.Setup(d => d.LoadedFilePath).Returns((string?)null);

        var result = await _service.WriteTableIdsAsync("Layout1", "H-123",
            new List<WritebackDetail>());

        Assert.False(result.Success);
    }

    // ── Phase 2: DWG-first behavior ──────────────────────────

    [Fact]
    public void RecommendedWriteStrategy_WhenCanWriteDwg_ReturnsDirectDwg()
    {
        _mockDwgFile.Setup(d => d.CanWriteDwg).Returns(true);
        Assert.Equal(WriteStrategy.DirectDwg, _service.RecommendedWriteStrategy);
    }

    [Fact]
    public async Task WriteTableIdsAsync_WhenCanWriteDwg_PrefersDwgExport()
    {
        _mockDwgFile.Setup(d => d.CanWriteDwg).Returns(true);
        _mockDwgFile.Setup(d => d.SaveAsDwgAsync(It.IsAny<string>()))
            .ReturnsAsync(true);

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };

        var result = await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        Assert.True(result.Success);
        Assert.Equal(WriteStrategy.DirectDwg, result.StrategyUsed);
        Assert.Equal(_tempDwgPath, result.OutputPath); // Writes back to original file

        // DWG was written to temp file — DXF should NOT be called
        _mockDwgFile.Verify(d => d.SaveAsDwgAsync(
            It.Is<string>(p => p.EndsWith(".tmp"))), Times.Once);
        _mockDwgFile.Verify(d => d.SaveAsDxfAsync(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public async Task WriteTableIdsAsync_WhenDwgFails_FallsBackToDxf()
    {
        _mockDwgFile.Setup(d => d.CanWriteDwg).Returns(true);
        _mockDwgFile.Setup(d => d.SaveAsDwgAsync(It.IsAny<string>()))
            .ReturnsAsync(false); // DWG export fails
        _mockDwgFile.Setup(d => d.SaveAsDxfAsync(It.IsAny<string>()))
            .ReturnsAsync(true); // DXF succeeds

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };

        var result = await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        Assert.True(result.Success);
        Assert.Equal(WriteStrategy.ExportDxf, result.StrategyUsed);
        Assert.Contains(".dxf", result.OutputPath);
    }

    [Fact]
    public async Task WriteTableIdsAsync_WhenBothFail_FallsBackToSidecar()
    {
        _mockDwgFile.Setup(d => d.CanWriteDwg).Returns(true);
        _mockDwgFile.Setup(d => d.SaveAsDwgAsync(It.IsAny<string>()))
            .ReturnsAsync(false);
        _mockDwgFile.Setup(d => d.SaveAsDxfAsync(It.IsAny<string>()))
            .ReturnsAsync(false);

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };

        var result = await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        Assert.True(result.Success); // Sidecar always succeeds
        Assert.Equal(WriteStrategy.SidecarJson, result.StrategyUsed);
    }

    [Fact]
    public async Task WriteTableIdsAsync_DwgSuccess_StillCreatesSidecarBackup()
    {
        _mockDwgFile.Setup(d => d.CanWriteDwg).Returns(true);
        _mockDwgFile.Setup(d => d.SaveAsDwgAsync(It.IsAny<string>()))
            .ReturnsAsync(true);

        var details = new List<WritebackDetail>
        {
            new() { ProductNo = "ST-001", DetailId = "D-001" }
        };

        await _service.WriteTableIdsAsync("Layout1", "H-123", details);

        // Even with successful DWG export, sidecar backup is created
        var sidecar = await _sidecarStore.LoadAsync(_tempDwgPath);
        Assert.NotNull(sidecar);
        Assert.Equal("H-123", sidecar!.Layouts[0].HeaderId);
    }

    [Fact]
    public async Task SaveAsync_WhenCanWriteDwg_PrefersDwgExport()
    {
        _mockDwgFile.Setup(d => d.CanWriteDwg).Returns(true);
        _mockDwgFile.Setup(d => d.SaveAsDwgAsync(It.IsAny<string>()))
            .ReturnsAsync(true);

        var result = await _service.SaveAsync();

        Assert.True(result);
        _mockDwgFile.Verify(d => d.SaveAsDwgAsync(
            It.Is<string>(p => p.EndsWith(".tmp"))), Times.Once);
        _mockDwgFile.Verify(d => d.SaveAsDxfAsync(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public async Task SaveAsync_WhenDwgFails_FallsBackToDxf()
    {
        _mockDwgFile.Setup(d => d.CanWriteDwg).Returns(true);
        _mockDwgFile.Setup(d => d.SaveAsDwgAsync(It.IsAny<string>()))
            .ReturnsAsync(false);
        _mockDwgFile.Setup(d => d.SaveAsDxfAsync(It.IsAny<string>()))
            .ReturnsAsync(true);

        var result = await _service.SaveAsync();

        Assert.True(result);
        _mockDwgFile.Verify(d => d.SaveAsDxfAsync(
            It.Is<string>(p => p.EndsWith(".dxf"))), Times.Once);
    }
}

public class FileDrawingDataServicePerLayoutTests : IDisposable
{
    private readonly string _tempDir;
    private readonly string _tempDwgPath;
    private readonly Mock<IDwgFileService> _mockDwgFile;
    private readonly SidecarIdStore _sidecarStore;
    private readonly FileDrawingDataService _service;

    public FileDrawingDataServicePerLayoutTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"fds_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _tempDwgPath = Path.Combine(_tempDir, "test.dwg");
        File.WriteAllText(_tempDwgPath, "dummy"); // sidecar needs file to exist at parent dir

        _mockDwgFile = new Mock<IDwgFileService>();
        _mockDwgFile.Setup(d => d.IsFileLoaded).Returns(true);
        _mockDwgFile.Setup(d => d.LoadedFilePath).Returns(_tempDwgPath);

        _sidecarStore = new SidecarIdStore();
        _service = new FileDrawingDataService(_mockDwgFile.Object, _sidecarStore);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, true); } catch { }
    }

    [Fact]
    public async Task SetAttributeValues_NullLayout_StoresGlobally()
    {
        var attrs = new Dictionary<string, string>
        {
            ["product_name"] = "Steel",
            ["spec"] = "3mm"
        };

        var result = await _service.SetAttributeValuesAsync(attrs, null);

        Assert.Equal(2, result.Count);

        // Verify sidecar has global attributes
        var sidecar = await _sidecarStore.LoadAsync(_tempDwgPath);
        Assert.NotNull(sidecar);
        Assert.Equal("Steel", sidecar!.Attributes["product_name"]);
        Assert.Equal("3mm", sidecar.Attributes["spec"]);
        Assert.Empty(sidecar.AttributesByLayout);
    }

    [Fact]
    public async Task SetAttributeValues_WithLayout_StoresPerLayout()
    {
        var attrs = new Dictionary<string, string>
        {
            ["product_name"] = "Aluminum",
            ["spec"] = "5mm"
        };

        var result = await _service.SetAttributeValuesAsync(attrs, "S405-201");

        Assert.Equal(2, result.Count);

        // Verify sidecar has per-layout attributes
        var sidecar = await _sidecarStore.LoadAsync(_tempDwgPath);
        Assert.NotNull(sidecar);
        Assert.Empty(sidecar!.Attributes); // global should be empty
        Assert.True(sidecar.AttributesByLayout.ContainsKey("S405-201"));
        Assert.Equal("Aluminum", sidecar.AttributesByLayout["S405-201"]["product_name"]);
    }

    [Fact]
    public async Task SetAttributeValues_MixedGlobalAndPerLayout_BothStored()
    {
        // First call: global
        await _service.SetAttributeValuesAsync(
            new Dictionary<string, string> { ["color_name"] = "Red" }, null);

        // Second call: per-layout
        await _service.SetAttributeValuesAsync(
            new Dictionary<string, string> { ["product_name"] = "Steel" }, "Layout1");

        var sidecar = await _sidecarStore.LoadAsync(_tempDwgPath);
        Assert.NotNull(sidecar);
        Assert.Equal("Red", sidecar!.Attributes["color_name"]);
        Assert.Equal("Steel", sidecar.AttributesByLayout["Layout1"]["product_name"]);
    }

    [Fact]
    public async Task SetAttributeValues_WithLayout_CallsApplyWithLayoutName()
    {
        var attrs = new Dictionary<string, string> { ["product_name"] = "Test" };

        await _service.SetAttributeValuesAsync(attrs, "Layout2");

        _mockDwgFile.Verify(d => d.ApplyAttributesToDocument(
            It.Is<Dictionary<string, string>>(a => a["product_name"] == "Test"),
            "Layout2"), Times.Once);
    }

    [Fact]
    public async Task SetAttributeValues_NullLayout_CallsApplyWithNullLayoutName()
    {
        var attrs = new Dictionary<string, string> { ["spec"] = "10mm" };

        await _service.SetAttributeValuesAsync(attrs, null);

        _mockDwgFile.Verify(d => d.ApplyAttributesToDocument(
            It.Is<Dictionary<string, string>>(a => a["spec"] == "10mm"),
            null), Times.Once);
    }
}
