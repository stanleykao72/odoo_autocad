using System;
using System.IO;
using System.Threading.Tasks;
using OdooAutoCAD.Core.AutoCAD;
using Xunit;

namespace OdooAutoCAD.Integration.Tests.Services;

public class DwgFileServiceTests : IDisposable
{
    private readonly string _tempDir;
    private readonly DwgFileService _service;

    public DwgFileServiceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "DwgFileTests_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(_tempDir);
        _service = new DwgFileService();
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, true); }
        catch { /* cleanup best-effort */ }
    }

    [Fact]
    public void IsFileLoaded_Default_ReturnsFalse()
    {
        Assert.False(_service.IsFileLoaded);
    }

    [Fact]
    public void LoadedFilePath_Default_ReturnsNull()
    {
        Assert.Null(_service.LoadedFilePath);
    }

    [Fact]
    public void DwgVersion_Default_ReturnsNull()
    {
        Assert.Null(_service.DwgVersion);
    }

    [Fact]
    public void CanWriteDwg_ReturnsTrue()
    {
        Assert.True(_service.CanWriteDwg);
    }

    [Fact]
    public async Task LoadFile_EmptyPath_ReturnsFalse()
    {
        Assert.False(await _service.LoadFileAsync(""));
    }

    [Fact]
    public async Task LoadFile_NullPath_ReturnsFalse()
    {
        Assert.False(await _service.LoadFileAsync(null!));
    }

    [Fact]
    public async Task LoadFile_NonExistentFile_ReturnsFalse()
    {
        Assert.False(await _service.LoadFileAsync(Path.Combine(_tempDir, "nonexistent.dwg")));
        Assert.False(_service.IsFileLoaded);
    }

    [Fact]
    public async Task LoadFile_NonDwgExtension_ReturnsFalse()
    {
        var txtPath = Path.Combine(_tempDir, "file.txt");
        await File.WriteAllTextAsync(txtPath, "not a dwg");

        Assert.False(await _service.LoadFileAsync(txtPath));
    }

    [Fact]
    public async Task LoadFile_InvalidDwgContent_ReturnsFalse()
    {
        var fakeDwg = Path.Combine(_tempDir, "fake.dwg");
        await File.WriteAllTextAsync(fakeDwg, "this is not a valid DWG file");

        Assert.False(await _service.LoadFileAsync(fakeDwg));
        Assert.False(_service.IsFileLoaded);
        Assert.Null(_service.LoadedFilePath);
    }

    [Fact]
    public async Task UnloadFile_ClearsState()
    {
        await _service.UnloadFileAsync();

        Assert.False(_service.IsFileLoaded);
        Assert.Null(_service.LoadedFilePath);
        Assert.Null(_service.DwgVersion);
    }

    [Fact]
    public async Task ExtractTableData_WhenNotLoaded_ReturnsEmpty()
    {
        var result = await _service.ExtractTableDataAsync("Layout1");
        Assert.Empty(result);
    }

    [Fact]
    public async Task GetHeaderIds_WhenNotLoaded_ReturnsEmpty()
    {
        var result = await _service.GetHeaderIdsAsync();
        Assert.Empty(result);
    }

    [Fact]
    public async Task GetPRNumber_WhenNotLoaded_ReturnsEmpty()
    {
        var result = await _service.GetPRNumberAsync();
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public async Task SaveAsDxf_WhenNotLoaded_ReturnsFalse()
    {
        var outputPath = Path.Combine(_tempDir, "output.dxf");
        Assert.False(await _service.SaveAsDxfAsync(outputPath));
    }

    [Fact]
    public void ApplyAttributes_WhenNotLoaded_DoesNotThrow()
    {
        var attrs = new System.Collections.Generic.Dictionary<string, string>
        {
            ["product_name"] = "Test"
        };

        // Should not throw even when no document is loaded
        _service.ApplyAttributesToDocument(attrs);
    }

    [Fact]
    public void ApplyAttributes_EmptyDict_DoesNotThrow()
    {
        _service.ApplyAttributesToDocument(new System.Collections.Generic.Dictionary<string, string>());
    }

    [Fact]
    public void ApplyAttributes_NullDict_DoesNotThrow()
    {
        _service.ApplyAttributesToDocument(null!);
    }

    // IDwgReaderService inherited methods
    [Fact]
    public void IsValidDwgFile_InvalidPath_ReturnsFalse()
    {
        Assert.False(_service.IsValidDwgFile(""));
    }

    [Fact]
    public void IsValidDwgFile_NonExistent_ReturnsFalse()
    {
        Assert.False(_service.IsValidDwgFile(Path.Combine(_tempDir, "nope.dwg")));
    }

    [Fact]
    public void GetLayouts_EmptyPath_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => _service.GetLayouts(""));
    }

    [Fact]
    public void GetLayouts_NonExistentFile_ThrowsFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() => _service.GetLayouts(Path.Combine(_tempDir, "nope.dwg")));
    }
}
