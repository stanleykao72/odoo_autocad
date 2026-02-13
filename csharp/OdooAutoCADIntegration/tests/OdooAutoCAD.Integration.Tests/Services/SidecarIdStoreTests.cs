using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using OdooAutoCAD.Core.AutoCAD;
using Xunit;

namespace OdooAutoCAD.Integration.Tests.Services;

public class SidecarIdStoreTests : IDisposable
{
    private readonly string _tempDir;
    private readonly SidecarIdStore _store;

    public SidecarIdStoreTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "SidecarTests_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(_tempDir);
        _store = new SidecarIdStore();
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, true); }
        catch { /* cleanup best-effort */ }
    }

    private string DwgPath(string name = "test") => Path.Combine(_tempDir, $"{name}.dwg");

    [Fact]
    public void GetSidecarPath_ReturnsCorrectPath()
    {
        var path = SidecarIdStore.GetSidecarPath(@"C:\Projects\Drawing.dwg");
        Assert.Equal(@"C:\Projects\Drawing.boq-ids.json", path);
    }

    [Fact]
    public void GetSidecarPath_HandlesNestedDirectories()
    {
        var path = SidecarIdStore.GetSidecarPath(@"C:\Deep\Nested\Path\File.dwg");
        Assert.Equal(@"C:\Deep\Nested\Path\File.boq-ids.json", path);
    }

    [Fact]
    public void GetSidecarPath_ThrowsOnEmpty()
    {
        Assert.Throws<ArgumentException>(() => SidecarIdStore.GetSidecarPath(""));
    }

    [Fact]
    public async Task Load_NonExistentFile_ReturnsNull()
    {
        var result = await _store.LoadAsync(DwgPath("nonexistent"));
        Assert.Null(result);
    }

    [Fact]
    public async Task SaveAndLoad_Roundtrip_PreservesAllData()
    {
        var dwgPath = DwgPath();
        var data = new SidecarData
        {
            Version = 1,
            SourceDwg = "test.dwg",
            ModifiedAt = new DateTime(2026, 2, 13, 10, 30, 0, DateTimeKind.Utc),
            Layouts = new List<SidecarLayout>
            {
                new()
                {
                    Name = "S405-201",
                    HeaderId = "BOQ-001-H",
                    Details = new List<SidecarDetail>
                    {
                        new() { ProductNo = "H-200x200", DetailId = "BOQ-001-D-01" },
                        new() { ProductNo = "PL-12", DetailId = "BOQ-001-D-02" }
                    }
                }
            },
            Attributes = new Dictionary<string, string>
            {
                ["product_name"] = "Steel Plate",
                ["spec"] = "SS400"
            }
        };

        await _store.SaveAsync(dwgPath, data);
        var loaded = await _store.LoadAsync(dwgPath);

        Assert.NotNull(loaded);
        Assert.Equal(1, loaded!.Version);
        Assert.Equal("test.dwg", loaded.SourceDwg);
        Assert.Single(loaded.Layouts);
        Assert.Equal("S405-201", loaded.Layouts[0].Name);
        Assert.Equal("BOQ-001-H", loaded.Layouts[0].HeaderId);
        Assert.Equal(2, loaded.Layouts[0].Details.Count);
        Assert.Equal("H-200x200", loaded.Layouts[0].Details[0].ProductNo);
        Assert.Equal("BOQ-001-D-01", loaded.Layouts[0].Details[0].DetailId);
        Assert.Equal(2, loaded.Attributes.Count);
        Assert.Equal("Steel Plate", loaded.Attributes["product_name"]);
    }

    [Fact]
    public async Task Save_CreatesJsonFile()
    {
        var dwgPath = DwgPath();
        var data = new SidecarData { SourceDwg = "test.dwg", ModifiedAt = DateTime.UtcNow };

        await _store.SaveAsync(dwgPath, data);

        var sidecarPath = SidecarIdStore.GetSidecarPath(dwgPath);
        Assert.True(File.Exists(sidecarPath));

        var json = await File.ReadAllTextAsync(sidecarPath);
        Assert.Contains("source_dwg", json);
        Assert.Contains("test.dwg", json);
    }

    [Fact]
    public async Task Save_OverwritesExistingFile()
    {
        var dwgPath = DwgPath();

        await _store.SaveAsync(dwgPath, new SidecarData { SourceDwg = "old.dwg", ModifiedAt = DateTime.UtcNow });
        await _store.SaveAsync(dwgPath, new SidecarData { SourceDwg = "new.dwg", ModifiedAt = DateTime.UtcNow });

        var loaded = await _store.LoadAsync(dwgPath);
        Assert.Equal("new.dwg", loaded!.SourceDwg);
    }

    [Fact]
    public async Task Exists_ReturnsTrue_WhenFileExists()
    {
        var dwgPath = DwgPath();
        await _store.SaveAsync(dwgPath, new SidecarData { ModifiedAt = DateTime.UtcNow });

        Assert.True(await _store.ExistsAsync(dwgPath));
    }

    [Fact]
    public async Task Exists_ReturnsFalse_WhenNoFile()
    {
        Assert.False(await _store.ExistsAsync(DwgPath("missing")));
    }

    [Fact]
    public async Task SaveAndLoad_AttributesByLayout_Roundtrip()
    {
        var dwgPath = DwgPath();
        var data = new SidecarData
        {
            SourceDwg = "test.dwg",
            ModifiedAt = DateTime.UtcNow,
            Attributes = new Dictionary<string, string>
            {
                ["global_key"] = "global_value"
            },
            AttributesByLayout = new Dictionary<string, Dictionary<string, string>>
            {
                ["Layout1"] = new() { ["product_name"] = "Steel", ["spec"] = "3mm" },
                ["Layout2"] = new() { ["product_name"] = "Aluminum" }
            }
        };

        await _store.SaveAsync(dwgPath, data);
        var loaded = await _store.LoadAsync(dwgPath);

        Assert.NotNull(loaded);
        Assert.Equal("global_value", loaded!.Attributes["global_key"]);
        Assert.Equal(2, loaded.AttributesByLayout.Count);
        Assert.Equal("Steel", loaded.AttributesByLayout["Layout1"]["product_name"]);
        Assert.Equal("3mm", loaded.AttributesByLayout["Layout1"]["spec"]);
        Assert.Equal("Aluminum", loaded.AttributesByLayout["Layout2"]["product_name"]);
    }

    [Fact]
    public async Task ConcurrentSaves_DoNotCorruptFile()
    {
        var dwgPath = DwgPath();
        var tasks = new List<Task>();

        for (int i = 0; i < 10; i++)
        {
            int idx = i;
            tasks.Add(Task.Run(async () =>
            {
                var data = new SidecarData
                {
                    SourceDwg = $"concurrent_{idx}.dwg",
                    ModifiedAt = DateTime.UtcNow,
                    Layouts = new List<SidecarLayout>
                    {
                        new() { Name = $"Layout_{idx}", HeaderId = $"H-{idx}" }
                    }
                };
                await _store.SaveAsync(dwgPath, data);
            }));
        }

        await Task.WhenAll(tasks);

        // Verify file is valid JSON
        var loaded = await _store.LoadAsync(dwgPath);
        Assert.NotNull(loaded);
        Assert.NotEmpty(loaded!.SourceDwg);
        Assert.Single(loaded.Layouts);
    }
}
