// OdooAutoCAD.Core/AutoCAD/SidecarIdStore.cs
// JSON sidecar file for storing table IDs alongside DWG files.
// Avoids modifying the original DWG — safe for file-mode writeback.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Root model for the sidecar JSON file.
/// Stores table IDs and attribute overrides alongside a DWG file.
/// </summary>
public class SidecarData
{
    [JsonPropertyName("version")]
    public int Version { get; set; } = 1;

    [JsonPropertyName("source_dwg")]
    public string SourceDwg { get; set; } = string.Empty;

    [JsonPropertyName("modified_at")]
    public DateTime ModifiedAt { get; set; }

    [JsonPropertyName("layouts")]
    public List<SidecarLayout> Layouts { get; set; } = new();

    [JsonPropertyName("attributes")]
    public Dictionary<string, string> Attributes { get; set; } = new();

    [JsonPropertyName("attributes_by_layout")]
    public Dictionary<string, Dictionary<string, string>> AttributesByLayout { get; set; } = new();
}

/// <summary>
/// Per-layout table ID data in the sidecar.
/// </summary>
public class SidecarLayout
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("header_id")]
    public string HeaderId { get; set; } = string.Empty;

    [JsonPropertyName("details")]
    public List<SidecarDetail> Details { get; set; } = new();
}

/// <summary>
/// Per-row detail ID mapping in the sidecar.
/// </summary>
public class SidecarDetail
{
    [JsonPropertyName("product_no")]
    public string ProductNo { get; set; } = string.Empty;

    [JsonPropertyName("detail_id")]
    public string DetailId { get; set; } = string.Empty;
}

/// <summary>
/// Reads and writes sidecar JSON files alongside DWG drawings.
/// Thread-safe via SemaphoreSlim for concurrent access.
/// </summary>
public class SidecarIdStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly SemaphoreSlim _lock = new(1, 1);

    /// <summary>
    /// Computes the sidecar file path for a given DWG file.
    /// Example: "drawing.dwg" → "drawing.boq-ids.json"
    /// </summary>
    public static string GetSidecarPath(string dwgFilePath)
    {
        if (string.IsNullOrWhiteSpace(dwgFilePath))
            throw new ArgumentException("DWG file path cannot be empty.", nameof(dwgFilePath));

        var dir = Path.GetDirectoryName(dwgFilePath) ?? string.Empty;
        var nameWithoutExt = Path.GetFileNameWithoutExtension(dwgFilePath);
        return Path.Combine(dir, nameWithoutExt + ".boq-ids.json");
    }

    /// <summary>
    /// Loads sidecar data from disk. Returns null if the file does not exist.
    /// </summary>
    public async Task<SidecarData?> LoadAsync(string dwgFilePath)
    {
        var sidecarPath = GetSidecarPath(dwgFilePath);
        if (!File.Exists(sidecarPath))
            return null;

        await _lock.WaitAsync().ConfigureAwait(false);
        try
        {
            var json = await File.ReadAllTextAsync(sidecarPath).ConfigureAwait(false);
            return JsonSerializer.Deserialize<SidecarData>(json, JsonOptions);
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    /// Saves sidecar data to disk. Creates or overwrites the sidecar file.
    /// </summary>
    public async Task SaveAsync(string dwgFilePath, SidecarData data)
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data));

        var sidecarPath = GetSidecarPath(dwgFilePath);

        await _lock.WaitAsync().ConfigureAwait(false);
        try
        {
            var json = JsonSerializer.Serialize(data, JsonOptions);
            await File.WriteAllTextAsync(sidecarPath, json).ConfigureAwait(false);
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    /// Checks whether a sidecar file exists for the given DWG.
    /// </summary>
    public Task<bool> ExistsAsync(string dwgFilePath)
    {
        var sidecarPath = GetSidecarPath(dwgFilePath);
        return Task.FromResult(File.Exists(sidecarPath));
    }
}
