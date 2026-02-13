// OdooAutoCAD.Core/AutoCAD/IDwgFileService.cs
// Extended file-based DWG service — adds file management, table extraction, and write support.

using System.Collections.Generic;
using System.Threading.Tasks;
using OdooAutoCAD.Core.BOQ;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Represents extracted table data from a layout, including header ID.
/// </summary>
public class LayoutTableData
{
    public string LayoutName { get; set; } = string.Empty;
    public string HeaderId { get; set; } = string.Empty;
    public List<Dictionary<string, string>> Rows { get; set; } = new();
}

/// <summary>
/// Extended DWG file service — adds stateful file management, table extraction, and DXF export.
/// Extends <see cref="IDwgReaderService"/> (layout listing + block attribute extraction).
/// </summary>
public interface IDwgFileService : IDwgReaderService
{
    // ── File info ─────────────────────────────────────────────

    /// <summary>True when a DWG file is currently loaded in memory.</summary>
    bool IsFileLoaded { get; }

    /// <summary>Full path to the currently loaded DWG file.</summary>
    string? LoadedFilePath { get; }

    /// <summary>AutoCAD version string from the DWG header (e.g. "AC1032").</summary>
    string? DwgVersion { get; }

    /// <summary>True if DwgWriter is available for direct DWG output (currently false).</summary>
    bool CanWriteDwg { get; }

    // ── Load / Unload ─────────────────────────────────────────

    /// <summary>
    /// Loads a DWG file into memory. Unloads any previously loaded file first.
    /// </summary>
    Task<bool> LoadFileAsync(string filePath);

    /// <summary>
    /// Unloads the currently loaded DWG file and clears state.
    /// </summary>
    Task UnloadFileAsync();

    // ── Table operations ──────────────────────────────────────

    /// <summary>
    /// Extracts table data from a specific layout.
    /// Returns 9-column tables with header validation.
    /// </summary>
    Task<List<LayoutTableData>> ExtractTableDataAsync(string layoutName);

    /// <summary>
    /// Gets header IDs from cell(0,8) of valid 9-column tables across all layouts.
    /// </summary>
    Task<IReadOnlyList<string>> GetHeaderIdsAsync();

    /// <summary>
    /// Extracts the PR number from block attributes in the first layout.
    /// </summary>
    Task<string> GetPRNumberAsync();

    // ── Write operations ──────────────────────────────────────

    /// <summary>
    /// Exports the in-memory document as a DXF file.
    /// </summary>
    Task<bool> SaveAsDxfAsync(string outputPath);

    /// <summary>
    /// Exports the in-memory document as a DWG file.
    /// Returns false if DWG writing is not supported for this document.
    /// </summary>
    Task<bool> SaveAsDwgAsync(string outputPath);

    /// <summary>
    /// Modifies TABLE cell values in the loaded document for a specific layout.
    /// Writes header_id to cell(0,8) and detail_id to cell(row,8) matched by product_no in col 1.
    /// Does NOT write to disk — call <see cref="SaveAsDxfAsync"/> afterward.
    /// </summary>
    void WriteTableIdsToDocument(string layoutName, string headerId, IList<WritebackDetail> details);

    /// <summary>
    /// Applies attribute modifications to the in-memory CadDocument.
    /// If layoutName is null, applies to all non-Model layouts.
    /// If layoutName is specified, applies only to that layout.
    /// Does NOT write to disk — call <see cref="SaveAsDxfAsync"/> afterward.
    /// </summary>
    void ApplyAttributesToDocument(Dictionary<string, string> attributes, string? layoutName = null);

    /// <summary>
    /// Returns diagnostic info about entity types found in a layout.
    /// Used to troubleshoot when table extraction finds nothing.
    /// </summary>
    string GetLayoutEntityDiagnostics(string layoutName);
}
