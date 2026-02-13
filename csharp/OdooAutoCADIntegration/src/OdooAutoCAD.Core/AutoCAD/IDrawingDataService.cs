// OdooAutoCAD.Core/AutoCAD/IDrawingDataService.cs
// Unified interface for ViewModels — abstracts COM (live AutoCAD) vs File (ACadSharp) backends.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OdooAutoCAD.Core.BOQ;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Operation mode for AutoCAD data access.
/// </summary>
public enum AutoCADOperationMode
{
    /// <summary>COM automation — requires running AutoCAD instance.</summary>
    COM,
    /// <summary>File-based — reads DWG via ACadSharp, no AutoCAD required.</summary>
    File
}

/// <summary>
/// Strategy used for write-back operations.
/// </summary>
public enum WriteStrategy
{
    /// <summary>Write directly to AutoCAD drawing via COM.</summary>
    DirectDwg,
    /// <summary>Export modified document as DXF.</summary>
    ExportDxf,
    /// <summary>Store IDs in a sidecar JSON file alongside the DWG.</summary>
    SidecarJson,
    /// <summary>Write is not supported in this mode.</summary>
    WriteUnsupported
}

/// <summary>
/// Result of a write-back operation.
/// </summary>
public record WritebackResult(
    bool Success,
    string Message,
    WriteStrategy StrategyUsed,
    string? OutputPath = null);

/// <summary>
/// Unified interface for AutoCAD drawing data access.
/// ViewModels consume this interface regardless of whether the backend is COM or file-based.
/// </summary>
public interface IDrawingDataService
{
    /// <summary>Current operation mode (COM or File).</summary>
    AutoCADOperationMode Mode { get; }

    /// <summary>True when the backend is ready (COM: connected; File: file loaded).</summary>
    bool IsReady { get; }

    /// <summary>Current data source (COM: document name; File: file path).</summary>
    string? CurrentSource { get; }

    /// <summary>True if the backend supports write operations.</summary>
    bool SupportsWrite { get; }

    /// <summary>Recommended strategy for table ID write-back.</summary>
    WriteStrategy RecommendedWriteStrategy { get; }

    // ── Connect / Load ──────────────────────────────────────────

    /// <summary>
    /// COM: Connect to running AutoCAD. File: Load a DWG file.
    /// </summary>
    /// <param name="filePath">DWG file path (required for File mode, ignored for COM).</param>
    Task<bool> ConnectOrLoadAsync(string? filePath = null);

    /// <summary>
    /// COM: Disconnect from AutoCAD. File: Unload the DWG file.
    /// </summary>
    Task DisconnectOrUnloadAsync();

    /// <summary>
    /// Returns current status information.
    /// </summary>
    Task<AutoCADStatus> GetStatusAsync();

    // ── READ operations ─────────────────────────────────────────

    /// <summary>
    /// Gets all non-Model layouts from the current drawing.
    /// </summary>
    Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync();

    /// <summary>
    /// Extracts block attributes and table data from a specific layout.
    /// </summary>
    Task<LayoutData> ExtractParametersAsync(string layoutName);

    /// <summary>
    /// Gets header IDs from all layouts (cell(0,8) of 9-column tables).
    /// </summary>
    Task<IReadOnlyList<string>> GetHeaderIdsAsync();

    /// <summary>
    /// Gets the PR number from block attributes.
    /// </summary>
    Task<string> GetPRNumberAsync();

    /// <summary>
    /// Reads block attributes (project_name, product_name, etc.) from a layout.
    /// </summary>
    Task<Dictionary<string, string>> GetAttributeBlockAsync(string? layoutName = null);

    // ── WRITE operations ────────────────────────────────────────

    /// <summary>
    /// Writes header_id and detail_id values back to the drawing.
    /// COM: writes to AutoCAD table cells directly.
    /// File: stores in sidecar JSON.
    /// </summary>
    Task<WritebackResult> WriteTableIdsAsync(
        string layoutName,
        string headerId,
        IList<WritebackDetail> details);

    /// <summary>
    /// Sets block attribute values in the drawing.
    /// COM: writes to AutoCAD blocks directly.
    /// File: stores in sidecar, exports DXF on save.
    /// </summary>
    Task<List<string>> SetAttributeValuesAsync(
        Dictionary<string, string> attributes,
        string? layoutName = null);

    /// <summary>
    /// Saves pending changes.
    /// COM: no-op (changes are immediate). File: exports DXF.
    /// </summary>
    Task<bool> SaveAsync();

    // ── Events ──────────────────────────────────────────────────

    /// <summary>
    /// Raised when IsReady changes (connected/disconnected, file loaded/unloaded).
    /// </summary>
    event EventHandler<bool>? ReadyStateChanged;
}
