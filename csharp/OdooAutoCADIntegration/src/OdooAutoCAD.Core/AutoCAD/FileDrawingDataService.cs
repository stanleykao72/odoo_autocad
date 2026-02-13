// OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs
// File backend for IDrawingDataService — wraps IDwgFileService + SidecarIdStore.
// Reads DWG via ACadSharp, writes via sidecar JSON + DXF export.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OdooAutoCAD.Core.BOQ;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// IDrawingDataService implementation backed by file-based ACadSharp operations.
/// Reads directly from DWG via IDwgFileService.
/// Writes table IDs to sidecar JSON; writes attributes via DXF export.
/// </summary>
public class FileDrawingDataService : IDrawingDataService
{
    private readonly IDwgFileService _dwgFileService;
    private readonly SidecarIdStore _sidecarStore;

    public FileDrawingDataService(IDwgFileService dwgFileService, SidecarIdStore sidecarStore)
    {
        _dwgFileService = dwgFileService ?? throw new ArgumentNullException(nameof(dwgFileService));
        _sidecarStore = sidecarStore ?? throw new ArgumentNullException(nameof(sidecarStore));
    }

    public AutoCADOperationMode Mode => AutoCADOperationMode.File;
    public bool IsReady => _dwgFileService.IsFileLoaded;
    public string? CurrentSource => _dwgFileService.LoadedFilePath;
    public bool SupportsWrite => true;
    public WriteStrategy RecommendedWriteStrategy =>
        _dwgFileService.CanWriteDwg ? WriteStrategy.DirectDwg : WriteStrategy.ExportDxf;

    public event EventHandler<bool>? ReadyStateChanged;

    // ── Connect / Load ───────────────────────────────────────

    public async Task<bool> ConnectOrLoadAsync(string? filePath = null)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return false;

        var result = await _dwgFileService.LoadFileAsync(filePath);
        ReadyStateChanged?.Invoke(this, result);
        return result;
    }

    public async Task DisconnectOrUnloadAsync()
    {
        await _dwgFileService.UnloadFileAsync();
        ReadyStateChanged?.Invoke(this, false);
    }

    public Task<AutoCADStatus> GetStatusAsync()
    {
        if (!_dwgFileService.IsFileLoaded)
        {
            return Task.FromResult(new AutoCADStatus(
                false, null, null, null, null, "未載入 DWG 檔案"));
        }

        var fileName = Path.GetFileName(_dwgFileService.LoadedFilePath ?? string.Empty);
        return Task.FromResult(new AutoCADStatus(
            IsConnected: true,
            ApplicationName: "ACadSharp (File Mode)",
            Version: _dwgFileService.DwgVersion,
            CurrentDocument: fileName,
            OpenDocuments: new List<string> { fileName },
            ErrorMessage: null,
            DocumentPath: _dwgFileService.LoadedFilePath));
    }

    // ── READ ─────────────────────────────────────────────────

    public Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync()
    {
        if (!_dwgFileService.IsFileLoaded || _dwgFileService.LoadedFilePath == null)
            return Task.FromResult<IReadOnlyList<LayoutInfo>>(Array.Empty<LayoutInfo>());

        var layouts = _dwgFileService.GetLayouts(_dwgFileService.LoadedFilePath);
        return Task.FromResult<IReadOnlyList<LayoutInfo>>(layouts);
    }

    public async Task<LayoutData> ExtractParametersAsync(string layoutName)
    {
        if (!_dwgFileService.IsFileLoaded || _dwgFileService.LoadedFilePath == null)
            return new LayoutData { LayoutName = layoutName };

        // Get block attributes (parameters)
        var result = _dwgFileService.ExtractParameters(_dwgFileService.LoadedFilePath, layoutName);

        // Also extract 9-column table data (BOQ tables)
        var layoutTables = await _dwgFileService.ExtractTableDataAsync(layoutName);

        // Diagnostic: report entity types when no tables found OR tables have 0 data rows
        var totalDataRows = layoutTables.Sum(t => t.Rows.Count);
        if (layoutTables.Count == 0 || totalDataRows == 0)
        {
            var diag = _dwgFileService.GetLayoutEntityDiagnostics(layoutName);
            if (!string.IsNullOrEmpty(diag))
                result.Parameters["_diagnostic_entities"] = diag;
        }

        foreach (var lt in layoutTables)
        {
            // Copy header_id from table data into parameters (matches COM path behavior)
            if (!string.IsNullOrWhiteSpace(lt.HeaderId) && !result.Parameters.ContainsKey("header_id"))
            {
                result.Parameters["header_id"] = lt.HeaderId;
            }

            // Build header row from column keys
            var headerRow = new List<string>
                { "Position", "Product No", "Width", "Height", "Length", "Thickness", "Qty", "Description", "Detail ID" };

            var cells = new List<List<string>> { headerRow };
            foreach (var row in lt.Rows)
            {
                cells.Add(new List<string>
                {
                    row.GetValueOrDefault("position", ""),
                    row.GetValueOrDefault("product_no", ""),
                    row.GetValueOrDefault("width", ""),
                    row.GetValueOrDefault("height", ""),
                    row.GetValueOrDefault("len", ""),
                    row.GetValueOrDefault("thickness", ""),
                    row.GetValueOrDefault("qty", ""),
                    row.GetValueOrDefault("desc", ""),
                    row.GetValueOrDefault("detail_id", "")
                });
            }

            result.Tables.Add(new TableData
            {
                Name = lt.LayoutName,
                RowCount = cells.Count,
                ColumnCount = 9,
                Cells = cells
            });
        }

        return result;
    }

    public Task<IReadOnlyList<string>> GetHeaderIdsAsync()
    {
        return _dwgFileService.GetHeaderIdsAsync();
    }

    public Task<string> GetPRNumberAsync()
    {
        return _dwgFileService.GetPRNumberAsync();
    }

    public Task<Dictionary<string, string>> GetAttributeBlockAsync(string? layoutName = null)
    {
        if (!_dwgFileService.IsFileLoaded || _dwgFileService.LoadedFilePath == null)
            return Task.FromResult(new Dictionary<string, string>());

        // Use the first non-Model layout if not specified
        var targetLayout = layoutName;
        if (string.IsNullOrEmpty(targetLayout))
        {
            var layouts = _dwgFileService.GetLayouts(_dwgFileService.LoadedFilePath);
            targetLayout = layouts.FirstOrDefault()?.Name;
        }

        if (string.IsNullOrEmpty(targetLayout))
            return Task.FromResult(new Dictionary<string, string>());

        var data = _dwgFileService.ExtractParameters(_dwgFileService.LoadedFilePath, targetLayout);
        var attrs = new Dictionary<string, string>();
        foreach (var kvp in data.Parameters)
        {
            attrs[kvp.Key] = kvp.Value?.ToString() ?? string.Empty;
        }
        return Task.FromResult(attrs);
    }

    // ── WRITE ────────────────────────────────────────────────

    public async Task<WritebackResult> WriteTableIdsAsync(
        string layoutName, string headerId, IList<WritebackDetail> details)
    {
        if (!_dwgFileService.IsFileLoaded || _dwgFileService.LoadedFilePath == null)
            return new WritebackResult(false, "未載入 DWG 檔案", WriteStrategy.SidecarJson);

        // 1. Modify TABLE cells in memory
        _dwgFileService.WriteTableIdsToDocument(layoutName, headerId, details);

        // 2. Try DWG export first (native format), fallback to DXF
        string? outputPath = null;
        WriteStrategy usedStrategy = WriteStrategy.SidecarJson;

        if (_dwgFileService.CanWriteDwg)
        {
            var dwgPath = _dwgFileService.LoadedFilePath + ".modified.dwg";
            if (await _dwgFileService.SaveAsDwgAsync(dwgPath))
            {
                outputPath = dwgPath;
                usedStrategy = WriteStrategy.DirectDwg;
            }
        }

        if (outputPath == null)
        {
            var dxfPath = Path.ChangeExtension(_dwgFileService.LoadedFilePath, ".dxf");
            if (await _dwgFileService.SaveAsDxfAsync(dxfPath))
            {
                outputPath = dxfPath;
                usedStrategy = WriteStrategy.ExportDxf;
            }
        }

        // 3. Backup to sidecar JSON (always, as fallback)
        var sidecar = await _sidecarStore.LoadAsync(_dwgFileService.LoadedFilePath)
                      ?? new SidecarData();

        sidecar.SourceDwg = Path.GetFileName(_dwgFileService.LoadedFilePath);
        sidecar.ModifiedAt = DateTime.UtcNow;

        var layoutEntry = sidecar.Layouts.FirstOrDefault(l => l.Name == layoutName);
        if (layoutEntry == null)
        {
            layoutEntry = new SidecarLayout { Name = layoutName };
            sidecar.Layouts.Add(layoutEntry);
        }

        layoutEntry.HeaderId = headerId;
        layoutEntry.Details = details.Select(d => new SidecarDetail
        {
            ProductNo = d.ProductNo,
            DetailId = d.DetailId ?? string.Empty
        }).ToList();

        await _sidecarStore.SaveAsync(_dwgFileService.LoadedFilePath, sidecar);

        // 4. Return result based on export success
        if (outputPath != null)
        {
            var msg = usedStrategy == WriteStrategy.DirectDwg
                ? $"ID 已寫入 DWG: {outputPath}"
                : $"ID 已寫入 DXF: {outputPath}";
            return new WritebackResult(true, msg, usedStrategy, outputPath);
        }

        var sidecarPath = SidecarIdStore.GetSidecarPath(_dwgFileService.LoadedFilePath);
        return new WritebackResult(
            true,
            "ID 已儲存至 sidecar JSON (DWG/DXF 匯出失敗)",
            WriteStrategy.SidecarJson,
            sidecarPath);
    }

    public async Task<List<string>> SetAttributeValuesAsync(
        Dictionary<string, string> attributes, string? layoutName = null)
    {
        if (!_dwgFileService.IsFileLoaded || _dwgFileService.LoadedFilePath == null)
            return new List<string>();

        // Store in sidecar for persistence
        var sidecar = await _sidecarStore.LoadAsync(_dwgFileService.LoadedFilePath)
                      ?? new SidecarData();

        sidecar.SourceDwg = Path.GetFileName(_dwgFileService.LoadedFilePath);
        sidecar.ModifiedAt = DateTime.UtcNow;

        if (string.IsNullOrEmpty(layoutName))
        {
            // Global: apply to all layouts
            foreach (var (key, value) in attributes)
                sidecar.Attributes[key] = value;
        }
        else
        {
            // Per-layout: store under attributes_by_layout
            if (!sidecar.AttributesByLayout.TryGetValue(layoutName, out var layoutAttrs))
            {
                layoutAttrs = new Dictionary<string, string>();
                sidecar.AttributesByLayout[layoutName] = layoutAttrs;
            }
            foreach (var (key, value) in attributes)
                layoutAttrs[key] = value;
        }

        await _sidecarStore.SaveAsync(_dwgFileService.LoadedFilePath, sidecar);

        // Also apply to in-memory document for immediate DXF export capability
        _dwgFileService.ApplyAttributesToDocument(attributes, layoutName);

        return attributes.Keys.ToList();
    }

    public async Task<bool> SaveAsync()
    {
        if (!_dwgFileService.IsFileLoaded || _dwgFileService.LoadedFilePath == null)
            return false;

        // Apply attribute modifications from sidecar
        var sidecar = await _sidecarStore.LoadAsync(_dwgFileService.LoadedFilePath);
        if (sidecar != null)
        {
            // Apply global attributes to all layouts
            if (sidecar.Attributes?.Count > 0)
            {
                _dwgFileService.ApplyAttributesToDocument(sidecar.Attributes);
            }

            // Apply per-layout attributes
            foreach (var (layout, attrs) in sidecar.AttributesByLayout)
            {
                if (attrs.Count > 0)
                {
                    _dwgFileService.ApplyAttributesToDocument(attrs, layout);
                }
            }
        }

        // Try DWG export first (to separate file), fallback to DXF
        if (_dwgFileService.CanWriteDwg)
        {
            var dwgPath = _dwgFileService.LoadedFilePath + ".modified.dwg";
            if (await _dwgFileService.SaveAsDwgAsync(dwgPath))
                return true;
        }

        var dxfPath = Path.ChangeExtension(_dwgFileService.LoadedFilePath, ".dxf");
        return await _dwgFileService.SaveAsDxfAsync(dxfPath);
    }
}
