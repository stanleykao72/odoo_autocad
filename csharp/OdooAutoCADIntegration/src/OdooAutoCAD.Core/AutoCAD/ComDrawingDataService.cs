// OdooAutoCAD.Core/AutoCAD/ComDrawingDataService.cs
// COM backend for IDrawingDataService — wraps existing IAutoCADService + IGUIProxy.
// All operations delegate to GUIProxy handlers running on the STA thread.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// IDrawingDataService implementation backed by COM automation via IGUIProxy.
/// Delegates all operations to existing AutoCADService handlers.
/// </summary>
public class ComDrawingDataService : IDrawingDataService
{
    private readonly IAutoCADService _autoCADService;
    private readonly IGUIProxy _guiProxy;

    public ComDrawingDataService(IAutoCADService autoCADService, IGUIProxy guiProxy)
    {
        _autoCADService = autoCADService ?? throw new ArgumentNullException(nameof(autoCADService));
        _guiProxy = guiProxy ?? throw new ArgumentNullException(nameof(guiProxy));
    }

    public AutoCADOperationMode Mode => AutoCADOperationMode.COM;
    public bool IsReady => _autoCADService.IsConnected;
    public string? CurrentSource => _autoCADService.GetCurrentDocumentPath();
    public bool SupportsWrite => true;
    public WriteStrategy RecommendedWriteStrategy => WriteStrategy.DirectDwg;

    public event EventHandler<bool>? ReadyStateChanged;

    // ── Connect / Disconnect ─────────────────────────────────

    public async Task<bool> ConnectOrLoadAsync(string? filePath = null)
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_connect", timeout: 15000);
        var connected = response.Success && response.Result is bool b && b;
        ReadyStateChanged?.Invoke(this, connected);
        return connected;
    }

    public async Task DisconnectOrUnloadAsync()
    {
        await _guiProxy.ExecuteInGuiAsync("autocad_disconnect");
        ReadyStateChanged?.Invoke(this, false);
    }

    public async Task<AutoCADStatus> GetStatusAsync()
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_status");
        if (response.Success && response.Result is AutoCADStatus status)
            return status;

        return new AutoCADStatus(false, null, null, null, null, response.ErrorMessage);
    }

    // ── READ ─────────────────────────────────────────────────

    public async Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync()
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_layouts");
        if (response.Success && response.Result is List<LayoutInfo> layouts)
            return layouts;

        return Array.Empty<LayoutInfo>();
    }

    public async Task<LayoutData> ExtractParametersAsync(string layoutName)
    {
        var parameters = new Dictionary<string, object?>
        {
            ["layoutName"] = layoutName
        };
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_extract_parameters", parameters);
        if (response.Success && response.Result is LayoutData data)
            return data;

        return new LayoutData { LayoutName = layoutName };
    }

    public async Task<IReadOnlyList<string>> GetHeaderIdsAsync()
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_header_ids");
        if (response.Success && response.Result is List<string> ids)
            return ids;

        return Array.Empty<string>();
    }

    public async Task<string> GetPRNumberAsync()
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_pr_number");
        if (response.Success && response.Result is string prNo)
            return prNo;

        return string.Empty;
    }

    public async Task<Dictionary<string, string>> GetAttributeBlockAsync(string? layoutName = null)
    {
        var parameters = layoutName != null
            ? new Dictionary<string, object?> { ["layoutName"] = layoutName }
            : null;

        var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_attribute_block", parameters);
        if (response.Success && response.Result is Dictionary<string, string> attrs)
            return attrs;

        return new Dictionary<string, string>();
    }

    // ── WRITE ────────────────────────────────────────────────

    public async Task<WritebackResult> WriteTableIdsAsync(
        string layoutName, string headerId, IList<WritebackDetail> details)
    {
        var parameters = new Dictionary<string, object?>
        {
            ["layout_name"] = layoutName,
            ["header_id"] = headerId,
            ["details"] = details
        };

        var response = await _guiProxy.ExecuteInGuiAsync("autocad_write_table_ids", parameters, timeout: 30000);
        if (response.Success)
            return new WritebackResult(true, "ID 已寫入 AutoCAD", WriteStrategy.DirectDwg);

        return new WritebackResult(false, response.ErrorMessage ?? "寫入失敗", WriteStrategy.DirectDwg);
    }

    public async Task<List<string>> SetAttributeValuesAsync(
        Dictionary<string, string> attributes, string? layoutName = null)
    {
        var parameters = new Dictionary<string, object?>
        {
            ["attributes"] = attributes
        };
        if (layoutName != null)
            parameters["layout_name"] = layoutName;

        var response = await _guiProxy.ExecuteInGuiAsync("autocad_set_attribute_values", parameters, timeout: 30000);
        if (response.Success && response.Result is List<string> results)
            return results;

        return attributes.Keys.ToList();
    }

    public Task<bool> SaveAsync()
    {
        // COM mode: changes are written immediately to AutoCAD, no separate save needed
        return Task.FromResult(true);
    }
}
