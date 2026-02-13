// OdooAutoCAD.Core/AutoCAD/DrawingDataServiceDispatcher.cs
// Runtime mode-switching dispatcher that delegates to COM or File backend.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OdooAutoCAD.Core.BOQ;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Dispatches IDrawingDataService calls to the currently active backend (COM or File).
/// Supports runtime mode switching with proper disconnect/unload of the old backend.
/// </summary>
public class DrawingDataServiceDispatcher : IDrawingDataService
{
    private readonly ComDrawingDataService _comService;
    private readonly FileDrawingDataService _fileService;
    private IDrawingDataService _active;

    /// <summary>
    /// Raised when the active mode changes.
    /// </summary>
    public event EventHandler<AutoCADOperationMode>? ModeChanged;

    public event EventHandler<bool>? ReadyStateChanged;

    public DrawingDataServiceDispatcher(
        ComDrawingDataService comService,
        FileDrawingDataService fileService,
        AutoCADOperationMode initialMode = AutoCADOperationMode.COM)
    {
        _comService = comService ?? throw new ArgumentNullException(nameof(comService));
        _fileService = fileService ?? throw new ArgumentNullException(nameof(fileService));
        _active = initialMode == AutoCADOperationMode.COM ? comService : fileService;

        // Forward ReadyStateChanged from both backends
        _comService.ReadyStateChanged += (_, ready) =>
        {
            if (_active == _comService)
                ReadyStateChanged?.Invoke(this, ready);
        };
        _fileService.ReadyStateChanged += (_, ready) =>
        {
            if (_active == _fileService)
                ReadyStateChanged?.Invoke(this, ready);
        };
    }

    // ── Mode switching ───────────────────────────────────────

    public AutoCADOperationMode Mode => _active.Mode;

    /// <summary>
    /// Switches to a different operation mode.
    /// Disconnects/unloads the current backend before switching.
    /// </summary>
    public async Task SwitchModeAsync(AutoCADOperationMode newMode)
    {
        if (_active.Mode == newMode)
            return;

        // Disconnect/unload current backend
        if (_active.IsReady)
            await _active.DisconnectOrUnloadAsync();

        _active = newMode == AutoCADOperationMode.COM ? _comService : _fileService;
        ModeChanged?.Invoke(this, newMode);
        ReadyStateChanged?.Invoke(this, _active.IsReady);
    }

    // ── Delegated properties ─────────────────────────────────

    public bool IsReady => _active.IsReady;
    public string? CurrentSource => _active.CurrentSource;
    public bool SupportsWrite => _active.SupportsWrite;
    public WriteStrategy RecommendedWriteStrategy => _active.RecommendedWriteStrategy;

    // ── Delegated methods ────────────────────────────────────

    public Task<bool> ConnectOrLoadAsync(string? filePath = null)
        => _active.ConnectOrLoadAsync(filePath);

    public Task DisconnectOrUnloadAsync()
        => _active.DisconnectOrUnloadAsync();

    public Task<AutoCADStatus> GetStatusAsync()
        => _active.GetStatusAsync();

    public Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync()
        => _active.GetLayoutsAsync();

    public Task<LayoutData> ExtractParametersAsync(string layoutName)
        => _active.ExtractParametersAsync(layoutName);

    public Task<IReadOnlyList<string>> GetHeaderIdsAsync()
        => _active.GetHeaderIdsAsync();

    public Task<string> GetPRNumberAsync()
        => _active.GetPRNumberAsync();

    public Task<Dictionary<string, string>> GetAttributeBlockAsync(string? layoutName = null)
        => _active.GetAttributeBlockAsync(layoutName);

    public Task<WritebackResult> WriteTableIdsAsync(
        string layoutName, string headerId, IList<WritebackDetail> details)
        => _active.WriteTableIdsAsync(layoutName, headerId, details);

    public Task<List<string>> SetAttributeValuesAsync(
        Dictionary<string, string> attributes, string? layoutName = null)
        => _active.SetAttributeValuesAsync(attributes, layoutName);

    public Task<bool> SaveAsync()
        => _active.SaveAsync();
}
