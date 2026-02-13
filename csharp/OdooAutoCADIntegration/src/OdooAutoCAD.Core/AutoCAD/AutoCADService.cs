// OdooAutoCAD.Core/AutoCAD/AutoCADService.cs
// AutoCAD COM Service Implementation - equivalent to Python util_autocad.py

using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// AutoCAD COM Service implementation.
/// Uses late binding (dynamic) for AutoCAD LT compatibility.
///
/// All COM operations run directly on the WPF STA thread via GUIProxy handlers.
/// The IDispatch QueryInterface warmup in GetActiveObject (lines 83-87) ensures
/// the COM proxy is fully initialized before any dynamic property access, preventing
/// deferred cross-process QI crashes. OleMessageFilter handles RPC_E_CALL_REJECTED
/// when AutoCAD is busy with WPF layout processing.
/// </summary>
public class AutoCADService : IAutoCADService, IDisposable
{
    private readonly ILogger<AutoCADService>? _logger;
    private readonly IGUIProxy _guiProxy;
    private dynamic? _acadApp;
    private dynamic? _acadDoc;
    private volatile bool _isConnected;

    private const string DefaultProgId = "AutoCAD.Application";
    private const int MaxRetryAttempts = 5;
    private const int RetryDelayMs = 1000;

    #region COM Interop for GetActiveObject (removed in .NET Core)

    // P/Invoke declarations matching Autodesk recommended pattern:
    // https://blog.autodesk.io/autocad-2025-marshalgetactiveobject-net-core/
    // https://chuongmep.com/posts/2024-05-02-use-com-api-autocad-netcore.html

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CLSIDFromProgIDEx(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid pclsid);

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid pclsid);

    /// <summary>
    /// Gets an active COM object by ProgID (replacement for Marshal.GetActiveObject in .NET Core).
    /// Uses CLSIDFromProgIDEx with fallback to CLSIDFromProgID for broader compatibility.
    ///
    /// IMPORTANT: Does NOT perform an explicit IDispatch QueryInterface here.
    /// Earlier versions forced a QI warmup immediately after GetActiveObject, but this
    /// triggered "Access Violation Reading 0x003f/0x0050" inside AutoCAD 2014's COM
    /// proxy DLL (53ac1e63h) — the proxy's vtable isn't ready for cross-apartment QI
    /// immediately after ROT lookup. Instead, ConnectInternalAsync uses stabilization
    /// delays before the first dynamic property access, letting .NET's DLR handle
    /// the IDispatch QI naturally when the proxy is ready.
    /// </summary>
    private static object GetActiveObject(string progId)
    {
        Guid clsid;
        try
        {
            CLSIDFromProgIDEx(progId, out clsid);
        }
        catch (Exception)
        {
            int hr = CLSIDFromProgID(progId, out clsid);
            if (hr < 0)
                Marshal.ThrowExceptionForHR(hr);
        }

        GetActiveObject(ref clsid, IntPtr.Zero, out object obj);

        return obj;
    }

    #endregion

    public AutoCADService(IGUIProxy guiProxy, ILogger<AutoCADService>? logger = null)
    {
        _guiProxy = guiProxy ?? throw new ArgumentNullException(nameof(guiProxy));
        _logger = logger;

        // Register GUI proxy handlers for AutoCAD operations
        RegisterGUIProxyHandlers();
    }

    /// <summary>
    /// Registers handlers for AutoCAD operations in the GUI proxy.
    /// All handlers execute COM operations directly on the STA/GUI thread.
    /// OleMessageFilter handles RPC_E_CALL_REJECTED retries automatically.
    /// </summary>
    private void RegisterGUIProxyHandlers()
    {
        _guiProxy.RegisterHandler("autocad_connect", async (parameters) =>
        {
            return await ConnectInternalAsync();
        });

        _guiProxy.RegisterHandler("autocad_disconnect", (parameters) =>
        {
            DisconnectInternal();
            return Task.FromResult<object?>(null);
        });

        _guiProxy.RegisterHandler("autocad_get_status", (parameters) =>
        {
            return Task.FromResult<object?>(GetStatusInternal());
        });

        _guiProxy.RegisterHandler("autocad_get_layouts", (parameters) =>
        {
            return Task.FromResult<object?>(GetLayoutsInternal());
        });

        _guiProxy.RegisterHandler("autocad_get_active_layout", (parameters) =>
        {
            return Task.FromResult<object?>(GetCurrentLayoutName());
        });

        _guiProxy.RegisterHandler("autocad_set_active_layout", (parameters) =>
        {
            parameters.TryGetValue("layoutName", out var val);
            var layoutName = val as string;
            if (layoutName != null)
            {
                return Task.FromResult<object?>(SwitchToLayout(layoutName));
            }
            return Task.FromResult<object?>(false);
        });

        _guiProxy.RegisterHandler("autocad_extract_parameters", (parameters) =>
        {
            parameters.TryGetValue("layoutName", out var val);
            var layoutName = val as string;
            if (layoutName != null)
            {
                return Task.FromResult<object?>(GetLayoutValuesInternal(layoutName));
            }
            return Task.FromResult<object?>(new LayoutData());
        });

        _guiProxy.RegisterHandler("autocad_open_document", (parameters) =>
        {
            parameters.TryGetValue("filePath", out var val);
            var filePath = val as string;
            if (string.IsNullOrEmpty(filePath) || !_isConnected)
                return Task.FromResult<object?>(false);

            try
            {
                _acadDoc = _acadApp!.Documents.Open(filePath);
                _logger?.LogInformation("Opened document: {FilePath}", filePath);
                return Task.FromResult<object?>(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to open document: {FilePath}", filePath);
                return Task.FromResult<object?>(false);
            }
        });

        _guiProxy.RegisterHandler("autocad_save_document", (parameters) =>
        {
            if (_acadDoc == null) return Task.FromResult<object?>(false);

            try
            {
                _acadDoc.Save();
                _logger?.LogInformation("Document saved");
                return Task.FromResult<object?>(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to save document");
                return Task.FromResult<object?>(false);
            }
        });

        _guiProxy.RegisterHandler("autocad_close_document", (parameters) =>
        {
            if (_acadDoc == null) return Task.FromResult<object?>(false);

            try
            {
                var save = true;
                if (parameters.TryGetValue("save", out var saveProp) && saveProp is bool s)
                    save = s;
                _acadDoc.Close(save);
                _acadDoc = _acadApp?.ActiveDocument;
                _logger?.LogInformation("Document closed");
                return Task.FromResult<object?>(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to close document");
                return Task.FromResult<object?>(false);
            }
        });

        _guiProxy.RegisterHandler("autocad_get_header_ids", (parameters) =>
        {
            if (_acadDoc == null)
                return Task.FromResult<object?>(new List<string>());

            var headerIds = new List<string>();

            try
            {
                foreach (dynamic layout in _acadDoc.Layouts)
                {
                    string layoutName = layout.Name;
                    if (layoutName == "Model") continue;

                    try
                    {
                        dynamic space = layout.Block;
                        foreach (dynamic entity in space)
                        {
                            try
                            {
                                string entityType = entity.EntityName;
                                if (entityType != "AcDbTable") continue;

                                int colCount = entity.Columns;
                                if (colCount != 9) continue;

                                string headerCell = entity.GetText(0, 7) ?? "";
                                if (!headerCell.Contains("HEADER_ID")) continue;

                                string headerId = LM_UnFormat(entity.GetText(0, 8) ?? "");
                                if (!string.IsNullOrWhiteSpace(headerId))
                                {
                                    headerIds.Add(headerId);
                                }

                                // Only process first valid table per layout
                                break;
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning(ex, "Error reading entity for header_id in {Layout}", layoutName);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning(ex, "Error accessing layout {Layout} for header_id extraction", layoutName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to extract header IDs from layouts");
            }

            return Task.FromResult<object?>(headerIds);
        });

        _guiProxy.RegisterHandler("autocad_write_table_ids", (parameters) =>
        {
            if (_acadDoc == null)
                return Task.FromResult<object?>("Error: No document open");

            var results = new List<string>();

            try
            {
                // Extract parameters
                parameters.TryGetValue("layout_name", out var lnVal);
                parameters.TryGetValue("header_id", out var hidVal);
                parameters.TryGetValue("details", out var detailsVal);

                var layoutName = lnVal as string;
                var headerId = hidVal?.ToString() ?? "";
                var details = detailsVal as IList<WritebackDetail> ?? new List<WritebackDetail>();

                if (string.IsNullOrEmpty(layoutName))
                    return Task.FromResult<object?>("Error: layout_name is required");

                // Switch to the target layout
                bool switched = SwitchToLayout(layoutName);
                if (!switched)
                    return Task.FromResult<object?>($"Error: Failed to switch to layout '{layoutName}'");

                dynamic layout = _acadDoc.Layouts.Item(layoutName);
                dynamic space = layout.Block;

                foreach (dynamic entity in space)
                {
                    try
                    {
                        string entityType = entity.EntityName;
                        if (entityType != "AcDbTable") continue;

                        int colCount = entity.Columns;
                        if (colCount != 9) continue;

                        // Validate: column 7 contains "HEADER_ID"
                        string headerCell = entity.GetText(0, 7) ?? "";
                        if (!headerCell.Contains("HEADER_ID")) continue;

                        // Write header_id to cell(0, 8)
                        try
                        {
                            entity.SetText(0, 8, headerId);
                            results.Add($"Header ID '{headerId}' written to cell(0,8)");
                        }
                        catch (Exception ex)
                        {
                            results.Add($"Failed to write header_id: {ex.Message}");
                        }

                        // Write detail_ids to matching rows (skip header rows 0-1, data starts at row 2)
                        int rowCount = entity.Rows;
                        for (int row = 2; row < rowCount; row++)
                        {
                            try
                            {
                                var cellProductNo = LM_UnFormat(entity.GetText(row, 1) ?? "");
                                if (string.IsNullOrWhiteSpace(cellProductNo)) continue;

                                // Find matching detail (case-insensitive)
                                var match = details.FirstOrDefault(d =>
                                    string.Equals(d.ProductNo, cellProductNo,
                                        StringComparison.OrdinalIgnoreCase));

                                if (match?.DetailId != null)
                                {
                                    entity.SetText(row, 8, match.DetailId);
                                    results.Add($"Row {row}: detail_id '{match.DetailId}' written for '{cellProductNo}'");
                                }
                            }
                            catch (Exception ex)
                            {
                                results.Add($"Row {row} writeback failed: {ex.Message}");
                            }
                        }

                        // Only process first valid table per layout
                        break;
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning(ex, "Error processing entity during writeback");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Writeback handler failed for layout");
                return Task.FromResult<object?>($"Error: {ex.Message}");
            }

            return Task.FromResult<object?>(results);
        });

        _guiProxy.RegisterHandler("autocad_get_pr_number", (parameters) =>
        {
            if (_acadDoc == null)
                return Task.FromResult<object?>(string.Empty);

            try
            {
                // Python logic: find block reference with Name == "pr_no",
                // then read the AcDbText inside the block definition.
                foreach (dynamic layout in _acadDoc.Layouts)
                {
                    string layoutName = layout.Name;
                    if (layoutName == "Model") continue;

                    try
                    {
                        dynamic space = layout.Block;
                        foreach (dynamic entity in space)
                        {
                            try
                            {
                                string entityType = entity.EntityName;
                                if (entityType != "AcDbBlockReference") continue;

                                string blockName = entity.Name;
                                if (!string.Equals(blockName, "pr_no", StringComparison.OrdinalIgnoreCase))
                                    continue;

                                // Found the pr_no block reference — read text from block definition
                                string effectiveName = entity.EffectiveName;
                                dynamic blocks = _acadDoc.Blocks;
                                dynamic blockDef = blocks.Item(effectiveName);

                                foreach (dynamic item in blockDef)
                                {
                                    try
                                    {
                                        string itemType = item.ObjectName;
                                        if (itemType == "AcDbText")
                                        {
                                            string text = LM_UnFormat(item.TextString);
                                            if (!string.IsNullOrWhiteSpace(text))
                                            {
                                                _logger?.LogInformation("PR number extracted: {PrNo}", text);
                                                return Task.FromResult<object?>(text);
                                            }
                                        }
                                    }
                                    catch { /* skip unreadable block items */ }
                                }

                                // Block found but no text — try attribute fallback
                                if ((bool)entity.HasAttributes)
                                {
                                    foreach (dynamic attr in entity.GetAttributes())
                                    {
                                        string attrText = LM_UnFormat(attr.TextString);
                                        if (!string.IsNullOrWhiteSpace(attrText))
                                        {
                                            _logger?.LogInformation("PR number from attribute: {PrNo}", attrText);
                                            return Task.FromResult<object?>(attrText);
                                        }
                                    }
                                }
                            }
                            catch { /* skip unreadable entities */ }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning(ex, "Error reading layout {Layout} for PR number", layoutName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to extract PR number");
            }

            return Task.FromResult<object?>(string.Empty);
        });

        _guiProxy.RegisterHandler("autocad_clear_table_ids", (parameters) =>
        {
            if (_acadDoc == null)
                return Task.FromResult<object?>(0);

            parameters.TryGetValue("layoutName", out var lnVal);
            var layoutName = lnVal as string;
            if (string.IsNullOrEmpty(layoutName))
                return Task.FromResult<object?>(0);

            int clearedCount = 0;
            try
            {
                bool switched = SwitchToLayout(layoutName);
                if (!switched)
                    return Task.FromResult<object?>(0);

                dynamic layout = _acadDoc.Layouts.Item(layoutName);
                dynamic space = layout.Block;

                foreach (dynamic entity in space)
                {
                    try
                    {
                        if (entity.EntityName != "AcDbTable") continue;
                        int colCount = entity.Columns;
                        if (colCount != 9) continue;

                        string headerCell = entity.GetText(0, 7) ?? "";
                        if (!headerCell.Contains("HEADER_ID")) continue;

                        int rowCount = entity.Rows;

                        // Clear header_id cell(0, 8)
                        try { entity.SetText(0, 8, ""); clearedCount++; } catch { }

                        // Clear detail_id cells (col 8) for data rows (skip header rows 0-1)
                        for (int row = 2; row < rowCount; row++)
                        {
                            try { entity.SetText(row, 8, ""); clearedCount++; } catch { }
                        }

                        break; // first valid table only
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to clear table IDs in layout {Layout}", layoutName);
            }

            return Task.FromResult<object?>(clearedCount);
        });

        _guiProxy.RegisterHandler("autocad_clear_all_table_ids", (parameters) =>
        {
            if (_acadDoc == null)
                return Task.FromResult<object?>(0);

            int totalCleared = 0;
            try
            {
                foreach (dynamic layout in _acadDoc.Layouts)
                {
                    string layoutName = layout.Name;
                    if (layoutName == "Model") continue;

                    try
                    {
                        dynamic space = layout.Block;
                        foreach (dynamic entity in space)
                        {
                            try
                            {
                                if (entity.EntityName != "AcDbTable") continue;
                                int colCount = entity.Columns;
                                if (colCount != 9) continue;

                                string headerCell = entity.GetText(0, 7) ?? "";
                                if (!headerCell.Contains("HEADER_ID")) continue;

                                int rowCount = entity.Rows;

                                // Clear header_id
                                try { entity.SetText(0, 8, ""); totalCleared++; } catch { }

                                // Clear detail_ids (skip header rows 0-1)
                                for (int row = 2; row < rowCount; row++)
                                {
                                    try { entity.SetText(row, 8, ""); totalCleared++; } catch { }
                                }

                                break;
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning(ex, "Error clearing IDs in layout {Layout}", layoutName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to clear all table IDs");
            }

            return Task.FromResult<object?>(totalCleared);
        });

        // Handler: Read current attribute values from the parameter block in a layout
        _guiProxy.RegisterHandler("autocad_get_attribute_block", (parameters) =>
        {
            if (_acadDoc == null)
                return Task.FromResult<object?>(new Dictionary<string, string>());

            var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Optional layout filter; if omitted, search all non-Model layouts
                parameters.TryGetValue("layoutName", out var lnVal);
                var targetLayout = lnVal as string;

                foreach (dynamic layout in _acadDoc.Layouts)
                {
                    string layoutName = layout.Name;
                    if (layoutName == "Model") continue;
                    if (targetLayout != null && !string.Equals(layoutName, targetLayout, StringComparison.OrdinalIgnoreCase))
                        continue;

                    try
                    {
                        dynamic space = layout.Block;
                        foreach (dynamic entity in space)
                        {
                            try
                            {
                                string entityType = entity.EntityName;
                                if (entityType != "AcDbBlockReference") continue;
                                if (!(bool)entity.HasAttributes) continue;

                                // Check if this block has project_name or job_working_plan_name attribute
                                bool isParamBlock = false;
                                foreach (dynamic attr in entity.GetAttributes())
                                {
                                    string tag = (attr.TagString ?? "").Trim();
                                    if (string.Equals(tag, "project_name", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(tag, "job_working_plan_name", StringComparison.OrdinalIgnoreCase))
                                    {
                                        isParamBlock = true;
                                        break;
                                    }
                                }

                                if (!isParamBlock) continue;

                                // Read all attributes from this block
                                foreach (dynamic attr in entity.GetAttributes())
                                {
                                    string tag = (attr.TagString ?? "").Trim();
                                    string value = LM_UnFormat(attr.TextString ?? "");
                                    if (!string.IsNullOrEmpty(tag))
                                    {
                                        attrs[tag] = value;
                                    }
                                }

                                _logger?.LogInformation("Read {Count} attributes from block in layout {Layout}",
                                    attrs.Count, layoutName);
                                return Task.FromResult<object?>(attrs);
                            }
                            catch { /* skip unreadable entities */ }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning(ex, "Error reading layout {Layout} for attribute block", layoutName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to get attribute block");
            }

            return Task.FromResult<object?>(attrs);
        });

        // Handler: Write attribute values to parameter block in a specific layout (or all non-Model layouts)
        _guiProxy.RegisterHandler("autocad_set_attribute_values", (parameters) =>
        {
            if (_acadDoc == null)
                return Task.FromResult<object?>(new List<string> { "Error: No document open" });

            var results = new List<string>();

            try
            {
                parameters.TryGetValue("attributes", out var attrsVal);
                var newValues = attrsVal as Dictionary<string, string>;
                if (newValues == null || newValues.Count == 0)
                    return Task.FromResult<object?>(new List<string> { "Error: No attributes provided" });

                // If layout_name is specified, only update that layout; otherwise update all
                parameters.TryGetValue("layout_name", out var layoutNameVal);
                var targetLayout = layoutNameVal as string;

                foreach (dynamic layout in _acadDoc.Layouts)
                {
                    string layoutName = layout.Name;
                    if (layoutName == "Model") continue;
                    if (!string.IsNullOrEmpty(targetLayout) &&
                        !string.Equals(layoutName, targetLayout, StringComparison.OrdinalIgnoreCase))
                        continue;

                    try
                    {
                        dynamic space = layout.Block;
                        foreach (dynamic entity in space)
                        {
                            try
                            {
                                string entityType = entity.EntityName;
                                if (entityType != "AcDbBlockReference") continue;
                                if (!(bool)entity.HasAttributes) continue;

                                // Check if this is the parameter block
                                bool isParamBlock = false;
                                foreach (dynamic attr in entity.GetAttributes())
                                {
                                    string tag = (attr.TagString ?? "").Trim();
                                    if (string.Equals(tag, "project_name", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(tag, "job_working_plan_name", StringComparison.OrdinalIgnoreCase))
                                    {
                                        isParamBlock = true;
                                        break;
                                    }
                                }

                                if (!isParamBlock) continue;

                                // Write matching attributes
                                int written = 0;
                                foreach (dynamic attr in entity.GetAttributes())
                                {
                                    string tag = (attr.TagString ?? "").Trim();
                                    foreach (var kv in newValues)
                                    {
                                        if (string.Equals(tag, kv.Key, StringComparison.OrdinalIgnoreCase))
                                        {
                                            try
                                            {
                                                attr.TextString = kv.Value;
                                                written++;
                                            }
                                            catch (Exception ex)
                                            {
                                                results.Add($"[{layoutName}] Failed to write '{kv.Key}': {ex.Message}");
                                            }
                                            break;
                                        }
                                    }
                                }

                                results.Add($"[{layoutName}] Wrote {written}/{newValues.Count} attributes");
                                break; // first param block per layout
                            }
                            catch { /* skip unreadable entities */ }
                        }
                    }
                    catch (Exception ex)
                    {
                        results.Add($"[{layoutName}] Error: {ex.Message}");
                        _logger?.LogWarning(ex, "Error writing attributes in layout {Layout}", layoutName);
                    }
                }

                _logger?.LogInformation("Set attribute values across layouts: {Results}", string.Join("; ", results));
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to set attribute values");
                results.Add($"Error: {ex.Message}");
            }

            return Task.FromResult<object?>(results);
        });

        _logger?.LogDebug("AutoCAD GUI proxy handlers registered");
    }

    // Only use the boolean flag — touching _acadApp from a non-STA thread triggers COM marshaling
    // and can cause access violations inside AutoCAD.
    public bool IsConnected => _isConnected;

    #region Connection Management

    public async Task<bool> ConnectAsync()
    {
        // IMPORTANT: Connects directly on the caller's STA thread, NOT through GUIProxy.
        // Python's pywin32 calls GetActiveObject on a regular thread with CoInitialize(),
        // and it works fine. Our earlier approach routed through GUIProxy's DispatcherTimer,
        // which caused "Access Violation Reading 0x003f/0x0050" in AutoCAD 2014's COM server.
        // The DispatcherTimer.Tick callback context interferes with cross-process COM
        // message pumping — AutoCAD's COM server can't properly negotiate OLE messages
        // when the caller is inside a WPF timer callback. Direct async call avoids this.
        return await ConnectInternalAsync();
    }

    /// <summary>
    /// Internal connection method — runs on the STA/GUI thread.
    /// Uses async Task.Delay for retries so the WPF message pump stays responsive.
    /// Matches the Python util_autocad.py connect_autocad() logic:
    ///   1. GetActiveObject to get running AutoCAD instance
    ///   2. Retry up to 5 times for ActiveDocument
    /// </summary>
    /// <summary>
    /// Connects to a running AutoCAD instance via COM with full retry logic.
    ///
    /// MUST be called directly on the STA thread (from ViewModel async commands),
    /// NOT through GUIProxy/DispatcherTimer. AutoCAD 2014's COM server crashes
    /// with "Access Violation Reading 0x003f/0x0050" (at 53ac1e63h) when the initial
    /// IDispatch QI happens inside a WPF DispatcherTimer.Tick callback — the timer
    /// context interferes with cross-process OLE message pumping. Direct async calls
    /// on the STA thread pump messages correctly, matching Python's pywin32 behavior.
    ///
    /// Retries up to 3 times with increasing delays, releasing COM objects between
    /// attempts to avoid stale proxy state.
    /// </summary>
    private async Task<bool> ConnectInternalAsync()
    {
        const int maxConnectRetries = 3;

        for (int connectAttempt = 1; connectAttempt <= maxConnectRetries; connectAttempt++)
        {
            try
            {
                _acadApp = GetActiveObject(DefaultProgId);
                _logger?.LogInformation("GetActiveObject succeeded (attempt {Attempt}/{Max})",
                    connectAttempt, maxConnectRetries);

                // Brief stabilization — let the COM proxy initialize.
                // First attempt uses a short delay; retries use longer delays
                // in case AutoCAD is recovering from a previous failed connection.
                var stabilizationMs = connectAttempt == 1 ? 200 : 500 * connectAttempt;
                await Task.Delay(stabilizationMs);

                // Retry for ActiveDocument (matches Python's retry pattern)
                for (int attempt = 1; attempt <= MaxRetryAttempts; attempt++)
                {
                    try
                    {
                        _acadDoc = _acadApp.ActiveDocument;
                        if (_acadDoc != null)
                        {
                            // Brief pause before first property access on document —
                            // the document proxy also needs time after first IDispatch QI.
                            await Task.Delay(100);

                            string docName = _acadDoc.Name;
                            _logger?.LogInformation("ActiveDocument: {Name}", docName);
                            break;
                        }
                    }
                    catch (Exception ex) when (attempt < MaxRetryAttempts)
                    {
                        _logger?.LogWarning("Retry {Attempt}/{Max}: ActiveDocument not ready — {Error}",
                            attempt, MaxRetryAttempts, ex.Message);
                        await Task.Delay(RetryDelayMs);
                    }
                }

                _isConnected = true;
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "COM connect attempt {Attempt}/{Max} failed",
                    connectAttempt, maxConnectRetries);

                // Release COM objects before retry — stale proxies can cause AVs
                ReleaseCOMObjects();

                if (connectAttempt < maxConnectRetries)
                {
                    // Increasing delay: 1s, 2s — gives AutoCAD time to recover
                    // after showing its "連線中斷" error dialog
                    var retryDelay = 1000 * connectAttempt;
                    _logger?.LogInformation("Waiting {Delay}ms before retry...", retryDelay);
                    await Task.Delay(retryDelay);
                }
            }
        }

        _logger?.LogError("Failed to connect to AutoCAD after {MaxRetries} attempts", maxConnectRetries);
        _isConnected = false;
        return false;
    }

    public async Task DisconnectAsync()
    {
        await _guiProxy.ExecuteInGuiAsync("autocad_disconnect", null, timeout: 5000);
    }

    /// <summary>
    /// Internal disconnection method executed on GUI thread.
    /// Properly releases COM objects to avoid leaks.
    /// </summary>
    private void DisconnectInternal()
    {
        _isConnected = false;
        ReleaseCOMObjects();
        _logger?.LogInformation("Disconnected from AutoCAD");
    }

    /// <summary>
    /// Safely releases COM object references.
    /// </summary>
    private void ReleaseCOMObjects()
    {
        if (_acadDoc != null)
        {
            try
            {
                if (Marshal.IsComObject(_acadDoc))
                    Marshal.ReleaseComObject(_acadDoc);
            }
            catch { /* ignore release errors */ }
            _acadDoc = null;
        }
        if (_acadApp != null)
        {
            try
            {
                if (Marshal.IsComObject(_acadApp))
                    Marshal.ReleaseComObject(_acadApp);
            }
            catch { /* ignore release errors */ }
            _acadApp = null;
        }
    }

    public async Task<AutoCADStatus> GetStatusAsync()
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_status", null, timeout: 5000);

        if (response.Success && response.Result is AutoCADStatus status)
        {
            return status;
        }

        return new AutoCADStatus(
            IsConnected: false,
            ApplicationName: null,
            Version: null,
            CurrentDocument: null,
            OpenDocuments: null,
            ErrorMessage: response.ErrorMessage ?? "Failed to get AutoCAD status");
    }

    /// <summary>
    /// Internal status retrieval method executed on GUI thread.
    /// </summary>
    private AutoCADStatus GetStatusInternal()
    {
        if (!_isConnected || _acadApp == null)
        {
            return new AutoCADStatus(
                IsConnected: false,
                ApplicationName: null,
                Version: null,
                CurrentDocument: null,
                OpenDocuments: null,
                ErrorMessage: "Not connected to AutoCAD");
        }

        try
        {
            string appName = _acadApp!.Name;
            string version = _acadApp.Version;
            string? currentDoc = _acadDoc?.Name;
            string? documentPath = null;
            try { documentPath = _acadDoc?.FullName; } catch { /* may fail if no doc */ }

            var openDocs = new List<string>();
            foreach (dynamic doc in _acadApp.Documents)
            {
                openDocs.Add(doc.Name);
            }

            return new AutoCADStatus(
                IsConnected: true,
                ApplicationName: appName,
                Version: version,
                CurrentDocument: currentDoc,
                OpenDocuments: openDocs,
                ErrorMessage: null,
                DocumentPath: documentPath);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to get AutoCAD status");
            return new AutoCADStatus(
                IsConnected: false,
                ApplicationName: null,
                Version: null,
                CurrentDocument: null,
                OpenDocuments: null,
                ErrorMessage: ex.Message);
        }
    }

    #endregion

    #region Document Management

    public async Task<bool> OpenDocumentAsync(string filePath)
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_open_document",
            new Dictionary<string, object?> { ["filePath"] = filePath }, timeout: 10000);
        return response.Success && response.Result is bool opened && opened;
    }

    public async Task<bool> SaveDocumentAsync()
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_save_document", null, timeout: 10000);
        return response.Success && response.Result is bool saved && saved;
    }

    public async Task<bool> CloseDocumentAsync(bool save = true)
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_close_document",
            new Dictionary<string, object?> { ["save"] = save }, timeout: 10000);
        return response.Success && response.Result is bool closed && closed;
    }

    public string? GetCurrentDocumentPath()
    {
        try
        {
            return _acadDoc?.FullName;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region Layout Operations

    public IReadOnlyList<LayoutInfo> GetLayouts()
    {
        return GetLayoutsInternal();
    }

    /// <summary>
    /// Internal method to get layouts. Excludes "Model" layout per requirements.
    /// </summary>
    private List<LayoutInfo> GetLayoutsInternal()
    {
        var layouts = new List<LayoutInfo>();

        try
        {
            if (_acadDoc == null) return layouts;

            foreach (dynamic layout in _acadDoc.Layouts)
            {
                string layoutName = layout.Name;
                
                // Exclude Model layout
                if (layoutName == "Model")
                    continue;

                layouts.Add(new LayoutInfo(
                    Name: layoutName,
                    TabOrder: layout.TabOrder,
                    IsModelSpace: false,
                    PlotConfigurationName: layout.ConfigName ?? ""));
            }

            return layouts.OrderBy(l => l.TabOrder).ToList();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to get layouts");
            return layouts;
        }
    }

    public bool SwitchToLayout(string layoutName)
    {
        try
        {
            if (_acadDoc == null) return false;

            _acadDoc.ActiveLayout = _acadDoc.Layouts.Item(layoutName);
            _logger?.LogDebug("Switched to layout: {LayoutName}", layoutName);
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to switch to layout: {LayoutName}", layoutName);
            return false;
        }
    }

    public string? GetCurrentLayoutName()
    {
        try
        {
            return _acadDoc?.ActiveLayout?.Name;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region Drawing Operations

    public string DrawLine(Point3D start, Point3D end, string layer = "0")
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var startPoint = new double[] { start.X, start.Y, start.Z };
            var endPoint = new double[] { end.X, end.Y, end.Z };

            var modelSpace = _acadDoc.ModelSpace;
            var line = modelSpace.AddLine(startPoint, endPoint);
            line.Layer = layer;

            return line.Handle;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to draw line");
            throw;
        }
    }

    public string DrawCircle(Point3D center, double radius, string layer = "0")
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var centerPoint = new double[] { center.X, center.Y, center.Z };

            var modelSpace = _acadDoc.ModelSpace;
            var circle = modelSpace.AddCircle(centerPoint, radius);
            circle.Layer = layer;

            return circle.Handle;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to draw circle");
            throw;
        }
    }

    public string DrawArc(Point3D center, double radius, double startAngle, double endAngle, string layer = "0")
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var centerPoint = new double[] { center.X, center.Y, center.Z };
            var startAngleRad = startAngle * Math.PI / 180;
            var endAngleRad = endAngle * Math.PI / 180;

            var modelSpace = _acadDoc.ModelSpace;
            var arc = modelSpace.AddArc(centerPoint, radius, startAngleRad, endAngleRad);
            arc.Layer = layer;

            return arc.Handle;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to draw arc");
            throw;
        }
    }

    public string DrawPolyline(IEnumerable<Point3D> points, bool closed = false, string layer = "0")
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var pointList = points.ToList();
            var vertices = new double[pointList.Count * 2];

            for (int i = 0; i < pointList.Count; i++)
            {
                vertices[i * 2] = pointList[i].X;
                vertices[i * 2 + 1] = pointList[i].Y;
            }

            var modelSpace = _acadDoc.ModelSpace;
            var pline = modelSpace.AddLightWeightPolyline(vertices);
            pline.Layer = layer;
            pline.Closed = closed;

            return pline.Handle;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to draw polyline");
            throw;
        }
    }

    public string DrawRectangle(Point3D corner1, Point3D corner2, string layer = "0")
    {
        var points = new List<Point3D>
        {
            corner1,
            new Point3D(corner2.X, corner1.Y),
            corner2,
            new Point3D(corner1.X, corner2.Y)
        };

        return DrawPolyline(points, closed: true, layer: layer);
    }

    #endregion

    #region Text Operations

    public string CreateText(
        Point3D position,
        string content,
        double height,
        double rotation = 0,
        TextAlignment alignment = TextAlignment.Left,
        string layer = "0")
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var insertPoint = new double[] { position.X, position.Y, position.Z };
            var rotationRad = rotation * Math.PI / 180;

            var modelSpace = _acadDoc.ModelSpace;
            var text = modelSpace.AddText(content, insertPoint, height);
            text.Layer = layer;
            text.Rotation = rotationRad;

            return text.Handle;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to create text");
            throw;
        }
    }

    public string CreateMText(
        Point3D position,
        string content,
        double width,
        double height,
        string layer = "0")
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var insertPoint = new double[] { position.X, position.Y, position.Z };

            var modelSpace = _acadDoc.ModelSpace;
            var mtext = modelSpace.AddMText(insertPoint, width, content);
            mtext.Layer = layer;
            mtext.Height = height;

            return mtext.Handle;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to create MText");
            throw;
        }
    }

    #endregion

    #region Dimension Operations

    public string AddDimension(
        DimensionType type,
        IEnumerable<Point3D> points,
        Point3D dimLineLocation,
        string layer = "0")
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var pointList = points.ToList();
            var modelSpace = _acadDoc.ModelSpace;
            dynamic dim;

            switch (type)
            {
                case DimensionType.Linear:
                    var p1 = new double[] { pointList[0].X, pointList[0].Y, pointList[0].Z };
                    var p2 = new double[] { pointList[1].X, pointList[1].Y, pointList[1].Z };
                    var dimLoc = new double[] { dimLineLocation.X, dimLineLocation.Y, dimLineLocation.Z };
                    dim = modelSpace.AddDimAligned(p1, p2, dimLoc);
                    break;

                case DimensionType.Radial:
                    var center = new double[] { pointList[0].X, pointList[0].Y, pointList[0].Z };
                    var chordPt = new double[] { pointList[1].X, pointList[1].Y, pointList[1].Z };
                    dim = modelSpace.AddDimRadial(center, chordPt, 0);
                    break;

                default:
                    throw new NotSupportedException($"Dimension type {type} not yet implemented");
            }

            dim.Layer = layer;
            return dim.Handle;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to add dimension");
            throw;
        }
    }

    #endregion

    #region Parameter Extraction

    /// <summary>
    /// Strips AutoCAD MText formatting codes (matches Python LM_UnFormat).
    /// Handles: \P (paragraph), \f/\F (font), \C/\c (color), \H (height),
    /// \A (alignment), \L/\l (underline), \O/\o (overline), \Q (oblique),
    /// \T (tracking), \W (width), \S (stacking), \p (paragraph style), {} braces.
    /// </summary>
    private static string LM_UnFormat(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var result = text;

        // Step 1: Preserve escaped backslashes (matches Python: r"\\\\" -> "\032")
        result = result.Replace("\\\\", "\x1A");

        // Step 2: Replace \P, newlines, tabs with space
        result = Regex.Replace(result, @"\\P|\n|\t", " ");

        // Step 3: Remove parameterized formatting codes ending with semicolon
        // Covers \A, \C, \c, \F, \f, \H, \L, \l, \O, \o, \p, \Q, \T, \W
        // e.g. \fPMingLiU|b0|i0|c136|p2; or \C1; or \H1.5x;
        result = Regex.Replace(result, @"\\[ACcFfHLlOopQTW][^\\;]*;", "");

        // Step 4: Remove toggle codes without semicolon (e.g. \L, \l, \O, \o)
        result = Regex.Replace(result, @"\\[ACcFfHLlOopQTW]", "");

        // Step 5: Remove stacking \S patterns
        result = Regex.Replace(result, @"\\S[^;]*;", "");

        // Step 6: Remove \~ (non-breaking space) → space
        result = result.Replace("\\~", " ");

        // Step 7: Remove braces (group delimiters)
        result = Regex.Replace(result, @"[{}]", "");

        // Step 8: Restore preserved backslashes
        result = result.Replace("\x1A", "\\");

        return result.Trim();
    }

    /// <summary>
    /// Extracts block attributes from the layout.
    /// Searches for standard attributes: pr_no, project_name, job_working_plan_name, 
    /// product_name, spec, color_name, unit, remarks, block_name, quantity
    /// </summary>
    private Dictionary<string, object> GetAttributeValues(dynamic layout)
    {
        var attributes = new Dictionary<string, object>();

        try
        {
            if (_acadDoc == null) return attributes;

            // Get the PaperSpace or ModelSpace depending on layout
            dynamic space = layout.Block;

            // Iterate through all entities in the layout
            foreach (dynamic entity in space)
            {
                try
                {
                    string entityType = entity.EntityName;

                    // Check if it's a block reference
                    if (entityType == "AcDbBlockReference")
                    {
                        string blockName = entity.Name;

                        // Special case: "pr_no" block stores value as AcDbText inside block def,
                        // not as an attribute (matches Python get_pr_no() pattern)
                        if (string.Equals(blockName, "pr_no", StringComparison.OrdinalIgnoreCase)
                            && !attributes.ContainsKey("pr_no"))
                        {
                            try
                            {
                                string effectiveName = entity.EffectiveName;
                                dynamic blocks = _acadDoc.Blocks;
                                dynamic blockDef = blocks.Item(effectiveName);
                                foreach (dynamic item in blockDef)
                                {
                                    try
                                    {
                                        string itemType = item.ObjectName;
                                        if (itemType == "AcDbText")
                                        {
                                            string text = LM_UnFormat(item.TextString);
                                            if (!string.IsNullOrWhiteSpace(text))
                                            {
                                                attributes["pr_no"] = text;
                                                break;
                                            }
                                        }
                                    }
                                    catch { /* skip unreadable block items */ }
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning(ex, "Failed to read pr_no block text");
                            }
                        }

                        // Read standard block attributes
                        if ((bool)entity.HasAttributes)
                        {
                            foreach (dynamic attr in entity.GetAttributes())
                            {
                                string tag = attr.TagString;
                                string value = LM_UnFormat(attr.TextString);

                                // Map standard attribute tags (store lowercase to match Python tag_list)
                                switch (tag.ToLower())
                                {
                                    case "pr_no":
                                    case "project_name":
                                    case "job_working_plan_name":
                                    case "product_name":
                                    case "product_catelog":
                                    case "spec":
                                    case "surface_treatment":
                                    case "operation_flow":
                                    case "color_name":
                                    case "color_no":
                                    case "unit":
                                    case "remarks":
                                    case "block_name":
                                    case "quantity":
                                        attributes[tag.ToLower()] = value;
                                        break;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // Skip entities that can't be read
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to extract attribute values");
        }

        return attributes;
    }

    /// <summary>
    /// Extracts table data from the layout.
    /// Validates: exactly 9 columns, HEADER_ID in column 7 (0-indexed).
    /// Rows 0-1 are title/header; data starts at row 2 (matching Python's "if i > 1").
    /// Returns rows with: Position, Product No, Width, Height, Length, Thickness, Qty, Description, Detail ID.
    /// Filters empty rows (both qty AND product_no empty).
    /// </summary>
    private (string? HeaderId, List<Dictionary<string, object>> Rows) GetTableData(dynamic layout)
    {
        var tableRows = new List<Dictionary<string, object>>();
        string? headerId = null;

        try
        {
            if (_acadDoc == null) return (null, tableRows);

            dynamic space = layout.Block;

            // Find table entities
            foreach (dynamic entity in space)
            {
                try
                {
                    string entityType = entity.EntityName;

                    if (entityType == "AcDbTable")
                    {
                        int rowCount = entity.Rows;
                        int colCount = entity.Columns;

                        // Validate: exactly 9 columns
                        if (colCount != 9)
                        {
                            _logger?.LogWarning("Table has {ColCount} columns, expected 9. Skipping.", colCount);
                            continue;
                        }

                        // Validate: column 7 (0-indexed) contains "HEADER_ID" in header
                        string headerCell = entity.GetText(0, 7) ?? "";
                        if (!headerCell.Contains("HEADER_ID"))
                        {
                            _logger?.LogWarning("Column 7 does not contain HEADER_ID. Found: '{Header}'", headerCell);
                            continue;
                        }

                        // Extract header_id from row 0, column 8 (matches Python get_table_data)
                        headerId = LM_UnFormat(entity.GetText(0, 8) ?? "");

                        // Extract data rows (skip title row 0 and header row 1; data starts at row 2)
                        for (int row = 2; row < rowCount; row++)
                        {
                            var position = LM_UnFormat(entity.GetText(row, 0) ?? "");
                            var productNo = LM_UnFormat(entity.GetText(row, 1) ?? "");
                            var width = LM_UnFormat(entity.GetText(row, 2) ?? "");
                            var height = LM_UnFormat(entity.GetText(row, 3) ?? "");
                            var length = LM_UnFormat(entity.GetText(row, 4) ?? "");
                            var thickness = LM_UnFormat(entity.GetText(row, 5) ?? "");
                            var qty = LM_UnFormat(entity.GetText(row, 6) ?? "");
                            var description = LM_UnFormat(entity.GetText(row, 7) ?? "");
                            var detailId = LM_UnFormat(entity.GetText(row, 8) ?? "");

                            // Filter empty rows (both qty AND product_no empty)
                            if (string.IsNullOrWhiteSpace(qty) && string.IsNullOrWhiteSpace(productNo))
                            {
                                continue;
                            }

                            var rowData = new Dictionary<string, object>
                            {
                                ["position"] = position,
                                ["product_no"] = productNo,
                                ["width"] = width,
                                ["height"] = height,
                                ["length"] = length,
                                ["thickness"] = thickness,
                                ["qty"] = qty,
                                ["description"] = description,
                                ["detail_id"] = detailId
                            };

                            tableRows.Add(rowData);
                        }

                        // Only process first valid table
                        break;
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "Error processing table entity");
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to extract table data");
        }

        return (headerId, tableRows);
    }

    /// <summary>
    /// Orchestrates extraction from specified layout.
    /// Switches to layout, extracts attributes and table data, combines into LayoutData.
    /// </summary>
    private LayoutData GetLayoutValuesInternal(string layoutName)
    {
        var data = new LayoutData();

        try
        {
            if (_acadDoc == null)
            {
                _logger?.LogWarning("No active document for layout extraction");
                return data;
            }

            // Switch to specified layout
            bool switched = SwitchToLayout(layoutName);
            if (!switched)
            {
                _logger?.LogWarning("Failed to switch to layout: {LayoutName}", layoutName);
                return data;
            }

            data.LayoutName = layoutName;

            // Get the layout object
            dynamic layout = _acadDoc.Layouts.Item(layoutName);

            // Extract block attributes
            data.Parameters = GetAttributeValues(layout);

            // Extract table data (cast to avoid dynamic tuple name loss)
            (string? headerId, List<Dictionary<string, object>> tableRows) = GetTableData((object)layout);

            // Store header_id in parameters for BuildImportRequest
            if (!string.IsNullOrWhiteSpace(headerId))
                data.Parameters["header_id"] = headerId;

            // Convert to TableData format
            if (tableRows.Count > 0)
            {
                var tableData = new TableData
                {
                    Name = "MainTable",
                    RowCount = tableRows.Count + 1, // +1 for header
                    ColumnCount = 9
                };

                // Add header row
                tableData.Cells.Add(new List<string>
                {
                    "Position", "Product No", "Width", "Height", "Length", 
                    "Thickness", "Qty", "Description", "HEADER_ID"
                });

                // Add data rows
                foreach (var row in tableRows)
                {
                    tableData.Cells.Add(new List<string>
                    {
                        row["position"]?.ToString() ?? "",
                        row["product_no"]?.ToString() ?? "",
                        row["width"]?.ToString() ?? "",
                        row["height"]?.ToString() ?? "",
                        row["length"]?.ToString() ?? "",
                        row["thickness"]?.ToString() ?? "",
                        row["qty"]?.ToString() ?? "",
                        row["description"]?.ToString() ?? "",
                        row["detail_id"]?.ToString() ?? ""
                    });
                }

                data.Tables.Add(tableData);
            }

            data.ExtractedAt = DateTime.UtcNow;
            int rowCount = tableRows.Count;
            int paramCount = data.Parameters.Count;
            _logger?.LogInformation("Extracted {ParamCount} parameters and {RowCount} table rows from {LayoutName}",
                (object)paramCount, (object)rowCount, (object)layoutName);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to extract layout values from {LayoutName}", layoutName);
        }

        return data;
    }

    public LayoutData GetLayoutsValues()
    {
        var data = new LayoutData();

        try
        {
            if (_acadDoc == null) return data;

            var currentLayout = GetCurrentLayoutName();
            data.LayoutName = currentLayout ?? "Unknown";

            // Extract parameters from title block and other entities
            // This is a simplified version - the full implementation would iterate through
            // all entities and extract text, attributes, and table data

            foreach (dynamic entity in _acadDoc.ModelSpace)
            {
                try
                {
                    string entityType = entity.EntityName;

                    if (entityType == "AcDbText" || entityType == "AcDbMText")
                    {
                        string textContent = entity.TextString;
                        // Parse and categorize text content
                    }
                    else if (entityType == "AcDbTable")
                    {
                        var tableData = ExtractTableData(entity);
                        data.Tables.Add(tableData);
                    }
                    else if (entityType == "AcDbBlockReference")
                    {
                        // Extract block attributes
                        ExtractBlockAttributes(entity, data.Parameters);
                    }
                }
                catch
                {
                    // Skip entities that can't be read
                }
            }

            data.ExtractedAt = DateTime.UtcNow;
            return data;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to extract layout values");
            return data;
        }
    }

    public LayoutData GetLayoutValues(string layoutName)
    {
        return GetLayoutValuesInternal(layoutName);
    }

    public async Task<IReadOnlyList<string>> ExportLayoutImagesAsync(string outputDirectory, string format = "PNG")
    {
        var exportedFiles = new List<string>();

        // Implementation would use plot functionality to export layouts as images
        // This is a placeholder for the actual implementation

        await Task.CompletedTask;
        return exportedFiles;
    }

    private TableData ExtractTableData(dynamic table)
    {
        var data = new TableData
        {
            RowCount = table.Rows,
            ColumnCount = table.Columns
        };

        for (int row = 0; row < data.RowCount; row++)
        {
            var rowData = new List<string>();
            for (int col = 0; col < data.ColumnCount; col++)
            {
                try
                {
                    rowData.Add(table.GetText(row, col) ?? "");
                }
                catch
                {
                    rowData.Add("");
                }
            }
            data.Cells.Add(rowData);
        }

        return data;
    }

    private void ExtractBlockAttributes(dynamic block, Dictionary<string, object> parameters)
    {
        try
        {
            if (!block.HasAttributes) return;

            foreach (dynamic attr in block.GetAttributes())
            {
                string tag = attr.TagString;
                string value = attr.TextString;
                parameters[tag] = value;
            }
        }
        catch
        {
            // Skip blocks without readable attributes
        }
    }

    #endregion

    #region Table Operations

    public void SetTableValue(string tableHandle, int row, int column, string value)
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var table = _acadDoc.HandleToObject(tableHandle);
            table.SetText(row, column, value);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to set table value");
            throw;
        }
    }

    public string GetTableValue(string tableHandle, int row, int column)
    {
        try
        {
            if (_acadDoc == null) throw new InvalidOperationException("No document open");

            var table = _acadDoc.HandleToObject(tableHandle);
            return table.GetText(row, column) ?? "";
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to get table value");
            throw;
        }
    }

    public IReadOnlyList<string> GetTables()
    {
        var tables = new List<string>();

        try
        {
            if (_acadDoc == null) return tables;

            foreach (dynamic entity in _acadDoc.ModelSpace)
            {
                if (entity.EntityName == "AcDbTable")
                {
                    tables.Add(entity.Handle);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to get tables");
        }

        return tables;
    }

    #endregion

    #region Layer Operations

    public bool CreateLayer(string name, int color = 7)
    {
        try
        {
            if (_acadDoc == null) return false;

            var layers = _acadDoc.Layers;
            var layer = layers.Add(name);
            layer.Color = color;

            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to create layer: {LayerName}", name);
            return false;
        }
    }

    public bool SetCurrentLayer(string name)
    {
        try
        {
            if (_acadDoc == null) return false;

            _acadDoc.ActiveLayer = _acadDoc.Layers.Item(name);
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to set current layer: {LayerName}", name);
            return false;
        }
    }

    public IReadOnlyList<string> GetLayers()
    {
        var layerNames = new List<string>();

        try
        {
            if (_acadDoc == null) return layerNames;

            foreach (dynamic layer in _acadDoc.Layers)
            {
                layerNames.Add(layer.Name);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to get layers");
        }

        return layerNames;
    }

    #endregion

    #region Selection Operations

    public void SelectEntities(IEnumerable<string> handles)
    {
        try
        {
            if (_acadDoc == null) return;

            var selSet = _acadDoc.SelectionSets.Add("TempSelection_" + Guid.NewGuid().ToString("N"));

            foreach (var handle in handles)
            {
                var entity = _acadDoc.HandleToObject(handle);
                selSet.AddItems(new object[] { entity });
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to select entities");
        }
    }

    public void ClearSelection()
    {
        try
        {
            if (_acadDoc == null) return;
            _acadDoc.SendCommand("_SELECT _NONE ");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to clear selection");
        }
    }

    public IReadOnlyList<string> GetSelectedEntities()
    {
        var handles = new List<string>();

        try
        {
            if (_acadDoc == null) return handles;

            var selSet = _acadDoc.ActiveSelectionSet;
            foreach (dynamic entity in selSet)
            {
                handles.Add(entity.Handle);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to get selected entities");
        }

        return handles;
    }

    #endregion

    #region Zoom Operations

    public void ZoomExtents()
    {
        try
        {
            if (_acadApp == null) return;
            _acadApp.ZoomExtents();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to zoom extents");
        }
    }

    public void ZoomWindow(Point3D corner1, Point3D corner2)
    {
        try
        {
            if (_acadApp == null) return;

            var pt1 = new double[] { corner1.X, corner1.Y, corner1.Z };
            var pt2 = new double[] { corner2.X, corner2.Y, corner2.Z };

            _acadApp.ZoomWindow(pt1, pt2);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to zoom window");
        }
    }

    public void ZoomSelected()
    {
        try
        {
            if (_acadDoc == null) return;
            _acadDoc.SendCommand("_ZOOM _OBJECT ");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to zoom selected");
        }
    }

    #endregion

    public void Dispose()
    {
        _isConnected = false;
        ReleaseCOMObjects();
        GC.SuppressFinalize(this);
    }
}
