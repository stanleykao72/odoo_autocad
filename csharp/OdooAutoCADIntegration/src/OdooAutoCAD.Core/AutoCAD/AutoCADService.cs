// OdooAutoCAD.Core/AutoCAD/AutoCADService.cs
// AutoCAD COM Service Implementation - equivalent to Python util_autocad.py

using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// AutoCAD COM Service implementation.
/// Uses late binding (dynamic) for AutoCAD LT compatibility.
///
/// All methods in this class must be called from the GUI (STA) thread
/// via the GUIProxy system to ensure thread safety.
/// </summary>
public class AutoCADService : IAutoCADService, IDisposable
{
    private readonly ILogger<AutoCADService>? _logger;
    private readonly IGUIProxy _guiProxy;
    private dynamic? _acadApp;
    private dynamic? _acadDoc;
    private bool _isConnected;

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
    /// This ensures all COM operations are executed on the STA thread.
    /// </summary>
    private void RegisterGUIProxyHandlers()
    {
        _guiProxy.RegisterHandler("autocad_connect", async (parameters) =>
        {
            return await Task.FromResult(ConnectInternal());
        });

        _guiProxy.RegisterHandler("autocad_disconnect", async (parameters) =>
        {
            DisconnectInternal();
            return await Task.FromResult<object?>(null);
        });

        _guiProxy.RegisterHandler("autocad_get_status", async (parameters) =>
        {
            return await Task.FromResult(GetStatusInternal());
        });

        _guiProxy.RegisterHandler("autocad_get_layouts", async (parameters) =>
        {
            return await Task.FromResult(GetLayoutsInternal());
        });

        _guiProxy.RegisterHandler("autocad_get_active_layout", async (parameters) =>
        {
            return await Task.FromResult(GetCurrentLayoutName());
        });

        _guiProxy.RegisterHandler("autocad_set_active_layout", async (parameters) =>
        {
            parameters.TryGetValue("layoutName", out var val);
            var layoutName = val as string;
            if (layoutName != null)
            {
                return await Task.FromResult(SwitchToLayout(layoutName));
            }
            return await Task.FromResult(false);
        });

        _guiProxy.RegisterHandler("autocad_extract_parameters", async (parameters) =>
        {
            parameters.TryGetValue("layoutName", out var val);
            var layoutName = val as string;
            if (layoutName != null)
            {
                return await Task.FromResult(GetLayoutValuesInternal(layoutName));
            }
            return await Task.FromResult(new LayoutData());
        });

        _guiProxy.RegisterHandler("autocad_open_document", async (parameters) =>
        {
            parameters.TryGetValue("filePath", out var val);
            var filePath = val as string;
            if (string.IsNullOrEmpty(filePath) || !_isConnected) return await Task.FromResult(false);
            try
            {
                _acadDoc = _acadApp!.Documents.Open(filePath);
                _logger?.LogInformation("Opened document: {FilePath}", filePath);
                return await Task.FromResult(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to open document: {FilePath}", filePath);
                return await Task.FromResult(false);
            }
        });

        _guiProxy.RegisterHandler("autocad_save_document", async (parameters) =>
        {
            if (_acadDoc == null) return await Task.FromResult(false);
            try
            {
                _acadDoc.Save();
                _logger?.LogInformation("Document saved");
                return await Task.FromResult(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to save document");
                return await Task.FromResult(false);
            }
        });

        _guiProxy.RegisterHandler("autocad_close_document", async (parameters) =>
        {
            if (_acadDoc == null) return await Task.FromResult(false);
            try
            {
                var save = true;
                if (parameters.TryGetValue("save", out var saveProp) && saveProp is bool s)
                    save = s;
                _acadDoc.Close(save);
                _acadDoc = _acadApp?.ActiveDocument;
                _logger?.LogInformation("Document closed");
                return await Task.FromResult(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to close document");
                return await Task.FromResult(false);
            }
        });

        _logger?.LogDebug("AutoCAD GUI proxy handlers registered");
    }

    // Only use the boolean flag — touching _acadApp from a non-STA thread triggers COM marshaling
    // and can cause access violations inside AutoCAD.
    public bool IsConnected => _isConnected;

    #region Connection Management

    public async Task<bool> ConnectAsync()
    {
        var response = await _guiProxy.ExecuteInGuiAsync("autocad_connect", null, timeout: 15000);
        return response.Success && response.Result is bool connected && connected;
    }

    /// <summary>
    /// Internal connection method executed on GUI thread.
    /// Implements retry logic for ActiveDocument (5 attempts, 1-second intervals).
    /// Uses GetActiveObject → CreateInstance fallback pattern.
    /// </summary>
    private bool ConnectInternal()
    {
        try
        {
            // Register COM message filter to handle rejected calls from AutoCAD.
            // AutoCAD (2010+) rejects COM calls during WPF layout processing,
            // which causes "Unhandled Access Violation" crashes without this filter.
            // The filter auto-retries rejected calls after 1 second.
            OleMessageFilter.Register();

            bool isNewInstance = false;

            // Try to get running instance first using GetActiveObject
            // (matches Python: client.GetActiveObject("AutoCAD.Application"))
            try
            {
                _acadApp = GetActiveObject(DefaultProgId);
                _logger?.LogInformation("Connected to existing AutoCAD instance via GetActiveObject");
            }
            catch (COMException ex)
            {
                _logger?.LogDebug(ex, "GetActiveObject failed, trying Activator.CreateInstance");

                // Fallback: Try to create new instance
                // (matches Python: client.Dispatch("AutoCAD.Application"))
                var acadType = Type.GetTypeFromProgID(DefaultProgId);
                if (acadType == null)
                {
                    _logger?.LogError("AutoCAD is not installed or ProgID not found");
                    return false;
                }

                _acadApp = Activator.CreateInstance(acadType);
                isNewInstance = true;
                _logger?.LogInformation("Created new AutoCAD instance");
            }

            if (_acadApp == null)
            {
                _logger?.LogError("Failed to obtain AutoCAD COM object");
                return false;
            }

            // Only set Visible on newly created instances (matching Python behavior).
            // Setting Visible on an already-running AutoCAD 2014 instance can trigger
            // an access violation (0x0050) inside AutoCAD's COM server.
            // Also: no Thread.Sleep here — blocking the STA message pump prevents COM
            // message processing and can cause the out-of-process server to crash.
            if (isNewInstance)
            {
                try
                {
                    _acadApp.Visible = true;
                    _logger?.LogDebug("Set AutoCAD.Visible = true (new instance)");
                }
                catch (COMException ex)
                {
                    _logger?.LogWarning(ex, "Failed to set AutoCAD.Visible (non-fatal, continuing)");
                }
            }

            // Retry logic for ActiveDocument (5 attempts, 1-second intervals)
            // Matches Python: retry_count = 5, time.sleep(1)
            for (int attempt = 1; attempt <= MaxRetryAttempts; attempt++)
            {
                try
                {
                    _acadDoc = _acadApp.ActiveDocument;
                    if (_acadDoc != null)
                    {
                        _logger?.LogInformation("ActiveDocument acquired on attempt {Attempt}", attempt);
                        break;
                    }
                }
                catch (COMException ex)
                {
                    _logger?.LogDebug(ex, "ActiveDocument attempt {Attempt} failed", attempt);
                }

                if (attempt < MaxRetryAttempts)
                {
                    Thread.Sleep(RetryDelayMs);
                }
            }

            if (_acadDoc == null)
            {
                _logger?.LogWarning("ActiveDocument is null after {Attempts} attempts", MaxRetryAttempts);
                // Still consider it connected even if no document is open
            }

            _isConnected = true;
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to connect to AutoCAD");
            _acadDoc = null;
            _acadApp = null;
            _isConnected = false;
            return false;
        }
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
            try { Marshal.ReleaseComObject(_acadDoc); } catch { /* ignore release errors */ }
            _acadDoc = null;
        }
        if (_acadApp != null)
        {
            try { Marshal.ReleaseComObject(_acadApp); } catch { /* ignore release errors */ }
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
            string appName = _acadApp.Name;
            string version = _acadApp.Version;
            string? currentDoc = _acadDoc?.Name;

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
                ErrorMessage: null);
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
    /// Strips AutoCAD MText formatting codes.
    /// Removes codes like \P (paragraph), \C (color), \F (font), \H (height), etc.
    /// </summary>
    private string LM_UnFormat(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        // Remove common MText formatting codes
        var result = text;

        // Remove \P (paragraph break) - replace with space
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\P", " ");

        // Remove \C# (color codes like \C1, \C255)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\C\d+;", "");

        // Remove \F (font)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\F[^;]*;", "");

        // Remove \H (height)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\H[^;]*;", "");

        // Remove \S (stacking)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\S[^;]*;", "");

        // Remove \Q (obliquing angle)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\Q\d+;", "");

        // Remove \T (tracking)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\T\d+;", "");

        // Remove \W (width factor)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\W\d+\.?\d*;", "");

        // Remove \A (alignment)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\\A\d+;", "");

        // Remove \L (underline on)
        result = result.Replace("\\L", "");

        // Remove \l (underline off)
        result = result.Replace("\\l", "");

        // Remove \O (overline on)
        result = result.Replace("\\O", "");

        // Remove \o (overline off)
        result = result.Replace("\\o", "");

        // Remove \~ (non-breaking space) - replace with space
        result = result.Replace("\\~", " ");

        // Remove curly braces {} (group delimiters)
        result = result.Replace("{", "").Replace("}", "");

        // Remove \\ (escaped backslash) - replace with single backslash
        result = result.Replace("\\\\", "\\");

        // Trim whitespace
        result = result.Trim();

        return result;
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

                    // Check if it's a block reference with attributes
                    if (entityType == "AcDbBlockReference")
                    {
                        if (entity.HasAttributes)
                        {
                            foreach (dynamic attr in entity.GetAttributes())
                            {
                                string tag = attr.TagString;
                                string value = LM_UnFormat(attr.TextString);

                                // Map standard attribute tags
                                switch (tag.ToLower())
                                {
                                    case "pr_no":
                                    case "project_name":
                                    case "job_working_plan_name":
                                    case "product_name":
                                    case "spec":
                                    case "color_name":
                                    case "unit":
                                    case "remarks":
                                    case "block_name":
                                    case "quantity":
                                        attributes[tag] = value;
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
    /// Validates: exactly 9 columns, HEADER_ID in column 7 (index 6).
    /// Returns rows with: Position, Product No, Width, Height, Length, Thickness, Qty, Description, Detail ID.
    /// Filters empty rows (both qty AND product_no empty).
    /// </summary>
    private List<Dictionary<string, object>> GetTableData(dynamic layout)
    {
        var tableRows = new List<Dictionary<string, object>>();

        try
        {
            if (_acadDoc == null) return tableRows;

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

                        // Validate: column 6 (index 6, 7th column) contains "HEADER_ID" in header
                        string headerCell = entity.GetText(0, 6) ?? "";
                        if (!headerCell.Contains("HEADER_ID"))
                        {
                            _logger?.LogWarning("Column 6 does not contain HEADER_ID. Found: {Header}", headerCell);
                            continue;
                        }

                        // Extract data rows (skip header row 0)
                        for (int row = 1; row < rowCount; row++)
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

        return tableRows;
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

            // Extract table data
            var tableRows = GetTableData(layout);

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
