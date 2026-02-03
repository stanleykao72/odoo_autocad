// OdooAutoCAD.Core/AutoCAD/AutoCADService.cs
// AutoCAD COM Service Implementation - equivalent to Python util_autocad.py

using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;

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
    private dynamic? _acadApp;
    private dynamic? _acadDoc;
    private bool _isConnected;

    private const string DefaultProgId = "AutoCAD.Application";

    public AutoCADService(ILogger<AutoCADService>? logger = null)
    {
        _logger = logger;
    }

    public bool IsConnected => _isConnected && _acadApp != null;

    #region Connection Management

    public async Task<bool> ConnectAsync()
    {
        return await Task.Run(() => Connect());
    }

    private bool Connect()
    {
        try
        {
            // Try to get running instance first
            try
            {
                _acadApp = Marshal.GetActiveObject(DefaultProgId);
                _logger?.LogInformation("Connected to existing AutoCAD instance");
            }
            catch (COMException)
            {
                // No running instance, try to create new
                var acadType = Type.GetTypeFromProgID(DefaultProgId);
                if (acadType == null)
                {
                    _logger?.LogError("AutoCAD is not installed or ProgID not found");
                    return false;
                }

                _acadApp = Activator.CreateInstance(acadType);
                _acadApp.Visible = true;
                _logger?.LogInformation("Started new AutoCAD instance");
            }

            _acadDoc = _acadApp?.ActiveDocument;
            _isConnected = true;

            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to connect to AutoCAD");
            _isConnected = false;
            return false;
        }
    }

    public async Task DisconnectAsync()
    {
        await Task.Run(() =>
        {
            _acadDoc = null;
            _acadApp = null;
            _isConnected = false;
            _logger?.LogInformation("Disconnected from AutoCAD");
        });
    }

    public async Task<AutoCADStatus> GetStatusAsync()
    {
        return await Task.Run(() =>
        {
            if (!IsConnected || _acadApp == null)
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
                return new AutoCADStatus(
                    IsConnected: false,
                    ApplicationName: null,
                    Version: null,
                    CurrentDocument: null,
                    OpenDocuments: null,
                    ErrorMessage: ex.Message);
            }
        });
    }

    #endregion

    #region Document Management

    public async Task<bool> OpenDocumentAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (!IsConnected) return false;

                _acadDoc = _acadApp.Documents.Open(filePath);
                _logger?.LogInformation("Opened document: {FilePath}", filePath);
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to open document: {FilePath}", filePath);
                return false;
            }
        });
    }

    public async Task<bool> SaveDocumentAsync()
    {
        return await Task.Run(() =>
        {
            try
            {
                if (_acadDoc == null) return false;

                _acadDoc.Save();
                _logger?.LogInformation("Document saved");
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to save document");
                return false;
            }
        });
    }

    public async Task<bool> CloseDocumentAsync(bool save = true)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (_acadDoc == null) return false;

                _acadDoc.Close(save);
                _acadDoc = _acadApp?.ActiveDocument;
                _logger?.LogInformation("Document closed");
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to close document");
                return false;
            }
        });
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
        var layouts = new List<LayoutInfo>();

        try
        {
            if (_acadDoc == null) return layouts;

            foreach (dynamic layout in _acadDoc.Layouts)
            {
                layouts.Add(new LayoutInfo(
                    Name: layout.Name,
                    TabOrder: layout.TabOrder,
                    IsModelSpace: layout.Name == "Model",
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
        SwitchToLayout(layoutName);
        return GetLayoutsValues();
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
        _acadDoc = null;
        _acadApp = null;
        _isConnected = false;

        GC.SuppressFinalize(this);
    }
}
