// OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs
// AutoCAD COM Service Interface - equivalent to Python util_autocad.py

using System.Collections.Generic;
using System.Threading.Tasks;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Represents a 3D point in AutoCAD coordinate space.
/// </summary>
public record Point3D(double X, double Y, double Z = 0);

/// <summary>
/// Information about an AutoCAD layout.
/// </summary>
public record LayoutInfo(
    string Name,
    int TabOrder,
    bool IsModelSpace,
    string PlotConfigurationName);

/// <summary>
/// Contains extracted data from a layout including tables and parameters.
/// </summary>
public class LayoutData
{
    public string LayoutName { get; set; } = string.Empty;
    public Dictionary<string, object> Parameters { get; set; } = new();
    public List<TableData> Tables { get; set; } = new();
    public DateTime ExtractedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Represents table data extracted from AutoCAD.
/// </summary>
public class TableData
{
    public string Name { get; set; } = string.Empty;
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public List<List<string>> Cells { get; set; } = new();
}

/// <summary>
/// Dimension types supported by AutoCAD.
/// </summary>
public enum DimensionType
{
    Linear,
    Aligned,
    Angular,
    Radial,
    Diameter,
    Ordinate
}

/// <summary>
/// Text alignment options.
/// </summary>
public enum TextAlignment
{
    Left,
    Center,
    Right,
    TopLeft,
    TopCenter,
    TopRight,
    MiddleLeft,
    MiddleCenter,
    MiddleRight,
    BottomLeft,
    BottomCenter,
    BottomRight
}

/// <summary>
/// AutoCAD connection status information.
/// </summary>
public record AutoCADStatus(
    bool IsConnected,
    string? ApplicationName,
    string? Version,
    string? CurrentDocument,
    List<string>? OpenDocuments,
    string? ErrorMessage);

/// <summary>
/// Interface for AutoCAD COM operations.
/// Equivalent to Python's util_autocad.py (1,301 lines).
/// All operations must be executed on STA (GUI) thread via GUIProxy.
/// </summary>
public interface IAutoCADService
{
    #region Connection Management

    /// <summary>
    /// Gets whether AutoCAD is currently connected.
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// Connects to a running AutoCAD instance or starts a new one.
    /// </summary>
    /// <returns>True if connection successful.</returns>
    Task<bool> ConnectAsync();

    /// <summary>
    /// Disconnects from AutoCAD.
    /// </summary>
    Task DisconnectAsync();

    /// <summary>
    /// Gets detailed status information about the AutoCAD connection.
    /// </summary>
    Task<AutoCADStatus> GetStatusAsync();

    #endregion

    #region Document Management

    /// <summary>
    /// Opens a drawing file.
    /// </summary>
    /// <param name="filePath">Path to the DWG file.</param>
    /// <returns>True if opened successfully.</returns>
    Task<bool> OpenDocumentAsync(string filePath);

    /// <summary>
    /// Saves the current document.
    /// </summary>
    Task<bool> SaveDocumentAsync();

    /// <summary>
    /// Closes the current document.
    /// </summary>
    /// <param name="save">Whether to save before closing.</param>
    Task<bool> CloseDocumentAsync(bool save = true);

    /// <summary>
    /// Gets the path of the currently active document.
    /// </summary>
    string? GetCurrentDocumentPath();

    #endregion

    #region Layout Operations

    /// <summary>
    /// Gets all layouts in the current document.
    /// </summary>
    IReadOnlyList<LayoutInfo> GetLayouts();

    /// <summary>
    /// Switches to a specified layout.
    /// </summary>
    /// <param name="layoutName">Name of the layout to switch to.</param>
    /// <returns>True if switch successful.</returns>
    bool SwitchToLayout(string layoutName);

    /// <summary>
    /// Gets the name of the currently active layout.
    /// </summary>
    string? GetCurrentLayoutName();

    #endregion

    #region Drawing Operations

    /// <summary>
    /// Draws a line between two points.
    /// </summary>
    /// <param name="start">Start point.</param>
    /// <param name="end">End point.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string DrawLine(Point3D start, Point3D end, string layer = "0");

    /// <summary>
    /// Draws a circle.
    /// </summary>
    /// <param name="center">Center point.</param>
    /// <param name="radius">Circle radius.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string DrawCircle(Point3D center, double radius, string layer = "0");

    /// <summary>
    /// Draws an arc.
    /// </summary>
    /// <param name="center">Center point.</param>
    /// <param name="radius">Arc radius.</param>
    /// <param name="startAngle">Start angle in degrees.</param>
    /// <param name="endAngle">End angle in degrees.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string DrawArc(Point3D center, double radius, double startAngle, double endAngle, string layer = "0");

    /// <summary>
    /// Draws a polyline through specified points.
    /// </summary>
    /// <param name="points">List of points.</param>
    /// <param name="closed">Whether to close the polyline.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string DrawPolyline(IEnumerable<Point3D> points, bool closed = false, string layer = "0");

    /// <summary>
    /// Draws a rectangle.
    /// </summary>
    /// <param name="corner1">First corner.</param>
    /// <param name="corner2">Opposite corner.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string DrawRectangle(Point3D corner1, Point3D corner2, string layer = "0");

    #endregion

    #region Text Operations

    /// <summary>
    /// Creates a single-line text entity.
    /// </summary>
    /// <param name="position">Insertion point.</param>
    /// <param name="content">Text content.</param>
    /// <param name="height">Text height.</param>
    /// <param name="rotation">Rotation angle in degrees.</param>
    /// <param name="alignment">Text alignment.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string CreateText(
        Point3D position,
        string content,
        double height,
        double rotation = 0,
        TextAlignment alignment = TextAlignment.Left,
        string layer = "0");

    /// <summary>
    /// Creates a multi-line text (MText) entity.
    /// </summary>
    /// <param name="position">Insertion point.</param>
    /// <param name="content">Text content.</param>
    /// <param name="width">Text width.</param>
    /// <param name="height">Text height.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string CreateMText(
        Point3D position,
        string content,
        double width,
        double height,
        string layer = "0");

    #endregion

    #region Dimension Operations

    /// <summary>
    /// Adds a dimension to the drawing.
    /// </summary>
    /// <param name="type">Type of dimension.</param>
    /// <param name="points">Points defining the dimension.</param>
    /// <param name="dimLineLocation">Location of dimension line.</param>
    /// <param name="layer">Layer name (default: "0").</param>
    /// <returns>Handle of created entity.</returns>
    string AddDimension(
        DimensionType type,
        IEnumerable<Point3D> points,
        Point3D dimLineLocation,
        string layer = "0");

    #endregion

    #region Parameter Extraction

    /// <summary>
    /// Extracts parameters and data from all layouts.
    /// This is the main method for extracting drawing data.
    /// </summary>
    /// <returns>Layout data containing parameters and tables.</returns>
    LayoutData GetLayoutsValues();

    /// <summary>
    /// Extracts parameters from a specific layout.
    /// </summary>
    /// <param name="layoutName">Layout name to extract from.</param>
    /// <returns>Layout data for the specified layout.</returns>
    LayoutData GetLayoutValues(string layoutName);

    /// <summary>
    /// Exports layout images to files.
    /// </summary>
    /// <param name="outputDirectory">Directory to save images.</param>
    /// <param name="format">Image format (PNG, JPEG, etc.).</param>
    /// <returns>List of exported file paths.</returns>
    Task<IReadOnlyList<string>> ExportLayoutImagesAsync(string outputDirectory, string format = "PNG");

    #endregion

    #region Table Operations

    /// <summary>
    /// Sets a value in a table cell.
    /// </summary>
    /// <param name="tableHandle">Handle of the table entity.</param>
    /// <param name="row">Row index (0-based).</param>
    /// <param name="column">Column index (0-based).</param>
    /// <param name="value">Value to set.</param>
    void SetTableValue(string tableHandle, int row, int column, string value);

    /// <summary>
    /// Gets a value from a table cell.
    /// </summary>
    /// <param name="tableHandle">Handle of the table entity.</param>
    /// <param name="row">Row index (0-based).</param>
    /// <param name="column">Column index (0-based).</param>
    /// <returns>Cell value.</returns>
    string GetTableValue(string tableHandle, int row, int column);

    /// <summary>
    /// Gets all tables in the current layout.
    /// </summary>
    /// <returns>List of table handles.</returns>
    IReadOnlyList<string> GetTables();

    #endregion

    #region Layer Operations

    /// <summary>
    /// Creates a new layer.
    /// </summary>
    /// <param name="name">Layer name.</param>
    /// <param name="color">Layer color index.</param>
    /// <returns>True if created successfully.</returns>
    bool CreateLayer(string name, int color = 7);

    /// <summary>
    /// Sets the current layer.
    /// </summary>
    /// <param name="name">Layer name.</param>
    /// <returns>True if set successfully.</returns>
    bool SetCurrentLayer(string name);

    /// <summary>
    /// Gets all layers in the document.
    /// </summary>
    /// <returns>List of layer names.</returns>
    IReadOnlyList<string> GetLayers();

    #endregion

    #region Selection Operations

    /// <summary>
    /// Selects entities by handles.
    /// </summary>
    /// <param name="handles">List of entity handles.</param>
    void SelectEntities(IEnumerable<string> handles);

    /// <summary>
    /// Clears the current selection.
    /// </summary>
    void ClearSelection();

    /// <summary>
    /// Gets handles of currently selected entities.
    /// </summary>
    /// <returns>List of entity handles.</returns>
    IReadOnlyList<string> GetSelectedEntities();

    #endregion

    #region Zoom Operations

    /// <summary>
    /// Zooms to fit all entities in the view.
    /// </summary>
    void ZoomExtents();

    /// <summary>
    /// Zooms to a specific window.
    /// </summary>
    /// <param name="corner1">First corner of zoom window.</param>
    /// <param name="corner2">Opposite corner of zoom window.</param>
    void ZoomWindow(Point3D corner1, Point3D corner2);

    /// <summary>
    /// Zooms to selected entities.
    /// </summary>
    void ZoomSelected();

    #endregion
}
