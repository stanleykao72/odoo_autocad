// OdooAutoCAD.Core/AutoCAD/IDwgReaderService.cs
// File-based DWG reader interface - no COM dependency, no AutoCAD instance required.
// Uses ACadSharp for pure C# DWG/DXF parsing.

using System.Collections.Generic;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Interface for file-based DWG reading without COM or a running AutoCAD instance.
/// Uses ACadSharp to parse DWG files directly.
/// Thread-safe: no COM objects, no STA thread requirement.
/// </summary>
public interface IDwgReaderService
{
    /// <summary>
    /// Gets all non-Model layouts from a DWG file.
    /// </summary>
    /// <param name="filePath">Path to the .dwg file.</param>
    /// <returns>List of layouts excluding Model space.</returns>
    List<LayoutInfo> GetLayouts(string filePath);

    /// <summary>
    /// Extracts block attribute parameters from a specific layout in a DWG file.
    /// </summary>
    /// <param name="filePath">Path to the .dwg file.</param>
    /// <param name="layoutName">Name of the layout to extract from.</param>
    /// <returns>Layout data with extracted parameters and table-like data.</returns>
    LayoutData ExtractParameters(string filePath, string layoutName);

    /// <summary>
    /// Validates that a file exists and can be read as a DWG file.
    /// </summary>
    /// <param name="filePath">Path to check.</param>
    /// <returns>True if the file is a valid, readable DWG file.</returns>
    bool IsValidDwgFile(string filePath);
}
