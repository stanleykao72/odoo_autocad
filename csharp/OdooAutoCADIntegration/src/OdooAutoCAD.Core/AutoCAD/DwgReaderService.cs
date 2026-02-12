// OdooAutoCAD.Core/AutoCAD/DwgReaderService.cs
// File-based DWG reader using ACadSharp — no COM, no running AutoCAD instance required.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using Microsoft.Extensions.Logging;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Reads DWG files using ACadSharp (pure C#, no COM).
/// Thread-safe: each call opens/reads/closes the file independently.
/// </summary>
public class DwgReaderService : IDwgReaderService
{
    private readonly ILogger<DwgReaderService>? _logger;

    /// <summary>
    /// Known block attribute tags that map to layout parameters.
    /// Matches the COM-based extraction in AutoCADService.
    /// </summary>
    private static readonly HashSet<string> KnownAttributeTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "pr_no", "project_name", "job_working_plan_name",
        "product_name", "spec", "color_name",
        "unit", "remarks", "block_name", "quantity"
    };

    public DwgReaderService(ILogger<DwgReaderService>? logger = null)
    {
        _logger = logger;
    }

    public List<LayoutInfo> GetLayouts(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException("DWG file not found.", filePath);

        _logger?.LogInformation("Reading layouts from: {FilePath}", filePath);

        var doc = DwgReader.Read(filePath);
        var layouts = new List<LayoutInfo>();

        if (doc.Layouts == null)
        {
            _logger?.LogWarning("No layouts collection found in: {FilePath}", filePath);
            return layouts;
        }

        foreach (var layout in doc.Layouts)
        {
            bool isModel = string.Equals(layout.Name, "Model", StringComparison.OrdinalIgnoreCase);

            if (!isModel)
            {
                layouts.Add(new LayoutInfo(
                    Name: layout.Name,
                    TabOrder: layout.TabOrder,
                    IsModelSpace: false,
                    PlotConfigurationName: string.Empty));
            }
        }

        layouts.Sort((a, b) => a.TabOrder.CompareTo(b.TabOrder));

        _logger?.LogInformation("Found {Count} paper-space layouts in: {FilePath}", layouts.Count, filePath);
        return layouts;
    }

    public LayoutData ExtractParameters(string filePath, string layoutName)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        if (string.IsNullOrWhiteSpace(layoutName))
            throw new ArgumentException("Layout name cannot be empty.", nameof(layoutName));
        if (!File.Exists(filePath))
            throw new FileNotFoundException("DWG file not found.", filePath);

        _logger?.LogInformation("Extracting parameters from layout '{Layout}' in: {FilePath}", layoutName, filePath);

        var doc = DwgReader.Read(filePath);
        var result = new LayoutData { LayoutName = layoutName };

        // Find the target layout
        var targetLayout = doc.Layouts?
            .FirstOrDefault(l => string.Equals(l.Name, layoutName, StringComparison.OrdinalIgnoreCase));

        if (targetLayout == null)
        {
            _logger?.LogWarning("Layout '{Layout}' not found in: {FilePath}", layoutName, filePath);
            return result;
        }

        var block = targetLayout.AssociatedBlock;
        if (block == null)
        {
            _logger?.LogWarning("Layout '{Layout}' has no associated block record", layoutName);
            return result;
        }

        // Extract block attributes from Insert entities
        var parameters = new Dictionary<string, object>();
        var tableRows = new List<List<string>>();

        foreach (var entity in block.Entities)
        {
            if (entity is Insert insert)
            {
                ExtractInsertAttributes(insert, parameters, tableRows);
            }
        }

        result.Parameters = parameters;

        // If we collected table-like rows, package them as TableData
        if (tableRows.Count > 0)
        {
            var table = new TableData
            {
                Name = "BlockAttributes",
                RowCount = tableRows.Count,
                ColumnCount = tableRows.Max(r => r.Count),
                Cells = tableRows
            };
            result.Tables.Add(table);
        }

        _logger?.LogInformation(
            "Extracted {ParamCount} parameters and {RowCount} table rows from layout '{Layout}'",
            parameters.Count, tableRows.Count, layoutName);

        return result;
    }

    public bool IsValidDwgFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            return false;

        try
        {
            var doc = DwgReader.Read(filePath);
            return doc != null;
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "File is not a valid DWG: {FilePath}", filePath);
            return false;
        }
    }

    /// <summary>
    /// Extracts attributes from an Insert entity.
    /// Known attribute tags go into the parameters dictionary.
    /// All attributes from a single insert are also collected as a table row.
    /// </summary>
    private void ExtractInsertAttributes(
        Insert insert,
        Dictionary<string, object> parameters,
        List<List<string>> tableRows)
    {
        if (insert.Attributes == null || insert.Attributes.Count == 0)
            return;

        var rowValues = new List<string>();

        foreach (var attr in insert.Attributes)
        {
            var tag = attr.Tag?.Trim() ?? string.Empty;
            var value = StripMTextFormatting(attr.Value ?? string.Empty);

            if (string.IsNullOrEmpty(tag))
                continue;

            // Store known parameters (first occurrence wins)
            if (KnownAttributeTags.Contains(tag) && !parameters.ContainsKey(tag))
            {
                parameters[tag] = value;
            }

            rowValues.Add(value);
        }

        if (rowValues.Count > 0)
        {
            tableRows.Add(rowValues);
        }
    }

    /// <summary>
    /// Strips AutoCAD MText formatting codes.
    /// Mirrors AutoCADService.LM_UnFormat() so results are consistent.
    /// </summary>
    private static string StripMTextFormatting(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var result = text;

        // \P (paragraph break) → space
        result = Regex.Replace(result, @"\\P", " ");

        // \C# (color)
        result = Regex.Replace(result, @"\\C\d+;", "");

        // \F (font)
        result = Regex.Replace(result, @"\\F[^;]*;", "");

        // \H (height)
        result = Regex.Replace(result, @"\\H[^;]*;", "");

        // \S (stacking)
        result = Regex.Replace(result, @"\\S[^;]*;", "");

        // \Q (obliquing angle)
        result = Regex.Replace(result, @"\\Q\d+;", "");

        // \T (tracking)
        result = Regex.Replace(result, @"\\T\d+;", "");

        // \W (width factor)
        result = Regex.Replace(result, @"\\W\d+\.?\d*;", "");

        // \A (alignment)
        result = Regex.Replace(result, @"\\A\d+;", "");

        // Toggle codes
        result = result.Replace("\\L", "");
        result = result.Replace("\\l", "");
        result = result.Replace("\\O", "");
        result = result.Replace("\\o", "");

        // \~ (non-breaking space) → space
        result = result.Replace("\\~", " ");

        // {} group delimiters
        result = result.Replace("{", "").Replace("}", "");

        // \\ → single backslash
        result = result.Replace("\\\\", "\\");

        return result.Trim();
    }
}
