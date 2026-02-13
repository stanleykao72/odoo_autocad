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
        "product_name", "product_catelog", "spec",
        "surface_treatment", "operation_flow",
        "color_name", "color_no",
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

        // Diagnostic: collect info about all Insert blocks in the layout
        var insertDiag = new List<string>();

        foreach (var entity in block.Entities)
        {
            if (entity is Insert insert)
            {
                var blockName = insert.Block?.Name ?? "(null)";
                var attrCount = insert.Attributes?.Count ?? 0;
                var attrTags = attrCount > 0
                    ? string.Join(", ", insert.Attributes!.Select(a => $"{a.Tag}"))
                    : "none";

                // Also check block definition for AttributeDefinitions (default values)
                var attrDefTags = new List<string>();
                if (insert.Block != null)
                {
                    foreach (var be in insert.Block.Entities)
                    {
                        if (be is AttributeDefinition attrDef)
                            attrDefTags.Add(attrDef.Tag ?? "?");
                    }
                }
                var defInfo = attrDefTags.Count > 0 ? $" defs=[{string.Join(",", attrDefTags)}]" : "";
                insertDiag.Add($"'{blockName}' attrs({attrCount})=[{attrTags}]{defInfo}");

                // Special case: "pr_no" block stores value as Text inside its block definition
                // (not as an attribute). Matches COM path and Python get_block_text() pattern.
                if (string.Equals(insert.Block?.Name, "pr_no", StringComparison.OrdinalIgnoreCase)
                    && !parameters.ContainsKey("pr_no"))
                {
                    var prText = ExtractBlockDefinitionText(insert);
                    if (!string.IsNullOrWhiteSpace(prText))
                    {
                        parameters["pr_no"] = prText;
                    }
                }

                ExtractInsertAttributes(insert, parameters, tableRows);
            }
        }

        // Store diagnostic info in parameters for UI logging
        if (insertDiag.Count > 0)
        {
            parameters["_diag_blocks"] = string.Join(" | ", insertDiag);
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
    /// Regex to strip AutoCAD duplicate-attribute suffixes like _001, _002 etc.
    /// When a block has multiple attribute definitions with the same tag,
    /// AutoCAD appends _NNN to make them unique in the DWG file.
    /// The COM API normalizes these back; we do the same here.
    /// </summary>
    private static readonly Regex AttributeSuffixPattern = new(@"_\d{3}$", RegexOptions.Compiled);

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

            // Normalize tag: strip _NNN suffix (e.g. project_name_001 → project_name)
            var normalizedTag = AttributeSuffixPattern.Replace(tag, "");

            // Store known parameters using normalized tag name (first occurrence wins)
            if (KnownAttributeTags.Contains(normalizedTag) && !parameters.ContainsKey(normalizedTag))
            {
                parameters[normalizedTag] = value;
            }

            rowValues.Add(value);
        }

        if (rowValues.Count > 0)
        {
            tableRows.Add(rowValues);
        }
    }

    /// <summary>
    /// Reads text from inside a block definition referenced by an Insert entity.
    /// Matches the COM path (AutoCADService.GetAttributeValues) and Python get_block_text():
    /// find the first AcDbText entity inside the block definition and return its text.
    /// Used for special blocks like "pr_no" that store values as text, not attributes.
    /// </summary>
    private string? ExtractBlockDefinitionText(Insert insert)
    {
        try
        {
            var blockDef = insert.Block;
            if (blockDef == null) return null;

            foreach (var entity in blockDef.Entities)
            {
                if (entity is TextEntity textEntity)
                {
                    var text = StripMTextFormatting(textEntity.Value ?? string.Empty);
                    if (!string.IsNullOrWhiteSpace(text))
                        return text;
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to read block definition text for Insert '{BlockName}'",
                insert.Block?.Name);
        }

        return null;
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

        // \P or \p (paragraph break) → space
        result = Regex.Replace(result, @"\\[Pp]", " ");

        // \C# or \c# (color)
        result = Regex.Replace(result, @"\\[Cc]\d+;", "");

        // \F or \f (font, e.g. \fPMingLiU;)
        result = Regex.Replace(result, @"\\[Ff][^;]*;", "");

        // \H or \h (height)
        result = Regex.Replace(result, @"\\[Hh][^;]*;", "");

        // \S or \s (stacking)
        result = Regex.Replace(result, @"\\[Ss][^;]*;", "");

        // \Q or \q (obliquing angle)
        result = Regex.Replace(result, @"\\[Qq]\d+;", "");

        // \T or \t (tracking)
        result = Regex.Replace(result, @"\\[Tt]\d+;", "");

        // \W or \w (width factor)
        result = Regex.Replace(result, @"\\[Ww]\d+\.?\d*;", "");

        // \A or \a (alignment)
        result = Regex.Replace(result, @"\\[Aa]\d+;", "");

        // Toggle codes
        result = Regex.Replace(result, @"\\[LlOo]", "");

        // \~ (non-breaking space) → space
        result = result.Replace("\\~", " ");

        // {} group delimiters
        result = result.Replace("{", "").Replace("}", "");

        // \\ → single backslash
        result = result.Replace("\\\\", "\\");

        return result.Trim();
    }
}
