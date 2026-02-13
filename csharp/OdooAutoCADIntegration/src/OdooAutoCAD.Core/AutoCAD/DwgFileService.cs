// OdooAutoCAD.Core/AutoCAD/DwgFileService.cs
// Stateful DWG file service — extends DwgReaderService with file management,
// table extraction (9-column validation), header IDs, PR number, and DXF export.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Objects;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.Core.BOQ;

using AcLayout = ACadSharp.Objects.Layout;

namespace OdooAutoCAD.Core.AutoCAD;

/// <summary>
/// Stateful DWG file service. Loads a DWG once and provides layout, table, attribute,
/// and DXF export operations on the in-memory document.
/// Extends <see cref="DwgReaderService"/> for layout listing and attribute extraction.
/// </summary>
public class DwgFileService : DwgReaderService, IDwgFileService
{
    private readonly ILogger<DwgFileService>? _logger;
    private CadDocument? _document;
    private string? _loadedFilePath;
    private readonly List<string> _readWarnings = new();

    /// <summary>
    /// Known attribute tags for block identification.
    /// A block reference is considered an "attribute block" if it contains one of these tags.
    /// </summary>
    private static readonly HashSet<string> IdentifyingTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "project_name", "job_working_plan_name"
    };

    public DwgFileService(ILogger<DwgFileService>? logger = null) : base(logger)
    {
        _logger = logger;
    }

    // ── File info ─────────────────────────────────────────────

    public bool IsFileLoaded => _document != null;
    public string? LoadedFilePath => _loadedFilePath;
    public string? DwgVersion => _document?.Header?.Version.ToString();
    public bool CanWriteDwg => true; // TABLE entity writer implemented

    // ── Load / Unload ─────────────────────────────────────────

    public Task<bool> LoadFileAsync(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return Task.FromResult(false);

        if (!File.Exists(filePath))
        {
            _logger?.LogWarning("DWG file not found: {FilePath}", filePath);
            return Task.FromResult(false);
        }

        if (!Path.GetExtension(filePath).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
        {
            _logger?.LogWarning("Not a .dwg file: {FilePath}", filePath);
            return Task.FromResult(false);
        }

        try
        {
            // Unload previous if any
            _document = null;
            _loadedFilePath = null;
            _readWarnings.Clear();

            // Use notification handler to capture ACadSharp parse warnings/errors
            _document = DwgReader.Read(filePath, (sender, e) =>
            {
                var msg = $"[ACadSharp {e.NotificationType}] {e.Message}";
                if (e.Exception != null)
                    msg += $" — {e.Exception.GetType().Name}: {e.Exception.Message}";
                _readWarnings.Add(msg);
                _logger?.LogWarning("ACadSharp: {Type} — {Message}", e.NotificationType, e.Message);
            });
            _loadedFilePath = filePath;

            _logger?.LogInformation("Loaded DWG: {FilePath} (version: {Version}, warnings: {WarnCount})",
                filePath, DwgVersion, _readWarnings.Count);
            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to load DWG: {FilePath}", filePath);
            _document = null;
            _loadedFilePath = null;
            return Task.FromResult(false);
        }
    }

    public Task UnloadFileAsync()
    {
        _document = null;
        _loadedFilePath = null;
        _logger?.LogInformation("DWG file unloaded");
        return Task.CompletedTask;
    }

    // ── Table operations ──────────────────────────────────────

    public Task<List<LayoutTableData>> ExtractTableDataAsync(string layoutName)
    {
        var results = new List<LayoutTableData>();

        if (_document == null)
            return Task.FromResult(results);

        try
        {
            var layout = FindLayout(layoutName);
            if (layout?.AssociatedBlock == null)
            {
                _logger?.LogWarning("Layout '{Layout}' not found or has no associated block", layoutName);
                return Task.FromResult(results);
            }

            // Diagnostic: log all entity types in this layout
            var entityTypes = layout.AssociatedBlock.Entities
                .GroupBy(e => e.GetType().Name)
                .Select(g => $"{g.Key}({g.Count()})");
            _logger?.LogInformation("Layout '{Layout}' entities: {Types}",
                layoutName, string.Join(", ", entityTypes));

            foreach (var entity in layout.AssociatedBlock.Entities)
            {
                if (entity is not TableEntity table)
                    continue;

                _logger?.LogInformation("Found TableEntity in '{Layout}': {Rows}x{Cols}",
                    layoutName, table.Rows.Count, table.Columns.Count);

                if (table.Columns.Count != 9)
                {
                    _logger?.LogInformation("Skipping table with {Cols} columns (expected 9) in layout '{Layout}'",
                        table.Columns.Count, layoutName);
                    continue;
                }

                // Validate: column 7 header must contain "HEADER_ID"
                var headerCheck = ReadCellText(table, 0, 7);
                _logger?.LogInformation("Table col 7 header in '{Layout}': '{Header}'", layoutName, headerCheck);
                if (!headerCheck.Contains("HEADER_ID", StringComparison.OrdinalIgnoreCase))
                {
                    _logger?.LogInformation("Table in layout '{Layout}' missing HEADER_ID in col 7 header (got: '{Header}')",
                        layoutName, headerCheck);
                    continue;
                }

                // Diagnostic: dump cell content structure for first few cells
                LogCellDiagnostics(table, layoutName);

                var tableData = new LayoutTableData
                {
                    LayoutName = layoutName,
                    HeaderId = StripFormatting(ReadCellText(table, 0, 8))
                };

                // Data rows start at row 2 (row 0=title/header_id, row 1=column labels)
                for (int row = 2; row < table.Rows.Count; row++)
                {
                    var position = StripFormatting(ReadCellText(table, row, 0));
                    var productNo = StripFormatting(ReadCellText(table, row, 1));
                    var qty = StripFormatting(ReadCellText(table, row, 6));

                    // Filter empty rows (both qty and product_no empty)
                    if (string.IsNullOrWhiteSpace(qty) && string.IsNullOrWhiteSpace(productNo))
                        continue;

                    var rowData = new Dictionary<string, string>
                    {
                        ["position"] = position,
                        ["product_no"] = productNo,
                        ["width"] = StripFormatting(ReadCellText(table, row, 2)),
                        ["height"] = StripFormatting(ReadCellText(table, row, 3)),
                        ["len"] = StripFormatting(ReadCellText(table, row, 4)),
                        ["thickness"] = StripFormatting(ReadCellText(table, row, 5)),
                        ["qty"] = qty,
                        ["desc"] = StripFormatting(ReadCellText(table, row, 7)),
                        ["detail_id"] = StripFormatting(ReadCellText(table, row, 8))
                    };
                    tableData.Rows.Add(rowData);
                }

                if (tableData.Rows.Count == 0 && table.Rows.Count > 2)
                {
                    _logger?.LogWarning(
                        "Table in '{Layout}' has {RowCount} rows but 0 extractable data rows. " +
                        "Cell content may be Field-type (requires COM mode) or Build step incomplete.",
                        layoutName, table.Rows.Count);
                }

                results.Add(tableData);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to extract table data from layout '{Layout}'", layoutName);
            // Return empty list on ACadSharp failures (AC-09)
        }

        return Task.FromResult(results);
    }

    public async Task<IReadOnlyList<string>> GetHeaderIdsAsync()
    {
        var headerIds = new List<string>();

        if (_document?.Layouts == null)
            return headerIds;

        foreach (var layout in _document.Layouts)
        {
            if (string.Equals(layout.Name, "Model", StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                var tables = await ExtractTableDataAsync(layout.Name).ConfigureAwait(false);
                foreach (var table in tables)
                {
                    if (!string.IsNullOrWhiteSpace(table.HeaderId))
                        headerIds.Add(table.HeaderId);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "Error reading header IDs from layout '{Layout}'", layout.Name);
            }
        }

        return headerIds;
    }

    public Task<string> GetPRNumberAsync()
    {
        if (_document?.Layouts == null)
            return Task.FromResult(string.Empty);

        // Search for pr_no in the first non-Model layout's attribute blocks
        foreach (var layout in _document.Layouts.OrderBy(l => l.TabOrder))
        {
            if (string.Equals(layout.Name, "Model", StringComparison.OrdinalIgnoreCase))
                continue;

            var block = layout.AssociatedBlock;
            if (block == null)
                continue;

            foreach (var entity in block.Entities)
            {
                if (entity is not Insert insert || insert.Attributes == null)
                    continue;

                foreach (var attr in insert.Attributes)
                {
                    if (string.Equals(attr.Tag?.Trim(), "pr_no", StringComparison.OrdinalIgnoreCase))
                    {
                        var value = StripFormatting(attr.Value ?? string.Empty);
                        if (!string.IsNullOrWhiteSpace(value))
                            return Task.FromResult(value);
                    }
                }
            }
        }

        return Task.FromResult(string.Empty);
    }

    // ── Write operations ──────────────────────────────────────

    public Task<bool> SaveAsDxfAsync(string outputPath)
    {
        if (_document == null)
            return Task.FromResult(false);

        try
        {
            using var writer = new DxfWriter(outputPath, _document, false);
            writer.Write();

            _logger?.LogInformation("DXF exported to: {OutputPath}", outputPath);
            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to export DXF: {OutputPath}", outputPath);
            return Task.FromResult(false);
        }
    }

    public Task<bool> SaveAsDwgAsync(string outputPath)
    {
        if (_document == null)
            return Task.FromResult(false);

        try
        {
            using var writer = new DwgWriter(outputPath, _document);
            writer.Write();

            _logger?.LogInformation("DWG exported to: {OutputPath}", outputPath);
            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to export DWG: {OutputPath}", outputPath);
            return Task.FromResult(false);
        }
    }

    public void WriteTableIdsToDocument(string layoutName, string headerId, IList<WritebackDetail> details)
    {
        if (_document == null)
            throw new InvalidOperationException("No document loaded");

        var layout = FindLayout(layoutName);
        if (layout?.AssociatedBlock == null)
        {
            _logger?.LogWarning("WriteTableIds: Layout '{Layout}' not found", layoutName);
            return;
        }

        // Find 9-column TABLE entities in layout
        var tables = layout.AssociatedBlock.Entities
            .OfType<TableEntity>()
            .Where(t => t.Columns.Count == 9);

        foreach (var table in tables)
        {
            // Write header_id to cell(0, 8)
            SetCellText(table, 0, 8, headerId);

            // Write detail_ids matched by product_no in col 1
            // Data rows start at row 2 (row 0 = title, row 1 = headers)
            for (int row = 2; row < table.Rows.Count; row++)
            {
                var productNo = StripFormatting(ReadCellText(table, row, 1));
                if (string.IsNullOrWhiteSpace(productNo))
                    continue;

                var detail = details.FirstOrDefault(d =>
                    string.Equals(d.ProductNo, productNo, StringComparison.OrdinalIgnoreCase));
                if (detail?.DetailId != null)
                {
                    SetCellText(table, row, 8, detail.DetailId);
                }
            }
        }

        _logger?.LogInformation("WriteTableIds: header_id={HeaderId}, {DetailCount} details written to layout '{Layout}'",
            headerId, details.Count, layoutName);
    }

    /// <summary>
    /// Sets a text value in a specific TABLE cell.
    /// Creates content if the cell has none; updates existing content otherwise.
    /// </summary>
    private static void SetCellText(TableEntity table, int row, int col, string value)
    {
        if (row < 0 || row >= table.Rows.Count || col < 0 || col >= table.Columns.Count)
            return;

        var cell = table.GetCell(row, col);

        if (cell.Contents.Count > 0 && cell.Content != null)
        {
            // Update existing content
            cell.Content.Value.Text = value;
            cell.Content.Value.Value = value;
            cell.Content.Value.ValueType = TableEntity.CellValueType.String;
        }
        else
        {
            // Create new content entry
            var content = new TableEntity.CellContent();
            content.Value.Text = value;
            content.Value.Value = value;
            content.Value.ValueType = TableEntity.CellValueType.String;
            cell.Contents.Add(content);
        }
    }

    public void ApplyAttributesToDocument(Dictionary<string, string> attributes, string? layoutName = null)
    {
        if (_document?.Layouts == null || attributes == null || attributes.Count == 0)
            return;

        foreach (var layout in _document.Layouts)
        {
            if (string.Equals(layout.Name, "Model", StringComparison.OrdinalIgnoreCase))
                continue;

            // If layoutName specified, only apply to that layout
            if (!string.IsNullOrEmpty(layoutName) &&
                !string.Equals(layout.Name, layoutName, StringComparison.OrdinalIgnoreCase))
                continue;

            var block = layout.AssociatedBlock;
            if (block == null)
                continue;

            foreach (var entity in block.Entities)
            {
                if (entity is not Insert insert || insert.Attributes == null)
                    continue;

                // Only modify blocks that have identifying tags
                bool isAttributeBlock = insert.Attributes
                    .Any(a => IdentifyingTags.Contains(a.Tag?.Trim() ?? string.Empty));

                if (!isAttributeBlock)
                    continue;

                foreach (var attr in insert.Attributes)
                {
                    var tag = attr.Tag?.Trim() ?? string.Empty;
                    if (attributes.TryGetValue(tag, out var newValue))
                    {
                        attr.Value = newValue;
                    }
                }
            }
        }

        var target = string.IsNullOrEmpty(layoutName) ? "all layouts" : layoutName;
        _logger?.LogInformation("Applied {Count} attribute modifications to {Target} in-memory document",
            attributes.Count, target);
    }

    // ── Diagnostics ────────────────────────────────────────────

    public string GetLayoutEntityDiagnostics(string layoutName)
    {
        if (_document == null)
            return "No document loaded";

        var layout = FindLayout(layoutName);
        if (layout?.AssociatedBlock == null)
            return $"Layout '{layoutName}' not found or has no block";

        var entities = layout.AssociatedBlock.Entities;
        var groups = entities
            .GroupBy(e => e.GetType().Name)
            .Select(g => $"{g.Key}({g.Count()})")
            .ToList();

        var summary = $"Layout '{layoutName}': {entities.Count()} entities — {string.Join(", ", groups)}";

        // Check for TableEntity specifically
        var tables = entities.OfType<TableEntity>().ToList();
        if (tables.Count == 0)
        {
            summary += " | No TableEntity found (ACAD_TABLE). Tables may be custom blocks.";
        }
        else
        {
            foreach (var t in tables)
            {
                summary += $" | Table: {t.Rows.Count}r x {t.Columns.Count}c";
                // Try reading first few cells for diagnostics
                for (int c = 0; c < Math.Min(t.Columns.Count, 9); c++)
                {
                    var cellText = ReadCellText(t, 0, c);
                    summary += $" [col{c}='{(string.IsNullOrEmpty(cellText) ? "(empty)" : cellText)}']";
                }

                // Check cell content structure for first data cell
                if (t.Rows.Count > 0 && t.Columns.Count > 0)
                {
                    try
                    {
                        var firstCell = t.GetCell(0, 0);
                        summary += $" | Cell(0,0): Contents={firstCell.Contents.Count}";
                        if (firstCell.Content != null)
                        {
                            summary += $", ContentType={firstCell.Content.ContentType}";
                            summary += $", ValueType={firstCell.Content.Value.ValueType}";
                            summary += $", Text='{firstCell.Content.Value.Text ?? "(null)"}'";
                            summary += $", Value='{firstCell.Content.Value.Value ?? "(null)"}'";
                        }
                        else
                        {
                            summary += ", Content=null";
                        }
                    }
                    catch (Exception ex)
                    {
                        summary += $" | Cell(0,0) error: {ex.Message}";
                    }
                }
            }
        }

        // Include ACadSharp read warnings (e.g., TABLE entity parse failures)
        if (_readWarnings.Count > 0)
        {
            var tableWarnings = _readWarnings
                .Where(w => w.Contains("TABLE", StringComparison.OrdinalIgnoreCase)
                         || w.Contains("Table", StringComparison.Ordinal))
                .ToList();
            if (tableWarnings.Count > 0)
            {
                summary += $" | TABLE-related warnings ({tableWarnings.Count}): " +
                           string.Join("; ", tableWarnings.Take(3));
            }

            summary += $" | Total read warnings: {_readWarnings.Count}";
        }

        return summary;
    }

    /// <summary>
    /// Returns all ACadSharp read warnings captured during file load.
    /// </summary>
    public IReadOnlyList<string> GetReadWarnings() => _readWarnings;

    // ── Helpers ───────────────────────────────────────────────

    /// <summary>
    /// Logs detailed cell content diagnostics for the first few cells of a table.
    /// This helps identify why ReadCellText returns empty.
    /// </summary>
    private void LogCellDiagnostics(TableEntity table, string layoutName)
    {
        try
        {
            // Check first 3 rows x first 3 columns
            var maxRow = Math.Min(table.Rows.Count, 4);
            var maxCol = Math.Min(table.Columns.Count, 3);

            for (int r = 0; r < maxRow; r++)
            {
                for (int c = 0; c < maxCol; c++)
                {
                    try
                    {
                        var cell = table.GetCell(r, c);
                        var contentsCount = cell.Contents.Count;
                        var content = cell.Content;
                        var readText = ReadCellText(table, r, c);

                        if (content != null)
                        {
                            _logger?.LogInformation(
                                "Cell({Row},{Col}) in '{Layout}': Contents={Count}, Type={ContentType}, " +
                                "ValueType={ValueType}, Text='{Text}', Value='{Value}', " +
                                "FormattedValue='{FmtVal}', ReadCellText='{ReadText}'",
                                r, c, layoutName, contentsCount,
                                content.ContentType, content.Value.ValueType,
                                content.Value.Text ?? "(null)",
                                content.Value.Value ?? "(null)",
                                content.Value.FormattedValue ?? "(null)",
                                readText);
                        }
                        else
                        {
                            _logger?.LogInformation(
                                "Cell({Row},{Col}) in '{Layout}': Contents={Count}, Content=null, ReadCellText='{ReadText}'",
                                r, c, layoutName, contentsCount, readText);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning(
                            "Cell({Row},{Col}) in '{Layout}' diagnostic error: {Error}",
                            r, c, layoutName, ex.Message);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning("Cell diagnostics failed for '{Layout}': {Error}", layoutName, ex.Message);
        }
    }

    private AcLayout? FindLayout(string layoutName)
    {
        return _document?.Layouts?
            .FirstOrDefault(l => string.Equals(l.Name, layoutName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Reads text content from a table cell. Returns empty string on failure.
    /// Tries multiple content paths: Text, Value (object), FormattedValue,
    /// and iterates Contents list if single Content is null.
    /// </summary>
    private static string ReadCellText(TableEntity table, int row, int col)
    {
        try
        {
            if (row < 0 || row >= table.Rows.Count || col < 0 || col >= table.Columns.Count)
                return string.Empty;

            var cell = table.GetCell(row, col);
            if (cell == null)
                return string.Empty;

            // Path 1: Single content (Content = Contents.FirstOrDefault())
            var content = cell.Content;
            if (content != null)
            {
                var result = ExtractTextFromContent(content);
                if (!string.IsNullOrEmpty(result))
                    return result;
            }

            // Path 2: Iterate all Contents (handles multi-content and edge cases)
            foreach (var c in cell.Contents)
            {
                var result = ExtractTextFromContent(c);
                if (!string.IsNullOrEmpty(result))
                    return result;
            }

            return string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Extracts text from a CellContent using all available value paths.
    /// </summary>
    private static string ExtractTextFromContent(TableEntity.CellContent content)
    {
        var value = content.Value;
        if (value == null)
            return string.Empty;

        // Try Text property (set by DXF reader path)
        if (!string.IsNullOrEmpty(value.Text))
            return value.Text;

        // Try Value object (set by DWG reader's readCustomTableDataValue)
        if (value.Value != null)
        {
            var str = value.Value.ToString();
            if (!string.IsNullOrEmpty(str))
                return str;
        }

        // Try FormattedValue (set by DWG reader for R2007+)
        if (!string.IsNullOrEmpty(value.FormattedValue))
            return value.FormattedValue;

        return string.Empty;
    }

    /// <summary>
    /// Strips MText formatting codes. Mirrors the base class StripMTextFormatting.
    /// </summary>
    private static string StripFormatting(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        var result = text;
        // All MText formatting codes — case-insensitive (AutoCAD uses both \F and \f)
        result = Regex.Replace(result, @"\\[Pp]", " ");           // paragraph break
        result = Regex.Replace(result, @"\\[Cc]\d+;", "");        // color
        result = Regex.Replace(result, @"\\[Ff][^;]*;", "");      // font (e.g. \fPMingLiU;)
        result = Regex.Replace(result, @"\\[Hh][^;]*;", "");      // text height
        result = Regex.Replace(result, @"\\[Ss][^;]*;", "");      // stacking/fractions
        result = Regex.Replace(result, @"\\[Qq]\d+;", "");        // oblique angle
        result = Regex.Replace(result, @"\\[Tt]\d+;", "");        // tracking
        result = Regex.Replace(result, @"\\[Ww]\d+\.?\d*;", "");  // width factor
        result = Regex.Replace(result, @"\\[Aa]\d+;", "");        // alignment
        result = Regex.Replace(result, @"\\[LlOo]", "");          // underline/overline toggle
        result = result.Replace("\\~", " ");                       // non-breaking space
        result = result.Replace("{", "").Replace("}", "");         // grouping braces
        result = result.Replace("\\\\", "\\");                     // escaped backslash
        return result.Trim();
    }
}
