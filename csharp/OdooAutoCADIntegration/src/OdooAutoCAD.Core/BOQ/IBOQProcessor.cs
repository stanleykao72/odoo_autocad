// OdooAutoCAD.Core/BOQ/IBOQProcessor.cs
// BOQ Processing Interface - equivalent to Python util_push_to_boq.py

using System.Collections.Generic;
using System.Threading.Tasks;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;

namespace OdooAutoCAD.Core.BOQ;

/// <summary>
/// BOQ generation options.
/// </summary>
public class BOQGenerationOptions
{
    /// <summary>
    /// Project ID to generate BOQ for.
    /// </summary>
    public int ProjectId { get; set; }

    /// <summary>
    /// Whether to include data from AutoCAD.
    /// </summary>
    public bool IncludeAutoCADData { get; set; } = true;

    /// <summary>
    /// Whether to validate products against Odoo.
    /// </summary>
    public bool ValidateProducts { get; set; } = true;

    /// <summary>
    /// Whether to auto-create missing products.
    /// </summary>
    public bool AutoCreateProducts { get; set; } = false;

    /// <summary>
    /// Default unit of measure for entries without one.
    /// </summary>
    public string DefaultUnitOfMeasure { get; set; } = "pcs";
}

/// <summary>
/// Result of BOQ generation.
/// </summary>
public class BOQGenerationResult
{
    public bool Success { get; set; }
    public List<BOQEntry> Entries { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public List<string> Errors { get; set; } = new();
    public int TotalItems { get; set; }
    public int ProcessedItems { get; set; }
    public int SkippedItems { get; set; }
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Validation result for BOQ entries.
/// </summary>
public class BOQValidationResult
{
    public bool IsValid { get; set; }
    public List<BOQValidationError> Errors { get; set; } = new();
}

/// <summary>
/// BOQ validation error details.
/// </summary>
public class BOQValidationError
{
    public int? EntryIndex { get; set; }
    public string Field { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string Severity { get; set; } = "Error"; // Error, Warning
}

/// <summary>
/// Interface for BOQ (Bill of Quantities) processing.
/// Equivalent to Python's util_push_to_boq.py.
/// </summary>
public interface IBOQProcessor
{
    /// <summary>
    /// Generates BOQ from AutoCAD layout data.
    /// </summary>
    /// <param name="layoutData">Layout data extracted from AutoCAD.</param>
    /// <param name="options">Generation options.</param>
    /// <returns>Generation result with BOQ entries.</returns>
    Task<BOQGenerationResult> GenerateBOQAsync(LayoutData layoutData, BOQGenerationOptions options);

    /// <summary>
    /// Generates BOQ for a project, optionally including AutoCAD data.
    /// This is the main MCP tool implementation.
    /// </summary>
    /// <param name="projectId">Project ID.</param>
    /// <param name="includeAutoCADData">Whether to extract and include AutoCAD data.</param>
    /// <returns>Generation result with BOQ entries.</returns>
    Task<BOQGenerationResult> GenerateBOQForProjectAsync(int projectId, bool includeAutoCADData = true);

    /// <summary>
    /// Validates BOQ entries before pushing to Odoo.
    /// </summary>
    /// <param name="entries">Entries to validate.</param>
    /// <returns>Validation result.</returns>
    Task<BOQValidationResult> ValidateBOQAsync(IEnumerable<BOQEntry> entries);

    /// <summary>
    /// Pushes validated BOQ entries to Odoo.
    /// </summary>
    /// <param name="entries">Entries to push.</param>
    /// <returns>Sync result.</returns>
    Task<SyncResult> PushToOdooAsync(IEnumerable<BOQEntry> entries);

    /// <summary>
    /// Maps a product name from AutoCAD to an Odoo product.
    /// </summary>
    /// <param name="autocadProductName">Product name from AutoCAD.</param>
    /// <returns>Matched Odoo product or null.</returns>
    Task<OdooProduct?> MapProductAsync(string autocadProductName);

    /// <summary>
    /// Gets product mapping rules.
    /// </summary>
    IReadOnlyDictionary<string, int> GetProductMappings();

    /// <summary>
    /// Sets a product mapping rule.
    /// </summary>
    /// <param name="autocadName">AutoCAD product name.</param>
    /// <param name="odooProductId">Odoo product ID.</param>
    void SetProductMapping(string autocadName, int odooProductId);

    /// <summary>
    /// Removes a product mapping rule.
    /// </summary>
    /// <param name="autocadName">AutoCAD product name to remove.</param>
    /// <returns>True if the mapping was found and removed.</returns>
    bool RemoveProductMapping(string autocadName);
}
