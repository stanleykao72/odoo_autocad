// OdooAutoCAD.Core/BOQ/BOQProcessor.cs
// BOQ Processing Implementation - equivalent to Python util_push_to_boq.py

using Microsoft.Extensions.Logging;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.Core.BOQ;

/// <summary>
/// BOQ (Bill of Quantities) Processor implementation.
/// Handles generation and validation of BOQ entries from AutoCAD data.
/// </summary>
public class BOQProcessor : IBOQProcessor
{
    private readonly ILogger<BOQProcessor>? _logger;
    private readonly IGUIProxy _guiProxy;
    private readonly IAutoCADService _autoCADService;
    private readonly IOdooService _odooService;
    private readonly Dictionary<string, int> _productMappings = new(StringComparer.OrdinalIgnoreCase);

    public BOQProcessor(
        IGUIProxy guiProxy,
        IAutoCADService autoCADService,
        IOdooService odooService,
        ILogger<BOQProcessor>? logger = null)
    {
        _guiProxy = guiProxy;
        _autoCADService = autoCADService;
        _odooService = odooService;
        _logger = logger;
    }

    public async Task<BOQGenerationResult> GenerateBOQAsync(LayoutData layoutData, BOQGenerationOptions options)
    {
        var result = new BOQGenerationResult();

        try
        {
            _logger?.LogInformation("Starting BOQ generation for project {ProjectId}", options.ProjectId);

            // Extract items from layout data
            var extractedItems = ExtractItemsFromLayout(layoutData);
            result.TotalItems = extractedItems.Count;

            foreach (var item in extractedItems)
            {
                try
                {
                    var entry = await ProcessItemAsync(item, options);
                    if (entry != null)
                    {
                        result.Entries.Add(entry);
                        result.ProcessedItems++;
                    }
                    else
                    {
                        result.SkippedItems++;
                        result.Warnings.Add($"Skipped item: {item.Name}");
                    }
                }
                catch (Exception ex)
                {
                    result.SkippedItems++;
                    result.Errors.Add($"Error processing {item.Name}: {ex.Message}");
                    _logger?.LogWarning(ex, "Error processing BOQ item: {ItemName}", item.Name);
                }
            }

            result.Success = result.Errors.Count == 0;
            _logger?.LogInformation("BOQ generation completed: {Processed}/{Total} items processed",
                result.ProcessedItems, result.TotalItems);

            return result;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "BOQ generation failed");
            result.Success = false;
            result.Errors.Add($"BOQ generation failed: {ex.Message}");
            return result;
        }
    }

    public async Task<BOQGenerationResult> GenerateBOQForProjectAsync(int projectId, bool includeAutoCADData = true)
    {
        var options = new BOQGenerationOptions
        {
            ProjectId = projectId,
            IncludeAutoCADData = includeAutoCADData,
            ValidateProducts = true
        };

        LayoutData? layoutData = null;

        if (includeAutoCADData)
        {
            // Execute AutoCAD extraction via GUI proxy for thread safety
            var response = await _guiProxy.ExecuteInGuiAsync("extract_autocad_parameters");

            if (response.Success && response.Result is LayoutData data)
            {
                layoutData = data;
            }
            else
            {
                _logger?.LogWarning("Failed to extract AutoCAD data: {Error}", response.ErrorMessage);
                layoutData = new LayoutData();
            }
        }
        else
        {
            layoutData = new LayoutData();
        }

        return await GenerateBOQAsync(layoutData, options);
    }

    public async Task<BOQValidationResult> ValidateBOQAsync(IEnumerable<BOQEntry> entries)
    {
        var result = new BOQValidationResult { IsValid = true };
        var entryList = entries.ToList();

        for (int i = 0; i < entryList.Count; i++)
        {
            var entry = entryList[i];

            // Validate required fields
            if (entry.ProjectId <= 0)
            {
                result.Errors.Add(new BOQValidationError
                {
                    EntryIndex = i,
                    Field = "ProjectId",
                    Message = "Project ID is required",
                    Severity = "Error"
                });
            }

            if (entry.ProductId <= 0)
            {
                result.Errors.Add(new BOQValidationError
                {
                    EntryIndex = i,
                    Field = "ProductId",
                    Message = "Product ID is required",
                    Severity = "Error"
                });
            }

            if (entry.Quantity <= 0)
            {
                result.Errors.Add(new BOQValidationError
                {
                    EntryIndex = i,
                    Field = "Quantity",
                    Message = "Quantity must be greater than zero",
                    Severity = "Error"
                });
            }

            // Validate product exists in Odoo
            if (entry.ProductId > 0)
            {
                var product = await _odooService.GetProductAsync(entry.ProductId);
                if (product == null)
                {
                    result.Errors.Add(new BOQValidationError
                    {
                        EntryIndex = i,
                        Field = "ProductId",
                        Message = $"Product with ID {entry.ProductId} not found in Odoo",
                        Severity = "Error"
                    });
                }
            }
        }

        result.IsValid = !result.Errors.Any(e => e.Severity == "Error");

        return result;
    }

    public async Task<SyncResult> PushToOdooAsync(IEnumerable<BOQEntry> entries)
    {
        // Validate before pushing
        var validation = await ValidateBOQAsync(entries);
        if (!validation.IsValid)
        {
            var errors = validation.Errors
                .Where(e => e.Severity == "Error")
                .Select(e => e.Message)
                .ToList();

            return new SyncResult(
                Success: false,
                RecordsProcessed: 0,
                RecordsCreated: 0,
                RecordsUpdated: 0,
                RecordsFailed: validation.Errors.Count,
                Errors: errors);
        }

        return await _odooService.ImportToBOQAsync(entries);
    }

    public async Task<OdooProduct?> MapProductAsync(string autocadProductName)
    {
        // Check local mapping first
        if (_productMappings.TryGetValue(autocadProductName, out var productId))
        {
            return await _odooService.GetProductAsync(productId);
        }

        // Try to find by name/code in Odoo
        var products = await _odooService.SearchProductsAsync(autocadProductName);
        if (products.Count > 0)
        {
            // Cache the mapping
            _productMappings[autocadProductName] = products[0].Id;
            return products[0];
        }

        return null;
    }

    public IReadOnlyDictionary<string, int> GetProductMappings()
    {
        return _productMappings;
    }

    public void SetProductMapping(string autocadName, int odooProductId)
    {
        _productMappings[autocadName] = odooProductId;
        _logger?.LogDebug("Product mapping set: {AutoCADName} -> {OdooProductId}", autocadName, odooProductId);
    }

    #region Private Helper Methods

    private List<ExtractedItem> ExtractItemsFromLayout(LayoutData layoutData)
    {
        var items = new List<ExtractedItem>();

        // Extract from parameters — AutoCAD block attributes have paired keys:
        // "product_name" + "quantity", or keys containing qty/quantity/count
        var parameters = layoutData.Parameters;
        string? productName = null;
        string? unit = null;

        // Look for explicit product_name + quantity pair first
        if (parameters.TryGetValue("product_name", out var pn))
            productName = pn?.ToString();
        if (parameters.TryGetValue("unit", out var u))
            unit = u?.ToString();

        if (!string.IsNullOrWhiteSpace(productName))
        {
            // Find paired quantity
            decimal qty = 0;
            if (parameters.TryGetValue("quantity", out var qv))
                qty = ParseQuantity(qv);
            else if (parameters.TryGetValue("qty", out var qv2))
                qty = ParseQuantity(qv2);

            if (qty > 0)
            {
                items.Add(new ExtractedItem
                {
                    Name = productName,
                    Quantity = qty,
                    Unit = unit,
                    Source = "Parameter"
                });
            }
        }
        else
        {
            // Fallback: scan for keys with quantity-like names
            foreach (var kvp in parameters)
            {
                if (IsQuantityParameter(kvp.Key))
                {
                    var name = ExtractProductName(kvp.Key);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        items.Add(new ExtractedItem
                        {
                            Name = name,
                            Quantity = ParseQuantity(kvp.Value),
                            Source = "Parameter"
                        });
                    }
                }
            }
        }

        // Extract from tables
        foreach (var table in layoutData.Tables)
        {
            var tableItems = ExtractItemsFromTable(table);
            items.AddRange(tableItems);
        }

        return items;
    }

    private List<ExtractedItem> ExtractItemsFromTable(TableData table)
    {
        var items = new List<ExtractedItem>();
        if (table.Cells.Count < 2) return items; // need header + at least one data row

        // Detect column indices from header row
        var headers = table.Cells[0];
        int nameCol = -1, qtyCol = -1, descCol = -1, unitCol = -1;

        for (int c = 0; c < headers.Count; c++)
        {
            var h = headers[c].Trim().ToLowerInvariant();
            if (nameCol < 0 && (h.Contains("product") || h == "name" || h == "item"))
                nameCol = c;
            else if (qtyCol < 0 && (h == "qty" || h.Contains("quantity") || h.Contains("count")))
                qtyCol = c;
            else if (descCol < 0 && (h.Contains("description") || h.Contains("desc")))
                descCol = c;
            else if (unitCol < 0 && (h.Contains("unit") || h == "uom"))
                unitCol = c;
        }

        // Fallback: if no header match, assume col 0 = name, col 1 = quantity
        if (nameCol < 0) nameCol = 0;
        if (qtyCol < 0) qtyCol = Math.Min(1, headers.Count - 1);

        // Extract data rows (skip header)
        for (int row = 1; row < table.Cells.Count; row++)
        {
            var cells = table.Cells[row];
            if (cells.Count <= Math.Max(nameCol, qtyCol)) continue;

            var name = cells[nameCol];
            var quantityStr = cells[qtyCol];

            if (!string.IsNullOrWhiteSpace(name) && decimal.TryParse(quantityStr, out var quantity) && quantity > 0)
            {
                items.Add(new ExtractedItem
                {
                    Name = name,
                    Quantity = quantity,
                    Source = "Table",
                    Unit = unitCol >= 0 && unitCol < cells.Count && !string.IsNullOrWhiteSpace(cells[unitCol])
                        ? cells[unitCol]
                        : "pcs"
                });
            }
        }

        return items;
    }

    private async Task<BOQEntry?> ProcessItemAsync(ExtractedItem item, BOQGenerationOptions options)
    {
        // Try to map to Odoo product
        var product = await MapProductAsync(item.Name);

        if (product == null && options.ValidateProducts)
        {
            _logger?.LogWarning("No product mapping found for: {ItemName}", item.Name);
            return null;
        }

        return new BOQEntry
        {
            ProjectId = options.ProjectId,
            ProductId = product?.Id ?? 0,
            ProductName = product?.Name ?? item.Name,
            Quantity = item.Quantity,
            UnitOfMeasure = item.Unit ?? options.DefaultUnitOfMeasure,
            UnitPrice = product?.ListPrice,
            Description = $"Extracted from AutoCAD ({item.Source})",
            Reference = item.Name
        };
    }

    private static bool IsQuantityParameter(string key)
    {
        var quantityKeywords = new[] { "qty", "quantity", "count", "amount", "num" };
        return quantityKeywords.Any(k => key.ToLowerInvariant().Contains(k));
    }

    private static string ExtractProductName(string parameterKey)
    {
        // Remove quantity-related suffixes
        var name = parameterKey
            .Replace("_qty", "")
            .Replace("_quantity", "")
            .Replace("_count", "")
            .Replace("Qty", "")
            .Replace("Quantity", "")
            .Replace("Count", "");

        return name.Trim('_', '-', ' ');
    }

    private static decimal ParseQuantity(object value)
    {
        return value switch
        {
            decimal d => d,
            double d => (decimal)d,
            int i => i,
            string s when decimal.TryParse(s, out var result) => result,
            _ => 0
        };
    }

    #endregion

    /// <summary>
    /// Internal class for extracted items before mapping.
    /// </summary>
    private class ExtractedItem
    {
        public string Name { get; set; } = string.Empty;
        public decimal Quantity { get; set; }
        public string? Unit { get; set; }
        public string Source { get; set; } = string.Empty;
    }
}
