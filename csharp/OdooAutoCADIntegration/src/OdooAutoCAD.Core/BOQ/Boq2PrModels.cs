// OdooAutoCAD.Core/BOQ/Boq2PrModels.cs
// Models for boq2pr_v2 Swagger API response

namespace OdooAutoCAD.Core.BOQ;

/// <summary>
/// Response from boq2pr_v2 API.
/// On success, All contains the created PRs with lines.
/// On error, ErrorCode and ErrorMessage describe the failure.
/// </summary>
public class Boq2PrResponse
{
    public bool Success { get; set; }
    public List<Boq2PrResult>? All { get; set; }
    public string? ErrorCode { get; set; }
    public string? ErrorMessage { get; set; }
}

/// <summary>
/// One PR created by boq2pr_v2.
/// </summary>
public class Boq2PrResult
{
    public int? PrId { get; set; }
    public string Reference { get; set; } = string.Empty;
    public string State { get; set; } = "draft";
    public List<Boq2PrLine> Lines { get; set; } = new();
}

/// <summary>
/// One line within a PR created by boq2pr_v2.
/// </summary>
public class Boq2PrLine
{
    public int ProductId { get; set; }
    public string ProductName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public string UnitOfMeasure { get; set; } = string.Empty;
    public decimal? UnitPrice { get; set; }
}
