// OdooAutoCAD.Core/BOQ/BoqImportModels.cs
// Models for import2boq_v2 Swagger API and AutoCAD ID writeback

namespace OdooAutoCAD.Core.BOQ;

/// <summary>
/// One layout's data for the import2boq_v2 API call.
/// Maps to the Python layout_dict structure.
/// </summary>
public class BoqImportLayout
{
    public string LayoutName { get; set; } = string.Empty;
    public string? HeaderId { get; set; }
    public string PrNo { get; set; } = string.Empty;
    public string ProjectName { get; set; } = string.Empty;
    public string JobWorkingPlanName { get; set; } = string.Empty;
    public string ProductName { get; set; } = string.Empty;
    public string ProductCatalog { get; set; } = string.Empty;
    public string Spec { get; set; } = string.Empty;
    public string SurfaceTreatment { get; set; } = string.Empty;
    public string OperationFlow { get; set; } = string.Empty;
    public string ColorName { get; set; } = string.Empty;
    public string ColorNo { get; set; } = string.Empty;
    public List<BoqImportDetail> Detail { get; set; } = new();
}

/// <summary>
/// One table row within a layout for import2boq_v2.
/// </summary>
public class BoqImportDetail
{
    public string Position { get; set; } = string.Empty;
    public string ProductNo { get; set; } = string.Empty;
    public string Width { get; set; } = string.Empty;
    public string Height { get; set; } = string.Empty;
    public string Length { get; set; } = string.Empty;
    public string Thickness { get; set; } = string.Empty;
    public string Qty { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string? DetailId { get; set; }
}

/// <summary>
/// Request payload for import2boq_v2 API.
/// </summary>
public class BoqImportRequest
{
    public List<BoqImportLayout> All { get; set; } = new();
}

/// <summary>
/// Response from import2boq_v2 API.
/// On success, All contains layouts with assigned header_id and detail_id values.
/// On error, ErrorCode and ErrorMessage describe the failure.
/// </summary>
public class BoqImportResponse
{
    public bool Success { get; set; }
    public List<BoqImportLayout>? All { get; set; }
    public string? ErrorCode { get; set; }
    public string? ErrorMessage { get; set; }
}

/// <summary>
/// Data to write back IDs from Odoo response to AutoCAD table cells.
/// One instance per layout.
/// </summary>
public class WritebackLayout
{
    public string LayoutName { get; set; } = string.Empty;
    public string? HeaderId { get; set; }
    public List<WritebackDetail> Details { get; set; } = new();
}

/// <summary>
/// Per-row writeback data: maps product_no to the Odoo-assigned detail_id.
/// </summary>
public class WritebackDetail
{
    public string ProductNo { get; set; } = string.Empty;
    public string? DetailId { get; set; }
}
