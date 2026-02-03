// OdooAutoCAD.Data/Entities/Entities.cs
// Data entities - equivalent to Python SQLAlchemy models in models/server.py

namespace OdooAutoCAD.Data.Entities;

/// <summary>
/// Server configuration cache.
/// Stores Odoo connection settings and other configurations.
/// </summary>
public class ServerConfig
{
    public int Id { get; set; }
    public string ConfigKey { get; set; } = string.Empty;
    public string? ConfigValue { get; set; }
    public string? Description { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Product mapping between AutoCAD names and Odoo products.
/// </summary>
public class ProductMapping
{
    public int Id { get; set; }
    public string AutoCADName { get; set; } = string.Empty;
    public int? OdooProductId { get; set; }
    public string? OdooProductCode { get; set; }
    public string? OdooProductName { get; set; }
    public bool IsActive { get; set; } = true;
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Local cache of BOQ entries.
/// </summary>
public class BOQCache
{
    public int Id { get; set; }
    public int ProjectId { get; set; }
    public string ProjectName { get; set; } = string.Empty;
    public int? OdooBOQId { get; set; }
    public int? ProductId { get; set; }
    public string ProductName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public string UnitOfMeasure { get; set; } = "pcs";
    public decimal? UnitPrice { get; set; }
    public string? Description { get; set; }
    public string? SourceDrawing { get; set; }
    public string? SourceLayout { get; set; }
    public bool IsSynced { get; set; }
    public DateTime? LastSyncTime { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Synchronization log for tracking data sync operations.
/// </summary>
public class SyncLog
{
    public int Id { get; set; }
    public string SyncType { get; set; } = string.Empty; // "boq", "product", "project"
    public string Direction { get; set; } = "upload"; // "upload", "download"
    public int RecordsProcessed { get; set; }
    public int RecordsCreated { get; set; }
    public int RecordsUpdated { get; set; }
    public int RecordsFailed { get; set; }
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string? Details { get; set; } // JSON details
    public DateTime SyncTime { get; set; } = DateTime.UtcNow;
    public TimeSpan Duration { get; set; }
}

/// <summary>
/// User preferences and settings.
/// </summary>
public class UserPreference
{
    public int Id { get; set; }
    public string PreferenceKey { get; set; } = string.Empty;
    public string? PreferenceValue { get; set; }
    public string? DataType { get; set; } // "string", "int", "bool", "json"
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
