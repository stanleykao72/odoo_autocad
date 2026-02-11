// OdooAutoCAD.Core/Odoo/IOdooService.cs
// Odoo REST API Client Interface - equivalent to Python util_odoo.py

using System.Collections.Generic;
using System.Threading.Tasks;
using OdooAutoCAD.Core.BOQ;

namespace OdooAutoCAD.Core.Odoo;

/// <summary>
/// Odoo connection status information.
/// </summary>
public record OdooStatus(
    bool IsConnected,
    string? ServerUrl,
    string? Database,
    string? Username,
    string? Version,
    string? ErrorMessage);

/// <summary>
/// Odoo project information.
/// </summary>
public record OdooProject(
    int Id,
    string Name,
    string? Code,
    string State,
    DateTime? DateStart,
    DateTime? DateEnd);

/// <summary>
/// Odoo product information.
/// </summary>
public record OdooProduct(
    int Id,
    string Name,
    string? Code,
    string? Description,
    decimal? ListPrice,
    string? UnitOfMeasure,
    string? Category);

/// <summary>
/// BOQ (Bill of Quantities) entry.
/// </summary>
public class BOQEntry
{
    public int? Id { get; set; }
    public int ProjectId { get; set; }
    public int ProductId { get; set; }
    public string ProductName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public string UnitOfMeasure { get; set; } = string.Empty;
    public decimal? UnitPrice { get; set; }
    public string? Description { get; set; }
    public string? Reference { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Purchase Requisition entry.
/// </summary>
public class PREntry
{
    public int? Id { get; set; }
    public int ProjectId { get; set; }
    public string Reference { get; set; } = string.Empty;
    public List<PRLine> Lines { get; set; } = new();
    public string State { get; set; } = "draft";
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Purchase Requisition line item.
/// </summary>
public class PRLine
{
    public int ProductId { get; set; }
    public string ProductName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public string UnitOfMeasure { get; set; } = string.Empty;
    public decimal? UnitPrice { get; set; }
}

/// <summary>
/// Sync result for data operations.
/// </summary>
public record SyncResult(
    bool Success,
    int RecordsProcessed,
    int RecordsCreated,
    int RecordsUpdated,
    int RecordsFailed,
    List<string>? Errors);

/// <summary>
/// Interface for Odoo REST API operations.
/// Equivalent to Python's util_odoo.py.
/// </summary>
public interface IOdooService
{
    #region Connection Management

    /// <summary>
    /// Raised when the connection state changes (true = connected, false = disconnected).
    /// </summary>
    event EventHandler<bool>? ConnectionStateChanged;

    /// <summary>
    /// Gets whether Odoo is currently connected (via JSON-RPC session or API token).
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// Gets whether Odoo has been authenticated via Swagger/OpenAPI (Basic Auth).
    /// This is separate from the JSON-RPC session-based connection.
    /// </summary>
    bool IsApiAuthenticated { get; }

    /// <summary>
    /// Marks Odoo as authenticated via Swagger/OpenAPI Basic Auth.
    /// Used by Dashboard and Odoo page after successful Swagger + Basic Auth test.
    /// </summary>
    void MarkApiAuthenticated(string serverUrl, string database);

    /// <summary>
    /// Clears API authentication state and fires ConnectionStateChanged(false).
    /// </summary>
    void ClearApiAuthentication();

    /// <summary>
    /// Connects to Odoo server with the specified credentials.
    /// </summary>
    /// <param name="serverUrl">Odoo server URL.</param>
    /// <param name="database">Database name.</param>
    /// <param name="username">Username.</param>
    /// <param name="password">Password or API key.</param>
    /// <returns>True if connection successful.</returns>
    Task<bool> ConnectAsync(string serverUrl, string database, string username, string password);

    /// <summary>
    /// Disconnects from Odoo.
    /// </summary>
    Task DisconnectAsync();

    /// <summary>
    /// Gets detailed status information about the Odoo connection.
    /// </summary>
    Task<OdooStatus> GetStatusAsync();

    /// <summary>
    /// Tests the connection to Odoo.
    /// </summary>
    /// <returns>True if connection is valid.</returns>
    Task<bool> TestConnectionAsync();

    /// <summary>
    /// Tests the connection to Odoo with specific credentials (non-persistent).
    /// </summary>
    /// <param name="serverUrl">Server URL to test.</param>
    /// <param name="database">Database name.</param>
    /// <param name="username">Username.</param>
    /// <param name="password">Password or API key.</param>
    /// <returns>OdooStatus with detailed connection information.</returns>
    Task<OdooStatus> TestConnectionAsync(string serverUrl, string database, string username, string password);

    #endregion

    #region Project Operations

    /// <summary>
    /// Gets all projects.
    /// </summary>
    /// <param name="activeOnly">If true, only returns active projects.</param>
    Task<IReadOnlyList<OdooProject>> GetProjectsAsync(bool activeOnly = true);

    /// <summary>
    /// Gets a project by ID.
    /// </summary>
    /// <param name="projectId">Project ID.</param>
    Task<OdooProject?> GetProjectAsync(int projectId);

    /// <summary>
    /// Searches for projects by name or code.
    /// </summary>
    /// <param name="searchTerm">Search term.</param>
    Task<IReadOnlyList<OdooProject>> SearchProjectsAsync(string searchTerm);

    #endregion

    #region Product Operations

    /// <summary>
    /// Gets all products.
    /// </summary>
    Task<IReadOnlyList<OdooProduct>> GetProductsAsync();

    /// <summary>
    /// Gets products via Swagger API with Basic Auth (no JSON-RPC session required).
    /// </summary>
    Task<IReadOnlyList<OdooProduct>> GetProductsViaApiAsync(
        string baseUrl, string basePath, string database, string userToken);

    /// <summary>
    /// Gets a product by ID.
    /// </summary>
    /// <param name="productId">Product ID.</param>
    Task<OdooProduct?> GetProductAsync(int productId);

    /// <summary>
    /// Searches for products by name or code.
    /// </summary>
    /// <param name="searchTerm">Search term.</param>
    Task<IReadOnlyList<OdooProduct>> SearchProductsAsync(string searchTerm);

    /// <summary>
    /// Gets products by category.
    /// </summary>
    /// <param name="categoryId">Category ID.</param>
    Task<IReadOnlyList<OdooProduct>> GetProductsByCategoryAsync(int categoryId);

    #endregion

    #region BOQ Operations

    /// <summary>
    /// Imports data to BOQ (Bill of Quantities).
    /// Equivalent to Python's import2boq method.
    /// </summary>
    /// <param name="entries">BOQ entries to import.</param>
    /// <returns>Sync result with details.</returns>
    Task<SyncResult> ImportToBOQAsync(IEnumerable<BOQEntry> entries);

    /// <summary>
    /// Imports BOQ data via the Swagger/BasicAuth import2boq_v2 API.
    /// Posts layout data and returns assigned header_id/detail_id values.
    /// </summary>
    Task<BoqImportResponse> ImportToBOQViaApiAsync(
        BoqImportRequest request, string baseUrl, string basePath,
        string database, string userToken);

    /// <summary>
    /// Gets BOQ entries for a project.
    /// </summary>
    /// <param name="projectId">Project ID.</param>
    Task<IReadOnlyList<BOQEntry>> GetBOQEntriesAsync(int projectId);

    /// <summary>
    /// Updates a BOQ entry.
    /// </summary>
    /// <param name="entry">Entry to update.</param>
    /// <returns>True if successful.</returns>
    Task<bool> UpdateBOQEntryAsync(BOQEntry entry);

    /// <summary>
    /// Deletes a BOQ entry.
    /// </summary>
    /// <param name="entryId">Entry ID to delete.</param>
    /// <returns>True if successful.</returns>
    Task<bool> DeleteBOQEntryAsync(int entryId);

    #endregion

    #region Purchase Requisition Operations

    /// <summary>
    /// Converts BOQ to Purchase Requisition.
    /// Equivalent to Python's boq2pr method.
    /// </summary>
    /// <param name="projectId">Project ID.</param>
    /// <param name="boqEntryIds">Optional list of specific BOQ entry IDs to convert.</param>
    /// <returns>Created PR entry.</returns>
    Task<PREntry?> ConvertBOQToPRAsync(int projectId, IEnumerable<int>? boqEntryIds = null);

    /// <summary>
    /// Gets Purchase Requisitions for a project.
    /// </summary>
    /// <param name="projectId">Project ID.</param>
    Task<IReadOnlyList<PREntry>> GetPurchaseRequisitionsAsync(int projectId);

    /// <summary>
    /// Submits a Purchase Requisition for approval.
    /// </summary>
    /// <param name="prId">PR ID to submit.</param>
    /// <returns>True if successful.</returns>
    Task<bool> SubmitPRAsync(int prId);

    #endregion

    #region Synchronization

    /// <summary>
    /// Syncs data to Odoo.
    /// Supports different sync types: parameters, boq, project.
    /// </summary>
    /// <param name="data">Data to sync.</param>
    /// <param name="syncType">Type of sync operation.</param>
    /// <returns>Sync result with details.</returns>
    Task<SyncResult> SyncToOdooAsync(object data, string syncType);

    /// <summary>
    /// Gets the last sync timestamp.
    /// </summary>
    DateTime? GetLastSyncTime();

    #endregion
}
