// OdooAutoCAD.Core/Odoo/OdooService.cs
// Odoo REST API Service Implementation - equivalent to Python util_odoo.py

using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.Core.BOQ;

namespace OdooAutoCAD.Core.Odoo;

/// <summary>
/// Odoo REST API Service implementation.
/// Uses HttpClient for REST API communication.
/// </summary>
public class OdooService : IOdooService, IDisposable
{
    private readonly ILogger<OdooService>? _logger;
    private readonly HttpClient _httpClient;

    private string? _serverUrl;
    private string? _database;
    private string? _username;
    private string? _sessionId;
    private int? _userId;
    private DateTime? _lastSyncTime;
    private bool _isConnected;
    private volatile bool _isApiAuthenticated;

    public OdooService(ILogger<OdooService>? logger = null, int timeoutSeconds = 30)
    {
        _logger = logger;
        _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(Math.Clamp(timeoutSeconds, 5, 120))
        };
        _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
    }

    public event EventHandler<bool>? ConnectionStateChanged;

    public bool IsConnected => (_isConnected && !string.IsNullOrEmpty(_sessionId)) || _isApiAuthenticated;

    public bool IsApiAuthenticated => _isApiAuthenticated;

    public void MarkApiAuthenticated(string serverUrl, string database)
    {
        _serverUrl = serverUrl;
        _database = database;
        _isApiAuthenticated = true;
        _logger?.LogInformation("Odoo API authenticated via Swagger: {ServerUrl}, Database: {Database}",
            serverUrl, database);
        ConnectionStateChanged?.Invoke(this, true);
    }

    public void ClearApiAuthentication()
    {
        _isApiAuthenticated = false;
        _logger?.LogInformation("Odoo API authentication cleared");
        ConnectionStateChanged?.Invoke(this, false);
    }

    #region Connection Management

    public async Task<bool> ConnectAsync(string serverUrl, string database, string username, string password)
    {
        try
        {
            _serverUrl = serverUrl.TrimEnd('/');
            _database = database;
            _username = username;

            // Authenticate using JSON-RPC
            var authRequest = new
            {
                jsonrpc = "2.0",
                method = "call",
                @params = new
                {
                    db = database,
                    login = username,
                    password = password
                },
                id = 1
            };

            var response = await PostJsonRpcAsync("/web/session/authenticate", authRequest);

            if (response?.TryGetProperty("result", out var result) == true)
            {
                if (result.TryGetProperty("uid", out var uid) && uid.ValueKind != JsonValueKind.False)
                {
                    _userId = uid.GetInt32();
                    _isConnected = true;
                    _logger?.LogInformation("Connected to Odoo: {ServerUrl}, Database: {Database}, User: {Username}",
                        _serverUrl, _database, _username);
                    ConnectionStateChanged?.Invoke(this, true);
                    return true;
                }
            }

            _logger?.LogWarning("Authentication failed for user: {Username}", username);
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to connect to Odoo");
            _isConnected = false;
            return false;
        }
    }

    public async Task DisconnectAsync()
    {
        _sessionId = null;
        _userId = null;
        _isConnected = false;
        _isApiAuthenticated = false;
        _logger?.LogInformation("Disconnected from Odoo");
        ConnectionStateChanged?.Invoke(this, false);
        await Task.CompletedTask;
    }

    public async Task<OdooStatus> GetStatusAsync()
    {
        if (!IsConnected)
        {
            return new OdooStatus(
                IsConnected: false,
                ServerUrl: _serverUrl,
                Database: _database,
                Username: _username,
                Version: null,
                ErrorMessage: "Not connected");
        }

        try
        {
            var versionRequest = new
            {
                jsonrpc = "2.0",
                method = "call",
                @params = new { },
                id = 1
            };

            var response = await PostJsonRpcAsync("/web/webclient/version_info", versionRequest);
            string? version = null;

            if (response?.TryGetProperty("result", out var result) == true)
            {
                if (result.TryGetProperty("server_version", out var serverVersion))
                {
                    version = serverVersion.GetString();
                }
            }

            return new OdooStatus(
                IsConnected: true,
                ServerUrl: _serverUrl,
                Database: _database,
                Username: _username,
                Version: version,
                ErrorMessage: null);
        }
        catch (Exception ex)
        {
            return new OdooStatus(
                IsConnected: false,
                ServerUrl: _serverUrl,
                Database: _database,
                Username: _username,
                Version: null,
                ErrorMessage: ex.Message);
        }
    }

    public async Task<bool> TestConnectionAsync()
    {
        try
        {
            var status = await GetStatusAsync();
            return status.IsConnected;
        }
        catch
        {
            return false;
        }
    }

    public async Task<OdooStatus> TestConnectionAsync(string serverUrl, string database, string username, string password)
    {
        try
        {
            // Create temporary HttpClient for non-persistent test
            using var httpClient = new HttpClient { Timeout = _httpClient.Timeout };
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            var authRequest = new
            {
                jsonrpc = "2.0",
                method = "call",
                @params = new
                {
                    db = database,
                    login = username,
                    password = password
                },
                id = 1
            };

            var json = JsonSerializer.Serialize(authRequest);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await httpClient.PostAsync($"{serverUrl.TrimEnd('/')}/web/session/authenticate", content);

            // Handle different failure scenarios with specific error messages
            if (!response.IsSuccessStatusCode)
            {
                return response.StatusCode switch
                {
                    System.Net.HttpStatusCode.Unauthorized or 
                    System.Net.HttpStatusCode.Forbidden =>
                        new OdooStatus(false, serverUrl, database, username, null, 
                            "Authentication failed. Please check username and password"),
                    System.Net.HttpStatusCode.NotFound =>
                        new OdooStatus(false, serverUrl, database, username, null, 
                            "Odoo not found at this URL"),
                    System.Net.HttpStatusCode.InternalServerError =>
                        new OdooStatus(false, serverUrl, database, username, null, 
                            "Server error occurred"),
                    _ =>
                        new OdooStatus(false, serverUrl, database, username, null, 
                            $"HTTP {(int)response.StatusCode}: {response.ReasonPhrase}")
                };
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonSerializer.Deserialize<JsonElement>(responseJson);

            // Check for authentication success
            if (jsonDoc.TryGetProperty("result", out var result))
            {
                if (result.TryGetProperty("uid", out var uid) && uid.ValueKind != JsonValueKind.False)
                {
                    // Get version info
                    var versionRequest = new
                    {
                        jsonrpc = "2.0",
                        method = "call",
                        @params = new { },
                        id = 2
                    };

                    var versionJson = JsonSerializer.Serialize(versionRequest);
                    var versionContent = new StringContent(versionJson, Encoding.UTF8, "application/json");
                    var versionResponse = await httpClient.PostAsync($"{serverUrl.TrimEnd('/')}/web/webclient/version_info", versionContent);

                    string? version = null;
                    if (versionResponse.IsSuccessStatusCode)
                    {
                        var versionResponseJson = await versionResponse.Content.ReadAsStringAsync();
                        var versionDoc = JsonSerializer.Deserialize<JsonElement>(versionResponseJson);
                        if (versionDoc.TryGetProperty("result", out var versionResult) &&
                            versionResult.TryGetProperty("server_version", out var serverVersion))
                        {
                            version = serverVersion.GetString();
                        }
                    }

                    return new OdooStatus(true, serverUrl, database, username, version, null);
                }
                else if (result.TryGetProperty("db", out var dbArray) && dbArray.ValueKind == JsonValueKind.Array)
                {
                    return new OdooStatus(false, serverUrl, database, username, null, 
                        $"Database '{database}' not found on server");
                }
            }

            return new OdooStatus(false, serverUrl, database, username, null, 
                "Authentication failed. Please check username and password");
        }
        catch (TaskCanceledException)
        {
            return new OdooStatus(false, serverUrl, database, username, null, 
                $"Connection timed out after {_httpClient.Timeout.TotalSeconds} seconds");
        }
        catch (HttpRequestException ex)
        {
            if (ex.InnerException is System.Security.Authentication.AuthenticationException)
            {
                return new OdooStatus(false, serverUrl, database, username, null, 
                    "SSL certificate validation failed");
            }
            return new OdooStatus(false, serverUrl, database, username, null, 
                $"Unable to reach Odoo server at {serverUrl}");
        }
        catch (JsonException)
        {
            return new OdooStatus(false, serverUrl, database, username, null, 
                "Invalid response from server");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Unexpected error testing Odoo connection");
            return new OdooStatus(false, serverUrl, database, username, null, 
                $"Unexpected error: {ex.Message}");
        }
    }

    #endregion

    #region Project Operations

    public async Task<IReadOnlyList<OdooProject>> GetProjectsAsync(bool activeOnly = true)
    {
        var domain = activeOnly
            ? new object[] { new object[] { "active", "=", true } }
            : Array.Empty<object>();

        var result = await SearchReadAsync("project.project", domain,
            new[] { "id", "name", "code", "state", "date_start", "date" });

        return result.Select(r => MapToOdooProject(r)).ToList();
    }

    public async Task<OdooProject?> GetProjectAsync(int projectId)
    {
        var result = await ReadAsync("project.project", new[] { projectId },
            new[] { "id", "name", "code", "state", "date_start", "date" });

        return result.FirstOrDefault() is { } data ? MapToOdooProject(data) : null;
    }

    public async Task<IReadOnlyList<OdooProject>> SearchProjectsAsync(string searchTerm)
    {
        var domain = new object[]
        {
            "|",
            new object[] { "name", "ilike", searchTerm },
            new object[] { "code", "ilike", searchTerm }
        };

        var result = await SearchReadAsync("project.project", domain,
            new[] { "id", "name", "code", "state", "date_start", "date" });

        return result.Select(r => MapToOdooProject(r)).ToList();
    }

    #endregion

    #region Product Operations

    public async Task<IReadOnlyList<OdooProduct>> GetProductsAsync()
    {
        var result = await SearchReadAsync("product.product", Array.Empty<object>(),
            new[] { "id", "name", "default_code", "description", "list_price", "uom_id", "categ_id" });

        return result.Select(r => MapToOdooProduct(r)).ToList();
    }

    public async Task<OdooProduct?> GetProductAsync(int productId)
    {
        var result = await ReadAsync("product.product", new[] { productId },
            new[] { "id", "name", "default_code", "description", "list_price", "uom_id", "categ_id" });

        return result.FirstOrDefault() is { } data ? MapToOdooProduct(data) : null;
    }

    public async Task<IReadOnlyList<OdooProduct>> SearchProductsAsync(string searchTerm)
    {
        var domain = new object[]
        {
            "|",
            new object[] { "name", "ilike", searchTerm },
            new object[] { "default_code", "ilike", searchTerm }
        };

        var result = await SearchReadAsync("product.product", domain,
            new[] { "id", "name", "default_code", "description", "list_price", "uom_id", "categ_id" });

        return result.Select(r => MapToOdooProduct(r)).ToList();
    }

    public async Task<IReadOnlyList<OdooProduct>> GetProductsByCategoryAsync(int categoryId)
    {
        var domain = new object[] { new object[] { "categ_id", "=", categoryId } };

        var result = await SearchReadAsync("product.product", domain,
            new[] { "id", "name", "default_code", "description", "list_price", "uom_id", "categ_id" });

        return result.Select(r => MapToOdooProduct(r)).ToList();
    }

    public async Task<IReadOnlyList<OdooProductCategory>> GetCategoriesAsync()
    {
        var result = await SearchReadAsync("product.category", Array.Empty<object>(),
            new[] { "id", "name", "parent_id" });

        return result.Select(r => new OdooProductCategory(
            r.GetProperty("id").GetInt32(),
            r.GetProperty("name").GetString() ?? "",
            r.TryGetProperty("parent_id", out var p) && p.ValueKind == JsonValueKind.Array
                ? p[0].GetInt32() : null
        )).ToList();
    }

    public async Task<IReadOnlyList<OdooProduct>> GetProductsViaApiAsync(
        string baseUrl, string endpointPath, string database, string userToken)
    {
        try
        {
            var credentials = Convert.ToBase64String(
                Encoding.UTF8.GetBytes($"{database}:{userToken}"));

            var resolvedPath = endpointPath.Replace("{method_name}", "get_product_list");
            var apiEndpoint = $"{baseUrl.TrimEnd('/')}{resolvedPath}";

            var request = new HttpRequestMessage(HttpMethod.Post, apiEndpoint);
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new { kwargs = new { user_token = userToken }, context = new { } }),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonSerializer.Deserialize<JsonElement>(responseJson);

            var products = new List<OdooProduct>();

            // Parse response — expect array of product objects
            if (jsonDoc.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in jsonDoc.EnumerateArray())
                {
                    products.Add(ParseApiProduct(item));
                }
            }
            else if (jsonDoc.TryGetProperty("result", out var resultArray) &&
                     resultArray.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in resultArray.EnumerateArray())
                {
                    products.Add(ParseApiProduct(item));
                }
            }

            _logger?.LogInformation("Fetched {Count} products via API", products.Count);
            return products;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to fetch products via API");
            throw;
        }
    }

    private static OdooProduct ParseApiProduct(JsonElement item)
    {
        return new OdooProduct(
            Id: item.TryGetProperty("id", out var id) && id.ValueKind == JsonValueKind.Number
                ? id.GetInt32() : 0,
            Name: item.TryGetProperty("name", out var name) ? name.GetString() ?? "" : "",
            Code: item.TryGetProperty("default_code", out var code) && code.ValueKind == JsonValueKind.String
                ? code.GetString() : null,
            Description: item.TryGetProperty("description", out var desc) && desc.ValueKind == JsonValueKind.String
                ? desc.GetString() : null,
            ListPrice: item.TryGetProperty("list_price", out var price) && price.ValueKind == JsonValueKind.Number
                ? price.GetDecimal() : null,
            UnitOfMeasure: item.TryGetProperty("uom_id", out var uom)
                ? (uom.ValueKind == JsonValueKind.Array ? uom[1].GetString()
                    : uom.ValueKind == JsonValueKind.String ? uom.GetString() : null)
                : null,
            Category: item.TryGetProperty("categ_id", out var cat)
                ? (cat.ValueKind == JsonValueKind.Array ? cat[1].GetString()
                    : cat.ValueKind == JsonValueKind.String ? cat.GetString() : null)
                : null);
    }

    #endregion

    #region BOQ Operations

    public async Task<SyncResult> ImportToBOQAsync(IEnumerable<BOQEntry> entries)
    {
        var entryList = entries.ToList();
        int created = 0, updated = 0, failed = 0;
        var errors = new List<string>();

        foreach (var entry in entryList)
        {
            try
            {
                var values = new Dictionary<string, object>
                {
                    ["project_id"] = entry.ProjectId,
                    ["product_id"] = entry.ProductId,
                    ["quantity"] = entry.Quantity,
                    ["description"] = entry.Description ?? ""
                };

                if (entry.Id.HasValue)
                {
                    // Update existing
                    await WriteAsync("boq.line", new[] { entry.Id.Value }, values);
                    updated++;
                }
                else
                {
                    // Create new
                    await CreateAsync("boq.line", values);
                    created++;
                }
            }
            catch (Exception ex)
            {
                failed++;
                errors.Add($"Failed to process entry {entry.ProductName}: {ex.Message}");
                _logger?.LogWarning(ex, "Failed to import BOQ entry: {ProductName}", entry.ProductName);
            }
        }

        _lastSyncTime = DateTime.UtcNow;

        return new SyncResult(
            Success: failed == 0,
            RecordsProcessed: entryList.Count,
            RecordsCreated: created,
            RecordsUpdated: updated,
            RecordsFailed: failed,
            Errors: errors.Count > 0 ? errors : null);
    }

    public async Task<BoqImportResponse> ImportToBOQViaApiAsync(
        BoqImportRequest request, string baseUrl, string endpointPath,
        string database, string userToken)
    {
        try
        {
            var credentials = Convert.ToBase64String(
                Encoding.UTF8.GetBytes($"{database}:{userToken}"));

            var resolvedPath = endpointPath.Replace("{method_name}", "import2boq_v2");
            var apiEndpoint = $"{baseUrl.TrimEnd('/')}{resolvedPath}";

            // Build layout_dict payload matching Python import2boq_v2 format
            var layoutDict = new Dictionary<string, object>();
            foreach (var layout in request.All)
            {
                var details = layout.Detail.Select(d => new Dictionary<string, object?>
                {
                    ["position"] = d.Position,
                    ["product_no"] = d.ProductNo,
                    ["width"] = d.Width,
                    ["height"] = d.Height,
                    ["length"] = d.Length,
                    ["thickness"] = d.Thickness,
                    ["qty"] = d.Qty,
                    ["description"] = d.Description,
                    ["detail_id"] = d.DetailId ?? ""
                }).ToList();

                layoutDict[layout.LayoutName] = new Dictionary<string, object?>
                {
                    ["header_id"] = layout.HeaderId ?? "",
                    ["pr_no"] = layout.PrNo,
                    ["project_name"] = layout.ProjectName,
                    ["job_working_plan_name"] = layout.JobWorkingPlanName,
                    ["product_name"] = layout.ProductName,
                    ["product_catalog"] = layout.ProductCatalog,
                    ["spec"] = layout.Spec,
                    ["surface_treatment"] = layout.SurfaceTreatment,
                    ["operation_flow"] = layout.OperationFlow,
                    ["color_name"] = layout.ColorName,
                    ["color_no"] = layout.ColorNo,
                    ["detail"] = details
                };
            }

            var body = new
            {
                args = new object[] { layoutDict },
                kwargs = new { user_token = userToken },
                context = new { }
            };

            var httpRequest = new HttpRequestMessage(HttpMethod.Post, apiEndpoint);
            httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            httpRequest.Content = new StringContent(
                JsonSerializer.Serialize(body),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.SendAsync(httpRequest);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonSerializer.Deserialize<JsonElement>(responseJson);

            // Check for error response
            if (jsonDoc.TryGetProperty("error_code", out var errorCode))
            {
                return new BoqImportResponse
                {
                    Success = false,
                    ErrorCode = errorCode.GetString(),
                    ErrorMessage = jsonDoc.TryGetProperty("error_message", out var errMsg)
                        ? errMsg.GetString() : "Unknown error"
                };
            }

            // Parse success response: { "all": [ { layout with header_id, detail with detail_id } ] }
            var resultLayouts = new List<BoqImportLayout>();

            JsonElement allElement;
            if (jsonDoc.TryGetProperty("all", out allElement) && allElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var layoutEl in allElement.EnumerateArray())
                {
                    var resultLayout = new BoqImportLayout
                    {
                        LayoutName = layoutEl.TryGetProperty("layout_name", out var ln)
                            ? ln.GetString() ?? "" : "",
                        HeaderId = layoutEl.TryGetProperty("header_id", out var hid)
                            ? hid.ToString() : null
                    };

                    if (layoutEl.TryGetProperty("detail", out var detailArr) &&
                        detailArr.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var detailEl in detailArr.EnumerateArray())
                        {
                            resultLayout.Detail.Add(new BoqImportDetail
                            {
                                ProductNo = detailEl.TryGetProperty("product_no", out var pn)
                                    ? pn.GetString() ?? "" : "",
                                DetailId = detailEl.TryGetProperty("detail_id", out var did)
                                    ? did.ToString() : null
                            });
                        }
                    }

                    resultLayouts.Add(resultLayout);
                }
            }

            _logger?.LogInformation("import2boq_v2 succeeded: {LayoutCount} layouts processed",
                resultLayouts.Count);

            return new BoqImportResponse
            {
                Success = true,
                All = resultLayouts
            };
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "import2boq_v2 API call failed");
            return new BoqImportResponse
            {
                Success = false,
                ErrorCode = "HTTP_ERROR",
                ErrorMessage = ex.Message
            };
        }
    }

    public async Task<IReadOnlyList<BOQEntry>> GetBOQEntriesAsync(int projectId)
    {
        var domain = new object[] { new object[] { "project_id", "=", projectId } };

        var result = await SearchReadAsync("boq.line", domain,
            new[] { "id", "project_id", "product_id", "quantity", "uom_id", "unit_price", "description" });

        return result.Select(r => MapToBOQEntry(r)).ToList();
    }

    public async Task<bool> UpdateBOQEntryAsync(BOQEntry entry)
    {
        if (!entry.Id.HasValue) return false;

        try
        {
            var values = new Dictionary<string, object>
            {
                ["quantity"] = entry.Quantity,
                ["description"] = entry.Description ?? ""
            };

            await WriteAsync("boq.line", new[] { entry.Id.Value }, values);
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to update BOQ entry: {EntryId}", entry.Id);
            return false;
        }
    }

    public async Task<bool> DeleteBOQEntryAsync(int entryId)
    {
        try
        {
            await UnlinkAsync("boq.line", new[] { entryId });
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to delete BOQ entry: {EntryId}", entryId);
            return false;
        }
    }

    #endregion

    #region Purchase Requisition Operations

    public async Task<PREntry?> ConvertBOQToPRAsync(int projectId, IEnumerable<int>? boqEntryIds = null)
    {
        try
        {
            var methodParams = new Dictionary<string, object>
            {
                ["project_id"] = projectId
            };

            if (boqEntryIds != null)
            {
                methodParams["boq_entry_ids"] = boqEntryIds.ToArray();
            }

            // Call custom Odoo method to convert BOQ to PR
            var response = await CallMethodAsync("boq.line", "convert_to_pr", methodParams);

            if (response?.TryGetProperty("result", out var result) == true)
            {
                var prId = result.GetInt32();
                var prEntries = await GetPurchaseRequisitionsAsync(projectId);
                return prEntries.FirstOrDefault(pr => pr.Id == prId);
            }

            return null;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to convert BOQ to PR for project: {ProjectId}", projectId);
            return null;
        }
    }

    public async Task<IReadOnlyList<PREntry>> GetPurchaseRequisitionsAsync(int projectId)
    {
        var domain = new object[] { new object[] { "project_id", "=", projectId } };

        var result = await SearchReadAsync("purchase.requisition", domain,
            new[] { "id", "name", "state", "line_ids" });

        var entries = new List<PREntry>();
        foreach (var r in result)
        {
            var entry = new PREntry
            {
                Id = r.GetProperty("id").GetInt32(),
                ProjectId = projectId,
                Reference = r.GetProperty("name").GetString() ?? "",
                State = r.GetProperty("state").GetString() ?? "draft"
            };
            entries.Add(entry);
        }

        return entries;
    }

    public async Task<bool> SubmitPRAsync(int prId)
    {
        try
        {
            await CallMethodAsync("purchase.requisition", "action_submit", new { id = prId });
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to submit PR: {PRId}", prId);
            return false;
        }
    }

    public async Task<Boq2PrResponse> ConvertBOQToPRViaApiAsync(
        List<string> headerIds, string baseUrl, string endpointPath,
        string database, string userToken)
    {
        try
        {
            var credentials = Convert.ToBase64String(
                Encoding.UTF8.GetBytes($"{database}:{userToken}"));

            var resolvedPath = endpointPath.Replace("{method_name}", "boq2pr_v2");
            var apiEndpoint = $"{baseUrl.TrimEnd('/')}{resolvedPath}";

            var body = new
            {
                args = new object[] { new { all = headerIds } },
                kwargs = new { user_token = userToken },
                context = new { }
            };

            var httpRequest = new HttpRequestMessage(HttpMethod.Post, apiEndpoint);
            httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            httpRequest.Content = new StringContent(
                JsonSerializer.Serialize(body),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.SendAsync(httpRequest);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonSerializer.Deserialize<JsonElement>(responseJson);

            // Check for error response
            if (jsonDoc.TryGetProperty("error_code", out var errorCode))
            {
                return new Boq2PrResponse
                {
                    Success = false,
                    ErrorCode = errorCode.GetString(),
                    ErrorMessage = jsonDoc.TryGetProperty("error_message", out var errMsg)
                        ? errMsg.GetString() : "Unknown error"
                };
            }

            // Parse success response: { "all": [ { pr_id, reference, state, lines } ] }
            var results = new List<Boq2PrResult>();

            if (jsonDoc.TryGetProperty("all", out var allElement) &&
                allElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var prEl in allElement.EnumerateArray())
                {
                    var prResult = new Boq2PrResult
                    {
                        PrId = prEl.TryGetProperty("pr_id", out var pid) && pid.ValueKind == JsonValueKind.Number
                            ? pid.GetInt32() : null,
                        Reference = prEl.TryGetProperty("reference", out var refVal)
                            ? refVal.GetString() ?? "" : "",
                        State = prEl.TryGetProperty("state", out var stateVal)
                            ? stateVal.GetString() ?? "draft" : "draft"
                    };

                    if (prEl.TryGetProperty("lines", out var linesArr) &&
                        linesArr.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var lineEl in linesArr.EnumerateArray())
                        {
                            prResult.Lines.Add(new Boq2PrLine
                            {
                                ProductId = lineEl.TryGetProperty("product_id", out var prodId) && prodId.ValueKind == JsonValueKind.Number
                                    ? prodId.GetInt32() : 0,
                                ProductName = lineEl.TryGetProperty("product_name", out var pn)
                                    ? pn.GetString() ?? "" : "",
                                Quantity = lineEl.TryGetProperty("quantity", out var qty) && qty.ValueKind == JsonValueKind.Number
                                    ? qty.GetDecimal() : 0,
                                UnitOfMeasure = lineEl.TryGetProperty("uom", out var uom)
                                    ? uom.GetString() ?? "" : "",
                                UnitPrice = lineEl.TryGetProperty("unit_price", out var price) && price.ValueKind == JsonValueKind.Number
                                    ? price.GetDecimal() : null
                            });
                        }
                    }

                    results.Add(prResult);
                }
            }

            _logger?.LogInformation("boq2pr_v2 succeeded: {PrCount} PRs created", results.Count);

            return new Boq2PrResponse
            {
                Success = true,
                All = results
            };
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "boq2pr_v2 API call failed");
            return new Boq2PrResponse
            {
                Success = false,
                ErrorCode = "HTTP_ERROR",
                ErrorMessage = ex.Message
            };
        }
    }

    public async Task<ProjectLookupResult> GetProjectViaApiAsync(
        string prNumber, string baseUrl, string endpointPath,
        string database, string userToken)
    {
        try
        {
            var credentials = Convert.ToBase64String(
                Encoding.UTF8.GetBytes($"{database}:{userToken}"));

            var resolvedPath = endpointPath.Replace("{method_name}", "get_project_v2");
            var apiEndpoint = $"{baseUrl.TrimEnd('/')}{resolvedPath}";

            // Match Python: get_project_v2 with exact name match
            var body = new
            {
                args = new object[] { new object[] { new object[] { "name", "=", prNumber } } },
                kwargs = new { user_token = userToken },
                context = new { }
            };

            var httpRequest = new HttpRequestMessage(HttpMethod.Post, apiEndpoint);
            httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            httpRequest.Content = new StringContent(
                JsonSerializer.Serialize(body),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.SendAsync(httpRequest);
            response.EnsureSuccessStatusCode();

            var responseJson = await response.Content.ReadAsStringAsync();
            _logger?.LogInformation("get_project_v2 response: {Response}", responseJson);
            var jsonDoc = JsonSerializer.Deserialize<JsonElement>(responseJson);

            // Check for error response
            if (jsonDoc.TryGetProperty("error_code", out _))
            {
                var errMsg = jsonDoc.TryGetProperty("error_message", out var em)
                    ? em.GetString() : "Unknown error";
                return new ProjectLookupResult(null, $"API error: {errMsg}");
            }

            // Parse response: { id, name, job_working_plan_id, job_working_plan_name }
            int? projectId = jsonDoc.TryGetProperty("id", out var idEl) && idEl.ValueKind == JsonValueKind.Number
                ? idEl.GetInt32() : null;
            string? projectName = jsonDoc.TryGetProperty("name", out var nameEl)
                ? nameEl.GetString() : null;
            int? jwpId = jsonDoc.TryGetProperty("job_working_plan_id", out var jwpIdEl) && jwpIdEl.ValueKind == JsonValueKind.Number
                ? jwpIdEl.GetInt32() : null;
            string? jwpName = jsonDoc.TryGetProperty("job_working_plan_name", out var jwpNameEl)
                ? jwpNameEl.GetString() : null;

            if (projectId == null || projectName == null)
            {
                // Truncate response for log display (max 500 chars)
                var truncated = responseJson.Length > 500 ? responseJson[..500] + "..." : responseJson;
                return new ProjectLookupResult(null, $"No id/name in response: {truncated}");
            }

            return new ProjectLookupResult(
                new OdooProjectInfo(projectId.Value, projectName, jwpId, jwpName),
                $"Found: {projectName} (ID: {projectId})");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "get_project_v2 API call failed for PR '{PrNumber}'", prNumber);
            return new ProjectLookupResult(null, $"API call failed: {ex.Message}");
        }
    }

    #endregion

    #region Synchronization

    public async Task<SyncResult> SyncToOdooAsync(object data, string syncType)
    {
        try
        {
            return syncType.ToLower() switch
            {
                "parameters" => await SyncParametersAsync(data),
                "boq" => await SyncBOQAsync(data),
                "project" => await SyncProjectAsync(data),
                _ => new SyncResult(false, 0, 0, 0, 0, new List<string> { $"Unknown sync type: {syncType}" })
            };
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Sync to Odoo failed: {SyncType}", syncType);
            return new SyncResult(false, 0, 0, 0, 0, new List<string> { ex.Message });
        }
    }

    public DateTime? GetLastSyncTime() => _lastSyncTime;

    private async Task<SyncResult> SyncParametersAsync(object data)
    {
        // Implementation for syncing parameters
        await Task.CompletedTask;
        return new SyncResult(true, 0, 0, 0, 0, null);
    }

    private async Task<SyncResult> SyncBOQAsync(object data)
    {
        if (data is IEnumerable<BOQEntry> entries)
        {
            return await ImportToBOQAsync(entries);
        }

        return new SyncResult(false, 0, 0, 0, 0, new List<string> { "Invalid BOQ data format" });
    }

    private async Task<SyncResult> SyncProjectAsync(object data)
    {
        // Implementation for syncing project data
        await Task.CompletedTask;
        return new SyncResult(true, 0, 0, 0, 0, null);
    }

    #endregion

    #region Private Helper Methods

    private async Task<JsonElement?> PostJsonRpcAsync(string endpoint, object request)
    {
        var json = JsonSerializer.Serialize(request);
        var content = new StringContent(json, Encoding.UTF8, "application/json");

        var response = await _httpClient.PostAsync($"{_serverUrl}{endpoint}", content);
        response.EnsureSuccessStatusCode();

        var responseJson = await response.Content.ReadAsStringAsync();
        return JsonSerializer.Deserialize<JsonElement>(responseJson);
    }

    private async Task<IReadOnlyList<JsonElement>> SearchReadAsync(
        string model, object[] domain, string[] fields, int? limit = null)
    {
        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                model,
                method = "search_read",
                args = new object[] { domain },
                kwargs = new { fields, limit = limit ?? 0 }
            },
            id = 1
        };

        var response = await PostJsonRpcAsync("/web/dataset/call_kw", request);

        if (response?.TryGetProperty("result", out var result) == true)
        {
            return result.EnumerateArray().ToList();
        }

        return Array.Empty<JsonElement>();
    }

    private async Task<IReadOnlyList<JsonElement>> ReadAsync(string model, int[] ids, string[] fields)
    {
        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                model,
                method = "read",
                args = new object[] { ids, fields }
            },
            id = 1
        };

        var response = await PostJsonRpcAsync("/web/dataset/call_kw", request);

        if (response?.TryGetProperty("result", out var result) == true)
        {
            return result.EnumerateArray().ToList();
        }

        return Array.Empty<JsonElement>();
    }

    private async Task<int> CreateAsync(string model, Dictionary<string, object> values)
    {
        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                model,
                method = "create",
                args = new object[] { values }
            },
            id = 1
        };

        var response = await PostJsonRpcAsync("/web/dataset/call_kw", request);

        if (response?.TryGetProperty("result", out var result) == true)
        {
            return result.GetInt32();
        }

        throw new Exception("Failed to create record");
    }

    private async Task WriteAsync(string model, int[] ids, Dictionary<string, object> values)
    {
        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                model,
                method = "write",
                args = new object[] { ids, values }
            },
            id = 1
        };

        await PostJsonRpcAsync("/web/dataset/call_kw", request);
    }

    private async Task UnlinkAsync(string model, int[] ids)
    {
        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                model,
                method = "unlink",
                args = new object[] { ids }
            },
            id = 1
        };

        await PostJsonRpcAsync("/web/dataset/call_kw", request);
    }

    private async Task<JsonElement?> CallMethodAsync(string model, string method, object args)
    {
        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                model,
                method,
                args = new object[] { args }
            },
            id = 1
        };

        return await PostJsonRpcAsync("/web/dataset/call_kw", request);
    }

    private static OdooProject MapToOdooProject(JsonElement data)
    {
        return new OdooProject(
            Id: data.GetProperty("id").GetInt32(),
            Name: data.GetProperty("name").GetString() ?? "",
            Code: data.TryGetProperty("code", out var code) && code.ValueKind == JsonValueKind.String
                ? code.GetString() : null,
            State: data.TryGetProperty("state", out var state) ? state.GetString() ?? "draft" : "draft",
            DateStart: data.TryGetProperty("date_start", out var dateStart) && dateStart.ValueKind == JsonValueKind.String
                ? DateTime.Parse(dateStart.GetString()!) : null,
            DateEnd: data.TryGetProperty("date", out var dateEnd) && dateEnd.ValueKind == JsonValueKind.String
                ? DateTime.Parse(dateEnd.GetString()!) : null);
    }

    private static OdooProduct MapToOdooProduct(JsonElement data)
    {
        return new OdooProduct(
            Id: data.GetProperty("id").GetInt32(),
            Name: data.GetProperty("name").GetString() ?? "",
            Code: data.TryGetProperty("default_code", out var code) && code.ValueKind == JsonValueKind.String
                ? code.GetString() : null,
            Description: data.TryGetProperty("description", out var desc) && desc.ValueKind == JsonValueKind.String
                ? desc.GetString() : null,
            ListPrice: data.TryGetProperty("list_price", out var price) && price.ValueKind == JsonValueKind.Number
                ? price.GetDecimal() : null,
            UnitOfMeasure: data.TryGetProperty("uom_id", out var uom) && uom.ValueKind == JsonValueKind.Array
                ? uom[1].GetString() : null,
            Category: data.TryGetProperty("categ_id", out var cat) && cat.ValueKind == JsonValueKind.Array
                ? cat[1].GetString() : null);
    }

    private static BOQEntry MapToBOQEntry(JsonElement data)
    {
        return new BOQEntry
        {
            Id = data.GetProperty("id").GetInt32(),
            ProjectId = data.TryGetProperty("project_id", out var projId) && projId.ValueKind == JsonValueKind.Array
                ? projId[0].GetInt32() : 0,
            ProductId = data.TryGetProperty("product_id", out var prodId) && prodId.ValueKind == JsonValueKind.Array
                ? prodId[0].GetInt32() : 0,
            ProductName = data.TryGetProperty("product_id", out var prodName) && prodName.ValueKind == JsonValueKind.Array
                ? prodName[1].GetString() ?? "" : "",
            Quantity = data.TryGetProperty("quantity", out var qty) ? qty.GetDecimal() : 0,
            UnitOfMeasure = data.TryGetProperty("uom_id", out var uom) && uom.ValueKind == JsonValueKind.Array
                ? uom[1].GetString() ?? "pcs" : "pcs",
            UnitPrice = data.TryGetProperty("unit_price", out var price) && price.ValueKind == JsonValueKind.Number
                ? price.GetDecimal() : null,
            Description = data.TryGetProperty("description", out var desc) && desc.ValueKind == JsonValueKind.String
                ? desc.GetString() : null
        };
    }

    #endregion

    public void Dispose()
    {
        _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }
}
