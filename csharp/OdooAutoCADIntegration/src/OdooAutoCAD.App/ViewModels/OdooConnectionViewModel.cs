using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Net.Http;
using System.Net.Http.Headers;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.Core.Odoo;

namespace OdooAutoCAD.App.ViewModels;

/// <summary>
/// ViewModel for Odoo Connection page with Swagger/OpenAPI + Basic Auth credential management.
/// Implements INotifyDataErrorInfo for validation.
/// </summary>
public partial class OdooConnectionViewModel : ObservableObject, INotifyDataErrorInfo
{
    private readonly IOdooService _odooService;
    private readonly ISettingsService _settingsService;
    private readonly ICredentialService _credentialService;
    private readonly IConfiguration _configuration;
    private readonly IAppLogService _logService;
    private readonly ILogger<OdooConnectionViewModel>? _logger;
    private readonly Dictionary<string, List<string>> _errors = new();

    // Cache of all products for restoring after search clear
    private List<OdooProduct> _allProducts = new();

    [ObservableProperty]
    private string _swaggerUrl = string.Empty;

    [ObservableProperty]
    private string _userToken = string.Empty;

    [ObservableProperty]
    private bool _showToken;

    [ObservableProperty]
    private bool _isConnected;

    [ObservableProperty]
    private bool _isTesting;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private bool _statusIsSuccess;

    [ObservableProperty]
    private string _odooVersion = string.Empty;

    [ObservableProperty]
    private bool _rememberMe;

    [ObservableProperty]
    private int _productCount;

    [ObservableProperty]
    private string _syncStatusMessage = string.Empty;

    [ObservableProperty]
    private bool _isSyncingProducts;

    #region Product Search & Category Filter (Sprint 8)

    [ObservableProperty]
    private string _productSearchTerm = string.Empty;

    [ObservableProperty]
    private bool _isSearchingProducts;

    [ObservableProperty]
    private string _searchStatusText = string.Empty;

    [ObservableProperty]
    private OdooProductCategory? _selectedCategory;

    [ObservableProperty]
    private bool _isLoadingCategories;

    #endregion

    #region Project Search (Sprint 8)

    [ObservableProperty]
    private string _projectSearchTerm = string.Empty;

    [ObservableProperty]
    private bool _isSearchingProjects;

    [ObservableProperty]
    private OdooProject? _selectedProject;

    [ObservableProperty]
    private string _projectStatusText = string.Empty;

    [ObservableProperty]
    private int _projectCount;

    #endregion

    #region Server Info & Sync Time (Sprint 8)

    [ObservableProperty]
    private string _serverVersion = string.Empty;

    [ObservableProperty]
    private string _connectedDatabase = string.Empty;

    [ObservableProperty]
    private string _connectedUsername = string.Empty;

    [ObservableProperty]
    private DateTime? _lastSyncTime;

    [ObservableProperty]
    private string _lastSyncDisplay = "Never synced";

    #endregion

    public ObservableCollection<OdooProduct> Products { get; } = new();
    public ObservableCollection<OdooProductCategory> Categories { get; } = new();
    public ObservableCollection<OdooProject> Projects { get; } = new();

    public OdooConnectionViewModel(
        IOdooService odooService,
        ISettingsService settingsService,
        ICredentialService credentialService,
        IConfiguration configuration,
        IAppLogService logService,
        ILogger<OdooConnectionViewModel>? logger = null)
    {
        _odooService = odooService;
        _settingsService = settingsService;
        _credentialService = credentialService;
        _configuration = configuration;
        _logService = logService;
        _logger = logger;

        // Initialize last sync time from service
        var syncTime = _odooService.GetLastSyncTime();
        if (syncTime.HasValue)
        {
            LastSyncTime = syncTime.Value;
            LastSyncDisplay = syncTime.Value.ToString("yyyy-MM-dd HH:mm:ss");
        }

        // Load settings on initialization
        _ = LoadSettingsAsync();
    }

    #region INotifyDataErrorInfo Implementation

    public bool HasErrors => _errors.Count > 0;

    public event EventHandler<DataErrorsChangedEventArgs>? ErrorsChanged;

    public IEnumerable GetErrors(string? propertyName)
    {
        if (string.IsNullOrEmpty(propertyName))
            return _errors.Values.SelectMany(e => e);

        return _errors.TryGetValue(propertyName, out var errors) ? errors : Array.Empty<string>();
    }

    private void AddError(string propertyName, string error)
    {
        if (!_errors.ContainsKey(propertyName))
            _errors[propertyName] = new List<string>();

        if (!_errors[propertyName].Contains(error))
        {
            _errors[propertyName].Add(error);
            OnErrorsChanged(propertyName);
        }
    }

    private void ClearErrors(string propertyName)
    {
        if (_errors.Remove(propertyName))
        {
            OnErrorsChanged(propertyName);
        }
    }

    private void OnErrorsChanged(string propertyName)
    {
        ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(propertyName));
        OnPropertyChanged(nameof(HasErrors));
    }

    #endregion

    #region Validation Methods

    partial void OnSwaggerUrlChanged(string value)
    {
        ValidateSwaggerUrl();
    }

    partial void OnUserTokenChanged(string value)
    {
        ValidateUserToken();
    }

    private void ValidateSwaggerUrl()
    {
        ClearErrors(nameof(SwaggerUrl));

        if (string.IsNullOrWhiteSpace(SwaggerUrl))
        {
            AddError(nameof(SwaggerUrl), "Swagger URL is required");
            return;
        }

        if (!Uri.TryCreate(SwaggerUrl, UriKind.Absolute, out var uri) ||
            (uri.Scheme != "http" && uri.Scheme != "https"))
        {
            AddError(nameof(SwaggerUrl), "Invalid URL format");
            return;
        }

        if (!SwaggerUrl.Contains("swagger.json", StringComparison.OrdinalIgnoreCase))
        {
            AddError(nameof(SwaggerUrl), "URL must point to a swagger.json endpoint");
            return;
        }

        var parsed = ParseSwaggerUrl(SwaggerUrl);
        if (parsed == null)
        {
            AddError(nameof(SwaggerUrl), "URL must include 'token' and 'db' query parameters");
        }
    }

    private void ValidateUserToken()
    {
        ClearErrors(nameof(UserToken));

        if (string.IsNullOrWhiteSpace(UserToken))
        {
            AddError(nameof(UserToken), "User token is required");
        }
    }

    private bool ValidateAll()
    {
        ValidateSwaggerUrl();
        ValidateUserToken();

        return !HasErrors;
    }

    #endregion

    #region URL Parsing

    /// <summary>
    /// Parses the Swagger URL to extract base URL, database name, and API token.
    /// Expected format: https://host/api/v1/boq_import_api/swagger.json?token=API_TOKEN&amp;db=DB_NAME
    /// </summary>
    internal static (string BaseUrl, string Database, string ApiToken)? ParseSwaggerUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return null;

        var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
        var apiToken = query["token"];
        var database = query["db"];

        if (string.IsNullOrEmpty(apiToken) || string.IsNullOrEmpty(database))
            return null;

        var baseUrl = $"{uri.Scheme}://{uri.Authority}";
        return (baseUrl, database, apiToken);
    }

    /// <summary>
    /// Resolves the full endpoint path for a Swagger operation from the spec JSON.
    /// Searches the "paths" object for the given operationId and returns basePath + path.
    /// Falls back to basePath + "/" + operationId if not found.
    /// </summary>
    internal static string ResolveSwaggerEndpoint(
        string specJson, string operationId = "callMethodForJobWorkingPlanBoqModel")
    {
        var doc = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(specJson);
        var basePath = doc.TryGetProperty("basePath", out var bp) ? bp.GetString() ?? "" : "";

        if (doc.TryGetProperty("paths", out var paths))
        {
            foreach (var pathEntry in paths.EnumerateObject())
            {
                foreach (var method in pathEntry.Value.EnumerateObject())
                {
                    if (method.Value.TryGetProperty("operationId", out var opId) &&
                        opId.GetString() == operationId)
                    {
                        return basePath + pathEntry.Name;
                    }
                }
            }
        }

        // Fallback: standard Odoo OpenAPI path pattern with {method_name} placeholder
        return basePath + "/job.working.plan.boq/call/{method_name}";
    }

    #endregion

    #region Commands

    [RelayCommand]
    private void ToggleTokenVisibility()
    {
        ShowToken = !ShowToken;
    }

    [RelayCommand(CanExecute = nameof(CanTestConnection))]
    private async Task TestConnectionAsync()
    {
        _logService.Log("Starting connection test...", "Odoo");

        if (!ValidateAll())
        {
            var missingFields = new List<string>();
            if (string.IsNullOrWhiteSpace(SwaggerUrl)) missingFields.Add("Swagger URL");
            if (string.IsNullOrWhiteSpace(UserToken)) missingFields.Add("User Token");

            if (missingFields.Any())
            {
                StatusMessage = $"Missing required fields: {string.Join(", ", missingFields)}";
                _logService.Log($"Validation failed - missing: {string.Join(", ", missingFields)}", "Odoo", AppLogLevel.Warning);
            }
            else
            {
                StatusMessage = "Please fix validation errors before testing connection";
                _logService.Log("Validation failed - fix errors above", "Odoo", AppLogLevel.Warning);
            }

            StatusIsSuccess = false;
            return;
        }

        IsTesting = true;
        StatusMessage = "Testing connection...";
        StatusIsSuccess = false;
        OdooVersion = string.Empty;

        try
        {
            var parsed = ParseSwaggerUrl(SwaggerUrl);
            if (parsed == null)
            {
                StatusMessage = "Invalid Swagger URL format";
                StatusIsSuccess = false;
                _logService.Log("FAILED - could not parse Swagger URL", "Odoo", AppLogLevel.Error);
                return;
            }

            var (baseUrl, database, apiToken) = parsed.Value;
            _logService.Log($"Base URL: {baseUrl} | DB: {database}", "Odoo");

            var status = await TestConnectionWithDetailsAsync(SwaggerUrl, baseUrl, database, apiToken, UserToken);

            if (status.IsConnected)
            {
                // Mark the service as API-authenticated so sidebar/dashboard timers see it
                _odooService.MarkApiAuthenticated(baseUrl, database);

                StatusMessage = $"Connected successfully! API: {status.Version ?? "available"}";
                StatusIsSuccess = true;
                OdooVersion = status.Version ?? "API Available";
                IsConnected = true;

                // Populate server info (Sprint 8)
                ServerVersion = status.Version ?? "Unknown";
                ConnectedDatabase = database;
                ConnectedUsername = UserToken.Length > 8
                    ? $"{UserToken[..4]}...{UserToken[^4..]}"
                    : "(token)";

                _logService.Log($"SUCCESS - {status.Version ?? "API available"}", "Odoo");
                _logger?.LogInformation("Test connection successful: {Version}", status.Version);
            }
            else
            {
                StatusMessage = $"Connection failed: {status.ErrorMessage}";
                StatusIsSuccess = false;
                IsConnected = false;
                _logService.Log($"FAILED - {status.ErrorMessage}", "Odoo", AppLogLevel.Error);
                _logger?.LogWarning("Test connection failed: {Error}", status.ErrorMessage);
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Unexpected error: {ex.Message}";
            StatusIsSuccess = false;
            IsConnected = false;
            _logService.Log($"EXCEPTION - {ex.GetType().Name}: {ex.Message}", "Odoo", AppLogLevel.Error);
            _logger?.LogError(ex, "Test connection error");
        }
        finally
        {
            IsTesting = false;
            _logService.Log("Test complete.", "Odoo");
        }
    }

    private bool CanTestConnection()
    {
        return !IsTesting;
    }

    [RelayCommand(CanExecute = nameof(CanDisconnect))]
    private async Task DisconnectAsync()
    {
        try
        {
            _odooService.ClearApiAuthentication();
            IsConnected = false;
            OdooVersion = string.Empty;
            StatusMessage = "Disconnected from Odoo";
            StatusIsSuccess = false;

            // Clear server info (Sprint 8)
            ServerVersion = string.Empty;
            ConnectedDatabase = string.Empty;
            ConnectedUsername = string.Empty;

            _logService.Log("Disconnected from Odoo", "Odoo");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error during Odoo disconnect");
            StatusMessage = $"Disconnect error: {ex.Message}";
        }
        await Task.CompletedTask;
    }

    private bool CanDisconnect() => IsConnected && !IsTesting;

    partial void OnIsConnectedChanged(bool value)
    {
        DisconnectCommand.NotifyCanExecuteChanged();
        SyncProductsCommand.NotifyCanExecuteChanged();
        SearchProductsCommand.NotifyCanExecuteChanged();
        LoadCategoriesCommand.NotifyCanExecuteChanged();
        SearchProjectsCommand.NotifyCanExecuteChanged();
        LoadProjectsCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanSyncProducts))]
    private async Task SyncProductsAsync()
    {
        IsSyncingProducts = true;
        SyncStatusMessage = "Syncing products...";

        try
        {
            var parsed = ParseSwaggerUrl(SwaggerUrl);
            if (parsed == null)
            {
                SyncStatusMessage = "Invalid Swagger URL";
                return;
            }

            var (baseUrl, database, apiToken) = parsed.Value;

            // Resolve endpoint path from swagger spec
            var endpointPath = "/api/v1/boq_import_api/job.working.plan.boq/call/{method_name}";
            try
            {
                using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
                var swaggerResponse = await httpClient.GetAsync(SwaggerUrl);
                if (swaggerResponse.IsSuccessStatusCode)
                {
                    var json = await swaggerResponse.Content.ReadAsStringAsync();
                    endpointPath = ResolveSwaggerEndpoint(json);
                    _logService.Log($"Swagger endpoint resolved: {endpointPath}", "Odoo");
                }
                else
                {
                    _logService.Log($"Swagger spec fetch returned HTTP {(int)swaggerResponse.StatusCode}, using default path", "Odoo", AppLogLevel.Warning);
                }
            }
            catch (Exception ex)
            {
                _logService.Log($"Swagger spec fetch failed ({ex.Message}), using default path", "Odoo", AppLogLevel.Warning);
            }

            var products = await _odooService.GetProductsViaApiAsync(baseUrl, endpointPath, database, UserToken);

            Products.Clear();
            _allProducts = products.ToList();
            foreach (var product in products)
            {
                Products.Add(product);
            }
            ProductCount = Products.Count;

            // Update sync time (Sprint 8)
            LastSyncTime = DateTime.Now;
            LastSyncDisplay = LastSyncTime.Value.ToString("yyyy-MM-dd HH:mm:ss");

            SyncStatusMessage = $"Synced {ProductCount} products";
            _logService.Log($"Product sync complete: {ProductCount} products", "Odoo");
        }
        catch (Exception ex)
        {
            SyncStatusMessage = $"Sync failed: {ex.Message}";
            _logService.Log($"Product sync failed: {ex.Message}", "Odoo", AppLogLevel.Error);
            _logger?.LogError(ex, "Product sync error");
        }
        finally
        {
            IsSyncingProducts = false;
        }
    }

    private bool CanSyncProducts() => IsConnected && !IsSyncingProducts;

    partial void OnIsSyncingProductsChanged(bool value)
    {
        SyncProductsCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    private async Task SaveSettingsAsync()
    {
        if (!ValidateAll())
        {
            StatusMessage = "Please fix validation errors before saving";
            StatusIsSuccess = false;
            return;
        }

        try
        {
            // Save to database (including sensitive user token)
            await _settingsService.SaveServerConfigsAsync(new Dictionary<string, string?>
            {
                ["odoo_swagger_url"] = SwaggerUrl,
                ["odoo_user_token"] = UserToken
            });

            // Also update appsettings.json (non-sensitive parts only)
            var appSettings = _settingsService.GetAppSettings();
            appSettings.Odoo.SwaggerUrl = SwaggerUrl;
            await _settingsService.SaveAppSettingsAsync(appSettings);

            // Handle DPAPI credential persistence
            if (RememberMe)
            {
                await _credentialService.SaveCredentialsAsync(SwaggerUrl, UserToken);
            }
            else
            {
                await _credentialService.ClearCredentialsAsync();
            }

            StatusMessage = "Settings saved successfully";
            StatusIsSuccess = true;
            _logger?.LogInformation("Odoo connection settings saved");
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to save settings: {ex.Message}";
            StatusIsSuccess = false;
            _logger?.LogError(ex, "Failed to save Odoo settings");
        }
    }

    #endregion

    #region Product Search & Category Filter Commands (Sprint 8)

    [RelayCommand(CanExecute = nameof(CanSearchProducts))]
    private async Task SearchProductsAsync()
    {
        if (string.IsNullOrWhiteSpace(ProductSearchTerm))
            return;

        IsSearchingProducts = true;
        SearchStatusText = "Searching...";

        try
        {
            var results = await _odooService.SearchProductsAsync(ProductSearchTerm);

            Products.Clear();
            foreach (var product in results)
            {
                Products.Add(product);
            }
            ProductCount = Products.Count;
            SearchStatusText = $"Found {ProductCount} products matching \"{ProductSearchTerm}\"";
            _logService.Log($"Product search \"{ProductSearchTerm}\": {ProductCount} results", "Odoo");
        }
        catch (Exception ex)
        {
            SearchStatusText = $"Search failed: {ex.Message}";
            _logService.Log($"Product search failed: {ex.Message}", "Odoo", AppLogLevel.Error);
            _logger?.LogError(ex, "Product search error");
        }
        finally
        {
            IsSearchingProducts = false;
        }
    }

    private bool CanSearchProducts() => IsConnected && !IsSearchingProducts && !string.IsNullOrWhiteSpace(ProductSearchTerm);

    partial void OnProductSearchTermChanged(string value)
    {
        SearchProductsCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    private void ClearProductSearch()
    {
        ProductSearchTerm = string.Empty;
        SearchStatusText = string.Empty;
        SelectedCategory = null;

        Products.Clear();
        foreach (var product in _allProducts)
        {
            Products.Add(product);
        }
        ProductCount = Products.Count;
    }

    [RelayCommand(CanExecute = nameof(CanLoadCategories))]
    private async Task LoadCategoriesAsync()
    {
        IsLoadingCategories = true;

        try
        {
            var categories = await _odooService.GetCategoriesAsync();

            Categories.Clear();
            foreach (var category in categories)
            {
                Categories.Add(category);
            }
            _logService.Log($"Loaded {categories.Count} product categories", "Odoo");
        }
        catch (Exception ex)
        {
            _logService.Log($"Load categories failed: {ex.Message}", "Odoo", AppLogLevel.Error);
            _logger?.LogError(ex, "Load categories error");
        }
        finally
        {
            IsLoadingCategories = false;
        }
    }

    private bool CanLoadCategories() => IsConnected && !IsLoadingCategories;

    partial void OnSelectedCategoryChanged(OdooProductCategory? value)
    {
        if (value != null)
        {
            _ = FilterByCategoryAsync(value.Id);
        }
    }

    private async Task FilterByCategoryAsync(int categoryId)
    {
        IsSearchingProducts = true;
        SearchStatusText = "Filtering by category...";

        try
        {
            var results = await _odooService.GetProductsByCategoryAsync(categoryId);

            Products.Clear();
            foreach (var product in results)
            {
                Products.Add(product);
            }
            ProductCount = Products.Count;
            SearchStatusText = $"Showing {ProductCount} products in category";
            _logService.Log($"Category filter: {ProductCount} products", "Odoo");
        }
        catch (Exception ex)
        {
            SearchStatusText = $"Filter failed: {ex.Message}";
            _logger?.LogError(ex, "Category filter error");
        }
        finally
        {
            IsSearchingProducts = false;
        }
    }

    #endregion

    #region Project Search Commands (Sprint 8)

    [RelayCommand(CanExecute = nameof(CanSearchProjects))]
    private async Task SearchProjectsAsync()
    {
        if (string.IsNullOrWhiteSpace(ProjectSearchTerm))
            return;

        IsSearchingProjects = true;
        ProjectStatusText = "Searching projects...";

        try
        {
            var results = await _odooService.SearchProjectsAsync(ProjectSearchTerm);

            Projects.Clear();
            foreach (var project in results)
            {
                Projects.Add(project);
            }
            ProjectCount = Projects.Count;
            ProjectStatusText = $"Found {ProjectCount} projects matching \"{ProjectSearchTerm}\"";
            _logService.Log($"Project search \"{ProjectSearchTerm}\": {ProjectCount} results", "Odoo");
        }
        catch (Exception ex)
        {
            ProjectStatusText = $"Search failed: {ex.Message}";
            _logService.Log($"Project search failed: {ex.Message}", "Odoo", AppLogLevel.Error);
            _logger?.LogError(ex, "Project search error");
        }
        finally
        {
            IsSearchingProjects = false;
        }
    }

    private bool CanSearchProjects() => IsConnected && !IsSearchingProjects && !string.IsNullOrWhiteSpace(ProjectSearchTerm);

    partial void OnProjectSearchTermChanged(string value)
    {
        SearchProjectsCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanLoadProjects))]
    private async Task LoadProjectsAsync()
    {
        IsSearchingProjects = true;
        ProjectStatusText = "Loading all projects...";

        try
        {
            var results = await _odooService.GetProjectsAsync();

            Projects.Clear();
            foreach (var project in results)
            {
                Projects.Add(project);
            }
            ProjectCount = Projects.Count;
            ProjectStatusText = $"Loaded {ProjectCount} active projects";
            _logService.Log($"Loaded {ProjectCount} projects", "Odoo");
        }
        catch (Exception ex)
        {
            ProjectStatusText = $"Load failed: {ex.Message}";
            _logService.Log($"Load projects failed: {ex.Message}", "Odoo", AppLogLevel.Error);
            _logger?.LogError(ex, "Load projects error");
        }
        finally
        {
            IsSearchingProjects = false;
        }
    }

    private bool CanLoadProjects() => IsConnected && !IsSearchingProjects;

    [RelayCommand]
    private async Task SelectProjectAsync()
    {
        if (SelectedProject == null)
        {
            ProjectStatusText = "No project selected";
            return;
        }

        try
        {
            await _settingsService.SaveServerConfigsAsync(new Dictionary<string, string?>
            {
                ["odoo_project_id"] = SelectedProject.Id.ToString(),
                ["odoo_project_name"] = SelectedProject.Name
            });

            ProjectStatusText = $"Selected: {SelectedProject.Name} (ID: {SelectedProject.Id})";
            _logService.Log($"Project selected: {SelectedProject.Name} (ID: {SelectedProject.Id})", "Odoo");
        }
        catch (Exception ex)
        {
            ProjectStatusText = $"Failed to save project selection: {ex.Message}";
            _logger?.LogError(ex, "Save project selection error");
        }
    }

    #endregion

    #region Private Methods

    private async Task LoadSettingsAsync()
    {
        try
        {
            // Load from appsettings.json via configuration
            var odooSection = _configuration.GetSection("Odoo");
            if (odooSection.Exists())
            {
                SwaggerUrl = odooSection["SwaggerUrl"] ?? string.Empty;
            }

            // Override with database values if they exist (including user token)
            var configs = await _settingsService.LoadServerConfigsAsync();
            if (configs.TryGetValue("odoo_swagger_url", out var url) && !string.IsNullOrEmpty(url))
                SwaggerUrl = url;
            if (configs.TryGetValue("odoo_user_token", out var token) && !string.IsNullOrEmpty(token))
                UserToken = token;

            // Try loading DPAPI-encrypted credentials (overrides if present)
            if (_credentialService.HasSavedCredentials)
            {
                var (savedUrl, savedToken) = await _credentialService.LoadCredentialsAsync();
                if (!string.IsNullOrEmpty(savedUrl))
                    SwaggerUrl = savedUrl;
                if (!string.IsNullOrEmpty(savedToken))
                    UserToken = savedToken;
                RememberMe = true;
            }

            if (string.IsNullOrEmpty(SwaggerUrl))
            {
                StatusMessage = "No saved configuration found. Please enter your Odoo connection details";
                StatusIsSuccess = false;
            }
            else
            {
                StatusMessage = "Settings loaded from configuration";
                StatusIsSuccess = true;
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to load Odoo settings");
            StatusMessage = $"Failed to load settings: {ex.Message}";
            StatusIsSuccess = false;
        }
    }

    /// <summary>
    /// Tests connection using Swagger spec fetch + Basic Auth test call.
    /// Step 1: Fetch swagger.json (validates server + API module + API access token)
    /// Step 2: POST a lightweight test call with Basic Auth (validates user token)
    /// </summary>
    private async Task<OdooStatus> TestConnectionWithDetailsAsync(
        string swaggerUrl, string baseUrl, string database, string apiToken, string userToken)
    {
        try
        {
            var timeoutSeconds = _configuration.GetValue<int>("Odoo:TimeoutSeconds", 30);
            var timeout = TimeSpan.FromSeconds(Math.Clamp(timeoutSeconds, 5, 120));
            _logService.Log($"Timeout: {timeoutSeconds}s", "Odoo");

            using var httpClient = new HttpClient { Timeout = timeout };

            // Step 1: Fetch swagger.json
            _logService.Log("Step 1: Fetching swagger spec...", "Odoo");
            _logService.Log($"GET {swaggerUrl}", "Odoo");

            var swaggerResponse = await httpClient.GetAsync(swaggerUrl);
            _logService.Log($"HTTP {(int)swaggerResponse.StatusCode} {swaggerResponse.ReasonPhrase}", "Odoo");

            if (!swaggerResponse.IsSuccessStatusCode)
            {
                var errorBody = await swaggerResponse.Content.ReadAsStringAsync();
                _logService.Log($"Response body: {Truncate(errorBody, 500)}", "Odoo", AppLogLevel.Debug);

                return swaggerResponse.StatusCode switch
                {
                    System.Net.HttpStatusCode.NotFound =>
                        new OdooStatus(false, baseUrl, database, null, null,
                            "Swagger endpoint not found. Is the boq_import_api module installed?"),
                    System.Net.HttpStatusCode.Forbidden or
                    System.Net.HttpStatusCode.Unauthorized =>
                        new OdooStatus(false, baseUrl, database, null, null,
                            "API access denied. Check the token parameter in the Swagger URL"),
                    _ =>
                        new OdooStatus(false, baseUrl, database, null, null,
                            $"HTTP {(int)swaggerResponse.StatusCode}: {swaggerResponse.ReasonPhrase}")
                };
            }

            var swaggerJson = await swaggerResponse.Content.ReadAsStringAsync();
            _logService.Log($"Swagger spec received ({swaggerJson.Length} bytes)", "Odoo");

            // Parse swagger spec to extract basePath and version
            string? basePath = null;
            string? apiVersion = null;
            try
            {
                var swaggerDoc = System.Text.Json.JsonDocument.Parse(swaggerJson);
                if (swaggerDoc.RootElement.TryGetProperty("basePath", out var bp))
                    basePath = bp.GetString();
                if (swaggerDoc.RootElement.TryGetProperty("info", out var info) &&
                    info.TryGetProperty("version", out var ver))
                    apiVersion = ver.GetString();

                _logService.Log($"API basePath: {basePath ?? "(none)"}", "Odoo");
                if (apiVersion != null)
                    _logService.Log($"API version: {apiVersion}", "Odoo");
            }
            catch (System.Text.Json.JsonException)
            {
                _logService.Log("Could not parse swagger spec as JSON", "Odoo", AppLogLevel.Warning);
            }

            // Step 2: Test Basic Auth with a lightweight API call
            _logService.Log("Step 2: Testing Basic Auth (user token)...", "Odoo");

            var credentials = Convert.ToBase64String(
                System.Text.Encoding.UTF8.GetBytes($"{database}:{userToken}"));

            // Use basePath from swagger spec, or fallback to common pattern
            var apiEndpoint = basePath != null
                ? $"{baseUrl}{basePath}"
                : $"{baseUrl}/api/v1/boq_import_api";

            // Try a method call to validate the user token
            var testUrl = $"{apiEndpoint}/callMethodForJobWorkingPlanBoqModel";
            _logService.Log($"POST {testUrl}", "Odoo");

            var testRequest = new HttpRequestMessage(HttpMethod.Post, testUrl);
            testRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            testRequest.Content = new StringContent(
                System.Text.Json.JsonSerializer.Serialize(new { method_name = "test_connection" }),
                System.Text.Encoding.UTF8,
                "application/json");

            var testResponse = await httpClient.SendAsync(testRequest);
            _logService.Log($"HTTP {(int)testResponse.StatusCode} {testResponse.ReasonPhrase}", "Odoo");

            var testBody = await testResponse.Content.ReadAsStringAsync();
            _logService.Log($"Response body: {Truncate(testBody, 500)}", "Odoo", AppLogLevel.Debug);

            if (testResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                testResponse.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                return new OdooStatus(false, baseUrl, database, null, apiVersion,
                    "User token authentication failed. Check your User Token value");
            }

            // Any non-auth-error response means the connection and auth work.
            // Even a 400/404/500 from the API method itself means we authenticated OK.
            var versionDisplay = apiVersion != null ? $"API v{apiVersion}" : "API Available";
            _logService.Log($"Authentication successful! {versionDisplay}", "Odoo");

            return new OdooStatus(true, baseUrl, database, null, versionDisplay, null);
        }
        catch (TaskCanceledException)
        {
            _logService.Log("TIMEOUT - request cancelled", "Odoo", AppLogLevel.Error);
            return new OdooStatus(false, baseUrl, database, null, null,
                $"Connection timed out after {_configuration.GetValue<int>("Odoo:TimeoutSeconds", 30)} seconds");
        }
        catch (HttpRequestException ex)
        {
            _logService.Log($"HTTP ERROR - {ex.Message}", "Odoo", AppLogLevel.Error);
            return new OdooStatus(false, baseUrl, database, null, null,
                $"Unable to reach server at {baseUrl}: {ex.Message}");
        }
        catch (System.Text.Json.JsonException ex)
        {
            _logService.Log($"JSON PARSE ERROR - {ex.Message}", "Odoo", AppLogLevel.Error);
            return new OdooStatus(false, baseUrl, database, null, null,
                $"Invalid response from server: {ex.Message}");
        }
        catch (Exception ex)
        {
            _logService.Log($"UNEXPECTED - {ex.GetType().Name}: {ex.Message}", "Odoo", AppLogLevel.Error);
            _logger?.LogError(ex, "Unexpected error testing Odoo connection");
            return new OdooStatus(false, baseUrl, database, null, null,
                $"Unexpected error: {ex.Message}");
        }
    }

    private static string Truncate(string value, int maxLength)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value.Length <= maxLength ? value : value[..maxLength] + "...";
    }

    #endregion
}
