using System.Collections;
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
    private readonly IConfiguration _configuration;
    private readonly IAppLogService _logService;
    private readonly ILogger<OdooConnectionViewModel>? _logger;
    private readonly Dictionary<string, List<string>> _errors = new();

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

    public OdooConnectionViewModel(
        IOdooService odooService,
        ISettingsService settingsService,
        IConfiguration configuration,
        IAppLogService logService,
        ILogger<OdooConnectionViewModel>? logger = null)
    {
        _odooService = odooService;
        _settingsService = settingsService;
        _configuration = configuration;
        _logService = logService;
        _logger = logger;

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
