using System.Collections.ObjectModel;
using System.Net.Http;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.App.ViewModels;

/// <summary>
/// ViewModel for the Parameter Configuration page (FR-009).
/// Loads material/spec/category/process/surface/color from Odoo APIs,
/// then writes 7 attribute values to all AutoCAD layout parameter blocks.
/// </summary>
public partial class ParameterConfigViewModel : ObservableObject
{
    private readonly IAutoCADService _autoCADService;
    private readonly IOdooService _odooService;
    private readonly IGUIProxy _guiProxy;
    private readonly ISettingsService _settingsService;
    private readonly IAppLogService _logService;
    private readonly ILogger<ParameterConfigViewModel>? _logger;

    // Dropdown collections
    public ObservableCollection<OdooProduct> Products { get; } = new();
    public ObservableCollection<OdooSetupValue> Specs { get; } = new();
    public ObservableCollection<OdooSetupValue> Categories { get; } = new();
    public ObservableCollection<OdooSetupValue> OperationFlows { get; } = new();
    public ObservableCollection<OdooSetupValue> SurfaceTreatments { get; } = new();
    public ObservableCollection<OdooColor> Colors { get; } = new();

    // Selected items
    [ObservableProperty]
    private OdooProduct? _selectedProduct;

    [ObservableProperty]
    private OdooSetupValue? _selectedSpec;

    [ObservableProperty]
    private OdooSetupValue? _selectedCategory;

    [ObservableProperty]
    private OdooSetupValue? _selectedOperationFlow;

    [ObservableProperty]
    private OdooSetupValue? _selectedSurfaceTreatment;

    [ObservableProperty]
    private OdooColor? _selectedColor;

    // Auto-fill (read-only display)
    [ObservableProperty]
    private string _unit = string.Empty;

    [ObservableProperty]
    private string _colorNo = string.Empty;

    // Layout info
    [ObservableProperty]
    private string _activeLayoutName = string.Empty;

    [ObservableProperty]
    private string _layoutSummary = string.Empty;

    // State
    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private bool _isSubmitting;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private string _lastSubmitResult = string.Empty;

    public bool IsAutoCADConnected => _autoCADService.IsConnected;
    public bool IsOdooConnected => _odooService.IsApiAuthenticated || _odooService.IsConnected;

    public bool AllFieldsFilled =>
        SelectedProduct != null &&
        SelectedSpec != null &&
        SelectedCategory != null &&
        SelectedOperationFlow != null &&
        SelectedSurfaceTreatment != null &&
        SelectedColor != null &&
        !string.IsNullOrEmpty(ColorNo);

    public bool CanSubmit => AllFieldsFilled && IsAutoCADConnected && !IsSubmitting;

    public ParameterConfigViewModel(
        IAutoCADService autoCADService,
        IOdooService odooService,
        IGUIProxy guiProxy,
        ISettingsService settingsService,
        IAppLogService logService,
        ILogger<ParameterConfigViewModel>? logger = null)
    {
        _autoCADService = autoCADService;
        _odooService = odooService;
        _guiProxy = guiProxy;
        _settingsService = settingsService;
        _logService = logService;
        _logger = logger;
    }

    partial void OnSelectedProductChanged(OdooProduct? value)
    {
        Unit = value?.UnitOfMeasure ?? string.Empty;
        OnPropertyChanged(nameof(AllFieldsFilled));
        OnPropertyChanged(nameof(CanSubmit));
    }

    partial void OnSelectedSpecChanged(OdooSetupValue? value)
    {
        OnPropertyChanged(nameof(AllFieldsFilled));
        OnPropertyChanged(nameof(CanSubmit));
    }

    partial void OnSelectedCategoryChanged(OdooSetupValue? value)
    {
        OnPropertyChanged(nameof(AllFieldsFilled));
        OnPropertyChanged(nameof(CanSubmit));
    }

    partial void OnSelectedOperationFlowChanged(OdooSetupValue? value)
    {
        OnPropertyChanged(nameof(AllFieldsFilled));
        OnPropertyChanged(nameof(CanSubmit));
    }

    partial void OnSelectedSurfaceTreatmentChanged(OdooSetupValue? value)
    {
        OnPropertyChanged(nameof(AllFieldsFilled));
        OnPropertyChanged(nameof(CanSubmit));
    }

    partial void OnSelectedColorChanged(OdooColor? value)
    {
        ColorNo = value?.ColorNo ?? string.Empty;
        OnPropertyChanged(nameof(AllFieldsFilled));
        OnPropertyChanged(nameof(CanSubmit));
    }

    partial void OnIsSubmittingChanged(bool value)
    {
        OnPropertyChanged(nameof(CanSubmit));
    }

    [RelayCommand]
    private async Task LoadOptionsAsync()
    {
        if (IsLoading) return;

        IsLoading = true;
        StatusMessage = "Loading options from Odoo...";
        LastSubmitResult = string.Empty;

        try
        {
            // Step 1: Load Swagger credentials
            _logService.Log("Loading configuration...", "ParamConfig");
            var configs = await _settingsService.LoadServerConfigsAsync();
            configs.TryGetValue("odoo_swagger_url", out var swaggerUrl);
            configs.TryGetValue("odoo_user_token", out var userToken);

            if (string.IsNullOrEmpty(swaggerUrl))
            {
                var appSettings = _settingsService.GetAppSettings();
                swaggerUrl = appSettings.Odoo.SwaggerUrl;
            }

            if (string.IsNullOrEmpty(swaggerUrl) || string.IsNullOrEmpty(userToken))
            {
                StatusMessage = "Missing Swagger URL or user token. Check Settings > Connection.";
                _logService.Log("Missing Swagger URL or user token", "ParamConfig", AppLogLevel.Error);
                return;
            }

            var parsed = OdooConnectionViewModel.ParseSwaggerUrl(swaggerUrl);
            if (parsed == null)
            {
                StatusMessage = "Invalid Swagger URL format.";
                _logService.Log($"Cannot parse Swagger URL: {swaggerUrl}", "ParamConfig", AppLogLevel.Error);
                return;
            }

            var (baseUrl, database, apiToken) = parsed.Value;
            _logService.Log($"Config: {baseUrl} | DB: {database}", "ParamConfig");

            // Step 2: Resolve endpoint path
            var endpointPath = "/api/v1/boq_import_api/job.working.plan.boq/call/{method_name}";
            try
            {
                using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
                var specJson = await httpClient.GetStringAsync(swaggerUrl);
                endpointPath = OdooConnectionViewModel.ResolveSwaggerEndpoint(specJson);
                _logService.Log($"Endpoint: {endpointPath}", "ParamConfig");
            }
            catch (Exception ex)
            {
                _logService.Log($"Swagger spec fetch failed ({ex.Message}), using default", "ParamConfig", AppLogLevel.Warning);
            }

            // Step 3: Resolve project ID for colors
            int projectId = 0;
            try
            {
                StatusMessage = "Reading PR number from AutoCAD...";
                var prResponse = await _guiProxy.ExecuteInGuiAsync("autocad_get_pr_number", null, timeout: 5000);
                var prNumber = prResponse.Result as string ?? "";
                if (!string.IsNullOrEmpty(prNumber))
                {
                    _logService.Log($"PR number: {prNumber}", "ParamConfig");
                    var projectResult = await _odooService.GetProjectViaApiAsync(
                        prNumber, baseUrl, endpointPath, database, userToken);
                    projectId = projectResult.Project?.Id ?? 0;
                    _logService.Log($"Project ID: {projectId}", "ParamConfig");
                }
                else
                {
                    _logService.Log("No PR number found in AutoCAD", "ParamConfig", AppLogLevel.Warning);
                }
            }
            catch (Exception ex)
            {
                _logService.Log($"Project lookup failed: {ex.Message}", "ParamConfig", AppLogLevel.Warning);
            }

            // Step 3b: Fetch active layout name
            try
            {
                var activeResponse = await _guiProxy.ExecuteInGuiAsync("autocad_get_active_layout", null, timeout: 5000);
                ActiveLayoutName = activeResponse.Result as string ?? "";

                LayoutSummary = !string.IsNullOrEmpty(ActiveLayoutName)
                    ? $"Current layout: {ActiveLayoutName}  —  Submit will update this layout only"
                    : "No active layout found";
                _logService.Log(LayoutSummary, "ParamConfig");
            }
            catch (Exception ex)
            {
                LayoutSummary = "Could not read layouts";
                _logService.Log($"Layout fetch failed: {ex.Message}", "ParamConfig", AppLogLevel.Warning);
            }

            // Step 4: Fetch each dropdown independently (partial success is OK)
            StatusMessage = "Fetching dropdown data...";
            var errors = new List<string>();

            await LoadCollectionAsync("Products", Products,
                () => _odooService.GetProductsViaApiAsync(baseUrl, endpointPath, database, userToken), errors);

            await LoadCollectionAsync("Specs", Specs,
                () => _odooService.GetSetupViaApiAsync("spec", baseUrl, endpointPath, database, userToken), errors);

            await LoadCollectionAsync("Categories", Categories,
                () => _odooService.GetSetupViaApiAsync("product_catelog", baseUrl, endpointPath, database, userToken), errors);

            await LoadCollectionAsync("OperationFlows", OperationFlows,
                () => _odooService.GetSetupViaApiAsync("operation_flow", baseUrl, endpointPath, database, userToken), errors);

            await LoadCollectionAsync("SurfaceTreatments", SurfaceTreatments,
                () => _odooService.GetSetupViaApiAsync("surface_treatment", baseUrl, endpointPath, database, userToken), errors);

            if (projectId > 0)
            {
                await LoadCollectionAsync("Colors", Colors,
                    () => _odooService.GetColorsViaApiAsync(projectId, baseUrl, endpointPath, database, userToken), errors);
            }
            else
            {
                Colors.Clear();
                _logService.Log("Colors skipped (no project ID)", "ParamConfig", AppLogLevel.Warning);
            }

            // Summary
            var counts = $"{Products.Count} products, {Specs.Count} specs, {Categories.Count} categories, " +
                         $"{OperationFlows.Count} flows, {SurfaceTreatments.Count} surfaces, {Colors.Count} colors";

            if (errors.Count > 0)
            {
                StatusMessage = $"Loaded with {errors.Count} error(s): {counts}";
                _logService.Log($"Load partial — {string.Join("; ", errors)}", "ParamConfig", AppLogLevel.Warning);
            }
            else
            {
                StatusMessage = $"Loaded: {counts}";
                _logService.Log($"All options loaded: {counts}", "ParamConfig");
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Load failed: {ex.Message}";
            _logger?.LogError(ex, "Failed to load parameter config options");
            _logService.Log($"Load failed: {ex.Message}", "ParamConfig", AppLogLevel.Error);
        }
        finally
        {
            IsLoading = false;
        }
    }

    /// <summary>
    /// Loads a single dropdown collection with independent error handling.
    /// Partial failures don't block other dropdowns.
    /// </summary>
    private async Task LoadCollectionAsync<T>(string name, ObservableCollection<T> collection,
        Func<Task<IReadOnlyList<T>>> fetcher, List<string> errors)
    {
        try
        {
            var items = await fetcher();
            collection.Clear();
            foreach (var item in items) collection.Add(item);
            _logService.Log($"{name}: {items.Count} items", "ParamConfig");
        }
        catch (Exception ex)
        {
            collection.Clear();
            errors.Add($"{name}: {ex.Message}");
            _logService.Log($"{name} failed: {ex.Message}", "ParamConfig", AppLogLevel.Error);
        }
    }

    [RelayCommand]
    private async Task SubmitAsync()
    {
        if (!CanSubmit) return;

        IsSubmitting = true;
        StatusMessage = "Writing parameters to AutoCAD...";
        LastSubmitResult = string.Empty;

        try
        {
            var attributes = new Dictionary<string, string>
            {
                ["product_name"] = SelectedProduct!.Name,
                ["spec"] = SelectedSpec!.Value,
                ["product_catelog"] = SelectedCategory!.Value,
                ["operation_flow"] = SelectedOperationFlow!.Value,
                ["surface_treatment"] = SelectedSurfaceTreatment!.Value,
                ["color_name"] = SelectedColor!.Name,
                ["color_no"] = ColorNo
            };

            // Re-read current active layout at submit time
            var currentLayoutResp = await _guiProxy.ExecuteInGuiAsync("autocad_get_active_layout", null, timeout: 5000);
            var currentLayout = currentLayoutResp.Result as string ?? ActiveLayoutName;
            if (!string.IsNullOrEmpty(currentLayout))
                ActiveLayoutName = currentLayout;

            var response = await _guiProxy.ExecuteInGuiAsync("autocad_set_attribute_values",
                new Dictionary<string, object?>
                {
                    ["attributes"] = attributes,
                    ["layout_name"] = ActiveLayoutName
                },
                timeout: 10000);

            if (response.Success && response.Result is List<string> results)
            {
                var summary = string.Join("\n", results);
                LastSubmitResult = $"Success: [{ActiveLayoutName}] updated";
                StatusMessage = LastSubmitResult;
                _logService.Log($"Parameters written to layout {ActiveLayoutName}", "ParamConfig");
                _logger?.LogInformation("Parameter submit results: {Results}", summary);
            }
            else
            {
                LastSubmitResult = $"Failed: {response.ErrorMessage ?? "Unknown error"}";
                StatusMessage = LastSubmitResult;
                _logService.Log($"Submit failed: {response.ErrorMessage}", "ParamConfig", AppLogLevel.Error);
            }
        }
        catch (Exception ex)
        {
            LastSubmitResult = $"Error: {ex.Message}";
            StatusMessage = LastSubmitResult;
            _logger?.LogError(ex, "Failed to submit parameters to AutoCAD");
            _logService.Log($"Submit error: {ex.Message}", "ParamConfig", AppLogLevel.Error);
        }
        finally
        {
            IsSubmitting = false;
        }
    }

    [RelayCommand]
    private void Cancel()
    {
        SelectedProduct = null;
        SelectedSpec = null;
        SelectedCategory = null;
        SelectedOperationFlow = null;
        SelectedSurfaceTreatment = null;
        SelectedColor = null;
        Unit = string.Empty;
        ColorNo = string.Empty;
        LastSubmitResult = string.Empty;
        StatusMessage = "Selections cleared";
    }
}
