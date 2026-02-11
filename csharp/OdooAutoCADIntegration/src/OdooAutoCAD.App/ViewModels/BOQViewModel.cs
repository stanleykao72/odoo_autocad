using System.Collections.ObjectModel;
using System.Net.Http;
using System.Text.Json;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.App.ViewModels;

/// <summary>
/// Enriched display model wrapping BOQEntry with validation status for DataGrid display.
/// </summary>
public partial class BOQDisplayItem : ObservableObject
{
    [ObservableProperty]
    private string _layoutName = string.Empty;

    [ObservableProperty]
    private string _productName = string.Empty;

    [ObservableProperty]
    private string _productCode = string.Empty;

    [ObservableProperty]
    private decimal _quantity;

    [ObservableProperty]
    private string _unitOfMeasure = string.Empty;

    [ObservableProperty]
    private decimal? _unitPrice;

    [ObservableProperty]
    private string _description = string.Empty;

    [ObservableProperty]
    private string _source = "AutoCAD";

    [ObservableProperty]
    private string _reference = string.Empty;

    [ObservableProperty]
    private string _validationStatus = "Pending";

    [ObservableProperty]
    private string _validationMessage = string.Empty;

    [ObservableProperty]
    private string? _detailId;

    [ObservableProperty]
    private string _position = string.Empty;

    [ObservableProperty]
    private string _width = string.Empty;

    [ObservableProperty]
    private string _height = string.Empty;

    [ObservableProperty]
    private string _length = string.Empty;

    [ObservableProperty]
    private string _thickness = string.Empty;

    public BOQEntry? OriginalEntry { get; set; }
}

/// <summary>
/// ViewModel for the BOQ (Bill of Quantities) page.
/// Handles BOQ extraction from AutoCAD, validation, push to Odoo, and ID writeback.
/// </summary>
public partial class BOQViewModel : ObservableObject
{
    private readonly IBOQProcessor _boqProcessor;
    private readonly IAutoCADService _autoCADService;
    private readonly IOdooService _odooService;
    private readonly IGUIProxy _guiProxy;
    private readonly IAppLogService _logService;
    private readonly ISettingsService _settingsService;
    private readonly ILogger<BOQViewModel>? _logger;

    /// <summary>
    /// Raw LayoutData from extraction, keyed by layout name.
    /// Used to rebuild the import2boq_v2 payload for push.
    /// </summary>
    private readonly Dictionary<string, LayoutData> _layoutDataMap = new();

    /// <summary>
    /// Tracks layouts that had writeback failures for retry.
    /// </summary>
    private readonly List<WritebackLayout> _failedWritebacks = new();

    private bool _validationCompleted;

    public ObservableCollection<BOQDisplayItem> BoqItems { get; } = new();

    // Extraction state
    [ObservableProperty]
    private bool _isExtracting;

    [ObservableProperty]
    private string _extractionStatus = string.Empty;

    [ObservableProperty]
    private double _extractionProgress;

    [ObservableProperty]
    private int _currentLayoutIndex;

    [ObservableProperty]
    private int _totalLayouts;

    // Validation state
    [ObservableProperty]
    private bool _isValidating;

    [ObservableProperty]
    private bool _hasValidationErrors;

    // Push state
    [ObservableProperty]
    private bool _isPushing;

    [ObservableProperty]
    private string _pushStatusText = string.Empty;

    [ObservableProperty]
    private string _lastPushResult = string.Empty;

    [ObservableProperty]
    private DateTime? _lastPushTime;

    // Writeback state
    [ObservableProperty]
    private bool _isWritingBack;

    // Summary
    [ObservableProperty]
    private int _totalItems;

    [ObservableProperty]
    private int _validItems;

    [ObservableProperty]
    private int _invalidItems;

    [ObservableProperty]
    private int _skippedItems;

    [ObservableProperty]
    private int _warningItems;

    [ObservableProperty]
    private int _layoutCount;

    [ObservableProperty]
    private int _skippedEmptyRows;

    [ObservableProperty]
    private int _skippedIllegalTables;

    // Connection state
    [ObservableProperty]
    private bool _isAutoCADConnected;

    [ObservableProperty]
    private bool _isOdooConnected;

    public bool HasData => TotalItems > 0;

    public bool HasNoData => TotalItems == 0;

    public string SkippedBreakdown =>
        $"Empty rows: {SkippedEmptyRows}\nIllegal tables: {SkippedIllegalTables}";

    public bool HasFailedWritebacks => _failedWritebacks.Count > 0;

    public BOQViewModel(
        IBOQProcessor boqProcessor,
        IAutoCADService autoCADService,
        IOdooService odooService,
        IGUIProxy guiProxy,
        IAppLogService logService,
        ISettingsService settingsService,
        ILogger<BOQViewModel>? logger = null)
    {
        _boqProcessor = boqProcessor;
        _autoCADService = autoCADService;
        _odooService = odooService;
        _guiProxy = guiProxy;
        _logService = logService;
        _settingsService = settingsService;
        _logger = logger;

        // Initial connection check
        RefreshConnectionStatus();
    }

    #region Property Changed Hooks

    partial void OnIsExtractingChanged(bool value)
    {
        ExtractBOQCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsAutoCADConnectedChanged(bool value)
    {
        ExtractBOQCommand.NotifyCanExecuteChanged();
    }

    partial void OnTotalItemsChanged(int value)
    {
        OnPropertyChanged(nameof(HasData));
        OnPropertyChanged(nameof(HasNoData));
    }

    partial void OnIsValidatingChanged(bool value)
    {
        ValidateCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsPushingChanged(bool value)
    {
        PushToOdooCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsWritingBackChanged(bool value)
    {
        RetryWritebackCommand.NotifyCanExecuteChanged();
    }

    partial void OnHasValidationErrorsChanged(bool value)
    {
        PushToOdooCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsOdooConnectedChanged(bool value)
    {
        PushToOdooCommand.NotifyCanExecuteChanged();
    }

    #endregion

    #region Extract Command

    [RelayCommand(CanExecute = nameof(CanExtractBOQ))]
    private async Task ExtractBOQAsync()
    {
        IsExtracting = true;
        ExtractionStatus = "Getting layouts...";
        ExtractionProgress = 0;
        BoqItems.Clear();
        _layoutDataMap.Clear();
        _failedWritebacks.Clear();
        _validationCompleted = false;
        HasValidationErrors = false;
        LastPushResult = string.Empty;
        PushStatusText = string.Empty;

        try
        {
            // Step 1: Get layouts via GUIProxy
            _logService.Log("BOQ: Getting AutoCAD layouts...", "BOQ");
            var layoutResponse = await _guiProxy.ExecuteInGuiAsync("autocad_get_layouts", null, timeout: 10000);

            // Handler returns List<LayoutInfo> — extract layout names
            List<string> layoutNames;
            if (layoutResponse.Success && layoutResponse.Result is IList<LayoutInfo> layoutInfos && layoutInfos.Count > 0)
            {
                layoutNames = layoutInfos.Select(l => l.Name).ToList();
            }
            else if (layoutResponse.Success && layoutResponse.Result is IList<string> stringLayouts && stringLayouts.Count > 0)
            {
                layoutNames = stringLayouts.ToList();
            }
            else
            {
                ExtractionStatus = "No layouts found in current drawing";
                _logService.Log("BOQ: No layouts found", "BOQ", AppLogLevel.Warning);
                return;
            }

            TotalLayouts = layoutNames.Count;
            LayoutCount = layoutNames.Count;
            _logService.Log($"BOQ: Found {layoutNames.Count} layouts", "BOQ");

            var allItems = new List<BOQDisplayItem>();

            // Step 2: For each layout, extract parameters
            for (var i = 0; i < layoutNames.Count; i++)
            {
                var layoutName = layoutNames[i];
                CurrentLayoutIndex = i + 1;
                ExtractionStatus = $"Extracting layout {CurrentLayoutIndex}/{TotalLayouts}: {layoutName}";
                ExtractionProgress = (double)i / layoutNames.Count;

                _logService.Log($"BOQ: Extracting layout '{layoutName}'...", "BOQ");

                var extractParams = new Dictionary<string, object?>
                {
                    ["layoutName"] = layoutName
                };

                var extractResponse = await _guiProxy.ExecuteInGuiAsync(
                    "autocad_extract_parameters", extractParams, timeout: 30000);

                if (!extractResponse.Success)
                {
                    _logService.Log($"BOQ: Failed to extract '{layoutName}': {extractResponse.ErrorMessage}", "BOQ", AppLogLevel.Warning);
                    continue;
                }

                if (extractResponse.Result is LayoutData layoutData)
                {
                    // Store raw LayoutData for push payload
                    _layoutDataMap[layoutName] = layoutData;

                    // Generate BOQ entries from layout data
                    var options = new BOQGenerationOptions
                    {
                        IncludeAutoCADData = true,
                        ValidateProducts = _odooService.IsConnected
                    };

                    var result = await _boqProcessor.GenerateBOQAsync(layoutData, options);

                    // Build display items, also populating detail fields from table data
                    int detailIndex = 0;
                    foreach (var entry in result.Entries)
                    {
                        var displayItem = new BOQDisplayItem
                        {
                            LayoutName = layoutName,
                            ProductName = entry.ProductName,
                            ProductCode = string.Empty,
                            Quantity = entry.Quantity,
                            UnitOfMeasure = entry.UnitOfMeasure,
                            UnitPrice = entry.UnitPrice,
                            Description = entry.Description ?? string.Empty,
                            Reference = entry.Reference ?? string.Empty,
                            Source = "AutoCAD",
                            ValidationStatus = "Pending",
                            ValidationMessage = string.Empty,
                            OriginalEntry = entry
                        };

                        // Populate detail fields from table data if available
                        if (layoutData.Tables.Count > 0)
                        {
                            var table = layoutData.Tables[0];
                            // detailIndex+1 because row 0 is header in Cells
                            int dataRow = detailIndex + 1;
                            if (dataRow < table.Cells.Count)
                            {
                                var cells = table.Cells[dataRow];
                                displayItem.Position = cells.Count > 0 ? cells[0] : "";
                                displayItem.ProductCode = cells.Count > 1 ? cells[1] : "";
                                displayItem.Width = cells.Count > 2 ? cells[2] : "";
                                displayItem.Height = cells.Count > 3 ? cells[3] : "";
                                displayItem.Length = cells.Count > 4 ? cells[4] : "";
                                displayItem.Thickness = cells.Count > 5 ? cells[5] : "";
                                displayItem.DetailId = cells.Count > 8 ? cells[8] : null;
                            }
                        }

                        allItems.Add(displayItem);
                        detailIndex++;
                    }

                    foreach (var warning in result.Warnings)
                    {
                        _logService.Log($"BOQ Warning: {warning}", "BOQ", AppLogLevel.Warning);
                    }
                }
            }

            // Populate collection
            foreach (var item in allItems)
            {
                BoqItems.Add(item);
            }

            ExtractionProgress = 1.0;
            UpdateSummary();

            ExtractionStatus = $"Extraction complete: {TotalItems} items from {LayoutCount} layouts";
            _logService.Log($"BOQ: Extraction complete — {TotalItems} items from {LayoutCount} layouts", "BOQ");
        }
        catch (Exception ex)
        {
            ExtractionStatus = $"Extraction failed: {ex.Message}";
            _logService.Log($"BOQ: Extraction failed — {ex.Message}", "BOQ", AppLogLevel.Error);
            _logger?.LogError(ex, "BOQ extraction error");
        }
        finally
        {
            IsExtracting = false;
        }
    }

    private bool CanExtractBOQ() => IsAutoCADConnected && !IsExtracting;

    #endregion

    #region Validate Command

    [RelayCommand(CanExecute = nameof(CanValidate))]
    private async Task ValidateAsync()
    {
        IsValidating = true;
        _logService.Log("BOQ: Starting validation...", "BOQ");

        try
        {
            // Reset all items to Pending
            foreach (var item in BoqItems)
            {
                item.ValidationStatus = "Pending";
                item.ValidationMessage = string.Empty;
            }

            int validCount = 0, warningCount = 0, errorCount = 0;

            foreach (var item in BoqItems)
            {
                // Check product mapping
                bool productValid = true;
                if (!string.IsNullOrWhiteSpace(item.ProductCode))
                {
                    var product = await _boqProcessor.MapProductAsync(item.ProductCode);
                    if (product == null && !string.IsNullOrWhiteSpace(item.ProductName))
                    {
                        product = await _boqProcessor.MapProductAsync(item.ProductName);
                    }
                    if (product == null)
                    {
                        item.ValidationStatus = "Error";
                        item.ValidationMessage = "Product not found in Odoo";
                        productValid = false;
                        errorCount++;
                    }
                }
                else if (!string.IsNullOrWhiteSpace(item.ProductName))
                {
                    var product = await _boqProcessor.MapProductAsync(item.ProductName);
                    if (product == null)
                    {
                        item.ValidationStatus = "Error";
                        item.ValidationMessage = "Product not found in Odoo";
                        productValid = false;
                        errorCount++;
                    }
                }
                else
                {
                    item.ValidationStatus = "Error";
                    item.ValidationMessage = "Missing product name/code";
                    productValid = false;
                    errorCount++;
                }

                if (!productValid) continue;

                // Check quantity
                if (item.Quantity < 0)
                {
                    item.ValidationStatus = "Error";
                    item.ValidationMessage = "Negative quantity";
                    errorCount++;
                }
                else if (item.Quantity == 0)
                {
                    item.ValidationStatus = "Warning";
                    item.ValidationMessage = "Zero quantity";
                    warningCount++;
                }
                else
                {
                    item.ValidationStatus = "Valid";
                    item.ValidationMessage = string.Empty;
                    validCount++;
                }
            }

            HasValidationErrors = errorCount > 0;
            _validationCompleted = true;
            UpdateSummary();

            PushToOdooCommand.NotifyCanExecuteChanged();

            _logService.Log(
                $"BOQ: Validation complete — {validCount} valid, {warningCount} warnings, {errorCount} errors",
                "BOQ");
        }
        catch (Exception ex)
        {
            _logService.Log($"BOQ: Validation failed — {ex.Message}", "BOQ", AppLogLevel.Error);
            _logger?.LogError(ex, "BOQ validation error");
        }
        finally
        {
            IsValidating = false;
        }
    }

    private bool CanValidate() => !IsValidating && TotalItems > 0;

    #endregion

    #region Push to Odoo Command

    [RelayCommand(CanExecute = nameof(CanPushToOdoo))]
    private async Task PushToOdooAsync()
    {
        IsPushing = true;
        PushStatusText = "Preparing push...";
        LastPushResult = string.Empty;
        _failedWritebacks.Clear();

        try
        {
            // Load Swagger URL + user token from settings
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
                PushStatusText = "Missing Swagger URL or user token in settings";
                LastPushResult = "Failed: Missing configuration";
                return;
            }

            // Parse swagger URL
            var parsed = OdooConnectionViewModel.ParseSwaggerUrl(swaggerUrl);
            if (parsed == null)
            {
                PushStatusText = "Invalid Swagger URL configuration";
                LastPushResult = "Failed: Invalid Swagger URL";
                return;
            }

            var (baseUrl, database, apiToken) = parsed.Value;

            // Fetch basePath from swagger spec (best-effort, fallback to default)
            var basePath = "/api/v1/boq_import_api";
            PushStatusText = "Fetching API configuration...";
            try
            {
                using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
                var specJson = await httpClient.GetStringAsync(swaggerUrl);
                var doc = JsonSerializer.Deserialize<JsonElement>(specJson);
                if (doc.TryGetProperty("basePath", out var bp))
                    basePath = bp.GetString() ?? basePath;
            }
            catch
            {
                // Use default basePath
            }

            // Build BoqImportRequest from layout data
            PushStatusText = "Building import request...";
            var request = BuildImportRequest();

            if (request.All.Count == 0)
            {
                PushStatusText = "No data to push";
                LastPushResult = "Failed: No layout data available";
                return;
            }

            // Call API
            PushStatusText = $"Pushing {request.All.Count} layouts to Odoo...";
            _logService.Log($"BOQ: Pushing {request.All.Count} layouts via import2boq_v2...", "BOQ");

            var response = await _odooService.ImportToBOQViaApiAsync(
                request, baseUrl, basePath, database, userToken);

            if (!response.Success)
            {
                PushStatusText = $"Push failed: {response.ErrorMessage}";
                LastPushResult = $"Failed: {response.ErrorCode} — {response.ErrorMessage}";
                _logService.Log($"BOQ: Push failed — {response.ErrorCode}: {response.ErrorMessage}", "BOQ", AppLogLevel.Error);
                return;
            }

            // Build writeback data from response
            var writebackData = BuildWritebackData(response);

            // Execute writeback to AutoCAD
            PushStatusText = "Writing IDs back to AutoCAD...";
            await ExecuteWritebackAsync(writebackData);

            // Update display items with returned IDs
            UpdateDisplayItemIds(response);

            LastPushTime = DateTime.Now;
            var failedCount = _failedWritebacks.Count;
            if (failedCount > 0)
            {
                PushStatusText = $"Push complete with {failedCount} writeback failure(s)";
                LastPushResult = $"Push OK — {response.All?.Count ?? 0} layouts. Writeback: {failedCount} failed.";
            }
            else
            {
                PushStatusText = "Push and writeback complete";
                LastPushResult = $"Success — {response.All?.Count ?? 0} layouts pushed and IDs written back";
            }

            OnPropertyChanged(nameof(HasFailedWritebacks));
            RetryWritebackCommand.NotifyCanExecuteChanged();

            _logService.Log($"BOQ: Push complete — {response.All?.Count ?? 0} layouts, {failedCount} writeback failures", "BOQ");
        }
        catch (Exception ex)
        {
            PushStatusText = $"Push failed: {ex.Message}";
            LastPushResult = $"Failed: {ex.Message}";
            _logService.Log($"BOQ: Push failed — {ex.Message}", "BOQ", AppLogLevel.Error);
            _logger?.LogError(ex, "BOQ push error");
        }
        finally
        {
            IsPushing = false;
        }
    }

    private bool CanPushToOdoo() =>
        !HasValidationErrors && !IsPushing && TotalItems > 0 &&
        IsOdooConnected && _validationCompleted;

    internal BoqImportRequest BuildImportRequest()
    {
        var request = new BoqImportRequest();

        foreach (var (layoutName, layoutData) in _layoutDataMap)
        {
            var importLayout = new BoqImportLayout
            {
                LayoutName = layoutName,
                // Pull header fields from layout parameters
                PrNo = GetParam(layoutData, "pr_no"),
                ProjectName = GetParam(layoutData, "project_name"),
                JobWorkingPlanName = GetParam(layoutData, "job_working_plan_name"),
                ProductName = GetParam(layoutData, "product_name"),
                ProductCatalog = GetParam(layoutData, "product_catalog"),
                Spec = GetParam(layoutData, "spec"),
                SurfaceTreatment = GetParam(layoutData, "surface_treatment"),
                OperationFlow = GetParam(layoutData, "operation_flow"),
                ColorName = GetParam(layoutData, "color_name"),
                ColorNo = GetParam(layoutData, "color_no")
            };

            // Build detail rows from table data
            if (layoutData.Tables.Count > 0)
            {
                var table = layoutData.Tables[0];
                // Skip header row (index 0)
                for (int i = 1; i < table.Cells.Count; i++)
                {
                    var cells = table.Cells[i];
                    importLayout.Detail.Add(new BoqImportDetail
                    {
                        Position = cells.Count > 0 ? cells[0] : "",
                        ProductNo = cells.Count > 1 ? cells[1] : "",
                        Width = cells.Count > 2 ? cells[2] : "",
                        Height = cells.Count > 3 ? cells[3] : "",
                        Length = cells.Count > 4 ? cells[4] : "",
                        Thickness = cells.Count > 5 ? cells[5] : "",
                        Qty = cells.Count > 6 ? cells[6] : "",
                        Description = cells.Count > 7 ? cells[7] : "",
                        DetailId = cells.Count > 8 && !string.IsNullOrWhiteSpace(cells[8]) ? cells[8] : null
                    });
                }
            }

            request.All.Add(importLayout);
        }

        return request;
    }

    private static string GetParam(LayoutData data, string key)
    {
        return data.Parameters.TryGetValue(key, out var val) ? val?.ToString() ?? "" : "";
    }

    internal static List<WritebackLayout> BuildWritebackData(BoqImportResponse response)
    {
        var result = new List<WritebackLayout>();
        if (response.All == null) return result;

        foreach (var layout in response.All)
        {
            var wb = new WritebackLayout
            {
                LayoutName = layout.LayoutName,
                HeaderId = layout.HeaderId
            };

            foreach (var detail in layout.Detail)
            {
                wb.Details.Add(new WritebackDetail
                {
                    ProductNo = detail.ProductNo,
                    DetailId = detail.DetailId
                });
            }

            result.Add(wb);
        }

        return result;
    }

    private async Task ExecuteWritebackAsync(List<WritebackLayout> writebackData)
    {
        IsWritingBack = true;
        try
        {
            foreach (var wb in writebackData)
            {
                try
                {
                    var parameters = new Dictionary<string, object?>
                    {
                        ["layout_name"] = wb.LayoutName,
                        ["header_id"] = wb.HeaderId ?? "",
                        ["details"] = wb.Details
                    };

                    var result = await _guiProxy.ExecuteInGuiAsync(
                        "autocad_write_table_ids", parameters, timeout: 30000);

                    if (!result.Success)
                    {
                        _failedWritebacks.Add(wb);
                        _logService.Log(
                            $"BOQ: Writeback failed for '{wb.LayoutName}': {result.ErrorMessage}",
                            "BOQ", AppLogLevel.Warning);
                    }
                }
                catch (Exception ex)
                {
                    _failedWritebacks.Add(wb);
                    _logService.Log(
                        $"BOQ: Writeback error for '{wb.LayoutName}': {ex.Message}",
                        "BOQ", AppLogLevel.Warning);
                }
            }
        }
        finally
        {
            IsWritingBack = false;
        }
    }

    private void UpdateDisplayItemIds(BoqImportResponse response)
    {
        if (response.All == null) return;

        foreach (var layout in response.All)
        {
            var items = BoqItems.Where(i =>
                string.Equals(i.LayoutName, layout.LayoutName, StringComparison.OrdinalIgnoreCase));

            foreach (var detail in layout.Detail)
            {
                // Find matching item by product_no (case-insensitive)
                var match = items.FirstOrDefault(i =>
                    string.Equals(i.ProductCode, detail.ProductNo, StringComparison.OrdinalIgnoreCase));
                if (match != null && detail.DetailId != null)
                {
                    match.DetailId = detail.DetailId;
                }
            }
        }
    }

    #endregion

    #region Retry Writeback Command

    [RelayCommand(CanExecute = nameof(CanRetryWriteback))]
    private async Task RetryWritebackAsync()
    {
        if (_failedWritebacks.Count == 0) return;

        var retryList = new List<WritebackLayout>(_failedWritebacks);
        _failedWritebacks.Clear();

        PushStatusText = $"Retrying writeback for {retryList.Count} layout(s)...";
        _logService.Log($"BOQ: Retrying writeback for {retryList.Count} layouts...", "BOQ");

        await ExecuteWritebackAsync(retryList);

        var stillFailed = _failedWritebacks.Count;
        PushStatusText = stillFailed > 0
            ? $"Retry complete: {stillFailed} still failing"
            : "All writebacks successful";

        OnPropertyChanged(nameof(HasFailedWritebacks));
        RetryWritebackCommand.NotifyCanExecuteChanged();
    }

    private bool CanRetryWriteback() => HasFailedWritebacks && !IsWritingBack;

    #endregion

    #region Refresh + Summary

    [RelayCommand]
    private void RefreshConnectionStatus()
    {
        IsAutoCADConnected = _autoCADService.IsConnected;
        IsOdooConnected = _odooService.IsConnected;
    }

    internal void UpdateSummary()
    {
        TotalItems = BoqItems.Count;
        ValidItems = BoqItems.Count(i => i.ValidationStatus == "Valid");
        InvalidItems = BoqItems.Count(i => i.ValidationStatus is "Invalid" or "Error");
        WarningItems = BoqItems.Count(i => i.ValidationStatus == "Warning");
        SkippedItems = BoqItems.Count(i => i.ValidationStatus == "Skipped");
        LayoutCount = BoqItems.Select(i => i.LayoutName).Distinct().Count();
        OnPropertyChanged(nameof(HasData));
        OnPropertyChanged(nameof(HasNoData));
        OnPropertyChanged(nameof(SkippedBreakdown));
    }

    /// <summary>
    /// Find a detail_id for a given product_no (case-insensitive).
    /// Used by writeback logic and tests.
    /// </summary>
    internal static string? FindDetailId(IEnumerable<WritebackDetail> details, string productNo)
    {
        return details.FirstOrDefault(d =>
            string.Equals(d.ProductNo, productNo, StringComparison.OrdinalIgnoreCase))?.DetailId;
    }

    #endregion
}
