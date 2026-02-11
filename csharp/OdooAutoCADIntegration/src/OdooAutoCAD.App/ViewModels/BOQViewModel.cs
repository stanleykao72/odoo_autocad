using System.Collections.ObjectModel;
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

    public BOQEntry? OriginalEntry { get; set; }
}

/// <summary>
/// ViewModel for the BOQ (Bill of Quantities) page.
/// Handles BOQ extraction from AutoCAD, review, and display.
/// </summary>
public partial class BOQViewModel : ObservableObject
{
    private readonly IBOQProcessor _boqProcessor;
    private readonly IAutoCADService _autoCADService;
    private readonly IOdooService _odooService;
    private readonly IGUIProxy _guiProxy;
    private readonly IAppLogService _logService;
    private readonly ILogger<BOQViewModel>? _logger;

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
    private int _layoutCount;

    // Connection state
    [ObservableProperty]
    private bool _isAutoCADConnected;

    [ObservableProperty]
    private bool _isOdooConnected;

    public bool HasData => TotalItems > 0;

    public BOQViewModel(
        IBOQProcessor boqProcessor,
        IAutoCADService autoCADService,
        IOdooService odooService,
        IGUIProxy guiProxy,
        IAppLogService logService,
        ILogger<BOQViewModel>? logger = null)
    {
        _boqProcessor = boqProcessor;
        _autoCADService = autoCADService;
        _odooService = odooService;
        _guiProxy = guiProxy;
        _logService = logService;
        _logger = logger;

        // Initial connection check
        RefreshConnectionStatus();
    }

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
    }

    [RelayCommand(CanExecute = nameof(CanExtractBOQ))]
    private async Task ExtractBOQAsync()
    {
        IsExtracting = true;
        ExtractionStatus = "Getting layouts...";
        ExtractionProgress = 0;
        BoqItems.Clear();

        try
        {
            // Step 1: Get layouts via GUIProxy
            _logService.Log("BOQ: Getting AutoCAD layouts...", "BOQ");
            var layoutResponse = await _guiProxy.ExecuteInGuiAsync("autocad_get_layouts", null, timeout: 10000);

            if (!layoutResponse.Success || layoutResponse.Result is not List<string> layouts || layouts.Count == 0)
            {
                ExtractionStatus = "No layouts found in current drawing";
                _logService.Log("BOQ: No layouts found", "BOQ", AppLogLevel.Warning);
                return;
            }

            TotalLayouts = layouts.Count;
            LayoutCount = layouts.Count;
            _logService.Log($"BOQ: Found {layouts.Count} layouts", "BOQ");

            var allItems = new List<BOQDisplayItem>();

            // Step 2: For each layout, extract parameters
            for (var i = 0; i < layouts.Count; i++)
            {
                var layoutName = layouts[i];
                CurrentLayoutIndex = i + 1;
                ExtractionStatus = $"Extracting layout {CurrentLayoutIndex}/{TotalLayouts}: {layoutName}";
                ExtractionProgress = (double)i / layouts.Count;

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
                    // Generate BOQ entries from layout data
                    var options = new BOQGenerationOptions
                    {
                        IncludeAutoCADData = true,
                        ValidateProducts = _odooService.IsConnected
                    };

                    var result = await _boqProcessor.GenerateBOQAsync(layoutData, options);

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
                            ValidationStatus = "Valid",
                            ValidationMessage = string.Empty,
                            OriginalEntry = entry
                        };
                        allItems.Add(displayItem);
                    }

                    // Mark skipped/invalid from warnings
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
            _logService.Log($"BOQ: Extraction complete — {TotalItems} items, {ValidItems} valid, {InvalidItems} invalid, {SkippedItems} skipped", "BOQ");
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
        InvalidItems = BoqItems.Count(i => i.ValidationStatus == "Invalid");
        SkippedItems = BoqItems.Count(i => i.ValidationStatus == "Skipped");
        LayoutCount = BoqItems.Select(i => i.LayoutName).Distinct().Count();
    }
}
