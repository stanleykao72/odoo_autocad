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
/// Display model for a Purchase Requisition in the DataGrid.
/// </summary>
public partial class PRDisplayItem : ObservableObject
{
    [ObservableProperty]
    private int _id;

    [ObservableProperty]
    private string _reference = string.Empty;

    [ObservableProperty]
    private string _state = "draft";

    [ObservableProperty]
    private int _lineCount;

    [ObservableProperty]
    private DateTime _createdAt = DateTime.Now;

    [ObservableProperty]
    private bool _isSelected;

    public PREntry? OriginalEntry { get; set; }
}

/// <summary>
/// Display model for a PR line item in the detail view DataGrid.
/// </summary>
public class PRLineDisplayItem
{
    public string ProductName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public string UnitOfMeasure { get; set; } = string.Empty;
    public decimal? UnitPrice { get; set; }
    public decimal LineTotal => Quantity * (UnitPrice ?? 0);
}

/// <summary>
/// ViewModel for the Purchase Requisition page.
/// Handles BOQ-to-PR conversion via Swagger, PR listing, and submission for approval.
/// </summary>
public partial class PurchaseRequisitionViewModel : ObservableObject
{
    private readonly IAutoCADService _autoCADService;
    private readonly IOdooService _odooService;
    private readonly IGUIProxy _guiProxy;
    private readonly IAppLogService _logService;
    private readonly ISettingsService _settingsService;
    private readonly ILogger<PurchaseRequisitionViewModel>? _logger;

    public ObservableCollection<PRDisplayItem> PrItems { get; } = new();
    public ObservableCollection<PRDisplayItem> FilteredPRs { get; } = new();
    public ObservableCollection<PRLineDisplayItem> SelectedPRLines { get; } = new();

    // Filter/Sort state
    [ObservableProperty]
    private string _selectedStateFilter = "All";

    [ObservableProperty]
    private string _selectedSortOrder = "Newest First";

    // Detail view
    [ObservableProperty]
    private decimal _selectedPRTotal;

    // Feedback enhancements
    [ObservableProperty]
    private string _statusMessageType = string.Empty;

    [ObservableProperty]
    private DateTime? _lastConversionTime;

    // Connection state
    [ObservableProperty]
    private bool _isAutoCADConnected;

    [ObservableProperty]
    private bool _isOdooConnected;

    // Convert state
    [ObservableProperty]
    private bool _isConverting;

    [ObservableProperty]
    private string _convertStatusText = string.Empty;

    [ObservableProperty]
    private string _lastConvertResult = string.Empty;

    // List state
    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private string _loadStatusText = string.Empty;

    // Submit state
    [ObservableProperty]
    private bool _isSubmitting;

    [ObservableProperty]
    private string _submitStatusText = string.Empty;

    // Selection
    [ObservableProperty]
    private PRDisplayItem? _selectedPR;

    // Summary
    [ObservableProperty]
    private int _totalPRs;

    [ObservableProperty]
    private int _draftCount;

    [ObservableProperty]
    private int _submittedCount;

    [ObservableProperty]
    private int _approvedCount;

    public bool HasData => TotalPRs > 0;

    public bool HasNoData => TotalPRs == 0;

    public PurchaseRequisitionViewModel(
        IAutoCADService autoCADService,
        IOdooService odooService,
        IGUIProxy guiProxy,
        IAppLogService logService,
        ISettingsService settingsService,
        ILogger<PurchaseRequisitionViewModel>? logger = null)
    {
        _autoCADService = autoCADService;
        _odooService = odooService;
        _guiProxy = guiProxy;
        _logService = logService;
        _settingsService = settingsService;
        _logger = logger;

        RefreshConnectionStatus();
    }

    #region Property Changed Hooks

    partial void OnIsConvertingChanged(bool value)
    {
        ConvertBOQToPRCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsAutoCADConnectedChanged(bool value)
    {
        ConvertBOQToPRCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsOdooConnectedChanged(bool value)
    {
        ConvertBOQToPRCommand.NotifyCanExecuteChanged();
        RefreshPRListCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsLoadingChanged(bool value)
    {
        RefreshPRListCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsSubmittingChanged(bool value)
    {
        SubmitPRCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedPRChanged(PRDisplayItem? value)
    {
        SubmitPRCommand.NotifyCanExecuteChanged();
        PopulateSelectedPRLines();
    }

    partial void OnSelectedStateFilterChanged(string value)
    {
        ApplyFilterAndSort();
    }

    partial void OnSelectedSortOrderChanged(string value)
    {
        ApplyFilterAndSort();
    }

    partial void OnTotalPRsChanged(int value)
    {
        OnPropertyChanged(nameof(HasData));
        OnPropertyChanged(nameof(HasNoData));
    }

    #endregion

    #region Convert BOQ to PR Command

    [RelayCommand(CanExecute = nameof(CanConvertBOQToPR))]
    private async Task ConvertBOQToPRAsync()
    {
        IsConverting = true;
        ConvertStatusText = "Extracting header IDs from AutoCAD...";
        LastConvertResult = string.Empty;

        try
        {
            // Step 1: Extract header_ids via GUIProxy
            _logService.Log("PR: Getting header IDs from AutoCAD tables...", "PR");
            var headerResponse = await _guiProxy.ExecuteInGuiAsync(
                "autocad_get_header_ids", null, timeout: 15000);

            List<string> headerIds;
            if (headerResponse.Success && headerResponse.Result is IList<string> ids && ids.Count > 0)
            {
                headerIds = ids.ToList();
            }
            else
            {
                ConvertStatusText = "No header IDs found in AutoCAD tables";
                LastConvertResult = "Failed: No header IDs found. Ensure BOQ has been pushed to Odoo first.";
                _logService.Log("PR: No header IDs found", "PR", AppLogLevel.Warning);
                return;
            }

            _logService.Log($"PR: Found {headerIds.Count} header IDs", "PR");

            // Step 2: Load Swagger credentials from settings
            ConvertStatusText = "Loading configuration...";
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
                ConvertStatusText = "Missing Swagger URL or user token in settings";
                LastConvertResult = "Failed: Missing configuration. Check Settings > Connection.";
                return;
            }

            // Step 3: Parse swagger URL
            var parsed = OdooConnectionViewModel.ParseSwaggerUrl(swaggerUrl);
            if (parsed == null)
            {
                ConvertStatusText = "Invalid Swagger URL configuration";
                LastConvertResult = "Failed: Invalid Swagger URL";
                return;
            }

            var (baseUrl, database, apiToken) = parsed.Value;

            // Step 4: Resolve endpoint path from swagger spec
            var endpointPath = "/api/v1/boq_import_api/callMethodForJobWorkingPlanBoqModel";
            ConvertStatusText = "Fetching API configuration...";
            try
            {
                using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
                var specJson = await httpClient.GetStringAsync(swaggerUrl);
                endpointPath = OdooConnectionViewModel.ResolveSwaggerEndpoint(specJson);
            }
            catch
            {
                // Use default endpoint path
            }

            // Step 5: Call boq2pr_v2
            ConvertStatusText = $"Converting {headerIds.Count} BOQ entries to PR...";
            _logService.Log($"PR: Calling boq2pr_v2 with {headerIds.Count} header IDs...", "PR");

            var response = await _odooService.ConvertBOQToPRViaApiAsync(
                headerIds, baseUrl, endpointPath, database, userToken);

            if (!response.Success)
            {
                ConvertStatusText = $"Conversion failed: {response.ErrorMessage}";
                LastConvertResult = $"Failed: {response.ErrorCode} — {response.ErrorMessage}";
                StatusMessageType = "error";
                _logService.Log($"PR: Conversion failed — {response.ErrorCode}: {response.ErrorMessage}", "PR", AppLogLevel.Error);
                return;
            }

            // Step 6: Populate PrItems from response
            PrItems.Clear();
            if (response.All != null)
            {
                foreach (var pr in response.All)
                {
                    var entry = new PREntry
                    {
                        Id = pr.PrId ?? 0,
                        Reference = pr.Reference,
                        State = pr.State,
                        CreatedAt = DateTime.Now
                    };
                    foreach (var line in pr.Lines)
                    {
                        entry.Lines.Add(new PRLine
                        {
                            ProductId = line.ProductId,
                            ProductName = line.ProductName,
                            Quantity = line.Quantity,
                            UnitOfMeasure = line.UnitOfMeasure,
                            UnitPrice = line.UnitPrice
                        });
                    }

                    PrItems.Add(new PRDisplayItem
                    {
                        Id = pr.PrId ?? 0,
                        Reference = pr.Reference,
                        State = pr.State,
                        LineCount = pr.Lines.Count,
                        CreatedAt = DateTime.Now,
                        OriginalEntry = entry
                    });
                }
            }

            UpdateSummary();

            if (PrItems.Count == 0)
            {
                ConvertStatusText = "Conversion returned no PRs";
                LastConvertResult = "No Purchase Requisitions were created from the selected BOQ entries";
                StatusMessageType = "info";
                _logService.Log("PR: Conversion returned 0 PRs", "PR", AppLogLevel.Warning);
            }
            else
            {
                var firstRef = PrItems.Count > 0 ? PrItems[0].Reference : "";
                ConvertStatusText = $"Conversion complete: {PrItems.Count} PR(s) created";
                LastConvertResult = $"Success — {PrItems.Count} Purchase Requisition(s) created (e.g. {firstRef})";
                StatusMessageType = "success";
                LastConversionTime = DateTime.Now;
                _logService.Log($"PR: Conversion complete — {PrItems.Count} PRs created", "PR");
            }
        }
        catch (TaskCanceledException)
        {
            ConvertStatusText = "Conversion timed out";
            LastConvertResult = "Failed: Operation timed out. Check network connection and try again.";
            StatusMessageType = "error";
            _logService.Log("PR: Conversion timed out", "PR", AppLogLevel.Error);
        }
        catch (Exception ex)
        {
            ConvertStatusText = $"Conversion failed: {ex.Message}";
            LastConvertResult = $"Failed: {ex.Message}";
            StatusMessageType = "error";
            _logService.Log($"PR: Conversion failed — {ex.Message}", "PR", AppLogLevel.Error);
            _logger?.LogError(ex, "BOQ-to-PR conversion error");
        }
        finally
        {
            IsConverting = false;
        }
    }

    private bool CanConvertBOQToPR() => IsAutoCADConnected && IsOdooConnected && !IsConverting;

    #endregion

    #region Refresh PR List Command

    [RelayCommand(CanExecute = nameof(CanRefreshPRList))]
    private async Task RefreshPRListAsync()
    {
        IsLoading = true;
        LoadStatusText = "Loading PRs from Odoo...";

        try
        {
            // Read project ID from settings
            var configs = await _settingsService.LoadServerConfigsAsync();
            configs.TryGetValue("current_project_id", out var projectIdStr);

            if (string.IsNullOrEmpty(projectIdStr) || !int.TryParse(projectIdStr, out var projectId))
            {
                LoadStatusText = "No project ID configured. Check Settings.";
                _logService.Log("PR: No project ID found in settings", "PR", AppLogLevel.Warning);
                return;
            }

            _logService.Log($"PR: Refreshing PRs for project {projectId}...", "PR");

            var prEntries = await _odooService.GetPurchaseRequisitionsAsync(projectId);

            PrItems.Clear();
            foreach (var entry in prEntries)
            {
                PrItems.Add(new PRDisplayItem
                {
                    Id = entry.Id ?? 0,
                    Reference = entry.Reference,
                    State = entry.State,
                    LineCount = entry.Lines.Count,
                    CreatedAt = entry.CreatedAt,
                    OriginalEntry = entry
                });
            }

            UpdateSummary();

            LoadStatusText = $"Loaded {PrItems.Count} PR(s)";
            _logService.Log($"PR: Loaded {PrItems.Count} PRs", "PR");
        }
        catch (Exception ex)
        {
            LoadStatusText = $"Failed to load PRs: {ex.Message}";
            _logService.Log($"PR: Failed to load PRs — {ex.Message}", "PR", AppLogLevel.Error);
            _logger?.LogError(ex, "PR list refresh error");
        }
        finally
        {
            IsLoading = false;
        }
    }

    private bool CanRefreshPRList() => IsOdooConnected && !IsLoading;

    #endregion

    #region Submit PR Command

    [RelayCommand(CanExecute = nameof(CanSubmitPR))]
    private async Task SubmitPRAsync()
    {
        if (SelectedPR == null) return;

        IsSubmitting = true;
        SubmitStatusText = $"Submitting {SelectedPR.Reference} for approval...";

        try
        {
            _logService.Log($"PR: Submitting {SelectedPR.Reference} (ID: {SelectedPR.Id})...", "PR");

            var success = await _odooService.SubmitPRAsync(SelectedPR.Id);

            if (success)
            {
                SelectedPR.State = "submitted";
                UpdateSummary();
                SubmitStatusText = $"{SelectedPR.Reference} submitted for approval";
                _logService.Log($"PR: {SelectedPR.Reference} submitted successfully", "PR");
            }
            else
            {
                SubmitStatusText = $"Failed to submit {SelectedPR.Reference}";
                _logService.Log($"PR: Failed to submit {SelectedPR.Reference}", "PR", AppLogLevel.Error);
            }
        }
        catch (Exception ex)
        {
            SubmitStatusText = $"Submit failed: {ex.Message}";
            _logService.Log($"PR: Submit failed — {ex.Message}", "PR", AppLogLevel.Error);
            _logger?.LogError(ex, "PR submit error");
        }
        finally
        {
            IsSubmitting = false;
        }
    }

    private bool CanSubmitPR() =>
        SelectedPR != null &&
        SelectedPR.State == "draft" &&
        !IsSubmitting;

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
        TotalPRs = PrItems.Count;
        DraftCount = PrItems.Count(i => i.State == "draft");
        SubmittedCount = PrItems.Count(i => i.State is "submitted" or "sent");
        ApprovedCount = PrItems.Count(i => i.State is "approved" or "done");
        OnPropertyChanged(nameof(HasData));
        OnPropertyChanged(nameof(HasNoData));
        ApplyFilterAndSort();
    }

    internal void ApplyFilterAndSort()
    {
        IEnumerable<PRDisplayItem> items = PrItems;

        // Filter by state
        if (!string.IsNullOrEmpty(SelectedStateFilter) && SelectedStateFilter != "All")
        {
            var filter = SelectedStateFilter.ToLowerInvariant();
            items = items.Where(i => string.Equals(i.State, filter, StringComparison.OrdinalIgnoreCase));
        }

        // Sort
        items = SelectedSortOrder switch
        {
            "Oldest First" => items.OrderBy(i => i.CreatedAt),
            "Reference A-Z" => items.OrderBy(i => i.Reference, StringComparer.OrdinalIgnoreCase),
            "Reference Z-A" => items.OrderByDescending(i => i.Reference, StringComparer.OrdinalIgnoreCase),
            _ => items.OrderByDescending(i => i.CreatedAt) // "Newest First" default
        };

        FilteredPRs.Clear();
        foreach (var item in items)
        {
            FilteredPRs.Add(item);
        }
    }

    private void PopulateSelectedPRLines()
    {
        SelectedPRLines.Clear();
        SelectedPRTotal = 0;

        if (SelectedPR?.OriginalEntry?.Lines == null) return;

        foreach (var line in SelectedPR.OriginalEntry.Lines)
        {
            SelectedPRLines.Add(new PRLineDisplayItem
            {
                ProductName = line.ProductName,
                Quantity = line.Quantity,
                UnitOfMeasure = line.UnitOfMeasure,
                UnitPrice = line.UnitPrice
            });
        }

        SelectedPRTotal = SelectedPRLines.Sum(l => l.LineTotal);
    }

    #endregion
}
