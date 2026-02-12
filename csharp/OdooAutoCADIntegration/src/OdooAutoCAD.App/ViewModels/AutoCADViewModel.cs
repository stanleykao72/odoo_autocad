// OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs
// ViewModel for AutoCAD connection management

using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.Core.AutoCAD;
using OdooAutoCAD.Core.Odoo;
using OdooAutoCAD.Core.Threading;

namespace OdooAutoCAD.App.ViewModels;

/// <summary>
/// ViewModel for AutoCAD connection page.
/// Manages connection status, AutoCAD information display, and connection commands.
/// All AutoCAD COM operations are executed through IGUIProxy for thread safety.
/// </summary>
public partial class AutoCADViewModel : ObservableObject
{
    private readonly IAutoCADService _autoCADService;
    private readonly IDwgReaderService _dwgReader;
    private readonly IGUIProxy _guiProxy;
    private readonly IOdooService _odooService;
    private readonly ISettingsService _settingsService;
    private readonly IAppLogService _logService;
    private readonly ILogger<AutoCADViewModel>? _logger;
    private DispatcherTimer? _healthTimer;

    [ObservableProperty]
    private bool _isConnected;

    [ObservableProperty]
    private string _statusMessage = "Click Connect to establish connection with AutoCAD";

    [ObservableProperty]
    private string _autoCADVersion = "N/A";

    [ObservableProperty]
    private string _currentDocument = "No document open";

    [ObservableProperty]
    private string _connectButtonText = "Connect to AutoCAD";

    [ObservableProperty]
    private bool _isConnecting;

    // Sprint 8: Drawing Info + COM Monitor
    [ObservableProperty]
    private string _documentPath = "N/A";

    [ObservableProperty]
    private bool _isMonitoring;

    [ObservableProperty]
    private string _comStatusText = "Not monitored";

    // Sprint 8: PR Project Info
    [ObservableProperty]
    private string _prNumber = "";

    [ObservableProperty]
    private string _projectName = "";

    [ObservableProperty]
    private string _jobWorkingPlanName = "";

    [ObservableProperty]
    private int _projectId;

    [ObservableProperty]
    private string _projectLookupStatus = "";

    // Sprint 8: Clear Table IDs
    [ObservableProperty]
    private string _clearIdsStatus = "";

    [ObservableProperty]
    private bool _isClearingIds;

    // Phase 2.2: Layouts
    [ObservableProperty]
    private ObservableCollection<LayoutInfo> _layouts = new();

    [ObservableProperty]
    private LayoutInfo? _selectedLayout;

    [ObservableProperty]
    private string _activeLayoutName = "N/A";

    // Phase 2.3: Parameter Extraction
    [ObservableProperty]
    private ObservableCollection<TableRowData> _tableData = new();

    [ObservableProperty]
    private Dictionary<string, object> _layoutAttributes = new();

    [ObservableProperty]
    private string _extractionStatus = "No extraction performed";

    [ObservableProperty]
    private bool _isExtracting;

    // DWG file-based extraction (ACadSharp, no COM)
    [ObservableProperty]
    private string _dwgFilePath = string.Empty;

    [ObservableProperty]
    private bool _isDwgFileLoaded;

    /// <summary>
    /// True when either COM is connected OR a DWG file is loaded.
    /// Used by the XAML to show/hide the layouts panel.
    /// </summary>
    public bool HasDataSource => IsConnected || IsDwgFileLoaded;

    partial void OnIsConnectedChanged(bool value)
    {
        OnPropertyChanged(nameof(HasDataSource));
        ExtractParametersCommand.NotifyCanExecuteChanged();
        OpenDwgFileCommand.NotifyCanExecuteChanged();
        ClearTableIdsCommand.NotifyCanExecuteChanged();
        ClearAllTableIdsCommand.NotifyCanExecuteChanged();
        ExtractPRInfoCommand.NotifyCanExecuteChanged();

        if (value)
        {
            ComStatusText = "Connected";
            StartMonitoring();
            // Fetch info + layouts regardless of how connection was established
            _ = FetchAutoCADDetailsAsync();
        }
        else
        {
            ComStatusText = "Disconnected";
            DocumentPath = "N/A";
            StopMonitoring();
        }
    }

    /// <summary>
    /// Fetches AutoCAD version/document info and layouts after connection is detected.
    /// Fire-and-forget from OnIsConnectedChanged — exceptions are logged, not thrown.
    /// </summary>
    private async Task FetchAutoCADDetailsAsync()
    {
        try
        {
            await UpdateAutoCADInfoAsync();
            await RefreshLayoutsAsync();
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to fetch AutoCAD details after connection");
        }
    }

    partial void OnIsDwgFileLoadedChanged(bool value)
    {
        OnPropertyChanged(nameof(HasDataSource));
        ExtractParametersCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedLayoutChanged(LayoutInfo? value)
    {
        ExtractParametersCommand.NotifyCanExecuteChanged();
        SelectLayoutCommand.NotifyCanExecuteChanged();
    }

    public AutoCADViewModel(
        IAutoCADService autoCADService,
        IDwgReaderService dwgReader,
        IGUIProxy guiProxy,
        IOdooService odooService,
        ISettingsService settingsService,
        IAppLogService logService,
        ILogger<AutoCADViewModel>? logger = null)
    {
        _autoCADService = autoCADService ?? throw new ArgumentNullException(nameof(autoCADService));
        _dwgReader = dwgReader ?? throw new ArgumentNullException(nameof(dwgReader));
        _guiProxy = guiProxy ?? throw new ArgumentNullException(nameof(guiProxy));
        _odooService = odooService ?? throw new ArgumentNullException(nameof(odooService));
        _settingsService = settingsService ?? throw new ArgumentNullException(nameof(settingsService));
        _logService = logService ?? throw new ArgumentNullException(nameof(logService));
        _logger = logger;

        // Initialize connection status
        UpdateConnectionStatus();
    }

    #region Connection Management

    /// <summary>
    /// Toggle connection command (Connect/Disconnect).
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanToggleConnection))]
    private async Task ToggleConnectionAsync()
    {
        if (IsConnected)
        {
            await DisconnectAsync();
        }
        else
        {
            await ConnectAsync();
        }
    }

    private bool CanToggleConnection() => !IsConnecting;

    /// <summary>
    /// Connects to AutoCAD with retry logic.
    /// All COM operations executed through IGUIProxy for thread safety.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanConnect))]
    private async Task ConnectAsync()
    {
        IsConnecting = true;
        StatusMessage = "Connecting to AutoCAD...";
        ConnectButtonText = "Connecting...";

        try
        {
            _logger?.LogInformation("Attempting to connect to AutoCAD");

            // Execute connection through GUI proxy for STA thread safety
            var response = await _guiProxy.ExecuteInGuiAsync("autocad_connect", null, timeout: 15000);

            if (response.Success && response.Result is bool connected && connected)
            {
                IsConnected = true;
                StatusMessage = "Successfully connected to AutoCAD";
                ConnectButtonText = "Disconnect";

                _logService.Log("AutoCAD connection established", "AutoCAD");
                _logger?.LogInformation("AutoCAD connection established");

                // Info + layouts are fetched via OnIsConnectedChanged → FetchAutoCADDetailsAsync
            }
            else
            {
                IsConnected = false;
                StatusMessage = response.ErrorMessage ?? "Failed to connect to AutoCAD. Ensure AutoCAD is running.";
                ConnectButtonText = "Connect to AutoCAD";

                _logService.Log($"AutoCAD connection failed: {response.ErrorMessage}", "AutoCAD", AppLogLevel.Warning);
                _logger?.LogWarning("AutoCAD connection failed: {Error}", response.ErrorMessage);
            }
        }
        catch (Exception ex)
        {
            IsConnected = false;
            StatusMessage = $"Connection error: {ex.Message}";
            ConnectButtonText = "Connect to AutoCAD";

            _logService.Log($"Connection error: {ex.Message}", "AutoCAD", AppLogLevel.Error);
            _logger?.LogError(ex, "Exception during AutoCAD connection");
        }
        finally
        {
            IsConnecting = false;
            ToggleConnectionCommand.NotifyCanExecuteChanged();
        }
    }

    private bool CanConnect() => !IsConnecting && !IsConnected;

    /// <summary>
    /// Disconnects from AutoCAD.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanDisconnect))]
    private async Task DisconnectAsync()
    {
        IsConnecting = true;
        StatusMessage = "Disconnecting from AutoCAD...";
        ConnectButtonText = "Disconnecting...";

        try
        {
            _logger?.LogInformation("Disconnecting from AutoCAD");

            // Execute disconnection through GUI proxy
            var response = await _guiProxy.ExecuteInGuiAsync("autocad_disconnect", null, timeout: 5000);

            if (response.Success)
            {
                IsConnected = false;
                StatusMessage = "Disconnected from AutoCAD";
                ConnectButtonText = "Connect to AutoCAD";
                AutoCADVersion = "N/A";
                CurrentDocument = "No document open";
                DocumentPath = "N/A";

                _logService.Log("AutoCAD disconnected", "AutoCAD");
                _logger?.LogInformation("AutoCAD disconnected successfully");
            }
            else
            {
                StatusMessage = response.ErrorMessage ?? "Failed to disconnect properly";
                _logger?.LogWarning("AutoCAD disconnection failed: {Error}", response.ErrorMessage);
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Disconnection error: {ex.Message}";
            _logger?.LogError(ex, "Exception during AutoCAD disconnection");
        }
        finally
        {
            IsConnecting = false;
            ToggleConnectionCommand.NotifyCanExecuteChanged();
        }
    }

    private bool CanDisconnect() => !IsConnecting && IsConnected;

    /// <summary>
    /// Updates AutoCAD information (version, document name, document path).
    /// Called after successful connection.
    /// </summary>
    private async Task UpdateAutoCADInfoAsync()
    {
        try
        {
            // Get AutoCAD status through GUI proxy
            var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_status", null, timeout: 5000);

            if (response.Success && response.Result is AutoCADStatus status)
            {
                AutoCADVersion = status.Version ?? "Unknown";
                CurrentDocument = status.CurrentDocument ?? "No document open";
                DocumentPath = status.DocumentPath ?? "N/A";

                _logger?.LogDebug("AutoCAD info updated - Version: {Version}, Document: {Document}",
                    AutoCADVersion, CurrentDocument);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Failed to update AutoCAD information");
            AutoCADVersion = "Error";
            CurrentDocument = "Error retrieving info";
            DocumentPath = "N/A";
        }
    }

    /// <summary>
    /// Updates connection status based on service state.
    /// </summary>
    private void UpdateConnectionStatus()
    {
        try
        {
            IsConnected = _autoCADService.IsConnected;
            ConnectButtonText = IsConnected ? "Disconnect" : "Connect to AutoCAD";
            StatusMessage = IsConnected
                ? "Connected to AutoCAD"
                : "Click Connect to establish connection with AutoCAD";
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error updating connection status");
            IsConnected = false;
            ConnectButtonText = "Connect to AutoCAD";
            StatusMessage = "AutoCAD status check failed. Click Connect to try again.";
        }
    }

    /// <summary>
    /// Refreshes the connection status and info.
    /// </summary>
    [RelayCommand]
    public async Task RefreshConnectionStatusAsync()
    {
        if (IsConnected)
        {
            await UpdateAutoCADInfoAsync();
            StatusMessage = "Status refreshed";
        }
        else
        {
            UpdateConnectionStatus();
        }
    }

    /// <summary>
    /// Refreshes the connection status (public method for external updates).
    /// </summary>
    public async Task RefreshStatusAsync()
    {
        await RefreshConnectionStatusAsync();
    }

    #endregion

    #region COM Health Monitoring

    private void StartMonitoring()
    {
        if (_healthTimer != null) return;

        _healthTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(10)
        };
        _healthTimer.Tick += OnHealthTimerTick;
        _healthTimer.Start();
        IsMonitoring = true;
        ComStatusText = "Connected (monitoring)";
        _logger?.LogDebug("COM health monitoring started");
    }

    private void StopMonitoring()
    {
        if (_healthTimer != null)
        {
            _healthTimer.Stop();
            _healthTimer.Tick -= OnHealthTimerTick;
            _healthTimer = null;
        }
        IsMonitoring = false;
        _logger?.LogDebug("COM health monitoring stopped");
    }

    private void OnHealthTimerTick(object? sender, EventArgs e)
    {
        if (!IsConnected) return;

        try
        {
            var isStillConnected = _autoCADService.IsConnected;
            if (!isStillConnected)
            {
                _logger?.LogWarning("COM health check: AutoCAD connection lost");
                IsConnected = false;
                StatusMessage = "AutoCAD connection lost (detected by health check)";
                ConnectButtonText = "Connect to AutoCAD";
                AutoCADVersion = "N/A";
                CurrentDocument = "No document open";
                DocumentPath = "N/A";
                ComStatusText = "Connection lost";

                _logService.Log("AutoCAD connection lost (health check)", "AutoCAD", AppLogLevel.Warning);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "COM health check error");
        }
    }

    #endregion

    #region PR Project Info (Sprint 8)

    [RelayCommand(CanExecute = nameof(CanExtractPRInfo))]
    private async Task ExtractPRInfoAsync()
    {
        ProjectLookupStatus = "Extracting PR number...";

        try
        {
            var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_pr_number", null, timeout: 10000);

            if (response.Success && response.Result is string prNum && !string.IsNullOrWhiteSpace(prNum))
            {
                PrNumber = prNum;
                _logService.Log($"PR number extracted: {prNum}", "AutoCAD");

                // Lookup project in Odoo via Swagger API (matches Python get_project_v2)
                ProjectLookupStatus = "Looking up project in Odoo...";
                try
                {
                    var configs = await _settingsService.LoadServerConfigsAsync();
                    configs.TryGetValue("odoo_swagger_url", out var swaggerUrl);
                    configs.TryGetValue("odoo_user_token", out var userToken);

                    if (string.IsNullOrEmpty(swaggerUrl))
                    {
                        var appSettings = _settingsService.GetAppSettings();
                        swaggerUrl = appSettings.Odoo.SwaggerUrl;
                    }

                    if (string.IsNullOrWhiteSpace(swaggerUrl) || string.IsNullOrWhiteSpace(userToken))
                    {
                        ProjectLookupStatus = "Odoo connection not configured (check Settings)";
                        _logService.Log("Odoo lookup skipped: missing Swagger URL or user token", "AutoCAD", AppLogLevel.Warning);
                        return;
                    }

                    var parsed = OdooConnectionViewModel.ParseSwaggerUrl(swaggerUrl!);
                    if (parsed == null)
                    {
                        ProjectLookupStatus = "Invalid Swagger URL (check Settings)";
                        _logService.Log("Odoo lookup skipped: invalid Swagger URL", "AutoCAD", AppLogLevel.Warning);
                        return;
                    }

                    var (baseUrl, database, apiToken) = parsed.Value;

                    // Resolve endpoint path from swagger spec
                    var endpointPath = "/api/v1/boq_import_api/callMethodForJobWorkingPlanBoqModel";
                    try
                    {
                        using var httpClient = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromSeconds(10) };
                        var specJson = await httpClient.GetStringAsync(swaggerUrl);
                        endpointPath = OdooConnectionViewModel.ResolveSwaggerEndpoint(specJson);
                    }
                    catch
                    {
                        // Use default endpoint path
                    }

                    _logService.Log($"Calling get_project_v2: base={baseUrl}, endpoint={endpointPath}, PR={prNum}", "AutoCAD");

                    var result = await _odooService.GetProjectViaApiAsync(
                        prNum, baseUrl, endpointPath, database, userToken!);

                    _logService.Log($"get_project_v2 result: {result.DiagnosticMessage}", "AutoCAD",
                        result.Project != null ? AppLogLevel.Info : AppLogLevel.Warning);

                    if (result.Project != null)
                    {
                        ProjectName = result.Project.Name;
                        ProjectId = result.Project.Id;
                        JobWorkingPlanName = result.Project.JobWorkingPlanName ?? "";
                        ProjectLookupStatus = $"Found: {result.Project.Name} (ID: {result.Project.Id})";
                    }
                    else
                    {
                        ProjectName = "";
                        ProjectId = 0;
                        JobWorkingPlanName = "";
                        ProjectLookupStatus = $"No project found for PR '{prNum}'";
                    }
                }
                catch (Exception ex)
                {
                    ProjectLookupStatus = $"Odoo lookup failed: {ex.Message}";
                    _logger?.LogWarning(ex, "Odoo project lookup failed for PR {PrNumber}", prNum);
                }
            }
            else
            {
                PrNumber = "";
                ProjectLookupStatus = "No PR number found in drawing";
                _logService.Log("No PR number found in drawing", "AutoCAD", AppLogLevel.Warning);
            }
        }
        catch (Exception ex)
        {
            ProjectLookupStatus = $"Extraction failed: {ex.Message}";
            _logger?.LogError(ex, "PR info extraction failed");
        }
    }

    private bool CanExtractPRInfo() => IsConnected && !IsConnecting;

    #endregion

    #region Clear Table IDs (Sprint 8)

    [RelayCommand(CanExecute = nameof(CanClearTableIds))]
    private async Task ClearTableIdsAsync()
    {
        if (SelectedLayout == null)
        {
            ClearIdsStatus = "Please select a layout first";
            return;
        }

        IsClearingIds = true;
        ClearIdsStatus = $"Clearing IDs in {SelectedLayout.Name}...";

        try
        {
            var parameters = new Dictionary<string, object?>
            {
                ["layoutName"] = SelectedLayout.Name
            };
            var response = await _guiProxy.ExecuteInGuiAsync("autocad_clear_table_ids", parameters, timeout: 10000);

            if (response.Success && response.Result is int count)
            {
                ClearIdsStatus = $"Cleared {count} cells in {SelectedLayout.Name}";
                _logService.Log($"Cleared {count} ID cells in {SelectedLayout.Name}", "AutoCAD");
            }
            else
            {
                ClearIdsStatus = response.ErrorMessage ?? "Clear operation failed";
            }
        }
        catch (Exception ex)
        {
            ClearIdsStatus = $"Error: {ex.Message}";
            _logger?.LogError(ex, "Clear table IDs failed");
        }
        finally
        {
            IsClearingIds = false;
        }
    }

    [RelayCommand(CanExecute = nameof(CanClearTableIds))]
    private async Task ClearAllTableIdsAsync()
    {
        IsClearingIds = true;
        ClearIdsStatus = "Clearing IDs in all layouts...";

        try
        {
            var response = await _guiProxy.ExecuteInGuiAsync("autocad_clear_all_table_ids", null, timeout: 30000);

            if (response.Success && response.Result is int count)
            {
                ClearIdsStatus = $"Cleared {count} cells across all layouts";
                _logService.Log($"Cleared {count} ID cells across all layouts", "AutoCAD");
            }
            else
            {
                ClearIdsStatus = response.ErrorMessage ?? "Clear all operation failed";
            }
        }
        catch (Exception ex)
        {
            ClearIdsStatus = $"Error: {ex.Message}";
            _logger?.LogError(ex, "Clear all table IDs failed");
        }
        finally
        {
            IsClearingIds = false;
        }
    }

    private bool CanClearTableIds() => IsConnected && !IsConnecting && !IsClearingIds;

    #endregion

    #region DWG File-Based Extraction

    /// <summary>
    /// Opens a DWG file and loads its layouts via ACadSharp (no COM needed).
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanOpenDwgFile))]
    private void OpenDwgFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Open DWG File",
            Filter = "AutoCAD Drawing (*.dwg)|*.dwg|All Files (*.*)|*.*",
            DefaultExt = ".dwg"
        };

        if (dialog.ShowDialog() != true)
            return;

        var filePath = dialog.FileName;

        try
        {
            _logger?.LogInformation("Opening DWG file: {FilePath}", filePath);

            if (!_dwgReader.IsValidDwgFile(filePath))
            {
                StatusMessage = "Selected file is not a valid DWG file.";
                _logService.Log($"Invalid DWG file: {filePath}", "AutoCAD", AppLogLevel.Warning);
                return;
            }

            var layouts = _dwgReader.GetLayouts(filePath);

            Layouts.Clear();
            foreach (var layout in layouts)
            {
                Layouts.Add(layout);
            }

            DwgFilePath = filePath;
            IsDwgFileLoaded = true;
            SelectedLayout = null;
            CurrentDocument = Path.GetFileName(filePath);
            StatusMessage = $"Loaded {layouts.Count} layouts from {Path.GetFileName(filePath)}";

            _logService.Log($"Opened DWG file: {Path.GetFileName(filePath)} ({layouts.Count} layouts)", "AutoCAD");
            _logger?.LogInformation("DWG file loaded: {Count} layouts", layouts.Count);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Failed to open DWG file: {ex.Message}";
            _logService.Log($"DWG open error: {ex.Message}", "AutoCAD", AppLogLevel.Error);
            _logger?.LogError(ex, "Failed to open DWG file: {FilePath}", filePath);
        }
    }

    private bool CanOpenDwgFile() => !IsConnecting && !IsExtracting;

    #endregion

    #region Phase 2.2: Layout Management

    /// <summary>
    /// Refreshes the list of layouts from AutoCAD.
    /// Excludes "Model" layout as per requirements.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanRefreshLayouts))]
    private async Task RefreshLayoutsAsync()
    {
        if (!IsConnected)
        {
            StatusMessage = "Not connected to AutoCAD";
            return;
        }

        try
        {
            _logger?.LogInformation("Refreshing layouts");

            var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_layouts", null, timeout: 5000);

            if (response.Success && response.Result is List<LayoutInfo> layouts)
            {
                // Exclude "Model" layout
                var filteredLayouts = layouts.Where(l => !l.IsModelSpace).ToList();

                Layouts.Clear();
                foreach (var layout in filteredLayouts)
                {
                    Layouts.Add(layout);
                }

                // Get active layout
                var activeResponse = await _guiProxy.ExecuteInGuiAsync("autocad_get_active_layout", null, timeout: 3000);
                if (activeResponse.Success && activeResponse.Result is string activeLayout)
                {
                    ActiveLayoutName = activeLayout;
                }

                _logService.Log($"Loaded {Layouts.Count} layouts", "AutoCAD");
                _logger?.LogInformation("Loaded {Count} layouts", Layouts.Count);
            }
            else
            {
                StatusMessage = response.ErrorMessage ?? "Failed to load layouts";
                _logger?.LogWarning("Failed to refresh layouts: {Error}", response.ErrorMessage);
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error loading layouts: {ex.Message}";
            _logger?.LogError(ex, "Exception while refreshing layouts");
        }
    }

    private bool CanRefreshLayouts() => IsConnected && !IsConnecting;

    /// <summary>
    /// Selects and switches to a layout.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanSelectLayout))]
    private async Task SelectLayoutAsync(LayoutInfo? layout)
    {
        if (layout == null || !IsConnected) return;

        try
        {
            _logger?.LogInformation("Switching to layout: {LayoutName}", layout.Name);

            var parameters = new Dictionary<string, object?>
            {
                ["layoutName"] = layout.Name
            };
            var response = await _guiProxy.ExecuteInGuiAsync("autocad_set_active_layout", parameters, timeout: 5000);

            if (response.Success)
            {
                ActiveLayoutName = layout.Name;
                SelectedLayout = layout;
                StatusMessage = $"Switched to layout: {layout.Name}";
                _logger?.LogInformation("Successfully switched to layout: {LayoutName}", layout.Name);
            }
            else
            {
                StatusMessage = response.ErrorMessage ?? $"Failed to switch to layout: {layout.Name}";
                _logger?.LogWarning("Failed to switch layout: {Error}", response.ErrorMessage);
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error switching layout: {ex.Message}";
            _logger?.LogError(ex, "Exception while switching to layout: {LayoutName}", layout.Name);
        }
    }

    private bool CanSelectLayout() => IsConnected && !IsConnecting;

    #endregion

    #region Phase 2.3: Parameter Extraction

    /// <summary>
    /// Extracts parameters from the selected layout.
    /// Uses file-based ACadSharp reader when a DWG file is loaded,
    /// or COM-based GUIProxy when connected to AutoCAD.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanExtractParameters))]
    private async Task ExtractParametersAsync()
    {
        if (SelectedLayout == null)
        {
            ExtractionStatus = "Please select a layout first";
            return;
        }

        IsExtracting = true;
        ExtractionStatus = "Extracting parameters...";

        try
        {
            _logService.Log($"Extracting parameters from layout: {SelectedLayout.Name}", "AutoCAD");
            _logger?.LogInformation("Extracting parameters from layout: {LayoutName}", SelectedLayout.Name);

            LayoutData? data = null;

            if (IsDwgFileLoaded)
            {
                // File-based extraction via ACadSharp (no COM)
                data = await Task.Run(() => _dwgReader.ExtractParameters(DwgFilePath, SelectedLayout.Name));
            }
            else if (IsConnected)
            {
                // COM-based extraction via GUIProxy
                var parameters = new Dictionary<string, object?>
                {
                    ["layoutName"] = SelectedLayout.Name
                };
                var response = await _guiProxy.ExecuteInGuiAsync("autocad_extract_parameters", parameters, timeout: 30000);

                if (response.Success && response.Result is LayoutData comData)
                {
                    data = comData;
                }
                else
                {
                    ExtractionStatus = response.ErrorMessage ?? "Extraction failed";
                    _logger?.LogWarning("Parameter extraction failed: {Error}", response.ErrorMessage);
                    return;
                }
            }

            if (data != null)
            {
                PopulateExtractionResults(data);
            }
        }
        catch (Exception ex)
        {
            ExtractionStatus = $"Extraction error: {ex.Message}";
            _logger?.LogError(ex, "Exception during parameter extraction");
        }
        finally
        {
            IsExtracting = false;
        }
    }

    private bool CanExtractParameters() =>
        (IsConnected || IsDwgFileLoaded) && !IsConnecting && !IsExtracting && SelectedLayout != null;

    /// <summary>
    /// Populates LayoutAttributes and TableData from extracted LayoutData.
    /// Shared by both COM and file-based extraction paths.
    /// </summary>
    private void PopulateExtractionResults(LayoutData data)
    {
        LayoutAttributes = data.Parameters;

        TableData.Clear();
        if (data.Tables.Count > 0)
        {
            var table = data.Tables[0];

            foreach (var row in table.Cells.Skip(1)) // Skip header row
            {
                if (row.Count >= 9)
                {
                    TableData.Add(new TableRowData
                    {
                        Position = row[0],
                        ProductNo = row[1],
                        Width = row[2],
                        Height = row[3],
                        Length = row[4],
                        Thickness = row[5],
                        Qty = row[6],
                        Description = row[7],
                        DetailId = row[8]
                    });
                }
            }
        }

        ExtractionStatus = $"Extracted {data.Parameters.Count} parameters, {TableData.Count} rows from {SelectedLayout!.Name}";
        _logService.Log($"Extracted {data.Parameters.Count} params, {TableData.Count} rows from {SelectedLayout.Name}", "AutoCAD");
        _logger?.LogInformation("Successfully extracted {ParamCount} params, {RowCount} rows",
            data.Parameters.Count, TableData.Count);
    }

    #endregion
}

/// <summary>
/// Represents a row in the extracted table data.
/// </summary>
public class TableRowData
{
    public string Position { get; set; } = string.Empty;
    public string ProductNo { get; set; } = string.Empty;
    public string Width { get; set; } = string.Empty;
    public string Height { get; set; } = string.Empty;
    public string Length { get; set; } = string.Empty;
    public string Thickness { get; set; } = string.Empty;
    public string Qty { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string DetailId { get; set; } = string.Empty; // Hidden column
}
