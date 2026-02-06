# FR-005: Purchase Requisition Page

> **Document Version**: 1.0
> **Last Updated**: 2026-02-06
> **Status**: Not Started
> **Priority**: P2

## 1. Overview

The Purchase Requisition (PR) Page manages the conversion of BOQ (Bill of Quantities) entries into Purchase Requisitions in Odoo, and provides a complete view for tracking, reviewing, and submitting PRs for approval. This is the final step in the core engineering workflow: parameters are extracted from AutoCAD, pushed to BOQ, and then transferred to Purchase Requisitions for the procurement pipeline.

In the Python application, this operation is a single-action "Transfer BOQ to PR" sidebar button that collects all `header_id` values from AutoCAD layout tables and sends them to Odoo's `boq2pr_v2` endpoint, which performs the actual conversion logic server-side. The C# implementation expands this into a full-featured page with PR listing, detail views, conversion controls, and submission workflow, while maintaining the same principle that the Odoo backend handles most of the conversion logic.

### Workflow Summary

```
AutoCAD Layouts (header_ids) --> Odoo boq2pr_v2 API --> Purchase Requisitions
        |                              |                        |
  get_layouts_header_id_to_pr()   Server-side conversion   PR list returned
```

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-005-01 | CAD Engineer | Convert BOQ entries to a Purchase Requisition | I can initiate the procurement process for the materials in my drawing |
| US-005-02 | CAD Engineer | See a list of all PRs for the current project | I can track which requisitions have been created |
| US-005-03 | CAD Engineer | View the line items within a specific PR | I can verify the correct products and quantities before submission |
| US-005-04 | Project Manager | Submit a PR for approval | The purchasing team can begin procurement |
| US-005-05 | Project Manager | Track the status of each PR (draft, submitted, approved, rejected) | I know where each requisition stands in the approval pipeline |
| US-005-06 | CAD Engineer | Filter and sort PRs by status, date, or reference | I can quickly find specific requisitions in large projects |
| US-005-07 | CAD Engineer | See a confirmation or error when conversion completes | I know whether the operation succeeded and what PRs were created |
| US-005-08 | Project Manager | View PR totals and line counts at a glance | I can assess the scope of each requisition without opening it |

## 3. Python Reference

### Source Files
- `utility/util_transfer_boq_to_pr.py` (~20 lines) - BOQ-to-PR transfer orchestration
- `utility/util_autocad.py` - `get_layouts_header_id_to_pr()` method (~60 lines, lines 604-661)
- `utility/util_odoo.py` - `boq2pr()` method (~20 lines, lines 111-129)
- `forms/form_main_modern.py` - Sidebar button and `transfer_boq_to_pr()` handler (line 358-361, 614-616)
- `forms/form_main.py` - Legacy sidebar button (line 86, 145-147)

### Key Classes and Functions

**UtilTransferBoqToPr** (`util_transfer_boq_to_pr.py`):
```python
class UtilTransferBoqToPr:
    def __init__(self, odoo_util, autocad_util, log_util):
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.log_util = log_util

    def transfer_boq_to_pr(self):
        header_id_dict = self.autocad_util.get_layouts_header_id_to_pr()
        # header_id_dict = {'all': [header_id_1, header_id_2, ...]}
        pr_list = self.odoo_util.boq2pr(header_id_dict)
        # pr_list = [pr_data_1, pr_data_2, ...] or error string
```

**get_layouts_header_id_to_pr()** (`util_autocad.py`):
- Iterates all document layouts (excluding "Model")
- For each layout, reads block attributes and extracts table data
- Collects `header_id` from valid tables (9 columns, "HEADER_ID" in column 7)
- Only includes `header_id` when the layout has non-empty detail rows
- Returns: `{'all': [header_id_1, header_id_2, ...]}`

**boq2pr()** (`util_odoo.py`):
- Calls Odoo API: `job_working_plan_boq.callMethodForJobWorkingPlanBoqModel`
- Method name: `boq2pr_v2`
- Request body: `{"args": [header_id_dict], "kwargs": {"user_token": token}, "context": {}}`
- Success response: `{'all': [pr_data_list]}` -- returns list of PR data
- Error response: `{'error_code': ..., 'error_message': ...}` -- returns error string

### Python Sidebar Button
- Label: "Transfer BOQ to PR"
- Color: Brown (#5D4037 fg, #3E2723 hover)
- Font: Microsoft JhengHei UI, 14pt bold
- Located in the main operations section of the sidebar
- Calls `transfer_boq_to_pr()` which delegates to `UtilTransferBoqToPr.transfer_boq_to_pr()`

## 4. Functional Requirements

### PR Generation from BOQ

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-005-001 | Page SHALL provide a "Convert BOQ to PR" button that triggers the BOQ-to-PR conversion | Must |
| FR-005-002 | Conversion SHALL call `IOdooService.ConvertBOQToPRAsync(projectId, boqEntryIds)` with the current project context | Must |
| FR-005-003 | If no specific BOQ entry IDs are provided, ALL BOQ entries for the project SHALL be included in the conversion | Must |
| FR-005-004 | User SHALL be able to select specific BOQ entries to convert (partial conversion) | Should |
| FR-005-005 | Conversion SHALL display a progress indicator while the Odoo backend processes the request | Must |
| FR-005-006 | On successful conversion, the page SHALL refresh the PR list and show the newly created PR highlighted | Must |
| FR-005-007 | On failure, the page SHALL display the error message returned by the Odoo API | Must |

### PR List Display

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-005-008 | Page SHALL display a list of all Purchase Requisitions for the current project | Must |
| FR-005-009 | Each PR list entry SHALL show: Reference, State, Line Count, and Created Date | Must |
| FR-005-010 | PR list SHALL load automatically when the page is opened and a project context exists | Must |
| FR-005-011 | PR list SHALL support pull-to-refresh or a manual refresh button | Must |
| FR-005-012 | Empty state SHALL display "No Purchase Requisitions found for this project" with guidance | Should |

### PR Detail View

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-005-013 | Selecting a PR from the list SHALL display its detail view with all line items | Must |
| FR-005-014 | PR detail SHALL display: Reference, State, Created Date, and total line count | Must |
| FR-005-015 | PR lines SHALL display in a DataGrid with columns: Product Name, Quantity, Unit of Measure, Unit Price | Must |
| FR-005-016 | PR detail SHALL show a computed total amount (sum of Quantity x UnitPrice for all lines) | Should |

### PR Submission

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-005-017 | Page SHALL provide a "Submit for Approval" button for PRs in "draft" state | Must |
| FR-005-018 | Submit SHALL call `IOdooService.SubmitPRAsync(prId)` and update the PR state on success | Must |
| FR-005-019 | Submit button SHALL be disabled for PRs that are not in "draft" state | Must |
| FR-005-020 | A confirmation dialog SHALL appear before submission: "Submit PR {reference} for approval?" | Must |
| FR-005-021 | On successful submission, the PR state SHALL update to "submitted" in the UI | Must |

### Status Tracking and Filtering

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-005-022 | PR states SHALL be visually distinguished with color-coded badges (draft=gray, submitted=blue, approved=green, rejected=red) | Must |
| FR-005-023 | Page SHALL provide a filter dropdown to filter PRs by state (All, Draft, Submitted, Approved, Rejected) | Should |
| FR-005-024 | Page SHALL provide sorting options: by Reference (A-Z), by Date (newest first), by State | Should |
| FR-005-025 | Page SHALL display a summary bar showing count of PRs by state (e.g., "3 Draft, 2 Submitted, 1 Approved") | Could |

## 5. UI Wireframe Description

```
+--------------------------------------------------------------+
|                   Purchase Requisition                         |
+--------------------------------------------------------------+
|                                                                |
|  Project Context                                               |
|  +------------------------------------------------------------+
|  | Project: Steel Frame Phase 2  |  PR No: PR-2025-001       |
|  | Total PRs: 6                  |  Last Conversion: 14:30   |
|  +------------------------------------------------------------+
|                                                                |
|  Conversion Controls                                           |
|  +------------------------------------------------------------+
|  | [Convert BOQ to PR]  [Refresh List]                        |
|  |  Status: Ready / Converting... / Error: {message}          |
|  +------------------------------------------------------------+
|                                                                |
|  Filter: [All States v]  Sort: [Newest First v]  [Search...] |
|                                                                |
|  PR List                       |  PR Detail                   |
|  +---------------------------+ |  +---------------------------+
|  | REF        STATE    DATE  | |  | PR-2025-006              |
|  | PR-006  [Draft]    02/06  | |  | State: Draft             |
|  | PR-005  [Submitted] 02/05 | |  | Created: 2026-02-06      |
|  | PR-004  [Approved]  02/04 | |  | Lines: 5 items           |
|  | PR-003  [Approved]  02/03 | |  | Total: $12,450.00        |
|  | PR-002  [Submitted] 02/02 | |  |                          |
|  | PR-001  [Approved]  02/01 | |  | Line Items:              |
|  |                           | |  | +------------------------+|
|  |                           | |  | |Product |Qty |UoM |Price||
|  |                           | |  | |H-200x  |4   |pcs |1200||
|  |                           | |  | |H-150x  |8   |pcs |950 ||
|  |                           | |  | |PL-12   |16  |pcs |300 ||
|  |                           | |  | |Bolt M16|64  |pcs |12  ||
|  |                           | |  | |Nut M16 |64  |pcs |5   ||
|  +---------------------------+ |  +------------------------+  |
|                                |                              |
|                                |  [Submit for Approval]       |
|                                +------------------------------+
|                                                                |
|  Summary: 1 Draft | 2 Submitted | 3 Approved | 0 Rejected    |
|                                                                |
+--------------------------------------------------------------+
```

### Layout Details
- **Project Context Panel**: Top banner showing current project info and conversion metadata
- **Conversion Controls**: Action bar with Convert button, refresh, and status text
- **Filter/Sort Bar**: Dropdown filters for state, sort order, and optional text search
- **PR List**: Left panel `ListView` with columns for Reference, State (badge), and Date
- **PR Detail Panel**: Right panel showing selected PR header info and line items DataGrid
- **Submit Button**: Positioned below the detail panel, conditionally visible/enabled
- **Summary Bar**: Bottom bar with state counts for quick overview

### State Badge Colors
| State | Background | Text | Border |
|-------|-----------|------|--------|
| Draft | #E0E0E0 (gray) | #424242 | #9E9E9E |
| Submitted | #BBDEFB (light blue) | #1565C0 | #42A5F5 |
| Approved | #C8E6C9 (light green) | #2E7D32 | #66BB6A |
| Rejected | #FFCDD2 (light red) | #C62828 | #EF5350 |

## 6. Data Model

### ViewModel Properties

```csharp
public class PurchaseRequisitionViewModel : ObservableObject
{
    // Project Context
    public int? ProjectId { get; set; }
    public string ProjectName { get; set; }
    public string PRNumber { get; set; }
    public DateTime? LastConversionTime { get; set; }

    // PR List
    public ObservableCollection<PRDisplayItem> PurchaseRequisitions { get; set; }
    public PRDisplayItem SelectedPR { get; set; }

    // Filtering & Sorting
    public string SelectedStateFilter { get; set; }          // "All", "draft", "submitted", "approved", "rejected"
    public string SelectedSortOrder { get; set; }             // "newest", "oldest", "reference_asc", "reference_desc"
    public string SearchText { get; set; }
    public ObservableCollection<PRDisplayItem> FilteredPRs { get; }  // Computed from filters

    // Status Summary
    public int DraftCount { get; set; }
    public int SubmittedCount { get; set; }
    public int ApprovedCount { get; set; }
    public int RejectedCount { get; set; }
    public int TotalCount { get; set; }

    // Conversion State
    public bool IsConverting { get; set; }
    public bool IsLoading { get; set; }
    public string StatusMessage { get; set; }

    // Detail View
    public ObservableCollection<PRLineDisplayItem> SelectedPRLines { get; set; }
    public decimal SelectedPRTotal { get; set; }

    // Commands
    public IAsyncRelayCommand ConvertBOQToPRCommand { get; }
    public IAsyncRelayCommand RefreshListCommand { get; }
    public IAsyncRelayCommand<PRDisplayItem> SelectPRCommand { get; }
    public IAsyncRelayCommand SubmitPRCommand { get; }
    public IRelayCommand<string> FilterByStateCommand { get; }
    public IRelayCommand<string> SortByCommand { get; }
}

/// <summary>
/// Display model for a PR in the list view.
/// Wraps PREntry with additional computed properties.
/// </summary>
public class PRDisplayItem : ObservableObject
{
    public int Id { get; set; }
    public string Reference { get; set; }
    public string State { get; set; }
    public string StateDisplay { get; set; }         // Localized display string
    public SolidColorBrush StateBadgeBackground { get; set; }
    public SolidColorBrush StateBadgeForeground { get; set; }
    public int LineCount { get; set; }
    public decimal? TotalAmount { get; set; }
    public DateTime CreatedAt { get; set; }
    public string CreatedAtDisplay { get; set; }     // Formatted date string
    public bool CanSubmit { get; set; }              // true only when State == "draft"
}

/// <summary>
/// Display model for a PR line item in the detail DataGrid.
/// </summary>
public class PRLineDisplayItem
{
    public int ProductId { get; set; }
    public string ProductName { get; set; }
    public decimal Quantity { get; set; }
    public string UnitOfMeasure { get; set; }
    public decimal? UnitPrice { get; set; }
    public decimal? LineTotal { get; set; }          // Computed: Quantity * UnitPrice
}
```

### C# Data Model Reference (from IOdooService.cs)

```csharp
// PREntry - defined in OdooAutoCAD.Core/Odoo/IOdooService.cs
public class PREntry
{
    public int? Id { get; set; }
    public int ProjectId { get; set; }
    public string Reference { get; set; } = string.Empty;
    public List<PRLine> Lines { get; set; } = new();
    public string State { get; set; } = "draft";
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}

// PRLine - defined in OdooAutoCAD.Core/Odoo/IOdooService.cs
public class PRLine
{
    public int ProductId { get; set; }
    public string ProductName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public string UnitOfMeasure { get; set; } = string.Empty;
    public decimal? UnitPrice { get; set; }
}
```

## 7. API/Service Dependencies

| Service | Interface | Methods Used | Purpose |
|---------|-----------|-------------|---------|
| Odoo Service | `IOdooService` | `ConvertBOQToPRAsync(projectId, boqEntryIds)` | Triggers BOQ-to-PR conversion on Odoo backend |
| Odoo Service | `IOdooService` | `GetPurchaseRequisitionsAsync(projectId)` | Fetches all PRs for the current project |
| Odoo Service | `IOdooService` | `SubmitPRAsync(prId)` | Submits a draft PR for approval |
| Odoo Service | `IOdooService` | `GetBOQEntriesAsync(projectId)` | Retrieves BOQ entries for selective conversion |
| AutoCAD Service | `IAutoCADService` | `GetLayoutsHeaderIds()` | Extracts header_ids from layouts (for header-based conversion path) |
| GUI Proxy | `IGUIProxy` | `ExecuteInGuiAsync()` | Thread-safe execution of AutoCAD COM calls for header_id extraction |
| Navigation Service | `INavigationService` | `NavigateTo()` | Navigation from dashboard quick-action or sidebar |
| BOQ Processor | `IBOQProcessor` | (indirect) | Upstream dependency -- BOQ must exist before PR conversion |

## 8. Validation Rules

| Rule | Description |
|------|-------------|
| VR-005-001 | Odoo connection MUST be active before any PR operation can proceed |
| VR-005-002 | A valid project context (ProjectId) MUST exist before conversion or PR listing |
| VR-005-003 | BOQ entries MUST exist for the project before conversion; if no BOQ data exists, show a warning and disable the Convert button |
| VR-005-004 | When using header_id-based conversion, at least one valid header_id must be extracted from AutoCAD layouts |
| VR-005-005 | AutoCAD connection MUST be active if the conversion path requires layout header_id extraction |
| VR-005-006 | Submit button SHALL only be enabled when the selected PR has State == "draft" |
| VR-005-007 | Duplicate conversion prevention: if a conversion is already in progress (IsConverting == true), the Convert button SHALL be disabled |
| VR-005-008 | PR Reference must be non-empty in all list and detail displays; show "N/A" if missing |
| VR-005-009 | Quantity values in PR lines must be positive (> 0); lines with zero or negative quantity should be flagged |
| VR-005-010 | Filter state values must be one of: "All", "draft", "submitted", "approved", "rejected" |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| Odoo not connected | "Cannot access Purchase Requisitions. Please connect to Odoo first." | Disable all controls, show Connect button link |
| No project context | "No project selected. Open a drawing in AutoCAD or select a project to view PRs." | Show empty state with guidance |
| No BOQ data for project | "No BOQ entries found for this project. Push parameters to BOQ first before converting to PR." | Disable Convert button, show link to BOQ page |
| AutoCAD not connected (header_id path) | "AutoCAD is not connected. Cannot extract layout header IDs for conversion." | Show warning, suggest connecting AutoCAD |
| No header_ids extracted | "No valid layout tables found in the current drawing. Ensure layouts have valid BOQ tables." | Show warning in status area |
| Conversion API error | "Conversion failed: {error_message from Odoo}" | Display Odoo error_message, enable retry |
| Conversion returns empty result | "Conversion completed but no Purchase Requisitions were created. The BOQ may already be fully converted." | Show informational message |
| Network timeout during conversion | "The conversion request timed out. The Odoo server may be processing a large request. Please try again." | Show retry option |
| PR list fetch fails | "Failed to load Purchase Requisitions: {error details}" | Show error with retry button |
| Submit fails | "Failed to submit PR {reference} for approval: {error details}" | Show error, PR remains in draft state |
| Submit rejected by backend | "PR {reference} could not be submitted. Reason: {server message}" | Display server-provided reason |
| PR has no lines | "PR {reference} has no line items and cannot be submitted." | Disable Submit, show warning |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/PurchaseRequisitionPage.xaml` - WPF page with master-detail layout
- `ViewModels/PurchaseRequisitionViewModel.cs` - MVVM ViewModel with commands and filtering logic

### Simple Pipeline Pattern

The core architectural insight is that the Python implementation follows a simple pipeline pattern where the client application acts as a thin orchestrator:

1. **Extract** header_ids from AutoCAD layouts (client-side COM)
2. **Send** header_ids to Odoo's `boq2pr_v2` endpoint (single API call)
3. **Receive** created PR data from Odoo (server returns results)
4. **Display** results to the user

The Odoo backend handles all of the heavy logic: validating BOQ entries, grouping items, creating PR records, computing prices, and managing state transitions. The C# application should maintain this pattern -- it is not responsible for PR creation logic, only for orchestrating the API calls and presenting results.

### Conversion Flow (C# equivalent of Python pipeline)

```csharp
// PurchaseRequisitionViewModel.ConvertBOQToPR()
private async Task ConvertBOQToPRAsync()
{
    IsConverting = true;
    StatusMessage = "Converting BOQ to Purchase Requisition...";

    try
    {
        // Option A: Project-based conversion (simpler, preferred for C#)
        var result = await _odooService.ConvertBOQToPRAsync(ProjectId.Value);

        // Option B: Header-ID-based conversion (matches Python exactly)
        // var headerIds = await _guiProxy.ExecuteInGuiAsync(
        //     () => _autoCADService.GetLayoutsHeaderIds());
        // var result = await _odooService.ConvertBOQToPRAsync(ProjectId.Value, headerIds);

        if (result != null)
        {
            StatusMessage = $"Successfully created PR: {result.Reference}";
            LastConversionTime = DateTime.Now;
            await RefreshListAsync();
            // Highlight newly created PR
            SelectedPR = PurchaseRequisitions.FirstOrDefault(p => p.Id == result.Id);
        }
        else
        {
            StatusMessage = "Conversion completed but no PR was created.";
        }
    }
    catch (Exception ex)
    {
        StatusMessage = $"Conversion failed: {ex.Message}";
    }
    finally
    {
        IsConverting = false;
    }
}
```

### MVVM Bindings
- Use `CommunityToolkit.Mvvm` for `ObservableObject`, `RelayCommand`, `AsyncRelayCommand`
- PR list bound via `{Binding FilteredPRs}` to `ListView`
- State badges via `{Binding StateBadgeBackground}` with pre-computed brush values
- Detail DataGrid bound via `{Binding SelectedPRLines}`
- Submit button visibility: `{Binding SelectedPR.CanSubmit}`
- Convert button enabled: `{Binding IsConverting, Converter={StaticResource InverseBoolConverter}}`
- Status text: `{Binding StatusMessage}`

### State Mapping
```csharp
// Map Odoo PR states to display values
private static readonly Dictionary<string, (string Display, string Background, string Foreground)> StateMap = new()
{
    ["draft"]     = ("Draft",     "#E0E0E0", "#424242"),
    ["submitted"] = ("Submitted", "#BBDEFB", "#1565C0"),
    ["approved"]  = ("Approved",  "#C8E6C9", "#2E7D32"),
    ["rejected"]  = ("Rejected",  "#FFCDD2", "#C62828"),
};
```

### Odoo API Mapping

| Python Method | Odoo Endpoint | C# Method |
|---------------|---------------|-----------|
| `boq2pr(header_id_dict)` | `job_working_plan_boq.boq2pr_v2` | `IOdooService.ConvertBOQToPRAsync()` |
| `get_layouts_header_id_to_pr()` | N/A (COM client-side) | `IAutoCADService.GetLayoutsHeaderIds()` |
| N/A | `purchase.requisition` search_read | `IOdooService.GetPurchaseRequisitionsAsync()` |
| N/A | `purchase.requisition` action_submit | `IOdooService.SubmitPRAsync()` |

### Key Differences from Python
- Python has a single button that runs the entire pipeline synchronously and logs results to the output console
- C# provides a full page with PR listing, detail view, filtering, and submission -- a richer user experience
- Python extracts `header_id` from AutoCAD layouts as the conversion input; C# also supports project-based conversion via `ConvertBOQToPRAsync(projectId)` for cases where AutoCAD is not connected
- Python error handling returns an error string; C# should catch exceptions and map API error responses to user-friendly messages
- Python has no PR listing or detail view; all feedback is via log output. C# must fetch and display PR data from Odoo
- Python does not support PR submission; C# adds `SubmitPRAsync()` for the approval workflow

### Thread Safety Considerations
- AutoCAD header_id extraction (if used) MUST go through `IGUIProxy.ExecuteInGuiAsync()` to avoid COM thread conflicts
- All Odoo API calls are async and should not block the UI thread
- PR list refresh should debounce rapid consecutive calls
- Conversion operations should prevent double-click by disabling the Convert button while `IsConverting` is true

### Dependencies on Other Features
- **FR-002 (AutoCAD Connection)**: Required if using header_id-based conversion path. The AutoCAD connection provides project context (PR No, Project ID).
- **FR-003 (Odoo Connection)**: Required for all PR operations. Odoo must be connected and authenticated.
- **FR-004 (BOQ Manager)**: BOQ entries must exist in Odoo before they can be converted to PRs. The BOQ page is the upstream workflow step.
