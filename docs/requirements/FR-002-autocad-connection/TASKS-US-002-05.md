# TASKS: US-002-05 — PR Project Info

> **Parent US**: [US-002-05](US-002-05-pr-project-info.md)
> **Parent FR**: [FR-002](FR-002-autocad-connection.md)
> **Priority**: P1
> **Tasks**: 5 | **Effort**: 3S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-002-03 (parameter extraction provides the PR number and project attributes)

## Acceptance Criteria
- [ ] AC-01: The page displays the extracted PR Number from the active layout in the Connection Status panel
- [ ] AC-02: The page displays the Project Name extracted from layout block attributes
- [ ] AC-03: The page displays the Job Working Plan Name extracted from layout block attributes
- [ ] AC-04: PR number and project info are shown in the format "PR No: PR-2025-001 | Project: Steel Frame Phase 2" as illustrated in the wireframe
- [ ] AC-05: If the PR number is not found in Odoo projects, a warning is displayed: "PR number '{pr_no}' not found in Odoo projects." but extraction continues
- [ ] AC-06: The ProjectId property is populated when a matching Odoo project is found via IOdooService.GetProjectAsync()
- [ ] AC-07: PR number and project info update when parameters are extracted from a different layout

---

## TASK-002-05-01: Add PR Number, Project Name, and Job Working Plan Name display fields

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Connection Status panel of `AutoCADPage.xaml`, add a row below the document info:
  - A `TextBlock` displaying PR and project in the wireframe format:
    ```xml
    <TextBlock>
        <Run Text="PR No: " FontWeight="SemiBold"/>
        <Run Text="{Binding PRNumber}"/>
        <Run Text=" | Project: "/>
        <Run Text="{Binding ProjectName}"/>
    </TextBlock>
    ```
  - A separate `TextBlock` for Job Working Plan Name:
    ```xml
    <TextBlock>
        <Run Text="Job Plan: " FontWeight="SemiBold"/>
        <Run Text="{Binding JobWorkingPlanName}"/>
    </TextBlock>
    ```
- Add a warning `TextBlock` for Odoo lookup failure:
  - Bound to `OdooWarningMessage` property
  - Styled with orange/yellow foreground color for warning visibility
  - Visibility bound to `OdooWarningMessage` being non-empty (use `StringToVisibilityConverter`)
- These fields should be visible only when `IsConnected == true` and values are populated

### How to verify
- [ ] PR Number and Project Name displayed in wireframe format (AC-01, AC-02, AC-04)
- [ ] Job Working Plan Name is displayed (AC-03)
- [ ] Warning message area is present for Odoo lookup failures (AC-05)

---

## TASK-002-05-02: Implement PRNumber, ProjectName, JobWorkingPlanName, and ProjectId properties

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-05-04 |

### What to do
- Add observable properties to `AutoCADViewModel` using `[ObservableProperty]`:
  - `private string _prNumber = string.Empty;` -- extracted PR number (e.g., "PR-2025-001")
  - `private string _projectName = string.Empty;` -- extracted project name
  - `private string _jobWorkingPlanName = string.Empty;` -- extracted job plan name
  - `private int? _projectId;` -- Odoo project ID from lookup
  - `private string _odooWarningMessage = string.Empty;` -- warning for Odoo lookup failure
- These properties are populated during the `ExtractParametersAsync()` flow (TASK-002-03-02):
  - After extracting layout attributes, pull `pr_no` -> `PRNumber`, `project_name` -> `ProjectName`, `job_working_plan_name` -> `JobWorkingPlanName` from the `LayoutAttributes` dictionary
- Implement `partial void OnPRNumberChanged(string value)` to trigger Odoo project lookup (TASK-002-05-04)

### How to verify
- [ ] PRNumber, ProjectName, JobWorkingPlanName properties are bindable (AC-01, AC-02, AC-03)
- [ ] Properties update when new parameters are extracted (AC-07)

---

## TASK-002-05-03: Implement process_pr_no equivalent in AutoCADService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Create a method that extracts PR number and project info from a layout:
  ```csharp
  public Dictionary<string, string> ProcessPRNumber(string layoutName)
  ```
- Implementation:
  - Switch to the specified layout using `SwitchToLayout(layoutName)`
  - Iterate block references in the layout's PaperSpace
  - For each block with attributes, search for the tags: `pr_no`, `project_name`, `job_working_plan_name`
  - Apply `LMUnFormat()` to clean the extracted values
  - Return a dictionary with keys: `"pr_no"`, `"project_name"`, `"job_working_plan_name"`
- Register IGUIProxy handler:
  - `"process_pr_no"` action: takes `layoutName` parameter, calls `ProcessPRNumber(layoutName)`, returns dictionary
- This can also be called as part of `GetLayoutValues()` to include PR info in the `LayoutData.Parameters`
- Handle cases where the block attributes are not found -- return empty strings

### How to verify
- [ ] PR number is extracted from layout block attributes (AC-01)
- [ ] Project name and job plan name are extracted (AC-02, AC-03)

---

## TASK-002-05-04: Integrate IOdooService.GetProjectAsync() lookup for PR-to-ProjectId resolution

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | TASK-002-05-02 |
| Blocks | TASK-002-05-05 |

### What to do
- Inject `IOdooService` into `AutoCADViewModel` constructor (alongside existing `IAutoCADService` and `IGUIProxy`)
- In `partial void OnPRNumberChanged(string value)` or in a dedicated method called after extraction:
  ```csharp
  private async Task LookupOdooProjectAsync(string prNumber)
  ```
- Implementation:
  - If `prNumber` is null or empty, clear `ProjectId` and return
  - Call `await _odooService.GetProjectAsync(prNumber)` to look up the Odoo project
  - If found: set `ProjectId` to the returned project ID, clear `OdooWarningMessage`
  - If not found: set `ProjectId = null`, set `OdooWarningMessage = $"PR number '{prNumber}' not found in Odoo projects."`
  - If Odoo is not connected or the call fails: set `OdooWarningMessage` with the failure reason, but do NOT block the extraction workflow
- This lookup is non-blocking: extraction continues regardless of Odoo lookup result
- Log the lookup result at Information level

### How to verify
- [ ] ProjectId is set when a matching Odoo project is found (AC-06)
- [ ] Extraction continues even when Odoo lookup fails (AC-05)
- [ ] PR info updates when switching layouts (AC-07)

---

## TASK-002-05-05: Add warning display when PR number is not found in Odoo projects

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | TASK-002-05-04 |
| Blocks | None |

### What to do
- Ensure the `OdooWarningMessage` property is properly managed:
  - Set to `$"PR number '{prNumber}' not found in Odoo projects."` when lookup fails (AC-05)
  - Set to empty string when lookup succeeds or when a new PR number is extracted
  - Set to a connection-related message if Odoo service is unavailable
- Add a `HasOdooWarning` computed property:
  ```csharp
  public bool HasOdooWarning => !string.IsNullOrEmpty(OdooWarningMessage);
  ```
- Implement `partial void OnOdooWarningMessageChanged(string value)` to:
  - Call `OnPropertyChanged(nameof(HasOdooWarning))`
- The warning should be displayed but should NOT prevent the user from continuing to work:
  - All other controls remain enabled
  - The user can still extract parameters, view table data, etc.
  - The warning is informational only

### How to verify
- [ ] Warning "PR number '{pr_no}' not found in Odoo projects." is displayed when lookup fails (AC-05)
- [ ] Warning does not block extraction or other operations (AC-05)

---

## Dependency Graph
```
TASK-002-05-01 (XAML PR/Project display fields)

TASK-002-05-03 (ProcessPRNumber service method)

TASK-002-05-02 (ViewModel PR properties)
    |
    +---> TASK-002-05-04 (Odoo project lookup)
              |
              +---> TASK-002-05-05 (Warning display)
```
