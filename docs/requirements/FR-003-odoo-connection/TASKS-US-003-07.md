# TASKS: US-003-07 — Search Projects

> **Parent US**: [US-003-07](US-003-07-search-projects.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 7 | **Effort**: 3S + 3M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (credential input form) must be completed
- [ ] US-003-03 (connection status and IsConnected property) must be completed

## Acceptance Criteria
- [ ] AC-01: A project search field is visible on the Odoo Integration Page
- [ ] AC-02: The search queries Odoo by project name or code via `SearchProjectsAsync(searchTerm)`
- [ ] AC-03: Search results display in a list/DataGrid with columns: ID, Name, Code, State, Date Start, Date End
- [ ] AC-04: The user can select a project from search results to set it as the active project context
- [ ] AC-05: The selected project is available for downstream operations (BOQ import, PR generation)
- [ ] AC-06: The search field is disabled when `IsConnected` is false
- [ ] AC-07: When no projects match, the message "No projects found matching '{searchTerm}'. Try a different search term." is displayed
- [ ] AC-08: Single project retrieval by ID is supported via `GetProjectAsync(projectId)` for navigation from other pages
- [ ] AC-09: The search operation runs asynchronously without blocking the UI thread

---

## TASK-003-07-01: Create Project Search panel in XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-003-07-06, TASK-003-07-07 |

### What to do
- Add a `GroupBox` or `Border` with `Header="Project Search"` in the middle section of `OdooConnectionPage.xaml` (left panel as per wireframe)
- Layout using `StackPanel` or `Grid`:
  - Row 0: `TextBox` bound to `{Binding ProjectSearchTerm, UpdateSourceTrigger=PropertyChanged}` with placeholder text "Search by PR No / Name..."
  - Row 1: "Search" `Button` bound to `{Binding SearchProjectsCommand}`
  - Row 2: `TextBlock` "Results:" label
  - Row 3: `DataGrid` with `AutoGenerateColumns="False"` bound to `{Binding ProjectSearchResults}`:
    - `DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="50"`
    - `DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"`
    - `DataGridTextColumn Header="Code" Binding="{Binding Code}" Width="80"`
    - `DataGridTextColumn Header="State" Binding="{Binding State}" Width="80"`
    - `DataGridTextColumn Header="Start" Binding="{Binding DateStart, StringFormat=d}" Width="90"`
    - `DataGridTextColumn Header="End" Binding="{Binding DateEnd, StringFormat=d}" Width="90"`
  - Row 4: "Select Project" `Button` bound to `{Binding SelectProjectCommand}`
- Add `DataGrid.SelectedItem` bound to `{Binding SelectedProject, Mode=TwoWay}`
- Add empty state `TextBlock` for no results message bound to `{Binding ProjectSearchEmptyMessage}`
- Wire `TextBox.KeyDown` for Enter key to trigger search

### How to verify
- [ ] Project search field is visible on the page (AC-01)
- [ ] Results DataGrid shows columns: ID, Name, Code, State, Date Start, Date End (AC-03)

---

## TASK-003-07-02: Add project search properties to ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-07-03, TASK-003-07-04 |

### What to do
- Add observable properties using `[ObservableProperty]`:
  - `private string _projectSearchTerm = string.Empty;`
  - `private OdooProject? _selectedProject;`
  - `private string _projectSearchEmptyMessage = string.Empty;`
- Add collection property (initialized in constructor):
  - `public ObservableCollection<OdooProject> ProjectSearchResults { get; } = new();`
- Implement `partial void OnSelectedProjectChanged(OdooProject? value)`:
  - Call `SelectProjectCommand.NotifyCanExecuteChanged()`
  - If value is not null, update any shared state or service that holds the active project context
- Implement `partial void OnProjectSearchTermChanged(string value)`:
  - Call `SearchProjectsCommand.NotifyCanExecuteChanged()`

### How to verify
- [ ] ProjectSearchTerm, ProjectSearchResults, SelectedProject properties exist (AC-01, AC-04)

---

## TASK-003-07-03: Implement SearchProjectsCommand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-07-02 |
| Blocks | None |

### What to do
- Add `SearchProjectsCommand` as `IAsyncRelayCommand` using `[RelayCommand(CanExecute = nameof(CanSearchProjects))]` attribute on `async Task SearchProjectsAsync()` method
- In `SearchProjectsAsync()`:
  - Call `var projects = await _odooService.SearchProjectsAsync(ProjectSearchTerm)`
  - Clear and repopulate `ProjectSearchResults` with results
  - If `projects.Count == 0`, set `ProjectSearchEmptyMessage = $"No projects found matching '{ProjectSearchTerm}'. Try a different search term."`
  - Else clear `ProjectSearchEmptyMessage`
- Add `bool CanSearchProjects() => IsConnected && !string.IsNullOrWhiteSpace(ProjectSearchTerm)`
- Validation rule VR-003-009: project search only available when `IsConnected` is true
- Handle `HttpRequestException` and other errors with user-friendly messages
- Log the search operation via `_logger`

### How to verify
- [ ] Search queries Odoo by project name or code via SearchProjectsAsync (AC-02)
- [ ] No results message displayed correctly (AC-07)
- [ ] Search runs asynchronously (AC-09)

---

## TASK-003-07-04: Implement SelectProjectCommand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-07-02 |
| Blocks | None |

### What to do
- Add `SelectProjectCommand` as `IRelayCommand` using `[RelayCommand(CanExecute = nameof(CanSelectProject))]` attribute on `void SelectProject()` method
- In `SelectProject()`:
  - Confirm `SelectedProject` is not null
  - Store the selected project in a shared service or state that other ViewModels can access (e.g., `IProjectContextService` with `ActiveProject` property)
  - Update `StatusMessage = $"Active project: {SelectedProject.Name} (ID: {SelectedProject.Id})"`
  - Optionally navigate to the BOQ page or enable BOQ/PR tabs (via `INavigationService.NavigateTo()`)
  - Log the project selection
- Add `bool CanSelectProject() => SelectedProject != null`
- Notify dependent commands (BOQ sync requires a selected project per VR-003-011)

### How to verify
- [ ] User can select a project from search results (AC-04)
- [ ] Selected project is available for downstream operations (AC-05)

---

## TASK-003-07-05: Implement GetProjectAsync for single-project retrieval

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Verify the existing `GetProjectAsync(int projectId)` implementation in `OdooService`:
  ```csharp
  public async Task<OdooProject?> GetProjectAsync(int projectId)
  {
      var result = await ReadAsync("project.project", new[] { projectId },
          new[] { "id", "name", "code", "state", "date_start", "date" });
      return result.FirstOrDefault() is { } data ? MapToOdooProject(data) : null;
  }
  ```
- The existing implementation correctly uses JSON-RPC `read` method with specific project ID
- Add a public method to the ViewModel that other pages can call for navigation:
  ```csharp
  public async Task LoadProjectByIdAsync(int projectId)
  {
      var project = await _odooService.GetProjectAsync(projectId);
      if (project != null) { SelectedProject = project; }
  }
  ```
- Handle `null` return (project not found) gracefully with logging
- This supports navigation from BOQ or PR pages that need to set the active project context

### How to verify
- [ ] Single project retrieval by ID works via GetProjectAsync (AC-08)
- [ ] Returns null for non-existent project ID (AC-08)

---

## TASK-003-07-06: Bind search field IsEnabled to IsConnected

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | TASK-003-07-01 |
| Blocks | None |

### What to do
- Bind the project search `TextBox.IsEnabled` to `{Binding IsConnected}`
- The "Search" button's enabled state is automatically controlled by `SearchProjectsCommand.CanExecute` which checks `IsConnected`
- The "Select Project" button's enabled state is controlled by `SelectProjectCommand.CanExecute` which checks `SelectedProject != null`
- When `IsConnected` is false, apply disabled visual styling to the entire Project Search panel

### How to verify
- [ ] Search field is disabled when IsConnected is false (AC-06)

---

## TASK-003-07-07: Add "Select Project" button below results DataGrid

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | TASK-003-07-01 |
| Blocks | None |

### What to do
- Add a `Button` labeled "Select Project" below the project search results `DataGrid`
- Bind `Command` to `{Binding SelectProjectCommand}`
- The button is automatically disabled when `SelectedProject` is null (via `CanSelectProject()`)
- Optionally add a `TextBlock` below the button showing the currently selected project: `"Active: {SelectedProject.Name}"` with visibility bound to `SelectedProject != null`
- Style the button with emphasis (primary button style) to indicate it confirms the selection

### How to verify
- [ ] "Select Project" button is visible below results (AC-04)
- [ ] Button is disabled when no project is selected (AC-04)

---

## Dependency Graph
```
TASK-003-07-02 (ViewModel properties)
    |
    +---> TASK-003-07-03 (SearchProjectsCommand)
    |
    +---> TASK-003-07-04 (SelectProjectCommand)

TASK-003-07-01 (XAML Project Search panel)
    |
    +---> TASK-003-07-06 (IsEnabled binding)
    +---> TASK-003-07-07 (Select Project button)

TASK-003-07-05 (GetProjectAsync single retrieval)
```
