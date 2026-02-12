# TASKS: US-009-02 — Parameter Configuration Form UI

> **Parent US**: [US-009-02](US-009-02-parameter-form-ui.md)
> **Parent FR**: [FR-009](FR-009-parameter-config.md)
> **Priority**: P2
> **Tasks**: 9 | **Effort**: 5S + 2M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] US-009-01 complete (API methods for dropdown data)
- [ ] Sprint 1 infrastructure (Navigation, DI, sidebar pattern)
- [ ] Existing page pattern (BOQPage, PRPage code-behind with DI)

## Acceptance Criteria
- [ ] AC-01: "參數配置" page accessible via sidebar navigation button
- [ ] AC-02: Page displays 3 GroupBox sections: 材料配置, 加工配置, 顏色配置
- [ ] AC-03: Material field is a searchable ComboBox populated from product API
- [ ] AC-04: Unit field is read-only and auto-fills on material selection
- [ ] AC-05: Spec, Category, Operation Flow, Surface Treatment are searchable ComboBoxes
- [ ] AC-06: Color field is searchable; Color No is read-only auto-filled
- [ ] AC-07: All options loaded in parallel on page load
- [ ] AC-08: Loading indicator shown during fetch
- [ ] AC-09: Submit enabled only when all 7 required fields are filled
- [ ] AC-10: Cancel clears all selections
- [ ] AC-11: Status message area for feedback

---

## TASK-009-02-01: Create ParameterConfigViewModel with all properties and collections

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/ParameterConfigViewModel.cs` |
| Estimate | L |
| Depends On | — |
| Blocks | TASK-009-02-06, TASK-009-02-07, TASK-009-02-08 |

### What to do
- Create `ParameterConfigViewModel` class extending `ObservableObject`:
  - Inject: `IOdooService`, `IAutoCADService`, `IGUIProxy`, `ISettingsService`, `IConfiguration`
  - **Material Section** properties:
    - `ObservableCollection<OdooProduct> Products`
    - `[ObservableProperty] private OdooProduct? _selectedProduct;`
    - `[ObservableProperty] private string _unit = string.Empty;` (read-only display)
    - `ObservableCollection<OdooSetupValue> Specs`
    - `[ObservableProperty] private OdooSetupValue? _selectedSpec;`
    - `ObservableCollection<OdooSetupValue> Categories`
    - `[ObservableProperty] private OdooSetupValue? _selectedCategory;`
  - **Processing Section** properties:
    - `ObservableCollection<OdooSetupValue> OperationFlows`
    - `[ObservableProperty] private OdooSetupValue? _selectedOperationFlow;`
    - `ObservableCollection<OdooSetupValue> SurfaceTreatments`
    - `[ObservableProperty] private OdooSetupValue? _selectedSurfaceTreatment;`
  - **Color Section** properties:
    - `ObservableCollection<OdooColor> Colors`
    - `[ObservableProperty] private OdooColor? _selectedColor;`
    - `[ObservableProperty] private string _colorNo = string.Empty;` (read-only display)
  - **State** properties:
    - `[ObservableProperty] private bool _isLoading;`
    - `[ObservableProperty] private bool _isSubmitting;`
    - `[ObservableProperty] private string _statusMessage = string.Empty;`
  - Commands (stub implementations):
    - `[RelayCommand] private async Task LoadOptionsAsync()` — stub
    - `[RelayCommand(CanExecute = nameof(CanSubmit))] private async Task SubmitAsync()` — stub
    - `[RelayCommand] private void Cancel()` — clears all selections
- Initialize all `ObservableCollection<T>` in constructor

### How to verify
- [ ] ViewModel compiles with all properties and commands (AC-02, AC-03, AC-05, AC-06)
- [ ] Collections are initialized (no null reference exceptions)
- [ ] Solution builds without errors

---

## TASK-009-02-02: Create ParameterConfigPage.xaml with 3 GroupBox sections

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/ParameterConfigPage.xaml` |
| Estimate | L |
| Depends On | — |
| Blocks | TASK-009-02-03 |

### What to do
- Create `ParameterConfigPage.xaml` as a `Page` with a `ScrollViewer` wrapping the content:
  - **Header row**: TextBlock "參數配置" with header style
  - **Material GroupBox** (材料配置):
    - Row 1: Label "材料:" + `ComboBox` (IsEditable=True, IsTextSearchEnabled=True, ItemsSource=Products, SelectedItem=SelectedProduct, DisplayMemberPath="Name")
    - Row 2: Label "單位:" + `TextBox` (IsReadOnly=True, Text=Unit)
    - Row 3: Label "材質:" + `ComboBox` (IsEditable=True, ItemsSource=Specs, SelectedItem=SelectedSpec, DisplayMemberPath="Value")
    - Row 4: Label "材料分類:" + `ComboBox` (IsEditable=True, ItemsSource=Categories, SelectedItem=SelectedCategory, DisplayMemberPath="Value")
  - **Processing GroupBox** (加工配置):
    - Row 1: Label "加工流程:" + `ComboBox` (ItemsSource=OperationFlows, SelectedItem=SelectedOperationFlow, DisplayMemberPath="Value")
    - Row 2: Label "表面處理:" + `ComboBox` (ItemsSource=SurfaceTreatments, SelectedItem=SelectedSurfaceTreatment, DisplayMemberPath="Value")
  - **Color GroupBox** (顏色配置):
    - Row 1: Label "顏色:" + `ComboBox` (ItemsSource=Colors, SelectedItem=SelectedColor, DisplayMemberPath="Name")
    - Row 2: Label "色號:" + `TextBox` (IsReadOnly=True, Text=ColorNo)
  - **Action bar**:
    - Button "確定提交" (Command=SubmitCommand)
    - Button "取消" (Command=CancelCommand, Style=SecondaryButton)
  - **Status area**: TextBlock bound to StatusMessage
  - **Loading overlay**: ProgressBar or spinner visible when IsLoading=True
- Use `Grid` layouts within each GroupBox for aligned labels + fields
- Apply CJK font chain and SettingsLabel style for labels
- Set ComboBox MinWidth="200" for consistent sizing

### How to verify
- [ ] Page renders 3 GroupBox sections (AC-02)
- [ ] All ComboBoxes are searchable (IsEditable, IsTextSearchEnabled) (AC-03, AC-05, AC-06)
- [ ] Read-only TextBoxes for Unit and ColorNo (AC-04, AC-06)
- [ ] Submit and Cancel buttons are present (AC-09, AC-10)
- [ ] Loading indicator area exists (AC-08)
- [ ] Status message area exists (AC-11)

---

## TASK-009-02-03: Code-behind with DI resolution pattern

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/ParameterConfigPage.xaml.cs` |
| Estimate | S |
| Depends On | TASK-009-02-01, TASK-009-02-02 |
| Blocks | — |

### What to do
- Create `ParameterConfigPage.xaml.cs` following the established pattern (BOQPage, PRPage):
  ```csharp
  public partial class ParameterConfigPage : Page
  {
      public ParameterConfigPage()
      {
          InitializeComponent();
          try
          {
              var viewModel = App.Services.GetRequiredService<ParameterConfigViewModel>();
              DataContext = viewModel;
              viewModel.LoadOptionsCommand.Execute(null);
          }
          catch (Exception ex)
          {
              // Fallback: display error message in page content
              Content = new TextBlock
              {
                  Text = $"Failed to initialize Parameter Config page: {ex.Message}",
                  Margin = new Thickness(20),
                  TextWrapping = TextWrapping.Wrap
              };
          }
      }
  }
  ```
- The try-catch fallback pattern matches BOQPage.xaml.cs and PRPage.xaml.cs

### How to verify
- [ ] Page resolves ViewModel from DI and sets DataContext
- [ ] LoadOptionsCommand is triggered on page load (AC-07)
- [ ] Error fallback shows a message instead of crashing

---

## TASK-009-02-04: Add sidebar navigation button for 參數配置

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | — |
| Blocks | — |

### What to do
- Add a new NavButton in the sidebar navigation section of `MainWindow.xaml`:
  ```xml
  <Button Content="參數配置"
          Tag="BtnParameterConfig"
          Command="{Binding NavigateCommand}"
          CommandParameter="ParameterConfig"
          Style="{StaticResource NavButton}"/>
  ```
- Place it after the existing "Odoo 連線" button and before "BOQ 管理" button (logical workflow order: connect → configure parameters → extract BOQ)
- Ensure the Tag follows the `Btn{PageName}` pattern for active indicator

### How to verify
- [ ] "參數配置" button appears in sidebar (AC-01)
- [ ] Clicking the button navigates to ParameterConfigPage
- [ ] Active indicator highlights correctly

---

## TASK-009-02-05: DI registration + page title mapping for ParameterConfig

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs`, `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | TASK-009-02-01 |
| Blocks | — |

### What to do
- In `App.xaml.cs` DI registration section, add:
  ```csharp
  services.AddTransient<ParameterConfigViewModel>();
  ```
- In `MainViewModel.cs` page title mapping dictionary, add:
  ```csharp
  { "ParameterConfig", "參數配置" }
  ```
- The NavigationService already handles assembly reflection for `Views.Pages.ParameterConfigPage`, so no changes needed there.

### How to verify
- [ ] ParameterConfigViewModel resolves from DI without error (AC-01)
- [ ] Page title displays "參數配置" in header bar when navigated to
- [ ] Navigation works end-to-end via sidebar button

---

## TASK-009-02-06: Implement parallel API loading in LoadOptionsCommand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/ParameterConfigViewModel.cs` |
| Estimate | M |
| Depends On | TASK-009-02-01 |
| Blocks | — |

### What to do
- Implement `LoadOptionsAsync()`:
  ```csharp
  [RelayCommand]
  private async Task LoadOptionsAsync()
  {
      IsLoading = true;
      StatusMessage = "正在載入選項...";
      try
      {
          // Get settings for Swagger URL, database, token
          var settings = await _settingsService.LoadSettingsAsync();
          var (baseUrl, endpointPath) = ParseSwaggerUrl(settings.SwaggerUrl);

          // Fetch all dropdown data in parallel
          var productsTask = _odooService.GetProductsViaApiAsync(baseUrl, endpointPath, settings.Database, settings.UserToken);
          var specsTask = _odooService.GetSetupViaApiAsync(baseUrl, endpointPath, settings.Database, settings.UserToken, "spec");
          var categoriesTask = _odooService.GetSetupViaApiAsync(baseUrl, endpointPath, settings.Database, settings.UserToken, "product_catelog");
          var flowsTask = _odooService.GetSetupViaApiAsync(baseUrl, endpointPath, settings.Database, settings.UserToken, "operation_flow");
          var surfaceTask = _odooService.GetSetupViaApiAsync(baseUrl, endpointPath, settings.Database, settings.UserToken, "surface_treatment");

          await Task.WhenAll(productsTask, specsTask, categoriesTask, flowsTask, surfaceTask);

          // Populate collections
          Products.Clear();
          foreach (var p in await productsTask) Products.Add(p);
          // ... same for Specs, Categories, OperationFlows, SurfaceTreatments

          // Colors loaded separately (needs project_id, may not be available yet)
          StatusMessage = $"已載入選項 (材料: {Products.Count}, 材質: {Specs.Count}, ...)";
      }
      catch (Exception ex)
      {
          StatusMessage = $"載入選項失敗: {ex.Message}";
      }
      finally
      {
          IsLoading = false;
      }
  }
  ```
- Colors are loaded separately via a `LoadColorsAsync(int projectId)` method called when project_id becomes available
- Reuse `ParseSwaggerUrl` and `ResolveSwaggerEndpoint` logic from OdooConnectionViewModel (extract to shared utility or duplicate inline)

### How to verify
- [ ] All 5 dropdown collections are populated in parallel (AC-07)
- [ ] Loading state is correctly set during fetch (AC-08)
- [ ] Status message shows results or error (AC-11)
- [ ] API failures are handled gracefully

---

## TASK-009-02-07: Implement selection changed auto-fill handlers (Unit, ColorNo)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/ParameterConfigViewModel.cs` |
| Estimate | S |
| Depends On | TASK-009-02-01 |
| Blocks | — |

### What to do
- Implement `partial void OnSelectedProductChanged(OdooProduct? value)`:
  ```csharp
  partial void OnSelectedProductChanged(OdooProduct? value)
  {
      Unit = value?.Unit ?? string.Empty;
      SubmitCommand.NotifyCanExecuteChanged();
  }
  ```
- Implement `partial void OnSelectedColorChanged(OdooColor? value)`:
  ```csharp
  partial void OnSelectedColorChanged(OdooColor? value)
  {
      ColorNo = value?.ColorNo ?? string.Empty;
      SubmitCommand.NotifyCanExecuteChanged();
  }
  ```
- Add `partial void OnSelected{Spec|Category|OperationFlow|SurfaceTreatment}Changed` to call `SubmitCommand.NotifyCanExecuteChanged()` on each selection change

### How to verify
- [ ] Unit auto-fills when material is selected (AC-04)
- [ ] ColorNo auto-fills when color is selected (AC-06)
- [ ] Submit button CanExecute updates on any selection change (AC-09)

---

## TASK-009-02-08: Implement AllFieldsFilled computed property + CanExecute for Submit

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/ParameterConfigViewModel.cs` |
| Estimate | S |
| Depends On | TASK-009-02-01 |
| Blocks | — |

### What to do
- Add computed property:
  ```csharp
  public bool AllFieldsFilled =>
      SelectedProduct != null &&
      SelectedSpec != null &&
      SelectedCategory != null &&
      SelectedOperationFlow != null &&
      SelectedSurfaceTreatment != null &&
      SelectedColor != null &&
      !string.IsNullOrEmpty(ColorNo);
  ```
- Implement `CanSubmit` method for the `[RelayCommand]` attribute:
  ```csharp
  private bool CanSubmit() => AllFieldsFilled && !IsSubmitting;
  ```
- Implement Cancel command:
  ```csharp
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
      StatusMessage = string.Empty;
  }
  ```

### How to verify
- [ ] Submit button is disabled when any field is empty (AC-09)
- [ ] Submit button is enabled when all 7 fields are filled (AC-09)
- [ ] Cancel resets all selections and read-only fields (AC-10)

---

## TASK-009-02-09: ViewModel unit tests for form state, loading, selection, validation

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/ViewModels/ParameterConfigViewModelTests.cs` |
| Estimate | M |
| Depends On | TASK-009-02-01, TASK-009-02-06, TASK-009-02-07, TASK-009-02-08 |
| Blocks | — |

### What to do
- Create `ParameterConfigViewModelTests` class with mocked dependencies:
  1. `LoadOptions_Success_PopulatesAllCollections` — verify all 5 collections populated
  2. `LoadOptions_ApiFailure_SetsErrorStatusMessage` — verify graceful error handling
  3. `LoadOptions_SetsIsLoading_DuringFetch` — verify IsLoading transitions
  4. `SelectedProduct_Changed_AutoFillsUnit` — verify Unit auto-fill
  5. `SelectedColor_Changed_AutoFillsColorNo` — verify ColorNo auto-fill
  6. `AllFieldsFilled_AllSelected_ReturnsTrue` — verify computed property
  7. `AllFieldsFilled_MissingField_ReturnsFalse` — verify for each missing field
  8. `Cancel_ClearsAllSelections` — verify all fields reset
  9. `CanSubmit_AllFieldsFilled_ReturnsTrue` — verify CanExecute
  10. `CanSubmit_IsSubmitting_ReturnsFalse` — verify disabled during submit

### How to verify
- [ ] All 10 tests pass
- [ ] Tests cover loading, selection, validation, and cancel (AC-03 through AC-11)

---

## Dependency Graph
```
TASK-009-02-01 (ViewModel)
    |
    +---> TASK-009-02-06 (LoadOptions)
    |
    +---> TASK-009-02-07 (Selection handlers)
    |
    +---> TASK-009-02-08 (AllFieldsFilled + Cancel)
    |
    +---> TASK-009-02-05 (DI + title mapping)
    |
    +---> TASK-009-02-09 (Tests — depends on 06, 07, 08)

TASK-009-02-02 (XAML)
    |
    +---> TASK-009-02-03 (Code-behind)

TASK-009-02-04 (Sidebar button — independent)
```
