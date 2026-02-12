# TASKS: US-009-03 — Write Parameters to AutoCAD

> **Parent US**: [US-009-03](US-009-03-write-parameters-to-autocad.md)
> **Parent FR**: [FR-009](FR-009-parameter-config.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 1S + 5M
> **Status**: Not Started

## Prerequisites
- [ ] US-009-02 complete (form UI with field selections)
- [ ] Sprint 3 infrastructure (AutoCADService, IGUIProxy, STA threading)
- [ ] GUIProxy handler registration pattern established (15+ handlers exist)

## Acceptance Criteria
- [ ] AC-01: Submit writes all 7 attribute values to AutoCAD
- [ ] AC-02: `autocad_get_attribute_block` handler finds the attribute block in a layout
- [ ] AC-03: `autocad_set_attribute_values` handler writes tag→value pairs to a block
- [ ] AC-04: Attributes written to all layouts in current drawing
- [ ] AC-05: All COM operations on GUI/STA thread via IGUIProxy
- [ ] AC-06: Success feedback displayed
- [ ] AC-07: Failure feedback displayed with details
- [ ] AC-08: Submit requires AutoCAD connected
- [ ] AC-09: project_id resolved from current drawing layout attributes

---

## TASK-009-03-01: Implement `autocad_get_attribute_block` GUIProxy handler

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | — |
| Blocks | TASK-009-03-02, TASK-009-03-03 |

### What to do
- Register a new GUIProxy handler `"autocad_get_attribute_block"` in `AutoCADService.RegisterHandlers()`:
  ```csharp
  _guiProxy.RegisterHandler("autocad_get_attribute_block", (parameters) =>
  {
      // Parameters: { "layout_name": "S405-201" }
      var layoutName = parameters?["layout_name"]?.ToString();

      // 1. Switch to the specified layout (or use active if null)
      if (!string.IsNullOrEmpty(layoutName))
      {
          _acadDoc.ActiveLayout = FindLayout(layoutName);
      }

      // 2. Iterate block references in the layout's Block
      dynamic layout = _acadDoc.ActiveLayout;
      dynamic block = layout.Block;

      for (int i = 0; i < (int)block.Count; i++)
      {
          dynamic entity = block.Item(i);
          // Check if entity is a BlockReference with attributes
          if ((int)entity.ObjectName == "AcDbBlockReference" && (bool)entity.HasAttributes)
          {
              dynamic attrs = entity.GetAttributes();
              // Check if this block has the expected parameter tags
              foreach (dynamic attr in attrs)
              {
                  if (attr.TagString == "product_name" || attr.TagString == "pr_no")
                  {
                      return Task.FromResult<object?>(true); // Block found
                  }
              }
          }
      }

      return Task.FromResult<object?>(false); // No attribute block found
  });
  ```
- The handler mirrors Python's `get_attribute_block(layout)` which finds the block reference containing parameter attribute tags
- Returns `true` if an attribute block exists in the layout, `false` otherwise
- Uses `Task.FromResult` (sync) — Pure STA Mode pattern

### How to verify
- [ ] Handler registered with name `"autocad_get_attribute_block"` (AC-02)
- [ ] Finds block references with attribute tags in a layout (AC-02)
- [ ] Runs on STA thread via Task.FromResult (AC-05)

---

## TASK-009-03-02: Implement `autocad_set_attribute_values` GUIProxy handler

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | TASK-009-03-01 |
| Blocks | TASK-009-03-03 |

### What to do
- Register a new GUIProxy handler `"autocad_set_attribute_values"`:
  ```csharp
  _guiProxy.RegisterHandler("autocad_set_attribute_values", (parameters) =>
  {
      // Parameters: { "values": { "product_name": "H-BEAM", "spec": "SS400", ... } }
      var valuesObj = parameters?["values"];
      if (valuesObj == null)
          return Task.FromResult<object?>(new { success = false, error = "No values provided" });

      var values = JsonSerializer.Deserialize<Dictionary<string, string>>(valuesObj.ToString()!);
      if (values == null || values.Count == 0)
          return Task.FromResult<object?>(new { success = false, error = "Empty values" });

      int updatedLayouts = 0;

      // Iterate all layouts (excluding "Model")
      dynamic layouts = _acadDoc.Layouts;
      for (int i = 0; i < (int)layouts.Count; i++)
      {
          dynamic layout = layouts.Item(i);
          if (layout.Name == "Model") continue;

          dynamic block = layout.Block;
          for (int j = 0; j < (int)block.Count; j++)
          {
              dynamic entity = block.Item(j);
              if ((int)entity.ObjectName == "AcDbBlockReference" && (bool)entity.HasAttributes)
              {
                  dynamic attrs = entity.GetAttributes();
                  bool isTargetBlock = false;

                  // First pass: check if this is the right block
                  foreach (dynamic attr in attrs)
                  {
                      if (attr.TagString == "product_name" || attr.TagString == "pr_no")
                      {
                          isTargetBlock = true;
                          break;
                      }
                  }

                  if (isTargetBlock)
                  {
                      // Second pass: set values
                      foreach (dynamic attr in attrs)
                      {
                          string tag = attr.TagString;
                          if (values.TryGetValue(tag, out var val))
                          {
                              attr.TextString = val;
                          }
                      }
                      updatedLayouts++;
                      break; // One block per layout
                  }
              }
          }
      }

      return Task.FromResult<object?>(new { success = true, updated_layouts = updatedLayouts });
  });
  ```
- The handler accepts a `values` dictionary with tag→value pairs and writes them to all layouts
- Returns `{ success: true, updated_layouts: N }` on success
- Uses `Task.FromResult` (sync) — Pure STA Mode pattern
- Mirrors Python's per-layout iteration: `for layout in layouts: block = get_attribute_block(layout); set_attribute_value(block, tag, val)`

### How to verify
- [ ] Handler registered with name `"autocad_set_attribute_values"` (AC-03)
- [ ] Writes all 7 tag→value pairs to attribute blocks (AC-01, AC-03)
- [ ] Iterates all layouts excluding "Model" (AC-04)
- [ ] Runs on STA thread via Task.FromResult (AC-05)
- [ ] Returns success count (AC-06)

---

## TASK-009-03-03: Implement SubmitCommand in ParameterConfigViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/ParameterConfigViewModel.cs` |
| Estimate | M |
| Depends On | TASK-009-03-01, TASK-009-03-02, TASK-009-02-08 |
| Blocks | — |

### What to do
- Implement `SubmitAsync()`:
  ```csharp
  [RelayCommand(CanExecute = nameof(CanSubmit))]
  private async Task SubmitAsync()
  {
      IsSubmitting = true;
      StatusMessage = "正在寫入 AutoCAD 屬性...";
      try
      {
          // 1. Build attribute values dictionary
          var values = new Dictionary<string, string>
          {
              { "product_name", SelectedProduct!.Name },
              { "spec", SelectedSpec!.Value },
              { "product_catelog", SelectedCategory!.Value },
              { "operation_flow", SelectedOperationFlow!.Value },
              { "surface_treatment", SelectedSurfaceTreatment!.Value },
              { "color_name", SelectedColor!.Name },
              { "color_no", ColorNo }
          };

          // 2. Call GUIProxy handler to write to AutoCAD
          var parameters = new Dictionary<string, object?>
          {
              { "values", values }
          };
          var response = await _guiProxy.ExecuteInGuiAsync("autocad_set_attribute_values", parameters);

          // 3. Process response
          if (response.Success)
          {
              var updatedCount = /* extract updated_layouts from response */;
              StatusMessage = $"已成功更新 {updatedCount} 個佈局的參數。";
          }
          else
          {
              StatusMessage = $"寫入失敗: {response.Error}";
          }
      }
      catch (Exception ex)
      {
          StatusMessage = $"更新 AutoCAD 屬性時發生錯誤: {ex.Message}";
      }
      finally
      {
          IsSubmitting = false;
          SubmitCommand.NotifyCanExecuteChanged();
      }
  }
  ```
- Submit is disabled during execution (`IsSubmitting` check in `CanSubmit`)
- Error handling follows established pattern (try-catch, user-friendly message)

### How to verify
- [ ] Submit writes all 7 attribute values via GUIProxy (AC-01)
- [ ] Success message shows layout count (AC-06)
- [ ] Failure message shows error details (AC-07)
- [ ] Submit is disabled during execution
- [ ] All COM operations go through GUIProxy (AC-05)

---

## TASK-009-03-04: Implement project_id resolution flow for color API

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/ParameterConfigViewModel.cs` |
| Estimate | M |
| Depends On | TASK-009-02-01 |
| Blocks | — |

### What to do
- Implement `LoadColorsAsync()` method for loading colors when project_id is available:
  ```csharp
  private async Task LoadColorsAsync()
  {
      // 1. Get project_id from AutoCAD layout attributes
      var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_layouts_values", null);
      if (!response.Success) return;

      // 2. Extract pr_no from first layout's attributes
      // (reuse existing layout data parsing from AutoCADService)
      var prNo = /* extract pr_no from response */;

      // 3. Look up project_id from Odoo using pr_no
      // Option A: Use existing IOdooService method if available
      // Option B: Parse project_id from layout attribute directly if stored

      if (projectId > 0)
      {
          var settings = await _settingsService.LoadSettingsAsync();
          var (baseUrl, endpointPath) = ParseSwaggerUrl(settings.SwaggerUrl);
          var colors = await _odooService.GetColorsViaApiAsync(
              baseUrl, endpointPath, settings.Database, settings.UserToken, projectId);

          Colors.Clear();
          foreach (var c in colors) Colors.Add(c);
      }
      else
      {
          StatusMessage = "無法取得專案 ID，顏色選項無法載入。";
      }
  }
  ```
- Called after `LoadOptionsAsync` completes AND AutoCAD is connected
- If AutoCAD is not connected, color section shows a "請先連接 AutoCAD" message
- If no project_id found, color dropdowns remain empty with a warning message

### How to verify
- [ ] project_id extracted from AutoCAD layout attributes (AC-09)
- [ ] Colors loaded when project_id is available
- [ ] Graceful handling when project_id is unavailable
- [ ] Color section shows appropriate message when disabled

---

## TASK-009-03-05: Unit tests for SubmitCommand (success, failure, disabled states)

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/ViewModels/ParameterConfigViewModelTests.cs` |
| Estimate | M |
| Depends On | TASK-009-03-03 |
| Blocks | — |

### What to do
- Add test methods to `ParameterConfigViewModelTests`:
  1. `Submit_AllFieldsFilled_CallsSetAttributeValues` — mock GUIProxy, verify handler called with correct values dictionary
  2. `Submit_Success_ShowsSuccessMessage` — mock successful response, verify StatusMessage contains layout count
  3. `Submit_Failure_ShowsErrorMessage` — mock error response, verify StatusMessage contains error
  4. `Submit_SetsIsSubmitting_DuringExecution` — verify IsSubmitting transitions
  5. `Submit_DisabledWhenFieldsMissing_CannotExecute` — verify CanSubmit returns false
  6. `Submit_DisabledDuringSubmission_CannotExecute` — verify CanSubmit returns false when IsSubmitting
  7. `Submit_CorrectAttributeDictionary_AllSevenTags` — verify all 7 tags present in values parameter

### How to verify
- [ ] All 7 tests pass
- [ ] Tests cover success, failure, and disabled states (AC-01, AC-06, AC-07, AC-08)

---

## TASK-009-03-06: Unit tests for handler registration (get_attribute_block, set_attribute_values)

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/AutoCADHandlerTests.cs` |
| Estimate | S |
| Depends On | TASK-009-03-01, TASK-009-03-02 |
| Blocks | — |

### What to do
- Add or extend `AutoCADHandlerTests`:
  1. `RegisterHandlers_ContainsGetAttributeBlock` — verify handler registered with correct name
  2. `RegisterHandlers_ContainsSetAttributeValues` — verify handler registered with correct name
  3. `SetAttributeValues_NullValues_ReturnsError` — verify error handling for null input
  4. `SetAttributeValues_EmptyValues_ReturnsError` — verify error handling for empty dict
- These tests verify handler registration without requiring actual AutoCAD COM objects

### How to verify
- [ ] All 4 tests pass
- [ ] Handler names are correctly registered (AC-02, AC-03)

---

## Dependency Graph
```
TASK-009-03-01 (get_attribute_block handler)
    |
    +---> TASK-009-03-02 (set_attribute_values handler)
    |         |
    |         +---> TASK-009-03-03 (SubmitCommand)
    |         |         |
    |         |         +---> TASK-009-03-05 (Submit tests)
    |         |
    |         +---> TASK-009-03-06 (Handler tests)
    |
    +---> TASK-009-03-06 (Handler tests)

TASK-009-03-04 (project_id resolution — independent of handlers)
```
