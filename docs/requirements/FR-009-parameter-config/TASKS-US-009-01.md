# TASKS: US-009-01 — Fetch Odoo Parameters

> **Parent US**: [US-009-01](US-009-01-fetch-odoo-parameters.md)
> **Parent FR**: [FR-009](FR-009-parameter-config.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 2S + 3M + 1S (total: 3S + 3M)
> **Status**: Not Started

## Prerequisites
- [ ] Sprint 3 infrastructure complete (IOdooService, Swagger API pattern, BasicAuth)
- [ ] `GetProductsViaApiAsync()` exists (Sprint 5)
- [ ] Swagger endpoint resolution pattern available (ParseSwaggerUrl/ResolveSwaggerEndpoint)

## Acceptance Criteria
- [ ] AC-01: `GetSetupViaApiAsync(setupName)` fetches setup values via Swagger `get_setup_v2` endpoint
- [ ] AC-02: `GetColorsViaApiAsync(projectId)` fetches color options via Swagger `get_color_v2` endpoint
- [ ] AC-03: Setup values returned as `IReadOnlyList<OdooSetupValue>` with `Value` and `SetupName`
- [ ] AC-04: Colors returned as `IReadOnlyList<OdooColor>` with `Name`, `ColorNo`, `ProjectId`
- [ ] AC-05: Both methods use BasicAuth + PATCH pattern
- [ ] AC-06: API failures return empty lists and log errors
- [ ] AC-07: `GetProductsViaApiAsync()` reused for material data

---

## TASK-009-01-01: Define OdooSetupValue and OdooColor record types

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/IOdooService.cs` |
| Estimate | S |
| Depends On | — |
| Blocks | TASK-009-01-02, TASK-009-01-03, TASK-009-01-04 |

### What to do
- Add two new record types alongside existing `OdooProduct` record in `IOdooService.cs`:
  ```csharp
  public record OdooSetupValue(string Value, string SetupName);
  public record OdooColor(string Name, string ColorNo, int ProjectId);
  ```
- `OdooSetupValue` represents a single setup option (e.g., spec="SS400", product_catelog="H型鋼")
- `OdooColor` represents a project-specific color (e.g., Name="RAL 7035", ColorNo="7035")
- Place these records near the existing `OdooProduct` record for consistency

### How to verify
- [ ] `OdooSetupValue` record exists with `Value` and `SetupName` properties (AC-03)
- [ ] `OdooColor` record exists with `Name`, `ColorNo`, and `ProjectId` properties (AC-04)
- [ ] Solution builds without errors

---

## TASK-009-01-02: Add GetSetupViaApiAsync and GetColorsViaApiAsync to IOdooService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/IOdooService.cs` |
| Estimate | S |
| Depends On | TASK-009-01-01 |
| Blocks | TASK-009-01-03, TASK-009-01-04 |

### What to do
- Add two new method signatures to the `IOdooService` interface:
  ```csharp
  /// <summary>
  /// Fetches setup values from Odoo via Swagger get_setup_v2 endpoint.
  /// </summary>
  /// <param name="setupName">Setup type: 'spec', 'product_catelog', 'operation_flow', 'surface_treatment'</param>
  Task<IReadOnlyList<OdooSetupValue>> GetSetupViaApiAsync(
      string baseUrl, string endpointPath, string database, string userToken, string setupName);

  /// <summary>
  /// Fetches project-specific colors from Odoo via Swagger get_color_v2 endpoint.
  /// </summary>
  Task<IReadOnlyList<OdooColor>> GetColorsViaApiAsync(
      string baseUrl, string endpointPath, string database, string userToken, int projectId);
  ```
- Follow the same parameter pattern as `GetProductsViaApiAsync(baseUrl, endpointPath, database, userToken)`
- `GetSetupViaApiAsync` adds `setupName` parameter for the setup type filter
- `GetColorsViaApiAsync` adds `projectId` parameter for the project filter

### How to verify
- [ ] Both method signatures exist in `IOdooService` interface (AC-01, AC-02)
- [ ] Parameter signatures match the established pattern (AC-05)
- [ ] Solution builds (OdooService will need stub implementations)

---

## TASK-009-01-03: Implement GetSetupViaApiAsync in OdooService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | TASK-009-01-01, TASK-009-01-02 |
| Blocks | TASK-009-01-05 |

### What to do
- Implement `GetSetupViaApiAsync` following the `GetProductsViaApiAsync` pattern exactly:
  1. Build the full URL from `baseUrl` + `endpointPath`
  2. Create `HttpRequestMessage` with `HttpMethod.Patch`
  3. Set BasicAuth header: `Authorization: Basic base64({database}:{userToken})`
  4. Build JSON body matching Python's Swagger call:
     ```json
     {
       "method_name": "get_setup_v2",
       "args": [[["setup_name", "=", "{setupName}"]]],
       "kwargs": {"user_token": "{userToken}"}
     }
     ```
  5. Send request, parse response JSON
  6. Map response items to `OdooSetupValue` records
  7. Return `IReadOnlyList<OdooSetupValue>`
- Handle errors: catch `HttpRequestException`, log error, return empty list
- Handle empty/null response gracefully

### How to verify
- [ ] Method builds correct Swagger request with `get_setup_v2` method name (AC-01)
- [ ] Uses BasicAuth + PATCH pattern (AC-05)
- [ ] Returns empty list on failure (AC-06)
- [ ] Returns `IReadOnlyList<OdooSetupValue>` with correct properties (AC-03)

---

## TASK-009-01-04: Implement GetColorsViaApiAsync in OdooService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | TASK-009-01-01, TASK-009-01-02 |
| Blocks | TASK-009-01-06 |

### What to do
- Implement `GetColorsViaApiAsync` following the same pattern as `GetSetupViaApiAsync`:
  1. Build the full URL from `baseUrl` + `endpointPath`
  2. Create `HttpRequestMessage` with `HttpMethod.Patch`
  3. Set BasicAuth header
  4. Build JSON body:
     ```json
     {
       "method_name": "get_color_v2",
       "args": [[["job_project_id", "=", {projectId}]]],
       "kwargs": {"user_token": "{userToken}"}
     }
     ```
  5. Send request, parse response JSON
  6. Map response items to `OdooColor` records (Name, ColorNo, ProjectId)
  7. Return `IReadOnlyList<OdooColor>`
- Handle errors: catch `HttpRequestException`, log error, return empty list
- Note: `project_id` is an integer, not a string — ensure correct JSON serialization

### How to verify
- [ ] Method builds correct Swagger request with `get_color_v2` method name (AC-02)
- [ ] Uses BasicAuth + PATCH pattern (AC-05)
- [ ] Returns empty list on failure (AC-06)
- [ ] Returns `IReadOnlyList<OdooColor>` with correct properties (AC-04)

---

## TASK-009-01-05: Unit tests for GetSetupViaApiAsync

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/OdooServiceSetupTests.cs` |
| Estimate | M |
| Depends On | TASK-009-01-03 |
| Blocks | — |

### What to do
- Create test class `OdooServiceSetupTests` with mocked `HttpMessageHandler`:
  1. `GetSetupViaApiAsync_ValidResponse_ReturnsSetupValues` — mock 200 response with sample setup data, verify correct parsing
  2. `GetSetupViaApiAsync_EmptyResponse_ReturnsEmptyList` — mock empty result array
  3. `GetSetupViaApiAsync_HttpError_ReturnsEmptyList` — mock 500 response, verify empty list returned (not exception)
  4. `GetSetupViaApiAsync_BuildsCorrectRequest_WithSetupName` — verify request URL, method (PATCH), BasicAuth header, body contains correct `setup_name` filter
  5. `GetSetupViaApiAsync_SpecType_ReturnsSpecValues` — test with setupName="spec"
  6. `GetSetupViaApiAsync_OperationFlowType_ReturnsFlowValues` — test with setupName="operation_flow"
- Use `MockHttpMessageHandler` pattern established in existing tests

### How to verify
- [ ] All 6 tests pass
- [ ] Tests cover success, empty, and error scenarios (AC-01, AC-03, AC-06)
- [ ] Tests verify correct Swagger request format (AC-05)

---

## TASK-009-01-06: Unit tests for GetColorsViaApiAsync

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/OdooServiceColorTests.cs` |
| Estimate | S |
| Depends On | TASK-009-01-04 |
| Blocks | — |

### What to do
- Create test class `OdooServiceColorTests` with mocked `HttpMessageHandler`:
  1. `GetColorsViaApiAsync_ValidResponse_ReturnsColors` — mock 200 response with sample color data, verify Name/ColorNo/ProjectId parsing
  2. `GetColorsViaApiAsync_EmptyResponse_ReturnsEmptyList` — mock empty result
  3. `GetColorsViaApiAsync_HttpError_ReturnsEmptyList` — mock 500 response
  4. `GetColorsViaApiAsync_BuildsCorrectRequest_WithProjectId` — verify body contains `job_project_id` filter with integer value
- Follow the same mock pattern as TASK-009-01-05

### How to verify
- [ ] All 4 tests pass
- [ ] Tests cover success, empty, and error scenarios (AC-02, AC-04, AC-06)
- [ ] Tests verify correct Swagger request format (AC-05)

---

## Dependency Graph
```
TASK-009-01-01 (Define records)
    |
    +---> TASK-009-01-02 (Interface methods)
              |
              +---> TASK-009-01-03 (GetSetupViaApiAsync impl)
              |         |
              |         +---> TASK-009-01-05 (Setup tests)
              |
              +---> TASK-009-01-04 (GetColorsViaApiAsync impl)
                        |
                        +---> TASK-009-01-06 (Color tests)
```
