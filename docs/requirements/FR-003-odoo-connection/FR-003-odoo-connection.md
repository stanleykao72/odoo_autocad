# FR-003: Odoo Integration Page

> **Document Version**: 1.0
> **Last Updated**: 2026-02-06
> **Status**: Not Started
> **Priority**: P1

## 1. Overview

The Odoo Integration Page provides the primary interface for connecting the desktop application to an Odoo ERP server via REST API. It manages server credentials, connection lifecycle, product catalog synchronization, project lookup, and data synchronization state. This page is the gateway through which all Odoo-dependent workflows (BOQ import, Purchase Requisition generation, product validation) are enabled.

The Python implementation (`utility/util_odoo.py`) communicates with Odoo through a Bravado/Swagger client that consumes a Swagger JSON endpoint. Authentication uses HTTP Basic Auth where the credentials are a Base64-encoded string of `{db_name}:{token}`. All API calls route through a single Swagger operation (`job_working_plan_boq.callMethodForJobWorkingPlanBoqModel`) with a `method_name` parameter to dispatch to specific server-side methods.

The C# implementation (`IOdooService.cs`, 262 lines) replaces the Swagger/Bravado client with `HttpClient` and Odoo's standard JSON-RPC protocol (`/web/dataset/call_kw`). Authentication shifts from token-based Basic Auth to session-based login via `/web/session/authenticate`. Configuration is loaded from `appsettings.json` instead of YAML files. The interface exposes structured async methods for connection management, project operations, product operations, BOQ operations, Purchase Requisition operations, and general synchronization.

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-003-01 | CAD Engineer | Enter Odoo server URL and credentials on a settings form | I can connect to our company's Odoo instance |
| US-003-02 | CAD Engineer | Test the Odoo connection before saving settings | I know the credentials are valid before proceeding |
| US-003-03 | CAD Engineer | See the current Odoo connection status at all times | I know whether data operations will succeed |
| US-003-04 | CAD Engineer | Disconnect from Odoo explicitly | I can switch to a different server or database |
| US-003-05 | CAD Engineer | Synchronize the product catalog from Odoo | I have an up-to-date list of products for BOQ entries |
| US-003-06 | CAD Engineer | Search and filter products by name or code | I can quickly find the product I need |
| US-003-07 | CAD Engineer | Search for a project by PR number or name | I can associate AutoCAD drawings with the correct Odoo project |
| US-003-08 | CAD Engineer | See the last successful sync timestamp | I know how fresh my local product data is |
| US-003-09 | Project Manager | View server information (version, database name) after connecting | I can verify we are connected to the correct environment |
| US-003-10 | System Admin | Configure Odoo connection settings in appsettings.json | The application can start with pre-configured defaults |
| US-003-11 | CAD Engineer | Filter products by category | I only see products relevant to construction/engineering |
| US-003-12 | CAD Engineer | Have connection credentials persisted securely between sessions | I do not have to re-enter credentials every launch |

## 3. Python Reference

### Source Files
- `utility/util_odoo.py` - Odoo API client with Bravado/Swagger integration
- `config/server.yaml` - Server connection configuration (host, db_name, url, token_file, auxiliary file paths)
- `config/token.yaml` - User authentication token and server_file reference
- `models/server.py` - SQLAlchemy models for local configuration caching
- `forms/form_main_modern.py` - Connection status display and quick-connect button

### Key Functions

| Python Function | Purpose | C# Equivalent |
|----------------|---------|---------------|
| `UtilOdoo.__init__(odoo_connection, log_util)` | Constructor; reads config dict and calls `connect_odoo()` | `OdooService(ILogger)` constructor + `ConnectAsync()` |
| `connect_odoo(odoo_connection)` | Establishes Bravado/Swagger client with Basic Auth (Base64 of `db_name:token`) | `ConnectAsync(serverUrl, database, username, password)` |
| `connected_odoo()` | Returns `True` if `self.odoo` is not `None` | `IsConnected` property |
| `string_to_base64(input_string)` | Encodes string to Base64 for Authorization header | Handled internally by `HttpClient` |
| `get_project(pr_no)` | Calls `get_project_v2` with domain `[['name', '=', pr_no]]` | `SearchProjectsAsync(searchTerm)` / `GetProjectAsync(id)` |
| `get_product()` | Calls `get_product_v2` with domain `[('categ_id', 'child_of', 27), ('active', '=', True)]` | `GetProductsAsync()` / `GetProductsByCategoryAsync(categoryId)` |
| `get_setup(setup_name)` | Calls `get_setup_v2` to retrieve configuration by name | No direct equivalent; configuration from `appsettings.json` |
| `get_color(project_id)` | Calls `get_color_v2` with domain `[('job_project_id', '=', project_id)]` | Custom method to be implemented |
| `import2boq(layout_dict)` | Calls `import2boq_v2` to push BOQ data | `ImportToBOQAsync(entries)` |
| `boq2pr(layout_dict)` | Calls `boq2pr_v2` to convert BOQ to Purchase Requisition | `ConvertBOQToPRAsync(projectId, boqEntryIds)` |
| `search_products(query, limit=50)` | Searches products with text query and result limit | `SearchProductsAsync(searchTerm)` |
| `push_boq_data(project_id, boq_data)` | Pushes BOQ entries to a specific project | `ImportToBOQAsync(entries)` |

### Python Configuration Structure

**server.yaml:**
```yaml
server:
  host: 'odoo-server.example.com'
  db_name: 'odoo-database-name'
  url: 'https://odoo-server.example.com/api/v1/boq_import_api/swagger.json?token=...&db=...'
  token_file: 'c:\odoo\config\token.yaml'
  StripMtext_file: '...\StripMtext v5-0b.lsp'
  read_csv_file: '...\read_csv.lsp'
  contract_product_file: '...\contract_product.lsp'
```

**token.yaml:**
```yaml
user:
  token: '<uuid-token>'
  server_file: 'path/to/server_prod.yaml'
```

### Python API Call Pattern
All Odoo operations use a single Swagger endpoint dispatched by `method_name`:
```python
self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
    method_name="<method_v2>",
    body={
        "args": [<positional_args>],
        "kwargs": {"user_token": self.user_token},
        "context": {}
    },
    _request_options=self.requestOptions
).response().incoming_response.json()
```

## 4. Functional Requirements

### Connection Management

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-003-001 | The page SHALL provide input fields for Server URL, Database Name, Username, and Password/API Key | Must |
| FR-003-002 | The page SHALL provide a "Test Connection" button that validates credentials without persisting the session | Must |
| FR-003-003 | The page SHALL provide a "Connect" button that establishes an authenticated session to Odoo via JSON-RPC (`/web/session/authenticate`) | Must |
| FR-003-004 | The page SHALL provide a "Disconnect" button that terminates the current Odoo session and clears session state | Must |
| FR-003-005 | The `IsConnected` property SHALL return `true` only when a valid session ID exists and authentication has succeeded | Must |
| FR-003-006 | Connection status SHALL be displayed as a colored indicator (green = connected, red/gray = disconnected) with descriptive text | Must |
| FR-003-007 | On successful connection, the page SHALL display server version, database name, and authenticated username in a status panel | Should |
| FR-003-008 | Connection settings SHALL be pre-populated from `appsettings.json` on application launch (`Odoo:ServerUrl`, `Odoo:Database`, `Odoo:Username`) | Must |
| FR-003-009 | The Password/API Key field SHALL use `PasswordBox` (masked input) and SHALL NOT be stored in `appsettings.json` in plaintext | Must |
| FR-003-010 | Connection timeout SHALL be configurable via `Odoo:TimeoutSeconds` in `appsettings.json` (default: 30 seconds) | Should |

### Product Management

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-003-011 | The page SHALL provide a "Sync Products" button that fetches the full product catalog from Odoo via `GetProductsAsync()` | Must |
| FR-003-012 | Products SHALL be displayed in a searchable, scrollable `DataGrid` with columns: ID, Name, Code, Description, Unit Price, UOM, Category | Must |
| FR-003-013 | The page SHALL provide a search box that filters products by name or code in real-time using `SearchProductsAsync(searchTerm)` | Must |
| FR-003-014 | The page SHALL support filtering products by category via `GetProductsByCategoryAsync(categoryId)`, matching the Python filter of `categ_id child_of 27` for construction products | Should |
| FR-003-015 | Product sync results SHALL display a summary: total records fetched, timestamp of sync, and any errors encountered | Should |
| FR-003-016 | Synced product data SHALL be cached locally (via SQLite or in-memory) to reduce repeated API calls within the same session | Should |

### Project Lookup

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-003-017 | The page SHALL provide a project search field that queries Odoo by project name or code via `SearchProjectsAsync(searchTerm)` | Must |
| FR-003-018 | Project search results SHALL display in a list with columns: ID, Name, Code, State, Date Start, Date End | Must |
| FR-003-019 | The user SHALL be able to select a project from search results to set it as the active project context for downstream operations (BOQ, PR) | Must |
| FR-003-020 | The page SHALL support retrieving a single project by ID via `GetProjectAsync(projectId)` when navigating from other pages | Should |

### Data Synchronization

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-003-021 | The page SHALL display the last synchronization timestamp via `GetLastSyncTime()`, formatted in the user's local timezone | Must |
| FR-003-022 | The page SHALL provide a general "Sync to Odoo" action supporting sync types: `parameters`, `boq`, and `project` via `SyncToOdooAsync(data, syncType)` | Must |
| FR-003-023 | Sync operations SHALL return a `SyncResult` displaying: Success/Failure status, RecordsProcessed, RecordsCreated, RecordsUpdated, RecordsFailed, and any error messages | Must |
| FR-003-024 | Sync operations SHALL be performed asynchronously with a progress indicator and SHALL NOT block the UI thread | Must |
| FR-003-025 | If a sync operation partially fails (some records succeed, some fail), the page SHALL display both the successful count and the list of errors | Should |
| FR-003-026 | The page SHALL disable the Sync and Product buttons when `IsConnected` is `false` | Must |

## 5. UI Wireframe Description

```
+----------------------------------------------------------------------+
|                      Odoo Integration Page                           |
+----------------------------------------------------------------------+
|                                                                      |
|  Connection Settings                                                 |
|  +----------------------------------------------------------------+  |
|  | Server URL:  [https://odoo.example.com_________________]       |  |
|  | Database:    [production_______________________________ ]       |  |
|  | Username:    [admin___________________________________ ]       |  |
|  | Password:    [********________________________________ ]       |  |
|  |                                                                |  |
|  | [Test Connection]  [Connect]  [Disconnect]                     |  |
|  +----------------------------------------------------------------+  |
|                                                                      |
|  Connection Status                                                   |
|  +----------------------------------------------------------------+  |
|  | Status: (o) Connected          Server Version: 17.0            |  |
|  | Database: production           Username: admin                 |  |
|  | Last Sync: 2026-02-06 14:30    Session: Active                 |  |
|  +----------------------------------------------------------------+  |
|                                                                      |
|  +-------------------------------+  +-----------------------------+  |
|  | Project Search                |  | Sync Controls               |  |
|  | [Search by PR No / Name____]  |  | Sync Type: [Products  v]   |  |
|  | [Search]                      |  | [Sync Now]                  |  |
|  |                               |  |                             |  |
|  | Results:                      |  | Last Sync: 2026-02-06 14:30 |  |
|  | +---------------------------+ |  | Records: 1,247 products     |  |
|  | | ID  | Name    | State    | |  | Status: Success             |  |
|  | |-----|---------|----------| |  |                             |  |
|  | | 42  | PR-2025 | active   | |  | Sync History:               |  |
|  | | 87  | PR-2026 | active   | |  | - Products  14:30  OK       |  |
|  | +---------------------------+ |  | - BOQ       14:25  OK       |  |
|  | [Select Project]             |  | - Params    14:20  2 errors |  |
|  +-------------------------------+  +-----------------------------+  |
|                                                                      |
|  Product Catalog                                                     |
|  +----------------------------------------------------------------+  |
|  | Search: [pipe_______________________] Category: [All       v]  |  |
|  | [Search] [Sync Products]                                       |  |
|  |                                                                |  |
|  | +------------------------------------------------------------+ |  |
|  | | ID  | Name         | Code  | UOM | Price  | Category       | |  |
|  | |-----|--------------|-------|-----|--------|----------------| |  |
|  | | 101 | Pipe 100mm   | P100  | m   | 45.00  | Piping         | |  |
|  | | 102 | Pipe 150mm   | P150  | m   | 62.50  | Piping         | |  |
|  | | 103 | Pipe 200mm   | P200  | m   | 85.00  | Piping         | |  |
|  | | 204 | Valve 50mm   | V050  | pcs | 120.00 | Valves         | |  |
|  | | ... | ...          | ...   | ... | ...    | ...            | |  |
|  | +------------------------------------------------------------+ |  |
|  | Showing 50 of 1,247 products                    [Load More]    |  |
|  +----------------------------------------------------------------+  |
|                                                                      |
+----------------------------------------------------------------------+
```

### Layout Details
- **Connection Settings**: `GroupBox` with `Grid` layout; 4 labeled input rows and a button bar
- **Connection Status**: `Border` panel with status indicators using colored `Ellipse` elements and text labels
- **Project Search**: Left panel using `StackPanel` with `TextBox`, `Button`, and `DataGrid` for results
- **Sync Controls**: Right panel with `ComboBox` for sync type selection, action button, and sync history `ListView`
- **Product Catalog**: Full-width bottom section with search bar, category filter `ComboBox`, and `DataGrid` with virtualization for large datasets
- **All panels**: Use WPF `GroupBox` or `Border` with rounded corners and consistent theme from `App.xaml` resources

## 6. Data Model

### Input Data (ViewModel Properties)

```csharp
public class OdooConnectionViewModel : ObservableObject
{
    // Connection Settings (bindable input fields)
    public string ServerUrl { get; set; }          // From appsettings.json Odoo:ServerUrl
    public string Database { get; set; }           // From appsettings.json Odoo:Database
    public string Username { get; set; }           // From appsettings.json Odoo:Username
    public string Password { get; set; }           // Masked input, not persisted in config
    public int TimeoutSeconds { get; set; }        // From appsettings.json Odoo:TimeoutSeconds

    // Connection Status (read-only display)
    public bool IsConnected { get; set; }
    public string ConnectionStatusText { get; set; }       // "Connected" / "Disconnected" / "Connecting..."
    public string ServerVersion { get; set; }
    public string ConnectedDatabase { get; set; }
    public string ConnectedUsername { get; set; }
    public string ErrorMessage { get; set; }

    // Project Search
    public string ProjectSearchTerm { get; set; }
    public ObservableCollection<OdooProject> ProjectSearchResults { get; set; }
    public OdooProject? SelectedProject { get; set; }

    // Product Catalog
    public string ProductSearchTerm { get; set; }
    public int? SelectedCategoryId { get; set; }
    public ObservableCollection<OdooProduct> Products { get; set; }
    public OdooProduct? SelectedProduct { get; set; }
    public int TotalProductCount { get; set; }

    // Sync Status
    public DateTime? LastSyncTime { get; set; }
    public string LastSyncStatus { get; set; }             // "Success" / "Failed" / "Partial"
    public string SelectedSyncType { get; set; }           // "parameters" / "boq" / "project"
    public bool IsSyncing { get; set; }
    public SyncResult? LastSyncResult { get; set; }
    public ObservableCollection<SyncHistoryEntry> SyncHistory { get; set; }

    // Commands
    public IAsyncRelayCommand TestConnectionCommand { get; }
    public IAsyncRelayCommand ConnectCommand { get; }
    public IAsyncRelayCommand DisconnectCommand { get; }
    public IAsyncRelayCommand SearchProjectsCommand { get; }
    public IRelayCommand SelectProjectCommand { get; }
    public IAsyncRelayCommand SyncProductsCommand { get; }
    public IAsyncRelayCommand SearchProductsCommand { get; }
    public IAsyncRelayCommand FilterByCategoryCommand { get; }
    public IAsyncRelayCommand SyncToOdooCommand { get; }
}

/// <summary>
/// Entry in the sync history log displayed on the page.
/// </summary>
public class SyncHistoryEntry
{
    public string SyncType { get; set; }          // "Products" / "BOQ" / "Parameters"
    public DateTime Timestamp { get; set; }
    public bool Success { get; set; }
    public int RecordsProcessed { get; set; }
    public int ErrorCount { get; set; }
    public string? Summary { get; set; }
}
```

### C# Records from IOdooService.cs

```csharp
public record OdooStatus(
    bool IsConnected,
    string? ServerUrl,
    string? Database,
    string? Username,
    string? Version,
    string? ErrorMessage);

public record OdooProject(
    int Id,
    string Name,
    string? Code,
    string State,
    DateTime? DateStart,
    DateTime? DateEnd);

public record OdooProduct(
    int Id,
    string Name,
    string? Code,
    string? Description,
    decimal? ListPrice,
    string? UnitOfMeasure,
    string? Category);

public record SyncResult(
    bool Success,
    int RecordsProcessed,
    int RecordsCreated,
    int RecordsUpdated,
    int RecordsFailed,
    List<string>? Errors);
```

## 7. API/Service Dependencies

| Service | Interface | Methods Used | Purpose |
|---------|-----------|-------------|---------|
| Odoo Service | `IOdooService` | `ConnectAsync()`, `DisconnectAsync()`, `TestConnectionAsync()`, `GetStatusAsync()`, `IsConnected` | Connection lifecycle management |
| Odoo Service | `IOdooService` | `GetProjectsAsync()`, `GetProjectAsync()`, `SearchProjectsAsync()` | Project lookup and selection |
| Odoo Service | `IOdooService` | `GetProductsAsync()`, `GetProductAsync()`, `SearchProductsAsync()`, `GetProductsByCategoryAsync()` | Product catalog retrieval and search |
| Odoo Service | `IOdooService` | `ImportToBOQAsync()`, `GetBOQEntriesAsync()`, `UpdateBOQEntryAsync()`, `DeleteBOQEntryAsync()` | BOQ data import and management |
| Odoo Service | `IOdooService` | `ConvertBOQToPRAsync()`, `GetPurchaseRequisitionsAsync()`, `SubmitPRAsync()` | Purchase Requisition generation |
| Odoo Service | `IOdooService` | `SyncToOdooAsync()`, `GetLastSyncTime()` | General synchronization operations |
| Configuration | `IConfiguration` | `GetSection("Odoo")` | Load default connection settings from `appsettings.json` |
| Navigation Service | `INavigationService` | `NavigateTo()` | Navigate to BOQ/PR pages after project selection |
| Credential Store | `ICredentialService` (optional) | `SaveCredentials()`, `LoadCredentials()` | Secure credential persistence via Windows Credential Manager or DPAPI |
| Logger | `ILogger<OdooConnectionViewModel>` | `LogInformation()`, `LogWarning()`, `LogError()` | Structured logging for connection events and errors |

## 8. Validation Rules

### Connection Settings Validation

| Rule | Field | Description |
|------|-------|-------------|
| VR-003-001 | Server URL | SHALL be a valid HTTPS URL (pattern: `^https?://[^\s/]+`). HTTP URLs SHALL trigger a warning but not be blocked. |
| VR-003-002 | Server URL | SHALL NOT be empty. "Connect" and "Test Connection" buttons SHALL be disabled when empty. |
| VR-003-003 | Server URL | SHALL have trailing slashes stripped before use (matching `OdooService.ConnectAsync` behavior: `serverUrl.TrimEnd('/')`) |
| VR-003-004 | Database | SHALL NOT be empty. Must contain only alphanumeric characters, hyphens, and underscores. |
| VR-003-005 | Username | SHALL NOT be empty when using session-based authentication. |
| VR-003-006 | Password | SHALL NOT be empty. Minimum length of 1 character. SHALL be masked in UI. |
| VR-003-007 | Timeout | SHALL be a positive integer between 5 and 120 seconds. Default: 30. |

### Sync Operation Validation

| Rule | Description |
|------|-------------|
| VR-003-008 | Product sync SHALL only be available when `IsConnected` is `true` |
| VR-003-009 | Project search SHALL only be available when `IsConnected` is `true` |
| VR-003-010 | "Sync to Odoo" SHALL require a valid `SelectedSyncType` from the allowed set: `parameters`, `boq`, `project` |
| VR-003-011 | BOQ sync SHALL require a selected project (`SelectedProject` is not null) |
| VR-003-012 | Product search term SHALL be at least 1 character before triggering server-side search (to prevent unbounded queries) |
| VR-003-013 | Concurrent sync operations SHALL be prevented; buttons SHALL be disabled while `IsSyncing` is `true` |
| VR-003-014 | Category filter SHALL accept `null` (all categories) or a valid positive integer category ID |

## 9. Error Handling

| Scenario | Error Source | User-Facing Message | Action |
|----------|-------------|---------------------|--------|
| Network unreachable | `HttpRequestException` | "Unable to reach Odoo server at {ServerUrl}. Please check your network connection and server URL." | Display in status panel. Enable retry via Connect button. |
| Authentication failed (invalid credentials) | JSON-RPC response with `uid = false` | "Authentication failed. Please verify your database name, username, and password." | Clear password field. Highlight credential inputs with error border. |
| Authentication failed (invalid database) | JSON-RPC error response | "Database '{Database}' not found on server. Please check the database name." | Highlight database field. |
| Connection timeout | `TaskCanceledException` / `TimeoutException` | "Connection timed out after {TimeoutSeconds} seconds. The server may be under heavy load." | Display timeout info. Suggest retry or increasing timeout in settings. |
| SSL/TLS certificate error | `HttpRequestException` (inner: `AuthenticationException`) | "SSL certificate validation failed for {ServerUrl}. Contact your system administrator." | Log full exception details for admin review. |
| Swagger endpoint not found (Python migration) | HTTP 404 on Swagger URL | "The legacy Swagger API endpoint was not found. The C# application uses JSON-RPC. Please verify the server URL points to the Odoo web interface (not the Swagger JSON endpoint)." | Show migration hint in error details. |
| Session expired during operation | JSON-RPC error mid-operation | "Your Odoo session has expired. Please reconnect." | Set `IsConnected = false`. Prompt reconnection. |
| Product sync partial failure | `SyncResult.RecordsFailed > 0` | "Product sync completed with {RecordsFailed} errors out of {RecordsProcessed} records." | Display error list in expandable panel. Log individual errors. |
| API rate limiting / server busy | HTTP 429 or 503 | "Odoo server is temporarily unavailable. Please try again in a few moments." | Implement exponential backoff for retries. Show retry countdown. |
| Invalid sync type | `SyncToOdooAsync` default case | "Unknown synchronization type: '{syncType}'. Supported types are: parameters, boq, project." | Log as warning. Display inline error. |
| BOQ import error | `ImportToBOQAsync` per-entry failure | "Failed to import BOQ entry '{ProductName}': {ErrorMessage}" | Display in SyncResult.Errors list. Continue processing remaining entries. |
| Project not found | Empty search results | "No projects found matching '{searchTerm}'. Try a different search term." | Display empty state message in results grid. |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/OdooConnectionPage.xaml` - WPF page with all connection, product, and sync UI elements
- `ViewModels/OdooConnectionViewModel.cs` - MVVM ViewModel with async commands and validation
- `OdooAutoCAD.Core/Odoo/IOdooService.cs` - Service interface (already exists, 262 lines)
- `OdooAutoCAD.Core/Odoo/OdooService.cs` - Service implementation (already exists, 675 lines)

### Swagger to JSON-RPC Migration

The Python application uses Bravado/Swagger to consume a custom Swagger endpoint at a URL like:
```
https://server/api/v1/boq_import_api/swagger.json?token=...&db=...
```

The C# application replaces this with Odoo's standard JSON-RPC protocol. All operations use `HttpClient.PostAsync` against two endpoints:

| Purpose | Python Endpoint | C# Endpoint |
|---------|----------------|-------------|
| Authentication | Basic Auth header on Swagger client | `POST /web/session/authenticate` |
| Server info | Not used | `POST /web/webclient/version_info` |
| All CRUD operations | `job_working_plan_boq.callMethodForJobWorkingPlanBoqModel` | `POST /web/dataset/call_kw` |

Each JSON-RPC call follows the pattern:
```json
{
    "jsonrpc": "2.0",
    "method": "call",
    "params": {
        "model": "<odoo.model.name>",
        "method": "<search_read|read|create|write|unlink>",
        "args": [...],
        "kwargs": { "fields": [...], "limit": N }
    },
    "id": 1
}
```

### Authentication Pattern

**Python (Token-based Basic Auth):**
1. Read `token` from `token.yaml` and `db_name` from `server.yaml`
2. Construct Basic Auth string: `Base64("{db_name}:{token}")`
3. Set `Authorization: Basic {base64_string}` header on all requests
4. No session management; each request is independently authenticated

**C# (Session-based JSON-RPC Auth):**
1. Read defaults from `appsettings.json` (`Odoo:ServerUrl`, `Odoo:Database`, `Odoo:Username`)
2. User provides password via UI (not stored in config)
3. Call `POST /web/session/authenticate` with `{ db, login, password }`
4. On success, extract `uid` from response; `HttpClient` retains session cookies
5. Subsequent requests authenticated via session cookie (managed by `HttpClient` cookie container)

### Configuration from appsettings.json

```json
{
    "Odoo": {
        "ServerUrl": "https://odoo.example.com",
        "Database": "production",
        "Username": "",
        "TimeoutSeconds": 30
    }
}
```

- `ServerUrl` pre-populates the URL field; user can override at runtime
- `Database` pre-populates the database field; user can override at runtime
- `Username` pre-populates if set; typically left blank for per-user entry
- `TimeoutSeconds` configures `HttpClient.Timeout`; defaults to 30 if missing
- Password is intentionally excluded from configuration for security

### MVVM Bindings

- Use `CommunityToolkit.Mvvm` for `ObservableObject`, `AsyncRelayCommand`, `RelayCommand`
- Connection status bound via `{Binding IsConnected}` with `BoolToVisibilityConverter` and `BoolToColorConverter`
- Input validation via `INotifyDataErrorInfo` on the ViewModel for real-time field validation
- Async commands use `IAsyncRelayCommand` to provide built-in `IsRunning` for progress indicators
- Product `DataGrid` bound to `ObservableCollection<OdooProduct>` with `VirtualizingStackPanel` for performance
- Sync history bound to `ObservableCollection<SyncHistoryEntry>` displayed in descending timestamp order

### Special Considerations

1. **Thread Safety**: `OdooService` uses a single `HttpClient` instance. All async operations must be thread-safe. The `HttpClient` is thread-safe for concurrent requests.

2. **Credential Security**: The password should not be stored in `appsettings.json`. Consider using Windows DPAPI (`ProtectedData`) or Windows Credential Manager for optional "Remember Me" functionality.

3. **Product Caching**: The Python version caches products to SQLite via SQLAlchemy. The C# version should cache to the local SQLite database (`Database:ConnectionString` in config) or use an in-memory cache with `IMemoryCache` to avoid repeated full catalog downloads.

4. **Category Filtering**: The Python code hardcodes `categ_id child_of 27` for construction products. The C# version should make this configurable (either via `appsettings.json` or a UI dropdown populated from Odoo's category tree).

5. **Error Response Handling**: The Python API returns `error_code` and `error_message` keys in the response JSON. The C# JSON-RPC implementation should check for the `error` key in JSON-RPC responses in addition to HTTP status codes.

6. **Connection State Events**: The ViewModel should subscribe to `IOdooService` connection state changes so that other pages (Dashboard, BOQ, PR) can reactively update their enabled state when connection drops or reconnects.

7. **Retry Logic**: The Python code notes "usually retrying a few times succeeds" for `ConnectionError`. The C# implementation should use Polly or a similar resilience library for configurable retry with exponential backoff, respecting the `RetryAttempts` pattern already established in the AutoCAD configuration section.

### Differences from Python

| Aspect | Python | C# |
|--------|--------|-----|
| HTTP Client | Bravado/Swagger | `System.Net.Http.HttpClient` |
| API Protocol | Custom Swagger endpoint + Basic Auth | Standard Odoo JSON-RPC + Session Auth |
| Configuration | YAML files (`server.yaml`, `token.yaml`) | `appsettings.json` with `IConfiguration` |
| Local Cache | SQLAlchemy + SQLite | EF Core or raw SQLite + `IMemoryCache` |
| Authentication | Token-based (UUID in `token.yaml`) | Session-based (username + password) |
| Product Filter | Hardcoded `categ_id child_of 27` | Configurable category via UI/config |
| Error Handling | Exception types (ConnectionError, JSONDecodeError, SwaggerValidationError) | `HttpRequestException`, JSON-RPC error objects, `TaskCanceledException` |
| Logging | Custom `log_util.safe_log_insert()` | `ILogger<T>` with structured logging |
| Concurrency | Synchronous calls on main thread | Async/await with `Task<T>` throughout |
