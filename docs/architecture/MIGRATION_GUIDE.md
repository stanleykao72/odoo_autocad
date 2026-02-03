# Migration Guide: Python to C#

> **Version**: 1.0
> **Last Updated**: February 2026
> **Target**: .NET 8.0 with WPF

## Table of Contents

1. [Overview](#overview)
2. [File Mapping](#file-mapping)
3. [API Equivalents](#api-equivalents)
4. [Testing Strategy](#testing-strategy)
5. [Phase-by-Phase Migration Plan](#phase-by-phase-migration-plan)
6. [Risk Mitigation](#risk-mitigation)
7. [Rollback Strategy](#rollback-strategy)

---

## Overview

This guide provides a comprehensive roadmap for migrating the OdooAutoCAD application from Python (Tkinter) to C# (.NET 8/WPF). The migration preserves all existing functionality while improving performance, maintainability, and type safety.

### Migration Goals

| Goal | Description | Success Criteria |
|------|-------------|------------------|
| Feature Parity | All existing features work in C# version | 100% feature coverage |
| Performance | Faster startup and runtime | < 3s startup, < 100ms UI response |
| Maintainability | Better code organization | Full test coverage, clear architecture |
| Thread Safety | Solve COM threading issues | Zero COM threading errors |
| User Experience | Maintain familiar UI/UX | User acceptance testing pass |

### Migration Approach

We follow a **Strangler Fig Pattern**:
1. Build new C# components alongside Python
2. Gradually route functionality to C#
3. Test each component thoroughly
4. Remove Python code after validation

---

## File Mapping

### Python to C# Component Mapping

| Python File | C# Project | C# File(s) | Notes |
|-------------|------------|------------|-------|
| `odoo.py` | OdooAutoCAD.Desktop | `App.xaml.cs`, `Program.cs` | Application entry point |
| `forms/form_main.py` | OdooAutoCAD.Desktop | `MainWindow.xaml`, `MainViewModel.cs` | Main application window |
| `forms/form_main_modern.py` | OdooAutoCAD.Desktop | `MainWindow.xaml`, `MainViewModel.cs` | Modern UI (merged) |
| `forms/form_autocad_param.py` | OdooAutoCAD.Desktop | `Views/ParameterDialog.xaml` | Parameter input dialog |
| `utility/util_autocad.py` | OdooAutoCAD.AutoCAD | `AutoCADComService.cs`, `DrawingExtractor.cs` | AutoCAD COM integration |
| `utility/util_odoo.py` | OdooAutoCAD.Odoo | `OdooRestClient.cs`, `OdooAuthenticator.cs` | Odoo API client |
| `utility/util_push_to_boq.py` | OdooAutoCAD.Core | `Services/BOQService.cs` | BOQ processing logic |
| `utility/util_transfer_boq_to_pr.py` | OdooAutoCAD.Core | `Services/PurchaseRequisitionService.cs` | PR generation |
| `utility/util_com_server.py` | OdooAutoCAD.AutoCAD | `ComObjectFactory.cs` | COM server utilities |
| `utility/util_gui_proxy.py` | OdooAutoCAD.GuiProxy | `GuiProxyService.cs`, `MessageQueue.cs` | GUI proxy system |
| `utility/util_mcp_sse_manager.py` | OdooAutoCAD.MCP | `McpServerHost.cs`, `SseConnectionManager.cs` | MCP SSE management |
| `mcp_server_fastmcp.py` | OdooAutoCAD.MCP | `Tools/*.cs`, `JsonRpcHandler.cs` | MCP tools and protocol |
| `models/server.py` | OdooAutoCAD.Data | `Models/*.cs`, `ApplicationDbContext.cs` | Data models |
| `ui/ui_theme.py` | OdooAutoCAD.Desktop | `Resources/Themes/*.xaml` | UI theming |
| `ui/ui_fonts.py` | OdooAutoCAD.Desktop | `Resources/Fonts.xaml` | Font management |
| `config/*.yaml` | OdooAutoCAD.Desktop | `appsettings.json`, `IConfiguration` | Configuration |

### Directory Structure Mapping

```
Python Structure                    C# Structure
================                    ============

odoo_autocad_source/               OdooAutoCAD/
├── odoo.py                        ├── src/
├── forms/                         │   ├── OdooAutoCAD.Desktop/
│   ├── form_main.py              │   │   ├── App.xaml(.cs)
│   ├── form_main_modern.py       │   │   ├── MainWindow.xaml(.cs)
│   └── form_autocad_param.py     │   │   ├── ViewModels/
├── utility/                       │   │   ├── Views/
│   ├── util_autocad.py           │   │   └── Resources/
│   ├── util_odoo.py              │   ├── OdooAutoCAD.Core/
│   ├── util_push_to_boq.py       │   │   ├── Interfaces/
│   ├── util_transfer_boq_to_pr.py│   │   ├── Models/
│   ├── util_com_server.py        │   │   └── Services/
│   ├── util_gui_proxy.py         │   ├── OdooAutoCAD.AutoCAD/
│   └── util_mcp_sse_manager.py   │   ├── OdooAutoCAD.Odoo/
├── mcp_server_fastmcp.py         │   ├── OdooAutoCAD.GuiProxy/
├── models/                        │   ├── OdooAutoCAD.MCP/
│   └── server.py                 │   └── OdooAutoCAD.Data/
├── ui/                           ├── tests/
│   ├── ui_theme.py               │   ├── OdooAutoCAD.Core.Tests/
│   └── ui_fonts.py               │   ├── OdooAutoCAD.AutoCAD.Tests/
├── config/                        │   └── ...
│   └── *.yaml                    └── docs/
├── db/                                └── architecture/
│   └── database.db
└── tests/
    └── ...
```

---

## API Equivalents

### AutoCAD COM Operations

#### Python (pywin32)

```python
# util_autocad.py
import win32com.client

class AutoCADUtil:
    def __init__(self):
        self.app = None
        self.doc = None

    def connect(self):
        try:
            self.app = win32com.client.GetActiveObject("AutoCAD.Application")
            self.doc = self.app.ActiveDocument
            return True
        except:
            return False

    def get_drawing_path(self):
        if self.doc:
            return self.doc.FullName
        return None

    def extract_parameters(self):
        params = []
        if self.doc:
            model_space = self.doc.ModelSpace
            for i in range(model_space.Count):
                entity = model_space.Item(i)
                # Extract parameters...
        return params
```

#### C# Equivalent

```csharp
// AutoCADComService.cs
using System.Runtime.InteropServices;

namespace OdooAutoCAD.AutoCAD;

public class AutoCADComService : IAutoCADService, IDisposable
{
    private dynamic? _application;
    private dynamic? _document;
    private bool _disposed;

    public bool IsConnected => _application != null;

    public AutoCADInfo? ApplicationInfo
    {
        get
        {
            if (_application == null) return null;

            return new AutoCADInfo(
                Version: _application.Version,
                ProductName: _application.Name,
                CurrentDrawing: _document?.FullName,
                IsDocumentOpen: _document != null
            );
        }
    }

    public Task<bool> ConnectAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            // Get running AutoCAD instance
            _application = Marshal.GetActiveObject("AutoCAD.Application");
            _document = _application?.ActiveDocument;
            return Task.FromResult(_application != null);
        }
        catch (COMException)
        {
            return Task.FromResult(false);
        }
    }

    public Task DisconnectAsync()
    {
        ReleaseComObjects();
        return Task.CompletedTask;
    }

    public Task<string?> GetCurrentDrawingPathAsync()
    {
        return Task.FromResult(_document?.FullName as string);
    }

    public Task<IReadOnlyList<DrawingParameter>> ExtractParametersAsync(
        CancellationToken cancellationToken = default)
    {
        var parameters = new List<DrawingParameter>();

        if (_document == null)
            return Task.FromResult<IReadOnlyList<DrawingParameter>>(parameters);

        var modelSpace = _document.ModelSpace;
        int count = modelSpace.Count;

        for (int i = 0; i < count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            dynamic entity = modelSpace.Item(i);
            // Extract parameters based on entity type...

            // Release each entity after processing
            Marshal.ReleaseComObject(entity);
        }

        return Task.FromResult<IReadOnlyList<DrawingParameter>>(parameters);
    }

    private void ReleaseComObjects()
    {
        if (_document != null)
        {
            Marshal.ReleaseComObject(_document);
            _document = null;
        }

        if (_application != null)
        {
            Marshal.ReleaseComObject(_application);
            _application = null;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        ReleaseComObjects();
    }
}
```

### Odoo REST API Client

#### Python (Bravado/Requests)

```python
# util_odoo.py
import requests
from bravado.client import SwaggerClient

class OdooUtil:
    def __init__(self, config):
        self.base_url = config['url']
        self.db = config['database']
        self.session = requests.Session()
        self.uid = None

    def authenticate(self, username, password):
        url = f"{self.base_url}/web/session/authenticate"
        payload = {
            "jsonrpc": "2.0",
            "params": {
                "db": self.db,
                "login": username,
                "password": password
            }
        }
        response = self.session.post(url, json=payload)
        result = response.json()
        if result.get('result', {}).get('uid'):
            self.uid = result['result']['uid']
            return True
        return False

    def get_products(self):
        # API call implementation
        pass
```

#### C# Equivalent

```csharp
// OdooRestClient.cs
using System.Net.Http.Json;
using System.Text.Json;

namespace OdooAutoCAD.Odoo;

public class OdooRestClient : IOdooClient
{
    private readonly HttpClient _httpClient;
    private readonly OdooConfiguration _config;
    private readonly ILogger<OdooRestClient> _logger;

    private int? _userId;
    private string? _sessionId;

    public bool IsConnected => _userId.HasValue;

    public OdooServerInfo? ServerInfo { get; private set; }

    public OdooRestClient(
        HttpClient httpClient,
        IOptions<OdooConfiguration> config,
        ILogger<OdooRestClient> logger)
    {
        _httpClient = httpClient;
        _config = config.Value;
        _logger = logger;

        _httpClient.BaseAddress = new Uri(_config.ServerUrl);
    }

    public async Task<bool> AuthenticateAsync(CancellationToken cancellationToken = default)
    {
        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                db = _config.Database,
                login = _config.Username,
                password = _config.Password
            },
            id = 1
        };

        try
        {
            var response = await _httpClient.PostAsJsonAsync(
                "/web/session/authenticate",
                request,
                cancellationToken);

            response.EnsureSuccessStatusCode();

            var result = await response.Content
                .ReadFromJsonAsync<JsonRpcResponse<AuthResult>>(cancellationToken);

            if (result?.Result?.UserId > 0)
            {
                _userId = result.Result.UserId;
                _sessionId = result.Result.SessionId;

                ServerInfo = new OdooServerInfo(
                    _config.ServerUrl,
                    _config.Database,
                    result.Result.ServerVersion ?? "Unknown",
                    _userId.Value
                );

                _logger.LogInformation(
                    "Authenticated to Odoo as user {UserId}", _userId);
                return true;
            }

            _logger.LogWarning("Authentication failed: No user ID returned");
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to authenticate with Odoo");
            return false;
        }
    }

    public async Task<IReadOnlyList<OdooProduct>> GetProductsAsync(
        CancellationToken cancellationToken = default)
    {
        EnsureAuthenticated();

        var request = new
        {
            jsonrpc = "2.0",
            method = "call",
            @params = new
            {
                model = "product.product",
                method = "search_read",
                args = new object[] { },
                kwargs = new
                {
                    fields = new[] { "id", "name", "default_code", "uom_id", "list_price" },
                    limit = 1000
                }
            },
            id = 2
        };

        var response = await _httpClient.PostAsJsonAsync(
            "/web/dataset/call_kw",
            request,
            cancellationToken);

        response.EnsureSuccessStatusCode();

        var result = await response.Content
            .ReadFromJsonAsync<JsonRpcResponse<List<OdooProductDto>>>(cancellationToken);

        return result?.Result?
            .Select(p => new OdooProduct(p.Id, p.Name, p.DefaultCode, p.UomId, p.ListPrice))
            .ToList()
            ?? new List<OdooProduct>();
    }

    private void EnsureAuthenticated()
    {
        if (!IsConnected)
            throw new OdooApiException("Not authenticated with Odoo");
    }
}

// DTOs
internal record JsonRpcResponse<T>(T? Result, JsonRpcError? Error);
internal record AuthResult(int UserId, string? SessionId, string? ServerVersion);
internal record OdooProductDto(int Id, string Name, string? DefaultCode, int[] UomId, decimal ListPrice);
```

### GUI Proxy System

#### Python

```python
# util_gui_proxy.py
import queue
import threading
from dataclasses import dataclass
from typing import Callable, Any

@dataclass
class ProxyRequest:
    request_id: str
    operation: Callable
    result_event: threading.Event
    result: Any = None
    error: Exception = None

class GuiProxy:
    def __init__(self, root):
        self.root = root
        self.request_queue = queue.Queue()
        self._running = False

    def start(self):
        self._running = True
        self._poll()

    def _poll(self):
        if not self._running:
            return

        try:
            while True:
                request = self.request_queue.get_nowait()
                self._process_request(request)
        except queue.Empty:
            pass

        # Schedule next poll
        self.root.after(50, self._poll)

    def _process_request(self, request):
        try:
            request.result = request.operation()
        except Exception as e:
            request.error = e
        finally:
            request.result_event.set()

    def execute(self, operation, timeout=30):
        request = ProxyRequest(
            request_id=str(uuid.uuid4()),
            operation=operation,
            result_event=threading.Event()
        )

        self.request_queue.put(request)

        if request.result_event.wait(timeout):
            if request.error:
                raise request.error
            return request.result
        else:
            raise TimeoutError("Operation timed out")
```

#### C# Equivalent

```csharp
// GuiProxyService.cs - See THREADING_MODEL.md for full implementation
// Key differences from Python:

// 1. Uses Channel<T> instead of Queue
private readonly Channel<ProxyRequest> _requestChannel;

// 2. Uses TaskCompletionSource instead of Event
public sealed class ProxyRequest
{
    public Guid RequestId { get; } = Guid.NewGuid();
    public Func<object?> Operation { get; init; } = null!;
    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(30);
    internal TaskCompletionSource<ProxyResponse> ResponseTcs { get; } = new();
}

// 3. Uses DispatcherTimer instead of root.after()
_processingTimer = new DispatcherTimer(
    _options.PollingInterval,
    DispatcherPriority.Normal,
    ProcessPendingRequests,
    _dispatcher);

// 4. Returns Task instead of blocking
public async Task<T> ExecuteAsync<T>(
    Func<T> func,
    TimeSpan? timeout = null,
    CancellationToken cancellationToken = default)
{
    var request = new ProxyRequest
    {
        Operation = () => func()!,
        Timeout = timeout ?? _options.DefaultTimeout,
        CancellationToken = cancellationToken
    };

    var response = await _messageQueue.EnqueueAsync(request, cancellationToken);

    if (!response.IsSuccess)
        throw response.Exception ?? new InvalidOperationException("Unknown error");

    return (T)response.Result!;
}
```

### Configuration Management

#### Python (YAML)

```python
# config loading
import yaml

def load_config(env='development'):
    with open(f'config/{env}.yaml', 'r') as f:
        return yaml.safe_load(f)

config = load_config()
odoo_url = config['odoo']['url']
```

```yaml
# config/development.yaml
odoo:
  url: https://odoo.example.com
  database: production_db
  username: api_user
```

#### C# Equivalent (appsettings.json + IConfiguration)

```csharp
// appsettings.json
{
    "Odoo": {
        "ServerUrl": "https://odoo.example.com",
        "Database": "production_db",
        "Username": "api_user",
        "Password": ""  // Set via environment or user secrets
    },
    "AutoCAD": {
        "ConnectionTimeout": 30,
        "ExtractTimeout": 120
    },
    "GuiProxy": {
        "DefaultTimeout": 30,
        "PollingInterval": 50,
        "MaxPendingRequests": 100
    },
    "Mcp": {
        "Port": 8084,
        "EnableAutoStart": false
    }
}

// Configuration classes
public class OdooConfiguration
{
    public string ServerUrl { get; set; } = string.Empty;
    public string Database { get; set; } = string.Empty;
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
}

// Registration
services.Configure<OdooConfiguration>(
    configuration.GetSection("Odoo"));

// Usage
public class OdooRestClient
{
    private readonly OdooConfiguration _config;

    public OdooRestClient(IOptions<OdooConfiguration> config)
    {
        _config = config.Value;
    }
}
```

### Database Operations

#### Python (SQLAlchemy)

```python
# models/server.py
from sqlalchemy import Column, Integer, String, create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

Base = declarative_base()

class ServerConfig(Base):
    __tablename__ = 'server_config'

    id = Column(Integer, primary_key=True)
    name = Column(String)
    url = Column(String)
    database = Column(String)

engine = create_engine('sqlite:///db/database.db')
Session = sessionmaker(bind=engine)

def get_configs():
    session = Session()
    try:
        return session.query(ServerConfig).all()
    finally:
        session.close()
```

#### C# Equivalent (EF Core)

```csharp
// Models/ServerConfiguration.cs
namespace OdooAutoCAD.Data.Models;

public class ServerConfiguration
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
    public string Database { get; set; } = string.Empty;
    public DateTime CreatedAt { get; set; }
    public DateTime? UpdatedAt { get; set; }
}

// ApplicationDbContext.cs
namespace OdooAutoCAD.Data;

public class ApplicationDbContext : DbContext
{
    public DbSet<ServerConfiguration> ServerConfigurations => Set<ServerConfiguration>();
    public DbSet<ProductCache> ProductCaches => Set<ProductCache>();
    public DbSet<UserPreference> UserPreferences => Set<UserPreference>();

    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options)
    {
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<ServerConfiguration>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.Name).IsRequired().HasMaxLength(100);
            entity.Property(e => e.Url).IsRequired().HasMaxLength(500);
        });
    }
}

// Repository pattern
public interface IServerConfigRepository
{
    Task<IReadOnlyList<ServerConfiguration>> GetAllAsync(CancellationToken ct = default);
    Task<ServerConfiguration?> GetByIdAsync(int id, CancellationToken ct = default);
    Task<ServerConfiguration> AddAsync(ServerConfiguration config, CancellationToken ct = default);
    Task UpdateAsync(ServerConfiguration config, CancellationToken ct = default);
    Task DeleteAsync(int id, CancellationToken ct = default);
}

public class ServerConfigRepository : IServerConfigRepository
{
    private readonly ApplicationDbContext _context;

    public ServerConfigRepository(ApplicationDbContext context)
    {
        _context = context;
    }

    public async Task<IReadOnlyList<ServerConfiguration>> GetAllAsync(
        CancellationToken ct = default)
    {
        return await _context.ServerConfigurations
            .OrderBy(c => c.Name)
            .ToListAsync(ct);
    }

    // ... other methods
}
```

---

## Testing Strategy

### Test Categories

| Category | Python Framework | C# Framework | Coverage Target |
|----------|------------------|--------------|-----------------|
| Unit Tests | pytest | xUnit | 90% |
| Integration Tests | pytest | xUnit + WebApplicationFactory | 80% |
| UI Tests | Manual | xUnit + FlaUI | 60% |
| Performance Tests | pytest-benchmark | BenchmarkDotNet | Key paths |

### Test Migration Approach

1. **Create test specifications** from Python tests
2. **Write C# tests first** (TDD approach)
3. **Implement C# code** to pass tests
4. **Run both test suites** during transition
5. **Deprecate Python tests** after C# validation

### Example Test Migration

#### Python Test

```python
# tests/unit/test_autocad_utils.py
import pytest
from unittest.mock import Mock, patch

class TestAutoCADUtil:
    @pytest.fixture
    def mock_com(self):
        with patch('utility.util_autocad.win32com.client') as mock:
            yield mock

    def test_connect_success(self, mock_com):
        mock_app = Mock()
        mock_com.GetActiveObject.return_value = mock_app

        util = AutoCADUtil()
        result = util.connect()

        assert result is True
        assert util.app == mock_app

    def test_connect_failure(self, mock_com):
        mock_com.GetActiveObject.side_effect = Exception("Not running")

        util = AutoCADUtil()
        result = util.connect()

        assert result is False
```

#### C# Equivalent

```csharp
// OdooAutoCAD.AutoCAD.Tests/AutoCADComServiceTests.cs
using Moq;
using Xunit;
using FluentAssertions;

namespace OdooAutoCAD.AutoCAD.Tests;

public class AutoCADComServiceTests
{
    [Fact]
    public async Task ConnectAsync_WhenAutoCADRunning_ReturnsTrue()
    {
        // Arrange
        var service = new TestableAutoCADComService(autoCADRunning: true);

        // Act
        var result = await service.ConnectAsync();

        // Assert
        result.Should().BeTrue();
        service.IsConnected.Should().BeTrue();
    }

    [Fact]
    public async Task ConnectAsync_WhenAutoCADNotRunning_ReturnsFalse()
    {
        // Arrange
        var service = new TestableAutoCADComService(autoCADRunning: false);

        // Act
        var result = await service.ConnectAsync();

        // Assert
        result.Should().BeFalse();
        service.IsConnected.Should().BeFalse();
    }

    [Fact]
    public async Task ExtractParametersAsync_WhenConnected_ReturnsParameters()
    {
        // Arrange
        var service = new TestableAutoCADComService(autoCADRunning: true);
        await service.ConnectAsync();

        // Act
        var parameters = await service.ExtractParametersAsync();

        // Assert
        parameters.Should().NotBeNull();
        parameters.Should().HaveCountGreaterThan(0);
    }

    [Fact]
    public async Task ExtractParametersAsync_WhenNotConnected_ReturnsEmpty()
    {
        // Arrange
        var service = new TestableAutoCADComService(autoCADRunning: false);

        // Act
        var parameters = await service.ExtractParametersAsync();

        // Assert
        parameters.Should().BeEmpty();
    }
}

// Test double for COM isolation
internal class TestableAutoCADComService : AutoCADComService
{
    private readonly bool _autoCADRunning;
    private bool _connected;

    public TestableAutoCADComService(bool autoCADRunning)
    {
        _autoCADRunning = autoCADRunning;
    }

    public override Task<bool> ConnectAsync(CancellationToken ct = default)
    {
        _connected = _autoCADRunning;
        return Task.FromResult(_connected);
    }

    public override bool IsConnected => _connected;
}
```

---

## Phase-by-Phase Migration Plan

### Phase 1: Foundation (Weeks 1-2)

**Objective**: Set up C# project structure and core infrastructure

**Tasks**:
- [ ] Create Visual Studio solution with all projects
- [ ] Configure NuGet packages and dependencies
- [ ] Set up logging infrastructure (Serilog)
- [ ] Implement configuration management
- [ ] Set up CI/CD pipeline for C# project
- [ ] Create database context and initial migrations

**Deliverables**:
- Compiling solution with basic structure
- Working unit test framework
- Configuration loading from appsettings.json

**Validation**:
- All projects compile successfully
- Unit test framework runs
- Configuration loads correctly

---

### Phase 2: Data Layer (Weeks 3-4)

**Objective**: Migrate database models and repository layer

**Tasks**:
- [ ] Create EF Core entity models
- [ ] Implement repository interfaces
- [ ] Write repository implementations
- [ ] Create database migrations
- [ ] Write unit tests for repositories
- [ ] Create data migration script from SQLite

**Deliverables**:
- Complete data layer implementation
- 95% test coverage on repositories
- Data migration script

**Validation**:
- All repository tests pass
- Data imports correctly from Python database
- CRUD operations work correctly

---

### Phase 3: GUI Proxy System (Weeks 5-6)

**Objective**: Implement thread-safe GUI proxy for COM operations

**Tasks**:
- [ ] Implement MessageQueue with Channel<T>
- [ ] Implement ProxyRequest/ProxyResponse types
- [ ] Implement GuiProxyService
- [ ] Add timeout handling
- [ ] Add error handling and logging
- [ ] Write comprehensive unit tests
- [ ] Write integration tests with mock COM

**Deliverables**:
- Complete GUI proxy implementation
- 95% test coverage
- Performance benchmarks

**Validation**:
- All tests pass
- No deadlocks under load
- Correct timeout behavior

---

### Phase 4: AutoCAD Integration (Weeks 7-9)

**Objective**: Migrate AutoCAD COM integration

**Tasks**:
- [ ] Implement IAutoCADService interface
- [ ] Implement AutoCADComService
- [ ] Implement DrawingExtractor
- [ ] Integrate with GUI Proxy
- [ ] Write unit tests with COM mocking
- [ ] Write integration tests with real AutoCAD
- [ ] Performance testing and optimization

**Deliverables**:
- Complete AutoCAD integration
- 90% test coverage
- Integration test suite

**Validation**:
- Connects to running AutoCAD
- Extracts parameters correctly
- No COM threading errors
- Performance meets targets

---

### Phase 5: Odoo Integration (Weeks 10-11)

**Objective**: Migrate Odoo REST API client

**Tasks**:
- [ ] Implement IOdooClient interface
- [ ] Implement OdooRestClient with HttpClient
- [ ] Implement authentication flow
- [ ] Implement product sync
- [ ] Implement BOQ operations
- [ ] Implement PR operations
- [ ] Add retry policies with Polly
- [ ] Write unit and integration tests

**Deliverables**:
- Complete Odoo client implementation
- 90% test coverage
- Integration tests against test Odoo instance

**Validation**:
- Authentication works
- Product sync works
- BOQ and PR operations work
- Error handling is robust

---

### Phase 6: MCP Server (Weeks 12-13)

**Objective**: Migrate MCP SSE server and tools

**Tasks**:
- [ ] Implement SSE server with ASP.NET Core
- [ ] Implement JSON-RPC handler
- [ ] Implement all 7 MCP tools
- [ ] Integrate with GUI Proxy for COM operations
- [ ] Write unit tests for each tool
- [ ] Write integration tests with Gemini CLI

**Deliverables**:
- Complete MCP server implementation
- All 7 tools working
- Integration with Gemini CLI validated

**Validation**:
- SSE connection establishes
- All tools respond correctly
- No threading issues with AutoCAD tools

---

### Phase 7: WPF UI (Weeks 14-17)

**Objective**: Build WPF user interface

**Tasks**:
- [ ] Create main window layout
- [ ] Implement MVVM ViewModels
- [ ] Create navigation system
- [ ] Implement connection status indicators
- [ ] Implement parameter extraction view
- [ ] Implement BOQ management view
- [ ] Implement settings view
- [ ] Create custom themes and styles
- [ ] Write UI automation tests

**Deliverables**:
- Complete WPF application
- Feature parity with Python UI
- UI test coverage

**Validation**:
- All features accessible
- UI is responsive
- Matches design specifications

---

### Phase 8: Integration and Testing (Weeks 18-19)

**Objective**: Full system integration and testing

**Tasks**:
- [ ] End-to-end testing of all workflows
- [ ] Performance testing and optimization
- [ ] Security review
- [ ] User acceptance testing
- [ ] Documentation updates
- [ ] Bug fixes from testing

**Deliverables**:
- Fully integrated application
- Test reports
- Updated documentation

**Validation**:
- All E2E tests pass
- Performance meets targets
- UAT sign-off

---

### Phase 9: Deployment and Transition (Weeks 20-21)

**Objective**: Deploy C# version and deprecate Python

**Tasks**:
- [ ] Create installer package
- [ ] Write deployment documentation
- [ ] Create user migration guide
- [ ] Parallel deployment (both versions)
- [ ] Monitor for issues
- [ ] Deprecate Python version

**Deliverables**:
- Installer package
- Deployment documentation
- User migration guide

**Validation**:
- Successful installations
- No critical issues in production
- Python version deprecated

---

## Risk Mitigation

### Identified Risks

| Risk | Impact | Probability | Mitigation |
|------|--------|-------------|------------|
| COM compatibility issues | High | Medium | Early prototyping, extensive testing |
| Performance regression | Medium | Low | Benchmarking at each phase |
| Feature gaps | High | Low | Comprehensive feature mapping |
| Learning curve | Medium | Medium | Team training, code reviews |
| Timeline slippage | Medium | Medium | Buffer time, phased approach |

### Contingency Plans

1. **COM Issues**: Fall back to Python COM wrapper with C# interop
2. **Performance**: Identify bottlenecks early, optimize critical paths
3. **Timeline**: Prioritize core features, defer enhancements
4. **Resources**: Cross-train team members, document thoroughly

---

## Rollback Strategy

### Rollback Triggers

- Critical bugs affecting core functionality
- Performance degradation > 50%
- Data corruption or loss
- Security vulnerabilities

### Rollback Procedure

1. **Immediate**: Switch users back to Python version
2. **Communication**: Notify stakeholders of rollback
3. **Analysis**: Root cause analysis of issues
4. **Planning**: Revise timeline and approach
5. **Resolution**: Fix issues before re-attempting migration

### Data Preservation

- Both versions use compatible SQLite database
- No schema changes during transition phase
- Regular database backups during migration

---

## Success Metrics

| Metric | Target | Measurement |
|--------|--------|-------------|
| Feature Parity | 100% | Feature checklist |
| Test Coverage | > 85% | Code coverage tools |
| Startup Time | < 3 seconds | Automated timing |
| UI Response | < 100ms | Performance tests |
| COM Errors | 0 | Error monitoring |
| User Satisfaction | > 4/5 | User surveys |

---

*Document Version: 1.0 | Created: February 2026*
