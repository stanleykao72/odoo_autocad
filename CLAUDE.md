# CLAUDE.md

> **版本**: 5.0 (MCP整合完整版)  
> **最後更新**: 2025年7月16日  
> **新功能**: AI助手整合，7個MCP工具完整實作，完整TDD覆蓋

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Windows desktop application that bridges Odoo ERP and AutoCAD for engineering/construction workflows. The application enables automated extraction of drawing parameters from AutoCAD, manages BOQ (Bill of Quantities), and handles Purchase Requisition processing through Odoo's REST API.

**v5.0 新增功能**: 完整的MCP (Model Context Protocol) 整合，提供AI助手功能，支援自然語言操作AutoCAD和Odoo系統。包含7個完整實作的MCP工具，可透過Gemini CLI進行實際AutoCAD和Odoo系統操作。

## Architecture

### Core Technologies
- Python 3.10+ with Tkinter for GUI
- Windows COM (pywin32) for AutoCAD integration  
- Bravado/Swagger client for Odoo REST API communication
- SQLAlchemy ORM with SQLite for local data caching
- YAML configuration management for multiple environments
- **MCP (Model Context Protocol)** with FastMCP SDK for AI助手整合
- **SSE (Server-Sent Events)** 傳輸協定用於MCP通訊

### Key Components

**Entry Point**
- `odoo.py` - Main application entry point, initializes database and launches GUI

**GUI Layer** 
- `forms/form_main.py` - Main application window with side navigation
- `forms/form_autocad_param.py` - Parameter input dialogs for AutoCAD data

**Business Logic**
- `utility/util_odoo.py` - Odoo API client with authentication and data synchronization
- `utility/util_autocad.py` - AutoCAD COM interface for drawing parameter extraction
- `utility/util_push_to_boq.py` - BOQ processing and validation logic
- `utility/util_transfer_boq_to_pr.py` - Purchase Requisition generation from BOQ data

**Data Layer**
- `models/server.py` - SQLAlchemy models for configuration caching
- `db/database.db` - SQLite database for local data storage
- `config/*.yaml` - Environment-specific Odoo connection configurations

**Integration**
- `utility/util_com_server.py` - Python COM server for AutoCAD integration
- `cuix/odoo_autocad.cuix` - AutoCAD UI customization and ribbon interface

**MCP/AI Assistant**
- `mcp_server_fastmcp.py` - FastMCP SSE伺服器，提供7個MCP工具
- `utility/util_mcp_sse_manager.py` - SSE伺服器管理器，與GUI整合

## Data Flow

### 傳統工作流程
1. **Configuration Setup**: Load YAML configs → Sync to SQLite → Establish Odoo/AutoCAD connections
2. **Parameter Extraction**: AutoCAD COM → Extract drawing parameters → Validate against Odoo products
3. **BOQ Processing**: Parameters → Generate BOQ entries → Push to Odoo project
4. **PR Generation**: BOQ data → Create Purchase Requisitions → Submit to Odoo purchasing workflow

### MCP AI助手工作流程 (v5.0新增)
1. **SSE Server啟動**: GUI啟動 → MCPSSEManager → FastMCP SSE Server (port 8083)
2. **AI整合**: Gemini CLI連接 → MCP協定 → 自然語言指令
3. **工具執行**: MCP工具呼叫 → 實際AutoCAD/Odoo操作 → 回傳結果
4. **即時回饋**: 操作結果 → 透過SSE回傳 → AI助手顯示

## Development Commands

### Setup and Installation
```bash
# Install Python dependencies (建議使用conda環境)
conda activate odoo_autocad
pip install -r requirements.txt

# For Windows development, install Windows-specific dependencies
pip install -r requirements-windows.txt

# Install testing dependencies for TDD workflow
pip install pytest pytest-cov pytest-mock pytest-watch
pip install coverage[toml] pytest-html pytest-xdist

# Install MCP dependencies for AI assistant integration
pip install fastmcp mcp starlette uvicorn
```

### Building Application
```bash
# 一鍵完整建置流程 (推薦方式) - 從原始碼到安裝包
build_and_package.bat           # Windows批次檔版本
build_and_package.ps1          # PowerShell版本

# 手動建置選項
python -m PyInstaller --onefile --windowed --name "odoo-autocad-integration" --icon "icon/odoo_autocad.ico" --add-data "config;config" --add-data "fonts;fonts" --add-data "icon;icon" --add-data "db;db" --hidden-import "customtkinter" --hidden-import "win32com.client" --hidden-import "win32com.gen_py" --hidden-import "pywintypes" --hidden-import "win32api" --hidden-import "tkinter" --hidden-import "tkinter.ttk" --hidden-import "sqlalchemy" --hidden-import "sqlalchemy.ext.declarative" --hidden-import "sqlalchemy.orm" --exclude-module "pytest" --exclude-module "unittest" --exclude-module "doctest" --exclude-module "pdb" --exclude-module "matplotlib" --exclude-module "numpy" --exclude-module "pandas" --clean --noconfirm --distpath "output" odoo.py

# 手動安裝包建置 (需先建置EXE)
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer/odoo-autocad-setup.iss               # 基本安裝包
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer/odoo-autocad-setup-with-signing.iss  # 含簽章的安裝包
```

**注意**: 舊的 `build/build_windows.bat`, `build/build_windows.ps1`, `build/build_exe.py` 檔案已移除，統一使用根目錄的一鍵建置腳本。

#### 實際建置流程說明
更新後的建置系統包含：
1. **環境檢查**: Python, PyInstaller, Inno Setup
2. **EXE建置**: 直接使用 PyInstaller 指令，不依賴額外的建置腳本
3. **文檔複製**: 自動複製 README.md, ANTIVIRUS_SOLUTION.md 等文檔到輸出目錄
4. **程式碼簽章**: 可選，需要有效的憑證和密碼
5. **安裝包建置**: 使用 Inno Setup 生成最終安裝程式

#### 建置輸出
- **EXE檔案**: `output/odoo-autocad-integration.exe`
- **安裝包**: `installer/odoo-autocad-integration-5.0-setup.exe`
- **建置資訊**: `output/build_info.json` (包含檔案大小和SHA256雜湊值)
- **文檔檔案**: `output/*.md`

### Running and Testing
```bash
# Run main application
python odoo.py

# Run main application with MCP SSE server auto-start
python odoo.py --enable-mcp

# Test modern UI implementations
python tests/test_modern_ui.py
python tests/test_enhanced_ui.py
python tests/test_ui_cross_platform.py

# Debug launch with enhanced logging
python debug_launch.py

# Start MCP SSE server manually (for testing)
python mcp_server_fastmcp.py --mode sse --port 8083
```

### AutoCAD Integration
```bash
# Register COM server for AutoCAD integration
python utility/util_com_server.py

# Verify AutoCAD connection (requires AutoCAD running)
# Test through main application UI connection status
```

### Database Operations
```bash
# Initialize database (handled automatically by odoo.py)
# Database file location: db/database.db
# Manual SQLite access: sqlite3 db/database.db
```

### Process Management and Debugging
```bash
# Stop specific processes by PID using PowerShell
powershell "Stop-Process -Id 5412 -Force"

# Check port usage
netstat -ano | findstr ":808"

# Test port availability
powershell "Test-NetConnection -ComputerName localhost -Port 8083 -InformationLevel Quiet"

# Stop Python processes occupying specific ports
powershell "Get-Process python | Where-Object {$_.Id -eq 1234} | Stop-Process -Force"

# Find processes using specific ports (alternative method)
powershell "Get-NetTCPConnection | Where-Object LocalPort -eq 8083 | Select-Object OwningProcess"

# Clear ports 8080-8083 for SSE testing
netstat -ano | findstr ":808"  # Check current usage
powershell "Stop-Process -Id <PID> -Force"  # Replace <PID> with actual process ID
```

### MCP SSE Server Management
```bash
# Start GUI with auto-start SSE server (RECOMMENDED)
python odoo.py --enable-mcp

# Start SSE server manually for debugging
python mcp_server_fastmcp.py --mode sse --port 8083

# Test SSE server health
curl http://localhost:8083/health

# Clean up SSE ports before testing
# 1. Find processes using ports 8080-8083
netstat -ano | findstr ":808"
# 2. Stop specific processes
powershell "Stop-Process -Id <PID1> -Force"
powershell "Stop-Process -Id <PID2> -Force"
powershell "Stop-Process -Id <PID3> -Force"
# 3. Verify ports are free
powershell "Test-NetConnection -ComputerName localhost -Port 8083 -InformationLevel Quiet"

# Test MCP tools using Gemini CLI (需要先設定Gemini CLI)
# 參考 doc/GEMINI_CLI_SETUP.md 進行設定
```

### MCP工具說明 (v5.0新增)
本專案提供7個完整實作的MCP工具，透過AI助手可進行自然語言操作：

1. **test_connection()** - 測試MCP連接狀態
2. **get_server_info()** - 獲取伺服器資訊和功能
3. **check_autocad_status()** - 檢查AutoCAD連接狀態，返回應用程式資訊
4. **check_odoo_status()** - 檢查Odoo連接狀態，返回伺服器資訊
5. **extract_autocad_parameters(drawing_path, use_current_drawing)** - 從AutoCAD圖檔提取參數
6. **sync_to_odoo(data, sync_type)** - 同步資料到Odoo系統 (參數/BOQ/專案)
7. **generate_boq(project_id, include_autocad_data)** - 生成工程量清單，可包含AutoCAD資料

### Test-Driven Development (TDD)

This project follows TDD principles inspired by best practices. Always follow the TDD cycle: **Red → Green → Refactor**.

#### TDD Workflow
1. **Red**: Write a simple failing test that describes the expected behavior
2. **Green**: Implement the minimum code necessary to make the test pass
3. **Refactor**: Improve code structure while maintaining passing tests
4. **Commit**: Commit behavioral changes and structural changes separately

#### Test Organization
```
tests/
├── unit/                    # Unit tests for individual components
│   ├── test_autocad_utils.py    # AutoCAD utility functions
│   ├── test_odoo_api.py         # Odoo API integration
│   ├── test_boq_processing.py   # BOQ data processing logic
│   ├── test_mcp_integration.py  # MCP server functionality
│   └── test_direct_sse_manager.py # Direct SSE manager integration
├── integration/             # Integration tests for workflows
│   ├── test_autocad_odoo_flow.py # End-to-end AutoCAD-Odoo workflow
│   ├── test_mcp_e2e.py          # MCP server end-to-end testing
│   ├── test_gui_integration.py   # GUI integration scenarios
│   ├── test_simple_sse.py       # Simple SSE functionality
│   └── test_gui_sse.py          # GUI SSE integration
├── performance/            # Performance and benchmark tests
│   ├── test_autocad_performance.py # AutoCAD operation performance
│   └── test_odoo_sync_performance.py # Odoo synchronization performance
├── ui/                     # UI-specific tests
│   ├── test_modern_ui.py
│   ├── test_enhanced_ui.py
│   └── test_ui_cross_platform.py
└── fixtures/               # Test data and mocks
    ├── sample_drawings/     # Sample AutoCAD files
    ├── mock_odoo_data/     # Mock Odoo responses
    └── test_configs/       # Test configuration files
```

#### Testing Commands
```bash
# Run all tests with pytest
python -m pytest tests/

# Run specific test categories
python -m pytest tests/unit/          # Unit tests only
python -m pytest tests/integration/   # Integration tests only
python -m pytest tests/performance/   # Performance tests only
python -m pytest tests/ui/           # UI tests only

# Run SSE-specific tests
cd tests && python run_sse_tests.py  # All SSE tests with custom runner
python -m pytest tests/ -k "sse" -v  # SSE tests with pytest

# Test coverage reporting
python -m pytest --cov=. --cov-report=html tests/
python -m pytest --cov=utility --cov-report=term-missing tests/unit/

# Development mode - watch for changes
python -m pytest-watch -- tests/

# Run tests with specific markers
python -m pytest -m "not slow" tests/  # Skip slow tests
python -m pytest -m "autocad" tests/   # Run only AutoCAD-related tests
```

#### Development Guidelines

**When adding new features:**
1. Write a failing test that describes the desired behavior
2. Implement the simplest solution to make the test pass
3. Refactor to improve clarity and eliminate duplication
4. Ensure all tests pass before committing

**When fixing bugs:**
1. Write a test that reproduces the problem
2. Implement the fix to make the test pass
3. Verify all existing tests still pass

**Test Writing Best Practices:**
- Use descriptive test names that explain the behavior being tested
- Keep tests focused on a single behavior or requirement
- Use meaningful assertions with clear failure messages
- Leverage fixtures and mocking for external dependencies (AutoCAD, Odoo)

### Legacy Testing (UI)
```bash
# UI testing (manual testing approach)
python tests/ui/test_modern_ui.py        # Modern CustomTkinter interface
python tests/ui/test_enhanced_ui.py      # Enhanced widget functionality  
python tests/ui/test_ui_cross_platform.py # Cross-platform compatibility

# Run all UI tests
cd tests/ui && python test_modern_ui.py && python test_enhanced_ui.py && python test_ui_cross_platform.py
```

## Configuration Management

The application uses a dual configuration system:
- **YAML files** in `config/` for different environments (development, production)
- **SQLite database** for runtime configuration caching and user preferences

Key configuration areas:
- Odoo server connections (URL, database, credentials)
- AutoCAD integration settings
- Product mapping and validation rules
- BOQ processing parameters

## Important Patterns

### Test-Driven Development Patterns

**Test First, Code Second:**
- Write failing tests before implementing functionality
- Use tests to drive the design of your APIs
- Keep tests simple and focused on single behaviors

**Mocking External Dependencies:**
```python
# Mock AutoCAD COM objects for testing
@pytest.fixture
def mock_autocad():
    with patch('utility.util_autocad.win32com.client') as mock:
        yield mock.Dispatch.return_value

# Mock Odoo API responses
@pytest.fixture  
def mock_odoo_response():
    return {
        'products': [{'id': 1, 'name': 'Test Product'}],
        'status': 'success'
    }
```

**Test Data Factories:**
```python
# Create consistent test data
def create_test_boq_entry(name="Test Item", quantity=10, unit="pcs"):
    return {
        'name': name,
        'quantity': quantity, 
        'unit': unit,
        'created_at': datetime.now()
    }
```

### Error Handling
- Network operations use try-catch with user-friendly error dialogs
- COM operations include connection state validation
- Configuration loading has fallback mechanisms
- **All error scenarios should be covered by tests**

### Data Synchronization
- Odoo product data is cached locally and synchronized on demand
- Configuration changes update both YAML and SQLite stores
- BOQ data maintains consistency between AutoCAD and Odoo
- **Data consistency logic is validated through integration tests**

### COM Integration
- AutoCAD connection state is monitored continuously
- COM server provides bidirectional communication
- Drawing parameter extraction uses robust COM object handling
- **COM operations are tested with mock objects to ensure reliability**

### Code Quality Patterns

**Single Responsibility:**
- Each module handles one specific aspect of functionality
- Classes and functions have clear, focused purposes
- Business logic is separated from UI and data access

**Dependency Injection:**
- Pass dependencies (AutoCAD util, Odoo util) as parameters
- Makes testing easier and reduces coupling
- Enables mocking for unit tests

**Fail Fast:**
- Validate inputs early and provide clear error messages
- Use assertions in development to catch programming errors
- Return meaningful error codes and descriptions

## UI Architecture

The application uses a main window with side navigation pattern:
- Left sidebar for function selection
- Main content area for forms and data display
- Modal dialogs for parameter input and configuration
- Status indicators for connection states (Odoo, AutoCAD)

## Localization

The application includes Chinese language support throughout the UI and uses appropriate fonts (Microsoft JhengHei) for proper character rendering on Windows systems.

## Deployment and Security

### Code Signing
- **Certificates**: Store certificates in `certs/` directory
- **Configuration**: Use `config/odoo-autocad-setup-with-signing.iss` for signed installers
- **Antivirus**: Follow `doc/ANTIVIRUS_SOLUTION.md` and `doc/CODE_SIGNING_GUIDE.md` for deployment best practices
- **Build info**: Generated builds include SHA256 hashes and metadata in `output/build_info.json`

### File Structure for Deployment
```
output/
├── odoo-autocad-integration.exe     # Main executable
├── build_info.json                  # Build metadata and hashes
├── README_ANTIVIRUS.txt             # Antivirus whitelist instructions
└── *.md                            # Documentation files
```

### PyInstaller Configuration
- **Spec file**: `odoo-autocad-integration.spec` contains build configuration
- **Hidden imports**: Required for COM, CustomTkinter, and SQLAlchemy
- **Excluded modules**: Testing and heavy libraries excluded to reduce size
- **UPX compression**: Optional if UPX is available on build system

## UI Modernization Progress (進行中)

### 已完成階段一 (CustomTkinter基礎)
- ✅ CustomTkinter 框架安裝和配置
- ✅ 新的色彩方案和主題系統 (`ui/ui_theme.py`)
- ✅ 跨平台字體管理系統 (`ui/ui_fonts.py`) 
- ✅ 現代化主表單重構 (`forms/form_main_modern.py`)
- ✅ 改善的連接狀態指示器
- ✅ **MCP SSE整合** (`utility/util_mcp_sse_manager.py`)
- ✅ **AI助手功能完整實作** (`mcp_server_fastmcp.py`)

### 待完成任務
- 🔄 **當前**: 測試新UI (需要conda環境配置tkinter)
- ⏳ 階段二: 響應式佈局和導航優化
- ⏳ 階段三: 互動體驗提升

### 重新開始指導
1. 確保conda環境: `conda activate odoo_autocad && conda install tk`
2. 安裝測試依賴: `pip install pytest pytest-cov pytest-mock`
3. 測試UI: `python -m pytest tests/ui/`
4. 查看詳細計劃: `doc/UI_IMPROVEMENT_PLAN.md`

### 新增的UI檔案
- `ui/` - UI主題和字體管理模組
- `forms/form_main_modern.py` - 現代化主表單 (含MCP SSE整合)
- `tests/` - UI測試檔案目錄
- `doc/UI_IMPROVEMENT_PLAN.md` - 完整改善計劃

### v5.0 MCP整合檔案
- `mcp_server_fastmcp.py` - MCP SSE伺服器，7個完整實作工具
- `utility/util_mcp_sse_manager.py` - SSE伺服器管理器，GUI整合

---

## 🧪 TDD 開發提醒

當使用此專案進行開發時，請始終遵循TDD原則：

### 每日開發檢查清單
- [ ] 為新功能先寫失敗測試
- [ ] 實現最小代碼使測試通過
- [ ] 重構改善代碼品質
- [ ] 確保所有測試通過後再提交
- [ ] 為修復的bug新增回歸測試

### 測試覆蓋率目標
- **核心業務邏輯**: 95%+（util_autocad.py, util_odoo.py）
- **UI組件**: 80%+（forms/, ui/）
- **整合流程**: 90%+（端對端工作流程）
- **MCP整合**: 95%+（mcp_server_fastmcp.py, util_mcp_sse_manager.py）

### 快速測試指令
```bash
# 開發時持續測試
python -m pytest-watch -- tests/unit/

# 提交前完整測試
python -m pytest tests/ --cov=. --cov-report=term-missing

# 效能回歸測試
python -m pytest tests/performance/ -v
```

記住：**好的測試是最好的文檔，TDD讓重構變得安全且快速。**

---

## 📚 文檔結構

專案文檔統一存放在 `doc/` 目錄下：

```
doc/
├── MCP_INTEGRATION_PLAN.md      # MCP整合完整計劃與進度
├── DEPLOYMENT_GUIDE.md          # 部署和建置指南  
├── GEMINI_CLI_SETUP.md          # Gemini CLI配置說明
├── RELEASE_NOTES_v5.0.md        # v5.0版本發布說明
├── README-DEVELOPMENT.md        # 開發環境設定指南
├── UI_IMPROVEMENT_PLAN.md       # UI現代化計劃
├── ANTIVIRUS_SOLUTION.md        # 防毒軟體解決方案
├── CODE_SIGNING_GUIDE.md        # 代碼簽名指南
└── INTERNAL_CA_GUIDE.md         # 內部CA憑證指南
```

### 快速導航

- **🚀 開始使用**: 參考根目錄 `README.md`
- **🤖 AI助手設定**: `doc/GEMINI_CLI_SETUP.md`
- **🏗️ 建置部署**: `doc/DEPLOYMENT_GUIDE.md`
- **📋 完整計劃**: `doc/MCP_INTEGRATION_PLAN.md`
- **🔧 開發指引**: 本檔案 `CLAUDE.md`