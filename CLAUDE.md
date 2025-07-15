# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Windows desktop application that bridges Odoo ERP and AutoCAD for engineering/construction workflows. The application enables automated extraction of drawing parameters from AutoCAD, manages BOQ (Bill of Quantities), and handles Purchase Requisition processing through Odoo's REST API.

## Architecture

### Core Technologies
- Python 3.10+ with Tkinter for GUI
- Windows COM (pywin32) for AutoCAD integration  
- Bravado/Swagger client for Odoo REST API communication
- SQLAlchemy ORM with SQLite for local data caching
- YAML configuration management for multiple environments

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

## Data Flow

1. **Configuration Setup**: Load YAML configs → Sync to SQLite → Establish Odoo/AutoCAD connections
2. **Parameter Extraction**: AutoCAD COM → Extract drawing parameters → Validate against Odoo products
3. **BOQ Processing**: Parameters → Generate BOQ entries → Push to Odoo project
4. **PR Generation**: BOQ data → Create Purchase Requisitions → Submit to Odoo purchasing workflow

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
```

### Building Application
```bash
# Build executable using optimized Python script
python build_exe.py

# Alternative: Build using Windows batch script
build_windows.bat

# Build installer with Inno Setup (requires Inno Setup installed)
iscc config/odoo-autocad-setup.iss

# Build installer with code signing (if certificates configured)
iscc config/odoo-autocad-setup-with-signing.iss
```

### Running and Testing
```bash
# Run main application
python odoo.py

# Test modern UI implementations
python tests/test_modern_ui.py
python tests/test_enhanced_ui.py
python tests/test_ui_cross_platform.py

# Debug launch with enhanced logging
python debug_launch.py
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
│   └── test_mcp_integration.py  # MCP server functionality
├── integration/             # Integration tests for workflows
│   ├── test_autocad_odoo_flow.py # End-to-end AutoCAD-Odoo workflow
│   ├── test_mcp_e2e.py          # MCP server end-to-end testing
│   └── test_gui_integration.py   # GUI integration scenarios
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
- **Antivirus**: Follow `ANTIVIRUS_SOLUTION.md` and `CODE_SIGNING_GUIDE.md` for deployment best practices
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

### 待完成任務
- 🔄 **當前**: 測試新UI (需要conda環境配置tkinter)
- ⏳ 階段二: 響應式佈局和導航優化
- ⏳ 階段三: 互動體驗提升

### 重新開始指導
1. 確保conda環境: `conda activate odoo_autocad && conda install tk`
2. 安裝測試依賴: `pip install pytest pytest-cov pytest-mock`
3. 測試UI: `python -m pytest tests/ui/`
4. 查看詳細計劃: `UI_IMPROVEMENT_PLAN.md`

### 新增的UI檔案
- `ui/` - UI主題和字體管理模組
- `forms/form_main_modern.py` - 現代化主表單
- `tests/` - UI測試檔案目錄
- `UI_IMPROVEMENT_PLAN.md` - 完整改善計劃

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
- **MCP整合**: 95%+（ai_assistant/）

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