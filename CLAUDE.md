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

### Testing
```bash
# UI testing (manual testing approach)
python tests/test_modern_ui.py        # Modern CustomTkinter interface
python tests/test_enhanced_ui.py      # Enhanced widget functionality  
python tests/test_ui_cross_platform.py # Cross-platform compatibility

# Run all UI tests
cd tests && python test_modern_ui.py && python test_enhanced_ui.py && python test_ui_cross_platform.py
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

### Error Handling
- Network operations use try-catch with user-friendly error dialogs
- COM operations include connection state validation
- Configuration loading has fallback mechanisms

### Data Synchronization
- Odoo product data is cached locally and synchronized on demand
- Configuration changes update both YAML and SQLite stores
- BOQ data maintains consistency between AutoCAD and Odoo

### COM Integration
- AutoCAD connection state is monitored continuously
- COM server provides bidirectional communication
- Drawing parameter extraction uses robust COM object handling

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
2. 測試UI: `python tests/test_modern_ui.py`
3. 查看詳細計劃: `UI_IMPROVEMENT_PLAN.md`

### 新增的UI檔案
- `ui/` - UI主題和字體管理模組
- `forms/form_main_modern.py` - 現代化主表單
- `tests/` - UI測試檔案目錄
- `UI_IMPROVEMENT_PLAN.md` - 完整改善計劃