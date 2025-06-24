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

# Build Windows installer (requires Inno Setup)
iscc config/odoo-autocad-setup.iss
```

### UI Development (現代化進行中)
```bash
# 測試新的現代化UI (需要tkinter支援)
python test_modern_ui.py

# 如果遇到tkinter問題，確保conda環境正確設置:
# conda install tk
```

### Database Management
```bash
# Initialize database (handled automatically by odoo.py)
python odoo.py
```

### Testing AutoCAD Integration
```bash
# Register COM server for AutoCAD integration
python utility/util_com_server.py
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
2. 測試UI: `python test_modern_ui.py`
3. 查看詳細計劃: `UI_IMPROVEMENT_PLAN.md`

### 新增的UI檔案
- `ui/` - UI主題和字體管理模組
- `forms/form_main_modern.py` - 現代化主表單
- `test_modern_ui.py` - UI測試檔案
- `UI_IMPROVEMENT_PLAN.md` - 完整改善計劃