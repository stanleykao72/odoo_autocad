# PRP: export_to_database

> **功能**: 將 AutoCAD 圖面元素匯出到 SQLite 資料庫  
> **類型**: MCP 工具  
> **優先級**: 高  
> **日期**: 2025年7月16日  
> **階段**: 第二階段進階繪圖工具

## 功能概述

### 目標
將掃描的 AutoCAD 圖面元素資料匯出到 SQLite 資料庫中，支援增量更新、幾何資料儲存和與 Odoo 系統的同步。這是第二階段進階工具的重要功能，為後續的 BOQ 生成和專案管理提供資料基礎。

### 使用場景
1. **資料持久化**: 將 AutoCAD 圖面資料永久保存到資料庫
2. **增量更新**: 支援圖面修改後的資料更新
3. **BOQ 準備**: 為工程量清單生成準備結構化資料
4. **系統整合**: 與 Odoo 系統進行資料同步
5. **AI 助手整合**: 透過自然語言查詢和匯出圖面資料

### 成功標準
- [ ] 能夠將掃描的元素資料儲存到 SQLite 資料庫
- [ ] 支援增量更新和版本控制
- [ ] 提供幾何資料的完整儲存和檢索
- [ ] 支援與 Odoo 系統的資料同步
- [ ] 提供匯出進度和統計資訊
- [ ] 與現有資料庫架構完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def export_to_database(
    drawing_name: str,
    include_geometry: bool = True,
    incremental_update: bool = True,
    sync_to_odoo: bool = False,
    element_types: list = None,
    layer_filter: str = None
) -> Dict[str, Any]:
    """將 AutoCAD 圖面元素匯出到 SQLite 資料庫"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| drawing_name | str | 是 | - | 圖面名稱，用作資料庫識別 |
| include_geometry | bool | 否 | True | 是否包含幾何資料 |
| incremental_update | bool | 否 | True | 是否進行增量更新 |
| sync_to_odoo | bool | 否 | False | 是否同步到 Odoo 系統 |
| element_types | list | 否 | None | 要匯出的元素類型列表 |
| layer_filter | str | 否 | None | 圖層過濾器 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "drawing_name": "建築平面圖",
        "database_path": "db/database.db",
        "export_summary": {
            "total_elements": 157,
            "new_elements": 23,
            "updated_elements": 8,
            "unchanged_elements": 126,
            "element_breakdown": {
                "line": 89,
                "circle": 34,
                "arc": 12,
                "text": 15,
                "dimension": 7
            }
        },
        "database_info": {
            "table_name": "autocad_elements",
            "version": "5.0.1",
            "last_updated": "2025-07-16T10:30:00"
        },
        "sync_info": {
            "odoo_sync": False,
            "sync_status": "not_requested",
            "last_sync": None
        },
        "export_settings": {
            "include_geometry": True,
            "incremental_update": True,
            "element_types": ["all"],
            "layer_filter": None
        },
        "performance": {
            "scan_time": 2.34,
            "export_time": 1.28,
            "total_time": 3.62
        },
        "exported_at": "2025-07-16T10:30:00"
    },
    "message": "成功匯出 157 個元素到資料庫 (新增: 23, 更新: 8)",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "圖面名稱不能為空",
    "error_code": "INVALID_DRAWING_NAME",
    "suggestion": "請提供有效的圖面名稱"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [x] 需要 SQLite 資料庫（使用現有的資料庫基礎設施）
- [x] 需要 `scan_elements` 功能（已實現）
- [ ] 需要 Odoo 連接（用於同步，使用現有的 `_odoo_util`）
- [ ] 需要新的資料庫匯出工具類別

### 核心邏輯
1. **參數驗證**: 檢查圖面名稱、元素類型列表格式
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **資料庫連接**: 建立或確認 SQLite 資料庫連接
4. **元素掃描**: 使用 `scan_elements` 獲取圖面元素資料
5. **增量處理**: 比較現有資料庫記錄，識別新增、更新、刪除
6. **資料儲存**: 將元素資料儲存到適當的資料庫表格
7. **Odoo 同步**: 如果需要，將資料同步到 Odoo 系統
8. **統計生成**: 生成匯出統計和效能資訊
9. **結果回傳**: 返回完整的匯出結果和統計資訊

### 錯誤處理
- **參數錯誤**: 圖面名稱無效、元素類型列表格式錯誤
- **連接錯誤**: AutoCAD 未運行、資料庫連接失敗、Odoo 連接問題
- **掃描錯誤**: 元素掃描失敗、資料格式錯誤
- **資料庫錯誤**: 寫入失敗、表格不存在、資料衝突
- **同步錯誤**: Odoo 同步失敗、網路問題、認證錯誤

## 資料庫架構

### 主要表格: `autocad_elements`
```sql
CREATE TABLE IF NOT EXISTS autocad_elements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    drawing_name TEXT NOT NULL,
    element_id TEXT NOT NULL,
    element_type TEXT NOT NULL,
    layer_name TEXT,
    color INTEGER,
    geometry_data TEXT,  -- JSON 格式儲存幾何資料
    properties_data TEXT,  -- JSON 格式儲存屬性資料
    bounds_data TEXT,  -- JSON 格式儲存邊界資料
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    version INTEGER DEFAULT 1,
    odoo_sync_status TEXT DEFAULT 'pending',
    odoo_sync_at TIMESTAMP NULL,
    UNIQUE(drawing_name, element_id)
);
```

### 索引設計
```sql
CREATE INDEX IF NOT EXISTS idx_autocad_elements_drawing ON autocad_elements(drawing_name);
CREATE INDEX IF NOT EXISTS idx_autocad_elements_type ON autocad_elements(element_type);
CREATE INDEX IF NOT EXISTS idx_autocad_elements_layer ON autocad_elements(layer_name);
CREATE INDEX IF NOT EXISTS idx_autocad_elements_sync ON autocad_elements(odoo_sync_status);
```

### 版本控制表: `autocad_drawing_versions`
```sql
CREATE TABLE IF NOT EXISTS autocad_drawing_versions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    drawing_name TEXT NOT NULL,
    version INTEGER NOT NULL,
    element_count INTEGER,
    export_timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    hash_value TEXT,  -- 資料完整性檢查
    notes TEXT,
    UNIQUE(drawing_name, version)
);
```

## 範例代碼

### 基本實現
```python
@mcp.tool()
def export_to_database(
    drawing_name: str,
    include_geometry: bool = True,
    incremental_update: bool = True,
    sync_to_odoo: bool = False,
    element_types: list = None,
    layer_filter: str = None
) -> Dict[str, Any]:
    """將 AutoCAD 圖面元素匯出到 SQLite 資料庫"""
    logger.info(f"export_to_database called with params: {locals()}")
    
    try:
        # 參數驗證
        if not drawing_name or not drawing_name.strip():
            return {
                "status": "error",
                "message": "圖面名稱不能為空",
                "error_code": "INVALID_DRAWING_NAME",
                "suggestion": "請提供有效的圖面名稱"
            }
        
        drawing_name = drawing_name.strip()
        
        # 驗證元素類型列表
        if element_types is not None:
            if not isinstance(element_types, list):
                return {
                    "status": "error",
                    "message": "元素類型必須是列表格式",
                    "error_code": "INVALID_ELEMENT_TYPES",
                    "suggestion": "請提供有效的元素類型列表，例如: ['line', 'circle']"
                }
            
            valid_types = ["line", "circle", "arc", "text", "dimension", "block"]
            invalid_types = [t for t in element_types if t not in valid_types]
            if invalid_types:
                return {
                    "status": "error",
                    "message": f"無效的元素類型: {invalid_types}",
                    "error_code": "INVALID_ELEMENT_TYPES",
                    "suggestion": f"有效的元素類型: {valid_types}"
                }
        
        # 檢查 AutoCAD 連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 開始計時
        start_time = datetime.now()
        
        # 步驟1: 掃描元素
        scan_start = datetime.now()
        if element_types:
            # 為每個元素類型分別掃描
            all_elements = []
            for element_type in element_types:
                scan_result = autocad_util.scan_elements(
                    element_type=element_type,
                    include_geometry=include_geometry,
                    include_properties=True,
                    layer_filter=layer_filter
                )
                all_elements.extend(scan_result.get("elements", []))
        else:
            # 掃描所有元素
            scan_result = autocad_util.scan_elements(
                element_type="all",
                include_geometry=include_geometry,
                include_properties=True,
                layer_filter=layer_filter
            )
            all_elements = scan_result.get("elements", [])
        
        scan_time = (datetime.now() - scan_start).total_seconds()
        
        # 步驟2: 資料庫匯出
        export_start = datetime.now()
        
        # 使用資料庫匯出工具
        db_util = _get_database_util()
        export_result = db_util.export_elements_to_database(
            drawing_name=drawing_name,
            elements=all_elements,
            incremental_update=incremental_update,
            include_geometry=include_geometry
        )
        
        export_time = (datetime.now() - export_start).total_seconds()
        total_time = (datetime.now() - start_time).total_seconds()
        
        # 步驟3: Odoo 同步（如果需要）
        sync_info = {
            "odoo_sync": sync_to_odoo,
            "sync_status": "not_requested",
            "last_sync": None
        }
        
        if sync_to_odoo:
            try:
                odoo_util = _odoo_util
                if odoo_util:
                    sync_result = odoo_util.sync_drawing_elements(
                        drawing_name=drawing_name,
                        elements=all_elements
                    )
                    sync_info.update({
                        "sync_status": "success" if sync_result.get("success") else "failed",
                        "last_sync": datetime.now().isoformat(),
                        "sync_details": sync_result
                    })
                else:
                    sync_info.update({
                        "sync_status": "failed",
                        "error": "Odoo 連接未建立"
                    })
            except Exception as sync_error:
                sync_info.update({
                    "sync_status": "failed",
                    "error": str(sync_error)
                })
        
        # 統計計算
        element_breakdown = {}
        for element in all_elements:
            element_type = element.get("type", "unknown")
            element_breakdown[element_type] = element_breakdown.get(element_type, 0) + 1
        
        return {
            "status": "success",
            "data": {
                "drawing_name": drawing_name,
                "database_path": export_result.get("database_path"),
                "export_summary": {
                    "total_elements": len(all_elements),
                    "new_elements": export_result.get("new_count", 0),
                    "updated_elements": export_result.get("updated_count", 0),
                    "unchanged_elements": export_result.get("unchanged_count", 0),
                    "element_breakdown": element_breakdown
                },
                "database_info": {
                    "table_name": "autocad_elements",
                    "version": export_result.get("version"),
                    "last_updated": export_result.get("last_updated")
                },
                "sync_info": sync_info,
                "export_settings": {
                    "include_geometry": include_geometry,
                    "incremental_update": incremental_update,
                    "element_types": element_types or ["all"],
                    "layer_filter": layer_filter
                },
                "performance": {
                    "scan_time": scan_time,
                    "export_time": export_time,
                    "total_time": total_time
                },
                "exported_at": datetime.now().isoformat()
            },
            "message": f"成功匯出 {len(all_elements)} 個元素到資料庫 (新增: {export_result.get('new_count', 0)}, 更新: {export_result.get('updated_count', 0)})",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in export_to_database: {e}")
        return {
            "status": "error",
            "message": f"匯出到資料庫失敗: {str(e)}",
            "error_code": "DATABASE_EXPORT_ERROR"
        }

def _get_database_util():
    """取得資料庫工具實例"""
    # 這裡需要實現資料庫工具類別
    return DatabaseExportUtil()
```

### 資料庫工具類別
```python
class DatabaseExportUtil:
    """資料庫匯出工具類別"""
    
    def __init__(self):
        self.db_path = "db/database.db"
        self.ensure_tables_exist()
    
    def ensure_tables_exist(self):
        """確保資料庫表格存在"""
        import sqlite3
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 創建主要表格
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS autocad_elements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                drawing_name TEXT NOT NULL,
                element_id TEXT NOT NULL,
                element_type TEXT NOT NULL,
                layer_name TEXT,
                color INTEGER,
                geometry_data TEXT,
                properties_data TEXT,
                bounds_data TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                version INTEGER DEFAULT 1,
                odoo_sync_status TEXT DEFAULT 'pending',
                odoo_sync_at TIMESTAMP NULL,
                UNIQUE(drawing_name, element_id)
            )
        ''')
        
        # 創建版本控制表格
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS autocad_drawing_versions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                drawing_name TEXT NOT NULL,
                version INTEGER NOT NULL,
                element_count INTEGER,
                export_timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                hash_value TEXT,
                notes TEXT,
                UNIQUE(drawing_name, version)
            )
        ''')
        
        # 創建索引
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_autocad_elements_drawing ON autocad_elements(drawing_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_autocad_elements_type ON autocad_elements(element_type)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_autocad_elements_layer ON autocad_elements(layer_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_autocad_elements_sync ON autocad_elements(odoo_sync_status)')
        
        conn.commit()
        conn.close()
    
    def export_elements_to_database(self, drawing_name, elements, incremental_update=True, include_geometry=True):
        """匯出元素到資料庫"""
        import sqlite3
        import json
        import hashlib
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            new_count = 0
            updated_count = 0
            unchanged_count = 0
            
            for element in elements:
                element_id = element.get("id")
                element_type = element.get("type")
                layer_name = element.get("layer")
                color = element.get("color")
                
                # 準備資料
                geometry_data = json.dumps(element.get("geometry", {})) if include_geometry else None
                properties_data = json.dumps(element.get("properties", {}))
                bounds_data = json.dumps(element.get("bounds", {}))
                
                if incremental_update:
                    # 檢查是否已存在
                    cursor.execute('''
                        SELECT id, geometry_data, properties_data, bounds_data 
                        FROM autocad_elements 
                        WHERE drawing_name = ? AND element_id = ?
                    ''', (drawing_name, element_id))
                    
                    existing = cursor.fetchone()
                    
                    if existing:
                        # 比較資料是否有變化
                        if (existing[1] != geometry_data or 
                            existing[2] != properties_data or 
                            existing[3] != bounds_data):
                            # 更新記錄
                            cursor.execute('''
                                UPDATE autocad_elements 
                                SET element_type = ?, layer_name = ?, color = ?,
                                    geometry_data = ?, properties_data = ?, bounds_data = ?,
                                    updated_at = CURRENT_TIMESTAMP, version = version + 1
                                WHERE id = ?
                            ''', (element_type, layer_name, color, geometry_data, 
                                  properties_data, bounds_data, existing[0]))
                            updated_count += 1
                        else:
                            unchanged_count += 1
                    else:
                        # 新增記錄
                        cursor.execute('''
                            INSERT INTO autocad_elements 
                            (drawing_name, element_id, element_type, layer_name, color,
                             geometry_data, properties_data, bounds_data)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (drawing_name, element_id, element_type, layer_name, color,
                              geometry_data, properties_data, bounds_data))
                        new_count += 1
                else:
                    # 非增量更新，直接插入或替換
                    cursor.execute('''
                        INSERT OR REPLACE INTO autocad_elements 
                        (drawing_name, element_id, element_type, layer_name, color,
                         geometry_data, properties_data, bounds_data)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (drawing_name, element_id, element_type, layer_name, color,
                          geometry_data, properties_data, bounds_data))
                    new_count += 1
            
            # 更新版本記錄
            cursor.execute('''
                SELECT MAX(version) FROM autocad_drawing_versions 
                WHERE drawing_name = ?
            ''', (drawing_name,))
            
            current_version = cursor.fetchone()[0] or 0
            new_version = current_version + 1
            
            # 計算雜湊值
            hash_data = f"{drawing_name}:{len(elements)}:{datetime.now().isoformat()}"
            hash_value = hashlib.md5(hash_data.encode()).hexdigest()
            
            cursor.execute('''
                INSERT INTO autocad_drawing_versions 
                (drawing_name, version, element_count, hash_value, notes)
                VALUES (?, ?, ?, ?, ?)
            ''', (drawing_name, new_version, len(elements), hash_value, 
                  f"增量更新: 新增{new_count}, 更新{updated_count}, 未變{unchanged_count}"))
            
            conn.commit()
            
            return {
                "database_path": self.db_path,
                "new_count": new_count,
                "updated_count": updated_count,
                "unchanged_count": unchanged_count,
                "version": new_version,
                "last_updated": datetime.now().isoformat()
            }
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
```

### 使用範例
```python
# 基本使用 - 匯出整個圖面
result = export_to_database(drawing_name="建築平面圖")

# 只匯出特定元素類型
result = export_to_database(
    drawing_name="結構圖",
    element_types=["line", "circle"]
)

# 匯出並同步到 Odoo
result = export_to_database(
    drawing_name="電氣圖",
    sync_to_odoo=True
)

# 完整匯出（非增量）
result = export_to_database(
    drawing_name="管線圖",
    incremental_update=False,
    include_geometry=True
)

# 過濾特定圖層
result = export_to_database(
    drawing_name="建築圖",
    layer_filter="WALLS",
    element_types=["line"]
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_export_to_database.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime
import os
import tempfile
import sqlite3

import mcp_server_fastmcp

class TestExportToDatabase:
    def test_export_to_database_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {
                    "id": "AcDbLine:1234567890",
                    "type": "line",
                    "layer": "0",
                    "color": 7,
                    "geometry": {
                        "start_point": [0, 0, 0],
                        "end_point": [100, 100, 0],
                        "length": 141.42,
                        "angle": 45.0
                    },
                    "properties": {
                        "linetype": "Continuous",
                        "lineweight": "Default"
                    },
                    "bounds": {
                        "min": [0, 0, 0],
                        "max": [100, 100, 0]
                    }
                }
            ],
            "summary": {
                "total_count": 1,
                "element_counts": {"line": 1}
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["drawing_name"] == "測試圖面"
        assert result["data"]["export_summary"]["total_elements"] == 1
        assert result["data"]["export_summary"]["element_breakdown"]["line"] == 1
        assert "database_path" in result["data"]
        assert "exported_at" in result["data"]
        assert "performance" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_export_to_database_parameter_validation_empty_name(self):
        """參數驗證測試 - 空白圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.export_to_database(drawing_name="")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        
    def test_export_to_database_parameter_validation_whitespace_name(self):
        """參數驗證測試 - 只有空白的圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.export_to_database(drawing_name="   ")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        
    def test_export_to_database_parameter_validation_invalid_element_types(self):
        """參數驗證測試 - 無效元素類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            element_types=["invalid_type"]
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ELEMENT_TYPES"
        
    def test_export_to_database_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_export_to_database_with_specific_element_types(self):
        """特定元素類型匯出測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {"id": "AcDbLine:1", "type": "line", "layer": "0", "color": 7},
                {"id": "AcDbCircle:2", "type": "circle", "layer": "0", "color": 7}
            ],
            "summary": {"total_count": 2}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            element_types=["line", "circle"]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["element_types"] == ["line", "circle"]
        
    def test_export_to_database_incremental_update_false(self):
        """非增量更新測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            incremental_update=False
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["incremental_update"] == False
        
    def test_export_to_database_with_layer_filter(self):
        """圖層過濾測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            layer_filter="WALLS"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["layer_filter"] == "WALLS"
        
    def test_export_to_database_without_geometry(self):
        """不包含幾何資料測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            include_geometry=False
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["include_geometry"] == False
        
    def test_export_to_database_with_odoo_sync(self):
        """Odoo 同步測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_elements.return_value = {"success": True}
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            sync_to_odoo=True
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_info"]["odoo_sync"] == True
        assert result["data"]["sync_info"]["sync_status"] == "success"
        
    def test_export_to_database_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'export_to_database')
        assert callable(mcp_server_fastmcp.export_to_database)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def export_to_database(
    drawing_name: str,
    include_geometry: bool = True,
    incremental_update: bool = True,
    sync_to_odoo: bool = False,
    element_types: list = None,
    layer_filter: str = None
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    if not drawing_name or not drawing_name.strip():
        return {
            "status": "error",
            "message": "圖面名稱不能為空",
            "error_code": "INVALID_DRAWING_NAME"
        }
    
    if element_types is not None:
        if not isinstance(element_types, list):
            return {
                "status": "error",
                "message": "元素類型必須是列表格式",
                "error_code": "INVALID_ELEMENT_TYPES"
            }
        
        valid_types = ["line", "circle", "arc", "text", "dimension", "block"]
        invalid_types = [t for t in element_types if t not in valid_types]
        if invalid_types:
            return {
                "status": "error",
                "message": f"無效的元素類型: {invalid_types}",
                "error_code": "INVALID_ELEMENT_TYPES"
            }
    
    if not _autocad_util:
        return {
            "status": "error",
            "message": "AutoCAD 連接未建立",
            "error_code": "AUTOCAD_NOT_CONNECTED"
        }
    
    # 模擬結果
    return {
        "status": "success",
        "data": {
            "drawing_name": drawing_name.strip(),
            "database_path": "db/database.db",
            "export_summary": {
                "total_elements": 0,
                "new_elements": 0,
                "updated_elements": 0,
                "unchanged_elements": 0,
                "element_breakdown": {}
            },
            "sync_info": {
                "odoo_sync": sync_to_odoo,
                "sync_status": "success" if sync_to_odoo else "not_requested"
            },
            "export_settings": {
                "include_geometry": include_geometry,
                "incremental_update": incremental_update,
                "element_types": element_types or ["all"],
                "layer_filter": layer_filter
            },
            "performance": {
                "scan_time": 0.0,
                "export_time": 0.0,
                "total_time": 0.0
            },
            "exported_at": datetime.now().isoformat()
        },
        "message": "成功匯出 0 個元素到資料庫 (新增: 0, 更新: 0)",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際資料庫操作
@mcp.tool()
def export_to_database(
    drawing_name: str,
    include_geometry: bool = True,
    incremental_update: bool = True,
    sync_to_odoo: bool = False,
    element_types: list = None,
    layer_filter: str = None
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_export_to_database.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_export_to_database.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_export_to_database.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_export_to_database.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 測試覆蓋率要求
- 單元測試覆蓋率: 95%+
- 整合測試覆蓋率: 90%+
- 所有錯誤路徑都要測試
- 所有參數組合都要測試

## 驗證標準

### 功能驗證
- [ ] 基本功能正常運作
- [ ] 參數驗證正確
- [ ] 錯誤處理完整
- [ ] 回傳值格式正確
- [ ] 支援所有指定的參數組合
- [ ] 與 AutoCAD 和資料庫正常互動
- [ ] 增量更新功能正確
- [ ] Odoo 同步功能正確
- [ ] 效能資訊準確
- [ ] 版本控制正確

### 品質標準
- [ ] 代碼覆蓋率 ≥ 95%
- [ ] 所有測試通過
- [ ] 符合代碼風格指南
- [ ] 包含完整日誌記錄
- [ ] 遵循 TDD 原則
- [ ] 無靜態分析警告

### 整合標準
- [ ] 與現有 MCP 系統整合
- [ ] 與 AutoCAD 正常連接
- [ ] 與 SQLite 資料庫正常互動
- [ ] 與 Odoo 系統相容
- [ ] 符合 CLAUDE.md 規則
- [ ] 不影響現有功能
- [ ] 日誌格式一致

## 實施檢查清單

### 開發前
- [ ] 完成 PRP 審查
- [ ] 準備測試資料
- [ ] 確認依賴項目
- [ ] 設定 TDD 環境
- [ ] 準備測試資料庫

### 開發中 (TDD 循環)
- [ ] 寫失敗測試 (Red)
- [ ] 最小實現 (Green)
- [ ] 重構改善 (Refactor)
- [ ] 遵循既有模式
- [ ] 實施錯誤處理
- [ ] 添加日誌記錄

### 開發後
- [ ] 執行所有測試
- [ ] 檢查代碼覆蓋率
- [ ] 進行代碼審查
- [ ] 更新文檔
- [ ] 整合測試
- [ ] 性能測試
- [ ] 資料庫完整性測試

## 相關資源

### 參考文件
- [AutoCAD MCP 整合方案](../../doc/AutoCAD_MCP_Integration_Plan.md)
- [CLAUDE.md](../../CLAUDE.md)
- [Context Engineering 分析](../../doc/Context_Engineering_Analysis.md)
- [TDD + Context Engineering 整合指南](../TDD_CONTEXT_ENGINEERING_GUIDE.md)

### 相關範例
- [MCP 工具範例](../examples/mcp_tool_template.py)
- [資料庫操作範例](../examples/database_operations.py)
- [SQLite 最佳實踐](../examples/sqlite_best_practices.py)

### 現有代碼
- `mcp_server_fastmcp.py` - 主要實施位置
- `utility/util_autocad.py` - AutoCAD 工具類別
- `models/server.py` - 資料庫模型
- `tests/unit/test_scan_elements.py` - 相關測試參考

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。實施時請嚴格遵循 TDD 原則：先寫測試，再寫實現。