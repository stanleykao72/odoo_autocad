# PRP: generate_boq_from_drawing

> **功能**: 從 AutoCAD 圖面自動生成工程量清單 (BOQ)  
> **類型**: MCP 工具  
> **優先級**: 高  
> **日期**: 2025年7月16日  
> **階段**: 第三階段 Odoo 深度整合

## 功能概述

### 目標
從 AutoCAD 圖面自動分析和提取工程量資訊，生成詳細的工程量清單 (Bill of Quantities, BOQ)。整合 Odoo 產品資料庫，支援自訂規則和計算邏輯，為工程估價和採購提供準確的數據基礎。

### 使用場景
1. **工程估價**: 從圖面自動計算工程量，生成成本估算
2. **採購計劃**: 生成材料清單，支援採購規劃
3. **施工管理**: 提供施工所需的材料和工時資訊
4. **專案管理**: 整合到 Odoo 專案管理系統
5. **AI 助手整合**: 透過自然語言指令生成 BOQ

### 成功標準
- [ ] 能夠從 AutoCAD 圖面自動提取工程量
- [ ] 支援多種元素類型的量化計算
- [ ] 整合 Odoo 產品資料庫
- [ ] 支援自訂計算規則
- [ ] 提供詳細的 BOQ 報告
- [ ] 支援成本計算和分析
- [ ] 支援匯出到多種格式
- [ ] 與現有系統完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def generate_boq_from_drawing(
    drawing_name: str,
    project_id: int = None,
    include_autocad_data: bool = True,
    calculation_rules: dict = None,
    element_types: list = None,
    layer_filter: str = None,
    output_format: str = "odoo",
    currency: str = "TWD"
) -> Dict[str, Any]:
    """從 AutoCAD 圖面自動生成工程量清單"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| drawing_name | str | 是 | - | 圖面名稱 |
| project_id | int | 否 | None | Odoo 專案 ID |
| include_autocad_data | bool | 否 | True | 是否包含 AutoCAD 元素資料 |
| calculation_rules | dict | 否 | None | 自訂計算規則 |
| element_types | list | 否 | None | 要計算的元素類型 |
| layer_filter | str | 否 | None | 圖層過濾器 |
| output_format | str | 否 | "odoo" | 輸出格式 ("odoo", "excel", "json") |
| currency | str | 否 | "TWD" | 貨幣單位 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "boq_info": {
            "drawing_name": "測試圖面",
            "project_id": 123,
            "generated_at": "2025-07-16T10:30:00",
            "calculation_rules": "standard",
            "currency": "TWD"
        },
        "boq_summary": {
            "total_items": 25,
            "total_quantity": 150.5,
            "total_amount": 125000.0,
            "currency": "TWD",
            "item_categories": {
                "materials": 15,
                "labor": 6,
                "equipment": 4
            }
        },
        "boq_items": [
            {
                "item_id": "BOQ_001",
                "description": "混凝土 C25",
                "unit": "m³",
                "quantity": 25.5,
                "unit_price": 3200.0,
                "total_price": 81600.0,
                "category": "materials",
                "odoo_product_id": 456,
                "autocad_elements": [
                    {
                        "element_id": "AcDbSolid:123",
                        "element_type": "solid",
                        "layer": "CONCRETE",
                        "volume": 25.5,
                        "calculation_method": "volume"
                    }
                ]
            },
            {
                "item_id": "BOQ_002",
                "description": "鋼筋 #4",
                "unit": "kg",
                "quantity": 1250.0,
                "unit_price": 35.0,
                "total_price": 43750.0,
                "category": "materials",
                "odoo_product_id": 789,
                "autocad_elements": [
                    {
                        "element_id": "AcDbLine:456",
                        "element_type": "line",
                        "layer": "REBAR",
                        "length": 500.0,
                        "calculation_method": "length_weight"
                    }
                ]
            }
        ],
        "calculation_details": {
            "total_elements_processed": 45,
            "matched_products": 20,
            "unmatched_elements": 5,
            "calculation_time": 3.2,
            "rules_applied": [
                "concrete_volume",
                "rebar_weight",
                "formwork_area"
            ]
        },
        "odoo_integration": {
            "project_updated": True,
            "boq_created": True,
            "boq_id": 789,
            "purchase_requisition_created": False
        },
        "export_info": {
            "format": "odoo",
            "file_path": null,
            "export_url": "https://odoo.example.com/boq/789"
        }
    },
    "message": "成功生成BOQ: 25項工程量清單，總金額 $125,000.00",
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
- [x] 需要 Odoo 連接（使用現有的 `_odoo_util`）
- [x] 需要現有的 scan_elements 功能
- [ ] 需要 BOQ 計算引擎
- [ ] 需要 Odoo 產品資料庫整合
- [ ] 需要工程量計算規則

### 核心邏輯
1. **參數驗證**: 檢查圖面名稱、專案ID、計算規則
2. **圖面分析**: 掃描和分析圖面元素
3. **元素分類**: 根據圖層和類型分類元素
4. **量化計算**: 應用計算規則計算工程量
5. **產品匹配**: 匹配 Odoo 產品資料庫
6. **成本計算**: 計算單價和總價
7. **BOQ 生成**: 生成結構化的工程量清單
8. **整合同步**: 同步到 Odoo 系統

### 錯誤處理
- **參數錯誤**: 圖面名稱無效、專案ID無效、計算規則格式錯誤
- **圖面錯誤**: 無法讀取圖面、圖面元素不完整
- **計算錯誤**: 無法計算工程量、計算規則錯誤
- **產品錯誤**: 無法匹配產品、價格資訊缺失
- **同步錯誤**: Odoo 連接失敗、資料同步失敗

## 範例代碼

### 基本實現
```python
@mcp.tool()
def generate_boq_from_drawing(
    drawing_name: str,
    project_id: int = None,
    include_autocad_data: bool = True,
    calculation_rules: dict = None,
    element_types: list = None,
    layer_filter: str = None,
    output_format: str = "odoo",
    currency: str = "TWD"
) -> Dict[str, Any]:
    """從 AutoCAD 圖面自動生成工程量清單"""
    logger.info(f"generate_boq_from_drawing called with params: {locals()}")
    
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
        
        # 驗證專案ID
        if project_id is not None:
            try:
                project_id = int(project_id)
                if project_id <= 0:
                    return {
                        "status": "error",
                        "message": "專案ID必須是正整數",
                        "error_code": "INVALID_PROJECT_ID"
                    }
            except (ValueError, TypeError):
                return {
                    "status": "error",
                    "message": "專案ID必須是整數",
                    "error_code": "INVALID_PROJECT_ID_TYPE"
                }
        
        # 驗證輸出格式
        valid_formats = ["odoo", "excel", "json"]
        if output_format not in valid_formats:
            return {
                "status": "error",
                "message": f"無效的輸出格式: {output_format}",
                "error_code": "INVALID_OUTPUT_FORMAT",
                "suggestion": f"輸出格式必須是 {valid_formats} 之一"
            }
        
        # 檢查系統連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        odoo_util = _odoo_util
        if not odoo_util:
            return {
                "status": "error",
                "message": "Odoo 連接未建立",
                "error_code": "ODOO_NOT_CONNECTED"
            }
        
        # 開始 BOQ 生成
        generation_start_time = datetime.now()
        
        # 1. 掃描圖面元素
        scan_result = scan_elements(
            element_type="all" if not element_types else element_types,
            include_geometry=True,
            include_properties=True,
            layer_filter=layer_filter
        )
        
        if scan_result["status"] != "success":
            return {
                "status": "error",
                "message": f"無法掃描圖面元素: {scan_result.get('message', '')}",
                "error_code": "SCAN_ELEMENTS_FAILED"
            }
        
        elements = scan_result["data"]["elements"]
        
        # 2. 應用計算規則
        if calculation_rules is None:
            calculation_rules = odoo_util.get_default_calculation_rules()
        
        boq_calculator = BOQCalculator(calculation_rules, currency)
        
        # 3. 計算工程量
        calculation_result = boq_calculator.calculate_quantities(elements)
        
        # 4. 匹配產品和價格
        boq_items = []
        total_amount = 0.0
        matched_products = 0
        unmatched_elements = 0
        
        for calc_item in calculation_result["items"]:
            try:
                # 從 Odoo 匹配產品
                product_match = odoo_util.match_product_by_description(
                    calc_item["description"],
                    calc_item["category"]
                )
                
                if product_match:
                    unit_price = product_match["list_price"]
                    total_price = calc_item["quantity"] * unit_price
                    odoo_product_id = product_match["id"]
                    matched_products += 1
                else:
                    # 使用預設價格或估算
                    unit_price = boq_calculator.estimate_unit_price(calc_item)
                    total_price = calc_item["quantity"] * unit_price
                    odoo_product_id = None
                    unmatched_elements += 1
                
                boq_item = {
                    "item_id": f"BOQ_{len(boq_items) + 1:03d}",
                    "description": calc_item["description"],
                    "unit": calc_item["unit"],
                    "quantity": calc_item["quantity"],
                    "unit_price": unit_price,
                    "total_price": total_price,
                    "category": calc_item["category"],
                    "odoo_product_id": odoo_product_id,
                    "autocad_elements": calc_item.get("elements", [])
                }
                
                boq_items.append(boq_item)
                total_amount += total_price
                
            except Exception as e:
                logger.warning(f"處理BOQ項目時發生錯誤: {calc_item.get('description', 'unknown')} - {str(e)}")
                unmatched_elements += 1
        
        # 5. 分類統計
        item_categories = {}
        for item in boq_items:
            category = item["category"]
            item_categories[category] = item_categories.get(category, 0) + 1
        
        # 6. Odoo 整合
        odoo_integration = {"project_updated": False, "boq_created": False, "boq_id": None}
        
        if output_format == "odoo" and project_id:
            try:
                # 創建或更新 Odoo BOQ
                boq_data = {
                    "project_id": project_id,
                    "drawing_name": drawing_name,
                    "items": boq_items,
                    "total_amount": total_amount,
                    "currency": currency,
                    "generated_at": generation_start_time.isoformat()
                }
                
                boq_result = odoo_util.create_or_update_boq(boq_data)
                odoo_integration = {
                    "project_updated": True,
                    "boq_created": True,
                    "boq_id": boq_result["boq_id"],
                    "purchase_requisition_created": False
                }
                
            except Exception as e:
                logger.error(f"Odoo整合失敗: {str(e)}")
                odoo_integration["error"] = str(e)
        
        # 7. 匯出資訊
        export_info = {
            "format": output_format,
            "file_path": None,
            "export_url": None
        }
        
        if output_format == "odoo" and odoo_integration.get("boq_id"):
            export_info["export_url"] = f"{odoo_util.base_url}/boq/{odoo_integration['boq_id']}"
        
        # 計算處理時間
        calculation_time = (datetime.now() - generation_start_time).total_seconds()
        
        return {
            "status": "success",
            "data": {
                "boq_info": {
                    "drawing_name": drawing_name,
                    "project_id": project_id,
                    "generated_at": generation_start_time.isoformat(),
                    "calculation_rules": "standard",
                    "currency": currency
                },
                "boq_summary": {
                    "total_items": len(boq_items),
                    "total_quantity": sum(item["quantity"] for item in boq_items),
                    "total_amount": total_amount,
                    "currency": currency,
                    "item_categories": item_categories
                },
                "boq_items": boq_items,
                "calculation_details": {
                    "total_elements_processed": len(elements),
                    "matched_products": matched_products,
                    "unmatched_elements": unmatched_elements,
                    "calculation_time": calculation_time,
                    "rules_applied": list(calculation_rules.keys()) if calculation_rules else []
                },
                "odoo_integration": odoo_integration,
                "export_info": export_info
            },
            "message": f"成功生成BOQ: {len(boq_items)}項工程量清單，總金額 ${total_amount:,.2f}",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in generate_boq_from_drawing: {e}")
        return {
            "status": "error",
            "message": f"生成BOQ失敗: {str(e)}",
            "error_code": "BOQ_GENERATION_ERROR"
        }
```

### BOQ 計算器類別
```python
class BOQCalculator:
    """BOQ 計算器"""
    
    def __init__(self, calculation_rules, currency):
        self.calculation_rules = calculation_rules or {}
        self.currency = currency
        self.default_rules = {
            "concrete_volume": self._calculate_concrete_volume,
            "rebar_weight": self._calculate_rebar_weight,
            "formwork_area": self._calculate_formwork_area,
            "paint_area": self._calculate_paint_area,
            "excavation_volume": self._calculate_excavation_volume
        }
    
    def calculate_quantities(self, elements):
        """計算工程量"""
        try:
            items = []
            total_quantity = 0.0
            
            # 按圖層分組元素
            layer_groups = {}
            for element in elements:
                layer = element.get("layer", "0")
                if layer not in layer_groups:
                    layer_groups[layer] = []
                layer_groups[layer].append(element)
            
            # 對每個圖層應用計算規則
            for layer, layer_elements in layer_groups.items():
                layer_items = self._calculate_layer_quantities(layer, layer_elements)
                items.extend(layer_items)
                total_quantity += sum(item["quantity"] for item in layer_items)
            
            return {
                "items": items,
                "total_quantity": total_quantity,
                "layer_count": len(layer_groups)
            }
            
        except Exception as e:
            raise Exception(f"計算工程量時發生錯誤: {str(e)}")
    
    def _calculate_layer_quantities(self, layer, elements):
        """計算圖層工程量"""
        items = []
        
        # 根據圖層名稱決定計算方法
        if "CONCRETE" in layer.upper():
            items.extend(self._calculate_concrete_items(elements))
        elif "REBAR" in layer.upper():
            items.extend(self._calculate_rebar_items(elements))
        elif "FORMWORK" in layer.upper():
            items.extend(self._calculate_formwork_items(elements))
        elif "PAINT" in layer.upper():
            items.extend(self._calculate_paint_items(elements))
        elif "EXCAVATION" in layer.upper():
            items.extend(self._calculate_excavation_items(elements))
        else:
            # 一般元素處理
            items.extend(self._calculate_general_items(layer, elements))
        
        return items
    
    def _calculate_concrete_items(self, elements):
        """計算混凝土工程量"""
        items = []
        total_volume = 0.0
        
        for element in elements:
            if element.get("type") == "solid":
                volume = element.get("geometry", {}).get("volume", 0.0)
                total_volume += volume
        
        if total_volume > 0:
            items.append({
                "description": "混凝土 C25",
                "unit": "m³",
                "quantity": total_volume,
                "category": "materials",
                "elements": elements
            })
        
        return items
    
    def _calculate_rebar_items(self, elements):
        """計算鋼筋工程量"""
        items = []
        total_weight = 0.0
        
        for element in elements:
            if element.get("type") == "line":
                length = element.get("geometry", {}).get("length", 0.0)
                # 假設 #4 鋼筋，每米重量 0.617 kg
                weight = length * 0.617
                total_weight += weight
        
        if total_weight > 0:
            items.append({
                "description": "鋼筋 #4",
                "unit": "kg",
                "quantity": total_weight,
                "category": "materials",
                "elements": elements
            })
        
        return items
    
    def _calculate_formwork_items(self, elements):
        """計算模板工程量"""
        items = []
        total_area = 0.0
        
        for element in elements:
            if element.get("type") in ["rectangle", "polygon"]:
                area = element.get("geometry", {}).get("area", 0.0)
                total_area += area
        
        if total_area > 0:
            items.append({
                "description": "模板",
                "unit": "m²",
                "quantity": total_area,
                "category": "materials",
                "elements": elements
            })
        
        return items
    
    def _calculate_paint_items(self, elements):
        """計算油漆工程量"""
        items = []
        total_area = 0.0
        
        for element in elements:
            if element.get("type") in ["rectangle", "polygon"]:
                area = element.get("geometry", {}).get("area", 0.0)
                total_area += area
        
        if total_area > 0:
            items.append({
                "description": "油漆",
                "unit": "m²",
                "quantity": total_area,
                "category": "materials",
                "elements": elements
            })
        
        return items
    
    def _calculate_excavation_items(self, elements):
        """計算挖土工程量"""
        items = []
        total_volume = 0.0
        
        for element in elements:
            if element.get("type") == "solid":
                volume = element.get("geometry", {}).get("volume", 0.0)
                total_volume += volume
        
        if total_volume > 0:
            items.append({
                "description": "挖土",
                "unit": "m³",
                "quantity": total_volume,
                "category": "labor",
                "elements": elements
            })
        
        return items
    
    def _calculate_general_items(self, layer, elements):
        """計算一般元素工程量"""
        items = []
        
        # 根據元素類型進行基本計算
        element_groups = {}
        for element in elements:
            element_type = element.get("type", "unknown")
            if element_type not in element_groups:
                element_groups[element_type] = []
            element_groups[element_type].append(element)
        
        for element_type, type_elements in element_groups.items():
            if element_type == "line":
                total_length = sum(
                    elem.get("geometry", {}).get("length", 0.0)
                    for elem in type_elements
                )
                if total_length > 0:
                    items.append({
                        "description": f"{layer} 線條",
                        "unit": "m",
                        "quantity": total_length,
                        "category": "materials",
                        "elements": type_elements
                    })
            
            elif element_type == "circle":
                total_area = sum(
                    elem.get("geometry", {}).get("area", 0.0)
                    for elem in type_elements
                )
                if total_area > 0:
                    items.append({
                        "description": f"{layer} 圓形",
                        "unit": "m²",
                        "quantity": total_area,
                        "category": "materials",
                        "elements": type_elements
                    })
        
        return items
    
    def estimate_unit_price(self, calc_item):
        """估算單價"""
        # 基本估價邏輯
        category = calc_item["category"]
        unit = calc_item["unit"]
        
        if category == "materials":
            if unit == "m³":
                return 3000.0  # 立方米材料估價
            elif unit == "m²":
                return 500.0   # 平方米材料估價
            elif unit == "m":
                return 100.0   # 米材料估價
            elif unit == "kg":
                return 35.0    # 公斤材料估價
        
        elif category == "labor":
            if unit == "m³":
                return 800.0   # 立方米工時估價
            elif unit == "m²":
                return 200.0   # 平方米工時估價
            elif unit == "m":
                return 50.0    # 米工時估價
        
        return 100.0  # 預設估價
```

### 使用範例
```python
# 基本 BOQ 生成
result = generate_boq_from_drawing(
    drawing_name="建築平面圖",
    project_id=123
)

# 自訂計算規則
custom_rules = {
    "concrete_volume": {"factor": 1.05, "waste": 0.05},
    "rebar_weight": {"unit_weight": 0.617, "waste": 0.03}
}

result = generate_boq_from_drawing(
    drawing_name="結構圖",
    project_id=123,
    calculation_rules=custom_rules,
    currency="USD"
)

# 特定圖層 BOQ
result = generate_boq_from_drawing(
    drawing_name="詳細圖",
    layer_filter="CONCRETE",
    element_types=["solid"],
    output_format="excel"
)

# 完整 BOQ 生成
result = generate_boq_from_drawing(
    drawing_name="總圖",
    project_id=456,
    include_autocad_data=True,
    calculation_rules=None,  # 使用預設規則
    element_types=None,      # 所有元素類型
    layer_filter=None,       # 所有圖層
    output_format="odoo",
    currency="TWD"
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_generate_boq_from_drawing.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime

import mcp_server_fastmcp

class TestGenerateBoqFromDrawing:
    def test_generate_boq_from_drawing_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_default_calculation_rules.return_value = {}
        mock_odoo_util.match_product_by_description.return_value = {
            "id": 456,
            "name": "混凝土 C25",
            "list_price": 3200.0
        }
        mock_odoo_util.create_or_update_boq.return_value = {"boq_id": 789}
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Mock scan_elements
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {
                    "elements": [
                        {
                            "id": "AcDbSolid:123",
                            "type": "solid",
                            "layer": "CONCRETE",
                            "geometry": {"volume": 25.5}
                        }
                    ]
                }
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面",
                project_id=123
            )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["boq_info"]["drawing_name"] == "測試圖面"
        assert result["data"]["boq_info"]["project_id"] == 123
        assert result["data"]["boq_summary"]["total_items"] > 0
        assert result["data"]["boq_summary"]["total_amount"] > 0
        assert "boq_items" in result["data"]
        assert "calculation_details" in result["data"]
        assert "odoo_integration" in result["data"]
        assert "export_info" in result["data"]
        assert "message" in result
        assert "timestamp" in result
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def generate_boq_from_drawing(
    drawing_name: str,
    project_id: int = None,
    include_autocad_data: bool = True,
    calculation_rules: dict = None,
    element_types: list = None,
    layer_filter: str = None,
    output_format: str = "odoo",
    currency: str = "TWD"
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    if not drawing_name or not drawing_name.strip():
        return {
            "status": "error",
            "message": "圖面名稱不能為空",
            "error_code": "INVALID_DRAWING_NAME"
        }
    
    if not _autocad_util:
        return {
            "status": "error",
            "message": "AutoCAD 連接未建立",
            "error_code": "AUTOCAD_NOT_CONNECTED"
        }
    
    if not _odoo_util:
        return {
            "status": "error",
            "message": "Odoo 連接未建立",
            "error_code": "ODOO_NOT_CONNECTED"
        }
    
    # 模擬結果
    return {
        "status": "success",
        "data": {
            "boq_info": {
                "drawing_name": drawing_name,
                "project_id": project_id,
                "generated_at": datetime.now().isoformat(),
                "calculation_rules": "standard",
                "currency": currency
            },
            "boq_summary": {
                "total_items": 25,
                "total_quantity": 150.5,
                "total_amount": 125000.0,
                "currency": currency,
                "item_categories": {"materials": 15, "labor": 6, "equipment": 4}
            },
            "boq_items": [
                {
                    "item_id": "BOQ_001",
                    "description": "混凝土 C25",
                    "unit": "m³",
                    "quantity": 25.5,
                    "unit_price": 3200.0,
                    "total_price": 81600.0,
                    "category": "materials",
                    "odoo_product_id": 456,
                    "autocad_elements": []
                }
            ],
            "calculation_details": {
                "total_elements_processed": 45,
                "matched_products": 20,
                "unmatched_elements": 5,
                "calculation_time": 3.2,
                "rules_applied": []
            },
            "odoo_integration": {
                "project_updated": True,
                "boq_created": True,
                "boq_id": 789,
                "purchase_requisition_created": False
            },
            "export_info": {
                "format": output_format,
                "file_path": None,
                "export_url": f"https://odoo.example.com/boq/789"
            }
        },
        "message": f"成功生成BOQ: 25項工程量清單，總金額 $125,000.00",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 BOQ 計算邏輯
@mcp.tool()
def generate_boq_from_drawing(
    drawing_name: str,
    project_id: int = None,
    include_autocad_data: bool = True,
    calculation_rules: dict = None,
    element_types: list = None,
    layer_filter: str = None,
    output_format: str = "odoo",
    currency: str = "TWD"
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_generate_boq_from_drawing.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_generate_boq_from_drawing.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_generate_boq_from_drawing.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_generate_boq_from_drawing.py --cov=mcp_server_fastmcp --cov-report=term-missing
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
- [ ] 支援所有指定的輸出格式
- [ ] 與 AutoCAD 和 Odoo 正常互動
- [ ] 工程量計算準確
- [ ] 產品匹配正確
- [ ] 成本計算準確
- [ ] BOQ 格式正確

### 品質標準
- [ ] 代碼覆蓋率 ≥ 95%
- [ ] 所有測試通過
- [ ] 符合代碼風格指南
- [ ] 包含完整日誌記錄
- [ ] 遵循 TDD 原則
- [ ] 無靜態分析警告

### 整合標準
- [ ] 與現有 MCP 系統整合
- [ ] 與 AutoCAD 和 Odoo 正常連接
- [ ] 與其他 MCP 工具相容
- [ ] 符合 CLAUDE.md 規則
- [ ] 不影響現有功能
- [ ] 日誌格式一致

## 實施檢查清單

### 開發前
- [ ] 完成 PRP 審查
- [ ] 準備測試資料
- [ ] 確認依賴項目
- [ ] 設定 TDD 環境

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

## 相關資源

### 參考文件
- [AutoCAD MCP 整合方案](../../doc/AutoCAD_MCP_Integration_Plan.md)
- [CLAUDE.md](../../CLAUDE.md)
- [Context Engineering 分析](../../doc/Context_Engineering_Analysis.md)
- [TDD + Context Engineering 整合指南](../TDD_CONTEXT_ENGINEERING_GUIDE.md)

### 相關範例
- [MCP 工具範例](../examples/mcp_tool_template.py)
- [BOQ 計算範例](../examples/boq_calculation_examples.py)
- [Odoo 整合範例](../examples/odoo_integration_examples.py)

### 現有代碼
- `mcp_server_fastmcp.py` - 主要實施位置
- `utility/util_odoo.py` - Odoo 工具類別
- `utility/util_autocad.py` - AutoCAD 工具類別
- `tests/unit/test_generate_*.py` - 相關測試參考

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。實施時請嚴格遵循 TDD 原則：先寫測試，再寫實現。