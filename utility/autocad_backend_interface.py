# -*- coding: utf-8 -*-
"""
AutoCAD Backend Interface — ABC 定義 COM/IPC 雙模式的統一介面契約

所有後端 (UtilAutoCAD, UtilAutoCADIPC) 必須繼承此介面，
確保 Dispatcher 和業務邏輯不需要 hasattr() 或模式判斷。
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any


class AutoCADBackendInterface(ABC):
    """AutoCAD 後端統一介面"""

    # === 連線 ===

    @abstractmethod
    def connected_autocad(self) -> bool:
        """檢查 AutoCAD 是否已連線"""
        ...

    @abstractmethod
    def connect_autocad(self, main_body=None):
        """建立 AutoCAD 連線"""
        ...

    # === Layout ===

    @abstractmethod
    def get_active_layout(self) -> Optional[str]:
        """取得目前啟用的 layout 名稱（字串）"""
        ...

    @abstractmethod
    def get_doc_layouts(self) -> List[str]:
        """取得所有非 Model layout 名稱列表"""
        ...

    # === 資料提取 ===

    @abstractmethod
    def get_layouts_values(self) -> dict:
        """提取所有 layout 的 TABLE + Block 資料。
        回傳格式必須為 {"all": [layout_dict, ...]}
        """
        ...

    @abstractmethod
    def get_single_layout_values(self, layout_name: str) -> dict:
        """提取單一 layout 的 TABLE + Block 資料"""
        ...

    @abstractmethod
    def get_block_attributes(self) -> dict:
        """讀取目前 layout 的屬性區塊值"""
        ...

    @abstractmethod
    def set_block_attributes(self, attrs: dict, layout_name: Optional[str] = None) -> bool:
        """寫入屬性區塊值"""
        ...

    # === ID 管理 ===

    @abstractmethod
    def set_layouts_tables_id(self, boq_list):
        """將 header_id + detail_id 寫回 TABLE"""
        ...

    @abstractmethod
    def get_layouts_header_id_to_pr(self) -> dict:
        """收集所有 header_id。回傳格式必須為 {"all": [...]}"""
        ...

    # === 清除 ===

    @abstractmethod
    def clear_table_id(self, layout=None):
        """清除指定 layout 的 TABLE ID"""
        ...

    @abstractmethod
    def clear_all_tables_id(self):
        """清除所有 layout 的 TABLE ID"""
        ...

    # === 繪圖操作 ===

    @abstractmethod
    def draw_line(self, start_point, end_point, layer="0", layout_name=None):
        """繪製直線。layout_name=None 表示目前作用中的配置（非 Model）"""
        ...

    @abstractmethod
    def draw_circle(self, center_point, radius, layer="0", layout_name=None):
        """繪製圓形。layout_name=None 表示目前作用中的配置（非 Model）"""
        ...

    @abstractmethod
    def set_layer(self, layer_name, color=7, create_if_not_exist=True):
        """設定或建立圖層"""
        ...

    @abstractmethod
    def list_layers(self, filter_type="all", sort_by="name", include_details=True):
        """列出圖層"""
        ...

    @abstractmethod
    def scan_elements(self, element_type="all", **kwargs):
        """掃描圖面元素"""
        ...
