# -*- coding: utf-8 -*-
"""
字體管理模組
處理字體載入、回退機制和跨平台相容性
"""
import platform
import tkinter.font as tkfont
from typing import Tuple, List
from .ui_theme import UITheme

class FontManager:
    """字體管理器"""
    
    # 跨平台字體回退清單
    FONT_FAMILIES = {
        'chinese': {
            'Windows': ['Microsoft JhengHei UI', 'Microsoft JhengHei', 'SimHei', 'Arial Unicode MS'],
            'Darwin': ['PingFang TC', 'Heiti TC', 'STHeiti', 'Arial Unicode MS'],
            'Linux': ['Noto Sans CJK TC', 'WenQuanYi Micro Hei', 'AR PL UMing TW', 'DejaVu Sans']
        },
        'english': {
            'Windows': ['Segoe UI', 'Tahoma', 'Arial', 'sans-serif'],
            'Darwin': ['SF Pro Display', 'Helvetica Neue', 'Arial', 'sans-serif'],
            'Linux': ['Ubuntu', 'DejaVu Sans', 'Liberation Sans', 'Arial', 'sans-serif']
        }
    }
    
    def __init__(self):
        self.system = platform.system()
        self.available_fonts = None
        self._font_cache = {}
        self._initialized = False
    
    def _ensure_initialized(self):
        """確保字體管理器已初始化"""
        if not self._initialized:
            try:
                self.available_fonts = set(tkfont.families())
                self._initialized = True
            except Exception:
                # 如果無法獲取字體列表，使用空集合
                self.available_fonts = set()
                self._initialized = True
    
    def get_best_font_family(self, font_type: str = 'chinese') -> str:
        """
        獲取最佳可用字體
        
        Args:
            font_type: 'chinese' 或 'english'
            
        Returns:
            最佳可用字體名稱
        """
        self._ensure_initialized()
        families = self.FONT_FAMILIES.get(font_type, {}).get(self.system, [])
        
        for family in families:
            if self.available_fonts and family in self.available_fonts:
                return family
        
        # 回退到系統預設
        return 'TkDefaultFont'
    
    def create_font(self, font_type: str, size_override: int = None) -> Tuple[str, int, str]:
        """
        創建字體配置
        
        Args:
            font_type: 字體類型 ('title', 'heading', 'body', etc.)
            size_override: 覆蓋字體大小
            
        Returns:
            字體元組 (family, size, weight)
        """
        cache_key = f"{font_type}_{size_override}"
        if cache_key in self._font_cache:
            return self._font_cache[cache_key]
        
        # 獲取基礎字體配置
        base_font = UITheme.get_font(font_type)
        _, size, weight = base_font
        
        # 使用最佳可用字體
        best_family = self.get_best_font_family('chinese')
        
        # 覆蓋大小如果提供
        if size_override:
            size = size_override
        
        font_config = (best_family, size, weight)
        self._font_cache[cache_key] = font_config
        return font_config
    
    def get_font_object(self, font_type: str, size_override: int = None) -> tkfont.Font:
        """
        創建tkinter.Font物件
        
        Args:
            font_type: 字體類型
            size_override: 覆蓋字體大小
            
        Returns:
            tkinter.Font物件
        """
        family, size, weight = self.create_font(font_type, size_override)
        return tkfont.Font(family=family, size=size, weight=weight)
    
    def validate_font_rendering(self) -> bool:
        """
        驗證中文字體渲染是否正常
        
        Returns:
            True if 中文字體可正常顯示
        """
        try:
            test_font = self.get_font_object('body')
            # 測試中文字符寬度
            test_char_width = test_font.measure('中')
            latin_char_width = test_font.measure('A')
            
            # 中文字符應該比拉丁字符寬
            return test_char_width > latin_char_width
        except Exception:
            return False
    
    def get_system_info(self) -> dict:
        """
        獲取系統字體資訊
        
        Returns:
            系統字體資訊字典
        """
        self._ensure_initialized()
        chinese_font = self.get_best_font_family('chinese')
        english_font = self.get_best_font_family('english')
        
        return {
            'system': self.system,
            'chinese_font': chinese_font,
            'english_font': english_font,
            'total_fonts': len(self.available_fonts) if self.available_fonts else 0,
            'chinese_rendering_ok': self.validate_font_rendering()
        }

# 全域字體管理器實例
font_manager = FontManager()

def get_app_font(font_type: str, size_override: int = None) -> Tuple[str, int, str]:
    """
    便利函數：獲取應用程式字體配置
    
    Args:
        font_type: 字體類型
        size_override: 覆蓋字體大小
        
    Returns:
        字體元組
    """
    return font_manager.create_font(font_type, size_override)

def get_app_font_object(font_type: str, size_override: int = None) -> tkfont.Font:
    """
    便利函數：獲取應用程式Font物件
    
    Args:
        font_type: 字體類型
        size_override: 覆蓋字體大小
        
    Returns:
        tkinter.Font物件
    """
    return font_manager.get_font_object(font_type, size_override)