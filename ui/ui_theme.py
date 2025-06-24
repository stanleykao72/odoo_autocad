# -*- coding: utf-8 -*-
"""
UI主題配置模組
定義應用程式的色彩方案、字體和樣式
"""
import customtkinter as ctk
from typing import Tuple

class UITheme:
    """UI主題配置類"""
    
    # 色彩方案
    COLORS = {
        # 主要色調
        'primary': '#2B579A',           # 專業藍
        'primary_dark': '#1e3a6f',      # 深藍
        'primary_light': '#4a7bc8',     # 淺藍
        
        # 次要色調
        'secondary': '#E8F4FD',         # 淺藍背景
        'secondary_dark': '#d1e7fa',    # 稍深的淺藍
        
        # 狀態色調
        'success': '#4CAF50',           # 成功綠
        'success_hover': '#45a049',     # 成功綠懸停
        'warning': '#FF9800',           # 警告橘
        'warning_hover': '#e68900',     # 警告橘懸停
        'error': '#F44336',             # 錯誤紅
        'error_hover': '#da190b',       # 錯誤紅懸停
        'info': '#2196F3',              # 資訊藍
        'info_hover': '#0b7dda',        # 資訊藍懸停
        
        # 中性色調
        'background': '#FFFFFF',        # 主背景
        'surface': '#F8F9FA',           # 表面色
        'border': '#E0E0E0',            # 邊框色
        'text_primary': '#212121',      # 主要文字
        'text_secondary': '#757575',    # 次要文字
        'text_disabled': '#BDBDBD',     # 禁用文字
        
        # 深色模式色調
        'dark_background': '#1E1E1E',   # 深色背景
        'dark_surface': '#2D2D2D',      # 深色表面
        'dark_border': '#404040',       # 深色邊框
        'dark_text_primary': '#FFFFFF', # 深色主要文字
        'dark_text_secondary': '#B0B0B0', # 深色次要文字
    }
    
    # 字體配置
    FONTS = {
        'title': ('Microsoft JhengHei UI', 16, 'bold'),
        'heading': ('Microsoft JhengHei UI', 14, 'bold'),
        'body': ('Microsoft JhengHei UI', 11, 'normal'),
        'small': ('Microsoft JhengHei UI', 9, 'normal'),
        'button': ('Microsoft JhengHei UI', 10, 'normal'),
        'menu': ('Microsoft JhengHei UI', 10, 'normal'),
    }
    
    # 尺寸配置
    SIZES = {
        'border_radius': 8,
        'button_height': 32,
        'entry_height': 32,
        'sidebar_width': 220,
        'topbar_height': 60,
        'status_bar_height': 30,
        'padding_small': 5,
        'padding_medium': 10,
        'padding_large': 20,
    }
    
    # 圖示配置 (使用Unicode字符)
    ICONS = {
        'connect': '🔌',
        'disconnect': '🔌',
        'success': '✓',
        'error': '✗',
        'warning': '⚠',
        'info': 'ℹ',
        'settings': '⚙',
        'help': '❓',
        'log': '📝',
        'refresh': '🔄',
        'export': '📤',
        'import': '📥',
        'search': '🔍',
        'clear': '🗑',
        'save': '💾',
        'load': '📂',
        'autocad': '📐',
        'odoo': '🏢',
        'boq': '📊',
        'pr': '📋',
    }
    
    @classmethod
    def setup_theme(cls, appearance_mode: str = "system", color_theme: str = "blue"):
        """
        設置CustomTkinter主題
        
        Args:
            appearance_mode: "system", "light", "dark"
            color_theme: "blue", "green", "dark-blue"
        """
        ctk.set_appearance_mode(appearance_mode)
        ctk.set_default_color_theme(color_theme)
    
    @classmethod
    def get_color(cls, color_name: str, mode: str = "light") -> str:
        """
        獲取指定的顏色
        
        Args:
            color_name: 顏色名稱
            mode: "light" 或 "dark"
            
        Returns:
            顏色的十六進制值
        """
        if mode == "dark" and f"dark_{color_name}" in cls.COLORS:
            return cls.COLORS[f"dark_{color_name}"]
        return cls.COLORS.get(color_name, cls.COLORS['primary'])
    
    @classmethod
    def get_font(cls, font_type: str) -> Tuple[str, int, str]:
        """
        獲取字體配置
        
        Args:
            font_type: 字體類型
            
        Returns:
            字體元組 (family, size, weight)
        """
        return cls.FONTS.get(font_type, cls.FONTS['body'])
    
    @classmethod
    def get_size(cls, size_name: str) -> int:
        """
        獲取尺寸配置
        
        Args:
            size_name: 尺寸名稱
            
        Returns:
            尺寸值
        """
        return cls.SIZES.get(size_name, 10)
    
    @classmethod
    def get_icon(cls, icon_name: str) -> str:
        """
        獲取圖示字符
        
        Args:
            icon_name: 圖示名稱
            
        Returns:
            Unicode圖示字符
        """
        return cls.ICONS.get(icon_name, '●')

# 主題配置實例
theme = UITheme()

# CustomTkinter 自定義色彩主題
CUSTOM_THEME = {
    "CTk": {
        "fg_color": [theme.COLORS['background'], theme.COLORS['dark_background']]
    },
    "CTkToplevel": {
        "fg_color": [theme.COLORS['background'], theme.COLORS['dark_background']]
    },
    "CTkFrame": {
        "corner_radius": theme.SIZES['border_radius'],
        "border_width": 0,
        "fg_color": [theme.COLORS['surface'], theme.COLORS['dark_surface']],
        "top_fg_color": [theme.COLORS['surface'], theme.COLORS['dark_surface']],
        "border_color": [theme.COLORS['border'], theme.COLORS['dark_border']]
    },
    "CTkButton": {
        "corner_radius": theme.SIZES['border_radius'],
        "border_width": 0,
        "fg_color": [theme.COLORS['primary'], theme.COLORS['primary']],
        "hover_color": [theme.COLORS['primary_dark'], theme.COLORS['primary_light']],
        "border_color": [theme.COLORS['primary'], theme.COLORS['primary']],
        "text_color": ["white", "white"],
        "text_color_disabled": [theme.COLORS['text_disabled'], theme.COLORS['text_disabled']]
    },
    "CTkLabel": {
        "corner_radius": 0,
        "fg_color": "transparent",
        "text_color": [theme.COLORS['text_primary'], theme.COLORS['dark_text_primary']]
    },
    "CTkEntry": {
        "corner_radius": theme.SIZES['border_radius'],
        "border_width": 1,
        "fg_color": [theme.COLORS['background'], theme.COLORS['dark_surface']],
        "border_color": [theme.COLORS['border'], theme.COLORS['dark_border']],
        "text_color": [theme.COLORS['text_primary'], theme.COLORS['dark_text_primary']],
        "placeholder_text_color": [theme.COLORS['text_secondary'], theme.COLORS['dark_text_secondary']]
    },
    "CTkComboBox": {
        "corner_radius": theme.SIZES['border_radius'],
        "border_width": 1,
        "fg_color": [theme.COLORS['background'], theme.COLORS['dark_surface']],
        "border_color": [theme.COLORS['border'], theme.COLORS['dark_border']],
        "button_color": [theme.COLORS['primary'], theme.COLORS['primary']],
        "button_hover_color": [theme.COLORS['primary_dark'], theme.COLORS['primary_light']],
        "text_color": [theme.COLORS['text_primary'], theme.COLORS['dark_text_primary']],
        "text_color_disabled": [theme.COLORS['text_disabled'], theme.COLORS['text_disabled']]
    }
}

def apply_custom_theme():
    """應用自定義主題到CustomTkinter"""
    # 注意：CustomTkinter 5.2.0+ 可能需要不同的主題應用方式
    # 這裡使用基本的主題設置
    UITheme.setup_theme("system", "blue")