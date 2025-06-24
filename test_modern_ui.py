# -*- coding: utf-8 -*-
"""
測試現代化UI的啟動檔案
"""
import sys
import os

# 確保可以導入專案模組
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from forms.form_main_modern import ModernFormMain
from ui.ui_fonts import font_manager

def test_ui():
    """測試現代化UI"""
    # 測試字體管理器
    print("=== 字體系統資訊 ===")
    font_info = font_manager.get_system_info()
    for key, value in font_info.items():
        print(f"{key}: {value}")
    
    print("\n=== 啟動現代化UI ===")
    
    # 測試用連接配置
    odoo_connection = {
        'host': 'localhost',
        'db_name': 'test_db',
        'url': 'http://localhost:8069',
        'token': 'test_token'
    }
    
    try:
        app = ModernFormMain(odoo_connection=odoo_connection)
        print("✅ 現代化UI創建成功")
        app.mainloop()
    except Exception as e:
        print(f"❌ UI啟動失敗: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_ui()