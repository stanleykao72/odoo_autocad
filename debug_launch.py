#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
除錯啟動腳本
用於診斷 PyInstaller 打包後的問題
"""

import os
import sys
import traceback
import logging
from pathlib import Path

# 設置詳細的日誌記錄
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('debug.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

def diagnose_environment():
    """診斷執行環境"""
    logger.info("=== 環境診斷開始 ===")
    
    # 1. Python 環境資訊
    logger.info(f"Python 版本: {sys.version}")
    logger.info(f"執行檔路徑: {sys.executable}")
    logger.info(f"當前工作目錄: {os.getcwd()}")
    logger.info(f"腳本位置: {__file__}")
    
    # 2. PyInstaller 相關
    if hasattr(sys, '_MEIPASS'):
        logger.info(f"PyInstaller 臨時目錄: {sys._MEIPASS}")
        logger.info("程式正在以打包模式執行")
    else:
        logger.info("程式正在以開發模式執行")
    
    # 3. 路徑檢查
    paths_to_check = [
        'config',
        'fonts', 
        'icon',
        'db',
        'forms',
        'utility',
        'models'
    ]
    
    logger.info("=== 路徑檢查 ===")
    for path in paths_to_check:
        if os.path.exists(path):
            logger.info(f"✅ {path} 存在")
            if os.path.isdir(path):
                try:
                    files = os.listdir(path)
                    logger.info(f"   包含檔案: {files[:5]}...")  # 只顯示前5個
                except PermissionError:
                    logger.warning(f"   無法讀取目錄內容")
        else:
            logger.warning(f"❌ {path} 不存在")
    
    # 4. 環境變數
    important_vars = ['APPDATA', 'USERPROFILE', 'TEMP', 'PATH']
    logger.info("=== 環境變數 ===")
    for var in important_vars:
        value = os.environ.get(var, '未設定')
        logger.info(f"{var}: {value[:100]}...")  # 截斷長路徑
    
    # 5. 模組導入測試
    logger.info("=== 模組導入測試 ===")
    modules_to_test = [
        'tkinter',
        'customtkinter', 
        'sqlalchemy',
        'yaml',
        'requests',
        'win32com.client'
    ]
    
    for module in modules_to_test:
        try:
            __import__(module)
            logger.info(f"✅ {module} 導入成功")
        except ImportError as e:
            logger.error(f"❌ {module} 導入失敗: {e}")
        except Exception as e:
            logger.error(f"❌ {module} 導入時發生其他錯誤: {e}")

def get_app_data_path():
    """獲取應用程式資料目錄路徑"""
    if sys.platform.startswith('win'):
        # Windows: 使用 %APPDATA% 目錄
        app_data = os.environ.get('APPDATA', os.path.expanduser('~'))
        app_dir = os.path.join(app_data, 'OdooAutoCAD')
    else:
        # macOS/Linux: 使用用戶家目錄
        app_dir = os.path.expanduser('~/.odoo_autocad')
    
    return app_dir

def check_database_setup():
    """檢查資料庫設置"""
    logger.info("=== 資料庫設置檢查 ===")
    
    app_data_dir = get_app_data_path()
    logger.info(f"應用程式資料目錄: {app_data_dir}")
    
    try:
        # 嘗試創建目錄
        os.makedirs(app_data_dir, exist_ok=True)
        logger.info(f"✅ 應用程式目錄創建成功")
        
        # 檢查寫入權限
        test_file = os.path.join(app_data_dir, 'test.txt')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        logger.info(f"✅ 目錄寫入權限正常")
        
        # 檢查資料庫檔案
        db_path = os.path.join(app_data_dir, 'database.db')
        if os.path.exists(db_path):
            logger.info(f"✅ 資料庫檔案存在: {db_path}")
            logger.info(f"   檔案大小: {os.path.getsize(db_path)} bytes")
        else:
            logger.info(f"ℹ️  資料庫檔案不存在，將會創建: {db_path}")
            
    except Exception as e:
        logger.error(f"❌ 資料庫設置檢查失敗: {e}")
        logger.error(traceback.format_exc())

def main():
    """主要函數"""
    logger.info("🚀 開始除錯啟動...")
    
    try:
        # 1. 診斷環境
        diagnose_environment()
        
        # 2. 檢查資料庫設置
        check_database_setup()
        
        # 3. 嘗試導入主程式
        logger.info("=== 嘗試啟動主程式 ===")
        
        # 先測試關鍵模組
        import tkinter
        import customtkinter
        logger.info("✅ GUI 模組導入成功")
        
        import sqlalchemy
        logger.info("✅ 資料庫模組導入成功")
        
        # 嘗試啟動主程式
        logger.info("正在啟動主程式...")
        import odoo
        logger.info("✅ 主程式啟動成功")
        
    except Exception as e:
        logger.error(f"❌ 啟動失敗: {e}")
        logger.error("完整錯誤資訊:")
        logger.error(traceback.format_exc())
        
        # 等待用戶按鍵
        input("\n按 Enter 鍵關閉...")
        sys.exit(1)

if __name__ == "__main__":
    main()