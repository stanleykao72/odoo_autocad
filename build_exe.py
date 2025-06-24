#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
優化的 EXE 建置腳本
減少防毒軟體誤報的可能性
"""

import os
import sys
import subprocess
import hashlib
import json
from pathlib import Path

def calculate_file_hash(file_path):
    """計算檔案的 SHA256 雜湊值"""
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def clean_build_dirs():
    """清理建置目錄"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            import shutil
            shutil.rmtree(dir_name)
            print(f"✅ 清理目錄: {dir_name}")

def build_executable():
    """建置可執行檔"""
    print("🚀 開始建置 Odoo-AutoCAD 整合工具...")
    
    # PyInstaller 優化參數
    pyinstaller_args = [
        'pyinstaller',
        '--onefile',                    # 單一檔案
        '--windowed',                   # Windows 應用程式（不顯示控制台）
        '--name', 'odoo-autocad-integration',
        '--icon', 'icon/odoo_autocad.ico',
        
        # 資料檔案
        '--add-data', 'config;config',
        '--add-data', 'fonts;fonts',
        '--add-data', 'icon;icon',
        
        # 隱藏的導入模組
        '--hidden-import', 'customtkinter',
        '--hidden-import', 'win32com.client',
        '--hidden-import', 'win32com.gen_py',
        '--hidden-import', 'pywintypes',
        '--hidden-import', 'win32api',
        '--hidden-import', 'tkinter',
        '--hidden-import', 'tkinter.ttk',
        '--hidden-import', 'PIL',
        '--hidden-import', 'PIL.Image',
        '--hidden-import', 'PIL.ImageTk',
        
        # 排除不需要的模組以減少檔案大小
        '--exclude-module', 'pytest',
        '--exclude-module', 'unittest',
        '--exclude-module', 'doctest',
        '--exclude-module', 'pdb',
        '--exclude-module', 'ipython',
        '--exclude-module', 'jupyter',
        '--exclude-module', 'matplotlib',
        '--exclude-module', 'numpy',
        '--exclude-module', 'pandas',
        
        # 其他優化
        '--clean',                      # 清理快取
        '--noconfirm',                  # 不要確認覆蓋
        '--distpath', 'output',         # 輸出目錄
        
        # 主程式檔案
        'odoo.py'
    ]
    
    # 檢查 UPX 是否可用（可選的壓縮工具）
    try:
        subprocess.run(['upx', '--version'], capture_output=True, check=True)
        pyinstaller_args.extend(['--upx-dir', 'C:/upx'])
        print("✅ 檢測到 UPX，將啟用壓縮")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("ℹ️  未檢測到 UPX，跳過壓縮")
    
    # 執行 PyInstaller
    try:
        result = subprocess.run(pyinstaller_args, check=True, capture_output=True, text=True)
        print("✅ PyInstaller 建置成功")
    except subprocess.CalledProcessError as e:
        print(f"❌ PyInstaller 建置失敗: {e}")
        print(f"錯誤輸出: {e.stderr}")
        return False
    
    return True

def generate_build_info():
    """生成建置資訊"""
    exe_path = Path("output/odoo-autocad-integration.exe")
    
    if not exe_path.exists():
        print("❌ 找不到建置的 EXE 檔案")
        return
    
    # 計算檔案雜湊
    file_hash = calculate_file_hash(exe_path)
    file_size = exe_path.stat().st_size
    
    build_info = {
        "version": "3.0",
        "build_date": str(os.path.getctime(exe_path)),
        "file_size": file_size,
        "sha256_hash": file_hash,
        "filename": exe_path.name,
        "build_environment": {
            "python_version": sys.version,
            "platform": sys.platform,
            "architecture": sys.maxsize > 2**32 and "64-bit" or "32-bit"
        }
    }
    
    # 儲存建置資訊
    info_path = Path("output/build_info.json")
    with open(info_path, 'w', encoding='utf-8') as f:
        json.dump(build_info, f, indent=2, ensure_ascii=False)
    
    print(f"✅ 建置完成！")
    print(f"📁 檔案位置: {exe_path}")
    print(f"📊 檔案大小: {file_size:,} bytes")
    print(f"🔒 SHA256: {file_hash}")
    print(f"📝 建置資訊: {info_path}")

def create_antivirus_readme():
    """創建防毒軟體說明檔"""
    readme_content = """# Odoo-AutoCAD 整合工具

## 防毒軟體誤報說明

此檔案是合法的 Odoo-AutoCAD 整合工具，由承暉精品股份有限公司開發。

### 檔案資訊
- 軟體名稱: Odoo and AutoCAD Integration
- 版本: 3.0
- 開發商: 承暉精品股份有限公司
- 官方網站: https://odoo-esmith.odoo.com/

### 如果被防毒軟體攔截
1. 將檔案標記為安全/信任
2. 將安裝目錄加入防毒軟體的排除清單
3. 聯繫技術支援獲得協助

### 檔案驗證
請核對 build_info.json 中的 SHA256 雜湊值以驗證檔案完整性。

### 技術支援
如有疑問，請聯繫技術支援團隊。
"""
    
    readme_path = Path("output/README_ANTIVIRUS.txt")
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print(f"✅ 創建防毒說明檔: {readme_path}")

def main():
    """主要建置流程"""
    print("🏗️  Odoo-AutoCAD 整合工具建置腳本")
    print("=" * 50)
    
    # 1. 清理舊的建置檔案
    clean_build_dirs()
    
    # 2. 建置可執行檔
    if not build_executable():
        sys.exit(1)
    
    # 3. 生成建置資訊
    generate_build_info()
    
    # 4. 創建防毒說明
    create_antivirus_readme()
    
    print("\n🎉 建置完成！檔案位於 output/ 目錄")
    print("\n💡 減少誤報建議:")
    print("   1. 使用程式碼簽章憑證簽署 EXE 檔案")
    print("   2. 提交檔案到 VirusTotal 進行掃描")
    print("   3. 向防毒廠商申請白名單")

if __name__ == "__main__":
    main()