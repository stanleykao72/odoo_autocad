#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
MCP Connection Diagnostic Tool
幫助診斷 Gemini CLI 與 AutoCAD-Odoo MCP 連接問題
"""

# Fix for Windows console encoding
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import json
import subprocess
import time
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))


def test_direct_execution():
    """測試直接執行應用程式"""
    print("1. 測試直接執行應用程式...")
    
    exe_path = "C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe"
    if os.path.exists(exe_path):
        print(f"   ✓ 執行檔存在: {exe_path}")
    else:
        print(f"   ✗ 執行檔不存在: {exe_path}")
        return False
    
    # 測試 --help
    try:
        result = subprocess.run(
            [exe_path, "--help"],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            print("   ✓ 執行檔可以正常運行")
            return True
        else:
            print(f"   ✗ 執行檔運行失敗: {result.stderr}")
            return False
    except Exception as e:
        print(f"   ✗ 執行錯誤: {e}")
        return False


def test_python_execution():
    """測試使用 Python 執行原始碼"""
    print("\n2. 測試使用 Python 執行原始碼...")
    
    src_path = Path(__file__).parent.parent.parent / "odoo.py"
    if not src_path.exists():
        print(f"   ✗ 原始碼不存在: {src_path}")
        return False
    
    print(f"   ✓ 原始碼存在: {src_path}")
    
    # 測試 --help
    try:
        result = subprocess.run(
            [sys.executable, str(src_path), "--help"],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            print("   ✓ Python 可以正常執行原始碼")
            return True
        else:
            print(f"   ✗ Python 執行失敗: {result.stderr}")
            return False
    except Exception as e:
        print(f"   ✗ 執行錯誤: {e}")
        return False


def test_mcp_stdio_communication():
    """測試 MCP stdio 通訊"""
    print("\n3. 測試 MCP stdio 通訊...")
    
    # 創建測試進程
    src_path = Path(__file__).parent.parent.parent / "odoo.py"
    process = subprocess.Popen(
        [sys.executable, str(src_path), "--enable-mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=0,
        env={**os.environ, "PYTHONUNBUFFERED": "1"}
    )
    
    try:
        # 給程式一些時間啟動
        time.sleep(2)
        
        # 檢查進程是否還在運行
        if process.poll() is not None:
            stderr = process.stderr.read()
            print(f"   ✗ 進程已結束，錯誤: {stderr}")
            return False
        
        print("   ✓ MCP 進程正在運行")
        
        # 發送測試請求
        test_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "1.0"
            }
        }
        
        print("   → 發送 initialize 請求...")
        process.stdin.write(json.dumps(test_request) + "\n")
        process.stdin.flush()
        
        # 等待回應（設置超時）
        import select
        if sys.platform == 'win32':
            # Windows doesn't support select on pipes
            # 直接嘗試讀取
            import threading
            response_data = []
            
            def read_response():
                try:
                    line = process.stdout.readline()
                    if line:
                        response_data.append(line)
                except:
                    pass
            
            reader_thread = threading.Thread(target=read_response)
            reader_thread.start()
            reader_thread.join(timeout=3)
            
            if response_data:
                print(f"   ← 收到回應: {response_data[0][:100]}...")
                return True
            else:
                print("   ✗ 沒有收到回應")
                return False
        
    finally:
        # 清理
        process.terminate()
        process.wait()


def suggest_config():
    """建議 Gemini CLI 配置"""
    print("\n4. 建議的 Gemini CLI 配置:")
    
    config = {
        "mcpServers": {
            "autocad-odoo": {
                "command": "python",
                "args": [
                    str(Path(__file__).parent.parent.parent / "odoo.py"),
                    "--enable-mcp"
                ],
                "env": {
                    "PYTHONUNBUFFERED": "1",
                    "PYTHONDONTWRITEBYTECODE": "1"
                }
            }
        }
    }
    
    print(json.dumps(config, indent=2))
    
    print("\n   或使用包裝腳本:")
    
    wrapper_config = {
        "mcpServers": {
            "autocad-odoo": {
                "command": "python",
                "args": [
                    str(Path(__file__).parent.parent.parent / "mcp_wrapper.py")
                ],
                "env": {
                    "PYTHONUNBUFFERED": "1"
                }
            }
        }
    }
    
    print(json.dumps(wrapper_config, indent=2))


def main():
    """主診斷流程"""
    print("=== AutoCAD-Odoo MCP 連接診斷工具 ===\n")
    
    # 執行各項測試
    exe_ok = test_direct_execution()
    py_ok = test_python_execution()
    mcp_ok = test_mcp_stdio_communication()
    
    print("\n=== 診斷結果摘要 ===")
    print(f"直接執行 EXE: {'✓ 成功' if exe_ok else '✗ 失敗'}")
    print(f"Python 執行: {'✓ 成功' if py_ok else '✗ 失敗'}")
    print(f"MCP 通訊: {'✓ 成功' if mcp_ok else '✗ 失敗'}")
    
    if not mcp_ok:
        print("\n⚠️  MCP 通訊測試失敗，可能原因：")
        print("   1. 應用程式的日誌輸出干擾了 stdio 通訊")
        print("   2. MCP 服務沒有正確初始化")
        print("   3. 需要使用包裝腳本來處理 stdio")
    
    # 顯示建議配置
    suggest_config()


if __name__ == "__main__":
    main()