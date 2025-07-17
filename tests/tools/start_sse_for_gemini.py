#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
啟動 SSE 伺服器供 Gemini CLI 連接
保持運行直到手動停止
"""

import time
import signal
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager

# 全域變數
manager = None

def signal_handler(sig, frame):
    """處理中斷信號"""
    print("\n\n[INFO] 收到中斷信號，正在停止 SSE 伺服器...")
    if manager:
        manager.stop_server()
    print("[INFO] SSE 伺服器已停止")
    sys.exit(0)

def main():
    """主函數"""
    global manager
    
    print("=== 啟動 SSE 伺服器供 Gemini CLI 連接 ===")
    print("端口: 8083")
    print("端點: http://localhost:8083/sse")
    print("按 Ctrl+C 停止伺服器\n")
    
    # 註冊信號處理器
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    # 創建 SSE 管理器
    manager = MCPSSEManager(port=8083)
    
    def status_callback(message, is_running):
        print(f"[SSE] {message}")
    
    manager.set_status_callback(status_callback)
    
    # 啟動伺服器
    print("[INFO] 正在啟動 SSE 伺服器...")
    success = manager.start_server()
    
    if not success:
        print("[ERROR] SSE 伺服器啟動失敗！")
        return 1
    
    # 等待伺服器完全啟動
    time.sleep(3)
    
    # 顯示狀態
    status = manager.get_server_status()
    print(f"\n[SUCCESS] SSE 伺服器已啟動！")
    print(f"運行狀態: {status['is_running']}")
    print(f"端口: {status['port']}")
    print(f"健康檢查: {status['health_check']}")
    
    if 'server_name' in status:
        print(f"伺服器名稱: {status['server_name']}")
        print(f"版本: {status['server_version']}")
    
    print(f"\n[INFO] Gemini CLI 現在可以連接到:")
    print(f"       http://localhost:{status['port']}/sse")
    print(f"\n[INFO] 在 Gemini CLI 中使用以下配置:")
    print(f'''{{
  "autocad-odoo-sse": {{
    "url": "http://localhost:{status['port']}/sse",
    "timeout": 30000,
    "description": "AutoCAD-Odoo Integration with SSE"
  }}
}}''')
    
    print(f"\n[INFO] 伺服器正在運行，按 Ctrl+C 停止...")
    
    # 保持運行
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass  # 由信號處理器處理

if __name__ == "__main__":
    sys.exit(main())