# -*- coding: utf-8 -*-
"""
Live Gemini CLI Testing Guide
實際的 Gemini CLI + MCP Server 測試指南
"""

import sys
import os
import time
import requests
import subprocess
from pathlib import Path

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

class LiveGeminiTester:
    """實際 Gemini CLI 測試器"""
    
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.mcp_port = 8083
        self.test_commands = [
            "畫一個半徑10的圓形",
            "從(0,0)到(100,100)畫一條線"
        ]
    
    def check_port_availability(self):
        """檢查端口是否可用"""
        print(f"🔍 檢查端口 {self.mcp_port} 是否可用...")
        
        try:
            response = requests.get(f"http://localhost:{self.mcp_port}/", timeout=3)
            print(f"✅ 端口 {self.mcp_port} 上有服務回應 (HTTP {response.status_code})")
            return True
        except requests.exceptions.ConnectionError:
            print(f"⚠️  端口 {self.mcp_port} 沒有服務回應")
            return False
        except Exception as e:
            print(f"❌ 檢查端口時發生錯誤: {e}")
            return False
    
    def start_application(self):
        """啟動應用程式並等待 MCP Server"""
        print("🚀 準備啟動應用程式...")
        print(f"📋 指令: python odoo.py")
        print(f"🎯 預期: MCP Server 會自動在端口 {self.mcp_port} 啟動")
        print()
        
        # 檢查應用程式是否已經在運行
        if self.check_port_availability():
            print("✅ MCP Server 已經在運行中！")
            return True
        
        print("💡 請在另一個終端視窗執行:")
        print(f"   cd {self.project_root}")
        print(f"   python odoo.py")
        print()
        print("⏳ 等待 MCP Server 啟動...")
        
        # 等待用戶啟動應用程式
        max_wait = 60  # 60 seconds
        wait_interval = 2  # Check every 2 seconds
        
        for i in range(0, max_wait, wait_interval):
            if self.check_port_availability():
                print(f"✅ MCP Server 已啟動！(等待時間: {i}秒)")
                return True
            
            if i % 10 == 0:  # Every 10 seconds
                print(f"⏳ 仍在等待... ({i}/{max_wait}秒)")
            
            time.sleep(wait_interval)
        
        print(f"❌ 等待 {max_wait} 秒後仍無法連接到 MCP Server")
        return False
    
    def test_mcp_direct_connection(self):
        """測試直接 MCP 連接"""
        print("\n🧪 測試直接 MCP 連接...")
        
        try:
            # Test basic connection
            response = requests.get(f"http://localhost:{self.mcp_port}/", timeout=10)
            print(f"✅ 基本連接: HTTP {response.status_code}")
            
            # Test if it's our MCP server by checking response
            if response.status_code in [200, 404, 405]:  # Common responses for MCP servers
                print(f"✅ 這看起來是我們的 MCP Server")
                return True
            else:
                print(f"⚠️  不確定這是否是我們的 MCP Server (狀態碼: {response.status_code})")
                return False
                
        except Exception as e:
            print(f"❌ 直接連接測試失敗: {e}")
            return False
    
    def provide_gemini_setup_guide(self):
        """提供 Gemini CLI 設定指南"""
        print("\n📋 Gemini CLI 設定指南")
        print("=" * 60)
        
        print("1️⃣ 安裝 Gemini CLI:")
        print("   • 如果尚未安裝，請參考: https://docs.anthropic.com/en/docs/claude-code")
        print("   • 確保 'claude --version' 能正常執行")
        print()
        
        print("2️⃣ 連接到 MCP Server:")
        print(f"   • Server URL: http://localhost:{self.mcp_port}")
        print("   • Protocol: SSE (Server-Sent Events)")
        print("   • Connection: 確認連接成功")
        print()
        
        print("3️⃣ 測試指令:")
        for i, cmd in enumerate(self.test_commands, 1):
            print(f"   {i}. \"{cmd}\"")
        print()
        
        print("4️⃣ 預期結果:")
        print("   ✅ Gemini CLI 成功連接到 MCP Server")
        print("   ✅ 指令被正確解析為繪圖操作")
        print("   ✅ 回應包含執行結果和時間資訊")
        print("   ✅ 如果 AutoCAD 開啟，會看到實際圖形")
        print()
    
    def provide_troubleshooting_guide(self):
        """提供故障排除指南"""
        print("🔧 故障排除指南")
        print("=" * 60)
        
        print("❌ 問題: Gemini CLI 無法連接")
        print("   🔍 檢查:")
        print(f"      • MCP Server 是否在端口 {self.mcp_port} 運行")
        print(f"      • 使用 netstat -ano | findstr :{self.mcp_port} 檢查")
        print(f"      • 嘗試 curl http://localhost:{self.mcp_port}")
        print()
        
        print("❌ 問題: 指令無法執行")
        print("   🔍 檢查:")
        print("      • MCP Server 日誌中是否有錯誤")
        print("      • AutoCAD 是否已啟動 (如需實際繪圖)")
        print("      • COM 連接是否正常")
        print()
        
        print("❌ 問題: AutoCAD 連接失敗")
        print("   🔍 解決:")
        print("      • 確保 AutoCAD 應用程式已開啟")
        print("      • 檢查 Windows 用戶權限")
        print("      • 查看應用程式日誌中的詳細錯誤")
        print()
    
    def run_interactive_test(self):
        """執行互動式測試"""
        print("🎯 Live Gemini CLI + MCP Server 測試")
        print("=" * 80)
        
        # Step 1: Check/Start MCP Server
        if not self.start_application():
            print("❌ 無法啟動或連接到 MCP Server")
            return False
        
        # Step 2: Test direct MCP connection
        if not self.test_mcp_direct_connection():
            print("❌ MCP Server 連接測試失敗")
            return False
        
        # Step 3: Provide Gemini setup guide
        self.provide_gemini_setup_guide()
        
        # Step 4: Wait for user to test
        print("⏳ 請按照上述指南設定並測試 Gemini CLI...")
        print("📝 測試完成後，請在下方輸入結果:")
        print()
        
        # Interactive testing
        test_results = []
        
        for i, cmd in enumerate(self.test_commands, 1):
            print(f"🧪 測試 {i}: {cmd}")
            
            while True:
                result = input(f"   這個指令測試結果如何？(成功=y, 失敗=n, 跳過=s): ").strip().lower()
                if result in ['y', 'yes', '成功']:
                    test_results.append(True)
                    print("   ✅ 記錄為成功")
                    break
                elif result in ['n', 'no', '失敗']:
                    test_results.append(False)
                    error = input("   請描述遇到的問題: ").strip()
                    print(f"   ❌ 記錄為失敗: {error}")
                    break
                elif result in ['s', 'skip', '跳過']:
                    print("   ⏭️  跳過此測試")
                    break
                else:
                    print("   請輸入 y/n/s")
        
        # Results summary
        print("\n" + "=" * 80)
        print("測試結果總結")
        print("=" * 80)
        
        success_count = sum(test_results)
        total_count = len(test_results)
        
        print(f"📊 測試結果: {success_count}/{total_count} 成功")
        
        for i, (cmd, result) in enumerate(zip(self.test_commands, test_results), 1):
            status = "✅ 成功" if result else "❌ 失敗"
            print(f"   {i}. {cmd} - {status}")
        
        if success_count == total_count:
            print("\n🎉 所有測試通過！Task 1 端到端測試成功完成！")
            print("✅ Gemini CLI → MCP Server → AutoCAD 整個流程運作正常")
            return True
        elif success_count > 0:
            print(f"\n⚠️  部分測試通過 ({success_count}/{total_count})")
            print("💡 建議檢查失敗的測試並進行故障排除")
            self.provide_troubleshooting_guide()
            return False
        else:
            print("\n❌ 所有測試失敗")
            print("🔧 請參考故障排除指南:")
            self.provide_troubleshooting_guide()
            return False

def main():
    """主程式"""
    tester = LiveGeminiTester()
    
    try:
        success = tester.run_interactive_test()
        
        if success:
            print("\n🏆 Task 1 里程碑測試完成！")
            print("🚀 準備進入 Story 1.1 的 Task 2 開發階段")
        
        return 0 if success else 1
        
    except KeyboardInterrupt:
        print("\n\n⚠️  測試被用戶中斷")
        return 1

if __name__ == "__main__":
    sys.exit(main())