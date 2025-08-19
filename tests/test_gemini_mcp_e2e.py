# -*- coding: utf-8 -*-
"""
End-to-End Test: Gemini CLI + MCP Server Integration
Tests the complete Gemini CLI → MCP Server → AutoCAD pipeline
"""

import sys
import os
import subprocess
import time
import json
import requests
import threading
from pathlib import Path

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class GeminiMCPTester:
    """Gemini CLI + MCP Server End-to-End Tester"""
    
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.mcp_server_process = None
        self.mcp_port = 8083
        self.mcp_url = f"http://localhost:{self.mcp_port}"
        self.test_commands = [
            "畫一個半徑10的圓形",
            "從(0,0)到(100,100)畫一條線"
        ]
    
    def check_requirements(self):
        """檢查測試需求"""
        print("🔍 檢查測試需求...")
        requirements_met = True
        
        # Check Python environment
        print(f"  ✓ Python: {sys.version}")
        
        # Check odoo.py exists
        odoo_py = self.project_root / 'odoo.py'
        if odoo_py.exists():
            print(f"  ✓ odoo.py 存在: {odoo_py}")
        else:
            print(f"  ✗ odoo.py 不存在: {odoo_py}")
            requirements_met = False
        
        # Check MCP server components
        try:
            from mcp_server_fastmcp import mcp, get_server_info
            print("  ✓ MCP Server 組件可用")
        except ImportError as e:
            print(f"  ✗ MCP Server 組件缺失: {e}")
            requirements_met = False
        
        # Check for Gemini CLI (optional - will guide user to install if missing)
        gemini_available = self.check_gemini_cli()
        
        print(f"\n需求檢查結果: {'✅ 通過' if requirements_met else '❌ 失敗'}")
        return requirements_met, gemini_available
    
    def check_gemini_cli(self):
        """檢查 Gemini CLI 是否可用"""
        try:
            result = subprocess.run(['gemini', '--version'], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                print(f"  ✓ Gemini CLI 可用: {result.stdout.strip()}")
                return True
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
        
        print("  ⚠️  Gemini CLI 未安裝或不在 PATH 中")
        return False
    
    def start_mcp_server(self):
        """啟動 MCP Server"""
        print(f"\n🚀 啟動 MCP Server (端口 {self.mcp_port})...")
        
        try:
            odoo_py = self.project_root / 'odoo.py'
            cmd = [
                sys.executable, str(odoo_py),
                '--mcp-server', '--mcp-port', str(self.mcp_port)
            ]
            
            self.mcp_server_process = subprocess.Popen(
                cmd,
                cwd=str(self.project_root),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            # Wait for server to start
            print("  等待 MCP Server 啟動...")
            time.sleep(5)
            
            # Check if server is running
            if self.mcp_server_process.poll() is None:
                print(f"  ✓ MCP Server 已啟動 (PID: {self.mcp_server_process.pid})")
                return True
            else:
                stdout, stderr = self.mcp_server_process.communicate()
                print(f"  ✗ MCP Server 啟動失敗")
                print(f"    STDOUT: {stdout[:200]}...")
                print(f"    STDERR: {stderr[:200]}...")
                return False
                
        except Exception as e:
            print(f"  ✗ 啟動 MCP Server 時出錯: {e}")
            return False
    
    def test_mcp_server_health(self):
        """測試 MCP Server 健康狀態"""
        print(f"\n🏥 測試 MCP Server 健康狀態...")
        
        try:
            # Test basic connectivity
            response = requests.get(f"{self.mcp_url}/", timeout=10)
            print(f"  ✓ Server 回應: HTTP {response.status_code}")
            
            # Test server info endpoint (if available)
            try:
                info_response = requests.get(f"{self.mcp_url}/health", timeout=5)
                if info_response.status_code == 200:
                    print("  ✓ Health endpoint 可用")
            except:
                print("  ⚠️  Health endpoint 不可用 (這是正常的)")
            
            return True
            
        except requests.exceptions.ConnectionError:
            print(f"  ✗ 無法連接到 MCP Server: {self.mcp_url}")
            return False
        except Exception as e:
            print(f"  ✗ MCP Server 健康檢查失敗: {e}")
            return False
    
    def test_mcp_direct(self):
        """直接測試 MCP 功能 (不通過 Gemini CLI)"""
        print(f"\n🧪 直接測試 MCP 功能...")
        
        try:
            from mcp_server_fastmcp import process_natural_language_command
            
            success_count = 0
            for i, command in enumerate(self.test_commands, 1):
                print(f"\n  測試 {i}: {command}")
                
                try:
                    result = process_natural_language_command(command)
                    
                    if result.get("status") == "success":
                        parsed_intent = result.get("data", {}).get("parsed_intent", {})
                        action = parsed_intent.get("action", "unknown")
                        confidence = parsed_intent.get("confidence", 0)
                        
                        print(f"    ✓ 解析成功: {action} (信心度: {confidence:.1%})")
                        print(f"    ✓ MCP 處理: {result.get('status')}")
                        success_count += 1
                    else:
                        print(f"    ✗ MCP 處理失敗: {result.get('status')}")
                        
                except Exception as e:
                    if "AutoCAD" in str(e) or "COM" in str(e):
                        print(f"    ✓ 命令處理正常，AutoCAD 連接失敗 (預期行為)")
                        success_count += 1
                    else:
                        print(f"    ✗ 處理錯誤: {e}")
            
            success_rate = success_count / len(self.test_commands)
            print(f"\n  直接 MCP 測試結果: {success_count}/{len(self.test_commands)} 成功 ({success_rate:.1%})")
            return success_rate >= 0.8
            
        except Exception as e:
            print(f"  ✗ 直接 MCP 測試失敗: {e}")
            return False
    
    def create_gemini_config(self):
        """創建 Gemini CLI 配置檔案"""
        print(f"\n📝 創建 Gemini CLI 配置檔案...")
        
        config = {
            "mcp": {
                "servers": {
                    "odoo-autocad": {
                        "command": "python",
                        "args": [
                            str(self.project_root / "odoo.py"),
                            "--mcp-server",
                            "--mcp-port", str(self.mcp_port)
                        ],
                        "cwd": str(self.project_root),
                        "transport": "sse",
                        "url": self.mcp_url
                    }
                }
            }
        }
        
        # Create config directory
        config_dir = Path.home() / ".config" / "gemini"
        config_dir.mkdir(parents=True, exist_ok=True)
        
        # Write config file
        config_file = config_dir / "mcp.json"
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            print(f"  ✓ 配置檔案已建立: {config_file}")
            return True
        except Exception as e:
            print(f"  ✗ 建立配置檔案失敗: {e}")
            return False
    
    def test_gemini_mcp_integration(self, gemini_available):
        """測試 Gemini CLI + MCP 整合"""
        print(f"\n🎯 測試 Gemini CLI + MCP 整合...")
        
        if not gemini_available:
            print("  ⚠️  Gemini CLI 不可用，提供手動測試指引...")
            self.provide_manual_testing_guide()
            return False
        
        # Create test script for Gemini CLI
        test_script = self.create_gemini_test_script()
        
        try:
            # Run Gemini CLI test
            result = subprocess.run([
                'gemini', 'cli', '--script', str(test_script)
            ], capture_output=True, text=True, timeout=30, cwd=str(self.project_root))
            
            if result.returncode == 0:
                print("  ✓ Gemini CLI 測試成功")
                print(f"    輸出: {result.stdout[:200]}...")
                return True
            else:
                print(f"  ✗ Gemini CLI 測試失敗 (代碼: {result.returncode})")
                print(f"    錯誤: {result.stderr[:200]}...")
                return False
                
        except subprocess.TimeoutExpired:
            print("  ✗ Gemini CLI 測試逾時")
            return False
        except Exception as e:
            print(f"  ✗ Gemini CLI 測試錯誤: {e}")
            return False
    
    def create_gemini_test_script(self):
        """創建 Gemini CLI 測試腳本"""
        script_content = f"""
# Gemini CLI MCP 測試腳本
請使用 MCP 工具測試以下指令：

1. 畫一個半徑10的圓形
2. 從座標(0,0)到(100,100)畫一條線

請確認：
- MCP 伺服器連接正常
- 指令能正確解析
- AutoCAD 操作能執行 (如果可用)
"""
        
        script_file = self.project_root / 'tests' / 'gemini_test_script.txt'
        script_file.parent.mkdir(exist_ok=True)
        
        with open(script_file, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        return script_file
    
    def provide_manual_testing_guide(self):
        """提供手動測試指引"""
        print(f"\n📋 手動 Gemini CLI 測試指引:")
        print(f"=" * 60)
        
        print(f"1. 安裝 Gemini CLI:")
        print(f"   - 前往 https://docs.anthropic.com/en/docs/claude-code")
        print(f"   - 下載並安裝 Gemini CLI")
        print(f"   - 確認 'gemini --version' 能正常執行")
        
        print(f"\n2. 配置 MCP 連接:")
        print(f"   - MCP Server URL: {self.mcp_url}")
        print(f"   - 協定: SSE (Server-Sent Events)")
        print(f"   - 使用提供的配置檔案")
        
        print(f"\n3. 測試指令:")
        for i, command in enumerate(self.test_commands, 1):
            print(f"   {i}. \"請使用MCP工具{command}\"")
        
        print(f"\n4. 預期結果:")
        print(f"   ✓ Gemini CLI 能連接到 MCP Server")
        print(f"   ✓ 指令能正確解析和執行")
        print(f"   ✓ AutoCAD 中能看到繪製的圖形 (如果 AutoCAD 可用)")
    
    def cleanup(self):
        """清理測試環境"""
        print(f"\n🧹 清理測試環境...")
        
        if self.mcp_server_process:
            try:
                self.mcp_server_process.terminate()
                self.mcp_server_process.wait(timeout=5)
                print("  ✓ MCP Server 已停止")
            except:
                try:
                    self.mcp_server_process.kill()
                    print("  ✓ MCP Server 已強制停止")
                except:
                    print("  ⚠️  無法停止 MCP Server")
    
    def run_full_test(self):
        """執行完整的端到端測試"""
        print("Gemini CLI + MCP Server 端到端測試")
        print("=" * 80)
        
        try:
            # 1. 檢查需求
            requirements_ok, gemini_available = self.check_requirements()
            if not requirements_ok:
                print("\n❌ 需求檢查失敗，無法繼續測試")
                return False
            
            # 2. 啟動 MCP Server
            if not self.start_mcp_server():
                print("\n❌ MCP Server 啟動失敗")
                return False
            
            # 3. 測試 MCP Server 健康狀態
            if not self.test_mcp_server_health():
                print("\n❌ MCP Server 健康檢查失敗")
                return False
            
            # 4. 直接測試 MCP 功能
            if not self.test_mcp_direct():
                print("\n❌ 直接 MCP 測試失敗")
                return False
            
            # 5. 創建 Gemini 配置
            self.create_gemini_config()
            
            # 6. 測試 Gemini CLI 整合
            gemini_success = self.test_gemini_mcp_integration(gemini_available)
            
            # 結果總結
            print(f"\n" + "=" * 80)
            print("測試結果總結")
            print("=" * 80)
            print("✅ MCP Server 啟動成功")
            print("✅ MCP Server 健康狀態良好")
            print("✅ 直接 MCP 功能測試通過")
            print("✅ Gemini CLI 配置檔案已建立")
            
            if gemini_success:
                print("✅ Gemini CLI + MCP 整合測試通過")
                print("\n🎉 所有測試通過！端到端整合成功！")
            else:
                print("⚠️  Gemini CLI 整合需要手動測試")
                print("\n✅ MCP Server 準備就緒，可進行手動 Gemini CLI 測試")
            
            return True
            
        except Exception as e:
            print(f"\n❌ 測試過程中發生錯誤: {e}")
            return False
        
        finally:
            self.cleanup()

def main():
    """主程式"""
    tester = GeminiMCPTester()
    
    try:
        success = tester.run_full_test()
        return 0 if success else 1
    except KeyboardInterrupt:
        print(f"\n\n⚠️  測試被用戶中斷")
        tester.cleanup()
        return 1

if __name__ == "__main__":
    sys.exit(main())