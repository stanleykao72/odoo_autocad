# -*- coding: utf-8 -*-
"""
Diagnose MCP Connection Issues
診斷 Gemini CLI MCP 連接問題
"""

import sys
import os
import requests
import subprocess
import json
from pathlib import Path

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

class MCPConnectionDiagnosis:
    """MCP 連接診斷器"""
    
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.ports_to_check = [8000, 8080, 8083, 3000]
    
    def check_all_ports(self):
        """檢查所有可能的端口"""
        print("🔍 檢查所有可能的 MCP Server 端口...")
        
        active_ports = []
        
        for port in self.ports_to_check:
            try:
                response = requests.get(f"http://localhost:{port}/", timeout=3)
                status = f"HTTP {response.status_code}"
                print(f"  ✅ 端口 {port}: {status}")
                active_ports.append((port, status, response.text[:100]))
                
                # Also try common MCP endpoints
                try:
                    sse_response = requests.get(f"http://localhost:{port}/sse", timeout=2)
                    print(f"      /sse: HTTP {sse_response.status_code}")
                except:
                    print(f"      /sse: 無回應")
                    
                try:
                    mcp_response = requests.get(f"http://localhost:{port}/mcp", timeout=2)
                    print(f"      /mcp: HTTP {mcp_response.status_code}")
                except:
                    print(f"      /mcp: 無回應")
                
            except requests.exceptions.ConnectionError:
                print(f"  ❌ 端口 {port}: 無服務")
            except Exception as e:
                print(f"  ⚠️  端口 {port}: {str(e)[:50]}")
        
        return active_ports
    
    def check_mcp_protocol(self, port):
        """檢查 MCP 協定支援"""
        print(f"\n🔍 檢查端口 {port} 的 MCP 協定支援...")
        
        base_url = f"http://localhost:{port}"
        
        # Test different MCP endpoints
        endpoints = [
            "/",
            "/sse",
            "/mcp", 
            "/health",
            "/status",
            "/tools"
        ]
        
        for endpoint in endpoints:
            try:
                url = f"{base_url}{endpoint}"
                response = requests.get(url, timeout=3)
                
                content_type = response.headers.get('content-type', '')
                
                print(f"  📡 {endpoint}: HTTP {response.status_code}")
                print(f"      Content-Type: {content_type}")
                
                # Check for SSE headers
                if 'text/event-stream' in content_type:
                    print(f"      ✅ SSE endpoint detected!")
                
                # Check response content
                if response.status_code == 200:
                    content = response.text[:200]
                    if 'mcp' in content.lower() or 'tool' in content.lower():
                        print(f"      ✅ MCP-related content detected")
                    
                print()
                
            except Exception as e:
                print(f"  ❌ {endpoint}: {str(e)[:50]}")
    
    def test_sse_connection(self, port):
        """測試 SSE 連接"""
        print(f"\n🌊 測試端口 {port} 的 SSE 連接...")
        
        try:
            import requests
            import time
            
            # Test SSE endpoint with proper headers
            headers = {
                'Accept': 'text/event-stream',
                'Cache-Control': 'no-cache',
                'Connection': 'keep-alive'
            }
            
            url = f"http://localhost:{port}/sse"
            print(f"  📡 嘗試連接: {url}")
            print(f"  📋 Headers: {headers}")
            
            # Make SSE request with stream=True
            response = requests.get(url, headers=headers, stream=True, timeout=5)
            
            print(f"  📊 回應狀態: HTTP {response.status_code}")
            print(f"  📋 回應 Headers: {dict(response.headers)}")
            
            if response.status_code == 200:
                print("  ✅ SSE 連接成功！")
                
                # Try to read a few events
                print("  📺 嘗試讀取 SSE 事件 (3秒內)...")
                
                start_time = time.time()
                for line in response.iter_lines(decode_unicode=True):
                    if time.time() - start_time > 3:  # Stop after 3 seconds
                        break
                    
                    if line:
                        print(f"      📨 Event: {line[:100]}")
                
                return True
            else:
                print(f"  ❌ SSE 連接失敗: HTTP {response.status_code}")
                print(f"      回應內容: {response.text[:200]}")
                return False
                
        except Exception as e:
            print(f"  ❌ SSE 測試錯誤: {e}")
            return False
    
    def generate_gemini_config(self, working_port):
        """生成正確的 Gemini 配置"""
        print(f"\n📝 生成 Gemini CLI 配置 (端口 {working_port})...")
        
        # Test different URL formats
        url_formats = [
            f"http://localhost:{working_port}",
            f"http://localhost:{working_port}/",
            f"http://localhost:{working_port}/sse",
            f"http://localhost:{working_port}/mcp"
        ]
        
        for url in url_formats:
            try:
                response = requests.get(url, timeout=3)
                print(f"  📡 {url}: HTTP {response.status_code}")
                
                if response.status_code in [200, 404]:  # Both can work for MCP
                    config = {
                        "mcpServers": {
                            "autocad": {
                                "url": url,
                                "timeout": 30000,
                                "description": f"AutoCAD-Odoo Integration (Port {working_port})"
                            }
                        }
                    }
                    
                    print(f"\n✅ 建議的 Gemini 配置:")
                    print(json.dumps(config, indent=2, ensure_ascii=False))
                    
                    # Write to file
                    config_file = self.project_root / "tests" / f"gemini_config_port_{working_port}.json"
                    with open(config_file, 'w', encoding='utf-8') as f:
                        json.dump(config, f, indent=2, ensure_ascii=False)
                    
                    print(f"\n📄 配置已儲存至: {config_file}")
                    return url
                    
            except Exception as e:
                print(f"  ❌ {url}: {str(e)[:50]}")
        
        return None
    
    def run_full_diagnosis(self):
        """執行完整診斷"""
        print("🔍 MCP Connection 完整診斷")
        print("=" * 80)
        
        # Step 1: Check all ports
        active_ports = self.check_all_ports()
        
        if not active_ports:
            print("\n❌ 沒有發現任何活躍的服務端口")
            print("💡 請確認 `python odoo.py` 已經啟動")
            return False
        
        print(f"\n✅ 發現 {len(active_ports)} 個活躍端口")
        
        # Step 2: Test each active port for MCP protocol
        best_port = None
        for port, status, content in active_ports:
            print(f"\n{'='*60}")
            print(f"測試端口 {port} (狀態: {status})")
            print(f"{'='*60}")
            
            # Check MCP protocol
            self.check_mcp_protocol(port)
            
            # Test SSE connection
            if self.test_sse_connection(port):
                best_port = port
                break
            else:
                # If SSE fails, still consider it if it has basic HTTP response
                if not best_port and status.startswith("HTTP"):
                    best_port = port
        
        # Step 3: Generate config for best port
        if best_port:
            print(f"\n🎯 推薦使用端口: {best_port}")
            recommended_url = self.generate_gemini_config(best_port)
            
            if recommended_url:
                print(f"\n🚀 請更新 Gemini CLI 設定:")
                print(f'   "url": "{recommended_url}"')
                return True
        
        print("\n❌ 無法找到合適的 MCP 端口配置")
        return False

def main():
    """主程式"""
    diagnosis = MCPConnectionDiagnosis()
    
    try:
        success = diagnosis.run_full_diagnosis()
        
        if success:
            print("\n🎉 診斷完成，請嘗試建議的配置！")
        else:
            print("\n⚠️  診斷未能找到解決方案，請檢查:")
            print("   1. python odoo.py 是否正常啟動")
            print("   2. 查看應用程式日誌中的 MCP 啟動訊息")
            print("   3. 檢查防火牆是否阻擋端口")
        
        return 0 if success else 1
        
    except KeyboardInterrupt:
        print("\n\n⚠️  診斷被用戶中斷")
        return 1

if __name__ == "__main__":
    sys.exit(main())