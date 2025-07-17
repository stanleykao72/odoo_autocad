#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
運行所有 SSE 相關測試的腳本
"""

import sys
import os
import subprocess
import time

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

def run_test(test_file, timeout=20):
    """運行單個測試並返回結果"""
    print(f"\n{'='*50}")
    print(f"運行測試: {test_file}")
    print('='*50)
    
    try:
        process = subprocess.Popen(
            [sys.executable, test_file],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8'
        )
        
        output, _ = process.communicate(timeout=timeout)
        return_code = process.returncode
        
        print(output)
        
        if return_code == 0:
            print(f"✓ {test_file} 測試通過")
            return True
        else:
            print(f"✗ {test_file} 測試失敗 (返回碼: {return_code})")
            return False
            
    except subprocess.TimeoutExpired:
        process.kill()
        print(f"⚠ {test_file} 測試超時 ({timeout}秒)")
        return False
    except Exception as e:
        print(f"✗ {test_file} 測試錯誤: {e}")
        return False

def main():
    """運行所有 SSE 測試"""
    print("開始運行 SSE 相關測試...")
    
    tests = [
        ("unit/test_direct_sse_manager.py", 15),
        ("integration/test_simple_sse.py", 10),
        ("integration/test_gui_sse.py", 20),
        ("tools/diagnose_gui_sse.py", 25),
        ("tools/test_mcp_sse_compliance.py", 30),
    ]
    
    results = []
    
    for test_file, timeout in tests:
        test_path = os.path.join(os.path.dirname(__file__), test_file)
        if os.path.exists(test_path):
            result = run_test(test_path, timeout)
            results.append((test_file, result))
        else:
            print(f"⚠ 測試檔案不存在: {test_file}")
            results.append((test_file, False))
        
        # 等待進程完全停止
        time.sleep(2)
    
    # 顯示總結
    print(f"\n{'='*50}")
    print("測試總結")
    print('='*50)
    
    passed = 0
    total = len(results)
    
    for test_file, result in results:
        status = "通過" if result else "失敗"
        icon = "✓" if result else "✗"
        print(f"{icon} {test_file}: {status}")
        if result:
            passed += 1
    
    print(f"\n總計: {passed}/{total} 個測試通過")
    
    if passed == total:
        print("🎉 所有測試都通過！")
        return 0
    else:
        print("❌ 部分測試失敗")
        return 1

if __name__ == "__main__":
    sys.exit(main())