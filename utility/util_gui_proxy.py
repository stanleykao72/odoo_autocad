# -*- coding: utf-8 -*-
"""
GUI代理執行系統 - 解決MCP Server與GUI間的COM線程問題

此模組提供線程安全的方式讓MCP Server通過GUI主線程執行AutoCAD操作
"""

import queue
import threading
import time
from typing import Any, Dict, Optional, Callable
import logging


class GUIProxy:
    """GUI代理執行器 - 處理跨線程COM操作"""

    # 預設等待 GUI 主線程回覆的秒數
    DEFAULT_TIMEOUT = 10.0

    def __init__(self, logger=None, timeout: float = None):
        self.logger = logger or logging.getLogger(__name__)

        # 請求隊列；每個請求自帶一個回覆 queue，避免共用 cache 造成洩漏
        self.request_queue = queue.Queue()
        self.request_id_counter = 0
        self.lock = threading.Lock()
        self.timeout = timeout if timeout is not None else self.DEFAULT_TIMEOUT

        # GUI處理器函數註冊
        self.handlers = {}

        # 狀態
        self.is_running = False

    def register_handler(self, action_name: str, handler_func: Callable):
        """註冊GUI處理器函數"""
        self.handlers[action_name] = handler_func
        if hasattr(self.logger, 'info'):
            self.logger.info(f"[GUI Proxy] 註冊處理器: {action_name}")
        
    def execute_in_gui(self, action: str, timeout: float = None, **kwargs) -> Dict[str, Any]:
        """
        在GUI線程中執行操作

        Args:
            action: 要執行的操作名稱
            timeout: 等待秒數（None 使用 self.timeout）
            **kwargs: 操作參數

        Returns:
            Dict[str, Any]: 操作結果
        """
        with self.lock:
            self.request_id_counter += 1
            request_id = f"req_{self.request_id_counter}"

        # 每個請求自帶回覆 queue：超時後 queue 隨請求一起被回收，
        # 不會像共用 cache 那樣留下無人取走的響應而持續累積記憶體。
        reply_queue = queue.Queue(maxsize=1)
        request = {
            "id": request_id,
            "action": action,
            "params": kwargs,
            "timestamp": time.time(),
            "reply_queue": reply_queue,
        }

        if hasattr(self.logger, 'info'):
            self.logger.info(f"[GUI Proxy] 發送請求: {action} (ID: {request_id})")

        # 發送到GUI線程
        self.request_queue.put(request)

        # 阻塞等待回覆（不再忙等輪詢）
        wait = self.timeout if timeout is None else timeout
        try:
            response = reply_queue.get(timeout=wait)
            if hasattr(self.logger, 'info'):
                self.logger.info(f"[GUI Proxy] 收到響應: {request_id}")
            return response
        except queue.Empty:
            if hasattr(self.logger, 'error'):
                self.logger.error(f"[GUI Proxy] 請求超時: {request_id}")
            return {
                "success": False,
                "error": f"GUI代理執行超時 (操作: {action})",
                "timeout": True
            }


    def process_requests(self):
        """
        處理請求隊列 (在GUI主線程中調用)
        
        Returns:
            int: 處理的請求數量
        """
        processed = 0
        
        try:
            while not self.request_queue.empty():
                try:
                    request = self.request_queue.get_nowait()
                    processed += 1
                    
                    request_id = request["id"]
                    action = request["action"]
                    params = request["params"]
                    
                    if hasattr(self.logger, 'info'):
                        self.logger.info(f"[GUI Proxy] 處理請求: {action} (ID: {request_id})")
                    
                    # 查找處理器
                    if action not in self.handlers:
                        response = {
                            "success": False,
                            "error": f"未知操作: {action}",
                            "available_actions": list(self.handlers.keys())
                        }
                    else:
                        try:
                            # 執行處理器
                            handler = self.handlers[action]
                            response = handler(**params)
                            
                            # 確保響應是dict格式
                            if not isinstance(response, dict):
                                response = {"success": True, "result": response}
                                
                        except Exception as e:
                            if hasattr(self.logger, 'error'):
                                self.logger.error(f"[GUI Proxy] 處理器執行失敗: {e}")
                            response = {
                                "success": False,
                                "error": f"處理器執行錯誤: {str(e)}",
                                "exception_type": type(e).__name__
                            }
                    
                    # 回傳響應；若呼叫端已超時離開，put_nowait 會滿而被丟棄
                    reply_queue = request.get("reply_queue")
                    if reply_queue is not None:
                        try:
                            reply_queue.put_nowait(response)
                        except queue.Full:
                            if hasattr(self.logger, 'warning'):
                                self.logger.warning(
                                    f"[GUI Proxy] 呼叫端已離開，丟棄響應: {request_id}")

                except queue.Empty:
                    break
                except Exception as e:
                    if hasattr(self.logger, 'error'):
                        self.logger.error(f"[GUI Proxy] 處理請求時發生錯誤: {e}")
                    
        except Exception as e:
            if hasattr(self.logger, 'error'):
                self.logger.error(f"[GUI Proxy] process_requests 發生錯誤: {e}")
            
        return processed


# 全域GUI代理實例
_gui_proxy_instance = None

def get_gui_proxy() -> GUIProxy:
    """獲取全域GUI代理實例"""
    global _gui_proxy_instance
    if _gui_proxy_instance is None:
        _gui_proxy_instance = GUIProxy()
    return _gui_proxy_instance


def setup_gui_proxy_handlers(autocad_util, logger):
    """設置GUI代理處理器"""
    proxy = get_gui_proxy()
    
    def log_message(msg):
        """兼容不同logger的日誌方法"""
        if hasattr(logger, 'info'):
            logger.info(msg)
        else:
            logger.safe_log_insert(f"{msg}\n")
    
    def handle_switch_layout(layout_name: str) -> Dict[str, Any]:
        """在GUI線程中處理layout切換"""
        try:
            log_message(f"[GUI Proxy] 開始切換layout: {layout_name}")
            
            # 檢查AutoCAD連接
            if not autocad_util.connected_autocad():
                return {"success": False, "error": "AutoCAD 未連接"}
            
            # 獲取當前document
            if not hasattr(autocad_util, 'acad') or not autocad_util.acad:
                return {"success": False, "error": "AutoCAD 應用程序不可用"}
                
            active_doc = autocad_util.acad.ActiveDocument
            if not active_doc:
                return {"success": False, "error": "無法獲取 ActiveDocument"}
            
            # 使用CTAB系統變數切換
            try:
                active_doc.SetVariable("CTAB", layout_name)
                log_message(f"[GUI Proxy] CTAB切換完成: {layout_name}")
                
                # 短暫等待
                time.sleep(0.1)
                
                # 驗證切換
                current_tab = active_doc.GetVariable("CTAB")
                if current_tab == layout_name:
                    return {
                        "success": True,
                        "message": f"成功切換到 layout: {layout_name}",
                        "current_layout": current_tab
                    }
                else:
                    return {
                        "success": False,
                        "error": f"切換未生效，當前: {current_tab}，預期: {layout_name}"
                    }
                    
            except Exception as e:
                log_message(f"[GUI Proxy] CTAB切換失敗: {e}")
                return {"success": False, "error": f"切換失敗: {str(e)}"}
            
        except Exception as e:
            log_message(f"[GUI Proxy] handle_switch_layout 錯誤: {e}")
            return {"success": False, "error": f"處理錯誤: {str(e)}"}
    
    def handle_get_current_layout() -> Dict[str, Any]:
        """獲取當前layout"""
        try:
            current_layout = autocad_util.get_active_layout()
            if current_layout:
                return {
                    "success": True,
                    "current_layout": {
                        "name": current_layout,
                    }
                }
            else:
                return {"success": False, "error": "無法獲取當前layout"}
        except Exception as e:
            return {"success": False, "error": f"錯誤: {str(e)}"}
    
    def handle_extract_parameters(drawing_path: str = None, use_current_drawing: bool = True) -> Dict[str, Any]:
        """提取AutoCAD參數"""
        try:
            log_message(f"[GUI Proxy] 開始提取參數，路徑: {drawing_path}, 使用當前圖檔: {use_current_drawing}")
            
            if not autocad_util.connected_autocad():
                return {"success": False, "error": "AutoCAD 未連接"}
                
            # 調用AutoCAD參數提取
            result = autocad_util.get_layouts_values()
            
            if result and len(result) > 0:
                log_message(f"[GUI Proxy] 成功提取 {len(result)} 個參數")
                return {
                    "success": True,
                    "parameters": result,
                    "count": len(result)
                }
            else:
                return {"success": False, "error": "未找到任何參數"}
                
        except Exception as e:
            log_message(f"[GUI Proxy] 參數提取失敗: {e}")
            return {"success": False, "error": f"參數提取失敗: {str(e)}"}
    
    def handle_get_autocad_status() -> Dict[str, Any]:
        """獲取AutoCAD詳細狀態"""
        try:
            if not autocad_util.connected_autocad():
                return {"success": False, "error": "AutoCAD 未連接", "connected": False}
            
            status = {
                "connected": True,
                "success": True
            }
            
            # 獲取文檔資訊
            try:
                if hasattr(autocad_util, 'acad') and autocad_util.acad:
                    if hasattr(autocad_util.acad, 'ActiveDocument'):
                        doc = autocad_util.acad.ActiveDocument
                        if doc and hasattr(doc, 'Name'):
                            status["document"] = {
                                "name": str(doc.Name),
                                "path": getattr(doc, 'Path', ''),
                                "saved": getattr(doc, 'Saved', True)
                            }
                    
                    # 獲取應用程序資訊
                    if hasattr(autocad_util.acad, 'Version'):
                        status["application"] = {
                            "version": str(autocad_util.acad.Version),
                            "name": getattr(autocad_util.acad, 'Name', 'AutoCAD')
                        }
            except Exception as e:
                log_message(f"[GUI Proxy] 獲取詳細狀態時出錯: {e}")
            
            return status
            
        except Exception as e:
            log_message(f"[GUI Proxy] 獲取AutoCAD狀態失敗: {e}")
            return {"success": False, "error": f"狀態獲取失敗: {str(e)}", "connected": False}
    
    def handle_draw_line(start_point: list, end_point: list, layer: str = None) -> Dict[str, Any]:
        """在AutoCAD中繪製直線"""
        try:
            log_message(f"[GUI Proxy] 繪製直線從 {start_point} 到 {end_point}")
            
            if not autocad_util.connected_autocad():
                return {"success": False, "error": "AutoCAD 未連接"}
            
            result = autocad_util.draw_line(start_point, end_point, layer)
            
            if result:
                return {"success": True, "message": "直線繪製成功", "object_created": True}
            else:
                return {"success": False, "error": "直線繪製失敗"}
                
        except Exception as e:
            log_message(f"[GUI Proxy] 繪製直線失敗: {e}")
            return {"success": False, "error": f"繪製失敗: {str(e)}"}
    
    def handle_draw_circle(center_point: list, radius: float, layer: str = None) -> Dict[str, Any]:
        """在AutoCAD中繪製圓形"""
        try:
            log_message(f"[GUI Proxy] 繪製圓形，中心: {center_point}, 半徑: {radius}")
            
            if not autocad_util.connected_autocad():
                return {"success": False, "error": "AutoCAD 未連接"}
            
            result = autocad_util.draw_circle(center_point, radius, layer)
            
            if result:
                return {"success": True, "message": "圓形繪製成功", "object_created": True}
            else:
                return {"success": False, "error": "圓形繪製失敗"}
                
        except Exception as e:
            log_message(f"[GUI Proxy] 繪製圓形失敗: {e}")
            return {"success": False, "error": f"繪製失敗: {str(e)}"}
    
    def handle_export_layout_image(layout_name: str, export_path: str = None, image_format: str = "wmf") -> Dict[str, Any]:
        """匯出指定layout為圖像檔案"""
        try:
            log_message(f"[GUI Proxy] 開始匯出layout '{layout_name}' 為圖像")
            
            if not autocad_util.connected_autocad():
                return {"success": False, "error": "AutoCAD 未連接"}
            
            # 首先獲取當前layout以便稍後恢復
            original_name = autocad_util.get_active_layout()

            try:
                # 切換到指定layout
                if original_name != layout_name:
                    # 獲取所有layouts (returns list of strings)
                    layouts = autocad_util.get_doc_layouts()

                    if layout_name not in layouts:
                        return {
                            "success": False,
                            "error": f"找不到名為 '{layout_name}' 的layout"
                        }

                    # 使用CTAB切換layout
                    if hasattr(autocad_util, 'acad') and autocad_util.acad:
                        active_doc = autocad_util.acad.ActiveDocument
                        if active_doc:
                            active_doc.SetVariable("CTAB", layout_name)
                            time.sleep(0.2)  # 等待切換完成
                
                # 設定匯出路徑
                if not export_path:
                    import os
                    export_path = os.path.join(os.getcwd(), f"layout_{layout_name}.{image_format}")
                
                # 執行匯出
                if hasattr(autocad_util, 'acad') and autocad_util.acad:
                    active_doc = autocad_util.acad.ActiveDocument
                    
                    # 獲取layout的範圍
                    current_layout = active_doc.ActiveLayout
                    
                    try:
                        # 獲取layout的可列印範圍
                        if hasattr(current_layout, 'GetExtents'):
                            min_ext, max_ext = current_layout.GetExtents()
                        else:
                            # 使用預設範圍 - 獲取所有實體的範圍
                            if current_layout.Name == "Model":
                                entities = active_doc.ModelSpace
                            else:
                                entities = active_doc.PaperSpace
                            
                            # 計算實體範圍
                            min_x, min_y = float('inf'), float('inf')
                            max_x, max_y = float('-inf'), float('-inf')
                            
                            entity_count = 0
                            for entity in entities:
                                try:
                                    if hasattr(entity, 'GetBoundingBox'):
                                        bbox = entity.GetBoundingBox()
                                        if len(bbox) >= 2:
                                            min_pt, max_pt = bbox[0], bbox[1]
                                            min_x = min(min_x, min_pt[0])
                                            min_y = min(min_y, min_pt[1])
                                            max_x = max(max_x, max_pt[0])
                                            max_y = max(max_y, max_pt[1])
                                            entity_count += 1
                                except Exception:
                                    continue
                            
                            if entity_count > 0:
                                min_ext = [min_x, min_y]
                                max_ext = [max_x, max_y]
                            else:
                                # 使用預設範圍
                                min_ext = [0, 0]
                                max_ext = [200, 100]
                        
                        # 執行匯出
                        active_doc.Export(export_path, image_format.upper(), min_ext + max_ext)
                        
                        # 檢查檔案是否成功建立
                        import os
                        if os.path.exists(export_path):
                            file_size = os.path.getsize(export_path)
                            log_message(f"[GUI Proxy] 圖像匯出成功: {export_path} ({file_size} bytes)")
                            
                            return {
                                "success": True,
                                "message": f"成功匯出layout '{layout_name}' 為圖像",
                                "export_path": export_path,
                                "file_size": file_size,
                                "image_format": image_format.upper(),
                                "layout_name": layout_name
                            }
                        else:
                            return {"success": False, "error": "圖像檔案未成功建立"}
                            
                    except Exception as export_error:
                        log_message(f"[GUI Proxy] 匯出過程發生錯誤: {export_error}")
                        return {"success": False, "error": f"匯出失敗: {str(export_error)}"}
                
                return {"success": False, "error": "無法存取AutoCAD文件"}
                
            finally:
                # 恢復到原始layout
                if original_name and original_name != layout_name:
                    try:
                        if hasattr(autocad_util, 'acad') and autocad_util.acad:
                            active_doc = autocad_util.acad.ActiveDocument
                            if active_doc:
                                active_doc.SetVariable("CTAB", original_name)
                                log_message(f"[GUI Proxy] 恢復到原始layout: {original_name}")
                    except Exception as restore_error:
                        log_message(f"[GUI Proxy] 恢復layout失敗: {restore_error}")
                
        except Exception as e:
            log_message(f"[GUI Proxy] 圖像匯出失敗: {e}")
            return {"success": False, "error": f"匯出失敗: {str(e)}"}
    
    def handle_extract_layout_parameters(layout_name: str) -> Dict[str, Any]:
        """提取指定layout的AutoCAD參數"""
        try:
            log_message(f"[GUI Proxy] 開始提取layout參數: {layout_name}")
            
            if not autocad_util.connected_autocad():
                return {"success": False, "error": "AutoCAD 未連接"}
            
            # 首先獲取當前layout以便稍後恢復
            original_name = autocad_util.get_active_layout()
            log_message(f"[GUI Proxy] 當前layout: {original_name}")

            try:
                # 切換到指定layout
                if original_name != layout_name:
                    log_message(f"[GUI Proxy] 切換到目標layout: {layout_name}")
                    # 獲取所有layouts (returns list of strings)
                    layouts = autocad_util.get_doc_layouts()

                    if layout_name not in layouts:
                        return {
                            "success": False,
                            "error": f"找不到名為 '{layout_name}' 的layout"
                        }

                    # 使用CTAB切換layout
                    if hasattr(autocad_util, 'acad') and autocad_util.acad:
                        active_doc = autocad_util.acad.ActiveDocument
                        if active_doc:
                            active_doc.SetVariable("CTAB", layout_name)
                            time.sleep(0.1)  # 等待切換完成
                
                # 提取當前layout的參數（使用統一的 get_single_layout_values）
                log_message(f"[GUI Proxy] 開始提取layout '{layout_name}' 的參數")

                data = autocad_util.get_single_layout_values(layout_name)
                if not data:
                    return {"success": False, "error": f"layout '{layout_name}' 無資料"}

                log_message(f"[GUI Proxy] 成功提取layout '{layout_name}' 的完整參數")

                return {
                    "success": True,
                    "layout_name": layout_name,
                    "parameters": {k: v for k, v in data.items()
                                   if k not in ('layout_name', 'header_id', 'detail')},
                    "table_data": {
                        'header_id': data.get('header_id'),
                        'detail_list': data.get('detail', []),
                    },
                    "extraction_method": "single_layout"
                }
                
            finally:
                # 恢復到原始layout
                if original_name and original_name != layout_name:
                    try:
                        if hasattr(autocad_util, 'acad') and autocad_util.acad:
                            active_doc = autocad_util.acad.ActiveDocument
                            if active_doc:
                                active_doc.SetVariable("CTAB", original_name)
                                log_message(f"[GUI Proxy] 恢復到原始layout: {original_name}")
                    except Exception as restore_error:
                        log_message(f"[GUI Proxy] 恢復layout失敗: {restore_error}")
                
        except Exception as e:
            log_message(f"[GUI Proxy] 提取layout參數失敗: {e}")
            return {"success": False, "error": f"提取失敗: {str(e)}"}
    
    # 註冊所有處理器
    proxy.register_handler("switch_layout", handle_switch_layout)
    proxy.register_handler("get_current_layout", handle_get_current_layout)
    proxy.register_handler("extract_parameters", handle_extract_parameters)
    proxy.register_handler("extract_layout_parameters", handle_extract_layout_parameters)
    proxy.register_handler("get_autocad_status", handle_get_autocad_status)
    proxy.register_handler("draw_line", handle_draw_line)
    proxy.register_handler("draw_circle", handle_draw_circle)
    proxy.register_handler("export_layout_image", handle_export_layout_image)
    
    log_message("[GUI Proxy] 處理器設置完成")
    return proxy