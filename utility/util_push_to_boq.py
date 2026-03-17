# -*- coding: utf-8 -*-


class UtilPushToBoq:
    def __init__(self, odoo_util, autocad_util, log_util):
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.log_util = log_util

    def push_to_boq(self, progress_callback=None):
        """推送到 BOQ — 支援 per-layout IPC 和進度回調

        Args:
            progress_callback: optional callable(value: float, message: str)
                               value 0.0~1.0, message 為顯示文字
        Returns:
            True on success, False on failure/cancel
        """
        def _log(msg):
            self.log_util.safe_log_insert(msg)

        def _progress(value, message):
            _log(f"[BOQ] ({value:.0%}) {message}\n")
            if progress_callback:
                progress_callback(value, message)

        _log("[BOQ] ========== 開始推送到 BOQ ==========\n")
        _progress(0.02, "取得 layout 清單...")

        # Per-layout extraction only works in IPC mode (COM uses bulk call)
        is_ipc = hasattr(self.autocad_util, 'mode') and self.autocad_util.mode == 'ipc'
        use_per_layout = is_ipc and hasattr(self.autocad_util, 'get_single_layout_values')
        _log(f"[BOQ] per-layout 支援: {use_per_layout} (mode={getattr(self.autocad_util, 'mode', 'unknown')})\n")
        layout_dic = []

        if use_per_layout:
            # Step 1: Get layout list
            _log("[BOQ] 呼叫 get_doc_layouts()...\n")
            layouts = self.autocad_util.get_doc_layouts()
            if layouts is None:
                layouts = []
            _log(f"[BOQ] get_doc_layouts() 回傳 {len(layouts)} 個 layout: {layouts}\n")

            if layouts:
                _progress(0.05, f"找到 {len(layouts)} 個 layout")

                # Step 2: Per-layout extraction
                for idx, name in enumerate(layouts):
                    # Check cancellation
                    if progress_callback and getattr(progress_callback, '_cancelled', False):
                        _log("[BOQ] ⚠ 使用者取消操作\n")
                        return False

                    frac = 0.05 + (idx / len(layouts)) * 0.65  # 5% ~ 70%
                    _progress(frac, f"提取 layout: {name} ({idx+1}/{len(layouts)})")

                    _log(f"[BOQ] 呼叫 get_single_layout_values('{name}')...\n")
                    data = self.autocad_util.get_single_layout_values(name)
                    if isinstance(data, dict):
                        _log(f"[BOQ] [{name}] 回傳 {len(data)} 個 key: {list(data.keys())}\n")
                    else:
                        _log(f"[BOQ] [{name}] 回傳型別: {type(data).__name__}, 值: {str(data)[:200]}\n")
                    if data:
                        # Inject pr_no if missing (AutoLISP pr_no text block
                        # may not be found during per-layout iteration)
                        if isinstance(data, dict) and 'pr_no' not in data:
                            pr_no = getattr(self.autocad_util, 'pr_no', None)
                            if pr_no:
                                data['pr_no'] = pr_no
                                _log(f"[BOQ] [{name}] 補充 pr_no={pr_no}\n")
                        layout_dic.append(data)
                    else:
                        _log(f"[BOQ] [{name}] ⚠ 空資料，跳過\n")

                _progress(0.70, f"提取完成，共 {len(layout_dic)}/{len(layouts)} 個有效 layout")
            else:
                # Fallback: odoo_get_layouts not available yet, use bulk call
                _log("[BOQ] ⚠ get_doc_layouts() 回傳空，可能 odoo_get_layouts 尚未載入 → 改用批次提取\n")
                use_per_layout = False

        if not use_per_layout:
            # Fallback: single bulk call (COM mode or old AutoLISP)
            _progress(0.10, "提取所有 layout 資料（批次）...")
            _log("[BOQ] 呼叫 get_layouts_values() (批次 odoo_extract_tables)...\n")
            layout_dic = self.autocad_util.get_layouts_values()
            _log(f"[BOQ] get_layouts_values() 回傳 {len(layout_dic) if isinstance(layout_dic, list) else type(layout_dic).__name__}\n")
            _progress(0.70, "提取完成")

        if not layout_dic:
            _log("[BOQ] ✘ 無法取得 layout 資料（可能 timeout 或無 TABLE），中止推送\n")
            _progress(1.0, "無資料，已中止")
            return False

        _log(f"[BOQ] 共 {len(layout_dic)} 筆 layout 資料待推送\n")

        # Step 3: Push to Odoo
        # Wrap in {"all": [...]} to match COM mode format expected by Odoo API
        if isinstance(layout_dic, list):
            layout_dic = {"all": layout_dic}
        _progress(0.75, "推送資料到 Odoo...")
        _log(f"[BOQ] 呼叫 import2boq(), keys={list(layout_dic.keys()) if isinstance(layout_dic, dict) else type(layout_dic).__name__}\n")
        boq_list = self.odoo_util.import2boq(layout_dic)
        _log(f"[BOQ] import2boq() 回傳: {str(boq_list)[:300]}\n")

        # Check if Odoo returned an error message instead of boq_list
        if isinstance(boq_list, str):
            _log(f"[BOQ] ✘ Odoo 回傳錯誤: {boq_list}\n")
            _progress(1.0, boq_list)
            raise RuntimeError(boq_list)

        if not boq_list:
            _log("[BOQ] ✘ import2boq() 回傳空值\n")
            _progress(1.0, "Odoo 回傳空值")
            return False

        # Step 4: Write IDs back
        _progress(0.90, "寫回 ID 到 AutoCAD...")
        _log("[BOQ] 呼叫 set_layouts_tables_id()...\n")
        self.autocad_util.set_layouts_tables_id(boq_list)
        _log("[BOQ] set_layouts_tables_id() 完成\n")

        _progress(1.0, "推送到 BOQ 完成")
        _log("[BOQ] ✔ 推送到 BOQ 完成\n")
        _log("[BOQ] ====================================\n")
        return True
