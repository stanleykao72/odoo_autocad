# -*- coding: utf-8 -*-


class UtilTransferBoqToPr:
    def __init__(self, odoo_util, autocad_util, log_util):
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.log_util = log_util

    def transfer_boq_to_pr(self, progress_callback=None):
        """轉移 BOQ 到 PR — 支援進度回調

        Args:
            progress_callback: optional callable(value: float, message: str)
        Returns:
            True on success, False on failure
        """
        def _log(msg):
            self.log_util.safe_log_insert(msg)

        def _progress(value, message):
            _log(f"[PR] ({value:.0%}) {message}\n")
            if progress_callback:
                progress_callback(value, message)

        _log("[PR] ========== 開始轉移 BOQ 到 PR ==========\n")

        _progress(0.10, "取得 header IDs...")
        _log("[PR] 呼叫 get_layouts_header_id_to_pr()...\n")
        header_id_dict = self.autocad_util.get_layouts_header_id_to_pr()
        _log(f"[PR] header_id_dict: {header_id_dict}\n")

        _progress(0.30, "檢查資料...")
        if not header_id_dict:
            _log("[PR] ✘ 無法取得 header IDs，中止轉移\n")
            _progress(1.0, "無資料，已中止")
            return False

        _log(f"[PR] 取得 header IDs 成功，準備呼叫 boq2pr()\n")
        _progress(0.40, "轉移 BOQ 到 PR...")
        _log("[PR] 呼叫 boq2pr()...\n")
        pr_list = self.odoo_util.boq2pr(header_id_dict)
        _log(f"[PR] boq2pr() 回傳: {str(pr_list)[:300]}\n")

        _progress(1.0, "轉移 BOQ 到 PR 完成")
        _log("[PR] ✔ 轉移 BOQ 到 PR 完成\n")
        _log("[PR] ========================================\n")
        return True
