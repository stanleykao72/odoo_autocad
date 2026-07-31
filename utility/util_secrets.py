# -*- coding: utf-8 -*-
"""憑證遮罩工具 — 避免 token 等機密以明文寫入日誌檔。"""


def mask_secret(value, keep=4):
    """遮罩機密字串，僅保留開頭數個字元。

    Args:
        value: 要遮罩的值（None / 空值直接回傳描述字串）
        keep: 保留的開頭字元數
    Returns:
        str — 例如 "abcd***(len=40)"
    """
    if not value:
        return "(未設定)"
    text = str(value)
    if len(text) <= keep:
        return "*" * len(text)
    return f"{text[:keep]}***(len={len(text)})"


def mask_url_token(url):
    """遮罩 URL query string 中的 token / password / api_key 等參數值。

    Odoo 的 swagger url 形如
        https://host/api/v1/...swagger.json?token=<uuid>&db=<name>
    直接記錄整段 URL 等同把 API token 寫進日誌。

    Args:
        url: 原始 URL
    Returns:
        str — token 值被遮罩後的 URL
    """
    if not url:
        return "(未設定)"
    import re
    return re.sub(
        r"((?:token|password|passwd|pwd|api_key|apikey|secret)=)([^&#]*)",
        lambda m: m.group(1) + mask_secret(m.group(2)),
        str(url),
        flags=re.IGNORECASE,
    )
