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
