# -*- coding: utf-8 -*-
"""
Tests for the 5 Odoo MCP tools in mcp_server_autocad.py.

迴歸測試重點：
- odoo_get_colors 之前呼叫 odoo.get_color() 缺少必填的 project_id → 必定 TypeError
- odoo_get_setup 之前把 int project_id 傳給期待 setup_name 的 API
- Odoo API 以字串回傳錯誤訊息時要轉成 success=False
"""
from unittest.mock import MagicMock

import pytest

mcp_server_autocad = pytest.importorskip("mcp_server_autocad")


@pytest.fixture
def odoo_mock(monkeypatch):
    odoo = MagicMock()
    monkeypatch.setattr(mcp_server_autocad, "_odoo_util", odoo, raising=False)
    return odoo


@pytest.fixture
def dispatcher_mock(monkeypatch):
    disp = MagicMock()
    monkeypatch.setattr(mcp_server_autocad, "_autocad_dispatcher", disp, raising=False)
    return disp


def _fn(tool_name):
    """取出工具函式（@mcp.tool() 會原樣回傳被裝飾的函式）"""
    return getattr(mcp_server_autocad, tool_name)


def _schema(tool_name):
    """取出 FastMCP 為該工具產生的 JSON Schema"""
    return mcp_server_autocad.mcp._tool_manager.get_tool(tool_name).parameters


class TestOdooGetColors:
    def test_passes_project_id_to_odoo(self, odoo_mock):
        odoo_mock.get_color.return_value = [{"id": 1, "name": "紅"}]

        result = _fn("odoo_get_colors")(project_id=42)

        odoo_mock.get_color.assert_called_once_with(42)
        assert result["success"] is True
        assert result["project_id"] == 42
        assert result["colors"][0]["name"] == "紅"

    def test_falls_back_to_drawing_project_id(self, odoo_mock, dispatcher_mock):
        dispatcher_mock.project_id = 7
        odoo_mock.get_color.return_value = []

        result = _fn("odoo_get_colors")()

        odoo_mock.get_color.assert_called_once_with(7)
        assert result["success"] is True

    def test_errors_when_no_project_id_available(self, odoo_mock, dispatcher_mock):
        dispatcher_mock.project_id = None

        result = _fn("odoo_get_colors")()

        assert result["success"] is False
        odoo_mock.get_color.assert_not_called()

    def test_odoo_error_string_becomes_failure(self, odoo_mock):
        odoo_mock.get_color.return_value = "權限不足"

        result = _fn("odoo_get_colors")(project_id=1)

        assert result["success"] is False
        assert result["error"] == "權限不足"


class TestOdooGetSetup:
    def test_passes_setup_name_through(self, odoo_mock):
        odoo_mock.get_setup.return_value = [{"id": 1}]

        result = _fn("odoo_get_setup")(setup_name="product_category")

        odoo_mock.get_setup.assert_called_once_with("product_category")
        assert result["success"] is True

    def test_odoo_error_string_becomes_failure(self, odoo_mock):
        odoo_mock.get_setup.return_value = "找不到設定"

        result = _fn("odoo_get_setup")(setup_name="nope")

        assert result["success"] is False
        assert result["error"] == "找不到設定"


class TestToolSchemas:
    def test_get_colors_project_id_is_optional(self):
        schema = _schema("odoo_get_colors")
        assert "project_id" in schema["properties"]
        assert "project_id" not in schema.get("required", [])

    def test_get_setup_requires_setup_name(self):
        schema = _schema("odoo_get_setup")
        assert schema["properties"]["setup_name"]["type"] == "string"
        assert "setup_name" in schema.get("required", [])
