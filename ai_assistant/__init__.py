# -*- coding: utf-8 -*-
"""
AI Assistant module for AutoCAD-Odoo Integration

This module provides MCP (Model Context Protocol) server functionality
to enable AI-driven AutoCAD and Odoo operations.

Architecture:
- mcp_server_manager: Core MCP server with TCP Socket + Named Pipe support
- reading_engine: AutoCAD entity scanning and data extraction
- drawing_engine: AutoCAD drawing and modification tools
- integration_bridge: Bridge between AutoCAD, Odoo, and MCP protocols
"""

__version__ = "1.0.0"
__author__ = "AutoCAD-Odoo Integration Team"

from .mcp_server_manager import MCPServerManager

__all__ = ['MCPServerManager']