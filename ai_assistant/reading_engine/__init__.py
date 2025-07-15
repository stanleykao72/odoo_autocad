# -*- coding: utf-8 -*-
"""
Reading Engine for AutoCAD entity scanning and data extraction

Based on Easy-MCP-AutoCad capabilities, this engine provides:
- Entity scanning and classification
- BOQ data extraction
- Drawing analysis and validation
"""

from .autocad_manager import AutoCADManager
from .entity_scanner import EntityScanner

__all__ = ['AutoCADManager', 'EntityScanner']