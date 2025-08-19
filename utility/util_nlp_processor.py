# -*- coding: utf-8 -*-
"""
Simple Natural Language Processor for CAD Drawing Commands
Task 1 Implementation - Basic keyword parsing for early testing
"""

import re
import logging
from typing import Dict, Any, List, Optional
from dataclasses import dataclass

logger = logging.getLogger(__name__)

@dataclass
class DrawingIntent:
    """基礎繪圖意圖資料模型"""
    action: str  # "draw_circle", "draw_line", "create_text"
    parameters: Dict[str, Any]
    confidence: float
    suggestions: List[str]
    
    def __post_init__(self):
        """確保必要的參數存在"""
        if not hasattr(self, 'suggestions'):
            self.suggestions = []

class NaturalLanguageProcessor:
    """簡化版自然語言處理器 - Task 1 基礎關鍵字解析"""
    
    def __init__(self):
        """初始化處理器"""
        self.patterns = {
            'circle': {
                'keywords': ['圓形', '圓', 'circle'],
                'radius_patterns': [
                    r'半徑\s*(\d+(?:\.\d+)?)',
                    r'直徑\s*(\d+(?:\.\d+)?)',
                    r'radius\s*(\d+(?:\.\d+)?)',
                    r'diameter\s*(\d+(?:\.\d+)?)'
                ],
                'center_patterns': [
                    r'中心\s*\(?([^)]*)\)?',
                    r'center\s*\(?([^)]*)\)?'
                ]
            },
            'line': {
                'keywords': ['線段', '線', '直線', 'line'],
                'point_patterns': [
                    r'從\s*\(?([^)]*)\)?\s*到\s*\(?([^)]*)\)?',
                    r'from\s*\(?([^)]*)\)?\s*to\s*\(?([^)]*)\)?',
                    r'\(?([^)]*)\)?\s*到\s*\(?([^)]*)\)?'
                ]
            },
            'text': {
                'keywords': ['文字', '文本', 'text'],
                'content_patterns': [
                    r'"([^"]*)"',
                    r"'([^']*)'",
                    r'內容\s*[：:]\s*(.+)',
                    r'content\s*[：:]\s*(.+)'
                ]
            }
        }
    
    def parse_drawing_intent(self, command: str) -> DrawingIntent:
        """解析自然語言繪圖指令"""
        command = command.strip()
        logger.info(f"Parsing command: {command}")
        
        # 檢測圓形指令
        if self._contains_keywords(command, self.patterns['circle']['keywords']):
            return self._parse_circle_command(command)
        
        # 檢測線段指令
        elif self._contains_keywords(command, self.patterns['line']['keywords']):
            return self._parse_line_command(command)
        
        # 檢測文字指令
        elif self._contains_keywords(command, self.patterns['text']['keywords']):
            return self._parse_text_command(command)
        
        else:
            # 無法識別的指令
            return DrawingIntent(
                action="unknown",
                parameters={},
                confidence=0.0,
                suggestions=[
                    "請嘗試：畫一個半徑10的圓形",
                    "請嘗試：從(0,0)到(100,100)畫一條線",
                    "請嘗試：在(0,0)寫文字\"Hello\""
                ]
            )
    
    def _contains_keywords(self, command: str, keywords: List[str]) -> bool:
        """檢查指令是否包含關鍵字"""
        for keyword in keywords:
            if keyword in command:
                return True
        return False
    
    def _parse_circle_command(self, command: str) -> DrawingIntent:
        """解析圓形繪製指令"""
        parameters = {}
        confidence = 0.7  # 基礎置信度
        suggestions = []
        
        # 解析半徑或直徑
        radius = None
        for pattern in self.patterns['circle']['radius_patterns']:
            match = re.search(pattern, command)
            if match:
                value = float(match.group(1))
                if '直徑' in pattern or 'diameter' in pattern:
                    radius = value / 2.0  # 直徑轉半徑
                else:
                    radius = value
                confidence = 0.9
                break
        
        if radius is None:
            radius = 10.0  # 預設半徑
            suggestions.append("未指定半徑，使用預設值10")
        
        # 解析中心點
        center = [0.0, 0.0, 0.0]  # 預設中心點
        for pattern in self.patterns['circle']['center_patterns']:
            match = re.search(pattern, command)
            if match:
                center_str = match.group(1)
                try:
                    # 嘗試解析座標 (x,y) 或 x,y
                    coords = re.findall(r'-?\d+(?:\.\d+)?', center_str)
                    if len(coords) >= 2:
                        center = [float(coords[0]), float(coords[1]), 0.0]
                        confidence = 0.95
                except:
                    suggestions.append("中心點格式不正確，使用預設值(0,0)")
        
        parameters = {
            "center_point": center,
            "radius": radius,
            "layer": "0"  # 預設圖層
        }
        
        return DrawingIntent(
            action="draw_circle",
            parameters=parameters,
            confidence=confidence,
            suggestions=suggestions
        )
    
    def _parse_line_command(self, command: str) -> DrawingIntent:
        """解析線段繪製指令"""
        parameters = {}
        confidence = 0.7
        suggestions = []
        
        # 解析起點和終點
        start_point = [0.0, 0.0, 0.0]
        end_point = [100.0, 100.0, 0.0]
        
        for pattern in self.patterns['line']['point_patterns']:
            match = re.search(pattern, command)
            if match:
                try:
                    # 解析起點
                    start_coords = re.findall(r'-?\d+(?:\.\d+)?', match.group(1))
                    if len(start_coords) >= 2:
                        start_point = [float(start_coords[0]), float(start_coords[1]), 0.0]
                    
                    # 解析終點
                    end_coords = re.findall(r'-?\d+(?:\.\d+)?', match.group(2))
                    if len(end_coords) >= 2:
                        end_point = [float(end_coords[0]), float(end_coords[1]), 0.0]
                    
                    confidence = 0.95
                except Exception as e:
                    suggestions.append(f"座標解析錯誤: {str(e)}")
                break
        
        if confidence < 0.9:
            suggestions.append("未找到明確座標，使用預設值")
        
        parameters = {
            "start_point": start_point,
            "end_point": end_point,
            "layer": "0"
        }
        
        return DrawingIntent(
            action="draw_line",
            parameters=parameters,
            confidence=confidence,
            suggestions=suggestions
        )
    
    def _parse_text_command(self, command: str) -> DrawingIntent:
        """解析文字創建指令"""
        parameters = {}
        confidence = 0.7
        suggestions = []
        
        # 解析文字內容
        text_content = "預設文字"
        for pattern in self.patterns['text']['content_patterns']:
            match = re.search(pattern, command)
            if match:
                text_content = match.group(1)
                confidence = 0.9
                break
        
        if confidence < 0.9:
            suggestions.append("未找到明確文字內容，使用預設值")
        
        # 預設位置
        position = [0.0, 0.0, 0.0]
        
        parameters = {
            "position": position,
            "text_content": text_content,
            "height": 2.5,
            "layer": "0"
        }
        
        return DrawingIntent(
            action="create_text",
            parameters=parameters,
            confidence=confidence,
            suggestions=suggestions
        )