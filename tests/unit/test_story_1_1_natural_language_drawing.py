# -*- coding: utf-8 -*-
"""
Unit tests for Story 1.1: Natural Language Drawing Commands
Tests the NLP processor and MCP tool integration
"""

import unittest
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from utility.util_nlp_processor import NaturalLanguageProcessor, DrawingIntent

class TestNaturalLanguageProcessor(unittest.TestCase):
    """Test the Natural Language Processor for CAD commands"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.nlp = NaturalLanguageProcessor()
    
    def test_circle_command_with_radius(self):
        """Test parsing circle command with radius"""
        command = "畫一個半徑10的圓形"
        intent = self.nlp.parse_drawing_intent(command)
        
        self.assertEqual(intent.action, "draw_circle")
        self.assertEqual(intent.parameters["radius"], 10.0)
        self.assertEqual(intent.parameters["center_point"], [0.0, 0.0, 0.0])
        self.assertGreater(intent.confidence, 0.8)
    
    def test_circle_command_with_diameter(self):
        """Test parsing circle command with diameter"""
        command = "畫一個直徑20的圓形"
        intent = self.nlp.parse_drawing_intent(command)
        
        self.assertEqual(intent.action, "draw_circle")
        self.assertEqual(intent.parameters["radius"], 10.0)  # diameter/2
        self.assertGreater(intent.confidence, 0.8)
    
    def test_line_command_with_coordinates(self):
        """Test parsing line command with coordinates"""
        command = "從(0,0)到(100,100)畫一條線"
        intent = self.nlp.parse_drawing_intent(command)
        
        self.assertEqual(intent.action, "draw_line")
        self.assertEqual(intent.parameters["start_point"], [0.0, 0.0, 0.0])
        self.assertEqual(intent.parameters["end_point"], [100.0, 100.0, 0.0])
        self.assertGreater(intent.confidence, 0.8)
    
    def test_text_command_with_content(self):
        """Test parsing text command with content"""
        command = '在(0,0)寫文字"Hello World"'
        intent = self.nlp.parse_drawing_intent(command)
        
        self.assertEqual(intent.action, "create_text")
        self.assertEqual(intent.parameters["text_content"], "Hello World")
        self.assertEqual(intent.parameters["position"], [0.0, 0.0, 0.0])
        self.assertGreater(intent.confidence, 0.7)
    
    def test_unknown_command(self):
        """Test handling unknown commands"""
        command = "做一些不相關的事情"
        intent = self.nlp.parse_drawing_intent(command)
        
        self.assertEqual(intent.action, "unknown")
        self.assertEqual(intent.confidence, 0.0)
        self.assertGreater(len(intent.suggestions), 0)
    
    def test_accuracy_target(self):
        """Test accuracy against target commands"""
        test_commands = [
            ("畫一個半徑10的圓形", "draw_circle"),
            ("從(0,0)到(100,100)畫線", "draw_line"),
            ("圓形半徑15", "draw_circle"),
            ("線段從原點到100,50", "draw_line"),
            ('寫文字"測試"', "create_text")
        ]
        
        correct = 0
        total = len(test_commands)
        
        for command, expected_action in test_commands:
            intent = self.nlp.parse_drawing_intent(command)
            if intent.action == expected_action and intent.confidence > 0.7:
                correct += 1
        
        accuracy = correct / total
        self.assertGreater(accuracy, 0.8, f"Accuracy {accuracy:.1%} below 80% target")
    
    def test_performance_target(self):
        """Test that processing time is under performance target"""
        import time
        
        command = "畫一個半徑10的圓形"
        
        start_time = time.time()
        intent = self.nlp.parse_drawing_intent(command)
        end_time = time.time()
        
        processing_time = end_time - start_time
        
        # Should process in under 0.5 seconds for simple commands
        self.assertLess(processing_time, 0.5, 
                       f"Processing time {processing_time:.3f}s exceeds 0.5s target")

class TestMCPToolIntegration(unittest.TestCase):
    """Test MCP tool integration (mock tests)"""
    
    def test_process_natural_language_command_structure(self):
        """Test that the MCP tool has the correct structure"""
        # Import the function
        try:
            from mcp_server_fastmcp import process_natural_language_command
            
            # Test function exists and is callable
            self.assertTrue(callable(process_natural_language_command))
            
            # Test function signature
            import inspect
            sig = inspect.signature(process_natural_language_command)
            params = list(sig.parameters.keys())
            
            self.assertIn('command', params)
            self.assertIn('user_context', params)
            
        except ImportError:
            self.fail("process_natural_language_command not properly imported")

if __name__ == '__main__':
    # Create test directory if it doesn't exist
    os.makedirs('logs', exist_ok=True)
    
    # Run tests
    unittest.main(verbosity=2)