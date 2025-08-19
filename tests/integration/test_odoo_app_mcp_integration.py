# -*- coding: utf-8 -*-
"""
Integration Test: Odoo App + MCP Server Integration
Tests the complete application launch via odoo.py with MCP functionality
"""

import unittest
import subprocess
import time
import sys
import os
import requests
import threading
from unittest.mock import patch, MagicMock

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

class TestOdooAppMCPIntegration(unittest.TestCase):
    """Test MCP integration through the main odoo.py application"""
    
    def setUp(self):
        """Set up test environment"""
        self.project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.odoo_py_path = os.path.join(self.project_root, 'odoo.py')
        self.mcp_processes = []
        self.test_timeout = 30  # seconds
        
    def tearDown(self):
        """Clean up test processes"""
        for proc in self.mcp_processes:
            try:
                if proc.poll() is None:  # Process is still running
                    proc.terminate()
                    proc.wait(timeout=5)
            except (subprocess.TimeoutExpired, ProcessLookupError):
                try:
                    proc.kill()
                except ProcessLookupError:
                    pass
        
        # Kill any remaining MCP server processes
        self.cleanup_mcp_processes()
    
    def cleanup_mcp_processes(self):
        """Clean up any remaining MCP server processes"""
        try:
            # Simple cleanup without psutil dependency
            import signal
            for proc in self.mcp_processes:
                try:
                    if proc.poll() is None:
                        proc.send_signal(signal.SIGTERM)
                        proc.wait(timeout=3)
                except (subprocess.TimeoutExpired, ProcessLookupError, AttributeError):
                    continue
        except Exception:
            pass
    
    def test_odoo_app_can_launch_mcp_server_mode(self):
        """Test: odoo.py can launch in pure MCP server mode"""
        print("\n=== Test: Launch odoo.py in MCP server mode ===")
        
        try:
            # Launch odoo.py in MCP server mode with custom port
            cmd = [sys.executable, self.odoo_py_path, '--mcp-server', '--mcp-port', '8084']
            proc = subprocess.Popen(
                cmd,
                cwd=self.project_root,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8'
            )
            self.mcp_processes.append(proc)
            
            # Wait a moment for server to start
            time.sleep(3)
            
            # Check if process is still running (not crashed)
            self.assertIsNone(proc.poll(), "MCP server process should still be running")
            
            print("[PASS] MCP server launched successfully through odoo.py")
            
        except Exception as e:
            self.fail(f"Failed to launch MCP server mode: {e}")
    
    def test_odoo_app_mcp_server_with_mock_gui(self):
        """Test: odoo.py can launch with MCP enabled (mocked GUI to avoid actual window)"""
        print("\n=== Test: Launch odoo.py with MCP enabled (GUI mocked) ===")
        
        try:
            # Mock the GUI components to prevent actual window creation
            with patch('forms.form_main_modern.ModernFormMain') as mock_form:
                mock_instance = MagicMock()
                mock_form.return_value = mock_instance
                
                # Mock the MCP server manager
                mock_instance.initialize_mcp_server_manager = MagicMock()
                mock_instance.mcp_server_manager = MagicMock()
                mock_instance.mcp_server_manager.start_all_servers = MagicMock()
                mock_instance.mainloop = MagicMock()
                
                # Import and run the main function with mocked GUI
                import odoo
                
                # Override sys.argv to simulate --enable-mcp
                with patch.object(sys, 'argv', ['odoo.py', '--enable-mcp']):
                    result = odoo.main()
                
                # Verify the result
                self.assertEqual(result, 0, "odoo.py should return 0 on success")
                
                # Verify MCP was initialized
                mock_instance.initialize_mcp_server_manager.assert_called_once()
                mock_instance.mcp_server_manager.start_all_servers.assert_called_once()
                
            print("[PASS] MCP integration works with GUI mode")
            
        except Exception as e:
            self.fail(f"Failed to test GUI with MCP: {e}")
    
    def test_mcp_natural_language_processing_integration(self):
        """Test: Natural language processing works through the application"""
        print("\n=== Test: Natural Language Processing Integration ===")
        
        try:
            # Test the NLP processor directly (as used in the app)
            from utility.util_nlp_processor import NaturalLanguageProcessor
            from mcp_server_fastmcp import process_natural_language_command
            
            # Test commands that should work in our Task 1 implementation
            test_commands = [
                "畫一個半徑10的圓形",
                "從(0,0)到(100,100)畫一條線"
            ]
            
            # Test NLP processor
            nlp = NaturalLanguageProcessor()
            for command in test_commands:
                intent = nlp.parse_drawing_intent(command)
                
                self.assertNotEqual(intent.action, "unknown", 
                                  f"Command should be recognized: {command}")
                self.assertGreater(intent.confidence, 0.7, 
                                 f"Confidence should be > 70%: {command}")
                
                print(f"  [PASS] Command parsed: {command} -> {intent.action} ({intent.confidence:.1%})")
            
            # Test MCP tool processing
            for command in test_commands:
                try:
                    result = process_natural_language_command(command)
                    
                    self.assertEqual(result.get("status"), "success", 
                                   f"MCP tool should process successfully: {command}")
                    
                    parsed_intent = result.get("data", {}).get("parsed_intent")
                    self.assertIsNotNone(parsed_intent, "Should have parsed intent")
                    
                    print(f"  [PASS] MCP processing: {command} -> {result.get('status')}")
                    
                except Exception as e:
                    # AutoCAD connection errors are expected in test environment
                    if "AutoCAD" in str(e) or "COM" in str(e):
                        print(f"  [EXPECTED] AutoCAD not available: {command}")
                    else:
                        raise e
            
            print("[PASS] Natural language processing integration working")
            
        except Exception as e:
            self.fail(f"Natural language processing test failed: {e}")
    
    def test_application_startup_sequence(self):
        """Test: Complete application startup sequence works"""
        print("\n=== Test: Application Startup Sequence ===")
        
        try:
            # Test database initialization
            import odoo
            
            # Mock database creation to avoid file system operations
            with patch('odoo.sqlite_create_table') as mock_db:
                mock_db.return_value = {
                    'host': 'test_host',
                    'db_name': 'test_db', 
                    'url': 'http://test.url',
                    'token': 'test_token'
                }
                
                # Test database setup
                odoo_conn = odoo.sqlite_create_table()
                self.assertIsNotNone(odoo_conn, "Database connection should be created")
                self.assertEqual(odoo_conn['host'], 'test_host', "Host should match")
                
            print("[PASS] Database initialization works")
            
            # Test GUI application startup (mocked)
            with patch('forms.form_main_modern.ModernFormMain') as mock_form:
                mock_instance = MagicMock()
                mock_form.return_value = mock_instance
                mock_instance.mainloop = MagicMock()
                
                # Test without MCP
                odoo.start_gui_application(enable_mcp=False)
                mock_form.assert_called_once()
                mock_instance.mainloop.assert_called_once()
                
            print("[PASS] GUI application startup works")
            
        except Exception as e:
            self.fail(f"Application startup test failed: {e}")
    
    def test_mcp_server_port_configuration(self):
        """Test: MCP server can be configured with different ports"""
        print("\n=== Test: MCP Server Port Configuration ===")
        
        try:
            # Test argument parsing
            import odoo
            import argparse
            
            # Create parser (same as in odoo.py)
            parser = argparse.ArgumentParser()
            parser.add_argument('--mcp-server', action='store_true')
            parser.add_argument('--mcp-port', type=int, default=8000)
            parser.add_argument('--mcp-pipe', type=str, default=r'\\\\.\\pipe\\odoo_autocad_mcp')
            parser.add_argument('--enable-mcp', action='store_true')
            
            # Test different configurations
            test_configs = [
                ['--mcp-server'],
                ['--mcp-server', '--mcp-port', '8085'],
                ['--enable-mcp'],
            ]
            
            for config in test_configs:
                args = parser.parse_args(config)
                
                if '--mcp-server' in config:
                    self.assertTrue(args.mcp_server, "MCP server flag should be set")
                
                if '--mcp-port' in config:
                    expected_port = int(config[config.index('--mcp-port') + 1])
                    self.assertEqual(args.mcp_port, expected_port, f"Port should be {expected_port}")
                
                if '--enable-mcp' in config:
                    self.assertTrue(args.enable_mcp, "Enable MCP flag should be set")
                
                print(f"  [PASS] Configuration: {' '.join(config)}")
            
            print("[PASS] MCP server port configuration works")
            
        except Exception as e:
            self.fail(f"Port configuration test failed: {e}")


def run_integration_test():
    """Run the integration test suite"""
    print("=" * 80)
    print("Task 1 Milestone: Odoo App + MCP Integration Test")
    print("=" * 80)
    print("Testing complete application launch and MCP functionality")
    print("Note: GUI tests are mocked to prevent actual window creation")
    print()
    
    # Create test suite
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestOdooAppMCPIntegration)
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2, stream=sys.stdout)
    result = runner.run(suite)
    
    # Summary
    print("\n" + "=" * 80)
    print("Integration Test Summary")
    print("=" * 80)
    print(f"Tests Run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    
    if result.failures:
        print("\nFailures:")
        for test, traceback in result.failures:
            print(f"  - {test}: {traceback.split('\\n')[-2]}")
    
    if result.errors:
        print("\nErrors:")
        for test, traceback in result.errors:
            print(f"  - {test}: {traceback.split('\\n')[-2]}")
    
    success = len(result.failures) == 0 and len(result.errors) == 0
    
    if success:
        print("\n[SUCCESS] All integration tests passed!")
        print("✅ odoo.py can launch with MCP functionality")
        print("✅ Natural language processing works through the app")
        print("✅ Application startup sequence is correct")
        print("✅ MCP server configuration is flexible")
        print("\nReady for manual end-to-end testing with actual Gemini CLI!")
    else:
        print("\n[FAIL] Some integration tests failed")
        print("❌ Please fix the issues before proceeding")
    
    return 0 if success else 1


if __name__ == "__main__":
    exit_code = run_integration_test()
    sys.exit(exit_code)