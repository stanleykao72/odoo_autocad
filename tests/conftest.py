# -*- coding: utf-8 -*-
"""
pytest configuration and shared fixtures for the entire test suite.
Following TDD principles from CLAUDE.md
"""
import pytest
import sys
import os
from unittest.mock import Mock, patch, MagicMock

# Add project root to Python path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture
def mock_autocad_util():
    """Mock AutoCAD utility for testing without actual AutoCAD dependency"""
    mock_util = Mock()
    mock_util.connected_autocad.return_value = True
    mock_util.connect_autocad.return_value = True
    mock_util.acad = Mock()
    mock_util.doc = Mock()
    return mock_util


@pytest.fixture
def mock_odoo_util():
    """Mock Odoo utility for testing without actual Odoo dependency"""
    mock_util = Mock()
    mock_util.connected_odoo.return_value = True
    mock_util.odoo = Mock()
    mock_util.user_token = "test_token"
    return mock_util


@pytest.fixture
def mock_log_util():
    """Mock logging utility for testing"""
    mock_log = Mock()
    mock_log.log.return_value = None
    mock_log.safe_log_insert.return_value = None
    return mock_log


@pytest.fixture
def mock_win32com():
    """Mock win32com.client for testing COM operations"""
    with patch('win32com.client') as mock:
        mock_app = Mock()
        mock_app.Visible = True
        mock.GetActiveObject.return_value = mock_app
        mock.Dispatch.return_value = mock_app
        yield mock


@pytest.fixture
def sample_boq_data():
    """Sample BOQ data for testing"""
    return [
        {
            'name': 'Test Item 1',
            'quantity': 10,
            'unit': 'pcs',
            'description': 'Test description 1'
        },
        {
            'name': 'Test Item 2', 
            'quantity': 5,
            'unit': 'meter',
            'description': 'Test description 2'
        }
    ]


@pytest.fixture
def sample_autocad_entities():
    """Sample AutoCAD entities data for testing"""
    return [
        {
            'type': 'LINE',
            'start_point': [0, 0, 0],
            'end_point': [100, 100, 0],
            'layer': 'Default'
        },
        {
            'type': 'CIRCLE',
            'center': [50, 50, 0],
            'radius': 25,
            'layer': 'Dimensions'
        }
    ]


@pytest.fixture
def mcp_test_config():
    """Test configuration for MCP server"""
    return {
        'tcp_port': 8000,
        'pipe_name': r'\\.\pipe\test_odoo_autocad_mcp',
        'timeout': 30,
        'max_connections': 5
    }


# Test markers for categorizing tests
def pytest_configure(config):
    """Configure test markers"""
    config.addinivalue_line("markers", "slow: marks tests as slow (deselect with '-m \"not slow\"')")
    config.addinivalue_line("markers", "autocad: marks tests requiring live AutoCAD COM connection")
    config.addinivalue_line("markers", "ipc: marks tests requiring live AutoCAD IPC connection")
    config.addinivalue_line("markers", "odoo: marks tests as Odoo-related")
    config.addinivalue_line("markers", "mcp: marks tests as MCP-related")
    config.addinivalue_line("markers", "integration: marks tests as integration tests")
    config.addinivalue_line("markers", "unit: marks tests as unit tests")