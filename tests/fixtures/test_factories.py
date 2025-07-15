# -*- coding: utf-8 -*-
"""
Test data factories for creating consistent test data.
Following TDD best practices from CLAUDE.md
"""
from datetime import datetime
from typing import Dict, List, Any


def create_test_boq_entry(
    name: str = "Test Item", 
    quantity: float = 10.0, 
    unit: str = "pcs",
    description: str = "Test description",
    **kwargs
) -> Dict[str, Any]:
    """Create a test BOQ entry with default values"""
    entry = {
        'name': name,
        'quantity': quantity,
        'unit': unit,
        'description': description,
        'created_at': datetime.now().isoformat(),
        'updated_at': datetime.now().isoformat(),
        'project_id': 1,
        'status': 'active'
    }
    entry.update(kwargs)
    return entry


def create_test_autocad_entity(
    entity_type: str = "LINE",
    layer: str = "Default",
    **kwargs
) -> Dict[str, Any]:
    """Create a test AutoCAD entity with default values"""
    entity = {
        'type': entity_type,
        'layer': layer,
        'handle': f"test_handle_{entity_type.lower()}",
        'object_id': 12345,
        'color': 'ByLayer'
    }
    
    # Add type-specific properties
    if entity_type == "LINE":
        entity.update({
            'start_point': [0, 0, 0],
            'end_point': [100, 100, 0]
        })
    elif entity_type == "CIRCLE":
        entity.update({
            'center': [50, 50, 0],
            'radius': 25
        })
    elif entity_type == "TEXT":
        entity.update({
            'text_string': "Test Text",
            'insertion_point': [10, 10, 0],
            'height': 10
        })
    
    entity.update(kwargs)
    return entity


def create_test_odoo_product(
    product_id: int = 1,
    name: str = "Test Product",
    **kwargs
) -> Dict[str, Any]:
    """Create a test Odoo product with default values"""
    product = {
        'id': product_id,
        'name': name,
        'default_code': f"PROD{product_id:03d}",
        'list_price': 100.0,
        'standard_price': 80.0,
        'uom_name': 'Unit(s)',
        'categ_id': [1, 'All / Saleable'],
        'active': True
    }
    product.update(kwargs)
    return product


def create_test_mcp_request(
    method: str = "scan_entities",
    params: Dict = None,
    request_id: str = "test_request_1"
) -> Dict[str, Any]:
    """Create a test MCP request"""
    if params is None:
        params = {}
    
    return {
        'jsonrpc': '2.0',
        'method': method,
        'params': params,
        'id': request_id
    }


def create_test_mcp_response(
    result: Any = None,
    error: Dict = None,
    request_id: str = "test_request_1"
) -> Dict[str, Any]:
    """Create a test MCP response"""
    response = {
        'jsonrpc': '2.0',
        'id': request_id
    }
    
    if error:
        response['error'] = error
    else:
        response['result'] = result if result is not None else {'success': True}
    
    return response


def create_multiple_boq_entries(count: int = 5) -> List[Dict[str, Any]]:
    """Create multiple BOQ entries for testing"""
    return [
        create_test_boq_entry(
            name=f"Test Item {i}",
            quantity=i * 10,
            description=f"Test description {i}"
        )
        for i in range(1, count + 1)
    ]


def create_multiple_autocad_entities(count: int = 3) -> List[Dict[str, Any]]:
    """Create multiple AutoCAD entities for testing"""
    entity_types = ["LINE", "CIRCLE", "TEXT"]
    return [
        create_test_autocad_entity(
            entity_type=entity_types[i % len(entity_types)],
            layer=f"Layer_{i}"
        )
        for i in range(count)
    ]