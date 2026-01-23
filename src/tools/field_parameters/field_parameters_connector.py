"""
DEPRECATED: This module has been moved to core/pbi_connector.py

This file exists only for backward compatibility.
Please update imports to use: from core.pbi_connector import get_connector
"""

# Re-export everything from the new location for backward compatibility
from core.pbi_connector import (
    ModelConnection,
    AvailableModel,
    TableInfo,
    FieldInfo,
    TableFieldsInfo,
    PowerBIConnector,
    get_connector,
)

__all__ = [
    'ModelConnection',
    'AvailableModel',
    'TableInfo',
    'FieldInfo',
    'TableFieldsInfo',
    'PowerBIConnector',
    'get_connector',
]

import warnings
warnings.warn(
    "field_parameters_connector is deprecated. Use 'from core.pbi_connector import get_connector' instead.",
    DeprecationWarning,
    stacklevel=2
)
