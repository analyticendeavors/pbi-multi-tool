"""
Field Parameters Data Models
Data classes representing Power BI model structures and connection state.

Built by Reid Havens of Analytic Endeavors

NOTE: These classes are now defined in core/pbi_connector.py to avoid circular imports.
This module re-exports them for backward compatibility.
"""

# Re-export from the canonical location
from core.pbi_connector import (
    ModelConnection,
    AvailableModel,
    TableInfo,
    FieldInfo,
    TableFieldsInfo,
)

__all__ = [
    'ModelConnection',
    'AvailableModel',
    'TableInfo',
    'FieldInfo',
    'TableFieldsInfo',
]
