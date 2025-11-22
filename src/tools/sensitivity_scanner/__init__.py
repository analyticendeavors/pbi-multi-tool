"""
Sensitivity Scanner Tool for Power BI PBIP Files

This tool scans Power BI semantic models (TMDL files) for sensitive content including:
- Personally Identifiable Information (PII)
- Financial data
- Credentials and security information
- Infrastructure details
- Row-Level Security expressions

Built by Reid Havens of Analytic Endeavors
"""

from .sensitivity_scanner_tool import SensitivityScannerTool

__all__ = ['SensitivityScannerTool']
__version__ = '1.0.0'
