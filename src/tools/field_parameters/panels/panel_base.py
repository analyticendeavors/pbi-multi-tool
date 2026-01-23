"""
Panel Base Mixin
Shared functionality for Field Parameters tool panels.

Built by Reid Havens of Analytic Endeavors
"""

# Re-export SectionPanelMixin from core.ui_base for backward compatibility
# The actual implementation now lives in core/ui_base.py
from core.ui_base import SectionPanelMixin

__all__ = ['SectionPanelMixin']
