"""
Connection Hot-Swap UI Dialogs

Modal dialog components for cloud/local model selection.
"""

# CloudBrowserDialog moved to core.cloud for reuse across tools
from core.cloud import CloudBrowserDialog
from tools.connection_hotswap.ui.dialogs.local_selector_dialog import LocalSelectorDialog

__all__ = ['CloudBrowserDialog', 'LocalSelectorDialog']
