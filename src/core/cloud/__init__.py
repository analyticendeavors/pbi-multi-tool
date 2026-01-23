"""
Core Cloud Module - Centralized cloud browser components
Built by Reid Havens of Analytic Endeavors

Provides reusable components for browsing Power BI Service workspaces
and datasets via the Fabric API.
"""

# Import models first (no dependencies on other modules)
from core.cloud.models import (
    CloudConnectionType,
    WorkspaceInfo,
    DatasetInfo,
    SwapTarget,
)

from core.cloud.cloud_workspace_browser import (
    CloudWorkspaceBrowser,
    _get_shared_msal_app,
    _save_token_cache,
)
from core.cloud.cloud_browser_dialog import CloudBrowserDialog

__all__ = [
    # Models
    'CloudConnectionType',
    'WorkspaceInfo',
    'DatasetInfo',
    'SwapTarget',
    # Browser components
    'CloudWorkspaceBrowser',
    'CloudBrowserDialog',
    '_get_shared_msal_app',
    '_save_token_cache',
]
