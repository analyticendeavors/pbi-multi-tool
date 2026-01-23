"""
AE Multi-Tool - Core Components
Built by Reid Havens of Analytic Endeavors

Copyright (c) 2024 Analytic Endeavors LLC. All rights reserved.
Unauthorized copying, modification, or distribution is prohibited.
"""

__version__ = "1.0.0"
__author__ = "Reid Havens"
__company__ = "Analytic Endeavors"
__fingerprint__ = "ae7f3c2d8b4e9a1f6c5d0e7b2a3f8c9d4e5b6a7c8d9e0f1a2b3c4d5e6f7a8b9c"

from core.ownership import verify_ownership, get_fingerprint, get_build_info
from core.widgets import AutoHideScrollbar, LoadingOverlay, SearchableTreeview, ToastNotification
from core.ui_base import RecentFilesManager, get_recent_files_manager

__all__ = [
    'verify_ownership',
    'get_fingerprint',
    'get_build_info',
    'AutoHideScrollbar',
    'LoadingOverlay',
    'SearchableTreeview',
    'ToastNotification',
    'RecentFilesManager',
    'get_recent_files_manager',
]
