"""
Advanced Copy Logic Module
Built by Reid Havens of Analytic Endeavors

This module contains all the core logic for the Advanced Copy tool.
All logic components are organized here for clean separation from UI.
"""

from .advanced_copy_core import AdvancedCopyEngine, AdvancedCopyError
from .advanced_copy_operations import PageOperations
from .advanced_copy_bookmark_analyzer import BookmarkAnalyzer
from .advanced_copy_bookmark_groups import BookmarkGroupManager
from .advanced_copy_visual_actions import VisualActionUpdater

__all__ = [
    'AdvancedCopyEngine',
    'AdvancedCopyError',
    'PageOperations',
    'BookmarkAnalyzer',
    'BookmarkGroupManager',
    'VisualActionUpdater'
]
