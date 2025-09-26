"""
Shared data classes for Report Cleanup Tool
Built by Reid Havens of Analytic Endeavors
"""

from dataclasses import dataclass
from typing import Dict, List, Any, Optional

@dataclass
class CleanupOpportunity:
    """Represents a cleanup opportunity found in the report"""
    item_type: str  # 'theme', 'custom_visual_build_pane', 'custom_visual_hidden', 'bookmark_guaranteed_unused', 'bookmark_likely_unused', 'visual_filter'
    item_name: str
    location: str   # File path or location description
    reason: str     # Why this can be removed
    safety_level: str = 'safe'  # 'safe', 'warning', 'risky'
    size_bytes: int = 0  # Estimated size savings
    visual_id: str = ""  # For custom visuals, store the actual ID
    bookmark_id: str = ""  # For bookmarks, store the actual ID
    filter_count: int = 0  # For visual filters, store the count of filters

@dataclass
class RemovalResult:
    """Result of a removal operation"""
    item_name: str
    item_type: str  # 'theme', 'custom_visual', 'bookmark', or 'visual_filter'
    success: bool
    error_message: str = ""
    bytes_freed: int = 0
    filters_hidden: int = 0  # For visual filters, count of filters hidden
