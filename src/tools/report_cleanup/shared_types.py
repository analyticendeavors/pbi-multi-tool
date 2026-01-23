"""
Shared data classes for Report Cleanup Tool
Built by Reid Havens of Analytic Endeavors
"""

from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional

@dataclass
class DuplicateImageGroup:
    """Represents a group of duplicate images with identical content"""
    group_id: str                           # UUID for the group
    image_hash: str                         # MD5 hex digest
    images: List[Dict[str, Any]] = field(default_factory=list)  # [{name, path, size_bytes, references}]
    selected_image: str = ""                # Auto-selected keeper (shortest filename)
    total_size_bytes: int = 0               # Combined size of all images in group
    savings_bytes: int = 0                  # Size savings after consolidation (total - selected)

@dataclass
class CleanupOpportunity:
    """Represents a cleanup opportunity found in the report"""
    item_type: str  # 'theme', 'custom_visual_build_pane', 'custom_visual_hidden', 'bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group', 'visual_filter', 'dax_query', 'tmdl_script', 'duplicate_image', 'unused_image'
    item_name: str
    location: str   # File path or location description
    reason: str     # Why this can be removed
    safety_level: str = 'safe'  # 'safe', 'warning', 'risky'
    size_bytes: int = 0  # Estimated size savings
    visual_id: str = ""  # For custom visuals, store the actual ID
    bookmark_id: str = ""  # For bookmarks, store the actual ID
    filter_count: int = 0  # For visual filters, store the count of filters
    # Image-specific fields
    duplicate_group_id: str = ""  # For duplicate images, links to DuplicateImageGroup
    image_path: str = ""  # Full path to image file
    references_count: int = 0  # Number of references to this image
    is_orphan: bool = False  # For unused images, True if file exists but not registered in report.json

@dataclass
class RemovalResult:
    """Result of a removal operation"""
    item_name: str
    item_type: str  # 'theme', 'custom_visual', 'bookmark', 'visual_filter', 'duplicate_image', 'unused_image'
    success: bool
    error_message: str = ""
    bytes_freed: int = 0
    filters_hidden: int = 0  # For visual filters, count of filters hidden
    references_updated: int = 0  # For duplicate images, count of references redirected
