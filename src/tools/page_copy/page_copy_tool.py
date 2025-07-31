"""
Page Copy Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for the Advanced Page Copy functionality.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
from tools.page_copy.page_copy_ui import AdvancedPageCopyTab


class PageCopyTool(BaseTool):
    """
    Advanced Page Copy Tool - duplicates pages with bookmarks within the same report
    """
    
    def __init__(self):
        super().__init__(
            tool_id="page_copy",
            name="Advanced Page Copy",
            description="Duplicate pages with bookmarks within the same Power BI report",
            version="1.1.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        """Create the Advanced Page Copy UI tab"""
        return AdvancedPageCopyTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        """Get the display title for the tab"""
        return "📋 Advanced Page Copy"
    
    def get_help_content(self) -> Dict[str, Any]:
        """Get help content for the Advanced Page Copy tool"""
        return {
            "title": "Advanced Page Copy - Help",
            "sections": [
                {
                    "title": "🚀 Quick Start",
                    "items": [
                        "1. Select a .pbip report file",
                        "2. Click 'ANALYZE REPORT' to scan for pages with bookmarks", 
                        "3. Select pages to copy from the list (Ctrl+Click for multiple)",
                        "4. Click 'EXECUTE COPY' to duplicate the selected pages"
                    ]
                },
                {
                    "title": "📋 What This Tool Does",
                    "items": [
                        "✅ Duplicates pages within the same report",
                        "✅ Preserves all bookmarks associated with each page",
                        "✅ Automatically renames copies to avoid conflicts",
                        "✅ Updates report metadata after copying"
                    ]
                },
                {
                    "title": "📁 File Requirements",
                    "items": [
                        "✅ Only .pbip files (enhanced PBIR format) are supported",
                        "✅ Only pages with bookmarks are shown for copying",
                        "❌ Pages without bookmarks cannot be copied with this tool"
                    ]
                },
                {
                    "title": "⚠️ Important Notes",
                    "items": [
                        "• Copied pages will have '(Copy)' suffix in their names",
                        "• Bookmarks are duplicated and renamed to avoid conflicts",
                        "• Always backup your report before making changes",
                        "• NOT officially supported by Microsoft"
                    ]
                }
            ],
            "warnings": [
                "This tool only works with PBIP enhanced report format (PBIR) files",
                "Only pages with bookmarks can be copied",
                "Always backup your report before making changes",
                "NOT officially supported by Microsoft"
            ]
        }
    
    def can_run(self) -> bool:
        """Check if the Advanced Page Copy tool can run"""
        try:
            # Check if we can import required modules
            from tools.page_copy.page_copy_core import AdvancedPageCopyEngine
            from tools.report_merger.merger_core import ValidationService, MergerEngine
            return True
        except ImportError as e:
            self.logger.error(f"Page Copy dependencies not available: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Page Copy validation failed: {e}")
            return False
