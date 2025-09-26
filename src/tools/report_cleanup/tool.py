"""
Report Cleanup Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for Report Cleanup functionality.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
from tools.report_cleanup.cleanup_ui import ReportCleanupTab


class ReportCleanupTool(BaseTool):
    """
    Report Cleanup Tool - Detects and removes unused themes and custom visuals from Power BI reports
    """
    
    def __init__(self):
        super().__init__(
            tool_id="report_cleanup",
            name="Report Cleanup",
            description="Detect and remove unused themes and custom visuals from Power BI reports",
            version="1.0.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        """Create the Report Cleanup UI tab"""
        return ReportCleanupTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        """Get the display title for the tab"""
        return "🧹 Report Cleanup"
    
    def get_help_content(self) -> Dict[str, Any]:
        """Get help content for the Report Cleanup tool"""
        return {
            "title": "Report Cleanup - Help",
            "sections": [
                {
                    "title": "🚀 Quick Start",
                    "items": [
                        "1. Select a .pbip report file",
                        "2. Click 'ANALYZE REPORT' to scan for cleanup opportunities", 
                        "3. Choose which items to remove: themes, custom visuals, or both",
                        "4. Click 'REMOVE SELECTED' to clean up your report"
                    ]
                },
                {
                    "title": "🧹 What This Tool Detects",
                    "items": [
                        "✅ Unused themes in BaseThemes and RegisteredResources",
                        "✅ Custom visuals that appear in build pane but are unused",
                        "✅ Hidden custom visuals (installed but not in build pane)",
                        "✅ Safely identifies items that can be removed without affecting functionality"
                    ]
                },
                {
                    "title": "🔍 Detection Logic",
                    "items": [
                        "✅ Compares active theme with available theme files",
                        "✅ Scans all pages for custom visual usage in visualContainers",
                        "✅ Identifies 'hidden' visuals taking up space in CustomVisuals folder",
                        "❌ Only removes themes and visuals that are truly unused"
                    ]
                },
                {
                    "title": "📁 File Requirements",
                    "items": [
                        "✅ Only .pbip files (enhanced PBIR format) are supported",
                        "✅ Must have corresponding .Report directory",
                        "✅ Requires StaticResources and CustomVisuals directories for full analysis"
                    ]
                },
                {
                    "title": "⚠️ Important Safety Notes",
                    "items": [
                        "• Always backup your PBIP file before cleanup",
                        "• Tool creates automatic backups with timestamps",
                        "• Only removes themes and visuals confirmed as unused",
                        "• Manual review recommended for custom visuals with complex dependencies",
                        "• NOT officially supported by Microsoft"
                    ]
                }
            ],
            "warnings": [
                "This tool only works with PBIP enhanced report format (PBIR) files",
                "Always backup your report before making changes",
                "Only removes themes and visuals that are confirmed unused",
                "Custom visuals with external dependencies may require manual verification",
                "NOT officially supported by Microsoft"
            ]
        }
    
    def can_run(self) -> bool:
        """Check if the Report Cleanup tool can run"""
        try:
            # Check if we can import required modules
            from tools.report_cleanup.report_analyzer import ReportAnalyzer
            from tools.report_cleanup.cleanup_engine import ReportCleanupEngine
            return True
        except ImportError as e:
            self.logger.error(f"Report Cleanup dependencies not available: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Report Cleanup validation failed: {e}")
            return False
