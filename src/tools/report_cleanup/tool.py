"""
Report Cleanup Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for Report Cleanup functionality.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
from tools.report_cleanup.cleanup_ui import ReportCleanupTab

# Ownership fingerprint
_AE_FP = "UmVwb3J0Q2xlYW51cDpBRS0yMDI0"


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
                        "1. Select a Power BI report file",
                        "2. Click 'ANALYZE REPORTS' to scan for cleanup opportunities",
                        "3. Choose which items to remove: themes, custom visuals, or both",
                        "4. Click 'CLEAN SELECTED' to clean up your report"
                    ]
                },
                {
                    "title": "🧹 What This Tool Detects",
                    "items": [
                        "✅ Unused themes in BaseThemes and RegisteredResources",
                        "✅ Custom visuals that appear in build pane but are unused",
                        "✅ Hidden custom visuals (installed but not in build pane)",
                        "✅ Unused bookmarks (missing pages or no navigation)",
                        "✅ Visual-level filters that can be hidden",
                        "✅ Safely identifies items that can be removed without affecting functionality"
                    ]
                },
                {
                    "title": "🔍 Detection Logic",
                    "items": [
                        "✅ Compares active theme with available theme files",
                        "✅ Scans all pages for custom visual usage in visualContainers",
                        "✅ Identifies 'hidden' visuals taking up space in CustomVisuals folder",
                        "✅ Analyzes bookmark navigation & pages",
                        "❌ Only removes themes and visuals that are truly unused"
                    ]
                },
                {
                    "title": "📁 File Requirements",
                    "items": [
                        "✅ PBIP files: Full support (analyze AND cleanup)",
                        "⚠️ PBIX files: Analysis only (cleanup requires PBIP)",
                        "✅ Must be saved in PBIR format (default since Jan 2025)",
                        "💡 To modify PBIX: Open in Desktop, Save As PBIP"
                    ]
                },
                {
                    "title": "⚠️ Important Safety Notes",
                    "items": [
                        "• Always backup your report before cleanup",
                        "• Tool creates automatic backups with timestamps",
                        "• Only removes themes and visuals confirmed as unused",
                        "• Manual review recommended for custom visuals with complex dependencies",
                        "• NOT officially supported by Microsoft"
                    ]
                }
            ],
            "warnings": [
                "PBIX: Analysis only - Power BI validates content integrity",
                "       and doesn't easily allow external modifications",
                "PBIP: Full support for analysis AND cleanup operations",
                "Always backup your report before making changes",
                "Only removes themes and visuals that are confirmed unused",
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
