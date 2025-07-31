"""
Report Merger Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for the Report Merger functionality.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
from tools.report_merger.merger_ui import ReportMergerTab


class ReportMergerTool(BaseTool):
    """
    Report Merger Tool - combines multiple Power BI reports into one
    """
    
    def __init__(self):
        super().__init__(
            tool_id="report_merger",
            name="Power BI Report Merger",
            description="Combine multiple PBIP reports into a single consolidated report",
            version="1.1.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        """Create the Report Merger UI tab"""
        return ReportMergerTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        """Get the display title for the tab"""
        return "📊 Report Merger"
    
    def get_help_content(self) -> Dict[str, Any]:
        """Get help content for the Report Merger tool"""
        return {
            "title": "Power BI Report Merger - Help",
            "sections": [
                {
                    "title": "🚀 Quick Start",
                    "items": [
                        "1. Select Report A and Report B (.pbip files)",
                        "2. Click 'ANALYZE REPORTS' to validate and preview",
                        "3. Choose theme if there's a conflict",
                        "4. Click 'EXECUTE MERGE' to create combined report"
                    ]
                },
                {
                    "title": "📁 File Requirements",
                    "items": [
                        "✅ Only .pbip files (enhanced PBIR format) are supported",
                        "✅ Reports must have definition\\ folder structure",
                        "❌ Legacy format with report.json files are NOT supported",
                        "❌ .pbix files are NOT supported"
                    ]
                },
                {
                    "title": "⚠️ Important Disclaimers",
                    "items": [
                        "• This tool ONLY works with PBIP enhanced report format (PBIR) files",
                        "• This is NOT officially supported by Microsoft - use at your own discretion",
                        "• Always keep backups of your original reports before merging",
                        "• Test thoroughly and validate merged results before production use"
                    ]
                }
            ],
            "warnings": [
                "This tool ONLY works with PBIP enhanced report format (PBIR) files",
                "This is NOT officially supported by Microsoft - use at your own discretion",
                "Look for .pbip files with definition\\ folder (not report.json files)",
                "Always keep backups of your original reports before merging",
                "Test thoroughly and validate merged results before production use",
                "Enable 'Store reports using enhanced metadata format (PBIR)' in Power BI Desktop"
            ]
        }
    
    def can_run(self) -> bool:
        """Check if the Report Merger tool can run"""
        try:
            # Check if we can import required modules
            from tools.report_merger.merger_core import MergerEngine, ValidationService
            return True
        except ImportError as e:
            self.logger.error(f"Report Merger dependencies not available: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Report Merger validation failed: {e}")
            return False
