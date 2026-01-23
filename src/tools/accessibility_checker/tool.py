"""
Accessibility Checker Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for Accessibility Checker functionality.
Analyzes Power BI reports for WCAG accessibility compliance.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab

# Ownership fingerprint
_AE_FP = "QWNjZXNzQ2hlY2tlclRvb2w6QUUtMjAyNA=="


class AccessibilityCheckerTool(BaseTool):
    """
    Accessibility Checker Tool - Analyzes Power BI reports for accessibility compliance
    and WCAG issues including tab order, alt text, color contrast, and more.
    """

    def __init__(self):
        super().__init__(
            tool_id="accessibility_checker",
            name="Accessibility Checker",
            description="Analyze Power BI reports for accessibility compliance and WCAG issues",
            version="1.0.0"
        )

    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        """Create the Accessibility Checker UI tab"""
        from tools.accessibility_checker.accessibility_ui import AccessibilityCheckerTab
        return AccessibilityCheckerTab(parent, main_app)

    def get_tab_title(self) -> str:
        """Get the display title for the tab"""
        return "Accessibility Checker"

    def get_help_content(self) -> Dict[str, Any]:
        """Get help content for the Accessibility Checker tool"""
        return {
            "title": "Accessibility Checker - Help",
            "sections": [
                {
                    "title": "Quick Start",
                    "items": [
                        "1. Select a Power BI report file",
                        "2. Click 'ANALYZE REPORT' to scan for accessibility issues",
                        "3. Review issues organized by category cards",
                        "4. Click category cards to see detailed issue list",
                        "5. Export report for documentation or remediation tracking"
                    ]
                },
                {
                    "title": "WCAG 2.1 Compliance Levels",
                    "items": [
                        "Level A: Minimum accessibility requirements (essential)",
                        "Level AA: Standard compliance target (recommended - most common)",
                        "Level AAA: Enhanced accessibility (best practice)"
                    ]
                },
                {
                    "title": "Check Categories & WCAG Criteria",
                    "items": [
                        "Tab Order (WCAG 2.4.3 Focus Order): Keyboard navigation must follow a logical sequence matching visual layout",
                        "Alt Text (WCAG 1.1.1 Non-text Content): All images, charts, and data visuals need text descriptions for screen readers",
                        "Color Contrast (WCAG 1.4.3 AA / 1.4.6 AAA): Text must have sufficient contrast ratio against background",
                        "Page Titles (WCAG 2.4.2 Page Titled): Pages need descriptive names, not generic like 'Page 1'",
                        "Visual Titles (WCAG 2.4.6 Headings and Labels): Charts and visuals need descriptive headings",
                        "Data Labels (WCAG 1.4.1 Use of Color): Don't rely solely on color to convey information",
                        "Bookmark Names (WCAG 2.4.6 Headings and Labels): Bookmarks need descriptive names for navigation",
                        "Hidden Pages (WCAG 2.4.1 Bypass Blocks): Ensure hidden pages don't contain essential content"
                    ]
                },
                {
                    "title": "Color Contrast Requirements",
                    "items": [
                        "Normal text (Level AA): 4.5:1 minimum contrast ratio",
                        "Large text (Level AA): 3:1 minimum (18pt+ or 14pt+ bold)",
                        "Normal text (Level AAA): 7:1 minimum contrast ratio",
                        "Large text (Level AAA): 4.5:1 minimum ratio",
                        "UI Components & Graphics: 3:1 minimum ratio"
                    ]
                },
                {
                    "title": "Issue Severity Levels",
                    "items": [
                        "ERROR (Red): Critical issues that fail WCAG compliance requirements",
                        "WARNING (Yellow): Issues that should be fixed for better accessibility",
                        "INFO (Blue): Suggestions, best practices, and informational notes"
                    ]
                },
                {
                    "title": "File Requirements",
                    "items": [
                        "PBIX and PBIP files saved in PBIR format are supported",
                        "Must have corresponding .Report directory (created automatically with PBIR format)",
                        "PBIR is the default format since January 2026 Power BI Desktop"
                    ]
                },
                {
                    "title": "Important Notes",
                    "items": [
                        "This tool is read-only - it does not modify your report",
                        "Automated checks are a starting point - manual review recommended",
                        "Some checks may not apply to all report types or configurations",
                        "Not officially supported by Microsoft"
                    ]
                }
            ],
            "warnings": [
                "This tool only works with reports saved in PBIR format",
                "Some checks require manual verification for complete compliance",
                "Color contrast checks analyze configured text colors vs background",
                "Not officially supported by Microsoft"
            ]
        }

    def can_run(self) -> bool:
        """Check if the Accessibility Checker tool can run"""
        try:
            # Check if we can import required modules
            from tools.accessibility_checker.accessibility_analyzer import AccessibilityAnalyzer
            from tools.accessibility_checker.accessibility_types import AccessibilityAnalysisResult
            return True
        except ImportError as e:
            self.logger.error(f"Accessibility Checker dependencies not available: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Accessibility Checker validation failed: {e}")
            return False
