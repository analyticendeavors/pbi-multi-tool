"""
Table Column Widths Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for Table Column Widths functionality.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
from tools.column_width.column_width_ui import TableColumnWidthsTab


class TableColumnWidthsTool(BaseTool):
    """
    Table Column Widths Tool - Column width standardization and visual formatting
    """
    
    def __init__(self):
        super().__init__(
            tool_id="column_width",
            name="Table Column Widths",
            description="Standardize column widths across tables and matrices in Power BI reports",
            version="2.0.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        """Create the Table Column Widths UI tab"""
        return TableColumnWidthsTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        """Get the display title for the tab"""
        return "📊 Table Column Widths"
    
    def get_help_content(self) -> Dict[str, Any]:
        """Get help content for the Table Column Widths tool"""
        return {
            "title": "Table Column Widths - Help",
            "sections": [
                {
                    "title": "🚀 Quick Start",
                    "items": [
                        "1. Select a .pbip report file",
                        "2. Click 'SCAN VISUALS' to analyze tables and matrices", 
                        "3. Choose preset widths or set custom values for categorical columns and measures",
                        "4. Select which visuals to update (individual, multiple, or all)",
                        "5. Click 'APPLY CHANGES' to standardize column widths"
                    ]
                },
                {
                    "title": "📊 What This Tool Does",
                    "items": [
                        "✅ Analyzes all Table and Matrix visuals in your report",
                        "✅ Automatically categorizes columns (categorical vs measures)",
                        "✅ Calculates optimal column widths based on text length and font settings",
                        "✅ Applies consistent width standards across multiple visuals",
                        "✅ Prevents text wrapping beyond 3 lines with intelligent sizing",
                        "✅ Supports both Table (tableEx) and Matrix (pivotTable) visual types"
                    ]
                },
                {
                    "title": "📏 Width Calculation Features",
                    "items": [
                        "✅ Reads actual font settings from each visual configuration",
                        "✅ Measures column header text length for precise calculations",
                        "✅ Auto-fit preset calculates optimal width to prevent wrapping",
                        "✅ Categorical columns: Dimensions, hierarchies, text fields",
                        "✅ Measure columns: Numeric measures and calculated fields",
                        "✅ Maximum width caps prevent extremely wide columns"
                    ]
                },
                {
                    "title": "🎯 Application Scope Options",
                    "items": [
                        "✅ Single Visual: Apply to one selected table/matrix",
                        "✅ Multiple Visuals: Choose specific visuals to update",
                        "✅ All Visuals: Apply standards to all tables and matrices",
                        "✅ Visual Type Filter: Tables only, matrices only, or both",
                        "✅ Preview mode shows changes before applying"
                    ]
                },
                {
                    "title": "📁 File Requirements",
                    "items": [
                        "✅ Only .pbip files (enhanced PBIR format) are supported",
                        "✅ Must have corresponding .Report directory",
                        "✅ Requires visual JSON files in pages/visuals directories"
                    ]
                },
                {
                    "title": "⚠️ Important Safety Notes",
                    "items": [
                        "• Always backup your PBIP file before applying changes",
                        "• Tool creates automatic backups with timestamps",
                        "• Column width changes only affect visual formatting, not data",
                        "• Changes are applied to the visual JSON configurations",
                        "• Compatible with Power BI Desktop and Power BI Service",
                        "• NOT officially supported by Microsoft"
                    ]
                }
            ],
            "warnings": [
                "This tool only works with PBIP enhanced report format (PBIR) files",
                "Always backup your report before making changes",
                "Column width changes only affect visual appearance, not data integrity",
                "Font settings are read from visual configurations for precise calculations",
                "NOT officially supported by Microsoft"
            ]
        }
    
    def can_run(self) -> bool:
        """Check if the Table Column Widths tool can run"""
        try:
            # Check if we can import required modules
            from tools.column_width.column_width_core import TableColumnWidthsEngine
            return True
        except ImportError as e:
            self.logger.error(f"Table Column Widths dependencies not available: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Table Column Widths validation failed: {e}")
            return False
