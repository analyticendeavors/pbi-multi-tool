"""
Field Parameters Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for Field Parameters functionality.
Provides a UI-driven approach to creating and editing Power BI Field Parameters.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
from tools.field_parameters.field_parameters_ui import FieldParametersTab

# Ownership fingerprint
_AE_FP = "RmllbGRQYXJhbWV0ZXJzOkFFLTIwMjQ="


class FieldParametersTool(BaseTool):
    """
    Field Parameters Tool - Create and edit Power BI Field Parameters with a visual interface
    """
    
    def __init__(self):
        super().__init__(
            tool_id="field_parameters",
            name="Field Parameters",
            description="Create and edit Power BI Field Parameters with drag-drop interface and multi-level categories",
            version="1.0.0"
        )
    
    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        """Create the Field Parameters UI tab"""
        return FieldParametersTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        """Get the display title for the tab"""
        return "🔢 Field Parameters"
    
    def get_help_content(self) -> Dict[str, Any]:
        """Get help content for the Field Parameters tool"""
        return {
            "title": "Field Parameters - Help",
            "sections": [
                {
                    "title": "🚀 Quick Start",
                    "items": [
                        "1. Connect to a Power BI model (via external tool launch or manual selection)",
                        "2. Choose 'Create New' or select existing field parameter to edit",
                        "3. Drag measures/columns from available fields to parameter builder",
                        "4. Assign categories and reorder items as needed",
                        "5. Copy generated TMDL code and paste into your model"
                    ]
                },
                {
                    "title": "📊 What This Tool Does",
                    "items": [
                        "✅ Visual interface for creating field parameters (no manual DAX coding)",
                        "✅ Drag-and-drop to reorder fields (controls sort order)",
                        "✅ Multi-level category support (unlimited nesting)",
                        "✅ Auto-generates proper TMDL/Tabular script code",
                        "✅ Edits existing parameters while preserving extra calculated columns",
                        "✅ Optional lineage tag preservation for version control"
                    ]
                },
                {
                    "title": "🎯 Key Features",
                    "items": [
                        "✅ Model Connection: Connects to open Power BI models via external tool integration",
                        "✅ Display Name Editing: Customize friendly names (different from measure names)",
                        "✅ Bulk Reset: 'Revert All' button to reset display names to original measure names",
                        "✅ Category Management: Create unlimited category levels with drag-drop reordering",
                        "✅ Smart Detection: Auto-identifies existing field parameters in model",
                        "✅ Code Generation: Outputs ready-to-paste TMDL code with proper formatting"
                    ]
                },
                {
                    "title": "📁 Model Connection",
                    "items": [
                        "✅ External Tool Launch: Automatically connects when launched from Power BI Desktop",
                        "✅ Manual Selection: Choose from dropdown of open Power BI models",
                        "✅ Live Validation: Only allows selection of measures/columns from connected model",
                        "✅ Parameter Detection: Finds tables with ParameterMetadata property"
                    ]
                },
                {
                    "title": "🏗️ Understanding Field Parameters",
                    "items": [
                        "A field parameter is a calculated table that contains:",
                        "  • Display column (what users see in slicers)",
                        "  • Fields column (NAMEOF references to actual measures/columns)",
                        "  • Order column (controls item sort order)",
                        "  • Optional: Category columns for grouping (multi-level supported)",
                        "",
                        "Structure: (\"Display Name\", NAMEOF('Table'[Field]), ItemOrder, Cat1Sort, \"Cat1\", Cat2Sort, \"Cat2\"...)"
                    ]
                },
                {
                    "title": "⚠️ Important Notes",
                    "items": [
                        "• Keep Lineage Tags: Check when editing existing parameters to preserve IDs",
                        "• Uncheck for new parameters to let Power BI generate fresh IDs",
                        "• Extra Columns: Tool preserves any calculated columns not part of parameter structure",
                        "• NAMEOF Syntax: Tool uses modern NAMEOF() syntax (Power BI 2020+)",
                        "• Model Format: Requires TMDL format (PBIP) for editing existing parameters",
                        "• Apply Changes: Copy generated code and paste into Tabular Editor or model definition"
                    ]
                },
                {
                    "title": "🔧 Workflow",
                    "items": [
                        "**Create New Parameter:**",
                        "1. Click 'Create New Parameter'",
                        "2. Enter parameter name",
                        "3. Add measures from available fields",
                        "4. Optionally create categories and assign items",
                        "5. Copy TMDL code",
                        "",
                        "**Edit Existing:**",
                        "1. Select parameter from dropdown",
                        "2. Modify display names, reorder items, or adjust categories",
                        "3. Keep lineage tags checked (default)",
                        "4. Copy updated TMDL code"
                    ]
                }
            ]
        }
