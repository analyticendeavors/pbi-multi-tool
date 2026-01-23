"""
Layout Optimizer Tool
Built by Reid Havens of Analytic Endeavors

This tool provides layout optimization for Power BI relationship diagrams.
"""

from core.tool_manager import BaseTool
from .layout_ui import PBIPLayoutOptimizerTab

# Ownership fingerprint
_AE_FP = "TGF5b3V0T3B0aW1pemVyOkFFLTIwMjQ="


class PBIPLayoutOptimizerTool(BaseTool):
    """Layout Optimizer Tool"""

    def __init__(self):
        super().__init__(
            tool_id="pbip_layout_optimizer",
            name="Layout Optimizer",
            description="Optimize Power BI relationship diagram layouts for PBIP projects",
            version="2.0.0"
        )
    
    def create_ui_tab(self, parent, main_app):
        """Create the layout optimizer UI tab"""
        return PBIPLayoutOptimizerTab(parent, main_app)
    
    def get_tab_title(self) -> str:
        """Get the tab title with emoji"""
        return "🎯 Layout Optimizer"
    
    def get_help_content(self) -> dict:
        """Get help content for this tool"""
        return {
            'title': 'Layout Optimizer Help',
            'description': 'Optimize Power BI relationship diagram layouts using advanced algorithms',
            'features': [
                'Layout Quality Analysis',
                'Automatic Table Positioning',
                'Haven\'s Middle-Out Design Philosophy',
                'Table Categorization',
                'Relationship Analysis'
            ],
            'requirements': [
                'PBIP projects (.pbip files)',
                'Semantic model with TMDL files',
                'Write permissions to the folder'
            ],
            'usage': [
                '1. Select your .pbip file using Browse button',
                '2. Click "Analyze Layout" to assess current quality',
                '3. Review analysis results and recommendations',
                '4. Click "Optimize Layout" to auto-arrange tables',
                '5. Open the PBIP in Power BI Desktop to see the improved layout'
            ]
        }
