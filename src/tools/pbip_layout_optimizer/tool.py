"""
PBIP Layout Optimizer Tool
Built by Reid Havens of Analytic Endeavors

This tool provides layout optimization for Power BI relationship diagrams.
"""

from core.tool_manager import BaseTool
from .layout_ui import PBIPLayoutOptimizerTab


class PBIPLayoutOptimizerTool(BaseTool):
    """PBIP Layout Optimizer Tool"""
    
    def __init__(self):
        super().__init__(
            tool_id="pbip_layout_optimizer",
            name="PBIP Layout Optimizer", 
            description="Optimize Power BI relationship diagram layouts using Haven's middle-out design philosophy",
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
            'title': 'PBIP Layout Optimizer Help',
            'description': 'Optimize Power BI relationship diagram layouts using advanced algorithms',
            'features': [
                'Layout Quality Analysis',
                'Automatic Table Positioning', 
                'Haven\'s Middle-Out Design Philosophy',
                'Table Categorization',
                'Relationship Analysis'
            ],
            'requirements': [
                'PBIP format files (.pbip folders)',
                'TMDL files in semantic model',
                'Write permissions to PBIP folder'
            ],
            'usage': [
                '1. Select your PBIP folder using Browse button',
                '2. Click "Analyze Layout" to assess current quality',
                '3. Review analysis results and recommendations', 
                '4. Click "Optimize Layout" to auto-arrange tables',
                '5. Check your Power BI file for improved layout'
            ]
        }
