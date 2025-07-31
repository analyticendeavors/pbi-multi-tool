"""PBIP Layout Optimizer Tool Package"""

from .layout_ui import PBIPLayoutOptimizerTab
from .tool import PBIPLayoutOptimizerTool

# Tool metadata for registration
TOOL_METADATA = {
    'name': 'PBIP Layout Optimizer',
    'description': 'Optimize Power BI relationship diagram layouts',
    'version': '2.0.0',
    'enabled': True
}
