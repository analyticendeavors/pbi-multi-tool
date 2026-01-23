"""
Connection Hot-Swap Tool
Built by Reid Havens of Analytic Endeavors

Enables hot-swapping Power BI live connections between cloud (Power BI Service/Fabric)
and local (Power BI Desktop) instances while the model is open using TOM.

Supports:
- Live connected reports (single DirectQuery connection)
- Composite reports (multiple live connections)
- Auto-matching local models by name
- Cloud workspace browsing via Fabric API
"""

from tools.connection_hotswap.connection_hotswap_tool import ConnectionHotswapTool

__all__ = ['ConnectionHotswapTool']
