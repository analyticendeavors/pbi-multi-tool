"""
Connection Hot-Swap Tool - Main tool implementation
Built by Reid Havens of Analytic Endeavors

This implements the BaseTool interface for Connection Hot-Swap functionality.
Enables toggling Power BI live connections between cloud and local models.
"""

from typing import Dict, Any
from core.tool_manager import BaseTool
from core.ui_base import BaseToolTab
from tools.connection_hotswap.ui.connection_hotswap_tab import ConnectionHotswapTab

# Ownership fingerprint
_AE_FP = "Q29ubmVjdGlvbkhvdFN3YXA6QUUtMjAyNA=="


class ConnectionHotswapTool(BaseTool):
    """
    Connection Hot-Swap Tool - Toggle live connections between cloud and local models
    """

    def __init__(self):
        super().__init__(
            tool_id="connection_hotswap",
            name="Connection Hot-Swap",
            description="Hot-swap Power BI live connections between cloud and local models while open",
            version="2.2.0"
        )

    def create_ui_tab(self, parent, main_app) -> BaseToolTab:
        """Create the Connection Hot-Swap UI tab"""
        return ConnectionHotswapTab(parent, main_app)

    def get_tab_title(self) -> str:
        """Get the display title for the tab"""
        return "🔄 Connection Hot-Swap"

    def get_help_content(self) -> Dict[str, Any]:
        """Get help content for the Connection Hot-Swap tool"""
        return {
            "title": "Connection Hot-Swap - Help",
            "sections": [
                {
                    "title": "🚀 Quick Start",
                    "items": [
                        "1. Open your Power BI report in Power BI Desktop",
                        "2. Launch this tool from the External Tools ribbon",
                        "3. The tool will detect live/DirectQuery connections",
                        "4. Choose target (local model or cloud workspace)",
                        "5. Click 'Swap' to change the connection"
                    ]
                },
                {
                    "title": "📊 What This Tool Does",
                    "items": [
                        "✅ Swap connections: cloud-to-local, local-to-cloud, or cloud-to-cloud",
                        "✅ Composite models: Hot-swap while open (TOM-based)",
                        "✅ Thin reports: Swap via PBIP file modification",
                        "✅ Connect to perspectives in ANY workspace (Pro, Premium, PPU, Fabric)",
                        "✅ Auto-match local models by name",
                        "✅ Save and reuse connection presets"
                    ]
                },
                {
                    "title": "🎯 Use Cases",
                    "items": [
                        "✅ Development: Switch from production cloud model to local dev copy",
                        "✅ Testing: Swap to local model with test data",
                        "✅ Offline Work: Use local model when cloud is unavailable",
                        "✅ Performance: Test with local model during development",
                        "✅ Composite Models: Swap individual connections independently"
                    ]
                },
                {
                    "title": "☁️ Cloud Model Selection",
                    "items": [
                        "Browse all workspaces you have access to",
                        "Filter by favorites or recently used",
                        "Search across all workspaces",
                        "Manual XMLA endpoint entry for direct access",
                        "Works with Pro, Premium, PPU, and Fabric workspaces"
                    ]
                },
                {
                    "title": "🖥️ Local Model Selection",
                    "items": [
                        "Automatically scans for open Power BI Desktop instances",
                        "Auto-matches models by name (suggest closest match)",
                        "Manual port entry for custom configurations",
                        "Shows model name from window title"
                    ]
                },
                {
                    "title": "🔀 Composite Model Support",
                    "items": [
                        "Detects all live connections in composite models",
                        "Shows mapping table for each connection",
                        "Swap connections individually or all at once",
                        "Supports different targets for different connections"
                    ]
                },
                {
                    "title": "⚡ Quick Swap Buttons",
                    "items": [
                        "'Swap All to Cloud': Switch all connections to cloud targets",
                        "'Swap All to Local': Switch all connections to local targets",
                        "Auto-match: Automatically suggests matching local models"
                    ]
                },
                {
                    "title": "⚠️ Important Notes",
                    "items": [
                        "• Composite models: Hot-swap while open in Desktop (TOM-based)",
                        "• Thin reports: Close & reopen after swap (PBIP can be edited while open)",
                        "• Perspectives work in ALL workspace types (Pro, Premium, PPU, Fabric)",
                        "• After swapping, refresh the report to load new data",
                        "• Changes are saved immediately - save your report afterward",
                        "• Original connection is backed up for rollback if swap fails"
                    ]
                },
                {
                    "title": "🔧 Connection String Formats",
                    "items": [
                        "**Local Connection:**",
                        "Provider=MSOLAP;Data Source=localhost:PORT;Initial Catalog=DATABASE",
                        "",
                        "**Cloud XMLA Endpoint:**",
                        "Provider=MSOLAP;Data Source=powerbi://api.powerbi.com/v1.0/myorg/WORKSPACE;Initial Catalog=DATASET"
                    ]
                }
            ]
        }

    def can_run(self) -> bool:
        """
        Check if the tool can run (TOM dependencies available).

        Returns True if:
        - pythonnet is installed
        - Power BI Desktop or TOM DLL is available
        """
        try:
            # Check for pythonnet
            import clr

            # Check for Power BI Desktop installation (most reliable source of TOM)
            from pathlib import Path
            pbi_path = Path(r"C:\Program Files\Microsoft Power BI Desktop\bin")
            if pbi_path.exists():
                return True

            # Fall back to checking for TOM DLL in common paths
            tom_paths = [
                Path(r"C:\Program Files\Tabular Editor 3\Microsoft.AnalysisServices.Tabular.dll"),
                Path(r"C:\Program Files (x86)\Microsoft SQL Server Management Studio 19\Common7\IDE\Extensions\Microsoft\SQLDB\DAC\Microsoft.AnalysisServices.Tabular.dll"),
            ]

            return any(p.exists() for p in tom_paths)

        except ImportError:
            # pythonnet not installed
            self.logger.warning("pythonnet not installed - Connection Hot-Swap requires pythonnet")
            return False
        except Exception as e:
            self.logger.warning(f"can_run exception: {e}")
            return False
