"""
Core Cloud Models - Shared data structures for cloud browsing
Built by Reid Havens of Analytic Endeavors

These models are used by the cloud browser components and can be imported
without triggering circular dependencies with tool packages.
"""

from dataclasses import dataclass
from typing import List, Optional, Literal
from enum import Enum


class CloudConnectionType(str, Enum):
    """
    Type of cloud connection used in Power BI.

    These represent fundamentally different connectors with different capabilities:
    - PBI_SEMANTIC_MODEL: Uses pbiServiceLive connector, works with Pro/PPU/Fabric
    - AAS_XMLA: Uses analysisServicesDatabaseLive with XMLA endpoint, supports perspectives
    """
    PBI_SEMANTIC_MODEL = "pbi_semantic_model"  # pbiServiceLive connector
    AAS_XMLA = "aas_xmla"                      # analysisServicesDatabaseLive with XMLA endpoint
    UNKNOWN = "unknown"                        # Could not determine type


@dataclass
class WorkspaceInfo:
    """
    Information about a Power BI Service workspace.

    Attributes:
        id: Workspace GUID
        name: Workspace display name
        type: Workspace type (Workspace, PersonalGroup, etc.)
        is_on_dedicated_capacity: True if on Premium/Fabric capacity
        is_favorite: True if user has favorited this workspace
        capacity_type: Workspace capacity type ('Premium', 'PPU', or None for Pro)
    """
    id: str
    name: str
    type: str = "Workspace"
    is_on_dedicated_capacity: bool = False
    is_favorite: bool = False
    capacity_type: Optional[str] = None  # 'Premium', 'PPU', or None for Pro

    @property
    def xmla_endpoint(self) -> str:
        """Build the XMLA endpoint URL for this workspace."""
        # URL-encode workspace name for the endpoint
        import urllib.parse
        encoded_name = urllib.parse.quote(self.name)
        return f"powerbi://api.powerbi.com/v1.0/myorg/{encoded_name}"


@dataclass
class DatasetInfo:
    """
    Information about a dataset/semantic model in Power BI Service.

    Attributes:
        id: Dataset GUID
        name: Dataset display name
        workspace_id: Parent workspace GUID
        workspace_name: Parent workspace name
        configured_by: User who configured the dataset
        is_refreshable: True if dataset can be refreshed
        is_effective_identity_required: True if RLS requires identity
        perspectives: List of perspective names (requires Premium/XMLA access to discover)
        perspectives_loaded: True if perspectives have been fetched (distinguishes empty from not loaded)
    """
    id: str
    name: str
    workspace_id: str
    workspace_name: str
    configured_by: Optional[str] = None
    is_refreshable: bool = True
    is_effective_identity_required: bool = False
    perspectives: Optional[List[str]] = None
    perspectives_loaded: bool = False

    @property
    def has_perspectives(self) -> bool:
        """Check if this dataset has perspectives available."""
        return self.perspectives is not None and len(self.perspectives) > 0

    def to_swap_target(self, workspace_info: Optional['WorkspaceInfo'] = None,
                       cloud_connection_type: Optional['CloudConnectionType'] = None) -> 'SwapTarget':
        """Convert to a SwapTarget for use in connection mapping."""
        xmla_endpoint = f"powerbi://api.powerbi.com/v1.0/myorg/{self.workspace_name}"
        # Default to PBI_SEMANTIC_MODEL as it's the most universal connector
        if cloud_connection_type is None:
            cloud_connection_type = CloudConnectionType.PBI_SEMANTIC_MODEL
        return SwapTarget(
            target_type="cloud",
            server=xmla_endpoint,
            database=self.name,
            display_name=f"{self.name} ({self.workspace_name})",
            workspace_name=self.workspace_name,
            workspace_id=self.workspace_id,
            dataset_id=self.id,
            cloud_connection_type=cloud_connection_type,
            perspective_name=None,  # Perspectives are selected separately
        )


@dataclass
class SwapTarget:
    """
    Target connection for a swap operation.

    Attributes:
        target_type: Whether this is a cloud or local target
        server: Server address (localhost:port or XMLA endpoint)
        database: Database/dataset name
        display_name: Friendly name for UI display
        workspace_name: Workspace name for cloud targets
        workspace_id: Workspace GUID for cloud targets
        dataset_id: Dataset GUID for cloud targets
        cloud_connection_type: Type of cloud connector (PBI Semantic Model vs AAS/XMLA)
        perspective_name: Perspective/cube name if connecting to specific perspective
    """
    target_type: Literal["cloud", "local"]
    server: str
    database: str
    display_name: str
    workspace_name: Optional[str] = None
    workspace_id: Optional[str] = None
    dataset_id: Optional[str] = None
    cloud_connection_type: Optional['CloudConnectionType'] = None
    perspective_name: Optional[str] = None

    def build_connection_string(self) -> str:
        """Build the MSOLAP connection string for this target."""
        base = f"Provider=MSOLAP;Data Source={self.server};Initial Catalog={self.database}"

        # Add Cube parameter for perspectives (works with both cloud and local targets)
        if self.perspective_name:
            base += f';Cube="{self.perspective_name}"'

        return base
