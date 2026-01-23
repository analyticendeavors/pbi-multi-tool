"""
Connection Hot-Swap Data Models
Built by Reid Havens of Analytic Endeavors

Data classes for representing connections, swap targets, and mappings.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Literal, Any
from enum import Enum

# Import shared cloud models from core.cloud (prevents circular imports)
# Re-exported here for backward compatibility
from core.cloud.models import (
    CloudConnectionType,
    WorkspaceInfo,
    DatasetInfo,
    SwapTarget,
)


class ConnectionType(Enum):
    """Type of data source connection in a Power BI model."""
    LIVE_CONNECTION = "LiveConnection"      # DirectQuery to Analysis Services/Semantic Model
    DIRECT_QUERY = "DirectQuery"            # DirectQuery to SQL/other relational sources
    IMPORT = "Import"                       # Import mode (not swappable)
    DUAL = "Dual"                           # Dual storage mode
    COMPOSITE = "Composite"                 # Model has multiple connection types
    UNKNOWN = "Unknown"                     # Could not determine type


class SwapStatus(Enum):
    """Status of a connection mapping."""
    PENDING = "pending"                     # Not yet configured
    MATCHED = "matched"                     # Auto-matched, awaiting confirmation
    READY = "ready"                         # Configured and validated
    SWAPPING = "swapping"                   # Swap in progress
    SUCCESS = "success"                     # Swap completed successfully
    ERROR = "error"                         # Swap failed


class TomReferenceType(Enum):
    """Type of TOM object reference stored for modification."""
    DATASOURCE = "DataSource"               # Standard DataSource object with ConnectionString
    EXPRESSION_SOURCE = "ExpressionSource"  # Named expression (M query) - modify Expression
    STRUCTURED_DATASOURCE = "StructuredDataSource"  # Uses ConnectionDetails instead
    LIVE_CONNECTION_MODEL = "LiveConnectionModel"  # Pure live connection (no TOM DataSource)
    UNKNOWN = "Unknown"                     # Unknown type - may not be modifiable


@dataclass
class DataSourceConnection:
    """
    Represents a single data source connection in the model.

    Attributes:
        name: DataSource name from TOM
        connection_type: Type of connection (Live, DirectQuery, etc.)
        server: Server/endpoint (e.g., localhost:52784 or powerbi://...)
        database: Database/dataset name or GUID
        provider: Connection provider (e.g., MSOLAP)
        is_cloud: True if this is a cloud XMLA endpoint
        connection_string: Full connection string from TOM
        workspace_name: Workspace name for cloud connections
        dataset_name: Friendly dataset name for cloud connections
        perspective_name: Perspective/cube name if connecting to specific perspective
        cloud_connection_type: Type of cloud connector (PBI Semantic Model vs AAS/XMLA)
        tom_datasource_ref: Reference to TOM DataSource object (not serializable)
    """
    name: str
    connection_type: ConnectionType
    server: str
    database: str
    provider: Optional[str] = None
    is_cloud: bool = False
    connection_string: str = ""
    workspace_name: Optional[str] = None
    dataset_name: Optional[str] = None
    perspective_name: Optional[str] = None
    cloud_connection_type: Optional['CloudConnectionType'] = None  # PBI Semantic Model vs AAS/XMLA
    tom_datasource_ref: Optional[Any] = field(default=None, repr=False, compare=False)
    tom_reference_type: 'TomReferenceType' = field(default=None)
    m_expression: Optional[str] = None  # Original M expression for ExpressionSource types

    @property
    def display_name(self) -> str:
        """User-friendly display name for the connection."""
        if self.is_cloud and self.dataset_name:
            return f"{self.dataset_name} ({self.workspace_name or 'Cloud'})"
        elif self.is_cloud:
            return f"{self.database} (Cloud)"
        elif self.dataset_name:
            # Local with friendly name - show name with "Local" tag
            return f"{self.dataset_name} (Local)"
        else:
            # Local fallback: show server:port first, then truncated GUID for readability
            db_display = self.database[:8] + "..." if len(self.database) > 12 else self.database
            return f"{self.server} ({db_display})"

    @property
    def is_swappable(self) -> bool:
        """Check if this connection can be hot-swapped."""
        # Pure live connection models cannot be swapped (no TOM DataSource to modify)
        if self.tom_reference_type == TomReferenceType.LIVE_CONNECTION_MODEL:
            return False

        # Must be a live or DirectQuery connection type
        if self.connection_type not in (
            ConnectionType.LIVE_CONNECTION,
            ConnectionType.DIRECT_QUERY,
        ):
            return False

        # Check for MSOLAP in provider OR connection string
        has_msolap = False
        if self.provider and 'MSOLAP' in self.provider.upper():
            has_msolap = True
        elif self.connection_string and 'MSOLAP' in self.connection_string.upper():
            has_msolap = True

        # Also allow cloud connections (XMLA endpoints) even without explicit MSOLAP
        if self.is_cloud and self.connection_type == ConnectionType.LIVE_CONNECTION:
            return True

        return has_msolap


# SwapTarget is imported from core.cloud.models above


@dataclass
class ConnectionMapping:
    """
    Mapping between a source connection and its swap target.

    Attributes:
        source: The current connection to be swapped
        target: The target connection to swap to (None if not configured)
        auto_matched: True if target was auto-matched by name
        status: Current status of the mapping
        error_message: Error message if status is ERROR
        original_connection_string: Backup of original connection for rollback
    """
    source: DataSourceConnection
    target: Optional[SwapTarget] = None
    auto_matched: bool = False
    status: SwapStatus = SwapStatus.PENDING
    error_message: Optional[str] = None
    original_connection_string: Optional[str] = None
    original_server: Optional[str] = None
    original_database: Optional[str] = None

    def __post_init__(self):
        """Store original connection details for potential rollback."""
        if self.original_connection_string is None:
            self.original_connection_string = self.source.connection_string
        if self.original_server is None:
            self.original_server = self.source.server
        if self.original_database is None:
            self.original_database = self.source.database

    @property
    def is_configured(self) -> bool:
        """Check if this mapping has a valid target configured."""
        return self.target is not None

    @property
    def is_ready(self) -> bool:
        """Check if this mapping is ready to execute."""
        return self.is_configured and self.status in (SwapStatus.READY, SwapStatus.MATCHED)


@dataclass
class ModelConnectionInfo:
    """
    Complete connection information for an open Power BI model.

    Attributes:
        model_name: Display name of the model
        server: Server this model is connected to
        database: Database name/GUID
        connection_type: Overall connection type of the model
        connections: List of all swappable data source connections
        is_composite: True if model has multiple live connections
        total_datasources: Total number of data sources in model
        swappable_count: Number of swappable connections
    """
    model_name: str
    server: str
    database: str
    connection_type: ConnectionType = ConnectionType.UNKNOWN
    connections: List[DataSourceConnection] = field(default_factory=list)
    is_composite: bool = False
    total_datasources: int = 0
    swappable_count: int = 0

    def __post_init__(self):
        """Calculate derived properties."""
        if not self.swappable_count:
            self.swappable_count = sum(1 for c in self.connections if c.is_swappable)
        if not self.is_composite:
            self.is_composite = len(self.connections) > 1


# WorkspaceInfo and DatasetInfo are imported from core.cloud.models above


@dataclass
class SwapResult:
    """
    Result of a connection swap operation.

    Attributes:
        success: True if swap completed successfully
        mapping: The mapping that was swapped
        message: Success or error message
        elapsed_ms: Time taken for the swap in milliseconds
    """
    success: bool
    mapping: ConnectionMapping
    message: str
    elapsed_ms: int = 0


class PresetStorageType(Enum):
    """Where a preset is stored."""
    USER = "user"          # User's AppData (personal, persists across projects)
    PROJECT = "project"    # Project folder (shareable, version-controllable)
    REPORT = "report"      # PBIP report folder (portable, travels with report)


class PresetScope(Enum):
    """
    Scope of a preset - determines how it's applied and what models it works with.

    GLOBAL: Works with any model that has a single live connection (not composite).
            Stores a single target that gets applied to the model's one connection.
            Useful for: "Always connect to Production Cloud" type presets.

    MODEL: Works with a specific model file (identified by hash).
           Stores the full mapping configuration (all connection-to-target mappings).
           Works with composite models. Like a "bookmark" for a specific model's config.
    """
    GLOBAL = "global"      # Single target, any live-connection model
    MODEL = "model"        # Full mapping snapshot, specific model file


@dataclass
class PresetTargetMapping:
    """
    Serializable mapping of a connection name to its target.

    Used within SwapPreset to store connection configurations.
    """
    connection_name: str          # Source connection name (matches DataSourceConnection.name)
    target_type: Literal["cloud", "local"]
    server: str
    database: str
    display_name: str
    workspace_name: Optional[str] = None
    workspace_id: Optional[str] = None
    dataset_id: Optional[str] = None
    cloud_connection_type: Optional[str] = None  # Stored as string for JSON serialization
    perspective_name: Optional[str] = None
    # Schema tracking for cloud connections (future-proofing)
    cloud_schema_fingerprint: Optional[str] = None  # MD5 fingerprint of key schema fields
    original_cloud_connection: Optional[dict] = None  # Full original connection for restoration

    def to_swap_target(self) -> SwapTarget:
        """Convert to a SwapTarget instance."""
        # Convert cloud_connection_type string back to enum if present
        cloud_type = None
        if self.cloud_connection_type:
            try:
                cloud_type = CloudConnectionType(self.cloud_connection_type)
            except ValueError:
                cloud_type = CloudConnectionType.UNKNOWN

        return SwapTarget(
            target_type=self.target_type,
            server=self.server,
            database=self.database,
            display_name=self.display_name,
            workspace_name=self.workspace_name,
            workspace_id=self.workspace_id,
            dataset_id=self.dataset_id,
            cloud_connection_type=cloud_type,
            perspective_name=self.perspective_name,
        )

    @classmethod
    def from_swap_target(cls, connection_name: str, target: SwapTarget) -> 'PresetTargetMapping':
        """Create from a SwapTarget."""
        # Convert cloud_connection_type enum to string for serialization
        cloud_type_str = None
        if target.cloud_connection_type:
            cloud_type_str = target.cloud_connection_type.value

        return cls(
            connection_name=connection_name,
            target_type=target.target_type,
            server=target.server,
            database=target.database,
            display_name=target.display_name,
            workspace_name=target.workspace_name,
            workspace_id=target.workspace_id,
            dataset_id=target.dataset_id,
            cloud_connection_type=cloud_type_str,
            perspective_name=target.perspective_name,
        )


@dataclass
class SwapPreset:
    """
    A saved environment preset for quick connection switching.

    Stores a named set of connection mappings that can be quickly applied.
    Example presets: "Development", "Testing", "Production"

    Attributes:
        name: Display name of the preset (e.g., "Dev", "Prod")
        description: Optional description of the preset
        mappings: List of connection-to-target mappings
        scope: Whether this is a global or model-specific preset
        model_hash: Hash of the model file (only for MODEL scope)
        model_name: Human-readable model name (only for MODEL scope)
        storage_type: Where the preset is stored (user or project)
        created_at: ISO timestamp when preset was created
        updated_at: ISO timestamp when preset was last modified

    Global presets (scope=GLOBAL):
        - Work with any model that has a single live connection
        - mappings contains just ONE entry (the single target)
        - model_hash and model_name are None

    Model presets (scope=MODEL):
        - Work only with the specific model identified by model_hash
        - mappings contains the full connection configuration snapshot
        - model_hash identifies the model, model_name is for display
    """
    name: str
    mappings: List[PresetTargetMapping] = field(default_factory=list)
    description: Optional[str] = None
    scope: PresetScope = PresetScope.MODEL  # Default to MODEL for backwards compatibility
    model_hash: Optional[str] = None        # Hash of model file (MODEL scope only)
    model_name: Optional[str] = None        # Display name of model (MODEL scope only)
    storage_type: PresetStorageType = PresetStorageType.USER
    created_at: Optional[str] = None
    updated_at: Optional[str] = None

    def __post_init__(self):
        """Set timestamps if not provided."""
        import datetime
        now = datetime.datetime.now().isoformat()
        if self.created_at is None:
            self.created_at = now
        if self.updated_at is None:
            self.updated_at = now

    def to_dict(self) -> dict:
        """Convert to serializable dictionary."""
        return {
            "name": self.name,
            "description": self.description,
            "scope": self.scope.value,
            "model_hash": self.model_hash,
            "model_name": self.model_name,
            "storage_type": self.storage_type.value,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "mappings": [
                {
                    "connection_name": m.connection_name,
                    "target_type": m.target_type,
                    "server": m.server,
                    "database": m.database,
                    "display_name": m.display_name,
                    "workspace_name": m.workspace_name,
                    "workspace_id": m.workspace_id,
                    "dataset_id": m.dataset_id,
                    "cloud_connection_type": m.cloud_connection_type,
                    "perspective_name": m.perspective_name,
                    "cloud_schema_fingerprint": m.cloud_schema_fingerprint,
                    "original_cloud_connection": m.original_cloud_connection,
                }
                for m in self.mappings
            ]
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'SwapPreset':
        """Create from dictionary."""
        mappings = [
            PresetTargetMapping(
                connection_name=m["connection_name"],
                target_type=m["target_type"],
                server=m["server"],
                database=m["database"],
                display_name=m["display_name"],
                workspace_name=m.get("workspace_name"),
                workspace_id=m.get("workspace_id"),
                dataset_id=m.get("dataset_id"),
                cloud_connection_type=m.get("cloud_connection_type"),
                perspective_name=m.get("perspective_name"),
                cloud_schema_fingerprint=m.get("cloud_schema_fingerprint"),
                original_cloud_connection=m.get("original_cloud_connection"),
            )
            for m in data.get("mappings", [])
        ]
        storage_type = PresetStorageType(data.get("storage_type", "user"))
        # Handle scope - default to MODEL for backwards compatibility with old presets
        scope_str = data.get("scope", "model")
        try:
            scope = PresetScope(scope_str)
        except ValueError:
            scope = PresetScope.MODEL
        return cls(
            name=data["name"],
            description=data.get("description"),
            mappings=mappings,
            scope=scope,
            model_hash=data.get("model_hash"),
            model_name=data.get("model_name"),
            storage_type=storage_type,
            created_at=data.get("created_at"),
            updated_at=data.get("updated_at"),
        )

    @property
    def is_global(self) -> bool:
        """Check if this is a global preset."""
        return self.scope == PresetScope.GLOBAL

    @property
    def is_model_specific(self) -> bool:
        """Check if this is a model-specific preset."""
        return self.scope == PresetScope.MODEL

    @property
    def target_type_display(self) -> str:
        """Get display string for target type (Cloud/Local/Mixed)."""
        if not self.mappings:
            return "Empty"
        types = set(m.target_type for m in self.mappings)
        if len(types) == 1:
            return "Cloud" if "cloud" in types else "Local"
        return "Mixed"

    def get_details_summary(self) -> str:
        """Get a human-readable summary of preset details."""
        lines = [f"Preset: {self.name}"]
        if self.description:
            lines.append(f"Description: {self.description}")
        lines.append(f"Scope: {'Global' if self.is_global else 'Model-Specific'}")
        if self.is_model_specific and self.model_name:
            lines.append(f"Model: {self.model_name}")
        lines.append(f"Type: {self.target_type_display}")
        lines.append(f"Connections: {len(self.mappings)}")
        lines.append("")
        lines.append("Mappings:")
        for m in self.mappings:
            target_indicator = "[Cloud]" if m.target_type == "cloud" else "[Local]"
            lines.append(f"  {m.connection_name} -> {m.display_name} {target_indicator}")
        return "\n".join(lines)


@dataclass
class SwapHistoryEntry:
    """
    A record of a past swap operation for rollback.

    Attributes:
        timestamp: ISO timestamp when swap occurred
        connection_name: Name of the swapped connection
        original_server: Server before swap
        original_database: Database before swap
        original_connection_string: Full connection string before swap
        new_server: Server after swap
        new_database: Database after swap
        new_connection_string: Full connection string after swap
        run_id: Unique ID to group batch swaps together (optional)
        model_file_path: File path of the model this swap was performed on (optional)
        source_type: Connection type before swap (Local, Cloud, XMLA)
        target_type: Connection type after swap (Local, Cloud, XMLA)
    """
    timestamp: str
    connection_name: str
    original_server: str
    original_database: str
    original_connection_string: str
    new_server: str
    new_database: str
    new_connection_string: str
    run_id: str = ""  # Groups batch swaps together
    model_file_path: str = ""  # File path of the model for filtering rollback entries
    source_type: str = ""  # Connection type: Local, Cloud, XMLA
    target_type: str = ""  # Connection type: Local, Cloud, XMLA

    @staticmethod
    def _determine_connection_type(
        server: str,
        is_cloud: bool = False,
        cloud_connection_type: Optional['CloudConnectionType'] = None,
        target_type: Optional[str] = None
    ) -> str:
        """Determine human-readable connection type.

        Args:
            server: Server address
            is_cloud: Whether connection is cloud-based
            cloud_connection_type: CloudConnectionType enum value
            target_type: "cloud" or "local" from SwapTarget

        Returns:
            "Local", "Cloud", or "XMLA"
        """
        # Check for local first
        if server and server.lower().startswith("localhost"):
            return "Local"
        if target_type == "local":
            return "Local"

        # Check cloud connection type if available
        if cloud_connection_type:
            if cloud_connection_type == CloudConnectionType.PBI_SEMANTIC_MODEL:
                return "Cloud"
            elif cloud_connection_type == CloudConnectionType.AAS_XMLA:
                return "XMLA"

        # Fallback to server string analysis
        if server:
            server_lower = server.lower()
            if "pbiazure://" in server_lower:
                return "Cloud"
            # Note: Both Semantic Model (pbiServiceXmlaStyleLive) and XMLA (analysisServicesDatabaseLive)
            # use powerbi://...v1.0/... URLs, so we can't distinguish them by URL alone.
            # Default to "Cloud" since it's the safer assumption (works with Pro workspaces).
            # If XMLA was explicitly selected, cloud_connection_type would be set above.
            if "powerbi://" in server_lower and "/v1.0/" in server_lower:
                return "Cloud"

        # If marked as cloud but no specific type
        if is_cloud or target_type == "cloud":
            return "Cloud"

        return "Unknown"

    @classmethod
    def from_mapping(cls, mapping: ConnectionMapping, run_id: str = "", model_file_path: str = "") -> 'SwapHistoryEntry':
        """Create history entry from a completed mapping.

        Uses mapping.original_server/original_database which are preserved from
        before the swap, since mapping.source gets updated to match target after swap.
        """
        import datetime

        # Determine source connection type
        source_type = cls._determine_connection_type(
            server=mapping.original_server or mapping.source.server,
            is_cloud=mapping.source.is_cloud,
            cloud_connection_type=mapping.source.cloud_connection_type
        )

        # Determine target connection type
        target_type = ""
        if mapping.target:
            target_type = cls._determine_connection_type(
                server=mapping.target.server,
                cloud_connection_type=mapping.target.cloud_connection_type,
                target_type=mapping.target.target_type
            )

        return cls(
            timestamp=datetime.datetime.now().isoformat(),
            connection_name=mapping.source.name,
            original_server=mapping.original_server or mapping.source.server,
            original_database=mapping.original_database or mapping.source.database,
            original_connection_string=mapping.original_connection_string or "",
            new_server=mapping.target.server if mapping.target else "",
            new_database=mapping.target.database if mapping.target else "",
            new_connection_string=mapping.target.build_connection_string() if mapping.target else "",
            run_id=run_id,
            model_file_path=model_file_path,
            source_type=source_type,
            target_type=target_type,
        )

    def to_dict(self) -> dict:
        """Convert to serializable dictionary."""
        return {
            "timestamp": self.timestamp,
            "connection_name": self.connection_name,
            "original_server": self.original_server,
            "original_database": self.original_database,
            "original_connection_string": self.original_connection_string,
            "new_server": self.new_server,
            "new_database": self.new_database,
            "new_connection_string": self.new_connection_string,
            "run_id": self.run_id,
            "model_file_path": self.model_file_path,
            "source_type": self.source_type,
            "target_type": self.target_type,
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'SwapHistoryEntry':
        """Create from dictionary."""
        return cls(
            timestamp=data["timestamp"],
            connection_name=data["connection_name"],
            original_server=data["original_server"],
            original_database=data["original_database"],
            original_connection_string=data["original_connection_string"],
            new_server=data["new_server"],
            new_database=data["new_database"],
            new_connection_string=data["new_connection_string"],
            run_id=data.get("run_id", ""),  # Backward compatible
            model_file_path=data.get("model_file_path", ""),  # Backward compatible
            source_type=data.get("source_type", ""),  # Backward compatible
            target_type=data.get("target_type", ""),  # Backward compatible
        )
