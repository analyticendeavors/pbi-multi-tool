"""
Connection Detector - Detect live/DirectQuery connections in Power BI models
Built by Reid Havens of Analytic Endeavors

Analyzes open Power BI models via TOM to identify swappable data source connections.
"""

import logging
import re
from typing import List, Optional, Tuple, Set
from urllib.parse import unquote

from tools.connection_hotswap.models import (
    ConnectionType,
    DataSourceConnection,
    ModelConnectionInfo,
    TomReferenceType,
)

# Ownership fingerprint
_AE_FP = "Q29ubmVjdGlvbkRldGVjdG9yOkFFLTIwMjQ="


class ConnectionDetector:
    """
    Detects live and DirectQuery connections in Power BI models via TOM.

    Uses the PowerBIConnector's TOM access to enumerate DataSources and
    identify which connections can be hot-swapped.
    """

    def __init__(self, connector: 'PowerBIConnector'):
        """
        Initialize the connection detector.

        Args:
            connector: PowerBIConnector instance with active model connection
        """
        self.connector = connector
        self.logger = logging.getLogger(__name__)

    def detect_connections(self) -> ModelConnectionInfo:
        """
        Analyze the connected model and detect all data source connections.

        Returns:
            ModelConnectionInfo with all detected connections and metadata
        """
        if not self.connector._model:
            self.logger.error("No model connected")
            return ModelConnectionInfo(
                model_name="Not Connected",
                server="",
                database="",
                connection_type=ConnectionType.UNKNOWN,
            )

        model = self.connector._model
        model_name = self.connector.current_connection.model_name if self.connector.current_connection else "Unknown"
        server = self.connector.current_connection.server if self.connector.current_connection else ""
        database = self.connector.current_connection.database if self.connector.current_connection else ""

        connections: List[DataSourceConnection] = []
        total_datasources = 0
        seen_datasource_names: Set[str] = set()

        try:
            # Method 1: Enumerate DataSources collection directly
            if hasattr(model, 'DataSources') and model.DataSources:
                total_datasources = model.DataSources.Count
                self.logger.info(f"Found {total_datasources} data source(s) in model.DataSources")

                for i in range(model.DataSources.Count):
                    datasource = model.DataSources[i]
                    ds_name = str(datasource.Name) if hasattr(datasource, 'Name') else f"DS_{i}"
                    ds_type = str(type(datasource).__name__)
                    self.logger.info(f"  DataSource[{i}]: {ds_name} (type: {ds_type})")

                    # Log all available properties for debugging
                    self._log_datasource_properties(datasource)

                    ds_info = self._extract_datasource_info(datasource)
                    if ds_info:
                        connections.append(ds_info)
                        seen_datasource_names.add(ds_name)
                        self.logger.info(f"  → Extracted: {ds_info.name} ({ds_info.connection_type.value}), provider={ds_info.provider}, swappable={ds_info.is_swappable}")

            # Method 2: Check partition sources for DirectQuery connections (composite models)
            partition_connections = self._detect_partition_connections(model, seen_datasource_names)
            if partition_connections:
                self.logger.info(f"Found {len(partition_connections)} additional connection(s) from partitions")
                connections.extend(partition_connections)
                total_datasources += len(partition_connections)

            # Method 3: Detect pure live connection models (non-composite)
            # These models have empty DataSources but are connected to a cloud semantic model
            if not connections:
                live_connection = self._detect_pure_live_connection(model, server, database)
                if live_connection:
                    self.logger.info(f"Detected pure live connection model: {live_connection.name}")
                    connections.append(live_connection)
                    total_datasources = 1

        except Exception as e:
            self.logger.error(f"Error enumerating data sources: {e}")
            import traceback
            self.logger.error(traceback.format_exc())

        # Determine overall connection type
        connection_type = self._detect_model_connection_type(connections)
        is_composite = len(connections) > 1 or self._has_mixed_storage_modes(model)

        return ModelConnectionInfo(
            model_name=model_name,
            server=server,
            database=database,
            connection_type=connection_type,
            connections=connections,
            is_composite=is_composite,
            total_datasources=total_datasources,
            swappable_count=sum(1 for c in connections if c.is_swappable),
        )

    def _log_datasource_properties(self, datasource):
        """Log all available properties of a datasource for debugging."""
        try:
            props = []
            if hasattr(datasource, 'ConnectionString'):
                conn_str = str(datasource.ConnectionString) if datasource.ConnectionString else "(empty)"
                # Truncate for logging
                if len(conn_str) > 100:
                    conn_str = conn_str[:100] + "..."
                props.append(f"ConnectionString={conn_str}")
            if hasattr(datasource, 'Type'):
                props.append(f"Type={datasource.Type}")
            if hasattr(datasource, 'ImpersonationMode'):
                props.append(f"ImpersonationMode={datasource.ImpersonationMode}")
            if hasattr(datasource, 'ConnectionDetails'):
                props.append("Has ConnectionDetails")
            if hasattr(datasource, 'Credential'):
                props.append("Has Credential")

            self.logger.info(f"    Properties: {', '.join(props) if props else '(none detected)'}")
        except Exception as e:
            self.logger.debug(f"    Error logging properties: {e}")

    def _detect_partition_connections(self, model, seen_names: Set[str]) -> List[DataSourceConnection]:
        """
        Detect connections from DirectQuery partitions in composite models.

        In composite models, DirectQuery connections may be referenced through
        partition sources rather than appearing directly in model.DataSources.

        Args:
            model: TOM Model object
            seen_names: Set of already-seen datasource names to avoid duplicates

        Returns:
            List of DataSourceConnection from partition sources
        """
        connections = []
        seen_sources = set()

        try:
            if not hasattr(model, 'Tables') or not model.Tables:
                return connections

            self.logger.info(f"Scanning {model.Tables.Count} table(s) for DirectQuery partitions...")

            for table in model.Tables:
                table_name = str(table.Name) if hasattr(table, 'Name') else "(unnamed)"

                if not hasattr(table, 'Partitions') or not table.Partitions:
                    continue

                for partition in table.Partitions:
                    partition_name = str(partition.Name) if hasattr(partition, 'Name') else "(unnamed)"

                    # Check partition mode
                    mode = str(partition.Mode) if hasattr(partition, 'Mode') else "Unknown"

                    # Check if this is a DirectQuery partition
                    if 'DirectQuery' in mode:
                        self.logger.info(f"  Found DirectQuery partition: {table_name}/{partition_name}")

                        # Get the source
                        source = partition.Source if hasattr(partition, 'Source') else None
                        if source:
                            source_type = str(type(source).__name__)
                            self.logger.info(f"    Source type: {source_type}")

                            # Check for DataSource reference
                            ds_ref = None
                            if hasattr(source, 'DataSource'):
                                ds_ref = source.DataSource
                            elif hasattr(source, 'Expression'):
                                # M expression - may reference a shared expression or datasource
                                expr = str(source.Expression) if source.Expression else ""
                                self.logger.info(f"    M Expression: {expr[:200]}...")

                                # Try to extract datasource name from M expression
                                # Pattern: #"DataSourceName" or Source = DataSourceName
                                import re
                                ds_match = re.search(r'#"([^"]+)"', expr)
                                if ds_match:
                                    ref_name = ds_match.group(1)
                                    self.logger.info(f"    Referenced source: {ref_name}")

                            # Check for EntityPartitionSource properties
                            if hasattr(source, 'EntityName'):
                                entity = str(source.EntityName) if source.EntityName else None
                                if entity:
                                    self.logger.info(f"    EntityName: {entity}")

                            if hasattr(source, 'ExpressionSource'):
                                expr_source = source.ExpressionSource
                                if expr_source:
                                    es_name = str(expr_source.Name) if hasattr(expr_source, 'Name') else "(unnamed)"
                                    self.logger.info(f"    ExpressionSource: {es_name}")

                                    # This is a shared expression (like a Power Query parameter or query)
                                    # Check if it contains a swappable connection
                                    if hasattr(expr_source, 'Expression'):
                                        shared_expr = str(expr_source.Expression) if expr_source.Expression else ""
                                        self.logger.info(f"    Shared expression: {shared_expr[:200]}...")

                                        # Check for Analysis Services connections in the expression
                                        if 'AnalysisServices.Database' in shared_expr or 'powerbi://' in shared_expr.lower():
                                            source_key = es_name
                                            if source_key not in seen_sources and source_key not in seen_names:
                                                seen_sources.add(source_key)
                                                conn = self._extract_from_m_expression(es_name, shared_expr, expr_source)
                                                if conn:
                                                    connections.append(conn)
                                                    self.logger.info(f"    → Created connection from shared expression: {conn.name}")

        except Exception as e:
            self.logger.error(f"Error detecting partition connections: {e}")
            import traceback
            self.logger.error(traceback.format_exc())

        return connections

    def _detect_pure_live_connection(self, model, server: str, database: str) -> Optional[DataSourceConnection]:
        """
        Detect pure live connection models (non-composite).

        These models have empty DataSources but connect to a cloud semantic model.
        The connection info comes from the connector's current connection.

        Args:
            model: TOM Model object
            server: Server from connector's current connection
            database: Database from connector's current connection

        Returns:
            DataSourceConnection if this is a pure live connection model, None otherwise
        """
        try:
            # Only check if this looks like a cloud connection
            if not self._is_cloud_connection(server, ""):
                self.logger.debug("Not a cloud connection - skipping live connection detection")
                return None

            # Verify model structure indicates live connection
            if not self._is_pure_live_connection_model(model):
                self.logger.debug("Model has import/DirectQuery partitions - not a pure live connection")
                return None

            self.logger.info(f"Detected pure live connection model to: {server}")

            # Extract cloud info from server/database
            workspace_name, dataset_name = self._extract_cloud_info(server, database)

            # Build connection string
            connection_string = f"Provider=MSOLAP;Data Source={server};Initial Catalog={database}"

            return DataSourceConnection(
                name="Live Connection",
                connection_type=ConnectionType.LIVE_CONNECTION,
                server=server,
                database=database,
                provider="MSOLAP",
                is_cloud=True,
                connection_string=connection_string,
                workspace_name=workspace_name,
                dataset_name=dataset_name or database,
                tom_datasource_ref=None,  # No TOM ref - connection is at connector level
                tom_reference_type=TomReferenceType.LIVE_CONNECTION_MODEL,
            )

        except Exception as e:
            self.logger.error(f"Error detecting pure live connection: {e}")
            return None

    def _is_pure_live_connection_model(self, model) -> bool:
        """
        Check if this is a pure live connection model (no import data, no DirectQuery partitions).

        Pure live connection models connect to a single external semantic model without
        any local data or DirectQuery sources.

        Args:
            model: TOM Model object

        Returns:
            True if this appears to be a pure live connection model
        """
        try:
            if not hasattr(model, 'Tables') or not model.Tables:
                return False

            # Check if any table has import or DirectQuery partitions
            # Pure live connections don't have these - all tables are "live"
            for table in model.Tables:
                if hasattr(table, 'Partitions') and table.Partitions:
                    for partition in table.Partitions:
                        mode = str(partition.Mode) if hasattr(partition, 'Mode') else ""
                        # If we find Import or DirectQuery, it's not a pure live model
                        if 'Import' in mode or 'DirectQuery' in mode:
                            return False

            # No import/DirectQuery partitions found - likely a pure live connection
            return True

        except Exception as e:
            self.logger.debug(f"Error checking for pure live connection model: {e}")
            return False

    def _extract_from_m_expression(self, name: str, expression: str, tom_ref) -> Optional[DataSourceConnection]:
        """
        Extract connection info from a Power Query M expression.

        Args:
            name: Name of the expression/datasource
            expression: M expression text
            tom_ref: TOM object reference for modification

        Returns:
            DataSourceConnection if extractable, None otherwise
        """
        try:
            server = ""
            database = ""
            is_cloud = False
            workspace_name = None
            dataset_name = None

            # Pattern 1: AnalysisServices.Database("server", "database")
            as_match = re.search(
                r'AnalysisServices\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"',
                expression
            )
            if as_match:
                server = as_match.group(1)
                database = as_match.group(2)
                is_cloud = self._is_cloud_connection(server, expression)
                if is_cloud:
                    workspace_name, dataset_name = self._extract_cloud_info(server, database)
                    self.logger.info(f"    Extracted AS connection: server={server}, db={database}, cloud={is_cloud}")

            # Pattern 2: Direct powerbi:// URL reference
            if not server:
                pbi_match = re.search(r'(powerbi://[^"]+)', expression, re.IGNORECASE)
                if pbi_match:
                    server = pbi_match.group(1)
                    is_cloud = True
                    workspace_name, _ = self._extract_cloud_info(server, "")
                    # Try to find database/catalog
                    cat_match = re.search(r'Initial Catalog\s*=\s*([^;"\)]+)', expression, re.IGNORECASE)
                    if cat_match:
                        database = cat_match.group(1).strip()
                        dataset_name = database

            if server:
                return DataSourceConnection(
                    name=name,
                    connection_type=ConnectionType.LIVE_CONNECTION,
                    server=server,
                    database=database,
                    provider="MSOLAP",  # AS connections use MSOLAP
                    is_cloud=is_cloud,
                    connection_string=f"Provider=MSOLAP;Data Source={server};Initial Catalog={database}",
                    workspace_name=workspace_name,
                    dataset_name=dataset_name or database,
                    tom_datasource_ref=tom_ref,
                    tom_reference_type=TomReferenceType.EXPRESSION_SOURCE,
                    m_expression=expression,
                )

        except Exception as e:
            self.logger.error(f"Error extracting from M expression: {e}")

        return None

    def _extract_datasource_info(self, datasource) -> Optional[DataSourceConnection]:
        """
        Extract connection info from a TOM DataSource object.

        Args:
            datasource: TOM DataSource object

        Returns:
            DataSourceConnection with parsed connection details, or None if not extractable
        """
        try:
            name = str(datasource.Name) if hasattr(datasource, 'Name') else "Unknown"
            connection_string = ""
            server = ""
            database = ""
            provider = ""
            is_cloud = False
            workspace_name = None
            dataset_name = None
            tom_reference_type = TomReferenceType.UNKNOWN

            # Get connection string based on datasource type
            ds_type = str(type(datasource).__name__)

            if hasattr(datasource, 'ConnectionString') and datasource.ConnectionString:
                connection_string = str(datasource.ConnectionString)
                tom_reference_type = TomReferenceType.DATASOURCE
                self.logger.debug(f"  DataSource '{name}' has ConnectionString, type=DATASOURCE")
            elif hasattr(datasource, 'ConnectionDetails') and datasource.ConnectionDetails:
                # StructuredDataSource has ConnectionDetails instead
                tom_reference_type = TomReferenceType.STRUCTURED_DATASOURCE
                self.logger.debug(f"  DataSource '{name}' has ConnectionDetails, type=STRUCTURED_DATASOURCE")
                details = datasource.ConnectionDetails
                if hasattr(details, 'Address'):
                    address = details.Address
                    if hasattr(address, 'Server'):
                        server = str(address.Server) if address.Server else ""
                    if hasattr(address, 'Database'):
                        database = str(address.Database) if address.Database else ""
            else:
                self.logger.debug(f"  DataSource '{name}' has neither ConnectionString nor ConnectionDetails, type=UNKNOWN")

            # Parse connection string if available
            perspective_name = None
            if connection_string:
                parsed = self._parse_connection_string(connection_string)
                server = parsed.get('server', server)
                database = parsed.get('database', database)
                provider = parsed.get('provider', '')
                perspective_name = parsed.get('perspective_name')

            # Detect if this is a cloud connection
            is_cloud = self._is_cloud_connection(server, connection_string)

            # Extract workspace/dataset names from cloud connections
            if is_cloud:
                workspace_name, dataset_name = self._extract_cloud_info(server, database)

            # Determine connection type
            connection_type = self._classify_connection(datasource, provider, is_cloud, connection_string)

            return DataSourceConnection(
                name=name,
                connection_type=connection_type,
                server=server,
                database=database,
                provider=provider,
                is_cloud=is_cloud,
                connection_string=connection_string,
                workspace_name=workspace_name,
                dataset_name=dataset_name,
                perspective_name=perspective_name,
                tom_datasource_ref=datasource,
                tom_reference_type=tom_reference_type,
            )

        except Exception as e:
            self.logger.error(f"Error extracting datasource info: {e}")
            return None

    def _parse_connection_string(self, connection_string: str) -> dict:
        """
        Parse a connection string into its components.

        Args:
            connection_string: MSOLAP or other connection string

        Returns:
            Dict with server, database, provider, perspective_name keys
        """
        result = {'server': '', 'database': '', 'provider': '', 'perspective_name': None}

        if not connection_string:
            return result

        # Handle standard key=value format
        # Provider=MSOLAP;Data Source=localhost:52784;Initial Catalog=DatabaseName;Cube=PerspectiveName
        patterns = {
            'provider': r'Provider\s*=\s*([^;]+)',
            'server': r'Data Source\s*=\s*([^;]+)',
            'database': r'Initial Catalog\s*=\s*([^;]+)',
        }

        for key, pattern in patterns.items():
            match = re.search(pattern, connection_string, re.IGNORECASE)
            if match:
                result[key] = match.group(1).strip()

        # Extract Cube/Perspective name - check quoted and unquoted forms
        cube_match = re.search(r'Cube\s*=\s*"([^"]+)"', connection_string, re.IGNORECASE)
        if not cube_match:
            cube_match = re.search(r'Cube\s*=\s*([^;]+)', connection_string, re.IGNORECASE)

        if cube_match:
            cube_value = cube_match.group(1).strip().strip('"')
            # "Model" is the default cube name, not a perspective
            if cube_value.lower() != 'model':
                result['perspective_name'] = cube_value

        return result

    def _is_cloud_connection(self, server: str, connection_string: str) -> bool:
        """
        Check if this is a cloud XMLA endpoint connection.

        Args:
            server: Server/Data Source value
            connection_string: Full connection string

        Returns:
            True if this is a cloud connection
        """
        cloud_indicators = [
            'powerbi://',
            'api.powerbi.com',
            'asazure.windows.net',
            'pbidedicated.windows.net',
        ]

        check_text = (server + connection_string).lower()
        return any(indicator in check_text for indicator in cloud_indicators)

    def _extract_cloud_info(self, server: str, database: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Extract workspace and dataset names from cloud connection.

        Args:
            server: Server URL (e.g., powerbi://api.powerbi.com/v1.0/myorg/WorkspaceName)
            database: Database name (dataset name)

        Returns:
            Tuple of (workspace_name, dataset_name)
        """
        workspace_name = None
        dataset_name = database if database else None

        # Parse XMLA endpoint URL
        # Format: powerbi://api.powerbi.com/v1.0/myorg/WorkspaceName
        if 'powerbi://' in server.lower():
            # Extract workspace name from URL path
            match = re.search(r'myorg/(.+?)(?:/|$)', server)
            if match:
                workspace_name = unquote(match.group(1))

        return workspace_name, dataset_name

    def _classify_connection(self, datasource, provider: str, is_cloud: bool, connection_string: str = "") -> ConnectionType:
        """
        Classify the connection type based on datasource properties.

        Args:
            datasource: TOM DataSource object
            provider: Provider string (e.g., MSOLAP)
            is_cloud: Whether this is a cloud connection
            connection_string: Connection string for additional analysis

        Returns:
            ConnectionType enum value
        """
        # Check for MSOLAP provider (Analysis Services/Semantic Model)
        if provider and 'MSOLAP' in provider.upper():
            return ConnectionType.LIVE_CONNECTION

        # Check connection string for Analysis Services indicators
        if connection_string:
            conn_upper = connection_string.upper()
            if 'MSOLAP' in conn_upper or 'ANALYSIS SERVICES' in conn_upper:
                return ConnectionType.LIVE_CONNECTION
            if 'POWERBI://' in conn_upper:
                return ConnectionType.LIVE_CONNECTION

        # Check datasource type name
        ds_type = str(type(datasource).__name__).lower()

        if 'provider' in ds_type:
            # ProviderDataSource - could be various types
            return ConnectionType.DIRECT_QUERY if not is_cloud else ConnectionType.LIVE_CONNECTION

        if 'structured' in ds_type:
            # StructuredDataSource - check if it might be an AS connection
            # Some composite models use StructuredDataSource for AS connections
            if is_cloud:
                return ConnectionType.LIVE_CONNECTION
            # Otherwise typically M/Power Query based imports
            return ConnectionType.IMPORT

        return ConnectionType.UNKNOWN

    def _detect_model_connection_type(self, connections: List[DataSourceConnection]) -> ConnectionType:
        """
        Determine the overall connection type of the model.

        Args:
            connections: List of detected connections

        Returns:
            Overall ConnectionType for the model
        """
        if not connections:
            return ConnectionType.IMPORT  # No external connections = import

        # Check if all connections are the same type
        types = set(c.connection_type for c in connections)

        if len(types) == 1:
            return types.pop()

        # Mixed types = composite
        return ConnectionType.COMPOSITE

    def _has_mixed_storage_modes(self, model) -> bool:
        """
        Check if the model has mixed storage modes (Import + DirectQuery).

        Args:
            model: TOM Model object

        Returns:
            True if model has tables with different storage modes
        """
        try:
            modes = set()

            if hasattr(model, 'Tables') and model.Tables:
                for table in model.Tables:
                    if hasattr(table, 'Partitions') and table.Partitions and table.Partitions.Count > 0:
                        # Get first partition to check mode
                        for partition in table.Partitions:
                            if hasattr(partition, 'Mode'):
                                mode = str(partition.Mode)
                                modes.add(mode)
                            break  # Only need first partition per table

            # If we have both Import and DirectQuery modes, it's composite
            return len(modes) > 1

        except Exception as e:
            self.logger.debug(f"Error checking storage modes: {e}")
            return False

    def is_swappable_connection(self, connection: DataSourceConnection) -> bool:
        """
        Check if a connection can be hot-swapped.

        Swappable connections:
        - Live connections to Analysis Services
        - DirectQuery to cloud semantic models
        - Any connection with MSOLAP provider

        Not swappable:
        - Import mode (data is embedded)
        - SQL DirectQuery (different connection paradigm)

        Args:
            connection: DataSourceConnection to check

        Returns:
            True if connection can be swapped
        """
        return connection.is_swappable

    def get_swappable_connections(self) -> List[DataSourceConnection]:
        """
        Get only the swappable connections from the model.

        Returns:
            List of DataSourceConnection objects that can be hot-swapped
        """
        info = self.detect_connections()
        return [c for c in info.connections if c.is_swappable]
