"""
Connection Swapper - TOM-based connection modification
Built by Reid Havens of Analytic Endeavors

Handles hot-swapping data source connections in open Power BI models via TOM.
"""

import logging
import time
import socket
import re
from typing import Tuple, Optional
from urllib.parse import quote

from tools.connection_hotswap.models import (
    ConnectionMapping,
    SwapTarget,
    SwapStatus,
    SwapResult,
    DataSourceConnection,
    TomReferenceType,
)

# Ownership fingerprint
_AE_FP = "Q29ubmVjdGlvblN3YXBwZXI6QUUtMjAyNA=="


class ConnectionSwapper:
    """
    Swaps data source connections in live Power BI models via TOM.

    Modifies DataSource.ConnectionString and calls Model.SaveChanges()
    to persist the change while the model is open.
    """

    def __init__(self, connector: 'PowerBIConnector'):
        """
        Initialize the connection swapper.

        Args:
            connector: PowerBIConnector instance with active model connection
        """
        self.connector = connector
        self.logger = logging.getLogger(__name__)

    def swap_connection(
        self,
        mapping: ConnectionMapping,
        dry_run: bool = False
    ) -> SwapResult:
        """
        Execute a connection swap.

        Modifies the DataSource connection string via TOM and saves the model.
        Handles different TOM reference types: DataSource, ExpressionSource, StructuredDataSource.

        Args:
            mapping: The connection mapping to apply
            dry_run: If True, validate only without applying

        Returns:
            SwapResult with success status and message
        """
        start_time = time.time()

        # Validate prerequisites
        if not mapping.target:
            return SwapResult(
                success=False,
                mapping=mapping,
                message="No target configured for this mapping",
                elapsed_ms=0,
            )

        if not mapping.source.tom_datasource_ref:
            return SwapResult(
                success=False,
                mapping=mapping,
                message="No TOM reference available for source connection",
                elapsed_ms=0,
            )

        if not self.connector._model:
            return SwapResult(
                success=False,
                mapping=mapping,
                message="No model connected",
                elapsed_ms=0,
            )

        # Build new connection string
        new_conn_string = mapping.target.build_connection_string()
        ref_type = mapping.source.tom_reference_type or TomReferenceType.UNKNOWN

        self.logger.info(f"Swapping connection '{mapping.source.name}' (type: {ref_type.value})")
        self.logger.debug(f"  From: {mapping.source.connection_string[:100]}...")
        self.logger.debug(f"  To: {new_conn_string[:100]}...")

        if dry_run:
            elapsed = int((time.time() - start_time) * 1000)
            return SwapResult(
                success=True,
                mapping=mapping,
                message=f"Dry run: Would swap to {mapping.target.display_name}",
                elapsed_ms=elapsed,
            )

        try:
            # Store original values for rollback/history (before we modify mapping.source)
            mapping.original_connection_string = mapping.source.connection_string
            mapping.original_server = mapping.source.server
            mapping.original_database = mapping.source.database

            # Re-fetch TOM reference to avoid stale COM objects
            tom_ref = self._refresh_tom_reference(mapping.source)
            if tom_ref is None:
                return SwapResult(
                    success=False,
                    mapping=mapping,
                    message="Could not refresh TOM reference - datasource may have been removed",
                    elapsed_ms=int((time.time() - start_time) * 1000),
                )

            # Apply the swap based on reference type
            if ref_type == TomReferenceType.EXPRESSION_SOURCE:
                success, message = self._swap_expression_source(tom_ref, mapping, new_conn_string)
            elif ref_type == TomReferenceType.STRUCTURED_DATASOURCE:
                success, message = self._swap_structured_datasource(tom_ref, mapping)
            else:
                # Default: standard DataSource with ConnectionString
                success, message = self._swap_standard_datasource(tom_ref, new_conn_string)

            if not success:
                mapping.status = SwapStatus.ERROR
                mapping.error_message = message
                return SwapResult(
                    success=False,
                    mapping=mapping,
                    message=message,
                    elapsed_ms=int((time.time() - start_time) * 1000),
                )

            # Save changes to the model
            self.logger.info("Calling model.SaveChanges()...")
            self.connector._model.SaveChanges()

            # Attempt to force reconnection by requesting a Calculate refresh
            # This may help Power BI Desktop recognize the connection change
            try:
                from Microsoft.AnalysisServices.Tabular import RefreshType
                self.logger.info("Requesting Calculate refresh to force reconnection...")
                self.connector._model.RequestRefresh(RefreshType.Calculate)
                self.connector._model.SaveChanges()
                self.logger.info("Calculate refresh requested successfully")
            except Exception as refresh_err:
                # Non-fatal - the metadata change was already saved
                self.logger.warning(f"RequestRefresh failed (non-fatal): {refresh_err}")

            # Validate the change persisted
            validation_ok, validation_msg = self._validate_swap_persisted(mapping, new_conn_string)
            if not validation_ok:
                self.logger.warning(f"Post-swap validation issue: {validation_msg}")
                # Don't fail the swap, but log warning

            # Update mapping status
            mapping.status = SwapStatus.SUCCESS

            # ========== Capture ALL values before any modifications ==========
            # What we're swapping TO (from target)
            new_server = mapping.target.server
            new_database = mapping.target.database
            new_is_cloud = (mapping.target.target_type == "cloud")
            new_display_name = mapping.target.display_name
            new_workspace_name = mapping.target.workspace_name

            # What we're swapping FROM (from source - these become the new target)
            old_server = mapping.original_server
            old_database = mapping.original_database
            old_is_cloud = mapping.source.is_cloud
            old_workspace_name = mapping.source.workspace_name
            old_workspace_id = getattr(mapping.source, 'workspace_id', None)
            old_dataset_id = getattr(mapping.source, 'dataset_id', None)
            old_dataset_name = mapping.source.dataset_name

            # Build display name for the old source (becomes new target)
            old_display_name = old_dataset_name or old_database
            if old_workspace_name:
                old_display_name = f"{old_display_name} ({old_workspace_name})"
            elif not old_is_cloud and ':' in old_server:
                # For local, extract port from server for display
                port = old_server.split(':')[-1]
                old_display_name = f"{old_display_name} ({port})"

            # ========== Now apply the updates ==========
            # Update source to reflect what we swapped TO
            mapping.source.connection_string = new_conn_string
            mapping.source.server = new_server
            mapping.source.database = new_database
            mapping.source.is_cloud = new_is_cloud
            mapping.source.workspace_name = new_workspace_name
            # Extract dataset name from display name (remove workspace suffix)
            mapping.source.dataset_name = new_display_name.split(' (')[0] if ' (' in new_display_name else new_display_name

            # Flip target to point to old source (so user can swap back)
            mapping.target = SwapTarget(
                target_type="cloud" if old_is_cloud else "local",
                server=old_server,
                database=old_database,
                display_name=old_display_name,
                workspace_name=old_workspace_name,
                workspace_id=old_workspace_id,
                dataset_id=old_dataset_id,
            )

            # Since we flipped the target, status should be READY (ready to swap back)
            mapping.status = SwapStatus.READY

            elapsed = int((time.time() - start_time) * 1000)

            self.logger.info(f"Successfully swapped connection in {elapsed}ms")

            return SwapResult(
                success=True,
                mapping=mapping,
                message=f"Swapped to {new_display_name}",
                elapsed_ms=elapsed,
            )

        except Exception as e:
            self.logger.error(f"Failed to swap connection: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            mapping.status = SwapStatus.ERROR
            mapping.error_message = str(e)

            elapsed = int((time.time() - start_time) * 1000)

            return SwapResult(
                success=False,
                mapping=mapping,
                message=f"Swap failed: {e}",
                elapsed_ms=elapsed,
            )

    def _refresh_tom_reference(self, source: DataSourceConnection):
        """
        Re-fetch TOM reference to avoid stale COM objects.

        Args:
            source: DataSourceConnection with name to look up

        Returns:
            Fresh TOM reference or None if not found
        """
        try:
            model = self.connector._model
            ref_type = source.tom_reference_type or TomReferenceType.UNKNOWN

            if ref_type == TomReferenceType.EXPRESSION_SOURCE:
                # Look for named expression in model.Expressions
                if hasattr(model, 'Expressions') and model.Expressions:
                    for expr in model.Expressions:
                        if hasattr(expr, 'Name') and str(expr.Name) == source.name:
                            self.logger.debug(f"Refreshed ExpressionSource reference for '{source.name}'")
                            return expr
            else:
                # Look in DataSources collection
                if hasattr(model, 'DataSources') and model.DataSources:
                    for ds in model.DataSources:
                        if hasattr(ds, 'Name') and str(ds.Name) == source.name:
                            self.logger.debug(f"Refreshed DataSource reference for '{source.name}'")
                            return ds

            # Fallback: return existing reference
            self.logger.debug(f"Could not refresh TOM reference, using existing")
            return source.tom_datasource_ref

        except Exception as e:
            self.logger.error(f"Error refreshing TOM reference: {e}")
            return source.tom_datasource_ref

    def _swap_standard_datasource(self, datasource, new_conn_string: str) -> Tuple[bool, str]:
        """
        Swap a standard DataSource by setting ConnectionString.

        Args:
            datasource: TOM DataSource object
            new_conn_string: New connection string to set

        Returns:
            Tuple of (success, message)
        """
        try:
            if not hasattr(datasource, 'ConnectionString'):
                return False, "DataSource does not have ConnectionString property"

            old_value = str(datasource.ConnectionString) if datasource.ConnectionString else "(empty)"
            self.logger.debug(f"Setting ConnectionString from '{old_value[:50]}...' to '{new_conn_string[:50]}...'")

            datasource.ConnectionString = new_conn_string

            return True, "ConnectionString updated"

        except Exception as e:
            return False, f"Failed to set ConnectionString: {e}"

    def _swap_expression_source(self, expr_source, mapping: ConnectionMapping, new_conn_string: str) -> Tuple[bool, str]:
        """
        Swap an ExpressionSource by modifying the M expression.

        Replaces server and database in AnalysisServices.Database() call.

        Args:
            expr_source: TOM NamedExpression object
            mapping: ConnectionMapping with target info
            new_conn_string: New connection string (for reference)

        Returns:
            Tuple of (success, message)
        """
        try:
            if not hasattr(expr_source, 'Expression'):
                return False, "ExpressionSource does not have Expression property"

            old_expression = str(expr_source.Expression) if expr_source.Expression else ""
            if not old_expression:
                return False, "ExpressionSource has empty Expression"

            self.logger.debug(f"Original M expression: {old_expression[:200]}...")

            # Build new M expression by replacing server and database
            new_expression = self._build_new_m_expression(
                old_expression,
                mapping.target.server,
                mapping.target.database
            )

            if new_expression == old_expression:
                # Check if pattern was found but resulted in no change (same target)
                pattern = r'AnalysisServices\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'
                match = re.search(pattern, old_expression)
                if match:
                    current_server = match.group(1)
                    current_db = match.group(2)
                    if current_server == mapping.target.server and current_db == mapping.target.database:
                        return False, f"Target matches current connection (already at {mapping.target.display_name})"
                    else:
                        return False, f"Expression unchanged after replacement (unexpected)"
                return False, "Could not modify M expression - pattern not found"

            self.logger.debug(f"New M expression: {new_expression[:200]}...")

            expr_source.Expression = new_expression

            # Update the stored M expression in the mapping
            mapping.source.m_expression = new_expression

            return True, "M expression updated"

        except Exception as e:
            return False, f"Failed to update M expression: {e}"

    def _build_new_m_expression(self, expression: str, new_server: str, new_database: str) -> str:
        """
        Build a new M expression with updated server and database.

        Args:
            expression: Original M expression
            new_server: New server/endpoint
            new_database: New database/dataset name

        Returns:
            Modified M expression
        """
        # Pattern: AnalysisServices.Database("server", "database")
        # or: AnalysisServices.Database("server", "database", [options])
        pattern = r'(AnalysisServices\.Database\s*\(\s*)"([^"]+)"\s*,\s*"([^"]+)"'

        def replacer(match):
            prefix = match.group(1)
            return f'{prefix}"{new_server}", "{new_database}"'

        new_expression = re.sub(pattern, replacer, expression)

        return new_expression

    def _swap_structured_datasource(self, datasource, mapping: ConnectionMapping) -> Tuple[bool, str]:
        """
        Swap a StructuredDataSource by modifying ConnectionDetails.

        Args:
            datasource: TOM StructuredDataSource object
            mapping: ConnectionMapping with target info

        Returns:
            Tuple of (success, message)
        """
        try:
            if not hasattr(datasource, 'ConnectionDetails'):
                return False, "StructuredDataSource does not have ConnectionDetails property"

            details = datasource.ConnectionDetails
            if not details:
                return False, "ConnectionDetails is null"

            if not hasattr(details, 'Address'):
                return False, "ConnectionDetails does not have Address property"

            address = details.Address
            if not address:
                return False, "ConnectionDetails.Address is null"

            # Update server and database
            if hasattr(address, 'Server'):
                old_server = str(address.Server) if address.Server else "(empty)"
                self.logger.debug(f"Setting Address.Server from '{old_server}' to '{mapping.target.server}'")
                address.Server = mapping.target.server

            if hasattr(address, 'Database'):
                old_db = str(address.Database) if address.Database else "(empty)"
                self.logger.debug(f"Setting Address.Database from '{old_db}' to '{mapping.target.database}'")
                address.Database = mapping.target.database

            return True, "ConnectionDetails updated"

        except Exception as e:
            return False, f"Failed to update ConnectionDetails: {e}"

    def _validate_swap_persisted(self, mapping: ConnectionMapping, expected_conn_string: str) -> Tuple[bool, str]:
        """
        Validate that the swap actually persisted by reading back the value.

        Args:
            mapping: The mapping that was swapped
            expected_conn_string: The connection string we tried to set

        Returns:
            Tuple of (is_valid, message)
        """
        try:
            ref_type = mapping.source.tom_reference_type or TomReferenceType.UNKNOWN
            tom_ref = self._refresh_tom_reference(mapping.source)

            if tom_ref is None:
                return False, "Could not find TOM reference to validate"

            if ref_type == TomReferenceType.EXPRESSION_SOURCE:
                # Check Expression property
                if hasattr(tom_ref, 'Expression'):
                    current = str(tom_ref.Expression) if tom_ref.Expression else ""
                    # Check if new server/database appear in expression
                    if mapping.target.server in current and mapping.target.database in current:
                        return True, "M expression validated"
                    else:
                        return False, f"M expression doesn't contain expected values"
            elif ref_type == TomReferenceType.STRUCTURED_DATASOURCE:
                # Check ConnectionDetails
                if hasattr(tom_ref, 'ConnectionDetails') and tom_ref.ConnectionDetails:
                    addr = tom_ref.ConnectionDetails.Address
                    if addr and hasattr(addr, 'Server'):
                        if str(addr.Server) == mapping.target.server:
                            return True, "ConnectionDetails validated"
                        else:
                            return False, f"Server mismatch: expected {mapping.target.server}, got {addr.Server}"
            else:
                # Check ConnectionString
                if hasattr(tom_ref, 'ConnectionString'):
                    current = str(tom_ref.ConnectionString) if tom_ref.ConnectionString else ""
                    if current == expected_conn_string:
                        return True, "ConnectionString validated"
                    else:
                        return False, f"ConnectionString mismatch: expected '{expected_conn_string[:50]}...', got '{current[:50]}...'"

            return True, "Validation skipped (unknown type)"

        except Exception as e:
            return False, f"Validation error: {e}"

    def rollback_connection(self, mapping: ConnectionMapping) -> SwapResult:
        """
        Rollback a connection to its original state.

        Handles different TOM reference types appropriately.

        Args:
            mapping: The mapping to rollback

        Returns:
            SwapResult with rollback status
        """
        start_time = time.time()
        ref_type = mapping.source.tom_reference_type or TomReferenceType.UNKNOWN

        # Check if we have original values to rollback to
        if ref_type == TomReferenceType.EXPRESSION_SOURCE:
            # For M expressions, we need the original M expression stored
            # Check if m_expression has the original (before swap)
            if not mapping.original_connection_string:
                return SwapResult(
                    success=False,
                    mapping=mapping,
                    message="No original M expression stored for rollback",
                    elapsed_ms=0,
                )
        else:
            if not mapping.original_connection_string:
                return SwapResult(
                    success=False,
                    mapping=mapping,
                    message="No original connection string stored for rollback",
                    elapsed_ms=0,
                )

        if not mapping.source.tom_datasource_ref:
            return SwapResult(
                success=False,
                mapping=mapping,
                message="No TOM reference available",
                elapsed_ms=0,
            )

        try:
            # Re-fetch TOM reference to avoid stale COM objects
            tom_ref = self._refresh_tom_reference(mapping.source)
            if tom_ref is None:
                return SwapResult(
                    success=False,
                    mapping=mapping,
                    message="Could not refresh TOM reference for rollback",
                    elapsed_ms=int((time.time() - start_time) * 1000),
                )

            # Rollback based on reference type
            if ref_type == TomReferenceType.EXPRESSION_SOURCE:
                # Parse original connection info from stored connection string
                # and rebuild original M expression
                original_parsed = self._parse_connection_string(mapping.original_connection_string)
                old_server = original_parsed.get('server', '')
                old_database = original_parsed.get('database', '')

                if hasattr(tom_ref, 'Expression') and tom_ref.Expression:
                    current_expr = str(tom_ref.Expression)
                    # Build rolled back expression with original values
                    rolled_back_expr = self._build_new_m_expression(
                        current_expr, old_server, old_database
                    )
                    tom_ref.Expression = rolled_back_expr
                    self.logger.info(f"Rolled back M expression to {old_server}/{old_database}")
            elif ref_type == TomReferenceType.STRUCTURED_DATASOURCE:
                # Parse and rollback ConnectionDetails
                original_parsed = self._parse_connection_string(mapping.original_connection_string)
                if hasattr(tom_ref, 'ConnectionDetails') and tom_ref.ConnectionDetails:
                    addr = tom_ref.ConnectionDetails.Address
                    if addr:
                        if hasattr(addr, 'Server'):
                            addr.Server = original_parsed.get('server', '')
                        if hasattr(addr, 'Database'):
                            addr.Database = original_parsed.get('database', '')
            else:
                # Standard DataSource - set ConnectionString
                tom_ref.ConnectionString = mapping.original_connection_string

            self.connector._model.SaveChanges()

            # Reset mapping state
            mapping.source.connection_string = mapping.original_connection_string
            mapping.status = SwapStatus.READY
            mapping.error_message = None

            elapsed = int((time.time() - start_time) * 1000)

            return SwapResult(
                success=True,
                mapping=mapping,
                message="Connection rolled back to original",
                elapsed_ms=elapsed,
            )

        except Exception as e:
            self.logger.error(f"Failed to rollback connection: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            elapsed = int((time.time() - start_time) * 1000)

            return SwapResult(
                success=False,
                mapping=mapping,
                message=f"Rollback failed: {e}",
                elapsed_ms=elapsed,
            )

    def _parse_connection_string(self, connection_string: str) -> dict:
        """
        Parse a connection string into its components.

        Args:
            connection_string: MSOLAP or other connection string

        Returns:
            Dict with server, database, provider keys
        """
        result = {'server': '', 'database': '', 'provider': ''}

        if not connection_string:
            return result

        patterns = {
            'provider': r'Provider\s*=\s*([^;]+)',
            'server': r'Data Source\s*=\s*([^;]+)',
            'database': r'Initial Catalog\s*=\s*([^;]+)',
        }

        for key, pattern in patterns.items():
            match = re.search(pattern, connection_string, re.IGNORECASE)
            if match:
                result[key] = match.group(1).strip()

        return result

    def validate_target_accessible(self, target: SwapTarget) -> Tuple[bool, str]:
        """
        Validate that the target connection is accessible.

        Args:
            target: SwapTarget to validate

        Returns:
            Tuple of (is_valid, message)
        """
        if target.target_type == "local":
            return self._validate_local_target(target)
        else:
            return self._validate_cloud_target(target)

    def _validate_local_target(self, target: SwapTarget) -> Tuple[bool, str]:
        """
        Validate a local Power BI Desktop target.

        Args:
            target: Local SwapTarget

        Returns:
            Tuple of (is_valid, message)
        """
        try:
            # Parse server address (localhost:port)
            if ':' in target.server:
                host, port_str = target.server.rsplit(':', 1)
                port = int(port_str)
            else:
                return False, f"Invalid server format: {target.server} (expected localhost:port)"

            # Try to connect to the port
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(2)
            result = sock.connect_ex((host, port))
            sock.close()

            if result == 0:
                return True, f"Local model accessible at {target.server}"
            else:
                return False, f"Cannot connect to {target.server} (is Power BI Desktop running?)"

        except ValueError as e:
            return False, f"Invalid port number: {e}"
        except Exception as e:
            return False, f"Validation error: {e}"

    def _validate_cloud_target(self, target: SwapTarget) -> Tuple[bool, str]:
        """
        Validate a cloud XMLA endpoint target.

        Note: Full validation requires authentication. This does basic format check.

        Args:
            target: Cloud SwapTarget

        Returns:
            Tuple of (is_valid, message)
        """
        # Basic format validation
        if not target.server:
            return False, "No XMLA endpoint specified"

        if 'powerbi://' not in target.server.lower():
            return False, f"Invalid XMLA endpoint format: {target.server}"

        if not target.database:
            return False, "No dataset name specified"

        # Full validation would require trying to connect
        # For now, assume format is valid
        return True, f"Cloud endpoint format valid: {target.workspace_name}/{target.database}"

    @staticmethod
    def build_local_connection_string(server: str, database: str) -> str:
        """
        Build connection string for local Power BI Desktop target.

        Args:
            server: Server address (localhost:port)
            database: Database name/GUID

        Returns:
            MSOLAP connection string
        """
        return f"Provider=MSOLAP;Data Source={server};Initial Catalog={database}"

    @staticmethod
    def build_cloud_connection_string(
        workspace_name: str,
        dataset_name: str,
        tenant: str = "myorg"
    ) -> str:
        """
        Build connection string for cloud XMLA endpoint target.

        Args:
            workspace_name: Power BI workspace name
            dataset_name: Dataset/semantic model name
            tenant: Tenant name (default: myorg)

        Returns:
            MSOLAP connection string for XMLA endpoint
        """
        # URL-encode workspace name for special characters
        encoded_workspace = quote(workspace_name)
        xmla_endpoint = f"powerbi://api.powerbi.com/v1.0/{tenant}/{encoded_workspace}"

        return f"Provider=MSOLAP;Data Source={xmla_endpoint};Initial Catalog={dataset_name}"

    def create_local_target(
        self,
        server: str,
        database: str,
        display_name: str
    ) -> SwapTarget:
        """
        Create a SwapTarget for a local Power BI Desktop model.

        Args:
            server: Server address (localhost:port)
            database: Database name/GUID
            display_name: Friendly display name

        Returns:
            SwapTarget configured for local connection
        """
        return SwapTarget(
            target_type="local",
            server=server,
            database=database,
            display_name=display_name,
        )

    def create_cloud_target(
        self,
        workspace_name: str,
        workspace_id: str,
        dataset_name: str,
        dataset_id: str,
        tenant: str = "myorg"
    ) -> SwapTarget:
        """
        Create a SwapTarget for a cloud Power BI Service model.

        Args:
            workspace_name: Workspace display name
            workspace_id: Workspace GUID
            dataset_name: Dataset display name
            dataset_id: Dataset GUID
            tenant: Tenant name (default: myorg)

        Returns:
            SwapTarget configured for cloud connection
        """
        encoded_workspace = quote(workspace_name)
        xmla_endpoint = f"powerbi://api.powerbi.com/v1.0/{tenant}/{encoded_workspace}"

        return SwapTarget(
            target_type="cloud",
            server=xmla_endpoint,
            database=dataset_name,
            display_name=f"{dataset_name} ({workspace_name})",
            workspace_name=workspace_name,
            workspace_id=workspace_id,
            dataset_id=dataset_id,
        )
