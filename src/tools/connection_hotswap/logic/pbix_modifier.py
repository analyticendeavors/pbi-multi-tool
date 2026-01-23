"""
PBIX/PBIP File Modifier for Thin Report Connection Swapping
Built by Reid Havens of Analytic Endeavors

Handles modification of .pbix and .pbip files to swap Analysis Services connections
for thin reports (live-connected Power BI reports).

PBIX files are ZIP archives containing a 'Connections' file.
PBIP files are Power BI Project format with loose JSON files including 'definition.pbir'.
"""

import json
import logging
import os
import random
import shutil
import zipfile
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, Tuple

from tools.connection_hotswap.logic.connection_cache import get_cache_manager

# Ownership fingerprint
_AE_FP = "UGJpeE1vZGlmaWVyOkFFLTIwMjQ="


@dataclass
class SwapResult:
    """Result of a connection swap operation."""
    success: bool
    message: str
    backup_path: Optional[str] = None
    file_type: Optional[str] = None  # 'pbix' or 'pbip'


# Cache for original cloud connections, keyed by normalized file path
_original_cloud_connections: dict = {}


class PbixModifier:
    """
    Handles modification of .pbix and .pbip files for connection swapping.

    Key differences between formats:
    - PBIX: ZIP archive, file is LOCKED while open in Power BI Desktop
    - PBIP: Loose JSON files, files are NOT locked while open (but need close/reopen to apply)
    """

    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def detect_file_type(self, file_path: str) -> Optional[str]:
        """
        Detect whether a file is a PBIX or PBIP.

        Args:
            file_path: Path to the Power BI file

        Returns:
            'pbix', 'pbip', or None if unknown/not found
        """
        if not file_path or not os.path.exists(file_path):
            return None

        lower_path = file_path.lower()
        if lower_path.endswith('.pbix'):
            return 'pbix'
        elif lower_path.endswith('.pbip'):
            return 'pbip'

        return None

    def backup_file(self, file_path: str) -> Tuple[bool, Optional[str]]:
        """
        Create a backup of the file before modification.

        Args:
            file_path: Path to the file to backup

        Returns:
            Tuple of (success, backup_path)
        """
        try:
            if not os.path.exists(file_path):
                self.logger.error(f"File not found for backup: {file_path}")
                return False, None

            # Create backup with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(file_path)[0]
            extension = os.path.splitext(file_path)[1]
            backup_path = f"{base_name}_backup_{timestamp}{extension}"

            # For PBIP, we need to backup both the .pbip file and the .Report folder
            file_type = self.detect_file_type(file_path)
            if file_type == 'pbip':
                # Backup the definition.pbir file specifically
                report_folder = self._get_report_folder(file_path)
                if report_folder:
                    pbir_path = os.path.join(report_folder, 'definition.pbir')
                    if os.path.exists(pbir_path):
                        backup_path = f"{os.path.splitext(pbir_path)[0]}_backup_{timestamp}.pbir"
                        shutil.copy2(pbir_path, backup_path)
                        self.logger.info(f"Created backup: {backup_path}")
                        return True, backup_path
                    else:
                        self.logger.error(f"definition.pbir not found in {report_folder}")
                        return False, None
            else:
                # Backup PBIX file
                shutil.copy2(file_path, backup_path)
                self.logger.info(f"Created backup: {backup_path}")
                return True, backup_path

        except Exception as e:
            self.logger.error(f"Error creating backup: {e}")
            return False, None

    def _get_report_folder(self, pbip_path: str) -> Optional[str]:
        """
        Get the .Report folder path from a .pbip file path.

        PBIP files have a companion .Report folder with the same base name.
        E.g., 'MyReport.pbip' -> 'MyReport.Report/'

        Args:
            pbip_path: Path to the .pbip file

        Returns:
            Path to the .Report folder, or None if not found
        """
        if not pbip_path.lower().endswith('.pbip'):
            return None

        # Replace .pbip with .Report
        report_folder = pbip_path[:-5] + '.Report'

        if os.path.isdir(report_folder):
            return report_folder

        # Try case-insensitive match (Windows)
        parent_dir = os.path.dirname(pbip_path)
        base_name = os.path.basename(pbip_path)[:-5]  # Remove .pbip

        try:
            for item in os.listdir(parent_dir):
                if item.lower() == f"{base_name.lower()}.report":
                    report_folder = os.path.join(parent_dir, item)
                    if os.path.isdir(report_folder):
                        return report_folder
        except Exception as e:
            self.logger.debug(f"Error searching for .Report folder: {e}")

        return None

    def swap_connection(self, file_path: str, server: str, database: str,
                        create_backup: bool = True, dataset_id: str = None,
                        is_cloud_target: bool = None,
                        use_cached_cloud: bool = True,
                        source_friendly_name: str = None,
                        perspective_name: str = None,
                        workspace_name: str = None,
                        cloud_connection_type: str = None) -> SwapResult:
        """
        Swap the connection in a PBIX or PBIP file to a target model.

        Automatically caches the original cloud connection (if present) before
        swapping, so it can be restored later using restore_cloud_connection().

        For cloud targets (Power BI Service/Fabric):
        - If we have a cached original cloud connection, restores EXACTLY that format
        - Otherwise, creates connection based on cloud_connection_type:
          - pbi_semantic_model: pbiServiceLive (no perspective) or pbiServiceXmlaStyleLive (with perspective)
          - aas_xmla: analysisServicesDatabaseLive with XMLA endpoint
        For local targets, creates analysisServicesDatabaseLive.

        Args:
            file_path: Path to the .pbix or .pbip file
            server: Target server address (e.g., 'localhost:52841' or 'powerbi://...')
            database: Target database/model identifier (e.g., GUID or friendly name)
            create_backup: Whether to create a backup before modification
            dataset_id: Optional dataset GUID for cloud connections (PbiServiceModelId)
            is_cloud_target: Explicitly set if target is cloud; auto-detects if None
            use_cached_cloud: If True and swapping to cloud, use cached original connection
            source_friendly_name: Optional friendly name of the source model (for GUID resolution)
            perspective_name: Optional perspective/cube name to connect to (e.g., "Finance View")
            workspace_name: Workspace name for cloud targets with perspectives (required for XMLA URL)
            cloud_connection_type: Type of cloud connection ('pbi_semantic_model' or 'aas_xmla')

        Returns:
            SwapResult with success status and message
        """
        file_type = self.detect_file_type(file_path)

        if not file_type:
            return SwapResult(
                success=False,
                message=f"Unknown or unsupported file type: {file_path}",
                file_type=None
            )

        # Auto-detect cloud target if not explicitly set
        if is_cloud_target is None:
            is_cloud_target = self._is_cloud_server(server)

        # Read current connection BEFORE modifying - cache if it's a cloud connection
        # This allows restoring to the original cloud connection later
        current_connection = self.read_current_connection(file_path)
        if current_connection and self.is_cloud_connection(current_connection):
            # Add friendly name to cache if provided (for GUID resolution later)
            if source_friendly_name:
                current_connection['friendly_name'] = source_friendly_name
            self.cache_original_cloud_connection(file_path, current_connection)
            self.logger.info("Cached original cloud connection before swapping to local")

        # Create backup if requested
        backup_path = None
        if create_backup:
            backup_success, backup_path = self.backup_file(file_path)
            if not backup_success:
                self.logger.warning("Backup failed, proceeding without backup")

        # For cloud targets, try to use cached original connection to preserve exact format
        # BUT only if we're restoring to the SAME cloud model (matching database GUID)
        # AND the same connection type (Semantic Model vs XMLA)
        if is_cloud_target and use_cached_cloud:
            cached_connection = self.get_cached_cloud_connection(file_path)
            if cached_connection:
                # Check if target database matches the cached original
                cached_database = cached_connection.get('database', '').lower()
                target_database_lower = (database or '').lower()

                # Check if connection type matches
                cached_conn_type = cached_connection.get('ConnectionType', '').lower()
                connection_type_matches = True  # Default to true if no type specified

                if cloud_connection_type:
                    # Map requested type to expected ConnectionType values
                    if cloud_connection_type == 'pbi_semantic_model':
                        # Semantic Model uses pbiServiceLive or pbiServiceXmlaStyleLive
                        connection_type_matches = cached_conn_type in ['pbiserviceLive'.lower(), 'pbiservicexmlastylelive']
                    elif cloud_connection_type == 'aas_xmla':
                        # XMLA uses analysisServicesDatabaseLive
                        connection_type_matches = cached_conn_type == 'analysisservicesdatabaselive'

                if cached_database and target_database_lower and cached_database == target_database_lower and connection_type_matches:
                    self.logger.info("Using cached original cloud connection for restoration (target matches cached)")
                    result = self.restore_cloud_connection(file_path, cached_connection, create_backup=False)
                    result.backup_path = backup_path
                    result.file_type = file_type
                    return result
                else:
                    if not connection_type_matches:
                        self.logger.info(f"Target connection type ({cloud_connection_type}) differs from cached ({cached_conn_type}) - building new connection")
                    else:
                        self.logger.info(f"Target database ({database}) differs from cached ({cached_connection.get('database')}) - building new connection")

        # Dispatch to appropriate method (new connection generation)
        if file_type == 'pbix':
            result = self._swap_connection_pbix(file_path, server, database, is_cloud_target, dataset_id, perspective_name, workspace_name, cloud_connection_type)
        else:
            result = self._swap_connection_pbip(file_path, server, database, is_cloud_target, dataset_id, perspective_name, workspace_name, cloud_connection_type)

        # Add warning for generic cloud connections (no cached schema available)
        if is_cloud_target and result.success:
            self.logger.warning("Generated generic cloud connection (no cached schema available)")
            result.message += (
                "\n\nNote: A generic cloud connection was created because no cached "
                "original connection schema was available. Some Power BI features "
                "(like OneLake catalog navigation) may not work as expected. "
                "Consider re-saving any presets after confirming the connection works."
            )

        result.backup_path = backup_path
        result.file_type = file_type
        return result

    def _is_cloud_server(self, server: str) -> bool:
        """Check if a server address is a cloud endpoint."""
        if not server:
            return False
        server_lower = server.lower()
        return ('powerbi://' in server_lower or
                'pbiazure://' in server_lower or
                'asazure://' in server_lower)

    def restore_cloud_connection(self, file_path: str, original_connection_info: dict,
                                  create_backup: bool = True) -> SwapResult:
        """
        Restore a cloud connection using the original connection format.

        This preserves the pbiServiceLive connection type and all associated
        properties (PbiServiceModelId, PbiModelVirtualServerName, etc.) which
        gives users the OneLake catalog navigation experience instead of a
        raw connection string dialog.

        Args:
            file_path: Path to the .pbix or .pbip file
            original_connection_info: Dict containing '_original_pbix_connection' or
                                     '_original_pbip_definition' from read_current_connection()
            create_backup: Whether to create a backup before modification

        Returns:
            SwapResult with success status and message
        """
        file_type = self.detect_file_type(file_path)

        if not file_type:
            return SwapResult(
                success=False,
                message=f"Unknown or unsupported file type: {file_path}",
                file_type=None
            )

        # Create backup if requested
        backup_path = None
        if create_backup:
            backup_success, backup_path = self.backup_file(file_path)
            if not backup_success:
                self.logger.warning("Backup failed, proceeding without backup")

        # Dispatch to appropriate method
        if file_type == 'pbix':
            result = self._restore_cloud_connection_pbix(file_path, original_connection_info)
        else:
            result = self._restore_cloud_connection_pbip(file_path, original_connection_info)

        result.backup_path = backup_path
        result.file_type = file_type
        return result

    def _restore_cloud_connection_pbix(self, pbix_path: str, original_info: dict) -> SwapResult:
        """
        Restore original cloud connection in a PBIX file.

        Preserves pbiServiceLive connection type with all Power BI Service properties.
        """
        try:
            # Get original connection data
            original_conn = original_info.get('_original_pbix_connection')
            original_version = original_info.get('_original_pbix_version', 2)

            if not original_conn:
                return SwapResult(
                    success=False,
                    message="No original PBIX connection data available for restoration"
                )

            # Check if file is accessible (not locked)
            try:
                with open(pbix_path, 'r+b'):
                    pass
            except (IOError, PermissionError) as e:
                return SwapResult(
                    success=False,
                    message=f"File is locked (Power BI Desktop may be open): {e}"
                )

            # Build the Connections content with original format
            connections_json = {
                "Version": original_version,
                "Connections": [original_conn]
            }

            connections_content = json.dumps(connections_json, indent=2)

            # Read existing ZIP, modify, and write back
            temp_path = pbix_path + '.tmp'

            with zipfile.ZipFile(pbix_path, 'r') as zip_read:
                with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zip_write:
                    for item in zip_read.infolist():
                        item_name_lower = item.filename.lower()

                        # Skip SecurityBindings if it exists
                        if item_name_lower == 'securitybindings':
                            self.logger.info("Removing SecurityBindings from PBIX")
                            continue

                        # Replace Connections file
                        if item_name_lower == 'connections':
                            self.logger.info("Restoring original cloud Connections in PBIX")
                            zip_write.writestr(item.filename, connections_content)
                        else:
                            data = zip_read.read(item.filename)
                            zip_write.writestr(item, data)

                    # If Connections file didn't exist, create it
                    existing_files = [f.filename.lower() for f in zip_read.infolist()]
                    if 'connections' not in existing_files:
                        self.logger.info("Creating new Connections file in PBIX")
                        zip_write.writestr('Connections', connections_content)

            # Replace original with modified file
            os.replace(temp_path, pbix_path)

            conn_type = original_conn.get('ConnectionType', 'cloud')
            self.logger.info(f"Successfully restored cloud connection ({conn_type}): {pbix_path}")
            return SwapResult(
                success=True,
                message=f"Cloud connection restored ({conn_type}). Close and reopen the file to apply changes."
            )

        except Exception as e:
            self.logger.error(f"Error restoring cloud connection in PBIX: {e}")
            temp_path = pbix_path + '.tmp'
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass
            return SwapResult(
                success=False,
                message=f"Error restoring cloud connection: {e}"
            )

    def _restore_cloud_connection_pbip(self, pbip_path: str, original_info: dict) -> SwapResult:
        """
        Restore original cloud connection in a PBIP file.

        Preserves the original definition.pbir format with all cloud properties.
        """
        try:
            # Get original PBIP definition
            original_def = original_info.get('_original_pbip_definition')

            if not original_def:
                return SwapResult(
                    success=False,
                    message="No original PBIP definition data available for restoration"
                )

            # Find the .Report folder
            report_folder = self._get_report_folder(pbip_path)
            if not report_folder:
                return SwapResult(
                    success=False,
                    message=f"Could not find .Report folder for: {pbip_path}"
                )

            # Find definition.pbir
            pbir_path = os.path.join(report_folder, 'definition.pbir')
            if not os.path.exists(pbir_path):
                return SwapResult(
                    success=False,
                    message=f"definition.pbir not found in: {report_folder}"
                )

            # Write the original content back
            with open(pbir_path, 'w', encoding='utf-8') as f:
                json.dump(original_def, f, indent=2)

            self.logger.info(f"Successfully restored cloud connection in PBIP: {pbir_path}")
            return SwapResult(
                success=True,
                message="Cloud connection restored. Close and reopen the file to apply changes."
            )

        except Exception as e:
            self.logger.error(f"Error restoring cloud connection in PBIP: {e}")
            return SwapResult(
                success=False,
                message=f"Error restoring cloud connection: {e}"
            )

    def is_cloud_connection(self, connection_info: dict) -> bool:
        """
        Determine if a connection is a cloud (Power BI Service/Fabric) connection.

        Cloud connections use:
        - pbiServiceLive connection type
        - pbiazure:// or powerbi:// data source URLs

        Args:
            connection_info: Connection dict from read_current_connection()

        Returns:
            True if this is a Power BI Service/Fabric cloud connection
        """
        if not connection_info:
            return False

        conn_type = connection_info.get('connection_type', '').lower()
        server = connection_info.get('server', '').lower()

        # Check for pbiServiceLive connection type
        if conn_type == 'pbiserviceLive'.lower():
            return True

        # Check for Power BI Service URLs
        if 'pbiazure://' in server or 'powerbi://api.powerbi.com' in server:
            return True

        # Check if original data indicates cloud connection
        original_pbix = connection_info.get('_original_pbix_connection', {})
        if original_pbix.get('ConnectionType', '').lower() == 'pbiserviceLive'.lower():
            return True

        return False

    def cache_original_cloud_connection(self, file_path: str, connection_info: dict):
        """
        Cache an original cloud connection for later restoration.

        This should be called BEFORE swapping to a local connection, so the
        original cloud connection format can be restored later.

        Caches to both:
        - In-memory dict (fast, session-only)
        - Persistent disk cache (survives app restart)

        Args:
            file_path: Path to the .pbix or .pbip file
            connection_info: Connection dict from read_current_connection()
        """
        global _original_cloud_connections

        if not file_path or not connection_info:
            return

        # Only cache if it's actually a cloud connection
        if not self.is_cloud_connection(connection_info):
            self.logger.debug(f"Not caching non-cloud connection for {file_path}")
            return

        # Normalize path for consistent lookup
        normalized_path = os.path.normpath(file_path).lower()

        # Cache in memory (fast access)
        _original_cloud_connections[normalized_path] = connection_info.copy()
        self.logger.info(f"Cached original cloud connection for: {file_path}")

        # Also persist to disk (survives app restart)
        file_type = self.detect_file_type(file_path) or "pbix"
        disk_cache = get_cache_manager()
        disk_cache.save_connection(file_path, connection_info, file_type)
        self.logger.debug(f"Persisted cloud connection to disk cache")

    def get_cached_cloud_connection(self, file_path: str) -> Optional[dict]:
        """
        Retrieve a cached original cloud connection.

        Checks in order:
        1. In-memory cache (fast, current session)
        2. Persistent disk cache (survives app restart)

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            Cached connection info dict, or None if not cached
        """
        global _original_cloud_connections

        if not file_path:
            return None

        normalized_path = os.path.normpath(file_path).lower()

        # Check in-memory cache first (fastest)
        cached = _original_cloud_connections.get(normalized_path)
        if cached:
            return cached

        # Check persistent disk cache (survives app restart)
        disk_cache = get_cache_manager()
        disk_cached = disk_cache.load_connection(file_path)
        if disk_cached:
            # Populate memory cache for faster subsequent access
            _original_cloud_connections[normalized_path] = disk_cached
            self.logger.debug(f"Loaded cloud connection from disk cache: {file_path}")
            return disk_cached

        return None

    def has_cached_cloud_connection(self, file_path: str) -> bool:
        """
        Check if we have a cached original cloud connection for a file.

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            True if we have cached cloud connection data
        """
        return self.get_cached_cloud_connection(file_path) is not None

    def clear_cached_cloud_connection(self, file_path: str):
        """
        Clear cached cloud connection for a file.

        Clears from both in-memory and persistent disk cache.

        Args:
            file_path: Path to the .pbix or .pbip file
        """
        global _original_cloud_connections

        if not file_path:
            return

        normalized_path = os.path.normpath(file_path).lower()

        # Clear from memory cache
        if normalized_path in _original_cloud_connections:
            del _original_cloud_connections[normalized_path]
            self.logger.info(f"Cleared cached cloud connection for: {file_path}")

        # Clear from disk cache
        disk_cache = get_cache_manager()
        disk_cache.clear_connection(file_path)

    def _swap_connection_pbix(self, pbix_path: str, server: str, database: str,
                               is_cloud_target: bool = False, dataset_id: str = None,
                               perspective_name: str = None, workspace_name: str = None,
                               cloud_connection_type: str = None) -> SwapResult:
        """
        Modify a .pbix ZIP archive to swap the connection.

        The PBIX file contains a 'Connections' file with the connection info.

        Connection types:
        - Local AS: analysisServicesDatabaseLive (Version 1)
        - Cloud Semantic Model (no perspective): pbiServiceLive (Version 2)
        - Cloud Semantic Model (with perspective): pbiServiceXmlaStyleLive (Version 4)
        - Cloud XMLA (with perspective): analysisServicesDatabaseLive (Version 1)

        Args:
            pbix_path: Path to the .pbix file
            server: Target server address
            database: Target database identifier
            is_cloud_target: Whether the target is a cloud endpoint
            dataset_id: Optional dataset GUID for cloud connections
            perspective_name: Optional perspective/cube name to connect to
            workspace_name: Workspace name for cloud connections with perspectives
            cloud_connection_type: 'pbi_semantic_model' or 'aas_xmla'

        Returns:
            SwapResult with success status and message
        """
        import urllib.parse

        try:
            # Check if file is accessible (not locked)
            try:
                with open(pbix_path, 'r+b'):
                    pass  # Just checking if we can open it
            except (IOError, PermissionError) as e:
                return SwapResult(
                    success=False,
                    message=f"File is locked (Power BI Desktop may be open): {e}"
                )

            # Build the new Connections content based on target type
            if is_cloud_target:
                # Default to Semantic Model if not specified
                use_xmla = cloud_connection_type == 'aas_xmla'

                if perspective_name:
                    # Cloud target WITH perspective - connection type determines format
                    if not workspace_name:
                        return SwapResult(
                            success=False,
                            message="Workspace name is required for cloud connections with perspectives"
                        )

                    # URL-encode workspace name for connection string
                    encoded_workspace = urllib.parse.quote(workspace_name)

                    if use_xmla:
                        # XMLA Endpoint: analysisServicesDatabaseLive (Version 1)
                        # Requires Premium/PPU/Fabric capacity
                        # Initial Catalog needs quotes around database name for XMLA
                        conn_str = f'Data Source=powerbi://api.powerbi.com/v1.0/myorg/{encoded_workspace};Initial Catalog="{database}";Cube={perspective_name};Access Mode=readonly'
                        connection = {
                            "Name": "EntityDataSource",
                            "ConnectionString": conn_str,
                            "ConnectionType": "analysisServicesDatabaseLive"
                        }
                        connections_json = {
                            "Version": 1,
                            "Connections": [connection]
                        }
                        self.logger.info(f"Building analysisServicesDatabaseLive (XMLA) connection for cloud perspective: {database} [{perspective_name}]")
                    else:
                        # Semantic Model: pbiServiceXmlaStyleLive (Version 4)
                        # Works with Pro workspaces, uses XMLA-style URL but PBI connector
                        escaped_perspective = perspective_name.replace('"', '\\"')
                        conn_str = (
                            f"Data Source=powerbi://api.powerbi.com/v1.0/myorg/{encoded_workspace};"
                            f"Initial Catalog={database};"
                            f'Cube="{escaped_perspective}";'
                            f"Access Mode=readonly;Integrated Security=ClaimsToken"
                        )
                        connection = {
                            "Name": "EntityDataSource",
                            "ConnectionString": conn_str,
                            "ConnectionType": "pbiServiceXmlaStyleLive",
                            "PbiModelVirtualServerName": "sobe_wowvirtualserver",
                            "PbiModelDatabaseName": dataset_id or database
                        }
                        connections_json = {
                            "Version": 4,
                            "Connections": [connection]
                        }
                        self.logger.info(f"Building pbiServiceXmlaStyleLive connection for cloud perspective: {database} [{perspective_name}]")
                else:
                    # Cloud target WITHOUT perspective
                    if use_xmla:
                        # XMLA Endpoint WITHOUT perspective: analysisServicesDatabaseLive with Cube=Model
                        if not workspace_name:
                            return SwapResult(
                                success=False,
                                message="Workspace name is required for XMLA connections"
                            )

                        encoded_workspace = urllib.parse.quote(workspace_name)
                        # Initial Catalog needs quotes around database name for XMLA
                        conn_str = (
                            f'Data Source=powerbi://api.powerbi.com/v1.0/myorg/{encoded_workspace};'
                            f'Initial Catalog="{database}";'
                            f'Cube=Model;Access Mode=readonly'
                        )
                        connection = {
                            "Name": "EntityDataSource",
                            "ConnectionString": conn_str,
                            "ConnectionType": "analysisServicesDatabaseLive"
                        }
                        connections_json = {
                            "Version": 1,
                            "Connections": [connection]
                        }
                        self.logger.info(f"Building analysisServicesDatabaseLive (XMLA) connection for cloud target: {database}")
                    else:
                        # Semantic Model: use standard pbiServiceLive
                        # PbiServiceModelId needs to be an integer (Power BI crashes with null)
                        # Generate random 8-digit integer - actual value doesn't matter, just needs to be present
                        # PbiModelVirtualServerName must be "sobe_wowvirtualserver" for pbiServiceLive
                        random_model_id = random.randint(10000000, 99999999)
                        # pbiServiceLive requires dataset GUID for both Initial Catalog and PbiModelDatabaseName
                        # Use dataset_id if available, otherwise fall back to database (friendly name)
                        db_identifier = dataset_id or database
                        # Connection string format matching Power BI's native format
                        conn_str = (
                            f'Data Source=pbiazure://api.powerbi.com;'
                            f'Initial Catalog={db_identifier};'
                            f'Identity Provider="https://login.microsoftonline.com/common, '
                            f'https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc";'
                            f'Integrated Security=ClaimsToken'
                        )
                        connection = {
                            "Name": "EntityDataSource",
                            "ConnectionString": conn_str,
                            "ConnectionType": "pbiServiceLive",
                            "PbiServiceModelId": random_model_id,
                            "PbiModelVirtualServerName": "sobe_wowvirtualserver",
                            "PbiModelDatabaseName": db_identifier  # Dataset GUID required for pbiServiceLive
                        }
                        connections_json = {
                            "Version": 2,
                            "Connections": [connection]
                        }
                        self.logger.info(f"Building pbiServiceLive connection for cloud target: {database} (DB: {db_identifier}, ModelId: {random_model_id})")
            else:
                # Local target: analysisServicesDatabaseLive connection
                # Use specified perspective or default to "Model"
                cube_value = perspective_name if perspective_name else "Model"
                connections_json = {
                    "Version": 1,
                    "Connections": [{
                        "Name": "EntityDataSource",
                        "ConnectionString": f"Data Source={server};Initial Catalog={database};Cube={cube_value}",
                        "ConnectionType": "analysisServicesDatabaseLive"
                    }]
                }
                perspective_info = f" [{perspective_name}]" if perspective_name else ""
                self.logger.info(f"Building analysisServicesDatabaseLive connection for local target: {server}{perspective_info}")

            connections_content = json.dumps(connections_json, indent=2)

            # Read existing ZIP, modify, and write back
            # We need to rewrite the entire ZIP because Python's zipfile doesn't support in-place updates

            temp_path = pbix_path + '.tmp'

            with zipfile.ZipFile(pbix_path, 'r') as zip_read:
                with zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zip_write:
                    for item in zip_read.infolist():
                        item_name_lower = item.filename.lower()

                        # Skip SecurityBindings if it exists (can cause issues)
                        if item_name_lower == 'securitybindings':
                            self.logger.info("Removing SecurityBindings from PBIX")
                            continue

                        # Replace Connections file
                        if item_name_lower == 'connections':
                            self.logger.info("Updating Connections file in PBIX")
                            zip_write.writestr(item.filename, connections_content)
                        else:
                            # Copy other files as-is
                            data = zip_read.read(item.filename)
                            zip_write.writestr(item, data)

                    # If Connections file didn't exist, create it
                    existing_files = [f.filename.lower() for f in zip_read.infolist()]
                    if 'connections' not in existing_files:
                        self.logger.info("Creating new Connections file in PBIX")
                        zip_write.writestr('Connections', connections_content)

            # Replace original with modified file
            os.replace(temp_path, pbix_path)

            conn_type = "pbiServiceLive" if is_cloud_target else "analysisServicesDatabaseLive"
            self.logger.info(f"Successfully modified PBIX ({conn_type}): {pbix_path}")
            return SwapResult(
                success=True,
                message=f"Connection updated to {database}. Close and reopen the file to apply changes."
            )

        except Exception as e:
            self.logger.error(f"Error modifying PBIX: {e}")
            # Clean up temp file if it exists
            temp_path = pbix_path + '.tmp'
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass
            return SwapResult(
                success=False,
                message=f"Error modifying PBIX: {e}"
            )

    def _swap_connection_pbip(self, pbip_path: str, server: str, database: str,
                               is_cloud_target: bool = False, dataset_id: str = None,
                               perspective_name: str = None, workspace_name: str = None,
                               cloud_connection_type: str = None) -> SwapResult:
        """
        Modify a .pbip project's definition.pbir to swap the connection.

        The definition.pbir file controls the data source connection.

        Connection types:
        - Local AS: analysisServicesDatabaseLive
        - Cloud Semantic Model (no perspective): pbiServiceLive
        - Cloud Semantic Model (with perspective): pbiServiceXmlaStyleLive
        - Cloud XMLA (with perspective): analysisServicesDatabaseLive

        Args:
            pbip_path: Path to the .pbip file
            server: Target server address
            database: Target database identifier
            is_cloud_target: Whether the target is a cloud endpoint
            dataset_id: Optional dataset GUID for cloud connections
            perspective_name: Optional perspective/cube name to connect to
            workspace_name: Workspace name for cloud connections with perspectives
            cloud_connection_type: 'pbi_semantic_model' or 'aas_xmla'

        Returns:
            SwapResult with success status and message
        """
        import urllib.parse

        try:
            # Find the .Report folder
            report_folder = self._get_report_folder(pbip_path)
            if not report_folder:
                return SwapResult(
                    success=False,
                    message=f"Could not find .Report folder for: {pbip_path}"
                )

            # Find definition.pbir
            pbir_path = os.path.join(report_folder, 'definition.pbir')
            if not os.path.exists(pbir_path):
                return SwapResult(
                    success=False,
                    message=f"definition.pbir not found in: {report_folder}"
                )

            # Build the new definition.pbir content based on target type
            if is_cloud_target:
                # Default to Semantic Model if not specified
                use_xmla = cloud_connection_type == 'aas_xmla'

                if perspective_name:
                    # Cloud target WITH perspective - connection type determines format
                    if not workspace_name:
                        return SwapResult(
                            success=False,
                            message="Workspace name is required for cloud connections with perspectives"
                        )

                    # URL-encode workspace name for connection string
                    encoded_workspace = urllib.parse.quote(workspace_name)

                    if use_xmla:
                        # XMLA Endpoint: analysisServicesDatabaseLive
                        # Requires Premium/PPU/Fabric capacity
                        # Initial Catalog needs quotes around database name for XMLA
                        conn_str = f'Data Source=powerbi://api.powerbi.com/v1.0/myorg/{encoded_workspace};Initial Catalog="{database}";Cube={perspective_name};Access Mode=readonly'
                        by_connection = {
                            "connectionString": conn_str,
                            "connectionType": "analysisServicesDatabaseLive",
                            "name": "EntityDataSource",
                            "pbiServiceModelId": None,
                            "pbiModelVirtualServerName": None,
                            "pbiModelDatabaseName": None
                        }
                        self.logger.info(f"Building analysisServicesDatabaseLive (XMLA) PBIP connection for cloud perspective: {database} [{perspective_name}]")
                    else:
                        # Semantic Model: pbiServiceXmlaStyleLive
                        # Works with Pro workspaces, uses XMLA-style URL but PBI connector
                        escaped_perspective = perspective_name.replace('"', '\\"')
                        conn_str = (
                            f"Data Source=powerbi://api.powerbi.com/v1.0/myorg/{encoded_workspace};"
                            f"Initial Catalog={database};"
                            f'Cube="{escaped_perspective}";'
                            f"Access Mode=readonly;Integrated Security=ClaimsToken"
                        )
                        by_connection = {
                            "connectionString": conn_str,
                            "connectionType": "pbiServiceXmlaStyleLive",
                            "name": "EntityDataSource",
                            "pbiServiceModelId": None,
                            "pbiModelVirtualServerName": "sobe_wowvirtualserver",
                            "pbiModelDatabaseName": dataset_id or database
                        }
                        self.logger.info(f"Building pbiServiceXmlaStyleLive PBIP connection for cloud perspective: {database} [{perspective_name}]")
                else:
                    # Cloud target WITHOUT perspective
                    if use_xmla:
                        # XMLA Endpoint WITHOUT perspective: analysisServicesDatabaseLive with Cube=Model
                        if not workspace_name:
                            return SwapResult(
                                success=False,
                                message="Workspace name is required for XMLA connections"
                            )

                        encoded_workspace = urllib.parse.quote(workspace_name)
                        # Initial Catalog needs quotes around database name for XMLA
                        conn_str = (
                            f'Data Source=powerbi://api.powerbi.com/v1.0/myorg/{encoded_workspace};'
                            f'Initial Catalog="{database}";'
                            f'Cube=Model;Access Mode=readonly'
                        )
                        by_connection = {
                            "connectionString": conn_str,
                            "connectionType": "analysisServicesDatabaseLive",
                            "name": "EntityDataSource",
                            "pbiServiceModelId": None,
                            "pbiModelVirtualServerName": None,
                            "pbiModelDatabaseName": None
                        }
                        self.logger.info(f"Building analysisServicesDatabaseLive (XMLA) PBIP connection for cloud target: {database}")
                    else:
                        # Semantic Model: use standard pbiServiceLive
                        # pbiServiceModelId needs to be an integer (Power BI crashes with null)
                        # Generate random 8-digit integer - actual value doesn't matter, just needs to be present
                        # pbiModelVirtualServerName must be "sobe_wowvirtualserver" for pbiServiceLive
                        random_model_id = random.randint(10000000, 99999999)
                        # pbiServiceLive requires dataset GUID for both Initial Catalog and pbiModelDatabaseName
                        # Use dataset_id if available, otherwise fall back to database (friendly name)
                        db_identifier = dataset_id or database
                        # Connection string format matching Power BI's native format
                        conn_str = (
                            f'Data Source=pbiazure://api.powerbi.com;'
                            f'Initial Catalog={db_identifier};'
                            f'Identity Provider="https://login.microsoftonline.com/common, '
                            f'https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc";'
                            f'Integrated Security=ClaimsToken'
                        )
                        by_connection = {
                            "connectionString": conn_str,
                            "connectionType": "pbiServiceLive",
                            "name": "EntityDataSource",
                            "pbiServiceModelId": random_model_id,
                            "pbiModelVirtualServerName": "sobe_wowvirtualserver",
                            "pbiModelDatabaseName": db_identifier  # Dataset GUID required for pbiServiceLive
                        }
                        self.logger.info(f"Building pbiServiceLive PBIP connection for cloud target: {database} (DB: {db_identifier}, ModelId: {random_model_id})")
            else:
                # Local target: analysisServicesDatabaseLive connection
                # Use specified perspective or default to "Model"
                cube_value = perspective_name if perspective_name else "Model"
                by_connection = {
                    "connectionString": f"Data Source={server};Initial Catalog={database};Cube={cube_value}",
                    "connectionType": "analysisServicesDatabaseLive",
                    "name": "EntityDataSource",
                    "pbiServiceModelId": None,
                    "pbiModelVirtualServerName": None,
                    "pbiModelDatabaseName": None
                }
                perspective_info = f" [{perspective_name}]" if perspective_name else ""
                self.logger.info(f"Building analysisServicesDatabaseLive PBIP connection for local target: {server}{perspective_info}")

            new_content = {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/1.0.0/schema.json",
                "version": "4.0",
                "datasetReference": {
                    "byConnection": by_connection
                }
            }

            # Write the new content
            with open(pbir_path, 'w', encoding='utf-8') as f:
                json.dump(new_content, f, indent=2)

            conn_type = "pbiServiceLive" if is_cloud_target else "analysisServicesDatabaseLive"
            self.logger.info(f"Successfully modified PBIP definition ({conn_type}): {pbir_path}")
            return SwapResult(
                success=True,
                message=f"Connection updated to {database}. Close and reopen the file to apply changes."
            )

        except Exception as e:
            self.logger.error(f"Error modifying PBIP: {e}")
            return SwapResult(
                success=False,
                message=f"Error modifying PBIP: {e}"
            )

    def read_current_connection(self, file_path: str) -> Optional[dict]:
        """
        Read the current connection info from a PBIX or PBIP file.

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            Dict with 'server', 'database', 'connection_type' or None if not found
        """
        file_type = self.detect_file_type(file_path)

        if file_type == 'pbix':
            return self._read_connection_pbix(file_path)
        elif file_type == 'pbip':
            return self._read_connection_pbip(file_path)

        return None

    def _read_connection_pbix(self, pbix_path: str) -> Optional[dict]:
        """Read connection info from a PBIX file."""
        try:
            with zipfile.ZipFile(pbix_path, 'r') as zf:
                # Find Connections file (case-insensitive)
                for name in zf.namelist():
                    if name.lower() == 'connections':
                        content = zf.read(name).decode('utf-8')
                        data = json.loads(content)

                        if 'Connections' in data and len(data['Connections']) > 0:
                            conn = data['Connections'][0]
                            conn_str = conn.get('ConnectionString', '')

                            # Parse connection string for basic info
                            result = self._parse_connection_string(
                                conn_str,
                                conn.get('ConnectionType', '')
                            )

                            # For pbiServiceLive connections, extract friendly name from
                            # PbiModelDatabaseName (not from connection string)
                            pbi_db_name = conn.get('PbiModelDatabaseName')
                            if pbi_db_name:
                                result['friendly_name'] = pbi_db_name
                                # Also use this as the database if not already set properly
                                if not result.get('database') or result['database'] == result.get('server'):
                                    result['database'] = pbi_db_name

                            # Store FULL original connection data for restoration
                            result['_original_pbix_connection'] = conn.copy()
                            result['_original_pbix_version'] = data.get('Version', 1)

                            return result

            return None

        except Exception as e:
            self.logger.debug(f"Error reading PBIX connection: {e}")
            return None

    def _read_connection_pbip(self, pbip_path: str) -> Optional[dict]:
        """Read connection info from a PBIP file."""
        try:
            report_folder = self._get_report_folder(pbip_path)
            if not report_folder:
                return None

            pbir_path = os.path.join(report_folder, 'definition.pbir')
            if not os.path.exists(pbir_path):
                return None

            with open(pbir_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            dataset_ref = data.get('datasetReference', {})

            # Check byConnection format (live connection)
            by_conn = dataset_ref.get('byConnection', {})
            if by_conn:
                result = self._parse_connection_string(
                    by_conn.get('connectionString', ''),
                    by_conn.get('connectionType', '')
                )

                # Store FULL original PBIP definition for restoration
                result['_original_pbip_definition'] = data.copy()

                return result

            # Check byPath format (local semantic model)
            by_path = dataset_ref.get('byPath', {})
            if by_path:
                return {
                    'server': 'local',
                    'database': by_path.get('path', ''),
                    'connection_type': 'byPath',
                    '_original_pbip_definition': data.copy()
                }

            return None

        except Exception as e:
            self.logger.debug(f"Error reading PBIP connection: {e}")
            return None

    def _parse_connection_string(self, conn_str: str, conn_type: str) -> dict:
        """Parse a connection string into server, database, and other components."""
        import re

        result = {
            'server': '',
            'database': '',
            'connection_type': conn_type,
            'cloud_connection_type': None,  # 'pbi_semantic_model' or 'aas_xmla'
            'perspective_name': None,
            'workspace_name': None
        }

        def strip_quotes(value: str) -> str:
            """Strip surrounding quotes from a value."""
            if value and len(value) >= 2:
                if (value.startswith('"') and value.endswith('"')) or \
                   (value.startswith("'") and value.endswith("'")):
                    return value[1:-1]
            return value

        # Parse Data Source - handle quoted values
        source_match = re.search(r'Data Source="([^"]+)"', conn_str, re.IGNORECASE)
        if not source_match:
            source_match = re.search(r'Data Source=([^;]+)', conn_str, re.IGNORECASE)
        if source_match:
            result['server'] = strip_quotes(source_match.group(1))

        # Parse Initial Catalog - handle quoted values (common in PBIP files)
        # This contains the friendly model name in cloud connections
        catalog_match = re.search(r'Initial Catalog="([^"]+)"', conn_str, re.IGNORECASE)
        if not catalog_match:
            catalog_match = re.search(r'Initial Catalog=([^;]+)', conn_str, re.IGNORECASE)
        if catalog_match:
            catalog_value = strip_quotes(catalog_match.group(1))
            result['database'] = catalog_value
            result['friendly_name'] = catalog_value  # Store as friendly name too

        # Parse semanticmodelid if present (this is the GUID, not the friendly name)
        modelid_match = re.search(r'semanticmodelid=([^;]+)', conn_str, re.IGNORECASE)
        if modelid_match:
            result['semantic_model_id'] = strip_quotes(modelid_match.group(1))

        # Parse Cube/Perspective name if present (used by AAS/XMLA connections)
        cube_match = re.search(r'Cube="([^"]+)"', conn_str, re.IGNORECASE)
        if not cube_match:
            cube_match = re.search(r'Cube=([^;]+)', conn_str, re.IGNORECASE)
        if cube_match:
            cube_value = strip_quotes(cube_match.group(1))
            # "Model" is the default, not a perspective
            if cube_value.lower() != 'model':
                result['perspective_name'] = cube_value

        # Extract workspace name from powerbi:// or pbiazure:// URL
        # Format: powerbi://api.powerbi.com/v1.0/myorg/WorkspaceName
        # Format: pbiazure://api.powerbi.com/v1.0/myorg/WorkspaceName
        server = result.get('server', '')
        if server:
            workspace_match = re.search(r'(?:powerbi|pbiazure)://api\.powerbi\.com/v1\.0/myorg/([^/;]+)', server, re.IGNORECASE)
            if workspace_match:
                result['workspace_name'] = workspace_match.group(1)

        # Determine cloud connection type
        conn_type_lower = conn_type.lower() if conn_type else ''
        server_lower = server.lower()

        if conn_type_lower == 'pbiserviceLive'.lower():
            # Power BI Semantic Model connector
            result['cloud_connection_type'] = 'pbi_semantic_model'
        elif conn_type_lower == 'analysisservicesdatabaselive':
            if 'localhost' in server_lower or server_lower.startswith('127.'):
                # Local Analysis Services - not a cloud connection
                result['cloud_connection_type'] = None
            elif 'powerbi://' in server_lower or 'asazure://' in server_lower:
                # AAS/XMLA endpoint (supports perspectives, requires Premium)
                result['cloud_connection_type'] = 'aas_xmla'
            elif 'pbiazure://' in server_lower:
                # Generic cloud endpoint - could be either type
                result['cloud_connection_type'] = 'pbi_semantic_model'

        return result


# Module-level instance for convenience
_modifier: Optional[PbixModifier] = None


def get_modifier() -> PbixModifier:
    """Get the global PbixModifier instance."""
    global _modifier
    if _modifier is None:
        _modifier = PbixModifier()
    return _modifier
