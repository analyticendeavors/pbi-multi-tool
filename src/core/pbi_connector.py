"""
Power BI Connector - Native Power BI Model Connection
Built by Reid Havens of Analytic Endeavors

Direct connection to Power BI Desktop via XMLA endpoint using Tabular Object Model (TOM).
Similar to how Tabular Editor and DAX Studio connect.

This is a shared connector used by multiple tools:
- Field Parameters Tool
- Connection Hotswap Tool
"""

import sys
import logging
import socket
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from pathlib import Path


# =============================================================================
# Data Classes - defined here to avoid circular imports
# =============================================================================

@dataclass(slots=True)
class ModelConnection:
    """Represents a connection to a Power BI model."""
    server: str  # localhost:port
    database: str  # Model name/GUID
    model_name: str  # Friendly name
    is_connected: bool = False


@dataclass(slots=True)
class AvailableModel:
    """Represents a discovered Power BI model."""
    server: str  # localhost:port
    database_name: str  # Model/database name
    display_name: str  # For UI display
    # Thin report attributes (live-connected reports with no local database)
    is_thin_report: bool = False  # True if this is a live-connected thin report
    thin_report_cloud_server: Optional[str] = None  # Cloud XMLA endpoint (if known)
    thin_report_cloud_database: Optional[str] = None  # Cloud dataset name (if known)
    # File path and process info (for thin report swapping)
    file_path: Optional[str] = None  # Full path to .pbix or .pbip file
    process_id: Optional[int] = None  # Power BI Desktop process ID


@dataclass(slots=True)
class TableInfo:
    """Information about a table in the model."""
    name: str
    is_hidden: bool
    table_type: str  # "Table", "Calculated", etc.
    is_measures_only: bool = False  # True if all columns hidden/deleted but has visible measures


@dataclass(slots=True)
class FieldInfo:
    """Information about a measure or column in a table."""
    name: str
    table_name: str
    field_type: str  # "Measure" or "Column"
    data_type: Optional[str] = None
    expression: Optional[str] = None
    is_hidden: bool = False
    display_folder: str = ""  # Display folder path (e.g., "Sales\\Revenue")


@dataclass(slots=True)
class TableFieldsInfo:
    """Complete info about a table's fields and metadata."""
    table_name: str
    fields: List[FieldInfo]  # All visible measures and columns with folder info
    is_measures_only: bool  # True if measures-only table (shows at top)
    sort_priority: int  # 0 = measures tables, 1 = regular tables


__all__ = [
    'ModelConnection',
    'AvailableModel',
    'TableInfo',
    'FieldInfo',
    'TableFieldsInfo',
    'PowerBIConnector',
    'get_connector',
]


class PowerBIConnector:
    """
    Native connector to Power BI Desktop models via XMLA endpoint.

    Uses Microsoft.AnalysisServices.Tabular (TOM) via pythonnet.
    Supports both local (Power BI Desktop) and cloud (Power BI Service) connections.
    """

    # Azure AD App Registration for Power BI
    # This is a well-known public client ID used by Power BI tools
    AZURE_CLIENT_ID = "a672d62c-fc7b-4e81-a576-e60dc46e951d"  # Power BI public app
    AZURE_AUTHORITY = "https://login.microsoftonline.com/common"
    AZURE_SCOPES = ["https://analysis.windows.net/powerbi/api/.default"]

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._server = None
        self._database = None
        self._model = None
        self.current_connection: Optional[ModelConnection] = None
        self._access_token = None  # For cloud connections
        self._is_cloud_connection = False

        # Cache for model metadata - invalidated on reconnect
        self._tables_cache: Optional[List[TableInfo]] = None
        self._fields_cache: Optional[Dict[str, TableFieldsInfo]] = None

        # Try to initialize .NET/TOM
        self._initialize_tom()
    
    def _initialize_tom(self) -> bool:
        """Initialize Tabular Object Model via pythonnet"""
        try:
            import clr
            
            # Power BI Desktop uses Microsoft.AnalysisServices.Server.Tabular.dll
            # Tabular Editor 3 uses Microsoft.AnalysisServices.Tabular.dll
            # We prefer Power BI Desktop's version for better compatibility
            
            # Try Power BI Desktop first
            pbi_path = Path(r"C:\Program Files\Microsoft Power BI Desktop\bin")
            if pbi_path.exists():
                sys.path.append(str(pbi_path))
                try:
                    clr.AddReference("Microsoft.AnalysisServices.Server.Tabular")
                    from Microsoft.AnalysisServices.Tabular import Server, Database, Model
                    
                    self.Server = Server
                    self.Database = Database
                    self.Model = Model
                    
                    self.logger.info(f"✅ Loaded TOM from Power BI Desktop: {pbi_path}")
                    self.logger.info("✅ Tabular Object Model initialized successfully")
                    return True
                except Exception as e:
                    self.logger.debug(f"Power BI Desktop DLL failed: {e}")
            
            # Fall back to other locations
            possible_paths = [
                # Tabular Editor installation
                Path(r"C:\Program Files\Tabular Editor 3"),
                # SSMS installation
                Path(r"C:\Program Files (x86)\Microsoft SQL Server Management Studio 19\Common7\IDE\Extensions\Microsoft\SQLDB\DAC"),
                Path(r"C:\Program Files (x86)\Microsoft SQL Server\160\SDK\Assemblies"),
            ]
            
            for path in possible_paths:
                dll_path = path / "Microsoft.AnalysisServices.Tabular.dll"
                if dll_path.exists():
                    sys.path.append(str(path))
                    try:
                        clr.AddReference("Microsoft.AnalysisServices.Tabular")
                        from Microsoft.AnalysisServices.Tabular import Server, Database, Model
                        
                        self.Server = Server
                        self.Database = Database
                        self.Model = Model
                        
                        self.logger.info(f"✅ Loaded TOM from: {dll_path}")
                        self.logger.info("✅ Tabular Object Model initialized successfully")
                        return True
                    except Exception as e:
                        self.logger.debug(f"Failed to load from {dll_path}: {e}")
                        continue
            
            # Try GAC as last resort
            self.logger.warning("TOM DLL not found in common paths - trying GAC")
            try:
                clr.AddReference("Microsoft.AnalysisServices.Tabular")
                from Microsoft.AnalysisServices.Tabular import Server, Database, Model
                self.Server = Server
                self.Database = Database
                self.Model = Model
                self.logger.info("✅ Loaded TOM from GAC")
                self.logger.info("✅ Tabular Object Model initialized successfully")
                return True
            except:
                pass
            
            self.logger.error("❌ Could not load Microsoft.AnalysisServices.Tabular DLL")
            self.logger.error("Install Power BI Desktop, SSMS, or Tabular Editor 3")
            return False
                
        except ImportError as e:
            # pythonnet is optional - log at debug level to avoid alarming users
            self.logger.debug(f"pythonnet not available: {e} (optional dependency)")
            return False
        except Exception as e:
            self.logger.error(f"Failed to initialize TOM: {e}")
            return False

    def _invalidate_cache(self):
        """Clear cached model metadata - call on connect/disconnect"""
        self._tables_cache = None
        self._fields_cache = None
        self.logger.debug("Model metadata cache invalidated")

    def _check_port(self, port: int) -> Optional[int]:
        """Check if a port is open (helper for parallel scanning)"""
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.settimeout(0.05)
                result = sock.connect_ex(('localhost', port))
                if result == 0:
                    return port
        except:
            pass
        return None
    
    def _get_listening_ports(self) -> List[int]:
        """
        Get all listening TCP ports on localhost using netstat.
        This is much faster than scanning ranges.
        
        Returns:
            List of port numbers that are listening
        """
        import subprocess
        
        listening_ports = []
        try:
            # Use netstat to find listening ports
            # -a = all connections, -n = numerical, -p TCP = protocol
            result = subprocess.run(
                ['netstat', '-an', '-p', 'TCP'],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    # Look for LISTENING state on localhost
                    if 'LISTENING' in line and '127.0.0.1:' in line:
                        try:
                            # Extract port from "127.0.0.1:PORT" or "[::1]:PORT"
                            parts = line.split()
                            for part in parts:
                                if '127.0.0.1:' in part or '[::1]:' in part:
                                    port_str = part.split(':')[-1]
                                    port = int(port_str)
                                    # Only care about ports in Power BI range
                                    if 49152 <= port <= 65535:
                                        listening_ports.append(port)
                                    break
                        except (ValueError, IndexError):
                            continue
            
            # Remove duplicates and sort
            listening_ports = sorted(set(listening_ports))
            self.logger.debug(f"Found {len(listening_ports)} listening ports in range 49152-65535")
            
        except subprocess.TimeoutExpired:
            self.logger.warning("netstat command timed out")
        except FileNotFoundError:
            self.logger.warning("netstat command not found")
        except Exception as e:
            self.logger.debug(f"Error getting listening ports: {e}")
        
        return listening_ports
    
    def _get_pbi_process_from_port(self, port: int) -> Optional[Tuple[int, 'psutil.Process']]:
        """
        Find the Power BI Desktop process from a port number.

        The port is owned by msmdsrv.exe (Analysis Services), and we need to
        find its parent PBIDesktop.exe process.

        Args:
            port: The port number the model is running on

        Returns:
            Tuple of (process_id, psutil.Process) for PBIDesktop.exe, or None if not found
        """
        try:
            import subprocess
            import psutil

            # Step 1: Find which process ID is listening on the port using netstat
            netstat_result = subprocess.run(
                ['netstat', '-ano', '-p', 'TCP'],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            msmdsrv_pid = None
            if netstat_result.returncode == 0:
                for line in netstat_result.stdout.split('\n'):
                    # Look for LISTENING state on our specific port
                    if 'LISTENING' in line and f':{port}' in line:
                        try:
                            # Last column in netstat -ano output is the PID
                            parts = line.split()
                            pid = parts[-1]
                            msmdsrv_pid = int(pid)
                            self.logger.debug(f"Port {port} is used by process ID: {msmdsrv_pid}")
                            break
                        except (ValueError, IndexError):
                            continue

            if not msmdsrv_pid:
                self.logger.debug(f"Could not find process ID for port {port}")
                return None

            # Step 2: Verify it's msmdsrv.exe and find parent PBIDesktop.exe
            process = psutil.Process(msmdsrv_pid)
            process_name = process.name()
            self.logger.debug(f"Process {msmdsrv_pid} is: {process_name}")

            # msmdsrv.exe is the Analysis Services engine spawned by PBIDesktop.exe
            if process_name.lower() == "msmdsrv.exe":
                # Find the parent PBIDesktop.exe process
                parent = process.parent()
                if parent:
                    parent_name = parent.name()
                    self.logger.debug(f"Parent process {parent.pid} is: {parent_name}")
                    if parent_name.lower() == "pbidesktop.exe":
                        return (parent.pid, parent)
                    else:
                        self.logger.debug(f"Parent is not PBIDesktop.exe, it's {parent_name}")
                        return None
                else:
                    self.logger.debug(f"Could not find parent process for msmdsrv.exe")
                    return None
            elif process_name.lower() == "pbidesktop.exe":
                return (msmdsrv_pid, process)
            else:
                self.logger.debug(f"Process {msmdsrv_pid} is neither PBIDesktop.exe nor msmdsrv.exe")
                return None

        except ImportError:
            self.logger.debug("psutil not available")
            return None
        except Exception as e:
            self.logger.debug(f"Error getting PBI process from port {port}: {e}")
            return None

    def _get_pbix_filepath_from_process(self, port: int) -> Tuple[Optional[str], Optional[int]]:
        """
        Get the full file path of the .pbix/.pbip file from the Power BI Desktop process.

        Uses WMI to query the CommandLine of the PBIDesktop.exe process, which contains
        the file path as a quoted argument.

        Args:
            port: The port number the model is running on

        Returns:
            Tuple of (file_path, process_id) or (None, None) if not found
        """
        try:
            import subprocess
            import re

            # Get the PBI Desktop process
            result = self._get_pbi_process_from_port(port)
            if not result:
                return None, None

            process_id, pbi_process = result
            self.logger.info(f"Getting file path for PBIDesktop.exe process {process_id}")

            # Use PowerShell/WMI to get the CommandLine
            ps_script = f'''
            $proc = Get-WmiObject Win32_Process -Filter "ProcessId = {process_id}"
            if ($proc) {{ $proc.CommandLine }}
            '''

            try:
                ps_result = subprocess.run(
                    ['powershell', '-Command', ps_script],
                    capture_output=True,
                    text=True,
                    timeout=10,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            except subprocess.TimeoutExpired:
                self.logger.warning(f"WMI query timed out for process {process_id} - Windows may need a restart")
                self._wmi_failures = getattr(self, '_wmi_failures', 0) + 1
                return None, process_id

            if ps_result.returncode != 0 or not ps_result.stdout.strip():
                self.logger.warning(f"WMI query failed for process {process_id} (returncode={ps_result.returncode})")
                if ps_result.stderr:
                    self.logger.warning(f"  WMI error: {ps_result.stderr.strip()[:200]}")
                self._wmi_failures = getattr(self, '_wmi_failures', 0) + 1
                return None, process_id

            cmd_line = ps_result.stdout.strip()
            self.logger.debug(f"CommandLine: {cmd_line[:200]}...")

            # Parse the file path from CommandLine
            # Format: "C:\...\PBIDesktop.exe" "C:\Users\...\report.pbix"
            # Or: "C:\...\PBIDesktop.exe" "C:\Users\...\report.pbip"
            # The file path is the second quoted argument

            # Find all quoted strings
            quoted_pattern = r'"([^"]+)"'
            matches = re.findall(quoted_pattern, cmd_line)

            for match in matches:
                # Skip the executable itself
                if match.lower().endswith('.exe'):
                    continue
                # Look for .pbix or .pbip files
                if match.lower().endswith('.pbix') or match.lower().endswith('.pbip'):
                    self.logger.info(f"Found file path: {match}")
                    return match, process_id

            # Fallback: look for unquoted paths (less common)
            # Pattern for paths like C:\...\file.pbix or C:\...\file.pbip
            path_pattern = r'([A-Za-z]:\\[^\s]+\.pbi[xp])'
            path_matches = re.findall(path_pattern, cmd_line, re.IGNORECASE)

            for match in path_matches:
                self.logger.info(f"Found file path (unquoted): {match}")
                return match, process_id

            self.logger.debug(f"No .pbix/.pbip file path found in CommandLine")
            return None, process_id

        except Exception as e:
            self.logger.debug(f"Error getting file path from process for port {port}: {e}")
            return None, None

    def _get_pbix_filename_from_process(self, port: int) -> Optional[str]:
        """
        Try to get the .pbix filename by finding the Power BI Desktop process
        that's using the specified port and extracting the filename from the window title.

        Args:
            port: The port number the model is running on

        Returns:
            The .pbix filename without extension, or None if not found
        """
        try:
            import subprocess

            # Get the PBI Desktop process
            result = self._get_pbi_process_from_port(port)
            if not result:
                return None

            process_id, pbi_process = result
            self.logger.info(f"Process {process_id} is: {pbi_process.name()}")
            
            # Step 3: Get the window title for this specific process ID
            ps_script = f"""
            Add-Type @"
                using System;
                using System.Runtime.InteropServices;
                using System.Text;
                public class Win32 {{
                    [DllImport("user32.dll", CharSet = CharSet.Auto)]
                    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
                    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
                    public static extern int GetWindowTextLength(IntPtr hWnd);
                }}
"@
            $process = Get-Process -Id {process_id} -ErrorAction SilentlyContinue
            if ($process) {{
                $length = [Win32]::GetWindowTextLength($process.MainWindowHandle)
                if ($length -gt 0) {{
                    $sb = New-Object System.Text.StringBuilder ($length + 1)
                    [Win32]::GetWindowText($process.MainWindowHandle, $sb, $sb.Capacity) | Out-Null
                    $sb.ToString()
                }}
            }}
            """
            
            ps_result = subprocess.run(
                ['powershell', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            self.logger.debug(f"PowerShell return code: {ps_result.returncode}")
            self.logger.debug(f"PowerShell stdout: '{ps_result.stdout.strip()}'")
            self.logger.debug(f"PowerShell stderr: '{ps_result.stderr.strip()}'")
            
            if ps_result.returncode == 0 and ps_result.stdout.strip():
                window_title = ps_result.stdout.strip()
                self.logger.info(f"Got window title: '{window_title}'")

                # Known suffixes that Power BI Desktop may use (varies by version)
                known_suffixes = [' - Power BI Desktop', ' - Power Query Editor']

                # Known titles that are NOT file names (just app names)
                excluded_titles = ['Power BI Desktop', 'Power Query Editor', 'Microsoft Power BI Desktop']

                # Try to strip known suffixes first
                filename = window_title
                for suffix in known_suffixes:
                    if suffix in window_title:
                        filename = window_title.replace(suffix, '').strip()
                        self.logger.info(f"Stripped suffix '{suffix}' -> '{filename}'")
                        break

                # Return filename if it's valid (not empty, not just an app name)
                if filename and filename not in excluded_titles:
                    self.logger.info(f"Port {port} -> Process {process_id} -> Filename: {filename}")
                    return filename
                else:
                    self.logger.debug(f"Window title '{window_title}' is not a valid filename")
            else:
                self.logger.debug(f"PowerShell failed or returned empty output for process {process_id}")

            return None
            
        except Exception as e:
            self.logger.debug(f"Error getting filename from process for port {port}: {e}")
            return None
    
    def _get_models_from_port(self, port: int) -> List[AvailableModel]:
        """Try to get models from a specific port (helper for parallel scanning)"""
        models = []
        server_address = f"localhost:{port}"
        try:
            temp_server = self.Server()
            # Add timeout to fail fast on non-AS ports (1 second for fast discovery)
            connection_string = f"Provider=MSOLAP;Data Source={server_address};Connect Timeout=1;"
            self.logger.info(f"Attempting TOM connection to {server_address}")
            temp_server.Connect(connection_string)

            # Get databases
            db_count = temp_server.Databases.Count
            self.logger.info(f"Port {port}: Connected, found {db_count} database(s)")

            if db_count == 0:
                # This might be a thin report (live-connected report)
                # Try to detect and handle it
                thin_report_model = self._detect_thin_report(temp_server, port, server_address)
                if thin_report_model:
                    models.append(thin_report_model)
                    self.logger.info(f"Detected thin report on port {port}")
                temp_server.Disconnect()
                return models

            for db in temp_server.Databases:
                database_guid = db.Name

                # Try multiple approaches to get a friendly name (like DAX Studio does)
                friendly_name = None

                # 1. Try to get .pbix filename from Power BI Desktop window title
                try:
                    pbix_name = self._get_pbix_filename_from_process(port)
                    if pbix_name:
                        friendly_name = pbix_name
                        self.logger.info(f"Got friendly name from process: {friendly_name}")
                except Exception as e:
                    self.logger.debug(f"Error getting filename from process: {e}")

                # 2. Try Database.FriendlyName property (if not already found)
                if not friendly_name:
                    try:
                        if hasattr(db, 'FriendlyName') and db.FriendlyName:
                            friendly_name = db.FriendlyName
                            self.logger.info(f"Got friendly name from db.FriendlyName: {friendly_name}")
                    except Exception as e:
                        self.logger.debug(f"Error getting db.FriendlyName: {e}")

                # 3. Try Model.Name (if still not found)
                if not friendly_name:
                    try:
                        if hasattr(db, 'Model') and db.Model and hasattr(db.Model, 'Name'):
                            friendly_name = db.Model.Name
                            self.logger.info(f"Got friendly name from Model.Name: {friendly_name}")
                    except Exception as e:
                        self.logger.debug(f"Error getting Model.Name: {e}")

                # Ultimate fallback to GUID
                if not friendly_name or friendly_name == database_guid:
                    friendly_name = database_guid

                # Get file path and process ID for preset persistence
                file_path, process_id = self._get_pbix_filepath_from_process(port)

                models.append(AvailableModel(
                    server=server_address,
                    database_name=database_guid,  # Still use GUID for connection
                    display_name=f"{friendly_name} ({port})",  # Show friendly name to user
                    file_path=file_path,
                    process_id=process_id
                ))
                self.logger.info(f"Found model: {friendly_name} on port {port}")

            temp_server.Disconnect()
        except Exception as e:
            # Log more detail about why connection failed
            error_type = type(e).__name__
            self.logger.info(f"Port {port}: Not AS or failed - {error_type}: {str(e)[:100]}")

        return models

    def _detect_thin_report(self, server, port: int, server_address: str) -> Optional[AvailableModel]:
        """
        Detect and create an AvailableModel for a thin report (live-connected report).

        Thin reports have 0 databases in the local AS instance because they're
        pass-through connections to a cloud semantic model.

        Args:
            server: Connected TOM Server object
            port: Port number
            server_address: Server address string (localhost:port)

        Returns:
            AvailableModel if thin report detected, None otherwise
        """
        try:
            # First verify this is a Power BI Desktop process (msmdsrv.exe)
            # and get the full file path for thin report file modification
            file_path, process_id = self._get_pbix_filepath_from_process(port)

            # Also get the friendly name from window title
            pbix_name = self._get_pbix_filename_from_process(port)

            if not pbix_name:
                self.logger.debug(f"Port {port}: 0 databases but not a PBI Desktop process")
                return None

            self.logger.info(f"Port {port}: Detected thin report: {pbix_name}")
            if file_path:
                self.logger.info(f"  File path: {file_path}")

            # Try to read cloud connection info from the file itself
            # This works for PBIP files (not locked) and may work for PBIX (if read-only access allowed)
            cloud_connection_info = None
            if file_path:
                cloud_connection_info = self._read_connection_from_file(file_path)

            # Fallback: try to extract from server properties
            if not cloud_connection_info:
                cloud_connection_info = self._extract_thin_report_connection(server)

            # Create AvailableModel with thin report marker
            # Use special database name format to indicate it's a thin report
            display_suffix = "(Thin Report)"
            cloud_server = ''
            cloud_database = ''
            cloud_friendly_name = ''

            if cloud_connection_info:
                cloud_server = cloud_connection_info.get('server', '')
                # Use friendly_name if available (model name), otherwise fall back to database
                cloud_friendly_name = cloud_connection_info.get('friendly_name', '')
                cloud_database = cloud_friendly_name or cloud_connection_info.get('database', '')

                # Check if this is actually a LOCAL connection (already swapped)
                # Local connections use localhost:PORT, not cloud URLs
                server_lower = cloud_server.lower() if cloud_server else ''
                is_actually_local = (
                    server_lower.startswith('localhost') or
                    server_lower.startswith('127.0.0.1') or
                    (server_lower and not any(x in server_lower for x in ['powerbi://', 'pbiazure://', 'asazure://']))
                )

                if is_actually_local:
                    self.logger.info(f"  Currently connected to LOCAL: {cloud_server}")
                    self.logger.info(f"  Database: {cloud_database}")
                    # Mark as local, not cloud - clear the cloud fields
                    cloud_server = ''
                    cloud_database = ''
                    display_suffix = "(Thin Report - Local)"
                else:
                    self.logger.info(f"  Cloud endpoint: {cloud_server}")
                    self.logger.info(f"  Cloud dataset: {cloud_database}")

            return AvailableModel(
                server=server_address,
                database_name=f"__thin_report__:{cloud_server}:{cloud_database}",
                display_name=f"{pbix_name} {display_suffix}",
                is_thin_report=True,
                thin_report_cloud_server=cloud_server if cloud_server else None,
                thin_report_cloud_database=cloud_database if cloud_database else None,
                file_path=file_path,
                process_id=process_id,
            )

        except Exception as e:
            self.logger.debug(f"Error detecting thin report on port {port}: {e}")
            return None

    def _read_connection_from_file(self, file_path: str) -> Optional[dict]:
        """
        Read cloud connection info directly from a PBIX or PBIP file.

        This method reads the connection string from the file itself rather than
        trying to extract it from the TOM server properties. For thin reports,
        this provides the actual cloud connection details.

        Note: PBIP files can be read while Power BI Desktop is open (not locked).
        PBIX files may be locked, in which case this will fail gracefully.

        Args:
            file_path: Full path to the .pbix or .pbip file

        Returns:
            Dict with 'server', 'database', 'connection_type' keys or None if not readable
        """
        try:
            # Import the PBIX modifier module
            from tools.connection_hotswap.logic.pbix_modifier import get_modifier

            modifier = get_modifier()
            connection_info = modifier.read_current_connection(file_path)

            if connection_info:
                self.logger.info(f"Read connection from file: server={connection_info.get('server', '')[:50]}...")
                return connection_info

            return None

        except ImportError as e:
            self.logger.debug(f"Could not import pbix_modifier: {e}")
            return None
        except Exception as e:
            # This is expected for locked PBIX files
            self.logger.debug(f"Could not read connection from file {file_path}: {e}")
            return None

    def _extract_thin_report_connection(self, server) -> Optional[dict]:
        """
        Try to extract cloud connection info from a thin report's AS server.

        Args:
            server: Connected TOM Server object

        Returns:
            Dict with 'server', 'database', 'dataset' keys if found, None otherwise
        """
        try:
            import re

            # Check for server properties that might contain connection info
            if hasattr(server, 'ServerProperties'):
                self.logger.debug(f"Checking {server.ServerProperties.Count} server properties for cloud connection")
                for prop in server.ServerProperties:
                    prop_name = str(prop.Name) if hasattr(prop, 'Name') else ''
                    prop_value = str(prop.Value) if hasattr(prop, 'Value') else ''

                    # Log interesting properties
                    if prop_value and len(prop_value) < 200:
                        self.logger.debug(f"  ServerProperty: {prop_name} = {prop_value[:100]}")

                    # Look for properties that might indicate the cloud connection
                    if 'powerbi://' in prop_value.lower():
                        self.logger.info(f"Found cloud connection in ServerProperty: {prop_name}")
                        match = re.search(r'(powerbi://[^\s;]+)', prop_value, re.IGNORECASE)
                        if match:
                            server_url = match.group(1)
                            # Try to extract workspace/dataset
                            workspace_match = re.search(r'myorg/([^/]+)', server_url)
                            workspace = workspace_match.group(1) if workspace_match else ''
                            return {
                                'server': server_url,
                                'database': workspace,  # Initial guess
                                'dataset': workspace,
                            }

            # Try to access connection string through other means
            if hasattr(server, 'ConnectionString'):
                conn_str = str(server.ConnectionString) if server.ConnectionString else ''
                self.logger.debug(f"Server.ConnectionString: {conn_str[:100] if conn_str else '(empty)'}")
                if 'powerbi://' in conn_str.lower():
                    self.logger.info("Found cloud connection in server ConnectionString")
                    match = re.search(r'Data Source=([^;]+)', conn_str, re.IGNORECASE)
                    if match:
                        server_url = match.group(1)
                        cat_match = re.search(r'Initial Catalog=([^;]+)', conn_str, re.IGNORECASE)
                        database = cat_match.group(1) if cat_match else ''
                        return {
                            'server': server_url,
                            'database': database,
                            'dataset': database,
                        }

            # The cloud connection info might not be accessible via TOM for thin reports
            # This is expected - thin reports don't expose their remote connection through local AS
            self.logger.info("Cloud connection info not available via TOM (this is normal for thin reports)")
            return None

        except Exception as e:
            self.logger.debug(f"Error extracting thin report connection: {e}")
            return None
    
    def discover_local_models_fast(self) -> List[AvailableModel]:
        """
        Fast discovery using msmdsrv.port.txt files (same method as DAX Studio).
        Power BI Desktop writes port info to temp files we can read.

        Returns:
            List of discovered models
        """
        import os
        import tempfile
        from pathlib import Path

        # Reset WMI failure counter at start of scan
        self.reset_wmi_failure_count()

        models = []
        
        # Power BI Desktop writes port files to TEMP\Power BI Desktop
        temp_dir = Path(tempfile.gettempdir()) / "Power BI Desktop"
        
        if not temp_dir.exists():
            self.logger.info("No Power BI Desktop temp folder found")
            return []
        
        self.logger.info(f"Checking for port files in: {temp_dir}")
        
        # List what's actually in the directory
        if temp_dir.exists():
            subdirs = list(temp_dir.iterdir())
            self.logger.info(f"Found {len(subdirs)} subdirectories in Power BI Desktop temp")
            for subdir in subdirs[:5]:  # Log first 5
                self.logger.info(f"  Subdir: {subdir.name}")
                port_file = subdir / "msmdsrv.port.txt"
                self.logger.info(f"    Port file exists: {port_file.exists()}")
        
        # Find all msmdsrv.port.txt files
        port_files_found = list(temp_dir.glob("*/msmdsrv.port.txt"))
        self.logger.info(f"Glob found {len(port_files_found)} port files")
        
        # Also try root level and .port files
        for pattern in ["msmdsrv.port.txt", "*.port.txt", "*/msmdsrv.port", "*/*.port"]:
            extra_files = list(temp_dir.glob(pattern))
            if extra_files:
                self.logger.info(f"Pattern '{pattern}' found {len(extra_files)} files")
                port_files_found.extend(extra_files)
        
        for port_file in port_files_found:
            try:
                # Read port number
                port = int(port_file.read_text().strip())
                self.logger.info(f"Found port file: {port_file.name} -> port {port}")
                
                # Try to connect
                server_address = f"localhost:{port}"
                try:
                    temp_server = self.Server()
                    conn_str = f"Provider=MSOLAP;Data Source={server_address};"
                    temp_server.Connect(conn_str)
                    
                    for db in temp_server.Databases:
                        database_guid = db.Name
                        
                        # Try multiple approaches to get a friendly name (like DAX Studio does)
                        friendly_name = None
                        
                        # 1. Try to get .pbix filename from Power BI Desktop window title
                        try:
                            pbix_name = self._get_pbix_filename_from_process(port)
                            if pbix_name:
                                friendly_name = pbix_name
                                self.logger.info(f"Got friendly name from process: {friendly_name}")
                        except Exception as e:
                            self.logger.debug(f"Error getting filename from process: {e}")
                        
                        # 2. Try Database.FriendlyName property (if not already found)
                        if not friendly_name:
                            try:
                                if hasattr(db, 'FriendlyName') and db.FriendlyName:
                                    friendly_name = db.FriendlyName
                                    self.logger.info(f"Got friendly name from db.FriendlyName: {friendly_name}")
                            except Exception as e:
                                self.logger.debug(f"Error getting db.FriendlyName: {e}")
                        
                        # 3. Try Model.Name (if still not found)
                        if not friendly_name:
                            try:
                                if hasattr(db, 'Model') and db.Model and hasattr(db.Model, 'Name'):
                                    friendly_name = db.Model.Name
                                    self.logger.info(f"Got friendly name from Model.Name: {friendly_name}")
                            except Exception as e:
                                self.logger.debug(f"Error getting Model.Name: {e}")
                        
                        # Ultimate fallback to GUID
                        if not friendly_name or friendly_name == database_guid:
                            friendly_name = database_guid
                        
                        models.append(AvailableModel(
                            server=server_address,
                            database_name=database_guid,  # Still use GUID for connection
                            display_name=f"{friendly_name} (:{port})"  # Show friendly name to user
                        ))
                        self.logger.info(f"✅ Found model: {friendly_name} on port {port}")
                    
                    temp_server.Disconnect()
                except Exception as e:
                    self.logger.debug(f"Port {port} connection failed: {e}")
                
            except Exception as e:
                self.logger.debug(f"Error reading port file {port_file}: {e}")
        
        self.logger.info(f"Fast discovery found {len(models)} model(s)")
        return models
    
    def get_wmi_failure_count(self) -> int:
        """Get the number of WMI query failures since last reset."""
        return getattr(self, '_wmi_failures', 0)

    def reset_wmi_failure_count(self):
        """Reset the WMI failure counter."""
        self._wmi_failures = 0

    def discover_local_models(self, quick_scan: bool = True, progress_callback=None) -> List[AvailableModel]:
        """
        Discover available Power BI models running on localhost.
        Uses smart port detection by querying OS for listening ports.

        Args:
            quick_scan: If True, scans common range (50000-55000). If False, scans full ephemeral range.
            progress_callback: Optional callback function(message: str) for progress updates

        Returns:
            List of discovered models
        """
        if not hasattr(self, 'Server'):
            self.logger.warning("TOM not initialized - cannot discover models")
            return []

        # Reset WMI failure counter at start of scan
        self.reset_wmi_failure_count()

        models = []
        
        # Try smart detection first: query OS for listening ports
        if progress_callback:
            progress_callback("Checking for listening ports...")
        
        listening_ports = self._get_listening_ports()
        if listening_ports:
            self.logger.info(f"Found {len(listening_ports)} listening ports via netstat")
            if progress_callback:
                progress_callback(f"Found {len(listening_ports)} listening ports, checking for Analysis Services...")
            
            # Check only the ports that are actually listening
            for i, port in enumerate(listening_ports, 1):
                if progress_callback:
                    progress_callback(f"Checking port {port} ({i}/{len(listening_ports)})...")
                
                port_models = self._get_models_from_port(port)
                if port_models:
                    models.extend(port_models)
                    if progress_callback:
                        progress_callback(f"✅ Found {len(port_models)} model(s) on port {port}")
            
            if models:
                self.logger.info(f"Smart detection found {len(models)} model(s)")
                if progress_callback:
                    progress_callback(f"Discovery complete: {len(models)} model(s) found")
                return models
        
        # Fallback to range scanning if smart detection finds nothing
        if progress_callback:
            progress_callback("No listening ports found, trying range scan...")
        
        # Smart port ranges - Power BI Desktop typically uses 50000-55000
        if quick_scan:
            start_port = 50000
            end_port = 55000  # Expanded from 52000 for better coverage
            self.logger.info(f"Quick scan: ports {start_port}-{end_port}...")
            if progress_callback:
                progress_callback(f"Quick scan: checking ports {start_port}-{end_port}...")
        else:
            # Full scan: IANA dynamic/ephemeral range
            start_port = 49152
            end_port = 65535
            self.logger.info(f"Full scan: ports {start_port}-{end_port} (this may take several minutes)...")
            if progress_callback:
                progress_callback(f"Full scan: ports {start_port}-{end_port} (this will take a while)...")
        
        # Phase 1: Parallel port checking (fast)
        if progress_callback:
            progress_callback("Phase 1: Finding open ports...")
        
        open_ports = []
        with ThreadPoolExecutor(max_workers=200) as executor:
            future_to_port = {executor.submit(self._check_port, port): port 
                             for port in range(start_port, end_port + 1)}
            
            for future in as_completed(future_to_port):
                result = future.result()
                if result is not None:
                    open_ports.append(result)
        
        self.logger.info(f"Found {len(open_ports)} open port(s)")
        
        # Phase 2: Check open ports for AS instances
        if open_ports:
            if progress_callback:
                progress_callback(f"Phase 2: Checking {len(open_ports)} open ports for Analysis Services...")
            
            for i, port in enumerate(sorted(open_ports), 1):
                if progress_callback:
                    progress_callback(f"Checking port {port} ({i}/{len(open_ports)})...")
                
                port_models = self._get_models_from_port(port)
                if port_models:
                    models.extend(port_models)
                    if progress_callback:
                        progress_callback(f"✅ Found {len(port_models)} model(s) on port {port}")
        else:
            self.logger.warning(f"No open ports found in range {start_port}-{end_port}")
            if progress_callback:
                progress_callback(f"No open ports found in range {start_port}-{end_port}")
        
        self.logger.info(f"Scan complete: Discovered {len(models)} local model(s)")
        if progress_callback:
            progress_callback(f"Discovery complete: {len(models)} model(s) found")
        
        return models
    
    def connect(self, server: str, database: str) -> bool:
        """
        Connect to a Power BI model.
        
        Args:
            server: Connection string (e.g., "localhost:52784")
            database: Database name or GUID
            
        Returns:
            True if connection successful
        """
        try:
            if not hasattr(self, 'Server'):
                self.logger.error("TOM not initialized - cannot connect")
                return False
            
            self.logger.info(f"Connecting to server: {server}, database: {database}")
            
            # Create server connection
            self._server = self.Server()
            connection_string = f"Provider=MSOLAP;Data Source={server};"
            self._server.Connect(connection_string)
            
            # Get database
            self._database = self._server.Databases.FindByName(database)
            if not self._database:
                # Try by ID if name lookup fails
                self._database = self._server.Databases.GetByName(database)
            
            if not self._database:
                self.logger.error(f"Database not found: {database}")
                return False
            
            self.logger.info(f"Database found: {self._database.Name}")
            
            # Get model
            self._model = self._database.Model
            self.logger.info(f"Model object: {type(self._model)}")
            self.logger.info(f"Model has Tables: {hasattr(self._model, 'Tables')}")
            
            if hasattr(self._model, 'Tables'):
                self.logger.info(f"Model.Tables.Count: {self._model.Tables.Count}")
            
            # Get friendly model name (like DAX Studio does)
            try:
                # Extract port from server address
                port = None
                if ':' in server:
                    try:
                        port = int(server.split(':')[-1])
                    except:
                        pass
                
                friendly_name = None
                
                # 1. Try to get .pbix filename from Power BI Desktop window title
                if port:
                    pbix_name = self._get_pbix_filename_from_process(port)
                    if pbix_name:
                        friendly_name = pbix_name
                
                # 2. Try Database.FriendlyName property
                if not friendly_name and hasattr(self._database, 'FriendlyName') and self._database.FriendlyName:
                    friendly_name = self._database.FriendlyName
                
                # 3. Try Model.Name
                if not friendly_name and self._model and hasattr(self._model, 'Name'):
                    friendly_name = self._model.Name
                
                # Ultimate fallback
                if not friendly_name:
                    friendly_name = self._database.Name
                    
            except Exception as e:
                self.logger.debug(f"Error getting friendly name: {e}")
                friendly_name = self._database.Name
            
            # Invalidate cache on new connection
            self._invalidate_cache()

            # Create connection info
            self.current_connection = ModelConnection(
                server=server,
                database=database,
                model_name=friendly_name,  # Use friendly name from Database.FriendlyName
                is_connected=True
            )

            self.logger.info(f"Connected to model: {friendly_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Connection failed: {e}", exc_info=True)
            return False
    
    def disconnect(self):
        """Disconnect from current model"""
        if self._server:
            try:
                self._server.Disconnect()
            except:
                pass

        self._server = None
        self._database = None
        self._model = None
        self._access_token = None
        self._is_cloud_connection = False
        if self.current_connection:
            self.current_connection.is_connected = False
        self.current_connection = None

        self.logger.info("Disconnected from model")

    # =========================================================================
    # Cloud/XMLA Endpoint Connection Methods
    # =========================================================================

    def _acquire_token_interactive(self) -> Optional[str]:
        """
        Acquire Azure AD access token using interactive browser authentication.

        Returns:
            Access token string or None if authentication failed
        """
        try:
            from msal import PublicClientApplication

            self.logger.info("Starting Azure AD interactive authentication...")

            # Create MSAL public client app
            app = PublicClientApplication(
                client_id=self.AZURE_CLIENT_ID,
                authority=self.AZURE_AUTHORITY
            )

            # Try to get token from cache first
            accounts = app.get_accounts()
            if accounts:
                self.logger.info(f"Found {len(accounts)} cached account(s), attempting silent auth")
                result = app.acquire_token_silent(self.AZURE_SCOPES, account=accounts[0])
                if result and "access_token" in result:
                    self.logger.info("✅ Acquired token from cache")
                    return result["access_token"]

            # Interactive login - opens browser
            self.logger.info("Opening browser for Azure AD login...")
            result = app.acquire_token_interactive(
                scopes=self.AZURE_SCOPES,
                prompt="select_account"  # Always show account picker
            )

            if result and "access_token" in result:
                self.logger.info("✅ Azure AD authentication successful")
                return result["access_token"]
            else:
                error = result.get("error_description", result.get("error", "Unknown error"))
                self.logger.error(f"❌ Azure AD authentication failed: {error}")
                return None

        except ImportError:
            self.logger.error("❌ MSAL library not installed. Install with: pip install msal")
            return None
        except Exception as e:
            self.logger.error(f"❌ Azure AD authentication error: {e}", exc_info=True)
            return None

    def connect_xmla(self, xmla_endpoint: str, dataset_name: Optional[str] = None) -> Tuple[bool, str]:
        """
        Connect to a Power BI Service model via XMLA endpoint.

        XMLA endpoint format:
            powerbi://api.powerbi.com/v1.0/myorg/WorkspaceName

        Args:
            xmla_endpoint: The XMLA endpoint URL (from Power BI workspace settings)
            dataset_name: Optional specific dataset name. If not provided, will list available.

        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            if not hasattr(self, 'Server'):
                return False, "TOM not initialized - cannot connect"

            # Use existing token if available, otherwise acquire new one
            if self._access_token:
                self.logger.info("Reusing existing Azure AD token...")
                access_token = self._access_token
            else:
                self.logger.info("Authenticating with Azure AD...")
                access_token = self._acquire_token_interactive()

                if not access_token:
                    return False, "Azure AD authentication failed. Please try again."

                self._access_token = access_token

            # Clean up the endpoint URL
            xmla_endpoint = xmla_endpoint.strip()
            if xmla_endpoint.endswith('/'):
                xmla_endpoint = xmla_endpoint[:-1]

            self.logger.info(f"Connecting to XMLA endpoint: {xmla_endpoint}")

            # Create server connection with OAuth token
            self._server = self.Server()

            # Connection string for Power BI Service with OAuth
            connection_string = (
                f"Provider=MSOLAP;"
                f"Data Source={xmla_endpoint};"
                f"Password={access_token};"
                f"Persist Security Info=True;"
                f"Impersonation Level=Impersonate"
            )

            self._server.Connect(connection_string)
            self.logger.info("✅ Connected to XMLA endpoint")

            # List available databases if no specific one requested
            if not dataset_name:
                db_names = []
                for db in self._server.Databases:
                    db_names.append(db.Name)
                self.logger.info(f"Available datasets: {db_names}")

                if len(db_names) == 0:
                    return False, "No datasets found in workspace. Check permissions."
                elif len(db_names) == 1:
                    dataset_name = db_names[0]
                    self.logger.info(f"Auto-selecting single dataset: {dataset_name}")
                else:
                    return False, f"Multiple datasets found. Please specify one:\n{chr(10).join(db_names)}"

            # Get the specific database
            self._database = self._server.Databases.FindByName(dataset_name)
            if not self._database:
                return False, f"Dataset '{dataset_name}' not found in workspace"

            self._model = self._database.Model
            self._is_cloud_connection = True

            # Invalidate cache on new connection
            self._invalidate_cache()

            # Create connection info
            self.current_connection = ModelConnection(
                server=xmla_endpoint,
                database=dataset_name,
                model_name=f"{dataset_name} (Cloud)",
                is_connected=True
            )

            self.logger.info(f"✅ Connected to cloud dataset: {dataset_name}")
            return True, f"Connected to '{dataset_name}'"

        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"❌ XMLA connection failed: {error_msg}", exc_info=True)

            # Provide helpful error messages
            if "403" in error_msg or "Forbidden" in error_msg:
                return False, (
                    "Access denied. Ensure:\n"
                    "1. XMLA read-write is enabled in workspace settings\n"
                    "2. You have Contributor or higher permissions\n"
                    "3. The workspace is on Premium or PPU capacity"
                )
            elif "401" in error_msg or "Unauthorized" in error_msg:
                return False, "Authentication failed. Please try signing in again."
            elif "not found" in error_msg.lower():
                return False, f"Workspace or dataset not found. Check the XMLA endpoint URL."
            else:
                return False, f"Connection failed: {error_msg}"

    def get_cloud_datasets(self, xmla_endpoint: str) -> Tuple[bool, List[str], str]:
        """
        Get list of available datasets from a cloud workspace.

        Args:
            xmla_endpoint: The XMLA endpoint URL

        Returns:
            Tuple of (success: bool, dataset_names: List[str], error_message: str)
        """
        try:
            if not hasattr(self, 'Server'):
                return False, [], "TOM not initialized"

            # Use existing token if available, otherwise acquire new one
            if not self._access_token:
                self._access_token = self._acquire_token_interactive()
                if not self._access_token:
                    return False, [], "Azure AD authentication failed"

            # Clean endpoint
            xmla_endpoint = xmla_endpoint.strip().rstrip('/')

            # Connect to server
            server = self.Server()
            connection_string = (
                f"Provider=MSOLAP;"
                f"Data Source={xmla_endpoint};"
                f"Password={self._access_token};"
                f"Persist Security Info=True;"
                f"Impersonation Level=Impersonate"
            )

            server.Connect(connection_string)

            # Get all database names
            datasets = [db.Name for db in server.Databases]

            server.Disconnect()

            return True, datasets, ""

        except Exception as e:
            return False, [], str(e)

    def is_cloud_connection(self) -> bool:
        """Check if current connection is to Power BI Service (cloud)"""
        return self._is_cloud_connection

    def is_connected(self) -> bool:
        """Check if currently connected to a model"""
        return self.current_connection is not None and self.current_connection.is_connected
    
    def get_tables(self, use_cache: bool = True) -> List[TableInfo]:
        """Get all tables from the model (cached)"""
        if not self.is_connected() or not self._model:
            self.logger.warning("get_tables: Not connected or no model")
            return []

        # Return cached result if available
        if use_cache and self._tables_cache is not None:
            self.logger.debug(f"Returning {len(self._tables_cache)} tables from cache")
            return self._tables_cache

        tables = []
        try:
            self.logger.info(f"Getting tables from model. Model type: {type(self._model)}")
            self.logger.info(f"Model has Tables collection: {hasattr(self._model, 'Tables')}")
            
            if hasattr(self._model, 'Tables'):
                table_count = self._model.Tables.Count
                self.logger.info(f"Model has {table_count} tables")
                
                for table in self._model.Tables:
                    # Determine table type
                    if hasattr(table, 'Partitions') and table.Partitions.Count > 0:
                        # Get first partition - can't use [0] on .NET NamedMetadataObjectCollection
                        partition = None
                        for p in table.Partitions:
                            partition = p
                            break
                        
                        if partition and hasattr(partition, 'SourceType'):
                            table_type = str(partition.SourceType)
                        else:
                            table_type = "Table"
                    else:
                        table_type = "Table"
                    
                    # Detect if this is a measures-only table
                    # (all columns hidden/deleted, but has visible measures)
                    is_measures_only = False
                    if hasattr(table, 'Columns') and hasattr(table, 'Measures'):
                        # Count visible columns (excluding RowNumber type)
                        visible_columns = 0
                        for col in table.Columns:
                            if hasattr(col, 'Type') and col.Type.ToString() == "RowNumber":
                                continue
                            if not (hasattr(col, 'IsHidden') and col.IsHidden):
                                visible_columns += 1
                        
                        # Count visible measures
                        visible_measures = 0
                        for measure in table.Measures:
                            if not (hasattr(measure, 'IsHidden') and measure.IsHidden):
                                visible_measures += 1
                        
                        # Measures-only if no visible columns but has visible measures
                        is_measures_only = (visible_columns == 0 and visible_measures > 0)
                    
                    tables.append(TableInfo(
                        name=table.Name,
                        is_hidden=table.IsHidden if hasattr(table, 'IsHidden') else False,
                        table_type=table_type,
                        is_measures_only=is_measures_only
                    ))
                
                self.logger.info(f"Retrieved {len(tables)} tables")
            else:
                self.logger.error("Model does not have Tables collection!")

            # Cache the result
            self._tables_cache = tables
            return tables

        except Exception as e:
            self.logger.error(f"Error getting tables: {e}", exc_info=True)
            return []
    
    def get_table_measures(self, table_name: str) -> List[FieldInfo]:
        """Get all measures from a specific table"""
        if not self.is_connected() or not self._model:
            return []
        
        measures = []
        try:
            table = self._model.Tables.Find(table_name)
            if not table:
                self.logger.warning(f"Table not found: {table_name}")
                return []
            
            for measure in table.Measures:
                measures.append(FieldInfo(
                    name=measure.Name,
                    table_name=table_name,
                    field_type="Measure",
                    expression=measure.Expression if hasattr(measure, 'Expression') else None,
                    is_hidden=measure.IsHidden if hasattr(measure, 'IsHidden') else False,
                    display_folder=measure.DisplayFolder if hasattr(measure, 'DisplayFolder') else ""
                ))
            
            self.logger.debug(f"Retrieved {len(measures)} measures from {table_name}")
            return measures
            
        except Exception as e:
            self.logger.error(f"Error getting measures from {table_name}: {e}")
            return []
    
    def get_table_columns(self, table_name: str) -> List[FieldInfo]:
        """Get all columns from a specific table"""
        if not self.is_connected() or not self._model:
            return []
        
        columns = []
        try:
            table = self._model.Tables.Find(table_name)
            if not table:
                self.logger.warning(f"Table not found: {table_name}")
                return []
            
            for column in table.Columns:
                # Skip system columns
                if column.Type.ToString() == "RowNumber":
                    continue
                
                columns.append(FieldInfo(
                    name=column.Name,
                    table_name=table_name,
                    field_type="Column",
                    data_type=column.DataType.ToString() if hasattr(column, 'DataType') else None,
                    expression=column.Expression if hasattr(column, 'Expression') else None,
                    is_hidden=column.IsHidden if hasattr(column, 'IsHidden') else False,
                    display_folder=column.DisplayFolder if hasattr(column, 'DisplayFolder') else ""
                ))
            
            self.logger.debug(f"Retrieved {len(columns)} columns from {table_name}")
            return columns
            
        except Exception as e:
            self.logger.error(f"Error getting columns from {table_name}: {e}")
            return []
    
    def get_all_fields_by_table(self, use_cache: bool = True) -> Dict[str, TableFieldsInfo]:
        """
        Get all measures and columns organized by table with complete metadata (cached).

        Returns:
            Dict mapping table name to TableFieldsInfo with fields, folders, and sort priority
        """
        # Return cached result if available
        if use_cache and self._fields_cache is not None:
            self.logger.debug(f"Returning fields from cache ({len(self._fields_cache)} tables)")
            return self._fields_cache

        result = {}

        tables = self.get_tables()
        for table in tables:
            # Skip hidden tables by default (can be made configurable)
            if table.is_hidden:
                continue
            
            all_fields = []
            
            # Get measures
            measures = self.get_table_measures(table.name)
            all_fields.extend([m for m in measures if not m.is_hidden])
            
            # Get columns
            columns = self.get_table_columns(table.name)
            all_fields.extend([c for c in columns if not c.is_hidden])
            
            if all_fields:  # Only include tables with visible fields
                # Measures-only tables get priority 0 (top), regular tables get priority 1
                sort_priority = 0 if table.is_measures_only else 1

                result[table.name] = TableFieldsInfo(
                    table_name=table.name,
                    fields=all_fields,
                    is_measures_only=table.is_measures_only,
                    sort_priority=sort_priority
                )

        # Cache the result
        self._fields_cache = result
        return result
    
    def detect_field_parameters(self) -> List[str]:
        """
        Detect existing field parameters in the model.
        
        Field parameters are identified by:
        1. Being calculated tables
        2. Having a column with extendedProperty ParameterMetadata
        3. Typically starting with ".Parameter" or "Parameter"
        
        Returns:
            List of field parameter table names
        """
        if not self.is_connected() or not self._model:
            return []
        
        parameters = []
        try:
            for table in self._model.Tables:
                # Check if it's a calculated table
                is_calculated = False
                if hasattr(table, 'Partitions') and table.Partitions.Count > 0:
                    # Get first partition - can't use [0] on .NET collection
                    partition = None
                    for p in table.Partitions:
                        partition = p
                        break
                    if partition and hasattr(partition, 'SourceType'):
                        is_calculated = str(partition.SourceType) == "Calculated"
                
                if not is_calculated:
                    continue
                
                # Check for ParameterMetadata extended property
                for column in table.Columns:
                    if hasattr(column, 'ExtendedProperties'):
                        for prop in column.ExtendedProperties:
                            if prop.Name == "ParameterMetadata":
                                parameters.append(table.Name)
                                break
                        if table.Name in parameters:
                            break
            
            self.logger.info(f"Detected {len(parameters)} field parameters")
            return parameters
            
        except Exception as e:
            self.logger.error(f"Error detecting field parameters: {e}")
            return []
    
    def get_table_tmdl(self, table_name: str) -> Optional[str]:
        """
        Get TMDL definition for a specific table.
        
        Constructs a minimal TMDL structure from TOM data that includes:
        - Table declaration
        - Column definitions with lineage tags
        - Partition with DAX expression
        - Annotations (PBI_Id, etc.)
        
        This format is compatible with FieldParameterParser.parse_tmdl().
        
        Args:
            table_name: Name of the table
            
        Returns:
            TMDL string or None if table not found
        """
        if not self.is_connected() or not self._model:
            return None
        
        try:
            table = self._model.Tables.Find(table_name)
            if not table:
                return None
            
            # TOM has serialization capabilities, but TMDL format requires
            # additional libraries or manual formatting
            # For field parameters, we can reconstruct from the partition source
            
            # Get partition source (the DAX table expression)
            if table.Partitions.Count > 0:
                # Get first partition - can't use [0] on .NET collection
                partition = None
                for p in table.Partitions:
                    partition = p
                    break
                if partition and hasattr(partition, 'Source') and hasattr(partition.Source, 'Expression'):
                    source_expression = partition.Source.Expression
                    
                    # Build minimal TMDL structure that the parser can understand
                    tmdl_lines = []
                    
                    # Table declaration with proper quoting
                    if ' ' in table_name or '-' in table_name:
                        tmdl_lines.append(f"table '{table_name}'")
                    else:
                        tmdl_lines.append(f"table {table_name}")
                    
                    # Add lineage tag if present
                    if hasattr(table, 'LineageTag') and table.LineageTag:
                        tmdl_lines.append(f"    lineageTag: {table.LineageTag}")
                    
                    # Add columns with metadata
                    for column in table.Columns:
                        if column.Type.ToString() == "RowNumber":
                            continue  # Skip system columns
                        
                        col_name = column.Name
                        if ' ' in col_name or '-' in col_name:
                            tmdl_lines.append(f"    column '{col_name}'")
                        else:
                            tmdl_lines.append(f"    column {col_name}")
                        
                        # Add data type
                        if hasattr(column, 'DataType'):
                            tmdl_lines.append(f"        dataType: {column.DataType}")
                        
                        # Add lineage tag
                        if hasattr(column, 'LineageTag') and column.LineageTag:
                            tmdl_lines.append(f"        lineageTag: {column.LineageTag}")
                        
                        # Check for ParameterMetadata extended property (indicates it's the main column)
                        if hasattr(column, 'ExtendedProperties'):
                            for prop in column.ExtendedProperties:
                                if prop.Name == "ParameterMetadata":
                                    tmdl_lines.append(f"        relatedColumnDetails")
                                    tmdl_lines.append(f"            groupByColumn: ...")
                    
                    # Add partition with DAX expression
                    part_name = partition.Name
                    if ' ' in part_name or '-' in part_name:
                        tmdl_lines.append(f"    partition '{part_name}' = calculated")
                    else:
                        tmdl_lines.append(f"    partition {part_name} = calculated")
                    
                    # Use 'source' instead of 'expression' to match parser expectations
                    tmdl_lines.append(f"        source = {source_expression}")
                    
                    # Add PBI_Id annotation if present
                    if hasattr(table, 'Annotations'):
                        for annotation in table.Annotations:
                            if annotation.Name == "PBI_Id":
                                tmdl_lines.append(f"    annotation PBI_Id = {annotation.Value}")
                    
                    return "\n".join(tmdl_lines)
            
            return None
            
        except Exception as e:
            self.logger.error(f"Error getting TMDL for {table_name}: {e}")
            return None
    
    def load_parameter_from_tom(self, table_name: str) -> Optional['FieldParameter']:
        """
        Load field parameter directly from TOM (bypasses TMDL parsing).
        
        This is the preferred method as it:
        - Reads directly from the partition DAX expression
        - Avoids fragile text parsing
        - Prepares for future save-back functionality
        
        Args:
            table_name: Name of the field parameter table
            
        Returns:
            FieldParameter object or None if not found/invalid
        """
        if not self.is_connected() or not self._model:
            self.logger.error("Not connected to model")
            return None
        
        try:
            # Import here to avoid circular dependency
            from tools.field_parameters.field_parameters_core import FieldParameter, FieldItem, CategoryLevel
            
            # Get table from model
            table = self._model.Tables.Find(table_name)
            if not table:
                self.logger.error(f"Table '{table_name}' not found in model")
                return None
            
            self.logger.info(f"Loading parameter table: {table_name}")
            
            # Get partition source (DAX expression)
            if table.Partitions.Count == 0:
                self.logger.error(f"Table '{table_name}' has no partitions")
                return None
            
            # Get first partition
            partition = None
            for p in table.Partitions:
                partition = p
                break
            
            if not partition or not hasattr(partition, 'Source') or not hasattr(partition.Source, 'Expression'):
                self.logger.error(f"Partition has no source expression")
                return None
            
            dax_expression = partition.Source.Expression.strip()
            self.logger.info(f"DAX expression length: {len(dax_expression)} chars")
            
            # Find parameter name (main display column with relatedColumnDetails)
            parameter_name = None
            for column in table.Columns:
                # Check for relatedColumnDetails or look for the non-hidden base column
                if not column.IsHidden and column.Name and 'Fields' not in column.Name and 'Order' not in column.Name:
                    parameter_name = column.Name
                    break
            
            if not parameter_name:
                # Fallback: use table name without prefix
                parameter_name = table_name.replace('.Parameter - ', '').replace('Parameter - ', '')
            
            self.logger.info(f"Parameter name: {parameter_name}")
            
            # Extract lineage tags
            lineage_tags = {}
            if hasattr(table, 'LineageTag') and table.LineageTag:
                lineage_tags['_table'] = table.LineageTag
            
            for column in table.Columns:
                if hasattr(column, 'LineageTag') and column.LineageTag:
                    lineage_tags[column.Name] = column.LineageTag
            
            # Extract PBI_Id
            pbi_id = None
            if hasattr(table, 'Annotations'):
                for annotation in table.Annotations:
                    if annotation.Name == 'PBI_Id':
                        pbi_id = annotation.Value
                        break
            
            # Parse DAX tuples from expression (returns fields and detected tuple format)
            fields, category_tuple_format = self._parse_dax_tuples(dax_expression)
            self.logger.info(f"Parsed {len(fields)} fields from DAX (format: {category_tuple_format})")

            # Detect category levels from columns
            category_levels = self._detect_category_levels(table, parameter_name)
            self.logger.info(f"Detected {len(category_levels)} category levels")

            # Extract unique labels for each category level from parsed fields
            self._populate_category_labels(category_levels, fields)
            for level in category_levels:
                self.logger.info(f"Category '{level.name}': {len(level.labels)} labels - {level.labels}")

            # Build FieldParameter object
            field_param = FieldParameter(
                table_name=table_name,
                parameter_name=parameter_name,
                fields=fields,
                category_levels=category_levels,
                keep_lineage_tags=True,
                lineage_tags=lineage_tags,
                pbi_id=pbi_id,
                category_tuple_format=category_tuple_format
            )
            
            return field_param
            
        except Exception as e:
            self.logger.error(f"Error loading parameter from TOM: {e}", exc_info=True)
            return None
    
    def _parse_dax_tuples(self, dax_expression: str) -> Tuple[List['FieldItem'], str]:
        """
        Parse DAX table expression tuples into FieldItem objects.

        Format: ("Display Name", NAMEOF('Table'[Field]), order, "Cat1", cat1_sort, "Cat2", cat2_sort, ...)

        Returns:
            Tuple of (list of FieldItem, category_tuple_format)
            category_tuple_format is "sort_first" (1, "Label") or "label_first" ("Label", 1)
        """
        from tools.field_parameters.field_parameters_core import FieldItem
        import re

        fields = []
        detected_format = "sort_first"  # Default
        format_votes = {"sort_first": 0, "label_first": 0}  # Vote for format based on what we see

        # Extract content between { }
        match = re.search(r'\{([^}]+)\}', dax_expression, re.DOTALL)
        if not match:
            self.logger.warning("Could not find { } in DAX expression")
            return fields, detected_format

        content = match.group(1)

        # Pattern to match tuples - handle multi-line
        # Format: ("Display", NAMEOF('Table'[Field]), order, additional values...)
        tuple_pattern = r'\(\s*"([^"]+)"\s*,\s*NAMEOF\(\'([^\']+)\'\[([^\]]+)\]\)\s*,\s*(\d+)([^)]*)\)'

        for match in re.finditer(tuple_pattern, content):
            display_name = match.group(1)
            table_name = match.group(2)
            field_name = match.group(3)
            order = int(match.group(4))
            remaining = match.group(5)  # Everything after order

            # Parse categories from remaining string
            # Categories come in pairs after the order number
            # Format can be either:
            #   - "label", sort (e.g., "Geography", 1) - label_first
            #   - sort, "label" (e.g., 1, "Geography") - sort_first (Power BI native)
            # We detect which format by looking at what comes first after commas
            categories = []

            # Extract all values after the order (remaining string)
            # Split by comma and parse alternating pairs
            if remaining.strip():
                # Find all quoted strings and numbers in order
                tokens = []
                # Match either quoted strings or numbers (including decimals like 10.1)
                token_pattern = r'"([^"]+)"|(\d+(?:\.\d+)?)'
                for token_match in re.finditer(token_pattern, remaining):
                    if token_match.group(1):  # quoted string
                        tokens.append(('str', token_match.group(1)))
                    elif token_match.group(2):  # number (int or decimal)
                        # Store as float to handle decimals, but keep as numeric
                        tokens.append(('num', float(token_match.group(2))))

                # Process tokens in pairs and detect format
                i = 0
                while i + 1 < len(tokens):
                    first = tokens[i]
                    second = tokens[i + 1]

                    if first[0] == 'str' and second[0] == 'num':
                        # Format: "label", sort (label_first)
                        categories.append((second[1], first[1]))
                        format_votes["label_first"] += 1
                    elif first[0] == 'num' and second[0] == 'str':
                        # Format: sort, "label" (sort_first)
                        categories.append((first[1], second[1]))
                        format_votes["sort_first"] += 1
                    # else: mismatched pair, skip

                    i += 2

            field_item = FieldItem(
                display_name=display_name,
                field_reference=f"NAMEOF('{table_name}'[{field_name}])",
                table_name=table_name,
                field_name=field_name,
                order_within_group=order,
                original_order_within_group=order,  # Preserve original for custom sort
                categories=categories
            )
            fields.append(field_item)

            self.logger.debug(f"Parsed field: {display_name} (order {order}, {len(categories)} categories)")

        # Determine format based on votes (majority wins)
        if format_votes["label_first"] > format_votes["sort_first"]:
            detected_format = "label_first"
        else:
            detected_format = "sort_first"

        self.logger.info(f"Detected category tuple format: {detected_format} (votes: {format_votes})")

        return fields, detected_format
    
    def _detect_category_levels(self, table, parameter_name: str) -> List['CategoryLevel']:
        """
        Detect category levels from table columns using SortByColumn relationships.

        Uses the SortByColumn property to intelligently identify category columns:
        - A category display column is one that has a SortByColumn pointing to another column
        - The column it points to is the sort column (hidden, numeric)
        - This approach is name-independent - works regardless of column naming conventions

        Standard field parameter columns are identified by:
        - Main display column: matches parameter_name exactly
        - Fields column: hidden, has ParameterMetadata extended property
        - Order column: the main display column's SortByColumn

        Calculated columns (with Expression attribute) are marked as is_calculated=True
        and cannot be edited by this tool.
        """
        from tools.field_parameters.field_parameters_core import CategoryLevel

        category_levels = []

        # PASS 1: Identify the standard field parameter columns
        # These are: main display, fields (hidden with ParameterMetadata), and order (sort column for main)

        main_display_col = None
        fields_col = None
        main_sort_col = None  # The Order column (what main display sorts by)

        # First find the fields column (has ParameterMetadata)
        for column in table.Columns:
            if column.Type.ToString() == "RowNumber":
                continue
            if hasattr(column, 'ExtendedProperties'):
                for prop in column.ExtendedProperties:
                    if prop.Name == "ParameterMetadata":
                        fields_col = column
                        self.logger.info(f"Found Fields column: {column.Name}")
                        break
                if fields_col:
                    break

        # Then find the main display column (matches parameter_name or is visible with SortByColumn)
        for column in table.Columns:
            if column.Type.ToString() == "RowNumber":
                continue
            if column.IsHidden:
                continue
            if column == fields_col:
                continue

            # Exact match to parameter name is definitive
            if column.Name == parameter_name:
                main_display_col = column
                if hasattr(column, 'SortByColumn') and column.SortByColumn:
                    main_sort_col = column.SortByColumn
                self.logger.info(f"Found Main display column (exact match): {column.Name}")
                break

        # If no exact match, find by ParameterMetadata relationship or first visible with sort
        if not main_display_col:
            for column in table.Columns:
                if column.Type.ToString() == "RowNumber":
                    continue
                if column.IsHidden:
                    continue
                if column == fields_col:
                    continue

                # Check if this column has relatedColumnDetails pointing to fields column
                # Or just take first visible column with SortByColumn as fallback
                if hasattr(column, 'SortByColumn') and column.SortByColumn:
                    main_display_col = column
                    main_sort_col = column.SortByColumn
                    self.logger.info(f"Found Main display column (fallback): {column.Name}")
                    break

        self.logger.info(f"Standard columns - Main: {main_display_col.Name if main_display_col else 'None'}, "
                        f"Fields: {fields_col.Name if fields_col else 'None'}, "
                        f"Order: {main_sort_col.Name if main_sort_col else 'None'}")

        # Build set of standard column names to exclude
        standard_columns = set()
        if main_display_col:
            standard_columns.add(main_display_col.Name)
        if fields_col:
            standard_columns.add(fields_col.Name)
        if main_sort_col:
            standard_columns.add(main_sort_col.Name)

        # PASS 2: Find category columns
        # A category column is a visible column (not in standard set) that has a SortByColumn
        # Also collect all hidden sort columns to exclude them

        sort_columns = set()  # Track which columns are used as sort columns

        # First, identify all sort columns (columns that other columns sort by)
        for column in table.Columns:
            if column.Type.ToString() == "RowNumber":
                continue
            if hasattr(column, 'SortByColumn') and column.SortByColumn:
                sort_columns.add(column.SortByColumn.Name)

        self.logger.debug(f"Identified sort columns: {sort_columns}")

        # Now find category columns
        # Log all columns for debugging
        self.logger.info(f"Scanning all columns for categories. Standard columns to skip: {standard_columns}")

        # Also check for columns in hierarchies - these might be category columns
        # that appear hidden at the column level but are used in a hierarchy
        hierarchy_columns = set()
        if hasattr(table, 'Hierarchies'):
            for hierarchy in table.Hierarchies:
                self.logger.info(f"  Found hierarchy: '{hierarchy.Name}'")
                if hasattr(hierarchy, 'Levels'):
                    for level in hierarchy.Levels:
                        if hasattr(level, 'Column') and level.Column:
                            col_name = level.Column.Name
                            hierarchy_columns.add(col_name)
                            self.logger.info(f"    Hierarchy level column: '{col_name}'")

        for column in table.Columns:
            col_name = column.Name
            col_type = column.Type.ToString() if hasattr(column, 'Type') else "Unknown"
            is_hidden = column.IsHidden if hasattr(column, 'IsHidden') else False
            in_hierarchy = col_name in hierarchy_columns

            self.logger.info(f"  Column: '{col_name}' | Type: {col_type} | Hidden: {is_hidden} | InHierarchy: {in_hierarchy}")

            # Skip system columns
            if col_type == "RowNumber":
                self.logger.info(f"    -> Skipped (RowNumber system column)")
                continue

            # Skip standard parameter columns
            if col_name in standard_columns:
                self.logger.info(f"    -> Skipped (standard parameter column)")
                continue

            # Skip hidden columns UNLESS they're in a hierarchy (hierarchy columns are often hidden)
            # A column in a hierarchy with a SortByColumn is likely a category column
            if is_hidden and not in_hierarchy:
                self.logger.info(f"    -> Skipped (hidden column, not in hierarchy)")
                continue

            # For hidden columns in hierarchies, check if they have a SortByColumn
            # If hidden but NOT in hierarchy, skip (it's just a sort column)
            if is_hidden and in_hierarchy:
                # Check if this column is used as a sort column by another column
                is_sort_column = col_name in sort_columns
                if is_sort_column:
                    self.logger.info(f"    -> Skipped (hidden sort column in hierarchy)")
                    continue
                # Otherwise, it's a category column that's hidden but in a hierarchy
                self.logger.info(f"    -> Category column (hidden but in hierarchy)")

            # Skip Value columns (unnamed intermediate columns)
            if col_name.startswith('Value'):
                self.logger.info(f"    -> Skipped (Value column)")
                continue

            # Check if this is a calculated column (has Expression attribute with DAX code)
            is_calculated = False
            if hasattr(column, 'Expression') and column.Expression:
                is_calculated = True

            # This is a category display column
            sort_by_col = None
            if hasattr(column, 'SortByColumn') and column.SortByColumn:
                sort_by_col = column.SortByColumn
                self.logger.info(f"    -> CATEGORY COLUMN! Sorts by '{sort_by_col.Name}'")
            else:
                self.logger.info(f"    -> CATEGORY COLUMN! (no SortByColumn)")

            sort_order = len(category_levels) + 1

            if is_calculated:
                self.logger.info(f"    -> Type: Calculated (read-only)")
            else:
                self.logger.info(f"    -> Type: Inline (editable)")

            category_levels.append(CategoryLevel(
                name=col_name,
                sort_order=sort_order,
                column_name=col_name,
                is_calculated=is_calculated
            ))

        return category_levels

    def _populate_category_labels(self, category_levels: List['CategoryLevel'], fields: List['FieldItem']):
        """
        Extract unique labels for each category level from the parsed field data.

        The field items have categories like [(1, "Geography"), (2, "Time")] where
        the first value is the sort order and the second is the label.

        This method extracts unique labels for each category level position.
        """
        if not category_levels or not fields:
            return

        # For each category level (by position)
        for level_idx, level in enumerate(category_levels):
            if level.is_calculated:
                continue  # Skip calculated columns - they generate their own values

            # Collect unique labels for this category level position
            unique_labels = {}  # label -> sort_order (keep track of sort order for ordering)

            for field_item in fields:
                if field_item.categories and len(field_item.categories) > level_idx:
                    sort_order, label = field_item.categories[level_idx]
                    if label and label not in unique_labels:
                        unique_labels[label] = sort_order

            # Sort labels by their sort order and store
            sorted_labels = sorted(unique_labels.items(), key=lambda x: x[1])
            level.labels = [label for label, _ in sorted_labels]

            self.logger.debug(f"Category level {level_idx} '{level.name}': found {len(level.labels)} unique labels")

    def _set_table_annotation(self, table, name: str, value: str):
        """
        Set or update an annotation on a table.

        Args:
            table: The TOM Table object
            name: Annotation name
            value: Annotation value
        """
        try:
            # Check if annotation already exists
            existing = None
            if hasattr(table, 'Annotations'):
                for annotation in table.Annotations:
                    if annotation.Name == name:
                        existing = annotation
                        break

            if existing:
                # Update existing annotation
                existing.Value = value
                self.logger.info(f"Updated annotation '{name}' on table '{table.Name}'")
            else:
                # Create new annotation
                # Need to import the Annotation class
                try:
                    from Microsoft.AnalysisServices.Tabular import Annotation
                    new_annotation = Annotation()
                    new_annotation.Name = name
                    new_annotation.Value = value
                    table.Annotations.Add(new_annotation)
                    self.logger.info(f"Added annotation '{name}' to table '{table.Name}'")
                except ImportError:
                    self.logger.warning("Could not import Annotation class - annotation not added")
        except Exception as e:
            self.logger.error(f"Error setting annotation '{name}': {e}")

    def refresh_table(self, table_name: str, refresh_type: str = "Calculate") -> Tuple[bool, str]:
        """
        Refresh a specific table in the connected model.

        For calculated tables (like field parameters), use 'Calculate' to re-evaluate the DAX.
        For tables with data sources, use 'Full' for a complete refresh.

        Args:
            table_name: Name of the table to refresh
            refresh_type: Type of refresh - 'Calculate', 'Full', 'DataOnly', or 'Automatic'

        Returns:
            Tuple of (success: bool, message: str)
        """
        if not self.is_connected() or not self._model:
            return False, "Not connected to a model"

        try:
            # Import RefreshType enum from TOM
            from Microsoft.AnalysisServices.Tabular import RefreshType

            # Find the table
            table = self._model.Tables.Find(table_name)
            if not table:
                return False, f"Table '{table_name}' not found in model"

            # Map string to RefreshType enum
            refresh_type_map = {
                'Calculate': RefreshType.Calculate,
                'Full': RefreshType.Full,
                'DataOnly': RefreshType.DataOnly,
                'Automatic': RefreshType.Automatic,
            }

            rt = refresh_type_map.get(refresh_type, RefreshType.Calculate)

            self.logger.info(f"Requesting {refresh_type} refresh for table '{table_name}'")

            # Request refresh on the table
            table.RequestRefresh(rt)

            # Execute the refresh by saving changes
            self._model.SaveChanges()

            self.logger.info(f"Successfully refreshed table '{table_name}'")
            return True, f"Table '{table_name}' refreshed successfully"

        except Exception as e:
            self.logger.error(f"Error refreshing table '{table_name}': {e}", exc_info=True)
            return False, f"Error refreshing table: {str(e)}"

    def apply_parameter_to_model(self, parameter: 'FieldParameter', auto_refresh_cloud: bool = True) -> Tuple[bool, str]:
        """
        Apply field parameter changes to the connected model via TOM.

        This creates or updates the field parameter table in the live model.
        The model must be connected and the table will be created/modified directly.

        For cloud connections, automatically triggers a table refresh after saving
        so the DAX calculated table is evaluated immediately.

        Args:
            parameter: The FieldParameter object with all configuration
            auto_refresh_cloud: If True and connected to cloud, automatically refresh
                               the table after saving (default: True)

        Returns:
            Tuple of (success: bool, message: str)
        """
        if not self.is_connected() or not self._model:
            return False, "Not connected to a model"

        try:
            from tools.field_parameters.field_parameters_core import FieldParameterGenerator

            # Generate the DAX expression for the partition (properly formatted for TOM)
            dax_expression = FieldParameterGenerator.generate_dax_expression(parameter)

            # Check if table already exists
            existing_table = self._model.Tables.Find(parameter.table_name)

            if existing_table:
                # Update existing table - modify the partition source
                self.logger.info(f"Updating existing table: {parameter.table_name}")

                if existing_table.Partitions.Count > 0:
                    # Get first partition
                    partition = None
                    for p in existing_table.Partitions:
                        partition = p
                        break

                    if partition and hasattr(partition, 'Source') and hasattr(partition.Source, 'Expression'):
                        partition.Source.Expression = dax_expression
                        self.logger.info("Updated partition expression")
                    else:
                        return False, "Could not find partition source to update"
                else:
                    return False, "Table has no partitions to update"

                # Add/update AE_MultiTool annotation
                self._set_table_annotation(existing_table, "AE_MultiTool", "Edited with Analytic Endeavors PBI Multi-Tool")

                # Save changes to the model
                self._model.SaveChanges()

                # For cloud connections, trigger table refresh so DAX is evaluated
                if auto_refresh_cloud and self.is_cloud_connection():
                    self.logger.info(f"Cloud connection detected - refreshing table '{parameter.table_name}'")
                    refresh_success, refresh_msg = self.refresh_table(parameter.table_name, "Calculate")
                    if refresh_success:
                        return True, f"Successfully updated and refreshed '{parameter.table_name}' in the cloud model"
                    else:
                        # Save succeeded but refresh failed - still report success with warning
                        return True, f"Successfully updated '{parameter.table_name}' in the model.\n\nNote: Auto-refresh failed ({refresh_msg}). You may need to manually refresh the table in Power BI Service."

                return True, f"Successfully updated '{parameter.table_name}' in the model"

            else:
                # Create new table - this is more complex as we need to create the full structure
                self.logger.info(f"Creating new table: {parameter.table_name}")

                # For new tables, we need to use the TOM API to create the structure
                # This requires creating Table, Columns, and Partition objects

                # Import TOM types
                try:
                    from Microsoft.AnalysisServices.Tabular import (
                        Table, Column, DataType, Partition, MPartitionSource
                    )
                except ImportError:
                    return False, "TOM types not available for creating new tables"

                # Create the table
                new_table = Table()
                new_table.Name = parameter.table_name

                # Add lineage tag if available
                if parameter.keep_lineage_tags and '_table' in parameter.lineage_tags:
                    new_table.LineageTag = parameter.lineage_tags['_table']

                # Create display column (main parameter column)
                display_col = Column()
                display_col.Name = parameter.parameter_name
                display_col.DataType = DataType.String
                display_col.SourceColumn = "[Value1]"
                display_col.SummarizeBy = "None"
                if parameter.keep_lineage_tags and parameter.parameter_name in parameter.lineage_tags:
                    display_col.LineageTag = parameter.lineage_tags[parameter.parameter_name]
                new_table.Columns.Add(display_col)

                # Create Fields column (hidden)
                fields_col = Column()
                fields_col.Name = f"{parameter.parameter_name} Fields"
                fields_col.DataType = DataType.String
                fields_col.SourceColumn = "[Value2]"
                fields_col.IsHidden = True
                new_table.Columns.Add(fields_col)

                # Create Order column (hidden)
                order_col = Column()
                order_col.Name = f"{parameter.parameter_name} Order"
                order_col.DataType = DataType.Int64
                order_col.SourceColumn = "[Value3]"
                order_col.IsHidden = True
                new_table.Columns.Add(order_col)

                # Create category columns if any
                value_idx = 4
                for cat_level in parameter.category_levels:
                    if cat_level.is_calculated:
                        continue  # Skip calculated columns

                    # Sort column
                    sort_col = Column()
                    sort_col.Name = f"{cat_level.column_name} Sort"
                    sort_col.DataType = DataType.Int64
                    sort_col.SourceColumn = f"[Value{value_idx}]"
                    sort_col.IsHidden = True
                    new_table.Columns.Add(sort_col)
                    value_idx += 1

                    # Display column
                    cat_col = Column()
                    cat_col.Name = cat_level.column_name
                    cat_col.DataType = DataType.String
                    cat_col.SourceColumn = f"[Value{value_idx}]"
                    new_table.Columns.Add(cat_col)
                    value_idx += 1

                # Create partition with DAX expression
                partition = Partition()
                partition.Name = parameter.table_name
                partition.Source = MPartitionSource()
                partition.Source.Expression = dax_expression
                new_table.Partitions.Add(partition)

                # Add table to model
                self._model.Tables.Add(new_table)

                # Save changes
                self._model.SaveChanges()

                # For cloud connections, trigger table refresh so DAX is evaluated
                if auto_refresh_cloud and self.is_cloud_connection():
                    self.logger.info(f"Cloud connection detected - refreshing new table '{parameter.table_name}'")
                    refresh_success, refresh_msg = self.refresh_table(parameter.table_name, "Calculate")
                    if refresh_success:
                        return True, f"Successfully created and refreshed '{parameter.table_name}' in the cloud model"
                    else:
                        # Save succeeded but refresh failed - still report success with warning
                        return True, f"Successfully created '{parameter.table_name}' in the model.\n\nNote: Auto-refresh failed ({refresh_msg}). You may need to manually refresh the table in Power BI Service."

                return True, f"Successfully created '{parameter.table_name}' in the model"

        except Exception as e:
            self.logger.error(f"Error applying parameter to model: {e}", exc_info=True)
            return False, f"Error: {str(e)}"

    def parse_external_tool_args(self, args: List[str]) -> Tuple[Optional[str], Optional[str]]:
        """
        Parse command line arguments from Power BI external tool launch.

        Args:
            args: sys.argv (command line arguments)

        Returns:
            Tuple of (server, database) or (None, None) if not found
            For thin reports, database may be empty string
        """
        # External tools receive: program.exe "localhost:port" "database-guid"
        # For thin reports, database may be empty or missing
        if len(args) >= 3:
            server = args[1].strip('"')
            database = args[2].strip('"')
            self.logger.info(f"Parsed external tool args: {server}, {database}")
            return server, database
        elif len(args) >= 2:
            # Thin report - only server passed, no database
            server = args[1].strip('"')
            self.logger.info(f"Parsed external tool args (thin report): {server}, (no database)")
            return server, ""

        self.logger.info("No external tool arguments found")
        return None, None
    
    def get_connection_info(self) -> Optional[ModelConnection]:
        """Get current connection information"""
        return self.current_connection


# Global connector instance
_connector: Optional[PowerBIConnector] = None


def get_connector() -> PowerBIConnector:
    """Get the global connector instance"""
    global _connector
    if _connector is None:
        _connector = PowerBIConnector()
    return _connector
