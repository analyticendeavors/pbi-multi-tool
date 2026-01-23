"""
Connection Health Checker - Background monitoring of connection accessibility
Built by Reid Havens of Analytic Endeavors

Monitors connection targets and provides live status indicators.
"""

import logging
import socket
import threading
import time
from typing import Dict, List, Callable, Optional
from enum import Enum

from tools.connection_hotswap.models import ConnectionMapping, SwapTarget


class HealthStatus(Enum):
    """Health status of a connection target."""
    UNKNOWN = "unknown"       # Not yet checked
    CHECKING = "checking"     # Check in progress
    HEALTHY = "healthy"       # Target is accessible
    UNHEALTHY = "unhealthy"  # Target is not accessible
    ERROR = "error"          # Error during check


class HealthCheckResult:
    """Result of a health check."""

    def __init__(
        self,
        target_id: str,
        status: HealthStatus,
        message: str = "",
        response_time_ms: int = 0
    ):
        self.target_id = target_id
        self.status = status
        self.message = message
        self.response_time_ms = response_time_ms
        self.last_checked = time.time()


class ConnectionHealthChecker:
    """
    Background health checker for connection targets.

    Periodically pings connection targets to verify accessibility
    and notifies callbacks of status changes.
    """

    DEFAULT_CHECK_INTERVAL = 30  # seconds
    SOCKET_TIMEOUT = 5  # seconds for socket check

    def __init__(
        self,
        check_interval: int = DEFAULT_CHECK_INTERVAL,
        on_status_change: Optional[Callable[[str, HealthCheckResult], None]] = None
    ):
        """
        Initialize the health checker.

        Args:
            check_interval: Seconds between health checks
            on_status_change: Callback when a target's status changes
        """
        self.logger = logging.getLogger(__name__)
        self.check_interval = check_interval
        self.on_status_change = on_status_change

        self._targets: Dict[str, SwapTarget] = {}
        self._results: Dict[str, HealthCheckResult] = {}
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.Lock()

    def add_target(self, target: SwapTarget) -> str:
        """
        Add a target to monitor.

        Args:
            target: SwapTarget to monitor

        Returns:
            Target ID for reference
        """
        target_id = f"{target.server}|{target.database}"
        with self._lock:
            self._targets[target_id] = target
            self._results[target_id] = HealthCheckResult(
                target_id=target_id,
                status=HealthStatus.UNKNOWN,
                message="Not yet checked"
            )
        return target_id

    def remove_target(self, target_id: str):
        """Remove a target from monitoring."""
        with self._lock:
            self._targets.pop(target_id, None)
            self._results.pop(target_id, None)

    def clear_targets(self):
        """Remove all targets."""
        with self._lock:
            self._targets.clear()
            self._results.clear()

    def get_status(self, target_id: str) -> Optional[HealthCheckResult]:
        """Get the current status of a target."""
        with self._lock:
            return self._results.get(target_id)

    def get_all_statuses(self) -> Dict[str, HealthCheckResult]:
        """Get all current statuses."""
        with self._lock:
            return dict(self._results)

    def check_target_now(self, target_id: str) -> HealthCheckResult:
        """
        Immediately check a specific target.

        Args:
            target_id: Target ID to check

        Returns:
            HealthCheckResult
        """
        with self._lock:
            target = self._targets.get(target_id)

        if not target:
            return HealthCheckResult(
                target_id=target_id,
                status=HealthStatus.ERROR,
                message="Target not found"
            )

        return self._perform_check(target_id, target)

    def check_all_now(self) -> Dict[str, HealthCheckResult]:
        """Check all targets immediately."""
        results = {}
        with self._lock:
            targets = dict(self._targets)

        for target_id, target in targets.items():
            result = self._perform_check(target_id, target)
            results[target_id] = result

        return results

    def start(self):
        """Start background health checking."""
        if self._running:
            return

        self._running = True
        self._thread = threading.Thread(target=self._check_loop, daemon=True)
        self._thread.start()
        self.logger.info(f"Health checker started (interval: {self.check_interval}s)")

    def stop(self):
        """Stop background health checking."""
        self._running = False
        if self._thread:
            self._thread.join(timeout=2)
            self._thread = None
        self.logger.info("Health checker stopped")

    def _check_loop(self):
        """Background loop that periodically checks all targets."""
        while self._running:
            try:
                self.check_all_now()
            except Exception as e:
                self.logger.error(f"Error in health check loop: {e}")

            # Sleep in small increments to allow quick shutdown
            for _ in range(int(self.check_interval)):
                if not self._running:
                    break
                time.sleep(1)

    def _perform_check(self, target_id: str, target: SwapTarget) -> HealthCheckResult:
        """
        Perform a health check on a target.

        Args:
            target_id: Target identifier
            target: SwapTarget to check

        Returns:
            HealthCheckResult
        """
        # Mark as checking
        with self._lock:
            if target_id in self._results:
                old_status = self._results[target_id].status
            else:
                old_status = HealthStatus.UNKNOWN

            self._results[target_id] = HealthCheckResult(
                target_id=target_id,
                status=HealthStatus.CHECKING,
                message="Checking..."
            )

        # Notify checking status
        if self.on_status_change:
            self.on_status_change(target_id, self._results[target_id])

        # Perform the check
        start_time = time.time()
        result = self._check_connectivity(target)
        elapsed_ms = int((time.time() - start_time) * 1000)

        result.target_id = target_id
        result.response_time_ms = elapsed_ms

        # Update stored result
        with self._lock:
            self._results[target_id] = result

        # Notify if status changed
        if self.on_status_change and result.status != old_status:
            self.on_status_change(target_id, result)

        return result

    def _check_connectivity(self, target: SwapTarget) -> HealthCheckResult:
        """
        Check if a target is accessible.

        For local targets: TCP socket connect to server:port
        For cloud targets: Check XMLA endpoint format validity

        Args:
            target: SwapTarget to check

        Returns:
            HealthCheckResult
        """
        try:
            if target.target_type == "local":
                return self._check_local_target(target)
            else:
                return self._check_cloud_target(target)

        except Exception as e:
            return HealthCheckResult(
                target_id="",
                status=HealthStatus.ERROR,
                message=f"Check error: {str(e)}"
            )

    def _check_local_target(self, target: SwapTarget) -> HealthCheckResult:
        """Check a local Power BI Desktop target."""
        try:
            # Parse server address
            if ':' in target.server:
                host, port_str = target.server.rsplit(':', 1)
                port = int(port_str)
            else:
                # Default SSAS port
                host = target.server
                port = 2383

            # Normalize localhost
            if host.lower() == 'localhost':
                host = '127.0.0.1'

            # TCP connect check
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(self.SOCKET_TIMEOUT)

            result = sock.connect_ex((host, port))
            sock.close()

            if result == 0:
                return HealthCheckResult(
                    target_id="",
                    status=HealthStatus.HEALTHY,
                    message=f"Connected to {host}:{port}"
                )
            else:
                return HealthCheckResult(
                    target_id="",
                    status=HealthStatus.UNHEALTHY,
                    message=f"Cannot connect to {host}:{port}"
                )

        except socket.timeout:
            return HealthCheckResult(
                target_id="",
                status=HealthStatus.UNHEALTHY,
                message="Connection timeout"
            )
        except Exception as e:
            return HealthCheckResult(
                target_id="",
                status=HealthStatus.ERROR,
                message=f"Error: {str(e)}"
            )

    def _check_cloud_target(self, target: SwapTarget) -> HealthCheckResult:
        """
        Check a cloud XMLA endpoint target.

        Note: Full XMLA connectivity requires authentication.
        This performs a basic format validation and DNS resolution.
        """
        try:
            # Validate XMLA endpoint format
            if not target.server.startswith("powerbi://"):
                return HealthCheckResult(
                    target_id="",
                    status=HealthStatus.ERROR,
                    message="Invalid XMLA endpoint format"
                )

            # Extract hostname from endpoint
            # Format: powerbi://api.powerbi.com/v1.0/myorg/WorkspaceName
            import urllib.parse
            parsed = urllib.parse.urlparse(target.server.replace("powerbi://", "https://"))
            hostname = parsed.hostname

            if not hostname:
                return HealthCheckResult(
                    target_id="",
                    status=HealthStatus.ERROR,
                    message="Cannot parse endpoint hostname"
                )

            # DNS resolution check
            try:
                socket.gethostbyname(hostname)
            except socket.gaierror:
                return HealthCheckResult(
                    target_id="",
                    status=HealthStatus.UNHEALTHY,
                    message=f"Cannot resolve {hostname}"
                )

            # Cloud endpoint is reachable (at DNS level)
            return HealthCheckResult(
                target_id="",
                status=HealthStatus.HEALTHY,
                message=f"XMLA endpoint valid ({hostname})"
            )

        except Exception as e:
            return HealthCheckResult(
                target_id="",
                status=HealthStatus.ERROR,
                message=f"Error: {str(e)}"
            )


def get_status_emoji(status: HealthStatus) -> str:
    """Get emoji representation of health status."""
    return {
        HealthStatus.UNKNOWN: "⚪",
        HealthStatus.CHECKING: "🟡",
        HealthStatus.HEALTHY: "🟢",
        HealthStatus.UNHEALTHY: "🔴",
        HealthStatus.ERROR: "⚠️",
    }.get(status, "⚪")


def get_status_color(status: HealthStatus, is_dark: bool = True) -> str:
    """Get color representation of health status."""
    if is_dark:
        return {
            HealthStatus.UNKNOWN: "#6b7280",
            HealthStatus.CHECKING: "#f59e0b",
            HealthStatus.HEALTHY: "#10b981",
            HealthStatus.UNHEALTHY: "#ef4444",
            HealthStatus.ERROR: "#f97316",
        }.get(status, "#6b7280")
    else:
        return {
            HealthStatus.UNKNOWN: "#9ca3af",
            HealthStatus.CHECKING: "#d97706",
            HealthStatus.HEALTHY: "#059669",
            HealthStatus.UNHEALTHY: "#dc2626",
            HealthStatus.ERROR: "#ea580c",
        }.get(status, "#9ca3af")
