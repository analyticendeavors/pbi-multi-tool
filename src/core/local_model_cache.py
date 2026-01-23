"""
Local Model Cache - Centralized cache for discovered local Power BI models
Built by Reid Havens of Analytic Endeavors

Shared between Connection Hot Swap and Field Parameters tools to avoid
redundant port scanning when switching between tools.
"""

import logging
import threading
import time
from typing import List, Optional, Callable, TYPE_CHECKING

if TYPE_CHECKING:
    from tools.field_parameters.models import AvailableModel


class LocalModelCache:
    """
    Singleton cache for discovered local Power BI Desktop models.

    Provides:
    - Shared cache between tools (scan once, use everywhere)
    - Staleness detection (auto-refresh after configurable timeout)
    - Thread-safe scanning with callbacks for UI updates
    - Optional callback registration for cache update notifications
    """

    _instance: Optional['LocalModelCache'] = None
    _lock = threading.Lock()

    # Default cache staleness timeout (seconds)
    DEFAULT_STALE_TIMEOUT = 60

    @classmethod
    def get_instance(cls) -> 'LocalModelCache':
        """Get the singleton instance of LocalModelCache."""
        if cls._instance is None:
            with cls._lock:
                # Double-check locking pattern
                if cls._instance is None:
                    cls._instance = cls()
        return cls._instance

    def __init__(self):
        """Initialize the cache. Use get_instance() instead of direct instantiation."""
        self.logger = logging.getLogger(__name__)
        self._cached_models: List['AvailableModel'] = []
        self._last_scan_time: Optional[float] = None
        self._scan_in_progress = False
        self._scan_lock = threading.Lock()
        self._update_callbacks: List[Callable[[List['AvailableModel']], None]] = []

    def get_models(self) -> List['AvailableModel']:
        """
        Get cached models.

        Returns:
            List of cached AvailableModel objects (may be empty)
        """
        return list(self._cached_models)

    def set_models(self, models: List['AvailableModel']) -> None:
        """
        Update the cache with new models.

        Args:
            models: List of AvailableModel objects from discovery
        """
        self._cached_models = list(models) if models else []
        self._last_scan_time = time.time()
        self.logger.info(f"LocalModelCache updated with {len(self._cached_models)} model(s)")

        # Notify registered callbacks
        self._notify_callbacks()

    def clear_cache(self) -> None:
        """Clear the cached models and reset scan time."""
        self._cached_models = []
        self._last_scan_time = None
        self.logger.info("LocalModelCache cleared")

    def is_empty(self) -> bool:
        """Check if cache has no models."""
        return len(self._cached_models) == 0

    def is_stale(self, max_age_seconds: int = None) -> bool:
        """
        Check if cache is stale (too old or never populated).

        Args:
            max_age_seconds: Maximum age in seconds before cache is stale.
                           Defaults to DEFAULT_STALE_TIMEOUT (60 seconds).

        Returns:
            True if cache should be refreshed, False if still valid
        """
        if max_age_seconds is None:
            max_age_seconds = self.DEFAULT_STALE_TIMEOUT

        if self._last_scan_time is None:
            return True

        age = time.time() - self._last_scan_time
        return age > max_age_seconds

    def is_scan_in_progress(self) -> bool:
        """Check if a scan is currently running."""
        return self._scan_in_progress

    def register_update_callback(self, callback: Callable[[List['AvailableModel']], None]) -> None:
        """
        Register a callback to be notified when cache is updated.

        Args:
            callback: Function that receives list of models when cache updates
        """
        if callback not in self._update_callbacks:
            self._update_callbacks.append(callback)

    def unregister_update_callback(self, callback: Callable[[List['AvailableModel']], None]) -> None:
        """
        Remove a previously registered callback.

        Args:
            callback: The callback function to remove
        """
        if callback in self._update_callbacks:
            self._update_callbacks.remove(callback)

    def _notify_callbacks(self) -> None:
        """Notify all registered callbacks of cache update."""
        for callback in self._update_callbacks:
            try:
                callback(self._cached_models)
            except Exception as e:
                self.logger.warning(f"Error in cache update callback: {e}")

    def scan_async(
        self,
        connector: 'PowerBIConnector',
        force: bool = False,
        on_complete: Optional[Callable[[List['AvailableModel']], None]] = None,
        on_error: Optional[Callable[[str], None]] = None
    ) -> bool:
        """
        Trigger an asynchronous scan for local models.

        Args:
            connector: PowerBIConnector instance for discovery
            force: If True, scan even if cache is fresh
            on_complete: Callback when scan completes (receives model list)
            on_error: Callback when scan fails (receives error message)

        Returns:
            True if scan was started, False if already in progress or cache is fresh
        """
        # Check if scan is needed
        if not force and not self.is_empty() and not self.is_stale():
            self.logger.debug("Cache is fresh, skipping scan")
            if on_complete:
                on_complete(self._cached_models)
            return False

        # Check if scan already in progress
        with self._scan_lock:
            if self._scan_in_progress:
                self.logger.debug("Scan already in progress")
                return False
            self._scan_in_progress = True

        def scan_thread():
            try:
                self.logger.info("Starting local model scan...")

                # Try fast discovery first
                models = connector.discover_local_models_fast()

                if not models:
                    # Fall back to smart port detection
                    models = connector.discover_local_models(quick_scan=True)

                # Update cache
                self.set_models(models)

                # Call completion callback
                if on_complete:
                    on_complete(models)

            except Exception as e:
                self.logger.error(f"Scan failed: {e}")
                if on_error:
                    on_error(str(e))
            finally:
                with self._scan_lock:
                    self._scan_in_progress = False

        thread = threading.Thread(target=scan_thread, daemon=True)
        thread.start()
        return True


def get_local_model_cache() -> LocalModelCache:
    """
    Convenience function to get the LocalModelCache singleton.

    Returns:
        The singleton LocalModelCache instance
    """
    return LocalModelCache.get_instance()
