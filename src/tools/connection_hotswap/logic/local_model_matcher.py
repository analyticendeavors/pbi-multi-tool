"""
Local Model Matcher - Auto-match local Power BI Desktop models by name
Built by Reid Havens of Analytic Endeavors

Discovers open Power BI Desktop instances and suggests matches based on name similarity.
"""

import logging
from difflib import SequenceMatcher
from typing import List, Optional, Callable

from core.local_model_cache import get_local_model_cache
from tools.field_parameters.models import AvailableModel
from tools.connection_hotswap.models import (
    DataSourceConnection,
    ConnectionMapping,
    SwapTarget,
    SwapStatus,
)


class LocalModelMatcher:
    """
    Auto-match local Power BI Desktop models by name.

    Uses the PowerBIConnector's discovery methods to find open models
    and matches them to cloud connections based on name similarity.
    """

    # Minimum similarity ratio for auto-matching (0.0 to 1.0)
    MATCH_THRESHOLD = 0.6

    def __init__(self, connector: 'PowerBIConnector'):
        """
        Initialize the local model matcher.

        Args:
            connector: PowerBIConnector instance for model discovery
        """
        self.connector = connector
        self.logger = logging.getLogger(__name__)
        self._cached_models: List[AvailableModel] = []

    def discover_local_models(
        self,
        progress_callback: Optional[Callable[[str], None]] = None,
        force_refresh: bool = False
    ) -> List[AvailableModel]:
        """
        Discover all open Power BI Desktop models.

        Reuses PowerBIConnector.discover_local_models() implementation.
        Uses shared LocalModelCache for cross-tool cache sharing.

        Args:
            progress_callback: Optional callback for progress updates
            force_refresh: If True, ignore cache and rescan

        Returns:
            List of discovered AvailableModel objects
        """
        # Check instance cache first
        if self._cached_models and not force_refresh:
            return self._cached_models

        # Check shared cache as fallback (may have been populated by Field Parameters)
        shared_cache = get_local_model_cache()
        if not force_refresh and not shared_cache.is_empty() and not shared_cache.is_stale():
            models = shared_cache.get_models()
            if models:
                self.logger.info(f"Using {len(models)} model(s) from shared cache")
                self._cached_models = models
                return models

        try:
            # Try fast discovery first (port file reading)
            models = self.connector.discover_local_models_fast()

            if not models:
                # Fall back to smart port detection
                if progress_callback:
                    progress_callback("Scanning for local models...")
                models = self.connector.discover_local_models(
                    quick_scan=True,
                    progress_callback=progress_callback
                )

            # Update both instance cache and shared cache
            self._cached_models = models
            shared_cache.set_models(models)
            self.logger.info(f"Discovered {len(models)} local model(s)")
            return models

        except Exception as e:
            self.logger.error(f"Error discovering local models: {e}")
            return []

    def find_matching_model(
        self,
        target_name: str,
        local_models: Optional[List[AvailableModel]] = None
    ) -> Optional[AvailableModel]:
        """
        Find a local model matching the given name.

        Matching strategies (in priority order):
        1. Exact name match (case-insensitive)
        2. Name contains match
        3. Fuzzy match using SequenceMatcher

        Args:
            target_name: Name to match against (e.g., dataset name from cloud)
            local_models: List of local models to search (uses cache if None)

        Returns:
            Best matching AvailableModel, or None if no match above threshold
        """
        if local_models is None:
            local_models = self._cached_models or self.discover_local_models()

        if not local_models or not target_name:
            return None

        target_lower = target_name.lower().strip()
        best_match: Optional[AvailableModel] = None
        best_score = 0.0

        for model in local_models:
            model_name = model.display_name.lower().strip()

            # Strategy 1: Exact match
            if model_name == target_lower:
                return model

            # Strategy 2: Contains match
            if target_lower in model_name or model_name in target_lower:
                score = 0.9  # High score for contains match
                if score > best_score:
                    best_score = score
                    best_match = model
                continue

            # Strategy 3: Fuzzy match
            score = SequenceMatcher(None, target_lower, model_name).ratio()
            if score > best_score:
                best_score = score
                best_match = model

        # Only return if above threshold
        if best_score >= self.MATCH_THRESHOLD:
            self.logger.debug(f"Matched '{target_name}' to '{best_match.display_name}' (score: {best_score:.2f})")
            return best_match

        self.logger.debug(f"No match found for '{target_name}' (best score: {best_score:.2f})")
        return None

    def suggest_matches(
        self,
        connections: List[DataSourceConnection],
        local_models: Optional[List[AvailableModel]] = None
    ) -> List[ConnectionMapping]:
        """
        Auto-suggest mappings for all connections.

        For each cloud connection:
        1. Extract dataset name
        2. Find matching local model
        3. Create ConnectionMapping with auto_matched=True if found

        Args:
            connections: List of connections to match
            local_models: Optional list of local models (discovers if None)

        Returns:
            List of ConnectionMapping objects with suggested targets
        """
        if local_models is None:
            local_models = self.discover_local_models()

        mappings: List[ConnectionMapping] = []

        for connection in connections:
            mapping = ConnectionMapping(source=connection)

            # Only try to match swappable connections
            if not connection.is_swappable:
                mapping.status = SwapStatus.PENDING
                mappings.append(mapping)
                continue

            # Get name to match against
            match_name = connection.dataset_name or connection.database or connection.name

            # Find matching local model
            match = self.find_matching_model(match_name, local_models)

            if match:
                # Create swap target from matched model
                target = SwapTarget(
                    target_type="local",
                    server=match.server,
                    database=match.database_name,
                    display_name=match.display_name,
                )
                mapping.target = target
                mapping.auto_matched = True
                # Set to READY directly - no confirmation needed for auto-match
                mapping.status = SwapStatus.READY
                self.logger.info(f"Auto-matched '{match_name}' → '{match.display_name}'")
            else:
                mapping.status = SwapStatus.PENDING

            mappings.append(mapping)

        return mappings

    def model_to_swap_target(self, model: AvailableModel) -> SwapTarget:
        """
        Convert an AvailableModel to a SwapTarget.

        Args:
            model: AvailableModel from discovery

        Returns:
            SwapTarget configured for local connection
        """
        return SwapTarget(
            target_type="local",
            server=model.server,
            database=model.database_name,
            display_name=model.display_name,
        )

    def validate_local_model_compatibility(
        self,
        source_connection: DataSourceConnection,
        local_model: AvailableModel
    ) -> tuple[bool, str]:
        """
        Validate that a local model is compatible with the source.

        Checks:
        - Model is accessible (can connect)
        - Basic structure compatibility (warning only)

        Args:
            source_connection: Original cloud connection
            local_model: Local model to validate

        Returns:
            Tuple of (is_compatible, message)
        """
        try:
            # Check if we can access the local model
            # This is a basic check - full validation would require connecting
            import socket

            if ':' in local_model.server:
                host, port_str = local_model.server.rsplit(':', 1)
                port = int(port_str)

                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2)
                result = sock.connect_ex((host, port))
                sock.close()

                if result != 0:
                    return False, f"Cannot connect to local model at {local_model.server}"

            return True, f"Local model '{local_model.display_name}' is accessible"

        except Exception as e:
            return False, f"Validation error: {e}"

    def clear_cache(self):
        """Clear the cached local models list."""
        self._cached_models = []

    def set_cache(self, models: List[AvailableModel]):
        """Pre-populate the cache with already-discovered models.

        This avoids redundant port scanning when models were already
        discovered (e.g., from the dropdown model selection).

        Args:
            models: List of AvailableModel objects to cache
        """
        self._cached_models = list(models) if models else []
