"""
Connection Cache Manager - Persistent storage for cloud connection schemas
Built by Reid Havens of Analytic Endeavors

Provides disk-persistent caching of original cloud connection schemas to ensure
proper restoration when swapping from local back to cloud connections, even
across application restarts.

Key Features:
- Stores complete original cloud connection format (PbiServiceModelId, etc.)
- Schema fingerprinting for change detection
- Survives application restarts (unlike in-memory caching)
- Supports both PBIX and PBIP file formats
"""

import datetime
import hashlib
import json
import logging
import os
import sys
from pathlib import Path
from typing import Optional, Dict, Any


# Schema fingerprint fields - these are the critical fields that identify a connection
SCHEMA_FINGERPRINT_FIELDS = [
    "ConnectionType",
    "PbiServiceModelId",
    "PbiModelVirtualServerName",
    "PbiModelDatabaseName",
]

# PBIP equivalent field names (lowercase)
PBIP_FINGERPRINT_FIELDS = [
    "connectionType",
    "pbiServiceModelId",
    "pbiModelVirtualServerName",
    "pbiModelDatabaseName",
]


class ConnectionCacheManager:
    """
    Manages persistent disk cache for cloud connection schemas.

    Cache is stored in user's AppData folder and survives application restarts.
    This ensures original cloud connection formats (with PbiServiceModelId,
    PbiModelVirtualServerName, etc.) can be restored properly.
    """

    CACHE_FILENAME = "connection_cache.json"
    CACHE_VERSION = "1.0"

    def __init__(self):
        """Initialize the connection cache manager."""
        self.logger = logging.getLogger(__name__)
        self._cache_dir = self._get_cache_dir()
        self._cache_file = self._cache_dir / self.CACHE_FILENAME

        # In-memory cache for fast access
        self._memory_cache: Dict[str, dict] = {}

        # Ensure cache directory exists
        self._cache_dir.mkdir(parents=True, exist_ok=True)

        # Load existing cache from disk
        self._load_from_disk()

    def _get_cache_dir(self) -> Path:
        """Get the cache directory path (same location as presets)."""
        base = Path(os.environ.get('APPDATA', ''))
        return base / 'Analytic Endeavors' / 'PBI Report Merger' / 'hotswap_presets'

    def _normalize_path(self, file_path: str) -> str:
        """Normalize file path for consistent cache lookup."""
        if not file_path:
            return ""
        # Normalize to lowercase with forward slashes
        return os.path.normpath(file_path).lower().replace('\\', '/')

    def _load_from_disk(self) -> None:
        """Load cache from disk into memory."""
        try:
            if self._cache_file.exists():
                with open(self._cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # Validate version
                if data.get('version') == self.CACHE_VERSION:
                    self._memory_cache = data.get('connections', {})
                    self.logger.debug(f"Loaded {len(self._memory_cache)} cached connections from disk")
                else:
                    # Version mismatch - start fresh but don't delete old file
                    self.logger.info(f"Cache version mismatch (found {data.get('version')}, expected {self.CACHE_VERSION})")
                    self._memory_cache = {}
        except Exception as e:
            self.logger.warning(f"Error loading connection cache from disk: {e}")
            self._memory_cache = {}

    def _save_to_disk(self) -> bool:
        """Persist cache to disk with atomic write."""
        try:
            cache_data = {
                "version": self.CACHE_VERSION,
                "connections": self._memory_cache
            }

            # Atomic write: write to temp file, then rename
            temp_path = self._cache_file.with_suffix('.tmp')
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, indent=2)

            # Rename temp to actual (atomic on most filesystems)
            temp_path.replace(self._cache_file)
            return True

        except Exception as e:
            self.logger.error(f"Error saving connection cache to disk: {e}")
            # Clean up temp file if it exists
            temp_path = self._cache_file.with_suffix('.tmp')
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except:
                    pass
            return False

    def save_connection(self, file_path: str, connection_info: dict, file_type: str = "pbix") -> bool:
        """
        Save a cloud connection to the persistent cache.

        Args:
            file_path: Path to the .pbix or .pbip file
            connection_info: Connection dict from read_current_connection()
            file_type: Either "pbix" or "pbip"

        Returns:
            True if save was successful
        """
        if not file_path or not connection_info:
            return False

        normalized_path = self._normalize_path(file_path)

        # Compute schema fingerprint
        fingerprint = self.compute_schema_fingerprint(connection_info, file_type)

        cache_entry = {
            "cached_at": datetime.datetime.utcnow().isoformat() + "Z",
            "file_type": file_type,
            "schema_fingerprint": fingerprint,
            "connection_data": connection_info.copy()
        }

        self._memory_cache[normalized_path] = cache_entry

        # Persist to disk
        success = self._save_to_disk()
        if success:
            self.logger.info(f"Cached cloud connection for: {file_path}")
        return success

    def load_connection(self, file_path: str) -> Optional[dict]:
        """
        Load a cached connection from the cache.

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            Cached connection_data dict, or None if not found
        """
        if not file_path:
            return None

        normalized_path = self._normalize_path(file_path)
        cache_entry = self._memory_cache.get(normalized_path)

        if cache_entry:
            return cache_entry.get('connection_data')
        return None

    def get_cache_entry(self, file_path: str) -> Optional[dict]:
        """
        Get the full cache entry including metadata.

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            Full cache entry dict with cached_at, file_type, fingerprint, connection_data
        """
        if not file_path:
            return None

        normalized_path = self._normalize_path(file_path)
        return self._memory_cache.get(normalized_path)

    def has_cached_connection(self, file_path: str) -> bool:
        """
        Check if we have a cached connection for a file.

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            True if cached connection exists
        """
        return self.load_connection(file_path) is not None

    def clear_connection(self, file_path: str) -> bool:
        """
        Remove a cached connection.

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            True if removed successfully
        """
        if not file_path:
            return False

        normalized_path = self._normalize_path(file_path)

        if normalized_path in self._memory_cache:
            del self._memory_cache[normalized_path]
            self._save_to_disk()
            self.logger.info(f"Cleared cached connection for: {file_path}")
            return True
        return False

    def get_cached_fingerprint(self, file_path: str) -> Optional[str]:
        """
        Get the schema fingerprint for a cached connection.

        Args:
            file_path: Path to the .pbix or .pbip file

        Returns:
            Schema fingerprint string, or None if not cached
        """
        cache_entry = self.get_cache_entry(file_path)
        if cache_entry:
            return cache_entry.get('schema_fingerprint')
        return None

    @staticmethod
    def compute_schema_fingerprint(connection_info: dict, file_type: str = "pbix") -> str:
        """
        Compute a fingerprint for schema comparison.

        The fingerprint captures the key fields that define a cloud connection's
        identity and format. Changes to these fields indicate the connection
        schema has changed.

        Args:
            connection_info: Connection dict (either raw connection or parsed info)
            file_type: Either "pbix" or "pbip" (affects field name lookup)

        Returns:
            16-character hex fingerprint string
        """
        # Try to get the original connection data if available
        original_conn = connection_info.get('_original_pbix_connection') or \
                       connection_info.get('_original_pbip_definition', {}).get('datasetReference', {}).get('byConnection', {})

        if not original_conn:
            original_conn = connection_info

        # Determine which field names to use
        if file_type == "pbip":
            fields = PBIP_FINGERPRINT_FIELDS
        else:
            fields = SCHEMA_FINGERPRINT_FIELDS

        # Build fingerprint data
        fingerprint_data = {}
        for field in fields:
            value = original_conn.get(field)
            # Normalize integers vs strings for PbiServiceModelId
            if field.lower() == 'pbiservicemodelid' and value is not None:
                try:
                    value = int(value)
                except (ValueError, TypeError):
                    pass
            fingerprint_data[field] = value

        # Also check uppercase/lowercase variants
        for field in SCHEMA_FINGERPRINT_FIELDS:
            if field not in fingerprint_data:
                value = original_conn.get(field)
                if value is not None:
                    fingerprint_data[field] = value

        # Create stable JSON for hashing
        stable_json = json.dumps(fingerprint_data, sort_keys=True, default=str)
        return hashlib.md5(stable_json.encode()).hexdigest()[:16]

    @staticmethod
    def compare_fingerprints(fingerprint1: Optional[str], fingerprint2: Optional[str]) -> bool:
        """
        Compare two schema fingerprints.

        Args:
            fingerprint1: First fingerprint
            fingerprint2: Second fingerprint

        Returns:
            True if fingerprints match (or both are None)
        """
        if fingerprint1 is None and fingerprint2 is None:
            return True
        if fingerprint1 is None or fingerprint2 is None:
            return False
        return fingerprint1 == fingerprint2

    def get_schema_diff(self, cached_connection: dict, current_connection: dict,
                        file_type: str = "pbix") -> Dict[str, Any]:
        """
        Get a detailed diff between cached and current connection schemas.

        Args:
            cached_connection: Cached connection data
            current_connection: Current connection data
            file_type: Either "pbix" or "pbip"

        Returns:
            Dict with 'matches' (bool) and 'differences' (list of field diffs)
        """
        # Get original connections
        cached_orig = cached_connection.get('_original_pbix_connection') or \
                     cached_connection.get('_original_pbip_definition', {}).get('datasetReference', {}).get('byConnection', {}) or \
                     cached_connection

        current_orig = current_connection.get('_original_pbix_connection') or \
                      current_connection.get('_original_pbip_definition', {}).get('datasetReference', {}).get('byConnection', {}) or \
                      current_connection

        fields = SCHEMA_FINGERPRINT_FIELDS if file_type == "pbix" else PBIP_FINGERPRINT_FIELDS

        differences = []
        for field in fields:
            cached_val = cached_orig.get(field)
            current_val = current_orig.get(field)

            # Normalize for comparison
            if field.lower() == 'pbiservicemodelid':
                try:
                    if cached_val is not None:
                        cached_val = int(cached_val)
                    if current_val is not None:
                        current_val = int(current_val)
                except (ValueError, TypeError):
                    pass

            if cached_val != current_val:
                differences.append({
                    'field': field,
                    'cached': cached_val,
                    'current': current_val
                })

        return {
            'matches': len(differences) == 0,
            'differences': differences
        }

    def clear_all(self) -> bool:
        """
        Clear all cached connections.

        Returns:
            True if cleared successfully
        """
        self._memory_cache = {}
        return self._save_to_disk()


# Global singleton instance for easy access
_cache_manager: Optional[ConnectionCacheManager] = None


def get_cache_manager() -> ConnectionCacheManager:
    """Get the global ConnectionCacheManager singleton."""
    global _cache_manager
    if _cache_manager is None:
        _cache_manager = ConnectionCacheManager()
    return _cache_manager
