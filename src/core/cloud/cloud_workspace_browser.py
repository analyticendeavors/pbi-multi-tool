"""
Cloud Workspace Browser - Fabric API workspace and dataset enumeration
Built by Reid Havens of Analytic Endeavors

Browses Power BI Service workspaces and datasets via the Fabric/Power BI REST API.
"""

import json
import logging
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
from urllib.parse import quote

try:
    import requests
except ImportError:
    requests = None

try:
    import msal
except ImportError:
    msal = None

from core.cloud.models import WorkspaceInfo, DatasetInfo, SwapTarget, CloudConnectionType

# Module-level shared MSAL app and token cache for authentication persistence
_shared_msal_app: Any = None
_shared_token_cache: Any = None
_token_cache_path: Optional[Path] = None


def _get_token_cache_path() -> Path:
    """Get the path for the persistent token cache file."""
    if getattr(sys, 'frozen', False):
        config_dir = Path(os.environ.get('APPDATA', '.')) / 'Analytic Endeavors' / 'PBI Report Merger'
    else:
        # From core/cloud/ -> core/ -> src/ (3 parents)
        config_dir = Path(__file__).parent.parent.parent / 'config'
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir / 'msal_token_cache.bin'


def _get_shared_msal_app() -> Optional[Any]:
    """Get or create the shared MSAL PublicClientApplication with persistent token cache."""
    global _shared_msal_app, _shared_token_cache, _token_cache_path

    if not msal:
        return None

    if _shared_msal_app is not None:
        return _shared_msal_app

    # Create persistent token cache
    _token_cache_path = _get_token_cache_path()
    _shared_token_cache = msal.SerializableTokenCache()

    # Load existing cache from disk
    if _token_cache_path.exists():
        try:
            with open(_token_cache_path, 'r', encoding='utf-8') as f:
                _shared_token_cache.deserialize(f.read())
        except Exception:
            pass  # Start fresh if cache is corrupted

    # Create MSAL app with the persistent cache
    _shared_msal_app = msal.PublicClientApplication(
        CloudWorkspaceBrowser.AZURE_CLIENT_ID,
        authority=CloudWorkspaceBrowser.AZURE_AUTHORITY,
        token_cache=_shared_token_cache
    )

    return _shared_msal_app


def _save_token_cache():
    """Save the token cache to disk if it has changed."""
    global _shared_token_cache, _token_cache_path

    if _shared_token_cache is None or _token_cache_path is None:
        return

    if _shared_token_cache.has_state_changed:
        try:
            with open(_token_cache_path, 'w', encoding='utf-8') as f:
                f.write(_shared_token_cache.serialize())
        except Exception:
            pass  # Ignore cache save errors


class CloudWorkspaceBrowser:
    """
    Browse Power BI Service workspaces and datasets via Fabric API.

    Uses MSAL for OAuth authentication and the Power BI REST API for
    workspace and dataset enumeration.
    """

    # Azure AD App Registration for Power BI (public client ID)
    AZURE_CLIENT_ID = "a672d62c-fc7b-4e81-a576-e60dc46e951d"
    AZURE_AUTHORITY = "https://login.microsoftonline.com/common"
    FABRIC_API_SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]

    # API endpoints
    API_BASE = "https://api.powerbi.com/v1.0/myorg"

    def __init__(self):
        """Initialize the cloud workspace browser."""
        self.logger = logging.getLogger(__name__)
        self._access_token: Optional[str] = None
        self._msal_app: Optional[Any] = None

        # Cached data
        self._workspaces: List[WorkspaceInfo] = []
        self._datasets_cache: Dict[str, List[DatasetInfo]] = {}

        # User preferences (persisted)
        self._favorites: List[str] = []  # Workspace IDs
        self._model_favorites: List[str] = []  # Dataset/Model IDs
        self._recent: List[Tuple[str, str]] = []  # (workspace_id, dataset_id) pairs

        # Session-level cache tracking (not persisted)
        self._session_workspaces_loaded: bool = False
        self._session_preload_started: bool = False

        # Load persisted preferences
        self._load_preferences()

    def _get_config_path(self) -> Path:
        """Get the path for storing configuration/preferences."""
        if getattr(sys, 'frozen', False):
            # Running as EXE
            config_dir = Path(os.environ.get('APPDATA', '.')) / 'Analytic Endeavors' / 'PBI Report Merger'
        else:
            # Running from source - from core/cloud/ -> core/ -> src/ (3 parents)
            config_dir = Path(__file__).parent.parent.parent / 'config'

        config_dir.mkdir(parents=True, exist_ok=True)
        return config_dir / 'cloud_browser_prefs.json'

    def _load_preferences(self):
        """Load persisted preferences (favorites, recent)."""
        try:
            config_path = self._get_config_path()
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self._favorites = data.get('favorites', [])
                    self._model_favorites = data.get('model_favorites', [])
                    self._recent = [tuple(r) for r in data.get('recent', [])]
        except Exception as e:
            self.logger.debug(f"Could not load preferences: {e}")

    def _save_preferences(self):
        """Save preferences to disk."""
        try:
            config_path = self._get_config_path()
            data = {
                'favorites': self._favorites,
                'model_favorites': self._model_favorites,
                'recent': [list(r) for r in self._recent],
            }
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self.logger.debug(f"Could not save preferences: {e}")

    def authenticate(self) -> Tuple[bool, str]:
        """
        Authenticate via MSAL interactive flow.

        Uses cached tokens when available to avoid re-authentication within a session.
        Token cache is shared across all instances and persisted to disk.

        Returns:
            Tuple of (success, message)
        """
        if not msal:
            return False, "MSAL library not installed. Run: pip install msal"

        try:
            # Use shared MSAL app with persistent token cache
            self._msal_app = _get_shared_msal_app()
            if self._msal_app is None:
                return False, "Failed to initialize MSAL"

            # Try to get token from cache first (silent auth)
            accounts = self._msal_app.get_accounts()
            if accounts:
                result = self._msal_app.acquire_token_silent(
                    self.FABRIC_API_SCOPE,
                    account=accounts[0]
                )
                if result and 'access_token' in result:
                    self._access_token = result['access_token']
                    _save_token_cache()  # Persist cache after successful silent auth
                    return True, "Authenticated using cached credentials"

            # Interactive login required
            result = self._msal_app.acquire_token_interactive(
                scopes=self.FABRIC_API_SCOPE,
            )

            if 'access_token' in result:
                self._access_token = result['access_token']
                _save_token_cache()  # Persist cache after successful login
                return True, "Authentication successful"
            else:
                error = result.get('error_description', result.get('error', 'Unknown error'))
                return False, f"Authentication failed: {error}"

        except Exception as e:
            self.logger.error(f"Authentication error: {e}")
            return False, f"Authentication error: {e}"

    def is_authenticated(self) -> bool:
        """Check if we have a valid access token."""
        return self._access_token is not None

    def try_silent_auth(self) -> bool:
        """
        Attempt silent authentication using cached MSAL tokens.

        Does not prompt for user interaction. Returns True if successful.
        Use this for background auth attempts (e.g., on tab switch).
        """
        if not msal:
            return False

        try:
            msal_app = _get_shared_msal_app()
            if msal_app is None:
                return False

            accounts = msal_app.get_accounts()
            if accounts:
                result = msal_app.acquire_token_silent(
                    self.FABRIC_API_SCOPE,
                    account=accounts[0]
                )
                if result and 'access_token' in result:
                    self._access_token = result['access_token']
                    self._msal_app = msal_app
                    _save_token_cache()
                    return True
            return False
        except Exception as e:
            self.logger.debug(f"Silent auth failed: {e}")
            return False

    def is_session_cached(self) -> bool:
        """Check if workspace data has been loaded this session."""
        return self._session_workspaces_loaded and len(self._workspaces) > 0

    def mark_session_loaded(self):
        """Mark that workspace data has been loaded this session."""
        self._session_workspaces_loaded = True

    def get_account_email(self) -> Optional[str]:
        """
        Get the email/username of the currently authenticated account.

        Returns:
            Email address if authenticated, None otherwise
        """
        msal_app = self._msal_app or _get_shared_msal_app()
        if msal_app:
            accounts = msal_app.get_accounts()
            if accounts:
                return accounts[0].get('username', None)
        return None

    def sign_out(self) -> bool:
        """
        Sign out by removing all accounts from MSAL cache.

        Clears the token cache and removes the current access token.

        Returns:
            True if sign out was successful, False otherwise
        """
        global _shared_token_cache, _token_cache_path
        msal_app = self._msal_app or _get_shared_msal_app()
        if msal_app:
            accounts = msal_app.get_accounts()
            for account in accounts:
                msal_app.remove_account(account)
            self._access_token = None
            _save_token_cache()  # Persist cleared cache
            return True
        return False

    def _try_refresh_token(self) -> bool:
        """
        Attempt to silently refresh the access token using MSAL cache.

        Returns:
            True if token was refreshed successfully, False otherwise
        """
        if not msal:
            return False

        # Use shared MSAL app if instance doesn't have one
        msal_app = self._msal_app or _get_shared_msal_app()
        if not msal_app:
            return False

        try:
            accounts = msal_app.get_accounts()
            if accounts:
                result = msal_app.acquire_token_silent(
                    self.FABRIC_API_SCOPE,
                    account=accounts[0]
                )
                if result and 'access_token' in result:
                    self._access_token = result['access_token']
                    self._msal_app = msal_app  # Store reference
                    _save_token_cache()  # Persist cache after refresh
                    return True
        except Exception as e:
            self.logger.debug(f"Silent token refresh failed: {e}")

        return False

    def _api_request(
        self,
        endpoint: str,
        method: str = "GET",
        params: Optional[Dict] = None,
        _retry_after_refresh: bool = True
    ) -> Tuple[Optional[Dict], Optional[str]]:
        """
        Make an authenticated API request.

        Args:
            endpoint: API endpoint (relative to API_BASE)
            method: HTTP method
            params: Query parameters
            _retry_after_refresh: Internal flag to prevent infinite retry loops

        Returns:
            Tuple of (response_json, error_message)
        """
        if not requests:
            return None, "requests library not installed"

        if not self._access_token:
            return None, "Not authenticated"

        url = f"{self.API_BASE}{endpoint}"
        headers = {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }

        try:
            response = requests.request(method, url, headers=headers, params=params, timeout=30)

            if response.status_code == 401:
                # Token expired, try silent refresh before failing
                if _retry_after_refresh and self._try_refresh_token():
                    # Token refreshed, retry the request once
                    return self._api_request(endpoint, method, params, _retry_after_refresh=False)
                # Silent refresh failed, clear token and require re-auth
                self._access_token = None
                return None, "Authentication expired. Please re-authenticate."

            if response.status_code >= 400:
                return None, f"API error {response.status_code}: {response.text}"

            try:
                return response.json(), None
            except (json.JSONDecodeError, ValueError) as e:
                return None, f"Invalid API response: {e}"

        except requests.RequestException as e:
            return None, f"Request failed: {e}"

    def _get_capacity_types(self) -> Dict[str, str]:
        """
        Fetch capacity info and return a map of capacityId -> capacity type.

        Queries the /capacities endpoint to distinguish Premium from PPU.
        SKU prefixes: P/EM = Premium, PP = PPU, F = Fabric

        Returns:
            Dict mapping capacityId to 'Premium', 'PPU', or 'Fabric'
        """
        capacity_map = {}
        try:
            data, error = self._api_request("/capacities")
            if error or not data:
                self.logger.debug(f"Could not fetch capacities: {error}")
                return capacity_map

            for cap in data.get('value', []):
                cap_id = cap.get('id')
                sku = cap.get('sku', '').upper()
                if not cap_id:
                    continue

                # Determine type from SKU prefix
                if sku.startswith('PP'):
                    # PP1, PP2, etc. = Premium Per User
                    capacity_map[cap_id] = 'PPU'
                elif sku.startswith('F'):
                    # F2, F4, F8, etc. = Fabric
                    capacity_map[cap_id] = 'Premium'  # Fabric shows as Premium
                else:
                    # P1, P2, EM1, EM2, etc. = Premium/Embedded
                    capacity_map[cap_id] = 'Premium'

            self.logger.debug(f"Loaded {len(capacity_map)} capacity types")
        except Exception as e:
            self.logger.debug(f"Error fetching capacities: {e}")

        return capacity_map

    def get_workspaces(
        self,
        filter_type: str = "all",
        force_refresh: bool = False
    ) -> Tuple[List[WorkspaceInfo], Optional[str]]:
        """
        Get accessible workspaces.

        Args:
            filter_type: "all", "favorites", or "recent"
            force_refresh: If True, bypass cache and fetch fresh data from API

        Returns:
            Tuple of (workspaces list, error message)
        """
        if filter_type == "favorites":
            return self._get_favorite_workspaces()
        elif filter_type == "recent":
            return self._get_recent_workspaces()

        # Return cached workspaces if available and not forcing refresh
        if self._workspaces and not force_refresh:
            return self._workspaces, None

        # Fetch all workspaces from API
        data, error = self._api_request("/groups")

        if error:
            return [], error

        # Fetch capacity info to distinguish Premium from PPU
        capacity_types = self._get_capacity_types()

        workspaces = []
        for ws in data.get('value', []):
            is_dedicated = ws.get('isOnDedicatedCapacity', False)
            capacity_id = ws.get('capacityId')

            # Determine capacity type based on capacity SKU
            capacity_type = None
            if is_dedicated and capacity_id:
                capacity_type = capacity_types.get(capacity_id, 'Premium')
            elif is_dedicated:
                # Has dedicated capacity but no capacityId - assume Premium
                capacity_type = 'Premium'
            # If not dedicated, capacity_type stays None (Pro workspace)
            # PPU will be detected via XMLA access later if this fails

            workspace = WorkspaceInfo(
                id=ws.get('id', ''),
                name=ws.get('name', ''),
                type=ws.get('type', 'Workspace'),
                is_on_dedicated_capacity=is_dedicated,
                is_favorite=ws.get('id', '') in self._favorites,
                capacity_type=capacity_type,
            )
            workspaces.append(workspace)

        # Sort by name
        workspaces.sort(key=lambda w: w.name.lower())

        self._workspaces = workspaces
        self.mark_session_loaded()  # Mark session data as loaded
        return workspaces, None

    def _get_favorite_workspaces(self) -> Tuple[List[WorkspaceInfo], Optional[str]]:
        """Get workspaces that have any favorited models."""
        if not self._workspaces:
            self.get_workspaces("all")

        # Return workspaces that have at least one favorited model
        favorites = [w for w in self._workspaces
                     if self.get_workspace_favorite_status(w.id) != 'none']
        return favorites, None

    def _get_recent_workspaces(self) -> Tuple[List[WorkspaceInfo], Optional[str]]:
        """Get workspaces from recent history."""
        if not self._workspaces:
            self.get_workspaces("all")

        recent_ws_ids = set(ws_id for ws_id, _ in self._recent)
        recent = [w for w in self._workspaces if w.id in recent_ws_ids]
        return recent, None

    def _update_workspace_capacity_type_if_ppu(self, workspace_name: str) -> None:
        """
        Update workspace capacity_type to 'PPU' if XMLA access worked
        but the workspace is not on dedicated capacity.

        This is called when XMLA access succeeds (e.g., perspective fetch works).
        If the workspace is not marked as dedicated capacity, it must be PPU.
        """
        for ws in self._workspaces:
            if ws.name == workspace_name:
                if not ws.is_on_dedicated_capacity and ws.capacity_type is None:
                    ws.capacity_type = 'PPU'
                    self.logger.debug(f"Workspace '{workspace_name}' detected as PPU capacity")
                break

    def get_workspace_datasets(
        self,
        workspace_id: str
    ) -> Tuple[List[DatasetInfo], Optional[str]]:
        """
        Get datasets in a workspace.

        Args:
            workspace_id: Workspace GUID

        Returns:
            Tuple of (datasets list, error message)
        """
        # Check cache
        if workspace_id in self._datasets_cache:
            return self._datasets_cache[workspace_id], None

        data, error = self._api_request(f"/groups/{workspace_id}/datasets")

        if error:
            return [], error

        # Find workspace name
        workspace_name = ""
        for ws in self._workspaces:
            if ws.id == workspace_id:
                workspace_name = ws.name
                break

        datasets = []
        for ds in data.get('value', []):
            dataset = DatasetInfo(
                id=ds.get('id', ''),
                name=ds.get('name', ''),
                workspace_id=workspace_id,
                workspace_name=workspace_name,
                configured_by=ds.get('configuredBy', ''),
                is_refreshable=ds.get('isRefreshable', True),
                is_effective_identity_required=ds.get('isEffectiveIdentityRequired', False),
            )
            datasets.append(dataset)

        # Sort by name
        datasets.sort(key=lambda d: d.name.lower())

        # Cache results
        self._datasets_cache[workspace_id] = datasets

        return datasets, None

    def search_datasets(
        self,
        query: str,
        max_results: int = 50
    ) -> Tuple[List[DatasetInfo], Optional[str]]:
        """
        Search across all accessible datasets by name.

        Args:
            query: Search query string
            max_results: Maximum results to return

        Returns:
            Tuple of (matching datasets, error message)
        """
        if not query:
            return [], None

        query_lower = query.lower()
        results: List[DatasetInfo] = []

        # Ensure we have workspaces
        if not self._workspaces:
            _, error = self.get_workspaces("all")
            if error:
                return [], error

        # Search through cached datasets first
        for ws_id, datasets in self._datasets_cache.items():
            for ds in datasets:
                if query_lower in ds.name.lower():
                    results.append(ds)
                    if len(results) >= max_results:
                        return results, None

        # If we don't have enough results, fetch more workspaces
        for ws in self._workspaces:
            if ws.id in self._datasets_cache:
                continue  # Already searched

            datasets, _ = self.get_workspace_datasets(ws.id)
            for ds in datasets:
                if query_lower in ds.name.lower():
                    results.append(ds)
                    if len(results) >= max_results:
                        return results, None

        return results, None

    def toggle_favorite(self, workspace_id: str) -> bool:
        """
        Toggle favorite status for a workspace.

        Args:
            workspace_id: Workspace GUID

        Returns:
            New favorite status (True if now favorite)
        """
        if workspace_id in self._favorites:
            self._favorites.remove(workspace_id)
            is_favorite = False
        else:
            self._favorites.append(workspace_id)
            is_favorite = True

        # Update workspace object
        for ws in self._workspaces:
            if ws.id == workspace_id:
                ws.is_favorite = is_favorite
                break

        self._save_preferences()
        return is_favorite

    def toggle_model_favorite(self, model_id: str) -> bool:
        """
        Toggle favorite status for a semantic model/dataset.

        Args:
            model_id: Dataset/Model GUID

        Returns:
            New favorite status (True if now favorite)
        """
        if model_id in self._model_favorites:
            self._model_favorites.remove(model_id)
            is_favorite = False
        else:
            self._model_favorites.append(model_id)
            is_favorite = True

        self._save_preferences()
        return is_favorite

    def is_model_favorite(self, model_id: str) -> bool:
        """Check if a model/dataset is a favorite."""
        return model_id in self._model_favorites

    def get_workspace_favorite_status(self, workspace_id: str) -> str:
        """
        Get the favorite status of a workspace based on its models.

        Args:
            workspace_id: Workspace GUID

        Returns:
            'all' if all models are favorited,
            'some' if some models are favorited,
            'none' if no models are favorited
        """
        if workspace_id not in self._datasets_cache:
            return 'none'

        datasets = self._datasets_cache[workspace_id]
        if not datasets:
            return 'none'

        favorited_count = sum(1 for ds in datasets if ds.id in self._model_favorites)

        if favorited_count == 0:
            return 'none'
        elif favorited_count == len(datasets):
            return 'all'
        else:
            return 'some'

    def favorite_all_models_in_workspace(self, workspace_id: str) -> int:
        """
        Favorite all models/datasets in a workspace.

        Args:
            workspace_id: Workspace GUID

        Returns:
            Number of models favorited
        """
        if workspace_id not in self._datasets_cache:
            return 0

        count = 0
        for ds in self._datasets_cache[workspace_id]:
            if ds.id not in self._model_favorites:
                self._model_favorites.append(ds.id)
                count += 1

        if count > 0:
            self._save_preferences()

        return count

    def unfavorite_all_models_in_workspace(self, workspace_id: str) -> int:
        """
        Unfavorite all models/datasets in a workspace.

        Args:
            workspace_id: Workspace GUID

        Returns:
            Number of models unfavorited
        """
        if workspace_id not in self._datasets_cache:
            return 0

        count = 0
        for ds in self._datasets_cache[workspace_id]:
            if ds.id in self._model_favorites:
                self._model_favorites.remove(ds.id)
                count += 1

        if count > 0:
            self._save_preferences()

        return count

    def add_recent(self, workspace_id: str, dataset_id: str):
        """
        Add to recent items list.

        Args:
            workspace_id: Workspace GUID
            dataset_id: Dataset GUID
        """
        item = (workspace_id, dataset_id)

        # Remove if already exists
        if item in self._recent:
            self._recent.remove(item)

        # Add to front
        self._recent.insert(0, item)

        # Limit to 20 recent items
        self._recent = self._recent[:20]

        self._save_preferences()

    def get_recent_datasets(self) -> List[DatasetInfo]:
        """
        Get recently used datasets.

        Returns:
            List of DatasetInfo for recent items
        """
        results = []
        for ws_id, ds_id in self._recent:
            if ws_id in self._datasets_cache:
                for ds in self._datasets_cache[ws_id]:
                    if ds.id == ds_id:
                        results.append(ds)
                        break

        return results

    def dataset_to_swap_target(self, dataset: DatasetInfo) -> SwapTarget:
        """
        Convert a DatasetInfo to a SwapTarget.

        Args:
            dataset: DatasetInfo from browsing

        Returns:
            SwapTarget configured for cloud connection
        """
        return dataset.to_swap_target()

    def create_manual_target(
        self,
        xmla_endpoint: str,
        dataset_name: str,
        cloud_connection_type: Optional[CloudConnectionType] = None,
        perspective_name: Optional[str] = None
    ) -> SwapTarget:
        """
        Create a SwapTarget from manual XMLA entry.

        Args:
            xmla_endpoint: Full XMLA endpoint URL or workspace name
            dataset_name: Dataset name
            cloud_connection_type: Type of cloud connector to use
            perspective_name: Optional perspective/cube name for the connection

        Returns:
            SwapTarget for the cloud connection
        """
        # Handle workspace name vs full URL
        if not xmla_endpoint.startswith('powerbi://'):
            # Assume it's just a workspace name
            xmla_endpoint = f"powerbi://api.powerbi.com/v1.0/myorg/{quote(xmla_endpoint)}"

        # Default to PBI_SEMANTIC_MODEL if not specified
        if cloud_connection_type is None:
            cloud_connection_type = CloudConnectionType.PBI_SEMANTIC_MODEL

        # Build display name including perspective if provided
        display_name = f"{dataset_name} (Manual Entry)"
        if perspective_name:
            display_name = f"{dataset_name} [{perspective_name}] (Manual Entry)"

        return SwapTarget(
            target_type="cloud",
            server=xmla_endpoint,
            database=dataset_name,
            display_name=display_name,
            cloud_connection_type=cloud_connection_type,
            perspective_name=perspective_name,
        )

    def clear_cache(self):
        """Clear cached workspaces and datasets."""
        self._workspaces = []
        self._datasets_cache = {}
        self._all_datasets_loaded = False

    def is_fully_cached(self) -> bool:
        """Check if all datasets have been loaded into cache."""
        return getattr(self, '_all_datasets_loaded', False)

    def preload_all_datasets(
        self,
        progress_callback: Optional[callable] = None
    ) -> Tuple[int, Optional[str]]:
        """
        Preload all datasets from all workspaces into cache.

        This enables fast searching without additional API calls.

        Args:
            progress_callback: Optional callback(current, total, workspace_name)
                              for progress updates

        Returns:
            Tuple of (total_datasets_loaded, error_message)
        """
        # Ensure we have workspaces first
        if not self._workspaces:
            _, error = self.get_workspaces("all")
            if error:
                return 0, error

        total_workspaces = len(self._workspaces)
        total_datasets = 0

        for i, ws in enumerate(self._workspaces):
            # Skip if already cached
            if ws.id in self._datasets_cache:
                total_datasets += len(self._datasets_cache[ws.id])
                if progress_callback:
                    progress_callback(i + 1, total_workspaces, ws.name)
                continue

            # Report progress
            if progress_callback:
                progress_callback(i + 1, total_workspaces, ws.name)

            # Load datasets for this workspace
            datasets, error = self.get_workspace_datasets(ws.id)
            if not error:
                total_datasets += len(datasets)

        self._all_datasets_loaded = True
        return total_datasets, None

    def search_datasets_cached(
        self,
        query: str,
        max_results: int = 100
    ) -> List[DatasetInfo]:
        """
        Search only cached datasets (no API calls).

        Use this for fast searching after preload_all_datasets() has been called.

        Args:
            query: Search query string
            max_results: Maximum results to return

        Returns:
            List of matching DatasetInfo
        """
        if not query:
            return []

        query_lower = query.lower()
        results: List[DatasetInfo] = []

        for datasets in self._datasets_cache.values():
            for ds in datasets:
                if query_lower in ds.name.lower():
                    results.append(ds)
                    if len(results) >= max_results:
                        return results

        # Sort by name for consistent ordering
        results.sort(key=lambda d: d.name.lower())
        return results

    def get_dataset_perspectives(
        self,
        workspace_name: str,
        dataset_name: str
    ) -> Tuple[List[str], Optional[str]]:
        """
        Fetch perspectives from a cloud dataset via XMLA/TOM.

        Requires Premium/PPU/Fabric capacity for XMLA endpoint access.
        Pro workspaces do not support XMLA and will return an empty list.

        Note: While perspective auto-discovery requires Premium, connecting
        to a perspective (using Cube= parameter) works with Pro workspaces
        if the user knows the perspective name.

        Args:
            workspace_name: Name of the workspace containing the dataset
            dataset_name: Name of the dataset/semantic model

        Returns:
            Tuple of (list of perspective names, error message if any)
            Empty list with no error = dataset has no perspectives or is Pro workspace
        """
        if not self._access_token:
            return [], "Not authenticated"

        perspectives = []

        try:
            # Try to import TOM via pbi_connector
            from core.pbi_connector import get_connector

            connector = get_connector()
            if not hasattr(connector, 'Server'):
                self.logger.debug("TOM not available - cannot fetch perspectives")
                return [], None  # Not an error, just not available

            # Build XMLA endpoint URL
            xmla_endpoint = f"powerbi://api.powerbi.com/v1.0/myorg/{workspace_name}"
            self.logger.info(f"Fetching perspectives from {xmla_endpoint}/{dataset_name}")

            # Connect to XMLA endpoint with OAuth token
            server = connector.Server()
            connection_string = (
                f"Provider=MSOLAP;"
                f"Data Source={xmla_endpoint};"
                f"Initial Catalog={dataset_name};"
                f"Password={self._access_token};"
                f"Persist Security Info=True;"
                f"Impersonation Level=Impersonate"
            )

            server.Connect(connection_string)

            # Get the database
            database = server.Databases.FindByName(dataset_name)
            if not database:
                server.Disconnect()
                return [], f"Dataset '{dataset_name}' not found"

            # Get perspectives from the model
            model = database.Model
            if model and hasattr(model, 'Perspectives'):
                for perspective in model.Perspectives:
                    perspectives.append(perspective.Name)
                self.logger.info(f"Found {len(perspectives)} perspectives: {perspectives}")

            server.Disconnect()

            # XMLA access succeeded - update workspace capacity_type if PPU
            self._update_workspace_capacity_type_if_ppu(workspace_name)

            return perspectives, None

        except Exception as e:
            error_str = str(e)
            self.logger.debug(f"Failed to fetch perspectives: {error_str}")

            # Handle common errors gracefully
            if "403" in error_str or "Forbidden" in error_str:
                # Pro workspace - XMLA not available
                self.logger.info("XMLA access denied (likely Pro workspace)")
                return [], None  # Not an error for callers - just unavailable
            elif "401" in error_str or "Unauthorized" in error_str:
                return [], "Authentication expired"
            else:
                # Log but don't propagate as error - perspectives are optional
                self.logger.warning(f"Could not fetch perspectives: {error_str}")
                return [], None

    def fetch_perspectives_for_dataset(self, dataset: DatasetInfo) -> DatasetInfo:
        """
        Fetch perspectives for a dataset and update the DatasetInfo object.

        This is a convenience method that updates the dataset in-place and
        marks perspectives_loaded = True regardless of whether perspectives
        were found (distinguishes "none" from "not checked").

        Args:
            dataset: The DatasetInfo to fetch perspectives for

        Returns:
            The same DatasetInfo with perspectives field populated
        """
        perspectives, error = self.get_dataset_perspectives(
            dataset.workspace_name,
            dataset.name
        )

        dataset.perspectives = perspectives if perspectives else []
        dataset.perspectives_loaded = True

        if error:
            self.logger.warning(f"Error fetching perspectives for {dataset.name}: {error}")

        return dataset

    def lookup_dataset_by_guid(
        self,
        dataset_guid: str,
        fetch_if_needed: bool = True
    ) -> Tuple[Optional[str], Optional[str]]:
        """
        Look up a dataset by its GUID to get the friendly name and workspace.

        This method:
        1. First checks the existing cache
        2. If not found and fetch_if_needed is True, tries silent authentication
        3. If authenticated, fetches workspaces and searches for the dataset

        Args:
            dataset_guid: The dataset GUID to look up
            fetch_if_needed: If True, try to authenticate and fetch if not in cache

        Returns:
            Tuple of (dataset_name, workspace_name) if found, (None, None) otherwise
        """
        guid_lower = dataset_guid.lower()

        # First check existing cache
        for workspace_id, datasets in self._datasets_cache.items():
            for dataset in datasets:
                if dataset.id.lower() == guid_lower:
                    workspace_name = getattr(dataset, 'workspace_name', None)
                    self.logger.info(f"Found dataset {dataset_guid} in cache: {dataset.name} ({workspace_name})")
                    return dataset.name, workspace_name

        if not fetch_if_needed:
            return None, None

        # Try silent authentication if not already authenticated
        if not self.is_authenticated():
            if not self._try_refresh_token():
                self.logger.debug("Silent auth failed, cannot look up dataset by GUID")
                return None, None
            self.logger.info("Silent auth successful, fetching workspaces to find dataset")

        # Fetch workspaces if we don't have any
        if not self._workspaces:
            _, error = self.get_workspaces("all")
            if error:
                self.logger.warning(f"Failed to fetch workspaces: {error}")
                return None, None

        # Search through workspaces (prioritize favorites first)
        favorites_first = sorted(self._workspaces, key=lambda w: (not w.is_favorite, w.name.lower()))

        for ws in favorites_first:
            # Fetch datasets for this workspace if not cached
            if ws.id not in self._datasets_cache:
                datasets, error = self.get_workspace_datasets(ws.id)
                if error:
                    continue

            # Check if the dataset is in this workspace
            if ws.id in self._datasets_cache:
                for dataset in self._datasets_cache[ws.id]:
                    if dataset.id.lower() == guid_lower:
                        workspace_name = getattr(dataset, 'workspace_name', None)
                        self.logger.info(f"Found dataset {dataset_guid}: {dataset.name} in {workspace_name}")
                        return dataset.name, workspace_name

        self.logger.debug(f"Dataset {dataset_guid} not found in any accessible workspace")
        return None, None
