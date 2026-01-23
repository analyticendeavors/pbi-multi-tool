"""
Preset Manager - Save and load environment presets
Built by Reid Havens of Analytic Endeavors

Manages SwapPreset persistence with dual storage support:
- User presets: stored in AppData (personal, persists across projects)
- Project presets: stored in project folder (shareable, version-controllable)

And two preset scopes:
- Global presets: Single target, works with any live-connection model
- Model presets: Full mapping snapshot, keyed by model file hash
"""

import datetime
import json
import logging
import os
import sys
from pathlib import Path
from typing import List, Optional, Dict

from tools.connection_hotswap.models import (
    SwapPreset,
    PresetStorageType,
    PresetScope,
    PresetTargetMapping,
    ConnectionMapping,
    SwapStatus,
)
from tools.connection_hotswap.logic.connection_cache import get_cache_manager, ConnectionCacheManager


class PresetManager:
    """
    Manages environment preset persistence.

    Supports two storage locations:
    - User: %APPDATA%/Analytic Endeavors/PBI Report Merger/hotswap_presets/
    - Project: .pbi-hotswap/presets/ in the project directory
    """

    PRESETS_FILENAME = "hotswap_presets.json"
    PROJECT_FOLDER_NAME = ".pbi-hotswap"

    def __init__(self, project_path: Optional[str] = None, report_path: Optional[str] = None):
        """
        Initialize the preset manager.

        Args:
            project_path: Optional path to project directory for project-level presets
            report_path: Optional path to PBIP report folder (.Report) for report-level presets
        """
        self.logger = logging.getLogger(__name__)
        self._project_path = project_path
        self._report_path = report_path

        # Initialize storage paths
        self._user_presets_dir = self._get_user_presets_dir()
        self._project_presets_dir = self._get_project_presets_dir()
        self._report_presets_dir = self._get_report_presets_dir()

        # Ensure user directory exists
        self._user_presets_dir.mkdir(parents=True, exist_ok=True)

    def _get_user_presets_dir(self) -> Path:
        """Get the user-level presets directory."""
        if getattr(sys, 'frozen', False):
            # Running as EXE
            base = Path(os.environ.get('APPDATA', ''))
        else:
            # Running from source
            base = Path(os.environ.get('APPDATA', ''))

        return base / 'Analytic Endeavors' / 'PBI Report Merger' / 'hotswap_presets'

    def _get_project_presets_dir(self) -> Optional[Path]:
        """Get the project-level presets directory."""
        if self._project_path:
            return Path(self._project_path) / self.PROJECT_FOLDER_NAME / 'presets'
        return None

    def _get_report_presets_dir(self) -> Optional[Path]:
        """Get the report-level presets directory (for PBIP files)."""
        if self._report_path:
            # Store in .pbi-hotswap folder at the PBIP root (parent of .Report folder)
            report_folder = Path(self._report_path)
            # If path ends with .Report, go to parent (PBIP root)
            if report_folder.name.endswith('.Report'):
                pbip_root = report_folder.parent
            else:
                pbip_root = report_folder
            return pbip_root / self.PROJECT_FOLDER_NAME / 'presets'
        return None

    def set_report_path(self, report_path: str) -> None:
        """
        Set the PBIP report folder path.

        Args:
            report_path: Path to the .Report folder or PBIP root
        """
        self._report_path = report_path
        self._report_presets_dir = self._get_report_presets_dir()
        self.logger.debug(f"Report presets dir set to: {self._report_presets_dir}")

    def set_project_path(self, project_path: str) -> None:
        """
        Set the project directory path.

        Args:
            project_path: Path to the project directory
        """
        self._project_path = project_path
        self._project_presets_dir = self._get_project_presets_dir()

    def _get_presets_file(self, storage_type: PresetStorageType) -> Path:
        """Get the presets file path for the given storage type."""
        if storage_type == PresetStorageType.USER:
            return self._user_presets_dir / self.PRESETS_FILENAME
        elif storage_type == PresetStorageType.REPORT:
            if not self._report_presets_dir:
                raise ValueError("Report path not set. Cannot access report presets.")
            return self._report_presets_dir / self.PRESETS_FILENAME
        else:  # PROJECT
            if not self._project_presets_dir:
                raise ValueError("Project path not set. Cannot access project presets.")
            return self._project_presets_dir / self.PRESETS_FILENAME

    def _load_presets_file(self, storage_type: PresetStorageType) -> Dict:
        """
        Load presets from a storage location.

        Returns a dict with structure:
        {
            'version': '2.0',
            'global_presets': { name: preset_dict, ... },
            'model_presets': { model_hash: { name: preset_dict, ... }, ... }
        }
        """
        try:
            file_path = self._get_presets_file(storage_type)
            if file_path.exists():
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # Handle v1.0 format migration
                if data.get('version', '1.0') == '1.0' and 'presets' in data:
                    # Migrate: all old presets become model presets (backwards compatible)
                    return {
                        'version': '2.0',
                        'global_presets': {},
                        'model_presets': {'_legacy': data.get('presets', {})},
                        'last_configs': {},
                        'settings': {}
                    }

                # Return v2.0 structure
                return {
                    'version': data.get('version', '2.0'),
                    'global_presets': data.get('global_presets', {}),
                    'model_presets': data.get('model_presets', {}),
                    'last_configs': data.get('last_configs', {}),
                    'settings': data.get('settings', {})
                }
        except Exception as e:
            self.logger.error(f"Error loading presets from {storage_type.value}: {e}")

        return {'version': '2.0', 'global_presets': {}, 'model_presets': {}, 'last_configs': {}, 'settings': {}}

    def _save_presets_file(self, storage_type: PresetStorageType, data: Dict) -> bool:
        """
        Save presets to a storage location.

        Args:
            storage_type: Where to save
            data: Dict with 'global_presets' and 'model_presets' keys
        """
        try:
            file_path = self._get_presets_file(storage_type)
            file_path.parent.mkdir(parents=True, exist_ok=True)

            save_data = {
                'version': '2.0',
                'global_presets': data.get('global_presets', {}),
                'model_presets': data.get('model_presets', {}),
                'last_configs': data.get('last_configs', {}),
                'settings': data.get('settings', {})
            }

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, indent=2)

            self.logger.info(f"Saved presets to {file_path}")
            return True

        except Exception as e:
            self.logger.error(f"Error saving presets to {storage_type.value}: {e}")
            return False

    def list_presets(
        self,
        include_project: bool = True,
        scope: Optional[PresetScope] = None,
        model_hash: Optional[str] = None
    ) -> List[SwapPreset]:
        """
        List available presets with optional filtering.

        Args:
            include_project: If True, include project-level presets
            scope: Filter by preset scope (GLOBAL or MODEL). None returns all.
            model_hash: For MODEL scope, filter to presets matching this model hash.
                       Also includes legacy presets (from v1.0 migration).

        Returns:
            List of SwapPreset objects matching the criteria
        """
        presets: List[SwapPreset] = []

        storage_types = [PresetStorageType.USER]
        if include_project and self._project_presets_dir:
            storage_types.append(PresetStorageType.PROJECT)

        for storage_type in storage_types:
            data = self._load_presets_file(storage_type)

            # Load global presets if scope allows
            if scope is None or scope == PresetScope.GLOBAL:
                for name, preset_dict in data.get('global_presets', {}).items():
                    try:
                        preset = SwapPreset.from_dict(preset_dict)
                        preset.storage_type = storage_type
                        preset.scope = PresetScope.GLOBAL
                        presets.append(preset)
                    except Exception as e:
                        self.logger.error(f"Error parsing global preset '{name}': {e}")

            # Load model presets if scope allows
            if scope is None or scope == PresetScope.MODEL:
                model_presets = data.get('model_presets', {})

                # Determine which model hashes to include
                if model_hash:
                    # Specific model + legacy presets
                    hash_keys = [model_hash]
                    if '_legacy' in model_presets:
                        hash_keys.append('_legacy')
                elif scope == PresetScope.MODEL:
                    # MODEL scope selected but no model connected - return nothing
                    # Model presets only show when their specific model is loaded
                    hash_keys = []
                else:
                    # scope is None (list all) - include all model presets
                    hash_keys = list(model_presets.keys())

                for hash_key in hash_keys:
                    if hash_key not in model_presets:
                        continue

                    for name, preset_dict in model_presets[hash_key].items():
                        try:
                            preset = SwapPreset.from_dict(preset_dict)
                            preset.storage_type = storage_type
                            preset.scope = PresetScope.MODEL
                            # Preserve model_hash from dict or use the key
                            if not preset.model_hash and hash_key != '_legacy':
                                preset.model_hash = hash_key
                            presets.append(preset)
                        except Exception as e:
                            self.logger.error(f"Error parsing model preset '{name}': {e}")

        return presets

    def get_preset(
        self,
        name: str,
        storage_type: Optional[PresetStorageType] = None,
        scope: Optional[PresetScope] = None,
        model_hash: Optional[str] = None
    ) -> Optional[SwapPreset]:
        """
        Get a preset by name.

        Args:
            name: Preset name
            storage_type: Where to look (if None, search both locations)
            scope: Preset scope to search in (if None, search all scopes)
            model_hash: For MODEL scope, the model hash to search under

        Returns:
            SwapPreset if found, None otherwise
        """
        storage_types = [storage_type] if storage_type else [
            PresetStorageType.PROJECT, PresetStorageType.USER
        ]

        for st in storage_types:
            if st == PresetStorageType.PROJECT and not self._project_presets_dir:
                continue

            data = self._load_presets_file(st)

            # Search global presets
            if scope is None or scope == PresetScope.GLOBAL:
                global_presets = data.get('global_presets', {})
                if name in global_presets:
                    preset = SwapPreset.from_dict(global_presets[name])
                    preset.storage_type = st
                    preset.scope = PresetScope.GLOBAL
                    return preset

            # Search model presets
            if scope is None or scope == PresetScope.MODEL:
                model_presets = data.get('model_presets', {})

                # Determine which hashes to search
                if model_hash:
                    hash_keys = [model_hash, '_legacy']
                else:
                    hash_keys = list(model_presets.keys())

                for hash_key in hash_keys:
                    if hash_key in model_presets and name in model_presets[hash_key]:
                        preset = SwapPreset.from_dict(model_presets[hash_key][name])
                        preset.storage_type = st
                        preset.scope = PresetScope.MODEL
                        if not preset.model_hash and hash_key != '_legacy':
                            preset.model_hash = hash_key
                        return preset

        return None

    def save_preset(self, preset: SwapPreset) -> bool:
        """
        Save a preset.

        Args:
            preset: The preset to save

        Returns:
            True if successful
        """
        preset.updated_at = datetime.datetime.now().isoformat()

        # Load existing presets for this storage type
        data = self._load_presets_file(preset.storage_type)

        if preset.scope == PresetScope.GLOBAL:
            # Global presets - simple name-keyed storage
            if 'global_presets' not in data:
                data['global_presets'] = {}
            data['global_presets'][preset.name] = preset.to_dict()
        else:
            # Model presets - keyed by model hash
            if 'model_presets' not in data:
                data['model_presets'] = {}

            # Use model_hash or '_legacy' for backwards compatibility
            hash_key = preset.model_hash or '_legacy'

            if hash_key not in data['model_presets']:
                data['model_presets'][hash_key] = {}

            data['model_presets'][hash_key][preset.name] = preset.to_dict()

        # Save back
        return self._save_presets_file(preset.storage_type, data)

    def delete_preset(
        self,
        name: str,
        storage_type: PresetStorageType,
        scope: Optional[PresetScope] = None,
        model_hash: Optional[str] = None
    ) -> bool:
        """
        Delete a preset.

        Args:
            name: Preset name
            storage_type: Where the preset is stored
            scope: Preset scope (if None, searches all scopes)
            model_hash: For MODEL scope, the model hash the preset is under

        Returns:
            True if deleted successfully
        """
        data = self._load_presets_file(storage_type)
        deleted = False

        # Try global presets first
        if scope is None or scope == PresetScope.GLOBAL:
            global_presets = data.get('global_presets', {})
            if name in global_presets:
                del global_presets[name]
                data['global_presets'] = global_presets
                deleted = True

        # Try model presets
        if not deleted and (scope is None or scope == PresetScope.MODEL):
            model_presets = data.get('model_presets', {})

            # Determine which hashes to search
            if model_hash:
                hash_keys = [model_hash, '_legacy']
            else:
                hash_keys = list(model_presets.keys())

            for hash_key in hash_keys:
                if hash_key in model_presets and name in model_presets[hash_key]:
                    del model_presets[hash_key][name]
                    # Clean up empty hash buckets
                    if not model_presets[hash_key]:
                        del model_presets[hash_key]
                    data['model_presets'] = model_presets
                    deleted = True
                    break

        if deleted:
            return self._save_presets_file(storage_type, data)

        return False

    def rename_preset(
        self,
        old_name: str,
        new_name: str,
        storage_type: PresetStorageType,
        scope: Optional[PresetScope] = None,
        model_hash: Optional[str] = None
    ) -> bool:
        """
        Rename a preset.

        Args:
            old_name: Current preset name
            new_name: New preset name
            storage_type: Where the preset is stored
            scope: Preset scope (if None, searches all scopes)
            model_hash: For MODEL scope, the model hash the preset is under

        Returns:
            True if renamed successfully
        """
        data = self._load_presets_file(storage_type)
        renamed = False

        # Try global presets first
        if scope is None or scope == PresetScope.GLOBAL:
            global_presets = data.get('global_presets', {})
            if old_name in global_presets:
                if new_name in global_presets:
                    self.logger.error(f"Global preset '{new_name}' already exists")
                    return False

                preset_dict = global_presets[old_name]
                preset_dict['name'] = new_name
                preset_dict['updated_at'] = datetime.datetime.now().isoformat()

                del global_presets[old_name]
                global_presets[new_name] = preset_dict
                data['global_presets'] = global_presets
                renamed = True

        # Try model presets
        if not renamed and (scope is None or scope == PresetScope.MODEL):
            model_presets = data.get('model_presets', {})

            # Determine which hashes to search
            if model_hash:
                hash_keys = [model_hash, '_legacy']
            else:
                hash_keys = list(model_presets.keys())

            for hash_key in hash_keys:
                if hash_key in model_presets and old_name in model_presets[hash_key]:
                    bucket = model_presets[hash_key]
                    if new_name in bucket:
                        self.logger.error(f"Model preset '{new_name}' already exists")
                        return False

                    preset_dict = bucket[old_name]
                    preset_dict['name'] = new_name
                    preset_dict['updated_at'] = datetime.datetime.now().isoformat()

                    del bucket[old_name]
                    bucket[new_name] = preset_dict
                    model_presets[hash_key] = bucket
                    data['model_presets'] = model_presets
                    renamed = True
                    break

        if renamed:
            return self._save_presets_file(storage_type, data)

        return False

    def create_preset_from_mappings(
        self,
        name: str,
        mappings: List[ConnectionMapping],
        storage_type: PresetStorageType = PresetStorageType.USER,
        description: Optional[str] = None,
        scope: PresetScope = PresetScope.MODEL,
        model_hash: Optional[str] = None,
        model_name: Optional[str] = None
    ) -> SwapPreset:
        """
        Create a new preset from current mappings.

        Args:
            name: Preset name
            mappings: Current connection mappings
            storage_type: Where to store the preset
            description: Optional description
            scope: Preset scope (GLOBAL for single-target, MODEL for full mapping)
            model_hash: Hash of the model file (required for MODEL scope)
            model_name: Display name of the model (optional)

        Returns:
            The created SwapPreset
        """
        preset_mappings: List[PresetTargetMapping] = []

        for mapping in mappings:
            if mapping.target:
                preset_mapping = PresetTargetMapping.from_swap_target(
                    connection_name=mapping.source.name,
                    target=mapping.target
                )
                preset_mappings.append(preset_mapping)

        preset = SwapPreset(
            name=name,
            description=description,
            mappings=preset_mappings,
            storage_type=storage_type,
            scope=scope,
            model_hash=model_hash if scope == PresetScope.MODEL else None,
            model_name=model_name if scope == PresetScope.MODEL else None,
        )

        return preset

    def apply_preset_to_mappings(
        self,
        preset: SwapPreset,
        mappings: List[ConnectionMapping]
    ) -> int:
        """
        Apply a preset's targets to a list of mappings.

        Args:
            preset: The preset to apply
            mappings: Current connection mappings to update

        Returns:
            Number of mappings updated
        """
        # Build lookup of preset mappings by connection name
        preset_lookup = {pm.connection_name: pm for pm in preset.mappings}

        updated = 0
        for mapping in mappings:
            if mapping.source.name in preset_lookup:
                preset_mapping = preset_lookup[mapping.source.name]
                mapping.target = preset_mapping.to_swap_target()
                mapping.status = SwapStatus.READY
                mapping.auto_matched = False
                updated += 1

        return updated

    def export_preset(self, preset: SwapPreset, file_path: str) -> bool:
        """
        Export a preset to a standalone JSON file.

        Args:
            preset: The preset to export
            file_path: Path to save the file

        Returns:
            True if successful
        """
        try:
            data = {
                'version': '1.0',
                'preset': preset.to_dict()
            }

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)

            return True

        except Exception as e:
            self.logger.error(f"Error exporting preset: {e}")
            return False

    def import_preset(
        self,
        file_path: str,
        storage_type: PresetStorageType = PresetStorageType.USER
    ) -> Optional[SwapPreset]:
        """
        Import a preset from a standalone JSON file.

        Args:
            file_path: Path to the preset file
            storage_type: Where to store the imported preset

        Returns:
            The imported SwapPreset, or None on error
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            preset_data = data.get('preset', data)  # Support both formats
            preset = SwapPreset.from_dict(preset_data)
            preset.storage_type = storage_type

            return preset

        except Exception as e:
            self.logger.error(f"Error importing preset: {e}")
            return None

    # -------------------------------------------------------------------------
    # Global Preset Helpers
    # -------------------------------------------------------------------------

    def create_global_preset(
        self,
        name: str,
        target_mapping: PresetTargetMapping,
        storage_type: PresetStorageType = PresetStorageType.USER,
        description: Optional[str] = None
    ) -> SwapPreset:
        """
        Create a global preset with a single target mapping.

        Global presets apply to any live-connection model (single connection).

        Args:
            name: Preset name
            target_mapping: The single target mapping for this global preset
            storage_type: Where to store the preset
            description: Optional description

        Returns:
            The created SwapPreset with GLOBAL scope
        """
        preset = SwapPreset(
            name=name,
            description=description,
            mappings=[target_mapping],
            storage_type=storage_type,
            scope=PresetScope.GLOBAL,
            model_hash=None,
            model_name=None,
        )

        return preset

    def can_apply_global_preset(
        self,
        preset: SwapPreset,
        mappings: List[ConnectionMapping]
    ) -> tuple[bool, str]:
        """
        Check if a global preset can be applied to the given mappings.

        Conditions for application:
        - Model must have exactly one connection (live connection, not composite)
        - The preset's target must be different from the current connection

        Args:
            preset: The global preset to check
            mappings: Current connection mappings

        Returns:
            Tuple of (can_apply: bool, reason: str)
        """
        if preset.scope != PresetScope.GLOBAL:
            return False, "Not a global preset"

        if len(mappings) != 1:
            return False, "Global presets only work with single-connection models"

        if not preset.mappings:
            return False, "Preset has no target mapping"

        # Get the preset's target
        preset_target = preset.mappings[0]
        current_mapping = mappings[0]

        # Check if current connection is the same as preset target
        if current_mapping.source:
            current_server = current_mapping.source.server or ""
            current_database = current_mapping.source.database or ""
            preset_server = preset_target.server or ""
            preset_database = preset_target.database or ""

            # Normalize for comparison
            if (current_server.lower() == preset_server.lower() and
                current_database.lower() == preset_database.lower()):
                return False, "Target connection is the same as current connection"

        return True, "OK"

    def apply_global_preset(
        self,
        preset: SwapPreset,
        mappings: List[ConnectionMapping]
    ) -> tuple[int, str]:
        """
        Apply a global preset to mappings.

        Args:
            preset: The global preset to apply
            mappings: Current connection mappings (should be exactly 1)

        Returns:
            Tuple of (updated_count: int, message: str)
        """
        can_apply, reason = self.can_apply_global_preset(preset, mappings)
        if not can_apply:
            return 0, reason

        # Apply the single target to the single mapping
        preset_mapping = preset.mappings[0]
        mapping = mappings[0]

        # Create target with the current connection's name
        mapping.target = preset_mapping.to_swap_target()
        mapping.status = SwapStatus.READY
        mapping.auto_matched = False

        return 1, "OK"

    # -------------------------------------------------------------------------
    # Last Config Memory
    # -------------------------------------------------------------------------

    def save_last_config(
        self,
        model_hash: str,
        mappings: List[ConnectionMapping],
        model_name: str,
        friendly_name: str = None,
        workspace_name: str = None
    ) -> bool:
        """
        Save the current connection configuration as the "last config" for a model.

        This saves the STARTING state before any swaps are made, allowing
        the user to restore their original configuration later.

        Args:
            model_hash: Hash identifying the model (from server+database)
            mappings: Current connection mappings to save
            model_name: Display name of the model
            friendly_name: Friendly display name of the model (e.g., "Cereal Model")
            workspace_name: Workspace name for cloud connections (e.g., "Fabric WS")

        Returns:
            True if saved successfully
        """
        if not model_hash:
            self.logger.warning("Cannot save last config: no model hash provided")
            return False

        # Build mapping data from current state
        mapping_data = []
        for mapping in mappings:
            if mapping.source:
                # Convert enum to string value for JSON serialization
                source_type = mapping.source.connection_type
                source_type_str = source_type.value if hasattr(source_type, 'value') else str(source_type)
                mapping_entry = {
                    'connection_name': mapping.source.name,
                    'source_server': mapping.source.server,
                    'source_database': mapping.source.database,
                    'source_type': source_type_str,
                    # Save friendly display info for proper restoration display
                    'source_friendly_name': friendly_name or mapping.source.database,
                    'source_workspace': workspace_name,
                    # Save cloud connection type for proper restoration (PBI Semantic Model vs XMLA)
                    'source_is_cloud': mapping.source.is_cloud,
                    # Save dataset_id (GUID) for cloud connections - required for pbiServiceLive
                    'source_dataset_id': getattr(mapping.source, 'dataset_id', None),
                }
                # Also save current target if any (this captures the starting state)
                if mapping.target:
                    mapping_entry['target_server'] = mapping.target.server
                    mapping_entry['target_database'] = mapping.target.database
                    mapping_entry['target_type'] = mapping.target.target_type
                    mapping_entry['target_display_name'] = mapping.target.display_name

                mapping_data.append(mapping_entry)

        # Load existing data and update last_configs section
        data = self._load_presets_file(PresetStorageType.USER)

        if 'last_configs' not in data:
            data['last_configs'] = {}

        data['last_configs'][model_hash] = {
            'model_name': model_name,
            'friendly_name': friendly_name,
            'workspace_name': workspace_name,
            'mappings': mapping_data,
            'saved_at': datetime.datetime.now().isoformat()
        }

        return self._save_presets_file(PresetStorageType.USER, data)

    def get_last_config(self, model_hash: str) -> Optional[Dict]:
        """
        Get the saved last config for a model.

        Args:
            model_hash: Hash identifying the model

        Returns:
            Dict with 'model_name', 'mappings', 'saved_at' if found, None otherwise
        """
        if not model_hash:
            return None

        data = self._load_presets_file(PresetStorageType.USER)
        last_configs = data.get('last_configs', {})

        return last_configs.get(model_hash)

    def has_last_config(self, model_hash: str) -> bool:
        """
        Check if a last config exists for a model.

        Args:
            model_hash: Hash identifying the model

        Returns:
            True if last config exists
        """
        return self.get_last_config(model_hash) is not None

    def update_last_config_workspace(self, model_hash: str, workspace_name: str, friendly_name: str = None) -> bool:
        """
        Update the workspace and friendly name in an existing Last Config.

        This preserves the original connection info while updating the workspace
        metadata, which is useful when swapping to a cloud target where
        the workspace info is known from the cloud browser selection.

        Args:
            model_hash: Hash identifying the model
            workspace_name: New workspace name to save
            friendly_name: Optional friendly name to update

        Returns:
            True if updated successfully
        """
        if not model_hash:
            return False

        data = self._load_presets_file(PresetStorageType.USER)
        last_configs = data.get('last_configs', {})

        if model_hash not in last_configs:
            return False

        # Update workspace (and optionally friendly name)
        last_configs[model_hash]['workspace_name'] = workspace_name
        if friendly_name:
            last_configs[model_hash]['friendly_name'] = friendly_name

        # Also update workspace in mapping entries if present
        mappings = last_configs[model_hash].get('mappings', [])
        for mapping_entry in mappings:
            mapping_entry['source_workspace'] = workspace_name
            if friendly_name:
                mapping_entry['source_friendly_name'] = friendly_name

        return self._save_presets_file(PresetStorageType.USER, data)

    def delete_last_config(self, model_hash: str) -> bool:
        """
        Delete the last config for a model.

        Args:
            model_hash: Hash identifying the model

        Returns:
            True if deleted successfully
        """
        if not model_hash:
            return False

        data = self._load_presets_file(PresetStorageType.USER)
        last_configs = data.get('last_configs', {})

        if model_hash in last_configs:
            del last_configs[model_hash]
            data['last_configs'] = last_configs
            return self._save_presets_file(PresetStorageType.USER, data)

        return False

    def apply_last_config_to_mappings(
        self,
        model_hash: str,
        mappings: List[ConnectionMapping]
    ) -> int:
        """
        Apply a saved last config to mappings (preview mode - sets targets).

        This sets the TARGET to the saved SOURCE, allowing users to swap back
        to their original connection state before the swap occurred.

        Args:
            model_hash: Hash identifying the model
            mappings: Current connection mappings to update

        Returns:
            Number of mappings updated
        """
        last_config = self.get_last_config(model_hash)
        if not last_config:
            return 0

        saved_mappings = last_config.get('mappings', [])

        # Build lookup by connection name
        saved_lookup = {m['connection_name']: m for m in saved_mappings}

        updated = 0
        for mapping in mappings:
            saved = None

            # First try exact name match
            if mapping.source and mapping.source.name in saved_lookup:
                saved = saved_lookup[mapping.source.name]
            # For thin reports: connection name changes between cloud/local states
            # If only one mapping exists in each, use it regardless of name
            elif len(mappings) == 1 and len(saved_mappings) == 1:
                saved = saved_mappings[0]
                self.logger.info(f"Last Config: Using single-mapping fallback for thin report")

            if saved and 'source_server' in saved:
                from tools.connection_hotswap.models import SwapTarget

                # Use the saved SOURCE as the target to swap back to
                # This allows restoring to the original connection state
                source_server = saved.get('source_server') or ''
                if not source_server:
                    self.logger.warning(f"Last Config: source_server is empty for mapping")
                    continue

                # Check for cloud endpoints: powerbi://, pbiazure://, asazure://, or use saved flag
                server_lower = source_server.lower()
                is_cloud = (
                    'powerbi://' in server_lower or
                    'pbiazure://' in server_lower or
                    'asazure://' in server_lower or
                    saved.get('source_is_cloud', False)
                )
                target_type = 'cloud' if is_cloud else 'local'

                # Build a friendly display name from saved info
                # Format: "Original: Model Name (Workspace)" for cloud
                # Format: "Original: database" for local
                friendly_name = saved.get('source_friendly_name')
                workspace = saved.get('source_workspace')

                # If workspace not saved, try to extract from server URL
                # e.g., pbiazure://api.powerbi.com/v1.0/myorg/Fabric%20WS -> "Fabric WS"
                if not workspace and is_cloud and source_server:
                    try:
                        import urllib.parse
                        # Look for workspace in the path segment after myorg/
                        if '/v1.0/' in source_server:
                            parts = source_server.split('/v1.0/')
                            if len(parts) > 1:
                                path_part = parts[1]
                                # Path is typically "myorg/WorkspaceName" or just "myorg"
                                segments = path_part.split('/')
                                if len(segments) >= 2 and segments[0] == 'myorg':
                                    workspace = urllib.parse.unquote(segments[1])
                    except Exception:
                        pass

                if friendly_name and workspace:
                    display_name = f"Original: {friendly_name} ({workspace})"
                elif friendly_name:
                    display_name = f"Original: {friendly_name}"
                elif workspace:
                    display_name = f"Original: {saved.get('source_database', 'Model')} ({workspace})"
                else:
                    display_name = f"Original: {saved.get('source_database', saved.get('connection_name', 'Original'))}"

                # Determine cloud connection type from server URL
                # pbiazure:// = PBI Semantic Model (pbiServiceLive)
                # powerbi:// or asazure:// = XMLA endpoint (analysisServicesDatabaseLive)
                from tools.connection_hotswap.models import CloudConnectionType
                cloud_conn_type = None
                if is_cloud:
                    if 'pbiazure://' in server_lower:
                        cloud_conn_type = CloudConnectionType.PBI_SEMANTIC_MODEL
                    else:
                        cloud_conn_type = CloudConnectionType.AAS_XMLA

                # Get dataset_id (GUID) if saved - required for pbiServiceLive connections
                dataset_id = saved.get('source_dataset_id')

                mapping.target = SwapTarget(
                    server=saved.get('source_server'),
                    database=saved.get('source_database'),
                    display_name=display_name,
                    target_type=target_type,
                    workspace_name=workspace,
                    cloud_connection_type=cloud_conn_type,
                    dataset_id=dataset_id
                )
                mapping.status = SwapStatus.READY
                mapping.auto_matched = False
                updated += 1

        return updated

    # -------------------------------------------------------------------------
    # Settings Management
    # -------------------------------------------------------------------------

    def get_setting(self, key: str, default=None):
        """
        Get a setting value.

        Args:
            key: Setting key (e.g., 'create_backup_before_swap')
            default: Default value if setting not found

        Returns:
            Setting value or default
        """
        data = self._load_presets_file(PresetStorageType.USER)
        settings = data.get('settings', {})
        return settings.get(key, default)

    def set_setting(self, key: str, value) -> bool:
        """
        Save a setting value.

        Args:
            key: Setting key (e.g., 'create_backup_before_swap')
            value: Value to save

        Returns:
            True if saved successfully
        """
        data = self._load_presets_file(PresetStorageType.USER)

        if 'settings' not in data:
            data['settings'] = {}

        data['settings'][key] = value

        return self._save_presets_file(PresetStorageType.USER, data)

    def get_backup_enabled(self) -> bool:
        """
        Get whether backup creation is enabled before swaps.

        Returns:
            True if backups should be created (default: False)
        """
        return self.get_setting('create_backup_before_swap', False)

    def set_backup_enabled(self, enabled: bool) -> bool:
        """
        Set whether backup creation is enabled before swaps.

        Args:
            enabled: True to enable backup creation

        Returns:
            True if saved successfully
        """
        return self.set_setting('create_backup_before_swap', enabled)

    # -------------------------------------------------------------------------
    # Bulk Global Preset Import/Export
    # -------------------------------------------------------------------------

    def export_all_global_presets(self, file_path: str) -> bool:
        """
        Export all global presets to a JSON file.

        Args:
            file_path: Path to save the export file

        Returns:
            True if successful
        """
        try:
            data = self._load_presets_file(PresetStorageType.USER)
            global_presets = data.get('global_presets', {})

            export_data = {
                'version': '1.0',
                'type': 'global_presets_export',
                'exported_at': datetime.datetime.now().isoformat(),
                'global_presets': global_presets
            }

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2)

            self.logger.info(f"Exported {len(global_presets)} global presets to {file_path}")
            return True

        except Exception as e:
            self.logger.error(f"Error exporting global presets: {e}")
            return False

    def import_all_global_presets(self, file_path: str) -> tuple[bool, int, str]:
        """
        Import global presets from a JSON file, replacing existing global presets.

        Args:
            file_path: Path to the import file

        Returns:
            Tuple of (success: bool, count: int, message: str)
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                import_data = json.load(f)

            # Validate file format
            if import_data.get('type') != 'global_presets_export':
                return False, 0, "Invalid file format. Expected global presets export file."

            imported_presets = import_data.get('global_presets', {})
            if not imported_presets:
                return False, 0, "No global presets found in file."

            # Load current data and replace global presets
            data = self._load_presets_file(PresetStorageType.USER)
            data['global_presets'] = imported_presets

            if self._save_presets_file(PresetStorageType.USER, data):
                count = len(imported_presets)
                self.logger.info(f"Imported {count} global presets from {file_path}")
                return True, count, f"Successfully imported {count} global preset(s)."
            else:
                return False, 0, "Failed to save imported presets."

        except json.JSONDecodeError:
            return False, 0, "Invalid JSON file."
        except Exception as e:
            self.logger.error(f"Error importing global presets: {e}")
            return False, 0, f"Error: {str(e)}"

    def get_global_preset_count(self) -> int:
        """Get the number of global presets currently stored."""
        data = self._load_presets_file(PresetStorageType.USER)
        return len(data.get('global_presets', {}))

    # -------------------------------------------------------------------------
    # Schema Validation for Cloud Connections
    # -------------------------------------------------------------------------

    def validate_preset_schema(
        self,
        preset: SwapPreset,
        file_path: str,
        file_type: str = "pbix"
    ) -> Dict:
        """
        Validate a preset's cloud connection schema against the current file.

        Compares the preset's stored schema fingerprint with the current file's
        connection schema to detect if the cloud connection format has changed.

        Args:
            preset: The preset to validate
            file_path: Path to the .pbix or .pbip file
            file_type: Either "pbix" or "pbip"

        Returns:
            Dict with:
                - matches: bool (True if schema matches or no validation needed)
                - has_cloud_mapping: bool (True if preset contains cloud mappings)
                - mismatched_mappings: List of mapping names with schema mismatches
                - details: str (human-readable explanation)
        """
        result = {
            'matches': True,
            'has_cloud_mapping': False,
            'mismatched_mappings': [],
            'details': ''
        }

        # Get the current file's cached connection (if any)
        disk_cache = get_cache_manager()
        current_cached = disk_cache.load_connection(file_path)

        mismatches = []
        details_parts = []

        for mapping in preset.mappings:
            if mapping.target_type != 'cloud':
                continue

            result['has_cloud_mapping'] = True

            # If preset has no fingerprint, skip validation (legacy preset)
            if not mapping.cloud_schema_fingerprint:
                self.logger.debug(f"Preset mapping '{mapping.connection_name}' has no schema fingerprint (legacy)")
                continue

            # Get current schema fingerprint
            current_fingerprint = None
            if current_cached:
                current_fingerprint = disk_cache.compute_schema_fingerprint(current_cached, file_type)

            # Compare fingerprints
            if current_fingerprint and mapping.cloud_schema_fingerprint != current_fingerprint:
                mismatches.append(mapping.connection_name)

                # Build detailed difference info
                if mapping.original_cloud_connection and current_cached:
                    diff = disk_cache.get_schema_diff(
                        {'_original_pbix_connection': mapping.original_cloud_connection},
                        current_cached,
                        file_type
                    )
                    if diff['differences']:
                        diff_strs = []
                        for d in diff['differences']:
                            diff_strs.append(f"  {d['field']}: preset={d['cached']} vs current={d['current']}")
                        details_parts.append(f"Mapping '{mapping.connection_name}':\n" + "\n".join(diff_strs))

        if mismatches:
            result['matches'] = False
            result['mismatched_mappings'] = mismatches
            result['details'] = "\n\n".join(details_parts) if details_parts else \
                f"Schema mismatch detected for: {', '.join(mismatches)}"

        return result

    def update_preset_schema(
        self,
        preset_name: str,
        file_path: str,
        file_type: str = "pbix",
        storage_type: PresetStorageType = PresetStorageType.USER,
        scope: PresetScope = PresetScope.MODEL,
        model_hash: Optional[str] = None
    ) -> bool:
        """
        Update a preset's cloud connection schema from the current file's cache.

        Called when user chooses to update a preset after a schema mismatch is detected.

        Args:
            preset_name: Name of the preset to update
            file_path: Path to the .pbix or .pbip file with the current schema
            file_type: Either "pbix" or "pbip"
            storage_type: Where the preset is stored
            scope: Preset scope
            model_hash: Model hash (required for MODEL scope)

        Returns:
            True if updated successfully
        """
        # Load the preset
        preset = self.get_preset(preset_name, storage_type, scope, model_hash)
        if not preset:
            self.logger.error(f"Preset '{preset_name}' not found")
            return False

        # Get the current file's cached connection
        disk_cache = get_cache_manager()
        current_cached = disk_cache.load_connection(file_path)

        if not current_cached:
            self.logger.warning(f"No cached connection found for {file_path}")
            return False

        # Compute current fingerprint
        current_fingerprint = disk_cache.compute_schema_fingerprint(current_cached, file_type)

        # Extract the original connection for storage
        original_conn = current_cached.get('_original_pbix_connection') or \
                       current_cached.get('_original_pbip_definition', {}).get('datasetReference', {}).get('byConnection', {})

        # Update each cloud mapping's schema
        updated = False
        for mapping in preset.mappings:
            if mapping.target_type == 'cloud':
                mapping.cloud_schema_fingerprint = current_fingerprint
                mapping.original_cloud_connection = original_conn.copy() if original_conn else None
                updated = True
                self.logger.info(f"Updated schema for mapping '{mapping.connection_name}'")

        if updated:
            preset.touch()  # Update timestamp
            return self.save_preset(preset)

        return True  # Nothing to update is not an error

    def capture_schema_for_preset(
        self,
        preset: SwapPreset,
        file_path: str,
        file_type: str = "pbix"
    ) -> SwapPreset:
        """
        Capture and store the current file's cloud connection schema in a preset.

        Called when creating a new preset to ensure the schema is captured
        for future validation.

        Args:
            preset: The preset to update (modified in place)
            file_path: Path to the .pbix or .pbip file
            file_type: Either "pbix" or "pbip"

        Returns:
            The updated preset
        """
        disk_cache = get_cache_manager()
        current_cached = disk_cache.load_connection(file_path)

        if not current_cached:
            self.logger.debug(f"No cached connection for {file_path} - schema capture skipped")
            return preset

        # Compute fingerprint
        fingerprint = disk_cache.compute_schema_fingerprint(current_cached, file_type)

        # Extract original connection
        original_conn = current_cached.get('_original_pbix_connection') or \
                       current_cached.get('_original_pbip_definition', {}).get('datasetReference', {}).get('byConnection', {})

        # Update cloud mappings
        for mapping in preset.mappings:
            if mapping.target_type == 'cloud':
                mapping.cloud_schema_fingerprint = fingerprint
                mapping.original_cloud_connection = original_conn.copy() if original_conn else None
                self.logger.debug(f"Captured schema for mapping '{mapping.connection_name}'")

        return preset
