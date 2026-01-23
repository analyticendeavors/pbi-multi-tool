"""
Report Cleanup Engine - Engine for removing unused themes and custom visuals
Built by Reid Havens of Analytic Endeavors

This module handles the actual removal of unused themes and custom visuals from Power BI reports.
"""

import json
import shutil
import os
import stat
import zipfile
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable
import datetime

# Import shared data classes
from tools.report_cleanup.shared_types import CleanupOpportunity, RemovalResult, DuplicateImageGroup


class ReportCleanupEngine:
    """
    Engine for removing unused themes and custom visuals from Power BI reports
    """
    
    def __init__(self, logger_callback: Optional[Callable[[str], None]] = None):
        self.log_callback = logger_callback or self._default_log
    
    def _default_log(self, message: str) -> None:
        """Default logging function"""
        print(message)
    
    def remove_unused_items(self, pbip_path: str, themes_to_remove: List[str] = None,
                           visuals_to_remove: List[CleanupOpportunity] = None,
                           bookmarks_to_remove: List[CleanupOpportunity] = None,
                           hide_visual_filters: bool = False,
                           dax_queries_to_remove: List[CleanupOpportunity] = None,
                           tmdl_scripts_to_remove: List[CleanupOpportunity] = None,
                           duplicate_groups: List[DuplicateImageGroup] = None,
                           duplicate_selections: Dict[str, str] = None,
                           unused_images_to_remove: List[CleanupOpportunity] = None,
                           create_backup: bool = True,
                           report_dir: Optional[Path] = None,
                           content_dir: Optional[Path] = None) -> List[RemovalResult]:
        """
        Remove unused themes, custom visuals, bookmarks, DAX queries, TMDL scripts, and images from a PBIP report

        Args:
            pbip_path: Path to the .pbip or .pbix file
            themes_to_remove: List of theme names to remove (None to skip)
            visuals_to_remove: List of CleanupOpportunity objects for visuals to remove (None to skip)
            bookmarks_to_remove: List of CleanupOpportunity objects for bookmarks to remove (None to skip)
            hide_visual_filters: Whether to hide all visual-level filters
            dax_queries_to_remove: List of CleanupOpportunity objects for DAX queries to remove (None to skip)
            tmdl_scripts_to_remove: List of CleanupOpportunity objects for TMDL scripts to remove (None to skip)
            duplicate_groups: List of DuplicateImageGroup objects to consolidate (None to skip)
            duplicate_selections: Dict mapping group_id to selected keeper image name (uses auto-selected if not provided)
            unused_images_to_remove: List of CleanupOpportunity objects for unused images to remove (None to skip)
            create_backup: Whether to create a backup before making changes
            report_dir: Optional pre-resolved report directory (for extracted PBIX PBIR)
            content_dir: Optional content directory for DAX queries and TMDL scripts (for extracted PBIX)

        Returns:
            List of RemovalResult objects
        """
        results = []

        pbip_file = Path(pbip_path)
        is_pbix = pbip_file.suffix.lower() == '.pbix'

        # Use provided report_dir or derive from pbip path
        if report_dir is None:
            report_dir = pbip_file.parent / f"{pbip_file.stem}.Report"

        # For semantic model operations, use content_dir if provided, otherwise derive from pbip
        if content_dir is not None:
            # For extracted PBIX: DAXQueries and TMDLScripts are at content_dir root
            semantic_model_dir = content_dir
        else:
            # For PBIP: they're in .SemanticModel folder
            semantic_model_dir = pbip_file.parent / f"{pbip_file.stem}.SemanticModel"

        if not report_dir.exists():
            raise FileNotFoundError(f"Report directory not found: {report_dir}")

        # Create backup if requested
        if create_backup:
            self._create_backup(pbip_path)

        # Remove themes
        if themes_to_remove:
            self.log_callback(f"🎨 Removing {len(themes_to_remove)} unused themes...")
            results.extend(self._remove_themes(report_dir, themes_to_remove))

        # Remove custom visuals (now handles different types)
        if visuals_to_remove:
            self.log_callback(f"🔮 Removing {len(visuals_to_remove)} unused custom visuals...")
            results.extend(self._remove_custom_visuals(report_dir, visuals_to_remove))

        # Remove bookmarks
        if bookmarks_to_remove:
            self.log_callback(f"📖 Removing {len(bookmarks_to_remove)} unused bookmarks...")
            results.extend(self._remove_bookmarks(report_dir, bookmarks_to_remove))

        # Hide visual filters
        if hide_visual_filters:
            self.log_callback("🎯 Hiding all visual-level filters...")
            results.extend(self._hide_visual_filters(report_dir))

        # Remove DAX queries
        if dax_queries_to_remove:
            self.log_callback(f"📝 Removing {len(dax_queries_to_remove)} DAX queries...")
            results.extend(self._remove_dax_queries(semantic_model_dir, dax_queries_to_remove))

        # Remove TMDL scripts
        if tmdl_scripts_to_remove:
            self.log_callback(f"📜 Removing {len(tmdl_scripts_to_remove)} TMDL scripts...")
            results.extend(self._remove_tmdl_scripts(semantic_model_dir, tmdl_scripts_to_remove))

        # Consolidate duplicate images (MUST update refs before deleting files)
        if duplicate_groups:
            self.log_callback(f"🔄 Consolidating {len(duplicate_groups)} duplicate image groups...")
            results.extend(self._consolidate_duplicate_images(report_dir, duplicate_groups, duplicate_selections or {}))

        # Remove unused images
        if unused_images_to_remove:
            self.log_callback(f"🗑️ Removing {len(unused_images_to_remove)} unused images...")
            results.extend(self._remove_unused_images(report_dir, unused_images_to_remove))

        # Clean up resourcePackages, publicCustomVisuals, and bookmarks in report.json
        if themes_to_remove or visuals_to_remove or bookmarks_to_remove:
            self._cleanup_report_references(report_dir, themes_to_remove or [], visuals_to_remove or [], bookmarks_to_remove or [])

        # For PBIX files, repack the modified content back into the archive
        if is_pbix and content_dir is not None:
            self.log_callback("📦 Repacking changes into PBIX file...")
            try:
                self._repack_pbix(pbip_file, content_dir)
                self.log_callback("  ✅ PBIX file updated successfully")
            except Exception as e:
                self.log_callback(f"  ❌ Failed to repack PBIX: {e}")
                # Add a failure result for the repack operation
                results.append(RemovalResult(
                    item_name="PBIX Repack",
                    item_type='repack',
                    success=False,
                    error_message=str(e),
                    bytes_freed=0
                ))

        self.log_callback("✅ Cleanup operation completed")
        return results
    
    def _create_backup(self, pbip_path: str) -> None:
        """Create a backup of the PBIP file and report directory"""
        self.log_callback("💾 Creating backup...")
        
        pbip_file = Path(pbip_path)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_suffix = f"_backup_{timestamp}"
        
        # Backup PBIP file
        backup_pbip = pbip_file.parent / f"{pbip_file.stem}{backup_suffix}.pbip"
        shutil.copy2(pbip_file, backup_pbip)
        
        # Backup Report directory
        report_dir = pbip_file.parent / f"{pbip_file.stem}.Report"
        backup_report_dir = pbip_file.parent / f"{pbip_file.stem}{backup_suffix}.Report"
        shutil.copytree(report_dir, backup_report_dir)
        
        self.log_callback(f"  ✅ Backup created: {backup_pbip.name}")
        self.log_callback(f"  ✅ Backup created: {backup_report_dir.name}")

    def _repack_pbix(self, pbix_path: Path, content_dir: Path) -> None:
        """
        Repack modified content back into a PBIX file.

        PBIX files are ZIP archives with specific structure. Power BI is VERY sensitive to:
        - File compression methods: DataModel MUST use ZIP_STORED, others use ZIP_DEFLATED
        - File metadata preservation

        STRATEGY: Use writestr() for all files with proper metadata preservation.
        Compare content to detect actual modifications vs unchanged files.
        Skip files that were deleted during cleanup.

        Args:
            pbix_path: Path to the original PBIX file
            content_dir: Path to the temp directory containing modified content
                         (with Report/, DAXQueries/, TMDLScripts/ subfolders)
        """
        import time

        # Create a temporary file for the new PBIX
        temp_pbix = pbix_path.parent / f"{pbix_path.stem}_temp.pbix"

        # Get list of folders we've modified (everything in content_dir)
        modified_prefixes = set()
        for item in content_dir.iterdir():
            if item.is_dir():
                modified_prefixes.add(item.name + '/')

        self.log_callback(f"    Modified prefixes: {modified_prefixes}")

        # Build set of files that exist in content_dir (for detecting deletions)
        existing_in_temp = set()
        modified_content = {}  # archive_path -> new_content (bytes)

        for prefix in modified_prefixes:
            folder_path = content_dir / prefix.rstrip('/')
            if folder_path.exists():
                for file_path in folder_path.rglob('*'):
                    if file_path.is_file():
                        rel_path = file_path.relative_to(content_dir)
                        archive_path = str(rel_path).replace('\\', '/')
                        existing_in_temp.add(archive_path)

        self.log_callback(f"    Files in temp dir: {len(existing_in_temp)}")

        try:
            # Track file counts and sizes for debugging
            copied_count = 0
            copied_bytes = 0
            skipped_count = 0
            skipped_files = []
            modified_count = 0
            modified_bytes = 0

            with zipfile.ZipFile(pbix_path, 'r') as original_zip:
                # Log original ZIP contents
                original_files = [f for f in original_zip.infolist() if not f.is_dir()]
                original_total_size = sum(f.file_size for f in original_files)
                self.log_callback(f"    Original PBIX: {len(original_files)} files, {original_total_size:,} bytes uncompressed")

                # Pre-read content from temp dir to check for modifications
                for archive_path in existing_in_temp:
                    temp_file_path = content_dir / archive_path.replace('/', os.sep)
                    with open(temp_file_path, 'rb') as f:
                        temp_content = f.read()

                    # Check if this file existed in original
                    try:
                        original_content = original_zip.read(archive_path)
                        if temp_content != original_content:
                            # Content was modified
                            modified_content[archive_path] = temp_content
                    except KeyError:
                        # New file (not in original)
                        modified_content[archive_path] = temp_content

                self.log_callback(f"    Files actually modified: {len(modified_content)}")

                # Open the new ZIP file
                with zipfile.ZipFile(temp_pbix, 'w') as new_zip:
                    # Process all files from original PBIX
                    for item in original_zip.infolist():
                        # Skip directories
                        if item.is_dir():
                            continue

                        # Check if this file is in one of the modified directories
                        in_modified_prefix = any(item.filename.startswith(prefix) for prefix in modified_prefixes)

                        if not in_modified_prefix:
                            # Not in modified prefix - copy from original
                            data = original_zip.read(item.filename)
                            self._write_zip_entry(new_zip, item, data)
                            copied_count += 1
                            copied_bytes += len(data)

                        elif item.filename in existing_in_temp:
                            # File is in modified prefix AND still exists in temp dir
                            if item.filename in modified_content:
                                # Content was actually modified - use new content
                                data = modified_content[item.filename]
                                self._write_zip_entry(new_zip, item, data)
                                modified_count += 1
                                modified_bytes += len(data)
                            else:
                                # Content unchanged - copy from original
                                data = original_zip.read(item.filename)
                                self._write_zip_entry(new_zip, item, data)
                                copied_count += 1
                                copied_bytes += len(data)
                        else:
                            # File was DELETED during cleanup - skip it
                            skipped_count += 1
                            skipped_files.append(f"{item.filename} ({item.file_size:,} bytes)")

                    # Add any NEW files that weren't in original
                    for archive_path, data in modified_content.items():
                        try:
                            # Check if we already wrote this file (existed in original)
                            original_zip.getinfo(archive_path)
                        except KeyError:
                            # New file - write it
                            now = time.localtime()
                            new_info = zipfile.ZipInfo(archive_path)
                            new_info.compress_type = zipfile.ZIP_DEFLATED
                            new_info.date_time = (now.tm_year, now.tm_mon, now.tm_mday,
                                                  now.tm_hour, now.tm_min, now.tm_sec)
                            new_zip.writestr(new_info, data)
                            modified_count += 1
                            modified_bytes += len(data)

            self.log_callback(f"    Copied {copied_count} unchanged files ({copied_bytes:,} bytes)")
            self.log_callback(f"    Modified {modified_count} files ({modified_bytes:,} bytes)")
            self.log_callback(f"    Skipped {skipped_count} deleted files")
            if skipped_files:
                for sf in skipped_files[:10]:  # Show first 10
                    self.log_callback(f"      - {sf}")
                if len(skipped_files) > 10:
                    self.log_callback(f"      ... and {len(skipped_files) - 10} more")
            self.log_callback(f"    Total: {copied_count + modified_count} files, {copied_bytes + modified_bytes:,} bytes")

            # Verify the new PBIX before replacing
            with zipfile.ZipFile(temp_pbix, 'r') as verify_zip:
                new_files = [f for f in verify_zip.infolist() if not f.is_dir()]
                new_total_size = sum(f.file_size for f in new_files)
                self.log_callback(f"    New PBIX: {len(new_files)} files, {new_total_size:,} bytes uncompressed")

                # Check for DataModel
                try:
                    dm_info = verify_zip.getinfo('DataModel')
                    self.log_callback(f"    DataModel: {dm_info.file_size:,} bytes, compression={dm_info.compress_type}")
                except KeyError:
                    self.log_callback("    WARNING: DataModel not found in new PBIX!")

            # Replace original with new PBIX
            pbix_path.unlink()
            temp_pbix.rename(pbix_path)

        except Exception as e:
            # Clean up temp file on failure
            if temp_pbix.exists():
                temp_pbix.unlink()
            raise

    def _write_zip_entry(self, new_zip: zipfile.ZipFile, orig_info: zipfile.ZipInfo, data: bytes) -> None:
        """
        Write a ZIP entry preserving original metadata.

        Args:
            new_zip: Target ZipFile to write to
            orig_info: Original ZipInfo with metadata to preserve
            data: File content to write (uncompressed)
        """
        # Clone ZipInfo attributes for compatibility
        new_info = zipfile.ZipInfo(orig_info.filename)

        # Preserve original compression, but FORCE DataModel to be STORED
        if orig_info.filename == 'DataModel':
            new_info.compress_type = zipfile.ZIP_STORED
        else:
            new_info.compress_type = orig_info.compress_type

        new_info.date_time = orig_info.date_time
        new_info.external_attr = orig_info.external_attr
        new_info.internal_attr = orig_info.internal_attr
        new_info.comment = orig_info.comment
        new_info.extra = orig_info.extra
        new_zip.writestr(new_info, data)

    def _remove_themes(self, report_dir: Path, themes_to_remove: List[str]) -> List[RemovalResult]:
        """Remove unused theme files"""
        results = []
        
        for theme_name in themes_to_remove:
            self.log_callback(f"  🗑️ Removing theme: {theme_name}")
            
            bytes_freed = 0
            success = False
            error_message = ""
            
            try:
                # Check SharedResources/BaseThemes
                base_theme_file = report_dir / "StaticResources" / "SharedResources" / "BaseThemes" / f"{theme_name}.json"
                if base_theme_file.exists():
                    bytes_freed += base_theme_file.stat().st_size
                    base_theme_file.unlink()
                    self.log_callback(f"    ✅ Removed BaseTheme file: {base_theme_file.name}")
                    success = True
                
                # Check RegisteredResources for custom themes
                reg_resources_dir = report_dir / "StaticResources" / "RegisteredResources"
                if reg_resources_dir.exists():
                    for theme_file in reg_resources_dir.glob("*.json"):
                        if theme_file.name == theme_name or theme_file.stem == theme_name:
                            bytes_freed += theme_file.stat().st_size
                            theme_file.unlink()
                            self.log_callback(f"    ✅ Removed CustomTheme file: {theme_file.name}")
                            success = True
                
                if not success:
                    error_message = f"Theme file not found: {theme_name}"
                    self.log_callback(f"    ⚠️ {error_message}")
            
            except Exception as e:
                error_message = str(e)
                self.log_callback(f"    ❌ Error removing theme {theme_name}: {error_message}")
            
            results.append(RemovalResult(
                item_name=theme_name,
                item_type='theme',
                success=success,
                error_message=error_message,
                bytes_freed=bytes_freed
            ))
        
        return results
    
    def _remove_custom_visuals(self, report_dir: Path, visuals_to_remove: List) -> List[RemovalResult]:
        """Remove unused custom visuals with support for different visual types"""
        results = []
        
        custom_visuals_dir = report_dir / "CustomVisuals"
        
        for visual_opportunity in visuals_to_remove:
            # Handle both CleanupOpportunity objects and simple strings for backwards compatibility
            if hasattr(visual_opportunity, 'visual_id'):
                visual_id = visual_opportunity.visual_id
                display_name = visual_opportunity.item_name
                item_type = visual_opportunity.item_type
            else:
                # Backwards compatibility for simple string IDs
                visual_id = str(visual_opportunity)
                display_name = visual_id
                item_type = 'custom_visual_build_pane'
            
            self.log_callback(f"  🗑️ Removing {item_type}: {display_name}")
            
            bytes_freed = 0
            success = False
            error_message = ""
            
            try:
                # Calculate bytes before removal attempt
                visual_dir = custom_visuals_dir / visual_id
                if visual_dir.exists() and visual_dir.is_dir():
                    bytes_freed = sum(f.stat().st_size for f in visual_dir.rglob('*') if f.is_file())
                    self.log_callback(f"    📁 Directory size: {self._format_bytes(bytes_freed)}")
                
                # Handle different types of custom visual removal
                if item_type == 'custom_visual_hidden':
                    # Hidden visuals: only remove from CustomVisuals directory
                    if visual_dir.exists() and visual_dir.is_dir():
                        self._force_remove_directory(visual_dir)
                        self.log_callback(f"    ✅ Removed hidden visual directory: {visual_id}")
                        success = True
                    else:
                        error_message = f"Hidden visual directory not found: {visual_id}"
                        self.log_callback(f"    ⚠️ {error_message}")
                
                elif item_type == 'custom_visual_build_pane':
                    # Build pane visuals: remove from CustomVisuals directory if exists
                    if visual_dir.exists() and visual_dir.is_dir():
                        self._force_remove_directory(visual_dir)
                        self.log_callback(f"    ✅ Removed visual directory: {visual_id}")
                    else:
                        # Even if no directory exists, removing from build pane is still successful
                        self.log_callback(f"    ✅ No local directory found (AppSource visual)")
                    
                    # Note: resourcePackages and publicCustomVisuals cleanup happens in _cleanup_report_references
                    success = True
                
                else:
                    # Fallback for unknown types
                    if visual_dir.exists() and visual_dir.is_dir():
                        self._force_remove_directory(visual_dir)
                        self.log_callback(f"    ✅ Removed visual directory: {visual_id}")
                        success = True
            
            except Exception as e:
                error_message = str(e)
                self.log_callback(f"    ❌ Error removing visual {display_name}: {error_message}")
            
            results.append(RemovalResult(
                item_name=display_name,
                item_type='custom_visual',
                success=success,
                error_message=error_message,
                bytes_freed=bytes_freed
            ))
        
        return results
    
    def _remove_bookmarks(self, report_dir: Path, bookmarks_to_remove: List) -> List[RemovalResult]:
        """Remove unused bookmarks from the report"""
        results = []
        
        for bookmark_opportunity in bookmarks_to_remove:
            bookmark_id = bookmark_opportunity.bookmark_id
            display_name = bookmark_opportunity.item_name
            item_type = bookmark_opportunity.item_type
            
            self.log_callback(f"  🗑️ Removing {item_type}: {display_name}")
            
            success = True  # Bookmark removal from report.json will be handled in _cleanup_report_references
            error_message = ""
            
            # Bookmarks are stored in report.json, so removal is handled there
            # We just track the success here
            try:
                self.log_callback(f"    ✅ Marked for removal: {display_name} (ID: {bookmark_id})")
            except Exception as e:
                error_message = str(e)
                success = False
                self.log_callback(f"    ❌ Error marking bookmark for removal {display_name}: {error_message}")
            
            results.append(RemovalResult(
                item_name=display_name,
                item_type='bookmark',
                success=success,
                error_message=error_message,
                bytes_freed=0  # Bookmarks don't take significant space
            ))
        
        return results
    
    def _force_remove_directory(self, directory_path: Path) -> None:
        """Force remove a directory, handling Windows file permission issues"""
        def handle_remove_readonly(func, path, exc):
            """Error handler for shutil.rmtree to handle read-only files"""
            try:
                # Clear the readonly bit and try again
                os.chmod(path, stat.S_IWRITE)
                func(path)
            except Exception:
                # If that doesn't work, try to force delete
                try:
                    if os.path.isfile(path):
                        os.remove(path)
                    elif os.path.isdir(path):
                        os.rmdir(path)
                except Exception:
                    # Last resort - leave it and continue
                    pass
        
        try:
            # First try normal removal
            shutil.rmtree(directory_path)
        except PermissionError:
            # If that fails, try with error handler for read-only files
            try:
                shutil.rmtree(directory_path, onerror=handle_remove_readonly)
            except Exception:
                # Final fallback - manually walk and delete
                self._manual_directory_cleanup(directory_path)
    
    def _manual_directory_cleanup(self, directory_path: Path) -> None:
        """Manually clean up directory contents when normal deletion fails"""
        try:
            # Walk through all files and subdirectories
            for root, dirs, files in os.walk(directory_path, topdown=False):
                # Remove files first
                for file_name in files:
                    file_path = Path(root) / file_name
                    try:
                        # Make file writable and delete
                        os.chmod(file_path, stat.S_IWRITE)
                        file_path.unlink()
                    except Exception:
                        # Skip files that can't be deleted
                        pass
                
                # Remove empty directories
                for dir_name in dirs:
                    dir_path = Path(root) / dir_name
                    try:
                        if dir_path.exists() and not any(dir_path.iterdir()):
                            dir_path.rmdir()
                    except Exception:
                        # Skip directories that can't be removed
                        pass
            
            # Finally try to remove the main directory
            if directory_path.exists():
                try:
                    directory_path.rmdir()
                except Exception:
                    # If we can't remove the main directory, that's okay
                    # The contents are cleaned up
                    pass
                    
        except Exception:
            # If manual cleanup fails, just continue
            pass
    
    def _format_bytes(self, bytes_count: int) -> str:
        """Format bytes into human readable format"""
        if bytes_count == 0:
            return "0 B"
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if bytes_count < 1024:
                return f"{bytes_count:.1f} {unit}"
            bytes_count /= 1024
        return f"{bytes_count:.1f} TB"
    
    def _cleanup_report_references(self, report_dir: Path, removed_themes: List[str], removed_visuals: List, removed_bookmarks: List = None) -> None:
        self.log_callback("🧹 Cleaning up report.json references...")
        
        report_json_path = report_dir / "definition" / "report.json"
        if not report_json_path.exists():
            return
        
        try:
            with open(report_json_path, 'r', encoding='utf-8') as f:
                report_data = json.load(f)
            
            # Extract visual IDs from opportunities
            visual_ids_to_remove = []
            for visual_item in removed_visuals:
                if hasattr(visual_item, 'visual_id'):
                    visual_ids_to_remove.append(visual_item.visual_id)
                else:
                    visual_ids_to_remove.append(str(visual_item))
            
            # Clean up publicCustomVisuals (AppSource visuals)
            if 'publicCustomVisuals' in report_data:
                original_public = report_data['publicCustomVisuals'][:]
                updated_public = [vid for vid in original_public if vid not in visual_ids_to_remove]
                
                if len(updated_public) != len(original_public):
                    report_data['publicCustomVisuals'] = updated_public
                    removed_count = len(original_public) - len(updated_public)
                    self.log_callback(f"    ✅ Removed {removed_count} references from publicCustomVisuals")
            
            # Clean up resourcePackages
            resource_packages = report_data.get('resourcePackages', [])
            updated_packages = []
            
            for package in resource_packages:
                package_type = package.get('type', '')
                package_name = package.get('name', '')
                
                # Handle CustomVisual packages
                if package_type == 'CustomVisual':
                    if package_name not in visual_ids_to_remove:
                        updated_packages.append(package)
                    else:
                        self.log_callback(f"    ✅ Removed CustomVisual package: {package_name}")
                
                # Handle SharedResources and RegisteredResources packages  
                elif package_type in ['SharedResources', 'RegisteredResources']:
                    items = package.get('items', [])
                    updated_items = []
                    
                    for item in items:
                        item_name = item.get('name', '')
                        item_type = item.get('type', '')
                        
                        # Remove theme items
                        if item_type in ['BaseTheme', 'CustomTheme'] and item_name in removed_themes:
                            self.log_callback(f"    ✅ Removed {item_type} reference: {item_name}")
                        else:
                            updated_items.append(item)
                    
                    # Only keep the package if it still has items or is not empty
                    if updated_items or package_type == 'SharedResources':  # Always keep SharedResources
                        package['items'] = updated_items
                        updated_packages.append(package)
                    else:
                        self.log_callback(f"    ✅ Removed empty {package_type} package")
                
                else:
                    # Keep other packages unchanged
                    updated_packages.append(package)
            
            # Clean up bookmarks from PBIP structure
            if removed_bookmarks:
                bookmark_ids_to_remove = []
                parent_groups_to_remove = []  # Track parent groups that become empty
                
                for bookmark_item in removed_bookmarks:
                    if hasattr(bookmark_item, 'bookmark_id'):
                        bookmark_id = bookmark_item.bookmark_id
                        bookmark_ids_to_remove.append(bookmark_id)
                        # Check if this is a parent group removal
                        if hasattr(bookmark_item, 'item_type') and bookmark_item.item_type == 'bookmark_empty_group':
                            parent_groups_to_remove.append(bookmark_id)
                    else:
                        bookmark_ids_to_remove.append(str(bookmark_item))
                
                # Remove bookmark files from bookmarks directory
                bookmarks_dir = report_dir / "definition" / "bookmarks"
                if bookmarks_dir.exists():
                    bookmarks_json_path = bookmarks_dir / "bookmarks.json"
                    if bookmarks_json_path.exists():
                        try:
                            with open(bookmarks_json_path, 'r', encoding='utf-8') as f:
                                bookmarks_data = json.load(f)
                            
                            # Process bookmark items in bookmarks.json
                            original_items = bookmarks_data.get('items', [])
                            updated_items = []
                            
                            for item in original_items:
                                item_id = item.get('name', '')
                                
                                # Skip items that are being completely removed (parent groups or individual bookmarks)
                                if item_id in bookmark_ids_to_remove:
                                    self.log_callback(f"    ✅ Removing bookmark/group from bookmarks.json: {item.get('displayName', item_id)}")
                                    continue
                                
                                # For parent groups that are NOT being removed, update their children lists
                                if 'children' in item:
                                    original_children = item.get('children', [])
                                    # Remove any children that are being deleted
                                    updated_children = [child_id for child_id in original_children if child_id not in bookmark_ids_to_remove]
                                    
                                    if len(updated_children) != len(original_children):
                                        removed_children_count = len(original_children) - len(updated_children)
                                        self.log_callback(f"    ✅ Removed {removed_children_count} child bookmark(s) from group: {item.get('displayName', item_id)}")
                                        
                                        # Update the children list
                                        item['children'] = updated_children
                                
                                # Keep this item (either unchanged or with updated children)
                                updated_items.append(item)
                            
                            # Update bookmarks.json with the modified structure
                            if len(updated_items) != len(original_items) or any('children' in item for item in updated_items):
                                bookmarks_data['items'] = updated_items
                                with open(bookmarks_json_path, 'w', encoding='utf-8') as f:
                                    json.dump(bookmarks_data, f, indent=2)
                                
                                total_removed = len(original_items) - len(updated_items)
                                self.log_callback(f"    ✅ Updated bookmarks.json: removed {total_removed} bookmark/group entries")
                            
                            # Remove individual bookmark files
                            for bookmark_id in bookmark_ids_to_remove:
                                bookmark_file = bookmarks_dir / f"{bookmark_id}.bookmark.json"
                                if bookmark_file.exists():
                                    try:
                                        bookmark_file.unlink()
                                        self.log_callback(f"    ✅ Removed bookmark file: {bookmark_id}.bookmark.json")
                                    except Exception as e:
                                        self.log_callback(f"    ⚠️ Could not remove bookmark file {bookmark_id}: {e}")
                                        
                        except Exception as e:
                            self.log_callback(f"    ⚠️ Could not update bookmarks.json: {e}")
            
            # Update report.json
            report_data['resourcePackages'] = updated_packages
            
            with open(report_json_path, 'w', encoding='utf-8') as f:
                json.dump(report_data, f, indent=2)
            
            self.log_callback("  ✅ Updated report.json references")
        
        except Exception as e:
            self.log_callback(f"  ⚠️ Warning: Could not clean up report references: {e}")
    
    def _hide_visual_filters(self, report_dir: Path) -> List[RemovalResult]:
        """Hide all visual-level filters by setting isHiddenInViewMode to true"""
        results = []
        
        pages_dir = report_dir / "definition" / "pages"
        if not pages_dir.exists():
            return results
        
        total_filters_hidden = 0
        total_visuals_with_filters = 0
        total_pages_processed = 0
        
        try:
            page_dirs = [d for d in pages_dir.iterdir() if d.is_dir() and d.name != "pages.json"]
            
            for page_dir in page_dirs:
                page_name = page_dir.name
                
                # Get page display name
                page_json = page_dir / "page.json"
                page_display_name = page_name
                if page_json.exists():
                    try:
                        with open(page_json, 'r', encoding='utf-8') as f:
                            page_data = json.load(f)
                        page_display_name = page_data.get('displayName', page_name)
                    except:
                        pass
                
                self.log_callback(f"  📊 Scanning page '{page_display_name}' for visual-level filters...")
                
                visuals_dir = page_dir / "visuals"
                if not visuals_dir.exists():
                    continue
                
                page_filters_hidden = 0
                page_visuals_with_filters = 0
                
                visual_dirs = [d for d in visuals_dir.iterdir() if d.is_dir()]
                
                for visual_dir in visual_dirs:
                    visual_json = visual_dir / "visual.json"
                    if not visual_json.exists():
                        continue
                    
                    try:
                        with open(visual_json, 'r', encoding='utf-8') as f:
                            visual_data = json.load(f)
                        
                        # Check for visual-level filters
                        filter_config = visual_data.get('filterConfig', {})
                        filters = filter_config.get('filters', [])
                        
                        if filters:
                            # Hide all visible filters
                            filters_hidden_in_visual = 0
                            for filter_obj in filters:
                                if not filter_obj.get('isHiddenInViewMode', False):
                                    filter_obj['isHiddenInViewMode'] = True
                                    filters_hidden_in_visual += 1
                            
                            if filters_hidden_in_visual > 0:
                                # Save the modified visual.json
                                with open(visual_json, 'w', encoding='utf-8') as f:
                                    json.dump(visual_data, f, indent=2)
                                
                                page_filters_hidden += filters_hidden_in_visual
                                page_visuals_with_filters += 1
                        
                    except Exception as e:
                        self.log_callback(f"    ❌ Error processing visual {visual_dir.name}: {e}")
                        continue
                
                if page_filters_hidden > 0:
                    total_filters_hidden += page_filters_hidden
                    total_visuals_with_filters += page_visuals_with_filters
                    total_pages_processed += 1
                    
                    self.log_callback(f"  ✅ Page '{page_display_name}': {page_filters_hidden} filters hidden on {page_visuals_with_filters} visuals")
            
            # Create a single result summarizing all filter hiding
            if total_filters_hidden > 0:
                result = RemovalResult(
                    item_name=f"{total_visuals_with_filters} visuals ({total_filters_hidden} filters)",
                    item_type='visual_filter',
                    success=True,
                    error_message="",
                    bytes_freed=0,
                    filters_hidden=total_filters_hidden
                )
                results.append(result)
                
                self.log_callback(f"  ✅ Successfully hidden {total_filters_hidden} visual-level filters across {total_visuals_with_filters} visuals on {total_pages_processed} pages")
            else:
                self.log_callback("  📌 No visible visual-level filters found to hide")
                
        except Exception as e:
            error_result = RemovalResult(
                item_name="Visual-level filters",
                item_type='visual_filter',
                success=False,
                error_message=str(e),
                bytes_freed=0,
                filters_hidden=0
            )
            results.append(error_result)
            self.log_callback(f"  ❌ Error hiding visual filters: {e}")

        return results

    def _remove_dax_queries(self, semantic_model_dir: Path, queries_to_remove: List[CleanupOpportunity]) -> List[RemovalResult]:
        """Remove DAX query files and update the manifest"""
        results = []

        dax_queries_dir = semantic_model_dir / "DAXQueries"
        manifest_path = dax_queries_dir / ".pbi" / "daxQueries.json"

        if not dax_queries_dir.exists():
            self.log_callback("  ⚠️ DAXQueries directory not found")
            return results

        # Load the manifest to update tab order
        manifest_data = None
        if manifest_path.exists():
            try:
                with open(manifest_path, 'r', encoding='utf-8') as f:
                    manifest_data = json.load(f)
            except Exception as e:
                self.log_callback(f"  ⚠️ Could not read DAX queries manifest: {e}")

        for query_opportunity in queries_to_remove:
            query_name = query_opportunity.item_name

            self.log_callback(f"  🗑️ Removing DAX query: {query_name}")

            bytes_freed = 0
            success = False
            error_message = ""

            try:
                # Find and remove the .dax file
                query_file = dax_queries_dir / f"{query_name}.dax"

                if query_file.exists():
                    bytes_freed = query_file.stat().st_size
                    query_file.unlink()
                    self.log_callback(f"    ✅ Removed DAX file: {query_name}.dax ({self._format_bytes(bytes_freed)})")
                    success = True

                    # Update manifest if loaded
                    if manifest_data and 'tabOrder' in manifest_data:
                        if query_name in manifest_data['tabOrder']:
                            manifest_data['tabOrder'].remove(query_name)
                            self.log_callback(f"    ✅ Removed from manifest tabOrder")

                        # Update defaultTab if it was this query
                        if manifest_data.get('defaultTab') == query_name:
                            if manifest_data['tabOrder']:
                                manifest_data['defaultTab'] = manifest_data['tabOrder'][0]
                            else:
                                manifest_data['defaultTab'] = ""
                else:
                    error_message = f"DAX query file not found: {query_name}.dax"
                    self.log_callback(f"    ⚠️ {error_message}")

            except Exception as e:
                error_message = str(e)
                self.log_callback(f"    ❌ Error removing DAX query {query_name}: {error_message}")

            results.append(RemovalResult(
                item_name=query_name,
                item_type='dax_query',
                success=success,
                error_message=error_message,
                bytes_freed=bytes_freed
            ))

        # Save updated manifest
        if manifest_data and manifest_path.exists():
            try:
                with open(manifest_path, 'w', encoding='utf-8') as f:
                    json.dump(manifest_data, f, indent=2)
                self.log_callback("  ✅ Updated DAX queries manifest")
            except Exception as e:
                self.log_callback(f"  ⚠️ Could not update DAX queries manifest: {e}")

        return results

    def _remove_tmdl_scripts(self, semantic_model_dir: Path, scripts_to_remove: List[CleanupOpportunity]) -> List[RemovalResult]:
        """Remove TMDL script files and update the manifest"""
        results = []

        tmdl_scripts_dir = semantic_model_dir / "TMDLScripts"
        manifest_path = tmdl_scripts_dir / ".pbi" / "tmdlScripts.json"

        if not tmdl_scripts_dir.exists():
            self.log_callback("  ⚠️ TMDLScripts directory not found")
            return results

        # Load the manifest to update tab order
        manifest_data = None
        if manifest_path.exists():
            try:
                with open(manifest_path, 'r', encoding='utf-8') as f:
                    manifest_data = json.load(f)
            except Exception as e:
                self.log_callback(f"  ⚠️ Could not read TMDL scripts manifest: {e}")

        for script_opportunity in scripts_to_remove:
            script_name = script_opportunity.item_name

            self.log_callback(f"  🗑️ Removing TMDL script: {script_name}")

            bytes_freed = 0
            success = False
            error_message = ""

            try:
                # Find and remove the .tmdl file
                script_file = tmdl_scripts_dir / f"{script_name}.tmdl"

                if script_file.exists():
                    bytes_freed = script_file.stat().st_size
                    script_file.unlink()
                    self.log_callback(f"    ✅ Removed TMDL file: {script_name}.tmdl ({self._format_bytes(bytes_freed)})")
                    success = True

                    # Update manifest if loaded
                    if manifest_data and 'tabOrder' in manifest_data:
                        if script_name in manifest_data['tabOrder']:
                            manifest_data['tabOrder'].remove(script_name)
                            self.log_callback(f"    ✅ Removed from manifest tabOrder")

                        # Update defaultTab if it was this script
                        if manifest_data.get('defaultTab') == script_name:
                            if manifest_data['tabOrder']:
                                manifest_data['defaultTab'] = manifest_data['tabOrder'][0]
                            else:
                                manifest_data['defaultTab'] = ""
                else:
                    error_message = f"TMDL script file not found: {script_name}.tmdl"
                    self.log_callback(f"    ⚠️ {error_message}")

            except Exception as e:
                error_message = str(e)
                self.log_callback(f"    ❌ Error removing TMDL script {script_name}: {error_message}")

            results.append(RemovalResult(
                item_name=script_name,
                item_type='tmdl_script',
                success=success,
                error_message=error_message,
                bytes_freed=bytes_freed
            ))

        # Save updated manifest
        if manifest_data and manifest_path.exists():
            try:
                with open(manifest_path, 'w', encoding='utf-8') as f:
                    json.dump(manifest_data, f, indent=2)
                self.log_callback("  ✅ Updated TMDL scripts manifest")
            except Exception as e:
                self.log_callback(f"  ⚠️ Could not update TMDL scripts manifest: {e}")

        return results

    def _find_image_display_name(self, report_dir: Path, item_name: str) -> str:
        """
        Find the display name currently used for an image in the report.

        Searches page.json files for references to the given ItemName and returns
        the associated display name. Falls back to the ItemName if no display name found.

        Args:
            report_dir: Path to the .Report directory
            item_name: The ItemName (actual filename) to search for

        Returns:
            The display name found, or item_name if not found
        """
        pages_dir = report_dir / "definition" / "pages"
        if not pages_dir.exists():
            return item_name

        def find_display_name_recursive(data: Any) -> str:
            """Recursively search for display name associated with item_name."""
            if isinstance(data, dict):
                # Check if this is an image container with both 'name' and 'url' keys
                if 'name' in data and 'url' in data:
                    url_resource = data.get('url', {}).get('expr', {}).get('ResourcePackageItem', {})
                    if url_resource.get('ItemName') == item_name:
                        # Found a match - extract display name
                        name_value = data.get('name', {}).get('expr', {}).get('Literal', {}).get('Value', '')
                        if name_value:
                            # Remove surrounding quotes if present
                            return name_value.strip("'")
                        return item_name

                # Recurse into dict values
                for value in data.values():
                    result = find_display_name_recursive(value)
                    if result:
                        return result

            elif isinstance(data, list):
                for item in data:
                    result = find_display_name_recursive(item)
                    if result:
                        return result

            return None

        # Search all page.json files
        for page_json in pages_dir.rglob("page.json"):
            try:
                with open(page_json, 'r', encoding='utf-8') as f:
                    page_data = json.load(f)
                result = find_display_name_recursive(page_data)
                if result:
                    return result
            except Exception:
                pass

        # Search all visual.json files as fallback
        for visual_json in pages_dir.rglob("visual.json"):
            try:
                with open(visual_json, 'r', encoding='utf-8') as f:
                    visual_data = json.load(f)
                result = find_display_name_recursive(visual_data)
                if result:
                    return result
            except Exception:
                pass

        return item_name

    def _consolidate_duplicate_images(self, report_dir: Path, groups: List[DuplicateImageGroup],
                                       selections: Dict[str, str]) -> List[RemovalResult]:
        """
        Consolidate duplicate images by redirecting references and removing duplicates.

        CRITICAL: Order of operations matters!
        1. Find keeper's display name from existing references
        2. Update ALL references in visual.json files (with consistent display name)
        3. Update ALL references in page.json files (with consistent display name)
        4. Delete duplicate image files AFTER refs are updated
        5. Update report.json to remove deleted image entries

        Args:
            report_dir: Path to the .Report directory
            groups: List of DuplicateImageGroup objects
            selections: Dict mapping group_id to selected keeper image name

        Returns:
            List of RemovalResult objects
        """
        results = []
        images_dir = report_dir / "StaticResources" / "RegisteredResources"

        for group in groups:
            # Get the keeper image (user selection or auto-selected)
            keeper = selections.get(group.group_id, group.selected_image)

            # Find the display name to use for ALL references to this keeper
            # This ensures all pages using the same image have the same display name
            keeper_display_name = self._find_image_display_name(report_dir, keeper)
            self.log_callback(f"  📦 Processing duplicate group (keeper: {keeper}, display: {keeper_display_name})")

            for img in group.images:
                img_name = img['name']

                # Skip the keeper image
                if img_name == keeper:
                    self.log_callback(f"    ✅ Keeping: {img_name}")
                    continue

                self.log_callback(f"    🔄 Consolidating: {img_name} -> {keeper}")

                bytes_freed = 0
                success = False
                error_message = ""
                references_updated = 0

                try:
                    # STEP 1: Update references in visual.json files (with consistent display name)
                    visual_refs = self._update_image_references_in_visuals(report_dir, img_name, keeper, keeper_display_name)
                    references_updated += visual_refs
                    if visual_refs > 0:
                        self.log_callback(f"      ✅ Updated {visual_refs} visual references")

                    # STEP 2: Update references in page.json files (with consistent display name)
                    page_refs = self._update_image_references_in_pages(report_dir, img_name, keeper, keeper_display_name)
                    references_updated += page_refs
                    if page_refs > 0:
                        self.log_callback(f"      ✅ Updated {page_refs} page references")

                    # STEP 3: Delete the duplicate image file
                    img_path = images_dir / img_name
                    if img_path.exists():
                        bytes_freed = img_path.stat().st_size
                        img_path.unlink()
                        self.log_callback(f"      ✅ Deleted file: {img_name} ({self._format_bytes(bytes_freed)})")

                    # STEP 4: Remove from report.json
                    self._remove_image_from_report_json(report_dir, img_name)

                    success = True

                except Exception as e:
                    error_message = str(e)
                    self.log_callback(f"      ❌ Error consolidating {img_name}: {error_message}")

                results.append(RemovalResult(
                    item_name=img_name,
                    item_type='duplicate_image',
                    success=success,
                    error_message=error_message,
                    bytes_freed=bytes_freed,
                    references_updated=references_updated
                ))

        return results

    def _update_image_references_in_visuals(self, report_dir: Path, old_name: str, new_name: str, new_display_name: str = None) -> int:
        """
        Update image references in all visual.json files.

        References are found at:
        visual.objects.general[].properties.imageUrl.expr.ResourcePackageItem.ItemName

        Args:
            report_dir: Path to the .Report directory
            old_name: Image name to find
            new_name: Image name to replace with
            new_display_name: Optional display name to set (for consolidation scenarios)

        Returns:
            Count of references updated
        """
        references_updated = 0
        pages_dir = report_dir / "definition" / "pages"

        if not pages_dir.exists():
            return 0

        # Find all visual.json files
        for visual_json in pages_dir.rglob("visual.json"):
            try:
                with open(visual_json, 'r', encoding='utf-8') as f:
                    visual_data = json.load(f)

                # Track if this file was modified
                modified = False

                # Recursively search and replace image references
                modified, count = self._replace_image_name_in_json(visual_data, old_name, new_name, new_display_name)

                if modified:
                    with open(visual_json, 'w', encoding='utf-8') as f:
                        json.dump(visual_data, f, indent=2)
                    references_updated += count

            except Exception as e:
                self.log_callback(f"      ⚠️ Error processing {visual_json}: {e}")

        return references_updated

    def _update_image_references_in_pages(self, report_dir: Path, old_name: str, new_name: str, new_display_name: str = None) -> int:
        """
        Update image references in all page.json files.

        References are found at:
        objects.outspace[].properties.image.image.url.expr.ResourcePackageItem.ItemName
        objects.background[].properties.image.image.url.expr.ResourcePackageItem.ItemName

        Args:
            report_dir: Path to the .Report directory
            old_name: Image name to find
            new_name: Image name to replace with
            new_display_name: Optional display name to set (for consolidation scenarios)

        Returns:
            Count of references updated
        """
        references_updated = 0
        pages_dir = report_dir / "definition" / "pages"

        if not pages_dir.exists():
            return 0

        # Find all page.json files
        for page_json in pages_dir.rglob("page.json"):
            try:
                with open(page_json, 'r', encoding='utf-8') as f:
                    page_data = json.load(f)

                # Track if this file was modified
                modified, count = self._replace_image_name_in_json(page_data, old_name, new_name, new_display_name)

                if modified:
                    with open(page_json, 'w', encoding='utf-8') as f:
                        json.dump(page_data, f, indent=2)
                    references_updated += count

            except Exception as e:
                self.log_callback(f"      ⚠️ Error processing {page_json}: {e}")

        return references_updated

    def _replace_image_name_in_json(self, data: Any, old_name: str, new_name: str, new_display_name: str = None) -> tuple:
        """
        Recursively search and replace image name references in JSON data.

        Updates ResourcePackageItem.ItemName (file reference).
        When new_display_name is provided, also updates the sibling display name
        (name.expr.Literal.Value) to ensure all pages using the same image have
        consistent display names - required to avoid Power BI rendering bugs when
        navigating between consecutive pages sharing the same wallpaper/background.

        Args:
            data: JSON data (dict, list, or primitive)
            old_name: Image name to find
            new_name: Image name to replace with
            new_display_name: Optional display name to set (for consolidation scenarios)

        Returns:
            Tuple of (was_modified, count_of_replacements)
        """
        modified = False
        count = 0

        if isinstance(data, dict):
            # Check if this is an image container with both 'name' and 'url' keys
            # This is the image.image level in page.json wallpaper/background structure
            if 'name' in data and 'url' in data:
                url_resource = data.get('url', {}).get('expr', {}).get('ResourcePackageItem', {})
                if url_resource.get('ItemName') == old_name:
                    # Update the ItemName
                    url_resource['ItemName'] = new_name
                    modified = True
                    count = 1
                    # Also update display name if provided (for consolidation)
                    # This ensures all pages using the same image have the same display name
                    if new_display_name:
                        name_literal = data.get('name', {}).get('expr', {}).get('Literal', {})
                        if 'Value' in name_literal:
                            name_literal['Value'] = f"'{new_display_name}'"

            # Check direct ResourcePackageItem (for visual.json and other patterns)
            if not modified and data.get('ResourcePackageItem', {}).get('ItemName') == old_name:
                data['ResourcePackageItem']['ItemName'] = new_name
                modified = True
                count = 1

            # Continue recursion if no modification at this level
            if not modified:
                for key, value in data.items():
                    child_modified, child_count = self._replace_image_name_in_json(value, old_name, new_name, new_display_name)
                    if child_modified:
                        modified = True
                        count += child_count

        elif isinstance(data, list):
            for item in data:
                child_modified, child_count = self._replace_image_name_in_json(item, old_name, new_name, new_display_name)
                if child_modified:
                    modified = True
                    count += child_count

        return modified, count

    def _remove_unused_images(self, report_dir: Path, images_to_remove: List[CleanupOpportunity]) -> List[RemovalResult]:
        """
        Remove unused images from the report.

        Args:
            report_dir: Path to the .Report directory
            images_to_remove: List of CleanupOpportunity objects for unused images

        Returns:
            List of RemovalResult objects
        """
        results = []
        images_dir = report_dir / "StaticResources" / "RegisteredResources"

        for img_opportunity in images_to_remove:
            img_name = img_opportunity.item_name
            is_orphan = getattr(img_opportunity, 'is_orphan', False)

            orphan_label = " (orphan file)" if is_orphan else ""
            self.log_callback(f"  🗑️ Removing unused image: {img_name}{orphan_label}")

            bytes_freed = 0
            success = False
            error_message = ""

            try:
                # Delete the image file
                img_path = images_dir / img_name
                if img_path.exists():
                    bytes_freed = img_path.stat().st_size
                    img_path.unlink()
                    self.log_callback(f"    ✅ Deleted file: {img_name} ({self._format_bytes(bytes_freed)})")
                else:
                    self.log_callback(f"    ⚠️ File not found: {img_name}")

                # Remove from report.json (only if NOT an orphan - orphans aren't registered)
                if not is_orphan:
                    self._remove_image_from_report_json(report_dir, img_name)
                else:
                    self.log_callback(f"    ✅ Skipped report.json update (orphan file not registered)")

                success = True

            except Exception as e:
                error_message = str(e)
                self.log_callback(f"    ❌ Error removing {img_name}: {error_message}")

            results.append(RemovalResult(
                item_name=img_name,
                item_type='unused_image',
                success=success,
                error_message=error_message,
                bytes_freed=bytes_freed
            ))

        return results

    def _remove_image_from_report_json(self, report_dir: Path, image_name: str) -> bool:
        """
        Remove an image entry from report.json resourcePackages.

        Args:
            report_dir: Path to the .Report directory
            image_name: Name of the image to remove

        Returns:
            True if entry was found and removed, False otherwise
        """
        report_json_path = report_dir / "definition" / "report.json"

        if not report_json_path.exists():
            return False

        try:
            with open(report_json_path, 'r', encoding='utf-8') as f:
                report_data = json.load(f)

            resource_packages = report_data.get('resourcePackages', [])
            entry_removed = False

            for package in resource_packages:
                package_type = package.get('type', '')

                # Only check RegisteredResources packages
                if package_type == 'RegisteredResources':
                    items = package.get('items', [])
                    original_count = len(items)

                    # Filter out the image
                    updated_items = [
                        item for item in items
                        if not (item.get('type') == 'Image' and item.get('name') == image_name)
                    ]

                    if len(updated_items) < original_count:
                        package['items'] = updated_items
                        entry_removed = True
                        self.log_callback(f"      ✅ Removed {image_name} from report.json")

            if entry_removed:
                with open(report_json_path, 'w', encoding='utf-8') as f:
                    json.dump(report_data, f, indent=2)

            return entry_removed

        except Exception as e:
            self.log_callback(f"      ⚠️ Error updating report.json: {e}")
            return False
