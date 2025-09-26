"""
Report Cleanup Engine - Engine for removing unused themes and custom visuals
Built by Reid Havens of Analytic Endeavors

This module handles the actual removal of unused themes and custom visuals from Power BI reports.
"""

import json
import shutil
import os
import stat
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable
import datetime

# Import shared data classes
from tools.report_cleanup.shared_types import CleanupOpportunity, RemovalResult


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
                           hide_visual_filters: bool = False, create_backup: bool = True) -> List[RemovalResult]:
        """
        Remove unused themes, custom visuals, and/or bookmarks from a PBIP report
        
        Args:
            pbip_path: Path to the .pbip file
            themes_to_remove: List of theme names to remove (None to skip)
            visuals_to_remove: List of CleanupOpportunity objects for visuals to remove (None to skip)
            bookmarks_to_remove: List of CleanupOpportunity objects for bookmarks to remove (None to skip)
            hide_visual_filters: Whether to hide all visual-level filters
            create_backup: Whether to create a backup before making changes
            
        Returns:
            List of RemovalResult objects
        """
        results = []
        
        pbip_file = Path(pbip_path)
        report_dir = pbip_file.parent / f"{pbip_file.stem}.Report"
        
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
        
        # Clean up resourcePackages, publicCustomVisuals, and bookmarks in report.json
        if themes_to_remove or visuals_to_remove or bookmarks_to_remove:
            self._cleanup_report_references(report_dir, themes_to_remove or [], visuals_to_remove or [], bookmarks_to_remove or [])
        
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
