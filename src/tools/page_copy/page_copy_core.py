"""
Advanced Page Copy Core Logic
Built by Reid Havens of Analytic Endeavors

Leverages existing merger_core logic for copying pages and bookmarks within the same report.
"""

import json
import shutil
import uuid
from pathlib import Path
from typing import Dict, List, Optional, Union, Tuple, Any, Callable

from tools.report_merger.merger_core import (
    ValidationService, MergerEngine, PBIPMergerError, 
    InvalidReportError, ValidationError, FileOperationError
)
from core.constants import AppConstants

class PageCopyError(PBIPMergerError):
    """Raised when page copy operations fail."""
    pass

class AdvancedPageCopyEngine:
    """
    Advanced Page Copy Engine - leverages MergerEngine logic for copying pages within a report
    """
    
    def __init__(self, logger_callback: Optional[Callable[[str], None]] = None):
        self.log_callback = logger_callback or self._default_log
        self.validation_service = ValidationService()
        self.merger_engine = MergerEngine(logger_callback=self.log_callback)
        self._selected_pages = []
        
        # CRITICAL: Track page ID mappings and bookmark order
        self._page_id_mapping = {}  # old_page_id -> new_page_id
        self._original_bookmark_names = []  # Track original bookmarks for ordering
        self._copied_bookmarks_order = []  # Track copied bookmarks in their creation order
        self._bookmark_copy_mapping = {}  # Maps copied bookmark ID to original bookmark ID
    
    def _default_log(self, message: str) -> None:
        print(message)
    
    def analyze_report_pages(self, report_path: str) -> Dict[str, Any]:
        """
        Analyze report to find pages that have bookmarks
        Returns analysis results including pages with bookmarks
        """
        self.log_callback("🔍 Analyzing report pages for bookmark content...")
        
        # Clean and validate path
        clean_path = self.merger_engine.clean_path(report_path)
        
        # Validate input
        self.validation_service._validate_single_pbip_path(clean_path, "Report")
        
        # Get report directory
        report_dir = Path(clean_path).parent / f"{Path(clean_path).stem}.Report"
        
        # Validate structure
        self.validation_service.validate_thin_report_structure(report_dir, Path(clean_path).stem)
        
        self.log_callback("✅ Report structure validated")
        
        # Check if source report has schema issues
        self._check_source_report_schemas(report_dir)
        
        # Analyze pages and their bookmarks
        pages_analysis = self._analyze_pages_with_bookmarks(report_dir)
        
        # Get report-level bookmarks
        report_bookmarks = self._get_report_level_bookmarks(report_dir)
        
        # Count total elements
        total_pages = len(pages_analysis['pages_with_bookmarks']) + len(pages_analysis['pages_without_bookmarks'])
        total_bookmarks = sum(page['bookmark_count'] for page in pages_analysis['pages_with_bookmarks'])
        total_bookmarks += len(report_bookmarks)
        
        self.log_callback(f"   📄 Total Pages: {total_pages}")
        self.log_callback(f"   🔖 Pages with Bookmarks: {len(pages_analysis['pages_with_bookmarks'])}")
        self.log_callback(f"   📋 Total Bookmarks: {total_bookmarks}")
        
        return {
            'report': {
                'path': clean_path,
                'name': Path(clean_path).stem,
                'total_pages': total_pages,
                'total_bookmarks': total_bookmarks
            },
            'pages_with_bookmarks': pages_analysis['pages_with_bookmarks'],
            'pages_without_bookmarks': pages_analysis['pages_without_bookmarks'],
            'report_bookmarks': report_bookmarks,
            'analysis_summary': {
                'copyable_pages': len(pages_analysis['pages_with_bookmarks']),
                'total_pages': total_pages,
                'total_bookmarks': total_bookmarks
            }
        }
    
    def copy_selected_pages(self, report_path: str, selected_page_names: List[str], 
                           analysis_results: Dict[str, Any]) -> bool:
        """
        Copy selected pages within the same report, including their bookmarks
        """
        try:
            self.log_callback("🚀 Starting advanced page copy operation...")
            
            clean_path = self.merger_engine.clean_path(report_path)
            report_dir = Path(clean_path).parent / f"{Path(clean_path).stem}.Report"
            
            # Validate selections
            if not selected_page_names:
                raise PageCopyError("No pages selected for copying")
            
            # Execute copy operations
            copy_stats = self._execute_page_copy_operations(
                report_dir, selected_page_names, analysis_results
            )
            
            # Update report metadata
            self._update_report_metadata_after_copy(report_dir, copy_stats)
            
            # Validate and fix any schema issues
            self.validate_and_fix_schemas(report_dir)
            
            # FINAL SAFEGUARD: One more check and fix
            self._final_schema_safeguard(report_dir)
            
            self.log_callback("🎉 PAGE COPY COMPLETED SUCCESSFULLY!")
            self.log_callback(f"📊 Copied {copy_stats['pages_copied']} pages with {copy_stats['bookmarks_copied']} bookmarks")
            
            return True
            
        except Exception as e:
            self.log_callback(f"❌ Page copy operation failed: {e}")
            return False
    
    def _analyze_pages_with_bookmarks(self, report_dir: Path) -> Dict[str, List[Dict]]:
        """Analyze pages to find which ones have associated bookmarks"""
        pages_dir = report_dir / "definition" / "pages"
        bookmarks_dir = report_dir / "definition" / "bookmarks"
        
        pages_with_bookmarks = []
        pages_without_bookmarks = []
        
        if not pages_dir.exists():
            return {
                'pages_with_bookmarks': pages_with_bookmarks,
                'pages_without_bookmarks': pages_without_bookmarks
            }
        
        # Get all bookmark files and their page references
        page_bookmark_map = self._create_page_bookmark_map(bookmarks_dir)
        
        # Analyze each page
        for page_dir in pages_dir.iterdir():
            if page_dir.is_dir() and page_dir.name != "pages.json":
                page_info = self._analyze_single_page(page_dir, page_bookmark_map)
                
                if page_info['bookmark_count'] > 0:
                    pages_with_bookmarks.append(page_info)
                else:
                    pages_without_bookmarks.append(page_info)
        
        return {
            'pages_with_bookmarks': pages_with_bookmarks,
            'pages_without_bookmarks': pages_without_bookmarks
        }
    
    def _create_page_bookmark_map(self, bookmarks_dir: Path) -> Dict[str, List[str]]:
        """Create mapping of page names to their bookmark files"""
        page_bookmark_map = {}
        
        if not bookmarks_dir.exists():
            return page_bookmark_map
        
        for bookmark_file in bookmarks_dir.glob("*.bookmark.json"):
            try:
                with open(bookmark_file, 'r', encoding='utf-8') as f:
                    bookmark_data = json.load(f)
                
                # Extract page reference from bookmark
                page_name = self._extract_page_name_from_bookmark(bookmark_data)
                
                if page_name:
                    if page_name not in page_bookmark_map:
                        page_bookmark_map[page_name] = []
                    page_bookmark_map[page_name].append(bookmark_file.stem.replace('.bookmark', ''))
                
            except Exception as e:
                self.log_callback(f"     ⚠️ Warning: Could not read bookmark {bookmark_file.name}: {e}")
        
        return page_bookmark_map
    
    def _extract_page_name_from_bookmark(self, bookmark_data: Dict) -> Optional[str]:
        """Extract page name from bookmark data using same logic as merger"""
        # Primary method: Look for activeSection in explorationState (this is the correct approach)
        if 'explorationState' in bookmark_data:
            exploration = bookmark_data['explorationState']
            if 'activeSection' in exploration:
                return exploration['activeSection']  # This is the page directory name
        
        # Fallback methods (from original logic)
        if 'config' in bookmark_data:
            config = bookmark_data['config']
            if 'name' in config:
                return config['name']
        
        # Look for displayName as fallback
        if 'displayName' in bookmark_data:
            return bookmark_data['displayName']
        
        return None
    
    def _analyze_single_page(self, page_dir: Path, page_bookmark_map: Dict[str, List[str]]) -> Dict[str, Any]:
        """Analyze a single page for bookmark content"""
        page_info = {
            'name': page_dir.name,  # This is the directory name (like 6a937bf7bb4cc8aa9829)
            'display_name': page_dir.name,  # Default to directory name
            'bookmark_count': 0,
            'bookmark_names': [],
            'page_path': page_dir
        }
        
        # Try to get display name from page.json
        page_json_file = page_dir / "page.json"
        if page_json_file.exists():
            try:
                with open(page_json_file, 'r', encoding='utf-8') as f:
                    page_data = json.load(f)
                
                # Use displayName for user-friendly display
                if 'displayName' in page_data:
                    page_info['display_name'] = page_data['displayName']
                else:
                    # Fallback to name if no displayName
                    page_info['display_name'] = page_data.get('name', page_dir.name)
                    
            except Exception as e:
                self.log_callback(f"     ⚠️ Warning: Could not read page.json for {page_dir.name}: {e}")
        
        # Check for associated bookmarks using the directory name (not display name)
        page_directory_name = page_dir.name  # This should match activeSection in bookmarks
        if page_directory_name in page_bookmark_map:
            page_info['bookmark_count'] = len(page_bookmark_map[page_directory_name])
            page_info['bookmark_names'] = page_bookmark_map[page_directory_name]
        
        return page_info
    
    def _get_report_level_bookmarks(self, report_dir: Path) -> List[Dict[str, Any]]:
        """Get report-level bookmarks (not tied to specific pages)"""
        bookmarks_dir = report_dir / "definition" / "bookmarks"
        report_bookmarks = []
        
        if not bookmarks_dir.exists():
            return report_bookmarks
        
        for bookmark_file in bookmarks_dir.glob("*.bookmark.json"):
            try:
                with open(bookmark_file, 'r', encoding='utf-8') as f:
                    bookmark_data = json.load(f)
                
                # Check if this is a report-level bookmark (no specific page reference)
                page_name = self._extract_page_name_from_bookmark(bookmark_data)
                
                if not page_name:  # Report-level bookmark
                    report_bookmarks.append({
                        'name': bookmark_file.stem.replace('.bookmark', ''),
                        'display_name': bookmark_data.get('displayName', bookmark_file.stem),
                        'file_path': bookmark_file
                    })
                    
            except Exception as e:
                self.log_callback(f"     ⚠️ Warning: Could not analyze bookmark {bookmark_file.name}: {e}")
        
        return report_bookmarks
    
    def _execute_page_copy_operations(self, report_dir: Path, selected_page_names: List[str], 
                                    analysis_results: Dict[str, Any]) -> Dict[str, int]:
        """Execute the actual page copying operations"""
        copy_stats = {'pages_copied': 0, 'bookmarks_copied': 0}
        
        pages_dir = report_dir / "definition" / "pages"
        bookmarks_dir = report_dir / "definition" / "bookmarks"
        
        # IMPORTANT: Capture original bookmark names BEFORE copying
        if bookmarks_dir.exists():
            # Read existing bookmarks.json to preserve exact order
            bookmarks_json = bookmarks_dir / "bookmarks.json"
            if bookmarks_json.exists():
                try:
                    with open(bookmarks_json, 'r', encoding='utf-8') as f:
                        bookmarks_data = json.load(f)
                    if 'items' in bookmarks_data:
                        # Extract all bookmark names, including those in groups
                        all_original_bookmarks = []
                        for item in bookmarks_data['items']:
                            if 'children' in item:
                                # This is a group - add all its children
                                all_original_bookmarks.extend(item['children'])
                            else:
                                # This is a regular bookmark
                                all_original_bookmarks.append(item['name'])
                        self._original_bookmark_names = all_original_bookmarks
                except:
                    # Fallback to file order
                    self._original_bookmark_names = [bf.stem.replace('.bookmark', '') 
                                                   for bf in sorted(bookmarks_dir.glob("*.bookmark.json"))]
            else:
                self._original_bookmark_names = [bf.stem.replace('.bookmark', '') 
                                               for bf in sorted(bookmarks_dir.glob("*.bookmark.json"))]
        
        # Reset copied bookmarks order for this operation
        self._copied_bookmarks_order = []
        self._bookmark_copy_mapping = {}
        
        # Find selected pages in analysis results
        pages_with_bookmarks = analysis_results['pages_with_bookmarks']
        selected_pages_data = [p for p in pages_with_bookmarks if p['name'] in selected_page_names]
        
        for page_data in selected_pages_data:
            try:
                # Copy the page
                new_page_name = self._copy_single_page(pages_dir, page_data)
                copy_stats['pages_copied'] += 1
                
                # Copy associated bookmarks
                bookmarks_copied = self._copy_page_bookmarks(
                    bookmarks_dir, page_data, new_page_name
                )
                copy_stats['bookmarks_copied'] += bookmarks_copied
                
                self.log_callback(f"   ✅ Copied page '{page_data['display_name']}' with {bookmarks_copied} bookmarks")
                
            except Exception as e:
                self.log_callback(f"   ❌ Failed to copy page '{page_data['name']}': {e}")
        
        return copy_stats
    
    def _copy_single_page(self, pages_dir: Path, page_data: Dict[str, Any]) -> str:
        """Copy a single page directory and generate unique names"""
        source_page_dir = page_data['page_path']
        old_page_id = source_page_dir.name  # Store the original page ID
        
        # Generate a completely new unique page ID in Power BI format (xxxxxxxxxxxxxxxxxxxxxxxx)
        # Format: 8chars + 4chars + 4chars + 4chars (no hyphens)
        new_page_id = str(uuid.uuid4()).replace('-', '')[:20]  # Take first 20 chars for consistency
        
        # CRITICAL: Store the mapping for bookmark updates
        self._page_id_mapping[old_page_id] = new_page_id
        
        # Copy page directory with new ID
        target_page_dir = pages_dir / new_page_id
        shutil.copytree(source_page_dir, target_page_dir)
        
        # Update page.json with new identifiers
        page_json_file = target_page_dir / "page.json"
        if page_json_file.exists():
            try:
                with open(page_json_file, 'r', encoding='utf-8') as f:
                    page_data_json = json.load(f)
                
                # Generate unique display name
                original_display_name = page_data_json.get('displayName', 'Page')
                counter = 1
                new_display_name = f"{original_display_name} (Copy)"
                
                # Check if this display name already exists
                while self._display_name_exists(pages_dir, new_display_name, new_page_id):
                    counter += 1
                    new_display_name = f"{original_display_name} (Copy) {counter}"
                
                # Update page metadata
                page_data_json['name'] = new_page_id  # Internal name matches directory
                page_data_json['displayName'] = new_display_name
                
                # CRITICAL: Ensure 'section' matches 'name' (Power BI requirement)
                if 'section' in page_data_json:
                    page_data_json['section'] = new_page_id
                    self.log_callback(f"     📝 Updated section to match name: {new_page_id}")
                
                # Generate new unique identifiers
                if 'id' in page_data_json:
                    page_data_json['id'] = str(uuid.uuid4()).replace('-', '')[:20]
                
                # Remove any non-standard properties
                non_standard_props = ['originalName', '_needsPageMapping']
                for prop in non_standard_props:
                    if prop in page_data_json:
                        del page_data_json[prop]
                
                with open(page_json_file, 'w', encoding='utf-8') as f:
                    json.dump(page_data_json, f, indent=2)
                
                self.log_callback(f"     📝 Created page: {new_display_name} (ID: {new_page_id})")
                    
            except Exception as e:
                self.log_callback(f"     ⚠️ Warning: Could not update page.json: {e}")
        
        return new_page_id
    
    def _copy_page_bookmarks(self, bookmarks_dir: Path, page_data: Dict[str, Any], 
                           new_page_id: str) -> int:
        """Copy bookmarks associated with a page with completely new IDs"""
        if not page_data['bookmark_names']:
            return 0
        
        bookmarks_copied = 0
        # Track the new bookmark names in the same order as originals
        new_bookmark_names_ordered = []
        
        # IMPORTANT: Copy bookmarks in the same order they appear in the original page
        for bookmark_name in page_data['bookmark_names']:
            try:
                source_bookmark_file = bookmarks_dir / f"{bookmark_name}.bookmark.json"
                
                if not source_bookmark_file.exists():
                    continue
                
                # Generate completely new bookmark ID in Power BI format
                new_bookmark_id = str(uuid.uuid4()).replace('-', '')[:20]
                
                target_bookmark_file = bookmarks_dir / f"{new_bookmark_id}.bookmark.json"
                
                # Copy and update bookmark
                with open(source_bookmark_file, 'r', encoding='utf-8') as f:
                    bookmark_data = json.load(f)
                
                # Update bookmark to reference new page
                self._update_bookmark_page_reference(bookmark_data, new_page_id)
                
                # Generate unique bookmark display name
                original_display_name = bookmark_data.get('displayName', 'Bookmark')
                counter = 1
                new_display_name = f"{original_display_name} (Copy)"
                
                # Check if this display name already exists
                while self._bookmark_display_name_exists(bookmarks_dir, new_display_name, new_bookmark_id):
                    counter += 1
                    new_display_name = f"{original_display_name} (Copy) {counter}"
                
                bookmark_data['displayName'] = new_display_name
                
                # Update bookmark IDs
                bookmark_data['name'] = new_bookmark_id
                if 'id' in bookmark_data:
                    bookmark_data['id'] = new_bookmark_id
                
                # Remove any non-standard properties
                non_standard_props = ['originalName', '_needsPageMapping']
                for prop in non_standard_props:
                    if prop in bookmark_data:
                        del bookmark_data[prop]
                
                with open(target_bookmark_file, 'w', encoding='utf-8') as f:
                    json.dump(bookmark_data, f, indent=2)
                
                bookmarks_copied += 1
                new_bookmark_names_ordered.append(new_bookmark_id)
                # Track the mapping of copied bookmark to original
                self._bookmark_copy_mapping[new_bookmark_id] = bookmark_name
                self.log_callback(f"       🔖 Created bookmark: {new_display_name} (ID: {new_bookmark_id})")
                
            except Exception as e:
                self.log_callback(f"     ⚠️ Warning: Could not copy bookmark {bookmark_name}: {e}")
        
        # Store the mapping of original bookmark order to new bookmark IDs
        if not hasattr(self, '_copied_bookmarks_order'):
            self._copied_bookmarks_order = []
        self._copied_bookmarks_order.extend(new_bookmark_names_ordered)
        
        return bookmarks_copied
    
    def _update_bookmark_page_reference(self, bookmark_data: Dict, new_page_name: str):
        """Update bookmark data to reference the new page"""
        # Update page reference in explorationState - this is the critical one for Power BI
        if 'explorationState' in bookmark_data:
            exploration = bookmark_data['explorationState']
            if 'activeSection' in exploration:
                old_page_ref = exploration['activeSection']
                # Use the mapping if this is a copied page
                if old_page_ref in self._page_id_mapping:
                    new_mapped_page_id = self._page_id_mapping[old_page_ref]
                    exploration['activeSection'] = new_mapped_page_id
                    self.log_callback(f"       🔄 Updated bookmark page reference: {old_page_ref} → {new_mapped_page_id}")
                    
                    # Update 'sections' object inside explorationState - it's a dict with page IDs as keys
                    if 'sections' in exploration and isinstance(exploration['sections'], dict):
                        old_sections = exploration['sections']
                        new_sections = {}
                        
                        # Update the page ID keys in the sections object
                        for section_key, section_data in old_sections.items():
                            if section_key == old_page_ref:
                                # Replace old page ID key with new one
                                new_sections[new_mapped_page_id] = section_data
                            else:
                                # Keep other sections as-is
                                new_sections[section_key] = section_data
                        
                        exploration['sections'] = new_sections
                        self.log_callback(f"       🔄 Updated sections object: {old_page_ref} → {new_mapped_page_id}")
                else:
                    # For new bookmarks being created
                    exploration['activeSection'] = new_page_name
                    
                    # Update 'sections' object inside explorationState
                    if 'sections' in exploration and isinstance(exploration['sections'], dict):
                        old_sections = exploration['sections']
                        new_sections = {}
                        
                        # Look for any sections that need updating
                        for section_key, section_data in old_sections.items():
                            # Since this is a new bookmark, we just preserve the structure
                            # but update any keys that match the old page reference
                            if section_key == old_page_ref:
                                new_sections[new_page_name] = section_data
                            else:
                                new_sections[section_key] = section_data
                        
                        exploration['sections'] = new_sections
            

        
        # Also update other references
        if 'config' in bookmark_data and 'name' in bookmark_data['config']:
            bookmark_data['config']['name'] = new_page_name
        
        if 'explorationState' in bookmark_data:
            exploration = bookmark_data['explorationState']
            if 'activeReportPage' in exploration:
                exploration['activeReportPage'] = new_page_name
    
    def _update_report_metadata_after_copy(self, report_dir: Path, copy_stats: Dict[str, int]):
        """Update report metadata files after copying pages"""
        # Update pages.json
        self._update_pages_json_after_copy(report_dir / "definition" / "pages")
        
        # Update bookmarks.json
        self._update_bookmarks_json_after_copy(report_dir / "definition" / "bookmarks")
        
        self.log_callback(f"   ✅ Report metadata updated")
    
    def _update_pages_json_after_copy(self, pages_dir: Path):
        """Update pages.json after copying pages"""
        # Use the existing logic from merger_core
        self.merger_engine._rebuild_pages_json(pages_dir)
        
        # DEBUG: Log that we're using merger_engine
        self.log_callback("   🔧 Updated pages.json using merger_engine")
    
    def _update_bookmarks_json_after_copy(self, bookmarks_dir: Path):
        """Update bookmarks.json after copying bookmarks"""
        if not bookmarks_dir.exists():
            return
        
        bookmark_files = list(bookmarks_dir.glob("*.bookmark.json"))
        
        if not bookmark_files:
            # If no bookmarks exist, remove bookmarks.json if present
            bookmarks_json_file = bookmarks_dir / "bookmarks.json"
            if bookmarks_json_file.exists():
                bookmarks_json_file.unlink()
            return
        
        # CRITICAL: Maintain proper order - originals first, then copies
        all_bookmark_names = [bf.stem.replace(".bookmark", "") for bf in bookmark_files]
        
        # Read existing bookmarks.json to get original order and groups if it exists
        bookmarks_json_file = bookmarks_dir / "bookmarks.json"
        original_items = []
        original_groups = {}
        has_groups = False
        
        if bookmarks_json_file.exists():
            try:
                with open(bookmarks_json_file, 'r', encoding='utf-8') as f:
                    existing_data = json.load(f)
                if 'items' in existing_data:
                    # Check if we have groups (items with 'children')
                    for item in existing_data['items']:
                        if 'children' in item:
                            has_groups = True
                            # This is a group
                            group_name = item['name']
                            original_groups[group_name] = {
                                'name': group_name,
                                'displayName': item.get('displayName', group_name),
                                'children': item['children'],
                                'original_children': item['children'][:] # Keep a copy
                            }
                        else:
                            # Regular bookmark
                            if item['name'] in self._original_bookmark_names:
                                original_items.append(item)
            except Exception as e:
                self.log_callback(f"   ⚠️ Error reading bookmarks.json: {e}")
        
        # Build the final items list
        final_items = []
        
        if has_groups:
            # Handle grouped bookmarks
            self._handle_grouped_bookmarks(original_groups, final_items, all_bookmark_names)
        else:
            # Handle flat bookmarks (no groups)
            self._handle_flat_bookmarks(original_items, final_items, all_bookmark_names)
        
        # CRITICAL: Use the correct schema URL (bookmarksMetadata, not bookmarks!)
        bookmarks_data = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmarksMetadata/1.0.0/schema.json",
            "items": final_items
        }
        
        bookmarks_json_file = bookmarks_dir / "bookmarks.json"
        with open(bookmarks_json_file, 'w', encoding='utf-8') as f:
            json.dump(bookmarks_data, f, indent=2, ensure_ascii=False)
        
        self.log_callback(f"   ✅ Updated bookmarks.json with {len(bookmark_files)} bookmarks")
        self.log_callback(f"   📋 Bookmark order: {len(self._original_bookmark_names)} originals, {len(self._copied_bookmarks_order)} copies")
        
        # DEBUG: Verify what was written
        with open(bookmarks_json_file, 'r', encoding='utf-8') as f:
            written_content = f.read()
            self.log_callback(f"   🔍 DEBUG - bookmarks.json content: {written_content[:200]}...")
            if '"$schema": "1.0.0"' in written_content:
                self.log_callback("   ❌ ERROR: Schema is still wrong! Just '1.0.0' found!")
            elif '"$schema": "https://' in written_content:
                self.log_callback("   ✅ Schema appears correct (has full URL)")
            else:
                self.log_callback("   ⚠️ WARNING: No schema found in bookmarks.json!")
    
    def _handle_flat_bookmarks(self, original_items: List[Dict], final_items: List[Dict], all_bookmark_names: List[str]):
        """Handle bookmarks when there are no groups"""
        # Get original bookmark names in order
        original_order = [item['name'] for item in original_items] if original_items else self._original_bookmark_names
        
        # Add original bookmarks
        for bookmark_name in original_order:
            if bookmark_name in all_bookmark_names:
                final_items.append({"name": bookmark_name})
        
        # Add copied bookmarks in their creation order
        for bookmark_name in self._copied_bookmarks_order:
            if bookmark_name in all_bookmark_names:
                final_items.append({"name": bookmark_name})
    
    def _handle_grouped_bookmarks(self, original_groups: Dict, final_items: List[Dict], all_bookmark_names: List[str]):
        """Handle bookmarks when there are groups"""
        # First, add all original groups with their original children
        for group_name, group_data in original_groups.items():
            group_item = {
                "name": group_data['name'],
                "displayName": group_data['displayName'],
                "children": [child for child in group_data['children'] if child in all_bookmark_names]
            }
            if group_item['children']:  # Only add group if it has children
                final_items.append(group_item)
        
        # Now create copy groups for each original group that has copied bookmarks
        if self._copied_bookmarks_order:
            # Create a mapping of original bookmark to its group
            bookmark_to_group = {}
            for group_name, group_data in original_groups.items():
                for child in group_data['original_children']:
                    bookmark_to_group[child] = group_data
            
            # Group the copied bookmarks by their original group
            copied_groups = {}
            for copied_bookmark in self._copied_bookmarks_order:
                # Find which original bookmark this is a copy of
                original_bookmark = self._bookmark_copy_mapping.get(copied_bookmark)
                if original_bookmark and original_bookmark in bookmark_to_group:
                    group_data = bookmark_to_group[original_bookmark]
                    group_name = group_data['name']
                    if group_name not in copied_groups:
                        copied_groups[group_name] = {
                            'original_display_name': group_data['displayName'],
                            'children': []
                        }
                    copied_groups[group_name]['children'].append(copied_bookmark)
            
            # Add the copied groups
            for original_group_name, copied_group_data in copied_groups.items():
                # Generate unique group ID
                new_group_id = str(uuid.uuid4()).replace('-', '')[:20]
                copied_group_item = {
                    "name": new_group_id,
                    "displayName": f"{copied_group_data['original_display_name']} (Copy)",
                    "children": copied_group_data['children']
                }
                final_items.append(copied_group_item)
                self.log_callback(f"   📁 Created bookmark group: {copied_group_item['displayName']}")
    
    def _display_name_exists(self, pages_dir: Path, display_name: str, exclude_page_id: str) -> bool:
        """Check if a display name already exists in other pages"""
        for page_dir in pages_dir.iterdir():
            if page_dir.is_dir() and page_dir.name != "pages.json" and page_dir.name != exclude_page_id:
                page_json = page_dir / "page.json"
                if page_json.exists():
                    try:
                        with open(page_json, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        if data.get('displayName') == display_name:
                            return True
                    except:
                        pass
        return False
    
    def _bookmark_display_name_exists(self, bookmarks_dir: Path, display_name: str, exclude_bookmark_id: str) -> bool:
        """Check if a bookmark display name already exists"""
        for bookmark_file in bookmarks_dir.glob("*.bookmark.json"):
            if bookmark_file.stem != exclude_bookmark_id:
                try:
                    with open(bookmark_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    if data.get('displayName') == display_name:
                        return True
                except:
                    pass
        return False
    
    def get_pages_for_selection(self, analysis_results: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Get list of pages available for selection (only those with bookmarks)"""
        return analysis_results.get('pages_with_bookmarks', [])
    
    def validate_and_fix_schemas(self, report_dir: Path):
        """Validate and fix all schema references in the report after copy operations."""
        self.log_callback("   🔧 Validating and fixing ALL schema references...")
        
        # Force fix bookmarks.json regardless of current content
        bookmarks_json = report_dir / "definition" / "bookmarks" / "bookmarks.json"
        if bookmarks_json.exists():
            try:
                # Read current content
                with open(bookmarks_json, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # Log current schema
                current_schema = data.get("$schema", "MISSING")
                self.log_callback(f"     🔍 Current bookmarks.json schema: {current_schema}")
                
                # FORCE correct schema (bookmarksMetadata!)
                data["$schema"] = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmarksMetadata/1.0.0/schema.json"
                
                # Also ensure we have 'items' not 'bookmarks'
                if "bookmarks" in data and "items" not in data:
                    data["items"] = data["bookmarks"]
                    del data["bookmarks"]
                
                # Write back
                with open(bookmarks_json, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                    
                self.log_callback(f"     ✅ FORCED correct schema in bookmarks.json")
                
                # Verify the fix
                with open(bookmarks_json, 'r', encoding='utf-8') as f:
                    verify_content = f.read()
                    if '"$schema": "1.0.0"' in verify_content:
                        self.log_callback("     ❌ CRITICAL: Schema is STILL wrong after fix!")
                    else:
                        self.log_callback("     ✅ Verified: Schema is now correct")
                        
            except Exception as e:
                self.log_callback(f"     ❌ Error fixing bookmarks.json: {e}")
        
        # Also fix pages.json
        pages_json = report_dir / "definition" / "pages" / "pages.json"
        if pages_json.exists():
            try:
                with open(pages_json, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                data["$schema"] = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json"
                
                with open(pages_json, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                    
                self.log_callback(f"     ✅ Fixed schema in pages.json")
            except Exception as e:
                self.log_callback(f"     ⚠️ Could not fix pages.json: {e}")
    
    def _check_source_report_schemas(self, report_dir: Path):
        """Check if the source report has schema issues before copying."""
        self.log_callback("   🔍 Checking source report schemas...")
        
        bookmarks_json = report_dir / "definition" / "bookmarks" / "bookmarks.json"
        if bookmarks_json.exists():
            try:
                with open(bookmarks_json, 'r', encoding='utf-8') as f:
                    content = f.read()
                    if '"$schema": "1.0.0"' in content:
                        self.log_callback("   ⚠️ WARNING: Source report has incorrect schema in bookmarks.json!")
                        self.log_callback("   🔧 This will be fixed during the copy operation.")
                    else:
                        self.log_callback("   ✅ Source report bookmarks.json schema is correct")
            except Exception as e:
                self.log_callback(f"   ⚠️ Could not check source bookmarks.json: {e}")
    
    def _final_schema_safeguard(self, report_dir: Path):
        """Final safeguard to ensure schemas are correct before completion."""
        self.log_callback("   🛽️ Running FINAL schema safeguard...")
        
        bookmarks_json = report_dir / "definition" / "bookmarks" / "bookmarks.json"
        if bookmarks_json.exists():
            try:
                # Read the file as text to see exactly what's there
                with open(bookmarks_json, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # If we find the bad schema, replace it directly in the text
                if '"$schema": "1.0.0"' in content:
                    self.log_callback("   ❌ CRITICAL: Found bad schema at final check!")
                    
                    # Replace the bad schema with the correct one (bookmarksMetadata!)
                    fixed_content = content.replace(
                        '"$schema": "1.0.0"',
                        '"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmarksMetadata/1.0.0/schema.json"'
                    )
                    
                    # Also fix the structure from 'bookmarks' to 'items'
                    fixed_content = fixed_content.replace('"bookmarks":', '"items":')
                    
                    # Fix any remaining wrong schema URLs
                    fixed_content = fixed_content.replace(
                        'definition/bookmarks/1.0.0/schema.json',
                        'definition/bookmarksMetadata/1.0.0/schema.json'
                    )
                    
                    # Write the fixed content back
                    with open(bookmarks_json, 'w', encoding='utf-8') as f:
                        f.write(fixed_content)
                    
                    self.log_callback("   ✅ Fixed bad schema using text replacement")
                else:
                    self.log_callback("   ✅ Final check passed - schema is correct")
                    
            except Exception as e:
                self.log_callback(f"   ❌ Final safeguard error: {e}")
