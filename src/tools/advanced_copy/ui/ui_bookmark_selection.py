"""
Bookmark Selection Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Handles bookmark selection UI for bookmark+visual copy mode.
"""

import tkinter as tk
from tkinter import ttk
import json
from pathlib import Path
from typing import List, Dict, Any, Optional

from core.constants import AppConstants


class BookmarkSelectionMixin:
    """
    Mixin for bookmark selection functionality.
    
    Methods extracted from AdvancedCopyTab:
    - _show_bookmark_mode_selection_ui()
    - _on_source_page_change()
    - _populate_target_pages()
    - _load_bookmarks_for_page()
    - _update_selection_status()
    """
    
    def _show_bookmark_mode_selection_ui(self, pages_with_bookmarks: List[Dict[str, Any]]):
        """Show bookmark + visual copy mode UI"""
        # Show the pages frame
        self.pages_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Adjust window height
        self._adjust_window_height(True)
        
        # CRITICAL: Store pages for later access
        self.available_pages = pages_with_bookmarks
        
        # Get ALL pages from analysis results (including pages without bookmarks)
        all_pages = []
        if self.analysis_results:
            all_pages.extend(self.analysis_results.get('pages_with_bookmarks', []))
            all_pages.extend(self.analysis_results.get('pages_without_bookmarks', []))
        self.all_report_pages = all_pages  # Store for target page selection
        
        self.log_message(f"\n📄 Pages available for bookmark copying:")
        self.log_message(f"   Source pages (with bookmarks): {len(pages_with_bookmarks)}")
        self.log_message(f"   Total pages in report: {len(all_pages)}")
        self.log_message(f"   Target pages available: {len(all_pages) - 1} (excluding source)")
        
        # Clear existing content
        for widget in self.pages_frame.winfo_children():
            widget.destroy()
        
        # Update frame title
        self.pages_frame.config(text="🔖 BOOKMARK + VISUAL COPY SETUP")
        
        # Main container
        content_frame = ttk.Frame(self.pages_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid - now only 2 columns instead of 3
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # COLUMN 1: Source Page Selection
        source_frame = tk.LabelFrame(content_frame, text="📄 Source Page",
                                    font=('Segoe UI', 10, 'bold'),
                                    fg=AppConstants.COLORS['text_primary'],
                                    bg='#dcdad5',
                                    bd=1, relief='solid', padx=10, pady=10)
        source_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Source page title
        tk.Label(source_frame, text="Select page with bookmarks:",
                 font=('Segoe UI', 9, 'bold'),
                 bg='#dcdad5',
                 fg=AppConstants.COLORS['text_primary']).pack(anchor=tk.W, pady=(0, 5))
        
        # Source page dropdown
        source_pages = [f"{p['display_name']} ({p['bookmark_count']} bookmarks)" 
                       for p in pages_with_bookmarks]
        source_combo = ttk.Combobox(source_frame, textvariable=self.source_page_var,
                                   values=source_pages, state='readonly', width=30)
        source_combo.pack(fill=tk.X, pady=(0, 10))
        source_combo.bind('<<ComboboxSelected>>', self._on_source_page_change)
        
        # Configure combobox to have white background
        style = ttk.Style()
        style.map('TCombobox', fieldbackground=[('readonly', 'white')])
        style.map('TCombobox', selectbackground=[('readonly', 'white')])
        style.map('TCombobox', selectforeground=[('readonly', 'black')])
        
        # Store reference for later use
        self._source_combo = source_combo
        
        # Bookmarks from source page with inline helper text
        title_row = tk.Frame(source_frame, bg='#dcdad5')
        title_row.pack(anchor=tk.W, fill=tk.X, pady=(5, 2))
        
        # Store reference to the bookmarks label so we can update it dynamically
        self.bookmarks_instruction_label = tk.Label(title_row, text="Select source page to begin",
                 font=('Segoe UI', 9, 'bold'),
                 bg='#dcdad5',
                 fg=AppConstants.COLORS['text_secondary'])
        self.bookmarks_instruction_label.pack(side=tk.LEFT)
        
        self.bookmarks_helper_label = tk.Label(title_row, text="",
                 font=('Segoe UI', 8),
                 bg='#dcdad5',
                 fg=AppConstants.COLORS['info'])
        self.bookmarks_helper_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Warning text about Selected visuals requirement
        warning_label = tk.Label(source_frame, 
                                text="⚠️ Only 'Selected visuals' configured bookmarks are supported",
                                font=('Segoe UI', 8),
                                bg='#dcdad5',
                                fg=AppConstants.COLORS['warning'])
        warning_label.pack(anchor=tk.W, pady=(0, 2))
        
        bookmark_list_frame = tk.Frame(source_frame, bg='#dcdad5')
        bookmark_list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Use Treeview instead of Listbox for hierarchical display
        self.bookmarks_treeview = ttk.Treeview(
            bookmark_list_frame,
            selectmode='extended',
            height=4,
            show='tree'  # Hide column headers
        )
        self.bookmarks_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # Configure treeview background to white
        style = ttk.Style()
        style.configure('Treeview', background='white', fieldbackground='white')
        
        bookmark_scroll = ttk.Scrollbar(bookmark_list_frame, orient=tk.VERTICAL,
                                       command=self.bookmarks_treeview.yview)
        bookmark_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.bookmarks_treeview.configure(yscrollcommand=bookmark_scroll.set)
        
        # COLUMN 2: Target Pages Selection
        target_frame = tk.LabelFrame(content_frame, text="🎯 Target Pages",
                                    font=('Segoe UI', 10, 'bold'),
                                    fg=AppConstants.COLORS['text_primary'],
                                    bg='#dcdad5',
                                    bd=1, relief='solid', padx=10, pady=10)
        target_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Target pages title with inline helper text
        title_row = tk.Frame(target_frame, bg='#dcdad5')
        title_row.pack(anchor=tk.W, fill=tk.X, pady=(0, 5))
        
        tk.Label(title_row, text="Copy TO these pages:",
                 font=('Segoe UI', 9, 'bold'),
                 bg='#dcdad5',
                 fg=AppConstants.COLORS['text_primary']).pack(side=tk.LEFT)
        
        tk.Label(title_row, text="Use Ctrl+Click for multiple",
                 font=('Segoe UI', 8),
                 bg='#dcdad5',
                 fg=AppConstants.COLORS['info']).pack(side=tk.LEFT, padx=(10, 0))
        
        target_list_frame = tk.Frame(target_frame, bg='#dcdad5')
        target_list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.target_pages_listbox = tk.Listbox(
            target_list_frame,
            selectmode=tk.EXTENDED,
            height=4,
            font=('Segoe UI', 9),
            bg='white',
            fg=AppConstants.COLORS['text_primary'],
            selectbackground=AppConstants.COLORS['accent'],
            exportselection=False
        )
        self.target_pages_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        target_scroll = ttk.Scrollbar(target_list_frame, orient=tk.VERTICAL,
                                     command=self.target_pages_listbox.yview)
        target_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.target_pages_listbox.configure(yscrollcommand=target_scroll.set)
        
        # Initially populate with ALL pages (will be filtered when source is selected)
        self._populate_target_pages(exclude_page_name=None)
        
        # Bind selection events - bind Button-1 FIRST to intercept clicks before selection
        self.bookmarks_treeview.bind('<ButtonPress-1>', self._on_tree_click, add='+')
        self.bookmarks_treeview.bind('<<TreeviewSelect>>', self._on_bookmark_tree_select)
        self.target_pages_listbox.bind('<<ListboxSelect>>', self._on_target_listbox_select)
        
        # Selection summary at bottom
        self.selection_label = ttk.Label(content_frame, text="Select source page to begin",
                                       font=('Segoe UI', 9),
                                       foreground=AppConstants.COLORS['text_secondary'])
        self.selection_label.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # If there's only one page with bookmarks, auto-select it
        if len(pages_with_bookmarks) == 1:
            source_combo.current(0)  # Select first item
            self._on_source_page_change()  # Manually trigger the event
    
    def _on_source_page_change(self, event=None):
        """Handle source page selection in bookmark mode"""
        if not self.source_page_var.get():
            self.log_message("   ⚠️ No source page selected")
            return
        
        # Find selected page
        selected_display = self.source_page_var.get().split(' (')[0]  # Remove bookmark count
        self.log_message(f"   🔍 Looking for page: '{selected_display}'")
        
        selected_page = next((p for p in self.available_pages 
                            if p['display_name'] == selected_display), None)
        
        if not selected_page:
            self.log_message(f"   ❌ Could not find page '{selected_display}' in available pages")
            self.log_message(f"   Available pages: {[p['display_name'] for p in self.available_pages]}")
            return
        
        self.log_message(f"   ✅ Found page: {selected_page['display_name']} ({selected_page['bookmark_count']} bookmarks)")
        
        # Update the instruction label to show next step
        if hasattr(self, 'bookmarks_instruction_label'):
            self.bookmarks_instruction_label.config(
                text="Bookmarks to copy:",
                fg=AppConstants.COLORS['text_primary']
            )
        if hasattr(self, 'bookmarks_helper_label'):
            self.bookmarks_helper_label.config(
                text="💡 Click groups to select all bookmarks"
            )
        
        # Load bookmarks for this page
        self._load_bookmarks_for_page(selected_page)
        
        # Update target pages list (exclude source page)
        self._populate_target_pages(exclude_page_name=selected_page['name'])
        
        # Update selection status
        self._update_selection_status()
    
    def _populate_target_pages(self, exclude_page_name: Optional[str] = None):
        """Populate target pages listbox with all pages except the excluded one"""
        if not self.target_pages_listbox:
            return
        
        # Determine which pages to use based on destination mode
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"
        
        if is_cross_pbip:
            # Use TARGET report's pages (from target_analysis_results)
            if not hasattr(self, 'target_analysis_results') or not self.target_analysis_results:
                self.log_message("   ⚠️ Target report not analyzed yet - cannot populate target pages")
                return
            
            # Get all pages from target analysis
            target_pages = []
            target_pages.extend(self.target_analysis_results.get('pages_with_bookmarks', []))
            target_pages.extend(self.target_analysis_results.get('pages_without_bookmarks', []))
            
            pages_to_use = target_pages
            self.log_message(f"   🌍 Using TARGET report pages ({len(pages_to_use)} pages)")
        else:
            # Use SOURCE report's pages (same PBIP mode)
            if not hasattr(self, 'all_report_pages') or not self.all_report_pages:
                self.log_message("   ⚠️ No pages available for target selection")
                return
            
            pages_to_use = self.all_report_pages
            self.log_message(f"   📄 Using SOURCE report pages ({len(pages_to_use)} pages)")
        
        # Set guard flag to prevent event interference
        self._updating_ui = True
        try:
            # Temporarily unbind the selection event
            self.target_pages_listbox.unbind('<<ListboxSelect>>')
            self.target_pages_listbox.unbind('<Button-1>')
            
            # Clear existing items
            self.target_pages_listbox.delete(0, tk.END)
            
            # Add "Select All" option as first item
            self.target_pages_listbox.insert(tk.END, "📌 (Select All)")
            
            # Add all pages except the excluded one (only exclude in same-PBIP mode)
            pages_added = 0
            for page in pages_to_use:
                # In cross-PBIP mode, don't exclude any pages (they're from different report)
                # In same-PBIP mode, exclude the source page
                should_add = is_cross_pbip or (exclude_page_name is None or page['name'] != exclude_page_name)
                
                if should_add:
                    self.target_pages_listbox.insert(tk.END, page['display_name'])
                    pages_added += 1
            
            self.log_message(f"   📄 Target pages available: {pages_added}")
            
            # Re-bind the selection events - bind Button-1 FIRST to intercept clicks
            self.target_pages_listbox.bind('<Button-1>', self._on_target_listbox_click, add='+')
            self.target_pages_listbox.bind('<<ListboxSelect>>', self._on_target_listbox_select)
        finally:
            self._updating_ui = False
    
    def _load_bookmarks_for_page(self, page_data: Dict[str, Any]):
        """Load bookmarks for selected page with group hierarchy and capture mode detection"""
        if not self.bookmarks_treeview:
            self.log_message("   ❌ Bookmarks treeview not initialized")
            return
        
        # Set guard flag to prevent event interference
        self._updating_ui = True
        try:
            # Temporarily unbind the selection events
            self.bookmarks_treeview.unbind('<<TreeviewSelect>>')
            self.bookmarks_treeview.unbind('<Button-1>')
            
            # Clear existing tree
            for item in self.bookmarks_treeview.get_children():
                self.bookmarks_treeview.delete(item)
            
            # Reset mapping
            self._bookmark_tree_mapping = {}
            
            # Get bookmark names from analysis
            bookmark_names = page_data.get('bookmark_names', [])
            self.log_message(f"   📋 Loading {len(bookmark_names)} bookmark(s): {bookmark_names}")
            
            if not bookmark_names:
                # Insert placeholder
                item_id = self.bookmarks_treeview.insert('', 'end', text='(No bookmarks found)')
                self.log_message("   ⚠️ No bookmark names found in page data")
                return
            
            # Load bookmark files and analyze capture modes
            report_path = self.clean_file_path(self.report_path.get())
            
            report_dir = Path(report_path).parent / f"{Path(report_path).stem}.Report"
            bookmarks_dir = report_dir / "definition" / "bookmarks"
            
            # Analyze bookmark capture modes
            capture_analysis = self.advanced_copy_engine.analyze_bookmark_capture_modes(
                bookmarks_dir, bookmark_names
            )
            
            # Log the analysis
            self.log_message(f"   📊 Bookmark Analysis:")
            self.log_message(f"      ✅ Selected Visuals: {capture_analysis['selected_visuals_count']}")
            self.log_message(f"      ⚠️ All Visuals: {capture_analysis['all_visuals_count']} (not supported)")
            
            if capture_analysis['selected_visuals_count'] > 0 and capture_analysis['all_visuals_count'] > 0:
                self.log_message(f"      💡 Compatible bookmarks will be listed first")
            
            # Load bookmark group structure
            bookmarks_json = bookmarks_dir / "bookmarks.json"
            bookmark_groups = {}  # bookmark_id -> group_info
            group_items = {}  # group_id -> {name, display_name, children[]}
            ungrouped_bookmarks = set(bookmark_names)  # Start with all, remove grouped ones
            
            if bookmarks_json.exists():
                try:
                    with open(bookmarks_json, 'r', encoding='utf-8') as f:
                        bookmarks_metadata = json.load(f)
                    
                    if 'items' in bookmarks_metadata:
                        for item in bookmarks_metadata['items']:
                            if 'children' in item:  # This is a group
                                group_id = item['name']
                                group_display = item.get('displayName', group_id)
                                children = [c for c in item['children'] if c in bookmark_names]
                                
                                if children:  # Only add groups that have bookmarks we're displaying
                                    group_items[group_id] = {
                                        'name': group_id,
                                        'display_name': group_display,
                                        'children': children
                                    }
                                    
                                    # Mark these bookmarks as grouped
                                    for child in children:
                                        bookmark_groups[child] = group_id
                                        ungrouped_bookmarks.discard(child)
                                        
                except Exception as e:
                    self.log_message(f"   ⚠️ Could not load bookmark groups: {e}")
            
            # Load individual bookmark display names
            bookmark_display_names = {}
            for bookmark_name in bookmark_names:
                bookmark_file = bookmarks_dir / f"{bookmark_name}.bookmark.json"
                try:
                    with open(bookmark_file, 'r', encoding='utf-8') as f:
                        bookmark_data = json.load(f)
                    bookmark_display_names[bookmark_name] = bookmark_data.get('displayName', bookmark_name)
                except Exception as e:
                    bookmark_display_names[bookmark_name] = bookmark_name
                    self.log_message(f"      ⚠️ Could not load {bookmark_name}: {e}")
            
            # Build tree: Sort groups by compatibility, then add ungrouped bookmarks
            groups_added = 0
            bookmarks_added = 0
            disabled_count = 0
            
            # Categorize groups by whether they have ANY compatible bookmarks
            compatible_groups = []  # Groups with at least one 'selected_visuals' bookmark
            incompatible_groups = []  # Groups with ONLY 'all_visuals' bookmarks
            
            for group_id, group_info in group_items.items():
                # Check if this group has ANY compatible bookmarks
                has_compatible = False
                for child_id in group_info['children']:
                    bookmark_info = capture_analysis['bookmark_details'].get(child_id, {})
                    mode = bookmark_info.get('mode', 'all_visuals')
                    if mode == 'selected_visuals':
                        has_compatible = True
                        break
                
                if has_compatible:
                    compatible_groups.append((group_id, group_info))
                else:
                    incompatible_groups.append((group_id, group_info))
            
            # Add compatible groups first, then incompatible groups
            for group_id, group_info in compatible_groups + incompatible_groups:
                # Add group node
                group_tree_id = self.bookmarks_treeview.insert(
                    '', 'end',
                    text=f"📁 {group_info['display_name']} ({len(group_info['children'])} bookmarks)",
                    open=True  # Expand by default
                )
                
                # Mark this as a group in mapping
                self._bookmark_tree_mapping[group_tree_id] = {
                    'type': 'group',
                    'group_id': group_id,
                    'bookmark_ids': group_info['children']
                }
                
                # Add child bookmarks
                for child_id in group_info['children']:
                    child_display = bookmark_display_names.get(child_id, child_id)
                    
                    # Get capture mode info
                    bookmark_info = capture_analysis['bookmark_details'].get(child_id, {})
                    mode = bookmark_info.get('mode', 'all_visuals')
                    visual_count = bookmark_info.get('visual_count', 0)
                    
                    # Format text based on mode
                    if mode == 'selected_visuals':
                        text = f"  🔖 {child_display} ({visual_count} visuals)"
                        tags = ()
                    else:
                        text = f"  ⚠️ {child_display} (All visuals - not supported)"
                        tags = ('disabled',)
                        disabled_count += 1
                    
                    child_tree_id = self.bookmarks_treeview.insert(
                        group_tree_id, 'end',
                        text=text,
                        tags=tags
                    )
                    
                    # Map tree item to bookmark ID
                    self._bookmark_tree_mapping[child_tree_id] = {
                        'type': 'bookmark',
                        'bookmark_id': child_id,
                        'group_id': group_id,
                        'mode': mode,
                        'selectable': mode == 'selected_visuals'
                    }
                    bookmarks_added += 1
                
                groups_added += 1
                self.log_message(f"      📁 Group: {group_info['display_name']} ({len(group_info['children'])} bookmarks)")
            
            # Add ungrouped bookmarks
            if ungrouped_bookmarks:
                for bookmark_id in sorted(ungrouped_bookmarks):
                    display_name = bookmark_display_names.get(bookmark_id, bookmark_id)
                    
                    # Get capture mode info
                    bookmark_info = capture_analysis['bookmark_details'].get(bookmark_id, {})
                    mode = bookmark_info.get('mode', 'all_visuals')
                    visual_count = bookmark_info.get('visual_count', 0)
                    
                    # Format text based on mode
                    if mode == 'selected_visuals':
                        text = f"🔖 {display_name} ({visual_count} visuals)"
                        tags = ()
                    else:
                        text = f"⚠️ {display_name} (All visuals - not supported)"
                        tags = ('disabled',)
                        disabled_count += 1
                    
                    tree_id = self.bookmarks_treeview.insert(
                        '', 'end',
                        text=text,
                        tags=tags
                    )
                    
                    # Map tree item to bookmark ID
                    self._bookmark_tree_mapping[tree_id] = {
                        'type': 'bookmark',
                        'bookmark_id': bookmark_id,
                        'group_id': None,
                        'mode': mode,
                        'selectable': mode == 'selected_visuals'
                    }
                    bookmarks_added += 1
            
            # Configure tree tags for disabled items
            self.bookmarks_treeview.tag_configure('disabled', foreground='#94a3b8')  # Grey color
            
            # Check if all bookmarks are disabled
            if disabled_count == bookmarks_added:
                # Disable the entire tree by unbinding events and showing a message
                self.bookmarks_treeview.unbind('<<TreeviewSelect>>')
                self.bookmarks_treeview.unbind('<Button-1>')
                self.log_message(f"   ⚠️ All bookmarks use 'All visuals' mode - selection disabled")
            # If there are selectable bookmarks, the events are already bound
            
            self.log_message(f"   ✅ Loaded {groups_added} groups, {bookmarks_added} bookmarks ({disabled_count} disabled)")
            
            # Re-bind the selection events - bind ButtonPress-1 FIRST
            self.bookmarks_treeview.bind('<ButtonPress-1>', self._on_tree_click, add='+')
            self.bookmarks_treeview.bind('<<TreeviewSelect>>', self._on_bookmark_tree_select)
        finally:
            self._updating_ui = False
    
    def _update_selection_status(self):
        """Update the selection status and button states (common logic)"""
        if self._updating_ui:
            return
        
        # Clear the pending update reference
        self._pending_update = None
            
        if not self.bookmarks_treeview or not self.target_pages_listbox:
            return
        
        try:
            self._updating_ui = True
            
            # Count selected bookmarks (not groups, and only selectable ones)
            selected_items = self.bookmarks_treeview.selection()
            bookmarks_selected = sum(1 for item in selected_items 
                                   if item in self._bookmark_tree_mapping 
                                   and self._bookmark_tree_mapping[item]['type'] == 'bookmark'
                                   and self._bookmark_tree_mapping[item].get('selectable', True))
            
            # Count selected target pages (excluding index 0 which is "Select All")
            target_selection = self.target_pages_listbox.curselection()
            targets_selected = sum(1 for idx in target_selection if idx > 0)
            
            # Update status message
            if not self.source_page_var.get():
                self.selection_label.config(text="Select source page to begin")
                if self.copy_button:
                    self.copy_button.config(state=tk.DISABLED)
            elif bookmarks_selected == 0:
                self.selection_label.config(text="Select bookmarks to copy")
                if self.copy_button:
                    self.copy_button.config(state=tk.DISABLED)
            elif targets_selected == 0:
                self.selection_label.config(text="Select target pages")
                if self.copy_button:
                    self.copy_button.config(state=tk.DISABLED)
            else:
                self.selection_label.config(
                    text=f"Ready: {bookmarks_selected} bookmark(s) → {targets_selected} page(s)")
                if self.copy_button:
                    self.copy_button.config(state=tk.NORMAL)
        finally:
            self._updating_ui = False
