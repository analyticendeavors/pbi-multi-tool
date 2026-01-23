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
from core.ui_base import ThemedScrollbar


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
        """Show bookmark + visual copy mode UI with modern styling matching Analysis Summary/Progress Log"""
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

        # Get dynamic theme colors
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.current_theme == 'dark'
        main_bg = '#0d0d1a' if is_dark else '#ffffff'
        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f5f5fa')
        content_bg = colors.get('background', '#0d0d1a' if is_dark else '#ffffff')
        option_bg = colors.get('option_bg', '#f5f5f7')

        # Update frame title with advanced copy icon (using labelwidget)
        # Title sits on section_bg (gray) background - matches Analysis & Progress style
        advanced_copy_icon = self._load_icon_for_button('advanced copy', size=16)

        title_frame = tk.Frame(self.pages_frame, bg=section_bg)
        if advanced_copy_icon:
            icon_label = tk.Label(title_frame, image=advanced_copy_icon, bg=section_bg)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            icon_label._icon_ref = advanced_copy_icon
            self._bookmark_title_icon = icon_label
        title_label = tk.Label(title_frame, text="Copy Configuration",
                              font=('Segoe UI Semibold', 11),
                              bg=section_bg, fg=colors['title_color'])
        title_label.pack(side=tk.LEFT)
        self.pages_frame.config(labelwidget=title_frame)
        self._bookmark_title_frame = title_frame
        self._bookmark_title_label = title_label

        # Keep Section.TLabelframe style (matches Analysis & Progress)
        self.pages_frame.configure(style='Section.TLabelframe')

        # Main container - uses content_bg (dark/white) like Analysis & Progress inner area
        # This matches the Section.TFrame pattern - padding 15 for inner content
        # Use fill=tk.X and expand=False to prevent vertical expansion (keeps windows compact)
        content_frame = ttk.Frame(self.pages_frame, style='Section.TFrame', padding="15")
        content_frame.pack(fill=tk.X, expand=False)

        # Store reference for theme updates
        self._bookmark_content_frame = content_frame

        # Consistent spacing for inner sections - matches content_frame padding (15)
        inner_pad = 15  # Padding between inner sections equals outer border padding

        # Configure grid - 3 equal columns for the three sections
        # Use uniform="cols" to guarantee equal column widths
        content_frame.columnconfigure(0, weight=1, uniform="cols")  # Select page + Copy Mode left
        content_frame.columnconfigure(1, weight=1, uniform="cols")  # Bookmarks to copy
        content_frame.columnconfigure(2, weight=1, uniform="cols")  # Target Pages (spans row 0-1)
        # Row 0: Select page | Bookmarks | Target Pages (Target spans row 0-1)
        # Row 1: Copy Mode (spans cols 0-1) | [Target cont]
        # Row 2: Status text (full width)
        content_frame.rowconfigure(0, weight=0)  # List windows row - fixed height (use canvas height)
        content_frame.rowconfigure(1, weight=0)  # Copy Mode row - fixed height

        # Common styling
        list_bg = '#1e1e2e' if is_dark else '#ffffff'

        # ============ COLUMN 1: Select page ============
        page_frame = tk.Frame(content_frame, bg=content_bg)
        page_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, inner_pad//2), pady=0)
        self._bookmark_page_frame = page_frame

        # Header with icon (matches Analysis Summary style)
        page_header = tk.Frame(page_frame, bg=content_bg)
        page_header.pack(anchor=tk.W, pady=(0, 8))
        self._bookmark_page_header = page_header

        page_icon = self._load_icon_for_button('file', size=16)
        if page_icon:
            page_icon_lbl = tk.Label(page_header, image=page_icon, bg=content_bg)
            page_icon_lbl.pack(side=tk.LEFT, padx=(0, 8))
            page_icon_lbl._icon_ref = page_icon
            self._bookmark_page_icon = page_icon_lbl

        self._bookmark_page_label = tk.Label(page_header, text="Select Page",
                 font=('Segoe UI Semibold', 11),
                 bg=content_bg,
                 fg=colors['title_color'])
        self._bookmark_page_label.pack(side=tk.LEFT)

        # Page list container (bordered inner window)
        page_list_container = tk.Frame(page_frame, bg=list_bg,
                                       highlightthickness=1,
                                       highlightbackground=colors['border'])
        page_list_container.pack(fill=tk.BOTH, expand=False)
        page_list_container.configure(height=160)  # +40px as requested
        page_list_container.pack_propagate(False)  # Enforce fixed height
        self._bookmark_page_list_container = page_list_container

        # Scrollable canvas for source pages
        page_canvas = tk.Canvas(page_list_container, bg=list_bg, highlightthickness=0)

        # Scrollbar for source pages - pack FIRST so it gets space before canvas expands
        page_scroll = ThemedScrollbar(page_list_container,
                                      command=page_canvas.yview,
                                      theme_manager=self._theme_manager,
                                      width=10, auto_hide=False)
        page_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._bookmark_page_scroll = page_scroll

        # Now pack canvas - it will take remaining space
        page_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 0), pady=3)
        page_canvas.configure(yscrollcommand=page_scroll.set)

        page_inner_frame = tk.Frame(page_canvas, bg=list_bg)
        page_window_id = page_canvas.create_window((3, 3), window=page_inner_frame, anchor='nw')
        self._bookmark_page_inner_frame = page_inner_frame
        self._bookmark_page_canvas = page_canvas

        # Configure canvas scrolling (height controlled by container)
        def on_page_frame_configure(event):
            page_canvas.configure(scrollregion=page_canvas.bbox('all'))
        page_inner_frame.bind('<Configure>', on_page_frame_configure)

        # Make inner frame expand to fill canvas width (responsive)
        def on_page_canvas_resize(event):
            page_canvas.itemconfig(page_window_id, width=event.width - 6)  # 3px padding each side
        page_canvas.bind('<Configure>', on_page_canvas_resize)

        # Mouse wheel scrolling for source pages (with bounds checking)
        def on_page_mousewheel(event):
            # Check current scroll position
            current_pos = page_canvas.yview()
            # Only scroll if not at boundaries
            if event.delta > 0 and current_pos[0] > 0:  # Scrolling up, not at top
                page_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.delta < 0 and current_pos[1] < 1:  # Scrolling down, not at bottom
                page_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        page_canvas.bind('<Enter>', lambda e: page_canvas.bind_all('<MouseWheel>', on_page_mousewheel))
        page_canvas.bind('<Leave>', lambda e: page_canvas.unbind_all('<MouseWheel>'))
        self._on_page_mousewheel = on_page_mousewheel  # Store reference

        # Store source page widgets and selection state
        self._source_page_widgets = []
        self._selected_source_page_idx = None

        # Create source page rows
        for idx, page in enumerate(pages_with_bookmarks):
            self._create_source_page_row(page_inner_frame, idx, page, colors, is_dark)

        # ============ COLUMN 2: Bookmarks to copy ============
        bookmarks_frame = tk.Frame(content_frame, bg=content_bg)
        bookmarks_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(inner_pad//2, inner_pad//2), pady=0)
        self._bookmark_bookmarks_frame = bookmarks_frame

        # Header with icon (matches Analysis Summary style)
        bookmarks_header = tk.Frame(bookmarks_frame, bg=content_bg)
        bookmarks_header.pack(anchor=tk.W, pady=(0, 8), fill=tk.X)
        self._bookmark_title_row = bookmarks_header

        bookmark_icon = self._load_icon_for_button('bookmark', size=16)
        if bookmark_icon:
            bookmark_icon_lbl = tk.Label(bookmarks_header, image=bookmark_icon, bg=content_bg)
            bookmark_icon_lbl.pack(side=tk.LEFT, padx=(0, 8))
            bookmark_icon_lbl._icon_ref = bookmark_icon
            self._bookmark_bookmarks_icon = bookmark_icon_lbl

        self.bookmarks_instruction_label = tk.Label(bookmarks_header, text="Select Bookmarks",
                 font=('Segoe UI Semibold', 11),
                 bg=content_bg,
                 fg=colors['title_color'])
        self.bookmarks_instruction_label.pack(side=tk.LEFT)

        self.bookmarks_helper_label = tk.Label(bookmarks_header, text="(select page first)",
                 font=('Segoe UI', 8),
                 bg=content_bg,
                 fg=colors['text_secondary'])
        self.bookmarks_helper_label.pack(side=tk.LEFT, padx=(6, 0))

        # Bookmark list frame (bordered inner window)
        bookmark_list_frame = tk.Frame(bookmarks_frame, bg=list_bg,
                                       highlightthickness=1,
                                       highlightbackground=colors['border'])
        bookmark_list_frame.pack(fill=tk.BOTH, expand=False)
        bookmark_list_frame.configure(height=160)  # +40px as requested
        bookmark_list_frame.pack_propagate(False)  # Enforce fixed height
        self._bookmark_list_frame = bookmark_list_frame

        # Modern treeview styling
        tree_bg = list_bg
        tree_fg = colors['text_primary']
        heading_bg = '#2a2a3e' if is_dark else '#f0f0f5'
        selected_bg = '#1a3a5c' if is_dark else '#e6f3ff'
        header_separator = main_bg

        style = ttk.Style()
        tree_style = 'BookmarkTree.Treeview'
        style.configure(tree_style,
                        background=tree_bg, foreground=tree_fg,
                        fieldbackground=tree_bg,
                        font=('Segoe UI', 9), rowheight=25,
                        relief='flat', borderwidth=0,
                        bordercolor=tree_bg, lightcolor=tree_bg, darkcolor=tree_bg)
        style.layout(tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])
        style.configure(f"{tree_style}.Heading",
                        background=heading_bg, foreground=tree_fg,
                        font=('Segoe UI', 9, 'bold'),
                        relief='groove', borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        padding=(8, 4))
        style.map(f"{tree_style}.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])
        style.map(tree_style,
                  background=[('selected', selected_bg)],
                  foreground=[('selected', tree_fg)])

        # Treeview for hierarchical display
        self.bookmarks_treeview = ttk.Treeview(
            bookmark_list_frame,
            selectmode='extended',
            height=3,
            show='tree',
            style=tree_style
        )
        self.bookmarks_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 0), pady=3)

        # Scrollbar for bookmarks
        bookmark_scroll = ThemedScrollbar(bookmark_list_frame,
                                          command=self.bookmarks_treeview.yview,
                                          theme_manager=self._theme_manager,
                                          width=10, auto_hide=False)
        bookmark_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.bookmarks_treeview.configure(yscrollcommand=bookmark_scroll.set)
        self._bookmark_scrollbar = bookmark_scroll

        # ============ COLUMN 3: Target Pages (spans rows 0-1 to align with Copy Mode) ============
        target_frame = tk.Frame(content_frame, bg=content_bg)
        target_frame.grid(row=0, column=2, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(inner_pad//2, 0), pady=0)
        self._bookmark_target_frame = target_frame

        # Header with icon (matches Analysis Summary style)
        target_header = tk.Frame(target_frame, bg=content_bg)
        target_header.pack(anchor=tk.W, pady=(0, 8))
        self._bookmark_target_header = target_header

        target_icon = self._load_icon_for_button('target', size=16)
        if target_icon:
            target_icon_lbl = tk.Label(target_header, image=target_icon, bg=content_bg)
            target_icon_lbl.pack(side=tk.LEFT, padx=(0, 8))
            target_icon_lbl._icon_ref = target_icon
            self._bookmark_target_icon = target_icon_lbl

        self._bookmark_target_hint = tk.Label(target_header, text="Target Pages",
                 font=('Segoe UI Semibold', 11),
                 bg=content_bg,
                 fg=colors['title_color'])
        self._bookmark_target_hint.pack(side=tk.LEFT)

        # Target list container (bordered inner window)
        target_list_container = tk.Frame(target_frame, bg=list_bg,
                                         highlightthickness=1,
                                         highlightbackground=colors['border'])
        target_list_container.pack(fill=tk.BOTH, expand=False)
        target_list_container.configure(height=215)  # Taller to align with Copy Mode bottom (+40px)
        target_list_container.pack_propagate(False)  # Enforce fixed height
        self._bookmark_target_list_frame = target_list_container

        # Scrollable canvas for target pages
        target_canvas = tk.Canvas(target_list_container, bg=list_bg, highlightthickness=0)

        # Scrollbar for target pages - pack FIRST so it gets space before canvas expands
        target_scroll = ThemedScrollbar(target_list_container,
                                        command=target_canvas.yview,
                                        theme_manager=self._theme_manager,
                                        width=10, auto_hide=False)
        target_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._bookmark_target_scrollbar = target_scroll

        # Now pack canvas - it will take remaining space
        target_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 0), pady=3)
        target_canvas.configure(yscrollcommand=target_scroll.set)

        target_inner_frame = tk.Frame(target_canvas, bg=list_bg)
        target_window_id = target_canvas.create_window((3, 3), window=target_inner_frame, anchor='nw')
        self._target_inner_frame = target_inner_frame
        self._target_canvas = target_canvas

        # Configure canvas scrolling (height controlled by container)
        def on_target_frame_configure(event):
            target_canvas.configure(scrollregion=target_canvas.bbox('all'))
        target_inner_frame.bind('<Configure>', on_target_frame_configure)

        # Make inner frame expand to fill canvas width (responsive)
        def on_target_canvas_resize(event):
            target_canvas.itemconfig(target_window_id, width=event.width - 6)  # 3px padding each side
        target_canvas.bind('<Configure>', on_target_canvas_resize)

        # Mouse wheel scrolling for target pages (with bounds checking)
        def on_target_mousewheel(event):
            # Check current scroll position
            current_pos = target_canvas.yview()
            # Only scroll if not at boundaries
            if event.delta > 0 and current_pos[0] > 0:  # Scrolling up, not at top
                target_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.delta < 0 and current_pos[1] < 1:  # Scrolling down, not at bottom
                target_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        target_canvas.bind('<Enter>', lambda e: target_canvas.bind_all('<MouseWheel>', on_target_mousewheel))
        target_canvas.bind('<Leave>', lambda e: target_canvas.unbind_all('<MouseWheel>'))
        self._on_target_mousewheel = on_target_mousewheel  # Store reference

        # Store target page widgets and selection state
        self._target_page_widgets = []
        self._selected_target_pages = set()
        self.target_pages_listbox = None  # No longer using Listbox

        # ============ ROW 1: Copy Mode (unified box spanning all columns) ============
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"

        # Copy Mode frame - only under Select page and Bookmarks (columns 0-1)
        mode_frame = tk.Frame(content_frame, bg=option_bg,
                              highlightbackground=colors['border'],
                              highlightthickness=1)
        mode_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=(0, inner_pad//2), pady=(inner_pad, 0))
        self._mode_frame = mode_frame  # Keep reference for theme updates
        self._mode_right_frame = None  # No longer split

        mode_inner = tk.Frame(mode_frame, bg=option_bg)
        mode_inner.pack(fill=tk.X, padx=12, pady=8)

        # Header (cogwheel icon + "Copy Mode" title)
        cogwheel_icon = self._button_icons.get('cogwheel')
        mode_header_frame = tk.Frame(mode_inner, bg=option_bg)
        mode_header_frame.pack(side=tk.LEFT)

        if cogwheel_icon:
            icon_lbl = tk.Label(mode_header_frame, image=cogwheel_icon, bg=option_bg)
            icon_lbl.pack(side=tk.LEFT, padx=(0, 6))
            icon_lbl._icon_ref = cogwheel_icon
            self._bookmark_mode_icon = icon_lbl
        title_lbl = tk.Label(mode_header_frame, text="Copy Mode:", bg=option_bg,
                            fg=colors['text_primary'], font=('Segoe UI', 9, 'bold'))
        title_lbl.pack(side=tk.LEFT)
        self._bookmark_mode_header = mode_header_frame
        self._bookmark_mode_title = title_lbl

        # Per-Page option
        self._bookmark_perpage_frame = tk.Frame(mode_inner, bg=option_bg)
        self._bookmark_perpage_frame.pack(side=tk.LEFT, padx=(24, 0))

        self._perpage_radio_canvas = self._create_svg_radio_bookmark(
            self._bookmark_perpage_frame, self.bookmark_copy_mode, "perpage", colors, option_bg
        )
        self._perpage_radio_canvas.pack(side=tk.LEFT)

        self._perpage_text = tk.Label(self._bookmark_perpage_frame, text="Per-Page",
                 font=('Segoe UI', 9), bg=option_bg,
                 fg=colors['text_primary'],
                 cursor='hand2', anchor=tk.W)
        self._perpage_text.pack(side=tk.LEFT, padx=(4, 0))
        self._perpage_text.bind('<Button-1>', lambda e: self._select_bookmark_mode("perpage"))

        # Hover underline effect for Per-Page
        def on_perpage_enter(e):
            self._perpage_text.configure(font=('Segoe UI', 9, 'underline'))
        def on_perpage_leave(e):
            self._perpage_text.configure(font=('Segoe UI', 9))
        self._perpage_text.bind('<Enter>', on_perpage_enter)
        self._perpage_text.bind('<Leave>', on_perpage_leave)

        # Description inline
        self._perpage_label = tk.Label(self._bookmark_perpage_frame,
                 text="- New bookmark per page",
                 font=('Segoe UI', 8),
                 bg=option_bg,
                 fg=colors['info'])
        self._perpage_label.pack(side=tk.LEFT, padx=(4, 0))

        # Cross-Page option (in same row, after Per-Page)
        self._bookmark_crosspage_frame = tk.Frame(mode_inner, bg=option_bg)
        self._bookmark_crosspage_frame.pack(side=tk.LEFT, padx=(32, 0))

        self._crosspage_radio_canvas = self._create_svg_radio_bookmark(
            self._bookmark_crosspage_frame, self.bookmark_copy_mode, "crosspage", colors, option_bg
        )
        self._crosspage_radio_canvas.pack(side=tk.LEFT)

        # Cross-Page is now enabled for both same-PBIP and cross-PBIP modes
        self._crosspage_text = tk.Label(self._bookmark_crosspage_frame, text="Cross-Page",
                 font=('Segoe UI', 9), bg=option_bg,
                 fg=colors['text_primary'],
                 cursor='hand2', anchor=tk.W)
        self._crosspage_text.pack(side=tk.LEFT, padx=(4, 0))
        self._crosspage_text.bind('<Button-1>', lambda e: self._select_bookmark_mode("crosspage"))

        # Hover underline effect for Cross-Page
        def on_crosspage_enter(e):
            self._crosspage_text.configure(font=('Segoe UI', 9, 'underline'))
        def on_crosspage_leave(e):
            self._crosspage_text.configure(font=('Segoe UI', 9))
        self._crosspage_text.bind('<Enter>', on_crosspage_enter)
        self._crosspage_text.bind('<Leave>', on_crosspage_leave)

        # Description inline
        self._crosspage_label = tk.Label(self._bookmark_crosspage_frame,
                 text="- One bookmark, all pages",
                 font=('Segoe UI', 8),
                 bg=option_bg,
                 fg=colors['info'])
        self._crosspage_label.pack(side=tk.LEFT, padx=(4, 0))

        # Set initial text colors based on default selection (perpage)
        self._update_bookmark_mode_colors()

        # ============ ROW 2: Status text (centered, below Copy Mode) ============
        status_frame = tk.Frame(content_frame, bg=content_bg)
        status_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(inner_pad, 0))
        self._bookmark_status_frame = status_frame

        self.selection_label = tk.Label(status_frame, text="📋 Select source page to begin",
                                       font=('Segoe UI', 9),
                                       bg=content_bg,
                                       fg=colors['text_secondary'])
        self.selection_label.pack(anchor=tk.CENTER)

        # Initially populate with ALL pages (will be filtered when source is selected)
        self._populate_target_pages_modern(exclude_page_name=None)

        # Bind selection events
        self.bookmarks_treeview.bind('<ButtonPress-1>', self._on_tree_click, add='+')
        self.bookmarks_treeview.bind('<<TreeviewSelect>>', self._on_bookmark_tree_select)

        # Auto-select first page if any compatible pages are found
        if len(pages_with_bookmarks) >= 1:
            self._select_source_page(0)
    
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
        
        # Update the helper label to show next step
        if hasattr(self, 'bookmarks_helper_label'):
            colors = self._theme_manager.colors
            self.bookmarks_helper_label.config(
                text="💡 Click groups to select all",
                fg=colors['info']
            )
        
        # Load bookmarks for this page
        self._load_bookmarks_for_page(selected_page)

        # Update target pages list (exclude source page)
        self._populate_target_pages_modern(exclude_page_name=selected_page['name'])

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
            pages_dir = report_dir / "definition" / "pages"

            # Analyze bookmark capture modes (including cross-page detection)
            capture_analysis = self.advanced_copy_engine.analyze_bookmark_capture_modes(
                bookmarks_dir, bookmark_names, pages_dir
            )

            # Initialize tooltip data storage for cross-page bookmarks
            self._bookmark_tooltip_data = {}

            # Load SVG icons for treeview items (smaller size for inline display)
            self._tree_folder_icon = self._load_icon_for_button('folder', size=14)
            self._tree_bookmark_icon = self._load_icon_for_button('bookmark', size=14)
            self._tree_crosspage_icon = self._load_icon_for_button('bookmark alt', size=14)
            self._tree_warning_icon = self._load_icon_for_button('warning', size=14)
            
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
                # Check if this is an incompatible group (only unsupported bookmarks)
                is_incompatible_group = (group_id, group_info) in incompatible_groups

                # Format group text - add "not supported" for incompatible groups
                bookmark_count = len(group_info['children'])
                if is_incompatible_group:
                    group_text = f" {group_info['display_name']} ({bookmark_count} bookmarks - not supported)"
                    group_tags = ('disabled',)  # Faded color like disabled bookmarks
                    group_open = False  # Collapsed by default
                else:
                    group_text = f" {group_info['display_name']} ({bookmark_count} bookmarks)"
                    group_tags = ()
                    group_open = True  # Expanded by default

                # Add group node with folder icon
                group_tree_id = self.bookmarks_treeview.insert(
                    '', 'end',
                    text=group_text,
                    image=self._tree_folder_icon if self._tree_folder_icon else '',
                    open=group_open,
                    tags=group_tags
                )

                # Mark this as a group in mapping (incompatible groups are not selectable)
                self._bookmark_tree_mapping[group_tree_id] = {
                    'type': 'group',
                    'group_id': group_id,
                    'bookmark_ids': group_info['children'],
                    'selectable': not is_incompatible_group
                }
                
                # Add child bookmarks
                for child_id in group_info['children']:
                    child_display = bookmark_display_names.get(child_id, child_id)

                    # Get capture mode info
                    bookmark_info = capture_analysis['bookmark_details'].get(child_id, {})
                    mode = bookmark_info.get('mode', 'all_visuals')
                    visual_count = bookmark_info.get('visual_count', 0)
                    is_cross_page = bookmark_info.get('is_cross_page', False)
                    page_count = bookmark_info.get('page_count', 0)
                    page_display_names_list = bookmark_info.get('page_display_names', [])

                    # Format text and select icon based on mode and cross-page status
                    if mode == 'selected_visuals':
                        if is_cross_page:
                            # Cross-page bookmark: show bookmark alt icon and page count
                            text = f"  {child_display} ({page_count} pages)"
                            icon = self._tree_crosspage_icon
                            tooltip_text = f"Cross-page: {', '.join(page_display_names_list)}"
                        else:
                            # Single-page bookmark: show bookmark icon and visual count
                            text = f"  {child_display} ({visual_count} visuals)"
                            icon = self._tree_bookmark_icon
                            tooltip_text = None
                        tags = ()
                    else:
                        text = f"  {child_display} (All visuals - not supported)"
                        icon = self._tree_warning_icon
                        tags = ('disabled',)
                        tooltip_text = None
                        disabled_count += 1

                    child_tree_id = self.bookmarks_treeview.insert(
                        group_tree_id, 'end',
                        text=text,
                        image=icon if icon else '',
                        tags=tags
                    )

                    # Store tooltip data for cross-page bookmarks
                    if tooltip_text:
                        self._bookmark_tooltip_data[child_tree_id] = tooltip_text

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
                    is_cross_page = bookmark_info.get('is_cross_page', False)
                    page_count = bookmark_info.get('page_count', 0)
                    page_display_names_list = bookmark_info.get('page_display_names', [])

                    # Format text and select icon based on mode and cross-page status
                    if mode == 'selected_visuals':
                        if is_cross_page:
                            # Cross-page bookmark: show bookmark alt icon and page count
                            text = f" {display_name} ({page_count} pages)"
                            icon = self._tree_crosspage_icon
                            tooltip_text = f"Cross-page: {', '.join(page_display_names_list)}"
                        else:
                            # Single-page bookmark: show bookmark icon and visual count
                            text = f" {display_name} ({visual_count} visuals)"
                            icon = self._tree_bookmark_icon
                            tooltip_text = None
                        tags = ()
                    else:
                        text = f" {display_name} (All visuals - not supported)"
                        icon = self._tree_warning_icon
                        tags = ('disabled',)
                        tooltip_text = None
                        disabled_count += 1

                    tree_id = self.bookmarks_treeview.insert(
                        '', 'end',
                        text=text,
                        image=icon if icon else '',
                        tags=tags
                    )

                    # Store tooltip data for cross-page bookmarks
                    if tooltip_text:
                        self._bookmark_tooltip_data[tree_id] = tooltip_text

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

            # Bind tooltip events for cross-page bookmarks
            self.bookmarks_treeview.bind('<Motion>', self._on_bookmark_tree_motion)
            self.bookmarks_treeview.bind('<Leave>', self._hide_bookmark_tooltip)
        finally:
            self._updating_ui = False
    
    def _update_selection_status(self):
        """Update the selection status and button states (common logic)"""
        if self._updating_ui:
            return

        # Clear the pending update reference
        self._pending_update = None

        if not self.bookmarks_treeview:
            return

        try:
            self._updating_ui = True

            # Count selected bookmarks (not groups, and only selectable ones)
            selected_items = self.bookmarks_treeview.selection()
            bookmarks_selected = sum(1 for item in selected_items
                                   if item in self._bookmark_tree_mapping
                                   and self._bookmark_tree_mapping[item]['type'] == 'bookmark'
                                   and self._bookmark_tree_mapping[item].get('selectable', True))

            # Count selected target pages (from checkbox selection set)
            targets_selected = len(getattr(self, '_selected_target_pages', set()))

            # Update status message with emojis for clarity
            colors = self._theme_manager.colors
            if not self.source_page_var.get():
                self.selection_label.config(text="📋 Select source page to begin", fg=colors['text_secondary'])
                if self.copy_button:
                    if hasattr(self.copy_button, 'set_enabled'):
                        self.copy_button.set_enabled(False)
                        self._copy_button_enabled = False
                    else:
                        self.copy_button.config(state=tk.DISABLED)
            elif bookmarks_selected == 0:
                self.selection_label.config(text="🔖 Select bookmarks to copy", fg=colors['text_secondary'])
                if self.copy_button:
                    if hasattr(self.copy_button, 'set_enabled'):
                        self.copy_button.set_enabled(False)
                        self._copy_button_enabled = False
                    else:
                        self.copy_button.config(state=tk.DISABLED)
            elif targets_selected == 0:
                self.selection_label.config(text="🎯 Select target pages", fg=colors['text_secondary'])
                if self.copy_button:
                    if hasattr(self.copy_button, 'set_enabled'):
                        self.copy_button.set_enabled(False)
                        self._copy_button_enabled = False
                    else:
                        self.copy_button.config(state=tk.DISABLED)
            else:
                self.selection_label.config(
                    text=f"✅ Ready: {bookmarks_selected} bookmark(s) → {targets_selected} page(s)",
                    fg=colors['success'])
                if self.copy_button:
                    if hasattr(self.copy_button, 'set_enabled'):
                        self.copy_button.set_enabled(True)
                        self._copy_button_enabled = True
                    else:
                        self.copy_button.config(state=tk.NORMAL)
        finally:
            self._updating_ui = False

    # ==================== TOOLTIP METHODS ====================

    def _on_bookmark_tree_motion(self, event):
        """Handle mouse motion over bookmark treeview - show tooltips for cross-page bookmarks."""
        # Get the item under the cursor
        item = self.bookmarks_treeview.identify_row(event.y)

        # Check if we have tooltip data for this item
        if hasattr(self, '_bookmark_tooltip_data') and item in self._bookmark_tooltip_data:
            tooltip_text = self._bookmark_tooltip_data[item]
            self._show_bookmark_tooltip(event, tooltip_text)
        else:
            self._hide_bookmark_tooltip()

    def _show_bookmark_tooltip(self, event, text):
        """Show tooltip near the cursor with theme-aware styling."""
        colors = self._theme_manager.colors

        # Theme-aware colors (matches Export Log tooltip)
        bg_color = colors.get('card_surface', '#ffffff')
        fg_color = colors.get('text_primary', '#1e1e1e')
        border_color = colors.get('border', '#e0e0e0')

        # Create tooltip window if it doesn't exist or was destroyed
        if not hasattr(self, '_bookmark_tooltip_window') or self._bookmark_tooltip_window is None:
            self._bookmark_tooltip_window = tk.Toplevel(self.bookmarks_treeview)
            self._bookmark_tooltip_window.wm_overrideredirect(True)
            self._bookmark_tooltip_window.wm_attributes('-topmost', True)

            # Create frame with border (like Tooltip class)
            self._bookmark_tooltip_frame = tk.Frame(
                self._bookmark_tooltip_window, bg=border_color, padx=1, pady=1
            )
            self._bookmark_tooltip_frame.pack(fill=tk.BOTH, expand=True)

            self._bookmark_tooltip_label = tk.Label(
                self._bookmark_tooltip_frame,
                text=text,
                background=bg_color,
                foreground=fg_color,
                font=('Segoe UI', 9),
                padx=8,
                pady=4
            )
            self._bookmark_tooltip_label.pack()
        else:
            # Update existing tooltip text and colors for theme changes
            self._bookmark_tooltip_frame.config(bg=border_color)
            self._bookmark_tooltip_label.config(text=text, background=bg_color, foreground=fg_color)

        # Position the tooltip near the cursor
        x = event.x_root + 15
        y = event.y_root + 10
        self._bookmark_tooltip_window.wm_geometry(f"+{x}+{y}")
        self._bookmark_tooltip_window.deiconify()

    def _hide_bookmark_tooltip(self, event=None):
        """Hide the tooltip window."""
        if hasattr(self, '_bookmark_tooltip_window') and self._bookmark_tooltip_window is not None:
            self._bookmark_tooltip_window.withdraw()

    def _update_bookmark_selection_theme(self, colors):
        """Update bookmark selection UI for theme changes."""
        is_dark = self._theme_manager.current_theme == 'dark'
        main_bg = '#0d0d1a' if is_dark else '#ffffff'
        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f5f5fa')

        # Treeview/listbox colors
        tree_bg = '#1e1e2e' if is_dark else '#ffffff'
        tree_fg = colors['text_primary']
        selected_bg = '#1a3a5c' if is_dark else '#e6f3ff'
        heading_bg = '#2a2a3e' if is_dark else '#f0f0f5'

        # Update main title frame (including icon) - uses section_bg like Analysis & Progress
        if hasattr(self, '_bookmark_title_frame') and self._bookmark_title_frame:
            self._bookmark_title_frame.config(bg=section_bg)
        if hasattr(self, '_bookmark_title_icon') and self._bookmark_title_icon:
            self._bookmark_title_icon.config(bg=section_bg)
        if hasattr(self, '_bookmark_title_label') and self._bookmark_title_label:
            self._bookmark_title_label.config(bg=section_bg, fg=colors['title_color'])

        # Update LabelFrame style for theme
        style = ttk.Style()
        style.configure('Borderless.TLabelframe.Label', background=main_bg)

        # Get option_bg for bordered inner sections (like Copy Mode)
        option_bg = colors.get('option_bg', '#f5f5f7')

        # Update page frame (column 1) - uses main_bg, no border
        if hasattr(self, '_bookmark_page_frame') and self._bookmark_page_frame:
            self._bookmark_page_frame.config(bg=main_bg)
        if hasattr(self, '_bookmark_page_header') and self._bookmark_page_header:
            self._bookmark_page_header.config(bg=main_bg)
        if hasattr(self, '_bookmark_page_icon') and self._bookmark_page_icon:
            self._bookmark_page_icon.config(bg=main_bg)
        if hasattr(self, '_bookmark_page_label') and self._bookmark_page_label:
            self._bookmark_page_label.config(bg=main_bg, fg=colors['title_color'])

        # Update bookmarks frame (column 2) - uses main_bg, no border
        if hasattr(self, '_bookmark_bookmarks_frame') and self._bookmark_bookmarks_frame:
            self._bookmark_bookmarks_frame.config(bg=main_bg)
        if hasattr(self, '_bookmark_title_row') and self._bookmark_title_row:
            self._bookmark_title_row.config(bg=main_bg)
        if hasattr(self, '_bookmark_bookmarks_icon') and self._bookmark_bookmarks_icon:
            self._bookmark_bookmarks_icon.config(bg=main_bg)
        if hasattr(self, 'bookmarks_instruction_label') and self.bookmarks_instruction_label:
            self.bookmarks_instruction_label.config(bg=main_bg, fg=colors['title_color'])
        if hasattr(self, 'bookmarks_helper_label') and self.bookmarks_helper_label:
            self.bookmarks_helper_label.config(bg=main_bg, fg=colors['text_secondary'])
        if hasattr(self, '_bookmark_list_frame') and self._bookmark_list_frame:
            self._bookmark_list_frame.config(bg=tree_bg, highlightbackground=colors['border'])

        # Update treeview styling
        style = ttk.Style()
        tree_style = 'BookmarkTree.Treeview'
        style.configure(tree_style,
                        background=tree_bg, foreground=tree_fg,
                        fieldbackground=tree_bg,
                        font=('Segoe UI', 9), rowheight=25,
                        relief='flat', borderwidth=0)
        style.configure(f"{tree_style}.Heading",
                        background=heading_bg, foreground=colors['text_primary'],
                        font=('Segoe UI', 9, 'bold'),
                        relief='groove', borderwidth=1,
                        padding=(8, 4))
        style.map(f"{tree_style}.Heading",
                  background=[('', heading_bg)])
        style.map(tree_style,
                  background=[('selected', selected_bg)],
                  foreground=[('selected', tree_fg)])

        # Update target frame (column 3) - uses main_bg, no border
        if hasattr(self, '_bookmark_target_frame') and self._bookmark_target_frame:
            self._bookmark_target_frame.config(bg=main_bg)
        if hasattr(self, '_bookmark_target_header') and self._bookmark_target_header:
            self._bookmark_target_header.config(bg=main_bg)
        if hasattr(self, '_bookmark_target_icon') and self._bookmark_target_icon:
            self._bookmark_target_icon.config(bg=main_bg)
        if hasattr(self, '_bookmark_target_hint') and self._bookmark_target_hint:
            self._bookmark_target_hint.config(bg=main_bg, fg=colors['title_color'])
        if hasattr(self, '_bookmark_target_list_frame') and self._bookmark_target_list_frame:
            self._bookmark_target_list_frame.config(bg=tree_bg, highlightbackground=colors['border'])

        # Update mode header (icon label) - uses option_bg not main_bg
        if hasattr(self, '_bookmark_mode_header') and self._bookmark_mode_header:
            try:
                self._bookmark_mode_header.config(bg=option_bg)
            except Exception:
                pass
            for child in self._bookmark_mode_header.winfo_children():
                child.config(bg=option_bg)
                if isinstance(child, tk.Label) and child.cget('text'):
                    child.config(fg=colors['text_primary'])

        # Update mode frame LEFT (row 1 col 0, uses option_bg with border)
        if hasattr(self, '_mode_frame') and self._mode_frame:
            self._mode_frame.config(bg=option_bg, highlightbackground=colors['border'])
            # Update mode_inner if it exists
            for child in self._mode_frame.winfo_children():
                if isinstance(child, tk.Frame):
                    child.config(bg=option_bg)
        # Note: _mode_right_frame no longer exists (unified mode frame)
        if hasattr(self, '_bookmark_mode_header') and self._bookmark_mode_header:
            self._bookmark_mode_header.config(bg=option_bg)
        if hasattr(self, '_bookmark_mode_icon') and self._bookmark_mode_icon:
            self._bookmark_mode_icon.config(bg=option_bg)
        if hasattr(self, '_bookmark_mode_title') and self._bookmark_mode_title:
            self._bookmark_mode_title.config(bg=option_bg, fg=colors['text_primary'])

        # Update per-page frame
        if hasattr(self, '_bookmark_perpage_frame') and self._bookmark_perpage_frame:
            self._bookmark_perpage_frame.config(bg=option_bg)
        if hasattr(self, '_perpage_label') and self._perpage_label:
            self._perpage_label.config(bg=option_bg, fg=colors['info'])

        # Update cross-page frame
        if hasattr(self, '_bookmark_crosspage_frame') and self._bookmark_crosspage_frame:
            self._bookmark_crosspage_frame.config(bg=option_bg)
        if hasattr(self, '_crosspage_label') and self._crosspage_label:
            is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"
            if is_cross_pbip:
                self._crosspage_label.config(bg=option_bg, fg=colors['warning'])
            else:
                self._crosspage_label.config(bg=option_bg, fg=colors['info'])

        # Update status frame and selection label (row 2, uses content_bg)
        if hasattr(self, '_bookmark_status_frame') and self._bookmark_status_frame:
            self._bookmark_status_frame.config(bg=main_bg)
        if hasattr(self, 'selection_label') and self.selection_label:
            self.selection_label.config(bg=main_bg, fg=colors['text_secondary'])

        # Update source page list
        if hasattr(self, '_bookmark_page_list_container') and self._bookmark_page_list_container:
            self._bookmark_page_list_container.config(
                bg=tree_bg,
                highlightbackground=colors['border']
            )
        if hasattr(self, '_bookmark_page_canvas') and self._bookmark_page_canvas:
            self._bookmark_page_canvas.config(bg=tree_bg)
        if hasattr(self, '_bookmark_page_inner_frame') and self._bookmark_page_inner_frame:
            self._bookmark_page_inner_frame.config(bg=tree_bg)

        # Update source page row widgets
        if hasattr(self, '_source_page_widgets'):
            radio_on = self._button_icons.get('radio-on')
            radio_off = self._button_icons.get('radio-off')
            for row_data in self._source_page_widgets:
                is_selected = row_data['idx'] == getattr(self, '_selected_source_page_idx', None)
                row_data['frame'].config(bg=tree_bg)
                row_data['text_label'].config(
                    bg=tree_bg,
                    fg=colors['title_color'] if is_selected else colors['text_primary']
                )
                # Update radio canvas
                radio_canvas = row_data.get('radio_canvas')
                if radio_canvas:
                    radio_canvas.config(bg=tree_bg)
                    radio_canvas.delete('all')
                    icon = radio_on if is_selected else radio_off
                    if icon:
                        radio_canvas.create_image(9, 9, image=icon, anchor=tk.CENTER, tags='radio')
                        radio_canvas._icon_ref = icon
                    else:
                        if is_selected:
                            radio_canvas.create_oval(2, 2, 16, 16, outline=colors['accent'], fill=colors['accent'], width=2, tags='radio')
                            radio_canvas.create_oval(6, 6, 12, 12, fill='white', outline='white', tags='radio_inner')
                        else:
                            radio_canvas.create_oval(2, 2, 16, 16, outline=colors['border'], fill=tree_bg, width=2, tags='radio')

        # Update target page list container
        if hasattr(self, '_bookmark_target_list_frame') and self._bookmark_target_list_frame:
            self._bookmark_target_list_frame.config(bg=tree_bg, highlightbackground=colors['border'])
        if hasattr(self, '_target_canvas') and self._target_canvas:
            self._target_canvas.config(bg=tree_bg)
        if hasattr(self, '_target_inner_frame') and self._target_inner_frame:
            self._target_inner_frame.config(bg=tree_bg)

        # Update separator line between Select All and pages
        if hasattr(self, '_target_separator') and self._target_separator:
            self._target_separator.config(bg=colors['border'])

        # Update target page row widgets
        if hasattr(self, '_target_page_widgets'):
            for row_data in self._target_page_widgets:
                is_select_all = row_data.get('is_select_all', False)
                is_all_selected = self._is_all_targets_selected()
                is_partial = self._is_partial_targets_selected() if hasattr(self, '_is_partial_targets_selected') else False

                if is_select_all:
                    is_selected = is_all_selected
                else:
                    is_selected = row_data['idx'] in getattr(self, '_selected_target_pages', set())

                row_data['frame'].config(bg=tree_bg)
                row_data['icon_label'].config(bg=tree_bg)  # Update checkbox label background
                row_data['text_label'].config(
                    bg=tree_bg,
                    fg=colors['title_color'] if is_selected else colors['text_primary']
                )
                # Update checkbox icon - use dark variants in dark mode, handle partial state for Select All
                if is_select_all and is_partial:
                    checkbox_icon = self._button_icons.get('box-partial-dark' if is_dark else 'box-partial')
                elif is_selected:
                    checkbox_icon = self._button_icons.get('box-checked-dark' if is_dark else 'box-checked')
                else:
                    checkbox_icon = self._button_icons.get('box-dark' if is_dark else 'box')
                if checkbox_icon:
                    row_data['icon_label'].configure(image=checkbox_icon)
                    row_data['icon_label']._icon_ref = checkbox_icon  # Prevent garbage collection

        # Update SVG radio canvases (use option_bg for Copy Mode section)
        if hasattr(self, '_perpage_radio_canvas') and self._perpage_radio_canvas:
            self._perpage_radio_canvas.config(bg=option_bg)
            self._update_svg_radio_bookmark(self._perpage_radio_canvas, colors)
        if hasattr(self, '_crosspage_radio_canvas') and self._crosspage_radio_canvas:
            self._crosspage_radio_canvas.config(bg=option_bg)
            self._update_svg_radio_bookmark(self._crosspage_radio_canvas, colors)
        if hasattr(self, '_perpage_text') and self._perpage_text:
            self._perpage_text.config(bg=option_bg)
        if hasattr(self, '_crosspage_text') and self._crosspage_text:
            self._crosspage_text.config(bg=option_bg)
        # Update Per-Page/Cross-Page text colors based on selection state
        if hasattr(self, '_update_bookmark_mode_colors'):
            self._update_bookmark_mode_colors()

    # ==================== NEW HELPER METHODS ====================

    def _create_source_page_row(self, parent, idx: int, page: dict, colors: dict, is_dark: bool):
        """Create a single source page row (radio-style single selection)"""
        bg_color = '#1e1e2e' if is_dark else '#ffffff'
        is_selected = idx == getattr(self, '_selected_source_page_idx', None)

        # Row frame
        row_frame = tk.Frame(parent, bg=bg_color, cursor='hand2')
        row_frame.pack(fill=tk.X, pady=(0, 2))

        # Radio indicator (circular)
        radio_on = self._button_icons.get('radio-on')
        radio_off = self._button_icons.get('radio-off')

        radio_canvas = tk.Canvas(row_frame, width=18, height=18, bg=bg_color,
                                  highlightthickness=0, cursor='hand2')
        radio_canvas.pack(side=tk.LEFT, padx=(8, 4), pady=4)

        # Draw initial state
        icon = radio_on if is_selected else radio_off
        if icon:
            radio_canvas.create_image(9, 9, image=icon, anchor=tk.CENTER, tags='radio')
            radio_canvas._icon_ref = icon  # Keep reference
        else:
            # Fallback: draw circle
            if is_selected:
                radio_canvas.create_oval(2, 2, 16, 16, outline=colors['accent'], fill=colors['accent'], width=2, tags='radio')
                radio_canvas.create_oval(6, 6, 12, 12, fill='white', outline='white', tags='radio_inner')
            else:
                radio_canvas.create_oval(2, 2, 16, 16, outline=colors['border'], fill=bg_color, width=2, tags='radio')

        # Text with bookmark count
        display_text = f"{page['display_name']} ({page['bookmark_count']})"
        fg_color = colors['primary'] if is_selected else colors['text_primary']

        text_label = tk.Label(row_frame, text=display_text, bg=bg_color, fg=fg_color,
                              font=('Segoe UI', 9), cursor='hand2', anchor='w', pady=4)
        text_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Bind click to select - all elements
        radio_canvas.bind('<Button-1>', lambda e, i=idx: self._select_source_page(i))
        text_label.bind('<Button-1>', lambda e, i=idx: self._select_source_page(i))
        row_frame.bind('<Button-1>', lambda e, i=idx: self._select_source_page(i))

        # Hover underline effect
        def on_enter(e, lbl=text_label):
            lbl.configure(font=('Segoe UI', 9, 'underline'))
        def on_leave(e, lbl=text_label):
            lbl.configure(font=('Segoe UI', 9))
        text_label.bind('<Enter>', on_enter)
        text_label.bind('<Leave>', on_leave)

        # Store widget references
        self._source_page_widgets.append({
            'frame': row_frame,
            'text_label': text_label,
            'radio_canvas': radio_canvas,
            'idx': idx,
            'page': page
        })

    def _select_source_page(self, idx: int):
        """Select a source page from the list"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.current_theme == 'dark'
        bg_color = '#1e1e2e' if is_dark else '#ffffff'

        # Update selection state
        self._selected_source_page_idx = idx

        # Get radio icons
        radio_on = self._button_icons.get('radio-on')
        radio_off = self._button_icons.get('radio-off')

        # Update visual state of all rows
        for row_data in self._source_page_widgets:
            is_selected = row_data['idx'] == idx
            row_data['text_label'].config(
                fg=colors['title_color'] if is_selected else colors['text_primary']
            )

            # Update radio indicator
            radio_canvas = row_data.get('radio_canvas')
            if radio_canvas:
                radio_canvas.delete('all')
                icon = radio_on if is_selected else radio_off
                if icon:
                    radio_canvas.create_image(9, 9, image=icon, anchor=tk.CENTER, tags='radio')
                    radio_canvas._icon_ref = icon
                else:
                    # Fallback: draw circle
                    if is_selected:
                        radio_canvas.create_oval(2, 2, 16, 16, outline=colors['accent'], fill=colors['accent'], width=2, tags='radio')
                        radio_canvas.create_oval(6, 6, 12, 12, fill='white', outline='white', tags='radio_inner')
                    else:
                        radio_canvas.create_oval(2, 2, 16, 16, outline=colors['border'], fill=bg_color, width=2, tags='radio')

        # Set the source_page_var for compatibility
        page = self.available_pages[idx]
        self.source_page_var.set(f"{page['display_name']} ({page['bookmark_count']} bookmarks)")

        # Trigger the change handler
        self._on_source_page_change()

    def _create_svg_radio_bookmark(self, parent, variable, value, colors, bg_color):
        """Create a custom SVG radio button for bookmark mode selection"""
        radio_on = self._button_icons.get('radio-on')
        radio_off = self._button_icons.get('radio-off')

        # Create canvas for radio button
        canvas = tk.Canvas(parent, width=18, height=18, bg=bg_color,
                          highlightthickness=0, cursor='hand2')

        # Store references
        canvas._variable = variable
        canvas._value = value
        canvas._icon_on = radio_on
        canvas._icon_off = radio_off

        # Draw initial state
        self._update_svg_radio_bookmark(canvas, colors)

        # Bind click event
        canvas.bind('<Button-1>', lambda e: self._on_svg_radio_click_bookmark(canvas))

        # Trace variable changes
        variable.trace_add('write', lambda *args: self._update_svg_radio_bookmark(canvas, self._theme_manager.colors))

        return canvas

    def _update_svg_radio_bookmark(self, canvas, colors):
        """Update SVG radio button display for bookmark mode"""
        # Check if canvas still exists (may have been destroyed when section was recreated)
        try:
            if not canvas.winfo_exists():
                return
        except tk.TclError:
            return  # Canvas was destroyed
        canvas.delete('all')
        is_selected = canvas._variable.get() == canvas._value
        icon = canvas._icon_on if is_selected else canvas._icon_off

        if icon:
            canvas.create_image(9, 9, image=icon, anchor=tk.CENTER)
        else:
            # Fallback: draw circle
            canvas_bg = canvas.cget('bg')
            if is_selected:
                canvas.create_oval(2, 2, 16, 16, outline=colors['accent'], fill=colors['accent'], width=2)
                canvas.create_oval(6, 6, 12, 12, fill='white', outline='white')
            else:
                canvas.create_oval(2, 2, 16, 16, outline=colors['border'], fill=canvas_bg, width=2)

    def _on_svg_radio_click_bookmark(self, canvas):
        """Handle SVG radio button click for bookmark mode"""
        # Cross-Page is now supported for both same-PBIP and cross-PBIP modes
        canvas._variable.set(canvas._value)
        self._update_bookmark_mode_colors()

    def _select_bookmark_mode(self, mode: str):
        """Select bookmark copy mode and update text colors"""
        # Cross-Page is now supported for both same-PBIP and cross-PBIP modes
        self.bookmark_copy_mode.set(mode)
        self._update_bookmark_mode_colors()

    def _update_bookmark_mode_colors(self):
        """Update Per-Page/Cross-Page text colors based on selection"""
        colors = self._theme_manager.colors
        mode = self.bookmark_copy_mode.get()

        # Update Per-Page text color
        if hasattr(self, '_perpage_text') and self._perpage_text:
            fg = colors['title_color'] if mode == "perpage" else colors['text_primary']
            self._perpage_text.configure(fg=fg)

        # Update Cross-Page text color (enabled for both same-PBIP and cross-PBIP)
        if hasattr(self, '_crosspage_text') and self._crosspage_text:
            fg = colors['title_color'] if mode == "crosspage" else colors['text_primary']
            self._crosspage_text.configure(fg=fg)

    def _populate_target_pages_modern(self, exclude_page_name: str = None):
        """Populate target pages with modern checkbox list (like Diagram Selection)"""
        if not hasattr(self, '_target_inner_frame') or not self._target_inner_frame:
            return

        # Determine which pages to use based on destination mode
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"

        if is_cross_pbip:
            # Use TARGET report's pages (from target_analysis_results)
            if not hasattr(self, 'target_analysis_results') or not self.target_analysis_results:
                self.log_message("   ⚠️ Target report not analyzed yet - cannot populate target pages")
                return

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

        # Clear existing widgets
        for widget in self._target_inner_frame.winfo_children():
            widget.destroy()
        self._target_page_widgets = []
        self._selected_target_pages = set()

        # Get theme colors
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.current_theme == 'dark'

        # Store pages for later reference
        self._target_pages_data = []

        # Add "Select All" row first
        self._create_target_page_row(self._target_inner_frame, -1, "(Select All)",
                                     colors, is_dark, is_select_all=True)

        # Add separator line between Select All and pages (like Layout Optimizer)
        # Use padx=(3, 3) for equal padding on both sides to span full width
        self._target_separator = tk.Frame(self._target_inner_frame, height=1,
                                          bg=colors.get('border', '#e0e0e0'))
        self._target_separator.pack(fill=tk.X, padx=(3, 3), pady=(4, 8))

        # Add all pages except the excluded one
        pages_added = 0
        for idx, page in enumerate(pages_to_use):
            # In cross-PBIP mode, don't exclude any pages
            # In same-PBIP mode, exclude the source page
            should_add = is_cross_pbip or (exclude_page_name is None or page['name'] != exclude_page_name)

            if should_add:
                self._target_pages_data.append(page)
                self._create_target_page_row(self._target_inner_frame, pages_added,
                                            page['display_name'], colors, is_dark)
                pages_added += 1

        self.log_message(f"   📄 Target pages available: {pages_added}")

    def _create_target_page_row(self, parent, idx: int, name: str, colors: dict,
                                is_dark: bool, is_select_all: bool = False):
        """Create a single target page row with checkbox (like Diagram Selection)"""
        bg_color = '#1e1e2e' if is_dark else '#ffffff'

        # Row frame
        row_frame = tk.Frame(parent, bg=bg_color)
        row_frame.pack(fill=tk.X, pady=(0, 2))

        # Checkbox icon (clickable) - use dark variants in dark mode, partial for Select All
        if is_select_all:
            is_all_selected = self._is_all_targets_selected()
            is_partial = self._is_partial_targets_selected()
        else:
            is_all_selected = idx in self._selected_target_pages
            is_partial = False

        if is_select_all and is_partial:
            checkbox_icon = self._button_icons.get('box-partial-dark' if is_dark else 'box-partial')
        elif is_all_selected:
            checkbox_icon = self._button_icons.get('box-checked-dark' if is_dark else 'box-checked')
        else:
            checkbox_icon = self._button_icons.get('box-dark' if is_dark else 'box')

        icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
        if checkbox_icon:
            icon_label.configure(image=checkbox_icon)
            icon_label._icon_ref = checkbox_icon
        else:
            icon_label.configure(text="☑️" if is_all_selected else "☐", font=('Segoe UI', 10))
        icon_label.pack(side=tk.LEFT, padx=(8, 10))

        # Text - partial shows as highlighted
        is_highlighted = is_all_selected or (is_select_all and is_partial)
        fg_color = colors['primary'] if is_highlighted else colors['text_primary']
        font_weight = 'bold' if is_select_all else 'normal'

        text_label = tk.Label(row_frame, text=name, bg=bg_color, fg=fg_color,
                              font=('Segoe UI', 9, font_weight), cursor='hand2', anchor='w')
        text_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Bind clicks
        if is_select_all:
            icon_label.bind('<Button-1>', lambda e: self._toggle_target_select_all())
            text_label.bind('<Button-1>', lambda e: self._toggle_target_select_all())
        else:
            icon_label.bind('<Button-1>', lambda e, i=idx: self._toggle_target_page_selection(i))
            text_label.bind('<Button-1>', lambda e, i=idx: self._toggle_target_page_selection(i))

        # Hover underline effect
        if is_select_all:
            def on_enter(e, lbl=text_label):
                lbl.configure(font=('Segoe UI', 9, 'bold underline'))
            def on_leave(e, lbl=text_label):
                lbl.configure(font=('Segoe UI', 9, 'bold'))
        else:
            def on_enter(e, lbl=text_label):
                lbl.configure(font=('Segoe UI', 9, 'underline'))
            def on_leave(e, lbl=text_label):
                lbl.configure(font=('Segoe UI', 9))
        text_label.bind('<Enter>', on_enter)
        text_label.bind('<Leave>', on_leave)

        # Store widget references
        self._target_page_widgets.append({
            'frame': row_frame,
            'icon_label': icon_label,
            'text_label': text_label,
            'idx': idx,
            'is_select_all': is_select_all
        })

    def _toggle_target_page_selection(self, idx: int):
        """Toggle selection state of a single target page"""
        if idx in self._selected_target_pages:
            self._selected_target_pages.discard(idx)
        else:
            self._selected_target_pages.add(idx)

        self._update_all_target_rows()
        self._update_selection_status()

    def _toggle_target_select_all(self):
        """Toggle all target page selections"""
        if self._is_all_targets_selected():
            # Deselect all
            self._selected_target_pages.clear()
        else:
            # Select all
            num_pages = len(self._target_pages_data) if hasattr(self, '_target_pages_data') else 0
            self._selected_target_pages = set(range(num_pages))

        self._update_all_target_rows()
        self._update_selection_status()

    def _is_all_targets_selected(self) -> bool:
        """Check if all target pages are selected"""
        num_pages = len(self._target_pages_data) if hasattr(self, '_target_pages_data') else 0
        return len(self._selected_target_pages) == num_pages and num_pages > 0

    def _is_partial_targets_selected(self) -> bool:
        """Check if some but not all target pages are selected"""
        num_pages = len(self._target_pages_data) if hasattr(self, '_target_pages_data') else 0
        num_selected = len(self._selected_target_pages)
        return num_selected > 0 and num_selected < num_pages

    def _update_all_target_rows(self):
        """Update all target page row visuals"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        for row_data in self._target_page_widgets:
            idx = row_data['idx']
            is_select_all = row_data.get('is_select_all', False)

            # Determine state for Select All vs regular rows
            if is_select_all:
                is_all_selected = self._is_all_targets_selected()
                is_partial = self._is_partial_targets_selected()
            else:
                is_all_selected = idx in self._selected_target_pages
                is_partial = False

            # Update icon - use dark variants in dark mode, partial icon for Select All
            if is_select_all and is_partial:
                checkbox_icon = self._button_icons.get('box-partial-dark' if is_dark else 'box-partial')
            elif is_all_selected:
                checkbox_icon = self._button_icons.get('box-checked-dark' if is_dark else 'box-checked')
            else:
                checkbox_icon = self._button_icons.get('box-dark' if is_dark else 'box')
            if checkbox_icon:
                row_data['icon_label'].configure(image=checkbox_icon)
                row_data['icon_label']._icon_ref = checkbox_icon
            else:
                row_data['icon_label'].configure(text="☑️" if is_all_selected else "☐")

            # Update text color - partial shows as selected (cyan)
            is_highlighted = is_all_selected or (is_select_all and is_partial)
            row_data['text_label'].config(
                fg=colors['title_color'] if is_highlighted else colors['text_primary']
            )
