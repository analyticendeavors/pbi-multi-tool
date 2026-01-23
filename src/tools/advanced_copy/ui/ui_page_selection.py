"""
Page Selection Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Handles page selection UI for full page copy mode.
Modern design with custom checkbox icons and themed spinbox.
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Any, Optional
from pathlib import Path

try:
    from PIL import Image, ImageTk
    import io
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False


class ThemedSpinbox(tk.Frame):
    """
    Custom spinbox with themed up/down arrow buttons matching scrollbar aesthetic.
    Uses canvas-drawn arrows for consistent theming.
    """

    def __init__(self, parent, from_: int = 1, to: int = 5, textvariable=None,
                 colors: dict = None, width: int = 5, **kwargs):
        # Use card_surface for parent frame background to blend with card container
        card_bg = colors.get('card_surface', '#ffffff') if colors else '#ffffff'
        super().__init__(parent, bg=card_bg)

        self.colors = colors or {}
        self._from = from_
        self._to = to
        self._var = textvariable

        # Entry for value display - use surface color for input field
        surface_color = self.colors.get('surface', self.colors.get('card_surface', '#ffffff'))
        self._entry = tk.Entry(
            self, width=width, justify='center',
            font=('Segoe UI', 10),
            bg=surface_color,
            fg=self.colors.get('text_primary', '#000000'),
            relief='flat', bd=1,
            highlightbackground=self.colors.get('border', '#cccccc'),
            highlightthickness=1,
            state='readonly',
            readonlybackground=surface_color
        )
        self._entry.pack(side=tk.LEFT, fill=tk.Y)

        # Display current value
        if self._var:
            self._update_display()
            self._var.trace_add('write', lambda *args: self._update_display())

        # Arrow buttons frame (thin border between entry and arrows)
        arrow_frame = tk.Frame(self, bg=self.colors.get('border', '#cccccc'))
        arrow_frame.pack(side=tk.LEFT, fill=tk.Y)
        self._arrow_frame = arrow_frame  # Store for theme updates

        # Arrow button colors
        btn_bg = self.colors.get('card_surface', '#f0f0f0')
        arrow_color = self.colors.get('text_secondary', '#666666')
        hover_bg = self.colors.get('accent', '#00a8a8')

        # Up arrow button
        self._up_btn = tk.Canvas(arrow_frame, width=18, height=12, bg=btn_bg,
                                 highlightthickness=0, cursor='hand2')
        self._up_btn.pack(side=tk.TOP)
        self._draw_arrow(self._up_btn, 'up', arrow_color)
        self._up_btn.bind('<Button-1>', lambda e: self._increment())
        self._up_btn.bind('<Enter>', lambda e: self._on_btn_hover(self._up_btn, True))
        self._up_btn.bind('<Leave>', lambda e: self._on_btn_hover(self._up_btn, False))

        # Separator line
        sep = tk.Frame(arrow_frame, height=1, bg=self.colors.get('border', '#cccccc'))
        sep.pack(fill=tk.X)

        # Down arrow button
        self._down_btn = tk.Canvas(arrow_frame, width=18, height=12, bg=btn_bg,
                                   highlightthickness=0, cursor='hand2')
        self._down_btn.pack(side=tk.TOP)
        self._draw_arrow(self._down_btn, 'down', arrow_color)
        self._down_btn.bind('<Button-1>', lambda e: self._decrement())
        self._down_btn.bind('<Enter>', lambda e: self._on_btn_hover(self._down_btn, True))
        self._down_btn.bind('<Leave>', lambda e: self._on_btn_hover(self._down_btn, False))

        # Store for theme updates
        self._arrow_color = arrow_color
        self._btn_bg = btn_bg
        self._hover_bg = hover_bg

    def _draw_arrow(self, canvas: tk.Canvas, direction: str, color: str):
        """Draw an arrow on the canvas"""
        canvas.delete('arrow')
        w, h = 18, 12
        if direction == 'up':
            points = [w//2, 3, w-4, h-3, 4, h-3]
        else:
            points = [4, 3, w-4, 3, w//2, h-3]
        canvas.create_polygon(points, fill=color, outline='', tags='arrow')

    def _on_btn_hover(self, canvas: tk.Canvas, hover: bool):
        """Handle button hover state"""
        if hover:
            canvas.configure(bg=self._hover_bg)
            self._draw_arrow(canvas, 'up' if canvas == self._up_btn else 'down', '#ffffff')
        else:
            canvas.configure(bg=self._btn_bg)
            self._draw_arrow(canvas, 'up' if canvas == self._up_btn else 'down', self._arrow_color)

    def _increment(self):
        """Increase value"""
        if self._var:
            val = self._var.get()
            if val < self._to:
                self._var.set(val + 1)

    def _decrement(self):
        """Decrease value"""
        if self._var:
            val = self._var.get()
            if val > self._from:
                self._var.set(val - 1)

    def _update_display(self):
        """Update entry display with current value"""
        self._entry.configure(state='normal')
        self._entry.delete(0, tk.END)
        self._entry.insert(0, str(self._var.get() if self._var else ''))
        self._entry.configure(state='readonly')

    def update_theme(self, colors: dict):
        """Update colors for theme change"""
        self.colors = colors
        self._btn_bg = colors.get('card_surface', '#f0f0f0')
        self._arrow_color = colors.get('text_secondary', '#666666')
        self._hover_bg = colors.get('accent', '#00a8a8')

        # Use card_surface for parent frame to blend with card container
        card_bg = colors.get('card_surface', '#ffffff')
        self.configure(bg=card_bg)

        # Entry uses surface color for the input field
        self._entry.configure(
            bg=colors.get('surface', colors.get('card_surface', '#ffffff')),
            fg=colors.get('text_primary', '#000000'),
            highlightbackground=colors.get('border', '#cccccc'),
            readonlybackground=colors.get('surface', colors.get('card_surface', '#ffffff'))
        )
        # Update arrow frame border
        if hasattr(self, '_arrow_frame'):
            self._arrow_frame.configure(bg=colors.get('border', '#cccccc'))

        self._up_btn.configure(bg=self._btn_bg)
        self._down_btn.configure(bg=self._btn_bg)
        self._draw_arrow(self._up_btn, 'up', self._arrow_color)
        self._draw_arrow(self._down_btn, 'down', self._arrow_color)


class PageSelectionMixin:
    """
    Mixin for page selection functionality.

    Methods extracted from AdvancedCopyTab:
    - _setup_page_selection()
    - _show_page_selection_ui()
    - _show_full_page_selection_ui()
    - _hide_page_selection_ui()
    - _adjust_window_height()
    - _on_page_selection_change()
    """

    def _setup_page_selection(self):
        """Setup page selection section (initially hidden)"""
        colors = self._theme_manager.colors

        # Load checkbox icons for page selection - themed for light/dark mode
        self._load_page_checkbox_icons()
        self._page_file_icon = self._load_icon_for_button('file', size=16)

        # Create section header with file icon
        page_header = self.create_section_header(self.frame, "Page Selection", "file")[0]

        self.pages_frame = ttk.LabelFrame(self.frame, labelwidget=page_header,
                                          style='Section.TLabelframe', padding="12")
        # Will be shown after analysis

        # Store mapping from tree item IDs to page indices
        self._page_tree_mapping = {}
        self._page_selection_vars = {}  # Track selection state
        self._copies_spinbox = None  # Store spinbox reference for theme updates

    def _load_page_checkbox_icons(self):
        """Load themed checkbox SVG icons for checked and unchecked states."""
        is_dark = self._theme_manager.is_dark

        # Select icon names based on theme
        box_name = 'box-dark' if is_dark else 'box'
        checked_name = 'box-checked-dark' if is_dark else 'box-checked'

        # Load icons using base class method
        self._page_checkbox_unchecked = self._load_icon_for_button(box_name, size=16)
        self._page_checkbox_checked = self._load_icon_for_button(checked_name, size=16)

    def _show_page_selection_ui(self, pages_with_bookmarks: List[Dict[str, Any]]):
        """Show page selection UI after analysis - routes based on content mode"""
        if self.copy_content_mode.get() == "full_page":
            self._show_full_page_selection_ui(pages_with_bookmarks)
        else:  # bookmark_visual mode
            self._show_bookmark_mode_selection_ui(pages_with_bookmarks)

    def _show_full_page_selection_ui(self, pages_with_bookmarks: List[Dict[str, Any]]):
        """Show page selection UI for full page copy mode with modern table design"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.current_theme == 'dark'

        # Show the pages frame
        self.pages_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        # Adjust window height to accommodate page selection
        self._adjust_window_height(True)

        # Clear existing content
        for widget in self.pages_frame.winfo_children():
            widget.destroy()

        # Use colors['background'] for labels - matches Section.TFrame background
        # Section.TFrame uses colors['background'] (white in light, #0d0d1a in dark)
        frame_bg = colors['background']
        # Canvas bg for rounded button corners - must match frame background exactly
        canvas_bg = '#0d0d1a' if is_dark else '#ffffff'

        # Content frame using ttk.Frame with Section.TFrame style (matches Analysis & Progress)
        # Padding="15" matches Analysis & Progress section
        # FIXED HEIGHT: Only sticky W, E (not N, S) - prevents vertical expansion when window resizes
        # This makes Page Selection stay fixed while Analysis & Progress grows/shrinks
        content_frame = ttk.Frame(self.pages_frame, style='Section.TFrame', padding="15")
        content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Full width: configure columns for horizontal expansion
        self.pages_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(0, weight=1)  # Instruction panel (left) takes less space
        content_frame.columnconfigure(1, weight=3)  # Page table (right) takes more space
        self._page_content_frame = content_frame  # Store for theme updates
        self._page_frame_bg = frame_bg  # Store for theme updates
        self._page_canvas_bg = canvas_bg  # Store for button theme updates

        # Instruction panel on LEFT side - use ttk.Frame for consistent styling
        instruction_frame = ttk.Frame(content_frame, style='Section.TFrame')
        instruction_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 15))
        self._instruction_frame = instruction_frame

        # Title for instruction panel - use frame_bg to match Section.TFrame
        title_label = tk.Label(instruction_frame, text="Selected Pages to Copy:",
                              font=('Segoe UI', 10, 'bold'),
                              bg=frame_bg, fg=colors['title_color'], anchor=tk.W)
        title_label.pack(fill=tk.X, pady=(0, 8))
        self._page_title_label = title_label

        # Helper text labels
        text_color = '#c0c0c0' if is_dark else colors['text_primary']
        helper_texts = [
            "Click a row to toggle selection",
            "Use Select All/None buttons for bulk selection",
            "Each selected page will be copied with all its visuals",
            "Bookmarks and page filters are preserved"
        ]
        self._instruction_labels = []
        for text in helper_texts:
            lbl = tk.Label(instruction_frame, text=f"• {text}",
                          font=('Segoe UI', 9),
                          bg=frame_bg, fg=text_color, anchor=tk.W,
                          wraplength=300, justify=tk.LEFT)
            lbl.pack(fill=tk.X, pady=2)
            self._instruction_labels.append(lbl)

        # Page table on RIGHT side - use ttk.Frame for consistent styling
        list_frame = ttk.Frame(content_frame, style='Section.TFrame')
        list_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        list_frame.columnconfigure(0, weight=1)
        self._page_list_frame = list_frame  # Store for theme updates

        # Import RoundedButton for styled buttons (need early for controls above table)
        from core.ui_base import RoundedButton

        # Controls row ABOVE the table - use ttk.Frame for consistent styling
        # Padding below buttons (15px) matches padding from top of frame to buttons
        controls_frame = ttk.Frame(list_frame, style='Section.TFrame')
        controls_frame.grid(row=0, column=0, pady=(0, 15), sticky=(tk.W, tk.E))
        self._controls_frame = controls_frame  # Store for theme updates

        # Select All / Select None buttons with proper rounded corners
        # Use button_secondary style for subtle appearance, radius=6 for nice rounding
        select_all_btn = RoundedButton(
            controls_frame, text="Select All",
            command=lambda: self._select_all_pages(True),
            bg=colors.get('button_secondary', colors.get('card_surface', '#1a1a2e')),
            hover_bg=colors.get('button_secondary_hover', colors.get('card_surface_hover', '#141424')),
            pressed_bg=colors.get('button_secondary_pressed', colors.get('card_surface_pressed', '#0e0e18')),
            fg=colors['text_primary'],
            width=80, height=28, radius=6, font=('Segoe UI', 9),
            canvas_bg=canvas_bg
        )
        select_all_btn.pack(side=tk.LEFT, padx=(0, 8))

        select_none_btn = RoundedButton(
            controls_frame, text="Select None",
            command=lambda: self._select_all_pages(False),
            bg=colors.get('button_secondary', colors.get('card_surface', '#1a1a2e')),
            hover_bg=colors.get('button_secondary_hover', colors.get('card_surface_hover', '#141424')),
            pressed_bg=colors.get('button_secondary_pressed', colors.get('card_surface_pressed', '#0e0e18')),
            fg=colors['text_primary'],
            width=90, height=28, radius=6, font=('Segoe UI', 9),
            canvas_bg=canvas_bg
        )
        select_none_btn.pack(side=tk.LEFT, padx=(0, 15))

        self._page_select_btns = [select_all_btn, select_none_btn]

        # Separator/spacer
        separator = tk.Frame(controls_frame, width=1, height=20, bg=colors['border'])
        separator.pack(side=tk.LEFT, padx=(0, 15))
        self._controls_separator = separator  # Store for theme updates

        # Number of copies label - use frame_bg to match Section.TFrame
        copies_label = tk.Label(controls_frame, text="Copies:",
                                font=('Segoe UI', 9, 'bold'),
                                bg=frame_bg,
                                fg=colors['text_primary'])
        copies_label.pack(side=tk.LEFT, padx=(0, 8))
        self._copies_label = copies_label  # Store for theme updates

        # Themed spinbox
        self._copies_spinbox = ThemedSpinbox(controls_frame, from_=1, to=5,
                                             textvariable=self.num_copies, colors=colors)
        self._copies_spinbox.pack(side=tk.LEFT, padx=(0, 8))

        # Description text (uses title_color - blue in dark, teal in light)
        desc_color = colors['title_color']
        copies_desc = tk.Label(controls_frame, text="(per selected page)",
                               font=('Segoe UI', 8),
                               bg=frame_bg,
                               fg=desc_color)
        copies_desc.pack(side=tk.LEFT)
        self._copies_desc_label = copies_desc  # Store for theme updates

        # Selection summary on the same row, far right - use frame_bg
        # Fixed width prevents text length changes from affecting table layout
        self.selection_label = tk.Label(controls_frame, text="All pages selected",
                                        font=('Segoe UI', 9),
                                        bg=frame_bg,
                                        fg=colors['text_secondary'],
                                        width=30,  # Fixed width to prevent layout shifts
                                        anchor=tk.E)  # Right-align text within fixed width
        self.selection_label.pack(side=tk.RIGHT)

        # Create tree container with subtle border and FIXED HEIGHT
        # Fixed height 140px - prevents section from expanding when window resizes
        tree_border = '#3a3a4a' if is_dark else '#d8d8e0'
        tree_container = tk.Frame(list_frame, bg=tree_border,
                                  highlightthickness=1, highlightbackground=tree_border,
                                  height=140)
        tree_container.grid(row=1, column=0, sticky=(tk.W, tk.E))
        tree_container.grid_propagate(False)  # CRITICAL: Enforce fixed height
        tree_container.columnconfigure(0, weight=1)  # Allow horizontal expansion
        tree_container.rowconfigure(0, weight=1)  # Let treeview fill the fixed height
        self._page_tree_container = tree_container

        # Configure treeview style for page selection - modern flat design
        style = ttk.Style()
        tree_style = "PageSelection.Treeview"

        if is_dark:
            tree_bg = '#1e1e2e'
            tree_fg = '#e0e0e0'
            tree_field_bg = '#1e1e2e'
            heading_bg = colors.get('section_bg', '#1a1a2a')
            heading_fg = '#e0e0e0'
            header_separator = '#0d0d1a'
            selected_bg = '#1a3a5c'  # Modern deep blue selection
            selected_fg = '#ffffff'
        else:
            tree_bg = '#ffffff'
            tree_fg = '#333333'
            tree_field_bg = '#ffffff'
            heading_bg = colors.get('section_bg', '#f5f5fa')
            heading_fg = '#333333'
            header_separator = '#ffffff'
            selected_bg = '#e6f3ff'  # Modern light blue selection
            selected_fg = '#1a1a2e'

        # Modern flat treeview style with extra row padding
        # indent=0 removes tree indentation so checkbox column stays fixed width
        style.configure(tree_style,
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=tree_field_bg,
                        font=('Segoe UI', 9),
                        rowheight=28,
                        indent=0,
                        relief='flat',
                        borderwidth=0,
                        bordercolor=tree_bg,
                        lightcolor=tree_bg,
                        darkcolor=tree_bg)

        # Modern header with groove relief
        style.configure(f"{tree_style}.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        font=('Segoe UI', 9, 'bold'),
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        padding=(8, 4))

        # Remove internal borders via layout
        style.layout(tree_style, [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

        style.map(tree_style,
                  background=[('selected', selected_bg)],
                  foreground=[('selected', selected_fg)])

        # Heading map for active/pressed states - maintains groove relief on interaction
        # Include ('', heading_bg) for Python 3.13+ compatibility
        style.map(f"{tree_style}.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])

        # Store colors for row highlighting (matching Table Column Widths design)
        self._page_selected_bg = selected_bg
        self._page_selected_fg = selected_fg
        self._page_unselected_bg = tree_bg
        self._page_unselected_fg = tree_fg
        self._hovered_item = None  # Track currently hovered item

        # Create treeview - compact height for better layout
        # Use separate columns: #0 for checkbox icon, page_name for text, then other data columns
        self.pages_listbox = ttk.Treeview(tree_container, height=2, selectmode='none',
                                          style=tree_style, columns=('page_name', 'bookmarks', 'filters', 'page_id'))
        self.pages_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure columns - Checkbox (#0), Page Name, Bookmarks count, Page-Level Filters count, and Page ID
        # #0 column is just for checkbox icon - fixed width to center under header
        # Width 44 gives room for 16px icon to center properly
        self.pages_listbox.heading('#0', text='', anchor=tk.CENTER)
        self.pages_listbox.column('#0', width=44, minwidth=44, stretch=False, anchor=tk.CENTER)

        # Page Name in its own column
        self.pages_listbox.heading('page_name', text='Page Name', anchor=tk.CENTER)
        self.pages_listbox.column('page_name', width=160, minwidth=100, anchor=tk.CENTER)

        self.pages_listbox.heading('bookmarks', text='Bookmarks', anchor=tk.CENTER)
        self.pages_listbox.heading('filters', text='Page-Level Filters', anchor=tk.CENTER)
        self.pages_listbox.heading('page_id', text='Page ID', anchor=tk.CENTER)
        self.pages_listbox.column('bookmarks', width=70, minwidth=55, anchor=tk.CENTER)
        self.pages_listbox.column('filters', width=100, minwidth=80, anchor=tk.CENTER)
        self.pages_listbox.column('page_id', width=140, minwidth=100, anchor=tk.CENTER)

        # Modern themed scrollbar with auto-hide
        from core.ui_base import ThemedScrollbar
        scrollbar = ThemedScrollbar(tree_container,
                                    command=self.pages_listbox.yview,
                                    theme_manager=self._theme_manager,
                                    width=12,
                                    auto_hide=True)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.pages_listbox.configure(yscrollcommand=scrollbar.set)
        self._page_scrollbar = scrollbar  # Store for theme updates

        # Populate treeview with pages
        self.available_pages = pages_with_bookmarks
        self._page_tree_mapping = {}
        self._page_selection_vars = {}

        for idx, page in enumerate(pages_with_bookmarks):
            # All pages start selected
            self._page_selection_vars[idx] = True

            # Use checkbox icons
            checkbox_icon = self._page_checkbox_checked if self._page_checkbox_checked else None

            # Get internal page ID (folder name) and filter count
            page_id = page.get('name', '')
            filter_count = page.get('filter_count', 0)

            # Values: (page_name, bookmarks, filters, page_id)
            # #0 column is just for checkbox icon, page name goes in page_name column
            if checkbox_icon:
                item_id = self.pages_listbox.insert('', 'end',
                                                    text='',  # #0 column empty, just icon
                                                    image=checkbox_icon,
                                                    values=(page['display_name'], str(page['bookmark_count']), str(filter_count), page_id),
                                                    tags=(str(idx), 'selected'))
            else:
                # Fallback without icon - use text checkbox in #0
                item_id = self.pages_listbox.insert('', 'end',
                                                    text='☑',
                                                    values=(page['display_name'], str(page['bookmark_count']), str(filter_count), page_id),
                                                    tags=(str(idx), 'selected'))

            self._page_tree_mapping[item_id] = idx

        # Configure tags for row highlighting (matching Table Column Widths design)
        # Hover = underline only (no background change)
        self.pages_listbox.tag_configure('hover', font=('Segoe UI', 9, 'underline'))
        # Selected = colored background + white text (checked rows)
        self.pages_listbox.tag_configure('selected', background=selected_bg, foreground=selected_fg)
        # Unselected = normal background (unchecked rows)
        self.pages_listbox.tag_configure('unselected', background=tree_bg, foreground=tree_fg)

        # Bind click events for checkbox toggle
        self.pages_listbox.bind('<Button-1>', self._on_page_tree_click)

        # Bind hover events for row highlighting
        self.pages_listbox.bind('<Motion>', self._on_page_tree_hover)
        self.pages_listbox.bind('<Leave>', self._on_page_tree_leave)

        # Enable copy button (all pages selected by default)
        self._update_page_selection_state()

    def _on_page_tree_click(self, event):
        """Handle click on page tree to toggle checkbox"""
        region = self.pages_listbox.identify_region(event.x, event.y)
        item = self.pages_listbox.identify_row(event.y)

        if not item:
            return

        # Toggle selection on any click within the row
        if item in self._page_tree_mapping:
            idx = self._page_tree_mapping[item]
            is_selected = self._page_selection_vars.get(idx, False)
            self._page_selection_vars[idx] = not is_selected

            # Update checkbox icon
            new_icon = self._page_checkbox_checked if not is_selected else self._page_checkbox_unchecked
            if new_icon:
                self.pages_listbox.item(item, image=new_icon)
            else:
                # Fallback text update - just checkbox in #0 column
                checkbox = "☑" if not is_selected else "☐"
                self.pages_listbox.item(item, text=checkbox)

            # Update row text color based on selection (like Table Column Widths)
            self._update_row_highlight(item, not is_selected)

            # Update selection state
            self._update_page_selection_state()

    def _on_page_tree_hover(self, event):
        """Handle mouse hover over treeview rows"""
        item = self.pages_listbox.identify_row(event.y)

        # If hovering over a different item
        if item != self._hovered_item:
            # Remove hover from previous item
            if self._hovered_item:
                current_tags = list(self.pages_listbox.item(self._hovered_item, 'tags'))
                if 'hover' in current_tags:
                    current_tags.remove('hover')
                    self.pages_listbox.item(self._hovered_item, tags=current_tags)

            # Add hover to new item
            if item:
                current_tags = list(self.pages_listbox.item(item, 'tags'))
                if 'hover' not in current_tags:
                    current_tags.append('hover')
                    self.pages_listbox.item(item, tags=current_tags)

            self._hovered_item = item

    def _on_page_tree_leave(self, event):
        """Handle mouse leaving treeview"""
        if self._hovered_item:
            current_tags = list(self.pages_listbox.item(self._hovered_item, 'tags'))
            if 'hover' in current_tags:
                current_tags.remove('hover')
                self.pages_listbox.item(self._hovered_item, tags=current_tags)
            self._hovered_item = None

    def _update_row_highlight(self, item, is_selected: bool):
        """Update row background/foreground based on selection state (like Table Column Widths)"""
        current_tags = list(self.pages_listbox.item(item, 'tags'))

        # Remove both selected and unselected tags first
        if 'selected' in current_tags:
            current_tags.remove('selected')
        if 'unselected' in current_tags:
            current_tags.remove('unselected')

        # Add appropriate tag based on selection state
        if is_selected:
            current_tags.append('selected')
        else:
            current_tags.append('unselected')

        self.pages_listbox.item(item, tags=current_tags)

    def _select_all_pages(self, select: bool):
        """Select or deselect all pages"""
        for item_id, idx in self._page_tree_mapping.items():
            self._page_selection_vars[idx] = select

            # Update checkbox icon
            icon = self._page_checkbox_checked if select else self._page_checkbox_unchecked
            if icon:
                self.pages_listbox.item(item_id, image=icon)
            else:
                # Fallback text - just checkbox in #0 column
                checkbox = "☑" if select else "☐"
                self.pages_listbox.item(item_id, text=checkbox)

            # Update row text color based on selection (like Table Column Widths)
            self._update_row_highlight(item_id, select)

        self._update_page_selection_state()

    def _update_page_selection_state(self):
        """Update selection summary and copy button state"""
        selected_count = sum(1 for v in self._page_selection_vars.values() if v)
        total_count = len(self._page_selection_vars)

        if selected_count == 0:
            self.selection_label.config(text="No pages selected")
            if self.copy_button:
                if hasattr(self.copy_button, 'set_enabled'):
                    self.copy_button.set_enabled(False)
                    self._copy_button_enabled = False
                else:
                    self.copy_button.config(state=tk.DISABLED)
        elif selected_count == 1:
            # Find the selected page name
            for idx, selected in self._page_selection_vars.items():
                if selected:
                    page_name = self.available_pages[idx]['display_name']
                    self.selection_label.config(text=f"Selected: {page_name}")
                    break
            if self.copy_button:
                if hasattr(self.copy_button, 'set_enabled'):
                    self.copy_button.set_enabled(True)
                    self._copy_button_enabled = True
                else:
                    self.copy_button.config(state=tk.NORMAL)
        elif selected_count == total_count:
            self.selection_label.config(text="All pages selected")
            if self.copy_button:
                if hasattr(self.copy_button, 'set_enabled'):
                    self.copy_button.set_enabled(True)
                    self._copy_button_enabled = True
                else:
                    self.copy_button.config(state=tk.NORMAL)
        else:
            self.selection_label.config(text=f"Selected: {selected_count} of {total_count} pages")
            if self.copy_button:
                if hasattr(self.copy_button, 'set_enabled'):
                    self.copy_button.set_enabled(True)
                    self._copy_button_enabled = True
                else:
                    self.copy_button.config(state=tk.NORMAL)

    def get_selected_page_indices(self) -> List[int]:
        """Get list of selected page indices (for compatibility with existing code)"""
        return [idx for idx, selected in self._page_selection_vars.items() if selected]

    def _hide_page_selection_ui(self):
        """Hide page selection UI"""
        if self.pages_frame:
            self.pages_frame.grid_remove()
            self._adjust_window_height(False)

    def _adjust_window_height(self, show_page_selection: bool):
        """Adjust window height based on page selection visibility

        Sets a FIXED height when showing page selection (not additive).
        This ensures the Execute Copy buttons at the bottom are fully visible.
        """
        try:
            if hasattr(self.main_app, 'root'):
                root = self.main_app.root

                # Define FIXED heights for different states
                base_height = 900  # Height before analysis (no page selection)
                post_analysis_height = 1185  # FIXED height after analysis (includes page selection + buttons + action area)

                # Add extra height when in cross-PBIP mode (target input showing)
                cross_pbip_adjustment = 55 if self.copy_destination_mode.get() == "cross_pbip" else 0

                if show_page_selection:
                    # Set FIXED post-analysis height (not additive!)
                    new_height = post_analysis_height + cross_pbip_adjustment
                else:
                    # Back to base height (no page selection)
                    new_height = base_height + cross_pbip_adjustment

                # Update window geometry
                current_geometry = root.geometry()
                parts = current_geometry.split('x')
                if len(parts) >= 2:
                    width = parts[0]
                    position = parts[1].split('+', 1)[1] if '+' in parts[1] else ''
                    new_geometry = f"{width}x{new_height}+{position}"
                    root.geometry(new_geometry)
        except Exception:
            pass

    def _on_page_selection_change(self, event=None):
        """Handle page selection changes - legacy compatibility method"""
        # Now handled by _update_page_selection_state()
        self._update_page_selection_state()

    def _update_page_selection_theme(self, colors: dict):
        """Update page selection UI colors for theme change"""
        is_dark = self._theme_manager.current_theme == 'dark'

        # Reload themed checkbox icons (light/dark variants)
        self._load_page_checkbox_icons()

        # Update existing treeview items with new checkbox icons
        if hasattr(self, 'pages_listbox') and self.pages_listbox and hasattr(self, '_page_selection_vars'):
            try:
                for item in self.pages_listbox.get_children():
                    if item in self._page_tree_mapping:
                        idx = self._page_tree_mapping[item]
                        is_selected = self._page_selection_vars.get(idx, False)
                        icon = self._page_checkbox_checked if is_selected else self._page_checkbox_unchecked
                        if icon:
                            self.pages_listbox.item(item, image=icon)
            except Exception:
                pass

        # Use colors['background'] for labels - matches Section.TFrame background
        frame_bg = colors['background']
        # Canvas bg for button corners
        canvas_bg = '#0d0d1a' if is_dark else '#ffffff'
        tree_border = '#3a3a4a' if is_dark else '#d8d8e0'

        # Note: content_frame, instruction_frame, list_frame, controls_frame all use
        # ttk.Frame with Section.TFrame style - background handled via ttk style system

        # Update tree container border
        if hasattr(self, '_page_tree_container') and self._page_tree_container:
            try:
                self._page_tree_container.configure(
                    bg=tree_border,
                    highlightbackground=tree_border
                )
            except Exception:
                pass

        # Update treeview style - modern flat design
        style = ttk.Style()
        tree_style = "PageSelection.Treeview"

        if is_dark:
            tree_bg = '#1e1e2e'
            tree_fg = '#e0e0e0'
            tree_field_bg = '#1e1e2e'
            heading_bg = colors.get('section_bg', '#1a1a2a')
            heading_fg = '#e0e0e0'
            header_separator = '#0d0d1a'
            selected_bg = '#1a3a5c'  # Modern deep blue selection
            selected_fg = '#ffffff'
        else:
            tree_bg = '#ffffff'
            tree_fg = '#333333'
            tree_field_bg = '#ffffff'
            heading_bg = colors.get('section_bg', '#f5f5fa')
            heading_fg = '#333333'
            header_separator = '#ffffff'
            selected_bg = '#e6f3ff'  # Modern light blue selection
            selected_fg = '#1a1a2e'

        style.configure(tree_style,
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=tree_field_bg,
                        font=('Segoe UI', 9),
                        rowheight=28,
                        relief='flat',
                        borderwidth=0)
        style.configure(f"{tree_style}.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        font=('Segoe UI', 9, 'bold'),
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        padding=(8, 4))

        # Heading map for active/pressed states - maintains groove relief on interaction
        # Include ('', heading_bg) for Python 3.13+ compatibility
        style.map(f"{tree_style}.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])

        style.map(tree_style,
                  background=[('selected', selected_bg)],
                  foreground=[('selected', selected_fg)])

        # Update selection colors for tag configuration (matching Table Column Widths design)
        if is_dark:
            selected_bg = '#1a3a5c'
            selected_fg = '#ffffff'
        else:
            selected_bg = '#e6f3ff'
            selected_fg = '#1a1a2e'

        # Store colors for row highlighting
        self._page_selected_bg = selected_bg
        self._page_selected_fg = selected_fg
        self._page_unselected_bg = tree_bg
        self._page_unselected_fg = tree_fg

        # Update tag colors on the treeview if it exists
        if hasattr(self, 'pages_listbox') and self.pages_listbox:
            try:
                # Hover = underline only (no background change)
                self.pages_listbox.tag_configure('hover', font=('Segoe UI', 9, 'underline'))
                # Selected = colored background + white text (checked rows)
                self.pages_listbox.tag_configure('selected', background=selected_bg, foreground=selected_fg)
                # Unselected = normal background (unchecked rows)
                self.pages_listbox.tag_configure('unselected', background=tree_bg, foreground=tree_fg)
            except Exception:
                pass

        # Update instruction labels - use frame_bg to match Section.TFrame
        if hasattr(self, '_page_title_label') and self._page_title_label:
            try:
                self._page_title_label.configure(bg=frame_bg, fg=colors['title_color'])
            except Exception:
                pass

        if hasattr(self, '_instruction_labels'):
            text_color = '#c0c0c0' if is_dark else colors['text_primary']
            for lbl in self._instruction_labels:
                try:
                    lbl.configure(bg=frame_bg, fg=text_color)
                except Exception:
                    pass

        # Update separator
        if hasattr(self, '_controls_separator') and self._controls_separator:
            try:
                self._controls_separator.configure(bg=colors['border'])
            except Exception:
                pass

        # Update select buttons with proper canvas_bg for rounded corners
        if hasattr(self, '_page_select_btns'):
            for btn in self._page_select_btns:
                try:
                    if hasattr(btn, 'update_colors'):
                        btn.update_colors(
                            bg=colors.get('button_secondary', colors.get('card_surface', '#1a1a2e')),
                            hover_bg=colors.get('button_secondary_hover', colors.get('card_surface_hover', '#141424')),
                            pressed_bg=colors.get('button_secondary_pressed', colors.get('card_surface_pressed', '#0e0e18')),
                            fg=colors['text_primary'],
                            canvas_bg=canvas_bg
                        )
                except Exception:
                    pass

        # Update copies label - use frame_bg
        if hasattr(self, '_copies_label') and self._copies_label:
            try:
                self._copies_label.configure(bg=frame_bg, fg=colors['text_primary'])
            except Exception:
                pass

        # Update copies description label (uses title_color - blue in dark, teal in light)
        if hasattr(self, '_copies_desc_label') and self._copies_desc_label:
            try:
                self._copies_desc_label.configure(bg=frame_bg, fg=colors['title_color'])
            except Exception:
                pass

        # Update themed spinbox
        if hasattr(self, '_copies_spinbox') and self._copies_spinbox:
            try:
                self._copies_spinbox.update_theme(colors)
            except Exception:
                pass

        # Update selection label - use frame_bg
        if hasattr(self, 'selection_label') and self.selection_label:
            try:
                self.selection_label.configure(
                    bg=frame_bg,
                    fg=colors['text_secondary']
                )
            except Exception:
                pass
