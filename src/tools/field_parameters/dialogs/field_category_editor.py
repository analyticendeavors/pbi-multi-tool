"""
Field Category Editor Dialog
Grid-based dialog for assigning category labels to fields with virtual scrolling,
drag-drop reordering, bulk editing, and column resizing.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Optional, Tuple, TYPE_CHECKING

from core.theme_manager import get_theme_manager
from core.ui_base import ThemedScrollbar, ThemedMessageBox, ThemedInputDialog, RoundedButton, ThemedContextMenu, Tooltip
from core.widgets import AutoHideScrollbar, ThemedCombobox
from tools.field_parameters.field_parameters_core import FieldItem, CategoryLevel
from tools.field_parameters.widgets import SVGCheckbox


class FieldCategoryEditorDialog(tk.Toplevel):
    """Dialog for assigning category labels to fields in a grid view"""

    # Virtual scrolling constants
    ROW_HEIGHT = 30  # Approximate height of each row
    BUFFER_ROWS = 10  # Extra rows to render above/below viewport for smoother scrolling
    VIRTUAL_SCROLL_THRESHOLD = 100  # Use virtual scrolling when item count exceeds this

    def __init__(self, parent, field_items: List[FieldItem], category_levels: List[CategoryLevel], on_save: callable):
        super().__init__(parent)
        self.withdraw()  # Hide immediately to prevent flicker before positioning
        self.field_items = field_items
        self.category_levels = category_levels
        self.on_save = on_save
        self._theme_manager = get_theme_manager()

        # Store category assignments: {field_id: {column_idx: label_str}}
        self.assignments: Dict[int, Dict[int, str]] = {}
        self._init_assignments()

        # Track selected rows for bulk edit
        self.selected_rows: set = set()
        self.last_clicked_row: Optional[int] = None  # For Shift+click range selection

        # Column widths in pixels (resizable)
        self.col_widths: List[int] = []
        self._init_column_widths()

        # Drag state for resizing columns
        self._resize_drag: Dict = {"active": False, "col_idx": None, "start_x": 0}

        # Drag state for row reordering
        self._row_drag: Dict = {
            "active": False,
            "start_y": 0,
            "field_id": None,
            "dragging": False,
            "pending_deselect": None
        }
        self._drop_indicator: Optional[tk.Frame] = None
        self._drop_position: Optional[int] = None
        self._drag_label: Optional[tk.Toplevel] = None

        # Virtual scrolling state
        self._virtual_mode = False
        self._visible_range = (0, 0)
        self._filtered_items: List[FieldItem] = []
        self._scroll_job = None

        self.title("Edit Field Categories")
        self.minsize(500, 300)
        self.resizable(True, True)
        self.transient(parent)

        # Set AE favicon
        try:
            from pathlib import Path
            base_path = Path(__file__).parent.parent.parent.parent
            favicon_path = base_path / "assets" / "favicon.ico"
            if favicon_path.exists():
                self.iconbitmap(str(favicon_path))
        except Exception:
            pass

        # Set dark title bar on Windows
        try:
            from ctypes import windll, byref, sizeof, c_int
            HWND = windll.user32.GetParent(self.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = c_int(1 if self._theme_manager.is_dark else 0)
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_USE_IMMERSIVE_DARK_MODE, byref(value), sizeof(value))
        except Exception:
            pass

        self._setup_ui()

        # Calculate center position and set geometry with size AND position
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        w = 775
        h = 515
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.deiconify()  # Show in correct position
        self.grab_set()

    def _init_assignments(self):
        """Initialize assignments from field_items' existing categories"""
        for field_item in self.field_items:
            field_id = id(field_item)
            self.assignments[field_id] = {}
            # Parse existing categories - field_item.categories is [(sort_order, label), ...]
            for col_idx, cat_level in enumerate(self.category_levels):
                # Find if this field has an assignment for this column
                assigned_label = None
                if field_item.categories:
                    # Categories are stored as list of (sort_order, label) tuples
                    # We need to match by column index
                    if col_idx < len(field_item.categories):
                        _, label = field_item.categories[col_idx]
                        if label in cat_level.labels:
                            assigned_label = label
                self.assignments[field_id][col_idx] = assigned_label or ""

    def _init_column_widths(self):
        """Initialize default column widths (in pixels)"""
        # Columns: [#/drag handle, Field Name, category columns...]
        self.col_widths = [45, 200]  # Order/drag column + Field name column
        for _ in self.category_levels:
            self.col_widths.append(180)  # Default category column width (wider for longer names like "Property Metrics Category")

    def _center_on_parent(self, parent):
        """Center dialog on parent window"""
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        w = self.winfo_width()
        h = self.winfo_height()
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.geometry(f"+{x}+{y}")

    def _setup_ui(self):
        """Setup the dialog UI"""
        colors = self._theme_manager.colors
        bg_color = colors['background']

        # Configure dialog background
        self.configure(bg=bg_color)

        main_frame = tk.Frame(self, bg=bg_color, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Row 1 - Explainer text
        info_frame = tk.Frame(main_frame, bg=bg_color)
        info_frame.pack(fill=tk.X, pady=(0, 2))

        tk.Label(
            info_frame,
            text="Click to select • Drag to reorder • Ctrl+click to toggle • Shift+click for range",
            font=("Segoe UI", 8, "italic"),
            bg=bg_color, fg=colors['text_muted']
        ).pack(side=tk.LEFT)

        # Row 2 - Filter on left, up/down arrows on right
        filter_frame = tk.Frame(main_frame, bg=bg_color)
        filter_frame.pack(fill=tk.X, pady=(0, 5))

        tk.Label(
            filter_frame, text="Filter:", font=("Segoe UI", 8),
            bg=bg_color, fg=colors['text_primary']
        ).pack(side=tk.LEFT, padx=(0, 3))

        # Right side controls frame (selection info + delete/up/down arrows)
        right_controls = tk.Frame(filter_frame, bg=bg_color)
        right_controls.pack(side=tk.RIGHT)
        self._right_controls = right_controls  # Store for theme updates

        # Arrow frame to group buttons together (packed to RIGHT, but buttons inside pack LEFT)
        arrow_frame = tk.Frame(right_controls, bg=bg_color)
        arrow_frame.pack(side=tk.RIGHT)
        self._arrow_frame = arrow_frame  # Store for theme updates

        # Compute disabled colors based on current theme
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a')

        # Delete button first (left position) - matches main Parameter Builder layout
        self.delete_btn = RoundedButton(
            arrow_frame, text="\u2716", font=("Segoe UI", 8),
            command=self._delete_selected_fields,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            width=26, height=24, radius=5,
            canvas_bg=bg_color
        )
        self.delete_btn.pack(side=tk.LEFT, padx=1)
        self.delete_btn.set_enabled(False)
        self._delete_tooltip = Tooltip(
            self.delete_btn,
            text=lambda: "Delete Fields" if len(self.selected_rows) > 1 else "Delete Field"
        )

        # Up arrow button (middle position)
        self.move_up_btn = RoundedButton(
            arrow_frame, text="\u25B2", font=("Segoe UI", 8),
            command=self._move_selected_up,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            width=26, height=24, radius=5,
            canvas_bg=bg_color
        )
        self.move_up_btn.pack(side=tk.LEFT, padx=1)
        self.move_up_btn.set_enabled(False)

        # Down arrow button (right position)
        self.move_down_btn = RoundedButton(
            arrow_frame, text="\u25BC", font=("Segoe UI", 8),
            command=self._move_selected_down,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            width=26, height=24, radius=5,
            canvas_bg=bg_color
        )
        self.move_down_btn.pack(side=tk.LEFT, padx=1)
        self.move_down_btn.set_enabled(False)

        # Track secondary buttons for theme updates
        self._secondary_buttons = [self.delete_btn, self.move_up_btn, self.move_down_btn]

        # Register for theme changes for live updates
        self._theme_manager.register_theme_callback(self._on_theme_changed)

        # Select All checkbox - SVGCheckbox for consistent styling
        self.select_all_var = tk.BooleanVar(value=False)
        self.select_all_cb = SVGCheckbox(
            right_controls,
            text="Select All",
            variable=self.select_all_var,
            command=self._toggle_select_all,
            bg=bg_color
        )
        self.select_all_cb.pack(side=tk.RIGHT, padx=(0, 8))

        # Selection label (e.g., "0 of 342 selected") - to the left of checkbox
        self.selection_label = tk.Label(
            right_controls,
            text="0 selected",
            font=("Segoe UI", 8, "italic"),
            bg=bg_color, fg=colors['text_muted']
        )
        self.selection_label.pack(side=tk.RIGHT, padx=(0, 10))

        # Modern hierarchical filter dropdown (same as main Parameter Builder)
        # Import here to avoid circular import (field_parameters_builder imports from dialogs)
        from tools.field_parameters.field_parameters_builder import FieldParameterFilterDropdown
        self._filter_dropdown = FieldParameterFilterDropdown(
            parent=filter_frame,
            theme_manager=self._theme_manager,
            on_filter_changed=self._on_dialog_filter_changed
        )
        self._filter_dropdown.pack(side=tk.LEFT, padx=(0, 3))

        # Set category levels for the filter dropdown
        self._filter_dropdown.set_category_levels(self.category_levels)

        # Filter state - None = show all
        self._dialog_category_filters: Optional[Dict] = None

        self.dialog_filter_count = tk.Label(
            filter_frame,
            text=f"Showing all {len(self.field_items)} fields",
            font=("Segoe UI", 8, "italic"),
            bg=bg_color, fg=colors['text_muted']
        )
        self.dialog_filter_count.pack(side=tk.LEFT)

        # Header container - includes header labels + spacer for scrollbar alignment
        # Use section_bg for header background (like treeview headers)
        is_dark = self._theme_manager.is_dark
        if is_dark:
            heading_bg = colors.get('section_bg', '#1a1a2a')
            heading_fg = colors.get('text_primary', '#e0e0e0')
            header_separator = colors.get('border', '#3a3a4a')
        else:
            heading_bg = colors.get('section_bg', '#f5f5fa')
            heading_fg = colors.get('text_primary', '#333333')
            header_separator = colors.get('border', '#d8d8e0')

        # Outer container with border for consistent styling
        header_outer = tk.Frame(main_frame, bg=header_separator)
        header_outer.pack(fill=tk.X, pady=(0, 5))

        header_container = tk.Frame(header_outer, bg=heading_bg)
        header_container.pack(fill=tk.X, padx=1, pady=1)

        # Fixed header row (outside scrollable area) - use grid for pixel-precise alignment
        self.header_frame = tk.Frame(header_container, bg=heading_bg)
        self.header_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Spacer to account for scrollbar width (keeps header aligned with rows)
        # Standard scrollbar is ~17-20 pixels, use 20 for safety
        scrollbar_spacer = tk.Frame(header_container, width=20, bg=heading_bg)
        scrollbar_spacer.pack(side=tk.RIGHT)
        scrollbar_spacer.pack_propagate(False)

        # Store header widgets for resizing
        self.header_labels: List[tk.Label] = []
        self.header_separators: List[tk.Frame] = []
        self._header_outer = header_outer  # Store for theme updates
        self._header_container_inner = header_container  # Store for theme updates

        # Build column headers - use pixel widths matching row layout
        # Columns: [#/drag, Field Name, category columns...]
        column_names = ["#", "Field Name"] + [cat_level.name for cat_level in self.category_levels]

        for col_idx, col_name in enumerate(column_names):
            # Determine padding to match row layout
            # Note: row_frame has highlightthickness=1, adding 1px border on left
            # So header needs +1px left padding to align with row content
            if col_idx == 0:
                # Order/drag column - minimal padding
                padx_val = (6, 2)
            elif col_idx == 1:
                # Field Name column - matches name_label padx=(5, 10) + 1px for row border
                padx_val = (0, 10)
            else:
                # Category columns - matches combo padx=(0, 5)
                padx_val = (0, 5)

            # Use a frame container with fixed pixel width for precise alignment
            col_frame = tk.Frame(self.header_frame, width=self.col_widths[col_idx], height=26, bg=heading_bg)
            col_frame.pack(side=tk.LEFT, padx=padx_val)
            col_frame.pack_propagate(False)  # Prevent children from changing frame size

            # Column header label inside the fixed-width frame - styled like treeview heading
            header_label = tk.Label(
                col_frame,
                text=col_name,
                font=("Segoe UI", 9, "bold"),
                anchor="center",  # Center all column headers
                bg=heading_bg,
                fg=heading_fg,
                pady=4
            )
            header_label.pack(fill=tk.BOTH, expand=True)
            self.header_labels.append(col_frame)  # Store frame for width updates

            # Add resize handle on right edge of column frame (all columns including last)
            # Create resize handle that overlays on the right edge of col_frame
            # Use place() to position it on the right edge without adding width
            separator_color = header_separator
            separator_hover = colors['text_muted']
            separator = tk.Frame(col_frame, width=4, bg=separator_color, cursor="sb_h_double_arrow")
            separator.place(relx=1.0, rely=0.15, relheight=0.7, anchor="ne")
            # Bind drag events for resizing
            separator.bind("<Button-1>", lambda e, idx=col_idx: self._start_resize(e, idx))
            separator.bind("<B1-Motion>", self._do_resize)
            separator.bind("<ButtonRelease-1>", self._end_resize)
            # Hover effect - darken on hover
            separator.bind("<Enter>", lambda e, s=separator, c=separator_hover: s.config(bg=c))
            separator.bind("<Leave>", lambda e, s=separator, c=separator_color: s.config(bg=c))
            self.header_separators.append(separator)

        # Separator below header
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=2)

        # Grid container with scrollbars (scrollable area for rows only)
        grid_container = tk.Frame(main_frame, bg=bg_color)
        grid_container.pack(fill=tk.BOTH, expand=True)

        # Canvas for scrolling with themed vertical scrollbar and auto-hide horizontal
        self.canvas = tk.Canvas(grid_container, highlightthickness=0, bg=bg_color)
        self._v_scrollbar = ThemedScrollbar(
            grid_container,
            command=self._on_scroll,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        h_scrollbar = AutoHideScrollbar(grid_container, orient=tk.HORIZONTAL, command=self.canvas.xview)

        self.canvas.configure(yscrollcommand=self._v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self._v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame inside canvas for grid rows (no header here anymore)
        # Use tk.Frame with explicit bg for dark mode consistency
        self.grid_frame = tk.Frame(self.canvas, bg=bg_color)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")

        self.grid_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        # Build the grid rows (header is already created above)
        self._build_grid()

        # Schedule initial scroll region update after widgets are realized
        self.after(10, self._initial_scroll_setup)

        # Separator before bottom section
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(10, 8))

        # Row 1: Bulk edit section - better layout with wider dropdowns
        bulk_frame = tk.Frame(main_frame, bg=bg_color)
        bulk_frame.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            bulk_frame, text="Bulk Edit:", font=("Segoe UI", 9, "bold"),
            bg=bg_color, fg=colors['text_primary']
        ).pack(side=tk.LEFT)

        # Column selector - wider
        tk.Label(
            bulk_frame, text="Column:", font=("Segoe UI", 8),
            bg=bg_color, fg=colors['text_primary']
        ).pack(side=tk.LEFT, padx=(15, 3))
        self.bulk_column_var = tk.StringVar()
        self.bulk_column_combo = ThemedCombobox(
            bulk_frame,
            textvariable=self.bulk_column_var,
            values=[level.name for level in self.category_levels],
            state="readonly",
            width=25,
            font=("Segoe UI", 9),
            theme_manager=self._theme_manager
        )
        if self.category_levels:
            self.bulk_column_combo.current(0)
        self.bulk_column_combo.pack(side=tk.LEFT, padx=(0, 15))

        # Value selector - wider
        tk.Label(
            bulk_frame, text="Value:", font=("Segoe UI", 8),
            bg=bg_color, fg=colors['text_primary']
        ).pack(side=tk.LEFT, padx=(0, 3))
        self.bulk_label_var = tk.StringVar()
        self.bulk_label_combo = ThemedCombobox(
            bulk_frame,
            textvariable=self.bulk_label_var,
            state="readonly",
            width=30,
            font=("Segoe UI", 9),
            theme_manager=self._theme_manager
        )
        self._update_bulk_labels()
        self.bulk_column_combo.bind("<<ComboboxSelected>>", lambda e: self._update_bulk_labels())
        self.bulk_label_combo.pack(side=tk.LEFT, padx=(0, 10))

        self.apply_bulk_btn = RoundedButton(
            bulk_frame,
            text="Apply",
            command=self._apply_bulk_edit,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=32, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=bg_color
        )
        self.apply_bulk_btn.pack(side=tk.LEFT)

        # Row 2: Centered Save/Cancel buttons
        bottom_frame = tk.Frame(main_frame, bg=bg_color)
        bottom_frame.pack(fill=tk.X, pady=(0, 0))

        # Center frame for buttons
        center_frame = tk.Frame(bottom_frame, bg=bg_color)
        center_frame.pack(anchor=tk.CENTER)

        save_btn = RoundedButton(
            center_frame,
            text="Save",
            command=self._save,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            height=38, radius=6,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=bg_color
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 8))

        cancel_btn = RoundedButton(
            center_frame,
            text="Cancel",
            command=self.destroy,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=38, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=bg_color
        )
        cancel_btn.pack(side=tk.LEFT)

    def _initial_scroll_setup(self):
        """Initial setup for scroll region after widgets are realized"""
        self.grid_frame.update_idletasks()
        self._on_canvas_configure(type('event', (), {'width': self.canvas.winfo_width()})())
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """Handle canvas resize"""
        # Calculate minimum width needed for all columns + padding
        # Field Name: padx=(5,10) = 15px + width[0]
        # Each combo: padx=(0,5) = 5px + width[i]
        # Row border: 2px (highlightthickness=1 on each side)
        min_width = 2 + 5 + self.col_widths[0] + 10  # border + left pad + field name + right pad
        for i in range(1, len(self.col_widths)):
            min_width += self.col_widths[i] + 5  # combo width + right pad

        # Set canvas window width to max of canvas width, content width, or calculated min
        content_width = max(event.width, self.grid_frame.winfo_reqwidth(), min_width)
        self.canvas.itemconfig(self.canvas_window, width=content_width)

    def _on_scroll(self, *args):
        """Handle scrollbar scroll events"""
        self.canvas.yview(*args)
        # Trigger virtual scroll update if in virtual mode
        if self._virtual_mode:
            self._schedule_virtual_update()

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling - only if content is scrollable"""
        bbox = self.canvas.bbox("all")
        if bbox:
            content_height = bbox[3] - bbox[1]
            visible_height = self.canvas.winfo_height()
            if content_height > visible_height:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                # Trigger virtual scroll update if in virtual mode
                if self._virtual_mode:
                    self._schedule_virtual_update()

    def _bind_mousewheel_recursive(self, widget):
        """Bind mousewheel event to widget and all its children recursively"""
        widget.bind("<MouseWheel>", self._on_mousewheel)
        for child in widget.winfo_children():
            self._bind_mousewheel_recursive(child)

    def _start_resize(self, event, col_idx: int):
        """Start column resize drag"""
        self._resize_drag["active"] = True
        self._resize_drag["col_idx"] = col_idx
        self._resize_drag["start_x"] = event.x_root
        self._resize_drag["start_width"] = self.col_widths[col_idx]

    def _do_resize(self, event):
        """Handle column resize drag motion"""
        if not self._resize_drag["active"]:
            return

        col_idx = self._resize_drag["col_idx"]
        delta = event.x_root - self._resize_drag["start_x"]
        new_width = max(60, self._resize_drag["start_width"] + delta)  # Min width 60px

        self.col_widths[col_idx] = new_width
        self._update_column_widths()

    def _end_resize(self, event):
        """End column resize drag"""
        self._resize_drag["active"] = False
        self._resize_drag["col_idx"] = None

    def _update_column_widths(self):
        """Update all column widths after resize"""
        # Update header frames (pixel-based)
        for col_idx, header_frame in enumerate(self.header_labels):
            header_frame.config(width=self.col_widths[col_idx])

        # Update row widgets (pixel-based frames)
        # Columns: [0: order/drag, 1: name, 2+: categories]
        for field_id, widgets in self.row_widgets.items():
            # Update order frame width (column 0)
            if "order_frame" in widgets:
                widgets["order_frame"].config(width=self.col_widths[0])

            # Update name frame width (column 1)
            if "name_frame" in widgets:
                widgets["name_frame"].config(width=self.col_widths[1])

            # Update combo frame widths (columns 2+)
            for combo_info in widgets["combos"]:
                combo_col_idx = combo_info["col_idx"] + 2  # +2 because col 0 is order, col 1 is name
                if combo_col_idx < len(self.col_widths) and "frame" in combo_info:
                    combo_info["frame"].config(width=self.col_widths[combo_col_idx])

        # Update canvas window width and scroll region after resize
        self.grid_frame.update_idletasks()
        self._on_canvas_configure(type('event', (), {'width': self.canvas.winfo_width()})())
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    # =========================================================================
    # FILTER MENU METHODS (Expandable multi-select like main Parameter Builder)
    # =========================================================================

    def _on_dialog_filter_changed(self):
        """Handle filter dropdown selection change - callback from FieldParameterFilterDropdown"""
        self._dialog_category_filters = self._filter_dropdown.get_filter_state()
        self._apply_dialog_filter()

    def _clear_dialog_filters(self):
        """Clear all filters and select (All)"""
        self._filter_dropdown.clear_filters()
        self._dialog_category_filters = None
        self._apply_dialog_filter()

    def _apply_dialog_filter(self):
        """Apply filter to show/hide rows in the dialog"""
        # Reset scroll to top
        self.canvas.yview_moveto(0)

        if self._virtual_mode:
            # Refresh virtual view after filter change
            self._refresh_virtual_view()
        else:
            # Non-virtual mode: show/hide rows directly
            for field_item in self.field_items:
                field_id = id(field_item)
                if field_id not in self.row_widgets:
                    continue

                frame = self.row_widgets[field_id]["frame"]
                should_show = self._dialog_field_matches_filter(field_item)

                if should_show:
                    if not frame.winfo_ismapped():
                        frame.pack(fill=tk.X, pady=1)
                else:
                    frame.pack_forget()

        self._update_dialog_filter_count()

    def _dialog_field_matches_filter(self, field_item: FieldItem) -> bool:
        """Check if a field matches the current filter (multi-column, multi-select)"""
        if self._dialog_category_filters is None:
            return True  # Show all

        # Get all category labels for this field
        field_labels = []
        if field_item.categories:
            for _, label in field_item.categories:
                field_labels.append(label or "")
        else:
            field_labels = [""]  # Uncategorized

        # Check if uncategorized filter is active
        if self._dialog_category_filters.get("__uncategorized__"):
            # Check if field is uncategorized (all labels empty)
            if all(lbl == "" for lbl in field_labels):
                return True

        # Check if field matches any selected label in any column
        for col_idx, selected_labels in self._dialog_category_filters.items():
            if col_idx == "__uncategorized__":
                continue

            if isinstance(selected_labels, set):
                # Get this field's label for this column
                if col_idx < len(field_labels):
                    field_label = field_labels[col_idx]
                    if field_label in selected_labels:
                        return True

        return False

    def _update_dialog_filter_count(self):
        """Update the filter count label in dialog"""
        if self._dialog_category_filters is None:
            self.dialog_filter_count.config(text=f"Showing all {len(self.field_items)} fields")
        else:
            visible_count = sum(1 for fi in self.field_items if self._dialog_field_matches_filter(fi))
            self.dialog_filter_count.config(text=f"Showing {visible_count} of {len(self.field_items)} fields")

    # =========================================================================
    # VIRTUAL SCROLLING METHODS
    # =========================================================================

    def _schedule_virtual_update(self):
        """Schedule a virtual scroll update with debouncing"""
        if self._scroll_job:
            self.after_cancel(self._scroll_job)
        # Debounce: wait 16ms (~60fps) before updating
        self._scroll_job = self.after(16, self._update_virtual_view)

    def _update_filtered_items_cache(self):
        """Update the cached list of items matching the current filter"""
        if self._dialog_category_filters is None:
            self._filtered_items = list(self.field_items)
        else:
            self._filtered_items = [fi for fi in self.field_items if self._dialog_field_matches_filter(fi)]

    def _get_visible_range(self) -> Tuple[int, int]:
        """Calculate which items should be visible based on scroll position"""
        if not self._filtered_items:
            return (0, 0)

        # Get current scroll position and viewport height
        viewport_height = self.canvas.winfo_height()
        if viewport_height <= 1:
            viewport_height = 400  # Default fallback

        # Get scroll position (0.0 to 1.0)
        try:
            scroll_top = self.canvas.yview()[0]
        except:
            scroll_top = 0.0

        total_items = len(self._filtered_items)
        total_height = total_items * self.ROW_HEIGHT

        # Calculate visible item range
        top_y = scroll_top * total_height
        bottom_y = top_y + viewport_height

        start_idx = max(0, int(top_y / self.ROW_HEIGHT) - self.BUFFER_ROWS)
        end_idx = min(total_items, int(bottom_y / self.ROW_HEIGHT) + self.BUFFER_ROWS + 1)

        return (start_idx, end_idx)

    def _update_virtual_view(self):
        """Update which widgets are rendered based on scroll position"""
        if not self._virtual_mode or not self._filtered_items:
            return

        new_range = self._get_visible_range()
        if new_range == self._visible_range:
            return  # No change needed

        old_start, old_end = self._visible_range
        new_start, new_end = new_range

        # Freeze UI updates during batch operation to prevent flicker
        self.grid_frame.update_idletasks()

        # FIRST: Show/create widgets that scrolled into view (place before hiding)
        # This prevents blank areas during scroll
        for i in range(new_start, new_end):
            if i < len(self._filtered_items):
                field_item = self._filtered_items[i]
                field_id = id(field_item)

                # Create widget if it doesn't exist
                if field_id not in self.row_widgets:
                    self._create_field_row_fast(field_item, i)
                else:
                    # Update order label if row was already created (might have different index now)
                    if "order_label" in self.row_widgets[field_id]:
                        self.row_widgets[field_id]["order_label"].config(text=f"{i + 1} \u2261")

                # Position widget using place() for absolute positioning
                frame = self.row_widgets[field_id]["frame"]
                y_pos = i * self.ROW_HEIGHT
                frame.place(x=0, y=y_pos, relwidth=1.0, height=self.ROW_HEIGHT - 2)

        # THEN: Hide widgets that scrolled out of view
        for i in range(old_start, old_end):
            if i < new_start or i >= new_end:
                if i < len(self._filtered_items):
                    field_item = self._filtered_items[i]
                    field_id = id(field_item)
                    if field_id in self.row_widgets:
                        self.row_widgets[field_id]["frame"].place_forget()

        self._visible_range = new_range

        # Apply selection highlighting to visible widgets
        self._update_row_highlights()

    def _setup_virtual_scrolling(self):
        """Setup virtual scrolling mode for large lists"""
        self._virtual_mode = True
        self._update_filtered_items_cache()

        # Set scroll region based on total item count
        total_height = len(self._filtered_items) * self.ROW_HEIGHT
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), total_height))

        # Configure grid_frame to fill the scroll region
        self.grid_frame.configure(height=total_height)

        # Initial render
        self._visible_range = (0, 0)
        self._update_virtual_view()

    def _refresh_virtual_view(self):
        """Refresh the virtual view after data changes (filter, reorder)"""
        if not self._virtual_mode:
            return

        # Update filtered items cache
        self._update_filtered_items_cache()

        # Update scroll region
        total_height = max(1, len(self._filtered_items) * self.ROW_HEIGHT)
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), total_height))
        self.grid_frame.configure(height=total_height)

        # Hide all currently visible widgets first
        old_start, old_end = self._visible_range
        for i in range(old_start, old_end):
            if i < len(self._filtered_items):
                field_item = self._filtered_items[i]
                field_id = id(field_item)
                if field_id in self.row_widgets:
                    self.row_widgets[field_id]["frame"].place_forget()

        # Reset visible range and re-render
        self._visible_range = (0, 0)
        self._update_virtual_view()

    def _create_field_row_fast(self, field_item: FieldItem, row_index: int = 0):
        """Create field row optimized for virtual scrolling - uses tk widgets for speed"""
        field_id = id(field_item)
        colors = self._theme_manager.colors
        # Zebra striping - matches Parameter Builder pattern
        row_bg = colors['card_surface'] if row_index % 2 == 0 else colors['surface']

        # Use tk.Frame (faster than ttk.Frame) - flat design without border
        row_frame = tk.Frame(self.grid_frame, bg=row_bg)
        # Don't pack - will use place() for virtual scrolling

        # Pre-create bound functions (avoid lambda overhead)
        click_handler = lambda e, fid=field_id: self._on_row_click(e, fid)
        motion_handler = lambda e, fid=field_id: self._on_row_drag_motion(e, fid)
        right_click_handler = lambda e, fid=field_id: self._on_row_right_click(e, fid)

        # Bind to frame
        row_frame.bind("<Button-1>", click_handler)
        row_frame.bind("<B1-Motion>", motion_handler)
        row_frame.bind("<ButtonRelease-1>", self._on_row_drag_end)
        row_frame.bind("<Button-3>", right_click_handler)
        row_frame.bind("<MouseWheel>", self._on_mousewheel)

        # Order/drag column - shows row number and drag handle
        order_frame = tk.Frame(row_frame, width=self.col_widths[0], height=22, bg=row_bg)
        order_frame.pack(side=tk.LEFT, padx=(5, 2), pady=3)
        order_frame.pack_propagate(False)

        order_label = tk.Label(
            order_frame,
            text=f"{row_index + 1} ≡",
            font=("Segoe UI", 9),
            anchor="w",
            bg=row_bg,
            fg=colors['text_muted'],
            cursor="hand2"
        )
        order_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        order_label.bind("<Button-1>", click_handler)
        order_label.bind("<B1-Motion>", motion_handler)
        order_label.bind("<ButtonRelease-1>", self._on_row_drag_end)
        order_label.bind("<Button-3>", right_click_handler)
        order_label.bind("<MouseWheel>", self._on_mousewheel)

        # Field name label - use pixel-based container frame for alignment
        name_frame = tk.Frame(row_frame, width=self.col_widths[1], height=22, bg=row_bg)
        name_frame.pack(side=tk.LEFT, padx=(0, 10), pady=3)
        name_frame.pack_propagate(False)

        name_label = tk.Label(
            name_frame,
            text=field_item.display_name,
            font=("Segoe UI", 9),
            anchor="w",
            bg=row_bg,
            fg=colors['text_primary'],
            cursor="hand2"
        )
        name_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        name_label.bind("<Button-1>", click_handler)
        name_label.bind("<B1-Motion>", motion_handler)
        name_label.bind("<ButtonRelease-1>", self._on_row_drag_end)
        name_label.bind("<Button-3>", right_click_handler)
        name_label.bind("<MouseWheel>", self._on_mousewheel)

        # Category dropdowns - use pixel-based container frames
        combos = []
        for col_idx, cat_level in enumerate(self.category_levels):
            combo_var = tk.StringVar(value=self.assignments[field_id].get(col_idx, ""))

            combo_col_idx = col_idx + 2  # +2 because col 0 is order, col 1 is name
            pixel_width = self.col_widths[combo_col_idx] if combo_col_idx < len(self.col_widths) else 150

            combo_frame = tk.Frame(row_frame, width=pixel_width, height=26, bg=row_bg)
            combo_frame.pack(side=tk.LEFT, padx=(0, 5), pady=2)
            combo_frame.pack_propagate(False)

            combo = ThemedCombobox(
                combo_frame,
                textvariable=combo_var,
                values=[""] + cat_level.labels,
                state="readonly",
                font=("Segoe UI", 8),
                theme_manager=self._theme_manager
            )
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
            combo.bind("<<ComboboxSelected>>",
                      lambda e, fid=field_id, cidx=col_idx, var=combo_var: self._on_combo_changed(fid, cidx, var))
            combo.bind("<MouseWheel>", self._on_mousewheel)

            combos.append({"combo": combo, "var": combo_var, "col_idx": col_idx, "frame": combo_frame})

        self.row_widgets[field_id] = {
            "frame": row_frame,
            "order_frame": order_frame,
            "order_label": order_label,
            "name_frame": name_frame,
            "name_label": name_label,
            "combos": combos,
            "field_item": field_item
        }

    def _build_grid(self):
        """Build the grid of field rows (header is created separately in _setup_ui)"""
        # Store row widgets for selection highlighting
        self.row_widgets: Dict[int, Dict] = {}  # field_id -> {frame, combos, ...}

        # Check if we should use virtual scrolling
        if len(self.field_items) > self.VIRTUAL_SCROLL_THRESHOLD:
            self._setup_virtual_scrolling()
        else:
            # Field rows only (no header - it's now fixed outside scrollable area)
            for idx, field_item in enumerate(self.field_items):
                self._create_field_row(field_item, idx)

    def _create_field_row(self, field_item: FieldItem, row_index: int = 0):
        """Create a row for a field with category dropdowns"""
        field_id = id(field_item)
        colors = self._theme_manager.colors
        # Zebra striping - matches Parameter Builder pattern
        row_bg = colors['card_surface'] if row_index % 2 == 0 else colors['surface']

        # Flat design without border - matches Parameter Builder
        row_frame = tk.Frame(self.grid_frame, bg=row_bg)
        row_frame.pack(fill=tk.X, pady=1)

        # Bind click and drag for selection/reordering
        row_frame.bind("<Button-1>", lambda e, fid=field_id: self._on_row_click(e, fid))
        row_frame.bind("<B1-Motion>", lambda e, fid=field_id: self._on_row_drag_motion(e, fid))
        row_frame.bind("<ButtonRelease-1>", self._on_row_drag_end)
        row_frame.bind("<Button-3>", lambda e, fid=field_id: self._on_row_right_click(e, fid))

        # Order/drag column - shows row number and drag handle
        order_frame = tk.Frame(row_frame, width=self.col_widths[0], height=22, bg=row_bg)
        order_frame.pack(side=tk.LEFT, padx=(5, 2), pady=3)
        order_frame.pack_propagate(False)

        order_label = tk.Label(
            order_frame,
            text=f"{row_index + 1} ≡",
            font=("Segoe UI", 9),
            anchor="w",
            bg=row_bg,
            fg=colors['text_muted'],
            cursor="hand2"
        )
        order_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        order_label.bind("<Button-1>", lambda e, fid=field_id: self._on_row_click(e, fid))
        order_label.bind("<B1-Motion>", lambda e, fid=field_id: self._on_row_drag_motion(e, fid))
        order_label.bind("<ButtonRelease-1>", self._on_row_drag_end)
        order_label.bind("<Button-3>", lambda e, fid=field_id: self._on_row_right_click(e, fid))

        # Field name label - use pixel-based container frame for alignment
        name_frame = tk.Frame(row_frame, width=self.col_widths[1], height=22, bg=row_bg)
        name_frame.pack(side=tk.LEFT, padx=(0, 10), pady=3)
        name_frame.pack_propagate(False)  # Fixed pixel width

        name_label = tk.Label(
            name_frame,
            text=field_item.display_name,
            font=("Segoe UI", 9),
            anchor="w",
            bg=row_bg,
            fg=colors['text_primary'],
            cursor="hand2"  # Indicate draggable
        )
        name_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        name_label.bind("<Button-1>", lambda e, fid=field_id: self._on_row_click(e, fid))
        name_label.bind("<B1-Motion>", lambda e, fid=field_id: self._on_row_drag_motion(e, fid))
        name_label.bind("<ButtonRelease-1>", self._on_row_drag_end)
        name_label.bind("<Button-3>", lambda e, fid=field_id: self._on_row_right_click(e, fid))

        # Category dropdowns - use pixel-based container frames
        combos = []
        combo_frames = []  # Store frames for width updates
        for col_idx, cat_level in enumerate(self.category_levels):
            combo_var = tk.StringVar(value=self.assignments[field_id].get(col_idx, ""))

            # Use pixel width from col_widths (col_idx + 2 because col 0 is order, col 1 is name)
            combo_col_idx = col_idx + 2
            pixel_width = self.col_widths[combo_col_idx] if combo_col_idx < len(self.col_widths) else 150

            # Container frame with fixed pixel width and height
            combo_frame = tk.Frame(row_frame, width=pixel_width, height=26, bg=row_bg)
            combo_frame.pack(side=tk.LEFT, padx=(0, 5), pady=2)
            combo_frame.pack_propagate(False)
            combo_frames.append(combo_frame)

            combo = ThemedCombobox(
                combo_frame,
                textvariable=combo_var,
                values=[""] + cat_level.labels,
                state="readonly",
                font=("Segoe UI", 8),
                theme_manager=self._theme_manager
            )
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Bind change event
            combo.bind("<<ComboboxSelected>>",
                      lambda e, fid=field_id, cidx=col_idx, var=combo_var: self._on_combo_changed(fid, cidx, var))

            combos.append({"combo": combo, "var": combo_var, "col_idx": col_idx, "frame": combo_frame})

        self.row_widgets[field_id] = {
            "frame": row_frame,
            "order_frame": order_frame,
            "order_label": order_label,
            "name_frame": name_frame,
            "name_label": name_label,
            "combos": combos,
            "field_item": field_item
        }

        # Bind mousewheel to row and all children so scrolling works anywhere
        self._bind_mousewheel_recursive(row_frame)

    def _on_row_click(self, event, field_id: int):
        """Handle row click for selection and prepare for potential drag"""
        ctrl_held = event.state & 0x4  # Control key
        shift_held = event.state & 0x1  # Shift key

        # Track if we should defer selection change until button release
        # This prevents deselecting items when starting a drag
        self._row_drag["pending_deselect"] = None

        if shift_held and self.last_clicked_row is not None:
            # Range selection from last clicked to current
            self._select_range(self.last_clicked_row, field_id)
        elif ctrl_held:
            # Toggle selection
            if field_id in self.selected_rows:
                self.selected_rows.discard(field_id)
            else:
                self.selected_rows.add(field_id)
            self.last_clicked_row = field_id
        else:
            # Single click behavior - but defer deselection if item is already selected
            # (user might be starting a drag of multiple selected items)
            if field_id in self.selected_rows:
                # Don't deselect yet - mark as pending and wait for button release
                self._row_drag["pending_deselect"] = field_id
            else:
                # Clicking unselected item - select only this item (clear others)
                self.selected_rows.clear()
                self.selected_rows.add(field_id)
            self.last_clicked_row = field_id

        self._update_row_highlights()
        self._update_selection_label()

        # Store for potential drag operation
        self._row_drag["field_id"] = field_id
        self._row_drag["start_y"] = event.y_root
        self._row_drag["dragging"] = False

    def _select_range(self, from_field_id: int, to_field_id: int):
        """Select all rows in a range between two field IDs"""
        # Get indices of both fields in the field_items list
        from_idx = None
        to_idx = None

        for idx, field_item in enumerate(self.field_items):
            fid = id(field_item)
            if fid == from_field_id:
                from_idx = idx
            if fid == to_field_id:
                to_idx = idx

        if from_idx is None or to_idx is None:
            return

        # Determine range bounds
        start_idx = min(from_idx, to_idx)
        end_idx = max(from_idx, to_idx)

        # Select all fields in range
        for idx in range(start_idx, end_idx + 1):
            field_id = id(self.field_items[idx])
            self.selected_rows.add(field_id)

    def _update_row_highlights(self):
        """Update visual highlighting of selected rows - preserves zebra striping"""
        colors = self._theme_manager.colors
        # Selection colors - match main Parameter Builder
        selection_bg = colors['selection_highlight']
        selection_border = colors['button_primary']  # Match drag/drop indicator color
        # Zebra striping colors - matches Parameter Builder pattern
        even_bg = colors['card_surface']
        odd_bg = colors['surface']
        normal_border = colors['border']

        # Iterate through field_items to get proper row index for zebra striping
        for row_idx, field_item in enumerate(self.field_items):
            field_id = id(field_item)
            if field_id not in self.row_widgets:
                continue

            widgets = self.row_widgets[field_id]
            if field_id in self.selected_rows:
                # Selected - use highlight color with visible border
                widgets["frame"].config(
                    bg=selection_bg,
                    highlightbackground=selection_border,
                    highlightcolor=selection_border,
                    highlightthickness=1
                )
                if "order_frame" in widgets:
                    widgets["order_frame"].config(bg=selection_bg)
                if "order_label" in widgets:
                    widgets["order_label"].config(bg=selection_bg)
                if "name_frame" in widgets:
                    widgets["name_frame"].config(bg=selection_bg)
                widgets["name_label"].config(bg=selection_bg)
                # Also update combo frames
                for combo_info in widgets.get("combos", []):
                    if "frame" in combo_info:
                        combo_info["frame"].config(bg=selection_bg)
            else:
                # Use zebra striping for non-selected rows
                row_bg = even_bg if row_idx % 2 == 0 else odd_bg
                widgets["frame"].config(
                    bg=row_bg,
                    highlightbackground=row_bg,
                    highlightcolor=row_bg,
                    highlightthickness=1
                )
                if "order_frame" in widgets:
                    widgets["order_frame"].config(bg=row_bg)
                if "order_label" in widgets:
                    widgets["order_label"].config(bg=row_bg)
                if "name_frame" in widgets:
                    widgets["name_frame"].config(bg=row_bg)
                widgets["name_label"].config(bg=row_bg)
                # Also update combo frames
                for combo_info in widgets.get("combos", []):
                    if "frame" in combo_info:
                        combo_info["frame"].config(bg=row_bg)

    def _update_selection_label(self):
        """Update the selection count label"""
        count = len(self.selected_rows)
        total = len(self.field_items)
        self.selection_label.config(text=f"{count} of {total} selected")

        # Update select all checkbox state
        if hasattr(self, 'select_all_var'):
            if count == total and total > 0:
                self.select_all_var.set(True)
            else:
                self.select_all_var.set(False)

        # Update up/down/delete button states (RoundedButton uses set_enabled)
        if hasattr(self, 'move_up_btn') and hasattr(self, 'move_down_btn'):
            if count > 0:
                self.move_up_btn.set_enabled(True)
                self.move_down_btn.set_enabled(True)
            else:
                self.move_up_btn.set_enabled(False)
                self.move_down_btn.set_enabled(False)
        if hasattr(self, 'delete_btn'):
            self.delete_btn.set_enabled(count > 0)

    def _toggle_select_all(self):
        """Toggle select all fields"""
        if self.select_all_var.get():
            # Select all
            for field_item in self.field_items:
                self.selected_rows.add(id(field_item))
        else:
            # Deselect all
            self.selected_rows.clear()

        self._update_row_highlights()
        self._update_selection_label()

    def _on_row_right_click(self, event, field_id: int):
        """Handle right-click for context menu on rows"""
        # If right-clicked item not in selection, select only it
        if field_id not in self.selected_rows:
            self.selected_rows.clear()
            self.selected_rows.add(field_id)
            self._update_row_highlights()
            self._update_selection_label()

        # Show themed context menu (modern dropdown design)
        menu = ThemedContextMenu(self, self._theme_manager)
        count = len(self.selected_rows)
        label_suffix = f" ({count} items)" if count > 1 else ""

        menu.add_command(label=f"Move to Top{label_suffix}", command=self._dialog_move_to_top)
        menu.add_command(label=f"Move to Position...{label_suffix}", command=self._dialog_move_to_position)
        menu.add_command(label=f"Move to Bottom{label_suffix}", command=self._dialog_move_to_bottom)
        menu.add_separator()
        menu.add_command(label=f"Delete{label_suffix}", command=self._delete_selected_fields)

        menu.show(event.x_root, event.y_root)

    def _dialog_move_to_top(self):
        """Move selected rows to top in dialog"""
        if not self.selected_rows:
            return

        # Get selected field items in their current order
        selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_rows]
        remaining_fields = [fi for fi in self.field_items if id(fi) not in self.selected_rows]

        # New order: selected first, then remaining
        self.field_items[:] = selected_fields + remaining_fields

        self._update_order_numbers()
        self._repack_rows()
        self._update_row_highlights()

    def _dialog_move_to_bottom(self):
        """Move selected rows to bottom in dialog"""
        if not self.selected_rows:
            return

        # Get selected field items in their current order
        selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_rows]
        remaining_fields = [fi for fi in self.field_items if id(fi) not in self.selected_rows]

        # New order: remaining first, then selected
        self.field_items[:] = remaining_fields + selected_fields

        self._update_order_numbers()
        self._repack_rows()
        self._update_row_highlights()

    def _dialog_move_to_position(self):
        """Move selected rows to a specific position"""
        if not self.selected_rows:
            return

        total_fields = len(self.field_items)
        current_position = None

        # Get current position of first selected item for default
        for idx, fi in enumerate(self.field_items):
            if id(fi) in self.selected_rows:
                current_position = idx + 1  # 1-based for user
                break

        # Prompt user for position
        position_str = ThemedInputDialog.askstring(
            self,
            "Move to Position",
            f"Enter position (1-{total_fields}):",
            initialvalue=str(current_position) if current_position else "1"
        )

        if not position_str:
            return

        try:
            position = int(position_str)
            if position < 1:
                position = 1
            elif position > total_fields:
                position = total_fields
        except ValueError:
            ThemedMessageBox.showerror(self, "Invalid Position", "Please enter a valid number.")
            return

        # Convert to 0-based index for insertion
        # User says "position 7" = item should END UP at index 6 (0-based)
        target_idx = position - 1

        # Get selected field items in their current order
        selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_rows]
        remaining_fields = [fi for fi in self.field_items if id(fi) not in self.selected_rows]

        # Simple approach: insert at target_idx in remaining list
        # Clamp to valid range
        insert_pos = min(target_idx, len(remaining_fields))

        # Insert selected items at the target position
        new_order = remaining_fields[:insert_pos] + selected_fields + remaining_fields[insert_pos:]
        self.field_items[:] = new_order

        self._update_order_numbers()
        self._repack_rows()
        self._update_row_highlights()

    def _update_order_numbers(self):
        """Update order number labels in the dialog"""
        for idx, field_item in enumerate(self.field_items, 1):
            field_id = id(field_item)
            if field_id in self.row_widgets:
                if "order_label" in self.row_widgets[field_id]:
                    label = self.row_widgets[field_id]["order_label"]
                    new_text = f"{idx} \u2261"  # Number + drag handle symbol
                    if label.cget("text") != new_text:
                        label.config(text=new_text)

    def _repack_rows(self):
        """Repack all rows in the correct order after reordering"""
        if self._virtual_mode:
            # Virtual mode - update filtered cache and refresh view
            self._update_filtered_items_cache()
            # Reset visible range to force re-render
            old_range = self._visible_range
            self._visible_range = (0, 0)
            # Update scroll region for new order
            total_height = len(self._filtered_items) * self.ROW_HEIGHT
            self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), total_height))
            self.grid_frame.configure(height=total_height)
            # Re-render visible items
            self._update_virtual_view()
        else:
            # Non-virtual mode - unpack and repack all rows
            # First, unpack all row widgets
            for field_id, widgets in self.row_widgets.items():
                widgets["frame"].pack_forget()

            # Then repack in new order (respecting filter)
            for field_item in self.field_items:
                field_id = id(field_item)
                if field_id in self.row_widgets:
                    if self._dialog_field_matches_filter(field_item):
                        self.row_widgets[field_id]["frame"].pack(fill=tk.X, pady=1, padx=1)

    # =========================================================================
    # ROW DRAG-DROP REORDERING METHODS
    # =========================================================================

    def _dialog_auto_scroll_if_near_edge(self, event):
        """Auto-scroll the dialog canvas when dragging near the edge of the visible area.

        Uses variable speed - faster scrolling the closer to the edge.
        """
        canvas = self.canvas
        canvas_top = canvas.winfo_rooty()
        canvas_height = canvas.winfo_height()
        canvas_bottom = canvas_top + canvas_height

        # Define edge zones (pixels from edge)
        edge_zone = 60
        mouse_y = event.y_root

        # Calculate distance into the edge zone
        if mouse_y < canvas_top + edge_zone:
            # Near top edge - scroll up (only if not already at top)
            scroll_top = canvas.yview()[0]
            if scroll_top > 0:
                distance_into_zone = canvas_top + edge_zone - mouse_y
                scroll_amount = max(1, min(3, int(distance_into_zone / 20)))
                canvas.yview_scroll(-scroll_amount, "units")
                # Update virtual view after scroll to prevent row displacement
                if self._virtual_mode:
                    self._schedule_virtual_update()
        elif mouse_y > canvas_bottom - edge_zone:
            # Near bottom edge - scroll down
            distance_into_zone = mouse_y - (canvas_bottom - edge_zone)
            scroll_amount = max(1, min(3, int(distance_into_zone / 20)))
            canvas.yview_scroll(scroll_amount, "units")
            # Update virtual view after scroll to prevent row displacement
            if self._virtual_mode:
                self._schedule_virtual_update()

    def _on_row_drag_motion(self, event, field_id: int):
        """Handle drag motion for row reordering"""
        # Only process if we have an item from click
        if self._row_drag.get("field_id") is None:
            return

        # Check if we've moved enough to start dragging
        start_y = self._row_drag.get("start_y", 0)
        if abs(event.y_root - start_y) < 5 and not self._row_drag.get("dragging"):
            return

        # Mark as actually dragging (not just a click)
        self._row_drag["dragging"] = True
        # Cancel any pending deselect since we're now dragging
        self._row_drag["pending_deselect"] = None

        # Auto-scroll if near edge of visible area
        self._dialog_auto_scroll_if_near_edge(event)

        # Create drag label if not exists
        if self._drag_label is None:
            count = len(self.selected_rows) if self.selected_rows else 1
            if count == 1:
                # Get the field name
                for fi in self.field_items:
                    if id(fi) in self.selected_rows:
                        label_text = fi.display_name
                        break
                else:
                    label_text = "1 field"
            else:
                label_text = f"{count} field(s)"
            self._create_drag_label(label_text)

        # Update drag label position
        if self._drag_label:
            self._drag_label.geometry(f"+{event.x_root + 15}+{event.y_root + 10}")

        # Calculate drop position from mouse Y
        canvas_y = self.canvas.canvasy(event.y_root - self.canvas.winfo_rooty())
        new_position = self._calculate_drop_position(canvas_y)

        # Show drop indicator
        if new_position != self._drop_position:
            self._drop_position = new_position
            self._show_drop_indicator(new_position)

    def _on_row_drag_end(self, event):
        """Handle end of drag operation"""
        drop_pos = self._drop_position

        # Hide drop indicator and drag label
        self._hide_drop_indicator()
        self._destroy_drag_label()

        # Only reorder if we actually dragged (not just clicked)
        if self._row_drag.get("dragging") and drop_pos is not None:
            self._perform_reorder(drop_pos)
        else:
            # Not a drag - handle pending deselect (click on already-selected item)
            pending_deselect = self._row_drag.get("pending_deselect")
            if pending_deselect is not None:
                self.selected_rows.clear()
                self.selected_rows.add(pending_deselect)
                self._update_row_highlights()
                self._update_selection_label()

        # Reset drag state
        self._row_drag = {
            "active": False,
            "start_y": 0,
            "field_id": None,
            "dragging": False,
            "pending_deselect": None
        }
        self._drop_position = None

    def _create_drag_label(self, text: str):
        """Create floating label during drag"""
        colors = self._theme_manager.colors
        self._drag_label = tk.Toplevel(self)
        self._drag_label.overrideredirect(True)
        self._drag_label.attributes('-topmost', True)
        self._drag_label.attributes('-alpha', 0.85)

        label = tk.Label(
            self._drag_label,
            text=f"  {text}  ",
            bg=colors['button_primary'],
            fg='#ffffff',
            font=("Segoe UI", 9, "bold"),
            padx=8,
            pady=4
        )
        label.pack()

    def _destroy_drag_label(self):
        """Destroy the floating drag label"""
        if self._drag_label:
            self._drag_label.destroy()
            self._drag_label = None

    def _calculate_drop_position(self, y: float) -> int:
        """Calculate which position the drop would occur at based on Y coordinate"""
        if not self.field_items:
            return 0

        if self._virtual_mode:
            # Virtual scrolling mode - calculate position based on fixed row height
            # Use filtered items list for accurate positioning
            items_to_check = self._filtered_items if self._filtered_items else self.field_items
            row_idx = int(y / self.ROW_HEIGHT)

            if row_idx < 0:
                # Clamp to first visible item's position in the full list
                if items_to_check:
                    return self.field_items.index(items_to_check[0])
                return 0
            if row_idx >= len(items_to_check):
                return len(self.field_items)

            # Check if we should insert before or after this row
            row_center_y = row_idx * self.ROW_HEIGHT + self.ROW_HEIGHT / 2
            if y < row_center_y:
                # Find the actual index in field_items
                if row_idx < len(items_to_check):
                    return self.field_items.index(items_to_check[row_idx])
                return len(self.field_items)
            else:
                if row_idx + 1 < len(items_to_check):
                    return self.field_items.index(items_to_check[row_idx + 1])
                return len(self.field_items)
        else:
            # Non-virtual mode - use widget heights
            cumulative_y = 0
            for idx, field_item in enumerate(self.field_items):
                field_id = id(field_item)
                if field_id in self.row_widgets:
                    frame = self.row_widgets[field_id]["frame"]
                    row_height = frame.winfo_height() + 2  # +2 for pady

                    # If mouse is in the upper half of this row, drop before it
                    if y < cumulative_y + row_height / 2:
                        return idx
                    cumulative_y += row_height

            # If past all rows, drop at end
            return len(self.field_items)

    def _show_drop_indicator(self, position: int):
        """Show a line indicating where the drop will occur"""
        self._hide_drop_indicator()

        items_to_check = self._filtered_items if self._filtered_items else self.field_items
        is_last_position = position >= len(self.field_items)

        if self._virtual_mode:
            # Virtual scrolling mode - calculate position based on fixed row height
            # Find the visual index for this position
            if items_to_check:
                first_filtered_pos = self.field_items.index(items_to_check[0])
                if position <= first_filtered_pos:
                    # Position is at or before first filtered item - indicator at top
                    visual_idx = 0
                else:
                    # Find which filtered item this position corresponds to
                    visual_idx = len(items_to_check)  # Default to end
                    for i, fi in enumerate(items_to_check):
                        if self.field_items.index(fi) >= position:
                            visual_idx = i
                            break
            else:
                visual_idx = 0

            y_pos = visual_idx * self.ROW_HEIGHT

            # For last position, place indicator just above the last item's bottom edge
            # This ensures it's visible within the grid_frame
            if is_last_position and len(items_to_check) > 0:
                y_pos = len(items_to_check) * self.ROW_HEIGHT
        else:
            # Non-virtual mode - calculate Y position based on widget heights
            y_pos = 0
            for idx, field_item in enumerate(self.field_items):
                if idx == position:
                    break
                field_id = id(field_item)
                if field_id in self.row_widgets:
                    frame = self.row_widgets[field_id]["frame"]
                    y_pos += frame.winfo_height() + 2

        # Create indicator line - use raise to ensure visibility
        colors = self._theme_manager.colors
        self._drop_indicator = tk.Frame(
            self.grid_frame,
            height=3,
            bg=colors['button_primary']
        )
        self._drop_indicator.place(x=0, y=y_pos, relwidth=1.0)
        self._drop_indicator.lift()  # Ensure indicator is on top

    def _hide_drop_indicator(self):
        """Hide the drop indicator"""
        if self._drop_indicator:
            self._drop_indicator.destroy()
            self._drop_indicator = None

    def _perform_reorder(self, target_position: int):
        """Reorder selected items to the target position"""
        if not self.selected_rows:
            return

        # Get selected field items in their current order
        selected_items = []
        non_selected_items = []

        for idx, field_item in enumerate(self.field_items):
            if id(field_item) in self.selected_rows:
                selected_items.append(field_item)
            else:
                non_selected_items.append((idx, field_item))

        if not selected_items:
            return

        # Calculate insertion point in the non-selected list
        # Adjust target position to account for removed selected items before it
        selected_indices = [self.field_items.index(fi) for fi in selected_items]
        items_before_target = sum(1 for idx in selected_indices if idx < target_position)
        adjusted_target = target_position - items_before_target

        # Build new list
        new_items = [fi for _, fi in non_selected_items]
        adjusted_target = min(adjusted_target, len(new_items))

        # Insert selected items at target position
        for i, item in enumerate(selected_items):
            new_items.insert(adjusted_target + i, item)

        # Update field_items
        self.field_items.clear()
        self.field_items.extend(new_items)

        # Rebuild the grid to reflect new order
        self._rebuild_grid()

    def _rebuild_grid(self):
        """Rebuild the grid after reordering - OPTIMIZED to repack instead of recreate"""
        if self._virtual_mode:
            # Virtual scrolling mode - use place() positioning
            self._refresh_virtual_view()
        else:
            # Non-virtual mode - use pack() layout
            # Unpack all row frames
            for widgets in self.row_widgets.values():
                widgets["frame"].pack_forget()

            # Repack in new order - respecting filter, and update order labels
            visible_idx = 0
            for field_item in self.field_items:
                field_id = id(field_item)
                if field_id in self.row_widgets:
                    if self._dialog_field_matches_filter(field_item):
                        self.row_widgets[field_id]["frame"].pack(fill=tk.X, pady=1)
                        # Update order label if present
                        if "order_label" in self.row_widgets[field_id]:
                            self.row_widgets[field_id]["order_label"].config(text=f"{visible_idx + 1} ≡")
                        visible_idx += 1

        # Selection is already preserved (row_widgets weren't destroyed)
        self._update_row_highlights()
        self._update_selection_label()

    def _on_combo_changed(self, field_id: int, col_idx: int, var: tk.StringVar):
        """Handle dropdown change"""
        self.assignments[field_id][col_idx] = var.get()

    def _update_bulk_labels(self):
        """Update bulk label dropdown based on selected column"""
        col_name = self.bulk_column_var.get()
        if not col_name and self.category_levels:
            # Default to first column if none selected
            col_name = self.category_levels[0].name
            self.bulk_column_var.set(col_name)

        for idx, level in enumerate(self.category_levels):
            if level.name == col_name:
                labels = level.labels if level.labels else []
                self.bulk_label_combo['values'] = ["(clear)"] + labels
                if labels:
                    self.bulk_label_combo.set(labels[0])
                else:
                    self.bulk_label_combo.set("")
                break

    def _apply_bulk_edit(self):
        """Apply bulk edit to selected rows"""
        if not self.selected_rows:
            ThemedMessageBox.showinfo(self, "No Selection", "Select one or more rows first (Ctrl+click to multi-select).")
            return

        col_name = self.bulk_column_var.get()
        label = self.bulk_label_var.get()

        # Handle "(clear)" option
        if label == "(clear)":
            label = ""

        # Find column index
        col_idx = None
        for idx, level in enumerate(self.category_levels):
            if level.name == col_name:
                col_idx = idx
                break

        if col_idx is None:
            return

        # Apply to all selected rows
        for field_id in self.selected_rows:
            self.assignments[field_id][col_idx] = label
            # Update the combo widget
            if field_id in self.row_widgets:
                for combo_info in self.row_widgets[field_id]["combos"]:
                    if combo_info["col_idx"] == col_idx:
                        combo_info["var"].set(label)
                        break

        # Update selection label to show success
        count = len(self.selected_rows)
        self.selection_label.config(text=f"Applied to {count} row(s)")

    # =========================================================================
    # DELETE FIELDS METHOD
    # =========================================================================

    def _delete_selected_fields(self):
        """Delete selected fields from the editor"""
        if not self.selected_rows:
            return

        count = len(self.selected_rows)
        msg = f"Are you sure you want to delete {count} field(s)?"
        if not ThemedMessageBox.askyesno(self, "Confirm Delete", msg):
            return

        # Get fields to delete
        fields_to_delete = [fi for fi in self.field_items if id(fi) in self.selected_rows]

        # Remove from field_items
        for field in fields_to_delete:
            field_id = id(field)
            self.field_items.remove(field)

            # Clean up widgets
            if field_id in self.row_widgets:
                self.row_widgets[field_id]["frame"].destroy()
                del self.row_widgets[field_id]

            # Clean up assignments
            if field_id in self.assignments:
                del self.assignments[field_id]

        # Clear selection and update UI
        self.selected_rows.clear()
        self._update_order_numbers()
        self._repack_rows()
        self._update_row_highlights()
        self._update_selection_label()
        self._update_dialog_filter_count()

    # =========================================================================
    # MOVE UP/DOWN METHODS
    # =========================================================================

    def _move_selected_up(self):
        """Move selected items up one position"""
        if not self.selected_rows:
            return

        # Get indices of selected items (sorted)
        selected_indices = sorted([
            self.field_items.index(self.row_widgets[fid]["field_item"])
            for fid in self.selected_rows
            if fid in self.row_widgets
        ])

        if not selected_indices or selected_indices[0] == 0:
            return  # Already at top

        # Move each selected item up
        for idx in selected_indices:
            if idx > 0:
                # Swap with item above
                self.field_items[idx], self.field_items[idx - 1] = \
                    self.field_items[idx - 1], self.field_items[idx]

        self._rebuild_grid()

    def _move_selected_down(self):
        """Move selected items down one position"""
        if not self.selected_rows:
            return

        # Get indices of selected items (sorted in reverse for moving down)
        selected_indices = sorted([
            self.field_items.index(self.row_widgets[fid]["field_item"])
            for fid in self.selected_rows
            if fid in self.row_widgets
        ], reverse=True)

        if not selected_indices or selected_indices[0] == len(self.field_items) - 1:
            return  # Already at bottom

        # Move each selected item down (process in reverse order)
        for idx in selected_indices:
            if idx < len(self.field_items) - 1:
                # Swap with item below
                self.field_items[idx], self.field_items[idx + 1] = \
                    self.field_items[idx + 1], self.field_items[idx]

        self._rebuild_grid()

    def _save(self):
        """Save assignments back to field items and close"""
        for field_item in self.field_items:
            field_id = id(field_item)
            if field_id in self.assignments:
                # Rebuild categories list as [(sort_order, label), ...]
                new_categories = []
                for col_idx, cat_level in enumerate(self.category_levels):
                    label = self.assignments[field_id].get(col_idx, "")
                    if label:
                        # Sort order is the position of the label in the labels list + 1
                        sort_order = cat_level.labels.index(label) + 1 if label in cat_level.labels else 1
                        new_categories.append((sort_order, label))
                    else:
                        # Empty/cleared assignment - empty label treated as uncategorized
                        new_categories.append((999.0, ""))
                field_item.categories = new_categories

        self.on_save()
        self.destroy()

    # =========================================================================
    # THEME HANDLING
    # =========================================================================

    def _on_theme_changed(self, theme_name: str = None):
        """Update UI elements when theme changes"""
        colors = self._theme_manager.colors
        bg_color = colors['background']
        is_dark = self._theme_manager.is_dark

        # FIRST: Update frame backgrounds BEFORE button colors
        # This ensures the parent containers have correct bg before buttons read it
        if hasattr(self, '_arrow_frame'):
            self._arrow_frame.config(bg=bg_color)
        if hasattr(self, '_right_controls'):
            self._right_controls.config(bg=bg_color)

        # Force Tk to process frame background changes
        self.update_idletasks()

        # Theme-aware disabled colors for buttons (is_dark already set above)
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Update secondary buttons (delete, up, down)
        for btn in self._secondary_buttons:
            btn.update_colors(
                bg=colors['button_secondary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                fg=colors['text_primary'],
                disabled_bg=disabled_bg,
                disabled_fg=disabled_fg,
                canvas_bg=bg_color
            )

        # CRITICAL: Use deferred callback to force button redraws in Toplevel window
        # tk.Toplevel dialogs have timing issues - canvas widgets need explicit redraw
        # Capture ALL colors for use in the deferred callback
        _colors = {
            'bg': colors['button_secondary'],
            'hover_bg': colors['button_secondary_hover'],
            'pressed_bg': colors['button_secondary_pressed'],
            'fg': colors['text_primary'],
            'disabled_bg': disabled_bg,
            'disabled_fg': disabled_fg,
            'canvas_bg': bg_color
        }

        def force_button_redraw():
            for btn in self._secondary_buttons:
                if btn.winfo_exists():
                    # Fully re-apply ALL colors
                    btn.bg_normal = _colors['bg']
                    btn.bg_hover = _colors['hover_bg']
                    btn.bg_pressed = _colors['pressed_bg']
                    btn.fg = _colors['fg']
                    btn.bg_disabled = _colors['disabled_bg']
                    btn.fg_disabled = _colors['disabled_fg']
                    btn._explicit_canvas_bg = _colors['canvas_bg']
                    # Update _current_bg based on enabled state
                    btn._current_bg = btn.bg_normal if btn._enabled else btn.bg_disabled
                    # Force canvas background update and redraw
                    btn.config(bg=_colors['canvas_bg'])
                    btn._draw_button()
            # Final update on the dialog itself
            self.update_idletasks()

        # Deferred execution - use longer delay to ensure all other updates complete
        self.after(50, force_button_redraw)

        # Refresh row highlights with new theme colors
        self._update_row_highlights()

    def destroy(self):
        """Clean up before closing dialog"""
        # Unregister from theme changes
        try:
            self._theme_manager.unregister_theme_callback(self._on_theme_changed)
        except Exception:
            pass  # Callback may not be registered or already removed

        super().destroy()

