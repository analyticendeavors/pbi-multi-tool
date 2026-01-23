"""
HierarchicalFilterDropdown - Reusable filter dropdown component.

A hierarchical filter dropdown with checkbox SVGs for filtering results.
Shows parent items (e.g., severity levels, categories) with child items
(e.g., rules, fields) nested underneath.

Uses two icon buttons: Filter (opens dropdown) and Clear (resets all filters).

Built by Reid Havens of Analytic Endeavors
Copyright (c) 2024 Analytic Endeavors LLC. All rights reserved.
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Set, Callable, Optional, Any
from pathlib import Path
import io

# Optional imports for SVG support
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False

# Import core widgets (avoid circular import by importing at module level)
from core.ui_base import SquareIconButton, ThemedScrollbar


class HierarchicalFilterDropdown:
    """
    A hierarchical filter dropdown with checkbox SVGs for filtering results.
    Shows parent items (groups) with child items nested underneath.
    Uses two icon buttons: Filter (opens dropdown) and Clear (resets all filters).

    This is a generic, reusable component that can be customized for different
    use cases by providing different group names, colors, and item data.
    """

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager,
        on_filter_changed: Callable,
        group_names: Optional[List[str]] = None,
        group_colors: Optional[Dict[str, str]] = None,
        header_text: str = "Filter Results",
        empty_message: str = "No items to filter.\nAdd items first."
    ):
        """
        Initialize the hierarchical filter dropdown.

        Args:
            parent: Parent widget to attach to
            theme_manager: Theme manager for colors
            on_filter_changed: Callback when filter selection changes
            group_names: List of group names in display order (default: ["High", "Medium", "Low"])
            group_colors: Dict mapping group name to color (default uses risk colors)
            header_text: Text shown in dropdown header
            empty_message: Message shown when no items to filter
        """
        self._parent = parent
        self._theme_manager = theme_manager
        self._on_filter_changed = on_filter_changed
        self._header_text = header_text
        self._empty_message = empty_message

        # Group configuration
        self._group_names = group_names or ["High", "Medium", "Low"]
        self._group_colors = group_colors  # Will use risk colors as fallback

        # Filter state - track what's selected
        self._selected_groups: Set[str] = set(self._group_names)  # All selected by default
        self._selected_items: Set[str] = set()  # All items selected by default
        self._all_items: Dict[str, str] = {}  # item_name -> group mapping

        # Collapsed state for group sections - all expanded by default
        self._collapsed_groups: Set[str] = set()

        # UI elements
        self._dropdown_popup = None
        self._checkbox_icons = {}
        self._checkbox_labels = {}  # Track labels for theme updates
        self._item_frames: Dict[str, List[tk.Frame]] = {}  # Track item frames for collapse/expand

        # Search/filter functionality
        self._search_var = None
        self._search_after_id = None  # For debouncing
        self._group_widgets = {}  # {group: {'frame': frame, 'arrow': arrow, 'checkbox': checkbox, ...}}
        self._item_widgets = {}  # {item_name: {'frame': frame, 'checkbox': checkbox, ...}}

        # Load icons - use 14px to match standard icon sizes
        self._load_checkbox_icons()
        self._filter_icon = self._load_icon("filter", 14)
        self._eraser_icon = self._load_icon("eraser", 14)
        self._search_icon = self._load_icon("magnifying-glass", 14)

        # Create the icon buttons
        self._create_filter_buttons()

    def _get_group_color(self, group_name: str) -> str:
        """Get color for a group, using custom colors or risk colors as fallback."""
        if self._group_colors and group_name in self._group_colors:
            return self._group_colors[group_name]

        colors = self._theme_manager.colors
        # Fallback to risk colors for common severity names
        risk_map = {
            "High": colors.get('risk_high', '#ef4444'),
            "Medium": colors.get('risk_medium', '#f59e0b'),
            "Low": colors.get('risk_low', '#22c55e')
        }
        return risk_map.get(group_name, colors.get('text_primary', '#ffffff'))

    def _load_checkbox_icons(self):
        """Load checkbox SVG icons for checked, unchecked, and partial states."""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return

        icons_dir = Path(__file__).parent.parent / "assets" / "Tool Icons"
        is_dark = self._theme_manager.is_dark

        # Select icons based on theme (includes partial checkbox for hierarchical parent items)
        box_name = "box-dark" if is_dark else "box"
        checked_name = "box-checked-dark" if is_dark else "box-checked"
        partial_name = "box-partial-dark" if is_dark else "box-partial"

        for name, key in [(box_name, "unchecked"), (checked_name, "checked"), (partial_name, "partial")]:
            svg_path = icons_dir / f"{name}.svg"
            if svg_path.exists():
                try:
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=64,  # Render at 4x for quality
                        output_height=64
                    )
                    img = Image.open(io.BytesIO(png_data))
                    img = img.resize((16, 16), Image.Resampling.LANCZOS)
                    self._checkbox_icons[key] = ImageTk.PhotoImage(img)
                except Exception:
                    pass

    def _create_filter_buttons(self):
        """Create the filter and clear icon buttons."""
        colors = self._theme_manager.colors

        # Create container frame
        self.frame = ttk.Frame(self._parent, style='Section.TFrame')

        # Filter button - opens dropdown (size=26 to match standard icons)
        self._filter_btn = SquareIconButton(
            self.frame,
            icon=self._filter_icon,
            command=self._toggle_dropdown,
            tooltip_text="Filter",
            size=26,
            radius=6
        )
        self._filter_btn.pack(side=tk.LEFT, padx=(0, 4))

        # Clear/Eraser button - resets all filters (size=26 to match standard icons)
        self._clear_btn = SquareIconButton(
            self.frame,
            icon=self._eraser_icon,
            command=self._reset_filters,
            tooltip_text="Clear Filters",
            size=26,
            radius=6
        )
        self._clear_btn.pack(side=tk.LEFT)

    def _load_icon(self, icon_name: str, size: int = 16):
        """Load an SVG icon."""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return None

        icons_dir = Path(__file__).parent.parent / "assets" / "Tool Icons"
        svg_path = icons_dir / f"{icon_name}.svg"

        if svg_path.exists():
            try:
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=size * 4,
                    output_height=size * 4
                )
                img = Image.open(io.BytesIO(png_data))
                img = img.resize((size, size), Image.Resampling.LANCZOS)
                return ImageTk.PhotoImage(img)
            except Exception:
                pass
        return None

    def _toggle_dropdown(self):
        """Toggle the dropdown popup visibility."""
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._close_dropdown()
        else:
            self._open_dropdown()

    def _open_dropdown(self):
        """Open the dropdown popup with scrollable content and collapsible sections."""
        if self._dropdown_popup:
            self._close_dropdown()

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create popup window
        self._dropdown_popup = tk.Toplevel(self._parent)
        self._dropdown_popup.withdraw()  # Hide until positioned
        self._dropdown_popup.overrideredirect(True)  # No window decorations

        # Configure popup background
        popup_bg = colors.get('surface', colors['background'])
        border_color = colors['border']

        # Create border frame
        border_frame = tk.Frame(
            self._dropdown_popup,
            bg=border_color,
            padx=1,
            pady=1
        )
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Main content frame with minimum width
        main_frame = tk.Frame(border_frame, bg=popup_bg)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header row with label and search on same line
        header_frame = tk.Frame(main_frame, bg=popup_bg)
        header_frame.pack(fill=tk.X, padx=8, pady=(8, 0))

        header = tk.Label(
            header_frame,
            text=self._header_text,
            font=('Segoe UI', 10, 'bold'),
            bg=popup_bg,
            fg=colors['text_primary']
        )
        header.pack(side=tk.LEFT)

        # Container for icon + search box on the right
        search_container = tk.Frame(header_frame, bg=popup_bg)
        search_container.pack(side=tk.RIGHT, padx=(16, 0))  # Minimum spacing from title
        self._search_container = search_container

        # Magnifying glass icon (outside the search box)
        self._search_icon_label = tk.Label(
            search_container,
            bg=popup_bg
        )
        if self._search_icon:
            self._search_icon_label.configure(image=self._search_icon)
            self._search_icon_label._icon_ref = self._search_icon
        self._search_icon_label.pack(side=tk.LEFT, padx=(0, 4))

        # Search entry with border matching the filter frame border
        entry_border = tk.Frame(search_container, bg=colors['border'])
        entry_border.pack(side=tk.LEFT)
        self._search_entry_border = entry_border

        # Inner frame for padding
        entry_bg = colors['background']
        entry_inner = tk.Frame(entry_border, bg=entry_bg)
        entry_inner.pack(padx=1, pady=1)
        self._search_entry_inner = entry_inner

        self._search_var = tk.StringVar()
        self._search_entry = tk.Entry(
            entry_inner,
            textvariable=self._search_var,
            font=('Segoe UI', 9),
            width=24,
            bg=entry_bg,
            fg=colors['text_primary'],
            insertbackground=colors['text_primary'],
            relief=tk.FLAT,
            highlightthickness=0
        )
        self._search_entry.pack(padx=4, pady=4)

        # Bind search with debounce
        self._search_var.trace_add('write', self._on_search_changed)

        # Separator
        sep = tk.Frame(main_frame, bg=border_color, height=1)
        sep.pack(fill=tk.X, padx=8, pady=(8, 4))

        # Check if there are any items to filter
        if not self._all_items:
            # Show empty state message
            empty_frame = tk.Frame(main_frame, bg=popup_bg)
            empty_frame.pack(fill=tk.X, padx=8, pady=16)
            tk.Label(
                empty_frame,
                text=self._empty_message,
                font=('Segoe UI', 9, 'italic'),
                bg=popup_bg,
                fg=colors['text_muted'],
                justify=tk.CENTER
            ).pack(anchor=tk.CENTER)

            # Position and show popup
            self._position_and_show_popup()
            return

        # Scrollable content area with max height
        MAX_HEIGHT = 350

        # Create canvas for scrolling
        canvas_frame = tk.Frame(main_frame, bg=popup_bg)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=8)

        canvas = tk.Canvas(canvas_frame, bg=popup_bg, highlightthickness=0, width=320)
        scrollbar = ThemedScrollbar(canvas_frame, command=canvas.yview,
                                    theme_manager=self._theme_manager, width=12)

        # Inner frame for content
        inner_frame = tk.Frame(canvas, bg=popup_bg)
        canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

        # Store references for height recalculation on expand/collapse
        self._dropdown_canvas = canvas
        self._dropdown_inner_frame = inner_frame
        self._dropdown_scrollbar = scrollbar
        self._dropdown_max_height = MAX_HEIGHT

        self._checkbox_labels = {}
        self._item_frames = {}
        self._group_widgets = {}
        self._item_widgets = {}

        for group in self._group_names:
            # Get items for this group
            items_in_group = [i for i, g in self._all_items.items() if g == group]

            if not items_in_group and group not in self._selected_groups:
                continue  # Skip empty groups that aren't checked

            # Group row (parent)
            group_frame = tk.Frame(inner_frame, bg=popup_bg)
            group_frame.pack(fill=tk.X, pady=2)

            # Expand/Collapse arrow
            is_collapsed = group in self._collapsed_groups
            arrow_text = ">" if is_collapsed else "v"
            arrow_label = tk.Label(
                group_frame,
                text=arrow_text,
                font=('Segoe UI', 8),
                bg=popup_bg,
                fg=colors['text_muted'],
                cursor='hand2'
            )
            arrow_label.pack(side=tk.LEFT, padx=(0, 4))
            arrow_label._group = group
            arrow_label.bind('<Button-1>', lambda e, g=group: self._toggle_collapse(g))

            # Checkbox icon
            is_checked = group in self._selected_groups
            icon = self._checkbox_icons.get('checked' if is_checked else 'unchecked')

            group_checkbox = tk.Label(
                group_frame,
                image=icon,
                bg=popup_bg,
                cursor='hand2'
            )
            group_checkbox.pack(side=tk.LEFT, padx=(0, 6))
            group_checkbox._checkbox_key = f"group:{group}"
            group_checkbox.bind('<Button-1>', lambda e, g=group: self._toggle_group(g))
            self._checkbox_labels[f"group:{group}"] = group_checkbox

            # Group label with color indicator
            group_color = self._get_group_color(group)
            color_dot = tk.Label(
                group_frame,
                text="*",
                font=('Segoe UI', 10),
                fg=group_color,
                bg=popup_bg
            )
            color_dot.pack(side=tk.LEFT, padx=(0, 4))

            group_label = tk.Label(
                group_frame,
                text=f"{group} ({len(items_in_group)})",
                font=('Segoe UI', 9, 'bold'),
                bg=popup_bg,
                fg=colors['text_primary'],
                cursor='hand2'
            )
            group_label.pack(side=tk.LEFT)
            group_label.bind('<Button-1>', lambda e, g=group: self._toggle_collapse(g))

            # Store arrow label for updating
            group_checkbox._arrow_label = arrow_label

            # Create container frame for items
            items_container = tk.Frame(inner_frame, bg=popup_bg)
            items_container.pack(fill=tk.X)
            if is_collapsed:
                items_container.pack_forget()

            self._item_frames[group] = {
                'container': items_container,
                'header': group_frame
            }

            # Store group widgets for search filtering
            self._group_widgets[group] = {
                'frame': group_frame,
                'container': items_container,
                'arrow': arrow_label,
                'checkbox': group_checkbox,
                'items': items_in_group
            }

            # Items under this group (indented)
            for item in sorted(items_in_group):
                item_frame = tk.Frame(items_container, bg=popup_bg)
                item_frame.pack(fill=tk.X, pady=1, padx=(22, 0))

                item_checked = item in self._selected_items
                item_icon = self._checkbox_icons.get('checked' if item_checked else 'unchecked')

                item_checkbox = tk.Label(
                    item_frame,
                    image=item_icon,
                    bg=popup_bg,
                    cursor='hand2'
                )
                item_checkbox.pack(side=tk.LEFT, padx=(0, 6))
                item_checkbox._checkbox_key = f"item:{item}"
                item_checkbox.bind('<Button-1>', lambda e, i=item: self._toggle_item(i))
                self._checkbox_labels[f"item:{item}"] = item_checkbox

                # Item name
                item_label = tk.Label(
                    item_frame,
                    text=item,
                    font=('Segoe UI', 9),
                    bg=popup_bg,
                    fg=colors['text_secondary'],
                    cursor='hand2',
                    anchor=tk.W
                )
                item_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
                item_label.bind('<Button-1>', lambda e, i=item: self._toggle_item(i))

                # Store item widgets for search filtering
                self._item_widgets[item] = {
                    'frame': item_frame,
                    'checkbox': item_checkbox,
                    'label': item_label,
                    'group': group
                }

        # Ensure partial checkbox icons are shown correctly for group parents
        self._update_checkboxes()

        # Update canvas scroll region after content is added
        inner_frame.update_idletasks()
        content_height = inner_frame.winfo_reqheight()

        # Only show scrollbar if content exceeds max height
        if content_height > MAX_HEIGHT:
            canvas.configure(height=MAX_HEIGHT, yscrollcommand=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            canvas.configure(height=content_height)
            canvas.pack(fill=tk.BOTH, expand=True)

        canvas.configure(scrollregion=canvas.bbox("all"))

        # Configure canvas width to match frame
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind('<Configure>', on_canvas_configure)

        # Mouse wheel scrolling
        def on_mousewheel(event):
            if content_height > MAX_HEIGHT:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all('<MouseWheel>', on_mousewheel)
        self._dropdown_popup._on_mousewheel = on_mousewheel

        # Bottom actions separator (outside scrollable area)
        footer_frame = tk.Frame(main_frame, bg=popup_bg)
        footer_frame.pack(fill=tk.X, padx=8, pady=(4, 8))

        sep2 = tk.Frame(footer_frame, bg=border_color, height=1)
        sep2.pack(fill=tk.X, pady=(0, 8))

        # Quick actions row
        actions_frame = tk.Frame(footer_frame, bg=popup_bg)
        actions_frame.pack(fill=tk.X)

        select_all_btn = tk.Label(
            actions_frame,
            text="Select All",
            font=('Segoe UI', 9),
            fg=colors['title_color'],
            bg=popup_bg,
            cursor='hand2'
        )
        select_all_btn.pack(side=tk.LEFT)
        select_all_btn.bind('<Button-1>', lambda e: self._select_all())

        tk.Label(actions_frame, text="  |  ", bg=popup_bg, fg=colors['text_muted']).pack(side=tk.LEFT)

        clear_all_btn = tk.Label(
            actions_frame,
            text="Clear All",
            font=('Segoe UI', 9),
            fg=colors['title_color'],
            bg=popup_bg,
            cursor='hand2'
        )
        clear_all_btn.pack(side=tk.LEFT)
        clear_all_btn.bind('<Button-1>', lambda e: self._clear_all())

        # Position and show popup
        self._position_and_show_popup()

    def _position_and_show_popup(self):
        """Position popup below button, right-anchored to filter icon, and show it."""
        self._dropdown_popup.update_idletasks()
        popup_width = self._dropdown_popup.winfo_reqwidth()
        btn_right_x = self._filter_btn.winfo_rootx() + self._filter_btn.winfo_width()
        btn_y = self._filter_btn.winfo_rooty() + self._filter_btn.winfo_height()

        # Anchor right edge of popup to right edge of button
        popup_x = btn_right_x - popup_width
        self._dropdown_popup.geometry(f"+{popup_x}+{btn_y}")
        self._dropdown_popup.deiconify()
        self._dropdown_popup.lift()
        self._dropdown_popup.focus_set()

        # Bind click outside to close
        self._parent.winfo_toplevel().bind('<Button-1>', self._on_click_outside, add='+')
        self._parent.winfo_toplevel().bind('<Configure>', self._on_window_configure, add='+')

    def _on_window_configure(self, event):
        """Handle parent window move/resize to keep dropdown anchored."""
        if not self._dropdown_popup or not self._dropdown_popup.winfo_exists():
            return

        try:
            popup_width = self._dropdown_popup.winfo_width()
            btn_right_x = self._filter_btn.winfo_rootx() + self._filter_btn.winfo_width()
            btn_y = self._filter_btn.winfo_rooty() + self._filter_btn.winfo_height()
            popup_x = btn_right_x - popup_width
            self._dropdown_popup.geometry(f"+{popup_x}+{btn_y}")
        except Exception:
            pass

    def _on_click_outside(self, event):
        """Handle click outside the dropdown."""
        if not self._dropdown_popup or not self._dropdown_popup.winfo_exists():
            return

        x, y = event.x_root, event.y_root
        dx = self._dropdown_popup.winfo_rootx()
        dy = self._dropdown_popup.winfo_rooty()
        dw = self._dropdown_popup.winfo_width()
        dh = self._dropdown_popup.winfo_height()

        bx = self._filter_btn.winfo_rootx()
        by = self._filter_btn.winfo_rooty()
        bw = self._filter_btn.winfo_width()
        bh = self._filter_btn.winfo_height()

        if not (dx <= x <= dx + dw and dy <= y <= dy + dh) and \
           not (bx <= x <= bx + bw and by <= y <= by + bh):
            self._close_dropdown()

    def _close_dropdown(self):
        """Close the dropdown popup."""
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            try:
                self._dropdown_popup.unbind_all('<MouseWheel>')
            except Exception:
                pass
            self._dropdown_popup.destroy()
        self._dropdown_popup = None

        try:
            self._parent.winfo_toplevel().unbind('<Button-1>')
            self._parent.winfo_toplevel().unbind('<Configure>')
        except Exception:
            pass

        if self._search_after_id:
            try:
                self._parent.after_cancel(self._search_after_id)
            except Exception:
                pass
            self._search_after_id = None

    def _on_search_changed(self, *args):
        """Handle search text change with debounce (500ms delay)."""
        if self._search_after_id:
            try:
                self._parent.after_cancel(self._search_after_id)
            except Exception:
                pass
        self._search_after_id = self._parent.after(500, self._apply_search_filter)

    def _apply_search_filter(self):
        """Apply the search filter to show/hide matching items."""
        self._search_after_id = None

        if not self._search_var:
            return

        search_text = self._search_var.get().strip().lower()

        if not search_text:
            self._show_all_items()
            return

        matched_items = set()
        for item_name in self._item_widgets:
            if search_text in item_name.lower():
                matched_items.add(item_name)

        groups_with_matches = set()

        for item_name, widgets in self._item_widgets.items():
            item_frame = widgets['frame']
            group = widgets['group']

            if item_name in matched_items:
                item_frame.pack(fill=tk.X, pady=1, padx=(22, 0))
                groups_with_matches.add(group)
            else:
                item_frame.pack_forget()

        for group, widgets in self._group_widgets.items():
            group_frame = widgets['frame']
            container = widgets['container']

            if group in groups_with_matches:
                group_frame.pack(fill=tk.X, pady=2)
                if group not in self._collapsed_groups:
                    container.pack(fill=tk.X)
            else:
                group_frame.pack_forget()
                container.pack_forget()

    def _show_all_items(self):
        """Show all items, respecting collapse state."""
        for group, widgets in self._group_widgets.items():
            group_frame = widgets['frame']
            container = widgets['container']

            group_frame.pack(fill=tk.X, pady=2)

            if group in self._collapsed_groups:
                container.pack_forget()
            else:
                container.pack(fill=tk.X)

        for item_name, widgets in self._item_widgets.items():
            item_frame = widgets['frame']
            group = widgets['group']

            if group not in self._collapsed_groups:
                item_frame.pack(fill=tk.X, pady=1, padx=(22, 0))

    def _toggle_group(self, group: str):
        """Toggle a group's selection."""
        if group in self._selected_groups:
            self._selected_groups.remove(group)
            items_to_remove = [i for i, g in self._all_items.items() if g == group]
            self._selected_items -= set(items_to_remove)
        else:
            self._selected_groups.add(group)
            items_to_add = [i for i, g in self._all_items.items() if g == group]
            self._selected_items |= set(items_to_add)

        self._update_checkboxes()
        self._on_filter_changed()

    def _toggle_item(self, item: str):
        """Toggle an item's selection."""
        if item in self._selected_items:
            self._selected_items.remove(item)
        else:
            self._selected_items.add(item)

        group = self._all_items.get(item)
        if group:
            items_in_group = [i for i, g in self._all_items.items() if g == group]
            selected_in_group = [i for i in items_in_group if i in self._selected_items]

            if len(selected_in_group) == len(items_in_group):
                self._selected_groups.add(group)
            else:
                self._selected_groups.discard(group)

        self._update_checkboxes()
        self._on_filter_changed()

    def _toggle_collapse(self, group: str):
        """Toggle the collapsed/expanded state of a group section."""
        if group in self._collapsed_groups:
            self._collapsed_groups.remove(group)
            if group in self._item_frames:
                frame_info = self._item_frames[group]
                container = frame_info['container']
                header = frame_info['header']
                container.pack(fill=tk.X, after=header)
        else:
            self._collapsed_groups.add(group)
            if group in self._item_frames:
                frame_info = self._item_frames[group]
                container = frame_info['container']
                container.pack_forget()

        for key, label in self._checkbox_labels.items():
            if key == f"group:{group}" and hasattr(label, '_arrow_label'):
                arrow_text = ">" if group in self._collapsed_groups else "v"
                label._arrow_label.configure(text=arrow_text)
                break

        self._update_dropdown_height()

    def _update_dropdown_height(self):
        """Recalculate and apply the dropdown height after expand/collapse changes."""
        if not hasattr(self, '_dropdown_canvas') or not self._dropdown_popup:
            return

        canvas = self._dropdown_canvas
        inner_frame = self._dropdown_inner_frame
        scrollbar = self._dropdown_scrollbar
        max_height = self._dropdown_max_height

        inner_frame.update_idletasks()
        content_height = inner_frame.winfo_reqheight()

        scrollbar_was_shown = scrollbar.winfo_ismapped()
        needs_scrollbar = content_height > max_height

        if needs_scrollbar:
            canvas.configure(height=max_height, yscrollcommand=scrollbar.set)
            if not scrollbar_was_shown:
                canvas.pack_forget()
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            canvas.configure(height=content_height, yscrollcommand='')
            if scrollbar_was_shown:
                scrollbar.pack_forget()
                canvas.pack_forget()
                canvas.pack(fill=tk.BOTH, expand=True)

        canvas.configure(scrollregion=canvas.bbox("all"))

        self._dropdown_popup.update_idletasks()
        popup_width = self._dropdown_popup.winfo_reqwidth()
        btn_right_x = self._filter_btn.winfo_rootx() + self._filter_btn.winfo_width()
        btn_y = self._filter_btn.winfo_rooty() + self._filter_btn.winfo_height()
        popup_x = btn_right_x - popup_width
        self._dropdown_popup.geometry(f"+{popup_x}+{btn_y}")

    def _update_checkboxes(self):
        """Update checkbox icons based on selection state."""
        for key, label in self._checkbox_labels.items():
            try:
                if key.startswith("group:"):
                    group = key.split(":")[1]
                    items_in_group = [i for i, g in self._all_items.items() if g == group]
                    selected_in_group = [i for i in items_in_group if i in self._selected_items]

                    if len(selected_in_group) == len(items_in_group) and len(items_in_group) > 0:
                        icon_key = 'checked'
                    elif len(selected_in_group) > 0:
                        icon_key = 'partial'
                    else:
                        icon_key = 'unchecked'
                else:
                    item = key.split(":")[1]
                    icon_key = 'checked' if item in self._selected_items else 'unchecked'

                icon = self._checkbox_icons.get(icon_key)
                if icon:
                    label.configure(image=icon)
            except Exception:
                pass

    def _select_all(self):
        """Select all groups and items."""
        self._selected_groups = set(self._group_names)
        self._selected_items = set(self._all_items.keys())
        self._update_checkboxes()
        self._on_filter_changed()

    def _clear_all(self):
        """Clear all selections (used in dropdown)."""
        self._selected_groups.clear()
        self._selected_items.clear()
        self._update_checkboxes()
        self._on_filter_changed()

    def _reset_filters(self):
        """Reset all filters to show everything (used by Clear button)."""
        self._select_all()
        self._close_dropdown()

    def set_items(self, items_by_group: Dict[str, List[str]], reset_selection: bool = True):
        """
        Set the available items grouped by group name.

        Args:
            items_by_group: Dict mapping group name to list of item names
            reset_selection: If True, reset selection to all items. If False, preserve current selection.
        """
        self._all_items = {}
        for group, items in items_by_group.items():
            for item in items:
                self._all_items[item] = group

        # Only reset selection if requested (default behavior)
        if reset_selection:
            self._selected_groups = set(self._group_names)
            self._selected_items = set(self._all_items.keys())

    def get_selected_groups(self) -> Set[str]:
        """Get currently selected groups."""
        return self._selected_groups.copy()

    def get_selected_items(self) -> Set[str]:
        """Get currently selected items."""
        return self._selected_items.copy()

    def is_item_visible(self, item_name: str) -> bool:
        """Check if an item should be visible based on current filters."""
        return item_name in self._selected_items

    def on_theme_changed(self):
        """Handle theme change - reload icons and update colors."""
        self._load_checkbox_icons()
        self._update_checkboxes()

        self._filter_icon = self._load_icon("filter", 14)
        self._eraser_icon = self._load_icon("eraser", 14)
        self._search_icon = self._load_icon("magnifying-glass", 14)

        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._close_dropdown()

    def pack(self, **kwargs):
        """Pack the filter dropdown frame."""
        self.frame.pack(**kwargs)

    def grid(self, **kwargs):
        """Grid the filter dropdown frame."""
        self.frame.grid(**kwargs)
