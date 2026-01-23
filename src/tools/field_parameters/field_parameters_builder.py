"""
Field Parameters Builder & Category Manager
Built by Reid Havens of Analytic Endeavors

Parameter builder with drag-drop reordering and category manager.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Optional, Tuple, TYPE_CHECKING, Dict
import logging
from pathlib import Path
import io

from tools.field_parameters.field_parameters_core import FieldItem, CategoryLevel
from tools.field_parameters.dialogs import AddLabelDialog, FieldCategoryEditorDialog
from tools.field_parameters.panels import CategoryManagerPanel
from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, ThemedScrollbar, ThemedMessageBox, LabeledRadioGroup, SquareIconButton, Tooltip, ThemedContextMenu, ThemedInputDialog
from core.filter_dropdown import HierarchicalFilterDropdown
from core.constants import ErrorMessages

# Optional PIL/CairoSVG for icon loading
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except (ImportError, OSError):
    CAIROSVG_AVAILABLE = False

if TYPE_CHECKING:
    from tools.field_parameters.field_parameters_ui import FieldParametersTab


class FieldParameterFilterDropdown(HierarchicalFilterDropdown):
    """
    Specialized filter dropdown for Field Parameters tool.

    Extends HierarchicalFilterDropdown to add:
    - "(All)" option - shows all fields regardless of category
    - "(Uncategorized)" option - shows fields with no category label assigned
    - Groups are category columns (e.g., "Property Metrics Category")
    - Items are labels within each category
    """

    SPECIAL_ALL = "__all__"
    SPECIAL_UNCATEGORIZED = "__uncategorized__"

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager,
        on_filter_changed,
    ):
        """
        Initialize the Field Parameter filter dropdown.

        Args:
            parent: Parent widget to attach to
            theme_manager: Theme manager for colors
            on_filter_changed: Callback when filter selection changes
        """
        # Initialize with empty groups - will be populated via set_category_levels
        super().__init__(
            parent=parent,
            theme_manager=theme_manager,
            on_filter_changed=on_filter_changed,
            group_names=[],  # Will be set via set_category_levels
            group_colors=None,  # Use default text color for category columns
            header_text="Filter by Category",
            empty_message="No categories defined.\nAdd category columns first."
        )

        # Track special filter states
        self._all_selected = True  # Start with "All" selected
        self._uncategorized_selected = False

        # Map from column index to group name
        self._col_idx_to_group: Dict[int, str] = {}
        self._group_to_col_idx: Dict[str, int] = {}

        # Track category levels for filter state building
        self._category_levels: List[CategoryLevel] = []

    def set_category_levels(self, category_levels: List['CategoryLevel']):
        """
        Set the category levels and rebuild the dropdown content.

        Args:
            category_levels: List of CategoryLevel objects from Field Parameters
        """
        # Check if category structure changed - only reset filter state if it did
        old_col_names = [cl.column_name for cl in self._category_levels] if self._category_levels else []
        new_col_names = [cl.column_name for cl in category_levels] if category_levels else []
        structure_changed = old_col_names != new_col_names

        self._category_levels = category_levels

        # Build group names and items from category levels
        group_names = []
        items_by_group: Dict[str, List[str]] = {}
        self._col_idx_to_group = {}
        self._group_to_col_idx = {}

        for col_idx, level in enumerate(category_levels):
            if level.is_calculated or not level.labels:
                continue

            group_name = level.name
            group_names.append(group_name)
            items_by_group[group_name] = list(level.labels)
            self._col_idx_to_group[col_idx] = group_name
            self._group_to_col_idx[group_name] = col_idx

        # Update base class state
        self._group_names = group_names

        # Only reset filter state if category structure actually changed
        if structure_changed:
            self._all_selected = True
            self._uncategorized_selected = False
            self._selected_groups = set(group_names)
            self._collapsed_groups = set()  # Start expanded

        # Set items (preserves selection state if structure unchanged)
        self.set_items(items_by_group, reset_selection=structure_changed)

    def _open_dropdown(self):
        """Override to add special (All) and (Uncategorized) options at the top."""
        # Close existing popup if any
        if self._dropdown_popup:
            self._close_dropdown()

        colors = self._theme_manager.colors

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

        # Main content frame
        main_frame = tk.Frame(border_frame, bg=popup_bg)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header row with label and search
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

        # Search container on the right
        search_container = tk.Frame(header_frame, bg=popup_bg)
        search_container.pack(side=tk.RIGHT, padx=(16, 0))
        self._search_container = search_container

        # Magnifying glass icon
        self._search_icon_label = tk.Label(search_container, bg=popup_bg)
        if self._search_icon:
            self._search_icon_label.configure(image=self._search_icon)
            self._search_icon_label._icon_ref = self._search_icon
        self._search_icon_label.pack(side=tk.LEFT, padx=(0, 4))

        # Search entry with border
        entry_border = tk.Frame(search_container, bg=colors['border'])
        entry_border.pack(side=tk.LEFT)
        self._search_entry_border = entry_border

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

        # ========================================
        # SPECIAL OPTIONS: (All) and (Uncategorized)
        # ========================================
        special_frame = tk.Frame(main_frame, bg=popup_bg)
        special_frame.pack(fill=tk.X, padx=8, pady=(4, 0))

        # "(All)" checkbox
        all_row = tk.Frame(special_frame, bg=popup_bg)
        all_row.pack(fill=tk.X, pady=2)

        all_icon = self._checkbox_icons.get('checked' if self._all_selected else 'unchecked')
        all_checkbox = tk.Label(all_row, image=all_icon, bg=popup_bg, cursor='hand2')
        all_checkbox.pack(side=tk.LEFT, padx=(0, 6))
        all_checkbox._checkbox_key = self.SPECIAL_ALL
        all_checkbox.bind('<Button-1>', lambda e: self._toggle_special(self.SPECIAL_ALL))
        self._checkbox_labels[self.SPECIAL_ALL] = all_checkbox

        all_label = tk.Label(
            all_row,
            text="(All)",
            font=('Segoe UI', 9, 'bold'),
            bg=popup_bg,
            fg=colors['text_primary'],
            cursor='hand2'
        )
        all_label.pack(side=tk.LEFT)
        all_label.bind('<Button-1>', lambda e: self._toggle_special(self.SPECIAL_ALL))

        # "(Uncategorized)" checkbox
        uncat_row = tk.Frame(special_frame, bg=popup_bg)
        uncat_row.pack(fill=tk.X, pady=2)

        uncat_icon = self._checkbox_icons.get('checked' if self._uncategorized_selected else 'unchecked')
        uncat_checkbox = tk.Label(uncat_row, image=uncat_icon, bg=popup_bg, cursor='hand2')
        uncat_checkbox.pack(side=tk.LEFT, padx=(0, 6))
        uncat_checkbox._checkbox_key = self.SPECIAL_UNCATEGORIZED
        uncat_checkbox.bind('<Button-1>', lambda e: self._toggle_special(self.SPECIAL_UNCATEGORIZED))
        self._checkbox_labels[self.SPECIAL_UNCATEGORIZED] = uncat_checkbox

        uncat_label = tk.Label(
            uncat_row,
            text="(Uncategorized)",
            font=('Segoe UI', 9),
            bg=popup_bg,
            fg=colors['text_primary'],
            cursor='hand2'
        )
        uncat_label.pack(side=tk.LEFT)
        uncat_label.bind('<Button-1>', lambda e: self._toggle_special(self.SPECIAL_UNCATEGORIZED))

        # Separator before category groups
        if self._group_names:
            sep2 = tk.Frame(main_frame, bg=border_color, height=1)
            sep2.pack(fill=tk.X, padx=8, pady=(8, 4))

        # Check if there are any items to filter
        if not self._all_items:
            # Position and show popup - no category items
            self._position_and_show_popup()
            return

        # ========================================
        # CATEGORY GROUPS (collapsible sections)
        # ========================================
        MAX_HEIGHT = 300

        # Create canvas for scrolling
        canvas_frame = tk.Frame(main_frame, bg=popup_bg)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=8)

        canvas = tk.Canvas(canvas_frame, bg=popup_bg, highlightthickness=0, width=320)
        scrollbar = ThemedScrollbar(canvas_frame, command=canvas.yview,
                                    theme_manager=self._theme_manager, width=12)

        # Inner frame for content
        inner_frame = tk.Frame(canvas, bg=popup_bg)
        canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

        # Store references
        self._dropdown_canvas = canvas
        self._dropdown_inner_frame = inner_frame
        self._dropdown_scrollbar = scrollbar
        self._dropdown_max_height = MAX_HEIGHT

        self._item_frames = {}
        self._group_widgets = {}
        self._item_widgets = {}

        for group in self._group_names:
            # Get items for this group
            items_in_group = [i for i, g in self._all_items.items() if g == group]

            if not items_in_group:
                continue

            # Group row (parent category column)
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

            # Checkbox icon for group - determine state
            group_items = [i for i, g in self._all_items.items() if g == group]
            selected_in_group = len([i for i in group_items if i in self._selected_items])

            if selected_in_group == 0:
                group_state = 'unchecked'
            elif selected_in_group == len(group_items):
                group_state = 'checked'
            else:
                group_state = 'partial'

            icon = self._checkbox_icons.get(group_state, self._checkbox_icons.get('unchecked'))

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

            # Group label
            group_label = tk.Label(
                group_frame,
                text=f"{group} ({len(items_in_group)})",
                font=('Segoe UI', 9, 'bold'),
                bg=popup_bg,
                fg=colors['text_primary'],
                cursor='hand2'
            )
            group_label.pack(side=tk.LEFT)
            group_label.bind('<Button-1>', lambda e, g=group: self._toggle_group(g))

            # Items container (create before storing in _group_widgets)
            items_container = tk.Frame(inner_frame, bg=popup_bg)
            if not is_collapsed:
                items_container.pack(fill=tk.X, padx=(20, 0))

            # Store group widgets - include 'container' for base class compatibility
            self._group_widgets[group] = {
                'frame': group_frame,
                'arrow': arrow_label,
                'checkbox': group_checkbox,
                'label': group_label,
                'container': items_container
            }

            # Store item frames with dict structure for base class compatibility
            self._item_frames[group] = {'container': items_container, 'header': group_frame}

            # Attach arrow_label to checkbox for base class _toggle_collapse
            group_checkbox._arrow_label = arrow_label

            # Add individual items (labels within category)
            for item in sorted(items_in_group):
                item_row = tk.Frame(items_container, bg=popup_bg)
                item_row.pack(fill=tk.X, pady=1)

                is_item_selected = item in self._selected_items
                item_icon = self._checkbox_icons.get('checked' if is_item_selected else 'unchecked')

                item_checkbox = tk.Label(
                    item_row,
                    image=item_icon,
                    bg=popup_bg,
                    cursor='hand2'
                )
                item_checkbox.pack(side=tk.LEFT, padx=(0, 6))
                item_checkbox._checkbox_key = f"item:{item}"
                item_checkbox.bind('<Button-1>', lambda e, i=item: self._toggle_item(i))
                self._checkbox_labels[f"item:{item}"] = item_checkbox

                item_label = tk.Label(
                    item_row,
                    text=item,
                    font=('Segoe UI', 9),
                    bg=popup_bg,
                    fg=colors['text_primary'],
                    cursor='hand2'
                )
                item_label.pack(side=tk.LEFT)
                item_label.bind('<Button-1>', lambda e, i=item: self._toggle_item(i))

                self._item_widgets[item] = {
                    'frame': item_row,
                    'checkbox': item_checkbox,
                    'label': item_label,
                    'group': group  # Required by base class _apply_search_filter
                }

        # Configure canvas scrolling
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Limit height
            content_height = inner_frame.winfo_reqheight()
            canvas_height = min(content_height, MAX_HEIGHT)
            canvas.configure(height=canvas_height)
            # Show scrollbar only if needed
            if content_height > MAX_HEIGHT:
                canvas.configure(yscrollcommand=scrollbar.set)  # Connect scrollbar thumb
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            else:
                canvas.configure(yscrollcommand='')  # Disconnect when not needed
                scrollbar.pack_forget()

        inner_frame.bind("<Configure>", configure_scroll_region)

        def configure_canvas_width(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", configure_canvas_width)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)
        self._dropdown_popup.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Position and show popup
        self._position_and_show_popup()

    def _toggle_special(self, special_type: str):
        """Toggle special (All) or (Uncategorized) filter options."""
        if special_type == self.SPECIAL_ALL:
            self._all_selected = not self._all_selected
            if self._all_selected:
                # Selecting "(All)" clears other selections
                self._uncategorized_selected = False
                self._selected_groups = set(self._group_names)
                self._selected_items = set(self._all_items.keys())
        elif special_type == self.SPECIAL_UNCATEGORIZED:
            self._uncategorized_selected = not self._uncategorized_selected
            if self._uncategorized_selected:
                # Selecting Uncategorized clears "(All)"
                self._all_selected = False

        self._update_checkboxes()
        self._on_filter_changed()

    def _toggle_group(self, group: str):
        """Override to handle "(All)" deselection when toggling groups."""
        # If selecting a group, deselect "(All)"
        if self._all_selected:
            self._all_selected = False

        # Toggle all items in the group
        group_items = [i for i, g in self._all_items.items() if g == group]
        all_selected = all(i in self._selected_items for i in group_items)

        if all_selected:
            # Deselect all in group
            for item in group_items:
                self._selected_items.discard(item)
        else:
            # Select all in group
            for item in group_items:
                self._selected_items.add(item)

        self._update_checkboxes()
        self._on_filter_changed()

    def _toggle_item(self, item: str):
        """Override to handle "(All)" deselection when toggling items."""
        # If selecting an item, deselect "(All)"
        if self._all_selected:
            self._all_selected = False

        # Toggle item
        if item in self._selected_items:
            self._selected_items.discard(item)
        else:
            self._selected_items.add(item)

        self._update_checkboxes()
        self._on_filter_changed()

    def _update_checkboxes(self):
        """Override to update special checkboxes along with regular ones."""
        # Helper to safely configure widget (check if it still exists)
        def safe_configure(widget, **kwargs):
            try:
                if widget.winfo_exists():
                    widget.configure(**kwargs)
            except Exception:
                pass  # Widget was destroyed

        # Update special checkboxes
        if self.SPECIAL_ALL in self._checkbox_labels:
            icon = self._checkbox_icons.get('checked' if self._all_selected else 'unchecked')
            if icon:
                safe_configure(self._checkbox_labels[self.SPECIAL_ALL], image=icon)

        if self.SPECIAL_UNCATEGORIZED in self._checkbox_labels:
            icon = self._checkbox_icons.get('checked' if self._uncategorized_selected else 'unchecked')
            if icon:
                safe_configure(self._checkbox_labels[self.SPECIAL_UNCATEGORIZED], image=icon)

        # Update group and item checkboxes
        for group in self._group_names:
            group_items = [i for i, g in self._all_items.items() if g == group]
            selected_in_group = len([i for i in group_items if i in self._selected_items])

            if selected_in_group == 0:
                state = 'unchecked'
            elif selected_in_group == len(group_items):
                state = 'checked'
            else:
                state = 'partial'

            key = f"group:{group}"
            if key in self._checkbox_labels:
                icon = self._checkbox_icons.get(state, self._checkbox_icons.get('unchecked'))
                if icon:
                    safe_configure(self._checkbox_labels[key], image=icon)

        # Update item checkboxes
        for item in self._all_items:
            key = f"item:{item}"
            if key in self._checkbox_labels:
                is_selected = item in self._selected_items
                icon = self._checkbox_icons.get('checked' if is_selected else 'unchecked')
                if icon:
                    safe_configure(self._checkbox_labels[key], image=icon)

    def _reset_filters(self):
        """Override to reset to "(All)" state."""
        self._all_selected = True
        self._uncategorized_selected = False
        self._selected_groups = set(self._group_names)
        self._selected_items = set(self._all_items.keys())
        self._update_checkboxes()
        self._close_dropdown()
        self._on_filter_changed()

    def _position_and_show_popup(self):
        """Override to position popup left-anchored instead of right-anchored."""
        self._dropdown_popup.update_idletasks()
        btn_x = self._filter_btn.winfo_rootx()
        btn_y = self._filter_btn.winfo_rooty() + self._filter_btn.winfo_height()

        # Left-anchor popup to left edge of button
        self._dropdown_popup.geometry(f"+{btn_x}+{btn_y}")
        self._dropdown_popup.deiconify()
        self._dropdown_popup.lift()
        self._dropdown_popup.focus_set()

        # Bind click outside to close
        self._parent.winfo_toplevel().bind('<Button-1>', self._on_click_outside, add='+')
        self._parent.winfo_toplevel().bind('<Configure>', self._on_window_configure, add='+')

    def _on_window_configure(self, event):
        """Override to keep popup left-anchored on window move/resize."""
        if not self._dropdown_popup or not self._dropdown_popup.winfo_exists():
            return

        try:
            btn_x = self._filter_btn.winfo_rootx()
            btn_y = self._filter_btn.winfo_rooty() + self._filter_btn.winfo_height()
            self._dropdown_popup.geometry(f"+{btn_x}+{btn_y}")
        except Exception:
            pass

    def _update_dropdown_height(self):
        """Override to reposition popup left-anchored after height changes."""
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

        # Reposition left-anchored
        self._dropdown_popup.update_idletasks()
        btn_x = self._filter_btn.winfo_rootx()
        btn_y = self._filter_btn.winfo_rooty() + self._filter_btn.winfo_height()
        self._dropdown_popup.geometry(f"+{btn_x}+{btn_y}")

    def get_filter_state(self) -> Optional[Dict]:
        """
        Get the current filter state in a format compatible with existing filter logic.

        Returns:
            None = show all ("All" is selected)
            {"__uncategorized__": True} = show uncategorized fields
            {col_idx: set(labels)} = show fields matching any label in any column
        """
        if self._all_selected:
            return None  # Show all

        filter_state = {}

        # Check uncategorized
        if self._uncategorized_selected:
            filter_state["__uncategorized__"] = True

        # Build label filters by column index
        for item in self._selected_items:
            if item in self._all_items:
                group = self._all_items[item]
                if group in self._group_to_col_idx:
                    col_idx = self._group_to_col_idx[group]
                    if col_idx not in filter_state:
                        filter_state[col_idx] = set()
                    filter_state[col_idx].add(item)

        # If nothing is selected, return None (show all)
        if not filter_state:
            return None

        return filter_state

    def clear_filters(self):
        """
        Clear all filter selections and select "(All)" to show all fields.
        This is a public method for external callers to reset the filter state.
        """
        self._all_selected = True
        self._uncategorized_selected = False
        self._selected_groups = set(self._group_names)
        self._selected_items = set(self._all_items.keys())

        # Update visual state if dropdown is open
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._update_checkboxes()

        # Update button text
        self._update_button_text()

        # Trigger filter change callback
        self._on_filter_changed()

    def has_categories(self) -> bool:
        """Check if there are any categories with labels."""
        return bool(self._group_names) or len(self._all_items) > 0


class ParameterBuilderPanel(ttk.LabelFrame):
    """Panel for building/editing field parameter with drag-drop"""

    # Virtual scrolling constants
    ROW_HEIGHT = 30  # Approximate height of each field row in pixels
    BUFFER_ROWS = 10  # Extra rows to render above/below viewport for smoother scrolling
    VIRTUAL_SCROLL_THRESHOLD = 100  # Use virtual scrolling when item count exceeds this

    def __init__(self, parent, main_tab: 'FieldParametersTab'):
        super().__init__(parent, style='Section.TLabelframe', padding="12")
        self.main_tab = main_tab
        self.logger = logging.getLogger(__name__)
        self._theme_manager = get_theme_manager()

        # Section header widget references for theme updates
        self._section_header_widgets = []
        self._header_icon = None

        # Button tracking for theme updates
        self._secondary_buttons = []

        # Create and set the section header labelwidget
        self._create_section_header(parent, "Parameter Builder", "Field Parameter")

        self.field_items: List[FieldItem] = []
        self.field_widgets: Dict[int, dict] = {}  # field_id -> {widgets}
        self.drag_data = {"item": None, "index": None}

        # Drop indicator
        self._drop_indicator = None
        self._drop_position = None  # Index where drop would occur

        # Category filter state - now supports multiple columns with multiple selections
        # Format: {column_index: set(selected_labels)} or None for show all
        self._category_filters: Optional[Dict[int, set]] = None  # None = show all
        self._category_levels: List['CategoryLevel'] = []  # Available category levels

        # Virtual scrolling state
        self._virtual_mode = False  # True when using virtual scrolling
        self._visible_range = (0, 0)  # (start_index, end_index) of currently rendered items
        self._filtered_items: List[FieldItem] = []  # Cached list of items matching current filter
        self._scroll_job = None  # Debounce scroll updates
        self._widget_pool: List[Dict] = []  # Pool of recycled row widgets for reuse

        # Grouped view state
        self._view_mode = "flat"  # "flat" or "grouped"
        self._group_by_category_idx = 0  # Which category column to group by
        self._combo_to_level_idx = []  # Maps combobox index to _category_levels index
        self._collapsed_groups: set = set()  # Set of collapsed group labels
        self._group_widgets: Dict[str, dict] = {}  # group_label -> {header_frame, fields_frame, ...}
        self._group_order: List[Tuple[float, str]] = []  # [(sort_order, label), ...] for ordering groups
        self._group_drag_data = {"item": None, "index": None, "dragging": False}
        self._group_drop_indicator = None
        self._selected_group: Optional[str] = None  # Currently selected group label in grouped view

        # Load icons for toolbar buttons
        self._button_icons = {}
        self._load_button_icons()

        self.setup_ui()

    def _load_button_icons(self):
        """Load SVG icons for toolbar buttons"""
        if not PIL_AVAILABLE:
            return

        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"

        icon_mappings = {
            'filter': ('filter', 16),
            'x-button': ('x-button', 14),
            'eraser': ('eraser', 14),
            'reset': ('reset', 16),
        }

        for icon_key, (svg_name, size) in icon_mappings.items():
            svg_path = icons_dir / f"{svg_name}.svg"
            try:
                if CAIROSVG_AVAILABLE and svg_path.exists():
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=size * 4,
                        output_height=size * 4
                    )
                    img = Image.open(io.BytesIO(png_data))
                    img = img.resize((size, size), Image.Resampling.LANCZOS)
                    self._button_icons[icon_key] = ImageTk.PhotoImage(img)
            except Exception as e:
                self.logger.debug(f"Failed to load toolbar icon {svg_name}: {e}")

    def setup_ui(self):
        """Setup parameter builder UI"""
        colors = self._theme_manager.colors

        # CRITICAL: Inner wrapper with Section.TFrame style fills LabelFrame with correct background
        # This is required per UI_DESIGN_PATTERNS.md - without it, section_bg shows in padding area
        self._content_wrapper = ttk.Frame(self, style='Section.TFrame', padding="15")
        self._content_wrapper.pack(fill=tk.BOTH, expand=True)

        # Background color for all tk.Frame widgets
        bg_color = colors['background']
        # Theme-aware disabled colors for buttons
        is_dark = colors.get('background', '') == '#0d0d1a'
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Track inner frames for theme updates
        self._inner_frames = []

        # =========================================================================
        # TOP TOOLBAR - Filter controls (left) and Navigation buttons (right)
        # Always visible
        # =========================================================================
        toolbar1 = tk.Frame(self._content_wrapper, bg=bg_color)
        toolbar1.pack(fill=tk.X, pady=(0, 3))
        self._toolbar1 = toolbar1  # Store reference for dropdown positioning
        self._inner_frames.append(toolbar1)

        # Initialize icon button list
        self._icon_buttons = []

        # Left side - Filter controls (always visible)
        self.filter_frame = tk.Frame(toolbar1, bg=bg_color)
        self.filter_frame.pack(side=tk.LEFT)
        self._inner_frames.append(self.filter_frame)

        # Hierarchical filter dropdown with search, collapsible groups, and SVG checkboxes
        self._filter_dropdown = FieldParameterFilterDropdown(
            parent=self.filter_frame,
            theme_manager=self._theme_manager,
            on_filter_changed=self._on_filter_changed
        )
        self._filter_dropdown.pack(side=tk.LEFT, padx=(0, 8))

        self.filter_count_label = tk.Label(
            self.filter_frame,
            text="",
            font=("Segoe UI", 8, "italic"),
            bg=bg_color, fg=colors['text_muted']
        )
        self.filter_count_label.pack(side=tk.LEFT)

        # Header category dropdown - only visible in grouped view (positioned after filter, before buttons)
        # Separate frame so it appears on the LEFT side when shown
        self._header_group_by_frame = tk.Frame(toolbar1, bg=bg_color)
        # Not packed initially - shown/hidden by _on_view_mode_changed

        self._header_group_by_combo = ttk.Combobox(
            self._header_group_by_frame, width=18, state="readonly", font=("Segoe UI", 8)
        )
        self._header_group_by_combo.pack(side=tk.LEFT)
        self._header_group_by_combo.bind("<<ComboboxSelected>>", self._on_header_group_by_changed)

        # Right side - Up/Down/Delete buttons (always visible)
        arrow_frame = tk.Frame(toolbar1, bg=bg_color)
        arrow_frame.pack(side=tk.RIGHT)
        self._arrow_frame = arrow_frame  # Store for theme updates

        # Delete button first (left position) - matches Category Columns layout
        self.delete_btn = RoundedButton(
            arrow_frame,
            text="\u2716",
            command=self._on_delete_selected,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=26, height=24, radius=5,
            font=('Segoe UI', 8),
            canvas_bg=bg_color
        )
        self.delete_btn.pack(side=tk.LEFT, padx=1)
        self._secondary_buttons.append(self.delete_btn)

        # Dynamic tooltip for delete button based on selection count
        self._delete_tooltip = Tooltip(
            self.delete_btn,
            text=lambda: "Delete Fields" if len(getattr(self, 'selected_items', set())) > 1 else "Delete Field"
        )

        self.move_up_btn = RoundedButton(
            arrow_frame,
            text="\u25B2",
            command=self._move_selected_up,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=26, height=24, radius=5,
            font=('Segoe UI', 8),
            canvas_bg=bg_color
        )
        self.move_up_btn.pack(side=tk.LEFT, padx=1)
        self._secondary_buttons.append(self.move_up_btn)

        self.move_down_btn = RoundedButton(
            arrow_frame,
            text="\u25BC",
            command=self._move_selected_down,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=26, height=24, radius=5,
            font=('Segoe UI', 8),
            canvas_bg=bg_color
        )
        self.move_down_btn.pack(side=tk.LEFT, padx=1)
        self._secondary_buttons.append(self.move_down_btn)

        # Scrollable canvas for fields - use tk.Frame with border
        border_color = colors.get('border', '#3a3a4a')
        canvas_frame = tk.Frame(
            self._content_wrapper, bg=bg_color,
            highlightbackground=border_color, highlightcolor=border_color,
            highlightthickness=1
        )
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        self._inner_frames.append(canvas_frame)
        self._canvas_frame = canvas_frame  # Store for theme updates

        self._canvas_scrollbar = ThemedScrollbar(
            canvas_frame,
            command=self._on_scroll,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._canvas_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas = tk.Canvas(
            canvas_frame,
            yscrollcommand=self._canvas_scrollbar.set,
            highlightthickness=0,
            bg=colors['background']  # Use darkest background so row colors stand out
        )
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame inside canvas to hold fields - use explicit bg to match canvas background
        container_bg = colors['background']
        self.fields_container = tk.Frame(self.canvas, bg=container_bg)
        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.fields_container,
            anchor="nw",
            width=self.canvas.winfo_reqwidth()
        )
        
        # Configure canvas scrolling
        self.fields_container.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Enable mouse wheel scrolling
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.fields_container.bind("<MouseWheel>", self._on_mousewheel)

        # Empty state message - drawn on canvas background (doesn't take up space)
        self._empty_text_id = self.canvas.create_text(
            0, 0,  # Will be repositioned in _on_canvas_configure
            text="No fields added yet.\n\nDouble-click or drag fields from\n'Available Fields' to add them here.",
            font=("Segoe UI", 10, "italic"),
            fill="gray",
            anchor="center",
            justify="center"
        )

        # =========================================================================
        # BOTTOM TOOLBAR - Edit Categories, Revert Names (left), View toggle (right)
        # Disabled until parameter is created/loaded
        # =========================================================================
        self.toolbar_bottom = tk.Frame(self._content_wrapper, bg=bg_color)
        self.toolbar_bottom.pack(fill=tk.X, pady=(8, 0))
        self._inner_frames.append(self.toolbar_bottom)

        # Left side - Action buttons
        left_btn_frame = tk.Frame(self.toolbar_bottom, bg=bg_color)
        left_btn_frame.pack(side=tk.LEFT)
        self._left_btn_frame = left_btn_frame  # Store for theme updates

        self.edit_categories_btn = RoundedButton(
            left_btn_frame,
            text="Edit Categories",
            command=self._on_edit_categories,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=bg_color
        )
        self.edit_categories_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._secondary_buttons.append(self.edit_categories_btn)

        # Revert Names button - text button with tooltip
        self.revert_names_btn = RoundedButton(
            left_btn_frame,
            text="Revert Names",
            command=self._on_revert_all,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=bg_color
        )
        self.revert_names_btn.pack(side=tk.LEFT)
        self._secondary_buttons.append(self.revert_names_btn)
        # Add tooltip
        Tooltip(self.revert_names_btn, "Revert all display names to their original values")

        # Right side - View mode toggle (always visible, disabled until parameter loaded)
        # Use tk.Frame (not ttk) with explicit bg to prevent background mismatch on initial load
        self.view_mode_frame = tk.Frame(self.toolbar_bottom, bg=bg_color)
        self.view_mode_frame.pack(side=tk.RIGHT)  # Pack immediately, disable via _set_radio_group_enabled

        # View: Flat/Grouped radio buttons - FAR RIGHT (packed first so it stays rightmost)
        # Use a tk.Frame (not ttk) with explicit background to match app styling
        self.view_toggle_frame = tk.Frame(self.view_mode_frame, bg=bg_color)
        self.view_toggle_frame.pack(side=tk.RIGHT)

        # Container for grouped-mode controls - child of view_mode_frame, positioned with place()
        # NOT visible initially (flat mode default) - shown/hidden via place/place_forget when switching view modes
        # Using place() (not pack) prevents it from affecting view_mode_frame's natural width
        self._grouped_controls_frame = tk.Frame(self.view_mode_frame, bg=bg_color)
        # Don't show yet - managed by _on_view_mode_changed via place()/place_forget()

        # Category dropdown for grouped view (which category to group by) - always packed inside frame
        self.group_by_combo = ttk.Combobox(
            self._grouped_controls_frame, width=14, state="readonly", font=("Segoe UI", 8)
        )
        self.group_by_combo.pack(side=tk.LEFT, padx=(0, 4), fill=tk.Y, pady=2)
        self.group_by_combo.bind("<<ComboboxSelected>>", self._on_group_by_changed)

        # Align to Category button - always packed inside frame
        self._align_btn_frame = tk.Frame(self._grouped_controls_frame, bg=bg_color)
        self._align_btn_frame.pack(side=tk.LEFT, pady=2)
        self.align_btn = RoundedButton(
            self._align_btn_frame,
            text="Align",
            command=self._on_align_to_category,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=50, height=22, radius=5,
            font=('Segoe UI', 8),
            canvas_bg=bg_color
        )
        self.align_btn.pack()
        self._secondary_buttons.append(self.align_btn)

        self.view_label = tk.Label(self.view_toggle_frame, text="View:", font=("Segoe UI", 8),
                bg=bg_color, fg=colors['text_primary'])
        self.view_label.pack(side=tk.LEFT, padx=(8, 3))

        self.view_mode_var = tk.StringVar(value="flat")

        # Radio Group for view mode toggle
        self.view_mode_radio_group = LabeledRadioGroup(
            self.view_toggle_frame,
            variable=self.view_mode_var,
            options=[
                ("flat", "Flat"),
                ("grouped", "Grouped"),
            ],
            command=self._on_view_mode_changed,
            orientation="horizontal",
            font=("Segoe UI", 8),
            padding=8
        )
        self.view_mode_radio_group.pack(side=tk.LEFT)

        # Set bottom toolbar to disabled initially (until parameter is created/loaded)
        self._set_bottom_toolbar_enabled(False)

    def _on_frame_configure(self, event):
        """Update canvas scroll region"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event):
        """Update canvas window width and reposition empty state text"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        # Center the empty state text
        if hasattr(self, '_empty_text_id'):
            self.canvas.coords(self._empty_text_id, event.width / 2, event.height / 2)

    def _on_scroll(self, *args):
        """Handle scrollbar scroll events"""
        self.canvas.yview(*args)
        # Trigger virtual scroll update if in virtual mode
        if self._virtual_mode:
            self._schedule_virtual_update()

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling - only if content is scrollable"""
        # Check if content is larger than visible area
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

    def _update_empty_state(self):
        """Show or hide the empty state text based on whether there are fields"""
        has_fields = bool(self.field_items)

        if hasattr(self, '_empty_text_id'):
            if has_fields:
                # Hide empty state when there are fields
                self.canvas.itemconfigure(self._empty_text_id, state='hidden')
            else:
                # Show empty state when no fields
                self.canvas.itemconfigure(self._empty_text_id, state='normal')

        # Enable/disable bottom toolbar based on field presence
        self._set_bottom_toolbar_enabled(has_fields)

    # =========================================================================
    # LOADING OVERLAY METHODS
    # =========================================================================

    def show_loading_overlay(self, message: str = "Loading..."):
        """Show a loading overlay on the canvas"""
        colors = self._theme_manager.colors
        overlay_bg = colors['card_surface']

        self._loading_overlay = tk.Frame(self.canvas, bg=overlay_bg)
        self._loading_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)

        loading_label = tk.Label(
            self._loading_overlay,
            text=message,
            font=("Segoe UI", 11),
            fg=colors['text_secondary'],
            bg=overlay_bg
        )
        loading_label.place(relx=0.5, rely=0.4, anchor="center")

        self._loading_progress = tk.Label(
            self._loading_overlay,
            text="Please wait...",
            font=("Segoe UI", 9, "italic"),
            fg=colors['text_muted'],
            bg=overlay_bg
        )
        self._loading_progress.place(relx=0.5, rely=0.5, anchor="center")

    def hide_loading_overlay(self):
        """Hide the loading overlay"""
        if hasattr(self, '_loading_overlay') and self._loading_overlay:
            self._loading_overlay.destroy()
            self._loading_overlay = None

    def update_loading_progress(self, text: str):
        """Update the loading progress text"""
        if hasattr(self, '_loading_progress') and self._loading_progress:
            self._loading_progress.config(text=text)
            self.update_idletasks()

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
        if self._category_filters is None:
            self._filtered_items = list(self.field_items)
        else:
            self._filtered_items = [fi for fi in self.field_items if self._field_matches_filter(fi)]

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

        # Capture current filtered items to avoid issues if list changes
        current_items = list(self._filtered_items)
        total_items = len(current_items)

        # FIRST: Show/create widgets that scrolled into view (place before hiding)
        # This prevents blank areas during scroll
        newly_created = []
        for i in range(new_start, min(new_end, total_items)):
            field_item = current_items[i]
            field_id = id(field_item)

            # Create widget if it doesn't exist
            if field_id not in self.field_widgets:
                # Find original index in field_items for order number
                orig_idx = self.field_items.index(field_item) if field_item in self.field_items else i
                self._create_field_widget_fast(field_item, orig_idx)
                newly_created.append(field_id)

            # Position widget using place() for absolute positioning
            frame = self.field_widgets[field_id]["frame"]
            y_pos = i * self.ROW_HEIGHT
            frame.place(x=3, y=y_pos, relwidth=1.0, width=-6, height=self.ROW_HEIGHT - 2)

        # Force update to show new widgets before hiding old ones
        self.fields_container.update_idletasks()

        # THEN: Hide widgets that scrolled out of view
        for i in range(old_start, min(old_end, total_items)):
            if i < new_start or i >= new_end:
                field_item = current_items[i]
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    self.field_widgets[field_id]["frame"].place_forget()

        self._visible_range = new_range

        # Apply selection highlighting to newly created widgets
        if newly_created and hasattr(self, 'selected_items'):
            self._apply_selection_to_widgets(newly_created)

    def _setup_virtual_scrolling(self):
        """Setup virtual scrolling mode for large lists"""
        self._virtual_mode = True
        self._update_filtered_items_cache()

        # Set scroll region based on total item count
        total_height = len(self._filtered_items) * self.ROW_HEIGHT
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), total_height))

        # Configure fields container to fill the scroll region
        self.fields_container.configure(height=total_height)

        # Initial render
        self._visible_range = (0, 0)
        self._update_virtual_view()

    def _disable_virtual_scrolling(self):
        """Disable virtual scrolling and use normal pack layout"""
        self._virtual_mode = False
        self._visible_range = (0, 0)
        self._filtered_items = []

        # Remove place() positioning and use pack() instead
        for field_item in self.field_items:
            field_id = id(field_item)
            if field_id in self.field_widgets:
                frame = self.field_widgets[field_id]["frame"]
                frame.place_forget()

    def _refresh_virtual_view(self):
        """Refresh the virtual view after data changes (add/remove/reorder)"""
        if not self._virtual_mode:
            return

        # Capture OLD visible items BEFORE updating the filtered items cache
        # This is critical: after filtering, _filtered_items will contain different items
        old_start, old_end = self._visible_range
        old_visible_items = list(self._filtered_items[old_start:old_end]) if self._filtered_items else []

        # Update filtered items cache (this changes _filtered_items)
        self._update_filtered_items_cache()

        # Update scroll region
        total_height = max(1, len(self._filtered_items) * self.ROW_HEIGHT)
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), total_height))
        self.fields_container.configure(height=total_height)

        # Hide all widgets that were previously visible (using saved references)
        for field_item in old_visible_items:
            field_id = id(field_item)
            if field_id in self.field_widgets:
                self.field_widgets[field_id]["frame"].place_forget()

        # Reset visible range and re-render
        self._visible_range = (0, 0)
        self._update_virtual_view()

    # =========================================================================
    # MOVE UP/DOWN METHODS
    # =========================================================================

    def _move_selected_up(self):
        """Move selected items up one position (or move selected group up)"""
        # In grouped view, check if a group is selected - move group instead
        if self._view_mode == "grouped" and self._selected_group:
            self._move_selected_group_up()
            return

        self._init_selection()
        if not self.selected_items:
            return

        # In grouped view, check alignment first
        if self._view_mode == "grouped":
            is_aligned, _, _ = self._check_category_alignment()
            if not is_aligned:
                return

        # Get indices of selected items (sorted)
        selected_indices = sorted([
            idx for idx, fi in enumerate(self.field_items)
            if id(fi) in self.selected_items
        ])

        if not selected_indices or selected_indices[0] == 0:
            return  # Already at top

        # Move each selected item up
        for idx in selected_indices:
            if idx > 0:
                # Swap with item above
                self.field_items[idx], self.field_items[idx - 1] = \
                    self.field_items[idx - 1], self.field_items[idx]

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        self._update_order_numbers()
        if self._view_mode == "grouped":
            self._render_grouped_view()
        else:
            self._repack_all_widgets()
        self._update_selection_visuals()
        self.main_tab.update_preview()

    def _move_selected_down(self):
        """Move selected items down one position (or move selected group down)"""
        # In grouped view, check if a group is selected - move group instead
        if self._view_mode == "grouped" and self._selected_group:
            self._move_selected_group_down()
            return

        self._init_selection()
        if not self.selected_items:
            return

        # In grouped view, check alignment first
        if self._view_mode == "grouped":
            is_aligned, _, _ = self._check_category_alignment()
            if not is_aligned:
                return

        # Get indices of selected items (sorted in reverse for moving down)
        selected_indices = sorted([
            idx for idx, fi in enumerate(self.field_items)
            if id(fi) in self.selected_items
        ], reverse=True)

        if not selected_indices or selected_indices[0] == len(self.field_items) - 1:
            return  # Already at bottom

        # Move each selected item down (process in reverse order)
        for idx in selected_indices:
            if idx < len(self.field_items) - 1:
                # Swap with item below
                self.field_items[idx], self.field_items[idx + 1] = \
                    self.field_items[idx + 1], self.field_items[idx]

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        self._update_order_numbers()
        if self._view_mode == "grouped":
            self._render_grouped_view()
        else:
            self._repack_all_widgets()
        self._update_selection_visuals()
        self.main_tab.update_preview()

    def _on_delete_selected(self):
        """Delete currently selected field(s) from parameter list"""
        self._init_selection()
        if not self.selected_items:
            return

        # Get selected items
        items_to_delete = [
            fi for fi in self.field_items
            if id(fi) in self.selected_items
        ]

        if not items_to_delete:
            return

        # Confirmation dialog
        count = len(items_to_delete)
        message = f"Delete {count} field{'s' if count > 1 else ''}?"
        if not ThemedMessageBox.askyesno(
            self.winfo_toplevel(),
            "Confirm Delete",
            message
        ):
            return

        # Remove from field_items list
        for item in items_to_delete:
            if item in self.field_items:
                self.field_items.remove(item)

        # Clear selection
        self.selected_items.clear()

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        # Refresh display
        self._update_order_numbers()
        if self._view_mode == "grouped":
            self._render_grouped_view()
        else:
            self._refresh_virtual_view()

        # Update empty state
        self._update_empty_state()

        # Trigger TMDL update immediately
        self.main_tab.update_preview()

    # =========================================================================
    # DROP INDICATOR METHODS
    # =========================================================================
    
    def show_drop_indicator(self, mouse_y: int):
        """
        Show a colored line indicating where the drop will occur.
        mouse_y is relative to the screen, we convert to widget coordinates.
        """
        # Convert screen Y to canvas Y
        canvas_y = mouse_y - self.canvas.winfo_rooty()

        # Account for scroll position
        scroll_y = self.canvas.canvasy(canvas_y)

        # In grouped view, use a different calculation and indicator
        if self._view_mode == "grouped":
            self._show_grouped_external_drop_indicator(scroll_y)
            return

        # Calculate drop position based on Y coordinate
        new_position = self._calculate_drop_position(scroll_y)

        # Only update if position changed (prevents flicker)
        if new_position == self._drop_position and self._drop_indicator is not None:
            return

        self._drop_position = new_position

        # Create indicator frame if it doesn't exist
        colors = self._theme_manager.colors
        if self._drop_indicator is None:
            self._drop_indicator = tk.Frame(
                self.fields_container,
                height=4,
                bg=colors['button_primary']  # Teal in light mode, blue in dark mode
            )

        # Position the indicator using place() for efficiency
        self._show_indicator_at_position(self._drop_position)
    
    def _show_indicator_at_position(self, position: int):
        """Show the indicator frame at the specified position using place() - OPTIMIZED

        Position is an index in the full field_items list. When filter is active,
        the indicator is placed relative to visible widgets.
        """
        if self._drop_indicator is None:
            return

        # In virtual scroll mode, calculate position based on filtered index and ROW_HEIGHT
        if self._virtual_mode:
            # Find the filtered index for this position
            filtered_idx = 0
            for idx, field_item in enumerate(self.field_items):
                if self._field_matches_filter(field_item):
                    if idx >= position:
                        break
                    filtered_idx += 1

            # Calculate Y position directly from index
            y_pos = filtered_idx * self.ROW_HEIGHT
            self._drop_indicator.place(x=3, y=y_pos, relwidth=1.0, width=-6, height=4)
            self._drop_indicator.lift()  # Ensure indicator is visible above other widgets
            return

        # Non-virtual mode: Build list of visible items with their indices
        visible_items = []
        for idx, field_item in enumerate(self.field_items):
            if self._field_matches_filter(field_item):
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    frame = self.field_widgets[field_id]["frame"]
                    if frame.winfo_ismapped():
                        visible_items.append((idx, frame))

        # Calculate Y position for the indicator
        y_pos = 0

        if not visible_items:
            # No visible items - show at top
            y_pos = 0
        else:
            # Find the visual position for the given list index
            # The indicator should appear between visible items
            last_visible_idx = visible_items[-1][0]

            if position > last_visible_idx:
                # Position is after all visible items - place after last visible
                last_frame = visible_items[-1][1]
                y_pos = last_frame.winfo_y() + last_frame.winfo_height() + 2
            else:
                # Find position before a visible item
                for i, (idx, frame) in enumerate(visible_items):
                    if idx >= position:
                        # Place indicator before this visible widget
                        y_pos = frame.winfo_y()
                        break

        # Use place() instead of pack() - much faster, no repacking needed
        self._drop_indicator.place(x=3, y=y_pos, relwidth=1.0, width=-6, height=4)
        self._drop_indicator.lift()  # Ensure indicator is visible
    
    def hide_drop_indicator(self):
        """Hide the drop indicator"""
        if self._drop_indicator:
            self._drop_indicator.place_forget()
        self._drop_position = None

    def _show_grouped_external_drop_indicator(self, y: float):
        """Show drop indicator for external drops in grouped view.

        In grouped view, we show the indicator at field positions within groups,
        finding which group and field position the mouse is over.
        """
        self.fields_container.update_idletasks()

        # Create indicator if needed
        colors = self._theme_manager.colors
        if self._drop_indicator is None:
            self._drop_indicator = tk.Frame(
                self.fields_container,
                height=4,
                bg=colors['button_primary']  # Teal in light mode, blue in dark mode
            )

        # Find first visible field for cases where y is above all content
        first_visible_field = None
        first_visible_y = float('inf')
        for field_item in self.field_items:
            if self._field_matches_filter(field_item):
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    frame = self.field_widgets[field_id]["frame"]
                    if frame.winfo_ismapped():
                        frame_y = frame.winfo_y()
                        if frame_y < first_visible_y:
                            first_visible_y = frame_y
                            first_visible_field = field_item
                        break  # First matching field is found

        # Find which field the mouse is over by checking all visible field widgets
        y_pos = 0
        target_position = len(self.field_items)  # Default to end

        # If y is above the first visible field, clamp to first position
        if first_visible_field and y < first_visible_y:
            try:
                target_position = self.field_items.index(first_visible_field)
            except ValueError:
                target_position = 0
            y_pos = first_visible_y
            self._drop_position = target_position
            self._drop_indicator.place(x=3, y=y_pos, relwidth=1.0, width=-6, height=4)
            self._drop_indicator.lift()
            return

        # Iterate through groups in order
        for sort_order, group_label in self._group_order:
            if group_label not in self._group_widgets:
                continue

            group_data = self._group_widgets[group_label]
            group_container = group_data.get("group_container")

            if not group_container or not group_container.winfo_ismapped():
                continue

            # Check if mouse is within this group container's header area
            header_frame = group_data.get("header_frame")
            if header_frame and header_frame.winfo_ismapped():
                header_y = header_frame.winfo_y()
                header_height = header_frame.winfo_height()
                header_center = header_y + header_height / 2

                if y < header_center:
                    # Place indicator before this group header
                    y_pos = header_y
                    # Find the first field in this group
                    group_fields = group_data.get("group_fields", [])
                    if group_fields:
                        try:
                            target_position = self.field_items.index(group_fields[0])
                        except ValueError:
                            pass
                    break

            # If group is expanded, check field positions within it
            if group_label not in self._collapsed_groups:
                group_fields = group_data.get("group_fields", [])
                for i, field_item in enumerate(group_fields):
                    field_id = id(field_item)
                    if field_id in self.field_widgets:
                        frame = self.field_widgets[field_id]["frame"]
                        if frame.winfo_ismapped():
                            frame_y = frame.winfo_y()
                            frame_height = frame.winfo_height()
                            frame_center = frame_y + frame_height / 2

                            if y < frame_center:
                                # Place indicator before this field
                                y_pos = frame_y
                                try:
                                    target_position = self.field_items.index(field_item)
                                except ValueError:
                                    pass
                                self._drop_position = target_position
                                self._drop_indicator.place(x=3, y=y_pos, relwidth=1.0, width=-6, height=4)
                                self._drop_indicator.lift()
                                return

                            # Update y_pos to be after this field (for "end of group" position)
                            y_pos = frame_y + frame_height + 2
                            try:
                                target_position = self.field_items.index(field_item) + 1
                            except ValueError:
                                pass

        # If we get here, place indicator at end of all content
        # Find actual bottom of all content
        max_y = 0
        for child in self.fields_container.winfo_children():
            if child.winfo_ismapped() and child != self._drop_indicator:
                child_bottom = child.winfo_y() + child.winfo_height()
                max_y = max(max_y, child_bottom)

        if max_y > 0:
            y_pos = max_y + 2

        self._drop_position = target_position
        self._drop_indicator.place(x=3, y=y_pos, relwidth=1.0, width=-6, height=4)
        self._drop_indicator.lift()

    def _calculate_drop_position(self, y: float) -> int:
        """
        Calculate the drop position (index in full list) based on Y coordinate.
        Returns the index where the new item should be inserted.

        When a filter is active, only visible items are considered for visual position,
        but the returned index is in terms of the full field_items list.
        """
        if not self.field_items:
            return 0

        # Build list of visible items with their indices in the full list
        visible_items = []
        for idx, field_item in enumerate(self.field_items):
            if self._field_matches_filter(field_item):
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    frame = self.field_widgets[field_id]["frame"]
                    if frame.winfo_ismapped():  # Only consider visible widgets
                        visible_items.append((idx, field_item, frame))

        if not visible_items:
            return 0

        # Check if mouse is above the first visible item
        first_idx, first_item, first_frame = visible_items[0]
        if y < first_frame.winfo_y():
            return first_idx

        # Find drop position based on visible widget positions
        for idx, field_item, frame in visible_items:
            frame_y = frame.winfo_y()
            frame_height = frame.winfo_height()
            frame_center = frame_y + frame_height / 2

            # If mouse is above the center of this widget, insert before it
            if y < frame_center:
                return idx

        # If we're past all visible widgets, insert after the last visible item
        last_visible_idx = visible_items[-1][0]
        return last_visible_idx + 1
    
    def get_drop_position(self) -> int:
        """Get the current drop position index"""
        return self._drop_position if self._drop_position is not None else len(self.field_items)
    
    def add_field(self, field_item: FieldItem, position: int = None):
        """
        Add a field to the builder at the specified position.
        If position is None, adds at the end.
        """
        # Hide empty state text
        self._update_empty_state()

        # Determine insertion position
        if position is None or position >= len(self.field_items):
            # Add at end
            self.field_items.append(field_item)
            insert_position = len(self.field_items) - 1
        else:
            # Insert at specific position
            self.field_items.insert(position, field_item)
            insert_position = position

        # In grouped view, assign categories from neighboring fields if the new field has none
        if self._view_mode == "grouped" and self._category_levels:
            self._assign_category_from_neighbors(field_item, insert_position)

        # Create the widget
        self._create_field_widget(field_item, insert_position)

        # Update order numbers for all fields
        self._update_order_numbers()

        # Repack all widgets in order - use grouped view render if in grouped mode
        if self._view_mode == "grouped":
            self._render_grouped_view()
        else:
            self._repack_all_widgets()

    def _assign_category_from_neighbors(self, field_item: FieldItem, position: int):
        """Assign category to a new field based on neighboring fields in grouped view.

        When adding a field in grouped view, we want it to inherit the category
        of its neighbors so it appears in the correct group.
        """
        if not self._category_levels or self._group_by_category_idx >= len(self._category_levels):
            return

        # If field already has categories for all levels, don't override
        if field_item.categories and len(field_item.categories) > self._group_by_category_idx:
            # Already has a category at this level
            return

        cat_idx = self._group_by_category_idx
        neighbor_category = None

        # Try to get category from field before this position
        if position > 0 and position <= len(self.field_items):
            # Note: field_item is already inserted at position, so check position-1
            neighbor = self.field_items[position - 1] if position > 0 else None
            if neighbor and neighbor != field_item:
                if neighbor.categories and len(neighbor.categories) > cat_idx:
                    neighbor_category = neighbor.categories[cat_idx]

        # If no category from before, try field after
        if not neighbor_category and position + 1 < len(self.field_items):
            neighbor = self.field_items[position + 1]
            if neighbor and neighbor != field_item:
                if neighbor.categories and len(neighbor.categories) > cat_idx:
                    neighbor_category = neighbor.categories[cat_idx]

        # Assign the category if found
        if neighbor_category:
            # Ensure categories list has enough slots
            while len(field_item.categories) < cat_idx:
                field_item.categories.append((0.0, ""))
            if len(field_item.categories) == cat_idx:
                field_item.categories.append(neighbor_category)
            else:
                field_item.categories[cat_idx] = neighbor_category

    def _create_field_widget(self, field_item: FieldItem, position: int):
        """Create the UI widget for a field item - compact single-line row"""
        field_id = id(field_item)
        colors = self._theme_manager.colors
        # Alternating row colors for zebra striping
        row_bg = colors['card_surface'] if position % 2 == 0 else colors['surface']

        # Create compact field row frame - flat design with invisible border
        # Note: highlightthickness must be set at creation to prevent height changes on selection
        field_frame = tk.Frame(
            self.fields_container,
            bg=row_bg,
            highlightthickness=2,
            highlightbackground=row_bg  # Invisible border matching background
        )
        # Don't pack yet - will be packed in _repack_all_widgets

        # Helper to bind all row events to a widget
        def bind_row_events(widget):
            widget.bind("<Button-1>", lambda e, fi=field_item: self._on_field_click(e, fi))
            widget.bind("<B1-Motion>", lambda e, fi=field_item: self._on_drag_motion(e, fi))
            widget.bind("<ButtonRelease-1>", self._end_internal_drag)
            widget.bind("<Button-3>", lambda e, fi=field_item: self._on_field_right_click(e, fi))

        # Bind to frame
        bind_row_events(field_frame)

        # Order number - use tk.Label for consistent styling
        order_label = tk.Label(
            field_frame,
            text=f"{position + 1}.",
            font=("Segoe UI", 9, "bold"),
            width=4,
            anchor="e",
            bg=row_bg,
            fg=colors['text_primary']
        )
        order_label.pack(side=tk.LEFT, padx=(4, 0), pady=2)
        bind_row_events(order_label)

        # Category position (within-group position) - shown in muted, smaller
        cat_pos_label = tk.Label(field_frame, text="", font=("Segoe UI", 7),
                                 fg=colors['text_muted'], width=4, anchor="w", bg=row_bg)
        cat_pos_label.pack(side=tk.LEFT, padx=(0, 2), pady=2)
        bind_row_events(cat_pos_label)

        # Drag handle - use tk.Label for consistent styling
        drag_handle = tk.Label(field_frame, text="☰", font=("Segoe UI", 10), cursor="hand2", bg=row_bg, fg=colors['text_secondary'])
        drag_handle.pack(side=tk.LEFT, padx=(0, 4), pady=2)
        bind_row_events(drag_handle)

        # Display name (editable label) - expands to fill available space
        display_var = tk.StringVar(value=field_item.display_name)
        display_label = tk.Label(
            field_frame,
            textvariable=display_var,
            font=("Segoe UI", 9, "bold"),
            fg=colors['title_color'],
            anchor="w",
            cursor="hand2",
            bg=row_bg
        )
        display_label.pack(side=tk.LEFT, padx=(0, 4), pady=2, fill=tk.X, expand=True)
        bind_row_events(display_label)
        # Double-click edits name
        display_label.bind("<Double-Button-1>",
                          lambda e, fi=field_item, dv=display_var: self._edit_display_name(fi, dv))

        # Field reference (read-only) - fixed width on right side
        ref_label = tk.Label(
            field_frame,
            text=field_item.field_reference,
            font=("Consolas", 8),
            fg=colors['text_muted'],
            anchor="w",
            width=35,
            bg=row_bg
        )
        ref_label.pack(side=tk.RIGHT, padx=(0, 2), pady=2)
        bind_row_events(ref_label)

        # Store widget references
        self.field_widgets[field_id] = {
            "frame": field_frame,
            "order_label": order_label,
            "cat_pos_label": cat_pos_label,
            "drag_handle": drag_handle,
            "display_var": display_var,
            "display_label": display_label,
            "ref_label": ref_label,
            "field_item": field_item
        }

        # Bind mousewheel to all child widgets so scrolling works anywhere
        self._bind_mousewheel_recursive(field_frame)
    
    def remove_field(self, field_item: FieldItem):
        """Remove a field from the builder"""
        field_id = id(field_item)
        if field_id in self.field_widgets:
            # Destroy widget
            self.field_widgets[field_id]["frame"].destroy()
            del self.field_widgets[field_id]

            # Remove from list
            if field_item in self.field_items:
                self.field_items.remove(field_item)

            # Remove from selection if selected
            if hasattr(self, 'selected_items') and field_id in self.selected_items:
                self.selected_items.discard(field_id)

            # Update order numbers
            self._update_order_numbers()

            # Update empty state visibility
            self._update_empty_state()
    
    def refresh_all_fields(self):
        """Refresh display of all fields (e.g., after bulk name change)"""
        for field_id, widgets in self.field_widgets.items():
            # Support both StringVar (legacy) and direct label (optimized)
            if "display_var" in widgets:
                widgets["display_var"].set(widgets["field_item"].display_name)
            else:
                widgets["display_label"].config(text=widgets["field_item"].display_name)
    
    def update_category_options(self, category_levels: List[CategoryLevel]):
        """Update category options for the hierarchical filter dropdown"""
        # Check if category levels structure changed (only reset filter if it did)
        # Compare by column names - if structure is same, preserve filter
        old_columns = [cl.column_name for cl in self._category_levels] if self._category_levels else []
        new_columns = [cl.column_name for cl in category_levels] if category_levels else []
        categories_changed = old_columns != new_columns

        self._category_levels = category_levels

        # Update the hierarchical filter dropdown with new category levels
        self._filter_dropdown.set_category_levels(category_levels)

        # Filter frame is always visible - dropdown contents update based on categories

        # Update group-by options (view_mode_frame is always visible)
        self._update_group_by_options()

        # Only reset filter if category structure actually changed
        if categories_changed:
            self._category_filters = None
            self._update_filter_count()

    def _on_filter_changed(self):
        """Handle filter dropdown selection change - callback from FieldParameterFilterDropdown"""
        self._category_filters = self._filter_dropdown.get_filter_state()
        self._apply_filter()

    def _apply_filter(self):
        """Apply the current filter to show/hide field widgets - OPTIMIZED"""
        # Freeze UI updates during bulk operation
        self.fields_container.update_idletasks()
        self.canvas.config(cursor="watch")

        if self._view_mode == "grouped":
            # Grouped mode - re-render the grouped view (will respect filter)
            self._render_grouped_view()
        elif self._virtual_mode:
            # Virtual scrolling mode - refresh the virtual view
            self._refresh_virtual_view()
        else:
            # Standard flat mode - unpack/repack widgets
            # First unpack all widgets
            for field_item in self.field_items:
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    self.field_widgets[field_id]["frame"].pack_forget()

            # Then repack only matching items in correct order
            for field_item in self.field_items:
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    if self._field_matches_filter(field_item):
                        self.field_widgets[field_id]["frame"].pack(fill=tk.X, pady=1, padx=3)

        self._update_filter_count()
        self._update_order_numbers()

        # Reset scroll to top when filter changes
        self.canvas.yview_moveto(0)
        self.canvas.config(cursor="")

    def _field_matches_filter(self, field_item: FieldItem) -> bool:
        """Check if a field matches the current filter (multi-column, multi-select)"""
        if self._category_filters is None:
            return True  # Show all

        # Get all category labels for this field
        field_labels = []
        if field_item.categories:
            for _, label in field_item.categories:
                field_labels.append(label or "")
        else:
            field_labels = [""]  # Uncategorized

        # Check if uncategorized filter is active
        if self._category_filters.get("__uncategorized__"):
            # Field is uncategorized if all its category labels are empty
            if all(label == "" for label in field_labels):
                return True

        # Check if field matches ANY selected label in ANY column
        for col_idx, selected_labels in self._category_filters.items():
            if col_idx == "__uncategorized__":
                continue
            if isinstance(col_idx, int) and col_idx < len(field_labels):
                if field_labels[col_idx] in selected_labels:
                    return True

        return False

    def _update_filter_count(self):
        """Update the filter count label"""
        if self._category_filters is None:
            self.filter_count_label.config(text=f"Showing all {len(self.field_items)} fields")
        else:
            visible_count = sum(1 for fi in self.field_items if self._field_matches_filter(fi))
            self.filter_count_label.config(text=f"Showing {visible_count} of {len(self.field_items)} fields")

    # =========================================================================
    # GROUPED VIEW METHODS
    # =========================================================================

    def _on_view_mode_changed(self):
        """Handle view mode toggle between flat and grouped"""
        new_mode = self.view_mode_var.get()
        if new_mode == self._view_mode:
            return

        self._view_mode = new_mode
        self.logger.info(f"View mode changed to: {new_mode}")

        # Clear group selection when switching views
        self._selected_group = None

        # Hide ALL field widgets immediately to prevent visual glitch
        for field_id, widgets in self.field_widgets.items():
            widgets["frame"].pack_forget()
            widgets["frame"].place_forget()  # Also clear place() positioning

        # Force UI update before rendering new view
        self.fields_container.update_idletasks()

        # Lock canvas_window width to prevent width changes from content differences
        canvas_width = self.canvas.winfo_width()
        if canvas_width > 10:
            self.canvas.itemconfig(self.canvas_window, width=canvas_width)

        # Reset scroll position when switching views
        self.canvas.yview_moveto(0)

        if new_mode == "grouped":
            # IMPORTANT: Disable virtual scrolling in grouped view
            # Grouped view uses pack() not place(), so virtual scroll updates would conflict
            self._virtual_mode = False
            self._visible_range = (0, 0)
            # Show grouped controls using place() to avoid affecting view_mode_frame width
            # Position to the left of view_toggle_frame
            self.view_toggle_frame.update_idletasks()
            toggle_width = self.view_toggle_frame.winfo_reqwidth()
            self._grouped_controls_frame.update_idletasks()
            controls_width = self._grouped_controls_frame.winfo_reqwidth()
            # Place at x offset from right edge, with some padding
            self._grouped_controls_frame.place(in_=self.view_mode_frame, relx=1.0, x=-(toggle_width + controls_width + 8), rely=0.5, anchor="w")
            # Show header dropdown for category selection using place() to avoid layout changes
            if hasattr(self, '_header_group_by_frame') and hasattr(self, '_toolbar1'):
                # Position after filter_frame without affecting pack layout
                self.filter_frame.update_idletasks()
                filter_width = self.filter_frame.winfo_reqwidth()
                self._header_group_by_frame.place(in_=self._toolbar1, x=filter_width + 8, rely=0.5, anchor="w")

            # Ensure all widgets have cat_pos_label - destroy and recreate any that don't
            widgets_to_recreate = []
            for field_id, widgets in list(self.field_widgets.items()):
                if "cat_pos_label" not in widgets:
                    widgets_to_recreate.append(field_id)
                    widgets["frame"].destroy()
                    del self.field_widgets[field_id]

            if widgets_to_recreate:
                self.logger.info(f"Recreating {len(widgets_to_recreate)} widgets missing cat_pos_label")
                for field_item in self.field_items:
                    if id(field_item) in widgets_to_recreate:
                        try:
                            orig_idx = self.field_items.index(field_item)
                        except ValueError:
                            orig_idx = 0
                        self._create_field_widget_fast(field_item, orig_idx)

            # Switch to grouped view (all collapsed initially)
            self._render_grouped_view(initial_render=True)
            # Update align button appearance
            self._update_align_button_style()
        else:
            # Hide grouped controls using place_forget (matches place() used when showing)
            self._grouped_controls_frame.place_forget()
            # Hide header dropdown
            if hasattr(self, '_header_group_by_frame'):
                self._header_group_by_frame.place_forget()
            # Switch back to flat view
            self._render_flat_view()

    def _on_align_to_category(self):
        """Handle click on Align button - sorts fields to match category order"""
        is_aligned, message, details = self._check_category_alignment(detailed=True)
        if is_aligned:
            # Get current category column name
            cat_name = ""
            if self._category_levels and self._group_by_category_idx < len(self._category_levels):
                cat_name = self._category_levels[self._group_by_category_idx].name
            ThemedMessageBox.showinfo(
                self.container.winfo_toplevel(),
                "Already Aligned",
                f"Field order is already aligned!\n\n"
                f"Fields are correctly grouped by '{cat_name}' - all fields in each "
                f"category appear together in a contiguous block.\n\n"
                f"No reordering is needed."
            )
            return

        # Build simple message showing only the misaligned fields
        cat_name = ""
        if self._category_levels and self._group_by_category_idx < len(self._category_levels):
            cat_name = self._category_levels[self._group_by_category_idx].name

        misaligned_fields = details.get('misaligned_fields', [])

        detail_msg = f"The following field(s) are out of place and breaking '{cat_name}' grouping:\n\n"

        # Show only the misaligned fields - these are the ones that need to move
        if misaligned_fields:
            for field_name, cat in misaligned_fields[:10]:
                detail_msg += f"  • '{field_name}' (in {cat})\n"
            if len(misaligned_fields) > 10:
                detail_msg += f"  ...and {len(misaligned_fields) - 10} more\n"

        detail_msg += f"\nAlign will move these {len(misaligned_fields)} field(s) to group with their category.\n\n"
        detail_msg += "Continue with alignment?"

        result = ThemedMessageBox.askyesno(self.container.winfo_toplevel(), "Align to Category Order", detail_msg)

        if result:
            self._align_to_category_order()
            ThemedMessageBox.showinfo(
                self.container.winfo_toplevel(),
                "Aligned",
                "Field order has been aligned to category grouping.\n\n"
                "All fields are now grouped by their category value."
            )

    def _on_group_by_changed(self, event=None):
        """Handle change in which category column to group by (bottom toolbar)"""
        combo_selection = self.group_by_combo.current()
        if combo_selection < 0:
            return
        # Map combobox index to actual category level index
        if not hasattr(self, '_combo_to_level_idx') or combo_selection >= len(self._combo_to_level_idx):
            return
        actual_level_idx = self._combo_to_level_idx[combo_selection]
        if actual_level_idx != self._group_by_category_idx:
            self._group_by_category_idx = actual_level_idx
            self._collapsed_groups.clear()  # Reset collapsed state
            # Sync header dropdown
            if hasattr(self, '_header_group_by_combo'):
                self._header_group_by_combo.current(combo_selection)
            # Render with initial_render=True to keep groups collapsed
            self._render_grouped_view(initial_render=True)

    def _on_header_group_by_changed(self, event=None):
        """Handle change in which category column to group by (header dropdown)"""
        combo_selection = self._header_group_by_combo.current()
        if combo_selection < 0:
            return
        # Map combobox index to actual category level index
        if not hasattr(self, '_combo_to_level_idx') or combo_selection >= len(self._combo_to_level_idx):
            return
        actual_level_idx = self._combo_to_level_idx[combo_selection]
        if actual_level_idx != self._group_by_category_idx:
            self._group_by_category_idx = actual_level_idx
            self._collapsed_groups.clear()  # Reset collapsed state
            # Sync bottom toolbar dropdown
            self.group_by_combo.current(combo_selection)
            # Render with initial_render=True to keep groups collapsed
            self._render_grouped_view(initial_render=True)

    def _update_align_button_style(self):
        """Update Align button style based on whether alignment is needed"""
        colors = self._theme_manager.colors
        bg_color = colors['background']
        is_aligned, _, _ = self._check_category_alignment()
        if is_aligned:
            # Normal style - no action needed
            self.align_btn.text = "\u21BB Align"
            self.align_btn.update_colors(
                bg=colors['button_secondary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                fg=colors['text_primary'],
                canvas_bg=bg_color
            )
        else:
            # Highlight style - alignment needed, reordering disabled
            self.align_btn.text = "\u26A0 Align"
            self.align_btn.update_colors(
                bg=colors['warning'],
                hover_bg=colors['warning_hover'] if 'warning_hover' in colors else colors['warning'],
                pressed_bg=colors['warning'],
                fg='#ffffff',  # White text for contrast
                canvas_bg=bg_color
            )

    def update_alignment_status(self):
        """Public method to update alignment button style when categories change.
        Called from main UI when field categories are modified."""
        if self._view_mode == "grouped":
            self._update_align_button_style()

    def _update_group_by_options(self):
        """Update the group-by dropdowns with available category columns"""
        if not self._category_levels:
            self.group_by_combo['values'] = []
            self.group_by_combo.set('')
            if hasattr(self, '_header_group_by_combo'):
                self._header_group_by_combo['values'] = []
                self._header_group_by_combo.set('')
            self._combo_to_level_idx = []  # Clear mapping
            return

        # Only include non-calculated categories with labels
        # Store mapping from combobox index to actual _category_levels index
        options = []
        self._combo_to_level_idx = []
        for level_idx, level in enumerate(self._category_levels):
            if not level.is_calculated and level.labels:
                options.append(level.name)
                self._combo_to_level_idx.append(level_idx)

        # Update both dropdowns
        self.group_by_combo['values'] = options
        if hasattr(self, '_header_group_by_combo'):
            self._header_group_by_combo['values'] = options

        if options:
            # Find the combobox index for current _group_by_category_idx
            combo_idx = 0
            if self._group_by_category_idx in self._combo_to_level_idx:
                combo_idx = self._combo_to_level_idx.index(self._group_by_category_idx)
            else:
                # Current selection not valid, reset to first
                self._group_by_category_idx = self._combo_to_level_idx[0] if self._combo_to_level_idx else 0

            self.group_by_combo.current(combo_idx)
            if hasattr(self, '_header_group_by_combo'):
                self._header_group_by_combo.current(combo_idx)

    def _render_flat_view(self):
        """Render the standard flat list view"""
        # Clear grouped view widgets
        self._clear_grouped_view()

        # Re-enable virtual scrolling if we have enough items
        if len(self.field_items) > self.VIRTUAL_SCROLL_THRESHOLD:
            self._setup_virtual_scrolling()
        else:
            # Make sure virtual mode is off for smaller lists
            self._virtual_mode = False

        # Repack all field widgets in order
        self._repack_all_widgets()

    def _render_grouped_view(self, initial_render: bool = False):
        """Render fields grouped by category with collapsible headers

        Args:
            initial_render: If True, collapse all groups initially
        """
        if not self._category_levels or self._group_by_category_idx >= len(self._category_levels):
            # No categories - fall back to flat view
            self._view_mode = "flat"
            self.view_mode_var.set("flat")
            self._render_flat_view()
            return

        # Clear existing layout
        self._clear_grouped_view()

        # Hide all field widgets first (both pack and place)
        for field_id, widgets in self.field_widgets.items():
            widgets["frame"].pack_forget()
            widgets["frame"].place_forget()

        # Force UI update to prevent visual glitch
        self.fields_container.update_idletasks()

        # Freeze container size during rendering to prevent width changes
        self.fields_container.pack_propagate(False)

        # Get category level for grouping
        cat_level = self._category_levels[self._group_by_category_idx]

        # Group fields by category
        fields_by_category = {}
        uncategorized = []

        for field_item in self.field_items:
            if not self._field_matches_filter(field_item):
                continue

            # Get category label at the group-by index
            cat_label = ""
            cat_sort = 999.0
            if field_item.categories and len(field_item.categories) > self._group_by_category_idx:
                cat_sort, cat_label = field_item.categories[self._group_by_category_idx]

            if cat_label:
                if cat_label not in fields_by_category:
                    fields_by_category[cat_label] = {"sort": cat_sort, "fields": []}
                fields_by_category[cat_label]["fields"].append(field_item)
            else:
                uncategorized.append(field_item)

        # Build group order from category level's label order (respects user-defined sort)
        self._group_order = []
        for label in cat_level.labels:
            if label in fields_by_category:
                sort_val = fields_by_category[label]["sort"]
                self._group_order.append((sort_val, label))

        # Sort by sort order (preserves category label order)
        self._group_order.sort(key=lambda x: x[0])

        # Add uncategorized at the end if any
        if uncategorized:
            self._group_order.append((999.0, "__uncategorized__"))
            fields_by_category["__uncategorized__"] = {"sort": 999.0, "fields": uncategorized}

        # On initial render, collapse all groups
        if initial_render:
            self._collapsed_groups = {label for _, label in self._group_order}

        # Store fields_by_category for later use (e.g., during expand/collapse)
        self._fields_by_category = fields_by_category

        # Check alignment once for all groups (avoid calling per-group)
        is_aligned, _, _ = self._check_category_alignment()

        # Render each group with its fields in a container frame
        for position, (sort_order, group_label) in enumerate(self._group_order, start=1):
            display_label = "(Uncategorized)" if group_label == "__uncategorized__" else group_label
            field_count = len(fields_by_category[group_label]["fields"])
            group_fields = fields_by_category[group_label]["fields"]
            self._create_group_with_fields(group_label, display_label, field_count, sort_order, group_fields, position, is_aligned)

        # Add a spacer at the bottom to ensure drop indicator is visible at end
        if hasattr(self, '_bottom_spacer') and self._bottom_spacer:
            self._bottom_spacer.destroy()
        # Use background color to match canvas (darkest for contrast with row colors)
        colors = self._theme_manager.colors
        self._bottom_spacer = tk.Frame(self.fields_container, height=10, bg=colors['background'])
        self._bottom_spacer.pack(fill=tk.X)

        # Update order numbers (including category positions) now that we're in grouped view
        self._update_order_numbers()

        # Schedule UI refresh to avoid blocking - update_idletasks is non-blocking
        self.fields_container.update_idletasks()
        self.canvas.update_idletasks()

        # Update align button style based on alignment status
        self._update_align_button_style()

        # Re-enable pack propagation after rendering complete
        self.fields_container.pack_propagate(True)

    def _create_group_with_fields(self, group_label: str, display_label: str, field_count: int, sort_order: float, group_fields: list, position: int = 0, is_aligned: bool = True):
        """Create a group container with header and fields inside"""
        colors = self._theme_manager.colors
        is_collapsed = group_label in self._collapsed_groups
        expand_icon = "►" if is_collapsed else "▼"

        # Theme-aware colors for group headers
        # Use background for container (darkest) so header row stands out
        group_bg = colors['background']
        header_bg = colors['surface']

        # Group container frame - contains header + fields content
        group_container = tk.Frame(self.fields_container, bg=group_bg)
        group_container.pack(fill=tk.X, pady=(2, 0))

        # Group header frame - flat design with subtle border
        header_frame = tk.Frame(
            group_container,
            bg=header_bg,
            highlightbackground=colors['border'],
            highlightcolor=colors['border'],
            highlightthickness=2  # Keep constant to prevent height changes on selection
        )
        header_frame.pack(fill=tk.X, pady=(2, 1), padx=3)

        # Drag handle for reordering groups - grayed out if not aligned
        if is_aligned:
            drag_handle = tk.Label(
                header_frame, text="☰", font=("Segoe UI", 10),
                cursor="hand2", bg=header_bg, fg=colors['text_primary']
            )
        else:
            drag_handle = tk.Label(
                header_frame, text="☰", font=("Segoe UI", 10),
                cursor="arrow", bg=header_bg, fg=colors['text_muted']  # Grayed out
            )
        drag_handle.pack(side=tk.LEFT, padx=(4, 4), pady=2)

        # Bind drag events to handle ONLY
        drag_handle.bind("<Button-1>", lambda e, gl=group_label: self._on_group_drag_start(e, gl))
        drag_handle.bind("<B1-Motion>", self._on_group_drag_motion)
        drag_handle.bind("<ButtonRelease-1>", self._on_group_drag_end)

        # Expand/collapse button - use Label for flat design (no 3D button appearance)
        expand_btn = tk.Label(
            header_frame, text=expand_icon, font=("Segoe UI", 10),
            bg=header_bg, fg=colors['text_primary'], width=2, cursor="hand2"
        )
        expand_btn.pack(side=tk.LEFT, padx=(0, 2), pady=2)
        expand_btn.bind("<Button-1>", lambda e, gl=group_label: self._toggle_group_collapse(gl))

        # Position number label (next to expand button)
        position_label = tk.Label(
            header_frame, text=str(position), font=("Segoe UI", 8),
            fg=colors['text_muted'], bg=header_bg
        )
        position_label.pack(side=tk.LEFT, padx=(0, 4), pady=2)

        # Group name - clicking also toggles
        name_label = tk.Label(
            header_frame, text=display_label, font=("Segoe UI", 9, "bold"),
            fg=colors['text_primary'], bg=header_bg, anchor="w", cursor="hand2"
        )
        name_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=2)
        name_label.bind("<Button-1>", lambda e, gl=group_label: self._toggle_group_collapse(gl))

        # Field count
        count_label = tk.Label(
            header_frame, text=f"({field_count})", font=("Segoe UI", 8),
            fg=colors['text_muted'], bg=header_bg
        )
        count_label.pack(side=tk.RIGHT, padx=(4, 8), pady=2)

        # Bind right-click context menu to header elements
        right_click_handler = lambda e, gl=group_label: self._on_group_header_right_click(e, gl)
        for widget in [header_frame, drag_handle, expand_btn, position_label, name_label, count_label]:
            widget.bind("<Button-3>", right_click_handler)

        # Bind left-click for group selection (allow selecting group for arrow key navigation)
        click_handler = lambda e, gl=group_label: self._on_group_header_click(e, gl)
        for widget in [header_frame, drag_handle, position_label, name_label, count_label]:
            widget.bind("<Button-1>", click_handler, add="+")

        # Store widget reference (fields_frame not used - we pack fields directly)
        self._group_widgets[group_label] = {
            "group_container": group_container,
            "header_frame": header_frame,
            "drag_handle": drag_handle,
            "expand_btn": expand_btn,
            "position_label": position_label,
            "name_label": name_label,
            "count_label": count_label,
            "sort_order": sort_order,
            "group_fields": group_fields
        }

        # Bind mousewheel to header widgets
        for widget in [group_container, header_frame, drag_handle, expand_btn, position_label, name_label, count_label]:
            widget.bind("<MouseWheel>", self._on_mousewheel)

        # Pack fields DIRECTLY into fields_container (not into a nested frame)
        # This avoids the pack(in_=) cross-parent clipping issues
        if not is_collapsed:
            for field_item in group_fields:
                field_id = id(field_item)
                # Create widget if it doesn't exist (may not exist if virtual scrolling was active)
                # OR recreate if it's missing cat_pos_label (backwards compatibility)
                needs_create = field_id not in self.field_widgets
                if not needs_create and "cat_pos_label" not in self.field_widgets[field_id]:
                    # Widget exists but is missing cat_pos_label - recreate it
                    self.field_widgets[field_id]["frame"].destroy()
                    del self.field_widgets[field_id]
                    needs_create = True

                if needs_create:
                    # Find position in field_items for order number
                    try:
                        orig_idx = self.field_items.index(field_item)
                    except ValueError:
                        orig_idx = 0
                    self._create_field_widget_fast(field_item, orig_idx)

                if field_id in self.field_widgets:
                    widget_frame = self.field_widgets[field_id]["frame"]
                    # Pack directly in fields_container (consistent padx with flat view)
                    widget_frame.pack(fill=tk.X, pady=1, padx=3)
                    # Add internal indent via order_label padding for grouped view
                    order_label = self.field_widgets[field_id]["order_label"]
                    order_label.pack_configure(padx=(21, 0))

    def _toggle_group_collapse(self, group_label: str):
        """Toggle collapse state of a group"""
        if group_label in self._collapsed_groups:
            self._collapsed_groups.discard(group_label)
        else:
            self._collapsed_groups.add(group_label)

        # Re-render grouped view
        self._render_grouped_view()

        # Explicitly update order numbers again after render to ensure cat_pos_labels are set
        self._update_order_numbers()
        self.fields_container.update_idletasks()

    def _clear_grouped_view(self):
        """Clear all group container widgets"""
        # First, unpack all field widgets (they are packed directly in fields_container)
        for field_id, widgets in self.field_widgets.items():
            widgets["frame"].pack_forget()

        # Unbind and destroy group containers (headers only - fields are children of fields_container)
        for group_label, widgets in self._group_widgets.items():
            if "group_container" in widgets:
                self._cleanup_widget_bindings(widgets["group_container"])
                widgets["group_container"].destroy()
        self._group_widgets.clear()

        # Hide group drop indicator (but don't destroy - reuse)
        if self._group_drop_indicator:
            self._group_drop_indicator.place_forget()

        # Destroy bottom spacer
        if hasattr(self, '_bottom_spacer') and self._bottom_spacer:
            self._bottom_spacer.destroy()
            self._bottom_spacer = None

    def _is_category_filter_active(self) -> bool:
        """Check if a category filter is currently active (not showing all)"""
        return self._category_filters is not None

    def _on_group_drag_start(self, event, group_label: str):
        """Start dragging a group header"""
        # Disable drag when filter is active
        if self._is_category_filter_active():
            self._group_drag_data = {"item": None, "dragging": False}
            return

        # Disable drag unless fields are aligned with categories
        is_aligned, _, _ = self._check_category_alignment()
        if not is_aligned:
            self._group_drag_data = {"item": None, "dragging": False}
            return

        self._group_drag_data = {
            "item": group_label,
            "start_y": event.y_root,
            "dragging": False
        }

    def _auto_scroll_if_near_edge(self, event):
        """Auto-scroll the canvas when dragging near the edge of the visible area.

        Uses variable speed - faster scrolling the closer to the edge.
        """
        canvas = self.canvas
        canvas_top = canvas.winfo_rooty()
        canvas_height = canvas.winfo_height()
        canvas_bottom = canvas_top + canvas_height

        # Define edge zones (pixels from edge)
        edge_zone = 60  # Larger zone for smoother experience
        mouse_y = event.y_root
        scrolled = False

        # Calculate distance into the edge zone
        if mouse_y < canvas_top + edge_zone:
            # Near top edge - scroll up (only if not already at top)
            scroll_top = canvas.yview()[0]
            if scroll_top > 0:
                distance_into_zone = canvas_top + edge_zone - mouse_y
                # Variable speed: 1-3 pixels based on how far into zone (slower overall)
                # Divide by 20 for much slower base speed
                scroll_amount = max(1, min(3, int(distance_into_zone / 20)))
                # Use yview_scroll with fractional "pages" for finer control
                canvas.yview_scroll(-scroll_amount, "units")
                scrolled = True
        elif mouse_y > canvas_bottom - edge_zone:
            # Near bottom edge - scroll down
            distance_into_zone = mouse_y - (canvas_bottom - edge_zone)
            scroll_amount = max(1, min(3, int(distance_into_zone / 20)))
            canvas.yview_scroll(scroll_amount, "units")
            scrolled = True

        # IMPORTANT: Trigger virtual view update if we scrolled in virtual mode
        if scrolled and self._virtual_mode:
            self._update_virtual_view()

    def _on_group_drag_motion(self, event):
        """Handle group drag motion"""
        if self._group_drag_data.get("item") is None:
            return

        # Check if we've moved enough to consider it a drag
        start_y = self._group_drag_data.get("start_y", 0)
        if abs(event.y_root - start_y) > 5:
            self._group_drag_data["dragging"] = True

        if not self._group_drag_data.get("dragging"):
            return

        # Auto-scroll if near edge of visible area
        self._auto_scroll_if_near_edge(event)

        # Show drop indicator
        canvas_y = self.canvas.canvasy(event.y_root - self.canvas.winfo_rooty())
        drop_pos = self._calculate_group_drop_position(canvas_y)

        if drop_pos != self._group_drag_data.get("drop_pos"):
            self._group_drag_data["drop_pos"] = drop_pos
            self._show_group_drop_indicator(drop_pos)

    def _on_group_drag_end(self, event):
        """End group drag and reorder if needed"""
        if not self._group_drag_data.get("dragging"):
            self._group_drag_data = {"item": None, "dragging": False}
            return

        group_label = self._group_drag_data.get("item")
        drop_pos = self._group_drag_data.get("drop_pos")

        # Hide drop indicator
        if self._group_drop_indicator:
            self._group_drop_indicator.place_forget()

        if group_label and drop_pos is not None:
            self._reorder_group(group_label, drop_pos)

        self._group_drag_data = {"item": None, "dragging": False}

    def _calculate_group_drop_position(self, y: float) -> int:
        """Calculate drop position index for group reordering"""
        if not self._group_order:
            return 0

        # Force geometry update to ensure winfo values are current
        self.fields_container.update_idletasks()

        # Find position based on group container positions
        # We need to find the FULL height of each group including expanded fields
        for idx, (sort_order, label) in enumerate(self._group_order):
            if label not in self._group_widgets:
                continue

            group_data = self._group_widgets[label]
            container = group_data.get("group_container") or group_data.get("header_frame")
            if not container or not container.winfo_ismapped():
                continue

            # Get the top of this group
            group_top = container.winfo_y()

            # Calculate the full bottom of this group (including expanded fields)
            group_bottom = group_top + container.winfo_height()

            # If group is expanded, find the bottom of its last field
            if label not in self._collapsed_groups:
                group_fields = group_data.get("group_fields", [])
                for field_item in group_fields:
                    field_id = id(field_item)
                    if field_id in self.field_widgets:
                        frame = self.field_widgets[field_id]["frame"]
                        if frame.winfo_ismapped():
                            field_bottom = frame.winfo_y() + frame.winfo_height()
                            group_bottom = max(group_bottom, field_bottom)

            # Calculate the center of the FULL group (header + expanded fields)
            group_center = (group_top + group_bottom) / 2

            if y < group_center:
                return idx

        # Below all groups - return position after last group
        return len(self._group_order)

    def _show_group_drop_indicator(self, position: int):
        """Show drop indicator at group position"""
        colors = self._theme_manager.colors
        if self._group_drop_indicator is None:
            self._group_drop_indicator = tk.Frame(
                self.fields_container, height=4, bg=colors['button_primary']
            )

        # Force geometry update to ensure winfo values are current
        self.fields_container.update_idletasks()

        # Calculate Y position based on group containers
        y_pos = 0
        if position < len(self._group_order):
            # Position before a specific group - show at top of that group
            label = self._group_order[position][1]
            if label in self._group_widgets:
                container = self._group_widgets[label].get("group_container") or self._group_widgets[label]["header_frame"]
                y_pos = container.winfo_y()
        elif self._group_order:
            # Position at the end - find the bottom of the last group
            last_label = self._group_order[-1][1]
            if last_label in self._group_widgets:
                group_data = self._group_widgets[last_label]
                container = group_data.get("group_container") or group_data.get("header_frame")

                # Start with the container bottom
                max_y = container.winfo_y() + container.winfo_height() if container and container.winfo_ismapped() else 0

                # If group is expanded, find the bottom of its last field
                if last_label not in self._collapsed_groups:
                    group_fields = group_data.get("group_fields", [])
                    for field_item in group_fields:
                        field_id = id(field_item)
                        if field_id in self.field_widgets:
                            frame = self.field_widgets[field_id]["frame"]
                            if frame.winfo_ismapped():
                                frame_bottom = frame.winfo_y() + frame.winfo_height()
                                max_y = max(max_y, frame_bottom)

                y_pos = max_y + 2 if max_y > 0 else 0

        self._group_drop_indicator.place(x=3, y=y_pos, relwidth=1.0, width=-6, height=4)
        self._group_drop_indicator.lift()  # Ensure indicator is visible above other widgets

    def _on_group_header_click(self, event, group_label: str):
        """Handle click on group header - select the group for arrow navigation"""
        if group_label == "__uncategorized__":
            return  # Can't select uncategorized group

        # Clear field selection when selecting a group
        self._init_selection()
        self.selected_items.clear()
        self._update_selection_visuals()

        # Select this group (or deselect if already selected)
        if self._selected_group == group_label:
            self._selected_group = None
        else:
            self._selected_group = group_label

        # Update visual selection for all groups
        self._update_group_selection_visuals()

    def _update_group_selection_visuals(self):
        """Update visual highlighting for selected group"""
        for label, widgets in self._group_widgets.items():
            header_frame = widgets.get("header_frame")
            drag_handle = widgets.get("drag_handle")
            expand_btn = widgets.get("expand_btn")
            name_label = widgets.get("name_label")
            position_label = widgets.get("position_label")
            count_label = widgets.get("count_label")

            if not header_frame:
                continue

            colors = self._theme_manager.colors
            if label == self._selected_group:
                # Selected - highlight with theme selection color (flat design)
                bg_color = colors['selection_highlight']
                selection_border = colors.get('selection_bg', colors['button_primary'])
                header_frame.config(
                    bg=bg_color,
                    highlightbackground=selection_border,
                    highlightcolor=selection_border,
                    highlightthickness=2
                )
                if drag_handle:
                    drag_handle.config(bg=bg_color)
                if expand_btn:
                    expand_btn.config(bg=bg_color)
                if name_label:
                    name_label.config(bg=bg_color)
                if position_label:
                    position_label.config(bg=bg_color)
                if count_label:
                    count_label.config(bg=bg_color)
            else:
                # Not selected - restore normal appearance (flat design)
                bg_color = colors['tree_heading_bg']
                header_frame.config(
                    bg=bg_color,
                    highlightbackground=colors['border'],
                    highlightcolor=colors['border'],
                    highlightthickness=2  # Keep constant to prevent height changes
                )
                if drag_handle:
                    drag_handle.config(bg=bg_color)
                if expand_btn:
                    expand_btn.config(bg=bg_color)
                if name_label:
                    name_label.config(bg=bg_color)
                if position_label:
                    position_label.config(bg=bg_color)
                if count_label:
                    count_label.config(bg=bg_color)

    def _clear_group_selection(self):
        """Clear the currently selected group"""
        self._selected_group = None
        self._update_group_selection_visuals()

    def _on_group_header_right_click(self, event, group_label: str):
        """Show context menu for category/group header"""
        if group_label == "__uncategorized__":
            return  # No context menu for uncategorized

        # Find current index of this group
        current_idx = None
        for idx, (sort_order, label) in enumerate(self._group_order):
            if label == group_label:
                current_idx = idx
                break

        if current_idx is None:
            return

        # Check if fields are aligned - disable reordering if not
        is_aligned, _, _ = self._check_category_alignment()

        # Create themed context menu
        colors = self._theme_manager.colors
        menu = tk.Menu(
            self, tearoff=0,
            bg=colors.get('surface', colors['background']),
            fg=colors['text_primary'],
            activebackground=colors.get('card_surface_hover', colors.get('surface', colors['background'])),
            activeforeground=colors['text_primary'],
            relief='flat',
            font=('Segoe UI', 9)
        )

        num_groups = len(self._group_order)

        # Determine state for move commands - disabled if not aligned or at boundary
        top_state = "normal" if is_aligned and current_idx > 0 else "disabled"
        bottom_state = "normal" if is_aligned and current_idx < num_groups - 1 else "disabled"
        position_state = "normal" if is_aligned else "disabled"

        # Move to Top
        menu.add_command(
            label="Move to Top",
            command=lambda: self._move_group_to_position(group_label, 0),
            state=top_state
        )

        # Move to Position...
        menu.add_command(
            label="Move to Position...",
            command=lambda: self._move_group_to_position_prompt(group_label, current_idx),
            state=position_state
        )

        # Move to Bottom
        menu.add_command(
            label="Move to Bottom",
            command=lambda: self._move_group_to_position(group_label, num_groups - 1),
            state=bottom_state
        )

        # Add hint if not aligned
        if not is_aligned:
            menu.add_separator()
            menu.add_command(
                label="(Align fields first to enable reordering)",
                state="disabled"
            )

        # Show the menu
        menu.tk_popup(event.x_root, event.y_root)

    def _move_group_to_position(self, group_label: str, new_position: int):
        """Move a group to a specific position"""
        self._reorder_group(group_label, new_position)

    def _move_selected_group_up(self):
        """Move the currently selected group up one position"""
        if not self._selected_group:
            return

        # Check alignment first
        is_aligned, _, _ = self._check_category_alignment()
        if not is_aligned:
            return

        # Find current index of selected group
        current_idx = None
        for idx, (sort_order, label) in enumerate(self._group_order):
            if label == self._selected_group:
                current_idx = idx
                break

        if current_idx is None or current_idx == 0:
            return  # Not found or already at top

        # Move group up one position
        self._reorder_group(self._selected_group, current_idx - 1)

        # Re-select the group after re-render (group_widgets was cleared and recreated)
        self._update_group_selection_visuals()

    def _move_selected_group_down(self):
        """Move the currently selected group down one position"""
        if not self._selected_group:
            return

        # Check alignment first
        is_aligned, _, _ = self._check_category_alignment()
        if not is_aligned:
            return

        # Find current index of selected group
        current_idx = None
        for idx, (sort_order, label) in enumerate(self._group_order):
            if label == self._selected_group:
                current_idx = idx
                break

        if current_idx is None or current_idx >= len(self._group_order) - 1:
            return  # Not found or already at bottom

        # Move group down one position (new_position is where it should end up)
        self._reorder_group(self._selected_group, current_idx + 2)

        # Re-select the group after re-render (group_widgets was cleared and recreated)
        self._update_group_selection_visuals()

    def _move_group_to_position_prompt(self, group_label: str, current_idx: int):
        """Prompt user for position and move group there"""
        num_groups = len(self._group_order)

        position_str = ThemedInputDialog.askstring(
            self.container.winfo_toplevel(),
            "Move Category to Position",
            f"Enter position (1-{num_groups}):",
            initialvalue=str(current_idx + 1)
        )

        if not position_str:
            return

        try:
            position = int(position_str)
            if position < 1:
                position = 1
            elif position > num_groups:
                position = num_groups
        except ValueError:
            ThemedMessageBox.showerror(self.container.winfo_toplevel(), ErrorMessages.INVALID_POSITION, ErrorMessages.INVALID_NUMBER)
            return

        # Convert to 0-based index
        self._reorder_group(group_label, position - 1)

    def _reorder_group(self, group_label: str, new_position: int):
        """Reorder a group to a new position and renumber all fields"""
        # Find current position
        current_idx = None
        for idx, (sort_order, label) in enumerate(self._group_order):
            if label == group_label:
                current_idx = idx
                break

        if current_idx is None:
            return

        # Check if move would result in no change
        # new_position is the insertion point (0 = before first, len = after last)
        # If current_idx == new_position, item would be inserted before itself (no change)
        # If current_idx == new_position - 1, item would be inserted after itself (no change)
        if current_idx == new_position or current_idx == new_position - 1:
            return

        # Remove from current position and insert at new position
        item = self._group_order.pop(current_idx)

        # Adjust target position if we removed from before it
        if current_idx < new_position:
            new_position -= 1

        self._group_order.insert(new_position, item)

        # Renumber all fields based on new group order
        self._renumber_fields_by_group_order()

        # Update the preview
        self.main_tab.update_preview()

        # Re-render
        self._render_grouped_view()

    def _renumber_fields_by_group_order(self):
        """Renumber all fields based on the current group order.

        This only reorders categories that are in the current group order (visible).
        Fields in non-visible categories are not affected.
        """
        if not self._group_order:
            return

        # Get category level for grouping
        cat_level = self._category_levels[self._group_by_category_idx]

        # Get all category labels that are being reordered (from _group_order)
        reordering_labels = {label for _, label in self._group_order if label != "__uncategorized__"}

        # Group ALL fields by category (not just filtered)
        all_fields_by_category = {}
        all_uncategorized = []

        for field_item in self.field_items:
            cat_label = ""
            if field_item.categories and len(field_item.categories) > self._group_by_category_idx:
                _, cat_label = field_item.categories[self._group_by_category_idx]

            if cat_label:
                if cat_label not in all_fields_by_category:
                    all_fields_by_category[cat_label] = []
                all_fields_by_category[cat_label].append(field_item)
            else:
                all_uncategorized.append(field_item)

        # Build new field order
        # First, add fields from the reordered groups in their new order
        new_order = []
        for sort_order, group_label in self._group_order:
            if group_label == "__uncategorized__":
                new_order.extend(all_uncategorized)
            elif group_label in all_fields_by_category:
                new_order.extend(all_fields_by_category[group_label])

        # Then add fields from categories NOT in the current group order
        # (these are filtered out but should keep their positions relative to each other)
        for field_item in self.field_items:
            field_id = id(field_item)
            if field_id not in {id(fi) for fi in new_order}:
                # This field is in a category not being reordered
                new_order.append(field_item)

        # Update field_items list
        self.field_items = new_order

        # Update the data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        # Update order numbers
        self._update_order_numbers()

        # Update category sort orders ONLY for categories in the group_order
        # This updates the sort value in field.categories tuples
        for new_idx, (old_sort, group_label) in enumerate(self._group_order):
            if group_label != "__uncategorized__" and group_label in reordering_labels:
                # Update sort order in field categories for ALL fields with this label
                for field_item in self.field_items:
                    if field_item.categories and len(field_item.categories) > self._group_by_category_idx:
                        _, label = field_item.categories[self._group_by_category_idx]
                        if label == group_label:
                            # Update sort order to match new position
                            field_item.categories[self._group_by_category_idx] = (float(new_idx + 1), label)

        # Update group_order with new sort values
        self._group_order = [(float(idx + 1), label) for idx, (_, label) in enumerate(self._group_order)]

        # Also update the category_level's label order to reflect the new order
        # This ensures the order persists when filters are cleared
        new_label_order = []
        for _, label in self._group_order:
            if label != "__uncategorized__" and label in cat_level.labels:
                new_label_order.append(label)
        # Add any labels not in the reordered set at the end
        for label in cat_level.labels:
            if label not in new_label_order:
                new_label_order.append(label)
        cat_level.labels = new_label_order

        self.logger.info(f"Renumbered {len(new_order)} fields based on new group order")

    def _get_group_bounds(self, field_item: FieldItem) -> Tuple[int, int]:
        """Get the index range for the group containing this field in grouped view.

        Returns:
            Tuple of (start_idx, end_idx) where end_idx is exclusive.
            Returns (0, len(field_items)) if not in grouped view or no categories.
        """
        if not hasattr(self, '_fields_by_category') or not self._fields_by_category:
            return (0, len(self.field_items))

        if not self._category_levels or self._group_by_category_idx >= len(self._category_levels):
            return (0, len(self.field_items))

        # Get this field's category label at the current group level
        cat_label = ""
        if field_item.categories and len(field_item.categories) > self._group_by_category_idx:
            _, cat_label = field_item.categories[self._group_by_category_idx]

        # Use "__uncategorized__" for fields without a category
        if not cat_label:
            cat_label = "__uncategorized__"

        # Find all items in the same group
        group_data = self._fields_by_category.get(cat_label, {})
        group_fields = group_data.get("fields", [])

        if not group_fields:
            return (0, len(self.field_items))

        # Get indices of all fields in this group
        indices = []
        for f in group_fields:
            try:
                indices.append(self.field_items.index(f))
            except ValueError:
                continue

        if not indices:
            return (0, len(self.field_items))

        return (min(indices), max(indices) + 1)  # +1 for exclusive end

    def _check_category_alignment(self, detailed: bool = False) -> Tuple[bool, str, dict]:
        """Check if field order aligns with category grouping.

        Args:
            detailed: If True, returns detailed info about misaligned fields

        Returns:
            Tuple of (is_aligned: bool, message: str, details: dict)
            details contains:
                - 'out_of_order_categories': list of category names that appear out of order
                - 'misaligned_fields': list of (field_name, category) tuples for scattered fields
                - 'current_order': list of (field_name, category) showing current sequence
                - 'expected_order': list of (field_name, category) showing aligned sequence
        """
        details = {
            'out_of_order_categories': [],
            'misaligned_fields': [],
            'current_order': [],
            'expected_order': []
        }

        if not self._category_levels or not self.field_items:
            return True, "", details

        cat_idx = self._group_by_category_idx
        if cat_idx >= len(self._category_levels):
            return True, "", details

        # Track last seen category for each group
        last_category = None
        seen_categories = set()
        out_of_order = []
        category_first_seen = {}  # Track where each category was first seen

        # Build current order for detailed view
        for i, field_item in enumerate(self.field_items):
            cat_label = ""
            if field_item.categories and len(field_item.categories) > cat_idx:
                _, cat_label = field_item.categories[cat_idx]

            display_name = field_item.display_name or field_item.field_name
            if detailed:
                details['current_order'].append((display_name, cat_label or "(uncategorized)"))

            if cat_label:
                if cat_label not in category_first_seen:
                    category_first_seen[cat_label] = i

                if cat_label != last_category:
                    if cat_label in seen_categories:
                        out_of_order.append(cat_label)
                        if detailed:
                            details['misaligned_fields'].append((display_name, cat_label))
                    seen_categories.add(cat_label)
                    last_category = cat_label

        details['out_of_order_categories'] = out_of_order

        if out_of_order:
            # Build expected order for comparison
            if detailed:
                cat_level = self._category_levels[cat_idx]
                fields_by_category = {}
                uncategorized = []

                for field_item in self.field_items:
                    cat_label = ""
                    cat_sort = 999.0
                    if field_item.categories and len(field_item.categories) > cat_idx:
                        cat_sort, cat_label = field_item.categories[cat_idx]

                    display_name = field_item.display_name or field_item.field_name
                    if cat_label:
                        if cat_label not in fields_by_category:
                            fields_by_category[cat_label] = {"sort": cat_sort, "fields": []}
                        fields_by_category[cat_label]["fields"].append((display_name, cat_label))
                    else:
                        uncategorized.append((display_name, "(uncategorized)"))

                sorted_categories = sorted(
                    [(data["sort"], label) for label, data in fields_by_category.items()],
                    key=lambda x: x[0]
                )

                for _, cat_label in sorted_categories:
                    details['expected_order'].extend(fields_by_category[cat_label]["fields"])
                details['expected_order'].extend(uncategorized)

            return False, f"Fields are not grouped by category. Categories appearing out of order: {', '.join(out_of_order[:3])}", details

        return True, "", details

    def _align_to_category_order(self):
        """Align field order to match category grouping"""
        if not self._category_levels or not self.field_items:
            return

        cat_idx = self._group_by_category_idx
        if cat_idx >= len(self._category_levels):
            return

        cat_level = self._category_levels[cat_idx]

        # Group fields by category
        fields_by_category = {}
        uncategorized = []

        for field_item in self.field_items:
            cat_label = ""
            cat_sort = 999.0
            if field_item.categories and len(field_item.categories) > cat_idx:
                cat_sort, cat_label = field_item.categories[cat_idx]

            if cat_label:
                if cat_label not in fields_by_category:
                    fields_by_category[cat_label] = {"sort": cat_sort, "fields": []}
                fields_by_category[cat_label]["fields"].append(field_item)
            else:
                uncategorized.append(field_item)

        # Build new order based on category sort order
        sorted_categories = sorted(
            [(data["sort"], label) for label, data in fields_by_category.items()],
            key=lambda x: x[0]
        )

        new_order = []
        for _, cat_label in sorted_categories:
            new_order.extend(fields_by_category[cat_label]["fields"])
        new_order.extend(uncategorized)

        # Update field_items
        self.field_items = new_order

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        # Update order numbers and refresh view
        self._update_order_numbers()

        if self._view_mode == "grouped":
            self._render_grouped_view()
        else:
            self._repack_all_widgets()

        self.main_tab.update_preview()
        self.logger.info("Aligned field order to category grouping")

    def _edit_display_name(self, field_item: FieldItem, display_var: tk.StringVar):
        """Edit display name inline (legacy - uses StringVar)"""
        new_name = ThemedInputDialog.askstring(
            self.container.winfo_toplevel(),
            "Edit Display Name",
            f"Enter new display name for:\n{field_item.field_name}",
            initialvalue=field_item.display_name,
            min_width=400
        )

        if new_name and new_name != field_item.display_name:
            display_var.set(new_name)
            self.main_tab.on_field_display_name_changed(field_item, new_name)

    def _edit_display_name_direct(self, field_item: FieldItem, display_label: tk.Label):
        """Edit display name inline (optimized - updates label directly)"""
        new_name = ThemedInputDialog.askstring(
            self.winfo_toplevel(),
            "Edit Display Name",
            f"Enter new display name for:\n{field_item.field_name}",
            initialvalue=field_item.display_name,
            min_width=400
        )

        if new_name and new_name != field_item.display_name:
            display_label.config(text=new_name)
            self.main_tab.on_field_display_name_changed(field_item, new_name)

    def _revert_single_name(self, field_item: FieldItem, display_var: tk.StringVar):
        """Revert single field name to original (legacy)"""
        display_var.set(field_item.field_name)
        self.main_tab.on_field_display_name_changed(field_item, field_item.field_name)
    
    def _on_revert_all(self):
        """Revert all display names"""
        self.main_tab.on_revert_all_display_names()
    
    def _on_category_changed(self, field_item: FieldItem, categories: list):
        """Handle category changes from table editor"""
        # Category assignment now handled via table editor popup
        self.main_tab.on_field_category_changed(field_item, categories)
    
    def _on_remove_field(self, field_item: FieldItem):
        """Handle remove button click"""
        self.main_tab.on_remove_field(field_item)
    
    def _update_order_numbers(self):
        """Update order number labels - shows actual position in full list - OPTIMIZED"""
        # Build category position map if in grouped view
        cat_positions = {}

        # Always calculate category positions if we have category levels, regardless of view mode
        # This ensures the data is ready when switching to grouped view
        if self._category_levels and self._group_by_category_idx < len(self._category_levels):
            cat_idx = self._group_by_category_idx
            # Group fields by category and track position within each
            fields_by_cat = {}
            for field_item in self.field_items:
                cat_label = ""
                if field_item.categories and len(field_item.categories) > cat_idx:
                    _, cat_label = field_item.categories[cat_idx]
                if not cat_label:
                    cat_label = "__uncategorized__"

                if cat_label not in fields_by_cat:
                    fields_by_cat[cat_label] = []
                fields_by_cat[cat_label].append(field_item)
                cat_positions[id(field_item)] = len(fields_by_cat[cat_label])

        # Update labels and order_within_group values
        for idx, field_item in enumerate(self.field_items, 1):
            # Update the actual order_within_group value to match new position
            field_item.order_within_group = idx

            field_id = id(field_item)
            if field_id in self.field_widgets:
                # Update main order label
                label = self.field_widgets[field_id]["order_label"]
                new_text = f"{idx}."
                # Only update if changed (avoids unnecessary redraws)
                if label.cget("text") != new_text:
                    label.config(text=new_text)

                # Update alternating row colors (idx is 1-based, so idx-1 for 0-based)
                self._update_row_colors(field_id, idx - 1)

                # Update category position label - recreate widget if cat_pos_label is missing
                if "cat_pos_label" not in self.field_widgets[field_id]:
                    # Widget was created before cat_pos_label feature - recreate it
                    old_frame = self.field_widgets[field_id]["frame"]
                    was_packed = old_frame.winfo_ismapped()
                    pack_info = None
                    if was_packed:
                        try:
                            pack_info = old_frame.pack_info()
                        except tk.TclError:
                            pack_info = None

                    # Destroy old widget
                    old_frame.destroy()
                    del self.field_widgets[field_id]

                    # Find position and recreate
                    try:
                        orig_idx = self.field_items.index(field_item)
                    except ValueError:
                        orig_idx = idx - 1
                    self._create_field_widget_fast(field_item, orig_idx)

                    # Repack if it was packed before
                    if was_packed and pack_info and field_id in self.field_widgets:
                        self.field_widgets[field_id]["frame"].pack(**pack_info)

                # Now update the category position text
                if field_id in self.field_widgets and "cat_pos_label" in self.field_widgets[field_id]:
                    cat_pos_label = self.field_widgets[field_id]["cat_pos_label"]
                    # Only show category position in grouped view
                    if self._view_mode == "grouped" and field_id in cat_positions:
                        cat_text = f"({cat_positions[field_id]})"
                    else:
                        cat_text = ""
                    cat_pos_label.config(text=cat_text)

    def _update_row_colors(self, field_id: int, position: int):
        """Update alternating row background colors for a field widget."""
        if field_id not in self.field_widgets:
            return

        colors = self._theme_manager.colors
        row_bg = colors['card_surface'] if position % 2 == 0 else colors['surface']

        widgets = self.field_widgets[field_id]
        frame = widgets.get("frame")
        if not frame:
            return

        # Update frame background and border
        try:
            current_bg = frame.cget("bg")
            if current_bg != row_bg:
                frame.config(bg=row_bg, highlightbackground=row_bg)

                # Update all label children
                for key in ("order_label", "cat_pos_label", "drag_handle", "display_label", "ref_label"):
                    label = widgets.get(key)
                    if label:
                        label.config(bg=row_bg)
        except tk.TclError:
            pass  # Widget may have been destroyed

    def _repack_all_widgets(self):
        """Repack all field widgets in the correct order, respecting filter - OPTIMIZED"""
        if self._virtual_mode:
            # Virtual mode - refresh the virtual view instead of repacking
            self._refresh_virtual_view()
            return

        # Standard mode - freeze container during repack
        self.fields_container.pack_propagate(False)

        # First unpack all
        for field_item in self.field_items:
            field_id = id(field_item)
            if field_id in self.field_widgets:
                self.field_widgets[field_id]["frame"].pack_forget()

        # Then pack in order - only items matching filter
        for field_item in self.field_items:
            field_id = id(field_item)
            if field_id in self.field_widgets:
                if self._field_matches_filter(field_item):
                    self.field_widgets[field_id]["frame"].pack(fill=tk.X, pady=1, padx=3)
                    # Reset internal padding to default for flat view
                    order_label = self.field_widgets[field_id]["order_label"]
                    order_label.pack_configure(padx=(4, 0))

        # Re-enable propagation
        self.fields_container.pack_propagate(True)

    def _repack_grouped_view(self):
        """Lightweight repack for within-group reorders - keeps headers intact.

        This is much faster than _render_grouped_view() because it:
        - Skips _clear_grouped_view() (keeps existing group headers)
        - Only repacks field widgets in their new order
        - No update_idletasks() calls (lets tkinter batch updates)
        """
        if not hasattr(self, '_fields_by_category') or not self._fields_by_category:
            # Fall back to full render if no category data
            self._render_grouped_view()
            return

        if not hasattr(self, '_group_widgets') or not self._group_widgets:
            # Fall back to full render if no group widgets
            self._render_grouped_view()
            return

        # Update _fields_by_category to reflect new order in field_items
        for group_label, group_data in self._fields_by_category.items():
            # Preserve only fields that are still in the group, in their new order
            # Use id() since FieldItem objects are not hashable
            old_field_ids = {id(f) for f in group_data["fields"]}
            group_data["fields"] = [f for f in self.field_items if id(f) in old_field_ids]

        # Unpack all field widgets (but NOT group containers/headers)
        for field_id, widgets in self.field_widgets.items():
            widgets["frame"].pack_forget()

        # Repack in group order, positioning fields AFTER their group_container
        # (Fields are packed into fields_container, not inside group_container)
        for _, group_label in self._group_order:
            if group_label in self._collapsed_groups:
                continue  # Skip collapsed groups

            # Get the group container to pack fields after
            group_widget_data = self._group_widgets.get(group_label, {})
            group_container = group_widget_data.get("group_container")
            if not group_container:
                continue

            group_fields = self._fields_by_category.get(group_label, {}).get("fields", [])

            # First field packs after group_container, subsequent after previous field
            prev_widget = group_container
            for field_item in group_fields:
                if not self._field_matches_filter(field_item):
                    continue
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    frame = self.field_widgets[field_id]["frame"]
                    frame.pack(fill=tk.X, pady=1, padx=3, after=prev_widget)
                    prev_widget = frame

    def _on_edit_categories(self):
        """Open table editor dialog for category assignments"""
        category_levels = self.main_tab.current_parameter.category_levels if self.main_tab.current_parameter else []

        def on_save():
            # Sync field order from dialog back to builder UI
            # (field_items list was modified in place by the dialog)
            self._update_order_numbers()
            # Re-render the appropriate view (grouped view needs full re-render for category changes)
            if self._view_mode == "grouped":
                self._render_grouped_view()
            else:
                self._repack_all_widgets()
            # Update data model
            if self.main_tab.current_parameter:
                self.main_tab.current_parameter.fields = self.field_items.copy()
            self.main_tab.update_preview()
            # Update alignment button if in grouped view mode
            self.update_alignment_status()

        FieldCategoryEditorDialog(
            self.winfo_toplevel(),
            self.field_items,
            category_levels,
            on_save
        )

    def update_edit_categories_button(self):
        """Update the Edit Categories button state based on fields and category columns"""
        has_fields = len(self.field_items) > 0
        has_categories = (
            self.main_tab.current_parameter and
            len(self.main_tab.current_parameter.category_levels) > 0
        )
        enabled = has_fields and has_categories
        self.edit_categories_btn.set_enabled(enabled)

    # =========================================================================
    # SELECTION AND DRAG-DROP
    # =========================================================================

    def _init_selection(self):
        """Initialize selection tracking"""
        if not hasattr(self, 'selected_items'):
            self.selected_items: set = set()  # Set of field_ids

    def _on_field_click(self, event, field_item: FieldItem):
        """Handle click on a field row for selection and prepare for potential drag"""
        self._init_selection()
        field_id = id(field_item)
        ctrl_held = event.state & 0x4
        shift_held = event.state & 0x1

        # Clear group selection when clicking a field
        if self._selected_group:
            self._clear_group_selection()

        # Track if we should defer selection change until button release
        # This prevents deselecting items when starting a drag
        self.drag_data["pending_deselect"] = None

        if ctrl_held:
            # Toggle selection immediately with Ctrl
            if field_id in self.selected_items:
                self.selected_items.discard(field_id)
            else:
                self.selected_items.add(field_id)
        elif shift_held and self.selected_items:
            # Range select from last selected to this one
            current_idx = self.field_items.index(field_item)
            selected_indices = [self.field_items.index(fi) for fi in self.field_items if id(fi) in self.selected_items]
            if selected_indices:
                min_idx = min(min(selected_indices), current_idx)
                max_idx = max(max(selected_indices), current_idx)
                for i in range(min_idx, max_idx + 1):
                    self.selected_items.add(id(self.field_items[i]))
        else:
            # Single click behavior - but defer deselection if item is already selected
            # (user might be starting a drag of multiple selected items)
            if field_id in self.selected_items:
                # Don't deselect yet - mark as pending and wait for button release
                # If user drags, we keep selection; if user releases without drag, we deselect
                self.drag_data["pending_deselect"] = field_id
            else:
                # Clicking unselected item - select only this item (clear others)
                self.selected_items.clear()
                self.selected_items.add(field_id)

        self._update_selection_visuals()

        # Store for potential drag operation
        idx = self.field_items.index(field_item) if field_item in self.field_items else -1
        self.drag_data["item"] = field_item
        self.drag_data["index"] = idx
        self.drag_data["dragging"] = False  # Will be True on first motion

    def _update_selection_visuals(self, force_update: bool = False):
        """Update visual highlighting of selected rows - OPTIMIZED for rapid clicks

        Args:
            force_update: If True, updates all selected items (for theme changes).
                         If False, only updates items that changed selection state.
        """
        self._init_selection()

        # Track which items actually changed to minimize UI updates
        if not hasattr(self, '_previous_selected'):
            self._previous_selected = set()

        if force_update:
            # On theme change, we need to update all selected AND all previously selected items
            # to apply new colors (selected get new border color, deselected get new zebra colors)
            changed_items = self.selected_items | self._previous_selected | set(self.field_widgets.keys())
        else:
            # Calculate which items changed state
            newly_selected = self.selected_items - self._previous_selected
            newly_deselected = self._previous_selected - self.selected_items
            changed_items = newly_selected | newly_deselected

        # Only update widgets that actually changed (much faster for rapid clicks)
        for field_id in changed_items:
            if field_id not in self.field_widgets:
                continue

            widgets = self.field_widgets[field_id]
            frame = widgets["frame"]
            field_item = widgets.get("field_item")
            is_selected = field_id in self.selected_items

            colors = self._theme_manager.colors
            if is_selected:
                # Highlight selected rows with selection color and border matching drag/drop
                bg_color = colors['selection_highlight']
                selection_border = colors['button_primary']  # Match drag/drop indicator color
                try:
                    frame.config(
                        bg=bg_color,
                        highlightbackground=selection_border,
                        highlightcolor=selection_border,
                        highlightthickness=2
                    )
                except:
                    pass
                # Apply background to all children (tk.Label or ttk.Label)
                for child in frame.winfo_children():
                    try:
                        if isinstance(child, tk.Label):
                            child.config(bg=bg_color)
                        elif isinstance(child, ttk.Label):
                            child.config(background=bg_color)
                    except:
                        pass
            else:
                # Reset to zebra stripe color (no border)
                # Determine row position for alternating colors
                try:
                    position = self.field_items.index(field_item) if field_item else 0
                except (ValueError, AttributeError):
                    position = 0
                default_bg = colors['card_surface'] if position % 2 == 0 else colors['surface']
                try:
                    frame.config(
                        bg=default_bg,
                        highlightbackground=default_bg,  # Match background for invisible border
                        highlightcolor=default_bg,
                        highlightthickness=2  # Keep constant to prevent height changes in grouped view
                    )
                except:
                    pass
                for child in frame.winfo_children():
                    try:
                        if isinstance(child, tk.Label):
                            child.config(bg=default_bg)
                        elif isinstance(child, ttk.Label):
                            child.config(background="")
                    except:
                        pass

        # Update previous state
        self._previous_selected = self.selected_items.copy()

    def _apply_selection_to_widgets(self, field_ids: list):
        """Apply selection highlighting to specific widget IDs (for virtual scrolling)"""
        self._init_selection()
        colors = self._theme_manager.colors
        for field_id in field_ids:
            if field_id not in self.field_widgets:
                continue

            widgets = self.field_widgets[field_id]
            frame = widgets["frame"]
            is_selected = field_id in self.selected_items

            if is_selected:
                # Highlight selected rows with selection color and border matching drag/drop
                bg_color = colors['selection_highlight']
                selection_border = colors['button_primary']  # Match drag/drop indicator color
                try:
                    frame.config(
                        bg=bg_color,
                        highlightbackground=selection_border,
                        highlightcolor=selection_border,
                        highlightthickness=2
                    )
                except:
                    pass
                for child in frame.winfo_children():
                    try:
                        if isinstance(child, tk.Label):
                            child.config(bg=bg_color)
                    except:
                        pass

    def _on_field_right_click(self, event, field_item: FieldItem):
        """Handle right-click for context menu"""
        self._init_selection()
        field_id = id(field_item)

        # If right-clicked item not in selection, select only it
        if field_id not in self.selected_items:
            self.selected_items.clear()
            self.selected_items.add(field_id)
            self._update_selection_visuals()

        # Show themed context menu
        menu = ThemedContextMenu(self, self._theme_manager)
        count = len(self.selected_items)
        label_suffix = f" ({count} items)" if count > 1 else ""

        # Check if move commands should be disabled in grouped view when not aligned
        move_enabled = True
        if self._view_mode == "grouped":
            is_aligned, _, _ = self._check_category_alignment()
            if not is_aligned:
                move_enabled = False

        menu.add_command(label=f"Move to Top{label_suffix}", command=self._move_selected_to_top, enabled=move_enabled)
        menu.add_command(label=f"Move to Position...{label_suffix}", command=self._move_selected_to_position, enabled=move_enabled)
        menu.add_command(label=f"Move to Bottom{label_suffix}", command=self._move_selected_to_bottom, enabled=move_enabled)

        # Add hint if move is disabled
        if not move_enabled:
            menu.add_separator()
            menu.add_command(label="(Click Align to enable reordering)", enabled=False)

        menu.add_separator()
        menu.add_command(label=f"Remove{label_suffix}", command=self._remove_selected)

        menu.show(event.x_root, event.y_root)

    def _move_selected_to_top(self):
        """Move selected items to the top of the list (or top of category in grouped view)"""
        self._init_selection()
        if not self.selected_items:
            return

        # In grouped view, move within category only
        if self._view_mode == "grouped" and self._category_levels:
            self._move_selected_within_category("top")
            return

        # Get selected field items in their current order
        selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_items]
        remaining_fields = [fi for fi in self.field_items if id(fi) not in self.selected_items]

        # New order: selected first, then remaining
        self.field_items = selected_fields + remaining_fields

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        self._update_order_numbers()
        self._repack_all_widgets()
        self._update_selection_visuals()
        self.main_tab.update_preview()

    def _move_selected_to_bottom(self):
        """Move selected items to the bottom of the list (or bottom of category in grouped view)"""
        self._init_selection()
        if not self.selected_items:
            return

        # In grouped view, move within category only
        if self._view_mode == "grouped" and self._category_levels:
            self._move_selected_within_category("bottom")
            return

        # Get selected field items in their current order
        selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_items]
        remaining_fields = [fi for fi in self.field_items if id(fi) not in self.selected_items]

        # New order: remaining first, then selected
        self.field_items = remaining_fields + selected_fields

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        self._update_order_numbers()
        self._repack_all_widgets()
        self._update_selection_visuals()
        self.main_tab.update_preview()

    def _move_selected_to_position(self):
        """Move selected items to a specific position (prompts user)"""
        self._init_selection()
        if not self.selected_items:
            return

        # In grouped view, move within category only
        if self._view_mode == "grouped" and self._category_levels:
            self._move_selected_within_category("position")
            return

        total_fields = len(self.field_items)
        current_position = None

        # Get current position of first selected item for default
        for idx, fi in enumerate(self.field_items):
            if id(fi) in self.selected_items:
                current_position = idx + 1  # 1-based for user
                break

        # Prompt user for position
        position_str = ThemedInputDialog.askstring(
            self.container.winfo_toplevel(),
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
            ThemedMessageBox.showerror(self.container.winfo_toplevel(), ErrorMessages.INVALID_POSITION, ErrorMessages.INVALID_NUMBER)
            return

        # Convert to 0-based index for insertion
        # User says "position 7" = item should END UP at index 6 (0-based)
        target_idx = position - 1

        # Get selected field items in their current order
        selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_items]
        remaining_fields = [fi for fi in self.field_items if id(fi) not in self.selected_items]

        # Simple approach: insert at target_idx in remaining list
        # Clamp to valid range
        insert_pos = min(target_idx, len(remaining_fields))

        # Insert selected items at the target position
        new_order = remaining_fields[:insert_pos] + selected_fields + remaining_fields[insert_pos:]
        self.field_items = new_order

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        self._update_order_numbers()
        if self._view_mode == "grouped":
            self._render_grouped_view()
        else:
            self._repack_all_widgets()
        self._update_selection_visuals()
        self.main_tab.update_preview()

        # Scroll to show the moved item at its new position
        if self._virtual_mode and self._filtered_items:
            # Find the first selected item in the filtered list
            for fi in selected_fields:
                if fi in self._filtered_items:
                    idx = self._filtered_items.index(fi)
                    total_height = len(self._filtered_items) * self.ROW_HEIGHT
                    if total_height > 0:
                        scroll_pos = (idx * self.ROW_HEIGHT) / total_height
                        self.canvas.yview_moveto(max(0, min(1, scroll_pos)))
                    break

    def _move_selected_within_category(self, move_type: str):
        """Move selected items within their category in grouped view.

        Args:
            move_type: "top", "bottom", or "position"
        """
        if not self.selected_items or not self._category_levels:
            return

        # Get the first selected item to determine the category
        first_selected = None
        for fi in self.field_items:
            if id(fi) in self.selected_items:
                first_selected = fi
                break

        if not first_selected:
            return

        # Get the category label for the selected item
        selected_cat_label = ""
        if first_selected.categories and len(first_selected.categories) > self._group_by_category_idx:
            _, selected_cat_label = first_selected.categories[self._group_by_category_idx]

        if not selected_cat_label:
            selected_cat_label = "__uncategorized__"

        # Get all fields in the same category (in their current order within field_items)
        category_fields = []
        category_indices = []  # Track original indices in field_items
        for idx, fi in enumerate(self.field_items):
            cat_label = ""
            if fi.categories and len(fi.categories) > self._group_by_category_idx:
                _, cat_label = fi.categories[self._group_by_category_idx]
            if not cat_label:
                cat_label = "__uncategorized__"

            if cat_label == selected_cat_label:
                category_fields.append(fi)
                category_indices.append(idx)

        if not category_fields:
            return

        # Separate selected and non-selected within this category
        selected_in_cat = [fi for fi in category_fields if id(fi) in self.selected_items]
        remaining_in_cat = [fi for fi in category_fields if id(fi) not in self.selected_items]

        if move_type == "top":
            # Selected items go to top of category
            new_cat_order = selected_in_cat + remaining_in_cat
        elif move_type == "bottom":
            # Selected items go to bottom of category
            new_cat_order = remaining_in_cat + selected_in_cat
        elif move_type == "position":
            # Prompt for position within category
            total_in_cat = len(category_fields)
            current_position = None

            # Find current position of first selected within category
            for idx, fi in enumerate(category_fields):
                if id(fi) in self.selected_items:
                    current_position = idx + 1  # 1-based for user
                    break

            cat_display = "(Uncategorized)" if selected_cat_label == "__uncategorized__" else selected_cat_label
            position_str = ThemedInputDialog.askstring(
                self.container.winfo_toplevel(),
                "Move to Position in Category",
                f"Enter position within '{cat_display}' (1-{total_in_cat}):",
                initialvalue=str(current_position) if current_position else "1"
            )

            if not position_str:
                return

            try:
                position = int(position_str)
                if position < 1:
                    position = 1
                elif position > total_in_cat:
                    position = total_in_cat
            except ValueError:
                ThemedMessageBox.showerror(self.container.winfo_toplevel(), ErrorMessages.INVALID_POSITION, ErrorMessages.INVALID_NUMBER)
                return

            # Convert to 0-based index
            target_idx = position - 1
            insert_pos = min(target_idx, len(remaining_in_cat))
            new_cat_order = remaining_in_cat[:insert_pos] + selected_in_cat + remaining_in_cat[insert_pos:]
        else:
            return

        # Now rebuild field_items with the new category order
        # We need to replace the fields at category_indices with new_cat_order
        new_field_items = list(self.field_items)
        for i, orig_idx in enumerate(category_indices):
            new_field_items[orig_idx] = new_cat_order[i]

        self.field_items = new_field_items

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        self._update_order_numbers()
        self._render_grouped_view()
        self._update_selection_visuals()
        self.main_tab.update_preview()

    def _remove_selected(self):
        """Remove all selected items"""
        self._init_selection()
        if not self.selected_items:
            return

        # Get items to remove
        items_to_remove = [
            fi for fi in self.field_items
            if id(fi) in self.selected_items
        ]

        if not items_to_remove:
            return

        # Confirmation dialog (same as X button)
        count = len(items_to_remove)
        message = f"Delete {count} field{'s' if count > 1 else ''}?"
        if not ThemedMessageBox.askyesno(
            self.winfo_toplevel(),
            "Confirm Delete",
            message
        ):
            return

        # Remove selected items
        for field_item in items_to_remove:
            self.main_tab.on_remove_field(field_item)

        self.selected_items.clear()

        # Force re-render to fix visual gaps after bulk removal
        if self._view_mode == "grouped":
            self._render_grouped_view()
        else:
            # Repack widgets to close gaps, then update order numbers
            self._repack_all_widgets()
            self._update_order_numbers()

    def _on_drag_motion(self, event, field_item: FieldItem):
        """Handle drag motion for reordering"""
        # Only process if we have an item from click
        if self.drag_data.get("item") is None:
            return

        # Mark as actually dragging (not just a click)
        self.drag_data["dragging"] = True
        # Cancel any pending deselect since we're now dragging
        self.drag_data["pending_deselect"] = None

        # Auto-scroll if near edge of visible area
        self._auto_scroll_if_near_edge(event)

        # Calculate drop position from mouse Y
        canvas_y = self.canvas.canvasy(event.y_root - self.canvas.winfo_rooty())
        new_position = self._calculate_drop_position(canvas_y)

        # In grouped view, constrain drop position to within the same group
        if self._view_mode == "grouped":
            drag_item = self.drag_data.get("item")
            if drag_item:
                start, end = self._get_group_bounds(drag_item)
                new_position = max(start, min(new_position, end))

        # Show drop indicator
        if new_position != self._drop_position:
            self._drop_position = new_position
            colors = self._theme_manager.colors
            if self._drop_indicator is None:
                self._drop_indicator = tk.Frame(self.fields_container, height=4, bg=colors['button_primary'])
            self._show_indicator_at_position(self._drop_position)

    def _end_internal_drag(self, event):
        """End internal dragging and reorder"""
        if self.drag_data.get("item") is None:
            return

        # Capture drop position BEFORE hiding indicator (which resets it)
        drop_pos = self._drop_position

        # Hide drop indicator
        self.hide_drop_indicator()

        # Only reorder if we actually dragged (not just clicked)
        if self.drag_data.get("dragging") and drop_pos is not None:
            self._perform_reorder(drop_pos)
        else:
            # Not a drag - handle pending deselect (click on already-selected item)
            pending_deselect = self.drag_data.get("pending_deselect")
            if pending_deselect is not None:
                self.selected_items.discard(pending_deselect)

        # Reset
        self.drag_data = {"item": None, "index": None, "dragging": False, "pending_deselect": None}
        self._update_selection_visuals()

    def _perform_reorder(self, target_position: int):
        """Reorder selected items to the target position"""
        self._init_selection()
        if not self.selected_items:
            return

        # In grouped view, block reorder if selected items span multiple groups
        if self._view_mode == "grouped":
            # Get all selected field items
            selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_items]
            if selected_fields:
                # Check that all selected items are in the same group
                first_bounds = self._get_group_bounds(selected_fields[0])
                for fi in selected_fields[1:]:
                    if self._get_group_bounds(fi) != first_bounds:
                        # Items span multiple groups - block operation
                        self.logger.debug("Reorder blocked: selection spans multiple groups")
                        return

                # Also verify target position is within group bounds
                start, end = first_bounds
                if target_position < start or target_position > end:
                    self.logger.debug("Reorder blocked: target position outside group bounds")
                    return

        # Store original order for comparison
        original_order = list(self.field_items)

        # Get selected items in current order
        selected_fields = [fi for fi in self.field_items if id(fi) in self.selected_items]
        remaining_fields = [fi for fi in self.field_items if id(fi) not in self.selected_items]

        # Calculate insert position in the remaining list
        # Count how many non-selected items are before target position
        insert_pos = 0
        for i, fi in enumerate(self.field_items):
            if i >= target_position:
                break
            if id(fi) not in self.selected_items:
                insert_pos += 1

        # Insert selected items at the calculated position
        new_order = remaining_fields[:insert_pos] + selected_fields + remaining_fields[insert_pos:]

        # Skip refresh if order didn't actually change (e.g., dropped in same spot)
        if new_order == original_order:
            return

        self.field_items = new_order

        # Update data model
        if self.main_tab.current_parameter:
            self.main_tab.current_parameter.fields = self.field_items.copy()

        self._update_order_numbers()
        # Use lightweight repack for grouped view (fast - keeps headers intact)
        if self._view_mode == "grouped":
            self._repack_grouped_view()
        else:
            self._repack_all_widgets()

        # Refresh zebra stripe colors based on new positions (force_update=True updates ALL items)
        self._update_selection_visuals(force_update=True)

        self.main_tab.update_preview()
    
    def load_parameter(self, parameter: 'FieldParameter'):
        """Load an existing parameter into the builder - OPTIMIZED for large datasets"""
        field_count = len(parameter.fields)
        self.logger.info(f"Loading parameter with {field_count} fields")

        # Clear existing (but keep any pre-existing loading overlay)
        self.clear()

        # Use existing overlay if present, otherwise create one for large datasets
        use_existing_overlay = hasattr(self, '_loading_overlay') and self._loading_overlay
        if field_count > 50 and not use_existing_overlay:
            self.show_loading_overlay(f"Loading {field_count} fields...")

        # 1. Add all field items to the list first (no widgets yet)
        self.field_items = list(parameter.fields)

        # 2. Freeze UI updates during loading
        self.canvas.config(cursor="watch")

        # Decide whether to use virtual scrolling based on item count
        use_virtual = field_count > self.VIRTUAL_SCROLL_THRESHOLD

        if use_virtual:
            # VIRTUAL SCROLLING MODE - only create widgets for visible items
            self.logger.info(f"Using virtual scrolling mode for {field_count} fields")
            self.update_loading_progress("Initializing virtual view...")

            # Setup virtual scrolling (will create only visible widgets)
            self._setup_virtual_scrolling()
        else:
            # STANDARD MODE - create all widgets
            self.fields_container.pack_propagate(False)

            # Create all widgets in one pass
            batch_size = 50
            for idx, field_item in enumerate(self.field_items):
                self._create_field_widget_fast(field_item, idx)
                if (idx + 1) % batch_size == 0:
                    self.update_loading_progress(f"Creating widgets... {idx + 1}/{field_count}")

            # Pack all widgets
            self.update_loading_progress("Rendering fields...")
            for field_item in self.field_items:
                field_id = id(field_item)
                if field_id in self.field_widgets:
                    self.field_widgets[field_id]["frame"].pack(fill=tk.X, pady=1, padx=3)

            self.fields_container.pack_propagate(True)

        self._update_empty_state()
        self.canvas.config(cursor="")

        # Remove loading overlay
        self.hide_loading_overlay()

        # Update category options if there are category levels
        if parameter.category_levels:
            self.logger.info(f"Updating category options for {len(parameter.category_levels)} levels")
            self.update_category_options(parameter.category_levels)

        # Update Edit Categories button state (enable if we have fields and categories)
        self.update_edit_categories_button()

        # Enable bottom toolbar now that parameter is loaded
        self._set_bottom_toolbar_enabled(True)

        self.logger.info(f"Finished loading parameter, {len(self.field_items)} fields in builder (virtual={use_virtual})")

    def _create_field_widget_fast(self, field_item: FieldItem, position: int):
        """Create field widget optimized for batch loading - minimal overhead"""
        field_id = id(field_item)
        colors = self._theme_manager.colors
        # Alternating row colors for zebra striping
        row_bg = colors['card_surface'] if position % 2 == 0 else colors['surface']

        # Use tk.Frame (faster than ttk.Frame) - flat design without border
        field_frame = tk.Frame(self.fields_container, bg=row_bg)

        # Pre-create bound functions once (avoid lambda overhead in bindings)
        click_handler = lambda e, fi=field_item: self._on_field_click(e, fi)
        motion_handler = lambda e, fi=field_item: self._on_drag_motion(e, fi)
        right_click_handler = lambda e, fi=field_item: self._on_field_right_click(e, fi)

        # Bind to frame - use add="+" to allow multiple bindings
        field_frame.bind("<Button-1>", click_handler)
        field_frame.bind("<B1-Motion>", motion_handler)
        field_frame.bind("<ButtonRelease-1>", self._end_internal_drag)
        field_frame.bind("<Button-3>", right_click_handler)
        field_frame.bind("<MouseWheel>", self._on_mousewheel)

        # Order number - use tk.Label (faster)
        order_label = tk.Label(field_frame, text=f"{position + 1}.", font=("Segoe UI", 9, "bold"),
                               width=4, anchor="e", bg=row_bg, fg=colors['text_primary'])
        order_label.pack(side=tk.LEFT, padx=(4, 0), pady=2)
        order_label.bind("<Button-1>", click_handler)
        order_label.bind("<B1-Motion>", motion_handler)
        order_label.bind("<ButtonRelease-1>", self._end_internal_drag)
        order_label.bind("<Button-3>", right_click_handler)
        order_label.bind("<MouseWheel>", self._on_mousewheel)

        # Category position (within-group position) - shown in muted, smaller
        cat_pos_label = tk.Label(field_frame, text="", font=("Segoe UI", 7),
                                 fg=colors['text_muted'], width=4, anchor="w", bg=row_bg)
        cat_pos_label.pack(side=tk.LEFT, padx=(0, 2), pady=2)
        cat_pos_label.bind("<Button-1>", click_handler)
        cat_pos_label.bind("<B1-Motion>", motion_handler)
        cat_pos_label.bind("<ButtonRelease-1>", self._end_internal_drag)
        cat_pos_label.bind("<Button-3>", right_click_handler)
        cat_pos_label.bind("<MouseWheel>", self._on_mousewheel)

        # Drag handle - use tk.Label
        drag_handle = tk.Label(field_frame, text="☰", font=("Segoe UI", 10),
                               cursor="hand2", bg=row_bg, fg=colors['text_secondary'])
        drag_handle.pack(side=tk.LEFT, padx=(0, 4), pady=2)
        drag_handle.bind("<Button-1>", click_handler)
        drag_handle.bind("<B1-Motion>", motion_handler)
        drag_handle.bind("<ButtonRelease-1>", self._end_internal_drag)
        drag_handle.bind("<Button-3>", right_click_handler)
        drag_handle.bind("<MouseWheel>", self._on_mousewheel)

        # Display name - use tk.Label directly (no StringVar overhead for initial load)
        display_label = tk.Label(
            field_frame, text=field_item.display_name, font=("Segoe UI", 9, "bold"),
            fg=colors['title_color'], anchor="w", cursor="hand2", bg=row_bg
        )
        display_label.pack(side=tk.LEFT, padx=(0, 4), pady=2, fill=tk.X, expand=True)
        display_label.bind("<Button-1>", click_handler)
        display_label.bind("<B1-Motion>", motion_handler)
        display_label.bind("<ButtonRelease-1>", self._end_internal_drag)
        display_label.bind("<Button-3>", right_click_handler)
        display_label.bind("<MouseWheel>", self._on_mousewheel)
        display_label.bind("<Double-Button-1>",
                          lambda e, fi=field_item: self._edit_display_name_direct(fi, display_label))

        # Field reference - fixed width, truncated
        ref_label = tk.Label(
            field_frame, text=field_item.field_reference, font=("Consolas", 8),
            fg=colors['text_muted'], anchor="w", width=35, bg=row_bg
        )
        ref_label.pack(side=tk.RIGHT, padx=(0, 2), pady=2)
        ref_label.bind("<Button-1>", click_handler)
        ref_label.bind("<B1-Motion>", motion_handler)
        ref_label.bind("<ButtonRelease-1>", self._end_internal_drag)
        ref_label.bind("<Button-3>", right_click_handler)
        ref_label.bind("<MouseWheel>", self._on_mousewheel)

        # Store widget references (no StringVar needed - update label directly)
        self.field_widgets[field_id] = {
            "frame": field_frame,
            "order_label": order_label,
            "cat_pos_label": cat_pos_label,
            "drag_handle": drag_handle,
            "display_label": display_label,
            "ref_label": ref_label,
            "field_item": field_item
        }
    
    def _cleanup_widget_bindings(self, widget):
        """Recursively unbind all events from widget and children to prevent memory leaks"""
        try:
            # Unbind common events
            for event in ('<Button-1>', '<B1-Motion>', '<ButtonRelease-1>',
                         '<Button-3>', '<MouseWheel>', '<Double-Button-1>',
                         '<Configure>', '<<ComboboxSelected>>'):
                try:
                    widget.unbind(event)
                except:
                    pass
            # Recursively clean children
            for child in widget.winfo_children():
                self._cleanup_widget_bindings(child)
        except:
            pass  # Widget may already be destroyed

    def clear(self):
        """Clear all fields and reset filter"""
        for widgets in list(self.field_widgets.values()):
            # Unbind events before destroying to prevent memory leaks
            self._cleanup_widget_bindings(widgets["frame"])
            widgets["frame"].destroy()

        self.field_widgets.clear()
        self.field_items.clear()
        self._update_empty_state()

        # Reset selection tracking
        if hasattr(self, 'selected_items'):
            self.selected_items.clear()
        if hasattr(self, '_previous_selected'):
            self._previous_selected.clear()

        # Reset virtual scrolling state
        self._virtual_mode = False
        self._visible_range = (0, 0)
        self._filtered_items = []
        if self._scroll_job:
            self.after_cancel(self._scroll_job)
            self._scroll_job = None

        # Reset grouped view state
        self._clear_grouped_view()
        self._view_mode = "flat"
        self.view_mode_var.set("flat")
        self._group_by_category_idx = 0
        self._combo_to_level_idx = []
        self._collapsed_groups.clear()
        self._group_order.clear()
        # Hide grouped controls children (frame stays packed to maintain width)
        self.group_by_combo.pack_forget()
        self._align_btn_frame.pack_forget()

        # Reset filter
        self._category_filters = None
        self._category_levels = []
        # Reset the hierarchical filter dropdown
        if hasattr(self, '_filter_dropdown'):
            self._filter_dropdown.set_category_levels([])
        # Filter frame stays visible - dropdown contents are cleared

    def set_enabled(self, enabled: bool):
        """Enable/disable panel"""
        # Individual widgets handle their own state
        pass

    def _set_bottom_toolbar_enabled(self, enabled: bool):
        """Enable/disable bottom toolbar based on parameter state"""
        # Enable/disable buttons
        if hasattr(self, 'edit_categories_btn'):
            self.edit_categories_btn.set_enabled(enabled)
        if hasattr(self, 'revert_names_btn'):
            self.revert_names_btn.set_enabled(enabled)

        # Handle radio group special disabled state
        self._set_radio_group_enabled(enabled)

    def _set_radio_group_enabled(self, enabled: bool):
        """Enable/disable radio group with proper visual state"""
        if not hasattr(self, 'view_mode_radio_group'):
            return

        if enabled:
            self.view_mode_radio_group.set_enabled(True)
        else:
            # Both radios should appear unselected when disabled
            self.view_mode_radio_group.set_enabled(False)
            # Force both to show as unselected (empty circles)
            if hasattr(self.view_mode_radio_group, '_radios'):
                for radio in self.view_mode_radio_group._radios.values():
                    if hasattr(radio, '_draw_radio'):
                        radio._draw_radio(selected=False, enabled=False)

    def lock_width(self, width: int = None):
        """
        Lock the panel width to prevent size changes when switching view modes.

        Args:
            width: Explicit width to lock to. If None, uses current width.
        """
        if width is None:
            width = self.winfo_width()

        if width > 10:  # Only lock if we have a valid width
            # Configure the LabelFrame with explicit width
            self.configure(width=width)
            # Prevent children from changing the frame size (both pack and grid)
            self.pack_propagate(False)
            self.grid_propagate(False)
            # Also lock content wrapper if it exists
            if hasattr(self, '_content_wrapper'):
                self._content_wrapper.configure(width=width - 30)  # Account for padding
                self._content_wrapper.pack_propagate(False)
            # Lock canvas frame to prevent content-driven width changes
            if hasattr(self, '_canvas_frame'):
                self._canvas_frame.pack_propagate(False)
            self._locked_width = width
            self.logger.debug(f"Builder panel width locked to {width}")

    def on_theme_changed(self):
        """Update widget colors when theme changes"""
        colors = self._theme_manager.colors
        # Use 'background' for inner content areas (pure white in light mode)
        bg_color = colors['background']
        fg_color = colors['text_primary']
        # Theme-aware disabled colors for buttons
        is_dark = colors.get('background', '') == '#0d0d1a'
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Update RoundedButton widgets (secondary style)
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

        # Update wrapper frames for buttons
        if hasattr(self, '_left_btn_frame'):
            self._left_btn_frame.config(bg=bg_color)
        if hasattr(self, '_grouped_controls_frame'):
            self._grouped_controls_frame.config(bg=bg_color)
        if hasattr(self, '_align_btn_frame'):
            self._align_btn_frame.config(bg=bg_color)
        if hasattr(self, '_clear_filter_wrapper'):
            self._clear_filter_wrapper.config(bg=bg_color)
        if hasattr(self, '_arrow_frame'):
            self._arrow_frame.config(bg=bg_color)
        if hasattr(self, '_header_group_by_frame'):
            self._header_group_by_frame.config(bg=bg_color)

        # Update view mode frames and label
        self.view_mode_frame.config(bg=bg_color)
        self.view_toggle_frame.config(bg=bg_color)
        self.view_label.config(bg=bg_color, fg=fg_color)

        # Update SVG radio group
        self.view_mode_radio_group.on_theme_changed()

        # Update inner frames
        if hasattr(self, '_inner_frames'):
            for frame in self._inner_frames:
                frame.config(bg=bg_color)

        # Update canvas frame border and background
        if hasattr(self, '_canvas_frame'):
            border_color = colors.get('border', '#3a3a4a')
            self._canvas_frame.config(
                bg=bg_color,
                highlightbackground=border_color,
                highlightcolor=border_color
            )

        # Update bottom toolbar frame and children
        if hasattr(self, 'toolbar_bottom'):
            self.toolbar_bottom.config(bg=bg_color)

        # Update revert names button (text button in bottom toolbar)
        if hasattr(self, 'revert_names_btn'):
            self.revert_names_btn.update_colors(
                bg=colors['button_secondary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                fg=colors['text_primary'],
                disabled_bg=disabled_bg,
                disabled_fg=disabled_fg,
                canvas_bg=bg_color
            )

        # Update delete button if icon button
        if hasattr(self, 'delete_btn') and hasattr(self.delete_btn, 'on_theme_changed'):
            self.delete_btn.on_theme_changed()

        # Update filter dropdown and count label
        if hasattr(self, '_filter_dropdown'):
            self._filter_dropdown.on_theme_changed()
        if hasattr(self, '_filter_label'):
            self._filter_label.config(bg=bg_color, fg=fg_color)
        self.filter_count_label.config(bg=bg_color, fg=colors['text_muted'])

        # Update icon buttons
        if hasattr(self, '_icon_buttons'):
            for btn in self._icon_buttons:
                if hasattr(btn, 'on_theme_changed'):
                    btn.on_theme_changed()

        # Update canvas and fields_container background (use darkest so row colors stand out)
        canvas_bg = colors['background']
        if hasattr(self, 'canvas'):
            self.canvas.config(bg=canvas_bg)
        if hasattr(self, 'fields_container'):
            self.fields_container.config(bg=canvas_bg)

        # Update ThemedScrollbar
        if hasattr(self, '_canvas_scrollbar'):
            self._canvas_scrollbar.on_theme_changed()

        # Update empty text if visible
        if hasattr(self, '_empty_text_id') and self._empty_text_id:
            self.canvas.itemconfig(self._empty_text_id, fill=colors['text_muted'])

        # Update drop indicator if it exists
        if self._drop_indicator:
            self._drop_indicator.config(bg=colors['button_primary'])

        # Update group drop indicator if it exists
        if self._group_drop_indicator:
            self._group_drop_indicator.config(bg=colors['button_primary'])

        # Update section header widgets - match BaseToolTab.on_theme_changed pattern exactly
        header_bg = colors.get('section_bg', colors['background'])
        for header_frame, icon_label, text_label in self._section_header_widgets:
            try:
                header_frame.configure(bg=header_bg)
                if icon_label:
                    icon_label.configure(bg=header_bg)
                text_label.configure(bg=header_bg, fg=colors['title_color'])
            except Exception:
                pass

        # Force the LabelFrame to re-apply its style
        try:
            self.configure(style='Section.TLabelframe')
        except Exception:
            pass

        # Update field row widgets (tk.Frame and tk.Label don't auto-update like ttk)
        row_bg = colors['card_surface']
        for field_id, widgets in self.field_widgets.items():
            try:
                frame = widgets.get('frame')
                if frame and frame.winfo_exists():
                    frame.config(bg=row_bg)
                # Update individual label widgets by their stored references
                for key in ['order_label', 'drag_handle']:
                    label = widgets.get(key)
                    if label and label.winfo_exists():
                        label.config(bg=row_bg, fg=colors['text_primary'] if key == 'order_label' else colors['text_secondary'])
                for key in ['cat_pos_label', 'ref_label']:
                    label = widgets.get(key)
                    if label and label.winfo_exists():
                        label.config(bg=row_bg, fg=colors['text_muted'])
                display = widgets.get('display_label')
                if display and display.winfo_exists():
                    display.config(bg=row_bg, fg=colors['title_color'])
            except Exception:
                pass  # Widget may have been destroyed

        # Update group header widgets (grouped view)
        # Use background for container (darkest) so header row stands out
        group_bg = colors['background']
        header_bg = colors['surface']
        for group_label, widgets in self._group_widgets.items():
            try:
                if widgets.get('group_container') and widgets['group_container'].winfo_exists():
                    widgets['group_container'].config(bg=group_bg)
                if widgets.get('header_frame') and widgets['header_frame'].winfo_exists():
                    widgets['header_frame'].config(bg=header_bg)
                for key in ['drag_handle', 'position_label', 'name_label', 'count_label']:
                    widget = widgets.get(key)
                    if widget and widget.winfo_exists():
                        widget.config(bg=header_bg, fg=colors['text_primary'])
            except Exception:
                pass  # Widget may have been destroyed

        # Refresh selection highlighting with new theme colors (force update all for theme change)
        self._update_selection_visuals(force_update=True)
        self._update_group_selection_visuals()

        # Update bottom spacer if it exists (grouped view only)
        if hasattr(self, '_bottom_spacer') and self._bottom_spacer and self._bottom_spacer.winfo_exists():
            self._bottom_spacer.config(bg=colors['background'])

        # Force ttk.Frame widgets to re-apply their style after theme change
        # ttk widgets cache their appearance and don't auto-update when style changes
        if hasattr(self, '_content_wrapper'):
            self._content_wrapper.configure(style='Section.TFrame')

    def _create_section_header(self, parent, text: str, icon_name: str):
        """Create a section header with icon and set it as the labelwidget"""
        colors = self._theme_manager.colors
        # Match BaseToolTab.create_section_header pattern - use section_bg
        bg_color = colors.get('section_bg', colors['background'])

        # Create header frame as child of parent (not self)
        header_frame = tk.Frame(parent, bg=bg_color)

        # Load icon
        icon = self._load_icon(icon_name, size=16)
        icon_label = None
        if icon:
            icon_label = tk.Label(header_frame, image=icon, bg=bg_color)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            # Store reference to prevent garbage collection
            icon_label._icon_ref = icon
            self._header_icon = icon

        # Text label with title styling
        text_label = tk.Label(
            header_frame, text=text, bg=bg_color,
            fg=colors['title_color'], font=('Segoe UI Semibold', 11)
        )
        text_label.pack(side=tk.LEFT)

        # Store for theme updates
        self._section_header_widgets.append((header_frame, icon_label, text_label))

        # Configure this LabelFrame to use the header widget
        self.configure(labelwidget=header_frame)

    def _load_icon(self, icon_name: str, size: int = 16):
        """Load an SVG icon for section header"""
        if not PIL_AVAILABLE:
            return None

        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"
        svg_path = icons_dir / f"{icon_name}.svg"
        png_path = icons_dir / f"{icon_name}.png"

        try:
            img = None

            # Try SVG first (if cairosvg available)
            if CAIROSVG_AVAILABLE and svg_path.exists():
                # Render at 4x size for quality, then downscale
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=size * 4,
                    output_height=size * 4
                )
                img = Image.open(io.BytesIO(png_data))
            elif png_path.exists():
                img = Image.open(png_path)

            if img is None:
                return None

            # Resize to target size with high-quality resampling
            img = img.resize((size, size), Image.Resampling.LANCZOS)

            return ImageTk.PhotoImage(img)
        except Exception as e:
            self.logger.debug(f"Failed to load icon {icon_name}: {e}")
            return None


# NOTE: AddLabelDialog and CategoryLabelEditorDialog moved to dialogs/


# NOTE: FieldCategoryEditorDialog moved to dialogs/field_category_editor.py



# NOTE: CategoryManagerPanel moved to panels/category_manager.py
