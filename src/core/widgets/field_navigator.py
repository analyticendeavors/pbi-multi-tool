"""
FieldNavigator - Reusable Field Selection Component
Built by Reid Havens of Analytic Endeavors

A shared widget for navigating and selecting Power BI model fields (measures and columns)
with search, filtering, context menu, and drag-and-drop support.

Usage:
    from core.widgets import FieldNavigator

    navigator = FieldNavigator(
        parent=self.frame,
        theme_manager=self._theme_manager,
        on_fields_selected=self.handle_selection,
        drop_target=self.builder_panel,  # Optional
    )
    navigator.pack(fill=tk.BOTH, expand=True)
    navigator.set_fields(tables_data)
"""

import tkinter as tk
from tkinter import ttk, simpledialog
from typing import Dict, List, Optional, Tuple, Callable, Protocol, runtime_checkable
import logging
from pathlib import Path
import io

from core.theme_manager import ThemeManager
from core.pbi_connector import FieldInfo, TableFieldsInfo
from core.ui_base import (
    RoundedButton, ThemedScrollbar, ThemedMessageBox,
    LabeledRadioGroup, ThemedContextMenu, SectionPanelMixin
)

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


@runtime_checkable
class DropTargetProtocol(Protocol):
    """Interface for widgets that can receive dropped fields."""

    def show_drop_indicator(self, screen_y: int) -> None:
        """Show visual drop indicator at the given screen Y position."""
        ...

    def hide_drop_indicator(self) -> None:
        """Hide the drop indicator."""
        ...

    def get_drop_position(self) -> Optional[int]:
        """Get the position where items should be inserted."""
        ...

    def winfo_rootx(self) -> int:
        """Get absolute X position of widget."""
        ...

    def winfo_rooty(self) -> int:
        """Get absolute Y position of widget."""
        ...

    def winfo_width(self) -> int:
        """Get widget width."""
        ...

    def winfo_height(self) -> int:
        """Get widget height."""
        ...


class FieldNavigator(SectionPanelMixin, ttk.LabelFrame):
    """
    Reusable field navigation component with search, filtering, context menu,
    drag-and-drop, and multi-select capabilities.

    Args:
        parent: Parent widget
        theme_manager: ThemeManager instance for theming
        on_fields_selected: Callback when fields are added - receives (fields: List[FieldInfo], position: Optional[int])
        drop_target: Optional widget implementing DropTargetProtocol for drag-drop support
        section_title: Header title (default: "Available Fields")
        section_icon: Icon name for header (default: "table")
        show_columns: Include columns in tree (default: True)
        show_add_button: Show "Add Selected" button (default: True)
        duplicate_checker: Optional callback to check if field is duplicate - receives (table_name, field_name)
        show_duplicate_dialogs: Show confirmation dialogs for duplicates (default: True)
        placeholder_text: Text to show when tree is empty
        can_add_validator: Optional callback to validate if adding fields is allowed - returns True if OK, False otherwise
    """

    def __init__(
        self,
        parent,
        theme_manager: ThemeManager,
        on_fields_selected: Callable[[List[FieldInfo], Optional[int]], None],
        drop_target: Optional[DropTargetProtocol] = None,
        section_title: str = "Available Fields",
        section_icon: str = "table",
        show_columns: bool = True,
        show_add_button: bool = True,
        duplicate_checker: Optional[Callable[[str, str], bool]] = None,
        show_duplicate_dialogs: bool = True,
        placeholder_text: str = "No tables or fields loaded.\n\nConnect to a model to view\navailable tables and fields.",
        can_add_validator: Optional[Callable[[], bool]] = None,
    ):
        ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe', padding="12")
        self._init_section_panel_mixin()

        # Store configuration
        self._theme_manager = theme_manager
        self._on_fields_selected = on_fields_selected
        self._drop_target = drop_target
        self._show_columns = show_columns
        self._show_add_button = show_add_button
        self._duplicate_checker = duplicate_checker
        self._show_duplicate_dialogs = show_duplicate_dialogs
        self._placeholder_text = placeholder_text
        self._can_add_validator = can_add_validator

        self.logger = logging.getLogger(__name__)

        # Button tracking for theme updates
        self._primary_buttons = []
        self._secondary_buttons = []

        # Tree icons (loaded during setup_ui)
        self._tree_icons = {}

        # Create section header
        self._create_section_header(parent, section_title, section_icon)

        # Store table metadata for field lookups
        self.tables_data: Dict[str, TableFieldsInfo] = {}

        self._setup_ui()

    def _setup_ui(self):
        """Setup the field navigator UI"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        content_bg = colors['background']
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Inner wrapper with Section.TFrame style
        self._content_wrapper = ttk.Frame(self, style='Section.TFrame', padding="15")
        self._content_wrapper.pack(fill=tk.BOTH, expand=True)

        # Track inner frames for theme updates
        self._inner_frames = []

        # Load tree icons
        self._load_tree_icons()

        # Search box
        search_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        search_frame.pack(fill=tk.X, pady=(0, 8))
        self._inner_frames.append(search_frame)
        self._search_frame = search_frame

        # Search icon
        search_icon = self._tree_icons.get('search')
        if search_icon:
            self._search_icon_label = tk.Label(search_frame, image=search_icon, bg=content_bg)
            self._search_icon_label.pack(side=tk.LEFT, padx=(0, 2))
        else:
            self._search_icon_label = tk.Label(
                search_frame, text="Q", font=("Segoe UI", 9),
                bg=content_bg, fg=colors['text_muted']
            )
            self._search_icon_label.pack(side=tk.LEFT, padx=(0, 2))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, style='Compact.TEntry')
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # Type filter (All / Measures / Columns)
        fg_color = colors['text_primary']
        self.filter_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        self.filter_frame.pack(fill=tk.X, pady=(0, 8))
        self._inner_frames.append(self.filter_frame)

        self.filter_label = tk.Label(self.filter_frame, text="Show:", font=("Segoe UI", 8),
                                     bg=content_bg, fg=fg_color)
        self.filter_label.pack(side=tk.LEFT, padx=(0, 5))

        self.type_filter_var = tk.StringVar(value="all")

        # Build filter options based on show_columns setting
        filter_options = [("all", "All"), ("measures", "Measures")]
        if self._show_columns:
            filter_options.append(("columns", "Columns"))

        self.type_filter_radio_group = LabeledRadioGroup(
            self.filter_frame,
            variable=self.type_filter_var,
            options=filter_options,
            command=self._on_type_filter_changed,
            orientation="horizontal",
            font=("Segoe UI", 8),
            padding=8
        )
        self.type_filter_radio_group.pack(side=tk.LEFT)

        # Tree view with scrollbar
        tree_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self._inner_frames.append(tree_frame)
        self._tree_frame = tree_frame

        tree_border = colors.get('tree_border', colors.get('border', colors['background']))
        tree_bg = colors.get('section_bg', colors.get('surface', content_bg))

        tree_container = tk.Frame(tree_frame, bg=tree_bg,
                                  highlightbackground=tree_border, highlightcolor=tree_border,
                                  highlightthickness=1)
        tree_container.pack(fill=tk.BOTH, expand=True)
        self._tree_container = tree_container
        self._tree_border_color = tree_border

        self._tree_scrollbar = ThemedScrollbar(
            tree_container,
            command=None,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            tree_container,
            columns=("type", "data1", "data2"),
            displaycolumns=(),
            yscrollcommand=self._tree_scrollbar.set,
            selectmode="extended",
            show='tree'
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(2, 4))
        self._tree_scrollbar._command = self.tree.yview

        self.tree.heading("#0", text="")
        self.tree.column("#0", width=220, minwidth=150, stretch=True)

        self._configure_treeview_style()

        # Bind events
        self.tree.bind("<Double-Button-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)

        # Drag and drop
        self.tree.bind("<ButtonPress-1>", self._on_drag_start)
        self.tree.bind("<B1-Motion>", self._on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_drag_end)

        # Placeholder overlay
        self._placeholder_label = tk.Label(
            tree_container,
            text=self._placeholder_text,
            font=("Segoe UI", 10, "italic"),
            fg=colors['text_muted'],
            bg=tree_bg,
            justify="center"
        )
        self._placeholder_label.place(relx=0.5, rely=0.5, anchor="center")
        self._placeholder_label.lower()

        # Add button (optional)
        if self._show_add_button:
            btn_frame = tk.Frame(self._content_wrapper, bg=content_bg)
            btn_frame.pack(fill=tk.X, pady=(8, 0))
            self._inner_frames.append(btn_frame)

            self.add_btn = RoundedButton(
                btn_frame,
                text="Add Selected Field(s)",
                command=self._on_add_clicked,
                bg=colors['button_secondary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                fg=colors['text_primary'],
                disabled_bg=disabled_bg,
                disabled_fg=disabled_fg,
                width=None, height=32, radius=6,
                font=('Segoe UI', 9),
                canvas_bg=content_bg
            )
            self.add_btn.pack(expand=True)
            self._secondary_buttons.append(self.add_btn)
            self._btn_frame = btn_frame

        self._update_empty_state()

    # =========================================================================
    # PUBLIC API
    # =========================================================================

    def set_fields(self, tables: Dict[str, TableFieldsInfo]) -> None:
        """Update the navigator with new field data."""
        self.tree.delete(*self.tree.get_children())
        self.tables_data = tables

        sorted_tables = sorted(
            tables.items(),
            key=lambda x: (x[1].sort_priority, x[0].lower())
        )

        for table_name, table_info in sorted_tables:
            table_display = table_name
            if table_info.is_measures_only:
                table_display += " (Measures)"

            table_icon = self._tree_icons.get('table')
            if table_icon:
                table_id = self.tree.insert(
                    "", "end",
                    text=f" {table_display}",
                    image=table_icon,
                    values=("table", table_name, ""),
                    tags=("table",)
                )
            else:
                table_id = self.tree.insert(
                    "", "end",
                    text=f"\U0001F5C3\uFE0F {table_display}",
                    values=("table", table_name, ""),
                    tags=("table",)
                )

            self._build_folder_hierarchy(table_id, table_name, table_info.fields)

        self._update_empty_state()

    def clear(self) -> None:
        """Clear all fields from the navigator."""
        self.tree.delete(*self.tree.get_children())
        self.search_var.set("")
        self.tables_data = {}
        self._update_empty_state()

    def set_duplicate_checker(self, checker: Callable[[str, str], bool]) -> None:
        """Set callback to check if a field is already selected."""
        self._duplicate_checker = checker

    def get_selected_fields(self) -> List[FieldInfo]:
        """Get currently selected fields in the tree."""
        fields = []
        for item_id in self.tree.selection():
            values = self.tree.item(item_id, "values")
            if values and values[0] == "field":
                table_name = values[1]
                field_name = values[2]
                table_info = self.tables_data.get(table_name)
                if table_info:
                    for field_info in table_info.fields:
                        if field_info.name == field_name:
                            fields.append(field_info)
                            break
        return fields

    def set_filter(self, filter_value: str) -> None:
        """Programmatically set the type filter (all/measures/columns)."""
        self.type_filter_var.set(filter_value)
        self._apply_all_filters()

    def set_search(self, search_text: str) -> None:
        """Programmatically set the search text."""
        self.search_var.set(search_text)

    def set_enabled(self, enabled: bool) -> None:
        """Enable/disable panel controls."""
        if self._show_add_button and hasattr(self, 'add_btn'):
            self.add_btn.set_enabled(enabled)

    def set_drop_target(self, drop_target: Optional[DropTargetProtocol]) -> None:
        """
        Set or update the drop target for drag-and-drop operations.

        Use this when the drop target widget is created after the FieldNavigator.

        Args:
            drop_target: Widget implementing DropTargetProtocol, or None to disable drag-drop
        """
        self._drop_target = drop_target

    # =========================================================================
    # TREE BUILDING
    # =========================================================================

    def _build_folder_hierarchy(self, table_id: str, table_name: str, fields: List[FieldInfo]):
        """Build hierarchical folder tree under a table node."""
        folder_nodes = {}
        fields_by_folder = {}

        for field_info in fields:
            # Skip columns if not showing them
            if not self._show_columns and field_info.field_type != "Measure":
                continue

            display_folder = field_info.display_folder.strip()

            if not display_folder:
                if "" not in fields_by_folder:
                    fields_by_folder[""] = []
                fields_by_folder[""].append(field_info)
            else:
                locations = [loc.strip() for loc in display_folder.split(';')]
                for location in locations:
                    if location:
                        if location not in fields_by_folder:
                            fields_by_folder[location] = []
                        fields_by_folder[location].append(field_info)

        # Create folder nodes
        all_folders = set()
        for folder_path in fields_by_folder.keys():
            if folder_path:
                parts = folder_path.split('\\')
                for i in range(len(parts)):
                    cumulative_path = '\\'.join(parts[:i+1])
                    all_folders.add(cumulative_path)

        sorted_folders = sorted(all_folders, key=lambda x: (x.count('\\'), x.lower()))

        for folder_path in sorted_folders:
            if folder_path in folder_nodes:
                continue

            parts = folder_path.split('\\')
            folder_name = parts[-1]

            if len(parts) == 1:
                parent_id = table_id
            else:
                parent_path = '\\'.join(parts[:-1])
                parent_id = folder_nodes.get(parent_path, table_id)

            folder_icon = self._tree_icons.get('folder')
            if folder_icon:
                folder_id = self.tree.insert(
                    parent_id, "end",
                    text=f" {folder_name}",
                    image=folder_icon,
                    values=("folder", table_name, folder_path),
                    tags=("folder",)
                )
            else:
                folder_id = self.tree.insert(
                    parent_id, "end",
                    text=f"\U0001F4C2 {folder_name}",
                    values=("folder", table_name, folder_path),
                    tags=("folder",)
                )
            folder_nodes[folder_path] = folder_id

        # Add fields to folders
        for folder_path, field_list in sorted(fields_by_folder.items()):
            if folder_path == "":
                parent_id = table_id
            else:
                parent_id = folder_nodes.get(folder_path, table_id)

            for field_info in sorted(field_list, key=lambda f: f.name.lower()):
                if field_info.field_type == "Measure":
                    field_icon = self._tree_icons.get('measure')
                    fallback_icon = "\U0001F9EE"
                else:
                    field_icon = self._tree_icons.get('column')
                    fallback_icon = "\u25AD"

                if field_icon:
                    self.tree.insert(
                        parent_id, "end",
                        text=f" {field_info.name}",
                        image=field_icon,
                        values=("field", table_name, field_info.name),
                        tags=("field",)
                    )
                else:
                    self.tree.insert(
                        parent_id, "end",
                        text=f"{fallback_icon} {field_info.name}",
                        values=("field", table_name, field_info.name),
                        tags=("field",)
                    )

    # =========================================================================
    # FILTERING
    # =========================================================================

    def _on_type_filter_changed(self, *args):
        """Handle type filter change."""
        self._apply_all_filters()

    def _on_search(self, *args):
        """Filter tree based on search text."""
        self._apply_all_filters()

    def _apply_all_filters(self):
        """Apply both search and type filters."""
        search_text = self.search_var.get().lower().strip()
        type_filter = self.type_filter_var.get()
        self._rebuild_tree_with_filters(search_text, type_filter)

    def _rebuild_tree_with_filters(self, search_text: str, type_filter: str):
        """Rebuild tree showing only items matching both filters."""
        if not self.tables_data:
            return

        self.tree.delete(*self.tree.get_children())

        sorted_tables = sorted(
            self.tables_data.items(),
            key=lambda x: (x[1].sort_priority, x[0].lower())
        )

        for table_name, table_info in sorted_tables:
            filtered_fields = []
            for field_info in table_info.fields:
                # Skip columns if not showing them
                if not self._show_columns and field_info.field_type != "Measure":
                    continue

                # Apply type filter
                if type_filter == "measures" and field_info.field_type != "Measure":
                    continue
                if type_filter == "columns" and field_info.field_type == "Measure":
                    continue

                # Apply search filter
                if search_text:
                    if search_text not in field_info.name.lower() and search_text not in table_name.lower():
                        continue

                filtered_fields.append(field_info)

            if not filtered_fields and not (search_text and search_text in table_name.lower()):
                continue

            table_display = table_name
            if table_info.is_measures_only:
                table_display += " (Measures)"

            table_icon = self._tree_icons.get('table')
            if table_icon:
                table_id = self.tree.insert(
                    "", "end",
                    text=f" {table_display}",
                    image=table_icon,
                    values=("table", table_name, ""),
                    tags=("table",)
                )
            else:
                table_id = self.tree.insert(
                    "", "end",
                    text=f"\U0001F5C3\uFE0F {table_display}",
                    values=("table", table_name, ""),
                    tags=("table",)
                )

            self._build_filtered_folder_hierarchy(table_id, table_name, filtered_fields, search_text)

            if search_text:
                self.tree.item(table_id, open=True)

    def _build_filtered_folder_hierarchy(self, table_id: str, table_name: str, fields: List[FieldInfo], search_text: str):
        """Build folder hierarchy for filtered fields."""
        folder_nodes = {}
        fields_by_folder = {}

        for field_info in fields:
            display_folder = field_info.display_folder.strip()

            if not display_folder:
                if "" not in fields_by_folder:
                    fields_by_folder[""] = []
                fields_by_folder[""].append(field_info)
            else:
                locations = [loc.strip() for loc in display_folder.split(';')]
                for location in locations:
                    if location:
                        if location not in fields_by_folder:
                            fields_by_folder[location] = []
                        fields_by_folder[location].append(field_info)

        all_folders = set()
        for folder_path in fields_by_folder.keys():
            if folder_path:
                parts = folder_path.split('\\')
                for i in range(len(parts)):
                    cumulative_path = '\\'.join(parts[:i+1])
                    all_folders.add(cumulative_path)

        sorted_folders = sorted(all_folders, key=lambda x: (x.count('\\'), x.lower()))

        for folder_path in sorted_folders:
            if folder_path in folder_nodes:
                continue

            parts = folder_path.split('\\')
            folder_name = parts[-1]

            if len(parts) == 1:
                parent_id = table_id
            else:
                parent_path = '\\'.join(parts[:-1])
                parent_id = folder_nodes.get(parent_path, table_id)

            folder_icon = self._tree_icons.get('folder')
            if folder_icon:
                folder_id = self.tree.insert(
                    parent_id, "end",
                    text=f" {folder_name}",
                    image=folder_icon,
                    values=("folder", table_name, folder_path),
                    tags=("folder",)
                )
            else:
                folder_id = self.tree.insert(
                    parent_id, "end",
                    text=f"\U0001F4C2 {folder_name}",
                    values=("folder", table_name, folder_path),
                    tags=("folder",)
                )
            folder_nodes[folder_path] = folder_id

            if search_text:
                self.tree.item(folder_id, open=True)

        for folder_path, field_list in sorted(fields_by_folder.items()):
            if folder_path == "":
                parent_id = table_id
            else:
                parent_id = folder_nodes.get(folder_path, table_id)

            for field_info in sorted(field_list, key=lambda f: f.name.lower()):
                if field_info.field_type == "Measure":
                    field_icon = self._tree_icons.get('measure')
                    fallback_icon = "\U0001F9EE"
                else:
                    field_icon = self._tree_icons.get('column')
                    fallback_icon = "\u25AD"

                if field_icon:
                    self.tree.insert(
                        parent_id, "end",
                        text=f" {field_info.name}",
                        image=field_icon,
                        values=("field", table_name, field_info.name),
                        tags=("field",)
                    )
                else:
                    self.tree.insert(
                        parent_id, "end",
                        text=f"{fallback_icon} {field_info.name}",
                        values=("field", table_name, field_info.name),
                        tags=("field",)
                    )

    # =========================================================================
    # HELPER METHODS
    # =========================================================================

    def _get_fields_from_item(self, item_id: str) -> List[Tuple[str, str]]:
        """Recursively collect all fields under an item."""
        fields = []
        values = self.tree.item(item_id, "values")
        item_type = values[0] if values else None

        if item_type == "field":
            table_name = values[1]
            field_name = values[2]
            fields.append((table_name, field_name))
        elif item_type in ("table", "folder"):
            for child_id in self.tree.get_children(item_id):
                fields.extend(self._get_fields_from_item(child_id))

        return fields

    def _count_field_types(self, fields: List[Tuple[str, str]]) -> Tuple[int, int]:
        """Count measures vs columns. Returns (num_measures, num_columns)."""
        num_measures = 0
        num_columns = 0

        for table_name, field_name in fields:
            table_info = self.tables_data.get(table_name)
            if table_info:
                for field_info in table_info.fields:
                    if field_info.name == field_name:
                        if field_info.field_type == "Measure":
                            num_measures += 1
                        else:
                            num_columns += 1
                        break

        return num_measures, num_columns

    def _get_field_info(self, table_name: str, field_name: str) -> Optional[FieldInfo]:
        """Get FieldInfo object for a table/field combination."""
        table_info = self.tables_data.get(table_name)
        if table_info:
            for field_info in table_info.fields:
                if field_info.name == field_name:
                    return field_info
        return None

    def _fields_to_field_info_list(self, fields: List[Tuple[str, str]]) -> List[FieldInfo]:
        """Convert list of (table, field) tuples to FieldInfo objects."""
        result = []
        for table_name, field_name in fields:
            field_info = self._get_field_info(table_name, field_name)
            if field_info:
                result.append(field_info)
        return result

    # =========================================================================
    # VALIDATION
    # =========================================================================

    def _validate_can_add(self) -> bool:
        """
        Validate that fields can be added.
        Returns True if OK to proceed, False if adding should be blocked.
        """
        if self._can_add_validator:
            return self._can_add_validator()
        return True

    # =========================================================================
    # DUPLICATE DETECTION
    # =========================================================================

    def _is_field_duplicate(self, table_name: str, field_name: str) -> bool:
        """Check if a field is already selected."""
        if self._duplicate_checker:
            return self._duplicate_checker(table_name, field_name)
        return False

    def _get_duplicates_from_list(self, fields: List[Tuple[str, str]]) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
        """Check which fields are duplicates. Returns (duplicates, new_fields)."""
        duplicates = []
        new_fields = []

        for table_name, field_name in fields:
            if self._is_field_duplicate(table_name, field_name):
                duplicates.append((table_name, field_name))
            else:
                new_fields.append((table_name, field_name))

        return duplicates, new_fields

    def _show_bulk_duplicate_dialog(self, duplicates: List[Tuple[str, str]], new_fields: List[Tuple[str, str]],
                                     item_name: str) -> str:
        """Show dialog for bulk add with some duplicates. Returns 'all', 'new_only', or 'cancel'."""
        if not self._show_duplicate_dialogs:
            return 'all'

        dialog = tk.Toplevel(self)
        dialog.title("Duplicate Fields Found")
        dialog.geometry("430x350")
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_rootx() + 50
        y = self.winfo_rooty() + 50
        dialog.geometry(f"+{x}+{y}")

        result = {'choice': 'cancel'}

        msg_frame = ttk.Frame(dialog, padding=15)
        msg_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            msg_frame,
            text=f"Adding fields from '{item_name}'",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w")

        ttk.Label(
            msg_frame,
            text=f"{len(duplicates)} of {len(duplicates) + len(new_fields)} field(s) already exist:",
            wraplength=380
        ).pack(anchor="w", pady=(10, 5))

        list_frame = ttk.Frame(msg_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=8)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        for tbl, fld in duplicates:
            listbox.insert(tk.END, f"  ! {fld} (already exists)")
        for tbl, fld in new_fields:
            listbox.insert(tk.END, f"  + {fld} (new)")

        btn_frame = ttk.Frame(msg_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        def choose_all():
            result['choice'] = 'all'
            dialog.destroy()

        def choose_new_only():
            result['choice'] = 'new_only'
            dialog.destroy()

        def cancel():
            result['choice'] = 'cancel'
            dialog.destroy()

        ttk.Button(
            btn_frame,
            text=f"Add All {len(duplicates) + len(new_fields)} (including duplicates)",
            command=choose_all
        ).pack(side=tk.LEFT, padx=(0, 5))

        if new_fields:
            ttk.Button(
                btn_frame,
                text=f"Add Only {len(new_fields)} New",
                command=choose_new_only,
                style="Accent.TButton"
            ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            btn_frame,
            text="Cancel",
            command=cancel
        ).pack(side=tk.LEFT)

        dialog.wait_window()
        return result['choice']

    # =========================================================================
    # EVENT HANDLERS
    # =========================================================================

    def _on_double_click(self, event):
        """Handle double-click on tree item."""
        selection = self.tree.selection()
        if not selection:
            return

        item_id = selection[0]
        values = self.tree.item(item_id, "values")

        if not values:
            return

        item_type = values[0]

        if item_type == "field":
            table_name = values[1]
            field_name = values[2]

            # Validate before adding
            if not self._validate_can_add():
                return

            if self._is_field_duplicate(table_name, field_name):
                if self._show_duplicate_dialogs:
                    if ThemedMessageBox.askyesno(
                        self,
                        "Duplicate Field",
                        f"'{field_name}' is already in the parameter.\n\nAdd it again?"
                    ):
                        field_info = self._get_field_info(table_name, field_name)
                        if field_info:
                            self._on_fields_selected([field_info], None)
            else:
                field_info = self._get_field_info(table_name, field_name)
                if field_info:
                    self._on_fields_selected([field_info], None)

        elif item_type == "folder":
            fields = self._get_fields_from_item(item_id)

            # Filter to measures only
            measure_fields = []
            for tbl, fld in fields:
                table_info = self.tables_data.get(tbl)
                if table_info:
                    for field_info in table_info.fields:
                        if field_info.name == fld and field_info.field_type == "Measure":
                            measure_fields.append((tbl, fld))
                            break

            if not measure_fields:
                folder_path = values[2]
                ThemedMessageBox.showinfo(
                    self,
                    "No Measures",
                    f"Folder '{folder_path}' contains no measures."
                )
                return

            duplicates, new_fields = self._get_duplicates_from_list(measure_fields)

            if duplicates:
                folder_path = values[2]
                choice = self._show_bulk_duplicate_dialog(
                    duplicates, new_fields, f"folder '{folder_path}'"
                )

                if choice == 'cancel':
                    return
                elif choice == 'new_only':
                    fields_to_add = new_fields
                else:
                    fields_to_add = measure_fields
            else:
                folder_path = values[2]
                if self._show_duplicate_dialogs:
                    if not ThemedMessageBox.askyesno(
                        self,
                        "Add Folder Measures",
                        f"Add all {len(measure_fields)} measure(s) from folder '{folder_path}'?"
                    ):
                        return
                fields_to_add = measure_fields

            # Validate before adding
            if not self._validate_can_add():
                return

            field_infos = self._fields_to_field_info_list(fields_to_add)
            if field_infos:
                self._on_fields_selected(field_infos, None)

        elif item_type == "table":
            # Toggle expand/collapse
            if self.tree.item(item_id, 'open'):
                self.tree.item(item_id, open=False)
            else:
                self.tree.item(item_id, open=True)

    def _on_add_clicked(self):
        """Handle add button click."""
        self._add_selected_fields()

    def _add_selected_fields(self, position: Optional[int] = None):
        """Add selected field(s)."""
        selection = self.tree.selection()
        if not selection:
            return

        # Validate before adding
        if not self._validate_can_add():
            return

        fields_to_add = []
        for item in selection:
            values = self.tree.item(item, "values")
            if values and values[0] == "field":
                table_name = values[1]
                field_name = values[2]
                fields_to_add.append((table_name, field_name))
            elif values and values[0] in ("folder", "table"):
                folder_fields = self._get_fields_from_item(item)
                # Filter to measures only
                for tbl, fld in folder_fields:
                    table_info = self.tables_data.get(tbl)
                    if table_info:
                        for field_info in table_info.fields:
                            if field_info.name == fld and field_info.field_type == "Measure":
                                fields_to_add.append((tbl, fld))
                                break

        if not fields_to_add:
            ThemedMessageBox.showinfo(self, "No Fields", "Please select one or more fields (measures)")
            return

        duplicates, new_fields = self._get_duplicates_from_list(fields_to_add)

        if duplicates and len(fields_to_add) == 1:
            if self._show_duplicate_dialogs:
                if not ThemedMessageBox.askyesno(
                    self,
                    "Duplicate Field",
                    f"'{fields_to_add[0][1]}' is already in the parameter.\n\nAdd it again?"
                ):
                    return
            final_fields = fields_to_add
        elif duplicates:
            choice = self._show_bulk_duplicate_dialog(duplicates, new_fields, f"{len(fields_to_add)} field(s)")
            if choice == 'cancel':
                return
            elif choice == 'new_only':
                final_fields = new_fields
            else:
                final_fields = fields_to_add
        else:
            final_fields = fields_to_add

        field_infos = self._fields_to_field_info_list(final_fields)
        if field_infos:
            self._on_fields_selected(field_infos, position)

    # =========================================================================
    # CONTEXT MENU
    # =========================================================================

    def _on_right_click(self, event):
        """Handle right-click to show context menu."""
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        current_selection = self.tree.selection()

        if item_id not in current_selection:
            self.tree.selection_set(item_id)
            current_selection = (item_id,)

        values = self.tree.item(item_id, "values")
        if not values:
            return

        item_type = values[0]
        multi_select = len(current_selection) > 1

        menu = ThemedContextMenu(self, self._theme_manager)

        if multi_select:
            fields_to_add = []
            for sel_id in current_selection:
                sel_values = self.tree.item(sel_id, "values")
                if sel_values and sel_values[0] == "field":
                    fields_to_add.append((sel_values[1], sel_values[2]))

            if fields_to_add:
                count = len(fields_to_add)
                menu.add_command(
                    label=f"Add {count} Selected Field(s)",
                    command=lambda f=fields_to_add: self._context_add_multiple_fields(f, position=None)
                )
                menu.add_separator()
                menu.add_command(
                    label=f"Add {count} to Top of List",
                    command=lambda f=fields_to_add: self._context_add_multiple_fields(f, position=0)
                )
                menu.add_command(
                    label=f"Add {count} to Position...",
                    command=lambda f=fields_to_add: self._context_add_multiple_fields_at_position(f)
                )
                menu.add_command(
                    label=f"Add {count} to Bottom of List",
                    command=lambda f=fields_to_add: self._context_add_multiple_fields(f, position=-1)
                )
            else:
                menu.add_section_header("(No fields selected)")

        elif item_type == "field":
            table_name = values[1]
            field_name = values[2]

            menu.add_command(
                label=f"Add '{field_name}'",
                command=lambda: self._context_add_field(table_name, field_name, position=None)
            )
            menu.add_separator()
            menu.add_command(
                label="Add to Top of List",
                command=lambda: self._context_add_field(table_name, field_name, position=0)
            )
            menu.add_command(
                label="Add to Position...",
                command=lambda t=table_name, f=field_name: self._context_add_field_at_position(t, f)
            )
            menu.add_command(
                label="Add to Bottom of List",
                command=lambda: self._context_add_field(table_name, field_name, position=-1)
            )

        elif item_type == "folder":
            fields = self._get_fields_from_item(item_id)
            num_measures, num_columns = self._count_field_types(fields)

            if num_measures > 0 and num_columns > 0:
                menu.add_section_header("(Mixed measures & columns)")
                menu.add_section_header("Use table-level add instead")
            elif num_measures > 0:
                menu.add_command(
                    label=f"Add All {num_measures} Measure(s)",
                    command=lambda: self._context_add_folder_measures(item_id, position=None)
                )
                menu.add_separator()
                menu.add_command(
                    label="Add All to Top of List",
                    command=lambda: self._context_add_folder_measures(item_id, position=0)
                )
                menu.add_command(
                    label="Add All to Bottom of List",
                    command=lambda: self._context_add_folder_measures(item_id, position=-1)
                )
            elif num_columns > 0:
                menu.add_command(
                    label=f"Add All {num_columns} Column(s)",
                    command=lambda: self._context_add_folder_columns(item_id, position=None)
                )
                menu.add_separator()
                menu.add_command(
                    label="Add All to Top of List",
                    command=lambda: self._context_add_folder_columns(item_id, position=0)
                )
                menu.add_command(
                    label="Add All to Bottom of List",
                    command=lambda: self._context_add_folder_columns(item_id, position=-1)
                )
            else:
                menu.add_section_header("(No fields in folder)")

        elif item_type == "table":
            fields = self._get_fields_from_item(item_id)
            num_measures, num_columns = self._count_field_types(fields)

            if num_measures > 0:
                menu.add_command(
                    label=f"Add All {num_measures} Measure(s)",
                    command=lambda: self._context_add_table_measures(item_id, position=None)
                )

            if num_columns > 0:
                menu.add_command(
                    label=f"Add All {num_columns} Column(s)",
                    command=lambda: self._context_add_table_columns(item_id, position=None)
                )

            if num_measures > 0 or num_columns > 0:
                menu.add_separator()
                if num_measures > 0:
                    menu.add_command(
                        label="Add Measures to Top",
                        command=lambda: self._context_add_table_measures(item_id, position=0)
                    )
                    menu.add_command(
                        label="Add Measures to Bottom",
                        command=lambda: self._context_add_table_measures(item_id, position=-1)
                    )
                if num_columns > 0:
                    menu.add_command(
                        label="Add Columns to Top",
                        command=lambda: self._context_add_table_columns(item_id, position=0)
                    )
                    menu.add_command(
                        label="Add Columns to Bottom",
                        command=lambda: self._context_add_table_columns(item_id, position=-1)
                    )

            if num_measures == 0 and num_columns == 0:
                menu.add_section_header("(No fields in table)")

        menu.show(event.x_root, event.y_root)

    def _context_add_field(self, table_name: str, field_name: str, position: Optional[int] = None):
        """Context menu: Add single field at specified position."""
        # Validate before adding
        if not self._validate_can_add():
            return

        if self._is_field_duplicate(table_name, field_name):
            if self._show_duplicate_dialogs:
                if not ThemedMessageBox.askyesno(
                    self,
                    "Duplicate Field",
                    f"'{field_name}' is already in the parameter.\n\nAdd it again?"
                ):
                    return

        actual_position = None if position == -1 else position
        field_info = self._get_field_info(table_name, field_name)
        if field_info:
            self._on_fields_selected([field_info], actual_position)

    def _context_add_field_at_position(self, table_name: str, field_name: str):
        """Context menu: Add single field at user-specified position."""
        position_str = simpledialog.askstring(
            "Add to Position",
            "Enter position (1, 2, 3, ...):",
            initialvalue="1"
        )

        if not position_str:
            return

        try:
            position = int(position_str)
            if position < 1:
                position = 1
        except ValueError:
            ThemedMessageBox.showerror(self, "Invalid Position", "Please enter a valid number.")
            return

        self._context_add_field(table_name, field_name, position=position - 1)

    def _context_add_multiple_fields(self, fields: List[Tuple[str, str]], position: Optional[int] = None):
        """Context menu: Add multiple fields at specified position."""
        # Validate before adding
        if not self._validate_can_add():
            return

        duplicates, new_fields = self._get_duplicates_from_list(fields)

        if duplicates:
            choice = self._show_bulk_duplicate_dialog(
                duplicates, new_fields, f"{len(fields)} selected field(s)"
            )
            if choice == 'cancel':
                return
            elif choice == 'new_only':
                fields_to_add = new_fields
            else:
                fields_to_add = fields
        else:
            fields_to_add = fields

        if not fields_to_add:
            return

        actual_position = None if position == -1 else position
        field_infos = self._fields_to_field_info_list(fields_to_add)
        if field_infos:
            self._on_fields_selected(field_infos, actual_position)

    def _context_add_multiple_fields_at_position(self, fields: List[Tuple[str, str]]):
        """Context menu: Add multiple fields at user-specified position."""
        if not fields:
            return

        position_str = simpledialog.askstring(
            "Add to Position",
            f"Enter position for {len(fields)} field(s):",
            initialvalue="1"
        )

        if not position_str:
            return

        try:
            position = int(position_str)
            if position < 1:
                position = 1
        except ValueError:
            ThemedMessageBox.showerror(self, "Invalid Position", "Please enter a valid number.")
            return

        self._context_add_multiple_fields(fields, position=position - 1)

    def _context_add_folder_measures(self, item_id: str, position: Optional[int] = None):
        """Context menu: Add all measures from folder."""
        # Validate before adding
        if not self._validate_can_add():
            return

        fields = self._get_fields_from_item(item_id)

        measure_fields = []
        for tbl, fld in fields:
            table_info = self.tables_data.get(tbl)
            if table_info:
                for field_info in table_info.fields:
                    if field_info.name == fld and field_info.field_type == "Measure":
                        measure_fields.append((tbl, fld))
                        break

        if not measure_fields:
            return

        duplicates, new_fields = self._get_duplicates_from_list(measure_fields)

        if duplicates:
            folder_path = self.tree.item(item_id, "values")[2]
            choice = self._show_bulk_duplicate_dialog(
                duplicates, new_fields, f"folder '{folder_path}'"
            )
            if choice == 'cancel':
                return
            elif choice == 'new_only':
                fields_to_add = new_fields
            else:
                fields_to_add = measure_fields
        else:
            fields_to_add = measure_fields

        actual_position = None if position == -1 else position
        field_infos = self._fields_to_field_info_list(fields_to_add)
        if field_infos:
            self._on_fields_selected(field_infos, actual_position)

    def _context_add_folder_columns(self, item_id: str, position: Optional[int] = None):
        """Context menu: Add all columns from folder."""
        # Validate before adding
        if not self._validate_can_add():
            return

        fields = self._get_fields_from_item(item_id)

        column_fields = []
        for tbl, fld in fields:
            table_info = self.tables_data.get(tbl)
            if table_info:
                for field_info in table_info.fields:
                    if field_info.name == fld and field_info.field_type == "Column":
                        column_fields.append((tbl, fld))
                        break

        if not column_fields:
            return

        duplicates, new_fields = self._get_duplicates_from_list(column_fields)

        if duplicates:
            folder_path = self.tree.item(item_id, "values")[2]
            choice = self._show_bulk_duplicate_dialog(
                duplicates, new_fields, f"folder '{folder_path}'"
            )
            if choice == 'cancel':
                return
            elif choice == 'new_only':
                fields_to_add = new_fields
            else:
                fields_to_add = column_fields
        else:
            fields_to_add = column_fields

        actual_position = None if position == -1 else position
        field_infos = self._fields_to_field_info_list(fields_to_add)
        if field_infos:
            self._on_fields_selected(field_infos, actual_position)

    def _context_add_table_measures(self, item_id: str, position: Optional[int] = None):
        """Context menu: Add all measures from table."""
        # Validate before adding
        if not self._validate_can_add():
            return

        fields = self._get_fields_from_item(item_id)

        measure_fields = []
        for tbl, fld in fields:
            table_info = self.tables_data.get(tbl)
            if table_info:
                for field_info in table_info.fields:
                    if field_info.name == fld and field_info.field_type == "Measure":
                        measure_fields.append((tbl, fld))
                        break

        if not measure_fields:
            return

        duplicates, new_fields = self._get_duplicates_from_list(measure_fields)

        if duplicates:
            table_name = self.tree.item(item_id, "values")[1]
            choice = self._show_bulk_duplicate_dialog(
                duplicates, new_fields, f"table '{table_name}'"
            )
            if choice == 'cancel':
                return
            elif choice == 'new_only':
                fields_to_add = new_fields
            else:
                fields_to_add = measure_fields
        else:
            fields_to_add = measure_fields

        actual_position = None if position == -1 else position
        field_infos = self._fields_to_field_info_list(fields_to_add)
        if field_infos:
            self._on_fields_selected(field_infos, actual_position)

    def _context_add_table_columns(self, item_id: str, position: Optional[int] = None):
        """Context menu: Add all columns from table."""
        # Validate before adding
        if not self._validate_can_add():
            return

        fields = self._get_fields_from_item(item_id)

        column_fields = []
        for tbl, fld in fields:
            table_info = self.tables_data.get(tbl)
            if table_info:
                for field_info in table_info.fields:
                    if field_info.name == fld and field_info.field_type != "Measure":
                        column_fields.append((tbl, fld))
                        break

        if not column_fields:
            return

        duplicates, new_fields = self._get_duplicates_from_list(column_fields)

        if duplicates:
            table_name = self.tree.item(item_id, "values")[1]
            choice = self._show_bulk_duplicate_dialog(
                duplicates, new_fields, f"table '{table_name}'"
            )
            if choice == 'cancel':
                return
            elif choice == 'new_only':
                fields_to_add = new_fields
            else:
                fields_to_add = column_fields
        else:
            fields_to_add = column_fields

        actual_position = None if position == -1 else position
        field_infos = self._fields_to_field_info_list(fields_to_add)
        if field_infos:
            self._on_fields_selected(field_infos, actual_position)

    # =========================================================================
    # DRAG AND DROP
    # =========================================================================

    def _on_drag_start(self, event):
        """Handle drag start."""
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        values = self.tree.item(item_id, "values")
        if not values:
            return

        item_type = values[0]

        if item_type in ("field", "folder", "table"):
            self.tree._drag_start_pos = (event.x, event.y)

            selection = self.tree.selection()
            if item_id not in selection:
                selected_items = [item_id]
            else:
                selected_items = list(selection)

            self.tree._drag_candidate = {
                'item_id': item_id,
                'selected_items': selected_items,
                'item_type': item_type,
                'table_name': values[1],
                'data': values[2] if len(values) > 2 else ""
            }
            self.tree._drag_activated = False
            self._drag_label = None

    def _on_drag_motion(self, event):
        """Handle drag motion."""
        if not hasattr(self.tree, '_drag_candidate') or not self.tree._drag_candidate:
            return

        if not hasattr(self.tree, '_drag_activated') or not self.tree._drag_activated:
            start_x, start_y = self.tree._drag_start_pos
            distance = ((event.x - start_x)**2 + (event.y - start_y)**2)**0.5

            if distance > 5:
                self.tree._drag_activated = True
                self.tree._drag_data = self.tree._drag_candidate.copy()
                self.tree._drag_data['is_dragging'] = True

                selected_items = self.tree._drag_candidate.get('selected_items', [])

                drag_fields = []
                seen_fields = set()

                for sel_item_id in selected_items:
                    sel_values = self.tree.item(sel_item_id, "values")
                    if not sel_values:
                        continue

                    sel_type = sel_values[0]

                    if sel_type == "field":
                        tbl = sel_values[1]
                        fld = sel_values[2]
                        key = (tbl, fld)
                        if key not in seen_fields:
                            drag_fields.append(key)
                            seen_fields.add(key)
                    else:
                        item_fields = self._get_fields_from_item(sel_item_id)
                        for tbl, fld in item_fields:
                            key = (tbl, fld)
                            if key in seen_fields:
                                continue
                            table_info = self.tables_data.get(tbl)
                            if table_info:
                                for field_info in table_info.fields:
                                    if field_info.name == fld and field_info.field_type == "Measure":
                                        drag_fields.append(key)
                                        seen_fields.add(key)
                                        break

                if len(drag_fields) == 1:
                    label_text = drag_fields[0][1]
                else:
                    label_text = f"{len(drag_fields)} field(s)"

                self.tree._drag_data['fields'] = drag_fields

                self._create_drag_label(label_text)
                self.tree.config(cursor="hand2")

        if hasattr(self.tree, '_drag_activated') and self.tree._drag_activated:
            self._update_drag_label_position(event)

            if self._drop_target:
                mouse_y = self.tree.winfo_rooty() + event.y
                if self._is_over_drop_target(event):
                    self._drop_target.show_drop_indicator(mouse_y)
                else:
                    self._drop_target.hide_drop_indicator()

    def _create_drag_label(self, field_name: str):
        """Create a floating label that follows the cursor during drag."""
        root = self.winfo_toplevel()

        self._drag_label = tk.Toplevel(root)
        self._drag_label.wm_overrideredirect(True)
        self._drag_label.wm_attributes('-topmost', True)

        try:
            self._drag_label.wm_attributes('-alpha', 0.85)
        except:
            pass

        colors = self._theme_manager.colors
        label = tk.Label(
            self._drag_label,
            text=f"{field_name}",
            bg=colors['button_primary'],
            fg='#ffffff',
            font=("Segoe UI", 9, "bold"),
            padx=8,
            pady=4,
            relief=tk.FLAT,
            borderwidth=0
        )
        label.pack()

        self._drag_label.geometry("+0+0")

    def _update_drag_label_position(self, event):
        """Update the floating drag label position."""
        if self._drag_label and self._drag_label.winfo_exists():
            x = self.tree.winfo_rootx() + event.x + 15
            y = self.tree.winfo_rooty() + event.y + 10
            self._drag_label.geometry(f"+{x}+{y}")

    def _destroy_drag_label(self):
        """Destroy the floating drag label."""
        if hasattr(self, '_drag_label') and self._drag_label:
            try:
                self._drag_label.destroy()
            except:
                pass
            self._drag_label = None

    def _on_drag_end(self, event):
        """Handle drag end."""
        self.tree.config(cursor="")
        self._destroy_drag_label()

        drop_position = None
        if self._drop_target:
            drop_position = self._drop_target.get_drop_position()
            self._drop_target.hide_drop_indicator()

        was_dragging = hasattr(self.tree, '_drag_activated') and self.tree._drag_activated

        if was_dragging:
            drag_data = getattr(self.tree, '_drag_data', None)
            if drag_data:
                drag_fields = drag_data.get('fields', [])

                if not drag_fields and drag_data.get('item_type') == 'field':
                    table_name = drag_data.get('table_name')
                    field_name = drag_data.get('data')
                    if table_name and field_name:
                        drag_fields = [(table_name, field_name)]

                if self._is_over_drop_target(event) and drag_fields:
                    # Validate before adding
                    if not self._validate_can_add():
                        self._clear_drag_data()
                        return

                    duplicates, new_fields = self._get_duplicates_from_list(drag_fields)

                    if duplicates and len(drag_fields) == 1:
                        if self._show_duplicate_dialogs:
                            if not ThemedMessageBox.askyesno(
                                self,
                                "Duplicate Field",
                                f"'{drag_fields[0][1]}' is already in the parameter.\n\nAdd it again?"
                            ):
                                self._clear_drag_data()
                                return
                        fields_to_add = drag_fields
                    elif duplicates:
                        item_name = f"{len(drag_fields)} field(s)"
                        choice = self._show_bulk_duplicate_dialog(duplicates, new_fields, item_name)
                        if choice == 'cancel':
                            self._clear_drag_data()
                            return
                        elif choice == 'new_only':
                            fields_to_add = new_fields
                        else:
                            fields_to_add = drag_fields
                    else:
                        fields_to_add = drag_fields

                    field_infos = self._fields_to_field_info_list(fields_to_add)
                    if field_infos:
                        self._on_fields_selected(field_infos, drop_position)

        self._clear_drag_data()

    def _is_over_drop_target(self, event) -> bool:
        """Check if the mouse is over the drop target."""
        if not self._drop_target:
            return False

        mouse_x = self.tree.winfo_rootx() + event.x
        mouse_y = self.tree.winfo_rooty() + event.y

        target_x = self._drop_target.winfo_rootx()
        target_y = self._drop_target.winfo_rooty()
        target_width = self._drop_target.winfo_width()
        target_height = self._drop_target.winfo_height()

        return (
            target_x <= mouse_x <= target_x + target_width and
            target_y <= mouse_y <= target_y + target_height
        )

    def _clear_drag_data(self):
        """Clear all drag-related data."""
        if hasattr(self.tree, '_drag_data'):
            self.tree._drag_data = None
        if hasattr(self.tree, '_drag_candidate'):
            self.tree._drag_candidate = None
        if hasattr(self.tree, '_drag_activated'):
            self.tree._drag_activated = False
        if hasattr(self.tree, '_drag_start_pos'):
            self.tree._drag_start_pos = None
        self._destroy_drag_label()

    # =========================================================================
    # THEME AND STYLING
    # =========================================================================

    def _update_empty_state(self):
        """Show or hide placeholder based on whether tree has content."""
        if hasattr(self, '_placeholder_label'):
            if self.tree.get_children():
                self._placeholder_label.lower()
            else:
                self._placeholder_label.lift()

    def _load_tree_icons(self):
        """Load SVG icons for tree items based on current theme."""
        if not PIL_AVAILABLE:
            return

        is_dark = self._theme_manager.is_dark
        icon_suffix = "-white" if is_dark else ""

        icon_mappings = {
            'table': ('table', 14),
            'folder': ('folder', 14),
            'measure': ('measure', 14),
            'column': ('columns', 14),
        }

        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Simple BW Icons"

        for icon_key, (svg_name, size) in icon_mappings.items():
            svg_path = icons_dir / f"{svg_name}{icon_suffix}.svg"
            png_path = icons_dir / f"{svg_name}{icon_suffix}.png"

            try:
                img = None

                if CAIROSVG_AVAILABLE and svg_path.exists():
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=size * 4,
                        output_height=size * 4
                    )
                    img = Image.open(io.BytesIO(png_data))
                elif png_path.exists():
                    img = Image.open(png_path)

                if img is not None:
                    img = img.resize((size, size), Image.Resampling.LANCZOS)
                    self._tree_icons[icon_key] = ImageTk.PhotoImage(img)

            except Exception as e:
                self.logger.debug(f"Failed to load tree icon {svg_name}: {e}")

        # Load search icon
        tool_icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"
        search_svg = tool_icons_dir / "magnifying-glass.svg"
        if CAIROSVG_AVAILABLE and search_svg.exists():
            try:
                png_data = cairosvg.svg2png(
                    url=str(search_svg),
                    output_width=14 * 4,
                    output_height=14 * 4
                )
                img = Image.open(io.BytesIO(png_data))
                img = img.resize((14, 14), Image.Resampling.LANCZOS)
                self._tree_icons['search'] = ImageTk.PhotoImage(img)
            except Exception as e:
                self.logger.debug(f"Failed to load search icon: {e}")

    def on_theme_changed(self):
        """Update widget colors when theme changes."""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        content_bg = colors['background']
        fg_color = colors['text_primary']
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        for frame in self._inner_frames:
            frame.config(bg=content_bg)

        self.filter_label.config(bg=content_bg, fg=fg_color)
        self.type_filter_radio_group.on_theme_changed()

        for btn in self._primary_buttons:
            btn.update_colors(
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg='#ffffff',
                disabled_bg=disabled_bg,
                disabled_fg=disabled_fg,
                canvas_bg=content_bg
            )

        for btn in self._secondary_buttons:
            btn.update_colors(
                bg=colors['button_secondary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                fg=colors['text_primary'],
                disabled_bg=disabled_bg,
                disabled_fg=disabled_fg,
                canvas_bg=content_bg
            )

        if hasattr(self, '_tree_scrollbar'):
            self._tree_scrollbar.on_theme_changed()

        if hasattr(self, '_tree_container'):
            tree_border = colors.get('border', '#3a3a4a')
            tree_bg = colors.get('section_bg', colors.get('surface', content_bg))
            self._tree_container.config(
                bg=tree_bg,
                highlightbackground=tree_border,
                highlightcolor=tree_border
            )

        if hasattr(self, '_placeholder_label'):
            tree_bg = colors.get('section_bg', colors.get('surface', content_bg))
            self._placeholder_label.config(
                fg=colors['text_muted'],
                bg=tree_bg
            )

        self._update_section_header_theme()
        self._configure_treeview_style()
        self._load_tree_icons()

        if hasattr(self, '_search_icon_label'):
            search_icon = self._tree_icons.get('search')
            if search_icon:
                self._search_icon_label.config(image=search_icon, bg=content_bg)
            else:
                try:
                    self._search_icon_label.config(bg=content_bg)
                except Exception:
                    pass

        if self.tables_data:
            self.set_fields(self.tables_data)

    def _configure_treeview_style(self):
        """Configure treeview selection colors based on current theme."""
        colors = self._theme_manager.colors
        style = ttk.Style()
        is_dark = self._theme_manager.is_dark

        tree_bg = colors.get('section_bg', colors.get('surface', colors['background']))
        text_color = colors.get('text_primary', '#000000')

        style.configure("Treeview",
                        background=tree_bg,
                        fieldbackground=tree_bg,
                        foreground=text_color)

        style.layout("Treeview", [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

        header_bg = colors.get('surface', colors.get('section_bg', colors['background']))
        style.configure("Treeview.Heading",
                        background=header_bg,
                        foreground=text_color,
                        relief='flat')
        style.map("Treeview.Heading",
                  background=[("active", header_bg)],
                  foreground=[("active", text_color)])

        selection_bg = colors.get('selection_highlight', colors.get('card_surface', colors['background']))
        selection_fg = '#ffffff' if is_dark else colors.get('text_primary', '#000000')

        style.map("Treeview",
                  background=[("selected", selection_bg)],
                  foreground=[("selected", selection_fg)])
