"""
Category Manager Panel
Panel for managing category columns and their labels with drag-drop reordering.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Optional, Tuple, TYPE_CHECKING
import logging

from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, ThemedScrollbar, ThemedMessageBox, Tooltip, ThemedInputDialog, ThemedContextMenu
from tools.field_parameters.field_parameters_core import CategoryLevel
from tools.field_parameters.dialogs import AddLabelDialog
from tools.field_parameters.panels.panel_base import SectionPanelMixin

if TYPE_CHECKING:
    from tools.field_parameters.field_parameters_ui import FieldParametersTab


class CategoryManagerPanel(SectionPanelMixin, ttk.LabelFrame):
    """Panel for managing category columns and their labels"""

    def __init__(self, parent, main_tab: 'FieldParametersTab'):
        ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe', padding="12")
        self._init_section_panel_mixin()
        self.main_tab = main_tab
        self.logger = logging.getLogger(__name__)

        # Button tracking for theme updates
        self._secondary_buttons = []

        # Create and set the section header labelwidget
        self._create_section_header(parent, "Category Columns", "folder")

        self.category_levels: List[CategoryLevel] = []

        self.setup_ui()

    def setup_ui(self):
        """Setup category manager UI"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        # Use 'background' (darkest) for inner content frames, not content_bg
        content_bg = colors['background']
        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # CRITICAL: Inner wrapper with Section.TFrame style fills LabelFrame with correct background
        # This is required per UI_DESIGN_PATTERNS.md - without it, section_bg shows in padding area
        self._content_wrapper = ttk.Frame(self, style='Section.TFrame', padding="15")
        self._content_wrapper.pack(fill=tk.BOTH, expand=True)

        # Track inner frames for theme updates
        self._inner_frames = []

        # Header row with arrow buttons on right (matches Parameter Builder spacing)
        header_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        header_frame.pack(fill=tk.X, pady=(0, 3))
        self._header_frame = header_frame  # Store for theme updates
        self._inner_frames.append(header_frame)

        # Right side - X/Up/Down arrow buttons (like Parameter Builder)
        arrow_frame = tk.Frame(header_frame, bg=content_bg)
        arrow_frame.pack(side=tk.RIGHT)
        self._arrow_frame = arrow_frame  # Store for theme updates

        self.remove_btn = RoundedButton(
            arrow_frame,
            text="\u2716",
            command=self._on_remove_item,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=26, height=24, radius=5,
            font=('Segoe UI', 8),
            canvas_bg=content_bg
        )
        self.remove_btn.pack(side=tk.LEFT, padx=1)
        self._secondary_buttons.append(self.remove_btn)

        # Dynamic tooltip for remove button based on selection count
        self._remove_tooltip = Tooltip(
            self.remove_btn,
            text=lambda: "Delete Categories" if hasattr(self, 'tree') and len(self.tree.selection()) > 1 else "Delete Category"
        )

        self.move_up_btn = RoundedButton(
            arrow_frame,
            text="\u25B2",
            command=self._on_move_up,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=26, height=24, radius=5,
            font=('Segoe UI', 8),
            canvas_bg=content_bg
        )
        self.move_up_btn.pack(side=tk.LEFT, padx=1)
        self._secondary_buttons.append(self.move_up_btn)

        self.move_down_btn = RoundedButton(
            arrow_frame,
            text="\u25BC",
            command=self._on_move_down,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=26, height=24, radius=5,
            font=('Segoe UI', 8),
            canvas_bg=content_bg
        )
        self.move_down_btn.pack(side=tk.LEFT, padx=1)
        self._secondary_buttons.append(self.move_down_btn)

        # Category columns treeview (with expandable labels)
        tree_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self._inner_frames.append(tree_frame)
        self._tree_frame = tree_frame

        # Get border color from theme (lighter border for unified look)
        tree_border = colors.get('border', '#3a3a4a')
        # Use section_bg to match treeview style background (consistent with _configure_treeview_style)
        tree_bg = colors.get('section_bg', colors.get('surface', content_bg))

        # Container with 1px border around tree and scrollbar (matches Parameter Builder style)
        tree_container = tk.Frame(tree_frame, bg=tree_bg,
                                  highlightbackground=tree_border, highlightcolor=tree_border,
                                  highlightthickness=1)
        tree_container.pack(fill=tk.BOTH, expand=True)
        self._tree_container = tree_container  # Store for theme updates

        self._tree_scrollbar = ThemedScrollbar(
            tree_container,
            command=None,  # Set after tree is created
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Style for treeview
        style = ttk.Style()
        style.configure("Category.Treeview", font=("Segoe UI", 9), rowheight=22)
        style.configure("Category.Treeview.Item", padding=(2, 2))

        # Configure selection colors (subtle gray instead of teal)
        self._configure_treeview_style()

        self.tree = ttk.Treeview(
            tree_container,
            yscrollcommand=self._tree_scrollbar.set,
            style="Category.Treeview",
            show="tree",  # Hide column headers
            selectmode="extended"  # Allow multi-select with Ctrl+click, Shift+click
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._tree_scrollbar._command = self.tree.yview

        # Configure column to stretch and fill available space
        # Note: minwidth removed - was causing panel to expand wider than grid allocation
        self.tree.column("#0", stretch=True)

        # Track expanded state for each category
        self._expanded_categories: Dict[int, bool] = {}

        # Drag-drop state
        self._drag_data = {"item": None, "is_label": False, "cat_idx": None, "label_idx": None}

        # Drop indicator for visual feedback during drag
        self._drop_indicator = None
        self._drop_target = None

        # Bind events
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
        self.tree.bind("<Double-Button-1>", self._on_tree_double_click)
        self.tree.bind("<F2>", self._on_f2_rename)

        # Right-click context menu
        self.tree.bind("<Button-3>", self._on_right_click)

        # Drag and drop
        self.tree.bind("<ButtonPress-1>", self._on_drag_start)
        self.tree.bind("<B1-Motion>", self._on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_drag_release)

        # Placeholder overlay - shown when tree is empty
        self._placeholder_label = tk.Label(
            tree_container,
            text="No category columns defined.\n\nClick 'Add Column' to create\ncategory columns for organizing fields.\n\nDouble-click to rename.",
            font=("Segoe UI", 10, "italic"),
            fg=colors['text_muted'],
            bg=tree_bg,
            justify="center"
        )
        # Use place to overlay on tree - initially visible (tree starts empty)
        self._placeholder_label.place(relx=0.5, rely=0.5, anchor="center")

        # Actions row - centered buttons: Add Column, Add Label, Rename
        actions_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        actions_frame.pack(fill=tk.X, pady=(8, 0))
        self._inner_frames.append(actions_frame)
        self._actions_frame = actions_frame

        # Inner frame to center the buttons
        btn_container = tk.Frame(actions_frame, bg=content_bg)
        btn_container.pack(expand=True)  # Center horizontally
        self._btn_container = btn_container

        # Add Column button (first)
        self.add_btn = RoundedButton(
            btn_container,
            text="Add Column",
            command=self._on_add_column,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=content_bg
        )
        self.add_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._secondary_buttons.append(self.add_btn)

        # Add Label button (second)
        self.add_label_btn = RoundedButton(
            btn_container,
            text="Add Label",
            command=self._on_add_label,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=content_bg
        )
        self.add_label_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.add_label_btn.set_enabled(False)  # Start disabled until selection
        self._secondary_buttons.append(self.add_label_btn)

        # Rename button (third)
        self.rename_btn = RoundedButton(
            btn_container,
            text="Rename",
            command=self._on_rename_item,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=content_bg
        )
        self.rename_btn.pack(side=tk.LEFT)
        self.rename_btn.set_enabled(False)  # Start disabled until selection
        self._secondary_buttons.append(self.rename_btn)

    def _on_add_column(self):
        """Add a new category column"""
        # Prompt for column name
        column_name = ThemedInputDialog.askstring(
            self,
            "Add Category Column",
            "Enter column name (e.g., 'Category', 'Subcategory'):"
        )

        if not column_name:
            return

        column_name = column_name.strip()

        # Validate: column name cannot match parameter name (would create duplicate columns)
        if self.main_tab.current_parameter:
            param_name = self.main_tab.current_parameter.parameter_name
            if column_name.lower() == param_name.lower():
                ThemedMessageBox.showerror(
                    self,
                    "Invalid Column Name",
                    f"Column name cannot be the same as the parameter name ('{param_name}').\n\n"
                    "This would create duplicate column names in the output."
                )
                return

        # Validate: column name cannot match existing column names
        for existing_level in self.category_levels:
            if column_name.lower() == existing_level.name.lower():
                ThemedMessageBox.showerror(
                    self,
                    "Duplicate Column Name",
                    f"A column named '{existing_level.name}' already exists.\n\n"
                    "Please choose a different name."
                )
                return

        # Create new CategoryLevel
        new_level = CategoryLevel(
            name=column_name,
            sort_order=len(self.category_levels) + 1,
            column_name=column_name,
            labels=[]
        )

        # Add to main tab
        self.main_tab.on_add_category_level(column_name, column_name)

        # Sync and refresh
        if self.main_tab.current_parameter:
            self.category_levels = self.main_tab.current_parameter.category_levels
        self._refresh_list()

        # Select the new column
        if self.category_levels:
            self.tree.selection_set(f"cat_{len(self.category_levels) - 1}")

    def _on_add_label(self):
        """Add new label(s) to the selected category with position options"""
        cat_idx = self._get_selected_category_index()
        if cat_idx is None or cat_idx >= len(self.category_levels):
            return

        level = self.category_levels[cat_idx]
        if level.is_calculated:
            return

        # Check if a label is currently selected (for above/below options)
        label_info = self._get_selected_label_index()
        selected_label_idx = label_info[1] if label_info and label_info[0] == cat_idx else None

        # Show dialog with position options
        dialog = AddLabelDialog(
            self.winfo_toplevel(),
            level.name,
            level.labels,
            selected_label_idx
        )
        self.wait_window(dialog)

        # Check if user confirmed (result_labels is now a list)
        if not dialog.result_labels:
            return

        # Insert label(s) at chosen position
        insert_idx = dialog.result_position
        for i, label_name in enumerate(dialog.result_labels):
            level.labels.insert(insert_idx + i, label_name)

        # Recalculate sort orders for all field items using this category
        self._recalculate_label_sort_orders(cat_idx)

        # Expand the category and refresh
        self._expanded_categories[cat_idx] = True
        self._refresh_list()

        # Select the new label(s)
        if len(dialog.result_labels) == 1:
            # Single label - select it
            self.tree.selection_set(f"cat_{cat_idx}_label_{insert_idx}")
            self.tree.see(f"cat_{cat_idx}_label_{insert_idx}")
        else:
            # Multiple labels - select all of them
            new_selection = [f"cat_{cat_idx}_label_{insert_idx + i}"
                           for i in range(len(dialog.result_labels))]
            self.tree.selection_set(*new_selection)
            self.tree.see(new_selection[0])

        self._on_selection_changed(None)
        self.main_tab.update_preview()

    def _on_rename_item(self):
        """Rename selected item (category or label)"""
        label_info = self._get_selected_label_index()

        if label_info:
            # Renaming a LABEL
            cat_idx, label_idx = label_info
            level = self.category_levels[cat_idx]
            old_name = level.labels[label_idx]

            new_name = ThemedInputDialog.askstring(
                self,
                "Rename Label",
                f"Enter new name for '{old_name}':",
                initialvalue=old_name
            )

            if not new_name:
                return

            new_name = new_name.strip()
            if new_name == old_name:
                return

            # Check for duplicate
            if new_name in level.labels:
                ThemedMessageBox.showerror(
                    self,
                    "Duplicate Label",
                    f"A label named '{new_name}' already exists in '{level.name}'.\n\n"
                    "Please choose a different name."
                )
                return

            # Also update any field items that use this label
            self._update_field_category_labels(cat_idx, old_name, new_name)

            # Update the label
            level.labels[label_idx] = new_name
            self._refresh_list()
            self.tree.selection_set(f"cat_{cat_idx}_label_{label_idx}")
            self.main_tab.update_preview()
        else:
            # Renaming a CATEGORY - delegate to existing method
            self._on_rename_column()

    def _update_field_category_labels(self, cat_idx: int, old_label: str, new_label: str):
        """Update all field items that use the old label to use the new label"""
        if not self.main_tab.current_parameter:
            return

        for field_item in self.main_tab.current_parameter.fields:
            if field_item.categories and len(field_item.categories) > cat_idx:
                sort_order, label = field_item.categories[cat_idx]
                if label == old_label:
                    field_item.categories[cat_idx] = (sort_order, new_label)

    def _on_rename_column(self):
        """Rename selected category column"""
        idx = self._get_selected_category_index()
        if idx is None or idx >= len(self.category_levels):
            return
        level = self.category_levels[idx]
        old_name = level.name

        # Prompt for new name
        new_name = ThemedInputDialog.askstring(
            self,
            "Rename Category Column",
            f"Enter new name for '{old_name}':",
            initialvalue=old_name
        )

        if not new_name:
            return

        new_name = new_name.strip()

        # Check if name unchanged
        if new_name == old_name:
            return

        # Validate: cannot match parameter name
        if self.main_tab.current_parameter:
            param_name = self.main_tab.current_parameter.parameter_name
            if new_name.lower() == param_name.lower():
                ThemedMessageBox.showerror(
                    self,
                    "Invalid Column Name",
                    f"Column name cannot be the same as the parameter name ('{param_name}').\n\n"
                    "This would create duplicate column names in the output."
                )
                return

        # Validate: cannot match other existing column names
        for i, existing_level in enumerate(self.category_levels):
            if i != idx and new_name.lower() == existing_level.name.lower():
                ThemedMessageBox.showerror(
                    self,
                    "Duplicate Column Name",
                    f"A column named '{existing_level.name}' already exists.\n\n"
                    "Please choose a different name."
                )
                return

        # Apply rename
        level.name = new_name
        level.column_name = new_name

        # Refresh display
        self._refresh_list()
        self.tree.selection_set(f"cat_{idx}")
        self.main_tab.update_preview()

    def _on_remove_item(self):
        """Remove selected item(s) (category or label(s))"""
        # Check for multi-select labels first
        multi_labels = self._get_selected_labels()
        if multi_labels and len(multi_labels[1]) > 1:
            # Removing multiple labels
            cat_idx, label_indices = multi_labels
            level = self.category_levels[cat_idx]
            label_names = [level.labels[i] for i in label_indices]

            confirm = ThemedMessageBox.askyesno(
                self,
                "Remove Labels",
                f"Remove {len(label_names)} labels from '{level.name}'?\n\n"
                f"Labels: {', '.join(label_names[:5])}{'...' if len(label_names) > 5 else ''}\n\n"
                "Any fields assigned to these labels will become uncategorized."
            )

            if confirm:
                # Clear labels from field items
                for label_name in label_names:
                    self._clear_field_category_label(cat_idx, label_name)

                # Remove labels (reverse order to maintain indices)
                for i in sorted(label_indices, reverse=True):
                    del level.labels[i]

                self._recalculate_label_sort_orders(cat_idx)
                self._refresh_list()
                self.tree.selection_set(f"cat_{cat_idx}")
                self.main_tab.update_preview()
            return

        # Single selection
        label_info = self._get_selected_label_index()

        if label_info:
            # Removing a single LABEL
            cat_idx, label_idx = label_info
            level = self.category_levels[cat_idx]
            label_name = level.labels[label_idx]

            confirm = ThemedMessageBox.askyesno(
                self,
                "Remove Label",
                f"Remove label '{label_name}' from '{level.name}'?\n\n"
                "Any fields assigned to this label will become uncategorized."
            )

            if confirm:
                # Clear this label from any field items that use it
                self._clear_field_category_label(cat_idx, label_name)

                # Remove the label
                del level.labels[label_idx]
                self._refresh_list()

                # Select the category
                self.tree.selection_set(f"cat_{cat_idx}")
                self.main_tab.update_preview()
        else:
            # Removing a CATEGORY
            idx = self._get_selected_category_index()
            if idx is None or idx >= len(self.category_levels):
                return

            level = self.category_levels[idx]

            confirm = ThemedMessageBox.askyesno(
                self,
                "Remove Category Column",
                f"Remove column '{level.name}'?\n\n"
                "All category assignments using this column will be cleared."
            )

            if confirm:
                self.main_tab.on_remove_category_level(idx)
                self._refresh_list()

    def _clear_field_category_label(self, cat_idx: int, label_name: str):
        """Clear this label from any field items that use it (set to empty)"""
        if not self.main_tab.current_parameter:
            return

        for field_item in self.main_tab.current_parameter.fields:
            if field_item.categories and len(field_item.categories) > cat_idx:
                sort_order, label = field_item.categories[cat_idx]
                if label == label_name:
                    # Clear the label (set to empty string)
                    field_item.categories[cat_idx] = (0, "")

    def _on_move_up(self):
        """Move selected item(s) up (category or label(s))"""
        # Check for multi-select labels first
        multi_labels = self._get_selected_labels()
        if multi_labels and len(multi_labels[1]) > 1:
            # Move multiple labels up as a block
            cat_idx, label_indices = multi_labels
            min_idx = min(label_indices)
            if min_idx == 0:
                return  # Already at top

            level = self.category_levels[cat_idx]
            # Move each selected label up by one (process in order from top to bottom)
            for old_idx in sorted(label_indices):
                new_idx = old_idx - 1
                label = level.labels.pop(old_idx)
                level.labels.insert(new_idx, label)

            # Recalculate all sort orders
            self._recalculate_label_sort_orders(cat_idx)

            # Keep category expanded and restore selection
            self._expanded_categories[cat_idx] = True
            new_selection = [f"cat_{cat_idx}_label_{idx - 1}" for idx in label_indices]
            self._refresh_list()
            self.tree.selection_set(*new_selection)
            self.tree.see(new_selection[0])
            self._on_selection_changed(None)
            self.main_tab.update_preview()
            return

        # Single label or category selection
        label_info = self._get_selected_label_index()

        if label_info:
            # Moving a single LABEL up within its category
            cat_idx, label_idx = label_info
            if label_idx == 0:
                return  # Already at top

            level = self.category_levels[cat_idx]
            # Swap labels
            level.labels[label_idx], level.labels[label_idx - 1] = \
                level.labels[label_idx - 1], level.labels[label_idx]

            # Update sort orders in field items that use these labels
            self._swap_label_sort_orders(cat_idx, label_idx - 1, label_idx)

            # Keep category expanded
            self._expanded_categories[cat_idx] = True
            new_selection = f"cat_{cat_idx}_label_{label_idx - 1}"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            # Update button states for new position
            self._on_selection_changed(None)
            self.main_tab.update_preview()
        else:
            # Moving a CATEGORY up
            idx = self._get_selected_category_index()
            if idx is None or idx == 0:
                return

            # Save current expanded states for both items
            was_expanded_current = self._expanded_categories.get(idx, False)
            was_expanded_above = self._expanded_categories.get(idx - 1, False)

            # Swap in the tracking dict (indices swap)
            self._expanded_categories[idx - 1] = was_expanded_current
            self._expanded_categories[idx] = was_expanded_above

            self.category_levels[idx], self.category_levels[idx-1] = \
                self.category_levels[idx-1], self.category_levels[idx]

            new_selection = f"cat_{idx-1}"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            # Update button states for new position
            self._on_selection_changed(None)
            self.main_tab.on_categories_reordered(self.category_levels)

    def _on_move_down(self):
        """Move selected item(s) down (category or label(s))"""
        # Check for multi-select labels first
        multi_labels = self._get_selected_labels()
        if multi_labels and len(multi_labels[1]) > 1:
            # Move multiple labels down as a block
            cat_idx, label_indices = multi_labels
            level = self.category_levels[cat_idx]
            max_idx = max(label_indices)
            if max_idx >= len(level.labels) - 1:
                return  # Already at bottom

            # Move each selected label down by one (process in reverse order from bottom to top)
            for old_idx in sorted(label_indices, reverse=True):
                new_idx = old_idx + 1
                label = level.labels.pop(old_idx)
                level.labels.insert(new_idx, label)

            # Recalculate all sort orders
            self._recalculate_label_sort_orders(cat_idx)

            # Keep category expanded and restore selection
            self._expanded_categories[cat_idx] = True
            new_selection = [f"cat_{cat_idx}_label_{idx + 1}" for idx in label_indices]
            self._refresh_list()
            self.tree.selection_set(*new_selection)
            self.tree.see(new_selection[-1])
            self._on_selection_changed(None)
            self.main_tab.update_preview()
            return

        # Single label or category selection
        label_info = self._get_selected_label_index()

        if label_info:
            # Moving a single LABEL down within its category
            cat_idx, label_idx = label_info
            level = self.category_levels[cat_idx]
            if label_idx >= len(level.labels) - 1:
                return  # Already at bottom

            # Swap labels
            level.labels[label_idx], level.labels[label_idx + 1] = \
                level.labels[label_idx + 1], level.labels[label_idx]

            # Update sort orders in field items that use these labels
            self._swap_label_sort_orders(cat_idx, label_idx, label_idx + 1)

            # Keep category expanded
            self._expanded_categories[cat_idx] = True
            new_selection = f"cat_{cat_idx}_label_{label_idx + 1}"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            # Update button states for new position
            self._on_selection_changed(None)
            self.main_tab.update_preview()
        else:
            # Moving a CATEGORY down
            idx = self._get_selected_category_index()
            if idx is None or idx >= len(self.category_levels) - 1:
                return

            # Save current expanded states for both items
            was_expanded_current = self._expanded_categories.get(idx, False)
            was_expanded_below = self._expanded_categories.get(idx + 1, False)

            # Swap in the tracking dict (indices swap)
            self._expanded_categories[idx + 1] = was_expanded_current
            self._expanded_categories[idx] = was_expanded_below

            self.category_levels[idx], self.category_levels[idx+1] = \
                self.category_levels[idx+1], self.category_levels[idx]

            new_selection = f"cat_{idx+1}"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            # Update button states for new position
            self._on_selection_changed(None)
            self.main_tab.on_categories_reordered(self.category_levels)

    def _swap_label_sort_orders(self, cat_idx: int, idx1: int, idx2: int):
        """Swap sort orders for labels at idx1 and idx2 in field items"""
        if not self.main_tab.current_parameter:
            return

        level = self.category_levels[cat_idx]
        label1 = level.labels[idx1]
        label2 = level.labels[idx2]

        # The sort order is based on label position, so we need to update
        # any field items using these labels to swap their sort orders
        for field_item in self.main_tab.current_parameter.fields:
            if field_item.categories and len(field_item.categories) > cat_idx:
                sort_order, label = field_item.categories[cat_idx]
                if label == label1:
                    # Update sort order to new position
                    field_item.categories[cat_idx] = (idx2 + 1, label)
                elif label == label2:
                    field_item.categories[cat_idx] = (idx1 + 1, label)

    # =========================================================================
    # Right-click Context Menu
    # =========================================================================

    def _on_right_click(self, event):
        """Show context menu on right-click"""
        item = self.tree.identify_row(event.y)
        if not item:
            return

        # Only change selection if clicked item is not already selected
        # This preserves multi-selection when right-clicking on selected items
        current_selection = self.tree.selection()
        if item not in current_selection:
            self.tree.selection_set(item)

        # Create themed context menu
        menu = ThemedContextMenu(self, self._theme_manager)

        # Check for multi-select labels
        multi_labels = self._get_selected_labels()
        if multi_labels and len(multi_labels[1]) > 1:
            # Multiple labels selected
            cat_idx, label_indices = multi_labels
            level = self.category_levels[cat_idx]
            num_labels = len(level.labels)
            min_idx = min(label_indices)
            max_idx = max(label_indices)

            # Only add Move to Top if not already at top
            if min_idx > 0:
                menu.add_command(label="Move to Top", command=self._move_to_top)
            # Only add Move to Bottom if not already at bottom
            if max_idx < num_labels - 1:
                menu.add_command(label="Move to Bottom", command=self._move_to_bottom)
            menu.add_command(label="Move to Position...", command=self._move_to_position)
            menu.add_separator()
            menu.add_command(label="Remove", command=self._on_remove_item)

            menu.show(event.x_root, event.y_root)
            return

        # Single selection
        label_info = self._get_selected_label_index()
        cat_idx = self._get_selected_category_index()

        if label_info:
            # Single label selected
            c_idx, l_idx = label_info
            level = self.category_levels[c_idx]
            num_labels = len(level.labels)

            # Only add Move to Top if not already at top
            if l_idx > 0:
                menu.add_command(label="Move to Top", command=self._move_to_top)
            # Only add Move to Bottom if not already at bottom
            if l_idx < num_labels - 1:
                menu.add_command(label="Move to Bottom", command=self._move_to_bottom)
            menu.add_command(label="Move to Position...", command=self._move_to_position)
            menu.add_separator()
            menu.add_command(label="Remove", command=self._on_remove_item)

        elif cat_idx is not None:
            # Category selected
            level = self.category_levels[cat_idx]
            if level.is_calculated:
                return  # No menu for calculated columns

            num_cats = len(self.category_levels)

            # Only add Move to Top if not already at top
            if cat_idx > 0:
                menu.add_command(label="Move to Top", command=self._move_to_top)
            # Only add Move to Bottom if not already at bottom
            if cat_idx < num_cats - 1:
                menu.add_command(label="Move to Bottom", command=self._move_to_bottom)
            menu.add_command(label="Move to Position...", command=self._move_to_position)
            menu.add_separator()
            menu.add_command(label="Remove", command=self._on_remove_item)

        # Show the menu
        menu.show(event.x_root, event.y_root)

    def _move_to_top(self):
        """Move selected item(s) to top of list"""
        # Check for multi-select labels first
        multi_labels = self._get_selected_labels()
        if multi_labels and len(multi_labels[1]) > 1:
            # Move multiple labels to top
            cat_idx, label_indices = multi_labels
            min_idx = min(label_indices)
            if min_idx == 0:
                return  # Already at top

            level = self.category_levels[cat_idx]
            # Extract selected labels in order
            selected_labels = [level.labels[i] for i in sorted(label_indices)]
            # Remove them from their positions (reverse order to maintain indices)
            for i in sorted(label_indices, reverse=True):
                level.labels.pop(i)
            # Insert at top in their original relative order
            for i, label in enumerate(selected_labels):
                level.labels.insert(i, label)

            self._recalculate_label_sort_orders(cat_idx)
            self._expanded_categories[cat_idx] = True
            new_selection = [f"cat_{cat_idx}_label_{i}" for i in range(len(selected_labels))]
            self._refresh_list()
            self.tree.selection_set(*new_selection)
            self.tree.see(new_selection[0])
            self._on_selection_changed(None)
            self.main_tab.update_preview()
            return

        # Single label or category
        label_info = self._get_selected_label_index()

        if label_info:
            # Move single label to top
            cat_idx, label_idx = label_info
            if label_idx == 0:
                return

            level = self.category_levels[cat_idx]
            label = level.labels.pop(label_idx)
            level.labels.insert(0, label)

            # Update all sort orders for this category's labels
            self._recalculate_label_sort_orders(cat_idx)

            self._expanded_categories[cat_idx] = True
            new_selection = f"cat_{cat_idx}_label_0"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            self._on_selection_changed(None)
            self.main_tab.update_preview()
        else:
            # Move category to top
            idx = self._get_selected_category_index()
            if idx is None or idx == 0:
                return

            # Save expanded state
            was_expanded = self._expanded_categories.get(idx, False)

            # Remove and insert at top
            level = self.category_levels.pop(idx)
            self.category_levels.insert(0, level)

            # Shift expanded states
            new_expanded = {0: was_expanded}
            for old_idx, expanded in self._expanded_categories.items():
                if old_idx < idx:
                    new_expanded[old_idx + 1] = expanded
                elif old_idx > idx:
                    new_expanded[old_idx] = expanded
            self._expanded_categories = new_expanded

            new_selection = "cat_0"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            self._on_selection_changed(None)
            self.main_tab.on_categories_reordered(self.category_levels)

    def _move_to_bottom(self):
        """Move selected item(s) to bottom of list"""
        # Check for multi-select labels first
        multi_labels = self._get_selected_labels()
        if multi_labels and len(multi_labels[1]) > 1:
            # Move multiple labels to bottom
            cat_idx, label_indices = multi_labels
            level = self.category_levels[cat_idx]
            max_idx = max(label_indices)
            if max_idx >= len(level.labels) - 1:
                return  # Already at bottom

            # Extract selected labels in order
            selected_labels = [level.labels[i] for i in sorted(label_indices)]
            # Remove them from their positions (reverse order to maintain indices)
            for i in sorted(label_indices, reverse=True):
                level.labels.pop(i)
            # Append at bottom in their original relative order
            for label in selected_labels:
                level.labels.append(label)

            self._recalculate_label_sort_orders(cat_idx)
            self._expanded_categories[cat_idx] = True
            num_labels = len(level.labels)
            num_selected = len(selected_labels)
            new_selection = [f"cat_{cat_idx}_label_{num_labels - num_selected + i}" for i in range(num_selected)]
            self._refresh_list()
            self.tree.selection_set(*new_selection)
            self.tree.see(new_selection[-1])
            self._on_selection_changed(None)
            self.main_tab.update_preview()
            return

        # Single label or category
        label_info = self._get_selected_label_index()

        if label_info:
            # Move single label to bottom
            cat_idx, label_idx = label_info
            level = self.category_levels[cat_idx]
            if label_idx >= len(level.labels) - 1:
                return

            label = level.labels.pop(label_idx)
            level.labels.append(label)

            # Update all sort orders for this category's labels
            self._recalculate_label_sort_orders(cat_idx)

            new_label_idx = len(level.labels) - 1
            self._expanded_categories[cat_idx] = True
            new_selection = f"cat_{cat_idx}_label_{new_label_idx}"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            self._on_selection_changed(None)
            self.main_tab.update_preview()
        else:
            # Move category to bottom
            idx = self._get_selected_category_index()
            if idx is None or idx >= len(self.category_levels) - 1:
                return

            # Save expanded state
            was_expanded = self._expanded_categories.get(idx, False)

            # Remove and append
            level = self.category_levels.pop(idx)
            self.category_levels.append(level)

            # Shift expanded states
            new_idx = len(self.category_levels) - 1
            new_expanded = {new_idx: was_expanded}
            for old_idx, expanded in self._expanded_categories.items():
                if old_idx > idx:
                    new_expanded[old_idx - 1] = expanded
                elif old_idx < idx:
                    new_expanded[old_idx] = expanded
            self._expanded_categories = new_expanded

            new_selection = f"cat_{new_idx}"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            self._on_selection_changed(None)
            self.main_tab.on_categories_reordered(self.category_levels)

    def _move_to_position(self):
        """Move selected item(s) to a specific position via dialog"""
        from tkinter import simpledialog

        label_info = self._get_selected_label_index()
        cat_idx = self._get_selected_category_index()

        if label_info:
            # Moving label(s)
            c_idx, l_idx = label_info
            level = self.category_levels[c_idx]
            num_labels = len(level.labels)

            # Check for multi-select
            multi_labels = self._get_selected_labels()
            if multi_labels and len(multi_labels[1]) > 1:
                current_pos = min(multi_labels[1]) + 1
                item_count = len(multi_labels[1])
                prompt = f"Move {item_count} labels to position (1-{num_labels}):"
            else:
                current_pos = l_idx + 1
                item_count = 1
                prompt = f"Move label to position (1-{num_labels}):"

            result = simpledialog.askinteger(
                "Move to Position",
                prompt,
                initialvalue=current_pos,
                minvalue=1,
                maxvalue=num_labels,
                parent=self
            )

            if result is None:
                return

            target_idx = result - 1  # Convert to 0-based

            if multi_labels and len(multi_labels[1]) > 1:
                # Move multiple labels
                label_indices = multi_labels[1]
                selected_labels = [level.labels[i] for i in sorted(label_indices)]

                # Remove in reverse order
                for i in sorted(label_indices, reverse=True):
                    level.labels.pop(i)

                # Insert at target position
                for i, label in enumerate(selected_labels):
                    level.labels.insert(target_idx + i, label)

                self._recalculate_label_sort_orders(c_idx)
                self._expanded_categories[c_idx] = True
                new_selection = [f"cat_{c_idx}_label_{target_idx + i}" for i in range(len(selected_labels))]
                self._refresh_list()
                self.tree.selection_set(*new_selection)
                self.tree.see(new_selection[0])
            else:
                # Move single label
                label = level.labels.pop(l_idx)
                level.labels.insert(target_idx, label)

                self._recalculate_label_sort_orders(c_idx)
                self._expanded_categories[c_idx] = True
                new_selection = f"cat_{c_idx}_label_{target_idx}"
                self._refresh_list()
                self.tree.selection_set(new_selection)
                self.tree.see(new_selection)

            self._on_selection_changed(None)
            self.main_tab.update_preview()

        elif cat_idx is not None:
            # Moving category
            num_cats = len(self.category_levels)
            current_pos = cat_idx + 1

            result = simpledialog.askinteger(
                "Move to Position",
                f"Move category to position (1-{num_cats}):",
                initialvalue=current_pos,
                minvalue=1,
                maxvalue=num_cats,
                parent=self
            )

            if result is None:
                return

            target_idx = result - 1  # Convert to 0-based

            if target_idx == cat_idx:
                return  # No change

            # Save expanded state
            was_expanded = self._expanded_categories.get(cat_idx, False)

            # Remove and insert
            level = self.category_levels.pop(cat_idx)
            self.category_levels.insert(target_idx, level)

            # Rebuild expanded states
            new_expanded = {}
            for old_idx, expanded in list(self._expanded_categories.items()):
                if old_idx == cat_idx:
                    new_expanded[target_idx] = was_expanded
                elif cat_idx < target_idx:
                    # Moving down: indices between old and new shift up
                    if old_idx > cat_idx and old_idx <= target_idx:
                        new_expanded[old_idx - 1] = expanded
                    else:
                        new_expanded[old_idx] = expanded
                else:
                    # Moving up: indices between new and old shift down
                    if old_idx >= target_idx and old_idx < cat_idx:
                        new_expanded[old_idx + 1] = expanded
                    else:
                        new_expanded[old_idx] = expanded
            new_expanded[target_idx] = was_expanded
            self._expanded_categories = new_expanded

            new_selection = f"cat_{target_idx}"
            self._refresh_list()
            self.tree.selection_set(new_selection)
            self.tree.see(new_selection)
            self._on_selection_changed(None)
            self.main_tab.on_categories_reordered(self.category_levels)

    def _recalculate_label_sort_orders(self, cat_idx: int):
        """Recalculate sort orders for all labels in a category after reordering"""
        if not self.main_tab.current_parameter:
            return

        level = self.category_levels[cat_idx]
        # Build a map of label -> new sort order (1-based)
        label_to_sort = {label: idx + 1 for idx, label in enumerate(level.labels)}

        # Update all field items
        for field_item in self.main_tab.current_parameter.fields:
            if field_item.categories and len(field_item.categories) > cat_idx:
                _, label = field_item.categories[cat_idx]
                if label and label in label_to_sort:
                    field_item.categories[cat_idx] = (label_to_sort[label], label)

    # =========================================================================
    # Drag and Drop
    # =========================================================================

    def _on_drag_start(self, event):
        """Start drag operation - supports multi-select for labels and categories"""
        item = self.tree.identify_row(event.y)
        if not item:
            self._drag_data = {"item": None}
            return

        # Check if click was on the expand/collapse arrow (indicator)
        # Only skip drag if clicking directly on the tree indicator AND item has children
        region = self.tree.identify_region(event.x, event.y)
        if region == "tree":
            # Check if this item has children (only then does it have an expand arrow)
            children = self.tree.get_children(item)
            if children:
                # Click was on the tree indicator (expand arrow) - don't drag
                self._drag_data = {"item": None}
                return

        # Check current selection before treeview changes it
        current_selection = self.tree.selection()

        # Check if it's a label or category
        parent = self.tree.parent(item)
        if parent:
            # It's a label - check if we're clicking on an already-selected item
            # to preserve multi-selection for drag
            if item in current_selection:
                # Check if we have multiple labels selected from same category
                multi_labels = self._get_selected_labels()
                if multi_labels and len(multi_labels[1]) > 1:
                    cat_idx, label_indices = multi_labels
                    self._drag_data = {
                        "item": item,
                        "is_label": True,
                        "is_multi": True,
                        "cat_idx": cat_idx,
                        "label_indices": label_indices,
                        "selection": list(current_selection)
                    }
                    return

            # Single label drag
            try:
                parts = item.split("_")
                cat_idx = int(parts[1])
                label_idx = int(parts[3])
                self._drag_data = {
                    "item": item,
                    "is_label": True,
                    "is_multi": False,
                    "cat_idx": cat_idx,
                    "label_idx": label_idx
                }
            except (IndexError, ValueError):
                self._drag_data = {"item": None}
        else:
            # It's a category - check for multi-select
            if item in current_selection:
                # Check if we have multiple categories selected
                selected_cats = self._get_selected_categories()
                if selected_cats and len(selected_cats) > 1:
                    # Multi-category drag
                    # Check none are calculated
                    for idx in selected_cats:
                        if self.category_levels[idx].is_calculated:
                            self._drag_data = {"item": None}
                            return
                    self._drag_data = {
                        "item": item,
                        "is_label": False,
                        "is_multi": True,
                        "cat_indices": selected_cats,
                        "selection": list(current_selection)
                    }
                    return

            # Single category drag
            try:
                cat_idx = int(item.split("_")[1])
                level = self.category_levels[cat_idx]
                if level.is_calculated:
                    self._drag_data = {"item": None}
                    return
                self._drag_data = {
                    "item": item,
                    "is_label": False,
                    "is_multi": False,
                    "cat_idx": cat_idx,
                    "label_idx": None
                }
            except (IndexError, ValueError):
                self._drag_data = {"item": None}

    def _on_drag_motion(self, event):
        """Handle drag motion - show visual feedback with drop indicator"""
        if not self._drag_data.get("item"):
            return

        # Change cursor to indicate drag
        self.tree.config(cursor="fleur")

        # Restore multi-selection during drag (treeview may have changed it on click)
        if self._drag_data.get("is_multi") and self._drag_data.get("selection"):
            current_sel = self.tree.selection()
            saved_sel = self._drag_data["selection"]
            if set(current_sel) != set(saved_sel):
                self.tree.selection_set(*saved_sel)

        # Show drop indicator
        self._show_tree_drop_indicator(event.y)

    def _show_tree_drop_indicator(self, y: int):
        """Show drop indicator line in treeview at the specified y position"""
        # Create indicator if it doesn't exist
        colors = self._theme_manager.colors
        if self._drop_indicator is None:
            self._drop_indicator = tk.Frame(
                self.tree,
                height=3,
                bg=colors['button_primary']  # Teal in light mode, blue in dark mode
            )

        # Find target item at this position
        target_item = self.tree.identify_row(y)
        source_is_label = self._drag_data.get("is_label", False)
        source_cat_idx = self._drag_data.get("cat_idx")

        if not target_item:
            # Below all items - show at bottom
            all_items = self.tree.get_children()
            if all_items:
                last_item = all_items[-1]
                # If last item is expanded, get its last child
                children = self.tree.get_children(last_item)
                if children:
                    last_item = children[-1]
                bbox = self.tree.bbox(last_item)
                if bbox:
                    y_pos = bbox[1] + bbox[3]  # Bottom of last item
                    self._drop_indicator.place(x=0, y=y_pos, relwidth=1.0, height=3)
                    self._drop_target = ("end", None)
                    return
            self._drop_indicator.place_forget()
            return

        # Get bounding box of target item
        bbox = self.tree.bbox(target_item)
        if not bbox:
            self._drop_indicator.place_forget()
            return

        item_y = bbox[1]
        item_height = bbox[3]
        item_center = item_y + item_height / 2

        # Determine if target is a label or category
        target_parent = self.tree.parent(target_item)
        target_is_label = bool(target_parent)

        if source_is_label:
            # Dragging a label - only show indicator for labels in same category
            if target_is_label:
                try:
                    parts = target_item.split("_")
                    target_cat_idx = int(parts[1])
                    if target_cat_idx != source_cat_idx:
                        self._drop_indicator.place_forget()
                        return
                except (IndexError, ValueError):
                    pass

            # Position indicator above or below target based on mouse position
            if y < item_center:
                # Above target
                y_pos = item_y
            else:
                # Below target
                y_pos = item_y + item_height

            # Indent for labels
            x_offset = 20 if target_is_label or source_is_label else 0
            self._drop_indicator.place(x=x_offset, y=y_pos, width=self.tree.winfo_width() - x_offset, height=3)
            self._drop_target = (target_item, y < item_center)
        else:
            # Dragging a category - only show indicator at category level
            if target_is_label:
                # Get the parent category
                target_item = target_parent
                bbox = self.tree.bbox(target_item)
                if not bbox:
                    self._drop_indicator.place_forget()
                    return
                item_y = bbox[1]
                item_height = bbox[3]
                item_center = item_y + item_height / 2

            # Position indicator above or below target
            if y < item_center:
                y_pos = item_y
            else:
                # Below - check if expanded and place after children
                children = self.tree.get_children(target_item)
                if children:
                    last_child = children[-1]
                    child_bbox = self.tree.bbox(last_child)
                    if child_bbox:
                        y_pos = child_bbox[1] + child_bbox[3]
                    else:
                        y_pos = item_y + item_height
                else:
                    y_pos = item_y + item_height

            self._drop_indicator.place(x=0, y=y_pos, relwidth=1.0, height=3)
            self._drop_target = (target_item, y < item_center)

    def _hide_tree_drop_indicator(self):
        """Hide the tree drop indicator"""
        if self._drop_indicator:
            self._drop_indicator.place_forget()
        self._drop_target = None

    def _on_drag_release(self, event):
        """Handle drag release - reorder items (supports multi-select)"""
        self.tree.config(cursor="")
        self._hide_tree_drop_indicator()

        if not self._drag_data.get("item"):
            return

        source_is_label = self._drag_data.get("is_label", False)
        source_is_multi = self._drag_data.get("is_multi", False)
        source_cat_idx = self._drag_data.get("cat_idx")

        target_item = self.tree.identify_row(event.y)
        drop_at_end = False
        drop_below_target = False  # True if dropping below/after the target item

        if not target_item:
            # Dropping below all items - need to handle "end" position
            if not source_is_label:
                # For categories, allow dropping at the end
                drop_at_end = True
            else:
                self._drag_data = {"item": None}
                return
        else:
            # Determine if we're dropping above or below the target item
            bbox = self.tree.bbox(target_item)
            if bbox:
                item_y = bbox[1]
                item_height = bbox[3]
                item_center = item_y + item_height / 2
                drop_below_target = event.y >= item_center

        # Determine target type (target_item may be None if drop_at_end)
        target_parent = self.tree.parent(target_item) if target_item else None

        if source_is_label:
            # Dragging label(s)
            if target_parent:
                # Target is also a label - get its indices
                try:
                    parts = target_item.split("_")
                    target_cat_idx = int(parts[1])
                    target_label_idx = int(parts[3])

                    # Only allow within same category
                    if target_cat_idx != source_cat_idx:
                        self._drag_data = {"item": None}
                        return

                    # If dropping below the target, we want to insert AFTER it
                    if drop_below_target:
                        target_label_idx += 1

                    level = self.category_levels[source_cat_idx]

                    if source_is_multi:
                        # Multi-label drag
                        source_label_indices = self._drag_data.get("label_indices", [])

                        # Don't drop on self (check original target)
                        orig_target = target_label_idx - 1 if drop_below_target else target_label_idx
                        if orig_target in source_label_indices:
                            self._drag_data = {"item": None}
                            return

                        # Extract selected labels in their current order
                        selected_labels = [level.labels[i] for i in sorted(source_label_indices)]

                        # Remove them (reverse order to maintain indices)
                        for i in sorted(source_label_indices, reverse=True):
                            level.labels.pop(i)

                        # Calculate adjusted target index after removal
                        removed_before = sum(1 for i in source_label_indices if i < target_label_idx)
                        adjusted_target = target_label_idx - removed_before

                        # Insert at target position
                        for i, label in enumerate(selected_labels):
                            level.labels.insert(adjusted_target + i, label)

                        self._recalculate_label_sort_orders(source_cat_idx)
                        self._expanded_categories[source_cat_idx] = True

                        # Select all moved labels at their new positions
                        new_selection = [f"cat_{source_cat_idx}_label_{adjusted_target + i}"
                                        for i in range(len(selected_labels))]
                        self._refresh_list()
                        self.tree.selection_set(*new_selection)
                        self._on_selection_changed(None)
                        self.main_tab.update_preview()
                    else:
                        # Single label drag
                        source_label_idx = self._drag_data.get("label_idx")

                        # Don't drop on self
                        if source_label_idx == target_label_idx or (drop_below_target and source_label_idx == target_label_idx - 1):
                            self._drag_data = {"item": None}
                            return

                        label = level.labels.pop(source_label_idx)
                        # Adjust target if we removed from before target
                        if source_label_idx < target_label_idx:
                            target_label_idx -= 1
                        level.labels.insert(target_label_idx, label)

                        self._recalculate_label_sort_orders(source_cat_idx)
                        self._expanded_categories[source_cat_idx] = True
                        new_selection = f"cat_{source_cat_idx}_label_{target_label_idx}"
                        self._refresh_list()
                        self.tree.selection_set(new_selection)
                        self._on_selection_changed(None)
                        self.main_tab.update_preview()
                except (IndexError, ValueError):
                    pass
            else:
                # Target is a category - labels stay in their category
                pass
        else:
            # Dragging a category (single or multi)
            if drop_at_end or not target_parent:
                # Target is also a category OR dropping at end
                try:
                    if drop_at_end:
                        # Drop at end - target index is after the last item
                        target_cat_idx = len(self.category_levels)
                    else:
                        target_cat_idx = int(target_item.split("_")[1])

                        # Don't drop onto calculated column
                        if self.category_levels[target_cat_idx].is_calculated:
                            self._drag_data = {"item": None}
                            return

                        # If dropping below the target, we want to insert AFTER it
                        if drop_below_target:
                            target_cat_idx += 1

                    source_is_multi = self._drag_data.get("is_multi", False)

                    if source_is_multi:
                        # Multi-category drag
                        source_cat_indices = self._drag_data.get("cat_indices", [])

                        # Don't drop on any of the selected items (check original target)
                        orig_target = target_cat_idx - 1 if drop_below_target and not drop_at_end else target_cat_idx
                        if orig_target in source_cat_indices:
                            self._drag_data = {"item": None}
                            return

                        # Save expanded states for all dragged categories
                        saved_expanded = {idx: self._expanded_categories.get(idx, False)
                                         for idx in source_cat_indices}

                        # Extract categories in their current order
                        selected_levels = [self.category_levels[i] for i in sorted(source_cat_indices)]

                        # Remove them (reverse order to maintain indices)
                        for i in sorted(source_cat_indices, reverse=True):
                            self.category_levels.pop(i)

                        # Calculate adjusted target index after removal
                        removed_before = sum(1 for i in source_cat_indices if i < target_cat_idx)
                        adjusted_target = target_cat_idx - removed_before

                        # Insert at target position
                        for i, level in enumerate(selected_levels):
                            self.category_levels.insert(adjusted_target + i, level)

                        # Update expanded states for new positions
                        self._expanded_categories = {}
                        for i, level in enumerate(selected_levels):
                            orig_idx = sorted(source_cat_indices)[i]
                            self._expanded_categories[adjusted_target + i] = saved_expanded.get(orig_idx, False)

                        # Select all moved categories at their new positions
                        new_selection = [f"cat_{adjusted_target + i}"
                                        for i in range(len(selected_levels))]
                        self._refresh_list()
                        self.tree.selection_set(*new_selection)
                        self._on_selection_changed(None)
                        self.main_tab.on_categories_reordered(self.category_levels)
                    else:
                        # Single category drag
                        source_cat_idx = self._drag_data.get("cat_idx")

                        # Don't drop on self (check if dropping on same position)
                        # When dropping below, target_cat_idx is already incremented
                        if source_cat_idx == target_cat_idx or (drop_below_target and source_cat_idx == target_cat_idx - 1):
                            self._drag_data = {"item": None}
                            return

                        # Save expanded state
                        was_expanded = self._expanded_categories.get(source_cat_idx, False)

                        # Reorder categories
                        level = self.category_levels.pop(source_cat_idx)

                        # Adjust target index if source was before target
                        if source_cat_idx < target_cat_idx:
                            target_cat_idx -= 1

                        self.category_levels.insert(target_cat_idx, level)

                        # Rebuild expanded states
                        new_expanded = {}
                        for idx, lv in enumerate(self.category_levels):
                            if lv == level:
                                new_expanded[idx] = was_expanded
                            else:
                                # Find original index
                                for old_idx, old_exp in self._expanded_categories.items():
                                    if old_idx != source_cat_idx:
                                        new_expanded[idx] = old_exp
                                        break
                        self._expanded_categories = {target_cat_idx: was_expanded}

                        new_selection = f"cat_{target_cat_idx}"
                        self._refresh_list()
                        self.tree.selection_set(new_selection)
                        self._on_selection_changed(None)
                        self.main_tab.on_categories_reordered(self.category_levels)
                except (IndexError, ValueError):
                    pass

        self._drag_data = {"item": None}

    def _get_selected_category_index(self) -> Optional[int]:
        """Get the index of the currently selected category (parent, not label)"""
        selection = self.tree.selection()
        if not selection:
            return None

        item_id = selection[0]
        # If it's a label item, get its parent category
        parent = self.tree.parent(item_id)
        if parent:
            item_id = parent

        # Extract index from item id (format: "cat_N")
        if item_id.startswith("cat_"):
            try:
                return int(item_id.split("_")[1])
            except (IndexError, ValueError):
                return None
        return None

    def _get_selected_label_index(self) -> Optional[Tuple[int, int]]:
        """Get (category_index, label_index) if a single label is selected, else None"""
        selection = self.tree.selection()
        if not selection:
            return None

        item_id = selection[0]
        # Check if it's a label (has parent)
        parent = self.tree.parent(item_id)
        if not parent:
            return None  # It's a category, not a label

        # Parse: "cat_N_label_M"
        try:
            parts = item_id.split("_")
            cat_idx = int(parts[1])
            label_idx = int(parts[3])
            return (cat_idx, label_idx)
        except (IndexError, ValueError):
            return None

    def _get_selected_labels(self) -> Optional[Tuple[int, List[int]]]:
        """
        Get all selected labels if they are all from the same category.
        Returns (category_index, [label_indices]) sorted by label index, or None.
        """
        selection = self.tree.selection()
        if not selection:
            return None

        cat_idx = None
        label_indices = []

        for item_id in selection:
            # Check if it's a label (has parent)
            parent = self.tree.parent(item_id)
            if not parent:
                # It's a category, not a label - mixed selection not supported
                return None

            # Parse: "cat_N_label_M"
            try:
                parts = item_id.split("_")
                this_cat_idx = int(parts[1])
                label_idx = int(parts[3])

                # All labels must be from same category
                if cat_idx is None:
                    cat_idx = this_cat_idx
                elif cat_idx != this_cat_idx:
                    return None  # Mixed categories not supported

                label_indices.append(label_idx)
            except (IndexError, ValueError):
                return None

        if cat_idx is not None and label_indices:
            return (cat_idx, sorted(label_indices))
        return None

    def _get_selected_categories(self) -> Optional[List[int]]:
        """
        Get all selected categories (not labels).
        Returns list of category indices sorted by index, or None if mixed selection.
        """
        selection = self.tree.selection()
        if not selection:
            return None

        cat_indices = []

        for item_id in selection:
            # Check if it's a category (has no parent)
            parent = self.tree.parent(item_id)
            if parent:
                # It's a label, not a category - mixed selection not supported
                return None

            # Parse: "cat_N"
            try:
                parts = item_id.split("_")
                cat_idx = int(parts[1])
                cat_indices.append(cat_idx)
            except (IndexError, ValueError):
                return None

        if cat_indices:
            return sorted(cat_indices)
        return None

    def _is_label_selected(self) -> bool:
        """Check if a label (child item) is selected"""
        return self._get_selected_label_index() is not None

    def _on_selection_changed(self, event):
        """Handle treeview selection change - context-aware button states

        Note: X, ^, v buttons (remove, move_up, move_down) stay always enabled
        per user request. They will silently do nothing if clicked without
        a valid selection.
        """
        selection = self.tree.selection()
        if not selection:
            # Nothing selected - disable text buttons, X/^/v stay enabled
            self.add_label_btn.set_enabled(False)
            self.rename_btn.set_enabled(False)
            return

        # Check for multi-select labels first
        multi_labels = self._get_selected_labels()
        if multi_labels and len(multi_labels[1]) > 1:
            # Multiple labels selected from same category
            # Add Label - enabled (adds to this category)
            self.add_label_btn.set_enabled(True)
            # Rename - disabled for multi-select
            self.rename_btn.set_enabled(False)
            return

        # Single selection handling
        label_info = self._get_selected_label_index()
        cat_idx = self._get_selected_category_index()

        if label_info:
            # A single LABEL is selected
            # Add Label button - enabled (adds to this category)
            self.add_label_btn.set_enabled(True)
            # Rename - enabled for labels
            self.rename_btn.set_enabled(True)

        elif cat_idx is not None and cat_idx < len(self.category_levels):
            # A CATEGORY is selected
            level = self.category_levels[cat_idx]

            if level.is_calculated:
                # Calculated columns are read-only
                self.add_label_btn.set_enabled(False)
                self.rename_btn.set_enabled(False)
            else:
                # Normal category
                self.add_label_btn.set_enabled(True)
                self.rename_btn.set_enabled(True)
        else:
            # Invalid selection
            self.add_label_btn.set_enabled(False)
            self.rename_btn.set_enabled(False)

    def _on_tree_double_click(self, event):
        """Handle double-click on tree item - rename label only (not categories)

        Categories use expand/collapse on click, so double-click would
        accidentally trigger rename. Use Rename button or F2 for categories.
        """
        item = self.tree.identify_row(event.y)
        if not item:
            return

        # Check if it's a category (parent) item
        parent = self.tree.parent(item)
        if parent == "":
            # It's a category - don't rename on double-click since that
            # conflicts with expand/collapse behavior. Use Rename button instead.
            return
        else:
            # It's a label - rename it (labels don't expand/collapse)
            self._on_rename_item()

    def _on_f2_rename(self, event):
        """Handle F2 key to rename selected category or label"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        parent = self.tree.parent(item)
        if parent == "":
            # It's a category
            self._on_rename_column()
        else:
            # It's a label
            self._on_rename_item()

    def _refresh_list(self):
        """Refresh category columns display with expandable labels"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, level in enumerate(self.category_levels):
            label_count = len(level.labels)
            label_text = f"({label_count} label{'s' if label_count != 1 else ''})"

            # Show calculated columns with indicator
            if level.is_calculated:
                display_text = f"{idx + 1}. {level.name} [Calc Column - read only]"
            else:
                display_text = f"{idx + 1}. {level.name} {label_text}"

            # Create parent item for category
            cat_id = f"cat_{idx}"
            self.tree.insert("", tk.END, iid=cat_id, text=display_text, open=self._expanded_categories.get(idx, True))

            # Add labels as children with order numbers
            if level.labels and not level.is_calculated:
                for label_idx, label in enumerate(level.labels):
                    label_id = f"cat_{idx}_label_{label_idx}"
                    # Show order number for each label
                    self.tree.insert(cat_id, tk.END, iid=label_id,
                                   text=f"    {label_idx + 1}. {label}", tags=("label",))

        # Configure tag for label items (same color as categories for readability)
        colors = self._theme_manager.colors
        self.tree.tag_configure("label", foreground=colors['text_primary'])

        # Force UI update to ensure tree is visually refreshed
        self.tree.update_idletasks()

        # Update placeholder visibility
        self._update_empty_state()

    def set_category_levels(self, levels: List[CategoryLevel]):
        """Set category levels"""
        self.category_levels = levels
        self._refresh_list()

    def load_categories(self, levels: List[CategoryLevel]):
        """Load category levels (alias for set_category_levels)"""
        self.set_category_levels(levels)

    def clear(self):
        """Clear all category levels"""
        self.category_levels.clear()
        self._expanded_categories.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._update_empty_state()

    def _update_empty_state(self):
        """Show or hide placeholder based on whether tree has content"""
        if hasattr(self, '_placeholder_label'):
            if self.tree.get_children():
                # Tree has items - hide placeholder
                self._placeholder_label.lower()
            else:
                # Tree is empty - show placeholder
                self._placeholder_label.lift()

    def set_enabled(self, enabled: bool):
        """Enable/disable panel"""
        self.add_btn.set_enabled(enabled)

    def on_theme_changed(self):
        """Update widget colors when theme changes"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        # Use 'section_bg' for inner content areas (matches Section.TFrame style)
        content_bg = colors['background']
        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Update inner frames
        for frame in self._inner_frames:
            frame.config(bg=content_bg)

        # Update arrow frame and button container
        if hasattr(self, '_arrow_frame'):
            self._arrow_frame.config(bg=content_bg)
        if hasattr(self, '_btn_container'):
            self._btn_container.config(bg=content_bg)

        # Update secondary buttons
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

        # Update drop indicator if it exists
        if self._drop_indicator:
            self._drop_indicator.config(bg=colors['button_primary'])

        # Update treeview label tag color
        self.tree.tag_configure("label", foreground=colors['text_primary'])

        # Update ThemedScrollbar
        if hasattr(self, '_tree_scrollbar'):
            self._tree_scrollbar.on_theme_changed()

        # Update tree container border - use section_bg to match treeview style
        if hasattr(self, '_tree_container'):
            tree_border = colors.get('border', '#3a3a4a')
            tree_bg = colors.get('section_bg', colors.get('surface', content_bg))
            self._tree_container.config(
                bg=tree_bg,
                highlightbackground=tree_border,
                highlightcolor=tree_border
            )

        # Update placeholder label - use same tree_bg as container
        if hasattr(self, '_placeholder_label'):
            tree_bg = colors.get('section_bg', colors.get('surface', content_bg))
            self._placeholder_label.config(
                fg=colors['text_muted'],
                bg=tree_bg
            )

        # Force ttk.Frame to re-apply style after theme change
        if hasattr(self, '_content_wrapper'):
            self._content_wrapper.configure(style='Section.TFrame')

        # Update section header widgets
        self._update_section_header_theme()

        # Update treeview selection colors
        self._configure_treeview_style()

    def _configure_treeview_style(self):
        """Configure treeview selection colors based on current theme"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        style = ttk.Style()

        # Use section background instead of pure white for treeview background (match Available Fields)
        tree_bg = colors.get('section_bg', colors.get('surface', colors['background']))
        text_color = colors.get('text_primary', '#000000')

        # Configure Category.Treeview background and text colors
        style.configure("Category.Treeview",
                        background=tree_bg,
                        fieldbackground=tree_bg,
                        foreground=text_color)

        # CRITICAL: Remove border from Treeview layout - style.configure doesn't work for this
        # The border is built into the layout elements, must be removed via layout()
        style.layout("Category.Treeview", [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

        # Use selection_highlight to match Available Fields panel selection color
        selection_bg = colors.get('selection_highlight', colors.get('card_surface', colors['background']))
        selection_fg = '#ffffff' if is_dark else colors.get('text_primary', '#000000')

        # Configure Category.Treeview selection colors
        style.map("Category.Treeview",
                  background=[("selected", selection_bg)],
                  foreground=[("selected", selection_fg)])
