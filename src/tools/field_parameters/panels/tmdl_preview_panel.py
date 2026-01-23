"""
TmdlPreviewPanel
Panel component for the Field Parameters tool.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import TYPE_CHECKING
import logging

from tools.field_parameters.widgets import SVGCheckbox
from core.theme_manager import get_theme_manager
from core.ui_base import ThemedScrollbar
from core.ui import Tooltip
from tools.field_parameters.panels.panel_base import SectionPanelMixin

if TYPE_CHECKING:
    from tools.field_parameters.field_parameters_ui import FieldParametersTab


class TmdlPreviewPanel(SectionPanelMixin, ttk.LabelFrame):
    """Panel showing generated TMDL code"""

    def __init__(self, parent, main_tab: 'FieldParametersTab'):
        ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe', padding="12")
        self._init_section_panel_mixin()
        self.main_tab = main_tab
        self.logger = logging.getLogger(__name__)

        # Create and set the section header labelwidget
        self._create_section_header(parent, "Generated TMDL Code", "export")

        self.setup_ui()

    def setup_ui(self):
        """Setup preview panel UI"""
        colors = self._theme_manager.colors
        # Use 'background' (darkest) for inner content frames, not section_bg
        content_bg = colors['background']

        # CRITICAL: Inner wrapper with Section.TFrame style fills LabelFrame with correct background
        # This is required per UI_DESIGN_PATTERNS.md - without it, section_bg shows in padding area
        self._content_wrapper = ttk.Frame(self, style='Section.TFrame', padding="15")
        self._content_wrapper.pack(fill=tk.BOTH, expand=True)

        # Track inner frames for theme updates
        self._inner_frames = []

        # Store tooltips for reference
        self._tooltips = []

        # TOP ROW - formatting options
        format_toolbar = tk.Frame(self._content_wrapper, bg=content_bg)
        format_toolbar.pack(fill=tk.X, pady=(0, 8))
        self._inner_frames.append(format_toolbar)

        # Group by category checkbox - SVG styled
        self.group_by_category_var = tk.BooleanVar(value=False)
        self.group_checkbox = SVGCheckbox(
            format_toolbar,
            text="Group by category:",
            variable=self.group_by_category_var,
            command=self._on_format_changed,
            state="disabled"  # Disabled initially until categories exist
        )
        self.group_checkbox.pack(side=tk.LEFT)
        # Add tooltip for "adds blank lines between groups" - bind to .frame since SVGCheckbox isn't a tk widget
        self._tooltips.append(Tooltip(self.group_checkbox.frame, "Adds blank lines between groups"))

        # Category column selector
        self.group_by_column_var = tk.StringVar()
        self.group_by_combo = ttk.Combobox(
            format_toolbar,
            textvariable=self.group_by_column_var,
            state="disabled",
            width=20
        )
        self.group_by_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.group_by_combo.bind("<<ComboboxSelected>>", lambda e: self._on_format_changed())

        # Separator
        ttk.Separator(format_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=(12, 12))

        # Update field order checkbox - SVG styled
        self.update_field_order_var = tk.BooleanVar(value=True)
        self.update_order_checkbox = SVGCheckbox(
            format_toolbar,
            text="Update field order",
            variable=self.update_field_order_var,
            command=self._on_format_changed
        )
        self.update_order_checkbox.pack(side=tk.LEFT)
        # Add tooltip - capitalize "Uncheck" - bind to .frame since SVGCheckbox isn't a tk widget
        self._tooltips.append(Tooltip(self.update_order_checkbox.frame, "Uncheck to preserve custom sort"))

        # Separator
        ttk.Separator(format_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=(12, 12))

        # Show uncategorized as blank checkbox - SVG styled
        self.show_unassigned_blank_var = tk.BooleanVar(value=False)
        self.show_unassigned_checkbox = SVGCheckbox(
            format_toolbar,
            text="Show uncategorized as BLANK",
            variable=self.show_unassigned_blank_var,
            command=self._on_format_changed
        )
        self.show_unassigned_checkbox.pack(side=tk.LEFT)
        # Add tooltip - bind to .frame since SVGCheckbox isn't a tk widget
        self._tooltips.append(Tooltip(self.show_unassigned_checkbox.frame, 'When checked, outputs "" instead of "Uncategorized"'))

        # MIDDLE - Text widget with scrollbar
        border_color = colors.get('border', '#3a3a4a')
        # Use section_bg to match Category Columns panel style
        tree_bg = colors.get('section_bg', colors.get('surface', content_bg))
        text_frame = tk.Frame(
            self._content_wrapper, bg=tree_bg,
            highlightbackground=border_color, highlightcolor=border_color,
            highlightthickness=1,
            bd=0, relief=tk.FLAT  # Prevent 3D bevel effect
        )
        text_frame.pack(fill=tk.BOTH, expand=True)  # No extra padding - content_wrapper has padding="15"
        self._inner_frames.append(text_frame)
        self._text_frame = text_frame  # Store for theme updates

        # Use grid layout for proper scrollbar corner handling
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)

        # Text widget with theme colors
        text_bg = colors.get('section_bg', colors.get('surface', colors['background']))
        text_fg = colors.get('text_primary', '#000000')

        self.text_widget = tk.Text(
            text_frame,
            wrap=tk.NONE,
            font=("Consolas", 9),
            height=15,
            bg=text_bg,
            fg=text_fg,
            insertbackground=text_fg,  # Cursor color
            selectbackground=colors.get('card_surface', '#3A3A3A'),
            selectforeground=text_fg,
            bd=0, relief=tk.FLAT, highlightthickness=0,  # Remove all borders from text widget
            padx=5, pady=5  # 5px internal padding so code doesn't touch borders
        )
        self.text_widget.grid(row=0, column=0, sticky='nsew')

        # ThemedScrollbar for vertical - auto_hide=False because it uses pack which conflicts with grid
        self._y_scrollbar = ThemedScrollbar(
            text_frame,
            command=self.text_widget.yview,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=False
        )
        self._y_scrollbar.grid(row=0, column=1, sticky='ns')

        # Horizontal scrollbar with auto-hide using grid_remove/grid
        self._x_scrollbar = ttk.Scrollbar(
            text_frame,
            orient=tk.HORIZONTAL,
            command=self.text_widget.xview
        )
        self._x_scrollbar.grid(row=1, column=0, sticky='ew')
        self._x_scrollbar_visible = True  # Track visibility state

        # Connect text widget to scrollbars with auto-hide for horizontal
        self.text_widget.config(yscrollcommand=self._y_scrollbar.set, xscrollcommand=self._on_xscroll)
        self._y_scrollbar._command = self.text_widget.yview

        # Make read-only
        self.text_widget.config(state=tk.DISABLED)

        # Placeholder overlay - shown when text is empty
        self._placeholder_label = tk.Label(
            text_frame,
            text="No parameter created yet.\n\nCreate a new parameter or load an existing one\nto see the TMDL code here.",
            font=("Segoe UI", 10, "italic"),
            fg=colors['text_muted'],
            bg=text_bg,
            justify="center"
        )
        # Use place to overlay on text widget - initially visible
        self._placeholder_label.place(relx=0.5, rely=0.5, anchor="center")

        # Note: Action buttons (Apply, Copy, Reset) are now in the main field_parameters_ui.py
        # outside this panel frame, following the standard centered button bar pattern

    def set_tmdl_code(self, code: str):
        """Set TMDL code in preview"""
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.delete("1.0", tk.END)
        self.text_widget.insert("1.0", code)
        self.text_widget.config(state=tk.DISABLED)
        # Hide placeholder when content is set
        if hasattr(self, '_placeholder_label'):
            self._placeholder_label.place_forget()
    
    def get_tmdl_code(self) -> str:
        """Get current TMDL code"""
        return self.text_widget.get("1.0", tk.END).strip()
    
    def set_error(self, error_msg: str):
        """Show error message in preview"""
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.delete("1.0", tk.END)
        self.text_widget.insert("1.0", f"ERROR:\n\n{error_msg}")
        self.text_widget.config(state=tk.DISABLED)
        # Hide placeholder when showing error
        if hasattr(self, '_placeholder_label'):
            self._placeholder_label.place_forget()
    
    def clear(self):
        """Clear preview and show placeholder overlay"""
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.delete("1.0", tk.END)
        self.text_widget.config(state=tk.DISABLED)
        # Show placeholder overlay
        if hasattr(self, '_placeholder_label'):
            self._placeholder_label.place(relx=0.5, rely=0.5, anchor="center")
    
    def _on_xscroll(self, first, last):
        """Handle horizontal scroll with auto-hide - shows/hides scrollbar based on content width"""
        # Update scrollbar position
        self._x_scrollbar.set(first, last)

        # Check if content needs horizontal scrolling (visible region < 100%)
        needs_scroll = (float(last) - float(first)) < 0.999

        # Show/hide horizontal scrollbar based on need
        if needs_scroll and not self._x_scrollbar_visible:
            self._x_scrollbar.grid()  # Show
            self._x_scrollbar_visible = True
        elif not needs_scroll and self._x_scrollbar_visible:
            self._x_scrollbar.grid_remove()  # Hide
            self._x_scrollbar_visible = False

    def _on_format_changed(self):
        """Handle format option change - update the parameter and regenerate preview"""
        if not self.main_tab.current_parameter:
            return

        # Update parameter with current format settings
        self.main_tab.current_parameter.group_by_category = self.group_by_category_var.get()
        self.main_tab.current_parameter.update_field_order = self.update_field_order_var.get()
        self.main_tab.current_parameter.show_unassigned_as_blank = self.show_unassigned_blank_var.get()

        # Update combo state
        if self.group_by_category_var.get():
            self.group_by_combo.config(state="normal")
            self.group_by_combo.bind('<Key>', lambda e: 'break')  # Block keyboard input
            # Get selected category index
            selected = self.group_by_column_var.get()
            if selected:
                for idx, level in enumerate(self.main_tab.current_parameter.category_levels):
                    if level.name == selected:
                        self.main_tab.current_parameter.group_by_category_index = idx
                        break
        else:
            self.group_by_combo.config(state="disabled")

        # Regenerate preview
        self.main_tab.update_preview()

    def update_category_columns(self, category_levels: list):
        """Update the category column dropdown with available columns"""
        if not category_levels:
            self.group_by_combo['values'] = []
            self.group_by_column_var.set("")
            self.group_by_combo.config(state="disabled")
            self.group_checkbox.config(state="disabled")
            return

        # Populate combo with category column names
        column_names = [level.name for level in category_levels]
        self.group_by_combo['values'] = column_names

        # Select first column by default if nothing selected
        if not self.group_by_column_var.get() and column_names:
            self.group_by_column_var.set(column_names[0])

        # Enable checkbox
        self.group_checkbox.config(state="normal")

        # Update combo state based on checkbox
        if self.group_by_category_var.get():
            self.group_by_combo.config(state="normal")
            self.group_by_combo.bind('<Key>', lambda e: 'break')  # Block keyboard input
        else:
            self.group_by_combo.config(state="disabled")

    def set_enabled(self, enabled: bool):
        """Enable/disable panel controls (checkboxes/combo only - buttons are in main UI)"""
        # Note: Buttons are now managed by field_parameters_ui.py
        pass

    def on_theme_changed(self):
        """Update widget colors when theme changes"""
        colors = self._theme_manager.colors
        # Use 'section_bg' for inner content areas (matches Section.TFrame style)
        content_bg = colors['background']

        # Update inner frames
        for frame in self._inner_frames:
            frame.config(bg=content_bg)

        # Update SVG checkboxes
        self.group_checkbox.on_theme_changed(bg=content_bg)
        self.update_order_checkbox.on_theme_changed(bg=content_bg)
        self.show_unassigned_checkbox.on_theme_changed(bg=content_bg)

        # Update ThemedScrollbar
        if hasattr(self, '_y_scrollbar'):
            self._y_scrollbar.on_theme_changed()

        # Update text frame border - use border color to match Category Columns panel
        if hasattr(self, '_text_frame'):
            border_color = colors.get('border', '#3a3a4a')
            tree_bg = colors.get('section_bg', colors.get('surface', content_bg))
            self._text_frame.config(
                bg=tree_bg,
                highlightbackground=border_color,
                highlightcolor=border_color
            )

        # Update text widget colors (match Available Fields style)
        if hasattr(self, 'text_widget'):
            text_bg = colors.get('section_bg', colors.get('surface', colors['background']))
            text_fg = colors.get('text_primary', '#000000')
            self.text_widget.config(
                bg=text_bg,
                fg=text_fg,
                insertbackground=text_fg,
                selectbackground=colors.get('card_surface', '#3A3A3A'),
                selectforeground=text_fg
            )

        # Update placeholder label
        if hasattr(self, '_placeholder_label'):
            text_bg = colors.get('section_bg', colors.get('surface', colors['background']))
            self._placeholder_label.config(
                fg=colors['text_muted'],
                bg=text_bg
            )

        # Update combobox style
        style = ttk.Style()
        combo_bg = colors.get('section_bg', colors.get('surface', colors['background']))
        style.configure('TCombobox',
                        fieldbackground=combo_bg,
                        background=combo_bg,
                        foreground=colors['text_primary'],
                        arrowcolor=colors['text_primary'])
        style.map('TCombobox',
                  fieldbackground=[('readonly', combo_bg), ('disabled', combo_bg)],
                  foreground=[('readonly', colors['text_primary']), ('disabled', colors['text_muted'])])
        # Force combobox to re-apply style
        if hasattr(self, 'group_by_combo'):
            self.group_by_combo.configure(style='TCombobox')

        # Force ttk.Frame to re-apply style after theme change
        if hasattr(self, '_content_wrapper'):
            self._content_wrapper.configure(style='Section.TFrame')

        # Update section header widgets
        self._update_section_header_theme()
