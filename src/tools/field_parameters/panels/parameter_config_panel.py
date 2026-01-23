"""
ParameterConfigPanel
Panel component for the Field Parameters tool.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import List, TYPE_CHECKING
import logging

from tools.field_parameters.widgets import SVGCheckbox
from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, ThemedMessageBox, LabeledRadioGroup
from tools.field_parameters.panels.panel_base import SectionPanelMixin

if TYPE_CHECKING:
    from tools.field_parameters.field_parameters_ui import FieldParametersTab


class ParameterConfigPanel(SectionPanelMixin, ttk.LabelFrame):
    """Panel for parameter configuration (new/edit, name, options)"""

    def __init__(self, parent, main_tab: 'FieldParametersTab'):
        ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe', padding="12")
        self._init_section_panel_mixin()
        self.main_tab = main_tab
        self.logger = logging.getLogger(__name__)

        # Button tracking for theme updates
        self._primary_buttons = []

        # Create and set the section header labelwidget
        self._create_section_header(parent, "Parameter Configuration", "cogwheel")

        self.mode = tk.StringVar(value="new")  # "new" or "edit"
        self.parameter_name = tk.StringVar()
        self.selected_parameter = tk.StringVar()
        self.keep_lineage = tk.BooleanVar(value=False)

        self.setup_ui()

    def setup_ui(self):
        """Setup configuration panel UI"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        # Use 'background' (darkest) for inner content frames, not section_bg
        # This creates proper contrast: section border (section_bg) vs content (background)
        content_bg = colors['background']
        fg_color = colors['text_primary']
        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # CRITICAL: Inner wrapper with Section.TFrame style fills LabelFrame with correct background
        # This is required per UI_DESIGN_PATTERNS.md - without it, section_bg shows in padding area
        self._content_wrapper = ttk.Frame(self, style='Section.TFrame', padding="15")
        self._content_wrapper.pack(fill=tk.BOTH, expand=True)

        # Track inner frames for theme updates
        self._inner_frames = []

        # Mode selection row
        self.mode_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        self.mode_frame.pack(fill=tk.X, pady=(0, 8))
        self._inner_frames.append(self.mode_frame)

        self.mode_label = tk.Label(self.mode_frame, text="Mode:",
                                   font=('Segoe UI', 9, 'bold'),
                                   bg=content_bg, fg=fg_color)
        self.mode_label.pack(side=tk.LEFT, padx=(0, 10))

        # Radio Group for mode selection
        self.mode_radio_group = LabeledRadioGroup(
            self.mode_frame,
            variable=self.mode,
            options=[
                ("new", "Create New"),
                ("edit", "Edit Existing"),
            ],
            command=self._on_mode_changed,
            orientation="horizontal",
            padding=15
        )
        self.mode_radio_group.pack(side=tk.LEFT)

        # Options (keep lineage) - SVG Checkbox
        self.lineage_checkbox = SVGCheckbox(
            self._content_wrapper,
            text="Keep Lineage Tags (only for existing parameters)",
            variable=self.keep_lineage,
            state="disabled"  # Disabled by default in "new" mode
        )
        self.lineage_checkbox.pack(anchor="w", pady=(0, 8))

        # New parameter section
        self.new_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        self.new_frame.pack(fill=tk.X, pady=(0, 5))
        self._inner_frames.append(self.new_frame)

        self._param_name_label = tk.Label(
            self.new_frame, text="Parameter Name:",
            font=('Segoe UI', 9),
            bg=content_bg, fg=fg_color
        )
        self._param_name_label.pack(side=tk.LEFT)
        # Configure compact entry style to match combobox height (padding 6,3 vs default 10,8)
        style = ttk.Style()
        style.configure('Compact.TEntry', padding=(6, 3))

        self.name_entry = ttk.Entry(
            self.new_frame,
            textvariable=self.parameter_name,
            width=30,
            style='Compact.TEntry'
        )
        self.name_entry.pack(side=tk.LEFT, padx=(5, 8), fill=tk.X, expand=True)

        # Create button (RoundedButton) - height matches REFRESH/CONNECT/CLOUD buttons
        self.create_btn = RoundedButton(
            self.new_frame,
            text="CREATE",
            command=self._on_create,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=content_bg
        )
        self.create_btn.pack(side=tk.LEFT)
        self._primary_buttons.append(self.create_btn)

        # Edit parameter section
        self.edit_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        # Don't pack yet - will show based on mode
        self._inner_frames.append(self.edit_frame)

        self._select_param_label = tk.Label(
            self.edit_frame, text="Select Parameter:",
            font=('Segoe UI', 9),
            bg=content_bg, fg=fg_color
        )
        self._select_param_label.pack(side=tk.LEFT)
        self.param_combo = ttk.Combobox(
            self.edit_frame,
            textvariable=self.selected_parameter,
            state="readonly",
            width=35
        )
        self.param_combo.pack(side=tk.LEFT, padx=(5, 8), fill=tk.X, expand=True)

        # Load button (RoundedButton) - height matches REFRESH/CONNECT/CLOUD buttons
        self.load_btn = RoundedButton(
            self.edit_frame,
            text="LOAD",
            command=self._on_load,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=content_bg
        )
        self.load_btn.pack(side=tk.LEFT)
        self._primary_buttons.append(self.load_btn)

        # Set initial mode
        self._on_mode_changed()
    
    def _on_mode_changed(self):
        """Handle mode change between new/edit"""
        if self.mode.get() == "new":
            self.edit_frame.pack_forget()
            self.new_frame.pack(fill=tk.X, pady=(0, 5))
            self.keep_lineage.set(False)
            self.lineage_checkbox.config(state="disabled")
        else:
            self.new_frame.pack_forget()
            self.edit_frame.pack(fill=tk.X, pady=(0, 5))
            self.keep_lineage.set(True)
            self.lineage_checkbox.config(state="normal")
    
    def _on_create(self):
        """Create new parameter"""
        param_name = self.parameter_name.get().strip()
        if not param_name:
            ThemedMessageBox.showwarning(self, "Missing Name", "Please enter a parameter name")
            return

        self.main_tab.on_create_new_parameter(param_name, self.keep_lineage.get())
    
    def _on_load(self):
        """Load existing parameter"""
        param_name = self.selected_parameter.get()
        if not param_name:
            ThemedMessageBox.showwarning(self, "No Selection", "Please select a parameter to edit")
            return

        self.main_tab.on_edit_existing_parameter(param_name, self.keep_lineage.get())
    
    def update_parameter_list(self, parameters: List[str]):
        """Update the list of available parameters"""
        self.param_combo['values'] = parameters
        if parameters:
            self.param_combo.current(0)
    
    def set_enabled(self, enabled: bool):
        """Enable/disable panel controls"""
        state = "normal" if enabled else "disabled"
        self.name_entry.config(state=state)
        self.create_btn.set_enabled(enabled)
        self.param_combo.config(state="readonly" if enabled else "disabled")
        self.load_btn.set_enabled(enabled)
    
    def clear(self):
        """Clear panel state"""
        self.parameter_name.set("")
        self.selected_parameter.set("")
        self.param_combo['values'] = []
        # Reset to Create New mode
        self.mode.set("new")
        self._on_mode_changed()

    def on_theme_changed(self):
        """Update widget colors when theme changes"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        # Use 'section_bg' for inner content areas (matches Section.TFrame style)
        content_bg = colors['background']
        fg_color = colors['text_primary']
        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Update inner frames
        for frame in self._inner_frames:
            frame.config(bg=content_bg)

        # Update mode label
        self.mode_label.config(bg=content_bg, fg=fg_color)

        # Update parameter name label
        self._param_name_label.config(bg=content_bg, fg=fg_color)

        # Update select parameter label
        self._select_param_label.config(bg=content_bg, fg=fg_color)

        # Update SVG radio group
        self.mode_radio_group.on_theme_changed()

        # Update SVG checkbox
        self.lineage_checkbox.on_theme_changed(bg=content_bg)

        # Update primary buttons
        for btn in self._primary_buttons:
            btn.update_colors(
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg=colors.get('button_text', '#ffffff'),
                disabled_bg=disabled_bg,
                disabled_fg=disabled_fg,
                canvas_bg=content_bg
            )

        # Force ttk.Frame to re-apply style after theme change
        if hasattr(self, '_content_wrapper'):
            self._content_wrapper.configure(style='Section.TFrame')

        # Update section header widgets
        self._update_section_header_theme()


