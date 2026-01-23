"""
Table Column Widths UI - User interface implementation
Built by Reid Havens of Analytic Endeavors

Provides the user interface for the Table Column Widths tool.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from typing import Dict, List, Any, Optional
import threading

from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin, RoundedButton, SquareIconButton, ThemedScrollbar, SplitLogSection, FileInputSection
from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from tools.column_width.column_width_core import (
    TableColumnWidthsEngine, WidthConfiguration, WidthPreset, 
    VisualInfo, FieldType, VisualType, FieldInfo
)


class TableColumnWidthsTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """
    Table Column Widths tool UI tab
    """
    
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "column_width", "Table Column Widths")
        
        # Initialize components
        self.engine: Optional[TableColumnWidthsEngine] = None
        self.visuals_info: List[VisualInfo] = []
        
        # UI State
        self.file_section = None  # FileInputSection instance
        self.pbip_path_var = None  # Set from file_section.path_var
        self.scan_button = None  # Set from file_section.action_button
        self.config = WidthConfiguration()
        
        # Visual selection variables
        self.visual_selection_vars: Dict[str, tk.BooleanVar] = {}
        self.visual_config_vars: Dict[str, Dict[str, tk.Variable]] = {}  # Per-visual configuration
        self.visual_tree = None
        self.current_selected_visual = None  # Track currently selected visual for per-visual config
        
        # Global configuration variables - defaulting to Fit to Header for categorical, Fit to Totals for measures
        self.global_categorical_preset_var = tk.StringVar(value=WidthPreset.AUTO_FIT.value)
        self.global_measure_preset_var = tk.StringVar(value=WidthPreset.FIT_TO_TOTALS.value)
        
        # Custom width variables
        self.categorical_custom_var = tk.StringVar(value="105")
        self.measure_custom_var = tk.StringVar(value="95")
        
        # Current active configuration (global by default, per-visual when selected)
        self.categorical_preset_var = self.global_categorical_preset_var
        self.measure_preset_var = self.global_measure_preset_var

        # Secondary buttons for theme updates (preview, reset, etc.)
        self._secondary_buttons = []

        # Preset rows for custom radio buttons (theme updates)
        self._preset_rows = {'categorical': [], 'measure': []}

        # Icon references list to prevent garbage collection
        self._icon_refs = []

        # Load button icons
        self._load_button_icons()

        # Setup UI
        self.setup_ui()
        self.setup_path_cleaning(self.pbip_path_var)
        self._show_welcome_message()
    
    def setup_ui(self) -> None:
        """Setup the Table Column Widths UI"""
        # IMPORTANT: Create action buttons FIRST with side=BOTTOM to ensure they're always visible
        self._setup_action_buttons()

        # File input section (now includes SCAN VISUALS button)
        self._setup_file_input_section()

        # Middle content: Two columns side by side (Configuration left, Visual Selection right)
        # Using 1:2 weight ratio - Column Width Config gets 1/3, Visual Selection gets 2/3
        # Use tk.Frame (not ttk) for more stable geometry during theme changes
        # CRITICAL: expand=False prevents height growth; weight=0 on row prevents vertical stretch
        middle_content = tk.Frame(self.frame, bg=self._theme_manager.colors['background'])
        middle_content.pack(fill=tk.X, expand=False)  # fill=X only, NOT BOTH, prevents height expansion
        middle_content.columnconfigure(0, weight=1)  # Initial 1/3 for Column Width Configuration
        middle_content.columnconfigure(1, weight=2)  # Initial 2/3 for Visual Selection
        middle_content.rowconfigure(0, weight=0)  # weight=0 prevents vertical expansion
        self._middle_content = middle_content  # Store for theme updates
        self._columns_width_locked = False  # Track if column widths have been locked

        # LEFT COLUMN: Column Width Configuration (no padx - use internal padding instead)
        left_column = tk.Frame(middle_content, bg=self._theme_manager.colors['background'])
        left_column.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._left_column = left_column  # Store for theme updates
        self._setup_configuration_section(left_column)

        # RIGHT COLUMN: Visual Selection (no padx - use internal padding instead)
        right_column = tk.Frame(middle_content, bg=self._theme_manager.colors['background'])
        right_column.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._right_column = right_column  # Store for theme updates
        self._setup_visual_selection_section(right_column)

        # Lock column widths after initial layout is complete to prevent theme toggle wiggle
        self.frame.after_idle(self._lock_column_widths)

        # Progress bar (full width, before log section)
        self.create_progress_bar(self.frame)

        # FULL WIDTH: Analysis & Progress at bottom (reduced top padding)
        self.log_section = SplitLogSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="Analysis & Progress",
            section_icon="analyze",
            summary_title="Column Summary",
            summary_icon="bar-chart",
            log_title="Progress Log",
            log_icon="log-file",
            summary_placeholder="Analyze a report to see column details"
        )
        self.log_section.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        # Connect log_text for log_message() calls
        self.log_text = self.log_section.log_text

    def _lock_column_widths(self):
        """Lock column widths AND height after initial layout to prevent layout drift.
        Called via after_idle to ensure geometry is finalized."""
        if not hasattr(self, '_middle_content') or not self._middle_content:
            return
        if self._columns_width_locked:
            return

        try:
            # Get actual rendered dimensions
            left_width = self._left_column.winfo_width()
            right_width = self._right_column.winfo_width()
            middle_height = self._middle_content.winfo_height()

            # Only lock if we have valid dimensions (widget is mapped and has size)
            if left_width > 10 and right_width > 10 and middle_height > 10:
                # CRITICAL: Set weight=0 AND minsize to truly lock the widths
                # weight=0 prevents proportional expansion
                # minsize sets both minimum AND effective fixed width when weight=0
                self._middle_content.columnconfigure(0, weight=0, minsize=left_width)
                self._middle_content.columnconfigure(1, weight=0, minsize=right_width)

                # ALSO set explicit width on the frame widgets themselves for extra stability
                self._left_column.configure(width=left_width)
                self._right_column.configure(width=right_width)

                # Prevent the frames from resizing based on their children
                self._left_column.grid_propagate(False)
                self._right_column.grid_propagate(False)

                # CRITICAL: Lock the HEIGHT of middle_content to prevent vertical expansion
                self._middle_content.configure(height=middle_height)
                self._middle_content.pack_propagate(False)  # Prevent pack children from affecting size
                self._middle_content.grid_propagate(False)  # Prevent grid children from affecting size

                self._locked_left_width = left_width
                self._locked_right_width = right_width
                self._locked_middle_height = middle_height
                self._columns_width_locked = True
        except Exception:
            pass

    def _load_button_icons(self):
        """Load SVG icons for buttons and section headers"""
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}

        # 16px icons for buttons and section headers
        icon_names_16 = [
            "folder", "magnifying-glass", "bar-chart", "filter", "warning",
            "eye", "execute", "reset", "save", "eraser", "question", "analyze",
            "Power-BI", "table", "table column widths", "earth"
        ]
        for name in icon_names_16:
            icon = self._load_icon_for_button(name, size=16)
            if icon:
                self._button_icons[name] = icon

        # Radio button icons for preset selection (16px) - reload on theme change
        self._load_radio_icons()

        # Checkbox icons for visual selection treeview (16px) - themed for light/dark mode
        self._load_checkbox_icons()

    def _load_radio_icons(self):
        """Load radio button SVG icons for selected and unselected states."""
        self._radio_on_icon = self._load_icon_for_button('radio-on', size=16)
        self._radio_off_icon = self._load_icon_for_button('radio-off', size=16)
        # Store in list to prevent garbage collection
        if self._radio_on_icon:
            self._icon_refs.append(self._radio_on_icon)
        if self._radio_off_icon:
            self._icon_refs.append(self._radio_off_icon)

    def _load_checkbox_icons(self):
        """Load themed checkbox SVG icons for checked and unchecked states."""
        is_dark = self._theme_manager.is_dark

        # Select icon names based on theme
        box_name = 'box-dark' if is_dark else 'box'
        checked_name = 'box-checked-dark' if is_dark else 'box-checked'

        # Load icons using base class method
        self._checkbox_unchecked_icon = self._load_icon_for_button(box_name, size=16)
        self._checkbox_checked_icon = self._load_icon_for_button(checked_name, size=16)
        # Store in list to prevent garbage collection
        if self._checkbox_unchecked_icon:
            self._icon_refs.append(self._checkbox_unchecked_icon)
        if self._checkbox_checked_icon:
            self._icon_refs.append(self._checkbox_checked_icon)

    def _create_preset_row(self, parent: tk.Widget, text: str, value: str,
                           preset_var: tk.StringVar, field_type: str) -> dict:
        """Create a single preset row with custom radio icon and text label"""
        colors = self._theme_manager.colors
        bg_color = colors.get('section_bg', colors['background'])

        # Row frame - more padding between options for better spacing
        row_frame = tk.Frame(parent, bg=bg_color)
        row_frame.pack(fill=tk.X, pady=3)

        # Radio button icon (clickable)
        is_selected = preset_var.get() == value
        icon = self._radio_on_icon if is_selected else self._radio_off_icon

        icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
        if icon:
            icon_label.configure(image=icon)
            icon_label._icon_ref = icon
        else:
            # Fallback to text if icons not available
            icon_label.configure(text="●" if is_selected else "○", font=('Segoe UI', 10))
        icon_label.pack(side=tk.LEFT, padx=(4, 8))

        # Text label (uses title_color when selected - blue in dark, teal in light)
        fg_color = colors['title_color'] if is_selected else colors['text_primary']
        text_label = tk.Label(row_frame, text=text, bg=bg_color, fg=fg_color,
                              font=('Segoe UI', 9), cursor='hand2', anchor='w')
        text_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Bind clicks to toggle selection
        def on_click(event=None):
            preset_var.set(value)
            self._update_preset_rows(field_type)

        # Hover underline effect with blue text color in dark mode
        def on_enter(event=None):
            # Get current theme colors for hover effect
            current_colors = self._theme_manager.colors
            # Always use title_color on hover (blue in dark, teal in light)
            text_label.configure(font=('Segoe UI', 9, 'underline'), fg=current_colors['title_color'])

        def on_leave(event=None):
            # Restore color based on selection state
            current_colors = self._theme_manager.colors
            is_selected = preset_var.get() == value
            fg_color = current_colors['title_color'] if is_selected else current_colors['text_primary']
            text_label.configure(font=('Segoe UI', 9), fg=fg_color)

        icon_label.bind('<Button-1>', on_click)
        text_label.bind('<Button-1>', on_click)
        row_frame.bind('<Enter>', on_enter)
        row_frame.bind('<Leave>', on_leave)

        row_data = {
            'frame': row_frame,
            'icon_label': icon_label,
            'text_label': text_label,
            'value': value,
            'field_type': field_type,
            'preset_var': preset_var
        }

        return row_data

    def _update_preset_rows(self, field_type: str):
        """Update all preset rows for a field type when selection changes"""
        colors = self._theme_manager.colors
        bg_color = colors.get('section_bg', colors['background'])

        if not hasattr(self, '_preset_rows') or field_type not in self._preset_rows:
            return

        for row_data in self._preset_rows[field_type]:
            preset_var = row_data['preset_var']
            is_selected = preset_var.get() == row_data['value']

            # Update icon
            icon = self._radio_on_icon if is_selected else self._radio_off_icon
            if icon:
                row_data['icon_label'].configure(image=icon)
                row_data['icon_label']._icon_ref = icon
            else:
                row_data['icon_label'].configure(text="●" if is_selected else "○")

            # Update text color (uses title_color when selected - blue in dark, teal in light)
            fg_color = colors['title_color'] if is_selected else colors['text_primary']
            row_data['text_label'].configure(fg=fg_color)

    def _create_preset_row_for_popup(self, parent: tk.Widget, text: str, value: str,
                                     preset_var: tk.StringVar, popup_key: str) -> dict:
        """Create a single preset row with custom radio icon for popup dialogs"""
        colors = self._theme_manager.colors
        # Use explicit colors matching popup frame backgrounds
        bg_color = colors['section_bg']

        # Row frame
        row_frame = tk.Frame(parent, bg=bg_color)
        row_frame.pack(fill=tk.X, pady=1)

        # Radio button icon (clickable)
        is_selected = preset_var.get() == value
        icon = self._radio_on_icon if is_selected else self._radio_off_icon

        icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
        if icon:
            icon_label.configure(image=icon)
            icon_label._icon_ref = icon
        else:
            # Fallback to text if icons not available
            icon_label.configure(text="●" if is_selected else "○", font=('Segoe UI', 10))
        icon_label.pack(side=tk.LEFT, padx=(4, 8))

        # Text label (uses title_color when selected - blue in dark, teal in light)
        fg_color = colors['title_color'] if is_selected else colors['text_primary']
        text_label = tk.Label(row_frame, text=text, bg=bg_color, fg=fg_color,
                              font=('Segoe UI', 9), cursor='hand2', anchor='w')
        text_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Bind clicks to toggle selection
        def on_click(event=None):
            preset_var.set(value)
            self._update_preset_rows_popup(popup_key)

        # Hover underline effect with blue text color in dark mode
        def on_enter(event=None):
            # Get current theme colors for hover effect
            current_colors = self._theme_manager.colors
            # Always use title_color on hover (blue in dark, teal in light)
            text_label.configure(font=('Segoe UI', 9, 'underline'), fg=current_colors['title_color'])

        def on_leave(event=None):
            # Restore color based on selection state
            current_colors = self._theme_manager.colors
            is_selected = preset_var.get() == value
            fg_color = current_colors['title_color'] if is_selected else current_colors['text_primary']
            text_label.configure(font=('Segoe UI', 9), fg=fg_color)

        icon_label.bind('<Button-1>', on_click)
        text_label.bind('<Button-1>', on_click)
        row_frame.bind('<Enter>', on_enter)
        row_frame.bind('<Leave>', on_leave)

        row_data = {
            'frame': row_frame,
            'icon_label': icon_label,
            'text_label': text_label,
            'value': value,
            'popup_key': popup_key,
            'preset_var': preset_var
        }

        return row_data

    def _update_preset_rows_popup(self, popup_key: str):
        """Update all preset rows for a popup when selection changes"""
        colors = self._theme_manager.colors
        # Use explicit colors matching popup frame backgrounds
        bg_color = colors['section_bg']

        if not hasattr(self, '_popup_preset_rows') or popup_key not in self._popup_preset_rows:
            return

        for row_data in self._popup_preset_rows[popup_key]:
            preset_var = row_data['preset_var']
            is_selected = preset_var.get() == row_data['value']

            # Update icon
            icon = self._radio_on_icon if is_selected else self._radio_off_icon
            if icon:
                row_data['icon_label'].configure(image=icon)
                row_data['icon_label']._icon_ref = icon
            else:
                row_data['icon_label'].configure(text="●" if is_selected else "○")

            # Update text color (uses title_color when selected - blue in dark, teal in light)
            fg_color = colors['title_color'] if is_selected else colors['text_primary']
            row_data['text_label'].configure(fg=fg_color)

    def on_theme_changed(self, theme: str):
        """Update theme-dependent widgets when theme changes"""
        super().on_theme_changed(theme)
        colors = self._theme_manager.colors
        # Inner content uses colors['background'] to match Section.TFrame style
        content_bg = colors['background']
        # section_bg for bordered sub-sections (Categorical/Measure frames, action buttons)
        section_bg = colors.get('section_bg', colors['background'])
        is_dark = self._theme_manager.is_dark

        # Use locked column widths to prevent layout shift during theme toggle
        # If not locked yet, try to lock them now
        if hasattr(self, '_middle_content') and self._middle_content:
            try:
                # If widths are already locked, re-apply the locked values with weight=0
                if hasattr(self, '_columns_width_locked') and self._columns_width_locked:
                    self._middle_content.columnconfigure(0, weight=0, minsize=self._locked_left_width)
                    self._middle_content.columnconfigure(1, weight=0, minsize=self._locked_right_width)
                    # ALSO re-apply explicit frame widths and propagate settings
                    if hasattr(self, '_left_column'):
                        self._left_column.configure(width=self._locked_left_width)
                        self._left_column.grid_propagate(False)
                    if hasattr(self, '_right_column'):
                        self._right_column.configure(width=self._locked_right_width)
                        self._right_column.grid_propagate(False)
                else:
                    # Try to lock now if not locked yet
                    self._lock_column_widths()

                # Update tk.Frame background colors
                self._middle_content.configure(bg=content_bg)
                if hasattr(self, '_left_column'):
                    self._left_column.configure(bg=content_bg)
                if hasattr(self, '_right_column'):
                    self._right_column.configure(bg=content_bg)
            except Exception:
                pass

        # Update scanning section frames (inner content uses colors['background'])
        if hasattr(self, '_scanning_content_frame'):
            try:
                self._scanning_content_frame.configure(bg=content_bg)
            except Exception:
                pass
        if hasattr(self, '_scanning_button_frame'):
            try:
                self._scanning_button_frame.configure(bg=content_bg)
            except Exception:
                pass

        # Reload themed checkbox icons (light/dark variants)
        self._load_checkbox_icons()

        # Update existing treeview items with new checkbox icons
        if hasattr(self, 'visual_tree') and self.visual_tree and self.visual_selection_vars:
            # Store icons on tree widget to prevent garbage collection
            self.visual_tree._checkbox_checked = self._checkbox_checked_icon
            self.visual_tree._checkbox_unchecked = self._checkbox_unchecked_icon
            try:
                for item in self.visual_tree.get_children():
                    tags = self.visual_tree.item(item, 'tags')
                    if tags:
                        visual_id = tags[0] if isinstance(tags, (list, tuple)) else tags
                        if visual_id in self.visual_selection_vars:
                            is_selected = self.visual_selection_vars[visual_id].get()
                            icon = self._checkbox_checked_icon if is_selected else self._checkbox_unchecked_icon
                            if icon:
                                self.visual_tree.item(item, image=icon)
                # Removed update_idletasks() - it triggers geometry recalculation causing width drift
            except Exception:
                pass

        # Update configuration section frames and labels (inner content uses colors['background'])
        if hasattr(self, '_config_inner_frame'):
            try:
                self._config_inner_frame.configure(bg=content_bg)
            except Exception:
                pass
        if hasattr(self, '_config_mode_frame'):
            try:
                self._config_mode_frame.configure(bg=content_bg)
            except Exception:
                pass
        if hasattr(self, '_mode_label'):
            try:
                self._mode_label.configure(bg=content_bg, fg=colors['text_primary'])
            except Exception:
                pass
        if hasattr(self, 'config_mode_label'):
            try:
                self.config_mode_label.configure(bg=content_bg, fg=colors['title_color'])
            except Exception:
                pass
        if hasattr(self, '_columns_frame'):
            try:
                self._columns_frame.configure(bg=content_bg)
            except Exception:
                pass

        # Note: Categorical and Measure title labels are updated in the bordered frames section below

        # Update selection section frames and labels (inner content uses colors['background'])
        if hasattr(self, '_selection_inner_frame'):
            try:
                self._selection_inner_frame.configure(bg=content_bg)
            except Exception:
                pass
        if hasattr(self, '_selection_instruction_frame'):
            try:
                self._selection_instruction_frame.configure(bg=content_bg)
            except Exception:
                pass
        if hasattr(self, '_selection_instruction_label'):
            try:
                self._selection_instruction_label.configure(bg=content_bg, fg=colors['title_color'])
            except Exception:
                pass
        if hasattr(self, '_selection_hint_label'):
            try:
                self._selection_hint_label.configure(bg=content_bg)  # fg stays #ff6600
            except Exception:
                pass
        if hasattr(self, '_selection_controls_frame'):
            try:
                self._selection_controls_frame.configure(bg=content_bg)
            except Exception:
                pass
        if hasattr(self, '_selection_tree_frame'):
            try:
                self._selection_tree_frame.configure(bg=content_bg)
            except Exception:
                pass

        # Update selection buttons (All/None/Tables/Matrices) for theme change
        if hasattr(self, '_selection_buttons') and self._selection_buttons:
            for btn in self._selection_buttons:
                try:
                    btn.update_colors(
                        bg=colors['card_surface'],
                        hover_bg=colors['card_surface_hover'],
                        pressed_bg=colors['card_surface_pressed'],
                        fg=colors['text_primary']
                    )
                except Exception:
                    pass

        # Update secondary buttons (Preview) colors - button_secondary style (like Reset All)
        if hasattr(self, 'preview_button') and self.preview_button:
            bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
            fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')
            try:
                self.preview_button.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    disabled_bg=bg_disabled,
                    disabled_fg=fg_disabled
                )
            except Exception:
                pass

        # Update reset button - button_secondary style (like Report Cleanup)
        if hasattr(self, 'reset_button') and self.reset_button:
            try:
                self.reset_button.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary']
                )
            except Exception:
                pass

        # Reload radio icons for theme change
        self._load_radio_icons()

        # Update preset rows (custom radio buttons) for theme change
        if hasattr(self, '_preset_rows') and self._preset_rows:
            for field_type in ['categorical', 'measure']:
                if field_type in self._preset_rows:
                    for row_data in self._preset_rows[field_type]:
                        try:
                            preset_var = row_data['preset_var']
                            is_selected = preset_var.get() == row_data['value']

                            # Update frame background
                            row_data['frame'].configure(bg=section_bg)

                            # Update icon label - separate bg and image updates
                            icon = self._radio_on_icon if is_selected else self._radio_off_icon
                            row_data['icon_label'].configure(bg=section_bg)
                            if icon:
                                # Apply image separately and store reference
                                row_data['icon_label'].configure(image=icon)
                                row_data['icon_label']._icon_ref = icon
                                # Also store on frame to prevent GC
                                row_data['frame']._icon_ref = icon
                            else:
                                # Fallback to text if icons not available
                                row_data['icon_label'].configure(text="●" if is_selected else "○")

                            # Update text color based on selection (uses title_color when selected)
                            fg_color = colors['title_color'] if is_selected else colors['text_primary']
                            row_data['text_label'].configure(bg=section_bg, fg=fg_color)
                        except Exception:
                            pass

        # Update treeview style for theme change - flat, modern design
        if hasattr(self, 'visual_tree') and self.visual_tree:
            # Set treeview colors based on theme - softer colors for modern look
            if is_dark:
                tree_bg = colors.get('surface', '#1e1e2e')
                tree_fg = colors.get('text_primary', '#e0e0e0')
                tree_field_bg = colors.get('surface', '#1e1e2e')
                heading_bg = colors.get('section_bg', '#1a1a2a')
                heading_fg = colors.get('text_primary', '#e0e0e0')
                selected_bg = '#1a3a5c'
                unselected_bg = colors.get('surface', '#1e1e2e')
                selected_fg = '#ffffff'
                unselected_fg = colors.get('text_primary', '#e0e0e0')
                tree_border = colors.get('border', '#3a3a4a')
                header_separator = '#0d0d1a'  # Faint column separator for dark mode
            else:
                tree_bg = colors.get('surface', '#ffffff')
                tree_fg = colors.get('text_primary', '#333333')
                tree_field_bg = colors.get('surface', '#ffffff')
                heading_bg = colors.get('section_bg', '#f5f5fa')
                heading_fg = colors.get('text_primary', '#333333')
                selected_bg = '#e6f3ff'
                unselected_bg = colors.get('surface', '#ffffff')
                selected_fg = '#1a1a2e'
                unselected_fg = colors.get('text_primary', '#333333')
                tree_border = '#d8d8e0'  # Softer border in light mode
                header_separator = '#ffffff'  # Faint column separator for light mode

            # Update tree container border and background
            if hasattr(self, '_tree_container') and self._tree_container:
                try:
                    self._tree_container.configure(bg=tree_bg,
                                                   highlightbackground=tree_border,
                                                   highlightcolor=tree_border)
                except Exception:
                    pass

            try:
                # Update treeview style - flat design with no internal borders
                style = ttk.Style()
                tree_style = "ColumnWidth.Treeview"
                style.configure(tree_style,
                                background=tree_bg,
                                foreground=tree_fg,
                                fieldbackground=tree_field_bg,
                                font=('Segoe UI', 9),
                                relief='flat',
                                borderwidth=0,
                                bordercolor=tree_bg,
                                lightcolor=tree_bg,
                                darkcolor=tree_bg)
                # Ensure layout removes treeview frame border
                style.layout(tree_style, [
                    ('Treeview.treearea', {'sticky': 'nswe'})
                ])
                style.configure(f"{tree_style}.Heading",
                                background=heading_bg,
                                foreground=heading_fg,
                                relief='groove',
                                borderwidth=1,
                                bordercolor=header_separator,
                                lightcolor=header_separator,
                                darkcolor=header_separator)
                style.map(f"{tree_style}.Heading",
                          relief=[('active', 'groove'), ('pressed', 'groove')],
                          background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])
                style.map(tree_style,
                          background=[('selected', selected_bg)],
                          foreground=[('selected', selected_fg)])

                # Update tag colors
                self.visual_tree.tag_configure('selected', background=selected_bg, foreground=selected_fg)
                self.visual_tree.tag_configure('unselected', background=unselected_bg, foreground=unselected_fg)
                # Keep hover tag consistent with font
                self.visual_tree.tag_configure('hover', font=('Segoe UI', 9, 'underline'))
            except Exception:
                pass

        # SplitLogSection handles its own theme updates internally

        # Update custom entry widgets for theme change
        if hasattr(self, '_custom_entries') and self._custom_entries:
            entry_bg = colors.get('card_surface', '#ffffff')
            entry_fg = colors.get('text_primary', '#000000')
            # Use lighter border in light mode (too intense otherwise)
            entry_border_color = colors.get('border', '#e0e0e0') if is_dark else '#c8c8d0'

            for field_type, widgets in self._custom_entries.items():
                try:
                    # Update frame background
                    widgets['frame'].configure(bg=section_bg)
                    # Update label colors
                    widgets['label'].configure(bg=section_bg, fg=colors['text_primary'])
                    # Update entry colors - use same border color for focus (no color change)
                    widgets['entry'].configure(
                        bg=entry_bg, fg=entry_fg,
                        insertbackground=entry_fg,
                        highlightbackground=entry_border_color,
                        highlightcolor=entry_border_color
                    )
                except Exception:
                    pass

        # Update Categorical and Measure bordered frames
        # Grid layout with uniform='config_cols' ensures equal widths are maintained
        border_color_frames = colors.get('border', '#e0e0e0')

        if hasattr(self, '_cat_outer_frame') and self._cat_outer_frame:
            try:
                self._cat_outer_frame.configure(bg=section_bg,
                                                highlightbackground=border_color_frames,
                                                highlightcolor=border_color_frames)
                self._cat_inner_frame.configure(bg=section_bg)
                self._cat_title_label.configure(bg=section_bg, fg=colors['title_color'])
            except Exception:
                pass

        if hasattr(self, '_measure_outer_frame') and self._measure_outer_frame:
            try:
                self._measure_outer_frame.configure(bg=section_bg,
                                                    highlightbackground=border_color_frames,
                                                    highlightcolor=border_color_frames)
                self._measure_inner_frame.configure(bg=section_bg)
                self._measure_title_label.configure(bg=section_bg, fg=colors['title_color'])
            except Exception:
                pass

        # Update action buttons section frames
        if hasattr(self, '_action_frame'):
            try:
                self._action_frame.configure(bg=section_bg)
            except Exception:
                pass
        if hasattr(self, '_button_container'):
            try:
                self._button_container.configure(bg=section_bg)
            except Exception:
                pass

        # Update Analysis Summary text colors for theme change
        self._apply_summary_text_colors()

    def _setup_file_input_section(self):
        """Setup file input section using FileInputSection template"""
        # Get scan icon
        scan_icon = self._button_icons.get('magnifying-glass')

        # Create FileInputSection with all components
        self.file_section = FileInputSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="PBIP File Source",
            section_icon="Power-BI",
            file_label="Project File (PBIP):",
            file_types=[("Power BI Project Files", "*.pbip")],
            action_button_text="SCAN VISUALS",
            action_button_command=self._scan_visuals,
            action_button_icon=scan_icon,
            help_command=self.show_help_dialog
        )
        self.file_section.pack(fill=tk.X, pady=(0, 15))

        # Store references for backward compatibility
        self.pbip_path_var = self.file_section.path_var
        self.scan_button = self.file_section.action_button
        self._file_source_section = self.file_section.section_frame

    def _setup_scanning_section(self, parent):
        """Setup visual scanning section"""
        colors = self._theme_manager.colors
        # Inner content uses colors['background'] to match Section.TFrame style
        content_bg = colors['background']

        # Create section with icon labelwidget
        header_widget = self.create_section_header(self.frame, "Visual Scanning", "magnifying-glass")[0]
        scan_frame = ttk.LabelFrame(parent, labelwidget=header_widget,
                                    style='Section.TLabelframe', padding="12")
        scan_frame.pack(fill=tk.X, pady=(0, 15))
        self._scanning_section = scan_frame  # Store for theme updates

        # Inner content frame - use Section.TFrame style for proper background + padding
        # Section.TFrame is configured with colors['background'] in theme_manager (auto-themes)
        content_frame = ttk.Frame(scan_frame, style='Section.TFrame', padding="15")
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Instructions
        instruction_label = tk.Label(content_frame, text="Scan Report for Tables and Matrices",
                                     bg=content_bg, fg=colors['title_color'],
                                     font=('Segoe UI', 10, 'bold'))
        instruction_label.pack(anchor=tk.W, pady=(0, 5))

        # Scan button and summary in horizontal layout
        button_frame = tk.Frame(content_frame, bg=content_bg)
        button_frame.pack(fill=tk.X, pady=(5, 0))
        self._scanning_button_frame = button_frame  # Store for theme updates

        # Scan button with icon
        scan_icon = self._button_icons.get('magnifying-glass')
        self.scan_button = RoundedButton(
            button_frame, text="SCAN VISUALS",
            command=self._scan_visuals,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=scan_icon
        )
        self.scan_button.pack(side=tk.LEFT)
        self._primary_buttons.append(self.scan_button)

        # Store widgets for theme updates
        self._scanning_widgets = {
            'instruction_label': instruction_label
        }
    
    def _setup_configuration_section(self, parent):
        """Setup width configuration section"""
        colors = self._theme_manager.colors
        # Inner content uses colors['background'] to match Section.TFrame style (creates contrast with section_bg border)
        content_bg = colors['background']

        # Create section with icon labelwidget - use "table column widths" icon
        header_widget = self.create_section_header(self.frame, "Column Width Configuration", "table column widths")[0]
        config_frame = ttk.LabelFrame(parent, labelwidget=header_widget,
                                      style='Section.TLabelframe', padding="12")
        config_frame.pack(fill=tk.BOTH, expand=True)  # expand=True so bottom aligns with Visual Selection
        self._config_section = config_frame  # Store for theme updates

        # Inner content frame - use Section.TFrame style for proper background + padding
        # Section.TFrame is configured with colors['background'] in theme_manager (auto-themes)
        inner_frame = ttk.Frame(config_frame, style='Section.TFrame', padding="15")
        inner_frame.pack(fill=tk.BOTH, expand=True)  # expand=True fills the config_frame so no gap appears

        # Configuration mode selector
        mode_frame = tk.Frame(inner_frame, bg=content_bg)
        mode_frame.pack(fill=tk.X, pady=(0, 8))
        self._config_mode_frame = mode_frame  # Store for theme updates

        mode_label = tk.Label(mode_frame, text="Configuration Mode:",
                              bg=content_bg, fg=colors['text_primary'],
                              font=('Segoe UI', 9, 'bold'))
        mode_label.pack(side=tk.LEFT)
        self._mode_label = mode_label  # Store for theme updates

        # Configuration mode label - will update dynamically
        self.config_mode_label = tk.Label(mode_frame, text="Global Settings (All Visuals)",
                                          bg=content_bg, fg=colors['title_color'],
                                          font=('Segoe UI', 9))
        self.config_mode_label.pack(side=tk.LEFT, padx=(10, 0))

        # Categorical and measure settings in horizontal layout using GRID with uniform
        # This ensures both columns always maintain equal widths regardless of content
        columns_frame = tk.Frame(inner_frame, bg=content_bg)
        columns_frame.pack(fill=tk.BOTH, expand=True)
        self._columns_frame = columns_frame  # Store for theme update layout stability

        # Configure grid columns with uniform group to enforce equal widths
        columns_frame.columnconfigure(0, weight=1, uniform='config_cols')
        columns_frame.columnconfigure(1, weight=1, uniform='config_cols')
        columns_frame.rowconfigure(0, weight=1)

        # Get theme colors for bordered frames
        frame_bg = colors.get('section_bg', colors['background'])
        border_color = colors.get('border', '#e0e0e0')

        # Categorical columns (left) - tk.Frame with border for visual grouping
        # Use same color for highlightcolor to prevent border change on focus
        cat_outer = tk.Frame(columns_frame, bg=frame_bg,
                            highlightbackground=border_color, highlightcolor=border_color,
                            highlightthickness=1)
        cat_outer.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 8))
        self._cat_outer_frame = cat_outer  # Store for theme updates

        # Inner padding frame
        cat_frame = tk.Frame(cat_outer, bg=frame_bg, padx=10, pady=8)
        cat_frame.pack(fill=tk.BOTH, expand=True)
        self._cat_inner_frame = cat_frame  # Store for theme updates

        # Title label - matches "Choose visuals to update" style
        cat_label = tk.Label(cat_frame, text="Categorical", bg=frame_bg,
                            fg=colors['title_color'], font=('Segoe UI', 10, 'bold'))
        cat_label.pack(anchor=tk.W, pady=(0, 8))
        self._cat_title_label = cat_label  # Store for theme updates

        self._setup_width_controls_with_custom(cat_frame, self.categorical_preset_var, 'categorical')

        # Measure columns (right) - tk.Frame with border for visual grouping
        # Use same color for highlightcolor to prevent border change on focus
        measure_outer = tk.Frame(columns_frame, bg=frame_bg,
                                highlightbackground=border_color, highlightcolor=border_color,
                                highlightthickness=1)
        measure_outer.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(8, 0))
        self._measure_outer_frame = measure_outer  # Store for theme updates

        # Inner padding frame
        measure_frame = tk.Frame(measure_outer, bg=frame_bg, padx=10, pady=8)
        measure_frame.pack(fill=tk.BOTH, expand=True)
        self._measure_inner_frame = measure_frame  # Store for theme updates

        # Title label - matches "Choose visuals to update" style
        measure_label = tk.Label(measure_frame, text="Measure", bg=frame_bg,
                                fg=colors['title_color'], font=('Segoe UI', 10, 'bold'))
        measure_label.pack(anchor=tk.W, pady=(0, 8))
        self._measure_title_label = measure_label  # Store for theme updates

        self._setup_width_controls_with_custom(measure_frame, self.measure_preset_var, 'measure')

        # Store widgets for theme updates
        self._config_widgets = {
            'mode_label': mode_label,
            'mode_value_label': self.config_mode_label
        }
    
    def _setup_width_controls_simple(self, parent: ttk.Widget, preset_var: tk.StringVar):
        """Setup simplified width control widgets"""
        
        # Enhanced preset options
        presets = [
            ("Narrow", WidthPreset.NARROW.value),
            ("Medium", WidthPreset.MEDIUM.value),
            ("Wide", WidthPreset.WIDE.value),
            ("Fit to Header", WidthPreset.AUTO_FIT.value),
            ("Fit to Totals", WidthPreset.FIT_TO_TOTALS.value),
            ("Custom", WidthPreset.CUSTOM.value)
        ]
        
        for text, value in presets:
            radio = ttk.Radiobutton(parent, text=text, variable=preset_var, value=value)
            radio.pack(anchor=tk.W, pady=1)
    
    def _setup_width_controls_with_custom(self, parent: ttk.Widget, preset_var: tk.StringVar, field_type: str):
        """Setup width control widgets with custom input field"""
        colors = self._theme_manager.colors

        # Enhanced preset options
        presets = [
            ("Narrow", WidthPreset.NARROW.value),
            ("Medium", WidthPreset.MEDIUM.value),
            ("Wide", WidthPreset.WIDE.value),
            ("Fit to Header", WidthPreset.AUTO_FIT.value),
            ("Fit to Totals", WidthPreset.FIT_TO_TOTALS.value),
            ("Custom", WidthPreset.CUSTOM.value)
        ]

        # Clear existing preset rows for this field type
        if hasattr(self, '_preset_rows'):
            self._preset_rows[field_type] = []

        # Create custom radio preset rows (like Layout Optimizer)
        for text, value in presets:
            row_data = self._create_preset_row(parent, text, value, preset_var, field_type)
            if hasattr(self, '_preset_rows'):
                self._preset_rows[field_type].append(row_data)

        # Custom width input - single line layout with proper theme colors
        bg_color = colors.get('section_bg', colors['background'])
        custom_frame = tk.Frame(parent, bg=bg_color)
        custom_frame.pack(anchor=tk.W, pady=(8, 0))

        # Custom label with matching background
        custom_label = tk.Label(custom_frame, text="Custom (px):", bg=bg_color,
                               fg=colors['text_primary'], font=('Segoe UI', 9))
        custom_label.pack(side=tk.LEFT)

        # Create custom width variable
        if field_type == 'categorical':
            custom_var = tk.StringVar(value="105")
            self.categorical_custom_var = custom_var
        else:
            custom_var = tk.StringVar(value="95")
            self.measure_custom_var = custom_var

        # Entry box with theme-aware colors
        # Use card_surface for input background (white in light, dark in dark mode)
        entry_bg = colors.get('card_surface', '#ffffff')
        entry_fg = colors.get('text_primary', '#000000')
        # Use lighter border in light mode (too intense otherwise)
        is_dark = self._theme_manager.is_dark
        border_color = colors.get('border', '#e0e0e0') if is_dark else '#c8c8d0'

        custom_entry = tk.Entry(custom_frame, textvariable=custom_var, width=6, font=('Segoe UI', 9),
                               relief='flat', borderwidth=0, bg=entry_bg, fg=entry_fg,
                               insertbackground=entry_fg, highlightbackground=border_color,
                               highlightcolor=border_color, highlightthickness=1)
        custom_entry.pack(side=tk.LEFT, padx=(8, 0))

        # Store references for theme updates
        if not hasattr(self, '_custom_entries'):
            self._custom_entries = {}
        self._custom_entries[field_type] = {
            'frame': custom_frame,
            'label': custom_label,
            'entry': custom_entry
        }

        # Delayed validation to prevent false positives while typing
        validation_timer = None

        def delayed_validate():
            """Actual validation function called after delay"""
            colors = self._theme_manager.colors
            entry_bg = colors.get('card_surface', '#ffffff')
            error_bg = '#ffebee' if not self._theme_manager.is_dark else '#4a1a1a'
            try:
                value = int(custom_var.get())
                if value < 50:  # Below minimum width
                    # Change background to error color (theme-aware)
                    custom_entry.config(bg=error_bg)
                    # Show warning in log
                    field_name = "categorical" if field_type == 'categorical' else "measure"
                    self.log_message(f"⚠️ Warning: {field_name} custom width {value}px is below minimum (50px). Will be increased to 50px.")
                else:
                    # Reset to normal background (theme-aware)
                    custom_entry.config(bg=entry_bg)
            except ValueError:
                # Reset to normal background for non-numeric values (theme-aware)
                custom_entry.config(bg=entry_bg)

        def validate_custom_width(*args):
            """Validation with delay - resets timer on each keystroke"""
            nonlocal validation_timer
            # Cancel any pending validation
            if validation_timer:
                self.frame.after_cancel(validation_timer)
            # Schedule new validation after 800ms delay
            validation_timer = self.frame.after(800, delayed_validate)

        custom_var.trace('w', validate_custom_width)
    
    def _setup_width_controls_with_custom_for_popup(self, parent: tk.Widget, preset_var: tk.StringVar, field_type: str, visual_id: str):
        """Setup width control widgets with custom input field for popup dialogs"""
        colors = self._theme_manager.colors

        # Enhanced preset options
        presets = [
            ("Narrow", WidthPreset.NARROW.value),
            ("Medium", WidthPreset.MEDIUM.value),
            ("Wide", WidthPreset.WIDE.value),
            ("Fit to Header", WidthPreset.AUTO_FIT.value),
            ("Fit to Totals", WidthPreset.FIT_TO_TOTALS.value),
            ("Custom", WidthPreset.CUSTOM.value)
        ]

        # Initialize popup preset rows tracking if needed
        popup_key = f"{visual_id}_{field_type}"
        if not hasattr(self, '_popup_preset_rows'):
            self._popup_preset_rows = {}
        self._popup_preset_rows[popup_key] = []

        # Create SVG radio button rows (like main tool)
        for text, value in presets:
            row_data = self._create_preset_row_for_popup(parent, text, value, preset_var, popup_key)
            self._popup_preset_rows[popup_key].append(row_data)
        
        # Custom width input - single line layout with theme colors
        is_dark = self._theme_manager.is_dark
        bg_color = colors['section_bg']
        custom_frame = tk.Frame(parent, bg=bg_color)
        custom_frame.pack(anchor=tk.W, pady=(8, 0))

        # Custom label with matching background
        custom_label = tk.Label(custom_frame, text="Custom (px):", bg=bg_color,
                               fg=colors['text_primary'], font=('Segoe UI', 9))
        custom_label.pack(side=tk.LEFT)

        # Get or create per-visual custom variables
        if visual_id not in self.visual_config_vars:
            self._create_per_visual_config_vars(visual_id)

        # Get/create custom variable for this visual
        if f'{field_type}_custom' not in self.visual_config_vars[visual_id]:
            default_value = "105" if field_type == 'categorical' else "95"
            self.visual_config_vars[visual_id][f'{field_type}_custom'] = tk.StringVar(value=default_value)

        custom_var = self.visual_config_vars[visual_id][f'{field_type}_custom']

        # Entry box with theme-aware colors matching frame background
        entry_bg = colors['section_bg']
        entry_fg = colors.get('text_primary', '#000000')
        border_color = colors.get('border', '#3d3d5c' if is_dark else '#d0d0e0')
        error_bg = '#4a1a1a' if is_dark else '#ffebee'

        custom_entry = tk.Entry(custom_frame, textvariable=custom_var, width=6, font=('Segoe UI', 9),
                               relief='flat', borderwidth=0, bg=entry_bg, fg=entry_fg,
                               insertbackground=entry_fg, highlightbackground=border_color,
                               highlightcolor=border_color, highlightthickness=1)
        custom_entry.pack(side=tk.LEFT, padx=(8, 0))

        # Delayed validation to prevent false positives while typing
        validation_timer = None

        def delayed_validate_popup():
            """Actual validation function called after delay for popup"""
            try:
                value = int(custom_var.get())
                if value < 50:  # Below minimum width
                    # Change background to error color (theme-aware)
                    custom_entry.config(bg=error_bg)
                    # Show warning in log
                    field_name = "categorical" if field_type == 'categorical' else "measure"
                    self.log_message(f"⚠️ Warning: Per-visual {field_name} custom width {value}px is below minimum (50px). Will be increased to 50px.")
                else:
                    # Reset to normal background (theme-aware)
                    custom_entry.config(bg=entry_bg)
            except ValueError:
                # Reset to normal background for non-numeric values (theme-aware)
                custom_entry.config(bg=entry_bg)
        
        def validate_popup_custom_width(*args):
            """Validation with delay - resets timer on each keystroke"""
            nonlocal validation_timer
            # Cancel any pending validation
            if validation_timer:
                self.frame.after_cancel(validation_timer)
            # Schedule new validation after 800ms delay
            validation_timer = self.frame.after(800, delayed_validate_popup)
        
        custom_var.trace('w', validate_popup_custom_width)
    
    def _setup_visual_selection_section(self, parent):
        """Setup visual selection section"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        # Inner content uses colors['background'] to match Section.TFrame style
        content_bg = colors['background']

        # Create section with icon labelwidget - use "table" icon
        header_widget = self.create_section_header(self.frame, "Visual Selection", "table")[0]
        selection_frame = ttk.LabelFrame(parent, labelwidget=header_widget,
                                         style='Section.TLabelframe', padding="12")
        selection_frame.pack(fill=tk.BOTH, expand=False)  # No padx - LabelFrame borders provide separation; expand=False prevents height growth
        self._selection_section = selection_frame  # Store for theme updates

        # Inner content frame - use Section.TFrame style for proper background + padding
        # Section.TFrame is configured with colors['background'] in theme_manager (auto-themes)
        inner_frame = ttk.Frame(selection_frame, style='Section.TFrame', padding="15")
        inner_frame.pack(fill=tk.BOTH, expand=False)  # expand=False prevents height growth

        # Instructions with double-click hint - single line
        instruction_frame = tk.Frame(inner_frame, bg=content_bg)
        instruction_frame.pack(fill=tk.X, pady=(0, 8))
        self._selection_instruction_frame = instruction_frame  # Store for theme updates

        instruction_label = tk.Label(instruction_frame, text="Choose visuals to update",
                                     bg=content_bg, fg=colors['title_color'],
                                     font=('Segoe UI', 10, 'bold'))
        instruction_label.pack(side=tk.LEFT)
        self._selection_instruction_label = instruction_label  # Store for theme updates

        hint_label = tk.Label(instruction_frame, text="Double-click any visual for per-visual configuration",
                              bg=content_bg, fg='#ff6600',
                              font=('Segoe UI', 9))
        hint_label.pack(side=tk.RIGHT)
        self._selection_hint_label = hint_label  # Store for theme updates

        # Selection controls - compact RoundedButtons
        controls_frame = tk.Frame(inner_frame, bg=content_bg)
        controls_frame.pack(fill=tk.X, pady=(8, 8))
        self._selection_controls_frame = controls_frame  # Store for theme updates

        # All button
        self._all_button = RoundedButton(
            controls_frame, text="All", command=self._select_all_visuals,
            bg=colors['card_surface'], hover_bg=colors['card_surface_hover'],
            pressed_bg=colors['card_surface_pressed'], fg=colors['text_primary'],
            width=58, height=26, radius=5, font=('Segoe UI', 9)
        )
        self._all_button.pack(side=tk.LEFT, padx=(0, 5))

        # None button
        self._none_button = RoundedButton(
            controls_frame, text="None", command=self._clear_all_visuals,
            bg=colors['card_surface'], hover_bg=colors['card_surface_hover'],
            pressed_bg=colors['card_surface_pressed'], fg=colors['text_primary'],
            width=58, height=26, radius=5, font=('Segoe UI', 9)
        )
        self._none_button.pack(side=tk.LEFT, padx=(0, 5))

        # Tables button
        self._tables_button = RoundedButton(
            controls_frame, text="Tables", command=self._select_tables_only,
            bg=colors['card_surface'], hover_bg=colors['card_surface_hover'],
            pressed_bg=colors['card_surface_pressed'], fg=colors['text_primary'],
            width=68, height=26, radius=5, font=('Segoe UI', 9)
        )
        self._tables_button.pack(side=tk.LEFT, padx=(0, 5))

        # Matrices button
        self._matrices_button = RoundedButton(
            controls_frame, text="Matrices", command=self._select_matrices_only,
            bg=colors['card_surface'], hover_bg=colors['card_surface_hover'],
            pressed_bg=colors['card_surface_pressed'], fg=colors['text_primary'],
            width=68, height=26, radius=5, font=('Segoe UI', 9)
        )
        self._matrices_button.pack(side=tk.LEFT)

        # Store selection buttons for theme updates
        self._selection_buttons = [
            self._all_button, self._none_button,
            self._tables_button, self._matrices_button
        ]

        # Treeview for visual selection with flat, modern design (matching Progress Log)
        # expand=False prevents vertical growth when data is loaded
        tree_frame = tk.Frame(inner_frame, bg=content_bg)
        tree_frame.pack(fill=tk.BOTH, expand=False)
        self._selection_tree_frame = tree_frame  # Store for theme updates

        # Configure treeview style for dark/light mode - flat design
        style = ttk.Style()
        tree_style = "ColumnWidth.Treeview"

        # Set treeview colors based on theme - softer colors for modern look
        if is_dark:
            tree_bg = colors.get('surface', '#1e1e2e')
            tree_fg = colors.get('text_primary', '#e0e0e0')
            tree_field_bg = colors.get('surface', '#1e1e2e')
            heading_bg = colors.get('section_bg', '#1a1a2a')
            heading_fg = colors.get('text_primary', '#e0e0e0')
            selected_bg = '#1a3a5c'
            unselected_bg = colors.get('surface', '#1e1e2e')
            selected_fg = '#ffffff'
            unselected_fg = colors.get('text_primary', '#e0e0e0')
            tree_border = colors.get('border', '#3a3a4a')
            header_separator = '#0d0d1a'  # Faint column separator for dark mode
        else:
            tree_bg = colors.get('surface', '#ffffff')
            tree_fg = colors.get('text_primary', '#333333')
            tree_field_bg = colors.get('surface', '#ffffff')
            heading_bg = colors.get('section_bg', '#f5f5fa')
            heading_fg = colors.get('text_primary', '#333333')
            selected_bg = '#e6f3ff'
            unselected_bg = colors.get('surface', '#ffffff')
            selected_fg = '#1a1a2e'
            unselected_fg = colors.get('text_primary', '#333333')
            tree_border = '#d8d8e0'  # Softer border in light mode
            header_separator = '#ffffff'  # Faint column separator for light mode

        # Configure flat treeview style - no 3D effects, no internal borders
        style.configure(tree_style,
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=tree_field_bg,
                        font=('Segoe UI', 9),
                        rowheight=25,
                        relief='flat',
                        borderwidth=0,
                        bordercolor=tree_bg,
                        lightcolor=tree_bg,
                        darkcolor=tree_bg)

        # Remove treeview frame border by styling the Item element
        style.layout(tree_style, [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

        # Heading style with subtle column separators using groove relief
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

        # Keep groove relief on active/pressed states for consistent separators
        # Include ('', heading_bg) for Python 3.13+ compatibility
        style.map(f"{tree_style}.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])

        style.map(tree_style,
                  background=[('selected', selected_bg)],
                  foreground=[('selected', selected_fg)])

        # Create single container with 1px border (matching Progress Log style)
        # Fixed height: rowheight(25) * rows(7) + heading(~25) = ~200px
        # Using explicit height to prevent ANY vertical growth when data loads
        tree_container = tk.Frame(tree_frame, bg=tree_bg,
                                  highlightbackground=tree_border, highlightcolor=tree_border,
                                  highlightthickness=1, height=200)
        tree_container.pack(fill=tk.X, expand=False)  # fill=X only, not BOTH - prevents height changes
        tree_container.pack_propagate(False)  # CRITICAL: prevents children from affecting container size
        self._tree_container = tree_container  # Store for theme updates
        self._tree_border_color = tree_border  # Store for theme updates

        # Add ThemedScrollbar FIRST (pack on right) so space is reserved before treeview expands
        self._tree_scrollbar = ThemedScrollbar(tree_container,
                                               theme_manager=self._theme_manager,
                                               width=12,
                                               auto_hide=False)  # Always show scrollbar
        self._tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview with custom styling (reduced height by ~25px)
        # Note: borderwidth/relief don't work on ttk.Treeview - handled via style
        self.visual_tree = ttk.Treeview(tree_container, height=7, selectmode='none',
                                        style=tree_style)
        self.visual_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=0, pady=0)  # Container is fixed size, so expand=True is safe here

        # Connect scrollbar to treeview (set _command directly since ThemedScrollbar is a Canvas)
        self._tree_scrollbar._command = self.visual_tree.yview
        self.visual_tree.configure(yscrollcommand=self._tree_scrollbar.set)

        # Configure columns - Checkbox (#0) is separate from Visual Name
        self.visual_tree['columns'] = ('visual_name', 'type', 'page', 'fields', 'config')

        # #0 column is just for checkbox icon - fixed width to center under header
        self.visual_tree.heading('#0', text='', anchor=tk.CENTER)
        self.visual_tree.column('#0', width=44, minwidth=44, stretch=False, anchor=tk.CENTER)

        # Visual Name in its own column (like Advanced Copy pattern)
        self.visual_tree.heading('visual_name', text='Visual Name', anchor=tk.CENTER)
        self.visual_tree.column('visual_name', width=145, minwidth=100, stretch=True, anchor=tk.CENTER)  # stretch=True to fill extra width

        self.visual_tree.heading('type', text='Type', anchor=tk.CENTER)
        self.visual_tree.heading('page', text='Page', anchor=tk.CENTER)
        self.visual_tree.heading('fields', text='Fields', anchor=tk.CENTER)
        self.visual_tree.heading('config', text='Config', anchor=tk.CENTER)

        # Configure column widths - all fixed except visual_name which absorbs width changes
        self.visual_tree.column('type', width=100, minwidth=80, stretch=False, anchor=tk.CENTER)
        self.visual_tree.column('page', width=80, minwidth=60, stretch=False, anchor=tk.CENTER)
        self.visual_tree.column('fields', width=70, minwidth=55, stretch=False, anchor=tk.CENTER)
        self.visual_tree.column('config', width=60, minwidth=55, stretch=False, anchor=tk.CENTER)

        # Configure tree tag styling for checkboxes
        self.visual_tree.tag_configure('selected', background=selected_bg, foreground=selected_fg)
        self.visual_tree.tag_configure('unselected', background=unselected_bg, foreground=unselected_fg)
        # Hover tag - same font/size with underline added
        self.visual_tree.tag_configure('hover', font=('Segoe UI', 9, 'underline'))

        # Track currently hovered item (cursor changes to hand2 on hover for click feedback)
        self._hovered_tree_item = None

        # Store widgets for theme updates
        self._selection_widgets = {
            'instruction_label': instruction_label,
            'hint_label': hint_label
        }

        # Bind events
        self.visual_tree.bind('<Button-1>', self._on_tree_click)
        self.visual_tree.bind('<Motion>', self._on_tree_hover)
        self.visual_tree.bind('<Leave>', self._on_tree_leave)  # Clear hover underline when leaving
        self.visual_tree.bind('<Double-Button-1>', self._on_tree_double_click)  # For per-visual config
    
    def _setup_action_buttons(self):
        """Setup action buttons at the bottom of the tab"""
        colors = self._theme_manager.colors
        section_bg = colors.get('section_bg', colors['background'])

        # Get disabled colors from theme for proper styling
        is_dark = self._theme_manager.is_dark
        bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Create action frame at bottom
        action_frame = tk.Frame(self.frame, bg=section_bg)
        action_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(15, 0))
        self._action_frame = action_frame  # Store for theme updates

        # Center container for buttons
        button_container = tk.Frame(action_frame, bg=section_bg)
        button_container.pack(anchor=tk.CENTER)
        self._button_container = button_container  # Store for theme updates

        # Bottom buttons sit on main outer background, need specific canvas_bg
        outer_canvas_bg = colors['section_bg']

        # Preview button with eye icon - uses button_secondary colors like Reset All
        eye_icon = self._button_icons.get('eye')
        self.preview_button = RoundedButton(
            button_container, text="PREVIEW",
            command=self._preview_changes,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=eye_icon,
            disabled_bg=bg_disabled, disabled_fg=fg_disabled,
            canvas_bg=outer_canvas_bg
        )
        self.preview_button.pack(side=tk.LEFT, padx=(0, 10))
        self.preview_button.set_enabled(False)  # Initially disabled
        self._secondary_buttons.append(self.preview_button)

        # Apply Changes button with execute icon (primary action)
        execute_icon = self._button_icons.get('execute')
        self.apply_button = RoundedButton(
            button_container, text="APPLY CHANGES",
            command=self._apply_changes,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=execute_icon,
            disabled_bg=bg_disabled, disabled_fg=fg_disabled,
            canvas_bg=outer_canvas_bg
        )
        self.apply_button.pack(side=tk.LEFT, padx=(0, 10))
        self.apply_button.set_enabled(False)  # Initially disabled
        self._primary_buttons.append(self.apply_button)

        # Reset button with reset icon - matches Report Cleanup style with button_secondary colors
        reset_icon = self._button_icons.get('reset')
        self.reset_button = RoundedButton(
            button_container, text="RESET ALL" if reset_icon else "RESET ALL",
            command=self.reset_tab,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=38, radius=6, font=('Segoe UI', 10),
            icon=reset_icon,
            canvas_bg=outer_canvas_bg
        )
        self.reset_button.pack(side=tk.LEFT)
        self._secondary_buttons.append(self.reset_button)

    def _handle_theme_change(self, theme: str):
        """Handle theme change - update outer button canvas_bg"""
        # Call base class first
        try:
            super()._handle_theme_change(theme)
        except Exception:
            pass

        # Update bottom action buttons canvas_bg for proper corner rounding on outer background
        colors = self._theme_manager.colors
        outer_canvas_bg = colors['section_bg']

        try:
            if hasattr(self, 'preview_button') and self.preview_button:
                self.preview_button.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    canvas_bg=outer_canvas_bg
                )
            if hasattr(self, 'apply_button') and self.apply_button:
                self.apply_button.update_colors(
                    bg=colors['button_primary'],
                    hover_bg=colors['button_primary_hover'],
                    pressed_bg=colors['button_primary_pressed'],
                    fg='#ffffff',
                    canvas_bg=outer_canvas_bg
                )
            if hasattr(self, 'reset_button') and self.reset_button:
                self.reset_button.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    canvas_bg=outer_canvas_bg
                )
        except Exception:
            pass

    def _show_welcome_message(self):
        """Show welcome message for Table Column Widths"""
        self.log_message("📊 Welcome to Table Column Widths!")
        self.log_message("=" * 60)
        self.log_message("📏 Standardize column widths across your Power BI tables and matrices")
        self.log_message("🎯 Features: Uniform column sizing with per-visual configuration")
        self.log_message("")
        self.log_message("👆 Start by selecting a .pbip file and clicking 'SCAN VISUALS'")
        self.log_message("🔄 Double-click any visual in the list for per-visual configuration")
    
    def _scan_visuals(self):
        """Scan the PBIP file for visuals"""
        try:
            pbip_path = self.clean_file_path(self.pbip_path_var.get())
            self.validate_pbip_file(pbip_path, "PBIP file")
            
            def scan_operation():
                self.update_progress(10, "Initializing table column widths engine...")
                self.engine = TableColumnWidthsEngine(pbip_path)
                
                self.update_progress(30, "Scanning report structure...")
                self.visuals_info = self.engine.scan_visuals()
                
                self.update_progress(70, "Analyzing visual configurations...")
                # Additional processing time for realism
                import time
                time.sleep(0.5)
                
                self.update_progress(100, "Visual scanning complete!")
                return len(self.visuals_info)
            
            def on_success(result):
                self._update_visual_tree()
                self._update_scan_summary()
                self._enable_configuration_ui()
                
                if result > 0:
                    self.log_message(f"✅ Successfully scanned {result} table/matrix visuals")
                    summary = self.engine.get_visual_summary()
                    self.log_message(f"📊 Found: {summary['table_count']} tables, {summary['matrix_count']} matrices")
                    self.log_message(f"📈 Total fields: {summary['total_fields']} ({summary['categorical_fields']} categorical, {summary['measure_fields']} measures)")
                else:
                    self.log_message("⚠️ No table or matrix visuals found in this report")
                    self.show_warning("No Visuals Found", 
                                    "No table or matrix visuals were found in the selected report.\n\n"
                                    "This tool only works with Table (tableEx) and Matrix (pivotTable) visual types.")
            
            def on_error(error):
                self.log_message(f"❌ Failed to scan visuals: {error}")
                self.show_error("Scan Error", f"Failed to scan visuals:\n\n{error}")
            
            self.run_in_background(scan_operation, on_success, on_error)
            
        except Exception as e:
            self.log_message(f"❌ Scan error: {e}")
            self.show_error("Scan Error", str(e))
    
    def _update_visual_tree(self):
        """Update the visual selection tree with proper page ordering"""
        # Clear existing items
        for item in self.visual_tree.get_children():
            self.visual_tree.delete(item)
        
        self.visual_selection_vars.clear()
        
        # Group visuals by page and sort by page order
        page_visuals = {}
        for visual_info in self.visuals_info:
            page_name = visual_info.page_name
            if page_name not in page_visuals:
                page_visuals[page_name] = []
            page_visuals[page_name].append(visual_info)
        
        # Sort pages by their natural order (Page 1, Page 2, etc.) or alphabetically
        sorted_pages = sorted(page_visuals.keys(), key=lambda x: self._extract_page_sort_key(x))
        
        # Add visuals to tree in page order
        for page_name in sorted_pages:
            for visual_info in page_visuals[page_name]:
                # Create selection variable
                var = tk.BooleanVar(value=True)  # Default to selected
                self.visual_selection_vars[visual_info.visual_id] = var
                
                # Determine visual type display
                if visual_info.visual_type == VisualType.TABLE:
                    type_text = "Table"
                else:
                    type_text = f"Matrix ({visual_info.layout_type})"
                
                # Count fields by type
                categorical_count = sum(1 for f in visual_info.fields if f.field_type == FieldType.CATEGORICAL)
                measure_count = sum(1 for f in visual_info.fields if f.field_type == FieldType.MEASURE)
                fields_text = f"{categorical_count}C + {measure_count}M"
                
                # Determine config status
                config_status = "Custom" if visual_info.visual_id in self.visual_config_vars else "Global"

                # Insert item with SVG checkbox icon (default selected)
                # #0 column is just for checkbox icon, visual_name is in values
                if self._checkbox_checked_icon:
                    item_id = self.visual_tree.insert('', 'end',
                                                      text='',  # #0 column empty, just icon
                                                      image=self._checkbox_checked_icon,
                                                      values=(visual_info.visual_name, type_text, visual_info.page_name, fields_text, config_status),
                                                      tags=(visual_info.visual_id, 'selected'))
                else:
                    # Fallback to emoji if icons not available
                    item_id = self.visual_tree.insert('', 'end',
                                                      text='',
                                                      values=(f"☑ {visual_info.visual_name}", type_text, visual_info.page_name, fields_text, config_status),
                                                      tags=(visual_info.visual_id, 'selected'))
    
    def _extract_page_sort_key(self, page_name: str) -> tuple:
        """Extract sorting key for page names to match tab order"""
        # Try to extract page number for proper sorting
        import re
        match = re.search(r'(\d+)', page_name)
        if match:
            page_num = int(match.group(1))
            return (page_num, page_name)  # Sort by number first, then name
        else:
            return (999, page_name)  # Non-numbered pages go to end
    
    def _update_scan_summary(self):
        """Update the scan summary display in the Analysis Summary panel"""
        if not hasattr(self, 'log_section') or not self.log_section:
            return

        summary_text = self.log_section.summary_text
        placeholder = self.log_section.placeholder_label

        if not summary_text:
            return

        # Hide placeholder and show summary text
        if placeholder:
            placeholder.grid_remove()

        summary_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Get scan results
        summary_text.config(state=tk.NORMAL)
        summary_text.delete(1.0, tk.END)

        if not self.visuals_info:
            summary_text.insert(tk.END, "No Visuals Found\n\n", 'header')
            summary_text.insert(tk.END, "No table or matrix visuals were found in this report.\n", 'info')
        else:
            # Get summary data from engine
            summary = self.engine.get_visual_summary() if self.engine else {}
            table_count = summary.get('table_count', 0)
            matrix_count = summary.get('matrix_count', 0)
            total_fields = summary.get('total_fields', 0)
            categorical_fields = summary.get('categorical_fields', 0)
            measure_fields = summary.get('measure_fields', 0)

            # Header - shorter divider to avoid wrapping
            summary_text.insert(tk.END, "Scan Results\n", 'header')
            summary_text.insert(tk.END, "─" * 28 + "\n\n", 'separator')

            # Configure tab stops for consistent two-column layout
            # Left column ~140px, right column starts at ~180px
            summary_text.configure(tabs=('180p',))

            # Side-by-side: Visuals Found (left) | Field Analysis (right)
            # Pad left column items to fixed width for alignment
            summary_text.insert(tk.END, "Visuals Found:", 'subheader')
            summary_text.insert(tk.END, "\t", 'item')
            summary_text.insert(tk.END, "Field Analysis:\n", 'subheader')

            # Fixed-width formatting for left column (20 chars) - aligns right column
            summary_text.insert(tk.END, f"  Tables: {table_count}", 'item')
            summary_text.insert(tk.END, "\t", 'item')
            summary_text.insert(tk.END, f"Categorical: {categorical_fields}\n", 'item')

            summary_text.insert(tk.END, f"  Matrices: {matrix_count}", 'item')
            summary_text.insert(tk.END, "\t", 'item')
            summary_text.insert(tk.END, f"Measures: {measure_fields}\n", 'item')

            summary_text.insert(tk.END, f"  Total: {len(self.visuals_info)} visuals", 'summary')
            summary_text.insert(tk.END, "\t", 'item')
            summary_text.insert(tk.END, f"Total: {total_fields}\n\n", 'summary')

            # Configuration hint - left aligned below
            summary_text.insert(tk.END, "Configuration:\n", 'subheader')
            summary_text.insert(tk.END, "  Use Global Settings for all visuals, or double-click a visual for per-visual config\n", 'info')

        # Apply theme-appropriate colors
        self._apply_summary_text_colors()

        summary_text.config(state=tk.DISABLED)

    def _apply_summary_text_colors(self):
        """Apply theme-appropriate colors to summary text tags"""
        if not hasattr(self, 'log_section') or not self.log_section:
            return

        summary_text = self.log_section.summary_text
        if not summary_text:
            return

        # Use centralized theme colors
        colors = self._theme_manager.colors
        title_color = colors['title_color']
        text_primary = colors['text_primary']
        text_secondary = colors['text_secondary']
        text_muted = colors['text_muted']
        success_color = colors['success']

        summary_text.tag_config('header', font=('Segoe UI', 11, 'bold'), foreground=title_color)
        summary_text.tag_config('subheader', font=('Segoe UI', 10, 'bold'), foreground=text_primary)
        summary_text.tag_config('separator', foreground=text_muted)
        summary_text.tag_config('item', font=('Segoe UI', 9), foreground=text_primary)
        summary_text.tag_config('info', font=('Segoe UI', 9, 'italic'), foreground=text_secondary)
        summary_text.tag_config('summary', font=('Segoe UI', 9, 'bold'), foreground=success_color)
    
    def _enable_configuration_ui(self):
        """Enable configuration UI after successful scan"""
        self.preview_button.set_enabled(True)
        self.apply_button.set_enabled(True)
    
    def _on_tree_double_click(self, event):
        """Handle double-click for per-visual configuration"""
        item = self.visual_tree.identify('item', event.x, event.y)
        if item:
            tags = self.visual_tree.item(item, 'tags')
            if tags:
                visual_id = tags[0]
                visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
                if visual_info:
                    self._show_per_visual_config_dialog(visual_info)
    
    def _show_per_visual_config_dialog(self, visual_info: VisualInfo):
        """Show per-visual configuration dialog"""
        colors = self._theme_manager.colors

        config_window = tk.Toplevel(self.main_app.root)
        config_window.withdraw()  # Hide until fully styled (prevents white flash)
        config_window.title(f"Configure: {visual_info.visual_name}")

        # Center dialog in main window
        dialog_width, dialog_height = 550, 400
        parent_x = self.main_app.root.winfo_rootx()
        parent_y = self.main_app.root.winfo_rooty()
        parent_width = self.main_app.root.winfo_width()
        parent_height = self.main_app.root.winfo_height()
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        config_window.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        config_window.resizable(False, False)
        config_window.transient(self.main_app.root)
        config_window.grab_set()

        # Apply theme-aware background for popup outer background
        # Popup outer background should use 'background' (white in light mode, dark in dark mode)
        is_dark = self._theme_manager.is_dark
        popup_outer_bg = colors['background']
        config_window.configure(bg=popup_outer_bg)

        # Set AE favicon icon
        try:
            config_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Set dark title bar if dark mode
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(config_window, self._theme_manager.is_dark)

        # Main frame - use tk.Frame with explicit background for consistency
        main_frame = tk.Frame(config_window, bg=popup_outer_bg, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header title
        tk.Label(main_frame, text=f"Configure: {visual_info.visual_name}",
                 font=('Segoe UI', 12, 'bold'), bg=popup_outer_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 10))

        visual_type = "Table" if visual_info.visual_type == VisualType.TABLE else "Matrix"
        tk.Label(main_frame, text=f"Type: {visual_type} | Page: {visual_info.page_name}",
                 font=('Segoe UI', 9), bg=popup_outer_bg,
                 fg=colors['text_secondary']).pack(anchor=tk.W, pady=(0, 15))

        # Get or create per-visual config variables
        if visual_info.visual_id not in self.visual_config_vars:
            self._create_per_visual_config_vars(visual_info.visual_id)

        visual_vars = self.visual_config_vars[visual_info.visual_id]

        # Get theme colors for bordered frames (same as main tool)
        frame_bg = colors['section_bg']
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#e0e0e0')

        # Configuration sections - use popup_outer_bg for consistency
        config_frame = tk.Frame(main_frame, bg=popup_outer_bg)
        config_frame.pack(fill=tk.X, pady=(0, 15))

        # Categorical columns (left) - tk.Frame with border for visual grouping
        cat_outer = tk.Frame(config_frame, bg=frame_bg,
                            highlightbackground=border_color, highlightcolor=border_color,
                            highlightthickness=1)
        cat_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # Inner padding frame
        cat_frame = tk.Frame(cat_outer, bg=frame_bg, padx=10, pady=8)
        cat_frame.pack(fill=tk.BOTH, expand=True)

        # Title label
        tk.Label(cat_frame, text="Categorical", bg=frame_bg,
                fg=colors['title_color'], font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(0, 8))

        self._setup_width_controls_with_custom_for_popup(cat_frame, visual_vars['categorical_preset'], 'categorical', visual_info.visual_id)

        # Measure columns (right) - tk.Frame with border for visual grouping
        measure_outer = tk.Frame(config_frame, bg=frame_bg,
                                highlightbackground=border_color, highlightcolor=border_color,
                                highlightthickness=1)
        measure_outer.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # Inner padding frame
        measure_frame = tk.Frame(measure_outer, bg=frame_bg, padx=10, pady=8)
        measure_frame.pack(fill=tk.BOTH, expand=True)

        # Title label
        tk.Label(measure_frame, text="Measure", bg=frame_bg,
                fg=colors['title_color'], font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(0, 8))

        self._setup_width_controls_with_custom_for_popup(measure_frame, visual_vars['measure_preset'], 'measure', visual_info.visual_id)

        # Buttons with RoundedButton styling - use tk.Frame for consistent background
        button_frame = tk.Frame(main_frame, bg=popup_outer_bg)
        button_frame.pack(fill=tk.X, pady=(20, 0))

        # Popup buttons sit on main window background, need explicit canvas_bg
        # Dark mode: #161627, Light mode: #ffffff (same as popup_outer_bg)
        popup_canvas_bg = popup_outer_bg

        # Apply button (primary style)
        execute_icon = self._button_icons.get('execute')
        apply_btn = RoundedButton(
            button_frame, text="Apply",
            command=lambda: self._apply_per_visual_config(visual_info.visual_id, config_window),
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 9, 'bold'),
            icon=execute_icon,
            canvas_bg=popup_canvas_bg
        )
        apply_btn.pack(side=tk.LEFT)

        # Apply Globally button (secondary style - button_secondary colors like Reset All)
        earth_icon = self._button_icons.get('earth')
        apply_global_btn = RoundedButton(
            button_frame, text="Apply Globally",
            command=lambda: self._copy_to_global_config(visual_info.visual_id),
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            icon=earth_icon,
            canvas_bg=popup_canvas_bg
        )
        apply_global_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Reset to Global button (secondary style - button_secondary colors like Reset All)
        reset_icon = self._button_icons.get('reset')
        reset_btn = RoundedButton(
            button_frame, text="Reset",
            command=lambda: self._reset_to_global_config(visual_info.visual_id, config_window),
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            icon=reset_icon,
            canvas_bg=popup_canvas_bg
        )
        reset_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Cancel button (secondary style - button_secondary colors like Reset All)
        cancel_btn = RoundedButton(
            button_frame, text="Cancel",
            command=config_window.destroy,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            canvas_bg=popup_canvas_bg
        )
        cancel_btn.pack(side=tk.RIGHT)

        # Show dialog now that it's fully styled
        config_window.deiconify()

    def _create_per_visual_config_vars(self, visual_id: str):
        """Create per-visual configuration variables"""
        self.visual_config_vars[visual_id] = {
            'categorical_preset': tk.StringVar(value=self.global_categorical_preset_var.get()),
            'measure_preset': tk.StringVar(value=self.global_measure_preset_var.get()),
            'categorical_custom': tk.StringVar(value="105"),
            'measure_custom': tk.StringVar(value="95")
        }
    
    def _apply_per_visual_config(self, visual_id: str, config_window):
        """Apply per-visual configuration and switch to it"""
        self.current_selected_visual = visual_id
        self._switch_to_per_visual_config(visual_id)
        
        # Update config mode label
        visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
        if visual_info:
            self.config_mode_label.config(text=f"Per-Visual: {visual_info.visual_name}")
        
        # Update tree display to show config status change
        self._update_tree_display()
        
        # Log the specific configuration settings that were applied
        if visual_id in self.visual_config_vars:
            visual_vars = self.visual_config_vars[visual_id]
            cat_setting = visual_vars['categorical_preset'].get()
            measure_setting = visual_vars['measure_preset'].get()
            
            self.log_message(f"⚙️ Applied per-visual configuration for {visual_info.visual_name if visual_info else visual_id}")
            
            # Show categorical setting with custom value if needed
            if cat_setting == WidthPreset.CUSTOM.value:
                cat_custom = visual_vars.get('categorical_custom', tk.StringVar(value="105")).get()
                self.log_message(f"📊 Categorical columns: {cat_setting} ({cat_custom}px)")
            else:
                self.log_message(f"📊 Categorical columns: {cat_setting}")
            
            # Show measure setting with custom value if needed
            if measure_setting == WidthPreset.CUSTOM.value:
                measure_custom = visual_vars.get('measure_custom', tk.StringVar(value="95")).get()
                self.log_message(f"📈 Measure columns: {measure_setting} ({measure_custom}px)")
            else:
                self.log_message(f"📈 Measure columns: {measure_setting}")
            
            self.log_message("ℹ️ This visual will use its own width settings when changes are applied")
        
        config_window.destroy()
    
    def _copy_to_global_config(self, visual_id: str):
        """Copy per-visual configuration to global settings"""
        if visual_id in self.visual_config_vars:
            visual_vars = self.visual_config_vars[visual_id]
            
            # Get the settings being copied
            cat_setting = visual_vars['categorical_preset'].get()
            measure_setting = visual_vars['measure_preset'].get()
            
            # Copy values to global variables
            self.global_categorical_preset_var.set(cat_setting)
            self.global_measure_preset_var.set(measure_setting)
            
            # Copy custom values too
            if 'categorical_custom' in visual_vars:
                self.categorical_custom_var.set(visual_vars['categorical_custom'].get())
            if 'measure_custom' in visual_vars:
                self.measure_custom_var.set(visual_vars['measure_custom'].get())
            
            # Switch back to global view
            self._switch_to_global_config()
            
            # Update tree display to potentially show config status changes
            self._update_tree_display()
            
            self.log_message("🌐 Per-visual configuration copied to global settings")
            
            # Show settings with custom values if applicable
            if cat_setting == WidthPreset.CUSTOM.value:
                cat_custom = visual_vars.get('categorical_custom', tk.StringVar(value="105")).get()
                self.log_message(f"📊 Categorical columns: {cat_setting} ({cat_custom}px)")
            else:
                self.log_message(f"📊 Categorical columns: {cat_setting}")
            
            if measure_setting == WidthPreset.CUSTOM.value:
                measure_custom = visual_vars.get('measure_custom', tk.StringVar(value="95")).get()
                self.log_message(f"📈 Measure columns: {measure_setting} ({measure_custom}px)")
            else:
                self.log_message(f"📈 Measure columns: {measure_setting}")
            
            self.log_message("📝 Now showing global configuration mode")
    
    def _reset_to_global_config(self, visual_id: str, config_window):
        """Reset visual to use global configuration"""
        # Remove per-visual configuration if it exists
        if visual_id in self.visual_config_vars:
            del self.visual_config_vars[visual_id]
        
        # Switch back to global configuration if this was the current visual
        if self.current_selected_visual == visual_id:
            self.current_selected_visual = None
            self._switch_to_global_config()
        
        # Update tree display to show config status change
        self._update_tree_display()
        
        config_window.destroy()
        
        visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
        self.log_message(f"🔄 Reset {visual_info.visual_name if visual_info else visual_id} to use global settings")
        self.log_message("🌐 This visual will now use global configuration when changes are applied")
    
    def _switch_to_global_config(self):
        """Switch configuration UI to global settings"""
        self.categorical_preset_var = self.global_categorical_preset_var
        self.measure_preset_var = self.global_measure_preset_var
        
        self.config_mode_label.config(text="Global Settings (All Visuals)")
    
    def _switch_to_per_visual_config(self, visual_id: str):
        """Switch configuration UI to per-visual settings"""
        if visual_id in self.visual_config_vars:
            visual_vars = self.visual_config_vars[visual_id]
            
            self.categorical_preset_var = visual_vars['categorical_preset']
            self.measure_preset_var = visual_vars['measure_preset']
    
    def _on_tree_click(self, event):
        """Handle tree item clicks for selection - only checkbox area should toggle"""
        item = self.visual_tree.identify('item', event.x, event.y)
        if item:
            # Only toggle if clicking in the checkbox area (first ~30 pixels)
            if event.x <= 30:
                # Get visual ID from tags
                tags = self.visual_tree.item(item, 'tags')
                if tags:
                    visual_id = tags[0]
                    if visual_id in self.visual_selection_vars:
                        # Toggle selection
                        var = self.visual_selection_vars[visual_id]
                        var.set(not var.get())
                        
                        # Update display and styling
                        visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
                        if visual_info:
                            self._update_tree_item_display(item, visual_info, var.get())
            # If clicking outside checkbox area, do nothing (allows double-click to work properly)
    
    def _update_tree_item_display(self, item, visual_info, is_selected):
        """Update tree item display with current selection and config status"""
        # Use SVG checkbox icons instead of emoji characters
        checkbox_icon = self._checkbox_checked_icon if is_selected else self._checkbox_unchecked_icon

        # Determine config status
        config_status = "Custom" if visual_info.visual_id in self.visual_config_vars else "Global"

        # Get current values for other columns
        # Values: (visual_name, type, page, fields, config) - 5 columns
        current_values = list(self.visual_tree.item(item, 'values'))
        if len(current_values) >= 5:
            current_values[4] = config_status  # Update config column (index 4)
        else:
            current_values.extend([''] * (5 - len(current_values)))
            current_values[4] = config_status

        # Update the item with SVG image and tags
        # #0 column is just for checkbox icon, visual_name stays in values[0]
        tag = 'selected' if is_selected else 'unselected'
        if checkbox_icon:
            self.visual_tree.item(item, text='', image=checkbox_icon,
                                  values=current_values, tags=(visual_info.visual_id, tag))
        else:
            # Fallback to emoji if icons not available - prepend to visual_name in values
            icon = "☑" if is_selected else "☐"
            current_values[0] = f"{icon} {visual_info.visual_name}"
            self.visual_tree.item(item, text='', values=current_values, tags=(visual_info.visual_id, tag))
    
    def _on_tree_hover(self, event):
        """Handle mouse hover over tree items - shows hand cursor and underline only"""
        item = self.visual_tree.identify('item', event.x, event.y)

        # Remove hover tag from previously hovered item
        if self._hovered_tree_item and self._hovered_tree_item != item:
            old_tags = list(self.visual_tree.item(self._hovered_tree_item, 'tags'))
            if 'hover' in old_tags:
                old_tags.remove('hover')
                self.visual_tree.item(self._hovered_tree_item, tags=tuple(old_tags))

        if item:
            self._hovered_tree_item = item
            # Change cursor to hand to indicate clickable
            self.visual_tree.config(cursor='hand2')
            # Add hover tag for underline effect only (no selection/color change)
            current_tags = list(self.visual_tree.item(item, 'tags'))
            if 'hover' not in current_tags:
                current_tags.append('hover')
                self.visual_tree.item(item, tags=tuple(current_tags))
        else:
            self._hovered_tree_item = None
            self.visual_tree.config(cursor='')

    def _on_tree_leave(self, event):
        """Handle mouse leaving the tree - clear hover effects"""
        # Remove hover tag from hovered item
        if self._hovered_tree_item:
            current_tags = list(self.visual_tree.item(self._hovered_tree_item, 'tags'))
            if 'hover' in current_tags:
                current_tags.remove('hover')
                self.visual_tree.item(self._hovered_tree_item, tags=tuple(current_tags))
        # Reset cursor
        self.visual_tree.config(cursor='')
        self._hovered_tree_item = None
    
    def _select_all_visuals(self):
        """Select all visuals"""
        for var in self.visual_selection_vars.values():
            var.set(True)
        self._update_tree_display()
        self.log_message("✓ Selected all visuals")
    
    def _clear_all_visuals(self):
        """Clear all visual selections"""
        for var in self.visual_selection_vars.values():
            var.set(False)
        self._update_tree_display()
        self.log_message("✗ Cleared all visual selections")
    
    def _select_tables_only(self):
        """Select only table visuals"""
        for visual_info in self.visuals_info:
            var = self.visual_selection_vars.get(visual_info.visual_id)
            if var:
                var.set(visual_info.visual_type == VisualType.TABLE)
        self._update_tree_display()
        self.log_message("📊 Selected tables only")
    
    def _select_matrices_only(self):
        """Select only matrix visuals"""
        for visual_info in self.visuals_info:
            var = self.visual_selection_vars.get(visual_info.visual_id)
            if var:
                var.set(visual_info.visual_type == VisualType.MATRIX)
        self._update_tree_display()
        self.log_message("📊 Selected matrices only")
    
    def _update_tree_display(self):
        """Update tree display after selection changes"""
        for item in self.visual_tree.get_children():
            tags = self.visual_tree.item(item, 'tags')
            if tags:
                visual_id = tags[0]
                var = self.visual_selection_vars.get(visual_id)
                visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
                
                if var and visual_info:
                    self._update_tree_item_display(item, visual_info, var.get())
    
    def _get_selected_visual_configs(self) -> Dict[str, Any]:
        """Get configuration for each selected visual (either per-visual or global)"""
        from tools.column_width.column_width_core import WidthConfiguration, WidthPreset
        
        configs = {}
        selected_ids = self._get_selected_visual_ids()
        
        for visual_id in selected_ids:
            if visual_id in self.visual_config_vars:
                # Use per-visual configuration
                visual_vars = self.visual_config_vars[visual_id]
                config = WidthConfiguration()
                config.categorical_preset = WidthPreset(visual_vars['categorical_preset'].get())
                config.measure_preset = WidthPreset(visual_vars['measure_preset'].get())
                
                # Get custom width values for this visual
                try:
                    config.categorical_custom = int(visual_vars['categorical_custom'].get())
                except (ValueError, KeyError):
                    config.categorical_custom = 105  # Default fallback
                    
                try:
                    config.measure_custom = int(visual_vars['measure_custom'].get())
                except (ValueError, KeyError):
                    config.measure_custom = 95  # Default fallback
                
                configs[visual_id] = config
            else:
                # Will use global configuration - don't add to configs dict
                pass
        
        return configs
    
    def _get_selected_visual_ids(self) -> List[str]:
        """Get list of selected visual IDs"""
        selected = []
        for visual_id, var in self.visual_selection_vars.items():
            if var.get():
                selected.append(visual_id)
        return selected
    
    def _get_global_config(self) -> Any:
        """Get global width configuration"""
        from tools.column_width.column_width_core import WidthConfiguration, WidthPreset
        
        config = WidthConfiguration()
        config.categorical_preset = WidthPreset(self.global_categorical_preset_var.get())
        config.measure_preset = WidthPreset(self.global_measure_preset_var.get())
        
        # Get custom width values
        try:
            config.categorical_custom = int(self.categorical_custom_var.get())
        except ValueError:
            config.categorical_custom = 105  # Default fallback
            
        try:
            config.measure_custom = int(self.measure_custom_var.get())
        except ValueError:
            config.measure_custom = 95  # Default fallback
        
        return config
    
    def _preview_changes(self):
        """Preview the changes that would be made"""
        if not self.engine:
            self.show_error("No Data", "Please scan visuals first.")
            return
            
        selected_ids = self._get_selected_visual_ids()
        if not selected_ids:
            self.show_warning("No Selection", "Please select at least one visual to preview.")
            return
            
        self.log_message(f"👁️ Generating preview for {len(selected_ids)} visual(s)...")
        
        # Check for custom values below minimum and warn
        per_visual_configs = self._get_selected_visual_configs()
        global_config = self._get_global_config()
        
        # Warn about global custom values below minimum
        if global_config.categorical_preset == WidthPreset.CUSTOM and global_config.categorical_custom < 50:
            self.log_message(f"⚠️ Global categorical custom width {global_config.categorical_custom}px increased to minimum 50px")
        if global_config.measure_preset == WidthPreset.CUSTOM and global_config.measure_custom < 50:
            self.log_message(f"⚠️ Global measure custom width {global_config.measure_custom}px increased to minimum 50px")
        
        # Warn about per-visual custom values below minimum
        for visual_id, config in per_visual_configs.items():
            visual_info = next((v for v in self.visuals_info if v.visual_id == visual_id), None)
            visual_name = visual_info.visual_name if visual_info else visual_id
            
            if config.categorical_preset == WidthPreset.CUSTOM and config.categorical_custom < 50:
                self.log_message(f"⚠️ {visual_name}: categorical custom width {config.categorical_custom}px increased to minimum 50px")
            if config.measure_preset == WidthPreset.CUSTOM and config.measure_custom < 50:
                self.log_message(f"⚠️ {visual_name}: measure custom width {config.measure_custom}px increased to minimum 50px")
        
        # Calculate preview data
        preview_data = []
        
        for visual_info in self.visuals_info:
            if visual_info.visual_id in selected_ids:
                # Determine which configuration to use - check if this specific visual has per-visual config
                if visual_info.visual_id in per_visual_configs:
                    config = per_visual_configs[visual_info.visual_id]
                    config_source = "Custom"
                else:
                    config = global_config
                    config_source = "Global"
                
                # Calculate optimal widths
                self.engine.calculate_optimal_widths(visual_info, config)
                preview_data.append((visual_info, config_source))
        
        # Show preview dialog
        self._show_preview_dialog(preview_data)
    
    def _show_preview_dialog(self, preview_data):
        """Show preview dialog with calculated changes"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Preserve parent column widths before opening dialog (prevents layout shift)
        saved_left_width = getattr(self, '_locked_left_width', None)
        saved_right_width = getattr(self, '_locked_right_width', None)

        # Force-apply saved widths BEFORE creating dialog to prevent drift (weight=0 disables expansion)
        if saved_left_width and saved_right_width and hasattr(self, '_middle_content'):
            self._middle_content.columnconfigure(0, weight=0, minsize=saved_left_width)
            self._middle_content.columnconfigure(1, weight=0, minsize=saved_right_width)
            # ALSO re-apply explicit frame widths and propagate settings
            if hasattr(self, '_left_column'):
                self._left_column.configure(width=saved_left_width)
                self._left_column.grid_propagate(False)
            if hasattr(self, '_right_column'):
                self._right_column.configure(width=saved_right_width)
                self._right_column.grid_propagate(False)
            self.main_app.root.update_idletasks()  # Ensure geometry is stable

        preview_window = tk.Toplevel(self.main_app.root)
        preview_window.withdraw()  # Hide until fully styled (prevents white flash)
        preview_window.title("Preview Column Width Changes")

        # Center dialog in main window
        dialog_width, dialog_height = 900, 600
        parent_x = self.main_app.root.winfo_rootx()
        parent_y = self.main_app.root.winfo_rooty()
        parent_width = self.main_app.root.winfo_width()
        parent_height = self.main_app.root.winfo_height()
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        preview_window.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        preview_window.resizable(True, False)  # Allow width resize but lock height
        preview_window.minsize(600, 600)  # Lock minimum height
        preview_window.maxsize(1600, 600)  # Lock maximum height (allow width to vary)
        preview_window.transient(self.main_app.root)
        preview_window.grab_set()

        # Restore parent column widths on dialog close to prevent layout drift (weight=0)
        def on_close():
            if saved_left_width and saved_right_width and hasattr(self, '_middle_content'):
                self._middle_content.columnconfigure(0, weight=0, minsize=saved_left_width)
                self._middle_content.columnconfigure(1, weight=0, minsize=saved_right_width)
                # ALSO re-apply explicit frame widths and propagate settings
                if hasattr(self, '_left_column'):
                    self._left_column.configure(width=saved_left_width)
                    self._left_column.grid_propagate(False)
                if hasattr(self, '_right_column'):
                    self._right_column.configure(width=saved_right_width)
                    self._right_column.grid_propagate(False)
            preview_window.destroy()

        preview_window.protocol("WM_DELETE_WINDOW", on_close)

        # Apply theme-aware background (white outer background in light mode)
        outer_bg = '#ffffff' if not is_dark else colors['background']
        preview_window.configure(bg=outer_bg)

        # Set AE favicon
        try:
            preview_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Set dark title bar if dark mode
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(preview_window, is_dark)

        # Main frame (use tk.Frame with explicit background for light mode white)
        main_frame = tk.Frame(preview_window, bg=outer_bg, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header with eye icon
        header_frame = tk.Frame(main_frame, bg=outer_bg)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        eye_icon = self._button_icons.get('eye')
        if eye_icon:
            icon_label = tk.Label(header_frame, image=eye_icon, bg=outer_bg)
            icon_label.pack(side=tk.LEFT, padx=(0, 8))
            icon_label._icon_ref = eye_icon

        tk.Label(header_frame, text="Preview Column Width Changes",
                 font=('Segoe UI', 14, 'bold'),
                 fg=colors['title_color'], bg=outer_bg).pack(side=tk.LEFT)

        tk.Label(main_frame, text=f"Changes for {len(preview_data)} selected visual(s)",
                 font=('Segoe UI', 9),
                 fg=colors['text_secondary'], bg=outer_bg).pack(anchor=tk.W, pady=(0, 15))

        # Preview tree with flat design and themed scrollbar (matching Visual Selection styling)
        tree_frame = tk.Frame(main_frame, bg=outer_bg)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Set treeview colors based on theme (matching Visual Selection table)
        if is_dark:
            tree_bg = colors.get('surface', '#1e1e2e')
            tree_fg = colors.get('text_primary', '#e0e0e0')
            tree_field_bg = colors.get('surface', '#1e1e2e')
            heading_bg = colors.get('section_bg', '#1a1a2a')
            heading_fg = colors.get('text_primary', '#e0e0e0')
            tree_border = colors.get('border', '#3a3a4a')
            header_separator = '#0d0d1a'
            # Selection colors for dark mode
            select_bg = '#3a3a5a'  # Darker selection background
            select_fg = '#ffffff'  # White text on dark selection
        else:
            tree_bg = colors.get('surface', '#ffffff')
            tree_fg = colors.get('text_primary', '#333333')
            tree_field_bg = colors.get('surface', '#ffffff')
            heading_bg = colors.get('section_bg', '#f5f5fa')
            heading_fg = colors.get('text_primary', '#333333')
            tree_border = '#d8d8e0'
            header_separator = '#ffffff'
            # Selection colors for light mode (keep text dark, not white)
            select_bg = '#cce4f7'  # Light blue selection background
            select_fg = '#333333'  # Dark text stays readable on selection

        # Configure preview treeview style - flat, no 3D effects (matching Visual Selection)
        style = ttk.Style()
        preview_tree_style = "PreviewDialog.Treeview"
        style.configure(preview_tree_style,
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=tree_field_bg,
                        font=('Segoe UI', 9),
                        rowheight=25,
                        relief='flat',
                        borderwidth=0,
                        bordercolor=tree_bg,
                        lightcolor=tree_bg,
                        darkcolor=tree_bg)

        # Remove treeview frame border by styling the Item element
        style.layout(preview_tree_style, [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

        # Heading style with subtle column separators
        style.configure(f"{preview_tree_style}.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        font=('Segoe UI', 9, 'bold'),
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        padding=(8, 4))

        style.map(f"{preview_tree_style}.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])

        # Selection colors for the treeview rows
        style.map(preview_tree_style,
                  background=[('selected', select_bg)],
                  foreground=[('selected', select_fg)])

        # Container with 1px border (matching Visual Selection style)
        tree_container = tk.Frame(tree_frame, bg=tree_bg,
                                  highlightbackground=tree_border, highlightcolor=tree_border,
                                  highlightthickness=1)
        tree_container.pack(fill=tk.BOTH, expand=True)

        preview_tree = ttk.Treeview(tree_container, height=15, style=preview_tree_style)
        preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=0, pady=0)

        # Configure columns
        preview_tree['columns'] = ('field_type', 'current_width', 'new_width', 'change')
        preview_tree.heading('#0', text='Page / Visual / Field')
        preview_tree.heading('field_type', text='Type')
        preview_tree.heading('current_width', text='Current')
        preview_tree.heading('new_width', text='New Width')
        preview_tree.heading('change', text='Change')

        # Configure column widths
        preview_tree.column('#0', width=300)
        preview_tree.column('field_type', width=80)
        preview_tree.column('current_width', width=80)
        preview_tree.column('new_width', width=100)
        preview_tree.column('change', width=80)

        # Add ThemedScrollbar with auto-hide (matching Visual Selection style)
        preview_scrollbar = ThemedScrollbar(tree_container,
                                            command=preview_tree.yview,
                                            theme_manager=self._theme_manager,
                                            auto_hide=True)
        preview_tree.configure(yscrollcommand=preview_scrollbar.set)
        
        # Populate preview data - organize by Page > Visual > Field hierarchy
        page_groups = {}
        
        # Group visuals by page
        for visual_info, config_source in preview_data:
            page_name = visual_info.page_name
            if page_name not in page_groups:
                page_groups[page_name] = []
            page_groups[page_name].append((visual_info, config_source))
        
        # Sort pages by their natural order
        sorted_pages = sorted(page_groups.keys(), key=lambda x: self._extract_page_sort_key(x))
        
        # Create hierarchy: Page > Visual > Field (no emojis)
        for page_name in sorted_pages:
            # Add page as top-level parent
            page_item = preview_tree.insert('', 'end',
                                           text=page_name,
                                           values=('', '', '', ''),
                                           open=True)

            # Add visuals under each page
            for visual_info, config_source in page_groups[page_name]:
                # Visual display with config type
                visual_display = f"{visual_info.visual_name} ({config_source})"
                visual_item = preview_tree.insert(page_item, 'end',
                                                  text=visual_display,
                                                  values=('', '', '', ''),
                                                  open=True)

                # Add fields under each visual
                for field in visual_info.fields:
                    if field.suggested_width is not None:
                        field_type = "Category" if field.field_type == FieldType.CATEGORICAL else "Measure"
                        current = f"{field.current_width:.0f}px" if field.current_width else "Not set"
                        new_width = f"{field.suggested_width:.0f}px"

                        # Calculate change
                        if field.current_width:
                            change_px = field.suggested_width - field.current_width
                            change = f"{change_px:+.0f}px"
                        else:
                            change = "New"

                        preview_tree.insert(visual_item, 'end',
                                            text=f"    {field.display_name}",
                                            values=(field_type, current, new_width, change))
        
        # Buttons with RoundedButton styling
        button_frame = tk.Frame(main_frame, bg=outer_bg)
        button_frame.pack(fill=tk.X)

        # Close button (secondary style - button_secondary colors like Reset All)
        close_btn = RoundedButton(
            button_frame, text="Close",
            command=on_close,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=36, radius=6, font=('Segoe UI', 10)
        )
        close_btn.pack(side=tk.RIGHT)

        # Apply button (primary style with execute icon)
        execute_icon = self._button_icons.get('execute')
        apply_btn = RoundedButton(
            button_frame, text="Apply Changes",
            command=lambda: [on_close(), self._apply_changes()],
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=36, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=execute_icon
        )
        apply_btn.pack(side=tk.RIGHT, padx=(0, 10))

        # Show dialog now that it's fully styled
        preview_window.deiconify()

    def _apply_changes(self):
        """Apply the column width changes"""
        selected_ids = self._get_selected_visual_ids()
        if not selected_ids:
            self.show_warning("No Selection", "Please select at least one visual to update.")
            return
            
        # Confirm action
        message = f"Apply column width changes to {len(selected_ids)} selected visual(s)?\n\n"
        message += "⚠️ This will modify your PBIP file permanently.\n"
        message += "🔒 Auto-size columns will be turned OFF to preserve settings."
        
        if not self.ask_yes_no("Confirm Changes", message):
            return
            
        self.log_message(f"⚙️ Applying changes to {len(selected_ids)} visual(s)...")
        
        # Log the configuration settings being used
        per_visual_configs = self._get_selected_visual_configs()
        global_config = self._get_global_config()
        
        # Show global settings
        self.log_message(f"🌐 Global settings - Categorical: {global_config.categorical_preset.value}, Measure: {global_config.measure_preset.value}")
        
        # Show per-visual overrides if any
        if per_visual_configs:
            self.log_message(f"🎨 {len(per_visual_configs)} visual(s) have custom per-visual settings")
        
        def apply_operation():
            self.update_progress(10, "Preparing configurations...")
            
            # Determine if we're using global or per-visual configurations
            per_visual_configs = self._get_selected_visual_configs()
            global_config = self._get_global_config()
            
            self.update_progress(30, "Applying width changes...")
            
            # Apply changes using the appropriate configuration method
            if any(visual_id in self.visual_config_vars for visual_id in selected_ids):
                # Some visuals have per-visual configurations
                results = self.engine.apply_width_changes(selected_ids, configs=per_visual_configs)
            else:
                # Use global configuration for all
                results = self.engine.apply_width_changes(selected_ids, global_config=global_config)
            
            self.update_progress(80, "Finalizing changes...")
            import time
            time.sleep(0.5)
            
            self.update_progress(100, "Changes applied successfully!")
            return results
        
        def on_success(results):
            if results["success"]:
                self.log_message(f"✅ Successfully updated {results['visuals_updated']} visuals")
                self.log_message(f"📊 Modified {results['fields_updated']} field width settings")
                self.log_message("🔒 Auto-size columns turned OFF to preserve settings")
                
                success_msg = (f"Column width changes applied successfully!\n\n"
                              f"• Visuals updated: {results['visuals_updated']}\n"
                              f"• Fields modified: {results['fields_updated']}\n"
                              f"• Auto-size columns: DISABLED")
                
                if results["errors"]:
                    success_msg += f"\n\nWarnings:\n" + "\n".join(f"• {error}" for error in results["errors"])
                
                self.show_info("Changes Applied", success_msg)
                
                # Optionally rescan to show updated status
                self._scan_visuals()
            else:
                error_msg = "Some errors occurred during application:\n\n" + "\n".join(f"• {error}" for error in results["errors"])
                self.show_error("Application Errors", error_msg)
        
        def on_error(error):
            self.log_message(f"❌ Failed to apply changes: {error}")
            self.show_error("Application Error", f"Failed to apply changes:\n\n{error}")
        
        self.run_in_background(apply_operation, on_success, on_error)
    
    def reset_tab(self):
        """Reset the tab to initial state"""
        # Clear data
        self.engine = None
        self.visuals_info.clear()
        self.visual_selection_vars.clear()
        self.visual_config_vars.clear()

        # Reset UI state
        self.pbip_path_var.set("")
        # Disable scan button (will be enabled when file path is entered)
        self.scan_button.set_enabled(False)
        
        # Reset configuration to defaults
        self.global_categorical_preset_var.set(WidthPreset.AUTO_FIT.value)
        self.global_measure_preset_var.set(WidthPreset.FIT_TO_TOTALS.value)
        
        # Reset to global configuration
        self.current_selected_visual = None
        self._switch_to_global_config()
        
        # Clear tree
        if self.visual_tree:
            for item in self.visual_tree.get_children():
                self.visual_tree.delete(item)
        
        # Disable buttons (RoundedButton uses set_enabled)
        self.preview_button.set_enabled(False)
        self.apply_button.set_enabled(False)

        # Hide progress
        self.update_progress(0, "", False)

        # Clear log and reset split log section
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)

        # Reset summary panel if using SplitLogSection
        if hasattr(self, 'log_section') and self.log_section:
            # Use the clear_summary method which handles hiding text and showing placeholder
            self.log_section.clear_summary()

        self._show_welcome_message()
        
        self.log_message("🔄 Tab reset to initial state")
    
    def show_help_dialog(self):
        """Show help dialog specific to Table Column Widths"""
        # Get the correct parent window
        parent_window = None
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'root'):
            parent_window = self.main_app.root
        elif hasattr(self, 'master'):
            parent_window = self.master
        elif hasattr(self, 'parent'):
            parent_window = self.parent
        else:
            # Fallback - find root window from frame
            parent_window = self.frame.winfo_toplevel()
        
        help_window = tk.Toplevel(parent_window)
        help_window.withdraw()  # Hide until fully styled (prevents white flash)
        help_window.title("Table Column Widths - Help")
        help_window.geometry("1000x760")  # Wider and shorter to match layout optimizer
        help_window.resizable(False, False)
        help_window.transient(parent_window)
        help_window.grab_set()

        # Set AE favicon icon
        try:
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'config'):
                help_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        self._create_help_content(help_window, parent_window)
    
    def _create_help_content(self, help_window, parent_window):
        """Create help content for table column widths"""
        from core.ui_base import RoundedButton
        colors = self._theme_manager.colors

        # Consistent help dialog background for all tools
        is_dark = self._theme_manager.is_dark
        help_bg = colors['background']
        help_window.configure(bg=help_bg)

        # Set dark title bar if dark mode
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, is_dark)

        # Main container - use tk.Frame with explicit bg for consistency
        container = tk.Frame(help_window, bg=help_bg)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header - use tk.Label with explicit bg for dark mode
        tk.Label(container, text="Table Column Widths - Help",
                 font=('Segoe UI', 16, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.CENTER, pady=(0, 15))

        # Orange warning box with white text
        warning_frame = tk.Frame(container, bg=help_bg)
        warning_frame.pack(fill=tk.X, pady=(0, 15))

        warning_bg = colors.get('warning_bg', '#d97706')
        warning_text = colors.get('warning_text', '#ffffff')
        warning_container = tk.Frame(warning_frame, bg=warning_bg,
                                   padx=15, pady=10, relief='flat', borderwidth=0)
        warning_container.pack(fill=tk.X)

        tk.Label(warning_container, text="IMPORTANT DISCLAIMERS & REQUIREMENTS",
                 font=('Segoe UI', 12, 'bold'),
                 bg=warning_bg,
                 fg=warning_text).pack(anchor=tk.W)

        warnings = [
            "This tool ONLY works with PBIP enhanced report format (PBIR) files",
            "This is NOT officially supported by Microsoft - use at your own discretion",
            "Requires TMDL files in semantic model definition folder",
            "Always keep backups of your original reports before optimization",
            "Test thoroughly and validate results before production use",
            "Enable 'Store reports using enhanced metadata format (PBIR)' in Power BI Desktop"
        ]

        for warning in warnings:
            tk.Label(warning_container, text=warning, font=('Segoe UI', 10),
                     bg=warning_bg,
                     fg=warning_text).pack(anchor=tk.W, padx=(12, 0), pady=1)
        
        # Single 2-column, 2-row grid for proper vertical alignment
        sections_frame = tk.Frame(container, bg=help_bg)
        sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        sections_frame.columnconfigure(0, weight=1)
        sections_frame.columnconfigure(1, weight=1)

        # ROW 0, COLUMN 0: What This Tool Does
        left_top_frame = tk.Frame(sections_frame, bg=help_bg)
        left_top_frame.grid(row=0, column=0, sticky='nwe', padx=(0, 10), pady=(0, 15))

        tk.Label(left_top_frame, text="What This Tool Does",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        what_items = [
            "Standardizes column widths across all tables and matrices in your report",
            "Provides intelligent width presets (Narrow, Medium, Wide, Fit to Header, Fit to Totals)",
            "Supports both global configuration and per-visual customization",
            "Auto-sizes columns based on headers or totals for optimal fit",
            "Disables auto-size to preserve your width settings",
            "Provides preview before applying changes"
        ]

        for item in what_items:
            tk.Label(left_top_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=450,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        # ROW 0, COLUMN 1: File Requirements
        right_top_frame = tk.Frame(sections_frame, bg=help_bg)
        right_top_frame.grid(row=0, column=1, sticky='nwe', padx=(10, 0), pady=(0, 15))

        tk.Label(right_top_frame, text="File Requirements",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        file_items = [
            "Only .pbip format files (.pbip folders) are supported",
            "Must contain semantic model definition folder with TMDL files",
            "Requires diagramLayout.json file for layout data",
            "Write permissions to PBIP folder (for saving changes)",
            "Legacy format with report.json files are NOT supported",
            ".pbix files are NOT supported",
            ".pbit files are NOT supported"
        ]

        for item in file_items:
            tk.Label(right_top_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=450,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        # ROW 1, COLUMN 0: Width Presets Explained
        left_bottom_frame = tk.Frame(sections_frame, bg=help_bg)
        left_bottom_frame.grid(row=1, column=0, sticky='nwe', padx=(0, 10))

        tk.Label(left_bottom_frame, text="Width Presets Explained",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        preset_items = [
            "Narrow: Compact columns (good for dense data tables)",
            "Medium: Balanced width for general use",
            "Wide: Spacious columns for readability",
            "Fit to Header: Auto-sizes to fit column header text",
            "Fit to Totals: Auto-sizes to fit total row (best for measures)",
            "Custom: Specify exact pixel width (minimum 50px)"
        ]

        for item in preset_items:
            tk.Label(left_bottom_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=450,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        # ROW 1, COLUMN 1: Important Notes
        right_bottom_frame = tk.Frame(sections_frame, bg=help_bg)
        right_bottom_frame.grid(row=1, column=1, sticky='nwe', padx=(10, 0))

        tk.Label(right_bottom_frame, text="Important Notes",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        notes_items = [
            "ONLY works with PBIP enhanced report format (PBIR)",
            "This tool is NOT officially supported by Microsoft",
            "Always backup your .pbip files before applying changes",
            "The tool modifies diagramLayout.json in your PBIP folder",
            "Test the optimized layout in Power BI Desktop before sharing",
            "Large models may take several minutes to analyze and optimize"
        ]

        for item in notes_items:
            tk.Label(right_bottom_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=450,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        help_window.bind('<Escape>', lambda event: help_window.destroy())

        # Center dialog and show (after all content built to prevent flash)
        help_window.update_idletasks()
        x = parent_window.winfo_rootx() + (parent_window.winfo_width() - 1000) // 2
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 760) // 2
        help_window.geometry(f"1000x760+{x}+{y}")

        # Set dark title bar BEFORE showing to prevent white flash
        help_window.update()
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled
