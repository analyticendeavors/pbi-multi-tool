"""
Layout Optimizer UI - Main Interface Tab
Built by Reid Havens of Analytic Endeavors

This tab provides layout optimization for Power BI relationship diagrams
with integrated advanced components when available.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
from typing import Dict, Any, Optional, List

from core.ui_base import BaseToolTab, SquareIconButton, ThemedScrollbar, ActionButtonBar, SplitLogSection, FileInputSection, ThemedMessageBox
from core.constants import AppConstants
from .enhanced_layout_core import EnhancedPBIPLayoutCore


class PBIPLayoutOptimizerTab(BaseToolTab):
    """Layout Optimizer main tab"""

    def __init__(self, parent, main_app=None):
        # Initialize with required parameters for BaseToolTab
        super().__init__(parent, main_app, "pbip_layout_optimizer", "Layout Optimizer")
        self.selected_pbip_folder = tk.StringVar()

        # Initialize logger
        import logging
        self.logger = logging.getLogger("pbip_layout_optimizer_ui")

        # Initialize enhanced core
        self.layout_core = EnhancedPBIPLayoutCore()

        # Canvas settings (simplified)
        self.canvas_width = tk.IntVar(value=1400)
        self.canvas_height = tk.IntVar(value=900)
        self.use_middle_out = tk.BooleanVar(value=True)  # Always enable middle-out
        self.preview_mode = tk.BooleanVar(value=False)   # Always save changes

        # Diagram selection state
        self.available_diagrams = []  # List of diagram info dicts
        self.diagram_scores = []  # Per-diagram quality scores from analysis
        self.selected_diagram_indices = set()  # Set of selected diagram indices
        self._diagram_items = {}  # Map tree item IDs to diagram indices
        self._viewed_diagram_index = 0  # Which diagram's analysis is shown in summary
        self._last_analysis_result = None  # Full analysis result for summary updates

        # Load icons for UI
        self._load_ui_icons()

        self._create_interface()

        # Force geometry calculation on initial load to fix alignment issues
        self.frame.update_idletasks()

    def _load_ui_icons(self):
        """Load SVG icons for buttons and section headers"""
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}

        # Section header icons (16px)
        section_icons = ['Power-BI', 'bar-chart', 'filter', 'analyze']
        for icon_name in section_icons:
            icon = self._load_icon_for_button(icon_name, size=16)
            if icon:
                self._button_icons[icon_name] = icon

        # Button icons (16px)
        button_icons = ['folder', 'magnifying-glass', 'execute', 'reset']
        for icon_name in button_icons:
            icon = self._load_icon_for_button(icon_name, size=16)
            if icon:
                self._button_icons[icon_name] = icon

        # Checkbox icons for diagram selection (16px) - themed for light/dark mode
        self._load_checkbox_icons()

    def _load_checkbox_icons(self):
        """Load themed checkbox SVG icons for checked, unchecked, and partial states."""
        is_dark = self._theme_manager.is_dark

        # Select icon names based on theme
        box_name = 'box-dark' if is_dark else 'box'
        checked_name = 'box-checked-dark' if is_dark else 'box-checked'
        partial_name = 'box-partial-dark' if is_dark else 'box-partial'

        # Load icons using base class method
        self._checkbox_off_icon = self._load_icon_for_button(box_name, size=16)
        self._checkbox_on_icon = self._load_icon_for_button(checked_name, size=16)
        self._checkbox_partial_icon = self._load_icon_for_button(partial_name, size=16)

    def _create_interface(self):
        """Create the main interface with responsive layout"""
        # Main container with scrolling - use BaseToolTab's frame
        main_frame = self.frame  # Use the frame from BaseToolTab

        # IMPORTANT: Pack buttons FIRST with side=BOTTOM
        # This ensures buttons are always visible even when window shrinks
        self._create_action_buttons_section(main_frame)

        # Create remaining sections from top (analysis will shrink with window)
        self._create_file_selection_section(main_frame)
        self._create_diagram_selection_section(main_frame)
        self._create_analysis_section(main_frame)

        # Force geometry calculation to fix initial alignment
        self.frame.update_idletasks()

    def _create_file_selection_section(self, parent):
        """Create PBIP file selection section using FileInputSection template widget"""
        # Load icons for the action button
        analyze_icon = self._button_icons.get('magnifying-glass')

        # Use FileInputSection template widget
        self.file_section = FileInputSection(
            parent=parent,
            theme_manager=self._theme_manager,
            section_title="PBIP File Source",
            section_icon="Power-BI",
            file_label="Project File (PBIP):",
            file_types=[("Power BI Project", "*.pbip")],
            action_button_text="ANALYZE LAYOUT",
            action_button_command=self._analyze_layout,
            action_button_icon=analyze_icon,
            help_command=self.show_help_dialog,
            on_file_selected=self._validate_folder
        )
        self.file_section.pack(fill=tk.X, pady=(0, 15))

        # Store references for compatibility with existing code
        self.selected_pbip_folder = self.file_section.path_var
        self.analyze_btn = self.file_section.action_button
        self.browse_btn = self.file_section.browse_button
        self._data_source_section = self.file_section.section_frame

        # Bind paste event to the entry
        self.file_section._path_entry.bind('<Control-v>', self._on_paste)

    def _create_diagram_selection_section(self, parent):
        """Create diagram selection section - modern scrollable list design"""
        colors = self._theme_manager.colors

        # Create icon labelwidget for section header
        diagram_header = self.create_section_header(self.frame, "Diagram Selection", "filter")[0]

        # Section frame with icon header - ALWAYS visible
        self._diagram_section = ttk.LabelFrame(parent, labelwidget=diagram_header,
                                               style='Section.TLabelframe', padding="12")
        self._diagram_section.pack(fill=tk.X, pady=(0, 15))

        # Content frame with padding="15" to match other sections - NO expand to keep compact
        content_frame = ttk.Frame(self._diagram_section, style='Section.TFrame', padding="15")
        content_frame.pack(fill=tk.X)  # fill=X only, no expand - keeps section compact
        self._diagram_content_frame = content_frame

        # Placeholder label (shown before analysis) - match frame background exactly
        frame_bg = colors['background']
        self._diagram_placeholder = tk.Label(
            content_frame,
            text="💡 Run analysis to detect available diagrams",
            font=('Segoe UI', 9, 'italic'),
            fg=colors['text_secondary'],
            bg=frame_bg,
            anchor='center'
        )
        self._diagram_placeholder.pack(fill=tk.X, pady=20)

        # Modern list container (hidden until analysis)
        # Container uses frame background to blend with section
        # FIXED HEIGHT container prevents vertical expansion (155px fits ~6 rows)
        # frame_bg already set above to colors['background']
        self._diagram_list_container = tk.Frame(content_frame, bg=frame_bg,
                                                 highlightthickness=0, height=155)
        self._diagram_list_container.pack_propagate(False)  # Prevent children from changing container size

        # Create canvas FIRST so scrollbar can reference it
        self._diagram_canvas = tk.Canvas(self._diagram_list_container, bg=colors['surface'],
                                          highlightthickness=0, height=100)

        # Themed scrollbar - ThemedScrollbar uses _command attribute directly (not configure)
        self._diagram_scrollbar = ThemedScrollbar(self._diagram_list_container,
                                                   command=self._diagram_canvas.yview,
                                                   theme_manager=self._theme_manager,
                                                   auto_hide=False)
        self._diagram_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Pack canvas AFTER scrollbar so it fills remaining horizontal space
        self._diagram_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._diagram_canvas.configure(yscrollcommand=self._diagram_scrollbar.set)

        # Inner frame for list items
        self._diagram_list_frame = tk.Frame(self._diagram_canvas, bg=colors['surface'])
        self._diagram_canvas_window = self._diagram_canvas.create_window((0, 0), window=self._diagram_list_frame,
                                                                          anchor='nw')

        # Bind canvas resize to update scroll region
        self._diagram_list_frame.bind('<Configure>', self._on_diagram_list_configure)
        self._diagram_canvas.bind('<Configure>', self._on_diagram_canvas_configure)

        # Enable mouse wheel scrolling
        self._diagram_canvas.bind('<Enter>', lambda e: self._diagram_canvas.bind_all('<MouseWheel>', self._on_diagram_mousewheel))
        self._diagram_canvas.bind('<Leave>', lambda e: self._diagram_canvas.unbind_all('<MouseWheel>'))

        # Store row widgets for theme updates
        self._diagram_row_widgets = []
        self._diagram_separator = None  # Will be created in _populate_diagram_list

        # Status row frame - holds status on left, tip on right (same line)
        # frame_bg already set above to colors['background']
        self._diagram_status_row = tk.Frame(content_frame, bg=frame_bg)

        # Selection status label (left side)
        self._diagram_status = tk.Label(
            self._diagram_status_row,
            text="Selected: 0 of 0 diagrams",
            font=('Segoe UI', 9),
            fg=colors['text_secondary'],
            bg=frame_bg
        )
        self._diagram_status.pack(side=tk.LEFT)

        # Tip label explaining click behavior (right side)
        self._diagram_tip = tk.Label(
            self._diagram_status_row,
            text="Tip: Click diagram name to view details in Analysis Summary below",
            font=('Segoe UI', 8, 'italic'),
            fg=colors['text_secondary'],
            bg=frame_bg
        )
        self._diagram_tip.pack(side=tk.RIGHT)
        # Status row will be packed after list is shown

    def _on_diagram_list_configure(self, event):
        """Update scroll region when list frame size changes"""
        self._diagram_canvas.configure(scrollregion=self._diagram_canvas.bbox('all'))

    def _on_diagram_canvas_configure(self, event):
        """Update inner frame width when canvas resizes"""
        self._diagram_canvas.itemconfig(self._diagram_canvas_window, width=event.width)

    def _on_diagram_mousewheel(self, event):
        """Handle mouse wheel scrolling with bounds checking"""
        # Only allow scrolling if content exceeds visible area
        bbox = self._diagram_canvas.bbox('all')
        if not bbox:
            return

        content_height = bbox[3] - bbox[1]
        canvas_height = self._diagram_canvas.winfo_height()

        # Don't scroll if content fits in canvas
        if content_height <= canvas_height:
            return

        # Check bounds before scrolling
        current = self._diagram_canvas.yview()
        if event.delta > 0 and current[0] <= 0:  # At top, trying to scroll up
            return
        if event.delta < 0 and current[1] >= 1:  # At bottom, trying to scroll down
            return

        self._diagram_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def _create_diagram_row(self, parent, idx: int, name: str, table_count: int,
                            quality_score: float = None, is_select_all: bool = False) -> dict:
        """Create a single diagram row with radio button, text, and optional quality score"""
        colors = self._theme_manager.colors
        bg_color = colors['surface']

        # Row frame
        row_frame = tk.Frame(parent, bg=bg_color)
        row_frame.pack(fill=tk.X, pady=(0, 4))  # 4px spacing between rows

        # Checkbox icon (clickable) - Select All row shows partial when some selected
        if is_select_all:
            if self._is_all_selected():
                icon = self._checkbox_on_icon
                is_selected = True
            elif self._is_some_selected():
                icon = self._checkbox_partial_icon
                is_selected = False
            else:
                icon = self._checkbox_off_icon
                is_selected = False
        else:
            is_selected = idx in self.selected_diagram_indices
            icon = self._checkbox_on_icon if is_selected else self._checkbox_off_icon

        icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
        if icon:
            icon_label.configure(image=icon)
            icon_label._icon_ref = icon
        else:
            icon_label.configure(text="☑️" if is_selected else "☐", font=('Segoe UI', 10))
        icon_label.pack(side=tk.LEFT, padx=(8, 10))  # 8px from left, 10px spacing to text

        # Build text with score if available - get rating text for full display
        if is_select_all:
            text = "Select All"
        elif quality_score is not None:
            # Get rating text based on score
            if quality_score >= 90:
                rating = "EXCELLENT"
            elif quality_score >= 80:
                rating = "GOOD"
            elif quality_score >= 60:
                rating = "OK"
            elif quality_score >= 30:
                rating = "BAD"
            else:
                rating = "VERY BAD"
            text = f"{name} ({table_count} tables) - Score: {quality_score:.1f}/100 ({rating})"
        else:
            text = f"{name} ({table_count} tables)"

        fg_color = colors['title_color'] if is_selected else colors['text_primary']
        font_weight = 'bold' if is_select_all else 'normal'

        text_label = tk.Label(row_frame, text=text, bg=bg_color, fg=fg_color,
                              font=('Segoe UI', 9, font_weight), cursor='hand2', anchor='w')
        text_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Bind clicks - ICON toggles selection, TEXT just shows analysis (doesn't toggle)
        if is_select_all:
            icon_label.bind('<Button-1>', lambda e: self._toggle_select_all())
            text_label.bind('<Button-1>', lambda e: self._toggle_select_all())
        else:
            icon_label.bind('<Button-1>', lambda e, i=idx: self._toggle_diagram_selection(i))
            text_label.bind('<Button-1>', lambda e, i=idx: self._show_diagram_analysis(i))  # Text only shows analysis

        # Hover underline effect (like Report Merger theme selection)
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

        return {
            'frame': row_frame,
            'icon_label': icon_label,
            'text_label': text_label,
            'idx': idx,
            'is_select_all': is_select_all
        }

    def _toggle_diagram_selection(self, idx: int):
        """Toggle selection state of a single diagram and show its analysis"""
        if idx in self.selected_diagram_indices:
            self.selected_diagram_indices.discard(idx)
        else:
            self.selected_diagram_indices.add(idx)

        self._update_all_diagram_rows()
        self._update_diagram_status()

        # Update Analysis Summary to show this diagram's stats
        self._show_diagram_analysis(idx)

    def _toggle_select_all(self):
        """Toggle all diagram selections"""
        if self._is_all_selected():
            # Deselect all
            self.selected_diagram_indices.clear()
        else:
            # Select all
            self.selected_diagram_indices = set(range(len(self.available_diagrams)))

        self._update_all_diagram_rows()
        self._update_diagram_status()

    def _is_all_selected(self) -> bool:
        """Check if all diagrams are selected"""
        return len(self.selected_diagram_indices) == len(self.available_diagrams) and len(self.available_diagrams) > 0

    def _is_some_selected(self) -> bool:
        """Check if some (but not all) diagrams are selected"""
        count = len(self.selected_diagram_indices)
        total = len(self.available_diagrams)
        return count > 0 and count < total

    def _update_all_diagram_rows(self):
        """Update all diagram row visuals"""
        colors = self._theme_manager.colors

        for row_data in self._diagram_row_widgets:
            idx = row_data['idx']
            is_select_all = row_data.get('is_select_all', False)

            # Determine icon based on selection state
            if is_select_all:
                # Select All row: checked if all, partial if some, unchecked if none
                if self._is_all_selected():
                    icon = self._checkbox_on_icon
                elif self._is_some_selected():
                    icon = self._checkbox_partial_icon
                else:
                    icon = self._checkbox_off_icon
                is_selected = self._is_all_selected()
            else:
                # Individual diagram rows: checked or unchecked
                is_selected = idx in self.selected_diagram_indices
                icon = self._checkbox_on_icon if is_selected else self._checkbox_off_icon

            # Update icon
            if icon:
                row_data['icon_label'].configure(image=icon)
                row_data['icon_label']._icon_ref = icon
            else:
                row_data['icon_label'].configure(text="☑️" if is_selected else "☐")

            # Update text color
            fg_color = colors['title_color'] if is_selected else colors['text_primary']
            row_data['text_label'].configure(fg=fg_color)

    def _update_diagram_status(self):
        """Update status label and optimize button state"""
        total_diagrams = len(self.available_diagrams)
        selected_count = len(self.selected_diagram_indices)
        self._diagram_status.configure(text=f"Selected: {selected_count} of {total_diagrams} diagrams")

        # Update optimize button state
        if selected_count > 0:
            self.optimize_btn.set_enabled(True)
        else:
            self.optimize_btn.set_enabled(False)

    def _filter_categorization_for_tables(self, categorization: Dict[str, Any], table_names: List[str]) -> Dict[str, Any]:
        """Filter categorization to only include tables in the given list"""
        if not categorization or not table_names:
            return categorization

        # Convert table_names to a set for O(1) lookup
        table_set = set(table_names)

        # Helper to filter a list and return count + filtered list
        def filter_tables(tables_list):
            if not tables_list:
                return []
            return [t for t in tables_list if t in table_set]

        filtered = {}

        # Fact tables
        fact_info = categorization.get('fact_tables', {})
        filtered_fact_tables = filter_tables(fact_info.get('tables', []))
        filtered['fact_tables'] = {
            'count': len(filtered_fact_tables),
            'tables': filtered_fact_tables
        }

        # Dimension tables
        dim_info = categorization.get('dimension_tables', {})
        filtered_l1 = filter_tables(dim_info.get('l1_tables', []))
        filtered_l2 = filter_tables(dim_info.get('l2_tables', []))
        filtered_l3 = filter_tables(dim_info.get('l3_tables', []))
        filtered_l4 = filter_tables(dim_info.get('l4_plus_tables', []))
        filtered['dimension_tables'] = {
            'l1_count': len(filtered_l1),
            'l2_count': len(filtered_l2),
            'l3_count': len(filtered_l3),
            'l4_plus_count': len(filtered_l4),
            'l1_tables': filtered_l1,
            'l2_tables': filtered_l2,
            'l3_tables': filtered_l3,
            'l4_plus_tables': filtered_l4
        }

        # Special tables
        special_info = categorization.get('special_tables', {})
        filtered_calendar = filter_tables(special_info.get('calendar_tables', []))
        filtered_metrics = filter_tables(special_info.get('metrics_tables', []))
        filtered_parameter = filter_tables(special_info.get('parameter_tables', []))
        filtered_calc_groups = filter_tables(special_info.get('calculation_groups', []))
        # Extension count from the full analysis (extensions don't have a separate tables list in special_info)
        filtered_extension = filter_tables(special_info.get('extension_tables', []))
        filtered['special_tables'] = {
            'calendar_count': len(filtered_calendar),
            'metrics_count': len(filtered_metrics),
            'parameter_count': len(filtered_parameter),
            'calculation_groups_count': len(filtered_calc_groups),
            'extension_count': len(filtered_extension),
            'calendar_tables': filtered_calendar,
            'metrics_tables': filtered_metrics,
            'parameter_tables': filtered_parameter,
            'calculation_groups': filtered_calc_groups,
            'extension_tables': filtered_extension
        }

        # Disconnected tables
        disconn_info = categorization.get('disconnected_tables', {})
        filtered_disconn = filter_tables(disconn_info.get('tables', []))
        filtered['disconnected_tables'] = {
            'count': len(filtered_disconn),
            'tables': filtered_disconn
        }

        # Excluded tables
        excl_info = categorization.get('excluded_tables', {})
        filtered_auto_date = filter_tables(excl_info.get('auto_date_tables', []))
        filtered['excluded_tables'] = {
            'auto_date_count': len(filtered_auto_date),
            'auto_date_tables': filtered_auto_date
        }

        return filtered

    def _show_diagram_analysis(self, idx: int):
        """Show analysis for a specific diagram in the Analysis Summary panel"""
        if not self.diagram_scores or idx < 0 or idx >= len(self.available_diagrams):
            return

        self._viewed_diagram_index = idx

        # Find the score info for this diagram
        score_info = next((s for s in self.diagram_scores if s['index'] == idx), None)
        if not score_info:
            return

        # Get full model categorization and filter to diagram's tables
        full_categorization = self._last_analysis_result.get('categorization', {}) if self._last_analysis_result else {}
        diagram_table_names = score_info.get('table_names', [])

        # Filter categorization to only include tables in this diagram
        filtered_categorization = self._filter_categorization_for_tables(full_categorization, diagram_table_names)

        # Build a result dict from diagram-specific data
        diagram_result = {
            'quality_score': score_info.get('quality_score', 0),
            'rating': score_info.get('rating', 'Unknown'),
            'layout_analysis': {
                'total_tables': score_info.get('table_count', 0),
                'positioned_tables': score_info.get('table_count', 0),
                'overlapping_tables': score_info.get('overlapping', 0),
                'average_spacing': score_info.get('avg_spacing', 0)
            },
            'diagram_name': score_info.get('name', f'Diagram {idx}'),
            'diagram_index': idx,
            # Use filtered categorization for this diagram's tables
            'categorization': filtered_categorization
        }

        # Update the summary panel with this diagram's data
        self._update_analysis_summary(diagram_result)

    def _populate_diagram_list(self):
        """Populate the modern diagram list with available diagrams"""
        colors = self._theme_manager.colors

        # Clear existing rows
        for widget in self._diagram_list_frame.winfo_children():
            widget.destroy()
        self._diagram_row_widgets.clear()
        self.selected_diagram_indices.clear()

        # Select all diagrams by default
        for idx in range(len(self.available_diagrams)):
            self.selected_diagram_indices.add(idx)

        # Add top spacer for padding from border (store reference for theme updates)
        self._diagram_top_spacer = tk.Frame(self._diagram_list_frame, height=8, bg=colors['surface'])
        self._diagram_top_spacer.pack(fill=tk.X)

        # Add "Select All" row first
        select_all_row = self._create_diagram_row(self._diagram_list_frame, -1, "Select All", 0, is_select_all=True)
        self._diagram_row_widgets.append(select_all_row)

        # Add separator line (store reference for theme updates)
        self._diagram_separator = tk.Frame(self._diagram_list_frame, height=1, bg=colors.get('border', '#e0e0e0'))
        self._diagram_separator.pack(fill=tk.X, padx=8, pady=(4, 8))

        # Add each diagram row with quality scores
        for idx, diagram in enumerate(self.available_diagrams):
            name = diagram.get('name', f"Diagram {diagram.get('ordinal', idx)}")
            table_count = diagram.get('table_count', 0)

            # Look up quality score from stored diagram_scores
            score = None
            if hasattr(self, 'diagram_scores') and self.diagram_scores:
                score_info = next((s for s in self.diagram_scores if s['index'] == idx), None)
                if score_info:
                    score = score_info.get('quality_score')

            row = self._create_diagram_row(self._diagram_list_frame, idx, name, table_count, quality_score=score)
            self._diagram_row_widgets.append(row)

        # Update status
        self._update_diagram_status()

    def _show_diagram_list(self):
        """Show the modern diagram list and hide placeholder"""
        self._diagram_placeholder.pack_forget()
        self._diagram_list_container.pack(fill=tk.X)  # Fixed height, don't expand
        self._diagram_status_row.pack(fill=tk.X, pady=(8, 0))  # Status and tip on same row

    def _hide_diagram_list(self):
        """Hide the diagram list and show placeholder"""
        self._diagram_list_container.pack_forget()
        self._diagram_status_row.pack_forget()
        self._diagram_placeholder.pack(fill=tk.X, pady=20)

    # Legacy aliases for compatibility
    def _show_diagram_tree(self):
        """Legacy alias - use _show_diagram_list"""
        self._show_diagram_list()

    def _hide_diagram_tree(self):
        """Legacy alias - use _hide_diagram_list"""
        self._hide_diagram_list()

    def _populate_diagram_tree(self):
        """Legacy alias - use _populate_diagram_list"""
        self._populate_diagram_list()

    def _create_analysis_section(self, parent):
        """Create split analysis section (40% summary / 60% progress log)"""
        colors = self._theme_manager.colors

        # Use SplitLogSection template widget
        self.log_section = SplitLogSection(
            parent=parent,
            theme_manager=self._theme_manager,
            section_title="Analysis & Progress",
            section_icon="analyze",
            summary_title="Layout Summary",
            summary_icon="bar-chart",
            log_title="Progress Log",
            log_icon="log-file",
            summary_placeholder="Analyze a model to see layout details"
        )
        self.log_section.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Store reference to text widget for logging (log_text is used by base class log_message)
        self.log_text = self.log_section.log_text
        self.analysis_text = self.log_text  # Alias for backwards compatibility

        # Create analysis summary content in the left panel
        self._create_analysis_summary_panel(self.log_section.summary_frame)

        # Initially show welcome message
        self._log_message("Welcome to the Layout Optimizer!")
        self._log_message("=" * 60)
        self._log_message("This tool optimizes Power BI semantic model diagram layouts:")
        self._log_message("- Apply Haven's middle-out design philosophy")
        self._log_message("- Auto-arrange tables for clean, professional layouts")
        self._log_message("- Categorize tables by type (Facts, Dimensions, Special)")
        self._log_message("")
        self._log_message("Start by selecting a .pbip file and clicking 'ANALYZE'")
        if not self.layout_core.mcp_available:
            self._log_message("Note: Using basic functionality - Enhanced components not available")

    def _create_analysis_summary_panel(self, parent):
        """Create analysis summary content panel with card-table structure (like Report Merger)"""
        colors = self._theme_manager.colors
        section_bg = colors.get('section_bg', colors['background'])

        # Clear existing widgets
        for widget in parent.winfo_children():
            widget.grid_forget()

        # Configure parent
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        # Main container frame for all content
        self._summary_container = tk.Frame(parent, bg=section_bg)
        self._summary_container.grid(row=0, column=0, sticky='nsew')

        # Placeholder (shown initially)
        self._summary_placeholder = tk.Label(
            self._summary_container,
            text="⏳ Waiting for analysis...",
            font=('Segoe UI', 9, 'italic'),
            fg=colors['text_secondary'],
            bg=section_bg,
            justify='center'
        )
        self._summary_placeholder.pack(expand=True)

        # Metrics frame (hidden until analysis)
        self._metrics_frame = tk.Frame(self._summary_container, bg=section_bg)

        # Store references for theme updates
        self._metric_labels = {}
        self._summary_table_frame = None
        self._summary_categorization_frame = None

    def _update_analysis_summary(self, result: Dict[str, Any]):
        """Update the analysis summary panel with card-table structure (like Report Merger)"""
        colors = self._theme_manager.colors
        section_bg = colors.get('section_bg', colors['background'])
        card_bg = colors.get('card_surface', colors['surface'])

        # Hide placeholder, show metrics
        self._summary_placeholder.pack_forget()
        self._metrics_frame.pack(fill=tk.BOTH, expand=True)

        # Clear existing content
        for widget in self._metrics_frame.winfo_children():
            widget.destroy()
        self._metric_labels.clear()

        # Get analysis data
        analysis = result.get('layout_analysis', {})
        quality_score = result.get('quality_score', 0)
        rating = result.get('rating', 'Unknown')
        diagram_name = result.get('diagram_name', 'All Tables')

        # === QUALITY SCORE CARD (main metrics) ===
        border_color = colors.get('border', '#333')
        self._summary_table_frame = tk.Frame(
            self._metrics_frame, bg=card_bg,
            highlightbackground=border_color, highlightcolor=border_color,
            highlightthickness=1, padx=12, pady=10
        )
        self._summary_table_frame.pack(fill=tk.X, pady=(0, 8))

        # Configure grid columns
        for col in range(2):
            self._summary_table_frame.columnconfigure(col, weight=1)

        # Diagram name header - shows which diagram's analysis is displayed
        diagram_header = tk.Label(
            self._summary_table_frame,
            text=f"Viewing: {diagram_name}",
            font=('Segoe UI Semibold', 9),
            fg=colors['title_color'], bg=card_bg
        )
        diagram_header.grid(row=0, column=0, sticky='w', pady=(0, 4))
        self._metric_labels['diagram_header'] = diagram_header

        # Hint text - on same row, right-aligned
        hint_label = tk.Label(
            self._summary_table_frame,
            text="Click a diagram above to view its analysis",
            font=('Segoe UI', 8, 'italic'),
            fg=colors['text_secondary'], bg=card_bg
        )
        hint_label.grid(row=0, column=1, sticky='e', pady=(0, 4))
        self._metric_labels['hint'] = hint_label

        # Quality Score row - color-coded by score
        score_color = colors.get('success', '#4CAF50') if quality_score >= 60 else (
            colors.get('warning', '#ff9800') if quality_score >= 40 else colors.get('error', '#f44336'))
        score_lbl = tk.Label(
            self._summary_table_frame,
            text=f"Quality Score: {quality_score}/100 ({rating})",
            font=('Segoe UI Semibold', 10),
            fg=score_color, bg=card_bg
        )
        score_lbl.grid(row=1, column=0, columnspan=2, sticky='w', pady=(0, 6))
        self._metric_labels['score'] = score_lbl

        # Metrics in 2-column grid layout (starting at row 2 after header and score)
        if analysis:
            metrics = [
                ("Tables:", analysis.get('total_tables', 0)),
                ("Positioned:", analysis.get('positioned_tables', 0)),
                ("Overlapping:", analysis.get('overlapping_tables', 0)),
                ("Avg Spacing:", f"{analysis.get('average_spacing', 0):.0f}px"),
            ]

            for i, (label_text, value) in enumerate(metrics):
                row = (i // 2) + 2  # Start at row 2 (after header at 0, score at 1)
                col = i % 2
                lbl = tk.Label(
                    self._summary_table_frame,
                    text=f"{label_text} {value}",
                    font=('Segoe UI', 9),
                    fg=colors['text_primary'], bg=card_bg, anchor='w'
                )
                lbl.grid(row=row, column=col, sticky='w', pady=2)
                self._metric_labels[label_text] = lbl

        # === SCORE EXPLANATION ===
        # Generate brief explanation of why score is what it is
        total_tables = analysis.get('total_tables', 0)
        overlapping = analysis.get('overlapping_tables', 0)
        spacing = analysis.get('average_spacing', 0)

        explanation_parts = []
        if total_tables <= 3:
            explanation_parts.append("Small diagram (1-3 tables) - simplified scoring")
            if overlapping == 0:
                explanation_parts.append("No overlaps (+20)")
        else:
            # Explain main scoring factors
            if overlapping > 0:
                explanation_parts.append(f"{overlapping} overlapping tables (penalty)")
            else:
                explanation_parts.append("No overlaps (good)")

            if spacing > 0:
                if spacing < 200:
                    explanation_parts.append("Tables too cramped")
                elif spacing > 2000:
                    explanation_parts.append("Tables too spread out")
                elif overlapping == 0:
                    # Only "Good spacing" if no overlaps - can't be good if tables overlap!
                    explanation_parts.append("Good spacing")

        if explanation_parts:
            explanation_text = " • ".join(explanation_parts)
            explain_lbl = tk.Label(
                self._summary_table_frame,
                text=explanation_text,
                font=('Segoe UI', 8, 'italic'),
                fg=colors['text_secondary'], bg=card_bg,
                wraplength=280, justify='left'
            )
            explain_lbl.grid(row=5, column=0, columnspan=2, sticky='w', pady=(6, 0))
            self._metric_labels['score_explanation'] = explain_lbl

        # === TABLE CATEGORIZATION CARD ===
        categorization = result.get('categorization', {})

        if categorization:
            cat_border_color = colors.get('border', '#333')
            self._summary_categorization_frame = tk.Frame(
                self._metrics_frame, bg=card_bg,
                highlightbackground=cat_border_color, highlightcolor=cat_border_color,
                highlightthickness=1, padx=12, pady=10
            )
            self._summary_categorization_frame.pack(fill=tk.X, pady=(0, 8))

            # Configure grid columns
            for col in range(2):
                self._summary_categorization_frame.columnconfigure(col, weight=1)

            # Header
            cat_header = tk.Label(
                self._summary_categorization_frame,
                text="TABLE CATEGORIZATION",
                font=('Segoe UI Semibold', 9),
                fg=colors['title_color'], bg=card_bg
            )
            cat_header.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 6))
            self._metric_labels['cat_header'] = cat_header

            # Build category items
            cat_items = []

            # Fact tables
            fact_info = categorization.get('fact_tables', {})
            if fact_info.get('count', 0) > 0:
                cat_items.append(f"FACT TABLES: {fact_info['count']}")

            # Dimension tables by level
            dim_info = categorization.get('dimension_tables', {})
            for level, count_key in [('L1', 'l1_count'), ('L2', 'l2_count'), ('L3', 'l3_count')]:
                count = dim_info.get(count_key, 0)
                if count > 0:
                    cat_items.append(f"{level} DIMENSIONS: {count}")

            # Special tables
            special_info = categorization.get('special_tables', {})
            special_types = [
                ('CALENDAR TABLES:', 'calendar_count'),
                ('PARAMETER TABLES:', 'parameter_count'),
                ('TABLE EXTENSIONS:', 'extension_count'),
            ]
            for label, count_key in special_types:
                count = special_info.get(count_key, 0)
                if count > 0:
                    cat_items.append(f"{label} {count}")

            # Display categories in 2-column grid
            for i, item_text in enumerate(cat_items):
                row = (i // 2) + 1
                col = i % 2
                lbl = tk.Label(
                    self._summary_categorization_frame,
                    text=item_text,
                    font=('Segoe UI', 9),
                    fg=colors['text_primary'], bg=card_bg, anchor='w'
                )
                lbl.grid(row=row, column=col, sticky='w', pady=2)
                self._metric_labels[f'cat_{i}'] = lbl

    def _reset_analysis_summary(self):
        """Reset the analysis summary panel to initial state"""
        # Clear metrics frame
        for widget in self._metrics_frame.winfo_children():
            widget.destroy()
        self._metric_labels.clear()

        # Hide metrics, show placeholder
        self._metrics_frame.pack_forget()
        self._summary_placeholder.pack(expand=True)
    
    def _create_action_buttons_section(self, parent):
        """Create action buttons section - packed at BOTTOM for responsive layout"""
        execute_icon = self._button_icons.get('execute')
        reset_icon = self._button_icons.get('reset')

        # Use centralized ActionButtonBar
        self.button_frame = ActionButtonBar(
            parent=parent,
            theme_manager=self._theme_manager,
            primary_text="OPTIMIZE LAYOUT",
            primary_command=self._optimize_layout,
            primary_icon=execute_icon,
            secondary_text="RESET ALL",
            secondary_command=self._reset_interface,
            secondary_icon=reset_icon,
            primary_starts_disabled=True
        )
        self.button_frame.pack(side=tk.BOTTOM, pady=(15, 0))

        # Expose buttons for compatibility with existing code
        self.optimize_btn = self.button_frame.primary_button
        self.reset_btn = self.button_frame.secondary_button

        # Track in button lists for any additional theme handling
        self._primary_buttons.append(self.optimize_btn)
        self._secondary_buttons.append(self.reset_btn)
    
    # _browse_file is now handled by FileInputSection template widget

    def _extract_folder_from_pbip_file(self, file_path: str) -> Optional[str]:
        """Extract the folder path from a .pbip file"""
        try:
            file_path_obj = Path(file_path)
            ext = file_path_obj.suffix.lower()

            # Check if file exists
            if not file_path_obj.exists():
                return None

            # Only handle PBIP files
            if ext != '.pbip':
                return None

            # The folder is the parent directory of the .pbip file
            folder_path = file_path_obj.parent

            # Verify it's a valid PBIP folder structure
            semantic_folders = list(folder_path.glob("*.SemanticModel"))
            if not semantic_folders:
                return None

            return str(folder_path)

        except Exception as e:
            self.logger.error(f"Error validating PBIP file: {e}")
            return None

    # Alias for backwards compatibility
    def _extract_folder_from_pbi_file(self, file_path: str) -> Optional[str]:
        """Alias for _extract_folder_from_pbip_file"""
        return self._extract_folder_from_pbip_file(file_path)
    
    def _on_paste(self, event):
        """Handle paste event to clean path quotes"""
        # Schedule validation after paste
        self.master.after(100, lambda: self._validate_folder(self.selected_pbip_folder.get()))
    
    def _validate_folder(self, file_path: str):
        """Validate selected PBIP file"""
        # Clean quotes from path
        file_path = file_path.strip().strip('"').strip("'")
        if file_path != self.selected_pbip_folder.get():
            self.selected_pbip_folder.set(file_path)

        if not file_path:
            return

        # Extract/validate path from PBIP file
        result_path = self._extract_folder_from_pbip_file(file_path)
        if not result_path:
            self.analyze_btn.set_enabled(False)
            self.optimize_btn.set_enabled(False)
            return

        # Validate PBIP folder structure
        validation = self.layout_core.validate_pbip_folder(result_path)
        if validation['valid']:
            self.analyze_btn.set_enabled(True)
        else:
            self.analyze_btn.set_enabled(False)
            self.optimize_btn.set_enabled(False)
    
    def _analyze_layout(self):
        """Analyze layout quality"""
        file_path = self.selected_pbip_folder.get()
        if not file_path:
            return  # Button should be disabled, but safety check

        # Extract folder path from file path
        folder_path = self._extract_folder_from_pbip_file(file_path)
        if not folder_path:
            return  # Button should be disabled, but safety check
        
        self.analyze_btn.set_enabled(False)

        # Run analysis in background
        self.run_in_background(
            target_func=lambda: self._analysis_thread_target(folder_path),
            success_callback=self._handle_analysis_result,
            error_callback=lambda e: self._handle_analysis_error(str(e))
        )
    
    def _analysis_thread_target(self, folder_path: str):
        """Background analysis logic for PBIP folders"""
        self.update_progress(10, "Validating PBIP folder...")
        self.update_progress(30, "Reading table definitions...")
        self.update_progress(50, "Analyzing layout quality...")
        layout_result = self.layout_core.analyze_layout_quality(folder_path)

        # If advanced components are available, also get table categorization
        if self.layout_core.mcp_available:
            self.update_progress(70, "Categorizing tables...")
            categorization_result = self.layout_core.analyze_table_categorization(folder_path)
            if categorization_result.get('success'):
                layout_result['categorization'] = categorization_result.get('categorization', {})
                layout_result['mcp_extensions'] = categorization_result.get('extensions', [])

        self.update_progress(90, "Calculating quality score...")
        self.update_progress(100, "Analysis complete!")
        return layout_result
    
    def _handle_analysis_result(self, result: Dict[str, Any]):
        """Handle analysis result"""
        if result['success']:
            # Merge extension count into categorization for display
            categorization = result.get('categorization', {})
            extensions = result.get('mcp_extensions', [])
            if categorization and extensions:
                special_tables = categorization.get('special_tables', {})
                special_tables['extension_count'] = len(extensions)
                # Store extension table names for filtering
                special_tables['extension_tables'] = [ext.get('extension_table', '') for ext in extensions]

            # Store for diagram-specific views
            self._last_analysis_result = result

            # Log analysis results to the single log area
            self._log_analysis_results(result)

            # Add default diagram name (All Tables is always diagram 0)
            result['diagram_name'] = 'All Tables'
            self._viewed_diagram_index = 0

            # Update analysis summary panel
            self._update_analysis_summary(result)

            # Get diagram list and populate selection UI
            file_path = self.selected_pbip_folder.get()
            folder_path = self._extract_folder_from_pbip_file(file_path)
            if folder_path:
                self.available_diagrams = self.layout_core.get_diagram_list(folder_path)
                self.diagram_scores = self.layout_core.analyze_diagram_quality(folder_path)
                diagram_scores = self.diagram_scores

                if self.available_diagrams:
                    # Populate the diagram tree with checkboxes
                    self._populate_diagram_tree()

                    # Show the diagram treeview (hide placeholder)
                    self._show_diagram_tree()

                    # Log diagram info with quality scores
                    self._log_message("")
                    self._log_message(f"📊 Found {len(self.available_diagrams)} diagram(s) in layout file:")
                    for i, diagram in enumerate(self.available_diagrams):
                        name = diagram.get('name', f"Diagram {diagram.get('ordinal', 0)}")
                        count = diagram.get('table_count', 0)

                        # Find matching quality score
                        score_info = next((s for s in diagram_scores if s['index'] == i), None)
                        if score_info:
                            score = score_info['quality_score']
                            rating = score_info['rating']
                            self._log_message(f"   • {name} ({count} tables) - Score: {score}/100 ({rating})")
                        else:
                            self._log_message(f"   • {name} ({count} tables)")
                    self._log_message("")

            # Enable optimization button (will be controlled by diagram selection)
            if self.selected_diagram_indices:
                self.optimize_btn.set_enabled(True)

        else:
            self._log_message(f"❌ Analysis failed: {result.get('error', 'Unknown error')}")

        self.analyze_btn.set_enabled(True)
    
    def _handle_analysis_error(self, error: str):
        """Handle analysis error"""
        self.analyze_btn.set_enabled(True)
        self._log_message(f"❌ Analysis Error: {error}")
    
    def _optimize_layout(self):
        """Optimize layout"""
        file_path = self.selected_pbip_folder.get()
        if not file_path:
            return  # Button should be disabled, but safety check

        # Extract folder path from file path
        folder_path = self._extract_folder_from_pbip_file(file_path)
        if not folder_path:
            return  # Button should be disabled, but safety check

        # Check diagram selection
        if not self.selected_diagram_indices:
            return  # Button should be disabled, but safety check

        self.optimize_btn.set_enabled(False)

        # Get selected indices as sorted list
        selected_indices = sorted(list(self.selected_diagram_indices))

        # Log which diagrams will be optimized
        self._log_message("")
        self._log_message(f"🎯 Optimizing {len(selected_indices)} diagram(s)...")
        for idx in selected_indices:
            if idx < len(self.available_diagrams):
                diagram = self.available_diagrams[idx]
                self._log_message(f"   • {diagram.get('name', f'Diagram {idx}')}")

        # Run optimization in background
        self.run_in_background(
            target_func=lambda: self._optimization_thread_target(folder_path, selected_indices),
            success_callback=self._handle_optimization_result,
            error_callback=lambda e: self._handle_optimization_error(str(e))
        )
    
    def _optimization_thread_target(self, folder_path: str, diagram_indices: List[int] = None):
        """Background optimization logic for PBIP folders"""
        self.update_progress(10, "Preparing layout optimization...")

        # Default to first diagram if no selection provided
        if diagram_indices is None:
            diagram_indices = [0]

        self.update_progress(30, "Analyzing table relationships...")
        self.update_progress(60, "Generating optimal positions...")

        # Always save changes (no preview mode in simplified version)
        save_changes = True

        # Always use middle-out design (no checkbox in simplified version)
        use_middle_out = True

        # Use enhanced optimization with diagram indices
        result = self.layout_core.optimize_layout(
            folder_path,
            self.canvas_width.get(),
            self.canvas_height.get(),
            save_changes,
            use_middle_out,
            diagram_indices=diagram_indices
        )

        self.update_progress(100, "Optimization complete!")
        return result
    
    def _handle_optimization_result(self, result: Dict[str, Any]):
        """Handle optimization result"""
        if result['success']:
            # Log optimization results to the single log area
            self._log_optimization_results(result)

            # Show success dialog with summary
            diagrams_count = len(result.get('diagrams_optimized', []))
            tables_count = result.get('tables_arranged', 0)
            self._show_optimization_complete_dialog(diagrams_count, tables_count)

        else:
            self._log_message(f"❌ Optimization failed: {result.get('error', 'Unknown error')}")
            self.show_toast("Optimization failed", toast_type='error', duration=4000)

        self.optimize_btn.set_enabled(True)

    def _handle_optimization_error(self, error: str):
        """Handle optimization error"""
        self.optimize_btn.set_enabled(True)
        self._log_message(f"❌ Optimization Error: {error}")

    def _show_optimization_complete_dialog(self, diagrams_count: int, tables_count: int):
        """Show a styled dialog popup when optimization completes"""
        # Get parent window
        parent_window = None
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'root'):
            parent_window = self.main_app.root
        elif hasattr(self, 'master'):
            parent_window = self.master
        elif hasattr(self, 'parent'):
            parent_window = self.parent
        else:
            parent_window = self.frame.winfo_toplevel()

        # Create dialog window
        dialog = tk.Toplevel(parent_window)
        dialog.withdraw()  # Hide until styled
        dialog.title("Optimization Complete")
        dialog.resizable(False, False)
        dialog.transient(parent_window)
        dialog.grab_set()

        # Set icon
        try:
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'config'):
                dialog.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Get theme colors
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        bg_color = colors.get('background', '#ffffff')

        dialog.configure(bg=bg_color)

        # Main content frame
        content = tk.Frame(dialog, bg=bg_color, padx=30, pady=25)
        content.pack(fill=tk.BOTH, expand=True)

        # Success icon and title row
        title_frame = tk.Frame(content, bg=bg_color)
        title_frame.pack(fill=tk.X, pady=(0, 15))

        # Checkmark icon
        tk.Label(
            title_frame,
            text="✓",
            font=('Segoe UI', 28, 'bold'),
            fg=colors.get('success', '#10b981'),
            bg=bg_color
        ).pack(side=tk.LEFT, padx=(0, 12))

        # Title text
        tk.Label(
            title_frame,
            text="Layout Optimization Complete",
            font=('Segoe UI Semibold', 14),
            fg=colors.get('text_primary', '#333333'),
            bg=bg_color
        ).pack(side=tk.LEFT, anchor='w')

        # Summary message
        summary_text = f"Successfully optimized {diagrams_count} diagram{'s' if diagrams_count != 1 else ''} with {tables_count} table{'s' if tables_count != 1 else ''}"
        tk.Label(
            content,
            text=summary_text,
            font=('Segoe UI', 11),
            fg=colors.get('text_secondary', '#666666'),
            bg=bg_color,
            wraplength=350
        ).pack(pady=(0, 20))

        # OK button
        ok_btn = self.create_action_button(content, "OK", lambda: dialog.destroy(), width=20)
        ok_btn.pack(pady=(0, 5))

        # Position dialog centered on parent
        dialog.update_idletasks()
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        x = parent_window.winfo_rootx() + (parent_window.winfo_width() - width) // 2
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - height) // 2
        dialog.geometry(f"+{x}+{y}")

        # Set dark title bar if needed
        dialog.update()
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(dialog, is_dark)
        dialog.deiconify()  # Show now that it's styled

        # Focus OK button for keyboard Enter
        ok_btn.focus_set()
        dialog.bind('<Return>', lambda e: dialog.destroy())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

        # Wait for dialog to close
        dialog.wait_window()

    def _log_analysis_results(self, result: Dict[str, Any]):
        """Log analysis results to the analysis text area"""
        self._log_message("📊 LAYOUT ANALYSIS COMPLETE:")
        self._log_message("=" * 50)
        
        # Basic analysis info
        analysis = result.get('layout_analysis', {})
        if analysis:
            self._log_message(f"📈 Quality Score: {result.get('quality_score', 0)}/100 ({result.get('rating', 'Unknown')})")
            self._log_message(f"📋 Total Tables: {analysis.get('total_tables', 0)}")
            self._log_message(f"📍 Positioned Tables: {analysis.get('positioned_tables', 0)}")
            self._log_message(f"⚠️ Overlapping Tables: {analysis.get('overlapping_tables', 0)}")
            self._log_message(f"🔄 Average Spacing: {analysis.get('average_spacing', 0):.1f}px")
        
        # Advanced categorization if available
        categorization = result.get('categorization', {})
        if categorization:
            self._log_message("")
            self._log_message("🏷️ TABLE CATEGORIZATION (Advanced):")
            self._log_message("=" * 40)
            
            # Facts - single line
            fact_info = categorization.get('fact_tables', {})
            if fact_info.get('count', 0) > 0:
                self._log_message(f"📊 FACT TABLES: {fact_info['count']}")
            
            # Dimensions - single line each
            dim_info = categorization.get('dimension_tables', {})
            for level, count_key in [('L1', 'l1_count'), ('L2', 'l2_count'), ('L3', 'l3_count'), ('L4+', 'l4_plus_count')]:
                count = dim_info.get(count_key, 0)
                if count > 0:
                    self._log_message(f"📁 {level} DIMENSIONS: {count}")
            
            # Special tables - single line each
            special_info = categorization.get('special_tables', {})
            if special_info.get('calendar_count', 0) > 0:
                self._log_message(f"📅 CALENDAR TABLES: {special_info['calendar_count']}")
            if special_info.get('metrics_count', 0) > 0:
                self._log_message(f"📊 METRICS TABLES: {special_info['metrics_count']}")
            if special_info.get('parameter_count', 0) > 0:
                self._log_message(f"⚙️ PARAMETER TABLES: {special_info['parameter_count']}")
            
            # Extensions - single line
            extensions = result.get('mcp_extensions', [])
            if extensions:
                self._log_message(f"🔗 TABLE EXTENSIONS: {len(extensions)}")
        
        self._log_message("")
        self._log_message("✅ Analysis complete! Ready for optimization.")
    
    def _log_optimization_results(self, result: Dict[str, Any]):
        """Log optimization results to the single log area"""
        self._log_message("")
        self._log_message("🎯 LAYOUT OPTIMIZATION COMPLETE:")
        self._log_message("=" * 50)
        
        # Basic optimization info
        method = result.get('layout_method', 'Enhanced')
        tables_arranged = result.get('tables_arranged', 0)
        changes_saved = result.get('changes_saved', False)
        
        self._log_message(f"🎯 Layout Method: {method}")
        self._log_message(f"📊 Tables Arranged: {tables_arranged}")
        self._log_message(f"💾 Changes Saved: {'Yes' if changes_saved else 'No'}")
        self._log_message(f"📌 Canvas Size: {self.canvas_width.get()}x{self.canvas_height.get()}")
        self._log_message(f"📆 Layout Design: Middle-Out Philosophy")
        
        # Layout features if available
        if result.get('layout_features'):
            self._log_message("")
            self._log_message("🔧 Layout Features:")
            features = result['layout_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                self._log_message(f"  {status} {feature.replace('_', ' ').title()}")
        
        # Advanced features if available
        if result.get('advanced_features'):
            self._log_message("")
            self._log_message("⚡ Advanced Features:")
            features = result['advanced_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                self._log_message(f"  {status} {feature.replace('_', ' ').title()}")
        
        self._log_message("")
        self._log_message("✅ Layout optimization completed successfully!")
        if changes_saved:
            self._log_message("💾 Your diagram layout has been updated.")
        else:
            self._log_message("📝 Preview mode - no changes were saved.")
    
    def _log_message(self, message: str):
        """Add message to analysis log - use base class method"""
        self.log_message(message)
    
    def _update_optimization_display(self, result: Dict[str, Any]):
        """Update optimization results display"""
        # Clear placeholder
        for widget in self.optimization_results_frame.winfo_children():
            widget.destroy()
        
        # Create results summary
        summary_frame = ttk.Frame(self.optimization_results_frame, style='Section.TFrame')
        summary_frame.pack(fill=tk.X)

        ttk.Label(summary_frame, text="Layout Optimization Complete",
                 style='Section.TLabel',
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['success']).pack(anchor=tk.W)

        # Results grid
        results_grid = ttk.Frame(summary_frame, style='Section.TFrame')
        results_grid.pack(fill=tk.X, pady=(10, 0))

        # Left column
        left_frame = ttk.Frame(results_grid, style='Section.TFrame')
        left_frame.pack(side=tk.LEFT, anchor=tk.NW)

        # Calculate tables arranged
        tables_arranged = result.get('tables_arranged', 0)

        ttk.Label(left_frame, text=f"Tables Arranged: {tables_arranged}",
                 style='Section.TLabel', font=('Segoe UI', 9)).pack(anchor=tk.W)
        ttk.Label(left_frame, text=f"Changes Saved: {'Yes' if result.get('changes_saved') else 'No'}",
                 style='Section.TLabel', font=('Segoe UI', 9)).pack(anchor=tk.W)
        ttk.Label(left_frame, text=f"Canvas Size: {self.canvas_width.get()}x{self.canvas_height.get()}",
                 style='Section.TLabel', font=('Segoe UI', 9)).pack(anchor=tk.W)

        # Right column
        right_frame = ttk.Frame(results_grid, style='Section.TFrame')
        right_frame.pack(side=tk.LEFT, anchor=tk.NW, padx=(50, 0))

        layout_method = result.get('layout_method', 'Enhanced')
        ttk.Label(right_frame, text=f"Layout Method: {layout_method}",
                 style='Section.TLabel', font=('Segoe UI', 9)).pack(anchor=tk.W)

        if result.get('advanced_features'):
            ttk.Label(right_frame, text="Advanced Features: Enabled",
                     style='Section.TLabel', font=('Segoe UI', 9)).pack(anchor=tk.W)
        else:
            ttk.Label(right_frame, text="Mode: Basic Layout",
                     style='Section.TLabel', font=('Segoe UI', 9)).pack(anchor=tk.W)
    
    def _format_analysis_results(self, result: Dict[str, Any]) -> str:
        """Format analysis results for detailed display"""
        text = "LAYOUT ANALYSIS RESULTS\n"
        text += "=" * 50 + "\n\n"
        
        text += f"PBIP Folder: {result.get('pbip_folder', '')}\n"
        text += f"Semantic Model: {result.get('semantic_model_path', '')}\n\n"
        
        # Quality assessment
        text += f"QUALITY ASSESSMENT:\n"
        text += f"Overall Score: {result.get('quality_score', 0)}/100 ({result.get('rating', 'Unknown')})\n\n"
        
        # Table statistics
        analysis = result.get('layout_analysis', {})
        if analysis:
            text += f"TABLE STATISTICS:\n"
            text += f"Total Tables: {analysis.get('total_tables', 0)}\n"
            text += f"Positioned Tables: {analysis.get('positioned_tables', 0)}\n"
            text += f"Overlapping Tables: {analysis.get('overlapping_tables', 0)}\n"
            text += f"Average Spacing: {analysis.get('average_spacing', 0):.1f}px\n\n"
        
        # Add advanced categorization if available
        if result.get('categorization') and self.layout_core.mcp_available:
            text += self._format_categorization_in_analysis(result['categorization'], result.get('mcp_extensions', []))
        
        if result.get('recommendations'):
            text += f"RECOMMENDATIONS:\n"
            for i, rec in enumerate(result['recommendations'], 1):
                text += f"{i}. {rec}\n"
            text += "\n"
        
        if result.get('table_names'):
            text += f"SAMPLE TABLES (first 10):\n"
            for table in result['table_names']:
                text += f"• {table}\n"
        
        return text
    
    def _format_categorization_in_analysis(self, categorization: Dict[str, Any], extensions: List[Dict[str, Any]]) -> str:
        """Format categorization data for inclusion in analysis results"""
        text = f"🏷️ TABLE CATEGORIZATION (Advanced Enhanced):\n"
        text += "=" * 40 + "\n"
        
        # Facts
        fact_info = categorization.get('fact_tables', {})
        fact_count = fact_info.get('count', 0)
        if fact_count > 0:
            text += f"📊 FACT TABLES ({fact_count}):\n"
            for table in fact_info.get('tables', [])[:5]:  # Show first 5
                text += f"   • {table}\n"
            if len(fact_info.get('tables', [])) > 5:
                text += f"   ... and {len(fact_info.get('tables', [])) - 5} more\n"
            text += "\n"
        
        # Dimensions by level
        dim_info = categorization.get('dimension_tables', {})
        for level, count_key, tables_key in [
            ('L1', 'l1_count', 'l1_tables'),
            ('L2', 'l2_count', 'l2_tables'), 
            ('L3', 'l3_count', 'l3_tables'),
            ('L4+', 'l4_plus_count', 'l4_plus_tables')
        ]:
            count = dim_info.get(count_key, 0)
            if count > 0:
                text += f"📁 {level} DIMENSION TABLES ({count}):\n"
                tables = dim_info.get(tables_key, [])
                for table in tables[:5]:  # Show first 5
                    text += f"   • {table}\n"
                if len(tables) > 5:
                    text += f"   ... and {len(tables) - 5} more\n"
                text += "\n"
        
        # Special tables
        special_info = categorization.get('special_tables', {})
        special_types = [
            ('📅 CALENDAR', 'calendar_count', 'calendar_tables'),
            ('📊 METRICS', 'metrics_count', 'metrics_tables'),
            ('⚙️ PARAMETERS', 'parameter_count', 'parameter_tables'),
            ('🧮 CALCULATION GROUPS', 'calculation_groups_count', 'calculation_groups')
        ]
        
        for label, count_key, tables_key in special_types:
            count = special_info.get(count_key, 0)
            if count > 0:
                text += f"{label} ({count}):\n"
                tables = special_info.get(tables_key, [])
                for table in tables[:3]:  # Show first 3
                    text += f"   • {table}\n"
                if len(tables) > 3:
                    text += f"   ... and {len(tables) - 3} more\n"
                text += "\n"
        
        # Extensions
        if extensions:
            text += f"🔗 TABLE EXTENSIONS ({len(extensions)}):\n"
            for ext in extensions[:3]:  # Show first 3
                text += f"   • {ext['extension_table']} → {ext['base_table']}\n"
            if len(extensions) > 3:
                text += f"   ... and {len(extensions) - 3} more\n"
            text += "\n"
        
        # Disconnected and excluded
        disconnected_count = categorization.get('disconnected_tables', {}).get('count', 0)
        excluded_count = categorization.get('excluded_tables', {}).get('auto_date_count', 0)
        
        if disconnected_count > 0:
            text += f"🔌 DISCONNECTED TABLES: {disconnected_count}\n"
        if excluded_count > 0:
            text += f"🗓️ AUTO DATE TABLES (excluded): {excluded_count}\n"
        
        text += "\n"
        return text
    
    def _format_optimization_results(self, result: Dict[str, Any]) -> str:
        """Format optimization results for detailed display"""
        text = "LAYOUT OPTIMIZATION RESULTS\n"
        text += "=" * 50 + "\n\n"
        
        text += f"Operation: {result.get('operation', 'Layout Optimization')}\n"
        text += f"PBIP Folder: {result.get('pbip_folder', '')}\n"
        text += f"Layout Method: {result.get('layout_method', 'Enhanced')}\n\n"
        
        text += f"OPTIMIZATION RESULTS:\n"
        text += f"Tables Arranged: {result.get('tables_arranged', 0)}\n"
        text += f"Changes Saved: {'Yes' if result.get('changes_saved') else 'No'}\n"
        text += f"Canvas Size: {self.canvas_width.get()}x{self.canvas_height.get()}\n\n"
        
        # Layout features
        if result.get('layout_features'):
            text += f"LAYOUT FEATURES:\n"
            features = result['layout_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                text += f"  {status} {feature.replace('_', ' ').title()}\n"
            text += "\n"
        
        # Advanced features
        if result.get('advanced_features'):
            text += f"ADVANCED FEATURES:\n"
            features = result['advanced_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                text += f"  {status} {feature.replace('_', ' ').title()}\n"
            text += "\n"
        
        # Enhancement status
        text += f"COMPONENT STATUS:\n"
        text += f"  ✅ Advanced Components: {'Available' if self.layout_core.mcp_available else 'Not Available'}\n"
        text += f"  ✅ Middle-Out Design: {'Enabled' if self.use_middle_out.get() else 'Disabled'}\n"
        text += f"  ✅ Table Categorization: {'Enabled' if self.layout_core.mcp_available else 'Basic'}\n"
        
        return text
    
    def _update_results_text(self, text: str):
        """Update the detailed results text area"""
        self.results_text.configure(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, text)
        self.results_text.configure(state=tk.DISABLED)
    
    def _export_log(self):
        """Export analysis log to file"""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Export Analysis Log",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if file_path:
                log_content = self.analysis_text.get(1.0, tk.END)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                
                ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Export Complete", f"Log exported to {file_path}")

        except Exception as e:
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Export Error", f"Failed to export log: {str(e)}")
    
    def _clear_log(self):
        """Clear the analysis log"""
        self.analysis_text.configure(state=tk.NORMAL)
        self.analysis_text.delete(1.0, tk.END)
        self.analysis_text.configure(state=tk.DISABLED)
        
        self._log_message("📝 Layout Optimizer - Log Cleared")
        self._log_message("🚀 Ready for new analysis...")
        self._log_message("")
    
    def _reset_interface(self):
        """Reset the entire interface"""
        # Reset folder selection
        self.selected_pbip_folder.set("")

        # Reset diagram selection (modern list)
        self.available_diagrams = []
        self.selected_diagram_indices = set()
        self._diagram_row_widgets.clear()
        self._diagram_separator = None  # Will be recreated in _populate_diagram_list
        for widget in self._diagram_list_frame.winfo_children():
            widget.destroy()
        self._hide_diagram_list()

        # Reset analysis summary
        self._reset_analysis_summary()

        # Reset buttons
        self.analyze_btn.set_enabled(False)
        self.optimize_btn.set_enabled(False)
    
    def _handle_theme_change(self, theme: str):
        """Override to update all custom components on theme change"""
        super()._handle_theme_change(theme)
        colors = self._theme_manager.colors
        # Use theme colors for proper updates
        bg_color = colors['section_bg']
        header_bg_color = colors['background']

        # SplitLogSection handles its own theme updates internally

        # Update diagram placeholder (match frame background exactly)
        if hasattr(self, '_diagram_placeholder') and self._diagram_placeholder:
            try:
                self._diagram_placeholder.configure(
                    fg=colors['text_secondary'],
                    bg=header_bg_color  # Use main background color
                )
            except Exception:
                pass

        # Update modern diagram list components
        surface_color = colors['card_surface']
        # Container uses header background to blend with section frame
        if hasattr(self, '_diagram_list_container') and self._diagram_list_container:
            try:
                self._diagram_list_container.configure(bg=header_bg_color)
            except Exception:
                pass
        if hasattr(self, '_diagram_canvas') and self._diagram_canvas:
            try:
                self._diagram_canvas.configure(bg=surface_color)
            except Exception:
                pass
        if hasattr(self, '_diagram_list_frame') and self._diagram_list_frame:
            try:
                self._diagram_list_frame.configure(bg=surface_color)
            except Exception:
                pass

        # Update diagram top spacer
        if hasattr(self, '_diagram_top_spacer') and self._diagram_top_spacer:
            try:
                self._diagram_top_spacer.configure(bg=surface_color)
            except Exception:
                pass

        # Update diagram row widgets
        if hasattr(self, '_diagram_row_widgets'):
            for row_data in self._diagram_row_widgets:
                try:
                    idx = row_data.get('idx', -1)
                    is_select_all = row_data.get('is_select_all', False)
                    is_selected = self._is_all_selected() if is_select_all else idx in self.selected_diagram_indices
                    fg_color = colors['title_color'] if is_selected else colors['text_primary']

                    if 'frame' in row_data:
                        row_data['frame'].configure(bg=surface_color)
                    if 'icon_label' in row_data:
                        row_data['icon_label'].configure(bg=surface_color)
                    if 'text_label' in row_data:
                        row_data['text_label'].configure(bg=surface_color, fg=fg_color)
                except Exception:
                    pass

        # Update diagram separator line
        if hasattr(self, '_diagram_separator') and self._diagram_separator:
            try:
                self._diagram_separator.configure(bg=colors.get('border', '#e0e0e0'))
            except Exception:
                pass

        # Update diagram status row (frame + labels)
        if hasattr(self, '_diagram_status_row') and self._diagram_status_row:
            try:
                self._diagram_status_row.configure(bg=header_bg_color)
            except Exception:
                pass
        if hasattr(self, '_diagram_status') and self._diagram_status:
            try:
                self._diagram_status.configure(fg=colors['text_secondary'], bg=header_bg_color)
            except Exception:
                pass
        if hasattr(self, '_diagram_tip') and self._diagram_tip:
            try:
                self._diagram_tip.configure(fg=colors['text_secondary'], bg=header_bg_color)
            except Exception:
                pass

        # Update analysis summary placeholder
        if hasattr(self, '_summary_placeholder') and self._summary_placeholder:
            try:
                self._summary_placeholder.configure(
                    fg=colors['text_secondary'],
                    bg=bg_color
                )
            except Exception:
                pass

        # Update metrics frame background
        if hasattr(self, '_metrics_frame') and self._metrics_frame:
            try:
                self._metrics_frame.configure(bg=bg_color)
            except Exception:
                pass

        # Update analysis summary card frames
        card_bg = colors.get('card_surface', colors['surface'])
        border_color = colors.get('border', '#333')
        if hasattr(self, '_summary_table_frame') and self._summary_table_frame:
            try:
                self._summary_table_frame.configure(bg=card_bg, highlightbackground=border_color, highlightcolor=border_color)
            except Exception:
                pass
        if hasattr(self, '_summary_categorization_frame') and self._summary_categorization_frame:
            try:
                self._summary_categorization_frame.configure(bg=card_bg, highlightbackground=border_color, highlightcolor=border_color)
            except Exception:
                pass
        if hasattr(self, '_summary_container') and self._summary_container:
            try:
                self._summary_container.configure(bg=bg_color)
            except Exception:
                pass
        if hasattr(self, '_metrics_frame') and self._metrics_frame:
            try:
                self._metrics_frame.configure(bg=bg_color)
            except Exception:
                pass

        # Update metric labels in analysis summary
        if hasattr(self, '_metric_labels'):
            for key, lbl in self._metric_labels.items():
                try:
                    # Labels inside card frames use card_bg
                    label_bg = card_bg
                    if key == 'cat_header':
                        lbl.configure(fg=colors['title_color'], bg=label_bg)
                    elif key == 'score':
                        # Keep score label color as-is (determined by quality)
                        lbl.configure(bg=label_bg)
                    else:
                        lbl.configure(fg=colors['text_primary'], bg=label_bg)
                except Exception:
                    pass

        # Update bottom action buttons canvas_bg for proper corner rounding on outer background
        outer_canvas_bg = colors['section_bg']
        try:
            if hasattr(self, 'optimize_btn') and self.optimize_btn:
                self.optimize_btn.update_colors(
                    bg=colors['button_primary'],
                    hover_bg=colors['button_primary_hover'],
                    pressed_bg=colors['button_primary_pressed'],
                    fg='#ffffff',
                    disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
                    disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
                    canvas_bg=outer_canvas_bg
                )
            if hasattr(self, 'reset_btn') and self.reset_btn:
                self.reset_btn.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    canvas_bg=outer_canvas_bg
                )
        except Exception:
            pass
        # Note: FileInputSection handles its own theme updates internally (browse_btn, analyze_btn, help button)

    # Required abstract method implementations for BaseToolTab
    def setup_ui(self):
        """Setup UI - already done in _create_interface"""
        pass  # Interface is created in __init__

    def reset_tab(self):
        """Reset tab to initial state"""
        self._reset_interface()

    def on_theme_changed(self, theme: str):
        """Update theme-dependent widgets when theme changes"""
        super().on_theme_changed(theme)

        # Reload themed checkbox icons (light/dark variants + partial)
        self._load_checkbox_icons()

        # Update existing diagram items with new icons if they exist
        if hasattr(self, '_diagram_row_widgets') and self._diagram_row_widgets:
            for row_data in self._diagram_row_widgets:
                if 'icon_label' in row_data:
                    try:
                        idx = row_data.get('idx')
                        is_select_all = row_data.get('is_select_all', False)

                        # Determine icon based on selection state
                        if is_select_all:
                            # Select All row: checked if all, partial if some, unchecked if none
                            if self._is_all_selected():
                                icon = self._checkbox_on_icon
                            elif self._is_some_selected():
                                icon = self._checkbox_partial_icon
                            else:
                                icon = self._checkbox_off_icon
                        else:
                            # Individual diagram rows: checked or unchecked
                            is_selected = idx in self.selected_diagram_indices if isinstance(idx, int) else False
                            icon = self._checkbox_on_icon if is_selected else self._checkbox_off_icon

                        if icon:
                            row_data['icon_label'].configure(image=icon)
                    except Exception:
                        pass

    def show_help_dialog(self):
        """Show context-sensitive help dialog"""
        # Get the correct parent window - try multiple possible parent references
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
        help_window.title("Layout Optimizer - Help")
        help_window.geometry("1000x750")  # Wider and shorter
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
        """Create help content for the layout optimizer"""
        from core.ui_base import RoundedButton
        colors = self._theme_manager.colors

        # Consistent help dialog background for all tools
        help_bg = colors['background']
        help_window.configure(bg=help_bg)

        # Main container - use tk.Frame with explicit bg for consistency
        container = tk.Frame(help_window, bg=help_bg)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header - use tk.Label with explicit bg for dark mode
        tk.Label(container, text="Layout Optimizer Help",
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
            "Requires PBIP format for TMDL relationship analysis",
            "This is NOT officially supported by Microsoft - use at your own discretion",
            "Uses TMDL files for table categorization and relationship data",
            "Always keep backups of your original projects before optimization",
            "Test thoroughly and validate optimized layouts before production use"
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
            "Analyzes your current relationship diagram layout quality",
            "Provides layout quality scoring (0-100 scale with ratings)",
            "Categorizes tables by type (Facts, Dimensions L1-L4+, Special tables)",
            "Applies Haven's middle-out design philosophy for professional layouts",
            "Optimizes spacing and reduces overlapping elements",
            "Designed for PBIP projects with TMDL files"
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
            "Select the .pbip file in your Power BI project folder",
            "Project must contain .SemanticModel folder with TMDL files",
            "TMDL files are required for table categorization and relationships",
            "Write permissions required for saving optimized layouts",
            "Backup your project before running optimization"
        ]

        for item in file_items:
            tk.Label(right_top_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=450,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        # ROW 1, COLUMN 0: Middle-Out Design
        left_bottom_frame = tk.Frame(sections_frame, bg=help_bg)
        left_bottom_frame.grid(row=1, column=0, sticky='nwe', padx=(0, 10))

        tk.Label(left_bottom_frame, text="Middle-Out Design",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        philosophy_items = [
            "Fact tables positioned centrally as the data foundation",
            "L1 dimensions arranged around facts for direct relationships",
            "L2+ dimensions positioned in outer layers by hierarchy",
            "Special tables (Calendar, Parameters, Metrics) grouped logically",
            "Optimized spacing prevents overlapping and improves readability",
            "Professional appearance suitable for executive presentations"
        ]

        for item in philosophy_items:
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
            "Requires PBIP format for TMDL relationship access",
            "This tool is NOT officially supported by Microsoft",
            "Always backup your project files before optimization",
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
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 750) // 2
        help_window.geometry(f"1000x750+{x}+{y}")

        # Set dark title bar BEFORE showing to prevent white flash
        help_window.update()
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled
