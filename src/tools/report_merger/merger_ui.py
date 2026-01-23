"""
Report Merger UI Tab - Refactored to use BaseToolTab
Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from pathlib import Path
from typing import Optional, Dict, Any

from core.constants import AppConstants
from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin, SquareIconButton, SVGToggle, RoundedButton, ActionButtonBar, SplitLogSection
from tools.report_merger.merger_core import MergerEngine, ValidationService, ValidationError


class ReportMergerTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """Report Merger tab - refactored to use new base architecture"""
    
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "report_merger", "Report Merger")
        
        # UI Variables
        self.report_a_path = tk.StringVar()
        self.report_b_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.theme_choice = tk.StringVar(value="report_a")
        
        # Core components
        self.validation_service = ValidationService()
        self.merger_engine = MergerEngine(logger_callback=self.log_message)
        
        # UI Components
        self.analyze_button = None
        self.merge_button = None
        self.theme_frame = None
        self.log_section = None  # Store log section for positioning
        
        # State
        self.analysis_results = None
        self.is_analyzing = False
        self.is_merging = False

        # Analysis summary table labels (for updating values)
        self._summary_labels = {}
        self._summary_title_label = None
        self._summary_status_label = None

        # Setup UI and events
        self.setup_ui()
        self._setup_events()

    def setup_ui(self) -> None:
        """Setup the UI for the report merger tab"""
        # Load UI icons for buttons and section headers
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}
        icon_names = ["Power-BI", "save", "paint", "folder", "magnifying-glass", "execute", "bar-chart", "reset", "analyze"]
        for name in icon_names:
            icon = self._load_icon_for_button(name, size=16)
            if icon:
                self._button_icons[name] = icon

        # IMPORTANT: Create action buttons FIRST with side=BOTTOM
        # This ensures buttons are always visible even when window shrinks
        execute_icon = self._button_icons.get('execute')
        reset_icon = self._button_icons.get('reset')

        # Use centralized ActionButtonBar
        self.button_frame = ActionButtonBar(
            parent=self.frame,
            theme_manager=self._theme_manager,
            primary_text="EXECUTE MERGE",
            primary_command=self.start_merge,
            primary_icon=execute_icon,
            secondary_text="RESET ALL",
            secondary_command=self.reset_tab,
            secondary_icon=reset_icon,
            primary_starts_disabled=True
        )
        self.button_frame.pack(side=tk.BOTTOM, pady=(15, 0))

        # Expose buttons for compatibility with existing code
        self.merge_button = self.button_frame.primary_button
        self.reset_btn = self.button_frame.secondary_button
        self._merge_button_enabled = False

        # Track in button lists for any additional theme handling
        self._primary_buttons.append(self.merge_button)
        self._secondary_buttons.append(self.reset_btn)

        # Now create other sections from top (they will fill remaining space)
        self._setup_data_sources()
        self._setup_theme_selection()
        self._setup_output_section()

        # Create split log section using SplitLogSection template widget
        self.log_section = SplitLogSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="Analysis & Progress",
            section_icon="analyze",
            summary_title="Analysis Summary",
            summary_icon="bar-chart",
            log_title="Progress Log",
            log_icon="log-file",
            summary_placeholder="Analyze reports to see comparison"
        )
        self.log_section.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Connect log_text to base class for log_message() method
        self.log_text = self.log_section.log_text

        # Create permanent analysis table structure in summary panel
        self._create_analysis_table()

        # Create progress bar using base class
        self.create_progress_bar(self.frame)

        # Show welcome message
        self._show_welcome_message()
    
    def _position_progress_frame(self):
        """Position progress frame specifically for Report Merger layout"""
        if self.progress_frame and self.button_frame:
            # Position progress bar above the button frame using side=BOTTOM
            # Must use side=BOTTOM since button_frame was packed with side=BOTTOM
            self.progress_frame.pack(side=tk.BOTTOM, before=self.button_frame, fill=tk.X, pady=(10, 5))
        elif self.progress_frame:
            # Fallback: position at bottom if button frame not found
            self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10))
    
    def _setup_data_sources(self):
        """Setup data sources section - single-column layout like Report Cleanup"""
        colors = self._theme_manager.colors

        # LabelFrame with custom labelwidget for icon + text header
        header_widget = self.create_section_header(self.frame, "PBIP File Source", "Power-BI")[0]
        section_frame = ttk.LabelFrame(self.frame, labelwidget=header_widget,
                                     style='Section.TLabelframe', padding="12")
        section_frame.pack(fill=tk.X, pady=(0, 15))
        self._data_source_section = section_frame

        content_frame = ttk.Frame(section_frame, style='Section.TFrame', padding="15")
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Load help icon and create button - positioned in upper right corner
        # Created AFTER content_frame to ensure proper stacking order
        help_icon = self._load_icon_for_button("question", size=14)
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}
        self._button_icons['question'] = help_icon

        self._help_button = SquareIconButton(
            section_frame, icon=help_icon, command=self.show_help_dialog,
            tooltip_text="Help", size=26, radius=6,
            bg_normal_override=AppConstants.CORNER_ICON_BG
        )
        # Position in upper-right corner of section title bar area (y=-6 places in title region)
        # Created after content_frame so it's already on top in stacking order
        self._help_button.place(relx=1.0, y=-35, anchor=tk.NE, x=-0)
        content_frame.columnconfigure(1, weight=1)

        # Import RoundedButton for browse buttons
        from core.ui_base import RoundedButton

        # canvas_bg for corner rounding: #0d0d1a dark / #ffffff light (main background)
        is_dark = self._theme_manager.is_dark
        file_section_canvas_bg = colors['background']  # Use theme color instead of hardcoded

        # Report A input row - consistent naming: "Project File A (PBIP):"
        ttk.Label(content_frame, text="Project File A (PBIP):",
                  style='Section.TLabel', font=('Segoe UI', 10)).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))

        entry_a = ttk.Entry(content_frame, textvariable=self.report_a_path,
                            font=('Segoe UI', 10), style='Section.TEntry')
        entry_a.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        folder_icon = self._button_icons.get('folder')
        browse_a_btn = RoundedButton(
            content_frame, text="Browse" if folder_icon else "Browse",
            command=lambda: self._browse_file(self.report_a_path, "Report A"),
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            width=90, height=32, radius=6, font=('Segoe UI', 10),
            icon=folder_icon,
            canvas_bg=file_section_canvas_bg
        )
        browse_a_btn.grid(row=0, column=2)
        self._primary_buttons.append(browse_a_btn)

        # Report B input row - consistent naming: "Project File B (PBIP):"
        ttk.Label(content_frame, text="Project File B (PBIP):",
                  style='Section.TLabel', font=('Segoe UI', 10)).grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(8, 0))

        entry_b = ttk.Entry(content_frame, textvariable=self.report_b_path,
                            font=('Segoe UI', 10), style='Section.TEntry')
        entry_b.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(8, 0))

        browse_b_btn = RoundedButton(
            content_frame, text="Browse" if folder_icon else "Browse",
            command=lambda: self._browse_file(self.report_b_path, "Report B"),
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            width=90, height=32, radius=6, font=('Segoe UI', 10),
            icon=folder_icon,
            canvas_bg=file_section_canvas_bg
        )
        browse_b_btn.grid(row=1, column=2, pady=(8, 0))
        self._primary_buttons.append(browse_b_btn)

        # Analyze button - using action button for proper hover/click states
        magnify_icon = self._button_icons.get('magnifying-glass')
        self.analyze_button = self.create_action_button(
            content_frame, "ANALYZE REPORTS" if magnify_icon else "🔍 ANALYZE REPORTS",
            self.analyze_reports, icon=magnify_icon)
        self.analyze_button.grid(row=2, column=0, columnspan=3, pady=(15, 0))
        # Start disabled until both paths are provided
        self.analyze_button.set_enabled(False)
        self._analyze_button_enabled = False

        # Setup path cleaning
        self.setup_path_cleaning(self.report_a_path)
        self.setup_path_cleaning(self.report_b_path)
    
    def _setup_theme_selection(self):
        """Setup theme selection - now inline in Analysis Summary panel"""
        # Theme selection is now shown inline in the Analysis Summary panel
        # instead of a separate section, to avoid pushing buttons off-screen
        self.theme_frame = None
        self._theme_selection_frame = None
        self._theme_radio_buttons = []
        # Store references to theme selection widgets for theme updates
        self._theme_warning_lbl = None
        self._theme_row_a = None
        self._theme_row_b = None
    
    def _setup_output_section(self):
        """Setup output section"""
        header_widget = self.create_section_header(self.frame, "Output Configuration", "save")[0]
        output_frame = ttk.LabelFrame(self.frame, labelwidget=header_widget,
                                    style='Section.TLabelframe', padding="12")
        output_frame.pack(fill=tk.X, pady=(0, 15))

        # Inner content frame with white background and internal padding
        content_frame = ttk.Frame(output_frame, style='Section.TFrame', padding="15")
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(1, weight=1)

        # Import RoundedButton for consistent browse button styling
        from core.ui_base import RoundedButton
        colors = self._theme_manager.colors

        # canvas_bg for corner rounding - use theme color
        is_dark = self._theme_manager.is_dark
        file_section_canvas_bg = colors['background']

        ttk.Label(content_frame, text="Output Path:",
                  style='Section.TLabel', font=('Segoe UI', 10)).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))

        ttk.Entry(content_frame, textvariable=self.output_path,
                  font=('Segoe UI', 10), style='Section.TEntry').grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        # Browse button with folder icon - consistent styling
        folder_icon = self._button_icons.get('folder')
        browse_out_btn = RoundedButton(
            content_frame, text="Browse" if folder_icon else "Browse",
            command=self.browse_output,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            width=90, height=32, radius=6, font=('Segoe UI', 10),
            icon=folder_icon,
            canvas_bg=file_section_canvas_bg
        )
        browse_out_btn.grid(row=0, column=2)
        self._primary_buttons.append(browse_out_btn)
        
        # Setup path cleaning for output
        self.setup_path_cleaning(self.output_path)
    
    def _setup_events(self):
        """Setup event handlers"""
        self.report_a_path.trace('w', lambda *args: self._on_path_change())
        self.report_b_path.trace('w', lambda *args: self._on_path_change())

    def _create_analysis_table(self):
        """Create permanent analysis table structure with placeholder values"""
        colors = self._theme_manager.colors
        summary_frame = self.log_section.summary_frame

        # Set section_bg for the container with padding (consistent with Advanced Copy)
        summary_frame.configure(bg=colors['section_bg'], padx=12, pady=8)

        # Clear any existing placeholder
        for widget in summary_frame.winfo_children():
            widget.destroy()

        # Title label (no background - transparent to section_bg)
        self._summary_title_label = tk.Label(
            summary_frame, text="Select reports to analyze",
            font=('Segoe UI', 10, 'bold'),
            bg=colors['section_bg'], fg=colors['text_primary'])
        self._summary_title_label.pack(anchor=tk.W, pady=(0, 10))

        # Table frame with card background (white in light mode) - stays at top, doesn't expand
        table_frame = tk.Frame(summary_frame, bg=colors['card_surface'],
                              highlightbackground=colors['border'],
                              highlightthickness=1, padx=12, pady=12)
        table_frame.pack(fill=tk.X, pady=(0, 10))

        # Configure columns for even distribution
        for col in range(4):
            table_frame.columnconfigure(col, weight=1, minsize=80)

        # Store reference to table frame for theme updates
        self._summary_table_frame = table_frame

        # Table headers
        headers = ["Metric", "Report A", "Report B", "Total"]
        self._summary_header_labels = []
        for col, header in enumerate(headers):
            lbl = tk.Label(table_frame, text=header, font=('Segoe UI', 9, 'bold'),
                          bg=colors['card_surface'], fg=colors['text_primary'])
            lbl.grid(row=0, column=col, padx=(0, 15), pady=(0, 8), sticky=tk.W)
            self._summary_header_labels.append(lbl)

        # Data rows - initialize with placeholders
        metrics = [
            ("pages", "📄 Pages"),
            ("bookmarks", "🔖 Bookmarks"),
            ("measures", "📐 Measures"),
        ]

        self._summary_labels = {}
        for row_idx, (key, label) in enumerate(metrics, start=1):
            # Metric name column
            metric_lbl = tk.Label(table_frame, text=label, font=('Segoe UI', 9),
                                 bg=colors['card_surface'], fg=colors['text_primary'])
            metric_lbl.grid(row=row_idx, column=0, padx=(0, 15), pady=4, sticky=tk.W)

            # Value columns - store references for later updates
            self._summary_labels[f'{key}_metric'] = metric_lbl
            for col_idx, col_key in enumerate(['a', 'b', 'total'], start=1):
                font = ('Segoe UI', 9) if col_key != 'total' else ('Segoe UI', 9, 'bold')
                val_lbl = tk.Label(table_frame, text="--", font=font,
                                  bg=colors['card_surface'], fg=colors['text_primary'])
                val_lbl.grid(row=row_idx, column=col_idx, padx=(0, 15), pady=4, sticky=tk.W)
                self._summary_labels[f'{key}_{col_key}'] = val_lbl

        # Report totals row (no separator, directly after measures)
        total_lbl = tk.Label(table_frame, text="Report Total", font=('Segoe UI', 9, 'bold'),
                            bg=colors['card_surface'], fg=colors['text_primary'])
        total_lbl.grid(row=len(metrics)+1, column=0, padx=(0, 15), pady=(8, 4), sticky=tk.W)
        self._summary_labels['column_total_metric'] = total_lbl

        for col_idx, col_key in enumerate(['a', 'b', 'total'], start=1):
            val_lbl = tk.Label(table_frame, text="--", font=('Segoe UI', 9, 'bold'),
                              bg=colors['card_surface'], fg=colors['text_primary'])
            val_lbl.grid(row=len(metrics)+1, column=col_idx, padx=(0, 15), pady=(8, 4), sticky=tk.W)
            self._summary_labels[f'column_total_{col_key}'] = val_lbl

        # Status label (centered, no background - transparent to section_bg)
        self._summary_status_label = tk.Label(
            summary_frame, text="⏳ Waiting for analysis...",
            font=('Segoe UI', 9, 'italic'),
            bg=colors['section_bg'], fg=colors['text_secondary'])
        self._summary_status_label.pack(anchor=tk.CENTER, pady=(5, 0))

    def on_theme_changed(self, theme: str):
        """Update theme-dependent widgets when theme changes"""
        super().on_theme_changed(theme)
        colors = self._theme_manager.colors

        # Update inline theme selection radio buttons if they exist
        if hasattr(self, '_theme_radio_buttons') and self._theme_radio_buttons:
            bg = colors['section_bg']  # Now inline in Analysis Summary panel
            is_dark = colors.get('background', '') == '#0d0d1a'
            hover_color = '#004466' if is_dark else colors['primary']
            for rb in self._theme_radio_buttons:
                try:
                    rb.configure(
                        fg=colors['text_primary'],
                        bg=bg,
                        selectcolor=bg,
                        activeforeground=hover_color,
                        activebackground=bg
                    )
                except Exception:
                    pass

        # Update theme selection frame and all its widgets
        if hasattr(self, '_theme_selection_frame') and self._theme_selection_frame:
            section_bg = colors['section_bg']
            try:
                self._theme_selection_frame.configure(bg=section_bg)

                # Update warning label (icon + text header)
                if hasattr(self, '_theme_warning_lbl') and self._theme_warning_lbl:
                    self._theme_warning_lbl.configure(bg=section_bg)
                    for child in self._theme_warning_lbl.winfo_children():
                        child.configure(bg=section_bg)
                        # Update text label fg color (warning orange)
                        if child.cget('text') if hasattr(child, 'cget') else '':
                            try:
                                child.configure(fg=colors.get('warning', '#f5751f'))
                            except Exception:
                                pass

                # Update row frames and labels
                if hasattr(self, '_theme_row_a') and self._theme_row_a:
                    self._theme_row_a.configure(bg=section_bg)
                if hasattr(self, '_theme_row_b') and self._theme_row_b:
                    self._theme_row_b.configure(bg=section_bg)
                if hasattr(self, '_theme_label_a') and self._theme_label_a:
                    self._theme_label_a.configure(bg=section_bg, fg=colors['text_primary'])
                if hasattr(self, '_theme_label_b') and self._theme_label_b:
                    self._theme_label_b.configure(bg=section_bg, fg=colors['text_primary'])

            except Exception:
                pass

        # Update Analysis Summary widgets
        self._update_analysis_summary_theme(colors)

        # Update bottom action buttons canvas_bg for proper corner rounding on outer background
        outer_canvas_bg = colors['section_bg']
        try:
            if hasattr(self, 'merge_button') and self.merge_button:
                self.merge_button.update_colors(
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

    def _update_analysis_summary_theme(self, colors):
        """Update all Analysis Summary widgets for theme change"""
        # SplitLogSection handles its own theme updates internally

        # Update title label
        if hasattr(self, '_summary_title_label') and self._summary_title_label:
            try:
                self._summary_title_label.configure(
                    bg=colors['section_bg'], fg=colors['text_primary'])
            except Exception:
                pass

        # Update table frame
        if hasattr(self, '_summary_table_frame') and self._summary_table_frame:
            try:
                self._summary_table_frame.configure(
                    bg=colors['card_surface'], highlightbackground=colors['border'])
            except Exception:
                pass

        # Update header labels
        if hasattr(self, '_summary_header_labels') and self._summary_header_labels:
            for lbl in self._summary_header_labels:
                try:
                    lbl.configure(bg=colors['card_surface'], fg=colors['text_primary'])
                except Exception:
                    pass

        # Update data labels
        if hasattr(self, '_summary_labels') and self._summary_labels:
            for lbl in self._summary_labels.values():
                try:
                    lbl.configure(bg=colors['card_surface'], fg=colors['text_primary'])
                except Exception:
                    pass

        # Update status label
        if hasattr(self, '_summary_status_label') and self._summary_status_label:
            try:
                self._summary_status_label.configure(bg=colors['section_bg'])
                current_text = str(self._summary_status_label.cget('text'))
                if 'Waiting' in current_text:
                    self._summary_status_label.configure(fg=colors['text_secondary'])
                elif 'Ready' in current_text:
                    self._summary_status_label.configure(fg=colors.get('success', '#22c55e'))
                elif 'Conflict' in current_text:
                    self._summary_status_label.configure(fg=colors.get('warning', '#f59e0b'))
            except Exception:
                pass

    def _show_welcome_message(self):
        """Show welcome message"""
        self.log_message("🔀 Welcome to Report Merger!")
        self.log_message("=" * 60)
        self.log_message("This tool merges two Power BI reports into one:")
        self.log_message("• Combine pages from Report A and Report B")
        self.log_message("• Merge themes, custom visuals, and bookmarks")
        self.log_message("• Preserve visual configurations and styles")
        self.log_message("")
        self.log_message("📁 Start by selecting your Report A and Report B files")
        self.log_message("⚠️ Requires PBIP format files only")
    
    def _on_path_change(self):
        """Handle path changes"""
        self._update_ui_state()
        self.auto_generate_output_path()

    def _on_theme_toggle_select(self, which: str, state: bool):
        """Handle theme toggle state change with mutual exclusion"""
        if which == 'a':
            if state:
                # A turned on - turn off B (no callback to avoid loop)
                self._theme_toggle_b.set_state(False, trigger_command=False)
                self.theme_choice.set("report_a")
            else:
                # A turned off - turn on B (no callback to avoid loop)
                self._theme_toggle_b.set_state(True, trigger_command=False)
                self.theme_choice.set("report_b")
        else:  # which == 'b'
            if state:
                # B turned on - turn off A (no callback to avoid loop)
                self._theme_toggle_a.set_state(False, trigger_command=False)
                self.theme_choice.set("report_b")
            else:
                # B turned off - turn on A (no callback to avoid loop)
                self._theme_toggle_a.set_state(True, trigger_command=False)
                self.theme_choice.set("report_a")

    def _select_theme(self, which: str):
        """Select a theme option (called from row/label clicks)"""
        if which == 'a':
            self._theme_toggle_a.set_state(True, trigger_command=False)
            self._theme_toggle_b.set_state(False, trigger_command=False)
            self.theme_choice.set("report_a")
        else:
            self._theme_toggle_a.set_state(False, trigger_command=False)
            self._theme_toggle_b.set_state(True, trigger_command=False)
            self.theme_choice.set("report_b")

    def _update_ui_state(self):
        """Update UI state"""
        has_both = bool(self.report_a_path.get() and self.report_b_path.get())
        if self.analyze_button:
            self.analyze_button.set_enabled(has_both)
            self._analyze_button_enabled = has_both
    
    def _browse_file(self, path_var: tk.StringVar, report_name: str):
        """Browse for report file"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title=f"Select {report_name} (.pbip file - PBIR format required)",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            path_var.set(file_path)
    
    def browse_output(self):
        """Browse for output location"""
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(
            title="Save Combined Report As", defaultextension=".pbip",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)
    
    def auto_generate_output_path(self):
        """Auto-generate output path"""
        report_a = self.clean_file_path(self.report_a_path.get())
        report_b = self.clean_file_path(self.report_b_path.get())
        
        if report_a and report_b:
            try:
                output_path = self.merger_engine.generate_output_path(report_a, report_b)
                self.output_path.set(output_path)
            except Exception:
                pass
    
    def analyze_reports(self):
        """Analyze selected reports"""
        try:
            report_a = self.clean_file_path(self.report_a_path.get())
            report_b = self.clean_file_path(self.report_b_path.get())
            self.validation_service.validate_input_paths(report_a, report_b)
        except Exception as e:
            self.show_error("Validation Error", str(e))
            return
        
        # Use base class background processing
        self.run_in_background(
            target_func=self._analyze_thread_target,
            success_callback=self._handle_analysis_complete,
            error_callback=lambda e: self.show_error("Analysis Error", str(e))
        )
    
    def _analyze_thread_target(self):
        """Background analysis logic"""
        self.update_progress(10, "Validating input files...")
        report_a = self.clean_file_path(self.report_a_path.get())
        report_b = self.clean_file_path(self.report_b_path.get())
        
        self.update_progress(30, "Reading source report...")
        
        self.update_progress(50, "Reading target report...")
        
        self.update_progress(80, "Analyzing compatibility...")
        results = self.merger_engine.analyze_reports(report_a, report_b)
        
        self.update_progress(100, "Analysis complete!")
        return results
    
    def _handle_analysis_complete(self, results):
        """Handle analysis completion"""
        self.analysis_results = results

        # Set theme choice if no conflict
        if not results['themes']['conflict']:
            self.theme_choice.set("same")

        # Enable merge button
        if self.merge_button:
            self.merge_button.set_enabled(True)
            self._merge_button_enabled = True

        # Show analysis summary with inline theme selection if conflict
        self._show_analysis_summary(results)
    def _show_analysis_summary(self, results):
        """Update analysis summary values in the existing table structure"""
        colors = self._theme_manager.colors

        report_a = results['report_a']
        report_b = results['report_b']
        totals = results['totals']
        themes = results['themes']

        # Update title
        if self._summary_title_label:
            self._summary_title_label.configure(
                text=f"Combining: {report_a['name']} + {report_b['name']}")

        # Update data values
        metrics = ['pages', 'bookmarks', 'measures']
        col_totals_a = 0
        col_totals_b = 0
        col_totals_total = 0

        for metric in metrics:
            val_a = report_a.get(metric, 0)
            val_b = report_b.get(metric, 0)
            val_total = totals.get(metric, val_a + val_b)

            col_totals_a += val_a
            col_totals_b += val_b
            col_totals_total += val_total

            # Update the label values
            if f'{metric}_a' in self._summary_labels:
                self._summary_labels[f'{metric}_a'].configure(text=str(val_a))
            if f'{metric}_b' in self._summary_labels:
                self._summary_labels[f'{metric}_b'].configure(text=str(val_b))
            if f'{metric}_total' in self._summary_labels:
                self._summary_labels[f'{metric}_total'].configure(text=str(val_total))

        # Update column totals
        if 'column_total_a' in self._summary_labels:
            self._summary_labels['column_total_a'].configure(text=str(col_totals_a))
        if 'column_total_b' in self._summary_labels:
            self._summary_labels['column_total_b'].configure(text=str(col_totals_b))
        if 'column_total_total' in self._summary_labels:
            self._summary_labels['column_total_total'].configure(text=str(col_totals_total))

        # Clean up existing theme selection frame if it exists
        if hasattr(self, '_theme_selection_frame') and self._theme_selection_frame:
            self._theme_selection_frame.destroy()
            self._theme_selection_frame = None
            self._theme_warning_lbl = None
            self._theme_row_a = None
            self._theme_row_b = None

        # Add inline theme selection if conflict detected
        if themes['conflict']:
            summary_frame = self.log_section.summary_frame

            # Create theme selection frame (between table and status)
            self._theme_selection_frame = tk.Frame(summary_frame, bg=colors['section_bg'])
            # Pack before status label
            self._theme_selection_frame.pack(fill=tk.X, pady=(8, 15), before=self._summary_status_label)

            # Warning label - uses warning.svg icon
            self._theme_warning_lbl = self.create_icon_label(
                self._theme_selection_frame, icon_name="warning",
                text="Theme Conflict - Select theme:",
                icon_size=16, font=('Segoe UI Semibold', 11),
                fg_color=colors.get('warning', '#f5751f'),
                bg_color=colors['section_bg']
            )
            self._theme_warning_lbl.pack(anchor=tk.W)

            # Get theme display names
            theme_a_display = themes['theme_a'].get('display', 'Theme A') if isinstance(themes['theme_a'], dict) else str(themes['theme_a'])
            theme_b_display = themes['theme_b'].get('display', 'Theme B') if isinstance(themes['theme_b'], dict) else str(themes['theme_b'])

            # SVG paths for toggles
            toggle_on_path = Path(__file__).parent.parent.parent / "assets" / "Tool Icons" / "toggle-on.svg"
            toggle_off_path = Path(__file__).parent.parent.parent / "assets" / "Tool Icons" / "toggle-off.svg"

            # Row A: Toggle | Label
            self._theme_row_a = tk.Frame(self._theme_selection_frame, bg=colors['section_bg'], cursor='hand2')
            self._theme_row_a.pack(fill=tk.X, pady=(5, 2))

            self._theme_toggle_a = SVGToggle(
                self._theme_row_a,
                svg_on=str(toggle_on_path),
                svg_off=str(toggle_off_path),
                command=lambda state: self._on_theme_toggle_select('a', state),
                initial_state=True,  # Default to Report A selected
                width=35, height=17,  # 30% smaller
                theme_manager=self._theme_manager
            )
            self._theme_toggle_a.pack(side=tk.LEFT, padx=(0, 8))

            self._theme_label_a = tk.Label(self._theme_row_a, text=f"Report A: {theme_a_display}",
                font=('Segoe UI', 9), fg=colors['text_primary'], bg=colors['section_bg'],
                cursor='hand2')
            self._theme_label_a.pack(side=tk.LEFT)

            # Row B: Toggle | Label
            self._theme_row_b = tk.Frame(self._theme_selection_frame, bg=colors['section_bg'], cursor='hand2')
            self._theme_row_b.pack(fill=tk.X, pady=(2, 0))

            self._theme_toggle_b = SVGToggle(
                self._theme_row_b,
                svg_on=str(toggle_on_path),
                svg_off=str(toggle_off_path),
                command=lambda state: self._on_theme_toggle_select('b', state),
                initial_state=False,  # Default to Report B not selected
                width=35, height=17,  # 30% smaller
                theme_manager=self._theme_manager
            )
            self._theme_toggle_b.pack(side=tk.LEFT, padx=(0, 8))

            self._theme_label_b = tk.Label(self._theme_row_b, text=f"Report B: {theme_b_display}",
                font=('Segoe UI', 9), fg=colors['text_primary'], bg=colors['section_bg'],
                cursor='hand2')
            self._theme_label_b.pack(side=tk.LEFT)

            # Click entire row to select that option
            self._theme_row_a.bind('<Button-1>', lambda e: self._select_theme('a'))
            self._theme_label_a.bind('<Button-1>', lambda e: self._select_theme('a'))
            self._theme_row_b.bind('<Button-1>', lambda e: self._select_theme('b'))
            self._theme_label_b.bind('<Button-1>', lambda e: self._select_theme('b'))

            # Hover effects - underline on hover
            def on_enter_a(e):
                self._theme_label_a.configure(font=('Segoe UI', 9, 'underline'))
            def on_leave_a(e):
                self._theme_label_a.configure(font=('Segoe UI', 9))
            def on_enter_b(e):
                self._theme_label_b.configure(font=('Segoe UI', 9, 'underline'))
            def on_leave_b(e):
                self._theme_label_b.configure(font=('Segoe UI', 9))

            self._theme_row_a.bind('<Enter>', on_enter_a)
            self._theme_row_a.bind('<Leave>', on_leave_a)
            self._theme_label_a.bind('<Enter>', on_enter_a)
            self._theme_label_a.bind('<Leave>', on_leave_a)
            self._theme_row_b.bind('<Enter>', on_enter_b)
            self._theme_row_b.bind('<Leave>', on_leave_b)
            self._theme_label_b.bind('<Enter>', on_enter_b)
            self._theme_label_b.bind('<Leave>', on_leave_b)

            # Set default selection
            self.theme_choice.set("report_a")

        # Update status
        if self._summary_status_label:
            if themes['conflict']:
                self._summary_status_label.configure(
                    text="✅ Ready to Merge",
                    fg=colors.get('success', '#22c55e'),
                    font=('Segoe UI', 9, 'bold'))
            else:
                self._summary_status_label.configure(
                    text="✅ Ready to Merge",
                    fg=colors.get('success', '#22c55e'),
                    font=('Segoe UI', 9, 'bold'))

        # Also log to progress log for reference
        self.log_message("✅ Analysis complete! Ready to merge.")

    def _clear_summary_panel(self):
        """Reset summary panel values to placeholders"""
        self._reset_analysis_values()

    def _restore_summary_placeholder(self):
        """Reset summary panel values to placeholders"""
        self._reset_analysis_values()

    def _reset_analysis_values(self):
        """Reset all analysis values to placeholder state"""
        colors = self._theme_manager.colors

        # Reset title
        if self._summary_title_label:
            self._summary_title_label.configure(text="📋 Select reports to analyze")

        # Reset all data values to "--"
        for key in self._summary_labels:
            if not key.endswith('_metric'):
                self._summary_labels[key].configure(text="--")

        # Clean up inline theme selection frame if it exists
        if hasattr(self, '_theme_selection_frame') and self._theme_selection_frame:
            self._theme_selection_frame.destroy()
            self._theme_selection_frame = None
            self._theme_warning_lbl = None
            self._theme_row_a = None
            self._theme_row_b = None
        self._theme_radio_buttons = []

        # Reset status
        if self._summary_status_label:
            self._summary_status_label.configure(
                text="⏳ Waiting for analysis...",
                fg=colors['text_secondary'],
                font=('Segoe UI', 9, 'italic'))
    
    def start_merge(self):
        """Start merge operation"""
        if not self.analysis_results:
            self.show_error("Error", "Please analyze reports first")
            return
        
        try:
            output_path = self.clean_file_path(self.output_path.get())
            self.validation_service.validate_output_path(output_path)
        except Exception as e:
            self.show_error("Output Validation Error", str(e))
            return
        
        totals = self.analysis_results['totals']
        # Use Report Merger icon in confirmation dialog
        icon_path = Path(__file__).parent.parent.parent / "assets" / "Tool Icons" / "report merger.svg"
        if not self.ask_yes_no("Confirm Merge",
                              f"Ready to merge reports?\n\n"
                              f"📊 Combined report will have:\n"
                              f"📄 {totals['pages']} pages\n"
                              f"🔖 {totals['bookmarks']} bookmarks\n"
                              f"📐 {totals.get('measures', 0)} measures\n\n"
                              f"💾 Output: {output_path}",
                              icon_path=str(icon_path)):
            return
        
        # Use base class background processing
        self.run_in_background(
            target_func=self._merge_thread_target,
            success_callback=self._handle_merge_complete,
            error_callback=lambda e: self.show_error("Merge Error", str(e))
        )
    
    def _merge_thread_target(self):
        """Background merge logic"""
        self.log_message("\n🚀 Starting merge operation...")
        
        self.update_progress(10, "Preparing merge operation...")
        report_a = self.clean_file_path(self.report_a_path.get())
        report_b = self.clean_file_path(self.report_b_path.get())
        output_path = self.clean_file_path(self.output_path.get())
        theme_choice = self.theme_choice.get()
        
        self.update_progress(30, "Reading source reports...")
        
        self.update_progress(50, "Merging reports...")
        
        self.update_progress(75, "Applying theme configuration...")
        
        success = self.merger_engine.merge_reports(
            report_a, report_b, output_path, theme_choice, self.analysis_results
        )
        
        self.update_progress(90, "Finalizing merged report...")
        
        self.update_progress(100, "Merge operation complete!")
        return {'success': success, 'output_path': output_path}
    
    def _handle_merge_complete(self, result):
        """Handle merge completion"""
        if result['success']:
            self.log_message("✅ MERGE COMPLETED SUCCESSFULLY!")
            self.log_message(f"💾 Output: {result['output_path']}")

            # Use Report Merger icon in success dialog
            icon_path = Path(__file__).parent.parent.parent / "assets" / "Tool Icons" / "report merger.svg"
            self.show_info("Merge Complete",
                          f"Merge completed successfully!\n\n💾 Output:\n{result['output_path']}",
                          icon_path=str(icon_path))
        else:
            self.show_error("Merge Failed", "The merge operation failed. Check the log for details.")
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        if self.is_busy:
            if not self.ask_yes_no("Confirm Reset", "An operation is in progress. Stop and reset?"):
                return
        
        # Clear state
        self.report_a_path.set("")
        self.report_b_path.set("")
        self.output_path.set("")
        self.theme_choice.set("report_a")
        self.analysis_results = None
        
        # Reset UI state
        if self.analyze_button:
            self.analyze_button.set_enabled(False)
            self._analyze_button_enabled = False
        if self.merge_button:
            self.merge_button.set_enabled(False)
            self._merge_button_enabled = False
        self._hide_theme_selection()  # This will also adjust window height
        
        # Clear log and show welcome
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)

        # Clear summary panel and restore placeholder
        self._clear_summary_panel()
        self._restore_summary_placeholder()

        self._show_welcome_message()
        self.log_message("✅ Report Merger reset successfully!")
    
    def show_help_dialog(self) -> None:
        """Show help dialog specific to report merger"""
        # Use dynamic colors from theme manager
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Consistent help dialog background for all tools
        help_bg = colors['background']

        def create_help_content(help_window):
            main_frame = tk.Frame(help_window, bg=help_bg, padx=20, pady=20)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Header - centered for middle-out design (consistent with other tools)
            tk.Label(main_frame, text="Report Merger - Help",
                     font=('Segoe UI', 16, 'bold'),
                     bg=help_bg,
                     fg=colors['title_color']).pack(anchor=tk.CENTER, pady=(0, 15))

            # PBIR Requirement Section
            pbir_frame = tk.Frame(main_frame, bg=help_bg)
            pbir_frame.pack(fill=tk.X, pady=(0, 20))

            # Darker orange warning box (#d97706) with white text for better readability
            warning_bg = '#d97706'
            warning_container = tk.Frame(pbir_frame, bg=warning_bg,
                                       padx=15, pady=10, relief='flat', borderwidth=0)
            warning_container.pack(fill=tk.X)

            # White text on darker orange background
            warning_text_color = '#ffffff'
            tk.Label(warning_container, text="⚠️  IMPORTANT DISCLAIMERS & REQUIREMENTS",
                     font=('Segoe UI', 12, 'bold'),
                     bg=warning_bg,
                     fg=warning_text_color).pack(anchor=tk.W)

            warnings = [
                "• This tool ONLY works with PBIP enhanced report format (PBIR) files",
                "• This is NOT officially supported by Microsoft - use at your own discretion",
                "• Look for .pbip files with definition\\ folder (not report.json files)",
                "• Always keep backups of your original reports before merging",
                "• Test thoroughly and validate merged results before production use",
                "• Enable 'Store reports using enhanced metadata format (PBIR)' in Power BI Desktop"
            ]

            for warning in warnings:
                tk.Label(warning_container, text=warning, font=('Segoe UI', 10),
                         bg=warning_bg,
                         fg=warning_text_color).pack(anchor=tk.W, pady=1)

            # Help sections
            help_sections = [
                  ("🎯 What This Tool Does", [
                      "✅ Combines multiple Power BI reports into a single unified report",
                      "✅ Merges pages, bookmarks, and report-level settings intelligently",
                      "✅ Handles theme conflicts with user-selectable resolution",
                      "✅ Preserves all visuals, data sources, and formatting",
                      "✅ Creates a new merged report while keeping originals intact",
                      "✅ Provides detailed analysis and preview before merging"
                  ]),
                ("📁 File Requirements", [
                    "✅ Only .pbip files (enhanced PBIR format) are supported",
                    "✅ Reports must have definition\\ folder structure",
                    "❌ Legacy format with report.json files are NOT supported",
                    "❌ .pbix files are NOT supported"
                ])
            ]

            for title, items in help_sections:
                section_frame = tk.Frame(main_frame, bg=help_bg)
                section_frame.pack(fill=tk.X, pady=(0, 15))

                tk.Label(section_frame, text=title, font=('Segoe UI', 12, 'bold'),
                         fg=colors['title_color'], bg=help_bg).pack(anchor=tk.W)

                for item in items:
                    tk.Label(section_frame, text=f"   {item}", font=('Segoe UI', 10),
                            fg=colors['text_primary'], bg=help_bg).pack(anchor=tk.W, pady=2)

        # Create custom help window for Report Merger (independent of base class)
        help_window = tk.Toplevel(self.main_app.root)
        help_window.withdraw()  # Hide initially to prevent flicker
        help_window.title("Power BI Report Merger - Help")
        help_window.geometry("670x680")  # Custom height for Report Merger
        help_window.resizable(False, False)
        help_window.transient(self.main_app.root)
        help_window.grab_set()
        help_window.configure(bg=help_bg)

        # Set AE favicon icon
        try:
            help_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Create content
        create_help_content(help_window)

        # Bind escape key
        help_window.bind('<Escape>', lambda event: help_window.destroy())

        # Center dialog on parent window after content is created
        help_window.update_idletasks()
        dialog_width = help_window.winfo_reqwidth()
        dialog_height = help_window.winfo_reqheight()
        parent_x = self.main_app.root.winfo_rootx()
        parent_y = self.main_app.root.winfo_rooty()
        parent_width = self.main_app.root.winfo_width()
        parent_height = self.main_app.root.winfo_height()
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        help_window.geometry(f"+{x}+{y}")

        # Set dark/light title bar BEFORE showing window to prevent white flash
        help_window.update()
        self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled
        help_window.focus_force()
