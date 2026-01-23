"""
Sensitivity Scanner UI Tab - Main user interface.

This module provides the SensitivityScannerTab class which implements
the user interface for the Sensitivity Scanner tool.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import io
from pathlib import Path
from typing import Optional, Dict, Any, Set, List, Callable

from datetime import datetime

from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin, ModernScrolledText, RoundedButton, ThemedScrollbar, SquareIconButton, Tooltip, FileInputSection, ThemedMessageBox
from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from core.filter_dropdown import HierarchicalFilterDropdown as BaseFilterDropdown
from tools.sensitivity_scanner.logic.tmdl_scanner import TmdlScanner
from tools.sensitivity_scanner.logic.models import ScanResult, Finding, RiskLevel

# PIL for icon loading (optional)
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# CairoSVG for SVG icon rendering (optional)
try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except (ImportError, OSError):
    CAIROSVG_AVAILABLE = False


class HierarchicalFilterDropdown(BaseFilterDropdown):
    """
    Sensitivity Scanner-specific filter dropdown wrapper.
    Extends the base HierarchicalFilterDropdown with Finding-specific methods.
    """

    def __init__(self, parent: tk.Widget, theme_manager, on_filter_changed: Callable):
        """
        Initialize the filter dropdown with severity-based configuration.

        Args:
            parent: Parent widget to attach to
            theme_manager: Theme manager for colors
            on_filter_changed: Callback when filter selection changes
        """
        # Initialize base class with severity-specific configuration
        super().__init__(
            parent=parent,
            theme_manager=theme_manager,
            on_filter_changed=on_filter_changed,
            group_names=["High", "Medium", "Low"],
            group_colors=None,  # Uses default risk colors
            header_text="Filter Results",
            empty_message="No findings to filter.\nRun a scan first."
        )

    def set_rules(self, rules_by_severity: Dict[str, List[str]]):
        """
        Set the available rules grouped by severity level.

        Args:
            rules_by_severity: Dict mapping severity name to list of rule names
        """
        self.set_items(rules_by_severity)

    def is_finding_visible(self, finding: 'Finding') -> bool:
        """
        Check if a finding should be visible based on current filters.

        Args:
            finding: Finding object with rule_name and severity attributes

        Returns:
            True if the finding should be visible
        """
        return self.is_item_visible(finding.rule_name)


class SensitivityScannerTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """
    Main UI tab for the Sensitivity Scanner tool.
    
    Provides:
    - File input for PBIP selection
    - Scan configuration options
    - Results display with risk levels
    - Export functionality
    - Detailed findings view
    """
    
    def __init__(self, parent, main_app):
        """
        Initialize the Sensitivity Scanner tab.
        
        Args:
            parent: Parent widget (notebook)
            main_app: Main application instance
        """
        super().__init__(parent, main_app, "sensitivity_scanner", "Sensitivity Scanner")
        
        # UI Variables
        self.pbip_path = tk.StringVar()
        self.scan_mode = tk.StringVar(value="full")  # full, tables, roles, expressions
        
        # Scanner components
        self.scanner: Optional[TmdlScanner] = None
        self.scan_result: Optional[ScanResult] = None
        
        # UI Components (will be created in setup_ui)
        self.file_input = None
        self.results_tree = None
        self.details_text = None

        # Button icons for styling
        self._button_icons = {}

        # Button references for theme updates
        self.browse_btn = None
        self.scan_button = None
        self.manage_rules_btn = None
        self.help_btn = None
        # Export and action buttons (replaces ActionButtonBar for 3-button layout)
        self.export_csv_button = None
        self.export_pdf_button = None
        self.reset_button = None

        # Setup UI
        self.setup_ui()
    
    def setup_ui(self) -> None:
        """Setup the complete UI layout."""
        # Load UI icons for buttons and section headers
        icon_names = ["magnifying-glass", "folder", "cogwheel", "bar-chart", "csv-file",
                      "log-file", "reset", "question", "warning", "table", "filter", "pdf"]
        for name in icon_names:
            icon = self._load_icon_for_button(name, size=16)
            if icon:
                self._button_icons[name] = icon

        # IMPORTANT: Create action buttons FIRST with side=BOTTOM
        # This ensures they are always visible regardless of window height
        self._create_action_buttons()

        # Create main content frame for the grid-based sections
        # Note: Don't use Section.TFrame here - that's for INNER frames inside LabelFrames
        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        # Configure grid weights for responsive layout
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(1, weight=3)  # Results section (3x)
        self.content_frame.rowconfigure(2, weight=2)  # Details/Log section gets more (2x)

        # Row 0: Combined File Input + Scan Options Section
        self._create_file_input_section()

        # Row 1: Results Section (expandable)
        self._create_results_section()

        # Row 2: Finding Details and Scan Log (side-by-side, expandable)
        self._create_details_and_log_section()

        # Row 3: Progress Bar
        self.create_progress_bar(self.content_frame)

        # Welcome message
        self._show_welcome_message()

    def _show_welcome_message(self):
        """Show welcome message in log"""
        self.log_message("🔍 Welcome to Sensitivity Scanner!")
        self.log_message("=" * 60)
        self.log_message("This tool scans Power BI semantic models for sensitive data:")
        self.log_message("• Detect PII, financial, and confidential data patterns")
        self.log_message("• Identify sensitive columns, tables, and expressions")
        self.log_message("• Generate compliance reports with custom patterns")
        self.log_message("")
        self.log_message("📂 Start by selecting a PBIP file or connecting to a model")
        self.log_message("⚠️ Note: This is STATIC ANALYSIS of model structure, not data values")

    def _create_file_input_section(self):
        """Create the file input section with scan mode on a separate row below.

        Uses FileInputSection template widget for file input, with custom
        options row (Manage Rules + Scan Mode) and Scan button added manually.
        """
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        content_bg = colors['background']

        # Load radio button icons
        self._radio_on_icon = self._load_icon_for_button('radio-on', 16)
        self._radio_off_icon = self._load_icon_for_button('radio-off', 16)
        self._scan_mode_radio_rows = []  # Track radio rows for theme updates

        # Use FileInputSection for file input (without action button - we add custom content)
        self.file_section = FileInputSection(
            parent=self.content_frame,
            theme_manager=self._theme_manager,
            section_title="PBIP File Source",
            section_icon="Power-BI",
            file_label="PBIP File:",
            file_types=[("Power BI Project Files", "*.pbip")],
            help_command=self.show_help_dialog,
            on_file_selected=lambda path: self._update_scan_button_state()
            # No action_button_text - we add custom content below
        )
        self.file_section.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Store references for compatibility
        self.pbip_path = self.file_section.path_var
        self.browse_btn = self.file_section.browse_button
        self._setup_section_frame = self.file_section.section_frame

        # Setup path cleaning
        self.setup_path_cleaning(self.pbip_path)

        # Get content frame to add custom content
        combined_frame = self.file_section.content_frame
        self._combined_frame = combined_frame  # Store for theme updates

        # ROW 1: Manage Rules + Scan Mode (all on one line, left aligned)
        # Use tk.Frame with explicit bg for proper theme handling
        options_frame = tk.Frame(combined_frame, bg=content_bg)
        options_frame.pack(fill=tk.X, pady=(12, 0))
        self._options_frame = options_frame  # Store for theme updates

        # Theme-aware canvas_bg for proper button corner rounding
        button_canvas_bg = content_bg  # Use content_bg for buttons in content area

        # Manage Rules button - positioned FIRST (left of Scan Mode), auto-sized
        cogwheel_icon = self._button_icons.get('cogwheel')
        self.manage_rules_btn = RoundedButton(
            options_frame,
            text="Manage Rules",
            command=self._open_pattern_manager,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=32,
            radius=6,
            font=('Segoe UI', 10),
            icon=cogwheel_icon,
            canvas_bg=button_canvas_bg
        )
        self.manage_rules_btn.pack(side=tk.LEFT, padx=(0, 20))

        # Scan mode label and radios - no visible container, blend with content
        mode_label = tk.Label(
            options_frame,
            text="Scan Mode:",
            bg=content_bg,
            fg=colors['text_primary'],
            font=('Segoe UI', 9, 'bold')
        )
        mode_label.pack(side=tk.LEFT, padx=(15, 10))
        self._scan_mode_label = mode_label  # Store for theme updates

        modes = [
            ("full", "Full Scan"),
            ("tables", "Tables Only"),
            ("roles", "RLS Only"),
            ("expressions", "Expressions Only")
        ]

        for value, label in modes:
            row_data = self._create_scan_mode_radio(options_frame, label, value, content_bg)
            self._scan_mode_radio_rows.append(row_data)

        # ROW 2: Scan for Issues button - centered on its own row
        scan_button_frame = tk.Frame(combined_frame, bg=content_bg)
        scan_button_frame.pack(fill=tk.X, pady=(15, 0))
        self._scan_button_frame = scan_button_frame  # Store for theme updates

        scan_icon = self._button_icons.get('magnifying-glass')
        self.scan_button = RoundedButton(
            scan_button_frame,
            text="SCAN FOR ISSUES",
            command=self._execute_scan,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            height=38,
            radius=6,
            font=('Segoe UI', 10, 'bold'),
            icon=scan_icon,
            canvas_bg=button_canvas_bg
        )
        self.scan_button.pack()
        # Start disabled until file is selected
        self.scan_button.set_enabled(False)

    def _create_scan_mode_radio(self, parent, text: str, value: str, bg_color: str) -> dict:
        """Create a single SVG radio button for scan mode."""
        colors = self._theme_manager.colors

        # Row frame
        row_frame = tk.Frame(parent, bg=bg_color)
        row_frame.pack(side=tk.LEFT, padx=(0, 12))

        # Radio icon (clickable)
        is_selected = self.scan_mode.get() == value
        icon = self._radio_on_icon if is_selected else self._radio_off_icon

        icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
        if icon:
            icon_label.configure(image=icon)
            icon_label._icon_ref = icon
        else:
            # Fallback to text if icons not available
            icon_label.configure(text="●" if is_selected else "○", font=('Segoe UI', 10))
        icon_label.pack(side=tk.LEFT, padx=(0, 4))

        # Text color - use title_color for selected (blue in dark mode, teal in light mode)
        text_fg = colors['title_color'] if is_selected else colors['text_primary']

        # Text label with underline support
        text_label = tk.Label(
            row_frame, text=text, bg=bg_color, fg=text_fg,
            font=('Segoe UI', 9), cursor='hand2', anchor='w'
        )
        text_label.pack(side=tk.LEFT)

        # Store row data
        row_data = {
            'frame': row_frame,
            'icon_label': icon_label,
            'text_label': text_label,
            'value': value
        }

        # Bind clicks
        def on_click(event=None):
            self.scan_mode.set(value)
            self._update_scan_mode_radios()

        icon_label.bind('<Button-1>', on_click)
        text_label.bind('<Button-1>', on_click)
        row_frame.bind('<Button-1>', on_click)

        # Hover underline effect for all items (including selected)
        def on_enter(event=None):
            text_label.configure(font=('Segoe UI', 9, 'underline'))

        def on_leave(event=None):
            text_label.configure(font=('Segoe UI', 9))

        text_label.bind('<Enter>', on_enter)
        text_label.bind('<Leave>', on_leave)
        icon_label.bind('<Enter>', on_enter)
        icon_label.bind('<Leave>', on_leave)
        row_frame.bind('<Enter>', on_enter)
        row_frame.bind('<Leave>', on_leave)

        return row_data

    def _update_scan_mode_radios(self):
        """Update all scan mode radio buttons when selection changes."""
        colors = self._theme_manager.colors
        # Use white background (colors['background']) for inner content area
        content_bg = colors['background']

        # Update label background
        if hasattr(self, '_scan_mode_label') and self._scan_mode_label:
            try:
                self._scan_mode_label.configure(bg=content_bg, fg=colors['text_primary'])
            except Exception:
                pass

        for row_data in self._scan_mode_radio_rows:
            is_selected = self.scan_mode.get() == row_data['value']

            # Update icon
            icon = self._radio_on_icon if is_selected else self._radio_off_icon
            if icon:
                row_data['icon_label'].configure(image=icon)
                row_data['icon_label']._icon_ref = icon
            else:
                row_data['icon_label'].configure(text="●" if is_selected else "○")

            # Update text color - use title_color for selected (blue in dark mode, teal in light mode)
            text_fg = colors['title_color'] if is_selected else colors['text_primary']
            row_data['text_label'].configure(fg=text_fg)

            # Update backgrounds with content_bg - blend with content area
            row_data['frame'].configure(bg=content_bg)
            row_data['icon_label'].configure(bg=content_bg)
            row_data['text_label'].configure(bg=content_bg)
    
    # _browse_pbip_file is now handled by FileInputSection template widget

    def _create_results_section(self):
        """Create results display section with modern flat treeview styling."""
        colors = self._theme_manager.colors
        section_bg = colors.get('section_bg', colors['background'])
        is_dark = self._theme_manager.is_dark

        results_header = self.create_section_header(self.frame, "Scan Results", "bar-chart")[0]
        outer_frame = ttk.LabelFrame(
            self.content_frame,
            labelwidget=results_header,
            style='Section.TLabelframe',
            padding="12"
        )
        outer_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self._results_section_frame = outer_frame  # Store for reference

        # Inner content frame with section background
        results_frame = ttk.Frame(outer_frame, style='Section.TFrame', padding="15")
        results_frame.pack(fill=tk.BOTH, expand=True)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)  # Allow tree to expand

        # Summary and filter row - use grid for precise alignment with tree below
        top_row = ttk.Frame(results_frame, style='Section.TFrame')
        top_row.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        top_row.columnconfigure(0, weight=1)  # Summary label expands
        top_row.columnconfigure(1, weight=0)  # Filter buttons fixed width

        # Summary label (left side, expands)
        self.summary_label = ttk.Label(
            top_row,
            text="No scan performed yet",
            style='Section.TLabel',
            font=('Segoe UI', 10),
            foreground=colors['text_secondary']
        )
        self.summary_label.grid(row=0, column=0, sticky=tk.W)

        # Hierarchical filter dropdown (right side, aligned with tree edge)
        self._filter_dropdown = HierarchicalFilterDropdown(
            top_row,
            self._theme_manager,
            on_filter_changed=self._apply_filters
        )
        self._filter_dropdown.grid(row=0, column=1, sticky=tk.E)

        # Configure treeview style for flat, modern look
        style = ttk.Style()
        tree_style = "ScanResults.Treeview"

        # Set treeview colors based on theme
        if is_dark:
            tree_bg = '#1e1e2e'
            tree_fg = '#e0e0e0'
            heading_bg = '#2a2a3c'
            heading_fg = '#e0e0e0'
            tree_border = '#3d3d5c'
            selected_bg = '#3d3d5c'
            header_separator = '#0d0d1a'  # Faint column separator for dark mode
        else:
            tree_bg = '#ffffff'
            tree_fg = '#333333'
            heading_bg = '#f0f0f0'
            heading_fg = '#333333'
            tree_border = '#d0d0d0'
            selected_bg = '#3b82f6'  # Blue selection highlight matching Progress Log
            header_separator = '#ffffff'  # Faint column separator for light mode

        # Configure treeview style - flat design
        style.configure(tree_style,
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=tree_bg,
                        borderwidth=0,
                        relief="flat",
                        rowheight=25,
                        lightcolor=tree_bg,
                        darkcolor=tree_bg,
                        bordercolor=tree_bg)

        # Remove treeview frame border
        style.layout(tree_style, [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

        # Heading style with groove relief for column dividers
        style.configure(f"{tree_style}.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        font=('Segoe UI', 9, 'bold'))

        style.map(f"{tree_style}.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])

        style.map(tree_style,
                  background=[('selected', selected_bg)])
        # Don't override foreground on selection - preserve risk colors (red/orange/green)

        # Create container with 1px border for modern look
        # Set highlightcolor=tree_border to prevent focus border color change
        tree_container = tk.Frame(results_frame, bg=tree_border,
                                  highlightbackground=tree_border,
                                  highlightcolor=tree_border,
                                  highlightthickness=1)
        tree_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)
        self._tree_container = tree_container  # Store for theme updates
        self._tree_border_color = tree_border

        # Tree view with flat styling
        self.results_tree = ttk.Treeview(
            tree_container,
            columns=("risk", "file", "location", "pattern", "confidence"),
            show="tree headings",
            height=15,
            style=tree_style
        )
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=0, pady=0)

        # Define columns - all headers centered
        self.results_tree.heading("#0", text="Finding", anchor=tk.CENTER)
        self.results_tree.heading("risk", text="Risk", anchor=tk.CENTER)
        self.results_tree.heading("file", text="File Type", anchor=tk.CENTER)
        self.results_tree.heading("location", text="Location", anchor=tk.CENTER)
        self.results_tree.heading("pattern", text="Pattern", anchor=tk.CENTER)
        self.results_tree.heading("confidence", text="Confidence", anchor=tk.CENTER)

        # Column widths
        self.results_tree.column("#0", width=200, minwidth=150)
        self.results_tree.column("risk", width=80, minwidth=60, anchor=tk.CENTER)
        self.results_tree.column("file", width=120, minwidth=100)
        self.results_tree.column("location", width=200, minwidth=150)
        self.results_tree.column("pattern", width=150, minwidth=100)
        self.results_tree.column("confidence", width=70, minwidth=60, anchor=tk.CENTER)

        # Add ThemedScrollbar with auto-hide
        self._tree_scrollbar = ThemedScrollbar(
            tree_container,
            command=self.results_tree.yview,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_tree.configure(yscrollcommand=self._tree_scrollbar.set)

        # Bind selection event
        self.results_tree.bind('<<TreeviewSelect>>', self._on_finding_selected)
    
    def _create_subsection_labelwidget(self, text: str, icon_name: str) -> ttk.Frame:
        """Create a labelwidget for subsection headers with icon + text (matches Report Merger pattern)."""
        colors = self._theme_manager.colors
        icon = self._button_icons.get(icon_name)
        # Use title_color for subsection headers (blue in dark mode, teal in light mode)
        title_color = colors.get('title_color', colors['primary'])

        # Use ttk.Frame with Section.TFrame style to inherit background properly
        header_frame = ttk.Frame(self.frame, style='Section.TFrame')

        icon_label = None
        if icon:
            # Use ttk.Label for proper style inheritance
            icon_label = ttk.Label(header_frame, image=icon, style='Section.TLabel')
            icon_label.pack(side=tk.LEFT, padx=(0, 4))
            icon_label._icon_ref = icon

        # Use ttk.Label with foreground color - background inherits from style
        text_label = ttk.Label(header_frame, text=text,
                              foreground=title_color, font=('Segoe UI Semibold', 11),
                              style='Section.TLabel')
        text_label.pack(side=tk.LEFT)

        self._section_header_widgets.append((header_frame, icon_label, text_label))
        return header_frame

    def _create_details_and_log_section(self):
        """Create finding details and scan log sections side-by-side with Analysis & Progress header."""
        colors = self._theme_manager.colors

        # Create outer LabelFrame with "Analysis & Progress" header - use bar-chart icon
        analysis_header = self.create_section_header(self.frame, "Analysis & Progress", "bar-chart")[0]
        analysis_outer = ttk.LabelFrame(
            self.content_frame,
            labelwidget=analysis_header,
            style='Section.TLabelframe',
            padding="12"
        )
        analysis_outer.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self._analysis_section_frame = analysis_outer  # Store for reference

        # Inner content frame with section background - padding must match gap between panels
        outer_frame = ttk.Frame(analysis_outer, style='Section.TFrame', padding="15")
        outer_frame.pack(fill=tk.BOTH, expand=True)
        outer_frame.columnconfigure(0, weight=1)
        outer_frame.columnconfigure(1, weight=1)
        outer_frame.rowconfigure(0, weight=1)  # Allow vertical expansion

        # Left: Finding Details container (NO border - matches Report Merger pattern)
        details_container = ttk.Frame(outer_frame, style='Section.TFrame')
        details_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        details_container.columnconfigure(0, weight=1)
        details_container.rowconfigure(0, minsize=30)  # Header row - fixed height
        details_container.rowconfigure(1, weight=1)  # Text area row expands

        # Details header - uses icon label with file.svg
        details_header = self._create_subsection_labelwidget("Finding Details", "file")
        details_header.grid(row=0, column=0, sticky=tk.W, pady=(0, 8), in_=details_container)

        # Details content frame (NO border) - use tk.Frame with internal padding like base class
        # Use section_bg for consistent background
        text_bg = colors['section_bg']
        details_frame = tk.Frame(details_container, bg=text_bg,
                                 highlightthickness=0, padx=8, pady=8)
        details_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(2, 0))
        details_frame.columnconfigure(0, weight=1)
        details_frame.rowconfigure(0, weight=1)
        self._details_frame = details_frame  # Store for theme updates

        # Use ModernScrolledText for themed scrollbar support (NO border on left panel)
        self.details_text = ModernScrolledText(
            details_frame,
            height=6,  # Minimum height, will expand with weight
            font=('Segoe UI', 9),
            wrap=tk.WORD,
            state=tk.DISABLED,
            relief='flat',
            borderwidth=0,
            bg=text_bg,
            fg=colors['text_primary'],
            selectbackground=colors.get('selection_bg', colors['accent']),
            highlightthickness=0,  # NO border on left panel
            padx=5,
            pady=5,
            theme_manager=self._theme_manager,
            auto_hide_scrollbar=True
        )
        self.details_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure text tags for formatting - use theme colors
        self.details_text.tag_configure('header', font=('Segoe UI', 10, 'bold'))
        self.details_text.tag_configure('high_risk', foreground=colors['risk_high'], font=('Segoe UI', 9, 'bold'))
        self.details_text.tag_configure('medium_risk', foreground=colors['risk_medium'], font=('Segoe UI', 9, 'bold'))
        self.details_text.tag_configure('low_risk', foreground=colors['risk_low'], font=('Segoe UI', 9, 'bold'))

        # Right: Progress Log container (HAS border on text widget - matches Report Merger pattern)
        log_outer_container = ttk.Frame(outer_frame, style='Section.TFrame')
        log_outer_container.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_outer_container.columnconfigure(0, weight=1)
        log_outer_container.rowconfigure(0, minsize=30)  # Header row - fixed height
        log_outer_container.rowconfigure(1, weight=1)  # Text area row expands

        # Log header row - contains label and icon buttons
        log_header_frame = ttk.Frame(log_outer_container, style='Section.TFrame')
        log_header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        log_header_frame.columnconfigure(0, weight=1)  # Label takes available space

        # Log header label (left-aligned) - uses icon label with log-file.svg
        log_header = self._create_subsection_labelwidget("Progress Log", "log-file")
        log_header.grid(row=0, column=0, sticky=tk.W, in_=log_header_frame)

        # Load icons for export/clear buttons
        save_icon = self._load_icon_for_button("save", size=14)
        eraser_icon = self._load_icon_for_button("eraser", size=14)
        self._button_icons['save'] = save_icon
        self._button_icons['eraser'] = eraser_icon

        # Icon buttons frame (right-aligned in header - no padding to align with log below)
        icon_buttons_frame = ttk.Frame(log_header_frame, style='Section.TFrame')
        icon_buttons_frame.grid(row=0, column=1, sticky=tk.E)

        # Export button (save icon)
        self._log_export_button = SquareIconButton(
            icon_buttons_frame, icon=save_icon,
            command=lambda: self._export_log(self.log_text),
            tooltip_text="Export Log", size=26, radius=6
        )
        self._log_export_button.pack(side=tk.LEFT, padx=(0, 4))

        # Clear button (eraser icon)
        self._log_clear_button = SquareIconButton(
            icon_buttons_frame, icon=eraser_icon,
            command=lambda: self._clear_log(self.log_text),
            tooltip_text="Clear Log", size=26, radius=6
        )
        self._log_clear_button.pack(side=tk.LEFT)

        # Log content container - no padding here, text widget has internal padx/pady
        log_container = ttk.Frame(log_outer_container, style='Section.TFrame')
        log_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)

        # Log text area - HAS border (right panel has border per design pattern)
        # Use section_bg (#161627 dark, #f5f5f7 light) for consistent background
        log_text_bg = colors['section_bg']
        self.log_text = ModernScrolledText(
            log_container,
            height=6,
            font=('Cascadia Code', 9),
            wrap=tk.WORD,  # Wrap text to avoid horizontal scroll
            state=tk.DISABLED,
            relief='flat',
            borderwidth=0,
            bg=log_text_bg,
            fg=colors['text_primary'],
            selectbackground=colors.get('selection_bg', colors['accent']),
            highlightthickness=1,  # HAS border on right panel
            highlightcolor=colors['border'],
            highlightbackground=colors['border'],
            padx=5,
            pady=5,
            theme_manager=self._theme_manager,
            auto_hide_scrollbar=False  # Always show scrollbar for log
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def _create_action_buttons(self):
        """Create action button section with export options (CSV, PDF, Reset)."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Bottom buttons sit on outer frame which uses outer_bg
        outer_canvas_bg = colors.get('outer_bg', '#161627' if is_dark else '#f5f5f7')

        # Disabled state colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Action frame at bottom - use tk.Frame with explicit bg for proper color matching
        action_frame = tk.Frame(self.frame, bg=outer_canvas_bg)
        action_frame.pack(side=tk.BOTTOM, pady=(15, 0))
        self._action_frame = action_frame  # Store for theme updates

        # Center container for buttons
        button_container = tk.Frame(action_frame, bg=outer_canvas_bg)
        button_container.pack(anchor=tk.CENTER, pady=5)
        self._button_container = button_container  # Store for theme updates

        # EXPORT CSV button (primary style)
        csv_icon = self._button_icons.get('csv-file')
        self.export_csv_button = RoundedButton(
            button_container, text="EXPORT CSV",
            command=self._export_csv,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=csv_icon,
            disabled_bg=disabled_bg, disabled_fg=disabled_fg,
            canvas_bg=outer_canvas_bg
        )
        self.export_csv_button.pack(side=tk.LEFT, padx=(0, 15))
        self.export_csv_button.set_enabled(False)  # Disabled until scan completes

        # EXPORT PDF button (primary style - main export action)
        pdf_icon = self._button_icons.get('pdf')
        self.export_pdf_button = RoundedButton(
            button_container, text="EXPORT PDF",
            command=self._export_html_report,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=pdf_icon,
            disabled_bg=disabled_bg, disabled_fg=disabled_fg,
            canvas_bg=outer_canvas_bg
        )
        self.export_pdf_button.pack(side=tk.LEFT, padx=(0, 15))
        self.export_pdf_button.set_enabled(False)  # Disabled until scan completes

        # RESET ALL button (secondary style)
        reset_icon = self._button_icons.get('reset')
        self.reset_button = RoundedButton(
            button_container, text="RESET ALL",
            command=self.reset_tab,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=38, radius=6, font=('Segoe UI', 10),
            icon=reset_icon,
            canvas_bg=outer_canvas_bg
        )
        self.reset_button.pack(side=tk.LEFT)
    
    def _position_progress_frame(self):
        """Position progress frame for this layout."""
        if self.progress_frame:
            self.progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 5), in_=self.content_frame)

    def _update_scan_button_state(self):
        """Enable Scan button only if a file path is provided"""
        if not hasattr(self, 'scan_button') or not self.scan_button:
            return
        has_file = bool(self.pbip_path.get().strip())
        self.scan_button.set_enabled(has_file)

    def _execute_scan(self):
        """Execute the sensitivity scan."""
        try:
            # Validate input
            pbip_file_path = self.clean_file_path(self.pbip_path.get())

            if not pbip_file_path:
                return  # Button should be disabled, but safety check
            
            self.validate_file_exists(pbip_file_path, "PBIP file")
            self.validate_pbip_file(pbip_file_path, "PBIP file")
            
            # Convert .pbip FILE to PBIP FOLDER path
            # The .pbip file is just metadata - the actual content is in folders
            # Modern PBIP format: .SemanticModel and .Report folders
            pbip_file = Path(pbip_file_path)
            pbip_base_name = pbip_file.stem
            pbip_parent = pbip_file.parent
            
            # Try to find the semantic model folder
            # Modern format: "Name.SemanticModel"
            semantic_model_folder = pbip_parent / f"{pbip_base_name}.SemanticModel"
            
            # Fallback: older format might just be "Name"
            if not semantic_model_folder.exists():
                semantic_model_folder = pbip_parent / pbip_base_name
            
            if not semantic_model_folder.exists():
                self.show_error(
                    "PBIP Folder Not Found", 
                    f"Could not find PBIP semantic model folder:\n"
                    f"Expected: {pbip_parent / (pbip_base_name + '.SemanticModel')}\n"
                    f"Or: {pbip_parent / pbip_base_name}\n\n"
                    f"The .pbip file must have a companion .SemanticModel folder."
                )
                return
            
            # Use the parent folder (which contains .pbip, .SemanticModel, .Report)
            # The PBIPReader expects the parent folder, not the .SemanticModel folder itself
            pbip_folder = pbip_parent
            
            # Clear previous results
            self._clear_results()
            
            # Log scan start
            self.log_message("=" * 60)
            self.log_message("🔍 Starting Sensitivity Scan")
            self.log_message(f"📁 File: {pbip_file.name}")
            self.log_message(f"📂 Folder: {pbip_folder.name}")
            self.log_message(f"⚙️  Mode: {self.scan_mode.get()}")
            self.log_message("=" * 60)
            
            # Run scan in background with FOLDER path
            self.run_in_background(
                target_func=lambda: self._perform_scan(str(pbip_folder)),
                success_callback=self._on_scan_success,
                error_callback=self._on_scan_error
            )
            
        except Exception as e:
            self.log_message(f"❌ Error: {e}")
            self.show_error("Scan Error", str(e))
    
    def _perform_scan(self, pbip_path: str) -> ScanResult:
        """
        Perform the actual scan (background thread).
        
        Args:
            pbip_path: Path to PBIP file
        
        Returns:
            ScanResult with findings
        """
        self.update_progress(0, "Initializing scanner...", show=True, persist=False)
        
        # Initialize scanner
        if self.scanner is None:
            self.scanner = TmdlScanner()
        
        self.update_progress(10, "Finding TMDL files...", show=True, persist=False)
        
        # Perform scan based on mode
        scan_mode = self.scan_mode.get()
        
        if scan_mode == "full":
            self.update_progress(25, "Scanning all TMDL files...", show=True, persist=False)
            result = self.scanner.scan_pbip(pbip_path)
        elif scan_mode == "tables":
            self.update_progress(25, "Scanning table TMDL files...", show=True, persist=False)
            result = self.scanner.scan_by_category(pbip_path, "tables")
        elif scan_mode == "roles":
            self.update_progress(25, "Scanning RLS role files...", show=True, persist=False)
            result = self.scanner.scan_by_category(pbip_path, "roles")
        elif scan_mode == "expressions":
            self.update_progress(25, "Scanning expression files...", show=True, persist=False)
            result = self.scanner.scan_by_category(pbip_path, "expressions")
        else:
            result = self.scanner.scan_pbip(pbip_path)
        
        self.update_progress(90, "Analyzing findings...", show=True, persist=False)
        
        self.update_progress(100, "Scan complete!", show=True, persist=False)
        
        return result
    
    def _on_scan_success(self, result: ScanResult):
        """
        Handle successful scan completion.
        
        Args:
            result: Scan result with findings
        """
        self.scan_result = result
        
        # Log summary
        self.log_message("=" * 60)
        self.log_message("✅ Scan Complete!")
        self.log_message(f"📊 Total Findings: {result.total_findings}")
        self.log_message(f"   🔴 HIGH Risk: {result.high_risk_count}")
        self.log_message(f"   🟡 MEDIUM Risk: {result.medium_risk_count}")
        self.log_message(f"   🟢 LOW Risk: {result.low_risk_count}")
        self.log_message(f"⏱️  Scan Duration: {result.scan_duration_seconds:.2f}s")
        self.log_message(f"📁 Files Scanned: {result.total_files_scanned}")
        self.log_message("=" * 60)
        
        # Update summary label
        summary_text = (
            f"Scan completed: {result.total_findings} findings "
            f"(🔴 {result.high_risk_count} HIGH, "
            f"🟡 {result.medium_risk_count} MEDIUM, "
            f"🟢 {result.low_risk_count} LOW)"
        )
        self.summary_label.config(text=summary_text)
        
        # Populate results tree
        self._populate_results_tree(result)
        
        # Enable export buttons
        self.export_csv_button.set_enabled(True)
        self.export_pdf_button.set_enabled(True)
        
        # Show summary dialog
        if result.has_high_risk:
            # Use sensitivity scanner icon for the warning dialog
            icon_path = Path(__file__).parent.parent.parent.parent / "assets" / "Tool Icons" / "sensivity scanner.svg"
            self.show_warning(
                "High Risk Findings Detected",
                f"Found {result.high_risk_count} HIGH RISK findings!\n\n"
                "Please review and address these before sharing your model.",
                icon_path=str(icon_path)
            )
        elif result.total_findings == 0:
            self.show_info(
                "No Findings",
                "No sensitive data patterns detected in this model.\n\n"
                "Note: This tool uses pattern matching and may not catch everything."
            )
    
    def _on_scan_error(self, error: Exception):
        """
        Handle scan error.
        
        Args:
            error: Exception that occurred
        """
        self.log_message("=" * 60)
        self.log_message(f"❌ Scan Failed: {error}")
        self.log_message("=" * 60)
        self.show_error("Scan Failed", f"An error occurred during scanning:\n\n{error}")
    
    def _populate_results_tree(self, result: ScanResult):
        """
        Populate the results tree with findings.
        
        Args:
            result: Scan result to display
        """
        # Clear existing items
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        if not result.findings:
            return
        
        # Group findings by risk level
        high_findings = result.get_findings_by_risk(RiskLevel.HIGH)
        medium_findings = result.get_findings_by_risk(RiskLevel.MEDIUM)
        low_findings = result.get_findings_by_risk(RiskLevel.LOW)
        
        # Add HIGH risk findings
        if high_findings:
            high_parent = self.results_tree.insert(
                "",
                tk.END,
                text=f"🔴 HIGH RISK ({len(high_findings)})",
                values=("HIGH", "", "", "", ""),
                tags=("high_risk",)
            )
            for finding in high_findings:
                self._add_finding_to_tree(high_parent, finding)
        
        # Add MEDIUM risk findings
        if medium_findings:
            medium_parent = self.results_tree.insert(
                "",
                tk.END,
                text=f"🟡 MEDIUM RISK ({len(medium_findings)})",
                values=("MEDIUM", "", "", "", ""),
                tags=("medium_risk",)
            )
            for finding in medium_findings:
                self._add_finding_to_tree(medium_parent, finding)
        
        # Add LOW risk findings
        if low_findings:
            low_parent = self.results_tree.insert(
                "",
                tk.END,
                text=f"🟢 LOW RISK ({len(low_findings)})",
                values=("LOW", "", "", "", ""),
                tags=("low_risk",)
            )
            for finding in low_findings:
                self._add_finding_to_tree(low_parent, finding)
        
        # Configure tags for coloring - use theme colors
        colors = self._theme_manager.colors
        self.results_tree.tag_configure("high_risk", foreground=colors['risk_high'])
        self.results_tree.tag_configure("medium_risk", foreground=colors['risk_medium'])
        self.results_tree.tag_configure("low_risk", foreground=colors['risk_low'])

        # Populate hierarchical filter dropdown with rules grouped by severity
        rules_by_severity = {
            "High": sorted(set(f.pattern_match.pattern_name for f in high_findings)),
            "Medium": sorted(set(f.pattern_match.pattern_name for f in medium_findings)),
            "Low": sorted(set(f.pattern_match.pattern_name for f in low_findings))
        }
        self._filter_dropdown.set_rules(rules_by_severity)

    def _apply_filters(self, event=None):
        """Apply hierarchical filters from the filter dropdown to the results tree."""
        if not self.scan_result or not self.scan_result.findings:
            return

        # Get filtered findings using the hierarchical dropdown
        filtered_findings = [
            finding for finding in self.scan_result.findings
            if self._filter_dropdown.is_finding_visible(finding)
        ]

        # Save expanded state of parent items before clearing
        expanded_state = {}
        for item in self.results_tree.get_children():
            tags = self.results_tree.item(item, 'tags')
            if tags:
                # Map tag to risk level for state preservation
                tag = tags[0] if isinstance(tags, (list, tuple)) else tags
                is_open = self.results_tree.item(item, 'open')
                expanded_state[tag] = is_open

        # Clear and repopulate tree with filtered results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)

        if not filtered_findings:
            return

        # Group filtered findings by risk level
        high_findings = [f for f in filtered_findings if f.risk_level == RiskLevel.HIGH]
        medium_findings = [f for f in filtered_findings if f.risk_level == RiskLevel.MEDIUM]
        low_findings = [f for f in filtered_findings if f.risk_level == RiskLevel.LOW]

        colors = self._theme_manager.colors

        # Add HIGH risk findings
        if high_findings:
            high_parent = self.results_tree.insert(
                "", tk.END,
                text=f"HIGH RISK ({len(high_findings)})",
                values=("HIGH", "", "", "", ""),
                tags=("high_risk",),
                open=expanded_state.get('high_risk', True)  # Default to expanded
            )
            for finding in high_findings:
                self._add_finding_to_tree(high_parent, finding)

        # Add MEDIUM risk findings
        if medium_findings:
            medium_parent = self.results_tree.insert(
                "", tk.END,
                text=f"MEDIUM RISK ({len(medium_findings)})",
                values=("MEDIUM", "", "", "", ""),
                tags=("medium_risk",),
                open=expanded_state.get('medium_risk', True)  # Default to expanded
            )
            for finding in medium_findings:
                self._add_finding_to_tree(medium_parent, finding)

        # Add LOW risk findings
        if low_findings:
            low_parent = self.results_tree.insert(
                "", tk.END,
                text=f"LOW RISK ({len(low_findings)})",
                values=("LOW", "", "", "", ""),
                tags=("low_risk",),
                open=expanded_state.get('low_risk', True)  # Default to expanded
            )
            for finding in low_findings:
                self._add_finding_to_tree(low_parent, finding)

        # Ensure tags are configured
        self.results_tree.tag_configure("high_risk", foreground=colors['risk_high'])
        self.results_tree.tag_configure("medium_risk", foreground=colors['risk_medium'])
        self.results_tree.tag_configure("low_risk", foreground=colors['risk_low'])

    def _add_finding_to_tree(self, parent, finding: Finding):
        """
        Add a single finding to the tree.
        
        Args:
            parent: Parent tree item
            finding: Finding to add
        """
        # Truncate matched text if too long
        matched_text = finding.pattern_match.matched_text
        if len(matched_text) > 50:
            matched_text = matched_text[:47] + "..."
        
        # Format confidence score
        confidence_pct = f"{finding.confidence_score * 100:.0f}%"
        
        # Add to tree
        self.results_tree.insert(
            parent,
            tk.END,
            text=f"{finding.pattern_match.pattern_name}: {matched_text}",
            values=(
                finding.risk_level.value,
                finding.file_type,
                finding.location_description,
                finding.pattern_match.pattern_name,
                confidence_pct
            ),
            tags=(f"{finding.risk_level.value.lower()}_risk",)
        )
    
    def _on_finding_selected(self, event):
        """Handle finding selection in tree."""
        selection = self.results_tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        item_values = self.results_tree.item(item_id)
        
        # Check if this is a parent node (risk level group)
        if not item_values['values'][1]:  # No file type means parent node
            return
        
        # Find the corresponding finding
        item_text = item_values['text']
        risk_level_str = item_values['values'][0]
        
        # Find matching finding
        finding = self._find_matching_finding(item_text, risk_level_str)
        
        if finding:
            self._display_finding_details(finding)
    
    def _find_matching_finding(self, item_text: str, risk_level_str: str) -> Optional[Finding]:
        """Find a finding matching the tree item."""
        if not self.scan_result:
            return None
        
        risk_level = RiskLevel[risk_level_str]
        findings = self.scan_result.get_findings_by_risk(risk_level)
        
        for finding in findings:
            finding_text = f"{finding.pattern_match.pattern_name}: {finding.pattern_match.matched_text}"
            if finding_text.startswith(item_text[:50]):  # Match truncated text
                return finding
        
        return None
    
    def _display_finding_details(self, finding: Finding):
        """
        Display details for a selected finding.
        
        Args:
            finding: Finding to display
        """
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)
        
        # Risk level header
        risk_tag = f"{finding.risk_level.value.lower()}_risk"
        self.details_text.insert(tk.END, f"Risk Level: {finding.risk_level.value}\n", risk_tag)
        self.details_text.insert(tk.END, f"Confidence: {finding.confidence_score * 100:.0f}%\n\n")
        
        # Pattern info
        self.details_text.insert(tk.END, "Pattern Detected:\n", "header")
        self.details_text.insert(tk.END, f"{finding.pattern_match.pattern_name}\n\n")
        
        # Location
        self.details_text.insert(tk.END, "Location:\n", "header")
        self.details_text.insert(tk.END, f"File: {finding.file_type}\n")
        self.details_text.insert(tk.END, f"{finding.location_description}\n\n")
        
        # Matched text
        self.details_text.insert(tk.END, "Matched Text:\n", "header")
        self.details_text.insert(tk.END, f"{finding.pattern_match.matched_text}\n\n")
        
        # Context
        if finding.pattern_match.context_before or finding.pattern_match.context_after:
            self.details_text.insert(tk.END, "Context:\n", "header")
            if finding.pattern_match.context_before:
                self.details_text.insert(tk.END, f"...{finding.pattern_match.context_before} ")
            self.details_text.insert(tk.END, f"[{finding.pattern_match.matched_text}]")
            if finding.pattern_match.context_after:
                self.details_text.insert(tk.END, f" {finding.pattern_match.context_after}...")
            self.details_text.insert(tk.END, "\n\n")
        
        # Description
        self.details_text.insert(tk.END, "Why This Matters:\n", "header")
        self.details_text.insert(tk.END, f"{finding.description}\n\n")
        
        # Recommendation
        self.details_text.insert(tk.END, "Recommended Action:\n", "header")
        self.details_text.insert(tk.END, f"{finding.recommendation}\n")
        
        self.details_text.config(state=tk.DISABLED)
    
    def _clear_results(self):
        """Clear all results and reset UI."""
        # Clear tree
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Clear details
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)
        self.details_text.config(state=tk.DISABLED)
        
        # Reset summary
        self.summary_label.config(text="Scanning...")
        
        # Disable export buttons
        self.export_csv_button.set_enabled(False)
        self.export_pdf_button.set_enabled(False)
        
        # Clear scan result
        self.scan_result = None
    
    def _export_report(self):
        """Export scan results to a text file."""
        if not self.scan_result:
            self.show_error("No Results", "No scan results to export")
            return
        
        try:
            # Ask for save location
            default_name = f"sensitivity_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            file_path = filedialog.asksaveasfilename(
                title="Export Scan Report",
                defaultextension=".txt",
                initialfile=default_name,
                filetypes=[
                    ("Text Files", "*.txt"),
                    ("All Files", "*.*")
                ]
            )
            
            if not file_path:
                return  # User cancelled
            
            # Generate report
            report = self._generate_text_report()
            
            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(report)
            
            self.log_message(f"✅ Report exported: {Path(file_path).name}")
            self.show_info("Export Successful", f"Report saved to:\n{file_path}")
            
        except Exception as e:
            self.log_message(f"❌ Export failed: {e}")
            self.show_error("Export Failed", f"Could not export report:\n{e}")

    def _export_csv(self):
        """Export scan results to CSV file."""
        if not self.scan_result:
            self.show_error("No Results", "No scan results to export")
            return

        try:
            # Ask for save location
            default_name = f"sensitivity_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            file_path = filedialog.asksaveasfilename(
                title="Export Scan Report (CSV)",
                defaultextension=".csv",
                initialfile=default_name,
                filetypes=[
                    ("CSV Files", "*.csv"),
                    ("All Files", "*.*")
                ]
            )

            if not file_path:
                return  # User cancelled

            # Write CSV
            import csv
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # Header row
                writer.writerow([
                    'Risk Level', 'Pattern', 'Matched Text', 'File Type',
                    'Location', 'Confidence', 'Categories', 'Description', 'Recommendation'
                ])

                # Data rows - all findings sorted by risk
                for finding in self.scan_result.findings:
                    categories = ', '.join([cat.value for cat in finding.categories]) if finding.categories else ''
                    writer.writerow([
                        finding.risk_level.value,
                        finding.pattern_match.pattern_name,
                        finding.pattern_match.matched_text,
                        finding.file_type,
                        finding.location_description,
                        f"{finding.confidence_score * 100:.0f}%",
                        categories,
                        finding.description,
                        finding.recommendation
                    ])

                # Add summary section
                writer.writerow([])
                writer.writerow(['--- SUMMARY ---'])
                writer.writerow(['Scan Date', self.scan_result.scan_time.strftime('%Y-%m-%d %H:%M:%S')])
                writer.writerow(['Source', self.scan_result.source_path])
                writer.writerow(['Files Scanned', self.scan_result.total_files_scanned])
                writer.writerow(['Total Findings', self.scan_result.total_findings])
                writer.writerow(['HIGH Risk', self.scan_result.high_risk_count])
                writer.writerow(['MEDIUM Risk', self.scan_result.medium_risk_count])
                writer.writerow(['LOW Risk', self.scan_result.low_risk_count])

            self.log_message(f"✅ CSV exported: {Path(file_path).name}")
            self.show_info("Export Successful", f"CSV report saved to:\n{file_path}")

        except Exception as e:
            self.log_message(f"❌ CSV export failed: {e}")
            self.show_error("Export Failed", f"Could not export CSV:\n{e}")

    def _export_html_report(self):
        """Export scan results as HTML (printable to PDF)."""
        if not self.scan_result:
            self.show_error("No Results", "No scan results to export")
            return

        try:
            # Ask for save location
            default_name = f"sensitivity_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
            file_path = filedialog.asksaveasfilename(
                title="Save Sensitivity Report (HTML)",
                defaultextension=".html",
                initialfile=default_name,
                filetypes=[
                    ("HTML Files", "*.html"),
                    ("All Files", "*.*")
                ]
            )

            if not file_path:
                return  # User cancelled

            # Generate HTML report
            html_content = self._generate_html_report()

            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # Open in browser for printing
            import webbrowser
            webbrowser.open(f'file:///{file_path}')

            self.log_message(f"✅ HTML report saved: {Path(file_path).name}")
            self.show_info("Export Complete",
                "Report saved as HTML and opened in browser.\n\n"
                "Use your browser's Print function (Ctrl+P) to save as PDF.")

        except Exception as e:
            self.log_message(f"❌ HTML export failed: {e}")
            self.show_error("Export Failed", f"Could not export HTML report:\n{e}")

    def _generate_html_report(self) -> str:
        """Generate a formatted HTML report of scan results."""
        if not self.scan_result:
            return ""

        result = self.scan_result

        # Calculate statistics
        total_findings = result.total_findings
        high_count = result.high_risk_count
        medium_count = result.medium_risk_count
        low_count = result.low_risk_count

        # Build HTML
        html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sensitivity Scan Report</title>
    <style>
        :root {{
            --high-color: #c62828;
            --high-bg: #ffebee;
            --medium-color: #ef6c00;
            --medium-bg: #fff3e0;
            --low-color: #1565c0;
            --low-bg: #e3f2fd;
            --border-color: #e0e0e0;
            --text-primary: #212121;
            --text-secondary: #616161;
            --bg-section: #fafafa;
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: var(--text-primary);
            background: #fff;
            padding: 40px;
            max-width: 1200px;
            margin: 0 auto;
        }}

        h1 {{
            color: #1a237e;
            font-size: 28px;
            margin-bottom: 8px;
            border-bottom: 3px solid #3f51b5;
            padding-bottom: 12px;
        }}

        .metadata {{
            color: var(--text-secondary);
            font-size: 14px;
            margin-bottom: 30px;
        }}

        .metadata span {{
            margin-right: 20px;
        }}

        /* Executive Summary */
        .summary-section {{
            background: var(--bg-section);
            border-radius: 8px;
            padding: 24px;
            margin-bottom: 30px;
        }}

        .summary-section h2 {{
            color: #1a237e;
            font-size: 20px;
            margin-bottom: 20px;
        }}

        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 16px;
        }}

        .summary-card {{
            background: #fff;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
            border: 1px solid var(--border-color);
        }}

        .summary-card.high {{
            border-left: 4px solid var(--high-color);
        }}

        .summary-card.medium {{
            border-left: 4px solid var(--medium-color);
        }}

        .summary-card.low {{
            border-left: 4px solid var(--low-color);
        }}

        .summary-card.total {{
            border-left: 4px solid #1a237e;
        }}

        .summary-card .count {{
            font-size: 36px;
            font-weight: bold;
            margin-bottom: 4px;
        }}

        .summary-card.high .count {{ color: var(--high-color); }}
        .summary-card.medium .count {{ color: var(--medium-color); }}
        .summary-card.low .count {{ color: var(--low-color); }}
        .summary-card.total .count {{ color: #1a237e; }}

        .summary-card .label {{
            color: var(--text-secondary);
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        /* Table of Contents */
        .toc {{
            background: #f5f5f5;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 30px;
        }}

        .toc h3 {{
            margin-bottom: 12px;
            color: #1a237e;
        }}

        .toc ul {{
            list-style: none;
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }}

        .toc a {{
            color: #3f51b5;
            text-decoration: none;
        }}

        .toc a:hover {{
            text-decoration: underline;
        }}

        .toc .badge {{
            display: inline-block;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 12px;
            margin-left: 4px;
        }}

        .toc .badge.high {{ background: var(--high-bg); color: var(--high-color); }}
        .toc .badge.medium {{ background: var(--medium-bg); color: var(--medium-color); }}
        .toc .badge.low {{ background: var(--low-bg); color: var(--low-color); }}

        /* Risk Sections */
        .risk-section {{
            margin-bottom: 40px;
            page-break-inside: avoid;
        }}

        .risk-section h2 {{
            padding: 12px 20px;
            border-radius: 8px 8px 0 0;
            color: #fff;
            font-size: 18px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        .risk-section.high h2 {{ background: var(--high-color); }}
        .risk-section.medium h2 {{ background: var(--medium-color); }}
        .risk-section.low h2 {{ background: var(--low-color); }}

        .findings-container {{
            border: 1px solid var(--border-color);
            border-top: none;
            border-radius: 0 0 8px 8px;
        }}

        .finding-card {{
            padding: 20px;
            border-bottom: 1px solid var(--border-color);
            page-break-inside: avoid;
        }}

        .finding-card:last-child {{
            border-bottom: none;
        }}

        .finding-header {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 12px;
        }}

        .finding-pattern {{
            font-weight: 600;
            font-size: 16px;
            color: var(--text-primary);
        }}

        .finding-confidence {{
            background: #e8eaf6;
            color: #3f51b5;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 500;
        }}

        .finding-details {{
            display: grid;
            grid-template-columns: 120px 1fr;
            gap: 8px;
            font-size: 14px;
        }}

        .finding-details dt {{
            color: var(--text-secondary);
            font-weight: 500;
        }}

        .finding-details dd {{
            color: var(--text-primary);
            word-break: break-word;
        }}

        .matched-text {{
            font-family: 'Cascadia Code', 'Consolas', monospace;
            background: #fff3e0;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 13px;
        }}

        .recommendation {{
            margin-top: 12px;
            padding: 12px;
            background: #e8f5e9;
            border-radius: 6px;
            border-left: 4px solid #4caf50;
        }}

        .recommendation-label {{
            font-weight: 600;
            color: #2e7d32;
            margin-bottom: 4px;
        }}

        /* Footer */
        .footer {{
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid var(--border-color);
            text-align: center;
            color: var(--text-secondary);
            font-size: 13px;
        }}

        .footer a {{
            color: #3f51b5;
        }}

        /* Print styles */
        @media print {{
            body {{
                padding: 20px;
            }}

            .toc {{
                page-break-after: always;
            }}

            .risk-section {{
                page-break-before: always;
            }}

            .risk-section:first-of-type {{
                page-break-before: avoid;
            }}

            .finding-card {{
                page-break-inside: avoid;
            }}

            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
    </style>
</head>
<body>
    <h1>Sensitivity Scan Report</h1>
    <div class="metadata">
        <span><strong>Scan Date:</strong> {result.scan_time.strftime('%Y-%m-%d %H:%M:%S')}</span>
        <span><strong>Source:</strong> {Path(result.source_path).name}</span>
        <span><strong>Duration:</strong> {result.scan_duration_seconds:.2f}s</span>
        <span><strong>Files Scanned:</strong> {result.total_files_scanned}</span>
    </div>

    <div class="summary-section">
        <h2>Executive Summary</h2>
        <div class="summary-grid">
            <div class="summary-card total">
                <div class="count">{total_findings}</div>
                <div class="label">Total Findings</div>
            </div>
            <div class="summary-card high">
                <div class="count">{high_count}</div>
                <div class="label">High Risk</div>
            </div>
            <div class="summary-card medium">
                <div class="count">{medium_count}</div>
                <div class="label">Medium Risk</div>
            </div>
            <div class="summary-card low">
                <div class="count">{low_count}</div>
                <div class="label">Low Risk</div>
            </div>
        </div>
    </div>

    <div class="toc">
        <h3>Table of Contents</h3>
        <ul>
'''
        # Add TOC links
        if high_count > 0:
            html += f'            <li><a href="#high-risk">HIGH Risk Findings</a> <span class="badge high">{high_count}</span></li>\n'
        if medium_count > 0:
            html += f'            <li><a href="#medium-risk">MEDIUM Risk Findings</a> <span class="badge medium">{medium_count}</span></li>\n'
        if low_count > 0:
            html += f'            <li><a href="#low-risk">LOW Risk Findings</a> <span class="badge low">{low_count}</span></li>\n'

        html += '''        </ul>
    </div>
'''

        # HIGH Risk Section
        if high_count > 0:
            html += self._generate_risk_section_html('high', 'HIGH', result.get_findings_by_risk(RiskLevel.HIGH))

        # MEDIUM Risk Section
        if medium_count > 0:
            html += self._generate_risk_section_html('medium', 'MEDIUM', result.get_findings_by_risk(RiskLevel.MEDIUM))

        # LOW Risk Section
        if low_count > 0:
            html += self._generate_risk_section_html('low', 'LOW', result.get_findings_by_risk(RiskLevel.LOW))

        # Footer
        html += '''
    <div class="footer">
        <p>Generated by <strong>AE Multi-Tool - Sensitivity Scanner</strong></p>
        <p>Built by Reid Havens of <a href="https://www.analyticendeavors.com">Analytic Endeavors</a></p>
    </div>
</body>
</html>'''

        return html

    def _generate_risk_section_html(self, risk_class: str, risk_label: str, findings) -> str:
        """Generate HTML for a risk level section."""
        html = f'''
    <div class="risk-section {risk_class}" id="{risk_class}-risk">
        <h2>{risk_label} Risk Findings ({len(list(findings))})</h2>
        <div class="findings-container">
'''
        # Need to iterate again since findings is a generator
        findings_list = list(self.scan_result.get_findings_by_risk(RiskLevel[risk_label]))

        for idx, finding in enumerate(findings_list, 1):
            categories = ', '.join([cat.value for cat in finding.categories]) if finding.categories else 'N/A'
            matched_text = finding.pattern_match.matched_text
            # Escape HTML entities
            matched_text = matched_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

            html += f'''            <div class="finding-card">
                <div class="finding-header">
                    <span class="finding-pattern">#{idx} - {finding.pattern_match.pattern_name}</span>
                    <span class="finding-confidence">{finding.confidence_score * 100:.0f}% Confidence</span>
                </div>
                <dl class="finding-details">
                    <dt>Matched Text:</dt>
                    <dd><code class="matched-text">{matched_text}</code></dd>
                    <dt>File Type:</dt>
                    <dd>{finding.file_type}</dd>
                    <dt>Location:</dt>
                    <dd>{finding.location_description}</dd>
                    <dt>Categories:</dt>
                    <dd>{categories}</dd>
                    <dt>Description:</dt>
                    <dd>{finding.description}</dd>
                </dl>
'''
            if finding.recommendation:
                html += f'''                <div class="recommendation">
                    <div class="recommendation-label">Recommendation</div>
                    <div>{finding.recommendation}</div>
                </div>
'''
            html += '''            </div>
'''

        html += '''        </div>
    </div>
'''
        return html

    def _generate_text_report(self) -> str:
        """Generate a formatted text report of scan results."""
        if not self.scan_result:
            return ""
        
        lines = []
        lines.append("=" * 80)
        lines.append("SENSITIVITY SCAN REPORT")
        lines.append("=" * 80)
        lines.append("")
        lines.append(f"Scan Date: {self.scan_result.scan_time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Source: {self.scan_result.source_path}")
        lines.append(f"Duration: {self.scan_result.scan_duration_seconds:.2f} seconds")
        lines.append(f"Files Scanned: {self.scan_result.total_files_scanned}")
        lines.append("")
        lines.append("-" * 80)
        lines.append("SUMMARY")
        lines.append("-" * 80)
        lines.append(f"Total Findings: {self.scan_result.total_findings}")
        lines.append(f"  HIGH Risk:   {self.scan_result.high_risk_count}")
        lines.append(f"  MEDIUM Risk: {self.scan_result.medium_risk_count}")
        lines.append(f"  LOW Risk:    {self.scan_result.low_risk_count}")
        lines.append("")
        
        # High risk findings
        if self.scan_result.high_risk_count > 0:
            lines.append("=" * 80)
            lines.append("HIGH RISK FINDINGS")
            lines.append("=" * 80)
            lines.append("")
            for idx, finding in enumerate(self.scan_result.get_findings_by_risk(RiskLevel.HIGH), 1):
                lines.extend(self._format_finding_for_report(idx, finding))
        
        # Medium risk findings
        if self.scan_result.medium_risk_count > 0:
            lines.append("=" * 80)
            lines.append("MEDIUM RISK FINDINGS")
            lines.append("=" * 80)
            lines.append("")
            for idx, finding in enumerate(self.scan_result.get_findings_by_risk(RiskLevel.MEDIUM), 1):
                lines.extend(self._format_finding_for_report(idx, finding))
        
        # Low risk findings
        if self.scan_result.low_risk_count > 0:
            lines.append("=" * 80)
            lines.append("LOW RISK FINDINGS")
            lines.append("=" * 80)
            lines.append("")
            for idx, finding in enumerate(self.scan_result.get_findings_by_risk(RiskLevel.LOW), 1):
                lines.extend(self._format_finding_for_report(idx, finding))
        
        lines.append("=" * 80)
        lines.append("END OF REPORT")
        lines.append("=" * 80)
        lines.append("")
        lines.append("Generated by AE Multi-Tool - Sensitivity Scanner")
        lines.append("Built by Reid Havens of Analytic Endeavors")
        lines.append("https://www.analyticendeavors.com")
        
        return "\n".join(lines)
    
    def _format_finding_for_report(self, idx: int, finding: Finding) -> list:
        """Format a single finding for the text report."""
        lines = []
        lines.append(f"Finding #{idx}")
        lines.append("-" * 40)
        lines.append(f"Pattern: {finding.pattern_match.pattern_name}")
        lines.append(f"Risk Level: {finding.risk_level.value}")
        lines.append(f"Confidence: {finding.confidence_score * 100:.0f}%")
        lines.append(f"File Type: {finding.file_type}")
        lines.append(f"Location: {finding.location_description}")
        lines.append(f"Matched Text: {finding.pattern_match.matched_text}")
        
        if finding.pattern_match.context_before or finding.pattern_match.context_after:
            context = ""
            if finding.pattern_match.context_before:
                context += f"...{finding.pattern_match.context_before} "
            context += f"[{finding.pattern_match.matched_text}]"
            if finding.pattern_match.context_after:
                context += f" {finding.pattern_match.context_after}..."
            lines.append(f"Context: {context}")
        
        lines.append(f"Description: {finding.description}")
        lines.append(f"Recommendation: {finding.recommendation}")
        lines.append("")
        
        return lines
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state."""
        # Clear inputs
        self.pbip_path.set("")
        self.scan_mode.set("full")
        
        # Clear results
        self._clear_results()
        self.summary_label.config(text="No scan performed yet")
        
        # Clear log
        self._clear_log(self.log_text)
        
        # Reset scanner
        self.scanner = None
        self.scan_result = None

        # Welcome message
        self._show_welcome_message()
    
    def _open_pattern_manager(self):
        """Open the Pattern Manager window."""
        try:
            from tools.sensitivity_scanner.ui.pattern_manager import PatternManager
            
            def on_patterns_updated():
                """Callback when patterns are updated."""
                self.log_message("Patterns updated! Please run a new scan to apply changes.")
                # Reinitialize scanner with new patterns
                self.scanner = None

            # Open pattern manager
            PatternManager(self.frame, self.scanner, on_patterns_updated)

        except Exception as e:
            self.log_message(f"Failed to open Pattern Manager: {e}")
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Error", f"Failed to open Pattern Manager:\n{e}")

    def on_theme_changed(self, theme: str) -> None:
        """Handle theme changes - update all themed widgets."""
        # Let base class handle tk-based section headers
        super().on_theme_changed(theme)

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        section_bg = colors.get('section_bg', colors['background'])
        outer_canvas_bg = colors.get('outer_bg', '#161627' if is_dark else '#f5f5f7')  # Bottom buttons sit on outer frame

        # Update action button frames
        if hasattr(self, '_action_frame') and self._action_frame:
            self._action_frame.configure(bg=outer_canvas_bg)
        if hasattr(self, '_button_container') and self._button_container:
            self._button_container.configure(bg=outer_canvas_bg)

        # Use title_color for section headers (blue in dark mode, teal in light mode)
        title_color = colors.get('title_color', colors['primary'])

        # Handle ttk-based subsection headers (foreground only - background from style)
        for header_frame, icon_label, text_label in self._section_header_widgets:
            try:
                # Skip tk.Frame (already handled by base class)
                header_frame.configure(bg=section_bg)
            except tk.TclError:
                # ttk.Frame/Label (subsection headers) - only update foreground
                try:
                    text_label.configure(foreground=title_color)
                except Exception:
                    pass

        # Inner content frames use white background (colors['background']), not section_bg
        content_bg = colors['background']

        # Update combined frame (inner content frame)
        if hasattr(self, '_combined_frame') and self._combined_frame:
            try:
                self._combined_frame.configure(bg=content_bg)
            except Exception:
                pass

        # Update options frame
        if hasattr(self, '_options_frame') and self._options_frame:
            try:
                self._options_frame.configure(bg=content_bg)
            except Exception:
                pass

        # Update scan button frame
        if hasattr(self, '_scan_button_frame') and self._scan_button_frame:
            try:
                self._scan_button_frame.configure(bg=content_bg)
            except Exception:
                pass

        # Reload radio icons for new theme and update scan mode radios
        # Note: _update_scan_mode_radios now handles container, label, and radio backgrounds
        if hasattr(self, '_scan_mode_radio_rows') and self._scan_mode_radio_rows:
            self._radio_on_icon = self._load_icon_for_button('radio-on', 16)
            self._radio_off_icon = self._load_icon_for_button('radio-off', 16)
            self._update_scan_mode_radios()

        # Update treeview risk tags and styling
        if hasattr(self, 'results_tree') and self.results_tree:
            self.results_tree.tag_configure("high_risk", foreground=colors['risk_high'])
            self.results_tree.tag_configure("medium_risk", foreground=colors['risk_medium'])
            self.results_tree.tag_configure("low_risk", foreground=colors['risk_low'])

            # Update treeview style for new theme - use centralized colors
            tree_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
            tree_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
            heading_bg = colors.get('section_bg', '#2a2a3c' if is_dark else '#f0f0f0')
            heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
            tree_border = colors.get('border', '#3d3d5c' if is_dark else '#d0d0d0')
            selected_bg = '#3d3d5c' if is_dark else '#3b82f6'  # Blue selection highlight
            header_separator = colors.get('background', '#0d0d1a' if is_dark else '#ffffff')

            style = ttk.Style()
            tree_style = "ScanResults.Treeview"
            style.configure(tree_style,
                            background=tree_bg,
                            foreground=tree_fg,
                            fieldbackground=tree_bg,
                            borderwidth=0,
                            relief="flat",
                            rowheight=25,
                            lightcolor=tree_bg,
                            darkcolor=tree_bg,
                            bordercolor=tree_bg)
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
                      background=[('selected', selected_bg)])
            # Don't override foreground on selection - preserve risk colors

            # Re-apply style to force tree to pick up changes
            self.results_tree.configure(style=tree_style)
            self.results_tree.update_idletasks()

            # Update tree container border (include highlightcolor to prevent focus border change)
            if hasattr(self, '_tree_container') and self._tree_container:
                try:
                    self._tree_container.configure(bg=tree_border, highlightbackground=tree_border,
                                                   highlightcolor=tree_border)
                except Exception:
                    pass

        # Update details frame background - use section_bg (#161627 dark, #f5f5f7 light)
        text_bg = colors['section_bg']
        if hasattr(self, '_details_frame') and self._details_frame:
            try:
                self._details_frame.configure(bg=text_bg)
            except Exception:
                pass

        # Update details text background and tags
        if hasattr(self, 'details_text') and self.details_text:
            try:
                self.details_text.configure(bg=text_bg, fg=colors['text_primary'])
                # Update outer Frame wrapper for ModernScrolledText (tk is already imported at module level)
                tk.Frame.configure(self.details_text, bg=text_bg)
            except Exception:
                pass
            self.details_text.tag_configure('high_risk', foreground=colors['risk_high'])
            self.details_text.tag_configure('medium_risk', foreground=colors['risk_medium'])
            self.details_text.tag_configure('low_risk', foreground=colors['risk_low'])

        # Update log text background - use section_bg (#161627 dark, #f5f5f7 light)
        if hasattr(self, 'log_text') and self.log_text:
            try:
                self.log_text.configure(
                    bg=text_bg,
                    fg=colors['text_primary'],
                    highlightcolor=colors['border'],
                    highlightbackground=colors['border']
                )
                # Update outer Frame wrapper for ModernScrolledText (tk is already imported at module level)
                tk.Frame.configure(self.log_text, bg=text_bg,
                                   highlightcolor=colors['border'],
                                   highlightbackground=colors['border'])
            except Exception:
                pass

        # Update radiobutton style
        style = ttk.Style()
        style.configure('Clean.TRadiobutton',
                       background=colors['section_bg'],
                       foreground=colors['text_primary'])

        # Update summary label
        if hasattr(self, 'summary_label') and self.summary_label:
            self.summary_label.configure(foreground=colors['text_secondary'])

        # Update hierarchical filter dropdown
        if hasattr(self, '_filter_dropdown') and self._filter_dropdown:
            self._filter_dropdown.on_theme_changed()

        # Update RoundedButtons with new colors and canvas backgrounds
        # Content area buttons (scan, manage rules) - use content_bg (colors['background'])
        # Note: browse_btn is handled by FileInputSection's internal theme updates
        content_buttons = [
            (self.scan_button, 'button_primary', 'button_primary_hover', 'button_primary_pressed', '#ffffff', content_bg),
            (self.manage_rules_btn, 'button_secondary', 'button_secondary_hover', 'button_secondary_pressed', colors['text_primary'], content_bg),
        ]

        # Get disabled colors for proper light/dark mode updates
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        for btn, bg_key, hover_key, pressed_key, fg, canvas_bg in content_buttons:
            if btn:
                try:
                    btn.update_colors(
                        bg=colors[bg_key],
                        hover_bg=colors[hover_key],
                        pressed_bg=colors[pressed_key],
                        fg=fg,
                        disabled_bg=disabled_bg,
                        disabled_fg=disabled_fg
                    )
                    btn.update_canvas_bg(canvas_bg)
                except Exception:
                    pass

        # Bottom buttons (export CSV, export PDF, reset, help)
        bottom_buttons = [
            (self.export_csv_button, 'button_primary', 'button_primary_hover', 'button_primary_pressed', '#ffffff'),
            (self.export_pdf_button, 'button_primary', 'button_primary_hover', 'button_primary_pressed', '#ffffff'),
            (self.reset_button, 'button_secondary', 'button_secondary_hover', 'button_secondary_pressed', colors['text_primary']),
            (self.help_btn, 'button_secondary', 'button_secondary_hover', 'button_secondary_pressed', colors['text_primary']),
        ]

        for btn, bg_key, hover_key, pressed_key, fg in bottom_buttons:
            if btn:
                try:
                    btn.update_colors(
                        bg=colors[bg_key],
                        hover_bg=colors[hover_key],
                        pressed_bg=colors[pressed_key],
                        fg=fg,
                        disabled_bg=disabled_bg,
                        disabled_fg=disabled_fg
                    )
                    btn.update_canvas_bg(outer_canvas_bg)
                except Exception:
                    pass

    def _set_window_icon(self, window: tk.Toplevel) -> None:
        """Set the sensitivity scanner SVG as a window icon.

        Args:
            window: The Toplevel window to set the icon on
        """
        try:
            # Get base path for assets
            base_path = Path(__file__).parent.parent.parent.parent

            # Load sensitivity scanner SVG and convert to icon
            icon_path = base_path / "assets" / "Tool Icons" / "sensivity scanner.svg"
            if icon_path.exists() and PIL_AVAILABLE and CAIROSVG_AVAILABLE:
                # Convert SVG to PNG at multiple sizes for best quality
                png_data = cairosvg.svg2png(url=str(icon_path), output_width=128, output_height=128)
                img = Image.open(io.BytesIO(png_data))

                # Create multiple sizes for the icon
                icon_sizes = [img.resize((size, size), Image.Resampling.LANCZOS)
                             for size in (16, 32, 48, 64)]
                photo_icons = [ImageTk.PhotoImage(icon) for icon in icon_sizes]

                # Keep references to prevent garbage collection
                if not hasattr(self, '_window_icons'):
                    self._window_icons = []
                self._window_icons.extend(photo_icons)

                # Set as window icon
                window.iconphoto(True, *photo_icons)
            else:
                # Fallback to AE favicon
                favicon_path = base_path / "assets" / "favicon.ico"
                if favicon_path.exists():
                    window.iconbitmap(str(favicon_path))
        except Exception:
            # Fallback to AE favicon on any error
            try:
                base_path = Path(__file__).parent.parent.parent.parent
                favicon_path = base_path / "assets" / "favicon.ico"
                if favicon_path.exists():
                    window.iconbitmap(str(favicon_path))
            except Exception:
                pass

    def show_help_dialog(self) -> None:
        """Show the help dialog with prominent disclaimer - modern styling with two-column layout."""
        from core.ui_base import RoundedButton
        from tools.sensitivity_scanner.sensitivity_scanner_tool import SensitivityScannerTool
        tool = SensitivityScannerTool()
        help_content = tool.get_help_content()

        colors = self._theme_manager.colors

        # Get correct parent window
        parent_window = self.frame.winfo_toplevel()

        # Consistent help dialog background for all tools
        is_dark = self._theme_manager.is_dark
        help_bg = colors['background']

        # Create help window
        help_window = tk.Toplevel(parent_window)
        help_window.withdraw()  # Hide until styled
        help_window.title(help_content['title'])
        help_window.geometry("850x650")
        help_window.resizable(False, False)
        help_window.transient(parent_window)
        help_window.grab_set()
        help_window.configure(bg=help_bg)

        # Set AE favicon icon (consistent with other tools)
        try:
            help_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Set dark title bar if dark mode
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, is_dark)

        # Main container - use consistent help_bg
        main_frame = tk.Frame(help_window, bg=help_bg, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header - centered for middle-out design
        tk.Label(main_frame, text="Sensitivity Scanner - Help",
                font=('Segoe UI', 16, 'bold'),
                bg=help_bg, fg=colors['title_color']).pack(anchor=tk.CENTER, pady=(0, 15))

        # Orange warning section - find disclaimer section
        warning_frame = tk.Frame(main_frame, bg=help_bg)
        warning_frame.pack(fill=tk.X, pady=(0, 15))

        warning_bg = '#d97706'
        warning_container = tk.Frame(warning_frame, bg=warning_bg,
                                   padx=15, pady=10, relief='flat', borderwidth=0)
        warning_container.pack(fill=tk.X)

        tk.Label(warning_container, text="IMPORTANT DISCLAIMERS & REQUIREMENTS",
                font=('Segoe UI', 12, 'bold'),
                bg=warning_bg, fg='#ffffff').pack(anchor=tk.W)

        # Find disclaimer section from help content
        disclaimer_items = []
        other_sections = []
        for section in help_content['sections']:
            if "IMPORTANT DISCLAIMERS" in section['title']:
                disclaimer_items = section['items']
            else:
                other_sections.append(section)

        # Two-column layout for warning items
        warning_cols = tk.Frame(warning_container, bg=warning_bg)
        warning_cols.pack(fill=tk.X, pady=(5, 0))
        warning_cols.columnconfigure(0, weight=1)
        warning_cols.columnconfigure(1, weight=1)

        mid = (len(disclaimer_items) + 1) // 2
        left_warnings = disclaimer_items[:mid]
        right_warnings = disclaimer_items[mid:]

        warning_left = tk.Frame(warning_cols, bg=warning_bg)
        warning_left.grid(row=0, column=0, sticky='nw')
        for item in left_warnings:
            tk.Label(warning_left, text=item if item.startswith("•") else f"• {item}",
                    font=('Segoe UI', 10), bg=warning_bg, fg='#ffffff',
                    wraplength=380, justify=tk.LEFT).pack(anchor=tk.W, pady=1)

        warning_right = tk.Frame(warning_cols, bg=warning_bg)
        warning_right.grid(row=0, column=1, sticky='nw')
        for item in right_warnings:
            tk.Label(warning_right, text=item if item.startswith("•") else f"• {item}",
                    font=('Segoe UI', 10), bg=warning_bg, fg='#ffffff',
                    wraplength=380, justify=tk.LEFT).pack(anchor=tk.W, pady=1)

        # Two-column layout for remaining sections
        columns_frame = tk.Frame(main_frame, bg=help_bg)
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)

        # Split sections into two columns
        mid_point = (len(other_sections) + 1) // 2
        left_sections = other_sections[:mid_point]
        right_sections = other_sections[mid_point:]

        # Left column
        left_col = tk.Frame(columns_frame, bg=help_bg)
        left_col.grid(row=0, column=0, sticky='nsew', padx=(0, 10))

        for section in left_sections:
            tk.Label(left_col, text=section['title'],
                    font=('Segoe UI', 12, 'bold'),
                    bg=help_bg, fg=colors['title_color']).pack(anchor=tk.W, pady=(10, 5))
            for item in section['items']:
                tk.Label(left_col, text=item,
                        font=('Segoe UI', 10),
                        bg=help_bg, fg=colors['text_primary'],
                        wraplength=380, justify=tk.LEFT).pack(anchor=tk.W, padx=(10, 0))

        # Right column
        right_col = tk.Frame(columns_frame, bg=help_bg)
        right_col.grid(row=0, column=1, sticky='nsew', padx=(10, 0))

        for section in right_sections:
            tk.Label(right_col, text=section['title'],
                    font=('Segoe UI', 12, 'bold'),
                    bg=help_bg, fg=colors['title_color']).pack(anchor=tk.W, pady=(10, 5))
            for item in section['items']:
                tk.Label(right_col, text=item,
                        font=('Segoe UI', 10),
                        bg=help_bg, fg=colors['text_primary'],
                        wraplength=380, justify=tk.LEFT).pack(anchor=tk.W, padx=(10, 0))

        help_window.bind('<Escape>', lambda e: help_window.destroy())

        # Center and show dialog
        help_window.update_idletasks()
        x = parent_window.winfo_rootx() + (parent_window.winfo_width() - 850) // 2
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 650) // 2
        help_window.geometry(f"850x650+{x}+{y}")
        help_window.deiconify()
