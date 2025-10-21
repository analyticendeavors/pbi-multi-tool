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

from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin
from core.constants import AppConstants
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
        self.pbip_path_var = tk.StringVar()
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
        
        # Setup UI
        self.setup_ui()
        self.setup_path_cleaning(self.pbip_path_var)
        self._show_welcome_message()
    
    def setup_ui(self) -> None:
        """Setup the Table Column Widths UI"""
        # File input section
        self._setup_file_input_section()
        
        # Main content in two columns
        main_content = ttk.Frame(self.frame)
        main_content.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        main_content.columnconfigure(0, weight=1)
        main_content.columnconfigure(1, weight=1)
        
        # LEFT COLUMN: Scanning + Configuration
        left_column = ttk.Frame(main_content)
        left_column.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Visual scanning section (left column)
        self._setup_scanning_section(left_column)
        
        # Configuration section (left column)
        self._setup_configuration_section(left_column)
        
        # Action buttons (left column)
        self._setup_action_buttons(left_column)
        
        # RIGHT COLUMN: Visual Selection + Log
        self.right_column = ttk.Frame(main_content)
        self.right_column.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        
        # Visual selection section (right column)
        self._setup_visual_selection_section(self.right_column)
        
        # Progress bar
        progress_info = self.create_progress_bar(self.right_column)
        
        # Log section (right column) - reduced height for better fit
        log_info = self.create_log_section(self.right_column, "📊 TABLE COLUMN WIDTHS LOG")
        # Customize log height for this tool
        log_info['text_widget'].config(height=10)  # Reduced from default 12
        log_info['frame'].pack(fill=tk.BOTH, expand=True, pady=(15, 0))
    
    def _setup_file_input_section(self):
        """Setup file input section"""
        # Create the section frame manually to have more control
        section_frame = ttk.LabelFrame(self.frame, text="📁 PBIP REPORT FILE SELECTION", 
                                     style='Section.TLabelframe', padding="20")
        section_frame.pack(fill=tk.X, pady=(0, 20))
        
        content_frame = ttk.Frame(section_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # LEFT: Guide text
        guide_frame = ttk.Frame(content_frame)
        guide_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 35))
        
        ttk.Label(guide_frame, text="📁 Select PBIP Report File:", 
                 font=('Segoe UI', 10, 'bold'), 
                 foreground=AppConstants.COLORS['info']).pack(anchor=tk.W)
        
        guide_text = [
            "1. Choose a .pbip file (enhanced report format)",
            "2. Ensure corresponding .Report directory exists",
            "3. File will be scanned for Table and Matrix visuals"
        ]
        
        for text in guide_text:
            ttk.Label(guide_frame, text=f"   {text}", font=('Segoe UI', 9),
                     foreground=AppConstants.COLORS['text_secondary'], 
                     wraplength=300).pack(anchor=tk.W, pady=1)
        
        # Add warning text at bottom in orange/red and italicized
        warning_label = ttk.Label(guide_frame, text="⚠️ Requires PBIP format with TMDL files",
                                font=('Segoe UI', 9, 'italic'),
                                foreground='#d97706')
        warning_label.pack(anchor=tk.W, pady=(5, 0))
        
        # RIGHT: File input
        input_frame = ttk.Frame(content_frame)
        input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        input_frame.columnconfigure(1, weight=1)
        
        # File input row
        ttk.Label(input_frame, text="File Path:").grid(row=0, column=0, sticky=tk.W, pady=8)
        
        self.file_entry = ttk.Entry(input_frame, textvariable=self.pbip_path_var, width=80)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        
        # Browse button connected to our method
        browse_button = ttk.Button(input_frame, text="📂 Browse",
                                  command=self._browse_file_custom)
        browse_button.grid(row=0, column=2, pady=8)
    
    def _browse_file_custom(self):
        """Custom browse file method that uses our path variable"""
        file_path = filedialog.askopenfilename(
            title="Select PBIP File",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.pbip_path_var.set(file_path)
    
    def _setup_scanning_section(self, parent):
        """Setup visual scanning section"""
        scan_frame = ttk.LabelFrame(parent, text="🔍 VISUAL SCANNING", 
                                   style='Section.TLabelframe', padding="15")
        scan_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Instructions
        ttk.Label(scan_frame, text="📋 Scan Report for Tables and Matrices", 
                 font=('Segoe UI', 10, 'bold'),
                 foreground=AppConstants.COLORS['info']).pack(anchor=tk.W, pady=(0, 5))
        
        # Scan button and summary in horizontal layout
        button_frame = ttk.Frame(scan_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.scan_button = ttk.Button(button_frame, text="🔍 SCAN VISUALS",
                                     command=self._scan_visuals,
                                     style='Action.TButton')
        self.scan_button.pack(side=tk.LEFT)
        
        # Summary display
        self.summary_label = ttk.Label(button_frame, text="No visuals scanned yet",
                                      font=('Segoe UI', 9),
                                      foreground=AppConstants.COLORS['text_secondary'])
        self.summary_label.pack(side=tk.LEFT, padx=(15, 0))
    
    def _setup_configuration_section(self, parent):
        """Setup width configuration section"""
        config_frame = ttk.LabelFrame(parent, text="⚙️ COLUMN WIDTH CONFIGURATION", 
                                     style='Section.TLabelframe', padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Configuration mode selector
        mode_frame = ttk.Frame(config_frame)
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(mode_frame, text="📝 Configuration Mode:", 
                 font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT)
        
        # Configuration mode label - will update dynamically
        self.config_mode_label = ttk.Label(mode_frame, text="Global Settings (All Visuals)", 
                                          font=('Segoe UI', 9),
                                          foreground=AppConstants.COLORS['info'])
        self.config_mode_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Categorical and measure settings in horizontal layout
        columns_frame = ttk.Frame(config_frame)
        columns_frame.pack(fill=tk.X)
        
        # Categorical columns (left)
        cat_frame = ttk.LabelFrame(columns_frame, text="📊 Categorical Columns", padding="10")
        cat_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self._setup_width_controls_with_custom(cat_frame, self.categorical_preset_var, 'categorical')
        
        # Measure columns (right)
        measure_frame = ttk.LabelFrame(columns_frame, text="📈 Measure Columns", padding="10")
        measure_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(5, 0))
        
        self._setup_width_controls_with_custom(measure_frame, self.measure_preset_var, 'measure')
    
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
        
        # Custom width input - single line layout
        custom_frame = ttk.Frame(parent)
        custom_frame.pack(anchor=tk.W, pady=(5, 0))
        
        # Custom label and entry on same line
        ttk.Label(custom_frame, text="Custom (px):", font=('Segoe UI', 9)).pack(side=tk.LEFT)
        
        # Create custom width variable
        if field_type == 'categorical':
            custom_var = tk.StringVar(value="105")
            self.categorical_custom_var = custom_var
        else:
            custom_var = tk.StringVar(value="95")
            self.measure_custom_var = custom_var
        
        # Entry box with validation - using tk.Entry for background color control
        custom_entry = tk.Entry(custom_frame, textvariable=custom_var, width=6, font=('Segoe UI', 9),
                               relief='solid', borderwidth=1, bg='white')
        custom_entry.pack(side=tk.LEFT, padx=(8, 0))
        
        # Delayed validation to prevent false positives while typing
        validation_timer = None
        
        def delayed_validate():
            """Actual validation function called after delay"""
            try:
                value = int(custom_var.get())
                if value < 50:  # Below minimum width
                    # Change background to light red
                    custom_entry.config(bg='#ffebee')
                    # Show warning in log
                    field_name = "categorical" if field_type == 'categorical' else "measure"
                    self.log_message(f"⚠️ Warning: {field_name} custom width {value}px is below minimum (50px). Will be increased to 50px.")
                else:
                    # Reset to normal background
                    custom_entry.config(bg='white')
            except ValueError:
                # Reset to normal background for non-numeric values
                custom_entry.config(bg='white')
        
        def validate_custom_width(*args):
            """Validation with delay - resets timer on each keystroke"""
            nonlocal validation_timer
            # Cancel any pending validation
            if validation_timer:
                self.frame.after_cancel(validation_timer)
            # Schedule new validation after 800ms delay
            validation_timer = self.frame.after(800, delayed_validate)
        
        custom_var.trace('w', validate_custom_width)
    
    def _setup_width_controls_with_custom_for_popup(self, parent: ttk.Widget, preset_var: tk.StringVar, field_type: str, visual_id: str):
        """Setup width control widgets with custom input field for popup dialogs"""
        
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
        
        # Custom width input - single line layout
        custom_frame = ttk.Frame(parent)
        custom_frame.pack(anchor=tk.W, pady=(5, 0))
        
        # Custom label and entry on same line
        ttk.Label(custom_frame, text="Custom (px):", font=('Segoe UI', 9)).pack(side=tk.LEFT)
        
        # Get or create per-visual custom variables
        if visual_id not in self.visual_config_vars:
            self._create_per_visual_config_vars(visual_id)
        
        # Get/create custom variable for this visual
        if f'{field_type}_custom' not in self.visual_config_vars[visual_id]:
            default_value = "105" if field_type == 'categorical' else "95"
            self.visual_config_vars[visual_id][f'{field_type}_custom'] = tk.StringVar(value=default_value)
        
        custom_var = self.visual_config_vars[visual_id][f'{field_type}_custom']
        
        # Entry box with validation - using tk.Entry for background color control
        custom_entry = tk.Entry(custom_frame, textvariable=custom_var, width=6, font=('Segoe UI', 9),
                               relief='solid', borderwidth=1, bg='white')
        custom_entry.pack(side=tk.LEFT, padx=(8, 0))
        
        # Delayed validation to prevent false positives while typing
        validation_timer = None
        
        def delayed_validate_popup():
            """Actual validation function called after delay for popup"""
            try:
                value = int(custom_var.get())
                if value < 50:  # Below minimum width
                    # Change background to light red
                    custom_entry.config(bg='#ffebee')
                    # Show warning in log
                    field_name = "categorical" if field_type == 'categorical' else "measure"
                    self.log_message(f"⚠️ Warning: Per-visual {field_name} custom width {value}px is below minimum (50px). Will be increased to 50px.")
                else:
                    # Reset to normal background
                    custom_entry.config(bg='white')
            except ValueError:
                # Reset to normal background for non-numeric values
                custom_entry.config(bg='white')
        
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
        selection_frame = ttk.LabelFrame(parent, text="🎯 VISUAL SELECTION", 
                                        style='Section.TLabelframe', padding="15")
        selection_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Instructions with double-click hint - single line
        instruction_frame = ttk.Frame(selection_frame)
        instruction_frame.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(instruction_frame, text="📋 Choose visuals to update", 
                 font=('Segoe UI', 10, 'bold'),
                 foreground=AppConstants.COLORS['info']).pack(side=tk.LEFT)
        
        ttk.Label(instruction_frame, text="🔄 Double-click any visual for per-visual configuration", 
                 font=('Segoe UI', 10, 'bold'),
                 foreground='#ff6600').pack(side=tk.RIGHT)
        
        # Selection controls - more compact
        controls_frame = ttk.Frame(selection_frame)
        controls_frame.pack(fill=tk.X, pady=(8, 8))
        
        ttk.Button(controls_frame, text="✓ All",
                  command=self._select_all_visuals,
                  style='Secondary.TButton', width=8).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(controls_frame, text="✗ None",
                  command=self._clear_all_visuals,
                  style='Secondary.TButton', width=8).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(controls_frame, text="📊 Tables",
                  command=self._select_tables_only,
                  style='Secondary.TButton', width=11).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(controls_frame, text="📊 Matrices",
                  command=self._select_matrices_only,
                  style='Secondary.TButton', width=11).pack(side=tk.LEFT)
        
        # Treeview for visual selection
        tree_frame = ttk.Frame(selection_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview with scrollbar
        tree_container = ttk.Frame(tree_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.visual_tree = ttk.Treeview(tree_container, height=10, selectmode='none')
        self.visual_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configure columns
        self.visual_tree['columns'] = ('type', 'page', 'fields', 'config')
        self.visual_tree.heading('#0', text='Visual Name')
        self.visual_tree.heading('type', text='Type')
        self.visual_tree.heading('page', text='Page')
        self.visual_tree.heading('fields', text='Fields')
        self.visual_tree.heading('config', text='Config')
        
        # Configure column widths
        self.visual_tree.column('#0', width=180, minwidth=150)
        self.visual_tree.column('type', width=100, minwidth=80)
        self.visual_tree.column('page', width=80, minwidth=60)
        self.visual_tree.column('fields', width=80, minwidth=60)
        self.visual_tree.column('config', width=60, minwidth=50)
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.visual_tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.visual_tree.configure(yscrollcommand=tree_scroll.set)
        
        # Configure tree styling for checkboxes
        self.visual_tree.tag_configure('selected', background='#e6f3ff')
        self.visual_tree.tag_configure('unselected', background='#f0f0f0')
        
        # Bind events
        self.visual_tree.bind('<Button-1>', self._on_tree_click)
        self.visual_tree.bind('<Motion>', self._on_tree_hover)
        self.visual_tree.bind('<Double-Button-1>', self._on_tree_double_click)  # For per-visual config
    
    def _setup_action_buttons(self, parent):
        """Setup action buttons"""
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(15, 0))
        
        # Action buttons in vertical layout for left column
        self.preview_button = ttk.Button(action_frame, text="👁️ PREVIEW",
                                        command=self._preview_changes,
                                        style='Secondary.TButton',
                                        state=tk.DISABLED)
        self.preview_button.pack(fill=tk.X, pady=(0, 6))
        
        self.apply_button = ttk.Button(action_frame, text="✅ APPLY CHANGES",
                                      command=self._apply_changes,
                                      style='Action.TButton',
                                      state=tk.DISABLED)
        self.apply_button.pack(fill=tk.X, pady=(0, 6))
        
        # Reset button
        ttk.Button(action_frame, text="🔄 RESET TAB",
                  command=self.reset_tab,
                  style='Secondary.TButton').pack(fill=tk.X)
    
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
                
                # Insert item with checkbox style
                display_text = f"☑️ {visual_info.visual_name}"
                
                item_id = self.visual_tree.insert('', 'end',
                                                  text=display_text,
                                                  values=(type_text, visual_info.page_name, fields_text, config_status),
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
        """Update the scan summary display"""
        if not self.engine:
            return
        
        summary = self.engine.get_visual_summary()
        summary_text = (f"{summary['total_visuals']} visuals found\n"
                       f"({summary['table_count']} tables, {summary['matrix_count']} matrices)")
        self.summary_label.config(text=summary_text)
    
    def _enable_configuration_ui(self):
        """Enable configuration UI after successful scan"""
        self.preview_button.config(state=tk.NORMAL)
        self.apply_button.config(state=tk.NORMAL)
    
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
        config_window = tk.Toplevel(self.main_app.root)
        config_window.title(f"Configure: {visual_info.visual_name}")
        config_window.geometry("550x395")
        config_window.resizable(False, False)
        config_window.transient(self.main_app.root)
        config_window.grab_set()
        
        # Center window
        config_window.geometry(f"+{self.main_app.root.winfo_rootx() + 100}+{self.main_app.root.winfo_rooty() + 100}")
        
        # Main frame
        main_frame = ttk.Frame(config_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text=f"⚙️ Configure: {visual_info.visual_name}", 
                 font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W)
        
        visual_type = "Table" if visual_info.visual_type == VisualType.TABLE else "Matrix"
        ttk.Label(header_frame, text=f"Type: {visual_type} | Page: {visual_info.page_name}", 
                 font=('Segoe UI', 9),
                 foreground=AppConstants.COLORS['text_secondary']).pack(anchor=tk.W, pady=(5, 0))
        
        # Get or create per-visual config variables
        if visual_info.visual_id not in self.visual_config_vars:
            self._create_per_visual_config_vars(visual_info.visual_id)
        
        visual_vars = self.visual_config_vars[visual_info.visual_id]
        
        # Configuration sections
        config_frame = ttk.Frame(main_frame)
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Categorical columns (left)
        cat_frame = ttk.LabelFrame(config_frame, text="📊 Categorical Columns", padding="10")
        cat_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self._setup_width_controls_with_custom_for_popup(cat_frame, visual_vars['categorical_preset'], 'categorical', visual_info.visual_id)
        
        # Measure columns (right)
        measure_frame = ttk.LabelFrame(config_frame, text="📈 Measure Columns", padding="10")
        measure_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(5, 0))
        
        self._setup_width_controls_with_custom_for_popup(measure_frame, visual_vars['measure_preset'], 'measure', visual_info.visual_id)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(button_frame, text="✅ Apply to This Visual",
                  command=lambda: self._apply_per_visual_config(visual_info.visual_id, config_window),
                  style='Action.TButton').pack(side=tk.LEFT)
        
        ttk.Button(button_frame, text="🌐 Copy to Global",
                  command=lambda: self._copy_to_global_config(visual_info.visual_id),
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(button_frame, text="🔄 Reset to Global",
                  command=lambda: self._reset_to_global_config(visual_info.visual_id, config_window),
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(button_frame, text="❌ Cancel",
                  command=config_window.destroy,
                  style='Secondary.TButton').pack(side=tk.RIGHT)
    
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
        icon = "☑️" if is_selected else "☐️"
        new_text = f"{icon} {visual_info.visual_name}"
        
        # Determine config status
        config_status = "Custom" if visual_info.visual_id in self.visual_config_vars else "Global"
        
        # Get current values for other columns
        current_values = list(self.visual_tree.item(item, 'values'))
        if len(current_values) >= 4:
            current_values[3] = config_status  # Update config column
        else:
            current_values.extend([''] * (4 - len(current_values)))
            current_values[3] = config_status
        
        # Update the item text and tags
        tag = 'selected' if is_selected else 'unselected'
        self.visual_tree.item(item, text=new_text, values=current_values, tags=(visual_info.visual_id, tag))
    
    def _on_tree_hover(self, event):
        """Handle mouse hover over tree items"""
        item = self.visual_tree.identify('item', event.x, event.y)
        if item:
            # Add hover effect
            self.visual_tree.selection_set(item)
        else:
            self.visual_tree.selection_remove(self.visual_tree.selection())
    
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
        preview_window = tk.Toplevel(self.main_app.root)
        preview_window.title("Preview Column Width Changes")
        preview_window.geometry("900x600")
        preview_window.resizable(True, True)
        preview_window.transient(self.main_app.root)
        preview_window.grab_set()
        
        # Center window
        preview_window.geometry(f"+{self.main_app.root.winfo_rootx() + 50}+{self.main_app.root.winfo_rooty() + 50}")
        
        # Main frame
        main_frame = ttk.Frame(preview_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text="👁️ Preview Column Width Changes", 
                 font=('Segoe UI', 14, 'bold')).pack(anchor=tk.W)
        
        ttk.Label(header_frame, text=f"Changes for {len(preview_data)} selected visual(s)", 
                 font=('Segoe UI', 9),
                 foreground=AppConstants.COLORS['text_secondary']).pack(anchor=tk.W, pady=(5, 0))
        
        # Preview tree
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        preview_tree = ttk.Treeview(tree_frame, height=15)
        preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
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
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=preview_tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        preview_tree.configure(yscrollcommand=tree_scroll.set)
        
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
        
        # Create hierarchy: Page > Visual > Field
        for page_name in sorted_pages:
            # Add page as top-level parent
            page_item = preview_tree.insert('', 'end',
                                           text=f"📝 {page_name}",
                                           values=('', '', '', ''),
                                           open=True)
            
            # Add visuals under each page
            for visual_info, config_source in page_groups[page_name]:
                # Visual display with config type
                visual_display = f"{visual_info.visual_name} ({config_source})"
                visual_item = preview_tree.insert(page_item, 'end',
                                                  text=f"📊 {visual_display}",
                                                  values=('', '', '', ''),
                                                  open=True)
                
                # Add fields under each visual
                for field in visual_info.fields:
                    if field.suggested_width is not None:
                        field_type = "📊 Category" if field.field_type == FieldType.CATEGORICAL else "📈 Measure"
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
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="❌ Close",
                  command=preview_window.destroy,
                  style='Secondary.TButton').pack(side=tk.RIGHT)
        
        ttk.Button(button_frame, text="✅ Apply These Changes",
                  command=lambda: [preview_window.destroy(), self._apply_changes()],
                  style='Action.TButton').pack(side=tk.RIGHT, padx=(0, 10))
    
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
        self.summary_label.config(text="No visuals scanned yet")
        
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
        
        # Disable buttons
        self.preview_button.config(state=tk.DISABLED)
        self.apply_button.config(state=tk.DISABLED)
        
        # Hide progress
        self.update_progress(0, "", False)
        
        # Clear and show welcome message
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
        
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
        help_window.title("Table Column Widths - Help")
        help_window.geometry("1000x830")  # Wider and shorter to match layout optimizer
        help_window.resizable(False, False)
        help_window.transient(parent_window)
        help_window.grab_set()
        
        # Center window
        help_window.geometry(f"+{parent_window.winfo_rootx() + 50}+{parent_window.winfo_rooty() + 50}")
        
        self._create_help_content(help_window)
    
    def _create_help_content(self, help_window):
        """Create help content for table column widths"""
        help_window.configure(bg=AppConstants.COLORS['background'])
        
        # Main container
        container = ttk.Frame(help_window, padding="20")
        container.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(container, text="📊 Table Column Widths - Help", 
                 font=('Segoe UI', 16, 'bold'), 
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 15))
        
        # Orange warning box
        warning_frame = ttk.Frame(container)
        warning_frame.pack(fill=tk.X, pady=(0, 15))
        
        warning_container = tk.Frame(warning_frame, bg=AppConstants.COLORS['warning'], 
                                   padx=15, pady=10, relief='solid', borderwidth=2)
        warning_container.pack(fill=tk.X)
        
        ttk.Label(warning_container, text="⚠️  IMPORTANT DISCLAIMERS & REQUIREMENTS", 
                 font=('Segoe UI', 12, 'bold'), 
                 background=AppConstants.COLORS['warning'],
                 foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W)
        
        warnings = [
            "• This tool ONLY works with PBIP enhanced report format (PBIR) files",
            "• This is NOT officially supported by Microsoft - use at your own discretion",
            "• Requires TMDL files in semantic model definition folder",
            "• Always keep backups of your original reports before optimization",
            "• Test thoroughly and validate results before production use",
            "• Enable 'Store reports using enhanced metadata format (PBIR)' in Power BI Desktop"
        ]
        
        for warning in warnings:
            ttk.Label(warning_container, text=warning, font=('Segoe UI', 10),
                     background=AppConstants.COLORS['warning'],
                     foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W, pady=1)
        
        # Top sections in 2-column layout
        top_sections_frame = ttk.Frame(container)
        top_sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        top_sections_frame.columnconfigure(0, weight=1)
        top_sections_frame.columnconfigure(1, weight=1)
        
        # LEFT COLUMN TOP: What This Tool Does
        left_top_frame = ttk.Frame(top_sections_frame)
        left_top_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(left_top_frame, text="🎯 What This Tool Does", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        what_items = [
            "✅ Standardizes column widths across all tables and matrices in your report",
            "✅ Provides intelligent width presets (Narrow, Medium, Wide, Fit to Header, Fit to Totals)",
            "✅ Supports both global configuration and per-visual customization",
            "✅ Auto-sizes columns based on headers or totals for optimal fit",
            "✅ Disables auto-size to preserve your width settings",
            "✅ Provides preview before applying changes"
        ]
        
        for item in what_items:
            ttk.Label(left_top_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # RIGHT COLUMN TOP: File Requirements
        right_top_frame = ttk.Frame(top_sections_frame)
        right_top_frame.grid(row=0, column=1, sticky=(tk.N, tk.W, tk.E), padx=(10, 0))
        
        ttk.Label(right_top_frame, text="📁 File Requirements", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        file_items = [
            "✅ Only .pbip format files (.pbip folders) are supported",
            "✅ Must contain semantic model definition folder with TMDL files",
            "✅ Requires diagramLayout.json file for layout data",
            "✅ Write permissions to PBIP folder (for saving changes)",
            "❌ Legacy format with report.json files are NOT supported",
            "❌ .pbix files are NOT supported",
            "❌ .pbit files are NOT supported"
        ]
        
        for item in file_items:
            ttk.Label(right_top_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # Bottom sections in 2-column layout
        bottom_sections_frame = ttk.Frame(container)
        bottom_sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        bottom_sections_frame.columnconfigure(0, weight=1)
        bottom_sections_frame.columnconfigure(1, weight=1)
        
        # LEFT COLUMN BOTTOM: Width Presets Explained
        left_bottom_frame = ttk.Frame(bottom_sections_frame)
        left_bottom_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(left_bottom_frame, text="⚙️ Width Presets Explained", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        preset_items = [
            "• Narrow: Compact columns (good for dense data tables)",
            "• Medium: Balanced width for general use",
            "• Wide: Spacious columns for readability",
            "• Fit to Header: Auto-sizes to fit column header text",
            "• Fit to Totals: Auto-sizes to fit total row (best for measures)",
            "• Custom: Specify exact pixel width (minimum 50px)"
        ]
        
        for item in preset_items:
            ttk.Label(left_bottom_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # RIGHT COLUMN BOTTOM: Important Notes
        right_bottom_frame = ttk.Frame(bottom_sections_frame)
        right_bottom_frame.grid(row=0, column=1, sticky=(tk.N, tk.W, tk.E), padx=(10, 0))
        
        ttk.Label(right_bottom_frame, text="⚠️ Important Notes", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        notes_items = [
            "• ONLY works with PBIP enhanced report format (PBIR)",
            "• This tool is NOT officially supported by Microsoft",
            "• Always backup your .pbip files before applying changes",
            "• The tool modifies diagramLayout.json in your PBIP folder",
            "• Test the optimized layout in Power BI Desktop before sharing",
            "• Large models may take several minutes to analyze and optimize"
        ]
        
        for item in notes_items:
            ttk.Label(right_bottom_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # Button frame at bottom
        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X, pady=(10, 0), side=tk.BOTTOM)
        
        ttk.Button(button_frame, text="❌ Close", 
                  command=help_window.destroy,
                  style='Action.TButton').pack(pady=(5, 0))
        
        help_window.bind('<Escape>', lambda event: help_window.destroy())
