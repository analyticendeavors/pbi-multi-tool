"""
Field Parameters UI - Main Tab Orchestrator
Built by Reid Havens of Analytic Endeavors

Main UI tab that coordinates all panels and manages state.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging
from typing import Optional, Dict, Any, List, Tuple

from core.constants import AppConstants
from core.ui_base import BaseToolTab, ThemedMessageBox, SquareIconButton, RoundedButton
from core.local_model_cache import get_local_model_cache
from tools.field_parameters.field_parameters_core import (
    FieldParameter, FieldItem, CategoryLevel, 
    FieldParameterParser, FieldParameterGenerator
)
from tools.field_parameters.panels import (
    CategoryManagerPanel, ModelConnectionPanel,
    ParameterConfigPanel, AvailableFieldsPanel, TmdlPreviewPanel
)
from tools.field_parameters.field_parameters_builder import ParameterBuilderPanel


class FieldParametersTab(BaseToolTab):
    """
    Main UI tab for Field Parameters tool
    
    Coordinates all panels and manages the overall state of the field parameter
    being created or edited.
    """
    
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "field_parameters", "Field Parameters")
        self.logger = logging.getLogger(__name__)
        
        # State
        self.current_parameter: Optional[FieldParameter] = None
        self.connected_model: Optional[str] = None
        self.available_tables: Dict[str, List[str]] = {}  # table_name -> [field_names]
        self.existing_parameters: List[str] = []
        
        # Panel references
        self.connection_panel: Optional[ModelConnectionPanel] = None
        self.config_panel: Optional[ParameterConfigPanel] = None
        self.fields_panel: Optional[AvailableFieldsPanel] = None
        self.builder_panel: Optional[ParameterBuilderPanel] = None
        self.category_panel: Optional[CategoryManagerPanel] = None
        self.preview_panel: Optional[TmdlPreviewPanel] = None

        # Bottom action buttons (outside preview panel frame, centered)
        self._button_container: Optional[tk.Frame] = None
        self.apply_btn: Optional[RoundedButton] = None
        self.copy_btn: Optional[RoundedButton] = None
        self.reset_btn: Optional[RoundedButton] = None
        
        self.setup_ui()

        # NOTE: Theme callback registration is handled by BaseToolTab._handle_theme_change
        # which already calls self.on_theme_changed(). Do NOT register again here
        # as double-registration causes style conflicts during navigation.

        self.logger.info("Field Parameters tab initialized")
    
    def setup_ui(self):
        """Setup the main UI layout"""
        # Create main container with padding - FIXED: use self.frame
        main_container = ttk.Frame(self.frame)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Top section: Model connection and parameter config - use grid for equal heights
        top_frame = ttk.Frame(main_container)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # Configure grid columns - equal 50/50 layout with uniform sizing
        top_frame.columnconfigure(0, weight=1, uniform='top_panels')  # Model connection
        top_frame.columnconfigure(1, weight=1, uniform='top_panels')  # Parameter configuration

        self.connection_panel = ModelConnectionPanel(top_frame, self)
        self.connection_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        self.config_panel = ParameterConfigPanel(top_frame, self)
        self.config_panel.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # Help button in upper right corner - placed on config_panel with negative y
        # This matches the pattern used by FileInputSection and other tools
        self._help_button = None
        help_icon = self._load_icon_for_button("question", size=14)
        if help_icon:
            self._button_icons['question'] = help_icon
            self._help_button = SquareIconButton(
                self.config_panel, icon=help_icon, command=self.show_help_dialog,
                tooltip_text="Help", size=26, radius=6,
                bg_normal_override=AppConstants.CORNER_ICON_BG
            )
            # Position in header area (y=-35 places it above section content)
            self._help_button.place(relx=1.0, y=-35, anchor=tk.NE, x=0)
        
        # Middle section: Three-column layout
        # Use tk.Frame (not ttk) for stable geometry during theme changes
        middle_frame = tk.Frame(main_container, bg=self._theme_manager.colors['background'])
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self._middle_frame = middle_frame  # Store for theme updates and width locking

        # Configure column weights for responsive layout (will be locked after initial render)
        middle_frame.columnconfigure(0, weight=3, minsize=300)  # Available fields (+20)
        middle_frame.columnconfigure(1, weight=1, minsize=230)  # Parameter builder (-60, narrowest)
        middle_frame.columnconfigure(2, weight=2, minsize=260)  # Categories (+30)
        middle_frame.rowconfigure(0, weight=1)  # Allow vertical expansion

        # Track if column widths have been locked
        self._columns_width_locked = False
        
        # Left: Available fields (no padx - LabelFrame borders provide separation)
        self.fields_panel = AvailableFieldsPanel(middle_frame, self)
        self.fields_panel.grid(row=0, column=0, sticky="nsew")

        # Center: Parameter builder (no padx - LabelFrame borders provide separation)
        self.builder_panel = ParameterBuilderPanel(middle_frame, self)
        self.builder_panel.grid(row=0, column=1, sticky="nsew")

        # Connect fields_panel drag-drop to builder_panel (must be done after builder_panel exists)
        self.fields_panel.set_drop_target(self.builder_panel)

        # Right: Category manager (no padx - LabelFrame borders provide separation)
        self.category_panel = CategoryManagerPanel(middle_frame, self)
        self.category_panel.grid(row=0, column=2, sticky="nsew")
        
        # Action buttons - pack FIRST with side=BOTTOM to reserve space before preview panel expands
        # This ensures buttons are always visible regardless of window height
        self._setup_action_buttons(main_container)

        # Bottom section: TMDL preview (expands to fill remaining space above buttons)
        self.preview_panel = TmdlPreviewPanel(main_container, self)
        self.preview_panel.pack(fill=tk.BOTH, expand=True)

        # Initial state
        self._set_initial_state()

        # Schedule column width locking after initial render (widgets must be visible first)
        self.frame.after(100, self._lock_column_widths)
    
    def _set_initial_state(self):
        """Set initial disabled state until model is connected"""
        self.config_panel.set_enabled(False)
        self.fields_panel.set_enabled(False)
        self.builder_panel.set_enabled(False)
        self.category_panel.set_enabled(False)
        self.preview_panel.set_enabled(False)
        # Buttons start disabled
        if self.apply_btn:
            self.apply_btn.set_enabled(False)
        if self.copy_btn:
            self.copy_btn.set_enabled(False)

    def _setup_action_buttons(self, parent: tk.Widget):
        """Setup centered action button bar below preview panel"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        canvas_bg = colors['section_bg']

        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Centered container for buttons - pack at BOTTOM to reserve space before preview panel expands
        self._button_container = tk.Frame(parent, bg=canvas_bg)
        self._button_container.pack(side=tk.BOTTOM, pady=(15, 0))

        # Load button icons
        execute_icon = self._load_icon_for_button('execute', size=16)
        coding_icon = self._load_icon_for_button('coding', size=16)
        reset_icon = self._load_icon_for_button('reset', size=16)

        # Store icons for theme updates
        if execute_icon:
            self._button_icons['execute'] = execute_icon
        if coding_icon:
            self._button_icons['coding'] = coding_icon
        if reset_icon:
            self._button_icons['reset'] = reset_icon

        # COPY TMDL button (secondary) - first in order
        self.copy_btn = RoundedButton(
            self._button_container,
            text="COPY TMDL",
            command=self.on_copy_tmdl,
            icon=coding_icon,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=38, radius=6,
            font=('Segoe UI', 10),
            canvas_bg=canvas_bg
        )
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 15))

        # APPLY TO MODEL button (primary) - center position
        self.apply_btn = RoundedButton(
            self._button_container,
            text="APPLY TO MODEL",
            command=self.on_apply_to_model,
            icon=execute_icon,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=38, radius=6,
            font=('Segoe UI', 10, 'bold'),
            canvas_bg=canvas_bg
        )
        self.apply_btn.pack(side=tk.LEFT, padx=(0, 15))

        # RESET ALL button (secondary) - last in order
        self.reset_btn = RoundedButton(
            self._button_container,
            text="RESET ALL",
            command=self.reset_tab,
            icon=reset_icon,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=38, radius=6,
            font=('Segoe UI', 10),
            canvas_bg=canvas_bg
        )
        self.reset_btn.pack(side=tk.LEFT)

    # =========================================================================
    # Required BaseToolTab Abstract Methods
    # =========================================================================
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        self.logger.info("Resetting Field Parameters tab")

        # Disconnect from model
        self.on_model_disconnected()

        # Reset all UI components (don't refresh model list - shared cache handles repopulation)
        if self.config_panel:
            self.config_panel.clear()
            self.config_panel.mode.set("new")
            self.config_panel.parameter_name.set("")
            self.config_panel.keep_lineage.set(False)

    def on_tab_activated(self) -> None:
        """Called when this tab becomes active - ensures theme is applied"""
        # Schedule full theme refresh after widgets are rendered
        # Use _handle_theme_change to ensure _setup_common_styling() is called
        # This fixes background color issues when navigating back to this tab
        current_theme = self._theme_manager.current_theme
        self.frame.after(150, lambda: self._handle_theme_change(current_theme))

        # Auto-populate model dropdown from shared cache
        # This provides instant model list if Hot Swap tab already scanned
        if self.connection_panel:
            self.connection_panel.populate_from_shared_cache()

    def show_help_dialog(self) -> None:
        """Show help dialog for Field Parameters tool"""
        colors = self._theme_manager.colors

        # Find parent window
        parent_window = self.frame.winfo_toplevel()

        help_window = tk.Toplevel(parent_window)
        help_window.withdraw()  # Hide until fully styled (prevents white flash)
        help_window.title("Field Parameters - Help")
        help_window.geometry("1000x690")
        help_window.resizable(False, False)
        help_window.transient(parent_window)
        help_window.grab_set()
        help_window.configure(bg=colors['background'])

        # Set AE favicon icon
        try:
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'config'):
                help_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Main container
        container = ttk.Frame(help_window, padding="20")
        container.pack(fill=tk.BOTH, expand=True)

        # Header
        ttk.Label(container, text="Field Parameters - Help",
                 font=('Segoe UI', 16, 'bold'),
                 foreground=colors['title_color']).pack(anchor=tk.W, pady=(0, 15))

        # Warning box with theme-aware colors
        warning_frame = ttk.Frame(container)
        warning_frame.pack(fill=tk.X, pady=(0, 15))

        warning_bg = colors.get('warning_bg', colors['warning'])
        warning_text = colors.get('warning_text', '#ffffff')
        warning_container = tk.Frame(warning_frame, bg=warning_bg,
                                   padx=15, pady=10, relief='flat', borderwidth=0)
        warning_container.pack(fill=tk.X)

        tk.Label(warning_container, text="IMPORTANT DISCLAIMERS & REQUIREMENTS",
                 font=('Segoe UI', 12, 'bold'),
                 bg=warning_bg,
                 fg=warning_text).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        # Configure 2-column layout for warnings
        warning_container.columnconfigure(0, weight=1)
        warning_container.columnconfigure(1, weight=1)

        warnings_left = [
            "Connects to Power BI models via Tabular Object Model (TOM)",
            "Power BI Desktop must be installed (provides required DLLs)",
            "Models must be open in Power BI Desktop or via XMLA endpoint"
        ]

        warnings_right = [
            "Always backup your models before applying changes",
            "Review TMDL code before pasting into model definitions",
            "Test generated parameters in Power BI Desktop before production"
        ]

        for i, warning in enumerate(warnings_left):
            tk.Label(warning_container, text=f"  {warning}", font=('Segoe UI', 10),
                     bg=warning_bg, fg=warning_text, anchor=tk.W).grid(
                         row=i+1, column=0, sticky=tk.W, pady=1)

        for i, warning in enumerate(warnings_right):
            tk.Label(warning_container, text=f"  {warning}", font=('Segoe UI', 10),
                     bg=warning_bg, fg=warning_text, anchor=tk.W).grid(
                         row=i+1, column=1, sticky=tk.W, pady=1, padx=(15, 0))

        # Top sections in 2-column layout
        top_sections_frame = ttk.Frame(container)
        top_sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        top_sections_frame.columnconfigure(0, weight=1)
        top_sections_frame.columnconfigure(1, weight=1)

        # LEFT COLUMN TOP: What This Tool Does
        left_top_frame = ttk.Frame(top_sections_frame)
        left_top_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 10))

        ttk.Label(left_top_frame, text="What This Tool Does",
                 font=('Segoe UI', 12, 'bold'),
                 foreground=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        what_items = [
            "Create Field Parameters with visual drag-and-drop interface",
            "Edit existing Field Parameters while preserving structure",
            "Support multi-level category hierarchies (unlimited nesting)",
            "Customize display names (different from measure names)",
            "Generate proper TMDL code with correct NAMEOF() syntax",
            "Preserve lineage tags for version control compatibility",
            "Bulk revert display names to original field names"
        ]

        for item in what_items:
            ttk.Label(left_top_frame, text=f"• {item}",
                     font=('Segoe UI', 10),
                     foreground=colors['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)

        # RIGHT COLUMN TOP: Connection Options
        right_top_frame = ttk.Frame(top_sections_frame)
        right_top_frame.grid(row=0, column=1, sticky=(tk.N, tk.W, tk.E), padx=(10, 0))

        ttk.Label(right_top_frame, text="Connection Options",
                 font=('Segoe UI', 12, 'bold'),
                 foreground=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        conn_items = [
            "External Tool Launch: Auto-connects when launched from Power BI",
            "Auto-Discovery: Click 'Refresh' to scan for open models",
            "Manual Connection: Use 'Connect...' to specify port manually",
            "Cloud XMLA: Connect to Premium/PPU workspaces",
            ".pbix files cannot be connected directly",
            "Requires Power BI Desktop or XMLA endpoint access"
        ]

        for item in conn_items:
            ttk.Label(right_top_frame, text=f"• {item}",
                     font=('Segoe UI', 10),
                     foreground=colors['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)

        # Bottom sections in 2-column layout
        bottom_sections_frame = ttk.Frame(container)
        bottom_sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        bottom_sections_frame.columnconfigure(0, weight=1)
        bottom_sections_frame.columnconfigure(1, weight=1)

        # LEFT COLUMN BOTTOM: Quick Start Guide
        left_bottom_frame = ttk.Frame(bottom_sections_frame)
        left_bottom_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 10))

        ttk.Label(left_bottom_frame, text="Quick Start Guide",
                 font=('Segoe UI', 12, 'bold'),
                 foreground=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        guide_items = [
            "1. Open a model in Power BI Desktop",
            "2. Click 'Refresh' or launch as External Tool",
            "3. Select model from dropdown and click 'Connect'",
            "4. Choose 'Create New' or 'Edit Existing' parameter",
            "5. Drag fields from Available Fields to Parameter Builder",
            "6. Optionally add categories via 'Edit Categories'",
            "7. Copy generated TMDL code and paste into model"
        ]

        for item in guide_items:
            ttk.Label(left_bottom_frame, text=item,
                     font=('Segoe UI', 10),
                     foreground=colors['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)

        # RIGHT COLUMN BOTTOM: Tips & Troubleshooting
        right_bottom_frame = ttk.Frame(bottom_sections_frame)
        right_bottom_frame.grid(row=0, column=1, sticky=(tk.N, tk.W, tk.E), padx=(10, 0))

        ttk.Label(right_bottom_frame, text="Tips & Troubleshooting",
                 font=('Segoe UI', 12, 'bold'),
                 foreground=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        tips_items = [
            "Keep Lineage Tags checked when editing existing parameters",
            "Uncheck for new parameters (let Power BI generate IDs)",
            "Category sort orders: Use 1, 2, 3 or 10, 20, 30 for flexibility",
            "Double-click items to rename display names",
            "Drag items to reorder (controls sort order in slicer)",
            "No models found? Ensure model is fully loaded in PBI Desktop",
            "Connection fails? Try manual port entry via 'Connect...'"
        ]

        for item in tips_items:
            ttk.Label(right_bottom_frame, text=f"• {item}",
                     font=('Segoe UI', 10),
                     foreground=colors['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)

        help_window.bind('<Escape>', lambda event: help_window.destroy())

        # Center dialog and show (after all content built to prevent flash)
        help_window.update_idletasks()
        x = parent_window.winfo_rootx() + (parent_window.winfo_width() - 1000) // 2
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 690) // 2
        help_window.geometry(f"1000x690+{x}+{y}")

        # Set dark title bar BEFORE showing to prevent white flash
        help_window.update()
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled
    
    # =========================================================================
    # Model Connection Methods
    # =========================================================================
    
    def on_model_connected(self, model_name: str, tables: Dict[str, List[str]],
                          existing_params: List[str]):
        """Called when model connection is established"""
        self.logger.info(f"Model connected: {model_name}")
        self.connected_model = model_name
        self.available_tables = tables
        self.existing_parameters = existing_params

        # Clear any parameter data from previous model (important for model switching)
        self.current_parameter = None
        self.builder_panel.clear()
        self.builder_panel.set_enabled(False)
        self.category_panel.clear()
        self.category_panel.set_enabled(False)
        self.preview_panel.clear()
        self.preview_panel.set_enabled(False)

        # Update config panel (clear first to reset mode to "Create New")
        self.config_panel.clear()
        self.config_panel.set_enabled(True)
        self.config_panel.update_parameter_list(existing_params)
        self.fields_panel.set_enabled(True)
        self.fields_panel.update_available_fields(tables)
        # Note: Success dialog is shown by the connection panel (local vs cloud specific)
    
    def on_model_disconnected(self):
        """Called when model is disconnected"""
        self.logger.info("Model disconnected")
        self.connected_model = None
        self.available_tables = {}
        self.existing_parameters = []
        
        # Reset UI state
        self._set_initial_state()
        self.config_panel.clear()
        self.fields_panel.clear()
        self.builder_panel.clear()
        self.category_panel.clear()
        self.preview_panel.clear()
    
    # =========================================================================
    # Parameter Configuration Methods
    # =========================================================================
    
    def on_create_new_parameter(self, parameter_name: str, keep_lineage: bool):
        """Create a new field parameter"""
        self.logger.info(f"Creating new parameter: {parameter_name}")
        
        self.current_parameter = FieldParameter(
            table_name=f".Parameter - {parameter_name}",
            parameter_name=parameter_name,
            keep_lineage_tags=keep_lineage
        )
        
        # Enable builder panels
        self.builder_panel.set_enabled(True)
        self.category_panel.set_enabled(True)
        self.preview_panel.set_enabled(True)

        # Enable action buttons
        if self.apply_btn:
            self.apply_btn.set_enabled(True)
        if self.copy_btn:
            self.copy_btn.set_enabled(True)

        # Clear and prepare for new data
        self.builder_panel.clear()
        self.category_panel.clear()
        self.preview_panel.clear()

        self.update_preview()
    
    def on_edit_existing_parameter(self, parameter_name: str, keep_lineage: bool):
        """Load and edit an existing field parameter using TOM direct loading"""
        self.logger.info(f"Editing existing parameter: {parameter_name}")

        try:
            # Show loading indicator IMMEDIATELY before any TOM operations
            self.builder_panel.show_loading_overlay(f"Loading '{parameter_name}'...")
            self.frame.update_idletasks()

            # Get connector
            from core.pbi_connector import get_connector

            connector = get_connector()

            # Load parameter directly from TOM (bypasses TMDL parsing)
            self.logger.info(f"Loading parameter from TOM: {parameter_name}")
            self.current_parameter = connector.load_parameter_from_tom(parameter_name)
            
            if not self.current_parameter:
                self.logger.error(f"Failed to load parameter: {parameter_name}")
                self.builder_panel.hide_loading_overlay()
                ThemedMessageBox.showerror(
                    self.frame.winfo_toplevel(),
                    "Load Failed",
                    f"Could not load parameter '{parameter_name}' from model.\n\n"
                    "The table may not exist, may not be a field parameter, or may not be accessible.\n\n"
                    "Check the log file for details."
                )
                return
            
            self.logger.info(f"Successfully loaded parameter with {len(self.current_parameter.fields)} fields")
            
            # Update lineage preference
            self.current_parameter.keep_lineage_tags = keep_lineage
            
            # Enable panels BEFORE loading (important!)
            self.builder_panel.set_enabled(True)
            self.category_panel.set_enabled(True)
            self.preview_panel.set_enabled(True)

            # Enable action buttons
            if self.apply_btn:
                self.apply_btn.set_enabled(True)
            if self.copy_btn:
                self.copy_btn.set_enabled(True)

            # Populate UI panels
            self.logger.info("Loading parameter into builder panel...")
            self.builder_panel.load_parameter(self.current_parameter)
            self.logger.info("Loading categories into category panel...")
            self.category_panel.load_categories(self.current_parameter.category_levels)
            
            # Update preview
            self.logger.info("Updating preview...")
            self.update_preview()
            
            self.logger.info(f"Successfully loaded parameter: {parameter_name}")

            # Check for custom sort order and alert user
            has_custom_sort = self._has_custom_sort_order()
            if has_custom_sort:
                ThemedMessageBox.showinfo(
                    self.frame.winfo_toplevel(),
                    "Parameter Loaded - Custom Sort Detected",
                    f"Successfully loaded parameter '{parameter_name}'\n\n"
                    f"Fields: {len(self.current_parameter.fields)}\n"
                    f"Categories: {len(self.current_parameter.category_levels)}\n\n"
                    "Custom field order detected!\n"
                    "The field order values are non-sequential (e.g., 1,1,1,2,2 or 1,2,3,10,11).\n\n"
                    "To preserve this custom sort when applying changes, uncheck\n"
                    "'Update field order' in the preview panel."
                )
            else:
                ThemedMessageBox.showinfo(
                    self.frame.winfo_toplevel(),
                    "Parameter Loaded",
                    f"Successfully loaded parameter '{parameter_name}'\n\n"
                    f"Fields: {len(self.current_parameter.fields)}\n"
                    f"Categories: {len(self.current_parameter.category_levels)}"
                )

        except Exception as e:
            self.logger.error(f"Error loading parameter: {e}", exc_info=True)
            self.builder_panel.hide_loading_overlay()
            ThemedMessageBox.showerror(
                self.frame.winfo_toplevel(),
                "Load Error",
                f"Failed to load parameter '{parameter_name}':\n\n{str(e)}\n\n"
                f"Check the log file for details."
            )
    
    # =========================================================================
    # Field Management Methods
    # =========================================================================

    def validate_can_add_field(self, show_warning: bool = True) -> bool:
        """
        Check if fields can be added (parameter exists).
        Call this ONCE before bulk add operations.

        Args:
            show_warning: If True, shows a warning dialog when validation fails

        Returns:
            True if a parameter exists and fields can be added, False otherwise
        """
        if not self.current_parameter:
            if show_warning:
                ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Parameter", "Please create or select a parameter first")
            return False
        return True

    def on_add_field(self, table_name: str, field_name: str, position: int = None):
        """
        Add a field to the parameter at the specified position.

        NOTE: Callers should call validate_can_add_field() ONCE before bulk operations.
        This method silently returns if no parameter exists.

        Args:
            table_name: The source table name
            field_name: The field name
            position: Where to insert (None=end, 0=top, etc.)
        """
        if not self.current_parameter:
            # Silently return - caller should have validated first
            return
        
        self.logger.info(f"Adding field: {table_name}.{field_name} at position={position}")
        
        # Note: Duplicate checking is now handled by the caller (panels)
        # This allows explicit duplicate additions when user confirms
        
        # Initialize categories with empty entries for each category level
        # (999.0 sort, empty label = uncategorized) - ensures DAX tuples have matching column counts
        categories = []
        if self.current_parameter.category_levels:
            for _ in self.current_parameter.category_levels:
                categories.append((999.0, ""))

        # Create field item with default display name
        field_item = FieldItem(
            display_name=field_name,
            field_reference=f"NAMEOF('{table_name}'[{field_name}])",
            table_name=table_name,
            field_name=field_name,
            order_within_group=1,
            original_order_within_group=1,
            categories=categories
        )
        
        # Insert at position in the data model
        if position is not None and position < len(self.current_parameter.fields):
            self.current_parameter.fields.insert(position, field_item)
        else:
            self.current_parameter.fields.append(field_item)
        
        # Add to UI at same position
        self.builder_panel.add_field(field_item, position=position)
        self.builder_panel.update_edit_categories_button()
        self.update_preview()
    
    def on_remove_field(self, field_item: FieldItem):
        """Remove a field from the parameter"""
        if not self.current_parameter:
            return
        
        self.logger.info(f"Removing field: {field_item.field_name}")
        
        if field_item in self.current_parameter.fields:
            self.current_parameter.fields.remove(field_item)
            self.builder_panel.remove_field(field_item)
            self.builder_panel.update_edit_categories_button()
            self.update_preview()
    
    def on_field_reordered(self, fields: List[FieldItem]):
        """Called when fields are reordered via drag-drop"""
        if not self.current_parameter:
            return
        
        self.logger.info("Fields reordered")
        self.current_parameter.fields = fields
        self.update_preview()
    
    def on_field_display_name_changed(self, field_item: FieldItem, new_name: str):
        """Called when a field's display name is changed"""
        self.logger.info(f"Display name changed: {field_item.field_name} -> {new_name}")
        field_item.display_name = new_name
        self.update_preview()
    
    def on_field_category_changed(self, field_item: FieldItem, category_path: List[Tuple[int, str]]):
        """Called when a field's category assignment changes"""
        self.logger.info(f"Category changed for {field_item.field_name}: {category_path}")
        field_item.categories = category_path
        self.update_preview()
        # Update alignment button style if in grouped view mode
        if hasattr(self, 'builder_panel') and self.builder_panel:
            self.builder_panel.update_alignment_status()
    
    def on_revert_all_display_names(self):
        """Revert all display names to original field names"""
        if not self.current_parameter:
            return
        
        confirm = ThemedMessageBox.askyesno(
            self.frame.winfo_toplevel(),
            "Revert All Display Names",
            "Are you sure you want to revert all display names to their original field names?\n\n"
            "This cannot be undone."
        )

        if confirm:
            self.logger.info("Reverting all display names")
            for field_item in self.current_parameter.fields:
                field_item.display_name = field_item.field_name

            self.builder_panel.refresh_all_fields()
            self.update_preview()
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Reverted", "All display names have been reset to original field names")
    
    # =========================================================================
    # Category Management Methods
    # =========================================================================
    
    def on_add_category_level(self, level_name: str, column_name: str):
        """Add a new category level"""
        if not self.current_parameter:
            return

        self.logger.info(f"Adding category level: {level_name}")

        category_level = CategoryLevel(
            name=level_name,
            sort_order=len(self.current_parameter.category_levels) + 1,
            column_name=column_name
        )

        self.current_parameter.category_levels.append(category_level)

        # Add empty category entry for all existing fields (999.0 sort, empty label = uncategorized)
        for field_item in self.current_parameter.fields:
            field_item.categories.append((999.0, ""))

        self.builder_panel.update_category_options(self.current_parameter.category_levels)
        self.builder_panel.update_edit_categories_button()
        self.update_preview()
    
    def on_remove_category_level(self, level_idx: int):
        """Remove a category level"""
        if not self.current_parameter or level_idx >= len(self.current_parameter.category_levels):
            return
        
        level = self.current_parameter.category_levels[level_idx]
        self.logger.info(f"Removing category level: {level.name}")
        
        # Remove from parameter
        self.current_parameter.category_levels.pop(level_idx)
        
        # Clear category assignments from fields that used this level
        for field_item in self.current_parameter.fields:
            if len(field_item.categories) > level_idx:
                field_item.categories = field_item.categories[:level_idx]
        
        self.builder_panel.update_category_options(self.current_parameter.category_levels)
        self.builder_panel.update_edit_categories_button()
        self.update_preview()

    def on_categories_reordered(self, category_levels: List[CategoryLevel]):
        """Called when categories are reordered"""
        if not self.current_parameter:
            return
        
        self.logger.info("Categories reordered")
        self.current_parameter.category_levels = category_levels
        self.update_preview()
    
    # =========================================================================
    # Preview and Output Methods
    # =========================================================================
    
    def update_preview(self):
        """Update the TMDL code preview"""
        if not self.current_parameter:
            self.preview_panel.clear()
            self.preview_panel.update_category_columns([])
            return

        try:
            # Update category columns dropdown in preview panel
            self.preview_panel.update_category_columns(self.current_parameter.category_levels)

            # Update builder panel filter with current category labels
            self.builder_panel.update_category_options(self.current_parameter.category_levels)

            # Generate TMDL code
            generator = FieldParameterGenerator()
            tmdl_code = generator.generate_tmdl(self.current_parameter)

            # Update preview
            self.preview_panel.set_tmdl_code(tmdl_code)

        except Exception as e:
            self.logger.error(f"Error generating TMDL: {e}", exc_info=True)
            self.preview_panel.set_error(f"Error generating code: {e}")
    
    def on_copy_tmdl(self):
        """Copy TMDL code to clipboard"""
        tmdl_code = self.preview_panel.get_tmdl_code()
        if tmdl_code:
            # Check for unmapped category values if categories exist
            if self.current_parameter and self.current_parameter.category_levels:
                unmapped_fields = self._get_fields_with_unmapped_categories()
                if unmapped_fields:
                    # Build warning message
                    field_list = "\n".join(f"  • {name}" for name in unmapped_fields[:10])
                    if len(unmapped_fields) > 10:
                        field_list += f"\n  ... and {len(unmapped_fields) - 10} more"

                    result = ThemedMessageBox.askyesno(
                        self.frame.winfo_toplevel(),
                        "Unmapped Categories",
                        f"{len(unmapped_fields)} field(s) have unmapped category values:\n\n"
                        f"{field_list}\n\n"
                        "Fields without category assignments will appear as 'Uncategorized' in the output.\n\n"
                        "Copy anyway?"
                    )
                    if not result:
                        return  # User cancelled

            # FIXED: use self.frame for clipboard methods
            self.frame.clipboard_clear()
            self.frame.clipboard_append(tmdl_code)
            ThemedMessageBox.showinfo(
                self.frame.winfo_toplevel(),
                "Copied",
                "TMDL code has been copied to clipboard.\n\n"
                "Paste it into Tabular Editor or your model definition file."
            )

    def on_apply_to_model(self):
        """Apply parameter changes directly to the connected model via TOM"""
        if not self.current_parameter:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Parameter", "Please create or load a parameter first")
            return

        if not self.connected_model:
            ThemedMessageBox.showwarning(
                self.frame.winfo_toplevel(),
                "Not Connected",
                "Not connected to a Power BI model.\n\n"
                "Use 'Copy TMDL' to copy the code and paste it manually into Tabular Editor."
            )
            return

        # Validate parameter
        if not self.validate_parameter():
            return

        # Check for custom sort order that will be overwritten
        if self.current_parameter.update_field_order and self._has_custom_sort_order():
            result = ThemedMessageBox.askyesno(
                self.frame.winfo_toplevel(),
                "Custom Sort Order Detected",
                "The original field order values are non-sequential (custom sorting detected).\n\n"
                "Examples: values like 1, 1, 1, 2, 2 instead of 1, 2, 3, 4, 5\n\n"
                "With 'Update field order' checked, applying will overwrite these with sequential values.\n\n"
                "To preserve custom sort, cancel and uncheck 'Update field order' in the preview panel.\n\n"
                "Continue and overwrite custom sort?"
            )
            if not result:
                return

        # Check for unmapped categories
        if self.current_parameter.category_levels:
            unmapped_fields = self._get_fields_with_unmapped_categories()
            if unmapped_fields:
                field_list = "\n".join(f"  • {name}" for name in unmapped_fields[:10])
                if len(unmapped_fields) > 10:
                    field_list += f"\n  ... and {len(unmapped_fields) - 10} more"

                result = ThemedMessageBox.askyesno(
                    self.frame.winfo_toplevel(),
                    "Unmapped Categories",
                    f"{len(unmapped_fields)} field(s) have unmapped category values:\n\n"
                    f"{field_list}\n\n"
                    "Fields without category assignments will appear as 'Uncategorized'.\n\n"
                    "Apply anyway?"
                )
                if not result:
                    return

        # Confirm apply
        table_name = self.current_parameter.table_name
        field_count = len(self.current_parameter.fields)

        # Build field order status message
        if self.current_parameter.update_field_order:
            order_status = "Field order: WILL BE UPDATED (sequential 1, 2, 3...)"
        else:
            order_status = "Field order: PRESERVED (using original values)"

        # Check if this is a cloud connection for appropriate messaging
        from core.pbi_connector import get_connector
        connector = get_connector()
        is_cloud = connector.is_cloud_connection()

        if is_cloud:
            effect_msg = "The table will be automatically refreshed so changes take effect immediately."
        else:
            effect_msg = "The changes will take effect immediately in Power BI Desktop."

        result = ThemedMessageBox.askyesno(
            self.frame.winfo_toplevel(),
            "Apply to Model",
            f"This will apply changes directly to the connected model:\n\n"
            f"Table: {table_name}\n"
            f"Fields: {field_count}\n"
            f"Model: {self.connected_model}\n\n"
            f"{order_status}\n\n"
            f"{effect_msg}\n\n"
            "Continue?"
        )

        if not result:
            return

        # Apply changes
        self.logger.info(f"Applying parameter '{table_name}' to model")

        try:
            # Reuse connector from earlier check
            success, message = connector.apply_parameter_to_model(self.current_parameter)

            if success:
                self.logger.info(f"Successfully applied parameter: {message}")
                ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Success", message)
            else:
                self.logger.error(f"Failed to apply parameter: {message}")
                ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Apply Failed", message)

        except Exception as e:
            self.logger.error(f"Error applying parameter: {e}", exc_info=True)
            ThemedMessageBox.showerror(
                self.frame.winfo_toplevel(),
                "Apply Error",
                f"Failed to apply parameter to model:\n\n{str(e)}\n\n"
                "Check the log file for details."
            )

    def _has_custom_sort_order(self) -> bool:
        """Check if original field order values are non-sequential (custom sorting).

        Returns True if:
        - Order values have duplicates (e.g., 1, 1, 1, 2, 2)
        - Order values are not sequential starting from 1 (e.g., 5, 10, 15)
        """
        if not self.current_parameter or not self.current_parameter.fields:
            return False

        original_orders = [f.original_order_within_group for f in self.current_parameter.fields]

        # Check for duplicates
        if len(original_orders) != len(set(original_orders)):
            return True

        # Check if sequential starting from 1
        expected = list(range(1, len(original_orders) + 1))
        if sorted(original_orders) != expected:
            return True

        return False

    def _get_fields_with_unmapped_categories(self) -> list:
        """Get list of field display names that have unmapped category values"""
        if not self.current_parameter or not self.current_parameter.category_levels:
            return []

        unmapped_fields = []
        num_categories = len(self.current_parameter.category_levels)

        for field_item in self.current_parameter.fields:
            # Check if field has all categories mapped
            has_unmapped = False

            if not field_item.categories:
                # No categories at all
                has_unmapped = True
            elif len(field_item.categories) < num_categories:
                # Missing some category assignments
                has_unmapped = True
            else:
                # Check if any category label is empty
                for sort_order, label in field_item.categories:
                    if not label:
                        has_unmapped = True
                        break

            if has_unmapped:
                unmapped_fields.append(field_item.display_name)

        return unmapped_fields
    
    # =========================================================================
    # Validation
    # =========================================================================
    
    def validate_parameter(self) -> bool:
        """Validate current parameter configuration"""
        if not self.current_parameter:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Parameter", "Please create or load a parameter first")
            return False

        if not self.current_parameter.parameter_name:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "Missing Name", "Please provide a parameter name")
            return False

        if not self.current_parameter.fields:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Fields", "Please add at least one field to the parameter")
            return False

        return True
    
    # =========================================================================
    # Drag and Drop Support
    # =========================================================================
    
    def set_drag_data(self, table_name: str, field_name: str):
        """Store drag data for cross-widget drag-and-drop"""
        self._drag_data = {
            'table_name': table_name,
            'field_name': field_name
        }
        self.logger.debug(f"Drag data set: {table_name}.{field_name}")
    
    def get_drag_data(self) -> Tuple[Optional[str], Optional[str]]:
        """Get stored drag data"""
        if hasattr(self, '_drag_data') and self._drag_data:
            return self._drag_data.get('table_name'), self._drag_data.get('field_name')
        return None, None
    
    def clear_drag_data(self):
        """Clear stored drag data"""
        if hasattr(self, '_drag_data'):
            self._drag_data = None
            self.logger.debug("Drag data cleared")

    # =========================================================================
    # Column Width Locking (prevents drift on theme change)
    # =========================================================================

    def _lock_column_widths(self):
        """
        Lock column widths and row height after initial render to prevent drift.
        Uses weight=0 with minsize to truly lock dimensions.
        """
        if not hasattr(self, '_middle_frame') or self._columns_width_locked:
            return

        # Get current dimensions of the panels
        col0_width = self.fields_panel.winfo_width()
        col1_width = self.builder_panel.winfo_width()
        col2_width = self.category_panel.winfo_width()
        row_height = self.fields_panel.winfo_height()

        # Only lock if widgets have valid dimensions (are mapped and sized)
        if col0_width > 10 and col1_width > 10 and col2_width > 10 and row_height > 10:
            # weight=0 on columns 0 and 1 to lock their widths
            # weight=1 on column 2 so it expands to fill extra window space (prevents black box gap)
            # minsize ensures columns don't shrink below locked widths
            self._middle_frame.columnconfigure(0, weight=0, minsize=col0_width)
            self._middle_frame.columnconfigure(1, weight=0, minsize=col1_width)
            self._middle_frame.columnconfigure(2, weight=1, minsize=col2_width)

            # Lock row height to prevent vertical resizing when content changes
            self._middle_frame.rowconfigure(0, weight=1, minsize=row_height)

            # Store locked values for theme change re-application
            self._locked_col0_width = col0_width
            self._locked_col1_width = col1_width
            self._locked_col2_width = col2_width
            self._locked_row_height = row_height
            self._columns_width_locked = True
            self.logger.debug(f"Dimensions locked: cols={col0_width}, {col1_width}, {col2_width}, row={row_height}")

            # Also lock the builder panel's internal width to prevent view mode changes from affecting it
            if self.builder_panel and hasattr(self.builder_panel, 'lock_width'):
                self.builder_panel.lock_width(col1_width)

    # =========================================================================
    # Theme Support
    # =========================================================================

    def on_theme_changed(self, theme: str = None):
        """Update widget colors when theme changes - propagate to all panels"""
        colors = self._theme_manager.colors

        # Update middle frame background only
        if hasattr(self, '_middle_frame'):
            self._middle_frame.configure(bg=colors['background'])

        # Propagate to all panels
        if self.connection_panel:
            self.connection_panel.on_theme_changed()
        if self.config_panel:
            self.config_panel.on_theme_changed()
        if self.fields_panel:
            self.fields_panel.on_theme_changed()
        if self.builder_panel:
            self.builder_panel.on_theme_changed()
        if self.category_panel:
            self.category_panel.on_theme_changed()
        if self.preview_panel:
            self.preview_panel.on_theme_changed()

        # Update action buttons
        self._update_button_theme()

        # Re-apply locked column widths to prevent drift from widget size changes
        # Schedule this after a short delay to let all theme updates complete
        if hasattr(self, '_columns_width_locked') and self._columns_width_locked:
            self.frame.after(50, self._reapply_locked_widths)

    def _reapply_locked_widths(self):
        """Re-apply locked column widths to prevent drift after theme changes"""
        if not hasattr(self, '_middle_frame') or not self._columns_width_locked:
            return

        # Re-apply the locked widths (weight=0 on cols 0,1; weight=1 on col 2 to fill extra space)
        self._middle_frame.columnconfigure(0, weight=0, minsize=self._locked_col0_width)
        self._middle_frame.columnconfigure(1, weight=0, minsize=self._locked_col1_width)
        self._middle_frame.columnconfigure(2, weight=1, minsize=self._locked_col2_width)
        self._middle_frame.rowconfigure(0, weight=1, minsize=self._locked_row_height)

    def _update_button_theme(self):
        """Update action button colors when theme changes"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        canvas_bg = colors['section_bg']

        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Update button container background
        if self._button_container:
            self._button_container.config(bg=canvas_bg)

        # Reload icons for new theme
        execute_icon = self._load_icon_for_button('execute', size=16)
        coding_icon = self._load_icon_for_button('coding', size=16)
        reset_icon = self._load_icon_for_button('reset', size=16)

        # Update primary button (Apply to Model)
        if self.apply_btn:
            self.apply_btn._icon = execute_icon
            self.apply_btn.update_colors(
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg='#ffffff',
                disabled_bg=disabled_bg,
                disabled_fg=disabled_fg,
                canvas_bg=canvas_bg
            )

        # Update secondary buttons (Copy TMDL, Reset All)
        for btn, icon in [(self.copy_btn, coding_icon), (self.reset_btn, reset_icon)]:
            if btn:
                btn._icon = icon
                btn.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    disabled_bg=disabled_bg,
                    disabled_fg=disabled_fg,
                    canvas_bg=canvas_bg
                )

