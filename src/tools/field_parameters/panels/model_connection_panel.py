"""
ModelConnectionPanel
Panel component for the Field Parameters tool.

Built by Reid Havens of Analytic Endeavors
"""

import io
import tkinter as tk
from tkinter import ttk
from typing import Optional, TYPE_CHECKING
import logging
import sys
import threading
from pathlib import Path

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

from core.pbi_connector import get_connector
from core.cloud import CloudWorkspaceBrowser, CloudBrowserDialog
from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, ThemedMessageBox
from core.local_model_cache import get_local_model_cache
from tools.field_parameters.panels.panel_base import SectionPanelMixin


class AnimatedDots(tk.Canvas):
    """Three animated dots for loading/scanning indication."""

    def __init__(self, parent, bg: str, dot_color: str, size: int = 6, spacing: int = 4):
        """
        Create animated dots indicator.

        Args:
            parent: Parent widget
            bg: Background color
            dot_color: Color for the dots
            size: Diameter of each dot in pixels
            spacing: Space between dots in pixels
        """
        width = size * 3 + spacing * 2 + 2  # +2 for edge padding
        super().__init__(parent, width=width, height=size + 1, bg=bg, highlightthickness=0)
        self._bg_color = bg
        self._dot_color = dot_color
        self._dim_color = self._dim_color_from(dot_color, bg)
        self._size = size
        self._spacing = spacing
        self._current = 0
        self._animating = False
        self._after_id = None
        self._dot_ids = []
        self._draw_dots()

    def _dim_color_from(self, color: str, bg: str) -> str:
        """Create a dimmed version of the color by blending toward background."""
        def parse_hex(c):
            if c.startswith('#') and len(c) == 7:
                return int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
            return None

        dot_rgb = parse_hex(color)
        bg_rgb = parse_hex(bg)

        if dot_rgb and bg_rgb:
            factor = 0.35
            r = int(dot_rgb[0] * factor + bg_rgb[0] * (1 - factor))
            g = int(dot_rgb[1] * factor + bg_rgb[1] * (1 - factor))
            b = int(dot_rgb[2] * factor + bg_rgb[2] * (1 - factor))
            return f'#{r:02x}{g:02x}{b:02x}'

        if dot_rgb:
            r, g, b = dot_rgb
            factor = 0.35
            return f'#{int(r*factor):02x}{int(g*factor):02x}{int(b*factor):02x}'

        return '#808080'

    def _draw_dots(self):
        """Draw the three dots."""
        self.delete("all")
        self._dot_ids = []
        for i in range(3):
            x = i * (self._size + self._spacing)
            color = self._dot_color if i == self._current else self._dim_color
            dot_id = self.create_oval(
                x, 0, x + self._size, self._size,
                fill=color, outline=color
            )
            self._dot_ids.append(dot_id)

    def _animate(self):
        """Advance to next dot and schedule next animation frame."""
        if not self._animating:
            return
        self._current = (self._current + 1) % 3
        self._draw_dots()
        self._after_id = self.after(150, self._animate)

    def start(self):
        """Start the animation."""
        if self._animating:
            return
        self._animating = True
        self._current = 0
        self._draw_dots()
        self._after_id = self.after(150, self._animate)

    def stop(self):
        """Stop the animation."""
        self._animating = False
        if self._after_id:
            self.after_cancel(self._after_id)
            self._after_id = None
        self._current = 0
        self._draw_dots()

    def update_colors(self, bg: str, dot_color: str):
        """Update colors for theme changes."""
        self.configure(bg=bg)
        self._bg_color = bg
        self._dot_color = dot_color
        self._dim_color = self._dim_color_from(dot_color, bg)
        self._draw_dots()

if TYPE_CHECKING:
    from tools.field_parameters.field_parameters_ui import FieldParametersTab


class ModelConnectionPanel(SectionPanelMixin, ttk.LabelFrame):
    """Panel for connecting to Power BI models"""

    def __init__(self, parent, main_tab: 'FieldParametersTab'):
        ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe', padding="12")
        self._init_section_panel_mixin()
        self.main_tab = main_tab
        self.logger = logging.getLogger(__name__)

        # Create and set the section header labelwidget
        self._create_section_header(parent, "Model Connection", "Power-BI")

        self.connection_status = tk.StringVar(value="Not Connected")
        self.selected_model = tk.StringVar()

        # Mapping from display name to connection string (server|database)
        self._display_to_connection = {}

        # Track current status state for theme changes
        self._current_status_color = 'error'  # 'error', 'success', 'text_muted'

        # Track buttons for theme updates
        self._primary_buttons = []
        self._secondary_buttons = []

        # Button icons (loaded in setup_ui)
        self._button_icons = {}

        # Cloud workspace browser (lazy initialized)
        self._cloud_browser: Optional[CloudWorkspaceBrowser] = None

        self.setup_ui()

    def _load_button_icons(self) -> dict:
        """Load SVG icons for buttons. Returns dict of icon_name -> PhotoImage."""
        icons = {}
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return icons

        # Path from panels/ -> field_parameters/ -> tools/ -> src/ (4 parents)
        icons_dir = Path(__file__).parent.parent.parent.parent / "assets" / "Tool Icons"

        # Icons to load (16px for buttons)
        icon_names = ["connect", "magnifying-glass", "cloud-computing"]

        # Determine icon variant based on theme
        is_dark = self._theme_manager.is_dark
        suffix = "" if is_dark else "-dark"

        for name in icon_names:
            try:
                # Try themed variant first, fallback to base
                icon_path = icons_dir / f"{name}{suffix}.svg"
                if not icon_path.exists():
                    icon_path = icons_dir / f"{name}.svg"

                if icon_path.exists():
                    # Render SVG to PNG at 16x16
                    png_data = cairosvg.svg2png(
                        url=str(icon_path),
                        output_width=16,
                        output_height=16
                    )
                    img = Image.open(io.BytesIO(png_data))
                    icons[name] = ImageTk.PhotoImage(img)
            except Exception as e:
                self.logger.debug(f"Could not load icon {name}: {e}")

        return icons

    def _apply_combobox_listbox_style(self):
        """Apply theme-aware styling to the combobox dropdown listbox (popup)."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Get appropriate colors for the dropdown list
        list_bg = colors.get('card_surface', '#2a2a3c' if is_dark else '#ffffff')
        list_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#1a1a2e')
        # Theme-aware selection colors - use teal for light mode to match brand
        select_bg = '#1a5a8a' if is_dark else '#009999'
        select_fg = '#ffffff'
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')

        # Apply listbox styling via option_add (affects NEW dropdown popups)
        root = self.winfo_toplevel()
        root.option_add('*TCombobox*Listbox.background', list_bg, 80)
        root.option_add('*TCombobox*Listbox.foreground', list_fg, 80)
        root.option_add('*TCombobox*Listbox.selectBackground', select_bg, 80)
        root.option_add('*TCombobox*Listbox.selectForeground', select_fg, 80)
        root.option_add('*TCombobox*Listbox.font', ('Segoe UI', 9), 80)
        # Remove extra border/relief on the popup
        root.option_add('*TCombobox*Listbox.relief', 'flat', 80)
        root.option_add('*TCombobox*Listbox.borderWidth', 0, 80)
        root.option_add('*TCombobox*Listbox.highlightThickness', 1, 80)
        root.option_add('*TCombobox*Listbox.highlightBackground', border_color, 80)
        root.option_add('*TCombobox*Listbox.highlightColor', border_color, 80)

        # Directly configure existing combobox listbox (for theme changes)
        if hasattr(self, 'model_combo'):
            self._configure_combobox_popdown(
                self.model_combo, list_bg, list_fg, select_bg, select_fg, border_color
            )

    def _configure_combobox_popdown(self, combo, bg, fg, select_bg, select_fg, border_color):
        """Directly configure an existing combobox's popdown listbox."""
        try:
            # Access the combobox's popdown window via Tk internal command
            popdown_path = combo.tk.call('ttk::combobox::PopdownWindow', combo)
            if popdown_path:
                # The listbox is typically at popdown_path.f.l
                listbox_path = f"{popdown_path}.f.l"
                try:
                    combo.tk.call(listbox_path, 'configure',
                                  '-background', bg,
                                  '-foreground', fg,
                                  '-selectbackground', select_bg,
                                  '-selectforeground', select_fg,
                                  '-highlightbackground', border_color,
                                  '-highlightcolor', border_color)
                except Exception:
                    pass  # Listbox may not exist yet (first open not happened)
        except Exception:
            pass  # Popdown window may not exist yet

    def setup_ui(self):
        """Setup connection panel UI"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        # Use 'background' (darkest) for inner content frames, not section_bg
        content_bg = colors['background']
        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Load button icons
        self._button_icons = self._load_button_icons()

        # CRITICAL: Inner wrapper with Section.TFrame style fills LabelFrame with correct background
        # This is required per UI_DESIGN_PATTERNS.md - without it, section_bg shows in padding area
        self._content_wrapper = ttk.Frame(self, style='Section.TFrame', padding="15")
        self._content_wrapper.pack(fill=tk.BOTH, expand=True)

        # Track inner frames for theme updates
        self._inner_frames = []

        # Status display
        status_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        status_frame.pack(fill=tk.X, pady=(0, 8))
        self._inner_frames.append(status_frame)

        self._status_text_label = tk.Label(
            status_frame, text="Status:",
            font=('Segoe UI', 9, 'bold'),
            bg=content_bg, fg=colors['text_primary']
        )
        self._status_text_label.pack(side=tk.LEFT)
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.connection_status,
            bg=content_bg,
            fg=colors['error'],
            font=("Segoe UI", 9, "bold")
        )
        self.status_label.pack(side=tk.LEFT, padx=(5, 0))

        # Model selection
        select_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        select_frame.pack(fill=tk.X, pady=(0, 8))
        self._inner_frames.append(select_frame)

        self._model_label = tk.Label(
            select_frame, text="Model:",
            font=('Segoe UI', 9),
            bg=content_bg, fg=colors['text_primary']
        )
        self._model_label.pack(side=tk.LEFT)
        self.model_combo = ttk.Combobox(
            select_frame,
            textvariable=self.selected_model,
            state="readonly",
            width=40
        )
        self.model_combo.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)

        # Prevent text highlight when dropdown closes
        self.model_combo.bind("<<ComboboxSelected>>", lambda e: self.model_combo.selection_clear())

        # Apply combobox dropdown styling
        self._apply_combobox_listbox_style()

        # Buttons
        button_frame = tk.Frame(self._content_wrapper, bg=content_bg)
        button_frame.pack(fill=tk.X)
        self._inner_frames.append(button_frame)

        # REFRESH button (left) - secondary style
        self.refresh_btn = RoundedButton(
            button_frame,
            text="REFRESH",
            command=self._on_refresh,
            icon=self._button_icons.get('magnifying-glass'),
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
        self.refresh_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._secondary_buttons.append(self.refresh_btn)

        # CONNECT button (middle) - primary style
        self.connect_btn = RoundedButton(
            button_frame,
            text="CONNECT",
            command=self._on_connect,
            icon=self._button_icons.get('connect'),
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            height=32, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=content_bg
        )
        self.connect_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.connect_btn.set_enabled(False)  # Disabled until models are discovered
        self._primary_buttons.append(self.connect_btn)

        # CLOUD button (right) - secondary style
        self.cloud_btn = RoundedButton(
            button_frame,
            text="CLOUD",
            command=self._on_cloud_connect,
            icon=self._button_icons.get('cloud-computing'),
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
        self.cloud_btn.pack(side=tk.LEFT)
        self._secondary_buttons.append(self.cloud_btn)

        # Scanning progress indicator with animated dots
        self.scan_status = tk.StringVar(value="")
        self.scan_status_label = tk.Label(
            button_frame,
            textvariable=self.scan_status,
            bg=content_bg,
            fg=colors['text_muted'],
            font=("Segoe UI", 8, "italic")
        )
        # Don't pack initially - will be shown/hidden dynamically

        # Animated dots indicator
        self._scanning_dots = AnimatedDots(
            button_frame,
            bg=content_bg,
            dot_color=colors.get('info', colors.get('button_primary', colors['text_primary']))
        )
        # Don't pack initially - will be shown/hidden dynamically

        # DON'T auto-discover on startup - only when user clicks Refresh
        # Just initialize with empty dropdown
        self.model_combo['values'] = []
    
    def _refresh_model_list(self, show_popup: bool = True):
        """Refresh available models from external tool context or auto-discover open models"""
        connector = get_connector()
        
        # Check if launched from Power BI (external tool)
        server, database = connector.parse_external_tool_args(sys.argv)
        
        if server and database:
            # Launched from Power BI - auto-connect
            self.logger.info(f"External tool launch detected: {server} / {database}")
            model_string = f"{server}|{database}"
            # Extract port for display name
            port = server.split(':')[-1] if ':' in server else server
            display_name = f"External Tool Model (:{port})"
            
            # Store mapping
            self._display_to_connection = {display_name: model_string}
            
            # Show friendly name in dropdown
            self.model_combo['values'] = [display_name]
            self.model_combo.current(0)
            self.connect_btn.set_enabled(True)  # Enable connect button
            # Auto-connect
            self._on_connect()
        else:
            # Manual mode - discover open models
            self.logger.info("Discovering open Power BI models...")
            discovered = connector.discover_local_models()
            
            if discovered:
                # Create mapping from display names to connection strings
                self._display_to_connection = {}
                display_names = []
                
                for model in discovered:
                    connection_string = f"{model.server}|{model.database_name}"
                    display_name = model.display_name
                    self._display_to_connection[display_name] = connection_string
                    display_names.append(display_name)
                
                # Show friendly display names in dropdown
                self.model_combo['values'] = display_names

                if display_names:
                    self.model_combo.current(0)
                    self.connect_btn.set_enabled(True)  # Enable connect button

                self.logger.info(f"Found {len(discovered)} open model(s)")
            else:
                # No models found
                self.logger.info("No open Power BI models found")
                self.model_combo['values'] = []
                self.connect_btn.set_enabled(False)  # Disable connect button

                # Only show popup if explicitly requested (user clicked refresh)
                if show_popup:
                    ThemedMessageBox.showinfo(
                        self.winfo_toplevel(),
                        "No Models Found",
                        "No open Power BI Desktop models detected.\n\n"
                        "Make sure Power BI Desktop is running with a model open.\n\n"
                        "The tool scans localhost ports 50000-60000 for Analysis Services instances."
                    )
    
    def _show_scanning_progress(self, message: str = "Scanning Connections"):
        """Show the animated dots indicator with a message - right aligned"""
        self.scan_status.set(message)
        # Pack from right: dots first (far right), then text to its left
        self._scanning_dots.pack(side=tk.RIGHT)
        self.scan_status_label.pack(side=tk.RIGHT, padx=(10, 6))
        self._scanning_dots.start()

    def _hide_scanning_progress(self):
        """Hide the animated dots indicator"""
        self._scanning_dots.stop()
        self._scanning_dots.pack_forget()
        self.scan_status_label.pack_forget()
        self.scan_status.set("")

    def populate_from_shared_cache(self):
        """
        Populate dropdown from shared LocalModelCache if available.

        Called on tab activation to instantly show models if already discovered
        by another tool (e.g., Hot Swap tab). Also triggers auto-scan if cache
        is empty or stale.
        """
        cache = get_local_model_cache()

        # If cache has fresh models, populate dropdown immediately
        if not cache.is_empty() and not cache.is_stale():
            models = cache.get_models()
            if models:
                self.logger.info(f"Populating dropdown from shared cache ({len(models)} models)")

                # Remember current selection
                current_selection = self.selected_model.get()

                # Create mapping from display names to connection strings
                self._display_to_connection = {}
                display_names = []

                for model in models:
                    connection_string = f"{model.server}|{model.database_name}"
                    display_name = model.display_name
                    self._display_to_connection[display_name] = connection_string
                    display_names.append(display_name)

                self.model_combo['values'] = display_names

                # Restore previous selection if still available
                if current_selection and current_selection in display_names:
                    idx = display_names.index(current_selection)
                    self.model_combo.current(idx)
                elif display_names:
                    self.model_combo.current(0)

                self.connect_btn.set_enabled(True)
                return

        # Cache is empty or stale - trigger auto-scan if not already in progress
        if not cache.is_scan_in_progress():
            self.logger.info("Shared cache empty/stale, triggering auto-scan")
            self._on_refresh()

    def _on_refresh(self):
        """Refresh model list in background thread to prevent UI freeze"""
        # Disable button during scan and show progress
        self.refresh_btn.config(state="disabled")
        self._show_scanning_progress("Scanning Connections")

        # Get current count for comparison later
        current_values = self.model_combo['values']
        current_count = len(current_values) if current_values else 0

        # Progress callback - we ignore detailed messages and just keep "Scanning Connections"
        # The animated dots provide visual feedback during the scan
        def progress_callback(message: str):
            pass  # Keep showing "Scanning Connections" with animated dots

        # Run discovery in background thread
        def discover_thread():
            try:
                self.logger.info("Starting background model discovery...")

                # Use smart discovery with progress callback
                discovered = get_connector().discover_local_models(
                    quick_scan=True,
                    progress_callback=progress_callback
                )
                
                self.logger.info(f"Discovery returned {len(discovered)} model(s)")
                
                # Update UI in main thread - use a lambda that captures values
                def update_ui():
                    self._discovery_complete(discovered, current_count)
                
                self.after(0, update_ui)
            except Exception as e:
                self.logger.error(f"Discovery thread error: {e}", exc_info=True)
                
                def show_error():
                    self._discovery_error(str(e))
                
                self.after(0, show_error)
        
        thread = threading.Thread(target=discover_thread, daemon=True)
        thread.start()
    
    def _discovery_complete(self, discovered, previous_count):
        """Handle discovery completion (runs in main thread)"""
        self.logger.info(f"_discovery_complete called with {len(discovered)} models")

        # Remember current selection before updating
        current_selection = self.selected_model.get()

        try:
            # Update shared cache so other tools can use these models
            cache = get_local_model_cache()
            cache.set_models(discovered)

            if discovered:
                # Create mapping from display names to connection strings
                self._display_to_connection = {}
                display_names = []

                for model in discovered:
                    connection_string = f"{model.server}|{model.database_name}"
                    display_name = model.display_name
                    self._display_to_connection[display_name] = connection_string
                    display_names.append(display_name)

                self.logger.info(f"Created {len(display_names)} dropdown entries")
                self.logger.info(f"First entry: {display_names[0] if display_names else 'none'}")

                self.model_combo['values'] = display_names
                self.logger.info(f"Updated combobox values")

                if display_names:
                    # Preserve previous selection if still available
                    if current_selection and current_selection in display_names:
                        idx = display_names.index(current_selection)
                        self.model_combo.current(idx)
                    else:
                        self.model_combo.current(0)
                    self.connect_btn.set_enabled(True)  # Enable connect button

                self.logger.info(f"Found {len(discovered)} open model(s)")

                # Update status to show discovered count with animated dots
                self.scan_status.set(f"Discovered {len(discovered)} model(s)")
            else:
                # No models found
                self.logger.info("No open Power BI models found")
                self.model_combo['values'] = []
                self.connect_btn.set_enabled(False)  # Disable connect button

                # Update status to show no models found
                self.scan_status.set("Discovered 0 models")

                # Show helpful message
                ThemedMessageBox.showinfo(
                    self.winfo_toplevel(),
                    "No Models Found",
                    "No open Power BI Desktop models detected.\n\n"
                    "Make sure Power BI Desktop is running with a model open.\n\n"
                    "Or use the 'Manual' button if you know the port number."
                )
        finally:
            # Re-enable button and hide progress
            self.refresh_btn.config(state="normal")
            self._hide_scanning_progress()
    
    def _discovery_error(self, error_msg):
        """Handle discovery error (runs in main thread)"""
        self.logger.error(f"Discovery failed: {error_msg}")
        self.refresh_btn.config(state="normal")
        self._hide_scanning_progress()

        ThemedMessageBox.showerror(
            self.winfo_toplevel(),
            "Discovery Error",
            f"Failed to discover models:\n{error_msg}\n\n"
            "Try using the 'Manual' button instead."
        )
    
    def _on_cloud_connect(self):
        """Open cloud workspace/model browser dialog"""
        # Get connector and check TOM first
        connector = get_connector()

        # Check if TOM is initialized
        if not hasattr(connector, 'Server'):
            ThemedMessageBox.showerror(
                self.winfo_toplevel(),
                "TOM Not Available",
                "Tabular Object Model (TOM) is not initialized.\n\n"
                "This requires Power BI Desktop to be installed.\n\n"
                "The Microsoft.AnalysisServices.Tabular.dll is found in:\n"
                "Power BI Desktop installation folder, SQL Server Management Studio, or Tabular Editor 3.\n\n"
                "Check the log for initialization errors."
            )
            return

        # Lazy initialize cloud browser (shared across dialog opens)
        if self._cloud_browser is None:
            self._cloud_browser = CloudWorkspaceBrowser()

        # Show cloud browser dialog (simple_mode hides perspective/connector options)
        dialog = CloudBrowserDialog(
            self.winfo_toplevel(),
            self._cloud_browser,
            simple_mode=True
        )
        self.wait_window(dialog)  # Wait for user to interact with dialog

        # Check if user selected a model
        result = dialog.result
        if result:
            # result is a SwapTarget with server (XMLA endpoint) and database (model name)
            xmla_endpoint = result.server
            dataset_name = result.database
            self._connect_to_cloud(xmla_endpoint, dataset_name)

    def _connect_local_manual(self, server_address: str, db_name: str):
        """Execute manual local connection"""
        connector = get_connector()

        try:
            self.connection_status.set("Connecting...")
            self._current_status_color = 'text_muted'
            self.status_label.config(foreground=self._theme_manager.colors['text_muted'])
            self.update_idletasks()

            success, message = connector.connect(server_address, db_name)

            if success:
                self.logger.info(f"Connected to {server_address}/{db_name}")

                # Get tables and existing field parameters
                tables = connector.get_all_fields_by_table()
                existing_params = connector.detect_field_parameters()

                self.logger.info(f"Found {len(tables)} tables, {len(existing_params)} field parameters")

                # Get friendly model name from connector (which obtained it via TOM)
                conn_info = connector.get_connection_info()
                model_display_name = conn_info.model_name if conn_info else db_name

                # Update status with friendly name
                self.connection_status.set(f"Connected: {model_display_name}")
                self._current_status_color = 'success'
                self.status_label.config(foreground=self._theme_manager.colors['success'])

                # Update dropdown with friendly name
                port = server_address.split(':')[-1] if ':' in server_address else server_address
                display_name = f"{model_display_name} (:{port})"
                self._display_to_connection[display_name] = f"{server_address}|{db_name}"
                self.model_combo['values'] = [display_name]
                self.model_combo.current(0)

                # Notify main tab with friendly name
                self.main_tab.on_model_connected(model_display_name, tables, existing_params)

                ThemedMessageBox.show(
                    self.winfo_toplevel(),
                    "Connected",
                    f"Successfully connected to {model_display_name}\n\n"
                    f"Tables: {len(tables)}\n"
                    f"Field Parameters: {len(existing_params)}",
                    msg_type="success",
                    custom_icon="Field Parameter.svg"
                )
            else:
                self.connection_status.set("Connection Failed")
                self._current_status_color = 'error'
                self.status_label.config(foreground=self._theme_manager.colors['error'])
                ThemedMessageBox.showerror(self.winfo_toplevel(), "Connection Failed", message)

        except Exception as e:
            self.logger.error(f"Manual connection failed: {e}", exc_info=True)
            self.connection_status.set("Connection Failed")
            self._current_status_color = 'error'
            self.status_label.config(foreground=self._theme_manager.colors['error'])
            ThemedMessageBox.showerror(self.winfo_toplevel(), "Connection Failed", str(e))

    def _connect_to_cloud(self, xmla_endpoint: str, dataset_name: Optional[str] = None):
        """Execute cloud connection"""
        connector = get_connector()

        # Check if TOM is initialized
        if not hasattr(connector, 'Server'):
            ThemedMessageBox.showerror(
                self.winfo_toplevel(),
                "TOM Not Available",
                "Tabular Object Model (TOM) is not initialized.\n\n"
                "Cloud connection requires the same prerequisites as local connection."
            )
            return

        # Update status
        self.connection_status.set("Authenticating...")
        self._current_status_color = 'text_muted'
        self.status_label.config(foreground=self._theme_manager.colors['text_muted'])
        self.update_idletasks()

        # Run connection directly (MSAL interactive auth needs main thread for browser)
        # Use after() to let the UI update first, then run connection
        def do_connect():
            try:
                success, message = connector.connect_xmla(xmla_endpoint, dataset_name)

                if success:
                    self._on_cloud_connection_success(connector, message)
                else:
                    self._on_cloud_connection_failed(message)
            except Exception as e:
                self.logger.error(f"Cloud connection error: {e}", exc_info=True)
                self._on_cloud_connection_failed(str(e))

        # Schedule connection after UI updates
        self.after(100, do_connect)

    def _on_cloud_connection_success(self, connector, message):
        """Handle successful cloud connection"""
        # Get tables and fields
        self.logger.info("Retrieving model metadata from cloud...")
        tables = connector.get_all_fields_by_table()
        existing_params = connector.detect_field_parameters()

        self.logger.info(f"Found {len(tables)} tables, {len(existing_params)} field parameters")

        # Update status
        conn_info = connector.get_connection_info()
        model_name = "Cloud Model"
        if conn_info:
            model_name = conn_info.model_name  # Use attribute access, not dict subscript
            self.connection_status.set(f"Connected: {model_name}")
            self._current_status_color = 'success'
            self.status_label.config(foreground=self._theme_manager.colors['success'])

            # Update dropdown to show connected model
            # Note: model_name from conn_info already includes " (Cloud)" suffix from connector
            display_name = model_name
            self._display_to_connection[display_name] = f"{conn_info.server}|{conn_info.database}"
            self.model_combo['values'] = [display_name]
            self.model_combo.current(0)

        # Notify main tab (with model_name as first param)
        self.main_tab.on_model_connected(model_name, tables, existing_params)

        ThemedMessageBox.show(
            self.winfo_toplevel(),
            "Cloud Connected",
            f"{message}\n\n"
            f"Tables: {len(tables)}\n"
            f"Field Parameters: {len(existing_params)}\n\n"
            "Changes will be applied directly to the published dataset.",
            msg_type="success",
            custom_icon="Field Parameter.svg"
        )

    def _on_cloud_connection_failed(self, error_message):
        """Handle failed cloud connection"""
        self.connection_status.set("Connection Failed")
        self._current_status_color = 'error'
        self.status_label.config(foreground=self._theme_manager.colors['error'])

        ThemedMessageBox.showerror(
            self.winfo_toplevel(),
            "Cloud Connection Failed",
            f"{error_message}\n\n"
            "Requirements:\n"
            "Power BI Premium or PPU capacity, XMLA read-write enabled, Contributor or higher permissions, msal library installed."
        )

    def _on_connect(self):
        """Connect to selected model"""
        display_name = self.selected_model.get()
        if not display_name:
            ThemedMessageBox.showwarning(self.winfo_toplevel(), "No Model", "Please select a model first")
            return
        
        try:
            # Look up connection string from display name
            model_info = self._display_to_connection.get(display_name)
            if not model_info:
                # Fallback: treat as direct connection string (backwards compatibility)
                model_info = display_name
            
            # Parse connection info (format: "server|database")
            if "|" in model_info:
                server, database = model_info.split("|")
            else:
                ThemedMessageBox.showerror(self.winfo_toplevel(), "Invalid Format", "Model format should be: server|database")
                return

            self.logger.info(f"Connecting to server: {server}, database: {database}")

            # Get connector and connect
            connector = get_connector()

            if not connector.connect(server, database):
                ThemedMessageBox.showerror(
                    self.winfo_toplevel(),
                    "Connection Failed",
                    "Could not connect to Power BI model.\n\n"
                    "Make sure Power BI Desktop is running and the model is open.\n"
                    "Try clicking 'Refresh' to re-scan for models."
                )
                return

            # Get tables and fields from model
            self.logger.info("Retrieving model metadata...")
            tables = connector.get_all_fields_by_table()
            existing_params = connector.detect_field_parameters()

            self.logger.info(f"Found {len(tables)} tables, {len(existing_params)} field parameters")

            # Update status
            conn_info = connector.get_connection_info()
            self.connection_status.set(f"Connected ({conn_info.model_name})")
            self._current_status_color = 'success'
            self.status_label.config(foreground=self._theme_manager.colors['success'])

            # Update dropdown with actual model name (for external tool launch)
            if conn_info.model_name and "External Tool Model" in display_name:
                port = server.split(':')[-1] if ':' in server else server
                new_display_name = f"{conn_info.model_name} ({port})"
                # Update the dropdown
                self._display_to_connection[new_display_name] = f"{server}|{database}"
                self.model_combo['values'] = [new_display_name]
                self.model_combo.current(0)
                self.selected_model.set(new_display_name)
                # Remove old generic entry
                if display_name in self._display_to_connection and display_name != new_display_name:
                    del self._display_to_connection[display_name]

            # Notify main tab
            self.main_tab.on_model_connected(conn_info.model_name, tables, existing_params)

            # Show success dialog
            ThemedMessageBox.show(
                self.winfo_toplevel(),
                "Connected",
                f"Successfully connected to {conn_info.model_name}\n\n"
                f"Tables: {len(tables)}\n"
                f"Field Parameters: {len(existing_params)}",
                msg_type="success",
                custom_icon="Field Parameter.svg"
            )

        except Exception as e:
            self.logger.error(f"Connection error: {e}", exc_info=True)
            ThemedMessageBox.showerror(self.winfo_toplevel(), "Connection Error", f"Failed to connect to model:\n{e}")

    def on_theme_changed(self):
        """Update widget colors when theme changes"""
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'
        # Use 'section_bg' for inner content areas (matches Section.TFrame style)
        content_bg = colors['background']
        # Theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Reload icons for new theme
        self._button_icons = self._load_button_icons()

        # Update button icons before color update
        if hasattr(self, 'connect_btn'):
            self.connect_btn._icon = self._button_icons.get('connect')
        if hasattr(self, 'refresh_btn'):
            self.refresh_btn._icon = self._button_icons.get('magnifying-glass')
        if hasattr(self, 'cloud_btn'):
            self.cloud_btn._icon = self._button_icons.get('cloud-computing')

        # Update combobox dropdown styling for new theme
        self._apply_combobox_listbox_style()

        # Update inner frames
        for frame in self._inner_frames:
            frame.config(bg=content_bg)

        # Update status text label
        self._status_text_label.config(bg=content_bg, fg=colors['text_primary'])

        # Update status label based on current status state
        status_color_map = {
            'error': colors['error'],
            'success': colors['success'],
            'text_muted': colors['text_muted']
        }
        current_color = status_color_map.get(self._current_status_color, colors['error'])
        self.status_label.config(bg=content_bg, fg=current_color)

        # Update model label
        self._model_label.config(bg=content_bg, fg=colors['text_primary'])

        # Update scan status label
        self.scan_status_label.config(bg=content_bg, fg=colors['text_muted'])

        # Force ttk.Frame to re-apply style after theme change
        if hasattr(self, '_content_wrapper'):
            self._content_wrapper.configure(style='Section.TFrame')

        # Update section header widgets
        self._update_section_header_theme()

        # Update primary buttons (update_colors triggers redraw with new icons)
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

        # Update animated dots indicator for theme
        if hasattr(self, '_scanning_dots') and self._scanning_dots:
            try:
                self._scanning_dots.update_colors(
                    bg=content_bg,
                    dot_color=colors.get('info', '#3b82f6')
                )
            except Exception:
                pass


