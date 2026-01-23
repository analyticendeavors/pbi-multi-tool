"""
Connection Hot-Swap Tab - Main UI Tab
Built by Reid Havens of Analytic Endeavors

Main UI tab that coordinates all panels for the Connection Hot-Swap tool.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
import threading
from pathlib import Path
from typing import Optional, List, Tuple

from core.ui_base import BaseToolTab, RoundedButton, ThemedScrollbar, SquareIconButton, Tooltip, ThemedMessageBox, SplitLogSection
from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from core.pbi_connector import get_connector
from tools.connection_hotswap.models import (
    ModelConnectionInfo,
    ConnectionMapping,
    SwapStatus,
    SwapPreset,
    PresetStorageType,
    PresetScope,
    SwapHistoryEntry,
    SwapTarget,
    DataSourceConnection,
    ConnectionType,
    TomReferenceType,
)
from tools.connection_hotswap.logic.connection_detector import ConnectionDetector
from tools.connection_hotswap.logic.connection_swapper import ConnectionSwapper
from tools.connection_hotswap.logic.local_model_matcher import LocalModelMatcher
from core.cloud import CloudWorkspaceBrowser
from tools.connection_hotswap.logic.preset_manager import PresetManager
from tools.connection_hotswap.logic.health_checker import (
    ConnectionHealthChecker,
    HealthStatus,
    HealthCheckResult,
    get_status_emoji,
)
from tools.connection_hotswap.ui.connection_diagram import ConnectionDiagram
from tools.connection_hotswap.ui.components.inline_target_picker import InlineTargetPicker
from tools.connection_hotswap.ui.dialogs.thin_report_dialog import ThinReportSwapDialog
from tools.field_parameters.models import AvailableModel
from tools.connection_hotswap.logic.schema_validator import (
    SchemaValidator,
    ValidationResult,
    ValidationSeverity,
)
from tools.connection_hotswap.logic.pbix_modifier import get_modifier as get_pbix_modifier
from tools.connection_hotswap.logic.process_control import get_controller as get_process_controller
from core.local_model_cache import get_local_model_cache


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
        # Parse both colors
        def parse_hex(c):
            if c.startswith('#') and len(c) == 7:
                return int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
            return None

        dot_rgb = parse_hex(color)
        bg_rgb = parse_hex(bg)

        if dot_rgb and bg_rgb:
            # Blend 65% toward background (35% dot color, 65% background)
            factor = 0.35
            r = int(dot_rgb[0] * factor + bg_rgb[0] * (1 - factor))
            g = int(dot_rgb[1] * factor + bg_rgb[1] * (1 - factor))
            b = int(dot_rgb[2] * factor + bg_rgb[2] * (1 - factor))
            return f'#{r:02x}{g:02x}{b:02x}'

        # Fallback: just reduce brightness
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
            # Highlight current dot, dim others
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


class ConnectionHotswapTab(BaseToolTab):
    """
    Main UI tab for Connection Hot-Swap tool.

    Coordinates model connection, connection detection, mapping configuration,
    and swap execution.
    """

    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "connection_hotswap", "Connection Hot-Swap")
        self.logger = logging.getLogger(__name__)
        self._theme_manager = get_theme_manager()

        # State
        self.model_info: Optional[ModelConnectionInfo] = None
        self.mappings: List[ConnectionMapping] = []
        self.auto_match_enabled = tk.BooleanVar(value=True)
        self._thin_report_context: Optional[dict] = None  # Context for thin report file modification
        self._current_model_file_path: Optional[str] = None  # File path for thin report modification

        # Logic components
        self.connector = get_connector()
        self.detector: Optional[ConnectionDetector] = None
        self.swapper: Optional[ConnectionSwapper] = None
        self.matcher: Optional[LocalModelMatcher] = None
        self.cloud_browser: Optional[CloudWorkspaceBrowser] = None
        self.preset_manager = PresetManager()
        self.health_checker: Optional[ConnectionHealthChecker] = None
        self.schema_validator: Optional[SchemaValidator] = None

        # Health status tracking
        self._health_statuses: dict = {}  # target_id -> HealthCheckResult

        # View mode (table or diagram)
        self._view_mode = tk.StringVar(value="table")  # "table" or "diagram"

        # Preset state
        self._preset_buttons: List[RoundedButton] = []
        self._current_presets: List[SwapPreset] = []
        self._preset_scope_filter: PresetScope = PresetScope.GLOBAL  # Current scope filter (Global/Model toggle) - defaults to GLOBAL

        # UI component references
        self.connection_status = tk.StringVar(value="Not Connected")
        self.selected_model = tk.StringVar()
        self._display_to_connection = {}

        self._primary_buttons = []

        # Load icons
        self._load_button_icons()

        self.setup_ui()
        self.logger.info("Connection Hot-Swap tab initialized")

        # Check for external tool launch and auto-connect after UI is ready
        self.frame.after(100, self._check_external_tool_launch)

    def _check_external_tool_launch(self):
        """Check if launched from Power BI external tool and auto-connect."""
        import sys
        server, database = self.connector.parse_external_tool_args(sys.argv)

        if not server:
            return

        # Check if server is a cloud endpoint (thin report launched from External Tools)
        # Power BI passes cloud endpoint like "pbiazure://api.powerbi.com" for thin reports
        is_cloud_endpoint = 'powerbi.com' in server.lower() or server.startswith('pbiazure://')

        if is_cloud_endpoint:
            # Thin report - Power BI passed cloud endpoint, need to discover local model
            self.logger.info(f"External tool launch detected (thin report via cloud endpoint): {server}")
            self._log_message("Auto-connecting to Power BI Desktop model...")
            self._log_message("Detecting thin report...")

            # Show progress indicator during discovery
            self._show_progress("Connecting...")

            cloud_server = server
            cloud_database = database

            def discover_thin_report():
                try:
                    self.frame.after(0, lambda: self._update_progress_message("Scanning for models..."))

                    # Discover local models to find the thin report
                    discovered = self.connector.discover_local_models()

                    self.logger.info(f"Looking for thin report with cloud_database: {cloud_database}")

                    # Find thin report matching the cloud database GUID
                    matching_model = None
                    for model in discovered:
                        if model.is_thin_report:
                            # Get the cloud database from the model (correct attribute name)
                            model_cloud_db = getattr(model, 'thin_report_cloud_database', None)
                            self.logger.info(f"  Thin report '{model.display_name}': cloud_db={model_cloud_db}")

                            # Match by database GUID if available
                            if cloud_database and model_cloud_db:
                                # Direct match
                                if model_cloud_db == cloud_database:
                                    matching_model = model
                                    self.logger.info(f"  -> MATCH by exact cloud_database")
                                    break
                                # Also check if GUID is embedded in database_name (format: __thin_report__:server:database)
                                if cloud_database in model.database_name:
                                    matching_model = model
                                    self.logger.info(f"  -> MATCH by database_name contains GUID")
                                    break

                    # Fallback: if no cloud match, just take the first thin report
                    if not matching_model:
                        self.logger.info("No exact match found, using first thin report as fallback")
                        for model in discovered:
                            if model.is_thin_report:
                                matching_model = model
                                break

                    if matching_model:
                        display_name = matching_model.display_name
                        local_server = matching_model.server
                        db_name = matching_model.database_name

                        self._display_to_connection = {display_name: f"{local_server}|{db_name}"}
                        self._dropdown_models_cache = discovered

                        def update_ui_and_connect():
                            self.model_combo['values'] = [display_name]
                            self.model_combo.current(0)
                            self.selected_model.set(display_name)
                            self._on_connect()

                        self.frame.after(0, update_ui_and_connect)
                    else:
                        self.frame.after(0, lambda: self._hide_progress())
                        self.frame.after(0, lambda: self._log_message(
                            "Could not detect thin report. Try clicking Refresh."))

                except Exception as e:
                    self.logger.error(f"Error discovering thin report: {e}")
                    self.frame.after(0, lambda: self._hide_progress())
                    self.frame.after(0, lambda: self._log_message(f"Error: {e}"))

            # Run discovery in background thread
            threading.Thread(target=discover_thin_report, daemon=True).start()

        elif server and database:
            # Standard composite model with localhost:port and database GUID
            self.logger.info(f"External tool launch detected: {server} / {database}")
            port = server.split(':')[-1] if ':' in server else server
            display_name = f"External Tool Model ({port})"

            # Store mapping
            self._display_to_connection = {display_name: f"{server}|{database}"}
            self._dropdown_models_cache = []

            # Update dropdown
            self.model_combo['values'] = [display_name]
            self.model_combo.current(0)
            self.selected_model.set(display_name)

            # Auto-connect with progress indicator
            self._log_message("Auto-connecting to Power BI Desktop model...")
            self._show_progress("Connecting...")
            self._on_connect()

        elif server:
            # Server only (no database) - localhost port without database (edge case)
            self.logger.info(f"External tool launch detected (server only): {server}")
            self._log_message("Auto-connecting to Power BI Desktop model...")

            port = server.split(':')[-1] if ':' in server else server
            self._log_message(f"Detecting model type on port {port}...")

            # Show progress indicator during discovery
            self._show_progress("Connecting...")

            def discover_and_connect():
                try:
                    self.frame.after(0, lambda: self._update_progress_message("Detecting model..."))

                    # Discover models to find the model on this port
                    discovered = self.connector.discover_local_models()

                    # Find model matching our port
                    matching_model = None
                    for model in discovered:
                        if model.server == server or model.server.endswith(f":{port}"):
                            matching_model = model
                            break

                    if matching_model:
                        display_name = matching_model.display_name
                        db_name = matching_model.database_name

                        self._display_to_connection = {display_name: f"{matching_model.server}|{db_name}"}
                        self._dropdown_models_cache = discovered

                        def update_ui_and_connect():
                            self.model_combo['values'] = [display_name]
                            self.model_combo.current(0)
                            self.selected_model.set(display_name)
                            self._on_connect()

                        self.frame.after(0, update_ui_and_connect)
                    else:
                        self.frame.after(0, lambda: self._hide_progress())
                        self.frame.after(0, lambda: self._log_message(
                            f"Could not detect model on port {port}. Try clicking Refresh."))

                except Exception as e:
                    self.logger.error(f"Error discovering model: {e}")
                    self.frame.after(0, lambda: self._hide_progress())
                    self.frame.after(0, lambda: self._log_message(f"Error: {e}"))

            # Run discovery in background thread
            threading.Thread(target=discover_and_connect, daemon=True).start()

    def _load_button_icons(self):
        """Load SVG icons for buttons and section headers"""
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}

        # 16px icons for buttons and section headers
        icon_names_16 = [
            "connection", "execute", "reset", "magnifying-glass",
            "Power-BI", "earth", "desktop", "analyze", "question",
            "hotswap", "save", "eye", "cogwheel", "link", "link alt", "export", "folder",
            "connect", "disconnect", "eraser",
            "letter-c", "letter-l",  # For context menu L/C indicators
            "cloud login",  # For cloud authentication button
            "cloud-computing"  # For cloud model button
        ]
        for name in icon_names_16:
            icon = self._load_icon_for_button(name, size=16)
            if icon:
                self._button_icons[name] = icon

        # Load checkbox icons for modern toggle
        self._load_checkbox_icons()

        # Load cloud auth icons (colored and grayscale)
        self._load_cloud_auth_icons()

    def _load_checkbox_icons(self):
        """Load themed checkbox SVG icons for checked and unchecked states."""
        is_dark = self._theme_manager.is_dark

        # Use theme-aware checkbox icons (box/box-dark pattern)
        box_name = 'box-dark' if is_dark else 'box'
        checked_name = 'box-checked-dark' if is_dark else 'box-checked'

        self._checkbox_off_icon = self._load_icon_for_button(box_name, size=16)
        self._checkbox_on_icon = self._load_icon_for_button(checked_name, size=16)

    def _load_cloud_auth_icons(self):
        """Load cloud auth icons - colored for signed-in, red for signed-out."""
        icons_dir = Path(__file__).parent.parent.parent.parent / "assets" / "Tool Icons"

        # Load signed-in icon (colored/green)
        self._cloud_icon_colored = self._load_icon_for_button("cloud login (signed in)", size=16)
        if not self._cloud_icon_colored:
            # Fall back to regular cloud login icon
            self._cloud_icon_colored = self._button_icons.get('cloud login')

        # Load signed-out icon (red version)
        self._cloud_icon_gray = self._load_icon_for_button("cloud login (signed out)", size=16)
        if not self._cloud_icon_gray:
            # Fall back to regular cloud login icon
            self._cloud_icon_gray = self._button_icons.get('cloud login')

    def _apply_combobox_listbox_style(self):
        """Apply theme-aware styling to the combobox dropdown listbox (popup) and entry."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Get appropriate colors for the dropdown list
        list_bg = colors.get('card_surface', '#2a2a3c' if is_dark else '#ffffff')
        list_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#1a1a2e')
        # Theme-aware selection colors: blue in dark mode, teal in light mode
        select_bg = '#1a5a8a' if is_dark else '#009999'
        select_fg = '#ffffff'

        # Apply listbox styling via option_add (affects the dropdown popup)
        # Note: These must be applied to the root window or parent
        # Use priority 80 (interactive) to override existing settings on theme change
        root = self.frame.winfo_toplevel()
        root.option_add('*TCombobox*Listbox.background', list_bg, 80)
        root.option_add('*TCombobox*Listbox.foreground', list_fg, 80)
        root.option_add('*TCombobox*Listbox.selectBackground', select_bg, 80)
        root.option_add('*TCombobox*Listbox.selectForeground', select_fg, 80)
        root.option_add('*TCombobox*Listbox.font', ('Segoe UI', 9), 80)
        # Remove extra border/relief on the popup
        root.option_add('*TCombobox*Listbox.relief', 'flat', 80)
        root.option_add('*TCombobox*Listbox.borderWidth', 0, 80)
        root.option_add('*TCombobox*Listbox.highlightThickness', 1, 80)
        root.option_add('*TCombobox*Listbox.highlightBackground', colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'), 80)
        root.option_add('*TCombobox*Listbox.highlightColor', colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'), 80)

        # Store colors for direct application to popdown when it opens
        self._combobox_list_colors = {
            'bg': list_bg,
            'fg': list_fg,
            'select_bg': select_bg,
            'select_fg': select_fg,
            'border': colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
        }

        # Destroy any cached popdown so it gets recreated with new theme colors
        if hasattr(self, 'model_combo'):
            self._destroy_combobox_popdown()
            # Force update the combobox entry text color by toggling state
            self._force_update_combobox_text_color()

    def _destroy_combobox_popdown(self):
        """Destroy cached combobox popdown so it gets recreated with new theme colors."""
        if not hasattr(self, 'model_combo'):
            return

        try:
            # Get the popdown window using internal ttk method
            popdown = self.model_combo.tk.call('ttk::combobox::PopdownWindow', self.model_combo)
            if popdown:
                # Destroy the cached popdown - next dropdown open will create a fresh one
                self.model_combo.tk.call('destroy', popdown)
                self.logger.debug("Destroyed combobox popdown for theme refresh")
        except Exception:
            pass  # Ignore errors if popdown doesn't exist yet

    def _force_update_combobox_text_color(self):
        """Force update the combobox entry text color after theme change."""
        if not hasattr(self, 'model_combo'):
            return

        try:
            # Toggle state to force ttk style re-evaluation
            current_state = self.model_combo.cget('state')
            # Temporarily change state and back to force style refresh
            self.model_combo.configure(state='disabled')
            self.model_combo.update_idletasks()
            self.model_combo.configure(state=current_state)
            self.model_combo.update_idletasks()
        except Exception:
            pass

    def _create_auto_match_toggle(self, parent: tk.Frame, bg_color: str, inline: bool = False):
        """Create a modern checkbox toggle for auto-match with SVG icons."""
        colors = self._theme_manager.colors

        # Container frame for toggle
        toggle_frame = tk.Frame(parent, bg=bg_color)
        if inline:
            # Pack inline (side=LEFT) when on same row as other elements
            toggle_frame.pack(side=tk.LEFT, padx=(20, 0))
        else:
            toggle_frame.pack(anchor=tk.W, pady=(0, 15))

        # Checkbox icon (clickable)
        is_checked = self.auto_match_enabled.get()
        icon = self._checkbox_on_icon if is_checked else self._checkbox_off_icon

        self._auto_match_icon_label = tk.Label(toggle_frame, bg=bg_color, cursor='hand2')
        if icon:
            self._auto_match_icon_label.configure(image=icon)
            self._auto_match_icon_label._icon_ref = icon
        else:
            # Fallback if icons not loaded
            self._auto_match_icon_label.configure(
                text="[x]" if is_checked else "[ ]",
                font=('Segoe UI', 9)
            )
        self._auto_match_icon_label.pack(side=tk.LEFT, padx=(0, 6))

        # Text label
        self._auto_match_text_label = tk.Label(
            toggle_frame,
            text="Auto-match local models",
            bg=bg_color,
            fg=colors['text_primary'],
            font=('Segoe UI', 9),
            cursor='hand2'
        )
        self._auto_match_text_label.pack(side=tk.LEFT)

        # Store frame reference for theme updates
        self._auto_match_toggle_frame = toggle_frame
        self._auto_match_toggle_bg = bg_color

        # Bind clicks to toggle
        def on_toggle_click(event=None):
            new_value = not self.auto_match_enabled.get()
            self.auto_match_enabled.set(new_value)
            self._update_auto_match_toggle()
            self._on_auto_match_changed()

        self._auto_match_icon_label.bind('<Button-1>', on_toggle_click)
        self._auto_match_text_label.bind('<Button-1>', on_toggle_click)

    def _update_auto_match_toggle(self):
        """Update the auto-match toggle icon based on current state."""
        is_checked = self.auto_match_enabled.get()
        icon = self._checkbox_on_icon if is_checked else self._checkbox_off_icon

        if icon and hasattr(self, '_auto_match_icon_label'):
            self._auto_match_icon_label.configure(image=icon)
            self._auto_match_icon_label._icon_ref = icon
        elif hasattr(self, '_auto_match_icon_label'):
            self._auto_match_icon_label.configure(text="[x]" if is_checked else "[ ]")

    def _create_preset_scope_toggle(self, parent: tk.Frame, bg_color: str):
        """Create a tab-style toggle between Global and Model preset scopes."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Container frame for the toggle - pack on LEFT side
        toggle_frame = tk.Frame(parent, bg=bg_color)
        toggle_frame.pack(side=tk.LEFT)

        # Tab container (no border - buttons provide their own styling)
        tab_container = tk.Frame(
            toggle_frame,
            bg=bg_color
        )
        tab_container.pack(side=tk.LEFT)

        # Tab colors - match CONNECT button using button_primary
        active_bg = colors.get('button_primary', '#0078d4' if is_dark else '#009999')
        inactive_bg = colors.get('surface', '#1e1e2e' if is_dark else '#f8f8fc')
        active_fg = '#ffffff'
        inactive_fg = colors.get('text_secondary', '#888888')
        hover_bg = colors.get('hover', '#2a2a3c' if is_dark else '#e8e8f0')
        primary_hover = colors.get('button_primary_hover', '#005a9e' if is_dark else '#007A7A')

        # Global tab button (left side - rounded top-left corner only) - ACTIVE by default
        self._scope_global_btn = RoundedButton(
            tab_container,
            text="GLOBAL",
            command=lambda: self._on_preset_scope_change(PresetScope.GLOBAL),
            bg=active_bg,
            fg=active_fg,
            hover_bg=primary_hover,
            pressed_bg=primary_hover,
            width=80, height=28, radius=5,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=active_bg,
            corners='top-left'
        )
        self._scope_global_btn.pack(side=tk.LEFT)

        # Add tooltip to GLOBAL button
        from core.ui_base import Tooltip
        Tooltip(
            self._scope_global_btn,
            "Universal presets - Apply to any single-connection model.\n"
            "Use for environment switching (e.g., Production, Development).",
            delay=400
        )

        # Model tab button (right side - rounded top-right corner only) - INACTIVE by default
        self._scope_model_btn = RoundedButton(
            tab_container,
            text="MODEL",
            command=lambda: self._on_preset_scope_change(PresetScope.MODEL),
            bg=inactive_bg,
            fg=inactive_fg,
            hover_bg=hover_bg,
            pressed_bg=hover_bg,
            width=80, height=28, radius=5,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=inactive_bg,
            corners='top-right'
        )
        self._scope_model_btn.pack(side=tk.LEFT)

        # Add tooltip to MODEL button
        Tooltip(
            self._scope_model_btn,
            "Model-specific presets - Only for this exact model file.\n"
            "Maps each connection by name to its saved target.\n"
            "Ideal for composite models: 'All Local', 'All Dev', 'All Prod', etc.",
            delay=400
        )

        # Store references for theme updates
        self._preset_scope_toggle_frame = toggle_frame
        self._preset_scope_tab_container = tab_container

        # Update visual state
        self._update_preset_scope_toggle()

    def _on_preset_scope_change(self, scope: PresetScope):
        """Handle preset scope toggle change."""
        if self._preset_scope_filter != scope:
            self._preset_scope_filter = scope
            self._update_preset_scope_toggle()
            self._refresh_preset_table()

    def _update_preset_scope_toggle(self):
        """Update the visual state of the scope tab buttons."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Use button_primary colors for consistency with CONNECT button
        active_bg = colors.get('button_primary', '#0078d4' if is_dark else '#009999')
        primary_hover = colors.get('button_primary_hover', '#005a9e' if is_dark else '#007A7A')
        # Pressed color: darker than hover - blue in dark mode, darker teal in light mode
        primary_pressed = colors.get('button_primary_pressed', '#004578' if is_dark else '#006666')
        inactive_bg = colors.get('surface', '#1e1e2e' if is_dark else '#f8f8fc')
        active_fg = '#ffffff'
        inactive_fg = colors.get('text_secondary', '#888888')
        hover_bg = colors.get('hover', '#2a2a3c' if is_dark else '#e8e8f0')
        # Use section_bg for parent background (inactive button canvas_bg)
        section_bg = colors.get('section_bg', '#1a1a2a' if is_dark else '#f5f5fa')

        is_global = self._preset_scope_filter == PresetScope.GLOBAL

        if hasattr(self, '_scope_global_btn'):
            if is_global:
                self._scope_global_btn.bg_normal = active_bg
                self._scope_global_btn.fg = active_fg
                self._scope_global_btn._current_bg = active_bg
                self._scope_global_btn.bg_hover = primary_hover
                self._scope_global_btn.bg_pressed = primary_pressed
                self._scope_global_btn.canvas_bg = section_bg  # Use section_bg for canvas
            else:
                self._scope_global_btn.bg_normal = inactive_bg
                self._scope_global_btn.fg = inactive_fg
                self._scope_global_btn._current_bg = inactive_bg
                self._scope_global_btn.bg_hover = hover_bg
                self._scope_global_btn.bg_pressed = hover_bg
                self._scope_global_btn.canvas_bg = section_bg  # Use section_bg for canvas
            self._scope_global_btn._draw_button()

        if hasattr(self, '_scope_model_btn'):
            if not is_global:
                self._scope_model_btn.bg_normal = active_bg
                self._scope_model_btn.fg = active_fg
                self._scope_model_btn._current_bg = active_bg
                self._scope_model_btn.bg_hover = primary_hover
                self._scope_model_btn.bg_pressed = primary_pressed
                self._scope_model_btn.canvas_bg = section_bg  # Use section_bg for canvas
            else:
                self._scope_model_btn.bg_normal = inactive_bg
                self._scope_model_btn.fg = inactive_fg
                self._scope_model_btn._current_bg = inactive_bg
                self._scope_model_btn.bg_hover = hover_bg
                self._scope_model_btn.bg_pressed = hover_bg
                self._scope_model_btn.canvas_bg = section_bg  # Use section_bg for canvas
            self._scope_model_btn._draw_button()

    def _update_preset_scope_button_counts(self):
        """Update the preset scope buttons with current preset counts."""
        import tkinter.font as tkfont

        if not hasattr(self, '_scope_global_btn') or not hasattr(self, '_scope_model_btn'):
            return

        # Get global preset count
        global_presets = self.preset_manager.list_presets(scope=PresetScope.GLOBAL)
        global_count = len(global_presets)

        # Get model preset count (only if connected to a model)
        model_hash = self._get_model_hash()
        if model_hash:
            model_presets = self.preset_manager.list_presets(
                scope=PresetScope.MODEL,
                model_hash=model_hash
            )
            model_count = len(model_presets)
        else:
            model_count = 0

        # Create new button text with counts
        global_text = f"GLOBAL ({global_count})"
        model_text = f"MODEL ({model_count})"

        # Calculate new widths needed (use the button's font)
        font_obj = tkfont.Font(font=self._scope_global_btn.btn_font)
        padding = 24  # Horizontal padding for button content

        global_width = font_obj.measure(global_text) + padding
        model_width = font_obj.measure(model_text) + padding

        # Ensure minimum width
        min_width = 80
        global_width = max(global_width, min_width)
        model_width = max(model_width, min_width)

        # Update global button
        self._scope_global_btn.text = global_text
        self._scope_global_btn.config(width=global_width)
        self._scope_global_btn._draw_button()

        # Update model button
        self._scope_model_btn.text = model_text
        self._scope_model_btn.config(width=model_width)
        self._scope_model_btn._draw_button()

    def _create_action_buttons(self):
        """Create action buttons at the very bottom of the tab"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Action frame at bottom - pack with side=BOTTOM first
        action_frame = ttk.Frame(self.frame)
        action_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(15, 0), padx=10)

        # Center container for buttons
        button_container = ttk.Frame(action_frame)
        button_container.pack(anchor=tk.CENTER)

        # Bottom buttons sit on main outer background
        outer_canvas_bg = colors.get('outer_bg', '#161627' if is_dark else '#f5f5f7')

        # Disabled state colors (use standard keys from constants.py)
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Swap Selected button (primary action) - with hotswap icon
        hotswap_icon = self._button_icons.get('hotswap')
        self.swap_selected_btn = RoundedButton(
            button_container, text="SWAP SELECTED",
            command=self._on_swap_selected,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=hotswap_icon, canvas_bg=outer_canvas_bg
        )
        self.swap_selected_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.swap_selected_btn.set_enabled(False)
        self._primary_buttons.append(self.swap_selected_btn)
        Tooltip(self.swap_selected_btn, "Apply target connections to selected mappings")

        # Rollback button - uses button_secondary colors to match Reset All buttons
        reset_icon = self._button_icons.get('reset')
        self.rollback_btn = RoundedButton(
            button_container, text="ROLLBACK",
            command=self._on_rollback_last,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            disabled_bg=disabled_bg, disabled_fg=disabled_fg,
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=reset_icon, canvas_bg=outer_canvas_bg
        )
        self.rollback_btn.pack(side=tk.LEFT)
        self.rollback_btn.set_enabled(False)
        Tooltip(self.rollback_btn, "Restore previous connections")

    def setup_ui(self):
        """Setup the main UI layout"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # IMPORTANT: Create action buttons FIRST with side=BOTTOM to ensure they're always visible
        self._create_action_buttons()

        # Create main container with consistent padding
        main_container = ttk.Frame(self.frame)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

        # TOP SECTION: Model Connection (full width)
        self._create_connection_section(main_container)

        # MIDDLE SECTION: Swap Configuration (master section containing Quick Swap & Connections)
        self._create_swap_configuration_section(main_container)

        # BOTTOM SECTION: Activity log
        self._create_log_section(main_container)

        # Initial state
        self._set_initial_state()

    def _create_connection_section(self, parent):
        """Create the model connection section (full width at top) - single row with auto-match below"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create section with icon labelwidget
        header_widget = self.create_section_header(self.frame, "Model Connection", "Power-BI")[0]

        conn_frame = ttk.LabelFrame(
            parent,
            labelwidget=header_widget,
            style='Section.TLabelframe',
            padding="12"
        )
        conn_frame.pack(fill=tk.X, pady=(0, 10))

        # Inner content frame FIRST (for proper stacking order with help button)
        # Use distinct content background: white for light, very dark navy for dark
        content_bg = '#0d0d1a' if is_dark else '#ffffff'
        inner_frame = tk.Frame(conn_frame, bg=content_bg, padx=15, pady=15)
        inner_frame.pack(fill=tk.BOTH, expand=True)
        self._conn_inner_frame = inner_frame  # Store for theme updates

        # Help button AFTER content frame to ensure proper stacking order (like Advanced Copy)
        help_icon = self._button_icons.get('question')
        if help_icon:
            self._help_button = SquareIconButton(
                conn_frame, icon=help_icon, command=self.show_help_dialog,
                tooltip_text="Help", size=26, radius=6,
                bg_normal_override=AppConstants.CORNER_ICON_BG
            )
            # Position in upper-right corner of section title bar area
            self._help_button.place(relx=1.0, y=-35, anchor=tk.NE, x=0)

        # Cloud authentication button (to the left of help button)
        # Shows cloud auth state and provides sign in/out dropdown
        initial_cloud_icon = self._cloud_icon_gray  # Start as signed out
        if hasattr(self, '_cloud_icon_colored') and self._cloud_icon_colored:
            self._cloud_auth_btn = SquareIconButton(
                conn_frame,
                icon=initial_cloud_icon,
                command=self._toggle_cloud_auth_dropdown,
                tooltip_text="Cloud Account",
                size=26,
                radius=6,
                bg_normal_override=AppConstants.CORNER_ICON_BG
            )
            # Position to the left of help button (help is at x=0, so cloud is at x=-32)
            self._cloud_auth_btn.place(relx=1.0, y=-35, anchor=tk.NE, x=-32)
            # Update state based on actual auth status after cloud_browser is initialized
            self.frame.after(500, self._update_cloud_auth_button_state)

        canvas_bg = content_bg

        # ROW 1: REFRESH | Model: | dropdown | CONNECT | DISCONNECT | Status
        row1_frame = tk.Frame(inner_frame, bg=content_bg)
        row1_frame.pack(fill=tk.X)
        self._conn_row_frame = row1_frame  # Store for theme updates

        # Left side controls container
        left_frame = tk.Frame(row1_frame, bg=content_bg)
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._conn_left_frame = left_frame  # Store for theme updates

        # Refresh button - LEFT of Model label
        refresh_icon = self._button_icons.get('magnifying-glass')
        self.refresh_btn = RoundedButton(
            left_frame, text="REFRESH",
            command=self._on_refresh_models,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            width=95, height=32, radius=6, font=('Segoe UI', 9),
            icon=refresh_icon, canvas_bg=canvas_bg
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=(0, 8))

        # Model label
        self._model_label = tk.Label(
            left_frame, text="Model:",
            fg=colors['text_primary'], bg=content_bg,
            font=('Segoe UI', 9)
        )
        self._model_label.pack(side=tk.LEFT, padx=(0, 8))

        # Model dropdown
        self.model_combo = ttk.Combobox(
            left_frame,
            textvariable=self.selected_model,
            state="readonly",
            font=('Segoe UI', 9),
            width=40
        )
        self.model_combo.pack(side=tk.LEFT, padx=(0, 12))
        # Bind selection change to enable connect when switching models
        self.model_combo.bind("<<ComboboxSelected>>", self._on_model_selection_changed)

        # Style the dropdown listbox (popup)
        self._apply_combobox_listbox_style()

        # Disabled button colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Connect button - to right of dropdown (with connect icon)
        connect_icon = self._button_icons.get('connect')
        self.connect_btn = RoundedButton(
            left_frame, text="CONNECT",
            command=self._on_connect,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            disabled_bg=disabled_bg, disabled_fg=disabled_fg,
            height=32, radius=6, font=('Segoe UI', 9),
            icon=connect_icon,
            canvas_bg=canvas_bg
        )
        self.connect_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.connect_btn.set_enabled(False)  # Disabled until model selected
        self._primary_buttons.append(self.connect_btn)

        # Disconnect button - to right of connect (with disconnect icon)
        # Uses muted red/rose colors for visibility when active (not as prominent as green CONNECT)
        disconnect_icon = self._button_icons.get('disconnect')
        disconnect_bg = '#503838' if is_dark else '#fef2f2'  # Muted red/rose background
        disconnect_hover = '#5a4040' if is_dark else '#fee2e2'  # Slightly lighter on hover
        disconnect_pressed = '#402828' if is_dark else '#fecaca'  # Darker on press
        disconnect_fg = colors.get('error', '#ef4444' if is_dark else '#dc2626')  # Error color text
        self.disconnect_btn = RoundedButton(
            left_frame, text="DISCONNECT",
            command=self._on_disconnect,
            bg=disconnect_bg, hover_bg=disconnect_hover,
            pressed_bg=disconnect_pressed, fg=disconnect_fg,
            disabled_bg=disabled_bg, disabled_fg=disabled_fg,
            height=32, radius=6, font=('Segoe UI', 9),
            icon=disconnect_icon,
            canvas_bg=canvas_bg
        )
        self.disconnect_btn.pack(side=tk.LEFT)
        self.disconnect_btn.set_enabled(False)

        # Right side: Status display (pushed to far right)
        status_frame = tk.Frame(row1_frame, bg=content_bg)
        status_frame.pack(side=tk.RIGHT)
        self._conn_button_frame = status_frame  # Store for theme updates (reusing name for compatibility)

        # Status indicator dot (shows connection state visually)
        self._status_dot = tk.Canvas(
            status_frame, width=10, height=10,
            bg=content_bg, highlightthickness=0
        )
        self._status_dot.pack(side=tk.LEFT, padx=(0, 4))
        self._draw_status_dot('error')  # Initial state: not connected

        self._status_text_label = tk.Label(
            status_frame,
            text="Model:",
            fg=colors['text_muted'],
            bg=content_bg,
            font=("Segoe UI", 9)
        )
        self._status_text_label.pack(side=tk.LEFT)

        self.status_label = tk.Label(
            status_frame,
            textvariable=self.connection_status,
            fg=colors['error'],
            bg=content_bg,
            font=("Segoe UI", 9, "bold")
        )
        self.status_label.pack(side=tk.LEFT, padx=(4, 0))

        # Track dropdown popup for cloud auth (button created in header area)
        self._cloud_auth_dropdown = None

        # Track current status state for theme updates
        self._current_status_state = 'error'

        # ROW 2: Auto-match checkbox (left) and Progress indicator (right)
        row2_frame = tk.Frame(inner_frame, bg=content_bg)
        row2_frame.pack(fill=tk.X, pady=(8, 0))
        self._conn_row2_frame = row2_frame  # Store for theme updates

        # Left side of row 2: Auto-match toggle
        row2_left = tk.Frame(row2_frame, bg=content_bg)
        row2_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._conn_row2_left = row2_left  # Store for theme updates

        # Auto-match toggle using checkbox icons
        self._create_auto_match_toggle_in_connection(row2_left, content_bg)

        # Right side of row 2: Progress indicator (hidden by default)
        row2_right = tk.Frame(row2_frame, bg=content_bg)
        row2_right.pack(side=tk.RIGHT)
        self._conn_row2_right = row2_right  # Store for theme updates

        self.progress_label = tk.Label(
            row2_right,
            text="",
            fg=colors['text_muted'],
            bg=content_bg,
            font=("Segoe UI", 9, "italic")
        )
        # Don't pack initially - will be shown/hidden dynamically

        # Animated dots indicator (replaces traditional progress bar)
        self._scanning_dots = AnimatedDots(
            row2_right,
            bg=content_bg,
            dot_color=colors.get('info', '#3b82f6')
        )
        # Don't pack initially - will be shown/hidden dynamically

        # Scan status label (hidden - kept for compatibility but not displayed)
        self.scan_status = tk.StringVar(value="")

    def _create_auto_match_toggle_in_connection(self, parent: tk.Frame, bg_color: str):
        """Create auto-match checkbox toggle in the Model Connection section."""
        colors = self._theme_manager.colors

        # Container frame for toggle
        toggle_frame = tk.Frame(parent, bg=bg_color)
        toggle_frame.pack(anchor=tk.W)

        # Checkbox icon (clickable)
        is_checked = self.auto_match_enabled.get()
        icon = self._checkbox_on_icon if is_checked else self._checkbox_off_icon

        self._auto_match_icon_label = tk.Label(toggle_frame, bg=bg_color, cursor='hand2')
        if icon:
            self._auto_match_icon_label.configure(image=icon)
            self._auto_match_icon_label._icon_ref = icon
        else:
            # Fallback if icons not loaded
            self._auto_match_icon_label.configure(
                text="[x]" if is_checked else "[ ]",
                font=('Segoe UI', 9)
            )
        self._auto_match_icon_label.pack(side=tk.LEFT, padx=(0, 6))

        # Text label
        self._auto_match_text_label = tk.Label(
            toggle_frame,
            text="Auto-match local models on connect",
            bg=bg_color,
            fg=colors['text_primary'],
            font=('Segoe UI', 9),
            cursor='hand2'
        )
        self._auto_match_text_label.pack(side=tk.LEFT)

        # Store frame reference for theme updates
        self._auto_match_toggle_frame = toggle_frame

        # Bind clicks to toggle
        def on_toggle_click(event=None):
            new_value = not self.auto_match_enabled.get()
            self.auto_match_enabled.set(new_value)
            self._update_auto_match_toggle()
            self._on_auto_match_changed()

        self._auto_match_icon_label.bind('<Button-1>', on_toggle_click)
        self._auto_match_text_label.bind('<Button-1>', on_toggle_click)

    def _draw_status_dot(self, state: str):
        """
        Draw the status indicator dot.

        Args:
            state: 'success' (green), 'error' (red), or 'info' (blue)
        """
        colors = self._theme_manager.colors
        color_map = {
            'success': colors.get('success', '#22c55e'),
            'error': colors.get('error', '#ef4444'),
            'info': colors.get('info', '#3b82f6')
        }
        color = color_map.get(state, color_map['error'])
        self._status_dot.delete("all")
        self._status_dot.create_oval(1, 1, 9, 9, fill=color, outline=color)
        self._current_status_state = state

    def _update_status_display(self, state: str, text: str):
        """
        Update both the status dot and text with consistent styling.

        Args:
            state: 'success', 'error', or 'info'
            text: The status text to display
        """
        colors = self._theme_manager.colors
        color_map = {
            'success': colors.get('success', '#22c55e'),
            'error': colors.get('error', '#ef4444'),
            'info': colors.get('info', '#3b82f6')
        }
        color = color_map.get(state, color_map['error'])

        # Update dot
        self._draw_status_dot(state)

        # Update text and color
        self.connection_status.set(text)
        self.status_label.configure(fg=color)

    def _create_swap_configuration_section(self, parent):
        """Create the swap configuration master section with Quick Swap and Connections subsections"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Master section with icon labelwidget (cogwheel icon for configuration)
        header_widget = self.create_section_header(self.frame, "Swap Configuration", "cogwheel")[0]
        config_frame = ttk.LabelFrame(
            parent,
            labelwidget=header_widget,
            style='Section.TLabelframe',
            padding=12
        )
        config_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Inner content frame with distinct content background
        # Use white for light mode, very dark navy for dark mode
        content_bg = '#0d0d1a' if is_dark else '#ffffff'
        content_frame = tk.Frame(config_frame, bg=content_bg, padx=15, pady=15)
        content_frame.pack(fill=tk.BOTH, expand=True)
        self._swap_config_content_frame = content_frame  # Store for theme updates

        # Use grid layout for 1/3 + 2/3 split (like Analysis Summary / Progress Log)
        content_frame.columnconfigure(0, weight=1, minsize=280, uniform="swap_cols")
        content_frame.columnconfigure(1, weight=2, minsize=400, uniform="swap_cols")
        content_frame.rowconfigure(0, weight=1, minsize=340)  # Taller minimum height for tables
        content_frame.rowconfigure(1, weight=0)  # Fixed height for button row

        # LEFT: Quick Swap & Presets subsection
        left_container = tk.Frame(content_frame, bg=content_bg)
        left_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        self._swap_left_container = left_container  # Store for theme updates
        self._create_quick_swap_subsection(left_container, content_bg)

        # RIGHT: Connections subsection
        right_container = tk.Frame(content_frame, bg=content_bg)
        right_container.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._swap_right_container = right_container  # Store for theme updates
        self._create_connections_subsection(right_container, content_bg)

        # BOTTOM: Unified button row spanning both columns
        self._create_unified_button_row(content_frame, content_bg)

    def _create_quick_swap_subsection(self, parent, section_bg):
        """Create the Quick Swap & Presets subsection (left side of Swap Configuration)"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Subsection header with execute icon
        header_frame = tk.Frame(parent, bg=section_bg)
        header_frame.pack(fill=tk.X, pady=(0, 12))
        self._quickswap_header_frame = header_frame  # Store for theme updates

        # Use execute icon for Quick Swap header
        execute_icon = self._button_icons.get('execute')
        icon_label = None
        if execute_icon:
            icon_label = tk.Label(header_frame, image=execute_icon, bg=section_bg)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            icon_label._icon_ref = execute_icon
        self._quickswap_header_icon = icon_label  # Store for theme updates

        header_text = tk.Label(
            header_frame,
            text="Quick Swap & Presets",
            font=('Segoe UI Semibold', 11),
            fg=colors['title_color'],
            bg=section_bg
        )
        header_text.pack(side=tk.LEFT)
        self._quickswap_header_text = header_text  # Store for theme updates

        # Import/Export icons will be added in the toggle row below, right-aligned to table

        # Content area with dark background (no border - matches parent section content)
        # Use the same dark background as the parent section
        # No extra padding - matches Analysis & Progress layout (padding comes from parent content_frame)
        content_bg = '#0d0d1a' if is_dark else '#ffffff'
        content_frame = tk.Frame(
            parent, bg=content_bg
        )
        content_frame.pack(fill=tk.BOTH, expand=True)
        self._quickswap_inner_frame = content_frame  # Store for theme updates

        canvas_bg = content_bg

        # Row 1: Global/Model toggle (left) + Import/Export icons (right)
        presets_header_row = tk.Frame(content_frame, bg=content_bg, highlightthickness=0)
        presets_header_row.pack(fill=tk.X, pady=(0, 8))
        self._presets_header_row = presets_header_row  # Store for theme updates

        # Preset scope toggle (left side)
        self._create_preset_scope_toggle(presets_header_row, content_bg)

        # Import/Export icon buttons (right side, aligned with table)
        header_icon_btn_frame = tk.Frame(presets_header_row, bg=content_bg)
        header_icon_btn_frame.pack(side=tk.RIGHT)
        self._header_icon_btn_frame = header_icon_btn_frame  # Store for theme updates

        # Import global presets button (folder icon)
        folder_icon = self._button_icons.get('folder')
        self._import_icon_btn = SquareIconButton(
            header_icon_btn_frame, icon=folder_icon, command=self._on_import_preset,
            tooltip_text="Import Global Presets", size=26, radius=6
        )
        self._import_icon_btn.pack(side=tk.LEFT, padx=(0, 4))

        # Export global presets button (save icon)
        save_icon = self._button_icons.get('save')
        self._export_icon_btn = SquareIconButton(
            header_icon_btn_frame, icon=save_icon, command=self._on_export_selected_preset,
            tooltip_text="Export Global Presets", size=26, radius=6
        )
        self._export_icon_btn.pack(side=tk.LEFT)

        # Presets table container with border - use Progress Log colors
        preset_tree_bg = '#161627' if is_dark else '#f5f5f7'
        preset_table_container = tk.Frame(
            content_frame,
            bg=preset_tree_bg,
            highlightbackground=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
            highlightcolor=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
            highlightthickness=1
        )
        preset_table_container.pack(fill=tk.BOTH, expand=True)

        # Configure preset table style - match Progress Log colors
        style = ttk.Style()
        preset_tree_style = "Presets.Treeview"

        preset_tree_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        preset_heading_bg = colors.get('section_bg', '#1a1a2a' if is_dark else '#f5f5fa')
        preset_heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        preset_header_sep = '#0d0d1a' if is_dark else '#ffffff'

        style.configure(preset_tree_style,
                        background=preset_tree_bg,
                        foreground=preset_tree_fg,
                        fieldbackground=preset_tree_bg,
                        font=('Segoe UI', 9),
                        relief='flat',
                        borderwidth=0,
                        rowheight=28)
        style.layout(preset_tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])

        # Modern heading style - match Connections table exactly
        style.configure(f"{preset_tree_style}.Heading",
                        background=preset_heading_bg,
                        foreground=preset_heading_fg,
                        relief='groove',
                        borderwidth=1,
                        bordercolor=preset_header_sep,
                        lightcolor=preset_header_sep,
                        darkcolor=preset_header_sep,
                        font=('Segoe UI', 9, 'bold'),
                        padding=(8, 8))
        style.map(f"{preset_tree_style}.Heading",
                  background=[('active', preset_heading_bg), ('pressed', preset_heading_bg), ('', preset_heading_bg)],
                  relief=[('active', 'groove'), ('pressed', 'groove')])
        # Custom heading layout to center image element vertically
        style.layout(f"{preset_tree_style}.Heading", [
            ('Treeheading.cell', {'sticky': 'nswe'}),
            ('Treeheading.border', {'sticky': 'nswe', 'children': [
                ('Treeheading.padding', {'sticky': 'nswe', 'children': [
                    ('Treeheading.image', {'sticky': ''}),
                    ('Treeheading.text', {'sticky': 'we'})
                ]})
            ]})
        ])
        style.map(preset_tree_style,
                  background=[('selected', '#1a3a5c' if is_dark else '#e6f3ff')],
                  foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])

        # Presets treeview table
        preset_columns = ("name", "type", "connections")
        self._preset_tree = ttk.Treeview(
            preset_table_container,
            columns=preset_columns,
            show="headings",
            selectmode="browse",
            height=3,
            style=preset_tree_style
        )

        self._preset_tree.heading("name", text="Preset Name")
        self._preset_tree.heading("type", text="Type")
        # Use smaller link alt icon for Connections column header (14px for better centering)
        link_header_icon = self._load_icon_for_button('link alt', size=14)
        if link_header_icon:
            self._preset_tree.heading("connections", image=link_header_icon, anchor="center")
            self._preset_tree._link_header_icon = link_header_icon  # Keep reference to prevent GC
        else:
            self._preset_tree.heading("connections", text="Connections")

        self._preset_tree.column("name", width=160, minwidth=80)
        self._preset_tree.column("type", width=55, minwidth=40, anchor="center")
        self._preset_tree.column("connections", width=20, minwidth=18, anchor="center")

        # Bind double-click for Quick Apply (swap now)
        self._preset_tree.bind("<Double-1>", self._on_preset_tree_double_click)

        # Bind selection change to enable Apply button
        self._preset_tree.bind("<<TreeviewSelect>>", self._on_preset_tree_select)

        # Bind right-click for context menu
        self._preset_tree.bind("<Button-3>", self._on_preset_tree_right_click)

        self._preset_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add tooltip for connections column header
        self._preset_header_tooltip = None
        self._preset_header_tooltip_id = None

        def on_preset_tree_motion(event):
            """Show tooltip when hovering over the connections header."""
            region = self._preset_tree.identify_region(event.x, event.y)
            if region == "heading":
                column = self._preset_tree.identify_column(event.x)
                if column == "#3":  # connections column
                    if not self._preset_header_tooltip:
                        # Create tooltip window
                        self._preset_header_tooltip = tk.Toplevel(self._preset_tree)
                        self._preset_header_tooltip.wm_overrideredirect(True)
                        self._preset_header_tooltip.wm_attributes("-topmost", True)

                        tip_bg = '#2a2a3c' if is_dark else '#f5f5dc'
                        tip_fg = colors['text_primary']
                        tip_label = tk.Label(
                            self._preset_header_tooltip,
                            text="Number of connections in preset",
                            bg=tip_bg, fg=tip_fg,
                            font=("Segoe UI", 9),
                            padx=6, pady=3,
                            relief=tk.SOLID, borderwidth=1
                        )
                        tip_label.pack()

                        # Position tooltip
                        x = self._preset_tree.winfo_rootx() + event.x + 10
                        y = self._preset_tree.winfo_rooty() + event.y + 10
                        self._preset_header_tooltip.wm_geometry(f"+{x}+{y}")
                else:
                    # Not over connections header, hide tooltip
                    if self._preset_header_tooltip:
                        self._preset_header_tooltip.destroy()
                        self._preset_header_tooltip = None
            else:
                # Not in header region, hide tooltip
                if self._preset_header_tooltip:
                    self._preset_header_tooltip.destroy()
                    self._preset_header_tooltip = None

        def on_preset_tree_leave(event):
            """Hide tooltip when leaving the treeview."""
            if self._preset_header_tooltip:
                self._preset_header_tooltip.destroy()
                self._preset_header_tooltip = None

        self._preset_tree.bind("<Motion>", on_preset_tree_motion)
        self._preset_tree.bind("<Leave>", on_preset_tree_leave)

        # Scrollbar area frame - match Progress Log width (ThemedScrollbar default is 12)
        scrollbar_area_bg = '#1a1a2e' if is_dark else '#f0f0f0'
        preset_scrollbar_area = tk.Frame(
            preset_table_container,
            bg=scrollbar_area_bg,
            width=12
        )
        preset_scrollbar_area.pack(side=tk.RIGHT, fill=tk.Y)
        preset_scrollbar_area.pack_propagate(False)  # Keep fixed width
        self._preset_scrollbar_area = preset_scrollbar_area

        # Small scrollbar for presets - inside scrollbar area
        preset_scrollbar = ThemedScrollbar(
            preset_scrollbar_area,
            command=self._preset_tree.yview,
            theme_manager=self._theme_manager,
            auto_hide=True
        )
        self._preset_tree.configure(yscrollcommand=preset_scrollbar.set)
        preset_scrollbar.pack(fill=tk.Y, expand=True)
        self._preset_scrollbar = preset_scrollbar

        # Store reference for theme updates
        self._preset_table_container = preset_table_container

        # Keep old reference for compatibility
        self._preset_buttons_frame = None
        self._preset_buttons = []

        # Load initial presets into table
        self._refresh_preset_table()

    def _create_connections_subsection(self, parent, section_bg):
        """Create the Connections subsection (right side of Swap Configuration)"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Subsection header with link icon + "Connections" text
        header_frame = tk.Frame(parent, bg=section_bg)
        header_frame.pack(fill=tk.X, pady=(0, 12))
        self._connections_header_frame = header_frame  # Store for theme updates

        # Use link icon for Connections header with text
        link_icon = self._button_icons.get('link')
        icon_label = None
        if link_icon:
            icon_label = tk.Label(header_frame, image=link_icon, bg=section_bg)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            icon_label._icon_ref = link_icon
        self._connections_header_icon = icon_label  # Store for theme updates

        # Connections text label
        header_text = tk.Label(
            header_frame,
            text="Connections",
            font=('Segoe UI Semibold', 11),
            fg=colors['title_color'],
            bg=section_bg
        )
        header_text.pack(side=tk.LEFT)
        self._connections_header_text = header_text  # Store for theme updates

        # Info label
        self.mapping_info = tk.StringVar(value="Connect to a model to see swappable connections")
        self._mapping_info_label = tk.Label(
            header_frame,
            textvariable=self.mapping_info,
            font=("Segoe UI", 9, "italic"),
            fg=colors['text_muted'],
            bg=section_bg
        )
        self._mapping_info_label.pack(side=tk.RIGHT)

        # Select All link button (right side, before info label)
        # Use title_color for consistency with section headers
        if is_dark:
            link_color = colors.get('title_color', '#0084b7')  # Light blue matching title
            hover_color = '#006691'  # Darker blue for hover
        else:
            link_color = colors.get('primary', '#009999')  # Teal for light mode
            hover_color = colors.get('primary_hover', '#007A7A')  # Darker teal for hover
        self._select_all_link = tk.Label(
            header_frame,
            text="Select All",
            font=("Segoe UI", 9, "underline"),
            fg=link_color,
            bg=section_bg,
            cursor='hand2'
        )
        self._select_all_link.pack(side=tk.RIGHT, padx=(0, 15))
        self._select_all_link.bind('<Button-1>', lambda e: self._on_select_all_connections())
        self._select_all_link.bind('<Enter>', lambda e: self._select_all_link.configure(fg=hover_color))
        self._select_all_link.bind('<Leave>', lambda e: self._select_all_link.configure(fg=link_color))

        # Content area with same background as Progress Log content
        content_bg = '#161627' if is_dark else '#f5f5f7'
        tree_border = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')

        # No spacer - connections table top aligns with GLOBAL/MODEL buttons row
        self._connections_spacer = None

        tree_container = tk.Frame(
            parent,
            bg=content_bg,
            highlightbackground=tree_border,
            highlightcolor=tree_border,
            highlightthickness=1
        )
        tree_container.pack(fill=tk.BOTH, expand=True)
        self._tree_container = tree_container
        self._mapping_inner_frame = parent  # Store for theme updates

        # Configure treeview style - flat, modern design
        style = ttk.Style()
        tree_style = "HotSwap.Treeview"
        style.configure(tree_style,
                        background=content_bg,
                        foreground=colors.get('text_primary', '#e0e0e0' if is_dark else '#333333'),
                        fieldbackground=content_bg,
                        font=('Segoe UI', 9),
                        relief='flat',
                        borderwidth=0,
                        bordercolor=content_bg,
                        lightcolor=content_bg,
                        darkcolor=content_bg,
                        rowheight=28)
        style.layout(tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])

        # Modern heading style with bold font, padding, and faint column separator
        heading_bg = colors.get('section_bg', '#1a1a2a' if is_dark else '#f5f5fa')
        heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        header_separator = '#0d0d1a' if is_dark else '#ffffff'

        style.configure(f"{tree_style}.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        font=('Segoe UI', 9, 'bold'),
                        padding=(8, 6))
        style.map(f"{tree_style}.Heading",
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)],
                  relief=[('active', 'groove'), ('pressed', 'groove')])
        style.map(tree_style,
                  background=[('selected', '#1a3a5c' if is_dark else '#e6f3ff')],
                  foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])

        # Treeview for mappings with inline target selection
        columns = ("source", "target", "type", "status")
        self.mapping_tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            selectmode="extended",
            height=4,
            style=tree_style
        )

        # Configure columns - center aligned for better readability
        self.mapping_tree.heading("source", text="Source Connection")
        self.mapping_tree.heading("target", text="Target (double-click to set)")
        self.mapping_tree.heading("type", text="Type")
        self.mapping_tree.heading("status", text="Status")

        self.mapping_tree.column("source", width=220, minwidth=120, anchor="center")
        self.mapping_tree.column("target", width=240, minwidth=120, anchor="center")
        self.mapping_tree.column("type", width=60, minwidth=50, anchor="center")
        self.mapping_tree.column("status", width=70, minwidth=50, anchor="center")

        self.mapping_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar area frame - match Progress Log width (ThemedScrollbar default is 12)
        scrollbar_area_bg = '#1a1a2e' if is_dark else '#f0f0f0'
        mapping_scrollbar_area = tk.Frame(
            tree_container,
            bg=scrollbar_area_bg,
            width=12
        )
        mapping_scrollbar_area.pack(side=tk.RIGHT, fill=tk.Y)
        mapping_scrollbar_area.pack_propagate(False)  # Keep fixed width
        self._mapping_scrollbar_area = mapping_scrollbar_area

        # Scrollbar - inside scrollbar area
        scrollbar = ThemedScrollbar(mapping_scrollbar_area, command=self.mapping_tree.yview, theme_manager=self._theme_manager, auto_hide=True)
        self.mapping_tree.configure(yscrollcommand=scrollbar.set)
        self._mapping_scrollbar = scrollbar
        scrollbar.pack(fill=tk.Y, expand=True)

        # Bind selection change and click for inline target picker
        self.mapping_tree.bind("<<TreeviewSelect>>", self._on_mapping_selected)
        self.mapping_tree.bind("<Button-1>", self._on_tree_click)
        self.mapping_tree.bind("<Double-1>", self._on_mapping_double_click)
        self.mapping_tree.bind("<Button-3>", self._on_right_click)

        # Context menu popup (modern styled - initialized on demand)
        self._context_popup = None

        # Initialize inline target picker
        self._inline_target_picker = InlineTargetPicker(
            parent=tree_container,
            get_local_models=self._get_local_models_for_picker,
            on_cloud_browse=self._on_inline_cloud_browse,
            on_target_selected=self._on_inline_target_selected
        )

        # Diagram container (hidden by default)
        self._diagram_container = tk.Frame(
            parent,
            bg=content_bg,
            highlightbackground=tree_border,
            highlightcolor=tree_border,
            highlightthickness=1
        )
        # Don't pack - will be shown when diagram view is selected

        # Connection diagram widget
        self._connection_diagram = ConnectionDiagram(
            self._diagram_container,
            on_node_click=self._on_diagram_node_click
        )
        self._connection_diagram.pack(fill=tk.BOTH, expand=True)

        # Cache for local models in UI
        self._local_models_cache = []
        self._dropdown_models_cache: List[AvailableModel] = []  # Full model objects from dropdown
        self._selected_local_target = None
        self._last_swapped_mapping = None
        self._swap_history: List[SwapHistoryEntry] = []
        self._max_history_entries = 50  # Store up to 50 individual swaps (covers ~10 batch runs)

        # Thin report context - stores info when connected to a thin report
        # This allows the preset system to work with thin reports
        self._thin_report_context: Optional[dict] = None  # {file_path, process_id, cloud_server, cloud_database, local_server}

        # Current model file path - for consistent preset hashing across all model types
        # This ensures presets persist for a file regardless of its current connection state
        self._current_model_file_path: Optional[str] = None

        # Swap history persistence
        self._swap_history_file = self._get_swap_history_file_path()
        self._load_swap_history()

    def _create_unified_button_row(self, parent, section_bg):
        """Create the unified button row spanning both columns at the bottom of Swap Configuration."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Button row using grid to span both columns
        # Use 12px top padding to match bottom padding of section (content_frame has pady=15)
        btn_row = tk.Frame(parent, bg=section_bg, highlightthickness=0)
        btn_row.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(12, 0))
        self._unified_btn_row = btn_row

        # Use same column weights as parent (1:2 ratio) for proper centering under tables
        btn_row.columnconfigure(0, weight=1, uniform="btn_cols")  # Left side (presets table width)
        btn_row.columnconfigure(1, weight=2, uniform="btn_cols")  # Right side (connections table width)

        # Left side container for APPLY PRESET and LAST CONFIG (centered under presets table)
        left_btn_container = tk.Frame(btn_row, bg=section_bg, highlightthickness=0)
        left_btn_container.grid(row=0, column=0, padx=(0, 15))  # No sticky - centers naturally

        # Right side container for SAVE MAPPING (centered under connections table)
        right_btn_container = tk.Frame(btn_row, bg=section_bg, highlightthickness=0)
        right_btn_container.grid(row=0, column=1)  # No sticky - centers naturally

        # Use theme-aware disabled colors
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # APPLY PRESET button
        eye_icon = self._button_icons.get('eye')
        self.apply_preset_btn = RoundedButton(
            left_btn_container, text="APPLY PRESET",
            command=self._on_apply_preset_to_mappings,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            icon=eye_icon, canvas_bg=section_bg,
            disabled_bg=disabled_bg, disabled_fg=disabled_fg
        )
        self.apply_preset_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.apply_preset_btn.set_enabled(False)
        Tooltip(self.apply_preset_btn, "Apply the selected preset's target connections to the mapping table")

        # LAST CONFIG button (restores saved starting configuration)
        reset_icon = self._button_icons.get('reset')
        self.last_config_btn = RoundedButton(
            left_btn_container, text="LAST CONFIG",
            command=self._on_apply_last_config,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            icon=reset_icon, canvas_bg=section_bg,
            disabled_bg=disabled_bg, disabled_fg=disabled_fg
        )
        self.last_config_btn.pack(side=tk.LEFT)
        self.last_config_btn.set_enabled(False)
        Tooltip(self.last_config_btn, self._get_last_config_tooltip)

        # SAVE MAPPING button
        save_icon = self._button_icons.get('save')
        self.save_mapping_btn = RoundedButton(
            right_btn_container, text="SAVE MAPPING",
            command=self._on_save_preset,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            icon=save_icon, canvas_bg=section_bg,
            disabled_bg=disabled_bg, disabled_fg=disabled_fg
        )
        self.save_mapping_btn.pack()  # Centers in container, container centers in grid cell
        self.save_mapping_btn.set_enabled(False)
        Tooltip(self.save_mapping_btn, "Save current mapping configuration as a preset")

        # Store references for theme updates
        self._left_btn_container = left_btn_container
        self._right_btn_container = right_btn_container

    def _on_tree_click(self, event):
        """Handle click on treeview - show inline picker for Target column"""
        # Identify which row and column was clicked
        region = self.mapping_tree.identify_region(event.x, event.y)
        column = self.mapping_tree.identify_column(event.x)
        item_id = self.mapping_tree.identify_row(event.y)

        # If clicking in empty area (not on a cell/row), clear selection and update buttons
        if region != "cell" or not item_id:
            # Use after_idle to let tkinter process the click first
            self.frame.after_idle(self._on_mapping_selected)
            return

        # Column #2 is the Target column (0-indexed internally as #1, #2, #3)
        if column == "#2":
            # Get cell bounding box for positioning
            bbox = self.mapping_tree.bbox(item_id, column)
            if bbox:
                # Convert to screen coordinates
                x = self.mapping_tree.winfo_rootx() + bbox[0]
                y = self.mapping_tree.winfo_rooty() + bbox[1] + bbox[3]  # Below the cell
                width = bbox[2]

                # Show the inline picker
                self._inline_target_picker.show_picker(item_id, x, y, max(width, 220))

    def _on_right_click(self, event):
        """Handle right-click on treeview - show modern styled context menu"""
        # Select the row under cursor
        item_id = self.mapping_tree.identify_row(event.y)
        if item_id:
            # Select this row if not already selected
            self.mapping_tree.selection_set(item_id)
            self._context_menu_item_id = item_id

            # Show modern context menu popup
            self._show_context_popup(event.x_root, event.y_root)

    def _show_context_popup(self, x: int, y: int):
        """Show a modern styled context menu popup at the given position."""
        # Close any existing popup
        self._close_context_popup()

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create popup window - attached to parent, not topmost
        self._context_popup = tk.Toplevel(self.mapping_tree)
        self._context_popup.withdraw()  # Hide until positioned
        self._context_popup.overrideredirect(True)  # No window decorations
        # Removed -topmost so popup doesn't float above other apps

        # Configure popup appearance
        popup_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
        text_color = colors['text_primary']
        hover_bg = colors.get('hover', '#2a2a3e' if is_dark else '#f0f0f5')
        primary_color = colors.get('primary', '#4a6cf5')

        # Border frame
        border_frame = tk.Frame(self._context_popup, bg=border_color, padx=1, pady=1)
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Main content frame
        main_frame = tk.Frame(border_frame, bg=popup_bg)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Menu items
        menu_items = [
            ("Edit Target...", self._on_context_edit_target),
            ("Clear Target", self._on_context_clear_target),
            None,  # Separator
            ("Swap Selected", self._on_swap_selected),
        ]

        for item in menu_items:
            if item is None:
                # Separator
                sep = tk.Frame(main_frame, height=1, bg=border_color)
                sep.pack(fill=tk.X, padx=8, pady=4)
            else:
                label_text, command = item
                item_frame = tk.Frame(main_frame, bg=popup_bg)
                item_frame.pack(fill=tk.X)

                label = tk.Label(
                    item_frame,
                    text=label_text,
                    font=('Segoe UI', 9),
                    fg=text_color,
                    bg=popup_bg,
                    anchor='w',
                    padx=12,
                    pady=6,
                    cursor='hand2'
                )
                label.pack(fill=tk.X)

                # Hover effects
                def on_enter(e, f=item_frame, l=label, hbg=hover_bg):
                    f.configure(bg=hbg)
                    l.configure(bg=hbg)

                def on_leave(e, f=item_frame, l=label, bg=popup_bg):
                    f.configure(bg=bg)
                    l.configure(bg=bg)

                def on_click(e, cmd=command):
                    self._close_context_popup()
                    cmd()

                for widget in [item_frame, label]:
                    widget.bind('<Enter>', on_enter)
                    widget.bind('<Leave>', on_leave)
                    widget.bind('<Button-1>', on_click)

        # Add Recents section - show available targets with type indicators
        recent_targets = self._get_recent_targets_for_context()
        if recent_targets:
            # Separator before Recents
            sep = tk.Frame(main_frame, height=1, bg=border_color)
            sep.pack(fill=tk.X, padx=8, pady=4)

            # Recents header (smaller, muted)
            recents_header = tk.Label(
                main_frame,
                text="Quick Set Target",
                font=('Segoe UI', 8),
                fg=colors.get('text_muted', '#888888'),
                bg=popup_bg,
                anchor='w',
                padx=12
            )
            recents_header.pack(fill=tk.X, pady=(4, 2))

            # Get pre-loaded SVG icons for cloud (C) and local (L)
            cloud_icon_img = self._button_icons.get('letter-c')
            local_icon_img = self._button_icons.get('letter-l')

            # Recent target items (max 5)
            for target in recent_targets[:5]:
                item_frame = tk.Frame(main_frame, bg=popup_bg)
                item_frame.pack(fill=tk.X)

                # Type indicator icon (cloud or local) - use pre-loaded SVG icons
                is_cloud = target.target_type == "cloud"
                icon_img = cloud_icon_img if is_cloud else local_icon_img

                if icon_img:
                    type_label = tk.Label(
                        item_frame,
                        image=icon_img,
                        bg=popup_bg
                    )
                    type_label.image = icon_img  # Keep reference
                else:
                    # Fallback to Unicode
                    type_icon = "\u24b8" if is_cloud else "\u24c1"
                    type_color = colors.get('primary', '#4a9eff') if is_cloud else colors.get('text_secondary', '#888888')
                    type_label = tk.Label(
                        item_frame,
                        text=type_icon,
                        font=('Segoe UI', 9),
                        fg=type_color,
                        bg=popup_bg,
                        width=2
                    )
                type_label.pack(side=tk.LEFT, padx=(12, 4))

                label = tk.Label(
                    item_frame,
                    text=target.display_name,
                    font=('Segoe UI', 9),
                    fg=text_color,
                    bg=popup_bg,
                    anchor='w',
                    pady=4,
                    cursor='hand2'
                )
                label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))

                # Hover effects
                def on_enter(e, f=item_frame, l=label, tl=type_label, hbg=hover_bg):
                    f.configure(bg=hbg)
                    l.configure(bg=hbg)
                    tl.configure(bg=hbg)

                def on_leave(e, f=item_frame, l=label, tl=type_label, bg=popup_bg):
                    f.configure(bg=bg)
                    l.configure(bg=bg)
                    tl.configure(bg=bg)

                def on_click(e, t=target):
                    self._close_context_popup()
                    item_id = getattr(self, '_context_menu_item_id', None)
                    if item_id:
                        self._on_inline_target_selected(item_id, t)

                for widget in [item_frame, type_label, label]:
                    widget.bind('<Enter>', on_enter)
                    widget.bind('<Leave>', on_leave)
                    widget.bind('<Button-1>', on_click)

        # Position popup
        self._context_popup.update_idletasks()
        self._context_popup.geometry(f"+{x}+{y}")
        self._context_popup.deiconify()
        # Lift relative to parent window only (not above other apps)
        self._context_popup.lift(self.frame.winfo_toplevel())
        self._context_popup.focus_set()

        # Bind events to close popup
        # Note: FocusOut removed - it fires unexpectedly and conflicts with click handling
        self._context_popup.bind('<Escape>', lambda e: self._close_context_popup())
        # Track bind ID for proper cleanup
        self._outside_click_bind_id = self.frame.winfo_toplevel().bind(
            '<Button-1>', self._on_context_outside_click, add='+'
        )

    def _on_context_outside_click(self, event):
        """Handle click outside context popup."""
        if not self._context_popup or not self._context_popup.winfo_exists():
            return

        # Check if click is outside popup
        px = self._context_popup.winfo_rootx()
        py = self._context_popup.winfo_rooty()
        pw = self._context_popup.winfo_width()
        ph = self._context_popup.winfo_height()

        if not (px <= event.x_root <= px + pw and py <= event.y_root <= py + ph):
            self._close_context_popup()

    def _close_context_popup(self):
        """Close the context menu popup."""
        if self._context_popup and self._context_popup.winfo_exists():
            self._context_popup.destroy()
        self._context_popup = None
        # Properly unbind using tracked bind ID
        if hasattr(self, '_outside_click_bind_id') and self._outside_click_bind_id:
            try:
                self.frame.winfo_toplevel().unbind('<Button-1>', self._outside_click_bind_id)
            except Exception:
                pass
            self._outside_click_bind_id = None

    def _on_context_edit_target(self):
        """Handle Edit Target from context menu - shows Select Target Type dialog"""
        item_id = getattr(self, '_context_menu_item_id', None)
        if not item_id or not self.mapping_tree.exists(item_id):
            return

        # Get the mapping index from the item_id
        try:
            mapping_index = int(item_id)
        except (ValueError, TypeError):
            return

        if mapping_index < 0 or mapping_index >= len(self.mappings):
            return

        mapping = self.mappings[mapping_index]

        # Show the Select Target Type dialog and proceed with selection
        self._show_target_selector(mapping, mapping_index)

    def _on_context_clear_target(self):
        """Handle Clear Target from context menu"""
        item_id = getattr(self, '_context_menu_item_id', None)
        if item_id:
            self._on_inline_target_selected(item_id, None)

    def _get_local_models_for_picker(self) -> List[SwapTarget]:
        """Get list of local models for the inline picker"""
        return self._local_models_cache

    def _get_recent_targets_for_context(self) -> List['SwapTarget']:
        """Get list of recent/available targets for context menu quick set."""
        targets = []
        seen_names = set()

        # First, add targets currently used in mappings (these are "recent")
        if hasattr(self, 'mappings') and self.mappings:
            for mapping in self.mappings:
                if mapping.target and mapping.target.display_name not in seen_names:
                    targets.append(mapping.target)
                    seen_names.add(mapping.target.display_name)

        # Then add local models from cache
        if hasattr(self, '_local_models_cache') and self._local_models_cache:
            for model in self._local_models_cache:
                if model.display_name not in seen_names:
                    targets.append(model)
                    seen_names.add(model.display_name)

        return targets

    def _on_inline_cloud_browse(self, item_id: str):
        """Handle Browse Cloud selection from inline picker"""
        # Store the item being edited
        self._inline_edit_item_id = item_id
        # Open cloud browser dialog
        self._on_browse_cloud()

    def _on_inline_target_selected(self, item_id: str, target: Optional[SwapTarget]):
        """Handle target selection from inline picker"""
        # The item_id is the string index of the mapping (set in _populate_mapping_table)
        try:
            mapping_index = int(item_id)
        except (ValueError, TypeError):
            return

        if mapping_index < 0 or mapping_index >= len(self.mappings):
            return

        mapping = self.mappings[mapping_index]

        if target is None:
            # Clear target
            mapping.target = None
            mapping.status = SwapStatus.PENDING
            self._log_message(f"Cleared target for: {mapping.source.name}")
        else:
            # Set new target
            mapping.target = target
            mapping.status = SwapStatus.READY
            self._log_message(f"Set target for {mapping.source.name}: {target.display_name}")

        # Update the display (preserving selection)
        self._populate_mapping_table(self.mappings)

        # Restore selection to the item that was just edited
        if self.mapping_tree.exists(item_id):
            self.mapping_tree.selection_set(item_id)
            self._on_mapping_selected()  # Update button states with restored selection

    def _create_log_section(self, parent):
        """Create the split activity log section (Connection Details + Progress Log)"""
        # Use SplitLogSection template widget
        self.log_section = SplitLogSection(
            parent=parent,
            theme_manager=self._theme_manager,
            section_title="Analysis & Progress",
            section_icon="analyze",
            summary_title="Connection Details",
            summary_icon="bar-chart",
            log_title="Progress Log",
            log_icon="log-file",
            summary_placeholder="Connect to a model to see connection details"
        )
        self.log_section.pack(fill=tk.BOTH, expand=True)

        # Store references for compatibility with existing methods
        self.log_text = self.log_section.log_text
        self._summary_text = self.log_section.summary_text
        self._summary_placeholder = self.log_section.placeholder_label

        # Show welcome message
        self._show_welcome_message()

    def _show_welcome_message(self):
        """Show welcome message for Connection Hot-Swap"""
        self.log_message("🔄 Welcome to Connection Hot-Swap!")
        self.log_message("=" * 60)
        self.log_message("♻️ Swap Power BI connections: cloud-to-local, local-to-cloud, or cloud-to-cloud")
        self.log_message("🎯 Composite models: Hot-swap while open | Thin reports: Requires PBIP format")
        self.log_message("☁️ Connect to perspectives in any workspace (Pro, Premium, PPU, or Fabric)")
        self.log_message("")
        self.log_message("👆 Start by selecting a semantic model from the dropdown")
        self.log_message("🔀 Configure swap mappings and click 'SWAP' to apply changes")

    def _update_connection_details(self):
        """Update the Connection Details panel with current model information"""
        if not self.model_info or not hasattr(self, '_summary_text'):
            return

        # Clear any existing card container (from selected connection view)
        if hasattr(self, '_card_container') and self._card_container:
            self._card_container.destroy()
            self._card_container = None

        # Hide placeholder and show text widget
        if hasattr(self, '_summary_placeholder'):
            self._summary_placeholder.grid_forget()
        self._summary_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Get model information
        model_name = self.selected_model.get() or "Unknown"
        model_type = self.model_info.connection_type.value if self.model_info.connection_type else "Unknown"
        total_datasources = self.model_info.total_datasources
        swappable_count = self.model_info.swappable_count

        # Build connection details text
        details = []
        details.append(f"Model: {model_name}")
        details.append(f"Type: {model_type}")
        details.append(f"Total Datasources: {total_datasources}")
        details.append(f"Swappable Connections: {swappable_count}")
        details.append("")

        # List each connection with its details
        if self.model_info.connections:
            details.append("Connections:")
            for conn in self.model_info.connections:
                swap_status = "[Swappable]" if conn.is_swappable else "[Fixed]"
                details.append(f"  {swap_status} {conn.display_name}")
                if conn.server:
                    details.append(f"    Server: {conn.server}")
                if conn.database:
                    details.append(f"    Database: {conn.database}")

        # Update the summary text
        self._summary_text.configure(state=tk.NORMAL)
        self._summary_text.delete('1.0', tk.END)
        self._summary_text.insert(tk.END, '\n'.join(details))
        self._summary_text.configure(state=tk.DISABLED)

    def _update_selected_connection_details(self, mapping: 'ConnectionMapping'):
        """Update the Connection Details panel with source/target info in two-column card layout."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Get the summary frame from the log section
        if not hasattr(self, 'log_section') or not hasattr(self.log_section, '_summary_frame'):
            return

        summary_frame = self.log_section._summary_frame

        # Hide placeholder
        if hasattr(self, '_summary_placeholder'):
            self._summary_placeholder.grid_forget()

        # Hide text widget if it exists
        if hasattr(self, '_summary_text'):
            self._summary_text.grid_forget()

        # Clear any existing card container
        if hasattr(self, '_card_container') and self._card_container:
            self._card_container.destroy()

        source = mapping.source
        target = mapping.target

        def trunc(text, max_len=28):
            if not text:
                return "-"
            return text[:max_len-2] + ".." if len(text) > max_len else text

        # Create card container
        self._card_container = tk.Frame(summary_frame, bg=colors['section_bg'])
        self._card_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._card_container.columnconfigure(0, weight=1, uniform="cards")
        self._card_container.columnconfigure(1, weight=1, uniform="cards")
        self._card_container.rowconfigure(0, weight=0)  # Header row
        self._card_container.rowconfigure(1, weight=1)  # Cards row

        # Header row with connection name/status on left and copy button on right
        header_frame = tk.Frame(self._card_container, bg=colors['section_bg'])
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 8))
        header_frame.columnconfigure(0, weight=1)

        status_text = mapping.status.value.upper()
        header_label = tk.Label(
            header_frame,
            text=f"{source.name}  [{status_text}]",
            font=('Segoe UI Semibold', 10),
            bg=colors['section_bg'],
            fg=colors['text_primary']
        )
        header_label.pack(side=tk.LEFT)

        # Copy button in upper right - load copy icon
        copy_icon = self._load_icon_for_button('copy', size=16)

        def copy_details():
            """Copy connection details to clipboard"""
            details = []
            details.append(f"Connection: {source.name}")
            details.append(f"Status: {status_text}")
            details.append("")
            details.append("SOURCE:")
            details.append(f"  Server: {source.server}")
            details.append(f"  Database: {source.database}")
            if hasattr(source, 'dataset_id') and source.dataset_id:
                details.append(f"  Dataset ID: {source.dataset_id}")
            if source.is_cloud and hasattr(source, 'dataset_name') and source.dataset_name:
                details.append(f"  Model: {source.dataset_name}")
            if source.workspace_name:
                details.append(f"  Workspace: {source.workspace_name}")
            if hasattr(source, 'perspective_name') and source.perspective_name:
                details.append(f"  Perspective: {source.perspective_name}")
            details.append("")
            details.append("TARGET:")
            if target:
                details.append(f"  Server: {target.server}")
                details.append(f"  Database: {target.database}")
                if hasattr(target, 'dataset_id') and target.dataset_id:
                    details.append(f"  Dataset ID: {target.dataset_id}")
                if target.target_type == "cloud" and hasattr(target, 'display_name') and target.display_name:
                    details.append(f"  Model: {target.display_name}")
                if target.workspace_name:
                    details.append(f"  Workspace: {target.workspace_name}")
                if target.perspective_name:
                    details.append(f"  Perspective: {target.perspective_name}")
            else:
                details.append("  (not configured)")

            self.frame.clipboard_clear()
            self.frame.clipboard_append('\n'.join(details))
            self._log_message("Connection details copied to clipboard")

        copy_btn = SquareIconButton(
            header_frame,
            icon=copy_icon,
            command=copy_details,
            tooltip_text="Copy Details",
            size=26,
            radius=6,
            bg_normal_override={'dark': '#0d0d1a', 'light': '#ffffff'}
        )
        copy_btn.pack(side=tk.RIGHT)

        # Card styling
        card_bg = colors['background']
        card_border = colors.get('border', '#3a3a4e' if is_dark else '#d0d0d8')
        label_fg = colors['text_secondary']
        value_fg = colors['text_primary']

        def create_card(parent, title, col):
            """Create a card frame with title"""
            card = tk.Frame(parent, bg=card_bg, highlightbackground=card_border,
                           highlightthickness=1, padx=10, pady=8)
            card.grid(row=1, column=col, sticky=(tk.W, tk.E, tk.N, tk.S),
                     padx=(0, 6) if col == 0 else (6, 0))

            # Card title - uses title_color which is teal in light mode, blue in dark mode
            title_label = tk.Label(card, text=title, font=('Segoe UI Semibold', 9),
                                  bg=card_bg, fg=colors['title_color'])
            title_label.pack(anchor=tk.W, pady=(0, 6))

            return card

        def add_row(card, label_text, value_text):
            """Add a label: value row to a card"""
            row = tk.Frame(card, bg=card_bg)
            row.pack(fill=tk.X, pady=1)

            label = tk.Label(row, text=label_text, font=('Segoe UI', 9),
                           bg=card_bg, fg=label_fg, width=10, anchor=tk.W)
            label.pack(side=tk.LEFT)

            value = tk.Label(row, text=value_text, font=('Segoe UI', 9),
                           bg=card_bg, fg=value_fg, anchor=tk.W)
            value.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Create SOURCE card - show cloud type if it's a cloud connection
        if source.is_cloud:
            from tools.connection_hotswap.models import CloudConnectionType
            if hasattr(source, 'cloud_connection_type') and source.cloud_connection_type:
                if source.cloud_connection_type == CloudConnectionType.PBI_SEMANTIC_MODEL:
                    src_type = "PBI Semantic Model"
                elif source.cloud_connection_type == CloudConnectionType.AAS_XMLA:
                    src_type = "AAS XMLA"
                else:
                    src_type = "Cloud"
            else:
                src_type = "Cloud"
        else:
            src_type = "Local"
        source_card = create_card(self._card_container, f"SOURCE ({src_type})", 0)
        add_row(source_card, "Server:", trunc(source.server))
        add_row(source_card, "Database:", trunc(source.database))
        # Show Dataset ID if available (cloud connections)
        if hasattr(source, 'dataset_id') and source.dataset_id:
            add_row(source_card, "Dataset ID:", trunc(source.dataset_id))
        # Show Model Name (friendly name) for cloud connections
        if source.is_cloud and hasattr(source, 'dataset_name') and source.dataset_name:
            add_row(source_card, "Model:", trunc(source.dataset_name))
        add_row(source_card, "Workspace:", trunc(source.workspace_name) if source.workspace_name else "-")

        # Show perspective if present
        src_persp = source.perspective_name if hasattr(source, 'perspective_name') and source.perspective_name else None
        if src_persp:
            add_row(source_card, "Perspective:", trunc(src_persp))

        # Create TARGET card - show cloud type if it's a cloud connection
        if target and target.target_type == "cloud":
            from tools.connection_hotswap.models import CloudConnectionType
            if hasattr(target, 'cloud_connection_type') and target.cloud_connection_type:
                if target.cloud_connection_type == CloudConnectionType.PBI_SEMANTIC_MODEL:
                    tgt_type = "PBI Semantic Model"
                elif target.cloud_connection_type == CloudConnectionType.AAS_XMLA:
                    tgt_type = "AAS XMLA"
                else:
                    tgt_type = "Cloud"
            else:
                tgt_type = "Cloud"
        else:
            tgt_type = target.target_type.title() if target else "Not Set"
        target_card = create_card(self._card_container, f"TARGET ({tgt_type})", 1)

        if target:
            add_row(target_card, "Server:", trunc(target.server))
            add_row(target_card, "Database:", trunc(target.database))
            # Show Dataset ID if available (cloud connections)
            if hasattr(target, 'dataset_id') and target.dataset_id:
                add_row(target_card, "Dataset ID:", trunc(target.dataset_id))
            # Show Model Name (friendly name) for cloud connections
            if target.target_type == "cloud" and hasattr(target, 'display_name') and target.display_name:
                add_row(target_card, "Model:", trunc(target.display_name))
            add_row(target_card, "Workspace:", trunc(target.workspace_name) if target.workspace_name else "-")

            # Show perspective if present
            tgt_persp = target.perspective_name if target.perspective_name else None
            if tgt_persp:
                add_row(target_card, "Perspective:", trunc(tgt_persp))
        else:
            # No target configured
            hint = tk.Label(target_card, text="Double-click Target\ncell to configure",
                          font=('Segoe UI', 9, 'italic'), bg=card_bg, fg=label_fg,
                          justify=tk.CENTER)
            hint.pack(expand=True, pady=10)

    def _clear_connection_details(self):
        """Clear the Connection Details panel and show placeholder"""
        # Destroy card container if it exists
        if hasattr(self, '_card_container') and self._card_container:
            self._card_container.destroy()
            self._card_container = None

        if hasattr(self, '_summary_text'):
            self._summary_text.configure(state=tk.NORMAL)
            self._summary_text.delete('1.0', tk.END)
            self._summary_text.configure(state=tk.DISABLED)
            self._summary_text.grid_forget()

        if hasattr(self, '_summary_placeholder'):
            self._summary_placeholder.configure(text="Connect to a model to see connection details")
            self._summary_placeholder.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

    def _set_initial_state(self):
        """Set initial disabled state"""
        self.last_config_btn.set_enabled(False)
        self.swap_selected_btn.set_enabled(False)

    # =========================================================================
    # Progress Indicators
    # =========================================================================

    def _show_progress(self, message: str = "Working..."):
        """Show the animated dots indicator with a message (on row 2, right side below Status)"""
        self.progress_label.configure(text=message)
        # Pack in row2_right frame (right side of row 2)
        self.progress_label.pack(side=tk.LEFT, padx=(0, 6))
        self._scanning_dots.pack(side=tk.LEFT)
        self._scanning_dots.start()

    def _hide_progress(self):
        """Hide the animated dots indicator"""
        self._scanning_dots.stop()
        self._scanning_dots.pack_forget()
        self.progress_label.pack_forget()

    def _update_progress_message(self, message: str):
        """Update the progress message"""
        self.progress_label.configure(text=message)

    # =========================================================================
    # Model Connection
    # =========================================================================

    def _on_refresh_models(self):
        """Refresh the list of available models"""
        self.refresh_btn.set_enabled(False)
        self._show_progress("⏳ Scanning for Power BI models...")
        self._log_message("⏳ Scanning for local Power BI models...")

        # Also trigger cloud preload in parallel (for cloud browser dialog)
        self._trigger_cloud_preload()

        def discover_thread():
            try:
                # Try fast discovery first
                self.frame.after(0, lambda: self._update_progress_message("Scanning Connections"))
                models = self.connector.discover_local_models_fast()
                if not models:
                    self.frame.after(0, lambda: self._update_progress_message("Scanning Connections"))
                    models = self.connector.discover_local_models(quick_scan=True)

                # Check for WMI failures (can cause thin reports to show incorrect connection info)
                wmi_failures = self.connector.get_wmi_failure_count()

                self.frame.after(0, lambda: self._update_model_list(models))

                # Warn user if WMI queries failed
                if wmi_failures > 0:
                    self.frame.after(0, lambda: self._warn_wmi_failures(wmi_failures))
            except Exception as e:
                self.frame.after(0, lambda: self._log_message(f"❌ Error discovering models: {e}"))
            finally:
                self.frame.after(0, lambda: self.refresh_btn.set_enabled(True))
                self.frame.after(0, lambda: self._hide_progress())

        threading.Thread(target=discover_thread, daemon=True).start()

    def _update_model_list(self, models):
        """Update the model dropdown with discovered models"""
        # Remember current selection to preserve it if still available
        current_selection = self.selected_model.get()

        # Update the shared cache for other tools to use
        cache = get_local_model_cache()
        cache.set_models(models)

        if models:
            self._display_to_connection = {}
            self._dropdown_models_cache = list(models)  # Store full model objects
            display_names = []

            for model in models:
                conn_str = f"{model.server}|{model.database_name}"
                self._display_to_connection[model.display_name] = conn_str
                display_names.append(model.display_name)

            self.model_combo['values'] = display_names

            # Restore previous selection if still available, otherwise select first
            if current_selection and current_selection in display_names:
                idx = display_names.index(current_selection)
                self.model_combo.current(idx)
            elif display_names:
                self.model_combo.current(0)

            self.connect_btn.set_enabled(True)  # Enable connect when model available
            self._log_message(f"Found {len(models)} model(s)")
        else:
            self.model_combo['values'] = []
            self._dropdown_models_cache = []
            self.connect_btn.set_enabled(False)  # Disable connect when no models
            self._log_message("No models found. Ensure Power BI Desktop is running.")

    def _warn_wmi_failures(self, failure_count: int):
        """Warn user about WMI query failures that may affect thin report detection."""
        self._log_message(f"WMI query failed for {failure_count} process(es) - thin report info may be incomplete")
        self._log_message("  If thin reports show incorrect connection info, try:")
        self._log_message("  1. Restart Power BI Desktop")
        self._log_message("  2. Restart Windows if the issue persists")

    def _on_model_selection_changed(self, event=None):
        """Handle model dropdown selection change - enable connect if different from current"""
        # Clear text selection immediately to remove highlight after dropdown closes
        self.model_combo.selection_clear()

        selected = self.selected_model.get()
        if not selected:
            self.connect_btn.set_enabled(False)
            return

        # Enable CONNECT button when a model is selected
        self.connect_btn.set_enabled(True)

        # If we're connected (any connected state), log message about switching
        # The _on_connect method will auto-disconnect first
        current_status = self.connection_status.get()
        connected_states = ["Connected", "Thin Report (Local)", "Thin Report (Cloud)"]
        if current_status in connected_states:
            self._log_message(f"Model changed - click CONNECT to switch to {selected}")

    def _on_connect(self):
        """Connect to the selected model"""
        selected = self.selected_model.get()
        if not selected or selected not in self._display_to_connection:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Model Selected", "Please select a model to connect to.")
            return

        # If already connected (any connected state), disconnect first
        # This enables switching models without manual disconnect
        current_status = self.connection_status.get()
        connected_states = ["Connected", "Thin Report (Local)", "Thin Report (Cloud)"]
        if current_status in connected_states:
            self._log_message("Switching models - disconnecting from current...")
            self._quick_disconnect()

        conn_str = self._display_to_connection[selected]
        server, database = conn_str.split('|', 1)

        # Find the model in cache and store file path for preset persistence
        self._current_model_file_path = None
        for model in self._dropdown_models_cache:
            if model.display_name == selected:
                self._current_model_file_path = model.file_path
                break

        # Check if this is a thin report (live-connected report with no local database)
        if database.startswith('__thin_report__:'):
            # Trigger cloud preload immediately since thin reports open cloud dialog
            self._trigger_cloud_preload()
            self._handle_thin_report_connection(selected, server, database)
            return

        self._show_progress("Connecting to model...")
        self._log_message(f"Connecting to {selected}...")
        self.connect_btn.set_enabled(False)

        def connect_thread():
            try:
                self.frame.after(0, lambda: self._update_progress_message("Establishing TOM connection..."))
                success = self.connector.connect(server, database)

                def on_success():
                    self._update_status_display('success', "Connected")
                    self.disconnect_btn.set_enabled(True)
                    self.connect_btn.set_enabled(False)

                    # Initialize logic components
                    self.detector = ConnectionDetector(self.connector)
                    self.swapper = ConnectionSwapper(self.connector)
                    self.matcher = LocalModelMatcher(self.connector)
                    # Only create cloud browser if not already preloaded (preserves auth/cache)
                    if not self.cloud_browser:
                        self.cloud_browser = CloudWorkspaceBrowser()
                    self.schema_validator = SchemaValidator(self.connector)

                    # Pre-populate matcher cache with models from dropdown (avoids redundant port scan)
                    if hasattr(self, '_dropdown_models_cache') and self._dropdown_models_cache:
                        self.matcher.set_cache(self._dropdown_models_cache)

                    # Initialize health checker with callback
                    self.health_checker = ConnectionHealthChecker(
                        check_interval=30,
                        on_status_change=self._on_health_status_change
                    )

                    # Update dropdown with actual model name from TOM (for external tool launch)
                    actual_model_name = self._get_model_display_name()
                    display_name = selected
                    if actual_model_name and "External Tool Model" in selected:
                        # Update dropdown and mapping with friendly name
                        port = server.split(':')[-1] if ':' in server else server
                        display_name = f"{actual_model_name} ({port})"
                        # Update the dropdown
                        self._display_to_connection[display_name] = f"{server}|{database}"
                        self.model_combo['values'] = [display_name]
                        self.model_combo.current(0)
                        self.selected_model.set(display_name)
                        # Remove old generic entry
                        if selected in self._display_to_connection and selected != display_name:
                            del self._display_to_connection[selected]

                    self._update_progress_message("⏳ Detecting connections...")
                    self._log_message(f"✅ Connected to {display_name}")

                    # Detect connections (this will hide progress when done)
                    self._detect_connections()

                def on_failure():
                    self._hide_progress()
                    self.connect_btn.set_enabled(True)
                    self._log_message(f"❌ Failed to connect to {selected}")

                if success:
                    self.frame.after(0, on_success)
                else:
                    self.frame.after(0, on_failure)

            except Exception as e:
                self.frame.after(0, lambda: self._hide_progress())
                self.frame.after(0, lambda: self.connect_btn.set_enabled(True))
                self.frame.after(0, lambda: self._log_message(f"❌ Connection error: {e}"))

        threading.Thread(target=connect_thread, daemon=True).start()

    def _handle_thin_report_connection(self, selected: str, local_server: str, database_info: str):
        """
        Handle connection to a thin report (live-connected report).

        Thin reports have no local TOM database - they're pass-through connections
        to a cloud semantic model. Instead of connecting via TOM, we:
        1. Store thin report context for file modification
        2. Create a virtual connection row showing the current connection (cloud OR local)
        3. Allow the normal preset/Quick Swap UI to work
        4. Execute swaps by modifying the PBIX/PBIP file

        Args:
            selected: Display name of the selected model
            local_server: Local server address (localhost:port)
            database_info: Special format: __thin_report__:cloud_server:cloud_database
                           For local thin reports: __thin_report__:: (empty cloud info)
        """
        # Parse the cloud connection info from the special database format
        # Format: __thin_report__:cloud_server:cloud_database
        # For local thin reports: __thin_report__:: (empty cloud info)
        # Note: cloud_server may contain colons (e.g., powerbi://api.powerbi.com/...)
        # So we split after the prefix, then split the remainder by the LAST colon
        if database_info.startswith('__thin_report__:'):
            remainder = database_info[len('__thin_report__:'):]
            # Handle empty case: remainder = ':' means no cloud info (local thin report)
            if remainder == ':' or not remainder:
                cloud_server = ''
                cloud_database = ''
            else:
                # Find the last colon to separate server from database
                last_colon_idx = remainder.rfind(':')
                if last_colon_idx > 0:
                    cloud_server = remainder[:last_colon_idx]
                    cloud_database = remainder[last_colon_idx + 1:]
                else:
                    cloud_server = remainder
                    cloud_database = ''
        else:
            cloud_server = ''
            cloud_database = ''

        # Determine if this thin report is currently connected to LOCAL (already swapped)
        is_currently_local = not cloud_server and not cloud_database

        self._show_progress("Loading thin report...")
        self._log_message(f"Detected thin report: {selected}")
        self._log_message(f"  Local AS port: {local_server}")
        if is_currently_local:
            self._log_message("  Currently connected to: LOCAL model")

        # Find the thin report model in cache to get file_path and process_id
        thin_report_model = None
        for model in self._dropdown_models_cache:
            if model.is_thin_report and selected in model.display_name:
                thin_report_model = model
                break

        if not thin_report_model:
            self._hide_progress()
            self._log_message("Could not find thin report model in cache")
            self._show_thin_report_help_message(selected, cloud_server, cloud_database,
                                                 "Could not find thin report model info")
            return

        # Read current connection from file to get perspective_name (if any)
        source_perspective_name = None
        source_cloud_connection_type = None
        if thin_report_model.file_path:
            from tools.connection_hotswap.logic.pbix_modifier import PbixModifier
            modifier = PbixModifier()
            file_conn = modifier.read_current_connection(thin_report_model.file_path)
            if file_conn:
                source_perspective_name = file_conn.get('perspective_name')
                source_cloud_connection_type = file_conn.get('cloud_connection_type')
                if source_perspective_name:
                    self._log_message(f"  Perspective: {source_perspective_name}")

        # For local thin reports, try to retrieve original cloud connection from cache
        original_cloud_server = ''
        original_cloud_database = ''
        original_cloud_friendly_name = ''
        if is_currently_local and thin_report_model.file_path:
            from tools.connection_hotswap.logic.pbix_modifier import PbixModifier
            modifier = PbixModifier()
            cached_cloud = modifier.get_cached_cloud_connection(thin_report_model.file_path)
            if cached_cloud:
                original_cloud_server = cached_cloud.get('server', '')
                original_cloud_database = cached_cloud.get('database', '')
                original_cloud_friendly_name = cached_cloud.get('friendly_name', '')
                self._log_message(f"  Original cloud (cached): {original_cloud_server}")
            else:
                self._log_message("  Original cloud connection not in cache (will need cloud login to swap back)")

        # Store thin report context for swap execution
        self._thin_report_context = {
            'file_path': thin_report_model.file_path,
            'process_id': thin_report_model.process_id,
            'cloud_server': cloud_server,
            'cloud_database': cloud_database,
            'local_server': local_server,
            'report_name': selected,
            'is_currently_local': is_currently_local,
            'original_cloud_server': original_cloud_server,
            'original_cloud_database': original_cloud_database,
            'original_cloud_friendly_name': original_cloud_friendly_name,
            'source_perspective_name': source_perspective_name,
            'source_cloud_connection_type': source_cloud_connection_type,
        }

        if thin_report_model.file_path:
            self._log_message(f"  File: {thin_report_model.file_path}")
        if cloud_server:
            self._log_message(f"  Cloud endpoint: {cloud_server}")
        if cloud_database:
            self._log_message(f"  Cloud dataset: {cloud_database}")

        # Update UI to show connected state
        if is_currently_local:
            self._update_status_display('success', "Thin Report (Local)")
        else:
            self._update_status_display('success', "Thin Report (Cloud)")
        self.disconnect_btn.set_enabled(True)
        self.connect_btn.set_enabled(False)

        # Force UI refresh immediately so status displays properly before any blocking calls
        self.frame.update_idletasks()

        # Initialize matcher for local model selection (we don't need full TOM connection)
        self.matcher = LocalModelMatcher(self.connector)
        # Only create cloud browser if not already preloaded (preserves auth/cache)
        if not self.cloud_browser:
            self.cloud_browser = CloudWorkspaceBrowser()

        # For cloud thin reports, cloud_database might be a GUID instead of friendly name
        # Try to resolve it using cloud browser cache (if available)
        # This is done asynchronously to prevent UI freezing during cloud auth
        if not is_currently_local and cloud_database and self._is_guid_format(cloud_database):
            # Start async resolution - UI stays responsive
            self._resolve_cloud_guid_async(
                cloud_database=cloud_database,
                cloud_server=cloud_server,
                is_currently_local=is_currently_local,
                thin_report_model=thin_report_model,
                source_perspective_name=source_perspective_name,
                source_cloud_connection_type=source_cloud_connection_type,
                original_cloud_server=original_cloud_server,
                original_cloud_database=original_cloud_database,
                original_cloud_friendly_name=original_cloud_friendly_name,
                local_server=local_server,
                selected=selected,
            )
            return  # Async path - completion handler will finish setup

        # Non-GUID path: complete setup synchronously
        display_name = cloud_database
        resolved_workspace_name = None
        self._complete_thin_report_setup(
            display_name=display_name,
            resolved_workspace_name=resolved_workspace_name,
            cloud_database=cloud_database,
            cloud_server=cloud_server,
            is_currently_local=is_currently_local,
            thin_report_model=thin_report_model,
            source_perspective_name=source_perspective_name,
            source_cloud_connection_type=source_cloud_connection_type,
            original_cloud_server=original_cloud_server,
            original_cloud_database=original_cloud_database,
            original_cloud_friendly_name=original_cloud_friendly_name,
            local_server=local_server,
            selected=selected,
        )

    def _resolve_cloud_guid_async(
        self,
        cloud_database: str,
        cloud_server: str,
        is_currently_local: bool,
        thin_report_model,
        source_perspective_name: Optional[str],
        source_cloud_connection_type: Optional[str],
        original_cloud_server: str,
        original_cloud_database: str,
        original_cloud_friendly_name: str,
        local_server: str,
        selected: str,
    ):
        """
        Resolve cloud GUID to friendly name in a background thread.

        This prevents UI freezing during silent authentication and workspace enumeration.
        The completion callback runs on the main thread to update UI.
        """
        # Show progress message - UI stays responsive, colored dots provide animation
        self._update_progress_message("Resolving cloud connection")
        self.frame.update_idletasks()

        def resolve_thread():
            """Background thread for cloud GUID resolution."""
            try:
                # Try silent resolution first (may take ~30s if fetching workspaces)
                resolved_name, resolved_workspace = self._try_resolve_guid_to_name(cloud_database)

                # Schedule callback on main thread
                self.frame.after(0, lambda: self._on_cloud_guid_resolved(
                    resolved_name=resolved_name,
                    resolved_workspace=resolved_workspace,
                    cloud_database=cloud_database,
                    cloud_server=cloud_server,
                    is_currently_local=is_currently_local,
                    thin_report_model=thin_report_model,
                    source_perspective_name=source_perspective_name,
                    source_cloud_connection_type=source_cloud_connection_type,
                    original_cloud_server=original_cloud_server,
                    original_cloud_database=original_cloud_database,
                    original_cloud_friendly_name=original_cloud_friendly_name,
                    local_server=local_server,
                    selected=selected,
                ))
            except Exception as e:
                self.logger.error(f"Cloud GUID resolution error: {e}")
                # On error, continue with GUID as display name
                self.frame.after(0, lambda: self._on_cloud_guid_resolved(
                    resolved_name=None,
                    resolved_workspace=None,
                    cloud_database=cloud_database,
                    cloud_server=cloud_server,
                    is_currently_local=is_currently_local,
                    thin_report_model=thin_report_model,
                    source_perspective_name=source_perspective_name,
                    source_cloud_connection_type=source_cloud_connection_type,
                    original_cloud_server=original_cloud_server,
                    original_cloud_database=original_cloud_database,
                    original_cloud_friendly_name=original_cloud_friendly_name,
                    local_server=local_server,
                    selected=selected,
                ))

        threading.Thread(target=resolve_thread, daemon=True).start()

    def _on_cloud_guid_resolved(
        self,
        resolved_name: Optional[str],
        resolved_workspace: Optional[str],
        cloud_database: str,
        cloud_server: str,
        is_currently_local: bool,
        thin_report_model,
        source_perspective_name: Optional[str],
        source_cloud_connection_type: Optional[str],
        original_cloud_server: str,
        original_cloud_database: str,
        original_cloud_friendly_name: str,
        local_server: str,
        selected: str,
    ):
        """
        Callback when cloud GUID resolution completes (runs on main thread).

        If silent resolution failed, prompts user for interactive login.
        Then continues with thin report setup.
        """
        display_name = cloud_database
        resolved_workspace_name = resolved_workspace

        if resolved_name:
            display_name = resolved_name
            self._log_message(f"  Cloud model: {resolved_name} (resolved from GUID)")
        elif self._is_guid_format(cloud_database):
            # GUID detected but not resolved - offer interactive cloud login
            self._log_message(f"  Cloud dataset appears to be a GUID: {cloud_database}")
            self._update_progress_message("GUID detected")

            # Prompt user to login to cloud to get the friendly name (interactive, main thread)
            resolved_name, resolved_workspace_name = self._prompt_cloud_login_for_guid(cloud_database, cloud_server)
            if resolved_name:
                display_name = resolved_name
                self._log_message(f"  Resolved via cloud: {resolved_name}")

        # Continue with thin report setup
        self._complete_thin_report_setup(
            display_name=display_name,
            resolved_workspace_name=resolved_workspace_name,
            cloud_database=cloud_database,
            cloud_server=cloud_server,
            is_currently_local=is_currently_local,
            thin_report_model=thin_report_model,
            source_perspective_name=source_perspective_name,
            source_cloud_connection_type=source_cloud_connection_type,
            original_cloud_server=original_cloud_server,
            original_cloud_database=original_cloud_database,
            original_cloud_friendly_name=original_cloud_friendly_name,
            local_server=local_server,
            selected=selected,
        )

    def _complete_thin_report_setup(
        self,
        display_name: str,
        resolved_workspace_name: Optional[str],
        cloud_database: str,
        cloud_server: str,
        is_currently_local: bool,
        thin_report_model,
        source_perspective_name: Optional[str],
        source_cloud_connection_type: Optional[str],
        original_cloud_server: str,
        original_cloud_database: str,
        original_cloud_friendly_name: str,
        local_server: str,
        selected: str,
    ):
        """
        Complete thin report setup after GUID resolution (if any).

        This contains the logic that was previously in _handle_thin_report_connection
        after the GUID resolution step.
        """
        # Update thin report context with resolved display name
        self._thin_report_context['display_name'] = display_name

        # If we resolved workspace from cloud cache, update Last Config if it exists but lacks workspace
        if resolved_workspace_name:
            model_hash = self._get_model_hash()
            if model_hash:
                last_config = self.preset_manager.get_last_config(model_hash)
                if last_config and not last_config.get('workspace_name'):
                    self.preset_manager.update_last_config_workspace(
                        model_hash,
                        workspace_name=resolved_workspace_name,
                        friendly_name=display_name
                    )
                    self._log_message(f"Updated Last Config with workspace: {resolved_workspace_name}")

        # Create a virtual DataSourceConnection representing the CURRENT connection
        # For cloud thin reports: show cloud as source
        # For local thin reports: show local as source
        if is_currently_local:
            # Read the actual connection from file to get the real server:database
            # that the thin report is connected to (e.g., localhost:62330, not the thin report's port)
            actual_local_server = local_server
            actual_local_database = None
            local_model_name = None

            if thin_report_model.file_path:
                from tools.connection_hotswap.logic.pbix_modifier import PbixModifier
                modifier = PbixModifier()
                current_conn = modifier.read_current_connection(thin_report_model.file_path)
                if current_conn:
                    actual_local_server = current_conn.get('server', local_server)
                    actual_local_database = current_conn.get('database', '')
                    self._log_message(f"  Connected to: {actual_local_server}")

            # Find the local model this is connected to for display purposes
            # Match by server (the actual model's port, not the thin report's port)
            for model in self._dropdown_models_cache:
                if not model.is_thin_report and model.server == actual_local_server:
                    local_model_name = model.display_name
                    self._log_message(f"  Local model: {local_model_name}")
                    break

            virtual_connection = DataSourceConnection(
                name=local_model_name or f"Local Model ({actual_local_server})",
                connection_type=ConnectionType.LIVE_CONNECTION,
                server=actual_local_server,
                database=actual_local_database or local_model_name or "local_model",
                is_cloud=False,
                connection_string=f"Data Source={actual_local_server}",
                dataset_name=local_model_name,
                workspace_name=None,
                tom_reference_type=TomReferenceType.LIVE_CONNECTION_MODEL
            )
        else:
            # Use resolved workspace name if available, otherwise try to extract from URL
            source_workspace = resolved_workspace_name or self._extract_workspace_from_server(cloud_server)

            # Additional fallback: get workspace from Last Config if not resolved yet
            if not source_workspace:
                model_hash = self._get_model_hash()
                if model_hash:
                    last_config = self.preset_manager.get_last_config(model_hash)
                    if last_config:
                        source_workspace = last_config.get('workspace_name', '')
                        if source_workspace:
                            self._log_message(f"Using workspace from Last Config: {source_workspace}")

            # Convert cloud_connection_type string to enum
            source_cloud_type_enum = None
            if source_cloud_connection_type:
                from tools.connection_hotswap.models import CloudConnectionType
                if source_cloud_connection_type == 'pbi_semantic_model':
                    source_cloud_type_enum = CloudConnectionType.PBI_SEMANTIC_MODEL
                elif source_cloud_connection_type == 'aas_xmla':
                    source_cloud_type_enum = CloudConnectionType.AAS_XMLA

            virtual_connection = DataSourceConnection(
                name=display_name or "Cloud Connection",
                connection_type=ConnectionType.LIVE_CONNECTION,
                server=cloud_server or "Unknown Cloud",
                database=cloud_database or "",
                is_cloud=True,
                connection_string=f"Data Source={cloud_server};Initial Catalog={cloud_database}",
                dataset_name=display_name,
                workspace_name=source_workspace,
                perspective_name=source_perspective_name,
                cloud_connection_type=source_cloud_type_enum,
                tom_reference_type=TomReferenceType.LIVE_CONNECTION_MODEL
            )

        # Override is_swappable check - thin reports ARE swappable via file modification
        # We'll handle this specially in the swap execution
        virtual_connection._is_thin_report_virtual = True  # Custom flag

        # Create model info with the virtual connection
        self.model_info = ModelConnectionInfo(
            model_name=selected,
            server=local_server,
            database=cloud_database or "thin_report",
            connection_type=ConnectionType.LIVE_CONNECTION,
            connections=[virtual_connection],
            is_composite=False,
            total_datasources=1,
            swappable_count=1
        )

        # Update Connection Details panel
        self._update_connection_details()

        # Create mapping for the virtual connection
        virtual_mapping = ConnectionMapping(
            source=virtual_connection,
            target=None,
            status=SwapStatus.PENDING
        )

        self._log_message("")
        if is_currently_local:
            self._log_message("Thin report loaded (currently LOCAL).")
            if original_cloud_server:
                self._log_message("Original cloud connection available - select target or use LAST CONFIG to swap back.")
            else:
                self._log_message("Select a cloud target to swap back, or use LAST CONFIG if available.")
        else:
            self._log_message("Thin report loaded. Select a local model target to swap.")
        self._log_message("Presets and Quick Swap are available.")

        # Use cached local models from dropdown (populated by REFRESH button)
        # This avoids redundant port scanning
        self._update_progress_message("Getting local models")
        if hasattr(self, '_dropdown_models_cache') and self._dropdown_models_cache:
            local_models = [m for m in self._dropdown_models_cache if not m.is_thin_report]
            self._log_message(f"Using {len(local_models)} cached local model(s)")
        else:
            # Fallback to scanning only if no cache available
            self._log_message("No cached models, scanning for local Power BI models...")
            local_models = self.matcher.discover_local_models()
            self._log_message(f"Found {len(local_models)} local model(s)")

        # For LOCAL thin reports with cached original cloud, auto-set as target for swap-back
        if is_currently_local and original_cloud_server and original_cloud_database:
            # Build friendly display name
            workspace_name = self._extract_workspace_from_server(original_cloud_server)

            # Get Last Config for fallback values (friendly name and workspace)
            model_hash = self._get_model_hash()
            last_config = None
            if model_hash:
                last_config = self.preset_manager.get_last_config(model_hash)

            # If no friendly name from in-memory cache, try to get from Last Config
            if not original_cloud_friendly_name and last_config:
                original_cloud_friendly_name = last_config.get('friendly_name', '')

            # Always try workspace from Last Config if not extracted from URL
            if not workspace_name and last_config:
                workspace_name = last_config.get('workspace_name', '')
                if workspace_name:
                    self._log_message(f"Using workspace from Last Config: {workspace_name}")

            if original_cloud_friendly_name:
                # Extract workspace name from server URL for display
                if workspace_name:
                    display = f"Original: {original_cloud_friendly_name} ({workspace_name})"
                else:
                    display = f"Original: {original_cloud_friendly_name}"
            else:
                display = f"Original Cloud ({original_cloud_database})"

            # Get cloud connection type and perspective from cached connection
            from tools.connection_hotswap.models import CloudConnectionType
            from tools.connection_hotswap.logic.pbix_modifier import PbixModifier
            modifier = PbixModifier()
            cloud_conn_type = None
            perspective_name = None
            if thin_report_model.file_path:
                cached_cloud = modifier.get_cached_cloud_connection(thin_report_model.file_path)
                if cached_cloud:
                    # Extract cloud connection type from cached data
                    cached_type = cached_cloud.get('cloud_connection_type')
                    if cached_type:
                        try:
                            cloud_conn_type = CloudConnectionType(cached_type)
                        except ValueError:
                            pass
                    # Extract perspective from cached data
                    perspective_name = cached_cloud.get('perspective_name')

            target = SwapTarget(
                target_type="cloud",
                server=original_cloud_server,
                database=original_cloud_database,
                display_name=display,
                workspace_name=workspace_name,
                cloud_connection_type=cloud_conn_type,
                perspective_name=perspective_name,
            )
            virtual_mapping.target = target
            virtual_mapping.auto_matched = True
            virtual_mapping.status = SwapStatus.READY
            self._log_message(f"Auto-set target to original cloud: {original_cloud_server}")
        # For CLOUD thin reports, auto-match if enabled
        elif not is_currently_local and self.auto_match_enabled.get() and display_name and local_models:
            self._update_progress_message("Auto-matching local models")
            match = self.matcher.find_matching_model(display_name, local_models)
            if match:
                target = SwapTarget(
                    target_type="local",
                    server=match.server,
                    database=match.database_name,
                    display_name=match.display_name,
                )
                virtual_mapping.target = target
                virtual_mapping.auto_matched = True
                virtual_mapping.status = SwapStatus.READY
                self._log_message(f"Auto-matched '{display_name}' -> '{match.display_name}'")

        # Populate the mapping table
        self._populate_mapping_table([virtual_mapping])

        # Auto-save Last Config on first connect (if not already saved)
        # This allows users to restore their original state after swapping
        # MUST be done BEFORE checking if Last Config exists for button enable
        self._save_last_config_if_new()

        # Enable LAST CONFIG button if saved config exists
        model_hash = self._get_model_hash()
        if model_hash and self.preset_manager.has_last_config(model_hash):
            self.last_config_btn.set_enabled(True)
            self._log_message("Saved target available - use LAST CONFIG to restore original connection")

        # Enable Rollback button if swap history exists from previous sessions
        if self._swap_history:
            self.rollback_btn.set_enabled(True)

        # Load model-specific presets
        self._load_model_presets_on_connect()

        # Refresh local models for Quick Swap panel
        self._refresh_local_models()

        # Hide progress - thin report is ready
        self._hide_progress()

    def _try_resolve_guid_to_name(self, potential_guid: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Try to resolve a GUID to a friendly dataset name using multiple sources.

        PBIX thin reports store the semantic model GUID instead of the friendly name.
        This method tries to resolve it using:
        1. Thin report context (cached cloud connection info from previous swap)
        2. Cloud browser cache (from browsing workspaces in this session)
        3. Last Config data (from previous connections)

        Args:
            potential_guid: A string that might be a GUID (dataset ID)

        Returns:
            Tuple of (friendly_dataset_name, workspace_name) if found, (None, None) otherwise
        """
        import re

        # Check if it looks like a GUID (8-4-4-4-12 hex pattern)
        guid_pattern = r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
        if not re.match(guid_pattern, potential_guid.lower()):
            return None, None  # Not a GUID, return None

        # Source 1: Check thin report context (most reliable for swapped thin reports)
        resolved_from_thin_report = None
        if self._thin_report_context:
            # Check if the GUID matches the original cloud database
            original_db = self._thin_report_context.get('original_cloud_database', '')
            if original_db and original_db.lower() == potential_guid.lower():
                friendly_name = self._thin_report_context.get('original_cloud_friendly_name', '')
                if friendly_name and friendly_name != potential_guid:
                    self.logger.info(f"Resolved GUID {potential_guid} -> {friendly_name} (from thin report cache)")
                    resolved_from_thin_report = friendly_name
            # Also check current cloud_database
            if not resolved_from_thin_report:
                current_db = self._thin_report_context.get('cloud_database', '')
                if current_db and current_db.lower() == potential_guid.lower():
                    display_name = self._thin_report_context.get('display_name', '')
                    if display_name and display_name != potential_guid:
                        self.logger.info(f"Resolved GUID {potential_guid} -> {display_name} (from thin report context)")
                        resolved_from_thin_report = display_name

        # Source 2: Look up in cloud browser (with silent auth if needed)
        # This will try to authenticate silently and fetch workspaces if cache is empty
        if hasattr(self, 'cloud_browser') and self.cloud_browser:
            cloud_name, cloud_workspace = self.cloud_browser.lookup_dataset_by_guid(
                potential_guid, fetch_if_needed=True
            )
            if cloud_workspace:
                # If we have a name from thin report cache, use it with the workspace from cloud
                if resolved_from_thin_report:
                    self.logger.info(f"Using thin report name '{resolved_from_thin_report}' with workspace '{cloud_workspace}' from cloud lookup")
                    return resolved_from_thin_report, cloud_workspace
                # Otherwise use the cloud name
                if cloud_name:
                    self.logger.info(f"Resolved GUID {potential_guid} -> {cloud_name} with workspace '{cloud_workspace}' (from cloud lookup)")
                    return cloud_name, cloud_workspace
            elif cloud_name and not resolved_from_thin_report:
                # Got name but no workspace
                self.logger.info(f"Resolved GUID {potential_guid} -> {cloud_name} (from cloud lookup, no workspace)")
                return cloud_name, None

        # Source 3: Check Last Config for this model (least reliable after swap)
        model_hash = self._get_model_hash()
        if model_hash:
            last_config = self.preset_manager.get_last_config(model_hash)
            if last_config:
                friendly_name = last_config.get('friendly_name', '')
                workspace_name = last_config.get('workspace_name', '')
                # If we have name from thin report but need workspace, check Last Config for workspace
                if resolved_from_thin_report and workspace_name:
                    self.logger.info(f"Using thin report name '{resolved_from_thin_report}' with workspace '{workspace_name}' from Last Config")
                    return resolved_from_thin_report, workspace_name
                # Only use Last Config if friendly_name looks like a real name (not a local model name)
                if friendly_name and friendly_name != potential_guid and not friendly_name.startswith('local'):
                    # Strip port number suffix if present (e.g., "Cereal Model (62330)" -> "Cereal Model")
                    # Port numbers are 5-digit numbers in parentheses at end, not workspace names
                    import re
                    port_pattern = r'\s*\(\d{4,5}\)$'
                    if re.search(port_pattern, friendly_name):
                        friendly_name = re.sub(port_pattern, '', friendly_name).strip()
                    self.logger.info(f"Resolved GUID {potential_guid} -> {friendly_name} (from Last Config)")
                    return friendly_name, workspace_name or None

        # If we have a name from thin report cache but couldn't find workspace, still return the name
        if resolved_from_thin_report:
            self.logger.info(f"Returning thin report name '{resolved_from_thin_report}' without workspace")
            return resolved_from_thin_report, None

        return None, None  # Not found in any cache

    def _is_guid_format(self, value: str) -> bool:
        """Check if a string matches GUID format (8-4-4-4-12 hex pattern)."""
        import re
        if not value:
            return False
        guid_pattern = r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
        return bool(re.match(guid_pattern, value.lower()))

    def _get_source_cloud_connection_type(self, source: 'DataSourceConnection') -> 'CloudConnectionType':
        """
        Determine the CloudConnectionType from a source DataSourceConnection.

        Uses the server string and connection type to detect whether the source
        is using PBI Semantic Model connector or AAS XMLA endpoint.

        Args:
            source: The source DataSourceConnection

        Returns:
            CloudConnectionType for the source, defaults to PBI_SEMANTIC_MODEL
        """
        from tools.connection_hotswap.models import CloudConnectionType

        if not source or not source.is_cloud:
            return CloudConnectionType.PBI_SEMANTIC_MODEL

        server_lower = source.server.lower() if source.server else ''

        # Check for XMLA endpoint indicators in the server string
        # AAS uses powerbi:// with analysisServicesDatabaseLive or asazure://
        if 'asazure://' in server_lower:
            return CloudConnectionType.AAS_XMLA

        # Check connection type - analysisServicesDatabaseLive with cloud server = XMLA
        conn_type_str = source.connection_type.value if source.connection_type else ''
        if conn_type_str.lower() == 'liveconnection' and 'powerbi://' in server_lower:
            # Power BI XMLA endpoint - could be either, check if source has explicit type
            if hasattr(source, 'cloud_connection_type') and source.cloud_connection_type:
                return source.cloud_connection_type
            # Default heuristic: if using XMLA endpoint format, assume AAS
            # Users connecting via Premium XMLA tend to need XMLA features
            return CloudConnectionType.AAS_XMLA

        # Default to PBI Semantic Model connector
        return CloudConnectionType.PBI_SEMANTIC_MODEL

    def _resolve_guid_display_names(self):
        """
        Post-process mappings to resolve any GUID display names to friendly names.

        This is used after applying Last Config or presets to ensure that targets
        with GUID databases show friendly names instead of the raw GUID.
        """
        if not self.mappings:
            return

        for mapping in self.mappings:
            if not mapping.target:
                continue

            # Check if the display_name contains what looks like a GUID
            display_name = mapping.target.display_name or ""

            # Extract potential GUID from display_name (e.g., "Original: abc-123-xyz")
            # or from database field
            potential_guid = mapping.target.database
            if self._is_guid_format(potential_guid):
                # Try to resolve the GUID to a friendly name
                resolved, resolved_ws = self._try_resolve_guid_to_name(potential_guid)
                if resolved:
                    # Update the display name with the resolved friendly name
                    # Use resolved workspace if available, otherwise use existing
                    workspace = resolved_ws or mapping.target.workspace_name
                    if resolved_ws and not mapping.target.workspace_name:
                        mapping.target.workspace_name = resolved_ws
                    if "Original:" in display_name:
                        if workspace:
                            mapping.target.display_name = f"Original: {resolved} ({workspace})"
                        else:
                            mapping.target.display_name = f"Original: {resolved}"
                    else:
                        if workspace:
                            mapping.target.display_name = f"{resolved} ({workspace})"
                        else:
                            mapping.target.display_name = resolved
                    self.logger.debug(f"Resolved GUID display name: {potential_guid} -> {resolved}")

    def _prompt_cloud_login_for_guid(self, guid: str, cloud_server: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Prompt user to login to cloud to resolve a GUID to a friendly dataset name.

        For PBIX thin reports, the dataset ID is stored as a GUID. This method
        offers the user a chance to authenticate with the cloud to look up the
        friendly name, which enables better auto-matching with local models.

        Args:
            guid: The GUID to resolve
            cloud_server: The cloud server URL (powerbi://...)

        Returns:
            Tuple of (friendly_dataset_name, workspace_name) if resolved, (None, None) otherwise
        """
        # Ask user if they want to login to cloud to resolve the model name
        result = self._show_cloud_login_prompt_dialog(guid)

        if not result:
            self._log_message("  User declined cloud login for GUID resolution")
            return None, None

        # Try to extract workspace from the cloud server URL
        workspace_name = self._extract_workspace_from_server(cloud_server)

        try:
            self._log_message("  Authenticating with Power BI...")

            # Initialize cloud browser
            if not hasattr(self, 'cloud_browser') or not self.cloud_browser:
                self.cloud_browser = CloudWorkspaceBrowser()

            # Use threaded authentication with progress dialog
            auth_result = self._authenticate_with_progress_dialog()
            if not auth_result:
                self._log_message("  Authentication failed or cancelled")
                return None, None

            self._log_message("  Authentication successful")
            self._update_progress_message("Fetching workspaces...")

            # Now fetch workspaces
            workspaces, ws_error = self.cloud_browser.get_workspaces()
            if ws_error:
                self._log_message(f"  Error fetching workspaces: {ws_error}")
            if not workspaces:
                self._log_message("  Could not retrieve workspaces from Power BI")
                self._show_themed_info_dialog(
                    "No Workspaces Found",
                    "You are signed in but no workspaces were found.\n\n"
                    "You may not have access to any Power BI workspaces."
                )
                return None, None

            self._log_message(f"  Retrieved {len(workspaces)} workspace(s)")

            # If we know the workspace, fetch datasets from just that workspace
            if workspace_name:
                self._update_progress_message(f"Looking up model in {workspace_name}...")
                for ws in workspaces:
                    if ws.name.lower() == workspace_name.lower():
                        datasets, _ = self.cloud_browser.get_workspace_datasets(ws.id)
                        if datasets:
                            for dataset in datasets:
                                if dataset.id.lower() == guid.lower():
                                    self._log_message(f"  Found model: {dataset.name}")
                                    return dataset.name, ws.name

            # If workspace not found or dataset not in expected workspace,
            # search all workspaces
            self._update_progress_message("Searching all workspaces for model...")
            for ws in workspaces:
                datasets, _ = self.cloud_browser.get_workspace_datasets(ws.id)
                if datasets:
                    for dataset in datasets:
                        if dataset.id.lower() == guid.lower():
                            self._log_message(f"  Found model: {dataset.name} in {ws.name}")
                            return dataset.name, ws.name

            self._log_message(f"  Could not find model with ID {guid}")
            self._show_themed_info_dialog(
                "Model Not Found",
                f"Could not find a semantic model with ID:\n{guid}\n\n"
                f"The model may have been deleted or you may not have access to it."
            )
            return None, None

        except Exception as e:
            self.logger.exception("Error during cloud login for GUID resolution")
            self._log_message(f"  Error: {str(e)}")
            return None, None

    def _show_cloud_login_prompt_dialog(self, guid: str) -> bool:
        """
        Show a themed Yes/No dialog asking user to login to cloud.

        Args:
            guid: The GUID to display

        Returns:
            True if user clicked Yes, False otherwise
        """
        from pathlib import Path

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create dialog - use withdraw/deiconify pattern for dark title bar
        dialog = tk.Toplevel(self.frame)
        dialog.withdraw()  # Hide until fully configured
        dialog.title("PBIX Thin Report Detected")
        dialog.geometry("450x255")
        dialog.resizable(False, False)
        dialog.transient(self.frame.winfo_toplevel())
        dialog.configure(bg=colors['background'])

        # Set AE favicon icon
        try:
            base_dir = Path(__file__).parent.parent.parent.parent
            icon_path = base_dir / 'assets' / 'favicon.ico'
            if icon_path.exists():
                dialog.iconbitmap(str(icon_path))
        except Exception:
            pass

        # Set dark/light title bar on Windows (must be before deiconify)
        try:
            import ctypes
            dialog.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(dialog.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if is_dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass

        # Result variable
        result_var = tk.BooleanVar(value=False)

        # Content frame
        content = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        content.pack(fill=tk.BOTH, expand=True)

        # Section background
        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f8f8fc')

        # Info section
        info_frame = tk.Frame(content, bg=section_bg, padx=12, pady=10)
        info_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            info_frame,
            text="This PBIX thin report uses a model ID (GUID)\ninstead of a friendly name.",
            font=("Segoe UI", 10),
            bg=section_bg,
            fg=colors['text_primary'],
            justify='left'
        ).pack(anchor='w')

        # GUID display
        guid_frame = tk.Frame(content, bg=colors['background'])
        guid_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            guid_frame,
            text="Model ID:",
            font=("Segoe UI", 9),
            bg=colors['background'],
            fg=colors['text_muted']
        ).pack(anchor='w')

        tk.Label(
            guid_frame,
            text=guid,
            font=("Consolas", 9),
            bg=colors['background'],
            fg=colors.get('primary', '#4a6cf5')
        ).pack(anchor='w')

        # Question
        tk.Label(
            content,
            text="To enable auto-matching with local models, would you\nlike to sign in to Power BI to look up the model name?",
            font=("Segoe UI", 10),
            bg=colors['background'],
            fg=colors['text_primary'],
            justify='left'
        ).pack(anchor='w', pady=(0, 16))

        # Buttons - centered
        button_frame = tk.Frame(content, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        # Inner frame for centering buttons
        button_inner = tk.Frame(button_frame, bg=colors['background'])
        button_inner.pack(anchor='center')

        def on_yes():
            result_var.set(True)
            dialog.destroy()

        def on_no():
            result_var.set(False)
            dialog.destroy()

        RoundedButton(
            button_inner,
            text="Sign In",
            command=on_yes,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors.get('button_primary_pressed', colors['button_primary_hover']),
            fg='#ffffff',
            height=32, radius=5,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=colors['background'],
            width=90
        ).pack(side=tk.LEFT, padx=(0, 10))

        RoundedButton(
            button_inner,
            text="No",
            command=on_no,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=32, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background'],
            width=90
        ).pack(side=tk.LEFT)

        # Center on parent and show
        dialog.update_idletasks()
        parent = self.frame.winfo_toplevel()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 255) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.deiconify()
        dialog.grab_set()

        # Wait for dialog to close
        dialog.wait_window()
        return result_var.get()

    def _show_themed_info_dialog(self, title: str, message: str):
        """
        Show a themed info dialog.

        Args:
            title: Dialog title
            message: Message to display
        """
        from pathlib import Path

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create dialog - use withdraw/deiconify pattern for dark title bar
        dialog = tk.Toplevel(self.frame)
        dialog.withdraw()  # Hide until fully configured
        dialog.title(title)
        dialog.geometry("400x155")
        dialog.resizable(False, False)
        dialog.transient(self.frame.winfo_toplevel())
        dialog.configure(bg=colors['background'])

        # Set AE favicon icon
        try:
            base_dir = Path(__file__).parent.parent.parent.parent
            icon_path = base_dir / 'assets' / 'favicon.ico'
            if icon_path.exists():
                dialog.iconbitmap(str(icon_path))
        except Exception:
            pass

        # Set dark/light title bar on Windows (must be before deiconify)
        try:
            import ctypes
            dialog.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(dialog.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if is_dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass

        # Content frame
        content = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        content.pack(fill=tk.BOTH, expand=True)

        # Message
        tk.Label(
            content,
            text=message,
            font=("Segoe UI", 10),
            bg=colors['background'],
            fg=colors['text_primary'],
            justify='left',
            wraplength=360
        ).pack(anchor='w', pady=(0, 20))

        # OK button - centered
        button_frame = tk.Frame(content, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        button_inner = tk.Frame(button_frame, bg=colors['background'])
        button_inner.pack(anchor='center')

        RoundedButton(
            button_inner,
            text="OK",
            command=dialog.destroy,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=32, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background'],
            width=80
        ).pack()

        # Center on parent and show
        dialog.update_idletasks()
        parent = self.frame.winfo_toplevel()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 155) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.deiconify()
        dialog.grab_set()

    def _extract_workspace_from_server(self, server: str) -> Optional[str]:
        """Extract workspace name from a Power BI XMLA endpoint URL.

        Handles multiple URL formats:
        - powerbi://api.powerbi.com/v1.0/myorg/WorkspaceName
        - pbiazure://api.powerbi.com/v1.0/myorg/WorkspaceName
        - asazure://region.asazure.windows.net/server
        """
        if not server:
            return None

        server_lower = server.lower()

        # Check for Power BI URL formats (powerbi:// or pbiazure://)
        if 'powerbi://' in server_lower or 'pbiazure://' in server_lower:
            try:
                import urllib.parse
                # Format: powerbi://api.powerbi.com/v1.0/myorg/WorkspaceName
                # or: pbiazure://api.powerbi.com/v1.0/myorg/WorkspaceName
                if '/v1.0/' in server:
                    parts = server.split('/v1.0/')
                    if len(parts) > 1:
                        path_part = parts[1]
                        segments = path_part.split('/')
                        # Path is typically "myorg/WorkspaceName"
                        if len(segments) >= 2 and segments[0] == 'myorg':
                            return urllib.parse.unquote(segments[1])
                # Fallback: try last segment
                parts = server.split('/')
                if len(parts) >= 5:
                    return urllib.parse.unquote(parts[-1])
            except Exception:
                pass

        return None

    def _authenticate_with_progress_dialog(self) -> bool:
        """
        Show a progress dialog while authenticating with Power BI.

        Uses threading to allow MSAL to open the browser properly.

        Returns:
            True if authentication was successful, False otherwise
        """
        import threading
        from pathlib import Path

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create progress dialog
        dialog = tk.Toplevel(self.frame)
        dialog.withdraw()
        dialog.title("Signing In")
        dialog.geometry("350x120")
        dialog.resizable(False, False)
        dialog.transient(self.frame.winfo_toplevel())
        dialog.configure(bg=colors['background'])

        # Set AE favicon icon
        try:
            base_dir = Path(__file__).parent.parent.parent.parent
            icon_path = base_dir / 'assets' / 'favicon.ico'
            if icon_path.exists():
                dialog.iconbitmap(str(icon_path))
        except Exception:
            pass

        # Set dark/light title bar
        try:
            import ctypes
            dialog.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(dialog.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if is_dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass

        # Result tracking
        auth_result = {'success': False, 'message': ''}

        # Content
        content = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        content.pack(fill=tk.BOTH, expand=True)

        # Spinner characters
        spinner_chars = ['|', '/', '-', '\\']
        spinner_idx = [0]

        spinner_label = tk.Label(
            content,
            text='|',
            font=('Consolas', 14),
            fg=colors.get('primary', '#009999'),
            bg=colors['background']
        )
        spinner_label.pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(
            content,
            text="Signing in to Power BI...\n\nA browser window should open.",
            font=("Segoe UI", 10),
            bg=colors['background'],
            fg=colors['text_primary'],
            justify='left'
        ).pack(side=tk.LEFT, anchor='w')

        # Animation
        animating = [True]

        def animate():
            if animating[0] and dialog.winfo_exists():
                spinner_idx[0] = (spinner_idx[0] + 1) % 4
                spinner_label.config(text=spinner_chars[spinner_idx[0]])
                dialog.after(150, animate)

        # Authentication thread
        def auth_thread():
            success, message = self.cloud_browser.authenticate()
            auth_result['success'] = success
            auth_result['message'] = message
            # Schedule dialog close on main thread
            if dialog.winfo_exists():
                dialog.after(0, dialog.destroy)

        # Center and show dialog
        dialog.update_idletasks()
        parent = self.frame.winfo_toplevel()
        x = parent.winfo_x() + (parent.winfo_width() - 350) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 120) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.deiconify()
        dialog.grab_set()

        # Start animation and auth
        animate()
        threading.Thread(target=auth_thread, daemon=True).start()

        # Wait for dialog to close
        dialog.wait_window()
        animating[0] = False

        # Show error if authentication failed
        if not auth_result['success']:
            self._show_themed_info_dialog(
                "Authentication Failed",
                f"Could not sign in to Power BI.\n\n{auth_result['message']}"
            )

        return auth_result['success']

    def _show_thin_report_help_message(self, report_name: str, cloud_server: str = '',
                                        cloud_database: str = '', error: str = ''):
        """
        Show a dialog for connecting thin reports to local models.

        Args:
            report_name: Name of the thin report
            cloud_server: Cloud server URL (if known)
            cloud_database: Cloud dataset name (if known)
            error: Optional error message from connection attempt
        """
        self._log_message("")
        self._log_message("=" * 50)
        self._log_message("THIN REPORT DETECTED")
        self._log_message("=" * 50)
        self._log_message("")
        self._log_message(f"'{report_name}' is a thin report (live-connected report).")

        if cloud_database:
            self._log_message(f"Connected to: {cloud_database} (Cloud)")

        # Find the thin report model in cache to get file_path and process_id
        thin_report_model = None
        for model in self._dropdown_models_cache:
            if model.is_thin_report and report_name in model.display_name:
                thin_report_model = model
                break

        thin_report_file_path = None
        thin_report_process_id = None
        if thin_report_model:
            thin_report_file_path = thin_report_model.file_path
            thin_report_process_id = thin_report_model.process_id
            if thin_report_file_path:
                self._log_message(f"File: {thin_report_file_path}")

        # Use cached models from dropdown (already discovered)
        # Filter to non-thin-reports only
        local_models = [m for m in self._dropdown_models_cache if not m.is_thin_report]
        self._log_message(f"Found {len(local_models)} local model(s) in cache")

        # Find matching models by cloud database name
        matching_models = []
        if cloud_database and local_models:
            cloud_name_lower = cloud_database.lower()
            for model in local_models:
                # Match if cloud database name appears in the display name
                if cloud_name_lower in model.display_name.lower():
                    matching_models.append(model)

            if matching_models:
                self._log_message(f"Found {len(matching_models)} matching model(s):")
                for m in matching_models:
                    self._log_message(f"  - {m.display_name} ({m.server})")

        self._log_message("")
        self._log_message("Opening local model selector...")

        # Callback for when user selects a local model to open
        def on_select_local(model):
            """Handle selection of a local model to open in Hot-Swap"""
            self._log_message(f"Switching to local model: {model.display_name}")
            # Find and select this model in the dropdown
            for i, combo_text in enumerate(self.model_combo['values']):
                if model.display_name in combo_text:
                    self.model_combo.current(i)
                    self._on_connect()
                    break

        # Show the thin report swap dialog
        dialog = ThinReportSwapDialog(
            parent=self.frame,
            report_name=report_name,
            cloud_server=cloud_server,
            cloud_database=cloud_database,
            local_models=local_models,
            matching_models=matching_models,
            error=error,
            on_select_local=on_select_local,
            thin_report_file_path=thin_report_file_path,
            thin_report_process_id=thin_report_process_id
        )

        # Log result
        if dialog.result:
            self._log_message(f"Selected: {dialog.result.display_name}")
        else:
            self._log_message("Dialog closed without selection")

    def _quick_disconnect(self):
        """Quick disconnect for switching models - clears state without full UI reset.

        Used when user selects a different model while already connected,
        allowing seamless model switching without manual disconnect.
        """
        # Stop health checker
        if self.health_checker:
            self.health_checker.stop()
            self.health_checker = None

        # Disconnect from current model
        try:
            self.connector.disconnect()
        except:
            pass

        # Clear state but don't update UI status (we're about to reconnect)
        self.model_info = None
        self.mappings = []
        self._health_statuses.clear()
        self._thin_report_context = None
        self._current_model_file_path = None
        self.mapping_tree.delete(*self.mapping_tree.get_children())

        # Clear connection details
        self._clear_connection_details()

        # Reset logic components
        self.detector = None
        self.swapper = None
        self.matcher = None
        self.cloud_browser = None
        self.schema_validator = None

    def _on_disconnect(self):
        """Disconnect from the current model"""
        # Stop health checker
        if self.health_checker:
            self.health_checker.stop()
            self.health_checker = None

        try:
            self.connector.disconnect()
        except:
            pass

        self._update_status_display('error', "Not Connected")
        self.disconnect_btn.set_enabled(False)
        # Only enable connect if a model is selected
        self.connect_btn.set_enabled(bool(self.selected_model.get()))

        # Clear state
        self.model_info = None
        self.mappings = []
        self._health_statuses.clear()
        self._thin_report_context = None  # Clear thin report context
        self._current_model_file_path = None  # Clear file path for preset hashing
        self.mapping_tree.delete(*self.mapping_tree.get_children())
        self.mapping_info.set("Connect to a model to see swappable connections")

        # Clear Connection Details panel
        self._clear_connection_details()

        self._set_initial_state()

        # Refresh preset table - MODEL scope will show no presets when disconnected
        self._refresh_preset_table()

        self._log_message("Disconnected")

    def _detect_connections(self):
        """Detect swappable connections in the connected model"""
        if not self.detector:
            self._hide_progress()
            return

        self._update_progress_message("⏳ Analyzing data sources...")
        self._log_message("⏳ Detecting data source connections...")

        try:
            self.model_info = self.detector.detect_connections()

            # Set friendly name on local connections from the selected model's display name
            selected_name = self.selected_model.get()
            if selected_name and self.model_info:
                for conn in self.model_info.connections:
                    if not conn.is_cloud and not conn.dataset_name:
                        conn.dataset_name = selected_name

            # Update the Connection Details panel with model info
            self._update_connection_details()

            if self.model_info.swappable_count > 0:
                self.mapping_info.set(
                    f"Found {self.model_info.swappable_count} swappable connection(s) "
                    f"({self.model_info.connection_type.value} model)"
                )

                # Enable Last Config button if a saved config exists for this model
                model_hash = self._get_model_hash()
                if model_hash and self.preset_manager.has_last_config(model_hash):
                    self.last_config_btn.set_enabled(True)

                # Enable Rollback button if swap history exists from previous sessions
                if self._swap_history:
                    self.rollback_btn.set_enabled(True)

                # Auto-match if enabled
                if self.auto_match_enabled.get():
                    self._auto_match_connections()
                else:
                    self._populate_mapping_table([
                        ConnectionMapping(source=c) for c in self.model_info.connections if c.is_swappable
                    ])
                    self._hide_progress()

                self._log_message(f"✅ Detected {self.model_info.swappable_count} swappable connection(s)")

                # Check for model-specific presets and auto-switch to MODEL scope if found
                self._load_model_presets_on_connect()
            else:
                self._hide_progress()
                self.mapping_info.set("No swappable connections found (Import mode only?)")
                self._log_message(f"⚠️ No swappable connections found in this model (total datasources: {self.model_info.total_datasources})")

        except Exception as e:
            self._hide_progress()
            self._log_message(f"❌ Error detecting connections: {e}")
            import traceback
            self.logger.error(traceback.format_exc())

    def _auto_match_connections(self):
        """Auto-match connections to local models"""
        if not self.matcher or not self.model_info:
            self._hide_progress()
            return

        self._update_progress_message("⏳ Auto-matching local models...")
        self._log_message("⏳ Auto-matching local models...")

        def match_thread():
            try:
                swappable = [c for c in self.model_info.connections if c.is_swappable]
                mappings = self.matcher.suggest_matches(swappable)
                self.frame.after(0, lambda: self._populate_mapping_table(mappings))
                self.frame.after(0, lambda: self._hide_progress())

                matched = sum(1 for m in mappings if m.auto_matched)
                self.frame.after(0, lambda: self._log_message(
                    f"✅ Auto-matched {matched}/{len(mappings)} connection(s)"
                ))
            except Exception as e:
                self.frame.after(0, lambda: self._hide_progress())
                self.frame.after(0, lambda: self._log_message(f"❌ Auto-match error: {e}"))

        threading.Thread(target=match_thread, daemon=True).start()

    def _populate_mapping_table(self, mappings: List[ConnectionMapping]):
        """Populate the mapping treeview"""
        self.mappings = mappings
        self.mapping_tree.delete(*self.mapping_tree.get_children())

        for i, mapping in enumerate(mappings):
            source_name = mapping.source.display_name
            target_name = mapping.target.display_name if mapping.target else "(not set)"

            # Determine target type (Local/Cloud)
            if mapping.target:
                target_type = "Cloud" if mapping.target.target_type == "cloud" else "Local"
            else:
                target_type = "--"

            status = mapping.status.value

            # Status indicator - use clear text/symbols
            status_emoji = {
                'pending': '--',      # Not configured yet
                'matched': '?',       # Auto-matched, needs confirmation
                'ready': 'Ready',     # Configured and validated
                'swapping': '...',    # Swap in progress
                'success': 'Done',    # Completed
                'error': 'ERR',       # Failed
            }.get(status, '--')

            self.mapping_tree.insert(
                "",
                "end",
                iid=str(i),
                values=(source_name, target_name, target_type, status_emoji)
            )

        # Enable buttons based on mapping state
        if mappings:
            # Enable save mapping button if any mappings have targets
            has_targets = any(m.target for m in mappings)
            if has_targets and hasattr(self, 'save_mapping_btn'):
                self.save_mapping_btn.set_enabled(True)

            # Update swap button based on selection (requires user to select rows)
            self._on_mapping_selected()

        # Refresh local models list when mappings are populated
        self._refresh_local_models()

        # Start health monitoring if any mappings have targets
        has_targets = any(m.target for m in mappings)
        if has_targets and self.health_checker:
            self._start_health_monitoring()

        # Auto-save Last Config on first connect (if not already saved)
        self._save_last_config_if_new()

    def _save_last_config_if_new(self):
        """Save the current configuration as Last Config if not already saved for this model."""
        model_hash = self._get_model_hash()
        if not model_hash:
            return

        # Only save if no last config exists for this model
        if not self.preset_manager.has_last_config(model_hash):
            model_name = self._get_model_display_name() or "Unknown Model"

            # Extract friendly name and workspace from source connection (if cloud)
            friendly_name = None
            workspace_name = None
            if self.mappings and self.mappings[0].source:
                source = self.mappings[0].source
                friendly_name = source.dataset_name or source.database
                workspace_name = source.workspace_name

            if self.preset_manager.save_last_config(
                model_hash, self.mappings, model_name,
                friendly_name=friendly_name, workspace_name=workspace_name
            ):
                self._log_message(f"Saved starting configuration for {model_name}")
                # Refresh preset table to show Last Config entry
                self._refresh_preset_table()

    # =========================================================================
    # Swap Operations
    # =========================================================================

    def _on_mapping_double_click(self, event):
        """Handle double-click on mapping row"""
        selection = self.mapping_tree.selection()
        if selection:
            self._on_select_target()

    def _on_select_target(self):
        """Open target selection dialog for selected mapping"""
        selection = self.mapping_tree.selection()
        if not selection:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Select Mapping", "Please select a connection to configure.")
            return

        idx = int(selection[0])
        if idx >= len(self.mappings):
            return

        mapping = self.mappings[idx]

        # Show target selection dialog
        self._show_target_selector(mapping, idx)

    def _show_target_selector(self, mapping: ConnectionMapping, idx: int):
        """Show dialog to select target (local or cloud)"""
        # Ask user which type of target they want
        target_type = self._ask_target_type()
        if not target_type:
            return

        if target_type == "local":
            from tools.connection_hotswap.ui.dialogs.local_selector_dialog import LocalSelectorDialog
            # Use cached models from dropdown (filter out thin reports and self)
            # Also exclude the thin report's own local server if applicable
            thin_report_server = None
            if self._thin_report_context:
                thin_report_server = self._thin_report_context.get('local_server')

            cached_local_models = [
                m for m in self._dropdown_models_cache
                if not m.is_thin_report and (not thin_report_server or m.server != thin_report_server)
            ]
            dialog = LocalSelectorDialog(
                self.frame.winfo_toplevel(),
                self.matcher,
                suggested_name=mapping.source.dataset_name or mapping.source.database,
                cached_models=cached_local_models if cached_local_models else None
            )
            result = dialog.result
        else:  # cloud
            from core.cloud import CloudBrowserDialog
            from tools.connection_hotswap.models import CloudConnectionType
            # Get current model status for dialog display
            current_status = self.connection_status.get()
            model_status = current_status if current_status and current_status != "Not Connected" else None

            # Build source context for dialog title (e.g., "Budget Report (Local)")
            source_name = mapping.source.dataset_name or mapping.source.database or mapping.source.name
            source_type = "Cloud" if mapping.source.is_cloud else "Local"
            source_context = f"{source_name} ({source_type})"

            # Always default to Semantic Model connector when manually selecting cloud target
            dialog = CloudBrowserDialog(
                self.frame.winfo_toplevel(),
                self.cloud_browser,
                default_connection_type=CloudConnectionType.PBI_SEMANTIC_MODEL,
                model_status=model_status,
                on_auth_change=self._update_cloud_auth_button_state,
                source_context=source_context
            )
            dialog.wait_window()
            result = dialog.result

        if result:
            mapping.target = result
            mapping.status = SwapStatus.READY
            self._update_mapping_row(idx, mapping)
            self._log_message(f"Target set: {mapping.source.name} → {result.display_name}")

    def _ask_target_type(self) -> str:
        """Ask user whether to select local or cloud target"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        parent = self.frame.winfo_toplevel()

        dialog_width = 380
        dialog_height = 120  # Compact height

        dialog = tk.Toplevel(parent)
        dialog.title("Select Target Type")
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

        # Set dialog icon (AE favicon)
        try:
            # Go up from ui -> connection_hotswap -> tools -> src, then to assets
            base_dir = Path(__file__).parent.parent.parent.parent
            icon_path = base_dir / 'assets' / 'favicon.ico'
            if icon_path.exists():
                dialog.iconbitmap(str(icon_path))
        except Exception:
            pass

        # Set dark title bar on Windows for theme matching
        try:
            from ctypes import windll, byref, sizeof, c_int
            HWND = windll.user32.GetParent(dialog.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = c_int(1 if is_dark else 0)
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_USE_IMMERSIVE_DARK_MODE, byref(value), sizeof(value))
        except Exception:
            pass

        # Center on parent
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - dialog_width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - dialog_height) // 2
        dialog.geometry(f"+{x}+{y}")

        result = {"value": None}

        # Content (no internal header - window title bar is sufficient)
        frame = tk.Frame(dialog, bg=colors['background'], padx=30, pady=18)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            frame,
            text="Choose a target type for this connection:",
            font=("Segoe UI", 10),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(pady=(0, 15))

        btn_frame = tk.Frame(frame, bg=colors['background'])
        btn_frame.pack()

        canvas_bg = colors['background']

        def select_local():
            result["value"] = "local"
            dialog.destroy()

        def select_cloud():
            result["value"] = "cloud"
            dialog.destroy()

        # Use folder icon for local model - auto-sized for consistent padding
        local_icon = self._button_icons.get('folder')
        local_btn = RoundedButton(
            btn_frame, text="LOCAL MODEL",
            command=select_local,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=36, radius=8, font=('Segoe UI', 10, 'bold'),
            icon=local_icon, canvas_bg=canvas_bg
        )
        local_btn.pack(side=tk.LEFT, padx=(0, 20))

        cloud_icon = self._button_icons.get('cloud-computing')
        cloud_btn = RoundedButton(
            btn_frame, text="CLOUD MODEL",
            command=select_cloud,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=36, radius=8, font=('Segoe UI', 10, 'bold'),
            icon=cloud_icon, canvas_bg=canvas_bg
        )
        cloud_btn.pack(side=tk.LEFT)

        dialog.wait_window()
        return result["value"]

    def _update_mapping_row(self, idx: int, mapping: ConnectionMapping):
        """Update a single row in the mapping table and diagram"""
        target_name = mapping.target.display_name if mapping.target else "(not set)"

        # Determine target type (Local/Cloud)
        if mapping.target:
            target_type = "Cloud" if mapping.target.target_type == "cloud" else "Local"
        else:
            target_type = "--"

        status = mapping.status.value
        status_emoji = {
            'pending': '--',      # Not configured yet
            'matched': '?',       # Auto-matched, needs confirmation
            'ready': 'Ready',     # Configured and validated
            'swapping': '...',    # Swap in progress
            'success': 'Done',    # Completed
            'error': 'ERR',       # Failed
        }.get(status, '--')

        self.mapping_tree.item(
            str(idx),
            values=(mapping.source.display_name, target_name, target_type, status_emoji)
        )

        # Update diagram if in diagram view
        self._update_diagram()

        # Update button states based on current mappings
        has_targets = any(m.target for m in self.mappings)
        if hasattr(self, 'save_mapping_btn'):
            self.save_mapping_btn.set_enabled(has_targets)
        # Swap button requires selection - defer to selection handler
        self._on_mapping_selected()

    def _on_swap_selected(self):
        """Swap the selected connection(s) - supports batch swap for multi-selection"""
        selection = self.mapping_tree.selection()
        if not selection:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Select Mapping", "Please select a connection to swap.")
            return

        # Gather all selected mappings that have targets
        ready_to_swap = []
        for item_id in selection:
            idx = int(item_id)
            if idx >= len(self.mappings):
                continue
            mapping = self.mappings[idx]
            if mapping.target:
                ready_to_swap.append((idx, mapping))

        if not ready_to_swap:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Targets", "Selected connection(s) have no target configured. Please set targets first.")
            return

        # If multiple selected, confirm batch swap
        if len(ready_to_swap) > 1:
            confirm = ThemedMessageBox.askyesno(
                self.frame.winfo_toplevel(),
                "Batch Swap",
                f"Swap {len(ready_to_swap)} connection(s)?"
            )
            if not confirm:
                return

            # Execute batch swap sequentially (to avoid XmlReader conflict)
            self._execute_batch_swap(ready_to_swap)
        else:
            # Single swap
            idx, mapping = ready_to_swap[0]
            self._execute_swap(mapping, idx)

    def _get_last_config_tooltip(self) -> str:
        """Get dynamic tooltip text for LAST CONFIG button based on current state"""
        model_hash = self._get_model_hash()
        if not model_hash:
            return "No saved target available\n(Connect to a model first)"

        if not self.preset_manager.has_last_config(model_hash):
            return "No saved target available\n(Starting config is saved on first connect)"

        # Get the last config details
        last_config = self.preset_manager.get_last_config(model_hash)
        if last_config:
            model_name = last_config.get('model_name', 'Unknown')
            mapping_count = len(last_config.get('mappings', []))
            saved_time = last_config.get('saved_at', '')
            # Format the saved time if present
            time_str = ""
            if saved_time:
                try:
                    from datetime import datetime
                    dt = datetime.fromisoformat(saved_time)
                    time_str = f"\nSaved: {dt.strftime('%m/%d/%Y %I:%M %p')}"
                except Exception:
                    pass
            return f"Set target to original connection\nModel: {model_name}\nConnections: {mapping_count}{time_str}"

        return "Set target to original connection"

    def _on_apply_last_config(self):
        """Apply the saved Last Config to the mapping table (preview mode)"""
        model_hash = self._get_model_hash()
        if not model_hash:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "Not Connected", "Please connect to a model first.")
            return

        if not self.preset_manager.has_last_config(model_hash):
            ThemedMessageBox.showinfo(
                self.frame.winfo_toplevel(),
                "No Saved Target",
                "No saved target configuration found for this model.\n"
                "The original connection is saved automatically when you first connect."
            )
            return

        # Apply the last config to current mappings (sets targets)
        applied_count = self.preset_manager.apply_last_config_to_mappings(model_hash, self.mappings)

        if applied_count > 0:
            # Check if any mappings have GUID databases that need resolution
            has_guids = any(
                mapping.target and self._is_guid_format(mapping.target.database or "")
                for mapping in self.mappings
            )

            if has_guids:
                # Resolve GUIDs asynchronously to prevent UI freeze
                self._show_progress("Resolving cloud connection")
                self._resolve_guid_display_names_async(applied_count)
            else:
                # No GUIDs to resolve - complete immediately
                self._complete_last_config_apply(applied_count)
        else:
            self._log_message("No mappings were updated from saved target")

    def _resolve_guid_display_names_async(self, applied_count: int):
        """Resolve GUID display names in background thread to prevent UI freeze."""
        def resolve_thread():
            try:
                # This calls _try_resolve_guid_to_name which may do cloud auth
                self._resolve_guid_display_names()
                # Callback to main thread when done
                self.frame.after(0, lambda: self._complete_last_config_apply(applied_count))
            except Exception as e:
                self.logger.error(f"GUID resolution error: {e}")
                # Still complete the apply even if resolution fails
                self.frame.after(0, lambda: self._complete_last_config_apply(applied_count))

        threading.Thread(target=resolve_thread, daemon=True).start()

    def _complete_last_config_apply(self, applied_count: int):
        """Complete the Last Config apply after GUID resolution (runs on main thread)."""
        self._hide_progress()

        # Refresh the mapping table and collect ready rows
        ready_rows = []
        for idx, mapping in enumerate(self.mappings):
            self._update_mapping_row(idx, mapping)
            if mapping.status == SwapStatus.READY:
                ready_rows.append(str(idx))

        self._log_message(f"Applied saved target to {applied_count} mapping(s)")

        # Auto-select all ready rows so user can immediately click SWAP SELECTED
        if ready_rows:
            self.mapping_tree.selection_set(ready_rows)
            self._on_mapping_selected()
            self._log_message("Ready to swap - click SWAP SELECTED to apply")

    def _on_swap_all(self, target_type: str):
        """Swap all connections to cloud or local"""
        ready_mappings = [(i, m) for i, m in enumerate(self.mappings) if m.is_ready]

        if not ready_mappings:
            ThemedMessageBox.showinfo(
                self.frame.winfo_toplevel(),
                "No Ready Mappings",
                f"No connections are ready to swap to {target_type}.\n"
                "Please configure targets first."
            )
            return

        confirm = ThemedMessageBox.askyesno(
            self.frame.winfo_toplevel(),
            "Confirm Swap",
            f"Swap {len(ready_mappings)} connection(s) to {target_type}?"
        )

        if confirm:
            # Execute batch swap sequentially (to avoid XmlReader conflict)
            self._execute_batch_swap(ready_mappings)

    def _execute_batch_swap(self, swaps_to_execute: List[tuple]):
        """Execute multiple swaps sequentially to avoid XmlReader conflict"""
        # For thin reports, route to single swap (they only have one connection)
        if self._thin_report_context:
            if swaps_to_execute:
                idx, mapping = swaps_to_execute[0]
                self._execute_thin_report_swap(mapping, idx)
            return

        if not self.swapper or not swaps_to_execute:
            return

        # Save origin state before swapping (so user can return via Last Config)
        self._save_origin_state()

        # Disable buttons during batch operation
        self.swap_selected_btn.set_enabled(False)
        self.last_config_btn.set_enabled(False)

        # Generate a unique run_id for this batch (groups all swaps together)
        self._current_batch_run_id = self._generate_run_id()

        # Mark all as swapping
        for idx, mapping in swaps_to_execute:
            mapping.status = SwapStatus.SWAPPING
            self._update_mapping_row(idx, mapping)

        total = len(swaps_to_execute)
        self._show_progress(f"⏳ Swapping {total} connection(s)...")
        self._log_message(f"⏳ Starting batch swap of {total} connection(s)...")

        def batch_swap_thread():
            results = []
            for i, (idx, mapping) in enumerate(swaps_to_execute, 1):
                self.frame.after(0, lambda n=mapping.source.name, c=i, t=total:
                    self._update_progress_message(f"⏳ Swapping {n} ({c}/{t})..."))
                result = self.swapper.swap_connection(mapping)
                results.append((idx, result))
                # Update UI after each swap completes
                self.frame.after(0, lambda r=result, ix=idx: self._handle_batch_swap_item(r, ix))

            # Final update after all complete
            self.frame.after(0, lambda: self._handle_batch_swap_complete(results))

        threading.Thread(target=batch_swap_thread, daemon=True).start()

    def _handle_batch_swap_item(self, result, idx: int):
        """Handle individual swap result during batch operation"""
        self._update_mapping_row(idx, result.mapping)
        self._log_message(f"{'✅' if result.success else '❌'} {result.message} ({result.elapsed_ms}ms)")

        # Track successful swap for rollback
        if result.success:
            self._last_swapped_mapping = result.mapping
            # Use run_id to group batch swaps together
            run_id = getattr(self, '_current_batch_run_id', '')
            history_entry = SwapHistoryEntry.from_mapping(
                result.mapping, run_id=run_id, model_file_path=self._current_model_file_path or ""
            )
            self._swap_history.insert(0, history_entry)
            if len(self._swap_history) > self._max_history_entries:
                self._swap_history = self._swap_history[:self._max_history_entries]

    def _handle_batch_swap_complete(self, results):
        """Handle completion of batch swap operation"""
        self._hide_progress()

        success_count = sum(1 for _, r in results if r.success)
        total = len(results)
        self._log_message(f"Batch swap complete: {success_count}/{total} succeeded")

        # Clear the current batch run_id
        self._current_batch_run_id = None

        # Persist swap history to file
        self._save_swap_history()

        # Re-enable buttons based on selection state
        if self.mappings:
            # Auto-select all ready mappings (ones with targets that can be swapped back)
            ready_items = [
                str(i) for i, m in enumerate(self.mappings)
                if m.status == SwapStatus.READY and m.target
            ]
            if ready_items:
                self.mapping_tree.selection_set(ready_items)

            self._on_mapping_selected()  # Check selection before enabling swap button
            if success_count > 0:
                self.rollback_btn.set_enabled(True)
            model_hash = self._get_model_hash()
            if model_hash and self.preset_manager.has_last_config(model_hash):
                self.last_config_btn.set_enabled(True)

    def _save_origin_state(self):
        """Save current state as Last Config before executing any swap.

        This allows users to return to their ORIGINAL starting state using Last Config.

        IMPORTANT: Only saves if no last_config exists yet, to preserve the original
        cloud connections. This prevents overwriting the starting config after swapping
        to local, which would lose the ability to swap back to cloud.
        """
        model_hash = self._get_model_hash()
        if not model_hash or not self.mappings:
            return

        # Only save if no last config exists - preserve the original starting state
        if self.preset_manager.has_last_config(model_hash):
            return

        model_name = self._get_model_display_name() or "Unknown Model"

        # Extract friendly name and workspace from source connection (if cloud)
        friendly_name = None
        workspace_name = None
        if self.mappings and self.mappings[0].source:
            source = self.mappings[0].source
            friendly_name = source.dataset_name or source.database
            workspace_name = source.workspace_name

        if self.preset_manager.save_last_config(
            model_hash, self.mappings, model_name,
            friendly_name=friendly_name, workspace_name=workspace_name
        ):
            self._log_message("Saved starting configuration as Last Config")

    def _execute_swap(self, mapping: ConnectionMapping, idx: int):
        """Execute a single swap"""
        # Check if this is a thin report swap
        if self._thin_report_context:
            self._execute_thin_report_swap(mapping, idx)
            return

        if not self.swapper:
            self._log_message("No swap handler available. Please reconnect to the model.")
            ThemedMessageBox.show(
                self.frame,
                title="Reconnect Required",
                message="Please reconnect to the model after the file reopens.\n\n"
                        "Use the REFRESH button or select the model from the dropdown.",
                msg_type="info",
                buttons=["OK"],
                custom_icon="hotswap.svg"
            )
            return

        # Save origin state before swapping (so user can return via Last Config)
        self._save_origin_state()

        mapping.status = SwapStatus.SWAPPING
        self._update_mapping_row(idx, mapping)
        self._show_progress(f"Swapping {mapping.source.name}...")
        self._log_message(f"Swapping {mapping.source.name}...")

        # Disable swap buttons during operation
        self.swap_selected_btn.set_enabled(False)
        self.last_config_btn.set_enabled(False)

        def swap_thread():
            self.frame.after(0, lambda: self._update_progress_message(f"Modifying connection..."))
            result = self.swapper.swap_connection(mapping)
            self.frame.after(0, lambda: self._handle_swap_result(result, idx))

        threading.Thread(target=swap_thread, daemon=True).start()

    def _handle_swap_result(self, result, idx: int):
        """Handle swap result"""
        self._hide_progress()
        self._update_mapping_row(idx, result.mapping)
        self._log_message(f"{'✅' if result.success else '❌'} {result.message} ({result.elapsed_ms}ms)")

        # Track successful swap for rollback
        if result.success:
            self._last_swapped_mapping = result.mapping
            self.rollback_btn.set_enabled(True)

            # Add to history with unique run_id for single swaps
            run_id = self._generate_run_id()
            history_entry = SwapHistoryEntry.from_mapping(
                result.mapping, run_id=run_id, model_file_path=self._current_model_file_path or ""
            )
            self._swap_history.insert(0, history_entry)  # Most recent first

            # Trim history to max entries
            if len(self._swap_history) > self._max_history_entries:
                self._swap_history = self._swap_history[:self._max_history_entries]

            # Persist to file
            self._save_swap_history()

        # Re-enable swap buttons based on selection
        if self.mappings:
            self._on_mapping_selected()  # Check selection before enabling swap button
            # Re-enable Last Config button if a saved config exists
            model_hash = self._get_model_hash()
            if model_hash and self.preset_manager.has_last_config(model_hash):
                self.last_config_btn.set_enabled(True)

    # =========================================================================
    # Thin Report Swap Execution
    # =========================================================================

    def _execute_thin_report_swap(self, mapping: ConnectionMapping, idx: int):
        """
        Execute a swap for a thin report by modifying the PBIX/PBIP file.

        For PBIP files: Modify directly (files not locked)
        For PBIX files: Show save/close confirmation, then modify
        """
        if not self._thin_report_context:
            self._log_message("Error: No thin report context available")
            return

        if not mapping.target:
            self._log_message("Error: No target selected for swap")
            return

        file_path = self._thin_report_context.get('file_path')
        if not file_path:
            self._log_message("Error: Cannot determine file path for thin report")
            ThemedMessageBox.showerror(
                self.frame.winfo_toplevel(),
                "File Path Unknown",
                "Could not determine the file path for this thin report.\n\n"
                "Please close and reopen Power BI Desktop, then try again."
            )
            return

        # Detect file type
        modifier = get_pbix_modifier()
        file_type = modifier.detect_file_type(file_path)

        if not file_type:
            self._log_message(f"Error: Unknown file type for {file_path}")
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Unknown File Type", f"Cannot determine file type for:\n{file_path}")
            return

        # Save origin state before swapping
        self._save_origin_state()

        target_server = mapping.target.server
        target_database = mapping.target.database

        if file_type == 'pbip':
            # PBIP files are not locked, but still need close/reopen for changes to take effect
            self._show_pbip_swap_confirmation(mapping, idx, file_path, target_server, target_database)
        else:
            # PBIX files are locked - need save/close confirmation
            self._show_pbix_swap_confirmation(mapping, idx, file_path, target_server, target_database)

    def _show_pbip_swap_confirmation(self, mapping: ConnectionMapping, idx: int,
                                      file_path: str, target_server: str, target_database: str):
        """Show inline confirmation for PBIP swap (requires close/reopen for changes to take effect)."""
        import os

        file_name = os.path.basename(file_path)

        # Get current backup setting as default for checkbox
        backup_default = self.preset_manager.get_backup_enabled()

        # Show confirmation dialog with auto-reconnect and backup options
        result = ThemedMessageBox.show(
            self.frame,
            title="PBIP Requires Close/Reopen",
            message=f"File: {file_name}\n\n"
                    f"PBIP files can be modified while open, but Power BI Desktop\n"
                    f"must be closed and reopened for changes to take effect.\n\n"
                    f"Steps:\n"
                    f"1. Save your work (optional)\n"
                    f"2. Close Power BI Desktop\n"
                    f"3. Modify the connection\n"
                    f"4. Reopen the file\n\n"
                    f"Reconnection required for further swaps.",
            msg_type="warning",
            buttons=["Save & Swap", "Swap Without Saving", "Cancel"],
            custom_icon="hotswap.svg",
            checkbox_text="Auto-reconnect after swap",
            checkbox_default=True,
            checkbox2_text="Create backup before swap",
            checkbox2_default=backup_default
        )

        # Result is (button_text, auto_reconnect, create_backup) tuple
        button_result, auto_reconnect, create_backup = result

        # Save backup preference for future use
        if create_backup != backup_default:
            self.preset_manager.set_backup_enabled(create_backup)

        if button_result == "Cancel" or button_result is None:
            self._log_message("Swap cancelled by user")
            return
        elif button_result == "Save & Swap":
            self._execute_pbip_swap(mapping, idx, file_path, target_server, target_database,
                                    save_first=True, auto_reconnect=auto_reconnect, create_backup=create_backup)
        else:
            self._execute_pbip_swap(mapping, idx, file_path, target_server, target_database,
                                    save_first=False, auto_reconnect=auto_reconnect, create_backup=create_backup)

    def _execute_pbip_swap(self, mapping: ConnectionMapping, idx: int,
                           file_path: str, target_server: str, target_database: str,
                           save_first: bool = True, auto_reconnect: bool = True, create_backup: bool = False):
        """Execute swap for PBIP file (not locked while open, but requires close/reopen)."""
        mapping.status = SwapStatus.SWAPPING
        self._update_mapping_row(idx, mapping)
        self._show_progress(f"Swapping {mapping.source.name}...")
        self._log_message(f"Starting PBIP swap process (save_first={save_first}, auto_reconnect={auto_reconnect}, backup={create_backup})...")

        self.swap_selected_btn.set_enabled(False)
        self.last_config_btn.set_enabled(False)

        process_id = self._thin_report_context.get('process_id')
        # Get dataset_id if this is a cloud target
        dataset_id = getattr(mapping.target, 'dataset_id', None) if mapping.target else None
        # Get perspective name if specified
        perspective_name = getattr(mapping.target, 'perspective_name', None) if mapping.target else None
        # Get workspace name for cloud targets with perspectives (needed for XMLA URL)
        workspace_name = getattr(mapping.target, 'workspace_name', None) if mapping.target else None
        # Get cloud connection type (pbi_semantic_model vs aas_xmla)
        cloud_conn_type = None
        if mapping.target:
            conn_type_enum = getattr(mapping.target, 'cloud_connection_type', None)
            if conn_type_enum:
                cloud_conn_type = conn_type_enum.value  # Get string value from enum
        # Get source friendly name from thin report context (for GUID resolution when swapping back)
        source_friendly_name = self._thin_report_context.get('display_name') if self._thin_report_context else None

        def swap_thread():
            import time
            start_time = time.time()

            controller = get_process_controller()
            modifier = get_pbix_modifier()

            success = False
            message = ""

            try:
                # Step 1: Save if requested
                if save_first and process_id:
                    self.frame.after(0, lambda: self._update_progress_message("Saving file"))
                    self._log_message("Sending save command to Power BI Desktop...")
                    save_result = controller.save_file(process_id, timeout=5.0)
                    if not save_result.success:
                        self._log_message(f"Warning: Save may not have completed: {save_result.message}")

                # Step 2: Close Power BI Desktop
                if process_id:
                    self.frame.after(0, lambda: self._update_progress_message("Closing Power BI Desktop"))
                    self._log_message("Closing Power BI Desktop...")
                    close_result = controller.close_gracefully(process_id, timeout=15.0)
                    if not close_result.success:
                        # Try force close
                        self._log_message("Graceful close failed, trying force close...")
                        close_result = controller.force_close(process_id)
                        if not close_result.success:
                            raise Exception(f"Could not close Power BI Desktop: {close_result.message}")

                # Step 3: Modify the file (PBIP files are not locked)
                self.frame.after(0, lambda: self._update_progress_message("Modifying connection"))
                self._log_message(f"Modifying PBIP to connect to {target_server}...")
                # Determine if we should use cached cloud connection
                is_cloud = modifier._is_cloud_server(target_server)
                swap_result = modifier.swap_connection(file_path, target_server, target_database,
                                                       create_backup=create_backup, dataset_id=dataset_id,
                                                       source_friendly_name=source_friendly_name,
                                                       perspective_name=perspective_name,
                                                       workspace_name=workspace_name,
                                                       cloud_connection_type=cloud_conn_type,
                                                       use_cached_cloud=is_cloud)
                if not swap_result.success:
                    raise Exception(f"File modification failed: {swap_result.message}")

                # Step 4: Reopen the file if auto-reconnect is enabled
                if auto_reconnect:
                    self.frame.after(0, lambda: self._update_progress_message("Reopening file"))
                    self._log_message("Reopening file in Power BI Desktop...")
                    reopen_result = controller.reopen_file(file_path)
                    if not reopen_result.success:
                        # Non-fatal - file was modified successfully
                        self._log_message(f"Warning: Could not auto-reopen file: {reopen_result.message}")

                success = True
                message = f"Connection swapped to {target_server}"
                if swap_result.backup_path:
                    message += f" (backup created)"

                self._log_message(f"PBIP swap thread completed successfully")

            except Exception as e:
                success = False
                message = str(e)
                self._log_message(f"Error during PBIP swap: {e}")

            # Wrap completion handling in try-except to ensure UI always updates
            try:
                elapsed_ms = int((time.time() - start_time) * 1000)

                # Update mapping status
                if success:
                    mapping.status = SwapStatus.SUCCESS
                    if mapping.target:
                        mapping.target.server = target_server
                        mapping.target.database = target_database
                else:
                    mapping.status = SwapStatus.ERROR

                # Capture values for lambda closure
                _success = success
                _message = message
                _elapsed_ms = elapsed_ms
                _auto_reconnect = auto_reconnect

                self._log_message(f"Scheduling UI completion callback...")

                def safe_completion_callback():
                    try:
                        self._handle_thin_report_swap_result(
                            _success, _message, mapping, idx, _elapsed_ms, file_path,
                            is_pbip=True, auto_reconnect=_auto_reconnect
                        )
                    except Exception as cb_error:
                        self.logger.error(f"Error in completion callback: {cb_error}")
                        try:
                            self._hide_progress()
                            self.swap_selected_btn.set_enabled(True)
                            self._log_message(f"Error updating UI: {cb_error}")
                        except:
                            pass

                self.frame.after(0, safe_completion_callback)

            except Exception as final_error:
                self.logger.error(f"Error in swap completion: {final_error}")
                err_msg = str(final_error)
                try:
                    self.frame.after(0, lambda: self._hide_progress())
                    self.frame.after(0, lambda: self.swap_selected_btn.set_enabled(True))
                    self.frame.after(0, lambda: self._log_message(f"Error: {err_msg}"))
                except:
                    pass

        threading.Thread(target=swap_thread, daemon=True).start()

    def _show_pbix_swap_confirmation(self, mapping: ConnectionMapping, idx: int,
                                      file_path: str, target_server: str, target_database: str):
        """Show inline confirmation for PBIX swap (requires close/reopen)."""
        import os

        file_name = os.path.basename(file_path)

        # Get current backup setting as default for checkbox
        backup_default = self.preset_manager.get_backup_enabled()

        # Show confirmation dialog with auto-reconnect and backup options
        result = ThemedMessageBox.show(
            self.frame,
            title="PBIX Requires Close/Reopen",
            message=f"File: {file_name}\n\n"
                    f"PBIX files are locked while open.\n\n"
                    f"Steps:\n"
                    f"1. Save your work (optional)\n"
                    f"2. Close Power BI Desktop\n"
                    f"3. Modify the connection\n"
                    f"4. Reopen the file\n\n"
                    f"Reconnection required for further swaps.",
            msg_type="warning",
            buttons=["Save & Swap", "Swap Without Saving", "Cancel"],
            custom_icon="hotswap.svg",
            checkbox_text="Auto-reconnect after swap",
            checkbox_default=True,
            checkbox2_text="Create backup before swap",
            checkbox2_default=backup_default
        )

        # Result is (button_text, auto_reconnect, create_backup) tuple
        button_result, auto_reconnect, create_backup = result

        # Save backup preference for future use
        if create_backup != backup_default:
            self.preset_manager.set_backup_enabled(create_backup)

        if button_result == "Cancel" or button_result is None:
            self._log_message("Swap cancelled by user")
            return
        elif button_result == "Save & Swap":
            self._execute_pbix_swap(mapping, idx, file_path, target_server, target_database,
                                    save_first=True, auto_reconnect=auto_reconnect, create_backup=create_backup)
        else:
            self._execute_pbix_swap(mapping, idx, file_path, target_server, target_database,
                                    save_first=False, auto_reconnect=auto_reconnect, create_backup=create_backup)

    def _execute_pbix_swap(self, mapping: ConnectionMapping, idx: int, file_path: str,
                           target_server: str, target_database: str, save_first: bool = True,
                           auto_reconnect: bool = True, create_backup: bool = False):
        """Execute swap for PBIX file (locked while open, requires close/reopen)."""
        mapping.status = SwapStatus.SWAPPING
        self._update_mapping_row(idx, mapping)
        self._show_progress(f"Swapping {mapping.source.name}...")
        self._log_message(f"Starting PBIX swap process (save_first={save_first}, auto_reconnect={auto_reconnect}, backup={create_backup})...")

        self.swap_selected_btn.set_enabled(False)
        self.last_config_btn.set_enabled(False)

        process_id = self._thin_report_context.get('process_id')
        # Get dataset_id if this is a cloud target
        dataset_id = getattr(mapping.target, 'dataset_id', None) if mapping.target else None
        # Get perspective name if specified
        perspective_name = getattr(mapping.target, 'perspective_name', None) if mapping.target else None
        # Get workspace name for cloud targets with perspectives (needed for XMLA URL)
        workspace_name = getattr(mapping.target, 'workspace_name', None) if mapping.target else None
        # Get cloud connection type (pbi_semantic_model vs aas_xmla)
        cloud_conn_type = None
        if mapping.target:
            conn_type_enum = getattr(mapping.target, 'cloud_connection_type', None)
            if conn_type_enum:
                cloud_conn_type = conn_type_enum.value  # Get string value from enum
        # Get source friendly name from thin report context (for GUID resolution when swapping back)
        source_friendly_name = self._thin_report_context.get('display_name') if self._thin_report_context else None

        def swap_thread():
            import time
            start_time = time.time()

            controller = get_process_controller()
            modifier = get_pbix_modifier()

            success = False
            message = ""

            try:
                # Step 1: Save if requested
                if save_first and process_id:
                    self.frame.after(0, lambda: self._update_progress_message("Saving file..."))
                    self._log_message("Sending save command to Power BI Desktop...")
                    save_result = controller.save_file(process_id, timeout=5.0)
                    if not save_result.success:
                        self._log_message(f"Warning: Save may not have completed: {save_result.message}")

                # Step 2: Close Power BI Desktop
                if process_id:
                    self.frame.after(0, lambda: self._update_progress_message("Closing Power BI Desktop..."))
                    self._log_message("Closing Power BI Desktop...")
                    close_result = controller.close_gracefully(process_id, timeout=15.0)
                    if not close_result.success:
                        # Try force close
                        self._log_message("Graceful close failed, trying force close...")
                        close_result = controller.force_close(process_id)
                        if not close_result.success:
                            raise Exception(f"Could not close Power BI Desktop: {close_result.message}")

                # Step 3: Wait for file to be unlocked
                self.frame.after(0, lambda: self._update_progress_message("Waiting for file unlock..."))
                self._log_message("Waiting for file to be unlocked...")
                unlock_result = controller.wait_for_file_unlock(file_path, timeout=30.0)
                if not unlock_result.success:
                    raise Exception(f"File did not unlock: {unlock_result.message}")

                # Step 4: Modify the file
                self.frame.after(0, lambda: self._update_progress_message("Modifying connection..."))
                self._log_message(f"Modifying PBIX to connect to {target_server}...")
                # Determine if we should use cached cloud connection
                # use_cached_cloud=True: Restore original connection (preserves PbiServiceModelId, etc.)
                # If no cache exists, swap_connection falls back to building a new connection
                is_cloud = modifier._is_cloud_server(target_server)
                swap_result = modifier.swap_connection(file_path, target_server, target_database,
                                                       create_backup=create_backup, dataset_id=dataset_id,
                                                       source_friendly_name=source_friendly_name,
                                                       perspective_name=perspective_name,
                                                       workspace_name=workspace_name,
                                                       cloud_connection_type=cloud_conn_type,
                                                       use_cached_cloud=is_cloud)
                if not swap_result.success:
                    raise Exception(f"File modification failed: {swap_result.message}")

                # Step 5: Reopen the file
                self.frame.after(0, lambda: self._update_progress_message("Reopening file..."))
                self._log_message("Reopening file in Power BI Desktop...")
                reopen_result = controller.reopen_file(file_path)
                if not reopen_result.success:
                    # Non-fatal - file was modified successfully
                    self._log_message(f"Warning: Could not auto-reopen file: {reopen_result.message}")

                success = True
                message = f"Connection swapped to {target_server}"
                if swap_result.backup_path:
                    message += f" (backup created)"

                self._log_message(f"PBIX swap thread completed successfully")

            except Exception as e:
                success = False
                message = str(e)
                self._log_message(f"Error during PBIX swap: {e}")

            # Wrap completion handling in try-except to ensure UI always updates
            try:
                elapsed_ms = int((time.time() - start_time) * 1000)

                # Update mapping status
                if success:
                    mapping.status = SwapStatus.SUCCESS
                    if mapping.target:
                        mapping.target.server = target_server
                        mapping.target.database = target_database
                else:
                    mapping.status = SwapStatus.ERROR

                # Capture values for lambda closure
                _success = success
                _message = message
                _elapsed_ms = elapsed_ms
                _auto_reconnect = auto_reconnect

                self._log_message(f"Scheduling UI completion callback...")

                def safe_completion_callback():
                    try:
                        self._handle_thin_report_swap_result(
                            _success, _message, mapping, idx, _elapsed_ms, file_path,
                            is_pbip=False, auto_reconnect=_auto_reconnect
                        )
                    except Exception as cb_error:
                        self.logger.error(f"Error in completion callback: {cb_error}")
                        # Ensure progress is hidden and buttons re-enabled even on error
                        try:
                            self._hide_progress()
                            self.swap_selected_btn.set_enabled(True)
                            self._log_message(f"Error updating UI: {cb_error}")
                        except:
                            pass

                self.frame.after(0, safe_completion_callback)

            except Exception as final_error:
                self.logger.error(f"Error in swap completion: {final_error}")
                # Emergency cleanup - ensure UI is restored
                # Use default argument to capture final_error value
                err_msg = str(final_error)
                try:
                    self.frame.after(0, lambda: self._hide_progress())
                    self.frame.after(0, lambda: self.swap_selected_btn.set_enabled(True))
                    self.frame.after(0, lambda msg=err_msg: self._log_message(f"Swap error: {msg}"))
                except:
                    pass

        threading.Thread(target=swap_thread, daemon=True).start()

    def _handle_thin_report_swap_result(self, success: bool, message: str,
                                         mapping: ConnectionMapping, idx: int,
                                         elapsed_ms: int, file_path: str,
                                         is_pbip: bool = False, auto_reconnect: bool = True):
        """Handle result of thin report swap operation."""
        self._log_message(f"Processing swap result: success={success}")

        # ALWAYS hide progress and re-enable buttons, even if there's an error later
        self._hide_progress()
        self.swap_selected_btn.set_enabled(True)  # Re-enable explicitly
        self._update_mapping_row(idx, mapping)

        status_icon = "Complete" if success else "Error"
        self._log_message(f"{status_icon}: {message} ({elapsed_ms}ms)")

        if success:
            self._last_swapped_mapping = mapping
            self.rollback_btn.set_enabled(True)

            # Capture original values for history before modifying source
            if not mapping.original_server:
                mapping.original_server = mapping.source.server
            if not mapping.original_database:
                mapping.original_database = mapping.source.database

            # Update the source connection to reflect the new state (it's now local)
            if mapping.target:
                mapping.source.server = mapping.target.server
                mapping.source.database = mapping.target.database
                # Detect if new server is local or cloud
                server_lower = mapping.target.server.lower()
                mapping.source.is_cloud = (
                    'powerbi://' in server_lower or
                    'asazure://' in server_lower
                )
                # Update cloud-specific fields based on new target type
                if not mapping.source.is_cloud:
                    # Clear cloud-specific fields if now local
                    mapping.source.workspace_name = None
                    mapping.source.dataset_name = None
                else:
                    # Copy cloud-specific fields from target when swapping TO cloud
                    mapping.source.workspace_name = mapping.target.workspace_name
                    # SwapTarget has 'database' not 'dataset_name'
                    mapping.source.dataset_name = mapping.target.database
                # Refresh the row display with updated source
                self._update_mapping_row(idx, mapping)

            # Add to history
            run_id = self._generate_run_id()
            history_entry = SwapHistoryEntry.from_mapping(
                mapping, run_id=run_id, model_file_path=file_path or self._current_model_file_path or ""
            )
            self._swap_history.insert(0, history_entry)
            if len(self._swap_history) > self._max_history_entries:
                self._swap_history = self._swap_history[:self._max_history_entries]
            self._save_swap_history()

            # Update Last Config with workspace info if swapping to cloud
            if mapping.source.is_cloud and mapping.source.workspace_name:
                model_hash = self._get_model_hash()
                if model_hash:
                    self.preset_manager.update_last_config_workspace(
                        model_hash,
                        workspace_name=mapping.source.workspace_name,
                        friendly_name=mapping.source.dataset_name
                    )
                    self._log_message(f"Updated Last Config with workspace: {mapping.source.workspace_name}")

            if is_pbip:
                # PBIP files: PBI was closed and reopened automatically (same flow as PBIX)
                # Keep thin_report_context with file_path so user can swap back
                if self._thin_report_context:
                    target_server = mapping.target.server if mapping.target else ''
                    is_now_local = 'localhost' in target_server.lower() or target_server.startswith('127.')
                    self._thin_report_context['is_currently_local'] = is_now_local
                    self._thin_report_context['current_server'] = target_server
                    self._thin_report_context['current_database'] = mapping.target.database if mapping.target else None

                reconnect_file_path = file_path

                if auto_reconnect:
                    ThemedMessageBox.show(
                        self.frame,
                        title="Swap Complete",
                        message=f"Connection successfully swapped!\n\n"
                                f"Target: {mapping.target.server}\n\n"
                                f"Power BI Desktop is reopening the file.",
                        msg_type="info",
                        buttons=["OK"],
                        custom_icon="hotswap.svg",
                        auto_close_seconds=5
                    )
                    # Start auto-reconnect after PBIP swap
                    if reconnect_file_path:
                        self._start_auto_reconnect(reconnect_file_path)
                else:
                    ThemedMessageBox.show(
                        self.frame,
                        title="Swap Complete",
                        message=f"Connection successfully swapped!\n\n"
                                f"Target: {mapping.target.server}\n\n"
                                f"Refresh and reconnect to make further swaps.",
                        msg_type="info",
                        buttons=["OK"],
                        custom_icon="hotswap.svg"
                    )
            else:
                # PBIX files: PBI was closed and reopened automatically
                # Keep thin_report_context with file_path so user can swap back
                # The file path is still valid, just update the connection info
                if self._thin_report_context:
                    # Determine if new connection is local or cloud based on target server
                    target_server = mapping.target.server if mapping.target else ''
                    is_now_local = 'localhost' in target_server.lower() or target_server.startswith('127.')
                    self._thin_report_context['is_currently_local'] = is_now_local
                    self._thin_report_context['current_server'] = target_server
                    self._thin_report_context['current_database'] = mapping.target.database if mapping.target else None

                # Store file path for auto-reconnect
                reconnect_file_path = file_path

                if auto_reconnect:
                    ThemedMessageBox.show(
                        self.frame,
                        title="Swap Complete",
                        message=f"Connection successfully swapped!\n\n"
                                f"Target: {mapping.target.server}\n\n"
                                f"Power BI Desktop is reopening the file.",
                        msg_type="info",
                        buttons=["OK"],
                        custom_icon="hotswap.svg",
                        auto_close_seconds=5
                    )
                    # Start auto-reconnect after PBIX swap
                    if reconnect_file_path:
                        self._start_auto_reconnect(reconnect_file_path)
                else:
                    ThemedMessageBox.show(
                        self.frame,
                        title="Swap Complete",
                        message=f"Connection successfully swapped!\n\n"
                                f"Target: {mapping.target.server}\n\n"
                                f"Refresh and reconnect to make further swaps.",
                        msg_type="info",
                        buttons=["OK"],
                        custom_icon="hotswap.svg"
                    )
        else:
            ThemedMessageBox.show(
                self.frame,
                title="Swap Failed",
                message=f"Failed to swap connection:\n\n{message}",
                msg_type="error",
                buttons=["OK"],
                custom_icon="hotswap.svg"
            )

        # Re-enable buttons
        if self.mappings:
            self._on_mapping_selected()
            model_hash = self._get_model_hash()
            if model_hash and self.preset_manager.has_last_config(model_hash):
                self.last_config_btn.set_enabled(True)

    def _start_auto_reconnect(self, file_path: str):
        """
        Auto-reconnect to a thin report after PBIX swap.

        Waits for Power BI Desktop to fully load the file, then rescans
        and auto-connects to the same model.

        Args:
            file_path: Path to the PBIX file to reconnect to
        """
        import os
        normalized_path = os.path.normpath(file_path).lower()

        self._show_progress("Waiting for Power BI Desktop to load...")
        self._log_message("Auto-reconnect: Waiting for Power BI Desktop to load the file...")

        def reconnect_thread():
            import time

            # Wait for PBI to fully load (give it time to open the file)
            wait_seconds = 8
            for i in range(wait_seconds):
                remaining = wait_seconds - i
                self.frame.after(0, lambda r=remaining: self._update_progress_message(
                    f"Waiting for Power BI Desktop... ({r}s)"
                ))
                time.sleep(1)

            self.frame.after(0, lambda: self._update_progress_message("Scanning for models..."))
            self.frame.after(0, lambda: self._log_message("Auto-reconnect: Scanning for models..."))

            # Rescan for models
            try:
                models = self.connector.discover_local_models()

                # Find the model with matching file path
                matching_model = None
                matching_display_name = None
                for model in models:
                    if model.file_path:
                        model_path = os.path.normpath(model.file_path).lower()
                        if model_path == normalized_path:
                            matching_model = model
                            matching_display_name = model.display_name
                            break

                if matching_model and matching_display_name:
                    # Capture values for lambda
                    dn = matching_display_name
                    mm = matching_model
                    am = models
                    self.frame.after(0, lambda d=dn, m=mm, a=am: self._complete_auto_reconnect(d, m, a))
                else:
                    self.frame.after(0, lambda: self._auto_reconnect_failed(
                        "Could not find the model after rescan. Please refresh manually."
                    ))

            except Exception as e:
                err_msg = str(e)
                self.frame.after(0, lambda msg=err_msg: self._auto_reconnect_failed(f"Error during rescan: {msg}"))

        threading.Thread(target=reconnect_thread, daemon=True).start()

    def _complete_auto_reconnect(self, display_name: str, model, all_models: list):
        """Complete auto-reconnect by updating dropdown and connecting."""
        self._log_message(f"Auto-reconnect: Found model '{display_name}'")

        # Update dropdown cache and selection
        self._dropdown_models_cache = all_models

        # Rebuild dropdown options
        self._display_to_connection.clear()
        self.model_combo['values'] = []

        display_names = []
        for m in all_models:
            display_names.append(m.display_name)
            if m.is_thin_report:
                cloud_server = m.thin_report_cloud_server or ''
                cloud_database = m.thin_report_cloud_database or ''
                conn_str = f"{m.server}|__thin_report__:{cloud_server}:{cloud_database}"
            else:
                conn_str = f"{m.server}|{m.database_name}"
            self._display_to_connection[m.display_name] = conn_str

        self.model_combo['values'] = display_names

        # Select and connect to the matching model
        self.selected_model.set(display_name)
        self._hide_progress()
        self._log_message(f"Auto-reconnect: Connecting to '{display_name}'...")

        # Trigger connection
        self._on_connect()

        # Restore rollback button state if history exists for this model
        if self._swap_history:
            model_path = self._current_model_file_path
            if model_path:
                # Check if any history entries match current model
                matching_entries = [e for e in self._swap_history
                                   if e.model_file_path and
                                   os.path.normpath(e.model_file_path).lower() ==
                                   os.path.normpath(model_path).lower()]
                if matching_entries:
                    self.rollback_btn.set_enabled(True)

    def _auto_reconnect_failed(self, message: str):
        """Handle auto-reconnect failure."""
        self._hide_progress()
        self._log_message(f"Auto-reconnect failed: {message}")

    def _on_auto_match_changed(self):
        """Handle auto-match checkbox change"""
        if self.auto_match_enabled.get() and self.model_info:
            self._auto_match_connections()

    # =========================================================================
    # Schema Validation
    # =========================================================================

    def _on_validate_selected(self):
        """Validate schema compatibility for selected mapping"""
        selection = self.mapping_tree.selection()
        if not selection:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Select Mapping", "Please select a connection to validate.")
            return

        idx = int(selection[0])
        if idx >= len(self.mappings):
            return

        mapping = self.mappings[idx]

        if not mapping.target:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "No Target", "Please select a target first before validating.")
            return

        if not self.schema_validator:
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Error", "Schema validator not initialized. Please reconnect.")
            return

        self._show_progress(f"⏳ Validating schema for {mapping.source.name}...")
        self._log_message(f"⏳ Validating schema compatibility...")

        def validate_thread():
            result = self.schema_validator.validate_mapping(mapping)
            self.frame.after(0, lambda: self._show_validation_result(result))

        threading.Thread(target=validate_thread, daemon=True).start()

    def _show_validation_result(self, result: ValidationResult):
        """Show the validation result in a dialog"""
        self._hide_progress()

        # Log summary
        self._log_message(f"🔍 Validation: {result.summary}")

        # Create validation result dialog
        dialog = tk.Toplevel(self.frame)
        dialog.title("Schema Validation Results")
        dialog.geometry("600x450")
        dialog.transient(self.frame.winfo_toplevel())
        dialog.grab_set()

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        bg_color = colors.get('background', '#1e1e2e' if is_dark else '#f5f5fa')
        surface_color = colors.get('surface', '#2a2a3a' if is_dark else '#ffffff')
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')

        dialog.configure(bg=bg_color)

        # Header
        header_frame = tk.Frame(dialog, bg=bg_color)
        header_frame.pack(fill=tk.X, padx=16, pady=(16, 8))

        # Status icon and title
        if result.is_compatible:
            status_icon = "✅"
            status_text = "Schemas Compatible"
            status_color = colors.get('success', '#10b981')
        elif result.has_errors:
            status_icon = "❌"
            status_text = "Compatibility Issues Found"
            status_color = colors.get('error', '#ef4444')
        else:
            status_icon = "⚠️"
            status_text = "Potential Issues Found"
            status_color = colors.get('warning', '#f59e0b')

        tk.Label(
            header_frame,
            text=f"{status_icon} {status_text}",
            font=('Segoe UI Semibold', 14),
            fg=status_color,
            bg=bg_color
        ).pack(anchor='w')

        tk.Label(
            header_frame,
            text=f"Source: {result.source_name}  →  Target: {result.target_name}",
            font=('Segoe UI', 10),
            fg=colors.get('text_muted', '#6b7280'),
            bg=bg_color
        ).pack(anchor='w', pady=(4, 0))

        # Summary line
        tk.Label(
            header_frame,
            text=result.summary,
            font=('Segoe UI', 10),
            fg=colors.get('text_primary', '#e0e0e0' if is_dark else '#333333'),
            bg=bg_color
        ).pack(anchor='w', pady=(4, 0))

        # Findings list
        findings_frame = tk.Frame(dialog, bg=surface_color, highlightbackground=border_color,
                                   highlightcolor=border_color, highlightthickness=1)
        findings_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(8, 16))

        # Findings listbox with scrollbar
        list_container = tk.Frame(findings_frame, bg=surface_color)
        list_container.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        findings_text = tk.Text(
            list_container,
            wrap=tk.WORD,
            bg=surface_color,
            fg=colors.get('text_primary', '#e0e0e0' if is_dark else '#333333'),
            font=('Segoe UI', 10),
            relief='flat',
            padx=8,
            pady=8,
            yscrollcommand=scrollbar.set
        )
        findings_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=findings_text.yview)

        # Define text tags for different severities
        findings_text.tag_configure('error', foreground=colors.get('error', '#ef4444'))
        findings_text.tag_configure('warning', foreground=colors.get('warning', '#f59e0b'))
        findings_text.tag_configure('info', foreground=colors.get('text_muted', '#6b7280'))
        findings_text.tag_configure('category', foreground=colors.get('primary', '#4a6cf5'), font=('Segoe UI Semibold', 10))

        # Populate findings
        if result.findings:
            for finding in result.findings:
                # Insert icon and message
                findings_text.insert(tk.END, f"{finding.icon} ")
                findings_text.insert(tk.END, f"[{finding.category}] ", 'category')

                tag = finding.severity.value
                findings_text.insert(tk.END, f"{finding.message}\n", tag)

                if finding.details:
                    findings_text.insert(tk.END, f"   {finding.details}\n", 'info')

                findings_text.insert(tk.END, "\n")
        else:
            findings_text.insert(tk.END, "No findings - schemas appear compatible.\n", 'info')

        findings_text.configure(state='disabled')

        # Button frame
        button_frame = tk.Frame(dialog, bg=bg_color)
        button_frame.pack(fill=tk.X, padx=16, pady=(0, 16))

        # Proceed anyway button (if there are warnings but no errors)
        if result.has_warnings and not result.has_errors:
            proceed_btn = RoundedButton(
                button_frame,
                text="PROCEED",
                command=lambda: self._proceed_after_validation(dialog, result),
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg='#ffffff',
                height=32, radius=6,
                font=('Segoe UI', 10, 'bold'),
                canvas_bg=bg_color
            )
            proceed_btn.pack(side=tk.LEFT, padx=(0, 8))

        close_btn = RoundedButton(
            button_frame,
            text="CLOSE",
            command=dialog.destroy,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=32, radius=6,
            font=('Segoe UI', 10),
            canvas_bg=bg_color
        )
        close_btn.pack(side=tk.RIGHT)

        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 600) // 2
        y = (dialog.winfo_screenheight() - 450) // 2
        dialog.geometry(f"600x450+{x}+{y}")

    def _proceed_after_validation(self, dialog, result: ValidationResult):
        """Close validation dialog and proceed to swap"""
        dialog.destroy()

        # Find and execute the swap for the selected mapping
        selection = self.mapping_tree.selection()
        if selection:
            idx = int(selection[0])
            if idx < len(self.mappings):
                mapping = self.mappings[idx]
                if mapping.target:
                    self._execute_swap(mapping, idx)

    # =========================================================================
    # Selection and Target Methods
    # =========================================================================

    def _on_select_all_connections(self):
        """Select all rows in the connections/mapping tree"""
        if hasattr(self, 'mapping_tree') and self.mapping_tree:
            all_items = self.mapping_tree.get_children()
            if all_items:
                self.mapping_tree.selection_set(all_items)
                # Trigger selection event to update button states
                self._on_mapping_selected()

    def _on_mapping_selected(self, event=None):
        """Handle selection change in the mapping treeview"""
        selection = self.mapping_tree.selection()

        # Update button states based on selection
        has_selection = len(selection) > 0
        has_ready_mappings = any(
            self.mappings[int(item_id)].status == SwapStatus.READY
            for item_id in selection
            if int(item_id) < len(self.mappings)
        )

        # Enable swap button if any selected mappings are ready
        self.swap_selected_btn.set_enabled(has_selection and has_ready_mappings)

        # Update connection details panel with selected mapping info
        if has_selection and selection:
            try:
                item_id = selection[0]
                idx = int(item_id)
                if idx < len(self.mappings):
                    self._update_selected_connection_details(self.mappings[idx])
            except (ValueError, IndexError):
                pass
        elif self.model_info:
            # No selection, show general model info
            self._update_connection_details()

    def _create_swap_target_from_local_model(self, model) -> SwapTarget:
        """Create a SwapTarget from a local model info object"""
        return SwapTarget(
            target_type="local",
            server=model.server,
            database=model.database_name,
            display_name=model.display_name
        )

    def _refresh_local_models(self):
        """Refresh the local models cache for the inline target picker"""
        if not self.matcher:
            return

        # Use cached models from dropdown if available (avoids redundant port scanning)
        if hasattr(self, '_dropdown_models_cache') and self._dropdown_models_cache:
            # Filter to non-thin-reports and models not the current connected model
            local_models = [
                m for m in self._dropdown_models_cache
                if not m.is_thin_report
            ]
            if local_models:
                self._log_message(f"Using {len(local_models)} cached local model(s)")
                self._update_local_models_cache(local_models)
                return

        # Only scan if no cached models available
        self._log_message("Scanning for local Power BI models...")

        def scan_thread():
            try:
                models = self.matcher.discover_local_models()
                self.frame.after(0, lambda: self._update_local_models_cache(models))
            except Exception as e:
                self.frame.after(0, lambda: self._log_message(f"Error scanning: {e}"))

        threading.Thread(target=scan_thread, daemon=True).start()

    def _update_local_models_cache(self, models):
        """Update the local models cache with discovered models"""
        # Convert discovered models to SwapTarget objects for the picker
        self._local_models_cache = [
            self._create_swap_target_from_local_model(model)
            for model in models
        ]

        if models:
            self._log_message(f"Found {len(models)} local model(s)")
        else:
            self._log_message("No local models found. Ensure Power BI Desktop has models open.")

    # -------------------------------------------------------------------------
    # Cloud Authentication Dropdown
    # -------------------------------------------------------------------------

    def _toggle_cloud_auth_dropdown(self):
        """Toggle the cloud authentication dropdown."""
        if self._cloud_auth_dropdown and self._cloud_auth_dropdown.winfo_exists():
            self._close_cloud_auth_dropdown()
        else:
            self._open_cloud_auth_dropdown()

    def _open_cloud_auth_dropdown(self):
        """Open the cloud auth dropdown menu."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create popup window
        popup = tk.Toplevel(self.frame)
        popup.withdraw()
        popup.overrideredirect(True)

        # Border frame (1px effect)
        border_frame = tk.Frame(popup, bg=colors['border'])
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Content frame
        surface_bg = colors.get('surface', colors['background'])
        content = tk.Frame(border_frame, bg=surface_bg)
        content.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # Get auth state
        email = self.cloud_browser.get_account_email() if self.cloud_browser else None
        is_signed_in = email is not None

        # If signed in, show email (non-clickable, muted)
        if is_signed_in:
            email_label = tk.Label(
                content, text=email,
                fg=colors['text_muted'], bg=surface_bg,
                font=('Segoe UI', 9), padx=8, pady=6, anchor='w'
            )
            email_label.pack(fill=tk.X)

            # Separator
            sep = tk.Frame(content, height=1, bg=colors['border'])
            sep.pack(fill=tk.X)

        # Sign In / Sign Out button
        action_text = "Sign Out" if is_signed_in else "Sign In"
        action_frame = tk.Frame(content, bg=surface_bg)
        action_frame.pack(fill=tk.X)

        action_label = tk.Label(
            action_frame, text=action_text,
            fg=colors['text_primary'], bg=surface_bg,
            font=('Segoe UI', 9), padx=8, pady=8, cursor='hand2', anchor='w'
        )
        action_label.pack(fill=tk.X)

        # Hover effect
        hover_bg = colors.get('card_surface_hover', surface_bg)

        def on_hover(entering: bool):
            bg = hover_bg if entering else surface_bg
            action_frame.config(bg=bg)
            action_label.config(bg=bg)

        action_frame.bind('<Enter>', lambda e: on_hover(True))
        action_frame.bind('<Leave>', lambda e: on_hover(False))
        action_label.bind('<Enter>', lambda e: on_hover(True))
        action_label.bind('<Leave>', lambda e: on_hover(False))

        # Click action
        def on_click(e=None):
            self._on_cloud_auth_action(is_signed_in)

        action_label.bind('<Button-1>', on_click)
        action_frame.bind('<Button-1>', on_click)

        # Position popup below button, right-aligned
        if hasattr(self, '_cloud_auth_btn'):
            btn = self._cloud_auth_btn
            popup.update_idletasks()

            btn_x = btn.winfo_rootx()
            btn_y = btn.winfo_rooty() + btn.winfo_height()
            popup_width = popup.winfo_reqwidth()

            # Right-align popup to button
            popup_x = btn_x + btn.winfo_width() - popup_width

            popup.geometry(f"+{popup_x}+{btn_y}")

        popup.deiconify()
        self._cloud_auth_dropdown = popup

        # Click-outside detection
        self._cloud_auth_click_handler = lambda e: self._on_cloud_auth_click_outside(e)
        self.frame.winfo_toplevel().bind('<Button-1>', self._cloud_auth_click_handler, add='+')

    def _on_cloud_auth_click_outside(self, event):
        """Handle click outside the dropdown."""
        if not self._cloud_auth_dropdown or not self._cloud_auth_dropdown.winfo_exists():
            return

        x, y = event.x_root, event.y_root

        # Check if click is inside dropdown
        dx = self._cloud_auth_dropdown.winfo_rootx()
        dy = self._cloud_auth_dropdown.winfo_rooty()
        dw = self._cloud_auth_dropdown.winfo_width()
        dh = self._cloud_auth_dropdown.winfo_height()

        if dx <= x <= dx + dw and dy <= y <= dy + dh:
            return  # Click inside dropdown

        # Check if click is on button
        if hasattr(self, '_cloud_auth_btn'):
            bx = self._cloud_auth_btn.winfo_rootx()
            by = self._cloud_auth_btn.winfo_rooty()
            bw = self._cloud_auth_btn.winfo_width()
            bh = self._cloud_auth_btn.winfo_height()

            if bx <= x <= bx + bw and by <= y <= by + bh:
                return  # Click on button (toggle handles it)

        # Click outside - close dropdown
        self._close_cloud_auth_dropdown()

    def _close_cloud_auth_dropdown(self):
        """Close the cloud auth dropdown."""
        if self._cloud_auth_dropdown and self._cloud_auth_dropdown.winfo_exists():
            self._cloud_auth_dropdown.destroy()
        self._cloud_auth_dropdown = None

        # Remove click handler
        if hasattr(self, '_cloud_auth_click_handler'):
            try:
                self.frame.winfo_toplevel().unbind('<Button-1>', self._cloud_auth_click_handler)
            except Exception:
                pass

    def _on_cloud_auth_action(self, is_signed_in: bool):
        """Handle sign in or sign out action."""
        self._close_cloud_auth_dropdown()

        if is_signed_in:
            # Sign out
            if self.cloud_browser:
                self.cloud_browser.sign_out()
                self._log_message("Signed out from cloud account")
            self._update_cloud_auth_button_state()
        else:
            # Sign in - just authenticate, don't open the cloud browser dialog
            if not self.cloud_browser:
                # Initialize cloud browser if needed
                self.cloud_browser = CloudWorkspaceBrowser()

            # Run authentication (opens browser window)
            success, message = self.cloud_browser.authenticate()
            if success:
                self._log_message(f"Cloud authentication: {message}")
                # Pre-cache workspaces so cloud browser opens instantly
                self._precache_cloud_workspaces()
            else:
                self._log_message(f"Cloud authentication failed: {message}")

            # Update button state
            self._update_cloud_auth_button_state()

    def _precache_cloud_workspaces(self):
        """
        Pre-fetch and cache cloud workspaces in the background.

        Called after successful authentication so the cloud browser dialog
        opens instantly with data already loaded.
        """
        import threading

        def _fetch_in_background():
            try:
                if self.cloud_browser:
                    workspaces, error = self.cloud_browser.get_workspaces("all")
                    if workspaces and not error:
                        self._log_message(f"Pre-cached {len(workspaces)} cloud workspaces")
            except Exception as e:
                # Silent failure - this is just a cache optimization
                pass

        # Run in background so it doesn't block the UI
        thread = threading.Thread(target=_fetch_in_background, daemon=True)
        thread.start()

    def _update_cloud_auth_button_state(self):
        """Update button icon based on auth state."""
        if not hasattr(self, '_cloud_auth_btn'):
            return

        if self.cloud_browser:
            email = self.cloud_browser.get_account_email()
            is_signed_in = email is not None
        else:
            is_signed_in = False

        # Update icon (colored vs grayscale)
        icon = self._cloud_icon_colored if is_signed_in else self._cloud_icon_gray
        if icon:
            self._cloud_auth_btn.update_icon(icon)

        # Update tooltip
        if is_signed_in and email:
            self._cloud_auth_btn._tooltip_text = f"Cloud Account: {email}"
        else:
            self._cloud_auth_btn._tooltip_text = "Cloud Account (Not signed in)"

    def _on_browse_cloud(self):
        """Open the cloud browser dialog"""
        if not self.cloud_browser:
            ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "Not Connected", "Please connect to a model first.")
            return

        from core.cloud import CloudBrowserDialog
        from tools.connection_hotswap.models import CloudConnectionType

        # Get current model status for dialog display
        current_status = self.connection_status.get()
        model_status = current_status if current_status and current_status != "Not Connected" else None

        # Build source context if we know which mapping we're editing
        source_context = None
        if hasattr(self, '_inline_edit_item_id') and self._inline_edit_item_id:
            try:
                idx = int(self._inline_edit_item_id)
                if 0 <= idx < len(self.mappings):
                    mapping = self.mappings[idx]
                    source_name = mapping.source.dataset_name or mapping.source.database or mapping.source.name
                    source_type = "Cloud" if mapping.source.is_cloud else "Local"
                    source_context = f"{source_name} ({source_type})"
            except (ValueError, IndexError):
                pass

        # Always default to Semantic Model connector when manually selecting cloud target
        dialog = CloudBrowserDialog(
            self.frame.winfo_toplevel(),
            self.cloud_browser,
            default_connection_type=CloudConnectionType.PBI_SEMANTIC_MODEL,
            model_status=model_status,
            on_auth_change=self._update_cloud_auth_button_state,
            source_context=source_context
        )
        dialog.wait_window()

        if dialog.result:
            # If called from inline picker, apply to specific item
            if hasattr(self, '_inline_edit_item_id') and self._inline_edit_item_id:
                self._on_inline_target_selected(self._inline_edit_item_id, dialog.result)
                self._inline_edit_item_id = None
            else:
                # Otherwise, just log the selection (legacy behavior)
                self._log_message(f"Selected cloud target: {dialog.result.display_name}")

    def _apply_target_to_mappings(self, target: SwapTarget, mappings_indices: List[int]):
        """Apply a target to specified mapping indices"""
        applied_count = 0
        for idx in mappings_indices:
            if idx >= len(self.mappings):
                continue

            mapping = self.mappings[idx]
            mapping.target = target
            mapping.status = SwapStatus.READY

            # Update the UI row
            self._update_mapping_row(idx, mapping)
            applied_count += 1

        # Log message
        if applied_count == 1:
            mapping = self.mappings[mappings_indices[0]]
            self._log_message(f"Target set: {mapping.source.name} -> {target.display_name}")
        else:
            self._log_message(f"Batch applied: {target.display_name} -> {applied_count} connection(s)")

        # Update swap button based on selection (requires user to select rows)
        self._on_mapping_selected()

        return applied_count

    def _on_rollback_last(self):
        """Show rollback history dialog or rollback last if only one entry"""
        if not self._swap_history:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No History", "No swap to rollback. Perform a swap first.")
            return

        # For thin reports, we don't need a swapper (rollback is done via file modification)
        if not self.swapper and not self._thin_report_context:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Connection", "Not connected to a model. Connect first to enable rollback.")
            return

        # Filter entries for current model
        import os
        current_model_path = self._current_model_file_path
        if current_model_path:
            current_model_path_normalized = os.path.normpath(current_model_path).lower()
        else:
            current_model_path_normalized = None

        filtered_entries = []
        for entry in self._swap_history:
            entry_path = getattr(entry, 'model_file_path', '')
            if not entry_path:
                # Old entries without model_file_path - include for backward compatibility
                filtered_entries.append(entry)
            elif current_model_path_normalized:
                entry_path_normalized = os.path.normpath(entry_path).lower()
                if entry_path_normalized == current_model_path_normalized:
                    filtered_entries.append(entry)

        if not filtered_entries:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No History", "No swap history for this model.")
            return

        # If only one filtered entry, rollback directly
        if len(filtered_entries) == 1:
            self._perform_rollback(filtered_entries[0])
        else:
            # Show history dialog
            self._show_rollback_history_dialog()

    def _show_rollback_history_dialog(self):
        """Show dialog with swap history for rollback selection"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        parent = self.frame.winfo_toplevel()

        # Wider and taller dialog as requested
        dialog_width = 700
        dialog_height = 450

        dialog = tk.Toplevel(parent)
        dialog.title("Swap History")
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.resizable(True, True)
        dialog.minsize(600, 400)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

        # Set AE favicon
        try:
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'config'):
                dialog.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Set dark title bar on Windows
        try:
            from ctypes import windll, byref, sizeof, c_int
            HWND = windll.user32.GetParent(dialog.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = c_int(1 if is_dark else 0)
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_USE_IMMERSIVE_DARK_MODE, byref(value), sizeof(value))
        except Exception:
            pass

        # Center on parent
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - dialog_width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - dialog_height) // 2
        dialog.geometry(f"+{x}+{y}")

        # Content
        frame = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        frame.pack(fill=tk.BOTH, expand=True)

        # Header matching connections table design (icon + text)
        header_frame = tk.Frame(frame, bg=colors['background'])
        header_frame.pack(fill=tk.X, pady=(0, 12))

        # Use reset icon for header (matches rollback functionality)
        reset_icon = self._button_icons.get('reset')
        if reset_icon:
            icon_label = tk.Label(header_frame, image=reset_icon, bg=colors['background'])
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            icon_label._icon_ref = reset_icon

        tk.Label(
            header_frame,
            text="Swap History",
            font=('Segoe UI Semibold', 11),
            fg=colors['title_color'],
            bg=colors['background']
        ).pack(side=tk.LEFT)

        # Info label (matching connections table header style)
        tk.Label(
            header_frame,
            text="Select an entry to rollback",
            font=("Segoe UI", 9, "italic"),
            fg=colors['text_muted'],
            bg=colors['background']
        ).pack(side=tk.RIGHT)

        # History list - styled to match connections table
        content_bg = '#161627' if is_dark else '#f5f5f7'
        tree_border = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')

        list_frame = tk.Frame(
            frame,
            bg=content_bg,
            highlightbackground=tree_border,
            highlightcolor=tree_border,
            highlightthickness=1
        )
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        # Configure treeview style to match connections table
        style = ttk.Style()
        tree_style = "HistoryDialog.Treeview"
        style.configure(tree_style,
                        background=content_bg,
                        foreground=colors.get('text_primary', '#e0e0e0' if is_dark else '#333333'),
                        fieldbackground=content_bg,
                        font=('Segoe UI', 9),
                        relief='flat',
                        borderwidth=0,
                        bordercolor=content_bg,
                        lightcolor=content_bg,
                        darkcolor=content_bg,
                        rowheight=28)
        style.layout(tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])

        # Modern heading style matching connections table
        heading_bg = colors.get('section_bg', '#1a1a2a' if is_dark else '#f5f5fa')
        heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        header_separator = '#0d0d1a' if is_dark else '#ffffff'

        style.configure(f"{tree_style}.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        font=('Segoe UI', 9, 'bold'),
                        padding=(8, 6))
        style.map(f"{tree_style}.Heading",
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)],
                  relief=[('active', 'groove'), ('pressed', 'groove')])
        style.map(tree_style,
                  background=[('selected', '#1a3a5c' if is_dark else '#e6f3ff')],
                  foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])

        # Treeview for history - columns: Date/Time, Swap Type, Source, Target
        columns = ("time", "swap_type", "source", "target")
        history_tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",  # No tree column, just headings
            selectmode="browse",
            height=8,
            style=tree_style
        )

        history_tree.heading("time", text="Date/Time")
        history_tree.heading("swap_type", text="Swap Type")
        history_tree.heading("source", text="Source")
        history_tree.heading("target", text="Target")

        # Column widths for the larger dialog
        history_tree.column("time", width=110, minwidth=90, anchor="center")
        history_tree.column("swap_type", width=120, minwidth=100, anchor="center")
        history_tree.column("source", width=180, minwidth=120, anchor="center")
        history_tree.column("target", width=180, minwidth=120, anchor="center")

        # Scrollbar area - styled to match connections table
        scrollbar_bg = '#1a1a2e' if is_dark else '#f0f0f0'
        scrollbar_area = tk.Frame(list_frame, bg=scrollbar_bg, width=12)
        scrollbar_area.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_area.pack_propagate(False)

        scrollbar = ThemedScrollbar(scrollbar_area, command=history_tree.yview, theme_manager=self._theme_manager, auto_hide=True)
        history_tree.configure(yscrollcommand=scrollbar.set)

        history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(fill=tk.Y, expand=True)

        # Filter entries to current model only - no backward compatibility, only show entries for this file
        from collections import OrderedDict
        from datetime import datetime
        import os

        current_model_path = self._current_model_file_path
        if current_model_path:
            current_model_path_normalized = os.path.normpath(current_model_path).lower()
        else:
            current_model_path_normalized = None

        # Filter swap history to only entries for this specific model
        filtered_history = []
        if current_model_path_normalized:
            for i, entry in enumerate(self._swap_history):
                entry_path = getattr(entry, 'model_file_path', '')
                if entry_path:
                    entry_path_normalized = os.path.normpath(entry_path).lower()
                    if entry_path_normalized == current_model_path_normalized:
                        filtered_history.append((i, entry))

        # Group entries by run_id
        runs = OrderedDict()  # run_id -> list of (index, entry)
        for i, entry in filtered_history:
            run_id = entry.run_id or f"single_{i}"  # Fallback for old entries without run_id
            if run_id not in runs:
                runs[run_id] = []
            runs[run_id].append((i, entry))

        # Helper to format timestamp
        def format_timestamp(timestamp):
            try:
                if 'T' in timestamp:
                    dt = datetime.fromisoformat(timestamp[:19])
                    return dt.strftime("%b %d, %H:%M")
                return timestamp[:16]
            except Exception:
                return timestamp[:16] if len(timestamp) > 16 else timestamp

        # Populate with grouped runs (limit to last 10 runs)
        # Each row shows: Date/Time, Model (File), Source, Target
        run_count = 0
        entry_to_run_mapping = {}  # Maps tree item id to list of entry indices for apply config

        for run_id, entries in runs.items():
            if run_count >= 10:
                break
            run_count += 1

            first_entry = entries[0][1]
            time_str = format_timestamp(first_entry.timestamp)

            if len(entries) == 1:
                # Single swap - show directly
                idx, entry = entries[0]

                # Get swap type using stored types (with fallback to server inference)
                swap_type = self._get_swap_type(
                    entry.original_server,
                    entry.new_server,
                    stored_source_type=entry.source_type,
                    stored_target_type=entry.target_type
                )

                # Get friendly display names for source and target
                source_display = self._get_friendly_name(entry.original_database, entry.original_server)
                target_display = self._get_friendly_name(entry.new_database, entry.new_server)

                item_id = f"entry_{idx}"
                history_tree.insert(
                    "",
                    "end",
                    iid=item_id,
                    values=(time_str, swap_type, source_display, target_display)
                )
                entry_to_run_mapping[item_id] = [idx]
            else:
                # Batch swap - show summary row
                # Aggregate source -> target info and determine swap type
                sources = set()
                targets = set()
                swap_types = set()
                for idx, entry in entries:
                    sources.add(self._get_friendly_name(entry.original_database, entry.original_server))
                    targets.add(self._get_friendly_name(entry.new_database, entry.new_server))
                    swap_types.add(self._get_swap_type(
                        entry.original_server,
                        entry.new_server,
                        stored_source_type=entry.source_type,
                        stored_target_type=entry.target_type
                    ))

                source_summary = list(sources)[0] if len(sources) == 1 else f"{len(entries)} connections"
                target_summary = list(targets)[0] if len(targets) == 1 else f"{len(entries)} targets"
                swap_type_summary = list(swap_types)[0] if len(swap_types) == 1 else "Mixed"

                run_item_id = f"run_{run_id}"
                history_tree.insert(
                    "",
                    "end",
                    iid=run_item_id,
                    values=(time_str, swap_type_summary, source_summary, target_summary)
                )
                entry_to_run_mapping[run_item_id] = [idx for idx, _ in entries]

        # Buttons with icons
        canvas_bg = colors['background']
        btn_frame = tk.Frame(frame, bg=colors['background'])
        btn_frame.pack(fill=tk.X)

        # Get icons for buttons
        apply_icon = self._button_icons.get('magnifying-glass')  # Same as Apply Preset
        eraser_icon = self._button_icons.get('eraser')

        def get_selected_history_entry():
            """Get the first history entry from the selected tree item."""
            selection = history_tree.selection()
            if not selection:
                return None
            item_id = selection[0]
            entry_indices = entry_to_run_mapping.get(item_id, [])
            if not entry_indices:
                return None
            idx = entry_indices[0]
            if idx < len(self._swap_history):
                return self._swap_history[idx]
            return None

        def get_selected_mapping():
            """Get the currently selected mapping from the main connection table."""
            selection = self.mapping_tree.selection()
            if not selection:
                return None, None
            try:
                idx = int(selection[0])
                if idx < len(self.mappings):
                    return idx, self.mappings[idx]
            except (ValueError, IndexError):
                pass
            return None, None

        def apply_history_to_connection(use_source: bool):
            """Apply the source or target from history entry to the selected connection."""
            # Check if a history entry is selected
            entry = get_selected_history_entry()
            if not entry:
                ThemedMessageBox.showinfo(dialog, "Select History Entry",
                    "Please select a history entry first.")
                return

            # Check if a connection is selected in the main table
            mapping_idx, mapping = get_selected_mapping()
            if mapping is None:
                ThemedMessageBox.showinfo(dialog, "Select Connection",
                    "Please select a connection in the main table first.\n\n"
                    "The history entry will be applied to that connection's target.")
                return

            # Get server, database, and connection type from history entry
            if use_source:
                server = entry.original_server
                database = entry.original_database
                history_type = entry.source_type  # "Local", "Cloud", or "XMLA"
                label = "source"
            else:
                server = entry.new_server
                database = entry.new_database
                history_type = entry.target_type  # "Local", "Cloud", or "XMLA"
                label = "target"

            # Determine if this is a local or cloud target based on the server
            is_cloud = ('powerbi://' in server.lower() or
                       'pbidedicated://' in server.lower() or
                       'pbiazure://' in server.lower())

            # Map history type string to CloudConnectionType enum for proper connection building
            cloud_connection_type = None
            if is_cloud:
                from tools.connection_hotswap.models import CloudConnectionType
                if history_type == "XMLA":
                    cloud_connection_type = CloudConnectionType.AAS_XMLA
                else:
                    # Default to Semantic Model for "Cloud" or unknown types
                    cloud_connection_type = CloudConnectionType.PBI_SEMANTIC_MODEL

            # Extract workspace name from cloud server URL if applicable
            workspace_name = None
            if is_cloud and '/v1.0/' in server:
                try:
                    import urllib.parse
                    parts = server.split('/v1.0/')
                    if len(parts) > 1:
                        path_part = parts[1]
                        segments = path_part.split('/')
                        if len(segments) >= 2 and segments[0] == 'myorg':
                            workspace_name = urllib.parse.unquote(segments[1])
                except Exception:
                    pass

            # Create a SwapTarget from the history entry
            target = SwapTarget(
                target_type="cloud" if is_cloud else "local",
                server=server,
                database=database,
                display_name=self._get_friendly_name(database, server),
                workspace_name=workspace_name,
                cloud_connection_type=cloud_connection_type
            )

            # Set the target on the mapping
            mapping.target = target
            mapping.status = SwapStatus.READY

            # Update the UI row
            self._update_mapping_row(mapping_idx, mapping)

            dialog.destroy()
            self._log_message(f"Applied history {label} to '{mapping.source.name}' - ready for swap")

        def update_apply_buttons_state(event=None):
            """Update the enabled state and tooltips of apply buttons based on selections."""
            history_selected = bool(history_tree.selection())
            mapping_idx, _ = get_selected_mapping()
            connection_selected = mapping_idx is not None

            if not history_selected:
                # No history selected - disable both with tooltip
                apply_source_btn.set_enabled(False)
                apply_target_btn.set_enabled(False)
                apply_source_btn._tooltip_text = "Select a history entry first"
                apply_target_btn._tooltip_text = "Select a history entry first"
            elif not connection_selected:
                # History selected but no connection - enable but tooltip explains
                apply_source_btn.set_enabled(True)
                apply_target_btn.set_enabled(True)
                apply_source_btn._tooltip_text = "Apply this history's SOURCE to selected connection\n(Select a connection in the main table first)"
                apply_target_btn._tooltip_text = "Apply this history's TARGET to selected connection\n(Select a connection in the main table first)"
            else:
                # Both selected - fully enabled with clear tooltips
                apply_source_btn.set_enabled(True)
                apply_target_btn.set_enabled(True)
                apply_source_btn._tooltip_text = "Apply this history's SOURCE to the selected connection"
                apply_target_btn._tooltip_text = "Apply this history's TARGET to the selected connection"

        # Bind selection change on history tree to update button states
        history_tree.bind('<<TreeviewSelect>>', update_apply_buttons_state)

        # APPLY SOURCE button
        apply_source_btn = RoundedButton(
            btn_frame, text="APPLY SOURCE",
            command=lambda: apply_history_to_connection(use_source=True),
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 10, 'bold'),
            canvas_bg=canvas_bg, icon=apply_icon
        )
        apply_source_btn.pack(side=tk.LEFT, padx=(0, 8))
        apply_source_btn.set_enabled(False)  # Start disabled
        apply_source_btn._tooltip_text = "Select a history entry first"

        # APPLY TARGET button
        apply_target_btn = RoundedButton(
            btn_frame, text="APPLY TARGET",
            command=lambda: apply_history_to_connection(use_source=False),
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 10, 'bold'),
            canvas_bg=canvas_bg, icon=apply_icon
        )
        apply_target_btn.pack(side=tk.LEFT, padx=(0, 8))
        apply_target_btn.set_enabled(False)  # Start disabled
        apply_target_btn._tooltip_text = "Select a history entry first"

        # Add dynamic tooltips - use callable to get current tooltip text
        Tooltip(apply_source_btn, lambda: apply_source_btn._tooltip_text)
        Tooltip(apply_target_btn, lambda: apply_target_btn._tooltip_text)

        def on_clear_history():
            """Clear history only for the currently connected model."""
            model_name = os.path.basename(current_model_path) if current_model_path else "this model"
            confirm = ThemedMessageBox.askyesno(
                dialog,
                "Clear History",
                f"Clear swap history for {model_name}?\n\nThis cannot be undone."
            )
            if confirm:
                # Only remove entries for the current model, keep others
                if current_model_path_normalized:
                    entries_to_keep = []
                    for entry in self._swap_history:
                        entry_path = getattr(entry, 'model_file_path', '')
                        if entry_path:
                            entry_path_normalized = os.path.normpath(entry_path).lower()
                            if entry_path_normalized != current_model_path_normalized:
                                entries_to_keep.append(entry)
                        # Keep entries without path (legacy) in the global history
                        else:
                            entries_to_keep.append(entry)
                    self._swap_history.clear()
                    self._swap_history.extend(entries_to_keep)
                else:
                    # No model connected - shouldn't happen, but just in case
                    pass

                self._last_swapped_mapping = None
                self.rollback_btn.set_enabled(False)
                self._log_message(f"Cleared swap history for {model_name}")
                dialog.destroy()

        clear_btn = RoundedButton(
            btn_frame, text="CLEAR",
            command=on_clear_history,
            bg=colors.get('error', '#ef4444'),
            hover_bg=colors.get('error_hover', '#dc2626'),
            pressed_bg=colors.get('error_pressed', '#b91c1c'),
            fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 10, 'bold'),
            canvas_bg=canvas_bg, icon=eraser_icon
        )
        clear_btn.pack(side=tk.LEFT, padx=(0, 8))

        close_btn = RoundedButton(
            btn_frame, text="CLOSE",
            command=dialog.destroy,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 10),
            canvas_bg=canvas_bg
        )
        close_btn.pack(side=tk.RIGHT)

    def _perform_rollback(self, entry: SwapHistoryEntry):
        """Perform the actual rollback from a history entry"""
        # Check if this is a thin report rollback
        if self._thin_report_context:
            self._perform_thin_report_rollback(entry)
            return

        if not self.swapper:
            self._log_message("❌ Cannot rollback: No connection to model")
            return

        # Find the matching mapping
        mapping = None
        for m in self.mappings:
            if m.source.name == entry.connection_name:
                mapping = m
                break

        if not mapping:
            self._log_message(f"❌ Cannot find mapping for {entry.connection_name}")
            return

        # Restore original connection string for rollback
        mapping.original_connection_string = entry.original_connection_string

        self._show_progress(f"⏳ Rolling back {entry.connection_name}...")
        self._log_message(f"⏳ Rolling back {entry.connection_name}...")

        def rollback_thread():
            try:
                result = self.swapper.rollback_connection(mapping)
                success = result.success
                message = result.message

                def on_complete():
                    self._hide_progress()
                    if success:
                        # Update the mapping source to reflect rolled-back state
                        for i, m in enumerate(self.mappings):
                            if m.source.name == entry.connection_name:
                                # Update source connection info to original values
                                m.source.server = entry.original_server
                                m.source.database = entry.original_database
                                m.source.connection_string = entry.original_connection_string
                                m.source.is_cloud = 'powerbi://' in entry.original_server.lower()
                                m.status = SwapStatus.PENDING
                                m.target = None
                                self._update_mapping_row(i, m)
                                break

                        # Remove this entry from history
                        self._swap_history = [h for h in self._swap_history
                                              if h.timestamp != entry.timestamp]

                        # Persist updated history
                        self._save_swap_history()

                        # Update button state
                        if not self._swap_history:
                            self._last_swapped_mapping = None
                            self.rollback_btn.set_enabled(False)

                        self._log_message(f"✅ Rollback successful: {message}")
                    else:
                        self._log_message(f"❌ Rollback failed: {message}")

                self.frame.after(0, on_complete)

            except Exception as e:
                self.frame.after(0, lambda: self._hide_progress())
                self.frame.after(0, lambda: self._log_message(f"❌ Rollback error: {e}"))

        threading.Thread(target=rollback_thread, daemon=True).start()

    def _perform_thin_report_rollback(self, entry: SwapHistoryEntry):
        """Perform rollback for a thin report by modifying the file."""
        file_path = self._thin_report_context.get('file_path')
        if not file_path:
            self._log_message("❌ Cannot rollback: File path not available")
            return

        # Parse original server and database from history entry
        original_server = entry.original_server
        original_database = entry.original_database

        if not original_server or not original_database:
            self._log_message("❌ Cannot rollback: Original connection info not available")
            return

        self._show_progress(f"⏳ Rolling back to {original_server}...")
        self._log_message(f"Rolling back to: {original_server}")

        # Detect file type
        modifier = get_pbix_modifier()
        file_type = modifier.detect_file_type(file_path)

        if file_type == 'pbip':
            # PBIP: Modify directly
            self._execute_thin_report_rollback_pbip(entry, file_path, original_server, original_database)
        elif file_type == 'pbix':
            # PBIX: Need to close/reopen
            self._show_thin_report_rollback_confirmation(entry, file_path, original_server, original_database)
        else:
            self._hide_progress()
            self._log_message(f"❌ Unknown file type for rollback: {file_path}")

    def _execute_thin_report_rollback_pbip(self, entry: SwapHistoryEntry, file_path: str,
                                           server: str, database: str):
        """Execute rollback for PBIP thin report."""
        create_backup = self.preset_manager.get_backup_enabled()

        def rollback_thread():
            try:
                modifier = get_pbix_modifier()
                result = modifier.swap_connection(file_path, server, database, create_backup=create_backup)

                def on_complete():
                    self._hide_progress()
                    if result.success:
                        self._complete_thin_report_rollback(entry, server, database, is_pbip=True)
                    else:
                        self._log_message(f"❌ Rollback failed: {result.message}")

                self.frame.after(0, on_complete)
            except Exception as e:
                self.frame.after(0, lambda: self._hide_progress())
                self.frame.after(0, lambda: self._log_message(f"❌ Rollback error: {e}"))

        threading.Thread(target=rollback_thread, daemon=True).start()

    def _show_thin_report_rollback_confirmation(self, entry: SwapHistoryEntry, file_path: str,
                                                 server: str, database: str):
        """Show confirmation dialog for PBIX thin report rollback."""
        from core.ui_base import ThemedMessageBox

        self._hide_progress()

        response = ThemedMessageBox.show(
            parent=self.frame,
            title="Rollback Thin Report",
            message=(
                f"Rolling back to: {server}\n\n"
                f"PBIX files require close/reopen. Choose an option:"
            ),
            buttons=["Save & Rollback", "Rollback (No Save)", "Cancel"],
            custom_icon="hotswap.svg"
        )

        if response == "Cancel":
            self._log_message("Rollback cancelled")
            return

        save_first = (response == "Save & Rollback")
        self._execute_thin_report_rollback_pbix(entry, file_path, server, database, save_first)

    def _execute_thin_report_rollback_pbix(self, entry: SwapHistoryEntry, file_path: str,
                                            server: str, database: str, save_first: bool):
        """Execute rollback for PBIX thin report with close/reopen."""
        self._show_progress("Rolling back PBIX file...")
        create_backup = self.preset_manager.get_backup_enabled()

        def rollback_thread():
            try:
                from tools.connection_hotswap.logic.process_control import get_controller

                modifier = get_pbix_modifier()
                controller = get_controller()
                process_id = self._thin_report_context.get('process_id')

                def modify_file(fp):
                    result = modifier.swap_connection(fp, server, database, create_backup=create_backup)
                    return result.success, result.message

                result = controller.save_close_and_reopen(
                    process_id=process_id,
                    file_path=file_path,
                    save_first=save_first,
                    modify_callback=modify_file
                )

                def on_complete():
                    self._hide_progress()
                    if result.success:
                        self._complete_thin_report_rollback(entry, server, database, is_pbip=False)
                        ThemedMessageBox.showinfo(
                            self.frame.winfo_toplevel(),
                            "Rollback Complete",
                            f"Rolled back to: {server}\n\n"
                            "Power BI Desktop is reopening the file."
                        )
                    else:
                        self._log_message(f"❌ Rollback failed: {result.message}")

                self.frame.after(0, on_complete)
            except Exception as e:
                self.frame.after(0, lambda: self._hide_progress())
                self.frame.after(0, lambda: self._log_message(f"❌ Rollback error: {e}"))

        threading.Thread(target=rollback_thread, daemon=True).start()

    def _complete_thin_report_rollback(self, entry: SwapHistoryEntry, server: str,
                                        database: str, is_pbip: bool):
        """Complete thin report rollback by updating UI state."""
        # Update the mapping source to reflect rolled-back state
        for i, m in enumerate(self.mappings):
            m.source.server = server
            m.source.database = database
            m.source.is_cloud = 'powerbi://' in server.lower()
            m.status = SwapStatus.PENDING
            m.target = None
            self._update_mapping_row(i, m)

        # Remove this entry from history
        self._swap_history = [h for h in self._swap_history if h.timestamp != entry.timestamp]
        self._save_swap_history()

        # Update button state
        if not self._swap_history:
            self._last_swapped_mapping = None
            self.rollback_btn.set_enabled(False)

        if is_pbip:
            self._log_message(f"✅ Rollback successful. Close and reopen the report to apply.")
            ThemedMessageBox.showinfo(
                self.frame.winfo_toplevel(),
                "Rollback Complete",
                f"Rolled back to: {server}\n\n"
                "Close and reopen the report in Power BI Desktop to apply changes."
            )
        else:
            self._log_message(f"✅ Rollback successful")

    def _perform_batch_rollback(self, entries: List[SwapHistoryEntry]):
        """Perform rollback for multiple history entries (batch rollback)"""
        if not self.swapper or not entries:
            return

        # Find all matching mappings
        rollback_pairs = []  # list of (entry, mapping, index) tuples
        for entry in entries:
            for i, m in enumerate(self.mappings):
                if m.source.name == entry.connection_name:
                    rollback_pairs.append((entry, m, i))
                    break

        if not rollback_pairs:
            self._log_message("❌ Cannot find mappings for batch rollback")
            return

        total = len(rollback_pairs)
        self._show_progress(f"⏳ Rolling back {total} connection(s)...")
        self._log_message(f"⏳ Starting batch rollback of {total} connection(s)...")

        def batch_rollback_thread():
            success_count = 0
            timestamps_to_remove = []

            for entry, mapping, idx in rollback_pairs:
                # Restore original connection string for rollback
                mapping.original_connection_string = entry.original_connection_string

                try:
                    result = self.swapper.rollback_connection(mapping)
                    success = result.success
                    message = result.message

                    def update_ui(success=success, message=message, entry=entry, mapping=mapping, idx=idx):
                        if success:
                            # Update the mapping source to reflect rolled-back state
                            mapping.source.server = entry.original_server
                            mapping.source.database = entry.original_database
                            mapping.source.connection_string = entry.original_connection_string
                            mapping.source.is_cloud = 'powerbi://' in entry.original_server.lower()
                            mapping.status = SwapStatus.PENDING
                            mapping.target = None
                            self._update_mapping_row(idx, mapping)
                            self._log_message(f"✅ Rolled back {entry.connection_name}")
                        else:
                            self._log_message(f"❌ Failed to rollback {entry.connection_name}: {message}")

                    self.frame.after(0, update_ui)

                    if success:
                        success_count += 1
                        timestamps_to_remove.append(entry.timestamp)

                except Exception as e:
                    self.frame.after(0, lambda e=e, name=entry.connection_name:
                        self._log_message(f"❌ Rollback error for {name}: {e}"))

            def on_complete():
                self._hide_progress()

                # Remove rolled-back entries from history
                self._swap_history = [h for h in self._swap_history
                                      if h.timestamp not in timestamps_to_remove]

                # Persist updated history
                self._save_swap_history()

                # Update button state
                if not self._swap_history:
                    self._last_swapped_mapping = None
                    self.rollback_btn.set_enabled(False)

                self._log_message(f"Batch rollback complete: {success_count}/{total} succeeded")

            self.frame.after(0, on_complete)

        threading.Thread(target=batch_rollback_thread, daemon=True).start()

    # =========================================================================
    # Health Monitoring Methods
    # =========================================================================

    def _on_check_health(self):
        """Manually trigger health check for all configured targets"""
        if not self.health_checker:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Not Ready", "Connect to a model first.")
            return

        if not self.mappings:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Mappings", "No connections to check.")
            return

        # Count targets to check
        targets_to_check = sum(1 for m in self.mappings if m.target)
        if targets_to_check == 0:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Targets", "Configure targets first using Apply Target.")
            return

        # Show progress
        self._show_progress(f"Checking {targets_to_check} target(s)...")
        self._log_message("Checking connection health...")

        def check_thread():
            try:
                # Ensure all targets are added
                for mapping in self.mappings:
                    if mapping.target:
                        self.health_checker.add_target(mapping.target)

                # Check all now
                results = self.health_checker.check_all_now()

                def on_complete():
                    self._hide_progress()

                    # Count statuses
                    healthy = sum(1 for r in results.values() if r.status == HealthStatus.HEALTHY)
                    unhealthy = sum(1 for r in results.values() if r.status == HealthStatus.UNHEALTHY)
                    errors = sum(1 for r in results.values() if r.status == HealthStatus.ERROR)
                    total = len(results)

                    # Update UI for all mappings
                    self._update_health_display()

                    # Show summary message
                    if total == 0:
                        self._log_message("No targets were checked.")
                    elif healthy == total:
                        self._log_message(f"All {healthy} target(s) healthy")
                    else:
                        status_parts = []
                        if healthy > 0:
                            status_parts.append(f"{healthy} healthy")
                        if unhealthy > 0:
                            status_parts.append(f"{unhealthy} unhealthy")
                        if errors > 0:
                            status_parts.append(f"{errors} error(s)")
                        summary = ", ".join(status_parts)
                        self._log_message(f"Health check: {summary}")

                self.frame.after(0, on_complete)

            except Exception as e:
                def on_error():
                    self._hide_progress()
                    self._log_message(f"Health check error: {e}")
                self.frame.after(0, on_error)

        threading.Thread(target=check_thread, daemon=True).start()

    def _on_health_status_change(self, target_id: str, result: HealthCheckResult):
        """Callback when health status changes (called from health checker thread)"""
        # Store the result
        self._health_statuses[target_id] = result

        # Schedule UI update on main thread
        self.frame.after(0, lambda: self._update_health_display())

    def _start_health_monitoring(self):
        """Start background health monitoring for configured targets"""
        if not self.health_checker:
            return

        # Clear and add targets
        self.health_checker.clear_targets()

        for mapping in self.mappings:
            if mapping.target:
                target_id = self.health_checker.add_target(mapping.target)
                self.logger.debug(f"Added target to health checker: {target_id}")

        # Start background monitoring
        self.health_checker.start()
        self._log_message("🩺 Health monitoring started (30s interval)")

    def _update_health_display(self):
        """Update the health status display in UI"""
        # Update mapping table with health indicators
        for i, mapping in enumerate(self.mappings):
            if mapping.target:
                target_id = f"{mapping.target.server}|{mapping.target.database}"
                result = self._health_statuses.get(target_id)

                if result:
                    # Update status column with health emoji
                    status = mapping.status.value
                    status_emoji = self._get_status_emoji_with_health(mapping.status, result.status)
                    # Determine target type (Local/Cloud)
                    target_type = "Cloud" if mapping.target.target_type == "cloud" else "Local"
                    self.mapping_tree.item(
                        str(i),
                        values=(mapping.source.display_name,
                                mapping.target.display_name,
                                target_type,
                                status_emoji)
                    )

        # Health status is now shown in the table's status column

    def _get_status_emoji_with_health(self, swap_status: SwapStatus, health_status: HealthStatus) -> str:
        """Get status indicator showing swap status (health indicators removed per user request)"""
        # Return only swap status - no colored circle health indicators
        return {
            SwapStatus.PENDING: '--',
            SwapStatus.MATCHED: '?',
            SwapStatus.READY: 'Ready',
            SwapStatus.SWAPPING: '...',
            SwapStatus.SUCCESS: 'Done',
            SwapStatus.ERROR: 'ERR',
        }.get(swap_status, '--')

    # =========================================================================
    # View Mode Methods (Table view only - diagram toggle removed)
    # =========================================================================

    def _set_view_mode(self, mode: str):
        """
        Switch between table and diagram view modes.
        Note: Diagram toggle UI removed, but method kept for compatibility.

        Args:
            mode: "table" or "diagram"
        """
        if mode == self._view_mode.get():
            return  # Already in this mode

        self._view_mode.set(mode)

        if mode == "table":
            # Show table, hide diagram
            self._diagram_container.pack_forget()
            self._tree_container.pack(fill=tk.BOTH, expand=True)
            self._log_message("Switched to table view")

        else:  # diagram
            # Show diagram, hide table
            self._tree_container.pack_forget()
            self._diagram_container.pack(fill=tk.BOTH, expand=True)

            # Update diagram with current mappings
            self._connection_diagram.update_mappings(self.mappings)

            # Sync selection
            selection = self.mapping_tree.selection()
            if selection:
                idx = int(selection[0])
                self._connection_diagram.set_selected(idx)

            self._log_message("Switched to diagram view")

    def _on_diagram_node_click(self, index: int):
        """Handle click on a diagram node."""
        if index < 0 or index >= len(self.mappings):
            return

        # Select the corresponding row in the tree (for state sync)
        self.mapping_tree.selection_set(str(index))

        # Trigger the selection event handler
        self._on_mapping_selected()

    def _update_diagram(self):
        """Update the diagram with current mappings."""
        if hasattr(self, '_connection_diagram') and self._view_mode.get() == "diagram":
            self._connection_diagram.update_mappings(self.mappings)

    # =========================================================================
    # Swap History Persistence
    # =========================================================================

    def _get_swap_history_file_path(self) -> Path:
        """Get the path to the swap history persistence file."""
        from pathlib import Path
        import os

        # Store in same location as presets: %APPDATA%/Analytic Endeavors/PBI Report Merger/
        base = Path(os.environ.get('APPDATA', ''))
        history_dir = base / 'Analytic Endeavors' / 'PBI Report Merger' / 'hotswap_history'
        history_dir.mkdir(parents=True, exist_ok=True)
        return history_dir / 'swap_history.json'

    def _load_swap_history(self) -> None:
        """Load swap history from persistence file."""
        import json

        if not self._swap_history_file.exists():
            return

        try:
            with open(self._swap_history_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            entries = data.get('history', [])
            self._swap_history = [
                SwapHistoryEntry.from_dict(entry)
                for entry in entries
            ]
            self.logger.debug(f"Loaded {len(self._swap_history)} swap history entries")
        except Exception as e:
            self.logger.warning(f"Error loading swap history: {e}")
            self._swap_history = []

    def _save_swap_history(self) -> None:
        """Save swap history to persistence file."""
        import json

        try:
            data = {
                'version': 1,
                'history': [entry.to_dict() for entry in self._swap_history]
            }

            with open(self._swap_history_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)

            self.logger.debug(f"Saved {len(self._swap_history)} swap history entries")
        except Exception as e:
            self.logger.warning(f"Error saving swap history: {e}")

    def _generate_run_id(self) -> str:
        """Generate a unique run ID for grouping batch swaps."""
        import datetime
        import uuid
        # Use timestamp + short UUID for uniqueness
        return f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"

    # =========================================================================
    # Preset Methods
    # =========================================================================

    def _get_model_hash(self) -> Optional[str]:
        """
        Get a hash of the current connected model for preset keying.

        Uses file path when available for stable hashing across connection changes.
        This ensures presets persist for a file regardless of its current connection
        (local, cloud, or after any swap).
        """
        import hashlib

        # PRIORITY 1: Use file path for stable hashing (works across connection changes)
        # Check thin report context first
        if self._thin_report_context:
            file_path = self._thin_report_context.get('file_path', '')
            if file_path:
                normalized_path = os.path.normpath(file_path).lower()
                identifier = f"file:{normalized_path}"
                hash_result = hashlib.md5(identifier.encode()).hexdigest()[:12]
                self.logger.debug(f"Model hash from thin_report file_path: {hash_result} (path: {normalized_path})")
                return hash_result

        # Check stored file path for regular models
        if self._current_model_file_path:
            normalized_path = os.path.normpath(self._current_model_file_path).lower()
            identifier = f"file:{normalized_path}"
            hash_result = hashlib.md5(identifier.encode()).hexdigest()[:12]
            self.logger.debug(f"Model hash from _current_model_file_path: {hash_result} (path: {normalized_path})")
            return hash_result

        # PRIORITY 2: Fallback to connection info if no file path available
        if self._thin_report_context:
            cloud_server = self._thin_report_context.get('cloud_server', '')
            cloud_database = self._thin_report_context.get('cloud_database', '')
            if cloud_server or cloud_database:
                identifier = f"thin:{cloud_server}|{cloud_database}".lower()
                hash_result = hashlib.md5(identifier.encode()).hexdigest()[:12]
                self.logger.debug(f"Model hash from thin_report cloud info: {hash_result}")
                return hash_result

        if not hasattr(self, 'connector') or not self.connector:
            self.logger.debug("Model hash: None (no connector)")
            return None

        conn = self.connector.current_connection
        if not conn or not conn.is_connected:
            self.logger.debug("Model hash: None (not connected)")
            return None

        # Use MD5 of server + database for a stable, short hash
        identifier = f"{conn.server}|{conn.database}".lower()
        hash_result = hashlib.md5(identifier.encode()).hexdigest()[:12]
        self.logger.debug(f"Model hash from connector: {hash_result}")
        return hash_result

    def _get_model_display_name(self) -> Optional[str]:
        """Get a display name for the current connected model."""
        # Check for thin report first
        if self._thin_report_context:
            return self._thin_report_context.get('report_name') or "Thin Report"

        if not hasattr(self, 'connector') or not self.connector:
            return None

        conn = self.connector.current_connection
        if not conn or not conn.is_connected:
            return None

        return conn.model_name

    def _get_friendly_name(self, database: str, server: str) -> str:
        """
        Convert database/server info to a friendly display name.

        For local models: "Local (port)" or "Local"
        For cloud models: Try to find matching target display_name, else truncate GUID

        Args:
            database: Database name or GUID
            server: Server address

        Returns:
            Friendly display name for the connection
        """
        if not database and not server:
            return "(unknown)"

        # Check if it's a local model (localhost)
        server_lower = (server or "").lower()
        if server_lower.startswith("localhost"):
            # Extract port if present
            if ":" in server:
                port = server.split(":")[-1]
                return f"Local ({port})"
            return "Local"

        # Check if it's a cloud model (powerbi.com)
        if "powerbi.com" in server_lower or "pbidedicated" in server_lower:
            # Try to find a matching target in current mappings
            for mapping in getattr(self, 'mappings', []):
                if mapping.target:
                    target_db = getattr(mapping.target, 'database', '')
                    if target_db and target_db.lower() == (database or "").lower():
                        return mapping.target.display_name

            # Try the _display_to_connection reverse lookup
            for display_name, conn_str in getattr(self, '_display_to_connection', {}).items():
                if database and database.lower() in conn_str.lower():
                    return display_name

            # Fall back to truncated GUID (first 8 chars)
            if database and len(database) > 10:
                return f"Cloud ({database[:8]}...)"
            return database or "(cloud)"

        # Default: truncate if too long
        if database and len(database) > 20:
            return f"{database[:17]}..."
        return database or "(unknown)"

    def _get_connection_type(self, server: str) -> str:
        """
        Determine connection type from server string.

        Returns:
            'Local', 'Cloud', or 'XMLA'
        """
        if not server:
            return "Unknown"

        server_lower = server.lower()

        # Local: localhost:port
        if server_lower.startswith("localhost"):
            return "Local"

        # XMLA endpoint: powerbi://api.powerbi.com/v1.0/myorg/...
        if "powerbi://" in server_lower and "/v1.0/" in server_lower:
            return "XMLA"

        # Cloud PBI Service: pbiazure://api.powerbi.com or powerbi:// without /v1.0/
        if "pbiazure://" in server_lower or "powerbi://" in server_lower:
            return "Cloud"

        return "Unknown"

    def _get_swap_type(
        self,
        source_server: str,
        target_server: str,
        stored_source_type: str = "",
        stored_target_type: str = ""
    ) -> str:
        """
        Get human-readable swap type from source and target.

        Args:
            source_server: Source server address (fallback for older entries)
            target_server: Target server address (fallback for older entries)
            stored_source_type: Stored source type from SwapHistoryEntry (preferred)
            stored_target_type: Stored target type from SwapHistoryEntry (preferred)

        Returns format like: "Local to Cloud", "XMLA to Local", etc.
        """
        # Use stored types if available, otherwise infer from server
        source_type = stored_source_type if stored_source_type else self._get_connection_type(source_server)
        target_type = stored_target_type if stored_target_type else self._get_connection_type(target_server)
        return f"{source_type} to {target_type}"

    def _load_model_presets_on_connect(self):
        """
        Check for model-specific presets after connecting and auto-switch to MODEL scope if found.

        This ensures that when a user connects to a model, any saved presets for that model
        are immediately visible in the preset list.
        """
        try:
            model_hash = self._get_model_hash()
            if not model_hash:
                # No model hash available, just refresh current view
                self._refresh_preset_table()
                return

            # Check if there are model-specific presets for this model
            model_presets = self.preset_manager.list_presets(
                scope=PresetScope.MODEL,
                model_hash=model_hash
            )

            if model_presets:
                # Model-specific presets found - switch to MODEL scope
                self._preset_scope_filter = PresetScope.MODEL
                self._update_preset_scope_toggle()
                self._refresh_preset_table()
                self._log_message(f"Loaded {len(model_presets)} model-specific preset(s)")
            else:
                # No model presets, stay on current scope but refresh to update UI
                self._refresh_preset_table()

        except Exception as e:
            self.logger.warning(f"Error loading model presets on connect: {e}")
            # Still refresh to ensure table is in sync
            self._refresh_preset_table()

    def _refresh_preset_buttons(self):
        """Legacy method - redirects to _refresh_preset_table for backwards compatibility"""
        self._refresh_preset_table()

    def _refresh_preset_table(self):
        """Refresh the preset table from saved presets based on current scope filter."""
        # Clear existing items
        if hasattr(self, '_preset_tree'):
            self._preset_tree.delete(*self._preset_tree.get_children())

        # Get current model hash for MODEL scope filtering
        model_hash = None
        if self._preset_scope_filter == PresetScope.MODEL:
            # Get model hash from connected model if available
            model_hash = self._get_model_hash()
            self.logger.info(f"Loading MODEL presets with hash: {model_hash}")

        # Load presets filtered by scope
        self._current_presets = self.preset_manager.list_presets(
            scope=self._preset_scope_filter,
            model_hash=model_hash
        )

        if not self._current_presets:
            # Still update button counts even if no presets for current scope
            self._update_preset_scope_button_counts()
            return

        # Populate table with presets
        for preset in self._current_presets:
            mapping_count = len(preset.mappings)

            # Determine preset type based on mappings
            # Check if any mapping targets cloud connections
            preset_type = "Local"
            for mapping in preset.mappings:
                # PresetTargetMapping has target_type attribute directly
                if mapping.target_type == "cloud":
                    preset_type = "Cloud"
                    break
                # Also check server string for cloud indicators
                if mapping.server:
                    server_lower = mapping.server.lower()
                    if 'powerbi://' in server_lower or 'asazure://' in server_lower:
                        preset_type = "Cloud"
                        break

            self._preset_tree.insert(
                "",
                "end",
                iid=preset.name,
                values=(preset.name, preset_type, mapping_count)
            )

        self.logger.info(f"Loaded {len(self._current_presets)} {self._preset_scope_filter.value} preset(s)")

        # Update button counts to reflect all presets
        self._update_preset_scope_button_counts()

    def _on_preset_tree_select(self, event):
        """Handle preset tree selection - enable/disable Apply button"""
        selection = self._preset_tree.selection()
        has_selection = len(selection) > 0

        # Enable Apply button only if a preset is selected AND we have mappings
        if hasattr(self, 'apply_preset_btn'):
            self.apply_preset_btn.set_enabled(has_selection and bool(self.mappings))

    def _on_apply_preset_to_mappings(self):
        """Apply selected preset to mapping table (preview mode - does not swap)"""
        selection = self._preset_tree.selection()
        if not selection:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Selection", "Please select a preset first.")
            return

        preset_name = selection[0]

        # Find the preset
        preset = None
        for p in self._current_presets:
            if p.name == preset_name:
                preset = p
                break

        if preset:
            # Use the existing _on_preset_click which applies targets without swapping
            self._on_preset_click(preset)

    def _on_preset_tree_double_click(self, event):
        """Handle double-click on preset table for Quick Apply (apply + swap immediately)"""
        selection = self._preset_tree.selection()
        if not selection:
            return

        preset_name = selection[0]

        # Find the preset by name
        preset = None
        for p in self._current_presets:
            if p.name == preset_name:
                preset = p
                break

        if preset:
            self._quick_apply_preset(preset)

    def _quick_apply_preset(self, preset: SwapPreset):
        """Quick Apply: apply preset targets and immediately execute swaps"""
        if not self.mappings:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Mappings", "Connect to a model first to apply presets.")
            return

        # First apply the preset to set targets (same as _on_preset_click logic)
        if preset.scope == PresetScope.GLOBAL:
            updated, message = self.preset_manager.apply_global_preset(preset, self.mappings)
            if updated == 0 and message != "OK":
                if "same as current" in message.lower():
                    self._log_message(f"Global preset '{preset.name}' skipped: target is same as current connection")
                elif "single-connection" in message.lower():
                    ThemedMessageBox.showinfo(
                        self.frame.winfo_toplevel(),
                        "Cannot Apply",
                        f"Global presets only work with single-connection models.\n"
                        f"This model has {len(self.mappings)} connections."
                    )
                else:
                    self._log_message(f"Could not apply global preset: {message}")
                return
        else:
            updated = self.preset_manager.apply_preset_to_mappings(preset, self.mappings)

        if updated == 0:
            self._log_message(f"Preset '{preset.name}' did not match any connections")
            return

        # Refresh mapping table to show new targets
        for i, mapping in enumerate(self.mappings):
            self._update_mapping_row(i, mapping)

        scope_label = "global" if preset.scope == PresetScope.GLOBAL else "model"
        self._log_message(f"Quick Apply: preset '{preset.name}' ({scope_label}) - {updated} mapping(s) configured")

        # Now execute swaps on all ready mappings
        ready_mappings = [(i, m) for i, m in enumerate(self.mappings) if m.is_ready]
        if not ready_mappings:
            self._log_message("No mappings ready to swap after applying preset")
            return

        # Confirm the quick swap
        if ThemedMessageBox.askyesno(
            self.frame.winfo_toplevel(),
            "Quick Apply",
            f"Apply preset '{preset.name}' and swap {len(ready_mappings)} connection(s) now?"
        ):
            # Execute batch swap sequentially (to avoid XmlReader conflict)
            self._execute_batch_swap(ready_mappings)

    def _on_preset_tree_right_click(self, event):
        """Handle right-click on preset table to show context menu"""
        # Identify the row under the cursor
        row_id = self._preset_tree.identify_row(event.y)
        if not row_id:
            return

        # Select the row
        self._preset_tree.selection_set(row_id)

        # Find the preset
        preset = None
        for p in self._current_presets:
            if p.name == row_id:
                preset = p
                break

        if not preset:
            return

        # Show modern styled context popup
        self._show_preset_context_popup(event.x_root, event.y_root, preset)

    def _show_preset_context_popup(self, x: int, y: int, preset: SwapPreset):
        """Show a modern styled context menu popup for preset."""
        # Close any existing popup
        self._close_preset_context_popup()

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create popup window
        self._preset_context_popup = tk.Toplevel(self._preset_tree)
        self._preset_context_popup.withdraw()
        self._preset_context_popup.overrideredirect(True)

        # Popup colors - themed
        popup_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        # Use themed border color instead of white
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
        text_color = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        hover_bg = colors.get('hover', '#2a2a3e' if is_dark else '#f0f0f5')

        # Border frame with themed color
        border_frame = tk.Frame(self._preset_context_popup, bg=border_color, padx=1, pady=1)
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Main content frame
        main_frame = tk.Frame(border_frame, bg=popup_bg)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Menu items with proper padding
        menu_items = [
            ("Quick Apply (swap now)", lambda: self._quick_apply_preset(preset)),
            ("Apply to Mappings (preview)", lambda: self._on_preset_click(preset)),
            None,  # Separator
            ("View Details", lambda: self._show_preset_details_dialog(preset)),
            None,  # Separator
            ("Delete", lambda: self._delete_preset_from_context(preset)),
        ]

        for item in menu_items:
            if item is None:
                # Separator
                sep = tk.Frame(main_frame, height=1, bg=border_color)
                sep.pack(fill=tk.X, padx=8, pady=4)
            else:
                label_text, command = item
                item_frame = tk.Frame(main_frame, bg=popup_bg)
                item_frame.pack(fill=tk.X)

                label = tk.Label(
                    item_frame,
                    text=label_text,
                    font=('Segoe UI', 9),
                    fg=text_color,
                    bg=popup_bg,
                    anchor='w',
                    padx=16,  # Good horizontal padding
                    pady=8,   # Good vertical padding
                    cursor='hand2'
                )
                label.pack(fill=tk.X)

                # Hover effects
                def on_enter(e, f=item_frame, l=label, hbg=hover_bg):
                    f.configure(bg=hbg)
                    l.configure(bg=hbg)

                def on_leave(e, f=item_frame, l=label, bg=popup_bg):
                    f.configure(bg=bg)
                    l.configure(bg=bg)

                def on_click(e, cmd=command):
                    self._close_preset_context_popup()
                    cmd()

                for widget in [item_frame, label]:
                    widget.bind('<Enter>', on_enter)
                    widget.bind('<Leave>', on_leave)
                    widget.bind('<Button-1>', on_click)

        # Position and show popup
        self._preset_context_popup.update_idletasks()
        self._preset_context_popup.geometry(f"+{x}+{y}")
        self._preset_context_popup.deiconify()

        # Close on click outside or focus loss
        self._preset_context_popup.bind('<FocusOut>', lambda e: self._close_preset_context_popup())
        self._preset_tree.bind('<Button-1>', lambda e: self._close_preset_context_popup(), add='+')
        self.frame.winfo_toplevel().bind('<Button-1>', lambda e: self._close_preset_context_popup(), add='+')

    def _close_preset_context_popup(self):
        """Close the preset context popup if open."""
        if hasattr(self, '_preset_context_popup') and self._preset_context_popup:
            try:
                self._preset_context_popup.destroy()
            except Exception:
                pass
            self._preset_context_popup = None

    def _delete_preset_from_context(self, preset: SwapPreset):
        """Delete a preset from the context menu"""
        if ThemedMessageBox.askyesno(
            self.frame.winfo_toplevel(),
            "Delete Preset",
            f"Delete preset '{preset.name}'?\n\nThis cannot be undone."
        ):
            if self.preset_manager.delete_preset(
                preset.name,
                preset.storage_type,
                preset.scope,
                preset.model_hash
            ):
                self._log_message(f"Deleted preset '{preset.name}'")
                self._refresh_preset_table()
            else:
                ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Delete Failed", "Could not delete preset.")

    def _on_export_selected_preset(self):
        """Export all global presets to a JSON file"""
        # Check if there are any global presets to export
        preset_count = self.preset_manager.get_global_preset_count()
        if preset_count == 0:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Presets", "No global presets to export.")
            return

        file_path = filedialog.asksaveasfilename(
            parent=self.frame.winfo_toplevel(),
            title="Export Global Presets",
            initialfile="hotswap_global_presets.json",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not file_path:
            return  # User cancelled

        if self.preset_manager.export_all_global_presets(file_path):
            self._log_message(f"Exported {preset_count} global preset(s) to {file_path}")
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Export Successful", f"Exported {preset_count} global preset(s) to:\n{file_path}")
        else:
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Export Failed", "Could not export presets to file.")

    def _on_import_preset(self):
        """Import global presets from a JSON file (replaces existing)"""
        file_path = filedialog.askopenfilename(
            parent=self.frame.winfo_toplevel(),
            title="Import Global Presets",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not file_path:
            return  # User cancelled

        # Confirm replacement
        current_count = self.preset_manager.get_global_preset_count()
        if current_count > 0:
            if not ThemedMessageBox.askyesno(
                self.frame.winfo_toplevel(),
                "Replace Global Presets?",
                f"This will replace your {current_count} existing global preset(s).\n\nContinue?"
            ):
                return

        # Import the presets
        success, count, message = self.preset_manager.import_all_global_presets(file_path)

        if success:
            self._log_message(f"Imported {count} global preset(s) from {file_path}")
            self._refresh_preset_table()
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "Import Successful", message)
        else:
            ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Import Failed", message)

    def _show_preset_details_dialog(self, preset: SwapPreset):
        """Show a dialog with preset details"""
        import os
        import sys
        import ctypes
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        parent = self.frame.winfo_toplevel()

        dialog = tk.Toplevel(parent)
        dialog.title(f"Preset Details: {preset.name}")
        dialog.geometry("450x415")
        dialog.resizable(True, True)
        dialog.transient(parent)
        dialog.configure(bg=colors['background'])

        # Set dialog icon (AE favicon)
        try:
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))),
                                     'assets', 'favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except Exception:
            pass  # Ignore if icon can't be loaded

        # Apply Windows dark/light mode title bar
        if sys.platform == 'win32':
            try:
                hwnd = ctypes.windll.user32.GetParent(dialog.winfo_id())
                DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                value = ctypes.c_int(1 if is_dark else 0)
                result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                    ctypes.byref(value), ctypes.sizeof(value)
                )
                if result != 0:
                    # Try older attribute for pre-20H1 Windows
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, 19, ctypes.byref(value), ctypes.sizeof(value)
                    )
            except Exception:
                pass  # Silently fail on unsupported systems

        # Center on parent
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 415) // 2
        dialog.geometry(f"+{x}+{y}")

        # Content
        frame = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        frame.pack(fill=tk.BOTH, expand=True)

        # Header row with icon and title
        header_frame = tk.Frame(frame, bg=colors['background'])
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # Load AE icon for header (use theme-aware color)
        ae_icon = self._load_icon_for_button('analyze', size=20)
        if ae_icon:
            icon_label = tk.Label(header_frame, image=ae_icon, bg=colors['background'])
            icon_label.pack(side=tk.LEFT, padx=(0, 10))
            icon_label._icon_ref = ae_icon  # Keep reference

        # Title
        tk.Label(
            header_frame,
            text=preset.name,
            font=("Segoe UI", 12, "bold"),
            fg=colors['title_color'],
            bg=colors['background']
        ).pack(side=tk.LEFT)

        # Info section
        info_frame = tk.Frame(frame, bg=colors['background'])
        info_frame.pack(fill=tk.X, pady=(0, 15))

        # Scope
        scope_label = "Global" if preset.scope == PresetScope.GLOBAL else "Model-specific"
        tk.Label(
            info_frame,
            text=f"Scope: {scope_label}",
            font=("Segoe UI", 9),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(anchor=tk.W)

        # Model name (if model-specific)
        if preset.scope == PresetScope.MODEL and preset.model_name:
            tk.Label(
                info_frame,
                text=f"Model: {preset.model_name}",
                font=("Segoe UI", 9),
                fg=colors['text_primary'],
                bg=colors['background']
            ).pack(anchor=tk.W)

        # Storage location
        storage_label = "Personal (AppData)" if preset.storage_type == PresetStorageType.USER else \
                       "Project Folder" if preset.storage_type == PresetStorageType.PROJECT else \
                       "With Report (PBIP)"
        tk.Label(
            info_frame,
            text=f"Storage: {storage_label}",
            font=("Segoe UI", 9),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(anchor=tk.W)

        # Updated date
        if preset.updated_at:
            date_str = preset.updated_at[:10] if len(preset.updated_at) >= 10 else preset.updated_at
            tk.Label(
                info_frame,
                text=f"Updated: {date_str}",
                font=("Segoe UI", 9),
                fg=colors.get('text_secondary', '#888888'),
                bg=colors['background']
            ).pack(anchor=tk.W)

        # Mappings section
        tk.Label(
            frame,
            text=f"Connection Mappings ({len(preset.mappings)}):",
            font=("Segoe UI", 9, "bold"),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(anchor=tk.W, pady=(0, 5))

        # Mappings list with scrollbar in bordered container (visible scroll bar)
        list_frame = tk.Frame(frame, bg=colors['background'])
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Container with visible border (like connections table)
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
        text_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')

        container = tk.Frame(
            list_frame,
            bg=text_bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=1
        )
        container.pack(fill=tk.BOTH, expand=True)

        # Create Text widget for mappings
        mappings_text = tk.Text(
            container,
            bg=text_bg,
            fg=colors['text_primary'],
            font=('Consolas', 9),
            wrap=tk.WORD,
            height=10,
            relief=tk.FLAT,
            padx=10,
            pady=10
        )

        # Scrollbar area with visible background
        scrollbar_bg = '#1a1a2e' if is_dark else '#f0f0f0'
        scrollbar_area = tk.Frame(container, bg=scrollbar_bg, width=12)
        scrollbar_area.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_area.pack_propagate(False)  # Maintain fixed width

        scrollbar = ThemedScrollbar(
            scrollbar_area,
            command=mappings_text.yview,
            theme_manager=self._theme_manager
        )
        mappings_text.configure(yscrollcommand=scrollbar.set)

        mappings_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(fill=tk.BOTH, expand=True)

        # Populate mappings
        for i, mapping in enumerate(preset.mappings):
            # Connection name
            conn_name = mapping.connection_name or f"Connection {i+1}"
            mappings_text.insert(tk.END, f"{conn_name}\n", "name")

            # Target info - show friendly details for cloud connections
            target_type = mapping.target_type or "local"
            mappings_text.insert(tk.END, f"  Type: {target_type.capitalize()}\n")

            if target_type == "cloud":
                # For cloud connections, show workspace and friendly name
                workspace = mapping.workspace_name or "N/A"
                # Use display_name if available, otherwise fallback to database
                model_name = mapping.display_name or mapping.database or "N/A"
                # If display_name includes workspace (e.g., "Model (Workspace)"), just use database
                if mapping.display_name and mapping.workspace_name and mapping.workspace_name in mapping.display_name:
                    model_name = mapping.database or "N/A"
                mappings_text.insert(tk.END, f"  Workspace: {workspace}\n")
                mappings_text.insert(tk.END, f"  Model: {model_name}\n")
                # Show cloud connection type (Semantic Model vs XMLA)
                conn_type = mapping.cloud_connection_type
                if conn_type:
                    conn_type_display = "Semantic Model" if conn_type == "pbi_semantic_model" else \
                                       "XMLA Endpoint" if conn_type == "aas_xmla" else conn_type
                    mappings_text.insert(tk.END, f"  Connector: {conn_type_display}\n")
                # Show perspective if one is set
                if mapping.perspective_name:
                    mappings_text.insert(tk.END, f"  Perspective: {mapping.perspective_name}\n")
            else:
                # For local connections, show server and database
                server = mapping.server or "N/A"
                database = mapping.database or "N/A"
                mappings_text.insert(tk.END, f"  Server: {server}\n")
                mappings_text.insert(tk.END, f"  Database: {database}\n")

            if i < len(preset.mappings) - 1:
                mappings_text.insert(tk.END, "\n")

        # Configure tag for connection names
        mappings_text.tag_configure("name", font=('Consolas', 9, 'bold'))

        # Make read-only
        mappings_text.configure(state=tk.DISABLED)

        # Button area
        button_area = tk.Frame(frame, bg=colors['background'])
        button_area.pack(fill=tk.X, pady=(5, 0))

        # OK button - auto-sized
        canvas_bg = colors['background']
        close_btn = RoundedButton(
            button_area, text="OK",
            command=dialog.destroy,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors.get('button_primary_pressed', colors['button_primary_hover']),
            fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 10, 'bold'),
            canvas_bg=canvas_bg
        )
        close_btn.pack(anchor=tk.E)

    def _on_preset_click(self, preset: SwapPreset):
        """Apply a preset to current mappings"""
        if not self.mappings:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Mappings", "Connect to a model first to apply presets.")
            return

        # Handle global vs model presets differently
        if preset.scope == PresetScope.GLOBAL:
            # Global presets use special application logic with same-connection guard
            updated, message = self.preset_manager.apply_global_preset(preset, self.mappings)

            if updated == 0 and message != "OK":
                # Show reason why global preset couldn't be applied
                if "same as current" in message.lower():
                    self._log_message(f"Global preset '{preset.name}' skipped: target is same as current connection")
                elif "single-connection" in message.lower():
                    ThemedMessageBox.showinfo(
                        self.frame.winfo_toplevel(),
                        "Cannot Apply",
                        f"Global presets only work with single-connection models.\n"
                        f"This model has {len(self.mappings)} connections."
                    )
                else:
                    self._log_message(f"Could not apply global preset: {message}")
                return
        else:
            # Model presets use standard application logic
            updated = self.preset_manager.apply_preset_to_mappings(preset, self.mappings)

        if updated > 0:
            # Refresh the mapping table
            for i, mapping in enumerate(self.mappings):
                self._update_mapping_row(i, mapping)

            scope_label = "global" if preset.scope == PresetScope.GLOBAL else "model"
            self._log_message(f"Applied {scope_label} preset '{preset.name}': {updated} mapping(s) configured")

            # Enable save button, swap button requires selection
            if hasattr(self, 'save_mapping_btn'):
                self.save_mapping_btn.set_enabled(True)
            self._on_mapping_selected()  # Check selection before enabling swap button
        else:
            self._log_message(f"Preset '{preset.name}' did not match any connections")

    def _on_save_preset(self):
        """Save current mappings as a new preset"""
        if not self.mappings:
            ThemedMessageBox.showinfo(self.frame.winfo_toplevel(), "No Mappings", "Connect to a model first to save presets.")
            return

        # Show save dialog - allow saving source or target connections
        # Target validation happens in the dialog based on selection
        self._show_save_preset_dialog()

    def _show_save_preset_dialog(self):
        """Show dialog to save a new preset"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        parent = self.frame.winfo_toplevel()

        dialog_width = 520
        dialog_height = 390  # Height with source/target selection and descriptions

        dialog = tk.Toplevel(parent)
        dialog.title("Save Preset")
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

        # Set dialog icon (AE favicon)
        try:
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))),
                                     'assets', 'favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except Exception:
            pass  # Ignore if icon can't be loaded

        # Set dark/light title bar on Windows (must be before deiconify)
        try:
            import ctypes
            dialog.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(dialog.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if is_dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass

        # Center on parent
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - dialog_width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - dialog_height) // 2
        dialog.geometry(f"+{x}+{y}")

        # Check if global presets are allowed (single connection only)
        can_save_global = len(self.mappings) == 1

        # Load radio icons
        radio_on_icon = self._load_icon_for_button('radio-on', size=16)
        radio_off_icon = self._load_icon_for_button('radio-off', size=16)

        # Content
        frame = tk.Frame(dialog, bg=colors['background'], padx=25, pady=15)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            frame,
            text="Save current configuration as a preset:",
            font=("Segoe UI", 10),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(anchor=tk.W, pady=(0, 12))

        # Name field
        tk.Label(
            frame,
            text="Preset Name:",
            font=("Segoe UI", 9),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(anchor=tk.W)

        name_var = tk.StringVar()
        # Use tk.Entry with no highlight instead of ttk.Entry
        entry_bg = colors.get('card_surface', '#2a2a3c' if is_dark else '#ffffff')
        entry_border = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
        name_entry = tk.Entry(
            frame, textvariable=name_var, font=('Segoe UI', 10),
            bg=entry_bg, fg=colors['text_primary'],
            insertbackground=colors['text_primary'],
            highlightthickness=1, highlightbackground=entry_border,
            highlightcolor=entry_border,  # Same color on focus - no highlight change
            relief=tk.FLAT, bd=0
        )
        name_entry.pack(fill=tk.X, pady=(4, 12), ipady=4)

        # Scope selection
        tk.Label(
            frame,
            text="Preset Scope:",
            font=("Segoe UI", 9),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(anchor=tk.W)

        scope_var = tk.StringVar(value="model")
        scope_frame = tk.Frame(frame, bg=colors['background'])
        scope_frame.pack(fill=tk.X, pady=(4, 4))

        # Custom radio button helper
        def create_radio_option(parent, text, value, variable, enabled=True):
            """Create a custom radio button with SVG icon"""
            opt_frame = tk.Frame(parent, bg=colors['background'])

            icon_label = tk.Label(opt_frame, bg=colors['background'], cursor='hand2' if enabled else '')
            is_selected = variable.get() == value
            icon = radio_on_icon if is_selected else radio_off_icon
            if icon:
                icon_label.configure(image=icon)
                icon_label._icon_ref = icon
            else:
                icon_label.configure(text="(*)" if is_selected else "( )", font=('Segoe UI', 9))
            icon_label.pack(side=tk.LEFT, padx=(0, 4))

            # Text color: normal or muted if disabled
            text_fg = colors['text_primary'] if enabled else colors.get('text_muted', '#888888')
            text_label = tk.Label(
                opt_frame, text=text, font=('Segoe UI', 9),
                fg=text_fg, bg=colors['background'],
                cursor='hand2' if enabled else ''
            )
            text_label.pack(side=tk.LEFT)

            def on_click(event=None):
                if enabled:
                    variable.set(value)
                    update_all_radios()

            if enabled:
                icon_label.bind('<Button-1>', on_click)
                text_label.bind('<Button-1>', on_click)

            return opt_frame, icon_label, text_label

        # Track radio widgets for updates
        scope_radios = []

        def update_all_radios():
            """Update all radio button icons based on current values"""
            for value, icon_lbl, _ in scope_radios:
                is_selected = scope_var.get() == value
                icon = radio_on_icon if is_selected else radio_off_icon
                if icon:
                    icon_lbl.configure(image=icon)
                    icon_lbl._icon_ref = icon

        # Side-by-side layout: Universal on left, Model-specific on right
        options_row = tk.Frame(scope_frame, bg=colors['background'])
        options_row.pack(anchor=tk.W, pady=(0, 4))

        # Universal radio - displayed on left
        universal_frame = tk.Frame(options_row, bg=colors['background'])
        universal_frame.pack(side=tk.LEFT, padx=(0, 30))

        universal_opt, universal_icon, universal_text = create_radio_option(
            universal_frame, "Universal", "global", scope_var, enabled=can_save_global
        )
        universal_opt.pack(anchor=tk.W)
        scope_radios.append(("global", universal_icon, universal_text))

        # Add tooltip to Universal option
        from core.ui_base import Tooltip
        universal_tooltip_text = (
            "Universal presets can be applied to any model with a single connection.\n"
            "Use for environment switching like 'Production', 'Development', etc."
        )
        if universal_text:
            Tooltip(universal_text, universal_tooltip_text, delay=400)
        Tooltip(universal_icon, universal_tooltip_text, delay=400)

        # Note for universal presets (only show if cannot save universal)
        if not can_save_global:
            note_label = tk.Label(
                universal_frame,
                text="(Requires single connection)",
                font=("Segoe UI", 8, "italic"),
                fg=colors.get('text_muted', '#888888'),
                bg=colors['background']
            )
            note_label.pack(anchor=tk.W, padx=(20, 0))

        # Model-specific radio - displayed on right
        model_frame = tk.Frame(options_row, bg=colors['background'])
        model_frame.pack(side=tk.LEFT)

        model_opt, model_icon, model_text = create_radio_option(
            model_frame, "Model-specific", "model", scope_var, enabled=True
        )
        model_opt.pack(anchor=tk.W)
        scope_radios.append(("model", model_icon, model_text))

        # Add tooltip to Model-specific option
        model_tooltip_text = (
            "Model-specific presets only work with this exact model file.\n"
            "Maps each connection by name to its saved target.\n"
            "Ideal for composite models: 'All Local', 'All Dev', 'All Prod', etc."
        )
        if model_text:
            Tooltip(model_text, model_tooltip_text, delay=400)
        Tooltip(model_icon, model_tooltip_text, delay=400)

        # Scope description text
        scope_desc_frame = tk.Frame(frame, bg=colors['background'])
        scope_desc_frame.pack(fill=tk.X, pady=(2, 8))

        scope_desc = tk.Label(
            scope_desc_frame,
            text="Universal: Apply to any single-connection model. Model-specific: Maps connections by name to saved targets.",
            font=("Segoe UI", 8, "italic"),
            fg=colors.get('text_muted', '#888888'),
            bg=colors['background'],
            wraplength=470,
            justify=tk.LEFT
        )
        scope_desc.pack(anchor=tk.W)

        # Connections to Save section (only for Universal presets)
        conn_section_label = tk.Label(
            frame,
            text="Connections to Save:",
            font=("Segoe UI", 9),
            fg=colors['text_primary'],
            bg=colors['background']
        )
        conn_section_label.pack(anchor=tk.W, pady=(10, 0))

        conn_source_var = tk.StringVar(value="target")  # Default to target connections
        conn_source_frame = tk.Frame(frame, bg=colors['background'])
        conn_source_frame.pack(fill=tk.X, pady=(4, 4))

        # Track connection source radios
        conn_source_radios = []

        def update_conn_source_radios():
            """Update connection source radio button icons"""
            for value, icon_lbl, _ in conn_source_radios:
                is_selected = conn_source_var.get() == value
                icon = radio_on_icon if is_selected else radio_off_icon
                if icon:
                    icon_lbl.configure(image=icon)
                    icon_lbl._icon_ref = icon

        # Side-by-side layout for connection source
        conn_options_row = tk.Frame(conn_source_frame, bg=colors['background'])
        conn_options_row.pack(anchor=tk.W, pady=(0, 4))

        # Source Connections radio
        source_conn_frame = tk.Frame(conn_options_row, bg=colors['background'])
        source_conn_frame.pack(side=tk.LEFT, padx=(0, 30))

        source_conn_opt = tk.Frame(source_conn_frame, bg=colors['background'])
        source_icon_label = tk.Label(source_conn_opt, bg=colors['background'], cursor='hand2')
        source_icon = radio_off_icon  # Not selected by default
        if source_icon:
            source_icon_label.configure(image=source_icon)
            source_icon_label._icon_ref = source_icon
        source_icon_label.pack(side=tk.LEFT, padx=(0, 4))
        source_text_label = tk.Label(
            source_conn_opt, text="Source Connections", font=('Segoe UI', 9),
            fg=colors['text_primary'], bg=colors['background'], cursor='hand2'
        )
        source_text_label.pack(side=tk.LEFT)
        source_conn_opt.pack(anchor=tk.W)
        conn_source_radios.append(("source", source_icon_label, source_text_label))

        def on_source_click(event=None):
            if scope_var.get() == "global":  # Only respond if Universal is selected
                conn_source_var.set("source")
                update_conn_source_radios()

        source_icon_label.bind('<Button-1>', on_source_click)
        source_text_label.bind('<Button-1>', on_source_click)

        # Add tooltip to Source Connections
        source_tooltip_text = (
            "Save the current source connections as the preset target.\n"
            "Use this to create a preset that points back to the current connection."
        )
        Tooltip(source_text_label, source_tooltip_text, delay=400)
        Tooltip(source_icon_label, source_tooltip_text, delay=400)

        # Target Connections radio
        target_conn_frame = tk.Frame(conn_options_row, bg=colors['background'])
        target_conn_frame.pack(side=tk.LEFT)

        target_conn_opt = tk.Frame(target_conn_frame, bg=colors['background'])
        target_icon_label = tk.Label(target_conn_opt, bg=colors['background'], cursor='hand2')
        target_icon = radio_on_icon  # Selected by default
        if target_icon:
            target_icon_label.configure(image=target_icon)
            target_icon_label._icon_ref = target_icon
        target_icon_label.pack(side=tk.LEFT, padx=(0, 4))
        target_text_label = tk.Label(
            target_conn_opt, text="Target Connections", font=('Segoe UI', 9),
            fg=colors['text_primary'], bg=colors['background'], cursor='hand2'
        )
        target_text_label.pack(side=tk.LEFT)
        target_conn_opt.pack(anchor=tk.W)
        conn_source_radios.append(("target", target_icon_label, target_text_label))

        def on_target_click(event=None):
            if scope_var.get() == "global":  # Only respond if Universal is selected
                conn_source_var.set("target")
                update_conn_source_radios()

        target_icon_label.bind('<Button-1>', on_target_click)
        target_text_label.bind('<Button-1>', on_target_click)

        # Add tooltip to Target Connections
        target_tooltip_text = (
            "Save the configured target connections as the preset target.\n"
            "Use this to create a preset that points to your selected target."
        )
        Tooltip(target_text_label, target_tooltip_text, delay=400)
        Tooltip(target_icon_label, target_tooltip_text, delay=400)

        # Connection source description text
        conn_desc_frame = tk.Frame(frame, bg=colors['background'])
        conn_desc_frame.pack(fill=tk.X, pady=(2, 0))

        conn_desc = tk.Label(
            conn_desc_frame,
            text="Source: Save current model connections (where it's pointing now). "
                 "Target: Save the configured target connections (where you want to point).",
            font=("Segoe UI", 8, "italic"),
            fg=colors.get('text_muted', '#888888'),
            bg=colors['background'],
            wraplength=470,
            justify=tk.LEFT
        )
        conn_desc.pack(anchor=tk.W)

        # Helper text
        helper_text = tk.Label(
            frame,
            text="Universal presets configure TARGET connections when applied.",
            font=("Segoe UI", 8, "italic"),
            fg=colors.get('text_muted', '#888888'),
            bg=colors['background']
        )
        helper_text.pack(anchor=tk.W, pady=(4, 0))

        # Function to enable/disable connection source section based on scope
        def update_conn_section_state():
            """Enable/disable connection source section based on scope selection"""
            is_universal = scope_var.get() == "global"
            muted_color = colors.get('text_muted', '#888888')
            normal_color = colors['text_primary']

            if is_universal:
                # Enable the section
                conn_section_label.configure(fg=normal_color)
                source_text_label.configure(fg=normal_color, cursor='hand2')
                target_text_label.configure(fg=normal_color, cursor='hand2')
                source_icon_label.configure(cursor='hand2')
                target_icon_label.configure(cursor='hand2')
                conn_desc.configure(
                    text="Source: Save current model connections (where it's pointing now). "
                         "Target: Save the configured target connections (where you want to point)."
                )
                helper_text.configure(text="Universal presets configure TARGET connections when applied.")
                # Update radio icons based on actual selection
                update_conn_source_radios()
            else:
                # Disable the section for Model-specific
                conn_section_label.configure(fg=muted_color)
                source_text_label.configure(fg=muted_color, cursor='')
                target_text_label.configure(fg=muted_color, cursor='')
                source_icon_label.configure(cursor='')
                target_icon_label.configure(cursor='')
                conn_desc.configure(
                    text="Maps each connection by name to its saved target (e.g., 'All Local', 'All Dev Cloud')."
                )
                helper_text.configure(text="Applies saved targets to matching connection names regardless of current source.")
                # Show BOTH radios as checked since model-specific saves both
                if radio_on_icon:
                    source_icon_label.configure(image=radio_on_icon)
                    source_icon_label._icon_ref = radio_on_icon
                    target_icon_label.configure(image=radio_on_icon)
                    target_icon_label._icon_ref = radio_on_icon

        # Update the scope radio click handlers to also update the conn section
        original_update_all_radios = update_all_radios

        def update_all_radios():
            """Update all radio button icons based on current values"""
            for value, icon_lbl, _ in scope_radios:
                is_selected = scope_var.get() == value
                icon = radio_on_icon if is_selected else radio_off_icon
                if icon:
                    icon_lbl.configure(image=icon)
                    icon_lbl._icon_ref = icon
            # Also update connection section state
            update_conn_section_state()

        # Initialize connection section state (Model-specific is default, so disable it)
        update_conn_section_state()

        # Buttons - centered
        canvas_bg = colors['background']
        btn_frame = tk.Frame(frame, bg=colors['background'])
        btn_frame.pack(pady=(10, 0))

        def get_preset_type_from_mappings(mappings_list):
            """Determine if preset is Cloud or Local based on TARGET mappings only.

            Presets save target connections, so type is determined by target, not source.
            """
            for mapping in mappings_list:
                # Only check target - presets save target connections
                if hasattr(mapping, 'target') and mapping.target:
                    if mapping.target.target_type == "cloud":
                        return "Cloud"
                    server = mapping.target.server or ""
                    if 'powerbi://' in server.lower() or 'asazure://' in server.lower():
                        return "Cloud"
            return "Local"

        def on_save():
            name = name_var.get().strip()
            if not name:
                ThemedMessageBox.showwarning(self.frame.winfo_toplevel(), "Name Required", "Please enter a preset name.")
                return

            # Determine scope
            scope_val = scope_var.get()
            scope = PresetScope.GLOBAL if scope_val == "global" else PresetScope.MODEL

            # Validate global scope
            if scope == PresetScope.GLOBAL and len(self.mappings) != 1:
                ThemedMessageBox.showwarning(
                    self.frame.winfo_toplevel(),
                    "Invalid Scope",
                    "Global presets can only be created for models with a single connection."
                )
                return

            # Always use USER storage (AppData)
            storage_type = PresetStorageType.USER

            # Get model info for model-specific presets
            model_hash = self._get_model_hash() if scope == PresetScope.MODEL else None
            model_name = self._get_model_display_name() if scope == PresetScope.MODEL else None

            # Determine which connections to save
            use_source = conn_source_var.get() == "source"

            # Validate target connections are configured if saving targets
            if not use_source:
                configured = [m for m in self.mappings if m.target]
                if not configured:
                    ThemedMessageBox.showwarning(
                        self.frame.winfo_toplevel(),
                        "No Targets Configured",
                        "Configure at least one target connection before saving.\n\n"
                        "Or select 'Source Connections' to save the current source connections."
                    )
                    return

            if use_source:
                # Create modified mappings where source connections become targets
                from tools.connection_hotswap.models import SwapTarget, CloudConnectionType, ConnectionMapping
                import copy

                modified_mappings = []
                for mapping in self.mappings:
                    source = mapping.source
                    # Create a SwapTarget from the source connection
                    source_as_target = SwapTarget(
                        target_type="cloud" if source.is_cloud else "local",
                        server=source.server,
                        database=source.database,
                        display_name=source.display_name,
                        workspace_name=source.workspace_name if source.is_cloud else None,
                        perspective_name=source.perspective_name,
                        # Infer cloud connection type from connection string
                        cloud_connection_type=CloudConnectionType.PBI_SEMANTIC_MODEL if (
                            source.is_cloud and 'powerbi://' in source.server.lower()
                        ) else (
                            CloudConnectionType.AAS_XMLA if source.is_cloud else None
                        )
                    )
                    # Create new mapping with source as target
                    new_mapping = ConnectionMapping(
                        source=source,
                        target=source_as_target,
                        auto_matched=False,
                        status=mapping.status,
                        original_connection_string=mapping.original_connection_string
                    )
                    modified_mappings.append(new_mapping)

                mappings_to_save = modified_mappings
            else:
                # Use current target mappings (existing behavior)
                mappings_to_save = self.mappings

            # Determine the type of preset being saved (Cloud or Local)
            new_preset_type = get_preset_type_from_mappings(mappings_to_save)

            # Check for existing preset with same name
            existing_preset = self.preset_manager.get_preset(
                name=name,
                storage_type=storage_type,
                scope=scope,
                model_hash=model_hash
            )

            if existing_preset:
                # Determine existing preset's type
                existing_type = "Local"
                for mapping in existing_preset.mappings:
                    if mapping.target_type == "cloud":
                        existing_type = "Cloud"
                        break
                    if mapping.server:
                        server_lower = mapping.server.lower()
                        if 'powerbi://' in server_lower or 'asazure://' in server_lower:
                            existing_type = "Cloud"
                            break

                if existing_type == new_preset_type:
                    # Same name AND same type - prompt for overwrite
                    result = ThemedMessageBox.askyesno(
                        self.frame.winfo_toplevel(),
                        "Preset Exists",
                        f"A {existing_type} preset named '{name}' already exists.\n\n"
                        f"Do you want to update/overwrite it?"
                    )
                    if not result:
                        return  # User cancelled
                else:
                    # Same name but different type - allow by appending type suffix
                    # Check if a preset with the suffixed name already exists
                    suffixed_name = f"{name} ({new_preset_type})"
                    suffixed_existing = self.preset_manager.get_preset(
                        name=suffixed_name,
                        storage_type=storage_type,
                        scope=scope,
                        model_hash=model_hash
                    )
                    if suffixed_existing:
                        # Even the suffixed name exists, prompt for overwrite
                        result = ThemedMessageBox.askyesno(
                            self.frame.winfo_toplevel(),
                            "Preset Exists",
                            f"A preset named '{suffixed_name}' already exists.\n\n"
                            f"Do you want to update/overwrite it?"
                        )
                        if not result:
                            return
                    else:
                        # Inform user we're saving with suffixed name
                        ThemedMessageBox.showinfo(
                            self.frame.winfo_toplevel(),
                            "Name Adjusted",
                            f"A {existing_type} preset named '{name}' already exists.\n\n"
                            f"Your {new_preset_type} preset will be saved as:\n'{suffixed_name}'"
                        )
                    name = suffixed_name

            # Create and save preset
            preset = self.preset_manager.create_preset_from_mappings(
                name=name,
                mappings=mappings_to_save,
                storage_type=storage_type,
                scope=scope,
                model_hash=model_hash,
                model_name=model_name
            )

            # Log model_hash for debugging preset persistence issues
            if scope == PresetScope.MODEL:
                self.logger.info(f"Saving MODEL preset '{name}' with hash: {model_hash}")

            if self.preset_manager.save_preset(preset):
                source_type = "source" if use_source else "target"
                self._log_message(f"Saved preset '{name}' ({source_type} connections)")
                self._refresh_preset_buttons()
                dialog.destroy()
            else:
                ThemedMessageBox.showerror(self.frame.winfo_toplevel(), "Save Failed", "Could not save preset. Check log for details.")

        # Inner frame to hold buttons for centering
        btn_inner = tk.Frame(btn_frame, bg=colors['background'])
        btn_inner.pack(anchor=tk.CENTER)

        cancel_btn = RoundedButton(
            btn_inner, text="CANCEL",
            command=dialog.destroy,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 10),
            canvas_bg=canvas_bg
        )
        cancel_btn.pack(side=tk.LEFT)

        save_btn = RoundedButton(
            btn_inner, text="SAVE",
            command=on_save,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 10, 'bold'),
            canvas_bg=canvas_bg
        )
        save_btn.pack(side=tk.LEFT, padx=(10, 0))

        name_entry.focus_set()

    def _on_manage_presets(self):
        """Show preset management dialog"""
        colors = self._theme_manager.colors
        parent = self.frame.winfo_toplevel()

        dialog = tk.Toplevel(parent)
        dialog.title("Manage Presets")
        dialog.geometry("500x400")
        dialog.resizable(True, True)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

        # Center on parent
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 500) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 400) // 2
        dialog.geometry(f"+{x}+{y}")

        # Content
        frame = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            frame,
            text="Environment Presets",
            font=("Segoe UI", 12, "bold"),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(anchor=tk.W, pady=(0, 10))

        # Preset list
        is_dark = self._theme_manager.is_dark
        list_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        list_border = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')

        list_frame = tk.Frame(
            frame,
            bg=list_bg,
            highlightbackground=list_border,
            highlightcolor=list_border,
            highlightthickness=1
        )
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Treeview for presets
        columns = ("name", "type", "mappings", "updated")
        preset_tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
            height=10
        )

        preset_tree.heading("name", text="Name")
        preset_tree.heading("type", text="Storage")
        preset_tree.heading("mappings", text="Mappings")
        preset_tree.heading("updated", text="Updated")

        preset_tree.column("name", width=150)
        preset_tree.column("type", width=80)
        preset_tree.column("mappings", width=70)
        preset_tree.column("updated", width=150)

        scrollbar = ThemedScrollbar(list_frame, command=preset_tree.yview, theme_manager=self._theme_manager)
        preset_tree.configure(yscrollcommand=scrollbar.set)

        preset_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Populate presets
        presets = self.preset_manager.list_presets()
        for preset in presets:
            storage_icon = "👤" if preset.storage_type == PresetStorageType.USER else "📁"
            updated = preset.updated_at[:10] if preset.updated_at else "Unknown"
            preset_tree.insert(
                "",
                "end",
                iid=preset.name,
                values=(preset.name, storage_icon, len(preset.mappings), updated)
            )

        # Buttons
        canvas_bg = colors['background']
        btn_frame = tk.Frame(frame, bg=colors['background'])
        btn_frame.pack(fill=tk.X)

        def on_delete():
            selection = preset_tree.selection()
            if not selection:
                return

            name = selection[0]
            preset = self.preset_manager.get_preset(name)
            if not preset:
                return

            confirm = ThemedMessageBox.askyesno(
                dialog,
                "Delete Preset",
                f"Delete preset '{name}'?\n\nThis cannot be undone."
            )

            if confirm:
                if self.preset_manager.delete_preset(name, preset.storage_type):
                    preset_tree.delete(name)
                    self._refresh_preset_buttons()
                    self._log_message(f"Deleted preset '{name}'")

        delete_btn = RoundedButton(
            btn_frame, text="DELETE",
            command=on_delete,
            bg=colors.get('error', '#ef4444'),
            hover_bg=colors.get('error_hover', '#dc2626'),
            pressed_bg=colors.get('error_pressed', '#b91c1c'),
            fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 10, 'bold'),
            canvas_bg=canvas_bg
        )
        delete_btn.pack(side=tk.LEFT, padx=(0, 10))

        close_btn = RoundedButton(
            btn_frame, text="CLOSE",
            command=dialog.destroy,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 10),
            canvas_bg=canvas_bg
        )
        close_btn.pack(side=tk.RIGHT)

    # =========================================================================
    # Logging
    # =========================================================================

    def _log_message(self, message: str):
        """Add message to the log"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")

        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        """Clear the log"""
        self.log_text.configure(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state="disabled")

    # =========================================================================
    # Required BaseToolTab Abstract Methods
    # =========================================================================

    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        self.logger.info("Resetting Connection Hot-Swap tab")
        self._on_disconnect()
        self._clear_log()
        self.model_combo['values'] = []
        self.selected_model.set("")

    def _trigger_cloud_preload(self):
        """
        Trigger cloud workspace and dataset preloading in background.
        Called from Refresh button and on tab activation.
        """
        # Initialize cloud_browser if needed
        if not self.cloud_browser:
            try:
                from core.cloud import CloudWorkspaceBrowser
                self.cloud_browser = CloudWorkspaceBrowser()
                self.logger.info("CloudWorkspaceBrowser initialized for preload")
            except Exception as e:
                self.logger.info(f"Could not initialize cloud browser: {e}")
                return

        # Skip if already authenticated and fully cached
        if self.cloud_browser.is_authenticated() and self.cloud_browser.is_fully_cached():
            self.logger.info("Cloud data already fully cached, skipping preload")
            return

        # Skip if preload already in progress
        if getattr(self, '_cloud_preload_in_progress', False):
            self.logger.info("Cloud preload already in progress")
            return
        self._cloud_preload_in_progress = True

        def preload_thread():
            try:
                # Try silent auth first
                if not self.cloud_browser.is_authenticated():
                    self.logger.info("Attempting silent auth for preload...")
                    success = self.cloud_browser.try_silent_auth()
                    if not success:
                        self.logger.info("Silent auth failed, user will auth manually")
                        return
                    self.logger.info("Silent auth succeeded")

                # Pre-load workspaces
                if not self.cloud_browser.is_session_cached():
                    self.logger.info("Pre-loading workspaces...")
                    workspaces, _ = self.cloud_browser.get_workspaces("all")
                    if workspaces:
                        self.logger.info(f"Pre-loaded {len(workspaces)} workspaces")

                # Pre-load all datasets (this is what makes the dialog fast)
                if not self.cloud_browser.is_fully_cached():
                    self.logger.info("Pre-loading datasets...")
                    total, error = self.cloud_browser.preload_all_datasets()
                    if not error:
                        self.logger.info(f"Pre-loaded {total} datasets")
                    else:
                        self.logger.info(f"Dataset preload error: {error}")
            except Exception as e:
                self.logger.info(f"Cloud preload failed: {e}")
            finally:
                self._cloud_preload_in_progress = False

        threading.Thread(target=preload_thread, daemon=True).start()
        self.logger.info("Cloud preload thread launched")

    def on_tab_activated(self) -> None:
        """
        Called when this tab becomes active.
        Attempts silent cloud authentication and pre-loads workspace data.
        Also triggers auto-scan for local Power BI models if cache is empty/stale.
        """
        self.logger.info("on_tab_activated called")
        self._trigger_cloud_preload()
        # Auto-scan for local models if cache is empty or stale
        self._auto_scan_if_needed()

    def _auto_scan_if_needed(self) -> None:
        """
        Auto-scan for local Power BI models if the shared cache is empty or stale.
        This provides immediate model discovery when switching to the Hot Swap tab.
        """
        cache = get_local_model_cache()

        # Skip if scan already in progress
        if cache.is_scan_in_progress():
            self.logger.debug("Scan already in progress, skipping auto-scan")
            return

        # Skip if cache is fresh and has models
        if not cache.is_empty() and not cache.is_stale():
            # Populate dropdown from cache without re-scanning
            self.logger.debug("Using cached models for dropdown")
            models = cache.get_models()
            if models:
                self.frame.after(0, lambda: self._update_model_list(models))
            return

        # Trigger scan in background
        self.logger.info("Auto-scanning for local Power BI models...")
        self._show_progress("Scanning for Power BI models...")

        def on_complete(models):
            self.frame.after(0, lambda: self._update_model_list(models))
            self.frame.after(0, self._hide_progress)

        def on_error(error_msg):
            self.logger.warning(f"Auto-scan failed: {error_msg}")
            self.frame.after(0, self._hide_progress)

        cache.scan_async(
            connector=self.connector,
            force=False,
            on_complete=on_complete,
            on_error=on_error
        )

    def show_help_dialog(self) -> None:
        """Show help dialog for Connection Hot-Swap tool"""
        # Get the correct parent window
        parent_window = self.frame.winfo_toplevel()

        help_window = tk.Toplevel(parent_window)
        help_window.withdraw()  # Hide until fully styled (prevents white flash)
        help_window.title("Connection Hot-Swap - Help")
        help_window.geometry("900x700")
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
        """Create help content for Connection Hot-Swap"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        help_bg = colors['background']
        help_window.configure(bg=help_bg)

        # Set dark title bar if dark mode
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, is_dark)

        # Main container
        container = tk.Frame(help_window, bg=help_bg)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        tk.Label(container, text="Connection Hot-Swap - Help",
                 font=('Segoe UI', 16, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.CENTER, pady=(0, 15))

        # Orange warning box with white text - two column layout
        warning_frame = tk.Frame(container, bg=help_bg)
        warning_frame.pack(fill=tk.X, pady=(0, 15))

        warning_bg = colors.get('warning_bg', '#d97706')
        warning_text = colors.get('warning_text', '#ffffff')
        warning_container = tk.Frame(warning_frame, bg=warning_bg,
                                   padx=15, pady=10, relief='flat', borderwidth=0)
        warning_container.pack(fill=tk.X)

        tk.Label(warning_container, text="IMPORTANT NOTES",
                 font=('Segoe UI', 12, 'bold'),
                 bg=warning_bg,
                 fg=warning_text).pack(anchor=tk.W, pady=(0, 5))

        # Two-column grid for warnings
        warnings_grid = tk.Frame(warning_container, bg=warning_bg)
        warnings_grid.pack(fill=tk.X)
        warnings_grid.columnconfigure(0, weight=1)
        warnings_grid.columnconfigure(1, weight=1)

        warnings_left = [
            "Composite models: Hot-swap while open (TOM-based)",
            "Thin reports: Close & reopen after swap (PBIP edits while open)",
            "Changes applied immediately - save report afterward"
        ]

        warnings_right = [
            "Perspectives supported in ALL workspaces (including Pro)",
            "Cloud connections: Pro, Premium, PPU, and Fabric",
            "Use Rollback to restore if issues occur"
        ]

        for i, warning in enumerate(warnings_left):
            tk.Label(warnings_grid, text=f"• {warning}", font=('Segoe UI', 10),
                     bg=warning_bg, fg=warning_text,
                     anchor='w').grid(row=i, column=0, sticky='w', padx=(0, 10), pady=1)

        for i, warning in enumerate(warnings_right):
            tk.Label(warnings_grid, text=f"• {warning}", font=('Segoe UI', 10),
                     bg=warning_bg, fg=warning_text,
                     anchor='w').grid(row=i, column=1, sticky='w', padx=(10, 0), pady=1)

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
            "Swap connections: cloud-to-local, local-to-cloud, or cloud-to-cloud",
            "Composite models hot-swap live; thin reports require reload after swap",
            "Connect to perspectives in ANY workspace type (Pro, Premium, PPU, Fabric)",
            "Support for PBI Semantic Model and XMLA endpoint connections",
            "Auto-match local models by name for quick setup",
            "Save and reuse connection presets for repeated workflows"
        ]

        for item in what_items:
            tk.Label(left_top_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=400,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        # ROW 0, COLUMN 1: Quick Start Guide
        right_top_frame = tk.Frame(sections_frame, bg=help_bg)
        right_top_frame.grid(row=0, column=1, sticky='nwe', padx=(10, 0), pady=(0, 15))

        tk.Label(right_top_frame, text="Quick Start Guide",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        quickstart_items = [
            "1. Open your Power BI report in Desktop",
            "2. Click REFRESH to find open models",
            "3. Select your model and click CONNECT",
            "4. Double-click Target column to set swap target",
            "5. Choose Local or Cloud target for each connection",
            "6. Click SWAP SELECTED to apply changes"
        ]

        for item in quickstart_items:
            tk.Label(right_top_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=400,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        # ROW 1, COLUMN 0: Quick Swap & Presets
        left_bottom_frame = tk.Frame(sections_frame, bg=help_bg)
        left_bottom_frame.grid(row=1, column=0, sticky='nwe', padx=(0, 10))

        tk.Label(left_bottom_frame, text="Quick Swap & Presets",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        preset_items = [
            "GLOBAL presets: Available across all reports",
            "MODEL presets: Specific to the current model",
            "APPLY PRESET: Load saved connection mappings",
            "LAST CONFIG: Set target to your original connection (before swapping)",
            "SAVE MAPPING: Save current targets as a new preset",
            "Import/Export: Share presets via JSON files"
        ]

        for item in preset_items:
            tk.Label(left_bottom_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=400,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        # ROW 1, COLUMN 1: Tips & Troubleshooting
        right_bottom_frame = tk.Frame(sections_frame, bg=help_bg)
        right_bottom_frame.grid(row=1, column=1, sticky='nwe', padx=(10, 0))

        tk.Label(right_bottom_frame, text="Tips & Troubleshooting",
                 font=('Segoe UI', 12, 'bold'),
                 bg=help_bg,
                 fg=colors['title_color']).pack(anchor=tk.W, pady=(0, 5))

        tips_items = [
            "Enable 'Auto-match local models' for automatic target detection",
            "Use Target Presets popup for quick local model selection",
            "Browse Cloud requires Azure AD authentication (one-time per session)",
            "After swap, refresh visuals in Desktop to load new data",
            "If swap fails, use ROLLBACK to restore the original connection",
            "Right-click rows for Edit Target and Clear Target options"
        ]

        for item in tips_items:
            tk.Label(right_bottom_frame, text=item,
                     font=('Segoe UI', 10),
                     bg=help_bg,
                     fg=colors['text_primary'],
                     wraplength=400,
                     justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

        help_window.bind('<Escape>', lambda event: help_window.destroy())

        # Center dialog and show (after all content built to prevent flash)
        help_window.update_idletasks()
        x = parent_window.winfo_rootx() + (parent_window.winfo_width() - 900) // 2
        y = parent_window.winfo_rooty() + (parent_window.winfo_height() - 700) // 2
        help_window.geometry(f"900x700+{x}+{y}")

        # Set dark title bar BEFORE showing to prevent white flash
        help_window.update()
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled

    def on_theme_changed(self, theme: str):
        """Handle theme change - updates all theme-dependent widgets"""
        super().on_theme_changed(theme)
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        section_bg = colors.get('section_bg', colors['background'])
        main_bg = colors['background']
        outer_bg = colors.get('outer_bg', '#161627' if is_dark else '#f5f5f7')
        # Distinct content background for Model Connection and Swap Configuration sections
        section_content_bg = '#0d0d1a' if is_dark else '#ffffff'

        # Update Model Connection section inner frame and labels (uses section_content_bg)
        conn_widgets = [
            '_conn_inner_frame', '_conn_row_frame', '_conn_left_frame',
            '_conn_button_frame', '_conn_row2_frame', '_conn_row2_left',
            '_conn_row2_right', '_conn_btn_container'
        ]
        for attr in conn_widgets:
            if hasattr(self, attr):
                try:
                    getattr(self, attr).configure(bg=section_content_bg)
                except Exception:
                    pass

        # Update Model Connection labels
        if hasattr(self, '_model_label'):
            try:
                self._model_label.configure(bg=section_content_bg, fg=colors['text_primary'])
            except Exception:
                pass
        if hasattr(self, '_status_text_label'):
            try:
                self._status_text_label.configure(bg=section_content_bg, fg=colors['text_muted'])
            except Exception:
                pass
        if hasattr(self, 'status_label') and self.status_label:
            try:
                current_status = self.connection_status.get()
                fg_color = colors['success'] if current_status == "Connected" else colors['error']
                self.status_label.configure(bg=section_content_bg, fg=fg_color)
            except Exception:
                pass

        # Update status dot for theme
        if hasattr(self, '_status_dot') and self._status_dot:
            try:
                self._status_dot.configure(bg=section_content_bg)
                # Redraw with current state to get correct theme colors
                if hasattr(self, '_current_status_state'):
                    self._draw_status_dot(self._current_status_state)
            except Exception:
                pass

        # Update cloud auth button for theme
        if hasattr(self, '_cloud_auth_btn'):
            try:
                # Close dropdown if open (will rebuild with new theme on reopen)
                if hasattr(self, '_cloud_auth_dropdown') and self._cloud_auth_dropdown:
                    self._close_cloud_auth_dropdown()

                # Reload cloud auth icons with new theme
                self._load_cloud_auth_icons()

                # Update button state (will apply correct icon)
                self._update_cloud_auth_button_state()
            except Exception:
                pass

        # Update progress label for theme (Deep Scan text)
        if hasattr(self, 'progress_label') and self.progress_label:
            try:
                self.progress_label.configure(bg=section_content_bg, fg=colors['text_muted'])
            except Exception:
                pass

        # Update model dropdown style for theme
        try:
            style = ttk.Style()
            text_color = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
            # Use button_primary for border (blue in dark mode, teal in light mode)
            border_color = colors.get('button_primary', '#00587C' if is_dark else '#009999')
            style.configure('TCombobox',
                          fieldbackground=colors.get('card_surface', '#2a2a3c' if is_dark else '#ffffff'),
                          background=colors.get('card_surface', '#2a2a3c' if is_dark else '#ffffff'),
                          foreground=text_color,
                          arrowcolor=colors.get('text_secondary', '#a0a0b0' if is_dark else '#666666'),
                          bordercolor=border_color)
            style.map('TCombobox',
                     fieldbackground=[('readonly', colors.get('card_surface', '#2a2a3c' if is_dark else '#ffffff')),
                                     ('disabled', colors.get('background', '#1e1e2e' if is_dark else '#f5f5f7'))],
                     foreground=[('readonly', text_color),
                                ('disabled', colors.get('text_muted', '#808080'))],
                     bordercolor=[('focus', border_color),
                                 ('!focus', border_color)])
        except Exception:
            pass

        # Update combobox dropdown listbox styling for theme
        try:
            self._apply_combobox_listbox_style()
        except Exception:
            pass

        # Update progress bar style for theme (kept for other components that may use it)
        try:
            style = ttk.Style()
            # Style the progress bar trough and bar colors
            bar_bg = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
            bar_fg = colors.get('primary', '#4a9eff')
            style.configure('TProgressbar',
                          troughcolor=bar_bg,
                          background=bar_fg,
                          bordercolor=bar_bg,
                          lightcolor=bar_fg,
                          darkcolor=bar_fg)
        except Exception:
            pass

        # Update animated dots indicator for theme
        if hasattr(self, '_scanning_dots') and self._scanning_dots:
            try:
                self._scanning_dots.update_colors(
                    bg=section_content_bg,
                    dot_color=colors.get('info', '#3b82f6')
                )
            except Exception:
                pass

        # Update Swap Configuration content area (uses section_content_bg)
        swap_config_widgets = [
            '_swap_config_content_frame', '_swap_left_container', '_swap_right_container'
        ]
        for attr in swap_config_widgets:
            if hasattr(self, attr):
                try:
                    getattr(self, attr).configure(bg=section_content_bg)
                except Exception:
                    pass

        # Update Quick Swap subsection content box (white/surface content_bg)
        content_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')

        # Update inner frame background (no border)
        quickswap_bg = '#0d0d1a' if is_dark else '#ffffff'
        if hasattr(self, '_quickswap_inner_frame'):
            try:
                self._quickswap_inner_frame.configure(bg=quickswap_bg)
            except Exception:
                pass

        # Update unified button row and containers (uses quickswap_bg)
        if hasattr(self, '_unified_btn_row'):
            try:
                self._unified_btn_row.configure(bg=quickswap_bg)
            except Exception:
                pass

        if hasattr(self, '_left_btn_container'):
            try:
                self._left_btn_container.configure(bg=quickswap_bg)
            except Exception:
                pass

        if hasattr(self, '_right_btn_container'):
            try:
                self._right_btn_container.configure(bg=quickswap_bg)
            except Exception:
                pass

        # Update preset action buttons canvas_bg and colors (inside the quickswap content frame)
        # Get disabled colors for theme
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        if hasattr(self, 'last_config_btn'):
            try:
                self.last_config_btn.update_canvas_bg(quickswap_bg)
                self.last_config_btn.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    disabled_bg=disabled_bg,
                    disabled_fg=disabled_fg
                )
            except Exception:
                pass
        # Note: apply_preset_btn canvas_bg is updated with other buttons in unified row later

        # Update Quick Swap subsection header (uses section_content_bg)
        if hasattr(self, '_quickswap_header_frame'):
            try:
                self._quickswap_header_frame.configure(bg=section_content_bg)
                if self._quickswap_header_icon:
                    self._quickswap_header_icon.configure(bg=section_content_bg)
                self._quickswap_header_text.configure(bg=section_content_bg, fg=colors['title_color'])
            except Exception:
                pass

        # Update header icon button frame (uses section_content_bg - same as header)
        if hasattr(self, '_header_icon_btn_frame'):
            try:
                self._header_icon_btn_frame.configure(bg=section_content_bg)
            except Exception:
                pass

        # Update Environment Presets label and header row (uses quickswap_bg)
        if hasattr(self, '_presets_header_row'):
            try:
                self._presets_header_row.configure(bg=quickswap_bg)
            except Exception:
                pass
        if hasattr(self, '_presets_label'):
            try:
                self._presets_label.configure(bg=quickswap_bg, fg=colors['text_primary'])
            except Exception:
                pass

        # Reload checkbox icons for new theme
        self._load_checkbox_icons()

        # Update auto-match toggle frame, labels, and icon (uses section_content_bg)
        if hasattr(self, '_auto_match_toggle_frame'):
            try:
                self._auto_match_toggle_frame.configure(bg=section_content_bg)
                self._auto_match_icon_label.configure(bg=section_content_bg)
                self._auto_match_text_label.configure(bg=section_content_bg, fg=colors['text_primary'])
                # Update toggle icon with new theme-aware icon
                self._update_auto_match_toggle()
            except Exception:
                pass

        # Update preset scope toggle (uses quickswap_bg)
        if hasattr(self, '_preset_scope_toggle_frame'):
            try:
                self._preset_scope_toggle_frame.configure(bg=quickswap_bg)
                if hasattr(self, '_preset_scope_tab_container'):
                    self._preset_scope_tab_container.configure(bg=quickswap_bg)
                # Update button colors and canvas_bg
                if hasattr(self, '_scope_global_btn'):
                    self._scope_global_btn.update_canvas_bg(quickswap_bg)
                if hasattr(self, '_scope_model_btn'):
                    self._scope_model_btn.update_canvas_bg(quickswap_bg)
                self._update_preset_scope_toggle()
            except Exception:
                pass

        # Update preset table container border (same as Progress Log content)
        if hasattr(self, '_preset_table_container'):
            try:
                tree_bg = '#161627' if is_dark else '#f5f5f7'
                tree_border = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
                self._preset_table_container.configure(
                    bg=tree_bg,
                    highlightbackground=tree_border,
                    highlightcolor=tree_border
                )
            except Exception:
                pass

        # Update preset scrollbar area background (matches Progress Log)
        if hasattr(self, '_preset_scrollbar_area'):
            try:
                scrollbar_area_bg = '#1a1a2e' if is_dark else '#f0f0f0'
                self._preset_scrollbar_area.configure(bg=scrollbar_area_bg)
            except Exception:
                pass

        # Update preset table style
        if hasattr(self, '_preset_tree'):
            try:
                # Use same colors as Progress Log content
                preset_tree_bg = '#161627' if is_dark else '#f5f5f7'
                preset_tree_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
                preset_heading_bg = colors.get('section_bg', '#1a1a2a' if is_dark else '#f5f5fa')
                preset_heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
                preset_header_sep = '#0d0d1a' if is_dark else '#ffffff'

                style = ttk.Style()
                preset_tree_style = "Presets.Treeview"
                style.configure(preset_tree_style,
                                background=preset_tree_bg,
                                foreground=preset_tree_fg,
                                fieldbackground=preset_tree_bg)
                style.configure(f"{preset_tree_style}.Heading",
                                background=preset_heading_bg,
                                foreground=preset_heading_fg,
                                relief='groove',
                                borderwidth=1,
                                bordercolor=preset_header_sep,
                                lightcolor=preset_header_sep,
                                darkcolor=preset_header_sep)
                style.map(f"{preset_tree_style}.Heading",
                          background=[('active', preset_heading_bg), ('pressed', preset_heading_bg), ('', preset_heading_bg)])
                # Custom heading layout to center image element vertically
                style.layout(f"{preset_tree_style}.Heading", [
                    ('Treeheading.cell', {'sticky': 'nswe'}),
                    ('Treeheading.border', {'sticky': 'nswe', 'children': [
                        ('Treeheading.padding', {'sticky': 'nswe', 'children': [
                            ('Treeheading.image', {'sticky': ''}),
                            ('Treeheading.text', {'sticky': 'we'})
                        ]})
                    ]})
                ])
                style.map(preset_tree_style,
                          background=[('selected', '#1a3a5c' if is_dark else '#e6f3ff')],
                          foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])
            except Exception:
                pass

        # Update Connections subsection header (uses section_content_bg)
        if hasattr(self, '_connections_header_frame'):
            try:
                self._connections_header_frame.configure(bg=section_content_bg)
                if self._connections_header_icon:
                    self._connections_header_icon.configure(bg=section_content_bg)
                if self._connections_header_text:
                    self._connections_header_text.configure(bg=section_content_bg, fg=colors['title_color'])
            except Exception:
                pass

        # Update connections spacer frame (uses section_content_bg) - if it exists
        if hasattr(self, '_connections_spacer') and self._connections_spacer:
            try:
                self._connections_spacer.configure(bg=section_content_bg)
            except Exception:
                pass

        # Update Connection Mappings section inner frame and labels (uses section_content_bg)
        mapping_widgets = ['_mapping_inner_frame', '_mapping_top_row']
        for attr in mapping_widgets:
            if hasattr(self, attr):
                try:
                    getattr(self, attr).configure(bg=section_content_bg)
                except Exception:
                    pass

        if hasattr(self, '_mapping_info_label'):
            try:
                self._mapping_info_label.configure(bg=section_content_bg, fg=colors['text_muted'])
            except Exception:
                pass

        # Update Select All link - use title_color for consistency with section headers
        if hasattr(self, '_select_all_link'):
            try:
                if is_dark:
                    link_color = colors.get('title_color', '#0084b7')  # Light blue matching title
                    hover_color = '#006691'  # Darker blue for hover
                else:
                    link_color = colors.get('primary', '#009999')  # Teal for light mode
                    hover_color = colors.get('primary_hover', '#007A7A')  # Darker teal for hover
                self._select_all_link.configure(bg=section_content_bg, fg=link_color)
                # Update hover bindings with new colors
                self._select_all_link.bind('<Enter>', lambda e: self._select_all_link.configure(fg=hover_color))
                self._select_all_link.bind('<Leave>', lambda e: self._select_all_link.configure(fg=link_color))
            except Exception:
                pass

        # Update tree container border and background (same as Progress Log content)
        if hasattr(self, '_tree_container') and self._tree_container:
            try:
                tree_bg = '#161627' if is_dark else '#f5f5f7'
                tree_border = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
                self._tree_container.configure(
                    bg=tree_bg,
                    highlightbackground=tree_border,
                    highlightcolor=tree_border
                )
            except Exception:
                pass

        # Update mapping scrollbar area background (matches Progress Log)
        if hasattr(self, '_mapping_scrollbar_area'):
            try:
                scrollbar_area_bg = '#1a1a2e' if is_dark else '#f0f0f0'
                self._mapping_scrollbar_area.configure(bg=scrollbar_area_bg)
            except Exception:
                pass

        # Update middle pane background
        if hasattr(self, '_middle_pane'):
            try:
                self._middle_pane.configure(bg=main_bg)
            except Exception:
                pass

        # Update treeview style for theme change
        if hasattr(self, 'mapping_tree') and self.mapping_tree:
            try:
                # Use same colors as Progress Log content
                tree_bg = '#161627' if is_dark else '#f5f5f7'
                tree_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
                heading_bg = colors.get('section_bg', '#1a1a2a' if is_dark else '#f5f5fa')
                heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
                header_separator = '#0d0d1a' if is_dark else '#ffffff'

                style = ttk.Style()
                tree_style = "HotSwap.Treeview"
                style.configure(tree_style,
                                background=tree_bg,
                                foreground=tree_fg,
                                fieldbackground=tree_bg,
                                bordercolor=tree_bg,
                                lightcolor=tree_bg,
                                darkcolor=tree_bg)
                style.configure(f"{tree_style}.Heading",
                                background=heading_bg,
                                foreground=heading_fg,
                                bordercolor=header_separator,
                                lightcolor=header_separator,
                                darkcolor=header_separator)
                style.map(f"{tree_style}.Heading",
                          background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])
                style.map(tree_style,
                          background=[('selected', '#1a3a5c' if is_dark else '#e6f3ff')],
                          foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])
            except Exception:
                pass

        # SplitLogSection handles its own theme updates internally

        # Update primary buttons with theme-aware disabled colors
        primary_disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        primary_disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        for btn in self._primary_buttons:
            try:
                btn.update_colors(
                    bg=colors['button_primary'],
                    hover_bg=colors['button_primary_hover'],
                    pressed_bg=colors['button_primary_pressed'],
                    fg='#ffffff',
                    disabled_bg=primary_disabled_bg,
                    disabled_fg=primary_disabled_fg
                )
            except Exception:
                pass

        # Update bottom action buttons (SWAP SELECTED, ROLLBACK) with outer_bg canvas
        if hasattr(self, 'swap_selected_btn'):
            try:
                self.swap_selected_btn.update_canvas_bg(outer_bg)
            except Exception:
                pass
        if hasattr(self, 'rollback_btn'):
            try:
                self.rollback_btn.update_canvas_bg(outer_bg)
                # Use button_secondary colors to match Reset All buttons
                rollback_disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
                rollback_disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')
                self.rollback_btn.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    disabled_bg=rollback_disabled_bg,
                    disabled_fg=rollback_disabled_fg
                )
            except Exception:
                pass

        # Update secondary buttons (REFRESH, APPLY TO MAPPINGS, SAVE MAPPING, IMPORT)
        # These are in Model Connection and Swap Configuration sections (use section_content_bg)
        # Light mode disabled colors: #c0c0cc bg, #9a9aa8 fg
        # Dark mode disabled colors: #3a3a4e bg, #6a6a7a fg
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # refresh_btn is in Model Connection section (section_content_bg)
        if hasattr(self, 'refresh_btn'):
            try:
                self.refresh_btn.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    disabled_bg=disabled_bg,
                    disabled_fg=disabled_fg
                )
                self.refresh_btn.update_canvas_bg(section_content_bg)
            except Exception:
                pass

        # apply_preset_btn and save_mapping_btn are in unified button row (quickswap_bg)
        for btn_name in ['apply_preset_btn', 'save_mapping_btn']:
            if hasattr(self, btn_name):
                try:
                    btn = getattr(self, btn_name)
                    btn.update_colors(
                        bg=colors['button_secondary'],
                        hover_bg=colors['button_secondary_hover'],
                        pressed_bg=colors['button_secondary_pressed'],
                        fg=colors['text_primary'],
                        disabled_bg=disabled_bg,
                        disabled_fg=disabled_fg
                    )
                    btn.update_canvas_bg(quickswap_bg)
                except Exception:
                    pass

        # Update disconnect button with muted red/rose colors for visibility
        if hasattr(self, 'disconnect_btn'):
            try:
                disconnect_bg = '#503838' if is_dark else '#fef2f2'
                disconnect_hover = '#5a4040' if is_dark else '#fee2e2'
                disconnect_pressed = '#402828' if is_dark else '#fecaca'
                disconnect_fg = colors.get('error', '#ef4444' if is_dark else '#dc2626')
                disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
                disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')
                self.disconnect_btn.update_colors(
                    bg=disconnect_bg,
                    hover_bg=disconnect_hover,
                    pressed_bg=disconnect_pressed,
                    fg=disconnect_fg,
                    disabled_bg=disabled_bg,
                    disabled_fg=disabled_fg
                )
                self.disconnect_btn.update_canvas_bg(section_content_bg)
            except Exception:
                pass

        # Update connect_btn canvas_bg (in Model Connection section)
        if hasattr(self, 'connect_btn'):
            try:
                self.connect_btn.update_canvas_bg(section_content_bg)
            except Exception:
                pass

        # Re-render Connection Details card container if visible (cards are dynamically created)
        if hasattr(self, '_card_container') and self._card_container:
            try:
                # Get currently selected mapping from tree
                if hasattr(self, 'mapping_tree') and self.mapping_tree:
                    selection = self.mapping_tree.selection()
                    if selection:
                        item_id = selection[0]
                        idx = int(item_id)
                        if idx < len(self.mappings):
                            # Re-render with new theme colors
                            self._update_selected_connection_details(self.mappings[idx])
            except Exception:
                pass

    def apply_theme(self) -> None:
        """Apply theme changes (called by on_theme_changed)"""
        self.on_theme_changed(self._theme_manager.current_theme)
