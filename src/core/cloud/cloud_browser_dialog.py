"""
Cloud Browser Dialog - Browse Power BI Service workspaces and datasets
Built by Reid Havens of Analytic Endeavors

Modal dialog for browsing and selecting cloud datasets via Fabric API.
"""

import ctypes
import io
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from pathlib import Path
from typing import Optional, List, Dict

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

from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, SquareIconButton, ThemedScrollbar, ThemedMessageBox, LabeledRadioGroup
from core.cloud.models import SwapTarget, WorkspaceInfo, DatasetInfo, CloudConnectionType


class CloudBrowserDialog(tk.Toplevel):
    """
    Dialog for browsing Power BI Service workspaces and datasets.

    Features:
    - Authenticate via OAuth
    - Browse workspaces (all, favorites, recent)
    - Search across datasets
    - Manual XMLA entry
    """

    def __init__(
        self,
        parent,
        browser: 'CloudWorkspaceBrowser',
        default_connection_type: Optional[CloudConnectionType] = None,
        model_status: Optional[str] = None,
        on_auth_change: Optional[callable] = None,
        source_context: Optional[str] = None,
        simple_mode: bool = False
    ):
        """
        Initialize the cloud browser dialog.

        Args:
            parent: Parent window
            browser: CloudWorkspaceBrowser instance
            default_connection_type: Optional default for the connection type radio buttons
            model_status: Optional current model connection status (e.g., "Thin Report (Cloud)")
            on_auth_change: Optional callback to sync auth state with parent
            source_context: Optional source connection name for title context (e.g., "Budget Report (Local)")
            simple_mode: If True, hides perspective selector and connector type options (for Field Parameters)
        """
        super().__init__(parent)

        self.browser = browser
        self.result: Optional[SwapTarget] = None
        self.workspaces: List[WorkspaceInfo] = []
        self.current_datasets: List[DatasetInfo] = []
        self._all_workspaces: List[WorkspaceInfo] = []  # Full list for filtering
        self._default_connection_type = default_connection_type or CloudConnectionType.PBI_SEMANTIC_MODEL
        self._selected_dataset: Optional[DatasetInfo] = None  # Currently selected dataset for perspective loading
        self._model_status = model_status  # Store model connection status
        self._on_auth_change = on_auth_change  # Callback to sync auth state with parent
        self._source_context = source_context  # Source connection name for title context
        self._simple_mode = simple_mode  # Hide perspective/connector options for simpler use cases

        # Get theme manager for theme-aware colors
        self._theme_manager = get_theme_manager()

        self._setup_window()
        self._setup_ui()

        # Set modal focus after UI is ready
        self.grab_set()
        self.focus_set()

        # Check authentication and load data
        import logging
        logger = logging.getLogger(__name__)
        if not self.browser.is_authenticated():
            logger.info("CloudBrowserDialog: Not authenticated, starting auth flow")
            self._authenticate()
        else:
            # Already authenticated - update model status indicator
            self._update_model_status_display()
            self._update_cloud_auth_button_state()

            # Use cached data immediately if available (never show empty dialog)
            is_cached = self.browser.is_session_cached()
            ws_count = len(self.browser._workspaces) if hasattr(self.browser, '_workspaces') else 0
            logger.info(f"CloudBrowserDialog: is_session_cached={is_cached}, workspaces_count={ws_count}")
            if is_cached:
                # Populate from cache synchronously for instant display
                logger.info("CloudBrowserDialog: Populating from cache")
                self._populate_workspaces_from_cache()
            else:
                # Not cached - show loading and fetch
                logger.info("CloudBrowserDialog: Not cached, calling _load_workspaces")
                self.after(50, self._load_workspaces)

        # Note: Caller is responsible for calling wait_window(dialog)
        # Do NOT call self.wait_window() here - causes TclError when dialog is destroyed

    def _setup_window(self):
        """Configure the dialog window"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Build title with optional source context
        title = "Select Cloud Semantic Model"
        if self._source_context:
            title = f"{title}  |  Source: {self._source_context}"
        self.title(title)
        self.geometry("925x585")
        self.resizable(True, True)
        self.minsize(700, 485)
        self.configure(bg=colors['background'])

        # Modal behavior - must be set after window is configured
        self.transient(self.master)

        # Center on parent
        self.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() - 925) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - 585) // 2
        self.geometry(f"+{x}+{y}")

        # Set dialog icon (AE favicon)
        try:
            # From core/cloud/ -> core/ -> src/ (3 parents)
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
                                     'assets', 'favicon.ico')
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass  # Ignore if icon can't be loaded

        # Apply Windows dark/light mode title bar
        self._apply_title_bar_theme(is_dark)

    def _apply_title_bar_theme(self, is_dark: bool):
        """
        Apply dark or light mode to the Windows title bar.

        Uses Windows DWM API to set the DWMWA_USE_IMMERSIVE_DARK_MODE attribute.
        Only works on Windows 10 build 17763+ and Windows 11.
        """
        if sys.platform != 'win32':
            return

        try:
            # Get the window handle
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())

            # DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Windows 10 20H1+)
            # For older Windows 10 builds, attribute 19 was used
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if is_dark else 0)

            # Try the modern attribute first (Windows 10 20H1+)
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value),
                ctypes.sizeof(value)
            )

            # If that failed, try the older attribute (Windows 10 pre-20H1)
            if result != 0:
                DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_USE_IMMERSIVE_DARK_MODE_OLD,
                    ctypes.byref(value),
                    ctypes.sizeof(value)
                )
        except Exception:
            pass  # Silently fail on unsupported systems

    def _load_cloud_auth_icons(self):
        """Load cloud auth icons - colored for signed-in, red for signed-out."""
        self._cloud_icon_colored = None
        self._cloud_icon_gray = None

        if not PIL_AVAILABLE:
            return

        # Path from core/cloud/ -> core/ -> src/ (3 parents)
        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"

        try:
            if CAIROSVG_AVAILABLE:
                import cairosvg

                # Load signed-in icon (colored/green version)
                signed_in_path = icons_dir / "cloud login (signed in).svg"
                if signed_in_path.exists():
                    png_data = cairosvg.svg2png(
                        url=str(signed_in_path),
                        output_width=64,
                        output_height=64
                    )
                    img = Image.open(io.BytesIO(png_data))
                    colored_img = img.resize((16, 16), Image.Resampling.LANCZOS)
                    self._cloud_icon_colored = ImageTk.PhotoImage(colored_img)

                # Load signed-out icon (red version)
                signed_out_path = icons_dir / "cloud login (signed out).svg"
                if signed_out_path.exists():
                    png_data = cairosvg.svg2png(
                        url=str(signed_out_path),
                        output_width=64,
                        output_height=64
                    )
                    img = Image.open(io.BytesIO(png_data))
                    gray_img = img.resize((16, 16), Image.Resampling.LANCZOS)
                    self._cloud_icon_gray = ImageTk.PhotoImage(gray_img)

                # Fallback: if no icons loaded, try regular cloud login
                if not self._cloud_icon_colored or not self._cloud_icon_gray:
                    svg_path = icons_dir / "cloud login.svg"
                    if svg_path.exists():
                        png_data = cairosvg.svg2png(
                            url=str(svg_path),
                            output_width=64,
                            output_height=64
                        )
                        img = Image.open(io.BytesIO(png_data))
                        fallback_img = img.resize((16, 16), Image.Resampling.LANCZOS)
                        fallback_icon = ImageTk.PhotoImage(fallback_img)
                        if not self._cloud_icon_colored:
                            self._cloud_icon_colored = fallback_icon
                        if not self._cloud_icon_gray:
                            self._cloud_icon_gray = fallback_icon
        except Exception:
            pass  # Silently fail

    def _load_connect_icon(self) -> 'ImageTk.PhotoImage':
        """Load the connect icon for the Connect button in simple_mode."""
        if not PIL_AVAILABLE:
            return None

        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"

        try:
            if CAIROSVG_AVAILABLE:
                svg_path = icons_dir / "connect.svg"
                if svg_path.exists():
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=64,
                        output_height=64
                    )
                    img = Image.open(io.BytesIO(png_data))
                    resized = img.resize((14, 14), Image.Resampling.LANCZOS)
                    return ImageTk.PhotoImage(resized)
        except Exception:
            pass
        return None

    def _get_capacity_indicator_for_workspace(self, workspace_id: str) -> str:
        """Get capacity text indicator for a workspace by its ID.

        Args:
            workspace_id: The workspace GUID

        Returns:
            Text indicator like " ◆" for Premium, " ◇" for PPU, or "" for Pro
        """
        # Look up workspace in the browser's workspace list
        for ws in self.browser._workspaces:
            if ws.id == workspace_id and ws.capacity_type:
                if ws.capacity_type == 'Premium':
                    return " ◆"  # Filled diamond for Premium
                elif ws.capacity_type == 'PPU':
                    return " ◇"  # Outlined diamond for PPU
        return ""

    def _setup_ui(self):
        """Setup the dialog UI"""
        colors = self._theme_manager.colors
        self._colors = colors  # Store for use in auth callback
        is_dark = self._theme_manager.is_dark

        # Load cloud auth icons
        self._load_cloud_auth_icons()
        self._cloud_auth_dropdown = None  # Track dropdown popup

        # Main container with theme background
        main_frame = tk.Frame(self, bg=colors['background'], padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Store colors for child panels
        self._is_dark = is_dark

        # ===== Single combined controls row =====
        # Layout: Search: [entry] [Search] | All Recent Favorites Manual | [auth indicator] [Sign In]
        controls_frame = tk.Frame(main_frame, bg=colors['background'])
        controls_frame.pack(fill=tk.X, pady=(0, 12))

        # Left side: Search controls
        search_frame = tk.Frame(controls_frame, bg=colors['background'])
        search_frame.pack(side=tk.LEFT)

        tk.Label(
            search_frame, text="Search:",
            bg=colors['background'], fg=colors['text_primary'],
            font=("Segoe UI", 9)
        ).pack(side=tk.LEFT)

        self.search_var = tk.StringVar()
        self._search_clear_timer = None  # Timer for auto-clear delay
        self._last_search_results = []  # Cache of last search results for workspace filtering
        self._selected_workspace_id = None  # Currently selected workspace for filtering
        self._sort_column = "workspace"  # Default sort column
        self._sort_reverse = False  # Sort direction (False = ascending)

        # Trace search_var for auto-clear when emptied
        self.search_var.trace_add('write', self._on_search_var_changed)

        self._search_placeholder = "Search all workspaces..."
        self.search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=("Segoe UI", 10),
            bg=colors.get('input_bg', colors['surface']),
            fg=colors['text_muted'],  # Start with muted for placeholder
            insertbackground=colors['text_primary'],
            relief='flat',
            highlightthickness=1,
            highlightbackground=colors['border'],
            highlightcolor=colors['border'],  # Same as background border - no teal highlight
            width=22
        )
        self.search_entry.insert(0, self._search_placeholder)
        self.search_entry.pack(side=tk.LEFT, padx=(8, 8), ipady=4, ipadx=9)  # ipadx=9 for left/right text padding
        self.search_entry.bind("<Return>", lambda e: self._search_datasets())
        self.search_entry.bind("<FocusIn>", self._on_search_focus_in)
        self.search_entry.bind("<FocusOut>", self._on_search_focus_out)

        # Theme-aware disabled colors for buttons
        is_dark = colors.get('background', '') == '#0d0d1a'
        disabled_bg = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        disabled_fg = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        self.search_btn = RoundedButton(
            search_frame,
            text="Search",
            command=self._search_datasets,
            bg=colors['button_secondary'],
            fg=colors['text_primary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            width=70,
            height=30,
            radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background']
        )
        # Pack later - buttons go on RIGHT side

        # Refresh button to clear cache and reload
        self.refresh_btn = RoundedButton(
            search_frame,
            text="Refresh",
            command=self._refresh_data,
            bg=colors['button_secondary'],
            fg=colors['text_primary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            disabled_bg=disabled_bg,
            disabled_fg=disabled_fg,
            width=70,
            height=30,
            radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background']
        )
        # Pack buttons on RIGHT side (refresh first so it's rightmost, then search)
        self.refresh_btn.pack(side=tk.RIGHT)
        self.search_btn.pack(side=tk.RIGHT, padx=(0, 6))

        # Right side: Model status indicator + cloud auth icon button
        auth_container = tk.Frame(controls_frame, bg=colors['background'])
        auth_container.pack(side=tk.RIGHT)

        self.auth_status = tk.StringVar(value="Not Connected")
        self._auth_indicator = tk.Canvas(
            auth_container, width=10, height=10,
            bg=colors['background'], highlightthickness=0
        )
        self._auth_indicator.pack(side=tk.LEFT, padx=(0, 6))
        self._draw_auth_indicator(False)

        self.auth_label = tk.Label(
            auth_container,
            textvariable=self.auth_status,
            font=("Segoe UI", 9),
            bg=colors['background'],
            fg=colors['text_muted']
        )
        self.auth_label.pack(side=tk.LEFT, padx=(0, 8))

        # Cloud auth icon button with dropdown (matches main tool)
        initial_icon = self._cloud_icon_gray or self._cloud_icon_colored
        if initial_icon:
            self._cloud_auth_btn = SquareIconButton(
                auth_container,
                icon=initial_icon,
                command=self._toggle_cloud_auth_dropdown,
                tooltip_text="Cloud Account",
                size=26,
                radius=6
            )
            self._cloud_auth_btn.pack(side=tk.LEFT)

        # Middle: View mode radio buttons (All, Recent, Favorites, Manual)
        tab_frame = tk.Frame(controls_frame, bg=colors['background'])
        tab_frame.pack(side=tk.RIGHT, padx=(20, 20))

        self.view_mode = tk.StringVar(value="all")

        # Radio buttons for view mode
        self._view_mode_radio = LabeledRadioGroup(
            tab_frame,
            variable=self.view_mode,
            options=[
                ("all", "All"),
                ("recent", "Recent"),
                ("favorites", "Favorites"),
                ("manual", "Manual"),
            ],
            command=self._on_view_change,
            orientation="horizontal",
            font=("Segoe UI", 9),
            padding=10
        )
        self._view_mode_radio.pack(side=tk.LEFT)

        # Content area (switches based on tab)
        self.content_frame = tk.Frame(main_frame, bg=colors['background'])
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        # Create content panels
        self._create_browse_panel()
        self._create_manual_panel()

        # Initially show browse panel
        self.browse_frame.pack(fill=tk.BOTH, expand=True)

        # Perspective entry section (between content and action buttons)
        # Skip in simple_mode (e.g., Field Parameters tool doesn't need this)
        if not self._simple_mode:
            self._create_perspective_section(main_frame, colors)

        # Bottom row: Action buttons and selection info
        # In simple_mode, use fixed height since buttons are place-managed (don't affect parent sizing)
        row_height = 40 if self._simple_mode else None
        action_row = tk.Frame(main_frame, bg=colors['background'], height=row_height)
        action_row.pack(fill=tk.X, pady=(0, 5))
        if self._simple_mode:
            action_row.pack_propagate(False)  # Prevent pack from shrinking frame

        # Radio variable for connection type - always create for _on_select compatibility
        self.cloud_conn_type = tk.StringVar(value=self._default_connection_type.value)

        if self._simple_mode:
            # Simple mode: Buttons left-aligned, selection info right-aligned
            # Use grid layout with middle column expanding

            # Track whether selected item is Pro (no XMLA access)
            self._selected_is_pro = False

            # Configure grid: buttons fixed on left, spacer expands, selection fixed on right
            action_row.columnconfigure(0, weight=0)  # Buttons - fixed
            action_row.columnconfigure(1, weight=1)  # Spacer - expands
            action_row.columnconfigure(2, weight=0)  # Selection - fixed

            # Column 0: Buttons (left-aligned)
            button_frame = tk.Frame(action_row, bg=colors['background'])
            button_frame.grid(row=0, column=0, sticky='w')
            self._action_row = action_row

            # Load connect icon for the button
            self._connect_icon = self._load_connect_icon()

            self.select_btn = RoundedButton(
                button_frame,
                text="Connect",
                command=self._on_select,
                bg=colors['button_primary'],
                fg='#ffffff',
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                width=100,
                height=32,
                radius=6,
                canvas_bg=colors['background'],
                icon=self._connect_icon
            )
            self.select_btn.pack(side=tk.LEFT, padx=(0, 10))

            # Add tooltip for disabled state (will be updated when Pro workspace selected)
            self._connect_tooltip = None
            self._setup_connect_tooltip()

            self.cancel_btn = RoundedButton(
                button_frame,
                text="Cancel",
                command=self._on_cancel,
                bg=colors['button_secondary'],
                fg=colors['text_primary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                width=90,
                height=32,
                radius=6,
                canvas_bg=colors['background']
            )
            self.cancel_btn.pack(side=tk.LEFT)

            # Column 2: Selection info (right-aligned)
            selection_frame = tk.Frame(action_row, bg=colors['background'])
            selection_frame.grid(row=0, column=2, sticky='e')

            tk.Label(
                selection_frame, text="Selected:",
                bg=colors['background'], fg=colors['text_primary'],
                font=("Segoe UI", 9)
            ).pack(side=tk.LEFT, padx=(0, 5))

            self.selection_var = tk.StringVar(value="No selection")
            self._full_selection_text = "No selection"
            self._selection_label = tk.Label(
                selection_frame,
                textvariable=self.selection_var,
                font=("Segoe UI", 9, "bold"),
                bg=colors['background'],
                fg=colors['primary'],
                anchor='w'
            )
            self._selection_label.pack(side=tk.LEFT)
            self._selection_frame = selection_frame

            # Bind to configure event to handle text truncation when resizing
            action_row.bind('<Configure>', self._update_selection_truncation)

        else:
            # Full mode: Buttons on left, connector options in center, selection on right
            # Use grid layout for predictable column widths (pack caused shrinking issues)
            action_row.columnconfigure(2, weight=1)  # Selection column expands

            # Column 0: Buttons frame
            button_frame = tk.Frame(action_row, bg=colors['background'])
            button_frame.grid(row=0, column=0, sticky='w')

            self.select_btn = RoundedButton(
                button_frame,
                text="Select",
                command=self._on_select,
                bg=colors['button_primary'],
                fg='#ffffff',
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                width=90,
                height=32,
                radius=6,
                canvas_bg=colors['background']
            )
            self.select_btn.pack(side=tk.LEFT, padx=(0, 10))

            self.cancel_btn = RoundedButton(
                button_frame,
                text="Cancel",
                command=self._on_cancel,
                bg=colors['button_secondary'],
                fg=colors['text_primary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                width=90,
                height=32,
                radius=6,
                canvas_bg=colors['background']
            )
            self.cancel_btn.pack(side=tk.LEFT)

            # Column 1: Connection type toggle
            conn_type_frame = tk.Frame(action_row, bg=colors['background'])
            conn_type_frame.grid(row=0, column=1, sticky='w', padx=(25, 15))

            tk.Label(
                conn_type_frame,
                text="Connector:",
                font=("Segoe UI", 9),
                bg=colors['background'],
                fg=colors['text_secondary']
            ).pack(side=tk.LEFT, padx=(0, 8))

            # Radio buttons for connection type
            self._conn_type_radio = LabeledRadioGroup(
                conn_type_frame,
                variable=self.cloud_conn_type,
                options=[
                    (CloudConnectionType.PBI_SEMANTIC_MODEL.value, "Semantic Model"),
                    (CloudConnectionType.AAS_XMLA.value, "XMLA Endpoint"),
                ],
                orientation="horizontal",
                font=("Segoe UI", 9),
                padding=12
            )
            self._conn_type_radio.pack(side=tk.LEFT)

            # Column 2: Selection info (expands to fill remaining space)
            selection_frame = tk.Frame(action_row, bg=colors['background'])
            selection_frame.grid(row=0, column=2, sticky='ew')

            tk.Label(
                selection_frame, text="Selected:",
                bg=colors['background'], fg=colors['text_primary'],
                font=("Segoe UI", 9)
            ).pack(side=tk.LEFT, padx=(0, 5))

            self.selection_var = tk.StringVar(value="No selection")
            self._full_selection_text = "No selection"
            self._selection_label = tk.Label(
                selection_frame,
                textvariable=self.selection_var,
                font=("Segoe UI", 9, "bold"),
                bg=colors['background'],
                fg=colors['primary'],
                anchor='w'
            )
            self._selection_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Bind resize event to update truncation
            selection_frame.bind('<Configure>', self._on_selection_frame_resize)
            self._selection_frame = selection_frame

    def _safe_after(self, callback):
        """
        Schedule a callback on the main thread, safely handling dialog destruction.

        If the dialog has been closed/destroyed, this silently ignores the callback
        instead of raising an error.
        """
        try:
            if self.winfo_exists():
                self.after(0, callback)
        except tk.TclError:
            pass  # Dialog was destroyed, ignore

    def _on_selection_frame_resize(self, event=None):
        """Handle resize of selection frame to update truncation with ellipsis."""
        if not hasattr(self, '_full_selection_text') or not hasattr(self, '_selection_label'):
            return

        # Get available width for the selection text
        frame_width = self._selection_frame.winfo_width()
        # Account for "Selected:" label (~60px) and padding
        available_width = frame_width - 70

        if available_width <= 0:
            return

        # Get the font for measuring text
        font = ('Segoe UI', 9, 'bold')
        full_text = self._full_selection_text

        # Measure full text width
        try:
            temp_label = tk.Label(self._selection_frame, text=full_text, font=font)
            text_width = temp_label.winfo_reqwidth()
            temp_label.destroy()
        except Exception:
            return

        # If full text fits, show it
        if text_width <= available_width:
            self.selection_var.set(full_text)
            return

        # Otherwise, truncate with ellipsis
        ellipsis = "..."
        truncated = full_text

        # Binary search for optimal truncation point
        low, high = 0, len(full_text)
        while low < high:
            mid = (low + high + 1) // 2
            test_text = full_text[:mid] + ellipsis
            try:
                temp_label = tk.Label(self._selection_frame, text=test_text, font=font)
                test_width = temp_label.winfo_reqwidth()
                temp_label.destroy()
            except Exception:
                break

            if test_width <= available_width:
                low = mid
            else:
                high = mid - 1

        if low > 0:
            truncated = full_text[:low] + ellipsis
        else:
            truncated = ellipsis

        self.selection_var.set(truncated)

    def _update_selection_truncation(self, event=None):
        """Handle resize of action row in simple mode to enforce half-width constraint.

        The selection text should use at most half the container width.
        """
        if not hasattr(self, '_full_selection_text') or not hasattr(self, '_selection_label'):
            return
        if not hasattr(self, '_action_row'):
            return

        # Get container width
        container_width = self._action_row.winfo_width()
        if container_width <= 1:
            return

        # Calculate max width for selection: half of container minus padding
        max_selection_width = (container_width // 2) - 30  # 30px for padding/margins

        if max_selection_width <= 0:
            return

        # Get the font for measuring text
        font = ('Segoe UI', 9, 'bold')
        full_text = self._full_selection_text

        # Account for "Selected:" label width (~60px)
        available_text_width = max_selection_width - 60

        if available_text_width <= 0:
            self.selection_var.set("...")
            return

        # Measure full text width
        try:
            temp_label = tk.Label(self._selection_frame, text=full_text, font=font)
            text_width = temp_label.winfo_reqwidth()
            temp_label.destroy()
        except Exception:
            return

        # If full text fits within max width, show it
        if text_width <= available_text_width:
            self.selection_var.set(full_text)
            return

        # Otherwise, truncate with ellipsis using binary search
        ellipsis = "..."
        low, high = 0, len(full_text)
        while low < high:
            mid = (low + high + 1) // 2
            test_text = full_text[:mid] + ellipsis
            try:
                temp_label = tk.Label(self._selection_frame, text=test_text, font=font)
                test_width = temp_label.winfo_reqwidth()
                temp_label.destroy()
            except Exception:
                break

            if test_width <= available_text_width:
                low = mid
            else:
                high = mid - 1

        if low > 0:
            truncated = full_text[:low] + ellipsis
        else:
            truncated = ellipsis

        self.selection_var.set(truncated)

    def _update_selection_display(self, full_text: str):
        """Update the selection display with the full text, applying truncation as needed."""
        self._full_selection_text = full_text
        # Trigger appropriate resize handler based on mode
        if self._simple_mode and hasattr(self, '_action_row'):
            self._update_selection_truncation()
        else:
            self._on_selection_frame_resize()

    def _create_browse_panel(self):
        """Create the workspace/dataset browse panel"""
        colors = self._colors
        is_dark = self._is_dark

        self.browse_frame = tk.Frame(self.content_frame, bg=colors['background'])

        # Split into workspace list and dataset list
        self._paned = ttk.PanedWindow(self.browse_frame, orient=tk.HORIZONTAL)
        self._paned.pack(fill=tk.BOTH, expand=True)

        # Use same colors as main connection table for consistency
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
        section_bg = colors.get('section_bg', '#161627' if is_dark else '#f5f5f7')
        surface_color = colors.get('surface', '#1a1a2e' if is_dark else '#f5f5f7')

        # ===== Workspace list section =====
        ws_outer = tk.Frame(self._paned, bg=colors['background'])
        self._paned.add(ws_outer, weight=1)

        # Section header - match main table header design
        self._ws_header_label = tk.Label(
            ws_outer,
            text="Workspaces",
            font=("Segoe UI", 9, "bold"),
            bg=colors['background'],
            fg=colors.get('title_color', '#0084b7' if is_dark else '#009999')
        )
        self._ws_header_label.pack(anchor=tk.W, pady=(0, 6))

        # Container with visible border (wraps listbox + scrollbar)
        ws_container = tk.Frame(
            ws_outer,
            bg=section_bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=1
        )
        ws_container.pack(fill=tk.BOTH, expand=True)

        # Workspace Treeview with icons (dark blue selection, not teal)
        select_bg = '#1a3a5c' if is_dark else '#e6f3ff'  # Dark blue instead of teal

        # Create Treeview style for workspace list
        style = ttk.Style()
        ws_tree_style = "WorkspaceList.Treeview"
        style.configure(ws_tree_style,
                       background=section_bg,
                       foreground=colors['text_primary'],
                       fieldbackground=section_bg,
                       font=("Segoe UI", 9),
                       rowheight=22,
                       borderwidth=0)
        style.map(ws_tree_style,
                 background=[('selected', select_bg)],
                 foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])
        # Remove inner border by using minimal layout
        style.layout(ws_tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])

        # Two columns: star (fixed) and name (stretches)
        self.workspace_list = ttk.Treeview(
            ws_container,
            columns=("star", "name"),
            selectmode='browse',
            show='',  # Hide tree column and headings - just show data columns
            style=ws_tree_style
        )
        # Star column - width sized to center star between border and text
        self.workspace_list.column("star", width=24, minwidth=24, stretch=False, anchor='center')
        # Name column for workspace name - stretches to fill
        self.workspace_list.column("name", width=180, minwidth=100, stretch=True, anchor='w')
        self.workspace_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar area with visible background
        scrollbar_bg = '#1a1a2e' if is_dark else '#f0f0f0'
        ws_scrollbar_area = tk.Frame(ws_container, bg=scrollbar_bg, width=12)
        ws_scrollbar_area.pack(side=tk.RIGHT, fill=tk.Y)
        ws_scrollbar_area.pack_propagate(False)

        ws_scroll = ThemedScrollbar(
            ws_scrollbar_area,
            command=self.workspace_list.yview,
            theme_manager=self._theme_manager
        )
        self.workspace_list.configure(yscrollcommand=ws_scroll.set)
        ws_scroll.pack(fill=tk.BOTH, expand=True)

        self.workspace_list.bind("<<TreeviewSelect>>", self._on_workspace_select)
        self.workspace_list.bind("<Double-1>", self._on_workspace_favorite_toggle)
        self.workspace_list.bind("<Button-3>", self._on_workspace_right_click)  # Right-click context menu

        # ===== Semantic model list section =====
        ds_outer = tk.Frame(self._paned, bg=colors['background'])
        self._paned.add(ds_outer, weight=2)

        # Section header row - title on left, help text on right
        ds_header_row = tk.Frame(ds_outer, bg=colors['background'])
        ds_header_row.pack(fill=tk.X, pady=(0, 6))

        self._ds_header_label = tk.Label(
            ds_header_row,
            text="Semantic Models",
            font=("Segoe UI", 9, "bold"),
            bg=colors['background'],
            fg=colors.get('title_color', '#0084b7' if is_dark else '#009999')
        )
        self._ds_header_label.pack(side=tk.LEFT)

        # Help text in upper right - always visible
        # Explains double-click behavior, right-click favorite, and capacity symbols
        help_hint_color = colors.get('text_secondary', '#888888' if is_dark else '#666666')
        self._ds_help_label = tk.Label(
            ds_header_row,
            text="Double-click model to select  |  Right-click to favorite \u2606  |  \u25C6 Premium/Fabric  \u25C7 PPU",
            font=("Segoe UI", 8),
            bg=colors['background'],
            fg=help_hint_color
        )
        self._ds_help_label.pack(side=tk.RIGHT)

        # Container with visible border (wraps treeview + scrollbar)
        ds_container = tk.Frame(
            ds_outer,
            bg=section_bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=1
        )
        ds_container.pack(fill=tk.BOTH, expand=True)

        # Configure Treeview style - match main connection table exactly
        style = ttk.Style()
        tree_style = "CloudDialog.Treeview"
        # Use same heading colors as main connection table
        heading_bg = colors.get('section_bg', '#161627' if is_dark else '#f5f5f7')
        heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        row_bg = section_bg
        row_fg = colors['text_primary']
        selected_bg = '#1a3a5c' if is_dark else '#e6f3ff'  # Dark blue selection (matching workspace)
        header_separator = '#0d0d1a' if is_dark else '#ffffff'  # Match main table separator

        style.configure(tree_style,
                       background=row_bg,
                       foreground=row_fg,
                       fieldbackground=row_bg,
                       rowheight=28,
                       borderwidth=0,
                       relief='flat',
                       font=('Segoe UI', 9))
        # Remove internal treeview borders via layout
        style.layout(tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])
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
        # Custom heading layout to ensure background colors are applied consistently
        style.layout(f"{tree_style}.Heading", [
            ('Treeheading.cell', {'sticky': 'nswe'}),
            ('Treeheading.border', {'sticky': 'nswe', 'children': [
                ('Treeheading.padding', {'sticky': 'nswe', 'children': [
                    ('Treeheading.image', {'side': 'left', 'sticky': ''}),
                    ('Treeheading.text', {'sticky': 'we'})
                ]})
            ]})
        ])
        style.map(tree_style,
                  background=[('selected', selected_bg)],
                  foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])  # Theme-aware text
        # Ensure heading text remains visible on all states (hover, pressed, normal)
        style.map(f"{tree_style}.Heading",
                  background=[('active', heading_bg), ('pressed', heading_bg), ('!active', heading_bg)],
                  foreground=[('active', heading_fg), ('pressed', heading_fg), ('!active', heading_fg), ('', heading_fg)],
                  relief=[('active', 'groove'), ('pressed', 'groove'), ('!active', 'groove')])

        columns = ("name", "workspace", "configured_by")
        self.dataset_tree = ttk.Treeview(
            ds_container,
            columns=columns,
            show="headings",
            selectmode="browse",
            style=tree_style
        )

        self.dataset_tree.heading("name", text="Semantic Model", command=lambda: self._sort_by_column("name"))
        self.dataset_tree.heading("workspace", text="Workspace", command=lambda: self._sort_by_column("workspace"))
        self.dataset_tree.heading("configured_by", text="Configured By", command=lambda: self._sort_by_column("configured_by"))

        self.dataset_tree.column("name", width=200, minwidth=100, anchor="center")
        self.dataset_tree.column("workspace", width=150, minwidth=100, anchor="center")
        self.dataset_tree.column("configured_by", width=120, minwidth=80, anchor="center")

        # Scrollbar area with visible background (inside border)
        ds_scrollbar_area = tk.Frame(ds_container, bg=scrollbar_bg, width=12)
        ds_scrollbar_area.pack(side=tk.RIGHT, fill=tk.Y)
        ds_scrollbar_area.pack_propagate(False)

        ds_scroll = ThemedScrollbar(
            ds_scrollbar_area,
            command=self.dataset_tree.yview,
            theme_manager=self._theme_manager
        )
        self.dataset_tree.configure(yscrollcommand=ds_scroll.set)

        self.dataset_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ds_scroll.pack(fill=tk.BOTH, expand=True)

        # Configure tag for Pro workspaces (disabled/grayed out appearance)
        # Only needed in simple_mode where Pro workspaces can't be used
        if self._simple_mode:
            pro_text_color = colors.get('text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')
            self.dataset_tree.tag_configure('pro_workspace', foreground=pro_text_color)

        self.dataset_tree.bind("<<TreeviewSelect>>", self._on_dataset_select)
        self.dataset_tree.bind("<Double-1>", self._on_dataset_double_click)
        self.dataset_tree.bind("<Button-3>", self._on_dataset_right_click)  # Right-click for context menu

    def _create_manual_panel(self):
        """Create the manual entry panel"""
        colors = self._colors
        is_dark = self._is_dark

        self.manual_frame = tk.Frame(self.content_frame, bg=colors['background'])

        # Instructions
        tk.Label(
            self.manual_frame,
            text="Enter XMLA endpoint and semantic model name manually:",
            font=("Segoe UI", 10),
            bg=colors['background'],
            fg=colors['text_primary']
        ).pack(anchor="w", pady=(0, 15))

        # Entry styling helper
        entry_bg = colors.get('input_bg', colors['surface'])
        entry_fg = colors['text_primary']

        # Workspace/XMLA entry
        row1 = tk.Frame(self.manual_frame, bg=colors['background'])
        row1.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            row1, text="Workspace or XMLA Endpoint:",
            width=25, anchor='w',
            bg=colors['background'], fg=colors['text_primary'],
            font=("Segoe UI", 9)
        ).pack(side=tk.LEFT)

        self.manual_workspace = tk.StringVar()
        tk.Entry(
            row1,
            textvariable=self.manual_workspace,
            font=("Segoe UI", 10),
            bg=entry_bg,
            fg=entry_fg,
            insertbackground=entry_fg,
            relief='flat',
            highlightthickness=1,
            highlightbackground=colors['border'],
            highlightcolor=colors['primary']
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)

        # Semantic model name entry
        row2 = tk.Frame(self.manual_frame, bg=colors['background'])
        row2.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            row2, text="Semantic Model:",
            width=25, anchor='w',
            bg=colors['background'], fg=colors['text_primary'],
            font=("Segoe UI", 9)
        ).pack(side=tk.LEFT)

        self.manual_dataset = tk.StringVar()
        tk.Entry(
            row2,
            textvariable=self.manual_dataset,
            font=("Segoe UI", 10),
            bg=entry_bg,
            fg=entry_fg,
            insertbackground=entry_fg,
            relief='flat',
            highlightthickness=1,
            highlightbackground=colors['border'],
            highlightcolor=colors['primary']
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)

        # Help text
        help_text = """Examples:
  - Workspace name: "My Workspace"
  - Full XMLA endpoint: powerbi://api.powerbi.com/v1.0/myorg/My%20Workspace

The workspace name will be automatically converted to an XMLA endpoint."""

        tk.Label(
            self.manual_frame,
            text=help_text,
            font=("Segoe UI", 8),
            bg=colors['background'],
            fg=colors['text_muted'],
            justify=tk.LEFT
        ).pack(anchor="w", pady=(15, 0))

    def _create_perspective_section(self, parent, colors):
        """Create the perspective entry section with dropdown and manual entry."""
        is_dark = self._is_dark

        perspective_frame = tk.Frame(parent, bg=colors['background'])
        perspective_frame.pack(fill=tk.X, pady=(9, 12))  # Reduced top padding by 3px

        # Left side: Label
        tk.Label(
            perspective_frame,
            text="Perspective (optional):",
            font=("Segoe UI", 9),
            bg=colors['background'],
            fg=colors['text_secondary']
        ).pack(side=tk.LEFT, padx=(0, 8))

        # Custom dropdown using tk.Entry for full cursor control (ttk.Combobox cursor styling
        # is broken on Windows). This gives us reliable insertbackground support.
        entry_bg = colors.get('input_bg', colors['surface'])
        entry_fg = colors['text_primary']
        border_color = colors['button_primary']  # Blue in dark, teal in light

        # Container for entry + dropdown button
        self._combo_container = tk.Frame(perspective_frame, bg=border_color)
        self._combo_container.pack(side=tk.LEFT, padx=(0, 12))

        # Inner frame for the entry and button (provides visual border)
        combo_inner = tk.Frame(self._combo_container, bg=entry_bg)
        combo_inner.pack(padx=1, pady=1)

        self.perspective_var = tk.StringVar()
        self._perspective_entry = tk.Entry(
            combo_inner,
            textvariable=self.perspective_var,
            font=("Segoe UI", 9),
            width=28,
            bg=entry_bg,
            fg=entry_fg,
            insertbackground=entry_fg,  # White cursor in dark mode
            relief='flat',
            highlightthickness=0,
            bd=0
        )
        self._perspective_entry.pack(side=tk.LEFT, padx=(4, 0), pady=2)

        # Dropdown button (arrow)
        arrow_char = "\u25BC"  # Down arrow
        self._dropdown_btn = tk.Label(
            combo_inner,
            text=arrow_char,
            font=("Segoe UI", 8),
            bg=entry_bg,
            fg=colors['text_secondary'],
            cursor='hand2',
            padx=6
        )
        self._dropdown_btn.pack(side=tk.LEFT, pady=2)
        self._dropdown_btn.bind('<Button-1>', self._toggle_perspective_dropdown)
        self._perspective_entry.bind('<Alt-Down>', self._toggle_perspective_dropdown)

        # Store values for dropdown
        self._perspective_values = []
        self._dropdown_popup = None

        # Compatibility property for existing code that uses self._perspective_combo['values']
        class ComboShim:
            """Shim to maintain compatibility with ttk.Combobox-style access."""
            def __init__(shim_self, entry, values_ref, parent_dialog):
                shim_self._entry = entry
                shim_self._values = values_ref
                shim_self._parent = parent_dialog

            def __getitem__(shim_self, key):
                if key == 'values':
                    return shim_self._parent._perspective_values
                return None

            def __setitem__(shim_self, key, value):
                if key == 'values':
                    shim_self._parent._perspective_values = list(value) if value else []

        self._perspective_combo = ComboShim(self._perspective_entry, self._perspective_values, self)

        # Hint text (changes based on workspace type and perspective availability)
        hint_color = colors.get('text_muted', '#888888')
        self._perspective_hint = tk.Label(
            perspective_frame,
            text="Select a dataset to see available perspectives",
            font=("Segoe UI", 8),
            bg=colors['background'],
            fg=hint_color
        )
        self._perspective_hint.pack(side=tk.LEFT)

        # Loading indicator (hidden by default)
        self._perspective_loading = tk.Label(
            perspective_frame,
            text="Loading perspectives...",
            font=("Segoe UI", 8, "italic"),
            bg=colors['background'],
            fg=colors['primary']
        )
        # Initially hidden - will be shown during loading

        # Store colors for dropdown popup
        self._dropdown_colors = {
            'bg': colors.get('card_surface', colors['surface']),
            'fg': colors['text_primary'],
            'select_bg': colors['button_primary'],
            'select_fg': '#ffffff',
            'border': colors['button_primary']
        }

    def _toggle_perspective_dropdown(self, event=None):
        """Toggle the perspective dropdown popup."""
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._close_perspective_dropdown()
            return

        # Don't show dropdown if no values
        if not self._perspective_values:
            return

        self._show_perspective_dropdown()

    def _show_perspective_dropdown(self):
        """Show the perspective dropdown popup listbox."""
        colors = self._dropdown_colors

        # Change arrow to point up (indicates dropdown is open)
        self._dropdown_btn.configure(text="\u25B2")

        # Create popup toplevel
        self._dropdown_popup = tk.Toplevel(self)
        self._dropdown_popup.withdraw()
        self._dropdown_popup.overrideredirect(True)

        # Create listbox frame with border
        frame = tk.Frame(self._dropdown_popup, bg=colors['border'])
        frame.pack(fill=tk.BOTH, expand=True)

        inner = tk.Frame(frame, bg=colors['bg'])
        inner.pack(padx=1, pady=1, fill=tk.BOTH, expand=True)

        # Calculate height based on items (max 8 visible)
        num_items = min(len(self._perspective_values), 8)
        listbox_height = num_items

        self._dropdown_listbox = tk.Listbox(
            inner,
            font=("Segoe UI", 9),
            bg=colors['bg'],
            fg=colors['fg'],
            selectbackground=colors['select_bg'],
            selectforeground=colors['select_fg'],
            relief='flat',
            bd=0,
            highlightthickness=0,
            height=listbox_height,
            activestyle='none'
        )
        self._dropdown_listbox.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # Populate listbox
        for value in self._perspective_values:
            self._dropdown_listbox.insert(tk.END, value)

        # Select current value if it exists
        current = self.perspective_var.get()
        if current in self._perspective_values:
            idx = self._perspective_values.index(current)
            self._dropdown_listbox.selection_set(idx)
            self._dropdown_listbox.see(idx)

        # Bind selection
        self._dropdown_listbox.bind('<ButtonRelease-1>', self._on_dropdown_select)
        self._dropdown_listbox.bind('<Return>', self._on_dropdown_select)
        self._dropdown_listbox.bind('<Escape>', lambda e: self._close_perspective_dropdown())

        # Position below the combo container (aligns dropdown with the full input box)
        self._combo_container.update_idletasks()
        x = self._combo_container.winfo_rootx()
        y = self._combo_container.winfo_rooty() + self._combo_container.winfo_height() + 2
        width = self._combo_container.winfo_width()

        self._dropdown_popup.geometry(f"{width}x{self._dropdown_listbox.winfo_reqheight() + 6}+{x}+{y}")
        self._dropdown_popup.deiconify()
        self._dropdown_listbox.focus_set()

        # Close on click outside
        self._dropdown_popup.bind('<FocusOut>', self._on_dropdown_focus_out)

    def _on_dropdown_select(self, event=None):
        """Handle selection from dropdown listbox."""
        selection = self._dropdown_listbox.curselection()
        if selection:
            idx = selection[0]
            value = self._perspective_values[idx]
            self.perspective_var.set(value)
        self._close_perspective_dropdown()

    def _on_dropdown_focus_out(self, event=None):
        """Close dropdown when focus leaves."""
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            # Small delay to allow click on listbox items to register
            self._dropdown_popup.after(100, self._check_dropdown_focus)

    def _check_dropdown_focus(self):
        """Check if focus is still in dropdown, close if not."""
        if not self._dropdown_popup or not self._dropdown_popup.winfo_exists():
            return
        try:
            focused = self.focus_get()
            if focused != self._dropdown_listbox:
                self._close_perspective_dropdown()
        except Exception:
            self._close_perspective_dropdown()

    def _close_perspective_dropdown(self):
        """Close the perspective dropdown popup."""
        # Change arrow back to point down (indicates dropdown is closed)
        self._dropdown_btn.configure(text="\u25BC")

        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._dropdown_popup.destroy()
        self._dropdown_popup = None

    def _update_perspective_dropdown(self, perspectives: List[str], is_premium: bool = True):
        """Update the perspective dropdown with available perspectives."""
        self._perspective_combo['values'] = perspectives
        # Store for validation when user manually types a name
        self._current_perspectives = perspectives
        self._can_validate_perspectives = is_premium  # True if we could fetch (even if empty)

        # Update hint text based on what's available
        if perspectives:
            self._perspective_hint.configure(
                text=f"{len(perspectives)} perspectives"
            )
        elif is_premium:
            self._perspective_hint.configure(
                text="No perspectives defined"
            )
        else:
            # Pro workspace - can't validate manual entries
            self._perspective_hint.configure(
                text="XMLA access required. Manual entries cannot be validated."
            )

    def _on_dataset_select_perspectives(self, dataset: DatasetInfo):
        """
        Handle perspective loading when a dataset is selected.

        For Premium workspaces, fetches perspectives via XMLA/TOM.
        For Pro workspaces, shows hint that manual entry is available.
        """
        self._selected_dataset = dataset

        # Skip perspective handling in simple_mode (no perspective UI exists)
        if self._simple_mode:
            return

        # Clear any previous perspective selection
        self.perspective_var.set("")
        self._perspective_combo['values'] = []

        # Check workspace capacity_type to immediately update XMLA option state
        # This gives instant feedback before perspective loading completes
        workspace_capacity_type = None
        for ws in self.browser._workspaces:
            if ws.name == dataset.workspace_name:
                workspace_capacity_type = ws.capacity_type
                break

        # If we know it's a Pro workspace (capacity_type is None), immediately disable XMLA
        # Note: We'll re-check after perspective loading in case XMLA access works (PPU detection)
        if workspace_capacity_type is None:
            self._update_xmla_option_state(False)
        else:
            self._update_xmla_option_state(True)

        # Check if perspectives are already loaded
        if dataset.perspectives_loaded:
            has_xmla_access = workspace_capacity_type is not None
            self._update_perspective_dropdown(dataset.perspectives or [], has_xmla_access)
            return

        # Try to load perspectives in background
        self._perspective_hint.pack_forget()
        self._perspective_loading.pack(side=tk.LEFT)

        def load_perspectives():
            try:
                perspectives, error = self.browser.get_dataset_perspectives(
                    dataset.workspace_name,
                    dataset.name
                )
                # Update the dataset with loaded perspectives
                dataset.perspectives = perspectives if perspectives else []
                dataset.perspectives_loaded = True

                # Determine if Premium/PPU based on workspace capacity_type after fetch
                # If XMLA access worked, browser sets capacity_type to Premium or PPU
                # If XMLA failed (403 Pro workspace), capacity_type remains None
                has_xmla_access = False
                for ws in self.browser._workspaces:
                    if ws.name == dataset.workspace_name:
                        has_xmla_access = ws.capacity_type is not None
                        break

                self._safe_after(lambda: self._on_perspectives_loaded(perspectives or [], has_xmla_access))
            except Exception as e:
                self._safe_after(lambda: self._on_perspectives_loaded([], False))

        import threading
        threading.Thread(target=load_perspectives, daemon=True).start()

    def _on_perspectives_loaded(self, perspectives: List[str], is_premium: bool):
        """Handle perspectives loaded callback on main thread."""
        # Hide loading, show hint
        self._perspective_loading.pack_forget()
        self._perspective_hint.pack(side=tk.LEFT)

        self._update_perspective_dropdown(perspectives, is_premium)

        # Enable/disable XMLA Endpoint option based on XMLA access
        self._update_xmla_option_state(is_premium)

        # Check if PPU was detected during perspective fetch (browser updates workspace)
        # Refresh displays to show updated capacity indicators
        self._refresh_capacity_indicators()

    def _update_xmla_option_state(self, has_xmla_access: bool):
        """
        Enable or disable the XMLA Endpoint connector option based on XMLA access.

        Pro workspaces do not support XMLA endpoints and require the Semantic Model
        connector. PPU and Premium/Fabric workspaces support both connectors.

        Args:
            has_xmla_access: True if the workspace supports XMLA endpoint connections
        """
        # Skip if in simple_mode (no connector radio exists)
        if not hasattr(self, '_conn_type_radio'):
            return

        xmla_value = CloudConnectionType.AAS_XMLA.value

        if has_xmla_access:
            # Enable XMLA option
            self._conn_type_radio.set_option_enabled(xmla_value, True)
        else:
            # Disable XMLA option with tooltip explaining why
            # Also force selection to Semantic Model if currently on XMLA
            if self.cloud_conn_type.get() == xmla_value:
                self.cloud_conn_type.set(CloudConnectionType.PBI_SEMANTIC_MODEL.value)

            tooltip = "XMLA endpoint not available for Pro workspaces.\nRequires Premium, PPU, or Fabric capacity."
            self._conn_type_radio.set_option_enabled(xmla_value, False, tooltip)

    def _draw_auth_indicator(self, connected: bool):
        """Draw the model connection status indicator dot."""
        self._auth_indicator.delete("all")
        color = '#4ade80' if connected else '#ef4444'  # Green / Red
        self._auth_indicator.create_oval(1, 1, 9, 9, fill=color, outline=color)

    def _update_model_status_display(self):
        """Update the model connection status label and indicator."""
        if self._model_status:
            # Model is connected - show the status
            self.auth_status.set(self._model_status)
            self._draw_auth_indicator(True)
            self.auth_label.configure(fg=self._colors.get('success', '#4ade80'))
        elif self._simple_mode and self.browser and self.browser.is_authenticated():
            # Simple mode with cloud auth - show authenticated status
            self.auth_status.set("Signed In")
            self._draw_auth_indicator(True)
            self.auth_label.configure(fg=self._colors.get('success', '#4ade80'))
        else:
            # No model connected / not authenticated
            self.auth_status.set("Not Connected")
            self._draw_auth_indicator(False)
            self.auth_label.configure(fg=self._colors.get('text_muted', '#888888'))

    def _update_cloud_auth_button_state(self):
        """Update the cloud auth button icon based on auth state."""
        if not hasattr(self, '_cloud_auth_btn'):
            return

        is_authenticated = self.browser.is_authenticated() if self.browser else False

        if is_authenticated:
            if self._cloud_icon_colored:
                self._cloud_auth_btn.update_icon(self._cloud_icon_colored)
            email = self.browser.get_account_email() if self.browser else None
            tooltip = f"Cloud Account: {email}" if email else "Cloud Account"
        else:
            if self._cloud_icon_gray:
                self._cloud_auth_btn.update_icon(self._cloud_icon_gray)
            tooltip = "Cloud Account (Not signed in)"

        self._cloud_auth_btn._tooltip_text = tooltip

        # Notify parent to sync its auth button state
        if self._on_auth_change:
            try:
                self._on_auth_change()
            except Exception:
                pass  # Don't let callback errors affect dialog

    def _toggle_cloud_auth_dropdown(self):
        """Toggle the cloud authentication dropdown."""
        if self._cloud_auth_dropdown and self._cloud_auth_dropdown.winfo_exists():
            self._close_cloud_auth_dropdown()
        else:
            self._open_cloud_auth_dropdown()

    def _open_cloud_auth_dropdown(self):
        """Open the cloud auth dropdown menu."""
        colors = self._colors

        # Create popup window
        popup = tk.Toplevel(self)
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
        email = self.browser.get_account_email() if self.browser else None
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
        self.bind('<Button-1>', self._cloud_auth_click_handler, add='+')

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
                self.unbind('<Button-1>', self._cloud_auth_click_handler)
            except Exception:
                pass

    def _on_cloud_auth_action(self, is_signed_in: bool):
        """Handle sign in or sign out action."""
        self._close_cloud_auth_dropdown()

        if is_signed_in:
            # Sign out
            if self.browser:
                self.browser.sign_out()
            self.auth_status.set("Not Connected")
            self._draw_auth_indicator(False)
            self.auth_label.configure(fg=self._colors['text_muted'])
            self._update_cloud_auth_button_state()
            # Clear lists
            self.workspace_list.delete(*self.workspace_list.get_children())
            self.dataset_list.delete(*self.dataset_list.get_children())
        else:
            # Sign in
            self._authenticate()

    def _on_view_change(self):
        """Handle view mode change"""
        mode = self.view_mode.get()

        # Hide all panels
        self.browse_frame.pack_forget()
        self.manual_frame.pack_forget()

        if mode == "manual":
            self.manual_frame.pack(fill=tk.BOTH, expand=True)
        else:
            self.browse_frame.pack(fill=tk.BOTH, expand=True)
            self._load_workspaces(mode)

    def _authenticate(self):
        """Authenticate with Azure AD"""
        self.auth_status.set("Authenticating...")

        def auth_thread():
            try:
                success, message = self.browser.authenticate()
                self._safe_after(lambda: self._handle_auth_result(success, message))
            except Exception as e:
                self._safe_after(lambda: self._handle_auth_result(False, f"Authentication error: {e}"))

        threading.Thread(target=auth_thread, daemon=True).start()

    def _handle_auth_result(self, success: bool, message: str):
        """Handle authentication result"""
        if success:
            # Update model status display (shows model connection, not auth status)
            self._update_model_status_display()
            self._update_cloud_auth_button_state()
            self._load_workspaces()
        else:
            self.auth_status.set("Auth Failed")
            self._draw_auth_indicator(False)
            self._update_cloud_auth_button_state()
            ThemedMessageBox.showerror(self, "Authentication Failed", message)

    def _load_workspaces(self, filter_type: str = "all"):
        """Load workspaces"""
        if not self.browser.is_authenticated():
            return

        self.workspace_list.delete(*self.workspace_list.get_children())

        def load_thread():
            try:
                workspaces, error = self.browser.get_workspaces(filter_type)
                self._safe_after(lambda: self._populate_workspaces(workspaces, error))

                # Preload all datasets in background for fast searching
                if not error and filter_type == "all" and not self.browser.is_fully_cached():
                    self._safe_after(lambda: self._start_preload())
            except Exception as e:
                self._safe_after(lambda: self._populate_workspaces([], f"Failed to load workspaces: {e}"))

        threading.Thread(target=load_thread, daemon=True).start()

    def _start_preload(self):
        """Start preloading all datasets in background."""
        self.search_btn.set_enabled(False)
        self.search_btn.update_text("Loading...")

        def preload_thread():
            try:
                def progress_callback(current, total, ws_name):
                    # Update UI on main thread, safely handling dialog closure
                    self._safe_after(lambda c=current, t=total: self._update_preload_progress(c, t))

                total_datasets, error = self.browser.preload_all_datasets(progress_callback)
                self._safe_after(lambda: self._on_preload_complete(total_datasets, error))
            except Exception as e:
                self._safe_after(lambda: self._on_preload_complete(0, f"Preload failed: {e}"))

        threading.Thread(target=preload_thread, daemon=True).start()

    def _update_preload_progress(self, current: int, total: int):
        """Update the search button with preload progress and show datasets progressively."""
        # Check if dialog is still alive before updating UI
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return

        if total > 0:
            percent = int((current / total) * 100)
            self.search_btn.update_text(f"{percent}%")

            # Progressively show datasets as they load (if no workspace selected)
            # Update every ~10% or every workspace if few workspaces
            update_interval = max(1, total // 10)
            if not self._selected_workspace_id and (current % update_interval == 0 or current == total):
                self._show_all_datasets()

    def _on_preload_complete(self, total_datasets: int, error: Optional[str]):
        """Handle preload completion."""
        # Check if dialog is still alive before updating UI
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return

        self.search_btn.set_enabled(True)
        self.search_btn.update_text("Search")

        # Show all datasets when preload completes (if no workspace selected)
        if not error and not self._selected_workspace_id:
            self._show_all_datasets()

    def _show_all_datasets(self):
        """
        Show all datasets from all workspaces, respecting current sort order.

        This displays all cached datasets in the semantic model list.
        Maintains the user's current sort column and direction.
        Note: _populate_datasets handles favorites filtering when in that view mode.
        """
        all_datasets = []

        # Get all datasets from cache
        for ws in self._all_workspaces:
            ws_datasets = self.browser._datasets_cache.get(ws.id, [])
            for ds in ws_datasets:
                all_datasets.append(ds)

        # Apply current sort order
        key_func = self._get_sort_key_func(self._sort_column)
        if key_func:
            all_datasets = sorted(all_datasets, key=key_func, reverse=self._sort_reverse)

        self._populate_datasets(all_datasets, None)

    def _get_sort_key_func(self, column: str):
        """Get the sort key function for a given column."""
        if column == "name":
            return lambda ds: (ds.name or "").lower()
        elif column == "workspace":
            return lambda ds: (ds.workspace_name or "").lower()
        elif column == "configured_by":
            return lambda ds: (ds.configured_by or "").lower()
        return None

    def _refresh_data(self):
        """Clear cache and reload all data."""
        self.browser.clear_cache()
        self.workspace_list.delete(*self.workspace_list.get_children())
        self.dataset_tree.delete(*self.dataset_tree.get_children())
        self._load_workspaces()

    def _populate_workspaces_from_cache(self):
        """Populate workspace list from browser's cached data immediately."""
        workspaces, _ = self.browser.get_workspaces("all")
        self._populate_workspaces(workspaces, None)

        # Start preload if not already running (datasets may not be cached yet)
        if not self.browser.is_fully_cached():
            self.after(100, self._start_preload)

    def _populate_workspaces(self, workspaces: List[WorkspaceInfo], error: Optional[str]):
        """Populate workspace list"""
        if error:
            ThemedMessageBox.showerror(self, "Error", error)
            return

        self.workspaces = workspaces
        self._all_workspaces = workspaces  # Keep full list for filtering
        self._refresh_workspace_list()

        # Default select "(All Workspaces)" entry
        children = self.workspace_list.get_children()
        if children:
            first_item = children[0]
            self.workspace_list.selection_set(first_item)
            self.workspace_list.see(first_item)

        # Adjust sash position: reduce workspace section by 15px
        # Schedule after idle to ensure layout is complete
        self.after_idle(self._adjust_paned_sash)

    def _adjust_paned_sash(self):
        """Adjust the paned window sash to reduce workspace section by 15px (once only)."""
        try:
            # Only adjust once on first load
            if getattr(self, '_sash_adjusted', False):
                return
            self._sash_adjusted = True

            if not hasattr(self, '_paned') or not self._paned.winfo_exists():
                return

            # Get current sash position (default is based on weight ratio 1:2)
            # For a 925px dialog with ~40px padding, content is ~885px
            # Default 1:2 ratio would put sash at ~295px
            # We want to shift it 15px left (workspace smaller, datasets larger)
            current_pos = self._paned.sashpos(0)
            if current_pos > 15:
                new_pos = current_pos - 15
                self._paned.sashpos(0, new_pos)
        except Exception:
            pass  # Ignore if sash adjustment fails

    def _refresh_capacity_indicators(self):
        """
        Refresh capacity indicators in workspace list and dataset table.

        Called after perspective loading when PPU might have been detected.
        Updates the display to include the correct capacity indicator.
        """
        # Update workspace list entries - values=(star, name_with_capacity)
        for item in self.workspace_list.get_children():
            item_values = self.workspace_list.item(item, 'values')
            if not item_values or len(item_values) < 2:
                continue
            item_star = item_values[0]
            item_name = item_values[1]
            # Skip "(All Workspaces)" entry
            if "(All Workspaces)" in item_name:
                continue

            # Find matching workspace by name (strip capacity indicator from name)
            base_name = item_name.rstrip(" \u25C6\u25C7").strip()
            for ws in self._all_workspaces:
                if ws.name == base_name:
                    # Build new values with name + capacity indicator
                    cap_indicator = ""
                    if ws.capacity_type == 'Premium':
                        cap_indicator = " \u25C6"
                    elif ws.capacity_type == 'PPU':
                        cap_indicator = " \u25C7"
                    new_name = f"{ws.name}{cap_indicator}"
                    if new_name != item_name:
                        self.workspace_list.item(item, values=(item_star, new_name))
                    break

        # Update dataset table entries
        if hasattr(self, 'current_datasets') and self.current_datasets:
            for item in self.dataset_tree.get_children():
                values = self.dataset_tree.item(item, 'values')
                if not values or len(values) < 2:
                    continue

                idx = int(item)
                if idx >= len(self.current_datasets):
                    continue

                ds = self.current_datasets[idx]
                cap_indicator = self._get_capacity_indicator_for_workspace(ds.workspace_id)
                workspace_display = (ds.workspace_name or "") + cap_indicator

                # Only update if changed
                if values[1] != workspace_display:
                    new_values = (values[0], workspace_display, values[2] if len(values) > 2 else "")
                    self.dataset_tree.item(item, values=new_values)

    def _refresh_workspace_list(self, filter_query: str = ""):
        """Refresh workspace list with optional filter by workspace name.

        Two-column layout: values=(star, name_with_capacity)
        """
        self.workspace_list.delete(*self.workspace_list.get_children())

        query = filter_query.lower()
        filtered = [ws for ws in self._all_workspaces if query in ws.name.lower()] if query else self._all_workspaces
        self.workspaces = filtered

        # Add "All Workspaces" option at the top (only when not filtering)
        if not query:
            self.workspace_list.insert('', 'end', values=("", "(All Workspaces)"), tags=('all',))

        for ws in filtered:
            # Star indicator: ★ if all models favorited, ☆ if some, empty if none
            star = self._get_workspace_star(ws.id)
            # Get capacity text indicator (trailing the name)
            cap_indicator = ""
            if ws.capacity_type == 'Premium':
                cap_indicator = " \u25C6"  # Filled diamond
            elif ws.capacity_type == 'PPU':
                cap_indicator = " \u25C7"  # Outlined diamond
            self.workspace_list.insert('', 'end', values=(star, f"{ws.name}{cap_indicator}"), tags=('ws',))

    def _get_workspace_star(self, workspace_id: str) -> str:
        """Get star indicator for workspace based on model favorites.

        Returns:
            ★ if all models favorited, ☆ if some, empty if none
        """
        fav_status = self.browser.get_workspace_favorite_status(workspace_id)
        if fav_status == 'all':
            return "\u2605"  # ★ filled star - all models favorited
        elif fav_status == 'some':
            return "\u2606"  # ☆ outline star - some models favorited
        return ""  # No star - no models favorited

    def _refresh_favorites_workspace_list(self):
        """Refresh workspace list in favorites mode.

        Only shows workspaces that have at least one individually favorited model.
        """
        self.workspace_list.delete(*self.workspace_list.get_children())

        # Filter to workspaces with at least one favorited model
        favorite_workspaces = [ws for ws in self._all_workspaces
                               if self.browser.get_workspace_favorite_status(ws.id) != 'none']

        # Update the active workspaces list
        self.workspaces = favorite_workspaces

        # Add "All Workspaces" option at the top
        self.workspace_list.insert('', 'end', values=("", "(All Workspaces)"), tags=('all',))

        for ws in favorite_workspaces:
            # Star indicator: ★ if all models favorited, ☆ if some
            star = self._get_workspace_star(ws.id)
            # Get capacity text indicator
            cap_indicator = ""
            if ws.capacity_type == 'Premium':
                cap_indicator = " \u25C6"  # Filled diamond
            elif ws.capacity_type == 'PPU':
                cap_indicator = " \u25C7"  # Outlined diamond
            self.workspace_list.insert('', 'end', values=(star, f"{ws.name}{cap_indicator}"), tags=('ws',))

    def _refresh_workspace_list_by_results(self, search_results: List):
        """Refresh workspace list to show only workspaces with matching models."""
        self.workspace_list.delete(*self.workspace_list.get_children())

        # Get unique workspace IDs from search results
        matching_workspace_ids = set()
        for ds in search_results:
            if hasattr(ds, 'workspace_id') and ds.workspace_id:
                matching_workspace_ids.add(ds.workspace_id)

        # Filter to only workspaces with matches
        filtered = [ws for ws in self._all_workspaces if ws.id in matching_workspace_ids]
        self.workspaces = filtered

        for ws in filtered:
            # Star indicator: ★ if all models favorited, ☆ if some, empty if none
            star = self._get_workspace_star(ws.id)
            # Get capacity text indicator (trailing the name)
            cap_indicator = ""
            if ws.capacity_type == 'Premium':
                cap_indicator = " \u25C6"  # Filled diamond
            elif ws.capacity_type == 'PPU':
                cap_indicator = " \u25C7"  # Outlined diamond
            self.workspace_list.insert('', 'end', values=(star, f"{ws.name}{cap_indicator}"), tags=('ws',))

    def _on_workspace_select(self, event):
        """Handle workspace selection - filter search results or load all datasets."""
        selection = self.workspace_list.selection()
        if not selection:
            return

        item_id = selection[0]
        item_values = self.workspace_list.item(item_id, 'values')
        # values=(star, name) - name is at index 1
        item_name = item_values[1] if item_values and len(item_values) > 1 else ""

        # Check if "(All Workspaces)" was selected
        if "(All Workspaces)" in item_name:
            self._selected_workspace_id = None
            # Show all datasets from all workspaces
            if self._last_search_results:
                self._populate_datasets(self._last_search_results, None)
            else:
                self._show_all_datasets()
            return

        # Find workspace index based on position in treeview
        children = self.workspace_list.get_children()
        idx = children.index(item_id)

        # Adjust index to account for "(All Workspaces)" entry at top
        ws_idx = idx - 1 if idx > 0 and not self._is_workspace_list_filtered() else idx
        if ws_idx < 0 or ws_idx >= len(self.workspaces):
            return

        workspace = self.workspaces[ws_idx]
        self._selected_workspace_id = workspace.id

        # If we have search results, filter them by selected workspace
        if self._last_search_results:
            filtered = [ds for ds in self._last_search_results
                       if hasattr(ds, 'workspace_id') and ds.workspace_id == workspace.id]
            self._populate_datasets(filtered, None)
        else:
            # No search active - load all datasets for this workspace
            self._load_datasets(workspace.id)

    def _is_workspace_list_filtered(self) -> bool:
        """Check if workspace list is currently filtered (no All Workspaces option)."""
        children = self.workspace_list.get_children()
        if children:
            first_values = self.workspace_list.item(children[0], 'values')
            # values=(star, name) - name is at index 1
            first_name = first_values[1] if first_values and len(first_values) > 1 else ""
            return "(All Workspaces)" not in first_name
        return True

    def _on_workspace_favorite_toggle(self, event):
        """Toggle favorite status for double-clicked workspace.

        Uses after() to let the selection event complete first, fixing the issue
        where the first double-click does nothing.
        """
        item_id = self.workspace_list.identify_row(event.y)
        if not item_id:
            return

        # Skip if "(All Workspaces)" was clicked - values=(star, name)
        item_values = self.workspace_list.item(item_id, 'values')
        item_name = item_values[1] if item_values and len(item_values) > 1 else ""
        if "(All Workspaces)" in item_name:
            return

        # Delay processing to let TreeviewSelect event complete first
        self.after(10, lambda: self._process_workspace_favorite_toggle(item_id))

    def _process_workspace_favorite_toggle(self, item_id: str):
        """Process the workspace favorite toggle after selection completes."""
        try:
            # Find index in list
            children = self.workspace_list.get_children()
            if item_id not in children:
                return
            idx = children.index(item_id)

            # Adjust index for "(All Workspaces)" entry at top
            ws_idx = idx - 1 if not self._is_workspace_list_filtered() else idx
            if ws_idx < 0 or ws_idx >= len(self.workspaces):
                return

            workspace = self.workspaces[ws_idx]
            self._toggle_workspace_favorite(workspace)
        except Exception as e:
            self.logger.debug(f"Error toggling workspace favorite: {e}")

    def _on_workspace_right_click(self, event):
        """Show context menu for workspace right-click"""
        item_id = self.workspace_list.identify_row(event.y)
        if not item_id:
            return

        # Skip if "(All Workspaces)" was right-clicked - values=(star, name)
        item_values = self.workspace_list.item(item_id, 'values')
        item_name = item_values[1] if item_values and len(item_values) > 1 else ""
        if "(All Workspaces)" in item_name:
            return

        # Find index in list
        children = self.workspace_list.get_children()
        idx = children.index(item_id)

        # Adjust index for "(All Workspaces)" entry at top
        ws_idx = idx - 1 if not self._is_workspace_list_filtered() else idx
        if ws_idx < 0 or ws_idx >= len(self.workspaces):
            return

        # Select the item that was right-clicked
        self.workspace_list.selection_set(item_id)

        workspace = self.workspaces[ws_idx]

        # Create custom popup menu (tk.Menu ignores styling on Windows)
        colors = self._colors
        is_dark = self._is_dark

        # Theme-appropriate colors
        if is_dark:
            menu_bg = '#1a1a2e'
            menu_fg = '#e0e0e0'
            highlight_bg = '#004466'  # Blue hover (matches button_primary_hover)
            highlight_fg = '#ffffff'
            border_color = '#3a3a4a'  # Match tree_border
        else:
            menu_bg = '#ffffff'
            menu_fg = '#333333'
            highlight_bg = '#007A7A'  # Teal hover (matches button_primary_hover)
            highlight_fg = '#ffffff'
            border_color = '#d8d8e0'  # Match tree_border

        # Custom popup window for full styling control
        popup = tk.Toplevel(self)
        popup.overrideredirect(True)  # Remove window decorations
        popup.configure(bg=border_color)  # Border color via background

        # Inner frame for content (creates border effect via padding)
        inner_frame = tk.Frame(popup, bg=menu_bg)
        inner_frame.pack(padx=1, pady=1)  # 1px border

        # Menu item label - based on model favorites status
        # "Remove from favorites" if any models favorited, "Add to favorites" if none
        fav_status = self.browser.get_workspace_favorite_status(workspace.id)
        label_text = "Remove from Favorites" if fav_status != 'none' else "Add to Favorites"
        menu_label = tk.Label(
            inner_frame,
            text=label_text,
            bg=menu_bg,
            fg=menu_fg,
            font=('Segoe UI', 9),
            padx=12,
            pady=6,
            cursor='hand2',
            anchor='center'
        )
        menu_label.pack(fill='x')

        # Hover effects
        def on_enter(e):
            menu_label.configure(bg=highlight_bg, fg=highlight_fg)

        def on_leave(e):
            menu_label.configure(bg=menu_bg, fg=menu_fg)

        def on_click(e):
            popup.destroy()
            self._toggle_workspace_favorite(workspace)

        menu_label.bind('<Enter>', on_enter)
        menu_label.bind('<Leave>', on_leave)
        menu_label.bind('<Button-1>', on_click)

        # Position at click location
        popup.geometry(f'+{event.x_root}+{event.y_root}')

        # Close popup when clicking elsewhere or losing focus
        def close_popup(e=None):
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except tk.TclError:
                pass

        popup.bind('<FocusOut>', close_popup)
        popup.bind('<Escape>', close_popup)
        popup.focus_set()

    def _toggle_workspace_favorite(self, workspace):
        """Toggle favorite status for all models in a workspace.

        Favoriting a workspace = favorite ALL models in it.
        If any models are already favorited, unfavorite ALL.
        """
        try:
            # Ensure datasets are loaded for this workspace before favoriting
            datasets, error = self.browser.get_workspace_datasets(workspace.id)
            if error or not datasets:
                # No datasets available - nothing to favorite
                return

            # Check current favorite status of workspace's models
            current_status = self.browser.get_workspace_favorite_status(workspace.id)

            if current_status == 'none':
                # No models favorited - favorite all
                self.browser.favorite_all_models_in_workspace(workspace.id)
            else:
                # Some or all models favorited - unfavorite all
                self.browser.unfavorite_all_models_in_workspace(workspace.id)

            # Update workspace star indicator
            self._update_workspace_star_indicator(workspace)

            # Update model stars in dataset tree if this workspace is selected
            self._refresh_dataset_stars()

            # Force UI update immediately
            self.update_idletasks()
        except Exception as e:
            ThemedMessageBox.showerror(self, "Error", f"Failed to toggle favorites: {e}")

    def _update_workspace_star_indicator(self, workspace):
        """Update just the star indicator for a workspace without rebuilding the list.

        Two-column values approach: values=(star, name_with_capacity)
        """
        for item_id in self.workspace_list.get_children():
            item_values = self.workspace_list.item(item_id, 'values')
            if not item_values or len(item_values) < 2:
                continue
            item_star = item_values[0]
            item_name = item_values[1]
            # Skip "(All Workspaces)" row
            if "(All Workspaces)" in item_name:
                continue
            # Strip capacity indicator to get base name for matching
            base_name = item_name.rstrip(" \u25C6\u25C7").strip()
            if base_name == workspace.name:
                # Update just the star column (values[0]) with new star indicator
                new_star = self._get_workspace_star(workspace.id)
                if new_star != item_star:
                    self.workspace_list.item(item_id, values=(new_star, item_name))
                break

    def _refresh_dataset_stars(self):
        """Refresh star indicators for all displayed datasets."""
        if not hasattr(self, 'current_datasets') or not self.current_datasets:
            return

        for item_id in self.dataset_tree.get_children():
            try:
                idx = int(item_id)
                if idx >= len(self.current_datasets):
                    continue
                ds = self.current_datasets[idx]
                values = list(self.dataset_tree.item(item_id, 'values'))
                if len(values) >= 1:
                    # Update model name with star indicator
                    is_fav = self.browser.is_model_favorite(ds.id)
                    base_name = values[0].lstrip("\u2605 ").strip()  # Strip existing star
                    new_name = f"\u2605 {base_name}" if is_fav else base_name  # ★ always for models
                    if values[0] != new_name:
                        values[0] = new_name
                        self.dataset_tree.item(item_id, values=values)
            except (ValueError, IndexError):
                continue

    def _on_dataset_double_click(self, event):
        """Handle double-click on dataset tree - only select if clicked on a data row."""
        # Check if we clicked on a valid data row (not header or empty space)
        region = self.dataset_tree.identify_region(event.x, event.y)
        if region != "cell" and region != "tree":
            # Clicked on header, separator, or nothing - ignore
            return

        item_id = self.dataset_tree.identify_row(event.y)
        if not item_id:
            # No row at click position
            return

        # Verify it's a valid dataset index
        try:
            idx = int(item_id)
            if idx < 0 or idx >= len(self.current_datasets):
                return
        except (ValueError, AttributeError):
            return

        # Valid row - proceed with selection
        self._on_select()

    def _on_dataset_right_click(self, event):
        """Show context menu for dataset/model right-click"""
        # Identify which row was clicked
        item_id = self.dataset_tree.identify_row(event.y)
        if not item_id:
            return

        # Select the item that was right-clicked
        self.dataset_tree.selection_set(item_id)

        # Get the dataset
        try:
            idx = int(item_id)
            if idx < 0 or idx >= len(self.current_datasets):
                return
            dataset = self.current_datasets[idx]
        except (ValueError, AttributeError):
            return

        # Create custom popup menu (same style as workspace context menu)
        colors = self._colors
        is_dark = self._is_dark

        # Theme-appropriate colors
        if is_dark:
            menu_bg = '#1a1a2e'
            menu_fg = '#e0e0e0'
            highlight_bg = '#004466'  # Blue hover
            highlight_fg = '#ffffff'
            border_color = '#3a3a4a'
        else:
            menu_bg = '#ffffff'
            menu_fg = '#333333'
            highlight_bg = '#007A7A'  # Teal hover
            highlight_fg = '#ffffff'
            border_color = '#d8d8e0'

        # Custom popup window
        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.configure(bg=border_color)

        inner_frame = tk.Frame(popup, bg=menu_bg)
        inner_frame.pack(padx=1, pady=1)

        # Check if model is favorite
        is_favorite = self.browser.is_model_favorite(dataset.id)
        label_text = "Remove from Favorites" if is_favorite else "Add to Favorites"

        menu_label = tk.Label(
            inner_frame,
            text=label_text,
            bg=menu_bg,
            fg=menu_fg,
            font=('Segoe UI', 9),
            padx=12,
            pady=6,
            cursor='hand2',
            anchor='center'
        )
        menu_label.pack(fill='x')

        def on_enter(e):
            menu_label.configure(bg=highlight_bg, fg=highlight_fg)

        def on_leave(e):
            menu_label.configure(bg=menu_bg, fg=menu_fg)

        def on_click(e):
            popup.destroy()
            self._toggle_model_favorite(dataset)

        menu_label.bind('<Enter>', on_enter)
        menu_label.bind('<Leave>', on_leave)
        menu_label.bind('<Button-1>', on_click)

        popup.geometry(f'+{event.x_root}+{event.y_root}')

        def close_popup(e=None):
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except tk.TclError:
                pass

        popup.bind('<FocusOut>', close_popup)
        popup.bind('<Escape>', close_popup)
        popup.focus_set()

    def _toggle_model_favorite(self, dataset):
        """Toggle favorite status for a semantic model/dataset"""
        try:
            is_favorite = self.browser.toggle_model_favorite(dataset.id)
            # Refresh the dataset tree to show star indicator (and filter in favorites mode)
            self._refresh_dataset_display()

            # Handle workspace list updates
            if hasattr(dataset, 'workspace_id') and dataset.workspace_id:
                if self.view_mode.get() == "favorites":
                    # In favorites mode, refresh workspace list to remove workspaces
                    # that no longer have any favorited models
                    self._refresh_favorites_workspace_list()
                else:
                    # In other modes, just update the star indicator
                    for ws in self._all_workspaces:
                        if ws.id == dataset.workspace_id:
                            self._update_workspace_star_indicator(ws)
                            break

            # Force UI update immediately
            self.update_idletasks()
        except Exception as e:
            ThemedMessageBox.showerror(self, "Error", f"Failed to toggle favorite: {e}")

    def _refresh_dataset_display(self):
        """Refresh the dataset tree display (update stars).

        When in favorites view mode, filters out unfavorited models.
        """
        if not hasattr(self, 'current_datasets') or not self.current_datasets:
            return

        # Determine which datasets to display
        datasets_to_show = self.current_datasets
        if self.view_mode.get() == "favorites":
            # Filter to only show individually favorited models
            datasets_to_show = [ds for ds in self.current_datasets if self.browser.is_model_favorite(ds.id)]

        # Re-populate with filtered datasets
        self.dataset_tree.delete(*self.dataset_tree.get_children())
        for i, ds in enumerate(datasets_to_show):
            # Add star indicator for favorites (★ filled star for models)
            star = "\u2605 " if self.browser.is_model_favorite(ds.id) else ""
            # Get capacity indicator for workspace (text suffix like " ◆")
            cap_indicator = self._get_capacity_indicator_for_workspace(ds.workspace_id)
            workspace_display = (ds.workspace_name or "") + cap_indicator
            self.dataset_tree.insert(
                "", "end", iid=str(i),
                values=(f"{star}{ds.name}", workspace_display, ds.configured_by or "")
            )

    def _load_datasets(self, workspace_id: str):
        """Load datasets for a workspace"""
        self.dataset_tree.delete(*self.dataset_tree.get_children())

        def load_thread():
            try:
                datasets, error = self.browser.get_workspace_datasets(workspace_id)
                self._safe_after(lambda: self._populate_datasets(datasets, error))
            except Exception as e:
                # Catch any unexpected errors and show them to the user
                self._safe_after(lambda: self._populate_datasets([], f"Failed to load datasets: {e}"))

        threading.Thread(target=load_thread, daemon=True).start()

    def _populate_datasets(self, datasets: List[DatasetInfo], error: Optional[str]):
        """Populate semantic model list.

        When in "favorites" view mode, only shows individually favorited models.
        Preserves current selection if the selected item is still in the new list.
        """
        # Restore search button state
        self.search_btn.set_enabled(True)
        self.search_btn.update_text("Search")

        if error:
            ThemedMessageBox.showerror(self, "Error", error)
            return

        # If in favorites view mode, filter to only show favorited models
        if self.view_mode.get() == "favorites":
            datasets = [ds for ds in datasets if self.browser.is_model_favorite(ds.id)]

        # Save current selection before clearing (by dataset ID for reliable matching)
        selected_dataset_id = None
        if hasattr(self, '_selected_dataset') and self._selected_dataset:
            selected_dataset_id = self._selected_dataset.id

        self.current_datasets = datasets
        self.dataset_tree.delete(*self.dataset_tree.get_children())

        # Track which index the previously selected dataset is now at
        restore_selection_idx = None

        for i, ds in enumerate(datasets):
            # Check if this is the previously selected dataset
            if selected_dataset_id and ds.id == selected_dataset_id:
                restore_selection_idx = i

            # Add star indicator for favorites (★ filled star for models)
            star = "\u2605 " if self.browser.is_model_favorite(ds.id) else ""
            # Get capacity indicator for workspace (text suffix like " ◆")
            cap_indicator = self._get_capacity_indicator_for_workspace(ds.workspace_id)
            workspace_display = (ds.workspace_name or "") + cap_indicator

            # In simple_mode, apply 'pro_workspace' tag to Pro items (grayed out)
            tags = ()
            if self._simple_mode and self._is_workspace_pro(ds.workspace_id):
                tags = ('pro_workspace',)

            self.dataset_tree.insert(
                "", "end", iid=str(i),
                values=(f"{star}{ds.name}", workspace_display, ds.configured_by or ""),
                tags=tags
            )

        # Restore selection if the previously selected item is still in the list
        if restore_selection_idx is not None:
            self.dataset_tree.selection_set(str(restore_selection_idx))
            self.dataset_tree.see(str(restore_selection_idx))

    def _sort_by_column(self, column: str):
        """Sort the dataset list by the specified column."""
        if not hasattr(self, 'current_datasets') or not self.current_datasets:
            return

        # Toggle direction if same column, otherwise sort ascending
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False

        # Get the sort key based on column
        key_func = self._get_sort_key_func(column)
        if not key_func:
            return

        # Sort the datasets
        sorted_datasets = sorted(self.current_datasets, key=key_func, reverse=self._sort_reverse)

        # Update the treeview
        self.dataset_tree.delete(*self.dataset_tree.get_children())
        self.current_datasets = sorted_datasets

        for i, ds in enumerate(sorted_datasets):
            # Add star indicator for favorites (★ filled star for models)
            star = "\u2605 " if self.browser.is_model_favorite(ds.id) else ""
            # Get capacity indicator for workspace (text suffix like " ◆")
            cap_indicator = self._get_capacity_indicator_for_workspace(ds.workspace_id)
            workspace_display = (ds.workspace_name or "") + cap_indicator

            # In simple_mode, apply 'pro_workspace' tag to Pro items (grayed out)
            tags = ()
            if self._simple_mode and self._is_workspace_pro(ds.workspace_id):
                tags = ('pro_workspace',)

            self.dataset_tree.insert(
                "", "end", iid=str(i),
                values=(f"{star}{ds.name}", workspace_display, ds.configured_by or ""),
                tags=tags
            )

        # Update column headers with sort direction indicators
        self._update_sort_indicators()

    def _update_sort_indicators(self):
        """Update column headers with sort direction arrows."""
        # Base header text for each column
        base_headers = {
            "name": "Semantic Model",
            "workspace": "Workspace",
            "configured_by": "Configured By"
        }

        for col, base_text in base_headers.items():
            if col == self._sort_column:
                # Add arrow indicator for sorted column (extra space for visual separation)
                arrow = "  \u25B2" if not self._sort_reverse else "  \u25BC"  # Up or Down arrow
                self.dataset_tree.heading(col, text=base_text + arrow)
            else:
                # Reset to base text
                self.dataset_tree.heading(col, text=base_text)

    def _on_dataset_select(self, event):
        """Handle dataset selection"""
        selection = self.dataset_tree.selection()
        if not selection:
            return

        idx = int(selection[0])
        if idx >= len(self.current_datasets):
            return

        dataset = self.current_datasets[idx]
        self._update_selection_display(f"{dataset.workspace_name} / {dataset.name}")

        # In simple_mode, check if workspace is Pro and update Connect button state
        if self._simple_mode:
            is_pro = self._is_workspace_pro(dataset.workspace_id)
            self._update_connect_button_state(is_pro)

        # Load perspectives for this dataset
        self._on_dataset_select_perspectives(dataset)

    def _on_search_focus_in(self, event):
        """Clear placeholder text when search entry gains focus."""
        if self.search_var.get() == self._search_placeholder:
            self.search_entry.delete(0, tk.END)
            self.search_entry.configure(fg=self._colors['text_primary'])

    def _on_search_focus_out(self, event):
        """Restore placeholder text when search entry loses focus and is empty."""
        if not self.search_var.get().strip():
            self.search_entry.delete(0, tk.END)
            self.search_entry.insert(0, self._search_placeholder)
            self.search_entry.configure(fg=self._colors['text_muted'])

    def _on_search_var_changed(self, *args):
        """Handle search variable changes for auto-clear functionality."""
        # Cancel any pending timer
        if self._search_clear_timer is not None:
            try:
                self.after_cancel(self._search_clear_timer)
            except Exception:
                pass
            self._search_clear_timer = None

        query = self.search_var.get().strip()
        # Ignore placeholder text
        if query == self._search_placeholder:
            return
        if not query:
            # Search box was cleared - start 1-second timer to reset filters
            self._search_clear_timer = self.after(1000, self._auto_clear_search)

    def _auto_clear_search(self):
        """Auto-clear search and reset to unfiltered state after delay."""
        self._search_clear_timer = None
        self._selected_workspace_id = None
        self._last_search_results = []

        # Reset workspace list to show all
        if hasattr(self, '_all_workspaces'):
            self._refresh_workspace_list("")

        # Show all datasets when search is cleared (if data is cached)
        if self.browser.is_fully_cached():
            self._show_all_datasets()
        else:
            self.dataset_tree.delete(*self.dataset_tree.get_children())

    def _search_datasets(self):
        """Search for semantic models and filter workspaces by matching models."""
        query = self.search_var.get().strip()

        # Treat placeholder text as empty query
        if query == self._search_placeholder:
            query = ""

        # Reset workspace selection when searching
        self._selected_workspace_id = None

        if not query:
            # Clear and reset - show all datasets if cached
            self._last_search_results = []
            if hasattr(self, '_all_workspaces'):
                self._refresh_workspace_list("")
            if self.browser.is_fully_cached():
                self._show_all_datasets()
            else:
                self.dataset_tree.delete(*self.dataset_tree.get_children())
            return

        # Use cached search if available (fast path)
        if self.browser.is_fully_cached():
            results = self.browser.search_datasets_cached(query)
            self._handle_search_results(results, None)
            return

        # Fall back to full search (with API calls)
        self.search_btn.set_enabled(False)
        self.search_btn.update_text("Searching...")
        self.dataset_tree.delete(*self.dataset_tree.get_children())

        def search_thread():
            try:
                results, error = self.browser.search_datasets(query)
                self._safe_after(lambda: self._handle_search_results(results, error))
            except Exception as e:
                self._safe_after(lambda: self._handle_search_results([], f"Search failed: {e}"))

        threading.Thread(target=search_thread, daemon=True).start()

    def _handle_search_results(self, results: List, error: Optional[str]):
        """Handle search results - filter workspaces and show all matching models."""
        # Restore search button state
        self.search_btn.set_enabled(True)
        self.search_btn.update_text("Search")

        if error:
            ThemedMessageBox.showerror(self, "Error", error)
            return

        # Cache the search results for workspace filtering
        self._last_search_results = results

        # Filter workspace list to only show workspaces with matching models
        if hasattr(self, '_all_workspaces'):
            self._refresh_workspace_list_by_results(results)

        # Show ALL matching models from all workspaces
        self._populate_datasets(results, None)

    def _on_select(self):
        """Confirm selection and close"""
        mode = self.view_mode.get()

        # Get selected cloud connection type
        conn_type_value = self.cloud_conn_type.get()
        try:
            cloud_conn_type = CloudConnectionType(conn_type_value)
        except ValueError:
            cloud_conn_type = CloudConnectionType.PBI_SEMANTIC_MODEL

        # Get perspective selection (from dropdown or manual entry)
        perspective_name = None
        if hasattr(self, 'perspective_var'):
            perspective_name = self.perspective_var.get().strip() or None

        # Validate perspective name if one was entered and we can validate
        if perspective_name and getattr(self, '_can_validate_perspectives', False):
            current_perspectives = getattr(self, '_current_perspectives', [])
            if current_perspectives:
                # Case-insensitive comparison
                known_names = [p.lower() for p in current_perspectives]
                if perspective_name.lower() not in known_names:
                    result = ThemedMessageBox.askyesno(
                        self,
                        "Perspective Not Found",
                        f"The perspective '{perspective_name}' was not found in auto-discovery.\n\n"
                        "This could mean:\n"
                        "- The name is misspelled\n"
                        "- The perspective doesn't exist in this model\n\n"
                        "Continue anyway?"
                    )
                    if not result:
                        return  # Let user correct the name

        if mode == "manual":
            # Manual entry
            workspace = self.manual_workspace.get().strip()
            dataset = self.manual_dataset.get().strip()

            if not workspace or not dataset:
                ThemedMessageBox.showwarning(self, "Missing Info", "Please enter workspace and dataset name.")
                return

            self.result = self.browser.create_manual_target(
                workspace, dataset,
                cloud_connection_type=cloud_conn_type,
                perspective_name=perspective_name
            )
            self.destroy()
            return

        # Tree selection
        selection = self.dataset_tree.selection()
        if not selection:
            ThemedMessageBox.showinfo(self, "Select Dataset", "Please select a dataset.")
            return

        idx = int(selection[0])
        if idx >= len(self.current_datasets):
            return

        dataset = self.current_datasets[idx]
        self.result = dataset.to_swap_target(cloud_connection_type=cloud_conn_type)

        # Add perspective to the result if selected
        if perspective_name:
            self.result.perspective_name = perspective_name
            # Update display name to include perspective
            self.result.display_name = f"{dataset.name} [{perspective_name}] ({dataset.workspace_name})"

        # Add to recent
        self.browser.add_recent(dataset.workspace_id, dataset.id)

        self.destroy()

    def _on_cancel(self):
        """Cancel and close"""
        self.result = None
        self.destroy()

    # ===== Simple Mode Pro Workspace Handling =====

    def _setup_connect_tooltip(self):
        """Setup tooltip for the Connect button in simple_mode."""
        if not self._simple_mode or not hasattr(self, 'select_btn'):
            return

        # Create tooltip label (hidden by default)
        self._connect_tooltip_text = tk.StringVar(value="")
        self._connect_tooltip_label = None

        def show_tooltip(event):
            if not self._selected_is_pro:
                return
            # Create tooltip near the button
            x = event.x_root + 10
            y = event.y_root + 10
            if self._connect_tooltip_label is None:
                self._connect_tooltip_label = tk.Toplevel(self)
                self._connect_tooltip_label.wm_overrideredirect(True)
                self._connect_tooltip_label.wm_attributes("-topmost", True)
                colors = self._theme_manager.colors
                is_dark = self._theme_manager.is_dark
                bg = colors.get('surface', '#2d2d3d' if is_dark else '#f5f5f5')
                fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
                label = tk.Label(
                    self._connect_tooltip_label,
                    text="Requires Premium, PPU, or Fabric capacity.\nPro workspaces do not support XMLA connections.",
                    bg=bg, fg=fg,
                    font=("Segoe UI", 9),
                    padx=8, pady=4,
                    justify=tk.LEFT,
                    relief=tk.SOLID,
                    borderwidth=1
                )
                label.pack()
            self._connect_tooltip_label.wm_geometry(f"+{x}+{y}")
            self._connect_tooltip_label.deiconify()

        def hide_tooltip(event):
            if self._connect_tooltip_label:
                self._connect_tooltip_label.withdraw()

        self.select_btn.bind("<Enter>", show_tooltip)
        self.select_btn.bind("<Leave>", hide_tooltip)

    def _is_workspace_pro(self, workspace_id: str) -> bool:
        """Check if a workspace is Pro (no XMLA access).

        Returns True if workspace has no capacity (capacity_type is None).
        """
        for ws in self.browser._workspaces:
            if ws.id == workspace_id:
                return ws.capacity_type is None
        return False  # Unknown, assume not Pro

    def _update_connect_button_state(self, is_pro: bool):
        """Update Connect button enabled/disabled state based on Pro workspace selection.

        Only applies in simple_mode.
        """
        if not self._simple_mode or not hasattr(self, 'select_btn'):
            return

        self._selected_is_pro = is_pro

        if is_pro:
            # Disable the Connect button
            self.select_btn.set_enabled(False)
            self.select_btn.config(cursor='arrow')
        else:
            # Enable the Connect button
            self.select_btn.set_enabled(True)
            self.select_btn.config(cursor='hand2')
