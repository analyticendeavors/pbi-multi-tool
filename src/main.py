"""
Enhanced Power BI Report Tools - Main Application with Sidebar Navigation
Built by Reid Havens of Analytic Endeavors

Modern sidebar-based UI with lazy loading and dark/light mode support.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path
from typing import Dict, Optional

# Pillow for icon loading (optional - falls back to emoji)
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# CairoSVG for SVG icon rendering (optional - falls back to PNG or emoji)
try:
    import cairosvg
    import io
    CAIROSVG_AVAILABLE = True
except (ImportError, OSError):
    CAIROSVG_AVAILABLE = False

# Add parent directory to Python path for organized imports
sys.path.append(str(Path(__file__).parent))
sys.path.append(str(Path(__file__).parent.parent))

from core.constants import AppConstants
from core.enhanced_base_tool import EnhancedBaseExternalTool, ToolConfiguration
from core.tool_manager import get_tool_manager
from core.theme_manager import get_theme_manager
from core.ui_base import RoundedNavButton, RoundedButton


class EnhancedPowerBIReportToolsApp(EnhancedBaseExternalTool):
    """
    Enhanced Power BI Report Tools with sidebar navigation.
    Features:
    - Modern sidebar navigation (replaces horizontal tabs)
    - Lazy loading for faster startup
    - Dark/light mode support
    - Wider layout optimized for laptops
    """

    def __init__(self):
        # Initialize with tool configuration
        config = ToolConfiguration(
            name="Enhanced Power BI Report Tools",
            version="2.0.0",
            description="Professional suite for Power BI report management",
            author="Reid Havens",
            website="https://www.analyticendeavors.com",
            icon_path="assets/favicon.ico"
        )

        super().__init__(config)

        # Get managers
        self.tool_manager = get_tool_manager()
        self.theme_manager = get_theme_manager()

        # UI Components
        self.sidebar_frame: Optional[tk.Frame] = None
        self.content_frame: Optional[ttk.Frame] = None
        self.tool_content_frame: Optional[ttk.Frame] = None
        self.current_tool_label: Optional[ttk.Label] = None
        self.theme_button: Optional[tk.Button] = None

        # Navigation state
        self.nav_buttons: Dict[str, tk.Button] = {}
        self.current_tool_id: Optional[str] = None
        self.sidebar_collapsed: bool = False

        # Sidebar widgets for collapse toggle
        self.brand_widgets: list = []  # Widgets to hide when collapsed
        self.collapse_btn: Optional[tk.Label] = None

        # Progress indicators for nav buttons
        self.nav_indicators: Dict[str, tk.Label] = {}

        # Tool ordering - use constant from AppConstants for central configuration
        self.tool_order = AppConstants.TOOL_ORDER

        # Tool icon mapping (tool_id -> icon filename without extension)
        self.tool_icons_map = {
            "report_merger": "report merger",
            "pbip_layout_optimizer": "layout optimizer",
            "report_cleanup": "report cleanup",
            "accessibility_checker": "witch-hat",
            "column_width": "table column widths",
            "advanced_copy": "advanced copy",
            "field_parameters": "Field Parameter",
            "connection_hotswap": "hotswap",
            "sensitivity_scanner": "sensivity scanner"  # Note: typo in filename
        }

        # Icon storage (prevents garbage collection)
        self._tool_icons: Dict[str, ImageTk.PhotoImage] = {}

        # Initialize tools
        self._initialize_tools()

    def _initialize_tools(self):
        """Initialize and register all tools"""
        registered_count = self.tool_manager.discover_and_register_tools("tools")

        if registered_count == 0:
            self.logger.log_security_event("No tools were discovered", "WARNING")

    def _load_tool_icons(self):
        """Load tool icons from SVG files (converted to 20x20 PNG)"""
        if not PIL_AVAILABLE:
            return

        icons_dir = Path(__file__).parent / "assets" / "Tool Icons"

        if not icons_dir.exists():
            return

        for tool_id, icon_name in self.tool_icons_map.items():
            svg_path = icons_dir / f"{icon_name}.svg"
            png_path = icons_dir / f"{icon_name}.png"

            try:
                img = None

                # Try SVG first (if cairosvg available)
                if CAIROSVG_AVAILABLE and svg_path.exists():
                    # Render at 4x size for quality, then downscale
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=80,
                        output_height=80
                    )
                    img = Image.open(io.BytesIO(png_data))

                # Fall back to PNG if SVG not available
                elif png_path.exists():
                    img = Image.open(png_path)

                if img is None:
                    continue

                # Ensure RGBA for transparency
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')

                # Resize to 20x20 with high-quality resampling
                img = img.resize((20, 20), Image.Resampling.LANCZOS)

                # Convert to PhotoImage
                photo = ImageTk.PhotoImage(img)
                self._tool_icons[tool_id] = photo

            except Exception:
                pass  # Silently fail - will use emoji fallback

        # Load theme toggle icons
        theme_icons = {"light_mode": "light-mode", "dark_mode": "night-mode"}
        for icon_id, icon_name in theme_icons.items():
            svg_path = icons_dir / f"{icon_name}.svg"
            png_path = icons_dir / f"{icon_name}.png"

            try:
                img = None

                if CAIROSVG_AVAILABLE and svg_path.exists():
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=80,
                        output_height=80
                    )
                    img = Image.open(io.BytesIO(png_data))
                elif png_path.exists():
                    img = Image.open(png_path)

                if img is None:
                    continue

                if img.mode != 'RGBA':
                    img = img.convert('RGBA')

                # Theme icons slightly smaller (16x16) than tool icons (20x20)
                img = img.resize((16, 16), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self._tool_icons[icon_id] = photo

            except Exception:
                pass  # Silently fail - will use emoji fallback

    def create_ui(self) -> tk.Tk:
        """Create the main UI with sidebar layout"""
        # Get the ttkbootstrap theme based on saved preference
        ttkb_theme = self.theme_manager.TTKB_THEMES.get(
            self.theme_manager.current_theme, 'cyborg'
        )
        root = self.create_secure_ui_base(themename=ttkb_theme)
        root.withdraw()  # Hide until fully positioned to prevent flash
        root.title(f"AE Power BI Tools v{self.config.version}")
        root.geometry(AppConstants.WINDOW_SIZE)
        root.minsize(*AppConstants.MIN_WINDOW_SIZE)

        # Load tool icons now that root window exists
        if PIL_AVAILABLE:
            self._load_tool_icons()

        # Initialize theme manager with root
        self.theme_manager.initialize(root)

        # Register for theme changes
        self.theme_manager.register_theme_callback(self._on_theme_changed)

        # Create main layout
        self._setup_main_layout(root)

        # Center window
        self._center_window(root)

        # Store root reference before setting title bar (needed for title bar method)
        self.root = root

        # Force window to fully render before setting title bar color
        root.update()

        # Set initial title bar color to match theme
        self._set_title_bar_color(self.theme_manager.is_dark)

        # Select first tool by default
        self._select_first_tool()

        # Force initial theme application to ensure all widgets have correct styling
        # This fixes the issue where UI looks wrong until theme is toggled
        self._on_theme_changed(self.theme_manager.current_theme)

        root.deiconify()  # Show window now that it's positioned and styled
        return root

    def _center_window(self, window):
        """Center the window horizontally, position near top vertically"""
        window.update_idletasks()

        # Parse window size from geometry string (e.g., "1450x1105")
        width, height = AppConstants.WINDOW_SIZE.split('x')
        window_width = int(width)
        screen_width = window.winfo_screenwidth()

        x = (screen_width - window_width) // 2
        y = 50

        x = max(0, x)
        y = max(0, y)

        # Only set position, don't override the size already set
        window.geometry(f"+{x}+{y}")

    def _setup_main_layout(self, root):
        """Setup the sidebar + content layout"""
        # Main container using tk.Frame for better control
        main_container = tk.Frame(root)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Use pack instead of grid for simpler fixed width sidebar
        # Left: Sidebar (fixed width)
        self._setup_sidebar(main_container)

        # Right: Content Area (fills remaining space)
        self._setup_content_area(main_container)

    def _setup_sidebar(self, parent):
        """Setup the left navigation sidebar"""
        colors = self.theme_manager.colors
        width = AppConstants.SIDEBAR_WIDTH

        # Clear brand widgets list
        self.brand_widgets = []

        # Sidebar container (fixed width using pack)
        self.sidebar_frame = tk.Frame(parent, bg=colors['sidebar_bg'], width=width)
        self.sidebar_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar_frame.pack_propagate(False)  # Prevent children from affecting size

        # ===== TOP BAR: Collapse button on right =====
        top_bar = tk.Frame(self.sidebar_frame, bg=colors['sidebar_bg'], height=30)
        top_bar.pack(fill=tk.X, pady=(8, 0))
        top_bar.pack_propagate(False)  # Prevent resizing during theme changes

        # Collapse/Expand button - Unicode chevron
        self.collapse_btn = tk.Label(
            top_bar,
            text="❮",  # Heavy left-pointing angle bracket
            font=('Segoe UI', 16),
            fg=colors['button_primary'],
            bg=colors['sidebar_bg'],
            cursor='hand2'
        )
        self.collapse_btn.pack(side=tk.RIGHT, padx=(0, 20), pady=2)
        self.collapse_btn.bind('<Button-1>', lambda e: self._toggle_sidebar())
        self.collapse_btn.bind('<Enter>', lambda e: self.collapse_btn.config(fg=colors['button_primary_hover']))
        self.collapse_btn.bind('<Leave>', lambda e: self.collapse_btn.config(fg=colors['button_primary']))

        # Store reference to top_bar for theme updates
        self.top_bar = top_bar

        # ===== TOP: Logo/Branding (centered) =====
        brand_frame = tk.Frame(self.sidebar_frame, bg=colors['sidebar_bg'])
        brand_frame.pack(fill=tk.X, pady=(10, 20))

        # Logo text - centered (always visible)
        # Use button_primary: blue in dark mode (#00587C), teal in light mode (#009999)
        self.logo_label = tk.Label(brand_frame, text="Æ",
                 font=('Segoe UI', 32, 'bold'),
                 fg=colors['button_primary'],
                 bg=colors['sidebar_bg'])
        self.logo_label.pack(anchor=tk.CENTER)

        # Brand text labels (hidden when collapsed)
        multi_tool_label = tk.Label(brand_frame, text="MULTI-TOOL",
                 font=('Segoe UI', 12, 'bold'),
                 fg=colors['nav_text'],
                 bg=colors['sidebar_bg'])
        multi_tool_label.pack(anchor=tk.CENTER)
        self.brand_widgets.append(multi_tool_label)

        suite_label = tk.Label(brand_frame, text="Power BI Suite",
                 font=('Segoe UI', 9),
                 fg=colors['text_muted'],
                 bg=colors['sidebar_bg'])
        suite_label.pack(anchor=tk.CENTER, pady=(2, 0))
        self.brand_widgets.append(suite_label)

        # Separator (hidden when collapsed)
        sep1 = tk.Frame(self.sidebar_frame, bg=colors['border'], height=1)
        sep1.pack(fill=tk.X, padx=15, pady=(10, 15))
        self.brand_widgets.append(sep1)

        # ===== MIDDLE: Navigation Items =====
        self.nav_frame = tk.Frame(self.sidebar_frame, bg=colors['sidebar_bg'])
        self.nav_frame.pack(fill=tk.BOTH, expand=True, padx=(8, 0))  # Left padding only - right edge touches border

        self._create_nav_buttons(self.nav_frame)

        # ===== BOTTOM: Links & Theme Toggle =====
        self.bottom_frame = tk.Frame(self.sidebar_frame, bg=colors['sidebar_bg'])
        self.bottom_frame.pack(fill=tk.X, pady=(10, 20), padx=20)

        # Separator above links (top divider)
        self.top_divider = tk.Frame(self.bottom_frame, bg=colors['divider_faint'], height=1)
        self.top_divider.pack(fill=tk.X, pady=(0, 12))

        # Footer links first (centered, just text with hover color change)
        link_frame = tk.Frame(self.bottom_frame, bg=colors['sidebar_bg'])
        link_frame.pack(pady=(0, 12))
        self.link_frame = link_frame  # Save reference - handled separately in _toggle_sidebar()

        self.footer_links = []
        for text, command in [("Git", self.open_github_repo),
                              ("Donate", self.open_donate_page),
                              ("Web", self.open_company_website)]:
            link = tk.Label(link_frame, text=text,
                           font=('Segoe UI', 9),
                           fg=colors['text_muted'],
                           bg=colors['sidebar_bg'],
                           cursor='hand2')
            link.pack(side=tk.LEFT, padx=8)
            link.bind('<Button-1>', lambda e, cmd=command: cmd())
            link.bind('<Enter>', lambda e, lbl=link: lbl.config(fg=colors['text_primary']))
            link.bind('<Leave>', lambda e, lbl=link: lbl.config(fg=colors['text_muted']))
            self.footer_links.append(link)

        # Faint divider line between links and theme toggle
        self.links_theme_divider = tk.Frame(self.bottom_frame, bg=colors['divider_faint'], height=1)
        self.links_theme_divider.pack(fill=tk.X, pady=(0, 12))
        # Note: Handled separately in _toggle_sidebar() since it needs fill=tk.X

        # Theme toggle at the bottom - separate icon and text labels for consistent spacing
        self.theme_toggle_frame = tk.Frame(self.bottom_frame, bg=colors['sidebar_bg'], cursor='hand2')
        self.theme_toggle_frame.pack(pady=(0, 0))

        # Get theme icon image or fallback to text symbols
        theme_icon_id = "light_mode" if self.theme_manager.is_dark else "dark_mode"
        theme_icon_image = self._tool_icons.get(theme_icon_id)
        theme_icon_text = "◑" if self.theme_manager.is_dark else "◐"  # Fallback
        theme_text = "Light Mode" if self.theme_manager.is_dark else "Dark Mode"

        self.theme_icon_label = tk.Label(
            self.theme_toggle_frame,
            image=theme_icon_image if theme_icon_image else None,
            text="" if theme_icon_image else theme_icon_text,
            font=('Segoe UI', 14),
            fg=colors['nav_text'],
            bg=colors['sidebar_bg'],
            cursor='hand2'
        )
        self.theme_icon_label.pack(side=tk.LEFT, padx=(0, 6), anchor='center')

        self.theme_text_label = tk.Label(
            self.theme_toggle_frame,
            text=theme_text,
            font=('Segoe UI', 10),
            fg=colors['nav_text'],
            bg=colors['sidebar_bg'],
            cursor='hand2'
        )
        self.theme_text_label.pack(side=tk.LEFT, anchor='center')

        # Bind click and hover to all theme toggle widgets
        for widget in [self.theme_toggle_frame, self.theme_icon_label, self.theme_text_label]:
            widget.bind('<Button-1>', lambda e: self._toggle_theme())
            widget.bind('<Enter>', lambda e: self._on_theme_toggle_hover(True))
            widget.bind('<Leave>', lambda e: self._on_theme_toggle_hover(False))

        # Store frame reference for collapsed mode updates
        self.theme_button = self.theme_toggle_frame

    def _create_nav_buttons(self, parent):
        """Create navigation buttons for each tool with optional icons"""
        colors = self.theme_manager.colors

        # Calculate button width based on sidebar width (minus left margin)
        btn_width = AppConstants.SIDEBAR_WIDTH - 5  # 5px left margin

        tools = {t.tool_id: t for t in self.tool_manager.get_enabled_tools()}

        for tool_id in self.tool_order:
            if tool_id not in tools:
                continue

            tool = tools[tool_id]

            # Get icon if available
            icon = self._tool_icons.get(tool_id)

            # Determine button text
            if icon:
                btn_text = tool.name
            else:
                btn_text = tool.get_tab_title()

            # Create RoundedNavButton with mode-aware corner rounding
            btn = RoundedNavButton(
                parent,
                text=btn_text,
                command=lambda tid=tool_id: self._on_nav_click(tid),
                bg=colors['sidebar_bg'],
                fg=colors['nav_text'],
                hover_bg=colors['sidebar_hover'],
                pressed_bg=colors['sidebar_pressed'],
                active_bg=colors['sidebar_active'],
                active_hover_bg=colors['sidebar_active_hover'],
                active_pressed_bg=colors['sidebar_active_pressed'],
                icon=icon,
                mode='expanded',
                width=btn_width,
                height=40,
                radius=6,
                font=('Segoe UI', 10)
            )
            # Pack with left margin only - right side touches edge
            btn.pack(fill=tk.X, pady=2, padx=(5, 0))

            # Store reference
            self.nav_buttons[tool_id] = btn

        # Add any tools not in the predefined order
        for tool in self.tool_manager.get_enabled_tools():
            if tool.tool_id not in self.nav_buttons:
                # Get icon if available
                icon = self._tool_icons.get(tool.tool_id)

                if icon:
                    btn_text = tool.name
                else:
                    btn_text = tool.get_tab_title()

                btn = RoundedNavButton(
                    parent,
                    text=btn_text,
                    command=lambda tid=tool.tool_id: self._on_nav_click(tid),
                    bg=colors['sidebar_bg'],
                    fg=colors['nav_text'],
                    hover_bg=colors['sidebar_hover'],
                    pressed_bg=colors['sidebar_pressed'],
                    active_bg=colors['sidebar_active'],
                    active_hover_bg=colors['sidebar_active_hover'],
                    active_pressed_bg=colors['sidebar_active_pressed'],
                    icon=icon,
                    mode='expanded',
                    width=btn_width,
                    height=40,
                    radius=6,
                    font=('Segoe UI', 10)
                )
                btn.pack(fill=tk.X, pady=2, padx=(5, 0))
                self.nav_buttons[tool.tool_id] = btn

        # Force initial geometry update - _on_map event handles initial redraw
        parent.update_idletasks()

    def _on_nav_hover(self, button: tk.Button, tool_id: str, entering: bool):
        """Handle navigation button hover"""
        colors = self.theme_manager.colors

        if tool_id == self.current_tool_id:
            # Active button - use active hover color
            if entering:
                button.config(bg=colors['sidebar_active_hover'])
            else:
                button.config(bg=colors['sidebar_active'])
        else:
            # Inactive button
            if entering:
                button.config(bg=colors['sidebar_hover'])
            else:
                button.config(bg=colors['sidebar_bg'])

    def _on_nav_press(self, button: tk.Button, tool_id: str):
        """Handle navigation button press (mouse down)"""
        colors = self.theme_manager.colors

        if tool_id == self.current_tool_id:
            # Active button - use active pressed color
            button.config(bg=colors['sidebar_active_pressed'])
        else:
            button.config(bg=colors['sidebar_pressed'])

    def _on_nav_release(self, button: tk.Button, tool_id: str):
        """Handle navigation button release (mouse up)"""
        colors = self.theme_manager.colors

        if tool_id == self.current_tool_id:
            # Active button - return to active hover (mouse still over)
            button.config(bg=colors['sidebar_active_hover'])
        else:
            # Return to hover state since mouse is still over button
            button.config(bg=colors['sidebar_hover'])

    def show_tool_progress(self, tool_id: str):
        """Show progress indicator on a tool's nav button"""
        btn = self.nav_buttons.get(tool_id)
        if not btn:
            return

        colors = self.theme_manager.colors

        # Create indicator if not exists
        if tool_id not in self.nav_indicators:
            indicator = tk.Label(
                btn.master,
                text="●",
                font=('Segoe UI', 8),
                fg=colors['primary'],
                bg=colors['sidebar_bg']
            )
            self.nav_indicators[tool_id] = indicator

        # Position indicator at right edge of button
        indicator = self.nav_indicators[tool_id]
        indicator.place(in_=btn, relx=0.95, rely=0.5, anchor='e')
        indicator.lift()

    def hide_tool_progress(self, tool_id: str):
        """Hide progress indicator on a tool's nav button"""
        indicator = self.nav_indicators.get(tool_id)
        if indicator:
            indicator.place_forget()

    def _setup_content_area(self, parent):
        """Setup the right content area"""
        colors = self.theme_manager.colors

        # Content container (fills remaining space after sidebar)
        self.content_frame = ttk.Frame(parent)
        self.content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Tool content area with padding (no breadcrumb - moved to title bar)
        self.tool_content_frame = ttk.Frame(self.content_frame)
        self.tool_content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(15, 20))

        # No header label needed - sidebar provides navigation context
        self.current_tool_label = None

        # Create tool placeholders (lazy loading)
        self.tool_manager.create_tool_placeholders(self.tool_content_frame, self)

    def _select_first_tool(self):
        """Select the first available tool OR connection_hotswap if launched from external tool."""
        # Check if launched from Power BI external tool
        from core.pbi_connector import get_connector
        server, database = get_connector().parse_external_tool_args(sys.argv)
        if server:
            # External tool launch (composite or thin report) - select Connection Hot-Swap tab
            if "connection_hotswap" in self.nav_buttons:
                self._on_nav_click("connection_hotswap")
                return

        # Default: select first tool in order
        for tool_id in self.tool_order:
            if tool_id in self.nav_buttons:
                self._on_nav_click(tool_id)
                return

        # Fallback: select first available
        if self.nav_buttons:
            first_tool_id = list(self.nav_buttons.keys())[0]
            self._on_nav_click(first_tool_id)

    def _on_nav_click(self, tool_id: str):
        """Handle navigation button click"""
        if self.current_tool_id == tool_id:
            return  # Already selected

        # Update nav button styles
        self._update_nav_selection(tool_id)

        # Hide current tool content and notify deactivation
        if self.current_tool_id:
            # Notify previous tab of deactivation
            prev_tab = self.tool_manager.get_tool_tab(self.current_tool_id)
            if prev_tab and hasattr(prev_tab, 'on_tab_deactivated'):
                try:
                    prev_tab.on_tab_deactivated()
                except Exception:
                    pass  # Don't let tab errors prevent navigation

            current_placeholder = self.tool_manager.get_tool_placeholder(self.current_tool_id)
            if current_placeholder:
                current_placeholder.pack_forget()

        # Show/load new tool
        placeholder = self.tool_manager.get_tool_placeholder(tool_id)
        if placeholder:
            # Lazy load if not already loaded
            if not self.tool_manager.is_tool_loaded(tool_id):
                self.tool_manager.load_tool_ui(tool_id)

            # Show the tool
            placeholder.pack(fill=tk.BOTH, expand=True)

        # Update current tool tracking
        self.current_tool_id = tool_id

        # Notify new tab of activation
        new_tab = self.tool_manager.get_tool_tab(tool_id)
        if new_tab and hasattr(new_tab, 'on_tab_activated'):
            try:
                new_tab.on_tab_activated()
            except Exception:
                pass  # Don't let tab errors break the app

        # Update window title with tool name
        tool = self.tool_manager.get_tool(tool_id)
        if tool:
            self.root.title(f"AE Power BI Tools v{self.config.version}  |  {tool.name}")

    def _update_nav_selection(self, active_tool_id: str):
        """Update navigation button visual states"""
        for tool_id, btn in self.nav_buttons.items():
            # RoundedNavButton handles all color/style logic internally
            btn.set_active(tool_id == active_tool_id)

    def _toggle_theme(self):
        """Toggle between dark and light themes"""
        # Execute theme change - callbacks will update all widgets
        self.theme_manager.toggle_theme()
        # Update Windows title bar to match theme
        self._set_title_bar_color(self.theme_manager.is_dark)
        # Force all pending updates to render at once
        self.root.update_idletasks()

    def _on_theme_toggle_hover(self, entering: bool):
        """Handle hover on theme toggle"""
        colors = self.theme_manager.colors
        fg = colors['text_primary'] if entering else colors['nav_text']
        if hasattr(self, 'theme_icon_label'):
            self.theme_icon_label.config(fg=fg)
        if hasattr(self, 'theme_text_label'):
            self.theme_text_label.config(fg=fg)

    def _set_title_bar_color(self, dark_mode: bool = True):
        """Set Windows title bar to dark or light mode for main window"""
        self._set_window_title_bar_color(self.root, dark_mode)

    def _set_window_title_bar_color(self, window, dark_mode: bool = True):
        """Set Windows title bar to dark or light mode for any window"""
        import sys
        if sys.platform != 'win32' or not window:
            return

        try:
            import ctypes
            # Get window handle
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())

            # Windows 10/11 dark mode attribute
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20

            # Set dark mode
            value = ctypes.c_int(1 if dark_mode else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value),
                ctypes.sizeof(value)
            )
        except Exception:
            pass  # Silently fail on older Windows versions

    def _toggle_sidebar(self):
        """Toggle sidebar between expanded and collapsed states"""
        self.sidebar_collapsed = not self.sidebar_collapsed
        colors = self.theme_manager.colors

        if self.sidebar_collapsed:
            # Collapse to icons only
            new_width = AppConstants.SIDEBAR_COLLAPSED_WIDTH
            self.sidebar_frame.config(width=new_width)

            # Update collapse button direction for narrower sidebar - center it
            if self.collapse_btn:
                self.collapse_btn.pack_forget()
                self.collapse_btn.pack(anchor=tk.CENTER, pady=(5, 2))
                self.collapse_btn.config(text="❯")  # Right-pointing (expand)

            # Hide brand widgets
            for widget in self.brand_widgets:
                widget.pack_forget()

            # Hide link_frame (Help/About/Web) and divider
            if hasattr(self, 'link_frame') and self.link_frame:
                self.link_frame.pack_forget()
            if hasattr(self, 'links_theme_divider') and self.links_theme_divider:
                self.links_theme_divider.pack_forget()

            # Hide nav_frame during transition to prevent cascade effect
            self.nav_frame.pack_forget()

            # Update nav buttons to icon-only mode (touches right edge, icon centered)
            for tool_id, btn in self.nav_buttons.items():
                btn.pack_forget()
                # Use expanded mode for shape (left-rounded, right flush with edge)
                btn.set_mode('expanded')
                # But draw only icon centered (no text)
                btn.set_icon_only(True)
                # Resize to square-ish for icons
                btn.set_size(52, 44)
                # Fill width to touch right edge, small left padding
                btn.pack(fill=tk.X, pady=3, padx=(5, 0))

            # Show nav_frame BEFORE bottom_frame to maintain position, and batch redraw
            self.nav_frame.pack(fill=tk.BOTH, expand=True, padx=(3, 0), before=self.bottom_frame)
            self.nav_frame.update_idletasks()
            self.root.after(10, self._redraw_all_nav_buttons)

            # Update theme toggle to just icon (hide text, center icon)
            if hasattr(self, 'theme_text_label'):
                self.theme_text_label.pack_forget()
            if hasattr(self, 'theme_icon_label'):
                theme_icon_id = "light_mode" if self.theme_manager.is_dark else "dark_mode"
                theme_icon_image = self._tool_icons.get(theme_icon_id)
                self.theme_icon_label.pack_forget()
                if theme_icon_image:
                    self.theme_icon_label.config(image=theme_icon_image, text="", padx=0)
                else:
                    theme_icon = "◑" if self.theme_manager.is_dark else "◐"
                    self.theme_icon_label.config(text=theme_icon, padx=0)
                self.theme_icon_label.pack(anchor='center')
            # Re-center the theme toggle frame
            if self.theme_button:
                self.theme_button.pack_forget()
                self.theme_button.pack(pady=(0, 12), anchor='center')

            # Shrink logo font
            if hasattr(self, 'logo_label'):
                self.logo_label.config(font=('Segoe UI', 20, 'bold'))

        else:
            # Expand to full
            new_width = AppConstants.SIDEBAR_WIDTH
            self.sidebar_frame.config(width=new_width)

            # Update collapse button direction for wider sidebar - back to right
            if self.collapse_btn:
                self.collapse_btn.pack_forget()
                self.collapse_btn.pack(side=tk.RIGHT, padx=(0, 20), pady=2)
                self.collapse_btn.config(text="❮")  # Left-pointing (collapse)

            # Restore brand widgets (pack_forget first to avoid duplication glitch)
            for widget in self.brand_widgets:
                widget.pack_forget()
                widget.pack(anchor=tk.CENTER)

            # Restore link_frame (Help/About/Web) and divider in correct order
            if hasattr(self, 'link_frame') and self.link_frame:
                self.link_frame.pack_forget()
                self.link_frame.pack(pady=(0, 12))
            if hasattr(self, 'links_theme_divider') and self.links_theme_divider:
                self.links_theme_divider.pack_forget()
                self.links_theme_divider.pack(fill=tk.X, pady=(0, 12))

            # Hide nav_frame during transition to prevent cascade effect
            self.nav_frame.pack_forget()

            # Restore nav buttons with full text (left-rounded, right flush)
            btn_width = AppConstants.SIDEBAR_WIDTH - 5  # Full width minus left margin
            for tool_id, btn in self.nav_buttons.items():
                btn.pack_forget()
                # Switch to expanded mode (left corners rounded only)
                btn.set_mode('expanded')
                # Show icon + text (not icon only)
                btn.set_icon_only(False)
                # Resize to full width
                btn.set_size(btn_width, 40)
                # Pack with left margin only - right side touches edge
                btn.pack(fill=tk.X, pady=2, padx=(5, 0))

            # Show nav_frame BEFORE bottom_frame to maintain position, and batch redraw
            self.nav_frame.pack(fill=tk.BOTH, expand=True, padx=(8, 0), before=self.bottom_frame)
            self.nav_frame.update_idletasks()
            self.root.after(10, self._redraw_all_nav_buttons)

            # Restore theme toggle with text (side-by-side layout)
            if hasattr(self, 'theme_icon_label'):
                theme_icon_id = "light_mode" if self.theme_manager.is_dark else "dark_mode"
                theme_icon_image = self._tool_icons.get(theme_icon_id)
                self.theme_icon_label.pack_forget()
                if theme_icon_image:
                    self.theme_icon_label.config(image=theme_icon_image, text="", anchor='center')
                else:
                    theme_icon = "◑" if self.theme_manager.is_dark else "◐"
                    self.theme_icon_label.config(text=theme_icon, width=2, anchor='center')
                self.theme_icon_label.pack(side=tk.LEFT, padx=(0, 6), anchor='center')
            if hasattr(self, 'theme_text_label'):
                theme_text = "Light Mode" if self.theme_manager.is_dark else "Dark Mode"
                self.theme_text_label.config(text=theme_text)
                self.theme_text_label.pack(side=tk.LEFT, anchor='center')
            # Restore theme toggle frame to left-aligned
            if self.theme_button:
                self.theme_button.pack_forget()
                self.theme_button.pack(pady=(0, 12))

            # Restore logo font
            if hasattr(self, 'logo_label'):
                self.logo_label.config(font=('Segoe UI', 32, 'bold'))

    def _redraw_all_nav_buttons(self):
        """Redraw all nav buttons at once to avoid cascade effect"""
        for btn in self.nav_buttons.values():
            btn._draw_button()

    def _on_theme_changed(self, theme: str):
        """Handle theme change - update UI elements"""
        colors = self.theme_manager.colors

        # Update top_bar and collapse_btn FIRST to avoid flickering
        if hasattr(self, 'top_bar') and self.top_bar:
            self.top_bar.config(bg=colors['sidebar_bg'])
        if self.collapse_btn:
            self.collapse_btn.config(
                fg=colors['button_primary'],
                bg=colors['sidebar_bg']
            )
            self.collapse_btn.bind('<Enter>', lambda e: self.collapse_btn.config(fg=colors['button_primary_hover']))
            self.collapse_btn.bind('<Leave>', lambda e: self.collapse_btn.config(fg=colors['button_primary']))

        # Update sidebar
        if self.sidebar_frame:
            self.sidebar_frame.config(bg=colors['sidebar_bg'])
            self._update_sidebar_colors(self.sidebar_frame, colors)

        # Update nav buttons colors for theme change
        for tool_id, btn in self.nav_buttons.items():
            btn.update_colors(
                bg=colors['sidebar_bg'],
                hover_bg=colors['sidebar_hover'],
                pressed_bg=colors['sidebar_pressed'],
                fg=colors['nav_text'],
                active_bg=colors['sidebar_active'],
                active_hover_bg=colors['sidebar_active_hover'],
                active_pressed_bg=colors['sidebar_active_pressed'],
                parent_bg=colors['sidebar_bg']
            )
        # Re-apply active selection state
        self._update_nav_selection(self.current_tool_id)

        # Update theme toggle labels
        if hasattr(self, 'theme_icon_label'):
            theme_icon_id = "light_mode" if self.theme_manager.is_dark else "dark_mode"
            theme_icon_image = self._tool_icons.get(theme_icon_id)
            theme_icon_text = "◑" if self.theme_manager.is_dark else "◐"  # Fallback
            self.theme_icon_label.config(
                image=theme_icon_image if theme_icon_image else "",
                text="" if theme_icon_image else theme_icon_text,
                fg=colors['nav_text'],
                bg=colors['sidebar_bg']
            )
        if hasattr(self, 'theme_text_label'):
            theme_text = "Light Mode" if self.theme_manager.is_dark else "Dark Mode"
            self.theme_text_label.config(
                text=theme_text,
                fg=colors['nav_text'],
                bg=colors['sidebar_bg']
            )
        if self.theme_button:
            self.theme_button.config(bg=colors['sidebar_bg'])

        # Update footer links
        if hasattr(self, 'footer_links'):
            for link in self.footer_links:
                link.config(fg=colors['text_muted'], bg=colors['sidebar_bg'])
                # Rebind hover with new colors
                link.bind('<Enter>', lambda e, lbl=link: lbl.config(fg=colors['text_primary']))
                link.bind('<Leave>', lambda e, lbl=link: lbl.config(fg=colors['text_muted']))

        # Update link_frame background
        if hasattr(self, 'link_frame') and self.link_frame:
            self.link_frame.config(bg=colors['sidebar_bg'])

        # Update divider lines
        if hasattr(self, 'top_divider') and self.top_divider:
            self.top_divider.config(bg=colors['divider_faint'])
        if hasattr(self, 'links_theme_divider') and self.links_theme_divider:
            self.links_theme_divider.config(bg=colors['divider_faint'])

    def _update_sidebar_colors(self, widget, colors):
        """Recursively update sidebar widget colors"""
        try:
            # Skip top_bar and collapse_btn - handled explicitly to avoid flickering
            if hasattr(self, 'top_bar') and widget == self.top_bar:
                return
            if hasattr(self, 'collapse_btn') and widget == self.collapse_btn:
                return

            if isinstance(widget, tk.Frame):
                widget.config(bg=colors['sidebar_bg'])
            elif isinstance(widget, tk.Label):
                # Special handling for logo label - use button_primary (blue dark, teal light)
                if hasattr(self, 'logo_label') and widget == self.logo_label:
                    widget.config(bg=colors['sidebar_bg'], fg=colors['button_primary'])
                else:
                    fg = widget.cget('fg')
                    # Check for button_primary colors (logo uses these)
                    if fg in [AppConstants.THEMES['dark']['button_primary'],
                              AppConstants.THEMES['light']['button_primary']]:
                        widget.config(bg=colors['sidebar_bg'], fg=colors['button_primary'])
                    # Preserve primary color labels
                    elif fg == colors.get('primary') or fg == AppConstants.THEMES['dark']['primary'] or fg == AppConstants.THEMES['light']['primary']:
                        widget.config(bg=colors['sidebar_bg'], fg=colors['primary'])
                    elif 'nav_text' in str(fg) or fg in ['#c0c0d0', '#4a4a5a']:
                        widget.config(bg=colors['sidebar_bg'], fg=colors['nav_text'])
                    else:
                        widget.config(bg=colors['sidebar_bg'], fg=colors['text_muted'])
            elif isinstance(widget, tk.Button):
                # Skip nav buttons (handled separately)
                if widget not in self.nav_buttons.values() and widget != self.theme_button:
                    widget.config(
                        bg=colors['sidebar_bg'],
                        fg=colors['text_muted'],
                        activebackground=colors['sidebar_pressed']
                    )

            # Recursively update children
            for child in widget.winfo_children():
                self._update_sidebar_colors(child, colors)
        except Exception:
            pass

    def perform_tool_operation(self, **kwargs) -> bool:
        """Implementation required by EnhancedBaseExternalTool"""
        return True

    # =========================================================================
    # DIALOGS
    # =========================================================================

    def open_company_website(self):
        """Open company website"""
        try:
            import webbrowser
            webbrowser.open(AppConstants.COMPANY_WEBSITE)
        except Exception as e:
            self.show_error("Error", f"Could not open website: {e}")

    def open_github_repo(self):
        """Open GitHub repository"""
        try:
            import webbrowser
            webbrowser.open("https://github.com/analyticendeavors/pbi-multi-tool")
        except Exception as e:
            self.show_error("Error", f"Could not open GitHub: {e}")

    def open_donate_page(self):
        """Open Buy Me a Coffee donation page"""
        try:
            import webbrowser
            webbrowser.open("https://buymeacoffee.com/reidhavens")
        except Exception as e:
            self.show_error("Error", f"Could not open donation page: {e}")

    def show_help_dialog(self):
        """Show context-sensitive help based on active tool"""
        try:
            if self.current_tool_id:
                tool_tab = self.tool_manager.get_tool_tab(self.current_tool_id)
                if tool_tab and hasattr(tool_tab, 'show_help_dialog'):
                    tool_tab.show_help_dialog()
                    return

            self.show_general_help()

        except Exception as e:
            self.logger.log_security_event(f"Help dialog error: {e}", "ERROR")
            self.show_general_help()

    def show_general_help(self):
        """Show general application help"""
        colors = self.theme_manager.colors

        help_window = tk.Toplevel(self.root)
        help_window.withdraw()  # Hide initially to prevent flicker
        help_window.title("AE Power BI Tools - Help")
        help_window.geometry("600x550")
        help_window.resizable(False, False)
        help_window.transient(self.root)
        help_window.grab_set()
        help_window.configure(bg=colors['background'])

        # Set favicon icon
        try:
            help_window.iconbitmap(self.config.icon_path)
        except Exception:
            pass

        # Main container
        container = tk.Frame(help_window, bg=colors['background'], padx=25, pady=25)
        container.pack(fill=tk.BOTH, expand=True)

        # Header
        tk.Label(container, text="AE Power BI Tools",
                 font=('Segoe UI', 18, 'bold'),
                 fg=colors['button_primary'],
                 bg=colors['background']).pack(anchor=tk.W, pady=(0, 5))

        tk.Label(container, text="Professional suite for Power BI report management",
                 font=('Segoe UI', 10),
                 fg=colors['text_secondary'],
                 bg=colors['background']).pack(anchor=tk.W, pady=(0, 20))

        # Tool list
        tk.Label(container, text="Available Tools:",
                 font=('Segoe UI', 12, 'bold'),
                 fg=colors['text_primary'],
                 bg=colors['background']).pack(anchor=tk.W, pady=(0, 10))

        tools_frame = tk.Frame(container, bg=colors['background'])
        tools_frame.pack(fill=tk.BOTH, expand=True)

        for tool in self.tool_manager.get_enabled_tools():
            tool_frame = tk.Frame(tools_frame, bg=colors['background'])
            tool_frame.pack(fill=tk.X, pady=4)

            tk.Label(tool_frame, text=f"• {tool.name}",
                     font=('Segoe UI', 10, 'bold'),
                     fg=colors['button_primary'],
                     bg=colors['background']).pack(anchor=tk.W)
            tk.Label(tool_frame, text=f"   {tool.description}",
                     font=('Segoe UI', 9),
                     fg=colors['text_secondary'],
                     bg=colors['background']).pack(anchor=tk.W)

        # Close button - use tk.Button with proper theme colors
        button_frame = tk.Frame(container, bg=colors['background'])
        button_frame.pack(fill=tk.X, pady=(20, 0))

        close_btn = tk.Button(
            button_frame,
            text="Close",
            font=('Segoe UI', 10),
            fg='#ffffff',
            bg=colors['button_primary'],
            activebackground=colors['button_primary_pressed'],
            activeforeground='#ffffff',
            relief='flat',
            bd=0,
            padx=20,
            pady=8,
            cursor='hand2',
            command=help_window.destroy
        )
        close_btn.pack()
        close_btn.bind('<Enter>', lambda e: close_btn.config(bg=colors['button_primary_hover']))
        close_btn.bind('<Leave>', lambda e: close_btn.config(bg=colors['button_primary']))

        help_window.bind('<Escape>', lambda e: help_window.destroy())

        # Center dialog on parent window after content is created (no flicker)
        help_window.update_idletasks()
        dialog_width = help_window.winfo_reqwidth()
        dialog_height = help_window.winfo_reqheight()
        parent_x = self.root.winfo_rootx()
        parent_y = self.root.winfo_rooty()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        help_window.geometry(f"+{x}+{y}")

        # Set dark/light title bar BEFORE showing window to prevent white flash
        help_window.update()
        self._set_window_title_bar_color(help_window, self.theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled

    def show_error(self, title: str, message: str):
        """Show themed error dialog"""
        self._show_themed_dialog(title, message, dialog_type='error')

    def show_info(self, title: str, message: str):
        """Show themed info dialog"""
        self._show_themed_dialog(title, message, dialog_type='info')

    def _show_themed_dialog(self, title: str, message: str, dialog_type: str = 'info'):
        """Create a themed dialog that matches the app's dark/light mode"""
        from core.ui_base import RoundedButton

        colors = self.theme_manager.colors
        is_dark = self.theme_manager.is_dark

        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg=colors['background'])

        # Set dark/light title bar
        self._set_window_title_bar_color(dialog, is_dark)

        # Center on parent
        dialog.geometry(f"+{self.root.winfo_rootx() + 100}+{self.root.winfo_rooty() + 100}")

        # Main frame
        main_frame = tk.Frame(dialog, bg=colors['background'], padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Icon and message row
        content_frame = tk.Frame(main_frame, bg=colors['background'])
        content_frame.pack(fill=tk.X, pady=(0, 20))

        # Icon based on type
        icon_colors = {
            'info': colors['info'],
            'error': colors['error'],
            'warning': colors['warning']
        }
        icons = {'info': 'ℹ️', 'error': '❌', 'warning': '⚠️'}

        icon_label = tk.Label(content_frame, text=icons.get(dialog_type, 'ℹ️'),
                             font=('Segoe UI', 24),
                             bg=colors['background'],
                             fg=icon_colors.get(dialog_type, colors['info']))
        icon_label.pack(side=tk.LEFT, padx=(0, 15))

        msg_label = tk.Label(content_frame, text=message,
                            font=('Segoe UI', 10),
                            bg=colors['background'],
                            fg=colors['text_primary'],
                            justify=tk.LEFT,
                            wraplength=350)
        msg_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Button frame
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        btn = RoundedButton(
            button_frame,
            text='OK',
            command=dialog.destroy,
            bg=colors['button_primary'],
            fg='#ffffff',
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            width=80,
            height=32,
            radius=6,
            font=('Segoe UI', 10)
        )
        btn.pack(side=tk.RIGHT)

        dialog.bind('<Escape>', lambda e: dialog.destroy())
        dialog.bind('<Return>', lambda e: dialog.destroy())

        dialog.wait_window()


def main():
    """Main entry point"""
    try:
        app = EnhancedPowerBIReportToolsApp()
        app.run_with_tool_manager(app.tool_manager)
    except Exception as e:
        print(f"Critical Error: {e}")
        import traceback
        traceback.print_exc()

        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Application Error",
                                 f"The application failed to start:\n\n{e}\n\n"
                                 f"Please check the logs for more details.")
        except:
            pass


if __name__ == "__main__":
    main()
