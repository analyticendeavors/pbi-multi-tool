"""
Inline Target Picker - Popup for inline target selection in Treeview
Built by Reid Havens of Analytic Endeavors

A dropdown-style popup that appears when clicking on the Target column in the
Connection Mappings table. Shows available local models and a "Browse Cloud..."
option for selecting cloud targets.
"""

import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional
from pathlib import Path

from core.theme_manager import get_theme_manager
from tools.connection_hotswap.models import SwapTarget


class InlineTargetPicker:
    """
    Popup menu for inline target selection in a Treeview cell.

    Shows local models and cloud browse option when the user clicks
    on the Target column.
    """

    def __init__(
        self,
        parent: tk.Widget,
        get_local_models: Callable[[], List[SwapTarget]],
        on_cloud_browse: Callable[[str], None],
        on_target_selected: Callable[[str, SwapTarget], None]
    ):
        """
        Initialize the inline target picker.

        Args:
            parent: Parent widget (the treeview container)
            get_local_models: Callback to get list of available local models
            on_cloud_browse: Callback when "Browse Cloud..." is selected, receives item_id
            on_target_selected: Callback when a target is selected, receives (item_id, SwapTarget)
        """
        self._parent = parent
        self._get_local_models = get_local_models
        self._on_cloud_browse = on_cloud_browse
        self._on_target_selected = on_target_selected
        self._theme_manager = get_theme_manager()

        self._popup: Optional[tk.Toplevel] = None
        self._current_item_id: Optional[str] = None
        self._icons: Dict[str, tk.PhotoImage] = {}
        self._load_icons()

    def _load_icons(self):
        """Load SVG icons for the picker menu."""
        try:
            import cairosvg
            from PIL import Image, ImageTk
            import io

            # Find icons directory
            icons_dir = Path(__file__).parent.parent.parent.parent.parent / "assets" / "icons"

            is_dark = self._theme_manager.is_dark
            icon_size = 16

            # Map icon names to SVG files
            icon_map = {
                'local': 'letter-l.svg',
                'cloud': 'letter-c.svg',
            }

            for icon_name, svg_file in icon_map.items():
                svg_path = icons_dir / svg_file
                if svg_path.exists():
                    try:
                        # Read and colorize SVG
                        svg_content = svg_path.read_text(encoding='utf-8')
                        # Apply theme color
                        icon_color = '#b0b0b0' if is_dark else '#666666'
                        svg_content = svg_content.replace('currentColor', icon_color)

                        # Convert SVG to PNG
                        png_data = cairosvg.svg2png(
                            bytestring=svg_content.encode('utf-8'),
                            output_width=icon_size,
                            output_height=icon_size
                        )

                        # Create PhotoImage
                        image = Image.open(io.BytesIO(png_data))
                        self._icons[icon_name] = ImageTk.PhotoImage(image)
                    except Exception:
                        pass
        except ImportError:
            pass  # cairosvg or PIL not available

    def show_picker(self, item_id: str, x: int, y: int, width: int = 200):
        """
        Show the picker popup at the specified position.

        Args:
            item_id: The treeview item ID for the row being edited
            x: X coordinate for popup position (screen coords)
            y: Y coordinate for popup position (screen coords)
            width: Width of the popup
        """
        # Close any existing popup
        self.hide_picker()

        self._current_item_id = item_id
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create popup window
        self._popup = tk.Toplevel(self._parent)
        self._popup.overrideredirect(True)  # No window decorations
        self._popup.attributes('-topmost', True)

        # Configure appearance
        popup_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')

        self._popup.configure(bg=border_color)  # Border via 1px bg showing through

        # Main container with slight padding for border effect
        container = tk.Frame(self._popup, bg=popup_bg, padx=1, pady=1)
        container.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # Get local models
        local_models = self._get_local_models()

        # Build menu items
        self._build_menu(container, local_models)

        # Position popup
        self._popup.geometry(f"+{x}+{y}")
        self._popup.update_idletasks()

        # Ensure minimum width
        current_width = self._popup.winfo_reqwidth()
        if current_width < width:
            self._popup.geometry(f"{width}x{self._popup.winfo_reqheight()}+{x}+{y}")

        # Bind events to close popup
        self._popup.bind('<FocusOut>', lambda e: self.hide_picker())
        self._popup.bind('<Escape>', lambda e: self.hide_picker())

        # Focus the popup
        self._popup.focus_set()

        # Bind global click to close
        self._popup.bind('<Button-1>', self._on_popup_click)
        self._parent.winfo_toplevel().bind('<Button-1>', self._on_outside_click, add='+')

    def _build_menu(self, container: tk.Frame, local_models: List[SwapTarget]):
        """Build the menu items in the popup."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        bg_color = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        hover_bg = colors.get('hover', '#2a2a3e' if is_dark else '#f0f0f5')
        text_color = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        muted_color = colors.get('text_muted', '#888888')
        primary_color = colors.get('primary', '#4a6cf5')

        # Popup title header
        title_header = tk.Label(
            container,
            text="Target Presets (Quick Swap)",
            font=('Segoe UI', 9, 'bold'),
            fg=text_color,
            bg=bg_color,
            anchor='w',
            padx=10,
            pady=6
        )
        title_header.pack(fill=tk.X)

        # Title separator
        title_sep = tk.Frame(container, height=1, bg=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'))
        title_sep.pack(fill=tk.X, padx=6, pady=(0, 4))

        # Local Models section header
        if local_models:
            header = tk.Label(
                container,
                text="Local Models",
                font=('Segoe UI', 8),
                fg=muted_color,
                bg=bg_color,
                anchor='w',
                padx=10,
                pady=4
            )
            header.pack(fill=tk.X)

            # Local model items
            for model in local_models:
                item_frame = tk.Frame(container, bg=bg_color)
                item_frame.pack(fill=tk.X)

                # Local icon (letter-l.svg)
                local_icon = self._icons.get('local')
                if local_icon:
                    icon_label = tk.Label(
                        item_frame,
                        image=local_icon,
                        bg=bg_color,
                        padx=6
                    )
                    icon_label._icon_ref = local_icon  # Keep reference
                else:
                    icon_label = tk.Label(
                        item_frame,
                        text="L",
                        font=('Segoe UI', 8),
                        fg=muted_color,
                        bg=bg_color,
                        padx=6
                    )
                icon_label.pack(side=tk.LEFT)

                # Model name
                name_label = tk.Label(
                    item_frame,
                    text=model.display_name,
                    font=('Segoe UI', 9),
                    fg=text_color,
                    bg=bg_color,
                    anchor='w',
                    padx=4,
                    pady=6
                )
                name_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

                # Bind hover effects
                for widget in [item_frame, icon_label, name_label]:
                    widget.bind('<Enter>', lambda e, f=item_frame: self._on_item_enter(f, hover_bg))
                    widget.bind('<Leave>', lambda e, f=item_frame: self._on_item_leave(f, bg_color))
                    widget.bind('<Button-1>', lambda e, m=model: self._select_target(m))
        else:
            # No local models message
            no_models = tk.Label(
                container,
                text="No local models detected",
                font=('Segoe UI', 9, 'italic'),
                fg=muted_color,
                bg=bg_color,
                anchor='w',
                padx=10,
                pady=8
            )
            no_models.pack(fill=tk.X)

        # Separator
        sep = tk.Frame(container, height=1, bg=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'))
        sep.pack(fill=tk.X, pady=4, padx=6)

        # Cloud section header
        cloud_header = tk.Label(
            container,
            text="Cloud",
            font=('Segoe UI', 8),
            fg=muted_color,
            bg=bg_color,
            anchor='w',
            padx=10,
            pady=4
        )
        cloud_header.pack(fill=tk.X)

        # Browse Cloud option
        cloud_frame = tk.Frame(container, bg=bg_color)
        cloud_frame.pack(fill=tk.X)

        # Cloud icon (letter-c.svg)
        cloud_icon_img = self._icons.get('cloud')
        if cloud_icon_img:
            cloud_icon = tk.Label(
                cloud_frame,
                image=cloud_icon_img,
                bg=bg_color,
                padx=6
            )
            cloud_icon._icon_ref = cloud_icon_img  # Keep reference
        else:
            cloud_icon = tk.Label(
                cloud_frame,
                text="C",
                font=('Segoe UI', 8),
                fg=primary_color,
                bg=bg_color,
                padx=6
            )
        cloud_icon.pack(side=tk.LEFT)

        cloud_label = tk.Label(
            cloud_frame,
            text="Browse Cloud...",
            font=('Segoe UI', 9),
            fg=primary_color,
            bg=bg_color,
            anchor='w',
            padx=4,
            pady=6
        )
        cloud_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Bind hover and click for cloud option
        for widget in [cloud_frame, cloud_icon, cloud_label]:
            widget.bind('<Enter>', lambda e, f=cloud_frame: self._on_item_enter(f, hover_bg))
            widget.bind('<Leave>', lambda e, f=cloud_frame: self._on_item_leave(f, bg_color))
            widget.bind('<Button-1>', lambda e: self._browse_cloud())

    def _on_item_enter(self, frame: tk.Frame, hover_bg: str):
        """Handle mouse enter on menu item."""
        frame.configure(bg=hover_bg)
        for child in frame.winfo_children():
            child.configure(bg=hover_bg)

    def _on_item_leave(self, frame: tk.Frame, normal_bg: str):
        """Handle mouse leave on menu item."""
        frame.configure(bg=normal_bg)
        for child in frame.winfo_children():
            child.configure(bg=normal_bg)

    def _select_target(self, target: SwapTarget):
        """Handle selection of a local model target."""
        if self._current_item_id:
            self._on_target_selected(self._current_item_id, target)
        self.hide_picker()

    def _browse_cloud(self):
        """Handle browse cloud selection."""
        item_id = self._current_item_id
        self.hide_picker()
        if item_id:
            self._on_cloud_browse(item_id)

    def _on_popup_click(self, event):
        """Handle click inside popup - don't close."""
        pass  # Prevent propagation

    def _on_outside_click(self, event):
        """Handle click outside popup - close it."""
        if self._popup and self._popup.winfo_exists():
            # Check if click is outside the popup
            popup_x = self._popup.winfo_rootx()
            popup_y = self._popup.winfo_rooty()
            popup_w = self._popup.winfo_width()
            popup_h = self._popup.winfo_height()

            if not (popup_x <= event.x_root <= popup_x + popup_w and
                    popup_y <= event.y_root <= popup_y + popup_h):
                self.hide_picker()

    def hide_picker(self):
        """Hide and destroy the picker popup."""
        if self._popup and self._popup.winfo_exists():
            # Unbind global click handler
            try:
                self._parent.winfo_toplevel().unbind('<Button-1>')
            except:
                pass
            self._popup.destroy()
            self._popup = None
        self._current_item_id = None

    def is_visible(self) -> bool:
        """Check if the picker is currently visible."""
        return self._popup is not None and self._popup.winfo_exists()
