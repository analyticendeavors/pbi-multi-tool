"""
ParameterSelectorDropdown
Custom themed dropdown for selecting Field Parameters.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable, List
import logging
from pathlib import Path
import io

from core.theme_manager import get_theme_manager

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


class ParameterSelectorDropdown:
    """
    A custom themed dropdown for selecting Field Parameters.
    Features:
    - Themed background and borders
    - Hover effects on items
    - Click-away to close
    - Theme change support
    - Field Parameter icon for each item
    """

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager=None,
        on_selection_changed: Optional[Callable[[str], None]] = None,
        width: int = 250
    ):
        """
        Initialize the parameter selector dropdown.

        Args:
            parent: Parent widget to attach to
            theme_manager: Theme manager for colors (optional, uses global if not provided)
            on_selection_changed: Callback when selection changes, receives parameter name
            width: Width of the dropdown button
        """
        self._parent = parent
        self._theme_manager = theme_manager or get_theme_manager()
        self._on_selection_changed = on_selection_changed
        self._width = width
        self._logger = logging.getLogger(__name__)

        # State
        self._parameters: List[str] = []
        self._selected_value: str = ""
        self._dropdown_popup = None
        self._item_labels: List[tk.Label] = []
        self._param_icon = None

        # Load Field Parameter icon
        self._load_icons()

        # Create the dropdown button
        self._create_dropdown_button()

    def _load_icons(self):
        """Load Field Parameter icon for dropdown items."""
        if not PIL_AVAILABLE:
            return

        icons_dir = Path(__file__).parent.parent.parent.parent / "assets" / "Tool Icons"
        svg_path = icons_dir / "Field Parameter.svg"
        png_path = icons_dir / "Field Parameter.png"

        try:
            img = None
            if CAIROSVG_AVAILABLE and svg_path.exists():
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=64,
                    output_height=64
                )
                img = Image.open(io.BytesIO(png_data))
            elif png_path.exists():
                img = Image.open(png_path)

            if img:
                img = img.resize((16, 16), Image.Resampling.LANCZOS)
                self._param_icon = ImageTk.PhotoImage(img)
        except Exception as e:
            self._logger.debug(f"Failed to load Field Parameter icon: {e}")

    def _create_dropdown_button(self):
        """Create the dropdown trigger button."""
        colors = self._theme_manager.colors

        # Main frame container
        self.frame = tk.Frame(self._parent, bg=colors['background'])

        # Button frame with border
        self._btn_frame = tk.Frame(
            self.frame,
            bg=colors['border'],
            padx=1,
            pady=1
        )
        self._btn_frame.pack(fill=tk.X)

        # Inner button area
        self._btn_inner = tk.Frame(
            self._btn_frame,
            bg=colors.get('input_bg', colors['surface']),
            cursor='hand2'
        )
        self._btn_inner.pack(fill=tk.X)

        # Icon (Field Parameter)
        if self._param_icon:
            self._icon_label = tk.Label(
                self._btn_inner,
                image=self._param_icon,
                bg=colors.get('input_bg', colors['surface'])
            )
            self._icon_label.pack(side=tk.LEFT, padx=(8, 4), pady=6)
            self._icon_label.bind('<Button-1>', lambda e: self._toggle_dropdown())

        # Text label showing current selection
        self._text_label = tk.Label(
            self._btn_inner,
            text="Select a parameter...",
            font=('Segoe UI', 9),
            bg=colors.get('input_bg', colors['surface']),
            fg=colors['text_muted'],
            anchor='w',
            cursor='hand2'
        )
        self._text_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0), pady=6)
        self._text_label.bind('<Button-1>', lambda e: self._toggle_dropdown())

        # Dropdown arrow
        self._arrow_label = tk.Label(
            self._btn_inner,
            text="\u25BC",  # Down arrow
            font=('Segoe UI', 8),
            bg=colors.get('input_bg', colors['surface']),
            fg=colors['text_muted'],
            cursor='hand2'
        )
        self._arrow_label.pack(side=tk.RIGHT, padx=(0, 8), pady=6)
        self._arrow_label.bind('<Button-1>', lambda e: self._toggle_dropdown())

        # Bind click on inner frame
        self._btn_inner.bind('<Button-1>', lambda e: self._toggle_dropdown())

    def _toggle_dropdown(self):
        """Toggle the dropdown popup visibility."""
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._close_dropdown()
        else:
            self._open_dropdown()

    def _open_dropdown(self):
        """Open the dropdown popup."""
        if self._dropdown_popup:
            self._close_dropdown()

        colors = self._theme_manager.colors
        popup_bg = colors.get('surface', colors['background'])
        border_color = colors['border']

        # Create popup window
        self._dropdown_popup = tk.Toplevel(self._parent)
        self._dropdown_popup.withdraw()
        self._dropdown_popup.overrideredirect(True)

        # Border frame
        border_frame = tk.Frame(
            self._dropdown_popup,
            bg=border_color,
            padx=1,
            pady=1
        )
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Content frame
        content_frame = tk.Frame(border_frame, bg=popup_bg)
        content_frame.pack(fill=tk.BOTH, expand=True)

        self._item_labels = []

        if not self._parameters:
            # Empty state
            empty_label = tk.Label(
                content_frame,
                text="No parameters found.\nConnect to a model first.",
                font=('Segoe UI', 9, 'italic'),
                bg=popup_bg,
                fg=colors['text_muted'],
                justify=tk.CENTER,
                pady=16,
                padx=16
            )
            empty_label.pack()
        else:
            # Create item for each parameter
            for param_name in self._parameters:
                item_frame = tk.Frame(content_frame, bg=popup_bg)
                item_frame.pack(fill=tk.X)

                # Icon
                if self._param_icon:
                    icon_lbl = tk.Label(
                        item_frame,
                        image=self._param_icon,
                        bg=popup_bg
                    )
                    icon_lbl.pack(side=tk.LEFT, padx=(8, 4), pady=4)
                    icon_lbl._param = param_name
                    icon_lbl.bind('<Button-1>', lambda e, n=param_name: self._select_item(n))
                    icon_lbl.bind('<Enter>', lambda e, f=item_frame: self._on_item_hover(f, True))
                    icon_lbl.bind('<Leave>', lambda e, f=item_frame: self._on_item_hover(f, False))

                # Text
                text_lbl = tk.Label(
                    item_frame,
                    text=param_name,
                    font=('Segoe UI', 9),
                    bg=popup_bg,
                    fg=colors['text_primary'],
                    anchor='w',
                    cursor='hand2'
                )
                text_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 8), pady=4)
                text_lbl._param = param_name
                text_lbl.bind('<Button-1>', lambda e, n=param_name: self._select_item(n))
                text_lbl.bind('<Enter>', lambda e, f=item_frame: self._on_item_hover(f, True))
                text_lbl.bind('<Leave>', lambda e, f=item_frame: self._on_item_hover(f, False))

                # Bind frame events
                item_frame.bind('<Button-1>', lambda e, n=param_name: self._select_item(n))
                item_frame.bind('<Enter>', lambda e, f=item_frame: self._on_item_hover(f, True))
                item_frame.bind('<Leave>', lambda e, f=item_frame: self._on_item_hover(f, False))

                self._item_labels.append((item_frame, text_lbl))

        # Position popup below button
        self._dropdown_popup.update_idletasks()
        btn_x = self._btn_frame.winfo_rootx()
        btn_y = self._btn_frame.winfo_rooty() + self._btn_frame.winfo_height()
        btn_width = self._btn_frame.winfo_width()

        # Make popup at least as wide as button
        popup_width = max(self._dropdown_popup.winfo_reqwidth(), btn_width)
        self._dropdown_popup.geometry(f"{popup_width}x{self._dropdown_popup.winfo_reqheight()}+{btn_x}+{btn_y}")

        self._dropdown_popup.deiconify()
        self._dropdown_popup.lift()
        self._dropdown_popup.focus_set()

        # Bind click outside to close
        self._parent.winfo_toplevel().bind('<Button-1>', self._on_click_outside, add='+')

    def _on_item_hover(self, frame: tk.Frame, entering: bool):
        """Handle hover effect on dropdown items."""
        colors = self._theme_manager.colors
        if entering:
            hover_bg = colors.get('card_surface_hover', colors.get('surface', colors['background']))
        else:
            hover_bg = colors.get('surface', colors['background'])

        frame.config(bg=hover_bg)
        for child in frame.winfo_children():
            try:
                child.config(bg=hover_bg)
            except tk.TclError:
                pass

    def _select_item(self, param_name: str):
        """Handle item selection."""
        self._selected_value = param_name
        self._update_display()
        self._close_dropdown()

        if self._on_selection_changed:
            self._on_selection_changed(param_name)

    def _update_display(self):
        """Update the button text to show current selection."""
        colors = self._theme_manager.colors
        if self._selected_value:
            self._text_label.config(
                text=self._selected_value,
                fg=colors['text_primary']
            )
        else:
            self._text_label.config(
                text="Select a parameter...",
                fg=colors['text_muted']
            )

    def _on_click_outside(self, event):
        """Handle click outside the dropdown."""
        if not self._dropdown_popup or not self._dropdown_popup.winfo_exists():
            return

        # Check if click is outside dropdown
        x, y = event.x_root, event.y_root
        dx = self._dropdown_popup.winfo_rootx()
        dy = self._dropdown_popup.winfo_rooty()
        dw = self._dropdown_popup.winfo_width()
        dh = self._dropdown_popup.winfo_height()

        # Also check if click is on the button itself
        bx = self._btn_frame.winfo_rootx()
        by = self._btn_frame.winfo_rooty()
        bw = self._btn_frame.winfo_width()
        bh = self._btn_frame.winfo_height()

        if not (dx <= x <= dx + dw and dy <= y <= dy + dh):
            if not (bx <= x <= bx + bw and by <= y <= by + bh):
                self._close_dropdown()

    def _close_dropdown(self):
        """Close the dropdown popup."""
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._dropdown_popup.destroy()
        self._dropdown_popup = None

        # Unbind handlers
        try:
            self._parent.winfo_toplevel().unbind('<Button-1>')
        except Exception:
            pass

    def set_parameters(self, parameters: List[str]):
        """
        Set the list of available parameters.

        Args:
            parameters: List of parameter names
        """
        self._parameters = parameters
        # If currently selected parameter is no longer available, clear selection
        if self._selected_value and self._selected_value not in parameters:
            self._selected_value = ""
            self._update_display()

    def set_selection(self, param_name: str):
        """Set the currently selected parameter by name."""
        self._selected_value = param_name
        self._update_display()

    def get_selection(self) -> str:
        """Get the currently selected parameter name."""
        return self._selected_value

    def clear(self):
        """Clear the selection and parameters list."""
        self._parameters = []
        self._selected_value = ""
        self._update_display()

    def config(self, **kwargs):
        """Configure widget options (for compatibility with ttk.Combobox)."""
        if 'state' in kwargs:
            state = kwargs['state']
            if state == 'disabled':
                self._btn_inner.config(cursor='')
                self._text_label.config(cursor='')
                self._arrow_label.config(cursor='')
            else:
                self._btn_inner.config(cursor='hand2')
                self._text_label.config(cursor='hand2')
                self._arrow_label.config(cursor='hand2')

    def on_theme_changed(self):
        """Update widget colors when theme changes."""
        colors = self._theme_manager.colors
        bg_color = colors['background']
        input_bg = colors.get('input_bg', colors['surface'])
        border_color = colors['border']

        # Update frame
        self.frame.config(bg=bg_color)
        self._btn_frame.config(bg=border_color)
        self._btn_inner.config(bg=input_bg)

        # Update icon label
        if hasattr(self, '_icon_label'):
            self._icon_label.config(bg=input_bg)

        # Update text label
        if self._selected_value:
            self._text_label.config(bg=input_bg, fg=colors['text_primary'])
        else:
            self._text_label.config(bg=input_bg, fg=colors['text_muted'])

        # Update arrow
        self._arrow_label.config(bg=input_bg, fg=colors['text_muted'])

        # Reload icons for theme
        self._load_icons()

        # Close popup if open (will reopen with new theme)
        if self._dropdown_popup and self._dropdown_popup.winfo_exists():
            self._close_dropdown()

    def pack(self, **kwargs):
        """Pack the dropdown frame."""
        self.frame.pack(**kwargs)

    def grid(self, **kwargs):
        """Grid the dropdown frame."""
        self.frame.grid(**kwargs)

    def place(self, **kwargs):
        """Place the dropdown frame."""
        self.frame.place(**kwargs)
