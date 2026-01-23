"""
SVG Controls for Field Parameters
SVG-based checkbox with theme support.

Note: Radio buttons have been consolidated into LabeledRadioGroup in core.ui_base.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from typing import Optional, Callable, Dict, Any
from pathlib import Path
import io
import logging

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


class SVGCheckbox:
    """
    An SVG-based checkbox with theme support.

    Usage:
        checkbox = SVGCheckbox(
            parent_frame,
            text="Enable feature",
            variable=my_bool_var,
            command=on_toggle
        )
    """

    def __init__(
        self,
        parent: tk.Widget,
        text: str,
        variable: tk.BooleanVar,
        command: Optional[Callable] = None,
        font: tuple = ('Segoe UI', 9),
        state: str = "normal",  # "normal" or "disabled"
        bg: Optional[str] = None  # Custom background color (defaults to colors['background'])
    ):
        self.parent = parent
        self.text = text
        self.variable = variable
        self.command = command
        self.font = font
        self._state = state
        self._custom_bg = bg  # Store custom bg for theme updates

        self.logger = logging.getLogger(__name__)
        self._theme_manager = get_theme_manager()

        # Store icons
        self._icons: Dict[str, Any] = {}

        # Load icons
        self._load_icons()

        # Create the widget - use section_bg as default for panel contexts
        colors = self._theme_manager.colors
        bg_color = bg if bg else colors.get('section_bg', colors['background'])

        self.frame = tk.Frame(parent, bg=bg_color)

        # Icon label
        is_checked = self.variable.get()
        icon = self._get_current_icon(is_checked)

        self.icon_label = tk.Label(self.frame, bg=bg_color, cursor='hand2')
        if icon:
            self.icon_label.configure(image=icon)
            self.icon_label._icon_ref = icon
        self.icon_label.pack(side=tk.LEFT, padx=(0, 6))

        # Text label
        text_fg = colors['text_muted'] if state == "disabled" else colors['text_primary']
        self.text_label = tk.Label(
            self.frame, text=text, bg=bg_color, fg=text_fg,
            font=self.font, cursor='hand2'
        )
        self.text_label.pack(side=tk.LEFT)

        # Click handler
        self.icon_label.bind('<Button-1>', self._on_click)
        self.text_label.bind('<Button-1>', self._on_click)

        # Hover underline effect
        self.text_label.bind('<Enter>', self._on_enter)
        self.text_label.bind('<Leave>', self._on_leave)

        # Trace variable for updates
        self.variable.trace_add("write", self._on_variable_changed)

    def _load_icons(self):
        """Load checkbox SVG icons"""
        if not PIL_AVAILABLE:
            return

        icons_dir = Path(__file__).parent.parent.parent.parent / "assets" / "Tool Icons"
        is_dark = self._theme_manager.is_dark

        # Load appropriate icons for current theme
        checked_icon = "box-checked-dark" if is_dark else "box-checked"
        unchecked_icon = "box-dark" if is_dark else "box"

        for icon_name, key in [(checked_icon, "checked"), (unchecked_icon, "unchecked")]:
            icon = self._load_svg_icon(icons_dir / f"{icon_name}.svg", size=16)
            if icon:
                self._icons[key] = icon

    def _load_svg_icon(self, path: Path, size: int = 16):
        """Load an SVG icon and return PhotoImage"""
        if not PIL_AVAILABLE:
            return None

        try:
            img = None

            if CAIROSVG_AVAILABLE and path.exists():
                png_data = cairosvg.svg2png(
                    url=str(path),
                    output_width=size * 4,
                    output_height=size * 4
                )
                img = Image.open(io.BytesIO(png_data))

            if img is None:
                return None

            img = img.resize((size, size), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception as e:
            self.logger.debug(f"Failed to load icon {path}: {e}")
            return None

    def _get_current_icon(self, is_checked: bool):
        """Get the appropriate icon for current state"""
        return self._icons.get('checked' if is_checked else 'unchecked')

    def _on_click(self, event=None):
        """Handle click event"""
        if self._state == "disabled":
            return

        # Toggle the value
        self.variable.set(not self.variable.get())

        if self.command:
            self.command()

    def _on_enter(self, event=None):
        """Handle mouse enter"""
        if self._state != "disabled":
            self.text_label.configure(font=self.font + ('underline',))

    def _on_leave(self, event=None):
        """Handle mouse leave"""
        self.text_label.configure(font=self.font)

    def _on_variable_changed(self, *args):
        """Update visual state when variable changes"""
        is_checked = self.variable.get()
        icon = self._get_current_icon(is_checked)

        if icon:
            self.icon_label.configure(image=icon)
            self.icon_label._icon_ref = icon

    def config(self, **kwargs):
        """Configure checkbox properties"""
        if 'state' in kwargs:
            self._state = kwargs['state']
            colors = self._theme_manager.colors

            if self._state == "disabled":
                self.text_label.configure(fg=colors['text_muted'])
                self.icon_label.configure(cursor='arrow')
                self.text_label.configure(cursor='arrow')
            else:
                self.text_label.configure(fg=colors['text_primary'])
                self.icon_label.configure(cursor='hand2')
                self.text_label.configure(cursor='hand2')

    def on_theme_changed(self, bg: Optional[str] = None):
        """Update widget colors when theme changes

        Args:
            bg: Optional background color override. If not provided, uses section_bg.
        """
        # Reload icons for new theme
        self._load_icons()

        colors = self._theme_manager.colors
        bg_color = bg if bg else colors.get('section_bg', colors['background'])

        self.frame.configure(bg=bg_color)
        self.icon_label.configure(bg=bg_color)

        text_fg = colors['text_muted'] if self._state == "disabled" else colors['text_primary']
        self.text_label.configure(bg=bg_color, fg=text_fg)

        # Update icon for current state
        is_checked = self.variable.get()
        icon = self._get_current_icon(is_checked)
        if icon:
            self.icon_label.configure(image=icon)
            self.icon_label._icon_ref = icon

    def pack(self, **kwargs):
        """Pack the container frame"""
        self.frame.pack(**kwargs)

    def grid(self, **kwargs):
        """Grid the container frame"""
        self.frame.grid(**kwargs)

    def place(self, **kwargs):
        """Place the container frame"""
        self.frame.place(**kwargs)
