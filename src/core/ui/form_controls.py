"""
Form Controls - Toggle switches and radio button groups with SVG icons.
Reusable form input components for the AE Multi-Tool application.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
import io
from pathlib import Path
from typing import Callable

from core.theme_manager import get_theme_manager
from core.constants import AppConstants

# Alias for theme colors from constants
THEME_COLORS = AppConstants.THEMES

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

# Deferred import to avoid circular dependency
def _get_tooltip():
    from core.ui.dialogs import Tooltip
    return Tooltip


class SVGToggle(tk.Canvas):
    """
    A toggle switch using SVG icons for on/off states.
    Reusable component for binary choices (e.g., theme conflict selection).
    """

    def __init__(self, parent,
                 svg_on: str,
                 svg_off: str,
                 command: Callable = None,
                 initial_state: bool = True,
                 width: int = 60,
                 height: int = 28,
                 theme_manager=None,
                 **kwargs):
        """
        Args:
            parent: Parent widget
            svg_on: Path to SVG for "on" state (e.g., toggle right/green)
            svg_off: Path to SVG for "off" state (e.g., toggle left/gray)
            command: Callback function(state: bool) called when toggled
            initial_state: Starting state (True = on, False = off)
            width: Canvas width
            height: Canvas height
            theme_manager: ThemeManager instance for theme-aware background
        """
        self._theme_manager = theme_manager or get_theme_manager()
        colors = self._theme_manager.colors

        # Get parent background
        self._bg_color = self._get_parent_background(parent, colors)

        super().__init__(parent, width=width, height=height,
                        bg=self._bg_color, highlightthickness=0, **kwargs)

        self._width = width
        self._height = height
        self._state = initial_state
        self.command = command
        self._enabled = True

        # Load SVG icons
        self._icon_on = self._load_svg(svg_on, width, height)
        self._icon_off = self._load_svg(svg_off, width, height)

        # Draw initial state
        self._draw_toggle()

        # Event bindings
        self.bind('<ButtonRelease-1>', self._on_click)
        self.config(cursor='hand2')

        # Register for theme changes
        if self._theme_manager:
            self._theme_manager.register_theme_callback(self._on_theme_change)

    def _get_parent_background(self, parent, colors) -> str:
        """Get parent background color for seamless blending"""
        try:
            return parent.cget('bg')
        except Exception:
            pass
        try:
            style = ttk.Style()
            parent_style = parent.cget('style')
            if parent_style:
                bg = style.lookup(parent_style, 'background')
                if bg:
                    return bg
        except Exception:
            pass
        return colors.get('section_bg', '#f5f5f7')

    def _load_svg(self, svg_path: str, width: int, height: int):
        """Load SVG and convert to PhotoImage"""
        try:
            path = Path(svg_path)

            if not path.exists():
                return None

            if CAIROSVG_AVAILABLE:
                # Render at 2x for quality, then the image displays at target size
                png_data = cairosvg.svg2png(
                    url=str(path),
                    output_width=width * 2,
                    output_height=height * 2
                )
                img = Image.open(io.BytesIO(png_data))
            elif PIL_AVAILABLE:
                # Fallback: try PNG with same name
                png_path = path.with_suffix('.png')
                if png_path.exists():
                    img = Image.open(png_path)
                else:
                    return None
            else:
                return None

            # Ensure RGBA
            if img.mode != 'RGBA':
                img = img.convert('RGBA')

            # Resize to target
            img = img.resize((width, height), Image.Resampling.LANCZOS)

            return ImageTk.PhotoImage(img)

        except Exception:
            return None

    def _draw_toggle(self):
        """Draw the toggle with current state's icon"""
        self.delete('all')
        icon = self._icon_on if self._state else self._icon_off
        if icon:
            self.create_image(self._width // 2, self._height // 2,
                            image=icon, anchor='center', tags='icon')

    def _on_click(self, event):
        """Handle click - toggle state and call command"""
        if not self._enabled:
            return

        self._state = not self._state
        self._draw_toggle()

        if self.command:
            self.command(self._state)

    def _on_theme_change(self, theme: str):
        """Update background when theme changes - match parent's background"""
        # Get background from parent widget (like ToggleSwitch does)
        try:
            parent_bg = self.master.cget('bg')
            self._bg_color = parent_bg
        except Exception:
            # Fallback to section_bg if parent bg can't be retrieved
            is_dark = theme == 'dark'
            colors = THEME_COLORS['dark'] if is_dark else THEME_COLORS['light']
            self._bg_color = colors.get('section_bg', '#f5f5f7')
        self.config(bg=self._bg_color)

    def get_state(self) -> bool:
        """Return current toggle state"""
        return self._state

    def set_state(self, state: bool, trigger_command: bool = True):
        """
        Set toggle state programmatically.

        Args:
            state: New state (True = on, False = off)
            trigger_command: Whether to call the command callback
        """
        if self._state != state:
            self._state = state
            self._draw_toggle()
            if trigger_command and self.command:
                self.command(self._state)

    def set_enabled(self, enabled: bool):
        """Enable or disable the toggle"""
        self._enabled = enabled
        self.config(cursor='hand2' if enabled else 'arrow')


class LabeledToggle(tk.Frame):
    """
    Toggle switch with optional text label, using SVG icons.
    Supports BooleanVar binding for easy state management.
    Uses toggle-on.svg and toggle-off.svg from assets/Tool Icons.
    """

    # Class-level icon cache to avoid reloading
    _icon_cache = {}

    def __init__(self, parent, variable: tk.BooleanVar = None, text: str = "",
                 command: Callable = None, icon_height: int = 18, **kwargs):
        """
        Args:
            parent: Parent widget
            variable: BooleanVar to bind toggle state (created if not provided)
            text: Optional text label displayed next to toggle
            command: Callback function called when toggled (no args)
            icon_height: Height of toggle icon in pixels (width auto-calculated)
        """
        self._theme_manager = get_theme_manager()
        colors = self._theme_manager.colors

        super().__init__(parent, bg=colors['background'], **kwargs)

        self._variable = variable or tk.BooleanVar(value=False)
        self._command = command
        self._text = text
        self._icon_height = icon_height
        self._enabled = True

        # Load icons (toggle SVG is 512x240, aspect ratio ~2.13:1)
        icon_width = int(icon_height * 2.13)
        self._load_icons(icon_width, icon_height)

        # Icon label (toggle button) - LEFT side
        self._icon_label = tk.Label(self, bg=colors['background'], cursor='hand2')
        self._icon_label.pack(side=tk.LEFT)

        # Text label if provided - RIGHT side
        self._label = None
        if text:
            self._label = tk.Label(self, text=text,
                                   bg=colors['background'], fg=colors['text_primary'],
                                   font=('Segoe UI', 9), cursor='hand2')
            self._label.pack(side=tk.LEFT, padx=(6, 0))

        # Update display
        self._update_display()

        # Bind events
        self._icon_label.bind('<Button-1>', self._on_click)
        if self._label:
            self._label.bind('<Button-1>', self._on_click)

        # Track variable changes
        self._variable.trace_add('write', lambda *args: self._update_display())

        # Register for theme changes
        self._theme_manager.register_theme_callback(self._handle_theme_change)

    def _load_icons(self, width: int, height: int):
        """Load toggle-on and toggle-off SVG icons"""
        cache_key = f"labeled_toggle_{width}x{height}"
        if cache_key in LabeledToggle._icon_cache:
            self._icon_on, self._icon_off = LabeledToggle._icon_cache[cache_key]
            return

        # Find assets directory - 3 levels up from core/ui/
        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"

        self._icon_on = self._load_svg_icon(icons_dir / "toggle-on.svg", width, height)
        self._icon_off = self._load_svg_icon(icons_dir / "toggle-off.svg", width, height)

        LabeledToggle._icon_cache[cache_key] = (self._icon_on, self._icon_off)

    def _load_svg_icon(self, path, width: int, height: int):
        """Load and resize an SVG icon"""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return None
        try:
            png_data = cairosvg.svg2png(url=str(path), output_width=width*2, output_height=height*2)
            img = Image.open(io.BytesIO(png_data))
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            img = img.resize((width, height), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception as e:
            return None

    def _update_display(self):
        """Update the icon based on current value"""
        is_on = self._variable.get()
        icon = self._icon_on if is_on else self._icon_off
        if icon:
            self._icon_label.configure(image=icon)
            self._icon_label.image = icon  # Keep reference

    def _on_click(self, event=None):
        """Handle click"""
        if self._enabled:
            self._variable.set(not self._variable.get())
            if self._command:
                self._command()

    def _handle_theme_change(self, theme: str):
        """Handle theme change"""
        colors = self._theme_manager.colors
        self.configure(bg=colors['background'])
        self._icon_label.configure(bg=colors['background'])
        if self._label:
            self._label.configure(bg=colors['background'], fg=colors['text_primary'])

    def get(self) -> bool:
        """Get current value"""
        return self._variable.get()

    def set(self, value: bool):
        """Set value"""
        self._variable.set(value)

    def set_enabled(self, enabled: bool):
        """Enable or disable the toggle"""
        self._enabled = enabled
        cursor = 'hand2' if enabled else 'arrow'
        self._icon_label.configure(cursor=cursor)
        if self._label:
            self._label.configure(cursor=cursor)


class LabeledRadioGroup:
    """
    Radio button group using SVG icons with optional descriptions.
    Uses radio-on.svg and radio-off.svg from assets/Tool Icons.

    Supports:
    - Horizontal or vertical orientation
    - Optional descriptions for each option
    - Disabling individual options with tooltips
    - Font and padding customization
    - Theme changes

    Usage:
        radio_group = LabeledRadioGroup(
            parent_frame,
            variable=my_string_var,
            options=[
                ("value1", "Option 1"),
                ("value2", "Option 2", "Optional description"),
            ],
            orientation="vertical",
            command=on_selection_changed
        )
        radio_group.pack()
    """

    # Class-level icon cache to avoid reloading
    _icon_cache = {}

    def __init__(
        self,
        parent: tk.Widget,
        variable: tk.StringVar = None,
        options: list = None,  # [(value, label) or (value, label, description), ...]
        command: Callable = None,
        orientation: str = "vertical",  # "horizontal" or "vertical"
        font: tuple = ('Segoe UI', 9),
        icon_size: int = 16,
        padding: int = 12,
        bg: str = None,  # Background color, defaults to theme's 'background'
        **kwargs
    ):
        self._theme_manager = get_theme_manager()
        colors = self._theme_manager.colors

        self._variable = variable or tk.StringVar()
        self._command = command
        self._options = options or []
        self._orientation = orientation
        self._font = font
        self._icon_size = icon_size
        self._padding = padding
        self._bg = bg  # Store custom bg, will use theme color if None
        self._radio_items = []
        self._disabled_options = {}
        self._option_tooltips = {}

        # Create the container frame - use background as default to match panel inner areas
        bg_color = bg if bg else colors['background']
        self.frame = tk.Frame(parent, bg=bg_color, **kwargs)

        # Load icons
        self._load_icons()

        # Create radio items
        for opt in self._options:
            if len(opt) == 2:
                value, label = opt
                description = None
            else:
                value, label, description = opt
            item = self._create_radio_item(value, label, description)
            self._radio_items.append(item)

        # Trace variable for updates
        self._trace_id = self._variable.trace_add("write", self._on_variable_changed)

        # Register for theme changes
        self._theme_manager.register_theme_callback(self._handle_theme_change)

    def _load_icons(self):
        """Load radio-on and radio-off SVG icons"""
        size = self._icon_size
        cache_key = f"radio_{size}x{size}"

        if cache_key in LabeledRadioGroup._icon_cache:
            self._icon_on, self._icon_off = LabeledRadioGroup._icon_cache[cache_key]
            return

        # Find assets directory - 3 levels up from core/ui/
        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"

        self._icon_on = self._load_svg_icon(icons_dir / "radio-on.svg", size)
        self._icon_off = self._load_svg_icon(icons_dir / "radio-off.svg", size)

        LabeledRadioGroup._icon_cache[cache_key] = (self._icon_on, self._icon_off)

    def _load_svg_icon(self, path: Path, size: int):
        """Load and resize an SVG icon"""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return None
        try:
            png_data = cairosvg.svg2png(url=str(path), output_width=size*2, output_height=size*2)
            img = Image.open(io.BytesIO(png_data))
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            img = img.resize((size, size), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def _create_radio_item(self, value: str, label: str, description: str = None):
        """Create a single radio item"""
        colors = self._theme_manager.colors
        bg_color = self._bg if self._bg else colors['background']

        row_frame = tk.Frame(self.frame, bg=bg_color)
        if self._orientation == "horizontal":
            row_frame.pack(side=tk.LEFT, padx=(0, self._padding))
        else:
            row_frame.pack(fill=tk.X, pady=2, anchor=tk.W)

        # Radio icon
        is_selected = self._variable.get() == value
        icon = self._icon_on if is_selected else self._icon_off

        icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
        if icon:
            icon_label.configure(image=icon)
            icon_label._icon_ref = icon
        icon_label.pack(side=tk.LEFT, padx=(0, 6))

        # Labels container
        labels_frame = tk.Frame(row_frame, bg=bg_color)
        labels_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Main label - selected uses title_color
        text_fg = colors['title_color'] if is_selected else colors['text_primary']
        main_label = tk.Label(
            labels_frame, text=label, bg=bg_color, fg=text_fg,
            font=self._font, anchor='w', cursor='hand2'
        )
        main_label.pack(side=tk.LEFT)

        # Optional description
        desc_label = None
        if description:
            desc_label = tk.Label(
                labels_frame, text=f"  -  {description}",
                bg=bg_color, fg=colors['text_secondary'],
                font=(self._font[0], self._font[1] - 1, 'italic'),
                anchor='w', cursor='hand2'
            )
            desc_label.pack(side=tk.LEFT)

        item = {
            'value': value,
            'frame': row_frame,
            'icon_label': icon_label,
            'labels_frame': labels_frame,
            'main_label': main_label,
            'desc_label': desc_label
        }

        # Click handler
        def on_click(event=None, v=value):
            if self._disabled_options.get(v, False):
                return
            self._variable.set(v)
            if self._command:
                self._command()

        icon_label.bind('<Button-1>', on_click)
        main_label.bind('<Button-1>', on_click)
        if desc_label:
            desc_label.bind('<Button-1>', on_click)
        row_frame.bind('<Button-1>', on_click)

        # Hover underline effect
        def on_enter(event=None, v=value):
            if self._disabled_options.get(v, False):
                return
            main_label.configure(font=self._font + ('underline',))

        def on_leave(event=None):
            main_label.configure(font=self._font)

        main_label.bind('<Enter>', on_enter)
        main_label.bind('<Leave>', on_leave)

        return item

    def _on_variable_changed(self, *args):
        """Update visual state when variable changes"""
        self._update_all()

    def _update_all(self):
        """Update all radio items"""
        colors = self._theme_manager.colors
        current_value = self._variable.get()

        for item in self._radio_items:
            is_selected = item['value'] == current_value
            is_disabled = self._disabled_options.get(item['value'], False)

            icon = self._icon_on if is_selected else self._icon_off
            if icon:
                item['icon_label'].configure(image=icon)
                item['icon_label']._icon_ref = icon

            # Determine text color
            if is_disabled:
                text_fg = colors['text_muted']
            elif is_selected:
                text_fg = colors['title_color']
            else:
                text_fg = colors['text_primary']

            item['main_label'].configure(fg=text_fg)

    def _handle_theme_change(self, theme: str):
        """Handle theme change"""
        colors = self._theme_manager.colors
        # If custom bg was set, don't update it (caller should handle)
        # Otherwise use background to match panel inner areas
        bg_color = self._bg if self._bg else colors['background']

        self.frame.configure(bg=bg_color)

        for item in self._radio_items:
            item['frame'].configure(bg=bg_color)
            item['icon_label'].configure(bg=bg_color)
            item['labels_frame'].configure(bg=bg_color)

            is_disabled = self._disabled_options.get(item['value'], False)
            is_selected = item['value'] == self._variable.get()

            if is_disabled:
                text_fg = colors['text_muted']
            elif is_selected:
                text_fg = colors['title_color']
            else:
                text_fg = colors['text_primary']

            item['main_label'].configure(bg=bg_color, fg=text_fg)
            if item['desc_label']:
                item['desc_label'].configure(bg=bg_color, fg=colors['text_secondary'])

    def on_theme_changed(self):
        """Public method for theme updates (alternative to callback registration)"""
        self._handle_theme_change(self._theme_manager.current_theme)

    def set_option_enabled(self, value: str, enabled: bool, tooltip: str = None):
        """
        Enable or disable a specific option by its value.

        Args:
            value: The option value to enable/disable
            enabled: True to enable, False to disable
            tooltip: Optional tooltip text to show when disabled
        """
        Tooltip = _get_tooltip()  # Deferred import
        colors = self._theme_manager.colors
        is_disabled = not enabled

        self._disabled_options[value] = is_disabled
        if tooltip:
            self._option_tooltips[value] = tooltip
        elif value in self._option_tooltips:
            del self._option_tooltips[value]

        # Find and update the row
        for item in self._radio_items:
            if item['value'] == value:
                if is_disabled:
                    item['main_label'].configure(fg=colors['text_muted'], cursor='arrow')
                    item['icon_label'].configure(cursor='arrow')
                    if item['desc_label']:
                        item['desc_label'].configure(cursor='arrow')

                    # Set up tooltip if provided
                    if tooltip:
                        item['_icon_tooltip'] = Tooltip(item['icon_label'], tooltip, delay=100)
                        item['_main_tooltip'] = Tooltip(item['main_label'], tooltip, delay=100)
                        item['_tooltip_bound'] = True
                else:
                    is_selected = self._variable.get() == value
                    text_fg = colors['title_color'] if is_selected else colors['text_primary']
                    item['main_label'].configure(fg=text_fg, cursor='hand2')
                    item['icon_label'].configure(cursor='hand2')
                    if item['desc_label']:
                        item['desc_label'].configure(cursor='hand2')

                    # Clear tooltips
                    if item.get('_tooltip_bound'):
                        if item.get('_icon_tooltip'):
                            item['_icon_tooltip'] = None
                        if item.get('_main_tooltip'):
                            item['_main_tooltip'] = None
                        item['_tooltip_bound'] = False
                break

    def get(self) -> str:
        """Get current value"""
        return self._variable.get()

    def set(self, value: str):
        """Set value"""
        self._variable.set(value)

    def set_enabled(self, enabled: bool):
        """Enable or disable the entire radio group"""
        colors = self._theme_manager.colors
        cursor = 'hand2' if enabled else 'arrow'

        for item in self._radio_items:
            value = item['value']
            self._disabled_options[value] = not enabled

            if enabled:
                is_selected = self._variable.get() == value
                text_fg = colors['title_color'] if is_selected else colors['text_primary']
            else:
                text_fg = colors['text_muted']

            item['main_label'].configure(fg=text_fg, cursor=cursor)
            item['icon_label'].configure(cursor=cursor)
            if item['desc_label']:
                item['desc_label'].configure(fg=colors['text_muted'] if not enabled else colors['text_secondary'], cursor=cursor)

    def set_bg(self, bg: str):
        """Update background color for all items (useful for theme changes with custom colors)"""
        self._bg = bg
        self.frame.configure(bg=bg)
        for item in self._radio_items:
            item['frame'].configure(bg=bg)
            item['icon_label'].configure(bg=bg)
            item['labels_frame'].configure(bg=bg)
            item['main_label'].configure(bg=bg)
            if item['desc_label']:
                item['desc_label'].configure(bg=bg)

    def pack(self, **kwargs):
        """Pack the container frame"""
        self.frame.pack(**kwargs)

    def grid(self, **kwargs):
        """Grid the container frame"""
        self.frame.grid(**kwargs)

    def place(self, **kwargs):
        """Place the container frame"""
        self.frame.place(**kwargs)
