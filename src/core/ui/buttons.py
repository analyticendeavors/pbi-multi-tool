"""
Button Components - Rounded and icon-based buttons
Built by Reid Havens of Analytic Endeavors

Provides reusable button components with theme support and hover/click states.
"""

import tkinter as tk
from typing import Callable

from core.theme_manager import get_theme_manager


class RoundedButton(tk.Canvas):
    """
    A button with rounded corners using Canvas.
    Provides full control over hover/click states with visual polish.

    Args:
        padding_x: Horizontal padding from button edge to content (default 16px).
                   If width is None, button auto-sizes to content + 2*padding_x.
                   Standard padding ensures consistent visual appearance across buttons.
    """

    # Standard horizontal padding for consistent button appearance
    DEFAULT_PADDING_X = 16

    def __init__(self, parent, text: str, command: Callable,
                 bg: str, fg: str, hover_bg: str, pressed_bg: str,
                 width: int = None, height: int = 32, radius: int = 6,
                 font: tuple = ('Segoe UI', 10),
                 icon: 'ImageTk.PhotoImage' = None,
                 disabled_bg: str = None, disabled_fg: str = None,
                 canvas_bg: str = None, padding_x: int = None,
                 corners: str = 'all', **kwargs):
        # Remove unsupported kwargs for Canvas
        kwargs.pop('padx', None)
        kwargs.pop('pady', None)

        # Use explicit canvas_bg if provided, otherwise will get from theme in _draw_button
        self._explicit_canvas_bg = canvas_bg
        self._parent = parent

        # Store padding - use default if not specified
        self._padding_x = padding_x if padding_x is not None else self.DEFAULT_PADDING_X

        # Calculate width if not provided (auto-size based on content + padding)
        self.btn_font = font
        self._icon = icon  # Store icon reference to prevent garbage collection

        if width is None:
            width = self._calculate_content_width(text, icon) + (2 * self._padding_x)

        # Initial bg will be set properly in _draw_button
        super().__init__(parent, width=width, height=height,
                        bg='#ffffff', highlightthickness=0, **kwargs)

        self.command = command
        self.text = text
        self.bg_normal = bg
        self.bg_hover = hover_bg
        self.bg_pressed = pressed_bg
        self.fg = fg
        self.bg_disabled = disabled_bg or '#3a3a4e'  # Default disabled bg
        self.fg_disabled = disabled_fg or '#6a6a7a'  # Default disabled text
        self.radius = radius
        self.corners = corners  # Which corners to round: 'all', 'left', 'right', 'none'
        self._current_bg = bg
        self._enabled = True  # State tracking
        self._last_width = None  # Track width to avoid unnecessary redraws on Configure

        # Draw initial state
        self._draw_button()

        # Bind events
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<ButtonPress-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)
        self.bind('<Configure>', self._on_configure)  # Redraw when resized (e.g., fill=tk.X)
        self.config(cursor='hand2')

    def _on_configure(self, event):
        """Handle resize events - redraw button to fill new dimensions."""
        # Only redraw if width actually changed (avoid redundant redraws)
        if event.width != self._last_width:
            self._last_width = event.width
            self._draw_button()

    def _calculate_content_width(self, text: str, icon) -> int:
        """Calculate the width of button content (icon + text)."""
        import tkinter.font as tkfont
        font_obj = tkfont.Font(font=self.btn_font)
        text_width = font_obj.measure(text)

        if icon:
            icon_width = icon.width()
            icon_spacing = 6  # Space between icon and text
            return icon_width + icon_spacing + text_width
        return text_width

    def _draw_rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        """Draw a rounded rectangle using ovals for corners and rectangles for body.

        Supports corner-specific rounding via self.corners:
        - 'all': round all corners (default)
        - 'left': round only left corners (top-left, bottom-left)
        - 'right': round only right corners (top-right, bottom-right)
        - 'none': no rounding
        """
        fill = kwargs.get('fill', '')
        outline = kwargs.get('outline', fill)
        corners = getattr(self, 'corners', 'all')

        r = min(radius, (x2-x1)//2, (y2-y1)//2)  # Ensure radius fits
        if r < 1 or corners == 'none':
            # No rounding, just a rectangle
            return self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, width=0)

        d = r * 2

        # Determine which corners to round
        # 'all': all corners, 'left': left side, 'right': right side
        # 'top-left': top-left only (for left tab), 'top-right': top-right only (for right tab)
        # 'top': top corners only
        round_tl = corners in ('all', 'left', 'top-left', 'top')
        round_tr = corners in ('all', 'right', 'top-right', 'top')
        round_bl = corners in ('all', 'left')
        round_br = corners in ('all', 'right')

        # Draw body rectangles FIRST (corners will overdraw edges)
        self.create_rectangle(x1+r, y1, x2-r, y2, fill=fill, outline=fill, width=0)
        self.create_rectangle(x1, y1+r, x2, y2-r, fill=fill, outline=fill, width=0)

        # Draw corner circles (ovals) on top for clean rounded edges
        if round_tl:
            self.create_oval(x1, y1, x1+d, y1+d, fill=fill, outline=fill, width=0)
        else:
            self.create_rectangle(x1, y1, x1+r, y1+r, fill=fill, outline=fill, width=0)

        if round_tr:
            self.create_oval(x2-d, y1, x2, y1+d, fill=fill, outline=fill, width=0)
        else:
            self.create_rectangle(x2-r, y1, x2, y1+r, fill=fill, outline=fill, width=0)

        if round_bl:
            self.create_oval(x1, y2-d, x1+d, y2, fill=fill, outline=fill, width=0)
        else:
            self.create_rectangle(x1, y2-r, x1+r, y2, fill=fill, outline=fill, width=0)

        if round_br:
            self.create_oval(x2-d, y2-d, x2, y2, fill=fill, outline=fill, width=0)
        else:
            self.create_rectangle(x2-r, y2-r, x2, y2, fill=fill, outline=fill, width=0)

    def _on_enter(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_hover
        self._draw_button()

    def _on_leave(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_normal
        self._draw_button()

    def _on_press(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_pressed
        self._draw_button()

    def _on_release(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_hover
        self._draw_button()
        if self.command:
            self.command()

    def set_enabled(self, enabled: bool):
        """Enable or disable the button"""
        self._enabled = enabled
        if enabled:
            self._current_bg = self.bg_normal
            self.config(cursor='hand2')
        else:
            # Use disabled background color for visual distinction
            self._current_bg = self.bg_disabled
            self.config(cursor='arrow')
        self._draw_button()

    def _draw_button(self):
        """Draw the button with current state"""
        # Determine canvas background color for rounded corners
        if self._explicit_canvas_bg:
            # Use explicitly provided canvas_bg
            theme_bg = self._explicit_canvas_bg
        else:
            # Try to get from theme - use 'background' to match ttk.Frame
            try:
                theme_bg = get_theme_manager().colors.get('background', '#ffffff')
            except Exception:
                theme_bg = '#ffffff'
        self.config(bg=theme_bg)

        self.delete('all')
        # Use actual dimensions (winfo_*) to handle fill=tk.X expansion
        # Fall back to configured dimensions if not yet mapped
        w = self.winfo_width() if self.winfo_width() > 1 else int(self.cget('width'))
        h = self.winfo_height() if self.winfo_height() > 1 else int(self.cget('height'))

        # Fill entire canvas with button color (no border visible)
        # When disabled, use disabled colors for visual distinction
        if self._enabled:
            draw_bg = self._current_bg
            draw_fg = self.fg
        else:
            # Enforce disabled colors - don't rely on _current_bg which can get out of sync
            draw_bg = self.bg_disabled
            draw_fg = self.fg_disabled

        self._draw_rounded_rect(0, 0, w, h, self.radius,
                               fill=draw_bg, outline=draw_bg)

        # Draw icon and text (icon to left of text if provided)
        if self._icon:
            # Calculate icon + text layout to center the combination
            icon_width = self._icon.width()
            icon_spacing = 6  # Space between icon and text

            # Get actual text width using font metrics
            import tkinter.font as tkfont
            font_obj = tkfont.Font(font=self.btn_font)
            text_width = font_obj.measure(self.text)

            # Total content width
            total_width = icon_width + icon_spacing + text_width

            # Start position to center the content
            start_x = (w - total_width) / 2

            # Icon position (centered vertically, at start of content)
            icon_x = start_x + icon_width / 2

            # Text position (after icon + spacing)
            text_x = start_x + icon_width + icon_spacing

            # Draw icon
            self.create_image(icon_x, h/2, image=self._icon, anchor='center')
            # Draw text to the right of icon
            self.create_text(text_x, h/2, text=self.text, fill=draw_fg,
                            font=self.btn_font, anchor='w')
        else:
            # No icon - draw text centered (original behavior)
            self.create_text(w/2, h/2, text=self.text, fill=draw_fg,
                            font=self.btn_font, anchor='center')

    def update_text(self, text: str):
        """Update the button text and redraw"""
        self.text = text
        self._draw_button()

    def update_colors(self, bg: str, hover_bg: str, pressed_bg: str, fg: str,
                      disabled_bg: str = None, disabled_fg: str = None,
                      canvas_bg: str = None):
        """Update button colors (for theme changes)"""
        self.bg_normal = bg
        self.bg_hover = hover_bg
        self.bg_pressed = pressed_bg
        self.fg = fg
        if disabled_bg:
            self.bg_disabled = disabled_bg
        if disabled_fg:
            self.fg_disabled = disabled_fg
        # Update canvas background if provided - also update _explicit_canvas_bg for persistence
        if canvas_bg:
            self._explicit_canvas_bg = canvas_bg
        # Respect current enabled state when updating colors
        self._current_bg = bg if self._enabled else self.bg_disabled
        self._draw_button()

    def update_canvas_bg(self, canvas_bg: str):
        """Update just the canvas background color (for theme changes).

        This is the background color that shows in the button corners
        and must match the parent container's background for seamless integration.
        """
        if canvas_bg:
            self._explicit_canvas_bg = canvas_bg
            self._draw_button()


class SquareIconButton(tk.Canvas):
    """
    A small square button with an icon and tooltip.
    Uses 6px rounded corners and theme-aware colors.
    """

    def __init__(self, parent, icon: 'ImageTk.PhotoImage', command: Callable,
                 tooltip_text: str = None,
                 size: int = 28, radius: int = 6,
                 bg_normal_override: dict = None, **kwargs):
        """
        Args:
            bg_normal_override: Optional dict with 'dark' and 'light' keys for custom
                               default background colors, e.g. {'dark': '#2f2f54', 'light': '#b6b6bd'}
        """
        # Get theme manager
        self._theme_manager = get_theme_manager()
        colors = self._theme_manager.colors
        self._parent = parent  # Store parent reference for background lookup
        self._bg_normal_override = bg_normal_override  # Store custom bg colors

        # Get actual parent background color - this is critical for hiding square corners
        self._section_bg = self._get_parent_background(parent, colors)

        super().__init__(parent, width=size, height=size,
                        bg=self._section_bg, highlightthickness=0, **kwargs)

        self.command = command
        self._icon = icon
        self._size = size
        self.radius = radius
        self._enabled = True
        self._tooltip = None

        # Theme-aware colors (subtle button style)
        self._update_colors_from_theme()
        # Update canvas bg to match (may differ from initial _get_parent_background)
        self.config(bg=self._section_bg)

        # Draw initial state
        self._current_bg = self.bg_normal
        self._draw_button()

        # Bind events
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<ButtonPress-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)
        self.config(cursor='hand2')

        # Create tooltip if text provided
        if tooltip_text:
            # Import Tooltip here to avoid circular import
            from core.ui.dialogs import Tooltip
            self._tooltip = Tooltip(self, tooltip_text)

        # Register for theme changes
        self._theme_manager.register_theme_callback(self._handle_theme_change)

    def _get_parent_background(self, parent, colors):
        """Get background color for canvas to match parent.

        NOTE: During theme changes, parent widgets may not be updated yet due to
        callback ordering. We use theme colors directly which are always correct,
        rather than reading from parent widget which may have stale values.
        """
        # First, try to get parent's bg directly (works for tk widgets)
        try:
            parent_bg = parent.cget('bg')
            # If parent bg matches a known theme color, it's been updated - use it
            if parent_bg in (colors['background'], colors['section_bg'], colors.get('card_surface', '')):
                return parent_bg
        except Exception:
            pass

        # For ttk widgets, try to get background from style
        try:
            style = parent.cget('style')
            if style:
                # Check for Section.TFrame which uses section_bg
                if 'Section' in style:
                    return colors.get('section_bg', colors['background'])
        except Exception:
            pass

        # Try to determine from widget class
        try:
            widget_class = parent.winfo_class()
            # ttk.Frame inside Section uses section_bg
            if widget_class == 'TFrame':
                # Check if parent is a section frame by looking at style
                try:
                    style = parent.cget('style')
                    if style and 'Section' in style:
                        return colors.get('section_bg', colors['background'])
                except Exception:
                    pass
        except Exception:
            pass

        # Default to background
        return colors.get('background', '#0d0d1a')

    def _update_colors_from_theme(self):
        """Update colors from current theme"""
        colors = self._theme_manager.colors

        # Determine if dark mode
        is_dark = colors.get('background', '') == '#0d0d1a'

        # Canvas bg must match parent to hide square corners
        self._section_bg = self._get_parent_background(self._parent, colors)

        # Button fill - use override if provided, otherwise card_surface
        if self._bg_normal_override:
            # Support both color keys (strings like 'background') and direct hex values
            override_value = self._bg_normal_override.get('dark' if is_dark else 'light')
            # If the override value is a color key, look it up; otherwise use it directly
            if override_value and override_value in colors:
                self.bg_normal = colors[override_value]
            elif override_value:
                self.bg_normal = override_value
            else:
                self.bg_normal = colors.get('card_surface', '#f5f5f7')
        else:
            # Default: card_surface - different from section_bg to make button visible
            # Dark: #1a1a2e (card_surface) vs #161627 (section_bg) - subtle contrast
            # Light: #e8e8f0 (card_surface) vs #f5f5f7 (section_bg) - subtle contrast
            self.bg_normal = colors.get('card_surface', colors.get('surface', '#f5f5f7'))

        # Hover: dark teal (#00587C) in dark, primary teal (#009999) in light
        self.bg_hover = '#00587C' if is_dark else '#009999'

        # Pressed: darker versions
        self.bg_pressed = '#003050' if is_dark else '#005C5C'
        self.fg_normal = colors.get('text_primary', '#1e1e1e')

        # Disabled: muted gray that's distinct in both themes
        self.bg_disabled = '#3a3a4e' if is_dark else '#c0c0cc'

    def _handle_theme_change(self, theme: str):
        """Handle theme change - theme parameter is 'dark' or 'light'"""
        # Update all colors based on new theme (this updates self._section_bg)
        self._update_colors_from_theme()

        # Reset to appropriate state with new colors (respect disabled state)
        self._current_bg = self.bg_disabled if not self._enabled else self.bg_normal
        self._draw_button()

        # Update canvas background using stored _section_bg (guaranteed to match bg_normal)
        self.config(bg=self._section_bg)

    def _on_enter(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_hover
        self._draw_button()

    def _on_leave(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_normal
        self._draw_button()

    def _on_press(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_pressed
        self._draw_button()

    def _on_release(self, event):
        if not self._enabled:
            return
        self._current_bg = self.bg_hover
        self._draw_button()
        if self.command:
            self.command()

    def _draw_rounded_rect(self, x1, y1, x2, y2, radius, fill):
        """Draw a rounded rectangle"""
        r = min(radius, (x2-x1)//2, (y2-y1)//2)
        if r < 1:
            return self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=fill, width=0)

        d = r * 2
        # Draw corner circles
        self.create_oval(x1, y1, x1+d, y1+d, fill=fill, outline=fill, width=0)
        self.create_oval(x2-d, y1, x2, y1+d, fill=fill, outline=fill, width=0)
        self.create_oval(x1, y2-d, x1+d, y2, fill=fill, outline=fill, width=0)
        self.create_oval(x2-d, y2-d, x2, y2, fill=fill, outline=fill, width=0)
        # Draw body rectangles
        self.create_rectangle(x1+r, y1, x2-r, y2, fill=fill, outline=fill, width=0)
        self.create_rectangle(x1, y1+r, x2, y2-r, fill=fill, outline=fill, width=0)

    def _draw_button(self):
        """Draw the button with current state"""
        self.delete('all')
        s = self._size

        # Draw rounded background
        self._draw_rounded_rect(0, 0, s, s, self.radius, self._current_bg)

        # Draw icon centered
        if self._icon:
            self.create_image(s // 2, s // 2, image=self._icon, anchor='center')

    def update_icon(self, icon: 'ImageTk.PhotoImage'):
        """Update the button icon"""
        self._icon = icon
        self._draw_button()

    def set_enabled(self, enabled: bool):
        """Enable or disable the button"""
        self._enabled = enabled
        self.config(cursor='hand2' if enabled else 'arrow')
        self._current_bg = self.bg_normal if enabled else self.bg_disabled
        self._draw_button()


class RoundedNavButton(tk.Canvas):
    """
    A navigation button with rounded corners that adapts based on sidebar state.

    Expanded mode: Left corners rounded (6px), right side flush with edge
    Collapsed mode: All 4 corners rounded (6px)
    """

    def __init__(self, parent, text: str, command: Callable,
                 bg: str, fg: str, hover_bg: str, pressed_bg: str,
                 active_bg: str, active_hover_bg: str, active_pressed_bg: str,
                 icon: tk.PhotoImage = None, mode: str = 'expanded',
                 width: int = 200, height: int = 40, radius: int = 6,
                 font: tuple = ('Segoe UI', 10), **kwargs):
        # Get parent background color for canvas
        try:
            parent_bg = parent.cget('bg')
        except Exception:
            parent_bg = '#1e1e1e'

        super().__init__(parent, width=width, height=height,
                        bg=parent_bg, highlightthickness=0, **kwargs)

        self.command = command
        self._text = text
        self._icon = icon
        self._mode = mode  # 'expanded' or 'collapsed'
        self._is_active = False
        self._parent_bg = parent_bg
        self.radius = radius
        self.btn_font = font
        self._requested_width = width
        self._requested_height = height

        # Color states
        self.bg_normal = bg
        self.bg_hover = hover_bg
        self.bg_pressed = pressed_bg
        self.active_bg = active_bg
        self.active_hover_bg = active_hover_bg
        self.active_pressed_bg = active_pressed_bg
        self.fg = fg
        self.fg_active = '#ffffff'  # White text when active

        self._current_bg = bg
        self._hover_state = False
        self._drawn = False
        self._icon_only = False  # When True, draw only icon centered (for collapsed sidebar with edge-touching buttons)

        # Bind events
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<ButtonPress-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)
        self.bind('<Configure>', self._on_configure)
        self.bind('<Map>', self._on_map)
        self.config(cursor='hand2')

    def _on_configure(self, event):
        """Handle resize - redraw button to fill available space"""
        # Only redraw if size actually changed
        new_w = event.width if event else self.winfo_width()
        new_h = event.height if event else self.winfo_height()
        if new_w > 1 and new_h > 1:
            self._draw_button()

    def _on_map(self, event):
        """Handle widget becoming visible - ensure initial draw"""
        # Multiple delayed draws to ensure rendering after geometry settles
        self.after(5, self._draw_button)
        self.after(50, self._draw_button)
        self.after(150, self._draw_button)

    def _draw_left_rounded_rect(self, x1, y1, x2, y2, radius, fill):
        """Draw rectangle with only LEFT corners rounded (for expanded mode)"""
        r = min(radius, (x2-x1)//2, (y2-y1)//2)
        if r < 1:
            self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=fill, width=0)
            return

        d = r * 2

        # Left corners - rounded (ovals)
        self.create_oval(x1, y1, x1+d, y1+d, fill=fill, outline=fill, width=0)  # Top-left
        self.create_oval(x1, y2-d, x1+d, y2, fill=fill, outline=fill, width=0)  # Bottom-left

        # Body rectangles - right side goes all the way to edge (no rounding)
        self.create_rectangle(x1+r, y1, x2, y2, fill=fill, outline=fill, width=0)  # Main body
        self.create_rectangle(x1, y1+r, x1+r, y2-r, fill=fill, outline=fill, width=0)  # Left edge fill

    def _draw_full_rounded_rect(self, x1, y1, x2, y2, radius, fill):
        """Draw rectangle with ALL corners rounded (for collapsed mode)"""
        r = min(radius, (x2-x1)//2, (y2-y1)//2)
        if r < 1:
            self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=fill, width=0)
            return

        d = r * 2

        # All 4 corners - rounded (ovals)
        self.create_oval(x1, y1, x1+d, y1+d, fill=fill, outline=fill, width=0)  # Top-left
        self.create_oval(x2-d, y1, x2, y1+d, fill=fill, outline=fill, width=0)  # Top-right
        self.create_oval(x1, y2-d, x1+d, y2, fill=fill, outline=fill, width=0)  # Bottom-left
        self.create_oval(x2-d, y2-d, x2, y2, fill=fill, outline=fill, width=0)  # Bottom-right

        # Body rectangles
        self.create_rectangle(x1+r, y1, x2-r, y2, fill=fill, outline=fill, width=0)
        self.create_rectangle(x1, y1+r, x2, y2-r, fill=fill, outline=fill, width=0)

    def _draw_button(self):
        """Draw the button with current state"""
        self.delete('all')

        # Use actual widget size - this handles fill=tk.X properly
        w = self.winfo_width()
        h = self.winfo_height()

        # Fallback to requested/configured size if widget not yet realized
        if w <= 1:
            w = self._requested_width if hasattr(self, '_requested_width') else int(self.cget('width'))
        if h <= 1:
            h = self._requested_height if hasattr(self, '_requested_height') else int(self.cget('height'))

        # Ensure minimum valid size
        if w < 10 or h < 10:
            return  # Don't draw if too small

        # Determine colors based on active state
        if self._is_active:
            draw_bg = self._current_bg
            draw_fg = self.fg_active
        else:
            draw_bg = self._current_bg
            draw_fg = self.fg

        # Draw background based on mode
        if self._mode == 'expanded':
            # Left corners rounded, right side flush
            self._draw_left_rounded_rect(0, 0, w, h, self.radius, draw_bg)
        else:
            # All corners rounded (collapsed mode)
            self._draw_full_rounded_rect(0, 0, w, h, self.radius, draw_bg)

        # Draw content (icon and/or text)
        # _icon_only overrides mode for content (but not shape)
        if self._icon_only or self._mode == 'collapsed':
            # Icon only, shifted left 3px to align with AE logo when collapsed
            icon_x = w/2 - 3
            if self._icon:
                self.create_image(icon_x, h/2, image=self._icon, anchor='center')
            else:
                # Fallback to first character/emoji
                display_text = self._text.strip().split()[0] if self._text else "•"
                self.create_text(icon_x, h/2, text=display_text, fill=draw_fg,
                                font=self.btn_font, anchor='center')
        else:
            # Expanded: Icon on left, text to the right
            text_x = 50 if self._icon else 15
            if self._icon:
                # Icon at x=15, centered vertically
                self.create_image(25, h/2, image=self._icon, anchor='center')
                # Text to the right of icon
                self.create_text(text_x, h/2, text=self._text, fill=draw_fg,
                                font=self.btn_font, anchor='w')
            else:
                # Just text, left-aligned with padding
                self.create_text(text_x, h/2, text=self._text, fill=draw_fg,
                                font=self.btn_font, anchor='w')


    def _on_enter(self, event):
        self._hover_state = True
        if self._is_active:
            self._current_bg = self.active_hover_bg
        else:
            self._current_bg = self.bg_hover
        self._draw_button()

    def _on_leave(self, event):
        self._hover_state = False
        if self._is_active:
            self._current_bg = self.active_bg
        else:
            self._current_bg = self.bg_normal
        self._draw_button()

    def _on_press(self, event):
        if self._is_active:
            self._current_bg = self.active_pressed_bg
        else:
            self._current_bg = self.bg_pressed
        self._draw_button()

    def _on_release(self, event):
        if self._is_active:
            self._current_bg = self.active_hover_bg
        else:
            self._current_bg = self.bg_hover
        self._draw_button()
        if self.command:
            self.command()

    def set_mode(self, mode: str):
        """Set the display mode ('expanded' or 'collapsed')"""
        self._mode = mode
        # Schedule redraw after geometry update
        self.after(1, self._draw_button)

    def set_icon_only(self, icon_only: bool):
        """Set whether to show only icon (centered) regardless of mode"""
        self._icon_only = icon_only

    def set_size(self, width: int, height: int):
        """Explicitly set button size (redraw handled externally via batch)"""
        self._requested_width = width
        self._requested_height = height
        self.config(width=width, height=height)

    def set_active(self, is_active: bool):
        """Set whether this button is the active/selected one"""
        self._is_active = is_active
        if is_active:
            self._current_bg = self.active_hover_bg if self._hover_state else self.active_bg
        else:
            self._current_bg = self.bg_hover if self._hover_state else self.bg_normal
        self._draw_button()

    def update_content(self, text: str = None, icon: tk.PhotoImage = None):
        """Update button text and/or icon"""
        if text is not None:
            self._text = text
        if icon is not None:
            self._icon = icon
        self._draw_button()

    def update_colors(self, bg: str, hover_bg: str, pressed_bg: str, fg: str,
                      active_bg: str, active_hover_bg: str, active_pressed_bg: str,
                      parent_bg: str = None):
        """Update button colors (for theme changes)"""
        self.bg_normal = bg
        self.bg_hover = hover_bg
        self.bg_pressed = pressed_bg
        self.fg = fg
        self.active_bg = active_bg
        self.active_hover_bg = active_hover_bg
        self.active_pressed_bg = active_pressed_bg
        if parent_bg:
            self._parent_bg = parent_bg
            self.config(bg=parent_bg)
        # Update current state
        if self._is_active:
            self._current_bg = self.active_hover_bg if self._hover_state else self.active_bg
        else:
            self._current_bg = self.bg_hover if self._hover_state else self.bg_normal
        self._draw_button()
