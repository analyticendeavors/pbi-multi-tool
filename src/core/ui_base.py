"""
UI Base Classes - Common UI patterns and base components
Built by Reid Havens of Analytic Endeavors

Provides reusable UI components and patterns for tool tabs.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import json
import io
from pathlib import Path
from typing import Optional, Dict, Any, Callable, List
from abc import ABC, abstractmethod

from core.constants import AppConstants
from core.theme_manager import get_theme_manager

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

# Alias for theme colors from constants
THEME_COLORS = AppConstants.THEMES

# Import extracted UI components (for backwards compatibility)
from core.ui.buttons import RoundedButton, SquareIconButton, RoundedNavButton
from core.ui.dialogs import Tooltip, ThemedMessageBox, ThemedInputDialog
from core.ui.template_widgets import ActionButtonBar, FileInputSection, SplitLogSection
from core.ui.menus import ThemedContextMenu
from core.ui.form_controls import SVGToggle, LabeledToggle, LabeledRadioGroup


class RecentFilesManager:
    """
    Manages recent files list with persistence.
    Stores recently used file paths per tool.
    """

    MAX_RECENT = 5
    _instance: Optional['RecentFilesManager'] = None

    def __init__(self):
        self._settings_path = Path.home() / "AppData" / "Local" / "AnalyticEndeavors" / "ae_multitool_settings.json"
        self._recent_files: Dict[str, List[str]] = {}
        self._load_recent_files()

    @classmethod
    def get_instance(cls) -> 'RecentFilesManager':
        """Get singleton instance"""
        if cls._instance is None:
            cls._instance = RecentFilesManager()
        return cls._instance

    def _load_recent_files(self):
        """Load recent files from settings"""
        try:
            if self._settings_path.exists():
                with open(self._settings_path, 'r') as f:
                    settings = json.load(f)
                    self._recent_files = settings.get('recent_files', {})
        except Exception:
            self._recent_files = {}

    def _save_recent_files(self):
        """Save recent files to settings"""
        try:
            self._settings_path.parent.mkdir(parents=True, exist_ok=True)

            # Load existing settings
            settings = {}
            if self._settings_path.exists():
                try:
                    with open(self._settings_path, 'r') as f:
                        settings = json.load(f)
                except Exception:
                    pass

            # Update recent files
            settings['recent_files'] = self._recent_files

            with open(self._settings_path, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass

    def add_file(self, tool_id: str, file_path: str):
        """Add a file to the recent list for a tool"""
        if tool_id not in self._recent_files:
            self._recent_files[tool_id] = []

        # Remove if already exists (to move to top)
        if file_path in self._recent_files[tool_id]:
            self._recent_files[tool_id].remove(file_path)

        # Add to beginning
        self._recent_files[tool_id].insert(0, file_path)

        # Trim to max size
        self._recent_files[tool_id] = self._recent_files[tool_id][:self.MAX_RECENT]

        self._save_recent_files()

    def get_recent(self, tool_id: str) -> List[str]:
        """Get recent files for a tool (only existing files)"""
        files = self._recent_files.get(tool_id, [])
        return [f for f in files if Path(f).exists()]

    def clear_recent(self, tool_id: str):
        """Clear recent files for a tool"""
        if tool_id in self._recent_files:
            self._recent_files[tool_id] = []
            self._save_recent_files()


def get_recent_files_manager() -> RecentFilesManager:
    """Get the global recent files manager instance"""
    return RecentFilesManager.get_instance()


class ThemedScrollbar(tk.Canvas):
    """
    A fully themeable scrollbar using Canvas for custom appearance.
    Supports dark/light mode with smooth colors that match the app theme.
    """

    def __init__(self, parent, command=None, theme_manager=None,
                 width: int = 12, auto_hide: bool = False, **kwargs):
        # Store theme manager reference
        self._theme_manager = theme_manager

        # Auto-hide settings
        self._auto_hide = auto_hide
        self._content_scrollable = False
        self._is_visible = not auto_hide  # Start hidden if auto_hide enabled
        self._parent = parent
        self._updating_visibility = False  # Guard against infinite loop

        # Get initial colors based on theme
        if theme_manager:
            colors = theme_manager.colors
            is_dark = theme_manager.is_dark
        else:
            colors = {}
            is_dark = True  # Default to dark

        # Scrollbar colors
        if is_dark:
            self._track_color = '#1a1a2e'  # Dark track
            self._thumb_color = '#3d3d5c'  # Medium gray thumb
            self._thumb_hover = '#4d4d6d'  # Lighter on hover
            self._thumb_active = '#5d5d7d'  # Even lighter when dragging
        else:
            self._track_color = '#f0f0f0'  # Light track
            self._thumb_color = '#c0c0c0'  # Gray thumb
            self._thumb_hover = '#a0a0a0'  # Darker on hover
            self._thumb_active = '#808080'  # Even darker when dragging

        super().__init__(parent, width=width, highlightthickness=0,
                        bg=self._track_color, **kwargs)

        self._command = command
        self._width = width
        self._thumb_id = None
        self._first = 0.0
        self._last = 1.0
        self._dragging = False
        self._drag_start_y = 0
        self._drag_start_first = 0.0

        # Bind events
        self.bind('<Configure>', self._on_configure)
        self.bind('<Button-1>', self._on_click)
        self.bind('<B1-Motion>', self._on_drag)
        self.bind('<ButtonRelease-1>', self._on_release)
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<MouseWheel>', self._on_mousewheel)

        # Register for theme changes
        if theme_manager:
            theme_manager.register_theme_callback(self._on_theme_change)

    def _on_theme_change(self, theme: str):
        """Update colors when theme changes"""
        is_dark = theme == 'dark'
        if is_dark:
            self._track_color = '#1a1a2e'
            self._thumb_color = '#3d3d5c'
            self._thumb_hover = '#4d4d6d'
            self._thumb_active = '#5d5d7d'
        else:
            self._track_color = '#f0f0f0'
            self._thumb_color = '#c0c0c0'
            self._thumb_hover = '#a0a0a0'
            self._thumb_active = '#808080'

        self.configure(bg=self._track_color)
        self._draw_thumb()

    def on_theme_changed(self):
        """Public method for theme updates (called manually when needed)"""
        if self._theme_manager:
            self._on_theme_change(self._theme_manager.current_theme)

    def set(self, first, last):
        """Set scrollbar position and handle auto-hide visibility"""
        self._first = float(first)
        self._last = float(last)

        # Check if content is scrollable (visible region < 100%)
        needs_scroll = (self._last - self._first) < 0.999

        # Handle auto-hide visibility changes (with guard to prevent infinite loop)
        if self._auto_hide and not self._updating_visibility:
            if needs_scroll and not self._is_visible:
                # Content needs scrolling, show scrollbar (deferred to break event chain)
                self._updating_visibility = True
                self.after(1, self._show_scrollbar)
            elif not needs_scroll and self._is_visible:
                # Content doesn't need scrolling, hide scrollbar (deferred)
                self._updating_visibility = True
                self.after(1, self._hide_scrollbar)

        # Only draw thumb if scrollbar is visible and content needs scrolling
        if self._is_visible and needs_scroll:
            self._draw_thumb()
        elif not needs_scroll:
            self.delete('thumb')  # Clear thumb when not needed

    def _show_scrollbar(self):
        """Show scrollbar (called deferred to prevent infinite loop)"""
        try:
            self._is_visible = True
            self._content_scrollable = True
            self.pack(side=tk.RIGHT, fill=tk.Y)
        finally:
            self._updating_visibility = False

    def _hide_scrollbar(self):
        """Hide scrollbar (called deferred to prevent infinite loop)"""
        try:
            self._is_visible = False
            self._content_scrollable = False
            self.pack_forget()
        finally:
            self._updating_visibility = False

    def _draw_thumb(self):
        """Draw the scrollbar thumb"""
        self.delete('thumb')

        height = self.winfo_height()
        if height <= 1:
            return

        # Calculate thumb size and position
        thumb_height = max(30, (self._last - self._first) * height)
        thumb_top = self._first * height
        thumb_bottom = thumb_top + thumb_height

        # Ensure thumb doesn't go past bounds
        if thumb_bottom > height:
            thumb_bottom = height
            thumb_top = thumb_bottom - thumb_height

        # Choose color based on state
        if self._dragging:
            color = self._thumb_active
        else:
            color = self._thumb_color

        # Draw rounded rectangle thumb with padding
        padding = 2
        radius = 4
        x1, y1 = padding, thumb_top + padding
        x2, y2 = self._width - padding, thumb_bottom - padding

        # Draw rounded rectangle
        self._thumb_id = self._create_rounded_rect(x1, y1, x2, y2, radius, color)

    def _create_rounded_rect(self, x1, y1, x2, y2, radius, color):
        """Create a rounded rectangle on the canvas"""
        # Ensure minimum size
        if y2 - y1 < radius * 2:
            radius = (y2 - y1) / 2

        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return self.create_polygon(points, fill=color, smooth=True, tags='thumb')

    def _on_configure(self, event):
        """Redraw when resized"""
        self._draw_thumb()

    def _on_click(self, event):
        """Handle click on scrollbar"""
        height = self.winfo_height()
        thumb_top = self._first * height
        thumb_height = max(30, (self._last - self._first) * height)
        thumb_bottom = thumb_top + thumb_height

        if thumb_top <= event.y <= thumb_bottom:
            # Clicked on thumb - start dragging
            self._dragging = True
            self._drag_start_y = event.y
            self._drag_start_first = self._first
            self._draw_thumb()
        else:
            # Clicked on track - page up/down
            if event.y < thumb_top:
                self._page_scroll(-1)
            else:
                self._page_scroll(1)

    def _on_drag(self, event):
        """Handle drag"""
        if self._dragging and self._command:
            height = self.winfo_height()
            delta_y = event.y - self._drag_start_y
            delta_scroll = delta_y / height

            new_first = self._drag_start_first + delta_scroll
            # Clamp to valid range
            max_first = 1.0 - (self._last - self._first)
            new_first = max(0.0, min(max_first, new_first))

            self._command('moveto', new_first)

    def _on_release(self, event):
        """Handle release"""
        self._dragging = False
        self._draw_thumb()

    def _on_enter(self, event):
        """Handle mouse enter - show hover state"""
        if not self._dragging:
            self.itemconfig('thumb', fill=self._thumb_hover)

    def _on_leave(self, event):
        """Handle mouse leave - restore normal state"""
        if not self._dragging:
            self.itemconfig('thumb', fill=self._thumb_color)

    def _on_mousewheel(self, event):
        """Handle mousewheel scroll"""
        if self._command:
            delta = -1 if event.delta > 0 else 1
            self._command('scroll', delta, 'units')

    def _page_scroll(self, direction):
        """Scroll by page"""
        if self._command:
            self._command('scroll', direction, 'pages')


class ModernScrolledText(tk.Frame):
    """
    A modern text widget with themed scrollbar that auto-hides when not needed.
    Uses ThemedScrollbar for full dark/light mode support.
    Provides the same interface as scrolledtext.ScrolledText.
    """

    def __init__(self, parent, height: int = 10, width: int = 40,
                 font: tuple = ('Segoe UI', 9), state: str = tk.NORMAL,
                 bg: str = '#ffffff', fg: str = '#1e1e1e',
                 selectbackground: str = '#0078d4', relief: str = 'flat',
                 borderwidth: int = 0, wrap: str = tk.WORD,
                 highlightthickness: int = 1, highlightcolor: str = '#e0e0e0',
                 highlightbackground: str = '#e0e0e0', auto_hide_scrollbar: bool = True,
                 theme_manager=None, **kwargs):
        # Store border settings for the outer frame (wraps both text and scrollbar)
        self._highlight_color = highlightcolor
        self._highlight_bg = highlightbackground
        self._highlight_thickness = highlightthickness

        # Get parent background for frame
        try:
            parent_bg = parent.cget('bg')
        except Exception:
            try:
                parent_bg = parent.cget('background')
            except Exception:
                parent_bg = bg

        # Put border on outer frame so it wraps BOTH text and scrollbar
        super().__init__(parent, bg=bg, highlightthickness=highlightthickness,
                        highlightcolor=highlightcolor, highlightbackground=highlightbackground,
                        borderwidth=0)

        self._auto_hide = auto_hide_scrollbar
        self._scrollbar_visible = True
        self._theme_manager = theme_manager

        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create text widget (no border - border is on outer frame)
        self._text = tk.Text(
            self, height=height, width=width, font=font, state=state,
            bg=bg, fg=fg, selectbackground=selectbackground, relief=relief,
            borderwidth=borderwidth, wrap=wrap, highlightthickness=0,
            **kwargs
        )
        self._text.grid(row=0, column=0, sticky='nsew')

        # Create vertical scrollbar (ThemedScrollbar for dark/light mode support)
        self._vscrollbar = ThemedScrollbar(self, command=self._text.yview,
                                           theme_manager=theme_manager)
        self._text.configure(yscrollcommand=self._on_scroll)

        # Create horizontal scrollbar if wrap is NONE (use ttk for horizontal)
        self._hscrollbar = None
        if wrap == tk.NONE:
            self._hscrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self._text.xview)
            self._text.configure(xscrollcommand=self._hscrollbar.set)
            self._hscrollbar.grid(row=1, column=0, sticky='ew')

        # Initially show scrollbar
        self._vscrollbar.grid(row=0, column=1, sticky='ns')

        # Bind events for auto-hide
        if self._auto_hide:
            self._text.bind('<Configure>', self._check_scrollbar_visibility)
            self._text.bind('<<Modified>>', self._on_text_modified)

        # Register theme callback to update border colors
        if self._theme_manager:
            self._theme_manager.register_theme_callback(self._on_theme_change)

    def _on_theme_change(self, theme: str):
        """Update colors when theme changes"""
        is_dark = theme == 'dark'
        colors = THEME_COLORS['dark'] if is_dark else THEME_COLORS['light']

        # Update text widget background and selection colors
        # Use section_bg (#161627 dark, #f5f5f7 light) for text areas
        # Use selection_bg (#1a5a8a dark, #3b82f6 light) for text selection - ALWAYS blue
        self._text.configure(
            bg=colors['section_bg'],
            fg=colors['text_primary'],
            selectbackground=colors['selection_bg'],
            selectforeground='#ffffff'
        )

        # Update outer frame border and background (border wraps text + scrollbar)
        # Must use tk.Frame.configure() directly since self.configure() is overridden
        # to delegate to the inner text widget
        tk.Frame.configure(self,
            bg=colors['section_bg'],
            highlightcolor=colors['border'],
            highlightbackground=colors['border']
        )

    def _on_scroll(self, first, last):
        """Handle scroll command and update scrollbar visibility"""
        self._vscrollbar.set(first, last)
        if self._auto_hide:
            self._update_scrollbar_visibility(float(first), float(last))

    def _update_scrollbar_visibility(self, first: float, last: float):
        """Show/hide scrollbar based on content"""
        # If first is 0 and last is 1, all content is visible - hide scrollbar
        should_show = not (first <= 0.0 and last >= 1.0)

        if should_show and not self._scrollbar_visible:
            self._vscrollbar.grid(row=0, column=1, sticky='ns')
            self._scrollbar_visible = True
        elif not should_show and self._scrollbar_visible:
            self._vscrollbar.grid_remove()
            self._scrollbar_visible = False

    def _check_scrollbar_visibility(self, event=None):
        """Check if scrollbar should be visible after resize"""
        self.after(10, self._delayed_visibility_check)

    def _delayed_visibility_check(self):
        """Delayed check for scrollbar visibility"""
        try:
            first, last = self._vscrollbar.get()
            self._update_scrollbar_visibility(first, last)
        except Exception:
            pass

    def _on_text_modified(self, event=None):
        """Handle text modification"""
        self.after(10, self._delayed_visibility_check)
        # Reset modified flag
        self._text.edit_modified(False)

    # Delegate common methods to the text widget
    def config(self, **kwargs):
        """Configure the text widget"""
        self._text.config(**kwargs)

    def configure(self, **kwargs):
        """Configure the text widget"""
        self._text.configure(**kwargs)

    def cget(self, key):
        """Get configuration value"""
        return self._text.cget(key)

    def insert(self, index, chars, *args):
        """Insert text"""
        self._text.insert(index, chars, *args)

    def delete(self, index1, index2=None):
        """Delete text"""
        self._text.delete(index1, index2)

    def get(self, index1, index2=None):
        """Get text"""
        return self._text.get(index1, index2)

    def see(self, index):
        """Scroll to make index visible"""
        self._text.see(index)

    def tag_config(self, tagName, **kwargs):
        """Configure a tag"""
        self._text.tag_config(tagName, **kwargs)

    def tag_configure(self, tagName, **kwargs):
        """Configure a tag"""
        self._text.tag_configure(tagName, **kwargs)

    def tag_add(self, tagName, index1, index2=None):
        """Add a tag"""
        self._text.tag_add(tagName, index1, index2)

    def tag_bind(self, tagName, sequence, func=None, add=None):
        """Bind an event to a tag"""
        return self._text.tag_bind(tagName, sequence, func, add)

    def tag_ranges(self, tagName):
        """Return list of ranges for a tag"""
        return self._text.tag_ranges(tagName)

    def index(self, index):
        """Return index"""
        return self._text.index(index)

    def yview(self, *args):
        """Vertical view"""
        return self._text.yview(*args)

    def xview(self, *args):
        """Horizontal view"""
        return self._text.xview(*args)

    def window_create(self, index, **kwargs):
        """Create an embedded window in the text widget"""
        return self._text.window_create(index, **kwargs)

    @property
    def text_widget(self):
        """Return the internal Text widget for creating embedded widgets"""
        return self._text


class BaseToolTab(ABC):
    """
    Base class for tool UI tabs.
    Provides common UI patterns and functionality.
    """
    
    def __init__(self, parent, main_app, tool_id: str, tool_name: str):
        self.parent = parent
        self.main_app = main_app
        self.tool_id = tool_id
        self.tool_name = tool_name

        # Create main frame for this tab
        self.frame = ttk.Frame(parent, padding="20")

        # Common UI state
        self.is_busy = False
        self.progress_bar = None
        self.progress_label = None  # Progress status label
        self.progress_frame = None  # Progress container
        self.progress_persist = False  # Whether to keep progress visible
        self.log_text = None

        # Track buttons for theme updates
        self._primary_buttons: List[RoundedButton] = []
        self._secondary_buttons: List[RoundedButton] = []

        # Track section header widgets for theme updates
        self._section_header_widgets: List[tuple] = []

        # Icon cache (populated by subclasses or create_section_header)
        if not hasattr(self, '_button_icons'):
            self._button_icons: Dict[str, Any] = {}

        # Theme support
        self._theme_manager = get_theme_manager()
        self._theme_manager.register_theme_callback(self._handle_theme_change)

        # Setup styling
        self._setup_common_styling()
    
    def get_frame(self) -> ttk.Frame:
        """Return the main frame for this tab"""
        return self.frame
    
    @abstractmethod
    def setup_ui(self) -> None:
        """Setup the UI for this tab - must be implemented by subclasses"""
        pass
    
    @abstractmethod
    def reset_tab(self) -> None:
        """Reset the tab to initial state - must be implemented by subclasses"""
        pass
    
    @abstractmethod
    def show_help_dialog(self) -> None:
        """Show help dialog for this tab - must be implemented by subclasses"""
        pass

    def on_tab_activated(self) -> None:
        """
        Called when this tab becomes active (user switches to it).
        Override in subclasses for tab-specific initialization.
        """
        pass

    def on_tab_deactivated(self) -> None:
        """
        Called when leaving this tab (user switches away).
        Override in subclasses for cleanup or state preservation.
        """
        pass

    def show_toast(self, message: str, toast_type: str = 'info', duration: int = 3000):
        """
        Show a non-blocking toast notification.

        Args:
            message: The message to display
            toast_type: One of 'success', 'error', 'warning', 'info'
            duration: How long to show the toast in milliseconds
        """
        try:
            from core.widgets import ToastNotification
            # Get the root window
            root = self.parent.winfo_toplevel()
            ToastNotification(root, message, toast_type, duration)
        except Exception:
            # Fallback to log if toast fails
            if self.log_text:
                self.log(message)

    def add_recent_file(self, file_path: str):
        """Add a file to the recent files list for this tool"""
        recent_manager = get_recent_files_manager()
        recent_manager.add_file(self.tool_id, file_path)

    def get_canvas_bg(self, context: str = 'section') -> str:
        """
        Get appropriate canvas background color for RoundedButton.

        RoundedButton requires a canvas_bg parameter to properly render
        rounded corners. This helper centralizes the logic for determining
        the correct background color based on where the button is placed.

        Args:
            context: Where the button is placed:
                - 'section' (default): Inside a section/card area
                - 'outer': On outer/main background area
                - 'dialog': Inside a popup/dialog window
                - 'input': Background for input field areas

        Returns:
            The appropriate hex color string for canvas background.
        """
        colors = self._theme_manager.colors
        if context == 'outer':
            return colors.get('outer_bg', colors['section_bg'])
        elif context == 'dialog':
            return colors.get('dialog_bg', colors['background'])
        elif context == 'input':
            return colors['background']
        # Default: section context
        return colors['section_bg']

    def show_error(self, title: str, message: str):
        """
        Show a themed error dialog.

        Use this instead of messagebox.showerror() for consistent theming.

        Args:
            title: Dialog title
            message: Error message to display
        """
        ThemedMessageBox.showerror(self.parent.winfo_toplevel(), title, message)

    def show_info(self, title: str, message: str):
        """
        Show a themed info dialog.

        Use this instead of messagebox.showinfo() for consistent theming.

        Args:
            title: Dialog title
            message: Info message to display
        """
        ThemedMessageBox.showinfo(self.parent.winfo_toplevel(), title, message)

    def show_warning(self, title: str, message: str):
        """
        Show a themed warning dialog.

        Use this instead of messagebox.showwarning() for consistent theming.

        Args:
            title: Dialog title
            message: Warning message to display
        """
        ThemedMessageBox.showwarning(self.parent.winfo_toplevel(), title, message)

    def show_success(self, title: str, message: str):
        """
        Show a themed success dialog.

        Args:
            title: Dialog title
            message: Success message to display
        """
        ThemedMessageBox.showsuccess(self.parent.winfo_toplevel(), title, message)

    def ask_yes_no(self, title: str, message: str) -> bool:
        """
        Show a themed yes/no confirmation dialog.

        Use this instead of messagebox.askyesno() for consistent theming.

        Args:
            title: Dialog title
            message: Question to ask

        Returns:
            True if user clicked Yes, False if No
        """
        return ThemedMessageBox.askyesno(self.parent.winfo_toplevel(), title, message) == "Yes"

    def create_secondary_button(self, parent, text: str, command: Callable, width: int = 16, height: int = 32,
                                  icon: 'ImageTk.PhotoImage' = None) -> RoundedButton:
        """
        Create a secondary button with rounded corners and proper hover/click states.
        Uses RoundedButton for full control - goes DARKER on hover, DARKEST on click.
        Used for Export Log, Clear Log, Reset All buttons.
        """
        import tkinter.font as tkfont

        colors = self._theme_manager.colors

        # Get the button-specific color values (different from card colors)
        bg_normal = colors['button_secondary']
        bg_hover = colors['button_secondary_hover']  # Darker
        bg_pressed = colors['button_secondary_pressed']  # Darkest
        fg = colors['text_primary']

        # Canvas background - secondary buttons sit inside frames
        # Light mode: #ffffff, Dark mode: #0d0d1a
        canvas_bg = colors['background']

        # Calculate width based on actual content for consistent padding
        btn_font = ('Segoe UI', 10)
        font_obj = tkfont.Font(font=btn_font)
        text_width = font_obj.measure(text)

        # Add icon width if provided
        icon_width = icon.width() if icon else 0
        icon_spacing = 6 if icon else 0  # Space between icon and text

        # Total content width + consistent padding (24px on each side = 48px total)
        content_width = text_width + icon_width + icon_spacing
        pixel_width = content_width + 48

        btn = RoundedButton(
            parent,
            text=text,
            command=command,
            bg=bg_normal,
            fg=fg,
            hover_bg=bg_hover,
            pressed_bg=bg_pressed,
            width=pixel_width,
            height=height,
            radius=6,
            font=btn_font,
            icon=icon,
            canvas_bg=canvas_bg
        )

        # Track for theme updates
        self._secondary_buttons.append(btn)

        return btn

    def create_primary_button(self, parent, text: str, command: Callable, width: int = 12,
                               icon: 'ImageTk.PhotoImage' = None) -> RoundedButton:
        """
        Create a primary button with rounded corners and proper hover/click states.
        Uses RoundedButton for full control - goes DARKER on hover, DARKEST on click.
        Used for action buttons like Browse.
        Dark mode: Blue (#00587C), Light mode: Teal (#009999)

        Args:
            icon: Optional icon to display to the left of text
        """
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'

        # Primary button colors - blue in dark mode, teal in light mode
        bg_normal = colors['button_primary']          # Blue or Teal based on theme
        bg_hover = colors['button_primary_hover']     # Darker
        bg_pressed = colors['button_primary_pressed'] # Darkest
        bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        fg = '#ffffff'  # White text
        fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Canvas background - primary buttons sit inside Section.TFrame (inside Section.TLabelframe)
        # Dynamically lookup the actual style background to ensure exact match
        try:
            style = ttk.Style()
            canvas_bg = style.lookup('Section.TFrame', 'background')
            if not canvas_bg:
                canvas_bg = colors['section_bg']
        except Exception:
            canvas_bg = colors['section_bg']

        # Convert character width to pixels (approx 8px per char + padding)
        pixel_width = width * 8 + 24

        btn = RoundedButton(
            parent,
            text=text,
            command=command,
            bg=bg_normal,
            fg=fg,
            hover_bg=bg_hover,
            pressed_bg=bg_pressed,
            width=pixel_width,
            height=32,
            radius=6,
            font=('Segoe UI', 10),
            icon=icon,
            disabled_bg=bg_disabled,
            disabled_fg=fg_disabled,
            canvas_bg=canvas_bg
        )

        # Track for theme updates
        self._primary_buttons.append(btn)

        return btn

    def create_action_button(self, parent, text: str, command: Callable, width: int = None,
                              icon: 'ImageTk.PhotoImage' = None) -> RoundedButton:
        """
        Create an action button (like ANALYZE REPORTS, EXECUTE MERGE) with rounded corners.
        Same style as primary buttons but larger - used for main action buttons.
        Dark mode: Blue (#00587C), Light mode: Teal (#009999)

        Uses RoundedButton auto-sizing with 24px horizontal padding for consistent appearance.

        Args:
            width: Optional explicit width. If None, auto-sizes based on content + padding.
            icon: Optional icon to display to the left of text
        """
        colors = self._theme_manager.colors
        is_dark = colors.get('background', '') == '#0d0d1a'

        # Same colors as primary buttons
        bg_normal = colors['button_primary']
        bg_hover = colors['button_primary_hover']
        bg_pressed = colors['button_primary_pressed']
        bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        fg = '#ffffff'  # White text
        fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Canvas background - action buttons typically sit on section frames
        # Use section_bg for proper corner blending
        canvas_bg = colors.get('section_bg', colors['background'])

        btn_font = ('Segoe UI', 10, 'bold')

        # Action buttons use default padding for consistent appearance across all buttons
        btn = RoundedButton(
            parent,
            text=text,
            command=command,
            bg=bg_normal,
            fg=fg,
            hover_bg=bg_hover,
            pressed_bg=bg_pressed,
            width=width,  # None = auto-size
            height=38,  # Taller than primary buttons
            radius=6,
            font=btn_font,
            icon=icon,
            disabled_bg=bg_disabled,
            disabled_fg=fg_disabled,
            canvas_bg=canvas_bg
            # Uses DEFAULT_PADDING_X (16px) for consistency with other buttons
        )

        # Track for theme updates (same as primary)
        self._primary_buttons.append(btn)

        return btn

    def create_section_header(
        self,
        parent: tk.Widget,
        text: str,
        icon_name: str,
        bg_color: str = None
    ) -> tuple:
        """
        Create a section header (icon + text) for use as LabelFrame labelwidget.

        This centralizes the pattern previously duplicated in _create_section_labelwidget()
        across all tool UI files. The header consists of an icon label and text label
        in a frame, styled to match the Section.TLabelframe.Label appearance.

        Args:
            parent: Parent widget (typically self.frame)
            text: Section title text (e.g., "PBIP File Source")
            icon_name: Name of icon in self._button_icons dict (e.g., "folder")
            bg_color: Background color override (defaults to colors['section_bg'])

        Returns:
            Tuple of (header_frame, icon_label, text_label) for theme updates.
            The header_frame should be used as the labelwidget parameter of ttk.LabelFrame.

        Example:
            header = self.create_section_header(self.frame, "PBIP File Source", "folder")
            section = ttk.LabelFrame(self.frame, labelwidget=header[0],
                                     style='Section.TLabelframe', padding="12")
        """
        colors = self._theme_manager.colors
        icon = self._button_icons.get(icon_name)
        if bg_color is None:
            bg_color = colors.get('section_bg', colors['background'])

        # Frame to hold icon + text
        header_frame = tk.Frame(parent, bg=bg_color)

        icon_label = None
        if icon:
            icon_label = tk.Label(header_frame, image=icon, bg=bg_color)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            # Store reference to prevent garbage collection
            icon_label._icon_ref = icon

        # Use title_color (blue in dark, teal in light) and Semibold font
        # to match Section.TLabelframe.Label style
        text_label = tk.Label(
            header_frame, text=text, bg=bg_color,
            fg=colors['title_color'], font=('Segoe UI Semibold', 11)
        )
        text_label.pack(side=tk.LEFT)

        # Store for theme updates
        self._section_header_widgets.append((header_frame, icon_label, text_label))

        return (header_frame, icon_label, text_label)

    def get_recent_files(self) -> List[str]:
        """Get recent files for this tool"""
        recent_manager = get_recent_files_manager()
        return recent_manager.get_recent(self.tool_id)

    def _setup_common_styling(self):
        """Setup common professional styling"""
        style = ttk.Style()
        # NOTE: Do NOT call style.theme_use('clam') here!
        # The theme_manager handles theme configuration globally.
        # Calling theme_use() resets ALL styles, wiping out TFrame background etc.
        colors = self._theme_manager.colors
        
        # Common styles
        # Note: In light mode, background=#ffffff (white), section_bg=#f5f5f7 (off-white)
        # User wants: outer background = off-white (#f5f5f7), inner content = white (#ffffff)
        # Title sits ABOVE the white frame, on the gray background
        styles = {
            'Section.TLabelframe': {
                'background': colors['section_bg'],  # Gray - title sits on this
                'borderwidth': 0,
                'relief': 'flat',
                'labelmargins': (0, 0, 0, 5)  # Small gap below title before content
            },
            'Section.TLabelframe.Label': {
                'background': colors['section_bg'],  # Gray background - matches outer area
                'foreground': colors['title_color'],  # Lighter blue in dark mode, teal in light mode
                'font': ('Segoe UI Semibold', 11),  # Semibold - modern, not too heavy
                'padding': (0, 2, 0, 4)  # Reduced padding - title closer to content
            },
            # Inner frames within sections - use background (white/dark) for outer containers
            'Section.TFrame': {
                'background': colors['background'],  # #0d0d1a dark, #ffffff light
            },
            # Labels inside sections - use background for outer container labels
            'Section.TLabel': {
                'background': colors['background'],  # #0d0d1a dark, #ffffff light
                'foreground': colors['text_primary'],
            },
            # Guide label styles - for Quick Start Guide with custom colors
            'GuideTitle.TLabel': {
                'background': colors['background'],
                'foreground': colors['info'],  # Blue title
            },
            'GuideWarning.TLabel': {
                'background': colors['background'],
                'foreground': colors['warning'],  # Orange warning
            },
            'GuideStep.TLabel': {
                'background': colors['background'],
                'foreground': colors['text_primary'],  # Normal text color
            },
            # Subsection titles - matches Section.TLabelframe.Label style for inner panel headers
            'Subsection.TLabel': {
                'background': colors['background'],  # White/dark to match content area
                'foreground': colors['title_color'],  # Same teal/blue as section headers
                'font': ('Segoe UI Semibold', 11),  # Same font as section headers
            },
            # Subsection warning variant - same font but warning color
            'SubsectionWarning.TLabel': {
                'background': colors['background'],
                'foreground': colors['warning'],  # Orange warning color
                'font': ('Segoe UI Semibold', 11),  # Same font as section headers
            },
            'Action.TButton': {
                'background': colors['primary'], 
                'foreground': colors['surface'], 
                'font': ('Segoe UI', 10, 'bold'), 
                'padding': (20, 10)
            },
            'Secondary.TButton': {
                'background': colors['border'], 
                'foreground': colors['text_primary'], 
                'font': ('Segoe UI', 10), 
                'padding': (15, 8)
            },
            'Brand.TButton': {
                'background': colors['accent'], 
                'foreground': colors['surface'], 
                'font': ('Segoe UI', 10, 'bold'), 
                'padding': (15, 8)
            },
            'Info.TButton': {
                'background': colors['info'], 
                'foreground': colors['surface'], 
                'font': ('Segoe UI', 9), 
                'padding': (12, 6)
            },
            'TProgressbar': {
                'background': colors['button_primary'],  # Blue in dark mode, teal in light mode
                'troughcolor': colors['border']
            },
            'TEntry': {
                'fieldbackground': colors['surface']
            },
            'TFrame': {
                'background': colors['section_bg']  # Off-white in light mode for outer areas
            },
            'TLabel': {
                'background': colors['section_bg']  # Off-white in light mode for outer areas
            },
            'TCheckbutton': {
                'background': colors['section_bg'],  # Match outer areas
                'foreground': colors['text_primary'],
                'focuscolor': 'none'  # Remove focus outline
            },
            'TRadiobutton': {
                'background': colors['background'],  # White in light mode for inner content
                'foreground': colors['text_primary'],
                'focuscolor': 'none'  # Remove focus outline
            }
        }
        
        for style_name, config in styles.items():
            style.configure(style_name, **config)

    def _handle_theme_change(self, theme: str):
        """Internal handler for theme changes - refreshes styling"""
        self._setup_common_styling()
        self._update_widget_colors()
        self.on_theme_changed(theme)
        # Force ttk widgets to pick up new style values
        self.frame.update_idletasks()

    def on_theme_changed(self, theme: str):
        """
        Hook for subclasses to respond to theme changes.
        Override this method to update custom widget colors.
        Subclasses should call super().on_theme_changed(theme) first.

        Args:
            theme: The new theme name ('dark' or 'light')
        """
        # Update section header widgets created via create_section_header()
        colors = self._theme_manager.colors
        bg_color = colors.get('section_bg', colors['background'])
        for header_frame, icon_label, text_label in self._section_header_widgets:
            try:
                header_frame.configure(bg=bg_color)
                if icon_label:
                    icon_label.configure(bg=bg_color)
                text_label.configure(bg=bg_color, fg=colors['title_color'])
            except Exception:
                pass  # Widget may have been destroyed

    def _update_widget_colors(self):
        """Update colors of non-ttk widgets (tk.Label, tk.Frame, etc.)"""
        colors = self._theme_manager.colors

        # Update log text widget if it exists (ModernScrolledText = tk.Frame with inner _text)
        # Use section_bg (#161627 dark, #f5f5f7 light) for consistent background
        # Must update BOTH the inner text widget AND the outer Frame wrapper
        if self.log_text:
            # Update inner text widget
            self.log_text.config(
                bg=colors['section_bg'],
                fg=colors['text_primary'],
                selectbackground=colors.get('selection_bg', colors['accent']),
                highlightbackground=colors['border'],
                highlightcolor=colors['border']
            )
            # Update outer Frame wrapper (ModernScrolledText.config only updates inner text!)
            try:
                tk.Frame.configure(self.log_text, bg=colors['section_bg'],
                                   highlightcolor=colors['border'], highlightbackground=colors['border'])
            except Exception:
                pass

        # Update summary_frame (from split log section) if it exists
        if hasattr(self, '_summary_frame') and self._summary_frame:
            try:
                if self._summary_frame.winfo_exists():
                    self._summary_frame.config(bg=colors['section_bg'])
            except Exception:
                pass  # Frame may have been destroyed

        # Update placeholder_label (from split log section) if it exists
        if hasattr(self, '_placeholder_label') and self._placeholder_label:
            try:
                if self._placeholder_label.winfo_exists():
                    self._placeholder_label.config(
                        bg=colors['section_bg'],
                        fg=colors.get('text_secondary', colors['text_primary'])
                    )
            except Exception:
                pass  # Label may have been destroyed

        # Update summary_text (from split log section) if it exists
        if hasattr(self, '_summary_text') and self._summary_text:
            try:
                if self._summary_text.winfo_exists():
                    self._summary_text.config(
                        bg=colors['section_bg'],
                        fg=colors['text_primary'],
                        selectbackground=colors.get('selection_bg', colors['accent'])
                    )
                    # Update outer Frame wrapper for ModernScrolledText
                    tk.Frame.configure(self._summary_text, bg=colors['section_bg'])
            except Exception:
                pass  # Widget may have been destroyed

        # Update summary_header icon label (from split log section) if it exists
        if hasattr(self, '_summary_header') and self._summary_header:
            try:
                if self._summary_header.winfo_exists():
                    self._summary_header.config(bg=colors['background'])
                    # Update all children (icon and text labels)
                    for child in self._summary_header.winfo_children():
                        try:
                            child.config(bg=colors['background'])
                        except Exception:
                            pass
            except Exception:
                pass  # Widget may have been destroyed

        # Update log_header icon label (from split log section) if it exists
        if hasattr(self, '_log_header') and self._log_header:
            try:
                if self._log_header.winfo_exists():
                    self._log_header.config(bg=colors['background'])
                    # Update all children (icon and text labels)
                    for child in self._log_header.winfo_children():
                        try:
                            child.config(bg=colors['background'])
                        except Exception:
                            pass
            except Exception:
                pass  # Widget may have been destroyed

        # Update primary buttons (Browse buttons and action buttons)
        # Use background for canvas since buttons are inside Section.TFrame content areas
        section_canvas_bg = colors['background']
        is_dark = colors.get('background', '') == '#0d0d1a'
        for btn in self._primary_buttons:
            try:
                if btn.winfo_exists():
                    btn.update_colors(
                        bg=colors['button_primary'],
                        hover_bg=colors['button_primary_hover'],
                        pressed_bg=colors['button_primary_pressed'],
                        fg='#ffffff',
                        disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
                        disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
                        canvas_bg=section_canvas_bg
                    )
            except Exception:
                pass  # Button may have been destroyed

        # Update secondary buttons (Export Log, Clear Log, Reset All)
        for btn in self._secondary_buttons:
            try:
                if btn.winfo_exists():
                    btn.update_colors(
                        bg=colors['button_secondary'],
                        hover_bg=colors['button_secondary_hover'],
                        pressed_bg=colors['button_secondary_pressed'],
                        fg=colors['text_primary'],
                        canvas_bg=section_canvas_bg
                    )
            except Exception:
                pass  # Button may have been destroyed

    def create_file_input_section(self, parent: ttk.Widget, title: str,
                                 file_types: List[tuple], guide_text: List[str] = None) -> Dict[str, Any]:
        """
        Create a standardized file input section.

        Returns:
            Dict with 'frame', 'path_var', 'entry', 'browse_button'
        """
        colors = self._theme_manager.colors

        section_frame = ttk.LabelFrame(parent, text=title,
                                     style='Section.TLabelframe', padding="12")

        # Inner content frame - uses white background
        content_frame = ttk.Frame(section_frame, style='Section.TFrame', padding="15")
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)

        # LEFT: Guide (if provided)
        if guide_text:
            guide_frame = ttk.Frame(content_frame, style='Section.TFrame')
            guide_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 35))

            # Use dynamic colors from theme manager
            colors = self._theme_manager.colors

            for i, text in enumerate(guide_text):
                if i == 0:  # Title
                    ttk.Label(guide_frame, text=text,
                             style='Section.TLabel',
                             font=('Segoe UI', 10, 'bold'),
                             foreground=colors['info']).pack(anchor=tk.W)
                elif text.startswith("⚠️"):  # Warning line - use ttk for theme support
                    ttk.Label(guide_frame, text=f"   {text}",
                            style='Section.TLabel',
                            font=('Segoe UI', 9, 'italic'),
                            foreground=colors['warning']).pack(anchor=tk.W, pady=1)
                else:  # Steps
                    # Handle bullet points with proper indentation for wrapped text
                    if text.startswith('•'):
                        # Create a frame for each bullet point to control indentation
                        bullet_frame = ttk.Frame(guide_frame, style='Section.TFrame')
                        bullet_frame.pack(anchor=tk.W, pady=1, fill=tk.X)

                        # Bullet symbol
                        ttk.Label(bullet_frame, text="•", font=('Segoe UI', 9),
                                 style='Section.TLabel',
                                 foreground=colors['text_secondary']).pack(side=tk.LEFT, anchor=tk.N)

                        # Text content with proper wrapping and indentation
                        text_content = text[1:].strip()  # Remove bullet and leading space
                        text_label = ttk.Label(bullet_frame, text=text_content, font=('Segoe UI', 9),
                                             style='Section.TLabel',
                                             foreground=colors['text_secondary'],
                                             wraplength=280, justify=tk.LEFT)
                        text_label.pack(side=tk.LEFT, anchor=tk.N, padx=(5, 0), fill=tk.X, expand=True)
                    else:
                        # Regular text
                        ttk.Label(guide_frame, text=f"   {text}", font=('Segoe UI', 9),
                                 style='Section.TLabel',
                                 foreground=colors['text_secondary'],
                                 wraplength=300).pack(anchor=tk.W, pady=1)

        # RIGHT: File input
        input_frame = ttk.Frame(content_frame, style='Section.TFrame')
        input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        input_frame.columnconfigure(1, weight=1)

        # File path variable
        path_var = tk.StringVar()

        # File input row
        ttk.Label(input_frame, text="File Path:", style='Section.TLabel').grid(row=0, column=0, sticky=tk.W, pady=8)
        
        entry = ttk.Entry(input_frame, textvariable=path_var, width=80)
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        
        # Create bound method to avoid lambda closure issues
        def browse_file_command():
            self._browse_file(path_var, file_types)

        # Use tk.Button with proper hover/click states (darker on hover, darkest on click)
        browse_button = self.create_primary_button(
            input_frame, "📂  Browse", browse_file_command, width=10)
        browse_button.grid(row=0, column=2, pady=8)
        
        return {
            'frame': section_frame,
            'path_var': path_var,
            'entry': entry,
            'browse_button': browse_button,
            'input_frame': input_frame
        }
    
    def create_log_section(self, parent: ttk.Widget, title: str = "📊 Analysis & Progress Log", labelwidget: tk.Widget = None) -> Dict[str, Any]:
        """
        Create a standardized log section.

        Args:
            parent: Parent widget to contain the log section
            title: Title string for the LabelFrame (ignored if labelwidget is provided)
            labelwidget: Optional custom widget to use as the label (icon + text frame)

        Returns:
            Dict with 'frame', 'text_widget', 'export_button', 'clear_button'
        """
        # Use dynamic colors from theme manager
        colors = self._theme_manager.colors

        # Create LabelFrame with either labelwidget or text
        if labelwidget:
            log_frame = ttk.LabelFrame(parent, labelwidget=labelwidget,
                                     style='Section.TLabelframe', padding="12")
        else:
            log_frame = ttk.LabelFrame(parent, text=title,
                                     style='Section.TLabelframe', padding="12")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        # Inner content frame - uses white background
        content_frame = ttk.Frame(log_frame, style='Section.TFrame', padding="15")
        content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        # Log text area with horizontal scrolling and no word wrap - flat, modern style
        # Use ModernScrolledText for themed scrollbar support
        # Use section_bg for background to match surrounding frame
        log_text = ModernScrolledText(
            content_frame, height=12, width=75, font=('Cascadia Code', 9), state=tk.DISABLED,
            bg=colors['section_bg'], fg=colors['text_primary'],
            selectbackground=colors.get('selection_bg', colors['accent']), relief='flat', borderwidth=0,
            wrap=tk.NONE, highlightthickness=1, highlightcolor=colors['border'],
            highlightbackground=colors['border'], padx=5, pady=5,
            theme_manager=self._theme_manager, auto_hide_scrollbar=False
        )
        log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Log controls
        log_controls = ttk.Frame(content_frame, style='Section.TFrame')
        log_controls.grid(row=0, column=1, sticky=tk.N, padx=(10, 0))
        
        # Create bound methods to avoid lambda closure issues
        def export_log_command():
            self._export_log(log_text)
        
        def clear_log_command():
            self._clear_log(log_text)
        
        # Use tk.Button for full control over hover/click states
        # Use two spaces after emoji for consistent icon-to-text spacing
        export_button = self.create_secondary_button(
            log_controls, "Export Log", export_log_command, width=14)
        export_button.pack(pady=(0, 5), anchor=tk.W)

        clear_button = self.create_secondary_button(
            log_controls, "Clear Log", clear_log_command, width=14)
        clear_button.pack(anchor=tk.W)

        self.log_text = log_text  # Store reference for logging

        return {
            'frame': log_frame,
            'text_widget': log_text,
            'export_button': export_button,
            'clear_button': clear_button
        }

    def create_log_section_with_options(self, parent: ttk.Widget, title: str = "📊 Analysis & Cleanup Log") -> Dict[str, Any]:
        """Create a log section with an options panel on the right (for cleanup tool)"""
        # Use dynamic colors from theme manager
        colors = self._theme_manager.colors

        log_frame = ttk.LabelFrame(parent, text=title,
                                 style='Section.TLabelframe', padding="12")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        # Inner content frame - uses white background
        main_container = ttk.Frame(log_frame, style='Section.TFrame', padding="15")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_container.columnconfigure(0, weight=1)  # Log area gets most space
        main_container.rowconfigure(0, weight=1)

        # Left side - Log area
        log_container = ttk.Frame(main_container, style='Section.TFrame')
        log_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 20))  # Increased padding for button space
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)

        # Log text area with proper scrollbars and sizing - flat, modern style
        # Use ModernScrolledText for themed scrollbar support
        # Use section_bg for background to match surrounding frame
        log_text = ModernScrolledText(
            log_container, height=10, width=55, font=('Cascadia Code', 9), state=tk.DISABLED,
            bg=colors['section_bg'], fg=colors['text_primary'],
            selectbackground=colors.get('selection_bg', colors['accent']), relief='flat', borderwidth=0,
            wrap=tk.NONE, exportselection=False,
            highlightthickness=1, highlightcolor=colors['border'],
            highlightbackground=colors['border'], padx=5, pady=5,
            theme_manager=self._theme_manager, auto_hide_scrollbar=False
        )
        log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 2), pady=(0, 2))  # Small padding for scrollbar space
        
        # Right side - Options area (wider to accommodate buttons)
        options_frame = ttk.Frame(main_container, style='Section.TFrame')
        options_frame.grid(row=0, column=1, sticky=(tk.N, tk.E), padx=(20, 0))  # Increased padding

        # Log controls in options area
        log_controls = ttk.Frame(options_frame, style='Section.TFrame')
        log_controls.pack(fill=tk.X, pady=(0, 20))
        
        # Create bound methods to avoid lambda closure issues
        def export_log_command():
            self._export_log(log_text)
        
        def clear_log_command():
            self._clear_log(log_text)
        
        # Use tk.Button for full control over hover/click states
        # Use two spaces after emoji for consistent icon-to-text spacing
        export_button = self.create_secondary_button(
            log_controls, "Export Log", export_log_command, width=14)
        export_button.pack(pady=(0, 5), fill=tk.X)

        clear_button = self.create_secondary_button(
            log_controls, "Clear Log", clear_log_command, width=14)
        clear_button.pack(fill=tk.X)
        
        self.log_text = log_text  # Store reference for logging

        return {
            'frame': log_frame,
            'text_widget': log_text,
            'export_button': export_button,
            'clear_button': clear_button,
            'options_frame': options_frame  # New: options area for additional content
        }

    def _load_icon_for_button(self, icon_name: str, size: int = 14) -> Optional['ImageTk.PhotoImage']:
        """Load an SVG or PNG icon for use in buttons.

        Args:
            icon_name: Name of the icon file (without extension)
            size: Target size in pixels

        Returns:
            ImageTk.PhotoImage or None if loading fails
        """
        if not PIL_AVAILABLE:
            return None

        # Get icon path
        icons_dir = Path(__file__).parent.parent / "assets" / "Tool Icons"
        svg_path = icons_dir / f"{icon_name}.svg"
        png_path = icons_dir / f"{icon_name}.png"

        try:
            img = None

            # Try SVG first (if cairosvg available)
            if CAIROSVG_AVAILABLE and svg_path.exists():
                # Render at 4x size for quality, then downscale
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=size * 4,
                    output_height=size * 4
                )
                img = Image.open(io.BytesIO(png_data))
            elif png_path.exists():
                img = Image.open(png_path)

            if img is None:
                return None

            # Ensure RGBA mode
            if img.mode != 'RGBA':
                img = img.convert('RGBA')

            # Resize with high-quality resampling
            img = img.resize((size, size), Image.Resampling.LANCZOS)

            return ImageTk.PhotoImage(img)

        except Exception as e:
            print(f"Failed to load icon {icon_name}: {e}")
            return None

    def create_icon_label(self, parent: tk.Widget, icon_name: str, text: str,
                          icon_size: int = 16, font: tuple = ('Segoe UI Semibold', 11),
                          fg_color: str = None, bg_color: str = None,
                          style: str = None) -> tk.Frame:
        """
        Create a label with an SVG icon and text.

        Args:
            parent: Parent widget
            icon_name: Name of the SVG icon file (without extension)
            text: Label text
            icon_size: Size of the icon in pixels
            font: Font tuple for the text
            fg_color: Foreground (text) color. If None, uses theme's title_color
            bg_color: Background color. If None, uses parent's background
            style: Optional ttk style name to get colors from

        Returns:
            tk.Frame containing the icon and label
        """
        colors = self._theme_manager.colors

        # Determine colors
        if bg_color is None:
            try:
                bg_color = parent.cget('bg')
            except Exception:
                bg_color = colors.get('background', '#ffffff')

        if fg_color is None:
            fg_color = colors.get('title_color', colors.get('accent', '#00a8a8'))

        # Create container frame
        frame = tk.Frame(parent, bg=bg_color)

        # Load icon
        icon_img = self._load_icon_for_button(icon_name, size=icon_size)

        if icon_img:
            # Store reference to prevent garbage collection
            if not hasattr(self, '_label_icons'):
                self._label_icons = {}
            self._label_icons[f'{icon_name}_{id(frame)}'] = icon_img

            # Create icon label
            icon_label = tk.Label(frame, image=icon_img, bg=bg_color)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))

        # Create text label
        text_label = tk.Label(frame, text=text, font=font, fg=fg_color, bg=bg_color)
        text_label.pack(side=tk.LEFT)

        # Store references for theme updates
        frame._icon_name = icon_name
        frame._icon_size = icon_size
        frame._text = text
        frame._font = font
        frame._fg_color_key = 'title_color'  # For theme updates

        return frame

    def create_split_log_section(self, parent: ttk.Widget, title: str = "Analysis & Progress",
                                   labelwidget: tk.Widget = None) -> Dict[str, Any]:
        """
        Create a split log section with summary panel (left) and progress log (right).

        Layout:
        - Left (60%): Summary/details panel with scrolledtext (word wrap, pill scrollbar)
        - Right (40%): Progress log with scrolledtext and icon buttons in header

        Args:
            parent: Parent widget
            title: Text for the section header (used if labelwidget not provided)
            labelwidget: Optional custom widget for section header (overrides title)

        Returns:
            Dict with 'frame', 'summary_frame', 'summary_text', 'log_text', 'export_button',
                   'clear_button', 'log_container', 'placeholder_label'
        """
        colors = self._theme_manager.colors

        # Main frame - use labelwidget if provided, otherwise create icon label
        if labelwidget:
            log_frame = ttk.LabelFrame(parent, labelwidget=labelwidget,
                                       style='Section.TLabelframe', padding="12")
        else:
            # Create section header with analyze.svg icon
            section_header = self.create_icon_label(
                parent, icon_name="analyze", text=title,
                icon_size=18, font=('Segoe UI Semibold', 12),
                bg_color=colors['section_bg']
            )
            log_frame = ttk.LabelFrame(parent, labelwidget=section_header,
                                       style='Section.TLabelframe', padding="12")
        log_frame.columnconfigure(0, weight=1)  # Summary gets 50%
        log_frame.columnconfigure(1, weight=1)  # Log gets 50%
        log_frame.rowconfigure(0, weight=1)

        # Inner content frame - uses white background
        content_frame = ttk.Frame(log_frame, style='Section.TFrame', padding="15")
        content_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Use uniform group to ensure columns stay equal width regardless of content
        content_frame.columnconfigure(0, weight=1, minsize=225, uniform="split_cols")  # Summary column (50%)
        content_frame.columnconfigure(1, weight=1, minsize=225, uniform="split_cols")  # Log column (50%)
        content_frame.rowconfigure(0, weight=1)

        # ===== LEFT SIDE: Summary/Details Panel =====
        summary_container = ttk.Frame(content_frame, style='Section.TFrame')
        summary_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        summary_container.columnconfigure(0, weight=1)
        summary_container.rowconfigure(0, minsize=30)  # Header row - fixed height to match right side
        summary_container.rowconfigure(1, weight=1)  # Text area row expands

        # Summary header frame - contains label and optional filter controls
        # Use background (outer container color) for header labels
        summary_header_frame = ttk.Frame(summary_container, style='Section.TFrame')
        summary_header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        summary_header_frame.columnconfigure(0, weight=1)  # Label takes available space

        # Summary header label inside frame
        summary_header = self.create_icon_label(
            summary_header_frame, icon_name="bar-chart", text="Analysis Summary",
            icon_size=16, font=('Segoe UI Semibold', 11),
            bg_color=colors['background']
        )
        summary_header.pack(side=tk.LEFT)

        # Store reference for external access (e.g., adding filter controls)
        self._summary_header_frame = summary_header_frame

        # Summary content frame (contains either placeholder or text widget)
        # Use section_bg (#161627 dark, #f5f5f7 light) for consistent background
        summary_frame = tk.Frame(summary_container, bg=colors['section_bg'],
                                 highlightthickness=0, padx=8, pady=8)
        summary_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        summary_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)

        # Placeholder message (hidden when summary_text is populated)
        # Use section_bg to match summary_frame
        placeholder_label = tk.Label(summary_frame,
                                     text="Run analysis to see results",
                                     font=('Segoe UI', 10, 'italic'),
                                     bg=colors['section_bg'],
                                     fg=colors.get('text_secondary', colors['text_primary']),
                                     anchor=tk.CENTER)
        placeholder_label.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        # Summary text widget - ModernScrolledText with themed scrollbar (no border - frame has it)
        # Use section_bg (#161627 dark, #f5f5f7 light) for consistent background
        summary_text = ModernScrolledText(
            summary_frame, height=10, width=40, font=('Segoe UI', 9), state=tk.DISABLED,
            bg=colors['section_bg'], fg=colors['text_primary'],
            selectbackground=colors.get('selection_bg', colors['accent']), relief='flat', borderwidth=0,
            wrap=tk.WORD, highlightthickness=0, padx=5, pady=5,
            theme_manager=self._theme_manager, auto_hide_scrollbar=True
        )
        # Don't grid yet - will be shown when placeholder is hidden

        # ===== RIGHT SIDE: Progress Log Panel =====
        log_outer_container = ttk.Frame(content_frame, style='Section.TFrame')
        log_outer_container.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_outer_container.columnconfigure(0, weight=1)
        log_outer_container.rowconfigure(0, minsize=30)  # Header row - fixed height to match left side
        log_outer_container.rowconfigure(1, weight=1)  # Text area row expands

        # Log header row - contains label and icon buttons
        log_header_frame = ttk.Frame(log_outer_container, style='Section.TFrame')
        log_header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        log_header_frame.columnconfigure(0, weight=1)  # Label takes available space

        # Log header label (left-aligned) - uses icon label with log-file.svg
        # Use background (outer container color) for header labels
        log_header_label = self.create_icon_label(
            log_header_frame, icon_name="log-file", text="Progress Log",
            icon_size=16, font=('Segoe UI Semibold', 11),
            bg_color=colors['background']
        )
        log_header_label.grid(row=0, column=0, sticky=tk.W)

        # Create bound methods for buttons
        def export_log_command():
            self._export_log(log_text)

        def clear_log_command():
            self._clear_log(log_text)

        # Load icons for buttons
        save_icon = self._load_icon_for_button("save", size=14)
        eraser_icon = self._load_icon_for_button("eraser", size=14)

        # Store icons to prevent garbage collection
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}
        self._button_icons['save'] = save_icon
        self._button_icons['eraser'] = eraser_icon

        # Icon buttons frame (right-aligned in header)
        icon_buttons_frame = ttk.Frame(log_header_frame, style='Section.TFrame')
        icon_buttons_frame.grid(row=0, column=1, sticky=tk.E)

        # Export button (save icon)
        export_button = SquareIconButton(
            icon_buttons_frame, icon=save_icon, command=export_log_command,
            tooltip_text="Export Log", size=26, radius=6
        )
        export_button.pack(side=tk.LEFT, padx=(0, 4))

        # Clear button (eraser icon)
        clear_button = SquareIconButton(
            icon_buttons_frame, icon=eraser_icon, command=clear_log_command,
            tooltip_text="Clear Log", size=26, radius=6
        )
        clear_button.pack(side=tk.LEFT)

        # Log content container
        log_container = ttk.Frame(log_outer_container, style='Section.TFrame')
        log_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)

        # Log text area - ModernScrolledText with themed scrollbar for dark/light mode
        # Use section_bg (#161627 dark, #f5f5f7 light) for consistent background
        log_text = ModernScrolledText(
            log_container, height=10, width=45, font=('Cascadia Code', 9), state=tk.DISABLED,
            bg=colors['section_bg'], fg=colors['text_primary'],
            selectbackground=colors.get('selection_bg', colors['accent']), relief='flat', borderwidth=0,
            wrap=tk.WORD, highlightthickness=1, highlightcolor=colors['border'],
            highlightbackground=colors['border'], padx=5, pady=5,
            theme_manager=self._theme_manager, auto_hide_scrollbar=False
        )
        log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.log_text = log_text  # Store reference for logging

        # Store references for theme updates (background color)
        self._summary_frame = summary_frame
        self._placeholder_label = placeholder_label
        self._summary_text = summary_text
        self._summary_header = summary_header
        self._log_header = log_header_label

        return {
            'frame': log_frame,
            'summary_frame': summary_frame,
            'summary_text': summary_text,
            'summary_header': summary_header,
            'summary_header_frame': summary_header_frame,  # Frame for adding filter controls
            'log_header': log_header_label,
            'log_text': log_text,
            'export_button': export_button,
            'clear_button': clear_button,
            'log_container': log_container,
            'placeholder_label': placeholder_label
        }

    def create_action_buttons(self, parent: ttk.Widget, buttons: List[Dict[str, Any]]) -> Dict[str, ttk.Button]:
        """
        Create standardized action buttons.
        
        Args:
            parent: Parent widget
            buttons: List of button configs with 'text', 'command', 'style', 'state'
            
        Returns:
            Dict mapping button text to button widget
        """
        button_frame = ttk.Frame(parent)
        button_widgets = {}
        
        for i, config in enumerate(buttons):
            button = ttk.Button(
                button_frame, 
                text=config['text'],
                command=config['command'],
                style=config.get('style', 'Action.TButton'),
                state=config.get('state', tk.NORMAL)
            )
            button.pack(side=tk.LEFT, padx=(0, 15) if i < len(buttons) - 1 else 0)
            button_widgets[config['text']] = button
        
        return button_widgets
    
    def create_progress_bar(self, parent: ttk.Widget) -> Dict[str, Any]:
        """Create a standardized enhanced progress bar
        
        Returns:
            Dict with 'frame', 'progress_bar'
        """
        # Progress container frame
        self.progress_frame = ttk.Frame(parent)
        
        # Just the progress bar - no status text
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', maximum=100)
        self.progress_bar.pack(fill='x')
        
        # Set progress_label to None since we're not using it
        self.progress_label = None
        
        return {
            'frame': self.progress_frame,
            'progress_bar': self.progress_bar
        }
    
    def log_message(self, message: str):
        """Log message to the tab's log area"""
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.config(state=tk.DISABLED)
            self.log_text.see(tk.END)
            self.frame.update_idletasks()
    
    def update_progress(self, progress_percent: int, message: str = "", show: bool = True, persist: bool = False):
        """Update progress bar - Universal progress system
        
        Args:
            progress_percent: Progress percentage (0-100)
            message: Status message to display (logged but not shown in UI)
            show: Whether to show or hide progress
            persist: If True, keep progress bar visible after operation (don't hide on 100%)
        """
        if show:
            # Show progress frame if not already visible
            if self.progress_frame and not self.progress_frame.winfo_viewable():
                # Auto-position progress frame - subclasses can override grid position
                self._position_progress_frame()
            
            # Update progress bar value
            if self.progress_bar:
                self.progress_bar['value'] = progress_percent
            
            # Set persistence flag if requested
            if persist:
                self.progress_persist = True
            
            # Log the progress message with percentage (no UI status text)
            if progress_percent > 0 and message:
                self.log_message(f"⏳ {progress_percent}% - {message}")
            elif message:
                self.log_message(f"⏳ {message}")
            elif progress_percent > 0:
                self.log_message(f"⏳ {progress_percent}% complete")

            # Set is_busy based on progress - 100% means operation is complete
            # This is separate from persist which controls progress bar visibility
            self.is_busy = progress_percent < 100
            
        else:
            # Hide progress frame only if not persisting
            if not self.progress_persist:
                if self.progress_bar:
                    self.progress_bar['value'] = 0  # Reset to empty
                if self.progress_frame:
                    self.progress_frame.grid_remove()
            else:
                # Keep progress bar visible but reset to 0
                if self.progress_bar:
                    self.progress_bar['value'] = 0
                
            self.is_busy = False
    
    def _position_progress_frame(self):
        """Position the progress frame - can be overridden by subclasses"""
        # Default positioning - subclasses should override this method
        # to position the progress frame appropriately for their layout
        if self.progress_frame:
            # Try to find a good default position - bottom of the frame
            self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))
    
    def run_in_background(self, target_func: Callable, 
                         success_callback: Callable = None,
                         error_callback: Callable = None,
                         progress_steps: List[tuple] = None):
        """
        Run a function in background thread with enhanced progress indication.
        
        Args:
            target_func: Function to run in background
            success_callback: Called on success with result
            error_callback: Called on error with exception
            progress_steps: List of (message, percentage) tuples for progress updates
        """
        import traceback
        
        def thread_target():
            """Thread target with proper error handling and closures."""
            result = None
            caught_error = None
            
            try:
                # Show initial progress only if no custom progress steps
                if progress_steps:
                    first_step = progress_steps[0]
                    self.frame.after(0, lambda: self.update_progress(first_step[1], first_step[0]))
                else:
                    # Don't show generic progress - let the target_func handle its own progress
                    pass
                
                result = target_func()
                
            except Exception as e:
                caught_error = e
                # Log the full traceback for debugging
                self.log_message(f"❌ Background operation failed: {e}")
                self.log_message(f"📋 Traceback: {traceback.format_exc()}")
            
            # Schedule callbacks on main thread with proper closures
            def schedule_callbacks():
                try:
                    if caught_error is not None:
                        # Handle error
                        if error_callback:
                            error_callback(caught_error)
                        else:
                            self._default_error_handler(caught_error)
                    else:
                        # Show completion progress if steps provided
                        if progress_steps:
                            self.update_progress(100, "Operation complete!")
                            import time
                            time.sleep(0.5)  # Brief pause to show completion
                        
                        # Handle success
                        if success_callback:
                            success_callback(result)
                
                except Exception as callback_error:
                    # Handle callback errors
                    self.log_message(f"❌ Callback error: {callback_error}")
                    self._default_error_handler(callback_error)
                
                finally:
                    # Only hide progress if not persisting
                    if not self.progress_persist:
                        self.update_progress(0, "", False)
            
            # Schedule on main thread
            if hasattr(self, 'frame') and self.frame:
                self.frame.after(0, schedule_callbacks)
            else:
                # Fallback if no frame available
                schedule_callbacks()
        
        thread = threading.Thread(target=thread_target, daemon=True)
        thread.start()
    
    def _default_error_handler(self, error: Exception):
        """Default error handler"""
        self.log_message(f"❌ Error: {error}")
        self.show_error("Error", str(error))
    
    def _browse_file(self, path_var: tk.StringVar, file_types: List[tuple]):
        """Common file browsing logic"""
        file_path = filedialog.askopenfilename(
            title="Select File",
            filetypes=file_types
        )
        if file_path:
            path_var.set(file_path)
    
    def _export_log(self, log_widget):
        """Export log content to file"""
        try:
            log_content = log_widget.get(1.0, tk.END)
            file_path = filedialog.asksaveasfilename(
                title="Export Log", defaultextension=".txt",
                filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
            )

            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"{self.tool_name} - Analysis Log\n")
                    f.write(f"Generated by {AppConstants.COMPANY_NAME}\n")
                    f.write(f"{'='*50}\n\n")
                    f.write(log_content)

                self.log_message(f"✅ Log exported to: {file_path}")
                self.show_info("Export Complete", f"Log exported successfully!\n\n{file_path}")

        except Exception as e:
            self.show_error("Export Error", f"Failed to export log: {e}")
    
    def _clear_log(self, log_widget):
        """Clear log content"""
        log_widget.config(state=tk.NORMAL)
        log_widget.delete(1.0, tk.END)
        log_widget.config(state=tk.DISABLED)
        self._show_welcome_message()
    
    def _show_welcome_message(self):
        """Show welcome message - can be overridden by subclasses"""
        self.log_message(f"🎉 Welcome to {self.tool_name}!")
        self.log_message("=" * 60)
    
    def create_help_window(self, title: str, content_creator: Callable) -> tk.Toplevel:
        """
        Create a standardized help window.

        Args:
            title: Window title
            content_creator: Function that creates content in the window

        Returns:
            The help window
        """
        # Use dynamic colors from theme manager
        colors = self._theme_manager.colors

        help_window = tk.Toplevel(self.main_app.root)
        help_window.withdraw()  # Hide initially to prevent flicker
        help_window.title(title)
        help_window.geometry("650x830")
        help_window.resizable(False, False)
        help_window.transient(self.main_app.root)
        help_window.grab_set()
        help_window.configure(bg=colors['background'])

        # Set AE favicon icon
        try:
            help_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Create content
        content_creator(help_window)

        # Bind escape key with proper closure
        def close_window(event=None):
            help_window.destroy()

        help_window.bind('<Escape>', close_window)

        # Center dialog on parent window after content is created
        help_window.update_idletasks()
        dialog_width = help_window.winfo_reqwidth()
        dialog_height = help_window.winfo_reqheight()
        parent_x = self.main_app.root.winfo_rootx()
        parent_y = self.main_app.root.winfo_rooty()
        parent_width = self.main_app.root.winfo_width()
        parent_height = self.main_app.root.winfo_height()
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        help_window.geometry(f"+{x}+{y}")

        # Set dark/light title bar BEFORE showing to prevent white flash
        help_window.update()
        self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled
        help_window.focus_force()

        return help_window
    
    def _set_dialog_title_bar_color(self, window, dark_mode: bool = True):
        """Set Windows title bar to dark or light mode"""
        try:
            import ctypes
            window.update()  # Ensure window is created
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if dark_mode else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass  # Ignore on non-Windows or if API fails

    def _create_themed_dialog(self, title: str, message: str, dialog_type: str = 'info',
                               buttons: list = None, icon_path: str = None,
                               max_content_height: int = None) -> any:
        """
        Create a themed dialog that matches the app's dark/light mode.

        Args:
            title: Dialog title
            message: Message to display
            dialog_type: 'info', 'error', 'warning', or 'question'
            buttons: List of button configs [{'text': str, 'value': any, 'style': 'primary'|'secondary'}]
            icon_path: Optional custom icon path (SVG or PNG). If None, uses default AE logo.
            max_content_height: Optional max height for message area. If set and content exceeds,
                               enables scrolling with modern themed scrollbar.

        Returns:
            The value of the clicked button, or None if closed
        """
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        dialog = tk.Toplevel(self.main_app.root)
        dialog.title(title)
        dialog.transient(self.main_app.root)
        # Allow resizing when scrollable content is enabled (both axes needed for Windows resize handles)
        is_resizable = max_content_height is not None
        dialog.resizable(is_resizable, is_resizable)
        dialog.configure(bg=colors['background'])

        # Withdraw dialog to prevent flash in upper-left before centering
        dialog.withdraw()

        # Set window icon to match main app (AE favicon)
        try:
            if hasattr(self.main_app, 'config') and hasattr(self.main_app.config, 'icon_path'):
                dialog.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass  # Ignore if icon cannot be set

        # Set dark/light title bar
        self._set_dialog_title_bar_color(dialog, is_dark)

        # Main frame
        main_frame = tk.Frame(dialog, bg=colors['background'], padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Icon and message row - expand vertically when scrollable content is enabled
        content_frame = tk.Frame(main_frame, bg=colors['background'])
        if max_content_height:
            content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        else:
            content_frame.pack(fill=tk.X, pady=(0, 20))

        # Load icon for dialog (custom icon_path or default AE logo)
        icon_image = None
        try:
            from pathlib import Path
            if icon_path:
                # Use custom icon path
                custom_path = Path(icon_path)
                if custom_path.exists() and PIL_AVAILABLE:
                    if custom_path.suffix.lower() == '.svg' and CAIROSVG_AVAILABLE:
                        # Convert SVG to PNG in memory
                        import cairosvg
                        import io
                        png_data = cairosvg.svg2png(url=str(custom_path), output_width=48, output_height=48)
                        img = Image.open(io.BytesIO(png_data))
                    else:
                        img = Image.open(custom_path)
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    img = img.resize((48, 48), Image.Resampling.LANCZOS)
                    icon_image = ImageTk.PhotoImage(img)

            if not icon_image:
                # Fall back to default AE logo
                logo_path = Path(__file__).parent.parent / "assets" / "Website Icon.png"
                if logo_path.exists() and PIL_AVAILABLE:
                    img = Image.open(logo_path)
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    # Resize to 32x32 for dialog icon
                    img = img.resize((32, 32), Image.Resampling.LANCZOS)
                    icon_image = ImageTk.PhotoImage(img)
        except Exception:
            pass  # Fall back to emoji if icon loading fails

        if icon_image:
            icon_label = tk.Label(content_frame, image=icon_image, bg=colors['background'])
            icon_label.image = icon_image  # Keep reference to prevent garbage collection
        else:
            # Fallback to emoji icons
            icons = {
                'info': 'ℹ️',
                'error': '❌',
                'warning': '⚠️',
                'question': '❓'
            }
            icon_colors = {
                'info': colors['info'],
                'error': colors['error'],
                'warning': colors['warning'],
                'question': colors['info']
            }
            icon_label = tk.Label(content_frame, text=icons.get(dialog_type, 'ℹ️'),
                                 font=('Segoe UI', 24),
                                 bg=colors['background'],
                                 fg=icon_colors.get(dialog_type, colors['info']))
        icon_label.pack(side=tk.LEFT, padx=(0, 15), anchor=tk.N)

        # Message - use scrollable text widget if max_content_height is set
        if max_content_height:
            # Create a frame to contain the scrollable text - allows expansion when window resized
            message_frame = tk.Frame(content_frame, bg=colors['background'], width=450)
            message_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            # Set minimum size for the content area (dialog minsize set later after geometry)

            # Use ModernScrolledText for themed scrollbar
            text_widget = ModernScrolledText(
                message_frame,
                wrap=tk.WORD,
                font=('Segoe UI', 10),
                bg=colors['background'],
                fg=colors['text_primary'],
                relief='flat',
                borderwidth=0,
                highlightthickness=0,
                padx=0,
                pady=0,
                theme_manager=self._theme_manager,
                auto_hide_scrollbar=True
            )
            text_widget.pack(fill=tk.BOTH, expand=True)
            text_widget.insert(1.0, message)
            text_widget.config(state=tk.DISABLED)
        else:
            msg_label = tk.Label(content_frame, text=message,
                                font=('Segoe UI', 10),
                                bg=colors['background'],
                                fg=colors['text_primary'],
                                justify=tk.LEFT,
                                wraplength=450)
            msg_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Result variable
        result = [None]

        # Button frame
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        # Default buttons if none provided
        if buttons is None:
            buttons = [{'text': 'OK', 'value': True, 'style': 'primary'}]

        def on_button_click(value):
            result[0] = value
            dialog.destroy()

        # Create buttons (right-aligned)
        for i, btn_config in enumerate(reversed(buttons)):
            btn_text = btn_config.get('text', 'OK')
            btn_value = btn_config.get('value', True)
            btn_style = btn_config.get('style', 'primary')

            if btn_style == 'primary':
                bg = colors['button_primary']
                hover_bg = colors['button_primary_hover']
                pressed_bg = colors['button_primary_pressed']
                fg = '#ffffff'
            else:
                bg = colors['button_secondary']
                hover_bg = colors['button_secondary_hover']
                pressed_bg = colors['button_secondary_pressed']
                fg = colors['text_primary']

            btn = RoundedButton(
                button_frame,
                text=btn_text,
                command=lambda v=btn_value: on_button_click(v),
                bg=bg,
                fg=fg,
                hover_bg=hover_bg,
                pressed_bg=pressed_bg,
                width=80,
                height=32,
                radius=6,
                font=('Segoe UI', 10)
            )
            btn.pack(side=tk.RIGHT, padx=(10, 0))  # 10px spacing between all buttons

        # Bind escape to close (returns None)
        dialog.bind('<Escape>', lambda e: dialog.destroy())
        # Bind Enter to first (primary) button
        dialog.bind('<Return>', lambda e: on_button_click(buttons[0].get('value', True)))

        # Center dialog on parent window after content is created
        dialog.update_idletasks()  # Force layout calculation
        dialog_width = dialog.winfo_reqwidth()
        dialog_height = dialog.winfo_reqheight()

        # When scrollable content is enabled, set larger default dimensions
        if max_content_height:
            # Default width: 580px (slightly wider than base 450px message width)
            # Default height: max_content_height + padding for icon, buttons, margins (~120px)
            default_width = 580
            default_height = max_content_height + 120
            dialog_width = max(dialog_width, default_width)
            dialog_height = max(dialog_height, default_height)

        parent_x = self.main_app.root.winfo_rootx()
        parent_y = self.main_app.root.winfo_rooty()
        parent_width = self.main_app.root.winfo_width()
        parent_height = self.main_app.root.winfo_height()
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2

        # Set both size and position when scrollable, just position otherwise
        if max_content_height:
            dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
            dialog.minsize(dialog_width, dialog_height)
        else:
            dialog.geometry(f"+{x}+{y}")

        # Now show the dialog (was withdrawn to prevent flash in upper-left)
        dialog.deiconify()
        dialog.grab_set()

        # Wait for dialog to close
        dialog.wait_window()

        return result[0]

    def show_error(self, title: str, message: str, icon_path: str = None):
        """Show themed error dialog"""
        self._create_themed_dialog(title, message, dialog_type='error', icon_path=icon_path)

    def show_info(self, title: str, message: str, icon_path: str = None):
        """Show themed info dialog"""
        self._create_themed_dialog(title, message, dialog_type='info', icon_path=icon_path)

    def show_warning(self, title: str, message: str, icon_path: str = None):
        """Show themed warning dialog"""
        self._create_themed_dialog(title, message, dialog_type='warning', icon_path=icon_path)

    def ask_yes_no(self, title: str, message: str, icon_path: str = None,
                   max_content_height: int = None) -> bool:
        """Show themed yes/no dialog

        Args:
            title: Dialog title
            message: Message to display
            icon_path: Optional custom icon path
            max_content_height: Optional max height for message area (enables scrolling if set)
        """
        result = self._create_themed_dialog(
            title, message, dialog_type='question',
            buttons=[
                {'text': 'Yes', 'value': True, 'style': 'primary'},
                {'text': 'No', 'value': False, 'style': 'secondary'}
            ],
            icon_path=icon_path,
            max_content_height=max_content_height
        )
        return result if result is not None else False
    
    def show_scrollable_info(self, title: str, content: str):
        """
        Show a scrollable info dialog with text content.
        Perfect for help dialogs, documentation, or long messages.

        Args:
            title: Window title
            content: Text content to display (can be multi-line)
        """
        # Use dynamic colors from theme manager
        colors = self._theme_manager.colors

        info_window = tk.Toplevel(self.main_app.root)
        info_window.title(title)
        info_window.geometry("700x650")
        info_window.resizable(True, True)
        info_window.transient(self.main_app.root)
        info_window.grab_set()

        # Center window
        info_window.geometry(f"+{self.main_app.root.winfo_rootx() + 50}+{self.main_app.root.winfo_rooty() + 50}")
        info_window.configure(bg=colors['background'])

        # Set dark/light title bar
        self._set_dialog_title_bar_color(info_window, self._theme_manager.is_dark)

        # Main frame
        main_frame = ttk.Frame(info_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_label = ttk.Label(main_frame, text=title,
                                font=('Segoe UI', 14, 'bold'),
                                foreground=colors['button_primary'])  # Blue in dark, teal in light
        header_label.pack(anchor=tk.W, pady=(0, 15))

        # Scrollable text area
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Create scrolled text widget - flat, modern style with themed scrollbar
        text_widget = ModernScrolledText(
            text_frame,
            wrap=tk.WORD,
            font=('Segoe UI', 10),
            bg=colors['surface'],
            fg=colors['text_primary'],
            selectbackground=colors.get('selection_bg', colors['accent']),
            relief='flat',
            borderwidth=0,
            padx=15,
            pady=15,
            highlightthickness=1,
            highlightcolor=colors['border'],
            highlightbackground=colors['border'],
            theme_manager=self._theme_manager,
            auto_hide_scrollbar=True
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        # Insert content
        text_widget.insert(1.0, content)
        text_widget.config(state=tk.DISABLED)  # Make read-only
        
        # Close button
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        close_button = ttk.Button(button_frame, text="❌ Close", 
                                 command=info_window.destroy,
                                 style='Action.TButton')
        close_button.pack()
        
        # Bind escape key to close
        info_window.bind('<Escape>', lambda event: info_window.destroy())


class FileInputMixin:
    """
    Mixin for tabs that need file input functionality.
    """
    
    def clean_file_path(self, path: str) -> str:
        """Clean file path by removing quotes and normalizing"""
        if not path:
            return path
        
        cleaned = path.strip()
        
        # Remove surrounding quotes
        if len(cleaned) >= 2:
            if (cleaned.startswith('"') and cleaned.endswith('"')) or \
               (cleaned.startswith("'") and cleaned.endswith("'")):
                cleaned = cleaned[1:-1]
        
        # Normalize path separators
        return str(Path(cleaned)) if cleaned else cleaned
    
    def setup_path_cleaning(self, path_var: tk.StringVar):
        """Setup automatic path cleaning on a StringVar"""
        def on_path_change(*args):
            current = path_var.get()
            cleaned = self.clean_file_path(current)
            if cleaned != current:
                path_var.set(cleaned)
        
        path_var.trace('w', on_path_change)


class ValidationMixin:
    """
    Mixin for tabs that need validation functionality.
    """
    
    def validate_file_exists(self, file_path: str, file_description: str = "File") -> None:
        """Validate that a file exists"""
        if not file_path:
            raise ValueError(f"{file_description} path is required")
        
        path_obj = Path(file_path)
        if not path_obj.exists():
            raise ValueError(f"{file_description} not found: {file_path}")
        
        if not path_obj.is_file():
            raise ValueError(f"{file_description} path must point to a file: {file_path}")
    
    def validate_pbip_file(self, file_path: str, file_description: str = "File") -> None:
        """Validate that a file is a valid PBIP file"""
        self.validate_file_exists(file_path, file_description)
        
        if not file_path.lower().endswith('.pbip'):
            raise ValueError(f"{file_description} must be a .pbip file")
        
        # Check for corresponding .Report directory
        path_obj = Path(file_path)
        report_dir = path_obj.parent / f"{path_obj.stem}.Report"
        if not report_dir.exists():
            raise ValueError(f"{file_description} missing corresponding .Report directory")


class SectionPanelMixin:
    """
    Mixin providing shared section header and icon loading functionality
    for tool panels.

    Usage:
        class MyPanel(SectionPanelMixin, ttk.LabelFrame):
            def __init__(self, parent):
                ttk.LabelFrame.__init__(self, parent, style='Section.TLabelframe', padding="12")
                self._init_section_panel_mixin()
                self._create_section_header(parent, "Section Title", "icon-name")
                ...
    """

    def _init_section_panel_mixin(self):
        """Initialize mixin state. Call this in __init__ before using other methods."""
        self._theme_manager = get_theme_manager()
        self._section_header_widgets: List[tuple] = []
        self._header_icon = None
        self._mixin_logger = None

    def _create_section_header(self, parent, text: str, icon_name: str):
        """
        Create a section header with icon and set it as the labelwidget.

        Args:
            parent: The parent widget (typically the frame containing this LabelFrame)
            text: Header text to display
            icon_name: Name of the SVG icon (without .svg extension)
        """
        colors = self._theme_manager.colors
        # Match BaseToolTab.create_section_header pattern - use section_bg
        bg_color = colors.get('section_bg', colors['background'])

        # Create header frame as child of parent (not self)
        header_frame = tk.Frame(parent, bg=bg_color)

        # Load icon
        icon = self._load_section_icon(icon_name, size=16)
        icon_label = None
        if icon:
            icon_label = tk.Label(header_frame, image=icon, bg=bg_color)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            # Store reference to prevent garbage collection
            icon_label._icon_ref = icon
            self._header_icon = icon

        # Text label with title styling
        text_label = tk.Label(
            header_frame, text=text, bg=bg_color,
            fg=colors['title_color'], font=('Segoe UI Semibold', 11)
        )
        text_label.pack(side=tk.LEFT)

        # Store for theme updates
        self._section_header_widgets.append((header_frame, icon_label, text_label))

        # Configure this LabelFrame to use the header widget
        self.configure(labelwidget=header_frame)

    def _load_section_icon(self, icon_name: str, size: int = 16) -> Optional['ImageTk.PhotoImage']:
        """
        Load an SVG icon for section header.

        Args:
            icon_name: Name of the icon file (without extension)
            size: Target size in pixels (default 16)

        Returns:
            PhotoImage if successful, None otherwise
        """
        if not PIL_AVAILABLE:
            return None

        icons_dir = Path(__file__).parent.parent / "assets" / "Tool Icons"
        svg_path = icons_dir / f"{icon_name}.svg"
        png_path = icons_dir / f"{icon_name}.png"

        try:
            img = None

            # Try SVG first (if cairosvg available)
            if CAIROSVG_AVAILABLE and svg_path.exists():
                # Render at 4x size for quality, then downscale
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=size * 4,
                    output_height=size * 4
                )
                img = Image.open(io.BytesIO(png_data))
            elif png_path.exists():
                img = Image.open(png_path)

            if img is None:
                return None

            # Resize to target size with high-quality resampling
            img = img.resize((size, size), Image.Resampling.LANCZOS)

            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def _update_section_header_theme(self):
        """
        Update section header colors when theme changes.
        Call this from on_theme_changed().
        Matches BaseToolTab.on_theme_changed() pattern exactly.
        """
        colors = self._theme_manager.colors
        bg_color = colors.get('section_bg', colors['background'])
        for header_frame, icon_label, text_label in self._section_header_widgets:
            try:
                header_frame.configure(bg=bg_color)
                if icon_label:
                    icon_label.configure(bg=bg_color)
                text_label.configure(bg=bg_color, fg=colors['title_color'])
            except Exception:
                pass

        # Force the LabelFrame to re-apply its style
        try:
            self.configure(style='Section.TLabelframe')
        except Exception:
            pass
