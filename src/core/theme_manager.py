"""
Theme Manager - Handles dark/light mode switching and persistence
Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
import json
from pathlib import Path
from typing import Callable, List, Optional

# ttkbootstrap for modern UI styling
try:
    import ttkbootstrap as ttkb
    from ttkbootstrap import Style as TtkbStyle
    TTKBOOTSTRAP_AVAILABLE = True
except ImportError:
    TTKBOOTSTRAP_AVAILABLE = False

from core.constants import AppConstants


class ThemeManager:
    """
    Manages application theming with dark/light mode support.
    Handles theme persistence and dynamic style updates.
    Uses ttkbootstrap for modern rounded buttons and styling when available.
    """

    _instance: Optional['ThemeManager'] = None

    # ttkbootstrap theme mapping - maps our dark/light to ttkbootstrap themes
    TTKB_THEMES = {
        'dark': 'cyborg',   # Dark theme with cyan accents (matches our teal brand)
        'light': 'cosmo'    # Clean light theme
    }

    def __init__(self):
        self._current_theme = AppConstants.DEFAULT_THEME
        self._theme_callbacks: List[Callable[[str], None]] = []
        self._settings_path = Path.home() / "AppData" / "Local" / "AnalyticEndeavors" / "ae_multitool_settings.json"
        self._root: Optional[tk.Tk] = None
        self._style: Optional[ttk.Style] = None
        self._ttkb_style: Optional['TtkbStyle'] = None

        # Load saved preference
        self._load_theme_preference()

    @classmethod
    def get_instance(cls) -> 'ThemeManager':
        """Get singleton instance"""
        if cls._instance is None:
            cls._instance = ThemeManager()
        return cls._instance

    @classmethod
    def reset_instance(cls):
        """Reset singleton instance (for testing)"""
        cls._instance = None

    def initialize(self, root: tk.Tk):
        """Initialize theme manager with root window"""
        self._root = root

        # Use ttkbootstrap style if available
        if TTKBOOTSTRAP_AVAILABLE:
            self._ttkb_style = TtkbStyle()
            self._style = self._ttkb_style
        else:
            self._style = ttk.Style()

        self._apply_theme(self._current_theme)

    @property
    def current_theme(self) -> str:
        return self._current_theme

    @property
    def colors(self) -> dict:
        """Get current theme colors"""
        return AppConstants.THEMES[self._current_theme]

    @property
    def is_dark(self) -> bool:
        return self._current_theme == 'dark'

    def toggle_theme(self):
        """Toggle between dark and light themes"""
        new_theme = 'light' if self._current_theme == 'dark' else 'dark'
        self.set_theme(new_theme)

    def set_theme(self, theme: str):
        """Set the application theme"""
        if theme not in AppConstants.THEMES:
            return

        self._current_theme = theme
        self._apply_theme(theme)
        self._save_theme_preference()

        # Notify all registered callbacks
        for callback in self._theme_callbacks:
            try:
                callback(theme)
            except Exception:
                # Silently skip failed callbacks (e.g., destroyed widget callbacks)
                pass

    def register_theme_callback(self, callback: Callable[[str], None]):
        """Register callback to be called when theme changes"""
        if callback not in self._theme_callbacks:
            self._theme_callbacks.append(callback)

    def unregister_theme_callback(self, callback: Callable[[str], None]):
        """Unregister a theme callback"""
        if callback in self._theme_callbacks:
            self._theme_callbacks.remove(callback)

    def _apply_theme(self, theme: str):
        """Apply theme to all ttk styles"""
        if not self._style:
            return

        colors = AppConstants.THEMES[theme]

        # Update the legacy COLORS dict for backward compatibility
        AppConstants.COLORS.clear()
        AppConstants.COLORS.update(colors)

        # Switch ttkbootstrap theme if available
        if TTKBOOTSTRAP_AVAILABLE and self._ttkb_style:
            ttkb_theme = self.TTKB_THEMES.get(theme, 'cyborg')
            try:
                self._ttkb_style.theme_use(ttkb_theme)
                # Override ttkbootstrap colors to our brand teal
                self._override_ttkbootstrap_colors(colors)
            except Exception:
                pass  # Theme may not be available

        # Configure custom styles (these override/extend ttkbootstrap)
        self._configure_base_styles(colors)
        self._configure_frame_styles(colors)
        self._configure_label_styles(colors)
        self._configure_button_styles(colors)
        self._configure_entry_styles(colors)
        self._configure_combobox_styles(colors)
        self._configure_notebook_styles(colors)
        self._configure_treeview_styles(colors)
        self._configure_progressbar_styles(colors)
        self._configure_scrollbar_styles(colors)
        self._configure_checkbutton_styles(colors)
        self._configure_labelframe_styles(colors)
        self._configure_sidebar_styles(colors)
        self._configure_header_styles(colors)

    def _override_ttkbootstrap_colors(self, colors: dict):
        """Override ttkbootstrap's default colors to our brand teal and secondary"""
        if not TTKBOOTSTRAP_AVAILABLE or not self._ttkb_style:
            return

        try:
            # Access the ttkbootstrap Colors object and override
            style = self._ttkb_style

            # Override primary color in ttkbootstrap's color system
            # This affects all ttkbootstrap widgets that use 'primary' color
            if hasattr(style, 'colors'):
                style.colors.primary = colors['primary']
                style.colors.info = colors['primary']  # Info often uses similar color
                # Override secondary to our card surface colors
                if hasattr(style.colors, 'secondary'):
                    style.colors.secondary = colors['card_surface']

            # Re-configure primary.TButton with our exact teal
            style.configure('primary.TButton',
                           background=colors['primary'])
            style.map('primary.TButton',
                     background=[('pressed', colors['primary_pressed']),
                                ('active', colors['primary_hover']),
                                ('disabled', colors['border'])])

            # Re-configure info.TButton with our teal
            style.configure('info.TButton',
                           background=colors['primary'])
            style.map('info.TButton',
                     background=[('pressed', colors['primary_pressed']),
                                ('active', colors['primary_hover'])])

            # Re-configure secondary.TButton - DARKER on hover, DARKEST on press
            style.configure('secondary.TButton',
                           background=colors['card_surface'],
                           foreground=colors['text_primary'])
            style.map('secondary.TButton',
                     background=[('pressed', colors['card_surface_pressed']),
                                ('active', colors['card_surface_hover']),
                                ('disabled', colors['border'])],
                     foreground=[('pressed', colors['text_primary']),
                                ('active', colors['text_primary'])])

        except Exception:
            pass  # Silently fail if ttkbootstrap structure has changed

    def _configure_base_styles(self, colors: dict):
        """Configure base widget styles"""
        # Only use clam theme if ttkbootstrap is not available
        if not TTKBOOTSTRAP_AVAILABLE:
            self._style.theme_use('clam')

        # Root configuration
        self._style.configure('.',
                              background=colors['background'],
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 10))

    def _configure_frame_styles(self, colors: dict):
        """Configure frame styles"""
        self._style.configure('TFrame',
                              background=colors['background'])

        self._style.configure('Card.TFrame',
                              background=colors['card_surface'])

        self._style.configure('Content.TFrame',
                              background=colors['background'])

    def _configure_label_styles(self, colors: dict):
        """Configure label styles"""
        self._style.configure('TLabel',
                              background=colors['background'],
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 10))

        self._style.configure('Title.TLabel',
                              background=colors['background'],
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 18, 'bold'))

        self._style.configure('Subtitle.TLabel',
                              background=colors['background'],
                              foreground=colors['text_secondary'],
                              font=('Segoe UI', 10))

        self._style.configure('Brand.TLabel',
                              background=colors['background'],
                              foreground=colors['primary'],
                              font=('Segoe UI', 16, 'bold'))

        self._style.configure('Muted.TLabel',
                              background=colors['background'],
                              foreground=colors['text_muted'],
                              font=('Segoe UI', 9))

        self._style.configure('HeaderTitle.TLabel',
                              background=colors['background'],
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 16, 'bold'))

        # Card-specific labels
        self._style.configure('Card.TLabel',
                              background=colors['card_surface'],
                              foreground=colors['text_primary'])

    def _configure_button_styles(self, colors: dict):
        """Configure button styles - flat, modern, no 3D effects"""
        # Primary button (teal) - same color family for all states
        self._style.configure('TButton',
                              background=colors['primary'],
                              foreground='#ffffff',
                              font=('Segoe UI', 10, 'bold'),
                              padding=(16, 10),
                              borderwidth=0,
                              relief='flat',
                              focuscolor='none')

        self._style.map('TButton',
                       background=[('pressed', colors['primary_pressed']),
                                  ('active', colors['primary_hover']),
                                  ('disabled', colors['border'])],
                       foreground=[('disabled', colors['text_muted']),
                                  ('pressed', '#ffffff'),
                                  ('active', '#ffffff'),
                                  ('!disabled', '#ffffff')],
                       relief=[('pressed', 'flat'), ('active', 'flat')])

        # Secondary button (gray) - MUST go darker on hover, darkest on click
        # Configure both our custom style and ttkbootstrap's secondary style
        secondary_bg = colors['card_surface']
        secondary_hover = colors['card_surface_hover']  # Darker than normal
        secondary_pressed = colors['card_surface_pressed']  # Darkest

        # Our custom Secondary.TButton style
        self._style.configure('Secondary.TButton',
                              background=secondary_bg,
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 10),
                              padding=(16, 10),
                              borderwidth=0,
                              relief='flat',
                              focuscolor='none')

        self._style.map('Secondary.TButton',
                       background=[('pressed', '!' + 'disabled', secondary_pressed),
                                  ('active', '!' + 'disabled', secondary_hover),
                                  ('disabled', colors['border'])],
                       foreground=[('pressed', colors['text_primary']),
                                  ('active', colors['text_primary']),
                                  ('!disabled', colors['text_primary'])],
                       relief=[('pressed', 'flat'), ('active', 'flat')])

        # Also override ttkbootstrap's secondary.TButton (lowercase)
        if TTKBOOTSTRAP_AVAILABLE:
            self._style.configure('secondary.TButton',
                                  background=secondary_bg,
                                  foreground=colors['text_primary'])
            self._style.map('secondary.TButton',
                           background=[('pressed', '!' + 'disabled', secondary_pressed),
                                      ('active', '!' + 'disabled', secondary_hover)])

        # Action button (teal CTA) - same color family for all states
        self._style.configure('Action.TButton',
                              background=colors['primary'],
                              foreground='#ffffff',
                              font=('Segoe UI', 10, 'bold'),
                              padding=(24, 12),
                              borderwidth=0,
                              relief='flat',
                              focuscolor='none')

        self._style.map('Action.TButton',
                       background=[('pressed', colors['primary_pressed']),
                                  ('active', colors['primary_hover']),
                                  ('disabled', colors['border'])],
                       foreground=[('disabled', colors['text_muted']),
                                  ('pressed', '#ffffff'),
                                  ('active', '#ffffff'),
                                  ('!disabled', '#ffffff')],
                       relief=[('pressed', 'flat'), ('active', 'flat')])

        # Brand button - accent teal colored, flat
        self._style.configure('Brand.TButton',
                              background=colors['accent'],
                              foreground='#ffffff',
                              font=('Segoe UI', 10, 'bold'),
                              padding=(16, 10),
                              borderwidth=0,
                              relief='flat',
                              focuscolor='none')

        self._style.map('Brand.TButton',
                       background=[('pressed', colors['primary_pressed']),
                                  ('active', colors['primary'])],
                       foreground=[('pressed', '#ffffff'),
                                  ('active', '#ffffff'),
                                  ('!disabled', '#ffffff')],
                       relief=[('pressed', 'flat'), ('active', 'flat')])

        # Info button - blue, smaller, flat
        self._style.configure('Info.TButton',
                              background=colors['info'],
                              foreground='#ffffff',
                              font=('Segoe UI', 9),
                              padding=(14, 8),
                              borderwidth=0,
                              relief='flat',
                              focuscolor='none')

        self._style.map('Info.TButton',
                       background=[('pressed', '#1e40af'),
                                  ('active', '#2563eb')],
                       foreground=[('pressed', '#ffffff'),
                                  ('active', '#ffffff'),
                                  ('!disabled', '#ffffff')],
                       relief=[('pressed', 'flat'), ('active', 'flat')])

        # Header action button (gray) - same color family for all states
        self._style.configure('HeaderAction.TButton',
                              background=colors['card_surface'],
                              foreground=colors['text_secondary'],
                              font=('Segoe UI', 9),
                              padding=(12, 6),
                              borderwidth=0,
                              relief='flat',
                              focuscolor='none')

        self._style.map('HeaderAction.TButton',
                       background=[('pressed', colors['card_surface_pressed']),
                                  ('active', colors['card_surface_hover'])],
                       foreground=[('pressed', colors['text_primary']),
                                  ('active', colors['text_primary'])],
                       relief=[('pressed', 'flat'), ('active', 'flat')])

    def _configure_entry_styles(self, colors: dict):
        """Configure entry styles - clean, flat borders"""
        self._style.configure('TEntry',
                              fieldbackground=colors['surface'],
                              foreground=colors['text_primary'],
                              insertcolor=colors['text_primary'],
                              bordercolor=colors['border'],
                              lightcolor=colors['border'],
                              darkcolor=colors['border'],
                              relief='flat',
                              padding=(10, 8))

        self._style.map('TEntry',
                       fieldbackground=[('focus', colors['surface']),
                                       ('disabled', colors['background'])],
                       bordercolor=[('focus', colors['primary']),
                                   ('!focus', colors['border'])])

        # Section.TEntry - for entries inside Section.TFrame content areas
        # Uses background color to blend seamlessly with content frame
        self._style.configure('Section.TEntry',
                              fieldbackground=colors['background'],
                              foreground=colors['text_primary'],
                              insertcolor=colors['text_primary'],
                              bordercolor=colors['border'],
                              lightcolor=colors['border'],
                              darkcolor=colors['border'],
                              relief='flat',
                              padding=(10, 8))

        self._style.map('Section.TEntry',
                       fieldbackground=[('readonly', colors['background']),
                                       ('focus', colors['background']),
                                       ('disabled', colors['background'])],
                       bordercolor=[('readonly', colors['border']),
                                   ('focus', colors['primary']),
                                   ('!focus', colors['border'])],
                       lightcolor=[('readonly', colors['border']),
                                  ('focus', colors['border']),
                                  ('!focus', colors['border'])],
                       darkcolor=[('readonly', colors['border']),
                                 ('focus', colors['border']),
                                 ('!focus', colors['border'])])

    def _configure_combobox_styles(self, colors: dict):
        """Configure combobox styles - flat, clean, theme-consistent"""
        # Border color: blue in dark mode (#00587C), teal in light mode (#009999)
        border_color = colors['button_primary']

        self._style.configure('TCombobox',
                              fieldbackground=colors['card_surface'],
                              background=colors['card_surface'],
                              foreground=colors['text_primary'],
                              arrowcolor=colors['text_secondary'],
                              bordercolor=border_color,
                              lightcolor=border_color,
                              darkcolor=border_color,
                              insertbackground=colors['text_primary'],  # Cursor visibility
                              relief='flat',
                              padding=(6, 3))  # Compact padding for smaller dropdowns

        # Use consistent border color in ALL states (blue in dark, teal in light)
        self._style.map('TCombobox',
                       fieldbackground=[('readonly', colors['card_surface']),
                                       ('disabled', colors['background'])],
                       foreground=[('readonly', colors['text_primary']),
                                  ('!disabled', colors['text_primary'])],
                       # Dropdown arrow button hover - use brand color for visibility
                       background=[('active', colors['button_primary_hover']),
                                  ('pressed', colors['button_primary_hover']),
                                  ('disabled', colors['background']),
                                  ('!active', colors['card_surface'])],
                       bordercolor=[('focus', border_color), ('!focus', border_color),
                                   ('pressed', border_color), ('active', border_color)],
                       lightcolor=[('focus', border_color), ('!focus', border_color),
                                  ('pressed', border_color), ('active', border_color)],
                       darkcolor=[('focus', border_color), ('!focus', border_color),
                                 ('pressed', border_color), ('active', border_color)])

        # Style the dropdown listbox (popup) - matches card_surface for theme consistency
        # Use button_primary (teal) for brand-consistent selection highlight
        if self._root:
            # Combobox entry cursor (insertion point) color - try multiple patterns
            # Use priority 100 to ensure these take precedence
            self._root.option_add('*TCombobox*Entry.insertBackground', colors['text_primary'], 100)
            self._root.option_add('*TCombobox.insertBackground', colors['text_primary'], 100)
            self._root.option_add('*Combobox.insertBackground', colors['text_primary'], 100)
            self._root.option_add('*Entry.insertBackground', colors['text_primary'], 60)
            # Listbox styling (priority 100 for higher precedence)
            self._root.option_add('*TCombobox*Listbox.background', colors['card_surface'], 100)
            self._root.option_add('*TCombobox*Listbox.foreground', colors['text_primary'], 100)
            self._root.option_add('*TCombobox*Listbox.selectBackground', colors['button_primary'], 100)
            self._root.option_add('*TCombobox*Listbox.selectForeground', '#ffffff', 100)
            self._root.option_add('*TCombobox*Listbox.font', ('Segoe UI', 9), 100)
            # Remove white corner - use 0 thickness and matching colors
            listbox_border = colors['card_surface']  # Match listbox background for seamless look
            self._root.option_add('*TCombobox*Listbox.relief', 'flat', 100)
            self._root.option_add('*TCombobox*Listbox.borderWidth', 0, 100)
            self._root.option_add('*TCombobox*Listbox.highlightThickness', 0, 100)
            self._root.option_add('*TCombobox*Listbox.highlightBackground', listbox_border, 100)
            self._root.option_add('*TCombobox*Listbox.highlightColor', listbox_border, 100)

    def _configure_notebook_styles(self, colors: dict):
        """Configure notebook styles (for tool internal tabs if needed)"""
        self._style.configure('TNotebook',
                              background=colors['background'],
                              borderwidth=0)

        self._style.configure('TNotebook.Tab',
                              background=colors['card_surface'],
                              foreground=colors['text_secondary'],
                              font=('Segoe UI', 10),
                              padding=(15, 8))

        self._style.map('TNotebook.Tab',
                       background=[('selected', colors['primary']),
                                  ('active', colors['accent'])],
                       foreground=[('selected', '#ffffff'),
                                  ('active', '#ffffff')])

    def _configure_treeview_styles(self, colors: dict):
        """Configure treeview styles - using centralized color constants"""
        # Use centralized list/tree colors with fallbacks for compatibility
        list_bg = colors.get('list_bg', colors['surface'])
        tree_heading_bg = colors.get('tree_heading_bg', colors['card_surface'])
        tree_border = colors.get('tree_border', colors['border'])
        selection_bg = colors.get('selection_highlight', colors['primary'])

        self._style.configure('Treeview',
                              background=list_bg,
                              foreground=colors['text_primary'],
                              fieldbackground=list_bg,
                              bordercolor=tree_border,
                              font=('Segoe UI', 10))

        self._style.configure('Treeview.Heading',
                              background=tree_heading_bg,
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 10, 'bold'))
        # Fix for Python 3.13+ where style.configure background isn't rendered
        self._style.map('Treeview.Heading',
                       background=[('', tree_heading_bg)])

        self._style.map('Treeview',
                       background=[('selected', selection_bg)],
                       foreground=[('selected', '#ffffff')])

    def _configure_progressbar_styles(self, colors: dict):
        """Configure progressbar styles - flat, modern"""
        self._style.configure('TProgressbar',
                              background=colors['primary'],
                              troughcolor=colors['surface'],
                              bordercolor=colors['surface'],
                              lightcolor=colors['primary'],
                              darkcolor=colors['primary'],
                              borderwidth=0,
                              thickness=6)

        self._style.configure('Success.TProgressbar',
                              background=colors['success'],
                              lightcolor=colors['success'],
                              darkcolor=colors['success'])

    def _configure_scrollbar_styles(self, colors: dict):
        """Configure scrollbar styles"""
        try:
            self._style.configure('TScrollbar',
                                  background=colors['card_surface'],
                                  troughcolor=colors['background'],
                                  bordercolor=colors['border'],
                                  arrowcolor=colors['text_secondary'])

            self._style.map('TScrollbar',
                           background=[('active', colors['accent']),
                                      ('pressed', colors['primary'])])
        except Exception:
            # ttkbootstrap may have already configured scrollbar elements
            pass

    def _configure_checkbutton_styles(self, colors: dict):
        """Configure checkbutton styles"""
        self._style.configure('TCheckbutton',
                              background=colors['background'],
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 10))

        self._style.map('TCheckbutton',
                       background=[('active', colors['background'])],
                       foreground=[('disabled', colors['text_muted'])])

    def _configure_labelframe_styles(self, colors: dict):
        """Configure labelframe styles - softer borders"""
        # Use section_bg if available, otherwise fall back to background
        section_bg = colors.get('section_bg', colors['background'])
        border_soft = colors.get('border_soft', colors['border'])

        # Standard TLabelframe - softer look
        self._style.configure('TLabelframe',
                              background=section_bg,
                              bordercolor=border_soft,
                              borderwidth=1,
                              relief='flat')

        self._style.configure('TLabelframe.Label',
                              background=section_bg,
                              foreground=colors['text_primary'],
                              font=('Segoe UI', 10, 'bold'))

        # Section.TLabelframe - primary section style (no visible border)
        # Set borderwidth=0 to completely hide the border in dark mode
        self._style.configure('Section.TLabelframe',
                              background=section_bg,
                              bordercolor=section_bg,
                              borderwidth=0,
                              relief='flat',
                              padding=(15, 10))

        # Use title_color for section labels (blue in dark mode, teal in light mode)
        title_color = colors.get('title_color', colors['primary'])
        self._style.configure('Section.TLabelframe.Label',
                              background=section_bg,
                              foreground=title_color,
                              font=('Segoe UI Semibold', 11))  # Semibold - consistent with ui_base

        # Section.TFrame - for content frames inside Section.TLabelframe
        # Uses colors['background'] (white/dark) to create contrast with section_bg border
        # Buttons inside should use canvas_bg=colors['background'] to match
        self._style.configure('Section.TFrame',
                              background=colors['background'])

        # Section.TLabel - for labels inside Section.TFrame (matches background)
        self._style.configure('Section.TLabel',
                              background=colors['background'])

        # AnalysisSummary.TFrame - borderless frame for analysis summary tables
        # Per design guide: Analysis Summary should NOT have a border to keep it separate
        self._style.configure('AnalysisSummary.TFrame',
                              background=section_bg,
                              relief='flat',
                              borderwidth=0)

        # ProgressLog.TFrame - frame with faint border for progress log sections
        # Per design guide: Progress Log gets a faint border and modern scrollbar
        log_border = colors.get('log_border', colors.get('border_soft', colors['border']))
        self._style.configure('ProgressLog.TFrame',
                              background=section_bg,
                              relief='flat')

        # Card.TFrame - for card containers with surface background
        self._style.configure('Card.TFrame',
                              background=colors['card_surface'])

        # Dialog.TFrame - for help dialogs and popups
        dialog_bg = colors.get('dialog_bg', colors['background'])
        self._style.configure('Dialog.TFrame',
                              background=dialog_bg)

    def _configure_sidebar_styles(self, colors: dict):
        """Configure sidebar-specific styles"""
        self._style.configure('Sidebar.TFrame',
                              background=colors['sidebar_bg'])

        self._style.configure('SidebarBrand.TLabel',
                              background=colors['sidebar_bg'],
                              foreground=colors['primary'],
                              font=('Segoe UI', 14, 'bold'))

        self._style.configure('SidebarSubtitle.TLabel',
                              background=colors['sidebar_bg'],
                              foreground=colors['nav_text'],
                              font=('Segoe UI', 9))

        self._style.configure('SidebarMuted.TLabel',
                              background=colors['sidebar_bg'],
                              foreground=colors['text_muted'],
                              font=('Segoe UI', 9))

    def _configure_header_styles(self, colors: dict):
        """Configure header-specific styles"""
        self._style.configure('Header.TFrame',
                              background=colors['background'])

    def _load_theme_preference(self):
        """Load saved theme preference"""
        try:
            if self._settings_path.exists():
                with open(self._settings_path, 'r') as f:
                    settings = json.load(f)
                    saved_theme = settings.get('theme', AppConstants.DEFAULT_THEME)
                    if saved_theme in AppConstants.THEMES:
                        self._current_theme = saved_theme
        except Exception:
            self._current_theme = AppConstants.DEFAULT_THEME

    def _save_theme_preference(self):
        """Save theme preference to file"""
        try:
            self._settings_path.parent.mkdir(parents=True, exist_ok=True)

            # Load existing settings if any
            settings = {}
            if self._settings_path.exists():
                try:
                    with open(self._settings_path, 'r') as f:
                        settings = json.load(f)
                except Exception:
                    pass

            # Update theme
            settings['theme'] = self._current_theme

            with open(self._settings_path, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass


# Convenience function
def get_theme_manager() -> ThemeManager:
    """Get the global theme manager instance"""
    return ThemeManager.get_instance()


# Export
__all__ = ['ThemeManager', 'get_theme_manager']
