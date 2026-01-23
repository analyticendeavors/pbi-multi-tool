"""
Power BI Report Merger - Application Constants
Built by Reid Havens of Analytic Endeavors
"""

class AppConstants:
    """Essential application constants - simplified and optimized."""

    # =============================================================================
    # CORE APPLICATION INFO
    # =============================================================================

    # Company & Branding
    COMPANY_NAME = "Analytic Endeavors"
    COMPANY_FOUNDER = "Reid Havens"
    COMPANY_WEBSITE = "https://www.analyticendeavors.com"

    # Application
    APP_NAME = "Power BI Report Merger"
    APP_VERSION = "v1.0 Enhanced"
    WINDOW_TITLE = f"{COMPANY_NAME} - {APP_NAME}"
    WINDOW_SIZE = "1450x1105"  # Wider layout for laptops, taller for inline theme selection (+5px for Grouped text)
    MIN_WINDOW_SIZE = (1250, 750)
    SIDEBAR_WIDTH = 210
    SIDEBAR_COLLAPSED_WIDTH = 72  # Icons-only mode (wider for square buttons)

    # =============================================================================
    # TOOL ORDER - Controls sidebar display order
    # =============================================================================

    TOOL_ORDER = [
        "report_cleanup",
        "accessibility_checker",
        "pbip_layout_optimizer",
        "column_width",
        "field_parameters",
        "connection_hotswap",
        "advanced_copy",
        "report_merger",
        "sensitivity_scanner",
    ]

    # =============================================================================
    # THEME SYSTEM - DARK AND LIGHT MODES
    # =============================================================================

    DEFAULT_THEME = 'dark'

    THEMES = {
        'dark': {
            # Primary brand colors - teal shades for button states
            'primary': '#009999',        # Teal (brand) - default
            'primary_hover': '#007A7A',  # Darker teal - hover
            'primary_pressed': '#005C5C', # Darkest teal - click
            'secondary': '#00587C',      # Dark Blue (brand)
            'accent': '#33ADAD',         # Light Teal
            'selection_bg': '#1a5a8a',   # Blue for text selection in dark mode (clearly different from teal)
            'accent_deep': '#003D56',    # Deep Blue

            # Interface colors
            'background': '#0d0d1a',     # Main background
            'card_bg': '#0d0d1a',
            'card_surface': '#1a1a2e',   # Card/panel backgrounds - default
            'card_surface_hover': '#141424',  # Hover state (darker)
            'card_surface_pressed': '#0e0e18', # Click state (darkest)
            'section_bg': '#161627',     # Section background (subtle diff from background)
            'surface': '#202036',        # Slightly lighter than card_surface for alternating rows
            'border': '#3d3d5c',
            'border_soft': '#2a2a40',    # Softer border for sections

            # Text colors
            'text_primary': '#ffffff',
            'text_secondary': '#c0c0d0',
            'text_muted': '#808090',

            # Status colors
            'success': '#10b981',
            'warning': '#f5751f',
            'error': '#ef4444',
            'info': '#3b82f6',

            # Risk level colors (for Sensitivity Scanner)
            'risk_high': '#dc2626',
            'risk_medium': '#d97706',
            'risk_low': '#059669',

            # Sidebar specific
            'sidebar_bg': '#0a0a14',
            'sidebar_hover': '#1a1a2e',       # Hover state (more visible)
            'sidebar_pressed': '#252540',     # Pressed state (even more visible)
            'sidebar_active': '#00587C',      # Brand dark blue for active nav items
            'sidebar_active_hover': '#004d6b', # Darker blue - active item hover
            'sidebar_active_pressed': '#003d56', # Darkest blue - active item click
            'nav_text': '#c0c0d0',
            'nav_text_active': '#ffffff',
            'divider_faint': '#2a2a3a',  # Faint divider line (dark mode)

            # Button colors for dark mode - use blue instead of teal
            'button_primary': '#00587C',        # Blue for Browse buttons
            'button_primary_hover': '#004466',  # Darker blue - hover
            'button_primary_pressed': '#003050', # Darkest blue - click
            'button_primary_disabled': '#3a3a4e',  # Muted gray - disabled state
            'button_secondary': '#2a2a40',       # Secondary buttons (Export Log, Clear Log, Reset All)
            'button_secondary_hover': '#222236', # Darker - hover
            'button_secondary_pressed': '#1a1a2a', # Darkest - click
            'button_text_disabled': '#6a6a7a',   # Muted text for disabled buttons

            # Title color - separate from button colors
            'title_color': '#0084b7',           # Light blue for section titles

            # Option panel background (for radio buttons, checkboxes)
            'option_bg': '#161627',              # Match section_bg for seamless integration

            # Canvas/outer backgrounds (for RoundedButton canvas corners)
            'outer_bg': '#161627',               # Outer section canvas background
            'dialog_bg': '#0d0d1a',              # Help/popup dialog background

            # Button text colors
            'button_text': '#ffffff',            # White text on primary buttons
            'button_text_secondary': '#ffffff',  # Text on secondary buttons (same for dark mode)

            # Warning/alert styling
            'warning_bg': '#d97706',             # Orange warning box background
            'warning_text': '#ffffff',           # White text on warning background

            # Disabled/muted states
            'text_disabled': '#94a3b8',          # Grayed out/disabled text

            # List/TreeView styling
            'list_bg': '#1e1e2e',                # List/tree background
            'tree_heading_bg': '#2a2a3e',        # Tree view header background
            'tree_border': '#3a3a4a',            # Tree view border color
            'selection_highlight': '#1a3a5c',    # Selected item background in lists

            # Progress log specific
            'log_border': '#2a2a40',             # Faint border for progress log (matches border_soft)
        },
        'light': {
            # Primary brand colors - teal shades for button states
            'primary': '#009999',        # Teal (brand) - default
            'primary_hover': '#007A7A',  # Darker teal - hover
            'primary_pressed': '#005C5C', # Darkest teal - click
            'secondary': '#00587C',
            'accent': '#33ADAD',
            'selection_bg': '#3b82f6',   # Blue for text selection in light mode (consistent with dark mode)
            'accent_deep': '#003D56',

            # Interface colors
            'background': '#ffffff',     # Main background (white)
            'card_bg': '#f5f5f7',
            'card_surface': '#e8e8f0',   # Default (light gray, not pure white)
            'card_surface_hover': '#d8d8e4',  # Hover state (darker)
            'card_surface_pressed': '#c8c8d8', # Click state (darkest)
            'section_bg': '#f5f5f7',     # Section background (off-white inside sections)
            'surface': '#f5f5f7',
            'border': '#d0d0e0',
            'border_soft': '#e0e0ec',    # Softer border for sections

            # Text colors
            'text_primary': '#1a1a2e',
            'text_secondary': '#5a5a6a',
            'text_muted': '#94a3b8',

            # Status colors
            'success': '#059669',
            'warning': '#f5751f',
            'error': '#dc2626',
            'info': '#2563eb',

            # Risk level colors (for Sensitivity Scanner)
            'risk_high': '#dc2626',
            'risk_medium': '#d97706',
            'risk_low': '#059669',

            # Sidebar specific
            'sidebar_bg': '#e8e8ed',
            'sidebar_hover': '#d0d0dc',       # Hover state (darker)
            'sidebar_pressed': '#c0c0cc',     # Pressed state (even darker)
            'sidebar_active': '#009999',
            'sidebar_active_hover': '#007a7a', # Darker teal - active item hover
            'sidebar_active_pressed': '#005c5c', # Darkest teal - active item click
            'nav_text': '#4a4a5a',
            'nav_text_active': '#ffffff',
            'divider_faint': '#c8c8d0',  # Faint divider line (light mode)

            # Button colors for light mode - keep teal
            'button_primary': '#009999',        # Teal for Browse buttons
            'button_primary_hover': '#007A7A',  # Darker teal - hover
            'button_primary_pressed': '#005C5C', # Darkest teal - click
            'button_primary_disabled': '#c0c0cc',  # Light gray - disabled state
            'button_secondary': '#d8d8e0',       # Secondary buttons
            'button_secondary_hover': '#c8c8d0', # Darker - hover
            'button_secondary_pressed': '#b8b8c0', # Darkest - click
            'button_text_disabled': '#9a9aa8',   # Muted text for disabled buttons

            # Title color - separate from button colors
            'title_color': '#009999',           # Teal for section titles (matches primary in light mode)

            # Option panel background (for radio buttons, checkboxes)
            'option_bg': '#f5f5f7',              # Light gray for radio visibility

            # Canvas/outer backgrounds (for RoundedButton canvas corners)
            'outer_bg': '#f5f5f7',               # Outer section canvas background
            'dialog_bg': '#f5f5f7',              # Help/popup dialog background

            # Button text colors
            'button_text': '#ffffff',            # White text on primary buttons
            'button_text_secondary': '#1a1a2e',  # Dark text on secondary buttons in light mode

            # Warning/alert styling
            'warning_bg': '#d97706',             # Orange warning box background
            'warning_text': '#ffffff',           # White text on warning background

            # Disabled/muted states
            'text_disabled': '#94a3b8',          # Grayed out/disabled text

            # List/TreeView styling
            'list_bg': '#ffffff',                # List/tree background (white)
            'tree_heading_bg': '#f0f0f5',        # Tree view header background
            'tree_border': '#d8d8e0',            # Tree view border color
            'selection_highlight': '#e6f3ff',    # Selected item background in lists

            # Progress log specific
            'log_border': '#e0e0ec',             # Faint border for progress log (matches border_soft)
        }
    }

    # Legacy COLORS for backward compatibility - defaults to dark theme
    # This will be dynamically updated by ThemeManager
    COLORS = THEMES['dark'].copy()

    # =============================================================================
    # CENTRALIZED FONT SYSTEM
    # =============================================================================
    # All fonts use Segoe UI (Windows default) with Helvetica fallback
    # These constants ensure consistent typography across all tools

    FONTS = {
        # Section and dialog headers
        'section_header': ('Segoe UI Semibold', 11),      # Section titles (e.g., "PBIP File Source")
        'dialog_title': ('Segoe UI', 16, 'bold'),         # Help dialog main title
        'dialog_section': ('Segoe UI', 12, 'bold'),       # Help dialog section headers

        # Buttons
        'button': ('Segoe UI', 10, 'bold'),               # Primary action buttons
        'button_secondary': ('Segoe UI', 10),             # Secondary/reset buttons
        'button_small': ('Segoe UI', 9),                  # Small action buttons

        # Labels and body text
        'label': ('Segoe UI', 10),                        # Standard labels
        'label_bold': ('Segoe UI', 10, 'bold'),           # Bold labels/values
        'body': ('Segoe UI', 9),                          # Body text, table cells
        'body_bold': ('Segoe UI', 9, 'bold'),             # Bold body text
        'body_italic': ('Segoe UI', 9, 'italic'),         # Italic tips/hints

        # Small/hint text
        'hint': ('Segoe UI', 8),                          # Very small hint text
        'tip': ('Segoe UI', 9, 'italic'),                 # Tip/guide text

        # Cards and special elements
        'card_title': ('Segoe UI', 10, 'bold'),           # Card titles
        'card_count': ('Segoe UI', 16, 'bold'),           # Large count display on cards
        'card_emoji': ('Segoe UI', 24),                   # Emoji icons on cards

        # Table styling
        'table_header': ('Segoe UI', 9, 'bold'),          # Table column headers
        'table_cell': ('Segoe UI', 9),                    # Table cell content

        # Log/console text
        'log': ('Consolas', 9),                           # Monospace log text
    }

    # =============================================================================
    # UI SIZING CONSTANTS
    # =============================================================================
    # Standard sizes for consistent UI elements across tools

    UI_SIZES = {
        # Button dimensions
        'button_height': 38,                   # Standard button height
        'button_height_small': 26,             # Small buttons (selection controls)
        'button_radius': 6,                    # Standard rounded corner radius
        'button_radius_small': 5,              # Smaller radius for compact buttons

        # Primary action button widths
        'button_width_action': 160,            # Execute/Analyze buttons
        'button_width_browse': 90,             # Browse buttons
        'button_width_reset': 130,             # Reset buttons
        'button_width_small': 58,              # Small action buttons (All/None)

        # Help button
        'help_button_size': 26,                # Square help button size
        'help_button_icon': 14,                # Help icon size
        'help_button_offset_y': -35,           # Y offset for upper-right positioning

        # Section icon sizes
        'section_icon': 16,                    # Icons in section headers
        'card_icon': 24,                       # Icons on cards
    }

    # =============================================================================
    # CORNER ICON BACKGROUND OVERRIDE
    # =============================================================================
    # Background override for corner icons (help, settings, login, etc.) in upper-right area.
    # Makes button "invisible" until hover by matching parent frame background.
    # Uses theme color key 'background' which resolves to #0d0d1a (dark) or #ffffff (light).
    CORNER_ICON_BG = {'dark': 'background', 'light': 'background'}

    @classmethod
    def get_colors(cls, theme: str = None) -> dict:
        """Get colors for specified theme (or current default)"""
        theme = theme or cls.DEFAULT_THEME
        return cls.THEMES.get(theme, cls.THEMES['dark'])
    
    # =============================================================================
    # UI TEXT CONTENT
    # =============================================================================
    
    # Header text
    BRAND_TEXT = f"📊 {COMPANY_NAME.upper()}"
    MAIN_TITLE = APP_NAME
    TAGLINE = "Professional-grade tool for intelligent PBIP report consolidation"
    BUILT_BY_TEXT = f"Built by {COMPANY_FOUNDER} of {COMPANY_NAME}"
    
    # Quick start steps (embedded in UI)
    QUICK_START_STEPS = [
        "1. Navigate to your .pbip file in File Explorer",
        "2. Right-click the .pbip file and select 'Copy as path'", 
        "3. Paste (Ctrl+V) into the path field above",
        "4. Path quotes will be automatically cleaned",
        "5. Repeat for the second report file",
        "6. Click 'Analyze Reports' to begin"
    ]
    
    # =============================================================================
    # TECHNICAL CONFIGURATION
    # =============================================================================
    
    # Power BI schema URLs (only the ones actually used)
    SCHEMA_URLS = {
        'platform': "https://developer.microsoft.com/json-schemas/fabric/item/platformMetadata/1.0.0/schema.json",
        'pbip': "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        'bookmarks': "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmarksMetadata/1.0.0/schema.json",
        'pages': "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        'report_extension': "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/reportExtension/1.0.0/schema.json"
    }
    
    # File settings
    SUPPORTED_EXTENSIONS = ['.pbip']
    MAX_CACHE_SIZE = 100


# =============================================================================
# CENTRALIZED ERROR MESSAGES
# =============================================================================
# Common error messages and dialog titles used across multiple tools.
# Use these constants to ensure consistent wording and enable easy updates.
# Usage: from core.constants import ErrorMessages
#        ErrorMessages.NO_SELECTION  # "No Selection"

class ErrorMessages:
    """Centralized error message strings for consistent UI messaging."""

    # Dialog titles (used as first parameter in messagebox calls)
    NO_SELECTION = "No Selection"
    INVALID_INPUT = "Invalid Input"
    INVALID_POSITION = "Invalid Position"
    FILE_NOT_FOUND = "File Not Found"
    CONNECTION_ERROR = "Connection Error"
    VALIDATION_ERROR = "Validation Error"
    OPERATION_FAILED = "Operation Failed"
    SUCCESS = "Success"
    INFO = "Information"
    WARNING = "Warning"

    # Common message bodies
    NO_FILE_SELECTED = "Please select a file first."
    NO_ITEM_SELECTED = "Please select an item first."
    NO_CONNECTION_SELECTED = "Please select a connection first."
    INVALID_NUMBER = "Please enter a valid number."
    INVALID_PORT = "Please enter a valid port number (1024-65535)"

    # File-related messages (use .format(path=...) for dynamic paths)
    PBIP_NOT_FOUND = "PBIP file not found: {path}"
    FILE_DOES_NOT_EXIST = "The file does not exist:\n{path}"
    FOLDER_NOT_FOUND = "Folder does not exist: {path}"

    # Operation messages
    OPERATION_COMPLETE = "Operation completed successfully."
    CHANGES_SAVED = "Changes saved successfully."
    SETTINGS_SAVED = "Settings saved successfully."
    RESET_COMPLETE = "Reset to defaults complete."


# Export main constants classes
__all__ = ['AppConstants', 'ErrorMessages']
