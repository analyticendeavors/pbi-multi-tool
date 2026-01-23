"""
Template Widgets - Reusable UI Section Components
Built by Reid Havens of Analytic Endeavors

Provides composable building blocks for tool tabs:
- ActionButtonBar: Standard action/reset button pair
- FileInputSection: File selection with browse and action buttons
- SplitLogSection: Two-column summary + progress log layout
"""

import tkinter as tk
from tkinter import ttk, filedialog
import io
from pathlib import Path
from typing import Optional, Callable, List

from core.constants import AppConstants
from core.theme_manager import get_theme_manager

# Import button components
from core.ui.buttons import RoundedButton, SquareIconButton

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


class ActionButtonBar(tk.Frame):
    """
    Standard action button bar with primary and optional secondary buttons.

    Provides consistent button styling and layout across all tools.
    Buttons are packed side=LEFT with standard spacing.

    Usage:
        button_bar = ActionButtonBar(
            parent=self.frame,
            theme_manager=self._theme_manager,
            primary_text="EXECUTE MERGE",
            primary_command=self.start_merge,
            primary_icon=execute_icon,
            secondary_text="RESET ALL",
            secondary_command=self.reset_tab,
            secondary_icon=reset_icon
        )
        button_bar.pack(side=tk.BOTTOM, pady=(15, 0))

        # Access buttons directly
        button_bar.primary_button.set_enabled(False)
        button_bar.secondary_button.set_enabled(True)
    """

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager,
        primary_text: str,
        primary_command: Callable,
        primary_icon=None,
        secondary_text: str = "RESET ALL",
        secondary_command: Callable = None,
        secondary_icon=None,
        primary_starts_disabled: bool = False,
        button_spacing: int = 15,
        **kwargs
    ):
        """
        Args:
            parent: Parent widget
            theme_manager: ThemeManager instance for colors
            primary_text: Text for primary (action) button
            primary_command: Callback for primary button
            primary_icon: Optional icon for primary button
            secondary_text: Text for secondary button (default "RESET ALL")
            secondary_command: Callback for secondary button (None to hide)
            secondary_icon: Optional icon for secondary button
            primary_starts_disabled: If True, primary button starts disabled
            button_spacing: Pixels between buttons (default 15)
        """
        # Get canvas background from theme (section_bg for bottom buttons)
        colors = theme_manager.colors
        canvas_bg = colors['section_bg']

        super().__init__(parent, bg=canvas_bg, **kwargs)

        self._theme_manager = theme_manager
        self._button_spacing = button_spacing
        self._primary_icon = primary_icon
        self._secondary_icon = secondary_icon

        # Get theme colors
        is_dark = theme_manager.is_dark
        bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        # Create primary button
        self._primary_button = RoundedButton(
            self,
            text=primary_text,
            command=primary_command,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            disabled_bg=bg_disabled,
            disabled_fg=fg_disabled,
            height=38,
            radius=6,
            font=('Segoe UI', 10, 'bold'),
            icon=primary_icon,
            canvas_bg=canvas_bg
        )
        self._primary_button.pack(side=tk.LEFT, padx=(0, button_spacing))

        if primary_starts_disabled:
            self._primary_button.set_enabled(False)

        # Create secondary button if command provided
        self._secondary_button = None
        if secondary_command is not None:
            self._secondary_button = RoundedButton(
                self,
                text=secondary_text,
                command=secondary_command,
                bg=colors['button_secondary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                fg=colors['text_primary'],
                height=38,
                radius=6,
                font=('Segoe UI', 10),
                icon=secondary_icon,
                canvas_bg=canvas_bg
            )
            self._secondary_button.pack(side=tk.LEFT)

        # Register for theme updates
        theme_manager.register_theme_callback(self._on_theme_changed)

    @property
    def primary_button(self) -> RoundedButton:
        """Access the primary action button"""
        return self._primary_button

    @property
    def secondary_button(self) -> Optional[RoundedButton]:
        """Access the secondary button (may be None)"""
        return self._secondary_button

    def set_primary_enabled(self, enabled: bool):
        """Enable or disable the primary button"""
        self._primary_button.set_enabled(enabled)

    def set_secondary_enabled(self, enabled: bool):
        """Enable or disable the secondary button"""
        if self._secondary_button:
            self._secondary_button.set_enabled(enabled)

    def _on_theme_changed(self, theme: str):
        """Update button colors when theme changes"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        canvas_bg = colors['section_bg']

        # Update frame background
        self.configure(bg=canvas_bg)

        # Update primary button
        bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

        self._primary_button.update_colors(
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            disabled_bg=bg_disabled,
            disabled_fg=fg_disabled,
            canvas_bg=canvas_bg
        )

        # Update secondary button
        if self._secondary_button:
            self._secondary_button.update_colors(
                bg=colors['button_secondary'],
                hover_bg=colors['button_secondary_hover'],
                pressed_bg=colors['button_secondary_pressed'],
                fg=colors['text_primary'],
                canvas_bg=canvas_bg
            )

    def destroy(self):
        """Unregister from theme manager before destroying"""
        try:
            self._theme_manager.unregister_theme_callback(self._on_theme_changed)
        except Exception:
            pass
        super().destroy()


class FileInputSection(tk.Frame):
    """
    Complete file input section with styled LabelFrame, entry, browse button,
    and optional action button. Handles theme changes automatically.

    This is a composable building block that encapsulates the common pattern
    of file selection used across multiple tools.

    Usage:
        file_section = FileInputSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="PBIP File Source",
            section_icon="Power-BI",
            file_label="Project File (PBIP):",
            file_types=[("PBIP files", "*.pbip")],
            action_button_text="ANALYZE REPORT",
            action_button_command=self.analyze,
            action_button_icon=analyze_icon,
            on_file_selected=self.validate_file
        )
        file_section.pack(fill=tk.X, pady=(0, 15))

        # Access path
        path = file_section.path
        file_section.path = "/new/path.pbip"

        # Control action button
        file_section.set_action_enabled(True)
    """

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager,
        section_title: str,
        section_icon: str,
        file_label: str = "Project File:",
        file_types: List[tuple] = None,
        action_button_text: str = None,
        action_button_command: Callable = None,
        action_button_icon=None,
        help_command: Callable = None,
        on_file_selected: Callable = None,
        browse_button_text: str = "Browse",
        **kwargs
    ):
        """
        Args:
            parent: Parent widget
            theme_manager: ThemeManager instance for colors
            section_title: Title text for the section header (e.g., "PBIP File Source")
            section_icon: Icon name for section header (e.g., "Power-BI", "folder")
            file_label: Label text for file input (e.g., "Project File (PBIP):")
            file_types: List of (description, pattern) tuples for file dialog
            action_button_text: Text for action button (None to hide)
            action_button_command: Callback for action button
            action_button_icon: Optional icon for action button
            help_command: Callback for help button (None to hide help button)
            on_file_selected: Callback when file is selected (receives path string)
            browse_button_text: Text for browse button (default "Browse")
        """
        super().__init__(parent, **kwargs)

        self._theme_manager = theme_manager
        self._file_types = file_types or [("All Files", "*.*")]
        self._on_file_selected = on_file_selected
        self._section_icon = section_icon

        # Path variable with trace for button enable/disable and file selection callback
        self._path_var = tk.StringVar()
        # Always set up trace to handle action button enable/disable
        self._path_var.trace('w', lambda *args: self._handle_path_change())

        # Get colors
        colors = theme_manager.colors
        is_dark = theme_manager.is_dark

        # Create section header using the pattern from BaseToolTab
        self._header_frame, self._header_icon, self._header_label = \
            self._create_section_header(section_title, section_icon)

        # Create LabelFrame with header
        self._section_frame = ttk.LabelFrame(
            self, labelwidget=self._header_frame,
            style='Section.TLabelframe', padding="12"
        )
        self._section_frame.pack(fill=tk.X, expand=True)

        # Content frame with padding
        self._content_frame = ttk.Frame(self._section_frame, style='Section.TFrame', padding="15")
        self._content_frame.pack(fill=tk.BOTH, expand=True)
        self._content_frame.columnconfigure(1, weight=1)  # Entry expands

        # Optional help button in upper right corner
        self._help_button = None
        if help_command:
            help_icon = self._load_icon("question", size=14)
            self._help_button = SquareIconButton(
                self._section_frame, icon=help_icon, command=help_command,
                tooltip_text="Help", size=26, radius=6,
                bg_normal_override=AppConstants.CORNER_ICON_BG
            )
            self._help_button.place(relx=1.0, y=-35, anchor=tk.NE, x=0)

        # File input row - use tk.Frame for explicit bg (needed for button corner rounding)
        # Uses colors['background'] to match Section.TFrame content area
        input_row_bg = colors['background']
        self._input_row = tk.Frame(self._content_frame, bg=input_row_bg)
        self._input_row.pack(fill=tk.X)
        self._input_row.columnconfigure(1, weight=1)

        # Label
        self._file_label = tk.Label(
            self._input_row, text=file_label, bg=input_row_bg,
            fg=colors['text_primary'], font=('Segoe UI', 10)
        )
        self._file_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        # Entry
        self._path_entry = ttk.Entry(
            self._input_row, textvariable=self._path_var,
            font=('Segoe UI', 10), style='Section.TEntry'
        )
        self._path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        # Browse button
        folder_icon = self._load_icon("folder", size=16)
        self._browse_button = RoundedButton(
            self._input_row,
            text=browse_button_text if folder_icon else f"📁 {browse_button_text}",
            command=self._browse_file,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            width=90, height=32, radius=6,
            font=('Segoe UI', 10),
            icon=folder_icon,
            canvas_bg=input_row_bg
        )
        self._browse_button.grid(row=0, column=2)

        # Action button (optional)
        self._action_button = None
        self._action_button_frame = None
        if action_button_text and action_button_command:
            # Wrapper frame for centering
            self._action_button_frame = tk.Frame(self._content_frame, bg=input_row_bg)
            self._action_button_frame.pack(fill=tk.X, pady=(15, 0))

            bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
            fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')

            self._action_button = RoundedButton(
                self._action_button_frame,
                text=action_button_text,
                command=action_button_command,
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg='#ffffff',
                disabled_bg=bg_disabled,
                disabled_fg=fg_disabled,
                height=38, radius=6,
                font=('Segoe UI', 10, 'bold'),
                icon=action_button_icon,
                canvas_bg=input_row_bg
            )
            self._action_button.pack()
            # Start disabled until file is selected
            self._action_button.set_enabled(False)

        # Store input_row_bg for theme updates
        self._input_row_bg = input_row_bg

        # Register for theme updates
        theme_manager.register_theme_callback(self._on_theme_changed)

    def _create_section_header(self, text: str, icon_name: str):
        """Create section header with icon + text for LabelFrame labelwidget"""
        colors = self._theme_manager.colors
        # Use section_bg to match Section.TLabelframe.Label style
        header_bg = colors.get('section_bg', colors['background'])

        header_frame = tk.Frame(self, bg=header_bg)

        # Try to load icon
        icon_label = None
        icon = self._load_icon(icon_name, size=16)
        if icon:
            icon_label = tk.Label(header_frame, image=icon, bg=header_bg)
            icon_label.pack(side=tk.LEFT, padx=(0, 6))
            icon_label._icon_ref = icon  # Prevent garbage collection

        # Text label
        text_label = tk.Label(
            header_frame, text=text,
            font=('Segoe UI Semibold', 11),
            bg=header_bg,
            fg=colors.get('title_color', colors.get('accent', '#00a8a8'))
        )
        text_label.pack(side=tk.LEFT)

        return header_frame, icon_label, text_label

    def _load_icon(self, name: str, size: int = 16):
        """Load an SVG icon from assets"""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return None

        try:
            # Path from core/ui/ up to src, then into assets
            icon_path = Path(__file__).parent.parent.parent / "assets" / "Tool Icons" / f"{name}.svg"
            if not icon_path.exists():
                return None

            png_data = cairosvg.svg2png(url=str(icon_path), output_width=size, output_height=size)
            image = Image.open(io.BytesIO(png_data))
            return ImageTk.PhotoImage(image)
        except Exception:
            return None

    def _browse_file(self):
        """Open file browser dialog"""
        file_path = filedialog.askopenfilename(
            title="Select File",
            filetypes=self._file_types + [("All Files", "*.*")]
        )
        if file_path:
            self._path_var.set(file_path)

    def _handle_path_change(self):
        """Handle path variable changes"""
        path = self._path_var.get().strip()

        # Enable/disable action button based on path
        if self._action_button:
            self._action_button.set_enabled(bool(path))

        # Call user callback
        if self._on_file_selected:
            self._on_file_selected(path)

    @property
    def path(self) -> str:
        """Get the current file path"""
        return self._path_var.get().strip()

    @path.setter
    def path(self, value: str):
        """Set the file path"""
        self._path_var.set(value)

    @property
    def path_var(self) -> tk.StringVar:
        """Get the path StringVar for external binding"""
        return self._path_var

    @property
    def action_button(self) -> Optional[RoundedButton]:
        """Get the action button (may be None)"""
        return self._action_button

    @property
    def browse_button(self) -> RoundedButton:
        """Get the browse button"""
        return self._browse_button

    @property
    def section_frame(self) -> ttk.LabelFrame:
        """Get the section LabelFrame for placing additional widgets"""
        return self._section_frame

    @property
    def content_frame(self) -> ttk.Frame:
        """Get the content frame for adding extra widgets"""
        return self._content_frame

    def set_action_enabled(self, enabled: bool):
        """Enable or disable the action button"""
        if self._action_button:
            self._action_button.set_enabled(enabled)

    def _on_theme_changed(self, theme: str):
        """Update colors when theme changes"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        # Header uses section_bg to match Section.TLabelframe.Label style
        header_bg = colors.get('section_bg', colors['background'])
        # Input row uses background to match Section.TFrame content area
        input_row_bg = colors['background']

        # Update header frame
        self._header_frame.configure(bg=header_bg)
        if self._header_icon:
            self._header_icon.configure(bg=header_bg)
        if self._header_label:
            self._header_label.configure(
                bg=header_bg,
                fg=colors.get('title_color', colors.get('accent', '#00a8a8'))
            )

        # Update input row
        self._input_row.configure(bg=input_row_bg)
        self._file_label.configure(bg=input_row_bg, fg=colors['text_primary'])

        # Update browse button
        self._browse_button.update_colors(
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            canvas_bg=input_row_bg
        )

        # Update action button frame and button
        if self._action_button_frame:
            self._action_button_frame.configure(bg=input_row_bg)

        if self._action_button:
            bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
            fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')
            self._action_button.update_colors(
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg='#ffffff',
                disabled_bg=bg_disabled,
                disabled_fg=fg_disabled,
                canvas_bg=input_row_bg
            )

    def destroy(self):
        """Unregister from theme manager before destroying"""
        try:
            self._theme_manager.unregister_theme_callback(self._on_theme_changed)
        except Exception:
            pass
        super().destroy()


class SplitLogSection(tk.Frame):
    """
    Split log section with summary panel (left) and progress log (right).

    This composable widget encapsulates the common pattern of a two-column
    log section with Analysis Summary on the left and Progress Log on the right.
    Includes export and clear buttons for the progress log.

    Usage:
        log_section = SplitLogSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="Analysis & Progress",
            summary_title="Analysis Summary",
            summary_icon="bar-chart",
            log_title="Progress Log",
            summary_placeholder="Run analysis to see results",
            on_export=self.export_log,  # Optional custom export handler
            on_clear=self.clear_log     # Optional custom clear handler
        )
        log_section.pack(fill=tk.BOTH, expand=True)

        # Log messages
        log_section.log("Processing file...")
        log_section.log("Complete!", "success")

        # Set summary content
        log_section.set_summary("Found 5 issues\\n- Issue 1\\n- Issue 2")

        # Clear log
        log_section.clear_log()

        # Access underlying widgets
        log_section.log_text  # ModernScrolledText for log
        log_section.summary_text  # ModernScrolledText for summary
    """

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager,
        section_title: str = "Analysis & Progress",
        section_icon: str = "analyze",
        summary_title: str = "Analysis Summary",
        summary_icon: str = "bar-chart",
        log_title: str = "Progress Log",
        log_icon: str = "log-file",
        summary_placeholder: str = "Run analysis to see results",
        on_export: Callable = None,
        on_clear: Callable = None,
        **kwargs
    ):
        """
        Args:
            parent: Parent widget
            theme_manager: ThemeManager instance for colors
            section_title: Title for the outer LabelFrame
            section_icon: Icon name for section header
            summary_title: Title for the summary panel header
            summary_icon: Icon name for summary header
            log_title: Title for the progress log header
            log_icon: Icon name for log header
            summary_placeholder: Placeholder text when summary is empty
            on_export: Optional custom export handler (default: save to file)
            on_clear: Optional custom clear handler (default: clear log text)
        """
        colors = theme_manager.colors
        super().__init__(parent, bg=colors['section_bg'], **kwargs)

        self._theme_manager = theme_manager
        self._summary_placeholder_text = summary_placeholder
        self._on_export_callback = on_export
        self._on_clear_callback = on_clear
        self._button_icons = {}

        # Configure column weights for the outer frame
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Create section header - use section_bg to match Section.TLabelframe.Label style
        header_bg = colors.get('section_bg', colors['background'])
        self._header_frame, self._header_icon, self._header_label = self._create_section_header(
            section_icon, section_title, header_bg, colors
        )

        # Main LabelFrame
        self._section_frame = ttk.LabelFrame(
            self,
            labelwidget=self._header_frame,
            style='Section.TLabelframe',
            padding="12"
        )
        self._section_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._section_frame.columnconfigure(0, weight=1)
        self._section_frame.rowconfigure(0, weight=1)

        # Inner content frame - white/card background
        self._content_frame = ttk.Frame(self._section_frame, style='Section.TFrame', padding="15")
        self._content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Use uniform group to ensure columns stay equal width regardless of content
        self._content_frame.columnconfigure(0, weight=1, minsize=225, uniform="split_cols")
        self._content_frame.columnconfigure(1, weight=1, minsize=225, uniform="split_cols")
        self._content_frame.rowconfigure(0, weight=1)

        # Create left side (summary)
        self._create_summary_panel(summary_title, summary_icon, colors)

        # Create right side (progress log)
        self._create_log_panel(log_title, log_icon, colors)

        # Register for theme updates
        theme_manager.register_theme_callback(self._on_theme_changed)

    def _create_section_header(self, icon_name: str, text: str, bg_color: str, colors: dict):
        """Create the section header with icon and text."""
        header_frame = tk.Frame(self, bg=bg_color)

        icon_label = None
        try:
            # Try to load the icon
            icon_path = Path(__file__).parent.parent.parent / 'assets' / 'Tool Icons' / f'{icon_name}.svg'
            if icon_path.exists():
                import cairosvg
                from PIL import Image, ImageTk
                import io

                png_data = cairosvg.svg2png(url=str(icon_path), output_width=18, output_height=18)
                image = Image.open(io.BytesIO(png_data))
                photo = ImageTk.PhotoImage(image)

                icon_label = tk.Label(header_frame, image=photo, bg=bg_color)
                icon_label.image = photo
                icon_label.pack(side=tk.LEFT, padx=(0, 8))
        except Exception:
            pass

        text_label = tk.Label(
            header_frame,
            text=text,
            font=('Segoe UI Semibold', 12),
            bg=bg_color,
            fg=colors.get('title_color', colors.get('accent', '#00a8a8'))
        )
        text_label.pack(side=tk.LEFT)

        return header_frame, icon_label, text_label

    def _create_icon_label(self, parent: tk.Widget, icon_name: str, text: str,
                           bg_color: str, colors: dict, icon_size: int = 16,
                           font: tuple = ('Segoe UI Semibold', 11)):
        """Create an icon + text label for sub-headers."""
        frame = tk.Frame(parent, bg=bg_color)

        icon_label = None
        try:
            icon_path = Path(__file__).parent.parent.parent / 'assets' / 'Tool Icons' / f'{icon_name}.svg'
            if icon_path.exists():
                import cairosvg
                from PIL import Image, ImageTk
                import io

                png_data = cairosvg.svg2png(url=str(icon_path), output_width=icon_size, output_height=icon_size)
                image = Image.open(io.BytesIO(png_data))
                photo = ImageTk.PhotoImage(image)

                icon_label = tk.Label(frame, image=photo, bg=bg_color)
                icon_label.image = photo
                icon_label.pack(side=tk.LEFT, padx=(0, 6))
        except Exception:
            pass

        text_label = tk.Label(
            frame,
            text=text,
            font=font,
            bg=bg_color,
            fg=colors.get('title_color', colors.get('accent', '#00a8a8'))
        )
        text_label.pack(side=tk.LEFT)

        return frame, icon_label, text_label

    def _create_summary_panel(self, title: str, icon_name: str, colors: dict):
        """Create the left summary panel."""
        # Deferred import to avoid circular dependency
        from core.ui_base import ModernScrolledText

        # Summary container
        self._summary_container = ttk.Frame(self._content_frame, style='Section.TFrame')
        self._summary_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        self._summary_container.columnconfigure(0, weight=1)
        self._summary_container.rowconfigure(0, minsize=30)
        self._summary_container.rowconfigure(1, weight=1)

        # Summary header frame (stretches full width for optional filter controls on right)
        header_bg = colors['background']
        self._summary_header_frame, self._summary_header_icon, self._summary_header_label = \
            self._create_icon_label(self._summary_container, icon_name, title, header_bg, colors)
        self._summary_header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))

        # Summary content frame with background
        self._summary_frame = tk.Frame(
            self._summary_container,
            bg=colors['section_bg'],
            highlightthickness=0,
            padx=8,
            pady=8
        )
        self._summary_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._summary_frame.columnconfigure(0, weight=1)
        self._summary_frame.rowconfigure(0, weight=1)

        # Placeholder label
        self._placeholder_label = tk.Label(
            self._summary_frame,
            text=self._summary_placeholder_text,
            font=('Segoe UI', 10, 'italic'),
            bg=colors['section_bg'],
            fg=colors.get('text_secondary', colors['text_primary']),
            anchor=tk.CENTER
        )
        self._placeholder_label.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        # Summary text widget (hidden until content is set)
        self._summary_text = ModernScrolledText(
            self._summary_frame,
            height=10,
            width=40,
            font=('Segoe UI', 9),
            state=tk.DISABLED,
            bg=colors['section_bg'],
            fg=colors['text_primary'],
            selectbackground=colors['selection_bg'],
            selectforeground='#ffffff',
            relief='flat',
            borderwidth=0,
            wrap=tk.WORD,
            highlightthickness=0,
            padx=5,
            pady=5,
            theme_manager=self._theme_manager,
            auto_hide_scrollbar=True
        )
        # Not gridded yet - shown when placeholder is hidden

    def _create_log_panel(self, title: str, icon_name: str, colors: dict):
        """Create the right progress log panel."""
        # Deferred import to avoid circular dependency
        from core.ui_base import ModernScrolledText

        # Log outer container
        self._log_container = ttk.Frame(self._content_frame, style='Section.TFrame')
        self._log_container.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._log_container.columnconfigure(0, weight=1)  # Header expands
        self._log_container.columnconfigure(1, weight=0)  # Buttons fixed width
        self._log_container.rowconfigure(0, minsize=30)
        self._log_container.rowconfigure(1, weight=1)

        # Log header label - direct child of container (matches summary pattern)
        header_bg = colors['background']
        self._log_header_frame, self._log_header_icon, self._log_header_label = \
            self._create_icon_label(self._log_container, icon_name, title, header_bg, colors)
        self._log_header_frame.grid(row=0, column=0, sticky=tk.W, pady=(0, 8))

        # Icon buttons frame - direct child of container
        self._icon_buttons_frame = tk.Frame(self._log_container, bg=header_bg)
        self._icon_buttons_frame.grid(row=0, column=1, sticky=tk.E, pady=(0, 8))

        # Load icons
        self._button_icons['save'] = self._load_icon("save", 14)
        self._button_icons['eraser'] = self._load_icon("eraser", 14)

        # Export button
        self._export_button = SquareIconButton(
            self._icon_buttons_frame,
            icon=self._button_icons['save'],
            command=self._handle_export,
            tooltip_text="Export Log",
            size=26,
            radius=6
        )
        self._export_button.pack(side=tk.LEFT, padx=(0, 4))

        # Clear button
        self._clear_button = SquareIconButton(
            self._icon_buttons_frame,
            icon=self._button_icons['eraser'],
            command=self._handle_clear,
            tooltip_text="Clear Log",
            size=26,
            radius=6
        )
        self._clear_button.pack(side=tk.LEFT)

        # Log text container - span both columns
        log_text_container = ttk.Frame(self._log_container, style='Section.TFrame')
        log_text_container.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_text_container.columnconfigure(0, weight=1)
        log_text_container.rowconfigure(0, weight=1)

        # Log text widget
        self._log_text = ModernScrolledText(
            log_text_container,
            height=10,
            width=45,
            font=('Cascadia Code', 9),
            state=tk.DISABLED,
            bg=colors['section_bg'],
            fg=colors['text_primary'],
            selectbackground=colors['selection_bg'],
            selectforeground='#ffffff',
            relief='flat',
            borderwidth=0,
            wrap=tk.WORD,
            highlightthickness=1,
            highlightcolor=colors['border'],
            highlightbackground=colors['border'],
            padx=5,
            pady=5,
            theme_manager=self._theme_manager,
            auto_hide_scrollbar=False
        )
        self._log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def _load_icon(self, name: str, size: int):
        """Load an icon by name."""
        try:
            icon_path = Path(__file__).parent.parent.parent / 'assets' / 'Tool Icons' / f'{name}.svg'
            if icon_path.exists():
                import cairosvg
                from PIL import Image, ImageTk
                import io

                png_data = cairosvg.svg2png(url=str(icon_path), output_width=size, output_height=size)
                image = Image.open(io.BytesIO(png_data))
                return ImageTk.PhotoImage(image)
        except Exception:
            pass
        return None

    def _handle_export(self):
        """Handle export button click."""
        if self._on_export_callback:
            self._on_export_callback()
        else:
            # Default export behavior
            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                title="Export Log"
            )
            if file_path:
                content = self._log_text.get("1.0", tk.END)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)

    def _handle_clear(self):
        """Handle clear button click."""
        if self._on_clear_callback:
            self._on_clear_callback()
        else:
            self.clear_log()

    # Public API

    @property
    def log_text(self):
        """Access the log text widget."""
        return self._log_text

    @property
    def summary_text(self):
        """Access the summary text widget."""
        return self._summary_text

    @property
    def summary_frame(self) -> tk.Frame:
        """Access the summary frame for custom content."""
        return self._summary_frame

    @property
    def section_frame(self) -> ttk.LabelFrame:
        """Access the outer LabelFrame."""
        return self._section_frame

    @property
    def export_button(self):
        """Access the export button."""
        return self._export_button

    @property
    def clear_button(self):
        """Access the clear button."""
        return self._clear_button

    @property
    def placeholder_label(self) -> tk.Label:
        """Access the placeholder label."""
        return self._placeholder_label

    @property
    def summary_header_frame(self) -> tk.Frame:
        """Access the summary header frame for adding custom controls (e.g., filters)."""
        return self._summary_header_frame

    def log(self, message: str, level: str = "info"):
        """
        Append a message to the progress log.

        Args:
            message: The message to log
            level: Log level (info, success, warning, error) - for future coloring
        """
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] {message}\n"

        self._log_text.configure(state=tk.NORMAL)
        self._log_text.insert(tk.END, formatted)
        self._log_text.see(tk.END)
        self._log_text.configure(state=tk.DISABLED)

    def clear_log(self):
        """Clear the progress log."""
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.delete("1.0", tk.END)
        self._log_text.configure(state=tk.DISABLED)

    def set_summary(self, content: str):
        """
        Set the summary panel content.

        Hides the placeholder and shows the summary text widget with content.

        Args:
            content: The summary text to display
        """
        # Hide placeholder, show text widget
        self._placeholder_label.grid_remove()
        self._summary_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Set content
        self._summary_text.configure(state=tk.NORMAL)
        self._summary_text.delete("1.0", tk.END)
        self._summary_text.insert("1.0", content)
        self._summary_text.configure(state=tk.DISABLED)

    def clear_summary(self):
        """Clear the summary and show placeholder."""
        self._summary_text.grid_remove()
        self._summary_text.configure(state=tk.NORMAL)
        self._summary_text.delete("1.0", tk.END)
        self._summary_text.configure(state=tk.DISABLED)
        self._placeholder_label.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

    def _on_theme_changed(self, theme: str):
        """Update colors when theme changes."""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        # Header uses section_bg to match Section.TLabelframe.Label style
        header_bg = colors.get('section_bg', colors['background'])
        inner_bg = colors['background']

        # Update outer frame
        self.configure(bg=colors['section_bg'])

        # Update section header
        self._header_frame.configure(bg=header_bg)
        if self._header_icon:
            self._header_icon.configure(bg=header_bg)
        if self._header_label:
            self._header_label.configure(
                bg=header_bg,
                fg=colors.get('title_color', colors.get('accent', '#00a8a8'))
            )

        # Update summary header
        self._summary_header_frame.configure(bg=inner_bg)
        if self._summary_header_icon:
            self._summary_header_icon.configure(bg=inner_bg)
        if self._summary_header_label:
            self._summary_header_label.configure(
                bg=inner_bg,
                fg=colors.get('title_color', colors.get('accent', '#00a8a8'))
            )

        # Update summary frame and placeholder (placeholder may be destroyed by tools
        # that replace it with custom content, so check winfo_exists first)
        try:
            self._summary_frame.configure(bg=colors['section_bg'])
        except Exception:
            pass
        try:
            if self._placeholder_label.winfo_exists():
                self._placeholder_label.configure(
                    bg=colors['section_bg'],
                    fg=colors.get('text_secondary', colors['text_primary'])
                )
        except Exception:
            pass

        # Update log header (same pattern as summary header - no wrapper)
        self._log_header_frame.configure(bg=inner_bg)
        if self._log_header_icon:
            self._log_header_icon.configure(bg=inner_bg)
        if self._log_header_label:
            self._log_header_label.configure(
                bg=inner_bg,
                fg=colors.get('title_color', colors.get('accent', '#00a8a8'))
            )

        # Update icon buttons frame
        self._icon_buttons_frame.configure(bg=inner_bg)

        # Explicitly update button canvas backgrounds to ensure correct color
        # (SquareIconButton callbacks may run before parent frame is updated)
        self._export_button.config(bg=inner_bg)
        self._clear_button.config(bg=inner_bg)

        # Update log text widget borders
        self._log_text.configure(
            highlightcolor=colors['border'],
            highlightbackground=colors['border']
        )

    def destroy(self):
        """Unregister from theme manager before destroying."""
        try:
            self._theme_manager.unregister_theme_callback(self._on_theme_changed)
        except Exception:
            pass
        super().destroy()
