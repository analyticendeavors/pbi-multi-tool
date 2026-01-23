"""
Pattern Manager UI - Manage sensitivity detection rules

Full CRUD functionality with Simple/Advanced modes for pattern management.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import re
import sys
import os
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, ThemedScrollbar, ThemedMessageBox, SquareIconButton

# Optional imports for icon loading
try:
    from PIL import Image, ImageTk
    import io
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False


class PatternManager:
    """Pattern Manager window for editing sensitivity patterns."""

    def __init__(self, parent, pattern_detector, on_patterns_updated_callback):
        """Initialize the Pattern Manager."""
        self.parent = parent
        self.pattern_detector = pattern_detector
        self.on_patterns_updated = on_patterns_updated_callback

        # Theme support
        self._theme_manager = get_theme_manager()

        # Track widgets for theme updates
        self._radio_rows = {'mode': [], 'risk': []}  # SVG radio row data
        self._form_buttons = []
        self._bottom_buttons = []
        self._text_widgets = []

        # Radio icons
        self._radio_on_icon = None
        self._radio_off_icon = None

        # Checkbox icons
        self._checkbox_icons = {}
        self._checkbox_labels = {}  # Track checkbox labels for updates

        # Delete icon for header button
        self._trash_icon = None

        # Base path for assets
        if getattr(sys, 'frozen', False):
            self.base_path = Path(sys._MEIPASS)
        else:
            self.base_path = Path(__file__).parent.parent.parent.parent

        # Paths - Handle both development and standalone .exe scenarios
        self.original_patterns_file = self._get_bundled_patterns_path()
        self.custom_patterns_file = self._get_writable_patterns_path()

        # Load radio icons
        self._load_radio_icons()

        # Load checkbox icons
        self._load_checkbox_icons()

        # Load trash icon for delete button
        self._load_trash_icon()

        # Load current patterns
        self.patterns = self._load_current_patterns()
        self.selected_pattern = None

        # Create window
        self._create_window()

        # Register for theme changes
        self._theme_manager.register_theme_callback(self.on_theme_changed)
    
    def _get_bundled_patterns_path(self) -> Path:
        """
        Get path to bundled default patterns.
        
        Handles both development (source files) and PyInstaller (.exe) scenarios.
        PyInstaller extracts data to a temporary folder accessible via sys._MEIPASS.
        """
        if getattr(sys, 'frozen', False):
            # Running in PyInstaller bundle
            base_path = Path(sys._MEIPASS)
        else:
            # Running in development
            base_path = Path(__file__).parent.parent.parent.parent
        
        return base_path / "data" / "sensitivity_patterns.json"
    
    def _get_writable_patterns_path(self) -> Path:
        """
        Get path to writable custom patterns file.
        
        Uses AppData for standalone .exe to ensure write permissions.
        In development, uses the data directory for convenience.
        """
        if getattr(sys, 'frozen', False):
            # Running in PyInstaller bundle - use AppData
            appdata = Path(os.environ.get('APPDATA', Path.home() / 'AppData' / 'Roaming'))
            app_dir = appdata / 'AE Power BI Multi-Tool' / 'Sensitivity Scanner'
            app_dir.mkdir(parents=True, exist_ok=True)
            return app_dir / 'sensitivity_patterns_custom.json'
        else:
            # Running in development - use data directory
            return Path(__file__).parent.parent.parent.parent / "data" / "sensitivity_patterns_custom.json"

    def _load_radio_icons(self):
        """Load SVG radio button icons."""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return

        icons_dir = self.base_path / "assets" / "Tool Icons"
        size = 16

        for icon_name in ['radio-on', 'radio-off']:
            svg_path = icons_dir / f"{icon_name}.svg"
            if svg_path.exists():
                try:
                    # Render at 4x size for quality, then downscale
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=size * 4,
                        output_height=size * 4
                    )
                    img = Image.open(io.BytesIO(png_data))
                    img = img.resize((size, size), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)

                    if icon_name == 'radio-on':
                        self._radio_on_icon = photo
                    else:
                        self._radio_off_icon = photo
                except Exception:
                    pass

    def _load_checkbox_icons(self):
        """Load SVG checkbox icons for checked and unchecked states."""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return

        icons_dir = self.base_path / "assets" / "Tool Icons"
        size = 16
        is_dark = self._theme_manager.is_dark

        # Use theme-appropriate icons
        unchecked_name = "box-dark" if is_dark else "box"
        checked_name = "box-checked-dark" if is_dark else "box-checked"

        for key, icon_name in [('unchecked', unchecked_name), ('checked', checked_name)]:
            svg_path = icons_dir / f"{icon_name}.svg"
            if svg_path.exists():
                try:
                    png_data = cairosvg.svg2png(
                        url=str(svg_path),
                        output_width=size * 4,
                        output_height=size * 4
                    )
                    img = Image.open(io.BytesIO(png_data))
                    img = img.resize((size, size), Image.Resampling.LANCZOS)
                    self._checkbox_icons[key] = ImageTk.PhotoImage(img)
                except Exception:
                    pass

    def _load_trash_icon(self):
        """Load SVG trash icon for delete button."""
        if not PIL_AVAILABLE or not CAIROSVG_AVAILABLE:
            return

        icons_dir = self.base_path / "assets" / "Tool Icons"
        svg_path = icons_dir / "trash.svg"
        size = 16

        if svg_path.exists():
            try:
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=size * 4,
                    output_height=size * 4
                )
                img = Image.open(io.BytesIO(png_data))
                img = img.resize((size, size), Image.Resampling.LANCZOS)
                self._trash_icon = ImageTk.PhotoImage(img)
            except Exception:
                pass

    def _set_title_bar_color(self, dark_mode: bool = True):
        """Set Windows title bar to dark or light mode."""
        try:
            import ctypes
            self.window.update()  # Ensure window is created
            hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if dark_mode else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass  # Ignore on non-Windows or if API fails

    def _set_window_icon(self):
        """Set the AE favicon as the window icon."""
        try:
            # Use AE favicon for Rule Manager
            favicon_path = self.base_path / "assets" / "favicon.ico"
            if favicon_path.exists():
                self.window.iconbitmap(str(favicon_path))
        except Exception:
            pass

    def _create_radio_row(self, parent, text: str, value: str,
                          var: tk.StringVar, group: str, fg_color=None) -> dict:
        """Create a single radio row with SVG icon and text label.

        Args:
            parent: Parent widget
            text: Label text
            value: Value when selected
            var: StringVar bound to this radio group
            group: Group name ('mode' or 'risk') for updating
            fg_color: Optional foreground color (for risk level coloring)

        Returns:
            dict with frame, icon_label, text_label, value, var
        """
        colors = self._theme_manager.colors
        bg_color = colors['background']  # Use popup background, not section_bg

        # Row frame
        row_frame = tk.Frame(parent, bg=bg_color)
        row_frame.pack(side=tk.LEFT, padx=(0, 15))

        # Radio button icon (clickable)
        is_selected = var.get() == value
        icon = self._radio_on_icon if is_selected else self._radio_off_icon

        icon_label = tk.Label(row_frame, bg=bg_color, cursor='hand2')
        if icon:
            icon_label.configure(image=icon)
            icon_label._icon_ref = icon
        else:
            # Fallback to text if icons not available
            icon_label.configure(text="●" if is_selected else "○", font=('Segoe UI', 10))
        icon_label.pack(side=tk.LEFT, padx=(0, 4))

        # Determine text color
        if fg_color:
            text_fg = fg_color
        else:
            text_fg = colors['title_color'] if is_selected else colors['text_primary']

        # Text label
        text_label = tk.Label(row_frame, text=text, bg=bg_color, fg=text_fg,
                              font=('Segoe UI', 9), cursor='hand2', anchor='w')
        text_label.pack(side=tk.LEFT)

        # Store fg_color for theme updates
        row_data = {
            'frame': row_frame,
            'icon_label': icon_label,
            'text_label': text_label,
            'value': value,
            'var': var,
            'group': group,
            'fg_color': fg_color  # None for mode, risk color for risk level
        }

        # Bind clicks to toggle selection
        def on_click(event=None):
            var.set(value)
            self._update_radio_rows(group)
            # Trigger mode change callback if applicable
            if group == 'mode':
                self._on_mode_change()

        icon_label.bind('<Button-1>', on_click)
        text_label.bind('<Button-1>', on_click)
        row_frame.bind('<Button-1>', on_click)

        # Add hover underline effect for all radio button text
        def on_enter(event, lbl=text_label):
            current_font = lbl.cget('font')
            if 'underline' not in current_font:
                lbl.configure(font=('Segoe UI', 9, 'underline'))
        def on_leave(event, lbl=text_label):
            lbl.configure(font=('Segoe UI', 9))
        text_label.bind('<Enter>', on_enter)
        text_label.bind('<Leave>', on_leave)

        return row_data

    def _update_radio_rows(self, group: str):
        """Update all radio rows in a group when selection changes."""
        colors = self._theme_manager.colors
        bg_color = colors['background']  # Use popup background, not section_bg

        if group not in self._radio_rows:
            return

        for row_data in self._radio_rows[group]:
            is_selected = row_data['var'].get() == row_data['value']

            # Update icon
            icon = self._radio_on_icon if is_selected else self._radio_off_icon
            if icon:
                row_data['icon_label'].configure(image=icon)
                row_data['icon_label']._icon_ref = icon
            else:
                row_data['icon_label'].configure(text="●" if is_selected else "○")

            # Update text color - get fresh risk colors from theme on theme change
            if group == 'risk':
                # Map value to risk color key
                risk_color_map = {
                    'high_risk': colors['risk_high'],
                    'medium_risk': colors['risk_medium'],
                    'low_risk': colors['risk_low']
                }
                text_fg = risk_color_map.get(row_data['value'], colors['text_primary'])
                # Update stored fg_color with current theme color
                row_data['fg_color'] = text_fg
            elif row_data.get('fg_color'):
                text_fg = row_data['fg_color']
            else:
                text_fg = colors['title_color'] if is_selected else colors['text_primary']
            row_data['text_label'].configure(fg=text_fg)

            # Update backgrounds
            row_data['frame'].configure(bg=bg_color)
            row_data['icon_label'].configure(bg=bg_color)
            row_data['text_label'].configure(bg=bg_color)

    def _load_current_patterns(self) -> Dict[str, Any]:
        """Load current patterns."""
        try:
            if self.custom_patterns_file.exists():
                with open(self.custom_patterns_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            with open(self.original_patterns_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            ThemedMessageBox.showerror(self.parent, "Error", f"Failed to load patterns: {e}")
            return {"patterns": {"high_risk": [], "medium_risk": [], "low_risk": []}}
    
    def _create_window(self):
        """Create the pattern manager window."""
        colors = self._theme_manager.colors

        self.window = tk.Toplevel(self.parent)
        self.window.title("Rule Manager - Sensitivity Scanner")
        self.window.transient(self.parent.winfo_toplevel())
        self.window.grab_set()

        # Set sensitivity scanner icon for this tool's popup
        self._set_window_icon()

        # Use background (white/dark) for main window - NOT section_bg (off-white)
        self.window.configure(bg=colors['background'])

        # Handle window close to unregister theme callback
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

        # Center window BEFORE showing (prevents flash)
        self.window.withdraw()  # Hide temporarily
        self.window.update_idletasks()
        root = self.parent.winfo_toplevel()
        x = root.winfo_rootx() + (root.winfo_width() - 1010) // 2
        y = root.winfo_rooty() + (root.winfo_height() - 805) // 2
        self.window.geometry(f"970x805+{x}+{y}")

        # Set dark/light title bar BEFORE showing to prevent white flash
        self._set_title_bar_color(self._theme_manager.is_dark)

        self.window.deiconify()  # Show at correct position

        # Configure style for the main container with matching background
        style = ttk.Style()
        style.configure('PatternManager.TFrame', background=colors['background'])
        style.configure('PatternManager.TLabelframe', background=colors['background'])
        style.configure('PatternManager.TLabelframe.Label', background=colors['background'],
                        foreground=colors['text_primary'], font=('Segoe UI', 10, 'bold'))

        # Main container
        container = ttk.Frame(self.window, padding="20", style='PatternManager.TFrame')
        container.pack(fill=tk.BOTH, expand=True)
        self._container = container  # Store for theme updates

        # Popup background color for all labels
        popup_bg = colors['background']
        self._popup_bg = popup_bg  # Store for child methods

        # Track labels for theme updates
        self._popup_labels = []  # Regular labels
        self._hint_labels = []   # Muted hint text labels
        self._title_labels = []  # Title-colored labels

        # Content (split view) - Fill remaining space
        content = ttk.Frame(container, style='PatternManager.TFrame')
        content.pack(fill=tk.BOTH, expand=True)
        content.columnconfigure(0, weight=0, minsize=425)  # Fixed width for pattern list
        content.columnconfigure(1, weight=1)  # Editor can expand
        content.rowconfigure(0, weight=1)  # Make row expandable
        
        # Left: Pattern list
        self._create_pattern_list(content)
        
        # Right: Editor
        self._create_editor(content)
        
        # Bottom buttons
        self._create_buttons(container)
    
    def _create_pattern_list(self, parent):
        """Create pattern list view."""
        colors = self._theme_manager.colors
        popup_bg = colors['background']

        # Outer container for the pattern list section
        list_container = tk.Frame(parent, bg=popup_bg)
        list_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=(0, 15))
        list_container.rowconfigure(1, weight=1)  # Make table row expandable
        list_container.columnconfigure(0, weight=1)

        # Header row: title/count on left, trash icon on right
        count = sum(len(self.patterns['patterns'].get(r, [])) for r in ['high_risk', 'medium_risk', 'low_risk'])
        is_custom = self.custom_patterns_file.exists()
        header_frame = tk.Frame(list_container, bg=popup_bg)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        header_frame.columnconfigure(0, weight=1)  # Let left side expand

        # Left side: title and count
        left_frame = tk.Frame(header_frame, bg=popup_bg)
        left_frame.grid(row=0, column=0, sticky=tk.W)

        header_title = tk.Label(left_frame, text="Current Patterns",
                                 font=('Segoe UI', 10, 'bold'), bg=popup_bg, fg=colors['text_primary'])
        header_title.pack(side=tk.LEFT)
        pattern_status = "custom" if is_custom else "defaults"
        self.count_label = tk.Label(left_frame, text=f"  ({count} patterns - {pattern_status})",
                                     font=('Segoe UI', 9), bg=popup_bg, fg=colors['text_muted'])
        self.count_label.pack(side=tk.LEFT)

        # Right side: trash icon
        self._header_delete_btn = SquareIconButton(
            header_frame,
            icon=self._trash_icon,
            command=self._delete_selected_patterns,
            size=26,
            tooltip_text="Delete selected patterns"
        )
        self._header_delete_btn.grid(row=0, column=1, sticky=tk.E)

        self._header_frame = header_frame  # Store for theme updates
        self._left_frame = left_frame  # Store for theme updates
        self._header_title = header_title  # Store for theme updates
        self._popup_labels.append(header_title)
        self._hint_labels.append(self.count_label)
        self._list_outer_frame = list_container  # Store for theme updates

        # Single border frame around tree + scrollbar (the accent border)
        border_frame = tk.Frame(list_container, bg=colors['border'], padx=1, pady=1)
        border_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        border_frame.rowconfigure(0, weight=1)
        border_frame.columnconfigure(0, weight=1)
        self._tree_border_frame = border_frame  # Store for theme updates

        tree_container = tk.Frame(border_frame, bg=colors['section_bg'])
        tree_container.pack(fill=tk.BOTH, expand=True)
        self._tree_container_frame = tree_container  # Store for theme updates

        # Configure modern treeview style with groove headers (like scan results table)
        style = ttk.Style()
        is_dark = self._theme_manager.is_dark
        if is_dark:
            heading_bg = '#2a2a3c'
            heading_fg = '#e0e0e0'
            header_separator = '#0d0d1a'  # Faint column separator for dark mode
        else:
            heading_bg = '#f0f0f0'
            heading_fg = '#333333'
            header_separator = '#ffffff'  # Faint column separator for light mode

        style.configure("Flat.Treeview",
                        borderwidth=0,
                        relief="flat",
                        rowheight=25,
                        background=colors['section_bg'],
                        fieldbackground=colors['section_bg'],
                        lightcolor=colors['section_bg'],
                        darkcolor=colors['section_bg'],
                        bordercolor=colors['section_bg'])

        # Remove treeview frame border (like scan results)
        style.layout("Flat.Treeview", [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])

        # Heading style with groove relief for column dividers (like scan results)
        style.configure("Flat.Treeview.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        font=('Segoe UI', 9, 'bold'))
        style.map("Flat.Treeview.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])

        # Configure selection to only change background, preserve foreground (tag colors)
        # This keeps risk color visible when rows are selected
        selection_bg = colors.get('card_surface', colors.get('surface', '#3A3A3A' if is_dark else '#E0E0E0'))
        style.map("Flat.Treeview",
                  background=[('selected', selection_bg)])
        # Note: NOT mapping foreground for 'selected' state preserves tag colors

        self.tree = ttk.Treeview(tree_container, columns=('id', 'name', 'risk'),
                                 show='headings', style="Flat.Treeview",
                                 selectmode=tk.EXTENDED)  # Enable multi-select

        # All headers centered
        self.tree.heading('id', text='Pattern ID', anchor=tk.CENTER)
        self.tree.heading('name', text='Pattern Name', anchor=tk.CENTER)
        self.tree.heading('risk', text='Risk', anchor=tk.CENTER)

        # Configure columns - name stretches to fill available space
        self.tree.column('id', width=120, minwidth=100, stretch=False)
        self.tree.column('name', width=200, minwidth=150, stretch=True)
        self.tree.column('risk', width=70, minwidth=60, stretch=False, anchor=tk.CENTER)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ThemedScrollbar (no padding between tree and scrollbar)
        self._tree_scrollbar = ThemedScrollbar(
            tree_container,
            command=self.tree.yview,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=self._tree_scrollbar.set)

        self.tree.bind('<<TreeviewSelect>>', self._on_select)

        self._populate_tree()
    
    def _create_editor(self, parent):
        """Create editor form."""
        colors = self._theme_manager.colors
        popup_bg = colors['background']

        # Outer container for the editor section (matches pattern list structure)
        editor_container = tk.Frame(parent, bg=popup_bg)
        editor_container.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.E, tk.W))
        editor_container.rowconfigure(1, weight=1)  # Make content row expandable
        editor_container.columnconfigure(0, weight=1)

        # Header row with "Pattern Editor" title (aligned with "Current Patterns")
        editor_header = tk.Label(editor_container, text="Pattern Editor",
                                  font=('Segoe UI', 10, 'bold'), bg=popup_bg, fg=colors['text_primary'])
        editor_header.grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        self._editor_header = editor_header  # Store for theme updates
        self._popup_labels.append(editor_header)

        # Content frame with padding (replaces LabelFrame padding)
        editor_frame = tk.Frame(editor_container, bg=popup_bg)
        editor_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        editor_frame.rowconfigure(0, weight=1)  # Make canvas expandable
        editor_frame.columnconfigure(0, weight=1)
        self._editor_outer_frame = editor_container  # Store for theme updates

        # Scrollable form with ThemedScrollbar
        self._editor_canvas = tk.Canvas(editor_frame, bg=popup_bg, highlightthickness=0)
        form = ttk.Frame(self._editor_canvas, style='PatternManager.TFrame')

        form.bind('<Configure>', lambda e: self._editor_canvas.configure(scrollregion=self._editor_canvas.bbox('all')))
        self._editor_canvas.create_window((0, 0), window=form, anchor='nw')

        # ThemedScrollbar (no padding between canvas and scrollbar)
        self._editor_scrollbar = ThemedScrollbar(
            editor_frame,
            command=self._editor_canvas.yview,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._editor_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._editor_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._editor_canvas.configure(yscrollcommand=self._editor_scrollbar.set)
        
        # Enable mouse wheel scrolling for editor - bind to all widgets
        def _on_mousewheel(event):
            self._editor_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(widget):
            """Recursively bind mouse wheel to widget and all its children."""
            widget.bind("<MouseWheel>", _on_mousewheel)
            for child in widget.winfo_children():
                _bind_to_mousewheel(child)
        
        # Bind mousewheel to editor frame and all children
        _bind_to_mousewheel(editor_frame)
        self._editor_canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # Form fields
        self.form_vars = {}
        row = 0
        popup_bg = colors['background']

        # Pattern Mode Toggle
        mode_frame = tk.Frame(form, bg=popup_bg)
        mode_frame.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        mode_label = tk.Label(mode_frame, text="Pattern Mode:", font=('Segoe UI', 10, 'bold'),
                              bg=popup_bg, fg=colors['text_primary'])
        mode_label.pack(side=tk.LEFT, padx=(0, 10))
        self._popup_labels.append(mode_label)
        self._mode_frame = mode_frame  # Store for theme updates
        self.form_vars['mode'] = tk.StringVar(value='simple')

        # Create SVG radio buttons for mode selection
        simple_row = self._create_radio_row(mode_frame, "Simple", "simple",
                                            self.form_vars['mode'], 'mode')
        self._radio_rows['mode'].append(simple_row)

        advanced_row = self._create_radio_row(mode_frame, "Advanced", "advanced",
                                              self.form_vars['mode'], 'mode')
        self._radio_rows['mode'].append(advanced_row)
        row += 1
        
        # Simple Mode Section
        self.simple_frame = ttk.Frame(form, style='PatternManager.TFrame')
        self.simple_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        self._create_simple_mode_ui(self.simple_frame)
        row += 1

        # Advanced Mode Section
        self.advanced_frame = ttk.Frame(form, style='PatternManager.TFrame')
        self.advanced_frame.grid(row=row, column=0, sticky=(tk.W, tk.E))
        self._create_advanced_mode_ui(self.advanced_frame)
        self.advanced_frame.grid_remove()  # Hide by default
        row += 1
        
        # Pattern ID (always visible)
        id_label = tk.Label(form, text="Pattern ID:*", font=('Segoe UI', 10, 'bold'),
                           bg=popup_bg, fg=colors['text_primary'])
        id_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(id_label)
        row += 1
        self.form_vars['id'] = tk.StringVar()
        ttk.Entry(form, textvariable=self.form_vars['id']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        id_hint = tk.Label(form, text="(Unique identifier: lowercase, underscores, no spaces. E.g., 'ssn_us')",
                 font=('Segoe UI', 8, 'italic'), bg=popup_bg, fg=colors['text_muted'])
        id_hint.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._hint_labels.append(id_hint)
        row += 1

        # Pattern Name (always visible)
        name_label = tk.Label(form, text="Pattern Name:*", font=('Segoe UI', 10, 'bold'),
                             bg=popup_bg, fg=colors['text_primary'])
        name_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(name_label)
        row += 1
        self.form_vars['name'] = tk.StringVar()
        ttk.Entry(form, textvariable=self.form_vars['name']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        name_hint = tk.Label(form, text="(Human-readable display name. E.g., 'US Social Security Number')",
                 font=('Segoe UI', 8, 'italic'), bg=popup_bg, fg=colors['text_muted'])
        name_hint.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._hint_labels.append(name_hint)
        row += 1

        # Risk Level
        risk_label = tk.Label(form, text="Risk Level:*", font=('Segoe UI', 10, 'bold'),
                             bg=popup_bg, fg=colors['text_primary'])
        risk_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(risk_label)
        row += 1
        self.form_vars['risk_level'] = tk.StringVar(value='medium_risk')
        risk_frame = tk.Frame(form, bg=popup_bg)
        risk_frame.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._risk_frame = risk_frame  # Store for theme updates

        # Create SVG radio buttons for risk level
        high_row = self._create_radio_row(risk_frame, "High", "high_risk",
                                          self.form_vars['risk_level'], 'risk',
                                          fg_color=colors['risk_high'])
        self._radio_rows['risk'].append(high_row)

        medium_row = self._create_radio_row(risk_frame, "Medium", "medium_risk",
                                            self.form_vars['risk_level'], 'risk',
                                            fg_color=colors['risk_medium'])
        self._radio_rows['risk'].append(medium_row)

        low_row = self._create_radio_row(risk_frame, "Low", "low_risk",
                                         self.form_vars['risk_level'], 'risk',
                                         fg_color=colors['risk_low'])
        self._radio_rows['risk'].append(low_row)
        row += 1

        # Description
        desc_label = tk.Label(form, text="Description:*", font=('Segoe UI', 10, 'bold'),
                             bg=popup_bg, fg=colors['text_primary'])
        desc_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(desc_label)
        row += 1
        self.form_vars['description'] = tk.Text(
            form, height=3, width=50, wrap=tk.WORD,
            font=('Segoe UI', 9), relief='flat', bd=0,
            highlightthickness=1, highlightbackground=colors['border'],
            highlightcolor=colors['primary'],
            bg=colors['background'], fg=colors['text_primary']
        )
        self.form_vars['description'].grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self._text_widgets.append(self.form_vars['description'])
        row += 1
        desc_hint = tk.Label(form, text="(What this pattern detects and why it matters)",
                 font=('Segoe UI', 8, 'italic'), bg=popup_bg, fg=colors['text_muted'])
        desc_hint.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._hint_labels.append(desc_hint)
        row += 1

        # Recommended Action
        action_label = tk.Label(form, text="Recommended Action:", font=('Segoe UI', 10, 'bold'),
                               bg=popup_bg, fg=colors['text_primary'])
        action_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(action_label)
        row += 1
        self.form_vars['recommendation'] = tk.Text(
            form, height=3, width=50, wrap=tk.WORD,
            font=('Segoe UI', 9), relief='flat', bd=0,
            highlightthickness=1, highlightbackground=colors['border'],
            highlightcolor=colors['primary'],
            bg=colors['background'], fg=colors['text_primary']
        )
        self.form_vars['recommendation'].grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self._text_widgets.append(self.form_vars['recommendation'])
        row += 1
        action_hint = tk.Label(form, text="(What action should users take when this is found? E.g., 'Implement RLS', 'Use parameters')",
                 font=('Segoe UI', 8, 'italic'), bg=popup_bg, fg=colors['text_muted'])
        action_hint.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._hint_labels.append(action_hint)
        row += 1

        # Examples
        examples_label = tk.Label(form, text="Examples (one per line):", font=('Segoe UI', 10, 'bold'),
                                 bg=popup_bg, fg=colors['text_primary'])
        examples_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(examples_label)
        row += 1
        self.form_vars['examples'] = tk.Text(
            form, height=3, width=50, wrap=tk.WORD,
            font=('Segoe UI', 9), relief='flat', bd=0,
            highlightthickness=1, highlightbackground=colors['border'],
            highlightcolor=colors['primary'],
            bg=colors['background'], fg=colors['text_primary']
        )
        self.form_vars['examples'].grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        self._text_widgets.append(self.form_vars['examples'])
        row += 1
        
        # Form buttons with RoundedButton - use popup background, not section_bg
        btn_frame = tk.Frame(form, bg=popup_bg)
        btn_frame.grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        self._form_btn_frame = btn_frame  # Store for theme updates

        # Add New button - primary style
        self.add_btn = RoundedButton(
            btn_frame, text="Add New", command=self._add_pattern,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=30, radius=6, font=('Segoe UI', 9, 'bold'),
            canvas_bg=popup_bg
        )
        self.add_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._form_buttons.append(self.add_btn)

        # Save button - primary style
        self.save_btn = RoundedButton(
            btn_frame, text="Save", command=self._save_pattern,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=30, radius=6, font=('Segoe UI', 9, 'bold'),
            canvas_bg=popup_bg
        )
        self.save_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._form_buttons.append(self.save_btn)

        # Delete button - secondary style
        self.delete_btn = RoundedButton(
            btn_frame, text="Delete", command=self._delete_pattern,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=30, radius=6, font=('Segoe UI', 9),
            canvas_bg=popup_bg
        )
        self.delete_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._form_buttons.append(self.delete_btn)

        # Clear button - secondary style
        self.clear_btn = RoundedButton(
            btn_frame, text="Clear", command=self._clear_form,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=30, radius=6, font=('Segoe UI', 9),
            canvas_bg=popup_bg
        )
        self.clear_btn.pack(side=tk.LEFT)
        self._form_buttons.append(self.clear_btn)
    
    def _create_simple_mode_ui(self, parent):
        """Create simple mode UI."""
        colors = self._theme_manager.colors
        popup_bg = colors['background']

        # Configure parent column to have fixed width
        parent.columnconfigure(0, weight=1, minsize=450)

        row = 0

        # Pattern Type
        type_label = tk.Label(parent, text="Pattern Type:*", font=('Segoe UI', 10, 'bold'),
                             bg=popup_bg, fg=colors['text_primary'])
        type_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(type_label)
        row += 1
        self.form_vars['pattern_type'] = tk.StringVar(value='Contains Text')
        pattern_types = [
            'Contains Text',
            'Starts With',
            'Ends With',
            'Email Address',
            'Phone Number (US)',
            'Credit Card',
            'URL/Link',
            'IP Address',
            'Date (MM/DD/YYYY)',
            'Date (DD/MM/YYYY)',
            'Date (YYYY-MM-DD)',
            'Custom Date Pattern',
        ]
        type_combo = ttk.Combobox(parent, textvariable=self.form_vars['pattern_type'],
                                 values=pattern_types, state='readonly', width=30)
        type_combo.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        type_combo.bind('<<ComboboxSelected>>', lambda e: self._update_simple_pattern())
        row += 1

        # Search Text
        search_label = tk.Label(parent, text="Search Text:", font=('Segoe UI', 10, 'bold'),
                               bg=popup_bg, fg=colors['text_primary'])
        search_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(search_label)
        row += 1
        self.form_vars['search_text'] = tk.StringVar()
        self.form_vars['search_text'].trace_add('write', lambda *args: self._update_simple_pattern())
        self.search_text_entry = ttk.Entry(parent, textvariable=self.form_vars['search_text'])
        self.search_text_entry.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        self.search_text_helper = tk.Label(parent, text="(Templates with '*' don't require search text - will be auto-generated)",
                 font=('Segoe UI', 8, 'italic'), bg=popup_bg, fg=colors['text_muted'])
        self.search_text_helper.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._hint_labels.append(self.search_text_helper)
        row += 1

        # Options
        options_label = tk.Label(parent, text="Options:", font=('Segoe UI', 10, 'bold'),
                                bg=popup_bg, fg=colors['text_primary'])
        options_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(options_label)
        row += 1

        # Use background color (white/dark) for options area - NOT section_bg
        options_bg = colors['background']

        # Use tk.Frame with explicit bg color instead of ttk.Frame
        self._options_frame = tk.Frame(parent, bg=options_bg)
        self._options_frame.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))

        self.form_vars['case_sensitive'] = tk.BooleanVar(value=False)
        self.form_vars['whole_word'] = tk.BooleanVar(value=False)

        # SVG checkbox for Case Sensitive
        case_icon = self._checkbox_icons.get('unchecked')
        self._case_checkbox = tk.Label(self._options_frame, image=case_icon, cursor='hand2',
                                       bg=options_bg)
        self._case_checkbox.pack(side=tk.LEFT, padx=(0, 4))
        if case_icon:
            self._case_checkbox._icon_ref = case_icon
        self._case_checkbox.bind('<Button-1>', lambda e: self._toggle_checkbox('case_sensitive'))
        self._checkbox_labels['case_sensitive'] = self._case_checkbox

        self._case_text = tk.Label(self._options_frame, text="Case sensitive", font=('Segoe UI', 9),
                            fg=colors['text_primary'], bg=options_bg, cursor='hand2')
        self._case_text.pack(side=tk.LEFT, padx=(0, 15))
        self._case_text.bind('<Button-1>', lambda e: self._toggle_checkbox('case_sensitive'))
        # Hover underline effect
        self._case_text.bind('<Enter>', lambda e: self._case_text.configure(font=('Segoe UI', 9, 'underline')))
        self._case_text.bind('<Leave>', lambda e: self._case_text.configure(font=('Segoe UI', 9)))

        # SVG checkbox for Match Whole Word
        word_icon = self._checkbox_icons.get('unchecked')
        self._word_checkbox = tk.Label(self._options_frame, image=word_icon, cursor='hand2',
                                       bg=options_bg)
        self._word_checkbox.pack(side=tk.LEFT, padx=(0, 4))
        if word_icon:
            self._word_checkbox._icon_ref = word_icon
        self._word_checkbox.bind('<Button-1>', lambda e: self._toggle_checkbox('whole_word'))
        self._checkbox_labels['whole_word'] = self._word_checkbox

        self._word_text = tk.Label(self._options_frame, text="Match whole word only", font=('Segoe UI', 9),
                            fg=colors['text_primary'], bg=options_bg, cursor='hand2')
        self._word_text.pack(side=tk.LEFT)
        self._word_text.bind('<Button-1>', lambda e: self._toggle_checkbox('whole_word'))
        # Hover underline effect
        self._word_text.bind('<Enter>', lambda e: self._word_text.configure(font=('Segoe UI', 9, 'underline')))
        self._word_text.bind('<Leave>', lambda e: self._word_text.configure(font=('Segoe UI', 9)))
        row += 1
        
        # Generated Pattern Preview
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(5, 10))
        row += 1
        gen_label = tk.Label(parent, text="Generated Pattern:", font=('Segoe UI', 9, 'italic'),
                            bg=popup_bg, fg=colors['text_primary'])
        gen_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(gen_label)
        row += 1
        self.preview_label = tk.Label(parent, text="", font=('Consolas', 9),
                                      bg=popup_bg, fg='#0066cc', wraplength=450)
        self.preview_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._popup_labels.append(self.preview_label)
        row += 1

        # Pattern Tester
        test_label = tk.Label(parent, text="Test Your Pattern:", font=('Segoe UI', 10, 'bold'),
                             bg=popup_bg, fg=colors['text_primary'])
        test_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(test_label)
        row += 1
        self.form_vars['test_input'] = tk.StringVar()
        self.form_vars['test_input'].trace_add('write', lambda *args: self._test_pattern())
        ttk.Entry(parent, textvariable=self.form_vars['test_input']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1

        self.test_result_label = tk.Label(parent, text="", font=('Segoe UI', 9),
                                         bg=popup_bg, fg=colors['text_primary'], wraplength=450)
        self.test_result_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._popup_labels.append(self.test_result_label)
        row += 1
    
    def _create_advanced_mode_ui(self, parent):
        """Create advanced mode UI."""
        colors = self._theme_manager.colors
        popup_bg = colors['background']

        # Configure parent column to have fixed width (same as simple mode)
        parent.columnconfigure(0, weight=1, minsize=450)

        row = 0

        # Regex Pattern
        regex_label = tk.Label(parent, text="Regex Pattern:*", font=('Segoe UI', 10, 'bold'),
                              bg=popup_bg, fg=colors['text_primary'])
        regex_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(regex_label)
        row += 1
        self.form_vars['pattern'] = tk.StringVar()
        self.form_vars['pattern'].trace_add('write', lambda *args: self._test_pattern())
        ttk.Entry(parent, textvariable=self.form_vars['pattern']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        row += 1

        # Pattern Tester (Advanced)
        test_adv_label = tk.Label(parent, text="Test Your Pattern:", font=('Segoe UI', 10, 'bold'),
                                 bg=popup_bg, fg=colors['text_primary'])
        test_adv_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        self._popup_labels.append(test_adv_label)
        row += 1
        self.form_vars['test_input_advanced'] = tk.StringVar()
        self.form_vars['test_input_advanced'].trace_add('write', lambda *args: self._test_pattern())
        ttk.Entry(parent, textvariable=self.form_vars['test_input_advanced']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1

        self.test_result_label_advanced = tk.Label(parent, text="", font=('Segoe UI', 9),
                                                  bg=popup_bg, fg=colors['text_primary'], wraplength=450)
        self.test_result_label_advanced.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        self._popup_labels.append(self.test_result_label_advanced)
        row += 1
    
    def _on_mode_change(self):
        """Handle mode change between simple and advanced."""
        mode = self.form_vars['mode'].get()
        
        if mode == 'simple':
            self.simple_frame.grid()
            self.advanced_frame.grid_remove()
            self._update_simple_pattern()
        else:
            self.simple_frame.grid_remove()
            self.advanced_frame.grid()
    
    def _convert_date_format_to_regex(self, format_str: str) -> str:
        """
        Convert user-friendly date format to regex.
        
        Supported tokens:
        - dd: Day (01-31)
        - mm: Month (01-12)
        - yyyy: Year (4 digits)
        - yy: Year (2 digits)
        
        Examples:
        'dd/mm/yyyy' -> r'\\b(?:0?[1-9]|[12][0-9]|3[01])/(?:0?[1-9]|1[0-2])/\\d{4}\\b'
        'mm-dd-yyyy' -> r'\\b(?:0?[1-9]|1[0-2])-(?:0?[1-9]|[12][0-9]|3[01])-\\d{4}\\b'
        """
        # Define regex components
        replacements = {
            'yyyy': r'(?:19|20)\d{2}',  # 1900-2099
            'yy': r'\d{2}',              # Any 2 digits
            'mm': r'(?:0?[1-9]|1[0-2])', # 01-12 or 1-12
            'dd': r'(?:0?[1-9]|[12][0-9]|3[01])'  # 01-31 or 1-31
        }
        
        # Escape the format string to handle special regex chars in separators
        regex = format_str.lower()
        
        # Replace date tokens with regex patterns (order matters - yyyy before yy)
        for token, pattern in sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True):
            if token in regex:
                # Temporarily replace with a placeholder to avoid double-replacement
                placeholder = f"___{token.upper()}___"
                regex = regex.replace(token, placeholder)
        
        # Escape any special regex characters in separators
        regex = re.escape(regex)
        
        # Replace placeholders with actual regex patterns
        for token, pattern in replacements.items():
            placeholder = f"___{token.upper()}___"
            regex = regex.replace(re.escape(placeholder), pattern)
        
        # Add word boundaries
        regex = r'\b' + regex + r'\b'
        
        return regex
    
    def _generate_regex(self, pattern_type, search_text, case_sensitive, whole_word):
        """Generate regex from simple mode inputs."""
        # Pre-built templates
        templates = {
            'Email Address': r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b',
            'Phone Number (US)': r'\b(?:\+?1[-.]?)?\(?([0-9]{3})\)?[-.]?([0-9]{3})[-.]?([0-9]{4})\b',
            'Credit Card': r'\b(?:\d{4}[-\s]?){3}\d{4}\b',
            'URL/Link': r'https?://(?:www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b(?:[-a-zA-Z0-9()@:%_\+.~#?&/=]*)',
            'IP Address': r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b',
            'Date (MM/DD/YYYY)': r'\b(?:0?[1-9]|1[0-2])/(?:0?[1-9]|[12][0-9]|3[01])/(?:19|20)?\d{2}\b',
            'Date (DD/MM/YYYY)': r'\b(?:0?[1-9]|[12][0-9]|3[01])/(?:0?[1-9]|1[0-2])/(?:19|20)?\d{2}\b',
            'Date (YYYY-MM-DD)': r'\b(?:19|20)\d{2}-(?:0?[1-9]|1[0-2])-(?:0?[1-9]|[12][0-9]|3[01])\b',
        }
        
        if pattern_type in templates:
            return templates[pattern_type]
        
        if not search_text:
            return ""
        
        # Handle Custom Date Pattern
        if pattern_type == 'Custom Date Pattern':
            return self._convert_date_format_to_regex(search_text)
        
        # Escape special regex characters
        escaped = re.escape(search_text)
        
        # Build pattern based on type
        if pattern_type == 'Contains Text':
            pattern = escaped
        elif pattern_type == 'Starts With':
            pattern = f'^{escaped}'
        elif pattern_type == 'Ends With':
            pattern = f'{escaped}$'
        else:
            pattern = escaped
        
        # Add word boundaries if needed
        if whole_word and pattern_type == 'Contains Text':
            pattern = f'\\b{pattern}\\b'
        
        # Add case insensitivity flag if needed
        if not case_sensitive:
            pattern = f'(?i){pattern}'
        
        return pattern
    
    def _toggle_checkbox(self, key):
        """Toggle checkbox state and update icon."""
        var = self.form_vars.get(key)
        if var:
            # Toggle the boolean value
            var.set(not var.get())
            # Update the icon
            self._update_checkbox_icon(key)
            # Trigger pattern update
            self._update_simple_pattern()
    
    def _update_checkbox_icon(self, key):
        """Update a single checkbox icon based on its state."""
        var = self.form_vars.get(key)
        label = self._checkbox_labels.get(key)
        if var and label:
            is_checked = var.get()
            icon = self._checkbox_icons.get('checked' if is_checked else 'unchecked')
            if icon:
                label.configure(image=icon)
                label._icon_ref = icon
    
    def _update_all_checkbox_icons(self):
        """Update all checkbox icons (called after theme change)."""
        for key in self._checkbox_labels:
            self._update_checkbox_icon(key)

    def _update_simple_pattern(self):
        """Update the pattern preview from simple mode inputs."""
        if self.form_vars['mode'].get() != 'simple':
            return
        
        pattern_type = self.form_vars['pattern_type'].get()
        search_text = self.form_vars['search_text'].get()
        case_sensitive = self.form_vars['case_sensitive'].get()
        whole_word = self.form_vars['whole_word'].get()
        
        # Check if this is a template pattern (doesn't need search text)
        template_patterns = ['Email Address', 'Phone Number (US)', 'Credit Card', 'URL/Link', 'IP Address', 
                            'Date (MM/DD/YYYY)', 'Date (DD/MM/YYYY)', 'Date (YYYY-MM-DD)']
        is_template = pattern_type in template_patterns
        
        # Special case: Custom Date Pattern needs search text
        needs_search_text = not is_template or pattern_type == 'Custom Date Pattern'
        
        # Get theme colors
        colors = self._theme_manager.colors

        # Enable/disable search text entry based on pattern type
        if is_template:
            self.search_text_entry.config(state='disabled')
            # Clear the search text when switching to a template
            self.form_vars['search_text'].set('')
            # Make it visually grey when disabled
            style = ttk.Style()
            style.map('Disabled.TEntry',
                     fieldbackground=[('disabled', colors.get('button_secondary', '#e0e0e0'))],
                     foreground=[('disabled', colors['text_muted'])])
            self.search_text_entry.config(style='Disabled.TEntry')
            self.search_text_helper.config(text="(This template auto-generates the pattern - no search text needed)", foreground=colors['success'])
        elif pattern_type == 'Custom Date Pattern':
            self.search_text_entry.config(state='normal', style='TEntry')
            self.search_text_helper.config(text="(Enter date format: Use 'dd' for day, 'mm' for month, 'yyyy' for year. E.g., 'dd-mm-yyyy')", foreground=colors['text_muted'])
        else:
            self.search_text_entry.config(state='normal', style='TEntry')
            self.search_text_helper.config(text="(Enter the text pattern you want to match)", foreground=colors['text_muted'])
        
        regex = self._generate_regex(pattern_type, search_text, case_sensitive, whole_word)
        
        # Update the advanced pattern field
        self.form_vars['pattern'].set(regex)
        
        # Update preview
        if regex:
            self.preview_label.config(text=regex)
        else:
            if is_template:
                self.preview_label.config(text="Select a pattern type to see generated regex...")
            else:
                self.preview_label.config(text="Enter search text to see pattern...")
        
        # Test pattern
        self._test_pattern()
    
    def _test_pattern(self):
        """Test the current pattern against test input."""
        colors = self._theme_manager.colors
        mode = self.form_vars['mode'].get()
        pattern_str = self.form_vars['pattern'].get()

        if mode == 'simple':
            test_input = self.form_vars['test_input'].get()
            result_label = self.test_result_label
        else:
            test_input = self.form_vars['test_input_advanced'].get()
            result_label = self.test_result_label_advanced

        if not pattern_str or not test_input:
            result_label.config(text="", foreground=colors['text_muted'])
            return

        try:
            compiled = re.compile(pattern_str)
            matches = compiled.findall(test_input)

            if matches:
                match_str = '", "'.join(str(m) for m in matches[:3])
                if len(matches) > 3:
                    match_str += f", ... ({len(matches)} total)"
                result_label.config(text=f'Match found: "{match_str}"', foreground=colors['success'])
            else:
                result_label.config(text='No match', foreground=colors['error'])
        except re.error as e:
            result_label.config(text=f'Invalid regex: {str(e)[:50]}', foreground=colors['warning'])
    
    def _create_buttons(self, parent):
        """Create bottom buttons."""
        colors = self._theme_manager.colors
        popup_bg = colors['background']  # Use popup background, not section_bg

        self._bottom_btn_frame = tk.Frame(parent, bg=popup_bg)
        self._bottom_btn_frame.pack(fill=tk.X, pady=(20, 0))

        left = tk.Frame(self._bottom_btn_frame, bg=popup_bg)
        left.pack(side=tk.LEFT)
        self._left_btn_frame = left  # Store for theme updates

        # Reset button - secondary style (auto-sized)
        self.reset_btn = RoundedButton(
            left, text="Reset to Defaults", command=self._reset,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            canvas_bg=popup_bg
        )
        self.reset_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._bottom_buttons.append(self.reset_btn)

        # Export button - secondary style (auto-sized)
        self.export_btn = RoundedButton(
            left, text="Export", command=self._export,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            canvas_bg=popup_bg
        )
        self.export_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._bottom_buttons.append(self.export_btn)

        # Import button - secondary style (auto-sized)
        self.import_btn = RoundedButton(
            left, text="Import", command=self._import,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            canvas_bg=popup_bg
        )
        self.import_btn.pack(side=tk.LEFT)
        self._bottom_buttons.append(self.import_btn)

        right = tk.Frame(self._bottom_btn_frame, bg=popup_bg)
        right.pack(side=tk.RIGHT)
        self._right_btn_frame = right  # Store for theme updates

        # Save & Close button - primary style (auto-sized)
        self.save_close_btn = RoundedButton(
            right, text="Save & Close", command=self._save_close,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            height=32, radius=6, font=('Segoe UI', 10, 'bold'),
            canvas_bg=popup_bg
        )
        self.save_close_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._bottom_buttons.append(self.save_close_btn)

        # Close button - secondary style (auto-sized)
        self.close_btn = RoundedButton(
            right, text="Close", command=self._on_close,
            bg=colors['button_secondary'], hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'], fg=colors['text_primary'],
            height=32, radius=6, font=('Segoe UI', 9),
            canvas_bg=popup_bg
        )
        self.close_btn.pack(side=tk.LEFT)
        self._bottom_buttons.append(self.close_btn)
    
    def _populate_tree(self):
        """Populate tree with patterns."""
        colors = self._theme_manager.colors

        for item in self.tree.get_children():
            self.tree.delete(item)

        # Configure tree tags for risk colors
        self.tree.tag_configure('high_risk', foreground=colors['risk_high'])
        self.tree.tag_configure('medium_risk', foreground=colors['risk_medium'])
        self.tree.tag_configure('low_risk', foreground=colors['risk_low'])

        for pattern in self.patterns['patterns'].get('high_risk', []):
            self.tree.insert('', tk.END, values=(pattern['id'], pattern['name'], 'HIGH'), tags=('high_risk',))
        for pattern in self.patterns['patterns'].get('medium_risk', []):
            self.tree.insert('', tk.END, values=(pattern['id'], pattern['name'], 'MEDIUM'), tags=('medium_risk',))
        for pattern in self.patterns['patterns'].get('low_risk', []):
            self.tree.insert('', tk.END, values=(pattern['id'], pattern['name'], 'LOW'), tags=('low_risk',))

        # Update count label with defaults/custom status
        if hasattr(self, 'count_label'):
            count = sum(len(self.patterns['patterns'].get(r, [])) for r in ['high_risk', 'medium_risk', 'low_risk'])
            pattern_status = "custom" if self.custom_patterns_file.exists() else "defaults"
            self.count_label.configure(text=f"  ({count} patterns - {pattern_status})")
    
    def _on_select(self, event):
        """Handle pattern selection."""
        selection = self.tree.selection()
        if not selection:
            return
        
        values = self.tree.item(selection[0], 'values')
        pattern_id = values[0]
        
        # Find pattern
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            for pattern in self.patterns['patterns'].get(risk_level, []):
                if pattern['id'] == pattern_id:
                    self._load_to_form(pattern, risk_level)
                    return
    
    def _load_to_form(self, pattern, risk_level):
        """Load pattern into form."""
        self.selected_pattern = pattern
        self.form_vars['id'].set(pattern.get('id', ''))
        self.form_vars['name'].set(pattern.get('name', ''))
        self.form_vars['pattern'].set(pattern.get('pattern', ''))
        self.form_vars['risk_level'].set(risk_level)
        
        desc = self.form_vars['description']
        desc.delete('1.0', tk.END)
        desc.insert('1.0', pattern.get('description', ''))
        
        # Load recommendation if it exists
        rec = self.form_vars['recommendation']
        rec.delete('1.0', tk.END)
        rec.insert('1.0', pattern.get('recommendation', ''))
        
        ex = self.form_vars['examples']
        ex.delete('1.0', tk.END)
        ex.insert('1.0', '\n'.join(pattern.get('examples', [])))
        
        # Switch to advanced mode when loading existing pattern
        self.form_vars['mode'].set('advanced')
        self._on_mode_change()
    
    def _clear_form(self):
        """Clear form."""
        self.selected_pattern = None
        self.form_vars['id'].set('')
        self.form_vars['name'].set('')
        self.form_vars['pattern'].set('')
        self.form_vars['risk_level'].set('medium_risk')
        self.form_vars['description'].delete('1.0', tk.END)
        self.form_vars['recommendation'].delete('1.0', tk.END)
        self.form_vars['examples'].delete('1.0', tk.END)
        
        # Clear simple mode fields
        if 'search_text' in self.form_vars:
            self.form_vars['search_text'].set('')
        if 'test_input' in self.form_vars:
            self.form_vars['test_input'].set('')
        if 'test_input_advanced' in self.form_vars:
            self.form_vars['test_input_advanced'].set('')
        
        # Reset to simple mode
        self.form_vars['mode'].set('simple')
        self._on_mode_change()
    
    def _validate(self):
        """Validate form."""
        if not self.form_vars['id'].get():
            ThemedMessageBox.showerror(self.window, "Error", "Pattern ID is required")
            return False
        if not self.form_vars['name'].get():
            ThemedMessageBox.showerror(self.window, "Error", "Pattern Name is required")
            return False
        if not self.form_vars['pattern'].get():
            ThemedMessageBox.showerror(self.window, "Error", "Regex Pattern is required")
            return False
        if not self.form_vars['description'].get('1.0', tk.END).strip():
            ThemedMessageBox.showerror(self.window, "Error", "Description is required")
            return False

        # Validate regex
        try:
            re.compile(self.form_vars['pattern'].get())
        except re.error as e:
            ThemedMessageBox.showerror(self.window, "Invalid Regex", f"Invalid regular expression:\n{e}")
            return False

        return True
    
    def _form_to_pattern(self):
        """Convert form to pattern dict."""
        examples = [l.strip() for l in self.form_vars['examples'].get('1.0', tk.END).split('\n') if l.strip()]
        recommendation = self.form_vars['recommendation'].get('1.0', tk.END).strip()
        
        pattern = {
            'id': self.form_vars['id'].get(),
            'name': self.form_vars['name'].get(),
            'pattern': self.form_vars['pattern'].get(),
            'description': self.form_vars['description'].get('1.0', tk.END).strip(),
            'categories': [],
            'examples': examples
        }
        
        # Add recommendation if provided
        if recommendation:
            pattern['recommendation'] = recommendation
        
        return pattern
    
    def _add_pattern(self):
        """Add new pattern."""
        if not self._validate():
            return
        
        pattern_id = self.form_vars['id'].get()
        
        # Check duplicate
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            for p in self.patterns['patterns'].get(risk_level, []):
                if p['id'] == pattern_id:
                    ThemedMessageBox.showerror(self.window, "Error", f"Pattern ID '{pattern_id}' already exists")
                    return

        pattern = self._form_to_pattern()
        risk_level = self.form_vars['risk_level'].get()
        self.patterns['patterns'][risk_level].append(pattern)

        self._save_patterns()
        self._populate_tree()
        self._clear_form()
        ThemedMessageBox.showinfo(self.window, "Success", "Pattern added!")
    
    def _save_pattern(self):
        """Save changes to existing pattern."""
        if not self.selected_pattern:
            ThemedMessageBox.showwarning(self.window, "No Selection", "Please select a pattern to edit")
            return
        
        if not self._validate():
            return
        
        updated = self._form_to_pattern()
        
        # Find and update
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            patterns = self.patterns['patterns'][risk_level]
            for i, p in enumerate(patterns):
                if p['id'] == self.selected_pattern['id']:
                    new_risk = self.form_vars['risk_level'].get()
                    if risk_level != new_risk:
                        patterns.pop(i)
                        self.patterns['patterns'][new_risk].append(updated)
                    else:
                        patterns[i] = updated
                    break
        
        self._save_patterns()
        self._populate_tree()
        self._clear_form()
        ThemedMessageBox.showinfo(self.window, "Success", "Pattern updated!")

    def _delete_pattern(self):
        """Delete selected pattern."""
        if not self.selected_pattern:
            ThemedMessageBox.showwarning(self.window, "No Selection", "Please select a pattern to delete")
            return

        if not ThemedMessageBox.askyesno(self.window, "Confirm", f"Delete pattern '{self.selected_pattern['name']}'?"):
            return
        
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            patterns = self.patterns['patterns'][risk_level]
            for i, p in enumerate(patterns):
                if p['id'] == self.selected_pattern['id']:
                    patterns.pop(i)
                    break
        
        self._save_patterns()
        self._populate_tree()
        self._clear_form()
        ThemedMessageBox.showinfo(self.window, "Success", "Pattern deleted!")

    def _delete_selected_patterns(self):
        """Delete all selected patterns from the treeview (multi-select)."""
        selection = self.tree.selection()
        if not selection:
            ThemedMessageBox.showwarning(self.window, "No Selection", "Please select one or more patterns to delete")
            return

        count = len(selection)
        confirm_msg = f"Delete {count} selected pattern(s)?" if count > 1 else f"Delete selected pattern?"
        if not ThemedMessageBox.askyesno(self.window, "Confirm Delete", confirm_msg):
            return

        # Collect all pattern IDs to delete
        ids_to_delete = set()
        for item in selection:
            values = self.tree.item(item, 'values')
            ids_to_delete.add(values[0])  # Pattern ID is first column

        # Remove patterns from all risk levels
        deleted_count = 0
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            patterns = self.patterns['patterns'].get(risk_level, [])
            original_len = len(patterns)
            self.patterns['patterns'][risk_level] = [p for p in patterns if p['id'] not in ids_to_delete]
            deleted_count += original_len - len(self.patterns['patterns'][risk_level])

        self._save_patterns()
        self._populate_tree()
        self._clear_form()
        ThemedMessageBox.showinfo(self.window, "Success", f"{deleted_count} pattern(s) deleted!")

    def _save_patterns(self):
        """Save patterns to custom file."""
        try:
            with open(self.custom_patterns_file, 'w', encoding='utf-8') as f:
                json.dump(self.patterns, f, indent=2)
        except Exception as e:
            ThemedMessageBox.showerror(self.window, "Error", f"Failed to save: {e}")
    
    def _reset(self):
        """Reset to defaults."""
        if not ThemedMessageBox.askyesno(self.window, "Reset", "Reset all patterns to defaults?\n\nThis will delete all custom patterns."):
            return

        if self.custom_patterns_file.exists():
            self.custom_patterns_file.unlink()

        self.patterns = self._load_current_patterns()
        self._populate_tree()
        self._clear_form()
        ThemedMessageBox.showinfo(self.window, "Success", "Reset to defaults!")
    
    def _export(self):
        """Export patterns."""
        path = filedialog.asksaveasfilename(
            title="Export Patterns",
            defaultextension=".json",
            initialfile=f"patterns_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            filetypes=[("JSON Files", "*.json")]
        )
        
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(self.patterns, f, indent=2)
                ThemedMessageBox.showinfo(self.window, "Success", f"Exported to:\n{path}")
            except Exception as e:
                ThemedMessageBox.showerror(self.window, "Error", f"Export failed: {e}")
    
    def _import(self):
        """Import patterns from JSON file."""
        path = filedialog.askopenfilename(
            title="Import Patterns",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        
        if not path:
            return
        
        try:
            with open(path, 'r', encoding='utf-8') as f:
                imported = json.load(f)
            
            # Validate structure
            if 'patterns' not in imported:
                ThemedMessageBox.showerror(self.window, "Invalid File", "File must contain 'patterns' key with high_risk, medium_risk, and low_risk arrays.")
                return
            
            # Ask user how to handle import
            result = messagebox.askyesnocancel(
                "Import Mode",
                "How would you like to import?\n\n"
                "YES = Replace all patterns with imported\n"
                "NO = Merge (add new patterns, keep existing)\n"
                "CANCEL = Abort import"
            )
            
            if result is None:  # Cancel
                return
            
            if result:  # Yes = Replace
                self.patterns = imported
                mode = "replaced"
            else:  # No = Merge
                added_count = 0
                for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
                    existing_ids = {p['id'] for p in self.patterns['patterns'].get(risk_level, [])}
                    for pattern in imported['patterns'].get(risk_level, []):
                        if pattern['id'] not in existing_ids:
                            self.patterns['patterns'][risk_level].append(pattern)
                            added_count += 1
                mode = f"merged ({added_count} new patterns added)"
            
            self._save_patterns()
            self._populate_tree()
            self._clear_form()
            
            # Update count
            count = sum(len(self.patterns['patterns'].get(r, []))
                       for r in ['high_risk', 'medium_risk', 'low_risk'])
            ThemedMessageBox.showinfo(self.window, "Success", f"Import complete!\nMode: {mode}\nTotal patterns: {count}")

        except json.JSONDecodeError as e:
            ThemedMessageBox.showerror(self.window, "Invalid JSON", f"File is not valid JSON:\n{e}")
        except Exception as e:
            ThemedMessageBox.showerror(self.window, "Import Error", f"Failed to import:\n{e}")
    
    def _save_close(self):
        """Save and close."""
        self._save_patterns()
        if self.on_patterns_updated:
            self.on_patterns_updated()
        # Show message before closing window
        ThemedMessageBox.showinfo(self.window, "Success", "Patterns saved! Re-scan to apply changes.")
        self._on_close()

    def _on_close(self):
        """Handle window close - unregister theme callback."""
        try:
            self._theme_manager.unregister_theme_callback(self.on_theme_changed)
        except (ValueError, AttributeError):
            pass
        self.window.destroy()

    def on_theme_changed(self):
        """Handle theme changes."""
        colors = self._theme_manager.colors

        # Update window title bar color
        self._set_title_bar_color(self._theme_manager.is_dark)

        # Update window background - use background (white/dark), NOT section_bg
        self.window.configure(bg=colors['background'])

        # Update container frame and LabelFrame styles
        style = ttk.Style()
        style.configure('PatternManager.TFrame', background=colors['background'])
        style.configure('PatternManager.TLabelframe', background=colors['background'])
        style.configure('PatternManager.TLabelframe.Label', background=colors['background'],
                        foreground=colors['text_primary'])

        # Popup background for all labels
        popup_bg = colors['background']

        # Update all popup labels (regular labels with text_primary)
        if hasattr(self, '_popup_labels'):
            for label in self._popup_labels:
                try:
                    label.configure(bg=popup_bg, fg=colors['text_primary'])
                except tk.TclError:
                    pass

        # Update all hint labels (muted text)
        if hasattr(self, '_hint_labels'):
            for label in self._hint_labels:
                try:
                    label.configure(bg=popup_bg, fg=colors['text_muted'])
                except tk.TclError:
                    pass

        # Update all title labels (title color)
        if hasattr(self, '_title_labels'):
            for label in self._title_labels:
                try:
                    label.configure(bg=popup_bg, fg=colors['title_color'])
                except tk.TclError:
                    pass

        # Update mode and risk frames
        if hasattr(self, '_mode_frame') and self._mode_frame:
            try:
                self._mode_frame.configure(bg=popup_bg)
            except tk.TclError:
                pass
        if hasattr(self, '_risk_frame') and self._risk_frame:
            try:
                self._risk_frame.configure(bg=popup_bg)
            except tk.TclError:
                pass

        # Update pattern list header frame
        if hasattr(self, '_header_frame'):
            self._header_frame.configure(bg=colors['background'])
        if hasattr(self, '_header_title'):
            try:
                self._header_title.configure(foreground=colors['text_primary'])
            except tk.TclError:
                pass
        if hasattr(self, 'count_label'):
            try:
                self.count_label.configure(foreground=colors['text_muted'])
            except tk.TclError:
                pass

        # Update header delete button
        if hasattr(self, '_header_delete_btn'):
            try:
                self._header_delete_btn.configure(bg=colors['background'])
            except tk.TclError:
                pass

        # Update editor canvas background
        if hasattr(self, '_editor_canvas'):
            self._editor_canvas.configure(bg=colors['background'])

        # Update SVG radio rows
        for group in ['mode', 'risk']:
            self._update_radio_rows(group)
        
        # Reload checkbox icons for new theme and update
        self._load_checkbox_icons()
        self._update_all_checkbox_icons()

        # Update checkbox icon label backgrounds (uses background, not section_bg)
        options_bg = colors['background']
        for key, label in self._checkbox_labels.items():
            try:
                label.configure(bg=options_bg)
            except tk.TclError:
                pass

        # Update options frame background
        if hasattr(self, '_options_frame') and self._options_frame:
            try:
                self._options_frame.configure(bg=options_bg)
            except tk.TclError:
                pass

        # Update checkbox text labels
        if hasattr(self, '_case_text') and self._case_text:
            try:
                self._case_text.configure(bg=options_bg, fg=colors['text_primary'])
            except tk.TclError:
                pass
        if hasattr(self, '_word_text') and self._word_text:
            try:
                self._word_text.configure(bg=options_bg, fg=colors['text_primary'])
            except tk.TclError:
                pass

        # Update text widgets - use background (white/dark) for inputs
        for tw in self._text_widgets:
            try:
                tw.configure(
                    bg=colors['background'],
                    fg=colors['text_primary'],
                    highlightbackground=colors['border'],
                    highlightcolor=colors['primary']
                )
            except tk.TclError:
                pass

        # Update form button frame and buttons
        if hasattr(self, '_form_btn_frame') and self._form_btn_frame:
            try:
                self._form_btn_frame.configure(bg=colors['background'])
            except tk.TclError:
                pass

        for btn in self._form_buttons:
            try:
                btn.update_canvas_bg(colors['background'])
            except (tk.TclError, AttributeError):
                pass

        # Update bottom button frame and buttons
        if hasattr(self, '_bottom_btn_frame'):
            try:
                self._bottom_btn_frame.configure(bg=colors['background'])
                for child in self._bottom_btn_frame.winfo_children():
                    if isinstance(child, tk.Frame):
                        child.configure(bg=colors['background'])
            except tk.TclError:
                pass

        for btn in self._bottom_buttons:
            try:
                btn.update_canvas_bg(colors['background'])
            except (tk.TclError, AttributeError):
                pass

        # Update treeview style with groove headers (matching scan results table)
        style = ttk.Style()
        is_dark = self._theme_manager.is_dark
        if is_dark:
            heading_bg = '#2a2a3c'
            heading_fg = '#e0e0e0'
            header_separator = '#0d0d1a'
        else:
            heading_bg = '#f0f0f0'
            heading_fg = '#333333'
            header_separator = '#ffffff'

        style.configure("Flat.Treeview",
                        background=colors['section_bg'],
                        fieldbackground=colors['section_bg'],
                        lightcolor=colors['section_bg'],
                        darkcolor=colors['section_bg'],
                        bordercolor=colors['section_bg'])
        style.configure("Flat.Treeview.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator)
        style.map("Flat.Treeview.Heading",
                  relief=[('active', 'groove'), ('pressed', 'groove')],
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)])

        # Configure selection to only change background, preserve foreground (tag colors)
        selection_bg = colors.get('card_surface', colors.get('surface', '#3A3A3A' if is_dark else '#E0E0E0'))
        style.map("Flat.Treeview",
                  background=[('selected', selection_bg)])

        # Update pattern list section
        # Note: _list_outer_frame is now a ttk.LabelFrame which themes itself via ttk styles

        if hasattr(self, '_tree_border_frame'):
            try:
                self._tree_border_frame.configure(bg=colors['border'])
            except tk.TclError:
                pass

        if hasattr(self, '_tree_container_frame'):
            try:
                self._tree_container_frame.configure(bg=colors['section_bg'])
            except tk.TclError:
                pass

        if hasattr(self, 'count_label'):
            try:
                self.count_label.configure(foreground=colors['text_muted'])
            except tk.TclError:
                pass

        # Update tree tags for risk colors
        if hasattr(self, 'tree'):
            self.tree.tag_configure('high_risk', foreground=colors['risk_high'])
            self.tree.tag_configure('medium_risk', foreground=colors['risk_medium'])
            self.tree.tag_configure('low_risk', foreground=colors['risk_low'])
