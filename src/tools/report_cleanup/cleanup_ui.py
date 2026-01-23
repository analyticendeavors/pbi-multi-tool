"""
Report Cleanup UI - User interface for report cleanup operations
Built by Reid Havens of Analytic Endeavors

This module provides the user interface for the Report Cleanup tool,
following the established patterns from other tools in the suite.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable

from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin, RoundedButton, SquareIconButton, SVGToggle, ActionButtonBar, SplitLogSection, FileInputSection, ThemedScrollbar, ThemedMessageBox
from core.constants import AppConstants
from core.theme_manager import get_theme_manager
from core.pbi_file_reader import PBIPReader
from tools.report_cleanup.shared_types import CleanupOpportunity, RemovalResult, DuplicateImageGroup
from tools.report_cleanup.report_analyzer import ReportAnalyzer
from tools.report_cleanup.cleanup_engine import ReportCleanupEngine

# PIL for thumbnail loading (optional)
try:
    from PIL import Image, ImageTk
    import io
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# CairoSVG for SVG thumbnail rendering (optional)
try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except (ImportError, OSError):
    CAIROSVG_AVAILABLE = False


class DuplicateImageConfigDialog:
    """
    Modal dialog for selecting which duplicate image to keep in each group.
    Shows thumbnails, filenames, sizes, and reference counts.
    """

    @staticmethod
    def show(parent, duplicate_groups: List[DuplicateImageGroup], report_dir: Path) -> Optional[Dict[str, str]]:
        """
        Show the duplicate image configuration dialog.

        Args:
            parent: Parent window
            duplicate_groups: List of DuplicateImageGroup objects
            report_dir: Path to the report directory (for loading thumbnails)

        Returns:
            Dictionary of {group_id: selected_image_name} or None if cancelled
        """
        theme_manager = get_theme_manager()
        colors = theme_manager.colors
        is_dark = theme_manager.is_dark

        # Light mode uses lighter background for image rows for better button visibility
        row_bg = colors['card_surface'] if is_dark else '#f5f5f7'

        result = [None]
        selections = {}  # group_id -> selected image name

        # Initialize selections with current defaults
        for group in duplicate_groups:
            selections[group.group_id] = group.selected_image

        # Create dialog
        dialog = tk.Toplevel(parent)
        dialog.withdraw()
        dialog.title("Configure Duplicate Images")
        dialog.resizable(True, True)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])
        dialog.minsize(600, 400)

        # Set AE favicon
        try:
            base_path = Path(__file__).parent.parent.parent
            favicon_path = base_path / "assets" / "favicon.ico"
            if favicon_path.exists():
                dialog.iconbitmap(str(favicon_path))
        except Exception:
            pass

        # Set dark title bar on Windows
        try:
            from ctypes import windll, byref, sizeof, c_int
            HWND = windll.user32.GetParent(dialog.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = c_int(1 if is_dark else 0)
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_USE_IMMERSIVE_DARK_MODE, byref(value), sizeof(value))
        except Exception:
            pass

        # Main container
        main_frame = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header text
        header_label = tk.Label(
            main_frame,
            text="Select which image to keep for each duplicate group.\nOther duplicates will be removed and references updated.",
            bg=colors['background'],
            fg=colors['text_primary'],
            font=('Segoe UI', 10),
            justify=tk.LEFT
        )
        header_label.pack(fill=tk.X, pady=(0, 15))

        # Scrollable frame for groups
        canvas_frame = tk.Frame(main_frame, bg=colors['background'])
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(canvas_frame, bg=colors['background'], highlightthickness=0)
        scrollbar = ThemedScrollbar(canvas_frame, command=canvas.yview, theme_manager=theme_manager)
        scrollable_frame = tk.Frame(canvas, bg=colors['background'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Bind canvas resize to expand inner frame
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)

        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Store references to prevent garbage collection
        thumbnail_refs = []

        # Load radio-on and radio-off SVG icons
        radio_on_icon = None
        radio_off_icon = None
        if PIL_AVAILABLE and CAIROSVG_AVAILABLE:
            try:
                icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"
                radio_size = 16
                for icon_name, icon_ref in [("radio-on.svg", "on"), ("radio-off.svg", "off")]:
                    icon_path = icons_dir / icon_name
                    if icon_path.exists():
                        png_data = cairosvg.svg2png(url=str(icon_path), output_width=radio_size*2, output_height=radio_size*2)
                        img = Image.open(io.BytesIO(png_data))
                        if img.mode != 'RGBA':
                            img = img.convert('RGBA')
                        img = img.resize((radio_size, radio_size), Image.Resampling.LANCZOS)
                        if icon_ref == "on":
                            radio_on_icon = ImageTk.PhotoImage(img)
                        else:
                            radio_off_icon = ImageTk.PhotoImage(img)
            except Exception:
                pass

        # Helper to load thumbnail
        def load_thumbnail(image_path: Path, size: int = 64) -> Optional['ImageTk.PhotoImage']:
            if not PIL_AVAILABLE:
                return None
            try:
                if image_path.suffix.lower() == '.svg':
                    if not CAIROSVG_AVAILABLE:
                        return None
                    # Render SVG to PNG
                    png_data = cairosvg.svg2png(url=str(image_path), output_width=size*2, output_height=size*2)
                    img = Image.open(io.BytesIO(png_data))
                else:
                    img = Image.open(image_path)

                # Convert to RGBA for transparency support
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')

                # Resize maintaining aspect ratio
                img.thumbnail((size, size), Image.Resampling.LANCZOS)
                return ImageTk.PhotoImage(img)
            except Exception:
                return None

        # Helper to format bytes
        def format_bytes(bytes_count: int) -> str:
            if bytes_count < 1024:
                return f"{bytes_count} B"
            elif bytes_count < 1024 * 1024:
                return f"{bytes_count / 1024:.1f} KB"
            else:
                return f"{bytes_count / (1024 * 1024):.2f} MB"

        # Build groups UI
        images_dir = report_dir / "StaticResources" / "RegisteredResources"

        # Track all radio labels by group_id for easy updating
        radio_labels_by_group: Dict[str, List[tk.Label]] = {}

        for i, group in enumerate(duplicate_groups):
            # Group frame with border
            group_frame = tk.Frame(
                scrollable_frame,
                bg=colors['card_surface'],
                highlightbackground=colors['border'],
                highlightthickness=1
            )
            group_frame.pack(fill=tk.X, pady=(0 if i == 0 else 10, 0), padx=5)

            # Determine image type from first image in group
            first_image_name = group.images[0]['name'] if group.images else ''
            image_type = Path(first_image_name).suffix.upper().lstrip('.') if first_image_name else ''
            type_str = f" ({image_type})" if image_type else ""

            # Group header with image type
            group_header = tk.Label(
                group_frame,
                text=f"Duplicate Group {i + 1}{type_str} - {len(group.images)} images, save {format_bytes(group.savings_bytes)}",
                bg=colors['card_surface'],
                fg=colors['text_primary'],
                font=('Segoe UI', 10, 'bold'),
                anchor='w'
            )
            group_header.pack(fill=tk.X, padx=10, pady=(10, 5))

            # Images container (horizontal layout)
            images_frame = tk.Frame(group_frame, bg=row_bg)
            images_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

            # Radio button variable for this group
            radio_var = tk.StringVar(value=group.selected_image)

            def make_selection_handler(gid, var):
                def handler(*args):
                    selections[gid] = var.get()
                return handler

            radio_var.trace_add('write', make_selection_handler(group.group_id, radio_var))

            # Initialize list to track radio labels for this group
            radio_labels_by_group[group.group_id] = []

            for img_info in group.images:
                # Image item frame
                img_frame = tk.Frame(images_frame, bg=row_bg)
                img_frame.pack(side=tk.LEFT, padx=(0, 15), pady=5)

                # Thumbnail
                img_path = images_dir / img_info['name']
                thumb = load_thumbnail(img_path)

                if thumb:
                    thumb_label = tk.Label(img_frame, image=thumb, bg=row_bg)
                    thumb_label.pack()
                    thumbnail_refs.append(thumb)  # Keep reference
                else:
                    # Placeholder for missing/unsupported image
                    placeholder = tk.Label(
                        img_frame,
                        text="[No Preview]",
                        width=10,
                        height=4,
                        bg=colors['surface'],
                        fg=colors['text_secondary'],
                        font=('Segoe UI', 8)
                    )
                    placeholder.pack()

                # Filename (truncated if too long)
                name = img_info['name']
                display_name = name if len(name) <= 20 else name[:17] + "..."
                name_label = tk.Label(
                    img_frame,
                    text=display_name,
                    bg=row_bg,
                    fg=colors['text_primary'],
                    font=('Segoe UI', 8)
                )
                name_label.pack()

                # Size and references
                size_str = format_bytes(img_info.get('size_bytes', 0))
                refs = img_info.get('references', 0)
                refs_count = refs if isinstance(refs, int) else len(refs)
                info_label = tk.Label(
                    img_frame,
                    text=f"{size_str} | {refs_count} refs",
                    bg=row_bg,
                    fg=colors['text_secondary'],
                    font=('Segoe UI', 8)
                )
                info_label.pack()

                # Custom radio button with SVG icons
                radio_frame = tk.Frame(img_frame, bg=row_bg)
                radio_frame.pack(pady=(5, 0))

                # Determine if this is the selected keeper
                is_selected = radio_var.get() == img_info['name']
                icon_to_use = radio_on_icon if is_selected else radio_off_icon

                # Radio icon label
                radio_icon_label = tk.Label(
                    radio_frame,
                    image=icon_to_use if icon_to_use else None,
                    text="" if icon_to_use else ("*" if is_selected else "o"),
                    bg=row_bg,
                    fg=colors['text_primary'],
                    font=('Segoe UI', 8),
                    cursor='hand2'
                )
                radio_icon_label.pack(side=tk.LEFT)
                if icon_to_use:
                    radio_icon_label._icon_ref = icon_to_use  # Keep reference

                # "Keep" text label
                radio_text_label = tk.Label(
                    radio_frame,
                    text="Keep",
                    bg=row_bg,
                    fg=colors['text_primary'],
                    font=('Segoe UI', 8),
                    cursor='hand2'
                )
                radio_text_label.pack(side=tk.LEFT, padx=(3, 0))

                # Store references for updating
                radio_icon_label._radio_on = radio_on_icon
                radio_icon_label._radio_off = radio_off_icon
                radio_icon_label._img_name = img_info['name']

                # Add to tracking list for this group
                radio_labels_by_group[group.group_id].append(radio_icon_label)

                # Click handler to select this image
                def make_click_handler(img_name, var, gid):
                    def handler(event=None):
                        var.set(img_name)
                        selections[gid] = img_name
                        # Update all radio icons in this group using the tracked list
                        for lbl in radio_labels_by_group.get(gid, []):
                            is_sel = lbl._img_name == img_name
                            new_icon = lbl._radio_on if is_sel else lbl._radio_off
                            if new_icon:
                                lbl.config(image=new_icon)
                                lbl._icon_ref = new_icon
                            else:
                                lbl.config(text="*" if is_sel else "o")
                    return handler

                click_handler = make_click_handler(img_info['name'], radio_var, group.group_id)
                radio_icon_label.bind('<Button-1>', click_handler)
                radio_text_label.bind('<Button-1>', click_handler)

        # Button frame
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X, pady=(15, 0))

        button_inner = tk.Frame(button_frame, bg=colors['background'])
        button_inner.pack(side=tk.RIGHT)

        def on_ok():
            result[0] = selections.copy()
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        # OK button (primary)
        ok_btn = RoundedButton(
            button_inner,
            text="OK",
            command=on_ok,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors.get('button_primary_pressed', colors['button_primary_hover']),
            fg='#ffffff',
            height=32, radius=6,
            font=('Segoe UI', 10, 'bold'),
            canvas_bg=colors['background']
        )
        ok_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Cancel button (secondary)
        cancel_btn = RoundedButton(
            button_inner,
            text="Cancel",
            command=on_cancel,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors.get('button_secondary_pressed', colors['button_secondary_hover']),
            fg=colors['text_primary'],
            height=32, radius=6,
            font=('Segoe UI', 10),
            canvas_bg=colors['background']
        )
        cancel_btn.pack(side=tk.LEFT)

        # Handle close button and escape
        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        dialog.bind('<Escape>', lambda e: on_cancel())
        dialog.bind('<Return>', lambda e: on_ok())

        # Unbind mousewheel when dialog closes
        def on_close():
            canvas.unbind_all("<MouseWheel>")
        dialog.bind('<Destroy>', lambda e: on_close() if e.widget == dialog else None)

        # Center and show - dialog is taller to show more groups at once
        dialog.update_idletasks()
        width = max(700, dialog.winfo_reqwidth())
        height = min(800, max(500, dialog.winfo_reqheight()))
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.deiconify()

        # Wait for dialog to close
        dialog.wait_window()

        return result[0]


class CleanupOpportunityCard(tk.Frame):
    """
    Visual card showing a cleanup opportunity category.
    Shows what CAN be cleaned before scan (with "--"), then updates with counts after scan.
    Clicking the card shows details in the left panel.
    """

    def __init__(self, parent, emoji: str, title: str, subtitle: str,
                 checkbox_var: tk.BooleanVar, on_checkbox_change: Callable = None,
                 on_card_click: Callable = None, card_key: str = None,
                 icon: 'ImageTk.PhotoImage' = None, has_config_button: bool = False,
                 on_config_click: Callable = None, config_icon: 'ImageTk.PhotoImage' = None,
                 **kwargs):
        self._theme_manager = get_theme_manager()
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Use takefocus=False to prevent focus border
        # Border: blue (#00587C) in dark mode, teal (#009999) in light mode
        border_color = '#00587C' if is_dark else '#009999'
        super().__init__(parent, bg=colors['card_surface'],
                        highlightbackground=border_color,
                        highlightthickness=1, takefocus=False, **kwargs)

        self.checkbox_var = checkbox_var
        self.on_checkbox_change = on_checkbox_change
        self.on_card_click = on_card_click
        self.card_key = card_key
        self._count = None
        self._size_bytes = None
        self._enabled = False
        self._emoji = emoji
        self._title = title
        self._subtitle = subtitle
        self._selected = False  # Track if card is selected for details view
        self._icon = icon  # Optional SVG icon (fallback to emoji if None)
        self._has_config_button = has_config_button
        self._on_config_click = on_config_click
        self._config_icon = config_icon

        # Build card UI
        self._build_card()

        # Bind card click to show details (not toggle - toggle is via switch only)
        self._bind_card_click()

        # Register for theme changes
        self._theme_manager.register_theme_callback(self._handle_theme_change)

    def _build_card(self):
        """Build the card UI elements"""
        colors = self._theme_manager.colors

        # Inner padding frame - 8px horizontal padding to allow longer text like "Duplicate/Unused"
        inner = tk.Frame(self, bg=colors['card_surface'], padx=8, pady=15, takefocus=False)
        inner.pack(fill=tk.BOTH, expand=True)

        # Config icon in upper-right corner with rounded hover effect
        self.config_button = None
        self._config_hover_id = None  # Track the hover rectangle
        if self._has_config_button and self._config_icon:
            # Use canvas for rounded hover effect
            icon_size = 16  # Matches cogwheel_16 icon size
            padding = 4
            canvas_size = icon_size + padding * 2

            self.config_button = tk.Canvas(
                self,
                width=canvas_size,
                height=canvas_size,
                bg=colors['card_surface'],
                highlightthickness=0,
                cursor='hand2'
            )
            self.config_button._icon_ref = self._config_icon  # Prevent GC

            # Center the icon on canvas
            self._config_icon_id = self.config_button.create_image(
                canvas_size // 2, canvas_size // 2,
                image=self._config_icon,
                anchor='center'
            )

            self.config_button.place(relx=1.0, x=-2, y=2, anchor='ne')
            self.config_button.bind('<Button-1>', self._on_config_button_click)

            # Rounded hover effect
            def on_config_enter(e):
                current_colors = self._theme_manager.colors
                hover_color = current_colors['card_surface_hover']
                # Draw rounded rectangle behind icon
                if self._config_hover_id:
                    self.config_button.delete(self._config_hover_id)
                radius = 4
                x1, y1, x2, y2 = 2, 2, canvas_size - 2, canvas_size - 2
                # Create rounded rectangle using polygon points
                self._config_hover_id = self._draw_rounded_rect(
                    self.config_button, x1, y1, x2, y2, radius, hover_color
                )
                # Raise icon above the hover background
                self.config_button.tag_raise(self._config_icon_id)

            def on_config_leave(e):
                if self._config_hover_id:
                    self.config_button.delete(self._config_hover_id)
                    self._config_hover_id = None

            self.config_button.bind('<Enter>', on_config_enter)
            self.config_button.bind('<Leave>', on_config_leave)

        # Icon (SVG) or Emoji fallback
        if self._icon:
            self.icon_label = tk.Label(inner, image=self._icon,
                                       bg=colors['card_surface'])
            self.icon_label.pack()
            # Store reference to prevent garbage collection
            self.icon_label._icon_ref = self._icon
            self.emoji_label = None  # For theme update compatibility
        else:
            # Emoji icon (large) - fallback
            self.emoji_label = tk.Label(inner, text=self._emoji, font=('Segoe UI', 24),
                                        bg=colors['card_surface'], fg=colors['text_primary'])
            self.emoji_label.pack()
            self.icon_label = None

        # Title (e.g., "Unused")
        self.title_label = tk.Label(inner, text=self._title, font=('Segoe UI', 10, 'bold'),
                                    bg=colors['card_surface'], fg=colors['text_primary'])
        self.title_label.pack(pady=(5, 0))

        # Subtitle (e.g., "Themes")
        self.subtitle_label = tk.Label(inner, text=self._subtitle, font=('Segoe UI', 9),
                                       bg=colors['card_surface'], fg=colors['text_secondary'])
        self.subtitle_label.pack()

        # Toggle switch frame (hidden initially, appears between subtitle and count)
        self.toggle_frame = tk.Frame(inner, bg=colors['card_surface'], takefocus=False)
        # Don't pack yet - shown after analysis

        # Use SVGToggle (same as Report Merger) instead of canvas-based ToggleSwitch
        toggle_on_path = Path(__file__).parent.parent.parent / "assets" / "Tool Icons" / "toggle-on.svg"
        toggle_off_path = Path(__file__).parent.parent.parent / "assets" / "Tool Icons" / "toggle-off.svg"

        self.toggle_switch = SVGToggle(
            self.toggle_frame,
            svg_on=str(toggle_on_path),
            svg_off=str(toggle_off_path),
            command=self._on_toggle_changed,
            initial_state=self.checkbox_var.get(),
            width=44,  # Same size as previous ToggleSwitch
            height=24,
            theme_manager=self._theme_manager
        )
        self.toggle_switch.pack(side=tk.LEFT)

        # Count label (shows "--" initially)
        self.count_label = tk.Label(inner, text="--", font=('Segoe UI', 16, 'bold'),
                                    bg=colors['card_surface'], fg=colors['text_muted'])
        self.count_label.pack(pady=(8, 0))

        # Size label (optional, for themes/visuals)
        self.size_label = tk.Label(inner, text="", font=('Segoe UI', 9),
                                   bg=colors['card_surface'], fg=colors['text_muted'])
        self.size_label.pack()

        # Store inner frame for theme updates
        self._inner = inner

    def _on_config_button_click(self, event=None):
        """Handle config button click"""
        if self._on_config_click and self._enabled:
            self._on_config_click(self.card_key)

    def _draw_rounded_rect(self, canvas, x1, y1, x2, y2, radius, fill):
        """Draw a rounded rectangle on a canvas and return its ID"""
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
            x1 + radius, y1,
        ]
        return canvas.create_polygon(points, fill=fill, smooth=True, outline='')

    def _bind_card_click(self):
        """Bind click events to show details (not toggle)"""
        # Bind to the card itself and all child widgets except toggle
        def on_click(event):
            # Don't trigger card click if clicking the toggle switch
            if isinstance(event.widget, SVGToggle):
                return
            if self._enabled and self.on_card_click:
                self.on_card_click(self.card_key)

        self.bind('<Button-1>', on_click)
        self._inner.bind('<Button-1>', on_click)
        # Bind to icon or emoji label (whichever exists)
        if self.icon_label:
            self.icon_label.bind('<Button-1>', on_click)
        if self.emoji_label:
            self.emoji_label.bind('<Button-1>', on_click)
        self.title_label.bind('<Button-1>', on_click)
        self.subtitle_label.bind('<Button-1>', on_click)
        self.count_label.bind('<Button-1>', on_click)
        self.size_label.bind('<Button-1>', on_click)

        # Change cursor to pointer on hover when enabled
        def on_enter(event):
            if self._enabled:
                self.config(cursor='hand2')

        def on_leave(event):
            self.config(cursor='')

        self.bind('<Enter>', on_enter)
        self.bind('<Leave>', on_leave)

    def _on_toggle_changed(self, state: bool):
        """Handle toggle switch change - sync with checkbox_var"""
        self.checkbox_var.set(state)
        if self.on_checkbox_change:
            self.on_checkbox_change()

    def set_selected(self, selected: bool):
        """Set the card as selected (showing details)"""
        self._selected = selected
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark
        # Border: blue (#00587C) in dark mode, teal (#009999) in light mode
        border_color = '#00587C' if is_dark else '#009999'

        if selected and self._enabled:
            # Highlight selected card with thicker border
            self.config(highlightthickness=2, highlightbackground=border_color)
        else:
            # Normal border
            self.config(highlightthickness=1, highlightbackground=border_color)

    def update_results(self, count: int, size_bytes: int = None):
        """Update card with analysis results"""
        colors = self._theme_manager.colors
        self._count = count
        self._size_bytes = size_bytes
        self._enabled = count > 0

        if count > 0:
            # Has opportunities - show count and toggle switch
            self.count_label.config(text=str(count), fg=colors['text_primary'])
            self.toggle_frame.pack(pady=(8, 0))  # Show toggle
            self.toggle_switch.set_enabled(True)

            if size_bytes and size_bytes > 0:
                size_text = f"~{size_bytes // 1024}KB" if size_bytes >= 1024 else f"~{size_bytes}B"
                self.size_label.config(text=size_text, fg=colors['text_secondary'])
            else:
                self.size_label.config(text="")

            # Normal card appearance
            self.config(bg=colors['card_surface'])
            self._inner.config(bg=colors['card_surface'])
            self._update_label_backgrounds(colors['card_surface'])
        else:
            # No opportunities - grey out
            self.count_label.config(text="0", fg=colors['text_muted'])
            self.toggle_frame.pack_forget()  # Hide toggle
            self.size_label.config(text="")

            # Grey out the card
            grey_bg = colors['card_surface_hover']
            self.config(bg=grey_bg)
            self._inner.config(bg=grey_bg)
            self._update_label_backgrounds(grey_bg)

    def _update_label_backgrounds(self, bg_color: str):
        """Update all label backgrounds"""
        if self.emoji_label:
            self.emoji_label.config(bg=bg_color)
        if self.icon_label:
            self.icon_label.config(bg=bg_color)
        if self.config_button:
            self.config_button.config(bg=bg_color)
        self.title_label.config(bg=bg_color)
        self.subtitle_label.config(bg=bg_color)
        self.count_label.config(bg=bg_color)
        self.size_label.config(bg=bg_color)
        self.toggle_frame.config(bg=bg_color)
        self.toggle_switch.config(bg=bg_color)

    def reset(self):
        """Reset card to initial state"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        self.count_label.config(text="--", fg=colors['text_muted'])
        self.size_label.config(text="")
        self.toggle_frame.pack_forget()
        self.checkbox_var.set(True)  # Reset to checked
        self.toggle_switch.set_state(True, trigger_command=False)  # Sync SVGToggle state
        self._count = None
        self._size_bytes = None
        self._enabled = False
        self._selected = False

        # Restore normal appearance: border blue (#00587C) in dark, teal (#009999) in light
        border_color = '#00587C' if is_dark else '#009999'
        self.config(bg=colors['card_surface'], highlightthickness=1, highlightbackground=border_color)
        self._inner.config(bg=colors['card_surface'])
        self._update_label_backgrounds(colors['card_surface'])

    def _handle_theme_change(self, theme: str):
        """Handle theme change - theme parameter is 'dark' or 'light'"""
        # Get current colors from theme manager
        colors = self._theme_manager.colors

        # Update colors based on current state
        if self._enabled:
            bg = colors['card_surface']
            self.count_label.config(fg=colors['text_primary'])
        elif self._count == 0:
            bg = colors['card_surface_hover']
            self.count_label.config(fg=colors['text_muted'])
        else:
            bg = colors['card_surface']
            self.count_label.config(fg=colors['text_muted'])

        # Update border based on selection state
        # Border: blue (#00587C) in dark mode, teal (#009999) in light mode
        is_dark = self._theme_manager.is_dark
        border_color = '#00587C' if is_dark else '#009999'
        if self._selected and self._enabled:
            self.config(bg=bg, highlightthickness=2, highlightbackground=border_color)
        else:
            self.config(bg=bg, highlightthickness=1, highlightbackground=border_color)

        self._inner.config(bg=bg)
        self._update_label_backgrounds(bg)

        if self.emoji_label:
            self.emoji_label.config(fg=colors['text_primary'])
        self.title_label.config(fg=colors['text_primary'])
        self.subtitle_label.config(fg=colors['text_secondary'])

        if self._size_bytes and self._size_bytes > 0:
            self.size_label.config(fg=colors['text_secondary'])

        # Update config button background if present
        if self.config_button:
            self.config_button.config(bg=bg)

    @property
    def is_enabled(self) -> bool:
        """Returns True if this card has opportunities"""
        return self._enabled

    @property
    def is_checked(self) -> bool:
        """Returns True if checkbox is checked and card has opportunities"""
        return self._enabled and self.checkbox_var.get()


class ReportCleanupTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """
    Report Cleanup UI Tab - provides interface for cleaning up Power BI reports
    """
    
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "report_cleanup", "Report Cleanup")
        
        # Tool instances
        self.analyzer = ReportAnalyzer(logger_callback=self.log_message)
        self.cleanup_engine = ReportCleanupEngine(logger_callback=self.log_message)
        
        # UI state
        self.current_opportunities: List[CleanupOpportunity] = []
        self._duplicate_groups: List[DuplicateImageGroup] = []  # For duplicate image config dialog
        self._report_dir: Optional[Path] = None  # Report directory for image operations
        self._duplicate_image_selections: Dict[str, str] = {}  # group_id -> keeper image name
        self.cleanup_options = {
            'remove_themes': tk.BooleanVar(value=True),
            'remove_custom_visuals': tk.BooleanVar(value=True),
            'remove_bookmarks': tk.BooleanVar(value=True),
            'hide_visual_filters': tk.BooleanVar(value=True),
            'remove_saved_scripts': tk.BooleanVar(value=True),  # Combined DAX queries and TMDL scripts
            'image_cleanup': tk.BooleanVar(value=True),  # Combined duplicate + unused images
        }

        # UI components
        self.file_section = None  # FileInputSection instance
        self.pbip_path_var = None  # Set from file_section.path_var
        self.analyze_button = None  # Set from file_section.action_button
        self.opportunity_cards = {}
        self.cards_hint_label = None
        self.clean_button = None
        self.reset_button = None
        self.progress_components = None

        # Setup UI and show welcome message
        self.setup_ui()
        self._show_welcome_message()

    def on_theme_changed(self, theme: str):
        """Update theme-dependent widgets when theme changes"""
        super().on_theme_changed(theme)
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # FileInputSection handles its own theme updates via _on_theme_changed

    def setup_ui(self) -> None:
        """Setup the Report Cleanup UI - redesigned with opportunity cards"""
        # Load UI icons for buttons and section headers
        if not hasattr(self, '_button_icons'):
            self._button_icons = {}

        # Load 16px icons for buttons/headers
        icon_names_16 = ["Power-BI", "broom", "bar-chart", "magnifying-glass", "reset", "analyze", "folder"]
        for name in icon_names_16:
            icon = self._load_icon_for_button(name, size=16)
            if icon:
                self._button_icons[name] = icon

        # Load 24px icons for cards (larger to match emoji size)
        # Order: Bookmarks, Filters, Themes, Visuals, Scripts, Duplicate/Unused Images
        icon_names_24 = ["bookmark", "filter", "paint", "bar-chart", "scripts", "delete-image"]
        for name in icon_names_24:
            icon = self._load_icon_for_button(name, size=24)
            if icon:
                self._button_icons[f"{name}_24"] = icon

        # Load 16px cogwheel icon for config button on cards
        cogwheel_icon = self._load_icon_for_button("cogwheel", size=16)
        if cogwheel_icon:
            self._button_icons["cogwheel_16"] = cogwheel_icon

        # IMPORTANT: Create action buttons FIRST with side=BOTTOM
        # This ensures buttons are always visible even when window shrinks
        self._setup_action_buttons()

        # Now create other sections from top (they will fill remaining space)
        self._setup_file_input_section()
        self._setup_opportunity_cards()
        self._setup_progress_section()
        self._setup_log_section()

        # Setup path cleaning
        self.setup_path_cleaning(self.pbip_path_var)
    
    def _setup_file_input_section(self):
        """Setup the file input section using FileInputSection template"""
        # Get analyze icon
        analyze_icon = self._button_icons.get('magnifying-glass')

        # Create FileInputSection with all components
        self.file_section = FileInputSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="Report Source",
            section_icon="Power-BI",
            file_label="Power BI File:",
            file_types=[("Power BI Files", "*.pbix *.pbip"), ("PBIX Files", "*.pbix"), ("PBIP Files", "*.pbip")],
            action_button_text="ANALYZE REPORTS",
            action_button_command=self._analyze_report,
            action_button_icon=analyze_icon,
            help_command=self.show_help_dialog
        )
        self.file_section.pack(fill=tk.X, pady=(0, 15))

        # Store references for backward compatibility
        self.pbip_path_var = self.file_section.path_var
        self.analyze_button = self.file_section.action_button
        self._data_source_section = self.file_section.section_frame

    def _setup_opportunity_cards(self):
        """Create the 4 cleanup opportunity cards"""
        colors = self._theme_manager.colors

        # Main section frame with icon + text header
        header_widget = self.create_section_header(self.frame, "Cleanup Opportunities", "broom")[0]
        section_frame = ttk.LabelFrame(self.frame, labelwidget=header_widget,
                                       style='Section.TLabelframe', padding="15")
        section_frame.pack(fill=tk.X, pady=(0, 15))
        self.opportunity_section = section_frame

        # Inner content frame - use tk.Frame for proper theme bg
        # Note: padding is handled via grid padx/pady for uniform spacing
        content_frame = tk.Frame(section_frame, bg=colors['background'])
        content_frame.pack(fill=tk.X, expand=True)
        self._cards_content_frame = content_frame

        # Cards container - center the cards, use grid for even spacing
        cards_container = tk.Frame(content_frame, bg=colors['background'])
        cards_container.pack(fill=tk.X, expand=True)
        self._cards_container = cards_container

        # Configure grid columns for even distribution
        for i in range(6):
            cards_container.columnconfigure(i, weight=1, uniform="cards")

        # Card configurations: (icon_key, emoji_fallback, title, subtitle, option_key)
        # Order: Bookmarks, Filters, Themes, Visuals, Scripts, Image Cleanup (duplicate + unused)
        card_configs = [
            ('bookmark_24', '📖', 'Unused', 'Bookmarks', 'remove_bookmarks'),
            ('filter_24', '🎯', 'Visual', 'Filters', 'hide_visual_filters'),
            ('paint_24', '🎨', 'Unused', 'Themes', 'remove_themes'),
            ('bar-chart_24', '📊', 'Unused', 'Visuals', 'remove_custom_visuals'),
            ('scripts_24', '📝', 'Saved Scripts', 'DAX / TMDL', 'remove_saved_scripts'),
            ('delete-image_24', '🖼️', 'Duplicate/Unused', 'Images', 'image_cleanup'),
        ]

        # Create cards dictionary
        self.opportunity_cards = {}

        # Get config icon for Duplicate Images card
        config_icon = self._button_icons.get("cogwheel_16")

        # Uniform card spacing: LabelFrame padding="15" + grid padx/pady=15 = 30px
        # Border to card: 15 + 15 = 30px, Between cards: 15 + 15 = 30px (uniform)
        card_padding = 15
        for i, (icon_key, emoji, title, subtitle, option_key) in enumerate(card_configs):
            card_icon = self._button_icons.get(icon_key)

            # Add config button only for Image Cleanup card (duplicate/unused)
            has_config = (option_key == 'image_cleanup')

            card = CleanupOpportunityCard(
                cards_container,
                emoji=emoji,
                title=title,
                subtitle=subtitle,
                checkbox_var=self.cleanup_options[option_key],
                on_checkbox_change=self._update_clean_button_state,
                on_card_click=self._on_card_clicked,
                icon=card_icon,
                card_key=option_key,
                has_config_button=has_config,
                on_config_click=self._on_duplicate_image_config_click if has_config else None,
                config_icon=config_icon if has_config else None
            )
            card.grid(row=0, column=i, padx=card_padding, pady=card_padding, sticky='nsew')
            self.opportunity_cards[option_key] = card

        # Hint text below cards - use same padding for consistency
        self.cards_hint_label = tk.Label(
            content_frame,
            text="Scan a report to discover cleanup opportunities",
            bg=colors['background'],
            fg=colors['text_secondary'],
            font=('Segoe UI', 10, 'italic')
        )
        self.cards_hint_label.pack(pady=(card_padding, card_padding))
        self._cards_hint_label_bg = colors['background']

        # Track currently selected card for details
        self._selected_card_key = None

    def _setup_action_buttons(self):
        """Create bottom action buttons (Clean Selected + Reset All)"""
        broom_icon = self._button_icons.get('broom')
        reset_icon = self._button_icons.get('reset')

        # Use centralized ActionButtonBar
        self.button_frame = ActionButtonBar(
            parent=self.frame,
            theme_manager=self._theme_manager,
            primary_text="CLEAN SELECTED",
            primary_command=self._remove_selected_items,
            primary_icon=broom_icon,
            secondary_text="RESET ALL",
            secondary_command=self.reset_tab,
            secondary_icon=reset_icon,
            primary_starts_disabled=True
        )
        self.button_frame.pack(side=tk.BOTTOM, pady=(15, 0))

        # Expose buttons for compatibility with existing code
        self.clean_button = self.button_frame.primary_button
        self.reset_button = self.button_frame.secondary_button

        # Track in button lists for any additional theme handling
        self._primary_buttons.append(self.clean_button)
        self._secondary_buttons.append(self.reset_button)

    def _update_clean_button_state(self):
        """Enable Clean button only if analysis done AND at least one checkbox checked"""
        if not hasattr(self, 'clean_button') or not self.clean_button:
            return

        # Must have opportunities found
        if not self.current_opportunities:
            self.clean_button.set_enabled(False)
            return

        # Check if any enabled card has its checkbox checked
        any_selected = False
        for option_key, card in self.opportunity_cards.items():
            if card.is_checked:
                any_selected = True
                break

        self.clean_button.set_enabled(any_selected)

    def _setup_progress_section(self):
        """Setup the progress section"""
        self.progress_components = self.create_progress_bar(self.frame)
        # Initially hidden
        self.progress_components['frame'].pack_forget()
    
    def _position_progress_frame(self):
        """Position the progress frame appropriately for this layout"""
        if self.progress_components and self.progress_components['frame']:
            # Position above the action buttons - pack with side=BOTTOM (no before=)
            # Since button_frame was packed first with side=BOTTOM, this will appear above it
            self.progress_components['frame'].pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 5))
    
    def _setup_log_section(self):
        """Setup the split log section with Details (left) and Progress Log (right)"""
        # Use SplitLogSection template widget
        self.log_section = SplitLogSection(
            parent=self.frame,
            theme_manager=self._theme_manager,
            section_title="Analysis & Progress",
            section_icon="analyze",
            summary_title="Cleanup Details",
            summary_icon="bar-chart",
            log_title="Progress Log",
            log_icon="log-file",
            summary_placeholder="Select an item to see details"
        )
        self.log_section.pack(fill=tk.BOTH, expand=True, pady=(0, 0))

        # Store references for compatibility
        self.log_section_frame = self.log_section
        self._summary_frame = self.log_section.summary_frame
        self._placeholder_label = self.log_section.placeholder_label

        # Use the provided summary_text (scrolledtext) for details
        self.details_text = self.log_section.summary_text

        # Store reference to log text (right side)
        self.log_text = self.log_section.log_text

        # Show initial placeholder in details panel
        self._show_details_placeholder()

    def _show_details_placeholder(self):
        """Show placeholder text in details panel - hide text widget, show placeholder label"""
        # Hide the details text widget
        if hasattr(self, 'details_text') and self.details_text:
            self.details_text.grid_remove()

        # Show the placeholder label (centered)
        if hasattr(self, '_placeholder_label') and self._placeholder_label:
            self._placeholder_label.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

    def _on_card_clicked(self, card_key: str):
        """Handle card click to show details in left panel"""
        if not card_key:
            return

        # Update selection state on all cards
        for key, card in self.opportunity_cards.items():
            card.set_selected(key == card_key)

        self._selected_card_key = card_key

        # Show details for the clicked card
        self._show_card_details(card_key)

    def _on_duplicate_image_config_click(self, card_key: str):
        """Handle config button click on Duplicate Images card to open selection dialog"""
        if not self._duplicate_groups:
            self.log_message("No duplicate image groups found.")
            return

        if not self._report_dir or not self._report_dir.exists():
            self.log_message("Report directory not available.")
            return

        # Open the config dialog
        selections = DuplicateImageConfigDialog.show(
            self.frame,
            self._duplicate_groups,
            self._report_dir
        )

        if selections is not None:
            # User clicked OK - store selections
            self._duplicate_image_selections = selections
            count = len([g for g in self._duplicate_groups if g.group_id in selections])
            self.log_message(f"Configured duplicate image selections for {count} group(s)")
        # If selections is None, user cancelled - keep previous selections

    def _show_card_details(self, card_key: str):
        """Show details for the selected card category in the details panel"""
        if not hasattr(self, 'details_text') or not self.details_text:
            return

        colors = self._theme_manager.colors

        # Hide placeholder and show details text widget
        if hasattr(self, '_placeholder_label') and self._placeholder_label:
            self._placeholder_label.grid_remove()
        self.details_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Get opportunities for this category
        category_map = {
            'remove_themes': ('theme', 'Unused Themes'),
            'remove_custom_visuals': (['custom_visual_build_pane', 'custom_visual_hidden'], 'Unused Visuals'),
            'remove_bookmarks': (['bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group'], 'Unused Bookmarks'),
            'hide_visual_filters': ('visual_filter', 'Visual Filters'),
            'remove_saved_scripts': (['dax_query', 'tmdl_script'], 'Saved Scripts'),
            'image_cleanup': (['duplicate_image', 'unused_image'], 'Duplicate/Unused Images'),
        }

        item_types, title = category_map.get(card_key, (None, 'Details'))

        # Filter opportunities for this category
        if isinstance(item_types, list):
            items = [op for op in self.current_opportunities if op.item_type in item_types]
        else:
            items = [op for op in self.current_opportunities if op.item_type == item_types]

        # Build details text
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)

        if not items:
            self.details_text.insert(tk.END, f"{title}\n\n", 'header')
            self.details_text.insert(tk.END, "No items found in this category.\n", 'info')
        else:
            self.details_text.insert(tk.END, f"{title} ({len(items)})\n", 'header')
            self.details_text.insert(tk.END, "─" * 30 + "\n\n", 'separator')

            if card_key == 'remove_themes':
                # Show theme details - total at top for visibility
                total_size = sum(op.size_bytes for op in items)
                self.details_text.insert(tk.END, f"Total: ~{self._format_bytes(total_size)}\n\n", 'summary')
                for op in items:
                    size_str = f" (~{op.size_bytes // 1024}KB)" if op.size_bytes >= 1024 else f" (~{op.size_bytes}B)" if op.size_bytes > 0 else ""
                    self.details_text.insert(tk.END, f"• {op.item_name}{size_str}\n", 'item')

            elif card_key == 'remove_custom_visuals':
                # Group visuals by type - total at top for visibility
                build_pane = [op for op in items if op.item_type == 'custom_visual_build_pane']
                hidden = [op for op in items if op.item_type == 'custom_visual_hidden']

                total_size = sum(op.size_bytes for op in items)
                self.details_text.insert(tk.END, f"Total: ~{self._format_bytes(total_size)}\n\n", 'summary')

                if build_pane:
                    self.details_text.insert(tk.END, "Build Pane Visuals:\n", 'subheader')
                    for op in build_pane:
                        size_str = f" (~{op.size_bytes // 1024}KB)" if op.size_bytes >= 1024 else ""
                        self.details_text.insert(tk.END, f"  🔮 {op.item_name}{size_str}\n", 'item')
                    self.details_text.insert(tk.END, "\n")

                if hidden:
                    self.details_text.insert(tk.END, "Hidden Visuals:\n", 'subheader')
                    for op in hidden:
                        size_str = f" (~{op.size_bytes // 1024}KB)" if op.size_bytes >= 1024 else ""
                        self.details_text.insert(tk.END, f"  👻 {op.item_name}{size_str}\n", 'item')

            elif card_key == 'remove_bookmarks':
                # Group bookmarks by type
                guaranteed = [op for op in items if op.item_type == 'bookmark_guaranteed_unused']
                likely = [op for op in items if op.item_type == 'bookmark_likely_unused']
                empty_groups = [op for op in items if op.item_type == 'bookmark_empty_group']

                if guaranteed:
                    self.details_text.insert(tk.END, "Guaranteed Unused (page missing):\n", 'subheader')
                    for op in guaranteed:
                        self.details_text.insert(tk.END, f"  ✅ {op.item_name}\n", 'item')
                    self.details_text.insert(tk.END, "\n")

                if likely:
                    self.details_text.insert(tk.END, "Likely Unused (no navigation):\n", 'subheader')
                    for op in likely:
                        self.details_text.insert(tk.END, f"  ⚠️ {op.item_name}\n", 'item')
                    self.details_text.insert(tk.END, "\n")

                if empty_groups:
                    self.details_text.insert(tk.END, "Empty Groups:\n", 'subheader')
                    for op in empty_groups:
                        self.details_text.insert(tk.END, f"  📁 {op.item_name}\n", 'item')

            elif card_key == 'hide_visual_filters':
                # Show filter count
                filter_op = items[0] if items else None
                if filter_op:
                    self.details_text.insert(tk.END, f"Total filters: {filter_op.filter_count}\n\n", 'item')
                    self.details_text.insert(tk.END, "Visual-level filters will be hidden (not removed) to clean up the filter pane interface.\n\n", 'info')
                    self.details_text.insert(tk.END, "💡 Filters can be shown again later if needed.\n", 'tip')

            elif card_key == 'remove_saved_scripts':
                # Show combined DAX queries and TMDL scripts details - total at top for visibility
                dax_queries = [op for op in items if op.item_type == 'dax_query']
                tmdl_scripts = [op for op in items if op.item_type == 'tmdl_script']

                total_size = sum(op.size_bytes for op in items)
                self.details_text.insert(tk.END, f"Total: ~{self._format_bytes(total_size)}\n\n", 'summary')

                if dax_queries:
                    self.details_text.insert(tk.END, "DAX Queries:\n", 'subheader')
                    for op in dax_queries:
                        size_str = f" (~{op.size_bytes // 1024}KB)" if op.size_bytes >= 1024 else f" (~{op.size_bytes}B)" if op.size_bytes > 0 else ""
                        self.details_text.insert(tk.END, f"  📝 {op.item_name}{size_str}\n", 'item')
                    self.details_text.insert(tk.END, "\n")

                if tmdl_scripts:
                    self.details_text.insert(tk.END, "TMDL Scripts:\n", 'subheader')
                    for op in tmdl_scripts:
                        size_str = f" (~{op.size_bytes // 1024}KB)" if op.size_bytes >= 1024 else f" (~{op.size_bytes}B)" if op.size_bytes > 0 else ""
                        self.details_text.insert(tk.END, f"  📜 {op.item_name}{size_str}\n", 'item')
                    self.details_text.insert(tk.END, "\n")

                self.details_text.insert(tk.END, "💡 These are development artifacts from DAX Query View and TMDL View.\n", 'tip')

            elif card_key == 'image_cleanup':
                # Combined card - show unused first, then duplicates
                unused_items = [op for op in items if op.item_type == 'unused_image']
                duplicate_items = [op for op in items if op.item_type == 'duplicate_image']

                # Calculate totals
                unused_size = sum(op.size_bytes for op in unused_items)
                duplicate_size = sum(op.size_bytes for op in duplicate_items)
                total_size = unused_size + duplicate_size
                self.details_text.insert(tk.END, f"Total: ~{self._format_bytes(total_size)}\n\n", 'summary')

                # Unused Images section (first)
                if unused_items:
                    self.details_text.insert(tk.END, f"Unused Images ({len(unused_items)}):\n", 'subheader')
                    for op in unused_items:
                        size_str = f" (~{op.size_bytes // 1024}KB)" if op.size_bytes >= 1024 else f" (~{op.size_bytes}B)" if op.size_bytes > 0 else ""
                        self.details_text.insert(tk.END, f"  🗑️ {op.item_name}{size_str}\n", 'item')
                    self.details_text.insert(tk.END, "\n")

                # Duplicate Images section (below)
                if duplicate_items:
                    self.details_text.insert(tk.END, f"Duplicate Images ({len(duplicate_items)}):\n", 'subheader')
                    # Group by duplicate_group_id
                    groups = {}
                    for op in duplicate_items:
                        gid = op.duplicate_group_id
                        if gid not in groups:
                            groups[gid] = []
                        groups[gid].append(op)

                    for group_id, group_items in groups.items():
                        self.details_text.insert(tk.END, f"  Group ({len(group_items)} images):\n", 'info')
                        for op in group_items:
                            size_str = f" (~{op.size_bytes // 1024}KB)" if op.size_bytes >= 1024 else f" (~{op.size_bytes}B)" if op.size_bytes > 0 else ""
                            refs_str = f" [{op.references_count} refs]" if op.references_count > 0 else ""
                            self.details_text.insert(tk.END, f"    🖼️ {op.item_name}{size_str}{refs_str}\n", 'item')
                        self.details_text.insert(tk.END, "\n")

                # Tips
                if unused_items:
                    self.details_text.insert(tk.END, "💡 Unused images are registered but never referenced.\n", 'tip')
                if duplicate_items:
                    self.details_text.insert(tk.END, "💡 Duplicates will be consolidated - references redirected.\n", 'tip')

        # Configure text tags
        self.details_text.tag_config('header', font=('Segoe UI', 11, 'bold'), foreground=colors['text_primary'])
        self.details_text.tag_config('subheader', font=('Segoe UI', 10, 'bold'), foreground=colors['text_primary'])
        self.details_text.tag_config('separator', foreground=colors['text_muted'])
        self.details_text.tag_config('item', font=('Segoe UI', 9), foreground=colors['text_primary'])
        self.details_text.tag_config('info', font=('Segoe UI', 9, 'italic'), foreground=colors['text_secondary'])
        self.details_text.tag_config('tip', font=('Segoe UI', 9), foreground=colors['info'])
        self.details_text.tag_config('summary', font=('Segoe UI', 9, 'bold'), foreground=colors['success'])

        self.details_text.config(state=tk.DISABLED)

    def _analyze_report(self):
        """Analyze the Power BI report for cleanup opportunities"""
        try:
            # Clean up any previous extracted temp directory
            if hasattr(self, '_current_validation') and hasattr(self, '_current_reader'):
                self._current_reader.cleanup_extracted_pbir(self._current_validation)

            # Validate input
            pbip_path = self.pbip_path_var.get().strip()
            if not pbip_path:
                return  # Button should be disabled, but safety check

            # Validate file using shared PBIPReader
            reader = PBIPReader()
            validation = reader.validate_report_file(pbip_path)
            if not validation['valid']:
                if validation.get('error_type') == 'not_pbir_format':
                    # Handle PBIR format error with Learn More button
                    self._show_pbir_format_error(validation['error'])
                    return
                else:
                    raise ValueError(validation['error'])

            # Store validation for cleanup and report_dir usage
            self._current_validation = validation
            self._current_reader = reader
            report_dir = validation['report_dir']

            # Clear previous results
            if hasattr(self, 'results_frame') and self.results_frame:
                self.results_frame.pack_forget()

            # Show immediate feedback
            self.log_message("🚀 Starting report analysis...")
            self.log_message(f"📁 Analyzing: {pbip_path}")

            # Run analysis in background
            def progress_callback(percent: int, message: str):
                """Thread-safe progress update callback"""
                self.frame.after(0, lambda: self.update_progress(percent, message))

            def run_analysis():
                analyzer = ReportAnalyzer(
                    logger_callback=self.log_message,
                    progress_callback=progress_callback
                )
                return analyzer.analyze_pbip_report(pbip_path, report_dir)

            def on_success(result):
                analysis_data, opportunities = result
                self.current_opportunities = opportunities
                # Store duplicate image data for config dialog
                self._duplicate_groups = analysis_data.get('images', {}).get('duplicate_groups', [])
                # Use validated report_dir (may be extracted temp dir for embedded PBIR)
                self._report_dir = report_dir
                self._duplicate_image_selections.clear()  # Reset selections for new analysis
                self._show_analysis_results(analysis_data, opportunities)
                self.update_progress(100, "Analysis complete!")
                self.log_message("✅ Analysis completed successfully!")

                # Warn about PBIX limitation (can analyze but not modify)
                if pbip_path.lower().endswith('.pbix'):
                    self.log_message("")
                    self.log_message("⚠️ NOTE: PBIX files can be analyzed but NOT modified.")
                    self.log_message("   Power BI validates PBIX content integrity and doesn't easily")
                    self.log_message("   allow external modifications - save as .pbip to clean.")

                # Note: Don't cleanup here - report_dir needed for cleanup actions
                # Cleanup happens when new analysis starts or tab closes

            def on_error(error):
                self.log_message(f"❌ Analysis failed: {error}")
                self.show_error("Analysis Failed", f"Report analysis failed:\n\n{error}")
                # Clean up extracted temp directory on error
                reader.cleanup_extracted_pbir(validation)

            # Show progress and run in background
            self.update_progress(10, "Starting analysis...", True)
            self.run_in_background(run_analysis, on_success, on_error)
            
        except Exception as e:
            self.log_message(f"❌ Validation error: {e}")
            self.show_error("Validation Error", str(e))

    def _show_pbir_format_error(self, error_message: str):
        """Show PBIR format error with Learn More button"""
        # Strip the marker prefix for display
        display_message = error_message.replace("PBIR_FORMAT_ERROR:", "")
        self.log_message(f"❌ Validation error: {display_message}")

        # Show dialog with Learn More button
        result = ThemedMessageBox.show(
            self.frame.winfo_toplevel(),
            "PBIR Format Required",
            f"{display_message}",
            msg_type="warning",
            buttons=["OK", "Learn More"]
        )

        # If user clicked Learn More, open the Microsoft documentation
        if result == "Learn More":
            import webbrowser
            webbrowser.open("https://learn.microsoft.com/en-us/power-bi/developer/projects/projects-report?tabs=v2%2Cdesktop#pbir-format")

    def _show_analysis_results(self, analysis_data: Dict[str, Any], opportunities: List[CleanupOpportunity]):
        """Display the analysis results by updating opportunity cards"""
        # Count opportunities by type
        theme_count = len([op for op in opportunities if op.item_type == 'theme'])
        visual_count = len([op for op in opportunities if op.item_type in ['custom_visual_build_pane', 'custom_visual_hidden']])
        bookmark_count = len([op for op in opportunities if op.item_type in ['bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group']])
        filter_count = sum(op.filter_count for op in opportunities if op.item_type == 'visual_filter')
        dax_query_count = len([op for op in opportunities if op.item_type == 'dax_query'])
        tmdl_script_count = len([op for op in opportunities if op.item_type == 'tmdl_script'])
        saved_scripts_count = dax_query_count + tmdl_script_count
        duplicate_image_count = len([op for op in opportunities if op.item_type == 'duplicate_image'])
        unused_image_count = len([op for op in opportunities if op.item_type == 'unused_image'])
        image_cleanup_count = duplicate_image_count + unused_image_count

        # Calculate sizes
        theme_size = sum(op.size_bytes for op in opportunities if op.item_type == 'theme')
        visual_size = sum(op.size_bytes for op in opportunities if op.item_type in ['custom_visual_build_pane', 'custom_visual_hidden'])
        dax_query_size = sum(op.size_bytes for op in opportunities if op.item_type == 'dax_query')
        tmdl_script_size = sum(op.size_bytes for op in opportunities if op.item_type == 'tmdl_script')
        saved_scripts_size = dax_query_size + tmdl_script_size
        duplicate_image_size = sum(op.size_bytes for op in opportunities if op.item_type == 'duplicate_image')
        unused_image_size = sum(op.size_bytes for op in opportunities if op.item_type == 'unused_image')
        image_cleanup_size = duplicate_image_size + unused_image_size

        # Update opportunity cards
        self.opportunity_cards['remove_themes'].update_results(theme_count, theme_size)
        self.opportunity_cards['remove_custom_visuals'].update_results(visual_count, visual_size)
        self.opportunity_cards['remove_bookmarks'].update_results(bookmark_count)
        self.opportunity_cards['hide_visual_filters'].update_results(filter_count)
        self.opportunity_cards['remove_saved_scripts'].update_results(saved_scripts_count, saved_scripts_size)
        self.opportunity_cards['image_cleanup'].update_results(image_cleanup_count, image_cleanup_size)

        # Hide hint text after analysis
        if hasattr(self, 'cards_hint_label') and self.cards_hint_label:
            self.cards_hint_label.pack_forget()

        # Update clean button state
        self._update_clean_button_state()

        # Log summary
        total = theme_count + visual_count + bookmark_count + filter_count + saved_scripts_count + image_cleanup_count
        if total > 0:
            self.log_message(f"\n📊 Found {total} cleanup opportunities:")
            if theme_count > 0:
                self.log_message(f"   🎨 {theme_count} unused themes (~{self._format_bytes(theme_size)})")
            if visual_count > 0:
                self.log_message(f"   📊 {visual_count} unused visuals (~{self._format_bytes(visual_size)})")
            if bookmark_count > 0:
                self.log_message(f"   📖 {bookmark_count} unused bookmarks")
            if filter_count > 0:
                self.log_message(f"   🎯 {filter_count} visual filters can be hidden")
            if saved_scripts_count > 0:
                self.log_message(f"   📝 {saved_scripts_count} saved scripts (~{self._format_bytes(saved_scripts_size)})")
                if dax_query_count > 0:
                    self.log_message(f"      - {dax_query_count} DAX queries")
                if tmdl_script_count > 0:
                    self.log_message(f"      - {tmdl_script_count} TMDL scripts")
            if image_cleanup_count > 0:
                self.log_message(f"   🖼️ {image_cleanup_count} image issues (~{self._format_bytes(image_cleanup_size)})")
                if duplicate_image_count > 0:
                    self.log_message(f"      - {duplicate_image_count} duplicate images (~{self._format_bytes(duplicate_image_size)} savings)")
                if unused_image_count > 0:
                    self.log_message(f"      - {unused_image_count} unused images (~{self._format_bytes(unused_image_size)})")
            self.log_message("\n💡 Select categories above and click CLEAN SELECTED")
        else:
            self.log_message("\n🎉 No cleanup needed! Your report is already optimized.")
    
    def _remove_selected_items(self):
        """Remove the selected themes, custom visuals, bookmarks, and/or saved scripts"""
        try:
            # Block cleanup for ALL PBIX files - Power BI validates content integrity
            # and ANY modification causes corruption. Only PBIP files can be modified.
            pbip_path = self.pbip_path_var.get().strip()
            if pbip_path.lower().endswith('.pbix'):
                self.show_error(
                    "PBIX Modification Not Supported",
                    "PBIX files cannot be modified by external tools.\n\n"
                    "Power BI validates the integrity of PBIX content, which prevents "
                    "any modifications from working correctly.\n\n"
                    "To use the cleanup feature:\n"
                    "1. Open this file in Power BI Desktop\n"
                    "2. Save As > Power BI Project (.pbip)\n"
                    "3. Run cleanup on the .pbip file\n\n"
                    "The analysis results above are still valid for understanding what could be cleaned."
                )
                return

            # Check if any options are selected
            remove_themes = self.cleanup_options['remove_themes'].get()
            remove_visuals = self.cleanup_options['remove_custom_visuals'].get()
            remove_bookmarks = self.cleanup_options['remove_bookmarks'].get()
            hide_filters = self.cleanup_options['hide_visual_filters'].get()
            remove_saved_scripts = self.cleanup_options['remove_saved_scripts'].get()
            image_cleanup = self.cleanup_options['image_cleanup'].get()
            # Combined toggle controls both duplicate consolidation and unused image removal
            consolidate_duplicates = image_cleanup
            remove_unused_images = image_cleanup

            if not remove_themes and not remove_visuals and not remove_bookmarks and not hide_filters and not remove_saved_scripts and not image_cleanup:
                self.show_warning("No Options Selected", "Please select at least one cleanup option.")
                return

            # Get items to remove based on options
            themes_to_remove = []
            visuals_to_remove = []  # Contains CleanupOpportunity objects
            bookmarks_to_remove = []  # Contains CleanupOpportunity objects
            filters_to_hide = False  # Boolean flag
            dax_queries_to_remove = []  # Contains CleanupOpportunity objects
            tmdl_scripts_to_remove = []  # Contains CleanupOpportunity objects
            duplicate_images_to_consolidate = []  # Contains CleanupOpportunity objects
            unused_images_to_remove = []  # Contains CleanupOpportunity objects

            for opportunity in self.current_opportunities:
                if opportunity.item_type == 'theme' and remove_themes:
                    themes_to_remove.append(opportunity.item_name)
                elif opportunity.item_type in ['custom_visual_build_pane', 'custom_visual_hidden'] and remove_visuals:
                    visuals_to_remove.append(opportunity)  # Pass the whole opportunity object
                elif opportunity.item_type in ['bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group'] and remove_bookmarks:
                    bookmarks_to_remove.append(opportunity)  # Pass the whole opportunity object
                elif opportunity.item_type == 'visual_filter' and hide_filters:
                    filters_to_hide = True  # Just set flag to hide all visual filters
                elif opportunity.item_type == 'dax_query' and remove_saved_scripts:
                    dax_queries_to_remove.append(opportunity)
                elif opportunity.item_type == 'tmdl_script' and remove_saved_scripts:
                    tmdl_scripts_to_remove.append(opportunity)
                elif opportunity.item_type == 'duplicate_image' and consolidate_duplicates:
                    duplicate_images_to_consolidate.append(opportunity)
                elif opportunity.item_type == 'unused_image' and remove_unused_images:
                    unused_images_to_remove.append(opportunity)

            if not themes_to_remove and not visuals_to_remove and not bookmarks_to_remove and not filters_to_hide and not dax_queries_to_remove and not tmdl_scripts_to_remove and not duplicate_images_to_consolidate and not unused_images_to_remove:
                self.show_warning("No Items to Process", "No items to remove or hide based on current selection.")
                return
            
            # Build confirmation message
            confirm_parts = []
            if themes_to_remove:
                theme_list = "\n".join([f"  🎨 {name}" for name in themes_to_remove])
                confirm_parts.append(f"THEMES ({len(themes_to_remove)}):\n{theme_list}")
            
            if visuals_to_remove:
                # Group visuals by type for better display
                build_pane = [v for v in visuals_to_remove if v.item_type == 'custom_visual_build_pane']
                hidden = [v for v in visuals_to_remove if v.item_type == 'custom_visual_hidden']
                
                visual_parts = []
                if build_pane:
                    build_list = "\n".join([f"  🔮 {v.item_name}" for v in build_pane])
                    visual_parts.append(f"BUILD PANE VISUALS ({len(build_pane)}):\n{build_list}")
                
                if hidden:
                    hidden_list = "\n".join([f"  😫 {v.item_name}" for v in hidden])
                    visual_parts.append(f"HIDDEN VISUALS ({len(hidden)}):\n{hidden_list}")
                
                if visual_parts:
                    confirm_parts.append("\n\n".join(visual_parts))
            
            if bookmarks_to_remove:
                # Group bookmarks by type for better display
                guaranteed = [b for b in bookmarks_to_remove if b.item_type == 'bookmark_guaranteed_unused']
                likely = [b for b in bookmarks_to_remove if b.item_type == 'bookmark_likely_unused']
                empty_groups = [b for b in bookmarks_to_remove if b.item_type == 'bookmark_empty_group']
                
                bookmark_parts = []
                if guaranteed:
                    guaranteed_list = "\n".join([f"  ✅ {b.item_name} (page missing)" for b in guaranteed])
                    bookmark_parts.append(f"GUARANTEED UNUSED ({len(guaranteed)}):  \n{guaranteed_list}")
                
                if likely:
                    likely_list = "\n".join([f"  ⚠️ {b.item_name} (no navigation found)" for b in likely])
                    bookmark_parts.append(f"LIKELY UNUSED ({len(likely)}):  \n{likely_list}")
                
                if empty_groups:
                    empty_list = "\n".join([f"  📁 {b.item_name} (empty group)" for b in empty_groups])
                    bookmark_parts.append(f"EMPTY GROUPS ({len(empty_groups)}):  \n{empty_list}")
                
                if bookmark_parts:
                    confirm_parts.append("\n\n".join(bookmark_parts))
            
            if filters_to_hide:
                # Get filter count from opportunity
                filter_opportunity = next((op for op in self.current_opportunities if op.item_type == 'visual_filter'), None)
                filter_count = filter_opportunity.filter_count if filter_opportunity else 0
                filter_text = f"VISUAL FILTERS ({filter_count}):\n  🎯 All visual-level filters will be hidden (not removed)"
                confirm_parts.append(filter_text)

            if dax_queries_to_remove or tmdl_scripts_to_remove:
                scripts_parts = []
                if dax_queries_to_remove:
                    dax_list = "\n".join([f"  📝 {d.item_name}" for d in dax_queries_to_remove])
                    scripts_parts.append(f"DAX Queries ({len(dax_queries_to_remove)}):\n{dax_list}")
                if tmdl_scripts_to_remove:
                    tmdl_list = "\n".join([f"  📜 {t.item_name}" for t in tmdl_scripts_to_remove])
                    scripts_parts.append(f"TMDL Scripts ({len(tmdl_scripts_to_remove)}):\n{tmdl_list}")
                total_scripts = len(dax_queries_to_remove) + len(tmdl_scripts_to_remove)
                confirm_parts.append(f"SAVED SCRIPTS ({total_scripts}):\n" + "\n".join(scripts_parts))

            if duplicate_images_to_consolidate:
                # Group by duplicate_group_id for display
                groups = {}
                for img in duplicate_images_to_consolidate:
                    gid = img.duplicate_group_id
                    if gid not in groups:
                        groups[gid] = []
                    groups[gid].append(img)

                dup_parts = []
                for group_id, group_items in groups.items():
                    group_list = "\n".join([f"    🖼️ {img.item_name}" for img in group_items])
                    dup_parts.append(f"  Group ({len(group_items)} duplicates):\n{group_list}")

                confirm_parts.append(f"DUPLICATE IMAGES ({len(duplicate_images_to_consolidate)}):\n" + "\n".join(dup_parts) + "\n  References will be redirected, duplicates removed.")

            if unused_images_to_remove:
                unused_list = "\n".join([f"  🗑️ {img.item_name}" for img in unused_images_to_remove])
                confirm_parts.append(f"UNUSED IMAGES ({len(unused_images_to_remove)}):\n{unused_list}")

            # Add special warning for likely unused bookmarks
            warning_note = ""
            if any(b.item_type == 'bookmark_likely_unused' for b in bookmarks_to_remove):
                warning_note = "\n\n📖 NOTE: 'Likely unused' bookmarks could still be accessed via the bookmark pane in Power BI Service, even without navigation buttons."
            
            if filters_to_hide:
                filter_note = "\n\n🎯 NOTE: Visual filters will be hidden (not deleted) and can be shown again later if needed."
                warning_note += filter_note
            
            action_text = "remove/hide these items" if filters_to_hide else "remove these items"
            confirm_message = (f"Are you sure you want to {action_text}?\n\n" +
                             "\n\n".join(confirm_parts) +
                             warning_note +
                             f"\n\n⚠️ This action will modify your report.\n"
                             f"💡 Consider creating a backup before proceeding.")
            
            # Use scrollable dialog with max height of 450px for long confirmation messages
            if not self.ask_yes_no("Confirm Cleanup", confirm_message, max_content_height=450):
                return
            
            # Perform removal
            pbip_path = self.pbip_path_var.get().strip()

            def run_removal():
                # Use the stored report_dir from analysis (may be extracted temp dir for PBIX)
                report_dir = getattr(self, '_report_dir', None)

                # For PBIX files, content_dir is the parent of report_dir (temp dir with DAXQueries/TMDLScripts)
                # For PBIP files, don't pass content_dir so it looks in .SemanticModel folder instead
                is_pbix = pbip_path.lower().endswith('.pbix')
                content_dir = report_dir.parent if (is_pbix and report_dir) else None

                return self.cleanup_engine.remove_unused_items(
                    pbip_path,
                    themes_to_remove if remove_themes else None,
                    visuals_to_remove if remove_visuals else None,
                    bookmarks_to_remove if remove_bookmarks else None,
                    hide_visual_filters=filters_to_hide,
                    dax_queries_to_remove=dax_queries_to_remove if remove_saved_scripts else None,
                    tmdl_scripts_to_remove=tmdl_scripts_to_remove if remove_saved_scripts else None,
                    duplicate_groups=self._duplicate_groups if consolidate_duplicates and duplicate_images_to_consolidate else None,
                    duplicate_selections=self._duplicate_image_selections if consolidate_duplicates else None,
                    unused_images_to_remove=unused_images_to_remove if remove_unused_images else None,
                    create_backup=False,  # No automatic backup
                    report_dir=report_dir,
                    content_dir=content_dir
                )

            def on_success(results: List[RemovalResult]):
                self._handle_removal_results(results)

            def on_error(error):
                self.show_error("Cleanup Failed", f"Cleanup operation failed:\n\n{error}")

            # Progress steps for removal
            progress_steps = [
                ("Removing themes...", 10) if remove_themes else None,
                ("Removing custom visuals...", 20) if remove_visuals else None,
                ("Removing bookmarks...", 30) if remove_bookmarks else None,
                ("Hiding visual filters...", 40) if filters_to_hide else None,
                ("Removing saved scripts...", 50) if remove_saved_scripts and (dax_queries_to_remove or tmdl_scripts_to_remove) else None,
                ("Consolidating duplicate images...", 65) if consolidate_duplicates and duplicate_images_to_consolidate else None,
                ("Removing unused images...", 80) if remove_unused_images and unused_images_to_remove else None,
                ("Cleaning up references...", 95),
                ("Cleanup complete!", 100)
            ]
            progress_steps = [step for step in progress_steps if step is not None]

            self.run_in_background(run_removal, on_success, on_error, progress_steps)
            
        except Exception as e:
            self.show_error("Cleanup Error", str(e))
    
    def _handle_removal_results(self, results: List[RemovalResult]):
        """Handle the results of cleanup operation"""
        successful_removals = [r for r in results if r.success]
        failed_removals = [r for r in results if not r.success]
        
        # Calculate total bytes freed
        total_bytes_freed = sum(r.bytes_freed for r in successful_removals)
        
        # Log results by type
        theme_results = [r for r in successful_removals if r.item_type == 'theme']
        visual_results = [r for r in successful_removals if r.item_type == 'custom_visual']
        bookmark_results = [r for r in successful_removals if r.item_type == 'bookmark']
        filter_results = [r for r in successful_removals if r.item_type == 'visual_filter']
        dax_query_results = [r for r in successful_removals if r.item_type == 'dax_query']
        tmdl_script_results = [r for r in successful_removals if r.item_type == 'tmdl_script']
        duplicate_image_results = [r for r in successful_removals if r.item_type == 'duplicate_image']
        unused_image_results = [r for r in successful_removals if r.item_type == 'unused_image']

        self.log_message(f"\n🎯 CLEANUP RESULTS:")

        if theme_results:
            self.log_message(f"🎨 Themes removed: {len(theme_results)}")
            for result in theme_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}")

        if visual_results:
            self.log_message(f"🔮 Custom visuals removed: {len(visual_results)}")
            for result in visual_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}")

        if bookmark_results:
            self.log_message(f"📖 Bookmarks removed: {len(bookmark_results)}")
            for result in bookmark_results:
                self.log_message(f"  ✅ {result.item_name}")

        if filter_results:
            for result in filter_results:
                self.log_message(f"🎯 Visual filters hidden: {result.filters_hidden}")
                self.log_message(f"  ✅ {result.item_name}")

        if dax_query_results:
            self.log_message(f"📝 DAX queries removed: {len(dax_query_results)}")
            for result in dax_query_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}")

        if tmdl_script_results:
            self.log_message(f"📜 TMDL scripts removed: {len(tmdl_script_results)}")
            for result in tmdl_script_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}")

        if duplicate_image_results:
            total_refs = sum(r.references_updated for r in duplicate_image_results)
            self.log_message(f"🖼️ Duplicate images consolidated: {len(duplicate_image_results)}")
            for result in duplicate_image_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                refs_str = f", {result.references_updated} refs updated" if result.references_updated > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}{refs_str}")

        if unused_image_results:
            self.log_message(f"🗑️ Unused images removed: {len(unused_image_results)}")
            for result in unused_image_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}")

        if failed_removals:
            self.log_message(f"❌ Failed removals: {len(failed_removals)}")
            for result in failed_removals:
                self.log_message(f"  ❌ {result.item_name}: {result.error_message}")
        
        self.log_message(f"💾 Total space freed: {self._format_bytes(total_bytes_freed)}")
        
        # Show completion message
        if successful_removals and not failed_removals:
            self.show_info("Cleanup Complete", 
                         f"Successfully cleaned up {len(successful_removals)} items!\n\n"
                         f"Space freed: {self._format_bytes(total_bytes_freed)}\n"
                         f"Your report has been cleaned up.")
        elif successful_removals and failed_removals:
            self.show_warning("Partial Success", 
                            f"Removed {len(successful_removals)} items successfully.\n"
                            f"{len(failed_removals)} items could not be removed.\n\n"
                            f"Check the log for details.")
        else:
            self.show_error("Cleanup Failed", 
                          f"Could not clean any of the selected items.\n\n"
                          f"Check the log for details.")
        
        # Refresh the analysis to show updated state
        self._analyze_report()

    def _format_bytes(self, bytes_count: int) -> str:
        """Format bytes into human readable format"""
        if bytes_count == 0:
            return "0 B"
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if bytes_count < 1024:
                return f"{bytes_count:.1f} {unit}"
            bytes_count /= 1024
        return f"{bytes_count:.1f} TB"
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        colors = self._theme_manager.colors

        # Clear input
        self.pbip_path_var.set("")

        # Clear opportunities
        self.current_opportunities.clear()

        # Clear image analysis state
        self._duplicate_groups.clear()
        self._report_dir = None
        self._duplicate_image_selections.clear()

        # Reset selected card state
        self._selected_card_key = None

        # Reset all opportunity cards to initial state
        for card in self.opportunity_cards.values():
            card.reset()

        # Show hint text again with current theme colors
        if hasattr(self, 'cards_hint_label') and self.cards_hint_label:
            self.cards_hint_label.config(bg=colors['background'], fg=colors['text_secondary'])
            self.cards_hint_label.pack(pady=(20, 5))

        # Disable clean button
        if hasattr(self, 'clean_button') and self.clean_button:
            self.clean_button.set_enabled(False)

        # Clear log and show welcome
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)

        # Reset details panel to placeholder
        self._show_details_placeholder()

        # Hide progress bar if visible
        if hasattr(self, 'progress_components') and self.progress_components:
            self.progress_components['frame'].pack_forget()

        # Force a UI update
        self.frame.update_idletasks()

        # Show welcome message
        self._show_welcome_message()

    def _handle_theme_change(self, theme: str):
        """Override to update card container backgrounds on theme change"""
        # Call base class first (may fail silently, so wrap in try)
        try:
            super()._handle_theme_change(theme)
        except Exception:
            pass

        # Get current theme colors
        colors = self._theme_manager.colors
        bg_color = colors.get('background', '#ffffff')

        # Update cards container backgrounds - these are critical for theme switching
        try:
            if hasattr(self, '_cards_content_frame') and self._cards_content_frame:
                self._cards_content_frame.config(bg=bg_color)
            if hasattr(self, '_cards_container') and self._cards_container:
                self._cards_container.config(bg=bg_color)
            if hasattr(self, 'cards_hint_label') and self.cards_hint_label:
                self.cards_hint_label.config(bg=bg_color, fg=colors['text_secondary'])
        except Exception:
            pass

        # SplitLogSection handles its own theme updates internally

        # Update details text widget colors (ModernScrolledText = tk.Frame with inner _text)
        # Use section_bg to match surrounding frame (eliminates visible border in dark mode)
        # Must update BOTH the inner text widget AND the outer Frame wrapper
        try:
            if hasattr(self, 'details_text') and self.details_text:
                # Update inner text widget
                self.details_text.config(
                    bg=colors['section_bg'], fg=colors['text_primary'],
                    highlightcolor=colors['border'], highlightbackground=colors['border'],
                    selectbackground=colors.get('selection_bg', colors['accent'])
                )
                # Update outer Frame wrapper (ModernScrolledText.config only updates inner text!)
                tk.Frame.configure(self.details_text, bg=colors['section_bg'],
                                   highlightcolor=colors['border'], highlightbackground=colors['border'])
                # Re-apply the current view to update tag colors
                if hasattr(self, '_selected_card_key') and self._selected_card_key:
                    self._show_card_details(self._selected_card_key)
                else:
                    self._show_details_placeholder()
        except Exception:
            pass

        # Update log_text outer Frame background (same issue as details_text)
        try:
            if hasattr(self, 'log_text') and self.log_text:
                # Update outer Frame wrapper for log text
                tk.Frame.configure(self.log_text, bg=colors['section_bg'],
                                   highlightcolor=colors['border'], highlightbackground=colors['border'])
        except Exception:
            pass

        # Update bottom action buttons canvas_bg for proper corner rounding on outer background
        try:
            outer_canvas_bg = colors['section_bg']
            if hasattr(self, 'clean_button') and self.clean_button:
                self.clean_button.update_colors(
                    bg=colors['button_primary'],
                    hover_bg=colors['button_primary_hover'],
                    pressed_bg=colors['button_primary_pressed'],
                    fg='#ffffff',
                    disabled_bg=colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc'),
                    disabled_fg=colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8'),
                    canvas_bg=outer_canvas_bg
                )
            if hasattr(self, 'reset_button') and self.reset_button:
                self.reset_button.update_colors(
                    bg=colors['button_secondary'],
                    hover_bg=colors['button_secondary_hover'],
                    pressed_bg=colors['button_secondary_pressed'],
                    fg=colors['text_primary'],
                    canvas_bg=outer_canvas_bg
                )
        except Exception:
            pass

    def show_help_dialog(self) -> None:
        """Show help dialog specific to report cleanup - two-column layout with modern styling"""
        from core.ui_base import RoundedButton
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Consistent help dialog background for all tools
        help_bg = colors['background']

        def create_help_content(help_window):
            # Main container - use tk.Frame with explicit bg for consistency
            main_frame = tk.Frame(help_window, bg=help_bg, padx=20, pady=15)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Header - centered for middle-out design
            tk.Label(main_frame, text="Report Cleanup - Help",
                     font=('Segoe UI', 16, 'bold'),
                     bg=help_bg,
                     fg=colors['title_color']).pack(anchor=tk.CENTER, pady=(0, 15))

            # ===== IMPORTANT SAFETY NOTES - Orange Warning Box =====
            warning_frame = tk.Frame(main_frame, bg=help_bg)
            warning_frame.pack(fill=tk.X, pady=(0, 15))

            warning_bg = '#d97706'  # Orange for warning
            warning_container = tk.Frame(warning_frame, bg=warning_bg,
                                       padx=15, pady=10, relief='flat', borderwidth=0)
            warning_container.pack(fill=tk.X)

            warning_text_color = '#ffffff'
            tk.Label(warning_container, text="IMPORTANT DISCLAIMERS & REQUIREMENTS",
                     font=('Segoe UI', 12, 'bold'),
                     bg=warning_bg, fg=warning_text_color).pack(anchor=tk.W)

            warnings = [
                "PBIX: Analysis only - Power BI validates content integrity",
                "       and doesn't easily allow external modifications",
                "PBIP: Full support for analysis AND cleanup operations",
                "NOT officially supported by Microsoft - use at your own discretion",
                "Always backup your report before cleanup"
            ]

            for warning in warnings:
                tk.Label(warning_container, text=f"  {warning}", font=('Segoe UI', 10),
                         bg=warning_bg, fg=warning_text_color).pack(anchor=tk.W, pady=1)

            # ===== Two-Column Layout for Help Sections =====
            columns_frame = tk.Frame(main_frame, bg=help_bg)
            columns_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            columns_frame.columnconfigure(0, weight=1)
            columns_frame.columnconfigure(1, weight=1)

            # LEFT COLUMN sections (emojis removed)
            left_sections = [
                ("Quick Start", [
                    "1. Select a Power BI report file",
                    "2. Click 'ANALYZE REPORT' to scan",
                    "3. Choose items to remove",
                    "4. Click 'CLEAN SELECTED' to clean"
                ]),
                ("Detection Logic", [
                    "Compares active theme with available files",
                    "Scans all pages for custom visual usage",
                    "Identifies hidden visuals in CustomVisuals",
                    "Analyzes bookmark navigation & pages"
                ])
            ]

            # RIGHT COLUMN sections (emojis removed)
            right_sections = [
                ("What This Tool Detects", [
                    "Unused themes in BaseThemes",
                    "Custom visuals in build pane but unused",
                    "Hidden custom visuals (not in pane)",
                    "Unused bookmarks (missing pages)",
                    "Visual-level filters that can be hidden"
                ]),
                ("File Requirements", [
                    "PBIP: Full support (analyze + cleanup)",
                    "PBIX: Analysis only (no modifications)",
                    "Must be saved in PBIR format",
                    "Legacy report.json NOT supported"
                ])
            ]

            # Create left column
            left_column = tk.Frame(columns_frame, bg=help_bg)
            left_column.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 10))

            for title, items in left_sections:
                section_frame = tk.Frame(left_column, bg=help_bg)
                section_frame.pack(fill=tk.X, pady=(0, 12), anchor=tk.W)

                tk.Label(section_frame, text=title, font=('Segoe UI', 12, 'bold'),
                         fg=colors['title_color'], bg=help_bg).pack(anchor=tk.W)

                for item in items:
                    tk.Label(section_frame, text=f"  {item}", font=('Segoe UI', 10),
                            fg=colors['text_primary'], bg=help_bg).pack(anchor=tk.W, pady=1)

            # Create right column
            right_column = tk.Frame(columns_frame, bg=help_bg)
            right_column.grid(row=0, column=1, sticky=(tk.N, tk.W, tk.E), padx=(10, 0))

            for title, items in right_sections:
                section_frame = tk.Frame(right_column, bg=help_bg)
                section_frame.pack(fill=tk.X, pady=(0, 12), anchor=tk.W)

                tk.Label(section_frame, text=title, font=('Segoe UI', 12, 'bold'),
                         fg=colors['title_color'], bg=help_bg).pack(anchor=tk.W)

                for item in items:
                    tk.Label(section_frame, text=f"  {item}", font=('Segoe UI', 10),
                            fg=colors['text_primary'], bg=help_bg).pack(anchor=tk.W, pady=1)

            # ===== Modern Close Button at Bottom =====
            button_frame = tk.Frame(main_frame, bg=help_bg)
            button_frame.pack(fill=tk.X, pady=(10, 0), side=tk.BOTTOM)

            close_btn = RoundedButton(
                button_frame,
                text="Close",
                command=help_window.destroy,
                bg=colors['button_primary'],
                fg='#ffffff',
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                height=36, radius=8,  # Auto-width, dialog standard
                font=('Segoe UI', 10),
                canvas_bg=help_bg
            )
            close_btn.pack(pady=(5, 0))

        # Create help window - wider for two-column layout
        help_window = tk.Toplevel(self.main_app.root)
        help_window.withdraw()
        help_window.title("Report Cleanup - Help")
        help_window.geometry("720x580")  # Wider for two columns, shorter height
        help_window.resizable(False, False)
        help_window.transient(self.main_app.root)
        help_window.grab_set()
        help_window.configure(bg=help_bg)

        # Set AE favicon icon
        try:
            help_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Create content
        create_help_content(help_window)

        # Bind escape key
        help_window.bind('<Escape>', lambda event: help_window.destroy())

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

        # Set dark/light title bar BEFORE showing window to prevent white flash
        help_window.update()
        self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled
        help_window.focus_force()
    
    def _show_welcome_message(self):
        """Show welcome message"""
        self.log_message("🧹 Welcome to Report Cleanup!")
        self.log_message("=" * 60)
        self.log_message("This tool helps you clean up your Power BI reports by:")
        self.log_message("• Detecting unused themes in BaseThemes and RegisteredResources")
        self.log_message("• Finding custom visuals that aren't used in any pages")
        self.log_message("• Identifying 'hidden' custom visuals taking up space")
        self.log_message("• Detecting unused bookmarks (missing pages or no navigation)")
        self.log_message("• Hiding visual-level filters to clean up the interface")
        self.log_message("• Providing detailed analysis and removal reports")
        self.log_message("")
        self.log_message("🎯 Start by selecting your Power BI file and clicking 'ANALYZE REPORTS'")
        self.log_message("💡 Consider backing up your report before making changes")
        self.log_message("")
