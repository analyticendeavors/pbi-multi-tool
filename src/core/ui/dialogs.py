"""
Dialog Components - Tooltip, MessageBox, and Input Dialog
Built by Reid Havens of Analytic Endeavors

Theme-aware dialog components that follow the app's dark/light mode.
"""

import tkinter as tk
import io
from pathlib import Path
from typing import List

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


class Tooltip:
    """
    Theme-aware tooltip that appears on hover.
    Follows the app's dark/light mode styling.
    Supports static text or callable that returns text dynamically.
    """

    def __init__(self, widget, text, delay: int = 250):
        """
        Args:
            widget: The widget to attach tooltip to
            text: Either a string or a callable that returns a string
            delay: Delay in ms before showing tooltip
        """
        self.widget = widget
        self._text = text  # Can be str or callable
        self.delay = delay
        self._tooltip_window = None
        self._tooltip_id = None
        self._theme_manager = get_theme_manager()

        # Bind events
        self.widget.bind('<Enter>', self._on_enter, add='+')
        self.widget.bind('<Leave>', self._on_leave, add='+')
        self.widget.bind('<ButtonPress>', self._on_leave, add='+')

    @property
    def text(self):
        """Get the tooltip text, evaluating callable if needed"""
        if callable(self._text):
            return self._text()
        return self._text

    @text.setter
    def text(self, value):
        """Set the tooltip text (can be str or callable)"""
        self._text = value

    def _on_enter(self, event=None):
        """Schedule tooltip to appear after delay"""
        self._cancel_scheduled()
        # Store mouse position for tooltip placement
        self._mouse_x = event.x_root if event else self.widget.winfo_rootx()
        self._mouse_y = event.y_root if event else self.widget.winfo_rooty()
        self._tooltip_id = self.widget.after(self.delay, self._show_tooltip)

    def _on_leave(self, event=None):
        """Hide tooltip and cancel any scheduled show"""
        self._cancel_scheduled()
        self._hide_tooltip()

    def _cancel_scheduled(self):
        """Cancel any scheduled tooltip show"""
        if self._tooltip_id:
            self.widget.after_cancel(self._tooltip_id)
            self._tooltip_id = None

    def _show_tooltip(self):
        """Show the tooltip near the widget"""
        if self._tooltip_window:
            return

        colors = self._theme_manager.colors

        # Create tooltip window
        self._tooltip_window = tk.Toplevel(self.widget)
        self._tooltip_window.wm_overrideredirect(True)
        self._tooltip_window.wm_attributes('-topmost', True)

        # Theme-aware colors
        bg_color = colors.get('card_surface', '#ffffff')
        fg_color = colors.get('text_primary', '#1e1e1e')
        border_color = colors.get('border', '#e0e0e0')

        # Create frame with border
        frame = tk.Frame(self._tooltip_window, bg=border_color, padx=1, pady=1)
        frame.pack(fill=tk.BOTH, expand=True)

        label = tk.Label(
            frame,
            text=self.text,
            background=bg_color,
            foreground=fg_color,
            font=('Segoe UI', 9),
            padx=8,
            pady=4
        )
        label.pack()

        # Position tooltip near the mouse cursor (slightly below and to the right)
        x = getattr(self, '_mouse_x', self.widget.winfo_rootx()) + 15
        y = getattr(self, '_mouse_y', self.widget.winfo_rooty()) + 20

        self._tooltip_window.wm_geometry(f"+{x}+{y}")

    def _hide_tooltip(self):
        """Hide and destroy the tooltip window"""
        if self._tooltip_window:
            self._tooltip_window.destroy()
            self._tooltip_window = None

    def update_text(self, text: str):
        """Update tooltip text"""
        self.text = text


class ThemedMessageBox:
    """
    Theme-aware message box that follows the app's dark/light mode.
    Replaces native tkinter messagebox for consistent styling.
    """

    @staticmethod
    def show(parent, title: str, message: str, msg_type: str = "info", buttons: List[str] = None,
             custom_icon: str = None, checkbox_text: str = None, checkbox_default: bool = True,
             checkbox_align: str = "left", checkbox2_text: str = None, checkbox2_default: bool = False,
             auto_close_seconds: int = None):
        """
        Show a themed message dialog.

        Args:
            parent: Parent window
            title: Dialog title
            message: Message to display
            msg_type: "info", "warning", "error", or "success"
            buttons: List of button labels (default: ["OK"])
            custom_icon: Optional custom icon filename (e.g., "hotswap.svg")
            checkbox_text: Optional text for first checkbox option
            checkbox_default: Default state of first checkbox (default: True)
            checkbox_align: Alignment of checkboxes - "left" or "right" (default: "left")
            checkbox2_text: Optional text for second checkbox option
            checkbox2_default: Default state of second checkbox (default: False)
            auto_close_seconds: If set, auto-close after this many seconds with countdown

        Returns:
            If no checkboxes: The label of the clicked button
            If one checkbox: Tuple of (button_label, checkbox_state)
            If two checkboxes: Tuple of (button_label, checkbox1_state, checkbox2_state)
        """
        # Import RoundedButton here to avoid circular imports
        from core.ui.buttons import RoundedButton

        theme_manager = get_theme_manager()
        colors = theme_manager.colors
        is_dark = theme_manager.is_dark

        # Default button
        if buttons is None:
            buttons = ["OK"]

        result = [None]

        # Create dialog
        dialog = tk.Toplevel(parent)
        dialog.withdraw()
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

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

        # Icon colors based on message type (fallback for canvas-drawn icons)
        icon_colors = {
            "info": colors['info'],
            "warning": colors['warning'],
            "error": colors['error'],
            "success": colors['success'],
        }
        icon_symbols = {
            "info": "i",
            "warning": "!",
            "error": "X",
            "success": "^",
        }
        # SVG icon filenames for each message type
        icon_svg_names = {
            "warning": "warning.svg",
            "error": "warning.svg",  # Use warning icon for errors too
            "info": "question.svg",
            "success": "box-checked.svg",
        }

        # Main container - no expand to prevent extra space distribution
        main_frame = tk.Frame(dialog, bg=colors['background'], padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH)

        # Content frame (icon + message)
        content_frame = tk.Frame(main_frame, bg=colors['background'])
        content_frame.pack(fill=tk.X, pady=(0, 15))

        # Try to load SVG icon, fall back to canvas-drawn icon
        icon_image = None
        icon_size = 40
        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"
        # Use custom icon if provided, otherwise use default for msg_type
        svg_filename = custom_icon if custom_icon else icon_svg_names.get(msg_type, "question.svg")
        svg_path = icons_dir / svg_filename

        if PIL_AVAILABLE and CAIROSVG_AVAILABLE and svg_path.exists():
            try:
                # Render SVG at 4x size for quality, then downscale
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=icon_size * 4,
                    output_height=icon_size * 4
                )
                image = Image.open(io.BytesIO(png_data))
                image = image.resize((icon_size, icon_size), Image.Resampling.LANCZOS)
                icon_image = ImageTk.PhotoImage(image)
            except Exception:
                icon_image = None

        if icon_image:
            # Use SVG icon
            icon_label = tk.Label(content_frame, image=icon_image, bg=colors['background'])
            icon_label.image = icon_image  # Keep reference
            icon_label.pack(side=tk.LEFT, padx=(0, 15))
        else:
            # Fallback to canvas-drawn icon circle
            icon_color = icon_colors.get(msg_type, colors['info'])
            icon_canvas = tk.Canvas(content_frame, width=icon_size, height=icon_size,
                                   bg=colors['background'], highlightthickness=0)
            icon_canvas.pack(side=tk.LEFT, padx=(0, 15))
            icon_canvas.create_oval(2, 2, icon_size-2, icon_size-2, fill=icon_color, outline="")
            icon_canvas.create_text(icon_size//2, icon_size//2, text=icon_symbols.get(msg_type, "i"),
                                   fill="#ffffff", font=('Segoe UI', 16, 'bold'))

        # Message text
        msg_label = tk.Label(content_frame, text=message,
                            bg=colors['background'], fg=colors['text_primary'],
                            font=('Segoe UI', 10), justify=tk.LEFT, wraplength=300)
        msg_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Optional checkboxes with SVG icons
        checkbox_var = tk.BooleanVar(value=checkbox_default)
        checkbox2_var = tk.BooleanVar(value=checkbox2_default)

        has_checkboxes = checkbox_text or checkbox2_text
        if has_checkboxes:
            # Options panel with subtle border for visual grouping
            border_color = '#2a2a3a' if is_dark else '#d0d0d0'

            options_frame = tk.Frame(main_frame, bg=colors['background'],
                                    highlightbackground=border_color,
                                    highlightcolor=border_color,
                                    highlightthickness=1)
            options_frame.pack(fill=tk.X, pady=(12, 0))

            # "Options" header
            options_header = tk.Label(options_frame, text="Options",
                                     font=('Segoe UI', 8),
                                     bg=colors['background'],
                                     fg=colors.get('text_secondary', '#888888' if is_dark else '#666666'))
            options_header.pack(anchor=tk.W, padx=10, pady=(6, 2))

            # Container for checkboxes inside the options panel
            checkbox_container = tk.Frame(options_frame, bg=colors['background'])
            checkbox_container.pack(fill=tk.X, padx=10, pady=(0, 8))

            # Load SVG checkbox icons once (reuse for both checkboxes)
            checkbox_size = 18
            checked_img = None
            unchecked_img = None

            # Select icons based on theme
            if is_dark:
                checked_svg = icons_dir / "box-checked-dark.svg"
                unchecked_svg = icons_dir / "box-dark.svg"
            else:
                checked_svg = icons_dir / "box-checked.svg"
                unchecked_svg = icons_dir / "box.svg"

            if PIL_AVAILABLE and CAIROSVG_AVAILABLE:
                try:
                    if checked_svg.exists():
                        png_data = cairosvg.svg2png(url=str(checked_svg), output_width=checkbox_size*2, output_height=checkbox_size*2)
                        img = Image.open(io.BytesIO(png_data)).resize((checkbox_size, checkbox_size), Image.Resampling.LANCZOS)
                        checked_img = ImageTk.PhotoImage(img)
                    if unchecked_svg.exists():
                        png_data = cairosvg.svg2png(url=str(unchecked_svg), output_width=checkbox_size*2, output_height=checkbox_size*2)
                        img = Image.open(io.BytesIO(png_data)).resize((checkbox_size, checkbox_size), Image.Resampling.LANCZOS)
                        unchecked_img = ImageTk.PhotoImage(img)
                except Exception:
                    pass

            # Helper to create SVG checkbox row
            def create_svg_checkbox(parent, text, var, is_first=True):
                row = tk.Frame(parent, bg=colors['background'])
                row.pack(anchor=tk.W, pady=(0 if is_first else 4, 0))

                icon_lbl = tk.Label(row, bg=colors['background'], cursor='hand2')
                icon_lbl.pack(side=tk.LEFT, padx=(0, 6))
                icon_lbl.checked_img = checked_img
                icon_lbl.unchecked_img = unchecked_img

                text_lbl = tk.Label(row, text=text, bg=colors['background'],
                                   fg=colors['text_primary'], font=('Segoe UI', 9), cursor='hand2')
                text_lbl.pack(side=tk.LEFT)

                def update_icon():
                    icon_lbl.configure(image=checked_img if var.get() else unchecked_img)

                def toggle(event=None):
                    var.set(not var.get())
                    update_icon()

                icon_lbl.bind('<Button-1>', toggle)
                text_lbl.bind('<Button-1>', toggle)
                update_icon()
                return row

            # Helper to create native checkbox row (fallback)
            def create_native_checkbox(parent, text, var, is_first=True):
                cb = tk.Checkbutton(
                    parent, text=text, variable=var,
                    bg=colors['background'], fg=colors['text_primary'],
                    activebackground=colors['background'], activeforeground=colors['text_primary'],
                    selectcolor=colors.get('surface', colors['background']), font=('Segoe UI', 9)
                )
                cb.pack(anchor=tk.W, pady=(0 if is_first else 4, 0))
                return cb

            # Create first checkbox
            if checkbox_text:
                if checked_img and unchecked_img:
                    create_svg_checkbox(checkbox_container, checkbox_text, checkbox_var, is_first=True)
                else:
                    create_native_checkbox(checkbox_container, checkbox_text, checkbox_var, is_first=True)

            # Create second checkbox
            if checkbox2_text:
                if checked_img and unchecked_img:
                    create_svg_checkbox(checkbox_container, checkbox2_text, checkbox2_var, is_first=False)
                else:
                    create_native_checkbox(checkbox_container, checkbox2_text, checkbox2_var, is_first=False)

        # Auto-close countdown label (above button, created first for layout order)
        countdown_label = None
        if auto_close_seconds and auto_close_seconds > 0:
            countdown_label = tk.Label(
                main_frame,
                text=f"Closing in {auto_close_seconds}...",
                font=('Segoe UI', 9, 'italic'),
                bg=colors['background'],
                fg=colors.get('text_muted', '#888888')
            )
            countdown_label.pack(pady=(0, 15))  # 15px gap above (from content_frame), 15px gap below

        # Button frame - centered, with spacing matching bottom padding
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        # No top padding if countdown label exists (it has its own padding)
        button_frame.pack(fill=tk.X, pady=(0 if countdown_label else 15, 0))

        # Inner frame for centered buttons
        button_inner = tk.Frame(button_frame, bg=colors['background'])
        button_inner.pack(expand=True)  # Center the buttons

        def on_button_click(btn_text):
            result[0] = btn_text
            dialog.destroy()

        # Button styling based on text (for PBIX dialog and others)
        # Warning/secondary action buttons (Swap Without Saving)
        warning_bg = '#4a3030' if is_dark else '#fef3c7'
        warning_hover = '#5a3838' if is_dark else '#fde68a'
        warning_pressed = '#3a2020' if is_dark else '#fcd34d'
        warning_fg = colors.get('warning', '#f59e0b' if is_dark else '#d97706')

        # Create buttons using RoundedButton for modern flat design
        canvas_bg = colors['background']
        for i, btn_text in enumerate(buttons):
            btn_lower = btn_text.lower()

            # Determine button style based on text
            if btn_lower in ('cancel', 'no', 'close'):
                # Cancel/No/Close buttons: secondary style like Rollback
                btn_bg = colors['button_secondary']
                btn_hover = colors['button_secondary_hover']
                btn_pressed = colors['button_secondary_pressed']
                btn_fg = colors['text_primary']
                btn_font = ('Segoe UI', 10)
            elif 'without' in btn_lower or 'only' in btn_lower:
                # "Without Saving" style: warning amber
                btn_bg = warning_bg
                btn_hover = warning_hover
                btn_pressed = warning_pressed
                btn_fg = warning_fg
                btn_font = ('Segoe UI', 10)
            elif i == 0:
                # First button is primary (teal)
                btn_bg = colors['button_primary']
                btn_hover = colors['button_primary_hover']
                btn_pressed = colors.get('button_primary_pressed', colors['button_primary_hover'])
                btn_fg = '#ffffff'
                btn_font = ('Segoe UI', 10, 'bold')
            else:
                # Other buttons: secondary
                btn_bg = colors['button_secondary']
                btn_hover = colors['button_secondary_hover']
                btn_pressed = colors.get('button_secondary_pressed', colors['button_secondary_hover'])
                btn_fg = colors['text_primary']
                btn_font = ('Segoe UI', 10)

            btn = RoundedButton(
                button_inner,
                text=btn_text,
                command=lambda t=btn_text: on_button_click(t),
                bg=btn_bg,
                hover_bg=btn_hover,
                pressed_bg=btn_pressed,
                fg=btn_fg,
                height=32, radius=6,
                font=btn_font,
                canvas_bg=canvas_bg
            )
            btn.pack(side=tk.LEFT, padx=(0, 10) if i < len(buttons) - 1 else 0)

        # Handle close button and escape
        dialog.protocol("WM_DELETE_WINDOW", lambda: on_button_click(buttons[-1]))
        dialog.bind('<Escape>', lambda e: on_button_click(buttons[-1]))
        dialog.bind('<Return>', lambda e: on_button_click(buttons[0]))

        # Auto-close countdown timer logic (label already created above buttons)
        if auto_close_seconds and auto_close_seconds > 0:
            # Countdown state
            remaining = [auto_close_seconds]
            timer_id = [None]

            def update_countdown():
                if not dialog.winfo_exists():
                    return
                remaining[0] -= 1
                if remaining[0] <= 0:
                    # Auto-close by clicking first button
                    on_button_click(buttons[0])
                else:
                    countdown_label.configure(text=f"Closing in {remaining[0]}...")
                    timer_id[0] = dialog.after(1000, update_countdown)

            # Start countdown
            timer_id[0] = dialog.after(1000, update_countdown)

            # Cancel countdown if user clicks any button or closes manually
            original_on_click = on_button_click

            def on_button_click_with_cancel(btn_text):
                if timer_id[0]:
                    dialog.after_cancel(timer_id[0])
                original_on_click(btn_text)

            # Re-bind buttons (they're already created, but we need to cancel timer)
            dialog.protocol("WM_DELETE_WINDOW", lambda: on_button_click_with_cancel(buttons[-1]))

        # Center and show
        dialog.update_idletasks()
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.deiconify()

        # Wait for dialog to close
        dialog.wait_window()

        # Return tuple based on checkbox configuration
        if checkbox_text and checkbox2_text:
            # Two checkboxes: (button, checkbox1_state, checkbox2_state)
            return (result[0], checkbox_var.get(), checkbox2_var.get())
        elif checkbox_text:
            # One checkbox: (button, checkbox_state)
            return (result[0], checkbox_var.get())
        return result[0]

    @staticmethod
    def showinfo(parent, title: str, message: str):
        """Show info message"""
        return ThemedMessageBox.show(parent, title, message, "info")

    @staticmethod
    def showwarning(parent, title: str, message: str):
        """Show warning message"""
        return ThemedMessageBox.show(parent, title, message, "warning")

    @staticmethod
    def showerror(parent, title: str, message: str):
        """Show error message"""
        return ThemedMessageBox.show(parent, title, message, "error")

    @staticmethod
    def showsuccess(parent, title: str, message: str):
        """Show success message"""
        return ThemedMessageBox.show(parent, title, message, "success")

    @staticmethod
    def askyesno(parent, title: str, message: str):
        """Show yes/no confirmation dialog"""
        result = ThemedMessageBox.show(parent, title, message, "warning", ["Yes", "No"])
        return result == "Yes"


class ThemedInputDialog:
    """
    Theme-aware input dialog that follows the app's dark/light mode.
    Replaces tkinter simpledialog.askstring for consistent styling.
    """

    @staticmethod
    def askstring(parent, title: str, prompt: str, initialvalue: str = "",
                  min_width: int = 300) -> str:
        """
        Show a themed input dialog.

        Args:
            parent: Parent window
            title: Dialog title
            prompt: Prompt text to display
            initialvalue: Initial value in the entry field
            min_width: Minimum width of the entry field (default: 300)

        Returns:
            The entered string, or None if cancelled
        """
        # Import RoundedButton here to avoid circular imports
        from core.ui.buttons import RoundedButton

        theme_manager = get_theme_manager()
        colors = theme_manager.colors
        is_dark = theme_manager.is_dark

        result = [None]

        # Create dialog
        dialog = tk.Toplevel(parent)
        dialog.withdraw()
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.configure(bg=colors['background'])

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
        main_frame.pack(fill=tk.BOTH)

        # Prompt label
        prompt_label = tk.Label(
            main_frame, text=prompt,
            bg=colors['background'], fg=colors['text_primary'],
            font=('Segoe UI', 10), justify=tk.LEFT, anchor='w'
        )
        prompt_label.pack(fill=tk.X, pady=(0, 10))

        # Entry field with theme styling
        entry_frame = tk.Frame(main_frame, bg=colors['border'], padx=1, pady=1)
        entry_frame.pack(fill=tk.X, pady=(0, 15))

        entry = tk.Entry(
            entry_frame,
            font=('Segoe UI', 10),
            bg=colors.get('surface', colors['background']),
            fg=colors['text_primary'],
            insertbackground=colors['text_primary'],
            relief='flat',
            width=max(40, min_width // 8)
        )
        entry.pack(fill=tk.X, padx=1, pady=1)
        entry.insert(0, initialvalue)
        entry.select_range(0, tk.END)
        entry.focus_set()

        # Button frame - centered
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        button_inner = tk.Frame(button_frame, bg=colors['background'])
        button_inner.pack(expand=True)

        def on_ok():
            result[0] = entry.get()
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

        # Key bindings
        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        dialog.bind('<Escape>', lambda e: on_cancel())
        dialog.bind('<Return>', lambda e: on_ok())

        # Center and show
        dialog.update_idletasks()
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.deiconify()

        # Wait for dialog to close
        dialog.wait_window()

        return result[0]
