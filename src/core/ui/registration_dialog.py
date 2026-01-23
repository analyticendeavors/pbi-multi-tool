"""
Registration Dialog - First-launch registration form
Built by Reid Havens of Analytic Endeavors
"""

import re
import io
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable
from pathlib import Path

from core.theme_manager import get_theme_manager
from core.ui.buttons import RoundedButton  # Use main RoundedButton with consistent padding

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


class RegistrationDialog:
    """
    Modal registration dialog shown on first launch.
    Collects: Name, Company, Job Title, Work Email
    """

    def __init__(self, parent: tk.Tk, registration_manager, app_version: str = ""):
        self.parent = parent
        self.reg_manager = registration_manager
        self.app_version = app_version
        self.success = False
        self.dialog: Optional[tk.Toplevel] = None
        self._logo_image = None  # Keep reference to prevent garbage collection
        self._user_icon = None  # Keep reference to prevent garbage collection
        self._verify_icon = None  # Keep reference to prevent garbage collection

        # Get theme colors
        self._theme_manager = get_theme_manager()

    def _load_icon(self, icon_name: str, size: int = 48) -> Optional['ImageTk.PhotoImage']:
        """Load an SVG icon by name"""
        if not PIL_AVAILABLE:
            return None

        # Get icon path
        icons_dir = Path(__file__).parent.parent.parent / "assets" / "Tool Icons"
        svg_path = icons_dir / f"{icon_name}.svg"

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

    def show(self) -> bool:
        """
        Show the registration dialog.
        Returns True if registration was successful, False if cancelled.
        """
        colors = self._theme_manager.colors

        self.dialog = tk.Toplevel(self.parent)
        self.dialog.withdraw()  # Hide initially to prevent flicker
        self.dialog.title("Register - AE Multi-Tool")
        self.dialog.geometry("450x480")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=colors['background'])
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # Set favicon icon
        try:
            import os
            # Look for icon relative to script location
            base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            icon_path = os.path.join(base_path, "assets", "favicon.ico")
            if os.path.exists(icon_path):
                self.dialog.iconbitmap(icon_path)
        except Exception:
            pass

        # Main frame
        main_frame = tk.Frame(self.dialog, bg=colors['background'], padx=30, pady=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header with logo on left and text aligned next to it
        header_frame = tk.Frame(main_frame, bg=colors['background'])
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Load and display logo (bigger size)
        self._logo_image = self._load_icon("analyticsnavytealicon", size=56)
        if self._logo_image:
            logo_label = tk.Label(
                header_frame,
                image=self._logo_image,
                bg=colors['background']
            )
            logo_label.pack(side=tk.LEFT, padx=(0, 15))

        # Text container (left-aligned next to logo)
        text_frame = tk.Frame(header_frame, bg=colors['background'])
        text_frame.pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(
            text_frame,
            text="Welcome to AE Multi-Tool",
            font=('Segoe UI', 16, 'bold'),
            fg=colors['text_primary'],
            bg=colors['background'],
            anchor='w'
        ).pack(anchor='w')

        tk.Label(
            text_frame,
            text="Please register to continue",
            font=('Segoe UI', 10),
            fg=colors['text_secondary'],
            bg=colors['background'],
            anchor='w'
        ).pack(anchor='w', pady=(2, 0))

        # Form fields frame
        form_frame = tk.Frame(main_frame, bg=colors['background'])
        form_frame.pack(fill=tk.X, pady=(0, 5))

        # Name field
        self.name_var = tk.StringVar()
        self._create_form_field(form_frame, "Name *", self.name_var, 0)

        # Company field
        self.company_var = tk.StringVar()
        self._create_form_field(form_frame, "Company *", self.company_var, 1)

        # Job Title field
        self.job_title_var = tk.StringVar()
        self._create_form_field(form_frame, "Job Title *", self.job_title_var, 2)

        # Email field
        self.email_var = tk.StringVar()
        self._create_form_field(form_frame, "Work Email *", self.email_var, 3)

        # Error label (hidden initially)
        self.error_label = tk.Label(
            main_frame,
            text="",
            font=('Segoe UI', 9),
            fg=colors['error'],
            bg=colors['background']
        )
        self.error_label.pack(pady=(0, 5))

        # Buttons frame - contains "Already Registered" link on left, buttons on right
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X, pady=(10, 0))

        # Load user icon for Register button
        self._user_icon = self._load_icon("user", size=16)

        # Register button (primary) - rounded with icon, height matches main tool (32px)
        self.register_btn = RoundedButton(
            button_frame,
            text="Register",
            command=self._on_register,
            bg=colors['button_primary'],
            fg='#ffffff',
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors.get('button_primary_pressed', '#00404f'),
            height=32,
            radius=6,
            icon=self._user_icon,
            canvas_bg=colors['background']
        )
        self.register_btn.pack(side=tk.RIGHT, padx=(6, 0))

        # Cancel button (secondary) - rounded, height matches main tool (32px)
        self.cancel_btn = RoundedButton(
            button_frame,
            text="Cancel",
            command=self._on_cancel,
            bg=colors['card_surface'],
            fg=colors['text_primary'],
            hover_bg=colors['card_surface_hover'],
            pressed_bg=colors.get('card_surface_pressed', colors['card_surface_hover']),
            height=32,
            radius=6,
            canvas_bg=colors['background']
        )
        self.cancel_btn.pack(side=tk.RIGHT)

        # "Already Registered" link on left side, vertically centered with buttons
        # Use title_color for a lighter blue (not teal)
        link_color = colors.get('title_color', '#0084b7')  # Lighter blue
        link_hover_color = '#60c8ff'  # Lighter on hover for better visibility
        already_label = tk.Label(
            button_frame,
            text="Already Registered?",
            font=('Segoe UI', 9, 'underline'),
            fg=link_color,
            bg=colors['background'],
            cursor='hand2'
        )
        already_label.pack(side=tk.LEFT, pady=(0, 0))
        already_label.bind('<Button-1>', lambda e: self._show_email_verification())
        already_label.bind('<Enter>', lambda e: already_label.config(fg=link_hover_color))
        already_label.bind('<Leave>', lambda e: already_label.config(fg=link_color))

        # Bind Enter key to register
        self.dialog.bind('<Return>', lambda e: self._on_register())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())

        # Center dialog on screen
        self.dialog.update_idletasks()
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        dialog_width = self.dialog.winfo_reqwidth()
        dialog_height = self.dialog.winfo_reqheight()
        x = (screen_width - dialog_width) // 2
        y = (screen_height - dialog_height) // 2
        self.dialog.geometry(f"+{x}+{y}")

        # Set dark/light title bar
        self.dialog.update()
        self._set_title_bar_color()
        self.dialog.deiconify()

        # Focus first field
        self.name_entry.focus_set()

        # Wait for dialog to close
        self.dialog.wait_window()

        return self.success

    def _create_form_field(self, parent: tk.Frame, label_text: str, variable: tk.StringVar, row: int):
        """Create a labeled form field"""
        colors = self._theme_manager.colors

        # Label
        label = tk.Label(
            parent,
            text=label_text,
            font=('Segoe UI', 9),
            fg=colors['text_secondary'],
            bg=colors['background'],
            anchor='w'
        )
        label.grid(row=row * 2, column=0, sticky='w', pady=(8, 2))

        # Entry container frame for better padding control
        entry_container = tk.Frame(
            parent,
            bg=colors.get('input_bg', colors['surface']),
            highlightthickness=1,
            highlightcolor=colors['button_primary'],
            highlightbackground=colors['border']
        )
        entry_container.grid(row=row * 2 + 1, column=0, sticky='ew', pady=(0, 0))

        # Entry with themed styling
        entry = tk.Entry(
            entry_container,
            textvariable=variable,
            font=('Segoe UI', 10),
            fg=colors['text_primary'],
            bg=colors.get('input_bg', colors['surface']),
            insertbackground=colors['text_primary'],
            relief='flat',
            bd=0,
            highlightthickness=0
        )
        entry.pack(fill=tk.X, padx=12, pady=8)

        # Configure column to expand
        parent.columnconfigure(0, weight=1)

        # Store reference to name entry for focus
        if row == 0:
            self.name_entry = entry

    def _validate_form(self) -> tuple[bool, str]:
        """Validate form fields"""
        name = self.name_var.get().strip()
        company = self.company_var.get().strip()
        job_title = self.job_title_var.get().strip()
        email = self.email_var.get().strip()

        if not name:
            return False, "Name is required"
        if not company:
            return False, "Company is required"
        if not job_title:
            return False, "Job Title is required"
        if not email:
            return False, "Work Email is required"

        # Basic email validation
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        if not re.match(email_pattern, email):
            return False, "Please enter a valid email address"

        return True, ""

    def _on_register(self):
        """Handle register button click"""
        colors = self._theme_manager.colors

        # Validate form
        valid, error_msg = self._validate_form()
        if not valid:
            self.error_label.config(text=error_msg)
            return

        # Clear error
        self.error_label.config(text="")

        # Disable button and show loading state
        self.register_btn.update_text("Registering")
        self.register_btn.set_enabled(False)
        self.dialog.update()

        # Attempt registration
        success, message = self.reg_manager.register_user(
            name=self.name_var.get().strip(),
            email=self.email_var.get().strip(),
            company=self.company_var.get().strip(),
            job_title=self.job_title_var.get().strip(),
            app_version=self.app_version
        )

        if success:
            self.success = True
            self.dialog.destroy()
        else:
            self.error_label.config(text=message)
            self.register_btn.update_text("Register")
            self.register_btn.set_enabled(True)

    def _on_cancel(self):
        """Handle cancel button click"""
        self.success = False
        self.dialog.destroy()

    def _show_email_verification(self):
        """Show email verification dialog for returning users"""
        colors = self._theme_manager.colors

        verify_dialog = tk.Toplevel(self.dialog)
        verify_dialog.withdraw()
        verify_dialog.title("Verify Registration")
        # Size set later with position for proper centering
        verify_dialog.resizable(False, False)
        verify_dialog.transient(self.dialog)
        verify_dialog.grab_set()
        verify_dialog.configure(bg=colors['background'])

        # Set icon
        try:
            import os
            base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            icon_path = os.path.join(base_path, "assets", "favicon.ico")
            if os.path.exists(icon_path):
                verify_dialog.iconbitmap(icon_path)
        except Exception:
            pass

        main_frame = tk.Frame(verify_dialog, bg=colors['background'], padx=30, pady=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main_frame,
            text="Enter your registered email",
            font=('Segoe UI', 11, 'bold'),
            fg=colors['text_primary'],
            bg=colors['background']
        ).pack(pady=(0, 15))

        email_var = tk.StringVar()

        # Entry container for better padding
        entry_container = tk.Frame(
            main_frame,
            bg=colors.get('input_bg', colors['surface']),
            highlightthickness=1,
            highlightcolor=colors['button_primary'],
            highlightbackground=colors['border']
        )
        entry_container.pack(fill=tk.X)

        email_entry = tk.Entry(
            entry_container,
            textvariable=email_var,
            font=('Segoe UI', 10),
            fg=colors['text_primary'],
            bg=colors.get('input_bg', colors['surface']),
            insertbackground=colors['text_primary'],
            relief='flat',
            bd=0,
            highlightthickness=0
        )
        email_entry.pack(fill=tk.X, padx=12, pady=10)

        error_label = tk.Label(
            main_frame,
            text="",
            font=('Segoe UI', 9),
            fg=colors['error'],
            bg=colors['background']
        )
        error_label.pack(pady=(10, 0))

        def on_verify():
            email = email_var.get().strip()
            if not email:
                error_label.config(text="Please enter your email")
                return

            verify_btn.update_text("Verifying")
            verify_btn.set_enabled(False)
            verify_dialog.update()

            success, message = self.reg_manager.verify_existing_email(email)
            if success:
                self.success = True
                verify_dialog.destroy()
                self.dialog.destroy()
            else:
                error_label.config(text=message)
                verify_btn.update_text("Verify")
                verify_btn.set_enabled(True)

        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X, pady=(15, 0))

        # Load user search icon for Verify button
        self._verify_icon = self._load_icon("user search", size=16)

        verify_btn = RoundedButton(
            button_frame,
            text="Verify",
            command=on_verify,
            bg=colors['button_primary'],
            fg='#ffffff',
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors.get('button_primary_pressed', '#00404f'),
            height=32,
            radius=6,
            icon=self._verify_icon,
            canvas_bg=colors['background']
        )
        verify_btn.pack(side=tk.RIGHT)

        back_btn = RoundedButton(
            button_frame,
            text="Back",
            command=verify_dialog.destroy,
            bg=colors['card_surface'],
            fg=colors['text_primary'],
            hover_bg=colors['card_surface_hover'],
            pressed_bg=colors.get('card_surface_pressed', colors['card_surface_hover']),
            height=32,
            radius=6,
            canvas_bg=colors['background']
        )
        back_btn.pack(side=tk.RIGHT, padx=(0, 10))

        verify_dialog.bind('<Return>', lambda e: on_verify())
        verify_dialog.bind('<Escape>', lambda e: verify_dialog.destroy())

        # Center on parent dialog (use explicit dimensions)
        verify_width = 400
        verify_height = 220
        self.dialog.update_idletasks()
        # Calculate parent's center point, then position verify dialog centered on it
        # Note: subtract 10px to compensate for window frame offset on Windows
        parent_center_x = self.dialog.winfo_rootx() + self.dialog.winfo_width() // 2
        parent_center_y = self.dialog.winfo_rooty() + self.dialog.winfo_height() // 2
        x = parent_center_x - verify_width // 2 - 10
        y = parent_center_y - verify_height // 2
        verify_dialog.geometry(f"{verify_width}x{verify_height}+{x}+{y}")

        verify_dialog.update()
        self._set_title_bar_color(verify_dialog)
        verify_dialog.deiconify()
        email_entry.focus_set()

    def _set_title_bar_color(self, window: Optional[tk.Toplevel] = None):
        """Set dark/light title bar color on Windows"""
        if window is None:
            window = self.dialog

        try:
            import ctypes
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            value = ctypes.c_int(1 if self._theme_manager.is_dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(value),
                ctypes.sizeof(value)
            )
        except Exception:
            pass  # Not on Windows or API not available
