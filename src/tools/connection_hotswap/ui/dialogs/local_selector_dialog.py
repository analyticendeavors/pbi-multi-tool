"""
Local Selector Dialog - Select local Power BI Desktop models
Built by Reid Havens of Analytic Endeavors

Modal dialog for discovering and selecting local Power BI Desktop models.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from typing import Optional, List

from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, ThemedScrollbar, ThemedMessageBox
from tools.field_parameters.models import AvailableModel
from tools.connection_hotswap.models import SwapTarget


class LocalSelectorDialog(tk.Toplevel):
    """
    Dialog for selecting local Power BI Desktop models.

    Features:
    - Scan for open models
    - Auto-highlight suggested match
    - Manual port entry
    - Test connection validation
    """

    def __init__(
        self,
        parent,
        matcher: 'LocalModelMatcher',
        suggested_name: Optional[str] = None,
        cached_models: Optional[List[AvailableModel]] = None
    ):
        """
        Initialize the local selector dialog.

        Args:
            parent: Parent window
            matcher: LocalModelMatcher instance for discovery
            suggested_name: Optional name to highlight as suggested match
            cached_models: Optional pre-cached models to display immediately
        """
        super().__init__(parent)

        self.matcher = matcher
        self.suggested_name = suggested_name
        self.result: Optional[SwapTarget] = None
        self.models: List[AvailableModel] = []
        self._has_cached_models = cached_models is not None and len(cached_models) > 0

        # Get theme manager
        self._theme_manager = get_theme_manager()
        self._colors = self._theme_manager.colors
        self._is_dark = self._theme_manager.is_dark

        self._setup_window()
        self._setup_ui()

        # Use cached models if available, otherwise scan
        if self._has_cached_models:
            self._populate_model_list(cached_models)
        else:
            self._scan_models()

        # Wait for dialog to close
        self.wait_window()

    def _setup_window(self):
        """Configure the dialog window"""
        colors = self._colors

        self.title("Select Local Model")
        # Window size adjusted for better content display
        self.geometry("450x470")
        self.resizable(True, True)
        self.minsize(400, 440)
        self.transient(self.master)
        self.grab_set()
        self.configure(bg=colors['background'])

        # Center on parent
        self.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() - 450) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - 470) // 2
        self.geometry(f"+{x}+{y}")

        # Set dialog icon (AE favicon)
        try:
            icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))),
                                     'assets', 'favicon.ico')
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass  # Ignore if icon can't be loaded

        # Set dark title bar on Windows for theme matching
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = ctypes.c_int(1 if self._is_dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(value), ctypes.sizeof(value))
        except Exception:
            pass

    def _setup_ui(self):
        """Setup the dialog UI with modern theme-aware styling"""
        colors = self._colors
        is_dark = self._is_dark

        # Main container with theme background
        main_frame = tk.Frame(self, bg=colors['background'], padx=15, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header row: "Available Models" label on left, Scan button on right
        # (Removed duplicate title - window title is sufficient)
        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f8f8fc')
        header_frame = tk.Frame(main_frame, bg=colors['background'])
        header_frame.pack(fill=tk.X, pady=(0, 8))

        # Section header - "Available Models" moved up, replaces duplicate title
        tk.Label(
            header_frame,
            text="Available Models",
            font=("Segoe UI", 10, "bold"),
            bg=colors['background'],
            fg=colors.get('title_color', '#0084b7' if is_dark else '#009999')
        ).pack(side=tk.LEFT)

        # Scan button in upper right (modern style)
        # Show RESCAN if we have cached models, otherwise SCAN
        scan_text = "RESCAN" if self._has_cached_models else "SCAN"
        canvas_bg = colors['background']
        self.scan_btn = RoundedButton(
            header_frame,
            text=scan_text,
            command=self._scan_models,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            width=70, height=28, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=canvas_bg
        )
        self.scan_btn.pack(side=tk.RIGHT)

        # Scan status label (theme-aware)
        self.scan_status = tk.StringVar(value="")
        self.scan_status_label = tk.Label(
            header_frame,
            textvariable=self.scan_status,
            fg=colors['text_muted'],
            bg=colors['background'],
            font=("Segoe UI", 9, "italic")
        )
        self.scan_status_label.pack(side=tk.RIGHT, padx=(0, 10))

        # Model list section with modern background (reduced top padding)
        list_frame = tk.Frame(main_frame, bg=section_bg, padx=10, pady=8)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Tree container with modern styling
        tree_container = tk.Frame(
            list_frame,
            bg=colors.get('surface', '#1e1e2e' if is_dark else '#ffffff'),
            highlightbackground=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
            highlightcolor=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
            highlightthickness=1
        )
        tree_container.pack(fill=tk.BOTH, expand=True)

        # Configure modern treeview style
        style = ttk.Style()
        tree_style = "LocalSelector.Treeview"

        tree_bg = '#161627' if is_dark else colors.get('surface', '#ffffff')
        tree_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        heading_bg = colors.get('section_bg', '#1a1a2a' if is_dark else '#f5f5fa')
        heading_fg = colors.get('text_primary', '#e0e0e0' if is_dark else '#333333')
        header_separator = '#0d0d1a' if is_dark else '#ffffff'

        style.configure(tree_style,
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=tree_bg,
                        font=('Segoe UI', 9),
                        relief='flat',
                        borderwidth=0,
                        bordercolor=tree_bg,
                        lightcolor=tree_bg,
                        darkcolor=tree_bg,
                        rowheight=28)  # Match connections table row height
        style.layout(tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])

        # Modern heading style
        style.configure(f"{tree_style}.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        relief='groove',
                        borderwidth=1,
                        bordercolor=header_separator,
                        lightcolor=header_separator,
                        darkcolor=header_separator,
                        font=('Segoe UI', 9, 'bold'),
                        padding=(8, 6))
        style.map(f"{tree_style}.Heading",
                  background=[('active', heading_bg), ('pressed', heading_bg), ('', heading_bg)],
                  relief=[('active', 'groove'), ('pressed', 'groove')])
        # Custom heading layout to ensure background colors are applied consistently
        style.layout(f"{tree_style}.Heading", [
            ('Treeheading.cell', {'sticky': 'nswe'}),
            ('Treeheading.border', {'sticky': 'nswe', 'children': [
                ('Treeheading.padding', {'sticky': 'nswe', 'children': [
                    ('Treeheading.image', {'side': 'left', 'sticky': ''}),
                    ('Treeheading.text', {'sticky': 'we'})
                ]})
            ]})
        ])
        style.map(tree_style,
                  background=[('selected', '#1a3a5c' if is_dark else '#e6f3ff')],
                  foreground=[('selected', '#ffffff' if is_dark else '#1a1a2e')])

        # Treeview for models
        columns = ("name", "server", "match")
        self.model_tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            selectmode="browse",
            height=6,
            style=tree_style
        )

        self.model_tree.heading("name", text="Model Name")
        self.model_tree.heading("server", text="Server")
        self.model_tree.heading("match", text="Match")

        self.model_tree.column("name", width=200, minwidth=120)
        self.model_tree.column("server", width=130, minwidth=80)
        self.model_tree.column("match", width=60, minwidth=50, anchor="center")

        # Modern themed scrollbar
        scrollbar = ThemedScrollbar(
            tree_container,
            command=self.model_tree.yview,
            theme_manager=self._theme_manager
        )
        self.model_tree.configure(yscrollcommand=scrollbar.set)

        self.model_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Double-click to select
        self.model_tree.bind("<Double-1>", lambda e: self._on_select())

        # Manual entry section (compact)
        manual_frame = tk.Frame(main_frame, bg=section_bg, padx=10, pady=8)
        manual_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            manual_frame,
            text="Manual Entry",
            font=("Segoe UI", 10, "bold"),
            bg=section_bg,
            fg=colors['title_color']
        ).pack(anchor='w', pady=(0, 6))

        entry_row = tk.Frame(manual_frame, bg=section_bg)
        entry_row.pack(fill=tk.X)

        tk.Label(
            entry_row, text="Server:",
            bg=section_bg, fg=colors['text_primary'],
            font=("Segoe UI", 9)
        ).pack(side=tk.LEFT)

        self.manual_server = tk.StringVar(value="localhost:")
        self.server_entry = tk.Entry(
            entry_row,
            textvariable=self.manual_server,
            width=16,
            font=("Segoe UI", 9),
            bg=colors.get('surface', '#1e1e2e' if is_dark else '#ffffff'),
            fg=colors['text_primary'],
            insertbackground=colors['text_primary'],
            relief='flat',
            highlightthickness=1,
            highlightbackground=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
            highlightcolor=colors.get('primary', '#4a6cf5')
        )
        self.server_entry.pack(side=tk.LEFT, padx=(5, 8))

        tk.Label(
            entry_row, text="Database:",
            bg=section_bg, fg=colors['text_primary'],
            font=("Segoe UI", 9)
        ).pack(side=tk.LEFT)

        self.manual_database = tk.StringVar()
        self.database_entry = tk.Entry(
            entry_row,
            textvariable=self.manual_database,
            width=16,
            font=("Segoe UI", 9),
            bg=colors.get('surface', '#1e1e2e' if is_dark else '#ffffff'),
            fg=colors['text_primary'],
            insertbackground=colors['text_primary'],
            relief='flat',
            highlightthickness=1,
            highlightbackground=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
            highlightcolor=colors.get('primary', '#4a6cf5')
        )
        self.database_entry.pack(side=tk.LEFT, padx=(5, 8))

        self.test_btn = RoundedButton(
            entry_row,
            text="TEST",
            command=self._test_manual_connection,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            width=50, height=24, radius=4,
            font=('Segoe UI', 8),
            canvas_bg=section_bg
        )
        self.test_btn.pack(side=tk.LEFT)

        # Action buttons
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X, pady=(5, 0))

        self.cancel_btn = RoundedButton(
            button_frame,
            text="CANCEL",
            command=self._on_cancel,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=32, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background']
        )
        self.cancel_btn.pack(side=tk.RIGHT)

        self.select_btn = RoundedButton(
            button_frame,
            text="SELECT",
            command=self._on_select,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            height=32, radius=5,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=colors['background']
        )
        self.select_btn.pack(side=tk.RIGHT, padx=(0, 10))

    def _scan_models(self):
        """Scan for local Power BI Desktop models"""
        self.scan_btn.set_enabled(False)
        self.scan_status.set("Scanning...")
        self.model_tree.delete(*self.model_tree.get_children())

        def scan_thread():
            try:
                models = self.matcher.discover_local_models(force_refresh=True)
                self.after(0, lambda: self._populate_model_list(models))
            except Exception as e:
                self.after(0, lambda: self.scan_status.set(f"Error: {e}"))
            finally:
                self.after(0, lambda: self.scan_btn.set_enabled(True))

        threading.Thread(target=scan_thread, daemon=True).start()

    def _populate_model_list(self, models: List[AvailableModel]):
        """Populate the model list"""
        self.models = models
        self.model_tree.delete(*self.model_tree.get_children())

        if not models:
            self.scan_status.set("No models found")
            return

        self.scan_status.set(f"Found {len(models)} model(s)")

        # Find suggested match
        suggested_match = None
        if self.suggested_name:
            suggested_match = self.matcher.find_matching_model(
                self.suggested_name,
                models
            )

        for i, model in enumerate(models):
            match_indicator = ""
            tags = ()

            if suggested_match and model.display_name == suggested_match.display_name:
                match_indicator = "Match"
                tags = ("suggested",)

            self.model_tree.insert(
                "",
                "end",
                iid=str(i),
                values=(model.display_name, model.server, match_indicator),
                tags=tags
            )

        # Note: Don't apply background styling to "suggested" tag
        # as it conflicts with treeview selection highlighting (double-highlight bug).
        # The "Match" text indicator in the match column is sufficient.

        # Select suggested match
        if suggested_match:
            for i, model in enumerate(models):
                if model.display_name == suggested_match.display_name:
                    self.model_tree.selection_set(str(i))
                    self.model_tree.see(str(i))
                    break

    def _test_manual_connection(self):
        """Test manual connection entry"""
        server = self.manual_server.get().strip()
        database = self.manual_database.get().strip()

        if not server or not database:
            ThemedMessageBox.showwarning(self, "Missing Info", "Please enter both server and database.")
            return

        import socket

        try:
            if ':' in server:
                host, port_str = server.rsplit(':', 1)
                port = int(port_str)
            else:
                ThemedMessageBox.showwarning(self, "Invalid Format", "Server should be localhost:PORT")
                return

            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(2)
            result = sock.connect_ex((host, port))
            sock.close()

            if result == 0:
                ThemedMessageBox.showinfo(self, "Success", f"Connection to {server} successful!")
            else:
                ThemedMessageBox.showerror(self, "Failed", f"Cannot connect to {server}")

        except ValueError:
            ThemedMessageBox.showerror(self, "Invalid Port", "Port must be a number")
        except Exception as e:
            ThemedMessageBox.showerror(self, "Error", f"Connection test failed: {e}")

    def _on_select(self):
        """Confirm selection and close"""
        # Check for tree selection first
        selection = self.model_tree.selection()
        if selection:
            idx = int(selection[0])
            if idx < len(self.models):
                model = self.models[idx]
                self.result = SwapTarget(
                    target_type="local",
                    server=model.server,
                    database=model.database_name,
                    display_name=model.display_name,
                )
                self.destroy()
                return

        # Check for manual entry
        server = self.manual_server.get().strip()
        database = self.manual_database.get().strip()

        if server and database:
            self.result = SwapTarget(
                target_type="local",
                server=server,
                database=database,
                display_name=f"Manual ({server})",
            )
            self.destroy()
            return

        ThemedMessageBox.showinfo(self, "Select Model", "Please select a model or enter connection details.")

    def _on_cancel(self):
        """Cancel and close"""
        self.result = None
        self.destroy()
