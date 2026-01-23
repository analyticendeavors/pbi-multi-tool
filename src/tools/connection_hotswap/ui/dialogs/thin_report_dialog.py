"""
Thin Report Swap Dialog - Connect thin reports to local models
Built by Reid Havens of Analytic Endeavors

Modal dialog for helping users connect thin reports (live-connected reports)
to local Power BI Desktop models.

Supports two modes:
1. Manual reconnection (Get Data > Analysis Services)
2. Automatic file swap (modify PBIX/PBIP connection directly)
"""

import io
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox
from typing import Callable, List, Optional

from core.theme_manager import get_theme_manager
from core.ui_base import RoundedButton, ThemedScrollbar, ThemedMessageBox
from tools.field_parameters.models import AvailableModel

# Optional imports for icon loading
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False


class ThinReportSwapDialog(tk.Toplevel):
    """
    Dialog for connecting thin reports to local models.

    Features:
    - List of available local models
    - Matching models highlighted (by cloud dataset name)
    - Copy connection string to clipboard
    - SWAP button to automatically modify the file connection
    - Different handling for PBIX (requires close) vs PBIP (can edit while open)
    """

    def __init__(
        self,
        parent,
        report_name: str,
        cloud_server: str = '',
        cloud_database: str = '',
        local_models: Optional[List[AvailableModel]] = None,
        matching_models: Optional[List[AvailableModel]] = None,
        error: str = '',
        on_select_local: Optional[Callable[[AvailableModel], None]] = None,
        thin_report_file_path: Optional[str] = None,
        thin_report_process_id: Optional[int] = None
    ):
        """
        Initialize the thin report swap dialog.

        Args:
            parent: Parent window
            report_name: Name of the thin report
            cloud_server: Cloud server URL (if known)
            cloud_database: Cloud dataset name (if known)
            local_models: List of available local models
            matching_models: List of models matching the cloud dataset name
            error: Optional error message from connection attempt
            on_select_local: Callback when user wants to open a local model
            thin_report_file_path: Full path to the thin report .pbix/.pbip file
            thin_report_process_id: Power BI Desktop process ID for the thin report
        """
        super().__init__(parent)

        self.report_name = report_name
        self.cloud_server = cloud_server
        self.cloud_database = cloud_database
        self.local_models = local_models or []
        self.matching_models = matching_models or []
        self.error = error
        self.on_select_local = on_select_local
        self.thin_report_file_path = thin_report_file_path
        self.thin_report_process_id = thin_report_process_id
        self.result: Optional[AvailableModel] = None
        self._sorted_models: List[AvailableModel] = []  # Initialized in _populate_models

        # Determine file type for swap behavior
        self._file_type = self._detect_file_type()
        self._swap_enabled = self.thin_report_file_path is not None

        # Check for cached cloud connection (for restore functionality)
        self._has_cached_cloud = False
        self._cached_cloud_info = None
        if self.thin_report_file_path:
            from tools.connection_hotswap.logic.pbix_modifier import get_modifier
            modifier = get_modifier()
            self._cached_cloud_info = modifier.get_cached_cloud_connection(self.thin_report_file_path)
            self._has_cached_cloud = self._cached_cloud_info is not None

        # Get theme manager
        self._theme_manager = get_theme_manager()
        self._colors = self._theme_manager.colors
        self._is_dark = self._theme_manager.is_dark

        # Load icons
        self._hotswap_icon = self._load_icon('hotswap', size=16)

        # Status tracking for integrated progress
        self._status_var = None
        self._status_label = None

        self._setup_window()
        self._setup_ui()

        # Wait for dialog to close
        self.wait_window()

    def _setup_window(self):
        """Configure the dialog window"""
        colors = self._colors

        self.title("Connect Thin Report to Local Model")
        self.geometry("520x480")
        self.resizable(True, True)
        self.minsize(450, 400)

        # Find the top-level window for proper centering and transient behavior
        top_window = self.master.winfo_toplevel()
        self.transient(top_window)
        self.grab_set()
        self.configure(bg=colors['background'])

        # Center on the top-level window
        self.update_idletasks()
        dialog_width = 520
        dialog_height = 480
        top_x = top_window.winfo_x()
        top_y = top_window.winfo_y()
        top_width = top_window.winfo_width()
        top_height = top_window.winfo_height()

        x = top_x + (top_width - dialog_width) // 2
        y = top_y + (top_height - dialog_height) // 2
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        # Ensure dialog appears on top
        self.lift()
        self.focus_force()

        # Set dialog icon (AE favicon)
        try:
            icon_path = os.path.join(
                os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))),
                'assets', 'favicon.ico'
            )
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass  # Ignore if icon can't be loaded

    def _setup_ui(self):
        """Setup the dialog UI"""
        colors = self._colors
        is_dark = self._is_dark
        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f8f8fc')

        # Main container
        main_frame = tk.Frame(self, bg=colors['background'], padx=15, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ===== Current Connection Info =====
        info_frame = tk.Frame(main_frame, bg=section_bg, padx=12, pady=10)
        info_frame.pack(fill=tk.X, pady=(0, 12))

        # Thin report name
        tk.Label(
            info_frame,
            text="Thin Report:",
            font=("Segoe UI", 9),
            bg=section_bg,
            fg=colors['text_muted']
        ).grid(row=0, column=0, sticky='w', pady=2)

        tk.Label(
            info_frame,
            text=self.report_name,
            font=("Segoe UI", 10, "bold"),
            bg=section_bg,
            fg=colors['text_primary']
        ).grid(row=0, column=1, sticky='w', padx=(8, 0), pady=2)

        # Cloud connection
        if self.cloud_database:
            tk.Label(
                info_frame,
                text="Connected to:",
                font=("Segoe UI", 9),
                bg=section_bg,
                fg=colors['text_muted']
            ).grid(row=1, column=0, sticky='w', pady=2)

            cloud_text = f"{self.cloud_database} (Cloud)"
            tk.Label(
                info_frame,
                text=cloud_text,
                font=("Segoe UI", 10),
                bg=section_bg,
                fg=colors.get('primary', '#4a6cf5')
            ).grid(row=1, column=1, sticky='w', padx=(8, 0), pady=2)

        # Error message if present
        if self.error:
            error_label = tk.Label(
                info_frame,
                text=f"Could not connect to cloud: {self.error[:50]}...",
                font=("Segoe UI", 9),
                bg=section_bg,
                fg=colors.get('error', '#ff4444')
            )
            error_label.grid(row=2, column=0, columnspan=2, sticky='w', pady=(4, 0))

        # ===== Local Models Section =====
        models_header = tk.Frame(main_frame, bg=colors['background'])
        models_header.pack(fill=tk.X, pady=(0, 6))

        header_text = "Available Local Models"
        if self.matching_models:
            header_text = f"Available Local Models ({len(self.matching_models)} match)"

        tk.Label(
            models_header,
            text=header_text,
            font=("Segoe UI", 10, "bold"),
            bg=colors['background'],
            fg=colors.get('title_color', '#0084b7' if is_dark else '#009999')
        ).pack(side=tk.LEFT)

        # Models list frame
        list_frame = tk.Frame(main_frame, bg=section_bg, padx=10, pady=8)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        if not self.local_models:
            # No models found message
            no_models_label = tk.Label(
                list_frame,
                text="No local Power BI Desktop models found.\n\nOpen a .pbix file that contains the data model\nyou want to connect to.",
                font=("Segoe UI", 10),
                bg=section_bg,
                fg=colors['text_muted'],
                justify='center'
            )
            no_models_label.pack(expand=True, pady=20)
        else:
            # Tree container
            tree_container = tk.Frame(
                list_frame,
                bg=colors.get('surface', '#1e1e2e' if is_dark else '#ffffff'),
                highlightbackground=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
                highlightcolor=colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0'),
                highlightthickness=1
            )
            tree_container.pack(fill=tk.BOTH, expand=True)

            # Configure treeview style
            style = ttk.Style()
            tree_style = "ThinReport.Treeview"

            tree_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
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
                            rowheight=28)
            style.layout(tree_style, [('Treeview.treearea', {'sticky': 'nswe'})])

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

            self.model_tree.column("name", width=220, minwidth=120)
            self.model_tree.column("server", width=140, minwidth=80)
            self.model_tree.column("match", width=60, minwidth=50, anchor="center")

            # Scrollbar
            scrollbar = ThemedScrollbar(
                tree_container,
                command=self.model_tree.yview,
                theme_manager=self._theme_manager
            )
            self.model_tree.configure(yscrollcommand=scrollbar.set)

            self.model_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # Populate models - matching ones first
            self._populate_models()

            # Bind selection for copy button state
            self.model_tree.bind("<<TreeviewSelect>>", self._on_selection_change)

        # ===== Copy Connection Section =====
        copy_frame = tk.Frame(main_frame, bg=section_bg, padx=12, pady=10)
        copy_frame.pack(fill=tk.X, pady=(0, 12))

        copy_row = tk.Frame(copy_frame, bg=section_bg)
        copy_row.pack(fill=tk.X)

        tk.Label(
            copy_row,
            text="Connection String:",
            font=("Segoe UI", 9),
            bg=section_bg,
            fg=colors['text_muted']
        ).pack(side=tk.LEFT)

        self.connection_var = tk.StringVar(value="(select a model above)")
        self.connection_label = tk.Label(
            copy_row,
            textvariable=self.connection_var,
            font=("Consolas", 9),
            bg=section_bg,
            fg=colors['text_primary'],
            anchor='w'
        )
        self.connection_label.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)

        self.copy_btn = RoundedButton(
            copy_row,
            text="COPY",
            command=self._copy_connection_string,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            width=60, height=26, radius=4,
            font=('Segoe UI', 9),
            canvas_bg=section_bg
        )
        self.copy_btn.pack(side=tk.RIGHT)
        self.copy_btn.set_enabled(False)

        # ===== Status Area (for swap progress) =====
        self._status_var = tk.StringVar(value="")
        self._status_label = tk.Label(
            main_frame,
            textvariable=self._status_var,
            font=("Segoe UI", 9),
            bg=colors['background'],
            fg=colors.get('primary', '#4a6cf5'),
            anchor='center'
        )
        self._status_label.pack(fill=tk.X, pady=(0, 8))
        self._status_label.pack_forget()  # Hidden by default

        # ===== Action Buttons (centered) =====
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        # Center container for buttons
        button_center = tk.Frame(button_frame, bg=colors['background'])
        button_center.pack(anchor='center')

        # RESTORE CLOUD button (if we have cached cloud connection)
        if self._has_cached_cloud and self._swap_enabled:
            cloud_color = colors.get('primary', '#4a6cf5')
            cloud_hover = colors.get('primary_hover', '#3d5dd5')
            self.restore_cloud_btn = RoundedButton(
                button_center,
                text="RESTORE CLOUD",
                command=self._on_restore_cloud,
                bg=cloud_color,
                hover_bg=cloud_hover,
                pressed_bg=cloud_hover,
                fg='#ffffff',
                height=32, radius=5,
                font=('Segoe UI', 9, 'bold'),
                canvas_bg=colors['background']
            )
            self.restore_cloud_btn.pack(side=tk.LEFT, padx=(0, 10))

        # SWAP button (if file path is available for auto-swap)
        if self._swap_enabled and self.local_models:
            self.swap_btn = RoundedButton(
                button_center,
                text="SWAP",
                command=self._on_swap_connection,
                bg=colors['button_primary'],
                hover_bg=colors['button_primary_hover'],
                pressed_bg=colors['button_primary_pressed'],
                fg='#ffffff',
                height=32, radius=5,
                font=('Segoe UI', 9, 'bold'),
                canvas_bg=colors['background'],
                icon=self._hotswap_icon
            )
            self.swap_btn.pack(side=tk.LEFT, padx=(0, 10))
            self.swap_btn.set_enabled(False)

        self.close_btn = RoundedButton(
            button_center,
            text="CLOSE",
            command=self._on_close,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=32, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background']
        )
        self.close_btn.pack(side=tk.LEFT)

    def _populate_models(self):
        """Populate the model tree with matching models first"""
        # Get matching model names for quick lookup
        matching_names = {m.display_name for m in self.matching_models}

        # Sort: matching models first, then others
        sorted_models = sorted(
            self.local_models,
            key=lambda m: (0 if m.display_name in matching_names else 1, m.display_name)
        )

        for i, model in enumerate(sorted_models):
            match_indicator = "Match" if model.display_name in matching_names else ""

            self.model_tree.insert(
                "",
                "end",
                iid=str(i),
                values=(model.display_name, model.server, match_indicator)
            )

        # Store sorted models for later reference
        self._sorted_models = sorted_models

        # Auto-select first matching model if available
        if self.matching_models:
            self.model_tree.selection_set("0")
            self._on_selection_change(None)

    def _on_selection_change(self, event):
        """Handle model selection change"""
        selection = self.model_tree.selection()
        if selection:
            idx = int(selection[0])
            if idx < len(self._sorted_models):
                model = self._sorted_models[idx]
                # Update connection string display
                self.connection_var.set(f"{model.server}")
                self.copy_btn.set_enabled(True)
                if hasattr(self, 'swap_btn'):
                    self.swap_btn.set_enabled(True)
        else:
            self.connection_var.set("(select a model above)")
            self.copy_btn.set_enabled(False)
            if hasattr(self, 'swap_btn'):
                self.swap_btn.set_enabled(False)

    def _copy_connection_string(self):
        """Copy the connection string to clipboard"""
        selection = self.model_tree.selection()
        if not selection:
            return

        idx = int(selection[0])
        if idx >= len(self._sorted_models):
            return

        model = self._sorted_models[idx]

        # Build connection string
        connection_info = f"Server: {model.server}\nDatabase: {model.database_name}"

        # Copy to clipboard
        self.clipboard_clear()
        self.clipboard_append(connection_info)

        # Show feedback
        original_text = self.copy_btn.text
        self.copy_btn.update_text("COPIED!")
        self.after(1500, lambda: self.copy_btn.update_text(original_text))

    def _on_close(self):
        """Close the dialog"""
        self.destroy()

    def _load_icon(self, icon_name: str, size: int = 16) -> Optional['ImageTk.PhotoImage']:
        """Load an SVG icon for use in buttons."""
        if not PIL_AVAILABLE:
            return None

        # Get icon path
        icons_dir = Path(__file__).parent.parent.parent.parent.parent / "assets" / "Tool Icons"
        svg_path = icons_dir / f"{icon_name}.svg"

        try:
            if CAIROSVG_AVAILABLE and svg_path.exists():
                # Render at 4x size for quality, then downscale
                png_data = cairosvg.svg2png(
                    url=str(svg_path),
                    output_width=size * 4,
                    output_height=size * 4
                )
                img = Image.open(io.BytesIO(png_data))
                img = img.resize((size, size), Image.Resampling.LANCZOS)
                return ImageTk.PhotoImage(img)
        except Exception:
            pass

        return None

    def _detect_file_type(self) -> Optional[str]:
        """Detect whether the thin report is a PBIX or PBIP file."""
        if not self.thin_report_file_path:
            return None

        lower_path = self.thin_report_file_path.lower()
        if lower_path.endswith('.pbix'):
            return 'pbix'
        elif lower_path.endswith('.pbip'):
            return 'pbip'
        return None

    def _on_restore_cloud(self):
        """Handle restore cloud connection button click."""
        if not self._cached_cloud_info:
            ThemedMessageBox.showerror(self, "Error", "No cached cloud connection available.")
            return

        if self._file_type == 'pbip':
            # PBIP can be edited while open
            self._restore_cloud_pbip()
        elif self._file_type == 'pbix':
            # PBIX needs to be closed first - show confirmation
            self._show_restore_cloud_pbix_dialog()
        else:
            ThemedMessageBox.showerror(self, "Error", "Unknown file type. Cannot restore connection.")

    def _restore_cloud_pbip(self):
        """Restore cloud connection for PBIP file (can edit while open)."""
        try:
            from tools.connection_hotswap.logic.pbix_modifier import get_modifier

            modifier = get_modifier()
            result = modifier.restore_cloud_connection(
                self.thin_report_file_path,
                self._cached_cloud_info,
                create_backup=True
            )

            if result.success:
                backup_msg = ""
                if result.backup_path:
                    backup_msg = f"\n\nBackup created: {os.path.basename(result.backup_path)}"

                # Clear the cache since we've restored
                modifier.clear_cached_cloud_connection(self.thin_report_file_path)

                ThemedMessageBox.showinfo(
                    self,
                    "Cloud Connection Restored",
                    f"Original cloud connection has been restored.\n\n"
                    f"Close and reopen the report to apply changes.{backup_msg}"
                )
                self.destroy()
            else:
                ThemedMessageBox.showerror(self, "Restore Failed", result.message)

        except Exception as e:
            ThemedMessageBox.showerror(self, "Error", f"Failed to restore cloud connection: {e}")

    def _show_restore_cloud_pbix_dialog(self):
        """Show confirmation dialog for PBIX cloud restore (requires close/reopen)."""
        colors = self._colors
        is_dark = self._is_dark

        # Create confirmation dialog
        confirm_dialog = tk.Toplevel(self)
        confirm_dialog.title("Restore Cloud Connection")
        confirm_dialog.geometry("400x250")
        confirm_dialog.resizable(False, False)
        confirm_dialog.transient(self)
        confirm_dialog.grab_set()
        confirm_dialog.configure(bg=colors['background'])

        # Center on parent
        confirm_dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 400) // 2
        y = self.winfo_y() + (self.winfo_height() - 250) // 2
        confirm_dialog.geometry(f"+{x}+{y}")

        # Set AE favicon icon
        try:
            base_dir = Path(__file__).parent.parent.parent.parent.parent
            icon_path = base_dir / 'assets' / 'favicon.ico'
            if icon_path.exists():
                confirm_dialog.iconbitmap(str(icon_path))
        except Exception:
            pass

        # Content
        content_frame = tk.Frame(confirm_dialog, bg=colors['background'], padx=20, pady=15)
        content_frame.pack(fill=tk.BOTH, expand=True)

        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f8f8fc')

        # Info section
        info_frame = tk.Frame(content_frame, bg=section_bg, padx=12, pady=10)
        info_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            info_frame,
            text="Restore Original Cloud Connection",
            font=("Segoe UI", 11, "bold"),
            bg=section_bg,
            fg=colors.get('primary', '#4a6cf5')
        ).pack(anchor='w')

        file_name = os.path.basename(self.thin_report_file_path) if self.thin_report_file_path else "Unknown"
        tk.Label(
            info_frame,
            text=f"File: {file_name}",
            font=("Segoe UI", 9),
            bg=section_bg,
            fg=colors['text_muted']
        ).pack(anchor='w', pady=(4, 0))

        # Explanation
        explain_text = (
            "This will restore the original Power BI Service connection.\n\n"
            "PBIX files are locked while open. We need to:\n"
            "1. Close Power BI Desktop\n"
            "2. Modify the file\n"
            "3. Reopen it"
        )
        tk.Label(
            content_frame,
            text=explain_text,
            font=("Segoe UI", 9),
            bg=colors['background'],
            fg=colors['text_primary'],
            justify='left'
        ).pack(anchor='w', pady=(0, 12))

        # Buttons
        button_frame = tk.Frame(content_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        def do_restore():
            confirm_dialog.destroy()
            self._execute_pbix_cloud_restore()

        RoundedButton(
            button_frame,
            text="RESTORE & REOPEN",
            command=do_restore,
            bg=colors.get('primary', '#4a6cf5'),
            hover_bg=colors.get('primary_hover', '#3d5dd5'),
            pressed_bg=colors.get('primary_pressed', '#3050c5'),
            fg='#ffffff',
            height=30, radius=5,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=colors['background']
        ).pack(side=tk.LEFT, padx=(0, 8))

        RoundedButton(
            button_frame,
            text="CANCEL",
            command=confirm_dialog.destroy,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=30, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background']
        ).pack(side=tk.LEFT)

    def _execute_pbix_cloud_restore(self):
        """Execute cloud restoration for PBIX file (close, modify, reopen)."""
        try:
            from tools.connection_hotswap.logic.pbix_modifier import get_modifier
            from tools.connection_hotswap.logic.process_control import PowerBIProcessController

            modifier = get_modifier()
            process_ctrl = PowerBIProcessController()

            # Close Power BI Desktop
            if self.thin_report_process_id:
                self._status_label.configure(text="Closing Power BI Desktop...")
                self._status_label.pack(fill=tk.X, pady=(0, 8))
                self.update()

                close_result = process_ctrl.close_powerbi_gracefully(
                    self.thin_report_process_id, timeout=10
                )
                if not close_result.success:
                    ThemedMessageBox.showwarning(
                        self,
                        "Could Not Close",
                        "Power BI Desktop could not be closed automatically.\n"
                        "Please save and close it manually, then try again."
                    )
                    self._status_label.pack_forget()
                    return

            # Wait for file to be unlocked
            self._status_label.configure(text="Waiting for file to be available...")
            self.update()

            unlock_result = process_ctrl.wait_for_file_unlock(
                self.thin_report_file_path, timeout=15
            )
            if not unlock_result.success:
                ThemedMessageBox.showerror(
                    self,
                    "File Locked",
                    "File is still locked. Please close Power BI Desktop and try again."
                )
                self._status_label.pack_forget()
                return

            # Restore the cloud connection
            self._status_label.configure(text="Restoring cloud connection...")
            self.update()

            result = modifier.restore_cloud_connection(
                self.thin_report_file_path,
                self._cached_cloud_info,
                create_backup=True
            )

            if result.success:
                # Clear the cache
                modifier.clear_cached_cloud_connection(self.thin_report_file_path)

                # Reopen the file
                self._status_label.configure(text="Reopening file...")
                self.update()

                process_ctrl.reopen_file(self.thin_report_file_path)

                ThemedMessageBox.showinfo(
                    self,
                    "Cloud Connection Restored",
                    "Original cloud connection has been restored.\n"
                    "The file is being reopened."
                )
                self.destroy()
            else:
                self._status_label.pack_forget()
                ThemedMessageBox.showerror(self, "Restore Failed", result.message)

        except ImportError as e:
            self._status_label.pack_forget()
            ThemedMessageBox.showerror(
                self,
                "Missing Module",
                f"Required module not available: {e}\n\n"
                "Please close Power BI Desktop manually, then restore."
            )
        except Exception as e:
            self._status_label.pack_forget()
            ThemedMessageBox.showerror(self, "Error", f"Failed to restore cloud connection: {e}")

    def _on_swap_connection(self):
        """Handle swap button click."""
        selection = self.model_tree.selection()
        if not selection:
            ThemedMessageBox.showinfo(self, "Select Model", "Please select a local model first.")
            return

        idx = int(selection[0])
        if idx >= len(self._sorted_models):
            return

        target_model = self._sorted_models[idx]

        if self._file_type == 'pbip':
            # PBIP can be edited while open - just modify and show message
            self._swap_pbip(target_model)
        elif self._file_type == 'pbix':
            # PBIX needs to be closed first - show confirmation dialog
            self._show_pbix_swap_dialog(target_model)
        else:
            ThemedMessageBox.showerror(self, "Error", "Unknown file type. Cannot swap connection.")

    def _swap_pbip(self, target_model: AvailableModel):
        """
        Swap PBIP connection directly (file is not locked while open).

        Args:
            target_model: The local model to connect to
        """
        try:
            from tools.connection_hotswap.logic.pbix_modifier import get_modifier

            modifier = get_modifier()
            result = modifier.swap_connection(
                self.thin_report_file_path,
                target_model.server,
                target_model.database_name,
                create_backup=True
            )

            if result.success:
                backup_msg = ""
                if result.backup_path:
                    backup_msg = f"\n\nBackup created: {os.path.basename(result.backup_path)}"

                ThemedMessageBox.showinfo(
                    self,
                    "Connection Updated",
                    f"Connection swapped to: {target_model.display_name}\n\n"
                    f"Close and reopen the report to apply changes.{backup_msg}"
                )
                self.destroy()
            else:
                ThemedMessageBox.showerror(self, "Swap Failed", result.message)

        except Exception as e:
            ThemedMessageBox.showerror(self, "Error", f"Failed to swap connection: {e}")

    def _show_pbix_swap_dialog(self, target_model: AvailableModel):
        """
        Show confirmation dialog for PBIX swap (requires close/reopen).

        Args:
            target_model: The local model to connect to
        """
        colors = self._colors
        is_dark = self._is_dark

        # Create confirmation dialog
        confirm_dialog = tk.Toplevel(self)
        confirm_dialog.title("PBIX Requires Close/Reopen")
        confirm_dialog.geometry("400x280")
        confirm_dialog.resizable(False, False)
        confirm_dialog.transient(self)
        confirm_dialog.grab_set()
        confirm_dialog.configure(bg=colors['background'])

        # Center on parent
        confirm_dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 400) // 2
        y = self.winfo_y() + (self.winfo_height() - 280) // 2
        confirm_dialog.geometry(f"+{x}+{y}")

        # Set AE favicon icon
        try:
            # Navigate from dialogs -> ui -> connection_hotswap -> tools -> src -> assets
            base_dir = Path(__file__).parent.parent.parent.parent.parent
            icon_path = base_dir / 'assets' / 'favicon.ico'
            if icon_path.exists():
                confirm_dialog.iconbitmap(str(icon_path))
        except Exception:
            pass

        # Content
        content_frame = tk.Frame(confirm_dialog, bg=colors['background'], padx=20, pady=15)
        content_frame.pack(fill=tk.BOTH, expand=True)

        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f8f8fc')

        # Info section
        info_frame = tk.Frame(content_frame, bg=section_bg, padx=12, pady=10)
        info_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            info_frame,
            text="PBIX File Requires Close/Reopen",
            font=("Segoe UI", 11, "bold"),
            bg=section_bg,
            fg=colors.get('warning', '#ff9800')
        ).pack(anchor='w')

        file_name = os.path.basename(self.thin_report_file_path) if self.thin_report_file_path else "Unknown"
        tk.Label(
            info_frame,
            text=f"File: {file_name}",
            font=("Segoe UI", 9),
            bg=section_bg,
            fg=colors['text_muted']
        ).pack(anchor='w', pady=(4, 0))

        # Explanation
        explain_text = (
            "PBIX files are locked while open. To swap the connection:\n\n"
            "1. Save your changes (optional)\n"
            "2. Close Power BI Desktop\n"
            "3. Modify the file\n"
            "4. Reopen the file"
        )
        tk.Label(
            content_frame,
            text=explain_text,
            font=("Segoe UI", 9),
            bg=colors['background'],
            fg=colors['text_primary'],
            justify='left'
        ).pack(anchor='w', pady=(0, 12))

        # Buttons
        button_frame = tk.Frame(content_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X)

        def on_cancel():
            confirm_dialog.destroy()

        def on_save_and_swap():
            confirm_dialog.destroy()
            self._execute_pbix_swap(target_model, save_first=True)

        def on_swap_without_save():
            confirm_dialog.destroy()
            self._execute_pbix_swap(target_model, save_first=False)

        cancel_btn = RoundedButton(
            button_frame,
            text="CANCEL",
            command=on_cancel,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=30, radius=5,
            font=('Segoe UI', 9),
            canvas_bg=colors['background']
        )
        cancel_btn.pack(side=tk.RIGHT)

        swap_no_save_btn = RoundedButton(
            button_frame,
            text="SWAP (NO SAVE)",
            command=on_swap_without_save,
            bg=colors.get('warning', '#ff9800'),
            hover_bg='#e68a00',
            pressed_bg='#cc7a00',
            fg='#ffffff',
            height=30, radius=5,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=colors['background']
        )
        swap_no_save_btn.pack(side=tk.RIGHT, padx=(0, 10))

        save_swap_btn = RoundedButton(
            button_frame,
            text="SAVE & SWAP",
            command=on_save_and_swap,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            height=30, radius=5,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=colors['background']
        )
        save_swap_btn.pack(side=tk.RIGHT, padx=(0, 10))

    def _execute_pbix_swap(self, target_model: AvailableModel, save_first: bool):
        """
        Execute the PBIX swap workflow: save -> close -> modify -> reopen.

        Args:
            target_model: The local model to connect to
            save_first: Whether to save before closing
        """
        from tools.connection_hotswap.logic.pbix_modifier import get_modifier
        from tools.connection_hotswap.logic.process_control import get_controller

        modifier = get_modifier()
        controller = get_controller()

        # Show integrated status
        self._show_status("Swapping connection...")

        # Disable buttons during swap
        if hasattr(self, 'swap_btn'):
            self.swap_btn.set_enabled(False)
        self.close_btn.set_enabled(False)

        def do_swap():
            try:
                # Update status for each step
                self.after(0, lambda: self._show_status("Saving file..." if save_first else "Closing Power BI..."))

                # Define the modification callback
                def modify_file(file_path):
                    self.after(0, lambda: self._show_status("Modifying file..."))
                    result = modifier.swap_connection(
                        file_path,
                        target_model.server,
                        target_model.database_name,
                        create_backup=True
                    )
                    return result.success, result.message

                # Execute the full workflow
                result = controller.save_close_and_reopen(
                    process_id=self.thin_report_process_id,
                    file_path=self.thin_report_file_path,
                    save_first=save_first,
                    modify_callback=modify_file
                )

                # Update UI on main thread
                self.after(0, lambda: self._on_swap_complete(result))

            except Exception as e:
                self.after(0, lambda: self._on_swap_error(str(e)))

        # Run in background thread
        thread = threading.Thread(target=do_swap, daemon=True)
        thread.start()

    def _show_status(self, message: str):
        """Show status message in the integrated status area."""
        if self._status_var and self._status_label:
            self._status_var.set(message)
            self._status_label.pack(fill=tk.X, pady=(0, 8))
            self.update_idletasks()

    def _hide_status(self):
        """Hide the integrated status area."""
        if self._status_label:
            self._status_label.pack_forget()

    def _on_swap_complete(self, result):
        """Handle swap completion."""
        self._hide_status()

        # Re-enable buttons
        if hasattr(self, 'swap_btn'):
            self.swap_btn.set_enabled(True)
        self.close_btn.set_enabled(True)

        if result.success:
            ThemedMessageBox.showinfo(self, "Success", "Connection swapped successfully!\n\nThe file has been reopened.")
            self.destroy()
        else:
            ThemedMessageBox.showerror(self, "Swap Failed", result.message)

    def _on_swap_error(self, error: str):
        """Handle swap error."""
        self._hide_status()

        # Re-enable buttons
        if hasattr(self, 'swap_btn'):
            self.swap_btn.set_enabled(True)
        self.close_btn.set_enabled(True)

        ThemedMessageBox.showerror(self, "Error", f"Swap failed: {error}")
