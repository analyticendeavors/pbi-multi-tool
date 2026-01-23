"""
Schema Mismatch Warning Dialog
Built by Reid Havens of Analytic Endeavors

Modal dialog for warning users when a preset's cloud connection schema
differs from the current file's schema. This helps ensure cloud connections
are restored with the correct format (PbiServiceModelId, etc.).

Options:
1. UPDATE PRESET - Update the preset with the current schema, then apply
2. APPLY ANYWAY - Apply the preset's stored connection (may have issues)
3. CANCEL - Abort the operation
"""

import os
import tkinter as tk
from tkinter import ttk
from typing import Optional
from enum import Enum

from core.theme_manager import get_theme_manager


class SchemaMismatchAction(Enum):
    """Actions the user can take when schema mismatch is detected."""
    UPDATE_PRESET = "update"
    APPLY_ANYWAY = "apply"
    CANCEL = "cancel"


class SchemaMismatchDialog(tk.Toplevel):
    """
    Warning dialog for schema mismatches between preset and current file.

    Shows a comparison of the differences and lets the user choose
    how to proceed.
    """

    def __init__(
        self,
        parent,
        preset_name: str,
        mismatched_mappings: list,
        details: str = "",
    ):
        """
        Initialize the schema mismatch dialog.

        Args:
            parent: Parent window
            preset_name: Name of the preset being applied
            mismatched_mappings: List of connection names with schema mismatches
            details: Detailed diff information
        """
        super().__init__(parent)

        self.preset_name = preset_name
        self.mismatched_mappings = mismatched_mappings
        self.details = details
        self.result: SchemaMismatchAction = SchemaMismatchAction.CANCEL

        # Get theme manager
        self._theme_manager = get_theme_manager()
        self._colors = self._theme_manager.colors
        self._is_dark = self._theme_manager.is_dark

        self._setup_window()
        self._setup_ui()

        # Wait for dialog to close
        self.wait_window()

    def _setup_window(self):
        """Configure the dialog window."""
        colors = self._colors

        self.title("Schema Mismatch Detected")
        self.geometry("500x380")
        self.resizable(True, True)
        self.minsize(450, 320)

        # Find the top-level window for proper centering
        top_window = self.master.winfo_toplevel()
        self.transient(top_window)
        self.grab_set()
        self.configure(bg=colors['background'])

        # Center on parent
        self.update_idletasks()
        dialog_width = 500
        dialog_height = 380
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

        # Set dialog icon
        try:
            icon_path = os.path.join(
                os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))),
                'assets', 'favicon.ico'
            )
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

    def _setup_ui(self):
        """Setup the dialog UI."""
        colors = self._colors
        is_dark = self._is_dark
        section_bg = colors.get('section_bg', '#1a1a2e' if is_dark else '#f8f8fc')
        warning_color = '#e6a700' if is_dark else '#d4a000'

        # Main container
        main_frame = tk.Frame(self, bg=colors['background'], padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Warning header
        header_frame = tk.Frame(main_frame, bg=colors['background'])
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # Warning icon (triangle with exclamation)
        tk.Label(
            header_frame,
            text="[!]",
            font=("Segoe UI", 16, "bold"),
            bg=colors['background'],
            fg=warning_color
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(
            header_frame,
            text="Cloud Connection Schema Mismatch",
            font=("Segoe UI", 12, "bold"),
            bg=colors['background'],
            fg=colors['text_primary']
        ).pack(side=tk.LEFT)

        # Explanation
        explanation_text = (
            f"The preset \"{self.preset_name}\" was saved with a different "
            "cloud connection schema than the current file.\n\n"
            "This may occur if:\n"
            "  - The dataset was republished to a different location\n"
            "  - A different workspace or model is being used\n"
            "  - Power BI updated the connection format"
        )

        explanation_label = tk.Label(
            main_frame,
            text=explanation_text,
            font=("Segoe UI", 9),
            bg=colors['background'],
            fg=colors['text_muted'],
            justify=tk.LEFT,
            wraplength=450
        )
        explanation_label.pack(fill=tk.X, pady=(0, 15))

        # Details section (expandable)
        if self.details:
            details_frame = tk.LabelFrame(
                main_frame,
                text="Differences",
                font=("Segoe UI", 9),
                bg=section_bg,
                fg=colors['text_muted'],
                padx=10,
                pady=8
            )
            details_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

            # Scrollable text for details
            text_frame = tk.Frame(details_frame, bg=section_bg)
            text_frame.pack(fill=tk.BOTH, expand=True)

            details_text = tk.Text(
                text_frame,
                font=("Consolas", 9),
                bg=section_bg,
                fg=colors['text_primary'],
                relief=tk.FLAT,
                wrap=tk.WORD,
                height=6
            )
            details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=details_text.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            details_text.configure(yscrollcommand=scrollbar.set)

            details_text.insert(tk.END, self.details)
            details_text.configure(state=tk.DISABLED)
        else:
            # Simple message if no details
            affected_text = f"Affected connections: {', '.join(self.mismatched_mappings)}"
            tk.Label(
                main_frame,
                text=affected_text,
                font=("Segoe UI", 9),
                bg=colors['background'],
                fg=colors['text_muted'],
                wraplength=450
            ).pack(fill=tk.X, pady=(0, 15))

        # Buttons frame
        button_frame = tk.Frame(main_frame, bg=colors['background'])
        button_frame.pack(fill=tk.X, pady=(10, 0))

        # Button styling
        button_font = ("Segoe UI", 9)
        button_padx = 15
        button_pady = 6

        # Cancel button (leftmost)
        cancel_btn = tk.Button(
            button_frame,
            text="Cancel",
            font=button_font,
            bg=colors.get('button_bg', '#3d3d5c'),
            fg=colors['text_primary'],
            relief=tk.FLAT,
            cursor="hand2",
            padx=button_padx,
            pady=button_pady,
            command=self._on_cancel
        )
        cancel_btn.pack(side=tk.LEFT)

        # Update Preset button (primary action - rightmost)
        update_btn = tk.Button(
            button_frame,
            text="Update Preset",
            font=button_font,
            bg=colors.get('accent', '#4a90d9'),
            fg='white',
            relief=tk.FLAT,
            cursor="hand2",
            padx=button_padx,
            pady=button_pady,
            command=self._on_update
        )
        update_btn.pack(side=tk.RIGHT)

        # Apply Anyway button (secondary - middle)
        apply_btn = tk.Button(
            button_frame,
            text="Apply Anyway",
            font=button_font,
            bg=colors.get('button_bg', '#3d3d5c'),
            fg=colors['text_primary'],
            relief=tk.FLAT,
            cursor="hand2",
            padx=button_padx,
            pady=button_pady,
            command=self._on_apply_anyway
        )
        apply_btn.pack(side=tk.RIGHT, padx=(0, 10))

        # Button hover effects
        for btn in [cancel_btn, apply_btn]:
            btn.bind('<Enter>', lambda e, b=btn: b.configure(bg=colors.get('button_hover', '#4d4d70')))
            btn.bind('<Leave>', lambda e, b=btn: b.configure(bg=colors.get('button_bg', '#3d3d5c')))

        update_btn.bind('<Enter>', lambda e: update_btn.configure(bg=colors.get('accent_hover', '#5a9de9')))
        update_btn.bind('<Leave>', lambda e: update_btn.configure(bg=colors.get('accent', '#4a90d9')))

        # Bind Escape to cancel
        self.bind('<Escape>', lambda e: self._on_cancel())

    def _on_update(self):
        """Handle Update Preset button click."""
        self.result = SchemaMismatchAction.UPDATE_PRESET
        self.destroy()

    def _on_apply_anyway(self):
        """Handle Apply Anyway button click."""
        self.result = SchemaMismatchAction.APPLY_ANYWAY
        self.destroy()

    def _on_cancel(self):
        """Handle Cancel button click."""
        self.result = SchemaMismatchAction.CANCEL
        self.destroy()


def show_schema_mismatch_warning(
    parent,
    preset_name: str,
    mismatched_mappings: list,
    details: str = ""
) -> SchemaMismatchAction:
    """
    Show the schema mismatch warning dialog and return the user's choice.

    Args:
        parent: Parent window
        preset_name: Name of the preset being applied
        mismatched_mappings: List of connection names with schema mismatches
        details: Detailed diff information

    Returns:
        SchemaMismatchAction enum value
    """
    dialog = SchemaMismatchDialog(
        parent,
        preset_name=preset_name,
        mismatched_mappings=mismatched_mappings,
        details=details
    )
    return dialog.result
