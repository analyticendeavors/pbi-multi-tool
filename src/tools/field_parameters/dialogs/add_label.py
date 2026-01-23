"""
Add Label Dialog
Modal dialog for adding new category labels with position options.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Optional

from core.theme_manager import get_theme_manager
from core.ui_base import ThemedMessageBox, RoundedButton, LabeledRadioGroup, ThemedScrollbar


class AddLabelDialog(tk.Toplevel):
    """Dialog for adding new label(s) with position options"""

    __slots__ = (
        'category_name', 'existing_labels', 'selected_label_idx',
        'result_labels', 'result_position', 'mode_var', 'single_frame',
        'multi_frame', 'name_var', 'name_entry', 'multi_text', 'position_var',
        '_theme_manager', 'mode_radio_group', 'position_radio_group',
        'relative_position_radio_group', '_scrollbar'
    )

    def __init__(self, parent, category_name: str, existing_labels: List[str],
                 selected_label_idx: Optional[int] = None):
        super().__init__(parent)
        self.withdraw()  # Hide until positioned to prevent flicker
        self._theme_manager = get_theme_manager()
        self.category_name = category_name
        self.existing_labels = existing_labels
        self.selected_label_idx = selected_label_idx

        # Results - can be single or multiple labels
        self.result_labels: List[str] = []
        self.result_position: Optional[int] = None

        self.title("Add Label(s)")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # Set AE favicon
        try:
            from pathlib import Path
            base_path = Path(__file__).parent.parent.parent.parent
            favicon_path = base_path / "assets" / "favicon.ico"
            if favicon_path.exists():
                self.iconbitmap(str(favicon_path))
        except Exception:
            pass

        # Set dark title bar on Windows
        try:
            from ctypes import windll, byref, sizeof, c_int
            HWND = windll.user32.GetParent(self.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            value = c_int(1 if self._theme_manager.is_dark else 0)
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_USE_IMMERSIVE_DARK_MODE, byref(value), sizeof(value))
        except Exception:
            pass

        self._setup_ui()

        # Size dialog and center on parent in one call to prevent flicker
        self.update_idletasks()
        width = max(380, self.winfo_reqwidth())
        height = self.winfo_reqheight() + 10
        # Calculate center position
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        x = px + (pw - width) // 2
        y = py + (ph - height) // 2
        # Set size and position in one call, then show
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.deiconify()

        self.name_entry.focus_set()
        self.bind("<Escape>", lambda e: self.destroy())

    def _center_on_parent(self, parent):
        """Center dialog on parent window"""
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    def _setup_ui(self):
        """Setup the dialog UI"""
        colors = self._theme_manager.colors
        bg_color = colors['background']

        # Configure dialog background
        self.configure(bg=bg_color)

        main_frame = tk.Frame(self, bg=bg_color, padx=12, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        tk.Label(
            main_frame, text=f"Add label(s) to '{self.category_name}':",
            font=("Segoe UI", 9, "bold"),
            bg=bg_color, fg=colors['text_primary']
        ).pack(anchor="w")

        # Input mode toggle
        self.mode_var = tk.StringVar(value="single")
        mode_frame = tk.Frame(main_frame, bg=bg_color)
        mode_frame.pack(fill=tk.X, pady=(8, 5))

        # Radio Group for mode toggle
        self.mode_radio_group = LabeledRadioGroup(
            mode_frame,
            variable=self.mode_var,
            options=[
                ("single", "Single label"),
                ("multi", "Multiple labels (one per line)"),
            ],
            command=self._on_mode_change,
            orientation="horizontal",
            font=('Segoe UI', 9),
            padding=12
        )
        self.mode_radio_group.pack(side=tk.LEFT)

        # Single entry frame
        self.single_frame = tk.Frame(main_frame, bg=bg_color)
        self.single_frame.pack(fill=tk.X, pady=(5, 10))

        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(self.single_frame, textvariable=self.name_var, width=45)
        self.name_entry.pack(fill=tk.X)
        self.name_entry.bind("<Return>", lambda e: self._on_ok())

        # Multi-line text frame (hidden by default)
        self.multi_frame = tk.Frame(main_frame, bg=bg_color)

        tk.Label(
            self.multi_frame, text="Enter each label on a separate line:",
            font=("Segoe UI", 8, "italic"),
            bg=bg_color, fg=colors['text_muted']
        ).pack(anchor="w")

        text_container = tk.Frame(self.multi_frame, bg=bg_color)
        text_container.pack(fill=tk.BOTH, expand=True, pady=(3, 0))

        text_bg = colors.get('section_bg', colors.get('surface', bg_color))
        self.multi_text = tk.Text(
            text_container, width=45, height=5, font=("Segoe UI", 9),
            bg=text_bg, fg=colors['text_primary'],
            insertbackground=colors['text_primary'],
            padx=4, pady=4
        )
        self.multi_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._scrollbar = ThemedScrollbar(
            text_container,
            command=self.multi_text.yview,
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.multi_text.config(yscrollcommand=self._scrollbar.set)

        # Position options section
        pos_section = tk.Frame(main_frame, bg=bg_color)
        pos_section.pack(fill=tk.X, pady=(5, 0))

        tk.Label(
            pos_section, text="Insert position:",
            font=("Segoe UI", 9),
            bg=bg_color, fg=colors['text_primary']
        ).pack(anchor="w")

        self.position_var = tk.StringVar(value="end")
        positions_frame = tk.Frame(pos_section, bg=bg_color)
        positions_frame.pack(fill=tk.X, pady=(3, 8))

        # Show relative position options if a label is selected
        self.relative_position_radio_group = None
        if self.selected_label_idx is not None and self.existing_labels:
            selected_name = self.existing_labels[self.selected_label_idx]
            display_name = selected_name if len(selected_name) <= 20 else selected_name[:17] + "..."

            row1 = tk.Frame(positions_frame, bg=bg_color)
            row1.pack(fill=tk.X)

            self.relative_position_radio_group = LabeledRadioGroup(
                row1,
                variable=self.position_var,
                options=[
                    ("above", f"Above '{display_name}'"),
                    ("below", f"Below '{display_name}'"),
                ],
                orientation="horizontal",
                font=('Segoe UI', 9),
                padding=12
            )
            self.relative_position_radio_group.pack(side=tk.LEFT)
            self.position_var.set("below")

        row2 = tk.Frame(positions_frame, bg=bg_color)
        row2.pack(fill=tk.X, pady=(2, 0))

        self.position_radio_group = LabeledRadioGroup(
            row2,
            variable=self.position_var,
            options=[
                ("top", "Top of list"),
                ("end", "Bottom of list"),
            ],
            orientation="horizontal",
            font=('Segoe UI', 9),
            padding=12
        )
        self.position_radio_group.pack(side=tk.LEFT)

        if self.selected_label_idx is None:
            self.position_var.set("end")

        # Buttons
        btn_frame = tk.Frame(main_frame, bg=bg_color)
        btn_frame.pack(fill=tk.X, pady=(8, 0))

        cancel_btn = RoundedButton(
            btn_frame, text="Cancel", command=self.destroy,
            bg=colors['button_secondary'],
            hover_bg=colors['button_secondary_hover'],
            pressed_bg=colors['button_secondary_pressed'],
            fg=colors['text_primary'],
            height=38, radius=6,
            font=('Segoe UI', 9),
            canvas_bg=bg_color
        )
        cancel_btn.pack(side=tk.RIGHT, padx=(8, 0))

        add_btn = RoundedButton(
            btn_frame, text="Add", command=self._on_ok,
            bg=colors['button_primary'],
            hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'],
            fg='#ffffff',
            height=38, radius=6,
            font=('Segoe UI', 9, 'bold'),
            canvas_bg=bg_color
        )
        add_btn.pack(side=tk.RIGHT)

    def _on_mode_change(self):
        """Handle mode toggle between single and multi-line"""
        if self.mode_var.get() == "multi":
            self.single_frame.pack_forget()
            self.multi_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 10),
                                 after=self.single_frame.master.winfo_children()[1])
            text = self.name_var.get().strip()
            if text:
                self.multi_text.insert("1.0", text)
            self.multi_text.focus_set()
            self.update_idletasks()
            self.geometry(f"{self.winfo_width()}x{self.winfo_reqheight() + 10}")
        else:
            self.multi_frame.pack_forget()
            self.single_frame.pack(fill=tk.X, pady=(5, 10),
                                  after=self.single_frame.master.winfo_children()[1])
            text = self.multi_text.get("1.0", tk.END).strip().split('\n')[0]
            self.name_var.set(text)
            self.name_entry.focus_set()
            self.update_idletasks()
            self.geometry(f"{self.winfo_width()}x{self.winfo_reqheight() + 10}")

    def _on_ok(self):
        """Handle Add button"""
        if self.mode_var.get() == "multi":
            text = self.multi_text.get("1.0", tk.END).strip()
            if not text:
                ThemedMessageBox.showwarning(self, "Missing Labels", "Please enter at least one label.")
                return

            lines = [line.strip() for line in text.split('\n') if line.strip()]
            if not lines:
                ThemedMessageBox.showwarning(self, "Missing Labels", "Please enter at least one label.")
                return

            # Check for duplicates with existing
            duplicates = [l for l in lines if l in self.existing_labels]
            if duplicates:
                ThemedMessageBox.showerror(
                    self,
                    "Duplicate Labels",
                    f"These labels already exist:\n\n" +
                    "\n".join(f"  - {d}" for d in duplicates[:5]) +
                    ("\n  ..." if len(duplicates) > 5 else "")
                )
                return

            # Check for duplicates within input
            seen = set()
            input_dupes = [l for l in lines if l in seen or seen.add(l)]
            if input_dupes:
                ThemedMessageBox.showerror(
                    self,
                    "Duplicate Input",
                    f"These labels are entered multiple times:\n\n" +
                    "\n".join(f"  - {d}" for d in input_dupes[:5])
                )
                return

            self.result_labels = lines
        else:
            name = self.name_var.get().strip()
            if not name:
                ThemedMessageBox.showwarning(self, "Missing Name", "Please enter a label name.")
                return

            if name in self.existing_labels:
                ThemedMessageBox.showerror(self, "Duplicate Label", f"A label named '{name}' already exists.")
                return

            self.result_labels = [name]

        # Calculate insert position
        pos = self.position_var.get()
        if pos == "top":
            self.result_position = 0
        elif pos == "end":
            self.result_position = len(self.existing_labels)
        elif pos == "above" and self.selected_label_idx is not None:
            self.result_position = self.selected_label_idx
        elif pos == "below" and self.selected_label_idx is not None:
            self.result_position = self.selected_label_idx + 1
        else:
            self.result_position = len(self.existing_labels)

        self.destroy()
