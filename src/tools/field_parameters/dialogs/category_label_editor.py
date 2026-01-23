"""
Category Label Editor Dialog
Modal dialog for editing labels within a category column.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk
from typing import TYPE_CHECKING

from core.theme_manager import get_theme_manager
from core.ui_base import ThemedScrollbar, ThemedMessageBox, ThemedInputDialog

if TYPE_CHECKING:
    from tools.field_parameters.field_parameters_core import CategoryLevel


class CategoryLabelEditorDialog(tk.Toplevel):
    """Dialog for editing labels within a category column"""

    __slots__ = (
        'category_level', 'on_save', 'labels', 'new_label_var', 'new_label_entry',
        'listbox', 'rename_btn', 'remove_btn', 'move_up_btn', 'move_down_btn',
        '_theme_manager'
    )

    def __init__(self, parent, category_level: 'CategoryLevel', on_save: callable):
        super().__init__(parent)
        self._theme_manager = get_theme_manager()
        self.category_level = category_level
        self.on_save = on_save
        self.labels = list(category_level.labels)  # Work with a copy

        self.title(f"Edit Labels: {category_level.name}")
        self.geometry("350x400")
        self.resizable(True, True)
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
        self._center_on_parent(parent)

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
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        ttk.Label(
            main_frame,
            text=f"Labels for '{self.category_level.name}'",
            font=("Segoe UI", 10, "bold")
        ).pack(fill=tk.X, pady=(0, 5))

        ttk.Label(
            main_frame,
            text="Add labels that can be assigned to fields.\nOrder determines sort value.",
            font=("Segoe UI", 8, "italic"),
            foreground=colors['text_muted']
        ).pack(fill=tk.X, pady=(0, 10))

        # Add label entry
        add_frame = ttk.Frame(main_frame)
        add_frame.pack(fill=tk.X, pady=(0, 5))

        self.new_label_var = tk.StringVar()
        self.new_label_entry = ttk.Entry(add_frame, textvariable=self.new_label_var)
        self.new_label_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.new_label_entry.bind("<Return>", lambda e: self._add_label())

        ttk.Button(add_frame, text="+ Add", width=8, command=self._add_label).pack(side=tk.RIGHT)

        # Labels listbox with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 5))

        self._scrollbar = ThemedScrollbar(
            list_frame,
            command=None,  # Set after listbox is created
            theme_manager=self._theme_manager,
            width=12,
            auto_hide=True
        )
        self._scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(
            list_frame,
            yscrollcommand=self._scrollbar.set,
            font=("Segoe UI", 9),
            selectmode=tk.SINGLE,
            bg=colors.get('section_bg', colors.get('surface', colors['background'])),
            fg=colors['text_primary'],
            selectbackground=colors['selection_highlight'],
            selectforeground=colors['text_primary'],
            highlightthickness=0,
            borderwidth=1,
            relief='solid'
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._scrollbar._command = self.listbox.yview
        self.listbox.bind("<<ListboxSelect>>", self._on_selection_changed)
        self.listbox.bind("<Double-Button-1>", lambda e: self._rename_label())

        # Action buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(5, 10))

        self.rename_btn = ttk.Button(
            action_frame, text="Rename", command=self._rename_label, state="disabled"
        )
        self.rename_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.remove_btn = ttk.Button(
            action_frame, text="Remove", command=self._remove_label, state="disabled"
        )
        self.remove_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.move_up_btn = ttk.Button(
            action_frame, text="^", width=3, command=self._move_up, state="disabled"
        )
        self.move_up_btn.pack(side=tk.LEFT, padx=(0, 2))

        self.move_down_btn = ttk.Button(
            action_frame, text="v", width=3, command=self._move_down, state="disabled"
        )
        self.move_down_btn.pack(side=tk.LEFT)

        # Save/Cancel buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="Save", command=self._save).pack(side=tk.RIGHT)

        # Populate list
        self._refresh_list()

    def _refresh_list(self):
        """Refresh the labels listbox"""
        self.listbox.delete(0, tk.END)
        for idx, label in enumerate(self.labels, 1):
            self.listbox.insert(tk.END, f"{idx}. {label}")

    def _add_label(self):
        """Add a new label"""
        label = self.new_label_var.get().strip()
        if not label:
            return
        if label in self.labels:
            ThemedMessageBox.showerror(
                self,
                "Duplicate Label",
                f"A label named '{label}' already exists.\n\nPlease choose a different name."
            )
            return
        self.labels.append(label)
        self._refresh_list()
        self.new_label_var.set("")
        self.new_label_entry.focus()

    def _remove_label(self):
        """Remove selected label"""
        selection = self.listbox.curselection()
        if selection:
            del self.labels[selection[0]]
            self._refresh_list()
            self._on_selection_changed(None)

    def _move_up(self):
        """Move selected label up"""
        selection = self.listbox.curselection()
        if selection and selection[0] > 0:
            idx = selection[0]
            self.labels[idx], self.labels[idx-1] = self.labels[idx-1], self.labels[idx]
            self._refresh_list()
            self.listbox.selection_set(idx-1)
            self._on_selection_changed(None)

    def _move_down(self):
        """Move selected label down"""
        selection = self.listbox.curselection()
        if selection and selection[0] < len(self.labels) - 1:
            idx = selection[0]
            self.labels[idx], self.labels[idx+1] = self.labels[idx+1], self.labels[idx]
            self._refresh_list()
            self.listbox.selection_set(idx+1)
            self._on_selection_changed(None)

    def _on_selection_changed(self, event):
        """Handle selection change"""
        selection = self.listbox.curselection()
        if selection:
            idx = selection[0]
            self.rename_btn.config(state="normal")
            self.remove_btn.config(state="normal")
            self.move_up_btn.config(state="normal" if idx > 0 else "disabled")
            self.move_down_btn.config(state="normal" if idx < len(self.labels)-1 else "disabled")
        else:
            self.rename_btn.config(state="disabled")
            self.remove_btn.config(state="disabled")
            self.move_up_btn.config(state="disabled")
            self.move_down_btn.config(state="disabled")

    def _rename_label(self):
        """Rename selected label"""
        selection = self.listbox.curselection()
        if not selection:
            return

        idx = selection[0]
        old_label = self.labels[idx]

        new_label = ThemedInputDialog.askstring(
            self,
            "Rename Label",
            f"Enter new name for '{old_label}':",
            initialvalue=old_label
        )

        if new_label and new_label.strip() and new_label.strip() != old_label:
            new_label = new_label.strip()
            if new_label in self.labels:
                ThemedMessageBox.showwarning(self, "Duplicate", f"Label '{new_label}' already exists.")
                return

            self.labels[idx] = new_label
            self._refresh_list()
            self.listbox.selection_set(idx)
            self._on_selection_changed(None)

    def _save(self):
        """Save labels and close"""
        self.category_level.labels = self.labels
        self.on_save()
        self.destroy()
