"""
Shared UI Widgets
Common reusable UI components for all tools.

Built by Reid Havens of Analytic Endeavors
Copyright (c) 2024 Analytic Endeavors LLC. All rights reserved.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable, List, Union


class AutoHideScrollbar(ttk.Scrollbar):
    """
    Scrollbar that automatically hides when content fits within the viewport.
    Supports both pack and grid geometry managers.
    """
    __slots__ = ('_pack_info', '_grid_info')

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._pack_info = {'side': tk.RIGHT, 'fill': tk.Y}
        self._grid_info = {}

    def set(self, first: str, last: str) -> None:
        """Override set to show/hide based on scroll position."""
        first_f, last_f = float(first), float(last)

        if first_f <= 0.0 and last_f >= 1.0:
            # Content fits entirely - hide scrollbar
            self.pack_forget()
        elif not self.winfo_ismapped():
            # Content needs scrolling and scrollbar is hidden - show it
            self.pack(**self._pack_info)

        super().set(first, last)

    def pack(self, **kwargs) -> None:
        """Store pack configuration for later re-packing."""
        if kwargs:
            self._pack_info = kwargs
        super().pack(**kwargs)

    def grid(self, **kwargs) -> None:
        """Store grid configuration for grid-based layouts."""
        if kwargs:
            self._grid_info = kwargs
        super().grid(**kwargs)

    def set_with_grid(self, first: str, last: str) -> None:
        """Alternative set method for grid-based layouts."""
        first_f, last_f = float(first), float(last)

        if first_f <= 0.0 and last_f >= 1.0:
            self.grid_remove()
        elif not self.winfo_ismapped() and self._grid_info:
            self.grid(**self._grid_info)

        super().set(first, last)

    def grid_set(self, first: str, last: str, row: int = 0, column: int = 1) -> None:
        """Legacy grid method - prefer set_with_grid() instead."""
        if float(first) <= 0.0 and float(last) >= 1.0:
            self.grid_remove()
        else:
            if not self.winfo_ismapped():
                self.grid(row=row, column=column, sticky="ns")
        ttk.Scrollbar.set(self, first, last)


class LoadingOverlay(tk.Frame):
    """
    Theme-aware loading overlay with animated spinner.
    """
    def __init__(self, parent, message: str = "Loading...", colors: dict = None):
        # Default colors for fallback
        if colors is None:
            colors = {
                'background': '#0d0d1a',
                'text_primary': '#ffffff',
                'text_muted': '#808090',
                'primary': '#009999'
            }

        super().__init__(parent, bg=colors.get('background', '#0d0d1a'))
        self._colors = colors

        # Animated spinner using Unicode characters
        self._spinner_chars = ['◐', '◓', '◑', '◒']
        self._spinner_idx = 0
        self._animating = False

        self._spinner_label = tk.Label(
            self,
            text='◐',
            font=('Segoe UI', 28),
            fg=colors.get('primary', '#009999'),
            bg=colors.get('background', '#0d0d1a')
        )
        self._spinner_label.place(relx=0.5, rely=0.42, anchor='center')

        self._message_label = tk.Label(
            self,
            text=message,
            font=('Segoe UI', 11),
            fg=colors.get('text_primary', '#ffffff'),
            bg=colors.get('background', '#0d0d1a')
        )
        self._message_label.place(relx=0.5, rely=0.52, anchor='center')

        self._progress_label = tk.Label(
            self,
            text='Please wait...',
            font=('Segoe UI', 9, 'italic'),
            fg=colors.get('text_muted', '#808090'),
            bg=colors.get('background', '#0d0d1a')
        )
        self._progress_label.place(relx=0.5, rely=0.58, anchor='center')

    def _animate(self) -> None:
        """Animate the spinner"""
        if self._animating:
            self._spinner_idx = (self._spinner_idx + 1) % 4
            self._spinner_label.config(text=self._spinner_chars[self._spinner_idx])
            self.after(120, self._animate)

    def update_message(self, message: str) -> None:
        """Update the main message"""
        self._message_label.config(text=message)

    def update_progress(self, text: str) -> None:
        """Update the progress text"""
        self._progress_label.config(text=text)
        self.update_idletasks()

    def show(self) -> None:
        """Show the overlay with animation"""
        self._animating = True
        self.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.lift()
        self._animate()

    def hide(self) -> None:
        """Hide the overlay"""
        self._animating = False
        self.place_forget()


class ToastNotification(tk.Toplevel):
    """
    Non-blocking toast notification that auto-dismisses.
    Appears at bottom-right of parent window.
    """

    # Toast type styling
    TOAST_STYLES = {
        'success': ('#10b981', '✓'),
        'error': ('#ef4444', '✗'),
        'warning': ('#f59e0b', '⚠'),
        'info': ('#009999', 'ℹ')
    }

    def __init__(self, parent, message: str, toast_type: str = 'info',
                 duration: int = 3000):
        super().__init__(parent)

        # Get styling for this toast type
        bg_color, icon = self.TOAST_STYLES.get(toast_type, self.TOAST_STYLES['info'])

        # Window setup - no decorations, positioned at bottom-right
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(bg=bg_color)

        # Rounded corners simulation with padding
        self.configure(padx=2, pady=2)

        # Content frame
        frame = tk.Frame(self, bg=bg_color, padx=15, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        # Icon
        tk.Label(
            frame,
            text=icon,
            font=('Segoe UI', 14),
            fg='white',
            bg=bg_color
        ).pack(side=tk.LEFT, padx=(0, 10))

        # Message
        tk.Label(
            frame,
            text=message,
            font=('Segoe UI', 10),
            fg='white',
            bg=bg_color,
            wraplength=280,
            justify=tk.LEFT
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Position bottom-right of parent
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + parent.winfo_width() - self.winfo_width() - 25
            y = parent.winfo_rooty() + parent.winfo_height() - self.winfo_height() - 70
            self.geometry(f'+{x}+{y}')
        except:
            # Fallback positioning
            self.geometry('+100+100')

        # Auto-dismiss after duration
        self.after(duration, self._fade_out)

    def _fade_out(self) -> None:
        """Destroy the toast"""
        try:
            self.destroy()
        except:
            pass


class SearchableTreeview(ttk.Frame):
    """
    Treeview with built-in search/filter functionality.
    """
    def __init__(self, parent, columns: tuple, show: str = "headings", **kwargs):
        super().__init__(parent)

        # Search bar
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._on_search_changed)

        search_frame = ttk.Frame(self)
        search_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(search_frame, text="🔍").pack(side=tk.LEFT, padx=(0, 5))
        self._search_entry = ttk.Entry(
            search_frame,
            textvariable=self._search_var,
            width=30
        )
        self._search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Treeview with scrollbar
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self._scrollbar = AutoHideScrollbar(tree_frame)
        self._scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show=show,
            yscrollcommand=self._scrollbar.set,
            **kwargs
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._scrollbar.config(command=self.tree.yview)

        self._all_items: list = []
        self._filter_callback: Optional[Callable] = None

    def set_filter_callback(self, callback: Callable[[str, dict], bool]) -> None:
        """Set custom filter function: callback(search_text, item_values) -> bool"""
        self._filter_callback = callback

    def _on_search_changed(self, *args) -> None:
        """Handle search text changes"""
        search_text = self._search_var.get().lower()

        # Clear current view
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Re-add matching items
        for item_id, values, tags in self._all_items:
            if self._matches_search(search_text, values):
                self.tree.insert("", tk.END, iid=item_id, values=values, tags=tags)

    def _matches_search(self, search_text: str, values: tuple) -> bool:
        """Check if item matches search"""
        if not search_text:
            return True

        if self._filter_callback:
            return self._filter_callback(search_text, values)

        # Default: check if search text is in any column
        return any(search_text in str(v).lower() for v in values)

    def insert_item(self, item_id: str, values: tuple, tags: tuple = ()) -> None:
        """Insert item (tracked for filtering)"""
        self._all_items.append((item_id, values, tags))
        self.tree.insert("", tk.END, iid=item_id, values=values, tags=tags)

    def clear(self) -> None:
        """Clear all items"""
        self._all_items.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)


class ThemedCombobox(tk.Frame):
    """
    A themed combobox that uses ThemedScrollbar for its dropdown list.
    Replaces ttk.Combobox to ensure consistent modern scrollbar styling.
    Supports readonly mode with a custom dropdown popup.
    """

    def __init__(
        self,
        parent,
        textvariable: Optional[tk.StringVar] = None,
        values: Optional[List[str]] = None,
        state: str = "readonly",
        font: Optional[tuple] = None,
        width: Optional[int] = None,
        theme_manager=None,
        **kwargs
    ):
        """
        Initialize the themed combobox.

        Args:
            parent: Parent widget
            textvariable: StringVar to bind to
            values: List of dropdown options
            state: "readonly" or "normal" (default readonly)
            font: Font tuple for display
            width: Character width of entry (approximate)
            theme_manager: Theme manager for colors
        """
        # Get theme manager
        if theme_manager is None:
            from core.theme_manager import get_theme_manager
            theme_manager = get_theme_manager()
        self._theme_manager = theme_manager

        # Get colors
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Frame colors - use darker/lighter backgrounds to stand out from row backgrounds
        if is_dark:
            # Use darkest background (#0d0d1a) to contrast with row zebra striping
            self._bg_color = '#0d0d1a'
            self._fg_color = colors.get('text_primary', '#e0e0e0')
            self._border_color = colors.get('border', '#3a3a4a')
            self._hover_bg = colors.get('surface', '#1a1a2e')
            self._dropdown_bg = '#0d0d1a'
            self._selection_bg = colors.get('selection_highlight', '#2a4a6a')
        else:
            # Use pure white (#ffffff) to contrast with row zebra striping
            self._bg_color = '#ffffff'
            self._fg_color = colors.get('text_primary', '#333333')
            self._border_color = colors.get('border', '#d8d8e0')
            self._hover_bg = colors.get('section_bg', '#f5f5fa')
            self._dropdown_bg = '#ffffff'
            self._selection_bg = colors.get('selection_highlight', '#cce5ff')

        super().__init__(parent, bg=self._border_color, **kwargs)

        # Store configuration
        self._values = list(values) if values else []
        self._state = state
        self._font = font or ("Segoe UI", 9)
        self._width = width or 20
        self._popup = None
        self._textvariable = textvariable or tk.StringVar()

        # Track if dropdown is open
        self._is_open = False

        # Create inner frame with background
        self._inner_frame = tk.Frame(self, bg=self._bg_color)
        self._inner_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # Create display label (for readonly mode)
        self._display = tk.Label(
            self._inner_frame,
            textvariable=self._textvariable,
            font=self._font,
            bg=self._bg_color,
            fg=self._fg_color,
            anchor="w",
            padx=5,
            pady=2
        )
        self._display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Calculate approximate pixel width from character width
        if width:
            # Approximate: 7 pixels per character for Segoe UI 9
            pixel_width = width * 7
            self._display.configure(width=width)

        # Dropdown arrow button
        self._arrow_btn = tk.Label(
            self._inner_frame,
            text="\u25BC",  # Down triangle
            font=("Segoe UI", 7),
            bg=self._bg_color,
            fg=self._fg_color,
            padx=5,
            pady=2,
            cursor="hand2"
        )
        self._arrow_btn.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind click events
        self._display.bind("<Button-1>", self._toggle_dropdown)
        self._arrow_btn.bind("<Button-1>", self._toggle_dropdown)
        self._inner_frame.bind("<Button-1>", self._toggle_dropdown)

        # Bind hover effects
        self._display.bind("<Enter>", self._on_hover_enter)
        self._display.bind("<Leave>", self._on_hover_leave)
        self._arrow_btn.bind("<Enter>", self._on_hover_enter)
        self._arrow_btn.bind("<Leave>", self._on_hover_leave)

        # Register for theme changes
        self._theme_manager.register_theme_callback(self._on_theme_changed)

    def _on_hover_enter(self, event=None):
        """Handle mouse enter for hover effect"""
        if self._state != "disabled":
            self._inner_frame.configure(bg=self._hover_bg)
            self._display.configure(bg=self._hover_bg)
            self._arrow_btn.configure(bg=self._hover_bg)

    def _on_hover_leave(self, event=None):
        """Handle mouse leave for hover effect"""
        if not self._is_open:
            self._inner_frame.configure(bg=self._bg_color)
            self._display.configure(bg=self._bg_color)
            self._arrow_btn.configure(bg=self._bg_color)

    def _toggle_dropdown(self, event=None):
        """Toggle the dropdown popup visibility"""
        if self._state == "disabled":
            return

        if self._popup and self._popup.winfo_exists():
            self._close_dropdown()
        else:
            self._open_dropdown()

    def _open_dropdown(self):
        """Open the dropdown popup with scrollable list"""
        if self._popup:
            self._close_dropdown()

        self._is_open = True
        colors = self._theme_manager.colors

        # Create popup window
        self._popup = tk.Toplevel(self)
        self._popup.withdraw()  # Hide until positioned
        self._popup.overrideredirect(True)  # No window decorations
        self._popup.configure(bg=self._border_color)

        # Main content frame
        main_frame = tk.Frame(self._popup, bg=self._dropdown_bg)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # Calculate max height for dropdown (show up to 10 items before scrolling)
        item_height = 22
        max_visible_items = 10
        visible_items = min(len(self._values), max_visible_items)
        list_height = max(visible_items * item_height, item_height)

        # Create listbox container
        list_frame = tk.Frame(main_frame, bg=self._dropdown_bg)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # Import ThemedScrollbar here to avoid circular import
        from core.ui_base import ThemedScrollbar

        # Create listbox
        self._listbox = tk.Listbox(
            list_frame,
            font=self._font,
            bg=self._dropdown_bg,
            fg=self._fg_color,
            selectbackground=self._selection_bg,
            selectforeground=self._fg_color,
            highlightthickness=0,
            borderwidth=0,
            activestyle="none",
            exportselection=False,
            height=visible_items if visible_items > 0 else 1
        )

        # Create themed scrollbar
        self._scrollbar = ThemedScrollbar(
            list_frame,
            command=self._listbox.yview,
            theme_manager=self._theme_manager,
            width=10,
            auto_hide=True
        )

        self._listbox.configure(yscrollcommand=self._scrollbar.set)

        # Pack listbox and scrollbar
        self._listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        if len(self._values) > max_visible_items:
            self._scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Populate listbox
        for value in self._values:
            self._listbox.insert(tk.END, value)

        # Select current value
        current = self._textvariable.get()
        if current in self._values:
            idx = self._values.index(current)
            self._listbox.selection_set(idx)
            self._listbox.see(idx)

        # Bind selection
        self._listbox.bind("<ButtonRelease-1>", self._on_select)
        self._listbox.bind("<Return>", self._on_select)
        self._listbox.bind("<Escape>", lambda e: self._close_dropdown())

        # Mouse wheel scrolling
        self._listbox.bind("<MouseWheel>", self._on_listbox_mousewheel)

        # Position popup below the combobox
        self.update_idletasks()
        self._popup.update_idletasks()

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        width = max(self.winfo_width(), 100)

        self._popup.geometry(f"{width}x{list_height}+{x}+{y}")
        self._popup.deiconify()
        self._popup.lift()
        self._listbox.focus_set()

        # Force scrollbar to update its thumb size after geometry is realized
        # This ensures accurate thumb representation on initial open
        def update_scrollbar():
            if self._popup and self._popup.winfo_exists() and hasattr(self, '_scrollbar'):
                self._popup.update_idletasks()
                # Get current listbox yview and re-apply to scrollbar
                first, last = self._listbox.yview()
                self._scrollbar.set(first, last)
                # Force scrollbar canvas to redraw
                self._scrollbar.update_idletasks()

        # Use longer delay to ensure geometry is fully calculated after theme changes
        self._popup.after(50, update_scrollbar)

        # Bind click outside to close
        self._popup.bind("<FocusOut>", self._on_focus_out)
        self.winfo_toplevel().bind("<Button-1>", self._on_click_outside, add="+")

    def _on_listbox_mousewheel(self, event):
        """Handle mousewheel scrolling in listbox"""
        self._listbox.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _on_select(self, event=None):
        """Handle item selection"""
        selection = self._listbox.curselection()
        if selection:
            idx = selection[0]
            value = self._values[idx]
            self._textvariable.set(value)
            # Generate <<ComboboxSelected>> event for compatibility
            self.event_generate("<<ComboboxSelected>>")
        self._close_dropdown()

    def _on_focus_out(self, event=None):
        """Handle focus leaving the popup"""
        # Delay close to allow click events to process
        if self._popup:
            self._popup.after(100, self._check_close)

    def _check_close(self):
        """Check if popup should close based on focus"""
        if self._popup and self._popup.winfo_exists():
            try:
                focus = self._popup.focus_get()
                if focus is None or not str(focus).startswith(str(self._popup)):
                    self._close_dropdown()
            except Exception:
                self._close_dropdown()

    def _on_click_outside(self, event):
        """Handle click outside the dropdown"""
        if not self._popup or not self._popup.winfo_exists():
            return

        # Check if click is inside dropdown or combobox
        x, y = event.x_root, event.y_root

        # Check dropdown bounds
        dx = self._popup.winfo_rootx()
        dy = self._popup.winfo_rooty()
        dw = self._popup.winfo_width()
        dh = self._popup.winfo_height()

        # Check combobox bounds
        cx = self.winfo_rootx()
        cy = self.winfo_rooty()
        cw = self.winfo_width()
        ch = self.winfo_height()

        in_dropdown = dx <= x <= dx + dw and dy <= y <= dy + dh
        in_combobox = cx <= x <= cx + cw and cy <= y <= cy + ch

        if not in_dropdown and not in_combobox:
            self._close_dropdown()

    def _close_dropdown(self):
        """Close the dropdown popup"""
        self._is_open = False
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None

        # Reset hover state
        self._on_hover_leave()

        # Unbind click outside
        try:
            self.winfo_toplevel().unbind("<Button-1>")
        except Exception:
            pass

    def _on_theme_changed(self, theme_name: str = None):
        """Update colors when theme changes"""
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Use darker/lighter backgrounds to stand out from row backgrounds
        if is_dark:
            # Use darkest background (#0d0d1a) to contrast with row zebra striping
            self._bg_color = '#0d0d1a'
            self._fg_color = colors.get('text_primary', '#e0e0e0')
            self._border_color = colors.get('border', '#3a3a4a')
            self._hover_bg = colors.get('surface', '#1a1a2e')
            self._dropdown_bg = '#0d0d1a'
            self._selection_bg = colors.get('selection_highlight', '#2a4a6a')
        else:
            # Use pure white (#ffffff) to contrast with row zebra striping
            self._bg_color = '#ffffff'
            self._fg_color = colors.get('text_primary', '#333333')
            self._border_color = colors.get('border', '#d8d8e0')
            self._hover_bg = colors.get('section_bg', '#f5f5fa')
            self._dropdown_bg = '#ffffff'
            self._selection_bg = colors.get('selection_highlight', '#cce5ff')

        # Update colors
        self.configure(bg=self._border_color)
        self._inner_frame.configure(bg=self._bg_color)
        self._display.configure(bg=self._bg_color, fg=self._fg_color)
        self._arrow_btn.configure(bg=self._bg_color, fg=self._fg_color)

        # Close dropdown if open (will reopen with new colors)
        if self._popup and self._popup.winfo_exists():
            self._close_dropdown()

    # =========================================================================
    # Public API (compatible with ttk.Combobox)
    # =========================================================================

    def get(self) -> str:
        """Get the current value"""
        return self._textvariable.get()

    def set(self, value: str):
        """Set the current value"""
        self._textvariable.set(value)

    def current(self, index: Optional[int] = None) -> Union[int, None]:
        """Get or set current selection by index"""
        if index is None:
            # Return current index
            current = self._textvariable.get()
            if current in self._values:
                return self._values.index(current)
            return -1
        else:
            # Set by index
            if 0 <= index < len(self._values):
                self._textvariable.set(self._values[index])

    def configure(self, **kwargs):
        """Configure widget options"""
        if 'values' in kwargs:
            self._values = list(kwargs.pop('values'))
        if 'state' in kwargs:
            self._state = kwargs.pop('state')
            if self._state == "disabled":
                self._display.configure(fg=self._theme_manager.colors.get('text_muted', '#808080'))
                self._arrow_btn.configure(fg=self._theme_manager.colors.get('text_muted', '#808080'))
            else:
                self._display.configure(fg=self._fg_color)
                self._arrow_btn.configure(fg=self._fg_color)
        if 'textvariable' in kwargs:
            self._textvariable = kwargs.pop('textvariable')
            self._display.configure(textvariable=self._textvariable)
        if 'font' in kwargs:
            self._font = kwargs.pop('font')
            self._display.configure(font=self._font)

        # Handle remaining kwargs with parent configure
        if kwargs:
            super().configure(**kwargs)

    config = configure  # Alias

    def __getitem__(self, key):
        """Support bracket notation for getting config"""
        if key == 'values':
            return tuple(self._values)
        elif key == 'state':
            return self._state
        return super().__getitem__(key)

    def __setitem__(self, key, value):
        """Support bracket notation for setting config"""
        self.configure(**{key: value})

    def bind(self, sequence=None, func=None, add=None):
        """Bind events - pass through to display for compatibility"""
        if sequence == "<<ComboboxSelected>>":
            # Bind to the frame itself for this virtual event
            return super().bind(sequence, func, add)
        else:
            # Bind to display label for other events
            return self._display.bind(sequence, func, add)

    def destroy(self):
        """Clean up on destroy"""
        self._close_dropdown()
        try:
            self._theme_manager.unregister_theme_callback(self._on_theme_changed)
        except Exception:
            pass
        super().destroy()


# Module fingerprint for integrity verification
_AE_SIG = "QUU6UmVpZEhhdmVuc0BBbmFseXRpY0VuZGVhdm9ycw=="
