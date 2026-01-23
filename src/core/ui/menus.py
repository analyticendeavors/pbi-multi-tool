"""
Themed Menu Components
Reusable context menu with modern styling.

Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk


class ThemedContextMenu:
    """
    Themed context menu using custom Toplevel for modern styling.

    Uses a Toplevel window instead of tk.Menu to provide better border control
    and consistent styling across platforms (tk.Menu has intense white borders on Windows).

    Features:
    - 1px themed border (uses 'border' color)
    - Hover effects on items
    - Separator support
    - Section headers (muted text)
    - Escape to close
    - Click-outside to close

    Usage:
        menu = ThemedContextMenu(parent, theme_manager)
        menu.add_command("Edit", on_edit)
        menu.add_separator()
        menu.add_section_header("Actions")
        menu.add_command("Delete", on_delete)
        menu.show(event.x_root, event.y_root)
    """

    def __init__(self, parent, theme_manager):
        """
        Create a themed context menu.

        Args:
            parent: Parent widget to attach the menu to
            theme_manager: ThemeManager instance for colors
        """
        self._parent = parent
        self._theme_manager = theme_manager
        self._popup = None
        self._outside_click_bind_id = None
        self._items = []  # Store items to build later: ('command', label, cmd, icon) or ('separator',) or ('header', text)

    def add_command(self, label: str, command, icon=None, enabled=True):
        """
        Add a clickable menu item.

        Args:
            label: Text to display
            command: Function to call when clicked
            icon: Optional PhotoImage icon to display
            enabled: Whether the item is clickable (default True). Disabled items appear muted.
        """
        self._items.append(('command', label, command, icon, enabled))

    def add_separator(self):
        """Add a visual separator line."""
        self._items.append(('separator',))

    def add_section_header(self, text: str):
        """
        Add a section header (muted, non-clickable text).

        Args:
            text: Header text to display
        """
        self._items.append(('header', text))

    def show(self, x: int, y: int):
        """
        Display the menu at the specified screen coordinates.

        Args:
            x: Screen x coordinate
            y: Screen y coordinate
        """
        # Close any existing popup
        self.close()

        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Create popup window - attached to parent, not topmost
        self._popup = tk.Toplevel(self._parent)
        self._popup.withdraw()  # Hide until positioned
        self._popup.overrideredirect(True)  # No window decorations

        # Configure popup appearance
        popup_bg = colors.get('surface', '#1e1e2e' if is_dark else '#ffffff')
        border_color = colors.get('border', '#3a3a4a' if is_dark else '#d8d8e0')
        text_color = colors['text_primary']
        hover_bg = colors.get('hover', '#2a2a3e' if is_dark else '#f0f0f5')
        muted_color = colors.get('text_muted', '#888888')

        # Border frame (1px border via padx/pady)
        border_frame = tk.Frame(self._popup, bg=border_color, padx=1, pady=1)
        border_frame.pack(fill=tk.BOTH, expand=True)

        # Main content frame
        main_frame = tk.Frame(border_frame, bg=popup_bg)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Build menu items
        for item in self._items:
            if item[0] == 'separator':
                # Separator line
                sep = tk.Frame(main_frame, height=1, bg=border_color)
                sep.pack(fill=tk.X, padx=8, pady=4)

            elif item[0] == 'header':
                # Section header (muted, non-clickable)
                header_text = item[1]
                header_label = tk.Label(
                    main_frame,
                    text=header_text,
                    font=('Segoe UI', 8),
                    fg=muted_color,
                    bg=popup_bg,
                    anchor='w',
                    padx=12
                )
                header_label.pack(fill=tk.X, pady=4)

            elif item[0] == 'command':
                # Clickable menu item
                label_text, command, icon, enabled = item[1], item[2], item[3], item[4]

                item_frame = tk.Frame(main_frame, bg=popup_bg)
                item_frame.pack(fill=tk.X)

                # Use muted color for disabled items
                item_fg = text_color if enabled else muted_color

                # Icon if provided
                if icon:
                    icon_label = tk.Label(
                        item_frame,
                        image=icon,
                        bg=popup_bg
                    )
                    icon_label.image = icon  # Keep reference
                    icon_label.pack(side=tk.LEFT, padx=(12, 4))
                    widgets_to_bind = [item_frame, icon_label]
                else:
                    widgets_to_bind = [item_frame]

                # Text label
                label = tk.Label(
                    item_frame,
                    text=label_text,
                    font=('Segoe UI', 9),
                    fg=item_fg,
                    bg=popup_bg,
                    anchor='w',
                    padx=12 if not icon else 0,
                    pady=6,
                    cursor='hand2' if enabled else ''
                )
                label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12) if icon else 0)
                widgets_to_bind.append(label)

                # Only add hover/click for enabled items
                if enabled:
                    # Hover effects - capture variables properly
                    def make_on_enter(frame, lbl, icon_lbl, hbg):
                        def on_enter(e):
                            frame.configure(bg=hbg)
                            lbl.configure(bg=hbg)
                            if icon_lbl:
                                icon_lbl.configure(bg=hbg)
                        return on_enter

                    def make_on_leave(frame, lbl, icon_lbl, bg):
                        def on_leave(e):
                            frame.configure(bg=bg)
                            lbl.configure(bg=bg)
                            if icon_lbl:
                                icon_lbl.configure(bg=bg)
                        return on_leave

                    def make_on_click(cmd):
                        def on_click(e):
                            self.close()
                            cmd()
                        return on_click

                    icon_label_ref = widgets_to_bind[1] if icon else None
                    on_enter = make_on_enter(item_frame, label, icon_label_ref, hover_bg)
                    on_leave = make_on_leave(item_frame, label, icon_label_ref, popup_bg)
                    on_click = make_on_click(command)

                    for widget in widgets_to_bind:
                        widget.bind('<Enter>', on_enter)
                        widget.bind('<Leave>', on_leave)
                        widget.bind('<Button-1>', on_click)

        # Position popup
        self._popup.update_idletasks()
        self._popup.geometry(f"+{x}+{y}")
        self._popup.deiconify()
        # Lift relative to parent window only (not above other apps)
        self._popup.lift(self._parent.winfo_toplevel())
        self._popup.focus_set()

        # Bind events to close popup
        self._popup.bind('<Escape>', lambda e: self.close())
        # Track bind ID for proper cleanup
        self._outside_click_bind_id = self._parent.winfo_toplevel().bind(
            '<Button-1>', self._on_outside_click, add='+'
        )

    def _on_outside_click(self, event):
        """Handle click outside context popup."""
        if not self._popup or not self._popup.winfo_exists():
            return

        # Check if click is outside popup
        px = self._popup.winfo_rootx()
        py = self._popup.winfo_rooty()
        pw = self._popup.winfo_width()
        ph = self._popup.winfo_height()

        if not (px <= event.x_root <= px + pw and py <= event.y_root <= py + ph):
            self.close()

    def close(self):
        """Close the context menu popup and cleanup bindings."""
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        # Properly unbind using tracked bind ID
        if self._outside_click_bind_id:
            try:
                self._parent.winfo_toplevel().unbind('<Button-1>', self._outside_click_bind_id)
            except Exception:
                pass
            self._outside_click_bind_id = None

    def destroy(self):
        """Full cleanup - close popup and clear items."""
        self.close()
        self._items.clear()
