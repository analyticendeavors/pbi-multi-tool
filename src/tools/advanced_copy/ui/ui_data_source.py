"""
Data Source Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Handles data source setup, file input, and guide text.
"""

import tkinter as tk
from tkinter import ttk, filedialog
from pathlib import Path

from core.constants import AppConstants
from core.ui_base import SquareIconButton


class DataSourceMixin:
    """
    Mixin for data source and file input functionality.

    Methods extracted from AdvancedCopyTab:
    - _setup_data_source()
    - _update_guide_text()
    - _show_target_pbip_input()
    - _hide_target_pbip_input()
    - _adjust_window_height_for_target_pbip()
    - browse_target_file()
    """

    def _setup_data_source(self):
        """Setup data source section"""
        colors = self._theme_manager.colors

        # Create section header with SVG icon using labelwidget
        header_widget = self.create_section_header(self.frame, "PBIP File Source", "Power-BI")[0]
        section_frame = ttk.LabelFrame(self.frame, labelwidget=header_widget,
                                       style='Section.TLabelframe', padding="12")
        section_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        self._data_source_section = section_frame

        # Content frame FIRST (for proper stacking order with help button)
        content_frame = ttk.Frame(section_frame, style='Section.TFrame', padding="15")
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(1, weight=1)

        # Help button AFTER content_frame to ensure proper stacking order
        help_icon = self._button_icons.get('question')
        if help_icon:
            self._help_button = SquareIconButton(
                section_frame, icon=help_icon, command=self.show_help_dialog,
                tooltip_text="Help", size=26, radius=6,
                bg_normal_override=AppConstants.CORNER_ICON_BG
            )
            # Position in upper-right corner of section title bar area
            self._help_button.place(relx=1.0, y=-35, anchor=tk.NE, x=0)

        # Import RoundedButton for browse button
        from core.ui_base import RoundedButton

        # Report file input row - matches Report Cleanup layout
        ttk.Label(content_frame, text="Project File (PBIP):",
                  style='Section.TLabel', font=('Segoe UI', 10)).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))

        entry = ttk.Entry(content_frame, textvariable=self.report_path,
                          font=('Segoe UI', 10), style='Section.TEntry')
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        # Browse button with folder icon - use RoundedButton directly like Report Cleanup
        # canvas_bg for corner rounding: #0d0d1a dark / #ffffff light (main background)
        is_dark = self._theme_manager.is_dark
        file_section_canvas_bg = '#0d0d1a' if is_dark else '#ffffff'
        folder_icon = self._button_icons.get('folder')
        browse_btn = RoundedButton(
            content_frame, text="Browse" if folder_icon else "Browse",
            command=self.browse_file,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            width=90, height=32, radius=6, font=('Segoe UI', 10),
            icon=folder_icon,
            canvas_bg=file_section_canvas_bg
        )
        browse_btn.grid(row=0, column=2)
        self._primary_buttons.append(browse_btn)

        # Track widgets that need theme updates
        self._options_widgets = []

        # Container for side-by-side Copy Content and Copy Destination sections
        options_container = ttk.Frame(content_frame, style='Section.TFrame')
        options_container.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        options_container.columnconfigure(0, weight=1)
        options_container.columnconfigure(1, weight=1)

        # Create Copy Content section (LEFT) with icon label header
        self._setup_copy_content_section(options_container, colors)

        # Create Copy Destination section (RIGHT) with icon label header
        self._setup_copy_destination_section(options_container, colors)

        # Analyze button BELOW the side-by-side sections
        magnify_icon = self._button_icons.get('magnifying-glass')
        bg_disabled = colors.get('button_primary_disabled', '#3a3a4e')
        fg_disabled = colors.get('button_text_disabled', '#6a6a7a')

        from core.ui_base import RoundedButton
        self.analyze_button = RoundedButton(
            content_frame, text="ANALYZE REPORT",
            command=self.analyze_report,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            disabled_bg=bg_disabled, disabled_fg=fg_disabled,
            height=38, radius=6, font=('Segoe UI', 10, 'bold'),
            icon=magnify_icon,
            canvas_bg=file_section_canvas_bg
        )
        self.analyze_button.grid(row=2, column=0, columnspan=3, pady=(15, 0))
        # Start disabled until file is selected
        self.analyze_button.set_enabled(False)
        self._analyze_button_enabled = False

        # Target PBIP input (for cross-PBIP mode) - will be shown/hidden at row 3
        self._setup_target_pbip_section(content_frame, colors)

        # Setup path cleaning
        self.setup_path_cleaning(self.report_path)
        self.setup_path_cleaning(self.target_pbip_path)

    def _setup_copy_content_section(self, parent, colors):
        """Setup Copy Content selection section with SVG icons (LEFT side)"""
        # Use a light background for better radio visibility
        option_bg = colors.get('option_bg', '#f5f5f7')

        # Content frame with border - LEFT column
        self._content_frame = tk.Frame(parent, bg=option_bg,
                                       highlightbackground=colors['border'],
                                       highlightthickness=1)
        self._content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))

        # Header row inside the frame
        file_icon = self._button_icons.get('file')
        header_frame = tk.Frame(self._content_frame, bg=option_bg)
        header_frame.pack(fill=tk.X, padx=10, pady=(8, 0))

        if file_icon:
            icon_lbl = tk.Label(header_frame, image=file_icon, bg=option_bg)
            icon_lbl.pack(side=tk.LEFT, padx=(0, 6))
            icon_lbl._icon_ref = file_icon
        else:
            icon_lbl = None
        title_lbl = tk.Label(header_frame, text="Copy Content", bg=option_bg,
                            fg=colors['text_primary'], font=('Segoe UI', 10, 'bold'))
        title_lbl.pack(side=tk.LEFT)
        self._content_header_frame = header_frame
        self._content_header_icon = icon_lbl
        self._content_header_title = title_lbl

        # Full Page Copy mode row - radio, text label, and description on same line
        self._full_page_row = tk.Frame(self._content_frame, bg=option_bg)
        self._full_page_row.pack(anchor=tk.W, pady=(6, 4), padx=20, fill=tk.X)

        # Radio button with SVG icon
        self._full_page_radio = self._create_svg_radio(
            self._full_page_row, self.copy_content_mode, "full_page",
            self._on_content_mode_change, colors, option_bg
        )
        self._full_page_radio.pack(side=tk.LEFT)

        # Determine initial color based on selection
        is_selected = self.copy_content_mode.get() == "full_page"
        text_fg = colors['title_color'] if is_selected else colors['text_primary']
        self._full_page_text = tk.Label(self._full_page_row, text="Full Page Copy",
                font=('Segoe UI', 9), fg=text_fg, bg=option_bg,
                cursor='hand2', width=18, anchor=tk.W)
        self._full_page_text.pack(side=tk.LEFT, padx=(4, 0))
        self._full_page_text.bind('<Button-1>', lambda e: self._radio_click('full_page', 'content'))
        # Add hover underline effect
        self._full_page_text.bind('<Enter>', lambda e: self._full_page_text.configure(font=('Segoe UI', 9, 'underline')))
        self._full_page_text.bind('<Leave>', lambda e: self._full_page_text.configure(font=('Segoe UI', 9)))

        # Description on same line (to the right)
        self._full_page_label = tk.Label(self._full_page_row,
                text="- Copies entire pages with bookmarks",
                font=('Segoe UI', 8), fg=colors['info'], bg=option_bg)
        self._full_page_label.pack(side=tk.LEFT, padx=(8, 0))

        # Bookmark + Visuals Copy mode row - radio, text label, and description on same line
        self._bookmark_row = tk.Frame(self._content_frame, bg=option_bg)
        self._bookmark_row.pack(anchor=tk.W, pady=(0, 8), padx=20, fill=tk.X)

        self._bookmark_radio = self._create_svg_radio(
            self._bookmark_row, self.copy_content_mode, "bookmark_visual",
            self._on_content_mode_change, colors, option_bg
        )
        self._bookmark_radio.pack(side=tk.LEFT)

        # Determine initial color based on selection
        is_selected = self.copy_content_mode.get() == "bookmark_visual"
        text_fg = colors['title_color'] if is_selected else colors['text_primary']
        self._bookmark_text = tk.Label(self._bookmark_row, text="Bookmark + Visuals",
                font=('Segoe UI', 9), fg=text_fg, bg=option_bg,
                cursor='hand2', width=18, anchor=tk.W)
        self._bookmark_text.pack(side=tk.LEFT, padx=(4, 0))
        self._bookmark_text.bind('<Button-1>', lambda e: self._radio_click('bookmark_visual', 'content'))
        # Add hover underline effect
        self._bookmark_text.bind('<Enter>', lambda e: self._bookmark_text.configure(font=('Segoe UI', 9, 'underline')))
        self._bookmark_text.bind('<Leave>', lambda e: self._bookmark_text.configure(font=('Segoe UI', 9)))

        # Description on same line (to the right)
        self._bookmark_label = tk.Label(self._bookmark_row,
                text="- Copies specific bookmarks + visuals",
                font=('Segoe UI', 8), fg=colors['info'], bg=option_bg)
        self._bookmark_label.pack(side=tk.LEFT, padx=(8, 0))

    def _setup_copy_destination_section(self, parent, colors):
        """Setup Copy Destination selection section with SVG icons (RIGHT side)"""
        # Use a light background for better radio visibility
        option_bg = colors.get('option_bg', '#f5f5f7')

        # Destination frame with border - RIGHT column
        self._destination_frame = tk.Frame(parent, bg=option_bg,
                                           highlightbackground=colors['border'],
                                           highlightthickness=1)
        self._destination_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))

        # Header row inside the frame
        target_icon = self._button_icons.get('target')
        header_frame = tk.Frame(self._destination_frame, bg=option_bg)
        header_frame.pack(fill=tk.X, padx=10, pady=(8, 0))

        if target_icon:
            icon_lbl = tk.Label(header_frame, image=target_icon, bg=option_bg)
            icon_lbl.pack(side=tk.LEFT, padx=(0, 6))
            icon_lbl._icon_ref = target_icon
        else:
            icon_lbl = None
        title_lbl = tk.Label(header_frame, text="Copy Destination", bg=option_bg,
                            fg=colors['text_primary'], font=('Segoe UI', 10, 'bold'))
        title_lbl.pack(side=tk.LEFT)
        self._destination_header_frame = header_frame
        self._destination_header_icon = icon_lbl
        self._destination_header_title = title_lbl

        # Same PBIP mode row - radio, text label, and description on same line
        self._same_pbip_row = tk.Frame(self._destination_frame, bg=option_bg)
        self._same_pbip_row.pack(anchor=tk.W, pady=(6, 4), padx=20, fill=tk.X)

        self._same_pbip_radio = self._create_svg_radio(
            self._same_pbip_row, self.copy_destination_mode, "same_pbip",
            self._on_destination_mode_change, colors, option_bg
        )
        self._same_pbip_radio.pack(side=tk.LEFT)

        # Determine initial color based on selection
        is_selected = self.copy_destination_mode.get() == "same_pbip"
        text_fg = colors['title_color'] if is_selected else colors['text_primary']
        self._same_pbip_text = tk.Label(self._same_pbip_row, text="Same PBIP",
                font=('Segoe UI', 9), fg=text_fg, bg=option_bg,
                cursor='hand2', width=11, anchor=tk.W)
        self._same_pbip_text.pack(side=tk.LEFT, padx=(4, 0))
        self._same_pbip_text.bind('<Button-1>', lambda e: self._radio_click('same_pbip', 'destination'))
        # Add hover underline effect
        self._same_pbip_text.bind('<Enter>', lambda e: self._same_pbip_text.configure(font=('Segoe UI', 9, 'underline')))
        self._same_pbip_text.bind('<Leave>', lambda e: self._same_pbip_text.configure(font=('Segoe UI', 9)))

        # Description on same line (to the right)
        self._same_pbip_label = tk.Label(self._same_pbip_row,
                text="- Copy within the same report",
                font=('Segoe UI', 8), fg=colors['info'], bg=option_bg)
        self._same_pbip_label.pack(side=tk.LEFT, padx=(8, 0))

        # Cross-PBIP mode row - radio, text label, and description on same line
        self._cross_pbip_row = tk.Frame(self._destination_frame, bg=option_bg)
        self._cross_pbip_row.pack(anchor=tk.W, pady=(0, 8), padx=20, fill=tk.X)

        self._cross_pbip_radio = self._create_svg_radio(
            self._cross_pbip_row, self.copy_destination_mode, "cross_pbip",
            self._on_destination_mode_change, colors, option_bg
        )
        self._cross_pbip_radio.pack(side=tk.LEFT)

        # Determine initial color based on selection
        is_selected = self.copy_destination_mode.get() == "cross_pbip"
        text_fg = colors['title_color'] if is_selected else colors['text_primary']
        self._cross_pbip_text = tk.Label(self._cross_pbip_row, text="Cross-PBIP",
                font=('Segoe UI', 9), fg=text_fg, bg=option_bg,
                cursor='hand2', width=11, anchor=tk.W)
        self._cross_pbip_text.pack(side=tk.LEFT, padx=(4, 0))
        self._cross_pbip_text.bind('<Button-1>', lambda e: self._radio_click('cross_pbip', 'destination'))
        # Add hover underline effect
        self._cross_pbip_text.bind('<Enter>', lambda e: self._cross_pbip_text.configure(font=('Segoe UI', 9, 'underline')))
        self._cross_pbip_text.bind('<Leave>', lambda e: self._cross_pbip_text.configure(font=('Segoe UI', 9)))

        # Description on same line (to the right)
        self._cross_pbip_label = tk.Label(self._cross_pbip_row,
                text="- Copy to different target report",
                font=('Segoe UI', 8), fg=colors['info'], bg=option_bg)
        self._cross_pbip_label.pack(side=tk.LEFT, padx=(8, 0))

    def _setup_target_pbip_section(self, parent, colors):
        """Setup Target PBIP input section (hidden by default)"""
        # Use option_bg to match Copy Content/Destination sections
        option_bg = colors.get('option_bg', '#f5f5f7')

        # Target PBIP frame with border - matches Copy Content/Destination style
        self.target_pbip_frame = tk.Frame(parent, bg=option_bg,
                                          highlightbackground=colors['border'],
                                          highlightthickness=1)
        # Will be shown/hidden based on destination mode - positioned at row 3 (before analyze button)

        # Store references for header (will be created inside frame when shown)
        self._target_header_frame = None
        self._target_header_icon = None
        self._target_header_title = None

    def _create_svg_radio(self, parent, variable, value, command, colors, bg_color=None):
        """Create a custom radio button using SVG icons"""
        radio_on = self._button_icons.get('radio-on')
        radio_off = self._button_icons.get('radio-off')

        # Use provided bg_color or fall back to card_surface
        canvas_bg = bg_color if bg_color else colors['card_surface']

        # Create canvas for radio button
        canvas = tk.Canvas(parent, width=18, height=18, bg=canvas_bg,
                          highlightthickness=0, cursor='hand2')

        # Store references
        canvas._variable = variable
        canvas._value = value
        canvas._command = command
        canvas._icon_on = radio_on
        canvas._icon_off = radio_off

        # Draw initial state
        self._update_svg_radio(canvas, colors)

        # Bind click event
        canvas.bind('<Button-1>', lambda e: self._on_svg_radio_click(canvas))

        # Trace variable changes
        variable.trace_add('write', lambda *args: self._update_svg_radio(canvas, self._theme_manager.colors))

        return canvas

    def _update_svg_radio(self, canvas, colors):
        """Update SVG radio button display"""
        canvas.delete('all')
        is_selected = canvas._variable.get() == canvas._value
        icon = canvas._icon_on if is_selected else canvas._icon_off

        if icon:
            canvas.create_image(9, 9, image=icon, anchor=tk.CENTER)
        else:
            # Fallback: draw circle - use canvas bg color
            canvas_bg = canvas.cget('bg')
            if is_selected:
                canvas.create_oval(2, 2, 16, 16, outline=colors['accent'], fill=colors['accent'], width=2)
                canvas.create_oval(6, 6, 12, 12, fill='white', outline='white')
            else:
                canvas.create_oval(2, 2, 16, 16, outline=colors['border'], fill=canvas_bg, width=2)

    def _on_svg_radio_click(self, canvas):
        """Handle SVG radio button click"""
        canvas._variable.set(canvas._value)
        # Update text colors based on which group this belongs to
        if canvas._variable == self.copy_content_mode:
            self._update_radio_text_colors('content')
        elif canvas._variable == self.copy_destination_mode:
            self._update_radio_text_colors('destination')
        if canvas._command:
            canvas._command()

    def _radio_click(self, value, mode_type):
        """Handle click on radio button text label"""
        if mode_type == 'content':
            self.copy_content_mode.set(value)
            self._update_radio_text_colors('content')
            self._on_content_mode_change()
        else:
            self.copy_destination_mode.set(value)
            self._update_radio_text_colors('destination')
            self._on_destination_mode_change()

    def _update_radio_text_colors(self, mode_type):
        """Update radio button text colors based on selection state.
        Selected items use title_color, unselected use text_primary.
        """
        colors = self._theme_manager.colors
        option_bg = colors.get('option_bg', '#f5f5f7')

        if mode_type == 'content':
            # Update Copy Content radio text colors
            is_full_page = self.copy_content_mode.get() == "full_page"
            if hasattr(self, '_full_page_text') and self._full_page_text:
                fg = colors['title_color'] if is_full_page else colors['text_primary']
                self._full_page_text.config(fg=fg, bg=option_bg)
            if hasattr(self, '_bookmark_text') and self._bookmark_text:
                fg = colors['title_color'] if not is_full_page else colors['text_primary']
                self._bookmark_text.config(fg=fg, bg=option_bg)
        else:
            # Update Copy Destination radio text colors
            is_same_pbip = self.copy_destination_mode.get() == "same_pbip"
            if hasattr(self, '_same_pbip_text') and self._same_pbip_text:
                fg = colors['title_color'] if is_same_pbip else colors['text_primary']
                self._same_pbip_text.config(fg=fg, bg=option_bg)
            if hasattr(self, '_cross_pbip_text') and self._cross_pbip_text:
                fg = colors['title_color'] if not is_same_pbip else colors['text_primary']
                self._cross_pbip_text.config(fg=fg, bg=option_bg)

    def _update_guide_text(self):
        """Update guide text based on selected copy modes - simplified version"""
        # Guide text is now static - just a single tip line
        # This method is kept for compatibility but doesn't need to do anything
        pass

    def _show_target_pbip_input(self):
        """Show target PBIP file input for cross-PBIP destination mode"""
        if not self.target_pbip_frame:
            return

        # Get dynamic theme colors
        colors = self._theme_manager.colors
        option_bg = colors.get('option_bg', '#f5f5f7')

        # Clear existing content
        for widget in self.target_pbip_frame.winfo_children():
            widget.destroy()

        # Configure frame background
        self.target_pbip_frame.configure(bg=option_bg, highlightbackground=colors['border'])

        # Row 1: Header row (icon + title on left, description on right)
        header_row = tk.Frame(self.target_pbip_frame, bg=option_bg)
        header_row.pack(fill=tk.X, padx=10, pady=(8, 6))

        # Icon and title on left
        target_icon = self._button_icons.get('target')
        if target_icon:
            icon_lbl = tk.Label(header_row, image=target_icon, bg=option_bg)
            icon_lbl.pack(side=tk.LEFT, padx=(0, 6))
            icon_lbl._icon_ref = target_icon
            self._target_header_icon = icon_lbl
        else:
            self._target_header_icon = None

        title_lbl = tk.Label(header_row, text="Target PBIP File", bg=option_bg,
                            fg=colors['text_primary'], font=('Segoe UI', 10, 'bold'))
        title_lbl.pack(side=tk.LEFT)
        self._target_header_title = title_lbl
        self._target_header_frame = header_row

        # Description on right of same row
        self._target_desc_label = tk.Label(header_row,
                text="Select the target PBIP file where content will be copied:",
                bg=option_bg, fg=colors['text_secondary'], font=('Segoe UI', 9))
        self._target_desc_label.pack(side=tk.RIGHT)

        # Row 2: Input row (label + entry + browse button)
        input_row = tk.Frame(self.target_pbip_frame, bg=option_bg)
        input_row.pack(fill=tk.X, padx=10, pady=(0, 8))
        input_row.columnconfigure(1, weight=1)

        self._target_file_label = tk.Label(input_row, text="Target Project File (PBIP):",
                bg=option_bg, fg=colors['text_primary'], font=('Segoe UI', 10))
        self._target_file_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        entry = ttk.Entry(input_row, textvariable=self.target_pbip_path,
                          font=('Segoe UI', 10), style='Section.TEntry')
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self._target_input_row = input_row  # Store for theme updates

        # Import RoundedButton for consistent browse button styling
        from core.ui_base import RoundedButton

        # Browse button with folder icon - consistent styling
        folder_icon = self._button_icons.get('folder')
        browse_target_btn = RoundedButton(
            input_row, text="Browse" if folder_icon else "Browse",
            command=self.browse_target_file,
            bg=colors['button_primary'], hover_bg=colors['button_primary_hover'],
            pressed_bg=colors['button_primary_pressed'], fg='#ffffff',
            width=90, height=32, radius=6, font=('Segoe UI', 10),
            icon=folder_icon,
            canvas_bg=option_bg
        )
        browse_target_btn.grid(row=0, column=2)
        self._primary_buttons.append(browse_target_btn)

        # Show the frame at row 2 (between options and analyze button)
        self.target_pbip_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

        # Move analyze button to row 3
        if hasattr(self, 'analyze_button') and self.analyze_button:
            self.analyze_button.grid(row=3, column=0, columnspan=3, pady=(15, 0))

        # Setup path cleaning for target
        self.target_pbip_path.trace_add('write', lambda *args: self._on_path_change())

        # Adjust window height to accommodate target PBIP input
        self._adjust_window_height_for_target_pbip(True)

        self.log_message("Cross-PBIP destination selected - select target PBIP file")

    def _hide_target_pbip_input(self):
        """Hide target PBIP file input"""
        if self.target_pbip_frame:
            self.target_pbip_frame.grid_remove()
            self.target_pbip_path.set("")  # Clear target path

            # Move analyze button back to row 2
            if hasattr(self, 'analyze_button') and self.analyze_button:
                self.analyze_button.grid(row=2, column=0, columnspan=3, pady=(15, 0))

            # Adjust window height back down
            self._adjust_window_height_for_target_pbip(False)

    def _adjust_window_height_for_target_pbip(self, show_target_input: bool):
        """Adjust window height when target PBIP input is shown/hidden

        Only adjusts height when the actual visibility changes, not on repeated clicks.
        Uses a flag to track current visibility state.
        """
        # Initialize the flag if it doesn't exist
        if not hasattr(self, '_target_pbip_was_visible'):
            self._target_pbip_was_visible = False

        # Only adjust if visibility actually changed
        if show_target_input == self._target_pbip_was_visible:
            return  # No change, don't adjust

        # Update the visibility flag
        self._target_pbip_was_visible = show_target_input

        try:
            if hasattr(self, 'main_app') and hasattr(self.main_app, 'root'):
                root = self.main_app.root

                # Height adjustment for target PBIP section
                target_pbip_height = 90  # Height of the target PBIP input section

                # Get current window geometry
                current_geometry = root.geometry()
                parts = current_geometry.split('x')
                if len(parts) >= 2:
                    width = parts[0]
                    height_and_pos = parts[1].split('+', 1)
                    current_height = int(height_and_pos[0])
                    position = height_and_pos[1] if len(height_and_pos) > 1 else ''

                    # Calculate new height
                    if show_target_input:
                        new_height = current_height + target_pbip_height
                    else:
                        new_height = current_height - target_pbip_height

                    # Apply new geometry
                    new_geometry = f"{width}x{new_height}+{position}"
                    root.geometry(new_geometry)
        except Exception:
            pass

    def browse_file(self):
        """Browse for report file"""
        file_path = filedialog.askopenfilename(
            title="Select Report (.pbip file - PBIP format required)",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.report_path.set(file_path)

    def browse_target_file(self):
        """Browse for target PBIP file"""
        file_path = filedialog.askopenfilename(
            title="Select Target PBIP File",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.target_pbip_path.set(file_path)

    def _update_options_theme(self):
        """Update options panel colors when theme changes"""
        colors = self._theme_manager.colors
        option_bg = colors.get('option_bg', '#f5f5f7')


        # Update analyze button colors including canvas_bg for corner rounding
        is_dark = self._theme_manager.is_dark
        bg_disabled = colors.get('button_primary_disabled', '#3a3a4e' if is_dark else '#c0c0cc')
        fg_disabled = colors.get('button_text_disabled', '#6a6a7a' if is_dark else '#9a9aa8')
        file_section_canvas_bg = '#0d0d1a' if is_dark else '#ffffff'
        if hasattr(self, 'analyze_button') and self.analyze_button:
            try:
                self.analyze_button.update_colors(
                    bg=colors['button_primary'],
                    hover_bg=colors['button_primary_hover'],
                    pressed_bg=colors['button_primary_pressed'],
                    fg='#ffffff',
                    disabled_bg=bg_disabled,
                    disabled_fg=fg_disabled,
                    canvas_bg=file_section_canvas_bg
                )
            except Exception:
                pass

        # Update Copy Content section header
        if hasattr(self, '_content_header_frame') and self._content_header_frame:
            self._content_header_frame.config(bg=option_bg)
        if hasattr(self, '_content_header_icon') and self._content_header_icon:
            self._content_header_icon.config(bg=option_bg)
        if hasattr(self, '_content_header_title') and self._content_header_title:
            self._content_header_title.config(bg=option_bg, fg=colors['text_primary'])

        # Update Copy Content frame and children
        if hasattr(self, '_content_frame') and self._content_frame:
            self._content_frame.config(bg=option_bg, highlightbackground=colors['border'])
        if hasattr(self, '_full_page_row') and self._full_page_row:
            self._full_page_row.config(bg=option_bg)
        if hasattr(self, '_full_page_radio') and self._full_page_radio:
            self._full_page_radio.config(bg=option_bg)
            self._update_svg_radio(self._full_page_radio, colors)
        # Update text colors based on selection state (selected = title_color)
        is_full_page = self.copy_content_mode.get() == "full_page"
        if hasattr(self, '_full_page_text') and self._full_page_text:
            fg = colors['title_color'] if is_full_page else colors['text_primary']
            self._full_page_text.config(fg=fg, bg=option_bg)
        if hasattr(self, '_full_page_label') and self._full_page_label:
            self._full_page_label.config(fg=colors['info'], bg=option_bg)
        if hasattr(self, '_bookmark_row') and self._bookmark_row:
            self._bookmark_row.config(bg=option_bg)
        if hasattr(self, '_bookmark_radio') and self._bookmark_radio:
            self._bookmark_radio.config(bg=option_bg)
            self._update_svg_radio(self._bookmark_radio, colors)
        if hasattr(self, '_bookmark_text') and self._bookmark_text:
            fg = colors['title_color'] if not is_full_page else colors['text_primary']
            self._bookmark_text.config(fg=fg, bg=option_bg)
        if hasattr(self, '_bookmark_label') and self._bookmark_label:
            self._bookmark_label.config(fg=colors['info'], bg=option_bg)

        # Update Destination section header
        if hasattr(self, '_destination_header_frame') and self._destination_header_frame:
            self._destination_header_frame.config(bg=option_bg)
        if hasattr(self, '_destination_header_icon') and self._destination_header_icon:
            self._destination_header_icon.config(bg=option_bg)
        if hasattr(self, '_destination_header_title') and self._destination_header_title:
            self._destination_header_title.config(bg=option_bg, fg=colors['text_primary'])

        # Update Destination frame and children
        if hasattr(self, '_destination_frame') and self._destination_frame:
            self._destination_frame.config(bg=option_bg, highlightbackground=colors['border'])
        if hasattr(self, '_same_pbip_row') and self._same_pbip_row:
            self._same_pbip_row.config(bg=option_bg)
        if hasattr(self, '_same_pbip_radio') and self._same_pbip_radio:
            self._same_pbip_radio.config(bg=option_bg)
            self._update_svg_radio(self._same_pbip_radio, colors)
        # Update text colors based on selection state (selected = title_color)
        is_same_pbip = self.copy_destination_mode.get() == "same_pbip"
        if hasattr(self, '_same_pbip_text') and self._same_pbip_text:
            fg = colors['title_color'] if is_same_pbip else colors['text_primary']
            self._same_pbip_text.config(fg=fg, bg=option_bg)
        if hasattr(self, '_same_pbip_label') and self._same_pbip_label:
            self._same_pbip_label.config(fg=colors['info'], bg=option_bg)
        if hasattr(self, '_cross_pbip_row') and self._cross_pbip_row:
            self._cross_pbip_row.config(bg=option_bg)
        if hasattr(self, '_cross_pbip_radio') and self._cross_pbip_radio:
            self._cross_pbip_radio.config(bg=option_bg)
            self._update_svg_radio(self._cross_pbip_radio, colors)
        if hasattr(self, '_cross_pbip_text') and self._cross_pbip_text:
            fg = colors['title_color'] if not is_same_pbip else colors['text_primary']
            self._cross_pbip_text.config(fg=fg, bg=option_bg)
        if hasattr(self, '_cross_pbip_label') and self._cross_pbip_label:
            self._cross_pbip_label.config(fg=colors['info'], bg=option_bg)

        # Update Target PBIP section (uses option_bg to match other sections)
        if hasattr(self, 'target_pbip_frame') and self.target_pbip_frame:
            self.target_pbip_frame.config(bg=option_bg, highlightbackground=colors['border'])
        if hasattr(self, '_target_header_row') and self._target_header_row:
            self._target_header_row.config(bg=option_bg)
        if hasattr(self, '_target_title_frame') and self._target_title_frame:
            self._target_title_frame.config(bg=option_bg)
        if hasattr(self, '_target_icon_label') and self._target_icon_label:
            self._target_icon_label.config(bg=option_bg)
        if hasattr(self, '_target_title_label') and self._target_title_label:
            self._target_title_label.config(bg=option_bg, fg=colors['text_primary'])
        if hasattr(self, '_target_desc_label') and self._target_desc_label:
            self._target_desc_label.config(fg=colors['text_secondary'], bg=option_bg)
        if hasattr(self, '_target_input_row') and self._target_input_row:
            self._target_input_row.config(bg=option_bg)
        if hasattr(self, '_target_file_label') and self._target_file_label:
            self._target_file_label.config(fg=colors['text_primary'], bg=option_bg)
