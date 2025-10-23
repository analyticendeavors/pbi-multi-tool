"""
Data Source Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Handles data source setup, file input, and guide text.
"""

import tkinter as tk
from tkinter import ttk, filedialog
from pathlib import Path

from core.constants import AppConstants


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
        # Guide text for the file input section
        guide_text = [
            "🚀 QUICK START GUIDE:",
            "1. Select your .pbip report file below",
            "2. Click 'Analyze Report' to scan for pages with bookmarks", 
            "3. Select which pages to copy from the list",
            "4. Pages will be duplicated with all bookmarks preserved",
            "5. Click 'Execute Copy' to create the duplicates",
            "6. New pages will have '(Copy)' suffix to avoid conflicts",
            "⚠️ Requires PBIP format with TMDLs files"
        ]
        
        # Create file input section using base class
        file_input = self.create_file_input_section(
            self.frame,
            "📁 PBIP FILE SOURCE",
            [("Power BI Project Files", "*.pbip"), ("All Files", "*.*")],
            []  # Pass empty guide_text to create section without guide
        )
        file_input['frame'].grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Get the content frame and manually add our custom guide text
        content_frame = file_input['frame'].winfo_children()[0]  # Get the content frame
        
        # Create custom guide frame with proper alignment and matching background color
        self.guide_frame = tk.Frame(content_frame, bg='#f8fafc')  # Match surrounding background
        self.guide_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 35))
        
        # Store reference to guide labels for dynamic updates
        self.guide_labels = []
        
        # Create Analyze button once (will be positioned by _update_guide_text)
        self.analyze_button = ttk.Button(self.guide_frame, text="🔍 ANALYZE REPORT",
                                       command=self.analyze_report, 
                                       style='Action.TButton', state=tk.DISABLED)
        
        # Initial guide text (will be updated based on mode)
        self._update_guide_text()
        
        # Modify the input frame to have a single report input
        input_frame = file_input['input_frame']
        
        # Clear the default input and create custom one
        for widget in input_frame.winfo_children():
            widget.destroy()
        
        input_frame.columnconfigure(1, weight=1)
        
        # Report file input
        ttk.Label(input_frame, text="Project File (PBIP):").grid(row=0, column=0, sticky=tk.W, pady=8)
        entry = ttk.Entry(input_frame, textvariable=self.report_path, width=80)
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        ttk.Button(input_frame, text="📂 Browse", 
                  command=self.browse_file).grid(row=0, column=2, pady=8)
        
        # Copy Content Selection
        content_frame = tk.LabelFrame(input_frame, text="📋 Copy Content", 
                                     font=('Segoe UI', 10, 'bold'),
                                     fg=AppConstants.COLORS['text_primary'],
                                     bg='#dcdad5',
                                     bd=1, relief='solid', padx=8, pady=6)
        content_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Full Page Copy mode
        full_page_row = tk.Frame(content_frame, bg='#dcdad5')
        full_page_row.pack(anchor=tk.W, pady=3, fill=tk.X)
        
        ttk.Radiobutton(full_page_row, text="📄 Full Page Copy", 
                       variable=self.copy_content_mode, value="full_page",
                       command=self._on_content_mode_change).pack(side=tk.LEFT)
        
        tk.Label(full_page_row, 
                text="Copies entire pages with all visuals and bookmarks",
                font=('Segoe UI', 9),
                fg=AppConstants.COLORS['info'],
                bg='#dcdad5').pack(side=tk.LEFT, padx=(10, 0))
        
        # Bookmark + Visuals Copy mode
        bookmark_row = tk.Frame(content_frame, bg='#dcdad5')
        bookmark_row.pack(anchor=tk.W, pady=(6, 3), fill=tk.X)
        
        ttk.Radiobutton(bookmark_row, text="🔖 Bookmark + Visuals Copy", 
                       variable=self.copy_content_mode, value="bookmark_visual",
                       command=self._on_content_mode_change).pack(side=tk.LEFT)
        
        tk.Label(bookmark_row, 
                text="Copies specific bookmarks and their visuals to target pages",
                font=('Segoe UI', 9),
                fg=AppConstants.COLORS['info'],
                bg='#dcdad5').pack(side=tk.LEFT, padx=(10, 0))
        
        # Copy Destination Selection
        destination_frame = tk.LabelFrame(input_frame, text="🎯 Copy Destination", 
                                         font=('Segoe UI', 10, 'bold'),
                                         fg=AppConstants.COLORS['text_primary'],
                                         bg='#dcdad5',
                                         bd=1, relief='solid', padx=8, pady=6)
        destination_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))
        
        # Same PBIP mode
        same_pbip_row = tk.Frame(destination_frame, bg='#dcdad5')
        same_pbip_row.pack(anchor=tk.W, pady=3, fill=tk.X)
        
        ttk.Radiobutton(same_pbip_row, text="📄 Same PBIP (within report)", 
                       variable=self.copy_destination_mode, value="same_pbip",
                       command=self._on_destination_mode_change).pack(side=tk.LEFT)
        
        tk.Label(same_pbip_row, 
                text="Copy within the same report file",
                font=('Segoe UI', 9),
                fg=AppConstants.COLORS['info'],
                bg='#dcdad5').pack(side=tk.LEFT, padx=(10, 0))
        
        # Cross-PBIP mode
        cross_pbip_row = tk.Frame(destination_frame, bg='#dcdad5')
        cross_pbip_row.pack(anchor=tk.W, pady=(6, 3), fill=tk.X)
        
        ttk.Radiobutton(cross_pbip_row, text="🔄 Cross-PBIP (between reports)", 
                       variable=self.copy_destination_mode, value="cross_pbip",
                       command=self._on_destination_mode_change).pack(side=tk.LEFT)
        
        tk.Label(cross_pbip_row, 
                text="Copy from source report to a different target report",
                font=('Segoe UI', 9),
                fg=AppConstants.COLORS['info'],
                bg='#dcdad5').pack(side=tk.LEFT, padx=(10, 0))
        
        # Target PBIP input (for cross-PBIP mode) - positioned after destination selection
        self.target_pbip_frame = tk.LabelFrame(input_frame, text="🎯 Target PBIP File",
                                              font=('Segoe UI', 10, 'bold'),
                                              fg=AppConstants.COLORS['text_primary'],
                                              bg='#dcdad5',
                                              bd=1, relief='solid', padx=8, pady=6)
        # Will be shown/hidden based on destination mode - positioned at row 3
        
        # Setup path cleaning
        self.setup_path_cleaning(self.report_path)
        self.setup_path_cleaning(self.target_pbip_path)
    
    def _update_guide_text(self):
        """Update guide text based on selected copy modes"""
        if not hasattr(self, 'guide_frame'):
            return
        
        # Clear existing labels
        for label in self.guide_labels:
            label.destroy()
        self.guide_labels = []
        
        # Unpack the analyze button temporarily if it exists
        if hasattr(self, 'analyze_button') and self.analyze_button:
            self.analyze_button.pack_forget()
        
        # Determine guide text based on content mode and destination
        is_full_page = self.copy_content_mode.get() == "full_page"
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"
        
        if is_full_page and not is_cross_pbip:
            guide_text = [
                "🚀 QUICK START GUIDE:",
                "1. Select your .pbip report file",
                "2. Click 'Analyze Report' to scan",
                "3. Select pages to copy",
                "4. Click 'Execute Copy'",
                "5. Pages copied with '(Copy)' suffix"
            ]
        elif is_full_page and is_cross_pbip:
            guide_text = [
                "🚀 QUICK START GUIDE:",
                "1. Select SOURCE .pbip file",
                "2. Select TARGET .pbip file",
                "3. Click 'Analyze Report'",
                "4. Select pages to copy",
                "5. Click 'Execute Copy'"
            ]
        elif not is_full_page and not is_cross_pbip:
            guide_text = [
                "🚀 QUICK START GUIDE:",
                "1. Select your .pbip report file",
                "2. Click 'Analyze Report'",
                "3. Pick source page + bookmarks",
                "4. Select target pages",
                "5. Click 'Execute Copy'"
            ]
        else:  # bookmark + cross-pbip
            guide_text = [
                "🚀 QUICK START GUIDE:",
                "1. Select SOURCE .pbip file",
                "2. Select TARGET .pbip file",
                "3. Click 'Analyze Report'",
                "4. Pick source page + bookmarks",
                "5. Select target pages in TARGET",
                "6. Click 'Execute Copy'"
            ]
        
        warning_text = "⚠️ Requires PBIP format with TMDLs files"
        
        # Create new labels
        for i, text in enumerate(guide_text):
            if i == 0:  # Title
                label = tk.Label(self.guide_frame, text=text, 
                                font=('Segoe UI', 10, 'bold'), 
                                foreground=AppConstants.COLORS['info'],
                                bg='#f8fafc')
                label.pack(anchor=tk.W)
                self.guide_labels.append(label)
            else:  # Steps
                label = tk.Label(self.guide_frame, text=f"   {text}", 
                                font=('Segoe UI', 9),
                                foreground=AppConstants.COLORS['text_secondary'],
                                bg='#f8fafc',
                                wraplength=300)
                label.pack(anchor=tk.W, pady=1)
                self.guide_labels.append(label)
        
        # Add warning text at bottom in orange/red and italicized
        warning_label = tk.Label(self.guide_frame, text=warning_text,
                                font=('Segoe UI', 9, 'italic'),
                                foreground='#d97706',
                                bg='#f8fafc')
        warning_label.pack(anchor=tk.W, pady=(5, 0))
        self.guide_labels.append(warning_label)
        
        # Position the analyze button at the bottom
        if hasattr(self, 'analyze_button') and self.analyze_button:
            self.analyze_button.pack(anchor=tk.W, pady=(15, 0))
    
    def _show_target_pbip_input(self):
        """Show target PBIP file input for cross-PBIP destination mode"""
        if not self.target_pbip_frame:
            return
        
        # Clear existing content
        for widget in self.target_pbip_frame.winfo_children():
            widget.destroy()
        
        # Configure grid
        self.target_pbip_frame.columnconfigure(1, weight=1)
        
        # Add target file input with label
        label_text = "Select the target PBIP file where content will be copied:"
        tk.Label(self.target_pbip_frame, text=label_text,
                bg='#dcdad5',
                fg=AppConstants.COLORS['text_secondary'],
                font=('Segoe UI', 9)).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        tk.Label(self.target_pbip_frame, text="Target File:",
                bg='#dcdad5',
                fg=AppConstants.COLORS['text_primary'],
                font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        entry = ttk.Entry(self.target_pbip_frame, textvariable=self.target_pbip_path, width=80)
        entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=5)
        
        ttk.Button(self.target_pbip_frame, text="📂 Browse",
                  command=self.browse_target_file).grid(row=1, column=2, pady=5)
        
        # Show the frame at row 3 (after destination selection)
        self.target_pbip_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))
        
        # Setup path cleaning for target
        self.target_pbip_path.trace('w', lambda *args: self._on_path_change())
        
        # Adjust window height to accommodate target PBIP input
        self._adjust_window_height_for_target_pbip(True)
        
        self.log_message("ℹ️ Cross-PBIP destination selected - select target PBIP file")
    
    def _hide_target_pbip_input(self):
        """Hide target PBIP file input"""
        if self.target_pbip_frame:
            self.target_pbip_frame.grid_remove()
            self.target_pbip_path.set("")  # Clear target path
            
            # Adjust window height back down
            self._adjust_window_height_for_target_pbip(False)
    
    def _adjust_window_height_for_target_pbip(self, show_target_input: bool):
        """Adjust layout when target PBIP input is shown/hidden - NO window resizing"""
        # DON'T resize the window - just let the frames reorganize themselves
        # The grid layout will handle the spacing automatically
        # User controls window size, we just show/hide sections
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
