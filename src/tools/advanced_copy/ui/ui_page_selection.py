"""
Page Selection Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Handles page selection UI for full page copy mode.
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Any

from core.constants import AppConstants


class PageSelectionMixin:
    """
    Mixin for page selection functionality.
    
    Methods extracted from AdvancedCopyTab:
    - _setup_page_selection()
    - _show_page_selection_ui()
    - _show_full_page_selection_ui()
    - _hide_page_selection_ui()
    - _adjust_window_height()
    - _on_page_selection_change()
    """
    
    def _setup_page_selection(self):
        """Setup page selection section (initially hidden)"""
        self.pages_frame = ttk.LabelFrame(self.frame, text="📋 PAGE SELECTION", 
                                        style='Section.TLabelframe', padding="10")
        # Will be shown after analysis
    
    def _show_page_selection_ui(self, pages_with_bookmarks: List[Dict[str, Any]]):
        """Show page selection UI after analysis - routes based on content mode"""
        if self.copy_content_mode.get() == "full_page":
            self._show_full_page_selection_ui(pages_with_bookmarks)
        else:  # bookmark_visual mode
            self._show_bookmark_mode_selection_ui(pages_with_bookmarks)
    
    def _show_full_page_selection_ui(self, pages_with_bookmarks: List[Dict[str, Any]]):
        """Show page selection UI for full page copy mode"""
        # Show the pages frame
        self.pages_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Adjust window height to accommodate page selection
        self._adjust_window_height(True)
        
        # Clear existing content
        for widget in self.pages_frame.winfo_children():
            widget.destroy()
        
        # Create selection interface
        content_frame = ttk.Frame(self.pages_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # LEFT: Instructions
        instruction_frame = ttk.Frame(content_frame)
        instruction_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 15))
        
        ttk.Label(instruction_frame, text="📋 SELECT PAGES TO COPY:", 
                 font=('Segoe UI', 11, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
        
        instructions = [
            "Pages shown below have bookmarks",
            "and can be safely copied with all",
            "their associated bookmarks.",
            "",
            "✅ Use Ctrl+Click for multiple",
            "✅ Use Shift+Click for ranges"
        ]
        
        for instruction in instructions:
            style = 'normal' if instruction and not instruction.startswith('✅') else 'info'
            color = AppConstants.COLORS['text_primary'] if style == 'normal' else AppConstants.COLORS['info']
            
            ttk.Label(instruction_frame, text=instruction, 
                     font=('Segoe UI', 9 if style == 'info' else 10),
                     foreground=color).pack(anchor=tk.W)
        
        # RIGHT: Page list with scrollbar
        list_frame = ttk.Frame(content_frame)
        list_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Create listbox with scrollbar
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_container.columnconfigure(0, weight=1)
        list_container.rowconfigure(0, weight=1)
        
        # Listbox for page selection
        self.pages_listbox = tk.Listbox(
            list_container, 
            selectmode=tk.EXTENDED,  # Allow multiple selection
            height=4,  # Reduced from 8 to 4 for more compact display
            font=('Segoe UI', 10),
            bg=AppConstants.COLORS['surface'],
            fg=AppConstants.COLORS['text_primary'],
            selectbackground=AppConstants.COLORS['accent'],
            relief='solid',
            borderwidth=1,
            exportselection=False  # Prevent selection interference
        )
        self.pages_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for listbox
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.pages_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.pages_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Populate listbox
        self.available_pages = pages_with_bookmarks
        for page in pages_with_bookmarks:
            display_text = f"{page['display_name']} ({page['bookmark_count']} bookmarks)"
            self.pages_listbox.insert(tk.END, display_text)
        
        # Bind selection events
        self.pages_listbox.bind('<<ListboxSelect>>', self._on_page_selection_change)
        
        # Number of Copies control (for Full Page Copy mode)
        copies_frame = ttk.Frame(list_frame)
        copies_frame.grid(row=2, column=0, pady=(5, 0), sticky=tk.W)
        
        ttk.Label(copies_frame, text="Number of copies per page:",
                 font=('Segoe UI', 9, 'bold'),
                 foreground=AppConstants.COLORS['text_primary']).pack(side=tk.LEFT, padx=(0, 10))
        
        # Spinbox for number of copies (1-5)
        copies_spinbox = ttk.Spinbox(copies_frame, from_=1, to=5, width=5,
                                    textvariable=self.num_copies, state='readonly')
        copies_spinbox.pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Label(copies_frame, text="💡 Each selected page will be duplicated this many times",
                font=('Segoe UI', 8),
                fg=AppConstants.COLORS['info'],
                bg=AppConstants.COLORS['background']).pack(side=tk.LEFT)
        
        # Selection summary
        self.selection_label = ttk.Label(list_frame, text="No pages selected", 
                                       font=('Segoe UI', 9),
                                       foreground=AppConstants.COLORS['text_secondary'])
        self.selection_label.grid(row=3, column=0, pady=(5, 0))
    
    def _hide_page_selection_ui(self):
        """Hide page selection UI"""
        if self.pages_frame:
            self.pages_frame.grid_remove()
            self._adjust_window_height(False)
    
    def _adjust_window_height(self, show_page_selection: bool):
        """Adjust window height based on page selection visibility
        
        Sets a FIXED height when showing page selection (not additive).
        This ensures the Execute Copy buttons at the bottom are fully visible.
        """
        try:
            if hasattr(self.main_app, 'root'):
                root = self.main_app.root
                
                # Define FIXED heights for different states
                base_height = 900  # Height before analysis (no page selection)
                post_analysis_height = 1175  # FIXED height after analysis (includes page selection + buttons + action area)
                
                # Add extra height when in cross-PBIP mode (target input showing)
                cross_pbip_adjustment = 55 if self.copy_destination_mode.get() == "cross_pbip" else 0
                
                if show_page_selection:
                    # Set FIXED post-analysis height (not additive!)
                    new_height = post_analysis_height + cross_pbip_adjustment
                else:
                    # Back to base height (no page selection)
                    new_height = base_height + cross_pbip_adjustment
                
                # Update window geometry
                current_geometry = root.geometry()
                parts = current_geometry.split('x')
                if len(parts) >= 2:
                    width = parts[0]
                    position = parts[1].split('+', 1)[1] if '+' in parts[1] else ''
                    new_geometry = f"{width}x{new_height}+{position}"
                    root.geometry(new_geometry)
        except Exception:
            pass
    
    def _on_page_selection_change(self, event=None):
        """Handle page selection changes"""
        if not self.pages_listbox:
            return
        
        selected_indices = self.pages_listbox.curselection()
        selected_count = len(selected_indices)
        
        if selected_count == 0:
            self.selection_label.config(text="No pages selected")
            if self.copy_button:
                self.copy_button.config(state=tk.DISABLED)
        elif selected_count == 1:
            page_name = self.available_pages[selected_indices[0]]['display_name']
            self.selection_label.config(text=f"Selected: {page_name}")
            if self.copy_button:
                self.copy_button.config(state=tk.NORMAL)
        else:
            self.selection_label.config(text=f"Selected: {selected_count} pages")
            if self.copy_button:
                self.copy_button.config(state=tk.NORMAL)
