"""
Advanced Page Copy UI Tab - Refactored to use BaseToolTab
Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional, Dict, Any, List

from core.constants import AppConstants
from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin
from tools.page_copy.page_copy_core import AdvancedPageCopyEngine, PageCopyError


class AdvancedPageCopyTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """Advanced Page Copy tab - refactored to use new base architecture"""
    
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "page_copy", "Advanced Page Copy")
        
        # UI Variables
        self.report_path = tk.StringVar()
        
        # Core components
        self.page_copy_engine = AdvancedPageCopyEngine(logger_callback=self.log_message)
        
        # UI Components
        self.analyze_button = None
        self.copy_button = None
        self.pages_listbox = None
        self.pages_frame = None
        self.selection_label = None
        
        # State
        self.analysis_results = None
        self.available_pages = []
        
        # Setup UI and events
        self.setup_ui()
        self._setup_events()
    
    def setup_ui(self) -> None:
        """Setup the UI for the advanced page copy tab"""
        # Configure grid weights
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(2, weight=1)  # Log section gets weight
        
        # Setup sections
        self._setup_data_source()
        self._setup_page_selection()
        
        # Create log section using base class
        log_components = self.create_log_section(self.frame)
        log_components['frame'].grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # Create action buttons using base class
        button_frame = ttk.Frame(self.frame)
        button_frame.grid(row=4, column=0, pady=(20, 0))  # Changed from row=3 to row=4
        
        self.copy_button = ttk.Button(button_frame, text="🚀 EXECUTE COPY", 
                                    command=self.start_copy,
                                    style='Action.TButton', state=tk.DISABLED)
        self.copy_button.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Button(button_frame, text="🔄 RESET ALL", command=self.reset_tab,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 15))
        
        # Create progress bar using base class
        self.create_progress_bar(self.frame)
        
        # Show welcome message
        self._show_welcome_message()
    
    def _position_progress_frame(self):
        """Position progress frame specifically for Advanced Page Copy layout"""
        if self.progress_frame:
            # Position between the log section (row 2) and action buttons (row 3)
            # Insert progress at row 3, and move buttons to row 4
            self.progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
            
            # Move button frame to row 4 if it exists
            for child in self.frame.winfo_children():
                if isinstance(child, ttk.Frame):
                    # Check if this frame contains buttons
                    has_buttons = any(isinstance(widget, ttk.Button) for widget in child.winfo_children())
                    if has_buttons:
                        child.grid_configure(row=4)
    
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
            "⚠️ Requires PBIP format"
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
        
        # Create custom guide frame with proper alignment
        guide_frame = ttk.Frame(content_frame)
        guide_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 35))
        
        # Add guide text with custom formatting
        guide_text = [
            "🚀 QUICK START GUIDE:",
            "1. Select your .pbip report file below",
            "2. Click 'Analyze Report' to scan for pages with",
            "   bookmarks", 
            "3. Select which pages to copy from the list",
            "4. Pages will be duplicated with all bookmarks",
            "   preserved",
            "5. Click 'Execute Copy' to create the duplicates",
            "6. New pages will have '(Copy)' suffix to avoid conflicts",
            "⚠️ Requires PBIP format"
        ]
        
        for i, text in enumerate(guide_text):
            if i == 0:  # Title
                ttk.Label(guide_frame, text=text, 
                         font=('Segoe UI', 10, 'bold'), 
                         foreground=AppConstants.COLORS['info']).pack(anchor=tk.W)
            else:  # Steps - use different padding for continuation lines
                if text.startswith('   '):  # Continuation line
                    ttk.Label(guide_frame, text=f"      {text.strip()}", font=('Segoe UI', 9),
                             foreground=AppConstants.COLORS['text_secondary'], 
                             wraplength=300).pack(anchor=tk.W, pady=1)
                else:  # Regular step
                    ttk.Label(guide_frame, text=f"   {text}", font=('Segoe UI', 9),
                             foreground=AppConstants.COLORS['text_secondary'], 
                             wraplength=300).pack(anchor=tk.W, pady=1)
        
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
        
        # Analyze button
        self.analyze_button = ttk.Button(input_frame, text="🔍 ANALYZE REPORT",
                                       command=self.analyze_report, 
                                       style='Action.TButton', state=tk.DISABLED)
        self.analyze_button.grid(row=1, column=0, columnspan=3, pady=(15, 0))
        
        # Setup path cleaning
        self.setup_path_cleaning(self.report_path)
    
    def _setup_page_selection(self):
        """Setup page selection section (initially hidden)"""
        self.pages_frame = ttk.LabelFrame(self.frame, text="📋 PAGE SELECTION", 
                                        style='Section.TLabelframe', padding="20")
        # Will be shown after analysis
    
    def _setup_events(self):
        """Setup event handlers"""
        self.report_path.trace('w', lambda *args: self._on_path_change())
    
    def _show_welcome_message(self):
        """Show welcome message"""
        messages = [
            "🎉 Welcome to Advanced Page Copy!",
            "📋 Select a report file to analyze pages with bookmarks...",
            "⚠️ Requires PBIP format files only",
            "🔄 This tool duplicates pages within the same report",
            "=" * 60
        ]
        for msg in messages:
            self.log_message(msg)
    
    def _on_path_change(self):
        """Handle path changes"""
        self._update_ui_state()
    
    def _update_ui_state(self):
        """Update UI state"""
        has_file = bool(self.report_path.get())
        if self.analyze_button:
            self.analyze_button.config(state=tk.NORMAL if has_file else tk.DISABLED)
    
    def _show_page_selection_ui(self, pages_with_bookmarks: List[Dict[str, Any]]):
        """Show page selection UI after analysis"""
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
        instruction_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 20))
        
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
            height=8,
            font=('Segoe UI', 10),
            bg=AppConstants.COLORS['surface'],
            fg=AppConstants.COLORS['text_primary'],
            selectbackground=AppConstants.COLORS['accent'],
            relief='solid',
            borderwidth=1
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
        
        # Selection summary
        self.selection_label = ttk.Label(list_frame, text="No pages selected", 
                                       font=('Segoe UI', 9),
                                       foreground=AppConstants.COLORS['text_secondary'])
        self.selection_label.grid(row=1, column=0, pady=(10, 0))
    
    def _hide_page_selection_ui(self):
        """Hide page selection UI"""
        if self.pages_frame:
            self.pages_frame.grid_remove()
            self._adjust_window_height(False)
    
    def _adjust_window_height(self, show_page_selection: bool):
        """Adjust window height based on page selection visibility"""
        try:
            if hasattr(self.main_app, 'root'):
                root = self.main_app.root
                base_height = 860
                page_selection_height = 250
                
                if show_page_selection:
                    new_height = base_height + page_selection_height
                else:
                    new_height = base_height
                
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
    
    # Public methods
    def browse_file(self):
        """Browse for report file"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Select Report (.pbip file - PBIP format required)",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.report_path.set(file_path)
    
    def analyze_report(self):
        """Analyze selected report"""
        try:
            report_path = self.clean_file_path(self.report_path.get())
            # Basic validation happens in the engine
        except Exception as e:
            self.show_error("Validation Error", str(e))
            return
        
        # Use base class background processing
        self.run_in_background(
            target_func=self._analyze_thread_target,
            success_callback=self._handle_analysis_complete,
            error_callback=lambda e: self.show_error("Analysis Error", str(e))
        )
    
    def _analyze_thread_target(self):
        """Background analysis logic"""
        self.update_progress(10, "Validating report file...", persist=True)
        report_path = self.clean_file_path(self.report_path.get())
        
        self.update_progress(30, "Reading report structure...", persist=True)
        
        self.update_progress(60, "Analyzing pages and bookmarks...", persist=True)
        results = self.page_copy_engine.analyze_report_pages(report_path)
        
        self.update_progress(90, "Preparing results...", persist=True)
        
        self.update_progress(100, "Analysis complete!", persist=True)
        return results
    
    def _handle_analysis_complete(self, results):
        """Handle analysis completion"""
        self.analysis_results = results
        pages_with_bookmarks = results['pages_with_bookmarks']
        
        if not pages_with_bookmarks:
            self.log_message("⚠️ No pages with bookmarks found!")
            self.log_message("   Only pages with bookmarks can be copied with this tool.")
            self.show_warning("No Copyable Pages", 
                             "No pages with bookmarks were found in this report.\n\n"
                             "This tool only copies pages that have associated bookmarks.")
            return
        
        # Show page selection UI
        self._show_page_selection_ui(pages_with_bookmarks)
        
        # Ensure progress bar stays visible after analysis
        if self.progress_frame and not self.progress_frame.winfo_viewable():
            self._position_progress_frame()
        
        # Show analysis summary
        self._show_analysis_summary(results)
    
    def _show_analysis_summary(self, results):
        """Show analysis summary"""
        self.log_message("\n📊 ANALYSIS SUMMARY")
        self.log_message("=" * 50)
        
        report = results['report']
        copyable_pages = len(results['pages_with_bookmarks'])
        total_pages = results['analysis_summary']['total_pages']
        total_bookmarks = results['analysis_summary']['total_bookmarks']
        
        self.log_message(f"📄 Report: {report['name']}")
        self.log_message(f"📄 Total Pages: {total_pages}")
        self.log_message(f"📋 Pages with Bookmarks: {copyable_pages}")
        self.log_message(f"🔖 Total Bookmarks: {total_bookmarks}")
        
        if copyable_pages > 0:
            self.log_message("\n📋 COPYABLE PAGES:")
            for page in results['pages_with_bookmarks']:
                self.log_message(f"   • {page['display_name']} ({page['bookmark_count']} bookmarks)")
            
            self.log_message(f"\n✅ Select pages above and click 'EXECUTE COPY' to proceed")
        else:
            self.log_message("\n⚠️ No pages available for copying")
    
    def start_copy(self):
        """Start copy operation"""
        if not self.analysis_results:
            self.show_error("Error", "Please analyze report first")
            return
        
        if not self.pages_listbox:
            self.show_error("Error", "No pages available for selection")
            return
        
        # Get selected pages
        selected_indices = self.pages_listbox.curselection()
        if not selected_indices:
            self.show_error("Error", "Please select at least one page to copy")
            return
        
        selected_pages = [self.available_pages[i] for i in selected_indices]
        page_names = [page['name'] for page in selected_pages]
        
        # Confirm operation
        if not self.ask_yes_no("Confirm Copy", 
                              f"Ready to copy {len(selected_pages)} page(s)?\n\n"
                              f"📋 Pages to copy:\n" + 
                              "\n".join([f"• {p['display_name']} ({p['bookmark_count']} bookmarks)" 
                                       for p in selected_pages[:5]]) +
                              (f"\n... and {len(selected_pages)-5} more" if len(selected_pages) > 5 else "") +
                              f"\n\n💾 Report: {Path(self.report_path.get()).name}"):
            return
        
        # Use base class background processing
        self.run_in_background(
            target_func=lambda: self._copy_thread_target(page_names),
            success_callback=self._handle_copy_complete,
            error_callback=lambda e: self.show_error("Copy Error", str(e))
        )
    
    def _copy_thread_target(self, selected_page_names: List[str]):
        """Background copy logic"""
        self.log_message("\n🚀 Starting page copy operation...")
        
        self.update_progress(10, "Preparing copy operation...")
        report_path = self.clean_file_path(self.report_path.get())
        
        self.update_progress(30, "Reading report data...")
        
        page_count = len(selected_page_names)
        if page_count == 1:
            self.update_progress(50, f"Copying page with bookmarks...")
        else:
            self.update_progress(50, f"Copying {page_count} pages with bookmarks...")
        
        success = self.page_copy_engine.copy_selected_pages(
            report_path, selected_page_names, self.analysis_results
        )
        
        self.update_progress(90, "Finalizing report updates...")
        
        self.update_progress(100, "Copy operation complete!")
        return {'success': success, 'report_path': report_path, 'page_count': len(selected_page_names)}
    
    def _handle_copy_complete(self, result):
        """Handle copy completion"""
        if result['success']:
            self.log_message("✅ PAGE COPY COMPLETED SUCCESSFULLY!")
            self.log_message(f"💾 Report updated: {result['report_path']}")
            
            self.show_info("Copy Complete", 
                          f"Page copy completed successfully!\n\n"
                          f"📋 Copied {result['page_count']} page(s) with bookmarks\n"
                          f"💾 Report: {Path(result['report_path']).name}\n\n"
                          f"The copied pages have been added to your report with '(Copy)' suffix.")
        else:
            self.show_error("Copy Failed", "The page copy operation failed. Check the log for details.")
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        if self.is_busy:
            if not self.ask_yes_no("Confirm Reset", "An operation is in progress. Stop and reset?"):
                return
        
        # Clear state
        self.report_path.set("")
        self.analysis_results = None
        self.available_pages = []
        
        # Reset progress persistence
        self.progress_persist = False
        if self.progress_frame:
            self.progress_frame.grid_remove()
        
        # Reset UI state
        if self.analyze_button:
            self.analyze_button.config(state=tk.DISABLED)
        if self.copy_button:
            self.copy_button.config(state=tk.DISABLED)
        if self.pages_listbox:
            self.pages_listbox.config(state=tk.NORMAL)
        self._hide_page_selection_ui()
        
        # Clear log and show welcome
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
        
        self._show_welcome_message()
        self.log_message("✅ Advanced Page Copy reset successfully!")
    
    def show_help_dialog(self) -> None:
        """Show help dialog specific to advanced page copy"""
        def create_help_content(help_window):
            # Main container that reserves space for the button
            container = ttk.Frame(help_window, padding="20")
            container.pack(fill=tk.BOTH, expand=True)
            
            # Content frame for scrollable content (if needed)
            content_frame = ttk.Frame(container)
            content_frame.pack(fill=tk.BOTH, expand=True)
            
            # Header
            ttk.Label(content_frame, text="📋 Advanced Page Copy - Help", 
                     font=('Segoe UI', 16, 'bold'), 
                     foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 20))
            
            # Orange warning box (similar to other help dialogs)
            warning_frame = ttk.Frame(content_frame)
            warning_frame.pack(fill=tk.X, pady=(0, 20))
            
            warning_container = tk.Frame(warning_frame, bg=AppConstants.COLORS['warning'], 
                                       padx=15, pady=10, relief='solid', borderwidth=2)
            warning_container.pack(fill=tk.X)
            
            ttk.Label(warning_container, text="⚠️  IMPORTANT DISCLAIMERS & REQUIREMENTS", 
                     font=('Segoe UI', 12, 'bold'), 
                     background=AppConstants.COLORS['warning'],
                     foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W)
            
            warnings = [
                "• This tool ONLY works with PBIP enhanced report format (PBIR) files",
                "• This is NOT officially supported by Microsoft - use at your own discretion",
                "• Look for .pbip files with definition\\ folder (not report.json files)",
                "• Always keep backups of your original reports before copying pages",
                "• Test thoroughly and validate copied pages before production use",
                "• Enable 'Store reports using enhanced metadata format (PBIR)' in Power BI Desktop"
            ]
            
            for warning in warnings:
                ttk.Label(warning_container, text=warning, font=('Segoe UI', 10),
                         background=AppConstants.COLORS['warning'],
                         foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W, pady=1)
            
            # Help sections
            help_sections = [
                ("📋 What This Tool Does", [
                    "✅ Duplicates pages within the same report",
                    "✅ Preserves all bookmarks associated with each page",
                    "✅ Automatically renames copies to avoid conflicts",
                    "✅ Updates report metadata after copying"
                ]),
                ("📁 File Requirements", [
                    "✅ Only .pbip files (enhanced PBIR format) are supported",
                    "✅ Only pages with bookmarks are shown for copying",
                    "❌ Pages without bookmarks cannot be copied with this tool"
                ]),
                ("⚠️ Important Notes", [
                    "• Copied pages will have '(Copy)' suffix in their names",
                    "• Bookmarks are duplicated and renamed to avoid conflicts",
                    "• Always backup your report before making changes",
                    "• NOT officially supported by Microsoft"
                ])
            ]
            
            for title, items in help_sections:
                section_frame = ttk.Frame(content_frame)
                section_frame.pack(fill=tk.X, pady=(0, 15))
                
                ttk.Label(section_frame, text=title, font=('Segoe UI', 12, 'bold'),
                         foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
                
                for item in items:
                    ttk.Label(section_frame, text=f"   {item}", font=('Segoe UI', 10)).pack(anchor=tk.W, pady=1)
            
            # Button frame at bottom - fixed position
            button_frame = ttk.Frame(container)
            button_frame.pack(fill=tk.X, pady=(20, 0), side=tk.BOTTOM)
            
            close_button = ttk.Button(button_frame, text="❌ Close", 
                                    command=help_window.destroy,
                                    style='Action.TButton')
            close_button.pack(pady=(10, 0))
        
        # Create custom help window for Advanced Page Copy (independent of base class)
        help_window = tk.Toplevel(self.main_app.root)
        help_window.title("Advanced Page Copy - Help")
        help_window.geometry("670x750")  # Adjusted height for optimal spacing
        help_window.resizable(False, False)
        help_window.transient(self.main_app.root)
        help_window.grab_set()
        
        # Center window
        help_window.geometry(f"+{self.main_app.root.winfo_rootx() + 50}+{self.main_app.root.winfo_rooty() + 50}")
        help_window.configure(bg=AppConstants.COLORS['background'])
        
        # Create content
        create_help_content(help_window)
        
        # Bind escape key
        help_window.bind('<Escape>', lambda event: help_window.destroy())
