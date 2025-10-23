"""
Helpers Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Contains helper methods for messages and dialogs.
"""

import tkinter as tk
from tkinter import ttk

from core.constants import AppConstants


class HelpersMixin:
    """
    Mixin for helper methods and UI utilities.
    
    Methods extracted from AdvancedCopyTab:
    - _show_welcome_message()
    - _show_analysis_summary()
    - _position_progress_frame()
    - show_help_dialog()
    """
    
    def _show_welcome_message(self):
        """Show welcome message"""
        messages = [
            "🎉 Welcome to Advanced Copy!",
            "📋 Select a report file and copy mode to begin...",
            "📄 Full Page Copy: Duplicates entire pages with bookmarks",
            "🔖 Bookmark+Visual Copy: Pop out bookmarks to multiple pages",
            "⚠️ Requires PBIP format files only",
            "=" * 60
        ]
        for msg in messages:
            self.log_message(msg)
    
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
        
        # Display bookmark navigator analysis
        navigator_analysis = results.get('navigator_analysis', {})
        if navigator_analysis and navigator_analysis['total_navigators'] > 0:
            self.log_message(f"\n🧭 BOOKMARK NAVIGATOR ANALYSIS:")
            self.log_message(f"   📊 Total Navigators: {navigator_analysis['total_navigators']}")
            
            for page_nav in navigator_analysis['pages_with_navigators']:
                self.log_message(f"   📄 Page '{page_nav['page_name']}' has {len(page_nav['navigators'])} bookmark navigator(s):")
                
                for nav in page_nav['navigators']:
                    # Show visual title (name from Power BI) if available, otherwise show ID
                    if nav['visual_title'] and nav['visual_title'] != nav['visual_name']:
                        nav_display = nav['visual_title']
                    else:
                        nav_display = f"Navigator (ID: {nav['visual_name']})"
                    
                    if nav['has_group']:
                        if nav['group_exists']:
                            self.log_message(f"      ✅ '{nav_display}' - Assigned to group (ID: {nav['group_id']})")
                        else:
                            self.log_message(f"      ⚠️ '{nav_display}' - References unknown group (ID: {nav['group_id']})")
                            self.log_message(f"         💡 Verify this group exists in bookmarks.json")
                    else:
                        self.log_message(f"      ⚠️ '{nav_display}' - NO GROUP ASSIGNED (shows all bookmarks)")
                        self.log_message(f"         💡 Bookmarks from both original and copied pages will appear in this navigator")
        
        if copyable_pages > 0:
            self.log_message("\n📋 COPYABLE PAGES:")
            for page in results['pages_with_bookmarks']:
                self.log_message(f"   • {page['display_name']} ({page['bookmark_count']} bookmarks)")
            
            self.log_message(f"\n✅ Select pages above and click 'EXECUTE COPY' to proceed")
            
            # Add prominent warning summary at BOTTOM if navigators without groups detected
            if navigator_analysis and navigator_analysis['navigators_without_groups'] > 0:
                self.log_message(f"\n{'='*50}")
                self.log_message(f"⚠️  WARNING: BOOKMARK NAVIGATORS WITHOUT GROUPS DETECTED")
                self.log_message(f"{'='*50}")
                self.log_message(f"   {navigator_analysis['navigators_without_groups']} navigator(s) show ALL bookmarks")
                self.log_message(f"   After copying, BOTH original and copied bookmarks will appear together")
                self.log_message(f"   ✅ RECOMMENDED: Assign groups to navigators before copying")
                self.log_message(f"{'='*50}")
        else:
            self.log_message("\n⚠️ No pages available for copying")
    
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
            ttk.Label(content_frame, text="📋 Advanced Copy - Help", 
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
            
            # Two-column layout for help sections
            columns_frame = ttk.Frame(content_frame)
            columns_frame.pack(fill=tk.BOTH, expand=True)
            columns_frame.columnconfigure(0, weight=1)
            columns_frame.columnconfigure(1, weight=1)
            
            # LEFT COLUMN
            left_column = ttk.Frame(columns_frame)
            left_column.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
            
            # RIGHT COLUMN
            right_column = ttk.Frame(columns_frame)
            right_column.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
            
            # Left column sections
            left_sections = [
                ("📋 What This Tool Does", [
                    "✅ Offers TWO copy modes: Full Page Copy and Bookmark+Visual Copy",
                    "✅ Duplicates complete pages OR specific bookmarks with visuals",
                    "✅ Preserves all bookmarks and their visual relationships",
                    "✅ Automatically renames copies to avoid conflicts",
                    "✅ Updates report metadata after copying",
                    "✅ Supports group selection - click groups to select all children"
                ]),
                ("📄 Full Page Copy Mode", [
                    "✅ Duplicates entire pages with all visuals and bookmarks",
                    "✅ Creates complete copies including all page elements",
                    "✅ Useful for creating template variations or backup pages",
                    "📝 Copied pages get '(Copy)' suffix to avoid name conflicts"
                ]),
                ("🔖 Bookmark + Visual Copy Mode", [
                    "✅ Copies specific bookmarks and their visuals to existing pages",
                    "✅ 'Pop out' bookmark controls to multiple pages",
                    "✅ Visuals keep SAME IDs across pages (shared visuals)",
                    "✅ Creates NEW bookmarks for each target page",
                    "💡 Perfect for dashboard navigation across multiple pages",
                    "⚠️ Only copies bookmark navigators relevant to selected bookmarks",
                    "⚠️ ONLY 'Selected visuals' configured bookmarks are supported",
                    "⚠️ Bookmarks set to 'All visuals' cannot be copied"
                ])
            ]
            
            # Right column sections
            right_sections = [
                ("📁 File Requirements", [
                    "✅ Only .pbip files (enhanced PBIR format) are supported",
                    "✅ Full Page Mode: Only pages with bookmarks shown for copying",
                    "✅ Bookmark Mode: Can target any page (with or without bookmarks)",
                    "⚠️ Bookmark Mode: Requires 'Selected visuals' configured bookmarks",
                    "❌ Standard .pbix files are NOT supported"
                ]),
                ("⚠️ Important Notes", [
                    "• Full Page: Creates complete duplicate pages",
                    "• Bookmark Mode: Shares visuals across pages (same IDs)",
                    "• Bookmark Mode: Only 'Selected visuals' bookmarks are supported",
                    "• 'All visuals' bookmarks cannot be copied (capture entire page state)",
                    "• Disabled bookmarks appear grayed out in the UI",
                    "• Bookmark groups are preserved and duplicated appropriately",
                    "• Always backup your report before making changes",
                    "• NOT officially supported by Microsoft"
                ]),
                ("🔖 Configuring Bookmarks for Copy", [
                    "To make bookmarks copyable in Bookmark+Visual mode:",
                    "1. Open bookmark in Power BI Desktop",
                    "2. Uncheck 'All visuals' option",
                    "3. Select specific visuals to include",
                    "4. Save the bookmark",
                    "5. Bookmark will now show as copyable in the tool"
                ])
            ]
            
            # Render left column sections
            for title, items in left_sections:
                section_frame = ttk.Frame(left_column)
                section_frame.pack(fill=tk.X, pady=(0, 15))
                
                ttk.Label(section_frame, text=title, font=('Segoe UI', 11, 'bold'),
                         foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
                
                for item in items:
                    ttk.Label(section_frame, text=f"   {item}", font=('Segoe UI', 9),
                             wraplength=380).pack(anchor=tk.W, pady=1)
            
            # Render right column sections
            for title, items in right_sections:
                section_frame = ttk.Frame(right_column)
                section_frame.pack(fill=tk.X, pady=(0, 15))
                
                ttk.Label(section_frame, text=title, font=('Segoe UI', 11, 'bold'),
                         foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
                
                for item in items:
                    ttk.Label(section_frame, text=f"   {item}", font=('Segoe UI', 9),
                             wraplength=380).pack(anchor=tk.W, pady=1)
            
            # Button frame at bottom - fixed position
            button_frame = ttk.Frame(container)
            button_frame.pack(fill=tk.X, pady=(20, 0), side=tk.BOTTOM)
            
            close_button = ttk.Button(button_frame, text="❌ Close", 
                                    command=help_window.destroy,
                                    style='Action.TButton')
            close_button.pack(pady=(10, 0))
        
        # Create custom help window for Advanced Copy (independent of base class)
        help_window = tk.Toplevel(self.main_app.root)
        help_window.title("Advanced Copy - Help")
        help_window.geometry("950x940")
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
        results = self.advanced_copy_engine.analyze_report_pages(report_path)
        
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
        
        # Check if we need to analyze target PBIP (Cross-PBIP + Bookmark mode)
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"
        is_bookmark_mode = self.copy_content_mode.get() == "bookmark_visual"
        
        if is_cross_pbip and is_bookmark_mode:
            # Need to analyze target PBIP to get its pages
            self.log_message("\n🌍 Cross-PBIP bookmark mode detected - analyzing TARGET report...")
            
            # Run target analysis in background
            self.run_in_background(
                target_func=self._analyze_target_thread_target,
                success_callback=lambda target_results: self._handle_target_analysis_complete(results, target_results),
                error_callback=lambda e: self.show_error("Target Analysis Error", str(e))
            )
        else:
            # No target analysis needed - proceed directly
            self._show_page_selection_ui(pages_with_bookmarks)
            
            # Ensure progress bar stays visible after analysis
            if self.progress_frame and not self.progress_frame.winfo_viewable():
                self._position_progress_frame()
            
            # Show analysis summary
            self._show_analysis_summary(results)
    
    def _analyze_target_thread_target(self):
        """Background analysis logic for target PBIP"""
        self.update_progress(10, "Validating target report file...", persist=True)
        target_path = self.clean_file_path(self.target_pbip_path.get())
        
        self.update_progress(30, "Reading target report structure...", persist=True)
        
        self.update_progress(60, "Analyzing target pages...", persist=True)
        results = self.advanced_copy_engine.analyze_report_pages(target_path)
        
        self.update_progress(90, "Preparing target results...", persist=True)
        
        self.update_progress(100, "Target analysis complete!", persist=True)
        return results
    
    def _handle_target_analysis_complete(self, source_results, target_results):
        """Handle target analysis completion (for Cross-PBIP bookmark mode)"""
        # Store target analysis
        self.target_analysis_results = target_results
        
        # Get pages from source
        pages_with_bookmarks = source_results['pages_with_bookmarks']
        
        # Log target analysis
        total_target_pages = len(target_results['pages_with_bookmarks']) + len(target_results['pages_without_bookmarks'])
        self.log_message(f"   ✅ Target report analyzed: {total_target_pages} pages available")
        
        # Show page selection UI with source pages
        self._show_page_selection_ui(pages_with_bookmarks)
        
        # Ensure progress bar stays visible after analysis
        if self.progress_frame and not self.progress_frame.winfo_viewable():
            self._position_progress_frame()
        
        # Show analysis summary for source
        self._show_analysis_summary(source_results)
    
    def start_copy(self):
        """Start copy operation - routes based on content mode"""
        if not self.analysis_results:
            self.show_error("Error", "Please analyze report first")
            return
        
        if self.copy_content_mode.get() == "full_page":
            self._start_full_page_copy()
        else:
            self._start_bookmark_copy()
    
    def _start_full_page_copy(self):
        """Start full page copy operation"""
        from pathlib import Path
        
        if not self.analysis_results:
            self.show_error("Error", "Please analyze report first")
            return
        
        # Check if cross-PBIP - now supported!
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"
        
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
        
        # Get number of copies
        num_copies = self.num_copies.get()
        total_pages_to_create = len(selected_pages) * num_copies
        
        # Validate target path if cross-PBIP
        if is_cross_pbip and not self.target_pbip_path.get():
            self.show_error("Error", "Please select a target PBIP file for cross-PBIP copy")
            return
        
        # Check for navigator warnings
        warning_message = ""
        if self.analysis_results:
            nav_analysis = self.analysis_results.get('navigator_analysis', {})
            if nav_analysis and nav_analysis['navigators_without_groups'] > 0:
                warning_message = (
                    f"\n\n{'─'*45}\n"
                    f"⚠️  WARNING: Bookmark Navigator Without Group Detected\n"
                    f"{'─'*45}\n"
                    f"This page contains navigators that show ALL bookmarks (not grouped).\n"
                    f"After copying, BOTH original and copied bookmarks will appear together.\n\n"
                    f"✅ RECOMMENDED: Cancel and assign groups to navigators first.\n"
                    f"{'─'*45}"
                )
        
        # Build confirmation message based on mode
        if is_cross_pbip:
            confirm_msg = f"Ready to copy {len(selected_pages)} page(s) from SOURCE to TARGET?{warning_message}\n\n"
            confirm_msg += f"📂 Source PBIP: {Path(self.report_path.get()).name}\n"
            confirm_msg += f"🎯 Target PBIP: {Path(self.target_pbip_path.get()).name}\n"
            if num_copies > 1:
                confirm_msg += f"🔁 Copies per page: {num_copies} (Total: {total_pages_to_create} pages will be created)\n"
            confirm_msg += "\n📋 Pages to copy:\n" + \
                          "\n".join([f"• {p['display_name']} ({p['bookmark_count']} bookmarks)" 
                                   for p in selected_pages[:5]]) + \
                          (f"\n... and {len(selected_pages)-5} more" if len(selected_pages) > 5 else "")
        else:
            confirm_msg = f"Ready to copy {len(selected_pages)} page(s)?{warning_message}\n\n"
            if num_copies > 1:
                confirm_msg += f"🔁 Copies per page: {num_copies} (Total: {total_pages_to_create} pages will be created)\n\n"
            confirm_msg += f"📋 Pages to copy:\n" + \
                          "\n".join([f"• {p['display_name']} ({p['bookmark_count']} bookmarks)" 
                                   for p in selected_pages[:5]]) + \
                          (f"\n... and {len(selected_pages)-5} more" if len(selected_pages) > 5 else "") + \
                          f"\n\n💾 Report: {Path(self.report_path.get()).name}"
        
        # Confirm operation
        if not self.ask_yes_no("Confirm Copy", confirm_msg):
            return
        
        # Use base class background processing - route based on destination mode
        if is_cross_pbip:
            self.run_in_background(
                target_func=lambda: self._copy_cross_pbip_thread_target(page_names, num_copies),
                success_callback=self._handle_copy_complete,
                error_callback=lambda e: self.show_error("Cross-PBIP Copy Error", str(e))
            )
        else:
            self.run_in_background(
                target_func=lambda: self._copy_thread_target(page_names, num_copies),
                success_callback=self._handle_copy_complete,
                error_callback=lambda e: self.show_error("Copy Error", str(e))
            )
    
    def _copy_cross_pbip_thread_target(self, selected_page_names, num_copies=1):
        """Background cross-PBIP copy logic for full pages"""
        self.log_message("\n🚀 Starting CROSS-PBIP full page copy operation...")
        
        self.update_progress(10, "Preparing cross-PBIP copy operation...")
        source_path = self.clean_file_path(self.report_path.get())
        target_path = self.clean_file_path(self.target_pbip_path.get())
        
        self.update_progress(30, "Reading source report data...")
        
        page_count = len(selected_page_names)
        total_pages_to_create = page_count * num_copies
        
        if num_copies > 1:
            self.log_message(f"   🔁 Creating {num_copies} copies of each page (Total: {total_pages_to_create} pages)")
        
        # Execute copies
        total_pages_created = 0
        for copy_num in range(1, num_copies + 1):
            if num_copies > 1:
                self.log_message(f"\n   📋 Creating copy batch {copy_num}/{num_copies}...")
            
            progress_base = 30 + (copy_num - 1) * (50 // num_copies)
            self.update_progress(progress_base, f"Copying pages (batch {copy_num}/{num_copies})...")
            
            success = self.advanced_copy_engine.copy_selected_pages_cross_pbip(
                source_path, target_path, selected_page_names, self.analysis_results
            )
            
            if success:
                total_pages_created += page_count
                self.log_message(f"   ✅ Batch {copy_num} complete ({page_count} pages copied)")
            else:
                self.log_message(f"   ⚠️ Batch {copy_num} failed")
        
        self.update_progress(90, "Finalizing target report updates...")
        
        self.update_progress(100, "Cross-PBIP copy complete!")
        return {
            'success': True,
            'source_path': source_path,
            'target_path': target_path,
            'page_count': total_pages_created,
            'num_copies': num_copies,
            'is_cross_pbip': True
        }
    
    def _copy_thread_target(self, selected_page_names, num_copies=1):
        """Background copy logic"""
        self.log_message("\n🚀 Starting page copy operation...")
        
        self.update_progress(10, "Preparing copy operation...")
        report_path = self.clean_file_path(self.report_path.get())
        
        self.update_progress(30, "Reading report data...")
        
        page_count = len(selected_page_names)
        total_pages_to_create = page_count * num_copies
        
        if num_copies > 1:
            self.log_message(f"   🔁 Creating {num_copies} copies of each page (Total: {total_pages_to_create} pages)")
        
        # Execute copies
        total_pages_created = 0
        for copy_num in range(1, num_copies + 1):
            if num_copies > 1:
                self.log_message(f"\n   📋 Creating copy batch {copy_num}/{num_copies}...")
            
            progress_base = 30 + (copy_num - 1) * (50 // num_copies)
            self.update_progress(progress_base, f"Copying pages (batch {copy_num}/{num_copies})...")
            
            success = self.advanced_copy_engine.copy_selected_pages(
                report_path, selected_page_names, self.analysis_results
            )
            
            if success:
                total_pages_created += page_count
                self.log_message(f"   ✅ Batch {copy_num} complete ({page_count} pages copied)")
            else:
                self.log_message(f"   ⚠️ Batch {copy_num} failed")
        
        self.update_progress(90, "Finalizing report updates...")
        
        self.update_progress(100, "Copy operation complete!")
        return {
            'success': True,
            'report_path': report_path,
            'page_count': total_pages_created,
            'num_copies': num_copies
        }
    
    def _handle_copy_complete(self, result):
        """Handle copy completion"""
        from pathlib import Path
        
        if result['success']:
            is_cross_pbip = result.get('is_cross_pbip', False)
            num_copies = result.get('num_copies', 1)
            
            if is_cross_pbip:
                self.log_message("✅ CROSS-PBIP PAGE COPY COMPLETED SUCCESSFULLY!")
                self.log_message(f"📂 Source: {result['source_path']}")
                self.log_message(f"🎯 Target: {result['target_path']}")
                
                if num_copies > 1:
                    msg = f"Cross-PBIP page copy completed successfully!\n\n" \
                          f"📋 Copied {result['page_count']} page(s) total ({num_copies} copies each)\n\n" \
                          f"📂 Source: {Path(result['source_path']).name}\n" \
                          f"🎯 Target: {Path(result['target_path']).name}\n\n" \
                          f"The copied pages have been added to the target report."
                else:
                    msg = f"Cross-PBIP page copy completed successfully!\n\n" \
                          f"📋 Copied {result['page_count']} page(s) with bookmarks\n\n" \
                          f"📂 Source: {Path(result['source_path']).name}\n" \
                          f"🎯 Target: {Path(result['target_path']).name}\n\n" \
                          f"The copied pages have been added to the target report."
                
                self.show_info("Cross-PBIP Copy Complete", msg)
            else:
                self.log_message("✅ PAGE COPY COMPLETED SUCCESSFULLY!")
                self.log_message(f"💾 Report updated: {result['report_path']}")
                
                if num_copies > 1:
                    msg = f"Page copy completed successfully!\n\n" \
                          f"📋 Copied {result['page_count']} page(s) total ({num_copies} copies each)\n" \
                          f"💾 Report: {Path(result['report_path']).name}\n\n" \
                          f"The copied pages have been added to your report with '(Copy)' suffix."
                else:
                    msg = f"Page copy completed successfully!\n\n" \
                          f"📋 Copied {result['page_count']} page(s) with bookmarks\n" \
                          f"💾 Report: {Path(result['report_path']).name}\n\n" \
                          f"The copied pages have been added to your report with '(Copy)' suffix."
                
                self.show_info("Copy Complete", msg)
        else:
            self.show_error("Copy Failed", "The page copy operation failed. Check the log for details.")
    
    def _start_bookmark_copy(self):
        """Start bookmark + visual copy operation"""
        from pathlib import Path
        
        # Validate selections
        if not self.source_page_var.get():
            self.show_error("Error", "Please select a source page")
            return
        
        if not self.bookmarks_treeview:
            self.show_error("Error", "Bookmark selection not available")
            return
        
        # Get selected tree items
        selected_items = self.bookmarks_treeview.selection()
        if not selected_items:
            self.show_error("Error", "Please select at least one bookmark to copy")
            return
        
        if not self.target_pages_listbox:
            self.show_error("Error", "Target page selection not available")
            return
        
        target_indices = self.target_pages_listbox.curselection()
        if not target_indices:
            self.show_error("Error", "Please select at least one target page")
            return
        
        # Get source page
        selected_display = self.source_page_var.get().split(' (')[0]
        source_page = next((p for p in self.available_pages 
                          if p['display_name'] == selected_display), None)
        
        if not source_page:
            self.show_error("Error", "Source page not found")
            return
        
        # Extract bookmark IDs from selected tree items (only bookmark items, not groups, and only selectable ones)
        selected_bookmark_ids = []
        for item_id in selected_items:
            if item_id in self._bookmark_tree_mapping:
                item_info = self._bookmark_tree_mapping[item_id]
                if item_info['type'] == 'bookmark' and item_info.get('selectable', True):
                    selected_bookmark_ids.append(item_info['bookmark_id'])
        
        if not selected_bookmark_ids:
            self.show_error("Error", "No bookmarks selected (groups are not copyable by themselves)")
            return
        
        # Get target page IDs (skip index 0 which is "Select All")
        # Determine which page list to use based on destination mode
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"
        
        if is_cross_pbip:
            # Use target report's pages
            if not hasattr(self, 'target_analysis_results') or not self.target_analysis_results:
                self.show_error("Error", "Target report not analyzed")
                return
            
            # Get all pages from target analysis
            target_pages_list = []
            target_pages_list.extend(self.target_analysis_results.get('pages_with_bookmarks', []))
            target_pages_list.extend(self.target_analysis_results.get('pages_without_bookmarks', []))
        else:
            # Use source report's pages (same PBIP)
            if not hasattr(self, 'all_report_pages') or not self.all_report_pages:
                self.show_error("Error", "Pages not available")
                return
            target_pages_list = self.all_report_pages
        
        target_page_names = []
        for idx in target_indices:
            if idx == 0:  # Skip "Select All" item
                continue
            target_display = self.target_pages_listbox.get(idx)
            # Look up page name from display name
            target_page = next((p for p in target_pages_list 
                              if p['display_name'] == target_display), None)
            if target_page:
                target_page_names.append(target_page['name'])
        
        if not target_page_names:
            self.show_error("Error", "No valid target pages selected")
            return
        
        # Determine if cross-PBIP or same-PBIP based on destination mode
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"
        
        # Confirm operation
        confirm_msg = f"Ready to copy bookmarks and visuals?\n\n"
        if is_cross_pbip:
            confirm_msg += f"📂 Source PBIP: {Path(self.report_path.get()).name}\n"
            confirm_msg += f"📂 Target PBIP: {Path(self.target_pbip_path.get()).name}\n"
        else:
            confirm_msg += f"📡 Report: {Path(self.report_path.get()).name}\n"
        
        confirm_msg += f"📄 Source Page: {source_page['display_name']}\n"
        confirm_msg += f"🔖 Bookmarks: {len(selected_bookmark_ids)}\n"
        confirm_msg += f"🎯 Target Pages: {len(target_page_names)}\n\n"
        confirm_msg += f"💡 Visuals keep same IDs, new bookmarks per page"
        
        if not self.ask_yes_no("Confirm Bookmark Copy", confirm_msg):
            return
        
        # Use base class background processing with appropriate target function
        if is_cross_pbip:
            self.run_in_background(
                target_func=lambda: self._cross_pbip_bookmark_copy_thread_target(
                    source_page['name'], selected_bookmark_ids, target_page_names
                ),
                success_callback=self._handle_bookmark_copy_complete,
                error_callback=lambda e: self.show_error("Cross-PBIP Copy Error", str(e))
            )
        else:
            self.run_in_background(
                target_func=lambda: self._bookmark_copy_thread_target(
                    source_page['name'], selected_bookmark_ids, target_page_names
                ),
                success_callback=self._handle_bookmark_copy_complete,
                error_callback=lambda e: self.show_error("Bookmark Copy Error", str(e))
            )
    
    def _bookmark_copy_thread_target(self, source_page_name, bookmark_ids, target_page_names):
        """Background bookmark copy logic"""
        self.log_message("\n🚀 Starting bookmark + visual copy operation...")
        
        self.update_progress(10, "Preparing bookmark copy operation...")
        report_path = self.clean_file_path(self.report_path.get())
        
        self.update_progress(30, "Reading bookmark and visual data...")
        
        self.update_progress(50, "Copying bookmarks and visuals (Per-Page mode)...")
        
        # Call core engine
        stats = self.advanced_copy_engine.copy_bookmarks_with_visuals(
            report_path=report_path,
            source_page_name=source_page_name,
            bookmark_names=bookmark_ids,
            target_page_names=target_page_names
        )
        
        self.update_progress(90, "Finalizing report updates...")
        
        self.update_progress(100, "Bookmark copy complete!")
        return {
            'success': True,
            'report_path': report_path,
            'stats': stats
        }
    
    def _cross_pbip_bookmark_copy_thread_target(self, source_page_name, bookmark_ids, target_page_names):
        """Background cross-PBIP bookmark copy logic"""
        self.log_message("\n🚀 Starting CROSS-PBIP bookmark + visual copy operation...")
        
        self.update_progress(10, "Preparing cross-PBIP copy operation...")
        source_pbip_path = self.clean_file_path(self.report_path.get())
        target_pbip_path = self.clean_file_path(self.target_pbip_path.get())
        
        self.update_progress(30, "Reading bookmark and visual data from source...")
        
        self.update_progress(50, "Copying bookmarks and visuals to target PBIP...")
        
        # Call cross-PBIP engine method
        stats = self.advanced_copy_engine.copy_bookmarks_with_visuals_cross_pbip(
            source_pbip_path=source_pbip_path,
            target_pbip_path=target_pbip_path,
            source_page_name=source_page_name,
            bookmark_names=bookmark_ids,
            target_page_names=target_page_names
        )
        
        self.update_progress(90, "Finalizing target report updates...")
        
        self.update_progress(100, "Cross-PBIP copy complete!")
        return {
            'success': True,
            'source_path': source_pbip_path,
            'target_path': target_pbip_path,
            'stats': stats,
            'is_cross_pbip': True
        }
    
    def _handle_bookmark_copy_complete(self, result):
        """Handle bookmark copy completion"""
        from pathlib import Path
        
        if result['success']:
            stats = result['stats']
            is_cross_pbip = result.get('is_cross_pbip', False)
            
            if is_cross_pbip:
                self.log_message("✅ CROSS-PBIP BOOKMARK + VISUAL COPY COMPLETED SUCCESSFULLY!")
                self.log_message(f"📂 Source: {result['source_path']}")
                self.log_message(f"📂 Target: {result['target_path']}")
                
                self.show_info("Cross-PBIP Copy Complete",
                              f"Cross-PBIP copy completed successfully!\n\n"
                              f"🔖 Bookmarks copied: {stats['bookmarks_copied']}\n"
                              f"📊 Visuals copied: {stats['visuals_copied']}\n"
                              f"📄 Pages updated: {stats['pages_updated']}\n\n"
                              f"💡 Visuals keep same IDs, new bookmarks per page\n\n"
                              f"📂 Source: {Path(result['source_path']).name}\n"
                              f"📂 Target: {Path(result['target_path']).name}")
            else:
                self.log_message("✅ BOOKMARK + VISUAL COPY COMPLETED SUCCESSFULLY!")
                self.log_message(f"📡 Report updated: {result['report_path']}")
                
                self.show_info("Bookmark Copy Complete",
                              f"Bookmark copy completed successfully!\n\n"
                              f"🔖 Bookmarks copied: {stats['bookmarks_copied']}\n"
                              f"📊 Visuals copied: {stats['visuals_copied']}\n"
                              f"📄 Pages updated: {stats['pages_updated']}\n\n"
                              f"💡 Visuals keep same IDs, new bookmarks per page\n\n"
                              f"📡 Report: {Path(result['report_path']).name}")
        else:
            self.show_error("Bookmark Copy Failed",
                          "The bookmark copy operation failed. Check the log for details.")
    
    # NOTE: _expand_window_after_analysis() method removed
    # Height is now managed entirely by _adjust_window_height() in ui_page_selection.py
    # which sets FIXED heights based on state (not additive)
