"""
Helpers Mixin for Advanced Copy UI
Built by Reid Havens of Analytic Endeavors

Contains helper methods for messages and dialogs.
"""

import tkinter as tk
from tkinter import ttk
from pathlib import Path

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

    @property
    def _tool_icon_path(self) -> str:
        """Get the Advanced Copy tool icon path for dialogs"""
        return str(Path(__file__).parent.parent.parent.parent / "assets" / "Tool Icons" / "advanced copy.svg")

    def _show_welcome_message(self):
        """Show welcome message in log panel"""
        self.log_message("📋 Welcome to Advanced Copy!")
        self.log_message("=" * 60)
        self.log_message("This tool provides advanced page and visual copying:")
        self.log_message("• Full Page Copy: Duplicates entire pages with bookmarks")
        self.log_message("• Bookmark+Visual Copy: Pop out bookmarks to multiple pages")
        self.log_message("")
        self.log_message("📂 Start by selecting a report file and choosing a copy mode")
        self.log_message("⚠️ Note: Requires PBIP format files only")

        # Reset summary table to placeholder values
        if hasattr(self, '_summary_title_label') and self._summary_title_label:
            self._summary_title_label.config(text="Select a report to analyze")

        if hasattr(self, '_summary_labels') and self._summary_labels:
            for key in self._summary_labels:
                self._summary_labels[key].config(text="-")

        # Reset status label
        if hasattr(self, '_summary_status_label') and self._summary_status_label:
            colors = self._theme_manager.colors
            self._summary_status_label.config(
                text="Waiting for analysis...",
                fg=colors['text_secondary'])
    
    def _show_analysis_summary(self, results):
        """Show analysis summary in both log and summary table"""
        self.log_message("\n ANALYSIS SUMMARY")
        self.log_message("=" * 50)

        report = results['report']
        copyable_pages = len(results['pages_with_bookmarks'])
        total_pages = results['analysis_summary']['total_pages']
        total_bookmarks = results['analysis_summary']['total_bookmarks']

        # Get navigator count
        navigator_analysis = results.get('navigator_analysis', {})
        total_navigators = navigator_analysis.get('total_navigators', 0) if navigator_analysis else 0

        # Update summary table
        if hasattr(self, '_summary_title_label') and self._summary_title_label:
            self._summary_title_label.config(text=f"Report: {report['name']}")

        if hasattr(self, '_summary_labels') and self._summary_labels:
            if 'Total Pages' in self._summary_labels:
                self._summary_labels['Total Pages'].config(text=str(total_pages))
            if 'Pages with Bookmarks' in self._summary_labels:
                self._summary_labels['Pages with Bookmarks'].config(text=str(copyable_pages))
            if 'Total Bookmarks' in self._summary_labels:
                self._summary_labels['Total Bookmarks'].config(text=str(total_bookmarks))
            if 'Navigators' in self._summary_labels:
                self._summary_labels['Navigators'].config(text=str(total_navigators))

            # Pages without Bookmarks
            if 'Pages without Bookmarks' in self._summary_labels:
                pages_without = len(results.get('pages_without_bookmarks', []))
                self._summary_labels['Pages without Bookmarks'].config(text=str(pages_without))

            # Page Filters - sum filter_count from all pages
            if 'Page Filters' in self._summary_labels:
                total_filters = sum(p.get('filter_count', 0) for p in results.get('pages_with_bookmarks', [])) + \
                                sum(p.get('filter_count', 0) for p in results.get('pages_without_bookmarks', []))
                self._summary_labels['Page Filters'].config(text=str(total_filters))

            # Report-Level Bookmarks
            if 'Report-Level Bookmarks' in self._summary_labels:
                report_bookmarks = len(results.get('report_bookmarks', []))
                self._summary_labels['Report-Level Bookmarks'].config(text=str(report_bookmarks))

            # Navigator Groups
            if 'Navigator Groups' in self._summary_labels:
                navigators_with_groups = navigator_analysis.get('navigators_with_groups', 0) if navigator_analysis else 0
                if total_navigators > 0:
                    self._summary_labels['Navigator Groups'].config(text=f"{navigators_with_groups} of {total_navigators}")
                else:
                    self._summary_labels['Navigator Groups'].config(text="-")

        # Update status label to show analysis complete
        if hasattr(self, '_summary_status_label') and self._summary_status_label:
            colors = self._theme_manager.colors
            self._summary_status_label.config(
                text="Ready to select pages",
                fg=colors.get('success', '#22c55e'))

        # Log to progress log panel
        self.log_message(f"Report: {report['name']}")
        self.log_message(f"Total Pages: {total_pages}")
        self.log_message(f"Pages with Bookmarks: {copyable_pages}")
        self.log_message(f"Total Bookmarks: {total_bookmarks}")
        
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
        from core.ui_base import RoundedButton

        def create_help_content(help_window, colors, help_bg):
            # Main container - use tk.Frame with explicit bg for consistency
            container = tk.Frame(help_window, bg=help_bg)
            container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

            # Content frame for scrollable content - use tk.Frame with explicit bg
            content_frame = tk.Frame(container, bg=help_bg)
            content_frame.pack(fill=tk.BOTH, expand=True)

            # Header - centered for middle-out design
            tk.Label(content_frame, text="Advanced Copy - Help",
                     font=('Segoe UI', 16, 'bold'),
                     bg=help_bg,
                     fg=colors['title_color']).pack(anchor=tk.CENTER, pady=(0, 20))

            # Orange warning box with white text
            warning_frame = tk.Frame(content_frame, bg=help_bg)
            warning_frame.pack(fill=tk.X, pady=(0, 20))

            warning_bg = '#d97706'
            warning_text = '#ffffff'
            warning_container = tk.Frame(warning_frame, bg=warning_bg,
                                       padx=15, pady=10, relief='flat', borderwidth=0)
            warning_container.pack(fill=tk.X)

            tk.Label(warning_container, text="IMPORTANT DISCLAIMERS & REQUIREMENTS",
                     font=('Segoe UI', 12, 'bold'),
                     bg=warning_bg,
                     fg=warning_text).pack(anchor=tk.W)

            warnings = [
                "This tool ONLY works with PBIP enhanced report format (PBIR) files",
                "This is NOT officially supported by Microsoft - use at your own discretion",
                "Look for .pbip files with definition\\ folder (not report.json files)",
                "Always keep backups of your original reports before copying pages",
                "Test thoroughly and validate copied pages before production use",
                "Enable 'Store reports using enhanced metadata format (PBIR)' in Power BI Desktop"
            ]

            for warning in warnings:
                tk.Label(warning_container, text=warning, font=('Segoe UI', 10),
                         bg=warning_bg,
                         fg=warning_text).pack(anchor=tk.W, pady=1)

            # Two-column layout for help sections - use tk.Frame with explicit bg
            columns_frame = tk.Frame(content_frame, bg=help_bg)
            columns_frame.pack(fill=tk.BOTH, expand=True)
            columns_frame.columnconfigure(0, weight=1)
            columns_frame.columnconfigure(1, weight=1)

            # LEFT COLUMN
            left_column = tk.Frame(columns_frame, bg=help_bg)
            left_column.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))

            # RIGHT COLUMN
            right_column = tk.Frame(columns_frame, bg=help_bg)
            right_column.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))

            # Left column sections (emojis removed)
            left_sections = [
                ("What This Tool Does", [
                    "Offers TWO copy modes: Full Page Copy and Bookmark+Visual Copy",
                    "Duplicates complete pages OR specific bookmarks with visuals",
                    "Preserves all bookmarks and their visual relationships",
                    "Automatically renames copies to avoid conflicts",
                    "Updates report metadata after copying",
                    "Supports group selection - click groups to select all children"
                ]),
                ("Full Page Copy Mode", [
                    "Duplicates entire pages with all visuals and bookmarks",
                    "Creates complete copies including all page elements",
                    "Useful for creating template variations or backup pages",
                    "Copied pages get '(Copy)' suffix to avoid name conflicts"
                ]),
                ("Bookmark + Visual Copy Mode", [
                    "Copies specific bookmarks and their visuals to existing pages",
                    "'Pop out' bookmark controls to multiple pages",
                    "Visuals keep SAME IDs across pages (shared visuals)",
                    "Creates NEW bookmarks for each target page",
                    "Perfect for dashboard navigation across multiple pages",
                    "Only copies bookmark navigators relevant to selected bookmarks",
                    "ONLY 'Selected visuals' configured bookmarks are supported",
                    "Bookmarks set to 'All visuals' cannot be copied"
                ])
            ]

            # Right column sections (emojis removed)
            right_sections = [
                ("File Requirements", [
                    "Only .pbip files (enhanced PBIR format) are supported",
                    "Full Page Mode: Only pages with bookmarks shown for copying",
                    "Bookmark Mode: Can target any page (with or without bookmarks)",
                    "Bookmark Mode: Requires 'Selected visuals' configured bookmarks",
                    "Standard .pbix files are NOT supported"
                ]),
                ("Important Notes", [
                    "Full Page: Creates complete duplicate pages",
                    "Bookmark Mode: Shares visuals across pages (same IDs)",
                    "Bookmark Mode: Only 'Selected visuals' bookmarks are supported",
                    "'All visuals' bookmarks cannot be copied (capture entire page state)",
                    "Disabled bookmarks appear grayed out in the UI",
                    "Bookmark groups are preserved and duplicated appropriately",
                    "Always backup your report before making changes",
                    "NOT officially supported by Microsoft"
                ]),
                ("Configuring Bookmarks for Copy", [
                    "To make bookmarks copyable in Bookmark+Visual mode:",
                    "1. Open bookmark in Power BI Desktop",
                    "2. Uncheck 'All visuals' option",
                    "3. Select specific visuals to include",
                    "4. Save the bookmark",
                    "5. Bookmark will now show as copyable in the tool"
                ])
            ]
            
            # Render left column sections - use tk.Frame and tk.Label with explicit bg
            for title, items in left_sections:
                section_frame = tk.Frame(left_column, bg=help_bg)
                section_frame.pack(fill=tk.X, pady=(0, 15))

                tk.Label(section_frame, text=title, font=('Segoe UI', 12, 'bold'),
                         bg=help_bg, fg=colors['title_color']).pack(anchor=tk.W)

                for item in items:
                    tk.Label(section_frame, text=item, font=('Segoe UI', 10),
                             bg=help_bg, fg=colors['text_primary'],
                             wraplength=380, justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)

            # Render right column sections - use tk.Frame and tk.Label with explicit bg
            for title, items in right_sections:
                section_frame = tk.Frame(right_column, bg=help_bg)
                section_frame.pack(fill=tk.X, pady=(0, 15))

                tk.Label(section_frame, text=title, font=('Segoe UI', 12, 'bold'),
                         bg=help_bg, fg=colors['title_color']).pack(anchor=tk.W)

                for item in items:
                    tk.Label(section_frame, text=item, font=('Segoe UI', 10),
                             bg=help_bg, fg=colors['text_primary'],
                             wraplength=380, justify=tk.LEFT).pack(anchor=tk.W, padx=(12, 0), pady=1)
        
        # Get theme colors
        colors = self._theme_manager.colors
        is_dark = self._theme_manager.is_dark

        # Consistent help dialog background for all tools
        help_bg = colors['background']

        # Create custom help window for Advanced Copy (independent of base class)
        help_window = tk.Toplevel(self.main_app.root)
        help_window.withdraw()  # Hide until fully styled (prevents white flash)
        help_window.title("Advanced Copy - Help")
        help_window.geometry("950x965")
        help_window.resizable(False, False)
        help_window.transient(self.main_app.root)
        help_window.grab_set()
        help_window.configure(bg=help_bg)

        # Set AE favicon icon
        try:
            help_window.iconbitmap(self.main_app.config.icon_path)
        except Exception:
            pass

        # Create content
        create_help_content(help_window, colors, help_bg)

        # Bind escape key
        help_window.bind('<Escape>', lambda event: help_window.destroy())

        # Center dialog and show (after all content built to prevent flash)
        help_window.update_idletasks()
        parent = self.main_app.root
        x = parent.winfo_rootx() + (parent.winfo_width() - 950) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 965) // 2
        help_window.geometry(f"950x965+{x}+{y}")

        # Set dark title bar BEFORE showing to prevent white flash
        help_window.update()
        if hasattr(self, '_set_dialog_title_bar_color'):
            self._set_dialog_title_bar_color(help_window, self._theme_manager.is_dark)
        help_window.deiconify()  # Show now that it's styled
    
    def analyze_report(self):
        """Analyze selected report"""
        try:
            report_path = self.clean_file_path(self.report_path.get())
            # Basic validation happens in the engine
        except Exception as e:
            self.show_error("Validation Error", str(e), icon_path=self._tool_icon_path)
            return
        
        # Use base class background processing
        self.run_in_background(
            target_func=self._analyze_thread_target,
            success_callback=self._handle_analysis_complete,
            error_callback=lambda e: self.show_error("Analysis Error", str(e), icon_path=self._tool_icon_path)
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
                             "This tool only copies pages that have associated bookmarks.",
                             icon_path=self._tool_icon_path)
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
                error_callback=lambda e: self.show_error("Target Analysis Error", str(e), icon_path=self._tool_icon_path)
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
            self.show_error("Error", "Please analyze report first", icon_path=self._tool_icon_path)
            return

        if self.copy_content_mode.get() == "full_page":
            self._start_full_page_copy()
        else:
            self._start_bookmark_copy()
    
    def _start_full_page_copy(self):
        """Start full page copy operation"""
        from pathlib import Path
        
        if not self.analysis_results:
            self.show_error("Error", "Please analyze report first", icon_path=self._tool_icon_path)
            return

        # Check if cross-PBIP - now supported!
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"

        if not self.pages_listbox:
            self.show_error("Error", "No pages available for selection", icon_path=self._tool_icon_path)
            return

        # Get selected pages using new method (compatible with Treeview-based selection)
        selected_indices = self.get_selected_page_indices()
        if not selected_indices:
            self.show_error("Error", "Please select at least one page to copy", icon_path=self._tool_icon_path)
            return

        selected_pages = [self.available_pages[i] for i in selected_indices]
        page_names = [page['name'] for page in selected_pages]

        # Get number of copies
        num_copies = self.num_copies.get()
        total_pages_to_create = len(selected_pages) * num_copies

        # Validate target path if cross-PBIP
        if is_cross_pbip and not self.target_pbip_path.get():
            self.show_error("Error", "Please select a target PBIP file for cross-PBIP copy", icon_path=self._tool_icon_path)
            return
        
        # Check for navigator warnings
        warning_section = ""
        if self.analysis_results:
            nav_analysis = self.analysis_results.get('navigator_analysis', {})
            if nav_analysis and nav_analysis['navigators_without_groups'] > 0:
                warning_section = (
                    f"\n\nWARNING: Navigator Without Group Detected\n\n"
                    f"This page has navigators showing ALL bookmarks.\n"
                    f"After copying, original and copied bookmarks will appear together.\n\n"
                    f"RECOMMENDED: Cancel and assign groups to navigators first."
                )

        # Build confirmation message based on mode (cleaner formatting)
        pages_list = "\n".join([f"  - {p['display_name']} ({p['bookmark_count']} bookmarks)"
                               for p in selected_pages[:5]])
        if len(selected_pages) > 5:
            pages_list += f"\n  ... and {len(selected_pages)-5} more"

        if is_cross_pbip:
            confirm_msg = f"Ready to copy {len(selected_pages)} page(s)?"
            confirm_msg += warning_section
            confirm_msg += f"\n\nPages to copy:\n{pages_list}"
            confirm_msg += f"\n\nReport: {Path(self.target_pbip_path.get()).name}"
        else:
            confirm_msg = f"Ready to copy {len(selected_pages)} page(s)?"
            confirm_msg += warning_section
            confirm_msg += f"\n\nPages to copy:\n{pages_list}"
            confirm_msg += f"\n\nReport: {Path(self.report_path.get()).name}"

        # Confirm operation
        if not self.ask_yes_no("Confirm Copy", confirm_msg, icon_path=self._tool_icon_path):
            return

        # Use base class background processing - route based on destination mode
        if is_cross_pbip:
            self.run_in_background(
                target_func=lambda: self._copy_cross_pbip_thread_target(page_names, num_copies),
                success_callback=self._handle_copy_complete,
                error_callback=lambda e: self.show_error("Cross-PBIP Copy Error", str(e), icon_path=self._tool_icon_path)
            )
        else:
            self.run_in_background(
                target_func=lambda: self._copy_thread_target(page_names, num_copies),
                success_callback=self._handle_copy_complete,
                error_callback=lambda e: self.show_error("Copy Error", str(e), icon_path=self._tool_icon_path)
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
                
                self.show_info("Cross-PBIP Copy Complete", msg, icon_path=self._tool_icon_path)
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

                self.show_info("Copy Complete", msg, icon_path=self._tool_icon_path)
        else:
            self.show_error("Copy Failed", "The page copy operation failed. Check the log for details.", icon_path=self._tool_icon_path)
    
    def _start_bookmark_copy(self):
        """Start bookmark + visual copy operation"""
        from pathlib import Path
        
        # Validate selections
        if not self.source_page_var.get():
            self.show_error("Error", "Please select a source page", icon_path=self._tool_icon_path)
            return

        if not self.bookmarks_treeview:
            self.show_error("Error", "Bookmark selection not available", icon_path=self._tool_icon_path)
            return
        
        # Get selected tree items
        selected_items = self.bookmarks_treeview.selection()
        if not selected_items:
            self.show_error("Error", "Please select at least one bookmark to copy", icon_path=self._tool_icon_path)
            return

        if not hasattr(self, '_selected_target_pages') or not hasattr(self, '_target_pages_data'):
            self.show_error("Error", "Target page selection not available", icon_path=self._tool_icon_path)
            return

        if not self._selected_target_pages:
            self.show_error("Error", "Please select at least one target page", icon_path=self._tool_icon_path)
            return

        # Get source page
        selected_display = self.source_page_var.get().split(' (')[0]
        source_page = next((p for p in self.available_pages
                          if p['display_name'] == selected_display), None)

        if not source_page:
            self.show_error("Error", "Source page not found", icon_path=self._tool_icon_path)
            return

        # Extract bookmark IDs from selected tree items (only bookmark items, not groups, and only selectable ones)
        selected_bookmark_ids = []
        for item_id in selected_items:
            if item_id in self._bookmark_tree_mapping:
                item_info = self._bookmark_tree_mapping[item_id]
                if item_info['type'] == 'bookmark' and item_info.get('selectable', True):
                    selected_bookmark_ids.append(item_info['bookmark_id'])

        if not selected_bookmark_ids:
            self.show_error("Error", "No bookmarks selected (groups are not copyable by themselves)", icon_path=self._tool_icon_path)
            return

        # Get target page IDs (skip index 0 which is "Select All")
        # Determine which page list to use based on destination mode
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"

        if is_cross_pbip:
            # Use target report's pages
            if not hasattr(self, 'target_analysis_results') or not self.target_analysis_results:
                self.show_error("Error", "Target report not analyzed", icon_path=self._tool_icon_path)
                return

            # Get all pages from target analysis
            target_pages_list = []
            target_pages_list.extend(self.target_analysis_results.get('pages_with_bookmarks', []))
            target_pages_list.extend(self.target_analysis_results.get('pages_without_bookmarks', []))
        else:
            # Use source report's pages (same PBIP)
            if not hasattr(self, 'all_report_pages') or not self.all_report_pages:
                self.show_error("Error", "Pages not available", icon_path=self._tool_icon_path)
                return
            target_pages_list = self.all_report_pages

        target_page_names = []
        for idx in self._selected_target_pages:
            # Get page data from the stored list (idx is already 0-based into _target_pages_data)
            if idx < len(self._target_pages_data):
                page_data = self._target_pages_data[idx]
                # Look up page name from display name
                target_page = next((p for p in target_pages_list
                                  if p['display_name'] == page_data['display_name']), None)
                if target_page:
                    target_page_names.append(target_page['name'])

        if not target_page_names:
            self.show_error("Error", "No valid target pages selected", icon_path=self._tool_icon_path)
            return
        
        # Determine if cross-PBIP or same-PBIP based on destination mode
        is_cross_pbip = self.copy_destination_mode.get() == "cross_pbip"

        # Get the selected copy mode (perpage or crosspage)
        copy_mode = self.bookmark_copy_mode.get()

        # For cross-page mode, check if bookmarks need "Current Page" disabled
        if copy_mode == 'crosspage':
            report_path = self.clean_file_path(self.report_path.get())
            report_dir = Path(report_path).parent / f"{Path(report_path).stem}.Report"
            bookmarks_dir = report_dir / "definition" / "bookmarks"

            self.log_message("\n🔍 Checking cross-page compatibility...")
            compatibility = self.advanced_copy_engine.bookmark_analyzer.check_bookmarks_for_cross_page_compatibility(
                bookmarks_dir, selected_bookmark_ids
            )

            if not compatibility['all_compatible']:
                # Some bookmarks have "Current Page" enabled - warn user
                needs_mod = compatibility['needs_modification']
                details = compatibility['details']

                warning_msg = (
                    f"⚠️ CROSS-PAGE COMPATIBILITY WARNING\n\n"
                    f"{len(needs_mod)} bookmark(s) have 'Current Page' enabled:\n\n"
                )

                for bm_id in needs_mod[:5]:  # Show max 5
                    bm_detail = details.get(bm_id, {})
                    warning_msg += f"  • {bm_detail.get('display_name', bm_id)}\n"

                if len(needs_mod) > 5:
                    warning_msg += f"  ... and {len(needs_mod) - 5} more\n"

                warning_msg += (
                    f"\nFor cross-page mode to work, 'Current Page' must be disabled.\n"
                    f"This allows the bookmark to apply on ANY page, not just navigate to one.\n\n"
                    f"Do you want to proceed? The tool will automatically disable\n"
                    f"'Current Page' for these bookmarks."
                )

                if not self.ask_yes_no("Cross-Page Compatibility", warning_msg, icon_path=self._tool_icon_path):
                    self.log_message("   ❌ Cross-page copy cancelled by user")
                    return

                self.log_message("   ✅ User confirmed - will disable 'Current Page' for affected bookmarks")

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

        # Mode-specific description
        if copy_mode == 'crosspage':
            confirm_msg += f"📋 Mode: Cross-Page\n"
            confirm_msg += f"💡 Bookmarks will work on source + {len(target_page_names)} target page(s)"
        else:
            confirm_msg += f"📋 Mode: Per-Page\n"
            confirm_msg += f"💡 New bookmarks created for each target page"

        if not self.ask_yes_no("Confirm Bookmark Copy", confirm_msg, icon_path=self._tool_icon_path):
            return

        # Use base class background processing with appropriate target function
        if is_cross_pbip:
            self.run_in_background(
                target_func=lambda: self._cross_pbip_bookmark_copy_thread_target(
                    source_page['name'], selected_bookmark_ids, target_page_names, copy_mode
                ),
                success_callback=self._handle_bookmark_copy_complete,
                error_callback=lambda e: self.show_error("Cross-PBIP Copy Error", str(e), icon_path=self._tool_icon_path)
            )
        else:
            self.run_in_background(
                target_func=lambda: self._bookmark_copy_thread_target(
                    source_page['name'], selected_bookmark_ids, target_page_names, copy_mode
                ),
                success_callback=self._handle_bookmark_copy_complete,
                error_callback=lambda e: self.show_error("Bookmark Copy Error", str(e), icon_path=self._tool_icon_path)
            )
    
    def _bookmark_copy_thread_target(self, source_page_name, bookmark_ids, target_page_names, mode='perpage'):
        """Background bookmark copy logic"""
        self.log_message("\n🚀 Starting bookmark + visual copy operation...")

        self.update_progress(10, "Preparing bookmark copy operation...")
        report_path = self.clean_file_path(self.report_path.get())

        self.update_progress(30, "Reading bookmark and visual data...")

        # Update progress message based on mode
        if mode == 'crosspage':
            self.update_progress(50, "Copying visuals and configuring cross-page bookmarks...")
        else:
            self.update_progress(50, "Copying bookmarks and visuals (Per-Page mode)...")

        # Call core engine with mode
        stats = self.advanced_copy_engine.copy_bookmarks_with_visuals(
            report_path=report_path,
            source_page_name=source_page_name,
            bookmark_names=bookmark_ids,
            target_page_names=target_page_names,
            mode=mode
        )
        
        self.update_progress(90, "Finalizing report updates...")

        self.update_progress(100, "Bookmark copy complete!")
        return {
            'success': True,
            'report_path': report_path,
            'stats': stats,
            'mode': mode
        }
    
    def _cross_pbip_bookmark_copy_thread_target(self, source_page_name, bookmark_ids, target_page_names, mode='perpage'):
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
            target_page_names=target_page_names,
            bookmark_copy_mode=mode
        )
        
        self.update_progress(90, "Finalizing target report updates...")
        
        self.update_progress(100, "Cross-PBIP copy complete!")
        return {
            'success': True,
            'source_path': source_pbip_path,
            'target_path': target_pbip_path,
            'stats': stats,
            'is_cross_pbip': True,
            'mode': mode
        }
    
    def _handle_bookmark_copy_complete(self, result):
        """Handle bookmark copy completion"""
        from pathlib import Path

        if result['success']:
            stats = result['stats']
            is_cross_pbip = result.get('is_cross_pbip', False)
            copy_mode = result.get('mode', 'perpage')

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
                              f"📂 Target: {Path(result['target_path']).name}",
                              icon_path=self._tool_icon_path)
            elif copy_mode == 'crosspage':
                # Cross-page mode (same PBIP)
                self.log_message("✅ CROSS-PAGE BOOKMARK CONFIGURATION COMPLETED SUCCESSFULLY!")
                self.log_message(f"📡 Report updated: {result['report_path']}")

                bookmarks_configured = stats.get('bookmarks_configured', 0)
                total_pages = stats['pages_updated'] + 1  # target pages + source page

                self.show_info("Cross-Page Copy Complete",
                              f"Cross-page configuration completed successfully!\n\n"
                              f"🔖 Bookmarks configured: {bookmarks_configured}\n"
                              f"📊 Visuals copied: {stats['visuals_copied']}\n"
                              f"📄 Pages included: {total_pages} (source + targets)\n\n"
                              f"💡 Bookmarks now work across all selected pages\n\n"
                              f"📡 Report: {Path(result['report_path']).name}",
                              icon_path=self._tool_icon_path)
            else:
                # Per-page mode (same PBIP)
                self.log_message("✅ BOOKMARK + VISUAL COPY COMPLETED SUCCESSFULLY!")
                self.log_message(f"📡 Report updated: {result['report_path']}")

                self.show_info("Bookmark Copy Complete",
                              f"Bookmark copy completed successfully!\n\n"
                              f"🔖 Bookmarks copied: {stats['bookmarks_copied']}\n"
                              f"📊 Visuals copied: {stats['visuals_copied']}\n"
                              f"📄 Pages updated: {stats['pages_updated']}\n\n"
                              f"💡 Visuals keep same IDs, new bookmarks per page\n\n"
                              f"📡 Report: {Path(result['report_path']).name}",
                              icon_path=self._tool_icon_path)
        else:
            self.show_error("Bookmark Copy Failed",
                          "The bookmark copy operation failed. Check the log for details.",
                          icon_path=self._tool_icon_path)
    
    # NOTE: _expand_window_after_analysis() method removed
    # Height is now managed entirely by _adjust_window_height() in ui_page_selection.py
    # which sets FIXED heights based on state (not additive)
