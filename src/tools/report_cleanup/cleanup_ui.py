"""
Report Cleanup UI - User interface for report cleanup operations
Built by Reid Havens of Analytic Endeavors

This module provides the user interface for the Report Cleanup tool,
following the established patterns from other tools in the suite.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, List, Any, Optional

from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin
from core.constants import AppConstants
from tools.report_cleanup.shared_types import CleanupOpportunity, RemovalResult
from tools.report_cleanup.report_analyzer import ReportAnalyzer
from tools.report_cleanup.cleanup_engine import ReportCleanupEngine


class ReportCleanupTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """
    Report Cleanup UI Tab - provides interface for cleaning up Power BI reports
    """
    
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "report_cleanup", "Report Cleanup")
        
        # Tool instances
        self.analyzer = ReportAnalyzer(logger_callback=self.log_message)
        self.cleanup_engine = ReportCleanupEngine(logger_callback=self.log_message)
        
        # UI state
        self.pbip_path_var = tk.StringVar()
        self.current_opportunities: List[CleanupOpportunity] = []
        self.cleanup_options = {
            'remove_themes': tk.BooleanVar(value=True),
            'remove_custom_visuals': tk.BooleanVar(value=True),
            'remove_bookmarks': tk.BooleanVar(value=True),
            'hide_visual_filters': tk.BooleanVar(value=True)
        }
        
        # UI components
        self.file_input_components = None
        self.analyze_button = None
        self.results_frame = None
        self.opportunities_listbox = None
        self.options_frame = None
        self.action_buttons = None
        self.progress_components = None
        
        # Setup UI and show welcome message
        self.setup_ui()
        self._show_welcome_message()
    
    def setup_ui(self) -> None:
        """Setup the Report Cleanup UI"""
        # File input section
        self._setup_file_input_section()
        
        # Analysis section
        self._setup_analysis_section()
        
        # Results section
        self._setup_results_section()
        
        # Progress section
        self._setup_progress_section()
        
        # Log section
        self._setup_log_section()
        
        # Setup path cleaning
        self.setup_path_cleaning(self.pbip_path_var)
    
    def _setup_file_input_section(self):
        """Setup the file input section"""
        guide_text = [
            "🎯 Quick Start Guide:",
            "• Select your .pbip file using the Browse button",
            "• Click 'ANALYZE REPORT' to scan for cleanup opportunities",
            "• Choose what to remove: themes, custom visuals, or both",
            "• Click 'REMOVE ITEMS' to clean up your report"
        ]
        
        self.file_input_components = self.create_file_input_section(
            self.frame,
            "📁 PBIP FILE SELECTION", 
            [("PBIP Files", "*.pbip"), ("All Files", "*.*")],
            guide_text
        )
        self.file_input_components['frame'].pack(fill=tk.X, pady=(0, 20))
        
        # Store the path variable
        self.pbip_path_var = self.file_input_components['path_var']
    
    def _setup_analysis_section(self):
        """Setup the analysis section with side-by-side layout"""
        analysis_frame = ttk.LabelFrame(self.frame, text="🔍 REPORT ANALYSIS", 
                                      style='Section.TLabelframe', padding="15")
        analysis_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Main container with left and right sections
        main_container = ttk.Frame(analysis_frame)
        main_container.pack(fill=tk.X)
        
        # Left section - Analysis description and button
        left_frame = ttk.Frame(main_container)
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 20))
        
        ttk.Label(left_frame, text="Report Cleanup Analysis", 
                 font=('Segoe UI', 11, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
        
        desc_text = ("Analyzes your PBIP report to identify unused themes and custom visuals. "
                    "Detects themes that aren't active and custom visuals that aren't used in any pages, "
                    "including hidden visuals taking up space in the CustomVisuals directory.")
        ttk.Label(left_frame, text=desc_text, font=('Segoe UI', 9),
                 foreground=AppConstants.COLORS['text_secondary'],
                 wraplength=500).pack(anchor=tk.W, pady=(5, 10))
        
        # Analysis button
        self.analyze_button = ttk.Button(
            left_frame, 
            text="🔍 ANALYZE REPORT",
            command=self._analyze_report,
            style='Action.TButton'
        )
        self.analyze_button.pack(anchor=tk.W)
        
        # Right section - Analysis summary (initially hidden)
        self.analysis_summary_frame = ttk.Frame(main_container)
        self.analysis_summary_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(20, 0))
        # Initially hidden - will be shown after analysis
        self.analysis_summary_frame.pack_forget()
    
    def _setup_results_section(self):
        """Setup the cleanup opportunities section"""
        self.results_frame = ttk.LabelFrame(self.frame, text="🎯 CLEANUP OPPORTUNITIES", 
                                          style='Section.TLabelframe', padding="15")
        self.results_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Initially hidden - will be shown after analysis
        self.results_frame.pack_forget()
        
        # Results will be dynamically created in _show_analysis_results()
    
    def _setup_progress_section(self):
        """Setup the progress section"""
        self.progress_components = self.create_progress_bar(self.frame)
        # Initially hidden
        self.progress_components['frame'].pack_forget()
    
    def _position_progress_frame(self):
        """Position the progress frame appropriately for this layout"""
        if self.progress_components and self.progress_components['frame']:
            # Position above the log section
            self.progress_components['frame'].pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10), before=self.log_section_frame)
    
    def _setup_log_section(self):
        """Setup the log section with cleanup options on the right"""
        log_components = self.create_log_section_with_options(self.frame, "📊 ANALYSIS & CLEANUP LOG")
        self.log_section_frame = log_components['frame']
        self.log_section_frame.pack(fill=tk.BOTH, expand=True)
        
        # Store reference to options frame for later use
        self.log_options_frame = log_components['options_frame']
    
    def _analyze_report(self):
        """Analyze the PBIP report for cleanup opportunities"""
        try:
            # Validate input
            pbip_path = self.pbip_path_var.get().strip()
            if not pbip_path:
                self.show_error("Input Required", "Please select a PBIP file first.")
                return
            
            # Validate file
            self.validate_pbip_file(pbip_path, "PBIP file")
            
            # Clear previous results
            if hasattr(self, 'results_frame') and self.results_frame:
                self.results_frame.pack_forget()
            
            # Show immediate feedback
            self.log_message("🚀 Starting report analysis...")
            self.log_message(f"📁 Analyzing: {pbip_path}")
            
            # Run analysis in background
            def run_analysis():
                analyzer = ReportAnalyzer(logger_callback=self.log_message)
                return analyzer.analyze_pbip_report(pbip_path)
            
            def on_success(result):
                analysis_data, opportunities = result
                self.current_opportunities = opportunities
                self._show_analysis_results(analysis_data, opportunities)
                self.log_message("✅ Analysis completed successfully!")
            
            def on_error(error):
                self.log_message(f"❌ Analysis failed: {error}")
                self.show_error("Analysis Failed", f"Report analysis failed:\n\n{error}")
            
            # Show progress and run in background
            self.update_progress(10, "Starting analysis...", True)
            self.run_in_background(run_analysis, on_success, on_error)
            
        except Exception as e:
            self.log_message(f"❌ Validation error: {e}")
            self.show_error("Validation Error", str(e))
    
    def _show_analysis_results(self, analysis_data: Dict[str, Any], opportunities: List[CleanupOpportunity]):
        """Display the analysis results"""
        # Show analysis summary in the right panel
        self._show_analysis_summary(analysis_data, opportunities)
        
        # Clear existing results content
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        # Show results frame if there are opportunities
        if opportunities:
            self.results_frame.pack(fill=tk.X, pady=(0, 15))
        
        if opportunities:
            # Show cleanup options in the right panel of the log section
            self._show_cleanup_options_in_log_panel(opportunities)
        
        # Ensure main window maintains minimum height to show all controls
        self.main_app.root.update_idletasks()
        current_height = self.main_app.root.winfo_height()
        min_height = 850  # Increased minimum height to ensure bottom buttons are visible
        if current_height < min_height:
            current_width = self.main_app.root.winfo_width()
            self.main_app.root.geometry(f"{current_width}x{min_height}")
        
        # Clean up any visual artifacts that might appear
        self._cleanup_visual_artifacts()
    
    def _show_cleanup_options_in_log_panel(self, opportunities: List[CleanupOpportunity]):
        """Show cleanup options in the right panel of the log section"""
        if not hasattr(self, 'log_options_frame') or not self.log_options_frame:
            return
        
        # Clear existing options content (except log controls)
        for widget in self.log_options_frame.winfo_children():
            if not isinstance(widget, ttk.Frame) or len(widget.winfo_children()) == 0:  # Don't remove log controls
                continue
            # Check if this is the log controls frame
            has_export_button = any(isinstance(child, ttk.Button) and "Export" in child.cget('text') 
                                  for child in widget.winfo_children())
            if not has_export_button:
                widget.destroy()
        
        # Analyze what opportunities exist
        has_themes = any(op.item_type == 'theme' for op in opportunities)
        has_visuals = any(op.item_type in ['custom_visual_build_pane', 'custom_visual_hidden'] for op in opportunities)
        has_bookmarks = any(op.item_type in ['bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group'] for op in opportunities)
        has_filters = any(op.item_type == 'visual_filter' for op in opportunities)
        
        # Cleanup options section
        cleanup_section = ttk.Frame(self.log_options_frame)
        cleanup_section.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(cleanup_section, text="Cleanup Options:", 
                 font=('Segoe UI', 10, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 8))
        
        # Checkboxes with conditional enabling
        theme_cb = ttk.Checkbutton(
            cleanup_section,
            text="🎨 Remove unused themes",
            variable=self.cleanup_options['remove_themes'],
            command=self._update_cleanup_button_state
        )
        theme_cb.pack(anchor=tk.W, pady=2)
        if not has_themes:
            theme_cb.config(state=tk.DISABLED)
            self.cleanup_options['remove_themes'].set(False)
        
        visual_cb = ttk.Checkbutton(
            cleanup_section,
            text="📊 Remove unused visuals",
            variable=self.cleanup_options['remove_custom_visuals'],
            command=self._update_cleanup_button_state
        )
        visual_cb.pack(anchor=tk.W, pady=2)
        if not has_visuals:
            visual_cb.config(state=tk.DISABLED)
            self.cleanup_options['remove_custom_visuals'].set(False)
        
        bookmark_cb = ttk.Checkbutton(
            cleanup_section,
            text="📖 Remove unused bookmarks",
            variable=self.cleanup_options['remove_bookmarks'],
            command=self._update_cleanup_button_state
        )
        bookmark_cb.pack(anchor=tk.W, pady=2)
        if not has_bookmarks:
            bookmark_cb.config(state=tk.DISABLED)
            self.cleanup_options['remove_bookmarks'].set(False)
        
        filter_cb = ttk.Checkbutton(
            cleanup_section,
            text="🎯 Hide visual level filters",
            variable=self.cleanup_options['hide_visual_filters'],
            command=self._update_cleanup_button_state
        )
        filter_cb.pack(anchor=tk.W, pady=2)
        if not has_filters:
            filter_cb.config(state=tk.DISABLED)
            self.cleanup_options['hide_visual_filters'].set(False)
        
        # Action buttons
        buttons_frame = ttk.Frame(self.log_options_frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.cleanup_action_button = ttk.Button(
            buttons_frame,
            text="🧹 CLEAN ITEMS",
            command=self._remove_selected_items,
            style='Action.TButton',
            width=20  # Increased width to prevent text cutoff
        )
        self.cleanup_action_button.pack(fill=tk.X, pady=(0, 5))
        
        info_button = ttk.Button(
            buttons_frame,
            text="ℹ️ What gets cleaned?",
            command=self._show_removal_info,
            style='Secondary.TButton',
            width=20  # Increased width to prevent text cutoff
        )
        info_button.pack(fill=tk.X)
        
        # Update button state
        self._update_cleanup_button_state()
    
    def _update_cleanup_button_state(self):
        """Update the cleanup button state and removal preview based on selected options"""
        if hasattr(self, 'cleanup_action_button'):
            remove_themes = self.cleanup_options['remove_themes'].get()
            remove_visuals = self.cleanup_options['remove_custom_visuals'].get()
            remove_bookmarks = self.cleanup_options['remove_bookmarks'].get()
            hide_filters = self.cleanup_options['hide_visual_filters'].get()
            
            if remove_themes or remove_visuals or remove_bookmarks or hide_filters:
                self.cleanup_action_button.config(state=tk.NORMAL)
            else:
                self.cleanup_action_button.config(state=tk.DISABLED)
        
        # Update the removal preview in the log
        self._update_removal_preview()
    
    def _update_removal_preview(self):
        """Update the removal preview at the bottom of the log"""
        if not hasattr(self, 'current_opportunities') or not self.current_opportunities:
            return
        
        # Check current selections
        remove_themes = self.cleanup_options['remove_themes'].get()
        remove_visuals = self.cleanup_options['remove_custom_visuals'].get()
        remove_bookmarks = self.cleanup_options['remove_bookmarks'].get()
        hide_filters = self.cleanup_options['hide_visual_filters'].get()
        
        # Always clear any existing preview first
        self._clear_removal_preview()
        
        # Get items that would be removed
        themes_to_remove = []
        visuals_to_remove = []
        bookmarks_to_remove = []
        filters_to_hide = []
        
        for opportunity in self.current_opportunities:
            if opportunity.item_type == 'theme' and remove_themes:
                themes_to_remove.append(opportunity.item_name)
            elif opportunity.item_type in ['custom_visual_build_pane', 'custom_visual_hidden'] and remove_visuals:
                visuals_to_remove.append(opportunity.item_name)
            elif opportunity.item_type in ['bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group'] and remove_bookmarks:
                bookmarks_to_remove.append(opportunity.item_name)
            elif opportunity.item_type == 'visual_filter' and hide_filters:
                filters_to_hide.append(opportunity.item_name)
        
        # Add preview header and content if there are items to remove/hide
        if themes_to_remove or visuals_to_remove or bookmarks_to_remove or filters_to_hide:
            self.log_message("")
            self.log_message("📋 CLEANUP PREVIEW - Items that will be removed/hidden:")
            
            # Summary lines that get updated
            if themes_to_remove:
                self.log_message(f"🎨 Themes ({len(themes_to_remove)}): {', '.join(themes_to_remove)}")
            
            if visuals_to_remove:
                self.log_message(f"🔮 Custom Visuals ({len(visuals_to_remove)}): {', '.join(visuals_to_remove[:3])}{'...' if len(visuals_to_remove) > 3 else ''}")
            
            if bookmarks_to_remove:
                self.log_message(f"📖 Bookmarks ({len(bookmarks_to_remove)}): {', '.join(bookmarks_to_remove[:3])}{'...' if len(bookmarks_to_remove) > 3 else ''}")
            
            if filters_to_hide:
                self.log_message(f"🎯 Visual Filters: {', '.join(filters_to_hide)} (will be hidden)")
            
            total_items = len(themes_to_remove) + len(visuals_to_remove) + len(bookmarks_to_remove) + len(filters_to_hide)
            action_text = "remove/hide" if filters_to_hide else "remove"
            self.log_message(f"📊 Total items to {action_text}: {total_items}")
    
    def _clear_removal_preview(self):
        """Clear any existing removal preview from the log"""
        if not self.log_text:
            return
        
        # Get current log content
        self.log_text.config(state=tk.NORMAL)
        content = self.log_text.get(1.0, tk.END)
        
        # Split into lines and remove preview section
        lines = content.rstrip('\n').split('\n')
        filtered_lines = []
        in_preview = False
        
        for line in lines:
            if '📋 CLEANUP PREVIEW' in line or '📋 REMOVAL PREVIEW' in line:
                in_preview = True
                continue
            elif in_preview and (line.startswith('🎨 Themes') or 
                               line.startswith('🔮 Custom Visuals') or 
                               line.startswith('📖 Bookmarks') or 
                               line.startswith('🎯 Visual Filters') or
                               line.startswith('📊 Total items') or
                               line.strip() == ''):
                continue  # Skip preview content lines
            else:
                in_preview = False
                filtered_lines.append(line)
        
        # Remove any trailing empty lines from filtered content
        while filtered_lines and filtered_lines[-1].strip() == '':
            filtered_lines.pop()
        
        # Update log content
        self.log_text.delete(1.0, tk.END)
        if filtered_lines:
            self.log_text.insert(1.0, '\n'.join(filtered_lines) + '\n')
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
    
    def _show_analysis_summary(self, analysis_data: Dict[str, Any], opportunities: List[CleanupOpportunity]):
        """Show analysis summary in the right panel of the analysis section"""
        # Clear existing summary content
        for widget in self.analysis_summary_frame.winfo_children():
            widget.destroy()
        
        # Show the summary frame
        self.analysis_summary_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(20, 0))
        
        # Summary header
        ttk.Label(self.analysis_summary_frame, text="Analysis Summary", 
                 font=('Segoe UI', 11, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
        
        # Calculate summary stats
        theme_opportunities = len([op for op in opportunities if op.item_type == 'theme'])
        build_pane_visual_opportunities = len([op for op in opportunities if op.item_type == 'custom_visual_build_pane'])
        hidden_visual_opportunities = len([op for op in opportunities if op.item_type == 'custom_visual_hidden'])
        total_visual_opportunities = build_pane_visual_opportunities + hidden_visual_opportunities
        bookmark_opportunities = len([op for op in opportunities if op.item_type in ['bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group']])
        filter_opportunities = len([op for op in opportunities if op.item_type == 'visual_filter'])
        total_filter_count = sum(op.filter_count for op in opportunities if op.item_type == 'visual_filter')
        
        # Get total counts from analysis data
        total_themes = len(analysis_data['themes']['available_themes'])
        # Get active theme info
        active_theme = analysis_data['themes'].get('active_theme')
        active_theme_count = 1 if active_theme else 0
        
        appsource_visuals = len(analysis_data['custom_visuals'].get('appsource_visuals', {}))
        build_pane_visuals = len(analysis_data['custom_visuals'].get('build_pane_visuals', {}))
        hidden_visuals = len(analysis_data['custom_visuals'].get('hidden_visuals', {}))
        total_visuals = appsource_visuals + build_pane_visuals + hidden_visuals
        
        total_bookmarks = len(analysis_data['bookmarks'].get('bookmarks', {}))
        
        total_size_bytes = sum(op.size_bytes for op in opportunities)
        
        # Create summary text based on whether cleanup is needed
        if not opportunities:
            # No cleanup needed - show success message
            active_theme_name = "Unknown"
            if active_theme:
                active_theme_name = active_theme[0]
            
            summary_text = (f"🎉 No cleanup needed!\n"
                           f"Active theme: {active_theme_name}\n"
                           f"Report is optimized.\n\n"
                           f"📊 Themes: {total_themes} (1 active)\n"
                           f"🔮 Visuals: {total_visuals} (all used)\n"
                           f"📖 Bookmarks: {total_bookmarks} (all used)\n"
                           f"🎯 Visual filters: 0 (all necessary)")
        else:
            # Cleanup needed - show standard summary
            filter_text = ""
            if total_filter_count > 0:
                filter_text = f"\n🎯 Visual filters: {total_filter_count} (can be hidden)"
            
            summary_text = (f"📊 Themes: {total_themes} (1 active, {theme_opportunities} unused)\n"
                           f"🔮 Visuals: {total_visuals} ({total_visual_opportunities} unused)\n"
                           f"   • Build pane: {appsource_visuals + build_pane_visuals} ({build_pane_visual_opportunities} unused)\n"
                           f"   • Hidden: {hidden_visuals} ({hidden_visual_opportunities} unused)\n"
                           f"📖 Bookmarks: {total_bookmarks} ({bookmark_opportunities} unused)\n"
                           f"💾 Space savings: {self._format_bytes(total_size_bytes)}" + filter_text)
        
        ttk.Label(self.analysis_summary_frame, text=summary_text, font=('Segoe UI', 9),
                 foreground=AppConstants.COLORS['text_secondary'],
                 wraplength=300).pack(anchor=tk.W, pady=(5, 0))
    

    
    def _remove_selected_items(self):
        """Remove the selected themes, custom visuals, and/or bookmarks"""
        try:
            # Check if any options are selected
            remove_themes = self.cleanup_options['remove_themes'].get()
            remove_visuals = self.cleanup_options['remove_custom_visuals'].get()
            remove_bookmarks = self.cleanup_options['remove_bookmarks'].get()
            hide_filters = self.cleanup_options['hide_visual_filters'].get()
            
            if not remove_themes and not remove_visuals and not remove_bookmarks and not hide_filters:
                self.show_warning("No Options Selected", "Please select at least one cleanup option (themes, custom visuals, bookmarks, or visual filters).")
                return
            
            # Get items to remove based on options
            themes_to_remove = []
            visuals_to_remove = []  # Contains CleanupOpportunity objects
            bookmarks_to_remove = []  # Contains CleanupOpportunity objects
            filters_to_hide = False  # Boolean flag
            
            for opportunity in self.current_opportunities:
                if opportunity.item_type == 'theme' and remove_themes:
                    themes_to_remove.append(opportunity.item_name)
                elif opportunity.item_type in ['custom_visual_build_pane', 'custom_visual_hidden'] and remove_visuals:
                    visuals_to_remove.append(opportunity)  # Pass the whole opportunity object
                elif opportunity.item_type in ['bookmark_guaranteed_unused', 'bookmark_likely_unused', 'bookmark_empty_group'] and remove_bookmarks:
                    bookmarks_to_remove.append(opportunity)  # Pass the whole opportunity object
                elif opportunity.item_type == 'visual_filter' and hide_filters:
                    filters_to_hide = True  # Just set flag to hide all visual filters
            
            if not themes_to_remove and not visuals_to_remove and not bookmarks_to_remove and not filters_to_hide:
                self.show_warning("No Items to Process", "No items to remove or hide based on current selection.")
                return
            
            # Build confirmation message
            confirm_parts = []
            if themes_to_remove:
                theme_list = "\n".join([f"  🎨 {name}" for name in themes_to_remove])
                confirm_parts.append(f"THEMES ({len(themes_to_remove)}):\n{theme_list}")
            
            if visuals_to_remove:
                # Group visuals by type for better display
                build_pane = [v for v in visuals_to_remove if v.item_type == 'custom_visual_build_pane']
                hidden = [v for v in visuals_to_remove if v.item_type == 'custom_visual_hidden']
                
                visual_parts = []
                if build_pane:
                    build_list = "\n".join([f"  🔮 {v.item_name}" for v in build_pane])
                    visual_parts.append(f"BUILD PANE VISUALS ({len(build_pane)}):\n{build_list}")
                
                if hidden:
                    hidden_list = "\n".join([f"  😫 {v.item_name}" for v in hidden])
                    visual_parts.append(f"HIDDEN VISUALS ({len(hidden)}):\n{hidden_list}")
                
                if visual_parts:
                    confirm_parts.append("\n\n".join(visual_parts))
            
            if bookmarks_to_remove:
                # Group bookmarks by type for better display
                guaranteed = [b for b in bookmarks_to_remove if b.item_type == 'bookmark_guaranteed_unused']
                likely = [b for b in bookmarks_to_remove if b.item_type == 'bookmark_likely_unused']
                empty_groups = [b for b in bookmarks_to_remove if b.item_type == 'bookmark_empty_group']
                
                bookmark_parts = []
                if guaranteed:
                    guaranteed_list = "\n".join([f"  ✅ {b.item_name} (page missing)" for b in guaranteed])
                    bookmark_parts.append(f"GUARANTEED UNUSED ({len(guaranteed)}):  \n{guaranteed_list}")
                
                if likely:
                    likely_list = "\n".join([f"  ⚠️ {b.item_name} (no navigation found)" for b in likely])
                    bookmark_parts.append(f"LIKELY UNUSED ({len(likely)}):  \n{likely_list}")
                
                if empty_groups:
                    empty_list = "\n".join([f"  📁 {b.item_name} (empty group)" for b in empty_groups])
                    bookmark_parts.append(f"EMPTY GROUPS ({len(empty_groups)}):  \n{empty_list}")
                
                if bookmark_parts:
                    confirm_parts.append("\n\n".join(bookmark_parts))
            
            if filters_to_hide:
                # Get filter count from opportunity
                filter_opportunity = next((op for op in self.current_opportunities if op.item_type == 'visual_filter'), None)
                filter_count = filter_opportunity.filter_count if filter_opportunity else 0
                filter_text = f"VISUAL FILTERS ({filter_count}):\n  🎯 All visual-level filters will be hidden (not removed)"
                confirm_parts.append(filter_text)
            
            # Add special warning for likely unused bookmarks
            warning_note = ""
            if any(b.item_type == 'bookmark_likely_unused' for b in bookmarks_to_remove):
                warning_note = "\n\n📖 NOTE: 'Likely unused' bookmarks could still be accessed via the bookmark pane in Power BI Service, even without navigation buttons."
            
            if filters_to_hide:
                filter_note = "\n\n🎯 NOTE: Visual filters will be hidden (not deleted) and can be shown again later if needed."
                warning_note += filter_note
            
            action_text = "remove/hide these items" if filters_to_hide else "remove these items"
            confirm_message = (f"Are you sure you want to {action_text}?\n\n" +
                             "\n\n".join(confirm_parts) +
                             warning_note +
                             f"\n\n⚠️ This action will modify your PBIP file.\n"
                             f"💡 Consider creating a backup copy of your PBIP file before proceeding.")
            
            if not self.ask_yes_no("Confirm Cleanup", confirm_message):
                return
            
            # Perform removal
            pbip_path = self.pbip_path_var.get().strip()
            
            def run_removal():
                return self.cleanup_engine.remove_unused_items(
                    pbip_path, 
                    themes_to_remove if remove_themes else None,
                    visuals_to_remove if remove_visuals else None,
                    bookmarks_to_remove if remove_bookmarks else None,
                    hide_visual_filters=filters_to_hide,
                    create_backup=False  # No automatic backup
                )
            
            def on_success(results: List[RemovalResult]):
                self._handle_removal_results(results)
            
            def on_error(error):
                self.show_error("Cleanup Failed", f"Cleanup operation failed:\n\n{error}")
            
            # Progress steps for removal
            progress_steps = [
                ("Removing themes...", 20) if remove_themes else None,
                ("Removing custom visuals...", 40) if remove_visuals else None,
                ("Removing bookmarks...", 60) if remove_bookmarks else None,
                ("Hiding visual filters...", 80) if filters_to_hide else None,
                ("Cleaning up references...", 90),
                ("Cleanup complete!", 100)
            ]
            progress_steps = [step for step in progress_steps if step is not None]
            
            self.run_in_background(run_removal, on_success, on_error, progress_steps)
            
        except Exception as e:
            self.show_error("Cleanup Error", str(e))
    
    def _handle_removal_results(self, results: List[RemovalResult]):
        """Handle the results of cleanup operation"""
        successful_removals = [r for r in results if r.success]
        failed_removals = [r for r in results if not r.success]
        
        # Calculate total bytes freed
        total_bytes_freed = sum(r.bytes_freed for r in successful_removals)
        
        # Log results by type
        theme_results = [r for r in successful_removals if r.item_type == 'theme']
        visual_results = [r for r in successful_removals if r.item_type == 'custom_visual']
        bookmark_results = [r for r in successful_removals if r.item_type == 'bookmark']
        filter_results = [r for r in successful_removals if r.item_type == 'visual_filter']
        
        self.log_message(f"\n🎯 CLEANUP RESULTS:")
        
        if theme_results:
            self.log_message(f"🎨 Themes removed: {len(theme_results)}")
            for result in theme_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}")
        
        if visual_results:
            self.log_message(f"🔮 Custom visuals removed: {len(visual_results)}")
            for result in visual_results:
                size_str = f" ({self._format_bytes(result.bytes_freed)})" if result.bytes_freed > 0 else ""
                self.log_message(f"  ✅ {result.item_name}{size_str}")
        
        if bookmark_results:
            self.log_message(f"📖 Bookmarks removed: {len(bookmark_results)}")
            for result in bookmark_results:
                self.log_message(f"  ✅ {result.item_name}")
        
        if filter_results:
            for result in filter_results:
                self.log_message(f"🎯 Visual filters hidden: {result.filters_hidden}")
                self.log_message(f"  ✅ {result.item_name}")
        
        if failed_removals:
            self.log_message(f"❌ Failed removals: {len(failed_removals)}")
            for result in failed_removals:
                self.log_message(f"  ❌ {result.item_name}: {result.error_message}")
        
        self.log_message(f"💾 Total space freed: {self._format_bytes(total_bytes_freed)}")
        
        # Show completion message
        if successful_removals and not failed_removals:
            self.show_info("Cleanup Complete", 
                         f"Successfully cleaned up {len(successful_removals)} items!\n\n"
                         f"Space freed: {self._format_bytes(total_bytes_freed)}\n"
                         f"Your PBIP report has been cleaned up.")
        elif successful_removals and failed_removals:
            self.show_warning("Partial Success", 
                            f"Removed {len(successful_removals)} items successfully.\n"
                            f"{len(failed_removals)} items could not be removed.\n\n"
                            f"Check the log for details.")
        else:
            self.show_error("Cleanup Failed", 
                          f"Could not clean any of the selected items.\n\n"
                          f"Check the log for details.")
        
        # Refresh the analysis to show updated state
        self._analyze_report()
    
    def _show_removal_info(self):
        """Show information about what gets cleaned"""
        info_message = (
            "🎨 THEMES:\n"
            "• Theme files (.json) in BaseThemes and RegisteredResources\n"
            "• References in report.json resourcePackages\n"
            "• Only inactive themes are identified for removal\n\n"
            
            "🔮 CUSTOM VISUALS:\n"
            "• Custom visual directories in CustomVisuals folder\n"
            "• References in report.json resourcePackages\n"
            "• Includes both 'hidden' visuals and unused visuals in build pane\n\n"
            
            "📖 BOOKMARKS:\n"
            "• Guaranteed unused: Bookmarks pointing to missing pages\n"
            "• Likely unused: Bookmarks with no navigation buttons found\n"
            "• References in report.json bookmarks array\n"
            "• Warning: 'Likely unused' could still be accessed via bookmark pane\n\n"
            
            "🎯 VISUAL FILTERS:\n"
            "• Sets isHiddenInViewMode: true on all visual-level filters\n"
            "• Filters are hidden but can be shown again later if needed\n"
            "• Does not affect page-level or report-level filters\n\n"
            
            "✅ SAFETY:\n"
            "• Only items confirmed as unused are marked for removal\n"
            "• Active themes are never removed\n"
            "• Custom visuals used in pages are never removed\n"
            "• Bookmarks with navigation found are never removed\n"
            "• Visual filters are hidden, not deleted\n"
            "• Consider creating a backup before cleanup"
        )
        
        self.show_info("What Gets Cleaned", info_message)
    
    def _format_bytes(self, bytes_count: int) -> str:
        """Format bytes into human readable format"""
        if bytes_count == 0:
            return "0 B"
        
        for unit in ['B', 'KB', 'MB', 'GB']:
            if bytes_count < 1024:
                return f"{bytes_count:.1f} {unit}"
            bytes_count /= 1024
        return f"{bytes_count:.1f} TB"
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        # Clear input
        self.pbip_path_var.set("")
        
        # Clear results
        self.current_opportunities.clear()
        
        # Reset options
        self.cleanup_options['remove_themes'].set(True)
        self.cleanup_options['remove_custom_visuals'].set(True)
        self.cleanup_options['remove_bookmarks'].set(True)
        self.cleanup_options['hide_visual_filters'].set(True)
        
        # Hide results and summary sections
        if self.results_frame:
            self.results_frame.pack_forget()
        if hasattr(self, 'analysis_summary_frame'):
            self.analysis_summary_frame.pack_forget()
        
        # Clear cleanup options from log panel
        if hasattr(self, 'log_options_frame') and self.log_options_frame:
            for widget in self.log_options_frame.winfo_children():
                # Keep log controls, remove cleanup options
                if not isinstance(widget, ttk.Frame):
                    continue
                has_export_button = any(isinstance(child, ttk.Button) and "Export" in child.cget('text') 
                                      for child in widget.winfo_children())
                if not has_export_button:
                    widget.destroy()
        
        # Reset buttons
        if self.analyze_button:
            self.analyze_button.config(state=tk.NORMAL)
        
        # Clear log completely and remove any preview
        self._clear_removal_preview()
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
        
        # Hide progress bar if visible
        if hasattr(self, 'progress_components') and self.progress_components:
            self.progress_components['frame'].pack_forget()
        
        # Force cleanup of any lingering widgets that might cause visual artifacts
        self._cleanup_visual_artifacts()
        
        # Force a UI update to remove any lingering widgets
        self.frame.update_idletasks()
        
        # Show welcome message
        self._show_welcome_message()
    
    def _cleanup_visual_artifacts(self):
        """Aggressively clean up widgets that might cause visual artifacts like black lines"""
        def cleanup_widget_recursively(widget):
            try:
                # Check all children recursively
                for child in widget.winfo_children():
                    cleanup_widget_recursively(child)
                
                # Check for problematic widgets
                widget_class = widget.winfo_class()
                widget_height = widget.winfo_height() if hasattr(widget, 'winfo_height') else 0
                widget_width = widget.winfo_width() if hasattr(widget, 'winfo_width') else 0
                
                # Remove widgets that might cause visual artifacts
                if (
                    # Very thin widgets that might be lines
                    (widget_height <= 5 and widget_width > 50) or
                    # Widgets with dark backgrounds
                    (widget_class in ['Frame', 'Separator', 'Label'] and widget_height < 10) or
                    # Empty frames with borders
                    (widget_class == 'Frame' and widget_height < 15 and not widget.winfo_children())
                ):
                    # Try to hide the widget
                    if hasattr(widget, 'pack_forget'):
                        widget.pack_forget()
                    if hasattr(widget, 'grid_forget'):
                        widget.grid_forget()
                    if hasattr(widget, 'place_forget'):
                        widget.place_forget()
                        
            except Exception:
                # Ignore errors during cleanup
                pass
        
        # Start cleanup from the main frame
        cleanup_widget_recursively(self.frame)
        
        # Also check the main app window
        if hasattr(self.main_app, 'root'):
            cleanup_widget_recursively(self.main_app.root)
    
    def show_help_dialog(self) -> None:
        def create_help_content(help_window):
            """Create help content"""
            main_frame = ttk.Frame(help_window, padding="30")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Header
            ttk.Label(main_frame, text="🧹 Report Cleanup - Help", 
                     font=('Segoe UI', 16, 'bold'),
                     foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 20))
            
            # Content sections
            help_content = self.main_app.tool_manager.get_tool_help_content("report_cleanup")
            
            if help_content and 'sections' in help_content:
                for section in help_content['sections']:
                    # Section header
                    ttk.Label(main_frame, text=section['title'], 
                             font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W, pady=(15, 5))
                    
                    # Section items
                    for item in section['items']:
                        ttk.Label(main_frame, text=f"  {item}", 
                                 font=('Segoe UI', 9),
                                 wraplength=500).pack(anchor=tk.W, pady=1)
            
            # Close button
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill=tk.X, pady=(20, 0))
            
            ttk.Button(button_frame, text="❌ Close", 
                      command=help_window.destroy,
                      style='Action.TButton').pack()
        
        self.create_help_window("Report Cleanup - Help", create_help_content)
    
    def _show_welcome_message(self):
        """Show welcome message"""
        self.log_message("🧹 Welcome to Report Cleanup!")
        self.log_message("=" * 60)
        self.log_message("This tool helps you clean up your Power BI reports by:")
        self.log_message("• Detecting unused themes in BaseThemes and RegisteredResources")
        self.log_message("• Finding custom visuals that aren't used in any pages")
        self.log_message("• Identifying 'hidden' custom visuals taking up space")
        self.log_message("• Detecting unused bookmarks (missing pages or no navigation)")
        self.log_message("• Hiding visual-level filters to clean up the interface")
        self.log_message("• Providing detailed analysis and removal reports")
        self.log_message("")
        self.log_message("🎯 Start by selecting your PBIP file and clicking 'ANALYZE REPORT'")
        self.log_message("💡 Consider backing up your PBIP file before making changes")
        self.log_message("")
