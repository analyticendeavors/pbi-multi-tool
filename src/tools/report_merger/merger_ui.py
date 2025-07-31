"""
Report Merger UI Tab - Refactored to use BaseToolTab
Built by Reid Havens of Analytic Endeavors
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from pathlib import Path
from typing import Optional, Dict, Any

from core.constants import AppConstants
from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin
from tools.report_merger.merger_core import MergerEngine, ValidationService, ValidationError


class ReportMergerTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """Report Merger tab - refactored to use new base architecture"""
    
    def __init__(self, parent, main_app):
        super().__init__(parent, main_app, "report_merger", "Report Merger")
        
        # UI Variables
        self.report_a_path = tk.StringVar()
        self.report_b_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.theme_choice = tk.StringVar(value="report_a")
        
        # Core components
        self.validation_service = ValidationService()
        self.merger_engine = MergerEngine(logger_callback=self.log_message)
        
        # UI Components
        self.analyze_button = None
        self.merge_button = None
        self.theme_frame = None
        self.log_components = None  # Store log components for positioning
        
        # State
        self.analysis_results = None
        self.is_analyzing = False
        self.is_merging = False
        
        # Setup UI and events
        self.setup_ui()
        self._setup_events()
    
    def setup_ui(self) -> None:
        """Setup the UI for the report merger tab"""
        # Setup sections using pack layout
        self._setup_data_sources()
        self._setup_theme_selection()
        self._setup_output_section()
        
        # Create log section using base class
        self.log_components = self.create_log_section(self.frame)
        self.log_components['frame'].pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Create action buttons using base class
        self.button_frame = ttk.Frame(self.frame)
        self.button_frame.pack(pady=(20, 0))
        
        self.merge_button = ttk.Button(self.button_frame, text="🚀 EXECUTE MERGE", 
                                     command=self.start_merge,
                                     style='Action.TButton', state=tk.DISABLED)
        self.merge_button.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Button(self.button_frame, text="🔄 RESET ALL", command=self.reset_tab,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 15))
        
        # Create progress bar using base class
        self.create_progress_bar(self.frame)
        
        # Show welcome message
        self._show_welcome_message()
    
    def _position_progress_frame(self):
        """Position progress frame specifically for Report Merger layout"""
        if self.progress_frame and self.button_frame:
            # Position progress bar right before the button frame to avoid layout conflicts
            self.progress_frame.pack(before=self.button_frame, fill=tk.X, pady=(10, 5))
        elif self.progress_frame:
            # Fallback: position at bottom if button frame not found
            self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10))
    
    def _setup_data_sources(self):
        """Setup data sources section"""
        # Guide text for the file input section
        guide_text = [
            "🚀 QUICK START GUIDE:",
            "1. Navigate to your .pbip file in File Explorer",
            "2. Right-click the .pbip file and select 'Copy as path'", 
            "3. Paste (Ctrl+V) into the path field",
            "4. Path quotes will be automatically cleaned",
            "5. Repeat for the second report file",
            "6. Click 'Analyze Reports' to begin",
            "⚠️ Requires PBIP format"
        ]
        
        # Create main section frame
        section_frame = ttk.LabelFrame(self.frame, text="📁 PBIP FILE SOURCE", 
                                     style='Section.TLabelframe', padding="20")
        section_frame.pack(fill=tk.X, pady=(0, 15))
        
        content_frame = ttk.Frame(section_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # LEFT: Guide text
        guide_frame = ttk.Frame(content_frame)
        guide_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 35))
        
        for i, text in enumerate(guide_text):
            if i == 0:  # Title
                ttk.Label(guide_frame, text=text, 
                         font=('Segoe UI', 10, 'bold'), 
                         foreground=AppConstants.COLORS['info']).pack(anchor=tk.W)
            else:  # Steps
                ttk.Label(guide_frame, text=f"   {text}", font=('Segoe UI', 9),
                         foreground=AppConstants.COLORS['text_secondary'], 
                         wraplength=300).pack(anchor=tk.W, pady=1)
        
        # RIGHT: File inputs
        input_frame = ttk.Frame(content_frame)
        input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        input_frame.columnconfigure(1, weight=1)
        
        # Report A input
        ttk.Label(input_frame, text="Project File (PBIP):").grid(row=0, column=0, sticky=tk.W, pady=8)
        entry_a = ttk.Entry(input_frame, textvariable=self.report_a_path, width=80)
        entry_a.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        ttk.Button(input_frame, text="📂 Browse", 
                  command=lambda: self._browse_file(self.report_a_path, "Report A")).grid(row=0, column=2, pady=8)
        
        # Report B input
        ttk.Label(input_frame, text="Project File (PBIP):").grid(row=1, column=0, sticky=tk.W, pady=8)
        entry_b = ttk.Entry(input_frame, textvariable=self.report_b_path, width=80)
        entry_b.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        ttk.Button(input_frame, text="📂 Browse", 
                  command=lambda: self._browse_file(self.report_b_path, "Report B")).grid(row=1, column=2, pady=8)
        
        # Analyze button
        self.analyze_button = ttk.Button(input_frame, text="🔍 ANALYZE REPORTS",
                                       command=self.analyze_reports, 
                                       style='Action.TButton', state=tk.DISABLED)
        self.analyze_button.grid(row=2, column=0, columnspan=3, pady=(15, 0))
        
        # Setup path cleaning
        self.setup_path_cleaning(self.report_a_path)
        self.setup_path_cleaning(self.report_b_path)
    
    def _setup_theme_selection(self):
        """Setup theme selection (initially hidden)"""
        self.theme_frame = ttk.LabelFrame(self.frame, text="🎨 THEME CONFIGURATION", 
                                        style='Section.TLabelframe', padding="20")
    
    def _setup_output_section(self):
        """Setup output section"""
        output_frame = ttk.LabelFrame(self.frame, text="📁 OUTPUT CONFIGURATION", 
                                    style='Section.TLabelframe', padding="20")
        output_frame.pack(fill=tk.X, pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Output Path:").grid(row=0, column=0, sticky=tk.W, pady=8)
        ttk.Entry(output_frame, textvariable=self.output_path, width=80).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        ttk.Button(output_frame, text="📂 Browse", 
                  command=self.browse_output).grid(row=0, column=2, pady=8)
        
        # Setup path cleaning for output
        self.setup_path_cleaning(self.output_path)
    
    def _setup_events(self):
        """Setup event handlers"""
        self.report_a_path.trace('w', lambda *args: self._on_path_change())
        self.report_b_path.trace('w', lambda *args: self._on_path_change())
    
    def _show_welcome_message(self):
        """Show welcome message"""
        messages = [
            "🎉 Welcome to the Report Merger!",
            "📁 Select your Report A and Report B files to begin merging...",
            "⚠️ Requires PBIP format files only",
            "=" * 60
        ]
        for msg in messages:
            self.log_message(msg)
    
    def _on_path_change(self):
        """Handle path changes"""
        self._update_ui_state()
        self.auto_generate_output_path()
    
    def _update_ui_state(self):
        """Update UI state"""
        has_both = bool(self.report_a_path.get() and self.report_b_path.get())
        if self.analyze_button:
            self.analyze_button.config(state=tk.NORMAL if has_both else tk.DISABLED)
    
    def _browse_file(self, path_var: tk.StringVar, report_name: str):
        """Browse for report file"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title=f"Select {report_name} (.pbip file - PBIR format required)",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            path_var.set(file_path)
    
    def browse_output(self):
        """Browse for output location"""
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(
            title="Save Combined Report As", defaultextension=".pbip",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)
    
    def auto_generate_output_path(self):
        """Auto-generate output path"""
        report_a = self.clean_file_path(self.report_a_path.get())
        report_b = self.clean_file_path(self.report_b_path.get())
        
        if report_a and report_b:
            try:
                output_path = self.merger_engine.generate_output_path(report_a, report_b)
                self.output_path.set(output_path)
            except Exception:
                pass
    
    def analyze_reports(self):
        """Analyze selected reports"""
        try:
            report_a = self.clean_file_path(self.report_a_path.get())
            report_b = self.clean_file_path(self.report_b_path.get())
            self.validation_service.validate_input_paths(report_a, report_b)
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
        self.update_progress(10, "Validating input files...")
        report_a = self.clean_file_path(self.report_a_path.get())
        report_b = self.clean_file_path(self.report_b_path.get())
        
        self.update_progress(30, "Reading source report...")
        
        self.update_progress(50, "Reading target report...")
        
        self.update_progress(80, "Analyzing compatibility...")
        results = self.merger_engine.analyze_reports(report_a, report_b)
        
        self.update_progress(100, "Analysis complete!")
        return results
    
    def _handle_analysis_complete(self, results):
        """Handle analysis completion"""
        self.analysis_results = results
        
        # Show theme selection if needed
        if results['themes']['conflict']:
            self._show_theme_selection(results['themes']['theme_a'], results['themes']['theme_b'])
        else:
            self._hide_theme_selection()
            self.theme_choice.set("same")
        
        # Enable merge button
        if self.merge_button:
            self.merge_button.config(state=tk.NORMAL)
        
        self._show_analysis_summary(results)
    
    def _show_theme_selection(self, theme_a: Dict[str, Any], theme_b: Dict[str, Any]):
        """Show theme selection UI"""
        # Position theme frame right above the log section
        if self.log_components and self.log_components['frame']:
            self.theme_frame.pack(before=self.log_components['frame'], fill=tk.X, pady=(0, 15))
        else:
            # Fallback positioning
            self.theme_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Expand window height to accommodate theme selection
        self._adjust_window_height(True)
        
        # Clear existing content
        for widget in self.theme_frame.winfo_children():
            widget.destroy()
        
        # Theme selection content
        content_frame = ttk.Frame(self.theme_frame)
        content_frame.pack(fill=tk.X, pady=5)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # LEFT: Warning text
        warning_frame = ttk.Frame(content_frame)
        warning_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 30))
        
        ttk.Label(warning_frame, text="⚠️ THEME CONFLICT DETECTED", 
                 font=('Segoe UI', 11, 'bold'),
                 foreground=AppConstants.COLORS['warning']).pack(anchor=tk.W)
        
        ttk.Label(warning_frame, text="Choose which theme to apply:", 
                 font=('Segoe UI', 10)).pack(anchor=tk.W, pady=(5, 0))
        
        # RIGHT: Theme options
        selection_frame = ttk.Frame(content_frame)
        selection_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        
        # Configure radiobutton style to remove gray background
        style = ttk.Style()
        style.configure('Clean.TRadiobutton', 
                       background=AppConstants.COLORS['background'],
                       foreground=AppConstants.COLORS['text_primary'])
        
        ttk.Radiobutton(selection_frame, text=f"📊 Report A: {theme_a['display']}",
                       variable=self.theme_choice, value="report_a",
                       style='Clean.TRadiobutton').pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(selection_frame, text=f"📊 Report B: {theme_b['display']}",
                       variable=self.theme_choice, value="report_b",
                       style='Clean.TRadiobutton').pack(anchor=tk.W, pady=2)
    
    def _hide_theme_selection(self):
        """Hide theme selection UI"""
        if self.theme_frame:
            self.theme_frame.pack_forget()
            # Contract window height when theme selection is hidden
            self._adjust_window_height(False)
    
    def _adjust_window_height(self, show_theme_selection: bool):
        """Adjust window height based on theme selection visibility"""
        try:
            if hasattr(self.main_app, 'root'):
                root = self.main_app.root
                base_height = 950  # More compact height for Report Merger
                theme_selection_height = 160  # Height for theme configuration section
                
                if show_theme_selection:
                    new_height = base_height + theme_selection_height
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
    
    def _show_analysis_summary(self, results):
        """Show analysis summary"""
        self.log_message("\n📊 ANALYSIS SUMMARY")
        self.log_message("=" * 50)
        
        report_a = results['report_a']
        report_b = results['report_b']
        totals = results['totals']
        themes = results['themes']
        
        self.log_message(f"🔄 Combining: {report_a['name']} + {report_b['name']}")
        self.log_message(f"📄 Total Pages: {totals['pages']} ({report_a['pages']} + {report_b['pages']})")
        self.log_message(f"🔖 Total Bookmarks: {totals['bookmarks']} ({report_a['bookmarks']} + {report_b['bookmarks']})")
        
        if 'measures' in totals:
            self.log_message(f"📐 Total Measures: {totals['measures']}")
        
        if themes['conflict']:
            self.log_message(f"⚠️ Theme Conflict: {themes['theme_a']['name']} vs {themes['theme_b']['name']}")
            self.log_message("   🎨 Please select your preferred theme above")
        else:
            self.log_message(f"✅ Consistent Theme: {themes['theme_a']['name']}")
        
        self.log_message("✅ Analysis complete! Ready to merge.")
    
    def start_merge(self):
        """Start merge operation"""
        if not self.analysis_results:
            self.show_error("Error", "Please analyze reports first")
            return
        
        try:
            output_path = self.clean_file_path(self.output_path.get())
            self.validation_service.validate_output_path(output_path)
        except Exception as e:
            self.show_error("Output Validation Error", str(e))
            return
        
        totals = self.analysis_results['totals']
        if not self.ask_yes_no("Confirm Merge", 
                              f"Ready to merge reports?\n\n"
                              f"📊 Combined report will have:\n"
                              f"📄 {totals['pages']} pages\n"
                              f"🔖 {totals['bookmarks']} bookmarks\n"
                              f"📐 {totals.get('measures', 0)} measures\n\n"
                              f"💾 Output: {output_path}"):
            return
        
        # Use base class background processing
        self.run_in_background(
            target_func=self._merge_thread_target,
            success_callback=self._handle_merge_complete,
            error_callback=lambda e: self.show_error("Merge Error", str(e))
        )
    
    def _merge_thread_target(self):
        """Background merge logic"""
        self.log_message("\n🚀 Starting merge operation...")
        
        self.update_progress(10, "Preparing merge operation...")
        report_a = self.clean_file_path(self.report_a_path.get())
        report_b = self.clean_file_path(self.report_b_path.get())
        output_path = self.clean_file_path(self.output_path.get())
        theme_choice = self.theme_choice.get()
        
        self.update_progress(30, "Reading source reports...")
        
        self.update_progress(50, "Merging reports...")
        
        self.update_progress(75, "Applying theme configuration...")
        
        success = self.merger_engine.merge_reports(
            report_a, report_b, output_path, theme_choice, self.analysis_results
        )
        
        self.update_progress(90, "Finalizing merged report...")
        
        self.update_progress(100, "Merge operation complete!")
        return {'success': success, 'output_path': output_path}
    
    def _handle_merge_complete(self, result):
        """Handle merge completion"""
        if result['success']:
            self.log_message("✅ MERGE COMPLETED SUCCESSFULLY!")
            self.log_message(f"💾 Output: {result['output_path']}")
            
            self.show_info("Merge Complete", 
                          f"Merge completed successfully!\n\n💾 Output:\n{result['output_path']}")
        else:
            self.show_error("Merge Failed", "The merge operation failed. Check the log for details.")
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state"""
        if self.is_busy:
            if not self.ask_yes_no("Confirm Reset", "An operation is in progress. Stop and reset?"):
                return
        
        # Clear state
        self.report_a_path.set("")
        self.report_b_path.set("")
        self.output_path.set("")
        self.theme_choice.set("report_a")
        self.analysis_results = None
        
        # Reset UI state
        if self.analyze_button:
            self.analyze_button.config(state=tk.DISABLED)
        if self.merge_button:
            self.merge_button.config(state=tk.DISABLED)
        self._hide_theme_selection()  # This will also adjust window height
        
        # Clear log and show welcome
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
        
        self._show_welcome_message()
        self.log_message("✅ Report Merger reset successfully!")
    
    def show_help_dialog(self) -> None:
        """Show help dialog specific to report merger"""
        def create_help_content(help_window):
            main_frame = ttk.Frame(help_window, padding="20")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Header
            ttk.Label(main_frame, text="❓ Power BI Report Merger - Help", 
                     font=('Segoe UI', 16, 'bold'), 
                     foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 20))
            
            # PBIR Requirement Section
            pbir_frame = ttk.Frame(main_frame)
            pbir_frame.pack(fill=tk.X, pady=(0, 20))
            
            warning_container = tk.Frame(pbir_frame, bg=AppConstants.COLORS['warning'], 
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
                "• Always keep backups of your original reports before merging",
                "• Test thoroughly and validate merged results before production use",
                "• Enable 'Store reports using enhanced metadata format (PBIR)' in Power BI Desktop"
            ]
            
            for warning in warnings:
                ttk.Label(warning_container, text=warning, font=('Segoe UI', 10),
                         background=AppConstants.COLORS['warning'],
                         foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W, pady=1)
            
            # Help sections
            help_sections = [
                  ("🎯 What This Tool Does", [
                      "✅ Combines multiple Power BI reports into a single unified report",
                      "✅ Merges pages, bookmarks, and report-level settings intelligently",
                      "✅ Handles theme conflicts with user-selectable resolution",
                      "✅ Preserves all visuals, data sources, and formatting",
                      "✅ Creates a new merged report while keeping originals intact",
                      "✅ Provides detailed analysis and preview before merging"
                  ]),
                ("📁 File Requirements", [
                    "✅ Only .pbip files (enhanced PBIR format) are supported",
                    "✅ Reports must have definition\\ folder structure",
                    "❌ Legacy format with report.json files are NOT supported",
                    "❌ .pbix files are NOT supported"
                ])
            ]
            
            for title, items in help_sections:
                section_frame = ttk.Frame(main_frame)
                section_frame.pack(fill=tk.X, pady=(0, 15))
                
                ttk.Label(section_frame, text=title, font=('Segoe UI', 12, 'bold'),
                         foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
                
                for item in items:
                    ttk.Label(section_frame, text=f"   {item}", font=('Segoe UI', 10)).pack(anchor=tk.W, pady=1)
            
            # Button frame at bottom - fixed position
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill=tk.X, pady=(20, 0), side=tk.BOTTOM)
            
            close_button = ttk.Button(button_frame, text="❌ Close", 
                                    command=help_window.destroy,
                                    style='Action.TButton')
            close_button.pack(pady=(10, 0))
        
        # Create custom help window for Report Merger (independent of base class)
        help_window = tk.Toplevel(self.main_app.root)
        help_window.title("Power BI Report Merger - Help")
        help_window.geometry("670x700")  # Custom height for Report Merger
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
