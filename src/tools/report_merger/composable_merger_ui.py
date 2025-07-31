"""
Report Merger UI Tab - Composition Architecture Implementation
Built by Reid Havens of Analytic Endeavors

This implementation uses the new composition architecture with inheritance
from base tool classes for improved error handling and modularity.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from pathlib import Path
from typing import Optional, Dict, Any

from core.constants import AppConstants
from core.composition.tool_composition import UIComposableTool
from tools.report_merger.merger_core import MergerEngine, ValidationService, ValidationError


class ComposableReportMergerTab(UIComposableTool):
    """
    Report Merger tab using composition architecture.
    Inherits from UIComposableTool and composes functionality using components.
    """
    
    def __init__(self, parent, main_app):
        # Initialize the composable tool
        super().__init__(parent, main_app, "report_merger", "Report Merger")
        
        # Initialize components
        self.initialize_components()
        
        # UI Variables
        self.report_a_path = tk.StringVar()
        self.report_b_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.theme_choice = tk.StringVar(value="report_a")
        
        # Core components - now using dependency injection pattern
        self.merger_engine = MergerEngine(logger_callback=self.log_message)
        self.validation_service = ValidationService()
        
        # UI Components
        self.analyze_button = None
        self.merge_button = None
        self.theme_frame = None
        
        # State
        self.analysis_results = None
        self.is_analyzing = False
        self.is_merging = False
        
        # Setup UI and events
        self.setup_ui()
        self._setup_events()
    
    def setup_ui(self) -> None:
        """Setup the UI for the report merger tab using composition components"""
        # Setup sections using pack layout
        self._setup_data_sources()
        self._setup_theme_selection()
        self._setup_output_section()
        
        # Create log section using composition
        self._create_log_section()
        
        # Create action buttons
        self._create_action_buttons()
        
        # Create progress bar using composition
        self.progress.create_progress_bar(self.frame)
        
        # Show welcome message
        self._show_welcome_message()
    
    def _setup_data_sources(self):
        """Setup data sources section using composition components"""
        # Create main section frame
        section_frame = ttk.LabelFrame(self.frame, text="📁 DATA SOURCES", 
                                     style='Section.TLabelframe', padding="20")
        section_frame.pack(fill=tk.X, pady=(0, 15))
        
        content_frame = ttk.Frame(section_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # LEFT: Guide text
        guide_frame = ttk.Frame(content_frame)
        guide_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 35))
        
        guide_text = [
            "🚀 QUICK START GUIDE:",
            "1. Navigate to your .pbip file in File Explorer",
            "2. Right-click the .pbip file and select 'Copy as path'", 
            "3. Paste (Ctrl+V) into the path field",
            "4. Path quotes will be automatically cleaned",
            "5. Repeat for the second report file",
            "6. Click 'Analyze Reports' to begin",
            "⚠️ Requires PBIR format"
        ]
        
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
        ttk.Label(input_frame, text="Report A (PBIR):").grid(row=0, column=0, sticky=tk.W, pady=8)
        entry_a = ttk.Entry(input_frame, textvariable=self.report_a_path, width=80)
        entry_a.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        
        # Use composition for file browsing
        def browse_a():
            self._browse_report_file(self.report_a_path, "Report A")
        
        ttk.Button(input_frame, text="📂 Browse", command=browse_a).grid(row=0, column=2, pady=8)
        
        # Report B input
        ttk.Label(input_frame, text="Report B (PBIR):").grid(row=1, column=0, sticky=tk.W, pady=8)
        entry_b = ttk.Entry(input_frame, textvariable=self.report_b_path, width=80)
        entry_b.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        
        # Use composition for file browsing
        def browse_b():
            self._browse_report_file(self.report_b_path, "Report B")
        
        ttk.Button(input_frame, text="📂 Browse", command=browse_b).grid(row=1, column=2, pady=8)
        
        # Analyze button
        self.analyze_button = ttk.Button(input_frame, text="🔍 ANALYZE REPORTS",
                                       command=self.analyze_reports, 
                                       style='Action.TButton', state=tk.DISABLED)
        self.analyze_button.grid(row=2, column=0, columnspan=3, pady=(15, 0))
        
        # Setup path cleaning using composition
        self.file_input.setup_path_cleaning(self.report_a_path)
        self.file_input.setup_path_cleaning(self.report_b_path)
    
    def _setup_theme_selection(self):
        """Setup theme selection (initially hidden)"""
        self.theme_frame = ttk.LabelFrame(self.frame, text="🎨 THEME CONFIGURATION", 
                                        style='Section.TLabelframe', padding="20")
    
    def _setup_output_section(self):
        """Setup output section using composition"""
        output_frame = ttk.LabelFrame(self.frame, text="📁 OUTPUT CONFIGURATION", 
                                    style='Section.TLabelframe', padding="20")
        output_frame.pack(fill=tk.X, pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Output Path:").grid(row=0, column=0, sticky=tk.W, pady=8)
        ttk.Entry(output_frame, textvariable=self.output_path, width=80).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        
        def browse_output():
            self._browse_output_file()
        
        ttk.Button(output_frame, text="📂 Browse", command=browse_output).grid(row=0, column=2, pady=8)
        
        # Setup path cleaning for output using composition
        self.file_input.setup_path_cleaning(self.output_path)
    
    def _create_log_section(self):
        """Create log section using composition"""
        log_frame = ttk.LabelFrame(self.frame, text="📊 ANALYSIS & PROGRESS LOG", 
                                 style='Section.TLabelframe', padding="15")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        content_frame = ttk.Frame(log_frame)
        content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(
            content_frame, height=12, width=85, font=('Consolas', 9), state=tk.DISABLED,
            bg=AppConstants.COLORS['surface'], fg=AppConstants.COLORS['text_primary'],
            selectbackground=AppConstants.COLORS['accent'], relief='solid', borderwidth=1
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Log controls
        log_controls = ttk.Frame(content_frame)
        log_controls.grid(row=0, column=1, sticky=tk.N, padx=(10, 0))
        
        def export_log():
            self._export_log()
        
        def clear_log():
            self._clear_log()
        
        ttk.Button(log_controls, text="📤 Export Log", command=export_log,
                  style='Secondary.TButton', width=16).pack(pady=(0, 5), anchor=tk.W)
        
        ttk.Button(log_controls, text="🗑️ Clear Log", command=clear_log,
                  style='Secondary.TButton', width=16).pack(anchor=tk.W)
    
    def _create_action_buttons(self):
        """Create action buttons"""
        button_frame = ttk.Frame(self.frame)
        button_frame.pack(pady=(20, 0))
        
        self.merge_button = ttk.Button(button_frame, text="🚀 EXECUTE MERGE", 
                                     command=self.start_merge,
                                     style='Action.TButton', state=tk.DISABLED)
        self.merge_button.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Button(button_frame, text="🔄 RESET ALL", command=self.reset_tool,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 15))
    
    def _setup_events(self):
        """Setup event handlers"""
        self.report_a_path.trace('w', lambda *args: self._on_path_change())
        self.report_b_path.trace('w', lambda *args: self._on_path_change())
    
    def _setup_common_styling(self):
        """Setup styling using composition"""
        style = ttk.Style()
        style.theme_use('clam')
        colors = AppConstants.COLORS
        
        styles = {
            'Section.TLabelframe': {
                'background': colors['background'], 
                'borderwidth': 1, 
                'relief': 'solid'
            },
            'Section.TLabelframe.Label': {
                'background': colors['background'], 
                'foreground': colors['primary'], 
                'font': ('Segoe UI', 12, 'bold')
            },
            'Action.TButton': {
                'background': colors['primary'], 
                'foreground': colors['surface'], 
                'font': ('Segoe UI', 10, 'bold'), 
                'padding': (20, 10)
            },
            'Secondary.TButton': {
                'background': colors['border'], 
                'foreground': colors['text_primary'], 
                'font': ('Segoe UI', 10), 
                'padding': (15, 8)
            }
        }
        
        for style_name, config in styles.items():
            style.configure(style_name, **config)
    
    def _show_welcome_message(self):
        """Show welcome message"""
        messages = [
            "🎉 Welcome to the Report Merger! (Composition Architecture)",
            "📁 Select your Report A and Report B files to begin merging...",
            "⚠️ Requires PBIP format (PBIR) files only",
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
    
    def _browse_report_file(self, path_var: tk.StringVar, report_name: str):
        """Browse for report file using composition"""
        file_path = filedialog.askopenfilename(
            title=f"Select {report_name} (.pbip file - PBIR format required)",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            path_var.set(file_path)
    
    def _browse_output_file(self):
        """Browse for output location using composition"""
        file_path = filedialog.asksaveasfilename(
            title="Save Combined Report As", defaultextension=".pbip",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)
    
    def auto_generate_output_path(self):
        """Auto-generate output path using composition"""
        report_a = self.file_input.clean_file_path(self.report_a_path.get())
        report_b = self.file_input.clean_file_path(self.report_b_path.get())
        
        if report_a and report_b:
            try:
                output_path = self.merger_engine.generate_output_path(report_a, report_b)
                self.output_path.set(output_path)
            except Exception:
                pass
    
    def analyze_reports(self):
        """Analyze selected reports using composition architecture"""
        try:
            report_a = self.file_input.clean_file_path(self.report_a_path.get())
            report_b = self.file_input.clean_file_path(self.report_b_path.get())
            
            # Use validation component
            self.validation.validate_pbip_file(report_a, "Report A")
            self.validation.validate_pbip_file(report_b, "Report B")
            
        except Exception as e:
            self.show_error("Validation Error", str(e))
            return
        
        # Use progress component with threading component
        def analyze_target():
            return self.merger_engine.analyze_reports(report_a, report_b)
        
        def on_success(results):
            self._handle_analysis_complete(results)
        
        def on_error(error):
            self.show_error("Analysis Error", str(error))
        
        # Use composition for background processing with progress
        self.progress.run_with_progress(
            target_func=analyze_target,
            success_callback=on_success,
            error_callback=on_error
        )
    
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
        self.theme_frame.pack(fill=tk.X, pady=(0, 15))
        
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
        
        ttk.Radiobutton(selection_frame, text=f"📊 Report A: {theme_a['display']}",
                       variable=self.theme_choice, value="report_a").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(selection_frame, text=f"📊 Report B: {theme_b['display']}",
                       variable=self.theme_choice, value="report_b").pack(anchor=tk.W, pady=2)
    
    def _hide_theme_selection(self):
        """Hide theme selection UI"""
        if self.theme_frame:
            self.theme_frame.pack_forget()
    
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
        """Start merge operation using composition architecture"""
        if not self.analysis_results:
            self.show_error("Error", "Please analyze reports first")
            return
        
        try:
            output_path = self.file_input.clean_file_path(self.output_path.get())
            self.validation.validate_output_path(output_path)
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
        
        # Use composition for background merge with progress
        def merge_target():
            self.log_message("\n🚀 Starting merge operation...")
            
            report_a = self.file_input.clean_file_path(self.report_a_path.get())
            report_b = self.file_input.clean_file_path(self.report_b_path.get())
            theme_choice = self.theme_choice.get()
            
            success = self.merger_engine.merge_reports(
                report_a, report_b, output_path, theme_choice, self.analysis_results
            )
            
            return {'success': success, 'output_path': output_path}
        
        def on_success(result):
            self._handle_merge_complete(result)
        
        def on_error(error):
            self.show_error("Merge Error", str(error))
        
        # Use composition for background processing with progress
        self.progress.run_with_progress(
            target_func=merge_target,
            success_callback=on_success,
            error_callback=on_error
        )
    
    def _handle_merge_complete(self, result):
        """Handle merge completion"""
        if result['success']:
            self.log_message("✅ MERGE COMPLETED SUCCESSFULLY!")
            self.log_message(f"💾 Output: {result['output_path']}")
            
            self.show_info("Merge Complete", 
                          f"Merge completed successfully!\n\n💾 Output:\n{result['output_path']}")
        else:
            self.show_error("Merge Failed", "The merge operation failed. Check the log for details.")
    
    def reset_tool(self) -> None:
        """Reset the tool to initial state using composition"""
        if self.progress.is_busy:
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
        self._hide_theme_selection()
        
        # Clear log
        self._clear_log()
        self.log_message("✅ Report Merger reset successfully!")
    
    def show_help_dialog(self) -> None:
        """Show help dialog specific to report merger"""
        def create_help_content(help_window):
            main_frame = ttk.Frame(help_window, padding="20")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Header
            ttk.Label(main_frame, text="❓ Power BI Report Merger - Help (Composition Architecture)", 
                     font=('Segoe UI', 16, 'bold'), 
                     foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 20))
            
            # Architecture note
            arch_frame = ttk.Frame(main_frame)
            arch_frame.pack(fill=tk.X, pady=(0, 20))
            
            info_container = tk.Frame(arch_frame, bg=AppConstants.COLORS['info'], 
                                    padx=15, pady=10, relief='solid', borderwidth=2)
            info_container.pack(fill=tk.X)
            
            ttk.Label(info_container, text="🏗️  NEW COMPOSITION ARCHITECTURE", 
                     font=('Segoe UI', 12, 'bold'), 
                     background=AppConstants.COLORS['info'],
                     foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W)
            
            arch_notes = [
                "• Built using composition pattern with component-based architecture",
                "• Improved error handling with proper closure management",
                "• Enhanced modularity and testability",
                "• Thread-safe background processing with progress indication",
                "• Composable validation, file handling, and UI components"
            ]
            
            for note in arch_notes:
                ttk.Label(info_container, text=note, font=('Segoe UI', 10),
                         background=AppConstants.COLORS['info'],
                         foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W, pady=1)
            
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
                ("🚀 Quick Start", [
                    "1. Select Report A and Report B (.pbip files)",
                    "2. Click 'ANALYZE REPORTS' to validate and preview",
                    "3. Choose theme if there's a conflict",
                    "4. Click 'EXECUTE MERGE' to create combined report"
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
            
            # Close button with padding
            button_frame = ttk.Frame(help_window)
            button_frame.pack(fill=tk.X, pady=(20, 20), padx=20)
            
            ttk.Button(button_frame, text="❌ Close", command=help_window.destroy).pack()
        
        # Create help window
        help_window = tk.Toplevel(self.main_app.root)
        help_window.title("Power BI Report Merger - Help (Composition Architecture)")
        help_window.geometry("750x750")
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
    
    def _export_log(self):
        """Export log content to file"""
        try:
            log_content = self.log_text.get(1.0, tk.END)
            file_path = filedialog.asksaveasfilename(
                title="Export Log", defaultextension=".txt",
                filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"{self.tool_name} - Analysis Log (Composition Architecture)\n")
                    f.write(f"Generated by {AppConstants.COMPANY_NAME}\n")
                    f.write(f"{'='*50}\n\n")
                    f.write(log_content)
                
                self.log_message(f"✅ Log exported to: {file_path}")
                self.show_info("Export Complete", f"Log exported successfully!\n\n{file_path}")
        
        except Exception as e:
            self.show_error("Export Error", f"Failed to export log: {e}")
    
    def _clear_log(self):
        """Clear log content"""
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
            self._show_welcome_message()


# Factory function for creating the composable report merger tab
def create_composable_report_merger_tab(parent, main_app) -> ComposableReportMergerTab:
    """
    Factory function to create a composable report merger tab.
    
    Args:
        parent: Parent widget
        main_app: Main application instance
        
    Returns:
        ComposableReportMergerTab instance
    """
    tab = ComposableReportMergerTab(parent, main_app)
    
    # Setup common styling
    tab._setup_common_styling()
    
    return tab
