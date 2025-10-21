"""
PBIP Layout Optimizer UI - Main Interface Tab
Built by Reid Havens of Analytic Endeavors

This tab provides layout optimization for Power BI relationship diagrams
with integrated advanced components when available.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
from typing import Dict, Any, Optional, List

from core.ui_base import BaseToolTab
from core.constants import AppConstants
from .enhanced_layout_core import EnhancedPBIPLayoutCore


class PBIPLayoutOptimizerTab(BaseToolTab):
    """PBIP Layout Optimizer main tab"""
    
    def __init__(self, parent, main_app=None):
        # Initialize with required parameters for BaseToolTab
        super().__init__(parent, main_app, "pbip_layout_optimizer", "PBIP Layout Optimizer")
        self.selected_pbip_folder = tk.StringVar()
        
        # Initialize logger
        import logging
        self.logger = logging.getLogger("pbip_layout_optimizer_ui")
        
        # Initialize enhanced core
        self.layout_core = EnhancedPBIPLayoutCore()
        
        # Canvas settings (simplified)
        self.canvas_width = tk.IntVar(value=1400)
        self.canvas_height = tk.IntVar(value=900)
        self.use_middle_out = tk.BooleanVar(value=True)  # Always enable middle-out
        self.preview_mode = tk.BooleanVar(value=False)   # Always save changes
        
        self._create_interface()
        
    def _create_interface(self):
        """Create the main interface"""
        # Main container with scrolling - use BaseToolTab's frame
        main_frame = self.frame  # Use the frame from BaseToolTab
        
        # Create sections
        self._create_file_selection_section(main_frame)
        self._create_analysis_section(main_frame)
        self._create_action_buttons_section(main_frame)
    
    def _create_file_selection_section(self, parent):
        """Create PBIP file selection section"""
        section_frame = ttk.LabelFrame(parent, text="📁 PBIP DATA SOURCE", 
                                     style='Section.TLabelframe', padding="20")
        section_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create main content frame
        content_frame = ttk.Frame(section_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # LEFT: Instructions
        instructions_frame = ttk.Frame(content_frame)
        instructions_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 35))
        
        ttk.Label(instructions_frame, text="🚀 QUICK START GUIDE:", 
                 font=('Segoe UI', 10, 'bold'),
                 foreground=AppConstants.COLORS['info']).pack(anchor=tk.W)
        
        instructions = [
            "1. Navigate to your .pbip file in File Explorer",
            "2. Right-click the .pbip file and select 'Copy as path'",
            "3. Paste (Ctrl+V) into the path field",
            "4. Path quotes will be automatically cleaned",
            "5. Click 'Analyze Layout' to assess current diagram",
            "6. Review recommendations and click 'Optimize Layout'"
        ]
        
        for instruction in instructions:
            ttk.Label(instructions_frame, text=instruction, 
                     font=('Segoe UI', 8),
                     foreground=AppConstants.COLORS['text_secondary'], 
                     wraplength=300).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        tk.Label(instructions_frame, text="⚠️ Requires PBIP format with TMDL files", 
                 font=('Segoe UI', 9, 'italic'),
                 foreground='#d97706',
                 background='#f8fafc').pack(anchor=tk.W, padx=(10, 0), pady=(5, 0))
        
        # RIGHT: File input
        input_frame = ttk.Frame(content_frame)
        input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        input_frame.columnconfigure(1, weight=1)
        
        # PBIP File input
        ttk.Label(input_frame, text="Project File (PBIP):").grid(row=0, column=0, sticky=tk.W, pady=8)
        
        self.folder_entry = ttk.Entry(input_frame, textvariable=self.selected_pbip_folder, 
                                     font=('Segoe UI', 9), width=80)
        self.folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        
        self.browse_btn = ttk.Button(input_frame, text="📂 Browse", 
                                    command=self._browse_file)
        self.browse_btn.grid(row=0, column=2, pady=8)
        
        # Analyze button
        self.analyze_btn = ttk.Button(input_frame, text="🔍 ANALYZE LAYOUT",
                                     command=self._analyze_layout, 
                                     style='Action.TButton', state='disabled')
        self.analyze_btn.grid(row=1, column=0, columnspan=3, pady=(15, 0))
        
        # Validation status (positioned under analyze button)
        self.validation_frame = ttk.Frame(input_frame)
        self.validation_frame.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        self.validation_label = ttk.Label(self.validation_frame, text="Select a PBIP file to begin analysis",
                                         font=('Segoe UI', 9, 'italic'),
                                         foreground=AppConstants.COLORS['text_secondary'])
        self.validation_label.pack()
        
        # Bind paste event
        self.folder_entry.bind('<Control-v>', self._on_paste)
    
    def _create_analysis_section(self, parent):
        """Create analysis section"""
        # Use base class log section like Report Merger
        self.log_components = self.create_log_section(parent, "🔍 ANALYSIS & PROGRESS LOG")
        self.log_components['frame'].pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Store reference to text widget for logging
        self.analysis_text = self.log_components['text_widget']
        
        # Initially show welcome message
        self._log_message("🎯 Welcome to the PBIP Layout Optimizer!")
        self._log_message("🔗 Select a PBIP folder to analyze and optimize your diagram layout...")
        self._log_message("🎯 Apply Haven's middle-out design philosophy for professional layouts")
        self._log_message("")
        
        # Show component status only if basic functionality
        if not self.layout_core.mcp_available:
            self._log_message("⚠️ Using basic functionality - Enhanced components not available")
        self._log_message("")
    
    def _create_action_buttons_section(self, parent):
        """Create action buttons section"""
        # Create button frame
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=(20, 0))
        
        # Main action button (like EXECUTE MERGE)
        self.optimize_btn = ttk.Button(button_frame, text="🎯 OPTIMIZE LAYOUT", 
                                      command=self._optimize_layout,
                                      style='Action.TButton',
                                      state='disabled')
        self.optimize_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        # Reset button (like RESET ALL)
        ttk.Button(button_frame, text="🔄 RESET ALL", 
                  command=self._reset_interface,
                  style='Secondary.TButton').pack(side=tk.LEFT)
    
    def _browse_file(self):
        """Browse for PBIP file"""
        file_path = filedialog.askopenfilename(
            title="Select PBIP File",
            initialdir=str(Path(self.selected_pbip_folder.get()).parent) if self.selected_pbip_folder.get() else str(Path.home()),
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.selected_pbip_folder.set(file_path)
            self._validate_folder(file_path)
    
    def _extract_folder_from_pbip_file(self, file_path: str) -> Optional[str]:
        """Extract the folder path from a .pbip file path"""
        try:
            file_path_obj = Path(file_path)
            
            # Check if the path is a .pbip file
            if not file_path_obj.suffix.lower() == '.pbip':
                return None
            
            # Check if file exists
            if not file_path_obj.exists():
                return None
            
            # The folder is the parent directory of the .pbip file
            folder_path = file_path_obj.parent
            
            # Verify it's a valid PBIP folder structure
            # Look for .SemanticModel folder in the same directory
            semantic_folders = list(folder_path.glob("*.SemanticModel"))
            if not semantic_folders:
                return None
            
            return str(folder_path)
            
        except Exception as e:
            self.logger.error(f"Error extracting folder from PBIP file: {e}")
            return None
    
    def _on_paste(self, event):
        """Handle paste event to clean path quotes"""
        # Schedule validation after paste
        self.master.after(100, lambda: self._validate_folder(self.selected_pbip_folder.get()))
    
    def _validate_folder(self, file_path: str):
        """Validate selected PBIP file and extract folder path"""
        # Clean quotes from path
        file_path = file_path.strip().strip('"').strip("'")
        if file_path != self.selected_pbip_folder.get():
            self.selected_pbip_folder.set(file_path)
        
        if not file_path:
            return
        
        # Extract folder path from .pbip file
        folder_path = self._extract_folder_from_pbip_file(file_path)
        if not folder_path:
            self.validation_label.configure(
                text="❌ Invalid PBIP file path",
                foreground=AppConstants.COLORS['error']
            )
            self.analyze_btn.configure(state='disabled')
            self.optimize_btn.configure(state='disabled')
            return
        
        validation = self.layout_core.validate_pbip_folder(folder_path)
        
        if validation['valid']:
            # Get table count for display
            tmdl_files = self.layout_core.find_tmdl_files(folder_path)
            table_count = len(self.layout_core.get_table_names_from_tmdl(tmdl_files))
            tmdl_count = len(tmdl_files) - 1  # Exclude model.tmdl
            
            self.validation_label.configure(
                text=f"✅ Valid PBIP file\n📋 Tables found: {table_count}\n📄 TMDL files: {tmdl_count}",
                foreground=AppConstants.COLORS['success']
            )
            
            # Enable analysis button
            self.analyze_btn.configure(state='normal')
            
        else:
            self.validation_label.configure(
                text=f"❌ {validation['error']}",
                foreground=AppConstants.COLORS['error']
            )
            
            # Disable buttons
            self.analyze_btn.configure(state='disabled')
            self.optimize_btn.configure(state='disabled')
    
    def _analyze_layout(self):
        """Analyze layout quality"""
        file_path = self.selected_pbip_folder.get()
        if not file_path:
            messagebox.showwarning("Error", "Please select a valid PBIP file first.")
            return
        
        # Extract folder path from file path
        folder_path = self._extract_folder_from_pbip_file(file_path)
        if not folder_path:
            messagebox.showerror("Error", "Invalid PBIP file. Please select a valid .pbip file.")
            return
        
        self.analyze_btn.configure(state='disabled')
        
        # Run analysis in background
        self.run_in_background(
            target_func=lambda: self._analysis_thread_target(folder_path),
            success_callback=self._handle_analysis_result,
            error_callback=lambda e: self._handle_analysis_error(str(e))
        )
    
    def _analysis_thread_target(self, folder_path: str):
        """Background analysis logic"""
        self.update_progress(10, "Validating PBIP folder...")
        
        self.update_progress(30, "Reading table definitions...")
        
        self.update_progress(50, "Analyzing layout quality...")
        layout_result = self.layout_core.analyze_layout_quality(folder_path)
        
        # If advanced components are available, also get table categorization
        if self.layout_core.mcp_available:
            self.update_progress(70, "Categorizing tables...")
            categorization_result = self.layout_core.analyze_table_categorization(folder_path)
            if categorization_result.get('success'):
                layout_result['categorization'] = categorization_result.get('categorization', {})
                layout_result['mcp_extensions'] = categorization_result.get('extensions', [])
        
        self.update_progress(90, "Calculating quality score...")
        
        self.update_progress(100, "Analysis complete!")
        return layout_result
    
    def _handle_analysis_result(self, result: Dict[str, Any]):
        """Handle analysis result"""
        if result['success']:
            # Log analysis results to the single log area
            self._log_analysis_results(result)
            
            # Enable optimization button
            self.optimize_btn.configure(state='normal')
            
        else:
            self._log_message(f"❌ Analysis failed: {result.get('error', 'Unknown error')}")
        
        self.analyze_btn.configure(state='normal')
    
    def _handle_analysis_error(self, error: str):
        """Handle analysis error"""
        self.analyze_btn.configure(state='normal')
        self._log_message(f"❌ Analysis Error: {error}")
    
    def _optimize_layout(self):
        """Optimize layout"""
        file_path = self.selected_pbip_folder.get()
        if not file_path:
            messagebox.showwarning("Error", "Please run analysis first.")
            return
        
        # Extract folder path from file path
        folder_path = self._extract_folder_from_pbip_file(file_path)
        if not folder_path:
            messagebox.showerror("Error", "Invalid PBIP file. Please select a valid .pbip file.")
            return
        
        self.optimize_btn.configure(state='disabled')
        
        # Run optimization in background
        self.run_in_background(
            target_func=lambda: self._optimization_thread_target(folder_path),
            success_callback=self._handle_optimization_result,
            error_callback=lambda e: self._handle_optimization_error(str(e))
        )
    
    def _optimization_thread_target(self, folder_path: str):
        """Background optimization logic"""
        self.update_progress(10, "Preparing layout optimization...")
        
        self.update_progress(30, "Analyzing table relationships...")
        
        self.update_progress(60, "Generating optimal positions...")
        
        # Always save changes (no preview mode in simplified version)
        save_changes = True
        
        # Always use middle-out design (no checkbox in simplified version)
        use_middle_out = True
        
        # Use enhanced optimization with spacing setting
        result = self.layout_core.optimize_layout(
            folder_path,
            self.canvas_width.get(),
            self.canvas_height.get(),
            save_changes,
            use_middle_out
        )
        
        self.update_progress(100, "Optimization complete!")
        return result
    
    def _handle_optimization_result(self, result: Dict[str, Any]):
        """Handle optimization result"""
        if result['success']:
            # Log optimization results to the single log area
            self._log_optimization_results(result)
            
        else:
            self._log_message(f"❌ Optimization failed: {result.get('error', 'Unknown error')}")
        
        self.optimize_btn.configure(state='normal')
    
    def _handle_optimization_error(self, error: str):
        """Handle optimization error"""
        self.optimize_btn.configure(state='normal')
        self._log_message(f"❌ Optimization Error: {error}")
    
    def _log_analysis_results(self, result: Dict[str, Any]):
        """Log analysis results to the analysis text area"""
        self._log_message("📊 LAYOUT ANALYSIS COMPLETE:")
        self._log_message("=" * 50)
        
        # Basic analysis info
        analysis = result.get('layout_analysis', {})
        if analysis:
            self._log_message(f"📈 Quality Score: {result.get('quality_score', 0)}/100 ({result.get('rating', 'Unknown')})")
            self._log_message(f"📋 Total Tables: {analysis.get('total_tables', 0)}")
            self._log_message(f"📍 Positioned Tables: {analysis.get('positioned_tables', 0)}")
            self._log_message(f"⚠️ Overlapping Tables: {analysis.get('overlapping_tables', 0)}")
            self._log_message(f"🔄 Average Spacing: {analysis.get('average_spacing', 0):.1f}px")
        
        # Advanced categorization if available
        categorization = result.get('categorization', {})
        if categorization:
            self._log_message("")
            self._log_message("🏷️ TABLE CATEGORIZATION (Advanced):")
            self._log_message("=" * 40)
            
            # Facts - single line
            fact_info = categorization.get('fact_tables', {})
            if fact_info.get('count', 0) > 0:
                self._log_message(f"📊 FACT TABLES: {fact_info['count']}")
            
            # Dimensions - single line each
            dim_info = categorization.get('dimension_tables', {})
            for level, count_key in [('L1', 'l1_count'), ('L2', 'l2_count'), ('L3', 'l3_count'), ('L4+', 'l4_plus_count')]:
                count = dim_info.get(count_key, 0)
                if count > 0:
                    self._log_message(f"📁 {level} DIMENSIONS: {count}")
            
            # Special tables - single line each
            special_info = categorization.get('special_tables', {})
            if special_info.get('calendar_count', 0) > 0:
                self._log_message(f"📅 CALENDAR TABLES: {special_info['calendar_count']}")
            if special_info.get('metrics_count', 0) > 0:
                self._log_message(f"📊 METRICS TABLES: {special_info['metrics_count']}")
            if special_info.get('parameter_count', 0) > 0:
                self._log_message(f"⚙️ PARAMETER TABLES: {special_info['parameter_count']}")
            
            # Extensions - single line
            extensions = result.get('mcp_extensions', [])
            if extensions:
                self._log_message(f"🔗 TABLE EXTENSIONS: {len(extensions)}")
        
        self._log_message("")
        self._log_message("✅ Analysis complete! Ready for optimization.")
    
    def _log_optimization_results(self, result: Dict[str, Any]):
        """Log optimization results to the single log area"""
        self._log_message("")
        self._log_message("🎯 LAYOUT OPTIMIZATION COMPLETE:")
        self._log_message("=" * 50)
        
        # Basic optimization info
        method = result.get('layout_method', 'Enhanced')
        tables_arranged = result.get('tables_arranged', 0)
        changes_saved = result.get('changes_saved', False)
        
        self._log_message(f"🎯 Layout Method: {method}")
        self._log_message(f"📊 Tables Arranged: {tables_arranged}")
        self._log_message(f"💾 Changes Saved: {'Yes' if changes_saved else 'No'}")
        self._log_message(f"📌 Canvas Size: {self.canvas_width.get()}x{self.canvas_height.get()}")
        self._log_message(f"📆 Layout Design: Middle-Out Philosophy")
        
        # Layout features if available
        if result.get('layout_features'):
            self._log_message("")
            self._log_message("🔧 Layout Features:")
            features = result['layout_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                self._log_message(f"  {status} {feature.replace('_', ' ').title()}")
        
        # Advanced features if available
        if result.get('advanced_features'):
            self._log_message("")
            self._log_message("⚡ Advanced Features:")
            features = result['advanced_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                self._log_message(f"  {status} {feature.replace('_', ' ').title()}")
        
        self._log_message("")
        self._log_message("✅ Layout optimization completed successfully!")
        if changes_saved:
            self._log_message("💾 Your diagram layout has been updated.")
        else:
            self._log_message("📝 Preview mode - no changes were saved.")
    
    def _log_message(self, message: str):
        """Add message to analysis log - use base class method"""
        self.log_message(message)
    
    def _update_optimization_display(self, result: Dict[str, Any]):
        """Update optimization results display"""
        # Clear placeholder
        for widget in self.optimization_results_frame.winfo_children():
            widget.destroy()
        
        # Create results summary
        summary_frame = ttk.Frame(self.optimization_results_frame)
        summary_frame.pack(fill=tk.X)
        
        ttk.Label(summary_frame, text="✅ Layout Optimization Complete", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['success']).pack(anchor=tk.W)
        
        # Results grid
        results_grid = ttk.Frame(summary_frame)
        results_grid.pack(fill=tk.X, pady=(10, 0))
        
        # Left column
        left_frame = ttk.Frame(results_grid)
        left_frame.pack(side=tk.LEFT, anchor=tk.NW)
        
        # Calculate tables arranged
        tables_arranged = result.get('tables_arranged', 0)
        
        ttk.Label(left_frame, text=f"📊 Tables Arranged: {tables_arranged}", 
                 font=('Segoe UI', 9)).pack(anchor=tk.W)
        ttk.Label(left_frame, text=f"💾 Changes Saved: {'Yes' if result.get('changes_saved') else 'No'}", 
                 font=('Segoe UI', 9)).pack(anchor=tk.W)
        ttk.Label(left_frame, text=f"📐 Canvas Size: {self.canvas_width.get()}x{self.canvas_height.get()}", 
                 font=('Segoe UI', 9)).pack(anchor=tk.W)
        
        # Right column
        right_frame = ttk.Frame(results_grid)
        right_frame.pack(side=tk.LEFT, anchor=tk.NW, padx=(50, 0))
        
        layout_method = result.get('layout_method', 'Enhanced')
        ttk.Label(right_frame, text=f"🎯 Layout Method: {layout_method}", 
                 font=('Segoe UI', 9)).pack(anchor=tk.W)
        
        if result.get('advanced_features'):
            ttk.Label(right_frame, text="⚡ Advanced Features: Enabled", 
                     font=('Segoe UI', 9)).pack(anchor=tk.W)
        else:
            ttk.Label(right_frame, text="⚡ Mode: Basic Layout", 
                     font=('Segoe UI', 9)).pack(anchor=tk.W)
    
    def _format_analysis_results(self, result: Dict[str, Any]) -> str:
        """Format analysis results for detailed display"""
        text = "LAYOUT ANALYSIS RESULTS\n"
        text += "=" * 50 + "\n\n"
        
        text += f"PBIP Folder: {result.get('pbip_folder', '')}\n"
        text += f"Semantic Model: {result.get('semantic_model_path', '')}\n\n"
        
        # Quality assessment
        text += f"QUALITY ASSESSMENT:\n"
        text += f"Overall Score: {result.get('quality_score', 0)}/100 ({result.get('rating', 'Unknown')})\n\n"
        
        # Table statistics
        analysis = result.get('layout_analysis', {})
        if analysis:
            text += f"TABLE STATISTICS:\n"
            text += f"Total Tables: {analysis.get('total_tables', 0)}\n"
            text += f"Positioned Tables: {analysis.get('positioned_tables', 0)}\n"
            text += f"Overlapping Tables: {analysis.get('overlapping_tables', 0)}\n"
            text += f"Average Spacing: {analysis.get('average_spacing', 0):.1f}px\n\n"
        
        # Add advanced categorization if available
        if result.get('categorization') and self.layout_core.mcp_available:
            text += self._format_categorization_in_analysis(result['categorization'], result.get('mcp_extensions', []))
        
        if result.get('recommendations'):
            text += f"RECOMMENDATIONS:\n"
            for i, rec in enumerate(result['recommendations'], 1):
                text += f"{i}. {rec}\n"
            text += "\n"
        
        if result.get('table_names'):
            text += f"SAMPLE TABLES (first 10):\n"
            for table in result['table_names']:
                text += f"• {table}\n"
        
        return text
    
    def _format_categorization_in_analysis(self, categorization: Dict[str, Any], extensions: List[Dict[str, Any]]) -> str:
        """Format categorization data for inclusion in analysis results"""
        text = f"🏷️ TABLE CATEGORIZATION (Advanced Enhanced):\n"
        text += "=" * 40 + "\n"
        
        # Facts
        fact_info = categorization.get('fact_tables', {})
        fact_count = fact_info.get('count', 0)
        if fact_count > 0:
            text += f"📊 FACT TABLES ({fact_count}):\n"
            for table in fact_info.get('tables', [])[:5]:  # Show first 5
                text += f"   • {table}\n"
            if len(fact_info.get('tables', [])) > 5:
                text += f"   ... and {len(fact_info.get('tables', [])) - 5} more\n"
            text += "\n"
        
        # Dimensions by level
        dim_info = categorization.get('dimension_tables', {})
        for level, count_key, tables_key in [
            ('L1', 'l1_count', 'l1_tables'),
            ('L2', 'l2_count', 'l2_tables'), 
            ('L3', 'l3_count', 'l3_tables'),
            ('L4+', 'l4_plus_count', 'l4_plus_tables')
        ]:
            count = dim_info.get(count_key, 0)
            if count > 0:
                text += f"📁 {level} DIMENSION TABLES ({count}):\n"
                tables = dim_info.get(tables_key, [])
                for table in tables[:5]:  # Show first 5
                    text += f"   • {table}\n"
                if len(tables) > 5:
                    text += f"   ... and {len(tables) - 5} more\n"
                text += "\n"
        
        # Special tables
        special_info = categorization.get('special_tables', {})
        special_types = [
            ('📅 CALENDAR', 'calendar_count', 'calendar_tables'),
            ('📊 METRICS', 'metrics_count', 'metrics_tables'),
            ('⚙️ PARAMETERS', 'parameter_count', 'parameter_tables'),
            ('🧮 CALCULATION GROUPS', 'calculation_groups_count', 'calculation_groups')
        ]
        
        for label, count_key, tables_key in special_types:
            count = special_info.get(count_key, 0)
            if count > 0:
                text += f"{label} ({count}):\n"
                tables = special_info.get(tables_key, [])
                for table in tables[:3]:  # Show first 3
                    text += f"   • {table}\n"
                if len(tables) > 3:
                    text += f"   ... and {len(tables) - 3} more\n"
                text += "\n"
        
        # Extensions
        if extensions:
            text += f"🔗 TABLE EXTENSIONS ({len(extensions)}):\n"
            for ext in extensions[:3]:  # Show first 3
                text += f"   • {ext['extension_table']} → {ext['base_table']}\n"
            if len(extensions) > 3:
                text += f"   ... and {len(extensions) - 3} more\n"
            text += "\n"
        
        # Disconnected and excluded
        disconnected_count = categorization.get('disconnected_tables', {}).get('count', 0)
        excluded_count = categorization.get('excluded_tables', {}).get('auto_date_count', 0)
        
        if disconnected_count > 0:
            text += f"🔌 DISCONNECTED TABLES: {disconnected_count}\n"
        if excluded_count > 0:
            text += f"🗓️ AUTO DATE TABLES (excluded): {excluded_count}\n"
        
        text += "\n"
        return text
    
    def _format_optimization_results(self, result: Dict[str, Any]) -> str:
        """Format optimization results for detailed display"""
        text = "LAYOUT OPTIMIZATION RESULTS\n"
        text += "=" * 50 + "\n\n"
        
        text += f"Operation: {result.get('operation', 'Layout Optimization')}\n"
        text += f"PBIP Folder: {result.get('pbip_folder', '')}\n"
        text += f"Layout Method: {result.get('layout_method', 'Enhanced')}\n\n"
        
        text += f"OPTIMIZATION RESULTS:\n"
        text += f"Tables Arranged: {result.get('tables_arranged', 0)}\n"
        text += f"Changes Saved: {'Yes' if result.get('changes_saved') else 'No'}\n"
        text += f"Canvas Size: {self.canvas_width.get()}x{self.canvas_height.get()}\n\n"
        
        # Layout features
        if result.get('layout_features'):
            text += f"LAYOUT FEATURES:\n"
            features = result['layout_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                text += f"  {status} {feature.replace('_', ' ').title()}\n"
            text += "\n"
        
        # Advanced features
        if result.get('advanced_features'):
            text += f"ADVANCED FEATURES:\n"
            features = result['advanced_features']
            for feature, enabled in features.items():
                status = "✅" if enabled else "❌"
                text += f"  {status} {feature.replace('_', ' ').title()}\n"
            text += "\n"
        
        # Enhancement status
        text += f"COMPONENT STATUS:\n"
        text += f"  ✅ Advanced Components: {'Available' if self.layout_core.mcp_available else 'Not Available'}\n"
        text += f"  ✅ Middle-Out Design: {'Enabled' if self.use_middle_out.get() else 'Disabled'}\n"
        text += f"  ✅ Table Categorization: {'Enabled' if self.layout_core.mcp_available else 'Basic'}\n"
        
        return text
    
    def _update_results_text(self, text: str):
        """Update the detailed results text area"""
        self.results_text.configure(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, text)
        self.results_text.configure(state=tk.DISABLED)
    
    def _export_log(self):
        """Export analysis log to file"""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Export Analysis Log",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if file_path:
                log_content = self.analysis_text.get(1.0, tk.END)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                
                messagebox.showinfo("Export Complete", f"Log exported to {file_path}")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export log: {str(e)}")
    
    def _clear_log(self):
        """Clear the analysis log"""
        self.analysis_text.configure(state=tk.NORMAL)
        self.analysis_text.delete(1.0, tk.END)
        self.analysis_text.configure(state=tk.DISABLED)
        
        self._log_message("📝 PBIP Layout Optimizer - Log Cleared")
        self._log_message("🚀 Ready for new analysis...")
        self._log_message("")
    
    def _reset_interface(self):
        """Reset the entire interface"""
        # Reset folder selection
        self.selected_pbip_folder.set("")
        
        # Reset validation
        self.validation_label.configure(
            text="Select a PBIP file to begin analysis",
            foreground=AppConstants.COLORS['text_secondary']
        )
        
        # Reset buttons
        self.analyze_btn.configure(state='disabled')
        self.optimize_btn.configure(state='disabled')
    
    # Required abstract method implementations for BaseToolTab
    def setup_ui(self):
        """Setup UI - already done in _create_interface"""
        pass  # Interface is created in __init__
    
    def reset_tab(self):
        """Reset tab to initial state"""
        self._reset_interface()
    
    def show_help_dialog(self):
        """Show context-sensitive help dialog"""
        # Get the correct parent window - try multiple possible parent references
        parent_window = None
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'root'):
            parent_window = self.main_app.root
        elif hasattr(self, 'master'):
            parent_window = self.master
        elif hasattr(self, 'parent'):
            parent_window = self.parent
        else:
            # Fallback - find root window from frame
            parent_window = self.frame.winfo_toplevel()
        
        help_window = tk.Toplevel(parent_window)
        help_window.title("PBIP Layout Optimizer - Help")
        help_window.geometry("1000x830")  # Wider and shorter
        help_window.resizable(False, False)
        help_window.transient(parent_window)
        help_window.grab_set()
        
        # Center window
        help_window.geometry(f"+{parent_window.winfo_rootx() + 50}+{parent_window.winfo_rooty() + 50}")
        
        self._create_help_content(help_window)
    
    def _create_help_content(self, help_window):
        """Create help content for the layout optimizer"""
        help_window.configure(bg=AppConstants.COLORS['background'])
        
        # Main container
        container = ttk.Frame(help_window, padding="20")
        container.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(container, text="📊 PBIP Layout Optimizer Help", 
                 font=('Segoe UI', 16, 'bold'), 
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 15))
        
        # Orange warning box
        warning_frame = ttk.Frame(container)
        warning_frame.pack(fill=tk.X, pady=(0, 15))
        
        warning_container = tk.Frame(warning_frame, bg=AppConstants.COLORS['warning'], 
                                   padx=15, pady=10, relief='solid', borderwidth=2)
        warning_container.pack(fill=tk.X)
        
        ttk.Label(warning_container, text="⚠️  IMPORTANT DISCLAIMERS & REQUIREMENTS", 
                 font=('Segoe UI', 12, 'bold'), 
                 background=AppConstants.COLORS['warning'],
                 foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W)
        
        warnings = [
            "• This tool ONLY works with PBIP format files (.pbip folders)",
            "• This is NOT officially supported by Microsoft - use at your own discretion",
            "• Requires TMDL files in semantic model definition folder",
            "• Always keep backups of your original reports before optimization",
            "• Test thoroughly and validate optimized layouts before production use",
            "• Enable 'Store reports using enhanced metadata format (PBIP)' in Power BI Desktop"
        ]
        
        for warning in warnings:
            ttk.Label(warning_container, text=warning, font=('Segoe UI', 10),
                     background=AppConstants.COLORS['warning'],
                     foreground=AppConstants.COLORS['surface']).pack(anchor=tk.W, pady=1)
        
        # Top sections in 2-column layout
        top_sections_frame = ttk.Frame(container)
        top_sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        top_sections_frame.columnconfigure(0, weight=1)
        top_sections_frame.columnconfigure(1, weight=1)
        
        # LEFT COLUMN TOP: What This Tool Does
        left_top_frame = ttk.Frame(top_sections_frame)
        left_top_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(left_top_frame, text="🎯 What This Tool Does", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        what_items = [
            "✅ Analyzes your current relationship diagram layout quality",
            "✅ Provides layout quality scoring (0-100 scale with ratings)",
            "✅ Categorizes tables by type (Facts, Dimensions L1-L4+, Special tables)",
            "✅ Applies Haven's middle-out design philosophy for professional layouts",
            "✅ Positions tables based on relationships and hierarchy",
            "✅ Optimizes spacing and reduces overlapping elements",
            "✅ Provides detailed analysis scoring and recommendations"
        ]
        
        for item in what_items:
            ttk.Label(left_top_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # RIGHT COLUMN TOP: File Requirements
        right_top_frame = ttk.Frame(top_sections_frame)
        right_top_frame.grid(row=0, column=1, sticky=(tk.N, tk.W, tk.E), padx=(10, 0))
        
        ttk.Label(right_top_frame, text="📁 File Requirements", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        file_items = [
            "✅ Only .pbip format files (.pbip folders) are supported",
            "✅ Must contain semantic model definition folder with TMDL files",
            "✅ Requires diagramLayout.json file for layout data",
            "✅ Write permissions to PBIP folder (for saving changes)",
            "❌ Legacy .pbix files are NOT supported",
            "❌ Reports without TMDL files cannot be optimized"
        ]
        
        for item in file_items:
            ttk.Label(right_top_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # Bottom sections in 2-column layout
        bottom_sections_frame = ttk.Frame(container)
        bottom_sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        bottom_sections_frame.columnconfigure(0, weight=1)
        bottom_sections_frame.columnconfigure(1, weight=1)
        
        # LEFT COLUMN BOTTOM: Haven's Middle-Out Design Philosophy
        left_bottom_frame = ttk.Frame(bottom_sections_frame)
        left_bottom_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(left_bottom_frame, text="🎯 Haven's Middle-Out Design Philosophy", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        philosophy_items = [
            "✅ Fact tables positioned centrally as the data foundation",
            "✅ L1 dimensions arranged around facts for direct relationships",
            "✅ L2+ dimensions positioned in outer layers by hierarchy",
            "✅ Special tables (Calendar, Parameters, Metrics) grouped logically",
            "✅ Optimized spacing prevents overlapping and improves readability",
            "✅ Professional appearance suitable for executive presentations"
        ]
        
        for item in philosophy_items:
            ttk.Label(left_bottom_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # RIGHT COLUMN BOTTOM: Important Notes
        right_bottom_frame = ttk.Frame(bottom_sections_frame)
        right_bottom_frame.grid(row=0, column=1, sticky=(tk.N, tk.W, tk.E), padx=(10, 0))
        
        ttk.Label(right_bottom_frame, text="⚠️ Important Notes", 
                 font=('Segoe UI', 12, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 5))
        
        notes_items = [
            "• ONLY works with PBIP format (not .pbix files)",
            "• This tool is NOT officially supported by Microsoft",
            "• Always backup your .pbip files before optimization",
            "• The tool modifies diagramLayout.json in your PBIP folder",
            "• Test the optimized layout in Power BI Desktop before sharing",
            "• Large models may take several minutes to analyze and optimize"
        ]
        
        for item in notes_items:
            ttk.Label(right_bottom_frame, text=item, 
                     font=('Segoe UI', 10),
                     foreground=AppConstants.COLORS['text_primary'],
                     wraplength=450).pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        # Button frame at bottom
        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X, pady=(10, 0), side=tk.BOTTOM)
        
        ttk.Button(button_frame, text="❌ Close", 
                  command=help_window.destroy,
                  style='Action.TButton').pack(pady=(5, 0))
        
        help_window.bind('<Escape>', lambda event: help_window.destroy())
