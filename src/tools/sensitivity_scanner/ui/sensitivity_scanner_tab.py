"""
Sensitivity Scanner UI Tab - Main user interface.

This module provides the SensitivityScannerTab class which implements
the user interface for the Sensitivity Scanner tool.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime

from core.ui_base import BaseToolTab, FileInputMixin, ValidationMixin
from core.constants import AppConstants
from tools.sensitivity_scanner.logic.tmdl_scanner import TmdlScanner
from tools.sensitivity_scanner.logic.models import ScanResult, Finding, RiskLevel


class SensitivityScannerTab(BaseToolTab, FileInputMixin, ValidationMixin):
    """
    Main UI tab for the Sensitivity Scanner tool.
    
    Provides:
    - File input for PBIP selection
    - Scan configuration options
    - Results display with risk levels
    - Export functionality
    - Detailed findings view
    """
    
    def __init__(self, parent, main_app):
        """
        Initialize the Sensitivity Scanner tab.
        
        Args:
            parent: Parent widget (notebook)
            main_app: Main application instance
        """
        super().__init__(parent, main_app, "sensitivity_scanner", "Sensitivity Scanner")
        
        # UI Variables
        self.pbip_path = tk.StringVar()
        self.scan_mode = tk.StringVar(value="full")  # full, tables, roles, expressions
        
        # Scanner components
        self.scanner: Optional[TmdlScanner] = None
        self.scan_result: Optional[ScanResult] = None
        
        # UI Components (will be created in setup_ui)
        self.file_input = None
        self.results_tree = None
        self.export_button = None
        self.details_text = None
        
        # Setup UI
        self.setup_ui()
    
    def setup_ui(self) -> None:
        """Setup the complete UI layout."""
        # Configure grid weights for responsive layout
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(1, weight=3)  # Results section (3x)
        self.frame.rowconfigure(2, weight=2)  # Details/Log section gets more (2x)
        
        # Row 0: Combined File Input + Scan Options Section
        self._create_file_input_section()
        
        # Row 1: Results Section (expandable)
        self._create_results_section()
        
        # Row 2: Finding Details and Scan Log (side-by-side, expandable)
        self._create_details_and_log_section()
        
        # Row 3: Progress Bar
        self.create_progress_bar(self.frame)
        
        # Row 4: Action Buttons
        self._create_action_buttons()
        
        # Welcome message
        self.log_message("🔍 Welcome to Sensitivity Scanner!")
        self.log_message("📊 This tool scans Power BI semantic models for sensitive data")
        self.log_message("⚠️  Remember: This is STATIC ANALYSIS of model structure, not data values")
    
    def _create_file_input_section(self):
        """Create the file input section."""
        # Create combined section for file input AND scan options
        combined_frame = ttk.LabelFrame(
            self.frame,
            text="📁 SENSITIVITY SCANNER SETUP",
            style='Section.TLabelframe',
            padding="15"
        )
        combined_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        combined_frame.columnconfigure(1, weight=1)
        
        # Row 0: File selection
        ttk.Label(
            combined_frame,
            text="PBIP File:",
            font=('Segoe UI', 9, 'bold')
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # File path entry
        self.pbip_path = tk.StringVar()
        file_entry = ttk.Entry(
            combined_frame,
            textvariable=self.pbip_path,
            font=('Segoe UI', 9),
            state='readonly'
        )
        file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Setup path cleaning
        self.setup_path_cleaning(self.pbip_path)
        
        # Browse button
        ttk.Button(
            combined_frame,
            text="📂 Browse",
            command=self._browse_pbip_file,
            style='Secondary.TButton',
            width=12
        ).grid(row=0, column=2, padx=(0, 10))
        
        # Scan button (narrower)
        ttk.Button(
            combined_frame,
            text="🔍 SCAN",
            command=self._execute_scan,
            style='Action.TButton',
            width=12
        ).grid(row=0, column=3)
        
        # Row 1: Scan mode options (horizontal radio buttons)
        ttk.Label(
            combined_frame,
            text="Scan Mode:",
            font=('Segoe UI', 9, 'bold')
        ).grid(row=1, column=0, sticky=tk.W, pady=(10, 0), padx=(0, 10))
        
        radio_frame = ttk.Frame(combined_frame)
        radio_frame.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(10, 0))
        
        # Row 2: Pattern Manager button
        ttk.Button(
            combined_frame,
            text="⚙️ MANAGE RULES",
            command=self._open_pattern_manager,
            style='Info.TButton',
            width=18
        ).grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        
        # Configure radiobutton style to remove gray background
        style = ttk.Style()
        style.configure('Clean.TRadiobutton', 
                       background=AppConstants.COLORS['background'],
                       foreground=AppConstants.COLORS['text_primary'])
        
        modes = [
            ("full", "🔍 Full Scan"),
            ("tables", "📊 Tables Only"),
            ("roles", "🔒 RLS Only"),
            ("expressions", "💫 Expressions Only")
        ]
        
        for idx, (value, label) in enumerate(modes):
            ttk.Radiobutton(
                radio_frame,
                text=label,
                variable=self.scan_mode,
                value=value,
                style='Clean.TRadiobutton'
            ).pack(side=tk.LEFT, padx=(0, 15))
    
    def _browse_pbip_file(self):
        """Browse for PBIP file."""
        file_path = filedialog.askopenfilename(
            title="Select PBIP File",
            filetypes=[("Power BI Project Files", "*.pbip"), ("All Files", "*.*")]
        )
        if file_path:
            self.pbip_path.set(file_path)
    

    def _create_results_section(self):
        """Create results display section."""
        results_frame = ttk.LabelFrame(
            self.frame,
            text="📊 SCAN RESULTS",
            style='Section.TLabelframe',
            padding="15"
        )
        results_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)  # Allow tree to expand
        
        # Summary label
        self.summary_label = ttk.Label(
            results_frame,
            text="No scan performed yet",
            font=('Segoe UI', 10),
            foreground=AppConstants.COLORS['text_secondary']
        )
        self.summary_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Create tree view for results
        tree_frame = ttk.Frame(results_frame)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)  # Allow tree to expand
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Tree view with columns - Set reasonable minimum height
        self.results_tree = ttk.Treeview(
            tree_frame,
            columns=("risk", "file", "location", "pattern", "confidence"),
            show="tree headings",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            height=9  # Minimum height, but will expand with weight
        )
        
        # Configure scrollbars
        vsb.config(command=self.results_tree.yview)
        hsb.config(command=self.results_tree.xview)
        
        # Define columns
        self.results_tree.heading("#0", text="Finding", anchor=tk.W)
        self.results_tree.heading("risk", text="Risk", anchor=tk.CENTER)
        self.results_tree.heading("file", text="File Type", anchor=tk.W)
        self.results_tree.heading("location", text="Location", anchor=tk.W)
        self.results_tree.heading("pattern", text="Pattern", anchor=tk.W)
        self.results_tree.heading("confidence", text="Confidence", anchor=tk.CENTER)
        
        # Column widths
        self.results_tree.column("#0", width=200, minwidth=150)
        self.results_tree.column("risk", width=80, minwidth=60, anchor=tk.CENTER)
        self.results_tree.column("file", width=120, minwidth=100)
        self.results_tree.column("location", width=200, minwidth=150)
        self.results_tree.column("pattern", width=150, minwidth=100)
        self.results_tree.column("confidence", width=70, minwidth=60, anchor=tk.CENTER)
        
        # Grid tree and scrollbars
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Bind selection event
        self.results_tree.bind('<<TreeviewSelect>>', self._on_finding_selected)
        
        # Export button
        button_frame = ttk.Frame(results_frame)
        button_frame.grid(row=2, column=0, pady=(10, 0))
        
        self.export_button = ttk.Button(
            button_frame,
            text="📄 EXPORT REPORT",
            command=self._export_report,
            style='Secondary.TButton',
            state=tk.DISABLED
        )
        self.export_button.pack()
    
    def _create_details_and_log_section(self):
        """Create finding details and scan log sections side-by-side."""
        container_frame = ttk.Frame(self.frame)
        container_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        container_frame.columnconfigure(0, weight=1)
        container_frame.columnconfigure(1, weight=1)
        container_frame.rowconfigure(0, weight=1)  # Allow vertical expansion
        
        # Left: Finding Details
        details_frame = ttk.LabelFrame(
            container_frame,
            text="📝 FINDING DETAILS",
            style='Section.TLabelframe',
            padding="15"
        )
        details_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        details_frame.columnconfigure(0, weight=1)
        details_frame.rowconfigure(0, weight=1)  # Allow text widget to expand
        
        from tkinter import scrolledtext
        self.details_text = scrolledtext.ScrolledText(
            details_frame,
            height=6,  # Minimum height, will expand with weight
            font=('Segoe UI', 9),
            wrap=tk.WORD,
            state=tk.DISABLED,
            relief='solid',
            borderwidth=1
        )
        self.details_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure text tags for formatting
        self.details_text.tag_configure('header', font=('Segoe UI', 10, 'bold'))
        self.details_text.tag_configure('high_risk', foreground='#dc2626', font=('Segoe UI', 9, 'bold'))
        self.details_text.tag_configure('medium_risk', foreground='#d97706', font=('Segoe UI', 9, 'bold'))
        self.details_text.tag_configure('low_risk', foreground='#059669', font=('Segoe UI', 9, 'bold'))
        
        # Right: Scan Log
        log_components = self.create_log_section(container_frame, "📋 SCAN LOG")
        log_components['frame'].grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        self.log_text = log_components['text_widget']
    
    def _create_action_buttons(self):
        """Create action button section."""
        button_frame = ttk.Frame(self.frame)
        button_frame.grid(row=4, column=0, pady=(15, 0))
        
        # Reset button
        ttk.Button(
            button_frame,
            text="🔄 RESET",
            command=self.reset_tab,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT, padx=5)
        
        # Help button
        ttk.Button(
            button_frame,
            text="❓ HELP",
            command=self.show_help_dialog,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT, padx=5)
    
    def _position_progress_frame(self):
        """Position progress frame for this layout."""
        if self.progress_frame:
            self.progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
    
    def _execute_scan(self):
        """Execute the sensitivity scan."""
        try:
            # Validate input
            pbip_file_path = self.clean_file_path(self.pbip_path.get())
            
            if not pbip_file_path:
                self.show_error("No File Selected", "Please select a PBIP file to scan")
                return
            
            self.validate_file_exists(pbip_file_path, "PBIP file")
            self.validate_pbip_file(pbip_file_path, "PBIP file")
            
            # Convert .pbip FILE to PBIP FOLDER path
            # The .pbip file is just metadata - the actual content is in folders
            # Modern PBIP format: .SemanticModel and .Report folders
            pbip_file = Path(pbip_file_path)
            pbip_base_name = pbip_file.stem
            pbip_parent = pbip_file.parent
            
            # Try to find the semantic model folder
            # Modern format: "Name.SemanticModel"
            semantic_model_folder = pbip_parent / f"{pbip_base_name}.SemanticModel"
            
            # Fallback: older format might just be "Name"
            if not semantic_model_folder.exists():
                semantic_model_folder = pbip_parent / pbip_base_name
            
            if not semantic_model_folder.exists():
                self.show_error(
                    "PBIP Folder Not Found", 
                    f"Could not find PBIP semantic model folder:\n"
                    f"Expected: {pbip_parent / (pbip_base_name + '.SemanticModel')}\n"
                    f"Or: {pbip_parent / pbip_base_name}\n\n"
                    f"The .pbip file must have a companion .SemanticModel folder."
                )
                return
            
            # Use the parent folder (which contains .pbip, .SemanticModel, .Report)
            # The PBIPReader expects the parent folder, not the .SemanticModel folder itself
            pbip_folder = pbip_parent
            
            # Clear previous results
            self._clear_results()
            
            # Log scan start
            self.log_message("=" * 60)
            self.log_message("🔍 Starting Sensitivity Scan")
            self.log_message(f"📁 File: {pbip_file.name}")
            self.log_message(f"📂 Folder: {pbip_folder.name}")
            self.log_message(f"⚙️  Mode: {self.scan_mode.get()}")
            self.log_message("=" * 60)
            
            # Run scan in background with FOLDER path
            self.run_in_background(
                target_func=lambda: self._perform_scan(str(pbip_folder)),
                success_callback=self._on_scan_success,
                error_callback=self._on_scan_error
            )
            
        except Exception as e:
            self.log_message(f"❌ Error: {e}")
            self.show_error("Scan Error", str(e))
    
    def _perform_scan(self, pbip_path: str) -> ScanResult:
        """
        Perform the actual scan (background thread).
        
        Args:
            pbip_path: Path to PBIP file
        
        Returns:
            ScanResult with findings
        """
        self.update_progress(0, "Initializing scanner...", show=True, persist=False)
        
        # Initialize scanner
        if self.scanner is None:
            self.scanner = TmdlScanner()
        
        self.update_progress(10, "Finding TMDL files...", show=True, persist=False)
        
        # Perform scan based on mode
        scan_mode = self.scan_mode.get()
        
        if scan_mode == "full":
            self.update_progress(25, "Scanning all TMDL files...", show=True, persist=False)
            result = self.scanner.scan_pbip(pbip_path)
        elif scan_mode == "tables":
            self.update_progress(25, "Scanning table TMDL files...", show=True, persist=False)
            result = self.scanner.scan_by_category(pbip_path, "tables")
        elif scan_mode == "roles":
            self.update_progress(25, "Scanning RLS role files...", show=True, persist=False)
            result = self.scanner.scan_by_category(pbip_path, "roles")
        elif scan_mode == "expressions":
            self.update_progress(25, "Scanning expression files...", show=True, persist=False)
            result = self.scanner.scan_by_category(pbip_path, "expressions")
        else:
            result = self.scanner.scan_pbip(pbip_path)
        
        self.update_progress(90, "Analyzing findings...", show=True, persist=False)
        
        self.update_progress(100, "Scan complete!", show=True, persist=False)
        
        return result
    
    def _on_scan_success(self, result: ScanResult):
        """
        Handle successful scan completion.
        
        Args:
            result: Scan result with findings
        """
        self.scan_result = result
        
        # Log summary
        self.log_message("=" * 60)
        self.log_message("✅ Scan Complete!")
        self.log_message(f"📊 Total Findings: {result.total_findings}")
        self.log_message(f"   🔴 HIGH Risk: {result.high_risk_count}")
        self.log_message(f"   🟡 MEDIUM Risk: {result.medium_risk_count}")
        self.log_message(f"   🟢 LOW Risk: {result.low_risk_count}")
        self.log_message(f"⏱️  Scan Duration: {result.scan_duration_seconds:.2f}s")
        self.log_message(f"📁 Files Scanned: {result.total_files_scanned}")
        self.log_message("=" * 60)
        
        # Update summary label
        summary_text = (
            f"Scan completed: {result.total_findings} findings "
            f"(🔴 {result.high_risk_count} HIGH, "
            f"🟡 {result.medium_risk_count} MEDIUM, "
            f"🟢 {result.low_risk_count} LOW)"
        )
        self.summary_label.config(text=summary_text)
        
        # Populate results tree
        self._populate_results_tree(result)
        
        # Enable export button
        self.export_button.config(state=tk.NORMAL)
        
        # Show summary dialog
        if result.has_high_risk:
            self.show_warning(
                "High Risk Findings Detected",
                f"Found {result.high_risk_count} HIGH RISK findings!\n\n"
                "Please review and address these before sharing your model."
            )
        elif result.total_findings == 0:
            self.show_info(
                "No Findings",
                "No sensitive data patterns detected in this model.\n\n"
                "Note: This tool uses pattern matching and may not catch everything."
            )
    
    def _on_scan_error(self, error: Exception):
        """
        Handle scan error.
        
        Args:
            error: Exception that occurred
        """
        self.log_message("=" * 60)
        self.log_message(f"❌ Scan Failed: {error}")
        self.log_message("=" * 60)
        self.show_error("Scan Failed", f"An error occurred during scanning:\n\n{error}")
    
    def _populate_results_tree(self, result: ScanResult):
        """
        Populate the results tree with findings.
        
        Args:
            result: Scan result to display
        """
        # Clear existing items
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        if not result.findings:
            return
        
        # Group findings by risk level
        high_findings = result.get_findings_by_risk(RiskLevel.HIGH)
        medium_findings = result.get_findings_by_risk(RiskLevel.MEDIUM)
        low_findings = result.get_findings_by_risk(RiskLevel.LOW)
        
        # Add HIGH risk findings
        if high_findings:
            high_parent = self.results_tree.insert(
                "",
                tk.END,
                text=f"🔴 HIGH RISK ({len(high_findings)})",
                values=("HIGH", "", "", "", ""),
                tags=("high_risk",)
            )
            for finding in high_findings:
                self._add_finding_to_tree(high_parent, finding)
        
        # Add MEDIUM risk findings
        if medium_findings:
            medium_parent = self.results_tree.insert(
                "",
                tk.END,
                text=f"🟡 MEDIUM RISK ({len(medium_findings)})",
                values=("MEDIUM", "", "", "", ""),
                tags=("medium_risk",)
            )
            for finding in medium_findings:
                self._add_finding_to_tree(medium_parent, finding)
        
        # Add LOW risk findings
        if low_findings:
            low_parent = self.results_tree.insert(
                "",
                tk.END,
                text=f"🟢 LOW RISK ({len(low_findings)})",
                values=("LOW", "", "", "", ""),
                tags=("low_risk",)
            )
            for finding in low_findings:
                self._add_finding_to_tree(low_parent, finding)
        
        # Configure tags for coloring
        self.results_tree.tag_configure("high_risk", foreground="#dc2626")
        self.results_tree.tag_configure("medium_risk", foreground="#d97706")
        self.results_tree.tag_configure("low_risk", foreground="#059669")
    
    def _add_finding_to_tree(self, parent, finding: Finding):
        """
        Add a single finding to the tree.
        
        Args:
            parent: Parent tree item
            finding: Finding to add
        """
        # Truncate matched text if too long
        matched_text = finding.pattern_match.matched_text
        if len(matched_text) > 50:
            matched_text = matched_text[:47] + "..."
        
        # Format confidence score
        confidence_pct = f"{finding.confidence_score * 100:.0f}%"
        
        # Add to tree
        self.results_tree.insert(
            parent,
            tk.END,
            text=f"{finding.pattern_match.pattern_name}: {matched_text}",
            values=(
                finding.risk_level.value,
                finding.file_type,
                finding.location_description,
                finding.pattern_match.pattern_name,
                confidence_pct
            ),
            tags=(f"{finding.risk_level.value.lower()}_risk",)
        )
    
    def _on_finding_selected(self, event):
        """Handle finding selection in tree."""
        selection = self.results_tree.selection()
        if not selection:
            return
        
        item_id = selection[0]
        item_values = self.results_tree.item(item_id)
        
        # Check if this is a parent node (risk level group)
        if not item_values['values'][1]:  # No file type means parent node
            return
        
        # Find the corresponding finding
        item_text = item_values['text']
        risk_level_str = item_values['values'][0]
        
        # Find matching finding
        finding = self._find_matching_finding(item_text, risk_level_str)
        
        if finding:
            self._display_finding_details(finding)
    
    def _find_matching_finding(self, item_text: str, risk_level_str: str) -> Optional[Finding]:
        """Find a finding matching the tree item."""
        if not self.scan_result:
            return None
        
        risk_level = RiskLevel[risk_level_str]
        findings = self.scan_result.get_findings_by_risk(risk_level)
        
        for finding in findings:
            finding_text = f"{finding.pattern_match.pattern_name}: {finding.pattern_match.matched_text}"
            if finding_text.startswith(item_text[:50]):  # Match truncated text
                return finding
        
        return None
    
    def _display_finding_details(self, finding: Finding):
        """
        Display details for a selected finding.
        
        Args:
            finding: Finding to display
        """
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)
        
        # Risk level header
        risk_tag = f"{finding.risk_level.value.lower()}_risk"
        self.details_text.insert(tk.END, f"Risk Level: {finding.risk_level.value}\n", risk_tag)
        self.details_text.insert(tk.END, f"Confidence: {finding.confidence_score * 100:.0f}%\n\n")
        
        # Pattern info
        self.details_text.insert(tk.END, "Pattern Detected:\n", "header")
        self.details_text.insert(tk.END, f"{finding.pattern_match.pattern_name}\n\n")
        
        # Location
        self.details_text.insert(tk.END, "Location:\n", "header")
        self.details_text.insert(tk.END, f"File: {finding.file_type}\n")
        self.details_text.insert(tk.END, f"{finding.location_description}\n\n")
        
        # Matched text
        self.details_text.insert(tk.END, "Matched Text:\n", "header")
        self.details_text.insert(tk.END, f"{finding.pattern_match.matched_text}\n\n")
        
        # Context
        if finding.pattern_match.context_before or finding.pattern_match.context_after:
            self.details_text.insert(tk.END, "Context:\n", "header")
            if finding.pattern_match.context_before:
                self.details_text.insert(tk.END, f"...{finding.pattern_match.context_before} ")
            self.details_text.insert(tk.END, f"[{finding.pattern_match.matched_text}]")
            if finding.pattern_match.context_after:
                self.details_text.insert(tk.END, f" {finding.pattern_match.context_after}...")
            self.details_text.insert(tk.END, "\n\n")
        
        # Description
        self.details_text.insert(tk.END, "Why This Matters:\n", "header")
        self.details_text.insert(tk.END, f"{finding.description}\n\n")
        
        # Recommendation
        self.details_text.insert(tk.END, "Recommended Action:\n", "header")
        self.details_text.insert(tk.END, f"{finding.recommendation}\n")
        
        self.details_text.config(state=tk.DISABLED)
    
    def _clear_results(self):
        """Clear all results and reset UI."""
        # Clear tree
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Clear details
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)
        self.details_text.config(state=tk.DISABLED)
        
        # Reset summary
        self.summary_label.config(text="Scanning...")
        
        # Disable export
        self.export_button.config(state=tk.DISABLED)
        
        # Clear scan result
        self.scan_result = None
    
    def _export_report(self):
        """Export scan results to a text file."""
        if not self.scan_result:
            self.show_error("No Results", "No scan results to export")
            return
        
        try:
            # Ask for save location
            default_name = f"sensitivity_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            file_path = filedialog.asksaveasfilename(
                title="Export Scan Report",
                defaultextension=".txt",
                initialfile=default_name,
                filetypes=[
                    ("Text Files", "*.txt"),
                    ("All Files", "*.*")
                ]
            )
            
            if not file_path:
                return  # User cancelled
            
            # Generate report
            report = self._generate_text_report()
            
            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(report)
            
            self.log_message(f"✅ Report exported: {Path(file_path).name}")
            self.show_info("Export Successful", f"Report saved to:\n{file_path}")
            
        except Exception as e:
            self.log_message(f"❌ Export failed: {e}")
            self.show_error("Export Failed", f"Could not export report:\n{e}")
    
    def _generate_text_report(self) -> str:
        """Generate a formatted text report of scan results."""
        if not self.scan_result:
            return ""
        
        lines = []
        lines.append("=" * 80)
        lines.append("SENSITIVITY SCAN REPORT")
        lines.append("=" * 80)
        lines.append("")
        lines.append(f"Scan Date: {self.scan_result.scan_time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Source: {self.scan_result.source_path}")
        lines.append(f"Duration: {self.scan_result.scan_duration_seconds:.2f} seconds")
        lines.append(f"Files Scanned: {self.scan_result.total_files_scanned}")
        lines.append("")
        lines.append("-" * 80)
        lines.append("SUMMARY")
        lines.append("-" * 80)
        lines.append(f"Total Findings: {self.scan_result.total_findings}")
        lines.append(f"  HIGH Risk:   {self.scan_result.high_risk_count}")
        lines.append(f"  MEDIUM Risk: {self.scan_result.medium_risk_count}")
        lines.append(f"  LOW Risk:    {self.scan_result.low_risk_count}")
        lines.append("")
        
        # High risk findings
        if self.scan_result.high_risk_count > 0:
            lines.append("=" * 80)
            lines.append("HIGH RISK FINDINGS")
            lines.append("=" * 80)
            lines.append("")
            for idx, finding in enumerate(self.scan_result.get_findings_by_risk(RiskLevel.HIGH), 1):
                lines.extend(self._format_finding_for_report(idx, finding))
        
        # Medium risk findings
        if self.scan_result.medium_risk_count > 0:
            lines.append("=" * 80)
            lines.append("MEDIUM RISK FINDINGS")
            lines.append("=" * 80)
            lines.append("")
            for idx, finding in enumerate(self.scan_result.get_findings_by_risk(RiskLevel.MEDIUM), 1):
                lines.extend(self._format_finding_for_report(idx, finding))
        
        # Low risk findings
        if self.scan_result.low_risk_count > 0:
            lines.append("=" * 80)
            lines.append("LOW RISK FINDINGS")
            lines.append("=" * 80)
            lines.append("")
            for idx, finding in enumerate(self.scan_result.get_findings_by_risk(RiskLevel.LOW), 1):
                lines.extend(self._format_finding_for_report(idx, finding))
        
        lines.append("=" * 80)
        lines.append("END OF REPORT")
        lines.append("=" * 80)
        lines.append("")
        lines.append("Generated by AE Multi-Tool - Sensitivity Scanner")
        lines.append("Built by Reid Havens of Analytic Endeavors")
        lines.append("https://www.analyticendeavors.com")
        
        return "\n".join(lines)
    
    def _format_finding_for_report(self, idx: int, finding: Finding) -> list:
        """Format a single finding for the text report."""
        lines = []
        lines.append(f"Finding #{idx}")
        lines.append("-" * 40)
        lines.append(f"Pattern: {finding.pattern_match.pattern_name}")
        lines.append(f"Risk Level: {finding.risk_level.value}")
        lines.append(f"Confidence: {finding.confidence_score * 100:.0f}%")
        lines.append(f"File Type: {finding.file_type}")
        lines.append(f"Location: {finding.location_description}")
        lines.append(f"Matched Text: {finding.pattern_match.matched_text}")
        
        if finding.pattern_match.context_before or finding.pattern_match.context_after:
            context = ""
            if finding.pattern_match.context_before:
                context += f"...{finding.pattern_match.context_before} "
            context += f"[{finding.pattern_match.matched_text}]"
            if finding.pattern_match.context_after:
                context += f" {finding.pattern_match.context_after}..."
            lines.append(f"Context: {context}")
        
        lines.append(f"Description: {finding.description}")
        lines.append(f"Recommendation: {finding.recommendation}")
        lines.append("")
        
        return lines
    
    def reset_tab(self) -> None:
        """Reset the tab to initial state."""
        # Clear inputs
        self.pbip_path.set("")
        self.scan_mode.set("full")
        
        # Clear results
        self._clear_results()
        self.summary_label.config(text="No scan performed yet")
        
        # Clear log
        self._clear_log(self.log_text)
        
        # Reset scanner
        self.scanner = None
        self.scan_result = None
        
        # Welcome message
        self.log_message("🔍 Welcome to Sensitivity Scanner!")
        self.log_message("📊 This tool scans Power BI semantic models for sensitive data")
        self.log_message("⚠️  Remember: This is STATIC ANALYSIS of model structure, not data values")
    
    def _open_pattern_manager(self):
        """Open the Pattern Manager window."""
        try:
            from tools.sensitivity_scanner.ui.pattern_manager import PatternManager
            
            def on_patterns_updated():
                """Callback when patterns are updated."""
                self.log_message("✅ Patterns updated! Please run a new scan to apply changes.")
                # Reinitialize scanner with new patterns
                self.scanner = None
            
            # Open pattern manager
            PatternManager(self.frame, self.scanner, on_patterns_updated)
            
        except Exception as e:
            self.log_message(f"❌ Failed to open Pattern Manager: {e}")
            messagebox.showerror("Error", f"Failed to open Pattern Manager:\n{e}")
    
    def show_help_dialog(self) -> None:
        """Show the help dialog with prominent disclaimer."""
        from tools.sensitivity_scanner.sensitivity_scanner_tool import SensitivityScannerTool
        tool = SensitivityScannerTool()
        help_content = tool.get_help_content()
        
        # Create help window
        help_window = tk.Toplevel(self.frame)
        help_window.title(help_content['title'])
        help_window.resizable(True, True)
        help_window.transient(self.frame.winfo_toplevel())
        help_window.grab_set()
        
        # Configure background
        help_window.configure(bg=AppConstants.COLORS['background'])
        
        # Center the window BEFORE showing (prevents flash)
        help_window.withdraw()  # Hide temporarily
        help_window.update_idletasks()
        root = self.frame.winfo_toplevel()
        x = root.winfo_rootx() + (root.winfo_width() - 550) // 2
        y = root.winfo_rooty() + (root.winfo_height() - 750) // 2
        help_window.geometry(f"550x750+{x}+{y}")
        help_window.deiconify()  # Show at correct position
        
        # Main container
        container = ttk.Frame(help_window, padding="20")
        container.pack(fill=tk.BOTH, expand=True)
        
        # Scrollable content area
        canvas_frame = ttk.Frame(container)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame, bg=AppConstants.COLORS['background'],
                          highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add help sections
        for section in help_content['sections']:
            # Special formatting for IMPORTANT DISCLAIMERS section
            if "IMPORTANT DISCLAIMERS" in section['title']:
                # Orange background box for disclaimer
                disclaimer_frame = tk.Frame(scrollable_frame, bg='#d97706', bd=2, relief=tk.SOLID)
                disclaimer_frame.pack(fill=tk.X, pady=(5, 20), padx=5)
                
                # Inner padding frame
                inner_frame = tk.Frame(disclaimer_frame, bg='#d97706')
                inner_frame.pack(fill=tk.X, padx=10, pady=10)
                
                # Title with warning icon
                title_label = tk.Label(inner_frame,
                                      text=section['title'],
                                      font=('Segoe UI', 11, 'bold'),
                                      fg='white',
                                      bg='#d97706',
                                      anchor=tk.W)
                title_label.pack(anchor=tk.W, pady=(0, 10))
                
                # Items in white text on orange background
                for item in section['items']:
                    item_label = tk.Label(inner_frame,
                                         text=item,
                                         font=('Segoe UI', 9),
                                         fg='white',
                                         bg='#d97706',
                                         anchor=tk.W,
                                         justify=tk.LEFT,
                                         wraplength=470)
                    item_label.pack(anchor=tk.W, pady=2)
            else:
                # Normal section formatting
                section_frame = ttk.Frame(scrollable_frame)
                section_frame.pack(fill=tk.X, pady=(15, 10))
                
                ttk.Label(section_frame,
                         text=section['title'],
                         font=('Segoe UI', 11, 'bold'),
                         foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
                
                ttk.Separator(section_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
                
                # Section items
                for item in section['items']:
                    ttk.Label(scrollable_frame,
                             text=f"• {item}",
                             font=('Segoe UI', 9),
                             foreground=AppConstants.COLORS['text_secondary'],
                             wraplength=470,
                             justify=tk.LEFT).pack(anchor=tk.W, padx=(10, 0), pady=2)
        
        # Close button at bottom
        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        close_button = ttk.Button(button_frame,
                                 text="✖ Close",
                                 command=help_window.destroy,
                                 style='Action.TButton')
        close_button.pack()
        
        # Bind escape key
        help_window.bind('<Escape>', lambda e: help_window.destroy())
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        help_window.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))
