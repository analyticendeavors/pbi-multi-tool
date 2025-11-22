"""
Pattern Manager UI - Manage sensitivity detection rules

Full CRUD functionality with Simple/Advanced modes for pattern management.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import re
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

from core.constants import AppConstants


class PatternManager:
    """Pattern Manager window for editing sensitivity patterns."""
    
    def __init__(self, parent, pattern_detector, on_patterns_updated_callback):
        """Initialize the Pattern Manager."""
        self.parent = parent
        self.pattern_detector = pattern_detector
        self.on_patterns_updated = on_patterns_updated_callback
        
        # Paths
        self.original_patterns_file = Path(__file__).parent.parent.parent.parent / "data" / "sensitivity_patterns.json"
        self.custom_patterns_file = Path(__file__).parent.parent.parent.parent / "data" / "sensitivity_patterns_custom.json"
        
        # Load current patterns
        self.patterns = self._load_current_patterns()
        self.selected_pattern = None
        
        # Create window
        self._create_window()
    
    def _load_current_patterns(self) -> Dict[str, Any]:
        """Load current patterns."""
        try:
            if self.custom_patterns_file.exists():
                with open(self.custom_patterns_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            with open(self.original_patterns_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load patterns: {e}")
            return {"patterns": {"high_risk": [], "medium_risk": [], "low_risk": []}}
    
    def _create_window(self):
        """Create the pattern manager window."""
        self.window = tk.Toplevel(self.parent)
        self.window.title("⚙️ Rule Manager - Sensitivity Scanner")
        self.window.transient(self.parent.winfo_toplevel())
        self.window.grab_set()
        
        self.window.configure(bg=AppConstants.COLORS['background'])
        
        # Center window BEFORE showing (prevents flash)
        self.window.withdraw()  # Hide temporarily
        self.window.update_idletasks()
        root = self.parent.winfo_toplevel()
        x = root.winfo_rootx() + (root.winfo_width() - 1035) // 2
        y = root.winfo_rooty() + (root.winfo_height() - 805) // 2
        self.window.geometry(f"1035x805+{x}+{y}")
        self.window.deiconify()  # Show at correct position
        
        # Main container
        container = ttk.Frame(self.window, padding="20")
        container.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(container, text="⚙️ Rule Manager",
                 font=('Segoe UI', 16, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
        
        status_text = "Using custom patterns" if self.custom_patterns_file.exists() else "Using default patterns"
        ttk.Label(container, text=f"📊 {status_text}",
                 font=('Segoe UI', 9, 'italic')).pack(anchor=tk.W, pady=(3, 20))
        
        # Content (split view) - Fill remaining space
        content = ttk.Frame(container)
        content.pack(fill=tk.BOTH, expand=True)
        content.columnconfigure(0, weight=0, minsize=425)  # Fixed width for pattern list
        content.columnconfigure(1, weight=1)  # Editor can expand
        content.rowconfigure(0, weight=1)  # Make row expandable
        
        # Left: Pattern list
        self._create_pattern_list(content)
        
        # Right: Editor
        self._create_editor(content)
        
        # Bottom buttons
        self._create_buttons(container)
    
    def _create_pattern_list(self, parent):
        """Create pattern list view."""
        list_frame = ttk.LabelFrame(parent, text="📋 Current Patterns", padding="10")
        list_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=(0, 10))
        list_frame.rowconfigure(0, weight=1)  # Make tree expandable
        
        # Tree
        tree_container = ttk.Frame(list_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(tree_container, columns=('id', 'name', 'risk'),
                                show='headings', yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.tree.yview)
        
        self.tree.heading('id', text='Pattern ID')
        self.tree.heading('name', text='Pattern Name')
        self.tree.heading('risk', text='Risk')
        
        self.tree.column('id', width=150)
        self.tree.column('name', width=200)
        self.tree.column('risk', width=80)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        
        self._populate_tree()
        
        # Count
        count = sum(len(self.patterns['patterns'].get(r, [])) for r in ['high_risk', 'medium_risk', 'low_risk'])
        ttk.Label(list_frame, text=f"Total: {count} patterns").pack(pady=(10, 0))
    
    def _create_editor(self, parent):
        """Create editor form."""
        editor_frame = ttk.LabelFrame(parent, text="✏️ Pattern Editor", padding="10")
        editor_frame.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.E, tk.W))
        editor_frame.rowconfigure(0, weight=1)  # Make canvas expandable
        
        # Scrollable form
        canvas = tk.Canvas(editor_frame, bg=AppConstants.COLORS['background'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(editor_frame, orient=tk.VERTICAL, command=canvas.yview)
        form = ttk.Frame(canvas)
        
        form.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=form, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Enable mouse wheel scrolling ONLY when mouse is over canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_mousewheel(event):
            canvas.bind("<MouseWheel>", _on_mousewheel)
        
        def _unbind_mousewheel(event):
            canvas.unbind("<MouseWheel>")
        
        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)
        
        # Form fields
        self.form_vars = {}
        row = 0
        
        # Pattern Mode Toggle
        mode_frame = ttk.Frame(form)
        mode_frame.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        ttk.Label(mode_frame, text="Pattern Mode:", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        self.form_vars['mode'] = tk.StringVar(value='simple')
        
        # Create radio buttons without background (like scan mode)
        simple_rb = tk.Radiobutton(mode_frame, text="🎯 Simple", variable=self.form_vars['mode'], value='simple',
                                  command=self._on_mode_change, relief=tk.FLAT, borderwidth=0,
                                  bg=AppConstants.COLORS['background'], fg=AppConstants.COLORS['text_primary'],
                                  selectcolor=AppConstants.COLORS['background'], activebackground=AppConstants.COLORS['background'],
                                  font=('Segoe UI', 9), cursor='hand2')
        simple_rb.pack(side=tk.LEFT, padx=(0, 15))
        
        advanced_rb = tk.Radiobutton(mode_frame, text="⚙️ Advanced", variable=self.form_vars['mode'], value='advanced',
                                    command=self._on_mode_change, relief=tk.FLAT, borderwidth=0,
                                    bg=AppConstants.COLORS['background'], fg=AppConstants.COLORS['text_primary'],
                                    selectcolor=AppConstants.COLORS['background'], activebackground=AppConstants.COLORS['background'],
                                    font=('Segoe UI', 9), cursor='hand2')
        advanced_rb.pack(side=tk.LEFT)
        row += 1
        
        # Simple Mode Section
        self.simple_frame = ttk.Frame(form)
        self.simple_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        self._create_simple_mode_ui(self.simple_frame)
        row += 1
        
        # Advanced Mode Section
        self.advanced_frame = ttk.Frame(form)
        self.advanced_frame.grid(row=row, column=0, sticky=(tk.W, tk.E))
        self._create_advanced_mode_ui(self.advanced_frame)
        self.advanced_frame.grid_remove()  # Hide by default
        row += 1
        
        # Pattern ID (always visible)
        ttk.Label(form, text="Pattern ID:*", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['id'] = tk.StringVar()
        ttk.Entry(form, textvariable=self.form_vars['id']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        ttk.Label(form, text="(Unique identifier: lowercase, underscores, no spaces. E.g., 'ssn_us')",
                 font=('Segoe UI', 8, 'italic'), foreground='#666').grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
        
        # Pattern Name (always visible)
        ttk.Label(form, text="Pattern Name:*", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['name'] = tk.StringVar()
        ttk.Entry(form, textvariable=self.form_vars['name']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        ttk.Label(form, text="(Human-readable display name. E.g., 'US Social Security Number')",
                 font=('Segoe UI', 8, 'italic'), foreground='#666').grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
        
        # Risk Level
        ttk.Label(form, text="Risk Level:*", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['risk_level'] = tk.StringVar(value='medium_risk')
        risk_frame = ttk.Frame(form)
        risk_frame.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        
        # Create radio buttons without background (like scan mode)
        high_rb = tk.Radiobutton(risk_frame, text="🔴 High", variable=self.form_vars['risk_level'], value='high_risk',
                                relief=tk.FLAT, borderwidth=0,
                                bg=AppConstants.COLORS['background'], fg=AppConstants.COLORS['text_primary'],
                                selectcolor=AppConstants.COLORS['background'], activebackground=AppConstants.COLORS['background'],
                                font=('Segoe UI', 9), cursor='hand2')
        high_rb.pack(side=tk.LEFT, padx=(0, 15))
        
        medium_rb = tk.Radiobutton(risk_frame, text="🟡 Medium", variable=self.form_vars['risk_level'], value='medium_risk',
                                  relief=tk.FLAT, borderwidth=0,
                                  bg=AppConstants.COLORS['background'], fg=AppConstants.COLORS['text_primary'],
                                  selectcolor=AppConstants.COLORS['background'], activebackground=AppConstants.COLORS['background'],
                                  font=('Segoe UI', 9), cursor='hand2')
        medium_rb.pack(side=tk.LEFT, padx=(0, 15))
        
        low_rb = tk.Radiobutton(risk_frame, text="🟢 Low", variable=self.form_vars['risk_level'], value='low_risk',
                               relief=tk.FLAT, borderwidth=0,
                               bg=AppConstants.COLORS['background'], fg=AppConstants.COLORS['text_primary'],
                               selectcolor=AppConstants.COLORS['background'], activebackground=AppConstants.COLORS['background'],
                               font=('Segoe UI', 9), cursor='hand2')
        low_rb.pack(side=tk.LEFT)
        row += 1
        
        # Description
        ttk.Label(form, text="Description:*", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['description'] = tk.Text(form, height=3, width=50, wrap=tk.WORD)
        self.form_vars['description'].grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        ttk.Label(form, text="(What this pattern detects and why it matters)",
                 font=('Segoe UI', 8, 'italic'), foreground='#666').grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
        
        # Recommended Action
        ttk.Label(form, text="Recommended Action:", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['recommendation'] = tk.Text(form, height=3, width=50, wrap=tk.WORD)
        self.form_vars['recommendation'].grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        ttk.Label(form, text="(What action should users take when this is found? E.g., 'Implement RLS', 'Use parameters')",
                 font=('Segoe UI', 8, 'italic'), foreground='#666').grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
        
        # Examples
        ttk.Label(form, text="Examples (one per line):", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['examples'] = tk.Text(form, height=3, width=50, wrap=tk.WORD)
        self.form_vars['examples'].grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        row += 1
        
        # Form buttons (no background style)
        btn_frame = ttk.Frame(form)
        btn_frame.grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        
        # Create buttons with no background
        add_btn = tk.Button(btn_frame, text="➕ Add New", command=self._add_pattern,
                           relief=tk.FLAT, borderwidth=0, padx=10, pady=5,
                           bg=AppConstants.COLORS['background'],
                           fg=AppConstants.COLORS['text_primary'],
                           cursor='hand2',
                           font=('Segoe UI', 9))
        add_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        save_btn = tk.Button(btn_frame, text="💾 Save", command=self._save_pattern,
                            relief=tk.FLAT, borderwidth=0, padx=10, pady=5,
                            bg=AppConstants.COLORS['background'],
                            fg=AppConstants.COLORS['text_primary'],
                            cursor='hand2',
                            font=('Segoe UI', 9))
        save_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        delete_btn = tk.Button(btn_frame, text="🗑️ Delete", command=self._delete_pattern,
                              relief=tk.FLAT, borderwidth=0, padx=10, pady=5,
                              bg=AppConstants.COLORS['background'],
                              fg=AppConstants.COLORS['text_primary'],
                              cursor='hand2',
                              font=('Segoe UI', 9))
        delete_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        clear_btn = tk.Button(btn_frame, text="🔄 Clear", command=self._clear_form,
                             relief=tk.FLAT, borderwidth=0, padx=10, pady=5,
                             bg=AppConstants.COLORS['background'],
                             fg=AppConstants.COLORS['text_primary'],
                             cursor='hand2',
                             font=('Segoe UI', 9))
        clear_btn.pack(side=tk.LEFT)
    
    def _create_simple_mode_ui(self, parent):
        """Create simple mode UI."""
        # Configure parent column to have fixed width
        parent.columnconfigure(0, weight=1, minsize=450)
        
        row = 0
        
        # Pattern Type
        ttk.Label(parent, text="Pattern Type:*", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['pattern_type'] = tk.StringVar(value='Contains Text')
        pattern_types = [
            'Contains Text',
            'Starts With',
            'Ends With',
            'Email Address',
            'Phone Number (US)',
            'Credit Card',
            'URL/Link',
            'IP Address',
            'Date (MM/DD/YYYY)',
            'Date (DD/MM/YYYY)',
            'Date (YYYY-MM-DD)',
            'Custom Date Pattern',
        ]
        type_combo = ttk.Combobox(parent, textvariable=self.form_vars['pattern_type'], 
                                 values=pattern_types, state='readonly', width=30)
        type_combo.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        type_combo.bind('<<ComboboxSelected>>', lambda e: self._update_simple_pattern())
        row += 1
        
        # Search Text
        ttk.Label(parent, text="Search Text:", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['search_text'] = tk.StringVar()
        self.form_vars['search_text'].trace_add('write', lambda *args: self._update_simple_pattern())
        self.search_text_entry = ttk.Entry(parent, textvariable=self.form_vars['search_text'])
        self.search_text_entry.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        self.search_text_helper = ttk.Label(parent, text="(Templates with '*' don't require search text - will be auto-generated)",
                 font=('Segoe UI', 8, 'italic'), foreground='#666')
        self.search_text_helper.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
        
        # Options
        ttk.Label(parent, text="Options:", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        options_frame = ttk.Frame(parent)
        options_frame.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        
        self.form_vars['case_sensitive'] = tk.BooleanVar(value=False)
        self.form_vars['whole_word'] = tk.BooleanVar(value=False)
        
        case_cb = ttk.Checkbutton(options_frame, text="Case sensitive", variable=self.form_vars['case_sensitive'],
                                 command=self._update_simple_pattern)
        case_cb.pack(side=tk.LEFT, padx=(0, 15))
        
        word_cb = ttk.Checkbutton(options_frame, text="Match whole word only", variable=self.form_vars['whole_word'],
                                 command=self._update_simple_pattern)
        word_cb.pack(side=tk.LEFT)
        row += 1
        
        # Generated Pattern Preview
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(5, 10))
        row += 1
        ttk.Label(parent, text="Generated Pattern:", font=('Segoe UI', 9, 'italic')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.preview_label = ttk.Label(parent, text="", font=('Consolas', 9),
                                      foreground='#0066cc', wraplength=450)
        self.preview_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
        
        # Pattern Tester
        ttk.Label(parent, text="Test Your Pattern:", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['test_input'] = tk.StringVar()
        self.form_vars['test_input'].trace_add('write', lambda *args: self._test_pattern())
        ttk.Entry(parent, textvariable=self.form_vars['test_input']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        
        self.test_result_label = ttk.Label(parent, text="", font=('Segoe UI', 9), wraplength=450)
        self.test_result_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
    
    def _create_advanced_mode_ui(self, parent):
        """Create advanced mode UI."""
        # Configure parent column to have fixed width (same as simple mode)
        parent.columnconfigure(0, weight=1, minsize=450)
        
        row = 0
        
        # Regex Pattern
        ttk.Label(parent, text="Regex Pattern:*", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['pattern'] = tk.StringVar()
        self.form_vars['pattern'].trace_add('write', lambda *args: self._test_pattern())
        ttk.Entry(parent, textvariable=self.form_vars['pattern']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        row += 1
        
        # Pattern Tester (Advanced)
        ttk.Label(parent, text="Test Your Pattern:", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        self.form_vars['test_input_advanced'] = tk.StringVar()
        self.form_vars['test_input_advanced'].trace_add('write', lambda *args: self._test_pattern())
        ttk.Entry(parent, textvariable=self.form_vars['test_input_advanced']).grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1
        
        self.test_result_label_advanced = ttk.Label(parent, text="", font=('Segoe UI', 9), wraplength=450)
        self.test_result_label_advanced.grid(row=row, column=0, sticky=tk.W, pady=(0, 15))
        row += 1
    
    def _on_mode_change(self):
        """Handle mode change between simple and advanced."""
        mode = self.form_vars['mode'].get()
        
        if mode == 'simple':
            self.simple_frame.grid()
            self.advanced_frame.grid_remove()
            self._update_simple_pattern()
        else:
            self.simple_frame.grid_remove()
            self.advanced_frame.grid()
    
    def _convert_date_format_to_regex(self, format_str: str) -> str:
        """
        Convert user-friendly date format to regex.
        
        Supported tokens:
        - dd: Day (01-31)
        - mm: Month (01-12)
        - yyyy: Year (4 digits)
        - yy: Year (2 digits)
        
        Examples:
        'dd/mm/yyyy' -> r'\b(?:0?[1-9]|[12][0-9]|3[01])/(?:0?[1-9]|1[0-2])/\d{4}\b'
        'mm-dd-yyyy' -> r'\b(?:0?[1-9]|1[0-2])-(?:0?[1-9]|[12][0-9]|3[01])-\d{4}\b'
        """
        # Define regex components
        replacements = {
            'yyyy': r'(?:19|20)\d{2}',  # 1900-2099
            'yy': r'\d{2}',              # Any 2 digits
            'mm': r'(?:0?[1-9]|1[0-2])', # 01-12 or 1-12
            'dd': r'(?:0?[1-9]|[12][0-9]|3[01])'  # 01-31 or 1-31
        }
        
        # Escape the format string to handle special regex chars in separators
        regex = format_str.lower()
        
        # Replace date tokens with regex patterns (order matters - yyyy before yy)
        for token, pattern in sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True):
            if token in regex:
                # Temporarily replace with a placeholder to avoid double-replacement
                placeholder = f"___{token.upper()}___"
                regex = regex.replace(token, placeholder)
        
        # Escape any special regex characters in separators
        regex = re.escape(regex)
        
        # Replace placeholders with actual regex patterns
        for token, pattern in replacements.items():
            placeholder = f"___{token.upper()}___"
            regex = regex.replace(re.escape(placeholder), pattern)
        
        # Add word boundaries
        regex = r'\b' + regex + r'\b'
        
        return regex
    
    def _generate_regex(self, pattern_type, search_text, case_sensitive, whole_word):
        """Generate regex from simple mode inputs."""
        # Pre-built templates
        templates = {
            'Email Address': r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b',
            'Phone Number (US)': r'\b(?:\+?1[-.]?)?\(?([0-9]{3})\)?[-.]?([0-9]{3})[-.]?([0-9]{4})\b',
            'Credit Card': r'\b(?:\d{4}[-\s]?){3}\d{4}\b',
            'URL/Link': r'https?://(?:www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b(?:[-a-zA-Z0-9()@:%_\+.~#?&/=]*)',
            'IP Address': r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b',
            'Date (MM/DD/YYYY)': r'\b(?:0?[1-9]|1[0-2])/(?:0?[1-9]|[12][0-9]|3[01])/(?:19|20)?\d{2}\b',
            'Date (DD/MM/YYYY)': r'\b(?:0?[1-9]|[12][0-9]|3[01])/(?:0?[1-9]|1[0-2])/(?:19|20)?\d{2}\b',
            'Date (YYYY-MM-DD)': r'\b(?:19|20)\d{2}-(?:0?[1-9]|1[0-2])-(?:0?[1-9]|[12][0-9]|3[01])\b',
        }
        
        if pattern_type in templates:
            return templates[pattern_type]
        
        if not search_text:
            return ""
        
        # Handle Custom Date Pattern
        if pattern_type == 'Custom Date Pattern':
            return self._convert_date_format_to_regex(search_text)
        
        # Escape special regex characters
        escaped = re.escape(search_text)
        
        # Build pattern based on type
        if pattern_type == 'Contains Text':
            pattern = escaped
        elif pattern_type == 'Starts With':
            pattern = f'^{escaped}'
        elif pattern_type == 'Ends With':
            pattern = f'{escaped}$'
        else:
            pattern = escaped
        
        # Add word boundaries if needed
        if whole_word and pattern_type == 'Contains Text':
            pattern = f'\\b{pattern}\\b'
        
        # Add case insensitivity flag if needed
        if not case_sensitive:
            pattern = f'(?i){pattern}'
        
        return pattern
    
    def _update_simple_pattern(self):
        """Update the pattern preview from simple mode inputs."""
        if self.form_vars['mode'].get() != 'simple':
            return
        
        pattern_type = self.form_vars['pattern_type'].get()
        search_text = self.form_vars['search_text'].get()
        case_sensitive = self.form_vars['case_sensitive'].get()
        whole_word = self.form_vars['whole_word'].get()
        
        # Check if this is a template pattern (doesn't need search text)
        template_patterns = ['Email Address', 'Phone Number (US)', 'Credit Card', 'URL/Link', 'IP Address', 
                            'Date (MM/DD/YYYY)', 'Date (DD/MM/YYYY)', 'Date (YYYY-MM-DD)']
        is_template = pattern_type in template_patterns
        
        # Special case: Custom Date Pattern needs search text
        needs_search_text = not is_template or pattern_type == 'Custom Date Pattern'
        
        # Enable/disable search text entry based on pattern type
        if is_template:
            self.search_text_entry.config(state='disabled')
            # Clear the search text when switching to a template
            self.form_vars['search_text'].set('')
            # Make it visually grey when disabled
            style = ttk.Style()
            style.map('Disabled.TEntry',
                     fieldbackground=[('disabled', '#e0e0e0')],
                     foreground=[('disabled', '#666666')])
            self.search_text_entry.config(style='Disabled.TEntry')
            self.search_text_helper.config(text="(✓ This template auto-generates the pattern - no search text needed)", foreground='#059669')
        elif pattern_type == 'Custom Date Pattern':
            self.search_text_entry.config(state='normal', style='TEntry')
            self.search_text_helper.config(text="(Enter date format: Use 'dd' for day, 'mm' for month, 'yyyy' for year. E.g., 'dd-mm-yyyy')", foreground='#666')
        else:
            self.search_text_entry.config(state='normal', style='TEntry')
            self.search_text_helper.config(text="(Enter the text pattern you want to match)", foreground='#666')
        
        regex = self._generate_regex(pattern_type, search_text, case_sensitive, whole_word)
        
        # Update the advanced pattern field
        self.form_vars['pattern'].set(regex)
        
        # Update preview
        if regex:
            self.preview_label.config(text=regex)
        else:
            if is_template:
                self.preview_label.config(text="Select a pattern type to see generated regex...")
            else:
                self.preview_label.config(text="Enter search text to see pattern...")
        
        # Test pattern
        self._test_pattern()
    
    def _test_pattern(self):
        """Test the current pattern against test input."""
        mode = self.form_vars['mode'].get()
        pattern_str = self.form_vars['pattern'].get()
        
        if mode == 'simple':
            test_input = self.form_vars['test_input'].get()
            result_label = self.test_result_label
        else:
            test_input = self.form_vars['test_input_advanced'].get()
            result_label = self.test_result_label_advanced
        
        if not pattern_str or not test_input:
            result_label.config(text="", foreground='#666666')
            return
        
        try:
            compiled = re.compile(pattern_str)
            matches = compiled.findall(test_input)
            
            if matches:
                match_str = '", "'.join(str(m) for m in matches[:3])
                if len(matches) > 3:
                    match_str += f", ... ({len(matches)} total)"
                result_label.config(text=f'✅ Match found: "{match_str}"', foreground='#16a34a')
            else:
                result_label.config(text='❌ No match', foreground='#dc2626')
        except re.error as e:
            result_label.config(text=f'⚠️ Invalid regex: {str(e)[:50]}', foreground='#d97706')
    
    def _create_buttons(self, parent):
        """Create bottom buttons."""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        
        left = ttk.Frame(btn_frame)
        left.pack(side=tk.LEFT)
        
        # Reset and Export with same style as right side
        ttk.Button(left, text="🔄 Reset to Defaults", command=self._reset).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(left, text="📤 Export", command=self._export).pack(side=tk.LEFT)
        
        right = ttk.Frame(btn_frame)
        right.pack(side=tk.RIGHT)
        
        # Save & Close button matching SCAN button style
        save_close_btn = tk.Button(right, text="💾 Save & Close", command=self._save_close,
                                  relief=tk.FLAT, borderwidth=0,
                                  bg=AppConstants.COLORS['primary'], fg='white',
                                  activebackground='#0a7ea3', activeforeground='white',
                                  padx=20, pady=8, cursor='hand2',
                                  font=('Segoe UI', 10, 'bold'))
        save_close_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(right, text="✖ Close", command=self.window.destroy).pack(side=tk.LEFT)
    
    def _populate_tree(self):
        """Populate tree with patterns."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for pattern in self.patterns['patterns'].get('high_risk', []):
            self.tree.insert('', tk.END, values=(pattern['id'], pattern['name'], '🔴 HIGH'))
        for pattern in self.patterns['patterns'].get('medium_risk', []):
            self.tree.insert('', tk.END, values=(pattern['id'], pattern['name'], '🟡 MEDIUM'))
        for pattern in self.patterns['patterns'].get('low_risk', []):
            self.tree.insert('', tk.END, values=(pattern['id'], pattern['name'], '🟢 LOW'))
    
    def _on_select(self, event):
        """Handle pattern selection."""
        selection = self.tree.selection()
        if not selection:
            return
        
        values = self.tree.item(selection[0], 'values')
        pattern_id = values[0]
        
        # Find pattern
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            for pattern in self.patterns['patterns'].get(risk_level, []):
                if pattern['id'] == pattern_id:
                    self._load_to_form(pattern, risk_level)
                    return
    
    def _load_to_form(self, pattern, risk_level):
        """Load pattern into form."""
        self.selected_pattern = pattern
        self.form_vars['id'].set(pattern.get('id', ''))
        self.form_vars['name'].set(pattern.get('name', ''))
        self.form_vars['pattern'].set(pattern.get('pattern', ''))
        self.form_vars['risk_level'].set(risk_level)
        
        desc = self.form_vars['description']
        desc.delete('1.0', tk.END)
        desc.insert('1.0', pattern.get('description', ''))
        
        # Load recommendation if it exists
        rec = self.form_vars['recommendation']
        rec.delete('1.0', tk.END)
        rec.insert('1.0', pattern.get('recommendation', ''))
        
        ex = self.form_vars['examples']
        ex.delete('1.0', tk.END)
        ex.insert('1.0', '\n'.join(pattern.get('examples', [])))
        
        # Switch to advanced mode when loading existing pattern
        self.form_vars['mode'].set('advanced')
        self._on_mode_change()
    
    def _clear_form(self):
        """Clear form."""
        self.selected_pattern = None
        self.form_vars['id'].set('')
        self.form_vars['name'].set('')
        self.form_vars['pattern'].set('')
        self.form_vars['risk_level'].set('medium_risk')
        self.form_vars['description'].delete('1.0', tk.END)
        self.form_vars['recommendation'].delete('1.0', tk.END)
        self.form_vars['examples'].delete('1.0', tk.END)
        
        # Clear simple mode fields
        if 'search_text' in self.form_vars:
            self.form_vars['search_text'].set('')
        if 'test_input' in self.form_vars:
            self.form_vars['test_input'].set('')
        if 'test_input_advanced' in self.form_vars:
            self.form_vars['test_input_advanced'].set('')
        
        # Reset to simple mode
        self.form_vars['mode'].set('simple')
        self._on_mode_change()
    
    def _validate(self):
        """Validate form."""
        if not self.form_vars['id'].get():
            messagebox.showerror("Error", "Pattern ID is required")
            return False
        if not self.form_vars['name'].get():
            messagebox.showerror("Error", "Pattern Name is required")
            return False
        if not self.form_vars['pattern'].get():
            messagebox.showerror("Error", "Regex Pattern is required")
            return False
        if not self.form_vars['description'].get('1.0', tk.END).strip():
            messagebox.showerror("Error", "Description is required")
            return False
        
        # Validate regex
        try:
            re.compile(self.form_vars['pattern'].get())
        except re.error as e:
            messagebox.showerror("Invalid Regex", f"Invalid regular expression:\n{e}")
            return False
        
        return True
    
    def _form_to_pattern(self):
        """Convert form to pattern dict."""
        examples = [l.strip() for l in self.form_vars['examples'].get('1.0', tk.END).split('\n') if l.strip()]
        recommendation = self.form_vars['recommendation'].get('1.0', tk.END).strip()
        
        pattern = {
            'id': self.form_vars['id'].get(),
            'name': self.form_vars['name'].get(),
            'pattern': self.form_vars['pattern'].get(),
            'description': self.form_vars['description'].get('1.0', tk.END).strip(),
            'categories': [],
            'examples': examples
        }
        
        # Add recommendation if provided
        if recommendation:
            pattern['recommendation'] = recommendation
        
        return pattern
    
    def _add_pattern(self):
        """Add new pattern."""
        if not self._validate():
            return
        
        pattern_id = self.form_vars['id'].get()
        
        # Check duplicate
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            for p in self.patterns['patterns'].get(risk_level, []):
                if p['id'] == pattern_id:
                    messagebox.showerror("Error", f"Pattern ID '{pattern_id}' already exists")
                    return
        
        pattern = self._form_to_pattern()
        risk_level = self.form_vars['risk_level'].get()
        self.patterns['patterns'][risk_level].append(pattern)
        
        self._save_patterns()
        self._populate_tree()
        self._clear_form()
        messagebox.showinfo("Success", "Pattern added!")
    
    def _save_pattern(self):
        """Save changes to existing pattern."""
        if not self.selected_pattern:
            messagebox.showwarning("No Selection", "Please select a pattern to edit")
            return
        
        if not self._validate():
            return
        
        updated = self._form_to_pattern()
        
        # Find and update
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            patterns = self.patterns['patterns'][risk_level]
            for i, p in enumerate(patterns):
                if p['id'] == self.selected_pattern['id']:
                    new_risk = self.form_vars['risk_level'].get()
                    if risk_level != new_risk:
                        patterns.pop(i)
                        self.patterns['patterns'][new_risk].append(updated)
                    else:
                        patterns[i] = updated
                    break
        
        self._save_patterns()
        self._populate_tree()
        self._clear_form()
        messagebox.showinfo("Success", "Pattern updated!")
    
    def _delete_pattern(self):
        """Delete selected pattern."""
        if not self.selected_pattern:
            messagebox.showwarning("No Selection", "Please select a pattern to delete")
            return
        
        if not messagebox.askyesno("Confirm", f"Delete pattern '{self.selected_pattern['name']}'?"):
            return
        
        for risk_level in ['high_risk', 'medium_risk', 'low_risk']:
            patterns = self.patterns['patterns'][risk_level]
            for i, p in enumerate(patterns):
                if p['id'] == self.selected_pattern['id']:
                    patterns.pop(i)
                    break
        
        self._save_patterns()
        self._populate_tree()
        self._clear_form()
        messagebox.showinfo("Success", "Pattern deleted!")
    
    def _save_patterns(self):
        """Save patterns to custom file."""
        try:
            with open(self.custom_patterns_file, 'w', encoding='utf-8') as f:
                json.dump(self.patterns, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {e}")
    
    def _reset(self):
        """Reset to defaults."""
        if not messagebox.askyesno("Reset", "Reset all patterns to defaults?\n\nThis will delete all custom patterns."):
            return
        
        if self.custom_patterns_file.exists():
            self.custom_patterns_file.unlink()
        
        self.patterns = self._load_current_patterns()
        self._populate_tree()
        self._clear_form()
        messagebox.showinfo("Success", "Reset to defaults!")
    
    def _export(self):
        """Export patterns."""
        path = filedialog.asksaveasfilename(
            title="Export Patterns",
            defaultextension=".json",
            initialfile=f"patterns_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            filetypes=[("JSON Files", "*.json")]
        )
        
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(self.patterns, f, indent=2)
                messagebox.showinfo("Success", f"Exported to:\n{path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {e}")
    
    def _save_close(self):
        """Save and close."""
        self._save_patterns()
        if self.on_patterns_updated:
            self.on_patterns_updated()
        self.window.destroy()
        messagebox.showinfo("Success", "Patterns saved! Re-scan to apply changes.")
