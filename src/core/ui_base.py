"""
UI Base Classes - Common UI patterns and base components
Built by Reid Havens of Analytic Endeavors

Provides reusable UI components and patterns for tool tabs.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
from pathlib import Path
from typing import Optional, Dict, Any, Callable, List
from abc import ABC, abstractmethod

from core.constants import AppConstants


class BaseToolTab(ABC):
    """
    Base class for tool UI tabs.
    Provides common UI patterns and functionality.
    """
    
    def __init__(self, parent, main_app, tool_id: str, tool_name: str):
        self.parent = parent
        self.main_app = main_app
        self.tool_id = tool_id
        self.tool_name = tool_name
        
        # Create main frame for this tab
        self.frame = ttk.Frame(parent, padding="20")
        
        # Common UI state
        self.is_busy = False
        self.progress_bar = None
        self.progress_label = None  # New: Progress status label
        self.progress_frame = None  # New: Progress container
        self.progress_persist = False  # New: Whether to keep progress visible
        self.log_text = None
        
        # Setup styling
        self._setup_common_styling()
    
    def get_frame(self) -> ttk.Frame:
        """Return the main frame for this tab"""
        return self.frame
    
    @abstractmethod
    def setup_ui(self) -> None:
        """Setup the UI for this tab - must be implemented by subclasses"""
        pass
    
    @abstractmethod
    def reset_tab(self) -> None:
        """Reset the tab to initial state - must be implemented by subclasses"""
        pass
    
    @abstractmethod
    def show_help_dialog(self) -> None:
        """Show help dialog for this tab - must be implemented by subclasses"""
        pass
    
    def _setup_common_styling(self):
        """Setup common professional styling"""
        style = ttk.Style()
        style.theme_use('clam')
        colors = AppConstants.COLORS
        
        # Common styles
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
            },
            'Brand.TButton': {
                'background': colors['accent'], 
                'foreground': colors['surface'], 
                'font': ('Segoe UI', 10, 'bold'), 
                'padding': (15, 8)
            },
            'Info.TButton': {
                'background': colors['info'], 
                'foreground': colors['surface'], 
                'font': ('Segoe UI', 9), 
                'padding': (12, 6)
            },
            'TProgressbar': {
                'background': colors['accent'], 
                'troughcolor': colors['border']
            },
            'TEntry': {
                'fieldbackground': colors['surface']
            },
            'TFrame': {
                'background': colors['background']
            },
            'TLabel': {
                'background': colors['background']
            }
        }
        
        for style_name, config in styles.items():
            style.configure(style_name, **config)
    
    def create_file_input_section(self, parent: ttk.Widget, title: str, 
                                 file_types: List[tuple], guide_text: List[str] = None) -> Dict[str, Any]:
        """
        Create a standardized file input section.
        
        Returns:
            Dict with 'frame', 'path_var', 'entry', 'browse_button'
        """
        section_frame = ttk.LabelFrame(parent, text=title, 
                                     style='Section.TLabelframe', padding="20")
        
        content_frame = ttk.Frame(section_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        
        # LEFT: Guide (if provided)
        if guide_text:
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
        
        # RIGHT: File input
        input_frame = ttk.Frame(content_frame)
        input_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        input_frame.columnconfigure(1, weight=1)
        
        # File path variable
        path_var = tk.StringVar()
        
        # File input row
        ttk.Label(input_frame, text="File Path:").grid(row=0, column=0, sticky=tk.W, pady=8)
        
        entry = ttk.Entry(input_frame, textvariable=path_var, width=80)
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(15, 10), pady=8)
        
        # Create bound method to avoid lambda closure issues
        def browse_file_command():
            self._browse_file(path_var, file_types)
        
        browse_button = ttk.Button(input_frame, text="📂 Browse",
                                  command=browse_file_command)
        browse_button.grid(row=0, column=2, pady=8)
        
        return {
            'frame': section_frame,
            'path_var': path_var,
            'entry': entry,
            'browse_button': browse_button,
            'input_frame': input_frame
        }
    
    def create_log_section(self, parent: ttk.Widget, title: str = "📊 ANALYSIS & PROGRESS LOG") -> Dict[str, Any]:
        """
        Create a standardized log section.
        
        Returns:
            Dict with 'frame', 'text_widget', 'export_button', 'clear_button'
        """
        log_frame = ttk.LabelFrame(parent, text=title, 
                                 style='Section.TLabelframe', padding="15")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        content_frame = ttk.Frame(log_frame)
        content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)
        
        # Log text area
        log_text = scrolledtext.ScrolledText(
            content_frame, height=12, width=85, font=('Consolas', 9), state=tk.DISABLED,
            bg=AppConstants.COLORS['surface'], fg=AppConstants.COLORS['text_primary'],
            selectbackground=AppConstants.COLORS['accent'], relief='solid', borderwidth=1
        )
        log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Log controls
        log_controls = ttk.Frame(content_frame)
        log_controls.grid(row=0, column=1, sticky=tk.N, padx=(10, 0))
        
        # Create bound methods to avoid lambda closure issues
        def export_log_command():
            self._export_log(log_text)
        
        def clear_log_command():
            self._clear_log(log_text)
        
        export_button = ttk.Button(log_controls, text="📤 Export Log", 
                                  command=export_log_command,
                                  style='Secondary.TButton', width=16)
        export_button.pack(pady=(0, 5), anchor=tk.W)
        
        clear_button = ttk.Button(log_controls, text="🗑️ Clear Log",
                                 command=clear_log_command,
                                 style='Secondary.TButton', width=16)
        clear_button.pack(anchor=tk.W)
        
        self.log_text = log_text  # Store reference for logging
        
        return {
            'frame': log_frame,
            'text_widget': log_text,
            'export_button': export_button,
            'clear_button': clear_button
        }
    
    def create_action_buttons(self, parent: ttk.Widget, buttons: List[Dict[str, Any]]) -> Dict[str, ttk.Button]:
        """
        Create standardized action buttons.
        
        Args:
            parent: Parent widget
            buttons: List of button configs with 'text', 'command', 'style', 'state'
            
        Returns:
            Dict mapping button text to button widget
        """
        button_frame = ttk.Frame(parent)
        button_widgets = {}
        
        for i, config in enumerate(buttons):
            button = ttk.Button(
                button_frame, 
                text=config['text'],
                command=config['command'],
                style=config.get('style', 'Action.TButton'),
                state=config.get('state', tk.NORMAL)
            )
            button.pack(side=tk.LEFT, padx=(0, 15) if i < len(buttons) - 1 else 0)
            button_widgets[config['text']] = button
        
        return button_widgets
    
    def create_progress_bar(self, parent: ttk.Widget) -> Dict[str, Any]:
        """Create a standardized enhanced progress bar
        
        Returns:
            Dict with 'frame', 'progress_bar'
        """
        # Progress container frame
        self.progress_frame = ttk.Frame(parent)
        
        # Just the progress bar - no status text
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', maximum=100)
        self.progress_bar.pack(fill='x')
        
        # Set progress_label to None since we're not using it
        self.progress_label = None
        
        return {
            'frame': self.progress_frame,
            'progress_bar': self.progress_bar
        }
    
    def log_message(self, message: str):
        """Log message to the tab's log area"""
        if self.log_text:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.config(state=tk.DISABLED)
            self.log_text.see(tk.END)
            self.frame.update_idletasks()
    
    def update_progress(self, progress_percent: int, message: str = "", show: bool = True, persist: bool = False):
        """Update progress bar - Universal progress system
        
        Args:
            progress_percent: Progress percentage (0-100)
            message: Status message to display (logged but not shown in UI)
            show: Whether to show or hide progress
            persist: If True, keep progress bar visible after operation (don't hide on 100%)
        """
        if show:
            # Show progress frame if not already visible
            if self.progress_frame and not self.progress_frame.winfo_viewable():
                # Auto-position progress frame - subclasses can override grid position
                self._position_progress_frame()
            
            # Update progress bar value
            if self.progress_bar:
                self.progress_bar['value'] = progress_percent
            
            # Set persistence flag if requested
            if persist:
                self.progress_persist = True
            
            # Log the progress message with percentage (no UI status text)
            if progress_percent > 0 and message:
                self.log_message(f"⏳ {progress_percent}% - {message}")
            elif message:
                self.log_message(f"⏳ {message}")
            elif progress_percent > 0:
                self.log_message(f"⏳ {progress_percent}% complete")
            
            self.is_busy = True
            
        else:
            # Hide progress frame only if not persisting
            if not self.progress_persist:
                if self.progress_bar:
                    self.progress_bar['value'] = 0  # Reset to empty
                if self.progress_frame:
                    self.progress_frame.grid_remove()
            else:
                # Keep progress bar visible but reset to 0
                if self.progress_bar:
                    self.progress_bar['value'] = 0
                
            self.is_busy = False
    
    def _position_progress_frame(self):
        """Position the progress frame - can be overridden by subclasses"""
        # Default positioning - subclasses should override this method
        # to position the progress frame appropriately for their layout
        if self.progress_frame:
            # Try to find a good default position - bottom of the frame
            self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))
    
    def run_in_background(self, target_func: Callable, 
                         success_callback: Callable = None,
                         error_callback: Callable = None,
                         progress_steps: List[tuple] = None):
        """
        Run a function in background thread with enhanced progress indication.
        
        Args:
            target_func: Function to run in background
            success_callback: Called on success with result
            error_callback: Called on error with exception
            progress_steps: List of (message, percentage) tuples for progress updates
        """
        import traceback
        
        def thread_target():
            """Thread target with proper error handling and closures."""
            result = None
            caught_error = None
            
            try:
                # Show initial progress only if no custom progress steps
                if progress_steps:
                    first_step = progress_steps[0]
                    self.frame.after(0, lambda: self.update_progress(first_step[1], first_step[0]))
                else:
                    # Don't show generic progress - let the target_func handle its own progress
                    pass
                
                result = target_func()
                
            except Exception as e:
                caught_error = e
                # Log the full traceback for debugging
                self.log_message(f"❌ Background operation failed: {e}")
                self.log_message(f"📋 Traceback: {traceback.format_exc()}")
            
            # Schedule callbacks on main thread with proper closures
            def schedule_callbacks():
                try:
                    if caught_error is not None:
                        # Handle error
                        if error_callback:
                            error_callback(caught_error)
                        else:
                            self._default_error_handler(caught_error)
                    else:
                        # Show completion progress if steps provided
                        if progress_steps:
                            self.update_progress(100, "Operation complete!")
                            import time
                            time.sleep(0.5)  # Brief pause to show completion
                        
                        # Handle success
                        if success_callback:
                            success_callback(result)
                
                except Exception as callback_error:
                    # Handle callback errors
                    self.log_message(f"❌ Callback error: {callback_error}")
                    self._default_error_handler(callback_error)
                
                finally:
                    # Only hide progress if not persisting
                    if not self.progress_persist:
                        self.update_progress(0, "", False)
            
            # Schedule on main thread
            if hasattr(self, 'frame') and self.frame:
                self.frame.after(0, schedule_callbacks)
            else:
                # Fallback if no frame available
                schedule_callbacks()
        
        thread = threading.Thread(target=thread_target, daemon=True)
        thread.start()
    
    def _default_error_handler(self, error: Exception):
        """Default error handler"""
        self.log_message(f"❌ Error: {error}")
        messagebox.showerror("Error", str(error))
    
    def _browse_file(self, path_var: tk.StringVar, file_types: List[tuple]):
        """Common file browsing logic"""
        file_path = filedialog.askopenfilename(
            title="Select File",
            filetypes=file_types
        )
        if file_path:
            path_var.set(file_path)
    
    def _export_log(self, log_widget: scrolledtext.ScrolledText):
        """Export log content to file"""
        try:
            log_content = log_widget.get(1.0, tk.END)
            file_path = filedialog.asksaveasfilename(
                title="Export Log", defaultextension=".txt",
                filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"{self.tool_name} - Analysis Log\n")
                    f.write(f"Generated by {AppConstants.COMPANY_NAME}\n")
                    f.write(f"{'='*50}\n\n")
                    f.write(log_content)
                
                self.log_message(f"✅ Log exported to: {file_path}")
                messagebox.showinfo("Export Complete", f"Log exported successfully!\n\n{file_path}")
        
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export log: {e}")
    
    def _clear_log(self, log_widget: scrolledtext.ScrolledText):
        """Clear log content"""
        log_widget.config(state=tk.NORMAL)
        log_widget.delete(1.0, tk.END)
        log_widget.config(state=tk.DISABLED)
        self._show_welcome_message()
    
    def _show_welcome_message(self):
        """Show welcome message - can be overridden by subclasses"""
        self.log_message(f"🎉 Welcome to {self.tool_name}!")
        self.log_message("=" * 60)
    
    def create_help_window(self, title: str, content_creator: Callable) -> tk.Toplevel:
        """
        Create a standardized help window.
        
        Args:
            title: Window title
            content_creator: Function that creates content in the window
            
        Returns:
            The help window
        """
        help_window = tk.Toplevel(self.main_app.root)
        help_window.title(title)
        help_window.geometry("650x620")
        help_window.resizable(False, False)
        help_window.transient(self.main_app.root)
        help_window.grab_set()
        
        # Center window
        help_window.geometry(f"+{self.main_app.root.winfo_rootx() + 50}+{self.main_app.root.winfo_rooty() + 50}")
        help_window.configure(bg=AppConstants.COLORS['background'])
        
        # Create content
        content_creator(help_window)
        
        # Bind escape key with proper closure
        def close_window(event=None):
            help_window.destroy()
        
        help_window.bind('<Escape>', close_window)
        
        return help_window
    
    def show_error(self, title: str, message: str):
        """Show error dialog"""
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str):
        """Show info dialog"""
        messagebox.showinfo(title, message)
    
    def show_warning(self, title: str, message: str):
        """Show warning dialog"""
        messagebox.showwarning(title, message)
    
    def ask_yes_no(self, title: str, message: str) -> bool:
        """Show yes/no dialog"""
        return messagebox.askyesno(title, message)


class FileInputMixin:
    """
    Mixin for tabs that need file input functionality.
    """
    
    def clean_file_path(self, path: str) -> str:
        """Clean file path by removing quotes and normalizing"""
        if not path:
            return path
        
        cleaned = path.strip()
        
        # Remove surrounding quotes
        if len(cleaned) >= 2:
            if (cleaned.startswith('"') and cleaned.endswith('"')) or \
               (cleaned.startswith("'") and cleaned.endswith("'")):
                cleaned = cleaned[1:-1]
        
        # Normalize path separators
        return str(Path(cleaned)) if cleaned else cleaned
    
    def setup_path_cleaning(self, path_var: tk.StringVar):
        """Setup automatic path cleaning on a StringVar"""
        def on_path_change(*args):
            current = path_var.get()
            cleaned = self.clean_file_path(current)
            if cleaned != current:
                path_var.set(cleaned)
        
        path_var.trace('w', on_path_change)


class ValidationMixin:
    """
    Mixin for tabs that need validation functionality.
    """
    
    def validate_file_exists(self, file_path: str, file_description: str = "File") -> None:
        """Validate that a file exists"""
        if not file_path:
            raise ValueError(f"{file_description} path is required")
        
        path_obj = Path(file_path)
        if not path_obj.exists():
            raise ValueError(f"{file_description} not found: {file_path}")
        
        if not path_obj.is_file():
            raise ValueError(f"{file_description} path must point to a file: {file_path}")
    
    def validate_pbip_file(self, file_path: str, file_description: str = "File") -> None:
        """Validate that a file is a valid PBIP file"""
        self.validate_file_exists(file_path, file_description)
        
        if not file_path.lower().endswith('.pbip'):
            raise ValueError(f"{file_description} must be a .pbip file")
        
        # Check for corresponding .Report directory
        path_obj = Path(file_path)
        report_dir = path_obj.parent / f"{path_obj.stem}.Report"
        if not report_dir.exists():
            raise ValueError(f"{file_description} missing corresponding .Report directory")
