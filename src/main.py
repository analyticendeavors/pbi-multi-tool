"""
Enhanced Power BI Report Tools - Main Application with Tool Manager Integration
Built by Reid Havens of Analytic Endeavors

This is the new main entry point that uses the ToolManager system for
automatic tool discovery and management.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

# Add parent directory to Python path for organized imports
sys.path.append(str(Path(__file__).parent))
sys.path.append(str(Path(__file__).parent.parent))

from core.constants import AppConstants
from core.enhanced_base_tool import EnhancedBaseExternalTool, ToolConfiguration
from core.tool_manager import get_tool_manager


class EnhancedPowerBIReportToolsApp(EnhancedBaseExternalTool):
    """
    Enhanced Power BI Report Tools with automatic tool discovery
    Uses the new ToolManager system for plugin-like architecture
    """
    
    def __init__(self):
        # Initialize with tool configuration
        config = ToolConfiguration(
            name="Enhanced Power BI Report Tools",
            version="1.2.0",
            description="Professional suite for Power BI report management with plugin architecture",
            author="Reid Havens",
            website="https://www.analyticendeavors.com",
            icon_path="assets/favicon.ico"
        )
        
        super().__init__(config)
        
        # Get tool manager
        self.tool_manager = get_tool_manager()
        
        # Application state
        self.notebook = None
        self.tool_tabs = {}
        
        # Tab state tracking for custom window sizes
        self.tab_states = {}  # Stores: {tool_id: {'custom_size': (width, height), 'is_modified': False}}
        self.current_tab_id = None  # Track current tab to detect changes
        
        # Initialize tools
        self._initialize_tools()
    
    def _initialize_tools(self):
        """Initialize and register all tools"""
        # Discover and register tools automatically
        registered_count = self.tool_manager.discover_and_register_tools("tools")
        
        if registered_count == 0:
            self.logger.log_security_event("No tools were discovered", "WARNING")
    
    def create_ui(self) -> tk.Tk:
        """Create the main tabbed user interface with tool manager integration"""
        # Create secure UI base
        root = self.create_secure_ui_base()
        root.title(f"Enhanced Power BI Report Tools v{self.config.version}")
        
        # Set initial size but don't position yet
        root.geometry("1100x950")  # Reduced width from 1150 to 1100 (50px narrower)
        
        self._setup_main_interface(root)
        
        # Center the window after everything is set up
        self._center_window(root)
        
        return root
    
    def _center_window(self, window):
        """Center the window horizontally, position near top vertically"""
        window.update_idletasks()  # Ensure window dimensions are calculated
        
        # Get window dimensions
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        
        # Get screen dimensions
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        # Calculate position: center horizontally, closer to top with comfortable margin
        x = (screen_width - window_width) // 2
        y = 65  # 65 pixels from top of screen
        
        # Ensure the window doesn't go off-screen
        x = max(0, x)
        y = max(0, y)
        
        # Set the window position - force it multiple times to ensure it takes effect
        window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        window.update()  # Force update
        window.geometry(f"+{x}+{y}")  # Force position again after update
    
    def _setup_main_interface(self, root):
        """Setup the main tabbed interface with tool manager"""
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Setup header
        self._setup_header(main_frame)
        
        # Setup tabbed notebook
        self._setup_notebook(main_frame)
        
        # Setup tabs using tool manager
        self._setup_tabs_with_tool_manager()
    
    def _setup_header(self, main_frame):
        """Setup common header for all tabs"""
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Left: Branding
        brand_frame = ttk.Frame(header_frame)
        brand_frame.pack(side=tk.LEFT)
        
        # Setup professional styling
        self._setup_header_styling()
        
        ttk.Label(brand_frame, text=f"📊 {AppConstants.COMPANY_NAME.upper()}", 
                 style='Brand.TLabel').pack(anchor=tk.W)
        
        # Title row with subtitle inline to the right
        title_row = ttk.Frame(brand_frame)
        title_row.pack(anchor=tk.W, pady=(5, 0))
        
        ttk.Label(title_row, text="Enhanced Power BI Report Tools", 
                 style='Title.TLabel').pack(side=tk.LEFT)
        ttk.Label(title_row, text=" —  Professional suite for Power BI report management", 
                 style='Subtitle.TLabel').pack(side=tk.LEFT)
        ttk.Label(brand_frame, text=f"Built by {AppConstants.COMPANY_FOUNDER} of {AppConstants.COMPANY_NAME}", 
                 style='Subtitle.TLabel').pack(anchor=tk.W, pady=(8, 0))
        
        # Right: Action buttons
        action_frame = ttk.Frame(header_frame)
        action_frame.pack(side=tk.RIGHT)
        
        ttk.Button(action_frame, text="🌐 WEBSITE", 
                  command=self.open_company_website, 
                  style='Brand.TButton').pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(action_frame, text="ℹ️ ABOUT", 
                  command=self.show_about_dialog, 
                  style='Info.TButton').pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(action_frame, text="❓ HELP", 
                  command=self.show_help_dialog, 
                  style='Info.TButton').pack(side=tk.RIGHT, padx=(5, 0))
    
    def _setup_header_styling(self):
        """Setup styling for header elements"""
        style = ttk.Style()
        colors = AppConstants.COLORS
        
        styles = {
            'Brand.TLabel': {'background': colors['background'], 'foreground': colors['primary'], 'font': ('Segoe UI', 16, 'bold')},
            'Title.TLabel': {'background': colors['background'], 'foreground': colors['text_primary'], 'font': ('Segoe UI', 18, 'bold')},
            'Subtitle.TLabel': {'background': colors['background'], 'foreground': colors['text_secondary'], 'font': ('Segoe UI', 10)},
            'Brand.TButton': {'background': colors['accent'], 'foreground': colors['surface'], 'font': ('Segoe UI', 10, 'bold'), 'padding': (15, 8)},
            'Info.TButton': {'background': colors['info'], 'foreground': colors['surface'], 'font': ('Segoe UI', 9), 'padding': (12, 6)},
        }
        
        for style_name, config in styles.items():
            style.configure(style_name, **config)
    
    def _setup_notebook(self, main_frame):
        """Setup the tabbed notebook widget"""
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Bind tab change event for dynamic height adjustment
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        
        # Style the notebook
        style = ttk.Style()
        colors = AppConstants.COLORS
        
        style.configure('TNotebook', 
                       background=colors['background'],
                       borderwidth=1,
                       relief='solid')
        
        style.configure('TNotebook.Tab',
                       background=colors['surface'],
                       foreground=colors['text_primary'],
                       font=('Segoe UI', 11, 'bold'),
                       padding=(25, 12),
                       borderwidth=1,
                       relief='solid')
        
        # Tab hover and selection effects
        style.map('TNotebook.Tab',
                 background=[('selected', colors['primary']),
                           ('active', colors['accent'])],
                 foreground=[('selected', colors['surface']),
                           ('active', colors['surface'])])
    
    def _setup_tabs_with_tool_manager(self):
        """Setup tabs using the tool manager with controlled ordering"""
        try:
            # Create tabs for all enabled tools
            self.tool_tabs = self.tool_manager.create_tool_tabs(self.notebook, self)
            
            if not self.tool_tabs:
                # No tools available - show error
                self._show_no_tools_error()
                return
            
            # Reorder tabs for better UX
            self._reorder_tabs_for_better_ux()
            
            # Set default tab
            self.notebook.select(0)
            
        except Exception as e:
            self.logger.log_security_event(f"Failed to setup tabs: {e}", "ERROR")
            self._show_tool_error(str(e))
    
    def _reorder_tabs_for_better_ux(self):
        """Reorder tabs to put most commonly used tools first"""
        desired_order = [
            "report_merger",           # Most important - should be first
            "pbip_layout_optimizer",   # Layout tool - second
            "report_cleanup",          # New cleanup tool - third
            "column_width",           # Column width tool - fourth
            "advanced_copy"               # Utility tool - last
        ]
        
        # Get current tabs
        current_tabs = []
        for i in range(self.notebook.index("end")):
            tab_widget = self.notebook.nametowidget(self.notebook.tabs()[i])
            tab_text = self.notebook.tab(i, "text")
            current_tabs.append((tab_widget, tab_text))
        
        # Clear notebook
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        
        # Find tools by ID and add in desired order
        tools = list(self.tool_manager.get_enabled_tools())
        tools_by_id = {tool.tool_id: tool for tool in tools}
        
        # Add tabs in desired order
        added_tools = set()
        for tool_id in desired_order:
            if tool_id in tools_by_id:
                tool = tools_by_id[tool_id]
                tab = self.tool_tabs.get(tool_id)
                if tab:
                    self.notebook.add(tab.get_frame(), text=tool.get_tab_title())
                    added_tools.add(tool_id)
        
        # Add any remaining tools that weren't in the desired order
        for tool in tools:
            if tool.tool_id not in added_tools:
                tab = self.tool_tabs.get(tool.tool_id)
                if tab:
                    self.notebook.add(tab.get_frame(), text=tool.get_tab_title())
    
    def _show_no_tools_error(self):
        """Show error when no tools are available"""
        error_frame = ttk.Frame(self.notebook)
        self.notebook.add(error_frame, text="❌ No Tools")
        
        ttk.Label(error_frame, text="No tools were discovered!", 
                 font=('Segoe UI', 16, 'bold'),
                 foreground=AppConstants.COLORS['error']).pack(pady=50)
        
        ttk.Label(error_frame, text="Please check the tools directory and restart the application.",
                 font=('Segoe UI', 12)).pack()
    
    def _show_tool_error(self, error_message: str):
        """Show error when tool setup fails"""
        error_frame = ttk.Frame(self.notebook)
        self.notebook.add(error_frame, text="❌ Tool Error")
        
        ttk.Label(error_frame, text="Tool Setup Failed!", 
                 font=('Segoe UI', 16, 'bold'),
                 foreground=AppConstants.COLORS['error']).pack(pady=50)
        
        ttk.Label(error_frame, text=f"Error: {error_message}",
                 font=('Segoe UI', 10)).pack()
    
    def _on_tab_changed(self, event=None):
        """Handle tab changes with smart state tracking - remembers custom sizes"""
        try:
            if not self.notebook:
                return
            
            # Get current tab and window size
            current_tab_id = self.notebook.select()
            current_width = self.root.winfo_width()
            current_height = self.root.winfo_height()
            
            # Save previous tab's current size if it was modified
            if self.current_tab_id:
                # Get previous tool ID
                prev_tool_id = self._get_tool_id_from_tab(self.current_tab_id)
                if prev_tool_id and prev_tool_id in self.tab_states:
                    # Check if tab was modified (has custom size different from default)
                    if self.tab_states[prev_tool_id]['is_modified']:
                        # Save the current window size as custom size
                        self.tab_states[prev_tool_id]['custom_size'] = (current_width, current_height)
            
            # Get tool ID for new tab
            tool_id = self._get_tool_id_from_tab(current_tab_id)
            if not tool_id:
                return
            
            # Initialize state tracking for this tab if not exists
            if tool_id not in self.tab_states:
                self.tab_states[tool_id] = {
                    'custom_size': None,
                    'is_modified': False,
                    'default_size': self._get_default_size_for_tool(tool_id)
                }
            
            # Check if tab has been modified by looking at the tool tab
            is_modified = self._check_if_tab_modified(tool_id)
            self.tab_states[tool_id]['is_modified'] = is_modified
            
            # Determine which size to use
            if is_modified and self.tab_states[tool_id]['custom_size']:
                # Tab is modified and has a custom size - use it
                width, height = self.tab_states[tool_id]['custom_size']
            else:
                # Tab is pristine - use default size
                width, height = self.tab_states[tool_id]['default_size']
            
            # Get current window position to preserve it
            current_x = self.root.winfo_x()
            current_y = self.root.winfo_y()
            
            # Apply the new geometry while preserving current position
            self.root.geometry(f"{width}x{height}+{current_x}+{current_y}")
            
            # Update current tab tracking
            self.current_tab_id = current_tab_id
            
            # Force window update
            self.root.update_idletasks()
            self.root.update()
                    
        except Exception as e:
            pass  # Silently handle tab change errors
    
    def _get_tool_id_from_tab(self, tab_id):
        """Get tool ID from tab ID by checking tab text"""
        try:
            current_tab_text = self.notebook.tab(tab_id, "text")
            
            # Map tab text to tool ID
            if "Report Merger" in current_tab_text:
                return "report_merger"
            elif "Advanced Copy" in current_tab_text or "Advanced Page Copy" in current_tab_text:
                return "advanced_copy"
            elif "Layout Optimizer" in current_tab_text:
                return "pbip_layout_optimizer"
            elif "Report Cleanup" in current_tab_text:
                return "report_cleanup"
            elif "Table Column Widths" in current_tab_text or "Visual Cleanup" in current_tab_text:
                return "column_width"
            else:
                # For any other tools, try to match by checking tool tabs
                for tid, tab in self.tool_tabs.items():
                    if tab.get_frame() == self.notebook.nametowidget(tab_id):
                        return tid
        except:
            pass
        return None
    
    def _get_default_size_for_tool(self, tool_id):
        """Get default window size for a specific tool"""
        default_sizes = {
            "report_merger": (1150, 950),
            "advanced_copy": (1175, 860),  # Increased from 820 to 860 for Cross-PBIP mode
            "pbip_layout_optimizer": (1130, 850),
            "report_cleanup": (1100, 1095),
            "column_width": (1200, 1035)
        }
        return default_sizes.get(tool_id, (1250, 1150))  # Default for unknown tools
    
    def _check_if_tab_modified(self, tool_id):
        """Check if a tab has been modified (file selected, button clicked, etc.)"""
        try:
            tool_tab = self.tool_tabs.get(tool_id)
            if not tool_tab:
                return False
            
            # Check if tool tab has any indication of modification
            # Each tool should have a method to check if it's in pristine state
            if hasattr(tool_tab, 'is_tab_pristine'):
                return not tool_tab.is_tab_pristine()  # Returns True if modified
            
            # Fallback: check for common indicators of modification
            # Most tools use StringVar for file paths - check those first
            if hasattr(tool_tab, '__dict__'):
                for attr_name, attr_value in tool_tab.__dict__.items():
                    # Check StringVar instances (used for file paths)
                    if isinstance(attr_value, tk.StringVar):
                        value = attr_value.get().strip()
                        if value:  # Has content
                            return True
                    # Check IntVar instances (used for options)
                    elif isinstance(attr_value, tk.IntVar):
                        # Check if it differs from 0 (common default)
                        if attr_value.get() != 0:
                            return True
            
            # Additional fallback: check widgets in frame
            frame = tool_tab.get_frame()
            for child in frame.winfo_children():
                if self._widget_has_content(child):
                    return True
            
            return False
        except:
            return False
    
    def _widget_has_content(self, widget):
        """Recursively check if a widget or its children have content"""
        try:
            # Check if it's an Entry widget with text
            if isinstance(widget, tk.Entry) and widget.get().strip():
                return True
            
            # Check if it's a Text widget with content
            if isinstance(widget, tk.Text):
                content = widget.get("1.0", tk.END).strip()
                if content:
                    return True
            
            # Check if it's a StringVar with value
            if hasattr(widget, 'get'):
                try:
                    value = widget.get()
                    if isinstance(value, str) and value.strip():
                        return True
                except:
                    pass
            
            # Recursively check children
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    if self._widget_has_content(child):
                        return True
        except:
            pass
        
        return False
    
    def perform_tool_operation(self, **kwargs) -> bool:
        """Implementation required by EnhancedBaseExternalTool"""
        # This could coordinate operations across tools if needed
        return True
    
    # Shared utility methods for tools
    def open_company_website(self):
        """Open company website"""
        try:
            import webbrowser
            webbrowser.open(AppConstants.COMPANY_WEBSITE)
            self.update_status(f"🌐 Opening {AppConstants.COMPANY_NAME} website...")
        except Exception as e:
            self.show_error("Error", f"Could not open website: {e}")
    
    def show_help_dialog(self):
        """Show context-sensitive help based on active tab"""
        try:
            current_tab_id = self.notebook.select()
            current_tab_text = self.notebook.tab(current_tab_id, "text")
            
            # Map tab text to tool ID (more reliable after reordering)
            tool_id = None
            if "Report Merger" in current_tab_text:
                tool_id = "report_merger"
            elif "Advanced Copy" in current_tab_text or "Advanced Page Copy" in current_tab_text:
                tool_id = "advanced_copy"
            elif "Layout Optimizer" in current_tab_text:
                tool_id = "pbip_layout_optimizer"
            elif "Report Cleanup" in current_tab_text:
                tool_id = "report_cleanup"
            elif "Table Column Widths" in current_tab_text or "Visual Cleanup" in current_tab_text:
                tool_id = "column_width"
            else:
                # For any other tools, try to match by checking tool tabs
                for tid, tab in self.tool_tabs.items():
                    if tab.get_frame() == self.notebook.nametowidget(current_tab_id):
                        tool_id = tid
                        break
            
            # Get the tool tab and show its help dialog
            if tool_id and tool_id in self.tool_tabs:
                tool_tab = self.tool_tabs[tool_id]
                if tool_tab and hasattr(tool_tab, 'show_help_dialog'):
                    tool_tab.show_help_dialog()
                    return
            
            # Fallback to general help
            self.show_general_help()
            
        except Exception as e:
            self.logger.log_security_event(f"Help dialog error: {e}", "ERROR")
            import traceback
            traceback.print_exc()
            self.show_general_help()
    
    def show_general_help(self):
        """Show general application help"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Enhanced Power BI Report Tools - Help")
        help_window.geometry("600x600")  # Increased height to accommodate close button
        help_window.resizable(False, False)
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Center window
        help_window.geometry(f"+{self.root.winfo_rootx() + 100}+{self.root.winfo_rooty() + 100}")
        
        self._create_general_help_content(help_window)
    
    def _create_general_help_content(self, help_window):
        """Create general help content"""
        help_window.configure(bg=AppConstants.COLORS['background'])
        
        # Main container that reserves space for the button
        container = ttk.Frame(help_window, padding="20")
        container.pack(fill=tk.BOTH, expand=True)
        
        # Content frame for scrollable content (if needed)
        content_frame = ttk.Frame(container)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(content_frame, text="📊 Enhanced Power BI Report Tools", 
                 font=('Segoe UI', 16, 'bold'), 
                 foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W, pady=(0, 20))
        
        # Tool list
        tools = self.tool_manager.get_enabled_tools()
        
        ttk.Label(content_frame, text="Available Tools:", 
                 font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        for tool in tools:
            tool_frame = ttk.Frame(content_frame)
            tool_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(tool_frame, text=f"• {tool.name}", 
                     font=('Segoe UI', 11, 'bold'),
                     foreground=AppConstants.COLORS['primary']).pack(anchor=tk.W)
            ttk.Label(tool_frame, text=f"  {tool.description}", 
                     font=('Segoe UI', 10)).pack(anchor=tk.W, padx=(20, 0))
        
        # Button frame at bottom - fixed position
        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X, pady=(20, 0), side=tk.BOTTOM)
        
        close_button = ttk.Button(button_frame, text="❌ Close", 
                                command=help_window.destroy,
                                style='Action.TButton')
        close_button.pack(pady=(10, 0))
        
        help_window.bind('<Escape>', lambda event: help_window.destroy())
    
    def show_about_dialog(self):
        """Show about dialog with tool manager info"""
        about_window = tk.Toplevel(self.root)
        about_window.title(f"About - Enhanced Power BI Report Tools")
        about_window.geometry("500x615")
        about_window.resizable(False, False)
        about_window.transient(self.root)
        about_window.grab_set()
        
        # Center window
        about_window.geometry(f"+{self.root.winfo_rootx() + 100}+{self.root.winfo_rooty() + 100}")
        
        self._create_about_content(about_window)
    
    def _create_about_content(self, about_window):
        """Create about content with tool manager information"""
        about_window.configure(bg=AppConstants.COLORS['background'])
        
        main_frame = ttk.Frame(about_window, padding="30 30 30 5")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(main_frame, text="🚀", font=('Segoe UI', 48)).pack()
        ttk.Label(main_frame, text="Enhanced Power BI Report Tools", 
                 font=('Segoe UI', 18, 'bold'),
                 foreground=AppConstants.COLORS['primary']).pack(pady=(10, 5))
        
        # Tool Manager Status
        status = self.tool_manager.get_status_summary()
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(pady=(20, 0))
        
        ttk.Label(status_frame, text="🔧 Tool Manager Status:", 
                 font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W)
        ttk.Label(status_frame, text=f"• {status['enabled_tools']} tools enabled", 
                 font=('Segoe UI', 10)).pack(anchor=tk.W, pady=1)
        ttk.Label(status_frame, text=f"• {status['active_tabs']} active tabs", 
                 font=('Segoe UI', 10)).pack(anchor=tk.W, pady=1)
        
        # Description
        desc_frame = ttk.Frame(main_frame)
        desc_frame.pack(pady=(20, 0))
        
        description = [
            "Professional suite with plugin architecture:",
            "• Automatic tool discovery and registration",
            "• Modular, extensible design",
            "• Enhanced security and error handling",
            "",
            "⚠️ Requires PBIP format (PBIR) files only",
            "⚠️ NOT officially supported by Microsoft",
            "",
            f"Built by {AppConstants.COMPANY_FOUNDER}",
            f"of {AppConstants.COMPANY_NAME}"
        ]
        
        for line in description:
            style = 'RequirementText.TLabel' if "⚠️" in line else None
            ttk.Label(desc_frame, text=line, font=('Segoe UI', 10), style=style).pack(anchor=tk.W, pady=1)
        
        # Footer with proper closures
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=(25, 0))
        
        def close_about():
            about_window.destroy()
        
        ttk.Button(footer_frame, text="🌐 Visit Website", 
                  command=self.open_company_website).pack(side=tk.LEFT)
        ttk.Button(footer_frame, text="❌ Close", 
                  command=close_about).pack(side=tk.RIGHT)
        
        about_window.bind('<Escape>', lambda event: close_about())
    
    def show_error(self, title: str, message: str):
        """Show error dialog"""
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str):
        """Show info dialog"""
        messagebox.showinfo(title, message)


def main():
    """Main entry point with enhanced error handling"""
    try:
        app = EnhancedPowerBIReportToolsApp()
        app.run_with_tool_manager(app.tool_manager)
    except Exception as e:
        print(f"Critical Error: {e}")
        
        # Show error dialog if possible
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()  # Hide the empty window
            messagebox.showerror("Application Error", 
                               f"The application failed to start:\n\n{e}\n\n"
                               f"Please check the logs for more details.")
        except:
            pass  # If even error dialog fails, just print


if __name__ == "__main__":
    main()
