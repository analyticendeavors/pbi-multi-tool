"""
Enhanced Base Tool Architecture for Power BI External Tools
Built by Reid Havens of Analytic Endeavors

Enhanced version that works with the new tool manager system.
"""

import os
import sys
import json
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Dict, List, Optional, Any
import datetime
import uuid

from core.tool_manager import ToolManager


@dataclass
class ToolConfiguration:
    """Configuration container for external tools"""
    name: str
    version: str
    description: str
    author: str
    website: str
    icon_path: Optional[str] = None
    log_level: str = "INFO"


class SecurityLogger:
    """Secure logging implementation with audit trail"""
    
    def __init__(self, tool_name: str, log_dir: Optional[str] = None):
        self.tool_name = tool_name
        
        # Create secure log directory
        if log_dir is None:
            log_dir = Path.home() / "AppData" / "Local" / "AnalyticEndeavors" / "Logs"
        
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # Configure logging with security best practices
        log_file = self.log_dir / f"{tool_name}_{datetime.date.today().isoformat()}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        self.logger = logging.getLogger(tool_name)
        self.logger.info(f"Security logger initialized for {tool_name}")
    
    def log_operation(self, operation: str, details: Dict[str, Any] = None):
        """Log operations with structured data"""
        message = f"Operation: {operation}"
        if details:
            message += f" | Details: {json.dumps(details, default=str)}"
        self.logger.info(message)
    
    def log_security_event(self, event: str, severity: str = "INFO"):
        """Log security-relevant events"""
        self.logger.log(
            getattr(logging, severity.upper()),
            f"SECURITY: {event}"
        )


class EnhancedBaseExternalTool(ABC):
    """
    Enhanced base class for Power BI External Tools that works with ToolManager.
    
    This maintains all the security features while being compatible with the 
    new tool management system.
    """
    
    def __init__(self, config: ToolConfiguration):
        self.config = config
        self.logger = SecurityLogger(config.name.replace(" ", ""))
        self.root = None
        self.is_running = False
        self.tool_manager = None
        
        # Initialize security context
        self._initialize_security_context()
        
        self.logger.log_operation("Tool initialization", {
            "name": config.name,
            "version": config.version,
            "author": config.author
        })
    
    def _initialize_security_context(self):
        """Initialize secure execution context"""
        # Validate execution environment
        try:
            # Check if running from expected location
            exe_path = Path(sys.executable if getattr(sys, 'frozen', False) else __file__)
            self.execution_path = exe_path.parent
            
            # Log security context
            self.logger.log_security_event(
                f"Execution context validated: {self.execution_path}"
            )
            
        except Exception as e:
            self.logger.log_security_event(
                f"Security context validation failed: {e}",
                "WARNING"
            )
    
    @abstractmethod
    def create_ui(self) -> tk.Tk:
        """Create the tool's user interface - must be implemented by subclasses"""
        pass
    
    @abstractmethod
    def perform_tool_operation(self, **kwargs) -> bool:
        """Perform the main tool operation - must be implemented by subclasses"""
        pass
    
    def validate_powerbi_integration(self) -> bool:
        """Validate Power BI Desktop integration"""
        try:
            # Check for Power BI Desktop installation
            powerbi_paths = [
                Path(os.environ.get('PROGRAMFILES', '')) / "Microsoft Power BI Desktop",
                Path(os.environ.get('PROGRAMFILES(X86)', '')) / "Microsoft Power BI Desktop",
                Path.home() / "AppData" / "Local" / "Microsoft" / "WindowsApps"
            ]
            
            powerbi_found = False
            for path in powerbi_paths:
                if path.exists():
                    # Look for Power BI Desktop executable
                    for pbi_exe in path.rglob("PBIDesktop.exe"):
                        powerbi_found = True
                        self.logger.log_operation("Power BI Desktop found", {"path": str(pbi_exe)})
                        break
            
            if not powerbi_found:
                self.logger.log_security_event(
                    "Power BI Desktop not found - tool may not integrate properly",
                    "WARNING"
                )
            
            return powerbi_found
            
        except Exception as e:
            self.logger.log_security_event(
                f"Power BI validation error: {e}",
                "ERROR"
            )
            return False
    
    def create_powerbi_tool_definition(self, install_path: Path) -> Dict[str, Any]:
        """Create Power BI External Tool definition"""
        return {
            "$schema": "https://powerbi.microsoft.com/schemas/pbitool.json",
            "name": self.config.name,
            "description": self.config.description,
            "path": str(install_path / f"{self.config.name.replace(' ', '')}.exe"),
            "iconData": "",
            "version": self.config.version,
            "website": self.config.website,
            "arguments": [],
            "metadata": {
                "author": self.config.author,
                "created": datetime.datetime.now().isoformat(),
                "security_validated": True
            }
        }
    
    def setup_error_handling(self):
        """Setup comprehensive error handling"""
        def handle_exception(exc_type, exc_value, exc_traceback):
            """Global exception handler"""
            if issubclass(exc_type, KeyboardInterrupt):
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return
            
            # Log the exception securely
            self.logger.log_security_event(
                f"Unhandled exception: {exc_type.__name__}: {exc_value}",
                "ERROR"
            )
            
            # Show user-friendly error dialog
            if self.root:
                messagebox.showerror(
                    "Application Error",
                    f"An unexpected error occurred.\n\n"
                    f"Error details have been logged for analysis.\n"
                    f"Please contact support if this problem persists.\n\n"
                    f"Reference: {datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
                )
        
        sys.excepthook = handle_exception
    
    def create_secure_ui_base(self) -> tk.Tk:
        """Create secure UI base with consistent styling"""
        root = tk.Tk()
        root.title(f"{self.config.name} v{self.config.version}")
        
        # Set professional styling
        style = ttk.Style()
        style.theme_use('vista' if 'vista' in style.theme_names() else 'default')
        
        # Configure window properties
        root.geometry("800x600")
        root.minsize(600, 400)
        
        # Set icon if available
        if self.config.icon_path and Path(self.config.icon_path).exists():
            try:
                root.iconbitmap(self.config.icon_path)
            except Exception:
                pass  # Icon loading is non-critical
        
        # Create status bar - REMOVED: Redundant with built-in logging systems
        # self.status_var = tk.StringVar()
        # self.status_var.set("Ready")
        # status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN)
        # status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        return root
    
    def update_status(self, message: str):
        """Update status bar with operation feedback"""
        if hasattr(self, 'status_var'):
            self.status_var.set(f"{datetime.datetime.now().strftime('%H:%M:%S')} - {message}")
            if self.root:
                self.root.update_idletasks()
    
    def run(self):
        """Main execution method with security checks"""
        try:
            self.logger.log_operation("Application startup")
            
            # Setup error handling
            self.setup_error_handling()
            
            # Validate Power BI integration
            self.validate_powerbi_integration()
            
            # Create UI
            self.root = self.create_ui()
            self.is_running = True
            
            # Center window
            self.root.eval('tk::PlaceWindow . center')
            
            self.logger.log_operation("UI created, starting main loop")
            
            # Start main loop
            self.root.mainloop()
            
        except Exception as e:
            self.logger.log_security_event(
                f"Application startup failed: {e}",
                "ERROR"
            )
            raise
        finally:
            self.is_running = False
            self.logger.log_operation("Application shutdown")
    
    def run_with_tool_manager(self, tool_manager: ToolManager):
        """Enhanced run method that works with tool manager"""
        self.tool_manager = tool_manager
        
        try:
            self.logger.log_operation("Application startup with tool manager")
            
            # Setup error handling
            self.setup_error_handling()
            
            # Validate Power BI integration
            self.validate_powerbi_integration()
            
            # Create UI with tool manager integration
            self.root = self.create_ui()
            self.is_running = True
            
            # Center window
            self.root.eval('tk::PlaceWindow . center')
            
            self.logger.log_operation("UI created with tool manager, starting main loop")
            
            # Start main loop
            self.root.mainloop()
            
        except Exception as e:
            self.logger.log_security_event(
                f"Application startup failed: {e}",
                "ERROR"
            )
            raise
        finally:
            self.is_running = False
            self.logger.log_operation("Application shutdown")
